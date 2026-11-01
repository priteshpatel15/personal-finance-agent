from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
from dataclasses import dataclass
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any

from personal_budget_plaid_excel.application.connection_states import reconcile_sync_states
from personal_budget_plaid_excel.application.import_source_selection import (
    SelectMonthlyImportSource,
    source_label,
)
from personal_budget_plaid_excel.application.monthly_review import BuildMonthlyReview
from personal_budget_plaid_excel.application.source_coverage import build_plaid_source_context
from personal_budget_plaid_excel.application.report import HtmlReportService
from personal_budget_plaid_excel.application.sync import SyncService
from personal_budget_plaid_excel.domain.models import PlaidSyncResult, SyncItemState
from personal_budget_plaid_excel.domain.review import ReviewBuildResult
from personal_budget_plaid_excel.domain.source_mode import ImportSourceMode, ReviewSourceContext
from personal_budget_plaid_excel.infrastructure.keychain import KeychainUnavailable, MacOSKeychain
from personal_budget_plaid_excel.infrastructure.plaid_rest import PlaidConfig, PlaidRestClient
from personal_budget_plaid_excel.infrastructure.telemetry import Telemetry
from personal_budget_plaid_excel.infrastructure.workbook import BudgetWorkbook
from personal_budget_plaid_excel.infrastructure.connection_registry import ConnectionRegistry
from personal_budget_plaid_excel.infrastructure.evidence_manifest import JsonEvidenceManifest
from personal_budget_plaid_excel.infrastructure.review_workbook import OpenpyxlReviewWorkbookWriter
from personal_budget_plaid_excel.ports.review import BudgetReviewSource

DEFAULT_WORKBOOK = Path(os.getenv("BUDGET_WORKBOOK", "budget.xlsx"))


@dataclass(frozen=True)
class MonthlyPaths:
    staging_workbook: Path
    connection_registry: Path
    review_workbook: Path


def main(argv: list[str] | None = None) -> int:
    parser = _parser()
    args = parser.parse_args(argv)
    workbook_path = Path(args.workbook).expanduser().resolve()
    try:
        return args.func(args, workbook_path)
    except Exception as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="personal-budget-plaid-excel")
    parser.add_argument("--workbook", default=str(DEFAULT_WORKBOOK), help="Path to the authority budget workbook (default: $BUDGET_WORKBOOK or ./budget.xlsx).")
    sub = parser.add_subparsers(dest="command", required=True)

    doctor = sub.add_parser("doctor", help="Validate local config and workbook shape.")
    doctor.add_argument("--month", help="Month as YYYY-MM; defaults to the local current month.")
    doctor.add_argument("--output-dir", type=Path, default=Path("monthly-budgets"))
    doctor.add_argument("--review-output-dir", type=Path, help="Directory for the standalone review workbook.")
    doctor.add_argument("--json", action="store_true")
    doctor.set_defaults(func=_doctor)

    connect = sub.add_parser("connect", help="Connect Plaid accounts through a localhost Link flow.")
    connect.add_argument("--port", type=int, default=8765)
    connect.add_argument("--connections", type=int, default=1)
    connect.set_defaults(func=_connect)

    sync = sub.add_parser("sync", help="Sync Plaid transactions into the workbook.")
    sync.add_argument("--sync-json", type=Path, help="Apply a local fake Plaid sync JSON file.")
    sync.add_argument("--init-only", action="store_true", help="Only migrate workbook structure and save.")
    sync.set_defaults(func=_sync)

    report = sub.add_parser("report", help="Generate the HTML budget report.")
    report.add_argument("--month", help="Budget month as YYYY-MM. Defaults to current month.")
    report.add_argument("--output-dir", type=Path, default=Path("reports"))
    report.set_defaults(func=_report)

    run = sub.add_parser("run", help="Sync, generate report, and emit Gmail payload JSON.")
    run.add_argument("--month", help="Budget month as YYYY-MM. Defaults to current month.")
    run.add_argument("--output-dir", type=Path, default=Path("reports"))
    run.add_argument("--skip-sync", action="store_true", help="Generate report without Plaid sync.")
    run.set_defaults(func=_run)

    monthly = sub.add_parser("monthly-run", help="Create or refresh the current month's standalone budget workbook.")
    monthly.add_argument("--output-dir", type=Path, default=Path("monthly-budgets"))
    monthly.add_argument("--month", help="Month as YYYY-MM; defaults to the local current month.")
    monthly.add_argument(
        "--source",
        help="Required import source: plaid or direct-portal.",
    )
    monthly.add_argument("--sync-json", type=Path, help=argparse.SUPPRESS)
    monthly.add_argument(
        "--portal-json",
        type=Path,
        help="Local month-scoped Direct portal import manifest.",
    )
    monthly.add_argument("--evidence-json", type=Path, help="Local, already-collected review evidence JSON.")
    monthly.add_argument("--review-output-dir", type=Path, help="Directory for the standalone review workbook.")
    monthly.set_defaults(func=_monthly_run)

    review = sub.add_parser("review-run", help="Build or refresh the standalone monthly review workbook offline.")
    review.add_argument("--month", help="Month as YYYY-MM; defaults to the local current month.")
    review.add_argument("--output-dir", type=Path, default=Path("monthly-budgets"))
    review.add_argument("--source-workbook", type=Path, help="Existing monthly workbook; defaults to --workbook.")
    review.add_argument("--evidence-json", type=Path, help="Local, already-collected review evidence JSON.")
    review.set_defaults(func=_review_run)
    return parser


def _doctor(args: argparse.Namespace, workbook_path: Path) -> int:
    month = _parse_month(args.month)
    paths = _monthly_paths(args.output_dir, args.review_output_dir, month)
    checks: dict[str, Any] = {
        "workbook_exists": workbook_path.exists(),
        "workbook": str(workbook_path),
        "month": month.strftime("%Y-%m"),
        "plaid_env_present": bool(_env("PLAID_CLIENT_ID") and _env("PLAID_SECRET")),
        "keychain_available": shutil.which("security") is not None,
        "python": sys.executable,
        "connection_registry": str(paths.connection_registry),
        "connection_registry_exists": paths.connection_registry.exists(),
        "monthly_workbook": str(paths.staging_workbook),
        "monthly_workbook_exists": paths.staging_workbook.exists(),
        "review_workbook": str(paths.review_workbook),
        "review_workbook_exists": paths.review_workbook.exists(),
        "staging_output_separate_from_authority": paths.staging_workbook != workbook_path,
        "review_output_separate_from_authority": paths.review_workbook != workbook_path,
    }
    authority_states: list[SyncItemState] = []
    if workbook_path.exists():
        workbook = BudgetWorkbook(workbook_path)
        checks["budget_categories"] = len(workbook.read_budgets())
        authority_states = workbook.read_sync_states()
        checks["required_sheets_present"] = all(
            sheet in workbook.workbook.sheetnames for sheet in ["Manual Budget", "Actual Charges"]
        )
    target_states = (
        BudgetWorkbook(paths.staging_workbook).read_sync_states()
        if paths.staging_workbook.exists()
        else []
    )
    registry_states = ConnectionRegistry(paths.connection_registry).read()
    intended_states = _reconcile_monthly_states(
        target_exists=paths.staging_workbook.exists(),
        target_states=target_states,
        registry_states=registry_states,
        authority_states=authority_states,
    )
    checks.update({
        "registry_connection_count": len(registry_states),
        "monthly_workbook_connection_count": len(target_states),
        "authority_connection_count": len(authority_states),
        "intended_connection_count": len(intended_states),
        "refresh_ready": bool(intended_states),
    })
    if args.json:
        print(json.dumps(checks, indent=2, sort_keys=True))
    else:
        for key, value in checks.items():
            print(f"{key}: {value}")
    workbook_ready = bool(
        checks["workbook_exists"]
        and checks.get("required_sheets_present", False)
        and checks["staging_output_separate_from_authority"]
        and checks["review_output_separate_from_authority"]
    )
    if not workbook_ready:
        return 1
    return 0 if checks["refresh_ready"] else 2


def _connect(args: argparse.Namespace, workbook_path: Path) -> int:
    from personal_budget_plaid_excel.cli.connect import run_link_flow

    result = run_link_flow(workbook_path, PlaidRestClient(_plaid_config()), port=args.port, connections=args.connections)
    print(json.dumps({"connected": True, **result}, indent=2))
    return 0


def _sync(args: argparse.Namespace, workbook_path: Path) -> int:
    telemetry = _telemetry(workbook_path)
    workbook = BudgetWorkbook(workbook_path)
    workbook.ensure_structure()
    if args.init_only:
        backup = workbook.save_with_backup(backup=True)
        print(json.dumps({"initialized": True, "backup": str(backup) if backup else None}, indent=2))
        return 0
    if args.sync_json:
        totals = {"items": 0, "added": 0, "modified": 0, "removed": 0}
        for payload in _fake_sync_payloads(args.sync_json):
            result = PlaidSyncResult(
                item_id=payload.get("item_id", "fake-item"),
                added=payload.get("added", []),
                modified=payload.get("modified", []),
                removed=payload.get("removed", []),
                accounts=payload.get("accounts", []),
                next_cursor=payload.get("next_cursor", "fake-cursor"),
                request_ids=payload.get("request_ids", ["fake-request"]),
            )
            counts = workbook.apply_plaid_sync(result)
            workbook.upsert_sync_state(
                SyncItemState(
                    item_id=result.item_id,
                    institution_name=payload.get("institution_name", "Fake Plaid"),
                    cursor=result.next_cursor,
                    access_token_key=f"plaid:{result.item_id}",
                    last_successful_sync=datetime.now(timezone.utc),
                    last_request_id=",".join(result.request_ids),
                    updated_at=datetime.now(timezone.utc),
                )
            )
            totals["items"] += 1
            for name in ("added", "modified", "removed"):
                totals[name] += counts[name]
        backup = workbook.save_with_backup(backup=True)
        telemetry.event("fake_sync_applied", counts=totals)
        print(json.dumps({"synced": True, "counts": totals, "backup": str(backup) if backup else None}, indent=2))
        return 0

    states = workbook.read_sync_states()
    if not states:
        workbook.save_with_backup(backup=True)
        print("No Plaid items are connected. Run `personal-budget-plaid-excel connect` first.", file=sys.stderr)
        return 2
    service = SyncService(workbook_path, PlaidRestClient(_plaid_config()), MacOSKeychain(), telemetry)
    print(json.dumps(service.sync_all_items(), indent=2))
    return 0


def _report(args: argparse.Namespace, workbook_path: Path) -> int:
    month = _parse_month(args.month)
    output_path = HtmlReportService(workbook_path).write_report(args.output_dir.resolve(), month)
    print(str(output_path))
    return 0


def _run(args: argparse.Namespace, workbook_path: Path) -> int:
    if not args.skip_sync:
        sync_args = argparse.Namespace(sync_json=None, init_only=False)
        sync_code = _sync(sync_args, workbook_path)
        if sync_code not in (0, 2):
            return sync_code
    month = _parse_month(args.month)
    output_path = HtmlReportService(workbook_path).write_report(args.output_dir.resolve(), month)
    payload = {
        "to": "me",
        "subject": f"Budget report for {month:%B %Y}",
        "body_file": str(output_path),
        "content_type": "text/html",
    }
    print(json.dumps({"gmail_send_email": payload}, indent=2))
    return 0


def _monthly_run(args: argparse.Namespace, workbook_path: Path) -> int:
    """Refresh only a standalone monthly deliverable; never mutate the authority workbook."""
    decision = SelectMonthlyImportSource().execute(args.source)
    if decision.source is None:
        print(json.dumps(decision.public_payload(), indent=2))
        return 2
    if decision.source is ImportSourceMode.DIRECT_PORTAL:
        if args.portal_json is None:
            print(json.dumps({
                "status": "source_manifest_required",
                "import_source": {
                    "id": decision.source.value,
                    "label": source_label(decision.source),
                },
                "required_argument": "--portal-json",
                "prompt": (
                    "Collect the month from signed-in Bank of America, Chase, and "
                    "American Express Chrome sessions, then provide the local portal manifest."
                ),
                "data_collection_started": False,
                "email_sent": False,
            }, indent=2))
            return 2
        return _monthly_run_direct_portal(args, workbook_path, decision.source)
    return _monthly_run_plaid(args, workbook_path, decision.source)


def _monthly_run_plaid(
    args: argparse.Namespace,
    workbook_path: Path,
    source_mode: ImportSourceMode,
) -> int:
    """Run the existing Plaid-backed import after explicit source selection."""
    month = _parse_month(args.month)
    paths = _monthly_paths(args.output_dir, args.review_output_dir, month)
    target = paths.staging_workbook
    if target == workbook_path:
        raise ValueError("monthly staging workbook must be separate from the authority workbook")
    if paths.review_workbook == workbook_path:
        raise ValueError("monthly review workbook must be separate from the authority workbook")
    target.parent.mkdir(parents=True, exist_ok=True)
    created = not target.exists()
    if created:
        shutil.copy2(workbook_path, target)
    registry = ConnectionRegistry(paths.connection_registry)
    target_workbook = BudgetWorkbook(target)
    target_states = target_workbook.read_sync_states()
    registry_states = registry.read()
    authority_states = BudgetWorkbook(workbook_path).read_sync_states()
    states = _reconcile_monthly_states(
        target_exists=not created,
        target_states=target_states,
        registry_states=registry_states,
        authority_states=authority_states,
    )
    for state in states:
        target_workbook.upsert_sync_state(state)
    target_workbook.ensure_structure(as_of=month)
    target_workbook.save_with_backup(backup=True)
    sync_args = argparse.Namespace(sync_json=args.sync_json, init_only=False)
    code = _sync(sync_args, target)
    if code not in (0, 2):
        return code
    _retain_month(target, month)
    if code == 0:
        registry.replace(BudgetWorkbook(target).read_sync_states())
    refreshed_workbook = BudgetWorkbook(target)
    source_context = build_plaid_source_context(
        month,
        refreshed_workbook.read_sync_states(),
        refreshed_workbook.read_transactions(),
    )
    review_result = _build_review(
        target,
        paths.review_workbook.parent,
        month,
        args.evidence_json,
        source_context=source_context,
    )
    source_complete = review_result.source_coverage_complete
    status = (
        "source_not_connected"
        if code != 0
        else "completed" if source_complete else "completed_incomplete_coverage"
    )
    print(json.dumps({
        "status": status,
        "import_source": {
            "id": source_mode.value,
            "label": source_label(source_mode),
        },
        "monthly_workbook": str(target),
        "review_workbook": review_result.output_path,
        "month": month.strftime("%Y-%m"),
        "created": created,
        "refresh_status": (
            "refreshed"
            if code == 0 and source_complete
            else "refreshed_incomplete_coverage" if code == 0 else "not_connected"
        ),
        "fresh": code == 0 and source_complete,
        "review": _review_result_payload(review_result),
        "email_sent": False,
    }))
    return code


def _monthly_run_direct_portal(
    args: argparse.Namespace,
    workbook_path: Path,
    source_mode: ImportSourceMode,
) -> int:
    """Build a review from an already-collected, local Direct portal manifest."""
    from personal_budget_plaid_excel.infrastructure.portal_manifest import JsonPortalManifest
    from personal_budget_plaid_excel.infrastructure.portal_source import DirectPortalBudgetSource

    month = _parse_month(args.month)
    batch = JsonPortalManifest(args.portal_json.expanduser().resolve()).read()
    if batch.month != month:
        raise ValueError("portal manifest month does not match requested month")

    paths = _monthly_paths(args.output_dir, args.review_output_dir, month)
    if paths.review_workbook == workbook_path:
        raise ValueError("monthly review workbook must be separate from the authority workbook")
    source = DirectPortalBudgetSource(BudgetWorkbook(workbook_path), batch)
    created = not paths.review_workbook.exists()
    result = _build_review_source(
        source,
        paths.review_workbook.parent,
        month,
        args.evidence_json,
    )
    print(json.dumps({
        "status": "completed" if result.source_coverage_complete else "completed_incomplete_coverage",
        "import_source": {
            "id": source_mode.value,
            "label": source_label(source_mode),
        },
        "monthly_workbook": None,
        "review_workbook": result.output_path,
        "month": month.strftime("%Y-%m"),
        "created": created,
        "fresh": result.source_coverage_complete,
        "review": _review_result_payload(result),
        "email_sent": False,
    }))
    return 0


def _review_run(args: argparse.Namespace, workbook_path: Path) -> int:
    month = _parse_month(args.month)
    source = args.source_workbook.expanduser().resolve() if args.source_workbook else workbook_path
    result = _build_review(source, args.output_dir.resolve(), month, args.evidence_json)
    print(json.dumps({
        "review_workbook": result.output_path,
        "month": month.strftime("%Y-%m"),
        "review": _review_result_payload(result),
        "email_sent": False,
    }))
    return 0


def _build_review(
    source: Path,
    output_dir: Path,
    month: date,
    evidence_json: Path | None,
    *,
    source_context: ReviewSourceContext | None = None,
) -> ReviewBuildResult:
    return _build_review_source(
        BudgetWorkbook(source),
        output_dir,
        month,
        evidence_json,
        source_context=source_context,
    )


def _build_review_source(
    source: BudgetReviewSource,
    output_dir: Path,
    month: date,
    evidence_json: Path | None,
    *,
    source_context: ReviewSourceContext | None = None,
) -> ReviewBuildResult:
    output = output_dir / f"{month:%B}-{month:%Y}-Personal-Budget-Review.xlsx"
    evidence_reader = JsonEvidenceManifest(evidence_json.expanduser().resolve()) if evidence_json else None
    telemetry = Telemetry(output_dir / "logs" / "personal-budget-plaid-excel.jsonl")
    return BuildMonthlyReview(OpenpyxlReviewWorkbookWriter()).execute(
        source,
        output,
        month,
        evidence_reader,
        telemetry,
        source_context=source_context,
    )


def _monthly_paths(
    output_dir: Path,
    review_output_dir: Path | None,
    month: date,
) -> MonthlyPaths:
    staging_dir = output_dir.expanduser().resolve()
    review_dir = (review_output_dir or output_dir).expanduser().resolve()
    return MonthlyPaths(
        staging_workbook=staging_dir / f"budget-{month:%Y-%m}.xlsx",
        connection_registry=staging_dir / "connection-registry.json",
        review_workbook=review_dir / f"{month:%B}-{month:%Y}-Personal-Budget-Review.xlsx",
    )


def _reconcile_monthly_states(
    *,
    target_exists: bool,
    target_states: list[SyncItemState],
    registry_states: list[SyncItemState],
    authority_states: list[SyncItemState],
) -> list[SyncItemState]:
    sources = (
        (target_states, registry_states, authority_states)
        if target_exists
        else (registry_states, target_states, authority_states)
    )
    return reconcile_sync_states(*sources)


def _review_result_payload(result) -> dict[str, Any]:
    return {
        "transaction_count": result.transaction_count,
        "category_count": result.category_count,
        "categorized_count": result.categorized_count,
        "uncategorized_count": result.uncategorized_count,
        "evidence_supported_count": result.evidence_supported_count,
        "gmail_link_count": result.gmail_link_count,
        "review_needed_count": result.review_needed_count,
        "excluded_non_spend_count": result.excluded_non_spend_count,
        "treatment_review_needed_count": result.treatment_review_needed_count,
        "backup_created": bool(result.backup_path),
        "source_mode": result.source_mode,
        "source_coverage_complete": result.source_coverage_complete,
        "source_expected_count": result.source_expected_count,
        "source_complete_count": result.source_complete_count,
        "source_partial_count": result.source_partial_count,
        "source_unavailable_count": result.source_unavailable_count,
        "reconciliation_removed_visible_count": result.reconciliation_removed_visible_count,
        "reconciliation_ambiguous_count": result.reconciliation_ambiguous_count,
    }


def _retain_month(path: Path, month: date) -> None:
    workbook = BudgetWorkbook(path)
    sheet = workbook.workbook["Actual Charges"]
    headers = workbook._headers(sheet)
    date_col = headers["Date"]
    for row in range(sheet.max_row, 5, -1):
        value = _as_date(sheet.cell(row, date_col).value)
        if value and (value.year, value.month) != (month.year, month.month):
            sheet.delete_rows(row)
    workbook.save_with_backup(backup=True)


def _as_date(value: object) -> date | None:
    """Accept the date values Excel and Plaid workbook imports commonly contain."""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        try:
            return datetime.fromisoformat(value).date()
        except ValueError:
            return None
    return None


def _fake_sync_payloads(path: Path) -> list[dict[str, Any]]:
    """Read one or many fake sync results used only by the automated test seam."""
    payload = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(payload, list):
        return payload
    if isinstance(payload, dict) and isinstance(payload.get("items"), list):
        return payload["items"]
    if isinstance(payload, dict):
        return [payload]
    raise ValueError("sync JSON must be an object, a list, or an object with an items list")


def _parse_month(value: str | None) -> date:
    if not value:
        today = date.today()
        return date(today.year, today.month, 1)
    return datetime.strptime(value, "%Y-%m").date().replace(day=1)


def _telemetry(workbook_path: Path) -> Telemetry:
    return Telemetry(workbook_path.parent / "logs" / "personal-budget-plaid-excel.jsonl")


def _plaid_config() -> PlaidConfig:
    return PlaidConfig.from_env(MacOSKeychain())


def _env(name: str) -> str:
    import os

    return os.getenv(name, "")


if __name__ == "__main__":
    raise SystemExit(main())
