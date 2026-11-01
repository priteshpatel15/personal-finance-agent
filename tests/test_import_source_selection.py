from __future__ import annotations

import io
import hashlib
import json
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from datetime import date, datetime, timezone
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook, load_workbook

from personal_budget_plaid_excel.application.import_source_selection import (
    SelectMonthlyImportSource,
)
from personal_budget_plaid_excel.application.monthly_review import BuildMonthlyReview
from personal_budget_plaid_excel.application.source_coverage import build_plaid_source_context
from personal_budget_plaid_excel.cli.main import main
from personal_budget_plaid_excel.domain.models import CategoryBudget, LedgerTransaction, SyncItemState
from personal_budget_plaid_excel.domain.source_mode import ImportSourceMode, PortalCoverageStatus
from personal_budget_plaid_excel.infrastructure.review_workbook import OpenpyxlReviewWorkbookWriter


class ImportSourceSelectionTests(unittest.TestCase):
    def test_missing_choice_returns_structured_choice_required_decision(self) -> None:
        decision = SelectMonthlyImportSource().execute(None)

        self.assertEqual(decision.status, "source_choice_required")
        self.assertIsNone(decision.source)
        self.assertFalse(decision.data_collection_allowed)
        self.assertEqual(decision.available_sources, tuple(ImportSourceMode))
        self.assertEqual(
            decision.public_payload(),
            {
                "status": "source_choice_required",
                "reason": "monthly_import_source_required",
                "prompt": "Choose an import source before collecting monthly transactions.",
                "choices": [
                    {"id": "plaid", "label": "Plaid"},
                    {"id": "direct-portal", "label": "Direct portal"},
                ],
                "data_collection_started": False,
            },
        )

    def test_each_explicit_choice_is_selected_without_aliasing_modes(self) -> None:
        plaid = SelectMonthlyImportSource().execute("plaid")
        portal = SelectMonthlyImportSource().execute("direct-portal")

        self.assertEqual(plaid.source, ImportSourceMode.PLAID)
        self.assertEqual(portal.source, ImportSourceMode.DIRECT_PORTAL)
        self.assertTrue(plaid.data_collection_allowed)
        self.assertTrue(portal.data_collection_allowed)
        self.assertEqual(plaid.public_payload()["import_source"], {"id": "plaid", "label": "Plaid"})
        self.assertEqual(
            portal.public_payload()["import_source"],
            {"id": "direct-portal", "label": "Direct portal"},
        )

    def test_monthly_run_without_choice_stops_before_creating_outputs(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            authority = root / "authority.xlsx"
            authority.write_bytes(b"not-opened-when-choice-is-missing")
            output = root / "monthly"
            review_output = root / "reviews"
            stdout = io.StringIO()

            with redirect_stdout(stdout):
                code = main([
                    "--workbook",
                    str(authority),
                    "monthly-run",
                    "--month",
                    "2026-08",
                    "--output-dir",
                    str(output),
                    "--review-output-dir",
                    str(review_output),
                ])

            self.assertEqual(code, 2)
            self.assertIn('"status": "source_choice_required"', stdout.getvalue())
            self.assertIn('"data_collection_started": false', stdout.getvalue())
            self.assertFalse(output.exists())
            self.assertFalse(review_output.exists())

    def test_direct_portal_choice_never_falls_back_to_plaid_or_mutates_outputs(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            authority = root / "authority.xlsx"
            authority.write_bytes(b"not-opened-during-browser-handoff")
            output = root / "monthly"
            stdout = io.StringIO()

            with redirect_stdout(stdout):
                code = main([
                    "--workbook",
                    str(authority),
                    "monthly-run",
                    "--month",
                    "2026-08",
                    "--source",
                    "direct-portal",
                    "--output-dir",
                    str(output),
                ])

            self.assertEqual(code, 2)
            text = stdout.getvalue()
            self.assertIn('"status": "source_manifest_required"', text)
            self.assertIn('"id": "direct-portal"', text)
            self.assertIn('"data_collection_started": false', text)
            self.assertIn('"required_argument": "--portal-json"', text)
            self.assertFalse(output.exists())

    def test_direct_portal_manifest_builds_review_without_a_plaid_staging_workbook(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            authority = root / "authority.xlsx"
            workbook = Workbook()
            manual = workbook.active
            manual.title = "Manual Budget"
            manual.append(["Cost Type", "Category", "Budget"])
            manual.append(["Flexible", "Groceries", 100])
            manual.append(["Budget", None, 100])
            workbook.save(authority)
            manifest = root / "portal.json"
            manifest.write_text(json.dumps({
                "schema_version": 1,
                "source": "direct-portal",
                "month": "2026-08",
                "coverage": [
                    _coverage("Bank of America", 1),
                    _coverage("Chase", 0),
                    _coverage("American Express", 0),
                ],
                "transactions": [{
                    "local_id": hashlib.sha256(b"synthetic-portal-row").hexdigest(),
                    "institution": "Bank of America",
                    "date": "2026-08-03",
                    "merchant": "Synthetic Market",
                    "account": "Synthetic Checking",
                    "amount": "10.00",
                    "status": "posted",
                    "kind": "purchase",
                }],
            }), encoding="utf-8")
            output = root / "monthly"
            stdout = io.StringIO()

            with redirect_stdout(stdout):
                code = main([
                    "--workbook", str(authority),
                    "monthly-run", "--source", "direct-portal",
                    "--month", "2026-08",
                    "--portal-json", str(manifest),
                    "--output-dir", str(output),
                ])

            self.assertEqual(code, 0)
            payload = json.loads(stdout.getvalue())
            self.assertEqual(payload["status"], "completed")
            self.assertEqual(payload["import_source"]["id"], "direct-portal")
            self.assertIsNone(payload["monthly_workbook"])
            self.assertEqual(payload["review"]["transaction_count"], 1)
            self.assertTrue(payload["review"]["source_coverage_complete"])
            self.assertFalse((output / "budget-2026-08.xlsx").exists())
            review = load_workbook(output / "August-2026-Personal-Budget-Review.xlsx")
            self.assertEqual(review["Budget Summary"]["H1"].value, "Direct portal")

    def test_direct_portal_never_overwrites_the_authority_workbook(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            authority = root / "August-2026-Personal-Budget-Review.xlsx"
            workbook = Workbook()
            manual = workbook.active
            manual.title = "Manual Budget"
            manual.append(["Cost Type", "Category", "Budget"])
            manual.append(["Flexible", "Groceries", 100])
            workbook.save(authority)
            before = hashlib.sha256(authority.read_bytes()).hexdigest()
            manifest = root / "portal.json"
            manifest.write_text(json.dumps({
                "schema_version": 1,
                "source": "direct-portal",
                "month": "2026-08",
                "coverage": [_coverage(name, 0) for name in (
                    "Bank of America", "Chase", "American Express"
                )],
                "transactions": [],
            }), encoding="utf-8")

            with redirect_stdout(io.StringIO()), redirect_stderr(io.StringIO()):
                code = main([
                    "--workbook", str(authority),
                    "monthly-run", "--source", "direct-portal",
                    "--month", "2026-08",
                    "--portal-json", str(manifest),
                    "--output-dir", str(root),
                ])

            self.assertEqual(code, 1)
            self.assertEqual(hashlib.sha256(authority.read_bytes()).hexdigest(), before)

    def test_plaid_coverage_uses_post_refresh_state_without_exposing_connection_ids(self) -> None:
        synced_at = datetime(2026, 8, 29, tzinfo=timezone.utc)
        states = [
            SyncItemState("secret-boa", "Bank of America", last_successful_sync=synced_at),
            SyncItemState("secret-chase", "Chase", last_successful_sync=synced_at),
            SyncItemState("secret-amex", "American Express", last_successful_sync=synced_at),
        ]
        transactions = [
            _ledger_transaction("secret-boa"),
            _ledger_transaction("secret-boa", removed=True),
            _ledger_transaction("secret-chase"),
        ]

        context = build_plaid_source_context(date(2026, 8, 1), states, transactions)

        self.assertIs(context.mode, ImportSourceMode.PLAID)
        self.assertTrue(context.coverage.complete)
        self.assertEqual(context.coverage.complete_count, 3)
        self.assertEqual(
            [row.posted_transaction_count for row in context.coverage.rows],
            [1, 1, 0],
        )
        self.assertEqual(
            [row.period_end for row in context.coverage.rows],
            [date(2026, 8, 29)] * 3,
        )
        self.assertNotIn("secret-", repr(context))

    def test_plaid_coverage_marks_mixed_error_partial_and_missing_unavailable(self) -> None:
        synced_at = datetime(2026, 8, 29, tzinfo=timezone.utc)
        states = [
            SyncItemState("boa-good", "Bank of America", last_successful_sync=synced_at),
            SyncItemState("boa-error", "Bank of America", error="SyntheticError"),
            SyncItemState("chase-good", "Chase", last_successful_sync=synced_at),
        ]

        context = build_plaid_source_context(date(2026, 8, 1), states, [])

        self.assertFalse(context.coverage.complete)
        self.assertEqual(
            [row.status for row in context.coverage.rows],
            [
                PortalCoverageStatus.PARTIAL,
                PortalCoverageStatus.COMPLETE,
                PortalCoverageStatus.UNAVAILABLE,
            ],
        )
        self.assertEqual(context.coverage.partial_count, 1)
        self.assertEqual(context.coverage.unavailable_count, 1)

    def test_plaid_coverage_does_not_claim_a_future_month_before_refresh(self) -> None:
        synced_at = datetime(2026, 8, 29, tzinfo=timezone.utc)
        states = [
            SyncItemState("boa", "Bank of America", last_successful_sync=synced_at),
            SyncItemState("chase", "Chase", last_successful_sync=synced_at),
            SyncItemState("amex", "American Express", last_successful_sync=synced_at),
        ]

        context = build_plaid_source_context(date(2026, 9, 1), states, [])

        self.assertFalse(context.coverage.complete)
        self.assertEqual(context.coverage.partial_count, 3)
        self.assertTrue(all(row.period_start is None for row in context.coverage.rows))
        self.assertTrue(all(row.period_end is None for row in context.coverage.rows))

    def test_plaid_coverage_is_visible_in_the_shared_review_workbook(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "review.xlsx"
            synced_at = datetime(2026, 8, 29, tzinfo=timezone.utc)
            context = build_plaid_source_context(
                date(2026, 8, 1),
                [
                    SyncItemState("one", "Bank of America", last_successful_sync=synced_at),
                    SyncItemState("two", "Chase", last_successful_sync=synced_at),
                    SyncItemState("three", "American Express", last_successful_sync=synced_at),
                ],
                [],
            )

            result = BuildMonthlyReview(OpenpyxlReviewWorkbookWriter()).execute(
                _SyntheticBudgetSource(),
                output,
                date(2026, 8, 1),
                source_context=context,
            )

            summary = load_workbook(output)["Budget Summary"]
            self.assertEqual(result.source_mode, "plaid")
            self.assertTrue(result.source_coverage_complete)
            self.assertEqual(summary["H1"].value, "Plaid")
            self.assertEqual(summary["H2"].value, "Complete")
            self.assertEqual(
                [summary.cell(row, 7).value for row in range(4, 7)],
                ["Bank of America", "Chase", "American Express"],
            )


def _ledger_transaction(item_id: str, *, removed: bool = False) -> LedgerTransaction:
    return LedgerTransaction(
        transaction_date=date(2026, 8, 3),
        month=date(2026, 8, 1),
        merchant="Synthetic Merchant",
        account="Synthetic Account",
        amount=Decimal("1.00"),
        personal_budget=True,
        is_removed=removed,
        raw={"Plaid Item ID": item_id},
    )


def _coverage(institution: str, count: int) -> dict[str, object]:
    return {
        "institution": institution,
        "status": "complete",
        "period_start": "2026-08-01",
        "period_end": "2026-08-31",
        "posted_transaction_count": count,
    }


class _SyntheticBudgetSource:
    def read_budgets(self) -> list[CategoryBudget]:
        return [CategoryBudget("Flexible", "Groceries", Decimal("100"))]

    def read_transactions(self) -> list[LedgerTransaction]:
        return []


if __name__ == "__main__":
    unittest.main()
