"""Read a local, normalized Direct portal import manifest.

The manifest is produced by an authorized browser collection step.  This reader
does not control a browser or contact an institution; it validates and
normalizes already-collected posted transactions for the monthly review path.
"""

from __future__ import annotations

import calendar
import hashlib
import json
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Mapping

from personal_budget_plaid_excel.domain.source_mode import (
    DIRECT_PORTAL_INSTITUTIONS,
    DirectPortalInstitution,
    ImportSourceMode,
    PortalCoverageStatus,
    PortalImportBatch,
    PortalAmountConvention,
    PortalInstitutionCoverage,
    PortalTransaction,
    PortalTransactionKind,
    ReconciliationClassification,
    ReconciliationCoverageRow,
    ReconciliationItem,
    ReconciliationReport,
    SourceCoverage,
)


class JsonPortalManifest:
    """Validate a versioned, month-scoped Direct portal JSON manifest."""

    def __init__(self, path: Path) -> None:
        self.path = Path(path)

    def read(self) -> PortalImportBatch:
        payload = json.loads(self.path.read_text(encoding="utf-8"))
        if not isinstance(payload, Mapping):
            raise ValueError("portal manifest must be a JSON object")
        schema_version = payload.get("schema_version")
        if schema_version not in (1, 2):
            raise ValueError("unsupported portal manifest schema version")
        if ImportSourceMode.parse(str(payload.get("source") or "")) is not ImportSourceMode.DIRECT_PORTAL:
            raise ValueError("portal manifest source must be direct-portal")

        month = _month(payload.get("month"))
        transactions = _transactions(payload.get("transactions"), month, schema_version)
        coverage = _coverage(payload.get("coverage"), month, transactions)
        reconciliation = _reconciliation(payload.get("reconciliation"))
        return PortalImportBatch(
            month=month,
            transactions=transactions,
            coverage=coverage,
            reconciliation=reconciliation,
        )


def _month(value: object) -> date:
    try:
        parsed = date.fromisoformat(f"{str(value)}-01")
    except ValueError as exc:
        raise ValueError("portal manifest month must use YYYY-MM") from exc
    return date(parsed.year, parsed.month, 1)


def _transactions(value: object, month: date, schema_version: int = 1) -> tuple[PortalTransaction, ...]:
    if not isinstance(value, list):
        raise ValueError("portal manifest transactions must be a list")

    by_local_id: dict[str, PortalTransaction] = {}
    for raw in value:
        if not isinstance(raw, Mapping):
            raise ValueError("portal manifest transaction must be an object")
        if str(raw.get("status") or "").strip().casefold() != "posted":
            raise ValueError("direct portal manifests accept posted transactions only")
        transaction_date = _date(raw.get("date"), "portal transaction date")
        if date(transaction_date.year, transaction_date.month, 1) != month:
            continue
        institution = DirectPortalInstitution.parse(str(raw.get("institution") or "")).value
        merchant = _required_text(raw.get("merchant"), "portal transaction merchant", 300)
        account = _required_text(raw.get("account"), "portal transaction account label", 180)
        amount = _decimal(raw.get("amount"), "portal transaction amount")
        supplied_local_id = str(raw.get("local_id") or "").strip().casefold()
        occurrence = _positive_int(raw.get("occurrence", 1), "portal transaction occurrence")
        transaction = PortalTransaction(
            local_id=(
                supplied_local_id
                or portal_transaction_local_id(
                    institution=institution,
                    transaction_date=transaction_date,
                    merchant=merchant,
                    account=account,
                    amount=amount,
                    occurrence=occurrence,
                )
            ),
            institution=institution,
            transaction_date=transaction_date,
            merchant=merchant,
            account=account,
            amount=amount,
            kind=PortalTransactionKind.parse(raw.get("kind")),
            amount_convention=(
                PortalAmountConvention.parse(raw.get("amount_convention"))
                if schema_version == 2
                else PortalAmountConvention.EXPENSE_POSITIVE
            ),
        )
        prior = by_local_id.get(transaction.local_id)
        if prior is not None and prior != transaction:
            raise ValueError("conflicting local transaction identity in portal manifest")
        by_local_id[transaction.local_id] = transaction

    institution_order = {name: index for index, name in enumerate(DIRECT_PORTAL_INSTITUTIONS)}
    return tuple(sorted(
        by_local_id.values(),
        key=lambda row: (
            row.transaction_date,
            institution_order[row.institution],
            row.account.casefold(),
            row.merchant.casefold(),
            row.amount,
            row.local_id,
        ),
    ))


def portal_transaction_local_id(
    *,
    institution: str,
    transaction_date: date,
    merchant: str,
    account: str,
    amount: Decimal,
    occurrence: int = 1,
) -> str:
    """Return the stable local identity used by Direct portal imports.

    The hash contains only normalized month-ledger fields and is never emitted
    in telemetry. ``occurrence`` distinguishes genuinely repeated identical
    portal postings while allowing repeated pagination captures to deduplicate.
    """

    if occurrence < 1:
        raise ValueError("portal transaction occurrence must be a positive integer")
    payload = [
        "direct-portal-v1",
        DirectPortalInstitution.parse(institution).value,
        transaction_date.isoformat(),
        _identity_text(account),
        _identity_text(merchant),
        _decimal_text(amount),
        occurrence,
    ]
    encoded = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    return hashlib.sha256(encoded.encode("utf-8")).hexdigest()


def _coverage(
    value: object,
    month: date,
    transactions: tuple[PortalTransaction, ...],
) -> SourceCoverage:
    if not isinstance(value, list):
        raise ValueError("portal manifest coverage must be a list")
    supplied: dict[str, PortalInstitutionCoverage] = {}
    observed = {institution: 0 for institution in DIRECT_PORTAL_INSTITUTIONS}
    for transaction in transactions:
        observed[transaction.institution] += 1

    month_end = date(month.year, month.month, calendar.monthrange(month.year, month.month)[1])
    for raw in value:
        if not isinstance(raw, Mapping):
            raise ValueError("portal coverage record must be an object")
        institution = DirectPortalInstitution.parse(str(raw.get("institution") or "")).value
        if institution in supplied:
            raise ValueError("portal coverage institutions must be unique")
        status = PortalCoverageStatus.parse(str(raw.get("status") or ""))
        period_start = _optional_date(raw.get("period_start"), "portal coverage period_start")
        period_end = _optional_date(raw.get("period_end"), "portal coverage period_end")
        declared_count = _nonnegative_int(raw.get("posted_transaction_count"), "posted transaction count")
        reason_code = ""

        if status is PortalCoverageStatus.COMPLETE and (
            period_start is None
            or period_end is None
            or period_start > month
            or period_end < month_end
        ):
            status = PortalCoverageStatus.PARTIAL
            reason_code = "date_coverage_incomplete"
        if status is PortalCoverageStatus.COMPLETE and declared_count != observed[institution]:
            status = PortalCoverageStatus.PARTIAL
            reason_code = "transaction_count_mismatch"
        if status is PortalCoverageStatus.UNAVAILABLE and (declared_count or observed[institution]):
            status = PortalCoverageStatus.PARTIAL
            reason_code = "unavailable_with_transactions"

        supplied[institution] = PortalInstitutionCoverage(
            institution=institution,
            status=status,
            period_start=period_start,
            period_end=period_end,
            posted_transaction_count=observed[institution],
            reason_code=reason_code,
        )

    rows = tuple(
        supplied.get(
            institution,
            PortalInstitutionCoverage(
                institution=institution,
                status=PortalCoverageStatus.UNAVAILABLE,
                period_start=None,
                period_end=None,
                posted_transaction_count=0,
                reason_code="source_not_available",
            ),
        )
        for institution in DIRECT_PORTAL_INSTITUTIONS
    )
    return SourceCoverage(mode=ImportSourceMode.DIRECT_PORTAL, month=month, rows=rows)


def _reconciliation(value: object) -> ReconciliationReport | None:
    if value is None:
        return None
    if not isinstance(value, Mapping):
        raise ValueError("portal reconciliation must be an object")
    coverage_value = value.get("coverage", [])
    item_value = value.get("items", [])
    if not isinstance(coverage_value, list) or not isinstance(item_value, list):
        raise ValueError("portal reconciliation coverage and items must be lists")

    coverage: list[ReconciliationCoverageRow] = []
    for raw in coverage_value:
        if not isinstance(raw, Mapping):
            raise ValueError("reconciliation coverage record must be an object")
        coverage.append(ReconciliationCoverageRow(
            institution=str(raw.get("institution") or ""),
            portal_status=_required_text(raw.get("portal_status"), "portal reconciliation status", 80),
            portal_posted_count=_nonnegative_int(raw.get("portal_posted_count"), "portal posted count"),
            active_plaid_count=_nonnegative_int(raw.get("active_plaid_count"), "active Plaid count"),
            matched_count=_nonnegative_int(raw.get("matched_count"), "matched count"),
            portal_only_count=_nonnegative_int(raw.get("portal_only_count"), "portal-only count"),
            plaid_only_count=_nonnegative_int(raw.get("plaid_only_count"), "Plaid-only count"),
            ambiguous_count=_nonnegative_int(raw.get("ambiguous_count"), "ambiguous count"),
            plaid_removed_portal_visible_count=_nonnegative_int(
                raw.get("plaid_removed_portal_visible_count"),
                "Plaid removed portal-visible count",
            ),
            plaid_only_pending_count=_nonnegative_int(
                raw.get("plaid_only_pending_count", 0),
                "Plaid-only pending count",
            ),
        ))

    items: list[ReconciliationItem] = []
    for raw in item_value:
        if not isinstance(raw, Mapping):
            raise ValueError("reconciliation item must be an object")
        items.append(ReconciliationItem(
            local_id=str(raw.get("local_id") or "").strip().casefold(),
            institution=str(raw.get("institution") or ""),
            transaction_date=_optional_date(raw.get("date"), "reconciliation date"),
            amount=_optional_decimal(raw.get("amount"), "reconciliation amount"),
            classification=ReconciliationClassification.parse(str(raw.get("classification") or "")),
            local_status=_required_text(raw.get("local_status"), "reconciliation local status", 120),
            budget_effect=_required_text(raw.get("budget_effect"), "reconciliation budget effect", 160),
            review_note=_required_text(raw.get("review_note"), "reconciliation review note", 300),
        ))
    return ReconciliationReport(coverage=tuple(coverage), items=tuple(items))


def _date(value: object, label: str) -> date:
    try:
        return date.fromisoformat(str(value or ""))
    except ValueError as exc:
        raise ValueError(f"{label} must use YYYY-MM-DD") from exc


def _optional_date(value: object, label: str) -> date | None:
    if value in (None, ""):
        return None
    return _date(value, label)


def _decimal(value: object, label: str) -> Decimal:
    try:
        parsed = Decimal(str(value))
    except (InvalidOperation, ValueError) as exc:
        raise ValueError(f"{label} must be numeric") from exc
    if not parsed.is_finite():
        raise ValueError(f"{label} must be finite")
    return parsed


def _optional_decimal(value: object, label: str) -> Decimal | None:
    if value in (None, ""):
        return None
    return _decimal(value, label)


def _nonnegative_int(value: object, label: str) -> int:
    if isinstance(value, bool):
        raise ValueError(f"{label} must be a nonnegative integer")
    try:
        parsed = int(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{label} must be a nonnegative integer") from exc
    if parsed < 0 or str(parsed) != str(value):
        raise ValueError(f"{label} must be a nonnegative integer")
    return parsed


def _positive_int(value: object, label: str) -> int:
    parsed = _nonnegative_int(value, label)
    if parsed < 1:
        raise ValueError(f"{label} must be a positive integer")
    return parsed


def _identity_text(value: object) -> str:
    return " ".join(str(value or "").replace("\x00", "").split()).casefold()


def _decimal_text(value: Decimal) -> str:
    return format(value.normalize(), "f") if value else "0"


def _required_text(value: object, label: str, limit: int) -> str:
    cleaned = " ".join(str(value or "").replace("\x00", "").split())
    if not cleaned:
        raise ValueError(f"{label} is required")
    return cleaned[:limit]
