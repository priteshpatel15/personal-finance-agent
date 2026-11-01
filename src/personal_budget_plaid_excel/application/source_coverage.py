from __future__ import annotations

from calendar import monthrange
from collections import defaultdict
from datetime import date
from typing import Iterable

from personal_budget_plaid_excel.domain.models import LedgerTransaction, SyncItemState
from personal_budget_plaid_excel.domain.source_mode import (
    DIRECT_PORTAL_INSTITUTIONS,
    DirectPortalInstitution,
    InstitutionCoverage,
    ImportSourceMode,
    PortalCoverageStatus,
    ReviewSourceContext,
    SourceCoverage,
)


def build_plaid_source_context(
    month: date,
    states: Iterable[SyncItemState],
    transactions: Iterable[LedgerTransaction],
) -> ReviewSourceContext:
    """Summarize post-refresh Plaid coverage without retaining connection IDs."""

    selected_month = date(month.year, month.month, 1)
    month_end = date(month.year, month.month, monthrange(month.year, month.month)[1])
    states_by_institution: dict[str, list[SyncItemState]] = defaultdict(list)
    institution_by_item: dict[str, str] = {}
    for state in states:
        try:
            institution = DirectPortalInstitution.parse(state.institution_name).value
        except ValueError:
            continue
        states_by_institution[institution].append(state)
        institution_by_item[state.item_id] = institution

    posted_counts = {institution: 0 for institution in DIRECT_PORTAL_INSTITUTIONS}
    for transaction in transactions:
        if transaction.is_removed or transaction.is_pending:
            continue
        if date(transaction.month.year, transaction.month.month, 1) != selected_month:
            continue
        item_id = str(transaction.raw.get("Plaid Item ID") or "")
        institution = institution_by_item.get(item_id)
        if institution:
            posted_counts[institution] += 1

    rows: list[InstitutionCoverage] = []
    for institution in DIRECT_PORTAL_INSTITUTIONS:
        institution_states = states_by_institution.get(institution, [])
        successful_states = [
            state
            for state in institution_states
            if state.last_successful_sync and not state.error
        ]
        successful_count = len(successful_states)
        failed_count = sum(bool(state.error) for state in institution_states)
        latest_success_date = (
            max(state.last_successful_sync.date() for state in successful_states)
            if successful_states
            else None
        )
        if successful_count and failed_count:
            status = PortalCoverageStatus.PARTIAL
            reason_code = "connection_refresh_partial"
        elif latest_success_date and latest_success_date < selected_month:
            status = PortalCoverageStatus.PARTIAL
            reason_code = "connection_refresh_before_requested_month"
        elif successful_count:
            status = PortalCoverageStatus.COMPLETE
            reason_code = ""
        else:
            status = PortalCoverageStatus.UNAVAILABLE
            reason_code = (
                "connection_refresh_failed" if failed_count else "connection_unavailable"
            )
        coverage_end = (
            min(month_end, latest_success_date)
            if latest_success_date and latest_success_date >= selected_month
            else None
        )
        rows.append(
            InstitutionCoverage(
                institution=institution,
                status=status,
                period_start=selected_month if coverage_end else None,
                period_end=coverage_end,
                posted_transaction_count=posted_counts[institution],
                reason_code=reason_code,
            )
        )

    return ReviewSourceContext(
        mode=ImportSourceMode.PLAID,
        coverage=SourceCoverage(
            mode=ImportSourceMode.PLAID,
            month=selected_month,
            rows=tuple(rows),
        ),
    )
