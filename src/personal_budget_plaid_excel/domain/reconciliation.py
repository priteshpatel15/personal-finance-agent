from __future__ import annotations

from collections import defaultdict
from decimal import Decimal

from personal_budget_plaid_excel.domain.source_mode import (
    ReconciliationItem,
    ReconciliationReport,
)


def merge_reconciliation_reports(
    existing: ReconciliationReport,
    incoming: ReconciliationReport,
) -> ReconciliationReport:
    """Retain the prior queue and append only conservatively identified new rows.

    Opaque source ids are authoritative when they match. Older workbooks did not
    retain those ids, so a unique institution/date/amount bridge is the only
    permitted fallback. Ambiguous bridges and conflicting dispositions fail
    closed before a workbook is replaced.
    """

    existing_by_id = {item.local_id: item for item in existing.items}
    matched_existing_ids: set[str] = set()
    matched_incoming_ids: set[str] = set()

    for item in incoming.items:
        prior = existing_by_id.get(item.local_id)
        if prior is None:
            continue
        if _identity_bridge(prior) != _identity_bridge(item):
            raise ValueError("conflicting reconciliation identity for the same stable key")
        _require_compatible_disposition(prior, item)
        matched_existing_ids.add(prior.local_id)
        matched_incoming_ids.add(item.local_id)

    prior_by_bridge: dict[tuple[str, object, Decimal | None], list[ReconciliationItem]] = defaultdict(list)
    incoming_by_bridge: dict[tuple[str, object, Decimal | None], list[ReconciliationItem]] = defaultdict(list)
    for item in existing.items:
        if item.local_id not in matched_existing_ids:
            prior_by_bridge[_identity_bridge(item)].append(item)
    for item in incoming.items:
        if item.local_id not in matched_incoming_ids:
            incoming_by_bridge[_identity_bridge(item)].append(item)

    for bridge in prior_by_bridge.keys() & incoming_by_bridge.keys():
        prior_candidates = prior_by_bridge[bridge]
        incoming_candidates = incoming_by_bridge[bridge]
        if len(prior_candidates) != 1 or len(incoming_candidates) != 1:
            raise ValueError("ambiguous reconciliation identity bridge")
        prior = prior_candidates[0]
        current = incoming_candidates[0]
        _require_compatible_disposition(prior, current)
        matched_existing_ids.add(prior.local_id)
        matched_incoming_ids.add(current.local_id)

    merged_items = list(existing.items)
    merged_items.extend(
        item for item in incoming.items if item.local_id not in matched_incoming_ids
    )
    coverage_by_institution = {row.institution: row for row in existing.coverage}
    coverage_by_institution.update({row.institution: row for row in incoming.coverage})
    return ReconciliationReport(
        coverage=tuple(coverage_by_institution.values()),
        items=tuple(merged_items),
    )


def _identity_bridge(item: ReconciliationItem) -> tuple[str, object, Decimal | None]:
    amount = item.amount.normalize() if item.amount is not None else None
    return item.institution, item.transaction_date, amount


def _require_compatible_disposition(
    existing: ReconciliationItem,
    incoming: ReconciliationItem,
) -> None:
    existing_disposition = (
        existing.classification,
        existing.local_status.strip(),
        existing.budget_effect.strip(),
    )
    incoming_disposition = (
        incoming.classification,
        incoming.local_status.strip(),
        incoming.budget_effect.strip(),
    )
    if existing_disposition != incoming_disposition:
        raise ValueError("conflicting reconciliation disposition")
