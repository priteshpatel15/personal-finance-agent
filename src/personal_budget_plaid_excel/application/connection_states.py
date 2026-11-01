"""Resolve durable Plaid connection state without exposing provider data."""

from __future__ import annotations

from dataclasses import replace
from datetime import datetime, timezone
from typing import Iterable

from personal_budget_plaid_excel.domain.models import SyncItemState


def reconcile_sync_states(*sources: Iterable[SyncItemState]) -> list[SyncItemState]:
    """Return the freshest state for every item across ordered state sources.

    Freshness wins first. When two records are equally fresh, the earlier source
    wins so callers can express context-specific precedence (for example, an
    existing target workbook before its registry and authority workbook).
    Blank display/keychain metadata is filled from the remaining records for
    the same item without replacing the winning cursor or sync timestamps.
    """

    candidates: dict[str, list[tuple[int, SyncItemState]]] = {}
    item_order: list[str] = []
    for source_priority, states in enumerate(sources):
        for state in states:
            if not state.item_id:
                continue
            if state.item_id not in candidates:
                candidates[state.item_id] = []
                item_order.append(state.item_id)
            candidates[state.item_id].append((source_priority, state))

    resolved: list[SyncItemState] = []
    for item_id in item_order:
        records = candidates[item_id]
        _, winner = max(
            records,
            key=lambda entry: (_freshness(entry[1]), -entry[0]),
        )
        fallback_states = [winner, *(state for _, state in records if state is not winner)]
        resolved.append(
            replace(
                winner,
                institution_name=_first_present(fallback_states, "institution_name"),
                access_token_key=(
                    _first_present(fallback_states, "access_token_key") or f"plaid:{item_id}"
                ),
            )
        )
    return resolved


def _freshness(state: SyncItemState) -> datetime:
    timestamps = [
        _utc(timestamp)
        for timestamp in (state.updated_at, state.last_successful_sync)
        if timestamp is not None
    ]
    return max(timestamps, default=datetime.min.replace(tzinfo=timezone.utc))


def _utc(value: datetime) -> datetime:
    if value.tzinfo is None:
        return value.replace(tzinfo=timezone.utc)
    return value.astimezone(timezone.utc)


def _first_present(states: Iterable[SyncItemState], field_name: str) -> str:
    for state in states:
        value = getattr(state, field_name)
        if value:
            return value
    return ""
