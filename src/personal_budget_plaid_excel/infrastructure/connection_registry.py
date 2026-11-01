"""Durable non-secret registry for monthly Plaid workbooks."""

from __future__ import annotations

import json
import os
from datetime import datetime
from pathlib import Path
from typing import Any

from personal_budget_plaid_excel.application.connection_states import reconcile_sync_states
from personal_budget_plaid_excel.domain.models import SyncItemState


class ConnectionRegistry:
    """Persist the minimum local state required for incremental syncs."""

    def __init__(self, path: Path) -> None:
        self.path = path

    def read(self) -> list[SyncItemState]:
        if not self.path.exists():
            return []
        rows = json.loads(self.path.read_text(encoding="utf-8"))
        if not isinstance(rows, list):
            raise ValueError("connection registry must contain a JSON list")
        return [_state_from_row(row) for row in rows]

    def replace(self, states: list[SyncItemState]) -> None:
        """Atomically replace the registry with sanitized connection state."""

        self.path.parent.mkdir(parents=True, exist_ok=True)
        rows = [_state_to_row(state) for state in states]
        temporary = self.path.with_suffix(f".{os.getpid()}.tmp{self.path.suffix}")
        temporary.write_text(json.dumps(rows, sort_keys=True), encoding="utf-8")
        os.replace(temporary, self.path)

    def merge(self, states: list[SyncItemState]) -> None:
        """Merge incoming state while retaining connections absent from it."""

        self.replace(reconcile_sync_states(states, self.read()))


def _state_from_row(row: Any) -> SyncItemState:
    if not isinstance(row, dict):
        raise ValueError("connection registry rows must be JSON objects")
    return SyncItemState(
        item_id=str(row.get("item_id") or ""),
        institution_name=str(row.get("institution_name") or ""),
        cursor=str(row.get("cursor") or ""),
        access_token_key=str(row.get("access_token_key") or ""),
        last_successful_sync=_parse_datetime(row.get("last_successful_sync")),
        last_request_id="",
        updated_at=_parse_datetime(row.get("updated_at")),
        error="",
    )


def _state_to_row(state: SyncItemState) -> dict[str, Any]:
    return {
        "item_id": state.item_id,
        "institution_name": state.institution_name,
        "cursor": state.cursor,
        "access_token_key": state.access_token_key or f"plaid:{state.item_id}",
        "last_successful_sync": _format_datetime(state.last_successful_sync),
        "last_request_id": "",
        "updated_at": _format_datetime(state.updated_at),
        "error": "",
    }


def _parse_datetime(value: Any) -> datetime | None:
    if isinstance(value, datetime):
        return value
    if not isinstance(value, str) or not value:
        return None
    try:
        return datetime.fromisoformat(value)
    except ValueError:
        return None


def _format_datetime(value: datetime | None) -> str | None:
    return value.isoformat() if value is not None else None
