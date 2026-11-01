from __future__ import annotations

from typing import Protocol

from personal_budget_plaid_excel.domain.models import PlaidSyncResult, SyncItemState


class PlaidPort(Protocol):
    def sync_item(self, state: SyncItemState, access_token: str) -> PlaidSyncResult:
        """Return all transaction updates for an item from its current cursor."""

