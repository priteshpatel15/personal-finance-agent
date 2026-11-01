from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal
from typing import Any


Money = Decimal


@dataclass(frozen=True)
class CategoryBudget:
    cost_type: str
    category: str
    monthly_budget: Money


@dataclass(frozen=True)
class LedgerTransaction:
    transaction_date: date
    month: date
    merchant: str
    account: str
    amount: Money
    personal_budget: bool
    saved_category: str | None = None
    selected_category: str | None = None
    suggested_category: str | None = None
    is_pending: bool = False
    is_removed: bool = False
    source: str = "workbook"
    plaid_transaction_id: str | None = None
    plaid_pending_transaction_id: str | None = None
    raw: dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class CategoryPosition:
    cost_type: str
    category: str
    monthly_budget: Money
    posted_spend: Money
    pending_impact: Money
    review_needed_count: int

    @property
    def remaining(self) -> Money:
        return self.monthly_budget - self.posted_spend

    @property
    def percent_used(self) -> Decimal:
        if self.monthly_budget == 0:
            return Decimal("0")
        return self.posted_spend / self.monthly_budget

    @property
    def status(self) -> str:
        if self.monthly_budget == 0 and self.posted_spend > 0:
            return "No Budget"
        if self.remaining < 0:
            return "Over"
        if self.percent_used >= Decimal("0.85"):
            return "Watch"
        return "OK"


@dataclass(frozen=True)
class MonthlyBudgetSummary:
    month: date
    positions: list[CategoryPosition]
    uncategorized_count: int
    last_synced_at: datetime | None = None
    sync_status: str = "not synced"

    @property
    def total_budget(self) -> Money:
        return sum((p.monthly_budget for p in self.positions), Decimal("0"))

    @property
    def total_posted_spend(self) -> Money:
        return sum((p.posted_spend for p in self.positions), Decimal("0"))

    @property
    def total_pending_impact(self) -> Money:
        return sum((p.pending_impact for p in self.positions), Decimal("0"))

    @property
    def total_remaining(self) -> Money:
        return self.total_budget - self.total_posted_spend


@dataclass(frozen=True)
class SyncItemState:
    item_id: str
    institution_name: str = ""
    cursor: str = ""
    access_token_key: str = ""
    last_successful_sync: datetime | None = None
    last_request_id: str = ""
    updated_at: datetime | None = None
    error: str = ""


@dataclass(frozen=True)
class PlaidSyncResult:
    item_id: str
    added: list[dict[str, Any]]
    modified: list[dict[str, Any]]
    removed: list[dict[str, Any]]
    accounts: list[dict[str, Any]]
    next_cursor: str
    request_ids: list[str]
