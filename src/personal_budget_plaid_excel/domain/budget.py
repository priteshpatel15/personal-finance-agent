from __future__ import annotations

from collections import defaultdict
from datetime import date
from decimal import Decimal

from .category import canonical_budget_category, is_budget_excluded
from .models import CategoryBudget, CategoryPosition, LedgerTransaction, MonthlyBudgetSummary


def month_start(value: date) -> date:
    return date(value.year, value.month, 1)


def monthly_budget_summary(
    budgets: list[CategoryBudget],
    transactions: list[LedgerTransaction],
    month: date,
    *,
    last_synced_at=None,
    sync_status: str = "not synced",
) -> MonthlyBudgetSummary:
    selected_month = month_start(month)
    posted = defaultdict(lambda: Decimal("0"))
    pending = defaultdict(lambda: Decimal("0"))
    review_by_category = defaultdict(int)
    uncategorized = 0

    for tx in transactions:
        if month_start(tx.month) != selected_month:
            continue
        if tx.is_removed or not tx.personal_budget:
            continue

        category = canonical_budget_category(tx.selected_category)
        if not category:
            uncategorized += 1
            suggested = canonical_budget_category(tx.suggested_category)
            if suggested:
                review_by_category[suggested] += 1
            continue
        if is_budget_excluded(category):
            continue
        if tx.is_pending:
            pending[category] += tx.amount
        else:
            posted[category] += tx.amount

    positions = [
        CategoryPosition(
            cost_type=budget.cost_type,
            category=budget.category,
            monthly_budget=budget.monthly_budget,
            posted_spend=posted[budget.category],
            pending_impact=pending[budget.category],
            review_needed_count=review_by_category[budget.category],
        )
        for budget in budgets
    ]
    return MonthlyBudgetSummary(
        month=selected_month,
        positions=positions,
        uncategorized_count=uncategorized,
        last_synced_at=last_synced_at,
        sync_status=sync_status,
    )

