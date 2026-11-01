from __future__ import annotations

import unittest
from datetime import date
from decimal import Decimal

from personal_budget_plaid_excel.domain.budget import monthly_budget_summary
from personal_budget_plaid_excel.domain.category import canonical_budget_category, is_budget_excluded
from personal_budget_plaid_excel.domain.models import CategoryBudget, LedgerTransaction


class DomainBudgetTests(unittest.TestCase):
    def test_bridges_and_exclusions(self) -> None:
        self.assertEqual(canonical_budget_category("Food & Dining"), "Food & Dining & Entertainment")
        self.assertEqual(canonical_budget_category("Insurance"), "Insurance (including car)")
        self.assertTrue(is_budget_excluded("Work Reimbursement"))

    def test_monthly_summary_counts_posted_pending_refunds_and_uncategorized(self) -> None:
        budgets = [
            CategoryBudget("Flexible", "Coffee", Decimal("100")),
            CategoryBudget("Flexible", "Shopping", Decimal("200")),
        ]
        txs = [
            LedgerTransaction(date(2026, 8, 1), date(2026, 8, 1), "Cafe", "Card", Decimal("12"), True, selected_category="Coffee"),
            LedgerTransaction(date(2026, 8, 2), date(2026, 8, 1), "Cafe", "Card", Decimal("-2"), True, selected_category="Coffee"),
            LedgerTransaction(date(2026, 8, 3), date(2026, 8, 1), "Store", "Card", Decimal("30"), True, selected_category="Shopping", is_pending=True),
            LedgerTransaction(date(2026, 8, 4), date(2026, 8, 1), "Unknown", "Card", Decimal("5"), True, suggested_category="Coffee"),
            LedgerTransaction(date(2026, 8, 4), date(2026, 8, 1), "Work", "Bank", Decimal("500"), False, selected_category="Work Reimbursement"),
            LedgerTransaction(date(2026, 7, 31), date(2026, 7, 1), "Old", "Card", Decimal("99"), True, selected_category="Coffee"),
        ]

        summary = monthly_budget_summary(budgets, txs, date(2026, 8, 6))
        coffee = next(row for row in summary.positions if row.category == "Coffee")
        shopping = next(row for row in summary.positions if row.category == "Shopping")

        self.assertEqual(coffee.posted_spend, Decimal("10"))
        self.assertEqual(coffee.pending_impact, Decimal("0"))
        self.assertEqual(coffee.remaining, Decimal("90"))
        self.assertEqual(coffee.review_needed_count, 1)
        self.assertEqual(shopping.posted_spend, Decimal("0"))
        self.assertEqual(shopping.pending_impact, Decimal("30"))
        self.assertEqual(summary.uncategorized_count, 1)


if __name__ == "__main__":
    unittest.main()

