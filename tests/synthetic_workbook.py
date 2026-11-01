"""Builds a synthetic authority workbook shaped like the real one (fake categories, merchants and amounts)."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

from openpyxl import Workbook

from personal_budget_plaid_excel.domain.models import SyncItemState
from personal_budget_plaid_excel.infrastructure.workbook import BASE_HEADERS, BudgetWorkbook

WORKBOOK_NAME = "budget.xlsx"

MANUAL_BUDGET = [
    ("Fixed", "Mortgage / Housing", 2000),
    ("Fixed", "Insurance (including car)", 250),
    ("Fixed", "Car Payment", 400),
    ("Fixed", "Utilities", 200),
    ("Flexible", "Food & Dining & Entertainment", 600),
    ("Flexible", "Coffee", 60),
    ("Flexible", "Groceries", 500),
    ("Flexible", "Shopping", 300),
    ("Flexible", "Professional Services", 150),
]

# (date, merchant, account, saved category, cost type, amount)
CHARGES = [
    (datetime(2026, 5, 2), "Example Consulting LLC", "Synthetic Card", "Professional Services", "Flexible", 120),
    (datetime(2026, 5, 3), "Example Mutual Insurance", "Synthetic Card", "Insurance", "Fixed", 210.5),
    (datetime(2026, 5, 7), "Sample Bistro", "Synthetic Card", "Food & Dining", "Flexible", 48.25),
    (datetime(2026, 5, 11), "Demo Grocer", "Synthetic Checking", "Groceries", "Flexible", 86.4),
    (datetime(2026, 5, 15), "Employer Expense Co", "Synthetic Card", "Work Reimbursement", "Flexible", 300),
]


def write_synthetic_workbook(path: Path) -> Path:
    workbook = Workbook()
    workbook.remove(workbook.active)
    manual = workbook.create_sheet("Manual Budget")
    manual.append(["Cost Type", "Category", "Budget"])
    for row in MANUAL_BUDGET:
        manual.append(list(row))
    manual.append(["Budget", None, f"=SUM(C2:C{len(MANUAL_BUDGET) + 1})"])

    actual = workbook.create_sheet("Actual Charges")
    actual["A1"] = "Actual Categorized Charges"
    headers = list(BASE_HEADERS)
    for column, header in enumerate(headers, start=1):
        actual.cell(5, column).value = header
    col = {header: index for index, header in enumerate(headers, start=1)}
    for offset, (when, merchant, account, saved, cost_type, amount) in enumerate(CHARGES):
        row = 6 + offset
        values = {
            "Date": when,
            "Month": when.replace(day=1),
            "Merchant": merchant,
            "Account": account,
            "Saved Category": saved,
            "Cost Type": cost_type,
            "Amount": amount,
            "Personal Budget?": "No" if saved == "Work Reimbursement" else "Yes",
            "Cancel?": "No",
            "Refund?": "No",
            "Won't Recur?": "No",
            "Evidence": "Synthetic",
        }
        for header, value in values.items():
            if header in col:
                actual.cell(row, col[header]).value = value
    workbook.save(path)
    # One previously linked (fake) Plaid item, so doctor/refresh paths have a connection to reconcile.
    budget = BudgetWorkbook(path)
    budget.upsert_sync_state(SyncItemState("synthetic-item", "Synthetic Bank", access_token_key="plaid:synthetic-item"))
    budget.save_with_backup(backup=False)
    return path
