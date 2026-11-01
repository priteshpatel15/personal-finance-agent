from __future__ import annotations

import tempfile
import unittest
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation

from personal_budget_plaid_excel.domain.models import PlaidSyncResult
from personal_budget_plaid_excel.infrastructure.workbook import (
    BASE_HEADERS,
    CATEGORY_MAP_SHEET,
    MAX_FORMULA_ROWS,
    BudgetWorkbook,
)


class WorkbookFormulaRobustnessTests(unittest.TestCase):
    def test_monthly_budget_preserves_noncontiguous_manual_budget_source_rows(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "synthetic-budget.xlsx"
            _write_synthetic_workbook(
                path,
                manual_rows=[
                    ("Flexible", "Coffee", 100),
                    (None, None, None),
                    ("Section", "Transportation", None),
                    ("Fixed", "Car Payment", 300),
                    ("Budget", None, 400),
                ],
            )

            workbook = BudgetWorkbook(path)
            workbook.ensure_structure(as_of=date(2026, 8, 1))
            monthly = workbook.workbook["Monthly Budget"]

            self.assertEqual(monthly["A6"].value, "='Manual Budget'!A2")
            self.assertEqual(monthly["B6"].value, "='Manual Budget'!B2")
            self.assertEqual(monthly["C6"].value, "='Manual Budget'!C2")
            self.assertEqual(monthly["A7"].value, "='Manual Budget'!A5")
            self.assertEqual(monthly["B7"].value, "='Manual Budget'!B5")
            self.assertEqual(monthly["C7"].value, "='Manual Budget'!C5")

    def test_monthly_formulas_follow_reordered_headers_and_current_extent_beyond_20000(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "synthetic-budget.xlsx"
            reordered_headers = [
                "Date",
                "Amount",
                "Merchant",
                "Account",
                "Saved Category",
                "Cost Type",
                "Month",
                "Selected Category",
                "Cancel?",
                "Refund?",
                "Won't Recur?",
                "Evidence",
                "Personal Budget?",
            ]
            _write_synthetic_workbook(path, actual_headers=reordered_headers)
            raw = BudgetWorkbook(path)
            actual = raw.workbook["Actual Charges"]
            initial_headers = _headers(actual)
            actual.cell(20001, initial_headers["Date"]).value = date(2026, 8, 15)
            actual.cell(20001, initial_headers["Month"]).value = date(2026, 8, 1)
            actual.cell(20001, initial_headers["Merchant"]).value = "Synthetic Merchant"
            actual.cell(20001, initial_headers["Amount"]).value = 10
            actual.cell(20001, initial_headers["Personal Budget?"]).value = "Yes"
            actual.cell(20001, initial_headers["Selected Category"]).value = "Coffee"

            raw.ensure_structure(as_of=date(2026, 8, 1))
            headers = _headers(actual)
            posted_formula = raw.workbook["Monthly Budget"]["D6"].value
            review_formula = raw.workbook["Monthly Budget"]["H6"].value

            for header in (
                "Selected Category",
                "Month",
                "Personal Budget?",
                "Plaid Removed?",
                "Budget Treatment",
                "Plaid Status",
                "Amount",
            ):
                self.assertIn(_range_for(headers, header, 20001), posted_formula)
            self.assertIn(_range_for(headers, "Suggested Category", 20001), review_formula)
            self.assertNotIn("$20000", posted_formula)

    def test_appending_a_sync_row_refreshes_formula_extent(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "synthetic-budget.xlsx"
            _write_synthetic_workbook(path)
            workbook = BudgetWorkbook(path)
            workbook.ensure_structure(as_of=date(2026, 8, 1))
            actual = workbook.workbook["Actual Charges"]
            headers = _headers(actual)
            _write_existing_transaction(actual, headers, row=6)
            workbook.ensure_structure(as_of=date(2026, 8, 1))
            self.assertIn("$6:$", workbook.workbook["Monthly Budget"]["D6"].value)

            workbook.apply_plaid_sync(
                PlaidSyncResult(
                    item_id="synthetic-item",
                    added=[
                        {
                            "transaction_id": "new-row",
                            "account_id": "synthetic-account",
                            "amount": 12.34,
                            "date": "2026-08-16",
                            "merchant_name": "Synthetic New Merchant",
                            "pending": False,
                        }
                    ],
                    modified=[],
                    removed=[],
                    accounts=[{"account_id": "synthetic-account", "name": "Synthetic Card"}],
                    next_cursor="synthetic-cursor",
                    request_ids=["synthetic-request"],
                )
            )

            self.assertEqual(actual.max_row, 7)
            self.assertIn("$7", workbook.workbook["Monthly Budget"]["D6"].value)

    def test_total_percent_used_is_total_posted_divided_by_total_budget(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "synthetic-budget.xlsx"
            _write_synthetic_workbook(
                path,
                manual_rows=[
                    ("Flexible", "Coffee", 100),
                    ("Fixed", "Car Payment", 300),
                    ("Budget", None, 400),
                ],
            )

            workbook = BudgetWorkbook(path)
            workbook.ensure_structure(as_of=date(2026, 8, 1))
            monthly = workbook.workbook["Monthly Budget"]

            self.assertEqual(monthly["A8"].value, "TOTAL")
            self.assertEqual(monthly["G8"].value, "=IF(C8=0,0,D8/C8)")
            self.assertNotIn("SUM(", monthly["G8"].value)

    def test_category_dropdown_migration_replaces_stale_managed_validations(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "synthetic-budget.xlsx"
            _write_synthetic_workbook(path)
            workbook = BudgetWorkbook(path)
            workbook.ensure_structure(as_of=date(2026, 8, 1))
            actual = workbook.workbook["Actual Charges"]
            headers = _headers(actual)
            selected_letter = _column_letter(headers["Selected Category"])

            managed = actual.data_validations.dataValidation[0]
            managed.formula1 = f"='{CATEGORY_MAP_SHEET}'!$E$2:$E$99"
            managed.sqref = f"{selected_letter}6:{selected_letter}999"
            duplicate = DataValidation(
                type="list",
                formula1=f"='{CATEGORY_MAP_SHEET}'!$E$2:$E$5",
                allow_blank=True,
            )
            duplicate.promptTitle = "Budget category"
            actual.add_data_validation(duplicate)
            duplicate.add(f"{selected_letter}6:{selected_letter}500")
            unrelated = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
            actual.add_data_validation(unrelated)
            unrelated.add("A6:A20")

            workbook.ensure_structure(as_of=date(2026, 8, 1))

            managed_after = [
                validation
                for validation in actual.data_validations.dataValidation
                if CATEGORY_MAP_SHEET in str(validation.formula1)
                or validation.promptTitle == "Budget category"
            ]
            self.assertEqual(len(managed_after), 1)
            self.assertEqual(
                managed_after[0].formula1,
                f"='{CATEGORY_MAP_SHEET}'!$E$2:$E$4",
            )
            self.assertEqual(
                str(managed_after[0].sqref),
                f"{selected_letter}6:{selected_letter}{MAX_FORMULA_ROWS}",
            )
            self.assertIn(unrelated, actual.data_validations.dataValidation)


def _write_synthetic_workbook(
    path: Path,
    *,
    manual_rows: list[tuple[object, object, object]] | None = None,
    actual_headers: list[str] | None = None,
) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)
    manual = workbook.create_sheet("Manual Budget")
    manual.append(["Cost Type", "Category", "Budget"])
    for row in manual_rows or [
        ("Flexible", "Coffee", 100),
        ("Flexible", "Shopping", 200),
        ("Fixed", "Car Payment", 300),
        ("Budget", None, 600),
    ]:
        manual.append(list(row))
    actual = workbook.create_sheet("Actual Charges")
    for column, header in enumerate(actual_headers or BASE_HEADERS, start=1):
        actual.cell(5, column).value = header
    workbook.save(path)


def _write_existing_transaction(sheet, headers: dict[str, int], *, row: int) -> None:
    values = {
        "Date": date(2026, 8, 1),
        "Month": date(2026, 8, 1),
        "Merchant": "Synthetic Existing Merchant",
        "Amount": 5,
        "Personal Budget?": "Yes",
        "Selected Category": "Coffee",
        "Plaid Status": "posted",
        "Plaid Removed?": "No",
    }
    for header, value in values.items():
        sheet.cell(row, headers[header]).value = value


def _headers(sheet) -> dict[str, int]:
    return {
        str(sheet.cell(5, column).value): column
        for column in range(1, sheet.max_column + 1)
        if sheet.cell(5, column).value
    }


def _column_letter(column: int) -> str:
    result = ""
    while column:
        column, remainder = divmod(column - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _range_for(headers: dict[str, int], header: str, end_row: int) -> str:
    letter = _column_letter(headers[header])
    return f"'Actual Charges'!${letter}$6:${letter}${end_row}"


if __name__ == "__main__":
    unittest.main()
