from __future__ import annotations

import os
import stat
import tempfile
import unittest
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

from personal_budget_plaid_excel.application.monthly_review import BuildMonthlyReview
from personal_budget_plaid_excel.infrastructure.review_workbook import OpenpyxlReviewWorkbookWriter
from personal_budget_plaid_excel.infrastructure.workbook import BudgetWorkbook
from tests.test_review_workflow import _transaction_rows_by_merchant, _write_synthetic_authority
from tests.test_workbook_non_spend import (
    _headers,
    _row_for_transaction,
    _sync,
    _transaction,
    _write_synthetic_workbook,
)


class WorkbookHardeningTests(unittest.TestCase):
    def test_review_edits_follow_transaction_identity_after_visible_table_sort(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "authority.xlsx"
            output = root / "review.xlsx"
            _write_synthetic_authority(source)
            service = BuildMonthlyReview(OpenpyxlReviewWorkbookWriter())
            service.execute(BudgetWorkbook(source), output, date(2026, 8, 1))

            workbook = load_workbook(output)
            sheet = workbook["Transactions"]
            rows = _transaction_rows_by_merchant(sheet)
            amazon_row = rows["Synthetic Amazon"]
            utility_row = rows["Synthetic Utility"]
            sheet.cell(amazon_row, 5).value = "Shopping"
            sheet.cell(amazon_row, 6).value = "Amazon reviewer note"
            sheet.cell(utility_row, 5).value = "Groceries, Household"
            sheet.cell(utility_row, 6).value = "Utility reviewer note"
            _swap_visible_rows(sheet, amazon_row, utility_row)
            workbook.save(output)

            service.execute(BudgetWorkbook(source), output, date(2026, 8, 1))

            refreshed = load_workbook(output, data_only=False)["Transactions"]
            refreshed_rows = _transaction_rows_by_merchant(refreshed)
            self.assertEqual(refreshed.cell(refreshed_rows["Synthetic Amazon"], 5).value, "Shopping")
            self.assertEqual(
                refreshed.cell(refreshed_rows["Synthetic Amazon"], 6).value,
                "Amazon reviewer note",
            )
            self.assertEqual(
                refreshed.cell(refreshed_rows["Synthetic Utility"], 5).value,
                "Groceries, Household",
            )
            self.assertEqual(
                refreshed.cell(refreshed_rows["Synthetic Utility"], 6).value,
                "Utility reviewer note",
            )
            self.assertTrue(refreshed.column_dimensions["I"].hidden)

    def test_review_escapes_untrusted_formulas_but_preserves_validated_gmail_formula(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "authority.xlsx"
            output = root / "review.xlsx"
            merchant_formula = '=HYPERLINK("https://example.invalid/merchant","merchant")'
            account_formula = '=WEBSERVICE("https://example.invalid/account")'
            unsafe_note_formula = '=WEBSERVICE("https://example.invalid/note")'
            smuggled_formula = (
                '=HYPERLINK("https://mail.google.com/mail/u/0/#all/syntheticFormula123",'
                '"Looks safe")+WEBSERVICE("https://example.invalid/smuggled")'
            )
            gmail_formula = (
                '=HYPERLINK("https://mail.google.com/mail/u/0/#all/syntheticFormula123",'
                '"Validated Gmail detail"&CHAR(10)&"Synthetic item"&CHAR(10)&'
                '"Open Gmail source")'
            )
            _write_synthetic_authority(source)
            authority = load_workbook(source)
            authority["Actual Charges"]["C6"] = merchant_formula
            authority["Actual Charges"]["D6"] = account_formula
            authority.save(source)
            service = BuildMonthlyReview(OpenpyxlReviewWorkbookWriter())
            service.execute(BudgetWorkbook(source), output, date(2026, 8, 1))

            workbook = load_workbook(output, data_only=False)
            transactions = workbook["Transactions"]
            self.assertEqual(transactions["B4"].value, merchant_formula)
            self.assertEqual(transactions["B4"].data_type, "s")
            self.assertEqual(transactions["C4"].value, account_formula)
            self.assertEqual(transactions["C4"].data_type, "s")
            transactions["F4"] = gmail_formula
            transactions["F5"] = unsafe_note_formula
            transactions["F6"] = smuggled_formula
            workbook.save(output)

            service.execute(BudgetWorkbook(source), output, date(2026, 8, 1))

            refreshed = load_workbook(output, data_only=False)
            transactions = refreshed["Transactions"]
            self.assertEqual(transactions["F4"].value, gmail_formula)
            self.assertEqual(transactions["F4"].data_type, "f")
            self.assertEqual(transactions["F5"].value, unsafe_note_formula)
            self.assertEqual(transactions["F5"].data_type, "s")
            self.assertEqual(transactions["F6"].value, smuggled_formula)
            self.assertEqual(transactions["F6"].data_type, "s")
            self.assertEqual(refreshed["Budget Summary"]["C4"].data_type, "f")

    def test_plaid_text_fields_are_stored_as_literal_cells(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "synthetic-budget.xlsx"
            merchant_formula = '=HYPERLINK("https://example.invalid/merchant","merchant")'
            description_formula = '=WEBSERVICE("https://example.invalid/description")'
            _write_synthetic_workbook(path)
            workbook = BudgetWorkbook(path)
            transaction = _transaction(
                "formula-text-1",
                "2026-08-07",
                merchant_formula,
                12,
                primary="GENERAL_MERCHANDISE",
                detailed="GENERAL_MERCHANDISE_OTHER_GENERAL_MERCHANDISE",
                original_description=description_formula,
            )

            workbook.apply_plaid_sync(_sync([transaction]))

            sheet = workbook.workbook["Actual Charges"]
            headers = _headers(sheet)
            row = _row_for_transaction(sheet, headers, "formula-text-1")
            merchant = sheet.cell(row, headers["Merchant"])
            description = sheet.cell(row, headers["Plaid Original Description"])
            self.assertEqual(merchant.value, merchant_formula)
            self.assertEqual(merchant.data_type, "s")
            self.assertEqual(description.value, description_formula)
            self.assertEqual(description.data_type, "s")

    def test_budget_atomic_save_and_backup_are_owner_only(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "synthetic-budget.xlsx"
            _write_synthetic_workbook(path)
            os.chmod(path, 0o600)

            backup = BudgetWorkbook(path).save_with_backup(backup=True)

            self.assertIsNotNone(backup)
            self.assertEqual(_mode(path), 0o600)
            self.assertEqual(_mode(backup), 0o600)
            self.assertEqual(_mode(path.parent / "backups"), 0o700)

    def test_review_output_and_backup_are_owner_only(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "authority.xlsx"
            output = root / "review.xlsx"
            _write_synthetic_authority(source)
            service = BuildMonthlyReview(OpenpyxlReviewWorkbookWriter())

            service.execute(BudgetWorkbook(source), output, date(2026, 8, 1))
            result = service.execute(BudgetWorkbook(source), output, date(2026, 8, 1))

            self.assertIsNotNone(result.backup_path)
            self.assertEqual(_mode(output), 0o600)
            self.assertEqual(_mode(Path(result.backup_path)), 0o600)
            self.assertEqual(_mode(output.parent / "backups"), 0o700)


def _swap_visible_rows(sheet, left: int, right: int) -> None:
    left_values = [sheet.cell(left, column).value for column in range(1, 9)]
    right_values = [sheet.cell(right, column).value for column in range(1, 9)]
    for column, value in enumerate(right_values, start=1):
        sheet.cell(left, column).value = value
    for column, value in enumerate(left_values, start=1):
        sheet.cell(right, column).value = value


def _mode(path: Path) -> int:
    return stat.S_IMODE(path.stat().st_mode)


if __name__ == "__main__":
    unittest.main()
