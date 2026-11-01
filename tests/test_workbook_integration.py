from __future__ import annotations

import tempfile
import unittest
from datetime import date
from decimal import Decimal
from pathlib import Path

from openpyxl import load_workbook

from personal_budget_plaid_excel.application.report import HtmlReportService
from personal_budget_plaid_excel.cli.main import _retain_month
from personal_budget_plaid_excel.domain.models import PlaidSyncResult, SyncItemState
from personal_budget_plaid_excel.infrastructure.workbook import BudgetWorkbook

from synthetic_workbook import WORKBOOK_NAME, write_synthetic_workbook
SOURCE_WORKBOOK = Path(WORKBOOK_NAME)


class WorkbookIntegrationTests(unittest.TestCase):
    def test_distinct_plaid_items_keep_distinct_non_secret_sync_records(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            copy = Path(tmp) / SOURCE_WORKBOOK.name
            write_synthetic_workbook(copy)
            workbook = BudgetWorkbook(copy)
            workbook.upsert_sync_state(SyncItemState("item-one", "Institution One", access_token_key="plaid:item-one"))
            workbook.upsert_sync_state(SyncItemState("item-two", "Institution Two", access_token_key="plaid:item-two"))
            workbook.save_with_backup(backup=False)

            states = workbook.read_sync_states()
            self.assertTrue({"item-one", "item-two"}.issubset({state.item_id for state in states}))
            self.assertTrue({"plaid:item-one", "plaid:item-two"}.issubset({state.access_token_key for state in states}))
            sync_sheet = workbook.workbook["_SyncState"]
            headers = [cell.value for cell in sync_sheet[1]]
            self.assertNotIn("Access Token", headers)

    def test_migration_preserves_manual_budget_and_adds_live_budget_surfaces(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            copy = Path(tmp) / SOURCE_WORKBOOK.name
            write_synthetic_workbook(copy)
            before = _manual_budget_snapshot(copy)

            workbook = BudgetWorkbook(copy)
            changed = workbook.ensure_structure(as_of=date(2026, 8, 6))
            backup = workbook.save_with_backup(backup=True)

            self.assertTrue(changed)
            self.assertTrue(backup and backup.exists())
            self.assertEqual(before, _manual_budget_snapshot(copy))

            wb = load_workbook(copy, data_only=False)
            self.assertIn("Monthly Budget", wb.sheetnames)
            self.assertIn("Category Map", wb.sheetnames)
            self.assertEqual(wb["_SyncState"].sheet_state, "hidden")

            actual = wb["Actual Charges"]
            headers = [actual.cell(5, col).value for col in range(1, actual.max_column + 1)]
            self.assertIn("Selected Category", headers)
            self.assertIn("Suggested Category", headers)
            selected_col = headers.index("Selected Category") + 1
            self.assertEqual(actual.cell(6, selected_col).value, "Professional Services")
            insurance_rows = [
                row for row in range(6, actual.max_row + 1)
                if actual.cell(row, headers.index("Saved Category") + 1).value == "Insurance"
            ]
            self.assertTrue(insurance_rows)
            self.assertEqual(actual.cell(insurance_rows[0], selected_col).value, "Insurance (including car)")

            category_map = wb["Category Map"]
            bridge_pairs = {
                (category_map.cell(row, 1).value, category_map.cell(row, 2).value)
                for row in range(2, category_map.max_row + 1)
            }
            self.assertIn(("Food & Dining", "Food & Dining & Entertainment"), bridge_pairs)
            self.assertIn(("Work Reimbursement", None), bridge_pairs)

            monthly = wb["Monthly Budget"]
            self.assertIn("SUMPRODUCT", monthly["D6"].value)
            self.assertEqual(monthly["B2"].value.date(), date(2026, 8, 1))
            self.assertGreater(len(actual.data_validations.dataValidation), 0)

    def test_fake_plaid_sync_appends_blank_selected_category_with_suggestion(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            copy = Path(tmp) / SOURCE_WORKBOOK.name
            write_synthetic_workbook(copy)
            workbook = BudgetWorkbook(copy)
            workbook.ensure_structure(as_of=date(2026, 8, 6))
            starting_rows = workbook.workbook["Actual Charges"].max_row
            result = PlaidSyncResult(
                item_id="item-1",
                added=[
                    {
                        "transaction_id": "tx-new",
                        "account_id": "acct-1",
                        "amount": 4.25,
                        "date": "2026-08-05",
                        "merchant_name": "New Coffee",
                        "pending": False,
                        "personal_finance_category": {
                            "primary": "FOOD_AND_DRINK",
                            "detailed": "FOOD_AND_DRINK_COFFEE",
                            "confidence_level": "VERY_HIGH",
                        },
                    }
                ],
                modified=[],
                removed=[],
                accounts=[{"account_id": "acct-1", "name": "Test Card"}],
                next_cursor="cursor-1",
                request_ids=["request-1"],
            )

            counts = workbook.apply_plaid_sync(result)
            ws = workbook.workbook["Actual Charges"]
            headers = [ws.cell(5, col).value for col in range(1, ws.max_column + 1)]
            row = starting_rows + 1

            self.assertEqual(counts["added"], 1)
            self.assertEqual(ws.cell(row, headers.index("Selected Category") + 1).value, "")
            self.assertEqual(ws.cell(row, headers.index("Suggested Category") + 1).value, "Coffee")
            self.assertEqual(ws.cell(row, headers.index("Plaid Status") + 1).value, "posted")
            self.assertEqual(ws.cell(row, headers.index("Plaid Transaction ID") + 1).value, "tx-new")

    def test_rerun_is_idempotent_and_preserves_existing_selected_category(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            copy = Path(tmp) / SOURCE_WORKBOOK.name
            write_synthetic_workbook(copy)
            workbook = BudgetWorkbook(copy)
            workbook.ensure_structure(as_of=date(2026, 8, 6))
            result = PlaidSyncResult(
                item_id="item-1",
                added=[
                    {
                        "transaction_id": "tx-rerun",
                        "account_id": "acct-1",
                        "amount": 4.25,
                        "date": "2026-08-05",
                        "merchant_name": "New Coffee",
                        "pending": False,
                    }
                ],
                modified=[],
                removed=[],
                accounts=[{"account_id": "acct-1", "name": "Test Card"}],
                next_cursor="cursor-1",
                request_ids=["request-1"],
            )

            workbook.apply_plaid_sync(result)
            ws = workbook.workbook["Actual Charges"]
            headers = [ws.cell(5, col).value for col in range(1, ws.max_column + 1)]
            transaction_col = headers.index("Plaid Transaction ID") + 1
            selected_col = headers.index("Selected Category") + 1
            row = next(row for row in range(6, ws.max_row + 1) if ws.cell(row, transaction_col).value == "tx-rerun")
            ws.cell(row, selected_col).value = "Coffee"

            rerun = workbook.apply_plaid_sync(result)

            self.assertEqual(rerun["added"], 0)
            self.assertEqual(ws.cell(row, selected_col).value, "Coffee")
            self.assertEqual(sum(ws.cell(r, transaction_col).value == "tx-rerun" for r in range(6, ws.max_row + 1)), 1)

    def test_month_filter_keeps_current_month_and_creates_backup(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            copy = Path(tmp) / SOURCE_WORKBOOK.name
            write_synthetic_workbook(copy)
            workbook = BudgetWorkbook(copy)
            workbook.ensure_structure(as_of=date(2026, 8, 1))
            result = PlaidSyncResult(
                item_id="item-1",
                added=[
                    {"transaction_id": "tx-august", "account_id": "acct-1", "amount": 10, "date": "2026-08-05", "merchant_name": "August"},
                    {"transaction_id": "tx-july", "account_id": "acct-1", "amount": 20, "date": "2026-07-05", "merchant_name": "July"},
                ],
                modified=[],
                removed=[],
                accounts=[{"account_id": "acct-1", "name": "Test Card"}],
                next_cursor="cursor-1",
                request_ids=["request-1"],
            )
            workbook.apply_plaid_sync(result)
            workbook.save_with_backup(backup=False)

            _retain_month(copy, date(2026, 8, 1))

            reloaded = BudgetWorkbook(copy)
            ws = reloaded.workbook["Actual Charges"]
            headers = [ws.cell(5, col).value for col in range(1, ws.max_column + 1)]
            transaction_col = headers.index("Plaid Transaction ID") + 1
            retained = {ws.cell(row, transaction_col).value for row in range(6, ws.max_row + 1)}
            self.assertIn("tx-august", retained)
            self.assertNotIn("tx-july", retained)
            self.assertTrue(list((copy.parent / "backups").glob(f"{copy.stem}.*{copy.suffix}")))

    def test_report_generates_html_from_existing_workbook_data(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            copy = Path(tmp) / SOURCE_WORKBOOK.name
            write_synthetic_workbook(copy)
            output = HtmlReportService(copy).write_report(Path(tmp) / "reports", date(2026, 5, 1))
            html = output.read_text(encoding="utf-8")

            self.assertIn("Budget Report: May 2026", html)
            self.assertIn("Monthly Budget", html)
            self.assertIn("Food &amp; Dining &amp; Entertainment", html)


def _manual_budget_snapshot(path: Path):
    wb = load_workbook(path, data_only=False)
    ws = wb["Manual Budget"]
    return tuple(
        tuple(_normalized_cell(ws.cell(row, col).value) for col in range(1, 4))
        for row in range(1, 37)
    )


def _normalized_cell(value):
    if isinstance(value, float):
        return Decimal(str(value)).quantize(Decimal("0.01"))
    return value


if __name__ == "__main__":
    unittest.main()
