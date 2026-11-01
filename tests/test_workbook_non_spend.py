from __future__ import annotations

import tempfile
import unittest
from datetime import date
from pathlib import Path

from openpyxl import Workbook

from personal_budget_plaid_excel.domain.models import PlaidSyncResult
from personal_budget_plaid_excel.domain.non_spend import TreatmentDisposition
from personal_budget_plaid_excel.infrastructure.workbook import BASE_HEADERS, BudgetWorkbook


class WorkbookNonSpendTests(unittest.TestCase):
    def test_credit_card_payment_is_excluded_without_overwriting_manual_category(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "synthetic-budget.xlsx"
            _write_synthetic_workbook(path)
            workbook = BudgetWorkbook(path)
            payment = _transaction(
                "payment-1",
                "2026-08-07",
                "Synthetic Card Payment",
                125,
                primary="LOAN_PAYMENTS",
                detailed="LOAN_PAYMENTS_CREDIT_CARD_PAYMENT",
                original_description="ONLINE PAYMENT THANK YOU",
                transaction_code="payment",
            )

            workbook.apply_plaid_sync(_sync([payment]))
            sheet = workbook.workbook["Actual Charges"]
            headers = _headers(sheet)
            row = _row_for_transaction(sheet, headers, "payment-1")

            self.assertEqual(
                sheet.cell(row, headers["Budget Treatment"]).value,
                TreatmentDisposition.EXCLUDED.value,
            )
            self.assertEqual(sheet.cell(row, headers["Treatment Basis"]).value, "Credit-card payment")
            self.assertEqual(sheet.cell(row, headers["Suggested Category"]).value, "")
            self.assertEqual(sheet.cell(row, headers["Selected Category"]).value, "")
            self.assertEqual(sheet.cell(row, headers["Plaid Original Description"]).value, "ONLINE PAYMENT THANK YOU")
            self.assertEqual(sheet.cell(row, headers["Plaid Transaction Code"]).value, "payment")
            self.assertEqual(sheet.cell(row, headers["Plaid Account Type"]).value, "credit")
            self.assertEqual(sheet.cell(row, headers["Plaid Account Subtype"]).value, "credit card")

            sheet.cell(row, headers["Selected Category"]).value = "Shopping"
            workbook.apply_plaid_sync(_sync([], modified=[payment]))

            self.assertEqual(sheet.cell(row, headers["Selected Category"]).value, "Shopping")
            self.assertEqual(
                sheet.cell(row, headers["Budget Treatment"]).value,
                TreatmentDisposition.EXCLUDED.value,
            )
            ledger_row = next(tx for tx in workbook.read_transactions() if tx.plaid_transaction_id == "payment-1")
            self.assertFalse(ledger_row.personal_budget)
            self.assertIn(TreatmentDisposition.EXCLUDED.value, workbook.workbook["Monthly Budget"]["D6"].value)

    def test_ambiguous_payment_stays_visible_counted_and_uncategorized(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "synthetic-budget.xlsx"
            _write_synthetic_workbook(path)
            workbook = BudgetWorkbook(path)
            ambiguous = _transaction(
                "ambiguous-1",
                "2026-08-08",
                "Synthetic Account Service",
                60,
                primary="LOAN_PAYMENTS",
                detailed="LOAN_PAYMENTS_OTHER_PAYMENT",
            )

            workbook.apply_plaid_sync(_sync([ambiguous]))
            sheet = workbook.workbook["Actual Charges"]
            headers = _headers(sheet)
            row = _row_for_transaction(sheet, headers, "ambiguous-1")

            self.assertEqual(
                sheet.cell(row, headers["Budget Treatment"]).value,
                TreatmentDisposition.REVIEW_NEEDED.value,
            )
            self.assertEqual(sheet.cell(row, headers["Treatment Basis"]).value, "Possible payment/transfer — verify")
            self.assertEqual(sheet.cell(row, headers["Selected Category"]).value, "")
            self.assertEqual(sheet.cell(row, headers["Suggested Category"]).value, "")
            ledger_row = next(tx for tx in workbook.read_transactions() if tx.plaid_transaction_id == "ambiguous-1")
            self.assertTrue(ledger_row.personal_budget)

    def test_unique_opposing_account_transfer_pair_is_excluded_on_both_sides(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "synthetic-budget.xlsx"
            _write_synthetic_workbook(path)
            workbook = BudgetWorkbook(path)
            outgoing = _transaction(
                "transfer-out",
                "2026-08-09",
                "Synthetic Transfer",
                250,
                account_id="checking-1",
                primary="TRANSFER_OUT",
                detailed="TRANSFER_OUT_ACCOUNT_TRANSFER",
                transaction_code="transfer",
            )
            incoming = _transaction(
                "transfer-in",
                "2026-08-10",
                "Synthetic Transfer",
                -250,
                account_id="savings-1",
                primary="TRANSFER_IN",
                detailed="TRANSFER_IN_ACCOUNT_TRANSFER",
                transaction_code="transfer",
            )

            workbook.apply_plaid_sync(_sync([outgoing, incoming], two_accounts=True))
            sheet = workbook.workbook["Actual Charges"]
            headers = _headers(sheet)
            rows = [
                _row_for_transaction(sheet, headers, "transfer-out"),
                _row_for_transaction(sheet, headers, "transfer-in"),
            ]

            self.assertEqual(
                [sheet.cell(row, headers["Budget Treatment"]).value for row in rows],
                [TreatmentDisposition.EXCLUDED.value, TreatmentDisposition.EXCLUDED.value],
            )
            self.assertEqual(
                {sheet.cell(row, headers["Treatment Basis"]).value for row in rows},
                {"Account transfer"},
            )

    def test_refund_and_similar_word_purchases_are_not_false_positive_exclusions(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "synthetic-budget.xlsx"
            _write_synthetic_workbook(path)
            workbook = BudgetWorkbook(path)
            refund = _transaction(
                "refund-1",
                "2026-08-11",
                "Synthetic Payment Systems Refund",
                -20,
                primary="GENERAL_MERCHANDISE",
                detailed="GENERAL_MERCHANDISE_OTHER_GENERAL_MERCHANDISE",
            )
            purchase = _transaction(
                "purchase-1",
                "2026-08-11",
                "Synthetic Transfer Station Cafe",
                20,
                account_id="checking-1",
                primary="FOOD_AND_DRINK",
                detailed="FOOD_AND_DRINK_COFFEE",
            )

            workbook.apply_plaid_sync(_sync([refund, purchase], two_accounts=True))
            sheet = workbook.workbook["Actual Charges"]
            headers = _headers(sheet)
            refund_row = _row_for_transaction(sheet, headers, "refund-1")
            purchase_row = _row_for_transaction(sheet, headers, "purchase-1")

            self.assertEqual(
                sheet.cell(refund_row, headers["Budget Treatment"]).value,
                TreatmentDisposition.INCLUDED.value,
            )
            self.assertEqual(sheet.cell(refund_row, headers["Treatment Basis"]).value, "Refund/credit offset")
            self.assertEqual(
                sheet.cell(purchase_row, headers["Budget Treatment"]).value,
                TreatmentDisposition.INCLUDED.value,
            )
            self.assertEqual(sheet.cell(purchase_row, headers["Suggested Category"]).value, "Coffee")

    def test_generic_transfer_without_unique_opposing_match_remains_review_needed(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "synthetic-budget.xlsx"
            _write_synthetic_workbook(path)
            workbook = BudgetWorkbook(path)
            transfer = _transaction(
                "transfer-review",
                "2026-08-12",
                "Synthetic App Transfer",
                75,
                primary="TRANSFER_OUT",
                detailed="TRANSFER_OUT_TRANSFER_OUT_FROM_APPS",
            )

            workbook.apply_plaid_sync(_sync([transfer]))
            sheet = workbook.workbook["Actual Charges"]
            headers = _headers(sheet)
            row = _row_for_transaction(sheet, headers, "transfer-review")

            self.assertEqual(
                sheet.cell(row, headers["Budget Treatment"]).value,
                TreatmentDisposition.REVIEW_NEEDED.value,
            )
            self.assertEqual(sheet.cell(row, headers["Suggested Category"]).value, "")

            self.assertTrue(
                next(tx for tx in workbook.read_transactions() if tx.plaid_transaction_id == "transfer-review").personal_budget
            )

    def test_negative_generic_payment_is_not_auto_marked_as_a_refund(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "synthetic-budget.xlsx"
            _write_synthetic_workbook(path)
            workbook = BudgetWorkbook(path)
            payment = _transaction(
                "negative-payment-review",
                "2026-08-12",
                "Synthetic Payment",
                -75,
                primary="LOAN_PAYMENTS",
                detailed="LOAN_PAYMENTS_OTHER_PAYMENT",
            )

            workbook.apply_plaid_sync(_sync([payment]))
            sheet = workbook.workbook["Actual Charges"]
            headers = _headers(sheet)
            row = _row_for_transaction(sheet, headers, "negative-payment-review")

            self.assertEqual(sheet.cell(row, headers["Refund?"]).value, "No")
            self.assertEqual(
                sheet.cell(row, headers["Budget Treatment"]).value,
                TreatmentDisposition.REVIEW_NEEDED.value,
            )
            self.assertEqual(sheet.cell(row, headers["Suggested Category"]).value, "")

            sheet.cell(row, headers["Refund?"]).value = "Yes"
            sheet.cell(row, headers["Selected Category"]).value = "Shopping"
            workbook.apply_plaid_sync(_sync([], modified=[payment]))

            self.assertEqual(sheet.cell(row, headers["Refund?"]).value, "Yes")
            self.assertEqual(sheet.cell(row, headers["Selected Category"]).value, "Shopping")
            self.assertEqual(
                sheet.cell(row, headers["Budget Treatment"]).value,
                TreatmentDisposition.REVIEW_NEEDED.value,
            )

    def test_removed_counterpart_no_longer_excludes_active_transfer(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "synthetic-budget.xlsx"
            _write_synthetic_workbook(path)
            workbook = BudgetWorkbook(path)
            active = _transaction(
                "active-transfer",
                "2026-08-12",
                "Synthetic App Transfer",
                250,
                account_id="checking-1",
                primary="TRANSFER_OUT",
                detailed="TRANSFER_OUT_TRANSFER_OUT_FROM_APPS",
            )
            counterpart = _transaction(
                "removed-counterpart",
                "2026-08-12",
                "Synthetic Account Transfer",
                -250,
                account_id="savings-1",
                primary="TRANSFER_IN",
                detailed="TRANSFER_IN_ACCOUNT_TRANSFER",
            )

            workbook.apply_plaid_sync(_sync([active, counterpart], two_accounts=True))
            sheet = workbook.workbook["Actual Charges"]
            headers = _headers(sheet)
            active_row = _row_for_transaction(sheet, headers, "active-transfer")
            self.assertEqual(
                sheet.cell(active_row, headers["Budget Treatment"]).value,
                TreatmentDisposition.EXCLUDED.value,
            )
            workbook.apply_plaid_sync(
                PlaidSyncResult(
                    item_id="synthetic-item",
                    added=[],
                    modified=[],
                    removed=[{"transaction_id": "removed-counterpart"}],
                    accounts=[],
                    next_cursor="synthetic-cursor-2",
                    request_ids=["synthetic-request-2"],
                )
            )
            removed_row = _row_for_transaction(sheet, headers, "removed-counterpart")

            self.assertEqual(
                sheet.cell(active_row, headers["Budget Treatment"]).value,
                TreatmentDisposition.REVIEW_NEEDED.value,
            )
            self.assertEqual(sheet.cell(active_row, headers["Suggested Category"]).value, "")
            self.assertEqual(sheet.cell(removed_row, headers["Plaid Removed?"]).value, "Yes")
            active_ledger = next(
                transaction
                for transaction in workbook.read_transactions()
                if transaction.plaid_transaction_id == "active-transfer"
            )
            self.assertTrue(active_ledger.personal_budget)

    def test_structure_migration_does_not_copy_saved_category_for_review_candidate(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "synthetic-budget.xlsx"
            _write_synthetic_workbook(path)
            workbook = BudgetWorkbook(path)
            workbook.ensure_structure(as_of=date(2026, 8, 1))
            sheet = workbook.workbook["Actual Charges"]
            headers = _headers(sheet)
            review_row = sheet.max_row + 1
            included_row = review_row + 1
            _write_existing_row(
                sheet,
                headers,
                review_row,
                merchant="Synthetic Account Payment",
                saved_category="Shopping",
                primary="LOAN_PAYMENTS",
                detailed="LOAN_PAYMENTS_OTHER_PAYMENT",
            )
            _write_existing_row(
                sheet,
                headers,
                included_row,
                merchant="Synthetic Ordinary Store",
                saved_category="Shopping",
                primary="GENERAL_MERCHANDISE",
                detailed="GENERAL_MERCHANDISE_OTHER_GENERAL_MERCHANDISE",
            )

            workbook.ensure_structure(as_of=date(2026, 8, 1))

            self.assertIn(sheet.cell(review_row, headers["Selected Category"]).value, (None, ""))
            self.assertEqual(sheet.cell(review_row, headers["Suggested Category"]).value, "")
            self.assertEqual(
                sheet.cell(review_row, headers["Budget Treatment"]).value,
                TreatmentDisposition.REVIEW_NEEDED.value,
            )
            self.assertEqual(sheet.cell(included_row, headers["Selected Category"]).value, "Shopping")


def _write_synthetic_workbook(path: Path) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)
    manual = workbook.create_sheet("Manual Budget")
    manual.append(["Cost Type", "Category", "Budget"])
    manual.append(["Flexible", "Coffee", 100])
    manual.append(["Flexible", "Shopping", 200])
    manual.append(["Fixed", "Car Payment", 300])
    manual.append(["Budget", None, 600])
    actual = workbook.create_sheet("Actual Charges")
    for column, header in enumerate(BASE_HEADERS, start=1):
        actual.cell(5, column).value = header
    workbook.save(path)


def _transaction(
    transaction_id: str,
    transaction_date: str,
    merchant: str,
    amount: float,
    *,
    account_id: str = "card-1",
    primary: str,
    detailed: str,
    original_description: str = "",
    transaction_code: str = "",
) -> dict[str, object]:
    return {
        "transaction_id": transaction_id,
        "account_id": account_id,
        "amount": amount,
        "date": transaction_date,
        "merchant_name": merchant,
        "name": merchant,
        "original_description": original_description,
        "transaction_code": transaction_code,
        "pending": False,
        "personal_finance_category": {
            "primary": primary,
            "detailed": detailed,
            "confidence_level": "VERY_HIGH",
        },
    }


def _sync(
    added: list[dict[str, object]],
    *,
    modified: list[dict[str, object]] | None = None,
    two_accounts: bool = False,
) -> PlaidSyncResult:
    accounts = [
        {"account_id": "card-1", "name": "Synthetic Card", "type": "credit", "subtype": "credit card"},
    ]
    if two_accounts:
        accounts.extend([
            {"account_id": "checking-1", "name": "Synthetic Checking", "type": "depository", "subtype": "checking"},
            {"account_id": "savings-1", "name": "Synthetic Savings", "type": "depository", "subtype": "savings"},
        ])
    return PlaidSyncResult(
        item_id="synthetic-item",
        added=added,
        modified=modified or [],
        removed=[],
        accounts=accounts,
        next_cursor="synthetic-cursor",
        request_ids=["synthetic-request"],
    )


def _headers(sheet) -> dict[str, int]:
    return {
        str(sheet.cell(5, column).value): column
        for column in range(1, sheet.max_column + 1)
        if sheet.cell(5, column).value
    }


def _row_for_transaction(sheet, headers: dict[str, int], transaction_id: str) -> int:
    return next(
        row
        for row in range(6, sheet.max_row + 1)
        if sheet.cell(row, headers["Plaid Transaction ID"]).value == transaction_id
    )


def _write_existing_row(
    sheet,
    headers: dict[str, int],
    row: int,
    *,
    merchant: str,
    saved_category: str,
    primary: str,
    detailed: str,
) -> None:
    values = {
        "Date": date(2026, 8, 13),
        "Month": date(2026, 8, 1),
        "Merchant": merchant,
        "Account": "Synthetic Card",
        "Saved Category": saved_category,
        "Amount": 15,
        "Personal Budget?": "Yes",
        "Refund?": "No",
        "Selected Category": "",
        "Suggested Category": saved_category,
        "Suggestion Reason": "Synthetic previous suggestion",
        "Plaid Status": "posted",
        "Plaid Removed?": "No",
        "Plaid PFC Primary": primary,
        "Plaid PFC Detailed": detailed,
        "Plaid PFC Confidence": "HIGH",
    }
    for header, value in values.items():
        sheet.cell(row, headers[header]).value = value


if __name__ == "__main__":
    unittest.main()
