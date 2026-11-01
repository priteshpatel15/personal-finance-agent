from __future__ import annotations

import unittest
from datetime import date
from decimal import Decimal

from personal_budget_plaid_excel.domain.category import build_merchant_history, infer_suggestion
from personal_budget_plaid_excel.domain.models import LedgerTransaction
from personal_budget_plaid_excel.domain.non_spend import (
    TreatmentDisposition,
    TreatmentKind,
    classify_budget_treatments,
)


class NonSpendTreatmentTests(unittest.TestCase):
    def test_strong_credit_card_payment_is_excluded(self) -> None:
        [treatment] = classify_budget_treatments([
            _transaction(
                "AUTOPAY PAYMENT",
                "500",
                pfc_primary="LOAN_PAYMENTS",
                pfc_detailed="LOAN_PAYMENTS_CREDIT_CARD_PAYMENT",
                confidence="HIGH",
            )
        ])

        self.assertEqual(treatment.disposition, TreatmentDisposition.EXCLUDED)
        self.assertEqual(treatment.kind, TreatmentKind.CREDIT_CARD_PAYMENT)
        self.assertTrue(treatment.excluded_from_budget)
        self.assertTrue(treatment.suppress_category_suggestion)

    def test_same_account_payment_cluster_promotes_low_confidence_variants(self) -> None:
        transactions = [
            _transaction(
                f"CARD PAYMENT {index}",
                str(100 + index),
                pfc_primary="LOAN_PAYMENTS",
                pfc_detailed=(
                    "LOAN_PAYMENTS_CREDIT_CARD_PAYMENT"
                    if index < 4
                    else "LOAN_PAYMENTS_OTHER_PAYMENT"
                ),
                confidence="HIGH" if index < 4 else "LOW",
            )
            for index in range(6)
        ]

        treatments = classify_budget_treatments(transactions)

        self.assertEqual(
            [treatment.disposition for treatment in treatments],
            [TreatmentDisposition.EXCLUDED] * 6,
        )
        self.assertTrue(all(treatment.kind is TreatmentKind.CREDIT_CARD_PAYMENT for treatment in treatments))

    def test_opposing_account_transfer_pair_is_excluded(self) -> None:
        treatments = classify_budget_treatments([
            _transaction(
                "ONLINE TRANSFER TO SAVINGS",
                "250",
                account="Checking",
                pfc_primary="TRANSFER_OUT",
                pfc_detailed="TRANSFER_OUT_ACCOUNT_TRANSFER",
                confidence="HIGH",
            ),
            _transaction(
                "TRANSFER FROM CHECKING",
                "-250",
                account="Savings",
                pfc_primary="TRANSFER_IN",
                pfc_detailed="TRANSFER_IN_ACCOUNT_TRANSFER",
                confidence="HIGH",
            ),
        ])

        self.assertTrue(all(treatment.disposition is TreatmentDisposition.EXCLUDED for treatment in treatments))
        self.assertTrue(all(treatment.kind is TreatmentKind.INTERNAL_TRANSFER for treatment in treatments))

    def test_refunds_income_loan_payments_and_similar_word_purchases_are_distinct(self) -> None:
        treatments = classify_budget_treatments([
            _transaction(
                "Synthetic Store Refund",
                "-25",
                selected_category="Shopping",
                pfc_primary="GENERAL_MERCHANDISE",
            ),
            _transaction(
                "Payroll Credit",
                "-1000",
                pfc_primary="INCOME",
                pfc_detailed="INCOME_SALARY",
                confidence="VERY_HIGH",
            ),
            _transaction(
                "Mortgage Payment",
                "1500",
                selected_category="Mortgage / Housing",
                pfc_primary="LOAN_PAYMENTS",
                pfc_detailed="LOAN_PAYMENTS_MORTGAGE_PAYMENT",
                confidence="VERY_HIGH",
            ),
            _transaction(
                "Payment Depot Market",
                "40",
                selected_category="Shopping",
                pfc_primary="GENERAL_MERCHANDISE",
                confidence="VERY_HIGH",
            ),
            _transaction(
                "Transfer Market Cafe",
                "15",
                selected_category="Food & Dining & Entertainment",
                pfc_primary="FOOD_AND_DRINK",
                confidence="HIGH",
            ),
        ])

        self.assertEqual(treatments[0].disposition, TreatmentDisposition.INCLUDED)
        self.assertEqual(treatments[0].kind, TreatmentKind.REFUND_CREDIT)
        self.assertEqual(treatments[1].disposition, TreatmentDisposition.EXCLUDED)
        self.assertEqual(treatments[1].kind, TreatmentKind.INCOME_CREDIT)
        self.assertEqual(treatments[2].disposition, TreatmentDisposition.INCLUDED)
        self.assertEqual(treatments[2].kind, TreatmentKind.BUDGETED_LOAN_PAYMENT)
        self.assertEqual(treatments[3].disposition, TreatmentDisposition.INCLUDED)
        self.assertEqual(treatments[4].disposition, TreatmentDisposition.INCLUDED)

    def test_ambiguous_payment_transfer_and_negative_direction_are_review_needed(self) -> None:
        treatments = classify_budget_treatments([
            _transaction(
                "PAYMENT",
                "30",
                pfc_primary="LOAN_PAYMENTS",
                pfc_detailed="LOAN_PAYMENTS_OTHER_PAYMENT",
                confidence="LOW",
            ),
            _transaction("APP TRANSFER", "50", pfc_primary="TRANSFER_OUT", confidence="LOW"),
            _transaction("Unknown Credit", "-10"),
        ])

        self.assertTrue(all(treatment.disposition is TreatmentDisposition.REVIEW_NEEDED for treatment in treatments))
        self.assertTrue(all(not treatment.excluded_from_budget for treatment in treatments))
        self.assertTrue(all(treatment.suppress_category_suggestion for treatment in treatments))

    def test_negative_payment_transfer_signals_are_not_treated_as_refunds(self) -> None:
        treatments = classify_budget_treatments([
            _transaction(
                "SYNTHETIC PAYMENT",
                "-30",
                selected_category="Shopping",
                pfc_primary="LOAN_PAYMENTS",
                pfc_detailed="LOAN_PAYMENTS_OTHER_PAYMENT",
                confidence="LOW",
            ),
            _transaction(
                "SYNTHETIC APP TRANSFER",
                "-50",
                pfc_primary="TRANSFER_IN",
                confidence="LOW",
                refund_flag="Yes",
            ),
            _transaction(
                "SYNTHETIC SAVED PAYMENT",
                "-40",
                saved_category="Shopping",
                pfc_primary="LOAN_PAYMENTS",
                pfc_detailed="LOAN_PAYMENTS_OTHER_PAYMENT",
                confidence="LOW",
            ),
            _transaction(
                "Synthetic Store Refund",
                "-20",
                selected_category="Shopping",
                pfc_primary="GENERAL_MERCHANDISE",
                refund_flag="Yes",
            ),
        ])

        self.assertEqual(
            [treatment.disposition for treatment in treatments],
            [
                TreatmentDisposition.REVIEW_NEEDED,
                TreatmentDisposition.REVIEW_NEEDED,
                TreatmentDisposition.REVIEW_NEEDED,
                TreatmentDisposition.INCLUDED,
            ],
        )
        self.assertEqual(
            [treatment.kind for treatment in treatments],
            [
                TreatmentKind.UNKNOWN,
                TreatmentKind.UNKNOWN,
                TreatmentKind.UNKNOWN,
                TreatmentKind.REFUND_CREDIT,
            ],
        )
        self.assertEqual(
            [treatment.reason_code for treatment in treatments],
            [
                "generic_loan_payment",
                "generic_transfer",
                "generic_loan_payment",
                "refund_credit_offset",
            ],
        )

    def test_removed_transaction_cannot_corroborate_an_active_transfer(self) -> None:
        treatments = classify_budget_treatments([
            _transaction(
                "SYNTHETIC APP TRANSFER",
                "250",
                account="Checking",
                pfc_primary="TRANSFER_OUT",
                pfc_detailed="TRANSFER_OUT_TRANSFER_OUT_FROM_APPS",
                confidence="LOW",
            ),
            _transaction(
                "SYNTHETIC ACCOUNT TRANSFER",
                "-250",
                account="Savings",
                pfc_primary="TRANSFER_IN",
                pfc_detailed="TRANSFER_IN_ACCOUNT_TRANSFER",
                confidence="HIGH",
                is_removed=True,
            ),
        ])

        self.assertEqual(treatments[0].disposition, TreatmentDisposition.REVIEW_NEEDED)
        self.assertEqual(treatments[1].disposition, TreatmentDisposition.EXCLUDED)

    def test_removed_rows_neither_promote_clusters_or_history_nor_block_a_unique_pair(self) -> None:
        cluster_treatments = classify_budget_treatments([
            _transaction(
                "CARD PAYMENT",
                "100",
                pfc_primary="LOAN_PAYMENTS",
                pfc_detailed="LOAN_PAYMENTS_CREDIT_CARD_PAYMENT",
                confidence="HIGH",
            ),
            _transaction(
                "CARD PAYMENT",
                "110",
                pfc_primary="LOAN_PAYMENTS",
                pfc_detailed="LOAN_PAYMENTS_OTHER_PAYMENT",
                confidence="LOW",
            ),
            _transaction(
                "CARD PAYMENT 1003",
                "120",
                pfc_primary="LOAN_PAYMENTS",
                pfc_detailed="LOAN_PAYMENTS_CREDIT_CARD_PAYMENT",
                confidence="HIGH",
                is_removed=True,
            ),
        ])
        history_treatments = classify_budget_treatments([
            _transaction(
                "CARD PAYMENT",
                "100",
                transaction_date=date(2026, 8, 1),
                pfc_primary="LOAN_PAYMENTS",
                pfc_detailed="LOAN_PAYMENTS_CREDIT_CARD_PAYMENT",
                confidence="HIGH",
            ),
            _transaction(
                "CARD PAYMENT",
                "110",
                transaction_date=date(2026, 8, 8),
                pfc_primary="LOAN_PAYMENTS",
                pfc_detailed="LOAN_PAYMENTS_CREDIT_CARD_PAYMENT",
                confidence="HIGH",
                is_removed=True,
            ),
            _transaction(
                "CARD PAYMENT",
                "120",
                transaction_date=date(2026, 8, 15),
                pfc_primary="LOAN_PAYMENTS",
                pfc_detailed="LOAN_PAYMENTS_OTHER_PAYMENT",
                confidence="LOW",
            ),
        ])
        pair_treatments = classify_budget_treatments([
            _transaction(
                "SYNTHETIC TRANSFER OUT",
                "250",
                account="Checking",
                pfc_primary="TRANSFER_OUT",
                confidence="LOW",
            ),
            _transaction(
                "SYNTHETIC TRANSFER IN",
                "-250",
                account="Savings",
                pfc_primary="TRANSFER_IN",
                pfc_detailed="TRANSFER_IN_ACCOUNT_TRANSFER",
                confidence="HIGH",
            ),
            _transaction(
                "SYNTHETIC TRANSFER DUPLICATE",
                "-250",
                account="Brokerage",
                pfc_primary="TRANSFER_IN",
                pfc_detailed="TRANSFER_IN_ACCOUNT_TRANSFER",
                confidence="HIGH",
                is_removed=True,
            ),
        ])

        self.assertEqual(cluster_treatments[1].disposition, TreatmentDisposition.REVIEW_NEEDED)
        self.assertEqual(history_treatments[2].disposition, TreatmentDisposition.REVIEW_NEEDED)
        self.assertEqual(history_treatments[2].reason_code, "possible_payment_description")
        self.assertEqual(pair_treatments[0].disposition, TreatmentDisposition.EXCLUDED)
        self.assertEqual(pair_treatments[1].disposition, TreatmentDisposition.EXCLUDED)
        self.assertEqual(pair_treatments[0].kind, TreatmentKind.INTERNAL_TRANSFER)

    def test_same_account_cluster_requires_consistent_description_and_respects_manual_category(self) -> None:
        treatments = classify_budget_treatments([
            _transaction(
                "CARD PAYMENT",
                "100",
                pfc_primary="LOAN_PAYMENTS",
                pfc_detailed="LOAN_PAYMENTS_CREDIT_CARD_PAYMENT",
                confidence="HIGH",
            ),
            _transaction(
                "CARD PAYMENT",
                "110",
                pfc_primary="LOAN_PAYMENTS",
                pfc_detailed="LOAN_PAYMENTS_CREDIT_CARD_PAYMENT",
                confidence="HIGH",
            ),
            _transaction(
                "AUTO LOAN PAYMENT",
                "120",
                pfc_primary="LOAN_PAYMENTS",
                pfc_detailed="LOAN_PAYMENTS_OTHER_PAYMENT",
                confidence="LOW",
            ),
            _transaction(
                "CARD PAYMENT",
                "130",
                selected_category="Shopping",
                pfc_primary="LOAN_PAYMENTS",
                pfc_detailed="LOAN_PAYMENTS_OTHER_PAYMENT",
                confidence="LOW",
            ),
        ])

        self.assertEqual(treatments[0].disposition, TreatmentDisposition.EXCLUDED)
        self.assertEqual(treatments[1].disposition, TreatmentDisposition.EXCLUDED)
        self.assertEqual(treatments[2].disposition, TreatmentDisposition.REVIEW_NEEDED)
        self.assertEqual(treatments[3].disposition, TreatmentDisposition.REVIEW_NEEDED)

    def test_same_account_cluster_does_not_override_a_confirmed_refund(self) -> None:
        treatments = classify_budget_treatments([
            _transaction(
                "CARD PAYMENT REFUND",
                "-100",
                pfc_primary="LOAN_PAYMENTS",
                pfc_detailed="LOAN_PAYMENTS_CREDIT_CARD_PAYMENT",
                confidence="HIGH",
            ),
            _transaction(
                "CARD PAYMENT REFUND",
                "-110",
                pfc_primary="LOAN_PAYMENTS",
                pfc_detailed="LOAN_PAYMENTS_CREDIT_CARD_PAYMENT",
                confidence="HIGH",
            ),
            _transaction(
                "CARD PAYMENT REFUND",
                "-120",
                pfc_primary="LOAN_PAYMENTS",
                pfc_detailed="LOAN_PAYMENTS_OTHER_PAYMENT",
                confidence="LOW",
            ),
        ])

        self.assertEqual(treatments[2].disposition, TreatmentDisposition.INCLUDED)
        self.assertEqual(treatments[2].kind, TreatmentKind.REFUND_CREDIT)
        self.assertEqual(treatments[2].reason_code, "refund_credit_offset")

    def test_confirmed_history_requires_consistent_amount_direction(self) -> None:
        treatments = classify_budget_treatments([
            _transaction(
                "CARD PAYMENT",
                "100",
                transaction_date=date(2026, 8, 1),
                pfc_primary="LOAN_PAYMENTS",
                pfc_detailed="LOAN_PAYMENTS_CREDIT_CARD_PAYMENT",
                confidence="HIGH",
            ),
            _transaction(
                "CARD PAYMENT",
                "110",
                transaction_date=date(2026, 8, 8),
                pfc_primary="LOAN_PAYMENTS",
                pfc_detailed="LOAN_PAYMENTS_CREDIT_CARD_PAYMENT",
                confidence="HIGH",
            ),
            _transaction(
                "CARD PAYMENT",
                "-120",
                transaction_date=date(2026, 8, 15),
                pfc_primary="LOAN_PAYMENTS",
                pfc_detailed="LOAN_PAYMENTS_OTHER_PAYMENT",
                confidence="LOW",
            ),
            _transaction(
                "CARD PAYMENT",
                "130",
                transaction_date=date(2026, 8, 22),
                selected_category="Shopping",
                pfc_primary="LOAN_PAYMENTS",
                pfc_detailed="LOAN_PAYMENTS_OTHER_PAYMENT",
                confidence="LOW",
            ),
        ])

        self.assertEqual(treatments[2].disposition, TreatmentDisposition.REVIEW_NEEDED)
        self.assertEqual(treatments[3].disposition, TreatmentDisposition.REVIEW_NEEDED)

    def test_confirmed_history_still_promotes_an_unmarked_consistent_candidate(self) -> None:
        treatments = classify_budget_treatments([
            _transaction(
                "CARD PAYMENT",
                "100",
                transaction_date=date(2026, 8, 1),
                pfc_primary="LOAN_PAYMENTS",
                pfc_detailed="LOAN_PAYMENTS_CREDIT_CARD_PAYMENT",
                confidence="HIGH",
            ),
            _transaction(
                "CARD PAYMENT",
                "110",
                transaction_date=date(2026, 8, 8),
                pfc_primary="LOAN_PAYMENTS",
                pfc_detailed="LOAN_PAYMENTS_CREDIT_CARD_PAYMENT",
                confidence="HIGH",
            ),
            _transaction(
                "CARD PAYMENT",
                "120",
                transaction_date=date(2026, 8, 15),
                pfc_primary="LOAN_PAYMENTS",
                pfc_detailed="LOAN_PAYMENTS_OTHER_PAYMENT",
                confidence="LOW",
            ),
        ])

        self.assertEqual(treatments[2].disposition, TreatmentDisposition.EXCLUDED)
        self.assertEqual(treatments[2].reason_code, "confirmed_historical_pattern")

    def test_payment_candidates_do_not_receive_or_seed_category_suggestions(self) -> None:
        budgets = {"Shopping"}
        merchant_history = {"autopay payment": "Shopping", "payment depot market": "Shopping"}
        payment = {
            "merchant_name": "AUTOPAY PAYMENT",
            "amount": 500,
            "personal_finance_category": {
                "primary": "LOAN_PAYMENTS",
                "detailed": "LOAN_PAYMENTS_CREDIT_CARD_PAYMENT",
                "confidence_level": "HIGH",
            },
        }
        purchase = {
            "merchant_name": "Payment Depot Market",
            "amount": 40,
            "personal_finance_category": {
                "primary": "GENERAL_MERCHANDISE",
                "detailed": "GENERAL_MERCHANDISE_SUPERSTORES",
                "confidence_level": "HIGH",
            },
        }

        self.assertEqual(infer_suggestion(payment, merchant_history, budgets), (None, ""))
        self.assertEqual(infer_suggestion(purchase, merchant_history, budgets)[0], "Shopping")

        history = build_merchant_history([
            _transaction(
                "AUTOPAY PAYMENT",
                "500",
                selected_category="Shopping",
                pfc_primary="LOAN_PAYMENTS",
                pfc_detailed="LOAN_PAYMENTS_CREDIT_CARD_PAYMENT",
                confidence="HIGH",
            ),
            _transaction(
                "Ordinary Shop",
                "20",
                selected_category="Shopping",
                pfc_primary="GENERAL_MERCHANDISE",
            ),
        ])
        self.assertNotIn("autopay payment", history)
        self.assertEqual(history["ordinary shop"], "Shopping")


def _transaction(
    merchant: str,
    amount: str,
    *,
    account: str = "Test Account",
    selected_category: str | None = None,
    saved_category: str | None = None,
    transaction_date: date = date(2026, 8, 15),
    is_removed: bool = False,
    pfc_primary: str = "",
    pfc_detailed: str = "",
    confidence: str = "",
    refund_flag: str = "No",
) -> LedgerTransaction:
    return LedgerTransaction(
        transaction_date=transaction_date,
        month=date(2026, 8, 1),
        merchant=merchant,
        account=account,
        amount=Decimal(amount),
        personal_budget=True,
        saved_category=saved_category,
        selected_category=selected_category,
        is_removed=is_removed,
        raw={
            "Plaid PFC Primary": pfc_primary,
            "Plaid PFC Detailed": pfc_detailed,
            "Plaid PFC Confidence": confidence,
            "Refund?": refund_flag,
        },
    )


if __name__ == "__main__":
    unittest.main()
