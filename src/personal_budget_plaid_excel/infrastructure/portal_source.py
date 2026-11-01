"""Adapter from normalized Direct portal data to the shared review source port."""

from __future__ import annotations

from datetime import date
from typing import Protocol

from personal_budget_plaid_excel.domain.models import CategoryBudget, LedgerTransaction
from personal_budget_plaid_excel.domain.source_mode import (
    PortalImportBatch,
    PortalTransaction,
    PortalTransactionKind,
    ReconciliationReport,
    ReviewSourceContext,
)


class BudgetAuthority(Protocol):
    def read_budgets(self) -> list[CategoryBudget]: ...


class DirectPortalBudgetSource:
    """Expose one portal-owned month without reading or merging Plaid rows."""

    def __init__(self, authority: BudgetAuthority, batch: PortalImportBatch) -> None:
        self.authority = authority
        self.batch = batch
        self.month = batch.month

    def read_budgets(self) -> list[CategoryBudget]:
        return self.authority.read_budgets()

    def read_transactions(self) -> list[LedgerTransaction]:
        return [_to_ledger(transaction, self.month) for transaction in self.batch.transactions]

    def review_source_context(self) -> ReviewSourceContext:
        return ReviewSourceContext(mode=self.batch.coverage.mode, coverage=self.batch.coverage)

    def review_reconciliation(self) -> ReconciliationReport | None:
        return self.batch.reconciliation


def _to_ledger(transaction: PortalTransaction, month: date) -> LedgerTransaction:
    treatment_evidence = _treatment_evidence(transaction)
    treatment_evidence["source_institution"] = transaction.institution
    return LedgerTransaction(
        transaction_date=transaction.transaction_date,
        month=month,
        merchant=transaction.merchant,
        account=transaction.account,
        amount=transaction.review_amount,
        personal_budget=True,
        is_pending=False,
        is_removed=False,
        source="direct-portal",
        raw=treatment_evidence,
    )


def _treatment_evidence(transaction: PortalTransaction) -> dict[str, object]:
    kind = transaction.kind
    if kind is PortalTransactionKind.UNSPECIFIED:
        return {}
    if kind is PortalTransactionKind.PURCHASE:
        return _pfc("GENERAL_MERCHANDISE", "GENERAL_MERCHANDISE_OTHER_GENERAL_MERCHANDISE", "VERY_HIGH")
    if kind is PortalTransactionKind.CREDIT_CARD_PAYMENT:
        return _pfc("LOAN_PAYMENTS", "LOAN_PAYMENTS_CREDIT_CARD_PAYMENT", "VERY_HIGH")
    if kind is PortalTransactionKind.ACCOUNT_TRANSFER:
        direction = "TRANSFER_OUT" if transaction.review_amount >= 0 else "TRANSFER_IN"
        return _pfc(direction, f"{direction}_ACCOUNT_TRANSFER", "VERY_HIGH")
    if kind is PortalTransactionKind.REFUND:
        evidence = _pfc("GENERAL_MERCHANDISE", "GENERAL_MERCHANDISE_OTHER_GENERAL_MERCHANDISE", "VERY_HIGH")
        evidence["Refund?"] = "Yes"
        return evidence
    if kind is PortalTransactionKind.INCOME:
        return _pfc("INCOME", "INCOME_WAGES", "VERY_HIGH")
    if kind is PortalTransactionKind.CAR_PAYMENT:
        return _pfc("LOAN_PAYMENTS", "LOAN_PAYMENTS_CAR_PAYMENT", "VERY_HIGH")
    if kind is PortalTransactionKind.MORTGAGE_PAYMENT:
        return _pfc("LOAN_PAYMENTS", "LOAN_PAYMENTS_MORTGAGE_PAYMENT", "VERY_HIGH")
    if kind is PortalTransactionKind.STUDENT_LOAN_PAYMENT:
        return _pfc("LOAN_PAYMENTS", "LOAN_PAYMENTS_STUDENT_LOAN_PAYMENT", "VERY_HIGH")
    if kind in {PortalTransactionKind.BUDGETED_LOAN_PAYMENT, PortalTransactionKind.AMBIGUOUS}:
        return _pfc("LOAN_PAYMENTS", "LOAN_PAYMENTS_OTHER_PAYMENT", "LOW")
    return {}


def _pfc(primary: str, detailed: str, confidence: str) -> dict[str, object]:
    return {
        "personal_finance_category": {
            "primary": primary,
            "detailed": detailed,
            "confidence_level": confidence,
        }
    }
