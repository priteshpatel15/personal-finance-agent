from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Protocol

from personal_budget_plaid_excel.domain.models import CategoryBudget, LedgerTransaction
from personal_budget_plaid_excel.domain.review import PurchaseEvidence, ReviewBuildResult, ReviewTransaction
from personal_budget_plaid_excel.domain.source_mode import ReconciliationReport, ReviewSourceContext


class BudgetReviewSource(Protocol):
    def read_budgets(self) -> list[CategoryBudget]: ...

    def read_transactions(self) -> list[LedgerTransaction]: ...


class PurchaseEvidenceReader(Protocol):
    def read(self) -> list[PurchaseEvidence]: ...


class ReviewWorkbookWriter(Protocol):
    def write(
        self,
        output_path: Path,
        month: date,
        budgets: list[CategoryBudget],
        transactions: list[ReviewTransaction],
        source_context: ReviewSourceContext | None = None,
        reconciliation: ReconciliationReport | None = None,
    ) -> ReviewBuildResult: ...


class ReviewTelemetry(Protocol):
    def event(self, name: str, **payload: object) -> None: ...
