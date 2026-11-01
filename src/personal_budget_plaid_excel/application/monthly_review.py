from __future__ import annotations

from datetime import date
from pathlib import Path

from personal_budget_plaid_excel.domain.review import ReviewBuildResult, build_review_transactions
from personal_budget_plaid_excel.domain.source_mode import ReconciliationReport, ReviewSourceContext
from personal_budget_plaid_excel.ports.review import (
    BudgetReviewSource,
    PurchaseEvidenceReader,
    ReviewTelemetry,
    ReviewWorkbookWriter,
)


class BuildMonthlyReview:
    """Build a standalone review artifact from an existing monthly budget source."""

    def __init__(self, writer: ReviewWorkbookWriter) -> None:
        self.writer = writer

    def execute(
        self,
        source: BudgetReviewSource,
        output_path: Path,
        month: date,
        evidence_reader: PurchaseEvidenceReader | None = None,
        telemetry: ReviewTelemetry | None = None,
        source_context: ReviewSourceContext | None = None,
        reconciliation: ReconciliationReport | None = None,
    ) -> ReviewBuildResult:
        source_path = getattr(source, "path", None)
        if source_path is not None and Path(source_path).resolve() == output_path.resolve():
            raise ValueError("review output must be separate from the budget authority workbook")

        selected_month = date(month.year, month.month, 1)
        budgets = source.read_budgets()
        transactions = [
            transaction
            for transaction in source.read_transactions()
            if not transaction.is_removed
            and date(transaction.month.year, transaction.month.month, 1) == selected_month
        ]
        evidence = evidence_reader.read() if evidence_reader else []
        review_transactions = build_review_transactions(transactions, evidence)
        resolved_source_context = source_context or _source_context(source)
        if resolved_source_context and resolved_source_context.coverage.month != selected_month:
            raise ValueError("review source coverage month does not match requested month")
        resolved_reconciliation = reconciliation or _reconciliation(source)
        result = self.writer.write(
            output_path,
            selected_month,
            budgets,
            review_transactions,
            source_context=resolved_source_context,
            reconciliation=resolved_reconciliation,
        )

        if telemetry:
            telemetry.event(
                "review_workbook_built",
                month=selected_month.strftime("%Y-%m"),
                transaction_count=result.transaction_count,
                category_count=result.category_count,
                categorized_count=result.categorized_count,
                uncategorized_count=result.uncategorized_count,
                evidence_supported_count=result.evidence_supported_count,
                gmail_link_count=result.gmail_link_count,
                review_needed_count=result.review_needed_count,
                excluded_non_spend_count=result.excluded_non_spend_count,
                treatment_review_needed_count=result.treatment_review_needed_count,
                backup_created=bool(result.backup_path),
                source_mode=result.source_mode,
                source_coverage_complete=result.source_coverage_complete,
                source_expected_count=result.source_expected_count,
                source_complete_count=result.source_complete_count,
                source_partial_count=result.source_partial_count,
                source_unavailable_count=result.source_unavailable_count,
                reconciliation_removed_visible_count=result.reconciliation_removed_visible_count,
                reconciliation_ambiguous_count=result.reconciliation_ambiguous_count,
            )
        return result


def _source_context(source: BudgetReviewSource) -> ReviewSourceContext | None:
    reader = getattr(source, "review_source_context", None)
    return reader() if callable(reader) else None


def _reconciliation(source: BudgetReviewSource) -> ReconciliationReport | None:
    reader = getattr(source, "review_reconciliation", None)
    return reader() if callable(reader) else None
