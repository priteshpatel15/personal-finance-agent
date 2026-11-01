from __future__ import annotations

import re
import hashlib
from copy import copy
from dataclasses import dataclass, replace
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from urllib.parse import urlsplit

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils.cell import range_boundaries

from personal_budget_plaid_excel.domain.models import CategoryBudget, LedgerTransaction
from personal_budget_plaid_excel.domain.non_spend import TreatmentDisposition
from personal_budget_plaid_excel.domain.reconciliation import merge_reconciliation_reports
from personal_budget_plaid_excel.domain.review import (
    REVIEW_NEEDED_DEFAULT,
    ReviewBuildResult,
    ReviewCellDetail,
    ReviewTransaction,
    conservative_cross_source_matches,
    _strong_merchant_match,
    _strong_account_institution_match,
    make_review_keys,
    transaction_identity,
)
from personal_budget_plaid_excel.infrastructure.excel_cells import set_literal_cell
from personal_budget_plaid_excel.domain.source_mode import (
    DIRECT_PORTAL_INSTITUTIONS,
    ImportSourceMode,
    PortalCoverageStatus,
    PortalInstitutionCoverage,
    ReconciliationClassification,
    ReconciliationCoverageRow,
    ReconciliationItem,
    ReconciliationReport,
    ReviewSourceContext,
    SourceCoverage,
)
from personal_budget_plaid_excel.infrastructure.private_files import (
    copy_private_file,
    ensure_private_directory,
    private_temporary_path,
    replace_private_file,
)


TRANSACTIONS_SHEET = "Transactions"
BUDGET_SUMMARY_SHEET = "Budget Summary"
REVIEW_STATE_SHEET = "_ReviewState"
RECONCILIATION_SHEET = "Reconciliation"
REVIEW_KEY_HEADER = "Review Key"
REVIEW_KEY_COLUMN = 9
SOURCE_PANEL_START_COLUMN = 7
RECONCILIATION_COVERAGE_TABLE = "ReconciliationCoverageTable"
RECONCILIATION_DETAIL_TABLE = "ReconciliationDetailTable"
PRESERVED_REVIEW_TABLE = "PreservedReviewSelections"
RECONCILIATION_KEY_HEADER = "Reconciliation Key"
RECONCILIATION_BRIDGE_HEADER = "Reconciliation Bridge"

RECONCILIATION_COVERAGE_HEADERS = [
    "Institution",
    "Portal Rows",
    "Matched Active",
    "Plaid Removed / Portal Visible",
    "Ambiguous",
    "Portal Only",
    "Plaid-only Posted",
    "Plaid-only Pending",
    "Portal Coverage",
]
RECONCILIATION_DETAIL_HEADERS = [
    "Institution",
    "Date",
    "Amount",
    "Classification",
    "Local Status",
    "Budget Effect",
    "Review Note",
]
RECONCILIATION_SAFE_FORMULAS = {
    "B5": "=SUM(B9:B11)",
    "E5": "=SUM(C9:C11)",
    "H5": "=SUM(F9:F11)",
    "B6": "=SUM(D9:D11)",
    "E6": "=SUM(E9:E11)",
    "H6": "=SUM(G9:H11)",
    "B12": "=SUM(B9:B11)",
    "C12": "=SUM(C9:C11)",
    "D12": "=SUM(D9:D11)",
    "E12": "=SUM(E9:E11)",
    "F12": "=SUM(F9:F11)",
    "G12": "=SUM(G9:G11)",
    "H12": "=SUM(H9:H11)",
}


@dataclass(frozen=True)
class ExistingReview:
    transaction: LedgerTransaction
    category: str
    detail_value: object
    hyperlink: str | None
    row_height: float | None
    budget_treatment: TreatmentDisposition | None
    treatment_basis: str
    cross_source_bridge: bool = False
    prior_review_key: str | None = None
    reconciliation_disposition: str = "Review needed — not active"
    reconciliation_note: str = "Source identity unresolved; original choice retained, not added to active spending."


class OpenpyxlReviewWorkbookWriter:
    """Write the standalone review artifact without mutating its source workbook."""

    def write(
        self,
        output_path: Path,
        month: date,
        budgets: list[CategoryBudget],
        transactions: list[ReviewTransaction],
        source_context: ReviewSourceContext | None = None,
        reconciliation: ReconciliationReport | None = None,
    ) -> ReviewBuildResult:
        if not budgets:
            raise ValueError("review workbook requires at least one Manual Budget category")
        categories = [budget.category for budget in budgets]
        if len(set(categories)) != len(categories):
            raise ValueError("Manual Budget categories must be unique")

        output_path = output_path.resolve()
        output_path.parent.mkdir(parents=True, exist_ok=True)
        existing = self._read_existing(output_path)
        original_existing = dict(existing)
        if source_context is None:
            source_context = self._read_existing_source_context(output_path)
        if (
            existing
            and source_context is not None
            and source_context.mode is ImportSourceMode.DIRECT_PORTAL
        ):
            existing = self._carry_forward_cross_source_reviews(transactions, existing)
        used_prior_keys = set()
        for transaction in transactions:
            key = transaction.review_key
            prior = existing.get(key)
            if prior is not None:
                used_prior_keys.add(prior.prior_review_key or key)
        preserved = {
            key: review for key, review in original_existing.items()
            if key not in used_prior_keys and (
                review.category or review.hyperlink or _formula_hyperlink(review.detail_value)
                or (review.detail_value and not _is_review_needed(review.detail_value))
            )
        }
        for key, prior in tuple(preserved.items()):
            if prior.reconciliation_note == "Source identity unresolved; original choice retained, not added to active spending.":
                nearby = [(row, current.transaction) for row, current in enumerate(transactions, 4)
                          if current.transaction.amount == prior.transaction.amount
                          and abs((current.transaction.transaction_date - prior.transaction.transaction_date).days) <= 3]
                merchant_matches = [(row, current) for row, current in nearby
                                    if _strong_merchant_match(current.merchant, prior.transaction.merchant)]
                if not nearby:
                    reason = "No active same-signed-amount/date candidate; check original source and existing reconciliation holds."
                elif not merchant_matches:
                    reason = "Amount/date candidates exist; merchant descriptions do not corroborate identity."
                elif not any(_strong_account_institution_match(current, prior.transaction) for _, current in merchant_matches):
                    reason = "Merchant/amount/date candidates exist; account identity is not corroborated."
                else:
                    reason = "Multiple possible source identities; a unique match is not established."
                candidates = ", ".join(str(row) for row, _ in nearby) or "none"
                preserved[key] = replace(prior, reconciliation_note=f"{reason} Transactions candidate rows: {candidates}. Not additional active spending; human identity decision required.")
        incoming_reconciliation = reconciliation
        existing_reconciliation = self._read_existing_reconciliation(output_path)
        if existing_reconciliation is not None:
            reconciliation = (
                existing_reconciliation
                if incoming_reconciliation is None
                else merge_reconciliation_reports(existing_reconciliation, incoming_reconciliation)
            )
        if preserved and reconciliation is None:
            reconciliation = ReconciliationReport(coverage=(), items=())
        backup_path = self._backup(output_path) if output_path.exists() else None
        workbook, counts = self._build(
            month,
            budgets,
            transactions,
            existing,
            source_context,
            reconciliation,
            existing_reconciliation_path=(output_path if existing_reconciliation is not None else None),
            existing_reconciliation_count=(
                len(existing_reconciliation.items) if existing_reconciliation is not None else 0
            ),
            refresh_reconciliation=(
                existing_reconciliation is not None and incoming_reconciliation is not None
            ),
        )

        self._write_preserved_reviews(workbook, preserved)
        with private_temporary_path(output_path) as temporary:
            workbook.save(temporary)
            self._validate_saved(
                temporary,
                len(budgets),
                len(transactions),
                source_context,
                reconciliation,
            )
            reloaded = self._read_existing(temporary)
            for key, prior in preserved.items():
                retained = reloaded.get(key)
                if retained is None or any(getattr(retained, field) != getattr(prior, field) for field in (
                    "transaction", "category", "detail_value", "hyperlink", "budget_treatment", "treatment_basis", "reconciliation_disposition", "reconciliation_note"
                )):
                    raise ValueError("prior review choices were not preserved safely")
            replace_private_file(temporary, output_path)

        return ReviewBuildResult(
            output_path=str(output_path),
            backup_path=str(backup_path) if backup_path else None,
            transaction_count=len(transactions),
            category_count=len(budgets),
            categorized_count=counts["categorized"],
            uncategorized_count=len(transactions) - counts["categorized"],
            evidence_supported_count=counts["supported"],
            gmail_link_count=counts["gmail_links"],
            review_needed_count=counts["review_needed"],
            excluded_non_spend_count=counts["excluded_non_spend"],
            treatment_review_needed_count=counts["treatment_review_needed"],
            source_mode=source_context.mode.value if source_context else "",
            source_coverage_complete=source_context.coverage.complete if source_context else False,
            source_expected_count=source_context.coverage.expected_count if source_context else 0,
            source_complete_count=source_context.coverage.complete_count if source_context else 0,
            source_partial_count=source_context.coverage.partial_count if source_context else 0,
            source_unavailable_count=source_context.coverage.unavailable_count if source_context else 0,
            reconciliation_removed_visible_count=(reconciliation.removed_visible_count if reconciliation else 0),
            reconciliation_ambiguous_count=(reconciliation.ambiguous_count if reconciliation else 0),
        )

    def _carry_forward_cross_source_reviews(
        self,
        transactions: list[ReviewTransaction],
        existing: dict[str, ExistingReview],
    ) -> dict[str, ExistingReview]:
        exact_keys = {
            transaction.review_key
            for transaction in transactions
            if transaction.review_key in existing
        }
        current_records = [
            (transaction.review_key, transaction.transaction)
            for transaction in transactions
            if transaction.review_key not in exact_keys
        ]
        previous_records = [
            (review_key, review)
            for review_key, review in existing.items()
            if review_key not in exact_keys
        ]
        matches = conservative_cross_source_matches(
            [transaction for _, transaction in current_records],
            [review.transaction for _, review in previous_records],
        )
        carried_forward = dict(existing)
        for current_index, previous_index in matches.items():
            current_key, _ = current_records[current_index]
            _, previous_review = previous_records[previous_index]
            carried_forward[current_key] = replace(
                previous_review,
                cross_source_bridge=True,
                prior_review_key=previous_records[previous_index][0],
            )
        return carried_forward

    def _build(
        self,
        month: date,
        budgets: list[CategoryBudget],
        transactions: list[ReviewTransaction],
        existing: dict[str, ExistingReview],
        source_context: ReviewSourceContext | None,
        reconciliation: ReconciliationReport | None,
        existing_reconciliation_path: Path | None = None,
        existing_reconciliation_count: int = 0,
        refresh_reconciliation: bool = False,
    ) -> tuple[Workbook, dict[str, int]]:
        if existing_reconciliation_path is not None:
            workbook = load_workbook(existing_reconciliation_path, data_only=False)
            for existing_sheet in tuple(workbook.worksheets):
                if existing_sheet.title != RECONCILIATION_SHEET:
                    workbook.remove(existing_sheet)
            if "BudgetCategoryList" in workbook.defined_names:
                del workbook.defined_names["BudgetCategoryList"]
            transaction_sheet = workbook.create_sheet(TRANSACTIONS_SHEET, 0)
            summary_sheet = workbook.create_sheet(BUDGET_SUMMARY_SHEET, 1)
            reconciliation_sheet = workbook[RECONCILIATION_SHEET]
            if PRESERVED_REVIEW_TABLE in reconciliation_sheet.tables:
                left, first, right, last = range_boundaries(reconciliation_sheet.tables[PRESERVED_REVIEW_TABLE].ref)
                for merged in tuple(reconciliation_sheet.merged_cells.ranges):
                    if first - 2 <= merged.min_row <= last:
                        reconciliation_sheet.unmerge_cells(str(merged))
                for row in reconciliation_sheet.iter_rows(min_row=first - 2, max_row=last, min_col=left, max_col=right):
                    for cell in row:
                        cell.value = None
                        cell.hyperlink = None
                del reconciliation_sheet.tables[PRESERVED_REVIEW_TABLE]
        else:
            workbook = Workbook()
            workbook.remove(workbook.active)
            transaction_sheet = workbook.create_sheet(TRANSACTIONS_SHEET)
            summary_sheet = workbook.create_sheet(BUDGET_SUMMARY_SHEET)
            reconciliation_sheet = (
                workbook.create_sheet(RECONCILIATION_SHEET)
                if reconciliation is not None
                else None
            )
        state_sheet = workbook.create_sheet(REVIEW_STATE_SHEET)
        state_sheet.sheet_state = "veryHidden"
        state_sheet.append([
            REVIEW_KEY_HEADER,
            RECONCILIATION_KEY_HEADER,
            RECONCILIATION_BRIDGE_HEADER,
        ])

        self._write_transaction_header(transaction_sheet, month)
        category_set = {budget.category for budget in budgets}
        counts = {
            "categorized": 0,
            "supported": 0,
            "gmail_links": 0,
            "review_needed": 0,
            "excluded_non_spend": 0,
            "treatment_review_needed": 0,
        }
        for index, review_transaction in enumerate(transactions, start=4):
            transaction = review_transaction.transaction
            treatment = review_transaction.budget_treatment
            prior = existing.get(review_transaction.review_key)
            if prior is not None:
                category = prior.category if prior.category in category_set or not prior.category else ""
            else:
                category = transaction.selected_category if transaction.selected_category in category_set else ""

            detail, hyperlink, row_height = self._select_detail(review_transaction.evidence_detail, prior)
            displayed_treatment = treatment.disposition
            treatment_basis = treatment.visible_basis
            if prior is not None and prior.budget_treatment is not None:
                displayed_treatment = prior.budget_treatment
                treatment_basis = prior.treatment_basis
            row = [
                transaction.transaction_date,
                transaction.merchant,
                transaction.account,
                float(transaction.amount),
                category,
                detail,
                displayed_treatment.value,
                treatment_basis,
                review_transaction.review_key,
            ]
            for column, value in enumerate(row, start=1):
                cell = transaction_sheet.cell(index, column)
                if column == 6 and _formula_hyperlink(value):
                    cell.value = value
                else:
                    set_literal_cell(cell, value)
            transaction_sheet.cell(index, 1).number_format = "yyyy-mm-dd"
            transaction_sheet.cell(index, 4).number_format = '$#,##0.00;[Red]($#,##0.00);-'
            detail_cell = transaction_sheet.cell(index, 6)
            detail_cell.alignment = Alignment(wrap_text=True, vertical="top")
            validated_hyperlink = _validated_gmail_url(hyperlink)
            if validated_hyperlink:
                detail_cell.hyperlink = validated_hyperlink
                detail_cell.font = Font(color="0563C1", underline="single")
            for column in (7, 8):
                transaction_sheet.cell(index, column).alignment = Alignment(wrap_text=True, vertical="top")
            if displayed_treatment is TreatmentDisposition.EXCLUDED:
                for column in (7, 8):
                    transaction_sheet.cell(index, column).fill = PatternFill("solid", fgColor="E7E6E6")
                counts["excluded_non_spend"] += 1
            elif displayed_treatment is TreatmentDisposition.REVIEW_NEEDED:
                for column in (7, 8):
                    transaction_sheet.cell(index, column).fill = PatternFill("solid", fgColor="FFF2CC")
                counts["treatment_review_needed"] += 1
            transaction_sheet.row_dimensions[index].height = row_height or _detail_row_height(str(detail or ""))
            state_sheet.append([review_transaction.review_key])

            if category:
                counts["categorized"] += 1
            if _is_review_needed(detail):
                counts["review_needed"] += 1
            elif detail:
                counts["supported"] += 1
            if _is_gmail_link(detail, hyperlink):
                counts["gmail_links"] += 1

        last_transaction_row = max(4, len(transactions) + 3)
        if not transactions:
            for column in range(1, REVIEW_KEY_COLUMN + 1):
                transaction_sheet.cell(4, column).value = None

        self._finish_transactions(transaction_sheet, last_transaction_row, month)
        self._write_summary(summary_sheet, month, budgets, last_transaction_row)
        if source_context is not None:
            self._write_source_context(summary_sheet, source_context)
        if reconciliation_sheet is not None and reconciliation is not None:
            if existing_reconciliation_path is None:
                self._write_reconciliation(reconciliation_sheet, month, reconciliation)
            elif refresh_reconciliation:
                self._update_preserved_reconciliation(
                    reconciliation_sheet,
                    reconciliation,
                    existing_reconciliation_count,
                )
            self._ensure_reconciliation_detail_table(
                reconciliation_sheet,
                max(16, 15 + len(reconciliation.items)),
            )
            self._write_reconciliation_state(state_sheet, reconciliation)
        self._add_category_validation(workbook, transaction_sheet, len(budgets), last_transaction_row)

        workbook.calculation.calcMode = "auto"
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        return workbook, counts

    def _write_transaction_header(self, sheet, month: date) -> None:
        sheet.merge_cells("A1:H1")
        sheet["A1"] = f"{month:%B %Y} Transactions"
        sheet["A1"].fill = PatternFill("solid", fgColor="1F4E78")
        sheet["A1"].font = Font(bold=True, color="FFFFFF", size=15)
        sheet["A1"].alignment = Alignment(vertical="center")
        sheet.row_dimensions[1].height = 28
        headers = [
            "Date",
            "Merchant",
            "Credit Card Account",
            "Amount",
            "Category",
            "Purchased Item Detail",
            "Budget Treatment",
            "Treatment Basis",
            REVIEW_KEY_HEADER,
        ]
        for column, header in enumerate(headers, start=1):
            cell = sheet.cell(3, column)
            cell.value = header
            cell.fill = PatternFill("solid", fgColor="0F6B78")
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(vertical="center")
        sheet.row_dimensions[3].height = 24

    def _finish_transactions(self, sheet, last_row: int, month: date) -> None:
        for row in range(4, last_row + 1):
            if row % 2 == 0:
                for column in range(1, REVIEW_KEY_COLUMN + 1):
                    cell = sheet.cell(row, column)
                    if column in (7, 8) and cell.fill.fill_type:
                        continue
                    cell.fill = PatternFill("solid", fgColor="DDEBF7")
            for column in range(1, REVIEW_KEY_COLUMN + 1):
                sheet.cell(row, column).alignment = Alignment(
                    vertical="top",
                    wrap_text=column in (6, 7, 8),
                )

        sheet.column_dimensions["A"].width = 13
        sheet.column_dimensions["B"].width = 30
        sheet.column_dimensions["C"].width = 30
        sheet.column_dimensions["D"].width = 14
        sheet.column_dimensions["E"].width = 31
        sheet.column_dimensions["F"].width = 72
        sheet.column_dimensions["G"].width = 23
        sheet.column_dimensions["H"].width = 32
        sheet.column_dimensions["I"].hidden = True
        sheet.freeze_panes = "A4"
        sheet.sheet_view.showGridLines = False

        table_name = _table_name(f"{month:%B}Transactions")
        table = Table(displayName=table_name, ref=f"A3:I{last_row}")
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        sheet.add_table(table)

    def _write_summary(self, sheet, month: date, budgets: list[CategoryBudget], last_transaction_row: int) -> None:
        sheet.merge_cells("A1:E1")
        sheet["A1"] = f"{month:%B %Y} Budget Summary"
        sheet["A1"].fill = PatternFill("solid", fgColor="1F4E78")
        sheet["A1"].font = Font(bold=True, color="FFFFFF", size=15)
        sheet["A1"].alignment = Alignment(vertical="center")
        sheet.row_dimensions[1].height = 28
        sheet.merge_cells("A2:E2")
        sheet["A2"] = (
            "Actual excludes only rows marked Excluded from Budget; "
            "Review needed rows remain included until verified."
        )
        sheet["A2"].font = Font(italic=True, color="666666", size=10)
        sheet["A2"].alignment = Alignment(wrap_text=True, vertical="center")
        sheet.row_dimensions[2].height = 30
        headers = ["Category", "Manual Budget", "Actual", "Remaining", "Variance"]
        for column, header in enumerate(headers, start=1):
            cell = sheet.cell(3, column)
            cell.value = header
            cell.fill = PatternFill("solid", fgColor="107C10")
            cell.font = Font(bold=True, color="FFFFFF")

        for row, budget in enumerate(budgets, start=4):
            set_literal_cell(sheet.cell(row, 1), budget.category)
            sheet.cell(row, 2).value = float(budget.monthly_budget)
            sheet.cell(row, 3).value = (
                f'=SUMIFS(\'Transactions\'!$D$4:$D${last_transaction_row},'
                f"'Transactions'!$E$4:$E${last_transaction_row},A{row},"
                f"'Transactions'!$G$4:$G${last_transaction_row},"
                f'"<>{TreatmentDisposition.EXCLUDED.value}")'
            )
            sheet.cell(row, 4).value = f"=B{row}-C{row}"
            sheet.cell(row, 5).value = f"=C{row}-B{row}"
            if row % 2 == 0:
                for column in range(1, 6):
                    sheet.cell(row, column).fill = PatternFill("solid", fgColor="E2F0D9")
            for column in range(2, 6):
                sheet.cell(row, column).number_format = '$#,##0.00;[Red]($#,##0.00);-'

        last_category_row = len(budgets) + 3
        total_row = last_category_row + 1
        sheet.cell(total_row, 1).value = "Overall Total"
        for column in range(2, 6):
            letter = chr(64 + column)
            sheet.cell(total_row, column).value = f"=SUM({letter}4:{letter}{last_category_row})"
            sheet.cell(total_row, column).number_format = '$#,##0.00;[Red]($#,##0.00);-'
        for column in range(1, 6):
            cell = sheet.cell(total_row, column)
            cell.fill = PatternFill("solid", fgColor="006B2E")
            cell.font = Font(bold=True, color="FFFFFF")
        sheet.row_dimensions[total_row].height = 28

        sheet.column_dimensions["A"].width = 38
        for column in ("B", "C", "D", "E"):
            sheet.column_dimensions[column].width = 20
        sheet.freeze_panes = "A4"
        sheet.sheet_view.showGridLines = False

        table = Table(displayName=_table_name(f"{month:%B}BudgetSummary"), ref=f"A3:E{last_category_row}")
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium4",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        sheet.add_table(table)

    def _write_source_context(self, sheet, context: ReviewSourceContext) -> None:
        sheet["G1"] = "Import Source"
        sheet["H1"] = context.mode.display_name
        sheet["G2"] = "Coverage Status"
        sheet["H2"] = "Complete" if context.coverage.complete else "Incomplete"
        headers = ["Institution", "Status", "Covered Dates", "Posted Rows"]
        for offset, header in enumerate(headers, start=SOURCE_PANEL_START_COLUMN):
            cell = sheet.cell(3, offset)
            cell.value = header
            cell.fill = PatternFill("solid", fgColor="5B6573")
            cell.font = Font(bold=True, color="FFFFFF")
        for row_number, coverage in enumerate(context.coverage.rows, start=4):
            set_literal_cell(sheet.cell(row_number, 7), coverage.institution)
            set_literal_cell(sheet.cell(row_number, 8), coverage.status.display_name)
            set_literal_cell(sheet.cell(row_number, 9), _coverage_period_text(coverage))
            set_literal_cell(sheet.cell(row_number, 10), coverage.posted_transaction_count)
            if coverage.status is not PortalCoverageStatus.COMPLETE:
                for column in range(7, 11):
                    sheet.cell(row_number, column).fill = PatternFill("solid", fgColor="FFF2CC")
        sheet["G1"].font = Font(bold=True)
        sheet["G2"].font = Font(bold=True)
        sheet.column_dimensions["G"].width = 24
        sheet.column_dimensions["H"].width = 16
        sheet.column_dimensions["I"].width = 27
        sheet.column_dimensions["J"].width = 14

    def _write_reconciliation(
        self,
        sheet,
        month: date,
        reconciliation: ReconciliationReport,
    ) -> None:
        sheet.merge_cells("A1:I1")
        sheet["A1"] = f"{month:%B %Y} Institution Reconciliation"
        sheet["A1"].fill = PatternFill("solid", fgColor="1F4E78")
        sheet["A1"].font = Font(bold=True, color="FFFFFF", size=15)
        sheet["A1"].alignment = Alignment(vertical="center")
        sheet.row_dimensions[1].height = 28
        sheet.merge_cells("A2:I2")
        sheet["A2"] = (
            "Review queue only — these rows never enter Transactions or Budget Summary automatically. "
            "Plaid-marked removed items remain outside active spend pending human disposition."
        )
        sheet["A2"].font = Font(italic=True, color="666666", size=10)
        sheet["A2"].alignment = Alignment(wrap_text=True, vertical="center")
        sheet.row_dimensions[2].height = 34

        coverage_rows = reconciliation.normalized_coverage
        active_plaid_rows = sum(row.matched_count + row.plaid_only_count for row in coverage_rows)
        removed_visible_rows = sum(row.plaid_removed_portal_visible_count for row in coverage_rows)
        ambiguous_rows = sum(row.ambiguous_count for row in coverage_rows)
        for row_number, (label, value) in enumerate((
            ("Active Plaid rows (matched + Plaid-only posted)", active_plaid_rows),
            ("Plaid marked removed — still visible at institution", removed_visible_rows),
            ("Ambiguous date / identity matches", ambiguous_rows),
        ), start=4):
            set_literal_cell(sheet.cell(row_number, 1), label)
            set_literal_cell(sheet.cell(row_number, 2), value)
            sheet.cell(row_number, 1).font = Font(bold=True)

        coverage_headers = [
            "Institution",
            "Portal Rows",
            "Matched Active",
            "Plaid Removed / Portal Visible",
            "Ambiguous",
            "Portal Only",
            "Plaid-only Posted",
            "Plaid-only Pending",
            "Portal Coverage",
        ]
        for column, header in enumerate(coverage_headers, start=1):
            cell = sheet.cell(8, column)
            cell.value = header
            cell.fill = PatternFill("solid", fgColor="0F6B78")
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        for row_number, row in enumerate(coverage_rows, start=9):
            values = [
                row.institution,
                row.portal_posted_count,
                row.matched_count,
                row.plaid_removed_portal_visible_count,
                row.ambiguous_count,
                row.portal_only_count,
                row.plaid_only_count,
                row.plaid_only_pending_count,
                row.portal_status,
            ]
            for column, value in enumerate(values, start=1):
                set_literal_cell(sheet.cell(row_number, column), value)
        totals = [
            "Overall",
            sum(row.portal_posted_count for row in coverage_rows),
            sum(row.matched_count for row in coverage_rows),
            removed_visible_rows,
            ambiguous_rows,
            sum(row.portal_only_count for row in coverage_rows),
            sum(row.plaid_only_count for row in coverage_rows),
            sum(row.plaid_only_pending_count for row in coverage_rows),
            "Complete" if all(row.portal_status == "Complete" for row in coverage_rows) else "Incomplete",
        ]
        for column, value in enumerate(totals, start=1):
            set_literal_cell(sheet.cell(12, column), value)
            sheet.cell(12, column).fill = PatternFill("solid", fgColor="D9EAF7")
            sheet.cell(12, column).font = Font(bold=True)
        coverage_table = Table(displayName=RECONCILIATION_COVERAGE_TABLE, ref="A8:I12")
        coverage_table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        sheet.add_table(coverage_table)

        detail_headers = [
            "Institution",
            "Date",
            "Amount",
            "Classification",
            "Local Status",
            "Budget Effect",
            "Review Note",
        ]
        for column, header in enumerate(detail_headers, start=1):
            cell = sheet.cell(15, column)
            cell.value = header
            cell.fill = PatternFill("solid", fgColor="9C5700")
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        for row_number, item in enumerate(reconciliation.items, start=16):
            values = [
                item.institution,
                item.transaction_date,
                float(item.amount) if item.amount is not None else None,
                item.classification.value,
                item.local_status,
                item.budget_effect,
                item.review_note,
            ]
            for column, value in enumerate(values, start=1):
                set_literal_cell(sheet.cell(row_number, column), value)
                sheet.cell(row_number, column).alignment = Alignment(
                    wrap_text=column in (4, 5, 6, 7),
                    vertical="top",
                )
            sheet.cell(row_number, 2).number_format = "yyyy-mm-dd"
            sheet.cell(row_number, 3).number_format = '$#,##0.00;[Red]($#,##0.00);-'
            sheet.row_dimensions[row_number].height = _detail_row_height(item.review_note)
        last_detail_row = max(16, 15 + len(reconciliation.items))
        detail_table = Table(
            displayName=RECONCILIATION_DETAIL_TABLE,
            ref=f"A15:G{last_detail_row}",
        )
        detail_table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        sheet.add_table(detail_table)

        widths = {"A": 24, "B": 13, "C": 15, "D": 48, "E": 24, "F": 28, "G": 48, "H": 20, "I": 20}
        for column, width in widths.items():
            sheet.column_dimensions[column].width = width
        sheet.freeze_panes = "A16"
        sheet.sheet_view.showGridLines = False

    def _update_preserved_reconciliation(
        self,
        sheet,
        reconciliation: ReconciliationReport,
        existing_item_count: int,
    ) -> None:
        coverage_rows = reconciliation.normalized_coverage
        for row_number, row in enumerate(coverage_rows, start=9):
            values = [
                row.institution,
                row.portal_posted_count,
                row.matched_count,
                row.plaid_removed_portal_visible_count,
                row.ambiguous_count,
                row.portal_only_count,
                row.plaid_only_count,
                row.plaid_only_pending_count,
                row.portal_status,
            ]
            for column, value in enumerate(values, start=1):
                cell = sheet.cell(row_number, column)
                if cell.data_type != "f":
                    set_literal_cell(cell, value)

        totals = [
            "Overall",
            sum(row.portal_posted_count for row in coverage_rows),
            sum(row.matched_count for row in coverage_rows),
            sum(row.plaid_removed_portal_visible_count for row in coverage_rows),
            sum(row.ambiguous_count for row in coverage_rows),
            sum(row.portal_only_count for row in coverage_rows),
            sum(row.plaid_only_count for row in coverage_rows),
            sum(row.plaid_only_pending_count for row in coverage_rows),
            "Complete" if all(row.portal_status == "Complete" for row in coverage_rows) else "Incomplete",
        ]
        for column, value in enumerate(totals, start=1):
            cell = sheet.cell(12, column)
            if cell.data_type != "f":
                set_literal_cell(cell, value)

        if not any(sheet[coordinate].data_type == "f" for coordinate in RECONCILIATION_SAFE_FORMULAS):
            sheet["B4"] = sum(row.matched_count + row.plaid_only_count for row in coverage_rows)
            sheet["B5"] = sum(row.plaid_removed_portal_visible_count for row in coverage_rows)
            sheet["B6"] = sum(row.ambiguous_count for row in coverage_rows)

        template_row = 16 if sheet.max_row >= 16 else 15
        for item_index in range(existing_item_count, len(reconciliation.items)):
            row_number = item_index + 16
            if template_row >= 16 and row_number != template_row:
                for column in range(1, 8):
                    source = sheet.cell(template_row, column)
                    target = sheet.cell(row_number, column)
                    target._style = copy(source._style)
                    target.alignment = copy(source.alignment)
                    target.number_format = source.number_format
            item = reconciliation.items[item_index]
            values = [
                item.institution,
                item.transaction_date,
                float(item.amount) if item.amount is not None else None,
                item.classification.value,
                item.local_status,
                item.budget_effect,
                item.review_note,
            ]
            for column, value in enumerate(values, start=1):
                set_literal_cell(sheet.cell(row_number, column), value)
            sheet.cell(row_number, 2).number_format = "yyyy-mm-dd"
            sheet.cell(row_number, 3).number_format = '$#,##0.00;[Red]($#,##0.00);-'
            sheet.row_dimensions[row_number].height = _detail_row_height(item.review_note)

    def _ensure_reconciliation_detail_table(self, sheet, last_detail_row: int) -> None:
        expected_ref = f"A15:G{last_detail_row}"
        if RECONCILIATION_DETAIL_TABLE in sheet.tables:
            table = sheet.tables[RECONCILIATION_DETAIL_TABLE]
            if table.ref != expected_ref:
                table.ref = expected_ref
            return
        table = Table(displayName=RECONCILIATION_DETAIL_TABLE, ref=expected_ref)
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        sheet.add_table(table)

    def _write_reconciliation_state(
        self,
        state_sheet,
        reconciliation: ReconciliationReport,
    ) -> None:
        for row_number, item in enumerate(reconciliation.items, start=2):
            set_literal_cell(state_sheet.cell(row_number, 2), item.local_id)
            set_literal_cell(
                state_sheet.cell(row_number, 3),
                _reconciliation_bridge_hash(
                    item.institution,
                    item.transaction_date,
                    item.amount,
                ),
            )

    def _add_category_validation(self, workbook, sheet, category_count: int, last_transaction_row: int) -> None:
        last_category_row = category_count + 3
        workbook.defined_names.add(
            DefinedName(
                "BudgetCategoryList",
                attr_text=f"'{BUDGET_SUMMARY_SHEET}'!$A$4:$A${last_category_row}",
            )
        )
        validation = DataValidation(type="list", formula1="=BudgetCategoryList", allow_blank=True)
        validation.errorStyle = "stop"
        validation.errorTitle = "Invalid category"
        validation.error = "Choose a category from the Budget Summary list."
        validation.promptTitle = "Budget category"
        validation.prompt = "Only the selected Category drives Budget Summary Actual totals."
        validation.showErrorMessage = True
        validation.showInputMessage = True
        sheet.add_data_validation(validation)
        validation.add(f"E4:E{last_transaction_row}")

    def _select_detail(
        self,
        evidence: ReviewCellDetail | None,
        existing: ExistingReview | None,
    ) -> tuple[object, str | None, float | None]:
        if (
            existing is not None
            and existing.detail_value not in (None, "")
            and (
                existing.cross_source_bridge
                or existing.detail_value != REVIEW_NEEDED_DEFAULT
                or existing.hyperlink is not None
                or _formula_hyperlink(existing.detail_value) is not None
            )
        ):
            return existing.detail_value, existing.hyperlink, existing.row_height
        if evidence is not None:
            return evidence.text, evidence.hyperlink, None
        if existing is not None and existing.detail_value not in (None, ""):
            return existing.detail_value, existing.hyperlink, existing.row_height
        return REVIEW_NEEDED_DEFAULT, None, None

    def _read_existing(self, path: Path) -> dict[str, ExistingReview]:
        if not path.exists():
            return {}
        workbook = load_workbook(path, data_only=False)
        if TRANSACTIONS_SHEET not in workbook.sheetnames:
            return {}
        sheet = workbook[TRANSACTIONS_SHEET]
        row_by_key = _rows_by_transaction_identity(sheet, _header_column(sheet, REVIEW_KEY_HEADER))
        groups = [(sheet, row_by_key, False)]
        if RECONCILIATION_SHEET in workbook.sheetnames:
            queue = workbook[RECONCILIATION_SHEET]
            if PRESERVED_REVIEW_TABLE in queue.tables:
                bounds = range_boundaries(queue.tables[PRESERVED_REVIEW_TABLE].ref)
                _, first, _, last = bounds
                state = workbook[REVIEW_STATE_SHEET]
                archived_rows = {}
                for row in range(2, state.max_row + 1):
                    key, position = state.cell(row, 4).value, state.cell(row, 5).value
                    if key in (None, "") and position in (None, ""):
                        continue
                    if not _is_review_key(key) or not isinstance(position, int) or not first < position <= last:
                        raise ValueError("invalid preserved review identity state")
                    if key in archived_rows or position in archived_rows.values():
                        raise ValueError("duplicate preserved review identity state")
                    archived_rows[key] = position
                if set(archived_rows.values()) != set(range(first + 1, last + 1)):
                    raise ValueError("preserved review identity state is incomplete")
                groups.append((queue, archived_rows, True))
        existing = {}
        for sheet, row_by_key, archived in groups:
            for key, row in row_by_key.items():
                cell = sheet.cell(row, 6)
                formula_hyperlink = _formula_hyperlink(cell.value)
                native_hyperlink = _validated_gmail_url(cell.hyperlink.target if cell.hyperlink else None)
                treatment_cell = sheet.cell(row, 7)
                try:
                    budget_treatment = (
                        TreatmentDisposition(str(treatment_cell.value))
                        if treatment_cell.data_type != "f"
                        else None
                    )
                except ValueError:
                    budget_treatment = None
                if key in existing:
                    raise ValueError("duplicate active/preserved review identity")
                existing[key] = ExistingReview(
                    transaction=_transaction_from_review_row(sheet, row),
                    category=str(sheet.cell(row, 5).value or ""),
                    detail_value=cell.value,
                    hyperlink=None if formula_hyperlink else native_hyperlink,
                    row_height=sheet.row_dimensions[row].height,
                    budget_treatment=budget_treatment,
                    treatment_basis=str(sheet.cell(row, 8).value or ""),
                    reconciliation_disposition=(str(sheet.cell(row, 9).value or "Review needed — not active") if archived else "Review needed — not active"),
                    reconciliation_note=(str(sheet.cell(row, 10).value or "") if archived else "Source identity unresolved; original choice retained, not added to active spending."),
                )
        return existing

    def _write_preserved_reviews(self, workbook, preserved: dict[str, ExistingReview]) -> None:
        if RECONCILIATION_SHEET not in workbook.sheetnames:
            return
        sheet = workbook[RECONCILIATION_SHEET]
        first = range_boundaries(sheet.tables[RECONCILIATION_DETAIL_TABLE].ref)[3] + 5
        if not preserved:
            return
        sheet.cell(first - 2, 1, "Prior classified transactions — preserved for identity review")
        sheet.cell(first - 1, 1, "Not active spending. Confirm identity before carrying a choice to Transactions; do not append duplicates.")
        for row in (first - 2, first - 1):
            sheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=10)
            sheet.cell(row, 1).alignment = Alignment(wrap_text=True, vertical="top")
            sheet.row_dimensions[row].height = 34
        sheet.column_dimensions["J"].width = 48
        sheet.row_dimensions[first].height = 42
        headers = ["Date", "Merchant", "Credit Card Account", "Amount", "Category", "Purchased Item Detail", "Budget Treatment", "Treatment Basis", "Disposition", "Review Note"]
        for column, header in enumerate(headers, 1):
            cell = sheet.cell(first, column, header)
            cell.fill = PatternFill("solid", fgColor="9C5700")
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        state = workbook[REVIEW_STATE_SHEET]
        state.cell(1, 4, "Preserved Review Key")
        state.cell(1, 5, "Preserved Review Row")
        for index, (key, prior) in enumerate(preserved.items(), 1):
            row = first + index
            transaction = prior.transaction
            values = [transaction.transaction_date, transaction.merchant, transaction.account, float(transaction.amount), prior.category, prior.detail_value, prior.budget_treatment.value if prior.budget_treatment else "", prior.treatment_basis, prior.reconciliation_disposition, prior.reconciliation_note]
            for column, value in enumerate(values, 1):
                cell = sheet.cell(row, column)
                if column == 6 and _formula_hyperlink(value):
                    cell.value = value
                else:
                    set_literal_cell(cell, value)
                cell.alignment = Alignment(wrap_text=True, vertical="top")
            if prior.hyperlink:
                sheet.cell(row, 6).hyperlink = prior.hyperlink
            sheet.cell(row, 1).number_format = "yyyy-mm-dd"
            sheet.cell(row, 4).number_format = '$#,##0.00;[Red]($#,##0.00);-'
            sheet.row_dimensions[row].height = max(prior.row_height or 70, 110, min(409.5, ((len(str(prior.detail_value or "")) + 23) // 24) * 15 + 20))
            state.cell(index + 1, 4, key)
            state.cell(index + 1, 5, row)
        table = Table(displayName=PRESERVED_REVIEW_TABLE, ref=f"A{first}:J{first + len(preserved)}")
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
        sheet.add_table(table)
        workbook[BUDGET_SUMMARY_SHEET]["G8"] = "Preserved prior review rows"
        workbook[BUDGET_SUMMARY_SHEET]["H8"] = len(preserved)
        workbook[BUDGET_SUMMARY_SHEET]["G9"] = "Outside active totals"
        workbook[BUDGET_SUMMARY_SHEET]["H9"] = "Review Reconciliation"

    def _read_existing_source_context(self, path: Path) -> ReviewSourceContext | None:
        if not path.exists():
            return None
        workbook = load_workbook(path, data_only=False)
        if BUDGET_SUMMARY_SHEET not in workbook.sheetnames:
            return None
        sheet = workbook[BUDGET_SUMMARY_SHEET]
        if sheet["G1"].value != "Import Source":
            return None
        display_mode = str(sheet["H1"].value or "").strip().casefold()
        if display_mode == "plaid":
            mode = ImportSourceMode.PLAID
        elif display_mode == "direct portal":
            mode = ImportSourceMode.DIRECT_PORTAL
        else:
            raise ValueError("existing review source metadata cannot be preserved safely")
        month = _month_from_summary_title(sheet["A1"].value)
        rows: list[PortalInstitutionCoverage] = []
        try:
            for row_number in range(4, 7):
                period_start, period_end = _coverage_period(str(sheet.cell(row_number, 9).value or ""))
                rows.append(PortalInstitutionCoverage(
                    institution=str(sheet.cell(row_number, 7).value or ""),
                    status=PortalCoverageStatus.parse(str(sheet.cell(row_number, 8).value or "")),
                    period_start=period_start,
                    period_end=period_end,
                    posted_transaction_count=int(sheet.cell(row_number, 10).value or 0),
                ))
        except (TypeError, ValueError) as exc:
            raise ValueError("existing review source metadata cannot be preserved safely") from exc
        return ReviewSourceContext(
            mode=mode,
            coverage=SourceCoverage(mode=mode, month=month, rows=tuple(rows)),
        )

    def _read_existing_reconciliation(self, path: Path) -> ReconciliationReport | None:
        if not path.exists():
            return None
        workbook = load_workbook(path, data_only=False)
        if RECONCILIATION_SHEET not in workbook.sheetnames:
            return None
        sheet = workbook[RECONCILIATION_SHEET]
        if RECONCILIATION_COVERAGE_TABLE not in sheet.tables:
            raise ValueError("existing reconciliation queue cannot be preserved safely")
        _validate_reconciliation_formula_safety(sheet)

        coverage_table = sheet.tables[RECONCILIATION_COVERAGE_TABLE]
        coverage_min_col, coverage_min_row, coverage_max_col, coverage_max_row = range_boundaries(coverage_table.ref)
        if coverage_table.ref != "A8:I12":
            raise ValueError("existing reconciliation coverage table is invalid")
        if [sheet.cell(8, column).value for column in range(1, 10)] != RECONCILIATION_COVERAGE_HEADERS:
            raise ValueError("existing reconciliation coverage headers are invalid")
        detail_min_col, detail_min_row, detail_max_col, detail_max_row = _reconciliation_detail_bounds(sheet)
        coverage: list[ReconciliationCoverageRow] = []
        try:
            for row_number in range(coverage_min_row + 1, min(coverage_max_row, coverage_min_row + 3) + 1):
                institution = str(sheet.cell(row_number, coverage_min_col).value or "")
                if institution == "Overall":
                    continue
                coverage.append(ReconciliationCoverageRow(
                    institution=institution,
                    portal_status=str(sheet.cell(row_number, coverage_min_col + 8).value or "Not reconciled"),
                    portal_posted_count=_integer_cell(sheet.cell(row_number, coverage_min_col + 1).value),
                    active_plaid_count=(
                        _integer_cell(sheet.cell(row_number, coverage_min_col + 2).value)
                        + _integer_cell(sheet.cell(row_number, coverage_min_col + 6).value)
                    ),
                    matched_count=_integer_cell(sheet.cell(row_number, coverage_min_col + 2).value),
                    portal_only_count=_integer_cell(sheet.cell(row_number, coverage_min_col + 5).value),
                    plaid_only_count=_integer_cell(sheet.cell(row_number, coverage_min_col + 6).value),
                    ambiguous_count=_integer_cell(sheet.cell(row_number, coverage_min_col + 4).value),
                    plaid_removed_portal_visible_count=_integer_cell(
                        sheet.cell(row_number, coverage_min_col + 3).value
                    ),
                    plaid_only_pending_count=_integer_cell(sheet.cell(row_number, coverage_min_col + 7).value),
                ))

            items: list[ReconciliationItem] = []
            bridge_occurrences: dict[str, int] = {}
            for row_number in range(detail_min_row + 1, detail_max_row + 1):
                values = [sheet.cell(row_number, column).value for column in range(detail_min_col, detail_max_col + 1)]
                if all(value in (None, "") for value in values):
                    continue
                transaction_date = _coerce_date(values[1])
                amount = Decimal(str(values[2])) if values[2] not in (None, "") else None
                bridge_hash = _reconciliation_bridge_hash(str(values[0] or ""), transaction_date, amount)
                occurrence = bridge_occurrences.get(bridge_hash, 0) + 1
                bridge_occurrences[bridge_hash] = occurrence
                identity = _stored_reconciliation_identity(
                    workbook,
                    len(items) + 2,
                    bridge_hash,
                ) or _reconciliation_identity(values[:4] + [occurrence])
                items.append(ReconciliationItem(
                    local_id=identity,
                    institution=str(values[0] or ""),
                    transaction_date=transaction_date,
                    amount=amount,
                    classification=ReconciliationClassification.parse(str(values[3] or "")),
                    local_status=str(values[4] or "Review needed"),
                    budget_effect=str(values[5] or "No automatic budget change"),
                    review_note=str(values[6] or "Review needed"),
                ))
        except (ArithmeticError, TypeError, ValueError) as exc:
            raise ValueError("existing reconciliation queue cannot be preserved safely") from exc
        return ReconciliationReport(coverage=tuple(coverage), items=tuple(items))

    def _backup(self, path: Path) -> Path:
        backup_dir = path.parent / "backups"
        ensure_private_directory(backup_dir)
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
        backup = backup_dir / f"{path.stem}.{stamp}{path.suffix}"
        copy_private_file(path, backup)
        return backup

    def _validate_saved(
        self,
        path: Path,
        category_count: int,
        transaction_count: int,
        source_context: ReviewSourceContext | None,
        reconciliation: ReconciliationReport | None,
    ) -> None:
        workbook = load_workbook(path, data_only=False)
        if workbook.sheetnames[:2] != [TRANSACTIONS_SHEET, BUDGET_SUMMARY_SHEET]:
            raise ValueError("review workbook visible sheet order is invalid")
        if REVIEW_STATE_SHEET not in workbook.sheetnames or workbook[REVIEW_STATE_SHEET].sheet_state != "veryHidden":
            raise ValueError("review workbook state sheet is missing or visible")
        transactions = workbook[TRANSACTIONS_SHEET]
        summary = workbook[BUDGET_SUMMARY_SHEET]
        last_transaction_row = max(4, transaction_count + 3)
        last_category_row = category_count + 3
        validations = list(transactions.data_validations.dataValidation)
        if len(validations) != 1 or validations[0].type != "list" or validations[0].formula1 != "=BudgetCategoryList":
            raise ValueError("review workbook category validation is invalid")
        validation_ref = str(validations[0].sqref)
        expected_ref = f"E4:E{last_transaction_row}"
        if expected_ref not in validation_ref and not (last_transaction_row == 4 and validation_ref == "E4"):
            raise ValueError("review workbook category validation range is invalid")
        if workbook.defined_names["BudgetCategoryList"].attr_text != f"'{BUDGET_SUMMARY_SHEET}'!$A$4:$A${last_category_row}":
            raise ValueError("review workbook category authority name is invalid")
        expected_headers = [
            "Date",
            "Merchant",
            "Credit Card Account",
            "Amount",
            "Category",
            "Purchased Item Detail",
            "Budget Treatment",
            "Treatment Basis",
            REVIEW_KEY_HEADER,
        ]
        if [transactions.cell(3, column).value for column in range(1, REVIEW_KEY_COLUMN + 1)] != expected_headers:
            raise ValueError("review workbook transaction headers are invalid")
        if not transactions.column_dimensions["I"].hidden:
            raise ValueError("review workbook identity column must be hidden")
        expected_table_ref = f"A3:I{last_transaction_row}"
        if not transactions.tables or any(table.ref != expected_table_ref for table in transactions.tables.values()):
            raise ValueError("review workbook transaction table range is invalid")
        allowed_treatments = {disposition.value for disposition in TreatmentDisposition}
        transaction_keys: list[str] = []
        for row in range(4, last_transaction_row + 1):
            if all(transactions.cell(row, column).value in (None, "") for column in range(1, 5)):
                continue
            if transactions.cell(row, 7).value not in allowed_treatments:
                raise ValueError("review workbook contains an invalid Budget Treatment")
            review_key = transactions.cell(row, REVIEW_KEY_COLUMN).value
            if not _is_review_key(review_key) or review_key in transaction_keys:
                raise ValueError("review workbook contains an invalid or duplicate Review Key")
            transaction_keys.append(review_key)
            for column in range(1, REVIEW_KEY_COLUMN + 1):
                cell = transactions.cell(row, column)
                if cell.data_type == "f" and not (column == 6 and _formula_hyperlink(cell.value)):
                    raise ValueError("review workbook contains an untrusted transaction formula")
        state = workbook[REVIEW_STATE_SHEET]
        state_keys = [
            state.cell(row, 1).value
            for row in range(2, state.max_row + 1)
            if state.cell(row, 1).value not in (None, "")
        ]
        if state_keys != transaction_keys:
            raise ValueError("review workbook state keys do not reconcile to transaction identities")
        for row in range(4, last_category_row + 1):
            expected = (
                f'=SUMIFS(\'Transactions\'!$D$4:$D${last_transaction_row},'
                f"'Transactions'!$E$4:$E${last_transaction_row},A{row},"
                f"'Transactions'!$G$4:$G${last_transaction_row},"
                f'"<>{TreatmentDisposition.EXCLUDED.value}")'
            )
            if summary.cell(row, 3).value != expected:
                raise ValueError("review workbook Actual formula is invalid")
        formulas = [
            cell.value
            for sheet in (transactions, summary)
            for row in sheet.iter_rows()
            for cell in row
            if cell.data_type == "f"
        ]
        if any("#REF!" in formula.upper() for formula in formulas):
            raise ValueError("review workbook contains a broken formula reference")
        self._validate_source_context(summary, source_context)
        self._validate_reconciliation(workbook, reconciliation)

    def _validate_source_context(self, summary, context: ReviewSourceContext | None) -> None:
        if context is None:
            return
        if summary["G1"].value != "Import Source" or summary["H1"].value != context.mode.display_name:
            raise ValueError("review workbook import source metadata is invalid")
        expected_status = "Complete" if context.coverage.complete else "Incomplete"
        if summary["G2"].value != "Coverage Status" or summary["H2"].value != expected_status:
            raise ValueError("review workbook source coverage status is invalid")
        for row_number, row in enumerate(context.coverage.rows, start=4):
            actual = [summary.cell(row_number, column).value for column in range(7, 11)]
            expected = [
                row.institution,
                row.status.display_name,
                _coverage_period_text(row),
                row.posted_transaction_count,
            ]
            if actual != expected:
                raise ValueError("review workbook source coverage rows are invalid")

    def _validate_reconciliation(self, workbook, reconciliation: ReconciliationReport | None) -> None:
        if reconciliation is None:
            if RECONCILIATION_SHEET in workbook.sheetnames:
                raise ValueError("unexpected reconciliation sheet")
            return
        if RECONCILIATION_SHEET not in workbook.sheetnames:
            raise ValueError("review workbook reconciliation sheet is missing")
        sheet = workbook[RECONCILIATION_SHEET]
        if sheet.tables[RECONCILIATION_COVERAGE_TABLE].ref != "A8:I12":
            raise ValueError("review workbook reconciliation coverage table is invalid")
        expected_detail_ref = f"A15:G{max(16, 15 + len(reconciliation.items))}"
        if sheet.tables[RECONCILIATION_DETAIL_TABLE].ref != expected_detail_ref:
            raise ValueError("review workbook reconciliation detail table is invalid")
        if [sheet.cell(8, column).value for column in range(1, 10)] != RECONCILIATION_COVERAGE_HEADERS:
            raise ValueError("review workbook reconciliation coverage headers are invalid")
        if [sheet.cell(15, column).value for column in range(1, 8)] != RECONCILIATION_DETAIL_HEADERS:
            raise ValueError("review workbook reconciliation detail headers are invalid")
        for row_number, item in enumerate(reconciliation.items, start=16):
            if sheet.cell(row_number, 4).value != item.classification.value:
                raise ValueError("review workbook reconciliation classification is invalid")
        _validate_reconciliation_formula_safety(sheet)


def _header_column(sheet, header: str) -> int | None:
    for column in range(1, sheet.max_column + 1):
        if sheet.cell(3, column).value == header:
            return column
    return None


def _rows_by_transaction_identity(sheet, key_column: int | None) -> dict[str, int]:
    visible_rows: list[tuple[int, LedgerTransaction]] = []
    for row in range(4, sheet.max_row + 1):
        if all(sheet.cell(row, column).value in (None, "") for column in range(1, 5)):
            continue
        tx_date = _coerce_date(sheet.cell(row, 1).value)
        amount = sheet.cell(row, 4).value
        if tx_date is None or not isinstance(amount, (int, float, Decimal)):
            continue
        visible_rows.append((row, _transaction_from_review_row(sheet, row)))

    calculated_keys = make_review_keys([transaction for _, transaction in visible_rows])
    allowed_by_identity: dict[tuple[str, str, str, str], list[str]] = {}
    for (_, transaction), key in zip(visible_rows, calculated_keys, strict=True):
        allowed_by_identity.setdefault(transaction_identity(transaction), []).append(key)

    rows_by_identity: dict[tuple[str, str, str, str], list[int]] = {}
    for row, transaction in visible_rows:
        rows_by_identity.setdefault(transaction_identity(transaction), []).append(row)

    rows_by_key: dict[str, int] = {}
    for identity, rows in rows_by_identity.items():
        unused = list(allowed_by_identity[identity])
        for row in rows:
            embedded = sheet.cell(row, key_column).value if key_column else None
            if _is_review_key(embedded) and embedded in unused:
                key = embedded
            else:
                key = unused[0]
            unused.remove(key)
            rows_by_key[key] = row
    return rows_by_key


def _transaction_from_review_row(sheet, row: int) -> LedgerTransaction:
    transaction_date = _coerce_date(sheet.cell(row, 1).value)
    amount = sheet.cell(row, 4).value
    if transaction_date is None or not isinstance(amount, (int, float, Decimal)):
        raise ValueError("existing review transaction identity cannot be preserved safely")
    return LedgerTransaction(
        transaction_date=transaction_date,
        month=date(transaction_date.year, transaction_date.month, 1),
        merchant=str(sheet.cell(row, 2).value or ""),
        account=str(sheet.cell(row, 3).value or ""),
        amount=Decimal(str(amount)),
        personal_budget=True,
        source="existing-review",
    )


def _is_review_key(value: object) -> bool:
    return isinstance(value, str) and re.fullmatch(r"[0-9a-f]{64}", value) is not None


def _coerce_date(value: object) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        try:
            return date.fromisoformat(value)
        except ValueError:
            return None
    return None


def _table_name(value: str) -> str:
    name = re.sub(r"[^A-Za-z0-9_]", "", value)
    if not name or name[0].isdigit():
        name = "Review" + name
    return name[:255]


def _detail_row_height(detail: str) -> float:
    lines = max(1, detail.count("\n") + 1)
    return min(90.0, max(18.0, 15.0 * lines))


def _coverage_period_text(coverage: PortalInstitutionCoverage) -> str:
    if coverage.period_start is None or coverage.period_end is None:
        return "Not available"
    return f"{coverage.period_start.isoformat()} to {coverage.period_end.isoformat()}"


def _coverage_period(value: str) -> tuple[date | None, date | None]:
    if not value or value == "Not available":
        return None, None
    match = re.fullmatch(r"(\d{4}-\d{2}-\d{2}) to (\d{4}-\d{2}-\d{2})", value)
    if not match:
        raise ValueError("invalid source coverage period")
    return date.fromisoformat(match.group(1)), date.fromisoformat(match.group(2))


def _month_from_summary_title(value: object) -> date:
    match = re.fullmatch(r"([A-Za-z]+) (\d{4}) Budget Summary", str(value or ""))
    if not match:
        raise ValueError("invalid review summary title")
    return datetime.strptime(f"{match.group(1)} {match.group(2)}", "%B %Y").date().replace(day=1)


def _integer_cell(value: object) -> int:
    if isinstance(value, bool):
        raise ValueError("invalid reconciliation count")
    parsed = int(value or 0)
    if parsed < 0:
        raise ValueError("invalid reconciliation count")
    return parsed


def _reconciliation_detail_bounds(sheet) -> tuple[int, int, int, int]:
    if [sheet.cell(15, column).value for column in range(1, 8)] != RECONCILIATION_DETAIL_HEADERS:
        raise ValueError("existing reconciliation detail headers are invalid")
    if RECONCILIATION_DETAIL_TABLE in sheet.tables:
        bounds = range_boundaries(sheet.tables[RECONCILIATION_DETAIL_TABLE].ref)
        if bounds[:3] != (1, 15, 7) or bounds[3] < 16:
            raise ValueError("existing reconciliation detail table is invalid")
        return bounds
    populated_rows = [
        row_number
        for row_number in range(16, sheet.max_row + 1)
        if any(sheet.cell(row_number, column).value not in (None, "") for column in range(1, 8))
    ]
    return 1, 15, 7, max(populated_rows, default=16)


def _validate_reconciliation_formula_safety(sheet) -> None:
    for row in sheet.iter_rows():
        for cell in row:
            if cell.data_type != "f":
                continue
            if PRESERVED_REVIEW_TABLE in sheet.tables:
                _, first, _, last = range_boundaries(sheet.tables[PRESERVED_REVIEW_TABLE].ref)
                if cell.column == 6 and first < cell.row <= last and _formula_hyperlink(cell.value):
                    continue
            expected = RECONCILIATION_SAFE_FORMULAS.get(cell.coordinate)
            if expected is None or _normalized_formula(cell.value) != _normalized_formula(expected):
                raise ValueError("existing reconciliation queue contains an unexpected formula")


def _normalized_formula(value: object) -> str:
    return re.sub(r"\s+", "", str(value or "")).upper()


def _stored_reconciliation_identity(
    workbook,
    state_row: int,
    bridge_hash: str,
) -> str | None:
    if REVIEW_STATE_SHEET not in workbook.sheetnames:
        return None
    state = workbook[REVIEW_STATE_SHEET]
    if (
        state.cell(1, 2).value != RECONCILIATION_KEY_HEADER
        or state.cell(1, 3).value != RECONCILIATION_BRIDGE_HEADER
    ):
        return None
    stored_key = state.cell(state_row, 2).value
    stored_bridge = state.cell(state_row, 3).value
    if not _is_review_key(stored_key) or stored_bridge != bridge_hash:
        raise ValueError("existing reconciliation identity state cannot be preserved safely")
    return stored_key


def _reconciliation_bridge_hash(
    institution: str,
    transaction_date: date | None,
    amount: Decimal | None,
) -> str:
    normalized_amount = format(amount.normalize(), "f") if amount is not None else ""
    normalized_date = transaction_date.isoformat() if transaction_date is not None else ""
    payload = "\x1f".join((institution.strip(), normalized_date, normalized_amount))
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _reconciliation_identity(values: list[object]) -> str:
    normalized: list[str] = []
    for value in values:
        if isinstance(value, datetime):
            normalized.append(value.date().isoformat())
        elif isinstance(value, date):
            normalized.append(value.isoformat())
        elif isinstance(value, Decimal):
            normalized.append(format(value.normalize(), "f"))
        else:
            normalized.append(str(value or "").strip())
    return hashlib.sha256("\x1f".join(normalized).encode("utf-8")).hexdigest()


def _is_review_needed(value: object) -> bool:
    return isinstance(value, str) and value.startswith("Review needed")


def _formula_hyperlink(value: object) -> str | None:
    if not isinstance(value, str):
        return None
    match = re.fullmatch(
        r'=HYPERLINK\("(https://mail\.google\.com/[^"\r\n]+)"\s*,\s*'
        r'"(?:[^"]|"")*"'
        r'(?:\s*&\s*CHAR\(10\)\s*&\s*"(?:[^"]|"")*")*\s*\)',
        value,
        re.IGNORECASE,
    )
    return _validated_gmail_url(match.group(1)) if match else None


def _validated_gmail_url(value: object) -> str | None:
    if not isinstance(value, str) or "\r" in value or "\n" in value:
        return None
    parsed = urlsplit(value)
    try:
        port = parsed.port
    except ValueError:
        return None
    if (
        parsed.scheme != "https"
        or parsed.hostname != "mail.google.com"
        or port not in (None, 443)
        or parsed.username
        or parsed.password
    ):
        return None
    return value


def _is_gmail_link(detail: object, hyperlink: str | None) -> bool:
    return bool(_validated_gmail_url(hyperlink) or _formula_hyperlink(detail))
