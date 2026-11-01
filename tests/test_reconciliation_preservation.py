from __future__ import annotations

import hashlib
import tempfile
import unittest
from datetime import date
from decimal import Decimal
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from personal_budget_plaid_excel.domain.models import CategoryBudget
from personal_budget_plaid_excel.domain.source_mode import (
    DIRECT_PORTAL_INSTITUTIONS,
    ReconciliationClassification,
    ReconciliationCoverageRow,
    ReconciliationItem,
    ReconciliationReport,
)
from personal_budget_plaid_excel.infrastructure.review_workbook import (
    RECONCILIATION_COVERAGE_TABLE,
    RECONCILIATION_DETAIL_TABLE,
    RECONCILIATION_SHEET,
    OpenpyxlReviewWorkbookWriter,
)


SAFE_RECONCILIATION_FORMULAS = {
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


class ReconciliationPreservationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.writer = OpenpyxlReviewWorkbookWriter()
        self.month = date(2026, 8, 1)
        self.budgets = [CategoryBudget("Fixed", "Synthetic Fixed", Decimal("1000"))]

    def test_recognized_formula_sheet_roundtrips_without_layout_or_annotation_loss(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "review.xlsx"
            report = _report(tuple(_item(index) for index in range(38)))
            self.writer.write(output, self.month, self.budgets, [], reconciliation=report)
            _install_recognized_formula_layout(output)
            before = _reconciliation_fingerprint(output)

            result = self.writer.write(output, self.month, self.budgets, [], reconciliation=None)

            self.assertEqual(result.reconciliation_removed_visible_count, 19)
            self.assertEqual(result.reconciliation_ambiguous_count, 19)
            self.assertEqual(before, _reconciliation_fingerprint(output))
            preserved = load_workbook(output, data_only=False)[RECONCILIATION_SHEET]
            self.assertEqual(preserved.tables[RECONCILIATION_COVERAGE_TABLE].ref, "A8:I12")
            self.assertEqual(preserved.tables[RECONCILIATION_DETAIL_TABLE].ref, "A15:G53")
            self.assertEqual(
                {cell.coordinate: cell.value for row in preserved for cell in row if cell.data_type == "f"},
                SAFE_RECONCILIATION_FORMULAS,
            )

    def test_missing_detail_table_is_restored_from_recognizable_static_queue(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "review.xlsx"
            report = _report(tuple(_item(index) for index in range(38)))
            self.writer.write(output, self.month, self.budgets, [], reconciliation=report)
            _install_recognized_formula_layout(output, remove_detail_table=True)

            result = self.writer.write(output, self.month, self.budgets, [], reconciliation=None)

            self.assertEqual(result.reconciliation_removed_visible_count, 19)
            restored = load_workbook(output, data_only=False)[RECONCILIATION_SHEET]
            self.assertEqual(restored.tables[RECONCILIATION_DETAIL_TABLE].ref, "A15:G53")
            self.assertEqual(restored["G53"].value, "Synthetic user note 37")
            self.assertEqual(
                {cell.coordinate: cell.value for row in restored for cell in row if cell.data_type == "f"},
                SAFE_RECONCILIATION_FORMULAS,
            )

    def test_new_report_merges_prior_queue_and_preserves_existing_user_note(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "review.xlsx"
            retained = _item(0)
            omitted_by_new_report = _item(1)
            self.writer.write(
                output,
                self.month,
                self.budgets,
                [],
                reconciliation=_report((retained, omitted_by_new_report)),
            )
            _install_recognized_formula_layout(output)
            workbook = load_workbook(output)
            workbook[RECONCILIATION_SHEET]["G16"] = "User-reviewed note that must remain authoritative"
            workbook.save(output)
            incoming_match = ReconciliationItem(
                local_id=retained.local_id,
                institution=retained.institution,
                transaction_date=retained.transaction_date,
                amount=retained.amount,
                classification=retained.classification,
                local_status=retained.local_status,
                budget_effect=retained.budget_effect,
                review_note="Fresh automated note must not overwrite the user note",
            )
            appended = _item(2)

            self.writer.write(
                output,
                self.month,
                self.budgets,
                [],
                reconciliation=_report((incoming_match, appended)),
            )

            sheet = load_workbook(output, data_only=False)[RECONCILIATION_SHEET]
            self.assertEqual(sheet.tables[RECONCILIATION_DETAIL_TABLE].ref, "A15:G18")
            self.assertEqual(sheet["G16"].value, "User-reviewed note that must remain authoritative")
            self.assertEqual(sheet["G17"].value, omitted_by_new_report.review_note)
            self.assertEqual(sheet["G18"].value, appended.review_note)
            self.assertEqual(
                {cell.coordinate: cell.value for row in sheet for cell in row if cell.data_type == "f"},
                SAFE_RECONCILIATION_FORMULAS,
            )

    def test_unique_legacy_identity_bridge_merges_but_ambiguous_bridge_fails_closed(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            unique_output = root / "unique.xlsx"
            existing = _item(0)
            self.writer.write(unique_output, self.month, self.budgets, [], reconciliation=_report((existing,)))
            _remove_reconciliation_identity_state(unique_output)
            bridged = ReconciliationItem(
                local_id=_local_id("replacement-source-id"),
                institution=existing.institution,
                transaction_date=existing.transaction_date,
                amount=existing.amount,
                classification=existing.classification,
                local_status=existing.local_status,
                budget_effect=existing.budget_effect,
                review_note="New source note",
            )

            self.writer.write(
                unique_output,
                self.month,
                self.budgets,
                [],
                reconciliation=_report((bridged,)),
            )

            unique_sheet = load_workbook(unique_output)[RECONCILIATION_SHEET]
            self.assertEqual(unique_sheet.tables[RECONCILIATION_DETAIL_TABLE].ref, "A15:G16")
            self.assertEqual(unique_sheet["G16"].value, existing.review_note)

            ambiguous_output = root / "ambiguous.xlsx"
            first = _item(10)
            second = ReconciliationItem(
                local_id=_local_id("duplicate-bridge"),
                institution=first.institution,
                transaction_date=first.transaction_date,
                amount=first.amount,
                classification=first.classification,
                local_status=first.local_status,
                budget_effect=first.budget_effect,
                review_note="Second prior queue row with the same economic bridge",
            )
            self.writer.write(
                ambiguous_output,
                self.month,
                self.budgets,
                [],
                reconciliation=_report((first, second)),
            )
            _remove_reconciliation_identity_state(ambiguous_output)
            before_hash = _sha256(ambiguous_output)
            incoming = ReconciliationItem(
                local_id=_local_id("ambiguous-incoming"),
                institution=first.institution,
                transaction_date=first.transaction_date,
                amount=first.amount,
                classification=first.classification,
                local_status=first.local_status,
                budget_effect=first.budget_effect,
                review_note="Cannot choose one prior row safely",
            )

            with self.assertRaisesRegex(ValueError, "ambiguous reconciliation identity"):
                self.writer.write(
                    ambiguous_output,
                    self.month,
                    self.budgets,
                    [],
                    reconciliation=_report((incoming,)),
                )
            self.assertEqual(_sha256(ambiguous_output), before_hash)

    def test_conflicting_disposition_and_unrecognized_formula_fail_before_replacement(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            conflict_output = root / "conflict.xlsx"
            existing = _item(0)
            self.writer.write(conflict_output, self.month, self.budgets, [], reconciliation=_report((existing,)))
            before_hash = _sha256(conflict_output)
            conflict = ReconciliationItem(
                local_id=existing.local_id,
                institution=existing.institution,
                transaction_date=existing.transaction_date,
                amount=existing.amount,
                classification=existing.classification,
                local_status="Conflicting resolved disposition",
                budget_effect=existing.budget_effect,
                review_note=existing.review_note,
            )

            with self.assertRaisesRegex(ValueError, "conflicting reconciliation disposition"):
                self.writer.write(
                    conflict_output,
                    self.month,
                    self.budgets,
                    [],
                    reconciliation=_report((conflict,)),
                )
            self.assertEqual(_sha256(conflict_output), before_hash)

            formula_output = root / "unsafe-formula.xlsx"
            self.writer.write(formula_output, self.month, self.budgets, [], reconciliation=_report((existing,)))
            workbook = load_workbook(formula_output)
            workbook[RECONCILIATION_SHEET]["B5"] = '=WEBSERVICE("https://invalid.example")'
            workbook.save(formula_output)
            formula_hash = _sha256(formula_output)

            with self.assertRaisesRegex(ValueError, "unexpected formula"):
                self.writer.write(formula_output, self.month, self.budgets, [], reconciliation=None)
            self.assertEqual(_sha256(formula_output), formula_hash)


def _report(items: tuple[ReconciliationItem, ...]) -> ReconciliationReport:
    coverage = tuple(
        ReconciliationCoverageRow(
            institution=institution,
            portal_status="Complete",
            portal_posted_count=sum(item.institution == institution for item in items),
            active_plaid_count=0,
            matched_count=0,
            portal_only_count=0,
            plaid_only_count=0,
            ambiguous_count=sum(
                item.institution == institution
                and item.classification is ReconciliationClassification.AMBIGUOUS
                for item in items
            ),
            plaid_removed_portal_visible_count=sum(
                item.institution == institution
                and item.classification is ReconciliationClassification.PLAID_REMOVED_PORTAL_VISIBLE
                for item in items
            ),
        )
        for institution in DIRECT_PORTAL_INSTITUTIONS
    )
    return ReconciliationReport(coverage=coverage, items=items)


def _item(index: int) -> ReconciliationItem:
    classification = (
        ReconciliationClassification.PLAID_REMOVED_PORTAL_VISIBLE
        if index % 2 == 0
        else ReconciliationClassification.AMBIGUOUS
    )
    return ReconciliationItem(
        local_id=_local_id(f"reconciliation-{index}"),
        institution=DIRECT_PORTAL_INSTITUTIONS[index % len(DIRECT_PORTAL_INSTITUTIONS)],
        transaction_date=date(2026, 8, (index % 28) + 1),
        amount=Decimal(index + 1).quantize(Decimal("0.00")),
        classification=classification,
        local_status="Review needed",
        budget_effect="No automatic budget change",
        review_note=f"Synthetic user note {index}",
    )


def _install_recognized_formula_layout(path: Path, *, remove_detail_table: bool = False) -> None:
    workbook = load_workbook(path)
    sheet = workbook[RECONCILIATION_SHEET]
    for coordinate, formula in SAFE_RECONCILIATION_FORMULAS.items():
        sheet[coordinate] = formula
    sheet["E5"].fill = PatternFill("solid", fgColor="DDEBF7")
    sheet["E5"].font = Font(bold=True, color="1F4E78")
    sheet["G20"].alignment = Alignment(wrap_text=True, vertical="center", indent=1)
    sheet.row_dimensions[20].height = 61.5
    sheet.column_dimensions["G"].width = 57.25
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.fitToWidth = 1
    sheet.sheet_view.zoomScale = 85
    if remove_detail_table:
        del sheet.tables[RECONCILIATION_DETAIL_TABLE]
    workbook.save(path)


def _remove_reconciliation_identity_state(path: Path) -> None:
    workbook = load_workbook(path)
    state = workbook["_ReviewState"]
    for row in range(1, state.max_row + 1):
        for column in range(2, state.max_column + 1):
            state.cell(row, column).value = None
    workbook.save(path)


def _reconciliation_fingerprint(path: Path) -> tuple[object, ...]:
    sheet = load_workbook(path, data_only=False)[RECONCILIATION_SHEET]
    cells = tuple(
        (
            cell.coordinate,
            cell.value,
            cell.data_type,
            tuple(cell._style),
            cell.hyperlink.target if cell.hyperlink else None,
            cell.comment.text if cell.comment else None,
        )
        for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column)
        for cell in row
        if cell.value not in (None, "") or cell.has_style or cell.hyperlink or cell.comment
    )
    rows = tuple(
        (index, dimension.height, dimension.hidden, dimension.outlineLevel)
        for index, dimension in sorted(sheet.row_dimensions.items())
    )
    columns = tuple(
        (index, dimension.width, dimension.hidden, dimension.outlineLevel)
        for index, dimension in sorted(sheet.column_dimensions.items())
    )
    tables = tuple(
        (table.name, table.ref, table.tableStyleInfo.name if table.tableStyleInfo else None)
        for table in sorted(sheet.tables.values(), key=lambda value: value.name)
    )
    return (
        cells,
        rows,
        columns,
        tuple(sorted(str(cell_range) for cell_range in sheet.merged_cells.ranges)),
        tables,
        str(sheet.freeze_panes),
        sheet.sheet_view.showGridLines,
        sheet.sheet_view.zoomScale,
        sheet.page_setup.fitToWidth,
        sheet.sheet_properties.pageSetUpPr.fitToPage,
    )


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _local_id(seed: str) -> str:
    return hashlib.sha256(seed.encode("utf-8")).hexdigest()


if __name__ == "__main__":
    unittest.main()
