from __future__ import annotations

import re
from dataclasses import replace
from datetime import date, datetime, timezone
from decimal import Decimal
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

from personal_budget_plaid_excel.domain.category import (
    build_category_map,
    build_merchant_history,
    canonical_budget_category,
    infer_suggestion,
    is_budget_excluded,
)
from personal_budget_plaid_excel.domain.models import (
    CategoryBudget,
    LedgerTransaction,
    Money,
    PlaidSyncResult,
    SyncItemState,
)
from personal_budget_plaid_excel.domain.non_spend import (
    TreatmentDisposition,
    classify_budget_treatments,
    classify_provider_payload,
)
from personal_budget_plaid_excel.infrastructure.excel_cells import set_literal_cell
from personal_budget_plaid_excel.infrastructure.private_files import (
    copy_private_file,
    ensure_private_directory,
    private_temporary_path,
    replace_private_file,
)

ACTUAL_CHARGES_SHEET = "Actual Charges"
MANUAL_BUDGET_SHEET = "Manual Budget"
CATEGORY_MAP_SHEET = "Category Map"
MONTHLY_BUDGET_SHEET = "Monthly Budget"
SYNC_STATE_SHEET = "_SyncState"
ACTUAL_TABLE_NAME = "ActualChargesTable"
SYNC_MIGRATION_VERSION = "1"
# Keep category entry convenient for new manual rows while formula ranges use the
# actual managed ledger extent and are rebuilt after each import.
MAX_FORMULA_ROWS = 20000
EXCLUDED_TREATMENT_VALUE = TreatmentDisposition.EXCLUDED.value

BASE_HEADERS = [
    "Date",
    "Month",
    "Merchant",
    "Account",
    "Saved Category",
    "Cost Type",
    "Amount",
    "Personal Budget?",
    "Cancel?",
    "Refund?",
    "Won't Recur?",
    "Evidence",
]

EXTENSION_HEADERS = [
    "Selected Category",
    "Suggested Category",
    "Suggestion Reason",
    "Plaid Transaction ID",
    "Plaid Pending Transaction ID",
    "Plaid Account ID",
    "Plaid Item ID",
    "Plaid Status",
    "Plaid Removed?",
    "Sync Source",
    "Last Synced At",
    "Plaid PFC Primary",
    "Plaid PFC Detailed",
    "Plaid PFC Confidence",
    "Budget Treatment",
    "Treatment Basis",
    "Plaid Original Description",
    "Plaid Transaction Code",
    "Plaid Account Type",
    "Plaid Account Subtype",
]

SYNC_HEADERS = [
    "Item ID",
    "Institution Name",
    "Cursor",
    "Access Token Key",
    "Last Successful Sync",
    "Last Request ID",
    "Updated At",
    "Error",
]


class BudgetWorkbook:
    def __init__(self, path: Path) -> None:
        self.path = path
        self.workbook = load_workbook(path)

    def ensure_structure(self, as_of: date | None = None) -> bool:
        changed = False
        budgets = self.read_budgets()
        changed |= self._ensure_actual_charges_columns(budgets)
        changed |= self._ensure_category_map(budgets)
        changed |= self._ensure_sync_state()
        changed |= self._ensure_monthly_budget(budgets, as_of or date.today())
        return changed

    def save_with_backup(self, *, backup: bool = True) -> Path | None:
        backup_path: Path | None = None
        if backup:
            backup_dir = self.path.parent / "backups"
            ensure_private_directory(backup_dir)
            stamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
            backup_path = backup_dir / f"{self.path.stem}.{stamp}{self.path.suffix}"
            copy_private_file(self.path, backup_path)

        with private_temporary_path(self.path) as tmp_path:
            self.workbook.save(tmp_path)
            replace_private_file(tmp_path, self.path)
        return backup_path

    def read_budgets(self) -> list[CategoryBudget]:
        return [budget for _, budget in self._read_budgets_with_source_rows()]

    def _read_budgets_with_source_rows(self) -> list[tuple[int, CategoryBudget]]:
        ws = self.workbook[MANUAL_BUDGET_SHEET]
        budgets: list[tuple[int, CategoryBudget]] = []
        for row in range(2, ws.max_row + 1):
            cost_type = ws.cell(row, 1).value
            category = ws.cell(row, 2).value
            amount = ws.cell(row, 3).value
            if cost_type not in {"Fixed", "Flexible"}:
                if cost_type == "Budget":
                    break
                continue
            if not category:
                continue
            budgets.append(
                (
                    row,
                    CategoryBudget(
                        cost_type=str(cost_type),
                        category=str(category),
                        monthly_budget=self._manual_budget_amount(row),
                    ),
                )
            )
        return budgets

    def read_transactions(self) -> list[LedgerTransaction]:
        transactions = self._read_transaction_rows()
        treatments = classify_budget_treatments(transactions)
        return [
            replace(
                transaction,
                personal_budget=transaction.personal_budget and not treatment.excluded_from_budget,
                suggested_category=(
                    None if treatment.suppress_category_suggestion else transaction.suggested_category
                ),
            )
            for transaction, treatment in zip(transactions, treatments, strict=True)
        ]

    def _read_transaction_rows(self) -> list[LedgerTransaction]:
        ws = self.workbook[ACTUAL_CHARGES_SHEET]
        headers = self._headers(ws)
        transactions: list[LedgerTransaction] = []
        for row in range(6, ws.max_row + 1):
            if not ws.cell(row, 1).value and not ws.cell(row, 3).value:
                continue
            raw = {header: ws.cell(row, col).value for header, col in headers.items()}
            tx_date = _coerce_date(raw.get("Date"))
            month = _coerce_date(raw.get("Month")) or _month_start(tx_date)
            saved = _optional_str(raw.get("Saved Category"))
            selected = _optional_str(raw.get("Selected Category"))
            suggested = _optional_str(raw.get("Suggested Category"))
            selected_for_budget = selected if "Selected Category" in headers else saved
            amount = _money(raw.get("Amount"))
            merchant = _optional_str(raw.get("Merchant")) or ""
            category_for_personal = selected_for_budget or saved
            personal_value = raw.get("Personal Budget?")
            personal_budget = str(personal_value).strip().lower() == "yes"
            if is_budget_excluded(category_for_personal):
                personal_budget = False
            status = (_optional_str(raw.get("Plaid Status")) or "posted").lower()
            transactions.append(
                LedgerTransaction(
                    transaction_date=tx_date,
                    month=month,
                    merchant=merchant,
                    account=_optional_str(raw.get("Account")) or "",
                    amount=amount,
                    personal_budget=personal_budget,
                    saved_category=saved,
                    selected_category=canonical_budget_category(selected_for_budget),
                    suggested_category=canonical_budget_category(suggested),
                    is_pending=status == "pending",
                    is_removed=str(raw.get("Plaid Removed?")).strip().lower() == "yes"
                    or status == "removed",
                    source=_optional_str(raw.get("Sync Source")) or "workbook",
                    plaid_transaction_id=_optional_str(raw.get("Plaid Transaction ID")),
                    plaid_pending_transaction_id=_optional_str(raw.get("Plaid Pending Transaction ID")),
                    raw=raw,
                )
            )
        return transactions

    def read_sync_states(self) -> list[SyncItemState]:
        if SYNC_STATE_SHEET not in self.workbook.sheetnames:
            return []
        ws = self.workbook[SYNC_STATE_SHEET]
        states: list[SyncItemState] = []
        for row in range(2, ws.max_row + 1):
            item_id = _optional_str(ws.cell(row, 1).value)
            if not item_id:
                continue
            states.append(
                SyncItemState(
                    item_id=item_id,
                    institution_name=_optional_str(ws.cell(row, 2).value) or "",
                    cursor=_optional_str(ws.cell(row, 3).value) or "",
                    access_token_key=_optional_str(ws.cell(row, 4).value) or f"plaid:{item_id}",
                    last_successful_sync=_coerce_datetime(ws.cell(row, 5).value),
                    last_request_id=_optional_str(ws.cell(row, 6).value) or "",
                    updated_at=_coerce_datetime(ws.cell(row, 7).value),
                    error=_optional_str(ws.cell(row, 8).value) or "",
                )
            )
        return states

    def upsert_sync_state(self, state: SyncItemState) -> None:
        self._ensure_sync_state()
        ws = self.workbook[SYNC_STATE_SHEET]
        row = self._find_sync_state_row(state.item_id) or ws.max_row + 1
        values = [
            state.item_id,
            state.institution_name,
            state.cursor,
            state.access_token_key or f"plaid:{state.item_id}",
            _dt_value(state.last_successful_sync),
            state.last_request_id,
            _dt_value(state.updated_at) or _dt_value(datetime.now(timezone.utc)),
            state.error,
        ]
        for col, value in enumerate(values, start=1):
            ws.cell(row, col).value = value
        self._style_sync_state()

    def apply_plaid_sync(self, result: PlaidSyncResult) -> dict[str, int]:
        if not self._has_runtime_structure():
            self.ensure_structure()
        ws = self.workbook[ACTUAL_CHARGES_SHEET]
        headers = self._headers(ws)
        budget_categories = {budget.category for budget in self.read_budgets()}
        merchant_history = build_merchant_history(self._read_transaction_rows())
        by_transaction_id = self._index_by_header(headers, "Plaid Transaction ID")
        by_pending_transaction_id = self._index_by_header(headers, "Plaid Pending Transaction ID")
        account_details = {
            account.get("account_id"): {
                "name": account.get("name") or account.get("official_name") or "",
                "type": account.get("type") or "",
                "subtype": account.get("subtype") or "",
            }
            for account in result.accounts
        }
        counts = {"added": 0, "modified": 0, "removed": 0}

        for tx in result.added:
            row = self._target_row_for_plaid_transaction(tx, by_transaction_id, by_pending_transaction_id)
            if row is None:
                row = ws.max_row + 1
                counts["added"] += 1
            else:
                counts["modified"] += 1
            self._write_plaid_transaction_row(
                ws,
                headers,
                row,
                tx,
                result.item_id,
                account_details,
                merchant_history,
                budget_categories,
                preserve_selected=True,
            )

        for tx in result.modified:
            row = self._target_row_for_plaid_transaction(tx, by_transaction_id, by_pending_transaction_id)
            if row is None:
                row = ws.max_row + 1
                counts["added"] += 1
            else:
                counts["modified"] += 1
            self._write_plaid_transaction_row(
                ws,
                headers,
                row,
                tx,
                result.item_id,
                account_details,
                merchant_history,
                budget_categories,
                preserve_selected=True,
            )

        for removed in result.removed:
            transaction_id = removed.get("transaction_id")
            row = by_transaction_id.get(transaction_id)
            if row:
                ws.cell(row, headers["Plaid Removed?"]).value = "Yes"
                ws.cell(row, headers["Plaid Status"]).value = "removed"
                ws.cell(row, headers["Last Synced At"]).value = _dt_value(datetime.now(timezone.utc))
                counts["removed"] += 1

        self._refresh_budget_treatments(ws, headers)
        self._resize_actual_table(ws)
        self._ensure_monthly_budget(self.read_budgets(), date.today())
        return counts

    def _has_runtime_structure(self) -> bool:
        required_sheets = {ACTUAL_CHARGES_SHEET, CATEGORY_MAP_SHEET, MONTHLY_BUDGET_SHEET, SYNC_STATE_SHEET}
        if not required_sheets.issubset(self.workbook.sheetnames):
            return False
        headers = self._headers(self.workbook[ACTUAL_CHARGES_SHEET])
        return all(header in headers for header in EXTENSION_HEADERS)

    def last_sync_context(self) -> tuple[datetime | None, str]:
        states = self.read_sync_states()
        if not states:
            return None, "not connected"
        failures = [state.error for state in states if state.error]
        latest = max((state.last_successful_sync for state in states if state.last_successful_sync), default=None)
        if failures:
            return latest, "review sync errors"
        return latest, "synced" if latest else "connected, not synced"

    def _ensure_actual_charges_columns(self, budgets: list[CategoryBudget]) -> bool:
        ws = self.workbook[ACTUAL_CHARGES_SHEET]
        headers = self._headers(ws)
        changed = False
        next_col = ws.max_column + 1
        for header in EXTENSION_HEADERS:
            if header not in headers:
                ws.cell(5, next_col).value = header
                headers[header] = next_col
                next_col += 1
                changed = True

        changed |= self._refresh_budget_treatments(ws, headers)
        selected_col = headers["Selected Category"]
        saved_col = headers["Saved Category"]
        personal_col = headers["Personal Budget?"]
        treatment_col = headers["Budget Treatment"]
        for row in range(6, ws.max_row + 1):
            if ws.cell(row, selected_col).value:
                continue
            if ws.cell(row, treatment_col).value != TreatmentDisposition.INCLUDED.value:
                continue
            saved_category = ws.cell(row, saved_col).value
            if saved_category:
                selected = canonical_budget_category(str(saved_category))
                if selected and not is_budget_excluded(selected):
                    ws.cell(row, selected_col).value = selected
                elif is_budget_excluded(selected):
                    ws.cell(row, personal_col).value = "No"
                changed = True

        self._apply_actual_charges_style(ws, headers, budgets)
        self._resize_actual_table(ws)
        return changed

    def _refresh_budget_treatments(self, ws, headers: dict[str, int]) -> bool:
        row_indices = [
            row
            for row in range(6, ws.max_row + 1)
            if ws.cell(row, headers.get("Date", 1)).value or ws.cell(row, headers.get("Merchant", 3)).value
        ]
        transactions = self._read_transaction_rows()
        treatments = classify_budget_treatments(transactions)
        if len(row_indices) != len(treatments):
            raise ValueError("budget treatment classification did not reconcile to workbook rows")
        changed = False
        treatment_col = headers["Budget Treatment"]
        basis_col = headers["Treatment Basis"]
        suggested_col = headers["Suggested Category"]
        suggestion_reason_col = headers["Suggestion Reason"]
        for row, treatment in zip(row_indices, treatments, strict=True):
            for column, value in (
                (treatment_col, treatment.disposition.value),
                (basis_col, treatment.visible_basis),
            ):
                if ws.cell(row, column).value != value:
                    ws.cell(row, column).value = value
                    changed = True
            if treatment.suppress_category_suggestion:
                for column in (suggested_col, suggestion_reason_col):
                    if ws.cell(row, column).value not in (None, ""):
                        ws.cell(row, column).value = ""
                        changed = True
        return changed

    def _ensure_category_map(self, budgets: list[CategoryBudget]) -> bool:
        if CATEGORY_MAP_SHEET in self.workbook.sheetnames:
            ws = self.workbook[CATEGORY_MAP_SHEET]
            ws.delete_rows(1, ws.max_row)
        else:
            ws = self.workbook.create_sheet(CATEGORY_MAP_SHEET)
        rows = build_category_map([budget.category for budget in budgets])
        headers = ["Source Category", "Budget Category", "Personal Budget?", "Notes", "Dropdown Categories"]
        ws.append(headers)
        dropdown_categories = [budget.category for budget in budgets]
        for idx, row in enumerate(rows, start=2):
            ws.cell(idx, 1).value = row[0]
            ws.cell(idx, 2).value = row[1]
            ws.cell(idx, 3).value = row[2]
            ws.cell(idx, 4).value = row[3]
        for idx, category in enumerate(dropdown_categories, start=2):
            ws.cell(idx, 5).value = category
        _style_header(ws, 1, 1, 5)
        widths = [34, 34, 18, 52, 34]
        for col, width in enumerate(widths, start=1):
            ws.column_dimensions[_col_letter(col)].width = width
        ws.freeze_panes = "A2"
        return True

    def _ensure_sync_state(self) -> bool:
        changed = False
        if SYNC_STATE_SHEET not in self.workbook.sheetnames:
            ws = self.workbook.create_sheet(SYNC_STATE_SHEET)
            changed = True
        else:
            ws = self.workbook[SYNC_STATE_SHEET]
        for col, header in enumerate(SYNC_HEADERS, start=1):
            if ws.cell(1, col).value != header:
                ws.cell(1, col).value = header
                changed = True
        ws["J1"] = "Migration Version"
        ws["K1"] = SYNC_MIGRATION_VERSION
        ws.sheet_state = "hidden"
        self._style_sync_state()
        return changed

    def _ensure_monthly_budget(self, budgets: list[CategoryBudget], as_of: date) -> bool:
        actual_ws = self.workbook[ACTUAL_CHARGES_SHEET]
        actual_headers = self._headers(actual_ws)
        actual_end_row = max(actual_ws.max_row, 6)
        budget_source_rows = self._read_budgets_with_source_rows()
        if [budget for _, budget in budget_source_rows] != budgets:
            raise ValueError("Manual Budget rows changed while Monthly Budget was being rebuilt")
        selected_month_value: Any = None
        if MONTHLY_BUDGET_SHEET in self.workbook.sheetnames:
            existing = self.workbook[MONTHLY_BUDGET_SHEET]
            selected_month_value = existing["B2"].value
            existing.delete_rows(1, existing.max_row)
            ws = existing
        else:
            ws = self.workbook.create_sheet(MONTHLY_BUDGET_SHEET, 0)
        selected_month = _coerce_date(selected_month_value) if selected_month_value else _month_start(as_of)
        ws["A1"] = "Monthly Budget"
        ws["A1"].font = Font(bold=True, size=16)
        ws["A2"] = "Selected Month"
        ws["B2"] = selected_month
        ws["B2"].number_format = "yyyy-mm"
        ws["D2"] = "Selected Category drives totals; confirmed non-spend rows marked Excluded are omitted."
        headers = [
            "Cost Type",
            "Category",
            "Monthly Budget",
            "Posted Spend",
            "Pending Impact",
            "Remaining",
            "% Used",
            "Review Needed",
            "Status",
        ]
        for col, header in enumerate(headers, start=1):
            ws.cell(5, col).value = header
        _style_header(ws, 5, 1, len(headers))
        start_row = 6
        for idx, (manual_row, budget) in enumerate(budget_source_rows, start=start_row):
            ws.cell(idx, 1).value = f"='Manual Budget'!A{manual_row}"
            ws.cell(idx, 2).value = f"='Manual Budget'!B{manual_row}"
            ws.cell(idx, 3).value = f"='Manual Budget'!C{manual_row}"
            ws.cell(idx, 4).value = _posted_formula(idx, actual_headers, actual_end_row)
            ws.cell(idx, 5).value = _pending_formula(idx, actual_headers, actual_end_row)
            ws.cell(idx, 6).value = f"=C{idx}-D{idx}"
            ws.cell(idx, 7).value = f'=IF(C{idx}=0,0,D{idx}/C{idx})'
            ws.cell(idx, 8).value = _review_formula(idx, actual_headers, actual_end_row)
            ws.cell(idx, 9).value = f'=IF(AND(C{idx}=0,D{idx}>0),"NO BUDGET",IF(F{idx}<0,"OVER",IF(G{idx}>=0.85,"WATCH","OK")))'
        total_row = start_row + len(budgets)
        ws.cell(total_row, 1).value = "TOTAL"
        for col in (3, 4, 5, 6, 8):
            ws.cell(total_row, col).value = f"=SUM({_col_letter(col)}{start_row}:{_col_letter(col)}{total_row - 1})"
        ws.cell(total_row, 7).value = f"=IF(C{total_row}=0,0,D{total_row}/C{total_row})"
        ws.cell(total_row, 9).value = f'=IF(F{total_row}<0,"OVER","OK")'
        for row in range(start_row, total_row + 1):
            for col in (3, 4, 5, 6):
                ws.cell(row, col).number_format = '"$"#,##0.00;[Red]("$"#,##0.00);-'
            ws.cell(row, 7).number_format = "0.0%"
        ws.freeze_panes = "A6"
        widths = [16, 34, 16, 16, 16, 16, 12, 16, 12]
        for col, width in enumerate(widths, start=1):
            ws.column_dimensions[_col_letter(col)].width = width
        ws.conditional_formatting.add(
            f"F{start_row}:F{total_row}",
            CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor="F4CCCC")),
        )
        ws.conditional_formatting.add(
            f"I{start_row}:I{total_row}",
            CellIsRule(operator="equal", formula=['"OVER"'], fill=PatternFill("solid", fgColor="F4CCCC")),
        )
        return True

    def _headers(self, ws) -> dict[str, int]:
        return {
            str(ws.cell(5, col).value): col
            for col in range(1, ws.max_column + 1)
            if ws.cell(5, col).value
        }

    def _manual_budget_amount(self, row: int) -> Money:
        return _cell_money(self.workbook[MANUAL_BUDGET_SHEET], row, 3)

    def _apply_actual_charges_style(self, ws, headers: dict[str, int], budgets: list[CategoryBudget]) -> None:
        _style_header(ws, 5, 1, ws.max_column)
        ws.freeze_panes = "A6"
        for col, width in {
            "A": 12,
            "B": 12,
            "C": 28,
            "D": 28,
            "E": 26,
            "G": 12,
            "M": 30,
            "N": 30,
            "O": 44,
            "P": 28,
            "T": 14,
            "U": 14,
        }.items():
            ws.column_dimensions[col].width = width
        for header, width in {
            "Budget Treatment": 18,
            "Treatment Basis": 32,
            "Plaid Original Description": 42,
            "Plaid Transaction Code": 22,
            "Plaid Account Type": 20,
            "Plaid Account Subtype": 22,
        }.items():
            if header in headers:
                ws.column_dimensions[_col_letter(headers[header])].width = width
        ws.auto_filter.ref = f"A5:{_col_letter(ws.max_column)}{max(ws.max_row, 6)}"
        selected_col = headers["Selected Category"]
        formula = f"='{CATEGORY_MAP_SHEET}'!$E$2:$E${len(budgets) + 1}"
        selected_letter = _col_letter(selected_col)
        validation_end_row = max(MAX_FORMULA_ROWS, ws.max_row)
        target_range = f"{selected_letter}6:{selected_letter}{validation_end_row}"
        _replace_managed_category_validation(ws, formula, target_range)

    def _resize_actual_table(self, ws) -> None:
        if ACTUAL_TABLE_NAME in ws.tables:
            ws.tables[ACTUAL_TABLE_NAME].ref = f"A5:{_col_letter(ws.max_column)}{max(ws.max_row, 6)}"

    def _index_by_header(self, headers: dict[str, int], header: str) -> dict[str, int]:
        ws = self.workbook[ACTUAL_CHARGES_SHEET]
        col = headers[header]
        index: dict[str, int] = {}
        for row in range(6, ws.max_row + 1):
            value = _optional_str(ws.cell(row, col).value)
            if value:
                index[value] = row
        return index

    def _target_row_for_plaid_transaction(
        self,
        tx: dict[str, Any],
        by_transaction_id: dict[str, int],
        by_pending_transaction_id: dict[str, int],
    ) -> int | None:
        transaction_id = tx.get("transaction_id")
        if transaction_id and transaction_id in by_transaction_id:
            return by_transaction_id[transaction_id]
        pending_id = tx.get("pending_transaction_id")
        if pending_id and pending_id in by_transaction_id:
            return by_transaction_id[pending_id]
        if pending_id and pending_id in by_pending_transaction_id:
            return by_pending_transaction_id[pending_id]
        return None

    def _write_plaid_transaction_row(
        self,
        ws,
        headers: dict[str, int],
        row: int,
        tx: dict[str, Any],
        item_id: str,
        account_details: dict[str, dict[str, Any]],
        merchant_history: dict[str, str],
        budget_categories: set[str],
        *,
        preserve_selected: bool,
    ) -> None:
        existing_selected = ws.cell(row, headers["Selected Category"]).value if preserve_selected else None
        pfc = tx.get("personal_finance_category") or {}
        tx_date = _parse_iso_date(tx.get("date"))
        month = _month_start(tx_date)
        merchant = tx.get("merchant_name") or tx.get("name") or ""
        account_id = tx.get("account_id") or ""
        account = account_details.get(account_id) or {}
        amount = _money(tx.get("amount"))
        existing_refund = ws.cell(row, headers["Refund?"]).value if preserve_selected else None
        classification_payload = dict(tx)
        classification_payload["account_name"] = account.get("name") or account_id
        classification_payload["account_type"] = account.get("type") or ""
        classification_payload["account_subtype"] = account.get("subtype") or ""
        treatment = classify_provider_payload(classification_payload)
        suggestion, reason = infer_suggestion(tx, merchant_history, budget_categories)
        values = {
            "Date": tx_date,
            "Month": month,
            "Merchant": merchant,
            "Account": account.get("name") or account_id,
            "Saved Category": "",
            "Cost Type": "",
            "Amount": amount,
            "Personal Budget?": "Yes",
            "Cancel?": "No",
            "Refund?": existing_refund if existing_refund not in (None, "") else "No",
            "Won't Recur?": "No",
            "Evidence": "Plaid",
            "Selected Category": existing_selected or "",
            "Suggested Category": suggestion or "",
            "Suggestion Reason": reason,
            "Plaid Transaction ID": tx.get("transaction_id") or "",
            "Plaid Pending Transaction ID": tx.get("pending_transaction_id") or "",
            "Plaid Account ID": account_id,
            "Plaid Item ID": item_id,
            "Plaid Status": "pending" if tx.get("pending") else "posted",
            "Plaid Removed?": "No",
            "Sync Source": "Plaid",
            "Last Synced At": _dt_value(datetime.now(timezone.utc)),
            "Plaid PFC Primary": pfc.get("primary") or "",
            "Plaid PFC Detailed": pfc.get("detailed") or "",
            "Plaid PFC Confidence": pfc.get("confidence_level") or "",
            "Budget Treatment": treatment.disposition.value,
            "Treatment Basis": treatment.visible_basis,
            "Plaid Original Description": tx.get("original_description") or "",
            "Plaid Transaction Code": tx.get("transaction_code") or "",
            "Plaid Account Type": account.get("type") or "",
            "Plaid Account Subtype": account.get("subtype") or "",
        }
        if is_budget_excluded(values["Selected Category"] or values["Suggested Category"]):
            values["Personal Budget?"] = "No"
        for header, value in values.items():
            if header in headers:
                set_literal_cell(ws.cell(row, headers[header]), value)
        ws.cell(row, headers["Date"]).number_format = "yyyy-mm-dd"
        ws.cell(row, headers["Month"]).number_format = "yyyy-mm"
        ws.cell(row, headers["Amount"]).number_format = '"$"#,##0.00;[Red]("$"#,##0.00);-'

    def _find_sync_state_row(self, item_id: str) -> int | None:
        ws = self.workbook[SYNC_STATE_SHEET]
        for row in range(2, ws.max_row + 1):
            if ws.cell(row, 1).value == item_id:
                return row
        return None

    def _style_sync_state(self) -> None:
        ws = self.workbook[SYNC_STATE_SHEET]
        _style_header(ws, 1, 1, len(SYNC_HEADERS))
        for col in range(1, len(SYNC_HEADERS) + 1):
            ws.column_dimensions[_col_letter(col)].width = 24


def _posted_formula(row: int, headers: dict[str, int], end_row: int) -> str:
    selected = _actual_charges_range(headers, "Selected Category", end_row)
    month = _actual_charges_range(headers, "Month", end_row)
    personal = _actual_charges_range(headers, "Personal Budget?", end_row)
    removed = _actual_charges_range(headers, "Plaid Removed?", end_row)
    treatment = _actual_charges_range(headers, "Budget Treatment", end_row)
    status = _actual_charges_range(headers, "Plaid Status", end_row)
    amount = _actual_charges_range(headers, "Amount", end_row)
    return (
        f"=SUMPRODUCT(({selected}=$B{row})*"
        f"({month}=$B$2)*"
        f"({personal}=\"Yes\")*"
        f"({removed}<>\"Yes\")*"
        f"({treatment}<>\"{EXCLUDED_TREATMENT_VALUE}\")*"
        f"(({status}=\"posted\")+"
        f"({status}=\"\"))*"
        f"{amount})"
    )


def _pending_formula(row: int, headers: dict[str, int], end_row: int) -> str:
    selected = _actual_charges_range(headers, "Selected Category", end_row)
    month = _actual_charges_range(headers, "Month", end_row)
    personal = _actual_charges_range(headers, "Personal Budget?", end_row)
    removed = _actual_charges_range(headers, "Plaid Removed?", end_row)
    treatment = _actual_charges_range(headers, "Budget Treatment", end_row)
    status = _actual_charges_range(headers, "Plaid Status", end_row)
    amount = _actual_charges_range(headers, "Amount", end_row)
    return (
        f"=SUMPRODUCT(({selected}=$B{row})*"
        f"({month}=$B$2)*"
        f"({personal}=\"Yes\")*"
        f"({removed}<>\"Yes\")*"
        f"({treatment}<>\"{EXCLUDED_TREATMENT_VALUE}\")*"
        f"({status}=\"pending\")*"
        f"{amount})"
    )


def _review_formula(row: int, headers: dict[str, int], end_row: int) -> str:
    month = _actual_charges_range(headers, "Month", end_row)
    personal = _actual_charges_range(headers, "Personal Budget?", end_row)
    removed = _actual_charges_range(headers, "Plaid Removed?", end_row)
    treatment = _actual_charges_range(headers, "Budget Treatment", end_row)
    selected = _actual_charges_range(headers, "Selected Category", end_row)
    suggested = _actual_charges_range(headers, "Suggested Category", end_row)
    return (
        f"=COUNTIFS({month},$B$2,"
        f"{personal},\"Yes\","
        f"{removed},\"<>Yes\","
        f"{treatment},\"<>{EXCLUDED_TREATMENT_VALUE}\","
        f"{selected},\"\","
        f"{suggested},$B{row})"
    )


def _actual_charges_range(headers: dict[str, int], header: str, end_row: int) -> str:
    if header not in headers:
        raise ValueError(f"Actual Charges is missing required header: {header}")
    column = _col_letter(headers[header])
    return f"'{ACTUAL_CHARGES_SHEET}'!${column}$6:${column}${max(end_row, 6)}"


def _style_header(ws, row: int, start_col: int, end_col: int) -> None:
    fill = PatternFill("solid", fgColor="1F4E78")
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row, col)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center")


def _replace_managed_category_validation(ws, formula: str, target_range: str) -> None:
    managed = [
        validation
        for validation in ws.data_validations.dataValidation
        if _is_managed_category_validation(validation)
    ]
    if (
        len(managed) == 1
        and managed[0].type == "list"
        and managed[0].formula1 == formula
        and str(managed[0].sqref) == target_range
    ):
        return
    for validation in managed:
        ws.data_validations.dataValidation.remove(validation)
    validation = DataValidation(type="list", formula1=formula, allow_blank=True)
    validation.error = "Choose a category from Category Map."
    validation.errorTitle = "Invalid category"
    validation.prompt = "Only Selected Category affects the Monthly Budget."
    validation.promptTitle = "Budget category"
    ws.add_data_validation(validation)
    validation.add(target_range)


def _is_managed_category_validation(validation: DataValidation) -> bool:
    if validation.type != "list":
        return False
    return (
        CATEGORY_MAP_SHEET in str(validation.formula1 or "")
        or validation.promptTitle == "Budget category"
        or validation.errorTitle == "Invalid category"
        or validation.prompt == "Only Selected Category affects the Monthly Budget."
    )


def _col_letter(col: int) -> str:
    result = ""
    while col:
        col, remainder = divmod(col - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _optional_str(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def _money(value: Any) -> Money:
    if value is None or value == "":
        return Decimal("0")
    return Decimal(str(value))


def _cell_money(ws, row: int, col: int) -> Money:
    value = ws.cell(row, col).value
    if value is None or value == "":
        return Decimal("0")
    if isinstance(value, str) and value.startswith("="):
        formula = value.replace("$", "").replace(" ", "")
        sum_match = re.fullmatch(r"=SUM\(C(\d+):C(\d+)\)(?:/(\d+(?:\.\d+)?))?", formula)
        if sum_match:
            start, end, divisor = sum_match.groups()
            total = sum((_cell_money(ws, idx, col) for idx in range(int(start), int(end) + 1)), Decimal("0"))
            if divisor:
                return total / Decimal(divisor)
            return total
        division_match = re.fullmatch(r"=(\d+(?:\.\d+)?)/(\d+(?:\.\d+)?)", formula)
        if division_match:
            left, right = division_match.groups()
            return Decimal(left) / Decimal(right)
        ref_match = re.fullmatch(r"=C(\d+)", formula)
        if ref_match:
            return _cell_money(ws, int(ref_match.group(1)), col)
        raise ValueError(f"Unsupported Manual Budget formula at {ws.cell(row, col).coordinate}: {value}")
    return _money(value)


def _coerce_date(value: Any) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str) and value:
        return _parse_iso_date(value[:10])
    return date.today()


def _parse_iso_date(value: Any) -> date:
    if isinstance(value, date):
        return value if not isinstance(value, datetime) else value.date()
    return datetime.strptime(str(value), "%Y-%m-%d").date()


def _month_start(value: date) -> date:
    return date(value.year, value.month, 1)


def _coerce_datetime(value: Any) -> datetime | None:
    if isinstance(value, datetime):
        return value
    if isinstance(value, str) and value:
        try:
            return datetime.fromisoformat(value)
        except ValueError:
            return None
    return None


def _dt_value(value: datetime | None) -> str:
    if value is None:
        return ""
    if value.tzinfo is None:
        value = value.replace(tzinfo=timezone.utc)
    return value.isoformat()
