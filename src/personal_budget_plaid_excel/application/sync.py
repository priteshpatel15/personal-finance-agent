from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from personal_budget_plaid_excel.domain.models import SyncItemState
from personal_budget_plaid_excel.infrastructure.telemetry import Telemetry
from personal_budget_plaid_excel.infrastructure.workbook import BudgetWorkbook
from personal_budget_plaid_excel.ports.plaid import PlaidPort
from personal_budget_plaid_excel.ports.secrets import SecretStore


class SyncService:
    def __init__(
        self,
        workbook_path: Path,
        plaid: PlaidPort,
        secrets: SecretStore,
        telemetry: Telemetry | None = None,
    ) -> None:
        self.workbook_path = workbook_path
        self.plaid = plaid
        self.secrets = secrets
        self.telemetry = telemetry or Telemetry()

    def sync_all_items(self) -> dict[str, int]:
        workbook = BudgetWorkbook(self.workbook_path)
        workbook.ensure_structure()
        states = workbook.read_sync_states()
        totals = {"items": len(states), "added": 0, "modified": 0, "removed": 0, "errors": 0}
        for state in states:
            try:
                access_token = self.secrets.get_secret(state.access_token_key or f"plaid:{state.item_id}")
                result = self.plaid.sync_item(state, access_token)
                counts = workbook.apply_plaid_sync(result)
                totals["added"] += counts["added"]
                totals["modified"] += counts["modified"]
                totals["removed"] += counts["removed"]
                workbook.upsert_sync_state(
                    SyncItemState(
                        item_id=state.item_id,
                        institution_name=state.institution_name,
                        cursor=result.next_cursor,
                        access_token_key=state.access_token_key,
                        last_successful_sync=datetime.now(timezone.utc),
                        last_request_id=",".join(result.request_ids[-3:]),
                        updated_at=datetime.now(timezone.utc),
                    )
                )
                self.telemetry.event(
                    "plaid_sync_item_success",
                    institution_name=state.institution_name,
                    counts=counts,
                )
            except Exception as exc:  # keep syncing other linked items
                totals["errors"] += 1
                diagnostics = _error_diagnostics(exc)
                workbook.upsert_sync_state(
                    SyncItemState(
                        item_id=state.item_id,
                        institution_name=state.institution_name,
                        cursor=state.cursor,
                        access_token_key=state.access_token_key,
                        last_successful_sync=state.last_successful_sync,
                        last_request_id=state.last_request_id,
                        updated_at=datetime.now(timezone.utc),
                        error=_error_summary(diagnostics),
                    )
                )
                self.telemetry.event(
                    "plaid_sync_item_error",
                    institution_name=state.institution_name,
                    **diagnostics,
                )
        workbook.save_with_backup(backup=True)
        self.telemetry.event("plaid_sync_complete", totals=totals)
        return totals


def _error_diagnostics(exc: Exception) -> dict[str, Any]:
    diagnostics: dict[str, Any] = {"error_type": type(exc).__name__}
    error_code = getattr(exc, "error_code", None)
    if isinstance(error_code, str) and error_code:
        diagnostics["error_code"] = error_code
    status = getattr(exc, "status", None)
    if isinstance(status, int):
        diagnostics["status"] = status
    return diagnostics


def _error_summary(diagnostics: dict[str, Any]) -> str:
    details: list[str] = []
    if diagnostics.get("error_code"):
        details.append(f"code={diagnostics['error_code']}")
    if diagnostics.get("status") is not None:
        details.append(f"status={diagnostics['status']}")
    suffix = f" ({', '.join(details)})" if details else ""
    return f"{diagnostics['error_type']}{suffix}"
