from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from personal_budget_plaid_excel.application.sync import SyncService
from personal_budget_plaid_excel.domain.models import PlaidSyncResult, SyncItemState
from personal_budget_plaid_excel.infrastructure.plaid_rest import PlaidApiError
from personal_budget_plaid_excel.infrastructure.telemetry import Telemetry
from personal_budget_plaid_excel.infrastructure.workbook import BudgetWorkbook


from synthetic_workbook import WORKBOOK_NAME, write_synthetic_workbook
SOURCE_WORKBOOK = Path(WORKBOOK_NAME)


class SyncServiceTelemetryTests(unittest.TestCase):
    def test_success_telemetry_contains_safe_context_and_aggregate_counts_only(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            workbook_path = _workbook_with_single_state(root)
            log_path = root / "telemetry.jsonl"
            service = SyncService(
                workbook_path,
                _SuccessfulPlaid(),
                _StaticSecrets(),
                Telemetry(log_path),
            )

            totals = service.sync_all_items()
            entries = [json.loads(line) for line in log_path.read_text(encoding="utf-8").splitlines()]
            serialized = log_path.read_text(encoding="utf-8")

        self.assertEqual(totals, {"items": 1, "added": 1, "modified": 0, "removed": 0, "errors": 0})
        success = next(entry for entry in entries if entry["event"] == "plaid_sync_item_success")
        self.assertEqual(
            success["payload"],
            {
                "institution_name": "Safe Institution",
                "counts": {"added": 1, "modified": 0, "removed": 0},
            },
        )
        for sensitive_value in (
            "sensitive-item-id",
            "sensitive-account-id",
            "sensitive-cursor",
            "sensitive-next-cursor",
            "sensitive-request-id",
            "Sensitive Merchant",
            "91.73",
            "sensitive-access-token",
        ):
            self.assertNotIn(sensitive_value, serialized)

    def test_error_telemetry_and_workbook_state_exclude_provider_exception_text(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            workbook_path = _workbook_with_single_state(root)
            log_path = root / "telemetry.jsonl"
            service = SyncService(
                workbook_path,
                _FailingPlaid(),
                _StaticSecrets(),
                Telemetry(log_path),
            )

            totals = service.sync_all_items()
            entries = [json.loads(line) for line in log_path.read_text(encoding="utf-8").splitlines()]
            [state] = BudgetWorkbook(workbook_path).read_sync_states()
            serialized = log_path.read_text(encoding="utf-8")

        self.assertEqual(totals["errors"], 1)
        failure = next(entry for entry in entries if entry["event"] == "plaid_sync_item_error")
        self.assertEqual(
            failure["payload"],
            {
                "institution_name": "Safe Institution",
                "error_type": "PlaidApiError",
                "error_code": "ITEM_LOGIN_REQUIRED",
                "status": 400,
            },
        )
        self.assertEqual(state.error, "PlaidApiError (code=ITEM_LOGIN_REQUIRED, status=400)")
        self.assertNotIn("provider raw text with sensitive-access-token", serialized)
        self.assertNotIn("sensitive-item-id", serialized)


class _StaticSecrets:
    def get_secret(self, account: str) -> str:
        return "sensitive-access-token"


class _SuccessfulPlaid:
    def sync_item(self, state: SyncItemState, access_token: str) -> PlaidSyncResult:
        return PlaidSyncResult(
            item_id=state.item_id,
            added=[
                {
                    "transaction_id": "sensitive-transaction-id",
                    "account_id": "sensitive-account-id",
                    "amount": 91.73,
                    "date": "2026-08-28",
                    "merchant_name": "Sensitive Merchant",
                }
            ],
            modified=[],
            removed=[],
            accounts=[{"account_id": "sensitive-account-id", "name": "Test Account"}],
            next_cursor="sensitive-next-cursor",
            request_ids=["sensitive-request-id"],
        )


class _FailingPlaid:
    def sync_item(self, state: SyncItemState, access_token: str) -> PlaidSyncResult:
        raise PlaidApiError(
            "provider raw text with sensitive-access-token",
            error_code="ITEM_LOGIN_REQUIRED",
            status=400,
        )


def _workbook_with_single_state(root: Path) -> Path:
    workbook_path = root / SOURCE_WORKBOOK.name
    write_synthetic_workbook(workbook_path)
    workbook = BudgetWorkbook(workbook_path)
    workbook.ensure_structure()
    sheet = workbook.workbook["_SyncState"]
    if sheet.max_row > 1:
        sheet.delete_rows(2, sheet.max_row - 1)
    workbook.upsert_sync_state(
        SyncItemState(
            "sensitive-item-id",
            "Safe Institution",
            cursor="sensitive-cursor",
            access_token_key="plaid:sensitive-item-id",
        )
    )
    workbook.save_with_backup(backup=False)
    return workbook_path


if __name__ == "__main__":
    unittest.main()
