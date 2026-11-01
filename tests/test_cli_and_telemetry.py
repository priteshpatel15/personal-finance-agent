from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from personal_budget_plaid_excel.cli.main import main
from personal_budget_plaid_excel.cli.connect import LinkLifecycle, _link_html
from personal_budget_plaid_excel.domain.models import SyncItemState
from personal_budget_plaid_excel.infrastructure.telemetry import Telemetry, redact
from personal_budget_plaid_excel.infrastructure.workbook import BudgetWorkbook

from synthetic_workbook import WORKBOOK_NAME, write_synthetic_workbook
SOURCE_WORKBOOK = Path(WORKBOOK_NAME)


class CliAndTelemetryTests(unittest.TestCase):
    def test_link_lifecycle_records_redacted_stage_diagnostics(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "connection.jsonl"
            lifecycle = LinkLifecycle(Telemetry(path))
            lifecycle.record("link_opened")
            lifecycle.record("callback_failed", public_token="do-not-store", error_type="RuntimeError")
            entries = [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines()]

        self.assertEqual(lifecycle.snapshot(), {"stage": "callback_failed"})
        self.assertEqual([entry["payload"]["stage"] for entry in entries], ["link_opened", "callback_failed"])
        self.assertEqual(entries[-1]["payload"]["public_token"], "[REDACTED]")

    def test_link_html_exposes_visible_status_stages(self) -> None:
        html = _link_html("test-token")
        for stage in ("link_opened", "institution_selected", "share_data_completed"):
            self.assertIn(stage, html)
        self.assertIn("Connection saved", html)

    def test_cli_fake_sync_report_and_run_payload(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            copy = tmp_path / SOURCE_WORKBOOK.name
            write_synthetic_workbook(copy)
            sync_json = tmp_path / "sync.json"
            sync_json.write_text(
                json.dumps(
                    {
                        "item_id": "item-cli",
                        "accounts": [{"account_id": "acct-cli", "name": "CLI Card"}],
                        "added": [
                            {
                                "transaction_id": "cli-tx",
                                "account_id": "acct-cli",
                                "amount": 19.99,
                                "date": "2026-08-04",
                                "merchant_name": "CLI Store",
                                "pending": True,
                                "personal_finance_category": {
                                    "primary": "GENERAL_MERCHANDISE",
                                    "detailed": "GENERAL_MERCHANDISE_SUPERSTORES",
                                    "confidence_level": "HIGH",
                                },
                            }
                        ],
                        "next_cursor": "cursor-cli",
                    }
                ),
                encoding="utf-8",
            )

            self.assertEqual(main(["--workbook", str(copy), "doctor"]), 0)
            self.assertEqual(main(["--workbook", str(copy), "sync", "--sync-json", str(sync_json)]), 0)
            reports_dir = tmp_path / "reports"
            self.assertEqual(main(["--workbook", str(copy), "report", "--month", "2026-08", "--output-dir", str(reports_dir)]), 0)
            self.assertTrue(list(reports_dir.glob("budget-report-2026-08-*.html")))
            self.assertEqual(main(["--workbook", str(copy), "run", "--skip-sync", "--month", "2026-08", "--output-dir", str(reports_dir)]), 0)

    def test_telemetry_redacts_sensitive_fields(self) -> None:
        payload = redact(
            {
                "access_token": "secret",
                "nested": {"account_id": "abc"},
                "item_id": "item",
                "cursor": "cursor",
                "request_ids": ["request"],
                "merchant_name": "merchant",
                "amount": 12.34,
                "raw_payload": {"date": "2026-08-01"},
                "error": "raw provider message",
                "error_type": "PlaidApiError",
                "error_code": "ITEM_LOGIN_REQUIRED",
                "status": 400,
                "count": 2,
            }
        )
        self.assertEqual(payload["access_token"], "[REDACTED]")
        self.assertEqual(payload["nested"]["account_id"], "[REDACTED]")
        for key in ("item_id", "cursor", "request_ids", "merchant_name", "amount", "raw_payload", "error"):
            self.assertEqual(payload[key], "[REDACTED]")
        self.assertEqual(payload["error_type"], "PlaidApiError")
        self.assertEqual(payload["error_code"], "ITEM_LOGIN_REQUIRED")
        self.assertEqual(payload["status"], 400)
        self.assertEqual(payload["count"], 2)

        with tempfile.TemporaryDirectory() as tmp:
            log_path = Path(tmp) / "telemetry.jsonl"
            telemetry = Telemetry(log_path)
            telemetry.event("test", access_token="secret", count=1)
            line = json.loads(log_path.read_text(encoding="utf-8"))
            self.assertEqual(line["payload"]["access_token"], "[REDACTED]")

    def test_monthly_run_aggregates_three_connections_and_is_idempotent(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / SOURCE_WORKBOOK.name
            write_synthetic_workbook(source)
            source_workbook = BudgetWorkbook(source)
            for item_id, institution in (("amex", "American Express"), ("boa", "Bank of America"), ("chase", "Chase")):
                source_workbook.upsert_sync_state(SyncItemState(item_id, institution, access_token_key=f"plaid:{item_id}"))
            source_workbook.save_with_backup(backup=False)
            sync_json = root / "three-items.json"
            sync_json.write_text(json.dumps({"items": [_fake_item("amex", "American Express", "tx-amex", "2026-08-03"), _fake_item("boa", "Bank of America", "tx-boa", "2026-08-04"), _fake_item("chase", "Chase", "tx-chase", "2026-08-05")]}), encoding="utf-8")
            output_dir = root / "monthly"
            command = [
                "--workbook", str(source),
                "monthly-run", "--source", "plaid",
                "--month", "2026-08",
                "--output-dir", str(output_dir),
                "--sync-json", str(sync_json),
            ]

            self.assertEqual(main(command), 0)
            target = output_dir / "budget-2026-08.xlsx"
            workbook = BudgetWorkbook(target)
            self.assertTrue({"amex", "boa", "chase"}.issubset({state.item_id for state in workbook.read_sync_states()}))
            sheet = workbook.workbook["Actual Charges"]
            headers = [sheet.cell(5, column).value for column in range(1, sheet.max_column + 1)]
            transaction_col = headers.index("Plaid Transaction ID") + 1
            category_col = headers.index("Selected Category") + 1
            row = next(row for row in range(6, sheet.max_row + 1) if sheet.cell(row, transaction_col).value == "tx-amex")
            sheet.cell(row, category_col).value = "Coffee"
            workbook.save_with_backup(backup=False)

            self.assertEqual(main(command), 0)
            rerun = BudgetWorkbook(target).workbook["Actual Charges"]
            rows = [row for row in range(6, rerun.max_row + 1) if rerun.cell(row, transaction_col).value == "tx-amex"]
            self.assertEqual(len(rows), 1)
            self.assertEqual(rerun.cell(rows[0], category_col).value, "Coffee")
            self.assertTrue(list((output_dir / "backups").glob("budget-2026-08.*.xlsx")))

    def test_monthly_run_rolls_over_to_a_separate_monthly_workbook(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / SOURCE_WORKBOOK.name
            write_synthetic_workbook(source)
            sync_json = root / "september.json"
            sync_json.write_text(json.dumps(_fake_item("amex", "American Express", "tx-september", "2026-09-03")), encoding="utf-8")
            output_dir = root / "monthly"

            self.assertEqual(main([
                "--workbook", str(source),
                "monthly-run", "--source", "plaid",
                "--month", "2026-09",
                "--output-dir", str(output_dir),
                "--sync-json", str(sync_json),
            ]), 0)

            target = output_dir / "budget-2026-09.xlsx"
            self.assertTrue(target.exists())
            self.assertFalse((output_dir / "budget-2026-08.xlsx").exists())
            sheet = BudgetWorkbook(target).workbook["Actual Charges"]
            headers = [sheet.cell(5, column).value for column in range(1, sheet.max_column + 1)]
            transaction_col = headers.index("Plaid Transaction ID") + 1
            self.assertIn("tx-september", {sheet.cell(row, transaction_col).value for row in range(6, sheet.max_row + 1)})


if __name__ == "__main__":
    unittest.main()


def _fake_item(item_id: str, institution: str, transaction_id: str, transaction_date: str) -> dict:
    return {
        "item_id": item_id,
        "institution_name": institution,
        "accounts": [{"account_id": f"account-{item_id}", "name": f"{institution} Card"}],
        "added": [{"transaction_id": transaction_id, "account_id": f"account-{item_id}", "amount": 12.34, "date": transaction_date, "merchant_name": "Monthly Test"}],
        "next_cursor": f"cursor-{item_id}",
    }
