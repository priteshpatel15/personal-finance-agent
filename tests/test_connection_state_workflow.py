from __future__ import annotations

import json
import shutil
import tempfile
import unittest
from datetime import datetime, timedelta, timezone
from pathlib import Path

from personal_budget_plaid_excel.application.connection_states import reconcile_sync_states
from personal_budget_plaid_excel.cli.main import main
from personal_budget_plaid_excel.domain.models import SyncItemState
from personal_budget_plaid_excel.infrastructure.connection_registry import ConnectionRegistry
from personal_budget_plaid_excel.infrastructure.workbook import BudgetWorkbook


from synthetic_workbook import WORKBOOK_NAME, write_synthetic_workbook
SOURCE_WORKBOOK = Path(WORKBOOK_NAME)


class ConnectionStateWorkflowTests(unittest.TestCase):
    def test_reconcile_prefers_freshest_state_and_unions_missing_connections(self) -> None:
        older = datetime(2026, 8, 20, tzinfo=timezone.utc)
        newer = older + timedelta(days=1)
        target = [
            SyncItemState(
                "shared-item",
                "Target Institution",
                cursor="target-current",
                access_token_key="plaid:shared-item",
                updated_at=newer,
            ),
            SyncItemState("target-only", "Target Only", cursor="target-only-cursor"),
        ]
        registry = [
            SyncItemState(
                "shared-item",
                "Registry Institution",
                cursor="registry-stale",
                access_token_key="plaid:shared-item",
                updated_at=older,
            ),
            SyncItemState("registry-only", "Registry Only", cursor="registry-only-cursor"),
        ]
        authority = [
            SyncItemState("shared-item", "Authority Institution", cursor="authority-stale"),
            SyncItemState("authority-only", "Authority Only", cursor="authority-only-cursor"),
        ]

        resolved = {state.item_id: state for state in reconcile_sync_states(target, registry, authority)}

        self.assertEqual(resolved["shared-item"].cursor, "target-current")
        self.assertEqual(resolved["shared-item"].institution_name, "Target Institution")
        self.assertEqual(
            set(resolved),
            {"shared-item", "target-only", "registry-only", "authority-only"},
        )

    def test_reconcile_uses_source_precedence_when_freshness_is_tied(self) -> None:
        timestamp = datetime(2026, 8, 20, tzinfo=timezone.utc)
        target = SyncItemState("shared-item", "Target", cursor="target-cursor", updated_at=timestamp)
        registry = SyncItemState("shared-item", "Registry", cursor="registry-cursor", updated_at=timestamp)

        resolved = reconcile_sync_states([target], [registry])

        self.assertEqual(len(resolved), 1)
        self.assertEqual(resolved[0].cursor, "target-cursor")
        self.assertEqual(resolved[0].institution_name, "Target")

    def test_registry_round_trip_preserves_cursor_and_freshness_but_not_diagnostics(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            registry = ConnectionRegistry(Path(tmp) / "connection-registry.json")
            synced_at = datetime(2026, 8, 29, 12, 30, tzinfo=timezone.utc)
            registry.replace(
                [
                    SyncItemState(
                        "item-one",
                        "Institution One",
                        cursor="current-cursor",
                        access_token_key="plaid:item-one",
                        last_successful_sync=synced_at,
                        last_request_id="request-id-must-not-persist",
                        updated_at=synced_at,
                        error="raw provider diagnostic must not persist",
                    )
                ]
            )

            [state] = registry.read()
            serialized = registry.path.read_text(encoding="utf-8")

        self.assertEqual(state.cursor, "current-cursor")
        self.assertEqual(state.last_successful_sync, synced_at)
        self.assertEqual(state.updated_at, synced_at)
        self.assertEqual(state.last_request_id, "")
        self.assertEqual(state.error, "")
        self.assertNotIn("request-id-must-not-persist", serialized)
        self.assertNotIn("raw provider diagnostic", serialized)

    def test_existing_month_prefers_target_then_persists_refreshed_registry(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / SOURCE_WORKBOOK.name
            write_synthetic_workbook(source)
            older = datetime(2026, 8, 20, tzinfo=timezone.utc)
            newer = older + timedelta(days=1)

            source_workbook = BudgetWorkbook(source)
            _replace_workbook_states(
                source_workbook,
                [
                    SyncItemState("shared-item", "Shared", cursor="authority-stale", updated_at=older),
                    SyncItemState("authority-only", "Authority Only", cursor="authority-cursor"),
                ],
            )
            source_workbook.save_with_backup(backup=False)

            output_dir = root / "monthly"
            output_dir.mkdir()
            target = output_dir / "budget-2026-08.xlsx"
            shutil.copy2(source, target)
            target_workbook = BudgetWorkbook(target)
            _replace_workbook_states(
                target_workbook,
                [SyncItemState("shared-item", "Shared", cursor="target-current", updated_at=newer)],
            )
            target_workbook.save_with_backup(backup=False)

            registry = ConnectionRegistry(output_dir / "connection-registry.json")
            registry.replace(
                [
                    SyncItemState("shared-item", "Shared", cursor="registry-stale", updated_at=older),
                    SyncItemState("registry-only", "Registry Only", cursor="registry-cursor"),
                ]
            )
            sync_json = root / "sync.json"
            sync_json.write_text(
                json.dumps(
                    {
                        "item_id": "registry-only",
                        "institution_name": "Registry Only",
                        "added": [],
                        "modified": [],
                        "removed": [],
                        "accounts": [],
                        "next_cursor": "registry-refreshed",
                        "request_ids": ["request-id-must-not-persist"],
                    }
                ),
                encoding="utf-8",
            )

            result = main(
                [
                    "--workbook",
                    str(source),
                    "monthly-run",
                    "--source",
                    "plaid",
                    "--month",
                    "2026-08",
                    "--output-dir",
                    str(output_dir),
                    "--sync-json",
                    str(sync_json),
                ]
            )

            workbook_states = {state.item_id: state for state in BudgetWorkbook(target).read_sync_states()}
            registry_states = {state.item_id: state for state in registry.read()}

        self.assertEqual(result, 0)
        self.assertEqual(workbook_states["shared-item"].cursor, "target-current")
        self.assertEqual(workbook_states["registry-only"].cursor, "registry-refreshed")
        self.assertIn("authority-only", workbook_states)
        self.assertEqual(registry_states["shared-item"].cursor, "target-current")
        self.assertEqual(registry_states["registry-only"].cursor, "registry-refreshed")
        self.assertIn("authority-only", registry_states)
        self.assertEqual(registry_states["registry-only"].last_request_id, "")


def _replace_workbook_states(workbook: BudgetWorkbook, states: list[SyncItemState]) -> None:
    workbook.ensure_structure()
    sheet = workbook.workbook["_SyncState"]
    if sheet.max_row > 1:
        sheet.delete_rows(2, sheet.max_row - 1)
    for state in states:
        workbook.upsert_sync_state(state)


if __name__ == "__main__":
    unittest.main()
