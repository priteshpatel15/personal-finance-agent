from __future__ import annotations

import unittest
from unittest.mock import patch

from personal_budget_plaid_excel.domain.models import SyncItemState
from personal_budget_plaid_excel.infrastructure.plaid_rest import PlaidApiError, PlaidConfig, PlaidRestClient


class PlaidRestClientTests(unittest.TestCase):
    def test_config_reads_keychain_when_environment_is_absent(self) -> None:
        class Store:
            def get_secret(self, account):
                return {"plaid-client-id": "client", "plaid-secret": "secret", "plaid-environment": "production"}[account]

        with patch.dict("os.environ", {}, clear=True):
            config = PlaidConfig.from_env(Store())

        self.assertEqual(config.client_id, "client")
        self.assertEqual(config.secret, "secret")
        self.assertEqual(config.environment, "production")

    def test_config_environment_overrides_keychain(self) -> None:
        class Store:
            def get_secret(self, account):
                raise AssertionError("keychain should not be consulted")

        with patch.dict("os.environ", {"PLAID_CLIENT_ID": "client", "PLAID_SECRET": "secret"}, clear=True):
            config = PlaidConfig.from_env(Store())

        self.assertEqual(config.client_id, "client")
    def test_sync_handles_pagination_and_accumulates_changes(self) -> None:
        calls = []

        def transport(path, payload):
            calls.append((path, payload.get("cursor")))
            if len(calls) == 1:
                return {
                    "added": [{"transaction_id": "t1"}],
                    "modified": [],
                    "removed": [],
                    "accounts": [{"account_id": "a1", "name": "Card"}],
                    "next_cursor": "c1",
                    "has_more": True,
                    "request_id": "r1",
                }
            return {
                "added": [],
                "modified": [{"transaction_id": "t2"}],
                "removed": [{"transaction_id": "t3"}],
                "accounts": [],
                "next_cursor": "c2",
                "has_more": False,
                "request_id": "r2",
            }

        client = PlaidRestClient(PlaidConfig("client", "secret"), transport=transport)
        result = client.sync_item(SyncItemState("item", cursor="c0"), "access-token")

        self.assertEqual([cursor for _, cursor in calls], ["c0", "c1"])
        self.assertEqual(result.next_cursor, "c2")
        self.assertEqual(len(result.added), 1)
        self.assertEqual(len(result.modified), 1)
        self.assertEqual(len(result.removed), 1)

    def test_sync_restarts_on_mutation_during_pagination(self) -> None:
        calls = []

        def transport(path, payload):
            calls.append(payload.get("cursor"))
            if len(calls) == 1:
                return {
                    "added": [{"transaction_id": "first-attempt"}],
                    "modified": [],
                    "removed": [],
                    "accounts": [],
                    "next_cursor": "page-2",
                    "has_more": True,
                    "request_id": "r1",
                }
            if len(calls) == 2:
                raise PlaidApiError("mutation", error_code="TRANSACTIONS_SYNC_MUTATION_DURING_PAGINATION")
            if len(calls) == 3:
                return {
                    "added": [{"transaction_id": "second-attempt"}],
                    "modified": [],
                    "removed": [],
                    "accounts": [],
                    "next_cursor": "page-2b",
                    "has_more": True,
                    "request_id": "r3",
                }
            return {
                "added": [],
                "modified": [],
                "removed": [],
                "accounts": [],
                "next_cursor": "done",
                "has_more": False,
                "request_id": "r4",
            }

        client = PlaidRestClient(PlaidConfig("client", "secret"), transport=transport)
        result = client.sync_item(SyncItemState("item", cursor="old"), "access-token")

        self.assertEqual(calls, ["old", "page-2", "old", "page-2b"])
        self.assertEqual([tx["transaction_id"] for tx in result.added], ["second-attempt"])
        self.assertEqual(result.next_cursor, "done")


if __name__ == "__main__":
    unittest.main()
