from __future__ import annotations

import json
import os
import urllib.error
import urllib.request
from dataclasses import dataclass
from typing import Any, Callable

from personal_budget_plaid_excel.domain.models import PlaidSyncResult, SyncItemState


class PlaidApiError(RuntimeError):
    def __init__(self, message: str, *, error_code: str = "", status: int | None = None) -> None:
        super().__init__(message)
        self.error_code = error_code
        self.status = status


@dataclass(frozen=True)
class PlaidConfig:
    client_id: str
    secret: str
    environment: str = "sandbox"
    client_name: str = "Personal Budget Excel"
    country_codes: tuple[str, ...] = ("US",)
    user_id: str = "personal-budget-owner"

    @property
    def base_url(self) -> str:
        env_map = {
            "sandbox": "https://sandbox.plaid.com",
            "development": "https://development.plaid.com",
            "production": "https://production.plaid.com",
        }
        if self.environment not in env_map:
            raise ValueError(f"Unsupported PLAID_ENV: {self.environment}")
        return env_map[self.environment]

    @classmethod
    def from_env(cls, credential_store: Any | None = None) -> "PlaidConfig":
        """Load explicit environment overrides, otherwise use the local keychain.

        Credentials are deliberately never written to the workbook or telemetry.
        """
        client_id = os.getenv("PLAID_CLIENT_ID", "")
        secret = os.getenv("PLAID_SECRET", "")
        environment = os.getenv("PLAID_ENV", "")
        if not (client_id and secret) and credential_store is not None:
            client_id = client_id or credential_store.get_secret("plaid-client-id")
            secret = secret or credential_store.get_secret("plaid-secret")
            environment = environment or credential_store.get_secret("plaid-environment")
        missing = [name for name, value in (("PLAID_CLIENT_ID", client_id), ("PLAID_SECRET", secret)) if not value]
        if missing:
            raise RuntimeError(f"Missing Plaid credentials: {', '.join(missing)}")
        return cls(
            client_id=client_id,
            secret=secret,
            environment=environment or "sandbox",
            client_name=os.getenv("PLAID_CLIENT_NAME", "Personal Budget Excel")[:30],
            country_codes=tuple(
                part.strip()
                for part in os.getenv("PLAID_COUNTRY_CODES", "US").split(",")
                if part.strip()
            ),
            user_id=os.getenv("PLAID_CLIENT_USER_ID", "personal-budget-owner"),
        )


Transport = Callable[[str, dict[str, Any]], dict[str, Any]]


class PlaidRestClient:
    def __init__(self, config: PlaidConfig, transport: Transport | None = None) -> None:
        self.config = config
        self._transport = transport

    def create_link_token(self) -> str:
        payload: dict[str, Any] = {
            "client_name": self.config.client_name,
            "country_codes": list(self.config.country_codes),
            "language": "en",
            "products": ["transactions"],
            "transactions": {"days_requested": 730},
            "user": {"client_user_id": self.config.user_id},
        }
        redirect_uri = os.getenv("PLAID_REDIRECT_URI")
        webhook = os.getenv("PLAID_WEBHOOK_URL")
        if redirect_uri:
            payload["redirect_uri"] = redirect_uri
        if webhook:
            payload["webhook"] = webhook
        response = self._post("/link/token/create", payload)
        return response["link_token"]

    def exchange_public_token(self, public_token: str) -> dict[str, str]:
        return self._post("/item/public_token/exchange", {"public_token": public_token})

    def sync_item(self, state: SyncItemState, access_token: str) -> PlaidSyncResult:
        original_cursor = state.cursor or None
        cursor = original_cursor
        added: list[dict[str, Any]] = []
        modified: list[dict[str, Any]] = []
        removed: list[dict[str, Any]] = []
        accounts_by_id: dict[str, dict[str, Any]] = {}
        request_ids: list[str] = []
        restart_count = 0

        while True:
            try:
                page = self._post(
                    "/transactions/sync",
                    {
                        "access_token": access_token,
                        "cursor": cursor,
                        "count": 500,
                    },
                )
            except PlaidApiError as exc:
                if exc.error_code == "TRANSACTIONS_SYNC_MUTATION_DURING_PAGINATION" and restart_count < 3:
                    restart_count += 1
                    cursor = original_cursor
                    added.clear()
                    modified.clear()
                    removed.clear()
                    accounts_by_id.clear()
                    request_ids.clear()
                    continue
                raise

            added.extend(page.get("added", []))
            modified.extend(page.get("modified", []))
            removed.extend(page.get("removed", []))
            for account in page.get("accounts", []):
                if account.get("account_id"):
                    accounts_by_id[account["account_id"]] = account
            if page.get("request_id"):
                request_ids.append(page["request_id"])
            cursor = page.get("next_cursor") or ""
            if not page.get("has_more"):
                return PlaidSyncResult(
                    item_id=state.item_id,
                    added=added,
                    modified=modified,
                    removed=removed,
                    accounts=list(accounts_by_id.values()),
                    next_cursor=cursor,
                    request_ids=request_ids,
                )

    def _post(self, path: str, payload: dict[str, Any]) -> dict[str, Any]:
        full_payload = {
            "client_id": self.config.client_id,
            "secret": self.config.secret,
            **{key: value for key, value in payload.items() if value is not None},
        }
        if self._transport is not None:
            return self._transport(path, full_payload)

        data = json.dumps(full_payload).encode("utf-8")
        request = urllib.request.Request(
            self.config.base_url + path,
            data=data,
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=45) as response:
                return json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            body = exc.read().decode("utf-8", errors="replace")
            try:
                parsed = json.loads(body)
            except json.JSONDecodeError:
                parsed = {"error_message": body}
            raise PlaidApiError(
                parsed.get("error_message") or str(exc),
                error_code=parsed.get("error_code", ""),
                status=exc.code,
            ) from exc
