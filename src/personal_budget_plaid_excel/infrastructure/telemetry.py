from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

SENSITIVE_KEY_FRAGMENTS = (
    "token",
    "secret",
    "password",
    "authorization",
    "account_id",
    "item_id",
    "transaction_id",
    "cursor",
    "request_id",
    "merchant",
    "amount",
    "payload",
)
SENSITIVE_EXACT_KEYS = {
    "accounts",
    "error",
    "exception",
    "message",
    "raw",
    "transactions",
}
SAFE_DIAGNOSTIC_KEYS = {"error_code", "error_type", "status"}


class Telemetry:
    def __init__(self, path: Path | None = None) -> None:
        self.path = path

    def event(self, name: str, **payload: Any) -> None:
        if self.path is None:
            return
        self.path.parent.mkdir(parents=True, exist_ok=True)
        row = {
            "timestamp": datetime.now(timezone.utc).isoformat(),
            "event": name,
            "payload": redact(payload),
        }
        with self.path.open("a", encoding="utf-8") as handle:
            handle.write(json.dumps(row, sort_keys=True) + "\n")


def redact(value: Any) -> Any:
    if isinstance(value, dict):
        return {
            key: "[REDACTED]" if _is_secret_key(key) else redact(item)
            for key, item in value.items()
        }
    if isinstance(value, list):
        return [redact(item) for item in value]
    return value


def _is_secret_key(key: str) -> bool:
    lowered = key.lower()
    if lowered in SAFE_DIAGNOSTIC_KEYS:
        return False
    return lowered in SENSITIVE_EXACT_KEYS or any(
        marker in lowered for marker in SENSITIVE_KEY_FRAGMENTS
    )
