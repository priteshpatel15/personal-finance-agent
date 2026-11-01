from __future__ import annotations

from typing import Protocol


class SecretStore(Protocol):
    def set_secret(self, account: str, value: str) -> None:
        """Store a secret value under an account key."""

    def get_secret(self, account: str) -> str:
        """Load a secret value by account key."""

