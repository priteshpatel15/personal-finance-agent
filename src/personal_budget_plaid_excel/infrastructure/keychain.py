from __future__ import annotations

import shutil
import subprocess


class KeychainUnavailable(RuntimeError):
    pass


class MacOSKeychain:
    def __init__(self, service: str = "personal-budget-plaid-excel") -> None:
        self.service = service
        if shutil.which("security") is None:
            raise KeychainUnavailable("macOS security command is not available")

    def set_secret(self, account: str, value: str) -> None:
        subprocess.run(
            [
                "security",
                "add-generic-password",
                "-a",
                account,
                "-s",
                self.service,
                "-w",
                value,
                "-U",
            ],
            check=True,
            capture_output=True,
            text=True,
        )

    def get_secret(self, account: str) -> str:
        result = subprocess.run(
            ["security", "find-generic-password", "-a", account, "-s", self.service, "-w"],
            check=True,
            capture_output=True,
            text=True,
        )
        return result.stdout.strip()

