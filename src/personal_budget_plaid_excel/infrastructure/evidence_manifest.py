from __future__ import annotations

import json
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

from personal_budget_plaid_excel.domain.review import EvidenceMatch, PurchaseEvidence


class JsonEvidenceManifest:
    """Read already-collected purchase evidence from a local, offline JSON file."""

    def __init__(self, path: Path) -> None:
        self.path = path

    def read(self) -> list[PurchaseEvidence]:
        payload = json.loads(self.path.read_text(encoding="utf-8"))
        records = payload.get("evidence") if isinstance(payload, dict) else payload
        if not isinstance(records, list):
            raise ValueError("evidence manifest must be a list or an object with an evidence list")
        return [self._record(record, index) for index, record in enumerate(records)]

    def _record(self, record: Any, index: int) -> PurchaseEvidence:
        if not isinstance(record, dict) or not isinstance(record.get("match"), dict):
            raise ValueError(f"evidence record {index} must contain a match object")
        match = record["match"]
        try:
            transaction_date = date.fromisoformat(str(match["date"]))
            amount = Decimal(str(match["amount"]))
            merchant = _required_text(match.get("merchant"))
        except (KeyError, ValueError, InvalidOperation) as exc:
            raise ValueError(f"evidence record {index} has invalid date, merchant, or amount") from exc

        account = _optional_text(match.get("account"))
        occurrence_value = match.get("occurrence")
        occurrence: int | None = None
        if occurrence_value is not None:
            try:
                occurrence = int(occurrence_value)
            except (TypeError, ValueError) as exc:
                raise ValueError(f"evidence record {index} occurrence must be a positive integer") from exc
            if occurrence <= 0:
                raise ValueError(f"evidence record {index} occurrence must be a positive integer")

        items_value = record.get("items", [])
        if not isinstance(items_value, list) or any(not isinstance(item, str) for item in items_value):
            raise ValueError(f"evidence record {index} items must be a list of strings")

        return PurchaseEvidence(
            match=EvidenceMatch(
                transaction_date=transaction_date,
                merchant=merchant,
                amount=amount,
                account=account,
                occurrence=occurrence,
            ),
            status=_required_text(record.get("status")),
            source=_required_text(record.get("source")),
            summary=_optional_text(record.get("summary")) or "",
            items=tuple(items_value),
            gmail_thread_id=_optional_text(record.get("gmail_thread_id")),
            review_reason=_optional_text(record.get("review_reason")) or "",
        )


def _required_text(value: Any) -> str:
    text = _optional_text(value)
    if text is None:
        raise ValueError("required text is missing")
    return text


def _optional_text(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text or None
