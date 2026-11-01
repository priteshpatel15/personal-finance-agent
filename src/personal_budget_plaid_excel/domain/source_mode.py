"""Source-neutral contracts for monthly review ingestion and reconciliation.

The browser or Plaid adapters collect data.  These contracts carry only the
normalized, local facts needed by the review builder; they intentionally do not
contain credentials, provider payloads, or raw account identifiers.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, replace
from datetime import date
from decimal import Decimal
from enum import Enum


class ImportSourceMode(str, Enum):
    PLAID = "plaid"
    DIRECT_PORTAL = "direct-portal"

    @classmethod
    def parse(cls, value: str) -> "ImportSourceMode":
        normalized = str(value or "").strip().casefold().replace("_", "-")
        try:
            return cls(normalized)
        except ValueError as exc:
            raise ValueError(f"unsupported import source: {value!r}") from exc

    @property
    def display_name(self) -> str:
        return "Plaid" if self is ImportSourceMode.PLAID else "Direct portal"


class DirectPortalInstitution(str, Enum):
    BANK_OF_AMERICA = "Bank of America"
    CHASE = "Chase"
    AMERICAN_EXPRESS = "American Express"

    @classmethod
    def parse(cls, value: str) -> "DirectPortalInstitution":
        normalized = " ".join(str(value or "").strip().casefold().split())
        aliases = {
            "bank of america": cls.BANK_OF_AMERICA,
            "bofa": cls.BANK_OF_AMERICA,
            "chase": cls.CHASE,
            "american express": cls.AMERICAN_EXPRESS,
            "amex": cls.AMERICAN_EXPRESS,
        }
        if normalized not in aliases:
            raise ValueError(f"unsupported portal institution: {value!r}")
        return aliases[normalized]


DIRECT_PORTAL_INSTITUTIONS = tuple(institution.value for institution in DirectPortalInstitution)


class PortalCoverageStatus(str, Enum):
    COMPLETE = "complete"
    PARTIAL = "partial"
    UNAVAILABLE = "unavailable"

    @classmethod
    def parse(cls, value: str) -> "PortalCoverageStatus":
        try:
            return cls(str(value or "").strip().casefold())
        except ValueError as exc:
            raise ValueError(f"unsupported portal coverage status: {value!r}") from exc

    @property
    def display_name(self) -> str:
        return self.value.capitalize()


class PortalAmountConvention(str, Enum):
    """Explicit source sign convention; never infer a transaction kind from sign."""

    EXPENSE_POSITIVE = "expense-positive"
    DEBIT_NEGATIVE = "debit-negative"

    @classmethod
    def parse(cls, value: object) -> "PortalAmountConvention":
        try:
            return cls(str(value or "").strip().casefold())
        except ValueError as exc:
            raise ValueError("unsupported or missing portal amount convention") from exc


class PortalTransactionKind(str, Enum):
    UNSPECIFIED = ""
    PURCHASE = "purchase"
    CREDIT_CARD_PAYMENT = "credit-card-payment"
    ACCOUNT_TRANSFER = "account-transfer"
    REFUND = "refund"
    INCOME = "income"
    BUDGETED_LOAN_PAYMENT = "budgeted-loan-payment"
    CAR_PAYMENT = "car-payment"
    MORTGAGE_PAYMENT = "mortgage-payment"
    STUDENT_LOAN_PAYMENT = "student-loan-payment"
    AMBIGUOUS = "ambiguous"

    @classmethod
    def parse(cls, value: str | None) -> "PortalTransactionKind":
        normalized = str(value or "").strip().casefold().replace("_", "-")
        try:
            return cls(normalized)
        except ValueError as exc:
            raise ValueError(f"unsupported portal transaction kind: {value!r}") from exc


@dataclass(frozen=True)
class PortalInstitutionCoverage:
    institution: str
    status: PortalCoverageStatus
    period_start: date | None
    period_end: date | None
    posted_transaction_count: int
    reason_code: str = ""

    def __post_init__(self) -> None:
        canonical = DirectPortalInstitution.parse(self.institution).value
        object.__setattr__(self, "institution", canonical)
        if self.posted_transaction_count < 0:
            raise ValueError("posted transaction count cannot be negative")
        if self.period_start and self.period_end and self.period_start > self.period_end:
            raise ValueError("portal coverage start cannot be after end")


@dataclass(frozen=True)
class SourceCoverage:
    mode: ImportSourceMode
    month: date
    rows: tuple[PortalInstitutionCoverage, ...]

    def __post_init__(self) -> None:
        normalized_month = date(self.month.year, self.month.month, 1)
        object.__setattr__(self, "month", normalized_month)
        institutions = tuple(row.institution for row in self.rows)
        if institutions != DIRECT_PORTAL_INSTITUTIONS:
            raise ValueError("source coverage must contain Bank of America, Chase, and American Express")

    @property
    def expected_count(self) -> int:
        return len(self.rows)

    @property
    def complete_count(self) -> int:
        return sum(row.status is PortalCoverageStatus.COMPLETE for row in self.rows)

    @property
    def partial_count(self) -> int:
        return sum(row.status is PortalCoverageStatus.PARTIAL for row in self.rows)

    @property
    def unavailable_count(self) -> int:
        return sum(row.status is PortalCoverageStatus.UNAVAILABLE for row in self.rows)

    @property
    def complete(self) -> bool:
        return bool(self.rows) and self.complete_count == self.expected_count


@dataclass(frozen=True)
class ReviewSourceContext:
    mode: ImportSourceMode
    coverage: SourceCoverage

    def __post_init__(self) -> None:
        if self.coverage.mode is not self.mode:
            raise ValueError("source mode and coverage mode do not match")


@dataclass(frozen=True)
class PortalTransaction:
    local_id: str
    institution: str
    transaction_date: date
    merchant: str
    account: str
    amount: Decimal
    kind: PortalTransactionKind = PortalTransactionKind.UNSPECIFIED
    amount_convention: PortalAmountConvention = PortalAmountConvention.EXPENSE_POSITIVE

    @property
    def review_amount(self) -> Decimal:
        """Expense-positive review value while ``amount`` retains the portal sign."""
        if self.amount_convention is PortalAmountConvention.DEBIT_NEGATIVE:
            return -self.amount
        return self.amount

    def __post_init__(self) -> None:
        if not re.fullmatch(r"[0-9a-f]{64}", self.local_id):
            raise ValueError("portal local_id must be a deterministic SHA-256 identity")
        object.__setattr__(self, "institution", DirectPortalInstitution.parse(self.institution).value)
        if not str(self.merchant).strip():
            raise ValueError("portal transaction merchant is required")
        if not str(self.account).strip():
            raise ValueError("portal transaction account label is required")


class ReconciliationClassification(str, Enum):
    PLAID_REMOVED_PORTAL_VISIBLE = "Plaid marked removed — still visible at institution"
    PORTAL_ONLY_POSTED = "Portal-only posted"
    PLAID_ONLY_POSTED = "Plaid-only posted"
    PLAID_ONLY_PENDING = "Plaid-only pending"
    AMBIGUOUS = "Ambiguous date / identity match"

    @classmethod
    def parse(cls, value: str) -> "ReconciliationClassification":
        try:
            return cls(str(value or "").strip())
        except ValueError as exc:
            raise ValueError(f"unsupported reconciliation classification: {value!r}") from exc


@dataclass(frozen=True)
class ReconciliationCoverageRow:
    institution: str
    portal_status: str
    portal_posted_count: int
    active_plaid_count: int
    matched_count: int
    portal_only_count: int
    plaid_only_count: int
    ambiguous_count: int
    plaid_removed_portal_visible_count: int
    plaid_only_pending_count: int = 0

    def __post_init__(self) -> None:
        object.__setattr__(self, "institution", DirectPortalInstitution.parse(self.institution).value)
        counts = (
            self.portal_posted_count,
            self.active_plaid_count,
            self.matched_count,
            self.portal_only_count,
            self.plaid_only_count,
            self.ambiguous_count,
            self.plaid_removed_portal_visible_count,
            self.plaid_only_pending_count,
        )
        if any(value < 0 for value in counts):
            raise ValueError("reconciliation counts cannot be negative")


@dataclass(frozen=True)
class ReconciliationItem:
    local_id: str
    institution: str
    transaction_date: date | None
    amount: Decimal | None
    classification: ReconciliationClassification
    local_status: str
    budget_effect: str
    review_note: str

    def __post_init__(self) -> None:
        if not re.fullmatch(r"[0-9a-f]{64}", self.local_id):
            raise ValueError("reconciliation local_id must be a deterministic SHA-256 identity")
        object.__setattr__(self, "institution", DirectPortalInstitution.parse(self.institution).value)


@dataclass(frozen=True)
class ReconciliationReport:
    coverage: tuple[ReconciliationCoverageRow, ...]
    items: tuple[ReconciliationItem, ...]

    def __post_init__(self) -> None:
        institutions = [row.institution for row in self.coverage]
        if len(institutions) != len(set(institutions)):
            raise ValueError("reconciliation coverage institutions must be unique")
        unknown = set(institutions) - set(DIRECT_PORTAL_INSTITUTIONS)
        if unknown:
            raise ValueError("reconciliation coverage contains an unsupported institution")
        item_ids = [item.local_id for item in self.items]
        if len(item_ids) != len(set(item_ids)):
            raise ValueError("reconciliation item identities must be unique")

    @property
    def normalized_coverage(self) -> tuple[ReconciliationCoverageRow, ...]:
        by_institution = {row.institution: row for row in self.coverage}
        return tuple(
            by_institution.get(
                institution,
                ReconciliationCoverageRow(
                    institution=institution,
                    portal_status="Not reconciled",
                    portal_posted_count=0,
                    active_plaid_count=0,
                    matched_count=0,
                    portal_only_count=0,
                    plaid_only_count=0,
                    ambiguous_count=0,
                    plaid_removed_portal_visible_count=0,
                    plaid_only_pending_count=0,
                ),
            )
            for institution in DIRECT_PORTAL_INSTITUTIONS
        )

    @property
    def removed_visible_count(self) -> int:
        return sum(
            item.classification is ReconciliationClassification.PLAID_REMOVED_PORTAL_VISIBLE
            for item in self.items
        )

    @property
    def ambiguous_count(self) -> int:
        return sum(item.classification is ReconciliationClassification.AMBIGUOUS for item in self.items)


@dataclass(frozen=True)
class PortalImportBatch:
    month: date
    transactions: tuple[PortalTransaction, ...]
    coverage: SourceCoverage
    reconciliation: ReconciliationReport | None = None

    def __post_init__(self) -> None:
        normalized_month = date(self.month.year, self.month.month, 1)
        object.__setattr__(self, "month", normalized_month)
        if self.coverage.mode is not ImportSourceMode.DIRECT_PORTAL:
            raise ValueError("portal import batch requires direct portal coverage")
        if self.coverage.month != normalized_month:
            raise ValueError("portal import month and coverage month do not match")

    def without_reconciliation(self) -> "PortalImportBatch":
        return replace(self, reconciliation=None)


# Neutral name for callers constructing either Plaid or Direct portal metadata.
InstitutionCoverage = PortalInstitutionCoverage
