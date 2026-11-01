from __future__ import annotations

import hashlib
import json
import re
from collections import Counter
from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from urllib.parse import quote

from .models import LedgerTransaction
from .non_spend import BudgetTreatment, classify_budget_treatments
from .source_mode import DirectPortalInstitution


REVIEW_NEEDED_DEFAULT = "Review needed — no verified purchase evidence"
CROSS_SOURCE_DATE_DRIFT_DAYS = 3

_TEXT_TOKEN = re.compile(r"[a-z0-9]+")
_MERCHANT_TOKEN_ALIASES = {
    "amzn": "amazon",
    "mktp": "marketplace",
    "mktplace": "marketplace",
}
_MERCHANT_NOISE_TOKENS = {
    "ach",
    "chkcard",
    "debit",
    "online",
    "pos",
    "purchase",
    "sq",
    "tst",
}
_GENERIC_MERCHANT_TOKENS = {
    "auto",
    "autopay",
    "card",
    "credit",
    "payment",
    "transaction",
}
_ACCOUNT_NOISE_TOKENS = {
    "account",
    "bank",
    "card",
    "credit",
    "debit",
    "ending",
    "express",
    "in",
    "mastercard",
    "of",
    "visa",
    "x",
    "xxxx",
}
_INSTITUTION_TOKENS = {
    "america",
    "american",
    "amex",
    "boa",
    "bofa",
    "chase",
    "jp",
    "jpmorgan",
    "morgan",
}
_GENERIC_ACCOUNT_TOKENS = {"checking", "savings"}


@dataclass(frozen=True)
class EvidenceMatch:
    transaction_date: date
    merchant: str
    amount: Decimal
    account: str | None = None
    occurrence: int | None = None


@dataclass(frozen=True)
class PurchaseEvidence:
    match: EvidenceMatch
    status: str
    source: str
    summary: str = ""
    items: tuple[str, ...] = ()
    gmail_thread_id: str | None = None
    review_reason: str = ""


@dataclass(frozen=True)
class ReviewCellDetail:
    text: str
    hyperlink: str | None = None
    supported: bool = False


@dataclass(frozen=True)
class ReviewTransaction:
    transaction: LedgerTransaction
    review_key: str
    budget_treatment: BudgetTreatment
    evidence_detail: ReviewCellDetail | None = None


@dataclass(frozen=True)
class ReviewBuildResult:
    output_path: str
    backup_path: str | None
    transaction_count: int
    category_count: int
    categorized_count: int
    uncategorized_count: int
    evidence_supported_count: int
    gmail_link_count: int
    review_needed_count: int
    excluded_non_spend_count: int
    treatment_review_needed_count: int
    source_mode: str = ""
    source_coverage_complete: bool = False
    source_expected_count: int = 0
    source_complete_count: int = 0
    source_partial_count: int = 0
    source_unavailable_count: int = 0
    reconciliation_removed_visible_count: int = 0
    reconciliation_ambiguous_count: int = 0


def build_review_transactions(
    transactions: list[LedgerTransaction],
    evidence: list[PurchaseEvidence],
) -> list[ReviewTransaction]:
    keys = make_review_keys(transactions)
    treatments = classify_budget_treatments(transactions)
    assigned: dict[int, ReviewCellDetail] = {}
    conflicts: set[int] = set()

    for record in evidence:
        candidates = [
            index
            for index, transaction in enumerate(transactions)
            if evidence_matches(record.match, transaction)
        ]
        selected: int | None = None
        if record.match.occurrence is not None:
            occurrence = record.match.occurrence
            if occurrence > 0 and occurrence <= len(candidates):
                selected = candidates[occurrence - 1]
        elif len(candidates) == 1:
            selected = candidates[0]

        if selected is None:
            continue
        if selected in assigned:
            conflicts.add(selected)
            continue
        assigned[selected] = detail_from_evidence(record)

    for index in conflicts:
        assigned[index] = review_needed_detail("conflicting evidence requires review")

    return [
        ReviewTransaction(
            transaction=transaction,
            review_key=keys[index],
            budget_treatment=treatments[index],
            evidence_detail=assigned.get(index),
        )
        for index, transaction in enumerate(transactions)
    ]


def make_review_keys(transactions: list[LedgerTransaction]) -> list[str]:
    occurrences: Counter[tuple[str, str, str, str]] = Counter()
    keys: list[str] = []
    for transaction in transactions:
        base = transaction_identity(transaction)
        occurrences[base] += 1
        payload = json.dumps([*base, occurrences[base]], ensure_ascii=False, separators=(",", ":"))
        keys.append(hashlib.sha256(payload.encode("utf-8")).hexdigest())
    return keys


def transaction_identity(transaction: LedgerTransaction) -> tuple[str, str, str, str]:
    return (
        transaction.transaction_date.isoformat(),
        normalize_match_text(transaction.merchant),
        normalize_match_text(transaction.account),
        decimal_text(transaction.amount),
    )


def evidence_matches(match: EvidenceMatch, transaction: LedgerTransaction) -> bool:
    if match.transaction_date != transaction.transaction_date:
        return False
    if normalize_match_text(match.merchant) != normalize_match_text(transaction.merchant):
        return False
    if match.amount != transaction.amount:
        return False
    if match.account is not None and normalize_match_text(match.account) != normalize_match_text(transaction.account):
        return False
    return True


def conservative_cross_source_matches(
    current: list[LedgerTransaction],
    previous: list[LedgerTransaction],
) -> dict[int, int]:
    """Return only bidirectionally unique, high-confidence cross-source matches.

    Exact review-key matches are intentionally handled by the workbook writer
    before this matcher is called.  Candidate edges require exact signed amount,
    at most three days of posting-date drift, strong merchant equivalence, and
    strong account/institution corroboration.  Any one-to-many or many-to-one
    candidate graph is omitted in full rather than guessed.
    """

    candidates_by_current: dict[int, list[int]] = {}
    current_count_by_previous: Counter[int] = Counter()
    for current_index, current_transaction in enumerate(current):
        candidates = [
            previous_index
            for previous_index, previous_transaction in enumerate(previous)
            if _is_conservative_cross_source_candidate(
                current_transaction,
                previous_transaction,
            )
        ]
        candidates_by_current[current_index] = candidates
        current_count_by_previous.update(candidates)

    return {
        current_index: candidates[0]
        for current_index, candidates in candidates_by_current.items()
        if len(candidates) == 1 and current_count_by_previous[candidates[0]] == 1
    }


def transaction_source_institution(transaction: LedgerTransaction) -> str | None:
    """Return one canonical institution only when the local evidence agrees."""

    institutions: set[str] = set()
    for key in ("source_institution", "Source Institution", "institution", "Institution"):
        raw_value = transaction.raw.get(key)
        if raw_value in (None, ""):
            continue
        try:
            institutions.add(DirectPortalInstitution.parse(str(raw_value)).value)
        except ValueError:
            return None
    institutions.update(_institutions_in_text(transaction.account))
    return next(iter(institutions)) if len(institutions) == 1 else None


def _is_conservative_cross_source_candidate(
    current: LedgerTransaction,
    previous: LedgerTransaction,
) -> bool:
    if current.amount != previous.amount:
        return False
    if abs((current.transaction_date - previous.transaction_date).days) > CROSS_SOURCE_DATE_DRIFT_DAYS:
        return False
    if not _strong_merchant_match(current.merchant, previous.merchant):
        return False
    return _strong_account_institution_match(current, previous)


def _strong_merchant_match(current: str, previous: str) -> bool:
    current_full = _tokens(current, aliases=_MERCHANT_TOKEN_ALIASES)
    previous_full = _tokens(previous, aliases=_MERCHANT_TOKEN_ALIASES)
    if current_full == previous_full and sum(map(len, current_full)) >= 3:
        return True

    current_core = tuple(
        token for token in current_full
        if token not in _MERCHANT_NOISE_TOKENS and not token.isdigit()
    )
    previous_core = tuple(
        token for token in previous_full
        if token not in _MERCHANT_NOISE_TOKENS and not token.isdigit()
    )
    current_set = set(current_core)
    previous_set = set(previous_core)
    shared = current_set & previous_set
    distinctive = shared - _GENERIC_MERCHANT_TOKENS
    if not distinctive or sum(map(len, shared)) < 6:
        return False
    if current_core == previous_core:
        return True
    if current_set <= previous_set or previous_set <= current_set:
        return True
    union = current_set | previous_set
    return bool(union) and len(shared) / len(union) >= 0.75


def _strong_account_institution_match(
    current: LedgerTransaction,
    previous: LedgerTransaction,
) -> bool:
    current_institution = transaction_source_institution(current)
    if current_institution is None:
        return False
    previous_institutions = _institutions_in_text(previous.account)
    previous_raw_institution = transaction_source_institution(previous)
    if previous_raw_institution:
        previous_institutions.add(previous_raw_institution)
    if len(previous_institutions) > 1:
        return False
    if previous_institutions and current_institution not in previous_institutions:
        return False

    current_last_four = _account_last_four(current.account)
    previous_last_four = _account_last_four(previous.account)
    if current_last_four and current_last_four == previous_last_four:
        return True

    current_tokens = set(_account_tokens(current.account))
    previous_tokens = set(_account_tokens(previous.account))
    shared = current_tokens & previous_tokens
    if not shared:
        return False
    shared_character_count = sum(map(len, shared))
    account_labels_agree = (
        current_tokens == previous_tokens
        or (
            shared_character_count >= 8
            and (current_tokens <= previous_tokens or previous_tokens <= current_tokens)
        )
        or (
            shared_character_count >= 10
            and len(shared) / len(current_tokens | previous_tokens) >= 0.75
        )
    )
    if not account_labels_agree:
        return False
    if previous_institutions:
        return shared_character_count >= 4
    distinctive = shared - _GENERIC_ACCOUNT_TOKENS
    return sum(map(len, distinctive)) >= 8


def _tokens(value: str, *, aliases: dict[str, str] | None = None) -> tuple[str, ...]:
    replacements = aliases or {}
    return tuple(
        replacements.get(token, token)
        for token in _TEXT_TOKEN.findall(str(value or "").casefold())
    )


def _account_tokens(value: str) -> tuple[str, ...]:
    return tuple(
        token
        for token in _tokens(value)
        if token not in _ACCOUNT_NOISE_TOKENS
        and token not in _INSTITUTION_TOKENS
        and not token.isdigit()
    )


def _account_last_four(value: str) -> str | None:
    groups = re.findall(r"(?<!\d)(\d{4})(?!\d)", str(value or ""))
    return groups[-1] if groups else None


def _institutions_in_text(value: str) -> set[str]:
    normalized = " ".join(_tokens(value))
    institutions: set[str] = set()
    if re.search(r"\b(?:bank of america|bofa|boa)\b", normalized):
        institutions.add(DirectPortalInstitution.BANK_OF_AMERICA.value)
    if re.search(r"\b(?:chase|jp morgan|jpmorgan)\b", normalized):
        institutions.add(DirectPortalInstitution.CHASE.value)
    if re.search(r"\b(?:american express|amex)\b", normalized):
        institutions.add(DirectPortalInstitution.AMERICAN_EXPRESS.value)
    return institutions


def detail_from_evidence(evidence: PurchaseEvidence) -> ReviewCellDetail:
    status = normalize_match_text(evidence.status)
    if status not in {"confirmed", "confident"}:
        reason = evidence.review_reason or ("ambiguous evidence" if status == "ambiguous" else "evidence not confirmed")
        return review_needed_detail(reason)

    source = normalize_match_text(evidence.source)
    summary = safe_text(evidence.summary, 300)
    items = tuple(value for value in (safe_text(item, 180) for item in evidence.items) if value)
    if not summary and not items:
        return review_needed_detail("confirmed evidence has no usable item detail")

    if source == "gmail":
        hyperlink = gmail_thread_url(evidence.gmail_thread_id)
        if hyperlink is None:
            return review_needed_detail("Gmail evidence is missing a verified thread identifier")
        lines = [f"Gmail evidence — {summary}" if summary else "Gmail evidence"]
        if items:
            lines.append("Items: " + "; ".join(items))
        lines.append("Open Gmail source")
        return ReviewCellDetail(text="\n".join(lines), hyperlink=hyperlink, supported=True)

    if source == "amazon":
        lines = [f"Amazon evidence — {summary}" if summary else "Amazon evidence"]
        if items:
            lines.append("Items: " + "; ".join(items))
        return ReviewCellDetail(text="\n".join(lines), supported=True)

    return review_needed_detail("unsupported evidence source")


def review_needed_detail(reason: str) -> ReviewCellDetail:
    cleaned = safe_text(reason, 180) or "verification is incomplete"
    return ReviewCellDetail(text=f"Review needed — {cleaned}")


def gmail_thread_url(thread_id: str | None) -> str | None:
    if not thread_id or not re.fullmatch(r"[A-Za-z0-9_-]{8,128}", thread_id):
        return None
    return f"https://mail.google.com/mail/u/0/#all/{quote(thread_id, safe='')}"


def safe_text(value: str, limit: int) -> str:
    cleaned = " ".join(str(value or "").replace("\x00", "").split())
    return cleaned[:limit]


def normalize_match_text(value: str) -> str:
    return " ".join(str(value or "").split()).casefold()


def decimal_text(value: Decimal) -> str:
    return format(value.normalize(), "f") if value else "0"
