"""Conservative, auditable classification of non-spend account activity."""

from __future__ import annotations

import re
from collections import defaultdict
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, InvalidOperation
from enum import Enum
from typing import Any, Iterable, Mapping

from .models import LedgerTransaction


class TreatmentDisposition(str, Enum):
    INCLUDED = "Included in Budget"
    EXCLUDED = "Excluded from Budget"
    REVIEW_NEEDED = "Review needed"


class TreatmentKind(str, Enum):
    PURCHASE = "Purchase"
    CREDIT_CARD_PAYMENT = "Credit-card payment"
    INTERNAL_TRANSFER = "Account transfer"
    REFUND_CREDIT = "Refund/credit offset"
    INCOME_CREDIT = "Income/account credit"
    BUDGETED_LOAN_PAYMENT = "Budgeted loan payment"
    UNKNOWN = "Possible payment/transfer"


@dataclass(frozen=True)
class BudgetTreatment:
    disposition: TreatmentDisposition
    kind: TreatmentKind
    reason_code: str

    @property
    def excluded_from_budget(self) -> bool:
        return self.disposition is TreatmentDisposition.EXCLUDED

    @property
    def suppress_category_suggestion(self) -> bool:
        return self.disposition is not TreatmentDisposition.INCLUDED

    @property
    def visible_basis(self) -> str:
        if self.reason_code == "ordinary_spend":
            return ""
        if self.reason_code == "negative_direction_unverified":
            return "Negative credit requires review"
        if self.disposition is TreatmentDisposition.REVIEW_NEEDED:
            return "Possible payment/transfer — verify"
        return self.kind.value


_STRONG_CONFIDENCE = {"HIGH", "VERY_HIGH"}
_CREDIT_CARD_PAYMENT_DETAILS = {"LOAN_PAYMENTS_CREDIT_CARD_PAYMENT"}
_OTHER_PAYMENT_DETAILS = {"LOAN_PAYMENTS_OTHER_PAYMENT"}
_BUDGETED_LOAN_DETAILS = {
    "LOAN_PAYMENTS_CAR_PAYMENT",
    "LOAN_PAYMENTS_MORTGAGE_PAYMENT",
    "LOAN_PAYMENTS_STUDENT_LOAN_PAYMENT",
}
_ACCOUNT_TRANSFER_DETAILS = {
    "TRANSFER_IN_ACCOUNT_TRANSFER",
    "TRANSFER_OUT_ACCOUNT_TRANSFER",
    "TRANSFER_IN_SAVINGS",
    "TRANSFER_OUT_SAVINGS",
    "TRANSFER_IN_INVESTMENT_AND_RETIREMENT_FUNDS",
    "TRANSFER_OUT_INVESTMENT_AND_RETIREMENT_FUNDS",
}
_PURCHASE_PRIMARIES = {
    "BANK_FEES",
    "ENTERTAINMENT",
    "FOOD_AND_DRINK",
    "GENERAL_MERCHANDISE",
    "GENERAL_SERVICES",
    "GOVERNMENT_AND_NON_PROFIT",
    "HOME_IMPROVEMENT",
    "MEDICAL",
    "PERSONAL_CARE",
    "RENT_AND_UTILITIES",
    "TRANSPORTATION",
    "TRAVEL",
}
_PAYMENT_TOKEN = re.compile(r"\b(?:auto\s*pay(?:ment)?|e\s*payment|payment|pmt|pymt)\b", re.IGNORECASE)
_TRANSFER_TOKEN = re.compile(r"\b(?:transfer|xfer|tfr)\b", re.IGNORECASE)
_REFUND_TOKEN = re.compile(r"\b(?:refund|returned?|rebate|cash\s*back|cashback|reversal)\b", re.IGNORECASE)
_CLEAR_PAYMENT = re.compile(
    r"^(?:(?:ach|online|mobile|electronic|automatic|scheduled|web)\s+)?"
    r"(?:(?:credit\s+card|card)\s+)?(?:payment|pmt|pymt)"
    r"(?:\s+(?:thank\s+you|received|confirmation))?(?:\s+[a-z0-9]+)?$",
    re.IGNORECASE,
)
_CLEAR_AUTOPAY = re.compile(r"^auto\s*pay(?:ment)?(?:\s+(?:thank\s+you|received))?$", re.IGNORECASE)
_CLEAR_TRANSFER = re.compile(
    r"^(?:(?:ach|online|mobile|internal|account|bank)\s+)?(?:transfer|xfer|tfr)"
    r"(?:\s+(?:to|from)\s+[a-z0-9 ]+)?$",
    re.IGNORECASE,
)


def classify_budget_treatments(transactions: Iterable[LedgerTransaction]) -> list[BudgetTreatment]:
    """Classify all rows together so account pairs and repeated patterns are usable evidence."""

    rows = list(transactions)
    treatments = [_base_treatment(transaction) for transaction in rows]
    _apply_same_account_payment_clusters(rows, treatments)
    _apply_opposing_account_pairs(rows, treatments)
    _apply_confirmed_history(rows, treatments)
    return treatments


def classify_provider_payload(payload: Mapping[str, Any]) -> BudgetTreatment:
    """Classify a single provider payload before category suggestion runs."""

    raw_date = str(payload.get("date") or "2000-01-01")
    try:
        transaction_date = date.fromisoformat(raw_date)
    except ValueError:
        transaction_date = date(2000, 1, 1)
    try:
        amount = Decimal(str(payload.get("amount") or 0))
    except (InvalidOperation, ValueError):
        amount = Decimal("0")
    transaction = LedgerTransaction(
        transaction_date=transaction_date,
        month=date(transaction_date.year, transaction_date.month, 1),
        merchant=str(
            payload.get("merchant_name")
            or payload.get("name")
            or payload.get("original_description")
            or ""
        ),
        account=str(payload.get("account_name") or ""),
        amount=amount,
        personal_budget=True,
        raw=dict(payload),
    )
    return classify_budget_treatments([transaction])[0]


def _base_treatment(transaction: LedgerTransaction) -> BudgetTreatment:
    primary, detailed, confidence = _pfc(transaction)
    description = _description(transaction)
    amount = transaction.amount

    if primary == "INCOME":
        return BudgetTreatment(
            TreatmentDisposition.EXCLUDED,
            TreatmentKind.INCOME_CREDIT,
            "pfc_income_credit",
        )

    if detailed in _CREDIT_CARD_PAYMENT_DETAILS:
        if confidence in _STRONG_CONFIDENCE or _has_clear_payment(description):
            return BudgetTreatment(
                TreatmentDisposition.EXCLUDED,
                TreatmentKind.CREDIT_CARD_PAYMENT,
                "pfc_credit_card_payment",
            )
        return _review("possible_credit_card_payment")

    if detailed in _ACCOUNT_TRANSFER_DETAILS:
        if confidence in _STRONG_CONFIDENCE or _has_clear_transfer(description):
            return BudgetTreatment(
                TreatmentDisposition.EXCLUDED,
                TreatmentKind.INTERNAL_TRANSFER,
                "pfc_account_transfer",
            )
        return _review("possible_account_transfer")

    if detailed in _BUDGETED_LOAN_DETAILS:
        return BudgetTreatment(
            TreatmentDisposition.INCLUDED,
            TreatmentKind.BUDGETED_LOAN_PAYMENT,
            "budgeted_loan_payment",
        )

    if amount < 0 and _REFUND_TOKEN.search(description):
        return BudgetTreatment(
            TreatmentDisposition.INCLUDED,
            TreatmentKind.REFUND_CREDIT,
            "refund_credit_offset",
        )

    clear_payment = _has_clear_payment(description)
    clear_transfer = _has_clear_transfer(description)
    purchase_signal = primary in _PURCHASE_PRIMARIES

    if clear_payment:
        if primary == "LOAN_PAYMENTS" and confidence in _STRONG_CONFIDENCE:
            return BudgetTreatment(
                TreatmentDisposition.EXCLUDED,
                TreatmentKind.CREDIT_CARD_PAYMENT,
                "supported_payment_description",
            )
        return _review("possible_payment_description")

    if clear_transfer:
        return _review("possible_transfer_description")

    if detailed in _OTHER_PAYMENT_DETAILS or primary == "LOAN_PAYMENTS":
        return _review("generic_loan_payment")

    if primary in {"TRANSFER_IN", "TRANSFER_OUT"}:
        return _review("generic_transfer")

    if amount < 0 and (
        purchase_signal
        or transaction.selected_category
        or transaction.saved_category
        or _raw_yes(transaction, "Refund?")
    ):
        return BudgetTreatment(
            TreatmentDisposition.INCLUDED,
            TreatmentKind.REFUND_CREDIT,
            "categorized_refund_credit_offset",
        )

    if purchase_signal:
        return BudgetTreatment(
            TreatmentDisposition.INCLUDED,
            TreatmentKind.PURCHASE,
            "confirmed_purchase",
        )

    if _PAYMENT_TOKEN.search(description) or _TRANSFER_TOKEN.search(description):
        return _review("generic_payment_transfer_text")

    if amount < 0:
        return BudgetTreatment(
            TreatmentDisposition.REVIEW_NEEDED,
            TreatmentKind.UNKNOWN,
            "negative_direction_unverified",
        )

    return BudgetTreatment(
        TreatmentDisposition.INCLUDED,
        TreatmentKind.PURCHASE,
        "ordinary_spend",
    )


def _apply_same_account_payment_clusters(
    transactions: list[LedgerTransaction],
    treatments: list[BudgetTreatment],
) -> None:
    clusters: dict[tuple[date, str, int, str], list[int]] = defaultdict(list)
    for index, transaction in enumerate(transactions):
        if transaction.is_removed or transaction.amount == 0:
            continue
        primary, detailed, _ = _pfc(transaction)
        description = _description(transaction)
        if primary != "LOAN_PAYMENTS":
            continue
        if detailed not in (_CREDIT_CARD_PAYMENT_DETAILS | _OTHER_PAYMENT_DETAILS):
            continue
        if not (_PAYMENT_TOKEN.search(description) or detailed in _CREDIT_CARD_PAYMENT_DETAILS):
            continue
        signature = _same_account_payment_signature(description)
        if not signature:
            continue
        clusters[
            (
                transaction.transaction_date,
                _normalize(transaction.account),
                _amount_direction(transaction.amount),
                signature,
            )
        ].append(index)

    for indexes in clusters.values():
        if len(indexes) < 3:
            continue
        strong_count = sum(
            _pfc(transactions[index])[1] in _CREDIT_CARD_PAYMENT_DETAILS
            and _pfc(transactions[index])[2] in _STRONG_CONFIDENCE
            for index in indexes
        )
        if strong_count < 2:
            continue
        for index in indexes:
            if treatments[index].disposition is not TreatmentDisposition.REVIEW_NEEDED:
                continue
            if _has_manual_spend_evidence(transactions[index]):
                continue
            treatments[index] = BudgetTreatment(
                TreatmentDisposition.EXCLUDED,
                TreatmentKind.CREDIT_CARD_PAYMENT,
                "same_account_payment_cluster",
            )


def _apply_opposing_account_pairs(
    transactions: list[LedgerTransaction],
    treatments: list[BudgetTreatment],
) -> None:
    candidates: dict[int, set[int]] = defaultdict(set)
    candidate_kinds: dict[tuple[int, int], TreatmentKind] = {}
    for left_index, left in enumerate(transactions):
        if left.is_removed:
            continue
        for right_index in range(left_index + 1, len(transactions)):
            right = transactions[right_index]
            if right.is_removed:
                continue
            if _normalize(left.account) == _normalize(right.account):
                continue
            if left.amount == 0 or right.amount == 0 or left.amount != -right.amount:
                continue
            if abs((left.transaction_date - right.transaction_date).days) > 3:
                continue
            if _opposing_pair_contradiction(treatments[left_index]) or _opposing_pair_contradiction(
                treatments[right_index]
            ):
                continue
            if _has_manual_spend_evidence(left) or _has_manual_spend_evidence(right):
                continue
            if not (_has_payment_transfer_signal(left) and _has_payment_transfer_signal(right)):
                continue
            kind = (
                TreatmentKind.CREDIT_CARD_PAYMENT
                if _is_credit_card_payment_signal(left) or _is_credit_card_payment_signal(right)
                else TreatmentKind.INTERNAL_TRANSFER
            )
            candidates[left_index].add(right_index)
            candidates[right_index].add(left_index)
            candidate_kinds[(left_index, right_index)] = kind

    for (left_index, right_index), kind in candidate_kinds.items():
        if candidates[left_index] != {right_index} or candidates[right_index] != {left_index}:
            continue
        for index in (left_index, right_index):
            treatments[index] = BudgetTreatment(
                TreatmentDisposition.EXCLUDED,
                kind,
                "opposing_account_pair",
            )


def _opposing_pair_contradiction(treatment: BudgetTreatment) -> bool:
    return treatment.kind in {
        TreatmentKind.PURCHASE,
        TreatmentKind.REFUND_CREDIT,
        TreatmentKind.INCOME_CREDIT,
        TreatmentKind.BUDGETED_LOAN_PAYMENT,
    }


def _apply_confirmed_history(
    transactions: list[LedgerTransaction],
    treatments: list[BudgetTreatment],
) -> None:
    groups: dict[tuple[str, str, int], list[int]] = defaultdict(list)
    for index, transaction in enumerate(transactions):
        if transaction.is_removed or transaction.amount == 0:
            continue
        signature = _normalize(_description(transaction))
        if signature:
            groups[
                (
                    _normalize(transaction.account),
                    signature,
                    _amount_direction(transaction.amount),
                )
            ].append(index)

    for indexes in groups.values():
        confirmed = [
            index
            for index in indexes
            if treatments[index].disposition is TreatmentDisposition.EXCLUDED
            and treatments[index].kind in {TreatmentKind.CREDIT_CARD_PAYMENT, TreatmentKind.INTERNAL_TRANSFER}
        ]
        conflicting = [
            index
            for index in indexes
            if treatments[index].disposition is TreatmentDisposition.INCLUDED
        ]
        confirmed_kinds = {treatments[index].kind for index in confirmed}
        if len(confirmed) < 2 or len(confirmed_kinds) != 1 or conflicting:
            continue
        for index in indexes:
            if treatments[index].disposition is not TreatmentDisposition.REVIEW_NEEDED:
                continue
            if _has_manual_spend_evidence(transactions[index]):
                continue
            treatments[index] = BudgetTreatment(
                TreatmentDisposition.EXCLUDED,
                treatments[confirmed[0]].kind,
                "confirmed_historical_pattern",
            )


def _review(reason_code: str) -> BudgetTreatment:
    return BudgetTreatment(
        TreatmentDisposition.REVIEW_NEEDED,
        TreatmentKind.UNKNOWN,
        reason_code,
    )


def _pfc(transaction: LedgerTransaction) -> tuple[str, str, str]:
    nested = transaction.raw.get("personal_finance_category") or {}
    primary = _upper(
        transaction.raw.get("Plaid PFC Primary")
        or (nested.get("primary") if isinstance(nested, Mapping) else "")
    )
    detailed = _upper(
        transaction.raw.get("Plaid PFC Detailed")
        or (nested.get("detailed") if isinstance(nested, Mapping) else "")
    )
    confidence = _upper(
        transaction.raw.get("Plaid PFC Confidence")
        or (nested.get("confidence_level") if isinstance(nested, Mapping) else "")
    )
    return primary, detailed, confidence


def _description(transaction: LedgerTransaction) -> str:
    return _normalize(
        str(
            transaction.raw.get("Plaid Original Description")
            or transaction.raw.get("original_description")
            or transaction.merchant
            or ""
        )
    )


def _has_clear_payment(description: str) -> bool:
    return bool(_CLEAR_PAYMENT.fullmatch(description) or _CLEAR_AUTOPAY.fullmatch(description))


def _has_clear_transfer(description: str) -> bool:
    return bool(_CLEAR_TRANSFER.fullmatch(description))


def _has_payment_transfer_signal(transaction: LedgerTransaction) -> bool:
    primary, detailed, _ = _pfc(transaction)
    description = _description(transaction)
    return bool(
        primary in {"LOAN_PAYMENTS", "TRANSFER_IN", "TRANSFER_OUT"}
        or detailed in (_CREDIT_CARD_PAYMENT_DETAILS | _OTHER_PAYMENT_DETAILS | _ACCOUNT_TRANSFER_DETAILS)
        or _PAYMENT_TOKEN.search(description)
        or _TRANSFER_TOKEN.search(description)
    )


def _is_credit_card_payment_signal(transaction: LedgerTransaction) -> bool:
    primary, detailed, _ = _pfc(transaction)
    return detailed in _CREDIT_CARD_PAYMENT_DETAILS or (
        primary == "LOAN_PAYMENTS" and _PAYMENT_TOKEN.search(_description(transaction)) is not None
    )


def _same_account_payment_signature(description: str) -> str:
    """Normalize volatile numeric suffixes without merging distinct payment descriptions."""

    return " ".join(token for token in _normalize(description).split() if not token.isdigit())


def _amount_direction(amount: Decimal) -> int:
    if amount > 0:
        return 1
    if amount < 0:
        return -1
    return 0


def _has_manual_spend_evidence(transaction: LedgerTransaction) -> bool:
    return bool(
        transaction.selected_category
        or transaction.saved_category
        or _raw_yes(transaction, "Refund?")
        or _REFUND_TOKEN.search(_description(transaction))
    )


def _raw_yes(transaction: LedgerTransaction, key: str) -> bool:
    return str(transaction.raw.get(key) or "").strip().casefold() == "yes"


def _normalize(value: object) -> str:
    return " ".join(re.sub(r"[^a-z0-9]+", " ", str(value or "").casefold()).split())


def _upper(value: object) -> str:
    return str(value or "").strip().upper()
