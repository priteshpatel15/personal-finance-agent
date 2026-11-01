from __future__ import annotations

from collections.abc import Iterable

from .models import LedgerTransaction
from .non_spend import TreatmentDisposition, classify_budget_treatments, classify_provider_payload

CATEGORY_BRIDGES = {
    "Food & Dining": "Food & Dining & Entertainment",
    "Insurance": "Insurance (including car)",
}

EXCLUDED_CATEGORIES = {"Work Reimbursement"}

PLAID_PRIMARY_TO_BUDGET = {
    "FOOD_AND_DRINK": "Food & Dining & Entertainment",
    "GENERAL_MERCHANDISE": "Shopping",
    "TRANSPORTATION": "Transportation / Rideshare",
    "TRAVEL": "Travel / Trip",
    "MEDICAL": "Medical / Health",
    "ENTERTAINMENT": "Entertainment",
    "RENT_AND_UTILITIES": "Utilities",
    "LOAN_PAYMENTS": "Car Payment",
    "BANK_FEES": "Interest & Fees",
    "TRANSFER_OUT": "Other / Needs Review",
    "GENERAL_SERVICES": "Professional Services",
    "PERSONAL_CARE": "Personal Care",
}

PLAID_DETAILED_TO_BUDGET = {
    "FOOD_AND_DRINK_COFFEE": "Coffee",
    "FOOD_AND_DRINK_GROCERIES": "Groceries",
    "FOOD_AND_DRINK_RESTAURANT": "Food & Dining & Entertainment",
    "FOOD_AND_DRINK_FAST_FOOD": "Food & Dining & Entertainment",
    "FOOD_AND_DRINK_DELIVERY": "Food Delivery",
    "TRANSPORTATION_GAS": "Gas",
    "TRANSPORTATION_PARKING": "Parking",
    "TRANSPORTATION_TAXIS_AND_RIDE_SHARES": "Transportation / Rideshare",
    "GENERAL_MERCHANDISE_CLOTHING_AND_ACCESSORIES": "Clothes",
}


def canonical_budget_category(category: str | None) -> str | None:
    if not category:
        return None
    stripped = str(category).strip()
    return CATEGORY_BRIDGES.get(stripped, stripped)


def is_budget_excluded(category: str | None) -> bool:
    return canonical_budget_category(category) in EXCLUDED_CATEGORIES or category in EXCLUDED_CATEGORIES


def build_category_map(budget_categories: Iterable[str]) -> list[tuple[str, str, str, str]]:
    rows: list[tuple[str, str, str, str]] = []
    seen: set[tuple[str, str]] = set()
    for category in budget_categories:
        if not category:
            continue
        key = (category, category)
        if key not in seen:
            rows.append((category, category, "Yes", "Budget category"))
            seen.add(key)
    for source, target in CATEGORY_BRIDGES.items():
        key = (source, target)
        if key not in seen:
            rows.append((source, target, "Yes", "Bridge from historical workbook category"))
            seen.add(key)
    rows.append(("Work Reimbursement", "", "No", "Excluded from personal budget spend"))
    return rows


def infer_suggestion(
    plaid_transaction: dict,
    merchant_history: dict[str, str],
    budget_categories: set[str],
) -> tuple[str | None, str]:
    if classify_provider_payload(plaid_transaction).suppress_category_suggestion:
        return None, ""
    merchant = _merchant_name(plaid_transaction)
    merchant_key = merchant.lower() if merchant else ""
    if merchant_key in merchant_history:
        return merchant_history[merchant_key], "Matched prior selected category for merchant"

    pfc = plaid_transaction.get("personal_finance_category") or {}
    detailed = pfc.get("detailed")
    primary = pfc.get("primary")
    if detailed in PLAID_DETAILED_TO_BUDGET:
        suggestion = PLAID_DETAILED_TO_BUDGET[detailed]
        if suggestion in budget_categories:
            return suggestion, f"Mapped Plaid detailed category {detailed}"
    if primary in PLAID_PRIMARY_TO_BUDGET:
        suggestion = PLAID_PRIMARY_TO_BUDGET[primary]
        if suggestion in budget_categories:
            return suggestion, f"Mapped Plaid primary category {primary}"
    return None, ""


def build_merchant_history(transactions: Iterable[LedgerTransaction]) -> dict[str, str]:
    transaction_rows = list(transactions)
    treatments = classify_budget_treatments(transaction_rows)
    history: dict[str, str] = {}
    for tx, treatment in zip(transaction_rows, treatments, strict=True):
        if treatment.disposition is not TreatmentDisposition.INCLUDED:
            continue
        if not tx.merchant:
            continue
        category = canonical_budget_category(tx.selected_category or tx.saved_category)
        if category and not is_budget_excluded(category):
            history.setdefault(tx.merchant.lower(), category)
    return history


def _merchant_name(plaid_transaction: dict) -> str:
    return (
        plaid_transaction.get("merchant_name")
        or plaid_transaction.get("name")
        or plaid_transaction.get("original_description")
        or ""
    ).strip()
