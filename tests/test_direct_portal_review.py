from __future__ import annotations

import hashlib
import json
import tempfile
import unittest
from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from pathlib import Path

from openpyxl import load_workbook

from personal_budget_plaid_excel.application.monthly_review import BuildMonthlyReview
from personal_budget_plaid_excel.domain.models import CategoryBudget, LedgerTransaction
from personal_budget_plaid_excel.domain.review import REVIEW_NEEDED_DEFAULT
from personal_budget_plaid_excel.domain.source_mode import (
    DIRECT_PORTAL_INSTITUTIONS,
    ImportSourceMode,
    PortalCoverageStatus,
    PortalInstitutionCoverage,
    ReconciliationClassification,
    ReviewSourceContext,
    SourceCoverage,
)
from personal_budget_plaid_excel.infrastructure.portal_manifest import JsonPortalManifest
from personal_budget_plaid_excel.infrastructure.portal_source import DirectPortalBudgetSource
from personal_budget_plaid_excel.infrastructure.evidence_manifest import JsonEvidenceManifest
from personal_budget_plaid_excel.infrastructure.review_workbook import OpenpyxlReviewWorkbookWriter
from personal_budget_plaid_excel.infrastructure.telemetry import Telemetry


@dataclass
class _Authority:
    budgets: list[CategoryBudget]

    def read_budgets(self) -> list[CategoryBudget]:
        return list(self.budgets)

    def read_transactions(self) -> list[LedgerTransaction]:
        raise AssertionError("Direct portal mode must not co-mingle authority or Plaid transactions")


@dataclass
class _BudgetSource:
    budgets: list[CategoryBudget]
    transactions: list[LedgerTransaction]

    def read_budgets(self) -> list[CategoryBudget]:
        return list(self.budgets)

    def read_transactions(self) -> list[LedgerTransaction]:
        return list(self.transactions)


class DirectPortalReviewTests(unittest.TestCase):
    def test_source_mode_parse_has_stable_canonical_values(self) -> None:
        self.assertIs(ImportSourceMode.parse("plaid"), ImportSourceMode.PLAID)
        self.assertIs(ImportSourceMode.parse(" DIRECT_PORTAL "), ImportSourceMode.DIRECT_PORTAL)
        self.assertEqual(ImportSourceMode.DIRECT_PORTAL.value, "direct-portal")
        with self.assertRaisesRegex(ValueError, "unsupported import source"):
            ImportSourceMode.parse("screen-scrape")

    def test_manifest_normalizes_missing_coverage_and_deduplicates_local_identity(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "portal.json"
            duplicate_id = _local_id("boa-one")
            path.write_text(json.dumps({
                "schema_version": 1,
                "source": "direct-portal",
                "month": "2026-08",
                "coverage": [
                    _coverage("Bank of America", "complete", 1),
                    _coverage("Chase", "partial", 1, period_end="2026-08-20"),
                ],
                "transactions": [
                    _portal_tx(duplicate_id, "Bank of America", "2026-08-04", "Synthetic Market", "Checking", "20.00"),
                    _portal_tx(duplicate_id, "Bank of America", "2026-08-04", "Synthetic Market", "Checking", "20.00"),
                    _portal_tx(_local_id("chase-one"), "Chase", "2026-08-03", "Synthetic Shop", "Card", "11.00"),
                ],
            }), encoding="utf-8")

            batch = JsonPortalManifest(path).read()

            self.assertEqual(batch.month, date(2026, 8, 1))
            self.assertEqual(tuple(row.institution for row in batch.coverage.rows), DIRECT_PORTAL_INSTITUTIONS)
            self.assertFalse(batch.coverage.complete)
            self.assertEqual(batch.coverage.complete_count, 1)
            self.assertEqual(batch.coverage.partial_count, 1)
            self.assertEqual(batch.coverage.unavailable_count, 1)
            self.assertEqual(batch.coverage.rows[-1].status, PortalCoverageStatus.UNAVAILABLE)
            self.assertEqual(len(batch.transactions), 2)
            self.assertEqual([row.transaction_date for row in batch.transactions], sorted(row.transaction_date for row in batch.transactions))

    def test_manifest_rejects_unknown_institution_and_conflicting_identity(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            unknown = root / "unknown.json"
            unknown.write_text(json.dumps({
                "schema_version": 1,
                "source": "direct-portal",
                "month": "2026-08",
                "coverage": [_coverage("Unknown Bank", "complete", 0)],
                "transactions": [],
            }), encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "unsupported portal institution"):
                JsonPortalManifest(unknown).read()

            conflict = root / "conflict.json"
            shared_id = _local_id("shared")
            conflict.write_text(json.dumps({
                "schema_version": 1,
                "source": "direct-portal",
                "month": "2026-08",
                "coverage": [_coverage(name, "complete", 0) for name in DIRECT_PORTAL_INSTITUTIONS],
                "transactions": [
                    _portal_tx(shared_id, "Chase", "2026-08-01", "Merchant One", "Card", "1.00"),
                    _portal_tx(shared_id, "Chase", "2026-08-02", "Merchant Two", "Card", "2.00"),
                ],
            }), encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "conflicting local transaction identity"):
                JsonPortalManifest(conflict).read()

    def test_manifest_computes_stable_ids_and_uses_occurrence_for_identical_postings(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            first_path = root / "first.json"
            second_path = root / "second.json"
            base = _portal_tx(
                _local_id("discarded"),
                "Bank of America",
                "2026-08-07",
                "Synthetic Transit",
                "Checking",
                "25.00",
            )
            base.pop("local_id")
            repeated_capture = dict(base)
            distinct_identical_posting = {**base, "occurrence": 2}
            coverage = [
                _coverage("Bank of America", "complete", 2),
                _coverage("Chase", "complete", 0),
                _coverage("American Express", "complete", 0),
            ]
            common = {
                "schema_version": 1,
                "source": "direct-portal",
                "month": "2026-08",
                "coverage": coverage,
            }
            first_path.write_text(json.dumps({
                **common,
                "transactions": [base, repeated_capture, distinct_identical_posting],
            }), encoding="utf-8")
            second_path.write_text(json.dumps({
                **common,
                "transactions": [distinct_identical_posting, repeated_capture, base],
            }), encoding="utf-8")

            first = JsonPortalManifest(first_path).read()
            second = JsonPortalManifest(second_path).read()

            first_ids = [transaction.local_id for transaction in first.transactions]
            self.assertEqual(len(first_ids), 2)
            self.assertEqual(first_ids, [transaction.local_id for transaction in second.transactions])
            self.assertTrue(all(len(local_id) == 64 for local_id in first_ids))

            invalid_path = root / "invalid-occurrence.json"
            invalid_path.write_text(json.dumps({
                **common,
                "transactions": [{**base, "occurrence": 0}],
            }), encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "occurrence must be a positive integer"):
                JsonPortalManifest(invalid_path).read()

    def test_direct_portal_build_is_source_exclusive_idempotent_and_displays_coverage(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            manifest = root / "portal.json"
            output = root / "August-2026-Personal-Budget-Review.xlsx"
            manifest.write_text(json.dumps({
                "schema_version": 1,
                "source": "direct-portal",
                "month": "2026-08",
                "coverage": [
                    _coverage("Bank of America", "complete", 1),
                    _coverage("Chase", "complete", 1),
                    _coverage("American Express", "complete", 1),
                ],
                "transactions": [
                    _portal_tx(_local_id("purchase"), "Bank of America", "2026-08-02", "Synthetic Market", "Checking", "40.00", kind="purchase"),
                    _portal_tx(_local_id("payment"), "Chase", "2026-08-03", "AUTOPAY PAYMENT", "Card", "500.00", kind="credit-card-payment"),
                    _portal_tx(_local_id("refund"), "American Express", "2026-08-04", "Synthetic Refund", "Card", "-5.00", kind="refund"),
                ],
            }), encoding="utf-8")
            authority = _Authority([
                CategoryBudget("Flexible", "Groceries", Decimal("200")),
                CategoryBudget("Flexible", "Shopping", Decimal("100")),
            ])
            batch = JsonPortalManifest(manifest).read()
            source = DirectPortalBudgetSource(authority, batch)
            service = BuildMonthlyReview(OpenpyxlReviewWorkbookWriter())

            first = service.execute(source, output, date(2026, 8, 1), telemetry=Telemetry(root / "events.jsonl"))
            workbook = load_workbook(output)
            transactions = workbook["Transactions"]
            merchant_rows = {transactions.cell(row, 2).value: row for row in range(4, transactions.max_row + 1)}
            transactions.cell(merchant_rows["Synthetic Market"], 5).value = "Groceries"
            transactions.cell(merchant_rows["Synthetic Market"], 6).value = "Manual local detail"
            workbook.save(output)

            second = service.execute(source, output, date(2026, 8, 1), telemetry=Telemetry(root / "events.jsonl"))

            self.assertEqual(first.transaction_count, 3)
            self.assertEqual(second.transaction_count, 3)
            self.assertEqual(second.source_mode, "direct-portal")
            self.assertTrue(second.source_coverage_complete)
            self.assertEqual(second.source_complete_count, 3)
            self.assertEqual(second.source_unavailable_count, 0)
            final = load_workbook(output, data_only=False)
            transactions = final["Transactions"]
            summary = final["Budget Summary"]
            rows = {transactions.cell(row, 2).value: row for row in range(4, transactions.max_row + 1)}
            self.assertEqual(len(rows), 3)
            self.assertEqual(transactions.cell(rows["Synthetic Market"], 5).value, "Groceries")
            self.assertEqual(transactions.cell(rows["Synthetic Market"], 6).value, "Manual local detail")
            self.assertEqual(transactions.cell(rows["AUTOPAY PAYMENT"], 7).value, "Excluded from Budget")
            self.assertEqual(transactions.cell(rows["Synthetic Refund"], 7).value, "Included in Budget")
            self.assertEqual(summary["G1"].value, "Import Source")
            self.assertEqual(summary["H1"].value, "Direct portal")
            self.assertEqual(summary["H2"].value, "Complete")
            self.assertEqual([summary.cell(row, 7).value for row in range(4, 7)], list(DIRECT_PORTAL_INSTITUTIONS))
            self.assertEqual([summary.cell(row, 8).value for row in range(4, 7)], ["Complete"] * 3)
            self.assertEqual(final["Transactions"].tables["AugustTransactions"].ref, "A3:I6")
            telemetry = (root / "events.jsonl").read_text(encoding="utf-8")
            self.assertIn('"source_mode": "direct-portal"', telemetry)
            for forbidden in ("Synthetic Market", "AUTOPAY PAYMENT", "500.00", _local_id("purchase")):
                self.assertNotIn(forbidden, telemetry)

    def test_unique_plaid_to_portal_bridge_preserves_every_user_owned_review_field(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            output = root / "review.xlsx"
            budgets = [
                CategoryBudget("Flexible", "Dining", Decimal("250")),
                CategoryBudget("Flexible", "Groceries", Decimal("400")),
            ]
            prior_source = _BudgetSource(
                budgets,
                [
                    _ledger_tx(
                        "2026-08-10",
                        "SQ * SYNTHETIC CAFE 042",
                        "Bank of America Synthetic Cash Rewards Visa",
                        "21.37",
                    ),
                    _ledger_tx(
                        "2026-08-15",
                        "TST* SYNTHETIC MARKET #0091",
                        "American Express Synthetic Rewards Preferred Card",
                        "63.25",
                    ),
                ],
            )
            service = BuildMonthlyReview(OpenpyxlReviewWorkbookWriter())
            service.execute(prior_source, output, date(2026, 8, 1))

            prior = load_workbook(output, data_only=False)
            self.assertIsNone(prior["Budget Summary"]["H1"].value)
            self.assertIsNone(prior["Budget Summary"]["H2"].value)
            sheet = prior["Transactions"]
            formula_link = (
                '=HYPERLINK("https://mail.google.com/mail/u/0/#all/syntheticFormula123",'
                '"Existing Gmail evidence"&CHAR(10)&"Open Gmail source")'
            )
            native_link = "https://mail.google.com/mail/u/0/#all/syntheticNative123"
            sheet["E4"] = "Dining"
            sheet["F4"] = formula_link
            sheet["G4"] = "Excluded from Budget"
            sheet["H4"] = "Existing confirmed payment treatment"
            sheet.row_dimensions[4].height = 61.5
            sheet["E5"] = "Groceries"
            sheet["F5"] = "Existing native Gmail evidence"
            sheet["F5"].hyperlink = native_link
            sheet["G5"] = "Review needed"
            sheet["H5"] = "Existing treatment review"
            sheet.row_dimensions[5].height = 73.0
            prior.save(output)

            manifest = root / "portal.json"
            manifest.write_text(json.dumps({
                "schema_version": 1,
                "source": "direct-portal",
                "month": "2026-08",
                "coverage": [
                    _coverage("Bank of America", "complete", 1),
                    _coverage("Chase", "complete", 0),
                    _coverage("American Express", "complete", 1),
                ],
                "transactions": [
                    _portal_tx(
                        _local_id("bridged-cafe"),
                        "Bank of America",
                        "2026-08-12",
                        "Synthetic Cafe",
                        "Synthetic Cash Rewards",
                        "21.37",
                        kind="purchase",
                    ),
                    _portal_tx(
                        _local_id("bridged-market"),
                        "American Express",
                        "2026-08-15",
                        "Synthetic Market",
                        "Synthetic Rewards Preferred",
                        "63.25",
                        kind="purchase",
                    ),
                ],
            }), encoding="utf-8")
            direct_source = DirectPortalBudgetSource(_Authority(budgets), JsonPortalManifest(manifest).read())

            service.execute(direct_source, output, date(2026, 8, 1))

            refreshed = load_workbook(output, data_only=False)["Transactions"]
            rows = {
                refreshed.cell(row, 2).value: row
                for row in range(4, refreshed.max_row + 1)
            }
            cafe_row = rows["Synthetic Cafe"]
            market_row = rows["Synthetic Market"]
            self.assertEqual(refreshed.cell(cafe_row, 5).value, "Dining")
            self.assertEqual(refreshed.cell(cafe_row, 6).value, formula_link)
            self.assertEqual(refreshed.cell(cafe_row, 6).data_type, "f")
            self.assertEqual(refreshed.cell(cafe_row, 7).value, "Excluded from Budget")
            self.assertEqual(refreshed.cell(cafe_row, 8).value, "Existing confirmed payment treatment")
            self.assertEqual(refreshed.row_dimensions[cafe_row].height, 61.5)
            self.assertEqual(refreshed.cell(market_row, 5).value, "Groceries")
            self.assertEqual(refreshed.cell(market_row, 6).value, "Existing native Gmail evidence")
            self.assertEqual(refreshed.cell(market_row, 6).hyperlink.target, native_link)
            self.assertEqual(refreshed.cell(market_row, 7).value, "Review needed")
            self.assertEqual(refreshed.cell(market_row, 8).value, "Existing treatment review")
            self.assertEqual(refreshed.row_dimensions[market_row].height, 73.0)

    def test_exact_key_preserves_valid_treatment_and_user_detail_over_incoming_evidence(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            output = root / "review.xlsx"
            evidence_path = root / "evidence.json"
            budgets = [CategoryBudget("Flexible", "Dining", Decimal("250"))]
            identity = ("2026-08-18", "Synthetic Exact Merchant", "Synthetic Exact Card", "32.10")
            prior_transaction = _ledger_tx(
                *identity,
                raw={
                    "personal_finance_category": {
                        "primary": "LOAN_PAYMENTS",
                        "detailed": "LOAN_PAYMENTS_CREDIT_CARD_PAYMENT",
                        "confidence_level": "VERY_HIGH",
                    }
                },
            )
            service = BuildMonthlyReview(OpenpyxlReviewWorkbookWriter())
            service.execute(_BudgetSource(budgets, [prior_transaction]), output, date(2026, 8, 1))

            prior = load_workbook(output, data_only=False)
            sheet = prior["Transactions"]
            prior_link = "https://mail.google.com/mail/u/0/#all/priorExactNative123"
            sheet["E4"] = "Dining"
            sheet["F4"] = "User-confirmed exact-key detail"
            sheet["F4"].hyperlink = prior_link
            sheet.row_dimensions[4].height = 66.0
            self.assertEqual(sheet["G4"].value, "Excluded from Budget")
            self.assertEqual(sheet["H4"].value, "Credit-card payment")
            prior.save(output)

            evidence_path.write_text(json.dumps({"evidence": [{
                "match": {
                    "date": identity[0],
                    "merchant": identity[1],
                    "account": identity[2],
                    "amount": identity[3],
                },
                "status": "confirmed",
                "source": "gmail",
                "summary": "Conflicting incoming evidence",
                "items": ["Incoming item must not overwrite the user detail"],
                "gmail_thread_id": "incomingExactGmail123",
            }]}), encoding="utf-8")
            reclassified_transaction = _ledger_tx(
                *identity,
                raw={
                    "personal_finance_category": {
                        "primary": "GENERAL_MERCHANDISE",
                        "detailed": "GENERAL_MERCHANDISE_OTHER_GENERAL_MERCHANDISE",
                        "confidence_level": "VERY_HIGH",
                    }
                },
            )

            service.execute(
                _BudgetSource(budgets, [reclassified_transaction]),
                output,
                date(2026, 8, 1),
                evidence_reader=JsonEvidenceManifest(evidence_path),
            )

            refreshed = load_workbook(output, data_only=False)["Transactions"]
            self.assertEqual(refreshed["E4"].value, "Dining")
            self.assertEqual(refreshed["F4"].value, "User-confirmed exact-key detail")
            self.assertEqual(refreshed["F4"].hyperlink.target, prior_link)
            self.assertEqual(refreshed["G4"].value, "Excluded from Budget")
            self.assertEqual(refreshed["H4"].value, "Credit-card payment")
            self.assertEqual(refreshed.row_dimensions[4].height, 66.0)

    def test_unique_bridge_preserves_annotations_when_prior_treatment_is_unparseable(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            output = root / "review.xlsx"
            budgets = [CategoryBudget("Flexible", "Dining", Decimal("250"))]
            prior_source = _BudgetSource(
                budgets,
                [
                    _ledger_tx(
                        "2026-08-20",
                        "SQ * SYNTHETIC BAKERY 0042",
                        "Bank of America Synthetic Cash Rewards Visa",
                        "14.80",
                    )
                ],
            )
            service = BuildMonthlyReview(OpenpyxlReviewWorkbookWriter())
            service.execute(prior_source, output, date(2026, 8, 1))
            prior = load_workbook(output, data_only=False)
            sheet = prior["Transactions"]
            annotation_link = "https://mail.google.com/mail/u/0/#all/legacyBridgeNative123"
            sheet["E4"] = "Dining"
            sheet["F4"] = "Legacy user annotation"
            sheet["F4"].hyperlink = annotation_link
            sheet["G4"] = "Legacy treatment value"
            sheet["H4"] = "Legacy basis must be recomputed"
            sheet.row_dimensions[4].height = 64.0
            prior.save(output)

            manifest = root / "portal.json"
            manifest.write_text(json.dumps({
                "schema_version": 1,
                "source": "direct-portal",
                "month": "2026-08",
                "coverage": [
                    _coverage("Bank of America", "complete", 1),
                    _coverage("Chase", "complete", 0),
                    _coverage("American Express", "complete", 0),
                ],
                "transactions": [
                    _portal_tx(
                        _local_id("legacy-bakery-bridge"),
                        "Bank of America",
                        "2026-08-22",
                        "Synthetic Bakery",
                        "Synthetic Cash Rewards",
                        "14.80",
                        kind="purchase",
                    )
                ],
            }), encoding="utf-8")
            direct_source = DirectPortalBudgetSource(_Authority(budgets), JsonPortalManifest(manifest).read())

            service.execute(direct_source, output, date(2026, 8, 1))

            refreshed = load_workbook(output, data_only=False)["Transactions"]
            self.assertEqual(refreshed["E4"].value, "Dining")
            self.assertEqual(refreshed["F4"].value, "Legacy user annotation")
            self.assertEqual(refreshed["F4"].hyperlink.target, annotation_link)
            self.assertEqual(refreshed["G4"].value, "Included in Budget")
            self.assertEqual(refreshed["H4"].value, "Purchase")
            self.assertEqual(refreshed.row_dimensions[4].height, 64.0)

    def test_ambiguous_plaid_to_portal_bridge_fails_closed_without_copying_annotations(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            output = root / "review.xlsx"
            budgets = [CategoryBudget("Flexible", "Dining", Decimal("250"))]
            prior_source = _BudgetSource(
                budgets,
                [
                    _ledger_tx(
                        "2026-08-10",
                        "SQ * SYNTHETIC DELI 001",
                        "Bank of America Synthetic Cash Rewards Visa",
                        "18.40",
                    ),
                    _ledger_tx(
                        "2026-08-12",
                        "SYNTHETIC DELI #002",
                        "Bank of America Synthetic Cash Rewards Visa",
                        "18.40",
                    ),
                ],
            )
            service = BuildMonthlyReview(OpenpyxlReviewWorkbookWriter())
            service.execute(
                prior_source,
                output,
                date(2026, 8, 1),
                source_context=_complete_source_context(ImportSourceMode.PLAID),
            )
            prior = load_workbook(output, data_only=False)
            sheet = prior["Transactions"]
            for row, detail in ((4, "First private annotation"), (5, "Second private annotation")):
                sheet.cell(row, 5).value = "Dining"
                sheet.cell(row, 6).value = detail
                sheet.cell(row, 7).value = "Excluded from Budget"
                sheet.cell(row, 8).value = f"Prior treatment {row}"
                sheet.row_dimensions[row].height = 55 + row
            prior.save(output)

            manifest = root / "portal.json"
            manifest.write_text(json.dumps({
                "schema_version": 1,
                "source": "direct-portal",
                "month": "2026-08",
                "coverage": [
                    _coverage("Bank of America", "complete", 1),
                    _coverage("Chase", "complete", 0),
                    _coverage("American Express", "complete", 0),
                ],
                "transactions": [
                    _portal_tx(
                        _local_id("ambiguous-deli"),
                        "Bank of America",
                        "2026-08-11",
                        "Synthetic Deli",
                        "Synthetic Cash Rewards",
                        "18.40",
                        kind="purchase",
                    ),
                ],
            }), encoding="utf-8")
            direct_source = DirectPortalBudgetSource(_Authority(budgets), JsonPortalManifest(manifest).read())

            service.execute(direct_source, output, date(2026, 8, 1))

            refreshed = load_workbook(output, data_only=False)["Transactions"]
            self.assertIsNone(refreshed["E4"].value)
            self.assertEqual(refreshed["F4"].value, REVIEW_NEEDED_DEFAULT)
            self.assertEqual(refreshed["G4"].value, "Included in Budget")
            self.assertEqual(refreshed["H4"].value, "Purchase")
            self.assertNotIn("private annotation", str(refreshed["F4"].value).casefold())
            self.assertEqual(refreshed.row_dimensions[4].height, 18.0)

    def test_incomplete_direct_portal_coverage_is_explicit_in_result_and_workbook(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            manifest = root / "portal.json"
            output = root / "review.xlsx"
            manifest.write_text(json.dumps({
                "schema_version": 1,
                "source": "direct-portal",
                "month": "2026-08",
                "coverage": [
                    _coverage("Bank of America", "complete", 0),
                    _coverage("Chase", "partial", 0, period_end="2026-08-20", note="Session ended before month completion"),
                ],
                "transactions": [],
            }), encoding="utf-8")
            authority = _Authority([CategoryBudget("Flexible", "Groceries", Decimal("200"))])
            source = DirectPortalBudgetSource(authority, JsonPortalManifest(manifest).read())

            result = BuildMonthlyReview(OpenpyxlReviewWorkbookWriter()).execute(source, output, date(2026, 8, 1))

            self.assertFalse(result.source_coverage_complete)
            self.assertEqual(result.source_complete_count, 1)
            self.assertEqual(result.source_partial_count, 1)
            self.assertEqual(result.source_unavailable_count, 1)
            summary = load_workbook(output)["Budget Summary"]
            self.assertEqual(summary["H2"].value, "Incomplete")
            self.assertEqual(summary["H6"].value, "Unavailable")
            self.assertNotIn("Session ended", str(summary["J5"].value))

    def test_reconciliation_queue_is_separate_and_preserved_without_new_input(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            manifest = root / "portal.json"
            output = root / "review.xlsx"
            manifest.write_text(json.dumps({
                "schema_version": 1,
                "source": "direct-portal",
                "month": "2026-08",
                "coverage": [_coverage(name, "complete", 0) for name in DIRECT_PORTAL_INSTITUTIONS],
                "transactions": [],
                "reconciliation": {
                    "coverage": [
                        {
                            "institution": "Bank of America",
                            "portal_status": "Complete",
                            "portal_posted_count": 1,
                            "active_plaid_count": 0,
                            "matched_count": 0,
                            "portal_only_count": 0,
                            "plaid_only_count": 0,
                            "ambiguous_count": 0,
                            "plaid_removed_portal_visible_count": 1,
                        }
                    ],
                    "items": [
                        {
                            "local_id": _local_id("car-payment-queue"),
                            "institution": "Bank of America",
                            "date": "2026-08-21",
                            "amount": "785.00",
                            "classification": ReconciliationClassification.PLAID_REMOVED_PORTAL_VISIBLE.value,
                            "local_status": "Plaid marked removed",
                            "budget_effect": "Not included in active spend",
                            "review_note": "Car payment — pending human disposition",
                        }
                    ],
                },
            }), encoding="utf-8")
            authority = _Authority([CategoryBudget("Fixed", "Car Payment", Decimal("800"))])
            batch = JsonPortalManifest(manifest).read()
            source = DirectPortalBudgetSource(authority, batch)
            service = BuildMonthlyReview(OpenpyxlReviewWorkbookWriter())

            first = service.execute(source, output, date(2026, 8, 1))
            self.assertEqual(first.transaction_count, 0)
            self.assertEqual(first.reconciliation_removed_visible_count, 1)
            workbook = load_workbook(output, data_only=False)
            self.assertIn("Reconciliation", workbook.sheetnames)
            sheet = workbook["Reconciliation"]
            self.assertEqual(sheet.tables["ReconciliationCoverageTable"].ref, "A8:I12")
            self.assertEqual(sheet.tables["ReconciliationDetailTable"].ref, "A15:G16")
            self.assertEqual(sheet["D16"].value, "Plaid marked removed — still visible at institution")
            self.assertEqual(sheet["G16"].value, "Car payment — pending human disposition")
            self.assertEqual(workbook["Budget Summary"]["C4"].value, (
                "=SUMIFS('Transactions'!$D$4:$D$4,'Transactions'!$E$4:$E$4,A4,"
                "'Transactions'!$G$4:$G$4,\"<>Excluded from Budget\")"
            ))

            source_without_report = DirectPortalBudgetSource(authority, batch.without_reconciliation())
            second = service.execute(source_without_report, output, date(2026, 8, 1))
            preserved = load_workbook(output, data_only=False)["Reconciliation"]
            self.assertEqual(second.reconciliation_removed_visible_count, 1)
            self.assertEqual(preserved["D16"].value, "Plaid marked removed — still visible at institution")
            self.assertEqual(preserved.tables["ReconciliationDetailTable"].ref, "A15:G16")


def _coverage(
    institution: str,
    status: str,
    count: int,
    *,
    period_start: str = "2026-08-01",
    period_end: str = "2026-08-31",
    note: str = "",
) -> dict[str, object]:
    return {
        "institution": institution,
        "status": status,
        "period_start": period_start,
        "period_end": period_end,
        "posted_transaction_count": count,
        "note": note,
    }


def _portal_tx(
    local_id: str,
    institution: str,
    transaction_date: str,
    merchant: str,
    account: str,
    amount: str,
    *,
    kind: str = "",
) -> dict[str, str]:
    return {
        "local_id": local_id,
        "institution": institution,
        "date": transaction_date,
        "merchant": merchant,
        "account": account,
        "amount": amount,
        "status": "posted",
        "kind": kind,
    }


def _local_id(seed: str) -> str:
    return hashlib.sha256(seed.encode("utf-8")).hexdigest()


def _ledger_tx(
    transaction_date: str,
    merchant: str,
    account: str,
    amount: str,
    *,
    raw: dict[str, object] | None = None,
) -> LedgerTransaction:
    parsed_date = date.fromisoformat(transaction_date)
    return LedgerTransaction(
        transaction_date=parsed_date,
        month=date(parsed_date.year, parsed_date.month, 1),
        merchant=merchant,
        account=account,
        amount=Decimal(amount),
        personal_budget=True,
        source="Plaid",
        raw=dict(raw or {}),
    )


def _complete_source_context(mode: ImportSourceMode) -> ReviewSourceContext:
    return ReviewSourceContext(
        mode=mode,
        coverage=SourceCoverage(
            mode=mode,
            month=date(2026, 8, 1),
            rows=tuple(
                PortalInstitutionCoverage(
                    institution=institution,
                    status=PortalCoverageStatus.COMPLETE,
                    period_start=date(2026, 8, 1),
                    period_end=date(2026, 8, 31),
                    posted_transaction_count=0,
                )
                for institution in DIRECT_PORTAL_INSTITUTIONS
            ),
        ),
    )
