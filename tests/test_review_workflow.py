from __future__ import annotations

import hashlib
import io
import json
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from datetime import date
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook, load_workbook

from personal_budget_plaid_excel.application.monthly_review import BuildMonthlyReview
from personal_budget_plaid_excel.cli.main import main
from personal_budget_plaid_excel.domain.models import SyncItemState
from personal_budget_plaid_excel.infrastructure.connection_registry import ConnectionRegistry
from personal_budget_plaid_excel.infrastructure.evidence_manifest import JsonEvidenceManifest
from personal_budget_plaid_excel.infrastructure.review_workbook import OpenpyxlReviewWorkbookWriter
from personal_budget_plaid_excel.infrastructure.telemetry import Telemetry
from personal_budget_plaid_excel.infrastructure.workbook import BASE_HEADERS, EXTENSION_HEADERS, BudgetWorkbook


class MonthlyReviewWorkflowTests(unittest.TestCase):
    def test_doctor_reports_connection_registry_and_intended_output_paths(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "authority.xlsx"
            output_dir = root / "monthly"
            review_dir = root / "reviews"
            _write_synthetic_authority(source)
            ConnectionRegistry(output_dir / "connection-registry.json").replace([
                SyncItemState("private-item-one", "Private Institution One"),
                SyncItemState("private-item-two", "Private Institution Two"),
            ])
            stdout = io.StringIO()

            with redirect_stdout(stdout):
                code = main([
                    "--workbook", str(source),
                    "doctor",
                    "--month", "2026-08",
                    "--output-dir", str(output_dir),
                    "--review-output-dir", str(review_dir),
                    "--json",
                ])

            payload = json.loads(stdout.getvalue())
            self.assertEqual(code, 0)
            self.assertEqual(payload["month"], "2026-08")
            self.assertEqual(payload["connection_registry"], str((output_dir / "connection-registry.json").resolve()))
            self.assertTrue(payload["connection_registry_exists"])
            self.assertEqual(payload["registry_connection_count"], 2)
            self.assertEqual(payload["monthly_workbook_connection_count"], 0)
            self.assertEqual(payload["authority_connection_count"], 0)
            self.assertEqual(payload["intended_connection_count"], 2)
            self.assertEqual(payload["monthly_workbook"], str((output_dir / "budget-2026-08.xlsx").resolve()))
            self.assertFalse(payload["monthly_workbook_exists"])
            self.assertEqual(
                payload["review_workbook"],
                str((review_dir / "August-2026-Personal-Budget-Review.xlsx").resolve()),
            )
            self.assertFalse(payload["review_workbook_exists"])
            self.assertTrue(payload["staging_output_separate_from_authority"])
            self.assertTrue(payload["review_output_separate_from_authority"])
            self.assertTrue(payload["refresh_ready"])
            self.assertNotIn("private-item", stdout.getvalue())
            self.assertNotIn("Private Institution", stdout.getvalue())

    def test_builds_review_contract_without_mutating_authority(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "authority.xlsx"
            output = root / "reviews" / "August-2026-Personal-Budget-Review.xlsx"
            evidence_path = root / "evidence.json"
            _write_synthetic_authority(source)
            _write_evidence(evidence_path)
            before_hash = _sha256(source)

            result = BuildMonthlyReview(OpenpyxlReviewWorkbookWriter()).execute(
                BudgetWorkbook(source),
                output,
                date(2026, 8, 1),
                JsonEvidenceManifest(evidence_path),
                Telemetry(root / "review.jsonl"),
            )

            self.assertEqual(_sha256(source), before_hash)
            self.assertEqual(result.transaction_count, 4)
            self.assertEqual(result.category_count, 3)
            self.assertEqual(result.evidence_supported_count, 2)
            self.assertEqual(result.gmail_link_count, 1)
            self.assertEqual(result.review_needed_count, 2)
            self.assertIsNone(result.backup_path)

            workbook = load_workbook(output, data_only=False)
            self.assertEqual(workbook.sheetnames[:2], ["Transactions", "Budget Summary"])
            self.assertEqual(workbook["_ReviewState"].sheet_state, "veryHidden")
            transactions = workbook["Transactions"]
            summary = workbook["Budget Summary"]

            self.assertEqual([transactions.cell(3, col).value for col in range(1, 9)], [
                "Date",
                "Merchant",
                "Credit Card Account",
                "Amount",
                "Category",
                "Purchased Item Detail",
                "Budget Treatment",
                "Treatment Basis",
            ])
            self.assertEqual(transactions.tables["AugustTransactions"].ref, "A3:I7")
            self.assertEqual(summary.tables["AugustBudgetSummary"].ref, "A3:E6")
            self.assertEqual([summary.cell(row, 1).value for row in range(4, 7)], [
                "Groceries, Household", "Utilities", "Shopping"
            ])
            self.assertEqual([summary.cell(row, 2).value for row in range(4, 7)], [100, 200, 50])
            self.assertEqual(summary["A7"].value, "Overall Total")
            self.assertEqual(summary["B7"].value, "=SUM(B4:B6)")
            self.assertEqual(summary["C7"].value, "=SUM(C4:C6)")

            for row in range(4, 7):
                self.assertEqual(
                    summary.cell(row, 3).value,
                    (
                        "=SUMIFS('Transactions'!$D$4:$D$7,"
                        f"'Transactions'!$E$4:$E$7,A{row},"
                        "'Transactions'!$G$4:$G$7,\"<>Excluded from Budget\")"
                    ),
                )
                self.assertEqual(summary.cell(row, 4).value, f"=B{row}-C{row}")
                self.assertEqual(summary.cell(row, 5).value, f"=C{row}-B{row}")

            defined_name = workbook.defined_names["BudgetCategoryList"]
            self.assertEqual(defined_name.attr_text, "'Budget Summary'!$A$4:$A$6")
            validations = list(transactions.data_validations.dataValidation)
            self.assertEqual(len(validations), 1)
            self.assertEqual(validations[0].type, "list")
            self.assertEqual(validations[0].formula1, "=BudgetCategoryList")
            self.assertTrue(validations[0].allow_blank)
            self.assertIn("E4:E7", str(validations[0].sqref))

            self.assertEqual([transactions.cell(row, 5).value for row in range(4, 8)], [
                "Groceries, Household", "Utilities", None, "Shopping"
            ])
            self.assertIsInstance(transactions["A4"].value, date)
            self.assertIsInstance(transactions["D4"].value, (int, float))
            self.assertTrue(transactions["F4"].value.startswith("Amazon evidence"))
            self.assertIsNone(transactions["F4"].hyperlink)
            self.assertTrue(transactions["F5"].value.startswith("Gmail evidence"))
            self.assertEqual(
                transactions["F5"].hyperlink.target,
                "https://mail.google.com/mail/u/0/#all/18abcdeffedcba01",
            )
            self.assertTrue(transactions["F6"].value.startswith("Review needed"))
            self.assertTrue(transactions["F7"].value.startswith("Review needed"))
            self.assertTrue(transactions["F4"].alignment.wrap_text)
            self.assertEqual(
                [transactions.cell(row, 7).value for row in range(4, 8)],
                ["Included in Budget"] * 4,
            )
            self.assertEqual(transactions["H7"].value, "Refund/credit offset")
            self.assertEqual(result.excluded_non_spend_count, 0)
            self.assertEqual(result.treatment_review_needed_count, 0)
            self.assertEqual(workbook.calculation.calcMode, "auto")

            telemetry_text = (root / "review.jsonl").read_text(encoding="utf-8")
            for forbidden in ("Synthetic Amazon", "18abcdeffedcba01", "23.45", "Paper Towels"):
                self.assertNotIn(forbidden, telemetry_text)

    def test_ambiguous_or_invalid_evidence_never_creates_a_link_or_guess(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "authority.xlsx"
            output = root / "review.xlsx"
            evidence_path = root / "evidence.json"
            _write_synthetic_authority(source, duplicate_amazon=True)
            evidence_path.write_text(json.dumps({"evidence": [
                {
                    "match": {"date": "2026-08-03", "merchant": "Synthetic Amazon", "amount": "23.45"},
                    "status": "confirmed",
                    "source": "gmail",
                    "summary": "=HYPERLINK(\"https://example.invalid\",\"unsafe\")",
                    "items": ["=2+2", "quoted \"item\""],
                    "gmail_thread_id": "../../unsafe",
                },
                {
                    "match": {"date": "2026-08-05", "merchant": "Uncategorized Shop", "amount": "9.99"},
                    "status": "ambiguous",
                    "source": "gmail",
                    "review_reason": "multiple possible receipts",
                },
            ]}), encoding="utf-8")

            result = BuildMonthlyReview(OpenpyxlReviewWorkbookWriter()).execute(
                BudgetWorkbook(source), output, date(2026, 8, 1), JsonEvidenceManifest(evidence_path)
            )

            workbook = load_workbook(output, data_only=False)
            transactions = workbook["Transactions"]
            self.assertEqual(result.evidence_supported_count, 0)
            self.assertEqual(result.gmail_link_count, 0)
            for row in range(4, transactions.max_row + 1):
                self.assertTrue(str(transactions.cell(row, 6).value).startswith("Review needed"))
                self.assertIsNone(transactions.cell(row, 6).hyperlink)
                self.assertFalse(str(transactions.cell(row, 6).value).startswith("="))

    def test_rerun_preserves_review_edits_appends_once_and_creates_unique_backups(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "authority.xlsx"
            output = root / "review.xlsx"
            _write_synthetic_authority(source)
            service = BuildMonthlyReview(OpenpyxlReviewWorkbookWriter())

            service.execute(BudgetWorkbook(source), output, date(2026, 8, 1))
            first = load_workbook(output)
            first["Transactions"]["E6"] = "Shopping"
            first["Transactions"]["F6"] = "Verified locally — preserve this note"
            first.save(output)

            _append_transaction(source, "2026-08-06", "New Store", "Test Card", 7.50, "")
            second = service.execute(BudgetWorkbook(source), output, date(2026, 8, 1))
            third = service.execute(BudgetWorkbook(source), output, date(2026, 8, 1))

            workbook = load_workbook(output, data_only=False)
            transactions = workbook["Transactions"]
            self.assertEqual(second.transaction_count, 5)
            self.assertEqual(third.transaction_count, 5)
            self.assertEqual(transactions.max_row, 8)
            self.assertEqual(transactions["E6"].value, "Shopping")
            self.assertEqual(transactions["F6"].value, "Verified locally — preserve this note")
            self.assertEqual(transactions.tables["AugustTransactions"].ref, "A3:I8")
            self.assertIn("E4:E8", str(list(transactions.data_validations.dataValidation)[0].sqref))
            backups = sorted((output.parent / "backups").glob("review.*.xlsx"))
            self.assertGreaterEqual(len(backups), 2)
            self.assertEqual(len({backup.name for backup in backups}), len(backups))

    def test_non_spend_treatment_updates_summary_without_losing_manual_review_edits(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "authority.xlsx"
            output = root / "review.xlsx"
            evidence_path = root / "evidence.json"
            _write_synthetic_non_spend_authority(source)
            evidence_path.write_text(json.dumps({"evidence": [{
                "match": {
                    "date": "2026-08-01",
                    "merchant": "Synthetic Market",
                    "amount": "40.00",
                    "account": "Test Card",
                },
                "status": "confirmed",
                "source": "gmail",
                "summary": "Synthetic grocery receipt",
                "items": ["Synthetic pantry item"],
                "gmail_thread_id": "syntheticNative123",
            }]}), encoding="utf-8")
            service = BuildMonthlyReview(OpenpyxlReviewWorkbookWriter())

            first_result = service.execute(
                BudgetWorkbook(source),
                output,
                date(2026, 8, 1),
                JsonEvidenceManifest(evidence_path),
            )
            first = load_workbook(output, data_only=False)
            first_transactions = first["Transactions"]
            first_rows = _transaction_rows_by_merchant(first_transactions)
            formula_detail = (
                '=HYPERLINK("https://mail.google.com/mail/u/0/#all/syntheticFormula123",'
                '"Synthetic formula detail")'
            )
            first_transactions.cell(first_rows["AUTOPAY PAYMENT"], 5).value = "Shopping"
            first_transactions.cell(first_rows["AUTOPAY PAYMENT"], 6).value = formula_detail
            first.save(output)

            result = service.execute(BudgetWorkbook(source), output, date(2026, 8, 1))

            workbook = load_workbook(output, data_only=False)
            transactions = workbook["Transactions"]
            summary = workbook["Budget Summary"]
            rows = _transaction_rows_by_merchant(transactions)
            market_row = rows["Synthetic Market"]
            refund_row = rows["Synthetic Refund"]
            payment_row = rows["AUTOPAY PAYMENT"]
            ambiguous_row = rows["ONLINE PAYMENT"]

            self.assertEqual(first_result.gmail_link_count, 1)
            self.assertEqual(result.excluded_non_spend_count, 1)
            self.assertEqual(result.treatment_review_needed_count, 1)
            self.assertEqual(
                [transactions.cell(3, column).value for column in range(1, 9)],
                [
                    "Date",
                    "Merchant",
                    "Credit Card Account",
                    "Amount",
                    "Category",
                    "Purchased Item Detail",
                    "Budget Treatment",
                    "Treatment Basis",
                ],
            )
            self.assertEqual(transactions.tables["AugustTransactions"].ref, "A3:I7")
            validation = list(transactions.data_validations.dataValidation)
            self.assertEqual(len(validation), 1)
            self.assertEqual(validation[0].formula1, "=BudgetCategoryList")
            self.assertEqual(str(validation[0].sqref), "E4:E7")

            self.assertEqual(transactions.cell(payment_row, 5).value, "Shopping")
            self.assertEqual(transactions.cell(payment_row, 6).value, formula_detail)
            self.assertEqual(transactions.cell(payment_row, 7).value, "Excluded from Budget")
            self.assertEqual(transactions.cell(payment_row, 8).value, "Credit-card payment")

            market_detail = transactions.cell(market_row, 6)
            self.assertTrue(str(market_detail.value).startswith("Gmail evidence"))
            self.assertEqual(
                market_detail.hyperlink.target,
                "https://mail.google.com/mail/u/0/#all/syntheticNative123",
            )
            self.assertTrue(market_detail.alignment.wrap_text)
            self.assertEqual(transactions.cell(refund_row, 7).value, "Included in Budget")
            self.assertEqual(transactions.cell(refund_row, 8).value, "Refund/credit offset")
            self.assertEqual(transactions.cell(ambiguous_row, 5).value, "Groceries, Household")
            self.assertEqual(transactions.cell(ambiguous_row, 7).value, "Review needed")
            self.assertIn("Possible payment/transfer", transactions.cell(ambiguous_row, 8).value)

            self.assertEqual(_included_actual(transactions, "Groceries, Household"), Decimal("65"))
            self.assertEqual(_included_actual(transactions, "Shopping"), Decimal("0"))
            for row in range(4, 7):
                self.assertEqual(
                    summary.cell(row, 3).value,
                    (
                        "=SUMIFS('Transactions'!$D$4:$D$7,"
                        f"'Transactions'!$E$4:$E$7,A{row},"
                        "'Transactions'!$G$4:$G$7,\"<>Excluded from Budget\")"
                    ),
                )
            self.assertEqual(summary["C7"].value, "=SUM(C4:C6)")
            self.assertTrue(result.backup_path)
            self.assertTrue(Path(result.backup_path).exists())
            formulas = [
                cell.value
                for sheet in (transactions, summary)
                for row in sheet.iter_rows()
                for cell in row
                if isinstance(cell.value, str) and cell.value.startswith("=")
            ]
            self.assertFalse(any("#REF!" in formula.upper() for formula in formulas))

    def test_review_run_cli_is_offline_and_emits_only_aggregate_result(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "monthly.xlsx"
            output_dir = root / "outputs"
            evidence_path = root / "evidence.json"
            _write_synthetic_authority(source)
            _write_evidence(evidence_path)

            code = main([
                "--workbook", str(source),
                "review-run",
                "--month", "2026-08",
                "--output-dir", str(output_dir),
                "--evidence-json", str(evidence_path),
            ])

            self.assertEqual(code, 0)
            self.assertTrue((output_dir / "August-2026-Personal-Budget-Review.xlsx").exists())

    def test_empty_month_keeps_a_valid_editable_review_row(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "authority.xlsx"
            output = root / "September-2026-Personal-Budget-Review.xlsx"
            _write_synthetic_authority(source)

            result = BuildMonthlyReview(OpenpyxlReviewWorkbookWriter()).execute(
                BudgetWorkbook(source), output, date(2026, 9, 1)
            )

            self.assertEqual(result.transaction_count, 0)
            workbook = load_workbook(output, data_only=False)
            transactions = workbook["Transactions"]
            summary = workbook["Budget Summary"]
            self.assertEqual(transactions.tables["SeptemberTransactions"].ref, "A3:I4")
            self.assertEqual(str(list(transactions.data_validations.dataValidation)[0].sqref), "E4")
            self.assertEqual(
                summary["C4"].value,
                (
                    "=SUMIFS('Transactions'!$D$4:$D$4,'Transactions'!$E$4:$E$4,A4,"
                    "'Transactions'!$G$4:$G$4,\"<>Excluded from Budget\")"
                ),
            )

    def test_monthly_run_creates_review_output_by_default_from_fake_sync_only(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "authority.xlsx"
            output_dir = root / "monthly"
            sync_json = root / "sync.json"
            _write_synthetic_authority(source)
            sync_json.write_text(json.dumps({
                "item_id": "synthetic-item",
                "institution_name": "Synthetic Bank",
                "accounts": [{"account_id": "synthetic-account", "name": "Synthetic Card"}],
                "added": [{
                    "transaction_id": "synthetic-new-transaction",
                    "account_id": "synthetic-account",
                    "amount": 4.25,
                    "date": "2026-08-08",
                    "merchant_name": "Synthetic New Merchant",
                }],
                "next_cursor": "synthetic-cursor",
            }), encoding="utf-8")
            before_hash = _sha256(source)

            code = main([
                "--workbook", str(source),
                "monthly-run",
                "--source", "plaid",
                "--month", "2026-08",
                "--output-dir", str(output_dir),
                "--sync-json", str(sync_json),
            ])

            self.assertEqual(code, 0)
            self.assertEqual(_sha256(source), before_hash)
            self.assertTrue((output_dir / "budget-2026-08.xlsx").exists())
            review = output_dir / "August-2026-Personal-Budget-Review.xlsx"
            self.assertTrue(review.exists())
            workbook = load_workbook(review, data_only=False)
            transactions = workbook["Transactions"]
            self.assertTrue(any(transactions.cell(row, 2).value == "Synthetic New Merchant" for row in range(4, transactions.max_row + 1)))
            self.assertNotIn("gmail_send_email", json.dumps(workbook.sheetnames))

    def test_monthly_run_without_connections_is_explicitly_not_fresh(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "authority.xlsx"
            output_dir = root / "monthly"
            _write_synthetic_authority(source)
            before_hash = _sha256(source)
            stdout = io.StringIO()
            stderr = io.StringIO()

            with redirect_stdout(stdout), redirect_stderr(stderr):
                code = main([
                    "--workbook", str(source),
                    "monthly-run",
                    "--source", "plaid",
                    "--month", "2026-08",
                    "--output-dir", str(output_dir),
                ])

            payload = json.loads(stdout.getvalue().splitlines()[-1])
            self.assertEqual(code, 2)
            self.assertEqual(payload["refresh_status"], "not_connected")
            self.assertFalse(payload["fresh"])
            self.assertEqual(_sha256(source), before_hash)
            self.assertTrue(Path(payload["monthly_workbook"]).exists())
            self.assertTrue(Path(payload["review_workbook"]).exists())
            self.assertIn("No Plaid items are connected", stderr.getvalue())

    def test_skill_routes_monthly_updates_to_review_and_keeps_external_sources_read_only(self) -> None:
        skill = (Path(__file__).resolve().parents[1] / "codex_skills" / "personal-budget-plaid-excel" / "SKILL.md").read_text(encoding="utf-8")
        self.assertIn("monthly-run --month YYYY-MM --source plaid", skill)
        self.assertIn("monthly-run --month YYYY-MM --source direct-portal", skill)
        self.assertIn("without `--source`", skill)
        self.assertIn("source-choice-required", skill)
        self.assertIn("review-run", skill)
        self.assertIn("Purchased Item Detail", skill)
        self.assertIn("read-only", skill)
        self.assertIn("Review needed", skill)
        self.assertNotIn("For a normal budget refresh/report request, run", skill)
        self.assertNotIn("Send the report through Gmail", skill)


def _write_synthetic_authority(path: Path, *, duplicate_amazon: bool = False) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)
    manual = workbook.create_sheet("Manual Budget")
    manual.append(["Cost Type", "Category", "Budget"])
    manual.append(["Flexible", "Groceries, Household", 100])
    manual.append(["Fixed", "Utilities", 200])
    manual.append(["Flexible", "Shopping", 50])
    manual.append(["Budget", None, 350])

    actual = workbook.create_sheet("Actual Charges")
    for col, header in enumerate(BASE_HEADERS + EXTENSION_HEADERS, start=1):
        actual.cell(5, col).value = header
    _append_actual_row(actual, "2026-08-03", "Synthetic Amazon", "Test Card", 23.45, "Groceries, Household", "tx-1")
    if duplicate_amazon:
        _append_actual_row(actual, "2026-08-03", "Synthetic Amazon", "Other Card", 23.45, "Groceries, Household", "tx-duplicate")
    _append_actual_row(actual, "2026-08-04", "Synthetic Utility", "Test Card", 80.00, "Utilities", "tx-2")
    _append_actual_row(actual, "2026-08-05", "Uncategorized Shop", "Test Card", 9.99, "", "tx-3")
    _append_actual_row(actual, "2026-08-05", "Synthetic Refund", "Test Card", -5.00, "Shopping", "tx-4", pending=True)
    workbook.save(path)


def _write_synthetic_non_spend_authority(path: Path) -> None:
    _write_synthetic_authority(path)
    workbook = load_workbook(path)
    actual = workbook["Actual Charges"]
    actual.delete_rows(6, max(0, actual.max_row - 5))
    _append_actual_row(
        actual,
        "2026-08-01",
        "Synthetic Market",
        "Test Card",
        40.00,
        "Groceries, Household",
        "tx-purchase",
        pfc_primary="GENERAL_MERCHANDISE",
        pfc_detailed="GENERAL_MERCHANDISE_SUPERSTORES",
        pfc_confidence="VERY_HIGH",
    )
    _append_actual_row(
        actual,
        "2026-08-02",
        "Synthetic Refund",
        "Test Card",
        -5.00,
        "Groceries, Household",
        "tx-refund",
        pfc_primary="GENERAL_MERCHANDISE",
        pfc_detailed="GENERAL_MERCHANDISE_SUPERSTORES",
        pfc_confidence="VERY_HIGH",
    )
    _append_actual_row(
        actual,
        "2026-08-03",
        "AUTOPAY PAYMENT",
        "Test Card",
        500.00,
        "Shopping",
        "tx-payment",
        pfc_primary="LOAN_PAYMENTS",
        pfc_detailed="LOAN_PAYMENTS_CREDIT_CARD_PAYMENT",
        pfc_confidence="VERY_HIGH",
    )
    _append_actual_row(
        actual,
        "2026-08-04",
        "ONLINE PAYMENT",
        "Test Card",
        30.00,
        "Groceries, Household",
        "tx-ambiguous",
        pfc_primary="LOAN_PAYMENTS",
        pfc_detailed="LOAN_PAYMENTS_OTHER_PAYMENT",
        pfc_confidence="LOW",
    )
    workbook.save(path)


def _append_transaction(path: Path, tx_date: str, merchant: str, account: str, amount: float, category: str) -> None:
    workbook = load_workbook(path)
    actual = workbook["Actual Charges"]
    _append_actual_row(actual, tx_date, merchant, account, amount, category, f"tx-{actual.max_row + 1}")
    workbook.save(path)


def _append_actual_row(
    sheet,
    tx_date: str,
    merchant: str,
    account: str,
    amount: float,
    category: str,
    transaction_id: str,
    *,
    pending: bool = False,
    pfc_primary: str = "",
    pfc_detailed: str = "",
    pfc_confidence: str = "",
) -> None:
    headers = {sheet.cell(5, col).value: col for col in range(1, sheet.max_column + 1)}
    row = sheet.max_row + 1
    values = {
        "Date": date.fromisoformat(tx_date),
        "Month": date(2026, 8, 1),
        "Merchant": merchant,
        "Account": account,
        "Saved Category": category,
        "Amount": amount,
        "Personal Budget?": "Yes",
        "Evidence": "Synthetic fixture",
        "Selected Category": category,
        "Plaid Transaction ID": transaction_id,
        "Plaid Status": "pending" if pending else "posted",
        "Plaid Removed?": "No",
        "Plaid PFC Primary": pfc_primary,
        "Plaid PFC Detailed": pfc_detailed,
        "Plaid PFC Confidence": pfc_confidence,
    }
    for header, value in values.items():
        sheet.cell(row, headers[header]).value = value


def _transaction_rows_by_merchant(sheet) -> dict[str, int]:
    return {
        str(sheet.cell(row, 2).value): row
        for row in range(4, sheet.max_row + 1)
        if sheet.cell(row, 2).value not in (None, "")
    }


def _included_actual(sheet, category: str) -> Decimal:
    return sum(
        (
            Decimal(str(sheet.cell(row, 4).value))
            for row in range(4, sheet.max_row + 1)
            if sheet.cell(row, 5).value == category
            and sheet.cell(row, 7).value != "Excluded from Budget"
        ),
        Decimal("0"),
    )


def _write_evidence(path: Path) -> None:
    path.write_text(json.dumps({"evidence": [
        {
            "match": {"date": "2026-08-03", "merchant": "Synthetic Amazon", "amount": "23.45", "account": "Test Card"},
            "status": "confirmed",
            "source": "amazon",
            "summary": "Household order",
            "items": ["Paper Towels", "Dish Soap"],
        },
        {
            "match": {"date": "2026-08-04", "merchant": "Synthetic Utility", "amount": "80.00"},
            "status": "confirmed",
            "source": "gmail",
            "summary": "Monthly utility confirmation",
            "items": ["Utility service"],
            "gmail_thread_id": "18abcdeffedcba01",
        },
        {
            "match": {"date": "2026-08-05", "merchant": "Uncategorized Shop", "amount": "9.99"},
            "status": "ambiguous",
            "source": "gmail",
            "review_reason": "multiple possible receipts",
        },
    ]}), encoding="utf-8")


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()
