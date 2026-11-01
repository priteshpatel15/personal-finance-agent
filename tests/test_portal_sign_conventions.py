"""Disposable synthetic checks for explicit portal amount conventions."""
import hashlib
import json
import tempfile
import unittest
from datetime import date
from decimal import Decimal
from pathlib import Path

from openpyxl import load_workbook

from personal_budget_plaid_excel.application.monthly_review import BuildMonthlyReview
from personal_budget_plaid_excel.domain.models import CategoryBudget
from personal_budget_plaid_excel.infrastructure.portal_manifest import JsonPortalManifest
from personal_budget_plaid_excel.infrastructure.portal_source import DirectPortalBudgetSource
from personal_budget_plaid_excel.infrastructure.review_workbook import OpenpyxlReviewWorkbookWriter
from personal_budget_plaid_excel.infrastructure.telemetry import Telemetry


class SyntheticAuthority:
    def read_budgets(self):
        return [CategoryBudget('Flexible', 'Synthetic category', Decimal('1000'))]


def payload(amount='-17.25', kind='purchase', convention='debit-negative', version=2):
    row = {'institution': 'Bank of America', 'date': '2026-09-03',
           'merchant': 'Synthetic merchant', 'account': 'Synthetic checking',
           'amount': amount, 'status': 'posted', 'kind': kind}
    if convention is not None:
        row['amount_convention'] = convention
    return {'schema_version': version, 'source': 'direct-portal', 'month': '2026-09',
            'coverage': [{'institution': 'Bank of America', 'status': 'partial',
                          'period_start': '2026-09-01', 'period_end': '2026-09-14',
                          'posted_transaction_count': 1}], 'transactions': [row]}


def read(root, content):
    manifest = root / 'synthetic.json'
    manifest.write_text(json.dumps(content))
    return JsonPortalManifest(manifest).read()


class PortalSignConventionTests(unittest.TestCase):
    def test_legacy_passthrough_and_raw_identity_remain_stable(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            legacy = read(root, payload(version=1, convention=None))
            explicit = read(root, payload())
            self.assertEqual(legacy.transactions[0].amount, Decimal('-17.25'))
            self.assertEqual(explicit.transactions[0].amount, Decimal('-17.25'))
            self.assertEqual(legacy.transactions[0].local_id, explicit.transactions[0].local_id)
            self.assertEqual(DirectPortalBudgetSource(SyntheticAuthority(), legacy).read_transactions()[0].amount, Decimal('-17.25'))
            self.assertEqual(DirectPortalBudgetSource(SyntheticAuthority(), explicit).read_transactions()[0].amount, Decimal('17.25'))

    def test_schema_two_requires_supported_explicit_convention(self):
        with tempfile.TemporaryDirectory() as tmp:
            for value in (None, '', 'guess-from-sign', 'absolute'):
                with self.subTest(convention=value), self.assertRaises(ValueError):
                    read(Path(tmp), payload(convention=value))

    def test_conventions_preserve_direction_for_all_kinds(self):
        cases = [('purchase', '-17.25', '17.25'), ('refund', '4.25', '-4.25'),
                 ('income', '25', '-25'), ('account-transfer', '6', '-6'),
                 ('account-transfer', '-6', '6'), ('credit-card-payment', '-9', '9'),
                 ('car-payment', '-12', '12'), ('mortgage-payment', '-12', '12'),
                 ('student-loan-payment', '-12', '12'), ('budgeted-loan-payment', '-12', '12'),
                 ('ambiguous', '4', '-4'), ('', '4', '-4'), ('purchase', '0', '0')]
        with tempfile.TemporaryDirectory() as tmp:
            for kind, raw, expected in cases:
                with self.subTest(kind=kind, raw=raw):
                    batch = read(Path(tmp), payload(raw, kind))
                    row = DirectPortalBudgetSource(SyntheticAuthority(), batch).read_transactions()[0]
                    self.assertEqual(row.amount, Decimal(expected))
                    self.assertEqual(batch.transactions[0].amount, Decimal(raw))
                    if kind == 'account-transfer':
                        self.assertEqual(row.raw['personal_finance_category']['primary'], 'TRANSFER_IN' if Decimal(expected) < 0 else 'TRANSFER_OUT')
            card = read(Path(tmp), payload('17.25', convention='expense-positive'))
            self.assertEqual(DirectPortalBudgetSource(SyntheticAuthority(), card).read_transactions()[0].amount, Decimal('17.25'))

    def test_invalid_nonfinite_amounts_still_fail(self):
        with tempfile.TemporaryDirectory() as tmp:
            for amount in ('NaN', 'Infinity', '-Infinity'):
                with self.subTest(amount=amount), self.assertRaises(ValueError):
                    read(Path(tmp), payload(amount))

    def test_duplicates_occurrences_and_conflicting_convention(self):
        with tempfile.TemporaryDirectory() as tmp:
            content = payload()
            content['transactions'] *= 2
            self.assertEqual(len(read(Path(tmp), content).transactions), 1)
            content['transactions'][1] = dict(content['transactions'][1], occurrence=2)
            self.assertEqual(len(read(Path(tmp), content).transactions), 2)
            content['transactions'][1] = dict(content['transactions'][0], amount_convention='expense-positive')
            with self.assertRaises(ValueError):
                read(Path(tmp), content)

    def test_review_rebuild_preserves_user_fields_and_excludes_only_supported_nonspend(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            content = payload()
            for index, (kind, amount) in enumerate([('refund', '4.25'), ('income', '25'), ('account-transfer', '-6'), ('ambiguous', '4')], start=1):
                row = dict(content['transactions'][0], kind=kind, amount=amount,
                           merchant=f'Synthetic merchant {index}')
                content['transactions'].append(row)
            content['coverage'][0]['posted_transaction_count'] = len(content['transactions'])
            batch = read(root, content)
            source = DirectPortalBudgetSource(SyntheticAuthority(), batch)
            output = root / 'review.xlsx'
            builder = BuildMonthlyReview(OpenpyxlReviewWorkbookWriter())
            first = builder.execute(source, output, date(2026, 9, 1), telemetry=Telemetry(root / 'events.jsonl'))
            wb = load_workbook(output)
            ws = wb['Transactions']
            merchant_rows = {ws.cell(r, 2).value: r for r in range(4, ws.max_row + 1)}
            r = merchant_rows['Synthetic merchant']
            ws.cell(r, 5, 'Synthetic category')
            ws.cell(r, 6, 'Review needed: user annotation')
            wb.save(output)
            second = builder.execute(source, output, date(2026, 9, 1), telemetry=Telemetry(root / 'events.jsonl'))
            result = load_workbook(output)
            ws = result['Transactions']
            actual = {ws.cell(r, 2).value: (ws.cell(r, 4).value, ws.cell(r, 7).value) for r in range(4, ws.max_row + 1)}
            self.assertEqual(actual['Synthetic merchant'], (17.25, 'Included in Budget'))
            self.assertEqual(actual['Synthetic merchant 1'], (-4.25, 'Included in Budget'))
            self.assertEqual(actual['Synthetic merchant 2'][1], 'Excluded from Budget')
            self.assertEqual(actual['Synthetic merchant 3'][1], 'Excluded from Budget')
            self.assertEqual(actual['Synthetic merchant 4'][1], 'Review needed')
            r = next(r for r in range(4, ws.max_row + 1) if ws.cell(r, 2).value == 'Synthetic merchant')
            self.assertEqual(ws.cell(r, 5).value, 'Synthetic category')
            self.assertEqual(ws.cell(r, 6).value, 'Review needed: user annotation')
            self.assertEqual(first.transaction_count, second.transaction_count)
            self.assertEqual(second.transaction_count, 5)
            self.assertTrue(second.backup_path)
            self.assertFalse(second.source_coverage_complete)
            self.assertIn('Excluded from Budget', result['Budget Summary']['C4'].value)
            self.assertNotIn('Synthetic merchant', (root / 'events.jsonl').read_text())
            self.assertNotIn('17.25', (root / 'events.jsonl').read_text())


if __name__ == '__main__':
    unittest.main()
