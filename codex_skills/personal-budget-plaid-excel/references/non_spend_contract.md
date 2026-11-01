# Non-Spend Treatment Contract

Use this contract for credit-card payments, account transfers, income credits, refunds, and descriptions that merely resemble them. Classification is auditable and conservative; it never hides a transaction or overwrites a manual Category.

## Visible dispositions

- `Included in Budget`: ordinary purchases, budgeted loan payments, and supported refund/credit offsets. These rows may affect Actuals when categorized.
- `Excluded from Budget`: confidently identified credit-card payments, internal account transfers, or income/account credits. Preserve any manual Category, but omit the row from Actuals and suppress category suggestions.
- `Review needed`: ambiguous payment/transfer or unexplained negative-direction activity. Keep the row visible, counted in Actuals if it already has a Category, and suppress category suggestions until reviewed.

Show the disposition in `Budget Treatment` and a concise, non-sensitive explanation in `Treatment Basis`. Never place raw descriptions, amounts, provider payloads, account identifiers, or evidence content in telemetry.

## Conservative precedence

Apply the strongest supported rule first:

1. Preserve the user's existing Category, Purchased Item Detail, and Gmail hyperlink regardless of treatment.
2. Treat an explicit provider income classification as `Excluded from Budget` with an income/account-credit basis.
3. Exclude a credit-card payment or internal transfer only when a specific provider classification has strong confidence, a tightly anchored account-payment/transfer description corroborates it, or a repeated/opposing-account pattern supplies equivalent corroboration.
4. Keep identified car, mortgage, student-loan, and other budgeted debt payments included; `LOAN_PAYMENTS` alone is not proof of an internal credit-card payment.
5. Keep a supported merchant refund or credit offset included so a manual category can net against spending. A negative amount alone is not proof of a refund.
6. Let ordinary purchase classifications defeat generic words such as `payment`, `credit`, or `transfer` unless stronger account-activity evidence exists. Processor or merchant names containing those words must not be excluded by text alone.
7. Send low-confidence, conflicting, generic-token, and otherwise ambiguous cases to `Review needed`. Never promote ambiguity to exclusion merely to balance totals.

Historical patterns may corroborate but not independently invent a classification. Same-account payment clusters require multiple consistent rows plus strong provider-supported examples. Opposing-account transfers require equal absolute amounts, opposite directions, different accounts, a narrow date window, and transfer evidence; common-amount collisions remain review-only.

## Summary, automation, and validation

- Budget Summary Actual formulas exclude only the exact `Excluded from Budget` disposition. `Review needed` rows stay counted until resolved.
- Category suggestion must short-circuit for both `Excluded from Budget` and `Review needed` rows.
- Reruns must recompute treatment from source evidence while preserving user-owned fields and backing up an existing review workbook before replacement.
- Emit only aggregate treatment counts and statuses. Do not log matched text, amounts, merchants, accounts, identifiers, or row-level evidence.
- Test payments, transfers, refunds/credits, budgeted loan payments, genuine purchases with similar words, historical corroboration, and ambiguous cases using disposable synthetic fixtures only. Do not call live Plaid, Gmail, or Amazon in tests.
