# Direct Portal Manifest Contract

Create this local JSON only after the user chooses `Direct portal` and the three institution reads have completed or stopped. Restrict the file to the local user. Never paste its transaction rows into chat, telemetry, or a report.

## Schema

Use schema version `1`, source `direct-portal`, and the requested `YYYY-MM`. Include one coverage record for each accessible institution; omitted institutions are normalized to `unavailable`.

```json
{
  "schema_version": 1,
  "source": "direct-portal",
  "month": "2026-08",
  "coverage": [
    {
      "institution": "Bank of America",
      "status": "complete",
      "period_start": "2026-08-01",
      "period_end": "2026-08-31",
      "posted_transaction_count": 1
    },
    {
      "institution": "Chase",
      "status": "partial",
      "period_start": "2026-08-01",
      "period_end": "2026-08-20",
      "posted_transaction_count": 0
    }
  ],
  "transactions": [
    {
      "institution": "Bank of America",
      "date": "2026-08-03",
      "merchant": "Synthetic example merchant",
      "account": "Checking ending in synthetic label",
      "amount": "12.34",
      "status": "posted",
      "kind": "purchase",
      "occurrence": 1
    }
  ]
}
```

The example is synthetic. The real local file may contain the minimum visible row fields needed for the workbook, but must not contain credentials, cookies, browser tokens, full account numbers, statement documents, or raw HTML.

## Field rules

- Institution is exactly `Bank of America`, `Chase`, or `American Express`.
- Coverage status is `complete`, `partial`, or `unavailable`. Complete requires the full requested period and a declared count equal to the accepted posted rows. The reader downgrades inconsistent complete claims to partial.
- Transactions must be clearly `posted` and dated within the requested month. Pending or out-of-month rows are not active monthly spend.
- Amount is the signed amount as shown by the portal and remains a JSON string to avoid binary rounding.
- Account is a non-secret visible label, never a full account number.
- `kind` may be `purchase`, `credit-card-payment`, `account-transfer`, `refund`, `income`, `budgeted-loan-payment`, `car-payment`, `mortgage-payment`, `student-loan-payment`, `ambiguous`, or blank. Use a specific kind only when portal evidence is clear; otherwise use `ambiguous` or blank.
- `local_id` is optional. When omitted, the reader hashes normalized institution, date, local account label, merchant, signed amount, and `occurrence`. A repeated capture with the same fields and occurrence deduplicates. Use increasing occurrence values only for genuinely distinct identical postings.

An optional `reconciliation` object may carry the supported aggregate coverage and review-queue rows defined by the workbook contract. It never authorizes adding a discrepancy to Transactions.

## Handoff and retention

Pass the absolute local path through `--portal-json`. Keep the normalized manifest permission-restricted and local while the review may need an idempotent rebuild. Never attach it to email or upload it. Remove it only at the user's request or under a cleanup policy disclosed before collection.
