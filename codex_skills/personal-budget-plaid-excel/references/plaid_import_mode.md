# Plaid Import Mode

Use this mode only after the user explicitly chooses `Plaid`.

## Preflight and collection

1. Run `doctor` with the requested month and intended output paths before mutation. Verify the actual saved connection registry, the existing month staging path, the intended standalone review path, and that the canonical Manual Budget workbook remains a separate authority.
2. The expected multi-connection set is Bank of America, Chase, and American Express. Confirm all three saved connections are ready. Do not assume a workbook connection list is the registry of record.
3. Run the current idempotent `monthly-run --source plaid`, never the legacy `run` report/email workflow. Append only unseen active transactions, apply provider modifications/removals, preserve existing user-owned fields, and leave newly imported Category values blank.
4. If any connection fails or is unavailable, surface the institution and mark source coverage incomplete. Do not claim the month is fully refreshed and do not silently switch to Direct portal.

```bash
PYTHONPATH=src <python> -m personal_budget_plaid_excel.cli.main \
  doctor --month YYYY-MM --output-dir monthly-budgets --json
PYTHONPATH=src <python> -m personal_budget_plaid_excel.cli.main \
  monthly-run --month YYYY-MM --source plaid
```

Confirm `refresh_ready`, the three intended connections, resolved staging/review paths, and aggregate per-institution refresh status before describing the output as fresh.

## Provider state and coverage

- Preserve provider-added, modified, pending, posted, and removed state locally as needed for idempotency and audit. Do not expose raw provider records, connection tokens, account identifiers, merchants, or amounts in chat or telemetry.
- A Plaid removal status is provider-supplied state, not a manual deletion. Keep relevant removed records visible in the reconciliation queue without counting them in active spending or duplicating them.
- When portal evidence shows a provider-removed transaction is still posted at the institution, label it exactly `Plaid marked removed — still visible at institution`. Keep it pending human disposition and do not restore it to spend automatically.
- Record selected mode `Plaid`, requested month, expected/available institution counts, per-institution refresh status, and completeness in the review workbook's source coverage metadata. Use aggregate status only.
- Run the shared review builder and workbook validation in `SKILL.md` after collection. Never send email as part of this mode.
