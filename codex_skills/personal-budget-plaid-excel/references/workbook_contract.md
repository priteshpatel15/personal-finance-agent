# Workbook Contract

## Authority and staging

- Canonical authority workbook: the path in `--workbook`, else `$BUDGET_WORKBOOK`, else `./budget.xlsx` (relative to the working directory).
- `Manual Budget` rows are the sole authority for category order, spelling, and budget amounts.
- `Actual Charges` is the transaction ledger. `Selected Category` is user-controlled; `Saved Category` is historical evidence; `Suggested Category` is advisory only.
- `monthly-run --month YYYY-MM --source plaid` refreshes `monthly-budgets/budget-YYYY-MM.xlsx`, never the canonical workbook. It reconciles the saved registry with target and authority connection state, preserves existing selections, appends unseen transactions, and leaves new selections blank. Direct portal mode skips the Plaid staging workbook and builds the same review output from its validated local manifest.
- `doctor --month YYYY-MM --output-dir ... --json` reports the resolved registry, staging, and review paths plus aggregate connection counts. A normal refresh proceeds only when `refresh_ready` is true. A no-connection `monthly-run` returns nonzero and labels its structural output `fresh: false` rather than claiming a fresh refresh.

## Standalone review output

- Default filename: `<Month>-<YYYY>-Personal-Budget-Review.xlsx`, separate from both canonical and staging workbooks.
- Required visible sheets: `Transactions` and `Budget Summary`; `Reconciliation` is also visible when a supported queue exists. A hidden implementation sheet may store opaque review keys only; raw Plaid/Gmail/account identifiers are forbidden.
- `Transactions` columns: Date, Merchant, Credit Card Account, Amount, Category, Purchased Item Detail, Budget Treatment, Treatment Basis. Dates and amounts remain typed values.
- Category values are blank or one of the Manual Budget categories. A workbook-level `BudgetCategoryList` points to the Budget Summary category cells; one list validation covers all Category data rows and permits blanks.
- Budget Treatment is one of `Included in Budget`, `Excluded from Budget`, or `Review needed`. Treatment Basis is a concise, non-sensitive reason visible beside it.
- Budget Summary preserves Manual Budget category order and amounts. Actual is a `SUMIFS` over Transactions Category and Amount that excludes only `Excluded from Budget`; `Review needed` remains counted until verified. Remaining, Variance, and Overall Total are formula-driven.
- Treatment does not overwrite Category. Preserve every manual Category, Purchased Item Detail, and Gmail hyperlink even when a confident non-spend row is excluded from Actuals.
- Do not auto-categorize rows marked `Excluded from Budget` or `Review needed`; leave ambiguous rows visible and auditable.
- Existing review Category and supported detail/link values are preserved by opaque stable transaction keys. New or unresolved rows receive `Review needed`.
- Gmail evidence links must be clickable and target only `https://mail.google.com/` using the actual connector-returned thread/message identifier. Amazon evidence is itemized text and does not invent a source link.
- Existing review outputs receive unique recoverable backups before atomic replacement.

## Import-source and reconciliation coverage

- Every monthly review visibly records the selected import mode (`Plaid` or `Direct portal`), requested month, expected institutions, per-institution availability/coverage status, and overall completeness. Bank of America, Chase, and American Express are the expected institution set unless the user explicitly changes it.
- Source metadata contains aggregate status only. Never put credentials, tokens, account identifiers, raw payloads, row descriptions, merchants, amounts, receipt contents, Gmail identifiers, or portal session data in metadata or telemetry.
- An unavailable institution, authentication handoff, failed Plaid connection, uncertain date range, or incomplete portal pagination makes overall coverage incomplete. The workbook and report must not describe incomplete coverage as a complete budget.
- Preserve reconciliation entries across idempotent rebuilds. Distinguish active imported spend, `Plaid marked removed — still visible at institution`, and ambiguous/source-only matches; do not describe provider-supplied removal state as a manual or local deletion.
- Provider-removed portal-visible rows remain in the reconciliation/review queue pending human disposition. Do not add them to active spend totals, silently restore them, or duplicate an active transaction.

## Validation

Confirm source workbook hash is unchanged, transaction/category/treatment counts reconcile, source mode and coverage metadata are present, reconciliation dispositions remain auditable, table and validation ranges end at the last data row, defined-name authority is correct, Gmail hyperlinks are valid, treatment values are valid, only confidently excluded rows are omitted from Actuals, formulas contain no broken references, and Excel recalculation is automatic. Report treatment and source results only as aggregate counts.
