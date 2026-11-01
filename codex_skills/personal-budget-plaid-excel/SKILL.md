---
name: personal-budget-plaid-excel
description: Import a monthly personal budget from Plaid or signed-in bank portals and build its standalone Excel review workbook while preserving Manual Budget authority and user choices. Use when the user asks to update, refresh, or review a monthly personal budget; optional Amazon and Gmail research is read-only and email sending is never the default.
---

# Personal Budget Plaid Excel

Use the Python CLI in this repository (`<repo-root>`). Read [references/workbook_contract.md](references/workbook_contract.md) for shared workbook invariants. Read only the reference for the chosen import source: [references/plaid_import_mode.md](references/plaid_import_mode.md) or [references/direct_portal_import_mode.md](references/direct_portal_import_mode.md). For Direct portal collection, also read [references/portal_manifest_contract.md](references/portal_manifest_contract.md). Read [references/non_spend_contract.md](references/non_spend_contract.md) when payments, transfers, credits, or category-suggestion behavior are in scope. Read [references/review_evidence_contract.md](references/review_evidence_contract.md) only when purchase research is in scope.

## Choose the Import Source First

Every monthly import uses exactly one source mode:

- `Plaid`: refresh all saved Plaid connections through the existing multi-connection sync. Read [references/plaid_import_mode.md](references/plaid_import_mode.md).
- `Direct portal`: collect posted transactions from the user's already signed-in Bank of America, Chase, and American Express Chrome sessions. Read [references/direct_portal_import_mode.md](references/direct_portal_import_mode.md).

Before any transaction collection, determine whether the user explicitly chose `Plaid` or `Direct portal`. Honor an explicit choice. If neither was stated, ask the user which of those two modes to use and wait for the answer; do not infer a default. Calling `monthly-run` without `--source` must return its source-choice-required result without collecting data.

Do not switch or fall back to the other source when the chosen mode is unavailable. Report the source-specific blocker or incomplete coverage and let the user decide whether to choose another mode.

## Shared Monthly Review

1. In Codex Desktop, call `load_workspace_dependencies` and use its bundled Python executable. Otherwise use the active project Python.
2. Resolve and verify the intended staging and review-output paths before mutation. Reuse the existing month artifacts rather than creating duplicate filenames.
3. Run the chosen idempotent import with an explicit source:

```bash
cd <repo-root>
PYTHONPATH=src <python> -m personal_budget_plaid_excel.cli.main \
  monthly-run --month YYYY-MM --source plaid

# Direct portal mode consumes the local manifest produced by the authorized Chrome read.
PYTHONPATH=src <python> -m personal_budget_plaid_excel.cli.main \
  monthly-run --month YYYY-MM --source direct-portal \
  --portal-json /absolute/path/to/local-portal-manifest.json
```

Both modes must feed the same downstream standalone review-workbook builder. The output must retain the Manual Budget categories and amounts, category dropdown and formulas, Purchased Item Detail and supported evidence, payment/transfer treatment, reconciliation queue, backups, and all user-owned values. It must not overwrite the canonical authority workbook or send email.

4. Record visible, non-sensitive source coverage in the review workbook: selected import mode, requested month, expected institutions, per-institution availability/coverage status, and whether the import is complete. Do not include connection tokens, account identifiers, raw payloads, full source descriptions, or row-level financial data in coverage metadata or telemetry. If any expected institution is unavailable or incompletely covered, label the workbook incomplete and name the institution at a high level.
5. Validate the review workbook before handoff: typed dates and amounts, source transaction count, preserved Category values, one authoritative dropdown covering every transaction row, visible Budget Treatment and Treatment Basis, auditable reconciliation rows, tables, formula-driven Budget Summary Actual/Remaining/Variance and Overall Total, readable wrapped Purchased Item Detail, valid Gmail hyperlinks, and zero broken formulas. Confirm that only confidently excluded non-spend rows are omitted from Actuals and that treatment-review rows remain visible and counted.
6. Open the final review workbook in Excel when requested or when visual verification is needed. Report aggregate counts and source coverage only.

## Optional Purchase Evidence

Workbook creation does not depend on Amazon or Gmail access. Without verified evidence, keep a concise `Review needed` note.

When the user authorized purchase research and access is already available:

- Use Amazon order history through the user's existing signed-in Chrome session and Gmail through an available read-only connector. Search and read only.
- Match conservatively on transaction date, amount, merchant/payee, and receipt/order language. Never infer or invent items.
- For confirmed Gmail evidence, retain itemized confirmation detail and the actual connector-returned message/thread identifier so the workbook can create the direct Gmail link.
- For ambiguous, conflicting, or unmatched evidence, keep `Review needed` and no link.
- Put collected evidence in a local JSON manifest following [references/review_evidence_contract.md](references/review_evidence_contract.md), then rebuild only the review workbook:

```bash
PYTHONPATH=src <python> -m personal_budget_plaid_excel.cli.main \
  --workbook /absolute/path/to/budget-YYYY-MM.xlsx \
  review-run --month YYYY-MM --output-dir /absolute/review/output/dir \
  --evidence-json /absolute/path/to/local-evidence.json
```

If a password, MFA, CAPTCHA, sign-in, or unexpected permission prompt appears, stop that source and ask the user to take over. If access is unavailable, finish the workbook with `Review needed`; do not loop or block the budget update.

## Authority and Preservation

- `Manual Budget` is the category and budget authority. Never edit it during a monthly refresh.
- Preserve every existing user-selected review Category and supported Purchased Item Detail on reruns. Append unseen transactions once; leave new categories blank.
- Preserve Gmail links and any manual Category even when a row is classified as non-spend. The treatment controls summary inclusion without erasing the user's choice.
- The Category dropdown must use all Manual Budget categories via the review workbook's `BudgetCategoryList` named range. Category and Budget Treatment together drive Budget Summary Actual totals.
- Never auto-categorize a row whose Budget Treatment is `Excluded from Budget` or `Review needed`. Ambiguity is a review state, not permission to suppress the row.
- Back up an existing month staging workbook and review workbook before replacement. Never overwrite the canonical authority workbook with a review artifact.
- Preserve prior reconciliation entries and their human disposition on idempotent rebuilds. Never silently promote a portal-only, Plaid-removed, or ambiguous row into spending; leave it visible for review until the user resolves it.

## Safety Boundaries

- Never put Plaid tokens/secrets, full account IDs, raw transaction payloads, provider message bodies, merchants, amounts, receipt contents, Gmail IDs/URLs, item details, or rule-input descriptions in chat or telemetry. Treatment telemetry is aggregate counts/status only.
- Never send, draft, label, archive, or modify Gmail; never purchase, return, cancel, or change Amazon settings; never make payments.
- Do not run legacy `run`, generate/send HTML email, or use a Gmail send tool unless the user separately and explicitly asks for that exact action.
- New review-workflow tests use disposable synthetic workbooks, fake Plaid sync JSON, synthetic portal manifests, and synthetic evidence manifests only. Never invoke live Plaid, bank portals, Gmail, or Amazon during tests.

## Connection Setup and Tests

If no Plaid item is connected and the user wants to connect one, use `connect`; credentials and tokens stay local and Keychain-backed. The localhost listener must remain available for the number of connections requested.

```bash
PYTHONPATH=src <python> -m personal_budget_plaid_excel.cli.main connect --connections N
PYTHONPATH=src <python> -m unittest discover -s tests -v
```
