# Direct Portal Import Mode

Use this mode only after the user explicitly chooses `Direct portal`.

## Access boundary

Before collection, tell the user to sign in to Bank of America, Chase, and American Express in Chrome. Use the `chrome:control-chrome` skill and the user's existing Chrome sessions only. This mode is read-only:

- Never enter credentials, handle MFA, solve a CAPTCHA, approve an unexpected permission request, download statements unnecessarily, change account settings, send messages, or make a payment.
- If a credential, sign-in, MFA, CAPTCHA, or permission prompt appears, stop that institution immediately and state the exact user takeover needed. Continue the other already accessible institutions instead of retrying in a loop.
- Never fall back to Plaid without a new user choice.

## Read and normalize

1. Open each institution's transaction activity and choose the requested month. Use a small, stated date range around both month boundaries when posting dates can straddle the month; do not broaden the read beyond what is needed for monthly attribution.
2. Collect only clearly posted transactions into a permission-restricted local portal manifest that follows [portal_manifest_contract.md](portal_manifest_contract.md). Treat pending rows as uncovered or review-only, not posted spend. Preserve enough local source evidence for conservative date, signed amount, institution, visible account label, description, and duplicate-occurrence matching, but never put that row-level content in chat or telemetry.
3. Let the manifest reader compute deterministic local keys when `local_id` is omitted. For genuinely distinct postings with otherwise identical fields, set stable `occurrence` values `1`, `2`, and so on; repeated pagination captures of the same posting reuse the same occurrence. Do not use or expose full account numbers, credentials, cookies, session data, or browser tokens.
4. For each institution, record `complete`, `partial`, or `unavailable`, the covered date range, and an aggregate posted-row count. A missing institution, partial pagination, an authentication handoff, or an uncertain date range makes overall Direct portal coverage incomplete.
5. Pass the local manifest to `monthly-run --source direct-portal --portal-json ...`. The command must not call Plaid in this mode.

```bash
PYTHONPATH=src <python> -m personal_budget_plaid_excel.cli.main \
  monthly-run --month YYYY-MM --source direct-portal \
  --portal-json /absolute/path/to/local-portal-manifest.json
```

## Review behavior

- Use the same review builder, Manual Budget authority, category dropdown, formulas, evidence preservation, payment/transfer treatment, and backup/validation workflow as Plaid mode.
- Record selected mode `Direct portal`, requested month, expected institutions, per-institution coverage/date range, and overall completeness in visible source coverage metadata. Do not claim the budget is complete when any source is incomplete or unavailable.
- Keep portal-only, ambiguous, and source-status discrepancies visible in the reconciliation queue. Do not silently add an unresolved discrepancy to spending or duplicate a transaction already represented by another stable row.
- Preserve existing evidence-backed Purchased Item Detail and Gmail links. Direct portal access does not authorize Amazon/Gmail research or any external writes.
- Keep the normalized manifest permission-restricted and local for an idempotent rebuild. Do not include its contents in reports, and delete it only when the user asks or a previously disclosed cleanup policy applies.
