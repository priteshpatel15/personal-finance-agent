# personal-finance-agent

A local-first budget pipeline: pull bank and card transactions through Plaid, reconcile them against an Excel workbook you already own, and build a monthly review workbook, all without your data leaving the machine.

## Why

Budgeting apps want your bank credentials and your data on their servers, and they replace the spreadsheet you trust with their categories. This keeps the spreadsheet as the source of truth: your `Manual Budget` sheet defines the categories and amounts, Plaid supplies read-only transactions, and the tool only ever writes separate staging and review workbooks beside it.

## Highlights

- **Plaid REST client with no SDK.** `infrastructure/plaid_rest.py` is under 200 lines of `urllib`: Link token creation, public-token exchange and cursor-based `/transactions/sync`, with an injectable transport so tests never touch the network. A localhost Link page (`connect`) handles one or several institutions in a single session.
- **Secrets in the macOS Keychain.** Plaid client credentials and per-item access tokens are stored with `security add-generic-password` under the `personal-budget-plaid-excel` service. The workbook keeps only a key name (`plaid:<item_id>`), never a token, and environment variables can override for CI or sandbox use.
- **Non-spend and transfer detection.** `domain/non_spend.py` classifies credit-card payments, account transfers, refunds, income and budgeted loan payments conservatively: it uses Plaid category codes, account types and matching opposite-sign legs across accounts, refuses to exclude a purchase just because the merchant name contains "payment" or "transfer", and sends ambiguous rows to `Review needed` instead of guessing.
- **Excel workbook as the source of truth.** Category order, spelling and budgets come only from `Manual Budget`. Reruns are idempotent: they append unseen transactions once, keep every category you chose by hand, back up before replacing anything, and never overwrite the authority workbook.
- **Two import sources, one review builder.** You can import from Plaid or from a local manifest of posted transactions read from signed-in bank portals. Both feed the same review workbook, which records per-institution coverage and flags the month incomplete when a source is missing.
- **Redacted telemetry.** The JSONL event log strips tokens, account IDs, merchants, amounts and provider error text, keeping aggregate counts and error codes only.

## Quickstart

Requires Python 3.11+. Plaid sync and Keychain storage need a Plaid account and macOS. Everything below runs offline on a synthetic workbook.

```bash
git clone <repository-url> personal-finance-agent && cd personal-finance-agent
python3 -m venv .venv && source .venv/bin/activate
pip install -e .

# Build a synthetic authority workbook (fake categories and transactions)
python -c "import sys, pathlib; sys.path.insert(0, 'tests'); from synthetic_workbook import write_synthetic_workbook; write_synthetic_workbook(pathlib.Path('budget.xlsx'))"

personal-budget-plaid-excel doctor --month 2026-05                      # validate workbook shape and config
personal-budget-plaid-excel report --month 2026-05 --output-dir reports  # HTML budget-vs-actual report
personal-budget-plaid-excel review-run --month 2026-05 --output-dir review  # standalone review workbook
```

With real data, point `--workbook` (or `BUDGET_WORKBOOK`) at your own workbook, store Plaid credentials (see `.env.example`), then:

```bash
personal-budget-plaid-excel connect --connections 2       # localhost Plaid Link; tokens go to Keychain
personal-budget-plaid-excel monthly-run --month 2026-08 --source plaid
```

The workbook needs a `Manual Budget` sheet (Cost Type / Category / Budget in columns A-C, ending with a `Budget` total row) and an `Actual Charges` sheet with headers on row 5. `tests/synthetic_workbook.py` shows the minimal shape.

## Architecture

Hexagonal layout under `src/personal_budget_plaid_excel/`:

- `domain/` holds pure policy: budgets, the category bridge and suggestions, non-spend treatment, reconciliation, review rows and source-coverage rules. It does no I/O.
- `application/` contains the use cases (sync, monthly review, report, import-source selection, connection-state reconciliation), written against ports.
- `ports/` defines the Plaid, secrets and review-source interfaces as Protocols.
- `infrastructure/` holds the adapters: the Plaid REST client, macOS Keychain, the openpyxl workbook and review-workbook writers, the portal manifest reader, a connection registry, private (0600) file helpers and telemetry.
- `cli/` is an argparse entry point (`doctor`, `connect`, `sync`, `report`, `monthly-run`, `review-run`) plus the localhost Link flow.

`codex_skills/` is an agent skill that drives the CLI for a monthly review. `docs/product/` is the original design package; it describes a broader target stack (web UI, SQLite, voice questions) than the Python CLI implemented here.

## Testing

```bash
python -m unittest discover -s tests      # or: pip install pytest && pytest -q
```

101 tests, all offline. They use synthetic workbooks, fake Plaid sync JSON and synthetic portal manifests, with no live Plaid, bank or email calls.

## License

MIT, see [LICENSE](LICENSE). Author: Pritesh Patel.
