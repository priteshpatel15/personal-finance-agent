# System Design

## Architecture

The local runtime uses Clean Architecture with dependencies directed inward:

1. **Domain (`packages/finance-domain`)** owns money, budget-period, inclusion, category-bridge, confidence, and affordability policy. It is deterministic and has no I/O dependencies.
2. **Application (`packages/finance-application`)** exposes `ImportBudgetWorkbook`, `LinkPlaidAccount`, `SyncTransactions`, `ClassifyTransaction`, `ReviewClassification`, `GetBudgetPosition`, and `AnswerPurchaseQuestion` use cases through ports.
3. **Ports (`packages/finance-contracts`)** specify typed repositories, encrypted-secret store, clock, job scheduler, Plaid sync, workbook reader, classifier, transcript source, telemetry, and transaction boundary contracts.
4. **Adapters (`packages/finance-infrastructure`)** own SQLite/Prisma migrations, the xlsx reader, Plaid Link/Transactions Sync, macOS Keychain, local classifier, launchd, JSONL telemetry, and testing fakes.
5. **Delivery (`apps/finance-web`)** presents authenticated local routes and a Codex command adapter. It calls application use cases only; UI components never query SQLite or Plaid.

No adapter is imported by the domain. Dependency tests enforce this direction.

## Feature impact

| Feature | Outcome | UI | Database | Authorization-sensitive |
| --- | --- | --- | --- | --- |
| `FEATURE-FOUNDATION` | private local ledger and operating controls | Yes | Yes | Yes |
| `FEATURE-WORKBOOK` | protected budget/history import | Yes | Yes | Yes |
| `FEATURE-SYNC` | Plaid read-only daily snapshots | Yes | Yes | Yes |
| `FEATURE-CATEGORIZATION` | trustworthy automatic categorization and review | Yes | Yes | Yes |
| `FEATURE-ADVICE` | explainable budget position and purchase advice | Yes | Yes | Yes |
| `FEATURE-REMOTE-VOICE` | Codex command and responsive voice-entry experience | Yes | Yes | Yes |

## Interfaces and compatibility

`BudgetRepository`, `LedgerRepository`, and `SnapshotRepository` publish opaque IDs, optimistic versions, and atomic write results. `PlaidGateway` supports Link-token creation, public-token exchange, and cursor-based sync only. `WorkbookReader` emits a signed import preview rather than raw cells. `CategoryClassifier` returns ranked, bridged categories plus a confidence and reasons. `QuestionAnswerer` returns a calculation object and a brief natural-language explanation. Every port accepts `OperationContext` carrying authorization, idempotency, trace, and telemetry fields.

The local HTTP contract is versioned under `/api/v1`; the Codex command is a thin local adapter over the same application input and never obtains a distinct privilege. A public-contract change requires a material amendment before implementation.

## Persistence and transactions

SQLite records accounts, Plaid items (without tokens), normalized transaction versions, category assignments and owner rules, budget plans and bridge mappings, sync cursors, immutable snapshots, review decisions, and redacted telemetry. Plaid access tokens live only in Keychain. A successful sync commits transaction changes and one immutable snapshot atomically; a failed sync retains the prior successful snapshot. The original workbook is read-only and its SHA-256 is captured in the import record.

Every database-affecting test starts from a fresh copy of a migrated real SQLite fixture, asserts migration and rollback behavior, persists and reopens state, and proves the source fixture and source workbook hashes are unchanged.

## UI and end-to-end testing

Affected user journeys compile into an exhaustive Playwright workflow manifest.
The final workflow matrix covers success, error, recovery, keyboard access,
responsive behavior, stale-version conflict, restart behavior, stale snapshot,
and interrupted sync states against one immutable release identity.

## Telemetry

Transactional operations emit redacted JSONL telemetry; the implemented contract is
`src/personal_budget_plaid_excel/infrastructure/telemetry.py`.
