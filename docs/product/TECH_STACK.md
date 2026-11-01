# Technical Stack

## Repository-discovered facts

| Concern | Observed repository truth |
| --- | --- |
| Languages | None observed; implementation is not yet initialized |
| Frameworks and test tools | None observed |
| Dependency manifests | None observed |
| Database markers | None observed |
| Telemetry markers | None observed |
| Budget source | `budget.xlsx`: `Manual Budget` and `Actual Charges` sheets |

## Approved engineering constraints

- High-level modules follow Clean Architecture dependency direction.
- Components are decoupled behind explicit ports and interfaces.
- Functions and classes follow SOLID design.
- Changes are developed test-first with requirement-linked RED and GREEN proof.
- UI changes require complete Playwright journey coverage.
- Database-affecting validation uses a fresh copy of the real database and
  proves that the source remained unchanged.
- Transactional operations emit redacted JSONL telemetry
  (`infrastructure/telemetry.py`).
- Runtime dependencies and toolchains are pinned for an active program.

## Approved implementation stack

- TypeScript on a pinned Node.js LTS runtime with pnpm workspaces.
- Next.js web application for the local, responsive dashboard and question surface.
- Pure TypeScript domain and application packages; adapters are the only layer that imports Plaid, SQLite, browser media, the filesystem, macOS Keychain, or framework APIs.
- SQLite with versioned migrations for locally persisted accounts, transactions, category decisions, budgets, and immutable sync snapshots. Tests clone a real migrated fixture database before every database-affecting run.
- Plaid Link and Transactions Sync behind a read-only adapter. Environment configuration contains no secrets; encrypted Plaid tokens live in macOS Keychain.
- A local, supervised categorization adapter trained from the imported labeled data; a deterministic rule layer and explicit owner corrections take precedence over a probabilistic prediction.
- `launchd` schedules daily refresh at 06:00 in the machine's configured local timezone, while a visible manual refresh remains available.
- Playwright covers Chromium desktop and a 390x844 mobile viewport. Vitest covers domain, application, adapter-contract, migration, and integration paths.

## Unresolved choices

Codex Remote is an owner-operated command channel, not a publicly exposed web-service guarantee. It invokes the same local question use case and accepts a typed question or a client-provided transcript. Native remote voice capture and phone-to-local browser networking are not assumed; a future network exposure or a cloud speech provider requires an amendment.
