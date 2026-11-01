# DESIGN-ARCHITECTURE: Local Clean Architecture runtime

- Status: Accepted
- Requirements: `REQ-FOUNDATION`, `REQ-WORKBOOK`, `REQ-SYNC`, `REQ-CATEGORIZATION`, `REQ-ADVICE`, `REQ-REMOTE-VOICE`

## Decision

Use a pnpm TypeScript workspace with Next.js only at the delivery boundary, pure domain and application packages, typed ports, and adapters for all I/O. SQLite is the sole local ledger; macOS Keychain holds secrets. The Codex command and web API invoke the same use cases.

## Alternatives

- A framework-only application would couple budget rules to UI and Plaid handling, making deterministic proof difficult.
- A cloud-first database or SaaS job runner would violate the stated local-first requirement.

## Consequences

Dependency tests enforce inward imports. Adapters can be replaced without rewriting affordability policy. The initial project has no code, so the stack is an explicit approved target rather than a repository-discovered fact.
