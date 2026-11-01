# DESIGN-PLAID-SYNC: Read-only Plaid cursor sync and immutable snapshots

- Status: Accepted
- Requirements: `REQ-SYNC`, `REQ-ADVICE`

## Decision

Use Plaid Link and Transactions Sync with read-only scopes. Store the access token only in Keychain, persist only the item reference/cursor in SQLite, and write a new immutable snapshot atomically after each successful synchronization. launchd runs at 06:00 local time.

## Alternatives

- Screen scraping statements is brittle and does not meet the requested Plaid integration.
- Overwriting the current ledger on failure loses the last trustworthy answer.

## Consequences

Answers cite the snapshot timestamp and show stale state on failure. Linking, re-linking, and manual refresh are visible owner-authorized actions.
