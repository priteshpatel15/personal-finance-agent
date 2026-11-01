# DESIGN-TELEMETRY: Redacted telemetry and isolated proof

- Status: Accepted
- Requirements: `REQ-FOUNDATION`, `REQ-SYNC`, `REQ-CATEGORIZATION`, `REQ-ADVICE`

## Decision

Emit the complete program telemetry envelope for all operations, but use hashed identifiers and counts instead of raw financial values, merchants, transcripts, credentials, or Plaid payloads. Every database and UI claim is proven from a fresh copy of the migrated SQLite fixture with non-cacheable unit, integration, and Playwright runs.

## Alternatives

- Logging raw data would undermine privacy.
- Testing shared local state could hide migration and idempotency defects.

## Consequences

Telemetry supports audit and recovery without data leakage. Release evidence proves source fixtures and source workbook bytes were unchanged.
