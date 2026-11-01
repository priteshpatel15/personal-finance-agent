# Local foundation and privacy controls

**Feature ID:** `FEATURE-FOUNDATION`

The owner starts a local runtime that presents a locked, local dashboard. It stores financial facts in SQLite and read credentials only in Keychain. It has no public listener, no payment capability, and no background task until the owner enables it. The dashboard shows storage health, latest snapshot time, scheduler state, and a recoverable error state.

Owner controls: unlock local data, enable/disable the daily job, inspect redacted telemetry, and clear a failed connection setup. Keyboard navigation, visible focus, labels, contrast, and empty/error/loading states are mandatory. Telemetry records the operation and outcome but hashes account and merchant identifiers and never records tokens, raw transcripts, or dollar amounts.
