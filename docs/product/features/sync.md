# Plaid linking, daily sync, and snapshots

**Feature ID:** `FEATURE-SYNC`

The owner starts Plaid Link from the dashboard and explicitly authorizes a read-only connection. The product exchanges and stores the credential locally, obtains transaction history through cursor-based sync, deduplicates transaction versions, and atomically writes an immutable snapshot. It schedules one daily 06:00 local sync and supports an owner-initiated refresh.

The dashboard distinguishes posted from pending activity and shows the as-of timestamp. A Link exit, expired consent, network failure, or partial sync keeps the last successful snapshot and offers re-link or retry; it never claims the data are current. The product never creates transfers, payments, trades, or account changes.
