# DESIGN-AFFORDABILITY: Explainable calendar-month affordability policy

- Status: Accepted
- Requirements: `REQ-ADVICE`

## Decision

Calculate category position by calendar month from personal posted spend, pending spend, accepted refunds, and approved categories. Exclude transfers, card payments, work reimbursement, and unresolved transactions from positive spending while exposing their value. Return Yes, Caution, or No using the documented freshness, confidence, and ten-percent-buffer rules.

## Alternatives

- A single global balance ignores category intent.
- Hiding pending or uncertain activity creates false certainty.

## Consequences

Every answer contains machine-readable inputs and a human-readable explanation. It cannot return a definitive approval with missing amount/category or stale data.
