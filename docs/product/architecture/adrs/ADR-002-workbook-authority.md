# DESIGN-WORKBOOK: Workbook authority and explicit bridge

- Status: Accepted
- Requirements: `REQ-WORKBOOK`, `REQ-CATEGORIZATION`, `REQ-ADVICE`

## Decision

Treat the supplied workbook as read-only import evidence. `Manual Budget` rows 2-32 columns A-C define starting monthly categories and amounts; `Actual Charges` beginning at row 5 supplies history. Map the two source-only labels explicitly and retain `Work Reimbursement` outside budget spend.

## Alternatives

- Inferring category equivalence by similarity would silently change the owner's budget.
- Mutating the source workbook would make it impossible to prove the import did not damage evidence.

## Consequences

The import UI requires a preview and confirmation. Its hash, source ranges, bridge version, and evaluated formula values are persisted with the plan.
