# DESIGN-CATEGORIZATION: Local supervised categorization with owner control

- Status: Accepted
- Requirements: `REQ-CATEGORIZATION`, `REQ-ADVICE`

## Decision

Classify locally using a precedence-ordered merchant-rule layer plus a supervised model trained on imported labels. Return confidence and reason codes. Only high-confidence classifications affect category totals automatically; everything else remains reviewable. Owner corrections override predictions and become durable rules.

## Alternatives

- A remote LLM-only categorizer sends financial data away and produces unbounded output.
- Automatically accepting all predictions risks a misleading budget.

## Consequences

The model is an adapter with no authority to alter a budget plan. The review queue is a first-class journey, and all prediction/version data are auditable.
