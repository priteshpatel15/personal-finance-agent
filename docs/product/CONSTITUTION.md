# Product Constitution

## Purpose and authority

This constitution is the highest product authority. It captures the purpose,
intended users, durable desires, experience promises, engineering principles,
and rules for changing them. Specifications interpret it; tasks never override
it.

## CONST-PURPOSE: Purpose

The product exists to give its owner a trustworthy daily answer to "where am I in my budget?" and "can I afford this purchase?" using locally held budget rules, current read-only transaction data, and clearly qualified recommendations.

## CONST-USER-VALUE: User value

The owner needs to understand category and overall monthly remaining budget before spending, without spreadsheet work or unexplained automation. Every answer identifies its data-as-of time, category assumption, current and proposed totals, and uncertainty.

## CONST-EXPERIENCE: Experience promise

The owner can ask a typed or transcribed voice question in seconds, review the calculation, correct any category, and recover from stale data, an incomplete connection, or a failed sync without losing a last-known-good budget view.

## CONST-LOCAL-PRIVACY: Local-first privacy and authority

Financial records, category history, snapshots, and Plaid access tokens remain on the owner's machine. Plaid access is read-only; the product never initiates transfers, payments, trades, or account changes. Account linking, re-linking, budget import, manual refresh, and category-rule changes require a visible owner action.

## CONST-TRUST: Truthful budget decisions

The workbook's `Manual Budget` sheet is the starting budget authority. Posted transaction facts, pending transactions, imported-history labels, predictions, owner corrections, and recommendations remain distinguishable. The product never silently turns a prediction into a user-approved category or presents stale data as current.

## CONST-ENGINEERING: Engineering integrity

Use Clean Architecture, decoupled modules, SOLID design, test-driven delivery, secure defaults, and complete telemetry.

## CONST-VERIFICATION: Proof before completion

A requirement is complete only when fresh release-bound evidence proves its acceptance criteria.

## Intended users and durable desires

- **Owner:** connect personal financial accounts once, see daily category position, and ask an immediate pre-purchase affordability question from desktop or mobile through the local Codex-operated workflow.
- **Owner as reviewer:** correct uncertain categories once so later matching improves without rebuilding the budget or corrupting historical evidence.
- **Owner as privacy steward:** retain exclusive authority over data, credentials, imports, refreshes, and any future external sharing.

## Governance

One explicit approval binds the exact design-package hashes and baseline commit.
A later approval is required only for a material change to a constitutional
clause, observable behavior, architecture, public contract, or permission
boundary. No checklist, task state, implementation, or evidence may weaken this
authority.

## Completion

Completion means locally verified against one immutable release identity. It
does not mean pushed, deployed, shipped, or released to users.
