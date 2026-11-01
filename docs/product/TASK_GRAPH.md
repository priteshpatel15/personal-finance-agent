# Compiled Task Graph

This is a deterministic preview compiled from the specification. It is not an
editable progress checklist. The development controller creates runtime task
state only after approval.

## TASK-FOUNDATION: Implement and qualify Local foundation and privacy controls

- Feature: `FEATURE-FOUNDATION`
- Milestone: `MILESTONE-FOUNDATION`
- Depends on: None
- Requirements: `REQ-FOUNDATION`
- Validation: `VALIDATE-FOUNDATION`
- Evidence contracts: `EVIDENCE-FOUNDATION`
- Packet: `PACKET-FOUNDATION`
- Write scope: `apps/finance-web/**`, `packages/finance-domain/**`, `packages/finance-application/**`, `packages/finance-contracts/**`, `packages/finance-infrastructure/**`, `tests/**`, `scripts/**`, `docs/operating/**`
- Protected paths: `budget.xlsx`, `.env*`, `.codex/**`, `docs/product/**`
- Logical locks: resource=['resource:foundation-local-ledger'], interface=['interface:foundation-ports'], schema=['schema:finance-sqlite']
- Globally exclusive: true

## TASK-SYNC: Implement and qualify Plaid read-only sync and daily snapshots

- Feature: `FEATURE-SYNC`
- Milestone: `MILESTONE-SYNC`
- Depends on: TASK-FOUNDATION
- Requirements: `REQ-SYNC`
- Validation: `VALIDATE-SYNC`
- Evidence contracts: `EVIDENCE-SYNC`
- Packet: `PACKET-SYNC`
- Write scope: `apps/finance-web/**`, `packages/finance-domain/**`, `packages/finance-application/**`, `packages/finance-contracts/**`, `packages/finance-infrastructure/**`, `tests/**`, `scripts/**`, `docs/operating/**`
- Protected paths: `budget.xlsx`, `.env*`, `.codex/**`, `docs/product/**`
- Logical locks: resource=['resource:sync-local-ledger'], interface=['interface:sync-ports'], schema=['schema:finance-sqlite']
- Globally exclusive: true

## TASK-WORKBOOK: Implement and qualify Workbook import and category authority

- Feature: `FEATURE-WORKBOOK`
- Milestone: `MILESTONE-WORKBOOK`
- Depends on: TASK-FOUNDATION
- Requirements: `REQ-WORKBOOK`
- Validation: `VALIDATE-WORKBOOK`
- Evidence contracts: `EVIDENCE-WORKBOOK`
- Packet: `PACKET-WORKBOOK`
- Write scope: `apps/finance-web/**`, `packages/finance-domain/**`, `packages/finance-application/**`, `packages/finance-contracts/**`, `packages/finance-infrastructure/**`, `tests/**`, `scripts/**`, `docs/operating/**`
- Protected paths: `budget.xlsx`, `.env*`, `.codex/**`, `docs/product/**`
- Logical locks: resource=['resource:workbook-local-ledger'], interface=['interface:workbook-ports'], schema=['schema:finance-sqlite']
- Globally exclusive: true

## TASK-CATEGORIZATION: Implement and qualify Automatic categorization and owner review

- Feature: `FEATURE-CATEGORIZATION`
- Milestone: `MILESTONE-CATEGORIZATION`
- Depends on: TASK-WORKBOOK, TASK-SYNC
- Requirements: `REQ-CATEGORIZATION`
- Validation: `VALIDATE-CATEGORIZATION`
- Evidence contracts: `EVIDENCE-CATEGORIZATION`
- Packet: `PACKET-CATEGORIZATION`
- Write scope: `apps/finance-web/**`, `packages/finance-domain/**`, `packages/finance-application/**`, `packages/finance-contracts/**`, `packages/finance-infrastructure/**`, `tests/**`, `scripts/**`, `docs/operating/**`
- Protected paths: `budget.xlsx`, `.env*`, `.codex/**`, `docs/product/**`
- Logical locks: resource=['resource:categorization-local-ledger'], interface=['interface:categorization-ports'], schema=['schema:finance-sqlite']
- Globally exclusive: true

## TASK-ADVICE: Implement and qualify Daily position and pre-purchase advice

- Feature: `FEATURE-ADVICE`
- Milestone: `MILESTONE-ADVICE`
- Depends on: TASK-CATEGORIZATION
- Requirements: `REQ-ADVICE`
- Validation: `VALIDATE-ADVICE`
- Evidence contracts: `EVIDENCE-ADVICE`
- Packet: `PACKET-ADVICE`
- Write scope: `apps/finance-web/**`, `packages/finance-domain/**`, `packages/finance-application/**`, `packages/finance-contracts/**`, `packages/finance-infrastructure/**`, `tests/**`, `scripts/**`, `docs/operating/**`
- Protected paths: `budget.xlsx`, `.env*`, `.codex/**`, `docs/product/**`
- Logical locks: resource=['resource:advice-local-ledger'], interface=['interface:advice-ports'], schema=['schema:finance-sqlite']
- Globally exclusive: true

## TASK-REMOTE-VOICE: Implement and qualify Codex command and voice-entry experience

- Feature: `FEATURE-REMOTE-VOICE`
- Milestone: `MILESTONE-REMOTE-VOICE`
- Depends on: TASK-ADVICE
- Requirements: `REQ-REMOTE-VOICE`
- Validation: `VALIDATE-REMOTE-VOICE`
- Evidence contracts: `EVIDENCE-REMOTE-VOICE`
- Packet: `PACKET-REMOTE-VOICE`
- Write scope: `apps/finance-web/**`, `packages/finance-domain/**`, `packages/finance-application/**`, `packages/finance-contracts/**`, `packages/finance-infrastructure/**`, `tests/**`, `scripts/**`, `docs/operating/**`
- Protected paths: `budget.xlsx`, `.env*`, `.codex/**`, `docs/product/**`
- Logical locks: resource=['resource:remote-voice-local-ledger'], interface=['interface:remote-voice-ports'], schema=['schema:finance-sqlite']
- Globally exclusive: true
