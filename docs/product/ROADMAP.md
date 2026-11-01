# Product Roadmap

The roadmap is compiled from approved requirements and dependencies. Completion
state is projected by the development controller; it is not edited here.

| Milestone | Outcome | Features | Depends on |
| --- | --- | --- | --- |
| `MILESTONE-FOUNDATION` | Local foundation and privacy controls | Local foundation and privacy controls | None |
| `MILESTONE-SYNC` | Plaid read-only sync and daily snapshots | Plaid read-only sync and daily snapshots | MILESTONE-FOUNDATION |
| `MILESTONE-WORKBOOK` | Workbook import and category authority | Workbook import and category authority | MILESTONE-FOUNDATION |
| `MILESTONE-CATEGORIZATION` | Automatic categorization and owner review | Automatic categorization and owner review | MILESTONE-WORKBOOK, MILESTONE-SYNC |
| `MILESTONE-ADVICE` | Daily position and pre-purchase advice | Daily position and pre-purchase advice | MILESTONE-CATEGORIZATION |
| `MILESTONE-REMOTE-VOICE` | Codex command and voice-entry experience | Codex command and voice-entry experience | MILESTONE-ADVICE |

## Release sequencing

A milestone may advance only after its dependency milestones have locally
verified evidence. Local verification does not authorize push, PR, deployment,
production mutation, messaging, or paid infrastructure.
