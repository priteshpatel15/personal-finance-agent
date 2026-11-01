# DESIGN-REMOTE-VOICE: Codex command with explicit mobile and voice boundaries

- Status: Accepted
- Requirements: `REQ-REMOTE-VOICE`, `REQ-ADVICE`

## Decision

Expose an owner-only `finance ask` local command that Codex Remote can invoke, and keep it behaviorally identical to the web question API. Add responsive web voice capture with transcript confirmation. Do not create a public server, shared URL, or automatic cloud speech upload.

## Alternatives

- Exposing the local database to the internet for phone access would create a new authentication and network boundary.
- Sending raw audio to a provider without explicit consent violates the local-first promise.

## Consequences

The owner gets remote command access while the machine is running. Native phone networking or provider-backed transcription is a future amendment, not an implied delivery claim.
