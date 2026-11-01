# Codex command and voice-entry experience

**Feature ID:** `FEATURE-REMOTE-VOICE`

The owner may call the local `finance ask` command through Codex Remote with a typed question or a client-provided transcript. The command produces the same calculation object and explanation as the web surface, with the snapshot timestamp and no elevated access.

The responsive local `/ask` route also supports accessible microphone capture, visible recording/transcribing/error states, transcript review and editing before submission, and keyboard-equivalent typed entry. It does not assume a phone can reach a machine-local web server. Remote browser networking, automatic speech provider upload, and any shared link are out of scope until separately approved.
