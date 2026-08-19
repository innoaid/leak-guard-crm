---
name: n8n API key in CLAUDE.md is truncated
description: Live n8n API key in CLAUDE.md is only the JWT header prefix; full key required for n8n /api/v1/* calls is not stored
type: project
originSessionId: d2026ad9-be64-45da-b7bd-5e342890e792
---
CLAUDE.md lists the n8n cloud API key as `eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9...` — only the JWT header prefix. The remainder of the key is not in CLAUDE.md or any other readable file. Authenticated calls to `https://leakguard.app.n8n.cloud/api/v1/*` (executions, workflows, etc.) cannot be made without it.

**Why:** The key is sensitive and was deliberately stored partial in CLAUDE.md as a soft-redaction.

**How to apply:** When the user asks for live n8n diagnostics (execution traces, workflow JSON, recent runs), ask them to either:
1. Paste the specific execution data themselves into chat (preferred — no key exposure), or
2. Paste the full API key in chat (sensitive — recommend rotating after the session)

Public, unauthenticated webhooks still work fine without the API key:
- `GET /webhook/lg-availability` — returns slot availability
- `POST /webhook/lg-booking` — submits a booking
- `POST /webhook/lg-wa-group-bot` — bot entry point (do not POST fake events to production)
