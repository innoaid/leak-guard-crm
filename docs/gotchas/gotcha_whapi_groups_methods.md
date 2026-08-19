---
name: Whapi groups events — PUT carries participants, PATCH carries metadata
description: Whapi event-method mapping is counter-intuitive; keep groups PUT for join detection, not PATCH
type: project
originSessionId: acc90124-d214-4973-b213-a7cdc5dbb905
---
When subscribing to Whapi webhook events for the Leak Guard bot:

- `groups PUT` event → body has **`groups_participants`** array (add / remove / promote / demote / join / leave). This is what `LG - WA Receiver` parses for new-customer detection.
- `groups PATCH` event → body has **`groups_updates`** (subject/photo/description metadata changes). Receiver doesn't parse this and routes to skip.
- `groups POST` event → group created (rarely fires; not used).

**Bug history (2026-04-26):** I trimmed Whapi events on the assumption that `groups PATCH` carried participant-add events. Wrong — Whapi sends those via `groups PUT`. After trim, all direct-add joins silently failed for ~3 hours (Muru, Annie). Real evidence in n8n exec metadata showed `event: {type:'groups', event:'put'}` for participant payloads.

**How to apply:**
- Minimal Whapi subscription needed: `messages POST` + `groups PUT`. That's it.
- If welcome msgs ever stop firing on direct-add, FIRST check Whapi `GET /settings` confirms `groups PUT` is subscribed. If `PATCH` is there instead, that's the bug.
- The `event_meta` field in the Whapi webhook body (`event: {type, event}`) is the authoritative source for which method fired. Useful for debugging which trim is correct.
