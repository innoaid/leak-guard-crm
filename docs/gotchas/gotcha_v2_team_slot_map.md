---
name: v2 agent slot table is KL-only by default
description: AI Agent's system prompt has KL slot times (9/10:30/12/2/3:30 PM) hardcoded in examples scattered everywhere. JB customers get wrong slot labels unless Verify Lead injects a team-specific slot map at runtime.
type: feedback
---

The v2 chat agent (LG - Customer Chat v2, `xhO6U0Xa8VmGPevy`) has its system
prompt calibrated for **KL slot times** — 5 slots at 9/10:30/12/2/3:30 PM.
Examples like "Confirm 12:00 PM Monday 4 May ya?" / "Friday, 1 May 2026: -
9:00 AM - 10:30 AM…" appear throughout the prompt as illustration.

For **JB customers** (groups starting with `SVJB-` or `QTJB-`), the slot map
is different — only 2 slots per day:
  - slot 1 = 11:30 AM
  - slot 2 = 2:00 PM

`check_availability` correctly returns `[1, 2]` for JB dates. But when the
agent emits `[PROPOSE date slot 2]` and labels it in the chat text, it writes
"10:30 AM" — the **KL** label for slot 2 — because that's what its prompt
examples taught it.

**Failure mode (round 33, 2026-04-30):** SVJB-0426-024 Nithia asked for
"next tuesday at 11am". Agent replied `[PROPOSE 2026-05-12 slot 2]\nConfirm
10:30 AM Tuesday 12 May ya?`. Two bugs in one reply:
  1. Date was wrong (separate `next\s+\w+` over-match bug in Verify Lead)
  2. Slot label was "10:30 AM" but JB slot 2 = 2:00 PM. Confusing for customer.

Note: the actual BOOKING (when [BOOK ... slot 2] commits) goes through
LG-Booking which re-derives time from team prefix → 2 PM. So the calendar
event is correct, but the chat conversation misleads the customer in the
meantime ("bot says 10:30 AM, calendar says 2 PM").

**Fix shipped (round 33):** Verify Lead now detects team from `groupName`
prefix (`/^(SVJB|QTJB)-/i`) and injects an explicit slot map line into
`agentInput`:

  `Team: JB. Slot map (USE THESE EXACT TIMES — ignore any other slot examples
  in your system prompt): slot 1 = 11:30 AM, slot 2 = 2:00 PM (THIS TEAM HAS
  ONLY 2 SLOTS PER DAY — there is no 9 AM, 10:30 AM, 12 PM, or 3:30 PM for JB).`

This appears as a top line in the agent's user message every turn, overriding
the prompt's KL-biased examples.

**How to apply:** when adding new agent capabilities or when introducing
additional teams (KL2, JBPenang, etc.), ALWAYS inject the team's slot map
dynamically in the per-turn agentInput rather than relying on the system
prompt to remember conditional logic. The agent can't reliably switch slot
maps based on group prefix from prompt examples alone — it needs the
authoritative map handed to it each turn.

**Test for regression:** for any non-KL team customer, after `[PROPOSE]`,
verify the chat label time matches the team's slot map AND matches what
LG-Booking commits to the calendar.
