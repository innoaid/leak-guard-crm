---
name: n8n Cloud has a hard monthly execution limit
description: Aggressive webhook testing can exhaust the cap and silence the live bot until reset or plan upgrade
type: project
originSessionId: acc90124-d214-4973-b213-a7cdc5dbb905
---
The Leak Guard n8n Cloud plan has a monthly execution quota. When exceeded, ANY webhook trigger fails with:
`Execution limit reached. Consider upgrading your plan` — the workflow stops at the Webhook node, no downstream nodes run.

**Hit on 2026-04-25** during Phase A verification of the existing-appointment intent fix. Cause: ~30+ test reschedule webhooks fired across LG - Customer Chat, LG - Booking, plus the hourly LG - Follow Up cron, all in one session.

**Why:** Each test webhook spawns a full execution; verification queries via `/api/v1/executions?includeData=true` count toward read-side rate limits separately but the executions themselves are the cap.

**How to apply:**
- Before doing a long debugging session that fires many test webhooks, ask the user about their n8n plan and remaining quota for the month.
- Prefer reading existing executions (`GET /api/v1/executions`) over creating new ones for verification — diagnose from real customer activity when possible.
- Recommend pausing the hourly Follow-Up cron during heavy testing days to slow burn.
- When the limit hits: bot is silent across the board (not a code bug). User options: upgrade plan, wait for monthly reset, or pause schedulers.
- The execution-limit error is silent from the customer's perspective — webhook returns 200 (n8n responded), but no logic ran. Easy to misdiagnose as "bot ignoring messages" if you don't check exec status.
