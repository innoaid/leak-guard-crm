---
name: n8n HTTP nodes neverError silently fails open
description: HTTP Request nodes with `options.response.response.neverError: true` swallow upstream API failures. Downstream code that defaults to empty arrays / falsy values cannot distinguish "API broke" from "no data". Always disable neverError, OR explicitly detect-and-fail-closed in the next Code node.
type: feedback
---

n8n's HTTP Request node has an option `options.response.response.neverError`.
When `true`, the node treats any HTTP response (including 4xx/5xx) as a
"successful" execution and passes whatever payload arrived to the next node.

**Failure mode:** Calendar/API calls that should determine "is this slot
free / does this resource exist" use this flag to keep workflows resilient.
The next node typically has:
```js
const items = Array.isArray(resp.items) ? resp.items : [];
const isFree = !items.some(...);
```
If the upstream API returned an error (rate-limited, 503, auth-token
expired mid-request), `resp` has no `items` array → `items = []` → `isFree
= true` → **the workflow proceeds AS IF the calendar were genuinely empty**.

**Concrete history (round 35, 2026-05-01):** All 7 calendar HTTP nodes in
LG-Availability + LG-Booking had `neverError: true`. A Google Calendar API
blip would silently make every slot look bookable on every check, and
LG-Booking's pre-create slot check would also pass empty → **double-booking
guaranteed during any API outage**.

**Fix policy (round 35):** Keep `neverError: true` (so n8n doesn't crash on
transient blips), but add explicit detect-and-fail-closed at the start of
the next Code node:

```js
const _errored = !!resp.error || !Array.isArray(resp.items);
if (_errored) {
  return [{ json: {
    error: 'calendar_unavailable',
    message: 'Service temporarily unavailable. Please try again in 1 minute.',
    availability: {},  // or whatever the empty-but-explicit shape is
  } }];
}
```

Caller code (booking page, chat agent, kanban) MUST also handle the
`error: 'calendar_unavailable'` response shape — show a clear retry message
instead of "no slots available" silence.

**Audit pattern:** to find this bug class, scan all active workflows for
HTTP nodes with `neverError: true`. For each, check the next Code node
that consumes its response and confirm explicit error detection exists.
Don't trust the existence of `Array.isArray(...) ? ... : []` fallbacks
— those LOOK defensive but actually paper over real failures.

**When to actually disable neverError instead of detecting-and-failing-closed:**
- When the upstream call is purely informational (analytics logging, sentry
  notifications) and a missing field is genuinely fine.
- When the workflow has a Wait + Retry node downstream specifically wired
  for transient errors.
- Otherwise: keep neverError=true AND fail-closed in code. The single-source
  defensive layer is the bug; you need TWO independent safeguards.
