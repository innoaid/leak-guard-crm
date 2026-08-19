---
name: MYT timezone double-shift in n8n Code nodes
description: `new Date(now.getTime() + 8*60*60*1000)` is correct for extracting MYT date PARTS (getUTC*, slice 0-10) but BUGGY when fed back to .toISOString() as a full timestamp — it double-encodes as UTC, putting the value 8h in the future.
type: feedback
---

n8n Code nodes use `const myt = new Date(now.getTime() + 8*60*60*1000)` as a
common idiom for "get current time in Malaysia". This is correct for some
uses, fatal for others. The trap:

**SAFE uses (extracting MYT calendar parts):**
```js
myt.getUTCFullYear()                  // ✓ MYT year
myt.getUTCMonth()                     // ✓ MYT month (0-11)
myt.getUTCDate()                      // ✓ MYT day-of-month
myt.getUTCDay()                       // ✓ MYT day-of-week
myt.getUTCHours()                     // ✓ MYT hour
myt.toISOString().slice(0,10)         // ✓ MYT YYYY-MM-DD
myt.toISOString().split('T')[0]       // ✓ MYT date string
new Date(myt.getTime() + d*86400000)  // ✓ MYT future-day offset
                                       // (when followed by getUTC* or .slice(0,10))
```

**BUGGY use (timestamp re-encoding):**
```js
const timeMin = myt.toISOString();    // ✗ NOW 8h in the future
                                       // (stored/sent as a UTC timestamp)
```

**Why:** `now.toISOString()` is already UTC. Adding 8h to the underlying
epoch ms creates a Date whose `.getUTCHours()` etc. return MYT-aligned
values — useful for *parts*. But its `.toISOString()` outputs the same epoch
ms as a *UTC timestamp* literal. Reading 09:08 UTC vs 09:08 MYT downstream
results in an 8h offset.

**Failure mode (round 34, 2026-05-01):** LG-Availability `Prep Time Range`
sent `timeMin = myt.toISOString()` to Google Calendar API. Result: timeMin
was 8h in the future. Calendar query missed every morning event. Booking
page + bot showed all today slots free even when blocked. SV-0426-028 hit
this when admin saw "all slots available" despite a real morning block.

**Rule:** when you need the **time-of-day in UTC for an external API or
database write**, use `now.toISOString()` directly (or `Date.now()`). When
you need to **extract MYT date parts for display/comparison**, use the
shifted `myt` followed by `getUTC*` or `.slice(0,10)`. Never feed shifted
`myt` into `.toISOString()` and treat the output as a real timestamp.

**Audit pattern:** to find this bug class, grep all active n8n Code-node
jsCode for `\+\s*8\s*\*\s*60\s*\*\s*60\s*\*\s*1000` and verify each match
is followed by `getUTC*` / `.slice(0,10)` / `.split('T')[0]`, not by
`.toISOString()` used as a timestamp downstream. Round-34 sweep across 26
active workflows found 13 matches; 12 were safe (date-part extraction),
1 was buggy (Prep Time Range).
