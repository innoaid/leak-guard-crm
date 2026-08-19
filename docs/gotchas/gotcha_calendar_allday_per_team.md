---
name: All-day events block per-calendar, NOT globally
description: LG/LGJB team calendars use all-day events as TRUE BLACKOUTS (public holidays, inspector leave) → block. Alvin's personal calendar uses all-day events as JOB-TRACKING NOTES → skip. Same per-calendar rule must apply in both lg-availability and lg-booking.
type: feedback
---

The booking pipeline reads 3 calendars (Team KL = LG + Alvin combined; Team JB
= LGJB only). All-day events on these calendars have **different semantics**:

**TRUE BLACKOUTS (block):**
- LG calendar (`bd98920b3c826b87f01e302106376ec1cbb64656847cd053fe724dbb0759546f@group.calendar.google.com`)
- LGJB calendar (`3e5d42664b8acdee8ca66ad466bfd7932df052778bd41ccd67199bcca99e1e45@group.calendar.google.com`)
- Use case: "Public Holiday — closed", "All inspectors on leave 5-7 May"

**JOB-TRACKING NOTES (do NOT block):**
- Alvin's personal calendar (`alvinlai.aid@gmail.com`)
- Use case: "Riza Sri Hartamas" (Alvin's other job today), "KS Lim Tmn Desa"
- Alvin can still squeeze in a site visit between/around these tasks. The
  all-day marker just helps him track context for the day.

**Implementation rule:**
```js
const _isAllDay = ev => !!(ev.start && ev.start.date && !ev.start.dateTime);
const alvinTimedOnly = alvinEvents.filter(ev => !_isAllDay(ev));
const teamEvents = team === 'JB'
  ? lgjbEvents.slice()                       // all-day blackouts kept
  : [...lgEvents, ...alvinTimedOnly];        // LG all-day kept; Alvin all-day dropped
```

Apply this in BOTH:
- `LG - Availability` (`VfAulssOzbmeoogV`) → `Build Availability`
- `LG - Booking` (`gh0pwqGygBDoNzJB`) → `Evaluate Slot`

**Failure mode (if forgotten):** If you treat all-day events uniformly, EITHER:
- Always block all-day → Alvin's job-tracking notes wipe out entire weeks of
  bookable slots, customers can't reach team (round 35.1 hit this for ~30 min
  before fix).
- Always skip all-day → public holiday blackouts don't block, customers booked
  for "closed" days, no inspector available (the original bug).

**Team-calendar policy (admin convention):**
- For TRUE blackouts: create the all-day event on the LG calendar (KL) or LGJB
  calendar (JB), NOT on a personal calendar.
- For job-tracking notes (Alvin's existing usage): keep on Alvin's personal
  calendar, will be ignored by the booking flow.

**If a new staff member is added with a calendar:**
- Decide upfront whether their all-day events are blackouts or notes.
- If notes: filter all-day from their event stream same as Alvin.
- If blackouts: include all-day events.
- Ideally formalize via a "Staff Calendars" CRM tab with a `treat_allday_as`
  column ('blackout' | 'note').
