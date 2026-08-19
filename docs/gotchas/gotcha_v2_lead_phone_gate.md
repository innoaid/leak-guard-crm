---
name: v2 must verify sender phone = lead phone, not just non-staff
description: WA Receiver routes ANY non-staff group message to v2. v2 Verify Lead matches the lead by groupId, not by sender phone — so messages from inspector / spouse / family / 3rd party in customer's group all triggered bot replies meant for the customer. Fix is in v2 Verify Lead: check senderPhone matches lead.Phone after lookup.
type: feedback
---

**Failure mode (round 37, 2026-05-01):** SV-0426-009 Amar group, exec 6337.
Field inspector (`601121806801`, not in staff allowlist) messaged
"datang dalam pukul 2.30pm@3.00pm boleh?" — they were asking the customer if
arriving 30 min early was OK. Bot misclassified as customer reschedule
intent (msgText matched BOOK_RX time-of-day regex, status was post_booking)
and replied with full week's available slots.

**Root cause:** WA Receiver `Parse & Route` rule:
```js
if (isGroup && !isStaff && !fromMe) → customer_message_v2
```
Treats EVERY non-staff group participant as "the customer". v2 Verify Lead
then matches the lead by `groupId === Group ID (AB)`, NOT by sender phone.
So spouse / family / inspector / mistakenly-added 3rd parties all trigger
bot responses meant for the lead.

**Fix shipped (round 37):** in v2 Verify Lead, immediately after lead lookup,
compare senderPhone to lead.Phone with existing normalization patterns:
```js
const _normSender = String(senderPhone || '').replace(/\D/g, '');
const _normLead   = String(lead.json['Phone'] || '').replace(/\D/g, '');
const _phoneMatches = !!_normSender && !!_normLead && (
  _normSender === _normLead
  || (_normSender.length >= 9 && _normLead.length >= 9
      && (_normSender.endsWith(_normLead.slice(-9))
          || _normLead.endsWith(_normSender.slice(-9))))
);
if (!_phoneMatches) {
  return [{ json: { shouldRespond: false, skipReason: 'non-lead sender (...)', ... } }];
}
```

Place AFTER bot loop check + lead lookup, BEFORE escalation regex / send-link
fast path / silent_ack / message classification.

**How to apply downstream:**
- ANY future fix or new feature in Verify Lead can assume "if we got past the
  gate, sender is the lead". No need to re-check senderPhone in downstream
  branches.
- DO NOT loosen this filter to handle "staff reply in group". Staff phones
  are already filtered upstream by WA Receiver's `isStaff` check (rule #11
  for -admin commands; non-staff filter on rule #13).
- If you ever want bot to reply to a non-lead sender (e.g. spouse), the
  correct path is to add their phone to a NEW lead row in CRM and link the
  group to that row — NOT to relax this gate.

**Edge cases handled:**
- Reactions converted to text by Parse & Route — senderPhone is the reactor.
  Reactor must still match lead phone, otherwise silent skip. Correct.
- Lead row's Phone column blank/wrong — gate rejects all messages → bot
  silent → admin should fix CRM. Acceptable failure mode.
- Phone format variants (0xxx vs 60xxx) — suffix-9 matching handles.

**Related work:** the audit `audit_funnel_v1.md` (round 38) already noted
this as crash-scenario A1 ("Customer DMs bot routed to v2 chat agent"). Fix
shipped at round 37 covers BOTH the in-group inspector case AND the future
funnel "customer DMs bot for group request" path — the latter will route
through a SEPARATE workflow, never v2.
