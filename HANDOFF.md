# Three-task plan: (1) Mute v2 bot post-Quotation-Sent + (2) Fix kanban Send-reschedule-link URL + (3) FB ad attribution

User instruction: ship Task 1 + Task 2 now. Task 3 (FB attribution) is research/setup, deferred until they answer the prerequisites.

---

---

## Task 1 — Mute v2 chat bot for ALL phases after Quotation Sent

### Context

After Quotation Sent, every downstream phase is admin-handled — payment, installation scheduling, on-site work, balance, completion. v2's link-only bot has nothing useful to say there. Mute it.

### Phase strings (verified via gviz read of live CRM today)

Live CRM contains: `New Lead` (487), `Pending Balance` (27), `Site Visit Confirmed` (27), `Pending Site Visit` (23), `Quotation Sent` (23), `Pending Invitation` (12), `Pending QT` (7), `Pending I.Date` (3), `Rejected` (2), `Human Handoff` (1). Kanban code also recognises `Pending Downpayment`, `I.Date Confirmed`, `Job In Progress`, `Job Complete`, `Receipt Sent`.

**Phases to mute (all post-Quotation-Sent, in funnel order)**:
```js
const _LATE_STAGE_MUTED = [
  'Pending Downpayment',
  'Pending I.Date',
  'I.Date Confirmed',
  'Job In Progress',
  'Pending Balance',
  'Job Complete',
  'Receipt Sent',
];
```

Phases that STAY chat-active: `Pending Invitation`, `Pending Site Visit`, `Site Visit Confirmed`, `Pending QT`, `Quotation Sent`. The existing Round 43 archived-status mute keeps catching `Lost` / `Cold Lead` / `Rejected` / `Out of Area` / `Human Handoff`.

### Patch

Single-node patch on `LG - Customer Chat v2` → `Verify Lead`, identical pattern to Round 43:
- New `_LATE_STAGE_MUTED` const
- After the archived check, return `shouldRespond: false, skipReason: 'muted: <status>', needsAdminAlert: true` (different from archived: admin still gets DM since these customers are still active)

File: `C:/Users/CY Lee/AppData/Local/Temp/n8n_v2/patch_t17_mute_pending_balance.py` (already drafted; needs the broader 7-phase list).

---

## Task 2 — Fix kanban Send-reschedule-link URL (regression)

### Context

The kanban lead detail modal has a "Send reschedule link" button that calls Apps Script `handleSendReschedule` (`kanban_code.gs:231`). It builds a URL and Whapi-sends it to the customer's WA group.

The URL it builds today is the **old GitHub-Pages-direct path**:
```
https://innoaid.github.io/leak-guard-crm/booking.html?phone=...&name=...&group=...
```

Round 18-24 migrated all other booking-link callsites to the canonical short form:
```
https://leakguard.my/appointment/?p=<phone>
```

This callsite was missed. Customers tapping the reschedule link from this button hit the long URL — which still works today (it redirects), but it's inconsistent with every other link the bot sends, and it'll break the day we drop the GitHub-Pages canonical redirect.

### Patch

In `kanban_code.gs` `handleSendReschedule`:

```js
// OLD
const params = [];
if (phone) params.push('phone=' + encodeURIComponent(phone));
if (name) params.push('name=' + encodeURIComponent(name));
if (groupName) params.push('group=' + encodeURIComponent(groupName));
const longUrl = 'https://innoaid.github.io/leak-guard-crm/booking.html?' + params.join('&');

const msg = 'Click the link below to manage your appointment — only takes 2 mins.\n\n' +
  '🚀 Express Booking: ' + longUrl;

// NEW
const url = 'https://leakguard.my/appointment/' + (phone ? ('?p=' + encodeURIComponent(phone)) : '');

const msg = 'Hi ' + (name || 'there') + ', here\'s the link to reschedule your site visit.\n\n' +
  'Check Real Time Availability / Book Your Slot Instantly: ' + url;
```

The new copy matches the Round 42c welcome-message wording so the customer sees the same call-to-action they're used to.

### File
- `C:/Projects/leak-guard-crm/kanban_code.gs` lines ~250-258 — replace the URL builder + the message string.
- After edit: user redeploys Apps Script Web App (same drill as previous handler additions). Single redeploy covers Task 1 and Task 2's URL fix together since neither touches WEBAPP_URL itself.

### Verification

1. After Apps Script redeploy: open kanban → SVC card → modal → tap "Send reschedule link" → check the customer's WA group.
2. Expected msg: `Hi <name>, here's the link to reschedule your site visit.\n\nCheck Real Time Availability / Book Your Slot Instantly: https://leakguard.my/appointment/?p=<phone>`
3. Tap the link in WA → loads `leakguard.my/appointment/` with their phone prefilled and their existing booking surfaced for reschedule.

---

## Task 3 — FB ad attribution to Pending Invitation conversion event (DEFERRED, needs your input)

### Context (the actual question)

User runs FB ads that funnel leads into the CRM sheet. They want to know **which ad creative converts the most leads to "Pending Invitation"** (the moment admin reviews a New Lead and decides to invite them). They want this metric visible **in Facebook Ads Manager**, not just internally.

`Pending Invitation` is the right anchor: it's the moment the admin commits to engaging the lead, so it's a much higher-quality signal than raw lead-form-fill.

### Architecture

To get the metric **inside Ads Manager**, FB needs a conversion event sent back to it. Three options:

| Approach | Ads Manager visibility | Effort | Recommended? |
|---|---|---|---|
| Facebook **Conversions API** (CAPI, server→server) | Yes, native | medium | **Yes** |
| Offline events CSV upload to FB Business Manager | Yes (manual) | low (recurring) | No — too manual |
| Internal-only dashboard (no FB attribution) | No | low | Only if FB CAPI access is blocked |

CAPI is the right fit. FB sends events back to the originating ad via `fbclid` + hashed `phone`/`email`.

### Required setup before we can ship

1. **FB Lead Ads** (most likely your case if leads "come into CRM as a list") — confirm.
2. **FB Business Manager access** with permission to: create Pixel events, generate System User Access Token, define Custom Conversions.
3. **Pixel ID** — find or create one in Events Manager.
4. **System User Access Token** with `ads_management` + `business_management` scopes.
5. **CRM sync source** — how leads currently land in the sheet (Zapier? Make? FB Lead Ads → n8n?). Need to capture `lead_id`, `ad_id`, `adset_id`, `campaign_id`, `fbclid` at intake.

### Implementation (3 components)

**(A) Capture FB attribution data at lead intake** — modify whatever currently syncs FB leads → Sheet to also save these CRM columns:
- `FB Lead ID` (form submission ID)
- `FB Ad ID`
- `FB Adset ID`
- `FB Campaign ID`
- `FB Click ID (fbc)` — derived from `fbclid` query param if present

If your current sync (Zapier / Make / n8n) doesn't pass these through, we extend it. If you don't have a sync at all and leads land manually, we build one — would be an n8n workflow with the FB Lead Ads webhook trigger.

**(B) Trigger CAPI on Status → Pending Invitation transition**

Apps Script `handleUpdateStatus` already runs whenever admin moves a lead via kanban. Add a small extension: when the new status is `Pending Invitation` (and the previous wasn't), POST a webhook to a new n8n workflow:

`POST https://leakguard.app.n8n.cloud/webhook/lg-fb-conversion`
```json
{ "secret": "ABC", "phone": "60179934386", "leadId": "...", "adId": "...", "fbc": "..." }
```

**(C) New n8n workflow `LG - FB Conversion Send`**

```
Webhook
  ↓
Verify Secret
  ↓
SHA-256 hash phone (+ email if available)
  ↓
HTTP POST graph.facebook.com/v18.0/{PIXEL_ID}/events
  body: {
    data: [{
      event_name: "Lead",
      event_time: <unix>,
      action_source: "system_generated",
      user_data: { ph: <hashed>, em: <hashed if available>, fbc: <click_id> },
      custom_data: { content_category: "pending_invitation", lead_status: "pending_invitation" }
    }],
    access_token: "<SYSTEM_USER_TOKEN>"
  }
  ↓
Respond 200 OK
```

Single Apps Script change + single new n8n workflow + 4-5 new CRM columns.

### Setup in FB Business Manager (one-time, you do it)

1. Events Manager → your Pixel → **Custom Conversions** → New
   - Source: pixel
   - Event: `Lead`
   - Rule: `custom_data.content_category` equals `pending_invitation`
   - Save as "Lead → Pending Invitation"
2. Ads Manager → Columns → Customise → add the new conversion to the column set
3. Per-ad Pending Invitation count starts populating after the first event fires

### Effort

- ~1 hour for the n8n CAPI workflow
- ~30 min for the Apps Script trigger (extending `handleUpdateStatus` to fire the webhook on Pending Invitation transition)
- ~30 min for adding columns + tweaking the existing FB lead intake (depends on how your current sync looks)
- ~30 min for FB Business Manager setup (your side, with my walkthrough)

### Information I need from you before I can ship

I'd rather not assume. Tell me:
1. **Ad type**: FB **Lead Ads** (in-FB form) / Click-to-WA / Landing-page-via-Website? Most likely Lead Ads given the "lead list to CRM" phrasing.
2. **Current sync mechanism**: how does a new FB lead get into your Sheet today? Manual entry / Zapier / Make / a webhook to n8n / FB Lead Ads native CRM Sync?
3. **FB Business Manager access**: do you have admin access to your Pixel + can generate a System User Access Token? If not, I'll guide.

Once you answer those three, I write code.

### Privacy note (worth flagging)

CAPI sends **hashed** customer phone (SHA-256) to FB so it can match to ad clicks. Standard practice; FB's own SDK does the same client-side via the Pixel. No raw PII leaves your system. Customers should be informed in your privacy policy that conversion events are sent for ad attribution — most companies running FB ads cover this in standard boilerplate.

---

## Suggested execution order

1. **Ship Task 1 first** (mute post-QS phases) — ~5 min, isolated, single n8n PUT, no dependencies.
2. **Then Task 2 setup** — 30 min FB admin work + 2 hours code. Needs your answers to the 3 questions above before I can wire correctly.

Both tasks are independent — neither blocks the other.
