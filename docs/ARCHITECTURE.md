# Architecture — Leak Guard CRM

System map. For deeper crash-scenario coverage see `audit_funnel_v1.md`.

---

## High-level flow

```
[FB ad / referral / TypeForm]
         │
         ▼
[CRM row created — Status: New Lead]
         │
         ▼
[Customer joins WA group] ──────► LG-Customer Join ──► Status: Pending Site Visit
         │                                                │
         │                                                ▼
         │                                       [Welcome msg with booking link]
         │                                                │
         ▼                                                ▼
[Customer messages in group] ──► WA Receiver ──► v2 Chat Agent
                                                     │
                                                     ▼
                                              [Conversation: AI handles
                                               reschedule / book / cancel]
                                                     │
                                                     ▼ [BOOK marker]
                                                LG-Booking ──► Calendar event
                                                     │       │
                                                     │       ▼
                                                     │  Status: Site Visit Confirmed
                                                     │       │
                                                     ▼       ▼
                                          [Confirmation msg + supervisor details]
                                                     │
                                                     ▼
                                          [Inspection happens]
                                                     │
                                                     ▼
                                  [Manual: kanban move to Pending QT]
                                                     │
                                                     ▼
                              [Quotation builder → quotation_builder.html]
                                                     │
                                                     ▼ [POST /lg-quotation-send]
                                          LG-Quotation Send → Whapi → Status: Quotation Sent
                                                     │
                                                     ▼
                                  [Manual transitions: Pending I.Date →
                                   Pending Downpayment → Pending Balance →
                                   Job Complete / Receipt Sent]
```

---

## Workflow inventory (active in n8n cloud)

| Workflow | ID | What | Trigger |
|---|---|---|---|
| **WA Receiver** | `gpnJPMa9w5FX1fDM` | Whapi webhook ingestion + routing fork | Whapi POST → `/lg-wa-group-bot` |
| **LG-Customer Chat v2** (LIVE) | `xhO6U0Xa8VmGPevy` | AI agent for in-group customer messages | called by WA Receiver |
| **LG-Customer Chat v1** (standby) | `Slbb6OljpwBuHjB2` | Pre-v2 link-only chat flow | dormant; rollback target |
| **LG-Customer Join** | `gJAArg6iDord57Ua` | Welcomes new customer when they join group | Whapi participant-add event |
| **LG-Booking** | `gh0pwqGygBDoNzJB` | Commits booking → calendar event + CRM update | called by booking page + v2 |
| **LG-Availability** | `VfAulssOzbmeoogV` | Returns slot availability for booking page + v2 tool | called externally + by v2 |
| **LG-Follow Up** | `31tPYuEY86hrcoAK` | Hourly cron sending PSV/QS follow-ups | scheduleTrigger + test webhook |
| **LG-Manual FU Send** | `ZNBodjojR7Ph0LF0` | Bulk FU triggered from kanban | webhook |
| **LG-Admin Commands** | `PILR7VRalAqggy5P` | -admin link / pause / reset / etc. | called by WA Receiver |
| **LG-Resend Welcome** | `M4LMhZL1pVAid9A2` | Resends welcome on PSV transition | webhook |
| **LG-Quotation Send** | `OvefOvUlbbH4ZxLm` | Receives PDF from quotation_builder + sends via Whapi | webhook |
| **LG-Customer Appointment Reminder** | `0l3d6TdqoFPNkUGU` | Daily cron — pings tomorrow's appts | scheduleTrigger |
| **LG-Daily Morning Summary** | `iqzsrQa1rivYcucj` | Morning summary DM to admin | scheduleTrigger |
| **LG-Hourly Last Msg Refresh** | `ujFfgkAJRgWxLnmy` | Refreshes Last Customer Msg from group history | scheduleTrigger |
| **LG-Refresh PSV Last Msgs** | `weBw4E3yz1Hcszcv` | Bulk msg refresh triggered from kanban | webhook |
| **LG-Post Visit Check** | `ln0CgGFBFwz5xivb` | Detects expired visits (date passed) | scheduleTrigger |
| **LG-SVC Leaving Cleanup** | `UNZoTmvF1ZocEiUn` | Deletes calendar event when SVC moves away | webhook |
| **LG-Quotation Create** | `Uab4acePy2tV4bvG` | Quotation flow init | webhook |

Inactive: `LG-Group Creator`, `LG-Whapi Group Bot`, `LG-Fonnte Group Bot` (legacy/disabled).

---

## CRM column map

| Header | Letter | Purpose | Set by |
|---|---|---|---|
| Phone | (B) | E.164 Malaysian phone | TypeForm import / kanban "+ New Lead" |
| Name | (C) | Customer name | same |
| Problem Type | (D) | Leak description | TypeForm |
| Location | (E) | KL / SGR / JB / etc. (used for team detection) | TypeForm |
| Status | (I) | The phase column (PSV, SVC, QS…) | various workflows |
| Slot Chosen | (J) | Human-readable slot string | LG-Booking |
| Date Appt Confirmed | (K) | Booking timestamp | LG-Booking |
| Group ID (AB) | (AB) | WA group `120363xxxxx@g.us` | LG-Admin Commands link flow |
| Group Name (AE) | (AE) | `SV-MMYY-NNN <Name> - <Loc>` | LG-Admin Commands Gen Seq Number |
| Cal Event ID (AH) | (AH) | Google Calendar event id | LG-Booking |
| Group Invite Link (AJ) | (AJ) | WA group invite URL | LG-Admin Commands |
| Pending Date (AF) | (AF) | v2 agent: server-side context date | Apps Script setPending |
| Pending Slot (AG) | (AG) | v2 agent: pending confirmation slot | Apps Script setPending |
| Pending Confirmation (AI) | (AI) | timestamp for pending TTL | Apps Script setPending |
| Tags | (?) | comma-separated: needs_reply, complaint, fu_paused, etc. | various |
| Last Bot Msg Time (AD) | (AD) | last bot message timestamp | Customer Join, etc. |
| Last Customer Msg (AM) | (AM) | last customer message in group | LG-Hourly Last Msg Refresh |
| Last Follow Up At (AL) | (AL) | last manual/auto FU sent | LG-Follow Up |
| Follow Up Count (AK) | (AK) | how many FUs sent | LG-Follow Up |

⚠️ **Always include the `(<letter>)` suffix when reading/writing via n8n
Sheets node**, otherwise silent no-op. See gotcha 2.

---

## Status / Tag semantics

### Statuses (in funnel order)

```
New Lead → Pending Invitation → Pending Site Visit → Site Visit Confirmed
       → Pending QT → Quotation Sent → Pending I.Date → Pending Downpayment
       → Pending Balance → Job Complete / Receipt Sent → (terminal)
```

Plus parallel terminal states: `Lost`, `Cold Lead`, `Rejected`, `Out of Area`,
`Human Handoff`, `Completed`.

### Tags (independent of status)

| Tag | Set by | Effect |
|---|---|---|
| `needs_reply` | LG-Customer Chat v1 (complaint/cancel) | red border + bot pauses |
| `complaint` | LG-Customer Chat v1 | red border |
| `cancel_request` | LG-Customer Chat v1 | red border on SVC |
| `cold_lead` | LG-Follow Up after FU#5 | filtered out of FU cron |
| `fu_paused` | kanban "Pause FU" | filtered out of FU cron |

Free namespace for new tags: `needs_group_creation`, `landing_form`,
`pixel_attributed_*`.

---

## Calendar architecture

3 Google Calendars consulted by `LG-Availability` and `LG-Booking`:

| Calendar | ID (truncated) | Owner | All-day events = blackout? |
|---|---|---|---|
| **LG (KL team)** | `bd989...0759546f` | shared team | YES (public holidays etc.) |
| **Alvin (KL personal)** | `alvinlai.aid@gmail.com` | Alvin | NO (job-tracking notes) |
| **LGJB (JB team)** | `3e5d4...99e1e45` | JB team | YES |

**Critical rule:** Alvin's all-day events are job-tracking notes ("Riza Sri
Hartamas") — they do NOT block site-visit slots. Team calendars' all-day
events ARE blackouts. See `docs/gotchas/gotcha_calendar_allday_per_team.md`.

---

## Slot maps per team

| Slot # | KL time | JB time |
|---|---|---|
| 1 | 9:00 AM – 10:00 AM | 11:30 AM – 12:30 PM |
| 2 | 10:30 AM – 11:30 AM | 2:00 PM – 3:00 PM |
| 3 | 12:00 PM – 1:00 PM | (none) |
| 4 | 2:00 PM – 3:00 PM | (none) |
| 5 | 3:30 PM – 4:30 PM | (none) |

Team detected from `Group Name (AE)` prefix:
- `SVJB-` / `QTJB-` → JB
- otherwise → KL

Verify Lead injects the correct slot map into agentInput per turn (round 33).

---

## Phone normalization

Two patterns (kept consistent across 19 nodes — see audit doc):

```js
// Strip non-digits
const norm = String(phone || '').replace(/\D/g, '');

// Suffix-9 fallback for 0xxx ↔ 60xxx variants
const matches = norm === target || (norm.length >= 9 && (
  norm.endsWith(target.slice(-9)) || target.endsWith(norm.slice(-9))));
```

Last-8 collisions are theoretically possible with 441+ rows; mitigated by
preferring active leads in tier-based lookups.

---

## v2 chat agent — internal flow

```
Whapi msg
  ▼
WA Receiver Parse & Route ──► route by sender + chatId
  ▼ (customer_message_v2)
v2 Webhook
  ▼
Read CRM (full sheet)
  ▼
Verify Lead ──► CRM lookup by groupId
            ──► sender phone match (round 37)
            ──► layered classification:
                  - Layer 1: escalation regex (warranty/price/complaint)
                  - Layer 2: send-link request
                  - Layer 3: silent_ack
                  - Layer 4: messageMode classifier
                  - Layer 5: post-booking gate
                  - Build agentInput (with team slot map, calendar ref, pending instruction)
  ▼
IF Should Respond
  ▼
Window Buffer Memory (k=10)
  ▼
OpenAI Chat Model (gpt-4o, temp 0.3, returnIntermediateSteps=true)
  ▼
AI Agent ──► tools: check_availability (only)
        ──► output markers: [PROPOSE], [CONTEXT], [BOOK], [CANCEL], [ESCALATE]
  ▼
Send Whapi Reply ──► strip ``` fences (round 36)
                ──► day-name auto-correct (round 38)
                ──► marker dispatcher:
                      [CONTEXT] → store in pending date
                      [PROPOSE] → store pending date+slot
                      [BOOK]    → POST to LG-Booking → confirmation msg + supervisor details
                      [CANCEL]  → admin DM
                      [ESCALATE] → silent + admin DM
                ──► race-verify (round 35B)
                ──► auto-pick guard (round 16)
                ──► confirm reprompt (round 25)
  ▼
Update Last Bot Msg
```

---

## Key external integrations

| Service | What | How |
|---|---|---|
| **Whapi** | WhatsApp gateway | bearer token, REST API |
| **OpenAI** | gpt-4o for v2 agent | n8n credential, ~$7/day at current volume |
| **Google Calendar API** | KL/JB/Alvin calendars | OAuth via n8n credential |
| **Google Sheets API** | CRM read/write | OAuth + Apps Script web app for mutations |
| **Apps Script web app** | CRM writes via JSON POST | shared secret `ABC`, hardcoded URL |
| **GitHub Pages** | leakguard.my booking page + kanban | static, custom domain |
| **TypeForm** | external lead capture (separate) | not in repo; integrates via Sheet append |

---

## Where to study deeper

- Workflow inventory + crash scenarios: `docs/audit_funnel_v1.md`
- Operational procedures: `docs/RUNBOOK.md`
- Hard-won bugs: `docs/gotchas/` (11 files; read all)
- Patch examples: `C:/Users/CY Lee/AppData/Local/Temp/n8n_v2/patch_*.py`
