# Audit: Funnel-Ready State (Step 0)

Generated 2026-05-01 (read-only audit; no system changes made).

Baseline reference for the upcoming FB-ad funnel build. All 10 active workflows verified healthy in last 5 executions each (0/5 errors). Use this doc as the rollback / regression-detection baseline.

---

## A. Workflow inventory

| Workflow | ID | Webhook | Routes by | Last updated |
|---|---|---|---|---|
| WA Receiver | `gpnJPMa9w5FX1fDM` | `/webhook/lg-wa-group-bot` | message content + sender phone + groupId | 2026-04-30 |
| LG-Customer Chat v2 | `xhO6U0Xa8VmGPevy` | `/webhook/lg-customer-chat-v2` | called by WA Receiver; routes by Status + Tag + msg content | 2026-05-01 |
| LG-Customer Chat v1 | `Slbb6OljpwBuHjB2` | `/webhook/lg-customer-chat` | standby (no traffic since cutover) | 2026-04-30 |
| LG-Customer Join | `gJAArg6iDord57Ua` | `/webhook/lg-customer-join` | participant phone in CRM | 2026-05-01 |
| LG-Booking | `gh0pwqGygBDoNzJB` | `/webhook/lg-booking` | called by booking page; resolves CRM by phone | 2026-05-01 |
| LG-Availability | `VfAulssOzbmeoogV` | `/webhook/lg-availability` | calendar query + CRM lookup by phone | 2026-05-01 |
| LG-Follow Up | `31tPYuEY86hrcoAK` | scheduleTrigger + `/webhook/lg-followup-test` | filters CRM by Status + Tags + intervals | 2026-04-30 |
| LG-Manual FU Send | `ZNBodjojR7Ph0LF0` | `/webhook/lg-manual-fu-send` | phone list from kanban | 2026-04-30 |
| LG-Admin Commands | `PILR7VRalAqggy5P` | `/webhook/lg-admin-commands` | `cmd` field switch (link / pause / reset / etc.) | 2026-05-01 |
| LG-Resend Welcome | `M4LMhZL1pVAid9A2` | `/webhook/lg-resend-welcome` | phone + status PSV | 2026-04-30 |

**Plug-in points for new funnel:**
- New `lg-form-submit` webhook → standalone workflow OR new node in LG-Customer Join
- New customer-DM routing → WA Receiver `Parse & Route` (line ~62-66, before "LIVE v2 path")
- Welcome branching → LG-Customer Join `Build Welcome with Slots` jsCode (lead has Cal Event ID test)

---

## B. Status / Tag transition map

### Statuses (current CRM distribution, n=441)

| Status | Count | Set by | Trigger |
|---|---|---|---|
| New Lead | 409 | Webform/TypeForm import | external; default for fresh leads |
| Site Visit Confirmed | 16 | LG-Booking `Update CRM` + LG-Admin Commands `Update CRM Appt` | booking page submit / admin -appt set |
| Pending Site Visit | 12 | LG-Customer Join `Update CRM Join` + LG-Admin Commands `Reset CRM` | customer joins WA group / admin reset |
| Pending Invitation | 2 | LG-Admin Commands `Update Group Name CRM` | -admin link command (group renamed, awaiting customer join) |
| Pending QT | 1 | manual / kanban move | post-visit |
| Quotation Sent | 1 | LG-Quotation Send `Update CRM` | quotation_builder.html submit |

**Implied funnel:** `New Lead → Pending Invitation → Pending Site Visit → Site Visit Confirmed → Pending QT → Quotation Sent → ...`

**Funnel reorder for FB ads:** the new flow has booking BEFORE group, so transition is `New Lead (form) → Site Visit Confirmed (booking) → [Tag: needs_group_creation] → SVC + group linked (admin) → existing flow`. Status doesn't go backwards; tag indicates "needs admin action".

### Tags (5 in use, n=6 rows tagged)

| Tag | Set by | Cleared by | Effect |
|---|---|---|---|
| `needs_reply` | LG-Customer Chat v1 (when complaint/cancel detected) | LG-Admin Commands `Clear Tag` | red border in kanban; pauses bot per tag-filter logic |
| `complaint` | LG-Customer Chat v1 `Tag CRM Complaint` | manual | red border; bot stays silent in group |
| `cancel_request` | LG-Customer Chat v1 `Tag CRM Cancel` | manual | red border on SVC cards |
| `cold_lead` | LG-Follow Up after FU#5 | manual | filtered OUT of follow-up cron |
| `fu_paused` | kanban "Pause FU" button (LG-Admin Commands `Pause Bot in CRM`) | kanban "Resume FU" | filtered OUT of follow-up cron |

**Free namespace for new tags:** `needs_group_creation`, `landing_form`, `pixel_attributed_*` — none collide with existing.

---

## C. Routing rules ledger (WA Receiver `Parse & Route`)

Evaluation order matters — first-match wins. Current rules:

| # | Predicate | Outcome | Notes |
|---|---|---|---|
| 1 | `groupsParticipants[].action ∈ ['add','join','invite']` | `customer_joined` (→ LG-Customer Join) | participant join event |
| 2 | other group_participants events | `skip:non-join event` | leave/promote/demote |
| 3 | `body.groups.length > 0` | `skip:group metadata` | group create/update events |
| 4 | `!messages.length` | `skip:no messages` | |
| 5 | `fromMe === true` | `skip:outgoing` | bot's own messages |
| 6 | reaction event (positive emoji) | `msgText='yes'` (continues to fork below) | round 22 |
| 7 | reaction event (negative emoji) | `msgText='cancel'` (continues) | |
| 8 | reaction event (neutral) | `skip:neutral_reaction_ignored` | |
| 9 | reaction from staff | `skip:staff_reaction_ignored` | |
| 10 | **`!isGroup && isAdmin`** | `admin_private` (→ LG-Admin Commands) | DM from admin only — line 54 |
| 11 | **`isGroup && isStaff && msgText.startsWith('-admin')`** | `admin_command` (→ LG-Admin Commands) | line 55 |
| 12 | `chatId === Annie sandbox` | `customer_message_v2` (→ LG-Chat v2) | line 59 |
| 13 | **`isGroup && !isStaff && !fromMe`** | `customer_message_v2` (→ LG-Chat v2) | line 63 — LIVE v2 path |
| 14 | else | `skip:unhandled` | fallback |

**Critical observation for new flow:** rule #10 only allows DMs from admin (60183639321). Customer DMs to bot are dropped with `skip:unhandled` (rule #14). To enable customer group-request DMs, **add a new rule between #10 and #11** that:
- Matches `!isGroup && !isStaff` AND msgText matches the "group request" keyword pattern
- Routes to a NEW workflow (e.g. `customer_dm_group_request` → `lg-customer-dm-handler` webhook)
- Does NOT loosen `isStaff` for `-admin` (rule #11) or `admin_private` (rule #10)

This keeps existing admin paths unchanged; new path is additive.

---

## D. Phone normalization audit

19 nodes do phone normalization. Two patterns dominate:

**Pattern 1: `replace(/\D/g, '')`** — 11 nodes
- LG-Admin Commands: Extract Group Names, Extract Pause Names, Find Lead, Process AI Cmd
- LG-Customer Chat v1: Classify Message
- LG-Customer Chat v2: Cancel/Escalate/Send Whapi Reply/SendLink/Verify Lead handlers
- LG-Manual FU Send: Build & Send
- LG-Resend Welcome: Find Lead & Build Msg

**Pattern 2: `endsWith(phone.slice(-8))`** — 8 nodes
- LG-Availability: Build Availability
- LG-Booking: Find Lead and Build Update, Resolve Effective Old Event
- LG-Customer Chat v1: Verify Lead
- LG-Customer Join: Match Customer
- LG-Customer Chat v2: indirect via Verify Lead
- LG-Manual FU Send: Build & Send (also)
- WA Receiver: Parse & Route reaction-staff filter

**Variants of the lookup:**

| Workflow / Node | Lookup formula |
|---|---|
| LG-Customer Join `Match Customer` | `phone===customerPhone \|\| phone==='6'+customerPhone \|\| phone.endsWith(customerPhone.slice(-8))` |
| LG-Booking `Find Lead and Build Update` | `rowPhone === phone \|\| (phone.length >= 8 && rowPhone.endsWith(phone.slice(-8)))` |
| LG-Admin Commands `Find Lead` | `String(r.json['Phone']\|\|'').replace(/\D/g, '').endsWith(cleanDigits)` |

**Risk:** Last-8 collision with 441 leads is unlikely BUT real. Of the 12 PSV leads I sampled, none collide. The Find Lead tier logic (active > all > unlinked) provides some defense by preferring an active row over collision artifacts. **For new customer DMs:** do exact phone match first; only fall back to last-8 if no match found, AND alert admin if multiple rows still match.

---

## E. Welcome message paths

| Trigger | Where | Current behavior | Conditional needed for new flow? |
|---|---|---|---|
| Customer joins group | LG-Customer Join `Build Welcome with Slots` | sends booking link msg (Variant C) | **YES** — branch on `Cal Event ID` exists → "your inspection is on [date+time]" instead |
| Admin runs `-admin invite resend` | LG-Resend Welcome `Find Lead & Build Msg` | resends Variant C | possibly — same logic as above |
| Customer says "send link" mid-chat | LG-Chat v2 `Debug Skip Echo` | sends booking link inline | no change needed — customer is already in group |
| Tool-call from agent | LG-Chat v2 `SendLink Handler` | sends booking link via webhook | no change needed |

**Step 1 scope (welcome branching) touches LG-Customer Join + LG-Resend Welcome.** These two share the same anchor pattern (`booking link` template literal); single patch script can update both.

---

## F. Ghost / orphan scenarios (current CRM)

Sampled all 441 rows from latest LG-Follow Up exec:

| Scenario | Count | Severity |
|---|---|---|
| PSV/SVC with no Group ID | 0 | clean |
| SVC with no Cal Event ID | 0 | clean |
| Cal Event ID set but Status outside SVC/QS pipeline | **2** | minor — likely archived leads with stale cal IDs |
| Phone normalization collisions (last-8) | not yet enumerated | low risk per spot-check |

The 2 weird Cal-Event-ID-with-non-SVC rows are probably from a customer who booked then was archived (Lost / Cold Lead) without explicit calendar cleanup. Not blocking — they're terminal-state, not active.

**For new flow:** when admin uses `-admin link` to attach group to an SVC lead with Cal Event ID, the existing flow handles it correctly via `Resolve Effective Old Event` (no race with the calendar event). Already shipped via round 35B.

---

## G. Crash-scenario map (current mitigation status)

### Class A — Existing flow regressions

| # | Scenario | Mitigation status |
|---|---|---|
| A1 | Customer DMs bot routed to v2 chat agent | **NOT YET** — need new routing rule between #10 and #11 |
| A2 | Non-staff DMs trigger LG-Admin Commands | safe — rule #10 `!isGroup && isAdmin` is strict; rule #11 requires both `isStaff && '-admin'` prefix |
| A3 | Customer joining group → wrong welcome | **NOT YET** — Step 1 fixes |
| A4 | LG-Follow Up fires PSV-style FU at SVC-already-booked-no-group | **PARTIAL** — Find Leads filters by Status only; doesn't check Group ID. New flow leads (SVC + no group) would be skipped because Status != PSV/QS. **Safe by accident.** |
| A5 | Manual FU Send for landing-page leads | only fires for PSV/QS Status — landing-page-form-submit creates `New Lead` Status which is filtered out |

### Class B — Bot DM handling failure modes

| # | Scenario | Mitigation needed |
|---|---|---|
| B1 | Customer DMs from phone NOT in CRM | new handler: lookup → if no match, silent skip + admin DM |
| B2 | Customer's WA phone differs from CRM phone | suffix-match last 8; alert admin on ambiguity |
| B3 | Customer DMs random text | strict keyword match (Q3 default) |
| B4 | Customer DMs 50× rapidly | rate limit per-phone in handler |
| B5 | Customer DMs after group already created | handler checks `Group ID` field; if set, reply "already in your group" |
| B6 | Customer DMs before booking complete | handler checks `Cal Event ID`; if missing, reply "complete booking first" |
| B7 | CRM has 2 rows with same phone (last-8 collision) | tier logic — pick most recent active |
| B8 | Whapi DM ack send fails | retry 3× backoff; admin sees pending tag anyway |

### Class C — Form submission

| # | Scenario | Mitigation status |
|---|---|---|
| C1 | Phone validation accepts bogus number | new — strict regex `^60\d{9,10}$` after normalization |
| C2 | Same phone fills form 5 times | new — Apps Script `handleLandingFormSubmit` uses smart-hybrid (existing pattern in `handleCreateLead`) |
| C3 | Form submit → redirect to booking fails | new — server returns booking URL in JSON; client JS does redirect with fallback display |
| C4 | Customer abandons after form, never books | acceptable — `New Lead` row visible in kanban; manual FU can re-engage |
| C5 | Honeypot caught a bot | silent reject (return 200 ok); admin DM count |
| C6 | Apps Script quota exhausted | move form submit to n8n webhook instead (higher quota) |

### Class D — Concurrency / race

| # | Scenario | Mitigation status |
|---|---|---|
| D1 | Customer fills form + admin already had them in CRM | smart-hybrid handles |
| D2 | Two customers grab same slot | round 35B post-create race-verify (already shipped) |
| D3 | Customer fills form twice → 2 bookings | round 35B force-reschedule on existing Cal Event ID (already shipped) |
| D4 | Bot creates group at same moment admin is creating | new flow: bot DM handler ONLY tags + DMs admin, NEVER creates group itself |

### Class E — Whapi platform limits

| # | Scenario | Mitigation status |
|---|---|---|
| E1 | Whapi exec quota overrun from ad burst | known — already documented (`gotcha_n8n_exec_limit.md`) |
| E2 | WA flags bot for too many group creates | new — kanban "Create Group" button rate-limit max 5/day |
| E3 | Customer's WA invite link expires | refresh via Whapi API on demand |
| E4 | Customer doesn't have WhatsApp | form ask "WA-enabled? yes/no" — defer for v2 |

### Class F — Funnel UX

| # | Scenario | Mitigation status |
|---|---|---|
| F1 | Customer lands on TQ directly (no booking) | new — TQ page checks query params; redirects to `/appointment/` if missing |
| F2 | Customer hits back → re-submits booking | already idempotent via `existingAppt` detection |
| F3 | Customer closes TQ before clicking group btn | acceptable — admin sees SVC-no-group lead in kanban |
| F4 | Customer fills form but exits before redirect | acceptable — `New Lead` visible in kanban |

---

## H. Recommendation for Step 1: Welcome message branching

### Why this first

- Lowest risk: change is invisible to current customers (no Cal Event ID at join time → existing welcome path unchanged)
- Defensive: ready when new-flow customers start joining groups already-booked
- Single workflow, single Code node
- Can be tested with a synthetic `lg-customer-join` webhook fire on Annie's group AFTER manually setting Annie's Cal Event ID
- Patch via Write-tool .py file (per heredoc gotcha — emoji-safe)

### Concrete patch

**File:** n8n `LG - Customer Join` (`gJAArg6iDord57Ua`) → `Build Welcome with Slots` jsCode

**Anchor (current):**
```js
const groupName = String(data.groupName || data['Group Name (AE)'] || '').trim();
```

**Add immediately after:**
```js
const calEventId = String(data['Cal Event ID (AH)'] || data.calEventId || '').trim();
const slotChosen = String(data['Slot Chosen'] || data.slotChosen || '').trim();
const hasBooking = !!(calEventId && slotChosen);
```

**Modify msg1 template** to branch:
```js
const msg1 = hasBooking
  ? `@${phone} Hi ${name} 👋 Welcome to your appointment group!\n\n📅 Site Visit confirmed: ${slotChosen}\n\nOur Senior Inspection Team will be there. Reply here if anything changes — we'll handle the rest.`
  : `@${phone} Hi ${name} 👋 We're Leak Guard, the waterproofing specialist team.\n\nWould like to set up a Free Site Inspection with our Senior Inspection Team.\n\n📋 View our team calendar and express-book your slot:\n${bookingUrl}\n\n💬 Or just tell us your preferred date — we'll book it for you.`;
```

**Caveat:** `data` here is the lead row from `Match Customer`. Need to confirm `Cal Event ID (AH)` and `Slot Chosen` are passed through Match Customer's output. From the audit, Match Customer only returns: `name, phone, problemType, location, groupName`. **Step 1 sub-task:** extend Match Customer to also pass `calEventId` and `slotChosen`. Both nodes change in same patch.

### Synthetic test (no live customer impact)

1. Manually set Annie's CRM row Cal Event ID + Slot Chosen via Apps Script setPending-style endpoint (or temporary direct sheet edit)
2. Fire `lg-customer-join` webhook with Annie's groupId + phone
3. Observe Build Welcome with Slots output → should produce the SVC variant
4. Annie's group receives the new "Welcome to your appointment group" message
5. Reset Annie's CRM via reset script

### Rollback

Single string revert. Set patch script `OLD/NEW` swap to undo.

### Risks

- If `Match Customer` extension breaks the existing data flow, customer joining a group could fail to receive welcome at all. Test thoroughly with fresh fields BEFORE PUT.

---

## Out of scope for this audit

- Sunday's leakguard.my migration (separate workstream, in progress)
- Quotation PDF testing (Phase 4 from earlier discussion)
- FB Pixel / ads campaign work
- Updating any code or config (this is a read-only audit)

---

## Quality gates (all met)

✓ Every active workflow appears in section A
✓ Every Status value in current use appears in section B with set-by reference
✓ Section C documents every Parse & Route fork in evaluation order
✓ Section H has concrete Step 1 plan with anchor strings + synthetic test plan + rollback
