# Leak Guard CRM — Claude Code primer

This file is auto-loaded into every Claude Code session opened in this repo.
Read top-to-bottom on first load; subsequent sessions can skim.

---

## What this project is

A **WhatsApp-based booking + CRM system for Leak Guard Sdn Bhd** (Malaysian
waterproofing company). Customers find us via FB ads / referrals → land on
`leakguard.my/appointment/` → book a free site inspection → join a WhatsApp
group → conversational AI agent (`LG-Customer Chat v2`) handles
reschedule/cancel/follow-up → site visit → quotation → installation → balance →
job complete.

Stack:
- **n8n cloud** (https://leakguard.app.n8n.cloud) — orchestrates Whapi +
  Google Calendar + Google Sheets. ~10 active workflows.
- **Apps Script** web app — CRM mutation handlers (status updates, tag writes,
  appointment edits) + a few read endpoints.
- **Google Sheets** — the CRM (`Leak Guard Leads` tab on a single spreadsheet).
- **Whapi.cloud** — WhatsApp gateway. One bot phone `+601119338657`.
- **Google Calendar** — three calendars: LG (KL), Alvin (KL personal), LGJB.
- **GitHub Pages** — serves `leakguard.my` (booking page, kanban, homepage).
- **Plesk (Exabytes)** — used briefly for hosting; phasing out, leaving
  domain-only at registrar.
- **OpenAI gpt-4o** — v2 chat agent's brain.

---

## Critical files (your map of the codebase)

```
booking.html → /appointment/index.html  (canonical booking page; relative logo.png)
team_kanban.html                         (admin kanban; full CRM dashboard)
kanban_code.gs                           (Apps Script source-of-truth)
quotation_builder.html                   (PDF quotation generator + boilerplate merge)
job_board.html                           (alt job-tracking; rarely used)
boilerplate.pdf                          (50MB; merged into customer quotations)
logo.png / logo.jpg                      (brand assets)
index.html                               (temp homepage at leakguard.my apex)
CNAME                                    (custom domain binding for GitHub Pages)
docs/RUNBOOK.md                          (ops procedures — read this BEFORE any patch)
docs/ARCHITECTURE.md                     (system map)
docs/audit_funnel_v1.md                  (Step 0 audit doc, references for funnel work)
docs/STARTER_TASKS.md                    (your first 5 onboarding tasks)
docs/gotchas/*.md                        (hard-won bug knowledge; READ ALL before patching)
```

n8n workflows are NOT in this repo — they live in n8n cloud. Use the n8n public
API to fetch / patch them. Patch script template in `RUNBOOK.md`.

---

## Conventions you MUST follow

These are non-negotiable and have caused production bugs when ignored. Each
has a backing gotcha doc in `docs/gotchas/`.

### 1. n8n Code-node patches use Write-tool .py files, NEVER inline heredoc

Bash heredoc + Python escape parsing collapses `\b`/`\n`/`\t` inside JS string
literals. Already caused 3 outages (rounds 17, 21, 31). Always:

```python
# Write to disk first via the Write tool, then run:
# python patch_xyz.py "$N8N_API_KEY"
```

Never `python << 'PY' ... PY` for n8n Code-node jsCode patches. See
`docs/gotchas/gotcha_n8n_jscode_escapes.md`.

### 2. Sheet header names include `(<col-letter>)` suffix — match exactly

`Group Name (AE)`, `Pending Date (AF)`, etc. Writing to `Status (I)` instead of
`Status` SILENTLY no-ops. Always read the actual header strings before writing.
See `docs/gotchas/gotcha_sheet_header_suffixes.md`.

### 3. MYT timezone idiom is treacherous

`new Date(now.getTime() + 8*60*60*1000)` is SAFE for `getUTC*()` /
`.slice(0,10)` (extracts MYT calendar parts) but BUGGY if fed to
`.toISOString()` as a full UTC timestamp (puts it 8h in the future). See
`docs/gotchas/gotcha_myt_tz_double_shift.md`.

### 4. HTTP nodes with `neverError: true` swallow API errors

Calendar reads use this so n8n doesn't crash on transient blips. Downstream
Code nodes MUST explicitly check `!Array.isArray(resp.items)` and fail closed,
or you get phantom-empty calendars and double-bookings. See
`docs/gotchas/gotcha_n8n_nevererror_fail_open.md`.

### 5. v2 agent must check sender phone matches lead phone

After CRM lookup in Verify Lead, compare `senderPhone` to `lead.Phone`. Skip
otherwise. Otherwise inspector / spouse / 3rd-party messages trigger bot
replies meant for the lead. See `docs/gotchas/gotcha_v2_lead_phone_gate.md`.

### 6. Always test in Annie's sandbox group, not live

Sandbox: `SV-0426-018 Annie - Ampang`, chatId `120363426626363115@g.us`,
phone `60179934386`. Live customer groups are NOT a test environment.

### 7. Git workflow: PR → owner-merge → patch live

You don't have direct merge rights on `main`. Push to `dev` or feature
branches. Owner reviews + merges + applies n8n / Apps Script patches.
DO NOT touch live n8n workflows or live Apps Script web app deployment
without explicit approval.

---

## Tooling cheatsheet

### Fetch a workflow's current jsCode

```bash
curl -s -H "X-N8N-API-KEY: $N8N_API_KEY" \
  "https://leakguard.app.n8n.cloud/api/v1/workflows/<WID>" \
  | python -c "import json,sys; w=json.load(sys.stdin); n=next(x for x in w['nodes'] if x['name']=='<NodeName>'); print(n['parameters']['jsCode'])"
```

### Smoke-test lg-availability

```bash
curl -s "https://leakguard.app.n8n.cloud/webhook/lg-availability?phone=60179934386" | python -m json.tool
```

### Reset Annie's CRM for a fresh sandbox test

```bash
PYTHONIOENCODING=utf-8 python "C:/Users/CY Lee/AppData/Local/Temp/reset_v2_test.py"
```

### Bounce a workflow (clear Window Buffer Memory etc.)

```bash
curl -s -X POST -H "X-N8N-API-KEY: $N8N_API_KEY" \
  "https://leakguard.app.n8n.cloud/api/v1/workflows/<WID>/deactivate"
curl -s -X POST -H "X-N8N-API-KEY: $N8N_API_KEY" \
  "https://leakguard.app.n8n.cloud/api/v1/workflows/<WID>/activate"
```

Full inventory of webhook URLs in `docs/ARCHITECTURE.md`.

---

## Plan-mode usage (Claude Code)

For changes touching n8n / Apps Script / live booking flow, use plan mode:
- Type `/plan` (or invoke ExitPlanMode after writing a plan file)
- Write the plan to `.claude/plans/<descriptive-name>.md`
- Owner reviews + approves
- Then execute

For trivial in-repo edits (typos, comments, doc updates), skip plan mode.

---

## Workflow IDs (memorize these)

| Workflow | ID | Webhook |
|---|---|---|
| WA Receiver | `gpnJPMa9w5FX1fDM` | `/webhook/lg-wa-group-bot` |
| LG-Customer Chat v2 | `xhO6U0Xa8VmGPevy` | `/webhook/lg-customer-chat-v2` |
| LG-Customer Chat v1 (standby) | `Slbb6OljpwBuHjB2` | `/webhook/lg-customer-chat` |
| LG-Customer Join | `gJAArg6iDord57Ua` | `/webhook/lg-customer-join` |
| LG-Booking | `gh0pwqGygBDoNzJB` | `/webhook/lg-booking` |
| LG-Availability | `VfAulssOzbmeoogV` | `/webhook/lg-availability` |
| LG-Follow Up | `31tPYuEY86hrcoAK` | scheduleTrigger + `/webhook/lg-followup-test` |
| LG-Manual FU Send | `ZNBodjojR7Ph0LF0` | `/webhook/lg-manual-fu-send` |
| LG-Admin Commands | `PILR7VRalAqggy5P` | `/webhook/lg-admin-commands` |
| LG-Resend Welcome | `M4LMhZL1pVAid9A2` | `/webhook/lg-resend-welcome` |

---

## Where credentials live (NOT in this repo)

- n8n API key: ask owner via DM. Truncated prefix in personal CLAUDE.md.
- Whapi token: hardcoded in n8n workflows (rotate after handover)
- Apps Script web app URL + secret: hardcoded in workflows (`AKfycbx…`)
- Google Calendar IDs: hardcoded in `Process Booking Request` jsCode
- OpenAI key: stored as n8n credential (not exposed)

NEVER commit any of these to the repo. Use environment variables or
n8n credentials store.

---

## Onboarding sequence

1. Read `docs/ARCHITECTURE.md` (15 min)
2. Read `docs/gotchas/*.md` (45 min — there are 11 of them, all worth it)
3. Read `docs/RUNBOOK.md` (30 min)
4. Pair with owner on first patch (1 h)
5. Pick first task from `docs/STARTER_TASKS.md`

Total ramp-up: ~1 day to be productive.

---

## Active TODOs (from owner's roadmap)

- **Phase 1**: Full kanban automation in WA group (5–7 weeks)
- **Phase 2**: FB ad funnel — landing form → CRM → booking → group
  (2–3 weeks; Step 0 audit done in `docs/audit_funnel_v1.md`)
- **DNS migration** to GitHub Pages: in progress
- **TinyURL retargeting** to leakguard.my: optional cleanup

Owner has the latest priorities. Ask before starting any unlisted work.
