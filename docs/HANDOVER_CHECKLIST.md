# Handover Checklist

For the new SE coming on board. Owner runs through this once with you on
day 1.

---

## Access you need

| What | Level | How to grant |
|---|---|---|
| GitHub repo `innoaid/leak-guard-crm` | **Triage** (read + open PRs, no merge) | owner adds via `Settings → Collaborators` |
| n8n cloud `leakguard.app.n8n.cloud` | **Editor** on `[DEV]`-suffixed workflows only | owner upgrades plan if needed; Free tier may not support per-workflow ACL — use sandbox-group gate as defense in depth instead |
| Google Sheet (CRM) | **Viewer** initially; **Editor** after first PR shipped | owner shares via Sheet UI |
| Apps Script project | **Editor** on a CLONED project `LG-CRM-Dev`; NOT live | owner makes a copy + shares |
| Whapi dashboard | **Read-only** to inspect token health | owner shares Whapi web login if needed |
| Plesk (if still relevant) | **None** — owner rotates creds; legacy | n/a |
| Google Calendar (3 calendars) | **Reader** | owner shares from Calendar UI |
| OpenAI account (for n8n credential) | **None** — never expose keys | n/a |

⚠️ **Never ask for the n8n API key in plain text or share it via Slack/email.
Owner will paste it in Claude Code chat as needed and rotate after handover.**

---

## Your local setup

1. Install **Claude Code** (https://claude.com/claude-code)
2. Clone the repo:
   ```
   git clone https://github.com/innoaid/leak-guard-crm.git
   cd leak-guard-crm
   ```
3. Open with Claude Code — it auto-loads `CLAUDE.md` (the project primer)
4. Read `docs/ARCHITECTURE.md` + `docs/gotchas/*.md`

You don't need a Python environment beyond the system Python (3.10+). All
patch scripts use stdlib only.

---

## Day-1 walkthrough (with owner, ~2 hours)

1. **Tour the live system** — owner opens kanban, walks through a real
   customer in each phase
2. **Pull up a recent v2 exec** — read the runData together; understand how
   Verify Lead → AI Agent → Send Whapi Reply flows
3. **Watch a patch ship live** — owner runs a small patch script while you
   watch the Bash output
4. **Inspect Whapi** — see the message log for Annie's sandbox
5. **Read 3 gotchas together** (heredoc, MYT TZ, Sheet headers) — these are
   the most common pitfalls
6. **Pick first starter task** from `docs/STARTER_TASKS.md`

---

## Communication norms

| Channel | What |
|---|---|
| GitHub PRs | code review, comments inline |
| WhatsApp DM with owner | quick questions, urgent issues |
| Daily standup | 5-min DM: what shipped yesterday, what's next, any blockers |
| Weekly demo | live walk-through of new features in Annie sandbox |

Don't push to `main`. Always feature branches → PR → owner merge.

---

## What "done" means for a task

A task is done when:
- ✅ Code shipped + verified in sandbox
- ✅ PR merged
- ✅ Smoke test passed on live
- ✅ Doc updated if behavior changed (`ARCHITECTURE.md`, `RUNBOOK.md`,
     or new `gotchas/*.md`)
- ✅ Memory entry added if a new gotcha discovered

Don't move on without all 5.

---

## Emergency protocol

If you accidentally ship a regression to live:

1. **Don't panic.** Most regressions are reversible.
2. **DM owner immediately** — even 2 AM. The cost of waking them is far less
   than the cost of an hour of broken bot.
3. **Check exec history** — is the bot crashing or silently misbehaving?
4. **Run the rollback**: every patch script keeps OLD/NEW; swap and re-run.
5. **Post-mortem**: write up what happened, what was missed in review, save
   to a new `gotchas/*.md` so it doesn't repeat.

---

## When to ask vs decide

**Always ask owner before:**
- Deploying to live (production n8n / Apps Script web app / kanban)
- Modifying staff phone allowlist
- Touching Whapi token / API keys
- Changing CRM column structure
- Bulk operations on >5 leads
- Anything labeled "out of scope" in a plan

**You can decide solo:**
- Refactoring within a single Code node
- Adding gotcha docs
- Writing tests
- Updating starter tasks list
- Adding kanban filter buttons (frontend only)
- Creating new dev-environment helpers

If unsure → ask. The cost of asking is 5 minutes. The cost of breaking
production is hours.

---

## Project North Star

We're building a **WhatsApp-first CRM** that lets Leak Guard handle 10× more
leads with the same staff size. Every feature should:
- Reduce manual admin work (kanban automation, bot taking over the routine)
- Improve customer experience (faster replies, no missed bookings)
- Stay rock-solid (zero double-bookings, no bot misfires)

We optimize for ZERO production bugs first; nice-to-haves second. The
architecture has gone through 38 rounds of patches getting here. Don't undo
that work in the name of refactoring.
