# Orphaned work — preserved, not applied

Changes that were made but never merged, kept here so they are recoverable
rather than lost. **Nothing in this folder is live.**

## `round58-fb-capi-conversion.patch`

Facebook CTWA referral capture + Conversions API fan-out. Round 58 vintage.

Found as an **uncommitted** modification in the stale `C:\projects\leak-guard-crm`
worktree on 2026-08-19. It is not in the trunk — `grep` for
`N8N_FB_CONVERSION_URL` or `handleUpdateLeadCtwa` in `kanban_code.gs` returns
nothing.

What it adds:

- `N8N_FB_CONVERSION_URL` → `https://leakguard.app.n8n.cloud/webhook/lg-fb-conversion`
- a `updateLeadCtwa` action + `handleUpdateLeadCtwa` handler, capturing FB CTWA
  referral fields on a lead's first message
- a `prevStatus` snapshot inside `handleUpdateStatus`, so the exact
  `Pending Invitation` transition can fire CAPI once without double-firing

**Do not `git apply` this.** It was written against a base ~2,795 lines shorter
than the current trunk; the context is long gone. Treat it as a specification to
re-implement if the feature is still wanted, not as a patch to replay.

The n8n side (`lg-fb-conversion`) may or may not still exist — check before
rebuilding.

## `CLAUDE-primer-2026-04.md`

A 211-line "Claude Code primer" found untracked in the same stale worktree —
a different document from the maintained 104-line `CLAUDE.md` at the repo root,
not an older revision of it. It carries stack detail the live file doesn't
(Whapi bot number, calendar list, hosting status, the v2 agent's model).

Kept for reference only. It is April 2026 vintage and unverified since — see
`../gotchas/gotcha_claudemd_stale.md` for how badly docs of this era drifted.
**The root `CLAUDE.md` remains the live one.**
