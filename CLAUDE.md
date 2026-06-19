# Kanban CRM — Project Memory

## What this project is
A Kanban-style CRM. The **database is a Google Sheet**. The `.gs` server code runs as a
**Google Apps Script web app** bound to that Sheet; the `.html` UI files are served as a
**static GitHub Pages site** at `leakguard.my`. WhatsApp/automation runs in **n8n cloud**.

**GitHub is the source of truth for all code.** New here / onboarding a fresh machine? See
**`README.md`** (quick start) and **`docs/ARCHITECTURE.md`** (system map).

---

## TWO deploy targets (don't confuse them)
- **`.html` files → GitHub Pages.** Just `git push origin main` — Pages auto-deploys to
  `leakguard.my` in ~1–3 min. **No Apps Script paste.** (`.nojekyll` at root keeps the build
  reliable; don't delete it. Hard-refresh / wait out the ~10-min CDN cache to see changes.)
- **`.gs` files → Apps Script (paste-only).** Paste into the Apps Script editor + redeploy
  (steps below). Pushing `.gs` to GitHub does **not** make the live Sheet run it.

---

## The copies of the code (these must be kept in sync)
1. **Laptop** local files
2. **Desktop** local files
3. **Live Apps Script** (the `.gs` code, bound to the Sheet)
4. **GitHub Pages** (the live `.html` UI at `leakguard.my`)

- **Git** keeps laptop ↔ desktop in sync (through GitHub). Work in **one folder** on the
  **`main`** branch only (no feature branches).
- **Deploying:** `.gs` → paste into Apps Script (below); `.html` → just `git push origin main`
  (GitHub Pages auto-deploys).
- The Sheet's CRM **data** is already cloud-synced automatically — both machines
  see the same records. Only the **code** needs manual syncing.

---

## GOLDEN RULE
**Always edit code locally (here, through Claude Code).
NEVER type fixes directly in the Apps Script web editor.**

The Apps Script editor is **paste-only**. If code is changed there and not copied
back to GitHub, GitHub falls behind, and the next `git pull` on the other machine
will be missing those changes.

---

## Routine — when I SIT DOWN at a machine
```
git pull
```
Get the latest from GitHub before touching anything. Then edit with Claude Code.

---

## Routine — push my `.gs` code to the LIVE app (so the Sheet runs it)
> Only for **`.gs`** files. `.html` files deploy by `git push origin main` (GitHub Pages) —
> do NOT paste them into Apps Script.
My normal manual method:
1. Open the updated `.gs` file (in this repo, or its GitHub raw page).
2. Copy the **full** contents.
3. In the Sheet: **Extensions → Apps Script**, open the matching `.gs` file,
   select all, paste over it.
4. **Repeat for every `.gs` file I changed.**
5. Save (Ctrl+S).
6. If this is a **web app**: Deploy → Manage deployments → edit the deployment →
   set version to **New version** → Deploy.
7. If I changed bot/automation code: also re-paste the n8n API key in the browser
   (n8n side), or the bot edits won't take effect. NEVER commit the n8n key to the repo.

> Optional faster way (automates steps 1–6): `clasp push`
> (Requires clasp installed + `clasp login` once per machine.)

---

## Routine — BEFORE I LEAVE a machine
```
git add -A
git commit -m "describe what I changed"
git push
```
If I skip this, the other machine will NOT see today's work.

---

## Full handoff cheat-sheet
- **Leaving the laptop after the meeting:** commit + push.
- **Arriving at the desktop:** `git pull`, then keep working.
- (Same in reverse for desktop → laptop.)

---

## If the code ever drifts and I'm unsure which copy is correct
- **Live Apps Script is newer** → copy the code from the editor back into the repo
  files, then commit + push.
- **Repo is newer** → paste repo files into Apps Script (manual method above).
- With clasp: `clasp pull` = cloud overwrites local; `clasp push` = local overwrites cloud.

---

## Script details
- Apps Script project type: **container-bound** to the Google Sheet.
- Script ID (Extensions → Apps Script → Project Settings → Script ID):
  `__________________________` (fill this in once)
