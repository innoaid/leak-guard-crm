# Kanban CRM — Project Memory

## What this project is
A Kanban-style CRM. The **database is a Google Sheet**, and the app is built with
**Google Apps Script that is container-bound to that Sheet**. This repo holds the
`.gs` server code and `.html` UI files.

**GitHub is the source of truth for all code.**

---

## The three copies of the code (these must be kept in sync)
1. **Laptop** local files
2. **Desktop** local files
3. **Live Apps Script** (in the cloud, bound to the Sheet)

- **Git** keeps laptop ↔ desktop in sync (through GitHub).
- **The deploy step** (below) keeps local files ↔ live Apps Script in sync.
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

## Routine — push my code to the LIVE app (so the Sheet runs it)
My normal manual method:
1. Open the updated file (in this repo, or its GitHub raw page).
2. Copy the **full** contents.
3. In the Sheet: **Extensions → Apps Script**, open the matching file,
   select all, paste over it.
4. **Repeat for every file I changed** (don't forget the `.html` files).
5. Save (Ctrl+S).
6. If this is a **web app**: Deploy → Manage deployments → edit the deployment →
   set version to **New version** → Deploy.

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
