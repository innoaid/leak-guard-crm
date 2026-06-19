# Leak Guard CRM

A Kanban-style CRM for a waterproofing business. **The database is a Google Sheet.** The UI
is a set of static HTML pages hosted on **GitHub Pages**; all writes go through a
**Google Apps Script web app**; WhatsApp messaging and automations run in **n8n cloud**.

> New here? Read this, then `docs/ARCHITECTURE.md` (system map) and `CLAUDE.md` (the working
> rules / GOLDEN RULE). **GitHub is the source of truth for all code.**

---

## The 4 moving pieces

| Piece | What | Where |
|---|---|---|
| **Google Sheet** | The actual database — every lead is a row | Sheet `Leak Guard Leads`, ID `1FnuiZcOSy5UMQpW81I7qtU6a7NGlnHtJbH2EkVM7PLQ` |
| **Apps Script web app** | All CRM **writes** + reminders + diagnostics. JSON `POST` router keyed on `action`, shared secret `ABC` | `*.gs` files; live URL is `WEBAPP_URL` in `team_kanban.html` |
| **GitHub Pages** | The **UI** — static pages served at `leakguard.my` | `*.html` files, deployed from the `main` branch |
| **n8n cloud** | WhatsApp send/receive, booking, follow-ups, quotation send, etc. | `leakguard.app.n8n.cloud` (not in this repo) |

Full system map, workflow inventory, and column reference: **`docs/ARCHITECTURE.md`**.

---

## Two deploy paths (important — they are different!)

Pushing to GitHub updates the **code**, but only HTML auto-deploys. Know which you changed:

### HTML (`*.html`) → GitHub Pages
1. Commit and `git push origin main`.
2. GitHub Pages rebuilds and serves it at `https://leakguard.my/<file>.html` in ~1–3 min.
3. That's the whole deploy — **no Apps Script paste**.

- `.nojekyll` (at repo root) disables Jekyll so the static build is reliable — don't remove it.
- To verify a deploy quickly, hit the **origin** (bypasses the `leakguard.my` CDN, which caches
  ~10 min): `curl -s https://innoaid.github.io/leak-guard-crm/<file>.html | grep "<a marker>"`.
- Then hard-refresh the page (Ctrl+Shift+R / reopen on mobile).

### Apps Script code (`*.gs`) → paste into the editor
The Apps Script editor is **paste-only** (see GOLDEN RULE in `CLAUDE.md`). After editing a
`.gs` here:
1. Sheet → **Extensions → Apps Script**, open the matching file.
2. Select-all, paste the new contents, **Save**.
3. If it's the web app: **Deploy → Manage deployments → edit → New version → Deploy**.

---

## Working on the project

- **One folder, one branch:** work in `C:\projects\leak-guard-crm` on **`main`**. There are no
  feature branches.
- **Each session:** `git pull` first → edit → commit → `git push origin main`.

### Continue on another machine (e.g. laptop)
1. First time: `git clone https://github.com/innoaid/leak-guard-crm.git`
   Already cloned: `git pull`.
2. Open Claude Code in the folder and keep working; commit + `git push origin main` when done.

> Note: Claude Code's saved *memory* is **per-machine** — it does not sync between desktop and
> laptop. This README, `CLAUDE.md`, and `docs/` are the portable knowledge that travels with
> the repo.

---

## Reading live CRM data (read-only, no credentials)

The Sheet is queryable via its public gviz endpoint — the same read the kanban board uses.
Handy for diagnosing a specific lead:

```
https://docs.google.com/spreadsheets/d/1FnuiZcOSy5UMQpW81I7qtU6a7NGlnHtJbH2EkVM7PLQ/gviz/tq?tqx=out:json&sheet=Leak%20Guard%20Leads
```

Response is JSONP — strip the `google.visualization.Query.setResponse(...)` wrapper, then
`table.cols[].label` are the headers and `table.rows[].c[].v` are the cell values.

---

## Key files

| File | What |
|---|---|
| `kanban_code.gs` | Main Apps Script web app — `doPost` action router, reminders, diagnostics |
| `autocount.gs` | AutoCount quotation/invoice integration handlers |
| `sync.gs`, `wa_admin_bot.gs` | Supporting Apps Script (sheet sync / WA admin bot) |
| `team_kanban.html` | The kanban board UI (holds `WEBAPP_URL` + the gviz read) |
| `estimation_builder.html` | On-site estimation builder (photos, service types) |
| `quotation_builder.html` | Quotation builder → posts to n8n `LG - Quotation Send` |
| `booking.html`, `caller.html` | Customer booking page / caller call-list page |
| `docs/ARCHITECTURE.md` | System map: workflows, CRM columns, calendars, v2 chat agent |

---

## Secrets — never commit

The repo holds **no** AutoCount credentials and **no** n8n API key — those live in n8n only.
Do not hardcode or commit them. (The Apps Script shared secret `ABC` is a soft gate, not real
auth.)
