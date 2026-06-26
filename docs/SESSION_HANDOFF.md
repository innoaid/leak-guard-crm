# Session Handoff — state & open items

_Last updated: 2026-06-27. Read this with `README.md`, `CLAUDE.md`, and `docs/ARCHITECTURE.md` to recall the project._

## Where things are
- **All code is on `main`** (this is the source of truth). Repo = Apps Script `.gs` + the HTML UIs.
- **Deploy model (important — two targets):**
  - `*.html` → **GitHub Pages**, auto-deploys on push to `main` (served at `leakguard.my`). `.nojekyll` is in the repo; hard-refresh to bust the ~10-min CDN cache.
  - `*.gs` → **paste into the Apps Script editor → Save → Deploy → Manage deployments → New version**. Pushing `.gs` to GitHub does NOT make the live Sheet run it.
  - **n8n** workflows are edited live (cloud) and backed up locally only.
- **Read live CRM data with no credentials** (handy for diagnosis): the public gviz endpoint
  `https://docs.google.com/spreadsheets/d/1FnuiZcOSy5UMQpW81I7qtU6a7NGlnHtJbH2EkVM7PLQ/gviz/tq?tqx=out:json&sheet=Leak%20Guard%20Leads`
  (strip the `google.visualization.Query.setResponse(...)` wrapper; parse the JSON).

## ⚠️ Pending Apps Script deploy (do this to activate the latest `.gs`)
Re-paste `kanban_code.gs` into the Apps Script editor and deploy a **New version**. The live
deployment is likely behind; the current file includes:
- **Ad Analysis Report** endpoint (`adReport`) — the kanban 📈 Ad Report button needs it.
- **Express phone normalization** (`_normMyPhone`) — booking phones stored canonical `60…`.
- **`expressLead`** upsert + serviceable-state routing.
- **Pending-group tag-clear** in `handleBulkLinkGroups`.

## n8n changes applied LIVE this session (no redeploy needed; backed up at `C:\projects\leak-guard-crm-n8n-backup`)
- **LG - Customer Join** — greet express leads on join + don't downgrade booked leads.
- **LG - Booking** — `Update CRM` matches the found row's stored phone (`matchedPhone`) + `alwaysOutputData` so a no-match never stalls the flow (fixes "network error" + duplicate calendar events).
- **LG - Bulk Link Groups** — returns `participants`; honors a `chosenPhone` (kanban member picker).
- **LG - Admin Commands** — `-admin silent` actually silent now (read flag from Route Admin); `Find Lead` prefers the named / furthest-in-funnel row over blank duplicates.
- **LG - Quotation Create** — `companyName` falls back when the lead name is blank (fixes AutoCount "CompanyName field is required" 400).

## Outstanding manual tasks
1. **Deploy `kanban_code.gs`** (New version) — see above.
2. **Delete orphaned Google Calendar events** (from the old retry-cascade): 4× `6012247174` (4 Jul), 3× `0183639951`, 1× `012552552255` (test). No CRM link.
3. **Re-sync Philip Yew's AutoCount QT** (60123284898) — re-send the estimation or use kanban "Sync to AutoCount"; the CompanyName fix lets it generate now + rename the group with the `QT-` prefix.
4. 🔒 **Rotate the Whapi token** — it's hardcoded in `kanban_code.gs` (line ~29) in this **public** repo, so it's exposed. Move to Script Properties.
5. **Add the express link** `https://leakguard.my/express/` to the WA first-engagement message (n8n bot text) so leads can choose express vs chatbot.
6. _Optional:_ assign the two pre-feature unassigned SVC leads (Mr Kean SV-Seremban, Stanley Tan).

## Decisions on file
- **Single folder + `main` only** — work in `C:\projects\leak-guard-crm`; the `round-69-qt-pdf` branch/worktree was retired.
- **MongoDB migration** — assessed; chose "just assess for now." It's a system-wide job (~24 n8n workflows / 95 Sheets nodes + a new API layer), best done staged; no rush at current scale (~1,500 leads).
- **Recurring root cause to watch:** duplicate phone rows + blank-name leads break phone-based matching/AutoCount. Matchers were hardened this session; de-duplicating leads at intake would remove the root cause.

## Not on GitHub (machine-local — won't transfer between machines)
- Claude Code's saved memory (`~/.claude/.../memory/`).
- The local n8n backups folder.
- The n8n API key (kept out of the repo by design).
