# Plan: Push handoff artefacts to GitHub for laptop resume

## Context

User wants to switch to laptop. Confirmed by direct check: neither `HANDOFF.md` (the current plan snapshot) nor `site_visit_mockup.html` (parked Site Visit Form mockup v2) are in the repo yet. Both need to be in GitHub so the laptop session can read them via `git clone` without any extra file transfer.

## Actions

1. Copy `C:/Users/CY Lee/.claude/plans/read-the-memory-and-squishy-shannon.md` → `C:/Projects/leak-guard-crm/HANDOFF.md`.
2. Copy `C:/Users/CY Lee/Desktop/site_visit_mockup.html` → `C:/Projects/leak-guard-crm/mockups/site_visit_mockup.html`.
3. `git add` both, commit, push.

**Where to put the mockup**: under a new `mockups/` subfolder so it's clearly not production code (the demo banner inside the file already says it's a mockup). On GitHub Pages it'd also be accessible at `https://leakguard.my/mockups/site_visit_mockup.html` if user wants to test directly on phone via real URL — bonus.

## Hotfix from earlier (already shipped, kept for reference)

Round 51 archive-view null-ref fix is committed `522a969` and pushed.

---

# Plan: Hotfix — Archive view renders empty due to null-ref in render()

## Context

User archived `SV-0426-016 Zach` to status `Rejected` via the new 🗄 Archive button (Round 43). gviz confirms CRM is correctly updated — Status=`Rejected`, `archived` tag set. But clicking 📦 Archive in the header shows an empty list.

## Root cause

`team_kanban.html` lines 2927-2935 — the archive-toggle handler replaces `btnArchive.innerHTML` with the literal string `'← Back to Kanban'` (no `<span id="archCount">` inside). Then `render()` runs:

```js
const archCount = allLeads.filter(...).length;
document.getElementById('archCount').textContent = archCount;  // ← throws TypeError
```

`getElementById('archCount')` returns `null` because the span was just removed. `.textContent` on null throws → `render()` aborts → `renderArchive()` is never called → archive list stays empty.

Symmetric bug exists when toggling BACK to kanban: line 2933 reads `document.getElementById('archCount').textContent` to preserve the count when rebuilding the button, but in archive mode the span doesn't exist → another null-ref.

## Fix

Two-line null-guard in `team_kanban.html`:

1. Line 726 (in `render()`): wrap the textContent assignment in an existence check.
2. Line 2933 (in `btnArchive.onclick`): cache the count via `allLeads` re-derivation when rebuilding the button, instead of reading from the (potentially missing) span.

Cleaner alternative: keep `archCount` as a top-level variable updated in `render()`, button rebuild reads from variable. Apply the cleaner version since it future-proofs against more null refs.

## File

| File | Change |
|---|---|
| `C:/Projects/leak-guard-crm/team_kanban.html` | Two-section edit: make archCount tolerant of the span being absent in archive mode. |

## Verification

1. Hard-refresh kanban → click 📦 Archive → Zach (`SV-0426-016`) appears with `Rejected` badge.
2. Click ← Back to Kanban → kanban view restored with badge count = 1.
3. Archive flow end-to-end with another card → after toast, click 📦 Archive → new card appears.
4. No console errors.

## Out of scope

The Site Visit Form mockup work is parked at `C:/Users/CY Lee/Desktop/site_visit_mockup.html` — pending user verdict on UX before backend wiring. Resume after this hotfix.
