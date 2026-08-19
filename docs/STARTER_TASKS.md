# Starter tasks — first 5 things to ship

These tasks are sized to take 1-3 hours each. Each touches a different part of
the system, so by task 5 you'll have hands-on with most of the codebase. Pair
with owner on task 1; tasks 2+ should be solo with PR review.

Goal: get from "read docs" to "shipping confidently" within ~1 week.

---

## Task 1: Read + reproduce a known bug

**What:** Pick any gotcha from `docs/gotchas/` (recommend
`gotcha_n8n_jscode_escapes.md` — most common pitfall). Read the failure
mode. Then reproduce it locally:

1. Write a test patch script that uses inline `python << 'PY'` heredoc
2. Have it write `\b` into a JS regex
3. Run it on a test/clone workflow (NOT live; ask owner to clone Verify Lead
   for you with `[DEV-TEST]` suffix)
4. Inspect the live state — confirm `\x08` shows up
5. Then redo via Write-tool .py file pattern, confirm `\b` survives

**Deliverable:** screenshot of both states + a 1-paragraph writeup posted as a
PR comment.

**Why this task:** the heredoc gotcha is the #1 cause of bugs we've fixed.
You'll never write inline heredoc patches again after seeing it firsthand.

---

## Task 2: Add a small Apps Script handler

**What:** Add a `setLocation` Apps Script handler to `kanban_code.gs`.
Accepts `{action:'setLocation', secret, phone, location}` and writes to the
Location column for the matching lead.

**Why:** lets us update Location later when customers move/clarify. Currently
no way to do this from kanban without manual sheet edit.

**How:**
1. Read `kanban_code.gs` — see how `handleSetPending` is structured
2. Mirror that pattern for `handleSetLocation`
3. Add `case 'setLocation': return handleSetLocation(body);` to `doPost` switch
4. Use the existing `setCellByHeader(sheet, rowNum, 'Location', body.location)` helper
5. Open a PR — owner reviews + redeploys Apps Script

**Test:**
```bash
curl -sL -X POST -H "Content-Type: text/plain" "$APPS_SCRIPT_URL" \
  -d '{"action":"setLocation","secret":"ABC","phone":"60179934386","location":"Petaling Jaya"}'
```

Expected: `{"status":"ok",...}` and Annie's Location cell updates.

---

## Task 3: Add a kanban column for "Days in stage"

**What:** The kanban already shows `Days in stage` per card (computed from
`Status Changed At`). But there's no SORT or FILTER by it. Add a "Stale leads"
filter button that shows only cards where `Days in stage > 7`.

**Why:** helps spot leads that are sitting too long in one phase.

**How:**
1. Open `team_kanban.html` (~3000 lines but well-commented)
2. Find the existing filter bar (search for `id="redFilter"` from round 27)
3. Add a new button or filter option
4. Hook into `passesFilters(lead)` to add the new condition
5. Test by opening the file directly in browser

**Test:** open `file:///C:/Projects/leak-guard-crm/team_kanban.html` (with
fresh CRM data), tick "Stale", verify only old cards show.

**Deploy:** push to main, GitHub Pages auto-rebuilds.

---

## Task 4: Write a unit test for `_isJbWord`

**What:** The JB detection regex (in LG-Admin Commands `Gen Seq Number`) has
been buggy twice. Write a small Python test harness that validates the
deployed regex behaves correctly across edge cases.

**Why:** prevents future regression. Pattern is reusable for other regex
fixes.

**How:**
1. Fetch the live regex via the n8n API
2. Translate to Python (or use `js2py`)
3. Test cases:
   - `"JB"` → match
   - `"jb"` → match (case insensitive)
   - `"Johor"` → match
   - `"JBHotel"` → no match (must be word)
   - `"Klang Valley JB"` → match (word boundary)
   - empty string → no match
4. Save to `tests/test_jb_regex.py`

**Deliverable:** test file + run output showing all pass.

---

## Task 5: Ship the FB-funnel Step 1 (welcome branching)

**What:** Audit recommended this as the first concrete funnel step
(`audit_funnel_v1.md` section H). Welcome message branches based on whether
the customer already has a Cal Event ID at join time.

**Why:** the larger funnel project depends on this; it's also a great
self-contained task to learn the patch-script pattern.

**How:** plan is already written in `audit_funnel_v1.md` section H. Concrete
anchors:
- File: LG-Customer Join (`gJAArg6iDord57Ua`) → `Build Welcome with Slots`
  jsCode
- ALSO extend `Match Customer` to pass `calEventId` and `slotChosen` through

**Steps:**
1. Write the patch script
2. Test with synthetic webhook fire on Annie's group AFTER manually setting
   her Cal Event ID via Apps Script setPending
3. Reset Annie afterwards
4. Open PR — owner reviews + merges + you run patch on live

**Bonus:** also update LG-Resend Welcome with the same branching logic (same
welcome template).

---

## When you're done with all 5

You've touched:
- Apps Script (Task 2)
- Kanban frontend (Task 3)
- Regex testing tooling (Task 4)
- n8n workflow patches via patch script (Tasks 1, 5)

You're ready for the bigger projects:
- Phase 1: Full kanban automation (5-7 weeks)
- Phase 2: FB ad funnel Steps 2-7 (2-3 weeks)

Ask owner to scope your next task based on priority.
