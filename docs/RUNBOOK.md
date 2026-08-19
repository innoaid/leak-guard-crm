# Runbook — Leak Guard CRM ops

How to deploy, rollback, and debug. Read end-to-end before your first patch.

---

## The patch script pattern (THE primary deploy mechanism)

Every n8n / Apps Script change goes through a single Python script that:
1. Fetches current workflow JSON via n8n API
2. Mutates the relevant Code node's `jsCode` (anchor-string find-and-replace)
3. PUTs back
4. Read-back verifies (no `\x08` backspace leaks, new code present, anchor absent)

### Skeleton

```python
# patch_xyz.py
import json, urllib.request, sys

API_KEY = sys.argv[1]
WID = '<workflow-id>'
BASE = 'https://leakguard.app.n8n.cloud/api/v1'

req = urllib.request.Request(f'{BASE}/workflows/{WID}',
                              headers={'X-N8N-API-KEY': API_KEY})
with urllib.request.urlopen(req) as r:
    wf = json.loads(r.read().decode('utf-8'))
node = next(n for n in wf['nodes'] if n['name'] == '<NodeName>')
code = node['parameters']['jsCode']

OLD = "..."   # exact substring you're replacing
NEW = "..."   # replacement
assert OLD in code, 'anchor missing — code shape changed'
node['parameters']['jsCode'] = code.replace(OLD, NEW)

put_body = {
    'name': wf['name'], 'nodes': wf['nodes'],
    'connections': wf['connections'],
    'settings': wf.get('settings') or {},
    'staticData': wf.get('staticData'),
}
req = urllib.request.Request(f'{BASE}/workflows/{WID}',
    data=json.dumps(put_body, ensure_ascii=False).encode('utf-8'),
    method='PUT',
    headers={'X-N8N-API-KEY': API_KEY, 'Content-Type': 'application/json'})
with urllib.request.urlopen(req) as r:
    print(f'PUT OK: {r.status}')

# Read-back verify
req = urllib.request.Request(f'{BASE}/workflows/{WID}',
                              headers={'X-N8N-API-KEY': API_KEY})
with urllib.request.urlopen(req) as r:
    wf2 = json.loads(r.read().decode('utf-8'))
code2 = next(n for n in wf2['nodes'] if n['name']=='<NodeName>')['parameters']['jsCode']
assert NEW in code2, 'NEW not in live'
assert code2.count(chr(8)) == 0, 'backspace leak'
print('VERIFY OK')
```

### Run it

```bash
python patch_xyz.py "$N8N_API_KEY"
```

### When NOT to use this pattern

- New WORKFLOW creation (use n8n UI or POST to `/api/v1/workflows`)
- Adding a brand-new NODE (paste the node JSON dict into `wf['nodes']`)
- Editing connections (mutate `wf['connections']` dict directly)

For both of those, see existing patches in `C:/Users/CY Lee/AppData/Local/Temp/n8n_v2/patch_*.py` for examples.

---

## Deploy a Code-node patch (live workflow)

1. **Pre-check**: `git status` clean? Owner approved? Test in sandbox first?
2. **Write** the patch script via Write tool (NOT inline heredoc — see CLAUDE.md gotcha #1)
3. **Test** via local synthetic exec or sandbox group first if possible
4. **Run** the patch
5. **Verify**: read latest exec, confirm new code path executed correctly
6. **Document**: add to `git log` with `Round NN: <description>` style

---

## Rollback a patch

Every patch script keeps OLD/NEW as constants. To revert:

1. Open the patch script
2. Swap OLD ↔ NEW in the call to `code.replace`
3. Re-run

For irreversible changes (e.g. a node added to workflow), manually delete the
node via the n8n UI or another patch script that pops it from `wf['nodes']`.

---

## Bounce a workflow

When you change `jsCode` in a Code node, n8n picks it up immediately. But:
- AI Agent Window Buffer Memory persists between executions
- Sometimes a workflow caches its compiled state

To force a fresh state:

```bash
curl -s -X POST -H "X-N8N-API-KEY: $N8N_API_KEY" \
  "https://leakguard.app.n8n.cloud/api/v1/workflows/<WID>/deactivate"
curl -s -X POST -H "X-N8N-API-KEY: $N8N_API_KEY" \
  "https://leakguard.app.n8n.cloud/api/v1/workflows/<WID>/activate"
```

This clears Window Buffer Memory (Annie's chat history etc).

---

## Inspect recent execs

List last N executions of a workflow:

```bash
curl -s -H "X-N8N-API-KEY: $N8N_API_KEY" \
  "https://leakguard.app.n8n.cloud/api/v1/executions?workflowId=<WID>&limit=10&includeData=false" \
  | python -m json.tool
```

Get full data of one exec (includes runData per node):

```bash
curl -s -H "X-N8N-API-KEY: $N8N_API_KEY" \
  "https://leakguard.app.n8n.cloud/api/v1/executions/<EID>?includeData=true" \
  | python -m json.tool > exec_<EID>.json
```

Then inspect node outputs:

```python
import json
d = json.load(open('exec_<EID>.json'))
nd = d['data']['resultData']['runData']
for k in nd:
    print(k, '|', nd[k][0].get('executionStatus'))
```

---

## Reset Annie's CRM (for sandbox testing)

```bash
PYTHONIOENCODING=utf-8 python "C:/Users/CY Lee/AppData/Local/Temp/reset_v2_test.py"
```

Clears Annie's Status, Cal Event ID, Pending fields. Lets you start a clean
test conversation.

---

## Smoke tests (run after every patch)

### lg-availability

```bash
curl -s "https://leakguard.app.n8n.cloud/webhook/lg-availability?phone=60179934386" \
  | python -c "import json,sys; d=json.load(sys.stdin); print('today:', d['availability'].get('2026-05-XX', 'NO_DATA'))"
```

Expected: real availability, not all-free, not all-blocked.

### Trigger a customer-join welcome (synthetic)

```bash
curl -X POST -H "Content-Type: application/json" \
  "https://leakguard.app.n8n.cloud/webhook/lg-customer-join" \
  -d '{"groupId":"120363426626363115@g.us","participants":[{"id":"60179934386@s.whatsapp.net"}]}'
```

Annie's group will get a welcome message. Don't fire on real customer groups.

### Fire a forced follow-up (no waiting for cron)

```bash
curl -X POST -H "Content-Type: application/json" -H "x-secret: ABC" \
  "https://leakguard.app.n8n.cloud/webhook/lg-followup-test" \
  -d '{"testGroup":"SV-0426-018 Annie - Ampang","force":true}'
```

---

## Whapi flakiness — what to do

Whapi cloud has sporadic timeouts (~5-10% of calls). Already mitigated:
- LG-Follow Up: 25s timeout + gate-on-sent (failed sends auto-retry next interval)
- v2 Send Whapi Reply: implicit retry via Whapi's own delivery layer

If you see `Whapi timeout 25s` errors in exec history:
- Single occurrences: ignore; auto-retry handles it
- Sustained pattern (>20% of recent execs): contact Whapi support
- Don't bump the timeout further — already at the practical max

---

## Apps Script redeploy

After editing `kanban_code.gs` in this repo:

1. Open the Apps Script editor (URL: ask owner)
2. Replace file contents with the new `kanban_code.gs`
3. **Deploy → Manage deployments → Edit (pencil) → Version: New version → Deploy**
4. Confirm Web app URL is unchanged (must end with `…WEhNPm7vrOYzZ6vFgQw-qA55Dv3mLB2Q/exec`)
5. Smoke test:
   ```bash
   curl -sL -X POST -H "Content-Type: text/plain" "<APPS_SCRIPT_URL>" \
     -d '{"action":"ping","secret":"ABC"}'
   ```
   Should return `{"status":"ok","pong":"..."}`

---

## DNS / hosting (post-Sunday-2026-05-04 migration)

Domain: `leakguard.my` is on GitHub Pages via custom domain.

To verify:
```
nslookup leakguard.my ns1.mschosting.cloud
```
Expect: `185.199.10[8-9].153` or `185.199.11[0-1].153`.

If returning `103.6.196.47` (Plesk IP), DNS replication broke; contact Exabytes
support.

---

## Common tripwires

| Symptom | Probable cause | First check |
|---|---|---|
| Bot stays silent in customer group | sender phone ≠ lead phone (round 37 gate) | exec → Verify Lead skipReason |
| All slots show free for today | calendar API errored, fail-closed didn't fire | LG-Availability latest exec |
| All slots show blocked | timezone bug returned (round 34) | check Prep Time Range jsCode |
| Day-name in bot reply wrong | gpt-4o slip; round 38 regex post-processor catches | check exec's reply field |
| Customer says "yes" but bot re-asks | pendingValid false; setPending didn't write | check Pending Date in CRM |
| Group renamed but Status didn't change | "Status (I)" vs "Status" header mismatch | check Update Group Name CRM node |

Each row maps to a gotcha doc in `docs/gotchas/`.

---

## Memory / state files

- n8n cloud workspace: lives in n8n's cloud DB; can be exported via API
- Apps Script: deployed via Google's infra; source-of-truth lives in repo
- CRM sheet: 441 rows currently; backed up by Google Sheets revision history
- Calendar events: live in Google Calendar; can't easily back up

For "true rollback" of a multi-piece change, keep the patch scripts AND tag
the git commit with `Round NN`.
