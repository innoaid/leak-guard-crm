---
name: n8n LangChain HTTP Tool body modes
description: jsonBody and bodyParameters keypair handle expressions vs LLM-fillable params differently — must use the right pattern
type: feedback
originSessionId: acc90124-d214-4973-b213-a7cdc5dbb905
---
n8n LangChain `toolHttpRequest` body modes are mutually exclusive on what they substitute:
- **Raw template jsonBody (no `={{}}` wrap):** placeholders `{name}` substituted from LLM, but inline `{{ n8n expression }}` blocks NOT evaluated — they pass through as literal text
- **Expression jsonBody (`={{ JSON.stringify({...}) }}` wrap):** evaluates the whole expression; `$fromAI()` calls inside DO substitute LLM values; placeholders `{name}` NOT recognized
- **bodyParameters keypair format:** ✅ correct way to mix static + LLM-filled:
  - Each entry needs explicit `valueProvider` field
  - `'fieldValue'` → server-side n8n expression (NOT in LLM schema)
  - `'modelRequired'` or `'modelOptional'` → LLM fills (IN schema, value uses `$fromAI('name', 'desc', 'type')`)

**Why:** discovered 2026-04-29 when LG-Booking received body with literal `{{ $('Verify Lead').item.json.phone }}` strings instead of resolved values. Calendar event title was built from JSON gibberish. Tried multiple body modes; only keypair with valueProvider works for mixed static/LLM fields.

**How to apply:** always use bodyParameters keypair for tools that combine context (Verify Lead state) with LLM-decided fields (date/slot/reason). Never mix `{{ }}` and `{name}` in raw jsonBody — pick one. Phantom node connections also crash agents — when renaming a tool node, also update its connection key in `wf['connections']`.
