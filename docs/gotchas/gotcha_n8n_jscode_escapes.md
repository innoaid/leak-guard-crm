---
name: n8n jsCode escape sequences via API
description: Patches that introduce JS escape sequences (\b, \n, \t, \r, \u...) into Code-node jsCode must be sent via Write-tool-created Python file, NEVER via inline `python << 'PY'` heredoc — the heredoc collapses escapes and breaks the regex/string literal in the live n8n state.
type: feedback
originSessionId: 0fffd8df-1d53-41af-a48d-dd528f0a4923
---
When patching n8n Code-node `parameters.jsCode` via the n8n API, JSON-encoding through Python's json.dumps preserves backslash-escapes correctly — BUT only if the Python source string contains literal backslash+letter, not the escape's interpreted form.

**Failure mode:** writing the patch script as an inline `bash python << 'PY' ... PY` heredoc results in `\b` (word boundary regex) being stored as ASCII 0x08 backspace in the live jsCode, and `\n` (newline regex/string escape) being stored as actual newline (LF) — breaking the regex or splitting the string literal across multiple lines.

**Why:** between the heredoc body, bash's `<<'PY'` quoting, and Python's string literal parsing, the double-backslash `\\b` somewhere becomes a single-backslash `\b` (interpreted as backspace by Python's literal grammar). When that 1-char backspace flows through json.dumps → JSON `"\b"` → n8n's JSON parser → backspace is preserved (per JSON spec).

**Confirmed examples:**
- Round 14/16 patches put `\b` in regex source, live state had `\x08` instead → `_CONFIRM_AT_START` and `_agentPickedTime` regexes never matched (round 17 fix used `(?=\s|$|[.,!?])` lookahead instead).
- Round 20 patch put `\n` inside a single-quoted JS string, live state had actual newlines → SyntaxError on every Verify Lead execution, bot completely dead (round 21 fix used backtick template literal + Write-tool patch script).
- Round 31 (2026-04-30): two more `\b → \x08` instances surfaced. (a) LG-Admin Commands `Gen Seq Number` regex `/(\bJB\b|JOHOR)/i` had become `/(\x08JB\x08|JOHOR)/i` — every JB lead's `-admin <phone>` rename produced SV- prefix instead of SVJB-. (b) LG-Availability `Build Availability` had `/\bJB\b/.test(location)` → `/\x08JB\x08/`. Both fixed via Write-tool Python patch using lookaround pattern `(?:^|[^A-Z])JB(?:[^A-Z]|$)` — completely escape-free regex body, immune to future heredoc collapse.

**How to apply:**
- Always use the **Write tool** to create a `.py` file with the patch contents, then run it via Bash. The Write tool writes UTF-8 bytes verbatim — no shell or quoting layer in between.
- Avoid `python << 'PY' ... PY` heredocs whenever the patch contains JS escape sequences inside string literals or regex source.
- After PUT, read back the live jsCode via API and verify with `code.count('\x08') == 0` AND inspect the regex/string with `repr()` to confirm `\b`/`\n` survived as 2-char `\b`/`\n` (not 1-char backspace/newline).
