---
name: CRM headers carry column-letter suffixes
description: Several CRM headers include their original column letter in parentheses; using the bare name in n8n Sheets nodes is a silent no-op
type: project
originSessionId: acc90124-d214-4973-b213-a7cdc5dbb905
---
Multiple CRM headers in `Leak Guard Leads` carry a `(<col-letter>)` suffix in the header text itself, not just as a comment. Confirmed examples:
- `Follow Up Count (AK)`
- `Last Follow Up At (AL)`
- `Group ID (AB)`
- `Group Name (AE)`
- `Last Bot Msg Time (AD)`
- `Pending Date (AF)`, `Pending Slot (AG)`

**Why:** Sheet was originally hand-built with column letters embedded in headers as documentation. Even though the sheet has since had columns inserted/shifted, the header strings retain the original letters. Existing workflows (e.g., LG - Follow Up's `Write Follow Up to CRM`) use the suffixed form because it's the literal cell value.

**How to apply:**
- When configuring n8n Google Sheets node `update` operation, the column key in `columns.value` must match the literal header text. Using `'Follow Up Count'` instead of `'Follow Up Count (AK)'` causes the update to silently no-op (execution shows `success` but no cell changes).
- When reading via gviz or n8n Sheets read, access keys use the literal header: `row.json['Follow Up Count (AK)']`.
- A few headers don't have the suffix (e.g., `Phone`, `Name`, `Status`, `Tags`). When in doubt, fetch headers via `gviz tq?...&headers=1` or copy the exact key from a working workflow's parameters.
