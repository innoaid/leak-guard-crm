---
name: CLAUDE.md sheet columns and workflow IDs are stale
description: Project CLAUDE.md has stale CRM column letters and typo'd workflow IDs — verify against live data before relying on them
type: project
originSessionId: acc90124-d214-4973-b213-a7cdc5dbb905
---
CLAUDE.md (project doc at C:\projects\n8n-workflows\CLAUDE.md) is partially stale as of 2026-04-25. Two specific drifts:

**1. CRM sheet has shifted columns** — a new column A "Handled by" was inserted, shifting everything. Real letters now (verified via gviz `select * limit 1`):
- A = Handled by (NEW)
- B = Timestamp, C = Phone, D = Name, E = Problem Type, F = Location, G = Full Address, H = Slab Size, I = Slot Chosen, J = Status (CLAUDE.md says I)
- AA = Cal Event ID (legacy), AB = Tags, AC = Group ID, AD = Flow Stage, AE = Last Bot Msg Time, AF = Group Name, AI = Cal Event ID (correct one)
- Headers retain old letters in parens, e.g. `'Group ID (AB)'` is now in column AC. n8n workflows still work because they read by header NAME, not index.

**2. Workflow IDs in CLAUDE.md have 1/l typos**. Real IDs (verified via `/api/v1/workflows`):
- LG - Customer Chat: `Slbb6OljpwBuHjB2` (CLAUDE.md says `S1bb6O1jpwBuHjB2` — wrong)
- LG - WA Receiver: `gpnJPMa9w5FX1fDM` (CLAUDE.md says `gpnJPMa9w5FXlfDM` — wrong)
- LG - Admin Commands: `PILR7VRalAqggy5P` (CLAUDE.md says `PILR7VRa1Aqqgy5P` — wrong)
- LG - Customer Appointment Reminder: `0l3d6TdqoFPNkUGU` (CLAUDE.md says `013d6TdqoFPNkUGU` — wrong)

**Why:** Manual transcription confused `1` with `l` and `O` with `0`. CLAUDE.md was last touched April 2026 and hasn't been re-verified since.

**How to apply:** When CLAUDE.md gives you a column letter or workflow ID, don't trust it for API calls. Look up the live value first — list workflows via `/api/v1/workflows` for IDs, or `gviz select * limit 1` on the sheet for column labels. Always read CRM cells by header name (e.g. `row.json['Group ID (AB)']`) not by index.
