// ================================================================
// LEAK GUARD KANBAN — Apps Script Web App
// ================================================================
// Standalone Apps Script project. Independent of sync.gs / wa_admin_bot.gs.
// Reads/writes the LIVE production sheet using HEADER NAMES (immune to
// column-letter shifts).
//
// SETUP (one-time):
//   1. https://script.google.com/home → New Project → Name: "Leak Guard Kanban"
//   2. Replace Code.gs with this file's contents
//   3. Save → Deploy → New deployment → type: Web app
//      - Execute as:   Me
//      - Who has access: Anyone
//   4. Copy the Web app URL → paste into team_kanban.html `WEBAPP_URL`
//   5. After any change here, redeploy: Deploy → Manage deployments → Edit → New version → Deploy
//
// SECURITY:
//   Soft secret in body.secret — must match SHARED_SECRET below.
//   Real auth (Google Sign-In + email allowlist) is v2.
// ================================================================

const LIVE_SHEET_ID = '1FnuiZcOSy5UMQpW81I7qtU6a7NGlnHtJbH2EkVM7PLQ';
const SHEET_NAME    = 'Leak Guard Leads';
const SHARED_SECRET = 'ABC'; // matches team_kanban.html
const N8N_BOOKING_URL = 'https://leakguard.app.n8n.cloud/webhook/lg-booking';
const N8N_WAGROUP_URL = 'https://gate.whapi.cloud/messages/text';
const WHAPI_TOKEN     = 'tjJeSotqcmnYBfQulcRxcFHHQ8QtDcC5';

// ================================================================
// doPost — single entry point. Routes by body.action.
// ================================================================

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    if (body.secret !== SHARED_SECRET) {
      return jsonResponse({status: 'error', message: 'unauthorized'});
    }

    switch (body.action) {
      case 'updateStatus':       return handleUpdateStatus(body);
      case 'updateTag':          return handleUpdateTag(body);
      case 'updateAssignee':     return handleUpdateAssignee(body);
      case 'updateNotes':        return handleUpdateNotes(body);
      case 'updateQuotation':    return handleUpdateQuotation(body);
      case 'archiveLead':        return handleArchive(body);
      case 'restoreLead':        return handleRestore(body);
      case 'sendRescheduleLink': return handleSendReschedule(body);
      case 'createLead':         return handleCreateLead(body);
      case 'resetTestLead':      return handleResetTestLead(body);
      case 'setPending':         return handleSetPending(body);  // round 12 — v2 agent pending-confirmation slot
      case 'bulkMoveStatus':     return handleBulkMoveStatus(body);  // round 32 — kanban bulk move-to-phase
      case 'updateLeadDetails':  return handleUpdateLeadDetails(body);  // task 2 — kanban edit-lead modal
      case 'cancelAppointment':  return handleCancelAppointment(body);  // round 45 — SVC -> PSV via kanban appt modal
      case 'bulkLinkGroups':     return handleBulkLinkGroups(body);  // round 48 — bulk-link pre-CRM WA groups
      case 'claimCooldown':      return handleClaimCooldown(body);  // round 54 — atomic check-and-set for v2 link cooldown
      case 'ping':               return jsonResponse({status: 'ok', pong: new Date().toISOString()});
      default:
        return jsonResponse({status: 'error', message: 'unknown action: ' + body.action});
    }
  } catch (err) {
    Logger.log('doPost error: ' + err.toString());
    return jsonResponse({status: 'error', message: err.toString()});
  }
}

// ================================================================
// Helpers
// ================================================================

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  return SpreadsheetApp.openById(LIVE_SHEET_ID).getSheetByName(SHEET_NAME);
}

// Returns {headers: string[], colByName: {name -> 1-based col index}}
function getHeaders(sheet) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function(h) { return String(h || '').trim(); });
  const colByName = {};
  for (let i = 0; i < headers.length; i++) {
    if (headers[i]) colByName[headers[i]] = i + 1;
  }
  return {headers: headers, colByName: colByName};
}

// Find lead row by phone (Phone column). Returns 1-based row number, or null.
function findRowByPhone(sheet, phone) {
  if (!phone) return null;
  const target = String(phone).trim();
  const last8 = target.length >= 8 ? target.slice(-8) : target;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const phoneCol = headers.indexOf('Phone');
  if (phoneCol === -1) return null;

  for (let i = 1; i < data.length; i++) {
    const cell = String(data[i][phoneCol] || '').trim();
    if (cell === target) return i + 1;
    if (cell.length >= 8 && cell.endsWith(last8)) return i + 1;
  }
  return null;
}

// Centralized cell update by header name
function setCellByHeader(sheet, rowNum, headerName, value) {
  const h = getHeaders(sheet);
  const col = h.colByName[headerName];
  if (!col) throw new Error('Header not found: ' + headerName);
  sheet.getRange(rowNum, col).setValue(value);
}

// ================================================================
// Action handlers
// ================================================================

function handleUpdateStatus(body) {
  // body: {action, phone, status, changedBy, secret}
  const sheet = getSheet();
  const rowNum = findRowByPhone(sheet, body.phone);
  if (!rowNum) return jsonResponse({status: 'error', message: 'lead not found'});

  const h = getHeaders(sheet);
  const statusCol    = h.colByName['Status'];
  const changedAt    = h.colByName['Status Changed At'];
  const changedBy    = h.colByName['Changed By'];

  if (!statusCol) return jsonResponse({status: 'error', message: 'Status column not found'});

  sheet.getRange(rowNum, statusCol).setValue(body.status);
  if (changedAt) sheet.getRange(rowNum, changedAt).setValue(new Date().toISOString());
  if (changedBy) sheet.getRange(rowNum, changedBy).setValue(body.changedBy || 'Kanban');

  return jsonResponse({status: 'ok', rowNum: rowNum});
}

function handleUpdateTag(body) {
  // body: {phone, tags, secret}  — tags is comma-separated string
  const sheet = getSheet();
  const rowNum = findRowByPhone(sheet, body.phone);
  if (!rowNum) return jsonResponse({status: 'error', message: 'lead not found'});

  setCellByHeader(sheet, rowNum, 'Tags', body.tags || '');
  return jsonResponse({status: 'ok'});
}

function handleUpdateAssignee(body) {
  // body: {phone, assignee, secret}
  const sheet = getSheet();
  const rowNum = findRowByPhone(sheet, body.phone);
  if (!rowNum) return jsonResponse({status: 'error', message: 'lead not found'});

  setCellByHeader(sheet, rowNum, 'Assigned To', body.assignee || '');
  return jsonResponse({status: 'ok'});
}

function handleUpdateNotes(body) {
  // body: {phone, notes, secret}
  const sheet = getSheet();
  const rowNum = findRowByPhone(sheet, body.phone);
  if (!rowNum) return jsonResponse({status: 'error', message: 'lead not found'});

  setCellByHeader(sheet, rowNum, 'Notes', body.notes || '');
  return jsonResponse({status: 'ok'});
}

function handleUpdateQuotation(body) {
  // body: {phone, quotation, secret}
  const sheet = getSheet();
  const rowNum = findRowByPhone(sheet, body.phone);
  if (!rowNum) return jsonResponse({status: 'error', message: 'lead not found'});

  setCellByHeader(sheet, rowNum, 'Quotation (RM)', body.quotation || '');
  return jsonResponse({status: 'ok'});
}

function handleArchive(body) {
  // body: {phone, archiveStatus, archiveNote, secret}
  //   archiveStatus = Lost / Cold Lead / Rejected / Out of Area / Human Handoff
  //   archiveNote   = optional admin reason; appended to Notes with date prefix
  const archiveStatus = body.archiveStatus || 'Lost';
  const sheet = getSheet();
  const rowNum = findRowByPhone(sheet, body.phone);
  if (!rowNum) return jsonResponse({status: 'error', message: 'lead not found'});

  // 1) Status flip + audit cols (reuses handleUpdateStatus path)
  handleUpdateStatus({
    phone: body.phone,
    status: archiveStatus,
    changedBy: body.changedBy || 'Kanban (archive)'
  });

  // 2) Tag with `archived` so v2 chat bot stays silent (defense-in-depth)
  try {
    const existingTags = String(sheet.getRange(rowNum, getHeaders(sheet).colByName['Tags'] || 0).getValue() || '').trim();
    const tagSet = new Set(existingTags.split(',').map(s => s.trim()).filter(Boolean));
    tagSet.add('archived');
    setCellByHeader(sheet, rowNum, 'Tags', Array.from(tagSet).join(','));
  } catch (_) {}

  // 3) Append archive note to Notes column
  if (body.archiveNote && String(body.archiveNote).trim()) {
    try {
      const existing = String(sheet.getRange(rowNum, getHeaders(sheet).colByName['Notes'] || 0).getValue() || '').trim();
      const datePrefix = new Date().toISOString().slice(0, 10);
      const entry = `[${datePrefix}] Archived (${archiveStatus}): ${String(body.archiveNote).trim()}`;
      const merged = existing ? (existing + '\n' + entry) : entry;
      setCellByHeader(sheet, rowNum, 'Notes', merged);
    } catch (_) {}
  }

  return jsonResponse({status: 'ok', rowNum: rowNum, archivedAs: archiveStatus});
}

function handleRestore(body) {
  // body: {phone, restoreToStatus, secret}
  return handleUpdateStatus({
    phone: body.phone,
    status: body.restoreToStatus || 'Pending Site Visit',
    changedBy: body.changedBy || 'Kanban (restore)'
  });
}

function handleSendReschedule(body) {
  // body: {phone, groupId, secret}
  // Sends a reschedule link to the lead's WhatsApp group via Whapi.
  // Pulls existing slot/calEventId from sheet for the bare URL.
  const sheet = getSheet();
  const rowNum = findRowByPhone(sheet, body.phone);
  if (!rowNum) return jsonResponse({status: 'error', message: 'lead not found'});

  const h = getHeaders(sheet);
  const row = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
  const get = function(name) { return h.colByName[name] ? row[h.colByName[name] - 1] : ''; };

  const phone     = String(get('Phone') || '').trim();
  const name      = String(get('Name') || '').trim();
  const groupName = String(get('Group Name (AE)') || '').trim();
  const groupId   = body.groupId || String(get('Group ID (AB)') || '').trim();

  if (!groupId) return jsonResponse({status: 'error', message: 'no group ID'});

  // Build the bare URL — booking page will detect existingAppt via /availability
  const params = [];
  if (phone) params.push('phone=' + encodeURIComponent(phone));
  if (name) params.push('name=' + encodeURIComponent(name));
  if (groupName) params.push('group=' + encodeURIComponent(groupName));
  const longUrl = 'https://innoaid.github.io/leak-guard-crm/booking.html?' + params.join('&');

  const msg = 'Click the link below to manage your appointment — only takes 2 mins.\n\n' +
    '🚀 Express Booking: ' + longUrl;

  try {
    UrlFetchApp.fetch(N8N_WAGROUP_URL, {
      method: 'post',
      contentType: 'application/json',
      headers: {'Authorization': 'Bearer ' + WHAPI_TOKEN},
      payload: JSON.stringify({to: groupId, body: msg, typing_time: 2}),
      muteHttpExceptions: true
    });
    return jsonResponse({status: 'ok', url: longUrl});
  } catch (err) {
    return jsonResponse({status: 'error', message: 'whapi: ' + err.toString()});
  }
}

// ================================================================
// Create New Lead (kanban "+ New Lead" button)
// ================================================================

// Returns ALL rows matching phone (last-8 digits), each {rowNum, status}.
function findAllMatchesByPhone(sheet, phone) {
  if (!phone) return [];
  const target = String(phone).replace(/\D/g, '');
  if (target.length < 8) return [];
  const last8 = target.slice(-8);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const phoneCol = headers.indexOf('Phone');
  const statusCol = headers.indexOf('Status');
  if (phoneCol === -1) return [];
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const cell = String(data[i][phoneCol] || '').replace(/\D/g, '');
    if (cell.length >= 8 && cell.endsWith(last8)) {
      out.push({rowNum: i + 1, status: String(data[i][statusCol] || '').trim()});
    }
  }
  return out;
}

function handleCreateLead(body) {
  // body: {action, name, phone, status, source, location, address, problemType,
  //        notes, assignedTo, allowDuplicate, changedBy, secret}
  if (!body.name || !body.phone || !body.status) {
    return jsonResponse({status: 'error', message: 'name, phone, status required'});
  }
  const ACTIVE = ['New Lead','Pending Invitation','Pending Site Visit','Site Visit Confirmed',
                  'Pending QT','Quotation Sent','Follow Up','Pending I.Date',
                  'I.Date Confirmed','Job In Progress'];
  const sheet = getSheet();

  // Server-side duplicate gate (defence-in-depth — frontend already filters)
  const matches = findAllMatchesByPhone(sheet, body.phone);
  const activeMatch = matches.find(function(m) { return ACTIVE.indexOf(m.status) !== -1; });
  if (activeMatch && !body.allowDuplicate) {
    return jsonResponse({
      status: 'error', code: 'active_duplicate',
      existingRow: activeMatch.rowNum, existingStatus: activeMatch.status
    });
  }

  const newRow = sheet.getLastRow() + 1;
  const nowIso = new Date().toISOString();
  const todayStr = nowIso.slice(0, 10);

  setCellByHeader(sheet, newRow, 'Timestamp', nowIso);
  setCellByHeader(sheet, newRow, 'Phone', body.phone);
  setCellByHeader(sheet, newRow, 'Name', body.name);
  setCellByHeader(sheet, newRow, 'Status', body.status);
  setCellByHeader(sheet, newRow, 'Status Changed At', nowIso);
  setCellByHeader(sheet, newRow, 'Changed By', body.changedBy || 'Kanban (create)');
  setCellByHeader(sheet, newRow, 'Date Lead In', todayStr);
  setCellByHeader(sheet, newRow, 'Source', body.source || 'Other');
  if (body.location)    setCellByHeader(sheet, newRow, 'Location', body.location);
  if (body.address)     setCellByHeader(sheet, newRow, 'Full Address', body.address);
  if (body.problemType) setCellByHeader(sheet, newRow, 'Problem Type', body.problemType);
  if (body.notes)       setCellByHeader(sheet, newRow, 'Notes', body.notes);
  if (body.assignedTo)  setCellByHeader(sheet, newRow, 'Assigned To', body.assignedTo);

  return jsonResponse({status: 'ok', rowNum: newRow});
}

// ================================================================
// Reset Test Lead — wipes chat-state CRM fields for v2 agent sandbox
// ================================================================

function handleResetTestLead(body) {
  // body: {action, groupName, secret, deleteCalEvent? (default false)}
  if (!body.groupName) {
    return jsonResponse({status: 'error', message: 'groupName required'});
  }
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const gnIdx = headers.indexOf('Group Name (AE)');
  if (gnIdx === -1) return jsonResponse({status: 'error', message: 'Group Name (AE) header missing'});

  let rowNum = null;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][gnIdx] || '').trim() === body.groupName.trim()) {
      rowNum = i + 1;
      break;
    }
  }
  if (!rowNum) return jsonResponse({status: 'error', message: 'group not found: ' + body.groupName});

  const calEventIdIdx = headers.indexOf('Cal Event ID (AH)');
  const oldCalEventId = calEventIdIdx >= 0 ? String(data[rowNum-1][calEventIdIdx] || '').trim() : '';

  // Fields to blank — chat lifecycle state
  const fieldsToBlank = [
    'Slot Chosen',
    'Date Appt Confirmed',
    'Cal Event ID (AH)',
    'Tags',
    'Pending Date (AF)',
    'Pending Slot (AG)',
    'Pending Confirmation (AI)',
    'Last Bot Msg Time (AD)',
    'Follow Up Count (AK)',
    'Last Customer Msg (AM)',
    'Last Follow Up At (AL)',
    'Bot Cooldown'
  ];
  const cleared = [];
  fieldsToBlank.forEach(function(f) {
    try {
      setCellByHeader(sheet, rowNum, f, '');
      cleared.push(f);
    } catch (_e) {
      // header missing — silently skip (defensive against future column rename)
    }
  });

  // Reset to known-good baseline
  setCellByHeader(sheet, rowNum, 'Status', 'Pending Site Visit');
  try { setCellByHeader(sheet, rowNum, 'Flow Stage (AC)', 'welcome_sent'); } catch (_e) {}
  setCellByHeader(sheet, rowNum, 'Status Changed At', new Date().toISOString());
  setCellByHeader(sheet, rowNum, 'Changed By', body.changedBy || 'v2 reset');

  return jsonResponse({
    status: 'ok',
    rowNum: rowNum,
    group: body.groupName,
    clearedFields: cleared,
    oldCalEventId: oldCalEventId,
    note: oldCalEventId ? 'Calendar event NOT deleted automatically — delete from Google Calendar UI if you want the slot freed for testing' : 'no calendar event was set'
  });
}

// ================================================================
// Set Pending — store/clear v2 agent's pending-confirmation slot
// ================================================================
// Called by n8n's Send Whapi Reply (when bot emits [PROPOSE] marker) and
// Debug Skip Echo (to clear pending on escalation interruption).
// Pass empty strings for pendingDate / pendingSlot / pendingCreatedAt to clear.

// ================================================================
// Round 54 — atomic cooldown claim for v2 chat link-send
// Body: { action, secret, groupName, durationMin? (default 60) }
// Returns: { status:'ok', claimed: true }  on fresh claim (timestamp written)
//          { status:'ok', claimed: false, ageMin } when still in cooldown
// LockService serialises concurrent calls so two parallel requests can't
// both claim. Eliminates the read-after-write race we hit with Sheets'
// eventual consistency (Nithia case — 2 messages 21s apart, both saw
// empty Pending Confirmation (AI) and both sent the link).
// ================================================================
function handleClaimCooldown(body) {
  if (!body.groupName) return jsonResponse({status: 'error', message: 'groupName required'});
  const durationMin = Number(body.durationMin) || 60;

  const lock = LockService.getScriptLock();
  try { lock.waitLock(5000); } catch (e) {
    return jsonResponse({status: 'error', message: 'lock timeout'});
  }
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const gnIdx = headers.indexOf('Group Name (AE)');
    const pcIdx = headers.indexOf('Pending Confirmation (AI)');
    if (gnIdx === -1 || pcIdx === -1) {
      return jsonResponse({status: 'error', message: 'columns missing'});
    }
    let rowNum = null;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][gnIdx] || '').trim() === body.groupName.trim()) {
        rowNum = i + 1; break;
      }
    }
    if (!rowNum) return jsonResponse({status: 'error', message: 'group not found'});

    const cellVal = String(data[rowNum-1][pcIdx] || '').trim();
    if (cellVal) {
      const lastTs = new Date(cellVal);
      const ageMs = Date.now() - lastTs.getTime();
      if (!isNaN(ageMs) && ageMs >= 0 && ageMs < durationMin * 60 * 1000) {
        return jsonResponse({status: 'ok', claimed: false, ageMin: Math.round(ageMs / 60000)});
      }
    }
    sheet.getRange(rowNum, pcIdx + 1).setValue(new Date().toISOString());
    SpreadsheetApp.flush();  // force the write to commit before releasing lock
    return jsonResponse({status: 'ok', claimed: true, rowNum: rowNum});
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function handleSetPending(body) {
  // body: {action: 'setPending', secret, groupName, pendingDate, pendingSlot, pendingCreatedAt}
  if (!body.groupName) {
    return jsonResponse({status: 'error', message: 'groupName required'});
  }
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const gnIdx = headers.indexOf('Group Name (AE)');
  if (gnIdx === -1) return jsonResponse({status: 'error', message: 'Group Name (AE) header missing'});

  let rowNum = null;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][gnIdx] || '').trim() === body.groupName.trim()) {
      rowNum = i + 1;
      break;
    }
  }
  if (!rowNum) return jsonResponse({status: 'error', message: 'group not found: ' + body.groupName});

  // Set or clear all three pending fields atomically. Empty string = clear.
  const pendingDate      = body.pendingDate || '';
  const pendingSlot      = body.pendingSlot || '';
  const pendingCreatedAt = body.pendingCreatedAt || '';

  try { setCellByHeader(sheet, rowNum, 'Pending Date (AF)',         pendingDate); } catch (_e) {}
  try { setCellByHeader(sheet, rowNum, 'Pending Slot (AG)',         pendingSlot); } catch (_e) {}
  try { setCellByHeader(sheet, rowNum, 'Pending Confirmation (AI)', pendingCreatedAt); } catch (_e) {}

  return jsonResponse({
    status: 'ok',
    rowNum: rowNum,
    group: body.groupName,
    pendingDate: pendingDate,
    pendingSlot: pendingSlot,
    pendingCreatedAt: pendingCreatedAt
  });
}

// Round 32 — kanban bulk move-to-phase
function handleBulkMoveStatus(body) {
  const phones = Array.isArray(body.phones) ? body.phones : [];
  const newStatus = String(body.newStatus || '').trim();
  if (!phones.length) return jsonResponse({status: 'error', message: 'no phones'});

  // Whitelist all known statuses (active + terminal). Reject anything else to
  // catch typos before they corrupt the sheet.
  const ALLOWED = [
    'New Lead','Pending Invitation','Pending Site Visit','Site Visit Confirmed',
    'Pending QT','Quotation Sent','Pending I.Date','Pending Downpayment',
    'Pending Balance','Completed with Complaint','Job Complete','Receipt Sent',
    'Lost','Cold Lead','Rejected','Out of Area','Human Handoff'
  ];
  if (ALLOWED.indexOf(newStatus) === -1) {
    return jsonResponse({status: 'error', message: 'invalid newStatus: ' + newStatus});
  }

  const sheet = getSheet();
  const ts = new Date().toISOString();
  let updated = 0;
  const missing = [];
  phones.forEach(function(phone) {
    const row = findRowByPhone(sheet, phone);
    if (!row) { missing.push(phone); return; }
    try { setCellByHeader(sheet, row, 'Status',            newStatus); } catch (_e) {}
    try { setCellByHeader(sheet, row, 'Status Changed At', ts);        } catch (_e) {}
    try { setCellByHeader(sheet, row, 'Changed By',        'Kanban_Bulk'); } catch (_e) {}
    updated++;
  });
  return jsonResponse({status: 'ok', updated: updated, missing: missing, newStatus: newStatus});
}

// ================================================================
// Round 48 — Bulk-link pre-CRM WhatsApp groups
// Body: {
//   action, secret, dryRun: bool, changedBy?,
//   groups: [
//     { kind: 'link_existing', phone, groupId, groupName, inviteLink, newStatus? },
//     { kind: 'create_orphan',  phone, name, status, groupId, groupName, inviteLink }
//   ]
// }
// Returns: { status:'ok', linked, created, skipped, errors:[{reason,payload}] }
// Per-row try/catch — partial failure surfaces in `errors`, not as 500.
// Idempotent: any groupId already present in CRM goes to `skipped` regardless of kind.
// ================================================================

function handleBulkLinkGroups(body) {
  if (!Array.isArray(body.groups)) {
    return jsonResponse({status: 'error', message: 'groups[] required'});
  }
  const result = { status: 'ok', linked: 0, created: 0, skipped: 0, errors: [] };

  // Dry-run: validate shape, count classifications, no writes.
  if (body.dryRun === true) {
    body.groups.forEach(function(g) {
      if (!g || !g.kind || !g.groupId) { result.errors.push({reason: 'invalid entry', payload: g}); return; }
      if (g.kind === 'link_existing') result.linked++;
      else if (g.kind === 'create_orphan') result.created++;
      else result.errors.push({reason: 'unknown kind: ' + g.kind, payload: g});
    });
    result.dryRun = true;
    return jsonResponse(result);
  }

  const sheet = getSheet();
  const ts = new Date().toISOString();
  const todayStr = ts.slice(0, 10);
  const changedBy = body.changedBy || 'Kanban_BulkLink';
  const headers = getHeaders(sheet);
  const gidCol = headers.colByName['Group ID (AB)'];
  if (!gidCol) return jsonResponse({status: 'error', message: 'Group ID (AB) column not found'});

  // Pre-fetch existing Group IDs once for idempotent skip.
  const allData = sheet.getDataRange().getValues();
  const existingGids = new Set();
  for (let i = 1; i < allData.length; i++) {
    const v = String(allData[i][gidCol - 1] || '').trim();
    if (v) existingGids.add(v);
  }

  body.groups.forEach(function(g) {
    try {
      if (!g || !g.kind || !g.groupId) {
        result.errors.push({reason: 'invalid entry', payload: g});
        return;
      }
      // Idempotent skip: any groupId already linked anywhere → skip.
      if (existingGids.has(g.groupId)) { result.skipped++; return; }

      if (g.kind === 'link_existing') {
        if (!g.phone) { result.errors.push({reason: 'phone required for link_existing', groupId: g.groupId}); return; }
        const row = findRowByPhone(sheet, g.phone);
        if (!row) {
          result.errors.push({reason: 'lead not found at write-time', phone: g.phone, groupId: g.groupId});
          return;
        }
        // Defensive: don't overwrite if row already has a different Group ID.
        const existingGid = String(sheet.getRange(row, gidCol).getValue() || '').trim();
        if (existingGid && existingGid !== g.groupId) {
          result.skipped++;
          return;
        }
        setCellByHeader(sheet, row, 'Group ID (AB)', g.groupId);
        if (g.groupName)  setCellByHeader(sheet, row, 'Group Name (AE)', g.groupName);
        if (g.inviteLink) setCellByHeader(sheet, row, 'Group Invite Link (AJ)', g.inviteLink);
        if (g.newStatus) {
          setCellByHeader(sheet, row, 'Status', g.newStatus);
          setCellByHeader(sheet, row, 'Status Changed At', ts);
        }
        setCellByHeader(sheet, row, 'Changed By', changedBy);
        result.linked++;
        existingGids.add(g.groupId);

      } else if (g.kind === 'create_orphan') {
        if (!g.phone || !g.name || !g.status) {
          result.errors.push({reason: 'orphan needs phone+name+status', payload: g});
          return;
        }
        // Server-side dup gate (race-safe — row may have appeared since Preview).
        const matches = findAllMatchesByPhone(sheet, g.phone);
        const ACTIVE = ['New Lead','Pending Invitation','Pending Site Visit','Site Visit Confirmed','Pending QT','Quotation Sent','Follow Up','Pending I.Date','I.Date Confirmed','Job In Progress'];
        const activeMatch = matches.find(function(m) { return ACTIVE.indexOf(m.status) !== -1; });
        if (activeMatch) {
          // Convert to link_existing on the fly.
          const row = activeMatch.rowNum;
          const existingGid = String(sheet.getRange(row, gidCol).getValue() || '').trim();
          if (!existingGid) {
            setCellByHeader(sheet, row, 'Group ID (AB)', g.groupId);
            if (g.groupName)  setCellByHeader(sheet, row, 'Group Name (AE)', g.groupName);
            if (g.inviteLink) setCellByHeader(sheet, row, 'Group Invite Link (AJ)', g.inviteLink);
            setCellByHeader(sheet, row, 'Changed By', changedBy);
            result.linked++;
          } else {
            result.skipped++;
          }
          existingGids.add(g.groupId);
          return;
        }
        // Truly orphan — append new row, mirror handleCreateLead's column set.
        const newRow = sheet.getLastRow() + 1;
        setCellByHeader(sheet, newRow, 'Timestamp', ts);
        setCellByHeader(sheet, newRow, 'Phone', g.phone);
        setCellByHeader(sheet, newRow, 'Name', g.name);
        setCellByHeader(sheet, newRow, 'Status', g.status);
        setCellByHeader(sheet, newRow, 'Status Changed At', ts);
        setCellByHeader(sheet, newRow, 'Changed By', changedBy);
        setCellByHeader(sheet, newRow, 'Date Lead In', todayStr);
        setCellByHeader(sheet, newRow, 'Source', 'Bulk Link (Pre-CRM)');
        setCellByHeader(sheet, newRow, 'Group ID (AB)', g.groupId);
        if (g.groupName)  setCellByHeader(sheet, newRow, 'Group Name (AE)', g.groupName);
        if (g.inviteLink) setCellByHeader(sheet, newRow, 'Group Invite Link (AJ)', g.inviteLink);
        result.created++;
        existingGids.add(g.groupId);

      } else {
        result.errors.push({reason: 'unknown kind: ' + g.kind, payload: g});
      }
    } catch (e) {
      result.errors.push({reason: String(e.message || e), payload: g});
    }
  });

  return jsonResponse(result);
}

// ================================================================
// Round 45 — Cancel an existing appointment (SVC -> PSV)
// Body: {phone, secret, changedBy}
// Effect: clears Cal Event ID + slot/date + pending fields, flips Status to
// 'Pending Site Visit'. Calendar-event deletion + WA messages happen in the
// n8n LG - Cancel Appointment workflow that calls this handler.
// ================================================================

function handleCancelAppointment(body) {
  const phone = String(body.phone || '').trim();
  if (!phone) return jsonResponse({status: 'error', message: 'phone required'});

  const sheet = getSheet();
  const row = findRowByPhone(sheet, phone);
  if (!row) return jsonResponse({status: 'error', message: 'lead not found', phone: phone});

  const ts = new Date().toISOString();
  try { setCellByHeader(sheet, row, 'Status',                    'Pending Site Visit'); } catch (_) {}
  try { setCellByHeader(sheet, row, 'Cal Event ID (AH)',         ''); } catch (_) {}
  try { setCellByHeader(sheet, row, 'Slot Chosen',               ''); } catch (_) {}
  try { setCellByHeader(sheet, row, 'Date Appt Confirmed',       ''); } catch (_) {}
  try { setCellByHeader(sheet, row, 'Pending Date (AF)',         ''); } catch (_) {}
  try { setCellByHeader(sheet, row, 'Pending Slot (AG)',         ''); } catch (_) {}
  try { setCellByHeader(sheet, row, 'Pending Confirmation (AI)', ''); } catch (_) {}
  try { setCellByHeader(sheet, row, 'Status Changed At',         ts); } catch (_) {}
  try { setCellByHeader(sheet, row, 'Changed By',                body.changedBy || 'Kanban_CancelAppt'); } catch (_) {}

  return jsonResponse({status: 'ok', row: row, newStatus: 'Pending Site Visit'});
}

// ================================================================
// Task 2 — Edit lead details from kanban modal
// Body: {action, secret, phone, name?, location?, fullAddress?, problemType?, slabSize?, source?, newGroupName?, changedBy?}
// Phone is the lookup key (immutable). Only fields present in body are written.
// ================================================================

function handleUpdateLeadDetails(body) {
  const phone = String(body.phone || '').trim();
  if (!phone) return jsonResponse({status: 'error', message: 'phone required'});

  const sheet = getSheet();
  const row = findRowByPhone(sheet, phone);
  if (!row) return jsonResponse({status: 'error', message: 'lead not found', phone: phone});

  const mapping = {
    'name':         'Name',
    'location':     'Location',
    'fullAddress':  'Full Address',
    'problemType':  'Problem Type',
    'slabSize':     'Slab Size (sqft)',
    'source':       'Source',
    'newGroupName': 'Group Name (AE)'
  };

  const updated = [];
  Object.keys(mapping).forEach(function(bodyKey) {
    const v = body[bodyKey];
    if (v === undefined || v === null) return;
    const header = mapping[bodyKey];
    try {
      setCellByHeader(sheet, row, header, String(v));
      updated.push(header);
    } catch (_e) {}
  });

  try { setCellByHeader(sheet, row, 'Status Changed At', new Date().toISOString()); } catch (_e) {}
  try { setCellByHeader(sheet, row, 'Changed By', body.changedBy || 'Kanban_Edit'); } catch (_e) {}

  return jsonResponse({status: 'ok', row: row, updated: updated});
}

// ================================================================
// Quick test (run from Apps Script editor manually)
// ================================================================

function testPing() {
  const fakeEvent = {
    postData: {
      contents: JSON.stringify({secret: SHARED_SECRET, action: 'ping'})
    }
  };
  const r = doPost(fakeEvent);
  Logger.log(r.getContent());
}

function testFindRow() {
  const sheet = getSheet();
  const rowNum = findRowByPhone(sheet, '60183639951');
  Logger.log('Row found: ' + rowNum);
}

function testSetPending() {
  // Manual test from editor — sets a pending slot for the test group
  const fakeEvent = {
    postData: {
      contents: JSON.stringify({
        secret: SHARED_SECRET,
        action: 'setPending',
        groupName: 'SV-0426-018 Annie - Ampang',
        pendingDate: '2026-05-13',
        pendingSlot: '4',
        pendingCreatedAt: new Date().toISOString()
      })
    }
  };
  const r = doPost(fakeEvent);
  Logger.log(r.getContent());
}
