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
const N8N_RENAME_GROUP_URL = 'https://leakguard.app.n8n.cloud/webhook/lg-rename-group'; // Round 70 — auto-rename WA group on QT-PDF detection
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
      case 'updateStatusByGroup': return handleUpdateStatusByGroup(body);  // round 61 — auto phase-shift from template detection
      case 'addTagByGroup':      return handleAddTagByGroup(body);  // round 63 — auto-tag from template detection
      case 'addPaymentVerification':   return handleAddPaymentVerification(body);  // round 64 — payment-noted template
      case 'clearPaymentVerification': return handleClearPaymentVerification(body);  // round 64 — account tick verify
      case 'updateLastAdminMsg': return handleUpdateLastAdminMsg(body);  // round 68 — admin-msg capture from Parse & Route
      case 'nextSeNumber':       return handleNextSeNumber(body);  // round 71 — atomic SE-MMYY-NNN counter for Estimation Builder
      case 'uploadEstimationPhoto': return handleUploadEstimationPhoto(body);  // round 72 — Drive upload from Estimation Builder
      case 'login':              return handleLogin(body);  // round 76 — kanban login gate (Users tab)
      case 'staffJobs':          return handleStaffJobs(body);  // round 76 phase 2/3 — staff assigned + repair cards
      case 'staffCommand':       return handleStaffCommand(body);  // round 76 phase 3 — WA -myjobs/-pending text command
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

// Round 70 — compute the new group name when a QT-MMYY-NNN PDF is detected.
// - Strips a leading SV-/SVJB- booking prefix
// - If a QT-MMYY-NNN[/NNN...] prefix already exists with the same MMYY,
//   slash-appends NNN; if NNN is already present, returns null (no-op)
// - Otherwise prepends QT-MMYY-NNN
function _buildQtGroupName(currentName, qtCode) {
  if (!/^\d{4}-\d{3}$/.test(String(qtCode || ''))) return null;
  const mmyy = qtCode.slice(0, 4);
  const nnn  = qtCode.slice(5);
  const work = String(currentName || '').replace(/^(SV|SVJB)-\d{4}-\d{3}\s*/i, '').trim();
  const m = work.match(/^(?:QT-)?(\d{4})-([\d\/]+)\b(.*)$/i);
  if (m) {
    const exMmyy = m[1];
    const exCodes = m[2].split('/').filter(function(s) { return s; });
    const tail = (m[3] || '').replace(/^\s+/, ' ');
    if (exMmyy === mmyy) {
      if (exCodes.indexOf(nnn) !== -1) return null;
      exCodes.push(nnn);
      return ('QT-' + mmyy + '-' + exCodes.join('/') + (tail ? ' ' + tail.trim() : '')).trim();
    }
    return ('QT-' + exMmyy + '-' + exCodes.join('/') + '/' + mmyy + '-' + nnn + (tail ? ' ' + tail.trim() : '')).trim();
  }
  return ('QT-' + mmyy + '-' + nnn + (work ? ' ' + work : '')).trim();
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

// Round 71 — atomic SE-MMYY-NNN counter for the Estimation Builder.
// Estimations live outside AutoCount; this handler owns the numbering.
// Resets count to 1 on MMYY rollover. Uses LockService for atomicity.
function handleNextSeNumber(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const sheet = SpreadsheetApp.openById(LIVE_SHEET_ID).getSheetByName('Counters');
    if (!sheet) return jsonResponse({status: 'error', message: 'Counters sheet missing — run bootstrapCountersSheet() once from the editor'});
    const data = sheet.getDataRange().getValues();
    let rowIdx = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim().toUpperCase() === 'SE') { rowIdx = i; break; }
    }
    if (rowIdx < 0) return jsonResponse({status: 'error', message: 'SE row missing in Counters — run bootstrapCountersSheet()'});
    // MYT-correct MMYY (CLAUDE.md gotcha #3 idiom — extract via getUTC* only)
    const mytNow = new Date(Date.now() + 8 * 60 * 60 * 1000);
    const mmyy = String(mytNow.getUTCMonth() + 1).padStart(2, '0') + String(mytNow.getUTCFullYear()).slice(-2);
    const sheetMmyy = String(data[rowIdx][1] || '').trim();
    let count = Number(data[rowIdx][2] || 0);
    if (sheetMmyy !== mmyy) count = 0;
    count = count + 1;
    sheet.getRange(rowIdx + 1, 2).setValue(mmyy);
    sheet.getRange(rowIdx + 1, 3).setValue(count);
    const docNo = 'SE-' + mmyy + '-' + String(count).padStart(3, '0');
    return jsonResponse({status: 'ok', docNo: docNo});
  } finally {
    lock.releaseLock();
  }
}

function handleUpdateQuotation(body) {
  // body: {phone, quotation, secret}
  const sheet = getSheet();
  const rowNum = findRowByPhone(sheet, body.phone);
  if (!rowNum) return jsonResponse({status: 'error', message: 'lead not found'});

  setCellByHeader(sheet, rowNum, 'Quotation (RM)', body.quotation || '');
  return jsonResponse({status: 'ok'});
}

// ================================================================
// Round 61 — phase shift by Group ID (called from WA Receiver template detect)
// ================================================================
// WA Receiver detects admin template messages in groups and fires this handler
// to flip the phase, audit Changed By, and DM the main admin a confirmation.
// Idempotent: if the row is already at the target status, the inner
// handleUpdateStatus is a no-op and the admin DM is skipped.

function handleUpdateStatusByGroup(body) {
  // body: {action, secret, groupId, status, changedBy?, notifyAdmin?}
  if (!body.groupId || !body.status) {
    return jsonResponse({status: 'error', message: 'groupId and status required'});
  }
  const sheet = getSheet();
  const headers = getHeaders(sheet);
  const groupIdCol = headers.colByName['Group ID (AB)'];
  const phoneCol   = headers.colByName['Phone'];
  const nameCol    = headers.colByName['Name'];
  const gnameCol   = headers.colByName['Group Name (AE)'];
  const statusCol  = headers.colByName['Status'];
  if (!groupIdCol || !phoneCol) {
    return jsonResponse({status: 'error', message: 'required columns missing'});
  }

  const data = sheet.getDataRange().getValues();
  const target = String(body.groupId).trim();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][groupIdCol - 1] || '').trim() !== target) continue;
    const phone     = String(data[i][phoneCol  - 1] || '').trim();
    const name      = nameCol  ? String(data[i][nameCol  - 1] || '').trim() : '';
    const groupName = gnameCol ? String(data[i][gnameCol - 1] || '').trim() : '';
    const prevStatus = statusCol ? String(data[i][statusCol - 1] || '').trim() : '';

    // Round 70: if Parse & Route attached a QT code (Round 70 n8n patch),
    // rebuild Group Name (AE) with "QT-MMYY-NNN" prefix (slash-append on
    // subsequent QTs) and push the new subject to the live WA group via
    // lg-rename-group. Runs BEFORE the regression guard so a revised QT
    // sent to a card already past QS still updates the group name even
    // when the status shift itself is blocked.
    if (body.qtCode && /^\d{4}-\d{3}$/.test(body.qtCode)) {
      const newName = _buildQtGroupName(groupName, body.qtCode);
      if (newName && newName !== groupName) {
        const rowNum = i + 1;
        try { setCellByHeader(sheet, rowNum, 'Group Name (AE)', newName); } catch (_e) {}
        try {
          UrlFetchApp.fetch(N8N_RENAME_GROUP_URL, {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify({secret: SHARED_SECRET, groupId: target, newGroupName: newName}),
            muteHttpExceptions: true,
          });
        } catch (_e) { /* best-effort; CRM is source of truth either way */ }
      }
    }

    // Round 69: refuse auto-template shifts that would regress a card backward
    // in the funnel. Prevents re-sent quotation PDFs (and accidental template
    // re-sends) from pulling Pending Downpayment / Balance / Completed cards
    // back to an earlier phase. Manual kanban clicks (changedBy='Kanban') are
    // unaffected; only calls tagged 'Auto-Template:*' from Parse & Route gate.
    if (String(body.changedBy || '').indexOf('Auto-Template') === 0) {
      const _funnel = ['New Lead','Pending Invitation','Pending Site Visit',
                       'Site Visit Confirmed','Pending QT','Quotation Sent',
                       'Pending Downpayment','Pending I.Date','I.Date Confirmed',
                       'Job In Progress','Pending Balance','Job Complete',
                       'Receipt Sent','Completed'];
      const _cur = _funnel.indexOf(prevStatus);
      const _tgt = _funnel.indexOf(body.status);
      if (_cur >= 0 && _tgt >= 0 && _tgt < _cur) {
        return jsonResponse({status: 'ignored',
                             reason: 'auto-template regression blocked',
                             prev: prevStatus, attempted: body.status});
      }
    }

    // Delegate to existing handleUpdateStatus — inherits Round 58 CAPI fan-out,
    // Status Changed At update, Changed By audit, prev-status snapshot.
    const result = handleUpdateStatus({
      action: 'updateStatus',
      phone: phone,
      status: body.status,
      changedBy: body.changedBy || 'Auto-Template'
    });

    // Round 64: when shifting INTO Completed (via the Google review template),
    // defensively clear pending_verification — by this point payment is settled
    // so the account-team queue should drop this row.
    if (body.status === 'Completed' && prevStatus !== 'Completed') {
      const rowNum = i + 1;
      const tagsColX = headers.colByName['Tags'];
      if (tagsColX) {
        const cur = String(data[i][tagsColX - 1] || '').trim();
        const cleaned = cur.split(',').map(function(t){ return t.trim(); }).filter(function(t){ return t && t !== 'pending_verification'; }).join(',');
        if (cleaned !== cur) {
          setCellByHeader(sheet, rowNum, 'Tags', cleaned);
          try { setCellByHeader(sheet, rowNum, 'Verification Amount', ''); } catch (_e) {}
          try { setCellByHeader(sheet, rowNum, 'Verification Date', ''); } catch (_e) {}
        }
      }
    }

    // Round 61.1: reset FU clock so LG-Follow Up doesn't fire moments after a
    // template-triggered phase change. Admin's template message is itself the
    // outbound communication — LG-Follow Up should start its window from now.
    // Touches: Last Bot Msg Time (AD), Last Follow Up At (AL), Follow Up Count (AK).
    if (prevStatus !== body.status) {
      const rowNum = i + 1;
      const nowIso = new Date().toISOString();
      try { setCellByHeader(sheet, rowNum, 'Last Bot Msg Time (AD)', nowIso); } catch (_e) {}
      try { setCellByHeader(sheet, rowNum, 'Last Follow Up At (AL)', nowIso); } catch (_e) {}
      try { setCellByHeader(sheet, rowNum, 'Follow Up Count (AK)', 0); } catch (_e) {}
    }

    // Round 62: admin DM disabled per user directive ("stop all notification msg
    // that will send back to admin"). To re-enable, restore the block below.
    // if (body.notifyAdmin !== false && prevStatus !== body.status) {
    //   try {
    //     UrlFetchApp.fetch(N8N_WAGROUP_URL, {
    //       method: 'post',
    //       contentType: 'application/json',
    //       headers: {'Authorization': 'Bearer ' + WHAPI_TOKEN},
    //       payload: JSON.stringify({
    //         to: '60183639321',  // main admin
    //         body: '✓ Auto-shifted ' + (groupName || ('group ' + target)) +
    //               '\n  ' + prevStatus + ' → ' + body.status +
    //               '\n  (template detected from ' + (body.changedBy || 'staff') + ')',
    //         typing_time: 0
    //       }),
    //       muteHttpExceptions: true
    //     });
    //   } catch (_e) { /* best-effort */ }
    // }

    return result;
  }
  return jsonResponse({status: 'error', message: 'group not found', groupId: target});
}

// ================================================================
// Round 63 — append a tag to a row by Group ID (called from WA Receiver
// template detection — currently used for the 'repair' tag fired by the
// senior-inspection-team callback template).
// ================================================================
// Idempotent: if the tag is already present in the comma-separated Tags
// cell, it's not added again (no duplicate). Returns the final tag list
// so the caller can audit if needed.

function handleAddTagByGroup(body) {
  // body: {action, secret, groupId, tag, changedBy?}
  if (!body.groupId || !body.tag) {
    return jsonResponse({status: 'error', message: 'groupId and tag required'});
  }
  const sheet = getSheet();
  const headers = getHeaders(sheet);
  const groupIdCol = headers.colByName['Group ID (AB)'];
  const tagsCol    = headers.colByName['Tags'];
  if (!groupIdCol || !tagsCol) {
    return jsonResponse({status: 'error', message: 'required columns missing'});
  }

  const data = sheet.getDataRange().getValues();
  const target = String(body.groupId).trim();
  const newTag = String(body.tag).trim();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][groupIdCol - 1] || '').trim() !== target) continue;
    const rowNum = i + 1;
    const existing = String(data[i][tagsCol - 1] || '').trim();
    const tagList = existing.split(',').map(function(t){ return t.trim(); }).filter(function(t){ return t; });
    let added = false;
    if (tagList.indexOf(newTag) === -1) {
      tagList.push(newTag);
      setCellByHeader(sheet, rowNum, 'Tags', tagList.join(','));
      added = true;
    }
    return jsonResponse({status: 'ok', rowNum: rowNum, tags: tagList.join(','), added: added});
  }
  return jsonResponse({status: 'error', message: 'group not found', groupId: target});
}

// ================================================================
// Round 68 — capture admin's last message into a CRM column
// ================================================================
// Parse & Route fires this for every staff message in a linked group (commands
// are routed away upstream). Apps Script writes "<senderName>: <text>" to the
// 'Last Admin Msg' column so the kanban can show what admin most recently said
// to that lead — symmetric to the existing 'Last Customer Msg (AM)' display.
// Idempotent overwrite: latest message wins. Silently no-ops if group not linked.

function handleUpdateLastAdminMsg(body) {
  // body: {action, secret, groupId, senderName, msgText}
  if (!body.groupId) return jsonResponse({status: 'error', message: 'groupId required'});
  const sheet = getSheet();
  const h = getHeaders(sheet);
  const gidCol = h.colByName['Group ID (AB)'];
  const lamCol = h.colByName['Last Admin Msg'];
  if (!gidCol) return jsonResponse({status: 'error', message: 'Group ID (AB) column missing'});
  if (!lamCol) return jsonResponse({status: 'error', message: 'Last Admin Msg column missing — run bootstrapLastAdminMsgColumn()'});
  const data = sheet.getDataRange().getValues();
  const target = String(body.groupId).trim();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][gidCol - 1] || '').trim() !== target) continue;
    const sender = String(body.senderName || '').trim() || 'Staff';
    const text = String(body.msgText || '').trim().slice(0, 500); // cap storage
    setCellByHeader(sheet, i + 1, 'Last Admin Msg', sender + ': ' + text);
    return jsonResponse({status: 'ok', rowNum: i + 1});
  }
  return jsonResponse({status: 'ok', skipped: true, reason: 'group not linked'});
}

// ================================================================
// Round 64 — payment verification queue (Pending Downpayment + Pending Balance)
// ================================================================
// WA Receiver detects admin's "Noted with thx" template, extracts the RM amount,
// and fires this handler. Adds 'pending_verification' tag + writes the amount
// + sent date to two CRM columns so account team can see the queue at a glance.
//
// Companion handler handleClearPaymentVerification is fired by the kanban
// ✓ Verify button when account confirms payment received.

function handleAddPaymentVerification(body) {
  // body: {action, secret, groupId, amount, changedBy?}
  if (!body.groupId) {
    return jsonResponse({status: 'error', message: 'groupId required'});
  }
  const sheet = getSheet();
  const headers = getHeaders(sheet);
  const groupIdCol = headers.colByName['Group ID (AB)'];
  const tagsCol    = headers.colByName['Tags'];
  const statusCol  = headers.colByName['Status'];
  const amountCol  = headers.colByName['Verification Amount'];
  const dateCol    = headers.colByName['Verification Date'];
  if (!groupIdCol || !tagsCol) {
    return jsonResponse({status: 'error', message: 'required columns missing'});
  }
  if (!amountCol || !dateCol) {
    return jsonResponse({status: 'error', message: 'run bootstrapVerificationColumns() first'});
  }

  const data = sheet.getDataRange().getValues();
  const target = String(body.groupId).trim();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][groupIdCol - 1] || '').trim() !== target) continue;
    const rowNum = i + 1;
    const status = String(data[i][statusCol - 1] || '').trim();

    // Gate: only fire for Pending Downpayment or Pending Balance phases.
    // If admin sends the template in another phase (e.g., wrong group), no-op.
    if (status !== 'Pending Downpayment' && status !== 'Pending Balance') {
      return jsonResponse({status: 'ok', skipped: true, reason: 'phase not eligible: ' + status, rowNum: rowNum});
    }

    // Round 65: route account verification to Alvin regardless of original sales owner
    try { setCellByHeader(sheet, rowNum, 'Assigned To', 'Alvin'); } catch (_e) {}

    // Append pending_verification tag (idempotent)
    const existing = String(data[i][tagsCol - 1] || '').trim();
    const tagList = existing.split(',').map(function(t){ return t.trim(); }).filter(function(t){ return t; });
    if (tagList.indexOf('pending_verification') === -1) tagList.push('pending_verification');
    setCellByHeader(sheet, rowNum, 'Tags', tagList.join(','));

    // Write amount (whatever bot parsed, e.g., "RM850") + date (now)
    if (body.amount) {
      setCellByHeader(sheet, rowNum, 'Verification Amount', String(body.amount));
    }
    setCellByHeader(sheet, rowNum, 'Verification Date', new Date().toISOString());

    return jsonResponse({status: 'ok', rowNum: rowNum, tags: tagList.join(','), amount: body.amount || ''});
  }
  return jsonResponse({status: 'error', message: 'group not found', groupId: target});
}

function handleClearPaymentVerification(body) {
  // body: {action, secret, phone, changedBy?}
  // Removes 'pending_verification' tag + clears Verification Amount + Verification Date.
  // Called from the kanban ✓ Verify button.
  if (!body.phone) return jsonResponse({status: 'error', message: 'phone required'});
  const sheet = getSheet();
  const rowNum = findRowByPhone(sheet, body.phone);
  if (!rowNum) return jsonResponse({status: 'error', message: 'lead not found'});

  const h = getHeaders(sheet);
  const tagsCol = h.colByName['Tags'];
  if (!tagsCol) return jsonResponse({status: 'error', message: 'Tags column missing'});

  // Round 65: capture current status before mutating so we can decide auto-jump.
  const _statusColR65 = h.colByName['Status'];
  const curStatus = _statusColR65 ? String(sheet.getRange(rowNum, _statusColR65).getValue() || '').trim() : '';

  const existing = String(sheet.getRange(rowNum, tagsCol).getValue() || '').trim();
  const next = existing.split(',').map(function(t){ return t.trim(); }).filter(function(t){ return t && t !== 'pending_verification'; });
  setCellByHeader(sheet, rowNum, 'Tags', next.join(','));
  // Round 66: don't clear amount/date — repurpose Verification Date as
  // "last verification activity timestamp" (queue OR completion). With the
  // pending_verification tag gone after this handler runs, the kanban's
  // notification panel reads: tag present + date = queued event; tag absent
  // + date = completed event. Amount stays so the panel can display "RM850".
  try { setCellByHeader(sheet, rowNum, 'Verification Date', new Date().toISOString()); } catch (_e) {}
  if (h.colByName['Changed By']) {
    sheet.getRange(rowNum, h.colByName['Changed By']).setValue(body.changedBy || 'Account-Verified');
  }

  // Round 65: only Pending Downpayment auto-jumps to Pending Balance on ✓ Verify.
  // Pending Balance is template-driven (Google-review template -> Completed via R64.2).
  let newStatus = curStatus;
  if (curStatus === 'Pending Downpayment') {
    handleUpdateStatus({
      action: 'updateStatus',
      phone: body.phone,
      status: 'Pending Balance',
      changedBy: body.changedBy || 'Kanban_Verify_AutoShift'
    });
    newStatus = 'Pending Balance';
  }
  return jsonResponse({status: 'ok', rowNum: rowNum, tags: next.join(','), newStatus: newStatus});
}

// Run ONCE from Apps Script editor to add the two new columns.
// Round 72 — one-time creation of the Drive folder for Estimation photos.
// Idempotent: looks for an existing folder by name, stores ID in
// ScriptProperties so handleUploadEstimationPhoto can find it without
// hardcoding. Run ONCE from the Apps Script editor.
function bootstrapEstimationPhotosFolder() {
  const props = PropertiesService.getScriptProperties();
  const existingId = props.getProperty('EST_PHOTOS_FOLDER_ID');
  if (existingId) {
    try {
      const f = DriveApp.getFolderById(existingId);
      Logger.log('bootstrapEstimationPhotosFolder: existing folder found — ' + f.getName() + ' (' + existingId + ')');
      return;
    } catch (_e) {
      Logger.log('bootstrapEstimationPhotosFolder: stored ID no longer valid, recreating');
    }
  }
  const name = 'Leak Guard Estimation Photos';
  const it = DriveApp.getFoldersByName(name);
  let folder;
  if (it.hasNext()) {
    folder = it.next();
    Logger.log('bootstrapEstimationPhotosFolder: reusing existing folder named ' + name);
  } else {
    folder = DriveApp.createFolder(name);
    Logger.log('bootstrapEstimationPhotosFolder: created folder ' + name);
  }
  props.setProperty('EST_PHOTOS_FOLDER_ID', folder.getId());
  Logger.log('bootstrapEstimationPhotosFolder: stored ID ' + folder.getId());
}

// Round 72 — accept a base64 dataURL photo from the Estimation Builder,
// save into the Drive folder, return a shareable URL for audit. Best-effort:
// caller fires-and-forgets, PDF render does not depend on this returning.
function handleUploadEstimationPhoto(body) {
  if (!body.phone || !body.dataUrl) {
    return jsonResponse({status: 'error', message: 'phone and dataUrl required'});
  }
  const props = PropertiesService.getScriptProperties();
  const folderId = props.getProperty('EST_PHOTOS_FOLDER_ID');
  if (!folderId) {
    return jsonResponse({status: 'error', message: 'EST_PHOTOS_FOLDER_ID not set — run bootstrapEstimationPhotosFolder() once'});
  }
  try {
    const folder = DriveApp.getFolderById(folderId);
    // dataUrl format: "data:image/jpeg;base64,<base64>"
    const comma = String(body.dataUrl).indexOf(',');
    if (comma < 0) return jsonResponse({status: 'error', message: 'malformed dataUrl'});
    const meta = String(body.dataUrl).slice(0, comma);
    const b64 = String(body.dataUrl).slice(comma + 1);
    const mime = (meta.match(/data:([^;]+);/) || [])[1] || 'image/jpeg';
    const ext = mime.indexOf('png') >= 0 ? 'png' : 'jpg';
    const ts = Utilities.formatDate(new Date(Date.now() + 8*60*60*1000), 'UTC', 'yyyyMMdd-HHmmss');
    const slabPart = String(body.slabName || 'slab').replace(/[^A-Za-z0-9-_]+/g, '-').slice(0, 30);
    const filename = String(body.phone) + '_' + slabPart + '_' + ts + '.' + ext;
    const blob = Utilities.newBlob(Utilities.base64Decode(b64), mime, filename);
    const file = folder.createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (_e) {}
    return jsonResponse({status: 'ok', url: file.getUrl(), id: file.getId()});
  } catch (e) {
    return jsonResponse({status: 'error', message: String(e && e.message || e)});
  }
}

// Round 71 — one-time creation of the Counters tab + seeding the SE row.
// Idempotent: skips if tab exists, only seeds SE row if missing.
function bootstrapCountersSheet() {
  const ss = SpreadsheetApp.openById(LIVE_SHEET_ID);
  let sheet = ss.getSheetByName('Counters');
  if (!sheet) {
    sheet = ss.insertSheet('Counters');
    sheet.getRange(1, 1, 1, 3).setValues([['Type', 'Last MMYY', 'Last Count']]);
    sheet.setFrozenRows(1);
    Logger.log('bootstrapCountersSheet: created Counters tab');
  }
  const data = sheet.getDataRange().getValues();
  let hasSe = false;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim().toUpperCase() === 'SE') { hasSe = true; break; }
  }
  if (!hasSe) {
    const mytNow = new Date(Date.now() + 8 * 60 * 60 * 1000);
    const mmyy = String(mytNow.getUTCMonth() + 1).padStart(2, '0') + String(mytNow.getUTCFullYear()).slice(-2);
    sheet.appendRow(['SE', mmyy, 0]);
    Logger.log('bootstrapCountersSheet: seeded SE row at MMYY=' + mmyy + ', count=0');
  } else {
    Logger.log('bootstrapCountersSheet: SE row already present, nothing to do');
  }
}

function bootstrapVerificationColumns() {
  const sheet = getSheet();
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h || '').trim(); });
  const wanted = ['Verification Amount', 'Verification Date'];
  const missing = wanted.filter(function(n){ return headers.indexOf(n) === -1; });
  if (missing.length === 0) {
    Logger.log('bootstrapVerificationColumns: both columns already present.');
    return;
  }
  sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
  Logger.log('bootstrapVerificationColumns: appended ' + missing.length + ' column(s) at col ' + (lastCol + 1) + ': ' + missing.join(', '));
}

// Round 68 — Run ONCE from Apps Script editor to add the Last Admin Msg column.
// Idempotent: re-running is a no-op if the column already exists.
function bootstrapLastAdminMsgColumn() {
  const sheet = getSheet();
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h || '').trim(); });
  if (headers.indexOf('Last Admin Msg') !== -1) {
    Logger.log('bootstrapLastAdminMsgColumn: column already present.');
    return;
  }
  sheet.getRange(1, lastCol + 1).setValue('Last Admin Msg');
  Logger.log('bootstrapLastAdminMsgColumn: appended at col ' + (lastCol + 1));
}

// Round 65.1 — one-off backfill: any row already tagged 'pending_verification'
// gets re-assigned to Alvin. Use this after the initial R65 deploy to fix
// pre-existing tagged rows that missed the auto-assign in handleAddPaymentVerification.
// Idempotent: re-running skips rows already assigned to Alvin.
function backfillPendingVerifAlvin() {
  const sheet = getSheet();
  const h = getHeaders(sheet);
  const tagsCol = h.colByName['Tags'];
  const assignedCol = h.colByName['Assigned To'];
  if (!tagsCol || !assignedCol) {
    Logger.log('backfillPendingVerifAlvin: Tags or Assigned To column missing.');
    return;
  }
  const data = sheet.getDataRange().getValues();
  let fixed = 0;
  for (let i = 1; i < data.length; i++) {
    const tags = String(data[i][tagsCol - 1] || '').toLowerCase();
    if (tags.indexOf('pending_verification') === -1) continue;
    const cur = String(data[i][assignedCol - 1] || '').trim();
    if (cur === 'Alvin') continue;
    setCellByHeader(sheet, i + 1, 'Assigned To', 'Alvin');
    Logger.log('Row ' + (i + 1) + ': "' + cur + '" -> Alvin');
    fixed++;
  }
  Logger.log('backfillPendingVerifAlvin: ' + fixed + ' row(s) reassigned.');
}

// Round 78 B.12 — one-off migration: rename every 'Assigned To = William'
// row to 'Team A' so the kanban filter + WA-msg lookup + reminders all
// match the new canonical name written by lg-booking. Idempotent:
// re-running skips rows already on 'Team A'.
function backfillRenameAssignedTo() {
  const sheet = getSheet();
  const h = getHeaders(sheet);
  const col = h.colByName['Assigned To'];
  if (!col) { Logger.log('backfillRenameAssignedTo: Assigned To column missing.'); return; }
  const data = sheet.getDataRange().getValues();
  let fixed = 0;
  for (let i = 1; i < data.length; i++) {
    const cur = String(data[i][col - 1] || '').trim();
    if (cur !== 'William') continue;
    setCellByHeader(sheet, i + 1, 'Assigned To', 'Team A');
    Logger.log('Row ' + (i + 1) + ': "William" -> "Team A"');
    fixed++;
  }
  Logger.log('backfillRenameAssignedTo: ' + fixed + ' row(s) renamed from William to Team A.');
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

  // Round 57: canonical short URL. Booking page reads ?p=<phone> and resolves
  // name/group/existingAppt via lg-availability. Matches every other booking-link
  // callsite migrated in rounds 18-24.
  const url = 'https://leakguard.my/appointment/' + (phone ? ('?p=' + encodeURIComponent(phone)) : '');

  const msg = 'Hi ' + (name || 'there') + ', here\'s the link to reschedule your site visit.\n\n' +
    'Check Real Time Availability / Book Your Slot Instantly: ' + url;

  try {
    UrlFetchApp.fetch(N8N_WAGROUP_URL, {
      method: 'post',
      contentType: 'application/json',
      headers: {'Authorization': 'Bearer ' + WHAPI_TOKEN},
      payload: JSON.stringify({to: groupId, body: msg, typing_time: 2}),
      muteHttpExceptions: true
    });
    return jsonResponse({status: 'ok', url: url});
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
    'Pending Balance','Completed','Job Complete','Receipt Sent',
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
        // Round 60.3: forward-fill invite link if n8n didn't supply one.
        {
          let _link = g.inviteLink || '';
          if (!_link) { _link = fetchInviteLinkForGroup(g.groupId); if (_link) Utilities.sleep(800); }
          if (_link) setCellByHeader(sheet, row, 'Group Invite Link (AJ)', _link);
        }
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
            // Round 60.3: forward-fill invite link.
            {
              let _link = g.inviteLink || '';
              if (!_link) { _link = fetchInviteLinkForGroup(g.groupId); if (_link) Utilities.sleep(800); }
              if (_link) setCellByHeader(sheet, row, 'Group Invite Link (AJ)', _link);
            }
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
        // Round 60.3: forward-fill invite link.
        {
          let _link = g.inviteLink || '';
          if (!_link) { _link = fetchInviteLinkForGroup(g.groupId); if (_link) Utilities.sleep(800); }
          if (_link) setCellByHeader(sheet, newRow, 'Group Invite Link (AJ)', _link);
        }
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
// Round 60.3 — Whapi invite-link helper (reused by backfill + bulk-link)
// ================================================================
// Returns a full WhatsApp group invite URL, or '' on any failure.
// Handles: 429 with 1× backoff (8s), invite_code → full-URL construction,
// network errors. Never throws.

function fetchInviteLinkForGroup(groupId) {
  if (!groupId) return '';
  let resp, code = 0, text = '';
  let retried = false;
  while (true) {
    try {
      resp = UrlFetchApp.fetch(
        'https://gate.whapi.cloud/groups/' + encodeURIComponent(groupId) + '/invite',
        { method: 'get', headers: { 'Authorization': 'Bearer ' + WHAPI_TOKEN }, muteHttpExceptions: true }
      );
    } catch (e) { return ''; }
    code = resp.getResponseCode();
    text = resp.getContentText();
    if (code !== 429 || retried) break;
    Utilities.sleep(8000);
    retried = true;
  }
  if (code !== 200) return '';
  try {
    const body = JSON.parse(text);
    if (body.link) return String(body.link);
    if (body.invite_link) return String(body.invite_link);
    if (body.url) return String(body.url);
    if (body.invite_code) return 'https://chat.whatsapp.com/' + String(body.invite_code).trim();
  } catch (_) {}
  return '';
}

// ================================================================
// Round 60 — backfill Group Invite Link via Whapi (run ONCE manually)
// ================================================================
// Walks rows with Group ID but no Group Invite Link, fetches via the helper,
// writes back. Idempotent (skips already-populated rows). Also runs from a
// weekly trigger as a safety net for any group that slips through.

function backfillInviteLinks() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(function(h){ return String(h || '').trim(); });
  const phoneIdx       = headers.indexOf('Phone');
  const nameIdx        = headers.indexOf('Name');
  const statusIdx      = headers.indexOf('Status');
  const groupIdIdx     = headers.indexOf('Group ID (AB)');
  const inviteLinkIdx  = headers.indexOf('Group Invite Link (AJ)');

  if (groupIdIdx === -1 || inviteLinkIdx === -1) {
    Logger.log('backfillInviteLinks: missing required headers (Group ID (AB) or Group Invite Link (AJ))');
    return;
  }

  const todo = [];
  for (let i = 1; i < data.length; i++) {
    const groupId = String(data[i][groupIdIdx] || '').trim();
    const invite  = String(data[i][inviteLinkIdx] || '').trim();
    if (groupId && !invite) {
      todo.push({
        row: i + 1,
        groupId: groupId,
        name: String(data[i][nameIdx] || ''),
        status: String(data[i][statusIdx] || '')
      });
    }
  }
  Logger.log('backfillInviteLinks: ' + todo.length + ' row(s) need backfill');
  if (!todo.length) return;

  let okCount = 0, failCount = 0;
  const failures = [];

  for (let j = 0; j < todo.length; j++) {
    const c = todo[j];
    const link = fetchInviteLinkForGroup(c.groupId);
    if (link) {
      sheet.getRange(c.row, inviteLinkIdx + 1).setValue(link);
      okCount++;
      Logger.log('OK row ' + c.row + ' [' + c.status + '] ' + c.name + ' → ' + link);
    } else {
      failures.push({ row: c.row, name: c.name, status: c.status, groupId: c.groupId, error: 'helper returned empty (likely 403/network/parse fail — check group admin status)' });
      failCount++;
    }
    // 1500ms steady-state pause = 40 req/min, comfortably under Whapi limits.
    Utilities.sleep(1500);
  }

  Logger.log('---');
  Logger.log('backfillInviteLinks DONE: ' + okCount + ' ok, ' + failCount + ' failed');
  if (failures.length) Logger.log('Failures: ' + JSON.stringify(failures, null, 2));
}

// ================================================================
// Round 76 — Kanban login gate + roles
// ================================================================
// A "Users" tab is the single source of truth for both the kanban
// login check and (Phase 2/3) the WA bot's phone<->name mapping.
// Columns: Username | PIN | Name | Role | Phone | AssignName
//   Role       = 'admin' (sees all) | 'supervisor' (filtered view)
//   AssignName = exact string used in the lead 'Assigned To' column so
//                the kanban filter matches; blank for admin.
// This is a practical gate (validates credentials, role-filters the
// view) — NOT crypto-grade; the raw gviz URL is still readable.
// ================================================================

const USERS_SHEET_NAME = 'Users';

// Run ONCE from the Apps Script editor. Idempotent: re-running only adds
// missing pieces. Seeds an admin row + 2 example supervisor placeholders.
function bootstrapUsersSheet() {
  const ss = SpreadsheetApp.openById(LIVE_SHEET_ID);
  let sheet = ss.getSheetByName(USERS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(USERS_SHEET_NAME);
    sheet.getRange(1, 1, 1, 6).setValues([['Username', 'PIN', 'Name', 'Role', 'Phone', 'AssignName']]);
    sheet.setFrozenRows(1);
    // PIN + Phone as plain text so leading zeros / long numbers survive.
    sheet.getRange('B:B').setNumberFormat('@');
    sheet.getRange('E:E').setNumberFormat('@');
    Logger.log('bootstrapUsersSheet: created Users tab');
  }
  // Seed rows only if the sheet has no data rows yet.
  if (sheet.getLastRow() < 2) {
    sheet.getRange(2, 1, 3, 6).setValues([
      ['admin',   '9999', 'Admin',    'admin',      '60183639321', ''],
      ['osment',  '1234', 'Osment',   'supervisor', '',            'Osment'],
      ['william', '5678', 'William',  'supervisor', '',            'William'],
    ]);
    Logger.log('bootstrapUsersSheet: seeded admin + 2 example supervisor rows (edit PINs/phones/AssignName)');
  } else {
    Logger.log('bootstrapUsersSheet: rows already present, nothing to seed');
  }
}

// body: {action:'login', username, pin, secret}
// Returns {status:'ok', name, role, assignName, phone} on match,
//         {status:'error', message:'invalid'} otherwise.
// Never returns any other user's row.
function handleLogin(body) {
  const username = String(body.username || '').trim().toLowerCase();
  const pin = String(body.pin || '').trim();
  if (!username || !pin) {
    return jsonResponse({status: 'error', message: 'username and pin required'});
  }
  const sheet = SpreadsheetApp.openById(LIVE_SHEET_ID).getSheetByName(USERS_SHEET_NAME);
  if (!sheet) {
    return jsonResponse({status: 'error', message: 'Users tab missing — run bootstrapUsersSheet() once from the editor'});
  }
  const data = sheet.getDataRange().getValues();
  // header: 0=Username 1=PIN 2=Name 3=Role 4=Phone 5=AssignName
  for (let i = 1; i < data.length; i++) {
    const u = String(data[i][0] || '').trim().toLowerCase();
    if (u !== username) continue;
    const rowPin = String(data[i][1] || '').trim();
    if (rowPin !== pin) {
      return jsonResponse({status: 'error', message: 'invalid'});
    }
    return jsonResponse({
      status: 'ok',
      name: String(data[i][2] || '').trim(),
      role: String(data[i][3] || '').trim().toLowerCase(),
      phone: String(data[i][4] || '').trim(),
      assignName: String(data[i][5] || '').trim(),
    });
  }
  return jsonResponse({status: 'error', message: 'invalid'});
}

// ================================================================
// Round 76 Phase 2 — Internal-staff job queries + daily WA reminder
// ================================================================
// handleStaffJobs is the web endpoint Phase 3's WA text-command flow
// will also call. The daily reminder runs entirely here via a
// time-driven trigger (no n8n) — reusing the Whapi send pattern from
// handleSendReschedule and the Users tab from Phase 1.

// Statuses we DON'T remind about / surface (archived + final).
const STAFF_DONE_STATUSES = [
  'Completed', 'Job Complete', 'Receipt Sent',
  'Lost', 'Cold Lead', 'Rejected', 'Out of Area', 'Human Handoff'
];

// Pure filter over already-read leads data.
//   data: sheet.getDataRange().getValues() (row 0 = headers)
//   h:    getHeaders(sheet) → {headers, colByName}
// Returns {assigned:[card...], repair:[card...]} where each card is
//   {name, phone, status, notes, groupLink, groupName}.
// A card can appear in BOTH lists (assigned to me AND repair-tagged).
function _filterStaffJobs(data, h, assignName) {
  const cName   = h.colByName['Name'];
  const cPhone  = h.colByName['Phone'];
  const cStatus = h.colByName['Status'];
  const cAssign = h.colByName['Assigned To'];
  const cTags   = h.colByName['Tags'];
  const cNotes  = h.colByName['Notes'];
  const cLink   = h.colByName['Group Invite Link (AJ)'];
  const cGName  = h.colByName['Group Name (AE)'];
  const cQtDate = h.colByName['Date QT Issued'];  // Round 79
  const want = String(assignName || '').trim();

  const assigned = [];
  const repair = [];
  const qtFollowups = [];  // Round 79: Quotation Sent + ≥3 days call list
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = cStatus ? String(row[cStatus - 1] || '').trim() : '';
    if (STAFF_DONE_STATUSES.indexOf(status) !== -1) continue;  // skip archived/final
    const card = {
      name:      cName  ? String(row[cName  - 1] || '').trim() : '',
      phone:     cPhone ? String(row[cPhone - 1] || '').trim() : '',
      status:    status,
      notes:     cNotes ? String(row[cNotes - 1] || '').trim() : '',
      groupLink: cLink  ? String(row[cLink  - 1] || '').trim() : '',
      groupName: cGName ? String(row[cGName - 1] || '').trim() : ''
    };
    const tags = cTags ? String(row[cTags - 1] || '').toLowerCase() : '';
    const isMine = !!(want && cAssign && String(row[cAssign - 1] || '').trim() === want);
    if (tags.indexOf('repair') !== -1) repair.push(card);
    if (isMine) assigned.push(card);
    // Round 79 — QT follow-up: my QS card aged ≥3 days since QT issued.
    if (isMine && status === 'Quotation Sent' && cQtDate) {
      const days = _daysSinceMyt(row[cQtDate - 1]);
      if (days >= 3) qtFollowups.push(Object.assign({}, card, {daysSinceQt: days}));
    }
  }
  return {assigned: assigned, repair: repair, qtFollowups: qtFollowups};
}

// body: {action:'staffJobs', phone OR name, secret}
// Resolves the Users row (by phone last-8 or exact name), then returns
// that person's assigned + repair cards. Used by Phase 3 WA commands.
function handleStaffJobs(body) {
  const phone = String(body.phone || '').replace(/\D/g, '');
  const name  = String(body.name || '').trim().toLowerCase();
  if (!phone && !name) {
    return jsonResponse({status: 'error', message: 'phone or name required'});
  }
  const usersSheet = SpreadsheetApp.openById(LIVE_SHEET_ID).getSheetByName(USERS_SHEET_NAME);
  if (!usersSheet) return jsonResponse({status: 'error', message: 'Users tab missing'});

  const u = usersSheet.getDataRange().getValues();
  // header: 0=Username 1=PIN 2=Name 3=Role 4=Phone 5=AssignName
  const last8 = phone.length >= 8 ? phone.slice(-8) : phone;
  let urow = null;
  for (let i = 1; i < u.length; i++) {
    const uPhone = String(u[i][4] || '').replace(/\D/g, '');
    const uName  = String(u[i][2] || '').trim().toLowerCase();
    if (phone && uPhone && (uPhone === phone || (uPhone.length >= 8 && uPhone.endsWith(last8)))) { urow = u[i]; break; }
    if (name && uName && uName === name) { urow = u[i]; break; }
  }
  if (!urow) return jsonResponse({status: 'error', message: 'not_recognised'});

  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const h = getHeaders(sheet);
  const jobs = _filterStaffJobs(data, h, String(urow[5] || '').trim());
  return jsonResponse({
    status: 'ok',
    name: String(urow[2] || '').trim(),
    role: String(urow[3] || '').trim().toLowerCase(),
    assigned: jobs.assigned,
    repair: jobs.repair
  });
}

// body: {action:'staffCommand', phone, command, secret}
// WA text-command handler (Phase 3). Resolves the sender via the Users
// tab, queries their jobs, and sends the WhatsApp reply itself (no n8n
// reply node needed). command ∈ '-myjobs' | '-pending' (else → help).
function handleStaffCommand(body) {
  const phone = String(body.phone || '').replace(/\D/g, '');
  const command = String(body.command || '').trim().toLowerCase();
  if (!phone) return jsonResponse({status: 'error', message: 'phone required'});

  const usersSheet = SpreadsheetApp.openById(LIVE_SHEET_ID).getSheetByName(USERS_SHEET_NAME);
  if (!usersSheet) return jsonResponse({status: 'error', message: 'Users tab missing'});
  const u = usersSheet.getDataRange().getValues();
  const last8 = phone.length >= 8 ? phone.slice(-8) : phone;
  let urow = null;
  for (let i = 1; i < u.length; i++) {
    const uPhone = String(u[i][4] || '').replace(/\D/g, '');
    if (uPhone && (uPhone === phone || (uPhone.length >= 8 && uPhone.endsWith(last8)))) { urow = u[i]; break; }
  }
  if (!urow) {
    _sendWhapiText(phone, "Sorry, this number isn't recognised as Leak Guard staff.");
    return jsonResponse({status: 'ok', recognised: false});
  }

  const name = String(urow[2] || '').trim();
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const h = getHeaders(sheet);
  const jobs = _filterStaffJobs(data, h, String(urow[5] || '').trim());

  let msg;
  if (command === '-myjobs') {
    msg = _buildCommandMessage('📋 Your assigned jobs', jobs.assigned, 'You have no assigned jobs right now. 👍');
  } else if (command === '-pending') {
    msg = _buildCommandMessage('🔧 Pending repair queue', jobs.repair, 'No pending repair jobs right now. 👍');
  } else {
    msg = 'Hi ' + (name || 'there') + '! Commands:\n' +
      '• *-myjobs* — your assigned jobs\n' +
      '• *-pending* — pending repair queue';
  }
  _sendWhapiText(phone, msg);
  return jsonResponse({status: 'ok', recognised: true, command: command});
}

// Build a command reply: "title (n)" + card lines, or an empty-state line.
function _buildCommandMessage(title, cards, emptyMsg) {
  if (!cards.length) return title + ' (0)\n\n' + emptyMsg;
  const lines = [title + ' (' + cards.length + ')', ''];
  cards.forEach(function(c) { _appendCardLines(lines, c); });
  return lines.join('\n');
}

// MYT YYYY-MM-DD (gotcha #3 idiom — getUTC* on a +8h-shifted timestamp).
function _mytDateStr() {
  const m = new Date(Date.now() + 8 * 60 * 60 * 1000);
  return m.getUTCFullYear() + '-' +
    String(m.getUTCMonth() + 1).padStart(2, '0') + '-' +
    String(m.getUTCDate()).padStart(2, '0');
}

// Round 79 — integer days from a YYYY-MM-DD MYT date (or Date cell) to
// today (MYT). Returns -1 if the input isn't a parseable YMD. Reuses
// _normalizeDateStr so a Date cell or a 'YYYY-MM-DD' string both work.
function _daysSinceMyt(v) {
  const s = _normalizeDateStr(v);
  const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return -1;
  const past = Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const t = new Date(Date.now() + 8 * 60 * 60 * 1000);
  const today = Date.UTC(t.getUTCFullYear(), t.getUTCMonth(), t.getUTCDate());
  return Math.floor((today - past) / 86400000);
}

// Trigger target — installed to fire every 30 min; self-gates to the
// 8 AM MYT hour and sends at most once per MYT day.
function sendSupervisorDailyReminders() {
  const mytHour = new Date(Date.now() + 8 * 60 * 60 * 1000).getUTCHours();
  if (mytHour !== 8) return;  // only the 8 AM MYT hour
  const props = PropertiesService.getScriptProperties();
  const today = _mytDateStr();
  if (props.getProperty('lastReminderDate') === today) return;  // already sent today
  const result = _doSupervisorReminders();
  props.setProperty('lastReminderDate', today);
  Logger.log('sendSupervisorDailyReminders: ' + JSON.stringify(result));
}

// Manual test from the editor — bypasses the hour + daily guards so it
// always sends right now.
function testSupervisorReminders() {
  Logger.log('testSupervisorReminders: ' + JSON.stringify(_doSupervisorReminders()));
}

// Core: DM each supervisor (Users tab, role=supervisor, has phone) their
// assigned + repair cards. Skips supervisors with nothing to show.
function _doSupervisorReminders() {
  const ss = SpreadsheetApp.openById(LIVE_SHEET_ID);
  const usersSheet = ss.getSheetByName(USERS_SHEET_NAME);
  if (!usersSheet) return {error: 'Users tab missing'};
  const u = usersSheet.getDataRange().getValues();

  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const h = getHeaders(sheet);

  let sent = 0, skipped = 0;
  for (let i = 1; i < u.length; i++) {
    if (String(u[i][3] || '').trim().toLowerCase() !== 'supervisor') continue;
    const phone = String(u[i][4] || '').replace(/\D/g, '');
    if (!phone) { skipped++; continue; }
    const jobs = _filterStaffJobs(data, h, String(u[i][5] || '').trim());
    if (!jobs.assigned.length && !jobs.repair.length && (!jobs.qtFollowups || !jobs.qtFollowups.length)) { skipped++; continue; }
    _sendWhapiText(phone, _buildReminderMessage(String(u[i][2] || '').trim(), jobs));
    sent++;
    Utilities.sleep(600);  // gentle pacing between sends
  }
  return {sent: sent, skipped: skipped};
}

// Round 79.1 — formatted layout: WhatsApp *bold* section titles,
// numbered items, double-blank-line section gaps, customer name bolded.
function _buildReminderMessage(name, jobs) {
  const lines = [];
  lines.push('🔔 *Good morning ' + (name || 'there') + '*');
  lines.push('_Your jobs today_');

  if (jobs.assigned.length) {
    lines.push('', '', '📋 *ASSIGNED — ' + jobs.assigned.length + '*', '');
    jobs.assigned.forEach(function(c, i) {
      lines.push((i + 1) + '. *' + (c.name || '(no name)') + '* — ' + c.status);
      if (c.notes)     lines.push('    📝 ' + _snippet(c.notes, 140));
      if (c.groupLink) lines.push('    🔗 ' + c.groupLink);
      if (i < jobs.assigned.length - 1) lines.push('');
    });
  }

  // Round 79 — QT follow-up: Quotation Sent + ≥3 days since QT issued.
  if (jobs.qtFollowups && jobs.qtFollowups.length) {
    lines.push('', '', '📞 *QT FOLLOW-UP — 3+ days — ' + jobs.qtFollowups.length + '*', '');
    jobs.qtFollowups.forEach(function(c, i) {
      lines.push((i + 1) + '. *' + (c.name || '(no name)') + '* — sent ' + c.daysSinceQt + ' days ago');
      if (c.phone)     lines.push('    📞 ' + c.phone);
      if (c.groupLink) lines.push('    🔗 ' + c.groupLink);
      if (c.notes)     lines.push('    📝 ' + _snippet(c.notes, 140));
      if (i < jobs.qtFollowups.length - 1) lines.push('');
    });
  }

  if (jobs.repair.length) {
    lines.push('', '', '🔧 *REPAIR QUEUE — ' + jobs.repair.length + '*', '');
    jobs.repair.forEach(function(c, i) {
      lines.push((i + 1) + '. *' + (c.name || '(no name)') + '* — ' + c.status);
      if (c.notes)     lines.push('    📝 ' + _snippet(c.notes, 140));
      if (c.groupLink) lines.push('    🔗 ' + c.groupLink);
      if (i < jobs.repair.length - 1) lines.push('');
    });
  }
  return lines.join('\n');
}

function _appendCardLines(lines, c) {
  lines.push('• ' + (c.name || '(no name)') + ' — ' + c.status);
  if (c.notes)     lines.push('  📝 ' + _snippet(c.notes, 140));
  if (c.groupLink) lines.push('  🔗 ' + c.groupLink);
}

function _snippet(s, max) {
  s = String(s || '').replace(/\s+/g, ' ').trim();
  return s.length > max ? s.slice(0, max - 1) + '…' : s;
}

function _sendWhapiText(to, body) {
  try {
    UrlFetchApp.fetch(N8N_WAGROUP_URL, {
      method: 'post',
      contentType: 'application/json',
      headers: {'Authorization': 'Bearer ' + WHAPI_TOKEN},
      payload: JSON.stringify({to: to, body: body}),
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log('_sendWhapiText error to ' + to + ': ' + e.toString());
  }
}

// Run ONCE from the editor — installs the time trigger that drives
// sendSupervisorDailyReminders (which self-gates to 8 AM MYT). Idempotent:
// removes any existing trigger for that function first.
function bootstrapDailyReminderTrigger() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendSupervisorDailyReminders') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  ScriptApp.newTrigger('sendSupervisorDailyReminders')
    .timeBased()
    .everyMinutes(30)
    .create();
  Logger.log('bootstrapDailyReminderTrigger: removed ' + removed + ' old trigger(s), installed 30-min trigger (self-gates to 8 AM MYT)');
}

// ================================================================
// Round 78 — Next-day appointment reminder (7 PM MYT)
// ================================================================
// Personal WA DM to each rep listing tomorrow's Site-Visit-Confirmed
// jobs assigned to them, sorted by appointment time. Mirrors the
// Phase 2 daily-reminder pattern (hour gate + once-per-day guard +
// 30-min trigger). Reuses _sendWhapiText / _snippet / _mytDateStr.

// Sheets cells for Date Appt Confirmed may come back as either a
// 'YYYY-MM-DD' string (n8n writes string) or a Date object (manual
// edits) — normalise either to 'YYYY-MM-DD' using its local Y/M/D.
function _normalizeDateStr(v) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return v.getFullYear() + '-' +
      String(v.getMonth() + 1).padStart(2, '0') + '-' +
      String(v.getDate()).padStart(2, '0');
  }
  return String(v || '').trim().slice(0, 10);
}

// Slot Chosen looks like 'Wednesday, 3 June 2026, 9:00 AM - 10:00 AM'.
// Last comma-piece = the time range we want to show in the reminder.
function _extractSlotTime(slot) {
  const parts = String(slot || '').split(',');
  return parts[parts.length - 1].trim();
}

// Minutes since midnight for the FIRST time in the string ('9:00 AM').
// Used to sort tomorrow's appointments chronologically.
function _slotTimeKey(t) {
  const m = String(t || '').match(/(\d{1,2}):(\d{2})\s*(AM|PM)/i);
  if (!m) return 9999;
  let h = parseInt(m[1], 10);
  const min = parseInt(m[2], 10);
  const ampm = m[3].toUpperCase();
  if (ampm === 'PM' && h !== 12) h += 12;
  if (ampm === 'AM' && h === 12) h = 0;
  return h * 60 + min;
}

function _formatHumanDate(yyyyMmDd) {
  const m = String(yyyyMmDd || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return yyyyMmDd;
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const days = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return days[d.getDay()] + ', ' + d.getDate() + ' ' + months[d.getMonth()] + ' ' + d.getFullYear();
}

// Pure filter: tomorrow's Site-Visit-Confirmed leads assigned to this
// rep, sorted by appointment time.
function _filterNextDayAppointments(data, h, assignName, tomorrowMyt) {
  const want = String(assignName || '').trim();
  if (!want) return [];
  const cName    = h.colByName['Name'];
  const cPhone   = h.colByName['Phone'];
  const cStatus  = h.colByName['Status'];
  const cAssign  = h.colByName['Assigned To'];
  const cDateApt = h.colByName['Date Appt Confirmed'];
  const cSlot    = h.colByName['Slot Chosen'];
  const cAddr    = h.colByName['Full Address'];
  const cLoc     = h.colByName['Location'];
  const cNotes   = h.colByName['Notes'];
  const cLink    = h.colByName['Group Invite Link (AJ)'];
  const cGName   = h.colByName['Group Name (AE)'];

  const items = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = cStatus ? String(row[cStatus - 1] || '').trim() : '';
    if (status !== 'Site Visit Confirmed') continue;
    const dateApt = cDateApt ? _normalizeDateStr(row[cDateApt - 1]) : '';
    if (dateApt !== tomorrowMyt) continue;
    const assigned = cAssign ? String(row[cAssign - 1] || '').trim() : '';
    if (assigned !== want) continue;
    const slot = cSlot ? String(row[cSlot - 1] || '').trim() : '';
    items.push({
      time:      _extractSlotTime(slot),
      slot:      slot,
      name:      cName  ? String(row[cName  - 1] || '').trim() : '',
      phone:     cPhone ? String(row[cPhone - 1] || '').trim() : '',
      address:   cAddr  ? String(row[cAddr  - 1] || '').trim() : '',
      location:  cLoc   ? String(row[cLoc   - 1] || '').trim() : '',
      notes:     cNotes ? String(row[cNotes - 1] || '').trim() : '',
      groupLink: cLink  ? String(row[cLink  - 1] || '').trim() : '',
      groupName: cGName ? String(row[cGName - 1] || '').trim() : ''
    });
  }
  items.sort(function(a, b) { return _slotTimeKey(a.time) - _slotTimeKey(b.time); });
  return items;
}

// Round 79.1 — formatted layout: WhatsApp *bold* header, numbered items,
// blank line between items, 4-space indented sub-details.
function _buildNextDayMessage(repName, tomorrowMyt, items) {
  const dt = _formatHumanDate(tomorrowMyt);
  const lines = [];
  lines.push('📅 *Your visits tomorrow*');
  lines.push('_' + dt + '_  —  ' + items.length + (items.length === 1 ? ' appointment' : ' appointments'));
  lines.push('', '');
  items.forEach(function(it, i) {
    lines.push((i + 1) + '. *' + (it.time || it.slot || 'time ?') + ' — ' + (it.name || '(no name)') + '*');
    if (it.address)       lines.push('    📍 ' + it.address);
    else if (it.location) lines.push('    📍 ' + it.location);
    if (it.phone)         lines.push('    📞 ' + it.phone);
    if (it.notes)         lines.push('    📝 ' + _snippet(it.notes, 140));
    if (it.groupLink)     lines.push('    🔗 ' + it.groupLink);
    if (i < items.length - 1) lines.push('');
  });
  return lines.join('\n');
}

// Core: DM each supervisor their next-day Site-Visit-Confirmed jobs.
// Skips reps with nothing tomorrow (no spammy empty DMs).
function _doNextDayReminders() {
  const ss = SpreadsheetApp.openById(LIVE_SHEET_ID);
  const usersSheet = ss.getSheetByName(USERS_SHEET_NAME);
  if (!usersSheet) return {error: 'Users tab missing'};
  const u = usersSheet.getDataRange().getValues();

  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const h = getHeaders(sheet);

  // Tomorrow in MYT (gotcha #3 idiom: +8h shift, then getUTC*).
  const t = new Date(Date.now() + 8 * 60 * 60 * 1000 + 24 * 60 * 60 * 1000);
  const tomorrowMyt = t.getUTCFullYear() + '-' +
    String(t.getUTCMonth() + 1).padStart(2, '0') + '-' +
    String(t.getUTCDate()).padStart(2, '0');

  let sent = 0, skipped = 0;
  for (let i = 1; i < u.length; i++) {
    if (String(u[i][3] || '').trim().toLowerCase() !== 'supervisor') continue;
    const phone = String(u[i][4] || '').replace(/\D/g, '');
    if (!phone) { skipped++; continue; }
    const items = _filterNextDayAppointments(data, h, String(u[i][5] || '').trim(), tomorrowMyt);
    if (!items.length) { skipped++; continue; }
    _sendWhapiText(phone, _buildNextDayMessage(String(u[i][2] || '').trim(), tomorrowMyt, items));
    sent++;
    Utilities.sleep(600);
  }
  return {sent: sent, skipped: skipped, tomorrowMyt: tomorrowMyt};
}

// Trigger target — installed to fire every 30 min; self-gates to the
// 7 PM MYT hour and sends at most once per MYT day.
function sendNextDayAppointmentReminders() {
  const mytHour = new Date(Date.now() + 8 * 60 * 60 * 1000).getUTCHours();
  if (mytHour !== 19) return;  // 7 PM MYT only
  const props = PropertiesService.getScriptProperties();
  const today = _mytDateStr();
  if (props.getProperty('lastNextDayReminderDate') === today) return;
  const result = _doNextDayReminders();
  props.setProperty('lastNextDayReminderDate', today);
  Logger.log('sendNextDayAppointmentReminders: ' + JSON.stringify(result));
}

// Manual test — bypasses the hour + daily guards so it always sends now.
function testNextDayAppointmentReminders() {
  Logger.log('testNextDayAppointmentReminders: ' + JSON.stringify(_doNextDayReminders()));
}

// Run ONCE from the editor — installs the 30-min time trigger that
// drives sendNextDayAppointmentReminders (which self-gates to 7 PM MYT).
function bootstrapNextDayReminderTrigger() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendNextDayAppointmentReminders') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  ScriptApp.newTrigger('sendNextDayAppointmentReminders')
    .timeBased()
    .everyMinutes(30)
    .create();
  Logger.log('bootstrapNextDayReminderTrigger: removed ' + removed + ' old trigger(s), installed 30-min trigger (self-gates to 7 PM MYT)');
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
