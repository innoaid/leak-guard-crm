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
  // body: {phone, archiveStatus, secret}  — archiveStatus = Lost / Cold Lead / etc.
  return handleUpdateStatus({
    phone: body.phone,
    status: body.archiveStatus || 'Lost',
    changedBy: body.changedBy || 'Kanban (archive)'
  });
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

  // Shorten via TinyURL (best-effort)
  let shortUrl = longUrl;
  try {
    const r = UrlFetchApp.fetch('https://tinyurl.com/api-create.php?url=' +
      encodeURIComponent(longUrl), {muteHttpExceptions: true});
    const s = String(r.getContentText()).trim();
    if (s.indexOf('http') === 0) shortUrl = s;
  } catch (_e) {}

  const msg = 'Click the link below to manage your appointment — only takes 2 mins.\n\n' +
    '🚀 Express Booking: ' + shortUrl;

  try {
    UrlFetchApp.fetch(N8N_WAGROUP_URL, {
      method: 'post',
      contentType: 'application/json',
      headers: {'Authorization': 'Bearer ' + WHAPI_TOKEN},
      payload: JSON.stringify({to: groupId, body: msg, typing_time: 2}),
      muteHttpExceptions: true
    });
    return jsonResponse({status: 'ok', shortUrl: shortUrl});
  } catch (err) {
    return jsonResponse({status: 'error', message: 'whapi: ' + err.toString()});
  }
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
