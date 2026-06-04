// ================================================================
// Round 83 — AutoCount sync for the Estimation Builder
// ================================================================
// After an estimation (SE-MMYY-NNN) is sent, the inspector/admin can
// sync it to AutoCount Cloud: find-or-create the debtor + create a
// quotation. Both happen inside the existing n8n workflow behind
// /webhook/lg-quotation-create (same one quotation_builder.html uses),
// so NO AutoCount credentials live in Apps Script.
//
// Two triggers, one code path:
//   A) estimation_builder.html — "Also create AutoCount quotation"
//      checkbox at send time (calls syncAutocount with seNo)
//   B) team_kanban.html — "Sync to AutoCount" button on the lead
//      modal (calls syncAutocount with phone; newest SE is used)
//
// Trigger B can fire days later from a different device, so every
// estimation send first persists its summary server-side via the
// saveEstimation action (Estimations tab, one row per SE). The
// membrane rate is per-device localStorage — the server can never
// recompute totals, which is why the frontend sends the final
// webhook-ready lineItems[] and totals.
//
// Idempotency: the Estimations row stores the AutoCount QT No once
// assigned. A re-sync of the same SE returns the existing QT (no
// duplicate quotation) but still re-applies the CRM side-effects
// (group rename / value tag / lead columns) — cheap and self-healing.
// A revision = new SE number = new row = new QT; the group name
// slash-appends the new code via _buildQtGroupName (Round 70).
//
// One-time setup (run from the editor): bootstrapEstimationsSheet(),
// bootstrapAutocountLeadColumns().
// ================================================================

const QT_CREATE_URL = 'https://leakguard.app.n8n.cloud/webhook/lg-quotation-create'; // debtor find-or-create + QT create in AutoCount

const ESTIMATIONS_SHEET = 'Estimations';
const ESTIMATIONS_HEADERS = [
  'SE No', 'Phone', 'Name', 'Timestamp', 'Submitted By', 'Line Items JSON',
  'Subtotal', 'Discount Type', 'Discount Value', 'Discount Amount',
  'MOQ Adjustment', 'Grand Total', 'Total Sqft', 'Membrane Rate',
  'AutoCount QT No', 'AutoCount Debtor Code', 'Synced At', 'Sync Status',
  'PDF URL',
];

// ================================================================
// saveEstimation — persist the SE summary (called on EVERY send)
// ================================================================
// body: {action, secret, seNo, phone, name, submittedBy,
//        lineItems: [{desc, qty, unit, price, furtherDescription}],
//        subtotal, discountType, discountValue, discountAmount,
//        moqAdjustment, grandTotal, totalSqft, membraneRate}
// Upserts by SE No: a re-save of the same SE overwrites the summary
// columns but never touches the QT No / Synced audit columns.
function handleSaveEstimation(body) {
  if (!body.seNo || !body.phone) {
    return jsonResponse({status: 'error', message: 'seNo and phone required'});
  }
  const sheet = SpreadsheetApp.openById(LIVE_SHEET_ID).getSheetByName(ESTIMATIONS_SHEET);
  if (!sheet) return jsonResponse({status: 'error', message: 'Estimations sheet missing — run bootstrapEstimationsSheet() once from the editor'});

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const h = getHeaders(sheet);
    let rowNum = _findEstimationRowBySe(sheet, h, body.seNo);
    if (!rowNum) rowNum = Math.max(sheet.getLastRow(), 1) + 1;

    const rowObj = {
      'SE No':           String(body.seNo).trim(),
      'Phone':           String(body.phone).trim(),
      'Name':            body.name || '',
      'Timestamp':       new Date().toISOString(),
      'Submitted By':    body.submittedBy || '',
      'Line Items JSON': JSON.stringify(body.lineItems || []),
      'Subtotal':        Number(body.subtotal) || 0,
      'Discount Type':   body.discountType || '',
      'Discount Value':  Number(body.discountValue) || 0,
      'Discount Amount': Number(body.discountAmount) || 0,
      'MOQ Adjustment':  Number(body.moqAdjustment) || 0,
      'Grand Total':     Number(body.grandTotal) || 0,
      'Total Sqft':      Number(body.totalSqft) || 0,
      'Membrane Rate':   Number(body.membraneRate) || 0,
      // AutoCount QT No / Debtor Code / Synced At / Sync Status are
      // deliberately NOT in this map — re-saves must not clear them.
    };
    Object.keys(rowObj).forEach(function(k) {
      if (h.colByName[k]) sheet.getRange(rowNum, h.colByName[k]).setValue(rowObj[k]);
    });
    return jsonResponse({status: 'ok', seNo: body.seNo, rowNum: rowNum});
  } finally {
    lock.releaseLock();
  }
}

// ================================================================
// syncAutocount — create the AutoCount QT + apply CRM side-effects
// ================================================================
// body: {action, secret, seNo}            (Trigger A — estimation builder)
//   or  {action, secret, phone}           (Trigger B — kanban; newest SE wins)
function handleSyncAutocount(body) {
  const sheet = SpreadsheetApp.openById(LIVE_SHEET_ID).getSheetByName(ESTIMATIONS_SHEET);
  if (!sheet) return jsonResponse({status: 'error', message: 'Estimations sheet missing — run bootstrapEstimationsSheet() once from the editor'});
  const h = getHeaders(sheet);

  let seNo = String(body.seNo || '').trim();
  if (!seNo && body.phone) {
    seNo = _latestSeForPhone(sheet, h, body.phone);
    if (!seNo) return jsonResponse({status: 'error', message: 'no estimation on file for this lead — generate one in the Estimation Builder first'});
  }
  if (!seNo) return jsonResponse({status: 'error', message: 'seNo or phone required'});

  // Lock covers the webhook round-trip so a double-click cannot create
  // two quotations for the same SE before the guard column is written.
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const rowNum = _findEstimationRowBySe(sheet, h, seNo);
    if (!rowNum) return jsonResponse({status: 'error', message: 'estimation not found: ' + seNo});

    const get = function(name) {
      return h.colByName[name] ? sheet.getRange(rowNum, h.colByName[name]).getValue() : '';
    };
    const phone = String(get('Phone') || '').trim();
    const name  = String(get('Name') || '').trim();

    // ── Idempotency guard: QT already created for this SE ──
    let qtNo = String(get('AutoCount QT No') || '').trim();
    let debtorCode = String(get('AutoCount Debtor Code') || '').trim();
    let agentWarning = '';
    const alreadySynced = !!qtNo;

    if (!alreadySynced) {
      let lineItems;
      try {
        lineItems = JSON.parse(String(get('Line Items JSON') || '[]'));
      } catch (e) {
        return jsonResponse({status: 'error', message: 'corrupt Line Items JSON for ' + seNo});
      }
      if (!lineItems.length) return jsonResponse({status: 'error', message: 'estimation ' + seNo + ' has no line items'});

      // Address comes from the lead row (not stored per-SE)
      const leadSheet = getSheet();
      const leadRow = findRowByPhone(leadSheet, phone);
      let address = '';
      if (leadRow) {
        const lh = getHeaders(leadSheet);
        const addrCol = lh.colByName['Full Address'] || lh.colByName['Location'];
        if (addrCol) address = String(leadSheet.getRange(leadRow, addrCol).getValue() || '').trim();
      }

      // Same payload shape quotation_builder.html sends (Phase 3 webhook).
      // Round 83.2 — the inspector (Submitted By) rides along as salesAgent;
      // the n8n workflow normalizes it ("William (KL)" -> "William") and
      // auto-creates the AutoCount agent when missing.
      const resp = UrlFetchApp.fetch(QT_CREATE_URL, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          phone:         phone,
          name:          name,
          address:       address,
          serviceHeader: 'Torch-On Membrane Waterproofing Services',
          salesAgent:    String(get('Submitted By') || '').trim(),
          lineItems:     lineItems,
        }),
        muteHttpExceptions: true,
      });
      const code = resp.getResponseCode();
      if (code !== 200) {
        _markSyncFailed(sheet, h, rowNum, 'HTTP ' + code);
        return jsonResponse({status: 'error', message: 'lg-quotation-create HTTP ' + code + ': ' + resp.getContentText().slice(0, 200)});
      }
      let acData;
      try { acData = JSON.parse(resp.getContentText()); } catch (e) { acData = null; }
      if (!acData || !acData.success || !acData.docNo) {
        const msg = (acData && acData.error) ? acData.error : String(resp.getContentText()).slice(0, 200);
        _markSyncFailed(sheet, h, rowNum, msg);
        return jsonResponse({status: 'error', message: 'AutoCount create failed: ' + msg});
      }
      qtNo = String(acData.docNo).trim();
      debtorCode = String(acData.debtorCode || '').trim();
      agentWarning = String(acData.agentWarning || '').trim();

      // Persist the guard BEFORE side-effects so a partial failure
      // downstream can never re-create the quotation on retry.
      if (h.colByName['AutoCount QT No'])        sheet.getRange(rowNum, h.colByName['AutoCount QT No']).setValue(qtNo);
      if (h.colByName['AutoCount Debtor Code'])  sheet.getRange(rowNum, h.colByName['AutoCount Debtor Code']).setValue(debtorCode);
      if (h.colByName['Synced At'])              sheet.getRange(rowNum, h.colByName['Synced At']).setValue(new Date().toISOString());
      if (h.colByName['Sync Status'])            sheet.getRange(rowNum, h.colByName['Sync Status']).setValue('synced');
    }

    // ── Side-effects (idempotent; safe to re-run) ──
    const grandRm   = Math.round(Number(get('Grand Total')) || 0);
    const totalSqft = Math.round(Number(get('Total Sqft')) || 0);
    const sideEffects = _applyCrmSideEffects(phone, qtNo, grandRm, totalSqft);

    return jsonResponse({
      status: 'ok', seNo: seNo, qtNo: qtNo, debtorCode: debtorCode,
      alreadySynced: alreadySynced, agentWarning: agentWarning, sideEffects: sideEffects,
    });
  } catch (err) {
    return jsonResponse({status: 'error', message: 'syncAutocount: ' + err.toString()});
  } finally {
    lock.releaseLock();
  }
}

// ================================================================
// Helpers
// ================================================================

// Find the Estimations row (1-based) for an SE No, or null.
function _findEstimationRowBySe(sheet, h, seNo) {
  const seCol = h.colByName['SE No'];
  if (!seCol) return null;
  const target = String(seNo).trim();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][seCol - 1] || '').trim() === target) return i + 1;
  }
  return null;
}

// Newest SE No for a phone (last-8-digit match, same idiom as
// findRowByPhone). "Newest" = greatest Timestamp.
function _latestSeForPhone(sheet, h, phone) {
  const seCol = h.colByName['SE No'];
  const phoneCol = h.colByName['Phone'];
  const tsCol = h.colByName['Timestamp'];
  if (!seCol || !phoneCol) return '';
  const target = String(phone).trim();
  const last8 = target.length >= 8 ? target.slice(-8) : target;

  const data = sheet.getDataRange().getValues();
  let best = '', bestTs = '';
  for (let i = 1; i < data.length; i++) {
    const cell = String(data[i][phoneCol - 1] || '').trim();
    const match = (cell === target) || (cell.length >= 8 && cell.endsWith(last8));
    if (!match) continue;
    const ts = tsCol ? String(data[i][tsCol - 1] || '') : '';
    if (!best || ts >= bestTs) { best = String(data[i][seCol - 1] || '').trim(); bestTs = ts; }
  }
  return best;
}

// 'QT-0526-087' -> '0526-087' (the shape _buildQtGroupName expects)
function _acQtCode(docNo) {
  const m = String(docNo || '').match(/(\d{4})-(\d{3})\b/);
  return m ? (m[1] + '-' + m[2]) : null;
}

function _markSyncFailed(sheet, h, rowNum, msg) {
  try {
    if (h.colByName['Sync Status']) sheet.getRange(rowNum, h.colByName['Sync Status']).setValue('failed: ' + String(msg).slice(0, 80));
  } catch (_e) { /* best-effort */ }
}

// Group rename + value tag + lead-column writeback. Every step is
// idempotent and best-effort: a failure in one never blocks the rest.
function _applyCrmSideEffects(phone, qtNo, grandRm, totalSqft) {
  const sheet = getSheet();
  const rowNum = findRowByPhone(sheet, phone);
  if (!rowNum) return {error: 'lead not found for ' + phone};
  const h = getHeaders(sheet);
  const out = {};

  // (a) Group rename — QT-MMYY-NNN prefix, slash-append on revisions
  // (identical pattern to the Round 70 QT-PDF-detect block in
  // handleUpdateStatusByGroup).
  const qtCode = _acQtCode(qtNo);
  if (qtCode) {
    const gnameCol = h.colByName['Group Name (AE)'];
    const gidCol   = h.colByName['Group ID (AB)'];
    const groupName = gnameCol ? String(sheet.getRange(rowNum, gnameCol).getValue() || '').trim() : '';
    const groupId   = gidCol   ? String(sheet.getRange(rowNum, gidCol).getValue() || '').trim() : '';
    const newName = _buildQtGroupName(groupName, qtCode);
    if (newName && newName !== groupName) {
      try { setCellByHeader(sheet, rowNum, 'Group Name (AE)', newName); } catch (_e) {}
      if (groupId) {
        try {
          UrlFetchApp.fetch(N8N_RENAME_GROUP_URL, {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify({secret: SHARED_SECRET, groupId: groupId, newGroupName: newName}),
            muteHttpExceptions: true,
          });
        } catch (_e) { /* best-effort; CRM is source of truth either way */ }
      }
      out.groupName = newName;
    }
  }

  // (b) Value tag — replace any previous RM…sqft tag, then append the
  // fresh one. Comma-free on purpose: the Tags column is comma-split
  // everywhere, so the tag uses a middot separator and no thousands
  // separators (e.g. "RM12500 · 850sqft").
  try {
    const tagsCol = h.colByName['Tags'];
    if (tagsCol) {
      const newTag = 'RM' + grandRm + ' · ' + totalSqft + 'sqft';
      const cur = String(sheet.getRange(rowNum, tagsCol).getValue() || '').trim();
      const list = cur.split(',')
        .map(function(t) { return t.trim(); })
        .filter(function(t) { return t && !/^RM\d.*sqft$/i.test(t); });
      list.push(newTag);
      sheet.getRange(rowNum, tagsCol).setValue(list.join(','));
      out.tag = newTag;
    }
  } catch (_e) { /* best-effort */ }

  // (c) Lead-column writeback
  try { setCellByHeader(sheet, rowNum, 'AutoCount QT No', qtNo); out.qtNoWritten = true; } catch (_e) {}
  try { setCellByHeader(sheet, rowNum, 'Quotation (RM)', grandRm); } catch (_e) {}
  try { setCellByHeader(sheet, rowNum, 'Total Sqft', totalSqft); } catch (_e) {}

  return out;
}

// ================================================================
// Round 83.3 — store the estimation PDF in Drive
// ================================================================
// The builder renders the PDF client-side and the only copy used to
// live in the customer's WA chat. This handler archives it: PDF goes
// into a Drive folder, the link lands in the Estimations row ('PDF
// URL') and the lead row ('Quotation PDF URL', shown on the kanban
// lead modal). Best-effort: the send flow never depends on it.

// Lazy find-or-create (no editor bootstrap needed — the DriveApp scope
// is already granted by the Round 72 photo upload feature).
function _estPdfFolder() {
  const props = PropertiesService.getScriptProperties();
  const existingId = props.getProperty('EST_PDFS_FOLDER_ID');
  if (existingId) {
    try { return DriveApp.getFolderById(existingId); } catch (_e) { /* recreate below */ }
  }
  const name = 'Leak Guard Estimation PDFs';
  const it = DriveApp.getFoldersByName(name);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(name);
  props.setProperty('EST_PDFS_FOLDER_ID', folder.getId());
  return folder;
}

// body: {action, secret, seNo, phone, pdfBase64, filename?}
function handleUploadEstimationPdf(body) {
  if (!body.seNo || !body.phone || !body.pdfBase64) {
    return jsonResponse({status: 'error', message: 'seNo, phone and pdfBase64 required'});
  }
  try {
    const folder = _estPdfFolder();
    const filename = String(body.filename || (body.seNo + '.pdf')).replace(/[^A-Za-z0-9 ._-]+/g, '-').slice(0, 80);
    const blob = Utilities.newBlob(Utilities.base64Decode(String(body.pdfBase64)), 'application/pdf', filename);
    const file = folder.createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (_e) {}
    const url = file.getUrl();

    // Estimations row ('PDF URL') — best-effort
    try {
      const est = SpreadsheetApp.openById(LIVE_SHEET_ID).getSheetByName(ESTIMATIONS_SHEET);
      if (est) {
        const h = getHeaders(est);
        const rowNum = _findEstimationRowBySe(est, h, body.seNo);
        if (rowNum && h.colByName['PDF URL']) est.getRange(rowNum, h.colByName['PDF URL']).setValue(url);
      }
    } catch (_e) { /* best-effort */ }

    // Lead row ('Quotation PDF URL') — best-effort
    try {
      const sheet = getSheet();
      const leadRow = findRowByPhone(sheet, body.phone);
      if (leadRow) setCellByHeader(sheet, leadRow, 'Quotation PDF URL', url);
    } catch (_e) { /* best-effort */ }

    return jsonResponse({status: 'ok', url: url, id: file.getId()});
  } catch (e) {
    return jsonResponse({status: 'error', message: String(e && e.message || e)});
  }
}

// ================================================================
// One-time bootstraps (run from the Apps Script editor)
// ================================================================

// Creates the Estimations tab with the frozen header row.
// Idempotent: re-running appends any headers that were added in later
// rounds (e.g. 'PDF URL', Round 83.3) without touching existing data.
function bootstrapEstimationsSheet() {
  const ss = SpreadsheetApp.openById(LIVE_SHEET_ID);
  let sheet = ss.getSheetByName(ESTIMATIONS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(ESTIMATIONS_SHEET);
    sheet.getRange(1, 1, 1, ESTIMATIONS_HEADERS.length).setValues([ESTIMATIONS_HEADERS]);
    sheet.setFrozenRows(1);
    Logger.log('bootstrapEstimationsSheet: created Estimations tab');
    return;
  }
  const h = getHeaders(sheet);
  const missing = ESTIMATIONS_HEADERS.filter(function(name) { return !h.colByName[name]; });
  if (!missing.length) {
    Logger.log('bootstrapEstimationsSheet: Estimations tab up to date, nothing to do');
    return;
  }
  const lastCol = sheet.getLastColumn();
  sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
  Logger.log('bootstrapEstimationsSheet: appended headers ' + missing.join(', '));
}

// Adds 'AutoCount QT No' + 'Total Sqft' to the lead sheet if missing.
// Idempotent: only appends headers that don't exist yet.
function bootstrapAutocountLeadColumns() {
  const sheet = getSheet();
  const h = getHeaders(sheet);
  const wanted = ['AutoCount QT No', 'Total Sqft'].filter(function(name) {
    return !h.colByName[name];
  });
  if (!wanted.length) {
    Logger.log('bootstrapAutocountLeadColumns: all columns already present, nothing to do');
    return;
  }
  const lastCol = sheet.getLastColumn();
  sheet.getRange(1, lastCol + 1, 1, wanted.length).setValues([wanted]);
  Logger.log('bootstrapAutocountLeadColumns: added ' + wanted.join(', '));
}
