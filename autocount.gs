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
    // Round 83.5 — collision guard: an existing row for this SE No must belong
    // to the SAME lead. A different phone means the SE counter issued a
    // duplicate (the June 2026 incident) — fail loudly instead of silently
    // overwriting another lead's estimation.
    if (rowNum && !_samePhone(sheet.getRange(rowNum, h.colByName['Phone']).getValue(), body.phone)) {
      return jsonResponse({status: 'error', message: 'SE collision: ' + body.seNo + ' already belongs to another lead — SE counter may be broken, check the Counters sheet'});
    }
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

    // Round 83.5 — when the caller supplied both seNo and phone (Trigger A
    // does), refuse to sync a row that belongs to a different lead. Stops a
    // duplicate SE from reusing another lead's quotation.
    if (body.seNo && body.phone && !_samePhone(phone, body.phone)) {
      return jsonResponse({status: 'error', message: 'SE collision: ' + seNo + ' on file belongs to another lead — not syncing. Check the Counters sheet.'});
    }

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
// Round 93 — manually link an EXISTING AutoCount QT to a card.
// ================================================================
// For job cards added manually (already past Quotation Sent) with no QT
// data and no PDF going into the group. Reads the QT from AutoCount by
// docNo (lg-quotation-read) and records value + sqft + chip. Does NOT
// change the phase, rename the group, or post anything to WhatsApp.
// body: {phone, qtNo, secret}
function handleLinkQt(body) {
  const phone = String(body.phone || '').trim();
  let docNo = String(body.qtNo || '').trim().toUpperCase().replace(/\s+/g, '');
  if (!phone) return jsonResponse({status: 'error', message: 'phone required'});
  if (!docNo) return jsonResponse({status: 'error', message: 'QT number required'});
  // Accept "0626-040" or "QT-0626-040".
  if (/^\d{4}-\d{3}$/.test(docNo)) docNo = 'QT-' + docNo;
  if (!/^QT-\d{4}-\d{3}$/.test(docNo)) {
    return jsonResponse({status: 'error', message: 'invalid QT number — expected QT-MMYY-NNN'});
  }

  // Read totals from AutoCount via the n8n webhook (creds live there).
  let acData;
  try {
    const resp = UrlFetchApp.fetch(N8N_QT_READ_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({docNo: docNo}),
      muteHttpExceptions: true,
    });
    if (resp.getResponseCode() !== 200) return jsonResponse({status: 'error', message: 'AutoCount read HTTP ' + resp.getResponseCode()});
    acData = JSON.parse(resp.getContentText());
  } catch (e) {
    return jsonResponse({status: 'error', message: 'AutoCount read failed: ' + String(e && e.message || e)});
  }
  if (!acData || !acData.success) {
    return jsonResponse({status: 'error', message: docNo + ' not found in AutoCount'});
  }

  const sheet = getSheet();
  const rowNum = findRowByPhone(sheet, phone);
  if (!rowNum) return jsonResponse({status: 'error', message: 'lead not found'});

  const total = Math.round(Number(acData.total) || 0);
  const sqft  = Math.round(Number(acData.totalSqft) || 0);
  let tag = '';
  try { setCellByHeader(sheet, rowNum, 'AutoCount QT No', docNo); } catch (_e) {}
  if (total) { try { setCellByHeader(sheet, rowNum, 'Quotation (RM)', total); } catch (_e) {} }
  if (sqft)  { try { setCellByHeader(sheet, rowNum, 'Total Sqft', sqft); } catch (_e) {} }
  if (total) {
    try {
      const h = getHeaders(sheet);
      const tagsCol = h.colByName['Tags'];
      if (tagsCol) {
        tag = 'RM' + total + ' · ' + sqft + 'sqft';
        const cur = String(sheet.getRange(rowNum, tagsCol).getValue() || '').trim();
        const list = cur.split(',').map(function(t){ return t.trim(); })
          .filter(function(t){ return t && !/^RM\d.*sqft$/i.test(t); });
        list.push(tag);
        sheet.getRange(rowNum, tagsCol).setValue(list.join(','));
      }
    } catch (_e) {}
  }
  return jsonResponse({status: 'ok', docNo: docNo, total: total, totalSqft: sqft, tag: tag});
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

// Last-8-digit phone match (same idiom as findRowByPhone).
function _samePhone(a, b) {
  const pa = String(a || '').trim();
  const pb = String(b || '').trim();
  if (!pa || !pb) return true; // nothing to compare — don't block
  if (pa === pb) return true;
  const la = pa.length >= 8 ? pa.slice(-8) : pa;
  const lb = pb.length >= 8 ? pb.slice(-8) : pb;
  return la === lb;
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
// Round 83.5 — ONE-TIME repair for the June 2026 SE-collision incident.
// ================================================================
// The SE counter issued SE-0626-001 to every estimation (Sheets stored
// MMYY '0626' as the number 626, so the rollover check reset the count
// on every call). All leads shared one Estimations row, and the
// idempotency guard reused Annie's test QT-0626-012 for Mr Ooi, Mr Oi,
// Wan and Sudip: groups renamed with the wrong code, no real AutoCount
// quotations created.
//
// Run ONCE from the editor AFTER deploying the Round 83.5 code fix.
// It: (1) clears the wrong 'AutoCount QT No' from every real lead
// stamped QT-0626-012 (Annie 60179934386, the test lead that actually
// owns it, is skipped); (2) strips the bogus 'QT-0626-012 ' prefix
// from their Group Name (AE) and pushes the rename to the live WA
// group; (3) un-syncs the shared Estimations row so the surviving
// estimation data (Sudip's) can re-sync into a real quotation.
// Idempotent — re-running is a no-op. Delete after a clean run.
function repairSeCollisionJune2026() {
  const WRONG_QT = 'QT-0626-012';
  const OWNER_PHONE_LAST8 = '79934386'; // Annie — the test lead that legitimately created QT-0626-012

  // (1) + (2) — lead rows
  const sheet = getSheet();
  const h = getHeaders(sheet);
  const qtCol = h.colByName['AutoCount QT No'];
  if (!qtCol) { Logger.log('repair: AutoCount QT No column missing — run bootstrapAutocountLeadColumns()'); return; }
  const data = sheet.getDataRange().getValues();
  let fixed = 0;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][qtCol - 1] || '').trim() !== WRONG_QT) continue;
    const phone = String(data[i][h.colByName['Phone'] - 1] || '').trim();
    if (phone.slice(-8) === OWNER_PHONE_LAST8) { Logger.log('repair: skipping owner lead ' + phone); continue; }
    const rowNum = i + 1;
    const name = String(data[i][h.colByName['Name'] - 1] || '').trim();

    // clear the wrong QT No
    sheet.getRange(rowNum, qtCol).setValue('');

    // strip the bogus prefix from the group name + push to the live WA group
    const gnameCol = h.colByName['Group Name (AE)'];
    const gidCol   = h.colByName['Group ID (AB)'];
    const curName = gnameCol ? String(data[i][gnameCol - 1] || '').trim() : '';
    const newName = curName.replace(new RegExp('^' + WRONG_QT + '\\s+'), '').trim();
    if (gnameCol && newName && newName !== curName) {
      sheet.getRange(rowNum, gnameCol).setValue(newName);
      const groupId = gidCol ? String(data[i][gidCol - 1] || '').trim() : '';
      if (groupId) {
        try {
          UrlFetchApp.fetch(N8N_RENAME_GROUP_URL, {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify({secret: SHARED_SECRET, groupId: groupId, newGroupName: newName}),
            muteHttpExceptions: true,
          });
        } catch (_e) { /* best-effort */ }
      }
      Logger.log('repair: ' + name + ' (' + phone + ') — QT cleared, group renamed to "' + newName + '"');
    } else {
      Logger.log('repair: ' + name + ' (' + phone + ') — QT cleared (group name already clean)');
    }
    fixed++;
  }

  // (3) — un-sync the shared Estimations row so its surviving data can re-sync
  const est = SpreadsheetApp.openById(LIVE_SHEET_ID).getSheetByName(ESTIMATIONS_SHEET);
  if (est) {
    const eh = getHeaders(est);
    const r = _findEstimationRowBySe(est, eh, 'SE-0626-001');
    if (r && String(est.getRange(r, eh.colByName['AutoCount QT No']).getValue() || '').trim() === WRONG_QT) {
      ['AutoCount QT No', 'AutoCount Debtor Code', 'Synced At', 'Sync Status'].forEach(function(c) {
        if (eh.colByName[c]) est.getRange(r, eh.colByName[c]).setValue('');
      });
      Logger.log('repair: Estimations row SE-0626-001 un-synced (data belongs to ' + est.getRange(r, eh.colByName['Name']).getValue() + ')');
    } else {
      Logger.log('repair: Estimations row already clean');
    }
  }

  Logger.log('repairSeCollisionJune2026: done — ' + fixed + ' lead(s) repaired');
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
