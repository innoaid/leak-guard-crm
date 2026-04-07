// ================================================================
// LEAK GUARD CRM — CHATSHERO → GOOGLE SHEETS SYNC
// ================================================================
//
// HOW TO FIND A GOOGLE SHEET ID:
//   Open the sheet in your browser. The URL looks like:
//   https://docs.google.com/spreadsheets/d/XXXXXXXXXX/edit
//   The XXXXXXXXXX part between /d/ and /edit is the Sheet ID.
//
// SETUP INSTRUCTIONS:
//   1. Open your destination Google Sheet
//   2. Go to Extensions → Apps Script
//   3. Paste this entire file into Code.gs
//   4. Run setupLeakGuardSheet() — creates all tabs
//   5. Run setupTrigger() — creates 5-minute auto-sync
//   6. Run testSyncOnce() — test with sample data
//   7. Set CONFIG.ACTIVE = true when ready for live sync
//
// ================================================================

// ── CONFIGURATION ────────────────────────────────────────────────
// Change these values to match your setup.

const CONFIG = {
  ACTIVE:           false,                                          // Set to true to enable auto-sync
  START_DATE:       '2026-01-01',                                   // Only sync leads created on or after this date
  LOOKBACK_DAYS:    365,                                            // Re-check existing leads within this window
  SOURCE_SHEET_ID:  '1XXHMfXfj2UMgB39RSufqDtGHNPzuse18NadCLz-FNbA', // ChatHero Google Sheet ID
  SOURCE_TAB:       'Sheet1',                                       // Tab name in ChatHero sheet
  DEST_TAB:         'Leak Guard Leads',                             // Destination tab for qualified leads
  STAFF_TAB:        'Staff List',                                   // Staff directory tab
  SCHEDULE_TAB:     'Schedule',                                     // Job scheduling tab
  LOG_TAB:          'Sync Log',                                     // Sync history log tab
};

// ── CHATHERO SOURCE COLUMNS (0-based index) ──────────────────────
// These map to the columns in the ChatHero Google Sheet.
// A=0, B=1, C=2, etc.

const CH = {
  CONV_ID:     0,   // A — Conversation ID (unique key)
  CREATE_DATE: 1,   // B — When the conversation was created
  RECENT_TIME: 2,   // C — Most recent message time
  APPT_DATE:   3,   // D — Appointment date (primary)
  APPT_TIME:   4,   // E — Appointment time
  NAME:        5,   // F — Customer name
  PHONE:       6,   // G — Phone number
  STATE:       7,   // H — State/location
  PROBLEMS:    8,   // I — Problem description
  SLAB_SIZE:   9,   // J — Slab size in sqft
  ADDRESS:     10,  // K — Full address
  QUOT_SV:     11,  // L — Quotation or site visit info
  CHAT_URL:    12,  // M — Link to ChatHero conversation
  QUOT_NO:     13,  // N — Quotation number
  DATE_SENT:   14,  // O — Date quotation was sent
  FOLLOWUP:    15,  // P — Follow-up date
  CH_STATUS:   16,  // Q — ChatHero status
  APPT_DATE2:  17,  // R — Appointment date (secondary/reschedule)
  COMPLETION:  18,  // S — Completion status
};

// ── DESTINATION SHEET HEADERS ────────────────────────────────────
// These are the column headers for the Leak Guard Leads tab.
// Order matters — each index maps to a column.

const HEADERS = [
  'Timestamp',        // A  — When the lead was first synced
  'Phone',            // B  — Customer phone number
  'Name',             // C  — Customer name
  'Problem Type',     // D  — Waterproofing problem description
  'Location',         // E  — State
  'Full Address',     // F  — Complete address
  'Slab Size (sqft)', // G  — Area in square feet
  'Slot Chosen',      // H  — Appointment date + time
  'Status',           // I  — CRM status (dropdown)
  'Assigned To',      // J  — Staff assignment (dropdown)
  'Quotation (RM)',   // K  — Quotation amount
  'Job Outcome',      // L  — Win/loss tracking (dropdown)
  'Notes',            // M  — Free-text notes
  'CH Conv ID',       // N  — ChatHero conversation ID (unique key)
  'CH Status',        // O  — Status from ChatHero
  'CH Chat URL',      // P  — Link to ChatHero conversation
  'Source',           // Q  — Lead source (ChatHero, Manual, etc.)
  'Last Synced',      // R  — When this row was last updated by sync
  'Date Lead In',        // S  — Filled on first sync (copy Timestamp)
  'Date Appt Confirmed', // T  — Filled when status = "Site Visit Confirmed"
  'Date QT Issued',      // U  — Filled when status = "Quotation Sent"
  'Date Confirmed',      // V  — Filled when status = "Pending I.Date"
  'Date Installed',      // W  — Filled when status = "Job Complete"
  'Status Changed At',   // X  — Filled every time status changes
  'Changed By',          // Y  — Filled by Make.com with staff name
];


// ================================================================
// FUNCTION 1 — SETUP: Create all tabs, headers, dropdowns
// ================================================================

function setupLeakGuardSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Rename the spreadsheet
  ss.rename('Leak Guard CRM');

  // ── Leak Guard Leads tab ──────────────────────────────────────
  let dest = ss.getSheetByName(CONFIG.DEST_TAB);
  if (!dest) dest = ss.insertSheet(CONFIG.DEST_TAB);
  dest.clear();
  dest.clearFormats();

  // Headers — dark purple row
  dest.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setBackground('#2D2A6E')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  dest.setFrozenRows(1);
  dest.setFrozenColumns(2);
  dest.setTabColor('#2D2A6E');

  // Status dropdown — col I (column 9)
  const statusList = [
    'New Lead', 'Pending Site Visit', 'Site Visit Confirmed',
    'Pending QT', 'Quotation Sent', 'Follow Up',
    'Pending I.Date', 'I.Date Confirmed',
    'Job In Progress', 'Job Complete', 'Receipt Sent',
    'Cold Lead', 'Lost', 'Rejected', 'Out of Area', 'Human Handoff'
  ];
  dest.getRange(2, 9, 500, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(statusList, true).build()
  );

  // Assigned To dropdown — col J (column 10)
  dest.getRange(2, 10, 500, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Ken', 'Admin', 'Team KL', 'Team JB', 'Unassigned'], true).build()
  );

  // Job Outcome dropdown — col L (column 12)
  dest.getRange(2, 12, 500, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList([
        'Pending', 'Won', 'Lost - Price',
        'Lost - No Response', 'Lost - Not Ready',
        'Lost - Out of Area', 'Referred Out'
      ], true).build()
  );

  // Column widths (18 original + 7 timeline columns)
  var widths = [160, 140, 130, 160, 120, 240, 120, 160, 150, 130, 120, 130, 200, 180, 130, 200, 100, 160, 130, 130, 130, 130, 130, 140, 120];
  widths.forEach(function(w, i) { dest.setColumnWidth(i + 1, w); });

  // ── Staff List tab ────────────────────────────────────────────
  let staff = ss.getSheetByName(CONFIG.STAFF_TAB);
  if (!staff) staff = ss.insertSheet(CONFIG.STAFF_TAB);
  staff.clear();
  staff.getRange(1, 1, 1, 3).setValues([['Phone', 'Name', 'Role']])
    .setBackground('#1D9E75')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  staff.setTabColor('#1D9E75');

  // ── Schedule tab ──────────────────────────────────────────────
  let sched = ss.getSheetByName(CONFIG.SCHEDULE_TAB);
  if (!sched) sched = ss.insertSheet(CONFIG.SCHEDULE_TAB);
  sched.clear();
  sched.getRange(1, 1, 1, 4).setValues([['Date', 'Time Slot', 'Status', 'Assigned Job']])
    .setBackground('#185FA5')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  sched.setTabColor('#185FA5');

  // ── Summary tab ───────────────────────────────────────────────
  let sum = ss.getSheetByName('Summary');
  if (!sum) sum = ss.insertSheet('Summary');
  sum.clear();
  sum.getRange('A1:B1').setValues([['Metric', 'Count']])
    .setBackground('#1D9E75')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');

  var sumData = [
    ['Total Leads',        "=COUNTA('Leak Guard Leads'!B2:B)"],
    ['New Lead',           "=COUNTIF('Leak Guard Leads'!I:I,\"New Lead\")"],
    ['Pending Site Visit', "=COUNTIF('Leak Guard Leads'!I:I,\"Pending Site Visit\")"],
    ['Confirmed',          "=COUNTIF('Leak Guard Leads'!I:I,\"Confirmed\")"],
    ['Quotation Sent',     "=COUNTIF('Leak Guard Leads'!I:I,\"Quotation Sent\")"],
    ['Job Complete',       "=COUNTIF('Leak Guard Leads'!I:I,\"Job Complete\")"],
    ['Lost',               "=COUNTIF('Leak Guard Leads'!I:I,\"Lost\")"],
    ['', ''],
    ['Total Revenue',      "=SUMIF('Leak Guard Leads'!I:I,\"Job Complete\",'Leak Guard Leads'!K:K)"],
    ['Conversion Rate',    "=IFERROR(COUNTIF('Leak Guard Leads'!I:I,\"Job Complete\")/COUNTA('Leak Guard Leads'!B2:B),0)"],
  ];
  sum.getRange(2, 1, sumData.length, 2).setValues(sumData);
  sum.getRange('B11').setNumberFormat('0.0%');
  sum.setTabColor('#1D9E75');

  // ── Job History tab ────────────────────────────────────────────
  let jobHist = ss.getSheetByName('Job History');
  if (!jobHist) jobHist = ss.insertSheet('Job History');
  jobHist.clear();
  jobHist.getRange(1, 1, 1, 7).setValues([['Timestamp', 'Phone', 'Customer Name', 'Old Status', 'New Status', 'Changed By', 'Notes']])
    .setBackground('#444441')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  jobHist.setTabColor('#444441');
  jobHist.setFrozenRows(1);

  // ── Sync Log tab ──────────────────────────────────────────────
  ensureLogTab(ss);

  // ── Delete default Sheet1 if it exists and is empty ───────────
  var sheet1 = ss.getSheetByName('Sheet1');
  if (sheet1 && ss.getSheets().length > 1) {
    var data1 = sheet1.getDataRange().getValues();
    if (data1.length <= 1 && (!data1[0] || !data1[0][0])) {
      ss.deleteSheet(sheet1);
    }
  }

  var setupMsg = 'Setup complete!\n\nTabs created:\n✓ Leak Guard Leads (with dropdowns + timeline columns)\n✓ Staff List\n✓ Schedule\n✓ Summary (with live formulas)\n✓ Job History\n✓ Sync Log\n\nNext steps:\n1. Run setupTrigger() to create auto-sync\n2. Run testSyncOnce() to test\n3. Set CONFIG.ACTIVE = true for live sync';
  try { SpreadsheetApp.getUi().alert(setupMsg); } catch(e) { Logger.log(setupMsg); }
}


// ================================================================
// FUNCTION 2 — SYNC: Pull qualified leads from ChatHero
// ================================================================
// Qualification criteria: Phone + Full Address + Problems must ALL be non-empty.
// New leads get Status=New Lead, Assigned To=Unassigned, Job Outcome=Pending.
// Existing leads: only ChatHero fields are updated. User columns (Status,
// Assigned To, Quotation RM, Job Outcome, Notes) are NEVER overwritten.

function hashData(rows) {
  return rows.map(function(r){ return r.join('|'); }).join('\n').length + '_' + rows.length;
}

function resetHash() {
  PropertiesService.getScriptProperties().deleteProperty('LAST_DATA_HASH');
  Logger.log('LAST_DATA_HASH cleared. Next sync will do a full scan.');
}

// ================================================================
// doPost — Web App endpoint for status updates from job_board.html
// ================================================================
// Deploy as Web App: Execute as Me, Anyone can access.
// After adding, redeploy: Deploy → Manage → Edit → New version → Deploy.

function doPost(e) {
  try {
    Logger.log('RAW PAYLOAD: ' + e.postData.contents);
    var body = JSON.parse(e.postData.contents);
    Logger.log('PARSED KEYS: ' + Object.keys(body).join(', '));
    if (body.messages || body.statuses || body.contacts) {
      return doPostWABot(e);
    }
    if (body.action) {
      return doPostJobBoard(e);
    }
    return ContentService
      .createTextOutput('ok')
      .setMimeType(ContentService.MimeType.TEXT);
  } catch(err) {
    Logger.log('doPost router error: ' + err.toString());
    return ContentService
      .createTextOutput('ok')
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

function doPostJobBoard(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    if (data.action !== 'updateStatus') throw new Error('Unknown action');

    var ss = SpreadsheetApp.openById('17lxgFT5bW-5mcnM-ks2hid1ZI0Icp6ieK7huD3ffWBE');
    var sheet = ss.getSheetByName('Leak Guard Leads');
    var rows = sheet.getDataRange().getValues();

    var rowIdx = -1;
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][1]).trim() === String(data.phone).trim()) {
        rowIdx = i + 1;
        break;
      }
    }
    if (rowIdx === -1) throw new Error('Lead not found: ' + data.phone);

    sheet.getRange(rowIdx, 9).setValue(data.status);
    sheet.getRange(rowIdx, 24).setValue(new Date(data.statusChangedAt));
    sheet.getRange(rowIdx, 25).setValue(data.changedBy || 'Web');
    if (data.dateCol && data.dateVal) {
      sheet.getRange(rowIdx, data.dateCol + 1).setValue(new Date(data.dateVal));
    }

    return ContentService.createTextOutput(JSON.stringify({status:'ok'}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({status:'error',message:err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function syncQualifiedLeads() {
  // Guard: don't run if sync is paused
  if (!CONFIG.ACTIVE) {
    Logger.log('Sync paused. Set CONFIG.ACTIVE = true to start.');
    return;
  }

 try {

  // Calculate effective date window
  var startDate = new Date(CONFIG.START_DATE);
  var lookbackDate = new Date();
  lookbackDate.setDate(lookbackDate.getDate() - CONFIG.LOOKBACK_DAYS);
  var effectiveStart = startDate > lookbackDate ? startDate : lookbackDate;

  // Open ChatHero source sheet
  var srcSS;
  try {
    srcSS = SpreadsheetApp.openById(CONFIG.SOURCE_SHEET_ID);
  } catch (e) {
    Logger.log('ERROR: Cannot open ChatHero sheet. Check SOURCE_SHEET_ID. ' + e);
    return;
  }

  var srcSheet = srcSS.getSheetByName(CONFIG.SOURCE_TAB);
  if (!srcSheet) {
    Logger.log('ERROR: Tab "' + CONFIG.SOURCE_TAB + '" not found in ChatHero sheet.');
    return;
  }

  // Open destination sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dest = ss.getSheetByName(CONFIG.DEST_TAB);
  if (!dest) {
    Logger.log('ERROR: "' + CONFIG.DEST_TAB + '" tab not found. Run setupLeakGuardSheet() first.');
    return;
  }

  ensureLogTab(ss);

  // Read source data + check for changes via hash
  var srcData = srcSheet.getDataRange().getValues();
  var props = PropertiesService.getScriptProperties();
  var currentHash = hashData(srcData);
  var lastHash = props.getProperty('LAST_DATA_HASH') || '';
  if (currentHash === lastHash) {
    Logger.log('No changes detected in ChatHero sheet, skipping sync.');
    return;
  }

  var destData = dest.getDataRange().getValues();
  var convIdCol = HEADERS.indexOf('CH Conv ID'); // column N = index 13

  // Build lookup: Conv ID → destination row number (1-based)
  var existingMap = {};
  for (var d = 1; d < destData.length; d++) {
    var id = destData[d][convIdCol];
    if (id) existingMap[String(id)] = d + 1;
  }

  // Build map of existing phone numbers → row number (1-based) for upsert
  var existingPhones = {};
  for (var p = 1; p < destData.length; p++) {
    var ph = String(destData[p][1] || '').trim();
    if (ph) existingPhones[ph] = p + 1;
  }

  var newCount = 0, updateCount = 0, skipCount = 0;

  // Process each ChatHero row
  for (var i = 1; i < srcData.length; i++) {
    var row = srcData[i];
    var convId     = String(row[CH.CONV_ID] || '').trim();
    var phone      = String(row[CH.PHONE] || '').trim();
    var address    = String(row[CH.ADDRESS] || '').trim();
    var problem    = String(row[CH.PROBLEMS] || '').trim();
    var createDate = row[CH.CREATE_DATE];

    // Skip rows with no conversation ID
    if (!convId) { skipCount++; continue; }

    // Skip rows older than the effective start date
    if (createDate && new Date(createDate) < effectiveStart) { skipCount++; continue; }

    // Qualification check: all 3 fields must be filled
    var qualified = phone && address && problem;

    // Skip unqualified rows that aren't already in the destination
    if (!qualified && !existingMap[convId]) { skipCount++; continue; }

    // Extract remaining fields from ChatHero
    var name     = row[CH.NAME] || '';
    var state    = row[CH.STATE] || '';
    var slabSize = row[CH.SLAB_SIZE] || '';
    var apptDate = row[CH.APPT_DATE] || row[CH.APPT_DATE2] || '';
    var apptTime = row[CH.APPT_TIME] || '';
    var slot     = apptDate ? (apptDate + ' ' + apptTime).trim() : '';
    var chatUrl  = row[CH.CHAT_URL] || '';
    var chStatus = row[CH.CH_STATUS] || '';
    var quotNo   = row[CH.QUOT_NO] || '';
    var now      = new Date();

    if (!existingMap[convId]) {
      // ── NEW LEAD ──────────────────────────────────────────────
      if (!qualified) { skipCount++; continue; }

      // Upsert: if phone already exists, update ChatHero fields on existing row
      if (phone && existingPhones[phone]) {
        var phoneRow = existingPhones[phone];
        // Update ChatHero-sourced cols only — never touch Status, Assigned To, Quotation, Job Outcome, Notes
        dest.getRange(phoneRow, 2, 1, 7).setValues([[phone, name, problem, state, address, slabSize, slot]]);
        dest.getRange(phoneRow, 14, 1, 5).setValues([[convId, chStatus, chatUrl, 'ChatHero', now]]);
        // Fill name only if blank
        if (!dest.getRange(phoneRow, 3).getValue() && name) dest.getRange(phoneRow, 3).setValue(name);
        updateCount++;
        continue;
      }

      var newRow = new Array(HEADERS.length).fill('');
      newRow[0]  = now;                                  // Timestamp
      newRow[1]  = phone;                                // Phone
      newRow[2]  = name;                                 // Name
      newRow[3]  = problem;                              // Problem Type
      newRow[4]  = state;                                // Location
      newRow[5]  = address;                              // Full Address
      newRow[6]  = slabSize;                             // Slab Size
      newRow[7]  = slot;                                 // Slot Chosen
      newRow[8]  = 'New Lead';                           // Status
      newRow[9]  = 'Unassigned';                         // Assigned To
      newRow[10] = '';                                   // Quotation (RM)
      newRow[11] = 'Pending';                            // Job Outcome
      newRow[12] = quotNo ? 'QT: ' + quotNo : '';       // Notes
      newRow[13] = convId;                               // CH Conv ID
      newRow[14] = chStatus;                             // CH Status
      newRow[15] = chatUrl;                              // CH Chat URL
      newRow[16] = 'ChatHero';                           // Source
      newRow[17] = now;                                  // Last Synced
      newRow[18] = now;                                  // Date Lead In
      newRow[19] = '';                                   // Date Appt Confirmed
      newRow[20] = '';                                   // Date QT Issued
      newRow[21] = '';                                   // Date Confirmed
      newRow[22] = '';                                   // Date Installed
      newRow[23] = now;                                  // Status Changed At
      newRow[24] = '';                                   // Changed By

      dest.appendRow(newRow);
      // Track new row for same-batch phone dedup
      if (phone) existingPhones[phone] = dest.getLastRow();
      newCount++;

    } else {
      // ── UPDATE EXISTING LEAD ──────────────────────────────────
      var rowNum = existingMap[convId];

      // Skip if outside lookback window
      if (createDate && new Date(createDate) < lookbackDate) { skipCount++; continue; }

      // Only update ChatHero-sourced columns — NEVER touch user columns
      // (Status, Assigned To, Quotation RM, Job Outcome, Notes are user-managed)
      var updates = [
        [4, problem],    // Problem Type (col D, 1-based = 5)
        [5, state],      // Location (col E, 1-based = 6)
        [6, address],    // Full Address (col F, 1-based = 7)
        [7, slabSize],   // Slab Size (col G, 1-based = 8)
        [8, slot],       // Slot Chosen (col H, 1-based = 9) — NOTE: only if from ChatHero
        [14, chStatus],  // CH Status (col O, 1-based = 15)
        [15, chatUrl],   // CH Chat URL (col P, 1-based = 16)
        [17, now],       // Last Synced (col R, 1-based = 18)
      ];
      updates.forEach(function(pair) {
        var col = pair[0], val = pair[1];
        if (val !== '' && val !== null && val !== undefined) {
          dest.getRange(rowNum, col + 1).setValue(val);
        }
      });

      // Fill name/phone only if currently blank in destination
      if (!dest.getRange(rowNum, 2).getValue() && phone) {
        dest.getRange(rowNum, 2).setValue(phone);
      }
      if (!dest.getRange(rowNum, 3).getValue() && name) {
        dest.getRange(rowNum, 3).setValue(name);
      }

      updateCount++;
    }

    // Rate limiting: pause every 50 rows to avoid quota issues
    if (i % 50 === 0) Utilities.sleep(100);
  }

  // Save hash after successful sync
  props.setProperty('LAST_DATA_HASH', currentHash);

  // Log this sync run
  logRun(ss, newCount, updateCount, skipCount);
  Logger.log('Sync complete — New: ' + newCount + ' | Updated: ' + updateCount + ' | Skipped: ' + skipCount);

 } catch(e) {
    if (e.message && e.message.indexOf('Service Spreadsheets') !== -1) {
      Logger.log('Sheets temporarily unavailable, retrying next run.');
      return;
    }
    throw e;
 }
}


// ================================================================
// FUNCTION 3 — TRIGGER: Set up automatic 5-minute sync
// ================================================================

function setupTrigger() {
  // Remove any existing syncQualifiedLeads triggers
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'syncQualifiedLeads') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Create new 5-minute trigger
  ScriptApp.newTrigger('syncQualifiedLeads')
    .timeBased()
    .everyMinutes(5)
    .create();

  var trigMsg = 'Trigger created!\n\nsyncQualifiedLeads will run every 5 minutes.\n\nRemember: set CONFIG.ACTIVE = true when ready for live sync.';
  try { SpreadsheetApp.getUi().alert(trigMsg); } catch(e) { Logger.log(trigMsg); }
}


// ================================================================
// FUNCTION 4 — TEST: Run one sync cycle and show results
// ================================================================

function testSyncOnce() {
  // Temporarily force ACTIVE=true for this test run
  var orig = CONFIG.ACTIVE;
  CONFIG.ACTIVE = true;

  syncQualifiedLeads();

  CONFIG.ACTIVE = orig;

  // Read the last log entry and display results
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = ss.getSheetByName(CONFIG.LOG_TAB);
  var last = log ? log.getRange(log.getLastRow(), 1, 1, 5).getValues()[0] : [];

  var testMsg = 'Test sync complete!\n\nNew rows: ' + (last[1]||0) + '\nUpdated: ' + (last[2]||0) + '\nSkipped: ' + (last[3]||0) + '\nTotal in sheet: ' + (last[4]||0);
  try { SpreadsheetApp.getUi().alert(testMsg); } catch(e) { Logger.log(testMsg); }
}


// ================================================================
// FUNCTION 5 — HELPER: Ensure Sync Log tab exists
// ================================================================

function ensureLogTab(ss) {
  var log = ss.getSheetByName(CONFIG.LOG_TAB);
  if (!log) {
    log = ss.insertSheet(CONFIG.LOG_TAB);
    log.setTabColor('#888780');
    log.getRange(1, 1, 1, 5)
      .setValues([['Timestamp', 'New', 'Updated', 'Skipped', 'Total']])
      .setBackground('#444441')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
  }
  return log;
}


// ================================================================
// FUNCTION 6 — HELPER: Log each sync run
// ================================================================
// Keeps a rolling log of sync results. Trims to 500 rows max.

function logRun(ss, newCount, updateCount, skipCount) {
  var log = ss.getSheetByName(CONFIG.LOG_TAB);
  if (!log) return;

  // Count total leads in destination
  var dest = ss.getSheetByName(CONFIG.DEST_TAB);
  var total = dest ? Math.max(0, dest.getLastRow() - 1) : 0;

  // Append log entry
  log.appendRow([new Date(), newCount, updateCount, skipCount, total]);

  // Trim old entries if log exceeds 500 data rows
  var rows = log.getLastRow();
  if (rows > 501) {
    log.deleteRows(2, rows - 501);
  }
}


// ================================================================
// ONE-TIME UTILITY — Remove duplicate phone entries
// ================================================================
// Run once from Apps Script editor, then delete this function.
// Keeps the first occurrence of each phone number, deletes later duplicates.

function removeDuplicatePhones() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Leak Guard Leads');
  var data = sheet.getDataRange().getValues();
  var seen = {};
  var rowsToDelete = [];

  for (var i = 1; i < data.length; i++) {
    var phone = String(data[i][1] || '').trim();
    if (!phone) continue;
    if (seen[phone]) {
      rowsToDelete.push(i + 1);
    } else {
      seen[phone] = true;
    }
  }

  rowsToDelete.reverse();
  rowsToDelete.forEach(function(rowNum) {
    sheet.deleteRow(rowNum);
    Logger.log('Deleted duplicate row: ' + rowNum);
  });

  Logger.log('Done. Removed ' + rowsToDelete.length + ' duplicate rows.');
  try {
    SpreadsheetApp.getUi().alert(
      'Done! Removed ' + rowsToDelete.length + ' duplicate phone entries. ' +
      'Kept first occurrence of each phone number.'
    );
  } catch(e) { Logger.log('Removed ' + rowsToDelete.length + ' duplicates'); }
}


// ================================================================
// ONE-TIME UTILITY — Rename statuses + update dropdown
// ================================================================
// Run once from Apps Script editor, then delete this function.

function renameStatuses() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Leak Guard Leads');
  var data = sheet.getDataRange().getValues();

  var renames = {
    'Pending Installation Date': 'Pending I.Date',
    'Downpayment Received': 'I.Date Confirmed'
  };

  for (var i = 1; i < data.length; i++) {
    var status = data[i][8];
    if (renames[status]) {
      sheet.getRange(i+1, 9).setValue(renames[status]);
      Logger.log('Row '+(i+1)+': '+status+' → '+renames[status]);
    }
  }

  var range = sheet.getRange(2, 9, sheet.getMaxRows()-1, 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList([
      'New Lead','Pending Site Visit','Site Visit Confirmed',
      'Pending QT','Quotation Sent','Follow Up',
      'Pending I.Date','I.Date Confirmed',
      'Job In Progress','Job Complete','Receipt Sent',
      'Cold Lead','Lost','Rejected','Out of Area','Human Handoff'
    ], true)
    .build();
  range.setDataValidation(rule);

  Logger.log('Done — statuses renamed and dropdown updated.');
  try { SpreadsheetApp.getUi().alert('Done! Statuses renamed successfully.'); }
  catch(e) { Logger.log('Statuses renamed successfully.'); }
}
