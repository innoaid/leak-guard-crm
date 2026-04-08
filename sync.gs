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
  ACTIVE:           true,                                           // Set to true to enable auto-sync
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
  'Cal Event ID',        // Z  — Google Calendar event ID for sync
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

function addTagsColumn() {
  var ss = SpreadsheetApp.openById(
    '17lxgFT5bW-5mcnM-ks2hid1ZI0Icp6ieK7huD3ffWBE');
  var sheet = ss.getSheetByName('Leak Guard Leads');
  sheet.getRange(1, 27).setValue('Tags');
  Logger.log('Tags column added to col AA');
}

function resetSyncTime() {
  PropertiesService.getScriptProperties()
    .deleteProperty('LAST_SYNC_TIME');
  Logger.log('Sync time reset. Next run will process all rows.');
}

// ================================================================
// doPost — Web App endpoint for status updates from job_board.html
// ================================================================
// Deploy as Web App: Execute as Me, Anyone can access.
// After adding, redeploy: Deploy → Manage → Edit → New version → Deploy.

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.action) {
      return doPostJobBoard(e);
    }
    if (body.object === 'whatsapp_business_account' ||
        body.entry ||
        body.messages ||
        body.statuses) {
      return doPostWABot(e);
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

    if (data.action === 'updateTags') {
      var ss2 = SpreadsheetApp.openById(
        '17lxgFT5bW-5mcnM-ks2hid1ZI0Icp6ieK7huD3ffWBE');
      var sheet2 = ss2.getSheetByName('Leak Guard Leads');
      var rows2 = sheet2.getDataRange().getValues();
      var rowIdx2 = -1;
      for (var i2 = 1; i2 < rows2.length; i2++) {
        if (String(rows2[i2][1]).trim() ===
            String(data.phone).trim()) {
          rowIdx2 = i2 + 1;
          break;
        }
      }
      if (rowIdx2 === -1) {
        return ContentService
          .createTextOutput(JSON.stringify({
            status:'error',message:'Lead not found'}))
          .setMimeType(ContentService.MimeType.JSON);
      }
      sheet2.getRange(rowIdx2, 27).setValue(data.tags||'');
      return ContentService
        .createTextOutput(JSON.stringify({status:'ok'}))
        .setMimeType(ContentService.MimeType.JSON);
    }

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

// ================================================================
// Google Calendar integration — LG Appointments
// ================================================================

function createLGCalendarEvent(name, phone, location,
  fullAddress, problemType, slabSize, dateTimeStr) {
  try {
    var calendars = CalendarApp
      .getCalendarsByName('LG Appointments');
    var cal = calendars.length > 0
      ? calendars[0]
      : CalendarApp.createCalendar('LG Appointments', {
          color: CalendarApp.Color.CYAN
        });

    var start = new Date(dateTimeStr);
    var end = new Date(start.getTime() + 60 * 60 * 1000);

    var hour = start.getHours();
    var validSlots = [9, 11, 13, 15];
    if (validSlots.indexOf(hour) === -1) {
      return {
        success: false,
        message: 'Invalid slot. Available: 9am, 11am, 1pm, 3pm'
      };
    }

    // Delete any existing LG Appointments events
    // for this lead using phone number in description
    var searchStart = new Date(start);
    searchStart.setDate(searchStart.getDate() - 60);
    var searchEnd = new Date(start);
    searchEnd.setDate(searchEnd.getDate() + 60);
    var allEvents = cal.getEvents(searchStart, searchEnd);
    var deleted = 0;
    allEvents.forEach(function(ev) {
      var desc = ev.getDescription() || '';
      var evPhone = String(phone).trim();
      if (desc.indexOf('Phone: ' + evPhone) !== -1) {
        Logger.log('Deleting existing event for phone: ' +
          evPhone + ' title: ' + ev.getTitle());
        ev.deleteEvent();
        deleted++;
      }
    });
    Logger.log('Deleted ' + deleted +
      ' existing events for ' + phone);

    // Check conflict — ignore if same phone
    var conflicts = cal.getEvents(start, end);
    var hasConflict = conflicts.some(function(ev){
      var desc = ev.getDescription() || '';
      return desc.indexOf('Phone: ' +
        String(phone).trim()) === -1;
    });
    if (hasConflict) {
      return {
        success: false,
        message: 'Slot already taken by another lead. ' +
          'Please choose a different time.'
      };
    }

    var title = (location || '') + ' - ' + (name || phone);
    var desc =
      'Phone: ' + (phone || '') + '\n' +
      'Address: ' + (fullAddress || location || '') + '\n' +
      'Problem: ' + (problemType || '') + '\n' +
      'Slab Size: ' + (slabSize || '') + '\n';

    var event = cal.createEvent(title, start, end, {
      description: desc
    });

    Logger.log('Calendar event created: ' + title +
      ' at ' + start);
    return {
      success: true,
      eventId: event.getId(),
      message: 'Event created: ' + title +
        ' on ' + start.toLocaleDateString('en-MY',{
          weekday:'short',day:'numeric',
          month:'short',year:'numeric'
        }) + ' ' + start.toLocaleTimeString('en-MY',{
          hour:'2-digit',minute:'2-digit'
        })
    };
  } catch(err) {
    Logger.log('Calendar error: ' + err.toString());
    return { success: false, message: err.toString() };
  }
}

function syncCalendarFromSheet() {
  try {
    var ss = SpreadsheetApp.openById(
      '17lxgFT5bW-5mcnM-ks2hid1ZI0Icp6ieK7huD3ffWBE');
    var sheet = ss.getSheetByName('Leak Guard Leads');
    var data = sheet.getDataRange().getValues();

    var calendars = CalendarApp
      .getCalendarsByName('LG Appointments');
    var cal = calendars.length > 0
      ? calendars[0]
      : CalendarApp.createCalendar('LG Appointments');

    var VALID_SLOTS = [9, 11, 13, 15];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var status      = String(row[8]  || '').trim();
      var name        = String(row[2]  || '').trim();
      var phone       = String(row[1]  || '').trim();
      var location    = String(row[4]  || '').trim();
      var fullAddress = String(row[5]  || '').trim();
      var problemType = String(row[3]  || '').trim();
      var slabSize    = String(row[6]  || '').trim();
      var dateAppt    = row[19]; // col T = dateApptConf
      var eventId     = String(row[25] || '').trim(); // col Z

      // Only process Site Visit Confirmed with a date
      if (status === 'Site Visit Confirmed' && dateAppt) {
        var apptDate = new Date(dateAppt);
        if (isNaN(apptDate)) continue;

        var hr = apptDate.getHours();
        if (VALID_SLOTS.indexOf(hr) === -1) {
          Logger.log('Row '+(i+1)+': invalid slot hour '+hr+
            ' for '+name+' — skipping');
          continue;
        }

        var start = new Date(
          apptDate.getFullYear(),
          apptDate.getMonth(),
          apptDate.getDate(),
          hr, 0, 0
        );
        var end = new Date(start.getTime() + 3600000);
        var title = (location||'') + ' - ' + (name||phone);
        var desc =
          'Phone: ' + phone + '\n' +
          'Address: ' + (fullAddress||location) + '\n' +
          'Problem: ' + problemType + '\n' +
          'Slab Size: ' + slabSize;

        if (eventId) {
          // Event exists — check if time changed
          try {
            var existing = cal.getEventById(eventId);
            if (existing) {
              var existStart = existing.getStartTime();
              if (existStart.getTime() !== start.getTime()) {
                // Time changed — update event
                existing.setTime(start, end);
                existing.setTitle(title);
                existing.setDescription(desc);
                Logger.log('Updated event: ' + title +
                  ' to ' + start);
              } else {
                Logger.log('No change: ' + title);
              }
            } else {
              // Event ID invalid — create new
              eventId = '';
            }
          } catch(e) {
            Logger.log('Event not found, creating new: ' + e);
            eventId = '';
          }
        }

        if (!eventId) {
          // No event yet — create new
          // Check for slot conflict first
          var conflicts = cal.getEvents(start, end);
          var hasConflict = conflicts.some(function(ev) {
            return ev.getTitle() !== title;
          });
          if (hasConflict) {
            Logger.log('Slot conflict for ' + title +
              ' at ' + start + ' — skipping');
            continue;
          }
          var newEvent = cal.createEvent(title, start, end, {
            description: desc
          });
          // Save event ID to col Z (index 25, getRange col 26)
          sheet.getRange(i+1, 26).setValue(newEvent.getId());
          Logger.log('Created event: ' + title + ' at ' + start);
        }

      } else if (status !== 'Site Visit Confirmed' && eventId) {
        // Status changed away — delete calendar event
        try {
          var evToDel = cal.getEventById(eventId);
          if (evToDel) {
            evToDel.deleteEvent();
            Logger.log('Deleted event for: ' + name +
              ' status changed to: ' + status);
          }
        } catch(e) {
          Logger.log('Delete error: ' + e);
        }
        // Clear event ID from sheet
        sheet.getRange(i+1, 26).setValue('');
      }
    }

    Logger.log('syncCalendarFromSheet complete.');

  } catch(err) {
    Logger.log('syncCalendarFromSheet error: ' +
      err.toString());
  }
}

function setupCalendarTrigger() {
  // Remove existing calendar sync triggers first
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'syncCalendarFromSheet') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // Create new 5-min trigger
  ScriptApp.newTrigger('syncCalendarFromSheet')
    .timeBased()
    .everyMinutes(5)
    .create();
  Logger.log('Calendar sync trigger created.');
}

function getAvailableSlots(dateStr) {
  try {
    var calendars = CalendarApp
      .getCalendarsByName('LG Appointments');
    if (!calendars.length) return [9, 11, 13, 15];
    var cal = calendars[0];
    var date = new Date(dateStr);
    var allSlots = [9, 11, 13, 15];
    var available = allSlots.filter(function(hour) {
      var start = new Date(date);
      start.setHours(hour, 0, 0, 0);
      var end = new Date(start.getTime() + 3600000);
      return cal.getEvents(start, end).length === 0;
    });
    return available;
  } catch(err) {
    return [9, 11, 13, 15];
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

  // Read source data + filter for rows updated since last sync (timestamp-based)
  var srcData = srcSheet.getDataRange().getValues();
  var sourceRows = srcData.slice(1);
  var props = PropertiesService.getScriptProperties();
  var lastSyncTime = props.getProperty('LAST_SYNC_TIME');
  var lastSync = lastSyncTime ? new Date(lastSyncTime) : new Date('2026-01-01');
  Logger.log('Last sync time: ' + lastSync);

  // Filter to only rows newer than last sync
  // ChatHero has no reliable last-updated timestamp; use
  // Create Chat Date (index 1) for new lead detection.
  var rowsToProcess = sourceRows.filter(function(row) {
    var rowTimestamp = row[1] ? new Date(row[1]) : null;
    if (!rowTimestamp || isNaN(rowTimestamp)) return false;
    return rowTimestamp > lastSync;
  });

  Logger.log('Total source rows: ' + sourceRows.length +
    ' | Rows to process: ' + rowsToProcess.length);

  // If no new rows, exit early
  if (rowsToProcess.length === 0) {
    Logger.log('No new or updated rows since last sync. Skipping.');
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

  // Process only the filtered rows (newer than last sync)
  for (var i = 0; i < rowsToProcess.length; i++) {
    var row = rowsToProcess[i];

    // ChatHero source columns (0-based):
    // 0=User NS, 1=Create Chat Date, 2=Create Chat Time,
    // 3=Appt Date, 4=Appt Time, 5=Name, 6=Phone, 7=State,
    // 8=Problems, 9=Slab Size, 10=Full Address,
    // 11=Quot/SiteVisit, 12=Chat URL, 13=Quot No,
    // 14=Date Sent, 15=Follow Up, 16=Status, 17=Completion
    var phone        = String(row[6] || '').trim();
    var name         = String(row[5] || '').trim();
    var problemType  = String(row[8] || '').trim();
    var location     = String(row[7] || '').trim();
    var fullAddress  = String(row[10] || '').trim();
    var slabSize     = String(row[9] || '').trim();
    var slotDate     = row[3] ? String(row[3]).trim() : '';
    var slotTime     = row[4] ? String(row[4]).trim() : '';
    var slotChosen   = slotDate && slotTime
      ? slotDate + ' ' + slotTime
      : (slotDate || slotTime || '');
    var chConvId     = String(row[0] || '').trim();
    var chStatus     = String(row[16] || '').trim();
    var chChatUrl    = String(row[12] || '').trim();
    var createDate   = row[1] ? new Date(row[1]) : new Date();

    // Skip rows with no conversation ID
    if (!chConvId) { skipCount++; continue; }

    // Skip rows older than the effective start date
    if (createDate && createDate < effectiveStart) { skipCount++; continue; }

    // Qualification: phone + problem + (full address OR location)
    var qualified = phone && problemType &&
      (fullAddress || location);

    // Skip unqualified rows that aren't already in the destination
    if (!qualified && !existingMap[chConvId]) { skipCount++; continue; }

    var now = new Date();

    if (!existingMap[chConvId]) {
      // ── NEW LEAD ──────────────────────────────────────────────
      if (!qualified) { skipCount++; continue; }

      // Upsert: if phone already exists, update ChatHero fields on existing row
      if (phone && existingPhones[phone]) {
        var phoneRow = existingPhones[phone];
        var existingRow = destData[phoneRow - 1] || [];
        // Skip if existing row already has matching ChatHero data
        if (String(existingRow[3] || '') === String(problemType) &&
            String(existingRow[5] || '') === String(fullAddress) &&
            String(existingRow[13] || '') === String(chConvId)) {
          skipCount++;
          continue;
        }
        // Update ChatHero-sourced cols only — never touch Status, Assigned To, Quotation, Job Outcome, Notes
        dest.getRange(phoneRow, 2, 1, 7).setValues([[phone, name, problemType, location, fullAddress, slabSize, slotChosen]]);
        dest.getRange(phoneRow, 14, 1, 5).setValues([[chConvId, chStatus, chChatUrl, 'ChatHero', now]]);
        // Fill name only if blank
        if (!dest.getRange(phoneRow, 3).getValue() && name) dest.getRange(phoneRow, 3).setValue(name);
        updateCount++;
        continue;
      }

      var newRow = new Array(HEADERS.length).fill('');
      newRow[0]  = createDate;       // Timestamp
      newRow[1]  = phone;            // Phone
      newRow[2]  = name;             // Name
      newRow[3]  = problemType;      // Problem Type
      newRow[4]  = location;         // Location
      newRow[5]  = fullAddress;      // Full Address
      newRow[6]  = slabSize;         // Slab Size
      newRow[7]  = slotChosen;       // Slot Chosen
      newRow[8]  = 'New Lead';       // Status - always New Lead
      newRow[9]  = '';               // Assigned To - blank
      newRow[10] = '';               // Quotation - blank
      newRow[11] = '';               // Job Outcome - blank
      newRow[12] = '';               // Notes - blank
      newRow[13] = chConvId;         // CH Conv ID
      newRow[14] = chStatus;         // CH Status
      newRow[15] = chChatUrl;        // CH Chat URL
      newRow[16] = 'ChatHero';       // Source
      newRow[17] = new Date();       // Last Synced
      newRow[18] = createDate;       // Date Lead In

      dest.appendRow(newRow);
      // Track new row for same-batch phone dedup
      if (phone) existingPhones[phone] = dest.getLastRow();
      newCount++;

    } else {
      // ── UPDATE EXISTING LEAD ──────────────────────────────────
      var rowNum = existingMap[chConvId];

      // Skip if outside lookback window
      if (createDate && createDate < lookbackDate) { skipCount++; continue; }

      // Only update ChatHero-sourced columns — NEVER touch user columns
      // (Status=8, Assigned To=9, Quotation=10, Job Outcome=11,
      //  Notes=12, or any date columns 18-24)
      // First value is 0-based destination index; getRange adds +1 for 1-based.
      var updates = [
        [3,  problemType],   // Problem Type (col D)
        [4,  location],      // Location (col E)
        [5,  fullAddress],   // Full Address (col F)
        [6,  slabSize],      // Slab Size (col G)
        [7,  slotChosen],    // Slot Chosen (col H)
        [14, chStatus],      // CH Status (col O)
        [15, chChatUrl],     // CH Chat URL (col P)
        [17, new Date()],    // Last Synced (col R)
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

  // Save sync time after successful sync
  props.setProperty('LAST_SYNC_TIME', new Date().toISOString());

  // Log this sync run
  logRun(ss, newCount, updateCount, skipCount);
  Logger.log('Sync complete: ' + newCount + ' inserted, ' + updateCount + ' updated, ' + skipCount + ' skipped.');

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
