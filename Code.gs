// ================================================================
// LEAK GUARD CRM — FULL SETUP + CHATSHERO SYNC SCRIPT
// ================================================================

const CONFIG = {
  ACTIVE: false,
  START_DATE: '2026-04-05',
  LOOKBACK_DAYS: 30,
  SOURCE_SHEET_ID: '1XXHMfXfj2UMgB39RSufqDtGHNPzuse18NadCLz-FNbA',
  SOURCE_TAB: 'Sheet1',
  DEST_TAB: 'Leak Guard Leads',
  STAFF_TAB: 'Staff List',
  SCHEDULE_TAB: 'Schedule',
  LOG_TAB: 'Sync Log',
};

// ChatHero column indices (0-based)
const CH = {
  CONV_ID:0, CREATE_DATE:1, RECENT_TIME:2,
  APPT_DATE:3, APPT_TIME:4, NAME:5,
  PHONE:6, STATE:7, PROBLEMS:8,
  SLAB_SIZE:9, ADDRESS:10, QUOT_SV:11,
  CHAT_URL:12, QUOT_NO:13, DATE_SENT:14,
  FOLLOWUP:15, CH_STATUS:16,
  APPT_DATE2:17, COMPLETION:18
};

// Destination headers
const HEADERS = [
  'Timestamp','Phone','Name','Problem Type',
  'Location','Full Address','Slab Size (sqft)',
  'Slot Chosen','Status','Assigned To',
  'Quotation (RM)','Job Outcome','Notes',
  'CH Conv ID','CH Status','CH Chat URL',
  'Source','Last Synced'
];

function setupLeakGuardSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Main leads tab
  let dest = ss.getSheetByName(CONFIG.DEST_TAB);
  if (!dest) dest = ss.insertSheet(CONFIG.DEST_TAB);
  dest.clear(); dest.clearFormats();
  dest.getRange(1,1,1,HEADERS.length).setValues([HEADERS])
    .setBackground('#2D2A6E').setFontColor('#FFFFFF')
    .setFontWeight('bold');
  dest.setFrozenRows(1);
  dest.setFrozenColumns(2);
  dest.setTabColor('#2D2A6E');

  // Status dropdown col I (index 8)
  const statusList = ['New Lead','Pending Site Visit','Confirmed',
    'Site Visit Done','Quotation Sent','Follow Up','Job Confirmed',
    'Downpayment Received','Job In Progress','Job Complete',
    'Receipt Sent','Cold Lead','Lost','Out of Area','Human Handoff'];
  dest.getRange(2,9,500,1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(statusList,true).build());

  // Assigned To dropdown col J (index 9)
  dest.getRange(2,10,500,1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Ken','Admin','Team KL','Team JB','Unassigned'],true).build());

  // Job Outcome dropdown col L (index 11)
  dest.getRange(2,12,500,1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pending','Won','Lost - Price',
        'Lost - No Response','Lost - Not Ready','Lost - Out of Area','Referred Out'],true).build());

  // Column widths
  [160,140,130,160,120,240,120,160,150,130,120,130,200,180,130,200,100,160]
    .forEach((w,i) => dest.setColumnWidth(i+1,w));

  // Staff List tab
  let staff = ss.getSheetByName(CONFIG.STAFF_TAB);
  if (!staff) staff = ss.insertSheet(CONFIG.STAFF_TAB);
  staff.clear();
  staff.getRange(1,1,1,3).setValues([['Phone','Name','Role']])
    .setBackground('#1D9E75').setFontColor('#FFFFFF').setFontWeight('bold');
  staff.setTabColor('#1D9E75');

  // Schedule tab
  let sched = ss.getSheetByName(CONFIG.SCHEDULE_TAB);
  if (!sched) sched = ss.insertSheet(CONFIG.SCHEDULE_TAB);
  sched.clear();
  sched.getRange(1,1,1,4).setValues([['Date','Time Slot','Status','Assigned Job']])
    .setBackground('#185FA5').setFontColor('#FFFFFF').setFontWeight('bold');
  sched.setTabColor('#185FA5');

  // Summary tab
  let sum = ss.getSheetByName('Summary');
  if (!sum) sum = ss.insertSheet('Summary');
  sum.clear();
  sum.getRange('A1:B1').setValues([['Metric','Count']])
    .setBackground('#1D9E75').setFontColor('#FFFFFF').setFontWeight('bold');
  const sumData = [
    ['Total Leads','=COUNTA(\'Leak Guard Leads\'!B2:B)'],
    ['New Lead','=COUNTIF(\'Leak Guard Leads\'!I:I,"New Lead")'],
    ['Pending Site Visit','=COUNTIF(\'Leak Guard Leads\'!I:I,"Pending Site Visit")'],
    ['Confirmed','=COUNTIF(\'Leak Guard Leads\'!I:I,"Confirmed")'],
    ['Quotation Sent','=COUNTIF(\'Leak Guard Leads\'!I:I,"Quotation Sent")'],
    ['Job Complete','=COUNTIF(\'Leak Guard Leads\'!I:I,"Job Complete")'],
    ['Lost','=COUNTIF(\'Leak Guard Leads\'!I:I,"Lost")'],
    ['',''],
    ['Total Revenue','=SUMIF(\'Leak Guard Leads\'!I:I,"Job Complete",\'Leak Guard Leads\'!K:K)'],
    ['Conversion Rate','=IFERROR(COUNTIF(\'Leak Guard Leads\'!I:I,"Job Complete")/COUNTA(\'Leak Guard Leads\'!B2:B),0)'],
  ];
  sum.getRange(2,1,sumData.length,2).setValues(sumData);
  sum.getRange('B11').setNumberFormat('0.0%');
  sum.setTabColor('#1D9E75');

  // Sync Log tab
  ensureLogTab(ss);

  SpreadsheetApp.getUi().alert('Setup complete!\n\nTabs created:\n✓ Leak Guard Leads\n✓ Staff List\n✓ Schedule\n✓ Summary\n✓ Sync Log\n\nNext: run setupTrigger()');
}

function syncQualifiedLeads() {
  if (!CONFIG.ACTIVE) {
    Logger.log('Sync paused. Set CONFIG.ACTIVE = true to start.');
    return;
  }
  const startDate = new Date(CONFIG.START_DATE);
  const lookbackDate = new Date();
  lookbackDate.setDate(lookbackDate.getDate() - CONFIG.LOOKBACK_DAYS);
  const effectiveStart = startDate > lookbackDate ? startDate : lookbackDate;

  let srcSS;
  try { srcSS = SpreadsheetApp.openById(CONFIG.SOURCE_SHEET_ID); }
  catch(e) { Logger.log('ERROR: Cannot open ChatHero sheet. Check SOURCE_SHEET_ID.'); return; }

  const srcSheet = srcSS.getSheetByName(CONFIG.SOURCE_TAB);
  if (!srcSheet) { Logger.log('ERROR: Tab not found: ' + CONFIG.SOURCE_TAB); return; }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dest = ss.getSheetByName(CONFIG.DEST_TAB);
  if (!dest) { Logger.log('ERROR: Run setupLeakGuardSheet() first.'); return; }

  ensureLogTab(ss);

  const srcData = srcSheet.getDataRange().getValues();
  const destData = dest.getDataRange().getValues();
  const convIdCol = HEADERS.indexOf('CH Conv ID'); // col N = index 13

  // Build lookup map: convId → dest row number (1-based)
  const existingMap = {};
  for (let i = 1; i < destData.length; i++) {
    const id = destData[i][convIdCol];
    if (id) existingMap[String(id)] = i + 1;
  }

  let newCount=0, updateCount=0, skipCount=0;

  for (let i = 1; i < srcData.length; i++) {
    const row = srcData[i];
    const convId  = String(row[CH.CONV_ID] || '').trim();
    const phone   = String(row[CH.PHONE] || '').trim();
    const address = String(row[CH.ADDRESS] || '').trim();
    const problem = String(row[CH.PROBLEMS] || '').trim();
    const createDate = row[CH.CREATE_DATE];

    if (!convId) { skipCount++; continue; }
    if (createDate && new Date(createDate) < effectiveStart) { skipCount++; continue; }

    // Qualification: all 3 must be filled
    const qualified = phone && address && problem;
    if (!qualified && !existingMap[convId]) { skipCount++; continue; }

    const name     = row[CH.NAME] || '';
    const state    = row[CH.STATE] || '';
    const slabSize = row[CH.SLAB_SIZE] || '';
    const apptDate = row[CH.APPT_DATE] || row[CH.APPT_DATE2] || '';
    const apptTime = row[CH.APPT_TIME] || '';
    const slot     = apptDate ? (apptDate + ' ' + apptTime).trim() : '';
    const chatUrl  = row[CH.CHAT_URL] || '';
    const chStatus = row[CH.CH_STATUS] || '';
    const quotNo   = row[CH.QUOT_NO] || '';
    const now      = new Date();

    if (!existingMap[convId]) {
      if (!qualified) { skipCount++; continue; }
      const newRow = new Array(HEADERS.length).fill('');
      newRow[0]  = now;
      newRow[1]  = phone;
      newRow[2]  = name;
      newRow[3]  = problem;
      newRow[4]  = state;
      newRow[5]  = address;
      newRow[6]  = slabSize;
      newRow[7]  = slot;
      newRow[8]  = 'New Lead';
      newRow[9]  = 'Unassigned';
      newRow[10] = '';
      newRow[11] = 'Pending';
      newRow[12] = quotNo ? 'QT: '+quotNo : '';
      newRow[13] = convId;
      newRow[14] = chStatus;
      newRow[15] = chatUrl;
      newRow[16] = 'ChatHero';
      newRow[17] = now;
      dest.appendRow(newRow);
      newCount++;
    } else {
      const rowNum = existingMap[convId];
      if (createDate && new Date(createDate) < lookbackDate) { skipCount++; continue; }
      // Only update ChatHero-sourced columns, never touch user columns
      [[4,problem],[5,state],[6,address],[7,slabSize],
       [8,slot],[14,chStatus],[15,chatUrl],[17,now]]
        .forEach(([col,val]) => {
          if (val !== '' && val !== null && val !== undefined)
            dest.getRange(rowNum, col+1).setValue(val);
        });
      // Fill name/phone only if currently blank
      if (!dest.getRange(rowNum,2).getValue() && phone)
        dest.getRange(rowNum,2).setValue(phone);
      if (!dest.getRange(rowNum,3).getValue() && name)
        dest.getRange(rowNum,3).setValue(name);
      updateCount++;
    }
    if (i % 50 === 0) Utilities.sleep(100);
  }

  logRun(ss, newCount, updateCount, skipCount);
  Logger.log('Done — New:'+newCount+' Updated:'+updateCount+' Skipped:'+skipCount);
}

function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncQualifiedLeads')
      ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncQualifiedLeads').timeBased().everyMinutes(5).create();
  SpreadsheetApp.getUi().alert('Trigger created!\nsyncQualifiedLeads runs every 5 minutes.\n\nRemember to set CONFIG.ACTIVE = true when ready.');
}

function testSyncOnce() {
  const orig = CONFIG.ACTIVE;
  CONFIG.ACTIVE = true;
  syncQualifiedLeads();
  CONFIG.ACTIVE = orig;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(CONFIG.LOG_TAB);
  const last = log ? log.getRange(log.getLastRow(),1,1,5).getValues()[0] : [];
  SpreadsheetApp.getUi().alert(
    'Test sync complete!\n\n' +
    'New rows: ' + (last[1]||0) + '\n' +
    'Updated: ' + (last[2]||0) + '\n' +
    'Skipped: ' + (last[3]||0) + '\n' +
    'Total in sheet: ' + (last[4]||0)
  );
}

function ensureLogTab(ss) {
  let log = ss.getSheetByName(CONFIG.LOG_TAB);
  if (!log) {
    log = ss.insertSheet(CONFIG.LOG_TAB);
    log.setTabColor('#888780');
    log.getRange(1,1,1,5).setValues([['Timestamp','New','Updated','Skipped','Total']])
      .setBackground('#444441').setFontColor('#FFFFFF').setFontWeight('bold');
  }
  return log;
}

function logRun(ss, n, u, s) {
  const log = ss.getSheetByName(CONFIG.LOG_TAB);
  if (!log) return;
  const dest = ss.getSheetByName(CONFIG.DEST_TAB);
  const total = dest ? Math.max(0, dest.getLastRow()-1) : 0;
  log.appendRow([new Date(), n, u, s, total]);
  const rows = log.getLastRow();
  if (rows > 501) log.deleteRows(2, rows-501);
}
