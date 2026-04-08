var BOT_CONFIG = {
  SHEET_ID: '17lxgFT5bW-5mcnM-ks2hid1ZI0Icp6ieK7huD3ffWBE',
  SHEET_NAME: 'Leak Guard Leads',
  WA_NUMBER: '60138938657',
  API_URL: 'https://waba-v2.360dialog.io/messages'
};

var DATE_REQUIRED = ['Site Visit Confirmed', 'I.Date Confirmed'];

function getProps() {
  return PropertiesService.getScriptProperties();
}

function getPendingConfirmation(senderPhone) {
  var raw = getProps().getProperty('PENDING_' + senderPhone);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch(e) { return null; }
}

function setPendingConfirmation(senderPhone, payload) {
  getProps().setProperty('PENDING_' + senderPhone,
    JSON.stringify(payload));
}

function clearPendingConfirmation(senderPhone) {
  getProps().deleteProperty('PENDING_' + senderPhone);
}

function findLeadByIdentifier(identifier, data) {
  var id = String(identifier).trim().toLowerCase();
  // Try exact name match first
  var byName = data.filter(function(r) {
    return r.name &&
      r.name.toLowerCase().indexOf(id) !== -1;
  });
  if (byName.length === 1) return byName[0];
  if (byName.length > 1) return byName[0];

  // Try last 5-7 digits of phone
  var digits = id.replace(/\D/g, '');
  if (digits.length >= 5) {
    var byPhone = data.filter(function(r) {
      return r.phone &&
        String(r.phone).slice(-digits.length) === digits;
    });
    if (byPhone.length === 1) return byPhone[0];
    if (byPhone.length > 1) return byPhone[0];
  }
  return null;
}

function getApiKey() {
  return getProps().getProperty('DIALOG360_API_KEY');
}

function getAnthropicKey() {
  return getProps().getProperty('ANTHROPIC_API_KEY');
}

function getAllowedNumbers() {
  var nums = getProps().getProperty('ALLOWED_NUMBERS') || '';
  return nums.split(',').map(function(n){ return n.trim(); });
}

function setupBotKeys() {
  var props = getProps();
  props.setProperty('DIALOG360_API_KEY', 'PASTE_360DIALOG_KEY_HERE');
  props.setProperty('ANTHROPIC_API_KEY', 'PASTE_ANTHROPIC_KEY_HERE');
  props.setProperty('ALLOWED_NUMBERS', '60146938657');
  Logger.log('Keys saved successfully.');
}

function doPostWABot(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.entry && body.entry[0] &&
        body.entry[0].changes && body.entry[0].changes[0]) {
      var value = body.entry[0].changes[0].value;
      if (value.messages) body.messages = value.messages;
      if (value.statuses) body.statuses = value.statuses;
      if (value.contacts) body.contacts = value.contacts;
    }
    if (body.statuses) {
      return ContentService.createTextOutput('ok')
        .setMimeType(ContentService.MimeType.TEXT);
    }
    var messages = body.messages;
    if (!messages || !messages.length) {
      return ContentService.createTextOutput('ok')
        .setMimeType(ContentService.MimeType.TEXT);
    }
    var msg = messages[0];
    var senderPhone = String(msg.from).trim();
    var msgId = msg.id || '';
    var msgTimestamp = parseInt(msg.timestamp || '0') * 1000;
    var props = getProps();

    // Block if message is older than 5 minutes
    var now = Date.now();
    if (msgTimestamp && (now - msgTimestamp) > 300000) {
      Logger.log('Stale message ignored: ' + msgId +
        ' age: ' + Math.round((now-msgTimestamp)/1000) + 's');
      return ContentService.createTextOutput('ok')
        .setMimeType(ContentService.MimeType.TEXT);
    }

    // Block duplicate message ID
    var lastMsgId = props.getProperty('LAST_MSG_' + senderPhone) || '';
    if (msgId && msgId === lastMsgId) {
      Logger.log('Duplicate message ignored: ' + msgId);
      return ContentService.createTextOutput('ok')
        .setMimeType(ContentService.MimeType.TEXT);
    }
    if (msgId) props.setProperty('LAST_MSG_' + senderPhone, msgId);
    var msgType = msg.type;
    var msgText = '';
    if (msgType === 'text') {
      msgText = msg.text.body;
    } else {
      sendWAMessage(senderPhone,
        'Sorry, I can only process text messages.');
      return ContentService.createTextOutput('ok')
        .setMimeType(ContentService.MimeType.TEXT);
    }
    var allowed = getAllowedNumbers();
    if (allowed.indexOf(senderPhone) === -1) {
      Logger.log('Unauthorized sender: ' + senderPhone);
      return ContentService.createTextOutput('ok')
        .setMimeType(ContentService.MimeType.TEXT);
    }
    var reply = processAdminMessage(senderPhone, msgText);
    sendWAMessage(senderPhone, reply);
    return ContentService.createTextOutput('ok')
      .setMimeType(ContentService.MimeType.TEXT);
  } catch(err) {
    Logger.log('doPost error: ' + err.toString());
    return ContentService.createTextOutput('ok')
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

function getSheetData() {
  var ss = SpreadsheetApp.openById(BOT_CONFIG.SHEET_ID);
  var sheet = ss.getSheetByName(BOT_CONFIG.SHEET_NAME);
  var rows = sheet.getDataRange().getValues();
  return rows.slice(1).map(function(row) {
    return {
      timestamp:       row[0]  ? String(row[0])  : '',
      phone:           row[1]  ? String(row[1])  : '',
      name:            row[2]  ? String(row[2])  : '',
      problemType:     row[3]  ? String(row[3])  : '',
      location:        row[4]  ? String(row[4])  : '',
      fullAddress:     row[5]  ? String(row[5])  : '',
      slabSize:        row[6]  ? String(row[6])  : '',
      slotChosen:      row[7]  ? String(row[7])  : '',
      status:          row[8]  ? String(row[8])  : '',
      assignedTo:      row[9]  ? String(row[9])  : '',
      quotation:       row[10] ? String(row[10]) : '',
      jobOutcome:      row[11] ? String(row[11]) : '',
      notes:           row[12] ? String(row[12]) : '',
      dateLeadIn:      row[18] ? String(row[18]) : '',
      dateApptConf:    row[19] ? String(row[19]) : '',
      dateQTIssued:    row[20] ? String(row[20]) : '',
      dateConfirmed:   row[21] ? String(row[21]) : '',
      dateInstalled:   row[22] ? String(row[22]) : '',
      statusChangedAt: row[23] ? String(row[23]) : '',
      changedBy:       row[24] ? String(row[24]) : '',
      tags:            row[26] ? String(row[26]).trim() : ''
    };
  }).filter(function(r){ return r.phone || r.name; });
}

function getDataSummary(data) {
  var today = new Date();
  var todayStr = today.toDateString();
  var tomorrow = new Date(today);
  tomorrow.setDate(tomorrow.getDate() + 1);
  var tomorrowStr = tomorrow.toDateString();

  function isDate(str, dateStr) {
    if (!str) return false;
    var d = new Date(str);
    return !isNaN(d) && d.toDateString() === dateStr;
  }

  var statusCounts = {};
  data.forEach(function(r) {
    if (r.status) {
      statusCounts[r.status] = (statusCounts[r.status]||0) + 1;
    }
  });

  var taggedLeads = data.filter(function(r){
    return r.tags && r.tags.trim() !== '';
  });

  return {
    total: data.length,
    statusCounts: statusCounts,
    taggedLeads: taggedLeads,
    todayAppts: data.filter(function(r){
      return isDate(r.dateApptConf, todayStr);
    }),
    tomorrowAppts: data.filter(function(r){
      return isDate(r.dateApptConf, tomorrowStr);
    }),
    todayLeads: data.filter(function(r){
      return isDate(r.dateLeadIn, todayStr);
    }),
    unassigned: data.filter(function(r){
      return !r.assignedTo &&
        ['Lost','Rejected','Out of Area','Cold Lead']
          .indexOf(r.status) === -1;
    }),
    todayDate: today.toLocaleDateString('en-MY',{
      weekday:'long',day:'numeric',
      month:'long',year:'numeric'
    }),
    tomorrowDate: tomorrow.toLocaleDateString('en-MY',{
      weekday:'long',day:'numeric',
      month:'long',year:'numeric'
    })
  };
}

function processAdminMessage(senderPhone, msgText) {
  var upperMsg = msgText.trim().toUpperCase();

  // Handle YES/NO confirmation for pending bulk moves
  var pending = getPendingConfirmation(senderPhone);
  if (pending && Date.now() <= pending.expires) {
    if (upperMsg === 'YES' || upperMsg === 'YA' ||
        upperMsg === 'OK' || upperMsg === 'CONFIRM') {
      return executeAction({type:'confirmPending'}, senderPhone);
    }
    if (upperMsg === 'NO' || upperMsg === 'CANCEL' ||
        upperMsg === 'BATAL') {
      return executeAction({type:'cancelPending'}, senderPhone);
    }
  }

  var data = getSheetData();
  var summary = getDataSummary(data);

  var systemPrompt =
    'You are a CRM assistant for Leak Guard waterproofing company Malaysia.\n' +
    'STRICT RULES - follow exactly:\n' +
    '1. Reply with ONE short message only. Never send multiple messages.\n' +
    '2. Never ask follow-up questions. Never say "How can I help you today?"\n' +
    '3. Never say "Is there anything else?" or similar.\n' +
    '4. After completing an action, just confirm it briefly and stop.\n' +
    '5. Only perform ONE action per reply. Never chain actions.\n' +
    '6. Reply in same language as user (English or BM).\n' +
    '7. Keep all replies under 5 lines.\n' +
    '8. Never repeat yourself.\n\n' +
    'Today is ' + summary.todayDate + '. Tomorrow is ' + summary.tomorrowDate + '.\n\n' +
    'CRM SUMMARY:\n' +
    'Total: ' + summary.total + ' leads\n' +
    'New Lead: ' + (summary.statusCounts['New Lead']||0) + '\n' +
    'Pending Site Visit: ' + (summary.statusCounts['Pending Site Visit']||0) + '\n' +
    'Site Visit Confirmed: ' + (summary.statusCounts['Site Visit Confirmed']||0) + '\n' +
    'Pending QT: ' + (summary.statusCounts['Pending QT']||0) + '\n' +
    'Quotation Sent: ' + (summary.statusCounts['Quotation Sent']||0) + '\n' +
    'Follow Up: ' + (summary.statusCounts['Follow Up']||0) + '\n' +
    'Pending I.Date: ' + (summary.statusCounts['Pending I.Date']||0) + '\n' +
    'I.Date Confirmed: ' + (summary.statusCounts['I.Date Confirmed']||0) + '\n' +
    'Job In Progress: ' + (summary.statusCounts['Job In Progress']||0) + '\n' +
    'Job Complete: ' + (summary.statusCounts['Job Complete']||0) + '\n' +
    'Receipt Sent: ' + (summary.statusCounts['Receipt Sent']||0) + '\n' +
    'Cold Lead: ' + (summary.statusCounts['Cold Lead']||0) + '\n' +
    'Lost: ' + (summary.statusCounts['Lost']||0) + '\n\n' +
    'TODAY APPTS (' + summary.todayAppts.length + '):\n' +
    (summary.todayAppts.length ? summary.todayAppts.map(function(r){
      return r.name + ' | ' + r.location + ' | ' + r.dateApptConf + ' | ' + r.phone;
    }).join('\n') : 'None') + '\n\n' +
    'TOMORROW APPTS (' + summary.tomorrowAppts.length + '):\n' +
    (summary.tomorrowAppts.length ? summary.tomorrowAppts.map(function(r){
      return r.name + ' | ' + r.location + ' | ' + r.dateApptConf + ' | ' + r.phone;
    }).join('\n') : 'None') + '\n\n' +
    'TODAY NEW LEADS (' + summary.todayLeads.length + '):\n' +
    (summary.todayLeads.length ? summary.todayLeads.map(function(r){
      return r.name + ' | ' + r.location + ' | ' + r.problemType;
    }).join('\n') : 'None') + '\n\n' +
    'LEAD LIST:\n' +
    data.map(function(r){
      return r.phone+'|'+(r.name||'')+'|'+r.status+'|'+r.location+'|'+(r.assignedTo||'');
    }).join('\n') + '\n\n' +
    'AVAILABLE SLOTS: 9am-10am, 11am-12pm, 1pm-2pm, 3pm-4pm\n' +
    'When setting Site Visit Confirmed:\n' +
    '- Always use one of the 4 exact slot times above\n' +
    '- morning = suggest 9am or 11am\n' +
    '- afternoon = suggest 1pm or 3pm\n' +
    '- ALWAYS format date with time as: 2026-04-09T09:00:00\n' +
    '- NEVER send date without time e.g. 2026-04-09 is WRONG\n' +
    '- Valid hours only: 09, 11, 13, 15\n' +
    '- Example: 9am on 9 Apr = 2026-04-09T09:00:00\n' +
    '- Example: 11am on 9 Apr = 2026-04-09T11:00:00\n' +
    '- Example: 1pm on 9 Apr = 2026-04-09T13:00:00\n' +
    '- Example: 3pm on 9 Apr = 2026-04-09T15:00:00\n\n' +
    'ACTIONS - add at end of reply, max ONE per reply:\n' +
    'SINGLE MOVE (1 lead, executes immediately):\n' +
    'ACTION:{"type":"updateStatus","phone":"60XX",' +
    '"status":"Status","date":"2026-04-08T09:00:00"}\n\n' +
    'BULK MOVE (2-10 leads, asks for confirmation):\n' +
    'ACTION:{"type":"bulkUpdateStatus","moves":[\n' +
    '  {"identifier":"name or last 5-7 digits","status":"Status"},\n' +
    '  {"identifier":"name or last 5-7 digits","status":"Status"}\n' +
    ']}\n' +
    'Rules:\n' +
    '- Use bulkUpdateStatus when user mentions 2+ leads\n' +
    '- identifier can be name or last 5-7 phone digits\n' +
    '- Max 10 leads, reject with reason if more\n' +
    '- Date-required statuses in bulk: skip that lead,\n' +
    '  move others, report skip with reason\n' +
    '- Date-required: Site Visit Confirmed, I.Date Confirmed\n' +
    '- Auto-date: Quotation Sent, Job Complete\n\n' +
    'CONFIRMATION: When user replies YES/YA/OK —\n' +
    'do NOT use any ACTION, just reply normally.\n' +
    'System handles YES/NO automatically.\n\n' +
    'CHECK SLOTS: ACTION:{"type":"checkSlots",' +
    '"date":"2026-04-08"}\n' +
    'Use when user asks what slots are free on a date.\n' +
    'REMINDER: ACTION:{"type":"sendReminders",' +
    '"phones":["60XX"]}\n' +
    'NEVER include more than one ACTION per reply.\n' +
    'NEVER perform any action without explicit instruction.\n\n' +
    'TAGGED LEADS (' + summary.taggedLeads.length + '):\n' +
    (summary.taggedLeads.length ? summary.taggedLeads.map(function(r){
      return '- ' + (r.name||r.phone) +
        ' | ' + r.status +
        ' | Tags: ' + r.tags +
        ' | ' + r.phone;
    }).join('\n') : 'None') + '\n\n' +
    'TAG MANAGEMENT:\n' +
    'Valid tags: Receipt, Complaint, Question, ' +
    'Price Nego, Special Req\n' +
    'Multiple tags separated by comma.\n' +
    'ACTION to update tags: ACTION:{"type":"updateTags",' +
    '"phone":"60XX","tags":"Receipt,Question"}\n' +
    'To clear all tags send empty string: "tags":""\n';

  var response = callClaude(systemPrompt, msgText);

  var actionMatch = response.match(/ACTION:(\{.*\})/);
  if (actionMatch) {
    try {
      var action = JSON.parse(actionMatch[1]);
      var actionResult = executeAction(action, senderPhone);
      response = response.replace(/ACTION:\{.*\}/, '').trim();
      if (actionResult) {
        response = (response ? response + '\n' : '') + actionResult;
      }
    } catch(err) {
      Logger.log('Action parse error: ' + err);
    }
  }
  return response;
}

function callClaude(systemPrompt, userMessage) {
  var apiKey = getAnthropicKey();
  if (!apiKey) return 'Anthropic API key not configured.';

  var payload = {
    model: 'claude-sonnet-4-20250514',
    max_tokens: 2048,
    system: systemPrompt,
    messages: [
      { role: 'user', content: userMessage }
    ]
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(
    'https://api.anthropic.com/v1/messages', options);
  var json = JSON.parse(response.getContentText());

  if (json.content && json.content[0]) {
    return json.content[0].text;
  }
  return 'Sorry, I could not process your request.';
}

function executeAction(action, senderPhone) {
  var data = getSheetData();

  // ── SINGLE STATUS UPDATE ──────────────────────
  if (action.type === 'updateStatus') {
    updateLeadStatus(
      action.phone, action.status, action.date||null);
    if (action.status === 'Site Visit Confirmed' &&
        action.date) {
      var lead = data.find(function(r){
        return String(r.phone).trim() ===
          String(action.phone).trim();
      });
      if (lead) {
        var result = createLGCalendarEvent(
          lead.name, lead.phone, lead.location,
          lead.fullAddress, lead.problemType,
          lead.slabSize, action.date);
        Logger.log('Calendar: ' + JSON.stringify(result));
      }
    }
    return '';
  }

  // ── BULK STATUS UPDATE ───────────────────────
  if (action.type === 'bulkUpdateStatus') {
    var moves = action.moves || [];

    // Reject if more than 10
    if (moves.length > 10) {
      return 'Too many leads (' + moves.length + '). ' +
        'Max 10 per bulk move. Please split into batches.';
    }

    // Resolve each identifier to a lead
    var resolved = [];
    var notFound = [];
    moves.forEach(function(m) {
      var lead = findLeadByIdentifier(m.identifier, data);
      if (lead) {
        resolved.push({
          lead: lead,
          status: m.status,
          date: m.date || null
        });
      } else {
        notFound.push(m.identifier);
      }
    });

    if (resolved.length === 0) {
      return 'No leads found. Please check the names ' +
        'or phone numbers and try again.';
    }

    // Single move — execute immediately, no confirmation
    if (resolved.length === 1 && notFound.length === 0) {
      var single = resolved[0];
      var dateRequiredSingle = DATE_REQUIRED.indexOf(single.status) !== -1;
      if (dateRequiredSingle && !single.date) {
        return 'Cannot move ' + (single.lead.name||single.lead.phone) +
          ' to ' + single.status +
          ' — this status requires a date. ' +
          'Please specify the date and time.';
      }
      updateLeadStatus(single.lead.phone, single.status, single.date||null);
      return (single.lead.name||single.lead.phone) +
        ' moved to ' + single.status + ' ✅';
    }

    // Multiple moves — store and ask for confirmation
    var confirmLines = [];
    var skipped = [];
    var toConfirm = [];

    resolved.forEach(function(mv) {
      var dateReq = DATE_REQUIRED.indexOf(mv.status) !== -1;
      if (dateReq && !mv.date) {
        skipped.push((mv.lead.name||mv.lead.phone) +
          ' → ' + mv.status +
          ' (requires date — skipped)');
      } else {
        toConfirm.push(mv);
        confirmLines.push(
          (toConfirm.length) + '. ' +
          (mv.lead.name||mv.lead.phone) +
          ' (' + mv.lead.phone + ')' +
          ' → ' + mv.status
        );
      }
    });

    if (toConfirm.length === 0) {
      return 'All leads require dates. ' +
        'Please move them individually:\n' +
        skipped.join('\n');
    }

    // Store pending confirmation
    setPendingConfirmation(senderPhone, {
      moves: toConfirm.map(function(mv) {
        return {
          phone: mv.lead.phone,
          status: mv.status,
          date: mv.date || null
        };
      }),
      expires: Date.now() + 300000 // 5 min
    });

    var reply = 'Confirm ' + toConfirm.length +
      ' moves?\n\n' + confirmLines.join('\n');
    if (skipped.length > 0) {
      reply += '\n\n⚠️ Skipped (' + skipped.length + '):\n' +
        skipped.join('\n');
    }
    if (notFound.length > 0) {
      reply += '\n\n❌ Not found: ' + notFound.join(', ');
    }
    reply += '\n\nReply YES to confirm or NO to cancel.';
    return reply;
  }

  // ── CONFIRM PENDING ───────────────────────────
  if (action.type === 'confirmPending') {
    var pending = getPendingConfirmation(senderPhone);
    if (!pending || Date.now() > pending.expires) {
      clearPendingConfirmation(senderPhone);
      return 'No pending confirmation found or it expired.';
    }
    var done = 0, failed = 0, failMsg = '';
    pending.moves.forEach(function(mv) {
      try {
        updateLeadStatus(mv.phone, mv.status, mv.date||null);
        done++;
      } catch(e) {
        failed++;
        failMsg = e.toString();
      }
    });
    clearPendingConfirmation(senderPhone);
    var res = 'Done ✅ ' + done + ' leads updated.';
    if (failed > 0) res += '\n❌ ' + failed +
      ' failed: ' + failMsg;
    return res;
  }

  // ── CANCEL PENDING ────────────────────────────
  if (action.type === 'cancelPending') {
    clearPendingConfirmation(senderPhone);
    return 'Cancelled. No changes made.';
  }

  // ── SEND REMINDERS ────────────────────────────
  if (action.type === 'sendReminders') {
    action.phones.forEach(function(phone){
      sendReminderToClient(phone);
    });
    return '';
  }

  // ── UPDATE TAGS ───────────────────────────────
  if (action.type === 'updateTags') {
    var ss = SpreadsheetApp.openById(BOT_CONFIG.SHEET_ID);
    var sheet = ss.getSheetByName(BOT_CONFIG.SHEET_NAME);
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][1]).trim() ===
          String(action.phone).trim()) {
        sheet.getRange(i+1, 27).setValue(action.tags||'');
        return 'Tags updated for ' +
          (rows[i][2]||action.phone) + ': ' +
          (action.tags||'cleared');
      }
    }
    return 'Lead not found: ' + action.phone;
  }

  // ── CHECK SLOTS ───────────────────────────────
  if (action.type === 'checkSlots') {
    var slots = getAvailableSlots(action.date);
    var slotNames = {
      9:'9am-10am', 11:'11am-12pm',
      13:'1pm-2pm', 15:'3pm-4pm'
    };
    var available = slots.map(function(h){
      return slotNames[h];
    }).join(', ');
    return 'Available slots on ' + action.date +
      ': ' + (available||'No slots available');
  }

  return '';
}

function updateLeadStatus(phone, status, dateVal) {
  var ss = SpreadsheetApp.openById(BOT_CONFIG.SHEET_ID);
  var sheet = ss.getSheetByName(BOT_CONFIG.SHEET_NAME);
  var rows = sheet.getDataRange().getValues();
  // 0-based column indices; +1 applied at write time for getRange
  var DATE_COL_MAP = {
    'Site Visit Confirmed': 19,  // col T = dateApptConf
    'Quotation Sent':       20,  // col U = dateQTIssued
    'I.Date Confirmed':     21,  // col V = dateConfirmed
    'Job Complete':         22   // col W = dateInstalled
  };
  Logger.log('updateLeadStatus - status: ' + status +
    ' dateVal: ' + dateVal +
    ' dateCol: ' + DATE_COL_MAP[status]);
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).trim() === String(phone).trim()) {
      sheet.getRange(i+1, 9).setValue(status);
      sheet.getRange(i+1, 24).setValue(new Date());
      sheet.getRange(i+1, 25).setValue('WA Bot');
      if (dateVal && DATE_COL_MAP[status] !== undefined) {
        var dv;
        if (dateVal instanceof Date) {
          dv = dateVal;
        } else {
          var raw = String(dateVal);
          // Extract date and hour from ISO string
          var m = raw.match(/(\d{4})-(\d{2})-(\d{2})T(\d{2})/);
          if (m) {
            dv = new Date(
              parseInt(m[1]),
              parseInt(m[2]) - 1,
              parseInt(m[3]),
              parseInt(m[4]),
              0, 0
            );
          } else {
            dv = new Date(raw);
          }
        }
        if (dv && !isNaN(dv)) {
          sheet.getRange(i+1, DATE_COL_MAP[status]+1)
            .setValue(dv);
          Logger.log('Date written: ' + dv +
            ' to col ' + (DATE_COL_MAP[status]+1));
        }
      }
      Logger.log('Updated ' + phone + ' to ' + status);
      return;
    }
  }
  Logger.log('Lead not found: ' + phone);
}

function sendReminderToClient(clientPhone) {
  var data = getSheetData();
  var lead = data.find(function(r){
    return String(r.phone).trim() === String(clientPhone).trim();
  });
  if (!lead) return;
  var msg =
    'Hi ' + (lead.name||'there') + '! 👋\n\n' +
    'This is a reminder from Leak Guard.\n\n' +
    'Our team will be visiting your property ' +
    'tomorrow for a waterproofing site visit.\n\n' +
    'Address: ' + (lead.fullAddress||lead.location) + '\n' +
    'Date: ' + lead.dateApptConf + '\n\n' +
    'Please ensure someone is available at the property. ' +
    'Contact us if you need to reschedule.\n\n' +
    'Thank you! 🙏';
  sendWAMessage(clientPhone, msg);
}

function sendWAMessage(toPhone, message) {
  var apiKey = getApiKey();
  if (!apiKey) {
    Logger.log('360dialog API key not configured');
    return;
  }
  var payload = {
    messaging_product: 'whatsapp',
    recipient_type: 'individual',
    to: toPhone,
    type: 'text',
    text: { body: message }
  };
  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'D360-API-KEY': apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  var response = UrlFetchApp.fetch(BOT_CONFIG.API_URL, options);
  Logger.log('WA send: ' + response.getContentText());
}

function testBot() {
  var reply = processAdminMessage(
    '60146938657',
    'show tomorrow appointments'
  );
  Logger.log(reply);
}

function testSendWA() {
  sendWAMessage('60146938657',
    'Test message from Leak Guard Admin Bot');
}
