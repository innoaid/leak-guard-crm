var BOT_CONFIG = {
  SHEET_ID: '17lxgFT5bW-5mcnM-ks2hid1ZI0Icp6ieK7huD3ffWBE',
  SHEET_NAME: 'Leak Guard Leads',
  WA_NUMBER: '60138938657',
  API_URL: 'https://waba-v2.360dialog.io/messages'
};

function getProps() {
  return PropertiesService.getScriptProperties();
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
      changedBy:       row[24] ? String(row[24]) : ''
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

  return {
    total: data.length,
    statusCounts: statusCounts,
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
    data.slice(0,200).map(function(r){
      return r.phone+'|'+(r.name||'')+'|'+r.status+'|'+r.location+'|'+(r.assignedTo||'');
    }).join('\n') + '\n\n' +
    'AVAILABLE SLOTS: 9am-10am, 11am-12pm, 1pm-2pm, 3pm-4pm\n' +
    'When setting Site Visit Confirmed:\n' +
    '- Always use one of the 4 exact slot times above\n' +
    '- morning = suggest 9am or 11am\n' +
    '- afternoon = suggest 1pm or 3pm\n' +
    '- Format date as ISO: 2026-04-08T09:00:00\n\n' +
    'ACTIONS - add at end of reply, max ONE per reply:\n' +
    'STATUS UPDATE: ACTION:{"type":"updateStatus",' +
    '"phone":"60XX","status":"Status",' +
    '"date":"2026-04-08T09:00:00"}\n' +
    'Date needed for: Site Visit Confirmed, I.Date Confirmed.\n' +
    'Auto-date: Quotation Sent, Job Complete.\n' +
    'CHECK SLOTS: ACTION:{"type":"checkSlots",' +
    '"date":"2026-04-08"}\n' +
    'Use when user asks what slots are free on a date.\n' +
    'REMINDER: ACTION:{"type":"sendReminders",' +
    '"phones":["60XX"]}\n' +
    'NEVER include more than one ACTION per reply.\n' +
    'NEVER perform any action without explicit instruction.\n';

  var response = callClaude(systemPrompt, msgText);

  var actionMatch = response.match(/ACTION:(\{.*\})/);
  if (actionMatch) {
    try {
      var action = JSON.parse(actionMatch[1]);
      executeAction(action);
      response = response.replace(/ACTION:\{.*\}/, '').trim();
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
    max_tokens: 1024,
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

function executeAction(action) {
  if (action.type === 'updateStatus') {
    updateLeadStatus(
      action.phone, action.status, action.date||null);
    if (action.status === 'Site Visit Confirmed' &&
        action.date) {
      Logger.log('Calendar attempt - phone: ' +
        action.phone + ' date: ' + action.date);

      var calDate = action.date;
      // If not ISO format, try to parse naturally
      if (calDate && calDate.indexOf('T') === -1) {
        var d = new Date(calDate);
        if (!isNaN(d)) calDate = d.toISOString();
      }
      Logger.log('Normalized date: ' + calDate);

      var data = getSheetData();
      var lead = data.find(function(r){
        return String(r.phone).trim() ===
          String(action.phone).trim();
      });
      if (lead) {
        var result = createLGCalendarEvent(
          lead.name,
          lead.phone,
          lead.location,
          lead.fullAddress,
          lead.problemType,
          lead.slabSize,
          calDate
        );
        Logger.log('Calendar result: ' +
          JSON.stringify(result));
        return result.message || '';
      }
    }
  } else if (action.type === 'checkSlots') {
    var slots = getAvailableSlots(action.date);
    var slotNames = {
      9:'9am-10am',
      11:'11am-12pm',
      13:'1pm-2pm',
      15:'3pm-4pm'
    };
    var available = slots.map(function(h){
      return slotNames[h];
    }).join(', ');
    return 'Available slots on ' + action.date +
      ': ' + (available || 'No slots available');
  } else if (action.type === 'sendReminders') {
    action.phones.forEach(function(phone){
      sendReminderToClient(phone);
    });
  }
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
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).trim() === String(phone).trim()) {
      sheet.getRange(i+1, 9).setValue(status);
      sheet.getRange(i+1, 24).setValue(new Date());
      sheet.getRange(i+1, 25).setValue('WA Bot');
      if (dateVal && DATE_COL_MAP[status] !== undefined) {
        sheet.getRange(i+1, DATE_COL_MAP[status] + 1)
          .setValue(new Date(dateVal));
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
