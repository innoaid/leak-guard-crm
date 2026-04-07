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
    var props = getProps();
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
    'You are an admin assistant for Leak Guard, ' +
    'a waterproofing company in Malaysia. ' +
    'You help staff query and update the CRM system via WhatsApp. ' +
    'Always reply in the same language the user writes in ' +
    '(English or Bahasa Malaysia). ' +
    'Keep replies concise and WhatsApp-friendly. ' +
    'Use line breaks for lists. No markdown bold or headers. ' +
    'Use emojis sparingly for clarity.\n\n' +
    'Today is ' + summary.todayDate + '.\n' +
    'Tomorrow is ' + summary.tomorrowDate + '.\n\n' +
    'CURRENT CRM DATA SUMMARY:\n' +
    'Total leads: ' + summary.total + '\n' +
    'New Leads: ' + (summary.statusCounts['New Lead']||0) + '\n' +
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
    'TODAY APPOINTMENTS (' + summary.todayAppts.length + '):\n' +
    (summary.todayAppts.length ? summary.todayAppts.map(function(r){
      return '- ' + (r.name||r.phone) +
        ' | ' + r.location +
        ' | ' + r.dateApptConf +
        ' | ' + (r.assignedTo||'Unassigned') +
        ' | ' + r.phone;
    }).join('\n') : 'None') + '\n\n' +
    'TOMORROW APPOINTMENTS (' + summary.tomorrowAppts.length + '):\n' +
    (summary.tomorrowAppts.length ? summary.tomorrowAppts.map(function(r){
      return '- ' + (r.name||r.phone) +
        ' | ' + r.location +
        ' | ' + r.dateApptConf +
        ' | ' + (r.assignedTo||'Unassigned') +
        ' | ' + r.phone;
    }).join('\n') : 'None') + '\n\n' +
    'TODAY NEW LEADS (' + summary.todayLeads.length + '):\n' +
    (summary.todayLeads.length ? summary.todayLeads.map(function(r){
      return '- ' + (r.name||r.phone) +
        ' | ' + r.location +
        ' | ' + r.problemType;
    }).join('\n') : 'None') + '\n\n' +
    'UNASSIGNED ACTIVE LEADS (' + summary.unassigned.length + '):\n' +
    summary.unassigned.slice(0,10).map(function(r){
      return '- ' + (r.name||r.phone) +
        ' | ' + r.status +
        ' | ' + r.location;
    }).join('\n') + '\n\n' +
    'IMPORTANT RULES:\n' +
    '- Only perform ONE action per message\n' +
    '- After performing an action, STOP and wait for next instruction\n' +
    '- Never ask follow-up questions like "Is there anything else?"\n' +
    '- Never ask "How can I help you today?"\n' +
    '- Keep replies short — one action confirmation or one answer, then stop\n' +
    '- Do not send multiple messages or repeat yourself\n' +
    '- Never chain multiple status updates automatically\n' +
    '- For status updates, always confirm what you did and ask if anything else needed\n' +
    '- Never assume what the next action should be\n' +
    '- If asked to move to next phase, move ONE step only then stop\n\n' +
    'AVAILABLE ACTIONS:\n' +
    '1. Answer questions about leads, appointments, counts\n' +
    '2. To update ONE lead status, add at END of reply:\n' +
    'ACTION:{"type":"updateStatus","phone":"60XXXXXXXXX",' +
    '"status":"New Status","date":"2026-04-08T10:00:00"}\n' +
    'Only include date if status requires it.\n' +
    'Date-required: Site Visit Confirmed, I.Date Confirmed\n' +
    'Auto-date: Quotation Sent, Job Complete\n' +
    '3. To send reminders:\n' +
    'ACTION:{"type":"sendReminders","phones":["60XXX"]}\n' +
    'NEVER include more than ONE ACTION per response.\n' +
    'NEVER perform follow-up actions automatically.\n\n' +
    'FULL LEAD LIST:\n' +
    data.slice(0,200).map(function(r){
      return r.phone + '|' + (r.name||'') +
        '|' + r.status +
        '|' + r.location +
        '|' + (r.assignedTo||'');
    }).join('\n');

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
    updateLeadStatus(action.phone, action.status, action.date||null);
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
  var DATE_COL_MAP = {
    'Site Visit Confirmed': 20,
    'Quotation Sent':       21,
    'I.Date Confirmed':     22,
    'Job Complete':         23
  };
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).trim() === String(phone).trim()) {
      sheet.getRange(i+1, 9).setValue(status);
      sheet.getRange(i+1, 24).setValue(new Date());
      sheet.getRange(i+1, 25).setValue('WA Bot');
      if (dateVal && DATE_COL_MAP[status]) {
        sheet.getRange(i+1, DATE_COL_MAP[status])
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
