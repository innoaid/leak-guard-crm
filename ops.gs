// ================================================================
// Operations Coordination + On-Site QC scheduling  (Phase 1)
// ----------------------------------------------------------------
// Coordinator confirms Installation Date + Expected Days + Subcon team
// for a "Pending I.Date" card (ops_confirm.html). The system
// random-assigns 1 of 2 QC (check-visit) teams that has a vacant slot
// on EVERY working day of the install window, then blocks one vacant
// slot per day on that team's Google Calendar. On-site team can
// extend/shorten (auto-reflow trailing check-slots; an extension that
// overflows a day is soft-allowed + flagged) and mark a check done.
//
// QC team = CHECK crew (1 slot = 1 check visit). Subcon = labour
// (chosen manually by coordinator, recorded on the card, NOT scheduled).
// System schedules only the 2 QC calendars. Sunday is never a work day.
//
// CALENDAR ACCESS goes through n8n (workflow "LG - QC Calendar",
// webhook lg-qc-cal) using the same Google Calendar credential as every
// other LG calendar — so the Apps Script project needs NO calendar
// permission, only the URL-fetch permission it already has. The proxy
// only ever touches the two allow-listed QC calendars.
//
// Router cases in kanban_code.gs doPost():
//   opsConfirmQueue · opsConfirmSave · opsOnsiteJobs ·
//   opsEditCompletion · opsMarkCheck · qcCalCheck
// ================================================================

const QC_TEAMS = [
  { name: 'QC Team A', calId: 'fac7acef517820bbc6d98f92a58dfce3a67f88882b6f80a704f0a6ac5c317a1f@group.calendar.google.com' },
  { name: 'QC Team B', calId: '25df23d5f659f8ba3598a6f94acd118366e77ba203f6681a4a21ac7348451fae@group.calendar.google.com' }
];

// 4 one-hour check slots/day (local MYT; end = start + 1h).
const QC_SLOTS = [
  { label: '9:00–10:00',  h: 9,  m: 0  },
  { label: '10:30–11:30', h: 10, m: 30 },
  { label: '1:00–2:00',   h: 13, m: 0  },
  { label: '3:00–4:00',   h: 15, m: 0  }
];

const SUBCON_TEAMS = ['Atik (A)', 'Atik (B)', 'Atik (C)', 'Mohiuddin', 'Danny'];
const OPS_DEST_STATUSES = ['Pending Downpayment', 'Job In Progress'];

const OPS_COL = {
  instDate:       'Installation Date',
  expComplete:    'Expected Completion Date',
  qcTeam:         'Assigned QC Team',
  subcon:         'Assigned Subcon Team',
  actualComplete: 'Actual Completion Date'
};

// n8n QC-calendar proxy (owns the Google Calendar credential).
const QC_PROXY_URL = 'https://leakguard.app.n8n.cloud/webhook/lg-qc-cal';
const QC_EVENT_MARK = 'LGQC';

// ---------------------------------------------------------------- sheet / date helpers

function _opsEnsureCol(sheet, name) {
  let col = getHeaders(sheet).colByName[name];
  if (!col) { col = sheet.getLastColumn() + 1; sheet.getRange(1, col).setValue(name); }
  return col;
}
function _opsParseDate(str) {
  const m = String(str || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 0, 0, 0, 0);
}
function _opsFmtDate(d) {
  const p = function(n) { return (n < 10 ? '0' : '') + n; };
  return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
}
function _opsAddDays(dateStr, n) { const d = _opsParseDate(dateStr); d.setDate(d.getDate() + n); return _opsFmtDate(d); }
// n working days from startStr (inclusive), skipping Sundays.
function _opsWorkingDays(startStr, n) {
  const out = []; const d = _opsParseDate(startStr);
  if (!d) return out;
  while (out.length < n) { if (d.getDay() !== 0) out.push(_opsFmtDate(d)); d.setDate(d.getDate() + 1); }
  return out;
}
// working days (Sun excluded) from start to end inclusive.
function _opsWorkingDaySpan(startStr, endStr) {
  const a = _opsParseDate(startStr), b = _opsParseDate(endStr);
  if (!a || !b || b < a) return 0;
  let n = 0; const d = new Date(a.getTime());
  while (d <= b) { if (d.getDay() !== 0) n++; d.setDate(d.getDate() + 1); }
  return n;
}
function _qcTeamByName(name) {
  const t = String(name || '').trim();
  for (let i = 0; i < QC_TEAMS.length; i++) if (QC_TEAMS[i].name.toLowerCase() === t.toLowerCase()) return QC_TEAMS[i];
  return null;
}

// ---------------------------------------------------------------- n8n calendar proxy

// TZ-independent datetime strings.
function _pad2(n) { return (n < 10 ? '0' : '') + n; }
function _mytLocal(dateStr, h, m)  { return dateStr + 'T' + _pad2(h) + ':' + _pad2(m) + ':00'; }         // event create (+ timeZone field)
function _mytOffset(dateStr, h, m) { return dateStr + 'T' + _pad2(h) + ':' + _pad2(m) + ':00+08:00'; }   // query bounds / comparisons
function _isoToMytDate(iso) {
  const d = new Date(iso); if (isNaN(d)) return String(iso).slice(0, 10);
  return new Date(d.getTime() + 8 * 3600000).toISOString().slice(0, 10);
}

// POST one op to the proxy; returns parsed Google API JSON ({} on parse fail).
function _qc(op, calId, extra) {
  const payload = Object.assign({ secret: 'ABC', op: op, calId: calId }, extra || {});
  const resp = UrlFetchApp.fetch(QC_PROXY_URL, {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(payload), muteHttpExceptions: true
  });
  let j = null; try { j = JSON.parse(resp.getContentText()); } catch (_e) {}
  return j || {};
}

// Busy intervals [{s,e} ms] for a calendar over [fromDate..toDate].
function _qcBusy(calId, fromDate, toDate) {
  const r = _qc('list', calId, { timeMin: _mytOffset(fromDate, 0, 0), timeMax: _mytOffset(toDate, 23, 59) });
  const items = r.items || [];
  const busy = [];
  items.forEach(function(ev) {
    if (ev.status === 'cancelled') return;
    let s, e;
    if (ev.start && ev.start.dateTime) { s = new Date(ev.start.dateTime).getTime(); e = new Date(ev.end.dateTime).getTime(); }
    else if (ev.start && ev.start.date) { s = new Date(ev.start.date + 'T00:00:00+08:00').getTime(); e = new Date((ev.end && ev.end.date ? ev.end.date : ev.start.date) + 'T00:00:00+08:00').getTime(); }
    else return;
    if (!isNaN(s) && !isNaN(e)) busy.push({ s: s, e: e });
  });
  return busy;
}
// Free slot indices for one day given a busy snapshot.
function _qcFreeSlots(busy, dateStr) {
  const free = [];
  for (let i = 0; i < QC_SLOTS.length; i++) {
    const ss = new Date(_mytOffset(dateStr, QC_SLOTS[i].h, QC_SLOTS[i].m)).getTime();
    const se = ss + 3600000;
    let taken = false;
    for (let j = 0; j < busy.length; j++) { if (busy[j].s < se && busy[j].e > ss) { taken = true; break; } }
    if (!taken) free.push(i);
  }
  return free;
}
// Create one check event; returns {id, error}.
function _qcCreateCheck(calId, dateStr, slotIdx, info) {
  const s = QC_SLOTS[slotIdx];
  const overflow = !!info.overflow;
  const title = (overflow ? '[OVERFLOW] ' : '') + (info.kind || 'WIP') + ' Check D' + info.dayN + '/' + info.total + ' - ' + (info.name || '(no name)');
  const desc = [QC_EVENT_MARK + ' ' + (info.kind || 'WIP'), 'Phone: ' + info.phone, 'Subcon: ' + (info.subcon || ''), 'Day ' + info.dayN + '/' + info.total, 'Address: ' + (info.address || '')].join('\n');
  const event = {
    summary: title, description: desc,
    start: { dateTime: _mytLocal(dateStr, s.h, s.m),     timeZone: 'Asia/Kuala_Lumpur' },
    end:   { dateTime: _mytLocal(dateStr, s.h + 1, s.m), timeZone: 'Asia/Kuala_Lumpur' },
    colorId: overflow ? '11' : '7',
    extendedProperties: { private: { lgqc: '1', phone: String(info.phone), day: String(info.dayN) } }
  };
  const r = _qc('create', calId, { event: event });
  return { id: r.id || null, error: r.error || null };
}
// All of a job's check events on one calendar, sorted by start.
function _qcFindJobEvents(calId, installDate, phone) {
  const from = _opsAddDays(installDate, -2), to = _opsAddDays(installDate, 90);
  const r = _qc('list', calId, { phone: String(phone), timeMin: _mytOffset(from, 0, 0), timeMax: _mytOffset(to, 23, 59) });
  const items = r.items || [];
  const out = items.map(function(ev) {
    const iso = (ev.start && ev.start.dateTime) ? ev.start.dateTime : ((ev.start && ev.start.date) ? ev.start.date : '');
    return { id: ev.id, mytDate: _isoToMytDate(iso), title: ev.summary || '', ts: new Date(iso).getTime() };
  });
  out.sort(function(a, b) { return (a.ts || 0) - (b.ts || 0); });
  return out;
}
function _qcDelete(calId, eventId) { return _qc('delete', calId, { eventId: eventId }); }
function _qcPatch(calId, eventId, patch) { return _qc('patch', calId, { eventId: eventId, patch: patch }); }

// ---------------------------------------------------------------- diagnostic

// action 'qcCalCheck' — can n8n reach both QC calendars?
function handleQcCalCheck(body) {
  const today = _mytDatePlus(0);
  const out = QC_TEAMS.map(function(t) {
    const r = _qc('list', t.calId, { timeMin: _mytOffset(today, 0, 0), timeMax: _mytOffset(today, 23, 59) });
    const accessible = !r.error && (r.items !== undefined);
    return { team: t.name, accessible: accessible, calName: r.summary || null, error: r.error ? (r.error.message || JSON.stringify(r.error)) : null };
  });
  return jsonResponse({ status: 'ok', calendars: out });
}

// ---------------------------------------------------------------- coordinator queue

function handleOpsConfirmQueue(body) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const h = getHeaders(sheet);
  const cName = h.colByName['Name'], cPhone = h.colByName['Phone'], cStatus = h.colByName['Status'],
        cAddr = h.colByName['Full Address'], cLoc = h.colByName['Location'], cProb = h.colByName['Problem Type'],
        cRm = h.colByName['Quotation (RM)'], cLink = h.colByName['Group Invite Link (AJ)'], cChg = h.colByName['Status Changed At'];
  const jobs = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if ((cStatus ? String(row[cStatus - 1] || '').trim() : '') !== 'Pending I.Date') continue;
    jobs.push({
      name:    cName  ? String(row[cName  - 1] || '').trim() : '',
      phone:   cPhone ? String(row[cPhone - 1] || '').trim() : '',
      address: cAddr  ? String(row[cAddr  - 1] || '').trim() : '',
      location:cLoc   ? String(row[cLoc   - 1] || '').trim() : '',
      problem: cProb  ? String(row[cProb  - 1] || '').trim() : '',
      rm:      cRm    ? String(row[cRm    - 1] || '').trim() : '',
      groupLink: cLink ? String(row[cLink - 1] || '').trim() : '',
      changedAt: cChg ? _normalizeDateStr(row[cChg - 1]) : ''
    });
  }
  return jsonResponse({
    status: 'ok', jobs: jobs,
    subconTeams: SUBCON_TEAMS, destStatuses: OPS_DEST_STATUSES,
    qcTeams: QC_TEAMS.map(function(t) { return t.name; }), slots: QC_SLOTS.map(function(s) { return s.label; })
  });
}

// ---------------------------------------------------------------- confirm + assign + block

// body: {phone, installDate, expectedDays, subcon, destStatus, by}
function handleOpsConfirmSave(body) {
  const phone = _normMyPhone(body.phone || '');
  const installDate = _normalizeDateStr(body.installDate || '');
  const expectedDays = Math.round(Number(body.expectedDays || 0));
  const subcon = String(body.subcon || '').trim();
  const destStatus = String(body.destStatus || '').trim();
  const by = String(body.by || 'Coordinator').trim().slice(0, 40);

  if (!phone) return jsonResponse({ status: 'error', message: 'phone required' });
  if (!/^\d{4}-\d{2}-\d{2}$/.test(installDate)) return jsonResponse({ status: 'error', message: 'installDate (YYYY-MM-DD) required' });
  if (!(expectedDays >= 1 && expectedDays <= 60)) return jsonResponse({ status: 'error', message: 'expectedDays must be 1–60' });
  if (SUBCON_TEAMS.indexOf(subcon) === -1) return jsonResponse({ status: 'error', message: 'pick a subcon team' });
  if (OPS_DEST_STATUSES.indexOf(destStatus) === -1) return jsonResponse({ status: 'error', message: 'destStatus must be Pending Downpayment or Job In Progress' });

  const start = _opsParseDate(installDate);
  if (start.getDay() === 0) return jsonResponse({ status: 'error', message: 'Installation date is a Sunday (non-working). Pick Mon–Sat.' });

  const sheet = getSheet();
  const row = findRowByPhone(sheet, phone);
  if (!row) return jsonResponse({ status: 'error', message: 'lead not found' });

  // Already scheduled? (Assigned QC Team set + events exist) → edit instead.
  const existingTeam = (function() { const c = getHeaders(sheet).colByName[OPS_COL.qcTeam]; return c ? String(sheet.getRange(row, c).getValue() || '').trim() : ''; })();
  if (existingTeam) {
    const t0 = _qcTeamByName(existingTeam);
    if (t0 && _qcFindJobEvents(t0.calId, installDate, phone).length) {
      return jsonResponse({ status: 'error', message: 'Already scheduled on ' + existingTeam + '. Use the on-site page to extend/shorten instead.' });
    }
  }

  const name = (function() { const c = getHeaders(sheet).colByName['Name']; return c ? String(sheet.getRange(row, c).getValue() || '').trim() : ''; })();
  const address = (function() { const c = getHeaders(sheet).colByName['Full Address']; return c ? String(sheet.getRange(row, c).getValue() || '').trim() : ''; })();

  const days = _opsWorkingDays(installDate, expectedDays);

  // Busy snapshot per team (one list call each), then eligibility.
  const teamBusy = {};
  QC_TEAMS.forEach(function(t) { teamBusy[t.name] = _qcBusy(t.calId, days[0], days[days.length - 1]); });
  const eligible = []; const fullReport = {};
  QC_TEAMS.forEach(function(t) {
    let ok = true; const fullDays = [];
    days.forEach(function(ds) { if (_qcFreeSlots(teamBusy[t.name], ds).length === 0) { ok = false; fullDays.push(ds); } });
    if (!ok) fullReport[t.name] = fullDays;
    if (ok) eligible.push(t);
  });
  if (!eligible.length) {
    return jsonResponse({ status: 'error', message: 'No QC team has a free slot on every day ' + days[0] + ' → ' + days[days.length - 1] + '. Pick another start date.', fullDays: fullReport });
  }

  const chosen = eligible[Math.floor(Math.random() * eligible.length)];
  const busyC = teamBusy[chosen.name];
  const created = [];
  days.forEach(function(ds, idx) {
    const free = _qcFreeSlots(busyC, ds);
    const slotIdx = free[Math.floor(Math.random() * free.length)];
    const r = _qcCreateCheck(chosen.calId, ds, slotIdx, { name: name, phone: phone, subcon: subcon, address: address, dayN: idx + 1, total: days.length, kind: 'WIP' });
    created.push({ date: ds, slot: QC_SLOTS[slotIdx].label, ok: !r.error });
  });

  const expCompletion = days[days.length - 1];
  _opsEnsureCol(sheet, OPS_COL.instDate);       setCellByHeader(sheet, row, OPS_COL.instDate, installDate);
  _opsEnsureCol(sheet, OPS_COL.expComplete);    setCellByHeader(sheet, row, OPS_COL.expComplete, expCompletion);
  _opsEnsureCol(sheet, OPS_COL.qcTeam);         setCellByHeader(sheet, row, OPS_COL.qcTeam, chosen.name);
  _opsEnsureCol(sheet, OPS_COL.subcon);         setCellByHeader(sheet, row, OPS_COL.subcon, subcon);
  _opsEnsureCol(sheet, OPS_COL.actualComplete);
  // Phase 2 — capture the site WA group link at the natural point (confirmation),
  // now that Whapi no longer auto-fills it. Optional; only writes when provided.
  if (body.groupLink) { try { _opsEnsureCol(sheet, 'Group Invite Link (AJ)'); setCellByHeader(sheet, row, 'Group Invite Link (AJ)', String(body.groupLink).trim().slice(0, 300)); } catch (_e) {} }

  let prevStatus = '';
  try { prevStatus = String(sheet.getRange(row, getHeaders(sheet).colByName['Status']).getValue() || '').trim(); } catch (_e) {}
  setCellByHeader(sheet, row, 'Status', destStatus);
  setCellByHeader(sheet, row, 'Status Changed At', new Date().toISOString());
  setCellByHeader(sheet, row, 'Changed By', 'OpsConfirm:' + by);
  try { _resetFuCadenceIfPhaseChanged(sheet, row, prevStatus, destStatus); } catch (_e) {}

  return jsonResponse({ status: 'ok', qcTeam: chosen.name, subcon: subcon, installDate: installDate, expectedCompletion: expCompletion, destStatus: destStatus, visits: created });
}

// ---------------------------------------------------------------- on-site active jobs

function handleOpsOnsiteJobs(body) {
  const team = String(body.team || '').trim();
  if (!_qcTeamByName(team)) return jsonResponse({ status: 'error', message: 'unknown QC team' });
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const h = getHeaders(sheet);
  const cName = h.colByName['Name'], cPhone = h.colByName['Phone'], cStatus = h.colByName['Status'],
        cAddr = h.colByName['Full Address'], cLink = h.colByName['Group Invite Link (AJ)'],
        cInst = h.colByName[OPS_COL.instDate], cExp = h.colByName[OPS_COL.expComplete],
        cQc = h.colByName[OPS_COL.qcTeam], cSub = h.colByName[OPS_COL.subcon];
  if (!cQc || !cInst) return jsonResponse({ status: 'ok', jobs: [] });
  const DONE = ['Completed', 'Job Complete', 'Receipt Sent', 'Rejected', 'Lost'];
  const jobs = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[cQc - 1] || '').trim().toLowerCase() !== team.toLowerCase()) continue;
    const inst = _normalizeDateStr(row[cInst - 1]); if (!inst) continue;
    const status = cStatus ? String(row[cStatus - 1] || '').trim() : '';
    if (DONE.indexOf(status) !== -1) continue;
    jobs.push({
      name:    cName  ? String(row[cName  - 1] || '').trim() : '',
      phone:   cPhone ? String(row[cPhone - 1] || '').trim() : '',
      address: cAddr  ? String(row[cAddr  - 1] || '').trim() : '',
      status:  status,
      groupLink: cLink ? String(row[cLink - 1] || '').trim() : '',
      instDate: inst,
      expComplete: cExp ? _normalizeDateStr(row[cExp - 1]) : '',
      subcon:  cSub   ? String(row[cSub  - 1] || '').trim() : ''
    });
  }
  jobs.sort(function(a, b) { return String(a.expComplete).localeCompare(String(b.expComplete)); });
  return jsonResponse({ status: 'ok', team: team, jobs: jobs, destStatuses: OPS_DEST_STATUSES });
}

// ---------------------------------------------------------------- extend / shorten

// body: {phone, newExpectedDays? , newCompletionDate? , by}
function handleOpsEditCompletion(body) {
  const phone = _normMyPhone(body.phone || '');
  const by = String(body.by || '').trim().slice(0, 40);
  if (!phone) return jsonResponse({ status: 'error', message: 'phone required' });

  const sheet = getSheet();
  const row = findRowByPhone(sheet, phone);
  if (!row) return jsonResponse({ status: 'error', message: 'lead not found' });
  const h = getHeaders(sheet);
  const instDate = h.colByName[OPS_COL.instDate] ? _normalizeDateStr(sheet.getRange(row, h.colByName[OPS_COL.instDate]).getValue()) : '';
  const teamName = h.colByName[OPS_COL.qcTeam] ? String(sheet.getRange(row, h.colByName[OPS_COL.qcTeam]).getValue() || '').trim() : '';
  if (!instDate || !teamName) return jsonResponse({ status: 'error', message: 'job not scheduled yet (no installation date / QC team)' });
  const team = _qcTeamByName(teamName);
  if (!team) return jsonResponse({ status: 'error', message: 'assigned QC team not recognised: ' + teamName });

  let newN;
  if (body.newCompletionDate) {
    newN = _opsWorkingDaySpan(instDate, _normalizeDateStr(body.newCompletionDate));
    if (!newN) return jsonResponse({ status: 'error', message: 'new completion date must be on/after the installation date' });
  } else {
    newN = Math.round(Number(body.newExpectedDays || 0));
  }
  if (!(newN >= 1 && newN <= 60)) return jsonResponse({ status: 'error', message: 'new duration must be 1–60 working days' });

  const name = h.colByName['Name'] ? String(sheet.getRange(row, h.colByName['Name']).getValue() || '').trim() : '';
  const address = h.colByName['Full Address'] ? String(sheet.getRange(row, h.colByName['Full Address']).getValue() || '').trim() : '';
  const subcon = h.colByName[OPS_COL.subcon] ? String(sheet.getRange(row, h.colByName[OPS_COL.subcon]).getValue() || '').trim() : '';

  const existing = _qcFindJobEvents(team.calId, instDate, phone);
  const curN = existing.length;
  const days = _opsWorkingDays(instDate, newN);
  let deleted = 0; const added = []; const overflowDays = [];

  if (newN < curN) {
    for (let i = newN; i < curN; i++) { _qcDelete(team.calId, existing[i].id); deleted++; }
  } else if (newN > curN) {
    const busyExt = _qcBusy(team.calId, days[curN], days[newN - 1]);
    for (let i = curN; i < newN; i++) {
      const ds = days[i];
      const free = _qcFreeSlots(busyExt, ds);
      const overflow = free.length === 0;
      const slotIdx = overflow ? 0 : free[Math.floor(Math.random() * free.length)];
      _qcCreateCheck(team.calId, ds, slotIdx, { name: name, phone: phone, subcon: subcon, address: address, dayN: i + 1, total: newN, kind: 'WIP', overflow: overflow });
      added.push({ date: ds, slot: QC_SLOTS[slotIdx].label, overflow: overflow });
      if (overflow) overflowDays.push(ds);
    }
  }

  const expCompletion = days[days.length - 1];
  _opsEnsureCol(sheet, OPS_COL.expComplete);
  setCellByHeader(sheet, row, OPS_COL.expComplete, expCompletion);
  try { setCellByHeader(sheet, row, 'Changed By', 'OpsEdit:' + by); } catch (_e) {}

  return jsonResponse({ status: 'ok', qcTeam: team.name, newDays: newN, prevDays: curN, deleted: deleted, added: added, overflowDays: overflowDays, expectedCompletion: expCompletion });
}

// ---------------------------------------------------------------- mark a check done

// body: {phone, date?, by, note?}
function handleOpsMarkCheck(body) {
  const phone = _normMyPhone(body.phone || '');
  const by = String(body.by || '').trim().slice(0, 40);
  const date = _normalizeDateStr(body.date || '') || _mytDatePlus(0);
  const note = String(body.note || '').slice(0, 300);
  if (!phone) return jsonResponse({ status: 'error', message: 'phone required' });

  const sheet = getSheet();
  const row = findRowByPhone(sheet, phone);
  const name = (row && getHeaders(sheet).colByName['Name']) ? String(sheet.getRange(row, getHeaders(sheet).colByName['Name']).getValue() || '').trim() : '';

  try { _callsSheet_().appendRow([new Date().toISOString(), phone, name, by, 'qc_check', date, note]); } catch (_e) {}

  // Phase 2 — if this check IS the defect inspection (repair tag + Repair Date == this
  // date), stamp 'Defect Inspected At' to start the warranty SLA clock (once).
  try {
    const h2 = getHeaders(sheet);
    const tags = (row && h2.colByName['Tags']) ? String(sheet.getRange(row, h2.colByName['Tags']).getValue() || '').toLowerCase() : '';
    const repDate = (row && h2.colByName['Repair Date']) ? _normalizeDateStr(sheet.getRange(row, h2.colByName['Repair Date']).getValue()) : '';
    const inspCol = h2.colByName['Defect Inspected At'];
    const alreadyInsp = inspCol ? String(sheet.getRange(row, inspCol).getValue() || '').trim() : '';
    if (row && tags.indexOf('repair') !== -1 && repDate && repDate === date && !alreadyInsp) {
      _opsEnsureCol(sheet, 'Defect Inspected At');
      setCellByHeader(sheet, row, 'Defect Inspected At', date);
    }
  } catch (_e) {}

  // Green the matching day's event (best-effort, via proxy).
  try {
    const h = getHeaders(sheet);
    const teamName = (row && h.colByName[OPS_COL.qcTeam]) ? String(sheet.getRange(row, h.colByName[OPS_COL.qcTeam]).getValue() || '').trim() : '';
    const instDate = (row && h.colByName[OPS_COL.instDate]) ? _normalizeDateStr(sheet.getRange(row, h.colByName[OPS_COL.instDate]).getValue()) : '';
    const team = _qcTeamByName(teamName);
    if (team && instDate) {
      _qcFindJobEvents(team.calId, instDate, phone).forEach(function(ev) {
        if (ev.mytDate === date) {
          const t = ev.title || '';
          _qcPatch(team.calId, ev.id, { colorId: '10', summary: (t.indexOf('✓') === -1 ? '✓ ' + t : t) });
        }
      });
    }
  } catch (_e) {}

  return jsonResponse({ status: 'ok', phone: phone, date: date });
}

// ================================================================
// Phase 2 — Coordinator Cockpit
// All handlers below are ADDITIVE (new router cases + lazily-created columns).
// New columns: 'Day1 Assessment' (JSON), 'WIP Phases' (JSON), 'Defect Inspected At',
// 'Defect Responded At'. Existing columns reused as-is.
// ================================================================

// +n working days (Sundays excluded) after a start date; n=0 returns start.
function _opsAddWorkingDays(startStr, n) {
  if (!startStr) return '';
  const d = _opsParseDate(startStr); let c = 0;
  if (n <= 0) return _opsFmtDate(d);
  while (c < n) { d.setDate(d.getDate() + 1); if (d.getDay() !== 0) c++; }
  return _opsFmtDate(d);
}

// Read-only snapshot for the 3 cockpit panels.
function handleOpsCockpit(body) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const h = getHeaders(sheet);
  const C = function (n) { return h.colByName[n]; };
  const cName=C('Name'), cPhone=C('Phone'), cAddr=C('Full Address'), cLoc=C('Location'),
        cStatus=C('Status'), cTags=C('Tags'), cLink=C('Group Invite Link (AJ)'),
        cInst=C(OPS_COL.instDate), cExp=C(OPS_COL.expComplete), cQc=C(OPS_COL.qcTeam),
        cSub=C(OPS_COL.subcon), cActual=C(OPS_COL.actualComplete),
        cA1=C('Day1 Assessment'), cWip=C('WIP Phases'),
        cRepDate=C('Repair Date'), cRepTeam=C('Repair QC Team'),
        cInsp=C('Defect Inspected At'), cResp=C('Defect Responded At');
  const today = _mytDatePlus(0);
  const V = function (row, c) { return c ? String(row[c-1]==null?'':row[c-1]).trim() : ''; };
  const D = function (row, c) { return c ? _normalizeDateStr(row[c-1]) : ''; };
  const J = function (s) { try { return JSON.parse(s); } catch (_e) { return null; } };
  const starting = [], wip = [], defect = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const phone = V(row, cPhone); if (!phone) continue;
    const status = V(row, cStatus);
    const tags = V(row, cTags).split(',').map(function (t) { return t.trim(); }).filter(Boolean);
    const inst = D(row, cInst);
    const actual = D(row, cActual);
    const base = { name: V(row, cName), phone: phone, address: V(row, cAddr) || V(row, cLoc),
                   groupLink: V(row, cLink), status: status, qcTeam: V(row, cQc), subcon: V(row, cSub) };
    // Panel 1 — Starting Today
    if (inst && inst === today && OPS_DEST_STATUSES.indexOf(status) !== -1 && !actual) {
      starting.push(Object.assign({}, base, { instDate: inst, assessment: J(V(row, cA1)) }));
    }
    // Panel 2 — Active WIP (installed before today, not complete)
    if (inst && inst < today && !actual && (status === 'Job In Progress' || status === 'Pending Downpayment')) {
      const a = J(V(row, cA1)); const ph = J(V(row, cWip)) || {};
      wip.push(Object.assign({}, base, { instDate: inst, expComplete: D(row, cExp),
        phases: { prep: ph.prep || '', primer: ph.primer || '', membrane: ph.membrane || '' },
        day1flags: (a && a.flags) ? a.flags : [] }));
    }
    // Panel 3 — Defect SLA
    if (tags.indexOf('repair') !== -1) {
      const insp = V(row, cInsp), resp = V(row, cResp);
      let respBy = '', repBy = '', respOver = false, repOver = false;
      if (insp) {
        const inspDate = insp.slice(0, 10);
        respBy = _opsAddWorkingDays(inspDate, 3);
        repBy  = _opsAddWorkingDays(inspDate, 10);
        respOver = !resp && respBy && respBy < today;
        repOver  = repBy && repBy < today;
      }
      defect.push(Object.assign({}, base, { repairDate: D(row, cRepDate), repairTeam: V(row, cRepTeam),
        inspectedAt: insp ? insp.slice(0, 10) : '', respondedAt: resp ? resp.slice(0, 10) : '',
        respondBy: respBy, repairBy: repBy, respondOverdue: respOver, repairOverdue: repOver }));
    }
  }
  starting.sort(function (a, b) { return String(a.name).localeCompare(String(b.name)); });
  wip.sort(function (a, b) { return String(a.expComplete).localeCompare(String(b.expComplete)); });
  defect.sort(function (a, b) { return String(a.repairBy || '9999').localeCompare(String(b.repairBy || '9999')); });
  return jsonResponse({ status: 'ok', today: today, starting: starting, wip: wip, defect: defect });
}

// QC on-site team records the day-1 substrate assessment.
// body: {phone, flags:[...], notes, scope:'extra'|'none', by}
function handleOpsSaveAssessment(body) {
  const phone = _normMyPhone(body.phone || ''); const by = String(body.by || '').slice(0, 40);
  if (!phone) return jsonResponse({ status: 'error', message: 'phone required' });
  const sheet = getSheet(); const row = findRowByPhone(sheet, phone);
  if (!row) return jsonResponse({ status: 'error', message: 'lead not found' });
  const flags = Array.isArray(body.flags) ? body.flags.map(function (f) { return String(f).slice(0, 30); }).slice(0, 20) : [];
  const payload = { flags: flags, notes: String(body.notes || '').slice(0, 1000),
                    scope: (body.scope === 'extra' ? 'extra' : 'none'), by: by, at: new Date().toISOString() };
  _opsEnsureCol(sheet, 'Day1 Assessment');
  setCellByHeader(sheet, row, 'Day1 Assessment', JSON.stringify(payload));
  return jsonResponse({ status: 'ok', phone: phone, assessment: payload });
}

// Coordinator ticks a WIP phase proof-photo. body: {phone, phase:'prep'|'primer'|'membrane', on, by}
function handleOpsTickPhase(body) {
  const phone = _normMyPhone(body.phone || ''); const phase = String(body.phase || '');
  const on = (body.on !== false);
  if (!phone || ['prep', 'primer', 'membrane'].indexOf(phase) === -1) return jsonResponse({ status: 'error', message: 'phone and valid phase required' });
  const sheet = getSheet(); const row = findRowByPhone(sheet, phone);
  if (!row) return jsonResponse({ status: 'error', message: 'lead not found' });
  _opsEnsureCol(sheet, 'WIP Phases');
  const c = getHeaders(sheet).colByName['WIP Phases'];
  let cur = {}; try { cur = JSON.parse(String(sheet.getRange(row, c).getValue() || '') || '{}') || {}; } catch (_e) { cur = {}; }
  cur[phase] = on ? new Date().toISOString() : '';
  setCellByHeader(sheet, row, 'WIP Phases', JSON.stringify(cur));
  return jsonResponse({ status: 'ok', phase: phase, on: on, phases: cur });
}

// Defect: coordinator marks they've responded to the customer. body: {phone, by}
function handleOpsMarkResponded(body) {
  const phone = _normMyPhone(body.phone || ''); const by = String(body.by || '').slice(0, 40);
  if (!phone) return jsonResponse({ status: 'error', message: 'phone required' });
  const sheet = getSheet(); const row = findRowByPhone(sheet, phone);
  if (!row) return jsonResponse({ status: 'error', message: 'lead not found' });
  _opsEnsureCol(sheet, 'Defect Responded At');
  const today = _mytDatePlus(0);
  setCellByHeader(sheet, row, 'Defect Responded At', today);
  return jsonResponse({ status: 'ok', phone: phone, respondedAt: today });
}

// Defect: mark repaired / close — drops it off the SLA panel (clears 'repair' tag). body: {phone, by, note?}
function handleOpsMarkRepaired(body) {
  const phone = _normMyPhone(body.phone || ''); const by = String(body.by || '').slice(0, 40);
  if (!phone) return jsonResponse({ status: 'error', message: 'phone required' });
  const sheet = getSheet(); const row = findRowByPhone(sheet, phone);
  if (!row) return jsonResponse({ status: 'error', message: 'lead not found' });
  const h = getHeaders(sheet); const tc = h.colByName['Tags'];
  if (tc) {
    const cur = String(sheet.getRange(row, tc).getValue() || '').split(',').map(function (t) { return t.trim(); })
      .filter(Boolean).filter(function (t) { return t !== 'repair'; });
    setCellByHeader(sheet, row, 'Tags', cur.join(','));
  }
  const name = h.colByName['Name'] ? String(sheet.getRange(row, h.colByName['Name']).getValue() || '').trim() : '';
  try { _callsSheet_().appendRow([new Date().toISOString(), phone, name, by, 'defect_repaired', _mytDatePlus(0), String(body.note || '').slice(0, 300)]); } catch (_e) {}
  return jsonResponse({ status: 'ok', phone: phone });
}

// Manual: coordinator sets/updates the site WA group link. body: {phone, link, by}
function handleSetGroupLink(body) {
  const phone = _normMyPhone(body.phone || ''); const link = String(body.link || '').trim().slice(0, 300);
  if (!phone || !link) return jsonResponse({ status: 'error', message: 'phone and link required' });
  const sheet = getSheet(); const row = findRowByPhone(sheet, phone);
  if (!row) return jsonResponse({ status: 'error', message: 'lead not found' });
  _opsEnsureCol(sheet, 'Group Invite Link (AJ)');
  setCellByHeader(sheet, row, 'Group Invite Link (AJ)', link);
  return jsonResponse({ status: 'ok', phone: phone, link: link });
}

// Manual: coordinator alerts staff on Telegram about an overdue SLA. body: {phone, message?}
function handleOpsAlertStaff(body) {
  const phone = _normMyPhone(body.phone || '');
  if (!phone) return jsonResponse({ status: 'error', message: 'phone required' });
  const sheet = getSheet(); const row = findRowByPhone(sheet, phone);
  const h = getHeaders(sheet);
  const name = (row && h.colByName['Name']) ? String(sheet.getRange(row, h.colByName['Name']).getValue() || '').trim() : '';
  const msg = String(body.message || ('⚠ Defect SLA overdue — ' + (name || phone))).slice(0, 500);
  // Route via the n8n LG-Staff-Alert webhook so the Telegram bot token stays out
  // of this (public) repo. The 'ABC' shared secret is already public on the pages.
  try {
    UrlFetchApp.fetch('https://leakguard.app.n8n.cloud/webhook/lg-staff-alert', {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ secret: 'ABC', text: msg }), muteHttpExceptions: true
    });
  } catch (_e) {}
  return jsonResponse({ status: 'ok' });
}
