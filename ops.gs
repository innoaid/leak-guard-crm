// ================================================================
// Operations Coordination + On-Site QC scheduling  (Phase 1)
// ----------------------------------------------------------------
// Adds the operations layer on top of the kanban:
//   • Coordinator confirms Installation Date + Expected Days + Subcon
//     team for a "Pending I.Date" card (ops_confirm.html).
//   • System random-assigns 1 of 2 QC (check-visit) teams that has a
//     vacant slot on EVERY working day of the install window, then
//     blocks one vacant slot per day on that team's Google Calendar.
//   • On-site team can extend / shorten the window (auto-reflow the
//     trailing check-slots; an extension that overflows a day is
//     soft-allowed + amber-flagged) and mark a daily check done.
//
// QC team = CHECK crew (one slot = one check visit). Subcon = labour
// (chosen manually by the coordinator, recorded on the card, NOT
// scheduled). System schedules only the 2 QC calendars.  Sunday is
// never a working day.
//
// Router cases live in kanban_code.gs doPost():
//   opsConfirmQueue · opsConfirmSave · opsOnsiteJobs ·
//   opsEditCompletion · opsMarkCheck · qcCalCheck
// ================================================================

const QC_TEAMS = [
  { name: 'QC Team A', calId: 'fac7acef517820bbc6d98f92a58dfce3a67f88882b6f80a704f0a6ac5c317a1f@group.calendar.google.com' },
  { name: 'QC Team B', calId: '25df23d5f659f8ba3598a6f94acd118366e77ba203f6681a4a21ac7348451fae@group.calendar.google.com' }
];

// 4 one-hour check slots/day (24h local = MYT; script TZ must be GMT+8).
const QC_SLOTS = [
  { label: '9:00–10:00',  h: 9,  m: 0  },
  { label: '10:30–11:30', h: 10, m: 30 },
  { label: '1:00–2:00',   h: 13, m: 0  },
  { label: '3:00–4:00',   h: 15, m: 0  }
];

const SUBCON_TEAMS = ['Atik (A)', 'Atik (B)', 'Atik (C)', 'Mohiuddin', 'Danny'];

// Where the coordinator may move the card on confirm (point 1).
const OPS_DEST_STATUSES = ['Pending Downpayment', 'Job In Progress'];

// New Leads-sheet columns (lazily created, like 'Booked By').
const OPS_COL = {
  instDate:       'Installation Date',
  expComplete:    'Expected Completion Date',
  qcTeam:         'Assigned QC Team',
  subcon:         'Assigned Subcon Team',
  actualComplete: 'Actual Completion Date'
};

// Stable marker written into every QC check event's description so we can
// find/extend/shorten a job's events later (search by this + 'Phone: <ph>').
const QC_EVENT_MARK = 'LGQC';

// ---------------------------------------------------------------- helpers

// Lazily ensure a header column exists; returns its 1-based index.
function _opsEnsureCol(sheet, name) {
  let col = getHeaders(sheet).colByName[name];
  if (!col) { col = sheet.getLastColumn() + 1; sheet.getRange(1, col).setValue(name); }
  return col;
}

// 'YYYY-MM-DD' -> Date at local midnight.
function _opsParseDate(str) {
  const m = String(str || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 0, 0, 0, 0);
}
// Date -> 'YYYY-MM-DD' (local parts).
function _opsFmtDate(d) {
  const p = function(n) { return (n < 10 ? '0' : '') + n; };
  return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
}

// n working days from startStr (inclusive), skipping Sundays. Returns ['YYYY-MM-DD', ...].
function _opsWorkingDays(startStr, n) {
  const out = [];
  const d = _opsParseDate(startStr);
  if (!d) return out;
  while (out.length < n) {
    if (d.getDay() !== 0) out.push(_opsFmtDate(d));   // 0 = Sunday
    d.setDate(d.getDate() + 1);
  }
  return out;
}

// Count working days (Sun excluded) from startStr to endStr inclusive. 0 if invalid/backwards.
function _opsWorkingDaySpan(startStr, endStr) {
  const a = _opsParseDate(startStr), b = _opsParseDate(endStr);
  if (!a || !b || b < a) return 0;
  let n = 0; const d = new Date(a.getTime());
  while (d <= b) { if (d.getDay() !== 0) n++; d.setDate(d.getDate() + 1); }
  return n;
}

// Resolve a QC team's calendar by id; throws a clear, fixable error if the
// Apps Script account can't see it (fix = share the calendar with the script
// owner account, "Make changes to events").
function _qcCal(calId) {
  const cal = CalendarApp.getCalendarById(calId);
  if (!cal) {
    throw new Error('QC calendar not accessible (' + calId + '). Share it with the Apps Script owner account with "Make changes to events" permission.');
  }
  return cal;
}
function _qcTeamByName(name) {
  const t = String(name || '').trim();
  for (let i = 0; i < QC_TEAMS.length; i++) if (QC_TEAMS[i].name.toLowerCase() === t.toLowerCase()) return QC_TEAMS[i];
  return null;
}

// Free slot indices for one team-calendar on one day (a slot is taken if any
// event overlaps its hour). Returns array of indices into QC_SLOTS.
function _qcDayFreeSlots(cal, dateStr) {
  const dayStart = _opsParseDate(dateStr);
  const dayEnd = new Date(dayStart.getTime()); dayEnd.setDate(dayEnd.getDate() + 1);
  const events = cal.getEvents(dayStart, dayEnd);
  const free = [];
  for (let i = 0; i < QC_SLOTS.length; i++) {
    const s = QC_SLOTS[i];
    const slotStart = new Date(dayStart.getTime()); slotStart.setHours(s.h, s.m, 0, 0);
    const slotEnd = new Date(slotStart.getTime() + 60 * 60 * 1000);
    let taken = false;
    for (let j = 0; j < events.length; j++) {
      if (events[j].getStartTime() < slotEnd && events[j].getEndTime() > slotStart) { taken = true; break; }
    }
    if (!taken) free.push(i);
  }
  return free;
}

// Create one QC check event in a given slot; returns the CalendarEvent.
function _qcCreateCheck(cal, teamName, dateStr, slotIdx, info) {
  const s = QC_SLOTS[slotIdx];
  const start = _opsParseDate(dateStr); start.setHours(s.h, s.m, 0, 0);
  const end = new Date(start.getTime() + 60 * 60 * 1000);
  const kind = info.kind || 'WIP';
  const overflow = !!info.overflow;
  const title = (overflow ? '[OVERFLOW] ' : '') +
    kind + ' Check D' + info.dayN + '/' + info.total + ' — ' + (info.name || '(no name)');
  const desc = [
    QC_EVENT_MARK + ' ' + kind,
    'Phone: ' + info.phone,
    'Subcon: ' + (info.subcon || ''),
    'Day ' + info.dayN + '/' + info.total,
    'Address: ' + (info.address || '')
  ].join('\n');
  const ev = cal.createEvent(title, start, end, { description: desc });
  try { ev.setColor(overflow ? CalendarApp.EventColor.RED : CalendarApp.EventColor.PALE_BLUE); } catch (_e) {}
  return ev;
}

// All QC check events for one job on one team calendar, sorted by start.
// Searches installDate-2 .. installDate+90 and matches the phone marker.
function _qcFindJobEvents(cal, installDateStr, phone) {
  const from = _opsParseDate(installDateStr); from.setDate(from.getDate() - 2);
  const to = _opsParseDate(installDateStr); to.setDate(to.getDate() + 90);
  const ph = String(phone).trim();
  const all = cal.getEvents(from, to);
  const mine = all.filter(function(ev) {
    const d = ev.getDescription() || '';
    return d.indexOf(QC_EVENT_MARK) !== -1 && d.indexOf('Phone: ' + ph) !== -1;
  });
  mine.sort(function(a, b) { return a.getStartTime() - b.getStartTime(); });
  return mine;
}

// ---------------------------------------------------------------- diagnostic

// action 'qcCalCheck' — verify the script can see both QC calendars.
function handleQcCalCheck(body) {
  const out = QC_TEAMS.map(function(t) {
    let cal = null;
    try { cal = CalendarApp.getCalendarById(t.calId); } catch (_e) {}
    return { team: t.name, calId: t.calId, accessible: !!cal, calName: cal ? cal.getName() : null };
  });
  return jsonResponse({ status: 'ok', calendars: out });
}

// ---------------------------------------------------------------- coordinator queue

// action 'opsConfirmQueue' — cards awaiting installation scheduling.
function handleOpsConfirmQueue(body) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const h = getHeaders(sheet);
  const cName = h.colByName['Name'], cPhone = h.colByName['Phone'],
        cStatus = h.colByName['Status'], cAddr = h.colByName['Full Address'],
        cLoc = h.colByName['Location'], cProb = h.colByName['Problem Type'],
        cRm = h.colByName['Quotation (RM)'], cLink = h.colByName['Group Invite Link (AJ)'],
        cChg = h.colByName['Status Changed At'],
        cInst = h.colByName[OPS_COL.instDate], cSub = h.colByName[OPS_COL.subcon],
        cQc = h.colByName[OPS_COL.qcTeam];
  const jobs = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = cStatus ? String(row[cStatus - 1] || '').trim() : '';
    if (status !== 'Pending I.Date') continue;
    jobs.push({
      name:    cName  ? String(row[cName  - 1] || '').trim() : '',
      phone:   cPhone ? String(row[cPhone - 1] || '').trim() : '',
      address: cAddr  ? String(row[cAddr  - 1] || '').trim() : '',
      location:cLoc   ? String(row[cLoc   - 1] || '').trim() : '',
      problem: cProb  ? String(row[cProb  - 1] || '').trim() : '',
      rm:      cRm    ? String(row[cRm    - 1] || '').trim() : '',
      groupLink: cLink ? String(row[cLink - 1] || '').trim() : '',
      changedAt: cChg ? _normalizeDateStr(row[cChg - 1]) : '',
      instDate: cInst ? _normalizeDateStr(row[cInst - 1]) : '',
      subcon:  cSub   ? String(row[cSub  - 1] || '').trim() : '',
      qcTeam:  cQc    ? String(row[cQc   - 1] || '').trim() : ''
    });
  }
  return jsonResponse({
    status: 'ok', jobs: jobs,
    subconTeams: SUBCON_TEAMS,
    destStatuses: OPS_DEST_STATUSES,
    qcTeams: QC_TEAMS.map(function(t) { return t.name; }),
    slots: QC_SLOTS.map(function(s) { return s.label; })
  });
}

// ---------------------------------------------------------------- confirm + assign + block

// action 'opsConfirmSave'
// body: {phone, installDate:'YYYY-MM-DD', expectedDays:int, subcon, destStatus, by}
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

  // Guard: already scheduled? (Assigned QC Team set + events exist) → ask to edit instead.
  const existingTeam = (function() { const c = getHeaders(sheet).colByName[OPS_COL.qcTeam]; return c ? String(sheet.getRange(row, c).getValue() || '').trim() : ''; })();
  if (existingTeam) {
    const t0 = _qcTeamByName(existingTeam);
    if (t0) {
      const prior = _qcFindJobEvents(_qcCal(t0.calId), installDate, phone);
      if (prior.length) return jsonResponse({ status: 'error', message: 'Already scheduled on ' + existingTeam + ' (' + prior.length + ' visits). Use the on-site page to extend/shorten instead.' });
    }
  }

  const name = (function() { const c = getHeaders(sheet).colByName['Name']; return c ? String(sheet.getRange(row, c).getValue() || '').trim() : ''; })();
  const address = (function() { const c = getHeaders(sheet).colByName['Full Address']; return c ? String(sheet.getRange(row, c).getValue() || '').trim() : ''; })();

  const days = _opsWorkingDays(installDate, expectedDays);

  // Eligibility: a team must have ≥1 free slot on EVERY day of the window.
  const cals = QC_TEAMS.map(function(t) { return { team: t, cal: _qcCal(t.calId) }; });
  const eligible = [];
  const fullReport = {};
  cals.forEach(function(c) {
    let ok = true; const fullDays = [];
    days.forEach(function(ds) { if (_qcDayFreeSlots(c.cal, ds).length === 0) { ok = false; fullDays.push(ds); } });
    if (!ok) fullReport[c.team.name] = fullDays;
    if (ok) eligible.push(c);
  });

  if (!eligible.length) {
    return jsonResponse({ status: 'error', message: 'No QC team has a free slot on every day ' + days[0] + ' → ' + days[days.length - 1] + '. Pick another start date.', fullDays: fullReport });
  }

  // Random among teams that fit (same spirit as inspection random-assign).
  const chosen = eligible[Math.floor(Math.random() * eligible.length)];
  const created = [];
  days.forEach(function(ds, idx) {
    const free = _qcDayFreeSlots(chosen.cal, ds);
    const slotIdx = free[Math.floor(Math.random() * free.length)];
    _qcCreateCheck(chosen.cal, chosen.team.name, ds, slotIdx, {
      name: name, phone: phone, subcon: subcon, address: address,
      dayN: idx + 1, total: days.length, kind: 'WIP'
    });
    created.push({ date: ds, slot: QC_SLOTS[slotIdx].label });
  });

  // Persist fields + move status. Ensure every ops column exists first
  // (setCellByHeader throws on a missing header).
  const expCompletion = days[days.length - 1];
  _opsEnsureCol(sheet, OPS_COL.instDate);       setCellByHeader(sheet, row, OPS_COL.instDate, installDate);
  _opsEnsureCol(sheet, OPS_COL.expComplete);    setCellByHeader(sheet, row, OPS_COL.expComplete, expCompletion);
  _opsEnsureCol(sheet, OPS_COL.qcTeam);         setCellByHeader(sheet, row, OPS_COL.qcTeam, chosen.team.name);
  _opsEnsureCol(sheet, OPS_COL.subcon);         setCellByHeader(sheet, row, OPS_COL.subcon, subcon);
  _opsEnsureCol(sheet, OPS_COL.actualComplete);

  let prevStatus = '';
  try { prevStatus = String(sheet.getRange(row, getHeaders(sheet).colByName['Status']).getValue() || '').trim(); } catch (_e) {}
  setCellByHeader(sheet, row, 'Status', destStatus);
  setCellByHeader(sheet, row, 'Status Changed At', new Date().toISOString());
  setCellByHeader(sheet, row, 'Changed By', 'OpsConfirm:' + by);
  try { _resetFuCadenceIfPhaseChanged(sheet, row, prevStatus, destStatus); } catch (_e) {}

  return jsonResponse({
    status: 'ok', qcTeam: chosen.team.name, subcon: subcon,
    installDate: installDate, expectedCompletion: expCompletion,
    destStatus: destStatus, visits: created
  });
}

// ---------------------------------------------------------------- on-site active jobs

// action 'opsOnsiteJobs'  body: {team}
// Active install jobs for a QC team (has Installation Date, not yet completed).
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
  if (!cQc || !cInst) return jsonResponse({ status: 'ok', jobs: [] });  // columns not created yet
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

// action 'opsEditCompletion'
// body: {phone, newExpectedDays? , newCompletionDate? , by}
// Reflows the trailing check-slots on the assigned QC team. Shorten → delete
// trailing visits. Extend → add visits; a day with no free slot is soft-allowed
// (created anyway, [OVERFLOW] + red) so the check for already-happening work is
// never refused.
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

  // New day count.
  let newN;
  if (body.newCompletionDate) {
    const nc = _normalizeDateStr(body.newCompletionDate);
    newN = _opsWorkingDaySpan(instDate, nc);
    if (!newN) return jsonResponse({ status: 'error', message: 'new completion date must be on/after the installation date' });
  } else {
    newN = Math.round(Number(body.newExpectedDays || 0));
  }
  if (!(newN >= 1 && newN <= 60)) return jsonResponse({ status: 'error', message: 'new duration must be 1–60 working days' });

  const cal = _qcCal(team.calId);
  const name = h.colByName['Name'] ? String(sheet.getRange(row, h.colByName['Name']).getValue() || '').trim() : '';
  const address = h.colByName['Full Address'] ? String(sheet.getRange(row, h.colByName['Full Address']).getValue() || '').trim() : '';
  const subcon = h.colByName[OPS_COL.subcon] ? String(sheet.getRange(row, h.colByName[OPS_COL.subcon]).getValue() || '').trim() : '';

  const existing = _qcFindJobEvents(cal, instDate, phone);
  const curN = existing.length;
  const days = _opsWorkingDays(instDate, newN);
  let deleted = 0; const added = []; const overflowDays = [];

  if (newN < curN) {
    // Shorten → remove the trailing (curN-newN) visits, free the slots.
    for (let i = newN; i < curN; i++) { try { existing[i].deleteEvent(); deleted++; } catch (_e) {} }
  } else if (newN > curN) {
    // Extend → add a visit for each new day; soft-overflow if the day is full.
    for (let i = curN; i < newN; i++) {
      const ds = days[i];
      const free = _qcDayFreeSlots(cal, ds);
      const overflow = free.length === 0;
      const slotIdx = overflow ? 0 : free[Math.floor(Math.random() * free.length)];
      _qcCreateCheck(cal, team.name, ds, slotIdx, {
        name: name, phone: phone, subcon: subcon, address: address,
        dayN: i + 1, total: newN, kind: 'WIP', overflow: overflow
      });
      added.push({ date: ds, slot: QC_SLOTS[slotIdx].label, overflow: overflow });
      if (overflow) overflowDays.push(ds);
    }
  }

  const expCompletion = days[days.length - 1];
  _opsEnsureCol(sheet, OPS_COL.expComplete);
  setCellByHeader(sheet, row, OPS_COL.expComplete, expCompletion);
  try { setCellByHeader(sheet, row, 'Changed By', 'OpsEdit:' + by); } catch (_e) {}

  return jsonResponse({
    status: 'ok', qcTeam: team.name, newDays: newN, prevDays: curN,
    deleted: deleted, added: added, overflowDays: overflowDays,
    expectedCompletion: expCompletion
  });
}

// ---------------------------------------------------------------- mark a check done

// action 'opsMarkCheck'  body: {phone, date?, by, note?}
// Append-only log of a completed daily check (feeds later reporting / photo
// double-check). Also greens the day's event if found.
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

  // Green the matching day's event (best-effort).
  try {
    const teamName = (row && getHeaders(sheet).colByName[OPS_COL.qcTeam]) ? String(sheet.getRange(row, getHeaders(sheet).colByName[OPS_COL.qcTeam]).getValue() || '').trim() : '';
    const instDate = (row && getHeaders(sheet).colByName[OPS_COL.instDate]) ? _normalizeDateStr(sheet.getRange(row, getHeaders(sheet).colByName[OPS_COL.instDate]).getValue()) : '';
    const team = _qcTeamByName(teamName);
    if (team && instDate) {
      const cal = _qcCal(team.calId);
      const evs = _qcFindJobEvents(cal, instDate, phone);
      evs.forEach(function(ev) {
        if (_opsFmtDate(ev.getStartTime()) === date) {
          try { ev.setColor(CalendarApp.EventColor.GREEN); } catch (_e) {}
          const t = ev.getTitle() || ''; if (t.indexOf('✓') === -1) { try { ev.setTitle('✓ ' + t); } catch (_e) {} }
        }
      });
    }
  } catch (_e) {}

  return jsonResponse({ status: 'ok', phone: phone, date: date });
}
