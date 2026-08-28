function getChecklistForEvent_(eventId) {
  const rows = sh_(APP.SHEETS.CHECKLIST).getDataRange().getValues();
  return rows.slice(1).filter(r => String(r[1]) === String(eventId)).map(r => ({
    id: r[0], order: r[2], task: r[3], category: r[4], dueDate: r[5], status: r[6],
    priority: r[7], source: r[8], autoKey: r[9], note: r[10], completedAt: r[11], updatedAt: r[12]
  })).sort((a,b) => Number(a.order || 0) - Number(b.order || 0));
}

function getChecklistProfileForEvent_(event) {
  const type = normalize_(event[APP.CALENDAR_HEADERS.TYPE]);
  const rows = sh_(APP.SHEETS.EVENT_TYPE_CONFIG).getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (normalize_(rows[i][0]) === type) return normalize_(rows[i][3]) || type;
  }
  return type;
}

function generateChecklistForSelectedEvent() {
  const event = selectedEvent_();
  const eventId = ensureEventId_(event);
  generateChecklistForEvent_(eventId, event);
  refreshEventSummary_(eventId, event._row);
}

function generateChecklistForPanel(eventId) {
  const found = findCalendarEventById_(eventId);
  if (!found) throw new Error('Evento non trovato.');
  generateChecklistForEvent_(eventId, found.event);
  refreshEventSummary_(eventId, found.row);
  return getEventPanelData(eventId);
}

function generateChecklistForEvent_(eventId, event) {
  const target = sh_(APP.SHEETS.CHECKLIST);
  const existing = getChecklistForEvent_(eventId);
  const existingKeys = new Set(existing.map(x => normalize_(x.autoKey || x.task)));
  const config = sh_(APP.SHEETS.CHECKLIST_CONFIG).getDataRange().getValues();
  const profile = getChecklistProfileForEvent_(event);
  const cls = normalize_(event[APP.CALENDAR_HEADERS.CLASS]);
  const start = event[APP.CALENDAR_HEADERS.START];
  const end = event[APP.CALENDAR_HEADERS.END];
  if (!profile || profile === 'NESSUNO') return;

  config.slice(1).forEach((r, idx) => {
    const cfgProfile = normalize_(r[1]);
    const cfgClass = normalize_(r[2]);
    const active = r[10] === true || normalize_(r[10]) === 'SI';
    if (!active || (cfgProfile !== profile && cfgProfile !== 'TUTTI')) return;
    if (cfgClass && cfgClass !== '*' && cfgClass !== cls) return;

    const task = String(r[3] || '').trim();
    const autoKey = String(r[9] || task).trim();
    if (!task || existingKeys.has(normalize_(autoKey))) return;

    const base = normalize_(r[5]);
    const offsetDays = r[6] === '' ? null : Number(r[6] || 0);
    let baseDate = base === 'FINE' ? end : start;
    if (!(baseDate instanceof Date)) baseDate = null;
    const due = baseDate && offsetDays !== null ? new Date(baseDate.getTime() + offsetDays * 86400000) : '';

    target.appendRow([
      'TASK-' + Utilities.getUuid(), eventId, Number(r[8] || (idx + 1) * 10), task,
      r[4] || '', due, 'DA FARE', r[7] || 'NORMALE', 'STANDARD', autoKey,
      r[11] || '', '', new Date()
    ]);
    existingKeys.add(normalize_(autoKey));
  });
}

function setChecklistItemStatus(taskId, completed) {
  const sheet = sh_(APP.SHEETS.CHECKLIST);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(taskId)) continue;
    const eventId = rows[i][1];
    sheet.getRange(i + 1, 7).setValue(completed ? 'FATTO' : 'DA FARE');
    sheet.getRange(i + 1, 12).setValue(completed ? new Date() : '');
    sheet.getRange(i + 1, 13).setValue(new Date());
    const found = findCalendarEventById_(eventId);
    if (found) refreshEventSummary_(eventId, found.row);
    return getEventPanelData(eventId);
  }
  throw new Error('Attività non trovata.');
}

function updateChecklistItem(taskId, dueDate, note) {
  const sheet = sh_(APP.SHEETS.CHECKLIST);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(taskId)) continue;
    const eventId = rows[i][1];
    sheet.getRange(i + 1, 6).setValue(parseClientDate_(dueDate) || '');
    sheet.getRange(i + 1, 11).setValue(note || '');
    sheet.getRange(i + 1, 13).setValue(new Date());
    const found = findCalendarEventById_(eventId);
    if (found) refreshEventSummary_(eventId, found.row);
    return getEventPanelData(eventId);
  }
  throw new Error('Attività non trovata.');
}

function addChecklistItem(eventId, task, dueDate, priority, note) {
  task = String(task || '').trim();
  if (!task) throw new Error('Inserisci il nome dell’attività.');
  const sheet = sh_(APP.SHEETS.CHECKLIST);
  const existing = getChecklistForEvent_(eventId);
  const maxOrder = existing.reduce((m, x) => Math.max(m, Number(x.order || 0)), 0);
  sheet.appendRow([
    'TASK-' + Utilities.getUuid(), eventId, maxOrder + 10, task, 'PERSONALIZZATA',
    parseClientDate_(dueDate) || '', 'DA FARE', priority || 'NORMALE', 'PERSONALIZZATA', '', note || '', '', new Date()
  ]);
  const found = findCalendarEventById_(eventId);
  if (found) refreshEventSummary_(eventId, found.row);
  return getEventPanelData(eventId);
}

function deleteChecklistItem(taskId) {
  const sheet = sh_(APP.SHEETS.CHECKLIST);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(taskId)) continue;
    const eventId = rows[i][1];
    sheet.deleteRow(i + 1);
    const found = findCalendarEventById_(eventId);
    if (found) refreshEventSummary_(eventId, found.row);
    return getEventPanelData(eventId);
  }
  throw new Error('Attività non trovata.');
}

function syncAutoChecklistTask_(eventId, autoKey, task, dueDate, shouldBeOpen, note) {
  const sheet = sh_(APP.SHEETS.CHECKLIST);
  const rows = sheet.getDataRange().getValues();
  const key = normalize_(autoKey);
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) !== String(eventId) || normalize_(rows[i][9]) !== key) continue;
    sheet.getRange(i + 1, 4).setValue(task);
    if (dueDate !== undefined) sheet.getRange(i + 1, 6).setValue(dueDate || '');
    sheet.getRange(i + 1, 7).setValue(shouldBeOpen ? 'DA FARE' : 'FATTO');
    sheet.getRange(i + 1, 11).setValue(note || '');
    sheet.getRange(i + 1, 12).setValue(shouldBeOpen ? '' : new Date());
    sheet.getRange(i + 1, 13).setValue(new Date());
    return;
  }
  if (!shouldBeOpen) return;
  const existing = getChecklistForEvent_(eventId);
  const maxOrder = existing.reduce((m, x) => Math.max(m, Number(x.order || 0)), 0);
  sheet.appendRow([
    'TASK-' + Utilities.getUuid(), eventId, maxOrder + 10, task, 'AMMINISTRAZIONE', dueDate || '',
    'DA FARE', 'ALTA', 'AUTO', autoKey, note || '', '', new Date()
  ]);
}

function refreshSelectedEventSummary() {
  const event = selectedEvent_();
  const id = ensureEventId_(event);
  refreshEventSummary_(id, event._row);
}

function refreshEventSummary_(eventId, row) {
  const items = getChecklistForEvent_(eventId);
  const cal = sh_(APP.SHEETS.CALENDAR);
  const map = headerMap_(cal);
  const isDone = x => ['COMPLETATA', 'FATTO'].includes(normalize_(x.status));
  const done = items.filter(isDone).length;
  const open = items.filter(x => !isDone(x));
  const now = new Date();
  const late = open.filter(x => x.dueDate instanceof Date && x.dueDate < now);
  let badge = '⚪ ' + done + '/' + items.length;
  if (items.length && open.length === 0) badge = '🟢 ' + done + '/' + items.length + ' completate';
  else if (late.length) badge = '🔴 ' + done + '/' + items.length + ' · ' + late.length + ' scadute';
  else if (open.length) badge = '🟠 ' + done + '/' + items.length + ' · ' + open.length + ' da fare';

  const next = open.sort((a,b) => {
    const ad = a.dueDate instanceof Date ? a.dueDate.getTime() : 9e15;
    const bd = b.dueDate instanceof Date ? b.dueDate.getTime() : 9e15;
    return ad - bd;
  })[0];

  cal.getRange(row, map[APP.CALENDAR_HEADERS.CHECKLIST]).setValue(badge);
  cal.getRange(row, map[APP.CALENDAR_HEADERS.NEXT_ACTION]).setValue(next ? next.task : '');
  cal.getRange(row, map[APP.CALENDAR_HEADERS.DEADLINE]).setValue(next && next.dueDate instanceof Date ? next.dueDate : '');
}
