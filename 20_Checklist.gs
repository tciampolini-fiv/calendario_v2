function getChecklistForEvent_(eventId) {
  const rows = sh_(APP.SHEETS.CHECKLIST).getDataRange().getValues();
  return rows.slice(1).filter(r => String(r[1]) === String(eventId)).map(r => ({
    id: r[0], task: r[3], dueDate: r[6], status: r[7], note: r[9], source: r[10]
  }));
}

function generateChecklistForSelectedEvent() {
  const event = selectedEvent_();
  const eventId = ensureEventId_(event);
  generateChecklistForEvent_(eventId, event);
  refreshEventSummary_(eventId, event._row);
}

function generateChecklistForEvent_(eventId, event) {
  const target = sh_(APP.SHEETS.CHECKLIST);
  const existing = getChecklistForEvent_(eventId);
  const existingKeys = new Set(existing.map(x => normalize_(x.task)));
  const config = sh_(APP.SHEETS.CHECKLIST_CONFIG).getDataRange().getValues();
  const type = normalize_(event[APP.CALENDAR_HEADERS.TYPE]);
  const start = event[APP.CALENDAR_HEADERS.START];
  config.slice(1).forEach((r, idx) => {
    const cfgType = normalize_(r[0]);
    if (cfgType !== type && cfgType !== 'TUTTI') return;
    const task = String(r[2] || '').trim();
    if (!task || existingKeys.has(normalize_(task))) return;
    const offsetDays = Number(r[5] || 0);
    const due = start instanceof Date ? new Date(start.getTime() + offsetDays * 86400000) : '';
    target.appendRow([
      'CHK-' + eventId + '-' + (idx + 1), eventId, r[1] || '', task, r[3] || '', r[4] || '', due,
      'DA FARE', false, '', 'STANDARD', new Date(), new Date()
    ]);
  });
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
  const done = items.filter(x => normalize_(x.status) === 'COMPLETATA' || normalize_(x.status) === 'FATTO').length;
  const open = items.filter(x => !(normalize_(x.status) === 'COMPLETATA' || normalize_(x.status) === 'FATTO'));
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
