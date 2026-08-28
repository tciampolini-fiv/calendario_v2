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
  if (!profile || profile === 'NESSUNO') return 0;

  let added = 0;
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
    added++;
  });
  return added;
}

/**
 * Una tantum / manutenzione: completa la checklist standard degli eventi
 * in corso o futuri (data fine >= oggi), senza duplicare le attività già
 * importate o personalizzate. Gli eventi già conclusi non vengono toccati.
 */
function initializeCurrentAndFutureChecklists() {
  const cal = sh_(APP.SHEETS.CALENDAR);
  const rows = cal.getDataRange().getValues();
  if (rows.length < 2) return;

  const headers = rows[0];
  const idx = {};
  headers.forEach((h, i) => idx[String(h).trim()] = i);
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let eventsTouched = 0;
  let tasksAdded = 0;

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const eventId = String(row[idx[APP.CALENDAR_HEADERS.ID]] || '').trim();
    if (!eventId) continue;

    const end = row[idx[APP.CALENDAR_HEADERS.END]];
    if (end instanceof Date) {
      const endDay = new Date(end);
      endDay.setHours(0, 0, 0, 0);
      if (endDay < today) continue;
    }

    const event = {};
    headers.forEach((h, c) => event[h] = row[c]);
    event._row = i + 1;
    const added = generateChecklistForEvent_(eventId, event);
    if (added > 0) {
      eventsTouched++;
      tasksAdded += added;
    }
  }

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    'Checklist aggiornate',
    'Eventi completati/aggiornati: ' + eventsTouched + '\nNuove attività create: ' + tasksAdded,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
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
    refreshEventSummary_(eventId);
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
    refreshEventSummary_(eventId);
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
  refreshEventSummary_(eventId);
  return getEventPanelData(eventId);
}

function deleteChecklistItem(taskId) {
  const sheet = sh_(APP.SHEETS.CHECKLIST);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(taskId)) continue;
    const eventId = rows[i][1];
    sheet.deleteRow(i + 1);
    refreshEventSummary_(eventId);
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
    refreshEventSummary_(eventId);
    return;
  }
  if (!shouldBeOpen) return;
  const existing = getChecklistForEvent_(eventId);
  const maxOrder = existing.reduce((m, x) => Math.max(m, Number(x.order || 0)), 0);
  sheet.appendRow([
    'TASK-' + Utilities.getUuid(), eventId, maxOrder + 10, task, 'AMMINISTRAZIONE', dueDate || '',
    'DA FARE', 'ALTA', 'AUTO', autoKey, note || '', '', new Date()
  ]);
  refreshEventSummary_(eventId);
}

function refreshSelectedEventSummary() {
  SpreadsheetApp.flush();
}

/**
 * I riepiloghi M:O del Calendario sono formule collegate a _CHECKLIST.
 * Manteniamo questa funzione per compatibilità con il resto del codice,
 * ma non scriviamo più direttamente nelle celle del calendario.
 */
function refreshEventSummary_(eventId, row) {
  SpreadsheetApp.flush();
}
