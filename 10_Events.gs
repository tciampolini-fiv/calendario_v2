function openEventSidebar() {
  const tpl = HtmlService.createTemplateFromFile('Sidebar');
  SpreadsheetApp.getUi().showSidebar(tpl.evaluate().setTitle('Dettaglio evento'));
}

function getCurrentEventPanelData() {
  const active = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!active || active.getName() !== APP.SHEETS.CALENDAR) {
    return { state: 'NO_EVENT', message: 'Seleziona una riga nel foglio Calendario.' };
  }

  const range = active.getActiveRange();
  const rowNumber = range ? range.getRow() : 0;
  if (rowNumber < 2) {
    return { state: 'NO_EVENT', message: 'Seleziona una riga evento.' };
  }

  const map = headerMap_(active);
  const values = active.getRange(rowNumber, 1, 1, active.getLastColumn()).getValues()[0];
  const event = {};
  Object.keys(map).forEach(h => event[h] = values[map[h] - 1]);
  event._row = rowNumber;

  const hasEventData = [
    APP.CALENDAR_HEADERS.START,
    APP.CALENDAR_HEADERS.TYPE,
    APP.CALENDAR_HEADERS.EVENT,
    APP.CALENDAR_HEADERS.CLASS,
    APP.CALENDAR_HEADERS.LOCATION
  ].some(h => String(event[h] || '').trim() !== '');

  if (!hasEventData) {
    return { state: 'NO_EVENT', message: 'La riga selezionata è vuota.' };
  }

  const id = ensureEventId_(event);
  const data = getEventPanelData(id);
  data.state = 'EVENT';
  data.row = rowNumber;
  return data;
}

function getEventPanelData(eventId) {
  const found = findCalendarEventById_(eventId);
  if (!found) throw new Error('Evento non trovato: ' + eventId);

  return clientSafe_({
    state: 'EVENT',
    eventId: String(eventId),
    row: found.row,
    event: found.event,
    refundLimit: getEventRefundLimit_(eventId, found.event),
    checklist: getChecklistForEvent_(eventId),
    expenses: getExpensesForEvent_(eventId),
    participants: getParticipantsForEvent_(eventId),
    cebOptions: getCebOptionsForEvent_(found.event)
  });
}

function findCalendarEventById_(eventId) {
  const cal = sh_(APP.SHEETS.CALENDAR);
  const rows = cal.getDataRange().getValues();
  if (!rows.length) return null;
  const headers = rows[0];
  const idCol = headers.indexOf(APP.CALENDAR_HEADERS.ID);
  if (idCol < 0) throw new Error('Colonna ID EVENTO non trovata nel Calendario.');

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idCol]) !== String(eventId)) continue;
    const event = {};
    headers.forEach((h, c) => event[h] = rows[i][c]);
    event._row = i + 1;
    return { row: i + 1, event: event };
  }
  return null;
}

function getEventRefundLimit_(eventId, event) {
  const meta = sh_(APP.SHEETS.EVENT_META).getDataRange().getValues();
  for (let i = 1; i < meta.length; i++) {
    if (String(meta[i][0]) === String(eventId)) {
      const value = Number(meta[i][1] || 0);
      const custom = meta[i][2] === true;
      if (custom || value > 0) {
        return { value: value, source: custom ? 'PERSONALIZZATO' : (meta[i][3] || 'STANDARD') };
      }
    }
  }
  return { value: resolveRefundStandard_(event), source: 'STANDARD' };
}

function resolveRefundStandard_(event) {
  const rows = sh_(APP.SHEETS.REFUND_CONFIG).getDataRange().getValues();
  const type = normalize_(event[APP.CALENDAR_HEADERS.TYPE]);
  const cls = normalize_(event[APP.CALENDAR_HEADERS.CLASS]);
  let best = null;

  for (let i = 1; i < rows.length; i++) {
    const active = rows[i][3] === true || normalize_(rows[i][3]) === 'SI';
    if (!active) continue;
    const rType = normalize_(rows[i][0]);
    const rClass = normalize_(rows[i][1]);
    if (rType !== type && rType !== 'ALTRO' && rType !== 'TUTTI') continue;
    if (rClass && rClass !== '*' && rClass !== cls) continue;
    const priority = Number(rows[i][4] || 0);
    if (!best || priority > best.priority) best = { value: Number(rows[i][2] || 0), priority: priority };
  }
  return best ? best.value : 0;
}

function setEventRefundLimit(eventId, value) {
  const sheet = sh_(APP.SHEETS.EVENT_META);
  const rows = sheet.getDataRange().getValues();
  const numeric = Number(value || 0);
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(eventId)) {
      sheet.getRange(i + 1, 2, 1, 3).setValues([[numeric, true, 'PERSONALIZZATO']]);
      sheet.getRange(i + 1, 8).setValue(new Date());
      return getEventPanelData(eventId);
    }
  }
  sheet.appendRow([eventId, numeric, true, 'PERSONALIZZATO', '', '', '', new Date()]);
  return getEventPanelData(eventId);
}

function resetEventRefundLimit(eventId) {
  const sheet = sh_(APP.SHEETS.EVENT_META);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(eventId)) {
      sheet.getRange(i + 1, 2, 1, 3).setValues([[0, false, '']]);
      sheet.getRange(i + 1, 8).setValue(new Date());
      break;
    }
  }
  return getEventPanelData(eventId);
}
