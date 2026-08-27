function openEventSidebar() {
  const event = selectedEvent_();
  const id = ensureEventId_(event);
  const tpl = HtmlService.createTemplateFromFile('Sidebar');
  tpl.eventId = id;
  SpreadsheetApp.getUi().showSidebar(tpl.evaluate().setTitle('Dettaglio evento'));
}

function getEventPanelData(eventId) {
  const cal = sh_(APP.SHEETS.CALENDAR);
  const rows = cal.getDataRange().getValues();
  const headers = rows[0];
  const idCol = headers.indexOf(APP.CALENDAR_HEADERS.ID);
  if (idCol < 0) throw new Error('Colonna ID EVENTO non trovata nel Calendario.');

  const row = rows.slice(1).find(r => String(r[idCol]) === String(eventId));
  if (!row) throw new Error('Evento non trovato: ' + eventId);

  const event = {};
  headers.forEach((h, i) => event[h] = row[i]);

  return clientSafe_({
    event: event,
    refundLimit: getEventRefundLimit_(eventId, event),
    checklist: getChecklistForEvent_(eventId),
    expenses: getExpensesForEvent_(eventId),
    participants: getParticipantsForEvent_(eventId)
  });
}

function getEventRefundLimit_(eventId, event) {
  const meta = sh_(APP.SHEETS.EVENT_META).getDataRange().getValues();
  for (let i = 1; i < meta.length; i++) {
    if (String(meta[i][0]) === String(eventId)) {
      const value = Number(meta[i][1] || 0);
      const custom = meta[i][2] === true;
      if (custom || value > 0) {
        return {
          value: value,
          source: custom ? 'PERSONALIZZATO' : (meta[i][3] || 'STANDARD')
        };
      }
    }
  }
  const standard = resolveRefundStandard_(event);
  return { value: standard, source: 'STANDARD' };
}

function resolveRefundStandard_(event) {
  const rows = sh_(APP.SHEETS.REFUND_CONFIG).getDataRange().getValues();
  const type = normalize_(event[APP.CALENDAR_HEADERS.TYPE]);
  const cls = normalize_(event[APP.CALENDAR_HEADERS.CLASS]);
  let best = null;

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][3] !== true) continue;
    const rType = normalize_(rows[i][0]);
    const rClass = normalize_(rows[i][1]);
    if (rType !== type && rType !== 'ALTRO') continue;
    if (rClass !== '*' && rClass !== cls) continue;
    const priority = Number(rows[i][4] || 0);
    if (!best || priority > best.priority) {
      best = { value: Number(rows[i][2] || 0), priority: priority };
    }
  }
  return best ? best.value : 0;
}

function setEventRefundLimit(eventId, value) {
  const sheet = sh_(APP.SHEETS.EVENT_META);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(eventId)) {
      sheet.getRange(i + 1, 2, 1, 3).setValues([[Number(value || 0), true, 'PERSONALIZZATO']]);
      return;
    }
  }
  sheet.appendRow([eventId, Number(value || 0), true, 'PERSONALIZZATO', '', '', '', new Date()]);
}
