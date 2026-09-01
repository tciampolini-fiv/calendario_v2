const EVENT_SHEET = Object.freeze({
  SHEETS: Object.freeze({
    TASKS: 'Attività',
    EXPENSES: 'Spese',
    GUESTS: 'Invitati',
    META: '_META'
  }),
  TASK_HEADERS: Object.freeze([
    'ID TASK','STATO','ATTIVITA','CATEGORIA','SCADENZA','NOTE','ORIGINE','AUTO KEY','ORDINE','DATA COMPLETAMENTO','ULTIMO AGGIORNAMENTO'
  ]),
  EXPENSE_HEADERS: Object.freeze([
    'ID SPESA','ID EVENTO','TIPO RECORD','CATEGORIA','CEB','DESCRIZIONE','BENEFICIARIO','ID PERSONA',
    'IMPORTO PREVENTIVO','IMPORTO EFFETTIVO','SCADENZA PAGAMENTO','STATO PAGAMENTO','DATA PAGAMENTO',
    'RIFERIMENTO','ALLEGATO / LINK','NOTE','DATA INSERIMENTO','ULTIMO AGGIORNAMENTO','STATO RIF',
    'DATA RICHIESTA RIF','CODICE RIF','FATTURA RICEVUTA','DATA FATTURA','CONTABILE RICHIESTA',
    'CONTABILE INVIATA','DATA CHIUSURA'
  ]),
  GUEST_HEADERS: Object.freeze([
    'ID PARTECIPAZIONE','ID EVENTO','ID PERSONA','NOME','COGNOME','CIRCOLO','RUOLO','EMAIL OVERRIDE',
    'MAX RIMBORSO PERSONALE','PROVENIENZA','FILE ORIGINE','NOTE'
  ])
});

function prepareEventSheetForSelectedEvent() {
  const event = selectedEvent_();
  const eventId = ensureEventId_(event);
  generateChecklistForEvent_(eventId, event);
  const folder = createWorkFolderForEvent_(eventId, event, event._row);
  const result = getOrCreateEventSheet_(eventId, event, folder.folderId);
  exportEventDataToEventSheet_(eventId, result.spreadsheet);
  setEventSheetLink_(event._row, result.spreadsheet.getUrl());
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    result.created ? 'Scheda evento creata' : 'Scheda evento aggiornata',
    'La scheda è nella cartella evento. Aprila dalla colonna SCHEDA EVENTO.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return { created: result.created, id: result.spreadsheet.getId(), url: result.spreadsheet.getUrl() };
}

function syncSelectedEventSheetToCalendar() {
  const event = selectedEvent_();
  const eventId = ensureEventId_(event);
  const folder = createWorkFolderForEvent_(eventId, event, event._row);
  const result = getOrCreateEventSheet_(eventId, event, folder.folderId);
  const child = result.spreadsheet;
  validateEventSheetIdentity_(child, eventId);

  const counts = {
    tasks: syncTasksFromEventSheet_(eventId, child),
    expenses: syncExpensesFromEventSheet_(eventId, child),
    guests: syncGuestsFromEventSheet_(eventId, child)
  };

  // Le regole temporali vengono valutate solo quando l'utente chiede l'aggiornamento.
  // Nessun timer e nessun polling: IN ATTESA diventa DA FARE quando serve un sollecito.
  syncChecklistLocksForEvent_(eventId);
  getExpensesForEvent_(eventId).forEach(syncExpenseTasks_);

  // Riporta nella scheda eventuali ID assegnati e gli stati calcolati dal sistema.
  exportEventDataToEventSheet_(eventId, child);
  setEventSheetLink_(event._row, child.getUrl());
  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    'Calendario aggiornato',
    'Attività: ' + counts.tasks + '\nSpese: ' + counts.expenses + '\nInvitati: ' + counts.guests,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return counts;
}

function getOrCreateEventSheet_(eventId, event, folderId) {
  const existing = findEventSheet_(eventId, event, folderId);
  if (existing) {
    ensureEventSheetStructure_(existing, eventId, folderId, event);
    return { spreadsheet: existing, created: false };
  }

  const name = buildEventSheetName_(event);
  const child = SpreadsheetApp.create(name);
  child.setSpreadsheetLocale('it_IT');
  child.setSpreadsheetTimeZone(APP.TZ);
  DriveApp.getFileById(child.getId()).moveTo(DriveApp.getFolderById(folderId));
  ensureEventSheetStructure_(child, eventId, folderId, event);
  return { spreadsheet: child, created: true };
}

function findEventSheet_(eventId, event, folderId) {
  const cal = sh_(APP.SHEETS.CALENDAR);
  const map = headerMap_(cal);
  const linkCol = map[APP.CALENDAR_HEADERS.EVENT_SHEET];
  if (linkCol && event && event._row) {
    const rich = cal.getRange(event._row, linkCol).getRichTextValue();
    const link = rich && rich.getLinkUrl();
    const displayed = cal.getRange(event._row, linkCol).getDisplayValue();
    const id = extractDriveId_(link || displayed);
    if (id) {
      try { return SpreadsheetApp.openById(id); } catch (e) {}
    }
  }

  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (files.hasNext()) {
    const file = files.next();
    try {
      const ss = SpreadsheetApp.openById(file.getId());
      const meta = ss.getSheetByName(EVENT_SHEET.SHEETS.META);
      if (meta && String(readMetaValue_(meta, 'EVENT_ID')) === String(eventId)) return ss;
    } catch (e) {}
  }
  return null;
}

function ensureEventSheetStructure_(child, eventId, folderId, event) {
  child.setSpreadsheetLocale('it_IT');
  child.setSpreadsheetTimeZone(APP.TZ);

  let tasks = child.getSheetByName(EVENT_SHEET.SHEETS.TASKS);
  if (!tasks) {
    const first = child.getSheets()[0];
    if (child.getSheets().length === 1 && first.getLastRow() === 0) {
      first.setName(EVENT_SHEET.SHEETS.TASKS);
      tasks = first;
    } else tasks = child.insertSheet(EVENT_SHEET.SHEETS.TASKS, 0);
  }
  let expenses = child.getSheetByName(EVENT_SHEET.SHEETS.EXPENSES);
  if (!expenses) expenses = child.insertSheet(EVENT_SHEET.SHEETS.EXPENSES);
  let guests = child.getSheetByName(EVENT_SHEET.SHEETS.GUESTS);
  if (!guests) guests = child.insertSheet(EVENT_SHEET.SHEETS.GUESTS);
  let meta = child.getSheetByName(EVENT_SHEET.SHEETS.META);
  if (!meta) meta = child.insertSheet(EVENT_SHEET.SHEETS.META);

  prepareChildSheet_(tasks, EVENT_SHEET.TASK_HEADERS, 'TASKS');
  prepareChildSheet_(expenses, EVENT_SHEET.EXPENSE_HEADERS, 'EXPENSES');
  prepareChildSheet_(guests, EVENT_SHEET.GUEST_HEADERS, 'GUESTS');

  writeMeta_(meta, {
    EVENT_ID: eventId,
    MASTER_SPREADSHEET_ID: APP.SPREADSHEET_ID,
    EVENT_FOLDER_ID: folderId,
    EVENT_SHEET_ID: child.getId(),
    SYNC_VERSION: '1',
    EVENT_LABEL: buildEventSheetLabel_(event)
  });
  meta.hideSheet();
}

function prepareChildSheet_(sheet, headers, kind) {
  if (sheet.getMaxColumns() < headers.length) sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#eeeeee')
    .setVerticalAlignment('middle');

  if (kind === 'TASKS') {
    sheet.hideColumns(1);
    sheet.hideColumns(7, 5);
    sheet.setColumnWidth(2, 115);
    sheet.setColumnWidth(3, 260);
    sheet.setColumnWidth(4, 150);
    sheet.setColumnWidth(5, 105);
    sheet.setColumnWidth(6, 320);
    sheet.getRange('B2:B1000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['DA FARE','IN ATTESA','FATTO'], true).setAllowInvalid(false).build());
    sheet.getRange('E2:E1000').setNumberFormat('dd/MM/yyyy');
    const rules = [
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($B2="IN ATTESA";$E2<>"";$E2<TODAY())').setBackground('#f4cccc').setRanges([sheet.getRange('B2:F1000')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('FATTO').setBackground('#d9ead3').setRanges([sheet.getRange('B2:B1000')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('IN ATTESA').setBackground('#fff2cc').setRanges([sheet.getRange('B2:B1000')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($B2="DA FARE";$E2<>"";$E2<TODAY())').setBackground('#f4cccc').setRanges([sheet.getRange('B2:F1000')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('DA FARE').setBackground('#fce5cd').setRanges([sheet.getRange('B2:B1000')]).build()
    ];
    sheet.setConditionalFormatRules(rules);
  }

  if (kind === 'EXPENSES') {
    sheet.hideColumns(1, 2);
    sheet.hideColumns(8);
    sheet.setColumnWidth(6, 220);
    sheet.setColumnWidth(7, 170);
    sheet.setColumnWidth(16, 260);
    sheet.getRange('I2:J1000').setNumberFormat('€ #,##0.00');
    ['K','M','Q','R','T','W','Z'].forEach(c => sheet.getRange(c + '2:' + c + '1000').setNumberFormat('dd/MM/yyyy'));
    sheet.getRange('L2:L1000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['DA DEFINIRE','DA PAGARE','PAGATO - FATTURA','PAGATO - AFOR','PAGATO CON CC','CHIUSO'], true).setAllowInvalid(false).build());
    ['V','X','Y'].forEach(c => sheet.getRange(c + '2:' + c + '1000').insertCheckboxes());
  }

  if (kind === 'GUESTS') {
    sheet.hideColumns(1, 3);
    sheet.hideColumns(10, 2);
    sheet.setColumnWidth(4, 140);
    sheet.setColumnWidth(5, 140);
    sheet.setColumnWidth(6, 220);
    sheet.setColumnWidth(8, 230);
    sheet.setColumnWidth(12, 260);
    sheet.getRange('G2:G1000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['ATLETA','TECNICO','ALTRO'], true).setAllowInvalid(false).build());
  }
}

function exportEventDataToEventSheet_(eventId, child) {
  const tasks = getChecklistRowsForEventRaw_(eventId).map(x => [
    x.id,
    normalizeTaskStatusForSheet_(x.status),
    x.task || '',
    x.category || '',
    x.dueDate || '',
    x.note || '',
    x.source || '',
    x.autoKey || '',
    x.order || '',
    x.completedAt || '',
    x.updatedAt || ''
  ]);

  const expenses = sh_(APP.SHEETS.EXPENSES).getDataRange().getValues().slice(1)
    .filter(r => String(r[1]) === String(eventId))
    .map(r => r.slice(0, EVENT_SHEET.EXPENSE_HEADERS.length));

  const guests = sh_(APP.SHEETS.PARTICIPANTS).getDataRange().getValues().slice(1)
    .filter(r => String(r[1]) === String(eventId))
    .map(r => r.slice(0, EVENT_SHEET.GUEST_HEADERS.length));

  replaceChildData_(child.getSheetByName(EVENT_SHEET.SHEETS.TASKS), tasks, EVENT_SHEET.TASK_HEADERS.length);
  replaceChildData_(child.getSheetByName(EVENT_SHEET.SHEETS.EXPENSES), expenses, EVENT_SHEET.EXPENSE_HEADERS.length);
  replaceChildData_(child.getSheetByName(EVENT_SHEET.SHEETS.GUESTS), guests, EVENT_SHEET.GUEST_HEADERS.length);

  const meta = child.getSheetByName(EVENT_SHEET.SHEETS.META);
  writeMeta_(meta, { LAST_SYNC: Utilities.formatDate(new Date(), APP.TZ, 'dd/MM/yyyy HH:mm') });
}

function replaceChildData_(sheet, rows, width) {
  const last = sheet.getLastRow();
  if (last > 1) sheet.getRange(2, 1, last - 1, width).clearContent();
  if (rows.length) sheet.getRange(2, 1, rows.length, width).setValues(rows);
}

function syncTasksFromEventSheet_(eventId, child) {
  const sheet = child.getSheetByName(EVENT_SHEET.SHEETS.TASKS);
  const values = sheet.getDataRange().getValues();
  const now = new Date();
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    const task = String(r[2] || '').trim();
    if (!task) continue;
    let id = String(r[0] || '').trim();
    if (!id) id = 'TASK-' + Utilities.getUuid();
    let status = normalizeTaskStatusForSheet_(r[1] || 'DA FARE');
    const source = String(r[6] || '').trim() || 'PERSONALIZZATA';
    let completed = r[9] || '';
    if (status === 'FATTO' && !(completed instanceof Date)) completed = now;
    if (status !== 'FATTO') completed = '';
    rows.push([
      id, eventId, Number(r[8] || (i * 10)), task, r[3] || '', r[4] || '', status,
      '', source, r[7] || '', r[5] || '', completed, now
    ]);
  }
  replaceCentralRowsForEvent_(sh_(APP.SHEETS.CHECKLIST), eventId, 2, rows, 13);
  return rows.length;
}

function syncExpensesFromEventSheet_(eventId, child) {
  const sheet = child.getSheetByName(EVENT_SHEET.SHEETS.EXPENSES);
  const values = sheet.getDataRange().getValues();
  const now = new Date();
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i].slice(0, 26);
    const hasData = r.slice(2).some(v => String(v === false ? '' : v || '').trim() !== '');
    if (!hasData) continue;
    if (!r[0]) r[0] = 'SPESA-' + Utilities.getUuid();
    r[1] = eventId;
    if (!r[16]) r[16] = now;
    r[17] = now;
    while (r.length < 26) r.push('');
    rows.push(r);
  }
  replaceCentralRowsForEvent_(sh_(APP.SHEETS.EXPENSES), eventId, 2, rows, 26);
  return rows.length;
}

function syncGuestsFromEventSheet_(eventId, child) {
  const sheet = child.getSheetByName(EVENT_SHEET.SHEETS.GUESTS);
  const values = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i].slice(0, 12);
    const name = String(r[3] || '').trim();
    const surname = String(r[4] || '').trim();
    if (!name && !surname) continue;
    if (!r[0]) r[0] = 'PAR-' + Utilities.getUuid();
    r[1] = eventId;
    if (!r[2]) {
      const known = findKnownPerson_(name, surname, '');
      if (known) {
        r[2] = known.id;
        if (!r[5]) r[5] = known.club || '';
        if (!r[7]) r[7] = known.email || '';
      }
    }
    if (!r[9]) r[9] = 'SCHEDA EVENTO';
    while (r.length < 12) r.push('');
    rows.push(r);
  }
  replaceCentralRowsForEvent_(sh_(APP.SHEETS.PARTICIPANTS), eventId, 2, rows, 12);
  return rows.length;
}

function replaceCentralRowsForEvent_(sheet, eventId, eventColumn, newRows, width) {
  const values = sheet.getDataRange().getValues();
  const targetRows = [];
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][eventColumn - 1]) === String(eventId)) targetRows.push(i + 1);
  }

  const reused = Math.min(targetRows.length, newRows.length);
  for (let i = 0; i < reused; i++) sheet.getRange(targetRows[i], 1, 1, width).setValues([newRows[i]]);
  for (let i = reused; i < targetRows.length; i++) sheet.getRange(targetRows[i], 1, 1, width).clearContent();
  if (newRows.length > reused) {
    const remaining = newRows.slice(reused);
    sheet.getRange(sheet.getLastRow() + 1, 1, remaining.length, width).setValues(remaining);
  }
}

function validateEventSheetIdentity_(child, eventId) {
  const meta = child.getSheetByName(EVENT_SHEET.SHEETS.META);
  if (!meta) throw new Error('La Scheda evento non contiene il foglio tecnico _META.');
  const stored = String(readMetaValue_(meta, 'EVENT_ID') || '').trim();
  if (stored && stored !== String(eventId)) throw new Error('La Scheda evento appartiene a un altro evento: ' + stored);
}

function readMetaValue_(sheet, key) {
  const values = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), 2).getValues();
  const target = normalize_(key);
  for (let i = 0; i < values.length; i++) if (normalize_(values[i][0]) === target) return values[i][1];
  return '';
}

function writeMeta_(sheet, object) {
  const values = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), 2).getValues();
  const rowsByKey = {};
  values.forEach((r, i) => { if (r[0]) rowsByKey[normalize_(r[0])] = i + 1; });
  if (!values[0] || normalize_(values[0][0]) !== 'CHIAVE') {
    sheet.getRange(1, 1, 1, 2).setValues([['CHIAVE','VALORE']]).setFontWeight('bold').setBackground('#eeeeee');
    rowsByKey.CHIAVE = 1;
  }
  Object.keys(object).forEach(key => {
    const normalized = normalize_(key);
    const row = rowsByKey[normalized] || sheet.getLastRow() + 1;
    sheet.getRange(row, 1, 1, 2).setValues([[key, object[key]]]);
    rowsByKey[normalized] = row;
  });
}

function normalizeTaskStatusForSheet_(status) {
  const s = normalize_(status);
  if (s === 'FATTO' || s === 'COMPLETATA') return 'FATTO';
  if (s === 'BLOCCATA' || s === 'IN ATTESA') return 'IN ATTESA';
  return 'DA FARE';
}

function buildEventSheetName_(event) {
  const type = toTitleCaseV2_(event[APP.CALENDAR_HEADERS.TYPE] || 'Evento');
  const cls = String(event[APP.CALENDAR_HEADERS.CLASS] || '').trim();
  const place = String(event[APP.CALENDAR_HEADERS.LOCATION] || '').trim();
  const dates = formatDocumentDateRange_(event[APP.CALENDAR_HEADERS.START], event[APP.CALENDAR_HEADERS.END]);
  return ['Scheda evento', type, cls, place, dates].filter(Boolean).join(' - ');
}

function buildEventSheetLabel_(event) {
  return [
    event[APP.CALENDAR_HEADERS.TYPE] || '',
    event[APP.CALENDAR_HEADERS.CLASS] || '',
    event[APP.CALENDAR_HEADERS.LOCATION] || '',
    formatDocumentDateRange_(event[APP.CALENDAR_HEADERS.START], event[APP.CALENDAR_HEADERS.END])
  ].filter(Boolean).join(' | ');
}

function setEventSheetLink_(row, url) {
  const sheet = sh_(APP.SHEETS.CALENDAR);
  const map = headerMap_(sheet);
  let col = map[APP.CALENDAR_HEADERS.EVENT_SHEET];
  if (!col) {
    col = sheet.getLastColumn() + 1;
    sheet.getRange(1, col).setValue(APP.CALENDAR_HEADERS.EVENT_SHEET);
  }
  const rich = SpreadsheetApp.newRichTextValue().setText('↗ APRI').setLinkUrl(url).build();
  sheet.getRange(row, col).setRichTextValue(rich);
}
