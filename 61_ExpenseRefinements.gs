const EVENT_SHEET_TEMPLATE_ID_V3 = '1obNkubjpGhC7o8AX5Tc986UBGdyJs26b476b-GZN3Sc';

function prepareEventSheetForSelectedEventV2() {
  const event = selectedEvent_();
  const eventId = ensureEventId_(event);
  const folder = createWorkFolderForEvent_(eventId, event, event._row);
  const result = getOrCreateEventSheetV3_(eventId, event, folder.folderId);
  const child = result.spreadsheet;

  ensureChecklistBackendHeadersV3_();
  ensureParticipantsBackendHeadersV2_();

  // Il trigger viene creato dal Calendario prima che l'utente apra la nuova scheda.
  // In questo modo l'autonumerazione task e lo stato DA FARE funzionano subito,
  // anche se il menu del progetto bound impiega alcuni secondi a comparire.
  ensureTaskEditTriggerV4_(child);

  // Una nuova Scheda evento deve essere una copia FEDELE del modello.
  // Non ricostruiamo layout, formati, tabelle, menu o convalide subito dopo la copia.
  // In questo modo ogni modifica fatta al MODELLO - Scheda evento viene ereditata davvero.
  if (!result.created) {
    refreshTasksV4FromBackend_(eventId,event,child);
    refreshParticipantsV3FromBackend_(eventId,event,child);
    refreshExpensesV4FromBackend_(eventId,event,child,folder.folderId);
    hideEventSheetTechnicalColumnsV3_(child);
  }

  setEventSheetLink_(event._row, child.getUrl());
  writeEventMetaV5_(child,eventId,event,folder.folderId);
  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    result.created ? 'Scheda evento creata dal modello' : 'Scheda evento aggiornata',
    result.created
      ? 'È stata creata una nuova copia del MODELLO - Scheda evento nella cartella dell evento.'
      : 'La scheda evento collegata è stata aggiornata con i dati del Calendario.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return {created:result.created,id:child.getId(),url:child.getUrl()};
}

function getLinkedEventSheetForSaveV2_(event) {
  if (!event || !event._row) throw new Error('Seleziona una riga evento valida nel Calendario.');
  const cal = sh_(APP.SHEETS.CALENDAR);
  const map = headerMap_(cal);
  const col = map[APP.CALENDAR_HEADERS.EVENT_SHEET];
  if (!col) throw new Error('Colonna SCHEDA EVENTO non trovata nel Calendario.');

  const cell = cal.getRange(event._row,col);
  const rich = cell.getRichTextValue();
  const url = (rich && rich.getLinkUrl()) || cell.getDisplayValue();
  const id = extractDriveId_(url);
  if (!id) {
    throw new Error('Questa riga non ha una Scheda evento collegata. Usa prima “Crea / aggiorna scheda ← Calendario”.');
  }
  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    throw new Error('Non riesco ad aprire la Scheda evento collegata. Verifica il link nella colonna SCHEDA EVENTO.');
  }
}

function getExplicitLinkedEventSheetV3_(event) {
  if (!event || !event._row) return null;
  const cal = sh_(APP.SHEETS.CALENDAR);
  const map = headerMap_(cal);
  const col = map[APP.CALENDAR_HEADERS.EVENT_SHEET];
  if (!col) return null;

  const cell = cal.getRange(event._row,col);
  const rich = cell.getRichTextValue();
  const url = (rich && rich.getLinkUrl()) || cell.getDisplayValue();
  const id = extractDriveId_(url);
  if (!id) return null;

  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    // Link non più valido: lo consideriamo assente e verrà creata una nuova copia del modello.
    return null;
  }
}

function syncSelectedEventSheetToCalendarV2() {
  const ui = SpreadsheetApp.getUi();
  const event = selectedEvent_();
  const eventId = ensureEventId_(event);

  const child = getLinkedEventSheetForSaveV2_(event);
  validateEventSheetIdentity_(child,eventId);

  ensureChecklistBackendHeadersV3_();
  ensureParticipantsBackendHeadersV2_();

  const participantCount = syncParticipantsV3ToBackend_(eventId,event,child);
  const taskCount = syncTasksV4ToBackend_(eventId,event,child);
  const expenseCount = syncExpensesV4ToBackend_(eventId,event,child);

  SpreadsheetApp.flush();

  const counts = {
    tasks:taskCount,
    participants:participantCount,
    expenses:expenseCount
  };

  ui.alert(
    'Calendario aggiornato',
    'I dati presenti nella Scheda evento sono stati salvati nel Calendario.\n\n' +
    'Attività: ' + counts.tasks + '\n' +
    'Spese: ' + counts.expenses + '\n' +
    'Partecipanti: ' + counts.participants + '\n\n' +
    'La Scheda evento non è stata ricaricata dal Calendario.',
    ui.ButtonSet.OK
  );
  return counts;
}

function writeEventMetaV5_(child,eventId,event,folderId) {
  const meta = child.getSheetByName(EVENT_SHEET.SHEETS.META);
  const start = event[APP.CALENDAR_HEADERS.START];
  const end = event[APP.CALENDAR_HEADERS.END];
  writeMeta_(meta,{
    EVENT_ID:eventId,
    MASTER_SPREADSHEET_ID:APP.SPREADSHEET_ID,
    EVENT_FOLDER_ID:folderId,
    EVENT_SHEET_ID:child.getId(),
    SYNC_VERSION:'5',
    EVENT_LABEL:buildEventSheetLabel_(event),
    EVENT_TYPE:String(event[APP.CALENDAR_HEADERS.TYPE]||''),
    EVENT_CLASS:String(event[APP.CALENDAR_HEADERS.CLASS]||''),
    EVENT_LOCATION:String(event[APP.CALENDAR_HEADERS.LOCATION]||''),
    EVENT_ZONE:String(event[APP.CALENDAR_HEADERS.ZONE]||''),
    EVENT_CLUB:String(event[APP.CALENDAR_HEADERS.CLUB]||''),
    EVENT_TECHNICIANS:String(event[APP.CALENDAR_HEADERS.TECHNICIANS]||''),
    EVENT_LODGING:String(event[APP.CALENDAR_HEADERS.LODGING]||''),
    EVENT_COMMITMENT:String(event[APP.CALENDAR_HEADERS.COMMITMENT]||''),
    EVENT_START:start instanceof Date ? Utilities.formatDate(start,APP.TZ,'yyyy-MM-dd') : '',
    EVENT_END:end instanceof Date ? Utilities.formatDate(end,APP.TZ,'yyyy-MM-dd') : '',
    LAST_SYNC:Utilities.formatDate(new Date(),APP.TZ,'dd/MM/yyyy HH:mm')
  });
}

function getOrCreateEventSheetV3_(eventId,event,folderId) {
  // Riutilizziamo SOLO la Scheda esplicitamente collegata nella colonna SCHEDA EVENTO.
  // Non cerchiamo più file orfani nella cartella evento: se il link è vuoto, la scheda è nuova.
  let child = getExplicitLinkedEventSheetV3_(event);
  let created = false;
  if (!child) {
    const folder = DriveApp.getFolderById(folderId);
    const templateFile = DriveApp.getFileById(EVENT_SHEET_TEMPLATE_ID_V3);
    const copy = templateFile.makeCopy(buildEventSheetName_(event), folder);
    child = SpreadsheetApp.openById(copy.getId());
    child.setSpreadsheetLocale('it_IT');
    child.setSpreadsheetTimeZone(APP.TZ);
    created = true;
  }
  ensureEventSheetBaseV3_(child,eventId,folderId,event);
  return {spreadsheet:child,created:created};
}

function ensureEventSheetBaseV3_(child,eventId,folderId,event) {
  child.setSpreadsheetLocale('it_IT');
  child.setSpreadsheetTimeZone(APP.TZ);

  let tasks = child.getSheetByName(EVENT_SHEET.SHEETS.TASKS);
  if (!tasks) {
    const first = child.getSheets()[0];
    if (child.getSheets().length === 1 && first.getLastRow() === 0) {
      first.setName(EVENT_SHEET.SHEETS.TASKS);
      tasks = first;
    } else {
      tasks = child.insertSheet(EVENT_SHEET.SHEETS.TASKS,0);
    }
  }
  if (!child.getSheetByName(EVENT_SHEET.SHEETS.EXPENSES)) child.insertSheet(EVENT_SHEET.SHEETS.EXPENSES);
  if (!child.getSheetByName(PARTICIPANTS_V3.SHEET)) child.insertSheet(PARTICIPANTS_V3.SHEET);
  let meta = child.getSheetByName(EVENT_SHEET.SHEETS.META);
  if (!meta) meta = child.insertSheet(EVENT_SHEET.SHEETS.META);

  writeEventMetaV5_(child,eventId,event,folderId);
  meta.hideSheet();
}

function hideEventSheetTechnicalColumnsV3_(child) {
  const tasks = child.getSheetByName(EVENT_SHEET.SHEETS.TASKS);
  if (tasks) {
    tasks.showColumns(1,TASKS_V4.VISIBLE_COLS);
    if (tasks.getMaxColumns() > TASKS_V4.VISIBLE_COLS) {
      tasks.hideColumns(TASKS_V4.VISIBLE_COLS + 1, tasks.getMaxColumns() - TASKS_V4.VISIBLE_COLS);
    }
  }
  const expenses = child.getSheetByName(EVENT_SHEET.SHEETS.EXPENSES);
  if (expenses) {
    expenses.showColumns(1,EXPENSES_V4.VISIBLE_COLS);
    if (expenses.getMaxColumns() > EXPENSES_V4.VISIBLE_COLS) {
      expenses.hideColumns(EXPENSES_V4.VISIBLE_COLS + 1, expenses.getMaxColumns() - EXPENSES_V4.VISIBLE_COLS);
    }
  }
  const participants = child.getSheetByName(PARTICIPANTS_V3.SHEET);
  if (participants) {
    participants.showColumns(1,PARTICIPANTS_V3.VISIBLE_COLS);
    if (participants.getMaxColumns() > PARTICIPANTS_V3.VISIBLE_COLS) {
      participants.hideColumns(PARTICIPANTS_V3.VISIBLE_COLS + 1, participants.getMaxColumns() - PARTICIPANTS_V3.VISIBLE_COLS);
    }
  }
}
