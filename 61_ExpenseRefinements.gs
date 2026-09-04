function prepareEventSheetForSelectedEventV2() {
  const event = selectedEvent_();
  const eventId = ensureEventId_(event);
  const folder = createWorkFolderForEvent_(eventId, event, event._row);
  const result = getOrCreateEventSheetV3_(eventId, event, folder.folderId);
  const child = result.spreadsheet;

  migrateLegacyEventSheetToV3_(eventId,event,child,folder.folderId);
  ensureChecklistBackendHeadersV3_();
  ensureParticipantsBackendHeadersV2_();
  refreshTasksV4FromBackend_(eventId,event,child);
  removeNativeTaskTableCompatibility_(child);
  refreshParticipantsV3FromBackend_(eventId,event,child);
  refreshExpensesV4FromBackend_(eventId,event,child,folder.folderId);
  ensureEventSheetEditTriggerV5_(child);
  hideEventSheetTechnicalColumnsV3_(child);
  setEventSheetLink_(event._row, child.getUrl());
  writeEventMetaV5_(child,eventId,event,folder.folderId);
  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    result.created ? 'Scheda evento creata' : 'Scheda evento aggiornata',
    'La scheda evento usa Attività, Spese e Partecipanti. Aprila dalla colonna SCHEDA EVENTO.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return {created:result.created,id:child.getId(),url:child.getUrl()};
}

function syncSelectedEventSheetToCalendarV2() {
  const event = selectedEvent_();
  const eventId = ensureEventId_(event);
  const folder = createWorkFolderForEvent_(eventId, event, event._row);
  const result = getOrCreateEventSheetV3_(eventId, event, folder.folderId);
  const child = result.spreadsheet;
  validateEventSheetIdentity_(child,eventId);

  migrateLegacyEventSheetToV3_(eventId,event,child,folder.folderId);

  // Prima i partecipanti: i rimborsi del foglio Partecipanti diventano poi righe Spese.
  const participantCount = syncParticipantsV3ToBackend_(eventId,event,child);
  const taskCount = syncTasksV4ToBackend_(eventId,event,child);
  removeNativeTaskTableCompatibility_(child);
  const expenseCount = syncExpensesV4ToBackend_(eventId,event,child);
  const counts = {
    tasks: taskCount,
    participants: participantCount,
    expenses: expenseCount
  };

  seedPresenceCheckTaskV3_(eventId,event,child);
  ensureTaskNumbersAndDefaultDependenciesV3_(eventId,event);
  refreshTasksV4FromBackend_(eventId,event,child);
  removeNativeTaskTableCompatibility_(child);
  refreshParticipantsV3FromBackend_(eventId,event,child);
  refreshExpensesV4FromBackend_(eventId,event,child,folder.folderId);
  ensureEventSheetEditTriggerV5_(child);
  hideEventSheetTechnicalColumnsV3_(child);

  setEventSheetLink_(event._row, child.getUrl());
  writeEventMetaV5_(child,eventId,event,folder.folderId);
  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    'Calendario aggiornato',
    'Attività: ' + counts.tasks + '\nSpese: ' + counts.expenses + '\nPartecipanti: ' + counts.participants,
    SpreadsheetApp.getUi().ButtonSet.OK
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
  let child = findEventSheet_(eventId,event,folderId);
  let created = false;
  if (!child) {
    child = SpreadsheetApp.create(buildEventSheetName_(event));
    child.setSpreadsheetLocale('it_IT');
    child.setSpreadsheetTimeZone(APP.TZ);
    DriveApp.getFileById(child.getId()).moveTo(DriveApp.getFolderById(folderId));
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

function migrateLegacyEventSheetToV3_(eventId,event,child,folderId) {
  // Le tabelle native di Sheets possono tipizzare le colonne e bloccare setNumberFormat().
  // La vista Attività non ha bisogno della tabella nativa: usiamo range, menu e formattazione condizionale.
  removeNativeTaskTableCompatibility_(child);

  const tasks = child.getSheetByName(EVENT_SHEET.SHEETS.TASKS);
  if (isLegacyTaskSheetV3_(tasks)) syncTasksFromEventSheet_(eventId,child);

  const expenses = child.getSheetByName(EVENT_SHEET.SHEETS.EXPENSES);
  if (isLegacyExpenseSheetV3_(expenses)) syncExpensesFromEventSheet_(eventId,child);

  // Se è ancora la vecchia vista Spese V3, la leggiamo prima di trasformarla nel nuovo cruscotto.
  if (expenses && normalize_(expenses.getRange('A1').getDisplayValue()) === 'IMPORTO') {
    syncExpensesV3ToBackend_(eventId,event,child);
  }

  ensureChecklistBackendHeadersV3_();
  ensureParticipantsBackendHeadersV2_();
  refreshTasksV4FromBackend_(eventId,event,child);
  removeNativeTaskTableCompatibility_(child);
  refreshParticipantsV3FromBackend_(eventId,event,child);
  refreshExpensesV4FromBackend_(eventId,event,child,folderId);
  hideEventSheetTechnicalColumnsV3_(child);
}

function removeNativeTaskTableCompatibility_(child) {
  try {
    const sheet = child.getSheetByName(EVENT_SHEET.SHEETS.TASKS);
    if (!sheet) return;
    const ss = Sheets.Spreadsheets.get(child.getId(), {
      fields: 'sheets(properties(sheetId),tables(tableId))'
    });
    const info = (ss.sheets || []).find(s =>
      s.properties && Number(s.properties.sheetId) === Number(sheet.getSheetId())
    );
    const tables = info && info.tables ? info.tables : [];
    if (!tables.length) return;
    Sheets.Spreadsheets.batchUpdate({
      requests: tables.map(t => ({deleteTable:{tableId:t.tableId}}))
    }, child.getId());
  } catch (e) {
    console.log('Rimozione tabella nativa Attività non necessaria/non riuscita: ' + e.message);
  }
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
