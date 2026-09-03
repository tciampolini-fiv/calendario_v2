function prepareEventSheetForSelectedEventV2() {
  const event = selectedEvent_();
  const eventId = ensureEventId_(event);
  const folder = createWorkFolderForEvent_(eventId, event, event._row);
  const result = getOrCreateEventSheetV3_(eventId, event, folder.folderId);
  const child = result.spreadsheet;

  migrateLegacyEventSheetToV3_(eventId,event,child,folder.folderId);
  ensureChecklistBackendHeadersV3_();
  ensureParticipantsBackendHeadersV2_();
  refreshTasksV3FromBackend_(eventId,event,child);
  refreshExpensesV3FromBackend_(eventId,event,child,folder.folderId);
  refreshParticipantsV2FromBackend_(eventId,event,child);
  hideEventSheetTechnicalColumnsV3_(child);
  setEventSheetLink_(event._row, child.getUrl());
  writeMeta_(child.getSheetByName(EVENT_SHEET.SHEETS.META),{
    LAST_SYNC:Utilities.formatDate(new Date(),APP.TZ,'dd/MM/yyyy HH:mm'),
    SYNC_VERSION:'3'
  });
  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    result.created ? 'Scheda evento creata' : 'Scheda evento aggiornata',
    'La scheda evento usa ora Attività, Spese e Partecipanti. Aprila dalla colonna SCHEDA EVENTO.',
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

  const counts = {
    tasks: syncTasksV3ToBackend_(eventId,event,child),
    expenses: syncExpensesV3ToBackend_(eventId,event,child),
    participants: syncParticipantsV2ToBackend_(eventId,event,child)
  };

  seedPresenceCheckTaskV3_(eventId,event,child);
  ensureTaskNumbersAndDefaultDependenciesV3_(eventId,event);
  recomputeTaskDependenciesV3_(eventId,event);
  refreshTasksV3FromBackend_(eventId,event,child);
  refreshExpensesV3FromBackend_(eventId,event,child,folder.folderId);
  refreshParticipantsV2FromBackend_(eventId,event,child);
  hideEventSheetTechnicalColumnsV3_(child);

  setEventSheetLink_(event._row, child.getUrl());
  writeMeta_(child.getSheetByName(EVENT_SHEET.SHEETS.META),{
    LAST_SYNC:Utilities.formatDate(new Date(),APP.TZ,'dd/MM/yyyy HH:mm'),
    SYNC_VERSION:'3'
  });
  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    'Calendario aggiornato',
    'Attività: ' + counts.tasks + '\nSpese: ' + counts.expenses + '\nPartecipanti: ' + counts.participants,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return counts;
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
  if (!child.getSheetByName(PARTICIPANTS_V2.SHEET)) child.insertSheet(PARTICIPANTS_V2.SHEET);
  let meta = child.getSheetByName(EVENT_SHEET.SHEETS.META);
  if (!meta) meta = child.insertSheet(EVENT_SHEET.SHEETS.META);

  writeMeta_(meta,{
    EVENT_ID:eventId,
    MASTER_SPREADSHEET_ID:APP.SPREADSHEET_ID,
    EVENT_FOLDER_ID:folderId,
    EVENT_SHEET_ID:child.getId(),
    SYNC_VERSION:'3',
    EVENT_LABEL:buildEventSheetLabel_(event)
  });
  meta.hideSheet();
}

function migrateLegacyEventSheetToV3_(eventId,event,child,folderId) {
  const tasks = child.getSheetByName(EVENT_SHEET.SHEETS.TASKS);
  if (isLegacyTaskSheetV3_(tasks)) {
    syncTasksFromEventSheet_(eventId,child);
  }

  const expenses = child.getSheetByName(EVENT_SHEET.SHEETS.EXPENSES);
  if (isLegacyExpenseSheetV3_(expenses)) {
    syncExpensesFromEventSheet_(eventId,child);
  }

  ensureChecklistBackendHeadersV3_();
  ensureParticipantsBackendHeadersV2_();
  refreshTasksV3FromBackend_(eventId,event,child);
  refreshExpensesV3FromBackend_(eventId,event,child,folderId);
  ensureParticipantV2Structure_(child);
  hideEventSheetTechnicalColumnsV3_(child);
}

function hideEventSheetTechnicalColumnsV3_(child) {
  const tasks = child.getSheetByName(EVENT_SHEET.SHEETS.TASKS);
  if (tasks) {
    tasks.showColumns(1,TASKS_V3.VISIBLE_COLS);
    if (tasks.getMaxColumns() > TASKS_V3.VISIBLE_COLS) {
      tasks.hideColumns(TASKS_V3.VISIBLE_COLS + 1, tasks.getMaxColumns() - TASKS_V3.VISIBLE_COLS);
    }
  }
  const expenses = child.getSheetByName(EVENT_SHEET.SHEETS.EXPENSES);
  if (expenses) {
    expenses.showColumns(1,EXPENSES_V3.VISIBLE_COLS);
    if (expenses.getMaxColumns() > EXPENSES_V3.VISIBLE_COLS) {
      expenses.hideColumns(EXPENSES_V3.VISIBLE_COLS + 1, expenses.getMaxColumns() - EXPENSES_V3.VISIBLE_COLS);
    }
  }
  const participants = child.getSheetByName(PARTICIPANTS_V2.SHEET);
  if (participants) {
    participants.showColumns(1,PARTICIPANTS_V2.VISIBLE_COLS);
    if (participants.getMaxColumns() > PARTICIPANTS_V2.VISIBLE_COLS) {
      participants.hideColumns(PARTICIPANTS_V2.VISIBLE_COLS + 1, participants.getMaxColumns() - PARTICIPANTS_V2.VISIBLE_COLS);
    }
  }
}
