const EVENT_PARTICIPANT_SHEET_V2 = 'Partecipanti';

function prepareEventSheetForSelectedEventV2() {
  const result = prepareEventSheetForSelectedEvent();
  const event = selectedEvent_();
  const eventId = ensureEventId_(event);
  const child = findSelectedEventSheetV2_();
  if (child) {
    ensureParticipantsBackendHeadersV2_();
    refreshParticipantsV2FromBackend_(eventId,event,child);
    refineExpenseSheetV2_(child);
    const legacy = child.getSheetByName(EVENT_SHEET.SHEETS.GUESTS);
    if (legacy) legacy.hideSheet();
  }
  applyOutstandingExpenseFormulaV2_();
  return result;
}

function syncSelectedEventSheetToCalendarV2() {
  const result = syncSelectedEventSheetToCalendar();
  const event = selectedEvent_();
  const eventId = ensureEventId_(event);
  const child = findSelectedEventSheetV2_();
  if (child) {
    syncParticipantsV2ToBackend_(eventId,event,child);
    refreshParticipantsV2FromBackend_(eventId,event,child);
    refineExpenseSheetV2_(child);
    const legacy = child.getSheetByName(EVENT_SHEET.SHEETS.GUESTS);
    if (legacy) legacy.hideSheet();
  }
  applyOutstandingExpenseFormulaV2_();
  return result;
}

function findSelectedEventSheetV2_() {
  try {
    const event = selectedEvent_();
    const eventId = ensureEventId_(event);
    const folder = createWorkFolderForEvent_(eventId, event, event._row);
    return findEventSheet_(eventId, event, folder.folderId);
  } catch (e) {
    return null;
  }
}

function getParticipantSheetV2_(child) {
  return child.getSheetByName(EVENT_PARTICIPANT_SHEET_V2) || child.getSheetByName(EVENT_SHEET.SHEETS.GUESTS);
}

function refineSelectedEventSheetV2_() {
  const event = selectedEvent_();
  const eventId = ensureEventId_(event);
  const folder = createWorkFolderForEvent_(eventId, event, event._row);
  const child = findEventSheet_(eventId, event, folder.folderId);
  if (!child) return;
  ensureParticipantsBackendHeadersV2_();
  refreshParticipantsV2FromBackend_(eventId,event,child);
  refineExpenseSheetV2_(child);
  const legacy = child.getSheetByName(EVENT_SHEET.SHEETS.GUESTS);
  if (legacy) legacy.hideSheet();
}

function refineExpenseSheetV2_(child) {
  const sheet = child.getSheetByName(EVENT_SHEET.SHEETS.EXPENSES);
  if (!sheet) return;

  // J e' testo (PROMEMORIA 2), non una data.
  sheet.getRange('J2:J1000').setNumberFormat('@');

  // Ripulisce eventuali vecchie validazioni/formati rimasti nelle colonne tecniche nascoste.
  if (sheet.getMaxColumns() >= 19) {
    sheet.getRange('L2:S1000').clearDataValidations();
    sheet.getRange('L2:P1000').setNumberFormat('@');
    sheet.getRange('Q2:S1000').setNumberFormat('dd/MM/yyyy');
  }
}

function applyOutstandingExpenseFormulaV2_() {
  const sheet = sh_(APP.SHEETS.CALENDAR);
  const map = headerMap_(sheet);
  const idCol = map[APP.CALENDAR_HEADERS.ID];
  const dueCol = map[APP.CALENDAR_HEADERS.TO_PAY];
  if (!idCol || !dueCol) return;

  const firstRow = 2;
  const lastRow = Math.max(sheet.getMaxRows(), 2);
  const firstFormula = buildOutstandingExpenseFormulaV2_(firstRow);
  sheet.getRange(firstRow, dueCol).setFormula(firstFormula);
  sheet.getRange(firstRow, dueCol).copyTo(
    sheet.getRange(firstRow, dueCol, lastRow - firstRow + 1, 1),
    SpreadsheetApp.CopyPasteType.PASTE_FORMULA,
    false
  );
}

function buildOutstandingExpenseFormulaV2_(row) {
  const a = '$A' + row;
  return '=IF(' + a + '="";"";IFERROR(SUM(FILTER(' +
    "'_SPESE'!$J$2:$J$1005;" +
    "'_SPESE'!$B$2:$B$1005=" + a + ';' +
    "'_SPESE'!$J$2:$J$1005<>0;" +
    "'_SPESE'!$L$2:$L$1005<>\"PAGATO - FATTURA\";" +
    "'_SPESE'!$L$2:$L$1005<>\"PAGATO - AFOR\";" +
    "'_SPESE'!$L$2:$L$1005<>\"PAGAMENTO FATTURA\";" +
    "'_SPESE'!$L$2:$L$1005<>\"PAGAMENTO AFOR\";" +
    "'_SPESE'!$L$2:$L$1005<>\"AFOR FATTO\";" +
    "'_SPESE'!$L$2:$L$1005<>\"PAGATO CON CC\";" +
    "'_SPESE'!$L$2:$L$1005<>\"INVIATO IN AMMINISTRAZIONE\";" +
    "'_SPESE'!$L$2:$L$1005<>\"CHIUSO\";" +
    "'_SPESE'!$L$2:$L$1005<>\"PAGATO\";" +
    "'_SPESE'!$L$2:$L$1005<>\"RIMBORSATO\"));0))";
}
