function prepareEventSheetForSelectedEventV2() {
  const result = prepareEventSheetForSelectedEvent();
  refineSelectedEventSheetV2_();
  applyOutstandingExpenseFormulaV2_();
  return result;
}

function syncSelectedEventSheetToCalendarV2() {
  const result = syncSelectedEventSheetToCalendar();
  refineSelectedEventSheetV2_();
  applyOutstandingExpenseFormulaV2_();
  return result;
}

function refineSelectedEventSheetV2_() {
  const event = selectedEvent_();
  const eventId = ensureEventId_(event);
  const folder = createWorkFolderForEvent_(eventId, event, event._row);
  const child = findEventSheet_(eventId, event, folder.folderId);
  if (!child) return;
  refineExpenseSheetV2_(child);
  refineGuestRefundSummaryV2_(child);
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

function refineGuestRefundSummaryV2_(child) {
  const sheet = child.getSheetByName(EVENT_SHEET.SHEETS.GUESTS);
  if (!sheet) return;
  if (sheet.getMaxColumns() < 13) sheet.insertColumnsAfter(sheet.getMaxColumns(), 13 - sheet.getMaxColumns());

  sheet.getRange('M1').setValue('RIMBORSO PASSATO');
  if (sheet.getMaxColumns() > 13) {
    sheet.getRange(1, 14, sheet.getMaxRows(), sheet.getMaxColumns() - 13).clearContent();
    sheet.hideColumns(14, sheet.getMaxColumns() - 13);
  }
  sheet.showColumns(13);
  sheet.getRange('M2:M1000').clearContent().setNumberFormat('€ #,##0.00');

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const formulas = [];
  for (let row = 2; row <= lastRow; row++) {
    const fullName = 'TRIM($D' + row + '&" "&$E' + row + ')';
    formulas.push([
      '=IF(AND($D' + row + '="";$E' + row + '="");"";IF($C' + row + '<>"";' +
      'SUMIFS(Spese!$E:$E;Spese!$N:$N;$C' + row + ';Spese!$B:$B;"RIMBORSO";Spese!$F:$F;"PAGATO");' +
      'SUMIFS(Spese!$E:$E;Spese!$C:$C;' + fullName + ';Spese!$B:$B;"RIMBORSO";Spese!$F:$F;"PAGATO")))'
    ]);
  }
  sheet.getRange(2, 13, formulas.length, 1).setFormulas(formulas);
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
    "'_SPESE'!$L$2:$L$1005<>\"RIMBORSATO\"));0))';
}
