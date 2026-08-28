function focusPanelEventRow(row, eventId) {
  row = Number(row || 0);
  if (row < 2) return false;

  const sheet = sh_(APP.SHEETS.CALENDAR);
  const map = headerMap_(sheet);
  const idCol = map[APP.CALENDAR_HEADERS.ID];
  if (!idCol) return false;

  const currentId = String(sheet.getRange(row, idCol).getDisplayValue() || '').trim();
  if (!currentId || (eventId && currentId !== String(eventId))) return false;

  const configCell = sh_(APP.SHEETS.CONFIG).getRange('Z1');
  const previousId = String(configCell.getDisplayValue() || '').trim();

  if (previousId && previousId !== currentId) {
    clearPanelEventBorderById_(sheet, idCol, previousId);
  }

  configCell.setValue(currentId);
  setPanelEventRowBorder_(sheet, row, true);
  sheet.setActiveRange(sheet.getRange(row, 1, 1, sheet.getLastColumn()));
  SpreadsheetApp.flush();
  return true;
}

function clearPanelEventRowHighlight() {
  const sheet = sh_(APP.SHEETS.CALENDAR);
  const map = headerMap_(sheet);
  const idCol = map[APP.CALENDAR_HEADERS.ID];
  if (!idCol) return false;

  const configCell = sh_(APP.SHEETS.CONFIG).getRange('Z1');
  const eventId = String(configCell.getDisplayValue() || '').trim();
  if (eventId) clearPanelEventBorderById_(sheet, idCol, eventId);
  configCell.clearContent();
  SpreadsheetApp.flush();
  return true;
}

function clearPanelEventBorderById_(sheet, idCol, eventId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2 || !eventId) return;
  const found = sheet.getRange(2, idCol, lastRow - 1, 1)
    .createTextFinder(String(eventId))
    .matchEntireCell(true)
    .findNext();
  if (found) setPanelEventRowBorder_(sheet, found.getRow(), false);
}

function setPanelEventRowBorder_(sheet, row, active) {
  const range = sheet.getRange(row, 1, 1, sheet.getLastColumn());
  if (active) {
    range.setBorder(
      true, true, true, true, null, null,
      '#1a73e8', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    );
  } else {
    range.setBorder(
      true, true, true, true, null, null,
      '#d5d7da', SpreadsheetApp.BorderStyle.SOLID
    );
  }
}
