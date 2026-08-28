function focusPanelEventRow(row, eventId) {
  row = Number(row || 0);
  if (row < 2) return false;

  const sheet = sh_(APP.SHEETS.CALENDAR);
  const map = headerMap_(sheet);
  const idCol = map[APP.CALENDAR_HEADERS.ID];
  if (!idCol) return false;

  const currentId = String(sheet.getRange(row, idCol).getDisplayValue() || '').trim();
  if (eventId && currentId !== String(eventId)) return false;

  sheet.setActiveRange(sheet.getRange(row, 1, 1, sheet.getLastColumn()));
  return true;
}
