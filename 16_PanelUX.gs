function focusPanelEventRow(row, eventId) {
  row = Number(row || 0);
  if (row < 2) return false;

  const sheet = sh_(APP.SHEETS.CALENDAR);
  const map = headerMap_(sheet);
  const idCol = map[APP.CALENDAR_HEADERS.ID];
  if (!idCol) return false;

  const currentId = String(sheet.getRange(row, idCol).getDisplayValue() || '').trim();
  if (eventId && currentId !== String(eventId)) return false;

  // Z1 e una cella tecnica del foglio _CONFIG: la formattazione condizionale
  // del Calendario la usa per mantenere evidenziata la riga aperta nel pannello,
  // anche se l'utente clicca successivamente in un'altra cella.
  sh_(APP.SHEETS.CONFIG).getRange('Z1').setValue(currentId);
  sheet.setActiveRange(sheet.getRange(row, 1, 1, sheet.getLastColumn()));
  SpreadsheetApp.flush();
  return true;
}
