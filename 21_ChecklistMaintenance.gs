function rebuildCurrentAndFutureStandardChecklists() {
  const cal = sh_(APP.SHEETS.CALENDAR);
  const rows = cal.getDataRange().getValues();
  if (rows.length < 2) return;

  const headers = rows[0];
  const idx = {};
  headers.forEach((h, i) => idx[String(h || '').trim()] = i);

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const targetEventIds = new Set();
  const events = [];

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const eventId = String(r[idx[APP.CALENDAR_HEADERS.ID]] || '').trim();
    if (!eventId) continue;

    const end = r[idx[APP.CALENDAR_HEADERS.END]];
    if (end instanceof Date) {
      const endDay = new Date(end);
      endDay.setHours(0, 0, 0, 0);
      if (endDay < today) continue;
    }

    const event = {};
    headers.forEach((h, c) => event[String(h || '').trim()] = r[c]);
    event._row = i + 1;
    targetEventIds.add(eventId);
    events.push({ id: eventId, event: event });
  }

  const checklist = sh_(APP.SHEETS.CHECKLIST);
  const checklistRows = checklist.getDataRange().getValues();
  const toDelete = [];
  for (let i = 1; i < checklistRows.length; i++) {
    const eventId = String(checklistRows[i][1] || '').trim();
    const source = normalize_(checklistRows[i][8]);
    if (targetEventIds.has(eventId) && source === 'STANDARD') toDelete.push(i + 1);
  }
  toDelete.sort((a, b) => b - a).forEach(r => checklist.deleteRow(r));

  let added = 0;
  events.forEach(x => { added += generateChecklistForEvent_(x.id, x.event); });
  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    'Checklist riallineate',
    'Sono state sostituite solo le checklist STANDARD degli eventi in corso e futuri.\n' +
    'Le voci LEGACY, PERSONALIZZATE e AUTO sono rimaste intatte.\n\n' +
    'Voci standard eliminate: ' + toDelete.length + '\nVoci standard ricreate: ' + added,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
