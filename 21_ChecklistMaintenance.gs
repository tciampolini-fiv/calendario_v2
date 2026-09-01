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
  const completedStandardKeys = {};
  const toDelete = [];
  let legacyOpenRemoved = 0;
  let autoRemoved = 0;

  for (let i = 1; i < checklistRows.length; i++) {
    const eventId = String(checklistRows[i][1] || '').trim();
    if (!targetEventIds.has(eventId)) continue;
    const source = normalize_(checklistRows[i][8]);
    const status = normalize_(checklistRows[i][6]);
    const done = status === 'FATTO' || status === 'COMPLETATA';

    if (source === 'STANDARD') {
      if (done) {
        const key = normalize_(checklistRows[i][9] || checklistRows[i][3]);
        if (key) {
          if (!completedStandardKeys[eventId]) completedStandardKeys[eventId] = new Set();
          completedStandardKeys[eventId].add(key);
        }
      }
      toDelete.push(i + 1);
      continue;
    }

    if (source === 'AUTO') {
      toDelete.push(i + 1);
      autoRemoved++;
      continue;
    }

    // Le vecchie voci LEGACY aperte sono la principale fonte di duplicati.
    // Quelle già concluse restano archiviate, ma non vengono più mostrate.
    if (source === 'LEGACY' && !done) {
      toDelete.push(i + 1);
      legacyOpenRemoved++;
    }
  }

  toDelete.sort((a, b) => b - a).forEach(r => checklist.deleteRow(r));

  let added = 0;
  events.forEach(x => { added += generateChecklistForEvent_(x.id, x.event); });

  // Ripristina le attività STANDARD già completate prima della pulizia.
  const rebuilt = checklist.getDataRange().getValues();
  const now = new Date();
  for (let i = 1; i < rebuilt.length; i++) {
    const eventId = String(rebuilt[i][1] || '').trim();
    const saved = completedStandardKeys[eventId];
    if (!saved || normalize_(rebuilt[i][8]) !== 'STANDARD') continue;
    const key = normalize_(rebuilt[i][9] || rebuilt[i][3]);
    if (!saved.has(key)) continue;
    checklist.getRange(i + 1, 7).setValue('FATTO');
    checklist.getRange(i + 1, 12, 1, 2).setValues([[rebuilt[i][11] || now, now]]);
  }

  // Ricrea soltanto l'automatismo ancora utile: fattura successiva a un AFOR.
  const expenses = sh_(APP.SHEETS.EXPENSES).getDataRange().getValues();
  for (let i = 1; i < expenses.length; i++) {
    const eventId = String(expenses[i][1] || '').trim();
    if (!targetEventIds.has(eventId)) continue;
    syncExpenseTasks_(expenseObjectFromPanelRow_(expenses[i]));
  }

  events.forEach(x => syncChecklistLocksForEvent_(x.id));
  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    'Checklist ripulite',
    'Eventi in corso/futuri: ' + events.length + '\n' +
    'Checklist STANDARD ricreate: ' + added + '\n' +
    'Vecchie attività LEGACY aperte eliminate: ' + legacyOpenRemoved + '\n' +
    'Vecchi automatismi eliminati: ' + autoRemoved + '\n\n' +
    'Le attività personalizzate e le attività già concluse sono state preservate.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
