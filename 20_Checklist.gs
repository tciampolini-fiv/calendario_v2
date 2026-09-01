function getChecklistRowsForEventRaw_(eventId) {
  const rows = sh_(APP.SHEETS.CHECKLIST).getDataRange().getValues();
  return rows.slice(1).filter(r => String(r[1]) === String(eventId)).map(r => ({
    id: r[0], eventId: r[1], order: r[2], task: r[3], category: r[4], dueDate: r[5], status: r[6],
    priority: r[7], source: r[8], autoKey: r[9], note: r[10], completedAt: r[11], updatedAt: r[12]
  }));
}

/**
 * La sidebar mostra solo le attività aperte e realmente eseguibili adesso.
 * Le attività completate e quelle BLOCCATE restano nel backend ma non vengono mostrate.
 */
function getChecklistForEvent_(eventId) {
  syncChecklistLocksForEvent_(eventId);
  return getChecklistRowsForEventRaw_(eventId)
    .filter(x => normalize_(x.status) === 'DA FARE')
    .sort((a, b) => {
      const aDue = a.dueDate instanceof Date ? a.dueDate.getTime() : Number.POSITIVE_INFINITY;
      const bDue = b.dueDate instanceof Date ? b.dueDate.getTime() : Number.POSITIVE_INFINITY;
      if (aDue !== bDue) return aDue - bDue;
      return Number(a.order || 0) - Number(b.order || 0);
    });
}

function getChecklistProfileForEvent_(event) {
  const type = normalize_(event[APP.CALENDAR_HEADERS.TYPE]);
  const rows = sh_(APP.SHEETS.EVENT_TYPE_CONFIG).getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (normalize_(rows[i][0]) === type) return normalize_(rows[i][3]) || type;
  }
  return type;
}

function getChecklistRulesForEvent_(event) {
  const config = sh_(APP.SHEETS.CHECKLIST_CONFIG).getDataRange().getValues();
  const profile = getChecklistProfileForEvent_(event);
  const cls = normalize_(event[APP.CALENDAR_HEADERS.CLASS]);
  const rules = {};

  config.slice(1).forEach(r => {
    const cfgProfile = normalize_(r[1]);
    const cfgClass = normalize_(r[2]);
    const active = r[10] === true || normalize_(r[10]) === 'SI';
    if (!active || (cfgProfile !== profile && cfgProfile !== 'TUTTI')) return;
    if (cfgClass && cfgClass !== '*' && cfgClass !== cls) return;

    const key = normalize_(r[9] || r[3]);
    if (!key) return;
    rules[key] = {
      dependsOn: normalize_(r[12] || ''),
      unlockDateBase: normalize_(r[13] || '')
    };
  });
  return rules;
}

function eventHasEnded_(event, today) {
  const end = event && event[APP.CALENDAR_HEADERS.END];
  if (!(end instanceof Date)) return false;
  const endDay = new Date(end);
  endDay.setHours(0, 0, 0, 0);
  const day = today instanceof Date ? new Date(today) : new Date();
  day.setHours(0, 0, 0, 0);
  return day > endDay;
}

/**
 * Gestisce gli sblocchi delle checklist STANDARD e della sola checklist AUTO
 * ammessa: richiesta fattura dopo pagamento AFOR.
 */
function syncChecklistLocksForEvent_(eventId) {
  const found = findCalendarEventById_(eventId);
  if (!found) return;

  const event = found.event;
  const rules = getChecklistRulesForEvent_(event);
  const sheet = sh_(APP.SHEETS.CHECKLIST);
  const rows = sheet.getDataRange().getValues();
  const eventRows = [];
  const byKey = {};

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) !== String(eventId)) continue;
    const item = {
      row: i + 1,
      values: rows[i],
      key: normalize_(rows[i][9] || rows[i][3]),
      source: normalize_(rows[i][8])
    };
    eventRows.push(item);
    if (item.key) byKey[item.key] = item;
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const ended = eventHasEnded_(event, today);
  const updates = [];

  eventRows.forEach(item => {
    const current = normalize_(item.values[6]);
    if (current === 'FATTO' || current === 'COMPLETATA') return;

    let managed = false;
    let blocked = false;

    if (item.source === 'STANDARD') {
      managed = true;
      const rule = rules[item.key] || {};
      if (rule.dependsOn) {
        const previous = byKey[rule.dependsOn];
        const previousStatus = previous ? normalize_(previous.values[6]) : '';
        if (previousStatus !== 'FATTO' && previousStatus !== 'COMPLETATA') blocked = true;
      }
      if (!blocked && rule.unlockDateBase === 'FINE' && !ended) blocked = true;
    } else if (item.source === 'AUTO' && item.key.indexOf('FATTURA_AFOR:') === 0) {
      managed = true;
      blocked = !ended;
    }

    if (!managed) return;
    const desired = blocked ? 'BLOCCATA' : 'DA FARE';
    if (current !== desired) updates.push({ row: item.row, status: desired });
  });

  updates.forEach(u => {
    sheet.getRange(u.row, 7).setValue(u.status);
    sheet.getRange(u.row, 13).setValue(new Date());
  });
  if (updates.length) SpreadsheetApp.flush();
}

/**
 * Corregge gli sblocchi quando si apre il file. Serve soprattutto alle attività
 * legate alla fine dell'evento, che possono diventare disponibili senza un edit.
 */
function syncAllChecklistLocks_() {
  const rows = sh_(APP.SHEETS.CHECKLIST).getDataRange().getValues();
  const ids = new Set();
  for (let i = 1; i < rows.length; i++) {
    const source = normalize_(rows[i][8]);
    const status = normalize_(rows[i][6]);
    if (!rows[i][1] || status === 'FATTO' || status === 'COMPLETATA') continue;
    if (source === 'STANDARD' || (source === 'AUTO' && normalize_(rows[i][9]).indexOf('FATTURA_AFOR:') === 0)) {
      ids.add(String(rows[i][1]));
    }
  }
  ids.forEach(syncChecklistLocksForEvent_);
}

function generateChecklistForEvent_(eventId, event) {
  const profile = getChecklistProfileForEvent_(event);
  if (!profile || profile === 'NESSUNO') return 0;

  const target = sh_(APP.SHEETS.CHECKLIST);
  const existing = getChecklistRowsForEventRaw_(eventId);
  const existingKeys = new Set(existing.map(x => normalize_(x.autoKey || x.task)));
  const config = sh_(APP.SHEETS.CHECKLIST_CONFIG).getDataRange().getValues();
  const cls = normalize_(event[APP.CALENDAR_HEADERS.CLASS]);
  const start = event[APP.CALENDAR_HEADERS.START];
  const end = event[APP.CALENDAR_HEADERS.END];
  const now = new Date();
  const appendRows = [];

  config.slice(1).forEach((r, idx) => {
    const cfgProfile = normalize_(r[1]);
    const cfgClass = normalize_(r[2]);
    const active = r[10] === true || normalize_(r[10]) === 'SI';
    if (!active || (cfgProfile !== profile && cfgProfile !== 'TUTTI')) return;
    if (cfgClass && cfgClass !== '*' && cfgClass !== cls) return;

    const task = String(r[3] || '').trim();
    const autoKey = String(r[9] || task).trim();
    if (!task || existingKeys.has(normalize_(autoKey))) return;

    const base = normalize_(r[5]);
    const offsetDays = r[6] === '' ? null : Number(r[6] || 0);
    let baseDate = base === 'FINE' ? end : start;
    if (!(baseDate instanceof Date)) baseDate = null;
    const due = baseDate && offsetDays !== null
      ? new Date(baseDate.getTime() + offsetDays * 86400000)
      : '';

    appendRows.push([
      'TASK-' + Utilities.getUuid(), eventId, Number(r[8] || (idx + 1) * 10), task,
      r[4] || '', due, 'DA FARE', '', 'STANDARD', autoKey, r[11] || '', '', now
    ]);
    existingKeys.add(normalize_(autoKey));
  });

  if (appendRows.length) {
    target.getRange(target.getLastRow() + 1, 1, appendRows.length, 13).setValues(appendRows);
  }
  syncChecklistLocksForEvent_(eventId);
  return appendRows.length;
}
