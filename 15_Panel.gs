function getSelectedEventShell() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!sheet || sheet.getName() !== APP.SHEETS.CALENDAR) {
    return { state: 'NO_EVENT', message: 'Seleziona una riga nel foglio Calendario.' };
  }
  const range = sheet.getActiveRange();
  const row = range ? range.getRow() : 0;
  if (row < 2) return { state: 'NO_EVENT', message: 'Seleziona una riga evento.' };
  return getEventShellByRow_(row);
}

function getEventShellByRow_(row, expectedEventId) {
  const sheet = sh_(APP.SHEETS.CALENDAR);
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const values = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
  const event = {};
  headers.forEach((h, i) => event[String(h || '').trim()] = values[i]);
  event._row = row;

  const hasData = values.slice(1).some(v => String(v || '').trim() !== '');
  if (!hasData) return { state: 'NO_EVENT', message: 'La riga selezionata è vuota.' };

  let eventId = String(event[APP.CALENDAR_HEADERS.ID] || '').trim();
  if (!eventId) eventId = ensureEventId_(event);
  if (expectedEventId && String(expectedEventId) !== eventId) {
    throw new Error('La riga selezionata non corrisponde più all’evento aperto. Ricarica la riga selezionata.');
  }

  return clientSafe_({
    state: 'EVENT',
    eventId: eventId,
    row: row,
    event: event,
    summary: {
      checklist: event[APP.CALENDAR_HEADERS.CHECKLIST] || '',
      nextAction: event[APP.CALENDAR_HEADERS.NEXT_ACTION] || '',
      budget: event[APP.CALENDAR_HEADERS.BUDGET] || 0,
      actual: event[APP.CALENDAR_HEADERS.ACTUAL] || 0,
      toPay: event[APP.CALENDAR_HEADERS.TO_PAY] || 0
    }
  });
}

function getEventFromPanelRow_(row, eventId) {
  const shell = getEventShellByRow_(Number(row), eventId);
  if (!shell || shell.state !== 'EVENT') throw new Error('Evento non disponibile.');
  const sheet = sh_(APP.SHEETS.CALENDAR);
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const values = sheet.getRange(Number(row), 1, 1, lastCol).getValues()[0];
  const event = {};
  headers.forEach((h, i) => event[String(h || '').trim()] = values[i]);
  event._row = Number(row);
  return event;
}

function getPanelChecklistSection(eventId) {
  return clientSafe_({ checklist: getChecklistForEvent_(eventId) });
}

function buildPanelExpensesSection_(eventId, event) {
  return {
    expenses: getExpensesForEvent_(eventId),
    participants: getParticipantsForEvent_(eventId),
    refundLimit: getEventRefundLimit_(eventId, event),
    standardRefundLimit: resolveRefundStandard_(event),
    cebOptions: getCebOptionsForEvent_(event)
  };
}

function getPanelExpensesSection(eventId, row) {
  const event = getEventFromPanelRow_(row, eventId);
  return clientSafe_(buildPanelExpensesSection_(eventId, event));
}

function getPanelParticipantsSection(eventId) {
  return clientSafe_({ participants: getParticipantsForEvent_(eventId) });
}

function saveEventPanelDraft(payload) {
  payload = payload || {};
  const eventId = String(payload.eventId || '').trim();
  const row = Number(payload.row || 0);
  if (!eventId || row < 2) throw new Error('Evento non valido.');

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) throw new Error('Il calendario è occupato da un altro salvataggio. Riprova tra qualche secondo.');

  try {
    const event = getEventFromPanelRow_(row, eventId);
    saveRefundLimitPanelDraft_(eventId, payload.refundLimitChange);
    saveParticipantRefundOverrides_(eventId, payload.participantRefundUpdates || []);
    saveChecklistPanelDraft_(eventId, payload.checklistUpdates || [], payload.newChecklist || [], payload.deletedChecklistIds || []);
    saveExpensesPanelDraft_(eventId, event, payload.expenseUpdates || [], payload.newExpenses || [], payload.newRefunds || []);
    SpreadsheetApp.flush();

    const result = {
      success: true,
      message: 'Modifiche salvate.',
      shell: getEventShellByRow_(row, eventId)
    };
    if (payload.returnChecklist) result.checklist = getChecklistForEvent_(eventId);
    if (payload.returnExpenses) result.expensesSection = buildPanelExpensesSection_(eventId, event);
    return clientSafe_(result);
  } finally {
    lock.releaseLock();
  }
}

function saveRefundLimitPanelDraft_(eventId, change) {
  if (!change || !change.mode) return;
  const sheet = sh_(APP.SHEETS.EVENT_META);
  const rows = sheet.getDataRange().getValues();
  let rowNumber = 0;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(eventId)) {
      rowNumber = i + 1;
      break;
    }
  }
  const now = new Date();
  if (normalize_(change.mode) === 'RESET') {
    if (rowNumber) {
      sheet.getRange(rowNumber, 2, 1, 3).setValues([[0, false, '']]);
      sheet.getRange(rowNumber, 8).setValue(now);
    }
    return;
  }
  const value = Number(change.value || 0);
  if (rowNumber) {
    sheet.getRange(rowNumber, 2, 1, 3).setValues([[value, true, 'PERSONALIZZATO']]);
    sheet.getRange(rowNumber, 8).setValue(now);
  } else {
    sheet.appendRow([eventId, value, true, 'PERSONALIZZATO', '', '', '', now]);
  }
}

function saveChecklistPanelDraft_(eventId, updates, newItems, deletedIds) {
  if (!updates.length && !newItems.length && !deletedIds.length) return;
  const sheet = sh_(APP.SHEETS.CHECKLIST);
  const rows = sheet.getDataRange().getValues();
  const byId = {};
  let maxOrder = 0;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) !== String(eventId)) continue;
    byId[String(rows[i][0])] = i + 1;
    maxOrder = Math.max(maxOrder, Number(rows[i][2] || 0));
  }

  const now = new Date();
  updates.forEach(u => {
    const r = byId[String(u.id || '')];
    if (!r) return;
    const old = rows[r - 1];
    const status = normalize_(u.status) === 'FATTO' ? 'FATTO' : 'DA FARE';
    const completedAt = status === 'FATTO' ? (old[11] || now) : '';
    sheet.getRange(r, 6, 1, 8).setValues([[
      parseClientDate_(u.dueDate) || '', status, old[7] || 'NORMALE', old[8] || '', old[9] || '',
      u.note || '', completedAt, now
    ]]);
  });

  const toDelete = deletedIds.map(id => byId[String(id)]).filter(Boolean).sort((a, b) => b - a);
  toDelete.forEach(r => sheet.deleteRow(r));

  const addRows = [];
  newItems.forEach(item => {
    const task = String(item.task || '').trim();
    if (!task) return;
    maxOrder += 10;
    addRows.push([
      'TASK-' + Utilities.getUuid(), eventId, maxOrder, task, 'PERSONALIZZATA',
      parseClientDate_(item.dueDate) || '', normalize_(item.status) === 'FATTO' ? 'FATTO' : 'DA FARE',
      item.priority || 'NORMALE', 'PERSONALIZZATA', '', item.note || '',
      normalize_(item.status) === 'FATTO' ? now : '', now
    ]);
  });
  if (addRows.length) sheet.getRange(sheet.getLastRow() + 1, 1, addRows.length, 13).setValues(addRows);
}

function expenseObjectFromPanelRow_(r) {
  return {
    id: r[0], eventId: r[1], type: r[2], category: r[3], ceb: r[4], description: r[5],
    beneficiary: r[6], personId: r[7], budget: Number(r[8] || 0), actual: Number(r[9] || 0),
    dueDate: r[10], status: r[11], paidDate: r[12], reference: r[13], attachment: r[14], notes: r[15],
    createdAt: r[16], updatedAt: r[17], rifStatus: r[18], rifRequestedAt: r[19], rifCode: r[20],
    invoiceReceived: r[21] === true, invoiceDate: r[22], receiptRequested: r[23] === true,
    receiptSent: r[24] === true, closedAt: r[25]
  };
}

function saveExpensesPanelDraft_(eventId, event, updates, newExpenses, newRefunds) {
  if (!updates.length && !newExpenses.length && !newRefunds.length) return;
  const sheet = sh_(APP.SHEETS.EXPENSES);
  const rows = sheet.getDataRange().getValues();
  const byId = {};
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) === String(eventId)) byId[String(rows[i][0])] = i + 1;
  }
  const now = new Date();
  const changedExpenses = [];

  updates.forEach(u => {
    const r = byId[String(u.id || '')];
    if (!r) return;
    const old = rows[r - 1].slice(0, 26);
    const budget = u.budget !== undefined ? Number(u.budget || 0) : Number(old[8] || 0);
    const actual = u.actual !== undefined ? Number(u.actual || 0) : Number(old[9] || 0);
    const rifNeeded = expenseNeedsRif_(budget, actual);
    const rifCode = u.rifCode !== undefined ? String(u.rifCode || '').trim() : String(old[20] || '').trim();
    let rifStatus = u.rifStatus !== undefined ? u.rifStatus : old[18];
    if (rifCode) rifStatus = 'RICEVUTO';
    else if (!rifNeeded) rifStatus = 'NON NECESSARIO';
    else if (!rifStatus || normalize_(rifStatus) === 'NON NECESSARIO') rifStatus = 'DA RICHIEDERE';

    const status = u.status !== undefined ? u.status : old[11];
    const invoiceReceived = u.invoiceReceived !== undefined ? !!u.invoiceReceived : old[21] === true;
    const receiptRequested = u.receiptRequested !== undefined ? !!u.receiptRequested : old[23] === true;
    const receiptSent = u.receiptSent !== undefined ? !!u.receiptSent : old[24] === true;
    const dueDate = u.dueDate !== undefined ? (parseClientDate_(u.dueDate) || '') : old[10];
    const category = u.category !== undefined ? u.category : old[3];
    const recordType = old[2] || 'SPESA';

    const updated = [
      old[0], eventId, recordType,
      category,
      resolveExpenseCeb_(event, old[4], category, recordType),
      u.description !== undefined ? u.description : old[5],
      u.beneficiary !== undefined ? u.beneficiary : old[6], old[7],
      budget, actual, dueDate, status,
      old[12], old[13], old[14], u.notes !== undefined ? u.notes : old[15],
      old[16] || now, now, rifStatus,
      (normalize_(rifStatus) === 'IN ATTESA' || normalize_(rifStatus) === 'RICEVUTO') ? (old[19] || now) : '',
      rifCode, invoiceReceived,
      invoiceReceived ? (old[22] || now) : '',
      receiptRequested, receiptSent,
      normalize_(status) === 'CHIUSO' ? (old[25] || now) : ''
    ];
    sheet.getRange(r, 1, 1, 26).setValues([updated]);
    changedExpenses.push(expenseObjectFromPanelRow_(updated));
  });

  const appendRows = [];
  newExpenses.forEach(p => {
    const description = String(p.description || '').trim();
    if (!description) return;
    const budget = 0;
    const actual = Number(p.actual || 0);
    const rifNeeded = expenseNeedsRif_(budget, actual);
    const rifCode = String(p.rifCode || '').trim();
    const rifStatus = rifCode ? 'RICEVUTO' : (rifNeeded ? (p.rifStatus || 'DA RICHIEDERE') : 'NON NECESSARIO');
    const status = p.status || 'DA DEFINIRE';
    const invoiceReceived = p.invoiceReceived === true;
    const receiptRequested = p.receiptRequested === true;
    const receiptSent = p.receiptSent === true;
    const dueDate = defaultPaymentDue_(event, p);
    const category = p.category || '';
    const ceb = resolveExpenseCeb_(event, '', category, 'SPESA');
    const row = [
      'SPESA-' + Utilities.getUuid(), eventId, 'SPESA', category, ceb,
      description, '', '', budget, actual, dueDate, status, '', '', '', p.notes || '', now, now,
      rifStatus, (normalize_(rifStatus) === 'IN ATTESA' || normalize_(rifStatus) === 'RICEVUTO') ? now : '', rifCode,
      invoiceReceived, invoiceReceived ? now : '', receiptRequested, receiptSent,
      normalize_(status) === 'CHIUSO' ? now : ''
    ];
    appendRows.push(row);
    changedExpenses.push(expenseObjectFromPanelRow_(row));
  });

  const currentRefundRows = rows.slice(1).filter(r => String(r[1]) === String(eventId) && normalize_(r[2]) === 'RIMBORSO');
  const runningRefunds = currentRefundRows.map(expenseObjectFromPanelRow_);

  newRefunds.forEach(p => {
    const beneficiary = String(p.beneficiary || '').trim();
    const amount = Number(p.amount || 0);
    const role = normalize_(p.role || 'ATLETA');
    if (!beneficiary || !amount) return;
    const paidDate = parseClientDate_(p.paidDate) || now;
    const sameBeneficiary = runningRefunds.filter(x => normalize_(x.beneficiary) === normalize_(beneficiary));
    if (sameBeneficiary.some(x => Number(x.actual || 0) === amount && sameDay_(x.paidDate, paidDate))) {
      throw new Error('Possibile rimborso duplicato per ' + beneficiary + '.');
    }
    const previous = sameBeneficiary.reduce((s, x) => s + Number(x.actual || 0), 0);
    const projected = previous + amount;
    const limitData = role === 'ATLETA'
      ? getParticipantRefundLimitForBeneficiary_(eventId, event, beneficiary)
      : { value: 0, participant: null };
    const limit = role === 'ATLETA' ? Number(limitData.value || 0) : 0;
    if (limit > 0 && projected > limit && p.confirmedOverLimit !== true) {
      throw new Error('Il rimborso di ' + beneficiary + ' supera il massimale atleta. Conferma prima il superamento nel pannello.');
    }
    const participant = limitData.participant || null;
    const row = [
      'RIM-' + Utilities.getUuid(), eventId, 'RIMBORSO', 'RIMBORSI ' + role, 'CEB.002',
      'Rimborso spese ' + role.toLowerCase(), beneficiary, participant ? participant.personId || '' : '', 0, amount, '', 'CHIUSO', paidDate,
      '', '', p.notes || '', now, now, 'NON NECESSARIO', '', '', true, now, false, false, now
    ];
    appendRows.push(row);
    runningRefunds.push(expenseObjectFromPanelRow_(row));
  });

  if (appendRows.length) sheet.getRange(sheet.getLastRow() + 1, 1, appendRows.length, 26).setValues(appendRows);
  changedExpenses.forEach(syncExpenseTasks_);
}
