function getExpensesForEvent_(eventId) {
  const rows = sh_(APP.SHEETS.EXPENSES).getDataRange().getValues();
  return rows.slice(1).filter(r => String(r[1]) === String(eventId)).map(r => ({
    id: r[0], eventId: r[1], type: r[2], category: r[3], ceb: r[4], description: r[5],
    beneficiary: r[6], personId: r[7], budget: Number(r[8] || 0), actual: Number(r[9] || 0),
    dueDate: r[10], status: r[11], paidDate: r[12], reference: r[13], attachment: r[14],
    notes: r[15], createdAt: r[16], updatedAt: r[17], rifStatus: r[18], rifRequestedAt: r[19],
    rifCode: r[20], invoiceReceived: r[21] === true, invoiceDate: r[22],
    receiptRequested: r[23] === true, receiptSent: r[24] === true, closedAt: r[25]
  }));
}

function getExpenseProfileForEvent_(event) {
  const type = normalize_(event[APP.CALENDAR_HEADERS.TYPE]);
  const rows = sh_(APP.SHEETS.EVENT_TYPE_CONFIG).getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (normalize_(rows[i][0]) === type) return normalize_(rows[i][4]) || type;
  }
  return type;
}

function getCebOptionsForEvent_(event) {
  const rows = sh_(APP.SHEETS.CEB_CONFIG).getDataRange().getValues();
  const type = normalize_(event[APP.CALENDAR_HEADERS.TYPE]);
  const profile = getExpenseProfileForEvent_(event);
  return rows.slice(1).filter(r => {
    const active = r[4] === true || normalize_(r[4]) === 'SI';
    if (!active) return false;
    const allowed = String(r[2] || '').split(';').map(normalize_).filter(Boolean);
    return allowed.includes('TUTTI') || allowed.includes(type) || allowed.includes(profile);
  }).map(r => ({
    code: r[0],
    description: r[1],
    recordType: r[3] || 'SPESA',
    categories: String(r[6] || '').split(';').map(normalize_).filter(Boolean)
  }));
}

function resolveExpenseCebForCategory_(event, category) {
  const profile = getExpenseProfileForEvent_(event);
  if (profile === 'FOIL ACADEMY') return 'CEB.033';
  const cat = normalize_(category);
  const options = getCebOptionsForEvent_(event).filter(o => normalize_(o.recordType) !== 'RIMBORSO');
  if (!options.length) return '';
  const exact = options.find(o => (o.categories || []).includes(cat));
  if (exact) return exact.code;
  const generic = options.find(o => (o.categories || []).includes('TUTTI'));
  if (generic) return generic.code;
  const travel = options.find(o => String(o.code) === 'CEB.001');
  return travel ? travel.code : options[0].code;
}

function resolveExpenseCeb_(event, requestedCeb, category, recordType) {
  const type = normalize_(recordType || 'SPESA');
  if (type === 'RIMBORSO') return 'CEB.002';
  if (category) return resolveExpenseCebForCategory_(event, category);
  if (getExpenseProfileForEvent_(event) === 'FOIL ACADEMY') return 'CEB.033';
  return requestedCeb || '';
}

function getRefundsForEvent_(eventId) {
  return getExpensesForEvent_(eventId).filter(x => normalize_(x.type) === 'RIMBORSO');
}

function getRefundSituation_(eventId, beneficiary) {
  const key = normalize_(beneficiary);
  const refunds = getRefundsForEvent_(eventId).filter(x => normalize_(x.beneficiary) === key);
  return { refunds: refunds, total: refunds.reduce((s, x) => s + Number(x.actual || 0), 0) };
}

function defaultPaymentDue_(event, payload) {
  const explicit = parseClientDate_(payload.dueDate);
  if (explicit) return explicit;
  const category = normalize_(payload.category);
  const description = normalize_(payload.description);
  const isHotel = category.indexOf('ALLOGGIO') >= 0 || category.indexOf('HOTEL') >= 0 || description.indexOf('HOTEL') >= 0 || description.indexOf('ALLOGGIO') >= 0;
  const start = event[APP.CALENDAR_HEADERS.START];
  if (isHotel && start instanceof Date) return new Date(start.getTime() - 7 * 86400000);
  return '';
}

function expenseNeedsRif_(budget, actual) {
  return Math.max(Number(budget || 0), Number(actual || 0)) > 1000;
}

/**
 * Stati operativi nuovi:
 * - DA PAGARE
 * - PAGATO - FATTURA: fattura già disponibile, pratica chiusa
 * - PAGATO - AFOR: pagamento effettuato, fattura da ricevere dopo l'evento
 * - CHIUSO: usato quando arriva la fattura successiva all'AFOR
 * Gli stati storici restano leggibili per compatibilità.
 */
function expensePaymentMode_(status) {
  const s = normalize_(status);
  if (['PAGATO - AFOR', 'PAGAMENTO AFOR', 'AFOR FATTO'].includes(s)) return 'AFOR';
  if (['PAGATO - FATTURA', 'PAGAMENTO FATTURA'].includes(s)) return 'FATTURA';
  if (['CHIUSO', 'PAGATO CON CC', 'INVIATO IN AMMINISTRAZIONE', 'PAGATO', 'RIMBORSATO'].includes(s)) return 'CHIUSO';
  return '';
}

function normalizeExpensePaymentState_(status, invoiceReceived) {
  let s = normalize_(status || 'DA PAGARE') || 'DA PAGARE';
  let invoice = !!invoiceReceived;
  if (s === 'PAGATO - FATTURA' || s === 'PAGAMENTO FATTURA') {
    s = 'PAGATO - FATTURA';
    invoice = true;
  } else if (s === 'PAGAMENTO AFOR' || s === 'AFOR FATTO') {
    s = 'PAGATO - AFOR';
  }
  if (s === 'PAGATO - AFOR' && invoice) s = 'CHIUSO';
  const mode = expensePaymentMode_(s);
  return {
    status: s,
    invoiceReceived: invoice,
    paid: mode !== '',
    closed: mode === 'FATTURA' || mode === 'CHIUSO'
  };
}

function invoiceDueAfterEvent_(event) {
  const end = event && event[APP.CALENDAR_HEADERS.END];
  if (!(end instanceof Date)) return '';
  return new Date(end.getTime() + 86400000);
}

/**
 * Unico automatismo spese mantenuto in checklist.
 * Dopo un pagamento AFOR prepara "Richiedere fattura", ma la task resta
 * BLOCCATA fino al giorno successivo alla fine dell'evento.
 */
function syncAforInvoiceTask_(expense) {
  if (!expense || normalize_(expense.type) === 'RIMBORSO') return;
  const found = findCalendarEventById_(expense.eventId);
  if (!found) return;

  const sheet = sh_(APP.SHEETS.CHECKLIST);
  const rows = sheet.getDataRange().getValues();
  const key = 'FATTURA_AFOR:' + expense.id;
  const normalizedKey = normalize_(key);
  let rowNumber = 0;
  let existingStatus = '';
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) === String(expense.eventId) && normalize_(rows[i][9]) === normalizedKey) {
      rowNumber = i + 1;
      existingStatus = normalize_(rows[i][6]);
      break;
    }
  }

  const needsInvoice = expensePaymentMode_(expense.status) === 'AFOR' && !expense.invoiceReceived;
  if (!needsInvoice) {
    if (rowNumber && existingStatus !== 'FATTO' && existingStatus !== 'COMPLETATA') {
      sheet.getRange(rowNumber, 7).setValue('FATTO');
      sheet.getRange(rowNumber, 12, 1, 2).setValues([[new Date(), new Date()]]);
    }
    return;
  }

  const ended = eventHasEnded_(found.event);
  const status = ended ? 'DA FARE' : 'BLOCCATA';
  const due = invoiceDueAfterEvent_(found.event);
  const label = expense.beneficiary || expense.description || 'spesa';
  const task = 'Richiedere fattura - ' + label;
  const note = 'Pagamento effettuato tramite AFOR. Richiedere la fattura dopo la conclusione dell’evento.';

  if (rowNumber) {
    // Se l'utente ha già completato la richiesta, non la riapriamo.
    if (existingStatus === 'FATTO' || existingStatus === 'COMPLETATA') return;
    sheet.getRange(rowNumber, 4).setValue(task);
    sheet.getRange(rowNumber, 6).setValue(due || '');
    sheet.getRange(rowNumber, 7).setValue(status);
    sheet.getRange(rowNumber, 11).setValue(note);
    sheet.getRange(rowNumber, 13).setValue(new Date());
    return;
  }

  const existing = getChecklistRowsForEventRaw_(expense.eventId);
  const maxOrder = existing.reduce((m, x) => Math.max(m, Number(x.order || 0)), 0);
  sheet.appendRow([
    'TASK-' + Utilities.getUuid(), expense.eventId, maxOrder + 10, task, 'AMMINISTRAZIONE',
    due || '', status, '', 'AUTO', key, note, '', new Date()
  ]);
}

function syncExpenseTasks_(expense) {
  syncAforInvoiceTask_(expense);
}

function addExpense(eventId, payload) {
  payload = payload || {};
  const description = String(payload.description || '').trim();
  if (!eventId || !description) throw new Error('Descrizione della spesa obbligatoria.');
  const found = findCalendarEventById_(eventId);
  if (!found) throw new Error('Evento non trovato.');

  const budget = Number(payload.budget || 0);
  const actual = Number(payload.actual || 0);
  const rifNeeded = expenseNeedsRif_(budget, actual);
  const rifCode = String(payload.rifCode || '').trim();
  const rifStatus = rifCode ? 'RICEVUTO' : (rifNeeded ? (payload.rifStatus || 'DA RICHIEDERE') : 'NON NECESSARIO');
  const now = new Date();
  const payment = normalizeExpensePaymentState_(payload.status || 'DA PAGARE', payload.invoiceReceived === true || String(payload.invoiceReceived) === 'true');
  const dueDate = defaultPaymentDue_(found.event, payload);
  const recordType = payload.type || 'SPESA';
  const category = payload.category || '';
  const ceb = resolveExpenseCeb_(found.event, payload.ceb, category, recordType);
  const paidDate = payment.paid ? (parseClientDate_(payload.paidDate) || now) : '';

  sh_(APP.SHEETS.EXPENSES).appendRow([
    'SPESA-' + Utilities.getUuid(), eventId, recordType, category, ceb,
    description, payload.beneficiary || '', payload.personId || '', budget, actual, dueDate,
    payment.status, paidDate, payload.reference || '', payload.attachment || '', payload.notes || '',
    now, now, rifStatus, (rifStatus === 'IN ATTESA' || rifStatus === 'RICEVUTO') ? now : '', rifCode,
    payment.invoiceReceived, payment.invoiceReceived ? (parseClientDate_(payload.invoiceDate) || now) : '',
    false, false, payment.closed ? now : ''
  ]);

  const created = getExpensesForEvent_(eventId).slice(-1)[0];
  syncExpenseTasks_(created);
  SpreadsheetApp.flush();
  return getEventPanelData(eventId);
}

function updateExpense(expenseId, payload) {
  payload = payload || {};
  const sheet = sh_(APP.SHEETS.EXPENSES);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(expenseId)) continue;
    const eventId = rows[i][1];
    const found = findCalendarEventById_(eventId);
    if (!found) throw new Error('Evento non trovato.');

    const budget = payload.budget !== undefined ? Number(payload.budget || 0) : Number(rows[i][8] || 0);
    const actual = payload.actual !== undefined ? Number(payload.actual || 0) : Number(rows[i][9] || 0);
    const rifNeeded = expenseNeedsRif_(budget, actual);
    const rifCode = payload.rifCode !== undefined ? String(payload.rifCode || '').trim() : String(rows[i][20] || '').trim();
    let rifStatus = payload.rifStatus !== undefined ? payload.rifStatus : rows[i][18];
    if (rifCode) rifStatus = 'RICEVUTO';
    else if (!rifNeeded) rifStatus = 'NON NECESSARIO';
    else if (!rifStatus || normalize_(rifStatus) === 'NON NECESSARIO') rifStatus = 'DA RICHIEDERE';

    const payment = normalizeExpensePaymentState_(
      payload.status !== undefined ? payload.status : rows[i][11],
      payload.invoiceReceived !== undefined ? !!payload.invoiceReceived : rows[i][21] === true
    );
    const now = new Date();
    const category = payload.category !== undefined ? payload.category : rows[i][3];
    const recordType = payload.type !== undefined ? payload.type : rows[i][2];
    const paidDate = payment.paid
      ? (payload.paidDate !== undefined ? (parseClientDate_(payload.paidDate) || now) : (rows[i][12] || now))
      : '';

    const updated = [
      rows[i][0], eventId, recordType, category,
      resolveExpenseCeb_(found.event, payload.ceb !== undefined ? payload.ceb : rows[i][4], category, recordType),
      payload.description !== undefined ? payload.description : rows[i][5],
      payload.beneficiary !== undefined ? payload.beneficiary : rows[i][6],
      payload.personId !== undefined ? payload.personId : rows[i][7],
      budget, actual,
      payload.dueDate !== undefined ? (parseClientDate_(payload.dueDate) || '') : rows[i][10],
      payment.status, paidDate,
      payload.reference !== undefined ? payload.reference : rows[i][13],
      payload.attachment !== undefined ? payload.attachment : rows[i][14],
      payload.notes !== undefined ? payload.notes : rows[i][15],
      rows[i][16] || now, now, rifStatus,
      (normalize_(rifStatus) === 'IN ATTESA' || normalize_(rifStatus) === 'RICEVUTO') ? (rows[i][19] || now) : '',
      rifCode, payment.invoiceReceived,
      payment.invoiceReceived ? (rows[i][22] || now) : '',
      false, false, payment.closed ? (rows[i][25] || now) : ''
    ];
    sheet.getRange(i + 1, 1, 1, 26).setValues([updated]);
    const expense = expenseObjectFromPanelRow_(updated);
    syncExpenseTasks_(expense);
    SpreadsheetApp.flush();
    return getEventPanelData(eventId);
  }
  throw new Error('Spesa non trovata.');
}

function updateExpenseStatus(expenseId, status, paidDate) {
  return updateExpense(expenseId, { status: status || 'DA PAGARE', paidDate: paidDate || '' });
}

function registerRefund(eventId, beneficiary, amount, paidDate, notes, role) {
  amount = Number(amount || 0);
  role = normalize_(role || 'ATLETA');
  if (!eventId || !beneficiary || !amount) throw new Error('Evento, beneficiario e importo sono obbligatori.');
  const found = findCalendarEventById_(eventId);
  if (!found) throw new Error('Evento non trovato.');

  const situation = getRefundSituation_(eventId, beneficiary);
  const limitData = role === 'ATLETA'
    ? getParticipantRefundLimitForBeneficiary_(eventId, found.event, beneficiary)
    : { value: 0, source: 'SENZA LIMITE' };
  const limit = role === 'ATLETA' ? Number(limitData.value || 0) : 0;
  const projected = situation.total + amount;
  const parsedPaidDate = parseClientDate_(paidDate) || new Date();
  if (situation.refunds.some(x => Number(x.actual || 0) === amount && sameDay_(x.paidDate, parsedPaidDate))) {
    throw new Error('Possibile duplicato: risulta già un rimborso dello stesso importo alla stessa persona per questo evento e data.');
  }

  const warning = limit > 0 && projected > limit
    ? 'ATTENZIONE: il totale rimborsato diventerebbe € ' + projected.toFixed(2) + ' rispetto al massimale atleta di € ' + limit.toFixed(2)
    : '';
  const now = new Date();
  const participant = limitData.participant || null;
  sh_(APP.SHEETS.EXPENSES).appendRow([
    'RIM-' + Utilities.getUuid(), eventId, 'RIMBORSO', 'RIMBORSI ' + role, 'CEB.002',
    'Rimborso spese ' + role.toLowerCase(), beneficiary, participant ? participant.personId || '' : '',
    0, amount, '', 'CHIUSO', parsedPaidDate, '', '', notes || '', now, now,
    'NON NECESSARIO', '', '', true, now, false, false, now
  ]);
  SpreadsheetApp.flush();
  return clientSafe_({
    success: true, previousTotal: situation.total, newTotal: projected, limit: limit,
    residual: limit > 0 ? limit - projected : null, warning: warning, data: getEventPanelData(eventId)
  });
}

function sameDay_(a, b) {
  if (!(a instanceof Date) || !(b instanceof Date)) return false;
  return Utilities.formatDate(a, APP.TZ, 'yyyyMMdd') === Utilities.formatDate(b, APP.TZ, 'yyyyMMdd');
}
