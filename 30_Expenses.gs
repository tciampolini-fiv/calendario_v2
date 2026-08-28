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
  }).map(r => ({ code: r[0], description: r[1], recordType: r[3] || 'SPESA' }));
}

function resolveExpenseCeb_(event, requestedCeb) {
  const profile = getExpenseProfileForEvent_(event);
  if (profile === 'FOIL ACADEMY') return 'CEB.033';
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
  const status = payload.status || 'DA DEFINIRE';
  const invoiceReceived = payload.invoiceReceived === true || String(payload.invoiceReceived) === 'true';
  const receiptRequested = payload.receiptRequested === true || String(payload.receiptRequested) === 'true';
  const receiptSent = payload.receiptSent === true || String(payload.receiptSent) === 'true';
  const dueDate = defaultPaymentDue_(found.event, payload);
  const ceb = resolveExpenseCeb_(found.event, payload.ceb);

  sh_(APP.SHEETS.EXPENSES).appendRow([
    'SPESA-' + Utilities.getUuid(), eventId, payload.type || 'SPESA', payload.category || '', ceb,
    description, payload.beneficiary || '', payload.personId || '', budget, actual, dueDate,
    status, parseClientDate_(payload.paidDate) || '', payload.reference || '', payload.attachment || '',
    payload.notes || '', now, now, rifStatus,
    rifStatus === 'IN ATTESA' || rifStatus === 'RICEVUTO' ? now : '', rifCode,
    invoiceReceived, invoiceReceived ? (parseClientDate_(payload.invoiceDate) || now) : '',
    receiptRequested, receiptSent, normalize_(status) === 'CHIUSO' ? now : ''
  ]);

  const items = getExpensesForEvent_(eventId);
  const created = items[items.length - 1];
  syncExpenseTasks_(created);
  refreshExpenseSummary_(eventId);
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

    const status = payload.status !== undefined ? payload.status : rows[i][11];
    const invoiceReceived = payload.invoiceReceived !== undefined ? !!payload.invoiceReceived : rows[i][21] === true;
    const receiptRequested = payload.receiptRequested !== undefined ? !!payload.receiptRequested : rows[i][23] === true;
    const receiptSent = payload.receiptSent !== undefined ? !!payload.receiptSent : rows[i][24] === true;
    const dueDate = payload.dueDate !== undefined ? (parseClientDate_(payload.dueDate) || '') : rows[i][10];
    const now = new Date();

    const values = [[
      rows[i][0], eventId, payload.type !== undefined ? payload.type : rows[i][2],
      payload.category !== undefined ? payload.category : rows[i][3],
      resolveExpenseCeb_(found.event, payload.ceb !== undefined ? payload.ceb : rows[i][4]),
      payload.description !== undefined ? payload.description : rows[i][5],
      payload.beneficiary !== undefined ? payload.beneficiary : rows[i][6],
      payload.personId !== undefined ? payload.personId : rows[i][7],
      budget, actual, dueDate, status,
      payload.paidDate !== undefined ? (parseClientDate_(payload.paidDate) || '') : rows[i][12],
      payload.reference !== undefined ? payload.reference : rows[i][13],
      payload.attachment !== undefined ? payload.attachment : rows[i][14],
      payload.notes !== undefined ? payload.notes : rows[i][15],
      rows[i][16] || now, now, rifStatus,
      (normalize_(rifStatus) === 'IN ATTESA' || normalize_(rifStatus) === 'RICEVUTO') ? (rows[i][19] || now) : '',
      rifCode, invoiceReceived,
      invoiceReceived ? (payload.invoiceDate ? parseClientDate_(payload.invoiceDate) : (rows[i][22] || now)) : '',
      receiptRequested, receiptSent,
      normalize_(status) === 'CHIUSO' ? (rows[i][25] || now) : ''
    ]];
    sheet.getRange(i + 1, 1, 1, 26).setValues(values);
    const expense = getExpensesForEvent_(eventId).find(x => String(x.id) === String(expenseId));
    syncExpenseTasks_(expense);
    refreshExpenseSummary_(eventId);
    return getEventPanelData(eventId);
  }
  throw new Error('Spesa non trovata.');
}

function updateExpenseStatus(expenseId, status, paidDate) {
  return updateExpense(expenseId, { status: status || 'DA DEFINIRE', paidDate: paidDate || '' });
}

function syncExpenseTasks_(expense) {
  if (!expense) return;
  const eventId = expense.eventId;
  const label = expense.beneficiary || expense.description || 'spesa';
  const paidOperationally = ['AFOR FATTO','PAGATO CON CC','INVIATO IN AMMINISTRAZIONE','CHIUSO'].includes(normalize_(expense.status));
  const paymentOpen = !!expense.dueDate && !paidOperationally;
  syncAutoChecklistTask_(eventId, 'PAY:' + expense.id, 'Pagamento ' + label, expense.dueDate || '', paymentOpen, expense.description || '');

  const rifOpen = expenseNeedsRif_(expense.budget, expense.actual) && normalize_(expense.rifStatus) !== 'RICEVUTO';
  syncAutoChecklistTask_(eventId, 'RIF:' + expense.id, 'Codice RIF - ' + label, '', rifOpen, rifOpen ? 'Spesa superiore a € 1.000' : '');

  const invoiceOpen = paidOperationally && !expense.invoiceReceived && normalize_(expense.type) !== 'RIMBORSO';
  syncAutoChecklistTask_(eventId, 'FATTURA:' + expense.id, 'Richiedere fattura - ' + label, '', invoiceOpen, 'Chiudere quando la fattura è stata ricevuta');

  const receiptOpen = expense.receiptRequested && !expense.receiptSent;
  syncAutoChecklistTask_(eventId, 'CONTABILE:' + expense.id, 'Recuperare e inviare contabile - ' + label, '', receiptOpen, 'Richiedere la contabile all’amministrazione e inoltrarla al fornitore');
  const found = findCalendarEventById_(eventId);
  if (found) refreshEventSummary_(eventId, found.row);
}

function refreshExpenseSummary_(eventId) {
  const found = findCalendarEventById_(eventId);
  if (!found) return;
  const items = getExpensesForEvent_(eventId);
  const budget = items.reduce((s, x) => s + Number(x.budget || 0), 0);
  const actual = items.reduce((s, x) => s + Number(x.actual || 0), 0);
  const operationallyDone = x => ['AFOR FATTO','PAGATO CON CC','INVIATO IN AMMINISTRAZIONE','CHIUSO','PAGATO','RIMBORSATO'].includes(normalize_(x.status));
  const toPay = items.filter(x => !operationallyDone(x)).reduce((s, x) => s + Number(x.actual || x.budget || 0), 0);
  const cal = sh_(APP.SHEETS.CALENDAR);
  const map = headerMap_(cal);
  cal.getRange(found.row, map[APP.CALENDAR_HEADERS.BUDGET]).setValue(budget);
  cal.getRange(found.row, map[APP.CALENDAR_HEADERS.ACTUAL]).setValue(actual);
  cal.getRange(found.row, map[APP.CALENDAR_HEADERS.TO_PAY]).setValue(toPay);
}

function registerRefund(eventId, beneficiary, amount, paidDate, notes, role) {
  amount = Number(amount || 0);
  role = normalize_(role || 'ATLETA');
  if (!eventId || !beneficiary || !amount) throw new Error('Evento, beneficiario e importo sono obbligatori.');
  const found = findCalendarEventById_(eventId);
  if (!found) throw new Error('Evento non trovato.');

  const situation = getRefundSituation_(eventId, beneficiary);
  const limitData = getEventRefundLimit_(eventId, found.event);
  const limit = role === 'ATLETA' ? Number(limitData.value || 0) : 0;
  const projected = situation.total + amount;
  const parsedPaidDate = parseClientDate_(paidDate) || new Date();
  const exactDuplicate = situation.refunds.some(x => Number(x.actual || 0) === amount && sameDay_(x.paidDate, parsedPaidDate));
  if (exactDuplicate) throw new Error('Possibile duplicato: risulta già un rimborso dello stesso importo alla stessa persona per questo evento e data.');

  const warning = limit > 0 && projected > limit
    ? 'ATTENZIONE: il totale rimborsato diventerebbe € ' + projected.toFixed(2) + ' rispetto al massimale atleta di € ' + limit.toFixed(2)
    : '';
  const now = new Date();
  sh_(APP.SHEETS.EXPENSES).appendRow([
    'RIM-' + Utilities.getUuid(), eventId, 'RIMBORSO', 'RIMBORSI ' + role, 'CEB.002',
    'Rimborso spese ' + role.toLowerCase(), beneficiary, '', 0, amount, '', 'CHIUSO', parsedPaidDate,
    '', '', notes || '', now, now, 'NON NECESSARIO', '', '', true, now, false, false, now
  ]);
  refreshExpenseSummary_(eventId);
  return clientSafe_({ success:true, previousTotal:situation.total, newTotal:projected, limit:limit,
    residual:limit > 0 ? limit - projected : null, warning:warning, data:getEventPanelData(eventId) });
}

function sameDay_(a, b) {
  if (!(a instanceof Date) || !(b instanceof Date)) return false;
  return Utilities.formatDate(a, APP.TZ, 'yyyyMMdd') === Utilities.formatDate(b, APP.TZ, 'yyyyMMdd');
}

function openRefundDialog() {
  const event = selectedEvent_();
  const id = ensureEventId_(event);
  const html = HtmlService.createTemplateFromFile('RefundDialog');
  html.eventId = id;
  SpreadsheetApp.getUi().showModalDialog(html.evaluate().setWidth(520).setHeight(600), 'Registra rimborso');
}
