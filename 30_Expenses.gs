function getExpensesForEvent_(eventId) {
  const rows = sh_(APP.SHEETS.EXPENSES).getDataRange().getValues();
  return rows.slice(1).filter(r => String(r[1]) === String(eventId)).map(r => ({
    id: r[0],
    eventId: r[1],
    type: r[2],
    category: r[3],
    ceb: r[4],
    description: r[5],
    beneficiary: r[6],
    personId: r[7],
    budget: Number(r[8] || 0),
    actual: Number(r[9] || 0),
    dueDate: r[10],
    status: r[11],
    paidDate: r[12],
    reference: r[13],
    attachment: r[14],
    notes: r[15],
    createdAt: r[16],
    updatedAt: r[17]
  }));
}

function getCebOptionsForEvent_(event) {
  const rows = sh_(APP.SHEETS.CEB_CONFIG).getDataRange().getValues();
  const type = normalize_(event[APP.CALENDAR_HEADERS.TYPE]);
  return rows.slice(1).filter(r => {
    const active = r[4] === true || normalize_(r[4]) === 'SI';
    if (!active) return false;
    const allowed = String(r[2] || '').split(';').map(normalize_).filter(Boolean);
    return allowed.includes('TUTTI') || allowed.includes(type);
  }).map(r => ({ code: r[0], description: r[1], recordType: r[3] || 'SPESA' }));
}

function getRefundsForEvent_(eventId) {
  return getExpensesForEvent_(eventId).filter(x => normalize_(x.type) === 'RIMBORSO');
}

function getRefundSituation_(eventId, beneficiary) {
  const key = normalize_(beneficiary);
  const refunds = getRefundsForEvent_(eventId).filter(x => normalize_(x.beneficiary) === key);
  const total = refunds.reduce((s, x) => s + Number(x.actual || 0), 0);
  return { refunds: refunds, total: total };
}

function addExpense(eventId, payload) {
  payload = payload || {};
  const description = String(payload.description || '').trim();
  if (!eventId || !description) throw new Error('Descrizione della spesa obbligatoria.');

  const sheet = sh_(APP.SHEETS.EXPENSES);
  const id = 'SPESA-' + Utilities.getUuid();
  const now = new Date();
  sheet.appendRow([
    id,
    eventId,
    payload.type || 'SPESA',
    payload.category || '',
    payload.ceb || '',
    description,
    payload.beneficiary || '',
    payload.personId || '',
    Number(payload.budget || 0),
    Number(payload.actual || 0),
    parseClientDate_(payload.dueDate) || '',
    payload.status || 'DA DEFINIRE',
    parseClientDate_(payload.paidDate) || '',
    payload.reference || '',
    payload.attachment || '',
    payload.notes || '',
    now,
    now
  ]);
  refreshExpenseSummary_(eventId);
  return getEventPanelData(eventId);
}

function updateExpenseStatus(expenseId, status, paidDate) {
  const sheet = sh_(APP.SHEETS.EXPENSES);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(expenseId)) continue;
    const eventId = rows[i][1];
    sheet.getRange(i + 1, 12).setValue(status || 'DA DEFINIRE');
    sheet.getRange(i + 1, 13).setValue(parseClientDate_(paidDate) || (normalize_(status) === 'PAGATO' ? new Date() : ''));
    sheet.getRange(i + 1, 18).setValue(new Date());
    refreshExpenseSummary_(eventId);
    return getEventPanelData(eventId);
  }
  throw new Error('Spesa non trovata.');
}

function refreshExpenseSummary_(eventId) {
  const found = findCalendarEventById_(eventId);
  if (!found) return;
  const items = getExpensesForEvent_(eventId);
  const budget = items.reduce((s, x) => s + Number(x.budget || 0), 0);
  const actual = items.reduce((s, x) => s + Number(x.actual || 0), 0);
  const toPay = items.filter(x => !['PAGATO', 'RIMBORSATO'].includes(normalize_(x.status)))
    .reduce((s, x) => s + Number(x.actual || x.budget || 0), 0);
  const cal = sh_(APP.SHEETS.CALENDAR);
  const map = headerMap_(cal);
  cal.getRange(found.row, map[APP.CALENDAR_HEADERS.BUDGET]).setValue(budget);
  cal.getRange(found.row, map[APP.CALENDAR_HEADERS.ACTUAL]).setValue(actual);
  cal.getRange(found.row, map[APP.CALENDAR_HEADERS.TO_PAY]).setValue(toPay);
}

function registerRefund(eventId, beneficiary, amount, paidDate, notes) {
  amount = Number(amount || 0);
  if (!eventId || !beneficiary || !amount) throw new Error('Evento, beneficiario e importo sono obbligatori.');

  const found = findCalendarEventById_(eventId);
  if (!found) throw new Error('Evento non trovato.');
  const limitData = getEventRefundLimit_(eventId, found.event);
  const situation = getRefundSituation_(eventId, beneficiary);
  const limit = Number(limitData.value || 0);
  const projected = situation.total + amount;
  const parsedPaidDate = parseClientDate_(paidDate) || new Date();
  const exactDuplicate = situation.refunds.some(x => Number(x.actual || 0) === amount && sameDay_(x.paidDate, parsedPaidDate));

  if (exactDuplicate) {
    throw new Error('Possibile duplicato: risulta già un rimborso dello stesso importo alla stessa persona per questo evento e data.');
  }

  const warning = limit > 0 && projected > limit
    ? 'ATTENZIONE: il totale rimborsato diventerebbe € ' + projected.toFixed(2) + ' rispetto al massimale di € ' + limit.toFixed(2)
    : '';

  const now = new Date();
  sh_(APP.SHEETS.EXPENSES).appendRow([
    'RIM-' + Utilities.getUuid(), eventId, 'RIMBORSO', 'RIMBORSI', 'CEB.002',
    'Rimborso spese', beneficiary, '', 0, amount, '', 'PAGATO', parsedPaidDate,
    '', '', notes || '', now, now
  ]);
  refreshExpenseSummary_(eventId);

  return clientSafe_({
    success: true,
    previousTotal: situation.total,
    newTotal: projected,
    limit: limit,
    residual: limit > 0 ? limit - projected : null,
    warning: warning,
    data: getEventPanelData(eventId)
  });
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
  SpreadsheetApp.getUi().showModalDialog(html.evaluate().setWidth(520).setHeight(560), 'Registra rimborso');
}
