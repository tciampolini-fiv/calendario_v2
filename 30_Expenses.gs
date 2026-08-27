function getExpensesForEvent_(eventId) {
  const rows = sh_(APP.SHEETS.EXPENSES).getDataRange().getValues();
  return rows.slice(1).filter(r => String(r[1]) === String(eventId)).map(r => ({
    id: r[0], type: r[3], category: r[4], ceb: r[5], description: r[6], beneficiary: r[7],
    budget: Number(r[8] || 0), actual: Number(r[9] || 0), dueDate: r[11], paidDate: r[12], status: r[13], notes: r[14]
  }));
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

function registerRefund(eventId, beneficiary, amount, paidDate, notes) {
  amount = Number(amount || 0);
  if (!eventId || !beneficiary || !amount) throw new Error('Evento, beneficiario e importo sono obbligatori.');
  const data = getEventPanelData(eventId);
  const situation = getRefundSituation_(eventId, beneficiary);
  const limit = Number((data.refundLimit && data.refundLimit.value) || 0);
  const projected = situation.total + amount;
  const exactDuplicate = situation.refunds.some(x => Number(x.actual || 0) === amount && sameDay_(x.paidDate, paidDate));
  if (exactDuplicate) {
    throw new Error('Possibile duplicato: risulta gia un rimborso dello stesso importo alla stessa persona per questo evento e data.');
  }
  const warning = limit > 0 && projected > limit
    ? 'ATTENZIONE: il totale rimborsato diventerebbe € ' + projected.toFixed(2) + ' rispetto al massimale di € ' + limit.toFixed(2)
    : '';
  const sheet = sh_(APP.SHEETS.EXPENSES);
  const id = 'RIM-' + Utilities.formatDate(new Date(), APP.TZ, 'yyyyMMdd-HHmmss');
  sheet.appendRow([
    id, eventId, new Date(), 'RIMBORSO', 'Indennita, diarie e rimborsi', 'CEB.002',
    'Rimborso spese', beneficiary, '', amount, '', '', paidDate || new Date(), 'PAGATO', notes || '', '', '', new Date()
  ]);
  return {
    success: true,
    previousTotal: situation.total,
    newTotal: projected,
    limit: limit,
    residual: limit > 0 ? limit - projected : null,
    warning: warning
  };
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
