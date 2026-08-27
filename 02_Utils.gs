function ss_() {
  return SpreadsheetApp.openById(APP.SPREADSHEET_ID);
}

function sh_(name) {
  const sheet = ss_().getSheetByName(name);
  if (!sheet) throw new Error('Foglio non trovato: ' + name);
  return sheet;
}

function headerMap_(sheet) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  return headers.reduce((m, h, i) => {
    if (h) m[h.trim()] = i + 1;
    return m;
  }, {});
}

function selectedEvent_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName() !== APP.SHEETS.CALENDAR) {
    throw new Error('Seleziona una riga nel foglio Calendario.');
  }
  const row = sheet.getActiveRange().getRow();
  if (row < 2) throw new Error('Seleziona una riga evento.');
  const map = headerMap_(sheet);
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const event = {};
  Object.keys(map).forEach(h => event[h] = values[map[h] - 1]);
  event._row = row;
  return event;
}

function ensureEventId_(event) {
  const idHeader = APP.CALENDAR_HEADERS.ID;
  if (event[idHeader]) return String(event[idHeader]);
  const sheet = sh_(APP.SHEETS.CALENDAR);
  const map = headerMap_(sheet);
  const id = 'EVT-' + Utilities.formatDate(new Date(), APP.TZ, 'yyyyMMdd-HHmmss') + '-' + event._row;
  sheet.getRange(event._row, map[idHeader]).setValue(id);
  event[idHeader] = id;
  return id;
}

function normalize_(value) {
  return String(value || '').trim().toUpperCase();
}

function samePersonKey_(name, surname) {
  return normalize_(name) + '|' + normalize_(surname);
}

/**
 * google.script.run non puo trasferire Date native tra server e browser.
 * Converte ricorsivamente Date in stringhe ISO e lascia invariati gli altri valori.
 */
function clientSafe_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, APP.TZ, "yyyy-MM-dd'T'HH:mm:ssXXX");
  }
  if (Array.isArray(value)) return value.map(clientSafe_);
  if (value && typeof value === 'object') {
    const out = {};
    Object.keys(value).forEach(k => out[k] = clientSafe_(value[k]));
    return out;
  }
  return value;
}

/** Converte una data ricevuta dal browser (yyyy-MM-dd o ISO) in Date. */
function parseClientDate_(value) {
  if (!value) return null;
  if (value instanceof Date) return value;
  const text = String(value).trim();
  if (!text) return null;
  const simple = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text);
  if (simple) {
    return new Date(Number(simple[1]), Number(simple[2]) - 1, Number(simple[3]), 12, 0, 0, 0);
  }
  const d = new Date(text);
  if (isNaN(d.getTime())) throw new Error('Data non valida: ' + text);
  return d;
}
