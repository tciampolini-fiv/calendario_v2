const EVENT_DRIVE = Object.freeze({
  ROOT: '1bV_kDC3MIdB6yy59pVUYJC_Xqwo3lAfJ',
  TRAINING_RACES: '1KSTURCZZvbbkZWVx4ZNf--Oqvj4Z6fYZ',
  TEMPLATES: '1U1T0TJOIVQdfajDDF0_JvP1Dcwm7Pelc',
  TESTS: '1u5r27NGwQGsSjXG0fdLMAJKhTbHkfmIV',
  MISSIONS: '1b4sk5piEy4LcakZyQlV_SJNI95R4Lmvp',
  FOIL_ACADEMY: '1qQQ3RkEPzDAlGubp2nCFEgiqF5X_ZyF1'
});

function createWorkFolderForSelectedEvent() {
  const event = selectedEvent_();
  const eventId = ensureEventId_(event);
  const result = createWorkFolderForEvent_(eventId, event, event._row);
  SpreadsheetApp.getUi().alert(
    result.created ? 'Cartella creata' : 'Cartella già presente',
    result.url,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return result;
}

function createWorkFolderForPanel(eventId, row) {
  const event = getEventFromPanelRow_(Number(row), String(eventId));
  return clientSafe_(createWorkFolderForEvent_(eventId, event, Number(row)));
}

function createWorkFolderForEvent_(eventId, event, row) {
  const sheet = sh_(APP.SHEETS.CALENDAR);
  const map = headerMap_(sheet);
  const folderCol = map[APP.CALENDAR_HEADERS.FOLDER];
  if (!folderCol) throw new Error('Colonna CARTELLA non trovata.');

  const existingUrl = String(event[APP.CALENDAR_HEADERS.FOLDER] || '').trim();
  if (existingUrl) {
    const existingId = extractDriveId_(existingUrl);
    if (existingId) {
      try {
        const folder = DriveApp.getFolderById(existingId);
        return { success: true, created: false, eventId: eventId, name: folder.getName(), url: folder.getUrl() };
      } catch (e) {
        // Se il link non è più valido, proseguiamo creando una nuova cartella.
      }
    }
  }

  const destination = resolveWorkFolderDestination_(event);
  const parent = destination.classSubfolder
    ? getOrCreateFolderByNameV2_(DriveApp.getFolderById(destination.folderId), destination.classSubfolder)
    : DriveApp.getFolderById(destination.folderId);

  const folderName = buildEventFolderName_(event);
  const folder = getOrCreateFolderByNameV2_(parent, folderName);
  const url = folder.getUrl();
  sheet.getRange(Number(row), folderCol).setValue(url);

  return { success: true, created: true, eventId: eventId, name: folder.getName(), url: url };
}

function resolveWorkFolderDestination_(event) {
  const type = normalize_(event[APP.CALENDAR_HEADERS.TYPE]);
  const cls = String(event[APP.CALENDAR_HEADERS.CLASS] || '').trim();

  if (type === 'FOIL ACADEMY') return { folderId: EVENT_DRIVE.FOIL_ACADEMY, classSubfolder: '' };
  if (type === 'TEST FISICI') return { folderId: EVENT_DRIVE.TESTS, classSubfolder: '' };
  if (type === 'RIUNIONE/MISSIONE' || type.indexOf('RIUNION') >= 0 || type.indexOf('MISSION') >= 0) {
    return { folderId: EVENT_DRIVE.MISSIONS, classSubfolder: '' };
  }
  if (type.indexOf('ALLENAMENTO') >= 0 || type === 'STAGE' || type.indexOf('REGATA') >= 0 || type.indexOf('OSSERVAZIONE') >= 0) {
    return { folderId: EVENT_DRIVE.TRAINING_RACES, classSubfolder: cls || 'SENZA CLASSE' };
  }
  return { folderId: EVENT_DRIVE.ROOT, classSubfolder: '' };
}

function buildEventFolderName_(event) {
  const type = toTitleCaseV2_(String(event[APP.CALENDAR_HEADERS.TYPE] || '').trim());
  const location = String(event[APP.CALENDAR_HEADERS.LOCATION] || '').trim();
  const dateLabel = formatEventDateRangeV2_(event[APP.CALENDAR_HEADERS.START], event[APP.CALENDAR_HEADERS.END]);
  return [type, location, dateLabel].filter(Boolean).join(' ').replace(/\s{2,}/g, ' ').trim() || 'Evento';
}

function formatEventDateRangeV2_(start, end) {
  if (!(start instanceof Date)) return '';
  const tz = APP.TZ || Session.getScriptTimeZone();
  const s = Utilities.formatDate(start, tz, 'dd-MM-yyyy');
  if (!(end instanceof Date)) return s;
  const e = Utilities.formatDate(end, tz, 'dd-MM-yyyy');
  return s === e ? s : s + ' - ' + e;
}

function getOrCreateFolderByNameV2_(parent, name) {
  const folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

function extractDriveId_(value) {
  const match = String(value || '').match(/[-\w]{25,}/);
  return match ? match[0] : '';
}

function toTitleCaseV2_(value) {
  return String(value || '').toLowerCase().replace(/(^|\s|[-/])([a-zà-öø-ÿ])/g, function(_, p, c) {
    return p + c.toUpperCase();
  });
}
