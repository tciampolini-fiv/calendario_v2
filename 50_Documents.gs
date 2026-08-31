const EVENT_DRIVE = Object.freeze({
  ROOT: '1bV_kDC3MIdB6yy59pVUYJC_Xqwo3lAfJ',
  TRAINING_RACES: '1KSTURCZZvbbkZWVx4ZNf--Oqvj4Z6fYZ',
  TEMPLATES: '1U1T0TJOIVQdfajDDF0_JvP1Dcwm7Pelc',
  TESTS: '1u5r27NGwQGsSjXG0fdLMAJKhTbHkfmIV',
  MISSIONS: '1b4sk5piEy4LcakZyQlV_SJNI95R4Lmvp',
  FOIL_ACADEMY: '1qQQ3RkEPzDAlGubp2nCFEgiqF5X_ZyF1'
});

const DOCUMENT_TEMPLATES = Object.freeze([
  { key:'CONV_ATLETI', label:'Convocazione atleti', id:'1SpOMrpDwe8aTW9QAyby865WufnD40AsPx6t8zm6Yw3E' },
  { key:'CONV_TECNICO', label:'Convocazione tecnico', id:'1Xb61H5TQ0avz8n5_THn-BSd1Xd4YBtmOD3Lu2uwUuJw' },
  { key:'GOMMONE', label:'Richiesta gommone alla Zona', id:'1uvK4J8WrWekpMrkPTkpUvsdYTggBhwzCDJaJmyIhUko' },
  { key:'OSPITALITA', label:'Richiesta ospitalità circolo', id:'1NuheJqcwtLxp8muTdbnpcZc22Z8HJgA-jbdt_Sq6ruk' },
  { key:'RINGRAZIAMENTO', label:'Ringraziamento circolo', id:'11Fqn8U3Hs8MgJPH0tQTKDEozmWDftPtmELoLjgMFg7c' }
]);

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
        return { success:true, created:false, eventId:eventId, name:folder.getName(), url:folder.getUrl(), folderId:folder.getId() };
      } catch (e) {}
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
  return { success:true, created:true, eventId:eventId, name:folder.getName(), url:url, folderId:folder.getId() };
}

function generateDocumentsForSelectedEvent() {
  const ui = SpreadsheetApp.getUi();
  const event = selectedEvent_();
  const eventId = ensureEventId_(event);
  const folderResult = createWorkFolderForEvent_(eventId, event, event._row);

  let prompt = 'Scegli i documenti da creare, separando i numeri con una virgola:\n\n';
  DOCUMENT_TEMPLATES.forEach((t, i) => prompt += (i + 1) + ' = ' + t.label + '\n');
  const response = ui.prompt('Genera documenti', prompt + '\nEsempio: 1,2,4', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const selected = parseDocumentSelection_(response.getResponseText());
  if (!selected.length) {
    ui.alert('Nessun documento valido selezionato.');
    return;
  }

  const result = generateDocumentsForEvent_(eventId, event, folderResult.folderId, selected);
  ui.alert(
    'Documenti creati',
    result.created.length + ' documento/i Google modificabile/i creato/i nella cartella evento.\n\n' +
    result.created.map(x => '• ' + x.name).join('\n'),
    ui.ButtonSet.OK
  );
  return result;
}

function generateDocumentsForEvent_(eventId, event, folderId, templateKeys) {
  const folder = DriveApp.getFolderById(folderId);
  const created = [];

  templateKeys.forEach(key => {
    const spec = DOCUMENT_TEMPLATES.find(x => x.key === key);
    if (!spec) return;
    const replacements = buildEventDocumentReplacements_(eventId, event, spec.key);
    const template = DriveApp.getFileById(spec.id);
    let name = buildGeneratedDocumentName_(spec.key, event, replacements);
    name = uniqueDocumentName_(folder, name);
    const copy = template.makeCopy(name, folder);
    fillGoogleDocPlaceholdersV2_(copy.getId(), replacements);
    registerGeneratedDocument_(eventId, spec, copy, name);
    created.push({ key:spec.key, name:name, id:copy.getId(), url:copy.getUrl() });
  });

  return { success:true, folderUrl:folder.getUrl(), created:created };
}

function buildEventDocumentReplacements_(eventId, event, templateKey) {
  const start = event[APP.CALENDAR_HEADERS.START];
  const end = event[APP.CALENDAR_HEADERS.END];
  const participants = getParticipantsForEvent_(eventId);
  const athletes = participants.filter(p => !normalize_(p.role) || normalize_(p.role) === 'ATLETA');
  const athleteLines = athletes.map(p => {
    const name = [p.name, p.surname].filter(Boolean).join(' ').trim();
    return p.club ? name + ' – ' + p.club : name;
  }).filter(Boolean);
  const technicians = String(event[APP.CALENDAR_HEADERS.TECHNICIANS] || '').trim();
  const refund = getEventRefundLimit_(eventId, event);
  const today = new Date();
  const zone = String(event[APP.CALENDAR_HEADERS.ZONE] || '').trim();

  const reps = {
    'tipo': toTitleCaseV2_(event[APP.CALENDAR_HEADERS.TYPE] || ''),
    'classe': String(event[APP.CALENDAR_HEADERS.CLASS] || '').trim(),
    'luogo': String(event[APP.CALENDAR_HEADERS.LOCATION] || '').trim(),
    'localita': String(event[APP.CALENDAR_HEADERS.LOCATION] || '').trim(),
    'zona': zone,
    'circolo': String(event[APP.CALENDAR_HEADERS.CLUB] || '').trim(),
    'tecnico': technicians,
    'lista tecnici': technicians,
    'lista atleti': athleteLines.join('\n'),
    'numero atleti': athletes.length ? String(athletes.length) : '',
    'hotel/struttura': String(event[APP.CALENDAR_HEADERS.LODGING] || '').trim(),
    'data inizio': formatDocumentDate_(start),
    'data fine': formatDocumentDate_(end),
    'data in': formatDocumentDate_(start),
    'data out': formatDocumentDate_(end),
    'date': formatDocumentDateRange_(start, end),
    'data di oggi': formatDocumentDate_(today),
    'data oggi': formatDocumentDate_(today),
    'rimborso massimo': refund && Number(refund.value || 0) > 0 ? formatDocumentMoney_(refund.value) : '',
    'note evento': String(event[APP.CALENDAR_HEADERS.EVENT] || '').trim()
  };

  // Nei modelli storici {{N}} ha due significati diversi.
  // La sostituzione va quindi definita per singolo modello.
  if (templateKey === 'GOMMONE') reps['N'] = zone;
  if (templateKey === 'OSPITALITA') reps['N'] = athletes.length ? String(athletes.length) : '';

  return reps;
}

function fillGoogleDocPlaceholdersV2_(docId, replacements) {
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  Object.keys(replacements).forEach(key => {
    const value = replacements[key];
    if (value === '' || value === null || value === undefined) return;
    body.replaceText('(?i)\\{\\{' + escapeRegexV2_(key) + '\\}\\}', safeDocReplacementV2_(value));
  });
  if (replacements.classe) body.replaceText('(?i)\\{classe\\}\\}', safeDocReplacementV2_(replacements.classe));
  doc.saveAndClose();
}

function buildGeneratedDocumentName_(key, event, replacements) {
  const type = replacements.tipo || 'Evento';
  const cls = replacements.classe;
  const place = replacements.luogo;
  const dates = replacements.date;
  if (key === 'CONV_ATLETI') return ['Convocazione atleti', type, cls, place, dates].filter(Boolean).join(' - ');
  if (key === 'CONV_TECNICO') return ['Convocazione tecnico', replacements.tecnico, type, cls, dates].filter(Boolean).join(' - ');
  if (key === 'GOMMONE') return ['Richiesta gommone', replacements.zona ? replacements.zona + ' Zona' : '', type, cls, dates].filter(Boolean).join(' - ');
  if (key === 'OSPITALITA') return ['Richiesta ospitalità', replacements.circolo, type, cls, dates].filter(Boolean).join(' - ');
  if (key === 'RINGRAZIAMENTO') return ['Ringraziamento', replacements.circolo, type, cls, dates].filter(Boolean).join(' - ');
  return [type, cls, dates].filter(Boolean).join(' - ');
}

function uniqueDocumentName_(folder, baseName) {
  if (!folder.getFilesByName(baseName).hasNext()) return baseName;
  const stamp = Utilities.formatDate(new Date(), APP.TZ || Session.getScriptTimeZone(), 'yyyyMMdd-HHmm');
  return baseName + ' - ' + stamp;
}

function registerGeneratedDocument_(eventId, spec, file, name) {
  const sheet = sh_(APP.SHEETS.DOCUMENTS);
  sheet.appendRow([
    'DOC-' + Utilities.getUuid(), eventId, spec.key, spec.id, name,
    file.getId(), file.getUrl(), 'BOZZA MODIFICABILE', new Date(), ''
  ]);
}

function parseDocumentSelection_(text) {
  const seen = {};
  return String(text || '').split(/[;,\s]+/).map(x => parseInt(x, 10)).filter(n => {
    if (!n || n < 1 || n > DOCUMENT_TEMPLATES.length || seen[n]) return false;
    seen[n] = true; return true;
  }).map(n => DOCUMENT_TEMPLATES[n - 1].key);
}

function formatDocumentDate_(value) {
  if (!(value instanceof Date)) return '';
  const months = ['gennaio','febbraio','marzo','aprile','maggio','giugno','luglio','agosto','settembre','ottobre','novembre','dicembre'];
  return value.getDate() + ' ' + months[value.getMonth()] + ' ' + value.getFullYear();
}

function formatDocumentDateRange_(start, end) {
  if (!(start instanceof Date)) return '';
  if (!(end instanceof Date)) return formatDocumentDate_(start);
  if (start.getTime() === end.getTime()) return formatDocumentDate_(start);
  if (start.getFullYear() === end.getFullYear() && start.getMonth() === end.getMonth()) {
    const months = ['gennaio','febbraio','marzo','aprile','maggio','giugno','luglio','agosto','settembre','ottobre','novembre','dicembre'];
    return start.getDate() + '-' + end.getDate() + ' ' + months[start.getMonth()] + ' ' + start.getFullYear();
  }
  return formatDocumentDate_(start) + ' - ' + formatDocumentDate_(end);
}

function formatDocumentMoney_(value) {
  const n = Number(value || 0);
  if (!isFinite(n)) return '';
  return n.toLocaleString('it-IT', { minimumFractionDigits:0, maximumFractionDigits:2 });
}

function resolveWorkFolderDestination_(event) {
  const type = normalize_(event[APP.CALENDAR_HEADERS.TYPE]);
  const cls = String(event[APP.CALENDAR_HEADERS.CLASS] || '').trim();
  if (type === 'FOIL ACADEMY') return { folderId:EVENT_DRIVE.FOIL_ACADEMY, classSubfolder:'' };
  if (type === 'TEST FISICI') return { folderId:EVENT_DRIVE.TESTS, classSubfolder:'' };
  if (type === 'RIUNIONE/MISSIONE' || type.indexOf('RIUNION') >= 0 || type.indexOf('MISSION') >= 0) return { folderId:EVENT_DRIVE.MISSIONS, classSubfolder:'' };
  if (type.indexOf('ALLENAMENTO') >= 0 || type === 'STAGE' || type.indexOf('REGATA') >= 0 || type.indexOf('OSSERVAZIONE') >= 0) return { folderId:EVENT_DRIVE.TRAINING_RACES, classSubfolder:cls || 'SENZA CLASSE' };
  return { folderId:EVENT_DRIVE.ROOT, classSubfolder:'' };
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
  return String(value || '').toLowerCase().replace(/(^|\s|[-/])([a-zà-öø-ÿ])/g, function(_, p, c) { return p + c.toUpperCase(); });
}

function escapeRegexV2_(value) {
  return String(value || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function safeDocReplacementV2_(value) {
  return String(value || '').replace(/\\/g, '\\\\').replace(/\$/g, '\\$');
}
