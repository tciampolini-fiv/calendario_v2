const EVENT_SHEET_TEMPLATE_ID_V3 = '1obNkubjpGhC7o8AX5Tc986UBGdyJs26b476b-GZN3Sc';

function prepareEventSheetForSelectedEventV2() {
  const event = selectedEvent_();
  const eventId = ensureEventId_(event);
  const folder = createWorkFolderForEvent_(eventId, event, event._row);
  const result = getOrCreateEventSheetV3_(eventId, event, folder.folderId);
  const child = result.spreadsheet;

  ensureChecklistBackendHeadersV3_();
  ensureParticipantsBackendHeadersV2_();

  // Per una nuova scheda generiamo subito la checklist standard centrale e poi
  // riversiamo SOLO i dati in Attività. Il dropdown nativo del modello non viene toccato.
  if (result.created) {
    generateChecklistForEvent_(eventId,event);
  }
  refreshTasksPreserveTemplateV6_(eventId,event,child);

  if (isEventSheetV7_(child)) {
    initializeEventSheetV7FromCalendar_(child,eventId,event,folder.folderId);
  } else if (!result.created) {
    // Compatibilità con le vecchie schede, senza alterare il nuovo modello V7.
    refreshParticipantsV3FromBackend_(eventId,event,child);
    refreshExpensesV4FromBackend_(eventId,event,child,folder.folderId);
    hideEventSheetTechnicalColumnsV3_(child);
  }

  ensureFastEventTaskTriggerV6_(child);
  setEventSheetLink_(event._row, child.getUrl());
  writeEventMetaV5_(child,eventId,event,folder.folderId);
  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    result.created ? 'Scheda evento creata dal modello' : 'Scheda evento aggiornata',
    result.created
      ? 'È stata creata una nuova Scheda evento con task, IMP e tecnici già inizializzati.'
      : 'La Scheda evento collegata è stata aggiornata senza modificare la struttura del modello.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return {created:result.created,id:child.getId(),url:child.getUrl()};
}

function isEventSheetV7_(child) {
  const expenses = child && child.getSheetByName(EVENT_SHEET.SHEETS.EXPENSES);
  const participants = child && child.getSheetByName(PARTICIPANTS_V3.SHEET);
  if (!expenses || !participants) return false;
  return normalize_(expenses.getRange('A13').getDisplayValue()) === 'PREVENTIVO' &&
    normalize_(expenses.getRange('E13').getDisplayValue()) === 'PAGAMENTO 1' &&
    normalize_(participants.getRange('H2').getDisplayValue()) === 'RUOLO';
}

function refreshTasksPreserveTemplateV6_(eventId,event,child) {
  ensureChecklistBackendHeadersV3_();
  seedPresenceCheckTaskV3_(eventId,event,child);
  ensureTaskNumbersAndDefaultDependenciesV3_(eventId,event);
  syncChecklistLocksForEvent_(eventId);

  const backendSheet = sh_(APP.SHEETS.CHECKLIST);
  const backend = backendSheet.getDataRange().getValues().slice(1)
    .filter(r => String(r[1]) === String(eventId));
  const byId = {};
  backend.forEach(r => { if (r[0]) byId[String(r[0])] = r; });
  backend.sort((a,b) => Number(a[2]||999999)-Number(b[2]||999999) || Number(a[13]||999999)-Number(b[13]||999999));

  const out = backend.map(r => {
    const dep = r[14] ? byId[String(r[14])] : null;
    return [
      r[3]||'',
      r[10]||'',
      r[13]||'',
      dep ? dep[13]||'' : '',
      normalize_(r[6]) === 'COMPLETATA' ? 'FATTO' : (r[6]||'DA FARE'),
      r[5]||''
    ];
  });

  const sheet = child.getSheetByName(EVENT_SHEET.SHEETS.TASKS);
  if (!sheet) throw new Error('Foglio Attività non trovato nella Scheda evento.');

  const endRow = Math.min(TASKS_V4.ENTRY_END_ROW, sheet.getMaxRows());
  if (endRow >= 2) sheet.getRange(2,1,endRow-1,TASKS_V4.VISIBLE_COLS).clearContent();
  if (out.length) sheet.getRange(2,1,out.length,TASKS_V4.VISIBLE_COLS).setValues(out);

  backend.forEach((r,index) => {
    const row = index + 2;
    const depId = String(r[14]||'').trim();
    const status = normalize_(r[6]);
    if (!depId || status === 'FATTO' || status === 'COMPLETATA') return;
    sheet.getRange(row,5).setFormula(fastTaskStatusFormulaV7_(row));
  });
  return out.length;
}

function ensureFastEventTaskTriggerV6_(child) {
  const sourceId = child.getId();
  const triggers = ScriptApp.getProjectTriggers();
  let fastExists = false;

  triggers.forEach(t => {
    let sameSource = false;
    try { sameSource = t.getTriggerSourceId() === sourceId; } catch (e) {}
    if (!sameSource) return;

    const handler = t.getHandlerFunction();
    if (handler === 'handleEventTaskEditV4') {
      ScriptApp.deleteTrigger(t);
      return;
    }
    if (handler === 'handleFastEventTaskEditV6') fastExists = true;
  });

  if (!fastExists) {
    ScriptApp.newTrigger('handleFastEventTaskEditV6')
      .forSpreadsheet(sourceId)
      .onEdit()
      .create();
  }
}

function fastTaskStatusFormulaV7_(row) {
  return '=IF(A' + row + '="";"";IF(D' + row + '="";"DA FARE";IFERROR(IF(INDEX($E$2:$E$500;MATCH(D' + row + ';$C$2:$C$500;0))="FATTO";"DA FARE";"IN ATTESA");"IN ATTESA")))';
}

function handleFastEventTaskEditV6(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== EVENT_SHEET.SHEETS.TASKS) return;

  const firstRow = Math.max(2,e.range.getRow());
  const lastRow = Math.min(TASKS_V4.ENTRY_END_ROW,e.range.getLastRow());
  const firstCol = e.range.getColumn();
  const lastCol = e.range.getLastColumn();
  if (lastRow < firstRow || firstCol > TASKS_V4.VISIBLE_COLS) return;

  if (firstCol <= 1 && lastCol >= 1) {
    const allNumbers = sheet.getRange(2,3,Math.max(1,Math.min(TASKS_V4.ENTRY_END_ROW,sheet.getMaxRows())-1),1).getValues();
    let nextNumber = allNumbers.reduce((max,row) => {
      const n = Number(row[0]);
      return Number.isFinite(n) && n > max ? n : max;
    },0) + 1;

    for (let row=firstRow;row<=lastRow;row++) {
      const description = String(sheet.getRange(row,1).getDisplayValue()||'').trim();
      if (!description) continue;
      if (!(Number(sheet.getRange(row,3).getValue()) > 0)) sheet.getRange(row,3).setValue(nextNumber++);
      autoLinkConfirmationTaskFastV7_(sheet,row);
      applyFastTaskDependencyV7_(sheet,row);
    }
  }

  if (firstCol <= 4 && lastCol >= 4) {
    for (let row=firstRow;row<=lastRow;row++) applyFastTaskDependencyV7_(sheet,row);
  }
  if (firstCol <= 5 && lastCol >= 5) refreshFastTaskDependenciesV7_(sheet);
}

function normalizeTaskTextFastV7_(value) {
  return String(value===null||value===undefined?'':value).trim().toUpperCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ');
}

function autoLinkConfirmationTaskFastV7_(sheet,row) {
  if (Number(sheet.getRange(row,4).getValue()||0)>0) return;
  const desc = normalizeTaskTextFastV7_(sheet.getRange(row,1).getDisplayValue());
  let parent = '';
  if (desc === 'CHECK CONFERMA PRESENZE') parent = 'CONVOCAZIONE ATLETI';
  else if (desc === 'CONFERMA OSPITALITA') parent = 'OSPITALITA CIRCOLO';
  else if (desc.indexOf('CONFERMA ') === 0) parent = desc.substring(9).trim();
  if (!parent) return;
  const count = Math.max(1,row-2);
  const values = sheet.getRange(2,1,count,3).getDisplayValues();
  for (let i=values.length-1;i>=0;i--) {
    if (normalizeTaskTextFastV7_(values[i][0])===parent && Number(values[i][2]||0)>0) {
      sheet.getRange(row,4).setValue(Number(values[i][2]));
      return;
    }
  }
}

function applyFastTaskDependencyV7_(sheet,row) {
  const description = String(sheet.getRange(row,1).getDisplayValue()||'').trim();
  if (!description) return;
  const status = normalize_(sheet.getRange(row,5).getDisplayValue());
  if (status === 'FATTO') return;
  const depNo = Number(sheet.getRange(row,4).getValue()||0);
  if (depNo>0) sheet.getRange(row,5).setFormula(fastTaskStatusFormulaV7_(row));
  else if (!status || status === 'IN ATTESA') sheet.getRange(row,5).setValue('DA FARE');
}

function refreshFastTaskDependenciesV7_(sheet) {
  const last = Math.min(500,Math.max(sheet.getLastRow(),2));
  for (let row=2;row<=last;row++) {
    const description = String(sheet.getRange(row,1).getDisplayValue()||'').trim();
    if (!description) continue;
    const depNo = Number(sheet.getRange(row,4).getValue()||0);
    const status = normalize_(sheet.getRange(row,5).getDisplayValue());
    if (depNo>0 && status !== 'FATTO') sheet.getRange(row,5).setFormula(fastTaskStatusFormulaV7_(row));
  }
}

function getLinkedEventSheetForSaveV2_(event) {
  if (!event || !event._row) throw new Error('Seleziona una riga evento valida nel Calendario.');
  const cal = sh_(APP.SHEETS.CALENDAR);
  const map = headerMap_(cal);
  const col = map[APP.CALENDAR_HEADERS.EVENT_SHEET];
  if (!col) throw new Error('Colonna SCHEDA EVENTO non trovata nel Calendario.');

  const cell = cal.getRange(event._row,col);
  const rich = cell.getRichTextValue();
  const url = (rich && rich.getLinkUrl()) || cell.getDisplayValue();
  const id = extractDriveId_(url);
  if (!id) throw new Error('Questa riga non ha una Scheda evento collegata. Usa prima “Crea / aggiorna scheda ← Calendario”.');
  try { return SpreadsheetApp.openById(id); }
  catch (e) { throw new Error('Non riesco ad aprire la Scheda evento collegata. Verifica il link nella colonna SCHEDA EVENTO.'); }
}

function getExplicitLinkedEventSheetV3_(event) {
  if (!event || !event._row) return null;
  const cal = sh_(APP.SHEETS.CALENDAR);
  const map = headerMap_(cal);
  const col = map[APP.CALENDAR_HEADERS.EVENT_SHEET];
  if (!col) return null;
  const cell = cal.getRange(event._row,col);
  const rich = cell.getRichTextValue();
  const url = (rich && rich.getLinkUrl()) || cell.getDisplayValue();
  const id = extractDriveId_(url);
  if (!id) return null;
  try { return SpreadsheetApp.openById(id); } catch (e) { return null; }
}

function syncSelectedEventSheetToCalendarV2() {
  const ui = SpreadsheetApp.getUi();
  const event = selectedEvent_();
  const eventId = ensureEventId_(event);
  const child = getLinkedEventSheetForSaveV2_(event);
  validateEventSheetIdentity_(child,eventId);

  if (isEventSheetV7_(child)) {
    ui.alert(
      'Salvataggio dalla Scheda evento',
      'Questa è una Scheda evento di nuova generazione. Per salvare usa il menu “Scheda evento” → “💾 Salva dati nel Calendario” direttamente nel file evento.\n\nIn questo modo vengono gestiti correttamente pagamenti parziali, residui e task automatiche.',
      ui.ButtonSet.OK
    );
    return null;
  }

  ensureChecklistBackendHeadersV3_();
  ensureParticipantsBackendHeadersV2_();
  const participantCount = syncParticipantsV3ToBackend_(eventId,event,child);
  const taskCount = syncTasksV4ToBackend_(eventId,event,child);
  const expenseCount = syncExpensesV4ToBackend_(eventId,event,child);
  ensureFastEventTaskTriggerV6_(child);
  SpreadsheetApp.flush();
  const counts = {tasks:taskCount,participants:participantCount,expenses:expenseCount};
  ui.alert('Calendario aggiornato','I dati presenti nella Scheda evento sono stati salvati nel Calendario.\n\nAttività: '+counts.tasks+'\nSpese: '+counts.expenses+'\nPartecipanti: '+counts.participants,ui.ButtonSet.OK);
  return counts;
}

function canonicalEventTypeV7_(value) {
  const t = normalizeTaskTextFastV7_(value);
  const aliases = {
    'RIUNIONE/MISSIONE':'RIUNIONI/MISSIONI',
    'RIUNIONI / MISSIONI':'RIUNIONI/MISSIONI',
    'OSSERVAZIONE REGATA':'OSS.REGATE ITA',
    'ISCRIZIONE REGATA':'ISCRIZIONI REGATE'
  };
  return aliases[t] || t;
}

function canonicalEventClassV7_(value) {
  const c = normalizeTaskTextFastV7_(value);
  return c === 'KITE' ? 'KITEFOIL' : c;
}

function resolveEventCommitmentV7_(event) {
  const type = canonicalEventTypeV7_(event[APP.CALENDAR_HEADERS.TYPE]);
  const cls = canonicalEventClassV7_(event[APP.CALENDAR_HEADERS.CLASS]);
  const config = sh_('_CONFIG_IMPEGNI').getDataRange().getDisplayValues();
  let wildcard = '';
  for (let i=1;i<config.length;i++) {
    if (canonicalEventTypeV7_(config[i][0]) !== type) continue;
    const cfgClass = canonicalEventClassV7_(config[i][1]);
    const imp = String(config[i][2]||'').trim();
    if (!imp) continue;
    if (cfgClass === cls || (cfgClass === 'TUTTE' && cls)) return imp;
    if (!cfgClass || cfgClass === '*') wildcard = imp;
  }
  return wildcard;
}

function technicianDirectoryV7_() {
  return [
    ['ZAGGIA','Leonardo','Zaggia'],['CRISI','Andrea','Crisi'],['RAVEGLIA','Matteo','Raveglia'],['CARICATO','Francesco','Caricato'],
    ['PICCIAU','Gianluigi','Picciau'],['SENSINI','Alessandra','Sensini'],['NUICOLUCCI','Matteo','Nuicolucci'],['CAMBONI','Mattia','Camboni'],
    ['CANGEMI','Antonino','Cangemi'],['LOPERFIDO','Daniel','Loperfido']
  ];
}

function resolveFullTechniciansV7_(raw) {
  const text = normalizeTaskTextFastV7_(raw);
  if (!text) return [];
  return technicianDirectoryV7_().filter(x=>text.indexOf(x[0])>=0);
}

function initializeEventSheetV7FromCalendar_(child,eventId,event,folderId) {
  writeEventMetaV5_(child,eventId,event,folderId);
  const imp = resolveEventCommitmentV7_(event) || String(event[APP.CALENDAR_HEADERS.COMMITMENT]||'').trim();
  if (imp) {
    const cal = sh_(APP.SHEETS.CALENDAR);
    const map = headerMap_(cal);
    const col = map[APP.CALENDAR_HEADERS.COMMITMENT];
    if (col && event._row) cal.getRange(event._row,col).setValue(imp).setNumberFormat('@');
    const expenses = child.getSheetByName(EVENT_SHEET.SHEETS.EXPENSES);
    if (expenses) expenses.getRange('N9').setValue(imp).setNumberFormat('@');
  }
  populateTechniciansInEventSheetV7_(child,event);
}

function populateTechniciansInEventSheetV7_(child,event) {
  const people = resolveFullTechniciansV7_(event[APP.CALENDAR_HEADERS.TECHNICIANS]);
  if (!people.length) return;
  const sheet = child.getSheetByName(PARTICIPANTS_V3.SHEET);
  if (!sheet || normalize_(sheet.getRange('H2').getDisplayValue()) !== 'RUOLO') return;
  const existing = sheet.getRange(3,1,15,2).getDisplayValues();
  const existingNames = new Set(existing.map(r=>normalizeTaskTextFastV7_((r[0]||'')+' '+(r[1]||''))));
  people.forEach(x=>{
    const fullKey = normalizeTaskTextFastV7_(x[1]+' '+x[2]);
    if (existingNames.has(fullKey)) return;
    let target = 0;
    for (let i=0;i<existing.length;i++) {
      if (!String(existing[i][0]||'').trim() && !String(existing[i][1]||'').trim()) { target=i+3; existing[i]=[x[1],x[2]]; break; }
    }
    if (!target) return;
    sheet.getRange(target,1).setValue(x[1]);
    sheet.getRange(target,2).setValue(x[2]);
    sheet.getRange(target,8).setValue('TECNICO');
    sheet.getRange(target,9).setValue('CONFERMATO');
    existingNames.add(fullKey);
  });
}

function writeEventMetaV5_(child,eventId,event,folderId) {
  const meta = child.getSheetByName(EVENT_SHEET.SHEETS.META);
  const start = event[APP.CALENDAR_HEADERS.START];
  const end = event[APP.CALENDAR_HEADERS.END];
  const imp = resolveEventCommitmentV7_(event) || String(event[APP.CALENDAR_HEADERS.COMMITMENT]||'').trim();
  const technicians = resolveFullTechniciansV7_(event[APP.CALENDAR_HEADERS.TECHNICIANS]);
  const fullTechnicians = technicians.length
    ? technicians.map(x=>x[1]+' '+x[2]).join(', ')
    : String(event[APP.CALENDAR_HEADERS.TECHNICIANS]||'');
  writeMeta_(meta,{
    EVENT_ID:eventId,
    MASTER_SPREADSHEET_ID:APP.SPREADSHEET_ID,
    EVENT_FOLDER_ID:folderId,
    EVENT_SHEET_ID:child.getId(),
    SYNC_VERSION:'7',
    EVENT_LABEL:buildEventSheetLabel_(event),
    EVENT_TYPE:String(event[APP.CALENDAR_HEADERS.TYPE]||''),
    EVENT_CLASS:String(event[APP.CALENDAR_HEADERS.CLASS]||''),
    EVENT_LOCATION:String(event[APP.CALENDAR_HEADERS.LOCATION]||''),
    EVENT_ZONE:String(event[APP.CALENDAR_HEADERS.ZONE]||''),
    EVENT_CLUB:String(event[APP.CALENDAR_HEADERS.CLUB]||''),
    EVENT_TECHNICIANS:fullTechnicians,
    EVENT_LODGING:String(event[APP.CALENDAR_HEADERS.LODGING]||''),
    EVENT_COMMITMENT:imp,
    EVENT_START:start instanceof Date ? Utilities.formatDate(start,APP.TZ,'yyyy-MM-dd') : '',
    EVENT_END:end instanceof Date ? Utilities.formatDate(end,APP.TZ,'yyyy-MM-dd') : '',
    LAST_SYNC:Utilities.formatDate(new Date(),APP.TZ,'dd/MM/yyyy HH:mm')
  });
}

function getOrCreateEventSheetV3_(eventId,event,folderId) {
  let child = getExplicitLinkedEventSheetV3_(event);
  let created = false;
  if (!child) {
    const folder = DriveApp.getFolderById(folderId);
    const templateFile = DriveApp.getFileById(EVENT_SHEET_TEMPLATE_ID_V3);
    const copy = templateFile.makeCopy(buildEventSheetName_(event), folder);
    child = SpreadsheetApp.openById(copy.getId());
    child.setSpreadsheetLocale('it_IT');
    child.setSpreadsheetTimeZone(APP.TZ);
    created = true;
  }
  ensureEventSheetBaseV3_(child,eventId,folderId,event);
  return {spreadsheet:child,created:created};
}

function ensureEventSheetBaseV3_(child,eventId,folderId,event) {
  child.setSpreadsheetLocale('it_IT');
  child.setSpreadsheetTimeZone(APP.TZ);
  let tasks = child.getSheetByName(EVENT_SHEET.SHEETS.TASKS);
  if (!tasks) {
    const first = child.getSheets()[0];
    if (child.getSheets().length === 1 && first.getLastRow() === 0) {
      first.setName(EVENT_SHEET.SHEETS.TASKS);
      tasks = first;
    } else tasks = child.insertSheet(EVENT_SHEET.SHEETS.TASKS,0);
  }
  if (!child.getSheetByName(EVENT_SHEET.SHEETS.EXPENSES)) child.insertSheet(EVENT_SHEET.SHEETS.EXPENSES);
  if (!child.getSheetByName(PARTICIPANTS_V3.SHEET)) child.insertSheet(PARTICIPANTS_V3.SHEET);
  let meta = child.getSheetByName(EVENT_SHEET.SHEETS.META);
  if (!meta) meta = child.insertSheet(EVENT_SHEET.SHEETS.META);
  writeEventMetaV5_(child,eventId,event,folderId);
  meta.hideSheet();
}

function hideEventSheetTechnicalColumnsV3_(child) {
  const tasks = child.getSheetByName(EVENT_SHEET.SHEETS.TASKS);
  if (tasks) {
    tasks.showColumns(1,TASKS_V4.VISIBLE_COLS);
    if (tasks.getMaxColumns() > TASKS_V4.VISIBLE_COLS) tasks.hideColumns(TASKS_V4.VISIBLE_COLS + 1, tasks.getMaxColumns() - TASKS_V4.VISIBLE_COLS);
  }
  const expenses = child.getSheetByName(EVENT_SHEET.SHEETS.EXPENSES);
  if (expenses) {
    expenses.showColumns(1,EXPENSES_V4.VISIBLE_COLS);
    if (expenses.getMaxColumns() > EXPENSES_V4.VISIBLE_COLS) expenses.hideColumns(EXPENSES_V4.VISIBLE_COLS + 1, expenses.getMaxColumns() - EXPENSES_V4.VISIBLE_COLS);
  }
  const participants = child.getSheetByName(PARTICIPANTS_V3.SHEET);
  if (participants) {
    participants.showColumns(1,PARTICIPANTS_V3.VISIBLE_COLS);
    if (participants.getMaxColumns() > PARTICIPANTS_V3.VISIBLE_COLS) participants.hideColumns(PARTICIPANTS_V3.VISIBLE_COLS + 1, participants.getMaxColumns() - PARTICIPANTS_V3.VISIBLE_COLS);
  }
}
