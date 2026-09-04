function installEventAppForSelectedEvent() {
  const ui = SpreadsheetApp.getUi();
  const event = selectedEvent_();
  const eventId = ensureEventId_(event);
  const folder = createWorkFolderForEvent_(eventId,event,event._row);
  const result = getOrCreateEventSheetV3_(eventId,event,folder.folderId);
  const child = result.spreadsheet;
  writeEventMetaV5_(child,eventId,event,folder.folderId);

  try {
    const installed = provisionEventBoundScriptV1_(child);
    ui.alert(
      'Funzioni scheda evento installate',
      'Apri o ricarica la scheda evento. Troverai il menu “Scheda evento” con Importa nuova spesa e Genera documenti.',
      ui.ButtonSet.OK
    );
    return installed;
  } catch (err) {
    const message = String(err && err.message || err);
    ui.alert(
      'Installazione non completata',
      message + '\n\nSe il messaggio indica che la Google Apps Script API non è abilitata, abilitala una sola volta e ripeti questo comando.',
      ui.ButtonSet.OK
    );
    throw err;
  }
}

function provisionEventBoundScriptV1_(child) {
  const meta = child.getSheetByName(EVENT_SHEET.SHEETS.META);
  if (!meta) throw new Error('Foglio _META non trovato nella scheda evento.');
  let scriptId = String(readMetaValue_(meta,'EVENT_SCRIPT_ID')||'').trim();
  const token = ScriptApp.getOAuthToken();
  const headers = {Authorization:'Bearer ' + token,'Content-Type':'application/json'};

  if (!scriptId) {
    const create = UrlFetchApp.fetch('https://script.googleapis.com/v1/projects',{
      method:'post',
      headers:headers,
      muteHttpExceptions:true,
      payload:JSON.stringify({
        title:'Automazioni - ' + child.getName(),
        parentId:child.getId()
      })
    });
    if (create.getResponseCode() < 200 || create.getResponseCode() >= 300) {
      throw new Error(appsScriptApiErrorV1_('Creazione progetto Apps Script',create));
    }
    const created = JSON.parse(create.getContentText()||'{}');
    scriptId = String(created.scriptId||'').trim();
    if (!scriptId) throw new Error('Google Apps Script API non ha restituito lo scriptId.');
    writeMeta_(meta,{EVENT_SCRIPT_ID:scriptId});
  }

  const source = HtmlService.createHtmlOutputFromFile('EventBoundCode').getContent();
  const manifest = {
    timeZone:APP.TZ,
    exceptionLogging:'STACKDRIVER',
    runtimeVersion:'V8'
  };
  const update = UrlFetchApp.fetch('https://script.googleapis.com/v1/projects/' + encodeURIComponent(scriptId) + '/content',{
    method:'put',
    headers:headers,
    muteHttpExceptions:true,
    payload:JSON.stringify({files:[
      {name:'appsscript',type:'JSON',source:JSON.stringify(manifest,null,2)},
      {name:'Code',type:'SERVER_JS',source:source}
    ]})
  });
  if (update.getResponseCode() < 200 || update.getResponseCode() >= 300) {
    throw new Error(appsScriptApiErrorV1_('Aggiornamento progetto Apps Script',update));
  }
  writeMeta_(meta,{
    EVENT_SCRIPT_ID:scriptId,
    EVENT_SCRIPT_VERSION:'1',
    EVENT_SCRIPT_UPDATED:Utilities.formatDate(new Date(),APP.TZ,'dd/MM/yyyy HH:mm')
  });
  return {scriptId:scriptId,spreadsheetId:child.getId()};
}

function appsScriptApiErrorV1_(label,response) {
  const code = response.getResponseCode();
  let detail = response.getContentText() || '';
  try {
    const parsed = JSON.parse(detail);
    if (parsed && parsed.error && parsed.error.message) detail = parsed.error.message;
  } catch (_) {}
  if (code === 403 || code === 404) {
    return label + ' non riuscita (' + code + '). Verifica che la Google Apps Script API sia abilitata per il progetto Google Cloud associato allo script. Dettaglio: ' + detail;
  }
  return label + ' non riuscita (' + code + '): ' + detail;
}
