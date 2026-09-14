const EVENT_SHEET_V22 = Object.freeze({
  TEMPLATE_ID: '195ZhB6VL-oygW8tW9iaKH73YkojtP9IDn30XZnnTTQs',
  SCHEMA_VERSION: '2.2',
  SYNC_VERSION: '14',
  OPERATIONAL_SOURCE: 'EVENT_SHEET'
});

function eventSheetTemplateIdV22_(){
  return EVENT_SHEET_V22.TEMPLATE_ID;
}

function eventSheetMarkContractV22_(child){
  if(!child)return;
  const meta=child.getSheetByName(EVENT_SHEET.SHEETS.META);
  if(!meta)return;
  writeMeta_(meta,{
    SCHEMA_VERSION:EVENT_SHEET_V22.SCHEMA_VERSION,
    SYNC_VERSION:EVENT_SHEET_V22.SYNC_VERSION,
    OPERATIONAL_SOURCE:EVENT_SHEET_V22.OPERATIONAL_SOURCE,
    LAST_SYNC:Utilities.formatDate(new Date(),APP.TZ,'dd/MM/yyyy HH:mm:ss')
  });
  meta.hideSheet();
}

function eventSheetIsV22_(child){
  if(!child||!isCurrentEventSheet_(child)||!isObjectiveActivitySheetV13_(child))return false;
  const meta=child.getSheetByName(EVENT_SHEET.SHEETS.META);
  if(!meta)return false;
  const source=normalize_(readMetaValue_(meta,'OPERATIONAL_SOURCE'));
  const schema=String(readMetaValue_(meta,'SCHEMA_VERSION')||'').trim();
  return source==='EVENT_SHEET'&&schema==='2.2';
}

/**
 * Regola architetturale v2.2:
 * - Calendario = master dei soli dati generali dell'evento.
 * - Scheda Evento = master di attivita, partecipanti, spese e processi.
 * - Fogli _CHECKLIST/_PARTECIPANTI/_SPESE = proiezione tecnica/cache usata
 *   dal cruscotto e dalle automazioni; non e' la fonte operativa primaria.
 */
function eventSheetSyncOperationalToBackendV22_(eventId,event,child){
  validateEventSheetIdentity_(child,eventId);
  if(!isCurrentEventSheet_(child))throw new Error('Scheda evento non compatibile con il contratto operativo corrente.');

  const participants=calendarSyncParticipantsV12_(eventId,child);
  let activities=0;
  if(isObjectiveActivitySheetV13_(child)){
    const result=syncObjectiveEventSheetV13_(child,eventId);
    activities=Number(result&&result.tasks||0);
  }else{
    activities=calendarSyncTasksV12_(eventId,child);
  }
  const expenses=calendarSyncExpensesV12_(eventId,event,child);
  eventSheetMarkContractV22_(child);
  SpreadsheetApp.flush();
  return{participants:participants,activities:activities,expenses:expenses};
}

function eventSheetPushGenericMetadataFromCalendarV22_(child,eventId,event,folderId){
  validateEventSheetIdentity_(child,eventId);
  writeEventMeta_(child,eventId,event,folderId);
  applyEventCommitment_(child,event);
  eventSheetMarkContractV22_(child);
}

function eventSheetSyncModeV22_(child){
  if(!child)return'MISSING';
  if(eventSheetIsV22_(child))return'V22';
  if(isObjectiveActivitySheetV13_(child)&&isCurrentEventSheet_(child))return'OBJECTIVES_LEGACY_META';
  if(typeof calendarIsCurrentEventSheetV12_==='function'&&calendarIsCurrentEventSheetV12_(child))return'LEGACY_V12';
  return'UNSUPPORTED';
}
