const ACTIVITY_V13=Object.freeze({VISIBLE:7,WIDTH:19,TYPE:8,OBJECTIVE_ID:9,TASK_ID:10,OBJECTIVE_ORDER:11,STEP_ORDER:12,COLOR:13,AUTO_DUE:14,PREVIOUS_ID:15,COMPLETED_AT:16,PATH:17,DUE_MODE:18,OFFSET:19});
const ACTIVITIES_EMERGENCY_MANUAL=true;

/**
 * MODALITA MANUALE DI EMERGENZA - 14/09/2026
 *
 * Per le schede gia esistenti il foglio Attivita NON viene mai rigenerato,
 * sincronizzato o riordinato da questo file. Il Calendario puo aggiornare i
 * metadati generali dell'evento, ma Attivita resta completamente manuale.
 */
function prepareEventSheetForSelectedEventV13(){
  const event=selectedEvent_();
  const eventId=ensureEventId_(event);
  const folder=createWorkFolderForEvent_(eventId,event,event._row);
  let child=getLinkedEventSheet_(event),created=false;

  if(!child){
    const copy=DriveApp.getFileById(EVENT_SHEET.TEMPLATE_ID).makeCopy(buildEventSheetName_(event),DriveApp.getFolderById(folder.folderId));
    child=SpreadsheetApp.openById(copy.getId());
    child.setSpreadsheetLocale('it_IT');
    child.setSpreadsheetTimeZone(APP.TZ);
    created=true;
  }

  if(!isCurrentEventSheet_(child))throw new Error('La Scheda evento collegata non usa il modello corrente.');
  validateEventSheetIdentity_(child,eventId);
  writeEventMeta_(child,eventId,event,folder.folderId);
  applyEventCommitment_(child,event);

  // Le altre sezioni restano come prima. Attivita non viene mai riscritta.
  if(created){
    writeCurrentParticipantsToEventSheet_(eventId,child);
    populateCurrentTechnicians_(child,event);
  }

  setEventSheetLink_(event._row,child.getUrl());
  SpreadsheetApp.flush();
  SpreadsheetApp.getActive().toast(
    created?'Scheda evento creata - Attivita manuale':'Dati generali aggiornati - Attivita non modificata',
    'Scheda evento',5
  );
  return{created:created,id:child.getId(),url:child.getUrl()};
}

function isObjectiveActivitySheetV13_(child){
  const sh=child&&child.getSheetByName(EVENT_SHEET.SHEETS.TASKS);
  return !!sh;
}

// In modalita emergenza nessuna funzione deve ricostruire o formattare Attivita.
function clearActivityRowGroupsV13_(sheet){return 0;}
function writeObjectivesToEventSheetV13_(eventId,child){return{manual:true};}
function syncObjectiveEventSheetV13_(child,eventId){return{manual:true,objectives:0,tasks:0};}

function syncSelectedEventSheetToCalendarV13(){
  SpreadsheetApp.getActive().toast('Attivita in modalita manuale: nessuna sincronizzazione eseguita.','Attivita',5);
  return{manual:true};
}

function syncAllEventSheetsToCalendarV13(){
  SpreadsheetApp.getUi().alert(
    'Attivita in modalita manuale',
    'La sincronizzazione automatica/manuale delle Attivita e temporaneamente disattivata. Nessuna scheda e stata modificata.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return{manual:true};
}
