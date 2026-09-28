const ACTIVITY_V13=Object.freeze({VISIBLE:7,WIDTH:19,TYPE:8,OBJECTIVE_ID:9,TASK_ID:10,OBJECTIVE_ORDER:11,STEP_ORDER:12,COLOR:13,AUTO_DUE:14,PREVIOUS_ID:15,COMPLETED_AT:16,PATH:17,DUE_MODE:18,OFFSET:19});
const EVENT_SHEET_SYNC_VERSION_V16='16';
const EVENT_SHEET_V16_HEADERS=Object.freeze(['OBIETTIVO','SCADENZA','ATTIVITÀ','NOTE','FATTO','STATO','','TIPO RIGA','ID OBIETTIVO','ID TASK','ORDINE OBIETTIVO','ORDINE STEP','COLORE','SCADENZA AUTOMATICA','ID TASK PRECEDENTE','DATA COMPLETAMENTO','PERCORSO','MODALITA SCADENZA','OFFSET GIORNI']);

function getEventSheetSchemaV15_(child){
  if(!child)return{kind:'MISSING',version:''};
  const sh=child.getSheetByName(EVENT_SHEET.SHEETS.TASKS),meta=child.getSheetByName(EVENT_SHEET.SHEETS.META);
  if(!sh||!meta)return{kind:'LEGACY',version:''};
  const headers=sh.getRange(1,1,1,ACTIVITY_V13.WIDTH).getDisplayValues()[0].map(normalize_);
  const exact=EVENT_SHEET_V16_HEADERS.every((h,i)=>headers[i]===h);
  const version=String(readMetaValue_(meta,'SYNC_VERSION')||'').trim();
  if(exact&&version===EVENT_SHEET_SYNC_VERSION_V16)return{kind:'V16',version:version};
  return{kind:'LEGACY',version:version};
}

function markLegacyCalendarRowV15_(row){
  const sheet=sh_(APP.SHEETS.CALENDAR),map=headerMap_(sheet);
  if(map[APP.CALENDAR_HEADERS.CHECKLIST])sheet.getRange(row,map[APP.CALENDAR_HEADERS.CHECKLIST]).setValue('SCHEDA VECCHIA');
  if(map[APP.CALENDAR_HEADERS.NEXT_ACTION])sheet.getRange(row,map[APP.CALENDAR_HEADERS.NEXT_ACTION]).setValue('Aprire la scheda per i dettagli');
  [APP.CALENDAR_HEADERS.BUDGET,APP.CALENDAR_HEADERS.ACTUAL,APP.CALENDAR_HEADERS.TO_PAY].forEach(h=>{if(map[h])sheet.getRange(row,map[h]).setValue('Aprire scheda');});
}

function prepareEventSheetForSelectedEventV15(){
  const event=selectedEvent_(),eventId=ensureEventId_(event),folder=createWorkFolderForEvent_(eventId,event,event._row);
  let child=getLinkedEventSheet_(event),created=false;
  if(child){
    const schema=getEventSheetSchemaV15_(child);
    if(schema.kind!=='V16'){
      markLegacyCalendarRowV15_(event._row);
      SpreadsheetApp.getUi().alert('Scheda evento vecchia','La scheda collegata resta invariata. Aprila direttamente per consultare attività, partecipanti e spese.',SpreadsheetApp.getUi().ButtonSet.OK);
      return{created:false,legacy:true,id:child.getId(),url:child.getUrl()};
    }
    syncObjectiveEventSheetV13_(child,eventId);
  }else{
    const copy=DriveApp.getFileById(EVENT_SHEET.TEMPLATE_ID).makeCopy(buildEventSheetName_(event),DriveApp.getFolderById(folder.folderId));
    child=SpreadsheetApp.openById(copy.getId());
    child.setSpreadsheetLocale('it_IT');
    child.setSpreadsheetTimeZone(APP.TZ);
    created=true;
    writeMeta_(child.getSheetByName(EVENT_SHEET.SHEETS.META),{SYNC_VERSION:EVENT_SHEET_SYNC_VERSION_V16});
  }

  if(!isCurrentEventSheet_(child)||!isObjectiveActivitySheetV13_(child))throw new Error('Il modello della Scheda evento nuova non e compatibile con V16.');
  validateEventSheetIdentity_(child,eventId);
  writeEventMeta_(child,eventId,event,folder.folderId);
  applyEventCommitment_(child,event);
  ensureDefaultObjectivesForEventV14_(eventId,event);
  writeObjectivesToEventSheetV13_(eventId,child);
  if(created)writeCurrentParticipantsToEventSheet_(eventId,child);
  populateCurrentTechnicians_(child,event);
  setEventSheetLink_(event._row,child.getUrl());
  refreshCalendarActivityDashboardV14_(eventId);
  SpreadsheetApp.flush();
  SpreadsheetApp.getActive().toast(created?'Scheda evento V16 creata':'Scheda evento V16 aggiornata','Scheda evento',5);
  return{created:created,legacy:false,id:child.getId(),url:child.getUrl()};
}

function isObjectiveActivitySheetV13_(child){
  const sh=child&&child.getSheetByName(EVENT_SHEET.SHEETS.TASKS);
  if(!sh)return false;
  const headers=sh.getRange(1,1,1,ACTIVITY_V13.WIDTH).getDisplayValues()[0].map(normalize_);
  return EVENT_SHEET_V16_HEADERS.every((h,i)=>headers[i]===h);
}

function syncSelectedEventSheetToCalendarV15(){
  const event=selectedEvent_(),eventId=ensureEventId_(event),child=getLinkedEventSheet_(event);
  if(!child)throw new Error('Nessuna Scheda evento collegata.');
  const schema=getEventSheetSchemaV15_(child);
  if(schema.kind!=='V16'){
    markLegacyCalendarRowV15_(event._row);
    SpreadsheetApp.getActive().toast('Scheda vecchia: nessuna sincronizzazione','Scheda evento',5);
    return{legacy:true};
  }
  syncObjectiveEventSheetV13_(child,eventId);
  SpreadsheetApp.getActive().toast('Scheda nuova sincronizzata','Attività',4);
  return{legacy:false};
}

function syncAllEventSheetsToCalendarV15(){
  const cal=sh_(APP.SHEETS.CALENDAR),map=headerMap_(cal),last=cal.getLastRow();
  let synced=0,legacy=0,missing=0;
  for(let row=2;row<=last;row++){
    const eventId=String(cal.getRange(row,map[APP.CALENDAR_HEADERS.ID]).getDisplayValue()||'').trim();
    if(!eventId)continue;
    const event={_row:row};Object.keys(map).forEach(h=>event[h]=cal.getRange(row,map[h]).getValue());
    const child=getLinkedEventSheet_(event);
    if(!child){missing++;continue;}
    if(getEventSheetSchemaV15_(child).kind!=='V16'){markLegacyCalendarRowV15_(row);legacy++;continue;}
    syncObjectiveEventSheetV13_(child,eventId);synced++;
  }
  SpreadsheetApp.getUi().alert('Sincronizzazione Schede evento','Schede nuove sincronizzate: '+synced+'\nSchede vecchie ignorate: '+legacy+'\nSchede assenti: '+missing,SpreadsheetApp.getUi().ButtonSet.OK);
}
function clearActivityRowGroupsV13_(sheet){try{for(let r=2;r<=Math.min(500,sheet.getMaxRows());r++){const g=sheet.getRowGroup(r,1);if(g)g.remove();}}catch(e){console.log('Pulizia gruppi righe: '+e.message);}}
function writeObjectivesToEventSheetV13_(eventId,child){
  const sheet=child.getSheetByName(EVENT_SHEET.SHEETS.TASKS);
  if(!sheet)throw new Error('Foglio Attività non trovato.');
  const objectives=getActivityObjectivesForEventV14_(eventId),rows=[];

  objectives.forEach(o=>{
    rows.push([o.name,o.dueDate||'','＋',o.note||'','',o.status,'','OBIETTIVO',o.id,'',o.order||'','',o.color||'#D9EAF7',false,'','CUSTOM','','','']);
    o.tasks.forEach(t=>rows.push(['',t.dueDate||'',t.task||'',t.note||'',activityIsDoneV14_(t.status),t.status||'IN ATTESA','','TASK',o.id,t.id||'',o.order||'',t.stepOrder||'',o.color||'',t.autoDue===true,t.previousTaskId||t.dependencyId||'',t.completedAt||'',t.path||'MANUALE',t.dueMode||'',t.offsetDays===''?'':t.offsetDays]));
  });
  rows.push(['','','','','','','','','','','','','','','','','','','']);
  rows.push(['＋ AGGIUNGI OBIETTIVO','','','','','','','AGGIUNGI_OBIETTIVO','','','','','','','','','','','']);
  const actionRow=rows.length+1;

  clearActivityRowGroupsV13_(sheet);
  const max=Math.min(500,sheet.getMaxRows());
  if(max>=2){
    sheet.getRange(2,1,max-1,ACTIVITY_V13.WIDTH).clearContent();
    sheet.getRange(2,5,max-1,1).clearDataValidations();
  }
  if(rows.length)sheet.getRange(2,1,rows.length,ACTIVITY_V13.WIDTH).setValues(rows);

  objectives.forEach(o=>{
    const idx=rows.findIndex(r=>r[7]==='OBIETTIVO'&&r[8]===o.id);
    if(idx<0)return;
    const headerRow=idx+2;
    sheet.getRange(headerRow,1,1,6).setBackground(o.color||'#D9EAF7').setFontWeight('bold');
    sheet.getRange(headerRow,3).setHorizontalAlignment('center').setFontWeight('bold');
    const count=o.tasks.length;
    if(count){
      const first=headerRow+1;
      sheet.getRange(first,5,count,1).insertCheckboxes();
      try{sheet.getRange(first,1,count,1).shiftRowGroupDepth(1);}catch(e){console.log('Gruppo righe: '+e.message);}
    }
  });

  sheet.getRange(actionRow,1,1,6).setFontWeight('bold');
  sheet.getRange(actionRow,1).setHorizontalAlignment('left');
  sheet.getRange(2,2,Math.max(rows.length,1),1).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(2,4,Math.max(rows.length,1),1).setWrap(true);
  sheet.setColumnWidth(4,320);
  try{sheet.hideColumns(7,13);}catch(e){}
}

function syncSelectedEventSheetToCalendarV13(){const event=selectedEvent_(),eventId=ensureEventId_(event),child=getLinkedEventSheet_(event);if(!child)throw new Error('Nessuna Scheda evento collegata.');syncObjectiveEventSheetV13_(child,eventId);SpreadsheetApp.getActive().toast('Scheda sincronizzata','Attività',4);}
function syncAllEventSheetsToCalendarV13(){const cal=sh_(APP.SHEETS.CALENDAR),map=headerMap_(cal),last=cal.getLastRow();let synced=0,skipped=0;for(let row=2;row<=last;row++){const eventId=String(cal.getRange(row,map[APP.CALENDAR_HEADERS.ID]).getDisplayValue()||'').trim();if(!eventId)continue;const event={_row:row};Object.keys(map).forEach(h=>event[h]=cal.getRange(row,map[h]).getValue());const child=getLinkedEventSheet_(event);if(!child||!isObjectiveActivitySheetV13_(child)){skipped++;continue;}syncObjectiveEventSheetV13_(child,eventId);synced++;}SpreadsheetApp.getUi().alert('Sincronizzazione attività','Schede sincronizzate: '+synced+'\nSchede non compatibili/assenti: '+skipped,SpreadsheetApp.getUi().ButtonSet.OK);}
function syncObjectiveEventSheetV13_(child,eventId){
  validateEventSheetIdentity_(child,eventId);const local=child.getSheetByName(EVENT_SHEET.SHEETS.TASKS);if(!local||!isObjectiveActivitySheetV13_(child))throw new Error('La Scheda non usa il layout Obiettivi.');const b=ensureActivityBackendV14_(),oldObjs=b.objectives.getDataRange().getValues(),oldTasks=b.checklist.getDataRange().getValues(),objById={},taskById={};oldObjs.slice(1).forEach(r=>{if(String(r[1])===String(eventId)&&r[0])objById[String(r[0])]=r;});oldTasks.slice(1).forEach(r=>{if(String(r[1])===String(eventId)&&r[0])taskById[String(r[0])]=r;});
  const last=Math.min(500,Math.max(local.getLastRow(),2)),values=local.getRange(2,1,last-1,ACTIVITY_V13.WIDTH).getValues(),now=new Date(),objs=[],tasks=[],validObjectives={};
  values.forEach((r,i)=>{if(normalize_(r[7])!=='OBIETTIVO'||!String(r[0]||'').trim())return;let id=String(r[8]||'').trim();if(!id){id='OBJ-'+Utilities.getUuid();local.getRange(i+2,9).setValue(id);}const old=objById[id]||[],visibleColor=local.getRange(i+2,1).getBackground(),color=visibleColor&&visibleColor!=='#ffffff'?visibleColor:String(r[12]||old[3]||activityColorV14_(objs.length));objs.push([id,eventId,String(r[0]).trim(),color,r[1] instanceof Date?r[1]:'',Number(r[10]||old[5]||((objs.length+1)*10)),old[6]||'SCHEDA EVENTO',old[7]||('CUSTOM:'+normalize_(r[0])),String(r[3]||old[8]||'').trim(),old[9]||now,now]);validObjectives[id]=true;local.getRange(i+2,13).setValue(color);});
  values.forEach((r,i)=>{if(normalize_(r[7])!=='TASK'||!String(r[2]||'').trim()||!validObjectives[String(r[8]||'')])return;let id=String(r[9]||'').trim();if(!id){id='TASK-'+Utilities.getUuid();local.getRange(i+2,10).setValue(id);}const old=taskById[id]||[],done=r[4]===true,status=done?'FATTO':(normalize_(r[5])||'IN ATTESA'),oid=String(r[8]),step=Number(r[11]||10),objOrder=Number(r[10]||10),prev=String(r[14]||'').trim(),completed=done?(r[15] instanceof Date?r[15]:(old[11] instanceof Date?old[11]:now)):'',path=String(r[16]||old[20]||'MANUALE').trim()||'MANUALE',mode=normalize_(r[17]||old[21]||'MANUAL'),offset=r[18]!==''?Number(r[18]):(old[22]!==''?Number(old[22]):2);tasks.push([id,eventId,objOrder*100+step,String(r[2]).trim(),old[4]||activityTaskCategoryV14_(r[2]),r[1] instanceof Date?r[1]:'',status,old[7]||'',old[8]||'SCHEDA EVENTO',old[9]||'',String(r[3]||'').trim(),completed,now,old[13]||tasks.length+1,prev,prev&&status!=='FATTO'?'DIPENDENZA':'',oid,step,r[13]===true,prev,path,mode,offset]);});
  replaceCentralRowsForEvent_(b.objectives,eventId,2,objs,11);replaceCentralRowsForEvent_(b.checklist,eventId,2,tasks,23);const found=findCalendarEventById_(eventId);syncActivityStatesForEventV14_(eventId,found&&found.event);refreshCalendarActivityDashboardV14_(eventId);SpreadsheetApp.flush();return{objectives:objs.length,tasks:tasks.length};
}
