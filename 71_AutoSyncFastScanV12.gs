const CALENDAR_FAST_SYNC_V12 = Object.freeze({
  LAST_SCAN_KEY:'CALENDAR_V12_LAST_DRIVE_SCAN',
  OVERLAP_MS:120000,
  MAX_AUTO_SYNC:8
});

function setupCalendarAutoSyncFastV12() {
  const ui=SpreadsheetApp.getUi();
  calendarEnsureSyncMetaV12_();
  const baseline=calendarInitializeSyncBaselineV12_();
  const handlers=['calendarAutoSyncOnOpenV12','calendarAutoSyncTimerV12','calendarAutoSyncFastOnOpenV12','calendarAutoSyncFastTimerV12'];
  ScriptApp.getProjectTriggers().forEach(t=>{if(handlers.includes(t.getHandlerFunction()))ScriptApp.deleteTrigger(t);});
  PropertiesService.getScriptProperties().setProperty(CALENDAR_FAST_SYNC_V12.LAST_SCAN_KEY,new Date().toISOString());
  ScriptApp.newTrigger('calendarAutoSyncFastOnOpenV12').forSpreadsheet(SpreadsheetApp.getActive().getId()).onOpen().create();
  ScriptApp.newTrigger('calendarAutoSyncFastTimerV12').timeBased().everyMinutes(5).create();
  ui.alert('Sincronizzazione automatica attivata',
    'Baseline registrata per '+baseline+' Schede evento.\n\nDa ora il Calendario cerca solo le Schede evento modificate dall ultimo controllo: all apertura e ogni 5 minuti.',
    ui.ButtonSet.OK);
}

function disableCalendarAutoSyncFastV12() {
  const handlers=['calendarAutoSyncOnOpenV12','calendarAutoSyncTimerV12','calendarAutoSyncFastOnOpenV12','calendarAutoSyncFastTimerV12'];
  let removed=0;
  ScriptApp.getProjectTriggers().forEach(t=>{if(handlers.includes(t.getHandlerFunction())){ScriptApp.deleteTrigger(t);removed++;}});
  SpreadsheetApp.getUi().alert('Sincronizzazione automatica disattivata','Trigger rimossi: '+removed,SpreadsheetApp.getUi().ButtonSet.OK);
}

function calendarAutoSyncFastOnOpenV12(){
  try{calendarSyncRecentlyModifiedV12_();}catch(err){console.log('Auto-sync apertura: '+(err.message||err));}
}

function calendarAutoSyncFastTimerV12(){
  try{calendarSyncRecentlyModifiedV12_();}catch(err){console.log('Auto-sync timer: '+(err.message||err));}
}

function calendarSyncRecentlyModifiedV12_(){
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(1000))return{synced:0,errors:[]};
  try{
    calendarEnsureSyncMetaV12_();
    const props=PropertiesService.getScriptProperties(),now=new Date();
    const raw=props.getProperty(CALENDAR_FAST_SYNC_V12.LAST_SCAN_KEY);
    const previous=raw?new Date(raw):new Date(now.getTime()-10*60*1000);
    const since=new Date(Math.max(0,previous.getTime()-CALENDAR_FAST_SYNC_V12.OVERLAP_MS));
    const query="modifiedDate > '"+calendarDriveQueryDateV12_(since)+"' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false";
    const files=DriveApp.searchFiles(query),modifiedById={};
    while(files.hasNext()){
      const file=files.next();
      modifiedById[file.getId()]={modified:file.getLastUpdated()};
    }
    props.setProperty(CALENDAR_FAST_SYNC_V12.LAST_SCAN_KEY,now.toISOString());
    const meta=calendarMetaMapV12_(),events=calendarCalendarRowsV12_();
    let synced=0;const errors=[];
    for(let i=0;i<events.length&&synced<CALENDAR_FAST_SYNC_V12.MAX_AUTO_SYNC;i++){
      const item=events[i],hit=modifiedById[item.sheetId];
      if(!item.sheetId||!hit)continue;
      const old=meta[item.eventId],last=old&&old.lastModified instanceof Date?old.lastModified:null;
      if(last&&hit.modified.getTime()<=last.getTime()+1000)continue;
      try{
        const child=SpreadsheetApp.openById(item.sheetId);
        if(!calendarIsCurrentEventSheetV12_(child))continue;
        calendarSyncOneEventSheetV12_(item.eventId,item.event,child);
        calendarRecordSuccessfulSyncV12_(item.eventId,item.sheetId);
        synced++;
      }catch(err){
        errors.push(item.eventId+': '+(err.message||String(err)));
        calendarRecordSyncErrorV12_(item.eventId,err.message||String(err));
      }
    }
    return{synced:synced,errors:errors};
  }finally{lock.releaseLock();}
}

function calendarDriveQueryDateV12_(date){
  return Utilities.formatDate(date,'GMT',"yyyy-MM-dd'T'HH:mm:ss");
}
