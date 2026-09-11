const CALENDAR_SYNC_V12 = Object.freeze({
  META_SHEET: '_EVENTI_META',
  META_LAST_MODIFIED_COL: 9,
  META_LAST_SYNC_COL: 10,
  META_ERROR_COL: 11,
  MAX_AUTO_SYNC: 8,
  MAX_MANUAL_SYNC: 40,
  HISTORY_START_ROW: 20,
  HISTORY_COLS: 18,
  CONV_START: 3,
  CONV_COUNT: 20,
  AGG_START: 25,
  AGG_COUNT: 20
});

function setupCalendarAutoSyncV12() {
  const ui = SpreadsheetApp.getUi();
  calendarEnsureSyncMetaV12_();
  const baseline = calendarInitializeSyncBaselineV12_();
  const handlers = ['calendarAutoSyncOnOpenV12','calendarAutoSyncTimerV12'];
  ScriptApp.getProjectTriggers().forEach(t => {
    if (handlers.includes(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('calendarAutoSyncOnOpenV12').forSpreadsheet(SpreadsheetApp.getActive().getId()).onOpen().create();
  ScriptApp.newTrigger('calendarAutoSyncTimerV12').timeBased().everyMinutes(5).create();
  ui.alert('Sincronizzazione automatica attivata',
    'Baseline registrata per ' + baseline + ' Schede evento.\n\nDa ora il Calendario controlla le modifiche all apertura e ogni 5 minuti, sincronizzando solo le schede cambiate.',
    ui.ButtonSet.OK);
}

function disableCalendarAutoSyncV12() {
  const handlers = ['calendarAutoSyncOnOpenV12','calendarAutoSyncTimerV12'];
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (handlers.includes(t.getHandlerFunction())) { ScriptApp.deleteTrigger(t); removed++; }
  });
  SpreadsheetApp.getUi().alert('Sincronizzazione automatica disattivata', 'Trigger rimossi: ' + removed, SpreadsheetApp.getUi().ButtonSet.OK);
}

function calendarAutoSyncOnOpenV12() {
  try { syncChangedEventSheetsV12_({silent:true, limit:CALENDAR_SYNC_V12.MAX_AUTO_SYNC}); }
  catch (err) { console.log('Auto-sync apertura: ' + (err.message || err)); }
}

function calendarAutoSyncTimerV12() {
  try { syncChangedEventSheetsV12_({silent:true, limit:CALENDAR_SYNC_V12.MAX_AUTO_SYNC}); }
  catch (err) { console.log('Auto-sync timer: ' + (err.message || err)); }
}

function syncChangedEventSheetsV12() {
  return syncChangedEventSheetsV12_({silent:false, limit:CALENDAR_SYNC_V12.MAX_MANUAL_SYNC});
}

function syncSelectedEventSheetToCalendarV2() {
  const event = selectedEvent_();
  const eventId = ensureEventId_(event);
  const child = calendarLinkedEventSheetV12_(event);
  if (!child) throw new Error('La riga selezionata non ha una Scheda evento collegata.');
  if (!calendarIsCurrentEventSheetV12_(child)) throw new Error('La Scheda evento collegata non usa il layout corrente V11.');
  const result = calendarSyncOneEventSheetV12_(eventId, event, child);
  calendarRecordSuccessfulSyncV12_(eventId, child.getId());
  SpreadsheetApp.getActive().toast(
    'Attivita: ' + result.tasks + ' | Partecipanti: ' + result.participants + ' | Movimenti: ' + result.expenses,
    'Scheda sincronizzata', 6);
  return result;
}

function syncChangedEventSheetsV12_(options) {
  options = options || {};
  const silent = options.silent === true;
  const limit = Number(options.limit || CALENDAR_SYNC_V12.MAX_AUTO_SYNC);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return {synced:0, skipped:0, errors:[]};
  try {
    calendarEnsureSyncMetaV12_();
    const candidates = calendarSyncCandidatesV12_();
    let synced = 0, skipped = 0;
    const errors = [];
    for (let i = 0; i < candidates.length && synced < limit; i++) {
      const item = candidates[i];
      if (!item.changed) { skipped++; continue; }
      try {
        const child = SpreadsheetApp.openById(item.sheetId);
        if (!calendarIsCurrentEventSheetV12_(child)) { skipped++; continue; }
        calendarSyncOneEventSheetV12_(item.eventId, item.event, child);
        calendarRecordSuccessfulSyncV12_(item.eventId, item.sheetId);
        synced++;
      } catch (err) {
        errors.push(item.eventId + ': ' + (err.message || String(err)));
        calendarRecordSyncErrorV12_(item.eventId, err.message || String(err));
      }
    }
    if (!silent) SpreadsheetApp.getActive().toast(
      'Aggiornate: ' + synced + (errors.length ? ' | Errori: ' + errors.length : ''),
      'Sincronizzazione completata', 7);
    return {synced:synced, skipped:skipped, errors:errors};
  } finally {
    lock.releaseLock();
  }
}

function calendarSyncOneEventSheetV12_(eventId, event, child) {
  calendarValidateEventIdentityV12_(child, eventId);
  const participants = calendarSyncParticipantsV12_(eventId, child);
  const tasks = calendarSyncTasksV12_(eventId, child);
  const expenses = calendarSyncExpensesV12_(eventId, event, child);
  const meta = child.getSheetByName('_META');
  if (meta) writeMeta_(meta, {LAST_SYNC:Utilities.formatDate(new Date(), APP.TZ, 'dd/MM/yyyy HH:mm')});
  SpreadsheetApp.flush();
  return {participants:participants, tasks:tasks, expenses:expenses};
}

function calendarSyncParticipantsV12_(eventId, child) {
  const local = child.getSheetByName('Partecipanti');
  if (!local) throw new Error('Foglio Partecipanti non trovato.');
  const backend = sh_(APP.SHEETS.PARTICIPANTS);
  const oldRows = backend.getDataRange().getValues();
  const oldById = {};
  for (let i=1;i<oldRows.length;i++) if (String(oldRows[i][1]||'')===String(eventId) && oldRows[i][0]) oldById[String(oldRows[i][0])] = oldRows[i];
  const now = new Date(), rows = [], convTech = [], aggTech = [];
  const conv = local.getRange(CALENDAR_SYNC_V12.CONV_START,1,CALENDAR_SYNC_V12.CONV_COUNT,19).getValues();
  const agg = local.getRange(CALENDAR_SYNC_V12.AGG_START,1,CALENDAR_SYNC_V12.AGG_COUNT,19).getValues();
  conv.forEach((r,index)=>{
    const parsed = calendarParticipantRowV12_(r,'CONVOCATO',eventId,oldById,now,CALENDAR_SYNC_V12.CONV_START+index);
    if (!parsed) { convTech.push(['','','','','','',''].slice(0,6)); return; }
    rows.push(parsed.backend); convTech.push(parsed.tech);
  });
  agg.forEach((r,index)=>{
    const parsed = calendarParticipantRowV12_(r,'AGGREGATO',eventId,oldById,now,CALENDAR_SYNC_V12.AGG_START+index);
    if (!parsed) { aggTech.push(['','','','','','',''].slice(0,6)); return; }
    rows.push(parsed.backend); aggTech.push(parsed.tech);
  });
  replaceCentralRowsForEvent_(backend,eventId,2,rows,17);
  local.getRange(CALENDAR_SYNC_V12.CONV_START,14,convTech.length,6).setValues(convTech);
  local.getRange(CALENDAR_SYNC_V12.AGG_START,14,aggTech.length,6).setValues(aggTech);
  return rows.length;
}

function calendarParticipantRowV12_(r,type,eventId,oldById,now,rowNumber) {
  const name=String(r[0]||'').trim(), surname=String(r[1]||'').trim();
  if (!name && !surname) return null;
  let id=String(r[13]||'').trim();
  const old=id&&oldById[id]?oldById[id]:null;
  if (!id) id='PAR-'+Utilities.getUuid();
  let personId=String(r[14]||'').trim()||(old?String(old[2]||''):'');
  let club=String(r[2]||'').trim(), email=String(r[6]||'').trim();
  if (!personId) {
    const known=findKnownPerson_(name,surname,'');
    if (known) { personId=known.id||''; if(!club)club=known.club||''; if(!email)email=known.email||''; }
  }
  const role=normalize_(r[7]||'ATLETA')||'ATLETA';
  let status=normalize_(r[8]);
  if (type==='CONVOCATO' && !['DA FARE','MANDATA CONVOCAZIONE','CONFERMATO','ASSENTE'].includes(status)) status=role==='TECNICO'?'CONFERMATO':'DA FARE';
  if (type==='AGGREGATO' && !['DA AUTORIZZARE','AUTORIZZATO','NON AUTORIZZATO'].includes(status)) status='DA AUTORIZZARE';
  const maxRefund=type==='CONVOCATO'&&r[9]!==''&&role!=='TECNICO'?Number(r[9]||0):'';
  const passed=type==='CONVOCATO'&&r[10]!==''&&role!=='TECNICO'?Number(r[10]||0):'';
  const passedDate=passed>0?(r[11] instanceof Date?r[11]:(old&&old[16] instanceof Date?old[16]:now)):'';
  const notes=String(type==='CONVOCATO'?(r[12]||''):(r[9]||'')).trim();
  const provenance=String(r[16]||'').trim()||(old?String(old[9]||''):'')||'SCHEDA EVENTO';
  const sourceFile=String(r[17]||'').trim()||(old?String(old[10]||''):'');
  const createdAt=r[18] instanceof Date?r[18]:(old&&old[12] instanceof Date?old[12]:now);
  return {
    backend:[id,eventId,personId,name,surname,club,role,email,maxRefund,provenance,sourceFile,notes,createdAt,type,status,passed,passedDate],
    tech:[id,personId,eventId,provenance,sourceFile,createdAt],
    rowNumber:rowNumber
  };
}

function calendarSyncTasksV12_(eventId, child) {
  const local=child.getSheetByName('Attività');
  if (!local) throw new Error('Foglio Attivita non trovato.');
  const backend=sh_(APP.SHEETS.CHECKLIST), oldRows=backend.getDataRange().getValues(), oldByNo={};
  let maxNo=0;
  for(let i=1;i<oldRows.length;i++){
    if(String(oldRows[i][1]||'')!==String(eventId))continue;
    const n=Number(oldRows[i][13]||0); if(n>0)oldByNo[n]=oldRows[i]; if(n>maxNo)maxNo=n;
  }
  const lastRow=Math.min(Math.max(local.getLastRow(),2),500), values=local.getRange(2,1,lastRow-1,6).getValues(), draft=[];
  values.forEach((r,i)=>{const description=String(r[0]||'').trim();if(!description)return;let no=Number(r[2]||0);if(!(no>0)){no=++maxNo;local.getRange(i+2,3).setValue(no);}maxNo=Math.max(maxNo,no);draft.push({row:i+2,values:r,no:no});});
  const seen={}; draft.forEach(x=>{if(seen[x.no])throw new Error('Numero task duplicato: '+x.no+'.');seen[x.no]=true;});
  const idByNo={}; draft.forEach(x=>{const old=oldByNo[x.no];x.old=old?old.slice(0,16):new Array(16).fill('');x.id=old&&old[0]?String(old[0]):'TASK-'+Utilities.getUuid();idByNo[x.no]=x.id;});
  const now=new Date();
  const rows=draft.map((x,index)=>{
    const r=x.values,old=x.old,depNo=Number(r[3]||0);
    if(depNo&&!idByNo[depNo])throw new Error('La task n. '+x.no+' dipende dalla task n. '+depNo+', che non esiste.');
    if(depNo===x.no)throw new Error('La task n. '+x.no+' non puo dipendere da se stessa.');
    let status=normalize_(r[4])||'DA FARE';if(!['DA FARE','IN ATTESA','FATTO'].includes(status))status='DA FARE';
    const completed=status==='FATTO'?(old[11] instanceof Date?old[11]:now):'';
    const description=String(r[0]||'').trim(), autoKey=old[9]||inferTaskAutoKeyV3_(description), source=old[8]||(autoKey?'AUTO':'MANUALE');
    return[x.id,eventId,(index+1)*10,description,old[4]||inferTaskProcessV3_(description,'ALTRO',autoKey),r[5] instanceof Date?r[5]:'',status,old[7]||'',source,autoKey,String(r[1]||'').trim(),completed,now,x.no,depNo?idByNo[depNo]:'',status!=='FATTO'&&depNo?'DIPENDENZA':''];
  });
  replaceCentralRowsForEvent_(backend,eventId,2,rows,16);
  return rows.length;
}

function calendarSyncExpensesV12_(eventId,event,child) {
  const local=child.getSheetByName('Spese');
  if (!local) throw new Error('Foglio Spese non trovato.');
  const backend=sh_(APP.SHEETS.EXPENSES), oldRows=backend.getDataRange().getValues(), oldById={};
  for(let i=1;i<oldRows.length;i++)if(String(oldRows[i][1]||'')===String(eventId)&&oldRows[i][0])oldById[String(oldRows[i][0])]=oldRows[i];
  const count=Math.max(0,local.getMaxRows()-CALENDAR_SYNC_V12.HISTORY_START_ROW+1);
  const history=count?local.getRange(CALENDAR_SYNC_V12.HISTORY_START_ROW,1,count,CALENDAR_SYNC_V12.HISTORY_COLS).getValues():[];
  const rows=[], now=new Date(), quotesByNo={};
  history.forEach(r=>{
    if(normalize_(r[1])!=='PREVENTIVO')return;
    const no=Number(r[0]||0); if(!(no>0))return;
    const id=String(r[12]||'').trim()||'PREV-'+Utilities.getUuid();
    quotesByNo[no]={id:id,budget:Number(r[2]||0),paid:0};
  });
  history.forEach(r=>{
    const no=Number(r[0]||0), movement=normalize_(r[1]);
    if(!quotesByNo[no]||movement==='PREVENTIVO'||normalize_(r[4])!=='PAGATO')return;
    quotesByNo[no].paid+=Number(r[2]||0);
  });
  const techUpdates=[];
  history.forEach((r,index)=>{
    const quoteNo=Number(r[0]||0), movement=normalize_(r[1]), amount=Number(r[2]||0), description=String(r[3]||'').trim();
    if(!(quoteNo>0)||!movement||!description)return;
    const q=quotesByNo[quoteNo];
    const id=String(r[12]||'').trim()||(movement==='PREVENTIVO'?(q?q.id:'PREV-'+Utilities.getUuid()):'PAG-'+Utilities.getUuid());
    const quoteId=String(r[13]||'').trim()||(movement==='PREVENTIVO'?id:(q?q.id:''));
    const old=oldById[id]?oldById[id].slice(0,26):new Array(26).fill('');
    const category=calendarExpenseCategoryV12_(r[11],description), paid=movement!=='PREVENTIVO'&&normalize_(r[4])==='PAGATO';
    const paidDate=paid&&r[5] instanceof Date?r[5]:'', due=!paid&&r[6] instanceof Date?r[6]:'', rif=String(r[7]||'').trim(), imp=String(r[8]||'').trim(), documents=String(r[9]||'').trim(), note=String(r[10]||'').trim();
    const ceb=String(r[14]||'').trim()||old[4]||resolveExpenseCeb_(event,'',category,'SPESA');
    const createdAt=r[16] instanceof Date?r[16]:(old[16] instanceof Date?old[16]:now);
    old[0]=id; old[1]=eventId; old[2]=movement==='PREVENTIVO'?'PREVENTIVO':'PAGAMENTO'; old[3]=category; old[4]=ceb; old[5]=description; old[6]=''; old[7]=String(r[15]||'').trim();
    old[8]=movement==='PREVENTIVO'?amount:0; old[9]=movement!=='PREVENTIVO'&&paid?amount:0; old[10]=due;
    if(movement==='PREVENTIVO'){
      const info=q||{budget:amount,paid:0}; old[11]=info.paid<=0.005?'DA SALDARE':(info.paid+0.005>=info.budget?'SALDATO':'PARZIALMENTE SALDATO');
    } else old[11]=paid?calendarBackendPaidStatusV12_(movement):'DA PAGARE';
    old[12]=paidDate; old[13]='PREV. '+quoteNo; old[14]=documents;
    old[15]='[MOVIMENTO='+movement+'] [ID_PREVENTIVO='+quoteId+'] [N_PREVENTIVO='+quoteNo+']'+(imp?' [IMP='+imp+']':'')+(note?' '+note:'');
    old[16]=createdAt; old[17]=now;
    if(rif)old[18]='RICEVUTO'; else if(movement==='PREVENTIVO'&&amount>1000)old[18]=old[18]||'DA RICHIEDERE'; else old[18]=old[18]||'NON NECESSARIO';
    old[20]=rif; old[21]=movement==='SALDO FATTURA'&&paid; old[22]=old[21]?paidDate:(old[22]||'');
    old[25]=paid&&movement!=='PREVENTIVO'?(old[25]||now):(movement==='PREVENTIVO'&&q&&q.paid+0.005>=q.budget?(old[25]||now):'');
    rows.push(old);
    techUpdates.push({row:CALENDAR_SYNC_V12.HISTORY_START_ROW+index,id:id,quoteId:quoteId,ceb:ceb,createdAt:createdAt,updatedAt:now});
  });
  const participants=child.getSheetByName('Partecipanti');
  if(participants)participants.getRange(CALENDAR_SYNC_V12.CONV_START,1,CALENDAR_SYNC_V12.CONV_COUNT,19).getValues().forEach((p,index)=>{
    const name=String(p[0]||'').trim(),surname=String(p[1]||'').trim(),role=normalize_(p[7]),passed=Number(p[10]||0);
    if(role==='TECNICO'||((!name&&!surname)||!(passed>0)))return;
    const participantId=String(p[13]||'').trim()||('ROW-'+(index+CALENDAR_SYNC_V12.CONV_START)),id='RIMBORSO-AUTO-'+participantId,old=oldById[id]?oldById[id].slice(0,26):new Array(26).fill(''),beneficiary=[name,surname].filter(Boolean).join(' ');
    old[0]=id;old[1]=eventId;old[2]='RIMBORSO';old[3]='RIMBORSO';old[4]='CEB.002';old[5]='Rimborso '+beneficiary;old[6]=beneficiary;old[7]=String(p[14]||'').trim();old[8]=0;old[9]=passed;old[10]='';old[11]='RIMBORSATO';old[12]=p[11] instanceof Date?p[11]:now;old[13]='';old[14]='';old[15]='[MOVIMENTO=RIMBORSO] [AUTO_RIMBORSO='+participantId+']'+(String(p[12]||'').trim()?' '+String(p[12]||'').trim():'');old[16]=old[16]||now;old[17]=now;old[18]='NON NECESSARIO';old[20]='';old[21]=false;old[25]=old[25]||now;rows.push(old);
  });
  replaceCentralRowsForEvent_(backend,eventId,2,rows,26);
  techUpdates.forEach(x=>local.getRange(x.row,13,1,6).setValues([[x.id,x.quoteId,x.ceb,'',x.createdAt,x.updatedAt]]));
  return rows.length;
}

function calendarExpenseCategoryV12_(category,description) {
  let c=normalize_(category),d=normalize_(description);
  if(c==='VITTO / ALLOGGIO')c=/(PASTI|VITTO|RISTORANTE|PRANZO|CENA)/.test(d)?'VITTO':'ALLOGGIO';
  if(c==='VIAGGIO'||c==='TRASPORTO')c='VIAGGI';
  return ['VIAGGI','VITTO','ALLOGGIO','NOLEGGI','ISCRIZIONI','ALTRO'].includes(c)?c:'ALTRO';
}

function calendarBackendPaidStatusV12_(movement) {
  if(movement==='AFOR')return'PAGATO - AFOR';
  if(movement==='SALDO FATTURA')return'PAGATO - FATTURA';
  if(movement==='CARTA DI CREDITO')return'PAGATO CON CC';
  return'PAGATO';
}

function calendarEnsureSyncMetaV12_() {
  const sheet=sh_(CALENDAR_SYNC_V12.META_SHEET);
  if(sheet.getMaxColumns()<CALENDAR_SYNC_V12.META_ERROR_COL)sheet.insertColumnsAfter(sheet.getMaxColumns(),CALENDAR_SYNC_V12.META_ERROR_COL-sheet.getMaxColumns());
  sheet.getRange(1,CALENDAR_SYNC_V12.META_LAST_MODIFIED_COL,1,3).setValues([['ULTIMA MODIFICA SCHEDA SINCRONIZZATA','ULTIMO SYNC SCHEDA','ERRORE SYNC SCHEDA']]);
  sheet.getRange(1,CALENDAR_SYNC_V12.META_LAST_MODIFIED_COL,1,3).setFontWeight('bold').setBackground('#eeeeee').setWrap(true);
  sheet.getRange(2,CALENDAR_SYNC_V12.META_LAST_MODIFIED_COL,Math.max(sheet.getMaxRows()-1,1),2).setNumberFormat('dd/MM/yyyy HH:mm:ss');
}

function calendarInitializeSyncBaselineV12_() {
  const items=calendarCalendarRowsV12_(), meta=calendarMetaMapV12_();
  let count=0;
  items.forEach(item=>{
    if(!item.sheetId)return;
    try{
      const modified=DriveApp.getFileById(item.sheetId).getLastUpdated();
      const row=calendarEnsureEventMetaRowV12_(item.eventId,meta);
      sh_(CALENDAR_SYNC_V12.META_SHEET).getRange(row,CALENDAR_SYNC_V12.META_LAST_MODIFIED_COL).setValue(modified);
      count++;
    }catch(e){}
  });
  return count;
}

function calendarSyncCandidatesV12_() {
  const items=calendarCalendarRowsV12_(), meta=calendarMetaMapV12_();
  return items.map(item=>{
    if(!item.sheetId)return Object.assign(item,{changed:false});
    try{
      const modified=DriveApp.getFileById(item.sheetId).getLastUpdated();
      const m=meta[item.eventId], syncedModified=m&&m.lastModified instanceof Date?m.lastModified:null;
      return Object.assign(item,{changed:!syncedModified||modified.getTime()>syncedModified.getTime()+1000,modified:modified});
    }catch(err){return Object.assign(item,{changed:false,error:err.message||String(err)});}
  });
}

function calendarCalendarRowsV12_() {
  const cal=sh_(APP.SHEETS.CALENDAR), last=cal.getLastRow();
  if(last<2)return[];
  const map=headerMap_(cal), idCol=map[APP.CALENDAR_HEADERS.ID], linkCol=map[APP.CALENDAR_HEADERS.EVENT_SHEET];
  const values=cal.getRange(2,1,last-1,cal.getLastColumn()).getValues();
  const rich=linkCol?cal.getRange(2,linkCol,last-1,1).getRichTextValues():[];
  const headers=cal.getRange(1,1,1,cal.getLastColumn()).getDisplayValues()[0];
  const out=[];
  values.forEach((r,i)=>{
    const eventId=String(r[idCol-1]||'').trim(); if(!eventId)return;
    let url=''; if(linkCol){const rt=rich[i]&&rich[i][0];url=(rt&&rt.getLinkUrl())||String(r[linkCol-1]||'');}
    const sheetId=extractDriveId_(url); const event={}; headers.forEach((h,c)=>event[h]=r[c]); event._row=i+2;
    out.push({eventId:eventId,event:event,sheetId:sheetId||'',row:i+2});
  });
  return out;
}

function calendarMetaMapV12_() {
  const sheet=sh_(CALENDAR_SYNC_V12.META_SHEET), values=sheet.getDataRange().getValues(), out={};
  for(let i=1;i<values.length;i++){
    const id=String(values[i][0]||'').trim();if(!id)continue;
    out[id]={row:i+1,lastModified:values[i][CALENDAR_SYNC_V12.META_LAST_MODIFIED_COL-1],lastSync:values[i][CALENDAR_SYNC_V12.META_LAST_SYNC_COL-1]};
  }
  return out;
}

function calendarEnsureEventMetaRowV12_(eventId,map) {
  if(map[eventId])return map[eventId].row;
  const sheet=sh_(CALENDAR_SYNC_V12.META_SHEET),row=sheet.getLastRow()+1;
  sheet.getRange(row,1).setValue(eventId); map[eventId]={row:row,lastModified:'',lastSync:''}; return row;
}

function calendarRecordSuccessfulSyncV12_(eventId,sheetId) {
  const meta=calendarMetaMapV12_(),row=calendarEnsureEventMetaRowV12_(eventId,meta),sheet=sh_(CALENDAR_SYNC_V12.META_SHEET);
  const modified=DriveApp.getFileById(sheetId).getLastUpdated(),now=new Date();
  sheet.getRange(row,CALENDAR_SYNC_V12.META_LAST_MODIFIED_COL,1,3).setValues([[modified,now,'']]);
}

function calendarRecordSyncErrorV12_(eventId,message) {
  const meta=calendarMetaMapV12_(),row=calendarEnsureEventMetaRowV12_(eventId,meta);
  sh_(CALENDAR_SYNC_V12.META_SHEET).getRange(row,CALENDAR_SYNC_V12.META_ERROR_COL).setValue(String(message||'').slice(0,500));
}

function calendarLinkedEventSheetV12_(event) {
  const cal=sh_(APP.SHEETS.CALENDAR),map=headerMap_(cal),col=map[APP.CALENDAR_HEADERS.EVENT_SHEET];
  if(!col)return null;
  const cell=cal.getRange(event._row,col),rich=cell.getRichTextValue(),id=extractDriveId_((rich&&rich.getLinkUrl())||cell.getDisplayValue());
  if(!id)return null;
  return SpreadsheetApp.openById(id);
}

function calendarIsCurrentEventSheetV12_(child) {
  const expense=child.getSheetByName('Spese'),participants=child.getSheetByName('Partecipanti'),meta=child.getSheetByName('_META');
  if(!expense||!participants||!meta)return false;
  return normalize_(expense.getRange('B12').getDisplayValue())==='N. PAGAMENTO' && normalize_(participants.getRange('H2').getDisplayValue())==='RUOLO';
}

function calendarValidateEventIdentityV12_(child,eventId) {
  const meta=child.getSheetByName('_META'); if(!meta)throw new Error('Foglio _META mancante.');
  const stored=String(readMetaValue_(meta,'EVENT_ID')||'').trim();
  if(stored&&stored!==String(eventId))throw new Error('La Scheda appartiene all evento '+stored+', non a '+eventId+'.');
}
