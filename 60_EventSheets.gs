const EVENT_SHEET = Object.freeze({
  TEMPLATE_ID:'1obNkubjpGhC7o8AX5Tc986UBGdyJs26b476b-GZN3Sc',
  SHEETS:Object.freeze({TASKS:'Attività',EXPENSES:'Spese',PARTICIPANTS:'Partecipanti',DOCUMENTS:'DOCUMENTI',META:'_META'}),
  PARTICIPANTS:Object.freeze({CONV_START:3,CONV_COUNT:20,AGG_START:25,AGG_COUNT:20})
});

function prepareEventSheetForSelectedEventV2(){
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

  if(!isCurrentEventSheet_(child))throw new Error('La Scheda evento collegata non usa il layout V11 corrente. Non viene modificata automaticamente.');
  validateEventSheetIdentity_(child,eventId);
  writeEventMeta_(child,eventId,event,folder.folderId);
  applyEventCommitment_(child,event);

  if(created){
    generateChecklistForEvent_(eventId,event);
    ensureCurrentTaskNumbersAndDependencies_(eventId,event);
    writeCurrentTasksToEventSheet_(eventId,child);
    writeCurrentParticipantsToEventSheet_(eventId,child);
  }
  populateCurrentTechnicians_(child,event);
  setEventSheetLink_(event._row,child.getUrl());
  SpreadsheetApp.flush();
  SpreadsheetApp.getActive().toast(created?'Scheda evento creata':'Dati evento aggiornati','Scheda evento',5);
  return{created:created,id:child.getId(),url:child.getUrl()};
}

function isCurrentEventSheet_(child){
  if(!child)return false;
  const expenses=child.getSheetByName(EVENT_SHEET.SHEETS.EXPENSES),participants=child.getSheetByName(EVENT_SHEET.SHEETS.PARTICIPANTS),meta=child.getSheetByName(EVENT_SHEET.SHEETS.META);
  if(!expenses||!participants||!meta)return false;
  return normalize_(expenses.getRange('A8').getDisplayValue())==='N. PREVENTIVO'&&
    normalize_(expenses.getRange('B12').getDisplayValue())==='N. PAGAMENTO'&&
    normalize_(expenses.getRange('A19').getDisplayValue())==='N. PREVENTIVO'&&
    normalize_(participants.getRange('H2').getDisplayValue())==='RUOLO';
}

function getLinkedEventSheet_(event){
  if(!event||!event._row)return null;
  const cal=sh_(APP.SHEETS.CALENDAR),map=headerMap_(cal),col=map[APP.CALENDAR_HEADERS.EVENT_SHEET];
  if(!col)return null;
  const cell=cal.getRange(event._row,col),rich=cell.getRichTextValue(),id=extractDriveId_((rich&&rich.getLinkUrl())||cell.getDisplayValue());
  if(!id)return null;
  try{return SpreadsheetApp.openById(id);}catch(e){return null;}
}

function writeEventMeta_(child,eventId,event,folderId){
  const meta=child.getSheetByName(EVENT_SHEET.SHEETS.META);
  if(!meta)throw new Error('Foglio _META non trovato nella Scheda evento.');
  const start=event[APP.CALENDAR_HEADERS.START],end=event[APP.CALENDAR_HEADERS.END];
  const commitment=resolveCommitmentCode_(event[APP.CALENDAR_HEADERS.TYPE],event[APP.CALENDAR_HEADERS.CLASS])||String(event[APP.CALENDAR_HEADERS.COMMITMENT]||'').trim();
  const technicians=resolveFullTechnicians_(event[APP.CALENDAR_HEADERS.TECHNICIANS]);
  writeMeta_(meta,{
    EVENT_ID:eventId,
    MASTER_SPREADSHEET_ID:APP.SPREADSHEET_ID,
    EVENT_FOLDER_ID:folderId,
    EVENT_SHEET_ID:child.getId(),
    SYNC_VERSION:'11',
    EVENT_LABEL:buildEventSheetLabel_(event),
    EVENT_TYPE:String(event[APP.CALENDAR_HEADERS.TYPE]||''),
    EVENT_CLASS:String(event[APP.CALENDAR_HEADERS.CLASS]||''),
    EVENT_LOCATION:String(event[APP.CALENDAR_HEADERS.LOCATION]||''),
    EVENT_ZONE:String(event[APP.CALENDAR_HEADERS.ZONE]||''),
    EVENT_CLUB:String(event[APP.CALENDAR_HEADERS.CLUB]||''),
    EVENT_TECHNICIANS:technicians.length?technicians.map(x=>x.name+' '+x.surname).join(', '):String(event[APP.CALENDAR_HEADERS.TECHNICIANS]||''),
    EVENT_LODGING:String(event[APP.CALENDAR_HEADERS.LODGING]||''),
    EVENT_COMMITMENT:commitment,
    EVENT_START:start instanceof Date?Utilities.formatDate(start,APP.TZ,'yyyy-MM-dd'):'',
    EVENT_END:end instanceof Date?Utilities.formatDate(end,APP.TZ,'yyyy-MM-dd'):''
  });
  meta.hideSheet();
}

function applyEventCommitment_(child,event){
  const commitment=resolveCommitmentCode_(event[APP.CALENDAR_HEADERS.TYPE],event[APP.CALENDAR_HEADERS.CLASS])||String(event[APP.CALENDAR_HEADERS.COMMITMENT]||'').trim();
  if(!commitment)return;
  const cal=sh_(APP.SHEETS.CALENDAR),map=headerMap_(cal),col=map[APP.CALENDAR_HEADERS.COMMITMENT];
  if(col&&event._row&&!String(cal.getRange(event._row,col).getDisplayValue()||'').trim())cal.getRange(event._row,col).setNumberFormat('@').setValue(String(commitment));
  const expenses=child.getSheetByName(EVENT_SHEET.SHEETS.EXPENSES);
  if(expenses&&!String(expenses.getRange('G9').getDisplayValue()||'').trim())expenses.getRange('G9').setNumberFormat('@').setValue(String(commitment));
}

function ensureCurrentTaskNumbersAndDependencies_(eventId,event){
  const sheet=sh_(APP.SHEETS.CHECKLIST),values=sheet.getDataRange().getValues(),items=[],used=new Set();
  sheet.getRange(1,14,1,3).setValues([['N TASK','ID DIPENDENZA','BLOCCO AUTO']]);
  for(let i=1;i<values.length;i++){
    if(String(values[i][1]||'')!==String(eventId))continue;
    const item={row:i+1,values:values[i]};items.push(item);
    const n=Number(values[i][13]||0);if(n>0)used.add(n);
  }
  items.sort((a,b)=>Number(a.values[2]||0)-Number(b.values[2]||0)||a.row-b.row);
  let next=1;
  items.forEach(x=>{let n=Number(x.values[13]||0);if(n>0)return;while(used.has(next))next++;n=next++;used.add(n);sheet.getRange(x.row,14).setValue(n);x.values[13]=n;});
  const byKey={};items.forEach(x=>{const key=normalize_(x.values[9]||x.values[3]);if(key)byKey[key]=x;});
  const rules=getChecklistRulesForEvent_(event);
  items.forEach(x=>{
    if(String(x.values[14]||'').trim())return;
    const key=normalize_(x.values[9]||x.values[3]),rule=rules[key];
    if(!rule||!rule.dependsOn)return;
    const parent=byKey[normalize_(rule.dependsOn)];if(!parent)return;
    sheet.getRange(x.row,15).setValue(String(parent.values[0]||''));x.values[14]=String(parent.values[0]||'');
  });
}

function writeCurrentTasksToEventSheet_(eventId,child){
  const backend=sh_(APP.SHEETS.CHECKLIST).getDataRange().getValues().slice(1).filter(r=>String(r[1]||'')===String(eventId));
  const byId={};backend.forEach(r=>{if(r[0])byId[String(r[0])]=r;});
  backend.sort((a,b)=>Number(a[2]||999999)-Number(b[2]||999999)||Number(a[13]||999999)-Number(b[13]||999999));
  const rows=backend.map(r=>{const parent=r[14]?byId[String(r[14])]:null;return[r[3]||'',r[10]||'',r[13]||'',parent?parent[13]||'':'',normalize_(r[6])==='COMPLETATA'?'FATTO':(r[6]||'DA FARE'),r[5]||''];});
  const sheet=child.getSheetByName(EVENT_SHEET.SHEETS.TASKS);if(!sheet)throw new Error('Foglio Attività non trovato.');
  const last=Math.min(500,sheet.getMaxRows());if(last>=2)sheet.getRange(2,1,last-1,6).clearContent();
  if(rows.length)sheet.getRange(2,1,rows.length,6).setValues(rows);
}

function writeCurrentParticipantsToEventSheet_(eventId,child){
  const sheet=child.getSheetByName(EVENT_SHEET.SHEETS.PARTICIPANTS);if(!sheet)return;
  const rows=sh_(APP.SHEETS.PARTICIPANTS).getDataRange().getValues().slice(1).filter(r=>String(r[1]||'')===String(eventId));
  const conv=rows.filter(r=>normalize_(r[13]||'CONVOCATO')!=='AGGREGATO').slice(0,EVENT_SHEET.PARTICIPANTS.CONV_COUNT);
  const agg=rows.filter(r=>normalize_(r[13])==='AGGREGATO').slice(0,EVENT_SHEET.PARTICIPANTS.AGG_COUNT);
  sheet.getRange(EVENT_SHEET.PARTICIPANTS.CONV_START,1,EVENT_SHEET.PARTICIPANTS.CONV_COUNT,19).clearContent();
  sheet.getRange(EVENT_SHEET.PARTICIPANTS.AGG_START,1,EVENT_SHEET.PARTICIPANTS.AGG_COUNT,19).clearContent();
  if(conv.length)sheet.getRange(EVENT_SHEET.PARTICIPANTS.CONV_START,1,conv.length,19).setValues(conv.map(r=>participantBackendToLocal_(r,false)));
  if(agg.length)sheet.getRange(EVENT_SHEET.PARTICIPANTS.AGG_START,1,agg.length,19).setValues(agg.map(r=>participantBackendToLocal_(r,true)));
}

function participantBackendToLocal_(r,isAggregated){
  if(isAggregated)return[r[3]||'',r[4]||'',r[5]||'','','','',r[7]||'',r[6]||'ATLETA',r[14]||'DA AUTORIZZARE',r[11]||'','','','',r[0]||'',r[2]||'',r[1]||'',r[9]||'',r[10]||'',r[12]||''];
  return[r[3]||'',r[4]||'',r[5]||'','','','',r[7]||'',r[6]||'ATLETA',r[14]||'DA FARE',r[8]||'',r[15]||'',r[16]||'',r[11]||'',r[0]||'',r[2]||'',r[1]||'',r[9]||'',r[10]||'',r[12]||''];
}

function technicianDirectory_(){return[
  ['ZAGGIA','Leonardo','Zaggia'],['CRISI','Andrea','Crisi'],['RAVEGLIA','Matteo','Raveglia'],['CARICATO','Francesco','Caricato'],['PICCIAU','Gianluigi','Picciau'],['SENSINI','Alessandra','Sensini'],['NUICOLUCCI','Matteo','Nuicolucci'],['CAMBONI','Mattia','Camboni'],['CANGEMI','Antonino','Cangemi'],['LOPERFIDO','Daniel','Loperfido']
];}
function resolveFullTechnicians_(raw){const text=normalize_(raw);if(!text)return[];return technicianDirectory_().filter(x=>text.indexOf(x[0])>=0).map(x=>({name:x[1],surname:x[2]}));}
function populateCurrentTechnicians_(child,event){
  const people=resolveFullTechnicians_(event[APP.CALENDAR_HEADERS.TECHNICIANS]);if(!people.length)return;
  const sheet=child.getSheetByName(EVENT_SHEET.SHEETS.PARTICIPANTS);if(!sheet)return;
  const start=EVENT_SHEET.PARTICIPANTS.CONV_START,count=EVENT_SHEET.PARTICIPANTS.CONV_COUNT,rows=sheet.getRange(start,1,count,19).getValues();
  const existing=new Set(rows.map(r=>normalize_((r[0]||'')+' '+(r[1]||''))));
  people.forEach(p=>{
    const key=normalize_(p.name+' '+p.surname);if(existing.has(key))return;
    const idx=rows.findIndex(r=>!String(r[0]||'').trim()&&!String(r[1]||'').trim());if(idx<0)return;
    const row=start+idx;sheet.getRange(row,1).setValue(p.name);sheet.getRange(row,2).setValue(p.surname);sheet.getRange(row,8).setValue('TECNICO');sheet.getRange(row,9).setValue('CONFERMATO');rows[idx][0]=p.name;rows[idx][1]=p.surname;existing.add(key);
  });
}

function inferTaskAutoKey_(task){
  const t=normalize_(task),map={'CONVOCAZIONE ATLETI':'CONV_ATLETI','CONVOCAZIONE TECNICO':'CONV_TECNICO','CHECK CONFERMA PRESENZE':'CHECK_CONFERME','CONFERMA PRESENZE TUTTI ATLETI':'CHECK_CONFERME','RINGRAZIAMENTO CIRCOLO':'RINGRAZIAMENTO_CIRCOLO'};
  return map[t]||'';
}
function inferTaskProcess_(task,category,autoKey){
  const key=normalize_(autoKey||task);
  if(key.indexOf('GOMMONE')>=0)return'GOMMONE';if(key.indexOf('SOGGIORNO')>=0||key.indexOf('ALLOGGIO')>=0||key.indexOf('HOTEL')>=0)return'HOTEL';if(key.indexOf('PASTI')>=0)return'PASTI';if(key.indexOf('VIAGGIO_TECNICO')>=0||key.indexOf('VIAGGIO TECNICO')>=0)return'VIAGGIO TECNICO';if(key.indexOf('CONV_')===0||key.indexOf('CONVOCAZ')>=0||key==='CHECK_CONFERME')return'CONVOCAZIONI';if(key.indexOf('CIRCOLO')>=0||key.indexOf('OSPITALITA')>=0||key.indexOf('RINGRAZIAMENTO')>=0)return'CIRCOLO';return String(category||'').trim()||'ALTRO';
}

function readMetaValue_(sheet,key){
  if(!sheet)return'';const values=sheet.getRange(1,1,Math.max(sheet.getLastRow(),1),2).getValues(),wanted=normalize_(key);for(let i=0;i<values.length;i++)if(normalize_(values[i][0])===wanted)return values[i][1];return'';
}
function writeMeta_(sheet,values){
  if(!sheet)return;const rows=sheet.getRange(1,1,Math.max(sheet.getLastRow(),1),2).getValues(),byKey={};rows.forEach((r,i)=>{if(r[0])byKey[normalize_(r[0])]=i+1;});Object.keys(values).forEach(key=>{const row=byKey[normalize_(key)];if(row)sheet.getRange(row,2).setValue(values[key]);else{sheet.appendRow([key,values[key]]);byKey[normalize_(key)]=sheet.getLastRow();}});
}
function validateEventSheetIdentity_(child,eventId){const meta=child&&child.getSheetByName(EVENT_SHEET.SHEETS.META);if(!meta)throw new Error('Foglio _META mancante.');const stored=String(readMetaValue_(meta,'EVENT_ID')||'').trim();if(stored&&stored!==String(eventId))throw new Error('La Scheda appartiene all evento '+stored+', non a '+eventId+'.');}
function buildEventSheetName_(event){return['Scheda evento',event[APP.CALENDAR_HEADERS.TYPE]||'',event[APP.CALENDAR_HEADERS.CLASS]||'',event[APP.CALENDAR_HEADERS.LOCATION]||'',formatDocumentDateRange_(event[APP.CALENDAR_HEADERS.START],event[APP.CALENDAR_HEADERS.END])].filter(Boolean).join(' - ');}
function buildEventSheetLabel_(event){return[event[APP.CALENDAR_HEADERS.TYPE]||'',event[APP.CALENDAR_HEADERS.CLASS]||'',event[APP.CALENDAR_HEADERS.LOCATION]||'',formatDocumentDateRange_(event[APP.CALENDAR_HEADERS.START],event[APP.CALENDAR_HEADERS.END])].filter(Boolean).join(' | ');}
function setEventSheetLink_(row,url){const sheet=sh_(APP.SHEETS.CALENDAR),map=headerMap_(sheet),col=map[APP.CALENDAR_HEADERS.EVENT_SHEET];if(!col)throw new Error('Colonna SCHEDA EVENTO non trovata.');sheet.getRange(row,col).setRichTextValue(SpreadsheetApp.newRichTextValue().setText('↗ APRI').setLinkUrl(url).build());}
function replaceCentralRowsForEvent_(sheet,eventId,eventColumn,newRows,width){const values=sheet.getDataRange().getValues(),target=[];for(let i=1;i<values.length;i++)if(String(values[i][eventColumn-1]||'')===String(eventId))target.push(i+1);const reused=Math.min(target.length,newRows.length);for(let i=0;i<reused;i++)sheet.getRange(target[i],1,1,width).setValues([newRows[i]]);for(let i=reused;i<target.length;i++)sheet.getRange(target[i],1,1,width).clearContent();if(newRows.length>reused){const rest=newRows.slice(reused);sheet.getRange(sheet.getLastRow()+1,1,rest.length,width).setValues(rest);}}
