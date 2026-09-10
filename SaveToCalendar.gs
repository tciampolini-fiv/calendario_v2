function onOpen() {
  SpreadsheetApp.getUi().createMenu('Scheda evento')
    .addItem('💾 Salva dati nel Calendario','eventSaveToCalendar')
    .addSeparator()
    .addItem('➕ Importa nuova spesa','eventImportExpenseV8')
    .addItem('🔄 Aggiorna cruscotto spese','eventRefreshExpenseDashboardV8_')
    .addSeparator()
    .addItem('📄 Genera documenti','eventGenerateDocuments')
    .addToUi();
  try { eventInitializeV7_(); } catch (err) { console.log('Inizializzazione Scheda evento: ' + (err.message || err)); }
}

function eventSaveToCalendar() {
  const ui=SpreadsheetApp.getUi();
  const child=SpreadsheetApp.getActive();
  const meta=eventMeta_();
  const eventId=String(meta.EVENT_ID||'').trim();
  const masterId=String(meta.MASTER_SPREADSHEET_ID||'').trim();
  if(!eventId) throw new Error('ID EVENTO mancante nel foglio _META.');
  if(!masterId) throw new Error('MASTER_SPREADSHEET_ID mancante nel foglio _META.');

  eventInitializeV7_();
  const master=SpreadsheetApp.openById(masterId);
  const counts={
    participants:eventSaveParticipants_(master,child,eventId),
    tasks:eventSaveTasks_(master,child,eventId),
    expenses:eventSaveExpensesV8_(master,child,eventId,meta)
  };
  eventWriteMetaValue_(child.getSheetByName(EVENT_APP.SHEETS.META),'LAST_SYNC',Utilities.formatDate(new Date(),Session.getScriptTimeZone()||'Europe/Rome','dd/MM/yyyy HH:mm'));
  SpreadsheetApp.flush();
  ui.alert('Calendario aggiornato','I dati della Scheda evento sono stati salvati nel Calendario.\n\nAttività: '+counts.tasks+'\nSpese/movimenti: '+counts.expenses+'\nPartecipanti: '+counts.participants+'\n\nLa Scheda evento non è stata ricaricata o sovrascritta.',ui.ButtonSet.OK);
  return counts;
}

function eventMasterSheet_(master,name) {
  const sheet=master.getSheetByName(name);
  if(!sheet) throw new Error('Foglio '+name+' non trovato nel Calendario master.');
  return sheet;
}
function eventNormalize_(value) { return String(value===null||value===undefined?'':value).trim().toUpperCase().replace(/\s+/g,' '); }

function eventReplaceRowsForEvent_(sheet,eventId,eventColumn,newRows,width) {
  const values=sheet.getDataRange().getValues();
  const targetRows=[];
  for(let i=1;i<values.length;i++) if(String(values[i][eventColumn-1]||'')===String(eventId)) targetRows.push(i+1);
  const reused=Math.min(targetRows.length,newRows.length);
  for(let i=0;i<reused;i++) sheet.getRange(targetRows[i],1,1,width).setValues([newRows[i]]);
  for(let i=reused;i<targetRows.length;i++) sheet.getRange(targetRows[i],1,1,width).clearContent();
  if(newRows.length>reused){const rest=newRows.slice(reused);sheet.getRange(sheet.getLastRow()+1,1,rest.length,width).setValues(rest);}
}

function eventSaveTasks_(master,child,eventId) {
  const local=child.getSheetByName(EVENT_APP.SHEETS.TASKS);
  if(!local) throw new Error('Foglio Attività non trovato nella Scheda evento.');
  eventAutoLinkConfirmationTasksV7_();
  eventRefreshTaskDependencyStatesV7_();
  const backend=eventMasterSheet_(master,'_CHECKLIST');
  const oldRows=backend.getDataRange().getValues();
  const oldByNo={}; let maxNo=0;
  for(let i=1;i<oldRows.length;i++){
    if(String(oldRows[i][1]||'')!==String(eventId)) continue;
    const n=Number(oldRows[i][13]||0); if(n>0)oldByNo[n]=oldRows[i]; if(n>maxNo)maxNo=n;
  }
  const lastRow=Math.min(Math.max(local.getLastRow(),2),500);
  const values=local.getRange(2,1,lastRow-1,6).getValues();
  const draft=[];
  values.forEach((r,i)=>{
    const description=String(r[0]||'').trim(); if(!description)return;
    let no=Number(r[2]||0); if(!(no>0)){no=++maxNo;local.getRange(i+2,3).setValue(no);} maxNo=Math.max(maxNo,no);
    draft.push({row:i+2,values:r,no:no});
  });
  const seen={}; draft.forEach(x=>{if(seen[x.no])throw new Error('Numero task duplicato: '+x.no+'.');seen[x.no]=true;});
  const idByNo={}; draft.forEach(x=>{const old=oldByNo[x.no];x.old=old?old.slice(0,16):new Array(16).fill('');x.id=old&&old[0]?String(old[0]):'TASK-'+Utilities.getUuid();idByNo[x.no]=x.id;});
  const now=new Date();
  const rows=draft.map((x,index)=>{
    const r=x.values,old=x.old,depNo=Number(r[3]||0);
    if(depNo&&!idByNo[depNo])throw new Error('La task n. '+x.no+' dipende dalla task n. '+depNo+', che non esiste.');
    if(depNo===x.no)throw new Error('La task n. '+x.no+' non può dipendere da se stessa.');
    let status=eventNormalize_(r[4])||'DA FARE'; if(!['DA FARE','IN ATTESA','FATTO'].includes(status))status='DA FARE';
    const completed=status==='FATTO'?(old[11] instanceof Date?old[11]:now):'';
    const description=String(r[0]||'').trim();
    const autoKey=old[9]||eventTaskAutoKey_(description);
    return [x.id,eventId,(index+1)*10,description,old[4]||eventTaskCategory_(description),r[5] instanceof Date?r[5]:'',status,old[7]||'',old[8]||(autoKey?'AUTO':'MANUALE'),autoKey,String(r[1]||'').trim(),completed,now,x.no,depNo?idByNo[depNo]:'',status!=='FATTO'&&depNo?'DIPENDENZA':''];
  });
  eventReplaceRowsForEvent_(backend,eventId,2,rows,16);
  return rows.length;
}

function eventTaskAutoKey_(description) {
  const d=eventNormalizeTextV7_(description);
  const direct={
    'GOMMONE':'GOMMONE','CONFERMA GOMMONE':'CONF_GOMMONE',
    'CONTATTARE HOTEL':'SOGGIORNO','CONFERMA HOTEL':'CONF_SOGGIORNO',
    'PASTI':'PASTI','CONFERMA PASTI':'CONF_PASTI','VIAGGIO TECNICO':'VIAGGIO_TECNICO',
    'CONVOCAZIONE ATLETI':'CONV_ATLETI','CONVOCAZIONE TECNICO':'CONV_TECNICO','CHECK CONFERMA PRESENZE':'CHECK_CONFERME',
    'OSPITALITA CIRCOLO':'OSPITALITA_CIRCOLO','CONFERMA OSPITALITA':'CONF_OSPITALITA','RINGRAZIAMENTO CIRCOLO':'RINGRAZIAMENTO_CIRCOLO',
    'PAGAMENTO AFOR':'PAGAMENTO_AFOR','PAGAMENTO FATTURA':'PAGAMENTO_FATTURA','PAGAMENTO':'PAGAMENTO',
    'INVIARE CONTABILE':'CONTABILE_INVIA','RICHIEDERE CONTABILE':'CONTABILE_RICHIEDI','RICHIEDERE FATTURA':'RICHIEDERE_FATTURA'
  };
  return direct[d]||'';
}
function eventTaskCategory_(description) {
  const d=eventNormalize_(description);
  if(/CONVOC|PRESENZ/.test(d))return'CONVOCAZIONI';
  if(/HOTEL|ALLOGG|PAST|OSPITAL|GOMMONE|TRASPORT/.test(d))return'LOGISTICA';
  if(/PAG|FATTUR|RIF|AFOR|RIMBORS|CONTABILE|SALDO/.test(d))return'AMMINISTRAZIONE';
  if(/CIRCOLO|ZONA|COMUNIC/.test(d))return'ORGANIZZAZIONE';
  return'ALTRO';
}

function eventSaveParticipants_(master,child,eventId) {
  const local=child.getSheetByName(EVENT_APP.SHEETS.PARTICIPANTS);
  if(!local)throw new Error('Foglio Partecipanti non trovato nella Scheda evento.');
  const backend=eventMasterSheet_(master,'_PARTECIPANTI');
  const oldRows=backend.getDataRange().getValues(),oldById={};
  for(let i=1;i<oldRows.length;i++)if(String(oldRows[i][1]||'')===String(eventId)&&oldRows[i][0])oldById[String(oldRows[i][0])]=oldRows[i];
  const now=new Date(),rows=[],convTech=[],aggTech=[];
  const conv=local.getRange(3,1,15,19).getValues(),agg=local.getRange(20,1,20,19).getValues();
  conv.forEach(r=>{const parsed=eventParticipantRowV7_(r,'CONVOCATO',eventId,oldById,now);if(!parsed){convTech.push(['','','','','','']);return;}rows.push(parsed.backend);convTech.push(parsed.tech);});
  agg.forEach(r=>{const parsed=eventParticipantRowV7_(r,'AGGREGATO',eventId,oldById,now);if(!parsed){aggTech.push(['','','','','','']);return;}rows.push(parsed.backend);aggTech.push(parsed.tech);});
  eventReplaceRowsForEvent_(backend,eventId,2,rows,17);
  local.getRange(3,14,convTech.length,6).setValues(convTech); local.getRange(20,14,aggTech.length,6).setValues(aggTech);
  return rows.length;
}
function eventParticipantRowV7_(r,type,eventId,oldById,now) {
  const name=String(r[0]||'').trim(),surname=String(r[1]||'').trim(); if(!name&&!surname)return null;
  let id=String(r[13]||'').trim(); const old=id&&oldById[id]?oldById[id]:null; if(!id)id='PAR-'+Utilities.getUuid();
  const personId=String(r[14]||'').trim()||(old?String(old[2]||''):'');
  const role=eventNormalize_(r[7])||'ATLETA'; let status=eventNormalize_(r[8]);
  if(type==='CONVOCATO'&&!['DA FARE','MANDATA CONVOCAZIONE','CONFERMATO','ASSENTE'].includes(status))status=role==='TECNICO'?'CONFERMATO':'DA FARE';
  if(type==='AGGREGATO'&&!['DA AUTORIZZARE','AUTORIZZATO','NON AUTORIZZATO'].includes(status))status='DA AUTORIZZARE';
  const provenance=String(r[16]||'').trim()||(old?String(old[9]||''):'')||'SCHEDA EVENTO';
  const sourceFile=String(r[17]||'').trim()||(old?String(old[10]||''):'');
  const createdAt=r[18] instanceof Date?r[18]:(old&&old[12] instanceof Date?old[12]:now);
  const maxRefund=type==='CONVOCATO'&&r[9]!==''&&role!=='TECNICO'?Number(r[9]||0):'';
  const passed=type==='CONVOCATO'&&r[10]!==''&&role!=='TECNICO'?Number(r[10]||0):'';
  const passedDate=passed>0?(r[11] instanceof Date?r[11]:(old&&old[16] instanceof Date?old[16]:now)):'';
  const notes=String(type==='CONVOCATO'?r[12]:r[9]||'').trim();
  const backend=[id,eventId,personId,name,surname,String(r[2]||'').trim(),role,String(r[6]||'').trim(),maxRefund,provenance,sourceFile,notes,createdAt,type,status,passed,passedDate];
  return {backend:backend,tech:[id,personId,eventId,provenance,sourceFile,createdAt]};
}

function eventExpenseCategoryV7_(category,description) {
  let c=eventNormalize_(category),d=eventNormalize_(description);
  if(c==='VITTO / ALLOGGIO'){if(/PASTI|VITTO|RISTORANTE|PRANZO|CENA/.test(d))return'VITTO';return'ALLOGGIO';}
  if(c==='VIAGGIO')c='VIAGGI'; return c||'ALTRO';
}
function eventResolveCeb_(master,meta,category,current) {
  if(current)return current;
  const type=eventNormalize_(meta.EVENT_TYPE||''); let profile=type;
  const typeCfg=master.getSheetByName('_CONFIG_TIPI_EVENTO');
  if(typeCfg){const rows=typeCfg.getDataRange().getValues();for(let i=1;i<rows.length;i++)if(eventNormalize_(rows[i][0])===type){profile=eventNormalize_(rows[i][4])||type;break;}}
  if(profile==='FOIL ACADEMY')return'CEB.033';
  const cfg=master.getSheetByName('_CONFIG_CEB'); if(!cfg)return eventNormalize_(category)==='RIMBORSO'?'CEB.002':'';
  const rows=cfg.getDataRange().getValues(),options=[];
  for(let i=1;i<rows.length;i++){
    const active=rows[i][4]===true||eventNormalize_(rows[i][4])==='SI'; if(!active)continue;
    const allowed=String(rows[i][2]||'').split(';').map(eventNormalize_).filter(Boolean);
    if(!(allowed.includes('TUTTI')||allowed.includes(type)||allowed.includes(profile)))continue;
    if(eventNormalize_(rows[i][3]||'SPESA')==='RIMBORSO')continue;
    options.push({code:String(rows[i][0]||''),cats:String(rows[i][6]||'').split(';').map(eventNormalize_).filter(Boolean)});
  }
  const cat=eventNormalize_(category),exact=options.find(o=>o.cats.includes(cat)); if(exact)return exact.code;
  const generic=options.find(o=>o.cats.includes('TUTTI')); if(generic)return generic.code;
  const travel=options.find(o=>o.code==='CEB.001'); return travel?travel.code:(options[0]?options[0].code:'');
}
function eventWriteMetaValue_(sheet,key,value) {
  if(!sheet)return; const rows=sheet.getRange(1,1,Math.max(sheet.getLastRow(),1),2).getValues(),target=eventNormalize_(key);
  for(let i=0;i<rows.length;i++)if(eventNormalize_(rows[i][0])===target){sheet.getRange(i+1,2).setValue(value);return;}
  sheet.getRange(sheet.getLastRow()+1,1,1,2).setValues([[key,value]]);
}
