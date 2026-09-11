function onOpen() {
  SpreadsheetApp.getUi().createMenu('Scheda evento')
    .addItem('↻ Sincronizza ora col Calendario','eventSaveToCalendar')
    .addSeparator()
    .addItem('➕ Nuovo obiettivo','eventAddObjectiveV13')
    .addItem('➕ Aggiungi attività all’obiettivo','eventAddTaskToObjectiveV13')
    .addSeparator()
    .addItem('➕ Importa nuova spesa','eventImportExpenseV8')
    .addItem('🔄 Aggiorna cruscotto spese','eventRefreshExpenseDashboardV8_')
    .addSeparator()
    .addItem('📄 Crea documento dal foglio DOCUMENTI','eventCreateDocumentFromDocumentsSheetV12')
    .addItem('📂 Aggiorna elenco documenti','eventRefreshDocumentListV12')
    .addToUi();
  try { eventInitializeV7_(); eventFormatActivityRowsV13_(); } catch (err) { console.log('Inizializzazione Scheda evento: ' + (err.message || err)); }
  try { eventEnsureDocumentsSheetV12_(); eventRefreshDocumentListV12_(false); } catch (err) { console.log('Documenti Scheda evento: ' + (err.message || err)); }
}

function eventSaveToCalendar() {
  const child=SpreadsheetApp.getActive(),meta=eventMeta_();
  const eventId=String(meta.EVENT_ID||'').trim(),masterId=String(meta.MASTER_SPREADSHEET_ID||'').trim();
  if(!eventId)throw new Error('ID EVENTO mancante nel foglio _META.');
  if(!masterId)throw new Error('MASTER_SPREADSHEET_ID mancante nel foglio _META.');
  const master=SpreadsheetApp.openById(masterId);
  const counts={participants:eventSaveParticipants_(master,child,eventId),tasks:eventSaveActivitiesV13_(master,child,eventId),expenses:eventSaveExpensesV8_(master,child,eventId,meta)};
  eventRefreshMasterActivitySummaryV13_(master,eventId,child);
  eventWriteMetaValue_(child.getSheetByName(EVENT_APP.SHEETS.META),'LAST_SYNC',Utilities.formatDate(new Date(),Session.getScriptTimeZone()||'Europe/Rome','dd/MM/yyyy HH:mm'));
  SpreadsheetApp.flush();
  child.toast('Attività: '+counts.tasks+' | Spese: '+counts.expenses+' | Partecipanti: '+counts.participants,'Calendario aggiornato',6);
  return counts;
}

function eventMasterSheet_(master,name){const sheet=master.getSheetByName(name);if(!sheet)throw new Error('Foglio '+name+' non trovato nel Calendario master.');return sheet;}
function eventNormalize_(value){return String(value===null||value===undefined?'':value).trim().toUpperCase().replace(/\s+/g,' ');}
function eventReplaceRowsForEvent_(sheet,eventId,eventColumn,newRows,width){
  const values=sheet.getDataRange().getValues(),targetRows=[];
  for(let i=1;i<values.length;i++)if(String(values[i][eventColumn-1]||'')===String(eventId))targetRows.push(i+1);
  const reused=Math.min(targetRows.length,newRows.length);
  for(let i=0;i<reused;i++)sheet.getRange(targetRows[i],1,1,width).setValues([newRows[i]]);
  for(let i=reused;i<targetRows.length;i++)sheet.getRange(targetRows[i],1,1,width).clearContent();
  if(newRows.length>reused){const rest=newRows.slice(reused);sheet.getRange(sheet.getLastRow()+1,1,rest.length,width).setValues(rest);}
}

function eventEnsureMasterObjectivesV13_(master){
  let sh=master.getSheetByName('_OBIETTIVI');if(!sh){sh=master.insertSheet('_OBIETTIVI');sh.getRange(1,1,1,11).setValues([['ID OBIETTIVO','ID EVENTO','NOME','COLORE','SCADENZA OBIETTIVO','ORDINE','ORIGINE','TEMPLATE KEY','NOTE','DATA INSERIMENTO','ULTIMO AGGIORNAMENTO']]);sh.setFrozenRows(1);}return sh;
}
function eventSaveActivitiesV13_(master,child,eventId){
  const local=child.getSheetByName(EVENT_APP.SHEETS.TASKS);if(!local)throw new Error('Foglio Attività non trovato nella Scheda evento.');eventRefreshActivitiesV13_();
  const objectiveBackend=eventEnsureMasterObjectivesV13_(master),taskBackend=eventMasterSheet_(master,'_CHECKLIST');
  if(taskBackend.getLastColumn()<20)taskBackend.insertColumnsAfter(taskBackend.getLastColumn(),20-taskBackend.getLastColumn());
  taskBackend.getRange(1,17,1,4).setValues([['ID OBIETTIVO','ORDINE STEP','SCADENZA AUTOMATICA','ID TASK PRECEDENTE']]);
  const oldObjRows=objectiveBackend.getDataRange().getValues(),oldObjById={};oldObjRows.slice(1).forEach(r=>{if(String(r[1])===String(eventId)&&r[0])oldObjById[String(r[0])]=r;});
  const oldTaskRows=taskBackend.getDataRange().getValues(),oldTaskById={};oldTaskRows.slice(1).forEach(r=>{if(String(r[1])===String(eventId)&&r[0])oldTaskById[String(r[0])]=r;});
  const last=Math.min(EVENT_APP.ACTIVITY.MAX,Math.max(local.getLastRow(),2)),rows=local.getRange(2,1,last-1,EVENT_APP.ACTIVITY.COLS).getValues(),now=new Date(),objectives=[],tasks=[],objectiveName={};
  rows.forEach((r,i)=>{
    const type=eventNormalize_(r[7]);if(type!=='OBIETTIVO')return;let oid=String(r[8]||'').trim();if(!oid){oid='OBJ-'+Utilities.getUuid();local.getRange(i+2,9).setValue(oid);}const name=String(r[0]||'').trim();if(!name)return;const old=oldObjById[oid]||[];objectives.push([oid,eventId,name,String(r[12]||old[3]||'#D9EAF7'),r[1] instanceof Date?r[1]:'',Number(r[10]||old[5]||((objectives.length+1)*10)),old[6]||'SCHEDA EVENTO',old[7]||'',old[8]||'',old[9]||now,now]);objectiveName[oid]=name;
  });
  rows.forEach((r,i)=>{
    if(eventNormalize_(r[7])!=='TASK')return;const task=String(r[2]||'').trim(),oid=String(r[8]||'').trim();if(!task||!oid)return;let tid=String(r[9]||'').trim();if(!tid){tid='TASK-'+Utilities.getUuid();local.getRange(i+2,10).setValue(tid);}const old=oldTaskById[tid]||[],checked=Boolean(r[4]),status=checked?'FATTO':(eventNormalize_(r[5])||'IN ATTESA'),step=Number(r[11]||10),objectiveOrder=Number(r[10]||10),previous=String(r[14]||'').trim(),completed=checked?(r[15] instanceof Date?r[15]:(old[11] instanceof Date?old[11]:now)):'';tasks.push([tid,eventId,objectiveOrder*100+step,task,old[4]||eventTaskCategory_(task),r[3] instanceof Date?r[3]:'',status,old[7]||'',old[8]||'SCHEDA EVENTO',old[9]||eventTaskAutoKey_(task),String(r[6]||'').trim(),completed,now,old[13]||tasks.length+1,previous,previous&&status!=='FATTO'?'DIPENDENZA':'',oid,step,r[13]===true,previous]);
  });
  eventReplaceRowsForEvent_(objectiveBackend,eventId,2,objectives,11);eventReplaceRowsForEvent_(taskBackend,eventId,2,tasks,20);return tasks.length;
}
function eventRefreshMasterActivitySummaryV13_(master,eventId,child){
  const cal=master.getSheetByName('Calendario');if(!cal)return;const headers=cal.getRange(1,1,1,cal.getLastColumn()).getDisplayValues()[0],idCol=headers.indexOf('ID EVENTO')+1,advCol=headers.indexOf('AVANZAMENTO')+1,nowCol=headers.indexOf('DA FARE ORA')+1;if(idCol<1)return;const ids=cal.getRange(2,idCol,Math.max(1,cal.getLastRow()-1),1).getDisplayValues().flat(),idx=ids.findIndex(x=>String(x)===String(eventId));if(idx<0)return;const row=idx+2,sh=child.getSheetByName(EVENT_APP.SHEETS.TASKS),last=Math.min(EVENT_APP.ACTIVITY.MAX,Math.max(sh.getLastRow(),2)),vals=sh.getRange(2,1,last-1,EVENT_APP.ACTIVITY.COLS).getValues(),objectives=[];
  vals.forEach(r=>{if(eventNormalize_(r[7])==='OBIETTIVO'&&r[8])objectives.push({id:String(r[8]),name:String(r[0]||''),status:eventNormalize_(r[5]),order:Number(r[10]||0),tasks:[]});});vals.forEach(r=>{if(eventNormalize_(r[7])==='TASK'&&r[8]){const o=objectives.find(x=>x.id===String(r[8]));if(o)o.tasks.push({checked:Boolean(r[4]),status:eventNormalize_(r[5]),name:String(r[2]||''),step:Number(r[11]||0)});}});objectives.sort((a,b)=>a.order-b.order);objectives.forEach(o=>o.tasks.sort((a,b)=>a.step-b.step));const progress=objectives.map(o=>(o.status==='COMPLETATO'?'✓':o.status==='IN RITARDO'?'⚠':'•')+' '+o.name+'  '+o.tasks.map(t=>t.checked?'☑':'☐').join('')).join('\n'),actions=[];objectives.forEach(o=>o.tasks.filter(t=>t.status==='DA FARE').forEach(t=>actions.push(o.name+' — '+t.name)));if(advCol>0)cal.getRange(row,advCol).setValue(progress).setWrap(true);if(nowCol>0)cal.getRange(row,nowCol).setValue(actions.join('\n')).setWrap(true);
}

function eventTaskAutoKey_(description) {
  const d=eventNormalizeTextV7_(description);
  const direct={'RICHIESTA GOMMONE':'GOMMONE','CHIEDERE DISPONIBILITA':'GOMMONE','CONFERMA GOMMONE':'CONF_GOMMONE','RICEVERE CONFERMA':'CONF_GOMMONE','RICHIESTA SOGGIORNO':'SOGGIORNO','CHIEDERE DISPONIBILITA HOTEL':'SOGGIORNO','CONFERMA SOGGIORNO':'CONF_SOGGIORNO','RICEVERE RISPOSTA HOTEL':'CONF_SOGGIORNO','SOLUZIONE PASTI':'PASTI','DEFINIRE SOLUZIONE PASTI':'PASTI','CONFERMA PASTI':'CONF_PASTI','VIAGGIO TECNICO':'VIAGGIO_TECNICO','DEFINIRE VIAGGIO':'VIAGGIO_TECNICO','CONVOCAZIONE ATLETI':'CONV_ATLETI','INVIARE CONVOCAZIONE':'CONV_ATLETI','CONVOCAZIONE TECNICO':'CONV_TECNICO','INVIARE CONVOCAZIONE TECNICO':'CONV_TECNICO','CONFERMA PRESENZE TUTTI ATLETI':'CHECK_CONFERME','OSPITALITA CIRCOLO':'OSPITALITA_CIRCOLO','CHIEDERE OSPITALITA':'OSPITALITA_CIRCOLO','RINGRAZIAMENTO CIRCOLO':'RINGRAZIAMENTO_CIRCOLO','INVIARE CONTABILE AFOR':'CONTABILE_AFOR','INVIARE CONTABILE':'CONTABILE_AFOR','PAGARE AFOR':'PAGARE_AFOR','RICHIEDERE FATTURA':'RICHIEDERE_FATTURA'};return direct[d]||'';
}
function eventTaskCategory_(description){const d=eventNormalize_(description);if(/CONVOC|PRESENZ/.test(d))return'CONVOCAZIONI';if(/SOGGIORNO|HOTEL|ALLOGG|PAST|OSPITAL|GOMMONE|TRASPORT|VIAGG|PRENOT/.test(d))return'LOGISTICA';if(/PAG|FATTUR|RIF|AFOR|RIMBORS|CONTABILE|SALDO/.test(d))return'AMMINISTRAZIONE';if(/CIRCOLO|ZONA|COMUNIC/.test(d))return'ORGANIZZAZIONE';return'ALTRO';}

function eventSaveParticipants_(master,child,eventId){
  const local=child.getSheetByName(EVENT_APP.SHEETS.PARTICIPANTS);if(!local)throw new Error('Foglio Partecipanti non trovato nella Scheda evento.');
  const backend=eventMasterSheet_(master,'_PARTECIPANTI'),oldRows=backend.getDataRange().getValues(),oldById={};for(let i=1;i<oldRows.length;i++)if(String(oldRows[i][1]||'')===String(eventId)&&oldRows[i][0])oldById[String(oldRows[i][0])]=oldRows[i];
  const now=new Date(),rows=[],convTech=[],aggTech=[],conv=local.getRange(3,1,20,19).getValues(),agg=local.getRange(25,1,20,19).getValues();
  conv.forEach(r=>{const parsed=eventParticipantRowV7_(r,'CONVOCATO',eventId,oldById,now);if(!parsed){convTech.push(['','','','','','']);return;}rows.push(parsed.backend);convTech.push(parsed.tech);});
  agg.forEach(r=>{const parsed=eventParticipantRowV7_(r,'AGGREGATO',eventId,oldById,now);if(!parsed){aggTech.push(['','','','','','']);return;}rows.push(parsed.backend);aggTech.push(parsed.tech);});
  eventReplaceRowsForEvent_(backend,eventId,2,rows,17);local.getRange(3,14,convTech.length,6).setValues(convTech);local.getRange(25,14,aggTech.length,6).setValues(aggTech);return rows.length;
}
function eventParticipantRowV7_(r,type,eventId,oldById,now){
  const name=String(r[0]||'').trim(),surname=String(r[1]||'').trim();if(!name&&!surname)return null;let id=String(r[13]||'').trim();const old=id&&oldById[id]?oldById[id]:null;if(!id)id='PAR-'+Utilities.getUuid();
  const personId=String(r[14]||'').trim()||(old?String(old[2]||''):''),role=eventNormalize_(r[7])||'ATLETA';let status=eventNormalize_(r[8]);
  if(type==='CONVOCATO'&&!['DA FARE','MANDATA CONVOCAZIONE','CONFERMATO','ASSENTE'].includes(status))status=role==='TECNICO'?'CONFERMATO':'DA FARE';if(type==='AGGREGATO'&&!['DA AUTORIZZARE','AUTORIZZATO','NON AUTORIZZATO'].includes(status))status='DA AUTORIZZARE';
  const provenance=String(r[16]||'').trim()||(old?String(old[9]||''):'')||'SCHEDA EVENTO',sourceFile=String(r[17]||'').trim()||(old?String(old[10]||''):''),createdAt=r[18] instanceof Date?r[18]:(old&&old[12] instanceof Date?old[12]:now),maxRefund=type==='CONVOCATO'&&r[9]!==''&&role!=='TECNICO'?Number(r[9]||0):'',passed=type==='CONVOCATO'&&r[10]!==''&&role!=='TECNICO'?Number(r[10]||0):'',passedDate=passed>0?(r[11] instanceof Date?r[11]:(old&&old[16] instanceof Date?old[16]:now)):'',notes=String(type==='CONVOCATO'?r[12]:r[9]||'').trim();
  return{backend:[id,eventId,personId,name,surname,String(r[2]||'').trim(),role,String(r[6]||'').trim(),maxRefund,provenance,sourceFile,notes,createdAt,type,status,passed,passedDate],tech:[id,personId,eventId,provenance,sourceFile,createdAt]};
}
function eventExpenseCategoryV7_(category,description){let c=eventNormalize_(category),d=eventNormalize_(description);if(c==='VITTO / ALLOGGIO'){if(/PASTI|VITTO|RISTORANTE|PRANZO|CENA/.test(d))return'VITTO';return'ALLOGGIO';}if(c==='VIAGGIO')c='VIAGGI';return c||'ALTRO';}
function eventResolveCeb_(master,meta,category,current){
  if(current)return current;const type=eventNormalize_(meta.EVENT_TYPE||'');let profile=type;const typeCfg=master.getSheetByName('_CONFIG_TIPI_EVENTO');if(typeCfg){const rows=typeCfg.getDataRange().getValues();for(let i=1;i<rows.length;i++)if(eventNormalize_(rows[i][0])===type){profile=eventNormalize_(rows[i][4])||type;break;}}
  if(profile==='FOIL ACADEMY')return'CEB.033';const cfg=master.getSheetByName('_CONFIG_CEB');if(!cfg)return eventNormalize_(category)==='RIMBORSO'?'CEB.002':'';const rows=cfg.getDataRange().getValues(),options=[];
  for(let i=1;i<rows.length;i++){const active=rows[i][4]===true||eventNormalize_(rows[i][4])==='SI';if(!active)continue;const allowed=String(rows[i][2]||'').split(';').map(eventNormalize_).filter(Boolean);if(!(allowed.includes('TUTTI')||allowed.includes(type)||allowed.includes(profile)))continue;if(eventNormalize_(rows[i][3]||'SPESA')==='RIMBORSO')continue;options.push({code:String(rows[i][0]||''),cats:String(rows[i][6]||'').split(';').map(eventNormalize_).filter(Boolean)});}
  const cat=eventNormalize_(category),exact=options.find(o=>o.cats.includes(cat));if(exact)return exact.code;const generic=options.find(o=>o.cats.includes('TUTTI'));if(generic)return generic.code;const travel=options.find(o=>o.code==='CEB.001');return travel?travel.code:(options[0]?options[0].code:'');
}
function eventWriteMetaValue_(sheet,key,value){if(!sheet)return;const rows=sheet.getRange(1,1,Math.max(sheet.getLastRow(),1),2).getValues(),target=eventNormalize_(key);for(let i=0;i<rows.length;i++)if(eventNormalize_(rows[i][0])===target){sheet.getRange(i+1,2).setValue(value);return;}sheet.getRange(sheet.getLastRow()+1,1,1,2).setValues([[key,value]]);}
