const EVENT_APP = Object.freeze({
  SHEETS:{TASKS:'Attività',EXPENSES:'Spese',PARTICIPANTS:'Partecipanti',DOCUMENTS:'DOCUMENTI',META:'_META'},
  MASTER_DOCUMENTS:'_DOCUMENTI',
  PARTICIPANTS:{CONV_START:3,CONV_COUNT:20,AGG_START:25,AGG_COUNT:20},
  ACTIVITY:{HEADER:1,START:2,MAX:500,COLS:19,TYPE:8,OBJECTIVE_ID:9,TASK_ID:10,OBJECTIVE_ORDER:11,STEP_ORDER:12,COLOR:13,AUTO_DUE:14,PREVIOUS_ID:15,COMPLETED_AT:16,PATH:17,DUE_MODE:18,OFFSET:19},
  TEMPLATE_SPECS:[
    {key:'CONV_ATLETI',label:'Convocazione atleti',id:'1SpOMrpDwe8aTW9QAyby865WufnD40AsPx6t8zm6Yw3E'},
    {key:'CONV_TECNICO',label:'Convocazione tecnico',id:'1Xb61H5TQ0avz8n5_THn-BSd1Xd4YBtmOD3Lu2uwUuJw'},
    {key:'GOMMONE',label:'Richiesta gommone alla Zona',id:'1uvK4J8WrWekpMrkPTkpUvsdYTggBhwzCDJaJmyIhUko'},
    {key:'OSPITALITA',label:'Richiesta ospitalità circolo',id:'1NuheJqcwtLxp8muTdbnpcZc22Z8HJgA-jbdt_Sq6ruk'},
    {key:'RINGRAZIAMENTO',label:'Ringraziamento circolo',id:'11Fqn8U3Hs8MgJPH0tQTKDEozmWDftPtmELoLjgMFg7c'}
  ]
});

function onEdit(e){
  if(!e||!e.range)return;
  const sheet=e.range.getSheet(),name=sheet.getName();
  if(name===EVENT_APP.SHEETS.TASKS){eventHandleActivitiesEditV14_(e);return;}
  if(name===EVENT_APP.SHEETS.EXPENSES){eventHandleExpenseEditV8_(e);return;}
  if(name===EVENT_APP.SHEETS.PARTICIPANTS){const row=e.range.getRow();if((row>=3&&row<=22)||(row>=25&&row<=44))eventRefreshExpenseDashboardV8_();}
}

function onSelectionChange(e){
  if(!e||!e.range)return;const sh=e.range.getSheet();if(sh.getName()!==EVENT_APP.SHEETS.TASKS||e.range.getNumRows()!==1||e.range.getNumColumns()!==1)return;
  const row=e.range.getRow(),col=e.range.getColumn();if(row<2)return;const type=eventNormalizeTextV7_(sh.getRange(row,EVENT_APP.ACTIVITY.TYPE).getDisplayValue());
  if(type==='AGGIUNGI_OBIETTIVO'&&col===1){eventInsertObjectiveV14_(sh,row);return;}
  if(type==='OBIETTIVO'&&col===3&&String(sh.getRange(row,3).getDisplayValue()).indexOf('＋')>=0)eventInsertTaskV14_(sh,row);
}

function eventInitializeV7_(){
  const meta=eventMeta_();if(!String(meta.EVENT_ID||'').trim())return false;
  eventApplyCommitmentV8_(meta);eventPopulateTechniciansV7_(meta);eventRefreshActivitiesV14_();eventFormatActivityRowsV13_();eventSyncExpenseTasksV8_();eventRefreshExpenseDashboardV8_();return true;
}

function eventActivityPaletteV14_(){return['#D9EAF7','#FCE8B2','#EADCF8','#D9EAD3','#F4CCCC','#D0E0E3','#FCE5CD','#D9D2E9','#CFE2F3','#E2F0D9'];}
function eventAddDaysV14_(d,n){if(!(d instanceof Date))return'';const x=new Date(d);x.setDate(x.getDate()+Number(n||0));return x;}
function eventDayV14_(d){if(!(d instanceof Date))return null;const x=new Date(d);x.setHours(0,0,0,0);return x;}
function eventMetaDateV14_(key){return eventParseIsoDate_(eventMeta_()[key]||'');}
function eventDefaultObjectiveDueV14_(){return eventAddDaysV14_(eventMetaDateV14_('EVENT_START'),-5);}
function eventActivityRowsV14_(sh){const last=Math.min(EVENT_APP.ACTIVITY.MAX,Math.max(sh.getLastRow(),2));return sh.getRange(2,1,last-1,EVENT_APP.ACTIVITY.COLS).getValues().map((v,i)=>({row:i+2,v:v}));}
function eventObjectiveTasksV14_(sh,oid){return eventActivityRowsV14_(sh).filter(x=>eventNormalizeTextV7_(x.v[7])==='TASK'&&String(x.v[8]||'')===String(oid)).sort((a,b)=>Number(a.v[11]||0)-Number(b.v[11]||0));}
function eventFindObjectiveHeaderRowV14_(sh,row){for(let r=row;r>=2;r--){if(eventNormalizeTextV7_(sh.getRange(r,EVENT_APP.ACTIVITY.TYPE).getDisplayValue())==='OBIETTIVO')return r;}return 0;}
function eventTaskMapV14_(sh,oid){const out={};eventObjectiveTasksV14_(sh,oid).forEach(x=>{if(x.v[9])out[String(x.v[9])]=x;});return out;}

function eventEnsureObjectiveRowV14_(sh,row){
  const name=String(sh.getRange(row,1).getDisplayValue()||'').trim();if(!name||name.indexOf('＋ AGGIUNGI OBIETTIVO')===0)return;
  sh.getRange(row,EVENT_APP.ACTIVITY.TYPE).setValue('OBIETTIVO');if(!sh.getRange(row,EVENT_APP.ACTIVITY.OBJECTIVE_ID).getValue())sh.getRange(row,EVENT_APP.ACTIVITY.OBJECTIVE_ID).setValue('OBJ-'+Utilities.getUuid());
  if(!(sh.getRange(row,2).getValue() instanceof Date)){const due=eventDefaultObjectiveDueV14_();if(due)sh.getRange(row,2).setValue(due).setNumberFormat('dd/MM/yyyy');}
  if(!(Number(sh.getRange(row,EVENT_APP.ACTIVITY.OBJECTIVE_ORDER).getValue())>0)){const vals=eventActivityRowsV14_(sh).filter(x=>eventNormalizeTextV7_(x.v[7])==='OBIETTIVO').map(x=>Number(x.v[10]||0)).filter(n=>n>0);sh.getRange(row,EVENT_APP.ACTIVITY.OBJECTIVE_ORDER).setValue(vals.length?Math.max.apply(null,vals)+10:10);}
  let color=String(sh.getRange(row,1).getBackground()||'').toLowerCase();if(!color||color==='#ffffff'){const p=eventActivityPaletteV14_(),n=eventActivityRowsV14_(sh).filter(x=>eventNormalizeTextV7_(x.v[7])==='OBIETTIVO').length;color=p[Math.max(0,n-1)%p.length];sh.getRange(row,1,1,7).setBackground(color);}sh.getRange(row,EVENT_APP.ACTIVITY.COLOR).setValue(color);sh.getRange(row,3).setValue('＋').setHorizontalAlignment('center');sh.getRange(row,1,1,7).setFontWeight('bold');
}
function eventEnsureTaskRowV14_(sh,row){
  const task=String(sh.getRange(row,3).getDisplayValue()||'').trim();if(!task)return;const header=eventFindObjectiveHeaderRowV14_(sh,row);if(!header)throw new Error('Aggiungi prima un obiettivo.');eventEnsureObjectiveRowV14_(sh,header);
  const oid=String(sh.getRange(header,EVENT_APP.ACTIVITY.OBJECTIVE_ID).getValue()||'');sh.getRange(row,EVENT_APP.ACTIVITY.TYPE).setValue('TASK');sh.getRange(row,EVENT_APP.ACTIVITY.OBJECTIVE_ID).setValue(oid);if(!sh.getRange(row,EVENT_APP.ACTIVITY.TASK_ID).getValue())sh.getRange(row,EVENT_APP.ACTIVITY.TASK_ID).setValue('TASK-'+Utilities.getUuid());
  sh.getRange(row,EVENT_APP.ACTIVITY.OBJECTIVE_ORDER).setValue(sh.getRange(header,EVENT_APP.ACTIVITY.OBJECTIVE_ORDER).getValue());if(!sh.getRange(row,EVENT_APP.ACTIVITY.PATH).getValue())sh.getRange(row,EVENT_APP.ACTIVITY.PATH).setValue('MANUALE');
  const tasks=eventObjectiveTasksV14_(sh,oid).filter(x=>x.row!==row),path=String(sh.getRange(row,EVENT_APP.ACTIVITY.PATH).getValue()||'MANUALE'),same=tasks.filter(x=>String(x.v[16]||'MANUALE')===path);let step=Number(sh.getRange(row,EVENT_APP.ACTIVITY.STEP_ORDER).getValue()||0);if(!(step>0)){step=tasks.length?Math.max.apply(null,tasks.map(x=>Number(x.v[11]||0)))+10:10;sh.getRange(row,EVENT_APP.ACTIVITY.STEP_ORDER).setValue(step);}const prev=same.length?same[same.length-1]:null;sh.getRange(row,EVENT_APP.ACTIVITY.PREVIOUS_ID).setValue(prev?prev.v[9]:'');if(prev&&!sh.getRange(row,EVENT_APP.ACTIVITY.DUE_MODE).getValue()){sh.getRange(row,EVENT_APP.ACTIVITY.DUE_MODE).setValue('PRECEDENTE');sh.getRange(row,EVENT_APP.ACTIVITY.OFFSET).setValue(2);}else if(!prev&&!sh.getRange(row,EVENT_APP.ACTIVITY.DUE_MODE).getValue())sh.getRange(row,EVENT_APP.ACTIVITY.DUE_MODE).setValue('MANUAL');
  sh.getRange(row,5).insertCheckboxes();if(sh.getRange(row,5).isBlank())sh.getRange(row,5).setValue(false);
}

function eventHandleActivitiesEditV14_(e){
  const sh=e.range.getSheet(),row=e.range.getRow(),col=e.range.getColumn();if(row<2||row>EVENT_APP.ACTIVITY.MAX)return;const type=eventNormalizeTextV7_(sh.getRange(row,EVENT_APP.ACTIVITY.TYPE).getDisplayValue());
  if(type==='OBIETTIVO'){if(col===1||col===2){eventEnsureObjectiveRowV14_(sh,row);eventRefreshActivitiesV14_();}return;}
  if(type==='TASK'){
    if(col===3)eventEnsureTaskRowV14_(sh,row);
    if(col===4){sh.getRange(row,EVENT_APP.ACTIVITY.AUTO_DUE).setValue(false);sh.getRange(row,EVENT_APP.ACTIVITY.DUE_MODE).setValue('MANUAL');}
    if(col===5)eventToggleTaskV14_(sh,row,e.value===true||String(e.value).toUpperCase()==='TRUE');
    eventRefreshActivitiesV14_();
  }
}

function eventToggleTaskV14_(sh,row,checked){
  eventEnsureTaskRowV14_(sh,row);const oid=String(sh.getRange(row,EVENT_APP.ACTIVITY.OBJECTIVE_ID).getValue()||''),tid=String(sh.getRange(row,EVENT_APP.ACTIVITY.TASK_ID).getValue()||''),map=eventTaskMapV14_(sh,oid),current=map[tid];if(!current)return;const prevId=String(current.v[14]||''),prev=prevId?map[prevId]:null;
  if(checked&&prev&&!Boolean(prev.v[4])){sh.getRange(row,5).setValue(false);SpreadsheetApp.getUi().alert('Attività ancora bloccata','Completa prima l’attività precedente nello stesso percorso.',SpreadsheetApp.getUi().ButtonSet.OK);return;}
  const descendants=[];let cursor=tid;while(cursor){const next=Object.values(map).find(x=>String(x.v[14]||'')===cursor);if(!next)break;descendants.push(next);cursor=String(next.v[9]||'');}
  if(!checked&&descendants.some(x=>Boolean(x.v[4]))){sh.getRange(row,5).setValue(true);SpreadsheetApp.getUi().alert('Sequenza non valida','Prima riapri le attività successive già completate nello stesso percorso.',SpreadsheetApp.getUi().ButtonSet.OK);return;}
  if(checked){const now=new Date();sh.getRange(row,6).setValue('FATTO');sh.getRange(row,EVENT_APP.ACTIVITY.COMPLETED_AT).setValue(now);const next=descendants[0];if(next&&!(next.v[3] instanceof Date)&&eventNormalizeTextV7_(next.v[17])==='PRECEDENTE'){const due=eventAddDaysV14_(now,Number(next.v[18]||2));sh.getRange(next.row,4).setValue(due).setNumberFormat('dd/MM/yyyy');sh.getRange(next.row,EVENT_APP.ACTIVITY.AUTO_DUE).setValue(true);}}
  else{sh.getRange(row,EVENT_APP.ACTIVITY.COMPLETED_AT).clearContent();descendants.forEach(x=>{if(x.v[13]===true&&eventNormalizeTextV7_(x.v[17])==='PRECEDENTE'){sh.getRange(x.row,4).clearContent();sh.getRange(x.row,EVENT_APP.ACTIVITY.AUTO_DUE).setValue(false);}});}
}

function eventRefreshActivitiesV14_(){
  const sh=SpreadsheetApp.getActive().getSheetByName(EVENT_APP.SHEETS.TASKS);if(!sh)return;const rows=eventActivityRowsV14_(sh),today=eventDayV14_(new Date()),objectives={};rows.forEach(x=>{if(eventNormalizeTextV7_(x.v[7])==='OBIETTIVO'&&x.v[8])objectives[String(x.v[8])]={row:x.row,due:x.v[1],tasks:[]};});rows.forEach(x=>{if(eventNormalizeTextV7_(x.v[7])==='TASK'&&x.v[8]&&objectives[String(x.v[8])])objectives[String(x.v[8])].tasks.push(x);});
  Object.keys(objectives).forEach(oid=>{const o=objectives[oid],map={};o.tasks.forEach(x=>{if(x.v[9])map[String(x.v[9])]=x;});o.tasks.forEach(x=>{const checked=Boolean(sh.getRange(x.row,5).getValue());if(checked){sh.getRange(x.row,6).setValue('FATTO');return;}const prev=String(x.v[14]||'')?map[String(x.v[14])]:null,prevDone=!prev||Boolean(sh.getRange(prev.row,5).getValue());let due=sh.getRange(x.row,4).getValue();const mode=eventNormalizeTextV7_(x.v[17]||''),offset=Number(x.v[18]||2);if(prevDone&&!(due instanceof Date)&&mode==='PRECEDENTE'&&prev){const completed=prev.v[15] instanceof Date?prev.v[15]:new Date();due=eventAddDaysV14_(completed,offset);sh.getRange(x.row,4).setValue(due).setNumberFormat('dd/MM/yyyy');sh.getRange(x.row,EVENT_APP.ACTIVITY.AUTO_DUE).setValue(true);}let state='IN ATTESA';if(prevDone){const d=eventDayV14_(due);state=!d||d<=today?'DA FARE':'IN ATTESA';}sh.getRange(x.row,6).setValue(state);});const done=o.tasks.filter(x=>Boolean(sh.getRange(x.row,5).getValue())).length;let state=o.tasks.length&&done===o.tasks.length?'COMPLETATO':(done?'IN CORSO':'DA AVVIARE');const od=eventDayV14_(o.due);if(state!=='COMPLETATO'&&od&&od<today)state='IN RITARDO';sh.getRange(o.row,6).setValue(state);});
}

function eventInsertObjectiveV14_(sh,actionRow){
  sh.insertRowBefore(actionRow);const row=actionRow,p=eventActivityPaletteV14_(),count=eventActivityRowsV14_(sh).filter(x=>eventNormalizeTextV7_(x.v[7])==='OBIETTIVO').length,color=p[count%p.length];sh.getRange(row,1,1,EVENT_APP.ACTIVITY.COLS).clearContent();sh.getRange(row,1).setValue('Nuovo obiettivo');sh.getRange(row,3).setValue('＋');sh.getRange(row,EVENT_APP.ACTIVITY.TYPE).setValue('OBIETTIVO');sh.getRange(row,EVENT_APP.ACTIVITY.COLOR).setValue(color);eventEnsureObjectiveRowV14_(sh,row);sh.getRange(row,1,1,7).setBackground(color).setFontWeight('bold');sh.setActiveRange(sh.getRange(row,1));
}
function eventInsertTaskV14_(sh,headerRow){
  eventEnsureObjectiveRowV14_(sh,headerRow);const oid=String(sh.getRange(headerRow,EVENT_APP.ACTIVITY.OBJECTIVE_ID).getValue()||''),tasks=eventObjectiveTasksV14_(sh,oid),insertAfter=tasks.length?tasks[tasks.length-1].row:headerRow;sh.insertRowAfter(insertAfter);const row=insertAfter+1;sh.getRange(row,1,1,EVENT_APP.ACTIVITY.COLS).clearContent();sh.getRange(row,3).setValue('Nuova attività');sh.getRange(row,EVENT_APP.ACTIVITY.TYPE).setValue('TASK');sh.getRange(row,EVENT_APP.ACTIVITY.OBJECTIVE_ID).setValue(oid);sh.getRange(row,EVENT_APP.ACTIVITY.PATH).setValue('MANUALE');eventEnsureTaskRowV14_(sh,row);sh.getRange(row,1,1,7).setBackground(null).setFontWeight('normal');try{sh.getRange(row,1).shiftRowGroupDepth(1);}catch(err){}sh.setActiveRange(sh.getRange(row,3));eventRefreshActivitiesV14_();
}
function eventFormatActivityRowsV13_(){
  const sh=SpreadsheetApp.getActive().getSheetByName(EVENT_APP.SHEETS.TASKS);if(!sh)return;eventActivityRowsV14_(sh).forEach(x=>{const type=eventNormalizeTextV7_(x.v[7]);if(type==='OBIETTIVO'){let visible=String(sh.getRange(x.row,1).getBackground()||'').toLowerCase(),stored=String(x.v[12]||'#D9EAF7');if(!visible||visible==='#ffffff'){visible=stored;sh.getRange(x.row,1,1,7).setBackground(visible);}else sh.getRange(x.row,EVENT_APP.ACTIVITY.COLOR).setValue(visible);sh.getRange(x.row,1,1,7).setFontWeight('bold');sh.getRange(x.row,3).setValue('＋').setHorizontalAlignment('center');}else if(type==='TASK'){sh.getRange(x.row,5).insertCheckboxes();}});try{sh.hideColumns(8,12);}catch(err){}
}
function eventHandleActivitiesEditV13_(e){return eventHandleActivitiesEditV14_(e);}
function eventRefreshActivitiesV13_(){return eventRefreshActivitiesV14_();}
function eventEnsureObjectiveRowV13_(sh,row){return eventEnsureObjectiveRowV14_(sh,row);}
function eventEnsureTaskRowV13_(sh,row){return eventEnsureTaskRowV14_(sh,row);}
function eventToggleTaskV13_(sh,row,checked){return eventToggleTaskV14_(sh,row,checked);}

function eventNormalizeTextV7_(value){return String(value===null||value===undefined?'':value).trim().toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ');}
function eventMeta_(){const sheet=SpreadsheetApp.getActive().getSheetByName(EVENT_APP.SHEETS.META);if(!sheet)throw new Error('Foglio _META non trovato.');const values=sheet.getRange(1,1,sheet.getLastRow(),2).getDisplayValues(),out={};values.forEach(r=>{if(r[0])out[String(r[0]).trim()]=r[1];});return out;}
function eventWriteLocalMetaV7_(key,value){const sheet=SpreadsheetApp.getActive().getSheetByName(EVENT_APP.SHEETS.META);if(!sheet)return;const rows=sheet.getRange(1,1,Math.max(sheet.getLastRow(),1),2).getDisplayValues(),target=eventNormalizeTextV7_(key);for(let i=0;i<rows.length;i++)if(eventNormalizeTextV7_(rows[i][0])===target){sheet.getRange(i+1,2).setValue(value);return;}sheet.appendRow([key,value]);}
function eventCommitmentMapV7_(){return {'RIUNIONI/MISSIONI|*':'2-1','ALLENAMENTO|29ER':'28','ALLENAMENTO|420':'29','ALLENAMENTO|NACRA 15':'30','ALLENAMENTO|IQFOIL':'31','ALLENAMENTO|ILCA 6':'32','ALLENAMENTO|ILCA 4':'33-1','ALLENAMENTO INTERZONALE|ILCA 4':'33-2','ALLENAMENTO|KITEFOIL':'34','ALLENAMENTO|WING':'35','TEST FISICI|*':'37','COLLEGIALE|*':'38','REGATA|29ER':'39-1','REGATA|420':'40','REGATA|NACRA 15':'41','REGATA|IQFOIL':'42','REGATA|ILCA 6':'43','REGATA|ILCA 4':'44','REGATA|KITEFOIL':'45','REGATA|ILCA 4':'44','REGATA|KITEFOIL':'45','REGATA|WING':'46','OSS.REGATE ITA|*':'36','CAMP. NAZ GIOVANILE SINGOLO|*':'48-1','CAMP. NAZ GIOVANILE DOPPIO|*':'48-1','FOIL ACADEMY|*':'50','ISCRIZIONI REGATE|29ER':'47-1','ISCRIZIONI REGATE|420':'47-2','ISCRIZIONI REGATE|NACRA 15':'47-3','ISCRIZIONI REGATE|IQFOIL':'47-4','ISCRIZIONI REGATE|ILCA 6':'47-5','ISCRIZIONI REGATE|ILCA 4':'47-6','ISCRIZIONI REGATE|KITEFOIL':'47-7','ISCRIZIONI REGATE|WING':'47-8','ISCRIZIONE YWC|TUTTE':'171','MANUTENZIONE ALLENAMENTI|*':'39-3','MANUTENZIONE REGATE|*':'39-4','NOLEGGI ALLENAMENTI|*':'52','NOLEGGI REGATE|*':'51'};}
function eventCanonicalTypeV7_(value){const t=eventNormalizeTextV7_(value),aliases={'RIUNIONE/MISSIONE':'RIUNIONI/MISSIONI','RIUNIONI / MISSIONI':'RIUNIONI/MISSIONI','OSSERVAZIONE REGATA':'OSS.REGATE ITA','ISCRIZIONE REGATA':'ISCRIZIONI REGATE'};return aliases[t]||t;}
function eventCanonicalClassV7_(value){const c=eventNormalizeTextV7_(value);return c==='KITE'?'KITEFOIL':c;}
function eventResolveCommitmentV7_(meta){const type=eventCanonicalTypeV7_(meta.EVENT_TYPE||''),cls=eventCanonicalClassV7_(meta.EVENT_CLASS||''),map=eventCommitmentMapV7_();return map[type+'|'+cls]||map[type+'|*']||map[type+'|TUTTE']||'';}
function eventTechnicianDirectoryV7_(){return [['ZAGGIA','Leonardo','Zaggia'],['CRISI','Andrea','Crisi'],['RAVEGLIA','Matteo','Raveglia'],['CARICATO','Francesco','Caricato'],['PICCIAU','Gianluigi','Picciau'],['SENSINI','Alessandra','Sensini'],['NUICOLUCCI','Matteo','Nuicolucci'],['CAMBONI','Mattia','Camboni'],['CANGEMI','Antonino','Cangemi'],['LOPERFIDO','Daniel','Loperfido']];}
function eventPopulateTechniciansV7_(meta){const raw=eventNormalizeTextV7_(meta.EVENT_TECHNICIANS||'');if(!raw)return[];const found=eventTechnicianDirectoryV7_().filter(x=>raw.indexOf(x[0])>=0);if(!found.length)return[];const full=found.map(x=>x[1]+' '+x[2]);eventWriteLocalMetaV7_('EVENT_TECHNICIANS',full.join(', '));const sheet=SpreadsheetApp.getActive().getSheetByName(EVENT_APP.SHEETS.PARTICIPANTS);if(!sheet)return full;const rows=sheet.getRange(3,1,20,19).getDisplayValues(),existing=new Set(rows.map(r=>eventNormalizeTextV7_((r[0]||'')+' '+(r[1]||''))));found.forEach(x=>{const key=eventNormalizeTextV7_(x[1]+' '+x[2]);if(existing.has(key))return;let target=-1;for(let i=0;i<rows.length;i++)if(!String(rows[i][0]||'').trim()&&!String(rows[i][1]||'').trim()){target=i+3;rows[i][0]=x[1];rows[i][1]=x[2];break;}if(target<0)return;sheet.getRange(target,1).setValue(x[1]);sheet.getRange(target,2).setValue(x[2]);sheet.getRange(target,8).setValue('TECNICO');sheet.getRange(target,9).setValue('CONFERMATO');existing.add(key);});return full;}
function eventDocumentReplacements_(meta){const participants=SpreadsheetApp.getActive().getSheetByName(EVENT_APP.SHEETS.PARTICIPANTS),rows=participants?participants.getRange(3,1,20,19).getDisplayValues():[],athletes=rows.filter(r=>{const has=String(r[0]||'').trim()||String(r[1]||'').trim(),role=eventNormalizeTextV7_(r[7]||'ATLETA'),status=eventNormalizeTextV7_(r[8]);return has&&role==='ATLETA'&&status!=='ASSENTE';}),athleteLines=athletes.map(r=>{const name=[r[0],r[1]].filter(Boolean).join(' ').trim();return r[2]?name+' – '+r[2]:name;}),start=eventParseIsoDate_(meta.EVENT_START),end=eventParseIsoDate_(meta.EVENT_END);return {'tipo':eventTitleCase_(meta.EVENT_TYPE||''),'classe':meta.EVENT_CLASS||'','luogo':meta.EVENT_LOCATION||'','localita':meta.EVENT_LOCATION||'','zona':meta.EVENT_ZONE||'','circolo':meta.EVENT_CLUB||'','tecnico':meta.EVENT_TECHNICIANS||'','lista tecnici':meta.EVENT_TECHNICIANS||'','lista atleti':athleteLines.join('\n'),'numero atleti':athletes.length?String(athletes.length):'','hotel/struttura':meta.EVENT_LODGING||'','data inizio':eventFormatDate_(start),'data fine':eventFormatDate_(end),'data in':eventFormatDate_(start),'data out':eventFormatDate_(end),'date':eventFormatDateRange_(start,end),'data di oggi':eventFormatDate_(new Date()),'data oggi':eventFormatDate_(new Date()),'impegno':meta.EVENT_COMMITMENT||''};}
function eventFillDocument_(docId,replacements,key){const doc=DocumentApp.openById(docId),body=doc.getBody(),reps=Object.assign({},replacements);if(key==='GOMMONE')reps.N=replacements.zona||'';if(key==='OSPITALITA')reps.N=replacements['numero atleti']||'';Object.keys(reps).forEach(k=>{const v=reps[k];if(v===''||v===null||v===undefined)return;body.replaceText('(?i)\\{\\{'+eventEscapeRegex_(k)+'\\}\\}',eventSafeReplacement_(v));});if(reps.classe)body.replaceText('(?i)\\{classe\\}\\}',eventSafeReplacement_(reps.classe));doc.saveAndClose();}
function eventDocumentName_(key,r){const base={CONV_ATLETI:['Convocazione atleti',r.tipo,r.classe,r.luogo,r.date],CONV_TECNICO:['Convocazione tecnico',r.tecnico,r.tipo,r.classe,r.date],GOMMONE:['Richiesta gommone',r.zona?r.zona+' Zona':'',r.tipo,r.classe,r.date],OSPITALITA:['Richiesta ospitalità',r.circolo,r.tipo,r.classe,r.date],RINGRAZIAMENTO:['Ringraziamento',r.circolo,r.tipo,r.classe,r.date]}[key]||[r.tipo,r.classe,r.date];return base.filter(Boolean).join(' - ');}
function eventRegisterDocument_(meta,spec,file){if(!meta.MASTER_SPREADSHEET_ID)return;const master=SpreadsheetApp.openById(meta.MASTER_SPREADSHEET_ID),sheet=master.getSheetByName(EVENT_APP.MASTER_DOCUMENTS);if(!sheet)return;sheet.appendRow(['DOC-'+Utilities.getUuid(),meta.EVENT_ID,spec.key,spec.id,file.getName(),file.getId(),file.getUrl(),'BOZZA MODIFICABILE',new Date(),'']);}
function eventParseIsoDate_(s){const m=String(s||'').match(/^(\d{4})-(\d{2})-(\d{2})$/);return m?new Date(Number(m[1]),Number(m[2])-1,Number(m[3]),12,0,0):null;}
function eventFormatDate_(d){if(!(d instanceof Date)||isNaN(d))return'';const months=['gennaio','febbraio','marzo','aprile','maggio','giugno','luglio','agosto','settembre','ottobre','novembre','dicembre'];return d.getDate()+' '+months[d.getMonth()]+' '+d.getFullYear();}
function eventFormatDateRange_(s,e){if(!(s instanceof Date))return'';if(!(e instanceof Date))return eventFormatDate_(s);if(s.getTime()===e.getTime())return eventFormatDate_(s);const months=['gennaio','febbraio','marzo','aprile','maggio','giugno','luglio','agosto','settembre','ottobre','novembre','dicembre'];if(s.getFullYear()===e.getFullYear()&&s.getMonth()===e.getMonth())return s.getDate()+'-'+e.getDate()+' '+months[s.getMonth()]+' '+s.getFullYear();return eventFormatDate_(s)+' - '+eventFormatDate_(e);}
function eventTitleCase_(s){return String(s||'').toLowerCase().replace(/(^|\s|[-/])([a-zà-öø-ÿ])/g,(_,p,c)=>p+c.toUpperCase());}
function eventEscapeRegex_(s){return String(s||'').replace(/[.*+?^${}()|[\]\\]/g,'\\$&');}
function eventSafeReplacement_(s){return String(s||'').replace(/\\/g,'\\\\').replace(/\$/g,'\\$');}
