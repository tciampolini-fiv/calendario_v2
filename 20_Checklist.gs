const CHECKLIST_DEPENDENCY_DELAY_DAYS = 2;
const OBJECTIVE_DEFAULT_DUE_DAYS = -5;
const CHECKLIST_V13 = Object.freeze({
  WIDTH:20,
  COL:{ID:1,EVENT:2,ORDER:3,TASK:4,CATEGORY:5,DUE:6,STATUS:7,PRIORITY:8,SOURCE:9,AUTO_KEY:10,NOTE:11,COMPLETED:12,UPDATED:13,NUMBER:14,DEPENDENCY:15,AUTO_BLOCK:16,OBJECTIVE:17,STEP:18,AUTO_DUE:19,PREVIOUS:20}
});

function getChecklistRowsForEventRaw_(eventId) {
  const rows = sh_(APP.SHEETS.CHECKLIST).getDataRange().getValues();
  return rows.slice(1).filter(r => String(r[1]) === String(eventId)).map(r => ({
    id:r[0], eventId:r[1], order:r[2], task:r[3], category:r[4], dueDate:r[5], status:r[6],
    priority:r[7], source:r[8], autoKey:r[9], note:r[10], completedAt:r[11], updatedAt:r[12],
    number:r[13], dependencyId:r[14], autoBlock:r[15], objectiveId:r[16], stepOrder:r[17],
    autoDue:r[18] === true, previousTaskId:r[19]
  }));
}

function getObjectivesForEvent_(eventId) {
  const sheet = sh_(APP.SHEETS.OBJECTIVES);
  const rows = sheet.getDataRange().getValues();
  const tasks = getChecklistRowsForEventRaw_(eventId), grouped={};
  tasks.forEach(t=>{if(!t.objectiveId)return;if(!grouped[t.objectiveId])grouped[t.objectiveId]=[];grouped[t.objectiveId].push(t);});
  return rows.slice(1).filter(r=>String(r[1])===String(eventId)).map(r=>{
    const list=(grouped[r[0]]||[]).sort((a,b)=>Number(a.stepOrder||a.order||0)-Number(b.stepOrder||b.order||0));
    const done=list.filter(t=>normalize_(t.status)==='FATTO'||normalize_(t.status)==='COMPLETATA').length,total=list.length;
    let state=total&&done===total?'COMPLETATO':(done?'IN CORSO':'DA AVVIARE');
    if(state!=='COMPLETATO'&&r[4] instanceof Date){const d=new Date(r[4]),today=new Date();d.setHours(0,0,0,0);today.setHours(0,0,0,0);if(d<today)state='IN RITARDO';}
    return{id:r[0],eventId:r[1],name:r[2],color:r[3],dueDate:r[4],order:r[5],source:r[6],templateKey:r[7],notes:r[8],createdAt:r[9],updatedAt:r[10],status:state,done:done,total:total,tasks:list};
  }).sort((a,b)=>Number(a.order||0)-Number(b.order||0));
}

function objectivePalette_(){return['#D9EAF7','#FCE8B2','#EADCF8','#D9EAD3','#F4CCCC','#D0E0E3','#FCE5CD','#D9D2E9','#CFE2F3','#E2F0D9'];}
function objectiveColorForIndex_(i){const p=objectivePalette_();return p[Math.max(0,Number(i)||0)%p.length];}
function addDays_(date,days){if(!(date instanceof Date))return'';const d=new Date(date);d.setDate(d.getDate()+Number(days||0));return d;}
function objectiveDueDate_(event){return addDays_(event&&event[APP.CALENDAR_HEADERS.START],OBJECTIVE_DEFAULT_DUE_DAYS);}

function objectiveKeyFromTask_(task,autoKey,category){
  const k=normalize_(autoKey||task),t=normalize_(task),c=normalize_(category);
  if(k.indexOf('GOMMONE')>=0||t.indexOf('GOMMONE')>=0)return'GOMMONE';
  if(k.indexOf('SOGGIORNO')>=0||t.indexOf('HOTEL')>=0||t.indexOf('SOGGIORNO')>=0||t.indexOf('ALLOGGIO')>=0)return'SOGGIORNO';
  if(k.indexOf('PASTI')>=0||t.indexOf('PASTI')>=0)return'PASTI';
  if(k.indexOf('OSPITALITA')>=0||t.indexOf('OSPITALITA')>=0||k.indexOf('RINGRAZIAMENTO')>=0)return'OSPITALITA_CIRCOLO';
  if(k.indexOf('CONV_ATLETI')>=0||k==='CHECK_CONFERME'||t.indexOf('CONVOCAZIONE ATLETI')>=0||t.indexOf('PRESENZE')>=0)return'CONV_ATLETI';
  if(k.indexOf('CONV_TECNICO')>=0||t.indexOf('CONVOCAZIONE TECNICO')>=0)return'CONV_TECNICO';
  if(k.indexOf('VIAGGIO_TECNICO')>=0||t.indexOf('VIAGGIO TECNICO')>=0)return'VIAGGIO_TECNICO';
  if(t.indexOf('NAVE')>=0||t.indexOf('CARRELLO')>=0||t.indexOf('MEZZ')>=0)return'TRASPORTO_MEZZI';
  if(c==='AMMINISTRAZIONE')return'AMMINISTRAZIONE';
  return'ALTRO';
}
function objectiveDisplayName_(key,task,note){
  if(key==='VIAGGIO_TECNICO'){
    const text=normalize_(String(task||'')+' '+String(note||'')),people=technicianDirectory_();
    for(let i=0;i<people.length;i++)if(text.indexOf(people[i][0])>=0)return'Viaggio tecnico – '+people[i][1]+' '+people[i][2];
  }
  return({GOMMONE:'Gommone',SOGGIORNO:'Soggiorno',PASTI:'Pasti',OSPITALITA_CIRCOLO:'Ospitalità Circolo',CONV_ATLETI:'Convocazione atleti',CONV_TECNICO:'Convocazione tecnico',VIAGGIO_TECNICO:'Viaggio tecnico',TRASPORTO_MEZZI:'Trasporto / Mezzi',AMMINISTRAZIONE:'Amministrazione',ALTRO:'Altro'})[key]||key;
}

function ensureObjectivesBackendV13_(){
  const ss=SpreadsheetApp.getActive();let obj=ss.getSheetByName(APP.SHEETS.OBJECTIVES);
  if(!obj){obj=ss.insertSheet(APP.SHEETS.OBJECTIVES);obj.getRange(1,1,1,11).setValues([['ID OBIETTIVO','ID EVENTO','NOME','COLORE','SCADENZA OBIETTIVO','ORDINE','ORIGINE','TEMPLATE KEY','NOTE','DATA INSERIMENTO','ULTIMO AGGIORNAMENTO']]);obj.setFrozenRows(1);}
  const checklist=sh_(APP.SHEETS.CHECKLIST),headers=checklist.getRange(1,1,1,Math.max(checklist.getLastColumn(),20)).getValues()[0];
  const extra=['ID OBIETTIVO','ORDINE STEP','SCADENZA AUTOMATICA','ID TASK PRECEDENTE'];for(let i=0;i<extra.length;i++)if(normalize_(headers[16+i])!==normalize_(extra[i]))checklist.getRange(1,17+i).setValue(extra[i]);
  return obj;
}

function migrateChecklistToObjectivesV13_(){
  const objSheet=ensureObjectivesBackendV13_(),check=sh_(APP.SHEETS.CHECKLIST),rows=check.getDataRange().getValues();
  if(rows.slice(1).some(r=>String(r[16]||'').trim()))return{migrated:false,message:'Migrazione già eseguita'};
  const cal=sh_(APP.SHEETS.CALENDAR),cr=cal.getDataRange().getValues(),headers=cr[0],events={};
  for(let i=1;i<cr.length;i++){const event={};headers.forEach((h,c)=>event[String(h||'').trim()]=cr[i][c]);const id=String(event[APP.CALENDAR_HEADERS.ID]||'').trim();if(id)events[id]=event;}
  const now=new Date(),objectiveRows=[],map={},tasksByObjective={};
  rows.slice(1).forEach((r,index)=>{
    const eventId=String(r[1]||'').trim(),task=String(r[3]||'').trim();if(!eventId||!task)return;
    const key=objectiveKeyFromTask_(task,r[9],r[4]),name=objectiveDisplayName_(key,task,r[10]),mk=eventId+'|'+normalize_(name);
    let o=map[mk];if(!o){const count=objectiveRows.filter(x=>x[1]===eventId).length;o=['OBJ-'+Utilities.getUuid(),eventId,name,objectiveColorForIndex_(count),objectiveDueDate_(events[eventId]),(count+1)*10,'MIGRATO',key,'',now,now];map[mk]=o;objectiveRows.push(o);tasksByObjective[o[0]]=[];}
    tasksByObjective[o[0]].push({row:index+2,id:r[0],order:Number(r[13]||r[2]||index+1)});
  });
  if(objectiveRows.length)objSheet.getRange(objSheet.getLastRow()+1,1,objectiveRows.length,11).setValues(objectiveRows);
  Object.keys(tasksByObjective).forEach(oid=>{const list=tasksByObjective[oid].sort((a,b)=>a.order-b.order);list.forEach((x,i)=>check.getRange(x.row,17,1,4).setValues([[oid,(i+1)*10,false,i?list[i-1].id:'']]));});
  syncAllChecklistLocks_();return{migrated:true,objectives:objectiveRows.length};
}

function syncChecklistLocksForEvent_(eventId){
  ensureObjectivesBackendV13_();const sheet=sh_(APP.SHEETS.CHECKLIST),rows=sheet.getDataRange().getValues(),groups={};
  for(let i=1;i<rows.length;i++){if(String(rows[i][1])!==String(eventId)||!rows[i][16])continue;const oid=String(rows[i][16]);if(!groups[oid])groups[oid]=[];groups[oid].push({row:i+1,v:rows[i]});}
  const today=new Date();today.setHours(0,0,0,0);const now=new Date();
  Object.keys(groups).forEach(oid=>{
    const list=groups[oid].sort((a,b)=>Number(a.v[17]||a.v[13]||a.v[2]||0)-Number(b.v[17]||b.v[13]||b.v[2]||0));
    list.forEach((item,i)=>{
      const current=normalize_(item.v[6]);if(current==='FATTO'||current==='COMPLETATA')return;
      const prev=i?list[i-1]:null,prevDone=!prev||['FATTO','COMPLETATA'].includes(normalize_(prev.v[6]));let due=item.v[5];
      if(prev&&prevDone&&!(due instanceof Date)){
        const base=prev.v[11] instanceof Date?prev.v[11]:now;due=addDays_(base,CHECKLIST_DEPENDENCY_DELAY_DAYS);sheet.getRange(item.row,6).setValue(due).setNumberFormat('dd/MM/yyyy');sheet.getRange(item.row,19).setValue(true);item.v[5]=due;
      }
      let desired='IN ATTESA';if(prevDone){if(!(due instanceof Date))desired=i===0?'DA FARE':'IN ATTESA';else{const d=new Date(due);d.setHours(0,0,0,0);desired=d<=today?'DA FARE':'IN ATTESA';}}
      if(current!==desired){sheet.getRange(item.row,7).setValue(desired);sheet.getRange(item.row,13).setValue(now);}
      const prevId=prev?String(prev.v[0]||''):'';if(String(item.v[19]||'')!==prevId)sheet.getRange(item.row,20).setValue(prevId);
    });
  });
}

function getChecklistForEvent_(eventId){syncChecklistLocksForEvent_(eventId);return getChecklistRowsForEventRaw_(eventId).sort((a,b)=>Number(a.order||0)-Number(b.order||0));}
function syncAllChecklistLocks_(){const rows=sh_(APP.SHEETS.CHECKLIST).getDataRange().getValues(),ids=new Set();for(let i=1;i<rows.length;i++)if(rows[i][1])ids.add(String(rows[i][1]));ids.forEach(syncChecklistLocksForEvent_);}
function getChecklistProfileForEvent_(event){const type=normalize_(event[APP.CALENDAR_HEADERS.TYPE]),rows=sh_(APP.SHEETS.EVENT_TYPE_CONFIG).getDataRange().getValues();for(let i=1;i<rows.length;i++)if(normalize_(rows[i][0])===type)return normalize_(rows[i][3])||type;return type;}
function getChecklistRulesForEvent_(){return{};}

function activityProfileForEvent_(event){const type=normalize_(event[APP.CALENDAR_HEADERS.TYPE]);if(type.indexOf('REGATA INT')>=0)return'REGATA_INTERNAZIONALE';if(type==='REGATA')return'REGATA';return'ALLENAMENTO';}
function standardObjectiveTemplatesForEvent_(event){
  const p=activityProfileForEvent_(event),international=p==='REGATA_INTERNAZIONALE',regatta=p==='REGATA';
  const base=international?-35:(regatta?-21:-14),stay=international?-30:(regatta?-21:-14),travel=international?-20:(regatta?-14:-10),conv=international?-15:(regatta?-12:-10),meals=international?-12:(regatta?-10:-7);
  return[
    {key:'GOMMONE',name:'Gommone',offset:base,steps:['Chiedere disponibilità','Ricevere conferma','Pagare AFOR','Inviare contabile']},
    {key:'SOGGIORNO',name:'Soggiorno',offset:stay,steps:['Chiedere disponibilità hotel','Ricevere risposta hotel','Confermare camere','Pagare anticipo']},
    {key:'PASTI',name:'Pasti',offset:meals,steps:['Definire soluzione pasti','Ricevere conferma']},
    {key:'VIAGGIO_TECNICO',name:'Viaggio tecnico',offset:travel,steps:['Definire viaggio','Prenotare','Inviare biglietto']},
    {key:'CONV_ATLETI',name:'Convocazione atleti',offset:conv,steps:['Inviare convocazione','Ricevere conferme']},
    {key:'CONV_TECNICO',name:'Convocazione tecnico',offset:conv,steps:['Inviare convocazione tecnico']},
    {key:'OSPITALITA_CIRCOLO',name:'Ospitalità Circolo',offset:base,steps:['Chiedere ospitalità','Ricevere conferma']}
  ];
}

function generateChecklistForEvent_(eventId,event){
  ensureObjectivesBackendV13_();const profile=getChecklistProfileForEvent_(event);if(!profile||profile==='NESSUNO')return 0;if(getObjectivesForEvent_(eventId).length)return 0;
  const objSheet=sh_(APP.SHEETS.OBJECTIVES),taskSheet=sh_(APP.SHEETS.CHECKLIST),start=event[APP.CALENDAR_HEADERS.START],now=new Date(),templates=standardObjectiveTemplatesForEvent_(event),techs=resolveFullTechnicians_(event[APP.CALENDAR_HEADERS.TECHNICIANS]),expanded=[];
  templates.forEach(t=>{if(t.key==='VIAGGIO_TECNICO'&&techs.length)techs.forEach(p=>expanded.push(Object.assign({},t,{name:'Viaggio tecnico – '+p.name+' '+p.surname,templateKey:t.key+':'+normalize_(p.surname)})));else expanded.push(Object.assign({},t,{templateKey:t.key}));});
  const objs=[],tasks=[];let globalNo=0;
  expanded.forEach((t,i)=>{const oid='OBJ-'+Utilities.getUuid();objs.push([oid,eventId,t.name,objectiveColorForIndex_(i),objectiveDueDate_(event),(i+1)*10,'STANDARD',t.templateKey,'',now,now]);let previous='';t.steps.forEach((step,j)=>{globalNo++;const id='TASK-'+Utilities.getUuid(),due=j===0?addDays_(start,t.offset):'',status=j===0&&due instanceof Date&&due<=now?'DA FARE':'IN ATTESA';tasks.push([id,eventId,(i+1)*100+(j+1)*10,step,'ALTRO',due,status,'','STANDARD',t.key+(j?':'+(j+1):''),'','',now,globalNo,previous,previous?'DIPENDENZA':'',oid,(j+1)*10,false,previous]);previous=id;});});
  if(objs.length)objSheet.getRange(objSheet.getLastRow()+1,1,objs.length,11).setValues(objs);if(tasks.length)taskSheet.getRange(taskSheet.getLastRow()+1,1,tasks.length,20).setValues(tasks);syncChecklistLocksForEvent_(eventId);return tasks.length;
}
function generateChecklistForSelectedEvent(){const event=selectedEvent_(),eventId=ensureEventId_(event),added=generateChecklistForEvent_(eventId,event);SpreadsheetApp.getUi().alert('Attività evento',added?'Nuove attività create: '+added:'Gli obiettivi sono già presenti.',SpreadsheetApp.getUi().ButtonSet.OK);}

function calendarObjectiveProgressText_(eventId){const objs=getObjectivesForEvent_(eventId);return objs.map(o=>{const symbol=o.status==='COMPLETATO'?'✓':o.status==='IN RITARDO'?'⚠':'•',boxes=o.tasks.map(t=>['FATTO','COMPLETATA'].includes(normalize_(t.status))?'☑':'☐').join('');return symbol+' '+o.name+(boxes?'  '+boxes:'');}).join('\n');}
function calendarTasksNowText_(eventId){syncChecklistLocksForEvent_(eventId);const objs=getObjectivesForEvent_(eventId),names={};objs.forEach(o=>names[o.id]=o.name);return getChecklistRowsForEventRaw_(eventId).filter(t=>normalize_(t.status)==='DA FARE').sort((a,b)=>Number(a.order||0)-Number(b.order||0)).map(t=>(names[t.objectiveId]?names[t.objectiveId]+' — ':'')+t.task).join('\n');}
function refreshCalendarActivityDashboardForEvent_(eventId){const found=findCalendarEventById_(eventId);if(!found)return;const sheet=sh_(APP.SHEETS.CALENDAR),headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0],a=headers.indexOf(APP.CALENDAR_HEADERS.CHECKLIST)+1,n=headers.indexOf(APP.CALENDAR_HEADERS.NEXT_ACTION)+1;if(a>0)sheet.getRange(found.row,a).setValue(calendarObjectiveProgressText_(eventId)).setWrap(true);if(n>0)sheet.getRange(found.row,n).setValue(calendarTasksNowText_(eventId)).setWrap(true);}
