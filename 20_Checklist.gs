const CHECKLIST_DEPENDENCY_DELAY_DAYS = 2;
const OBJECTIVE_DEFAULT_DUE_DAYS = -5;

function getChecklistRowsForEventRaw_(eventId) {
  const rows = sh_(APP.SHEETS.CHECKLIST).getDataRange().getValues();
  return rows.slice(1).filter(r => String(r[1]) === String(eventId)).map(r => ({
    id:r[0], eventId:r[1], objectiveId:r[2], order:r[3], task:r[4], dueDate:r[5], status:r[6],
    note:r[7], source:r[8], autoKey:r[9], completedAt:r[10], updatedAt:r[11], autoDue:r[12] === true,
    previousTaskId:r[13]
  }));
}

function getObjectivesForEvent_(eventId) {
  const sheet = sh_(APP.SHEETS.OBJECTIVES);
  const rows = sheet.getDataRange().getValues();
  const tasks = getChecklistRowsForEventRaw_(eventId);
  const byObjective = {};
  tasks.forEach(t => { if (!byObjective[t.objectiveId]) byObjective[t.objectiveId] = []; byObjective[t.objectiveId].push(t); });
  return rows.slice(1).filter(r => String(r[1]) === String(eventId)).map(r => {
    const list = (byObjective[r[0]] || []).sort((a,b)=>Number(a.order||0)-Number(b.order||0));
    const done = list.filter(t=>normalize_(t.status)==='FATTO').length;
    const total = list.length;
    let state = total && done === total ? 'COMPLETATO' : (done ? 'IN CORSO' : 'DA AVVIARE');
    const due = r[4];
    if (state !== 'COMPLETATO' && due instanceof Date) {
      const d = new Date(due); d.setHours(0,0,0,0); const today = new Date(); today.setHours(0,0,0,0);
      if (d < today) state = 'IN RITARDO';
    }
    return {id:r[0], eventId:r[1], name:r[2], color:r[3], dueDate:r[4], order:r[5], source:r[6], templateKey:r[7], notes:r[8], createdAt:r[9], updatedAt:r[10], status:state, done:done, total:total, tasks:list};
  }).sort((a,b)=>Number(a.order||0)-Number(b.order||0));
}

function objectivePalette_(){return ['#D9EAF7','#FCE8B2','#EADCF8','#D9EAD3','#F4CCCC','#D0E0E3','#FCE5CD','#D9D2E9','#CFE2F3','#E2F0D9'];}
function objectiveColorForIndex_(index){const p=objectivePalette_();return p[Math.max(0,index)%p.length];}
function objectiveDueDate_(event){const start=event && event[APP.CALENDAR_HEADERS.START];if(!(start instanceof Date))return'';const d=new Date(start);d.setDate(d.getDate()+OBJECTIVE_DEFAULT_DUE_DAYS);return d;}
function objectiveKeyFromTask_(task,autoKey,category){
  const k=normalize_(autoKey||task), t=normalize_(task), c=normalize_(category);
  if(k.indexOf('GOMMONE')>=0||t.indexOf('GOMMONE')>=0)return'GOMMONE';
  if(k.indexOf('SOGGIORNO')>=0||t.indexOf('HOTEL')>=0||t.indexOf('SOGGIORNO')>=0||t.indexOf('ALLOGGIO')>=0)return'SOGGIORNO';
  if(k.indexOf('PASTI')>=0||t.indexOf('PASTI')>=0)return'PASTI';
  if(k.indexOf('OSPITALITA')>=0||t.indexOf('OSPITALITA')>=0||t.indexOf('CIRCOLO')>=0||k.indexOf('RINGRAZIAMENTO')>=0)return'OSPITALITA CIRCOLO';
  if(k.indexOf('CONV_ATLETI')>=0||k==='CHECK_CONFERME'||t.indexOf('CONVOCAZIONE ATLETI')>=0||t.indexOf('PRESENZE')>=0)return'CONVOCAZIONE ATLETI';
  if(k.indexOf('CONV_TECNICO')>=0||t.indexOf('CONVOCAZIONE TECNICO')>=0)return'CONVOCAZIONE TECNICO';
  if(k.indexOf('VIAGGIO_TECNICO')>=0||t.indexOf('VIAGGIO TECNICO')>=0)return'VIAGGIO TECNICO';
  if(t.indexOf('NAVE')>=0||t.indexOf('CARRELLO')>=0||t.indexOf('MEZZ')>=0)return'TRASPORTO / MEZZI';
  if(c==='AMMINISTRAZIONE')return'AMMINISTRAZIONE';
  return'ALTRO';
}
function objectiveDisplayName_(key,task,note){
  if(key==='VIAGGIO TECNICO'){
    const text=(String(task||'')+' '+String(note||'')).trim();
    const people=technicianDirectory_();
    for(let i=0;i<people.length;i++)if(normalize_(text).indexOf(people[i][0])>=0)return'Viaggio tecnico – '+people[i][1]+' '+people[i][2];
  }
  const names={'GOMMONE':'Gommone','SOGGIORNO':'Soggiorno','PASTI':'Pasti','OSPITALITA CIRCOLO':'Ospitalità Circolo','CONVOCAZIONE ATLETI':'Convocazione atleti','CONVOCAZIONE TECNICO':'Convocazione tecnico','VIAGGIO TECNICO':'Viaggio tecnico','TRASPORTO / MEZZI':'Trasporto / Mezzi','AMMINISTRAZIONE':'Amministrazione','ALTRO':'Altro'};
  return names[key]||key;
}

function migrateChecklistToObjectivesV13_(){
  const cal=sh_(APP.SHEETS.CALENDAR), calRows=cal.getDataRange().getValues(), headers=calRows[0], idx={};headers.forEach((h,i)=>idx[String(h||'').trim()]=i);
  const events={};for(let i=1;i<calRows.length;i++){const id=String(calRows[i][idx[APP.CALENDAR_HEADERS.ID]]||'').trim();if(!id)continue;const e={};headers.forEach((h,c)=>e[String(h||'').trim()]=calRows[i][c]);events[id]=e;}
  const checklist=sh_(APP.SHEETS.CHECKLIST), rows=checklist.getDataRange().getValues();
  if(normalize_(rows[0][2])==='ID OBIETTIVO')return {migrated:false};
  const objSheet=sh_(APP.SHEETS.OBJECTIVES), now=new Date(), objectives=[], objectiveMap={}, taskRows=[];
  rows.slice(1).forEach((r,index)=>{
    const eventId=String(r[1]||'').trim();if(!eventId||!String(r[3]||'').trim())return;
    const key=objectiveKeyFromTask_(r[3],r[9],r[4]);const display=objectiveDisplayName_(key,r[3],r[10]);const mapKey=eventId+'|'+normalize_(display);
    let obj=objectiveMap[mapKey];
    if(!obj){const sameEventCount=objectives.filter(x=>x[1]===eventId).length;obj=['OBJ-'+Utilities.getUuid(),eventId,display,objectiveColorForIndex_(sameEventCount),objectiveDueDate_(events[eventId]),(sameEventCount+1)*10,'MIGRATO',key,'',now,now];objectiveMap[mapKey]=obj;objectives.push(obj);}
    const existingOrder=Number(r[13]||0)>0?Number(r[13]):Number(r[2]||0);taskRows.push([r[0]||'TASK-'+Utilities.getUuid(),eventId,obj[0],existingOrder||((index+1)*10),r[3]||'',r[5]||'',normalize_(r[6])==='COMPLETATA'?'FATTO':(r[6]||'IN ATTESA'),r[10]||'',r[8]||'MIGRATO',r[9]||'',r[11]||'',r[12]||now,false,r[14]||'']);
  });
  if(objSheet.getLastRow()>1)objSheet.getRange(2,1,objSheet.getLastRow()-1,11).clearContent();
  if(objectives.length)objSheet.getRange(2,1,objectives.length,11).setValues(objectives);
  checklist.clearContents();
  checklist.getRange(1,1,1,14).setValues([['ID TASK','ID EVENTO','ID OBIETTIVO','ORDINE STEP','ATTIVITA','SCADENZA','STATO','NOTE','ORIGINE','AUTO KEY','DATA COMPLETAMENTO','ULTIMO AGGIORNAMENTO','SCADENZA AUTOMATICA','ID TASK PRECEDENTE']]);
  if(taskRows.length)checklist.getRange(2,1,taskRows.length,14).setValues(taskRows);
  Object.keys(events).forEach(syncChecklistLocksForEvent_);SpreadsheetApp.flush();return{migrated:true,objectives:objectives.length,tasks:taskRows.length};
}

function syncChecklistLocksForEvent_(eventId){
  const sheet=sh_(APP.SHEETS.CHECKLIST),rows=sheet.getDataRange().getValues(),items=[];
  for(let i=1;i<rows.length;i++)if(String(rows[i][1])===String(eventId))items.push({row:i+1,v:rows[i]});
  const groups={};items.forEach(x=>{const id=String(x.v[2]||'');if(!groups[id])groups[id]=[];groups[id].push(x);});
  const today=new Date();today.setHours(0,0,0,0);const now=new Date();
  Object.keys(groups).forEach(id=>{
    const list=groups[id].sort((a,b)=>Number(a.v[3]||0)-Number(b.v[3]||0));
    for(let i=0;i<list.length;i++){
      const item=list[i],status=normalize_(item.v[6]);if(status==='FATTO')continue;
      const prev=i?list[i-1]:null;const prevDone=!prev||normalize_(prev.v[6])==='FATTO';let due=item.v[5];
      if(prev&&prevDone&&!(due instanceof Date)){
        const base=prev.v[10] instanceof Date?prev.v[10]:now;due=new Date(base);due.setHours(12,0,0,0);due.setDate(due.getDate()+CHECKLIST_DEPENDENCY_DELAY_DAYS);
        sheet.getRange(item.row,6).setValue(due).setNumberFormat('dd/MM/yyyy');sheet.getRange(item.row,13).setValue(true);item.v[5]=due;
      }
      let desired='IN ATTESA';
      if(prevDone){if(!(due instanceof Date))desired=i===0?'DA FARE':'IN ATTESA';else{const d=new Date(due);d.setHours(0,0,0,0);desired=d<=today?'DA FARE':'IN ATTESA';}}
      if(status!==desired)sheet.getRange(item.row,7).setValue(desired);
      const prevId=prev?String(prev.v[0]||''):'';if(String(item.v[13]||'')!==prevId)sheet.getRange(item.row,14).setValue(prevId);
      sheet.getRange(item.row,12).setValue(now);
    }
  });
}

function getChecklistForEvent_(eventId){syncChecklistLocksForEvent_(eventId);return getChecklistRowsForEventRaw_(eventId).sort((a,b)=>Number(a.order||0)-Number(b.order||0));}
function syncAllChecklistLocks_(){const rows=sh_(APP.SHEETS.CHECKLIST).getDataRange().getValues(),ids=new Set();for(let i=1;i<rows.length;i++)if(rows[i][1])ids.add(String(rows[i][1]));ids.forEach(syncChecklistLocksForEvent_);}

function getChecklistProfileForEvent_(event){const type=normalize_(event[APP.CALENDAR_HEADERS.TYPE]),rows=sh_(APP.SHEETS.EVENT_TYPE_CONFIG).getDataRange().getValues();for(let i=1;i<rows.length;i++)if(normalize_(rows[i][0])===type)return normalize_(rows[i][3])||type;return type;}
function getChecklistRulesForEvent_(){return {};}

function standardObjectiveTemplatesForEvent_(event){
  const type=normalize_(event[APP.CALENDAR_HEADERS.TYPE]);const international=type.indexOf('REGATA INT')>=0||type==='REGATA';
  const base=international?-35:-14, stay=international?-30:-14, travel=international?-20:-10, conv=international?-15:-10, meals=international?-12:-7;
  return [
    {key:'GOMMONE',name:'Gommone',first:'Chiedere disponibilità',offset:base,steps:['Chiedere disponibilità','Ricevere conferma','Pagare AFOR','Inviare contabile']},
    {key:'SOGGIORNO',name:'Soggiorno',first:'Chiedere disponibilità hotel',offset:stay,steps:['Chiedere disponibilità hotel','Ricevere risposta hotel','Confermare camere','Pagare anticipo']},
    {key:'PASTI',name:'Pasti',first:'Definire soluzione pasti',offset:meals,steps:['Definire soluzione pasti','Ricevere conferma']},
    {key:'VIAGGIO_TECNICO',name:'Viaggio tecnico',first:'Definire viaggio',offset:travel,steps:['Definire viaggio','Prenotare','Inviare biglietto']},
    {key:'CONV_ATLETI',name:'Convocazione atleti',first:'Inviare convocazione',offset:conv,steps:['Inviare convocazione','Ricevere conferme']},
    {key:'CONV_TECNICO',name:'Convocazione tecnico',first:'Inviare convocazione tecnico',offset:conv,steps:['Inviare convocazione tecnico']},
    {key:'OSPITALITA_CIRCOLO',name:'Ospitalità Circolo',first:'Chiedere ospitalità',offset:base,steps:['Chiedere ospitalità','Ricevere conferma']}
  ];
}
function addDays_(d,n){if(!(d instanceof Date))return'';const x=new Date(d);x.setDate(x.getDate()+Number(n||0));return x;}
function generateChecklistForEvent_(eventId,event){
  const profile=getChecklistProfileForEvent_(event);if(!profile||profile==='NESSUNO')return 0;
  const objSheet=sh_(APP.SHEETS.OBJECTIVES), taskSheet=sh_(APP.SHEETS.CHECKLIST);const existing=getObjectivesForEvent_(eventId);if(existing.length)return 0;
  const start=event[APP.CALENDAR_HEADERS.START],now=new Date(),templates=standardObjectiveTemplatesForEvent_(event),techs=resolveFullTechnicians_(event[APP.CALENDAR_HEADERS.TECHNICIANS]);
  const expanded=[];templates.forEach(t=>{if(t.key==='VIAGGIO_TECNICO'&&techs.length){techs.forEach(p=>expanded.push(Object.assign({},t,{name:'Viaggio tecnico – '+p.name+' '+p.surname,templateKey:t.key+':'+normalize_(p.surname)})));}else expanded.push(Object.assign({},t,{templateKey:t.key}));});
  const objRows=[],taskRows=[];expanded.forEach((t,i)=>{const oid='OBJ-'+Utilities.getUuid();objRows.push([oid,eventId,t.name,objectiveColorForIndex_(i),objectiveDueDate_(event),(i+1)*10,'STANDARD',t.templateKey,'',now,now]);t.steps.forEach((step,j)=>taskRows.push(['TASK-'+Utilities.getUuid(),eventId,oid,(j+1)*10,step,j===0?addDays_(start,t.offset):'',j===0&&addDays_(start,t.offset) instanceof Date&&addDays_(start,t.offset)<=now?'DA FARE':'IN ATTESA','', 'STANDARD',t.key+(j?':'+(j+1):''),'',now,false,'']));});
  if(objRows.length)objSheet.getRange(objSheet.getLastRow()+1,1,objRows.length,11).setValues(objRows);if(taskRows.length)taskSheet.getRange(taskSheet.getLastRow()+1,1,taskRows.length,14).setValues(taskRows);syncChecklistLocksForEvent_(eventId);return taskRows.length;
}
function generateChecklistForSelectedEvent(){const event=selectedEvent_(),eventId=ensureEventId_(event),added=generateChecklistForEvent_(eventId,event);SpreadsheetApp.getUi().alert('Attività evento',added?'Nuove attività create: '+added:'Gli obiettivi sono già presenti.',SpreadsheetApp.getUi().ButtonSet.OK);}

function calendarObjectiveProgressText_(eventId){
  const objs=getObjectivesForEvent_(eventId);if(!objs.length)return'';
  return objs.map(o=>{const mark=o.status==='COMPLETATO'?'✓':o.status==='IN RITARDO'?'⚠':'•';const boxes=o.tasks.map(t=>normalize_(t.status)==='FATTO'?'☑':'☐').join('');return mark+' '+o.name+'  '+boxes;}).join('\n');
}
function calendarTasksNowText_(eventId){syncChecklistLocksForEvent_(eventId);const objs=getObjectivesForEvent_(eventId),names={};objs.forEach(o=>names[o.id]=o.name);return getChecklistRowsForEventRaw_(eventId).filter(t=>normalize_(t.status)==='DA FARE').sort((a,b)=>Number(a.order||0)-Number(b.order||0)).map(t=>(names[t.objectiveId]?names[t.objectiveId]+' — ':'')+t.task).join('\n');}
