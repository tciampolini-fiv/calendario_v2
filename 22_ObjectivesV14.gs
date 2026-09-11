const ACTIVITY_BACKEND_V14 = Object.freeze({
  WIDTH: 23,
  COL: Object.freeze({
    ID:1, EVENT:2, ORDER:3, TASK:4, CATEGORY:5, DUE:6, STATUS:7, PRIORITY:8,
    SOURCE:9, AUTO_KEY:10, NOTE:11, COMPLETED:12, UPDATED:13, NUMBER:14,
    DEPENDENCY:15, AUTO_BLOCK:16, OBJECTIVE:17, STEP:18, AUTO_DUE:19,
    PREVIOUS:20, PATH:21, DUE_MODE:22, OFFSET:23
  })
});

function activityPaletteV14_(){
  return ['#D9EAF7','#FCE8B2','#EADCF8','#D9EAD3','#F4CCCC','#D0E0E3','#FCE5CD','#D9D2E9','#CFE2F3','#E2F0D9'];
}
function activityColorV14_(i){const p=activityPaletteV14_();return p[Math.max(0,Number(i)||0)%p.length];}
function activityAddDaysV14_(date,days){if(!(date instanceof Date))return'';const d=new Date(date);d.setDate(d.getDate()+Number(days||0));return d;}
function activityDayV14_(date){if(!(date instanceof Date))return null;const d=new Date(date);d.setHours(0,0,0,0);return d;}
function activityIsDoneV14_(value){const s=normalize_(value);return s==='FATTO'||s==='COMPLETATA'||s==='COMPLETATO';}

function ensureActivityBackendV14_(){
  const ss=SpreadsheetApp.getActive();
  let objectives=ss.getSheetByName(APP.SHEETS.OBJECTIVES);
  if(!objectives){
    objectives=ss.insertSheet(APP.SHEETS.OBJECTIVES);
    objectives.getRange(1,1,1,11).setValues([['ID OBIETTIVO','ID EVENTO','NOME','COLORE','SCADENZA OBIETTIVO','ORDINE','ORIGINE','TEMPLATE KEY','NOTE','DATA INSERIMENTO','ULTIMO AGGIORNAMENTO']]);
    objectives.setFrozenRows(1);objectives.hideSheet();
  }
  const checklist=sh_(APP.SHEETS.CHECKLIST);
  if(checklist.getMaxColumns()<ACTIVITY_BACKEND_V14.WIDTH)checklist.insertColumnsAfter(checklist.getMaxColumns(),ACTIVITY_BACKEND_V14.WIDTH-checklist.getMaxColumns());
  checklist.getRange(1,17,1,7).setValues([['ID OBIETTIVO','ORDINE STEP','SCADENZA AUTOMATICA','ID TASK PRECEDENTE','PERCORSO','MODALITA SCADENZA','OFFSET GIORNI']]);
  return {objectives:objectives,checklist:checklist};
}

function getActivityTasksForEventV14_(eventId){
  const rows=sh_(APP.SHEETS.CHECKLIST).getDataRange().getValues();
  return rows.slice(1).filter(r=>String(r[1])===String(eventId)&&String(r[0]||'').trim()).map(r=>({
    id:r[0],eventId:r[1],order:r[2],task:r[3],category:r[4],dueDate:r[5],status:r[6],priority:r[7],
    source:r[8],autoKey:r[9],note:r[10],completedAt:r[11],updatedAt:r[12],number:r[13],dependencyId:r[14],
    autoBlock:r[15],objectiveId:r[16],stepOrder:r[17],autoDue:r[18]===true,previousTaskId:r[19],
    path:r[20]||'MANUALE',dueMode:r[21]||'',offsetDays:r[22]
  }));
}

function getActivityObjectivesForEventV14_(eventId){
  ensureActivityBackendV14_();
  const tasks=getActivityTasksForEventV14_(eventId),byObjective={};
  tasks.forEach(t=>{if(!t.objectiveId)return;(byObjective[t.objectiveId]=byObjective[t.objectiveId]||[]).push(t);});
  const today=activityDayV14_(new Date());
  return sh_(APP.SHEETS.OBJECTIVES).getDataRange().getValues().slice(1)
    .filter(r=>String(r[1])===String(eventId)&&String(r[0]||'').trim())
    .map(r=>{
      const list=(byObjective[r[0]]||[]).sort((a,b)=>Number(a.stepOrder||a.order||0)-Number(b.stepOrder||b.order||0));
      const done=list.filter(t=>activityIsDoneV14_(t.status)).length,total=list.length;
      let status=total&&done===total?'COMPLETATO':(done?'IN CORSO':'DA AVVIARE');
      const due=activityDayV14_(r[4]);if(status!=='COMPLETATO'&&due&&due<today)status='IN RITARDO';
      return {id:r[0],eventId:r[1],name:r[2],color:r[3],dueDate:r[4],order:r[5],source:r[6],templateKey:r[7],notes:r[8],createdAt:r[9],updatedAt:r[10],status:status,done:done,total:total,tasks:list};
    }).sort((a,b)=>Number(a.order||0)-Number(b.order||0));
}

function activityProfileForEventV14_(event){
  const type=normalize_(event&&event[APP.CALENDAR_HEADERS.TYPE]);
  const rows=sh_(APP.SHEETS.EVENT_TYPE_CONFIG).getDataRange().getValues();
  for(let i=1;i<rows.length;i++)if(normalize_(rows[i][0])===type)return normalize_(rows[i][3])||'NESSUNO';
  return 'NESSUNO';
}

function activityConfigForProfileV14_(profile){
  const rows=sh_(APP.SHEETS.CHECKLIST_CONFIG).getDataRange().getValues();
  return rows.slice(1).filter(r=>normalize_(r[0])===normalize_(profile)&&(r[12]===true||normalize_(r[12])==='SI')).map(r=>({
    profile:r[0],objectiveKey:String(r[1]||'').trim(),objectiveName:String(r[2]||'').trim(),objectiveOrder:Number(r[3]||0),
    color:String(r[4]||''),objectiveDueBase:normalize_(r[5]),objectiveDueOffset:Number(r[6]||0),path:String(r[7]||'MANUALE').trim(),
    step:Number(r[8]||0),task:String(r[9]||'').trim(),dueMode:normalize_(r[10]),offset:Number(r[11]||0),autoKey:String(r[13]||'').trim(),note:String(r[14]||'').trim()
  })).sort((a,b)=>a.objectiveOrder-b.objectiveOrder||a.step-b.step);
}

function canonicalObjectiveKeyV14_(templateKey,name){
  const k=normalize_(templateKey),n=normalize_(name);
  if(k.indexOf('VIAGGIO_TECNICO:')===0)return String(templateKey);
  if(k==='VIAGGIO_TECNICO'||n.indexOf('VIAGGIO TECNICO')===0)return 'VIAGGIO_TECNICO';
  if(k==='CIRCOLO'||k.indexOf('OSPITALITA')>=0||k.indexOf('RINGRAZIAMENTO')>=0||n==='CIRCOLO'||n.indexOf('OSPITALITA CIRCOLO')>=0||n.indexOf('RINGRAZIAMENTO CIRCOLO')>=0)return 'CIRCOLO';
  if(k==='VITTO_ALLOGGIO'||k==='SOGGIORNO'||k==='PASTI'||n==='SOGGIORNO'||n==='PASTI'||n.indexOf('VITTO')>=0)return 'VITTO_ALLOGGIO';
  if(k==='CONVOCAZIONI'||k.indexOf('CONV_')===0||n.indexOf('CONVOCAZION')>=0)return 'CONVOCAZIONI';
  if(k.indexOf('GOMMONE')>=0||n==='GOMMONE')return 'GOMMONE';
  if(k==='ZONA'||n==='ZONA')return 'ZONA';
  return String(templateKey||'').trim()||('CUSTOM:'+normalize_(name));
}
function canonicalObjectiveNameV14_(key,current){
  const k=normalize_(key);
  if(k==='CIRCOLO')return'Circolo';
  if(k==='VITTO_ALLOGGIO')return'Vitto-Alloggio';
  if(k==='GOMMONE')return'Gommone';
  if(k==='CONVOCAZIONI')return'Convocazioni';
  if(k==='ZONA')return'Zona';
  return current||key;
}
function canonicalTaskKeyV14_(task,autoKey){
  const k=normalize_(autoKey),t=normalize_(task);
  const direct={
    'GOMMONE':'GOMMONE_RICHIESTA','GOMMONE_RICHIESTA':'GOMMONE_RICHIESTA','CONF_GOMMONE':'GOMMONE_CONFERMA','GOMMONE_CONFERMA':'GOMMONE_CONFERMA',
    'OSPITALITA_CIRCOLO':'CIRCOLO_RICHIESTA','CIRCOLO_RICHIESTA':'CIRCOLO_RICHIESTA','CONF_OSPITALITA':'CIRCOLO_CONFERMA','CIRCOLO_CONFERMA':'CIRCOLO_CONFERMA','RINGRAZIAMENTO_CIRCOLO':'CIRCOLO_RINGRAZIAMENTO','CIRCOLO_RINGRAZIAMENTO':'CIRCOLO_RINGRAZIAMENTO',
    'SOGGIORNO':'ALLOGGIO_RICHIESTA','ALLOGGIO_RICHIESTA':'ALLOGGIO_RICHIESTA','CONF_SOGGIORNO':'ALLOGGIO_CONFERMA','ALLOGGIO_CONFERMA':'ALLOGGIO_CONFERMA','ALLOGGIO_PAGAMENTO':'ALLOGGIO_PAGAMENTO','ALLOGGIO_FATTURA':'ALLOGGIO_FATTURA',
    'PASTI':'PASTI_RICHIESTA','PASTI_RICHIESTA':'PASTI_RICHIESTA','CONF_PASTI':'PASTI_CONFERMA','PASTI_CONFERMA':'PASTI_CONFERMA','PASTI_PAGAMENTO':'PASTI_PAGAMENTO',
    'CONV_ATLETI':'CONV_ATLETI_INVIO','CONV_ATLETI_INVIO':'CONV_ATLETI_INVIO','CHECK_CONFERME':'CONV_ATLETI_CONFERMA','CONV_ATLETI_CONFERMA':'CONV_ATLETI_CONFERMA','CONV_TECNICO':'CONV_TECNICI_INVIO','CONV_TECNICI_INVIO':'CONV_TECNICI_INVIO',
    'VIAGGIO_TECNICO':'VIAGGIO_PRENOTAZIONE','VIAGGIO_PRENOTAZIONE':'VIAGGIO_PRENOTAZIONE','VIAGGIO_PAGAMENTO':'VIAGGIO_PAGAMENTO','ZONA_CONFERMA':'ZONA_CONFERMA'
  };
  if(direct[k])return direct[k];
  if(t==='RICHIESTA GOMMONE')return'GOMMONE_RICHIESTA';if(t==='CONFERMA GOMMONE')return'GOMMONE_CONFERMA';
  if(t.indexOf('OSPITALITA')>=0&&t.indexOf('CONFERMA')<0)return'CIRCOLO_RICHIESTA';if(t.indexOf('CONFERMA OSPITALITA')>=0)return'CIRCOLO_CONFERMA';if(t.indexOf('RINGRAZIAMENTO')>=0)return'CIRCOLO_RINGRAZIAMENTO';
  if(t.indexOf('RICHIESTA SOGGIORNO')>=0||t.indexOf('CONTATTARE HOTEL')>=0)return'ALLOGGIO_RICHIESTA';if(t.indexOf('CONFERMA SOGGIORNO')>=0||t.indexOf('CONFERMA HOTEL')>=0)return'ALLOGGIO_CONFERMA';
  if(t.indexOf('SOLUZIONE PASTI')>=0)return'PASTI_RICHIESTA';if(t.indexOf('CONFERMA PASTI')>=0)return'PASTI_CONFERMA';
  if(t.indexOf('CONVOCAZIONE ATLETI')>=0)return'CONV_ATLETI_INVIO';if(t.indexOf('CONFERMA PRESENZE')>=0||t.indexOf('CHECK CONFERMA')>=0)return'CONV_ATLETI_CONFERMA';if(t.indexOf('CONVOCAZIONE TECNICO')>=0)return'CONV_TECNICI_INVIO';
  if(t==='VIAGGIO TECNICO')return'VIAGGIO_PRENOTAZIONE';
  return k;
}

function compactLegacyObjectivesV14_(eventId){
  const b=ensureActivityBackendV14_(),objRows=b.objectives.getDataRange().getValues(),taskRows=b.checklist.getDataRange().getValues(),groups={};
  for(let i=1;i<objRows.length;i++){
    const r=objRows[i];if(String(r[1])!==String(eventId)||!r[0])continue;
    const key=canonicalObjectiveKeyV14_(r[7],r[2]);
    if(!['CIRCOLO','VITTO_ALLOGGIO','CONVOCAZIONI','GOMMONE','ZONA'].includes(normalize_(key)))continue;
    (groups[key]=groups[key]||[]).push({row:i+1,id:String(r[0]),values:r});
  }
  const deleteRows=[];
  Object.keys(groups).forEach(key=>{
    const list=groups[key].sort((a,b)=>Number(a.values[5]||0)-Number(b.values[5]||0));if(!list.length)return;
    const keep=list[0];b.objectives.getRange(keep.row,3).setValue(canonicalObjectiveNameV14_(key,keep.values[2]));b.objectives.getRange(keep.row,8).setValue(key);
    list.slice(1).forEach(drop=>{
      for(let i=1;i<taskRows.length;i++)if(String(taskRows[i][16]||'')===drop.id)b.checklist.getRange(i+1,17).setValue(keep.id);
      deleteRows.push(drop.row);
    });
  });
  deleteRows.sort((a,b)=>b-a).forEach(r=>b.objectives.deleteRow(r));
}

function activityTechniciansV14_(raw){
  const resolved=typeof resolveFullTechnicians_==='function'?resolveFullTechnicians_(raw):[];
  if(resolved&&resolved.length)return resolved.map(x=>({name:x.name,surname:x.surname,full:(x.name+' '+x.surname).trim()}));
  return String(raw||'').split(/[,;\n]+/).map(x=>x.trim()).filter(Boolean).map(x=>{const p=x.split(/\s+/);return{name:p.slice(0,-1).join(' '),surname:p[p.length-1],full:x};});
}
function activityObjectiveDueV14_(event,base,offset){const d=normalize_(base)==='FINE'?event[APP.CALENDAR_HEADERS.END]:event[APP.CALENDAR_HEADERS.START];return activityAddDaysV14_(d,offset);}
function activityTaskDueV14_(event,mode,offset){const m=normalize_(mode);if(m==='INIZIO')return activityAddDaysV14_(event[APP.CALENDAR_HEADERS.START],offset);if(m==='FINE')return activityAddDaysV14_(event[APP.CALENDAR_HEADERS.END],offset);return'';}
function activityTaskCategoryV14_(task){const t=normalize_(task);if(t.indexOf('CONVOCAZ')>=0||t.indexOf('ATLETI')>=0||t.indexOf('TECNIC')>=0)return'CONVOCAZIONI';if(t.indexOf('PAGAMENTO')>=0||t.indexOf('FATTURA')>=0)return'AMMINISTRAZIONE';if(t.indexOf('CIRCOLO')>=0||t.indexOf('OSPITALITA')>=0||t.indexOf('RINGRAZIAMENTO')>=0||t.indexOf('COMITATO DI ZONA')>=0)return'ORGANIZZAZIONE';return'LOGISTICA';}

function ensureDefaultObjectivesForEventV14_(eventId,event){
  const backend=ensureActivityBackendV14_();compactLegacyObjectivesV14_(eventId);
  const profile=activityProfileForEventV14_(event);if(!profile||profile==='NESSUNO')return 0;
  const config=activityConfigForProfileV14_(profile);if(!config.length)return 0;
  const grouped={};config.forEach(c=>(grouped[c.objectiveKey]=grouped[c.objectiveKey]||[]).push(c));
  let objectives=backend.objectives.getDataRange().getValues(),existing={};
  for(let i=1;i<objectives.length;i++)if(String(objectives[i][1])===String(eventId)&&objectives[i][0])existing[canonicalObjectiveKeyV14_(objectives[i][7],objectives[i][2])]={row:i+1,values:objectives[i]};
  const techs=activityTechniciansV14_(event[APP.CALENDAR_HEADERS.TECHNICIANS]),specs=[];
  Object.keys(grouped).forEach(key=>{
    const rows=grouped[key];
    if(normalize_(key)==='VIAGGIO_TECNICO'){
      techs.forEach((p,i)=>{const full=p.full||(p.name+' '+p.surname).trim(),suffix=normalize_(full),surname=p.surname||full;specs.push({key:'VIAGGIO_TECNICO:'+suffix,name:'Viaggio tecnico – '+surname,rows:rows,order:rows[0].objectiveOrder+i});});
    }else specs.push({key:key,name:rows[0].objectiveName,rows:rows,order:rows[0].objectiveOrder});
  });
  const now=new Date();let added=0;
  specs.forEach((spec,specIndex)=>{
    let found=existing[spec.key],oid;
    const first=spec.rows[0],due=activityObjectiveDueV14_(event,first.objectiveDueBase,first.objectiveDueOffset);
    if(!found){
      oid='OBJ-'+Utilities.getUuid();backend.objectives.appendRow([oid,eventId,spec.name,first.color||activityColorV14_(specIndex),due,spec.order,'STANDARD',spec.key,'',now,now]);
      found={row:backend.objectives.getLastRow(),values:[oid,eventId,spec.name,first.color,due,spec.order,'STANDARD',spec.key,'',now,now]};existing[spec.key]=found;
    }else{
      oid=String(found.values[0]);const source=normalize_(found.values[6]);
      if(source==='STANDARD'||source==='MIGRATO'){backend.objectives.getRange(found.row,3).setValue(spec.name);backend.objectives.getRange(found.row,8).setValue(spec.key);if(!(found.values[4] instanceof Date)&&due)backend.objectives.getRange(found.row,5).setValue(due).setNumberFormat('dd/MM/yyyy');}
    }
    let taskRows=backend.checklist.getDataRange().getValues(),assigned=[];
    for(let i=1;i<taskRows.length;i++)if(String(taskRows[i][1])===String(eventId)&&String(taskRows[i][16]||'')===oid)assigned.push({row:i+1,values:taskRows[i]});
    const used=new Set(),prevByPath={};
    spec.rows.forEach((cfg,j)=>{
      let match=assigned.find(x=>!used.has(x.row)&&canonicalTaskKeyV14_(x.values[3],x.values[9])===normalize_(cfg.autoKey));
      const previous=prevByPath[cfg.path]||'';
      if(match){
        used.add(match.row);const r=match.values,hasDue=r[5] instanceof Date,dueValue=hasDue?r[5]:activityTaskDueV14_(event,cfg.dueMode,cfg.offset);
        backend.checklist.getRange(match.row,4).setValue(cfg.task);backend.checklist.getRange(match.row,5).setValue(activityTaskCategoryV14_(cfg.task));
        if(!hasDue&&dueValue)backend.checklist.getRange(match.row,6).setValue(dueValue).setNumberFormat('dd/MM/yyyy');
        backend.checklist.getRange(match.row,10).setValue(cfg.autoKey);backend.checklist.getRange(match.row,17,1,7).setValues([[oid,cfg.step,!hasDue&&!!dueValue,previous,cfg.path,cfg.dueMode,cfg.offset]]);
        prevByPath[cfg.path]=String(r[0]);
      }else{
        const id='TASK-'+Utilities.getUuid(),taskDue=activityTaskDueV14_(event,cfg.dueMode,cfg.offset),today=activityDayV14_(new Date()),dueDay=activityDayV14_(taskDue),status=previous?'IN ATTESA':(!dueDay||dueDay<=today?'DA FARE':'IN ATTESA');
        backend.checklist.appendRow([id,eventId,spec.order*100+cfg.step,cfg.task,activityTaskCategoryV14_(cfg.task),taskDue,status,'','STANDARD',cfg.autoKey,cfg.note||'','',now,'',previous,previous?'DIPENDENZA':'',oid,cfg.step,!!taskDue,previous,cfg.path,cfg.dueMode,cfg.offset]);
        prevByPath[cfg.path]=id;added++;
      }
    });
  });
  syncActivityStatesForEventV14_(eventId,event);refreshCalendarActivityDashboardV14_(eventId);return added;
}

function syncActivityStatesForEventV14_(eventId,event){
  const b=ensureActivityBackendV14_();if(!event){const found=findCalendarEventById_(eventId);event=found?found.event:null;}if(!event)return;
  const rows=b.checklist.getDataRange().getValues(),groups={};
  for(let i=1;i<rows.length;i++){
    const r=rows[i];if(String(r[1])!==String(eventId)||!r[16]||!r[0])continue;
    const path=String(r[20]||'MANUALE'),key=String(r[16])+'|'+path;(groups[key]=groups[key]||[]).push({row:i+1,v:r});
  }
  const today=activityDayV14_(new Date()),now=new Date();
  Object.keys(groups).forEach(key=>{
    const list=groups[key].sort((a,b)=>Number(a.v[17]||a.v[2]||0)-Number(b.v[17]||b.v[2]||0));
    list.forEach((item,i)=>{
      const r=item.v,done=activityIsDoneV14_(r[6]),prev=i?list[i-1]:null,prevDone=!prev||activityIsDoneV14_(prev.v[6]),mode=normalize_(r[21]||''),offset=Number(r[22]||2);let due=r[5],auto=r[18]===true;
      const prevId=prev?String(prev.v[0]||''):'';
      if(String(r[19]||'')!==prevId)b.checklist.getRange(item.row,20).setValue(prevId);
      if(done)return;
      if(!prevDone){
        if(mode==='PRECEDENTE'&&auto&&due instanceof Date){b.checklist.getRange(item.row,6).clearContent();b.checklist.getRange(item.row,19).setValue(false);due='';auto=false;}
        if(normalize_(r[6])!=='IN ATTESA')b.checklist.getRange(item.row,7).setValue('IN ATTESA');return;
      }
      if(!(due instanceof Date)){
        if(mode==='PRECEDENTE'&&prev){const base=prev.v[11] instanceof Date?prev.v[11]:now;due=activityAddDaysV14_(base,offset);auto=true;}
        else{due=activityTaskDueV14_(event,mode,offset);auto=!!due;}
        if(due){b.checklist.getRange(item.row,6).setValue(due).setNumberFormat('dd/MM/yyyy');b.checklist.getRange(item.row,19).setValue(auto);}
      }
      const dueDay=activityDayV14_(due),desired=!dueDay||dueDay<=today?'DA FARE':'IN ATTESA';
      if(normalize_(r[6])!==desired){b.checklist.getRange(item.row,7).setValue(desired);b.checklist.getRange(item.row,13).setValue(now);}
    });
  });
}

function refreshCalendarActivityDashboardV14_(eventId){
  syncActivityStatesForEventV14_(eventId);const found=findCalendarEventById_(eventId);if(!found)return;
  const sheet=sh_(APP.SHEETS.CALENDAR),map=headerMap_(sheet),objectives=getActivityObjectivesForEventV14_(eventId);
  const progress=objectives.map(o=>{const symbol=o.status==='COMPLETATO'?'✓':o.status==='IN RITARDO'?'⚠':'•',boxes=o.tasks.map(t=>activityIsDoneV14_(t.status)?'☑':'☐').join('');return symbol+' '+o.name+(boxes?'  '+boxes:'');}).join('\n');
  const actions=[];objectives.forEach(o=>o.tasks.filter(t=>normalize_(t.status)==='DA FARE').forEach(t=>actions.push(o.name+' — '+t.task)));
  if(map[APP.CALENDAR_HEADERS.CHECKLIST])sheet.getRange(found.row,map[APP.CALENDAR_HEADERS.CHECKLIST]).setValue(progress).setWrap(true).setVerticalAlignment('top');
  if(map[APP.CALENDAR_HEADERS.NEXT_ACTION])sheet.getRange(found.row,map[APP.CALENDAR_HEADERS.NEXT_ACTION]).setValue(actions.join('\n')).setWrap(true).setVerticalAlignment('top');
}
function refreshAllCalendarActivityDashboardsV14_(){const rows=sh_(APP.SHEETS.CALENDAR).getDataRange().getValues(),headers=rows[0],idIndex=headers.indexOf(APP.CALENDAR_HEADERS.ID);for(let i=1;i<rows.length;i++){const id=String(rows[i][idIndex]||'').trim();if(id&&getActivityObjectivesForEventV14_(id).length)refreshCalendarActivityDashboardV14_(id);}}
