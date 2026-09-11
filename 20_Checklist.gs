const CHECKLIST_DEPENDENCY_DELAY_DAYS = 2;
const OBJECTIVE_DEFAULT_DUE_DAYS = -5;

function getChecklistRowsForEventRaw_(eventId){return getActivityTasksForEventV14_(eventId);}
function getObjectivesForEvent_(eventId){return getActivityObjectivesForEventV14_(eventId);}
function ensureObjectivesBackendV13_(){return ensureActivityBackendV14_().objectives;}
function objectivePalette_(){return activityPaletteV14_();}
function objectiveColorForIndex_(i){return activityColorV14_(i);}
function addDays_(date,days){return activityAddDaysV14_(date,days);}
function objectiveDueDate_(event){return activityAddDaysV14_(event&&event[APP.CALENDAR_HEADERS.START],OBJECTIVE_DEFAULT_DUE_DAYS);}

function eventHasEnded_(event){const end=event&&event[APP.CALENDAR_HEADERS.END];if(!(end instanceof Date))return false;const d=new Date(end),today=new Date();d.setHours(0,0,0,0);today.setHours(0,0,0,0);return d<today;}
function dependentTaskDueDate_(completedAt,delayDays){return activityAddDaysV14_(completedAt,delayDays===undefined?CHECKLIST_DEPENDENCY_DELAY_DAYS:delayDays);}

function syncChecklistLocksForEvent_(eventId){
  const found=findCalendarEventById_(eventId);syncActivityStatesForEventV14_(eventId,found&&found.event);
}
function syncAllChecklistLocks_(){
  const rows=sh_(APP.SHEETS.CHECKLIST).getDataRange().getValues(),ids=new Set();for(let i=1;i<rows.length;i++)if(rows[i][1])ids.add(String(rows[i][1]));ids.forEach(syncChecklistLocksForEvent_);
}
function getChecklistForEvent_(eventId){syncChecklistLocksForEvent_(eventId);return getActivityTasksForEventV14_(eventId).sort((a,b)=>Number(a.order||0)-Number(b.order||0));}
function getChecklistProfileForEvent_(event){return activityProfileForEventV14_(event);}
function getChecklistRulesForEvent_(){return{};}

function generateChecklistForEvent_(eventId,event){return ensureDefaultObjectivesForEventV14_(eventId,event);}
function generateChecklistForSelectedEvent(){const event=selectedEvent_(),eventId=ensureEventId_(event),added=ensureDefaultObjectivesForEventV14_(eventId,event);SpreadsheetApp.getUi().alert('Attività evento',added?'Nuove attività create: '+added:'Tutti gli obiettivi previsti sono già presenti.',SpreadsheetApp.getUi().ButtonSet.OK);}

function calendarObjectiveProgressText_(eventId){
  return getActivityObjectivesForEventV14_(eventId).map(o=>{const symbol=o.status==='COMPLETATO'?'✓':o.status==='IN RITARDO'?'⚠':'•',boxes=o.tasks.map(t=>activityIsDoneV14_(t.status)?'☑':'☐').join('');return symbol+' '+o.name+(boxes?'  '+boxes:'');}).join('\n');
}
function calendarTasksNowText_(eventId){
  syncChecklistLocksForEvent_(eventId);const names={};getActivityObjectivesForEventV14_(eventId).forEach(o=>names[o.id]=o.name);return getActivityTasksForEventV14_(eventId).filter(t=>normalize_(t.status)==='DA FARE').sort((a,b)=>Number(a.order||0)-Number(b.order||0)).map(t=>(names[t.objectiveId]?names[t.objectiveId]+' — ':'')+t.task).join('\n');
}
function refreshCalendarActivityDashboardForEvent_(eventId){return refreshCalendarActivityDashboardV14_(eventId);}

function migrateChecklistToObjectivesV13_(){
  ensureActivityBackendV14_();const ids=new Set(),rows=sh_(APP.SHEETS.CHECKLIST).getDataRange().getValues();for(let i=1;i<rows.length;i++)if(rows[i][1])ids.add(String(rows[i][1]));let changed=0;ids.forEach(id=>{const found=findCalendarEventById_(id);if(found)changed+=ensureDefaultObjectivesForEventV14_(id,found.event);});return{migrated:true,added:changed};
}
