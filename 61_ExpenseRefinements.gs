const EVENT_SHEET_TEMPLATE_ID_V3 = '1obNkubjpGhC7o8AX5Tc986UBGdyJs26b476b-GZN3Sc';

function prepareEventSheetForSelectedEventV2() {
  const event=selectedEvent_(),eventId=ensureEventId_(event),folder=createWorkFolderForEvent_(eventId,event,event._row),result=getOrCreateEventSheetV3_(eventId,event,folder.folderId),child=result.spreadsheet;
  ensureChecklistBackendHeadersV3_();ensureParticipantsBackendHeadersV2_();
  if(result.created)generateChecklistForEvent_(eventId,event);
  refreshTasksPreserveTemplateV6_(eventId,event,child);
  if(isEventSheetV7_(child))initializeEventSheetV7FromCalendar_(child,eventId,event,folder.folderId);
  else if(!result.created){refreshParticipantsV3FromBackend_(eventId,event,child);refreshExpensesV4FromBackend_(eventId,event,child,folder.folderId);hideEventSheetTechnicalColumnsV3_(child);}
  ensureFastEventTaskTriggerV9_(child);setEventSheetLink_(event._row,child.getUrl());writeEventMetaV5_(child,eventId,event,folder.folderId);SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(result.created?'Scheda evento creata dal modello':'Scheda evento aggiornata',result.created?'È stata creata una nuova Scheda evento con task, IMP e tecnici già inizializzati.':'La Scheda evento collegata è stata aggiornata senza modificarne la struttura.',SpreadsheetApp.getUi().ButtonSet.OK);
  return{created:result.created,id:child.getId(),url:child.getUrl()};
}

function isEventSheetV7_(child){
  const expenses=child&&child.getSheetByName(EVENT_SHEET.SHEETS.EXPENSES),participants=child&&child.getSheetByName(PARTICIPANTS_V3.SHEET);
  if(!expenses||!participants||normalize_(participants.getRange('H2').getDisplayValue())!=='RUOLO')return false;
  const isV10=normalize_(expenses.getRange('A19').getDisplayValue())==='N. PREVENTIVO'&&normalize_(expenses.getRange('B19').getDisplayValue())==='TIPO PAGAMENTO'&&normalize_(expenses.getRange('C19').getDisplayValue())==='IMPORTO';
  const isV8=normalize_(expenses.getRange('A19').getDisplayValue())==='DESCRIZIONE'&&normalize_(expenses.getRange('C19').getDisplayValue())==='MOVIMENTO';
  const isV7=normalize_(expenses.getRange('A13').getDisplayValue())==='PREVENTIVO'&&normalize_(expenses.getRange('E13').getDisplayValue())==='PAGAMENTO 1';
  return isV10||isV8||isV7;
}
function isEventSheetV8_(child){
  const expenses=child&&child.getSheetByName(EVENT_SHEET.SHEETS.EXPENSES);
  if(!expenses)return false;
  const isV10=normalize_(expenses.getRange('A19').getDisplayValue())==='N. PREVENTIVO'&&normalize_(expenses.getRange('B19').getDisplayValue())==='TIPO PAGAMENTO'&&normalize_(expenses.getRange('C19').getDisplayValue())==='IMPORTO';
  const isV8=normalize_(expenses.getRange('A19').getDisplayValue())==='DESCRIZIONE'&&normalize_(expenses.getRange('C19').getDisplayValue())==='MOVIMENTO';
  return isV10||isV8;
}

function refreshTasksPreserveTemplateV6_(eventId,event,child){
  ensureChecklistBackendHeadersV3_();seedPresenceCheckTaskV3_(eventId,event,child);ensureTaskNumbersAndDefaultDependenciesV3_(eventId,event);syncChecklistLocksForEvent_(eventId);
  const backend=sh_(APP.SHEETS.CHECKLIST).getDataRange().getValues().slice(1).filter(r=>String(r[1])===String(eventId)),byId={};backend.forEach(r=>{if(r[0])byId[String(r[0])]=r;});backend.sort((a,b)=>Number(a[2]||999999)-Number(b[2]||999999)||Number(a[13]||999999)-Number(b[13]||999999));
  const out=backend.map(r=>{const dep=r[14]?byId[String(r[14])]:null;return[r[3]||'',r[10]||'',r[13]||'',dep?dep[13]||'':'',normalize_(r[6])==='COMPLETATA'?'FATTO':(r[6]||'DA FARE'),r[5]||''];});
  const sheet=child.getSheetByName(EVENT_SHEET.SHEETS.TASKS);if(!sheet)throw new Error('Foglio Attività non trovato nella Scheda evento.');
  const endRow=Math.min(TASKS_V4.ENTRY_END_ROW,sheet.getMaxRows());if(endRow>=2)sheet.getRange(2,1,endRow-1,TASKS_V4.VISIBLE_COLS).clearContent();if(out.length)sheet.getRange(2,1,out.length,TASKS_V4.VISIBLE_COLS).setValues(out);
  backend.forEach((r,index)=>{const depId=String(r[14]||'').trim(),status=normalize_(r[6]);if(!depId||status==='FATTO'||status==='COMPLETATA')return;sheet.getRange(index+2,5).setFormula(fastTaskStatusFormulaV9_(index+2));});
  if(!isEventSheetV8_(child))ensureNativeTaskTableV4_(child,sheet);
  return out.length;
}

function ensureFastEventTaskTriggerV9_(child){
  const sourceId=child.getId(),triggers=ScriptApp.getProjectTriggers();let exists=false;
  triggers.forEach(t=>{let same=false;try{same=t.getTriggerSourceId()===sourceId;}catch(e){}if(!same)return;const h=t.getHandlerFunction();if(['handleEventTaskEditV4','handleFastEventTaskEditV6'].includes(h)){ScriptApp.deleteTrigger(t);return;}if(h==='handleFastEventTaskEditV9')exists=true;});
  if(!exists)ScriptApp.newTrigger('handleFastEventTaskEditV9').forSpreadsheet(sourceId).onEdit().create();
}
function ensureFastEventTaskTriggerV6_(child){return ensureFastEventTaskTriggerV9_(child);}

function fastTaskStatusFormulaV9_(row){return'=IF(A'+row+'="";"";IF(D'+row+'="";"DA FARE";IFERROR(IF(AND(INDEX($E$2:$E$500;MATCH(D'+row+';$C$2:$C$500;0))="FATTO";$F'+row+'<>"";$F'+row+'<=TODAY());"DA FARE";"IN ATTESA");"IN ATTESA")))';}
function fastTaskStatusFormulaV7_(row){return fastTaskStatusFormulaV9_(row);}

function handleFastEventTaskEditV9(e){
  if(!e||!e.range)return;const sheet=e.range.getSheet();if(sheet.getName()!==EVENT_SHEET.SHEETS.TASKS)return;
  const firstRow=Math.max(2,e.range.getRow()),lastRow=Math.min(TASKS_V4.ENTRY_END_ROW,e.range.getLastRow()),firstCol=e.range.getColumn(),lastCol=e.range.getLastColumn();if(lastRow<firstRow||firstCol>TASKS_V4.VISIBLE_COLS)return;
  if(firstCol<=1&&lastCol>=1){const nums=sheet.getRange(2,3,Math.max(1,Math.min(TASKS_V4.ENTRY_END_ROW,sheet.getMaxRows())-1),1).getValues();let next=nums.reduce((m,r)=>{const n=Number(r[0]);return Number.isFinite(n)&&n>m?n:m;},0)+1;for(let row=firstRow;row<=lastRow;row++){const d=String(sheet.getRange(row,1).getDisplayValue()||'').trim();if(!d)continue;if(!(Number(sheet.getRange(row,3).getValue())>0))sheet.getRange(row,3).setValue(next++);autoLinkConfirmationTaskFastV9_(sheet,row);applyFastTaskDependencyV9_(sheet,row);}}
  if(firstCol<=4&&lastCol>=4)for(let row=firstRow;row<=lastRow;row++)applyFastTaskDependencyV9_(sheet,row);
  if(firstCol<=5&&lastCol>=5){for(let row=firstRow;row<=lastRow;row++)if(normalizeTaskTextFastV7_(sheet.getRange(row,5).getDisplayValue())==='FATTO')scheduleFastDependentTasksV9_(sheet,row);refreshFastTaskDependenciesV9_(sheet);}
  if(firstCol<=6&&lastCol>=6)refreshFastTaskDependenciesV9_(sheet);
}
function handleFastEventTaskEditV6(e){return handleFastEventTaskEditV9(e);}
function normalizeTaskTextFastV7_(value){return String(value===null||value===undefined?'':value).trim().toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ');}
function fastDateReachedV9_(value){if(!(value instanceof Date))return false;const d=new Date(value),t=new Date();d.setHours(0,0,0,0);t.setHours(0,0,0,0);return d.getTime()<=t.getTime();}

function autoLinkConfirmationTaskFastV9_(sheet,row){
  if(Number(sheet.getRange(row,4).getValue()||0)>0)return;const desc=normalizeTaskTextFastV7_(sheet.getRange(row,1).getDisplayValue());
  const parents={'CONFERMA GOMMONE':'RICHIESTA GOMMONE','CONFERMA SOGGIORNO':'RICHIESTA SOGGIORNO','CONFERMA PASTI':'SOLUZIONE PASTI','CONFERMA PRESENZE TUTTI ATLETI':'CONVOCAZIONE ATLETI','CONFERMA OSPITALITA CIRCOLO':'OSPITALITA CIRCOLO'};const parent=parents[desc]||'';if(!parent)return;
  const values=sheet.getRange(2,1,Math.max(1,row-2),3).getDisplayValues();for(let i=values.length-1;i>=0;i--)if(normalizeTaskTextFastV7_(values[i][0])===parent&&Number(values[i][2]||0)>0){sheet.getRange(row,4).setValue(Number(values[i][2]));return;}
}
function autoLinkConfirmationTaskFastV7_(sheet,row){return autoLinkConfirmationTaskFastV9_(sheet,row);}

function scheduleFastDependentTasksV9_(sheet,parentRow){
  const parentNo=Number(sheet.getRange(parentRow,3).getValue()||0);if(!(parentNo>0))return;const last=Math.min(500,Math.max(sheet.getLastRow(),2)),rows=sheet.getRange(2,1,last-1,6).getValues();
  rows.forEach((r,i)=>{if(Number(r[3]||0)!==parentNo)return;const row=i+2,desc=normalizeTaskTextFastV7_(r[0]),dueCell=sheet.getRange(row,6),existing=dueCell.getValue();if(desc!=='CONFERMA PRESENZE TUTTI ATLETI'&&!(existing instanceof Date)){const due=new Date();due.setHours(12,0,0,0);due.setDate(due.getDate()+2);dueCell.setValue(due).setNumberFormat('dd/MM/yyyy');}if(normalize_(sheet.getRange(row,5).getDisplayValue())!=='FATTO')sheet.getRange(row,5).setFormula(fastTaskStatusFormulaV9_(row));});
}
function applyFastTaskDependencyV9_(sheet,row){const description=String(sheet.getRange(row,1).getDisplayValue()||'').trim();if(!description)return;const status=normalize_(sheet.getRange(row,5).getDisplayValue());if(status==='FATTO')return;const depNo=Number(sheet.getRange(row,4).getValue()||0);if(depNo>0)sheet.getRange(row,5).setFormula(fastTaskStatusFormulaV9_(row));else if(!status||status==='IN ATTESA')sheet.getRange(row,5).setValue('DA FARE');}
function applyFastTaskDependencyV7_(sheet,row){return applyFastTaskDependencyV9_(sheet,row);}
function refreshFastTaskDependenciesV9_(sheet){const last=Math.min(500,Math.max(sheet.getLastRow(),2));for(let row=2;row<=last;row++){const description=String(sheet.getRange(row,1).getDisplayValue()||'').trim();if(!description)continue;const depNo=Number(sheet.getRange(row,4).getValue()||0),status=normalize_(sheet.getRange(row,5).getDisplayValue());if(depNo>0&&status!=='FATTO')sheet.getRange(row,5).setFormula(fastTaskStatusFormulaV9_(row));}}
function refreshFastTaskDependenciesV7_(sheet){return refreshFastTaskDependenciesV9_(sheet);}

function getLinkedEventSheetForSaveV2_(event){if(!event||!event._row)throw new Error('Seleziona una riga evento valida nel Calendario.');const cal=sh_(APP.SHEETS.CALENDAR),map=headerMap_(cal),col=map[APP.CALENDAR_HEADERS.EVENT_SHEET];if(!col)throw new Error('Colonna SCHEDA EVENTO non trovata nel Calendario.');const cell=cal.getRange(event._row,col),rich=cell.getRichTextValue(),url=(rich&&rich.getLinkUrl())||cell.getDisplayValue(),id=extractDriveId_(url);if(!id)throw new Error('Questa riga non ha una Scheda evento collegata. Usa prima “Crea / aggiorna scheda ← Calendario”.');try{return SpreadsheetApp.openById(id);}catch(e){throw new Error('Non riesco ad aprire la Scheda evento collegata. Verifica il link nella colonna SCHEDA EVENTO.');}}
function getExplicitLinkedEventSheetV3_(event){if(!event||!event._row)return null;const cal=sh_(APP.SHEETS.CALENDAR),map=headerMap_(cal),col=map[APP.CALENDAR_HEADERS.EVENT_SHEET];if(!col)return null;const cell=cal.getRange(event._row,col),rich=cell.getRichTextValue(),url=(rich&&rich.getLinkUrl())||cell.getDisplayValue(),id=extractDriveId_(url);if(!id)return null;try{return SpreadsheetApp.openById(id);}catch(e){return null;}}

function syncSelectedEventSheetToCalendarV2(){const ui=SpreadsheetApp.getUi(),event=selectedEvent_(),eventId=ensureEventId_(event),child=getLinkedEventSheetForSaveV2_(event);validateEventSheetIdentity_(child,eventId);if(isEventSheetV7_(child)){ui.alert('Salvataggio dalla Scheda evento','Questa è una Scheda evento di nuova generazione. Per salvare usa il menu “Scheda evento” → “💾 Salva dati nel Calendario” direttamente nel file evento.\n\nIn questo modo vengono gestiti correttamente preventivi, pagamenti multipli, residui e task automatiche.',ui.ButtonSet.OK);return null;}ensureChecklistBackendHeadersV3_();ensureParticipantsBackendHeadersV2_();const participantCount=syncParticipantsV3ToBackend_(eventId,event,child),taskCount=syncTasksV4ToBackend_(eventId,event,child),expenseCount=syncExpensesV4ToBackend_(eventId,event,child);ensureFastEventTaskTriggerV9_(child);SpreadsheetApp.flush();const counts={tasks:taskCount,participants:participantCount,expenses:expenseCount};ui.alert('Calendario aggiornato','I dati presenti nella Scheda evento sono stati salvati nel Calendario.\n\nAttività: '+counts.tasks+'\nSpese: '+counts.expenses+'\nPartecipanti: '+counts.participants,ui.ButtonSet.OK);return counts;}

function canonicalEventTypeV7_(value){const t=normalizeTaskTextFastV7_(value),aliases={'RIUNIONE/MISSIONE':'RIUNIONI/MISSIONI','RIUNIONI / MISSIONI':'RIUNIONI/MISSIONI','OSSERVAZIONE REGATA':'OSS.REGATE ITA','ISCRIZIONE REGATA':'ISCRIZIONI REGATE'};return aliases[t]||t;}
function canonicalEventClassV7_(value){const c=normalizeTaskTextFastV7_(value);return c==='KITE'?'KITEFOIL':c;}
function resolveEventCommitmentV7_(event){const type=canonicalEventTypeV7_(event[APP.CALENDAR_HEADERS.TYPE]),cls=canonicalEventClassV7_(event[APP.CALENDAR_HEADERS.CLASS]),config=sh_('_CONFIG_IMPEGNI').getDataRange().getDisplayValues();let wildcard='';for(let i=1;i<config.length;i++){if(canonicalEventTypeV7_(config[i][0])!==type)continue;const cfgClass=canonicalEventClassV7_(config[i][1]),imp=String(config[i][2]||'').trim();if(!imp)continue;if(cfgClass===cls||(cfgClass==='TUTTE'&&cls))return imp;if(!cfgClass||cfgClass==='*')wildcard=imp;}return wildcard;}
function technicianDirectoryV7_(){return[['ZAGGIA','Leonardo','Zaggia'],['CRISI','Andrea','Crisi'],['RAVEGLIA','Matteo','Raveglia'],['CARICATO','Francesco','Caricato'],['PICCIAU','Gianluigi','Picciau'],['SENSINI','Alessandra','Sensini'],['NUICOLUCCI','Matteo','Nuicolucci'],['CAMBONI','Mattia','Camboni'],['CANGEMI','Antonino','Cangemi'],['LOPERFIDO','Daniel','Loperfido']];}
function resolveFullTechniciansV7_(raw){const text=normalizeTaskTextFastV7_(raw);if(!text)return[];return technicianDirectoryV7_().filter(x=>text.indexOf(x[0])>=0);}
function initializeEventSheetV7FromCalendar_(child,eventId,event,folderId){writeEventMetaV5_(child,eventId,event,folderId);const imp=resolveEventCommitmentV7_(event)||String(event[APP.CALENDAR_HEADERS.COMMITMENT]||'').trim();if(imp){const cal=sh_(APP.SHEETS.CALENDAR),map=headerMap_(cal),col=map[APP.CALENDAR_HEADERS.COMMITMENT];if(col&&event._row)cal.getRange(event._row,col).setValue(imp).setNumberFormat('@');const expenses=child.getSheetByName(EVENT_SHEET.SHEETS.EXPENSES);if(expenses){if(isEventSheetV8_(child))expenses.getRange('E9').setValue(imp).setNumberFormat('@');else expenses.getRange('N9').setValue(imp).setNumberFormat('@');}}populateTechniciansInEventSheetV7_(child,event);}
function populateTechniciansInEventSheetV7_(child,event){const people=resolveFullTechniciansV7_(event[APP.CALENDAR_HEADERS.TECHNICIANS]);if(!people.length)return;const sheet=child.getSheetByName(PARTICIPANTS_V3.SHEET);if(!sheet||normalize_(sheet.getRange('H2').getDisplayValue())!=='RUOLO')return;const existing=sheet.getRange(3,1,15,2).getDisplayValues(),existingNames=new Set(existing.map(r=>normalizeTaskTextFastV7_((r[0]||'')+' '+(r[1]||''))));people.forEach(x=>{const fullKey=normalizeTaskTextFastV7_(x[1]+' '+x[2]);if(existingNames.has(fullKey))return;let target=0;for(let i=0;i<existing.length;i++)if(!String(existing[i][0]||'').trim()&&!String(existing[i][1]||'').trim()){target=i+3;existing[i]=[x[1],x[2]];break;}if(!target)return;sheet.getRange(target,1).setValue(x[1]);sheet.getRange(target,2).setValue(x[2]);sheet.getRange(target,8).setValue('TECNICO');sheet.getRange(target,9).setValue('CONFERMATO');existingNames.add(fullKey);});}

function writeEventMetaV5_(child,eventId,event,folderId){const meta=child.getSheetByName(EVENT_SHEET.SHEETS.META),start=event[APP.CALENDAR_HEADERS.START],end=event[APP.CALENDAR_HEADERS.END],imp=resolveEventCommitmentV7_(event)||String(event[APP.CALENDAR_HEADERS.COMMITMENT]||'').trim(),technicians=resolveFullTechniciansV7_(event[APP.CALENDAR_HEADERS.TECHNICIANS]),fullTechnicians=technicians.length?technicians.map(x=>x[1]+' '+x[2]).join(', '):String(event[APP.CALENDAR_HEADERS.TECHNICIANS]||'');writeMeta_(meta,{EVENT_ID:eventId,MASTER_SPREADSHEET_ID:APP.SPREADSHEET_ID,EVENT_FOLDER_ID:folderId,EVENT_SHEET_ID:child.getId(),SYNC_VERSION:'10',EVENT_LABEL:buildEventSheetLabel_(event),EVENT_TYPE:String(event[APP.CALENDAR_HEADERS.TYPE]||''),EVENT_CLASS:String(event[APP.CALENDAR_HEADERS.CLASS]||''),EVENT_LOCATION:String(event[APP.CALENDAR_HEADERS.LOCATION]||''),EVENT_ZONE:String(event[APP.CALENDAR_HEADERS.ZONE]||''),EVENT_CLUB:String(event[APP.CALENDAR_HEADERS.CLUB]||''),EVENT_TECHNICIANS:fullTechnicians,EVENT_LODGING:String(event[APP.CALENDAR_HEADERS.LODGING]||''),EVENT_COMMITMENT:imp,EVENT_START:start instanceof Date?Utilities.formatDate(start,APP.TZ,'yyyy-MM-dd'):'',EVENT_END:end instanceof Date?Utilities.formatDate(end,APP.TZ,'yyyy-MM-dd'):'',LAST_SYNC:Utilities.formatDate(new Date(),APP.TZ,'dd/MM/yyyy HH:mm')});}
function getOrCreateEventSheetV3_(eventId,event,folderId){let child=getExplicitLinkedEventSheetV3_(event),created=false;if(!child){const folder=DriveApp.getFolderById(folderId),templateFile=DriveApp.getFileById(EVENT_SHEET_TEMPLATE_ID_V3),copy=templateFile.makeCopy(buildEventSheetName_(event),folder);child=SpreadsheetApp.openById(copy.getId());child.setSpreadsheetLocale('it_IT');child.setSpreadsheetTimeZone(APP.TZ);created=true;}ensureEventSheetBaseV3_(child,eventId,folderId,event);return{spreadsheet:child,created:created};}
function ensureEventSheetBaseV3_(child,eventId,folderId,event){child.setSpreadsheetLocale('it_IT');child.setSpreadsheetTimeZone(APP.TZ);let tasks=child.getSheetByName(EVENT_SHEET.SHEETS.TASKS);if(!tasks){const first=child.getSheets()[0];if(child.getSheets().length===1&&first.getLastRow()===0){first.setName(EVENT_SHEET.SHEETS.TASKS);tasks=first;}else tasks=child.insertSheet(EVENT_SHEET.SHEETS.TASKS,0);}if(!child.getSheetByName(EVENT_SHEET.SHEETS.EXPENSES))child.insertSheet(EVENT_SHEET.SHEETS.EXPENSES);if(!child.getSheetByName(PARTICIPANTS_V3.SHEET))child.insertSheet(PARTICIPANTS_V3.SHEET);let meta=child.getSheetByName(EVENT_SHEET.SHEETS.META);if(!meta)meta=child.insertSheet(EVENT_SHEET.SHEETS.META);writeEventMetaV5_(child,eventId,event,folderId);meta.hideSheet();}
function hideEventSheetTechnicalColumnsV3_(child){const tasks=child.getSheetByName(EVENT_SHEET.SHEETS.TASKS);if(tasks){tasks.showColumns(1,TASKS_V4.VISIBLE_COLS);if(tasks.getMaxColumns()>TASKS_V4.VISIBLE_COLS)tasks.hideColumns(TASKS_V4.VISIBLE_COLS+1,tasks.getMaxColumns()-TASKS_V4.VISIBLE_COLS);}const expenses=child.getSheetByName(EVENT_SHEET.SHEETS.EXPENSES);if(expenses){expenses.showColumns(1,EXPENSES_V4.VISIBLE_COLS);if(expenses.getMaxColumns()>EXPENSES_V4.VISIBLE_COLS)expenses.hideColumns(EXPENSES_V4.VISIBLE_COLS+1,expenses.getMaxColumns()-EXPENSES_V4.VISIBLE_COLS);}const participants=child.getSheetByName(PARTICIPANTS_V3.SHEET);if(participants){participants.showColumns(1,PARTICIPANTS_V3.VISIBLE_COLS);if(participants.getMaxColumns()>PARTICIPANTS_V3.VISIBLE_COLS)participants.hideColumns(PARTICIPANTS_V3.VISIBLE_COLS+1,participants.getMaxColumns()-PARTICIPANTS_V3.VISIBLE_COLS);}}