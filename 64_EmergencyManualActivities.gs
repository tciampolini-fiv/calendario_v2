// MODALITA' EMERGENZA - 14/09/2026
// Obiettivo: rendere il foglio Attivita completamente manuale per il lavoro
// sulle vecchie Schede Evento. Nessuna automazione deve riscrivere Attivita.

function activateManualActivitiesEmergency(){
  const removed=disableCalendarAutomationEmergency_();
  const event=selectedEvent_();
  const child=getLinkedEventSheet_(event);
  if(!child)throw new Error('La riga selezionata non ha una Scheda evento collegata.');
  const result=manualizeActivitySheetEmergency_(child);
  SpreadsheetApp.getUi().alert(
    'Modalita manuale Attivita',
    'Scheda sbloccata: '+child.getName()+'\n'+
    'Protezioni rimosse: '+result.protections+'\n'+
    'Gruppi righe rimossi: '+result.groups+'\n'+
    'Trigger automatici rimossi: '+removed+'\n\n'+
    'Da questo momento il foglio Attivita si modifica manualmente.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return result;
}

function activateManualActivitiesEmergencyAll(){
  const removed=disableCalendarAutomationEmergency_();
  const cal=sh_(APP.SHEETS.CALENDAR),map=headerMap_(cal),last=cal.getLastRow();
  const col=map[APP.CALENDAR_HEADERS.EVENT_SHEET];
  if(!col)throw new Error('Colonna SCHEDA EVENTO non trovata nel Calendario.');
  const started=Date.now(),seen={},errors=[];let done=0,skipped=0;

  for(let row=2;row<=last;row++){
    if(Date.now()-started>260000){errors.push('Interrotto per limite di tempo dopo '+done+' schede. Rieseguire il comando per completare le restanti.');break;}
    const cell=cal.getRange(row,col),rich=cell.getRichTextValue();
    const id=extractDriveId_((rich&&rich.getLinkUrl())||cell.getDisplayValue());
    if(!id||seen[id]){skipped++;continue;}
    seen[id]=true;
    try{
      const child=SpreadsheetApp.openById(id);
      manualizeActivitySheetEmergency_(child);
      done++;
    }catch(err){errors.push('Riga '+row+': '+(err.message||String(err)));}
  }

  SpreadsheetApp.getUi().alert(
    'Modalita manuale Attivita',
    'Schede sbloccate: '+done+'\nTrigger automatici rimossi: '+removed+
    (errors.length?'\nProblemi: '+errors.length+'\n'+errors.slice(0,8).join('\n'):''),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return{done:done,skipped:skipped,errors:errors,removedTriggers:removed};
}

function disableCalendarAutomationEmergency_(){
  let removed=0;
  ScriptApp.getProjectTriggers().forEach(t=>{
    const h=String(t.getHandlerFunction()||'');
    if(/calendarAutoSync|syncChangedEventSheets/i.test(h)){
      ScriptApp.deleteTrigger(t);removed++;
    }
  });
  return removed;
}

function manualizeActivitySheetEmergency_(child){
  if(!child)throw new Error('Scheda evento non disponibile.');
  const sheet=child.getSheetByName('Attività');
  if(!sheet)throw new Error('Foglio Attivita non trovato in '+child.getName()+'.');

  let protections=0,groups=0;
  [SpreadsheetApp.ProtectionType.SHEET,SpreadsheetApp.ProtectionType.RANGE].forEach(type=>{
    sheet.getProtections(type).forEach(p=>{
      try{p.remove();protections++;}catch(err){console.log('Protezione non rimossa: '+err.message);}
    });
  });

  // Nessun raggruppamento o ordine strutturale imposto dallo script.
  try{
    for(let r=1;r<=Math.min(sheet.getMaxRows(),600);r++){
      const g=sheet.getRowGroup(r,1);
      if(g){g.remove();groups++;}
    }
  }catch(err){console.log('Rimozione gruppi: '+err.message);}

  // Nessun dropdown/checkbox/convalida che possa impedire la scrittura manuale.
  const usedRows=Math.max(1,Math.min(sheet.getMaxRows(),Math.max(sheet.getLastRow(),50)));
  const usedCols=Math.max(1,Math.min(sheet.getMaxColumns(),19));
  sheet.getRange(1,1,usedRows,usedCols).clearDataValidations();

  // Elimina la formattazione automatica/condizionale, ma conserva i formati
  // numerici (es. date) e soprattutto tutti i contenuti inseriti dall'utente.
  try{sheet.setConditionalFormatRules([]);}catch(err){console.log('Formattazione condizionale: '+err.message);}
  const visible=sheet.getRange(1,1,usedRows,Math.min(7,sheet.getMaxColumns()));
  visible.setBackground(null).setFontWeight('normal').setFontStyle('normal');

  // Rimuove soltanto i vecchi controlli grafici "+"; non cancella task/obiettivi.
  const last=Math.min(sheet.getLastRow(),600);
  if(last>=2){
    const range=sheet.getRange(2,1,last-1,8),values=range.getValues();
    let changed=false;
    values.forEach(r=>{
      if(String(r[2]||'').trim()==='＋'||String(r[2]||'').trim()==='+'){r[2]='';changed=true;}
      const a=String(r[0]||'').trim();
      if(/^\s*[＋+]\s*AGGIUNGI OBIETTIVO/i.test(a)||normalize_(r[7])==='AGGIUNGI_OBIETTIVO'){
        for(let c=0;c<8;c++)r[c]='';changed=true;
      }
    });
    if(changed)range.setValues(values);
  }

  try{sheet.showRows(1,sheet.getMaxRows());}catch(err){}
  SpreadsheetApp.flush();
  return{protections:protections,groups:groups};
}
