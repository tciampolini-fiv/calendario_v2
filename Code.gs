const EVENT_APP = Object.freeze({
  SHEETS:{TASKS:'Attività',EXPENSES:'Spese',PARTICIPANTS:'Partecipanti',META:'_META'},
  MASTER_DOCUMENTS:'_DOCUMENTI',
  INPUT_ROW:9,
  LIST_START_ROW:13,
  LIST_END_ROW:500,
  PARTICIPANTS:{CONV_START:3,CONV_COUNT:15,AGG_START:20,AGG_COUNT:20},
  TEMPLATE_SPECS:[
    {key:'CONV_ATLETI',label:'Convocazione atleti',id:'1SpOMrpDwe8aTW9QAyby865WufnD40AsPx6t8zm6Yw3E'},
    {key:'CONV_TECNICO',label:'Convocazione tecnico',id:'1Xb61H5TQ0avz8n5_THn-BSd1Xd4YBtmOD3Lu2uwUuJw'},
    {key:'GOMMONE',label:'Richiesta gommone alla Zona',id:'1uvK4J8WrWekpMrkPTkpUvsdYTggBhwzCDJaJmyIhUko'},
    {key:'OSPITALITA',label:'Richiesta ospitalità circolo',id:'1NuheJqcwtLxp8muTdbnpcZc22Z8HJgA-jbdt_Sq6ruk'},
    {key:'RINGRAZIAMENTO',label:'Ringraziamento circolo',id:'11Fqn8U3Hs8MgJPH0tQTKDEozmWDftPtmELoLjgMFg7c'}
  ]
});

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();

  if (sheet.getName() === EVENT_APP.SHEETS.TASKS) {
    eventHandleTaskEditV7_(e);
    return;
  }

  if (sheet.getName() === EVENT_APP.SHEETS.EXPENSES) {
    if (e.range.getA1Notation() === 'P7' && e.value === 'TRUE') {
      try { eventImportExpense(); }
      catch (err) {
        sheet.getRange('P7').setValue(false);
        SpreadsheetApp.getActive().toast(err.message || String(err),'Spese',6);
      }
      return;
    }
    const row = e.range.getRow();
    if (row === EVENT_APP.INPUT_ROW || (row >= EVENT_APP.LIST_START_ROW && row <= EVENT_APP.LIST_END_ROW)) {
      if (e.range.getColumn() <= 16) {
        eventNormalizeExpenseRowV7_(sheet,row);
        if (row >= EVENT_APP.LIST_START_ROW) eventSyncExpenseTasksV7_();
        eventRefreshExpenseDashboard();
      }
    }
    return;
  }

  if (sheet.getName() === EVENT_APP.SHEETS.PARTICIPANTS) {
    const row = e.range.getRow();
    if ((row >= 3 && row <= 17) || (row >= 20 && row <= 39)) eventRefreshExpenseDashboard();
  }
}

function eventInitializeV7_() {
  const meta = eventMeta_();
  if (!String(meta.EVENT_ID||'').trim()) return false; // non inizializzare il MODELLO stesso
  eventApplyCommitmentV7_(meta);
  eventPopulateTechniciansV7_(meta);
  eventAutoLinkConfirmationTasksV7_();
  eventRefreshTaskDependencyStatesV7_();
  eventSyncExpenseTasksV7_();
  eventRefreshExpenseDashboard();
  return true;
}

function eventTaskStatusFormulaV7_(row) {
  return '=IF(A' + row + '="";"";IF(D' + row + '="";"DA FARE";IFERROR(IF(INDEX($E$2:$E$500;MATCH(D' + row + ';$C$2:$C$500;0))="FATTO";"DA FARE";"IN ATTESA");"IN ATTESA")))';
}

function eventHandleTaskEditV7_(e) {
  const sheet = e.range.getSheet();
  const firstRow = Math.max(2,e.range.getRow());
  const lastRow = Math.min(500,e.range.getLastRow());
  const firstCol = e.range.getColumn();
  const lastCol = e.range.getLastColumn();
  if (lastRow < firstRow || firstCol > 6) return;

  if (firstCol <= 1 && lastCol >= 1) {
    for (let row=firstRow;row<=lastRow;row++) eventPrepareTaskRowV7_(sheet,row);
  }
  if (firstCol <= 4 && lastCol >= 4) {
    for (let row=firstRow;row<=lastRow;row++) eventApplyTaskDependencyV7_(sheet,row);
  }
  if (firstCol <= 5 && lastCol >= 5) eventRefreshTaskDependencyStatesV7_();
}

function eventPrepareTaskRowV7_(sheet,row) {
  const description = String(sheet.getRange(row,1).getDisplayValue()||'').trim();
  if (!description) return;
  const noCell = sheet.getRange(row,3);
  if (!(Number(noCell.getValue()) > 0)) {
    const nums = sheet.getRange(2,3,499,1).getValues().flat().map(Number).filter(n=>n>0);
    noCell.setValue(nums.length ? Math.max.apply(null,nums)+1 : 1);
  }
  eventAutoLinkConfirmationTaskRowV7_(sheet,row);
  eventApplyTaskDependencyV7_(sheet,row);
}

function eventApplyTaskDependencyV7_(sheet,row) {
  const description = String(sheet.getRange(row,1).getDisplayValue()||'').trim();
  if (!description) return;
  const status = String(sheet.getRange(row,5).getDisplayValue()||'').trim().toUpperCase();
  if (status === 'FATTO') return;
  const depNo = Number(sheet.getRange(row,4).getValue()||0);
  if (depNo > 0) sheet.getRange(row,5).setFormula(eventTaskStatusFormulaV7_(row));
  else if (!status || status === 'IN ATTESA') sheet.getRange(row,5).setValue('DA FARE');
}

function eventRefreshTaskDependencyStatesV7_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(EVENT_APP.SHEETS.TASKS);
  if (!sheet) return;
  for (let row=2;row<=Math.min(500,Math.max(sheet.getLastRow(),2));row++) {
    const description = String(sheet.getRange(row,1).getDisplayValue()||'').trim();
    if (!description) continue;
    const depNo = Number(sheet.getRange(row,4).getValue()||0);
    const status = String(sheet.getRange(row,5).getDisplayValue()||'').trim().toUpperCase();
    if (depNo > 0 && status !== 'FATTO') sheet.getRange(row,5).setFormula(eventTaskStatusFormulaV7_(row));
    else if (!depNo && !status) sheet.getRange(row,5).setValue('DA FARE');
  }
}

function eventAutoLinkConfirmationTasksV7_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(EVENT_APP.SHEETS.TASKS);
  if (!sheet) return;
  const last = Math.min(500,Math.max(sheet.getLastRow(),2));
  for (let row=2;row<=last;row++) eventAutoLinkConfirmationTaskRowV7_(sheet,row);
}

function eventAutoLinkConfirmationTaskRowV7_(sheet,row) {
  if (Number(sheet.getRange(row,4).getValue()||0) > 0) return;
  const desc = eventNormalizeTextV7_(sheet.getRange(row,1).getDisplayValue());
  if (!desc) return;
  let parent = '';
  if (desc === 'CHECK CONFERMA PRESENZE') parent = 'CONVOCAZIONE ATLETI';
  else if (desc === 'CONFERMA OSPITALITA') parent = 'OSPITALITA CIRCOLO';
  else if (desc.indexOf('CONFERMA ') === 0) parent = desc.substring(9).trim();
  if (!parent) return;

  const values = sheet.getRange(2,1,Math.max(1,row-2),3).getDisplayValues();
  for (let i=values.length-1;i>=0;i--) {
    if (eventNormalizeTextV7_(values[i][0]) === parent && Number(values[i][2]||0)>0) {
      sheet.getRange(row,4).setValue(Number(values[i][2]));
      return;
    }
  }
}

function eventNormalizeTextV7_(value) {
  return String(value===null||value===undefined?'':value).trim().toUpperCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ');
}

function eventMeta_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(EVENT_APP.SHEETS.META);
  if (!sheet) throw new Error('Foglio _META non trovato.');
  const values = sheet.getRange(1,1,sheet.getLastRow(),2).getDisplayValues();
  const out = {};
  values.forEach(r=>{ if (r[0]) out[String(r[0]).trim()] = r[1]; });
  return out;
}

function eventWriteLocalMetaV7_(key,value) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(EVENT_APP.SHEETS.META);
  if (!sheet) return;
  const rows = sheet.getRange(1,1,Math.max(sheet.getLastRow(),1),2).getDisplayValues();
  const target = eventNormalizeTextV7_(key);
  for (let i=0;i<rows.length;i++) {
    if (eventNormalizeTextV7_(rows[i][0])===target) { sheet.getRange(i+1,2).setValue(value); return; }
  }
  sheet.appendRow([key,value]);
}

function eventCommitmentMapV7_() {
  return {
    'RIUNIONI/MISSIONI|*':'2-1','ALLENAMENTO|29ER':'28','ALLENAMENTO|420':'29','ALLENAMENTO|NACRA 15':'30',
    'ALLENAMENTO|IQFOIL':'31','ALLENAMENTO|ILCA 6':'32','ALLENAMENTO|ILCA 4':'33-1','ALLENAMENTO INTERZONALE|ILCA 4':'33-2',
    'ALLENAMENTO|KITEFOIL':'34','ALLENAMENTO|WING':'35','TEST FISICI|*':'37','COLLEGIALE|*':'38',
    'REGATA|29ER':'39-1','REGATA|420':'40','REGATA|NACRA 15':'41','REGATA|IQFOIL':'42','REGATA|ILCA 6':'43',
    'REGATA|ILCA 4':'44','REGATA|KITEFOIL':'45','REGATA|WING':'46','OSS.REGATE ITA|*':'36',
    'CAMP. NAZ GIOVANILE SINGOLO|*':'48-1','CAMP. NAZ GIOVANILE DOPPIO|*':'48-1','FOIL ACADEMY|*':'50',
    'ISCRIZIONI REGATE|29ER':'47-1','ISCRIZIONI REGATE|420':'47-2','ISCRIZIONI REGATE|NACRA 15':'47-3',
    'ISCRIZIONI REGATE|IQFOIL':'47-4','ISCRIZIONI REGATE|ILCA 6':'47-5','ISCRIZIONI REGATE|ILCA 4':'47-6',
    'ISCRIZIONI REGATE|KITEFOIL':'47-7','ISCRIZIONI REGATE|WING':'47-8','ISCRIZIONE YWC|TUTTE':'171',
    'MANUTENZIONE ALLENAMENTI|*':'39-3','MANUTENZIONE REGATE|*':'39-4','NOLEGGI ALLENAMENTI|*':'52','NOLEGGI REGATE|*':'51'
  };
}

function eventCanonicalTypeV7_(value) {
  const t = eventNormalizeTextV7_(value);
  const aliases = {'RIUNIONE/MISSIONE':'RIUNIONI/MISSIONI','RIUNIONI / MISSIONI':'RIUNIONI/MISSIONI','OSSERVAZIONE REGATA':'OSS.REGATE ITA','ISCRIZIONE REGATA':'ISCRIZIONI REGATE'};
  return aliases[t] || t;
}
function eventCanonicalClassV7_(value) {
  const c = eventNormalizeTextV7_(value);
  if (c==='KITE') return 'KITEFOIL';
  return c;
}
function eventResolveCommitmentV7_(meta) {
  const type = eventCanonicalTypeV7_(meta.EVENT_TYPE||'');
  const cls = eventCanonicalClassV7_(meta.EVENT_CLASS||'');
  const map = eventCommitmentMapV7_();
  return map[type+'|'+cls] || map[type+'|*'] || map[type+'|TUTTE'] || '';
}
function eventApplyCommitmentV7_(meta) {
  const imp = eventResolveCommitmentV7_(meta) || String(meta.EVENT_COMMITMENT||'').trim();
  if (!imp) return '';
  eventWriteLocalMetaV7_('EVENT_COMMITMENT',imp);
  const expense = SpreadsheetApp.getActive().getSheetByName(EVENT_APP.SHEETS.EXPENSES);
  if (expense) {
    expense.getRange('N9').setValue(imp).setNumberFormat('@');
    const vals = expense.getRange(13,2,488,1).getDisplayValues();
    vals.forEach((r,i)=>{ if (String(r[0]||'').trim() && !String(expense.getRange(i+13,14).getDisplayValue()||'').trim()) expense.getRange(i+13,14).setValue(imp).setNumberFormat('@'); });
  }
  return imp;
}

function eventTechnicianDirectoryV7_() {
  return [
    ['ZAGGIA','Leonardo','Zaggia'],['CRISI','Andrea','Crisi'],['RAVEGLIA','Matteo','Raveglia'],['CARICATO','Francesco','Caricato'],
    ['PICCIAU','Gianluigi','Picciau'],['SENSINI','Alessandra','Sensini'],['NUICOLUCCI','Matteo','Nuicolucci'],['CAMBONI','Mattia','Camboni'],
    ['CANGEMI','Antonino','Cangemi'],['LOPERFIDO','Daniel','Loperfido']
  ];
}
function eventPopulateTechniciansV7_(meta) {
  const raw = eventNormalizeTextV7_(meta.EVENT_TECHNICIANS||'');
  if (!raw) return [];
  const found = eventTechnicianDirectoryV7_().filter(x=>raw.indexOf(x[0])>=0);
  if (!found.length) return [];
  const full = found.map(x=>x[1]+' '+x[2]);
  eventWriteLocalMetaV7_('EVENT_TECHNICIANS',full.join(', '));

  const sheet = SpreadsheetApp.getActive().getSheetByName(EVENT_APP.SHEETS.PARTICIPANTS);
  if (!sheet) return full;
  const rows = sheet.getRange(3,1,15,19).getDisplayValues();
  const existing = new Set(rows.map(r=>eventNormalizeTextV7_((r[0]||'')+' '+(r[1]||''))));
  found.forEach(x=>{
    const key = eventNormalizeTextV7_(x[1]+' '+x[2]);
    if (existing.has(key)) return;
    let target = -1;
    for (let i=0;i<rows.length;i++) if (!String(rows[i][0]||'').trim() && !String(rows[i][1]||'').trim()) { target=i+3; rows[i][0]=x[1]; rows[i][1]=x[2]; break; }
    if (target<0) return;
    sheet.getRange(target,1).setValue(x[1]);
    sheet.getRange(target,2).setValue(x[2]);
    sheet.getRange(target,8).setValue('TECNICO');
    sheet.getRange(target,9).setValue('CONFERMATO');
    existing.add(key);
  });
  return full;
}

function eventPaymentTypeV7_(value) {
  const t = eventNormalizeTextV7_(value);
  if (t==='FATTURA' || t==='PAGAMENTO FATTURA') return 'SALDO FATTURA';
  if (t==='CARTA' || t==='CC') return 'CARTA DI CREDITO';
  return t;
}
function eventIsFullPaymentV7_(type) { return type==='CARTA DI CREDITO' || type==='SALDO FATTURA'; }
function eventNormalizeExpenseRowV7_(sheet,row) {
  if (!sheet || (row!==9 && (row<13 || row>500))) return;
  const r = sheet.getRange(row,1,1,16).getValues()[0];
  const budget = Number(r[0]||0);
  if (!(budget>0) && !String(r[1]||'').trim()) { sheet.getRange(row,11).clearContent(); return; }
  let p1 = eventPaymentTypeV7_(r[4]);
  let a1 = Number(r[5]||0);
  let d1 = r[6] instanceof Date ? r[6] : '';
  let p2 = eventPaymentTypeV7_(r[7]);
  let a2 = Number(r[8]||0);
  let d2 = r[9] instanceof Date ? r[9] : '';
  const today = new Date(); today.setHours(12,0,0,0);
  if (p1 && eventIsFullPaymentV7_(p1) && !(a1>0) && budget>0) a1=budget;
  if (p1 && a1>0 && !d1) d1=today;
  if (p2 && eventIsFullPaymentV7_(p2) && !(a2>0) && budget>0) a2=Math.max(budget-a1,0);
  if (p2 && a2>0 && !d2) d2=today;
  if (p1 && p1!==String(r[4]||'')) sheet.getRange(row,5).setValue(p1);
  if (a1!==Number(r[5]||0)) sheet.getRange(row,6).setValue(a1);
  if (d1 && !(r[6] instanceof Date)) sheet.getRange(row,7).setValue(d1).setNumberFormat('dd/MM/yyyy');
  if (p2 && p2!==String(r[7]||'')) sheet.getRange(row,8).setValue(p2);
  if (a2!==Number(r[8]||0)) sheet.getRange(row,9).setValue(a2);
  if (d2 && !(r[9] instanceof Date)) sheet.getRange(row,10).setValue(d2).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(row,11).setValue(Math.max(budget-a1-a2,0)).setNumberFormat('€ #,##0.00');
  if (!String(r[13]||'').trim()) {
    const imp = eventResolveCommitmentV7_(eventMeta_());
    if (imp) sheet.getRange(row,14).setValue(imp).setNumberFormat('@');
  }
}

function eventImportExpense() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(EVENT_APP.SHEETS.EXPENSES);
  if (!sheet) throw new Error('Foglio Spese non trovato.');
  eventNormalizeExpenseRowV7_(sheet,9);
  const input = sheet.getRange(9,1,1,16).getValues()[0];
  const budget = Number(input[0]||0);
  const description = String(input[1]||'').trim();
  if (!(budget>0)) throw new Error('Inserisci il PREVENTIVO totale.');
  if (!description) throw new Error('La DESCRIZIONE è obbligatoria.');

  let target=13;
  const descriptions=sheet.getRange(13,2,488,1).getDisplayValues();
  while(target<=500 && String(descriptions[target-13][0]||'').trim()) target++;
  if (target>500) throw new Error('Elenco spese pieno.');

  const now=new Date();
  sheet.getRange(target,1,1,16).setValues([input]);
  sheet.getRange(target,17,1,6).setValues([['SPESA-'+Utilities.getUuid(),'','',now,now,'']]);
  sheet.getRange(target,1).setNumberFormat('€ #,##0.00');
  sheet.getRange(target,6).setNumberFormat('€ #,##0.00');
  sheet.getRange(target,9).setNumberFormat('€ #,##0.00');
  sheet.getRange(target,11).setNumberFormat('€ #,##0.00');
  sheet.getRange(9,1,1,16).clearContent();
  const imp=eventResolveCommitmentV7_(eventMeta_());
  if (imp) sheet.getRange('N9').setValue(imp).setNumberFormat('@');
  sheet.getRange('P7').setValue(false);
  eventNormalizeExpenseRowV7_(sheet,target);
  eventSyncExpenseTasksV7_();
  eventRefreshExpenseDashboard();
  ss.toast('Spesa aggiunta. Il residuo e le scadenze sono stati aggiornati.','Spese',4);
}

function eventRefreshExpenseDashboard() {
  const ss=SpreadsheetApp.getActive();
  const expense=ss.getSheetByName(EVENT_APP.SHEETS.EXPENSES);
  const participants=ss.getSheetByName(EVENT_APP.SHEETS.PARTICIPANTS);
  if (!expense || String(expense.getRange('A1').getDisplayValue())!=='RIEPILOGO SPESE EVENTO') return;
  const rows=expense.getRange(13,1,488,16).getValues();
  const buckets={VIAGGI:0,VITTO:0,ALLOGGIO:0,NOLEGGI:0,RIMBORSI:0,ALTRO:0};
  let forecast=0,paid=0;
  rows.forEach(r=>{
    const budget=Number(r[0]||0); if(!budget && !String(r[1]||'').trim()) return;
    const category=eventNormalizeTextV7_(r[2]||'ALTRO');
    const bucket=['VIAGGI','VITTO','ALLOGGIO','NOLEGGI'].includes(category)?category:'ALTRO';
    forecast+=budget; buckets[bucket]+=budget; paid+=Number(r[5]||0)+Number(r[8]||0);
  });
  if (participants) {
    const pRows=participants.getRange(3,1,15,19).getValues();
    pRows.forEach(r=>{
      const name=String(r[0]||'').trim(),surname=String(r[1]||'').trim();
      if(!name&&!surname) return;
      const role=eventNormalizeTextV7_(r[7]); if(role==='TECNICO') return;
      const status=eventNormalizeTextV7_(r[8]);
      const max=Number(r[9]||0),passed=Number(r[10]||0);
      const current=status==='ASSENTE'?0:(passed>0?passed:max);
      forecast+=current; buckets.RIMBORSI+=current; if(passed>0) paid+=passed;
    });
  }
  expense.getRange('C2').setValue('PREVENTIVO EVENTO'); expense.getRange('D2').setValue(forecast).setNumberFormat('€ #,##0.00');
  expense.getRange('E2').setValue('PAGATO'); expense.getRange('F2').setValue(paid).setNumberFormat('€ #,##0.00');
  expense.getRange('G2').setValue('DA SALDARE'); expense.getRange('H2').setValue(Math.max(forecast-paid,0)).setNumberFormat('€ #,##0.00');
  expense.getRange('A5:F5').setValues([[buckets.VIAGGI,buckets.VITTO,buckets.ALLOGGIO,buckets.NOLEGGI,buckets.RIMBORSI,buckets.ALTRO]]).setNumberFormat('€ #,##0.00');
}

function eventSyncExpenseTasksV7_() {
  const ss=SpreadsheetApp.getActive();
  const expense=ss.getSheetByName(EVENT_APP.SHEETS.EXPENSES);
  const tasks=ss.getSheetByName(EVENT_APP.SHEETS.TASKS);
  if(!expense||!tasks) return;
  const rows=expense.getRange(13,1,488,16).getValues();
  let latestAfor=null;
  rows.forEach(r=>{
    const description=String(r[1]||'').trim(); if(!description) return;
    const p1=eventPaymentTypeV7_(r[4]),p2=eventPaymentTypeV7_(r[7]);
    if(p1==='AFOR' && r[6] instanceof Date && (!latestAfor||r[6]>latestAfor)) latestAfor=r[6];
    if(p2==='AFOR' && r[9] instanceof Date && (!latestAfor||r[9]>latestAfor)) latestAfor=r[9];
    const residual=Math.max(Number(r[0]||0)-Number(r[5]||0)-Number(r[8]||0),0);
    const due=r[11] instanceof Date?r[11]:null;
    const taskDesc='Saldo '+description;
    const taskRow=eventFindTaskRowV7_(tasks,taskDesc);
    if(residual>0 && due) eventEnsureTaskV7_(tasks,taskDesc,due,'DA FARE');
    else if(residual<=0 && taskRow>0 && eventNormalizeTextV7_(tasks.getRange(taskRow,5).getDisplayValue())!=='FATTO') tasks.getRange(taskRow,5).setValue('FATTO');
  });
  if(latestAfor) {
    const due=new Date(latestAfor); due.setDate(due.getDate()+1);
    eventEnsureTaskV7_(tasks,'Inviare contabile',due,'DA FARE');
    const end=eventParseIsoDate_(eventMeta_().EVENT_END);
    if(end) eventEnsureTaskV7_(tasks,'Richiedere contabile',end,new Date()>=end?'DA FARE':'IN ATTESA');
  }
}
function eventFindTaskRowV7_(sheet,description) {
  const target=eventNormalizeTextV7_(description); const vals=sheet.getRange(2,1,499,1).getDisplayValues();
  for(let i=0;i<vals.length;i++) if(eventNormalizeTextV7_(vals[i][0])===target) return i+2;
  return 0;
}
function eventEnsureTaskV7_(sheet,description,due,status) {
  let row=eventFindTaskRowV7_(sheet,description);
  if(!row) {
    const vals=sheet.getRange(2,1,499,1).getDisplayValues(); row=vals.findIndex(r=>!String(r[0]||'').trim())+2;
    if(row<2) return 0;
    sheet.getRange(row,1).setValue(description); eventPrepareTaskRowV7_(sheet,row);
  }
  if(due instanceof Date) sheet.getRange(row,6).setValue(due).setNumberFormat('dd/MM/yyyy');
  const current=eventNormalizeTextV7_(sheet.getRange(row,5).getDisplayValue());
  if(current!=='FATTO') sheet.getRange(row,5).setValue(status||'DA FARE');
  return row;
}

function eventGenerateDocuments() {
  const ui=SpreadsheetApp.getUi(); const meta=eventMeta_();
  let prompt='Scegli i documenti da creare, separando i numeri con una virgola:\n\n';
  EVENT_APP.TEMPLATE_SPECS.forEach((t,i)=>prompt+=(i+1)+' = '+t.label+'\n');
  const response=ui.prompt('Genera documenti',prompt+'\nEsempio: 1,2,4',ui.ButtonSet.OK_CANCEL);
  if(response.getSelectedButton()!==ui.Button.OK) return;
  const keys=eventParseDocumentSelection_(response.getResponseText());
  if(!keys.length){ui.alert('Nessun documento valido selezionato.');return;}
  const folder=DriveApp.getFolderById(meta.EVENT_FOLDER_ID); const replacements=eventDocumentReplacements_(meta); const created=[];
  keys.forEach(key=>{const spec=EVENT_APP.TEMPLATE_SPECS.find(x=>x.key===key);if(!spec)return;const copy=DriveApp.getFileById(spec.id).makeCopy(eventDocumentName_(key,replacements),folder);eventFillDocument_(copy.getId(),replacements,key);eventRegisterDocument_(meta,spec,copy);created.push(copy.getName());});
  ui.alert('Documenti creati',created.map(x=>'• '+x).join('\n'),ui.ButtonSet.OK);
}
function eventParseDocumentSelection_(text){const seen={};return String(text||'').split(/[;,\s]+/).map(x=>parseInt(x,10)).filter(n=>{if(!n||n<1||n>EVENT_APP.TEMPLATE_SPECS.length||seen[n])return false;seen[n]=true;return true;}).map(n=>EVENT_APP.TEMPLATE_SPECS[n-1].key);}
function eventDocumentReplacements_(meta){
  const participants=SpreadsheetApp.getActive().getSheetByName(EVENT_APP.SHEETS.PARTICIPANTS);
  const rows=participants?participants.getRange(3,1,15,19).getDisplayValues():[];
  const athletes=rows.filter(r=>{const has=String(r[0]||'').trim()||String(r[1]||'').trim();const role=eventNormalizeTextV7_(r[7]||'ATLETA');const status=eventNormalizeTextV7_(r[8]);return has&&role==='ATLETA'&&status!=='ASSENTE';});
  const athleteLines=athletes.map(r=>{const name=[r[0],r[1]].filter(Boolean).join(' ').trim();return r[2]?name+' – '+r[2]:name;});
  const start=eventParseIsoDate_(meta.EVENT_START),end=eventParseIsoDate_(meta.EVENT_END);
  return {'tipo':eventTitleCase_(meta.EVENT_TYPE||''),'classe':meta.EVENT_CLASS||'','luogo':meta.EVENT_LOCATION||'','localita':meta.EVENT_LOCATION||'','zona':meta.EVENT_ZONE||'','circolo':meta.EVENT_CLUB||'','tecnico':meta.EVENT_TECHNICIANS||'','lista tecnici':meta.EVENT_TECHNICIANS||'','lista atleti':athleteLines.join('\n'),'numero atleti':athletes.length?String(athletes.length):'','hotel/struttura':meta.EVENT_LODGING||'','data inizio':eventFormatDate_(start),'data fine':eventFormatDate_(end),'data in':eventFormatDate_(start),'data out':eventFormatDate_(end),'date':eventFormatDateRange_(start,end),'data di oggi':eventFormatDate_(new Date()),'data oggi':eventFormatDate_(new Date()),'impegno':meta.EVENT_COMMITMENT||''};
}
function eventFillDocument_(docId,replacements,key){const doc=DocumentApp.openById(docId),body=doc.getBody(),reps=Object.assign({},replacements);if(key==='GOMMONE')reps.N=replacements.zona||'';if(key==='OSPITALITA')reps.N=replacements['numero atleti']||'';Object.keys(reps).forEach(k=>{const v=reps[k];if(v===''||v===null||v===undefined)return;body.replaceText('(?i)\\{\\{'+eventEscapeRegex_(k)+'\\}\\}',eventSafeReplacement_(v));});if(reps.classe)body.replaceText('(?i)\\{classe\\}\\}',eventSafeReplacement_(reps.classe));doc.saveAndClose();}
function eventDocumentName_(key,r){const base={CONV_ATLETI:['Convocazione atleti',r.tipo,r.classe,r.luogo,r.date],CONV_TECNICO:['Convocazione tecnico',r.tecnico,r.tipo,r.classe,r.date],GOMMONE:['Richiesta gommone',r.zona?r.zona+' Zona':'',r.tipo,r.classe,r.date],OSPITALITA:['Richiesta ospitalità',r.circolo,r.tipo,r.classe,r.date],RINGRAZIAMENTO:['Ringraziamento',r.circolo,r.tipo,r.classe,r.date]}[key]||[r.tipo,r.classe,r.date];return base.filter(Boolean).join(' - ');}
function eventRegisterDocument_(meta,spec,file){if(!meta.MASTER_SPREADSHEET_ID)return;const master=SpreadsheetApp.openById(meta.MASTER_SPREADSHEET_ID),sheet=master.getSheetByName(EVENT_APP.MASTER_DOCUMENTS);if(!sheet)return;sheet.appendRow(['DOC-'+Utilities.getUuid(),meta.EVENT_ID,spec.key,spec.id,file.getName(),file.getId(),file.getUrl(),'BOZZA MODIFICABILE',new Date(),'']);}
function eventParseIsoDate_(s){const m=String(s||'').match(/^(\d{4})-(\d{2})-(\d{2})$/);return m?new Date(Number(m[1]),Number(m[2])-1,Number(m[3]),12,0,0):null;}
function eventFormatDate_(d){if(!(d instanceof Date)||isNaN(d))return'';const months=['gennaio','febbraio','marzo','aprile','maggio','giugno','luglio','agosto','settembre','ottobre','novembre','dicembre'];return d.getDate()+' '+months[d.getMonth()]+' '+d.getFullYear();}
function eventFormatDateRange_(s,e){if(!(s instanceof Date))return'';if(!(e instanceof Date))return eventFormatDate_(s);if(s.getTime()===e.getTime())return eventFormatDate_(s);const months=['gennaio','febbraio','marzo','aprile','maggio','giugno','luglio','agosto','settembre','ottobre','novembre','dicembre'];if(s.getFullYear()===e.getFullYear()&&s.getMonth()===e.getMonth())return s.getDate()+'-'+e.getDate()+' '+months[s.getMonth()]+' '+s.getFullYear();return eventFormatDate_(s)+' - '+eventFormatDate_(e);}
function eventTitleCase_(s){return String(s||'').toLowerCase().replace(/(^|\s|[-/])([a-zà-öø-ÿ])/g,(_,p,c)=>p+c.toUpperCase());}
function eventEscapeRegex_(s){return String(s||'').replace(/[.*+?^${}()|[\]\\]/g,'\\$&');}
function eventSafeReplacement_(s){return String(s||'').replace(/\\/g,'\\\\').replace(/\$/g,'\\$');}
