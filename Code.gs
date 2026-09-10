const EVENT_APP = Object.freeze({
  SHEETS:{TASKS:'Attività',EXPENSES:'Spese',PARTICIPANTS:'Partecipanti',META:'_META'},
  MASTER_DOCUMENTS:'_DOCUMENTI',
  INPUT_ROW:9,
  LIST_START_ROW:20,
  LIST_END_ROW:500,
  PARTICIPANTS:{CONV_START:3,CONV_COUNT:20,AGG_START:25,AGG_COUNT:20},
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
  const sheet=e.range.getSheet();
  if (sheet.getName()===EVENT_APP.SHEETS.TASKS) { eventHandleTaskEditV9_(e); return; }
  if (sheet.getName()===EVENT_APP.SHEETS.EXPENSES) { eventHandleExpenseEditV8_(e); return; }
  if (sheet.getName()===EVENT_APP.SHEETS.PARTICIPANTS) {
    const row=e.range.getRow();
    if ((row>=3&&row<=22)||(row>=25&&row<=44)) eventRefreshExpenseDashboardV8_();
  }
}

function eventInitializeV7_() {
  const meta=eventMeta_();
  if (!String(meta.EVENT_ID||'').trim()) return false;
  eventApplyCommitmentV8_(meta);
  eventPopulateTechniciansV7_(meta);
  eventAutoLinkConfirmationTasksV9_();
  eventRefreshTaskDependencyStatesV9_();
  eventSyncExpenseTasksV8_();
  eventRefreshExpenseDashboardV8_();
  return true;
}

function eventTaskStatusFormulaV9_(row) {
  return '=IF(A'+row+'="";"";IF(D'+row+'="";"DA FARE";IFERROR(IF(AND(INDEX($E$2:$E$500;MATCH(D'+row+';$C$2:$C$500;0))="FATTO";$F'+row+'<>"";$F'+row+'<=TODAY());"DA FARE";"IN ATTESA");"IN ATTESA")))';
}

function eventHandleTaskEditV9_(e) {
  const sheet=e.range.getSheet();
  const firstRow=Math.max(2,e.range.getRow());
  const lastRow=Math.min(500,e.range.getLastRow());
  const firstCol=e.range.getColumn();
  const lastCol=e.range.getLastColumn();
  if (lastRow<firstRow||firstCol>6) return;

  if (firstCol<=1&&lastCol>=1) {
    for(let row=firstRow;row<=lastRow;row++) eventPrepareTaskRowV9_(sheet,row);
  }
  if (firstCol<=4&&lastCol>=4) {
    for(let row=firstRow;row<=lastRow;row++) eventApplyTaskDependencyV9_(sheet,row);
  }
  if (firstCol<=5&&lastCol>=5) {
    for(let row=firstRow;row<=lastRow;row++) {
      if (eventNormalizeTextV7_(sheet.getRange(row,5).getDisplayValue())==='FATTO') {
        eventScheduleDependentTasksV9_(sheet,row);
      }
    }
    eventRefreshTaskDependencyStatesV9_();
  }
  if (firstCol<=6&&lastCol>=6) eventRefreshTaskDependencyStatesV9_();
}

function eventPrepareTaskRowV9_(sheet,row) {
  const description=String(sheet.getRange(row,1).getDisplayValue()||'').trim();
  if(!description)return;
  const noCell=sheet.getRange(row,3);
  if (!(Number(noCell.getValue())>0)) {
    const nums=sheet.getRange(2,3,499,1).getValues().flat().map(Number).filter(n=>n>0);
    noCell.setValue(nums.length?Math.max.apply(null,nums)+1:1);
  }
  eventAutoLinkConfirmationTaskRowV9_(sheet,row);
  eventApplyTaskDependencyV9_(sheet,row);
}

function eventApplyTaskDependencyV9_(sheet,row) {
  const description=String(sheet.getRange(row,1).getDisplayValue()||'').trim();
  if(!description)return;
  const status=eventNormalizeTextV7_(sheet.getRange(row,5).getDisplayValue());
  if(status==='FATTO')return;
  const depNo=Number(sheet.getRange(row,4).getValue()||0);
  if(depNo>0) sheet.getRange(row,5).setFormula(eventTaskStatusFormulaV9_(row));
  else if(!status||status==='IN ATTESA') sheet.getRange(row,5).setValue('DA FARE');
}

function eventRefreshTaskDependencyStatesV9_() {
  const sheet=SpreadsheetApp.getActive().getSheetByName(EVENT_APP.SHEETS.TASKS);
  if(!sheet)return;
  const last=Math.min(500,Math.max(sheet.getLastRow(),2));
  const values=sheet.getRange(2,1,last-1,6).getValues();
  const rowByNo={};
  values.forEach((r,i)=>{const n=Number(r[2]||0);if(n>0)rowByNo[n]=i+2;});

  for(let row=2;row<=last;row++) {
    const description=String(sheet.getRange(row,1).getDisplayValue()||'').trim();
    if(!description)continue;
    const status=eventNormalizeTextV7_(sheet.getRange(row,5).getDisplayValue());
    if(status==='FATTO')continue;
    const depNo=Number(sheet.getRange(row,4).getValue()||0);
    if(depNo>0) {
      const parentRow=rowByNo[depNo]||0;
      const parentDone=parentRow>0&&eventNormalizeTextV7_(sheet.getRange(parentRow,5).getDisplayValue())==='FATTO';
      const due=sheet.getRange(row,6).getValue();
      const dueReached=due instanceof Date&&eventDateReachedV9_(due);
      if(parentDone&&dueReached) sheet.getRange(row,5).setValue('DA FARE');
      else sheet.getRange(row,5).setFormula(eventTaskStatusFormulaV9_(row));
    } else if(!status) {
      sheet.getRange(row,5).setValue('DA FARE');
    }
  }
}

function eventDateReachedV9_(date) {
  if(!(date instanceof Date))return false;
  const due=new Date(date); due.setHours(0,0,0,0);
  const today=new Date(); today.setHours(0,0,0,0);
  return due.getTime()<=today.getTime();
}

function eventScheduleDependentTasksV9_(sheet,parentRow) {
  const parentNo=Number(sheet.getRange(parentRow,3).getValue()||0);
  if(!(parentNo>0))return;
  const last=Math.min(500,Math.max(sheet.getLastRow(),2));
  const rows=sheet.getRange(2,1,last-1,6).getValues();
  rows.forEach((r,i)=>{
    const row=i+2;
    if(Number(r[3]||0)!==parentNo)return;
    const description=eventNormalizeTextV7_(r[0]);
    const dueCell=sheet.getRange(row,6);
    const existingDue=dueCell.getValue();
    const isPresence=description==='CONFERMA PRESENZE TUTTI ATLETI';
    if(!isPresence&&!(existingDue instanceof Date)) {
      const due=new Date(); due.setHours(12,0,0,0); due.setDate(due.getDate()+2);
      dueCell.setValue(due).setNumberFormat('dd/MM/yyyy');
    }
    if(eventNormalizeTextV7_(sheet.getRange(row,5).getDisplayValue())!=='FATTO') {
      sheet.getRange(row,5).setFormula(eventTaskStatusFormulaV9_(row));
    }
  });
}

function eventAutoLinkConfirmationTasksV9_() {
  const sheet=SpreadsheetApp.getActive().getSheetByName(EVENT_APP.SHEETS.TASKS);
  if(!sheet)return;
  for(let row=2;row<=Math.min(500,Math.max(sheet.getLastRow(),2));row++) eventAutoLinkConfirmationTaskRowV9_(sheet,row);
}

function eventAutoLinkConfirmationTaskRowV9_(sheet,row) {
  if(Number(sheet.getRange(row,4).getValue()||0)>0)return;
  const desc=eventNormalizeTextV7_(sheet.getRange(row,1).getDisplayValue());
  if(!desc)return;
  const parents={
    'CONFERMA GOMMONE':'RICHIESTA GOMMONE',
    'CONFERMA SOGGIORNO':'RICHIESTA SOGGIORNO',
    'CONFERMA PASTI':'SOLUZIONE PASTI',
    'CONFERMA PRESENZE TUTTI ATLETI':'CONVOCAZIONE ATLETI',
    'CONFERMA OSPITALITA CIRCOLO':'OSPITALITA CIRCOLO'
  };
  const parent=parents[desc]||'';
  if(!parent)return;
  const values=sheet.getRange(2,1,Math.max(1,row-2),3).getDisplayValues();
  for(let i=values.length-1;i>=0;i--) {
    if(eventNormalizeTextV7_(values[i][0])===parent&&Number(values[i][2]||0)>0) {
      sheet.getRange(row,4).setValue(Number(values[i][2]));
      return;
    }
  }
}

function eventNormalizeTextV7_(value){return String(value===null||value===undefined?'':value).trim().toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ');}
function eventMeta_(){const sheet=SpreadsheetApp.getActive().getSheetByName(EVENT_APP.SHEETS.META);if(!sheet)throw new Error('Foglio _META non trovato.');const values=sheet.getRange(1,1,sheet.getLastRow(),2).getDisplayValues();const out={};values.forEach(r=>{if(r[0])out[String(r[0]).trim()]=r[1];});return out;}
function eventWriteLocalMetaV7_(key,value){const sheet=SpreadsheetApp.getActive().getSheetByName(EVENT_APP.SHEETS.META);if(!sheet)return;const rows=sheet.getRange(1,1,Math.max(sheet.getLastRow(),1),2).getDisplayValues(),target=eventNormalizeTextV7_(key);for(let i=0;i<rows.length;i++)if(eventNormalizeTextV7_(rows[i][0])===target){sheet.getRange(i+1,2).setValue(value);return;}sheet.appendRow([key,value]);}
function eventCommitmentMapV7_(){return {'RIUNIONI/MISSIONI|*':'2-1','ALLENAMENTO|29ER':'28','ALLENAMENTO|420':'29','ALLENAMENTO|NACRA 15':'30','ALLENAMENTO|IQFOIL':'31','ALLENAMENTO|ILCA 6':'32','ALLENAMENTO|ILCA 4':'33-1','ALLENAMENTO INTERZONALE|ILCA 4':'33-2','ALLENAMENTO|KITEFOIL':'34','ALLENAMENTO|WING':'35','TEST FISICI|*':'37','COLLEGIALE|*':'38','REGATA|29ER':'39-1','REGATA|420':'40','REGATA|NACRA 15':'41','REGATA|IQFOIL':'42','REGATA|ILCA 6':'43','REGATA|ILCA 4':'44','REGATA|KITEFOIL':'45','REGATA|WING':'46','OSS.REGATE ITA|*':'36','CAMP. NAZ GIOVANILE SINGOLO|*':'48-1','CAMP. NAZ GIOVANILE DOPPIO|*':'48-1','FOIL ACADEMY|*':'50','ISCRIZIONI REGATE|29ER':'47-1','ISCRIZIONI REGATE|420':'47-2','ISCRIZIONI REGATE|NACRA 15':'47-3','ISCRIZIONI REGATE|IQFOIL':'47-4','ISCRIZIONI REGATE|ILCA 6':'47-5','ISCRIZIONI REGATE|KITEFOIL':'47-6','ISCRIZIONI REGATE|WING':'47-8','ISCRIZIONE YWC|TUTTE':'171','MANUTENZIONE ALLENAMENTI|*':'39-3','MANUTENZIONE REGATE|*':'39-4','NOLEGGI ALLENAMENTI|*':'52','NOLEGGI REGATE|*':'51'};}
function eventCanonicalTypeV7_(value){const t=eventNormalizeTextV7_(value),aliases={'RIUNIONE/MISSIONE':'RIUNIONI/MISSIONI','RIUNIONI / MISSIONI':'RIUNIONI/MISSIONI','OSSERVAZIONE REGATA':'OSS.REGATE ITA','ISCRIZIONE REGATA':'ISCRIZIONI REGATE'};return aliases[t]||t;}
function eventCanonicalClassV7_(value){const c=eventNormalizeTextV7_(value);return c==='KITE'?'KITEFOIL':c;}
function eventResolveCommitmentV7_(meta){const type=eventCanonicalTypeV7_(meta.EVENT_TYPE||''),cls=eventCanonicalClassV7_(meta.EVENT_CLASS||''),map=eventCommitmentMapV7_();return map[type+'|'+cls]||map[type+'|*']||map[type+'|TUTTE']||'';}
function eventTechnicianDirectoryV7_(){return [['ZAGGIA','Leonardo','Zaggia'],['CRISI','Andrea','Crisi'],['RAVEGLIA','Matteo','Raveglia'],['CARICATO','Francesco','Caricato'],['PICCIAU','Gianluigi','Picciau'],['SENSINI','Alessandra','Sensini'],['NUICOLUCCI','Matteo','Nuicolucci'],['CAMBONI','Mattia','Camboni'],['CANGEMI','Antonino','Cangemi'],['LOPERFIDO','Daniel','Loperfido']];}
function eventPopulateTechniciansV7_(meta){const raw=eventNormalizeTextV7_(meta.EVENT_TECHNICIANS||'');if(!raw)return[];const found=eventTechnicianDirectoryV7_().filter(x=>raw.indexOf(x[0])>=0);if(!found.length)return[];const full=found.map(x=>x[1]+' '+x[2]);eventWriteLocalMetaV7_('EVENT_TECHNICIANS',full.join(', '));const sheet=SpreadsheetApp.getActive().getSheetByName(EVENT_APP.SHEETS.PARTICIPANTS);if(!sheet)return full;const rows=sheet.getRange(3,1,20,19).getDisplayValues(),existing=new Set(rows.map(r=>eventNormalizeTextV7_((r[0]||'')+' '+(r[1]||''))));found.forEach(x=>{const key=eventNormalizeTextV7_(x[1]+' '+x[2]);if(existing.has(key))return;let target=-1;for(let i=0;i<rows.length;i++)if(!String(rows[i][0]||'').trim()&&!String(rows[i][1]||'').trim()){target=i+3;rows[i][0]=x[1];rows[i][1]=x[2];break;}if(target<0)return;sheet.getRange(target,1).setValue(x[1]);sheet.getRange(target,2).setValue(x[2]);sheet.getRange(target,8).setValue('TECNICO');sheet.getRange(target,9).setValue('CONFERMATO');existing.add(key);});return full;}
function eventGenerateDocuments(){const ui=SpreadsheetApp.getUi(),meta=eventMeta_();let prompt='Scegli i documenti da creare, separando i numeri con una virgola:\n\n';EVENT_APP.TEMPLATE_SPECS.forEach((t,i)=>prompt+=(i+1)+' = '+t.label+'\n');const response=ui.prompt('Genera documenti',prompt+'\nEsempio: 1,2,4',ui.ButtonSet.OK_CANCEL);if(response.getSelectedButton()!==ui.Button.OK)return;const keys=eventParseDocumentSelection_(response.getResponseText());if(!keys.length){ui.alert('Nessun documento valido selezionato.');return;}const folder=DriveApp.getFolderById(meta.EVENT_FOLDER_ID),replacements=eventDocumentReplacements_(meta),created=[];keys.forEach(key=>{const spec=EVENT_APP.TEMPLATE_SPECS.find(x=>x.key===key);if(!spec)return;const copy=DriveApp.getFileById(spec.id).makeCopy(eventDocumentName_(key,replacements),folder);eventFillDocument_(copy.getId(),replacements,key);eventRegisterDocument_(meta,spec,copy);created.push(copy.getName());});ui.alert('Documenti creati',created.map(x=>'• '+x).join('\n'),ui.ButtonSet.OK);}
function eventParseDocumentSelection_(text){const seen={};return String(text||'').split(/[;,\s]+/).map(x=>parseInt(x,10)).filter(n=>{if(!n||n<1||n>EVENT_APP.TEMPLATE_SPECS.length||seen[n])return false;seen[n]=true;return true;}).map(n=>EVENT_APP.TEMPLATE_SPECS[n-1].key);}
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
