const EXPENSES_V8 = Object.freeze({
  INPUT_ROW:9,
  PAYMENT_START_ROW:13,
  PAYMENT_END_ROW:16,
  IMPORT_CHECKBOX:'K7',
  HISTORY_HEADER_ROW:19,
  HISTORY_START_ROW:20,
  HISTORY_END_ROW:500,
  HISTORY_COLS:18,
  PAYMENT_TYPES:Object.freeze(['AFOR','CARTA DI CREDITO','SALDO FATTURA','ALTRO PAGAMENTO'])
});

function eventHandleExpenseEditV8_(e) {
  const sheet=e.range.getSheet();
  if (e.range.getA1Notation()===EXPENSES_V8.IMPORT_CHECKBOX && e.value==='TRUE') {
    try { eventImportExpenseV8(); }
    catch (err) {
      sheet.getRange(EXPENSES_V8.IMPORT_CHECKBOX).setValue(false);
      SpreadsheetApp.getActive().toast(err.message||String(err),'Spese',7);
    }
    return;
  }
  const row=e.range.getRow(),col=e.range.getColumn();
  if (row>=EXPENSES_V8.PAYMENT_START_ROW && row<=EXPENSES_V8.PAYMENT_END_ROW && col>=2 && col<=7) {
    eventNormalizePaymentInputRowV8_(sheet,row);
    return;
  }
  if (row>=EXPENSES_V8.HISTORY_START_ROW && row<=EXPENSES_V8.HISTORY_END_ROW && col<=12) {
    eventNormalizeHistoryPaymentRowV8_(sheet,row);
    eventRecalculateQuoteStatusesV10_(sheet);
    eventSortHistoryV10_(sheet);
    eventSyncExpenseTasksV8_();
    eventRefreshExpenseDashboardV8_();
  }
}

function eventApplyCommitmentV8_(meta) {
  const imp=eventResolveCommitmentV7_(meta)||String(meta.EVENT_COMMITMENT||'').trim();
  if (!imp) return '';
  eventWriteLocalMetaV7_('EVENT_COMMITMENT',imp);
  const sheet=SpreadsheetApp.getActive().getSheetByName(EVENT_APP.SHEETS.EXPENSES);
  if (sheet && !String(sheet.getRange('E9').getDisplayValue()||'').trim()) sheet.getRange('E9').setValue(imp).setNumberFormat('@');
  return imp;
}

/** Checkbox PAGATO = unica fonte di verità. */
function eventNormalizePaymentInputRowV8_(sheet,row) {
  const quoteNo=Number(sheet.getRange(row,2).getValue()||0);
  const typeCell=sheet.getRange(row,3);
  const type=eventPaymentTypeV8_(typeCell.getDisplayValue());
  if(type && type!==typeCell.getDisplayValue())typeCell.setValue(type);

  const amountCell=sheet.getRange(row,4);
  let amount=Number(amountCell.getValue()||0);
  const paid=sheet.getRange(row,5).getValue()===true;
  const dateCell=sheet.getRange(row,6);
  const paymentDate=dateCell.getValue();

  // Per CC / saldo fattura compila solo l'importo residuo; NON marca PAGATO.
  if ((type==='CARTA DI CREDITO'||type==='SALDO FATTURA') && !(amount>0)) {
    const quote=quoteNo>0?eventFindQuoteByNumberV10_(sheet,quoteNo):null;
    const budget=quote?quote.amount:Number(sheet.getRange('A9').getValue()||0);
    if(budget>0){
      const alreadyAllocated=eventAllocatedInputAmountV10_(sheet,row,quoteNo);
      const historyPaid=quote?eventPaidHistoryForQuoteV10_(sheet,quoteNo):0;
      amount=Math.max(budget-historyPaid-alreadyAllocated,0);
      if(amount>0)amountCell.setValue(amount).setNumberFormat('€ #,##0.00');
    }
  }

  if(paid){
    if(!(paymentDate instanceof Date))dateCell.setValue(new Date()).setNumberFormat('dd/MM/yyyy');
  }else if(paymentDate instanceof Date||String(dateCell.getDisplayValue()||'').trim()){
    dateCell.clearContent();
  }
  SpreadsheetApp.flush();
}

function eventAllocatedInputAmountV10_(sheet,excludeRow,quoteNo) {
  let total=0;
  const rows=sheet.getRange(EXPENSES_V8.PAYMENT_START_ROW,2,4,3).getValues();
  rows.forEach((r,i)=>{
    const row=EXPENSES_V8.PAYMENT_START_ROW+i;
    if(row===excludeRow)return;
    const q=Number(r[0]||0);
    if((quoteNo>0&&q!==quoteNo)||(quoteNo<=0&&q>0))return;
    total+=Number(r[2]||0);
  });
  return total;
}

function eventPaidHistoryForQuoteV10_(sheet,quoteNo){
  let total=0;
  const rows=sheet.getRange(EXPENSES_V8.HISTORY_START_ROW,1,EXPENSES_V8.HISTORY_END_ROW-EXPENSES_V8.HISTORY_START_ROW+1,8).getValues();
  rows.forEach(r=>{if(Number(r[0]||0)===quoteNo&&eventNormalizeTextV7_(r[1])!=='PREVENTIVO'&&eventNormalizeTextV7_(r[5])==='PAGATO')total+=Number(r[2]||0);});
  return total;
}

function eventPaymentTypeV8_(value) {
  const t=eventNormalizeTextV7_(value);
  if(t==='FATTURA'||t==='PAGAMENTO FATTURA')return'SALDO FATTURA';
  if(t==='CARTA'||t==='CC')return'CARTA DI CREDITO';
  return EXPENSES_V8.PAYMENT_TYPES.includes(t)?t:'';
}

function eventQuoteStatusV10_(budget,paid){
  const b=Number(budget||0),p=Number(paid||0);
  if(!(b>0)||p<=0.005)return'DA PAGARE';
  if(p+0.005>=b)return'PAGATO';
  return'PARZIALMENTE PAGATO';
}

function eventNormalizeHistoryPaymentRowV8_(sheet,row){
  const movement=eventNormalizeTextV7_(sheet.getRange(row,2).getDisplayValue());
  if(!movement||movement==='PREVENTIVO')return;
  const statusCell=sheet.getRange(row,6);
  let status=eventNormalizeTextV7_(statusCell.getDisplayValue());
  if(status==='PARZIALMENTE PAGATO'){status='DA PAGARE';statusCell.setValue(status);}
  const dateCell=sheet.getRange(row,7),paymentDate=dateCell.getValue();
  if(status==='PAGATO'){
    if(!(paymentDate instanceof Date))dateCell.setValue(new Date()).setNumberFormat('dd/MM/yyyy');
  }else if(paymentDate instanceof Date||String(dateCell.getDisplayValue()||'').trim())dateCell.clearContent();
}

function eventImportExpenseV8(){
  const ss=SpreadsheetApp.getActive(),sheet=ss.getSheetByName(EVENT_APP.SHEETS.EXPENSES);
  if(!sheet)throw new Error('Foglio Spese non trovato.');

  const shared=sheet.getRange(9,1,1,9).getValues()[0];
  const budget=Number(shared[0]||0),description=String(shared[1]||'').trim(),category=String(shared[2]||'').trim(),rif=String(shared[3]||'').trim();
  const meta=eventMeta_();
  const imp=String(shared[4]||'').trim()||eventResolveCommitmentV7_(meta)||String(meta.EVENT_COMMITMENT||'').trim();
  const documents=String(shared[5]||'').trim(),notes=String(shared[6]||'').trim();
  const creatingQuote=budget>0||description||category||rif||documents||notes;

  if(creatingQuote){
    if(!(budget>0))throw new Error('Inserisci il PREVENTIVO totale.');
    if(!description)throw new Error('La DESCRIZIONE è obbligatoria. Inserisci qui anche il nome del fornitore.');
    if(!category)throw new Error('Seleziona la TIPOLOGIA.');
  }

  for(let row=13;row<=16;row++)eventNormalizePaymentInputRowV8_(sheet,row);
  const rawPayments=sheet.getRange(13,2,4,6).getValues(),payments=[];
  const newQuoteNo=creatingQuote?eventNextQuoteNumberV10_(sheet):0;

  rawPayments.forEach((r,i)=>{
    const requestedQuoteNo=Number(r[0]||0),type=eventPaymentTypeV8_(r[1]),amount=Number(r[2]||0),paid=r[3]===true,paymentDate=paid&&r[4] instanceof Date?r[4]:'',due=r[5] instanceof Date?r[5]:'';
    const hasAnything=requestedQuoteNo>0||!!type||amount>0||paid||!!due;
    if(!hasAnything)return;
    if(!type)throw new Error('Seleziona il tipo del pagamento n. '+(i+1)+'.');
    if(!(amount>0))throw new Error('Inserisci l’importo del pagamento n. '+(i+1)+'.');
    if(!paid&&!due)throw new Error('Il pagamento n. '+(i+1)+' non è pagato: inserisci una SCADENZA.');
    const quoteNo=requestedQuoteNo>0?requestedQuoteNo:newQuoteNo;
    if(!(quoteNo>0))throw new Error('Nel pagamento n. '+(i+1)+' indica il N. PREVENTIVO.');
    const existing=eventFindQuoteByNumberV10_(sheet,quoteNo);
    if(!existing&&quoteNo!==newQuoteNo)throw new Error('Preventivo n. '+quoteNo+' non trovato.');
    payments.push({quoteNo:quoteNo,type:type,amount:amount,paid:paid,paymentDate:paymentDate,due:due,number:i+1});
  });

  if(!creatingQuote&&!payments.length)throw new Error('Inserisci un preventivo oppure almeno un pagamento.');

  if(creatingQuote){
    const existingAllocated=payments.filter(p=>p.quoteNo===newQuoteNo).reduce((s,p)=>s+p.amount,0);
    if(existingAllocated>budget+0.005)throw new Error('La somma dei pagamenti del nuovo preventivo supera il preventivo.');
  }

  const byExistingQuote={};
  payments.forEach(p=>{if(p.quoteNo!==newQuoteNo)(byExistingQuote[p.quoteNo]||(byExistingQuote[p.quoteNo]=[])).push(p);});
  Object.keys(byExistingQuote).forEach(k=>{
    const q=eventFindQuoteByNumberV10_(sheet,Number(k));
    const historyAllocated=eventAllocatedHistoryForQuoteV10_(sheet,Number(k));
    const adding=byExistingQuote[k].reduce((s,p)=>s+p.amount,0);
    if(q&&historyAllocated+adding>q.amount+0.005)throw new Error('I pagamenti del preventivo n. '+k+' superano il suo importo.');
  });

  const append=[];
  const now=new Date();
  let newQuoteId='';
  if(creatingQuote){
    newQuoteId='PREV-'+Utilities.getUuid();
    const paidTotal=payments.filter(p=>p.quoteNo===newQuoteNo&&p.paid).reduce((s,p)=>s+p.amount,0);
    append.push([newQuoteNo,'PREVENTIVO',budget,description,category,eventQuoteStatusV10_(budget,paidTotal),'','',rif,imp,documents,notes,newQuoteId,newQuoteId,'','',now,now]);
  }

  payments.forEach(p=>{
    const q=p.quoteNo===newQuoteNo?{id:newQuoteId,description:description,category:category,rif:rif,imp:imp,documents:documents,notes:notes}:eventFindQuoteByNumberV10_(sheet,p.quoteNo);
    append.push([p.quoteNo,p.type,p.amount,q.description,q.category,p.paid?'PAGATO':'DA PAGARE',p.paid?p.paymentDate:'',p.due,q.rif,q.imp,q.documents,q.notes,'PAG-'+Utilities.getUuid(),q.id,'','',now,now]);
  });

  const target=eventFindHistoryBlockV8_(sheet,append.length);
  if(!target)throw new Error('Elenco spese pieno.');
  sheet.getRange(target,1,append.length,EXPENSES_V8.HISTORY_COLS).setValues(append);
  sheet.getRange(target,3,append.length,1).setNumberFormat('€ #,##0.00');
  sheet.getRange(target,7,append.length,2).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(target,10,append.length,1).setNumberFormat('@');

  sheet.getRange('A9:G9').clearContent();
  sheet.getRange('B13:D16').clearContent();
  sheet.getRange('E13:E16').setValue(false);
  sheet.getRange('F13:G16').clearContent();
  sheet.getRange(EXPENSES_V8.IMPORT_CHECKBOX).setValue(false);
  eventApplyCommitmentV8_(meta);
  eventRecalculateQuoteStatusesV10_(sheet);
  eventSortHistoryV10_(sheet);
  eventSyncExpenseTasksV8_();
  eventRefreshExpenseDashboardV8_();
  ss.toast((creatingQuote?'Preventivo n. '+newQuoteNo+' e ':'')+payments.length+' pagamento/i registrati.','Spese',5);
}

function eventNextQuoteNumberV10_(sheet){
  const values=sheet.getRange(EXPENSES_V8.HISTORY_START_ROW,1,EXPENSES_V8.HISTORY_END_ROW-EXPENSES_V8.HISTORY_START_ROW+1,1).getValues();
  let max=0;values.forEach(r=>{const n=Number(r[0]||0);if(Number.isFinite(n)&&n>max)max=n;});return max+1;
}

function eventFindQuoteByNumberV10_(sheet,quoteNo){
  if(!(quoteNo>0))return null;
  const rows=sheet.getRange(EXPENSES_V8.HISTORY_START_ROW,1,EXPENSES_V8.HISTORY_END_ROW-EXPENSES_V8.HISTORY_START_ROW+1,EXPENSES_V8.HISTORY_COLS).getValues();
  for(let i=0;i<rows.length;i++){
    const r=rows[i];
    if(Number(r[0]||0)!==quoteNo||eventNormalizeTextV7_(r[1])!=='PREVENTIVO')continue;
    return{id:String(r[13]||'').trim(),number:quoteNo,amount:Number(r[2]||0),description:String(r[3]||'').trim(),category:String(r[4]||'').trim(),rif:String(r[8]||'').trim(),imp:String(r[9]||'').trim(),documents:String(r[10]||'').trim(),notes:String(r[11]||'').trim()};
  }
  return null;
}

function eventAllocatedHistoryForQuoteV10_(sheet,quoteNo){
  let total=0;const rows=sheet.getRange(EXPENSES_V8.HISTORY_START_ROW,1,EXPENSES_V8.HISTORY_END_ROW-EXPENSES_V8.HISTORY_START_ROW+1,3).getValues();
  rows.forEach(r=>{if(Number(r[0]||0)===quoteNo&&eventNormalizeTextV7_(r[1])!=='PREVENTIVO')total+=Number(r[2]||0);});return total;
}

function eventFindHistoryBlockV8_(sheet,needed){
  const values=sheet.getRange(EXPENSES_V8.HISTORY_START_ROW,1,EXPENSES_V8.HISTORY_END_ROW-EXPENSES_V8.HISTORY_START_ROW+1,1).getDisplayValues();
  let run=0;for(let i=0;i<values.length;i++){if(!String(values[i][0]||'').trim())run++;else run=0;if(run>=needed)return EXPENSES_V8.HISTORY_START_ROW+i-needed+1;}return 0;
}

function eventSortHistoryV10_(sheet){
  const count=EXPENSES_V8.HISTORY_END_ROW-EXPENSES_V8.HISTORY_START_ROW+1;
  const range=sheet.getRange(EXPENSES_V8.HISTORY_START_ROW,1,count,EXPENSES_V8.HISTORY_COLS);
  const all=range.getValues();
  const rows=all.filter(r=>Number(r[0]||0)>0&&String(r[1]||'').trim());
  rows.sort((a,b)=>{
    const qa=Number(a[0]||0),qb=Number(b[0]||0);if(qa!==qb)return qa-qb;
    const ra=eventNormalizeTextV7_(a[1])==='PREVENTIVO'?0:1,rb=eventNormalizeTextV7_(b[1])==='PREVENTIVO'?0:1;if(ra!==rb)return ra-rb;
    const da=a[16] instanceof Date?a[16].getTime():0,db=b[16] instanceof Date?b[16].getTime():0;return da-db;
  });
  range.clearContent();
  if(rows.length)sheet.getRange(EXPENSES_V8.HISTORY_START_ROW,1,rows.length,EXPENSES_V8.HISTORY_COLS).setValues(rows);
  eventFormatAllHistoryGroupsV10_(sheet,rows.length);
}

function eventFormatAllHistoryGroupsV10_(sheet,rowCount){
  const full=sheet.getRange(EXPENSES_V8.HISTORY_START_ROW,1,Math.max(1,EXPENSES_V8.HISTORY_END_ROW-EXPENSES_V8.HISTORY_START_ROW+1),12);
  full.setBorder(false,false,false,false,false,false);
  if(!(rowCount>0))return;
  const vals=sheet.getRange(EXPENSES_V8.HISTORY_START_ROW,1,rowCount,2).getValues();
  let start=0;
  while(start<vals.length){
    const quoteNo=Number(vals[start][0]||0);let end=start;
    while(end+1<vals.length&&Number(vals[end+1][0]||0)===quoteNo)end++;
    const r=sheet.getRange(EXPENSES_V8.HISTORY_START_ROW+start,1,end-start+1,12);
    r.setBorder(true,true,true,true,false,false,'#5f6368',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sheet.getRange(EXPENSES_V8.HISTORY_START_ROW+start,1,1,12).setFontWeight('bold');
    if(end>start)sheet.getRange(EXPENSES_V8.HISTORY_START_ROW+start+1,1,end-start,12).setFontWeight('normal');
    start=end+1;
  }
}

function eventRecalculateQuoteStatusesV10_(sheet){
  if(!sheet)return;
  const count=EXPENSES_V8.HISTORY_END_ROW-EXPENSES_V8.HISTORY_START_ROW+1;
  const rows=sheet.getRange(EXPENSES_V8.HISTORY_START_ROW,1,count,EXPENSES_V8.HISTORY_COLS).getValues();
  const quotes={};
  rows.forEach((r,i)=>{if(eventNormalizeTextV7_(r[1])==='PREVENTIVO'){const no=Number(r[0]||0);if(no>0)quotes[no]={row:EXPENSES_V8.HISTORY_START_ROW+i,budget:Number(r[2]||0),paid:0};}});
  rows.forEach(r=>{const no=Number(r[0]||0);if(quotes[no]&&eventNormalizeTextV7_(r[1])!=='PREVENTIVO'&&eventNormalizeTextV7_(r[5])==='PAGATO')quotes[no].paid+=Number(r[2]||0);});
  Object.keys(quotes).forEach(no=>{const q=quotes[no];sheet.getRange(q.row,6).setValue(eventQuoteStatusV10_(q.budget,q.paid));});
}

function eventRefreshExpenseDashboardV8_(){
  const ss=SpreadsheetApp.getActive(),expense=ss.getSheetByName(EVENT_APP.SHEETS.EXPENSES),participants=ss.getSheetByName(EVENT_APP.SHEETS.PARTICIPANTS);if(!expense)return;
  eventRecalculateQuoteStatusesV10_(expense);
  const rows=expense.getRange(EXPENSES_V8.HISTORY_START_ROW,1,EXPENSES_V8.HISTORY_END_ROW-EXPENSES_V8.HISTORY_START_ROW+1,12).getValues();
  const buckets={VIAGGI:0,VITTO:0,ALLOGGIO:0,NOLEGGI:0,RIMBORSI:0,ALTRO:0};let forecast=0,paid=0;
  rows.forEach(r=>{if(!Number(r[0]||0))return;const movement=eventNormalizeTextV7_(r[1]),amount=Number(r[2]||0);if(movement==='PREVENTIVO'){forecast+=amount;const cat=eventNormalizeTextV7_(r[4]||'ALTRO'),bucket=['VIAGGI','VITTO','ALLOGGIO','NOLEGGI'].includes(cat)?cat:'ALTRO';buckets[bucket]+=amount;}else if(eventNormalizeTextV7_(r[5])==='PAGATO')paid+=amount;});
  if(participants)participants.getRange(3,1,15,19).getValues().forEach(r=>{const name=String(r[0]||'').trim(),surname=String(r[1]||'').trim();if(!name&&!surname)return;if(eventNormalizeTextV7_(r[7])==='TECNICO')return;const status=eventNormalizeTextV7_(r[8]),max=Number(r[9]||0),passed=Number(r[10]||0),current=status==='ASSENTE'?0:(passed>0?passed:max);forecast+=current;buckets.RIMBORSI+=current;if(passed>0)paid+=passed;});
  expense.getRange('D2').setValue(forecast).setNumberFormat('€ #,##0.00');expense.getRange('F2').setValue(paid).setNumberFormat('€ #,##0.00');expense.getRange('H2').setValue(Math.max(forecast-paid,0)).setNumberFormat('€ #,##0.00');expense.getRange('A5:F5').setValues([[buckets.VIAGGI,buckets.VITTO,buckets.ALLOGGIO,buckets.NOLEGGI,buckets.RIMBORSI,buckets.ALTRO]]).setNumberFormat('€ #,##0.00');
}

function eventFormatCurrencyV10_(amount){const n=Math.max(0,Number(amount||0)),parts=n.toFixed(2).split('.');parts[0]=parts[0].replace(/\B(?=(\d{3})+(?!\d))/g,'.');return'€ '+parts[0]+','+parts[1];}
function eventExpenseTaskNoteV10_(description,amount){return String(description||'').trim()+' — '+eventFormatCurrencyV10_(amount);}

function eventSyncExpenseTasksV8_(){
  const ss=SpreadsheetApp.getActive(),expense=ss.getSheetByName(EVENT_APP.SHEETS.EXPENSES),tasks=ss.getSheetByName(EVENT_APP.SHEETS.TASKS);if(!expense||!tasks)return;
  eventRecalculateQuoteStatusesV10_(expense);
  const rows=expense.getRange(EXPENSES_V8.HISTORY_START_ROW,1,EXPENSES_V8.HISTORY_END_ROW-EXPENSES_V8.HISTORY_START_ROW+1,EXPENSES_V8.HISTORY_COLS).getValues(),quotes={},quoteHasAfor={};
  rows.forEach(r=>{if(eventNormalizeTextV7_(r[1])==='PREVENTIVO'&&r[13])quotes[String(r[13])]={number:Number(r[0]||0),amount:Number(r[2]||0),description:String(r[3]||'').trim()};});
  rows.forEach(r=>{
    if(eventPaymentTypeV8_(r[1])!=='AFOR')return;const quoteId=String(r[13]||'').trim();if(quoteId)quoteHasAfor[quoteId]=true;
    const amount=Number(r[2]||0),description=String(r[3]||'').trim(),status=eventNormalizeTextV7_(r[5]),paidDate=r[6] instanceof Date?r[6]:null,due=r[7] instanceof Date?r[7]:null,movementId=String(r[12]||'').trim();if(!movementId)return;
    const noteText=eventExpenseTaskNoteV10_(description,amount),payKey='AFOR_PAY:'+movementId,contKey='AFOR_CONT:'+movementId;
    if(status==='DA PAGARE'&&due)eventEnsureAutoTaskV8_(tasks,'Pagare AFOR',due,eventDateReachedV9_(due)?'DA FARE':'IN ATTESA',payKey,noteText);
    else if(status==='PAGATO'){eventMarkAutoTaskDoneV8_(tasks,payKey);if(paidDate){const contDue=new Date(paidDate);contDue.setHours(12,0,0,0);contDue.setDate(contDue.getDate()+1);eventEnsureAutoTaskV8_(tasks,'Inviare contabile AFOR',contDue,eventDateReachedV9_(contDue)?'DA FARE':'IN ATTESA',contKey,noteText);}}
  });
  const end=eventParseIsoDate_(eventMeta_().EVENT_END);
  if(end)Object.keys(quoteHasAfor).forEach(quoteId=>{const q=quotes[quoteId];if(!q)return;eventEnsureAutoTaskV8_(tasks,'Richiedere fattura',end,eventDateReachedV9_(end)?'DA FARE':'IN ATTESA','FATTURA_AFOR:'+quoteId,eventExpenseTaskNoteV10_(q.description,q.amount));});
}

function eventAutoTaskMarkerV8_(key){return'[AUTO_KEY='+key+']';}
function eventFindAutoTaskRowV8_(sheet,key){const marker=eventAutoTaskMarkerV8_(key),range=sheet.getRange(2,2,499,1),cellNotes=range.getNotes(),visible=range.getDisplayValues();for(let i=0;i<cellNotes.length;i++)if(String(cellNotes[i][0]||'').indexOf(marker)>=0)return i+2;for(let i=0;i<visible.length;i++)if(String(visible[i][0]||'').indexOf(marker)>=0){sheet.getRange(i+2,2).setNote(marker);return i+2;}return 0;}
function eventEnsureAutoTaskV8_(sheet,description,due,status,key,noteText){let row=eventFindAutoTaskRowV8_(sheet,key);if(!row){const descs=sheet.getRange(2,1,499,1).getDisplayValues(),idx=descs.findIndex(r=>!String(r[0]||'').trim());if(idx<0)return 0;row=idx+2;sheet.getRange(row,1).setValue(description);eventPrepareTaskRowV9_(sheet,row);}sheet.getRange(row,2).setValue(noteText||'').setNote(eventAutoTaskMarkerV8_(key));if(due instanceof Date)sheet.getRange(row,6).setValue(due).setNumberFormat('dd/MM/yyyy');if(eventNormalizeTextV7_(sheet.getRange(row,5).getDisplayValue())!=='FATTO'){const desired=due instanceof Date?(eventDateReachedV9_(due)?'DA FARE':'IN ATTESA'):(status||'IN ATTESA');sheet.getRange(row,5).setValue(desired);}return row;}
function eventMarkAutoTaskDoneV8_(sheet,key){const row=eventFindAutoTaskRowV8_(sheet,key);if(row>0)sheet.getRange(row,5).setValue('FATTO');}

function eventSaveExpensesV8_(master,child,eventId,meta){
  const local=child.getSheetByName(EVENT_APP.SHEETS.EXPENSES);if(!local)throw new Error('Foglio Spese non trovato nella Scheda evento.');
  eventApplyCommitmentV8_(meta);eventRecalculateQuoteStatusesV10_(local);eventSortHistoryV10_(local);eventSyncExpenseTasksV8_();eventRefreshExpenseDashboardV8_();
  const backend=eventMasterSheet_(master,'_SPESE'),oldRows=backend.getDataRange().getValues(),oldById={};for(let i=1;i<oldRows.length;i++)if(String(oldRows[i][1]||'')===String(eventId)&&oldRows[i][0])oldById[String(oldRows[i][0])]=oldRows[i];
  const history=local.getRange(20,1,481,EXPENSES_V8.HISTORY_COLS).getValues(),now=new Date(),quoteTotals={};
  history.forEach(r=>{if(eventNormalize_(r[1])==='PREVENTIVO'&&r[13])quoteTotals[String(r[13])]={budget:Number(r[2]||0),paid:0,quoteNo:Number(r[0]||0)};});
  history.forEach(r=>{const q=String(r[13]||'');if(q&&eventNormalize_(r[1])!=='PREVENTIVO'&&eventNormalize_(r[5])==='PAGATO'&&quoteTotals[q])quoteTotals[q].paid+=Number(r[2]||0);});
  const rows=[];
  history.forEach((r,index)=>{
    const quoteNo=Number(r[0]||0),movement=eventNormalize_(r[1]),amount=Number(r[2]||0),description=String(r[3]||'').trim();if(!description||!movement)return;
    const category=eventExpenseCategoryV7_(r[4],description),localStatus=eventNormalize_(r[5]),paid=movement!=='PREVENTIVO'&&localStatus==='PAGATO',paidDate=paid&&r[6] instanceof Date?r[6]:'',due=!paid&&r[7] instanceof Date?r[7]:'',userRif=String(r[8]||'').trim(),imp=String(r[9]||'').trim(),documents=String(r[10]||'').trim(),note=String(r[11]||'').trim();
    const id=String(r[12]||'').trim()||(movement==='PREVENTIVO'?'PREV-':'PAG-')+Utilities.getUuid(),quoteId=String(r[13]||'').trim()||(movement==='PREVENTIVO'?id:''),old=oldById[id]?oldById[id].slice(0,26):new Array(26).fill(''),ceb=eventResolveCeb_(master,meta,category,String(r[14]||'').trim()||old[4]),createdAt=r[16] instanceof Date?r[16]:(old[16] instanceof Date?old[16]:now),qInfo=quoteTotals[quoteId];
    old[0]=id;old[1]=eventId;old[2]=movement==='PREVENTIVO'?'PREVENTIVO':'PAGAMENTO';old[3]=category;old[4]=ceb;old[5]=description;old[6]='';old[7]=String(r[15]||'').trim();old[8]=movement==='PREVENTIVO'?amount:0;old[9]=movement!=='PREVENTIVO'&&paid?amount:0;old[10]=due;
    if(movement==='PREVENTIVO'){const qp=qInfo?Number(qInfo.paid||0):0,qb=qInfo?Number(qInfo.budget||amount):amount;old[11]=qp<=0.005?'DA SALDARE':(qp+0.005>=qb?'SALDATO':'PARZIALMENTE SALDATO');}else old[11]=paid?eventBackendPaidStatusV8_(movement):'DA PAGARE';
    old[12]=paidDate;old[13]='PREV. '+quoteNo;old[14]=documents;old[15]='[MOVIMENTO='+movement+'] [ID_PREVENTIVO='+quoteId+'] [N_PREVENTIVO='+quoteNo+']'+(imp?' [IMP='+imp+']':'')+(note?' '+note:'');old[16]=createdAt;old[17]=now;if(userRif)old[18]='RICEVUTO';else if(movement==='PREVENTIVO'&&amount>1000)old[18]=old[18]||'DA RICHIEDERE';else old[18]=old[18]||'NON NECESSARIO';old[20]=userRif;old[21]=movement==='SALDO FATTURA'&&paid;old[22]=old[21]?paidDate:(old[22]||'');old[25]=paid&&movement!=='PREVENTIVO'?(old[25]||now):(movement==='PREVENTIVO'&&qInfo&&qInfo.paid+0.005>=qInfo.budget?(old[25]||now):'');rows.push(old);
    if(!String(r[12]||'').trim())local.getRange(20+index,13).setValue(id);if(!String(r[13]||'').trim())local.getRange(20+index,14).setValue(quoteId);if(!String(r[14]||'').trim()&&ceb)local.getRange(20+index,15).setValue(ceb);local.getRange(20+index,17).setValue(createdAt);local.getRange(20+index,18).setValue(now);
  });
  const participants=child.getSheetByName(EVENT_APP.SHEETS.PARTICIPANTS);if(participants)participants.getRange(3,1,15,19).getValues().forEach((p,index)=>{const name=String(p[0]||'').trim(),surname=String(p[1]||'').trim(),role=eventNormalize_(p[7]),passed=Number(p[10]||0);if(role==='TECNICO'||((!name&&!surname)||!(passed>0)))return;const participantId=String(p[13]||'').trim()||('ROW-'+(index+3)),id='RIMBORSO-AUTO-'+participantId,old=oldById[id]?oldById[id].slice(0,26):new Array(26).fill(''),beneficiary=[name,surname].filter(Boolean).join(' ');old[0]=id;old[1]=eventId;old[2]='RIMBORSO';old[3]='RIMBORSO';old[4]='CEB.002';old[5]='Rimborso '+beneficiary;old[6]=beneficiary;old[7]=String(p[14]||'').trim();old[8]=0;old[9]=passed;old[10]='';old[11]='RIMBORSATO';old[12]=p[11] instanceof Date?p[11]:now;old[13]='';old[14]='';old[15]='[MOVIMENTO=RIMBORSO] [AUTO_RIMBORSO='+participantId+']'+(String(p[12]||'').trim()?' '+String(p[12]||'').trim():'');old[16]=old[16]||now;old[17]=now;old[18]='NON NECESSARIO';old[20]='';old[21]=false;old[25]=old[25]||now;rows.push(old);});
  eventReplaceRowsForEvent_(backend,eventId,2,rows,26);SpreadsheetApp.flush();return rows.length;
}

function eventBackendPaidStatusV8_(movement){if(movement==='AFOR')return'PAGATO - AFOR';if(movement==='SALDO FATTURA')return'PAGATO - FATTURA';if(movement==='CARTA DI CREDITO')return'PAGATO CON CC';return'PAGATO';}
