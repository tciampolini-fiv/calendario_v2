const EXPENSES_V8 = Object.freeze({
  INPUT_ROW: 9,
  PAYMENT_START_ROW: 13,
  PAYMENT_END_ROW: 16,
  IMPORT_CHECKBOX: 'K7',
  HISTORY_HEADER_ROW: 19,
  HISTORY_START_ROW: 20,
  HISTORY_COLS: 18,
  VISIBLE_HISTORY_COLS: 11,
  PAYMENT_TYPES: Object.freeze(['AFOR','CARTA DI CREDITO','SALDO FATTURA','ALTRO PAGAMENTO']),
  CATEGORIES: Object.freeze(['VIAGGI','VITTO','ALLOGGIO','NOLEGGI','ISCRIZIONI','ALTRO'])
});

const EXPENSE_COLS_V11 = Object.freeze({
  input: Object.freeze({quoteNo:1,amount:3,description:4,category:5,rif:6,imp:7,documents:8,notes:9,paid:10,due:11}),
  payment: Object.freeze({quoteNo:1,paymentNo:2,amount:3,description:4,type:5,paid:6,paymentDate:7,due:8}),
  history: Object.freeze({quoteNo:1,movement:2,amount:3,description:4,status:5,paymentDate:6,due:7,rif:8,imp:9,documents:10,notes:11,category:12,movementId:13,quoteId:14,ceb:15,personId:16,createdAt:17,updatedAt:18})
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
  if (row===EXPENSES_V8.INPUT_ROW && col===EXPENSE_COLS_V11.input.quoteNo) {
    eventSyncPaymentQuoteInputsV11_(sheet,true);
    return;
  }
  if (row>=EXPENSES_V8.PAYMENT_START_ROW && row<=EXPENSES_V8.PAYMENT_END_ROW && col>=1 && col<=8) {
    eventNormalizePaymentInputRowV8_(sheet,row);
    return;
  }
  if (row>=EXPENSES_V8.HISTORY_START_ROW && col<=EXPENSES_V8.VISIBLE_HISTORY_COLS) {
    eventNormalizeHistoryPaymentRowV8_(sheet,row);
    eventRecalculateQuoteStatusesV10_(sheet);
    eventSortHistoryV10_(sheet);
    eventSyncPaymentQuoteInputsV11_(sheet,false);
    eventSyncExpenseTasksV8_();
    eventRefreshExpenseDashboardV8_();
  }
}

function eventEnsureExpenseLayoutV11_(sheet) {
  if(!sheet)return;
  if(String(sheet.getRange('B12').getDisplayValue()||'').trim()!=='N. PAGAMENTO')sheet.getRange('B12').setValue('N. PAGAMENTO');
  sheet.getRange('L19:R19').setValues([['TIPOLOGIA','ID MOVIMENTO','ID PREVENTIVO','CEB','ID PERSONA','DATA INSERIMENTO','ULTIMO AGGIORNAMENTO']]);
  try{sheet.hideColumns(12,7);}catch(e){}

  const jFormula='=IF(C9="";0;SUMIFS($C$13:$C$16;$F$13:$F$16;TRUE;$A$13:$A$16;$A$9))';
  const kFormula='=IF(C9="";0;MAX(C9-J9;0))';
  if(sheet.getRange('J9').getFormula()!==jFormula)sheet.getRange('J9').setFormula(jFormula).setNumberFormat('€ #,##0.00');
  if(sheet.getRange('K9').getFormula()!==kFormula)sheet.getRange('K9').setFormula(kFormula).setNumberFormat('€ #,##0.00');

  if(!String(sheet.getRange('A9').getDisplayValue()||'').trim()&&!eventHasPaymentDraftV11_(sheet))eventRestoreNextQuoteFormulaV11_(sheet);
  sheet.getRange('B13:B16').setValues([[1],[2],[3],[4]]).setNumberFormat('0');

  const typeRule=SpreadsheetApp.newDataValidation().requireValueInList(EXPENSES_V8.PAYMENT_TYPES,true).setAllowInvalid(false).build();
  sheet.getRange('E13:E16').setDataValidation(typeRule);
  sheet.getRange('F13:F16').insertCheckboxes();
  sheet.getRange('C13:C16').setNumberFormat('€ #,##0.00');
  sheet.getRange('G13:H16').setNumberFormat('dd/MM/yyyy');
  sheet.getRange('C9').setNumberFormat('€ #,##0.00');
  sheet.getRange('A9').setNumberFormat('0');

  eventEnsureHistoryValidationsV11_(sheet);
  eventSyncPaymentQuoteInputsV11_(sheet,false);
}

function eventHasPaymentDraftV11_(sheet){
  const rows=sheet.getRange(EXPENSES_V8.PAYMENT_START_ROW,1,EXPENSES_V8.PAYMENT_END_ROW-EXPENSES_V8.PAYMENT_START_ROW+1,8).getValues();
  return rows.some(r=>Number(r[0]||0)>0||Number(r[2]||0)>0||String(r[3]||'').trim()||String(r[4]||'').trim()||r[5]===true||r[6] instanceof Date||r[7] instanceof Date);
}

function eventRestoreNextQuoteFormulaV11_(sheet){
  sheet.getRange('A9').setFormula('=IFERROR(MAX(FILTER($A$20:$A;$B$20:$B="PREVENTIVO"))+1;1)').setNumberFormat('0');
}

function eventExistingQuoteNumbersV11_(sheet){
  const count=Math.max(0,sheet.getMaxRows()-EXPENSES_V8.HISTORY_START_ROW+1);
  if(!count)return[];
  const rows=sheet.getRange(EXPENSES_V8.HISTORY_START_ROW,1,count,2).getValues();
  const set=new Set();
  rows.forEach(r=>{const n=Number(r[0]||0);if(n>0&&eventNormalizeTextV7_(r[1])==='PREVENTIVO')set.add(n);});
  return Array.from(set).sort((a,b)=>a-b);
}

function eventSyncPaymentQuoteInputsV11_(sheet,clearWhenExistingMode){
  const q=Number(sheet.getRange('A9').getValue()||0),range=sheet.getRange('A13:A16');
  if(q>0){
    range.clearDataValidations();
    range.setValues([[q],[q],[q],[q]]).setNumberFormat('0');
    return;
  }
  if(clearWhenExistingMode)range.clearContent();
  const numbers=eventExistingQuoteNumbersV11_(sheet);
  if(numbers.length){
    const rule=SpreadsheetApp.newDataValidation().requireValueInList(numbers.map(String),true).setAllowInvalid(false).build();
    range.setDataValidation(rule).setNumberFormat('0');
  }else range.clearDataValidations();
}

function eventEnsureHistoryValidationsV11_(sheet){
  const count=Math.max(0,sheet.getMaxRows()-EXPENSES_V8.HISTORY_START_ROW+1);if(!count)return;
  const movementRule=SpreadsheetApp.newDataValidation().requireValueInList(['PREVENTIVO'].concat(EXPENSES_V8.PAYMENT_TYPES),true).setAllowInvalid(false).build();
  const statusRule=SpreadsheetApp.newDataValidation().requireValueInList(['DA PAGARE','PARZIALMENTE PAGATO','PAGATO'],true).setAllowInvalid(false).build();
  sheet.getRange(EXPENSES_V8.HISTORY_START_ROW,2,count,1).setDataValidation(movementRule);
  sheet.getRange(EXPENSES_V8.HISTORY_START_ROW,5,count,1).setDataValidation(statusRule).setNumberFormat('@');
  sheet.getRange(EXPENSES_V8.HISTORY_START_ROW,3,count,1).setNumberFormat('€ #,##0.00');
  sheet.getRange(EXPENSES_V8.HISTORY_START_ROW,6,count,2).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(EXPENSES_V8.HISTORY_START_ROW,8,count,2).setNumberFormat('@');
}

function eventApplyCommitmentV8_(meta) {
  const imp=eventResolveCommitmentV7_(meta)||String(meta.EVENT_COMMITMENT||'').trim();
  const sheet=SpreadsheetApp.getActive().getSheetByName(EVENT_APP.SHEETS.EXPENSES);
  if(sheet)eventEnsureExpenseLayoutV11_(sheet);
  if (!imp) return '';
  eventWriteLocalMetaV7_('EVENT_COMMITMENT',imp);
  if (sheet && !String(sheet.getRange('G9').getDisplayValue()||'').trim()) sheet.getRange('G9').setValue(imp).setNumberFormat('@');
  return imp;
}

/** Checkbox PAGATO = unica fonte di verita. */
function eventNormalizePaymentInputRowV8_(sheet,row) {
  const c=EXPENSE_COLS_V11.payment,newQuote=Number(sheet.getRange('A9').getValue()||0);
  if(newQuote>0&&Number(sheet.getRange(row,c.quoteNo).getValue()||0)!==newQuote)sheet.getRange(row,c.quoteNo).setValue(newQuote);
  const quoteNo=Number(sheet.getRange(row,c.quoteNo).getValue()||0);
  const typeCell=sheet.getRange(row,c.type),type=eventPaymentTypeV8_(typeCell.getDisplayValue());
  if(type&&type!==typeCell.getDisplayValue())typeCell.setValue(type);

  const amountCell=sheet.getRange(row,c.amount);let amount=Number(amountCell.getValue()||0);
  const paid=sheet.getRange(row,c.paid).getValue()===true,dateCell=sheet.getRange(row,c.paymentDate),paymentDate=dateCell.getValue();

  if((type==='CARTA DI CREDITO'||type==='SALDO FATTURA')&&!(amount>0)){
    const quote=quoteNo>0?eventFindQuoteByNumberV10_(sheet,quoteNo):null;
    const budget=quote?quote.amount:Number(sheet.getRange('C9').getValue()||0);
    if(budget>0){
      const historyAllocated=quote?eventAllocatedHistoryForQuoteV10_(sheet,quoteNo):0;
      const inputAllocated=eventAllocatedInputAmountV10_(sheet,row,quoteNo);
      amount=Math.max(budget-historyAllocated-inputAllocated,0);
      if(amount>0)amountCell.setValue(amount).setNumberFormat('€ #,##0.00');
    }
  }

  if(paid){
    if(!(paymentDate instanceof Date))dateCell.setValue(new Date()).setNumberFormat('dd/MM/yyyy');
  }else if(paymentDate instanceof Date||String(dateCell.getDisplayValue()||'').trim())dateCell.clearContent();
  SpreadsheetApp.flush();
}

function eventAllocatedInputAmountV10_(sheet,excludeRow,quoteNo) {
  let total=0;
  const rows=sheet.getRange(EXPENSES_V8.PAYMENT_START_ROW,1,EXPENSES_V8.PAYMENT_END_ROW-EXPENSES_V8.PAYMENT_START_ROW+1,3).getValues();
  rows.forEach((r,i)=>{
    const row=EXPENSES_V8.PAYMENT_START_ROW+i;if(row===excludeRow)return;
    const q=Number(r[0]||0);if((quoteNo>0&&q!==quoteNo)||(quoteNo<=0&&q>0))return;
    total+=Number(r[2]||0);
  });
  return total;
}

function eventHistoryRowsV11_(sheet,width){
  const count=Math.max(0,sheet.getMaxRows()-EXPENSES_V8.HISTORY_START_ROW+1);
  return count?sheet.getRange(EXPENSES_V8.HISTORY_START_ROW,1,count,width||EXPENSES_V8.HISTORY_COLS).getValues():[];
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
  if(row>sheet.getMaxRows())return;
  const movement=eventNormalizeTextV7_(sheet.getRange(row,2).getDisplayValue());
  if(!movement||movement==='PREVENTIVO')return;
  const statusCell=sheet.getRange(row,5);let status=eventNormalizeTextV7_(statusCell.getDisplayValue());
  if(status==='PARZIALMENTE PAGATO'){status='DA PAGARE';statusCell.setValue(status);}
  const dateCell=sheet.getRange(row,6),paymentDate=dateCell.getValue();
  if(status==='PAGATO'){
    if(!(paymentDate instanceof Date))dateCell.setValue(new Date()).setNumberFormat('dd/MM/yyyy');
  }else if(paymentDate instanceof Date||String(dateCell.getDisplayValue()||'').trim())dateCell.clearContent();
}

function eventImportExpenseV8(){
  const ss=SpreadsheetApp.getActive(),sheet=ss.getSheetByName(EVENT_APP.SHEETS.EXPENSES);
  if(!sheet)throw new Error('Foglio Spese non trovato.');
  eventEnsureExpenseLayoutV11_(sheet);

  const shared=sheet.getRange(9,1,1,9).getValues()[0];
  const quoteNoInput=Number(shared[0]||0),budget=Number(shared[2]||0),description=String(shared[3]||'').trim(),category=String(shared[4]||'').trim(),rif=String(shared[5]||'').trim();
  const meta=eventMeta_();
  const imp=String(shared[6]||'').trim()||eventResolveCommitmentV7_(meta)||String(meta.EVENT_COMMITMENT||'').trim();
  const documents=String(shared[7]||'').trim(),notes=String(shared[8]||'').trim();
  const creatingQuote=budget>0||description||category||rif||documents||notes;

  if(creatingQuote){
    if(!(budget>0))throw new Error('Inserisci l IMPORTO totale del preventivo.');
    if(!description)throw new Error('La DESCRIZIONE del preventivo e obbligatoria.');
    if(!category)throw new Error('Seleziona la TIPOLOGIA.');
  }

  const nextQuoteNo=eventNextQuoteNumberV10_(sheet);
  const newQuoteNo=creatingQuote?(quoteNoInput>0?quoteNoInput:nextQuoteNo):0;
  if(creatingQuote&&eventFindQuoteByNumberV10_(sheet,newQuoteNo))throw new Error('Il preventivo n. '+newQuoteNo+' esiste gia. Usa i PAGAMENTI per aggiungere un movimento a quel preventivo.');

  for(let row=EXPENSES_V8.PAYMENT_START_ROW;row<=EXPENSES_V8.PAYMENT_END_ROW;row++)eventNormalizePaymentInputRowV8_(sheet,row);
  const rawPayments=sheet.getRange(EXPENSES_V8.PAYMENT_START_ROW,1,EXPENSES_V8.PAYMENT_END_ROW-EXPENSES_V8.PAYMENT_START_ROW+1,8).getValues(),payments=[];

  rawPayments.forEach((r,i)=>{
    const requestedQuoteNo=Number(r[0]||0),paymentNo=Number(r[1]||i+1),amount=Number(r[2]||0),paymentDescription=String(r[3]||'').trim(),type=eventPaymentTypeV8_(r[4]),paid=r[5]===true,paymentDate=paid&&r[6] instanceof Date?r[6]:'',due=r[7] instanceof Date?r[7]:'';
    const hasAnything=amount>0||!!paymentDescription||!!type||paid||r[6] instanceof Date||!!due;
    if(!hasAnything)return;
    if(!type)throw new Error('Seleziona il TIPO PAGAMENTO per il pagamento n. '+paymentNo+'.');
    if(!(amount>0))throw new Error('Inserisci l IMPORTO del pagamento n. '+paymentNo+'.');
    if(!paid&&!due)throw new Error('Il pagamento n. '+paymentNo+' non e pagato: inserisci una SCADENZA.');
    const quoteNo=requestedQuoteNo>0?requestedQuoteNo:newQuoteNo;
    if(!(quoteNo>0))throw new Error('Nel pagamento n. '+paymentNo+' seleziona il N. PREVENTIVO.');
    const existing=eventFindQuoteByNumberV10_(sheet,quoteNo);
    if(!existing&&quoteNo!==newQuoteNo)throw new Error('Preventivo n. '+quoteNo+' non trovato nello storico.');
    payments.push({quoteNo:quoteNo,paymentNo:paymentNo,type:type,amount:amount,description:paymentDescription,paid:paid,paymentDate:paymentDate,due:due});
  });

  if(!creatingQuote&&!payments.length)throw new Error('Inserisci un preventivo oppure almeno un pagamento.');

  if(creatingQuote){
    const allocated=payments.filter(p=>p.quoteNo===newQuoteNo).reduce((s,p)=>s+p.amount,0);
    if(allocated>budget+0.005)throw new Error('La somma dei pagamenti del nuovo preventivo supera il suo importo.');
  }

  const byExistingQuote={};
  payments.forEach(p=>{if(p.quoteNo!==newQuoteNo)(byExistingQuote[p.quoteNo]||(byExistingQuote[p.quoteNo]=[])).push(p);});
  Object.keys(byExistingQuote).forEach(k=>{
    const q=eventFindQuoteByNumberV10_(sheet,Number(k)),historyAllocated=eventAllocatedHistoryForQuoteV10_(sheet,Number(k)),adding=byExistingQuote[k].reduce((s,p)=>s+p.amount,0);
    if(q&&historyAllocated+adding>q.amount+0.005)throw new Error('I pagamenti del preventivo n. '+k+' superano il suo importo.');
  });

  const append=[],now=new Date();let newQuoteId='';
  if(creatingQuote){
    newQuoteId='PREV-'+Utilities.getUuid();
    const paidTotal=payments.filter(p=>p.quoteNo===newQuoteNo&&p.paid).reduce((s,p)=>s+p.amount,0);
    append.push([newQuoteNo,'PREVENTIVO',budget,description,eventQuoteStatusV10_(budget,paidTotal),'','',rif,imp,documents,notes,category,newQuoteId,newQuoteId,'','',now,now]);
  }

  payments.forEach(p=>{
    const q=p.quoteNo===newQuoteNo?{id:newQuoteId,description:description,category:category,rif:rif,imp:imp,documents:documents,notes:notes}:eventFindQuoteByNumberV10_(sheet,p.quoteNo);
    const movementDescription=p.description||q.description;
    append.push([p.quoteNo,p.type,p.amount,movementDescription,p.paid?'PAGATO':'DA PAGARE',p.paid?p.paymentDate:'',p.due,q.rif,q.imp,q.documents,q.notes,q.category,'PAG-'+Utilities.getUuid(),q.id,'','',now,now]);
  });

  const target=eventFindHistoryBlockV8_(sheet,append.length);
  if(!target)throw new Error('Impossibile trovare spazio nello storico spese.');
  sheet.getRange(target,1,append.length,EXPENSES_V8.HISTORY_COLS).setValues(append);
  sheet.getRange(target,3,append.length,1).setNumberFormat('€ #,##0.00');
  sheet.getRange(target,6,append.length,2).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(target,9,append.length,1).setNumberFormat('@');

  sheet.getRange('A9:I9').clearContent();
  sheet.getRange('A13:H16').clearContent();
  sheet.getRange('B13:B16').setValues([[1],[2],[3],[4]]).setNumberFormat('0');
  sheet.getRange('F13:F16').setValue(false);
  sheet.getRange(EXPENSES_V8.IMPORT_CHECKBOX).setValue(false);
  eventRestoreNextQuoteFormulaV11_(sheet);
  eventSyncPaymentQuoteInputsV11_(sheet,false);
  eventApplyCommitmentV8_(meta);
  eventRecalculateQuoteStatusesV10_(sheet);
  eventSortHistoryV10_(sheet);
  eventSyncPaymentQuoteInputsV11_(sheet,false);
  eventSyncExpenseTasksV8_();
  eventRefreshExpenseDashboardV8_();
  ss.toast((creatingQuote?'Preventivo n. '+newQuoteNo+' e ':'')+payments.length+' pagamento/i registrati.','Spese',5);
}

function eventNextQuoteNumberV10_(sheet){
  let max=0;eventHistoryRowsV11_(sheet,2).forEach(r=>{if(eventNormalizeTextV7_(r[1])!=='PREVENTIVO')return;const n=Number(r[0]||0);if(Number.isFinite(n)&&n>max)max=n;});return max+1;
}

function eventFindQuoteByNumberV10_(sheet,quoteNo){
  if(!(quoteNo>0))return null;
  const rows=eventHistoryRowsV11_(sheet,EXPENSES_V8.HISTORY_COLS);
  for(let i=0;i<rows.length;i++){
    const r=rows[i];if(Number(r[0]||0)!==quoteNo||eventNormalizeTextV7_(r[1])!=='PREVENTIVO')continue;
    return{id:String(r[13]||'').trim(),number:quoteNo,amount:Number(r[2]||0),description:String(r[3]||'').trim(),category:String(r[11]||'').trim(),rif:String(r[7]||'').trim(),imp:String(r[8]||'').trim(),documents:String(r[9]||'').trim(),notes:String(r[10]||'').trim()};
  }
  return null;
}

function eventAllocatedHistoryForQuoteV10_(sheet,quoteNo){
  let total=0;eventHistoryRowsV11_(sheet,3).forEach(r=>{if(Number(r[0]||0)===quoteNo&&eventNormalizeTextV7_(r[1])!=='PREVENTIVO')total+=Number(r[2]||0);});return total;
}

function eventFindHistoryBlockV8_(sheet,needed){
  if(!(needed>0))return 0;
  for(let pass=0;pass<2;pass++){
    const count=Math.max(0,sheet.getMaxRows()-EXPENSES_V8.HISTORY_START_ROW+1),values=count?sheet.getRange(EXPENSES_V8.HISTORY_START_ROW,1,count,1).getDisplayValues():[];
    let run=0;
    for(let i=0;i<values.length;i++){
      if(!String(values[i][0]||'').trim())run++;else run=0;
      if(run>=needed)return EXPENSES_V8.HISTORY_START_ROW+i-needed+1;
    }
    eventExpandHistoryV11_(sheet,Math.max(50,needed+10));
  }
  return 0;
}

function eventExpandHistoryV11_(sheet,howMany){
  const add=Math.max(10,Number(howMany||50)),oldMax=sheet.getMaxRows();
  if(oldMax>=EXPENSES_V8.HISTORY_START_ROW+1)sheet.insertRowsBefore(oldMax,add);else sheet.insertRowsAfter(oldMax,add);
  eventEnsureHistoryValidationsV11_(sheet);
}

function eventSortHistoryV10_(sheet){
  const count=Math.max(0,sheet.getMaxRows()-EXPENSES_V8.HISTORY_START_ROW+1);if(!count)return;
  const range=sheet.getRange(EXPENSES_V8.HISTORY_START_ROW,1,count,EXPENSES_V8.HISTORY_COLS),all=range.getValues();
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
  const total=Math.max(1,sheet.getMaxRows()-EXPENSES_V8.HISTORY_START_ROW+1),full=sheet.getRange(EXPENSES_V8.HISTORY_START_ROW,1,total,EXPENSES_V8.VISIBLE_HISTORY_COLS);
  full.setBorder(false,false,false,false,false,false);
  if(!(rowCount>0))return;
  const vals=sheet.getRange(EXPENSES_V8.HISTORY_START_ROW,1,rowCount,2).getValues();let start=0;
  while(start<vals.length){
    const quoteNo=Number(vals[start][0]||0);let end=start;while(end+1<vals.length&&Number(vals[end+1][0]||0)===quoteNo)end++;
    const r=sheet.getRange(EXPENSES_V8.HISTORY_START_ROW+start,1,end-start+1,EXPENSES_V8.VISIBLE_HISTORY_COLS);
    r.setBorder(true,true,true,true,false,false,'#5f6368',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sheet.getRange(EXPENSES_V8.HISTORY_START_ROW+start,1,1,EXPENSES_V8.VISIBLE_HISTORY_COLS).setFontWeight('bold');
    if(end>start)sheet.getRange(EXPENSES_V8.HISTORY_START_ROW+start+1,1,end-start,EXPENSES_V8.VISIBLE_HISTORY_COLS).setFontWeight('normal');
    start=end+1;
  }
}

function eventRecalculateQuoteStatusesV10_(sheet){
  if(!sheet)return;
  const rows=eventHistoryRowsV11_(sheet,EXPENSES_V8.HISTORY_COLS),quotes={};
  rows.forEach((r,i)=>{if(eventNormalizeTextV7_(r[1])==='PREVENTIVO'){const no=Number(r[0]||0);if(no>0)quotes[no]={row:EXPENSES_V8.HISTORY_START_ROW+i,budget:Number(r[2]||0),paid:0};}});
  rows.forEach(r=>{const no=Number(r[0]||0);if(quotes[no]&&eventNormalizeTextV7_(r[1])!=='PREVENTIVO'&&eventNormalizeTextV7_(r[4])==='PAGATO')quotes[no].paid+=Number(r[2]||0);});
  Object.keys(quotes).forEach(no=>{const q=quotes[no];sheet.getRange(q.row,5).setValue(eventQuoteStatusV10_(q.budget,q.paid));});
}

function eventRefreshExpenseDashboardV8_(){
  const ss=SpreadsheetApp.getActive(),expense=ss.getSheetByName(EVENT_APP.SHEETS.EXPENSES),participants=ss.getSheetByName(EVENT_APP.SHEETS.PARTICIPANTS);if(!expense)return;
  eventEnsureExpenseLayoutV11_(expense);eventRecalculateQuoteStatusesV10_(expense);
  const rows=eventHistoryRowsV11_(expense,EXPENSES_V8.HISTORY_COLS),buckets={VIAGGI:0,VITTO:0,ALLOGGIO:0,NOLEGGI:0,RIMBORSI:0,ALTRO:0};let forecast=0,paid=0;
  rows.forEach(r=>{
    if(!Number(r[0]||0))return;const movement=eventNormalizeTextV7_(r[1]),amount=Number(r[2]||0);
    if(movement==='PREVENTIVO'){
      forecast+=amount;const cat=eventNormalizeTextV7_(r[11]||'ALTRO'),bucket=['VIAGGI','VITTO','ALLOGGIO','NOLEGGI'].includes(cat)?cat:'ALTRO';buckets[bucket]+=amount;
    }else if(eventNormalizeTextV7_(r[4])==='PAGATO')paid+=amount;
  });
  if(participants)participants.getRange(EVENT_APP.PARTICIPANTS.CONV_START,1,EVENT_APP.PARTICIPANTS.CONV_COUNT,19).getValues().forEach(r=>{const name=String(r[0]||'').trim(),surname=String(r[1]||'').trim();if(!name&&!surname)return;if(eventNormalizeTextV7_(r[7])==='TECNICO')return;const status=eventNormalizeTextV7_(r[8]),max=Number(r[9]||0),passed=Number(r[10]||0),current=status==='ASSENTE'?0:(passed>0?passed:max);forecast+=current;buckets.RIMBORSI+=current;if(passed>0)paid+=passed;});
  expense.getRange('D2').setValue(forecast).setNumberFormat('€ #,##0.00');
  expense.getRange('F2').setValue(paid).setNumberFormat('€ #,##0.00');
  expense.getRange('H2').setValue(Math.max(forecast-paid,0)).setNumberFormat('€ #,##0.00');
  expense.getRange('A5:F5').setValues([[buckets.VIAGGI,buckets.VITTO,buckets.ALLOGGIO,buckets.NOLEGGI,buckets.RIMBORSI,buckets.ALTRO]]).setNumberFormat('€ #,##0.00');
}

function eventFormatCurrencyV10_(amount){const n=Math.max(0,Number(amount||0)),parts=n.toFixed(2).split('.');parts[0]=parts[0].replace(/\B(?=(\d{3})+(?!\d))/g,'.');return'€ '+parts[0]+','+parts[1];}
function eventExpenseTaskNoteV10_(description,amount){return String(description||'').trim()+' — '+eventFormatCurrencyV10_(amount);}

function eventSyncExpenseTasksV8_(){
  const ss=SpreadsheetApp.getActive(),expense=ss.getSheetByName(EVENT_APP.SHEETS.EXPENSES),tasks=ss.getSheetByName(EVENT_APP.SHEETS.TASKS);if(!expense||!tasks)return;
  eventRecalculateQuoteStatusesV10_(expense);
  const rows=eventHistoryRowsV11_(expense,EXPENSES_V8.HISTORY_COLS),quotes={},quoteHasAfor={};
  rows.forEach(r=>{if(eventNormalizeTextV7_(r[1])==='PREVENTIVO'&&r[13])quotes[String(r[13])]={number:Number(r[0]||0),amount:Number(r[2]||0),description:String(r[3]||'').trim()};});
  rows.forEach(r=>{
    if(eventPaymentTypeV8_(r[1])!=='AFOR')return;
    const quoteId=String(r[13]||'').trim();if(quoteId)quoteHasAfor[quoteId]=true;
    const amount=Number(r[2]||0),description=String(r[3]||'').trim(),status=eventNormalizeTextV7_(r[4]),paidDate=r[5] instanceof Date?r[5]:null,due=r[6] instanceof Date?r[6]:null,movementId=String(r[12]||'').trim();if(!movementId)return;
    const noteText=eventExpenseTaskNoteV10_(description,amount),payKey='AFOR_PAY:'+movementId,contKey='AFOR_CONT:'+movementId;
    if(status==='DA PAGARE'&&due)eventEnsureAutoTaskV8_(tasks,'Pagare AFOR',due,eventDateReachedV9_(due)?'DA FARE':'IN ATTESA',payKey,noteText);
    else if(status==='PAGATO'){
      eventMarkAutoTaskDoneV8_(tasks,payKey);
      if(paidDate){const contDue=new Date(paidDate);contDue.setHours(12,0,0,0);contDue.setDate(contDue.getDate()+1);eventEnsureAutoTaskV8_(tasks,'Inviare contabile AFOR',contDue,eventDateReachedV9_(contDue)?'DA FARE':'IN ATTESA',contKey,noteText);}
    }
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
  eventEnsureExpenseLayoutV11_(local);eventApplyCommitmentV8_(meta);eventRecalculateQuoteStatusesV10_(local);eventSortHistoryV10_(local);eventSyncExpenseTasksV8_();eventRefreshExpenseDashboardV8_();
  const backend=eventMasterSheet_(master,'_SPESE'),oldRows=backend.getDataRange().getValues(),oldById={};for(let i=1;i<oldRows.length;i++)if(String(oldRows[i][1]||'')===String(eventId)&&oldRows[i][0])oldById[String(oldRows[i][0])]=oldRows[i];
  const history=eventHistoryRowsV11_(local,EXPENSES_V8.HISTORY_COLS),now=new Date(),quoteTotals={};
  history.forEach(r=>{if(eventNormalize_(r[1])==='PREVENTIVO'&&r[13])quoteTotals[String(r[13])]={budget:Number(r[2]||0),paid:0,quoteNo:Number(r[0]||0)};});
  history.forEach(r=>{const q=String(r[13]||'');if(q&&eventNormalize_(r[1])!=='PREVENTIVO'&&eventNormalize_(r[4])==='PAGATO'&&quoteTotals[q])quoteTotals[q].paid+=Number(r[2]||0);});
  const rows=[];
  history.forEach((r,index)=>{
    const quoteNo=Number(r[0]||0),movement=eventNormalize_(r[1]),amount=Number(r[2]||0),description=String(r[3]||'').trim();if(!description||!movement)return;
    const category=eventExpenseCategoryV7_(r[11],description),localStatus=eventNormalize_(r[4]),paid=movement!=='PREVENTIVO'&&localStatus==='PAGATO',paidDate=paid&&r[5] instanceof Date?r[5]:'',due=!paid&&r[6] instanceof Date?r[6]:'',userRif=String(r[7]||'').trim(),imp=String(r[8]||'').trim(),documents=String(r[9]||'').trim(),note=String(r[10]||'').trim();
    const id=String(r[12]||'').trim()||(movement==='PREVENTIVO'?'PREV-':'PAG-')+Utilities.getUuid(),quoteId=String(r[13]||'').trim()||(movement==='PREVENTIVO'?id:''),old=oldById[id]?oldById[id].slice(0,26):new Array(26).fill(''),ceb=eventResolveCeb_(master,meta,category,String(r[14]||'').trim()||old[4]),createdAt=r[16] instanceof Date?r[16]:(old[16] instanceof Date?old[16]:now),qInfo=quoteTotals[quoteId];
    old[0]=id;old[1]=eventId;old[2]=movement==='PREVENTIVO'?'PREVENTIVO':'PAGAMENTO';old[3]=category;old[4]=ceb;old[5]=description;old[6]='';old[7]=String(r[15]||'').trim();old[8]=movement==='PREVENTIVO'?amount:0;old[9]=movement!=='PREVENTIVO'&&paid?amount:0;old[10]=due;
    if(movement==='PREVENTIVO'){const qp=qInfo?Number(qInfo.paid||0):0,qb=qInfo?Number(qInfo.budget||amount):amount;old[11]=qp<=0.005?'DA SALDARE':(qp+0.005>=qb?'SALDATO':'PARZIALMENTE SALDATO');}else old[11]=paid?eventBackendPaidStatusV8_(movement):'DA PAGARE';
    old[12]=paidDate;old[13]='PREV. '+quoteNo;old[14]=documents;old[15]='[MOVIMENTO='+movement+'] [ID_PREVENTIVO='+quoteId+'] [N_PREVENTIVO='+quoteNo+']'+(imp?' [IMP='+imp+']':'')+(note?' '+note:'');old[16]=createdAt;old[17]=now;if(userRif)old[18]='RICEVUTO';else if(movement==='PREVENTIVO'&&amount>1000)old[18]=old[18]||'DA RICHIEDERE';else old[18]=old[18]||'NON NECESSARIO';old[20]=userRif;old[21]=movement==='SALDO FATTURA'&&paid;old[22]=old[21]?paidDate:(old[22]||'');old[25]=paid&&movement!=='PREVENTIVO'?(old[25]||now):(movement==='PREVENTIVO'&&qInfo&&qInfo.paid+0.005>=qInfo.budget?(old[25]||now):'');rows.push(old);
    const localRow=EXPENSES_V8.HISTORY_START_ROW+index;if(!String(r[12]||'').trim())local.getRange(localRow,13).setValue(id);if(!String(r[13]||'').trim())local.getRange(localRow,14).setValue(quoteId);if(!String(r[14]||'').trim()&&ceb)local.getRange(localRow,15).setValue(ceb);local.getRange(localRow,17).setValue(createdAt);local.getRange(localRow,18).setValue(now);
  });
  const participants=child.getSheetByName(EVENT_APP.SHEETS.PARTICIPANTS);if(participants)participants.getRange(EVENT_APP.PARTICIPANTS.CONV_START,1,EVENT_APP.PARTICIPANTS.CONV_COUNT,19).getValues().forEach((p,index)=>{const name=String(p[0]||'').trim(),surname=String(p[1]||'').trim(),role=eventNormalize_(p[7]),passed=Number(p[10]||0);if(role==='TECNICO'||((!name&&!surname)||!(passed>0)))return;const participantId=String(p[13]||'').trim()||('ROW-'+(index+EVENT_APP.PARTICIPANTS.CONV_START)),id='RIMBORSO-AUTO-'+participantId,old=oldById[id]?oldById[id].slice(0,26):new Array(26).fill(''),beneficiary=[name,surname].filter(Boolean).join(' ');old[0]=id;old[1]=eventId;old[2]='RIMBORSO';old[3]='RIMBORSO';old[4]='CEB.002';old[5]='Rimborso '+beneficiary;old[6]=beneficiary;old[7]=String(p[14]||'').trim();old[8]=0;old[9]=passed;old[10]='';old[11]='RIMBORSATO';old[12]=p[11] instanceof Date?p[11]:now;old[13]='';old[14]='';old[15]='[MOVIMENTO=RIMBORSO] [AUTO_RIMBORSO='+participantId+']'+(String(p[12]||'').trim()?' '+String(p[12]||'').trim():'');old[16]=old[16]||now;old[17]=now;old[18]='NON NECESSARIO';old[20]='';old[21]=false;old[25]=old[25]||now;rows.push(old);});
  eventReplaceRowsForEvent_(backend,eventId,2,rows,26);SpreadsheetApp.flush();return rows.length;
}

function eventBackendPaidStatusV8_(movement){if(movement==='AFOR')return'PAGATO - AFOR';if(movement==='SALDO FATTURA')return'PAGATO - FATTURA';if(movement==='CARTA DI CREDITO')return'PAGATO CON CC';return'PAGATO';}
