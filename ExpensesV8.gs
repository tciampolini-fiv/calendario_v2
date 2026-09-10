const EXPENSES_V8 = Object.freeze({
  INPUT_ROW:9,
  PAYMENT_START_ROW:13,
  PAYMENT_END_ROW:16,
  IMPORT_CHECKBOX:'K7',
  HISTORY_HEADER_ROW:19,
  HISTORY_START_ROW:20,
  HISTORY_END_ROW:500,
  HISTORY_COLS:17,
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
  if (row>=EXPENSES_V8.PAYMENT_START_ROW && row<=EXPENSES_V8.PAYMENT_END_ROW && col>=2 && col<=6) {
    eventNormalizePaymentInputRowV8_(sheet,row);
    return;
  }
  if (row>=EXPENSES_V8.HISTORY_START_ROW && row<=EXPENSES_V8.HISTORY_END_ROW && col<=11) {
    eventNormalizeHistoryPaymentRowV8_(sheet,row);
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

function eventNormalizePaymentInputRowV8_(sheet,row) {
  const type=eventPaymentTypeV8_(sheet.getRange(row,2).getDisplayValue());
  if (!type) return;
  if (type!==sheet.getRange(row,2).getDisplayValue()) sheet.getRange(row,2).setValue(type);
  const budget=Number(sheet.getRange('A9').getValue()||0);
  let amount=Number(sheet.getRange(row,3).getValue()||0);
  let paid=sheet.getRange(row,4).getValue()===true;
  let paymentDate=sheet.getRange(row,5).getValue();
  if ((type==='CARTA DI CREDITO'||type==='SALDO FATTURA') && !(amount>0) && budget>0) {
    const otherPaid=eventPaidInputAmountV8_(sheet,row);
    amount=Math.max(budget-otherPaid,0);
    if (amount>0) sheet.getRange(row,3).setValue(amount).setNumberFormat('€ #,##0.00');
    paid=true;
    sheet.getRange(row,4).setValue(true);
  }
  if (paymentDate instanceof Date && !paid) {
    paid=true;
    sheet.getRange(row,4).setValue(true);
  }
  if (paid && !(paymentDate instanceof Date)) {
    paymentDate=new Date();
    sheet.getRange(row,5).setValue(paymentDate).setNumberFormat('dd/MM/yyyy');
  }
  SpreadsheetApp.flush();
}

function eventPaidInputAmountV8_(sheet,excludeRow) {
  let total=0;
  const rows=sheet.getRange(EXPENSES_V8.PAYMENT_START_ROW,2,4,3).getValues();
  rows.forEach((r,i)=>{if(EXPENSES_V8.PAYMENT_START_ROW+i!==excludeRow&&r[2]===true)total+=Number(r[1]||0);});
  return total;
}

function eventPaymentTypeV8_(value) {
  const t=eventNormalizeTextV7_(value);
  if (t==='FATTURA'||t==='PAGAMENTO FATTURA') return 'SALDO FATTURA';
  if (t==='CARTA'||t==='CC') return 'CARTA DI CREDITO';
  return EXPENSES_V8.PAYMENT_TYPES.includes(t)?t:'';
}

function eventNormalizeHistoryPaymentRowV8_(sheet,row) {
  const movement=eventNormalizeTextV7_(sheet.getRange(row,3).getDisplayValue());
  if (!movement) return;
  const status=eventNormalizeTextV7_(sheet.getRange(row,5).getDisplayValue());
  const dateCell=sheet.getRange(row,6),paymentDate=dateCell.getValue();
  if (status==='PAGATO' && movement!=='PREVENTIVO' && !(paymentDate instanceof Date)) dateCell.setValue(new Date()).setNumberFormat('dd/MM/yyyy');
}

function eventImportExpenseV8() {
  const ss=SpreadsheetApp.getActive(),sheet=ss.getSheetByName(EVENT_APP.SHEETS.EXPENSES);
  if (!sheet) throw new Error('Foglio Spese non trovato.');
  const shared=sheet.getRange(9,1,1,9).getValues()[0];
  const budget=Number(shared[0]||0),description=String(shared[1]||'').trim(),category=String(shared[2]||'').trim(),rif=String(shared[3]||'').trim();
  const meta=eventMeta_();
  const imp=String(shared[4]||'').trim()||eventResolveCommitmentV7_(meta)||String(meta.EVENT_COMMITMENT||'').trim();
  const documents=String(shared[5]||'').trim(),notes=String(shared[6]||'').trim();
  if (!(budget>0)) throw new Error('Inserisci il PREVENTIVO totale.');
  if (!description) throw new Error('La DESCRIZIONE è obbligatoria. Inserisci qui anche il nome del fornitore.');
  if (!category) throw new Error('Seleziona la TIPOLOGIA.');

  for(let row=13;row<=16;row++) eventNormalizePaymentInputRowV8_(sheet,row);
  const rawPayments=sheet.getRange(13,2,4,5).getValues(),payments=[];
  let allocated=0,paidTotal=0;
  rawPayments.forEach((r,i)=>{
    const type=eventPaymentTypeV8_(r[0]),amount=Number(r[1]||0),paid=r[2]===true,paymentDate=r[3] instanceof Date?r[3]:'',due=r[4] instanceof Date?r[4]:'';
    const hasAnything=!!type||amount>0||paid||!!paymentDate||!!due;
    if(!hasAnything)return;
    if(!type)throw new Error('Seleziona il tipo del pagamento n. '+(i+1)+'.');
    if(!(amount>0))throw new Error('Inserisci l’importo del pagamento n. '+(i+1)+'.');
    if(!paid&&!due)throw new Error('Il pagamento n. '+(i+1)+' non è pagato: inserisci una SCADENZA.');
    allocated+=amount;if(paid)paidTotal+=amount;
    payments.push({type:type,amount:amount,paid:paid,paymentDate:paymentDate,due:due,number:i+1});
  });
  if(allocated>budget+0.005)throw new Error('La somma dei 4 pagamenti supera il preventivo.');

  const target=eventFindHistoryBlockV8_(sheet,1+payments.length);
  if(!target)throw new Error('Elenco spese pieno.');
  const quoteId='PREV-'+Utilities.getUuid(),now=new Date();
  const quoteStatus=paidTotal+0.005>=budget?'PAGATO':'DA PAGARE';
  const rows=[[description,category,'PREVENTIVO',budget,quoteStatus,'','',rif,imp,documents,notes,quoteId,quoteId,'','',now,now]];
  payments.forEach(p=>rows.push([description,category,p.type,p.amount,p.paid?'PAGATO':'DA PAGARE',p.paid?p.paymentDate:'',p.due,rif,imp,documents,notes,'PAG-'+Utilities.getUuid(),quoteId,'','',now,now]));
  sheet.getRange(target,1,rows.length,EXPENSES_V8.HISTORY_COLS).setValues(rows);
  sheet.getRange(target,4,rows.length,1).setNumberFormat('€ #,##0.00');
  sheet.getRange(target,6,rows.length,2).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(target,9,rows.length,1).setNumberFormat('@');
  sheet.getRange('A9:G9').clearContent();
  sheet.getRange('B13:C16').clearContent();
  sheet.getRange('D13:D16').setValue(false);
  sheet.getRange('E13:F16').clearContent();
  sheet.getRange(EXPENSES_V8.IMPORT_CHECKBOX).setValue(false);
  eventApplyCommitmentV8_(meta);
  eventSyncExpenseTasksV8_();
  eventRefreshExpenseDashboardV8_();
  ss.toast('Preventivo registrato con '+payments.length+' pagamento/i nello storico.','Spese',5);
}

function eventFindHistoryBlockV8_(sheet,needed) {
  const values=sheet.getRange(EXPENSES_V8.HISTORY_START_ROW,1,EXPENSES_V8.HISTORY_END_ROW-EXPENSES_V8.HISTORY_START_ROW+1,1).getDisplayValues();
  let run=0;
  for(let i=0;i<values.length;i++){if(!String(values[i][0]||'').trim())run++;else run=0;if(run>=needed)return EXPENSES_V8.HISTORY_START_ROW+i-needed+1;}
  return 0;
}

function eventRefreshExpenseDashboardV8_() {
  const ss=SpreadsheetApp.getActive(),expense=ss.getSheetByName(EVENT_APP.SHEETS.EXPENSES),participants=ss.getSheetByName(EVENT_APP.SHEETS.PARTICIPANTS);
  if(!expense)return;
  const rows=expense.getRange(EXPENSES_V8.HISTORY_START_ROW,1,EXPENSES_V8.HISTORY_END_ROW-EXPENSES_V8.HISTORY_START_ROW+1,11).getValues();
  const buckets={VIAGGI:0,VITTO:0,ALLOGGIO:0,NOLEGGI:0,RIMBORSI:0,ALTRO:0};
  let forecast=0,paid=0;
  rows.forEach(r=>{
    if(!String(r[0]||'').trim())return;
    const movement=eventNormalizeTextV7_(r[2]),amount=Number(r[3]||0);
    if(movement==='PREVENTIVO'){
      forecast+=amount;
      const cat=eventNormalizeTextV7_(r[1]||'ALTRO'),bucket=['VIAGGI','VITTO','ALLOGGIO','NOLEGGI'].includes(cat)?cat:'ALTRO';
      buckets[bucket]+=amount;
    }else if(eventNormalizeTextV7_(r[4])==='PAGATO')paid+=amount;
  });
  if(participants){participants.getRange(3,1,15,19).getValues().forEach(r=>{
    const name=String(r[0]||'').trim(),surname=String(r[1]||'').trim();if(!name&&!surname)return;
    if(eventNormalizeTextV7_(r[7])==='TECNICO')return;
    const status=eventNormalizeTextV7_(r[8]),max=Number(r[9]||0),passed=Number(r[10]||0),current=status==='ASSENTE'?0:(passed>0?passed:max);
    forecast+=current;buckets.RIMBORSI+=current;if(passed>0)paid+=passed;
  });}
  expense.getRange('D2').setValue(forecast).setNumberFormat('€ #,##0.00');
  expense.getRange('F2').setValue(paid).setNumberFormat('€ #,##0.00');
  expense.getRange('H2').setValue(Math.max(forecast-paid,0)).setNumberFormat('€ #,##0.00');
  expense.getRange('A5:F5').setValues([[buckets.VIAGGI,buckets.VITTO,buckets.ALLOGGIO,buckets.NOLEGGI,buckets.RIMBORSI,buckets.ALTRO]]).setNumberFormat('€ #,##0.00');
}

function eventSyncExpenseTasksV8_() {
  const ss=SpreadsheetApp.getActive(),expense=ss.getSheetByName(EVENT_APP.SHEETS.EXPENSES),tasks=ss.getSheetByName(EVENT_APP.SHEETS.TASKS);
  if(!expense||!tasks)return;
  const rows=expense.getRange(EXPENSES_V8.HISTORY_START_ROW,1,EXPENSES_V8.HISTORY_END_ROW-EXPENSES_V8.HISTORY_START_ROW+1,17).getValues();
  let hasAfor=false;
  rows.forEach(r=>{
    const description=String(r[0]||'').trim();if(!description)return;
    const movement=eventPaymentTypeV8_(r[2]);if(movement!=='AFOR')return;
    hasAfor=true;
    const status=eventNormalizeTextV7_(r[4]),paidDate=r[5] instanceof Date?r[5]:null,due=r[6] instanceof Date?r[6]:null,movementId=String(r[11]||'').trim();
    if(!movementId)return;
    const payKey='AFOR_PAY:'+movementId,contKey='AFOR_CONT:'+movementId;
    if(status==='DA PAGARE'&&due){
      eventEnsureAutoTaskV8_(tasks,'Pagare AFOR',due,eventDateReachedV9_(due)?'DA FARE':'IN ATTESA',payKey,description);
    }else if(status==='PAGATO'){
      eventMarkAutoTaskDoneV8_(tasks,payKey);
      if(paidDate){const contDue=new Date(paidDate);contDue.setHours(12,0,0,0);contDue.setDate(contDue.getDate()+1);eventEnsureAutoTaskV8_(tasks,'Inviare contabile AFOR',contDue,eventDateReachedV9_(contDue)?'DA FARE':'IN ATTESA',contKey,description);}
    }
  });
  const end=eventParseIsoDate_(eventMeta_().EVENT_END);
  if(hasAfor&&end)eventEnsureAutoTaskV8_(tasks,'Richiedere fattura',end,eventDateReachedV9_(end)?'DA FARE':'IN ATTESA','FATTURA_AFOR_EVENTO','Evento');
}

function eventAutoTaskMarkerV8_(key){return '[AUTO_KEY='+key+']';}
function eventFindAutoTaskRowV8_(sheet,key){const marker=eventAutoTaskMarkerV8_(key),notes=sheet.getRange(2,2,499,1).getDisplayValues();for(let i=0;i<notes.length;i++)if(String(notes[i][0]||'').indexOf(marker)>=0)return i+2;return 0;}
function eventEnsureAutoTaskV8_(sheet,description,due,status,key,context){
  let row=eventFindAutoTaskRowV8_(sheet,key);
  if(!row){const descs=sheet.getRange(2,1,499,1).getDisplayValues(),idx=descs.findIndex(r=>!String(r[0]||'').trim());if(idx<0)return 0;row=idx+2;sheet.getRange(row,1).setValue(description);sheet.getRange(row,2).setValue(eventAutoTaskMarkerV8_(key)+(context?' '+context:''));eventPrepareTaskRowV9_(sheet,row);}
  if(due instanceof Date)sheet.getRange(row,6).setValue(due).setNumberFormat('dd/MM/yyyy');
  if(eventNormalizeTextV7_(sheet.getRange(row,5).getDisplayValue())!=='FATTO')sheet.getRange(row,5).setValue(status||'IN ATTESA');
  return row;
}
function eventMarkAutoTaskDoneV8_(sheet,key){const row=eventFindAutoTaskRowV8_(sheet,key);if(row>0)sheet.getRange(row,5).setValue('FATTO');}

function eventSaveExpensesV8_(master,child,eventId,meta) {
  const local=child.getSheetByName(EVENT_APP.SHEETS.EXPENSES);if(!local)throw new Error('Foglio Spese non trovato nella Scheda evento.');
  eventApplyCommitmentV8_(meta);eventSyncExpenseTasksV8_();eventRefreshExpenseDashboardV8_();
  const backend=eventMasterSheet_(master,'_SPESE'),oldRows=backend.getDataRange().getValues(),oldById={};
  for(let i=1;i<oldRows.length;i++)if(String(oldRows[i][1]||'')===String(eventId)&&oldRows[i][0])oldById[String(oldRows[i][0])]=oldRows[i];
  const history=local.getRange(20,1,481,17).getValues(),now=new Date(),quoteTotals={};
  history.forEach(r=>{if(eventNormalize_(r[2])==='PREVENTIVO'&&r[12])quoteTotals[String(r[12])]={budget:Number(r[3]||0),paid:0};});
  history.forEach(r=>{const q=String(r[12]||'');if(q&&eventNormalize_(r[2])!=='PREVENTIVO'&&eventNormalize_(r[4])==='PAGATO'&&quoteTotals[q])quoteTotals[q].paid+=Number(r[3]||0);});
  const rows=[];
  history.forEach((r,index)=>{
    const description=String(r[0]||'').trim();if(!description)return;
    const category=eventExpenseCategoryV7_(r[1],description),movement=eventNormalize_(r[2]),amount=Number(r[3]||0),localStatus=eventNormalize_(r[4]),paid=localStatus==='PAGATO';
    const paidDate=paid&&r[5] instanceof Date?r[5]:'',due=!paid&&r[6] instanceof Date?r[6]:'',userRif=String(r[7]||'').trim(),imp=String(r[8]||'').trim(),documents=String(r[9]||'').trim(),note=String(r[10]||'').trim();
    const id=String(r[11]||'').trim()||(movement==='PREVENTIVO'?'PREV-':'PAG-')+Utilities.getUuid(),quoteId=String(r[12]||'').trim()||(movement==='PREVENTIVO'?id:'');
    const old=oldById[id]?oldById[id].slice(0,26):new Array(26).fill(''),ceb=eventResolveCeb_(master,meta,category,String(r[13]||'').trim()||old[4]),createdAt=r[15] instanceof Date?r[15]:(old[16] instanceof Date?old[16]:now),qInfo=quoteTotals[quoteId];
    old[0]=id;old[1]=eventId;old[2]=movement==='PREVENTIVO'?'PREVENTIVO':'PAGAMENTO';old[3]=category;old[4]=ceb;old[5]=description;old[6]='';old[7]=String(r[14]||'').trim();
    old[8]=movement==='PREVENTIVO'?amount:0;old[9]=movement!=='PREVENTIVO'&&paid?amount:0;old[10]=due;
    old[11]=movement==='PREVENTIVO'?(qInfo&&qInfo.paid+0.005>=qInfo.budget?'SALDATO':'DA SALDARE'):(paid?eventBackendPaidStatusV8_(movement):'DA PAGARE');
    old[12]=paidDate;old[13]=quoteId;old[14]=documents;old[15]='[MOVIMENTO='+movement+'] [ID_PREVENTIVO='+quoteId+']'+(imp?' [IMP='+imp+']':'')+(note?' '+note:'');old[16]=createdAt;old[17]=now;
    if(userRif)old[18]='RICEVUTO';else if(movement==='PREVENTIVO'&&amount>1000)old[18]=old[18]||'DA RICHIEDERE';else old[18]=old[18]||'NON NECESSARIO';
    old[20]=userRif;old[21]=movement==='SALDO FATTURA'&&paid;old[22]=old[21]?paidDate:(old[22]||'');old[25]=paid&&movement!=='PREVENTIVO'?(old[25]||now):(movement==='PREVENTIVO'&&qInfo&&qInfo.paid+0.005>=qInfo.budget?(old[25]||now):'');rows.push(old);
    if(!String(r[11]||'').trim())local.getRange(20+index,12).setValue(id);if(!String(r[12]||'').trim())local.getRange(20+index,13).setValue(quoteId);if(!String(r[13]||'').trim()&&ceb)local.getRange(20+index,14).setValue(ceb);local.getRange(20+index,16).setValue(createdAt);local.getRange(20+index,17).setValue(now);
  });
  const participants=child.getSheetByName(EVENT_APP.SHEETS.PARTICIPANTS);
  if(participants)participants.getRange(3,1,15,19).getValues().forEach((p,index)=>{
    const name=String(p[0]||'').trim(),surname=String(p[1]||'').trim(),role=eventNormalize_(p[7]),passed=Number(p[10]||0);if(role==='TECNICO'||((!name&&!surname)||!(passed>0)))return;
    const participantId=String(p[13]||'').trim()||('ROW-'+(index+3)),id='RIMBORSO-AUTO-'+participantId,old=oldById[id]?oldById[id].slice(0,26):new Array(26).fill(''),beneficiary=[name,surname].filter(Boolean).join(' ');
    old[0]=id;old[1]=eventId;old[2]='RIMBORSO';old[3]='RIMBORSO';old[4]='CEB.002';old[5]='Rimborso '+beneficiary;old[6]=beneficiary;old[7]=String(p[14]||'').trim();old[8]=0;old[9]=passed;old[10]='';old[11]='RIMBORSATO';old[12]=p[11] instanceof Date?p[11]:now;old[13]='';old[14]='';old[15]='[MOVIMENTO=RIMBORSO] [AUTO_RIMBORSO='+participantId+']'+(String(p[12]||'').trim()?' '+String(p[12]||'').trim():'');old[16]=old[16]||now;old[17]=now;old[18]='NON NECESSARIO';old[20]='';old[21]=false;old[25]=old[25]||now;rows.push(old);
  });
  eventReplaceRowsForEvent_(backend,eventId,2,rows,26);SpreadsheetApp.flush();return rows.length;
}

function eventBackendPaidStatusV8_(movement){if(movement==='AFOR')return'PAGATO - AFOR';if(movement==='SALDO FATTURA')return'PAGATO - FATTURA';if(movement==='CARTA DI CREDITO')return'PAGATO CON CC';return'PAGATO';}
