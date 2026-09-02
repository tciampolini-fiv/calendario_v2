const EVENT_SHEET = Object.freeze({
  SHEETS: Object.freeze({TASKS:'Attività', EXPENSES:'Spese', GUESTS:'Invitati', META:'_META'}),
  TASK_HEADERS: Object.freeze([
    'ID TASK','STATO','ATTIVITA','CATEGORIA','SCADENZA','NOTE','ORIGINE','AUTO KEY','ORDINE','DATA COMPLETAMENTO','ULTIMO AGGIORNAMENTO'
  ]),
  EXPENSE_HEADERS: Object.freeze([
    'DESCRIZIONE','TIPOLOGIA','FORNITORE / PERSONA','TIPO PAGAMENTO','IMPORTO','STATO',
    'RIFERIMENTO / NOTE','PROMEMORIA 1','SCADENZA 1','PROMEMORIA 2','SCADENZA 2',
    'ID SPESA','CEB','ID PERSONA','ID TASK PROM. 1','ID TASK PROM. 2','DATA PAGAMENTO','DATA INSERIMENTO','ULTIMO AGGIORNAMENTO'
  ]),
  GUEST_HEADERS: Object.freeze([
    'ID PARTECIPAZIONE','ID EVENTO','ID PERSONA','NOME','COGNOME','CIRCOLO','RUOLO','EMAIL OVERRIDE',
    'MAX RIMBORSO PERSONALE','PROVENIENZA','FILE ORIGINE','NOTE',
    'SPESE PAGATE','RIMBORSO DA PASSARE','RIMBORSO PASSATO','STATO RIMBORSO'
  ]),
  EXPENSE_CATEGORIES: Object.freeze(['VITTO / ALLOGGIO','VIAGGI','NOLEGGI','ISCRIZIONI','TRASPORTO','RIMBORSO','ALTRO']),
  EXPENSE_MOVEMENTS: Object.freeze(['PREVENTIVO','AFOR','FATTURA','CARTA DI CREDITO','ALTRO PAGAMENTO']),
  PAYMENT_STATES: Object.freeze(['DA PAGARE','PAGATO'])
});

function prepareEventSheetForSelectedEvent() {
  const event = selectedEvent_();
  const eventId = ensureEventId_(event);
  generateChecklistForEvent_(eventId, event);
  const folder = createWorkFolderForEvent_(eventId, event, event._row);
  const result = getOrCreateEventSheet_(eventId, event, folder.folderId);
  exportEventDataToEventSheet_(eventId, result.spreadsheet);
  setEventSheetLink_(event._row, result.spreadsheet.getUrl());
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    result.created ? 'Scheda evento creata' : 'Scheda evento aggiornata',
    'La scheda è nella cartella evento. Aprila dalla colonna SCHEDA EVENTO.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return {created:result.created, id:result.spreadsheet.getId(), url:result.spreadsheet.getUrl()};
}

function syncSelectedEventSheetToCalendar() {
  const event = selectedEvent_();
  const eventId = ensureEventId_(event);
  const folder = createWorkFolderForEvent_(eventId, event, event._row);
  const result = getOrCreateEventSheet_(eventId, event, folder.folderId);
  const child = result.spreadsheet;
  validateEventSheetIdentity_(child, eventId);

  const counts = {
    tasks: syncTasksFromEventSheet_(eventId, child),
    expenses: syncExpensesFromEventSheet_(eventId, child),
    guests: syncGuestsFromEventSheet_(eventId, child)
  };
  syncExpenseRemindersFromChild_(eventId, child);
  syncChecklistLocksForEvent_(eventId);
  exportEventDataToEventSheet_(eventId, child);
  setEventSheetLink_(event._row, child.getUrl());
  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    'Calendario aggiornato',
    'Attività: ' + counts.tasks + '\nSpese: ' + counts.expenses + '\nInvitati: ' + counts.guests,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return counts;
}

function getOrCreateEventSheet_(eventId, event, folderId) {
  const existing = findEventSheet_(eventId, event, folderId);
  if (existing) {
    ensureEventSheetStructure_(existing, eventId, folderId, event);
    return {spreadsheet:existing, created:false};
  }
  const child = SpreadsheetApp.create(buildEventSheetName_(event));
  child.setSpreadsheetLocale('it_IT');
  child.setSpreadsheetTimeZone(APP.TZ);
  DriveApp.getFileById(child.getId()).moveTo(DriveApp.getFolderById(folderId));
  ensureEventSheetStructure_(child, eventId, folderId, event);
  return {spreadsheet:child, created:true};
}

function findEventSheet_(eventId, event, folderId) {
  const cal = sh_(APP.SHEETS.CALENDAR);
  const map = headerMap_(cal);
  const linkCol = map[APP.CALENDAR_HEADERS.EVENT_SHEET];
  if (linkCol && event && event._row) {
    const rich = cal.getRange(event._row, linkCol).getRichTextValue();
    const link = rich && rich.getLinkUrl();
    const displayed = cal.getRange(event._row, linkCol).getDisplayValue();
    const id = extractDriveId_(link || displayed);
    if (id) {
      try { return SpreadsheetApp.openById(id); } catch (e) {}
    }
  }
  const files = DriveApp.getFolderById(folderId).getFilesByType(MimeType.GOOGLE_SHEETS);
  while (files.hasNext()) {
    const file = files.next();
    try {
      const ss = SpreadsheetApp.openById(file.getId());
      const meta = ss.getSheetByName(EVENT_SHEET.SHEETS.META);
      if (meta && String(readMetaValue_(meta, 'EVENT_ID')) === String(eventId)) return ss;
    } catch (e) {}
  }
  return null;
}

function ensureEventSheetStructure_(child, eventId, folderId, event) {
  child.setSpreadsheetLocale('it_IT');
  child.setSpreadsheetTimeZone(APP.TZ);
  let tasks = child.getSheetByName(EVENT_SHEET.SHEETS.TASKS);
  if (!tasks) {
    const first = child.getSheets()[0];
    if (child.getSheets().length === 1 && first.getLastRow() === 0) {
      first.setName(EVENT_SHEET.SHEETS.TASKS);
      tasks = first;
    } else tasks = child.insertSheet(EVENT_SHEET.SHEETS.TASKS, 0);
  }
  let expenses = child.getSheetByName(EVENT_SHEET.SHEETS.EXPENSES);
  if (!expenses) expenses = child.insertSheet(EVENT_SHEET.SHEETS.EXPENSES);
  let guests = child.getSheetByName(EVENT_SHEET.SHEETS.GUESTS);
  if (!guests) guests = child.insertSheet(EVENT_SHEET.SHEETS.GUESTS);
  let meta = child.getSheetByName(EVENT_SHEET.SHEETS.META);
  if (!meta) meta = child.insertSheet(EVENT_SHEET.SHEETS.META);

  migrateLegacyChildExpenseLayout_(expenses);
  prepareChildSheet_(tasks, EVENT_SHEET.TASK_HEADERS, 'TASKS');
  prepareChildSheet_(expenses, EVENT_SHEET.EXPENSE_HEADERS, 'EXPENSES');
  prepareChildSheet_(guests, EVENT_SHEET.GUEST_HEADERS, 'GUESTS');
  writeMeta_(meta, {
    EVENT_ID:eventId,
    MASTER_SPREADSHEET_ID:APP.SPREADSHEET_ID,
    EVENT_FOLDER_ID:folderId,
    EVENT_SHEET_ID:child.getId(),
    SYNC_VERSION:'2',
    EVENT_LABEL:buildEventSheetLabel_(event)
  });
  meta.hideSheet();
}

function migrateLegacyChildExpenseLayout_(sheet) {
  if (normalize_(sheet.getRange(1,1).getDisplayValue()) !== 'ID SPESA') return;
  if (sheet.getLastRow() > 1) sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getMaxColumns()).clearContent();
}

function prepareChildSheet_(sheet, headers, kind) {
  if (sheet.getMaxColumns() < headers.length) sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  sheet.getRange(1,1,1,headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#eeeeee').setVerticalAlignment('middle').setWrap(true);

  if (kind === 'TASKS') {
    sheet.showColumns(1, Math.min(11, sheet.getMaxColumns()));
    sheet.hideColumns(1);
    sheet.hideColumns(7,5);
    sheet.setColumnWidth(2,115); sheet.setColumnWidth(3,260); sheet.setColumnWidth(4,150); sheet.setColumnWidth(5,105); sheet.setColumnWidth(6,320);
    sheet.getRange('B2:B999').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['DA FARE','IN ATTESA','FATTO'],true).setAllowInvalid(false).build());
    sheet.getRange('E2:E999').setNumberFormat('dd/MM/yyyy');
    sheet.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$B2="FATTO"').setBackground('#d9ead3').setRanges([sheet.getRange('B2:F999')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$B2="IN ATTESA"').setBackground('#fff2cc').setRanges([sheet.getRange('B2:F999')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($B2="DA FARE",$E2<>"",$E2<TODAY())').setBackground('#f4cccc').setRanges([sheet.getRange('B2:F999')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($B2="DA FARE",OR($E2="",$E2>=TODAY()))').setBackground('#fce5cd').setRanges([sheet.getRange('B2:F999')]).build()
    ]);
    return;
  }

  if (kind === 'EXPENSES') {
    sheet.showColumns(1, Math.min(11, sheet.getMaxColumns()));
    if (sheet.getMaxColumns() > 11) sheet.hideColumns(12, sheet.getMaxColumns()-11);
    [245,150,190,155,105,105,280,245,105,245,105].forEach((w,i)=>sheet.setColumnWidth(i+1,w));
    sheet.getRange('B2:B1000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(EVENT_SHEET.EXPENSE_CATEGORIES,true).setAllowInvalid(false).build());
    sheet.getRange('D2:D1000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(EVENT_SHEET.EXPENSE_MOVEMENTS,true).setAllowInvalid(true).build());
    sheet.getRange('F2:F1000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(EVENT_SHEET.PAYMENT_STATES,true).setAllowInvalid(true).build());
    sheet.getRange('E2:E1000').setNumberFormat('€ #,##0.00');
    sheet.getRange('I2:I1000').setNumberFormat('dd/MM/yyyy');
    sheet.getRange('K2:K1000').setNumberFormat('dd/MM/yyyy');
    sheet.getRange('A2:K1000').setVerticalAlignment('top').setWrap(true);
    sheet.getRange('B1').setNote('La tipologia determina automaticamente il CEB, che resta nascosto.');
    sheet.getRange('D1').setNote('PREVENTIVO = totale stimato; AFOR/FATTURA/CARTA DI CREDITO = singolo pagamento. Per un rimborso scegli TIPOLOGIA = RIMBORSO.');
    sheet.getRange('F1').setNote('Usa DA PAGARE o PAGATO per i pagamenti. Sul PREVENTIVO può restare vuoto.');
    sheet.getRange('H1').setNote('Se compilato, al comando AGGIORNA CALENDARIO viene creata una task collegata alla spesa.');
    sheet.getRange('J1').setNote('Secondo promemoria opzionale della stessa spesa/pagamento.');
    sheet.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$F2="PAGATO"').setBackground('#d9ead3').setRanges([sheet.getRange('A2:K1000')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($H2<>"",$I2<>"",$I2<TODAY())').setBackground('#f4cccc').setRanges([sheet.getRange('H2:I1000')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($J2<>"",$K2<>"",$K2<TODAY())').setBackground('#f4cccc').setRanges([sheet.getRange('J2:K1000')]).build()
    ]);
    return;
  }

  if (kind === 'GUESTS') {
    sheet.showColumns(1, Math.min(headers.length,sheet.getMaxColumns()));
    sheet.hideColumns(1,3);
    sheet.hideColumns(10,2);
    sheet.setColumnWidth(4,140); sheet.setColumnWidth(5,140); sheet.setColumnWidth(6,220); sheet.setColumnWidth(7,110);
    sheet.setColumnWidth(8,230); sheet.setColumnWidth(9,135); sheet.setColumnWidth(12,250);
    sheet.setColumnWidth(13,120); sheet.setColumnWidth(14,135); sheet.setColumnWidth(15,130); sheet.setColumnWidth(16,125);
    sheet.getRange('G2:G1000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['ATLETA','TECNICO','ALTRO'],true).setAllowInvalid(false).build());
    sheet.getRange('M2:O1000').setNumberFormat('€ #,##0.00');
  }
}

function exportEventDataToEventSheet_(eventId, child) {
  const tasks = getChecklistRowsForEventRaw_(eventId).map(x=>[
    x.id, normalizeTaskStatusForSheet_(x.status), x.task||'', x.category||'', x.dueDate||'', x.note||'',
    x.source||'', x.autoKey||'', x.order||'', x.completedAt||'', x.updatedAt||''
  ]);

  const expenseSheet = child.getSheetByName(EVENT_SHEET.SHEETS.EXPENSES);
  const preserved = readChildExpenseState_(expenseSheet);
  const expenses = getExpensesForEvent_(eventId).map(x=>mapBackendExpenseToChild_(x,preserved[String(x.id)]||null));
  const guests = sh_(APP.SHEETS.PARTICIPANTS).getDataRange().getValues().slice(1)
    .filter(r=>String(r[1])===String(eventId)).map(r=>r.slice(0,12));

  replaceChildData_(child.getSheetByName(EVENT_SHEET.SHEETS.TASKS),tasks,EVENT_SHEET.TASK_HEADERS.length);
  replaceChildData_(expenseSheet,expenses,EVENT_SHEET.EXPENSE_HEADERS.length);
  replaceChildData_(child.getSheetByName(EVENT_SHEET.SHEETS.GUESTS),guests,12);
  refreshGuestSummaryFormulas_(child.getSheetByName(EVENT_SHEET.SHEETS.GUESTS),guests.length);
  updateExpenseBeneficiaryValidation_(expenseSheet,guests);
  writeMeta_(child.getSheetByName(EVENT_SHEET.SHEETS.META),{LAST_SYNC:Utilities.formatDate(new Date(),APP.TZ,'dd/MM/yyyy HH:mm')});
}

function replaceChildData_(sheet, rows, width) {
  const last = sheet.getLastRow();
  if (last>1) sheet.getRange(2,1,last-1,width).clearContent();
  if (rows.length) sheet.getRange(2,1,rows.length,width).setValues(rows);
}

function readChildExpenseState_(sheet) {
  const values = sheet.getDataRange().getValues();
  const out = {};
  for (let i=1;i<values.length;i++) {
    const r=values[i], id=String(r[11]||'').trim();
    if (!id) continue;
    out[id]={movement:r[3]||'', reminder1:r[7]||'', due1:r[8]||'', reminder2:r[9]||'', due2:r[10]||'', task1:r[14]||'', task2:r[15]||''};
  }
  return out;
}

function mapBackendExpenseToChild_(expense,preserved) {
  const movement = preserved&&preserved.movement ? normalizeExpenseMovement_(preserved.movement) : inferExpenseMovement_(expense);
  const isQuote = movement==='PREVENTIVO';
  const amount = isQuote ? Number(expense.budget||expense.actual||0) : Number(expense.actual||0);
  const category = normalize_(expense.type)==='RIMBORSO' ? 'RIMBORSO' : (expense.category||'ALTRO');
  const reference = expense.reference || stripMovementMarker_(expense.notes||'');
  return [
    expense.description||'', category, expense.beneficiary||'', movement, amount||'', childPaymentState_(expense,movement), reference,
    preserved?preserved.reminder1||'':'', preserved?preserved.due1||'':'', preserved?preserved.reminder2||'':'', preserved?preserved.due2||'':'',
    expense.id||'', expense.ceb||'', expense.personId||'', preserved?preserved.task1||'':'', preserved?preserved.task2||'':'',
    expense.paidDate||'', expense.createdAt||'', expense.updatedAt||''
  ];
}

function inferExpenseMovement_(expense) {
  const marker=extractMovementMarker_(expense.notes||'');
  if (marker) return marker;
  if (normalize_(expense.type)==='RIMBORSO') return 'RIMBORSO';
  const s=normalize_(expense.status);
  if (['PAGATO - AFOR','PAGAMENTO AFOR','AFOR FATTO'].includes(s)) return 'AFOR';
  if (['PAGATO - FATTURA','PAGAMENTO FATTURA'].includes(s)||expense.invoiceReceived===true) return 'FATTURA';
  if (s==='PAGATO CON CC') return 'CARTA DI CREDITO';
  if (Number(expense.budget||0)>0||s==='DA DEFINIRE') return 'PREVENTIVO';
  return 'ALTRO PAGAMENTO';
}

function extractMovementMarker_(notes) {
  const m=String(notes||'').match(/\[MOVIMENTO=([^\]]+)\]/i);
  return m?normalizeExpenseMovement_(m[1]):'';
}
function stripMovementMarker_(notes) { return String(notes||'').replace(/\s*\[MOVIMENTO=[^\]]+\]\s*/ig,' ').trim(); }
function normalizeExpenseMovement_(value) {
  const m=normalize_(value);
  if (EVENT_SHEET.EXPENSE_MOVEMENTS.includes(m)||m==='RIMBORSO') return m;
  if (m==='CC'||m==='CARTA') return 'CARTA DI CREDITO';
  if (m==='PAGAMENTO') return 'ALTRO PAGAMENTO';
  return m||'PREVENTIVO';
}
function childPaymentState_(expense,movement) {
  if (movement==='PREVENTIVO') return '';
  const s=normalize_(expense.status);
  return ['PAGATO - AFOR','PAGAMENTO AFOR','AFOR FATTO','PAGATO - FATTURA','PAGAMENTO FATTURA','PAGATO CON CC','CHIUSO','PAGATO','RIMBORSATO','INVIATO IN AMMINISTRAZIONE'].includes(s)?'PAGATO':'DA PAGARE';
}

function syncTasksFromEventSheet_(eventId,child) {
  const sheet=child.getSheetByName(EVENT_SHEET.SHEETS.TASKS), values=sheet.getDataRange().getValues(), now=new Date(), rows=[];
  for (let i=1;i<values.length;i++) {
    const r=values[i], task=String(r[2]||'').trim();
    if (!task) continue;
    const id=String(r[0]||'').trim()||'TASK-'+Utilities.getUuid();
    const status=normalizeTaskStatusForSheet_(r[1]||'DA FARE');
    const source=String(r[6]||'').trim()||'PERSONALIZZATA';
    let completed=r[9]||'';
    if (status==='FATTO'&&!(completed instanceof Date)) completed=now;
    if (status!=='FATTO') completed='';
    rows.push([id,eventId,Number(r[8]||(i*10)),task,r[3]||'',r[4]||'',status,'',source,r[7]||'',r[5]||'',completed,now]);
  }
  replaceCentralRowsForEvent_(sh_(APP.SHEETS.CHECKLIST),eventId,2,rows,13);
  return rows.length;
}

function syncExpensesFromEventSheet_(eventId,child) {
  const found=findCalendarEventById_(eventId);
  if (!found) throw new Error('Evento non trovato: '+eventId);
  const sheet=child.getSheetByName(EVENT_SHEET.SHEETS.EXPENSES), values=sheet.getDataRange().getValues();
  const oldRows=sh_(APP.SHEETS.EXPENSES).getDataRange().getValues(), oldById={};
  for (let i=1;i<oldRows.length;i++) if (String(oldRows[i][1])===String(eventId)&&oldRows[i][0]) oldById[String(oldRows[i][0])]=oldRows[i];
  const now=new Date(), centralRows=[], techRows=[];

  for (let i=1;i<values.length;i++) {
    const r=values[i], description=String(r[0]||'').trim(), amount=Number(r[4]||0);
    const hasData=description||amount||String(r[3]||'').trim()||String(r[7]||'').trim()||String(r[9]||'').trim();
    if (!hasData) continue;
    if (!description) throw new Error('Descrizione mancante nella riga '+(i+1)+' del foglio Spese.');

    const categoryInput=String(r[1]||'').trim()||'ALTRO';
    let movement=normalizeExpenseMovement_(r[3]||(normalize_(categoryInput)==='RIMBORSO'?'RIMBORSO':(amount?'PREVENTIVO':'ALTRO PAGAMENTO')));
    if (normalize_(categoryInput)==='RIMBORSO') movement='RIMBORSO';
    let id=String(r[11]||'').trim();
    if (!id) id='SPESA-'+Utilities.getUuid();
    const old=oldById[id]?oldById[id].slice(0,26):new Array(26).fill('');
    const recordType=movement==='RIMBORSO'?'RIMBORSO':'SPESA';
    const category=recordType==='RIMBORSO'?'RIMBORSO':categoryInput;
    const ceb=resolveExpenseCeb_(found.event,'',category,recordType);
    const beneficiary=String(r[2]||'').trim();
    const personId=String(r[13]||'').trim()||resolveExpensePersonId_(eventId,beneficiary);
    const paymentState=normalize_(r[5]);
    const paid=paymentState==='PAGATO'&&movement!=='PREVENTIVO';
    const budget=movement==='PREVENTIVO'?amount:0, actual=movement==='PREVENTIVO'?0:amount;
    const status=backendPaymentStatus_(movement,paymentState);
    const paidDate=paid?(r[16] instanceof Date?r[16]:(old[12] instanceof Date?old[12]:now)):'';
    const createdAt=r[17] instanceof Date?r[17]:(old[16] instanceof Date?old[16]:now);
    const reference=String(r[6]||'').trim();
    const legacyNotes=stripMovementMarker_(old[15]||'');
    const notes='[MOVIMENTO='+movement+']'+(legacyNotes?' '+legacyNotes:'');
    const rifNeeded=expenseNeedsRif_(budget,actual), rifCode=String(old[20]||'').trim();
    let rifStatus=old[18]||'';
    if (rifCode) rifStatus='RICEVUTO';
    else if (!rifNeeded) rifStatus='NON NECESSARIO';
    else if (!rifStatus||normalize_(rifStatus)==='NON NECESSARIO') rifStatus='DA RICHIEDERE';
    const invoiceReceived=movement==='FATTURA'||old[21]===true;
    const closed=['PAGATO - FATTURA','PAGATO CON CC','PAGATO','RIMBORSATO','CHIUSO'].includes(normalize_(status));

    old[0]=id; old[1]=eventId; old[2]=recordType; old[3]=category; old[4]=ceb; old[5]=description; old[6]=beneficiary; old[7]=personId;
    old[8]=budget; old[9]=actual; old[10]=''; old[11]=status; old[12]=paidDate; old[13]=reference; old[15]=notes; old[16]=createdAt; old[17]=now;
    old[18]=rifStatus; old[20]=rifCode; old[21]=invoiceReceived; old[22]=invoiceReceived?(old[22]||now):''; old[25]=closed?(old[25]||now):'';
    centralRows.push(old);
    techRows.push([id,ceb,personId,String(r[14]||'').trim(),String(r[15]||'').trim(),paidDate,createdAt,now]);
  }
  replaceCentralRowsForEvent_(sh_(APP.SHEETS.EXPENSES),eventId,2,centralRows,26);
  if (techRows.length) sheet.getRange(2,12,techRows.length,8).setValues(techRows);
  return centralRows.length;
}

function backendPaymentStatus_(movement,state) {
  const paid=normalize_(state)==='PAGATO';
  if (movement==='PREVENTIVO') return 'DA DEFINIRE';
  if (movement==='AFOR') return paid?'PAGATO - AFOR':'DA PAGARE';
  if (movement==='FATTURA') return paid?'PAGATO - FATTURA':'DA PAGARE';
  if (movement==='CARTA DI CREDITO') return paid?'PAGATO CON CC':'DA PAGARE';
  if (movement==='RIMBORSO') return paid?'RIMBORSATO':'DA PAGARE';
  return paid?'PAGATO':'DA PAGARE';
}

function resolveExpensePersonId_(eventId,beneficiary) {
  const key=normalize_(beneficiary);
  if (!key) return '';
  const p=getParticipantsForEvent_(eventId).find(x=>normalize_([x.name,x.surname].filter(Boolean).join(' '))===key);
  return p?String(p.personId||''):'';
}

function syncExpenseRemindersFromChild_(eventId,child) {
  const expenseSheet=child.getSheetByName(EVENT_SHEET.SHEETS.EXPENSES), values=expenseSheet.getDataRange().getValues();
  const check=sh_(APP.SHEETS.CHECKLIST), checkRows=check.getDataRange().getValues(), byId={};
  let maxOrder=0;
  for (let i=1;i<checkRows.length;i++) {
    if (String(checkRows[i][1])!==String(eventId)) continue;
    if (checkRows[i][0]) byId[String(checkRows[i][0])]={row:i+1,values:checkRows[i]};
    maxOrder=Math.max(maxOrder,Number(checkRows[i][2]||0));
  }
  const today=new Date(); today.setHours(0,0,0,0);
  const ids=[];
  for (let i=1;i<values.length;i++) {
    const r=values[i], expenseId=String(r[11]||'').trim();
    if (!expenseId) { ids.push(['','']); continue; }
    const rowIds=[];
    [[7,8,14,1],[9,10,15,2]].forEach(spec=>{
      const textIndex=spec[0], dueIndex=spec[1], idIndex=spec[2], slot=spec[3];
      const reminder=String(r[textIndex]||'').trim(), due=r[dueIndex] instanceof Date?r[dueIndex]:'';
      let taskId=String(r[idIndex]||'').trim(), existing=taskId&&byId[taskId]?byId[taskId]:null;
      if (!reminder) {
        if (existing&&!['FATTO','COMPLETATA'].includes(normalize_(existing.values[6]))) {
          check.getRange(existing.row,7).setValue('FATTO');
          check.getRange(existing.row,12,1,2).setValues([[new Date(),new Date()]]);
        }
        rowIds.push(taskId); return;
      }
      if (existing&&['FATTO','COMPLETATA'].includes(normalize_(existing.values[6]))) {
        const sameText=normalize_(existing.values[3])===normalize_(reminder);
        const oldDue=existing.values[5] instanceof Date?existing.values[5]:'';
        const sameDue=(!oldDue&&!due)||(oldDue&&due&&oldDue.getTime()===due.getTime());
        if (!sameText||!sameDue) { taskId=''; existing=null; }
      }
      if (!taskId) taskId='TASK-'+Utilities.getUuid();
      let status=existing&&['FATTO','COMPLETATA'].includes(normalize_(existing.values[6]))?'FATTO':'DA FARE';
      if (status!=='FATTO'&&due instanceof Date) { const d=new Date(due); d.setHours(0,0,0,0); if (d>today) status='IN ATTESA'; }
      const description=String(r[0]||'').trim(), reference=String(r[6]||'').trim();
      const note=['Spesa: '+description,reference].filter(Boolean).join(' | '), autoKey='PROMEMORIA_SPESA:'+expenseId+':'+slot;
      if (existing) {
        check.getRange(existing.row,4).setValue(reminder); check.getRange(existing.row,5).setValue('PAGAMENTI'); check.getRange(existing.row,6).setValue(due||'');
        check.getRange(existing.row,7).setValue(status); check.getRange(existing.row,9).setValue('PERSONALIZZATA'); check.getRange(existing.row,10).setValue(autoKey);
        check.getRange(existing.row,11).setValue(note); if (status!=='FATTO') check.getRange(existing.row,12).clearContent(); check.getRange(existing.row,13).setValue(new Date());
      } else {
        maxOrder+=10;
        check.appendRow([taskId,eventId,maxOrder,reminder,'PAGAMENTI',due||'',status,'','PERSONALIZZATA',autoKey,note,status==='FATTO'?new Date():'',new Date()]);
        byId[taskId]={row:check.getLastRow(),values:[]};
      }
      rowIds.push(taskId);
    });
    ids.push(rowIds);
  }
  if (ids.length) expenseSheet.getRange(2,15,ids.length,2).setValues(ids);
}

function syncGuestsFromEventSheet_(eventId,child) {
  const sheet=child.getSheetByName(EVENT_SHEET.SHEETS.GUESTS), values=sheet.getDataRange().getValues(), rows=[];
  for (let i=1;i<values.length;i++) {
    const r=values[i].slice(0,12), name=String(r[3]||'').trim(), surname=String(r[4]||'').trim();
    if (!name&&!surname) continue;
    if (!r[0]) r[0]='PAR-'+Utilities.getUuid();
    r[1]=eventId;
    if (!r[2]) {
      const known=findKnownPerson_(name,surname,'');
      if (known) { r[2]=known.id; if (!r[5]) r[5]=known.club||''; if (!r[7]) r[7]=known.email||''; }
    }
    if (!r[9]) r[9]='SCHEDA EVENTO';
    while (r.length<12) r.push('');
    rows.push(r);
  }
  replaceCentralRowsForEvent_(sh_(APP.SHEETS.PARTICIPANTS),eventId,2,rows,12);
  return rows.length;
}

function refreshGuestSummaryFormulas_(sheet,guestCount) {
  const clearRows=Math.max(sheet.getLastRow()-1,guestCount,1);
  sheet.getRange(2,13,clearRows,4).clearContent();
  if (!guestCount) return;
  const formulas=[];
  for (let i=0;i<guestCount;i++) {
    const row=i+2, person='TRIM($D'+row+'&" "&$E'+row+')';
    formulas.push([
      '=IF(AND($D'+row+'="",$E'+row+'=""),"",SUMIFS(Spese!$E:$E,Spese!$C:$C,'+person+',Spese!$B:$B,"<>RIMBORSO",Spese!$D:$D,"<>PREVENTIVO",Spese!$F:$F,"PAGATO"))',
      '=IF(AND($D'+row+'="",$E'+row+'=""),"",SUMIFS(Spese!$E:$E,Spese!$C:$C,'+person+',Spese!$B:$B,"RIMBORSO",Spese!$F:$F,"DA PAGARE"))',
      '=IF(AND($D'+row+'="",$E'+row+'=""),"",SUMIFS(Spese!$E:$E,Spese!$C:$C,'+person+',Spese!$B:$B,"RIMBORSO",Spese!$F:$F,"PAGATO"))',
      '=IF(AND($D'+row+'="",$E'+row+'=""),"",IF(N'+row+'>0,"DA PASSARE",IF(O'+row+'>0,"PASSATO","—")))'
    ]);
  }
  sheet.getRange(2,13,guestCount,4).setFormulas(formulas);
}

function updateExpenseBeneficiaryValidation_(sheet,guests) {
  const names=guests.map(r=>[r[3],r[4]].filter(Boolean).join(' ').trim()).filter(Boolean);
  if (!names.length) { sheet.getRange('C2:C1000').clearDataValidations(); return; }
  sheet.getRange('C2:C1000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(names,true).setAllowInvalid(true).build());
}

function replaceCentralRowsForEvent_(sheet,eventId,eventColumn,newRows,width) {
  const values=sheet.getDataRange().getValues(), targetRows=[];
  for (let i=1;i<values.length;i++) if (String(values[i][eventColumn-1])===String(eventId)) targetRows.push(i+1);
  const reused=Math.min(targetRows.length,newRows.length);
  for (let i=0;i<reused;i++) sheet.getRange(targetRows[i],1,1,width).setValues([newRows[i]]);
  for (let i=reused;i<targetRows.length;i++) sheet.getRange(targetRows[i],1,1,width).clearContent();
  if (newRows.length>reused) {
    const remaining=newRows.slice(reused);
    sheet.getRange(sheet.getLastRow()+1,1,remaining.length,width).setValues(remaining);
  }
}

function validateEventSheetIdentity_(child,eventId) {
  const meta=child.getSheetByName(EVENT_SHEET.SHEETS.META);
  if (!meta) throw new Error('La Scheda evento non contiene il foglio tecnico _META.');
  const stored=String(readMetaValue_(meta,'EVENT_ID')||'').trim();
  if (stored&&stored!==String(eventId)) throw new Error('La Scheda evento appartiene a un altro evento: '+stored);
}
function readMetaValue_(sheet,key) {
  const values=sheet.getRange(1,1,Math.max(sheet.getLastRow(),1),2).getValues(), target=normalize_(key);
  for (let i=0;i<values.length;i++) if (normalize_(values[i][0])===target) return values[i][1];
  return '';
}
function writeMeta_(sheet,object) {
  const values=sheet.getRange(1,1,Math.max(sheet.getLastRow(),1),2).getValues(), rowsByKey={};
  values.forEach((r,i)=>{if(r[0]) rowsByKey[normalize_(r[0])]=i+1;});
  if (!values[0]||normalize_(values[0][0])!=='CHIAVE') {
    sheet.getRange(1,1,1,2).setValues([['CHIAVE','VALORE']]).setFontWeight('bold').setBackground('#eeeeee'); rowsByKey.CHIAVE=1;
  }
  Object.keys(object).forEach(key=>{const n=normalize_(key), row=rowsByKey[n]||sheet.getLastRow()+1; sheet.getRange(row,1,1,2).setValues([[key,object[key]]]); rowsByKey[n]=row;});
}
function normalizeTaskStatusForSheet_(status) {
  const s=normalize_(status); if (s==='FATTO'||s==='COMPLETATA') return 'FATTO'; if (s==='BLOCCATA'||s==='IN ATTESA') return 'IN ATTESA'; return 'DA FARE';
}
function buildEventSheetName_(event) {
  const type=toTitleCaseV2_(event[APP.CALENDAR_HEADERS.TYPE]||'Evento'), cls=String(event[APP.CALENDAR_HEADERS.CLASS]||'').trim(), place=String(event[APP.CALENDAR_HEADERS.LOCATION]||'').trim();
  return ['Scheda evento',type,cls,place,formatDocumentDateRange_(event[APP.CALENDAR_HEADERS.START],event[APP.CALENDAR_HEADERS.END])].filter(Boolean).join(' - ');
}
function buildEventSheetLabel_(event) {
  return [event[APP.CALENDAR_HEADERS.TYPE]||'',event[APP.CALENDAR_HEADERS.CLASS]||'',event[APP.CALENDAR_HEADERS.LOCATION]||'',formatDocumentDateRange_(event[APP.CALENDAR_HEADERS.START],event[APP.CALENDAR_HEADERS.END])].filter(Boolean).join(' | ');
}
function setEventSheetLink_(row,url) {
  const sheet=sh_(APP.SHEETS.CALENDAR), map=headerMap_(sheet); let col=map[APP.CALENDAR_HEADERS.EVENT_SHEET];
  if (!col) { col=sheet.getLastColumn()+1; sheet.getRange(1,col).setValue(APP.CALENDAR_HEADERS.EVENT_SHEET); }
  sheet.getRange(row,col).setRichTextValue(SpreadsheetApp.newRichTextValue().setText('↗ APRI').setLinkUrl(url).build());
}
