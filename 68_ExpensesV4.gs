const EXPENSES_V4 = Object.freeze({
  SHEET: 'Spese',
  DASHBOARD_TITLE_ROW: 1,
  METRIC_ROW: 2,
  CATEGORY_LABEL_ROW: 4,
  CATEGORY_VALUE_ROW: 5,
  INPUT_TITLE_ROW: 7,
  INPUT_HEADER_ROW: 8,
  INPUT_ROW: 9,
  LIST_TITLE_ROW: 11,
  LIST_HEADER_ROW: 12,
  LIST_START_ROW: 13,
  LIST_END_ROW: 500,
  VISIBLE_COLS: 10,
  TECH_START_COL: 11,
  TECH_COLS: 6,
  HEADERS: Object.freeze(['IMPORTO','DESCRIZIONE','TIPOLOGIA','FORNITORE','TIPO','STATO','RIF','IMP','DOCUMENTI','NOTE']),
  TECH_HEADERS: Object.freeze(['ID SPESA','CEB','ID PERSONA','DATA PAGAMENTO','DATA INSERIMENTO','ULTIMO AGGIORNAMENTO']),
  CATEGORIES: Object.freeze(['VIAGGI','VITTO','ALLOGGIO','NOLEGGI','ISCRIZIONI','TRASPORTO','RIMBORSO','ALTRO']),
  TYPES: Object.freeze(['PREVENTIVO','AFOR','FATTURA','CARTA DI CREDITO','ALTRO PAGAMENTO']),
  STATES: Object.freeze(['DA PAGARE','PAGATO']),
  DASHBOARD_CATEGORIES: Object.freeze(['VIAGGI','VITTO','ALLOGGIO','NOLEGGI','RIMBORSI','ALTRO'])
});

function normalizeExpenseCategoryV4_(category,description) {
  const c = normalize_(category);
  const d = normalize_(description);
  if (c === 'VITTO / ALLOGGIO') {
    if (/(HOTEL|OSTELLO|ALLOGGIO|CAMERA|SOGGIORNO)/.test(d)) return 'ALLOGGIO';
    if (/(PASTI|VITTO|RISTORANTE|PRANZO|CENA)/.test(d)) return 'VITTO';
    return 'ALLOGGIO';
  }
  if (c === 'VIAGGIO') return 'VIAGGI';
  if (EXPENSES_V4.CATEGORIES.includes(c)) return c;
  return c || 'ALTRO';
}

function dashboardBucketV4_(category) {
  const c = normalize_(category);
  if (c === 'VIAGGI') return 'VIAGGI';
  if (c === 'VITTO') return 'VITTO';
  if (c === 'ALLOGGIO') return 'ALLOGGIO';
  if (c === 'NOLEGGI') return 'NOLEGGI';
  if (c === 'RIMBORSO') return 'RIMBORSI';
  return 'ALTRO';
}

function ensureExpenseV4Structure_(child,event,folderId) {
  let sheet = child.getSheetByName(EXPENSES_V4.SHEET);
  if (!sheet) sheet = child.insertSheet(EXPENSES_V4.SHEET);
  if (sheet.getMaxRows() < EXPENSES_V4.LIST_END_ROW) {
    sheet.insertRowsAfter(sheet.getMaxRows(),EXPENSES_V4.LIST_END_ROW-sheet.getMaxRows());
  }
  if (sheet.getMaxColumns() < 16) sheet.insertColumnsAfter(sheet.getMaxColumns(),16-sheet.getMaxColumns());

  ['A1:J1','A7:H7','A11:J11'].forEach(a1=>sheet.getRange(a1).breakApart());
  sheet.getRange('A1:J1').merge();
  sheet.getRange('A7:H7').merge();
  sheet.getRange('A11:J11').merge();

  sheet.getRange('A1').setValue('RIEPILOGO SPESE EVENTO');
  sheet.getRange('A2:H2').setValues([[
    'PREVENTIVO INIZIO ANNO','',
    'PREVISIONE ATTUALE','',
    'PAGATO','',
    'IMPEGNO',''
  ]]);
  sheet.getRange('A4:F4').setValues([EXPENSES_V4.DASHBOARD_CATEGORIES]);
  sheet.getRange('A7').setValue('NUOVA SPESA');
  sheet.getRange('I7').setValue('IMPORTA SPESA');
  sheet.getRange('J7').insertCheckboxes();
  sheet.getRange('J7').setValue(false).setNote('Spunta questa casella dopo aver compilato la riga 9. La spesa viene trasferita nell elenco e la riga di inserimento si libera.');
  sheet.getRange('A8:J8').setValues([EXPENSES_V4.HEADERS]);
  sheet.getRange('A11').setValue('ELENCO SPESE');
  sheet.getRange('A12:J12').setValues([EXPENSES_V4.HEADERS]);
  sheet.getRange('K12:P12').setValues([EXPENSES_V4.TECH_HEADERS]);

  sheet.setFrozenRows(5);
  sheet.showColumns(1,10);
  sheet.hideColumns(11,6);

  sheet.getRange('A1:J1').setFontWeight('bold').setFontSize(13).setBackground('#d9eaf7').setVerticalAlignment('middle');
  sheet.setRowHeight(1,30);
  sheet.getRange('A7:H7').setFontWeight('bold').setBackground('#e8f0fe');
  sheet.getRange('I7:J7').setFontWeight('bold').setBackground('#e8f0fe').setHorizontalAlignment('center');
  sheet.getRange('A11:J11').setFontWeight('bold').setBackground('#eeeeee');
  sheet.getRange('A8:J8').setFontWeight('bold').setBackground('#eeeeee').setWrap(true);
  sheet.getRange('A12:P12').setFontWeight('bold').setBackground('#eeeeee').setWrap(true);
  sheet.getRange('A2:H2').setWrap(true).setVerticalAlignment('middle');
  sheet.getRange('A4:F5').setHorizontalAlignment('center').setWrap(true);

  [105,235,125,180,135,100,125,75,210,270].forEach((w,i)=>sheet.setColumnWidth(i+1,w));
  sheet.getRange('A9:J9').setVerticalAlignment('top').setWrap(true).setBackground('#fff9e6');
  sheet.getRange('A13:J500').setVerticalAlignment('top').setWrap(true);
  sheet.getRange('A9:A9').setNumberFormat('€ #,##0.00');
  sheet.getRange('A13:A500').setNumberFormat('€ #,##0.00');
  sheet.getRange('B2:B2').setNumberFormat('€ #,##0.00');
  sheet.getRange('D2:D2').setNumberFormat('€ #,##0.00');
  sheet.getRange('F2:F2').setNumberFormat('€ #,##0.00');
  sheet.getRange('A5:F5').setNumberFormat('€ #,##0.00');
  sheet.getRange('H9:H9').setNumberFormat('@').setBackground('#f3f4f4');
  sheet.getRange('H13:H500').setNumberFormat('@').setBackground('#f3f4f4');
  sheet.getRange('N13:P500').setNumberFormat('dd/MM/yyyy');

  const categoryRule = SpreadsheetApp.newDataValidation().requireValueInList(EXPENSES_V4.CATEGORIES,true).setAllowInvalid(false).build();
  const typeRule = SpreadsheetApp.newDataValidation().requireValueInList(EXPENSES_V4.TYPES,true).setAllowInvalid(true).build();
  const stateRule = SpreadsheetApp.newDataValidation().requireValueInList(EXPENSES_V4.STATES,true).setAllowInvalid(true).build();
  sheet.getRange('C9').setDataValidation(categoryRule);
  sheet.getRange('E9').setDataValidation(typeRule);
  sheet.getRange('F9').setDataValidation(stateRule);
  sheet.getRange('C13:C500').setDataValidation(categoryRule);
  sheet.getRange('E13:E500').setDataValidation(typeRule);
  sheet.getRange('F13:F500').setDataValidation(stateRule);

  const commitment = String(event && event[APP.CALENDAR_HEADERS.COMMITMENT] || '').trim() || resolveExpenseCommitmentV3_(event,'');
  sheet.getRange('H2').setValue(commitment).setNumberFormat('@');
  sheet.getRange('H9').setValue(commitment).setNumberFormat('@');
  sheet.getRange('H13:H500').setNumberFormat('@');

  if (folderId) {
    const url = 'https://drive.google.com/drive/folders/' + folderId;
    const richInput = SpreadsheetApp.newRichTextValue().setText('DOCUMENTI').setLinkUrl(url).build();
    const richList = SpreadsheetApp.newRichTextValue().setText('DOCUMENTI').setLinkUrl(url).build();
    sheet.getRange('I8').setRichTextValue(richInput);
    sheet.getRange('I12').setRichTextValue(richList);
    sheet.getRange('I8').setNote('Clicca qui per aprire la cartella evento. Carica il file e incolla il link nella cella DOCUMENTI della spesa.');
  }

  sheet.getRange('B2').setNote('Valore iniziale dell evento. Viene salvato una volta e resta modificabile da qui.');
  sheet.getRange('D2').setNote('Stima corrente del costo finale: spese inserite + rimborsi previsti.');
  sheet.getRange('F2').setNote('Somma delle spese segnate PAGATO + rimborsi già passati.');

  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$F13="PAGATO"')
      .setBackground('#d9ead3').setRanges([sheet.getRange('A13:J500')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$F13="DA PAGARE"')
      .setBackground('#fff2cc').setRanges([sheet.getRange('A13:J500')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($E13="PREVENTIVO",$F13="")')
      .setBackground('#e8f0fe').setRanges([sheet.getRange('A13:J500')]).build()
  ]);
  return sheet;
}

function ensureInitialEventBudgetV4_(child,event) {
  const meta = child.getSheetByName(EVENT_SHEET.SHEETS.META);
  const saved = readMetaValue_(meta,'INITIAL_BUDGET');
  if (saved !== '' && saved !== null && saved !== undefined) return Number(saved||0);
  const current = Number(event && event[APP.CALENDAR_HEADERS.BUDGET] || 0);
  writeMeta_(meta,{INITIAL_BUDGET:current || 0});
  return current || 0;
}

function saveInitialBudgetFromSheetV4_(child) {
  const sheet = child.getSheetByName(EXPENSES_V4.SHEET);
  const meta = child.getSheetByName(EVENT_SHEET.SHEETS.META);
  if (!sheet || !meta) return;
  const value = Number(sheet.getRange('B2').getValue()||0);
  writeMeta_(meta,{INITIAL_BUDGET:value});
}

function readExpenseListV4_(sheet) {
  const values = sheet.getRange(EXPENSES_V4.LIST_START_ROW,1,
    EXPENSES_V4.LIST_END_ROW-EXPENSES_V4.LIST_START_ROW+1,16).getValues();
  return values.filter(r=>{
    const id = String(r[10]||'').trim();
    const hasData = Number(r[0]||0) || String(r[1]||'').trim() || String(r[2]||'').trim();
    return hasData && id.indexOf('RIMBORSO-AUTO-') !== 0;
  });
}

function readParticipantRefundForecastV4_(child) {
  const sheet = child.getSheetByName(PARTICIPANTS_V3.SHEET);
  if (!sheet) return {forecast:0,paid:0,rows:[]};
  const values = sheet.getRange(PARTICIPANTS_V3.CONVOCATI_START_ROW,1,
    PARTICIPANTS_V3.CONVOCATI_END_ROW-PARTICIPANTS_V3.CONVOCATI_START_ROW+1,16).getValues();
  const rows = [];
  let forecast = 0;
  let paid = 0;
  values.forEach((r,index)=>{
    const name = String(r[0]||'').trim();
    const surname = String(r[1]||'').trim();
    if (!name && !surname) return;
    const status = normalize_(r[5]);
    const max = Number(r[6]||0);
    const passed = Number(r[7]||0);
    const amountForecast = status === 'ASSENTE' ? 0 : (passed > 0 ? passed : max);
    forecast += amountForecast;
    if (passed > 0) paid += passed;
    rows.push({
      rowNumber:PARTICIPANTS_V3.CONVOCATI_START_ROW+index,
      name:name,
      surname:surname,
      role:String(r[4]||'ATLETA').trim(),
      status:status,
      max:max,
      passed:passed,
      paidDate:r[8] instanceof Date ? r[8] : '',
      notes:String(r[9]||'').trim(),
      participantId:String(r[10]||'').trim(),
      personId:String(r[11]||'').trim(),
      forecast:amountForecast
    });
  });
  return {forecast:forecast,paid:paid,rows:rows};
}

function calculateExpenseDashboardV4_(child) {
  const sheet = child.getSheetByName(EXPENSES_V4.SHEET);
  if (!sheet) return null;
  const list = readExpenseListV4_(sheet);
  const buckets = {VIAGGI:0,VITTO:0,ALLOGGIO:0,NOLEGGI:0,RIMBORSI:0,ALTRO:0};
  let ordinaryForecast = 0;
  let ordinaryPaid = 0;
  list.forEach(r=>{
    const amount = Number(r[0]||0);
    const category = normalizeExpenseCategoryV4_(r[2],r[1]);
    const bucket = dashboardBucketV4_(category);
    ordinaryForecast += amount;
    buckets[bucket] = Number(buckets[bucket]||0) + amount;
    if (normalize_(r[5]) === 'PAGATO') ordinaryPaid += amount;
  });
  const refunds = readParticipantRefundForecastV4_(child);
  buckets.RIMBORSI = refunds.forecast;
  return {
    forecast:ordinaryForecast + refunds.forecast,
    paid:ordinaryPaid + refunds.paid,
    buckets:buckets,
    refunds:refunds
  };
}

function refreshExpenseDashboardV4_(child,event) {
  const sheet = child.getSheetByName(EXPENSES_V4.SHEET);
  if (!sheet) return;
  const initial = ensureInitialEventBudgetV4_(child,event);
  const calc = calculateExpenseDashboardV4_(child) || {forecast:0,paid:0,buckets:{}};
  sheet.getRange('B2').setValue(initial).setNumberFormat('€ #,##0.00');
  sheet.getRange('D2').setValue(calc.forecast||0).setNumberFormat('€ #,##0.00');
  sheet.getRange('F2').setValue(calc.paid||0).setNumberFormat('€ #,##0.00');
  const values = EXPENSES_V4.DASHBOARD_CATEGORIES.map(k=>Number(calc.buckets[k]||0));
  sheet.getRange('A5:F5').setValues([values]).setNumberFormat('€ #,##0.00');
}

function refreshExpensesV4FromBackend_(eventId,event,child,folderId) {
  const sheet = ensureExpenseV4Structure_(child,event,folderId);
  const expenses = getExpensesForEvent_(eventId);
  const ordinary = expenses.filter(x=>normalize_(x.type)!=='RIMBORSO');
  const out = ordinary.map(x=>{
    const movement = visibleExpenseTypeV3_(x);
    const isQuote = movement === 'PREVENTIVO';
    const amount = isQuote ? Number(x.budget||0) : Number(x.actual||0);
    const state = isQuote ? '' : (expenseStatusIsPaidV3_(x.status) ? 'PAGATO' : 'DA PAGARE');
    const category = normalizeExpenseCategoryV4_(x.category,x.description);
    return [
      amount||'',x.description||'',category,x.beneficiary||'',movement,state,
      visibleExpenseRifV3_(x),resolveExpenseCommitmentV3_(event,category),x.attachment||'',
      stripMovementMarker_(x.notes||''),x.id||'',x.ceb||'',x.personId||'',x.paidDate||'',x.createdAt||'',x.updatedAt||''
    ];
  });

  const participantRefunds = readParticipantRefundForecastV4_(child).rows.filter(x=>x.passed>0);
  participantRefunds.forEach(p=>{
    const beneficiary = [p.name,p.surname].filter(Boolean).join(' ');
    out.push([
      p.passed,'Rimborso ' + beneficiary,'RIMBORSO',beneficiary,'RIMBORSO','PAGATO','',
      String(event[APP.CALENDAR_HEADERS.COMMITMENT]||''), '',p.notes||'',
      'RIMBORSO-AUTO-' + (p.participantId || p.rowNumber),'CEB.002',p.personId||'',p.paidDate||'','',''
    ]);
  });

  const clearRows = EXPENSES_V4.LIST_END_ROW-EXPENSES_V4.LIST_START_ROW+1;
  sheet.getRange(EXPENSES_V4.LIST_START_ROW,1,clearRows,16).clearContent();
  if (out.length) sheet.getRange(EXPENSES_V4.LIST_START_ROW,1,out.length,16).setValues(out);
  sheet.getRange('A13:A500').setNumberFormat('€ #,##0.00');
  sheet.getRange('H13:H500').setNumberFormat('@').setBackground('#f3f4f4');
  sheet.getRange('N13:P500').setNumberFormat('dd/MM/yyyy');
  sheet.getRange('A9:J9').clearContent();
  sheet.getRange('H9').setValue(String(event[APP.CALENDAR_HEADERS.COMMITMENT]||'')).setNumberFormat('@');
  sheet.getRange('J7').setValue(false);
  refreshExpenseDashboardV4_(child,event);
  return out.length;
}

function importExpenseInputV4_(child,event) {
  const sheet = ensureExpenseV4Structure_(child,event,readMetaValue_(child.getSheetByName(EVENT_SHEET.SHEETS.META),'EVENT_FOLDER_ID'));
  const input = sheet.getRange(EXPENSES_V4.INPUT_ROW,1,1,10).getValues()[0];
  const amount = Number(input[0]||0);
  const description = String(input[1]||'').trim();
  if (!amount && !description) {
    sheet.getRange('J7').setValue(false);
    throw new Error('Compila almeno IMPORTO e DESCRIZIONE prima di importare la spesa.');
  }
  if (!description) {
    sheet.getRange('J7').setValue(false);
    throw new Error('La DESCRIZIONE è obbligatoria.');
  }
  input[2] = normalizeExpenseCategoryV4_(input[2],description);
  if (!input[4]) input[4] = 'PREVENTIVO';
  if (normalize_(input[4]) !== 'PREVENTIVO' && !input[5]) input[5] = 'DA PAGARE';
  input[7] = String(event[APP.CALENDAR_HEADERS.COMMITMENT]||'') || resolveExpenseCommitmentV3_(event,input[2]);

  let row = EXPENSES_V4.LIST_START_ROW;
  const descriptions = sheet.getRange(EXPENSES_V4.LIST_START_ROW,2,
    EXPENSES_V4.LIST_END_ROW-EXPENSES_V4.LIST_START_ROW+1,1).getDisplayValues();
  while (row <= EXPENSES_V4.LIST_END_ROW && String(descriptions[row-EXPENSES_V4.LIST_START_ROW][0]||'').trim()) row++;
  if (row > EXPENSES_V4.LIST_END_ROW) throw new Error('Elenco spese pieno.');

  const now = new Date();
  sheet.getRange(row,1,1,10).setValues([input]);
  sheet.getRange(row,11,1,6).setValues([[
    'SPESA-' + Utilities.getUuid(),
    resolveExpenseCeb_(event,'',input[2],'SPESA'),
    '', '', now, now
  ]]);
  sheet.getRange(row,1).setNumberFormat('€ #,##0.00');
  sheet.getRange(row,8).setNumberFormat('@').setBackground('#f3f4f4');
  sheet.getRange(row,14,1,3).setNumberFormat('dd/MM/yyyy');

  sheet.getRange(EXPENSES_V4.INPUT_ROW,1,1,10).clearContent();
  sheet.getRange('H9').setValue(String(event[APP.CALENDAR_HEADERS.COMMITMENT]||'')).setNumberFormat('@');
  sheet.getRange('J7').setValue(false);
  refreshExpenseDashboardV4_(child,event);
  return row;
}

function syncExpensesV4ToBackend_(eventId,event,child) {
  const found = findCalendarEventById_(eventId);
  if (!found) throw new Error('Evento non trovato: ' + eventId);
  const folderId = readMetaValue_(child.getSheetByName(EVENT_SHEET.SHEETS.META),'EVENT_FOLDER_ID');
  const sheet = ensureExpenseV4Structure_(child,event,folderId);
  saveInitialBudgetFromSheetV4_(child);

  const values = sheet.getRange(EXPENSES_V4.LIST_START_ROW,1,
    EXPENSES_V4.LIST_END_ROW-EXPENSES_V4.LIST_START_ROW+1,16).getValues();
  const oldRows = sh_(APP.SHEETS.EXPENSES).getDataRange().getValues();
  const oldById = {};
  for (let i=1;i<oldRows.length;i++) {
    if (String(oldRows[i][1])===String(eventId) && oldRows[i][0]) oldById[String(oldRows[i][0])] = oldRows[i];
  }

  const now = new Date();
  const centralRows = [];
  values.forEach((r,index)=>{
    const idInput = String(r[10]||'').trim();
    if (idInput.indexOf('RIMBORSO-AUTO-') === 0) return;
    const amount = Number(r[0]||0);
    const description = String(r[1]||'').trim();
    const hasData = amount || description || String(r[2]||'').trim() || String(r[3]||'').trim() || String(r[4]||'').trim();
    if (!hasData) return;
    if (!description) throw new Error('Descrizione mancante nella riga ' + (EXPENSES_V4.LIST_START_ROW+index) + ' del foglio Spese.');

    const category = normalizeExpenseCategoryV4_(r[2],description);
    const movement = normalizeExpenseMovement_(r[4] || 'PREVENTIVO');
    const id = idInput || 'SPESA-' + Utilities.getUuid();
    const old = oldById[id] ? oldById[id].slice(0,26) : new Array(26).fill('');
    const budget = movement === 'PREVENTIVO' ? amount : 0;
    const actual = movement === 'PREVENTIVO' ? 0 : amount;
    const paymentState = normalize_(r[5]);
    const paid = paymentState === 'PAGATO' && movement !== 'PREVENTIVO';
    const status = backendPaymentStatus_(movement,paymentState);
    const paidDate = paid ? (r[13] instanceof Date ? r[13] : (old[12] instanceof Date ? old[12] : now)) : '';
    const createdAt = r[14] instanceof Date ? r[14] : (old[16] instanceof Date ? old[16] : now);
    const rifCode = String(r[6]||'').trim();
    const attachment = String(r[8]||'').trim();
    const freeNotes = String(r[9]||'').trim();
    const notes = '[MOVIMENTO=' + movement + ']' + (freeNotes ? ' ' + freeNotes : '');
    const rifNeeded = expenseNeedsRif_(budget,actual);
    let rifStatus = old[18]||'';
    if (rifCode) rifStatus = 'RICEVUTO';
    else if (!rifNeeded) rifStatus = 'NON NECESSARIO';
    else if (!rifStatus || normalize_(rifStatus)==='NON NECESSARIO') rifStatus = 'DA RICHIEDERE';
    const invoiceReceived = movement === 'FATTURA';
    const closed = ['PAGATO - FATTURA','PAGATO CON CC','PAGATO','CHIUSO'].includes(normalize_(status));

    old[0]=id; old[1]=eventId; old[2]='SPESA'; old[3]=category; old[4]=resolveExpenseCeb_(found.event,'',category,'SPESA');
    old[5]=description; old[6]=String(r[3]||'').trim(); old[7]=String(r[12]||'').trim();
    old[8]=budget; old[9]=actual; old[10]=''; old[11]=status; old[12]=paidDate; old[13]=''; old[14]=attachment; old[15]=notes;
    old[16]=createdAt; old[17]=now; old[18]=rifStatus; old[20]=rifCode; old[21]=invoiceReceived;
    old[22]=invoiceReceived ? (old[22]||now) : ''; old[25]=closed ? (old[25]||now) : '';
    centralRows.push(old);
  });

  // I rimborsi sono governati esclusivamente dal foglio Partecipanti.
  const refunds = readParticipantRefundForecastV4_(child).rows.filter(p=>p.passed>0);
  refunds.forEach(p=>{
    const participantKey = p.participantId || ('ROW-' + p.rowNumber);
    const id = 'RIMBORSO-AUTO-' + participantKey;
    const old = oldById[id] ? oldById[id].slice(0,26) : new Array(26).fill('');
    const beneficiary = [p.name,p.surname].filter(Boolean).join(' ');
    old[0]=id; old[1]=eventId; old[2]='RIMBORSO'; old[3]='RIMBORSO'; old[4]='CEB.002';
    old[5]='Rimborso ' + beneficiary; old[6]=beneficiary; old[7]=p.personId||'';
    old[8]=0; old[9]=p.passed; old[10]=''; old[11]='RIMBORSATO'; old[12]=p.paidDate||now;
    old[13]=''; old[14]=''; old[15]='[MOVIMENTO=RIMBORSO] [AUTO_RIMBORSO=' + participantKey + ']' + (p.notes ? ' ' + p.notes : '');
    old[16]=old[16]||now; old[17]=now; old[18]='NON NECESSARIO'; old[20]=''; old[21]=false; old[25]=old[25]||now;
    centralRows.push(old);
  });

  replaceCentralRowsForEvent_(sh_(APP.SHEETS.EXPENSES),eventId,2,centralRows,26);
  refreshExpenseDashboardV4_(child,event);
  return centralRows.length;
}

function handleExpenseSheetEditV4_(e,event) {
  if (!e || !e.range) return false;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== EXPENSES_V4.SHEET) return false;
  const row = e.range.getRow();
  const col = e.range.getColumn();
  const child = e.source;

  if (row === EXPENSES_V4.INPUT_TITLE_ROW && col === 10 && e.value === 'TRUE') {
    importExpenseInputV4_(child,event);
    return true;
  }
  if (row === 2 && col === 2) {
    saveInitialBudgetFromSheetV4_(child);
    refreshExpenseDashboardV4_(child,event);
    return true;
  }
  if (row >= EXPENSES_V4.LIST_START_ROW && row <= EXPENSES_V4.LIST_END_ROW && col <= 10) {
    refreshExpenseDashboardV4_(child,event);
    return true;
  }
  return true;
}
