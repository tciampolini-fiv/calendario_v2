const EXPENSES_V3 = Object.freeze({
  HEADERS: Object.freeze([
    'IMPORTO','DESCRIZIONE','TIPOLOGIA','FORNITORE','TIPO','STATO','RIF','IMP','DOCUMENTI','NOTE',
    'ID SPESA','CEB','ID PERSONA','DATA PAGAMENTO','DATA INSERIMENTO','ULTIMO AGGIORNAMENTO'
  ]),
  VISIBLE_COLS: 10,
  TECH_START_COL: 11,
  CATEGORIES: Object.freeze(['VITTO / ALLOGGIO','VIAGGI','NOLEGGI','ISCRIZIONI','TRASPORTO','RIMBORSO','ALTRO']),
  TYPES: Object.freeze(['PREVENTIVO','AFOR','FATTURA','CARTA DI CREDITO','ALTRO PAGAMENTO']),
  STATES: Object.freeze(['DA PAGARE','PAGATO'])
});

function isLegacyExpenseSheetV3_(sheet) {
  return sheet && normalize_(sheet.getRange('A1').getDisplayValue()) === 'DESCRIZIONE';
}

function ensureExpenseV3Structure_(child, folderId) {
  let sheet = child.getSheetByName(EVENT_SHEET.SHEETS.EXPENSES);
  if (!sheet) sheet = child.insertSheet(EVENT_SHEET.SHEETS.EXPENSES);
  if (sheet.getMaxColumns() < EXPENSES_V3.HEADERS.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), EXPENSES_V3.HEADERS.length - sheet.getMaxColumns());
  }

  const isV3 = normalize_(sheet.getRange('A1').getDisplayValue()) === 'IMPORTO';
  if (!isV3) {
    sheet.getRange(1,1,sheet.getMaxRows(),EXPENSES_V3.HEADERS.length).clearContent().clearDataValidations();
  }

  sheet.getRange(1,1,1,EXPENSES_V3.HEADERS.length).setValues([EXPENSES_V3.HEADERS]);
  sheet.setFrozenRows(1);
  sheet.showColumns(1,EXPENSES_V3.VISIBLE_COLS);
  if (sheet.getMaxColumns() >= EXPENSES_V3.TECH_START_COL) {
    sheet.hideColumns(EXPENSES_V3.TECH_START_COL, EXPENSES_V3.HEADERS.length - EXPENSES_V3.TECH_START_COL + 1);
  }

  sheet.getRange(1,1,1,EXPENSES_V3.HEADERS.length)
    .setFontWeight('bold').setBackground('#eeeeee').setVerticalAlignment('middle').setWrap(true);
  [110,250,150,190,155,105,135,80,230,300].forEach((w,i)=>sheet.setColumnWidth(i+1,w));
  sheet.getRange('A2:J999').setVerticalAlignment('top').setWrap(true);
  sheet.getRange('A2:A999').setNumberFormat('€ #,##0.00');
  sheet.getRange('H2:H999').setNumberFormat('@').setBackground('#f3f4f4');
  sheet.getRange('N2:P999').setNumberFormat('dd/MM/yyyy');

  sheet.getRange('C2:C999').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(EXPENSES_V3.CATEGORIES,true).setAllowInvalid(false).build()
  );
  sheet.getRange('E2:E999').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(EXPENSES_V3.TYPES,true).setAllowInvalid(true).build()
  );
  sheet.getRange('F2:F999').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(EXPENSES_V3.STATES,true).setAllowInvalid(true).build()
  );

  sheet.getRange('A1').setNote('Importo del preventivo o del singolo movimento di spesa.');
  sheet.getRange('E1').setNote('PREVENTIVO, AFOR, FATTURA, CARTA DI CREDITO oppure ALTRO PAGAMENTO.');
  sheet.getRange('F1').setNote('Per i preventivi può restare vuoto. Per i pagamenti usa DA PAGARE o PAGATO.');
  sheet.getRange('G1').setNote('Codice RIF, quando previsto.');
  sheet.getRange('H1').setNote('Impegno calcolato automaticamente in base all evento e alla configurazione degli impegni.');
  sheet.getRange('I1').setNote('Clicca sul titolo DOCUMENTI per aprire la cartella evento, carica il file dal PC e incolla qui il link Drive. Puoi inserire più link separati da un a capo.');

  if (folderId) {
    const url = 'https://drive.google.com/drive/folders/' + folderId;
    const rich = SpreadsheetApp.newRichTextValue().setText('DOCUMENTI').setLinkUrl(url).build();
    sheet.getRange('I1').setRichTextValue(rich);
  }

  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$F2="PAGATO"')
      .setBackground('#d9ead3').setRanges([sheet.getRange('A2:J999')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$F2="DA PAGARE"')
      .setBackground('#fff2cc').setRanges([sheet.getRange('A2:J999')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($E2="PREVENTIVO",$F2="")')
      .setBackground('#e8f0fe').setRanges([sheet.getRange('A2:J999')]).build()
  ]);
  return sheet;
}

function resolveExpenseCommitmentV3_(event, category) {
  const rows = sh_(APP.SHEETS.COMMITMENT_CONFIG).getDataRange().getDisplayValues();
  const type = normalize_(event[APP.CALENDAR_HEADERS.TYPE]);
  const cls = normalize_(event[APP.CALENDAR_HEADERS.CLASS]);
  const cat = normalize_(category);
  const typeAliases = [type];
  if (type.indexOf('REGATA INT.') === 0) typeAliases.push('REGATA');
  if (type === 'STAGE') typeAliases.push('ALLENAMENTO');

  let fallback = '';
  for (let i=1;i<rows.length;i++) {
    const rType = normalize_(rows[i][0]);
    const rClass = normalize_(rows[i][1]);
    const code = String(rows[i][2]||'').trim();
    const expenseType = normalize_(rows[i][4]||'');
    if (!code || !typeAliases.includes(rType)) continue;
    if (rClass && rClass !== '*' && rClass !== cls) continue;
    if (expenseType && expenseType !== '*' && expenseType !== cat) continue;
    if (expenseType === cat && cat) return code;
    if (!fallback) fallback = code;
  }
  if (fallback) return fallback;
  const eventCode = String(event[APP.CALENDAR_HEADERS.COMMITMENT]||'').trim();
  if (eventCode) return eventCode;
  return resolveCommitmentCode_(type,cls);
}

function visibleExpenseTypeV3_(expense) {
  const notes = String(expense.notes||'');
  const marker = notes.match(/\[MOVIMENTO=([^\]]+)\]/i);
  if (marker) return normalizeExpenseMovement_(marker[1]);
  return inferExpenseMovement_(expense);
}

function visibleExpenseRifV3_(expense) {
  const explicit = String(expense.rifCode||'').trim();
  if (explicit) return explicit;
  const reference = String(expense.reference||'').trim();
  return /\bRIF\b/i.test(reference) ? reference : '';
}

function expenseStatusIsPaidV3_(status) {
  const s = normalize_(status);
  return ['PAGATO - FATTURA','PAGATO - AFOR','PAGAMENTO FATTURA','PAGAMENTO AFOR','AFOR FATTO','PAGATO CON CC','INVIATO IN AMMINISTRAZIONE','CHIUSO','PAGATO','RIMBORSATO'].includes(s);
}

function refreshExpensesV3FromBackend_(eventId, event, child, folderId) {
  const sheet = ensureExpenseV3Structure_(child,folderId);
  const expenses = getExpensesForEvent_(eventId);
  const out = expenses.map(x=>{
    const movement = visibleExpenseTypeV3_(x);
    const isQuote = movement === 'PREVENTIVO';
    const amount = isQuote ? Number(x.budget||0) : Number(x.actual||0);
    const state = isQuote ? '' : (expenseStatusIsPaidV3_(x.status) ? 'PAGATO' : 'DA PAGARE');
    const category = x.category || (x.type === 'RIMBORSO' ? 'RIMBORSO' : 'ALTRO');
    return [
      amount || '', x.description||'', category, x.beneficiary||'', movement, state,
      visibleExpenseRifV3_(x), resolveExpenseCommitmentV3_(event,category), x.attachment||'',
      stripMovementMarker_(x.notes||''), x.id||'', x.ceb||'', x.personId||'', x.paidDate||'', x.createdAt||'', x.updatedAt||''
    ];
  });
  const clearRows = Math.max(sheet.getLastRow()-1,out.length,1);
  sheet.getRange(2,1,clearRows,EXPENSES_V3.HEADERS.length).clearContent();
  if (out.length) sheet.getRange(2,1,out.length,EXPENSES_V3.HEADERS.length).setValues(out);
  sheet.getRange('A2:A999').setNumberFormat('€ #,##0.00');
  sheet.getRange('H2:H999').setNumberFormat('@').setBackground('#f3f4f4');
  sheet.getRange('N2:P999').setNumberFormat('dd/MM/yyyy');
  return out.length;
}

function syncExpensesV3ToBackend_(eventId, event, child) {
  const found = findCalendarEventById_(eventId);
  if (!found) throw new Error('Evento non trovato: ' + eventId);
  const sheet = ensureExpenseV3Structure_(child, readMetaValue_(child.getSheetByName(EVENT_SHEET.SHEETS.META),'EVENT_FOLDER_ID'));
  const values = sheet.getDataRange().getValues();
  const oldRows = sh_(APP.SHEETS.EXPENSES).getDataRange().getValues();
  const oldById = {};
  for (let i=1;i<oldRows.length;i++) {
    if (String(oldRows[i][1])===String(eventId) && oldRows[i][0]) oldById[String(oldRows[i][0])] = oldRows[i];
  }

  const now = new Date();
  const centralRows = [];
  const techRows = [];
  for (let i=1;i<values.length;i++) {
    const r = values[i];
    const amount = Number(r[0]||0);
    const description = String(r[1]||'').trim();
    const hasData = amount || description || String(r[2]||'').trim() || String(r[3]||'').trim() || String(r[4]||'').trim();
    if (!hasData) continue;
    if (!description) throw new Error('Descrizione mancante nella riga ' + (i+1) + ' del foglio Spese.');

    const categoryInput = String(r[2]||'').trim() || 'ALTRO';
    let movement = normalizeExpenseMovement_(r[4] || (normalize_(categoryInput)==='RIMBORSO' ? 'RIMBORSO' : 'PREVENTIVO'));
    if (normalize_(categoryInput)==='RIMBORSO') movement = 'RIMBORSO';
    let id = String(r[10]||'').trim();
    if (!id) id = 'SPESA-' + Utilities.getUuid();
    const old = oldById[id] ? oldById[id].slice(0,26) : new Array(26).fill('');
    const recordType = movement === 'RIMBORSO' ? 'RIMBORSO' : 'SPESA';
    const category = recordType === 'RIMBORSO' ? 'RIMBORSO' : categoryInput;
    const ceb = resolveExpenseCeb_(found.event,'',category,recordType);
    const beneficiary = String(r[3]||'').trim();
    const personId = String(r[12]||'').trim() || resolveExpensePersonId_(eventId,beneficiary);
    const paymentState = normalize_(r[5]);
    const paid = paymentState === 'PAGATO' && movement !== 'PREVENTIVO';
    const budget = movement === 'PREVENTIVO' ? amount : 0;
    const actual = movement === 'PREVENTIVO' ? 0 : amount;
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
    const closed = ['PAGATO - FATTURA','PAGATO CON CC','PAGATO','RIMBORSATO','CHIUSO'].includes(normalize_(status));

    old[0]=id; old[1]=eventId; old[2]=recordType; old[3]=category; old[4]=ceb; old[5]=description; old[6]=beneficiary; old[7]=personId;
    old[8]=budget; old[9]=actual; old[10]=''; old[11]=status; old[12]=paidDate; old[13]=''; old[14]=attachment; old[15]=notes;
    old[16]=createdAt; old[17]=now; old[18]=rifStatus; old[20]=rifCode; old[21]=invoiceReceived;
    old[22]=invoiceReceived ? (old[22]||now) : ''; old[25]=closed ? (old[25]||now) : '';
    centralRows.push(old);
    techRows.push([id,ceb,personId,paidDate,createdAt,now]);
  }

  replaceCentralRowsForEvent_(sh_(APP.SHEETS.EXPENSES),eventId,2,centralRows,26);
  if (techRows.length) sheet.getRange(2,11,techRows.length,6).setValues(techRows);
  return centralRows.length;
}
