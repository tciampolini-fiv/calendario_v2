const TASKS_V4 = Object.freeze({
  TABLE_NAME: 'AttivitaEvento',
  HEADERS: Object.freeze(['DESCRIZIONE TASK','NOTE','N.','DIPENDE DA','STATO','SCADENZA']),
  VISIBLE_COLS: 6,
  ENTRY_END_ROW: 500,
  STATUS: Object.freeze(['DA FARE','IN ATTESA','FATTO'])
});

function ensureTaskV4Structure_(child) {
  let sheet = child.getSheetByName(EVENT_SHEET.SHEETS.TASKS);
  if (!sheet) sheet = child.insertSheet(EVENT_SHEET.SHEETS.TASKS, 0);
  if (sheet.getMaxRows() < TASKS_V4.ENTRY_END_ROW) {
    sheet.insertRowsAfter(sheet.getMaxRows(), TASKS_V4.ENTRY_END_ROW - sheet.getMaxRows());
  }

  sheet.getRange(1,1,1,TASKS_V4.VISIBLE_COLS).setValues([TASKS_V4.HEADERS]);
  sheet.setFrozenRows(1);
  sheet.showColumns(1,TASKS_V4.VISIBLE_COLS);
  if (sheet.getMaxColumns() > TASKS_V4.VISIBLE_COLS) {
    sheet.hideColumns(TASKS_V4.VISIBLE_COLS + 1, sheet.getMaxColumns() - TASKS_V4.VISIBLE_COLS);
  }

  sheet.getRange(1,1,1,TASKS_V4.VISIBLE_COLS)
    .setFontWeight('bold')
    .setVerticalAlignment('middle')
    .setWrap(true);
  [320,330,55,95,115,105].forEach((w,i)=>sheet.setColumnWidth(i+1,w));
  sheet.getRange('A2:F' + TASKS_V4.ENTRY_END_ROW).setVerticalAlignment('top').setWrap(true);
  sheet.getRange('C2:D' + TASKS_V4.ENTRY_END_ROW).setNumberFormat('0');
  sheet.getRange('E2:E' + TASKS_V4.ENTRY_END_ROW).setNumberFormat('@');
  sheet.getRange('F2:F' + TASKS_V4.ENTRY_END_ROW).setNumberFormat('dd/MM/yyyy');

  sheet.getRange('A1').setNote('Scegli una task frequente dal menu oppure scrivila/modificala liberamente.');
  sheet.getRange('C1').setNote('Numero stabile della task. Una nuova task riceve sempre il numero massimo esistente + 1, anche se la inserisci in mezzo alle altre.');
  sheet.getRange('D1').setNote('Indica il numero della task da cui dipende. Il menu legge direttamente i numeri presenti nella colonna N.');
  sheet.getRange('E1').setNote('Le task dipendenti restano IN ATTESA finché la task precedente non è FATTO, poi passano a DA FARE. Puoi segnare FATTO manualmente.');
  sheet.getRange('F1').setNote('Per Check conferma presenze inserisci qui la data limite indicata nella convocazione.');

  applyTaskV4Validations_(sheet);
  applyTaskV4ConditionalFormatting_(sheet);
  return sheet;
}

function applyTaskV4Validations_(sheet) {
  const activityChoices = getTaskPresetChoicesV3_();
  sheet.getRange('A2:A' + TASKS_V4.ENTRY_END_ROW).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(activityChoices,true)
      .setAllowInvalid(true)
      .build()
  );
  sheet.getRange('D2:D' + TASKS_V4.ENTRY_END_ROW).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(sheet.getRange('C2:C' + TASKS_V4.ENTRY_END_ROW),true)
      .setAllowInvalid(true)
      .build()
  );
  sheet.getRange('E2:E' + TASKS_V4.ENTRY_END_ROW).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(TASKS_V4.STATUS,true)
      .setAllowInvalid(true)
      .build()
  );
}

function applyTaskV4ConditionalFormatting_(sheet) {
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$E2="FATTO"')
      .setBackground('#d9ead3')
      .setRanges([sheet.getRange('A2:F' + TASKS_V4.ENTRY_END_ROW)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($E2="DA FARE",$F2<>"",$F2<TODAY())')
      .setBackground('#f4cccc')
      .setRanges([sheet.getRange('A2:F' + TASKS_V4.ENTRY_END_ROW)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$E2="DA FARE"')
      .setBackground('#fff2cc')
      .setRanges([sheet.getRange('A2:F' + TASKS_V4.ENTRY_END_ROW)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$E2="IN ATTESA"')
      .setBackground('#d9eaf7')
      .setRanges([sheet.getRange('A2:F' + TASKS_V4.ENTRY_END_ROW)])
      .build()
  ]);
}

function taskLiveStatusFormulaV4_(row, isPresenceCheck) {
  const unlocked = isPresenceCheck
    ? 'IF(OR(F' + row + '="";F' + row + '>TODAY());"IN ATTESA";"DA FARE")'
    : '"DA FARE"';
  return '=IF(A' + row + '="";"";' +
    'IF(D' + row + '="";"DA FARE";' +
      'IF(D' + row + '=C' + row + ';"IN ATTESA";' +
        'IFERROR(' +
          'IF(INDEX($E$2:$E$' + TASKS_V4.ENTRY_END_ROW + ';MATCH(D' + row + ';$C$2:$C$' + TASKS_V4.ENTRY_END_ROW + ';0))="FATTO";' + unlocked + ';"IN ATTESA")' +
        ';"IN ATTESA")' +
      ')' +
    ')' +
  ')';
}

function refreshTasksV4FromBackend_(eventId, event, child) {
  ensureChecklistBackendHeadersV3_();
  seedPresenceCheckTaskV3_(eventId,event,child);
  ensureTaskNumbersAndDefaultDependenciesV3_(eventId,event);

  const backendSheet = sh_(APP.SHEETS.CHECKLIST);
  const backend = backendSheet.getDataRange().getValues().slice(1)
    .filter(r=>String(r[1])===String(eventId));
  const byId = {};
  backend.forEach(r=>{ if (r[0]) byId[String(r[0])] = r; });
  backend.sort((a,b)=>Number(a[2]||999999)-Number(b[2]||999999) || Number(a[13]||999999)-Number(b[13]||999999));

  const sheet = ensureTaskV4Structure_(child);
  sheet.getRange(2,1,TASKS_V4.ENTRY_END_ROW-1,TASKS_V4.VISIBLE_COLS).clearContent();

  const out = backend.map(r=>{
    const dep = r[14] ? byId[String(r[14])] : null;
    return [
      r[3]||'',
      r[10]||'',
      r[13]||'',
      dep ? dep[13]||'' : '',
      normalize_(r[6]) === 'COMPLETATA' ? 'FATTO' : (r[6]||'DA FARE'),
      r[5]||''
    ];
  });
  if (out.length) sheet.getRange(2,1,out.length,TASKS_V4.VISIBLE_COLS).setValues(out);

  backend.forEach((r,index)=>{
    const row = index + 2;
    const depId = String(r[14]||'').trim();
    const status = normalize_(r[6]);
    if (!depId || status === 'FATTO' || status === 'COMPLETATA') return;
    const isPresenceCheck = normalize_(r[9]) === 'CHECK_CONFERME';
    sheet.getRange(row,5).setFormula(taskLiveStatusFormulaV4_(row,isPresenceCheck));
  });

  applyTaskV4Validations_(sheet);
  ensureNativeTaskTableV4_(child, sheet, Math.max(out.length + 1, 2));
  ensureTaskEditTriggerV4_(child);
  return out.length;
}

function syncTasksV4ToBackend_(eventId, event, child) {
  ensureChecklistBackendHeadersV3_();
  const sheet = ensureTaskV4Structure_(child);
  const lastRow = Math.min(Math.max(sheet.getLastRow(),2), TASKS_V4.ENTRY_END_ROW);
  const values = sheet.getRange(2,1,lastRow-1,TASKS_V4.VISIBLE_COLS).getValues();
  const formulas = sheet.getRange(2,1,lastRow-1,TASKS_V4.VISIBLE_COLS).getFormulas();

  const backendSheet = sh_(APP.SHEETS.CHECKLIST);
  const oldRows = backendSheet.getDataRange().getValues();
  const oldByNo = {};
  let maxNo = 0;
  for (let i=1;i<oldRows.length;i++) {
    if (String(oldRows[i][1]) !== String(eventId)) continue;
    const n = Number(oldRows[i][13]||0);
    if (n > 0) oldByNo[n] = oldRows[i];
    if (n > maxNo) maxNo = n;
  }

  const draft = [];
  values.forEach((r,i)=>{
    const description = String(r[0]||'').trim();
    if (!description) return;
    let no = Number(r[2]||0);
    if (!(no > 0)) {
      maxNo++;
      no = maxNo;
      sheet.getRange(i+2,3).setValue(no);
    }
    if (no > maxNo) maxNo = no;
    draft.push({sheetRow:i+2, values:r, formulas:formulas[i], no:no});
  });

  const seen = new Set();
  draft.forEach(x=>{
    if (seen.has(x.no)) throw new Error('Numero task duplicato: ' + x.no + '.');
    seen.add(x.no);
  });

  const idByNo = {};
  draft.forEach(x=>{
    const old = oldByNo[x.no];
    x.old = old ? old.slice(0,16) : new Array(16).fill('');
    x.id = old && old[0] ? String(old[0]) : 'TASK-' + Utilities.getUuid();
    idByNo[x.no] = x.id;
  });

  const now = new Date();
  const rows = [];
  draft.forEach((x,index)=>{
    const r = x.values;
    const depNo = Number(r[3]||0);
    if (depNo && !idByNo[depNo]) throw new Error('La task n. ' + x.no + ' dipende dalla task n. ' + depNo + ', che non esiste.');
    if (depNo === x.no) throw new Error('La task n. ' + x.no + ' non può dipendere da se stessa.');

    const old = x.old;
    const depId = depNo ? idByNo[depNo] : '';
    let status = normalize_(r[4]) || 'DA FARE';
    if (!TASKS_V4.STATUS.includes(status)) status = 'DA FARE';
    const due = r[5] instanceof Date ? r[5] : '';
    const completed = status === 'FATTO'
      ? (old[11] instanceof Date ? old[11] : now)
      : '';
    const source = old[8] || 'MANUALE';
    const autoKey = old[9] || inferTaskAutoKeyV3_(r[0]);
    const category = old[4] || inferTaskProcessV3_(r[0],'ALTRO',autoKey);
    const hasStatusFormula = !!String((x.formulas && x.formulas[4]) || '').trim();
    let autoBlock = '';
    if (status !== 'FATTO' && depNo && hasStatusFormula) {
      const parent = draft.find(y=>y.no===depNo);
      const parentDone = parent && normalize_(parent.values[4]) === 'FATTO';
      if (!parentDone) autoBlock = 'DIPENDENZA';
      else if (normalize_(autoKey) === 'CHECK_CONFERME' && (!(due instanceof Date) || due > now)) autoBlock = 'DATA';
    }

    rows.push([
      x.id,
      eventId,
      (index + 1) * 10,
      String(r[0]||'').trim(),
      category,
      due,
      status,
      old[7]||'',
      source,
      autoKey,
      String(r[1]||'').trim(),
      completed,
      now,
      x.no,
      depId,
      autoBlock
    ]);
  });

  replaceCentralRowsForEvent_(backendSheet,eventId,2,rows,16);
  ensureNativeTaskTableV4_(child, sheet, Math.max(rows.length + 1, 2));
  return rows.length;
}

function ensureNativeTaskTableV4_(child, sheet, endRow) {
  endRow = Math.max(2, Math.min(Number(endRow||2), TASKS_V4.ENTRY_END_ROW));
  try {
    const filter = sheet.getFilter();
    if (filter) filter.remove();
  } catch (e) {}

  try {
    const ss = Sheets.Spreadsheets.get(child.getId(), {
      fields: 'sheets(properties(sheetId,title),tables(tableId,name,range))'
    });
    const sheetInfo = (ss.sheets || []).find(s=>s.properties && Number(s.properties.sheetId)===Number(sheet.getSheetId()));
    const tables = sheetInfo && sheetInfo.tables ? sheetInfo.tables : [];
    const table = tables.find(t=>String(t.name||'')===TASKS_V4.TABLE_NAME) || tables[0];
    const range = {
      sheetId: sheet.getSheetId(),
      startRowIndex: 0,
      endRowIndex: endRow,
      startColumnIndex: 0,
      endColumnIndex: TASKS_V4.VISIBLE_COLS
    };
    if (table) {
      Sheets.Spreadsheets.batchUpdate({requests:[{
        updateTable:{table:{tableId:table.tableId,range:range},fields:'range'}
      }]}, child.getId());
    } else {
      Sheets.Spreadsheets.batchUpdate({requests:[{
        addTable:{table:{name:TASKS_V4.TABLE_NAME,range:range}}
      }]}, child.getId());
    }
  } catch (e) {
    console.log('Tabella nativa Attività non creata/aggiornata: ' + e.message);
  }
}

function ensureTaskEditTriggerV4_(child) {
  const sourceId = child.getId();
  const exists = ScriptApp.getProjectTriggers().some(t=>{
    if (t.getHandlerFunction() !== 'handleEventTaskEditV4') return false;
    try { return t.getTriggerSourceId() === sourceId; } catch (e) { return false; }
  });
  if (!exists) ScriptApp.newTrigger('handleEventTaskEditV4').forSpreadsheet(sourceId).onEdit().create();
}

function handleEventTaskEditV4(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== EVENT_SHEET.SHEETS.TASKS) return;
  const row = e.range.getRow();
  if (row < 2 || row > TASKS_V4.ENTRY_END_ROW) return;
  const firstCol = e.range.getColumn();
  const lastCol = firstCol + e.range.getNumColumns() - 1;
  if (firstCol > TASKS_V4.VISIBLE_COLS) return;

  applyTaskV4ValidationToRow_(sheet,row);
  const description = String(sheet.getRange(row,1).getDisplayValue()||'').trim();
  if (!description) return;

  const noCell = sheet.getRange(row,3);
  let no = Number(noCell.getValue()||0);
  if (!(no > 0)) {
    const last = Math.max(sheet.getLastRow(),2);
    const nums = sheet.getRange(2,3,last-1,1).getValues()
      .flat()
      .map(v=>Number(v||0))
      .filter(v=>v>0);
    no = nums.length ? Math.max.apply(null,nums) + 1 : 1;
    noCell.setValue(no).setNumberFormat('0');
  }

  // Se l utente modifica direttamente STATO, la sua scelta resta manuale.
  if (firstCol <= 5 && lastCol >= 5) return;

  const depNo = Number(sheet.getRange(row,4).getValue()||0);
  const statusCell = sheet.getRange(row,5);
  if (depNo) {
    const isPresenceCheck = normalize_(description) === 'CHECK CONFERMA PRESENZE';
    statusCell.setFormula(taskLiveStatusFormulaV4_(row,isPresenceCheck));
  } else if (!String(statusCell.getDisplayValue()||'').trim()) {
    statusCell.setValue('DA FARE');
  }
}

function applyTaskV4ValidationToRow_(sheet,row) {
  const activityChoices = getTaskPresetChoicesV3_();
  sheet.getRange(row,1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(activityChoices,true).setAllowInvalid(true).build()
  );
  sheet.getRange(row,4).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInRange(sheet.getRange('C2:C' + TASKS_V4.ENTRY_END_ROW),true).setAllowInvalid(true).build()
  );
  sheet.getRange(row,5).setNumberFormat('@').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(TASKS_V4.STATUS,true).setAllowInvalid(true).build()
  );
  sheet.getRange(row,6).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(row,1,1,TASKS_V4.VISIBLE_COLS).setVerticalAlignment('top').setWrap(true);
}
