const TASKS_V4 = Object.freeze({
  TABLE_NAME: 'AttivitaEvento',
  HEADERS: Object.freeze(['DESCRIZIONE TASK','NOTE','N.','DIPENDE DA','STATO','SCADENZA']),
  VISIBLE_COLS: 6,
  ENTRY_END_ROW: 500,
  STATUS: Object.freeze(['DA FARE','IN ATTESA','FATTO'])
});

function getNativeTaskTableStateV4_(child, sheet) {
  try {
    const ss = Sheets.Spreadsheets.get(child.getId(), {
      fields: 'sheets(properties(sheetId,title),tables(tableId,name,range,columnProperties))'
    });
    const sheetInfo = (ss.sheets || []).find(s =>
      s.properties && Number(s.properties.sheetId) === Number(sheet.getSheetId())
    );
    const tables = sheetInfo && sheetInfo.tables ? sheetInfo.tables : [];
    const table = tables.find(t => String(t.name || '') === TASKS_V4.TABLE_NAME) || tables[0] || null;
    return {ok:true, table:table};
  } catch (e) {
    console.log('Lettura tabella nativa Attivita non riuscita: ' + e.message);
    return {ok:false, table:null};
  }
}

function getNativeTaskTableV4_(child, sheet) {
  return getNativeTaskTableStateV4_(child, sheet).table;
}

function hasNativeTaskTableV4_(sheet) {
  const state = getNativeTaskTableStateV4_(sheet.getParent(), sheet);
  // Se non riusciamo a verificare, ci comportiamo in modo conservativo:
  // evitiamo modifiche Range che potrebbero essere vietate su colonne tipizzate.
  return !state.ok || !!state.table;
}

function ensureTaskV4Structure_(child) {
  let sheet = child.getSheetByName(EVENT_SHEET.SHEETS.TASKS);
  if (!sheet) sheet = child.insertSheet(EVENT_SHEET.SHEETS.TASKS, 0);
  if (sheet.getMaxRows() < TASKS_V4.ENTRY_END_ROW) {
    sheet.insertRowsAfter(sheet.getMaxRows(), TASKS_V4.ENTRY_END_ROW - sheet.getMaxRows());
  }

  const state = getNativeTaskTableStateV4_(child, sheet);
  const nativeTable = state.table;

  sheet.setFrozenRows(1);
  sheet.showColumns(1, TASKS_V4.VISIBLE_COLS);
  if (sheet.getMaxColumns() > TASKS_V4.VISIBLE_COLS) {
    sheet.hideColumns(TASKS_V4.VISIBLE_COLS + 1, sheet.getMaxColumns() - TASKS_V4.VISIBLE_COLS);
  }
  [320,330,55,95,115,105].forEach((w,i) => sheet.setColumnWidth(i + 1, w));

  // Le operazioni cella-per-cella vengono eseguite solo prima della creazione
  // della tabella nativa. Dopo, tipo e dropdown sono gestiti dalla tabella stessa.
  if (state.ok && !nativeTable) {
    sheet.getRange(1,1,1,TASKS_V4.VISIBLE_COLS).setValues([TASKS_V4.HEADERS]);
    sheet.getRange(1,1,1,TASKS_V4.VISIBLE_COLS)
      .setFontWeight('bold')
      .setVerticalAlignment('middle')
      .setWrap(true);
    sheet.getRange('A2:F' + TASKS_V4.ENTRY_END_ROW).setVerticalAlignment('top').setWrap(true);

    sheet.getRange('A1').setNote('Scegli una task frequente dal menu oppure scrivila/modificala liberamente.');
    sheet.getRange('C1').setNote('Numero stabile della task. Una nuova task riceve sempre il numero massimo esistente + 1, anche se la inserisci in mezzo alle altre.');
    sheet.getRange('D1').setNote('Indica il numero della task da cui dipende. Il menu propone i numeri presenti nella colonna N.');
    sheet.getRange('E1').setNote('Le task dipendenti restano IN ATTESA finche la task precedente non e FATTO, poi passano a DA FARE. Puoi segnare FATTO manualmente.');
    sheet.getRange('F1').setNote('Per Check conferma presenze inserisci qui la data limite indicata nella convocazione.');

    applyTaskV4ConditionalFormatting_(sheet);
  }
  return sheet;
}

function uniqueTaskChoicesV4_(values) {
  const seen = new Set();
  const out = [];
  values.forEach(value => {
    const text = String(value === null || value === undefined ? '' : value).trim();
    if (!text) return;
    const key = normalize_(text);
    if (seen.has(key)) return;
    seen.add(key);
    out.push(text);
  });
  return out;
}

function taskDropdownRuleV4_(values) {
  const choices = uniqueTaskChoicesV4_(values);
  if (!choices.length) return null;
  return {
    condition: {
      type: 'ONE_OF_LIST',
      values: choices.map(value => ({userEnteredValue:String(value)}))
    }
  };
}

function taskTableColumnPropertiesV4_(sheet) {
  const existingDescriptions = sheet.getRange(2,1,TASKS_V4.ENTRY_END_ROW-1,1)
    .getDisplayValues().flat();
  const activityChoices = uniqueTaskChoicesV4_(
    getTaskPresetChoicesV3_().concat(existingDescriptions)
  );

  const dependencyChoices = uniqueTaskChoicesV4_(
    sheet.getRange(2,3,TASKS_V4.ENTRY_END_ROW-1,1)
      .getValues().flat()
      .map(value => {
        const n = Number(value || 0);
        return n > 0 ? String(n) : '';
      })
  );

  const activityRule = taskDropdownRuleV4_(activityChoices);
  const statusRule = taskDropdownRuleV4_(TASKS_V4.STATUS);
  const dependencyRule = taskDropdownRuleV4_(dependencyChoices);

  return [
    {
      columnIndex:0,
      columnName:TASKS_V4.HEADERS[0],
      columnType:'DROPDOWN',
      dataValidationRule:activityRule
    },
    {
      columnIndex:1,
      columnName:TASKS_V4.HEADERS[1],
      columnType:'TEXT'
    },
    {
      columnIndex:2,
      columnName:TASKS_V4.HEADERS[2],
      columnType:'DOUBLE'
    },
    dependencyRule ? {
      columnIndex:3,
      columnName:TASKS_V4.HEADERS[3],
      columnType:'DROPDOWN',
      dataValidationRule:dependencyRule
    } : {
      columnIndex:3,
      columnName:TASKS_V4.HEADERS[3],
      columnType:'DOUBLE'
    },
    {
      columnIndex:4,
      columnName:TASKS_V4.HEADERS[4],
      columnType:'DROPDOWN',
      dataValidationRule:statusRule
    },
    {
      columnIndex:5,
      columnName:TASKS_V4.HEADERS[5],
      columnType:'DATE'
    }
  ];
}

function clearLegacyTaskValidationsOutsideTableV4_(sheet, table) {
  // Le vecchie versioni applicavano convalide Range fino alla riga 500.
  // Prima di ampliare una tabella esistente rimuoviamo solo quelle che sono
  // ancora FUORI dalla tabella, evitando qualsiasi operazione su celle tipizzate.
  const currentEndRow = table && table.range && table.range.endRowIndex
    ? Number(table.range.endRowIndex)
    : 1;
  const firstOutsideRow = Math.max(2, currentEndRow + 1);
  if (firstOutsideRow > TASKS_V4.ENTRY_END_ROW) return;
  sheet.getRange(
    firstOutsideRow,
    1,
    TASKS_V4.ENTRY_END_ROW - firstOutsideRow + 1,
    TASKS_V4.VISIBLE_COLS
  ).clearDataValidations();
}

function applyTaskV4Validations_(sheet) {
  // Compatibilita con chiamate legacy: i dropdown delle Attivita sono ora
  // proprieta delle colonne della tabella nativa, non convalide Range.
  if (hasNativeTaskTableV4_(sheet)) return false;
  return true;
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
    .filter(r => String(r[1]) === String(eventId));
  const byId = {};
  backend.forEach(r => { if (r[0]) byId[String(r[0])] = r; });
  backend.sort((a,b) => Number(a[2]||999999)-Number(b[2]||999999) || Number(a[13]||999999)-Number(b[13]||999999));

  const out = backend.map(r => {
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

  const sheet = ensureTaskV4Structure_(child);
  sheet.getRange(2,1,TASKS_V4.ENTRY_END_ROW-1,TASKS_V4.VISIBLE_COLS).clearContent();
  if (out.length) sheet.getRange(2,1,out.length,TASKS_V4.VISIBLE_COLS).setValues(out);

  backend.forEach((r,index) => {
    const row = index + 2;
    const depId = String(r[14]||'').trim();
    const status = normalize_(r[6]);
    if (!depId || status === 'FATTO' || status === 'COMPLETATA') return;
    const isPresenceCheck = normalize_(r[9]) === 'CHECK_CONFERME';
    sheet.getRange(row,5).setFormula(taskLiveStatusFormulaV4_(row,isPresenceCheck));
  });

  ensureNativeTaskTableV4_(child, sheet);
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
  for (let i=1; i<oldRows.length; i++) {
    if (String(oldRows[i][1]) !== String(eventId)) continue;
    const n = Number(oldRows[i][13]||0);
    if (n > 0) oldByNo[n] = oldRows[i];
    if (n > maxNo) maxNo = n;
  }

  const draft = [];
  values.forEach((r,i) => {
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
  draft.forEach(x => {
    if (seen.has(x.no)) throw new Error('Numero task duplicato: ' + x.no + '.');
    seen.add(x.no);
  });

  const idByNo = {};
  draft.forEach(x => {
    const old = oldByNo[x.no];
    x.old = old ? old.slice(0,16) : new Array(16).fill('');
    x.id = old && old[0] ? String(old[0]) : 'TASK-' + Utilities.getUuid();
    idByNo[x.no] = x.id;
  });

  const now = new Date();
  const rows = [];
  draft.forEach((x,index) => {
    const r = x.values;
    const depNo = Number(r[3]||0);
    if (depNo && !idByNo[depNo]) throw new Error('La task n. ' + x.no + ' dipende dalla task n. ' + depNo + ', che non esiste.');
    if (depNo === x.no) throw new Error('La task n. ' + x.no + ' non puo dipendere da se stessa.');

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
      const parent = draft.find(y => y.no===depNo);
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
  ensureNativeTaskTableV4_(child, sheet);
  return rows.length;
}

function ensureNativeTaskTableV4_(child, sheet) {
  try {
    const filter = sheet.getFilter();
    if (filter) filter.remove();
  } catch (e) {}

  const state = getNativeTaskTableStateV4_(child, sheet);
  if (!state.ok) {
    console.log('Tabella nativa Attivita non aggiornata: stato non verificabile.');
    return false;
  }

  try {
    clearLegacyTaskValidationsOutsideTableV4_(sheet, state.table);

    const range = {
      sheetId: sheet.getSheetId(),
      startRowIndex: 0,
      endRowIndex: TASKS_V4.ENTRY_END_ROW,
      startColumnIndex: 0,
      endColumnIndex: TASKS_V4.VISIBLE_COLS
    };
    const columnProperties = taskTableColumnPropertiesV4_(sheet);

    if (state.table) {
      Sheets.Spreadsheets.batchUpdate({requests:[{
        updateTable:{
          table:{
            tableId:state.table.tableId,
            range:range,
            columnProperties:columnProperties
          },
          fields:'range,columnProperties'
        }
      }]}, child.getId());
    } else {
      Sheets.Spreadsheets.batchUpdate({requests:[{
        addTable:{table:{
          name:TASKS_V4.TABLE_NAME,
          range:range,
          columnProperties:columnProperties
        }}
      }]}, child.getId());
    }
    return true;
  } catch (e) {
    console.log('Tabella nativa Attivita non creata/aggiornata: ' + e.message);
    return false;
  }
}

function ensureTaskEditTriggerV4_(child) {
  const sourceId = child.getId();
  const exists = ScriptApp.getProjectTriggers().some(t => {
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

  const description = String(sheet.getRange(row,1).getDisplayValue()||'').trim();
  if (!description) return;

  const noCell = sheet.getRange(row,3);
  let no = Number(noCell.getValue()||0);
  if (!(no > 0)) {
    const last = Math.max(sheet.getLastRow(),2);
    const nums = sheet.getRange(2,3,last-1,1).getValues()
      .flat()
      .map(v => Number(v||0))
      .filter(v => v>0);
    no = nums.length ? Math.max.apply(null,nums) + 1 : 1;
    noCell.setValue(no);
    // Aggiorna subito il dropdown DIPENDE DA con il nuovo numero.
    ensureNativeTaskTableV4_(sheet.getParent(), sheet);
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
  // Compatibilita legacy: nelle tabelle native la validazione e definita
  // a livello di colonna e non va piu applicata alla singola cella.
  return false;
}
