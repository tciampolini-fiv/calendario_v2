const TASKS_V3 = Object.freeze({
  HEADERS: Object.freeze([
    'N.','PROCESSO','ATTIVITA','DIPENDE DA','STATO','SCADENZA','NOTE',
    'ID TASK','ORIGINE','AUTO KEY','ID DIPENDENZA','DATA COMPLETAMENTO','ULTIMO AGGIORNAMENTO','ORDINE','BLOCCO AUTO'
  ]),
  VISIBLE_COLS: 7,
  TECH_START_COL: 8,
  ENTRY_END_ROW: 120,
  STATUS: Object.freeze(['DA FARE','IN ATTESA','FATTO']),
  PROCESSES: Object.freeze(['HOTEL','GOMMONE','PASTI','VIAGGIO TECNICO','CONVOCAZIONI','CIRCOLO','ISCRIZIONI','TRASPORTO','ALTRO']),
  COMMON_TASKS: Object.freeze([
    'Contattare hotel','Conferma hotel','Pagamento AFOR','Inviare contabile','Richiedere fattura',
    'Gommone','Conferma Gommone','Pasti','Conferma Pasti','Viaggio Tecnico',
    'Convocazione atleti','Convocazione Tecnico','Check conferma presenze',
    'Ospitalità Circolo','Conferma Ospitalità','Ringraziamento Circolo'
  ])
});

function ensureChecklistBackendHeadersV3_() {
  const sheet = sh_(APP.SHEETS.CHECKLIST);
  sheet.getRange(1,14,1,3).setValues([['N TASK','ID DIPENDENZA','BLOCCO AUTO']]);
}

function isLegacyTaskSheetV3_(sheet) {
  return sheet && normalize_(sheet.getRange('A1').getDisplayValue()) === 'ID TASK';
}

function ensureTaskV3Structure_(child) {
  let sheet = child.getSheetByName(EVENT_SHEET.SHEETS.TASKS);
  if (!sheet) sheet = child.insertSheet(EVENT_SHEET.SHEETS.TASKS, 0);
  if (sheet.getMaxColumns() < TASKS_V3.HEADERS.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), TASKS_V3.HEADERS.length - sheet.getMaxColumns());
  }
  if (sheet.getMaxRows() < TASKS_V3.ENTRY_END_ROW) {
    sheet.insertRowsAfter(sheet.getMaxRows(), TASKS_V3.ENTRY_END_ROW - sheet.getMaxRows());
  }

  const isV3 = normalize_(sheet.getRange('A1').getDisplayValue()) === 'N.';
  if (!isV3) {
    sheet.getRange(1,1,sheet.getMaxRows(),TASKS_V3.HEADERS.length).clearContent().clearDataValidations();
  }

  sheet.getRange(1,1,1,TASKS_V3.HEADERS.length).setValues([TASKS_V3.HEADERS]);
  sheet.setFrozenRows(1);
  sheet.showColumns(1,TASKS_V3.VISIBLE_COLS);
  if (sheet.getMaxColumns() >= TASKS_V3.TECH_START_COL) {
    sheet.hideColumns(TASKS_V3.TECH_START_COL, TASKS_V3.HEADERS.length - TASKS_V3.TECH_START_COL + 1);
  }

  sheet.getRange(1,1,1,TASKS_V3.HEADERS.length)
    .setFontWeight('bold').setBackground('#eeeeee').setVerticalAlignment('middle').setWrap(true);
  [55,145,285,95,115,105,330].forEach((w,i)=>sheet.setColumnWidth(i+1,w));
  sheet.getRange('A2:G' + TASKS_V3.ENTRY_END_ROW).setVerticalAlignment('top').setWrap(true);
  sheet.getRange('A2:A' + TASKS_V3.ENTRY_END_ROW).setNumberFormat('0');
  sheet.getRange('D2:D' + TASKS_V3.ENTRY_END_ROW).setNumberFormat('0');
  sheet.getRange('F2:F' + TASKS_V3.ENTRY_END_ROW).setNumberFormat('dd/MM/yyyy');
  sheet.getRange('L2:M' + TASKS_V3.ENTRY_END_ROW).setNumberFormat('dd/MM/yyyy');

  sheet.getRange('A1').setNote('Numero stabile della task. Le righe vuote vengono numerate automaticamente appena inserisci l attività.');
  sheet.getRange('C1').setNote('Puoi scegliere una voce frequente dal menu oppure scrivere liberamente una nuova task. Per aggiungere una task usa semplicemente la prima riga vuota.');
  sheet.getRange('D1').setNote('Indica il numero della task da cui questa dipende. Il menu legge direttamente la colonna N. e quindi si aggiorna automaticamente.');
  sheet.getRange('E1').setNote('Per le task dipendenti lo stato è automatico: IN ATTESA finché la task precedente non è FATTO, poi DA FARE. Puoi sempre sovrascriverlo manualmente.');
  sheet.getRange('F1').setNote('Scadenza della task. Per Check conferma presenze inserisci qui la data limite indicata nella convocazione.');

  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$E2="FATTO"')
      .setBackground('#d9ead3').setRanges([sheet.getRange('A2:G' + TASKS_V3.ENTRY_END_ROW)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($E2="DA FARE",$F2<>"",$F2<TODAY())')
      .setBackground('#f4cccc').setRanges([sheet.getRange('A2:G' + TASKS_V3.ENTRY_END_ROW)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$E2="DA FARE"')
      .setBackground('#fff2cc').setRanges([sheet.getRange('A2:G' + TASKS_V3.ENTRY_END_ROW)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$E2="IN ATTESA"')
      .setBackground('#d9eaf7').setRanges([sheet.getRange('A2:G' + TASKS_V3.ENTRY_END_ROW)]).build()
  ]);
  return sheet;
}

function applyTaskV3Validations_(sheet) {
  const activityChoices = getTaskPresetChoicesV3_();
  sheet.getRange('B2:B' + TASKS_V3.ENTRY_END_ROW).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(TASKS_V3.PROCESSES,true).setAllowInvalid(true).build()
  );
  sheet.getRange('C2:C' + TASKS_V3.ENTRY_END_ROW).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(activityChoices,true).setAllowInvalid(true).build()
  );
  sheet.getRange('D2:D' + TASKS_V3.ENTRY_END_ROW).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(sheet.getRange('A2:A' + TASKS_V3.ENTRY_END_ROW),true)
      .setAllowInvalid(true)
      .build()
  );
  sheet.getRange('E2:E' + TASKS_V3.ENTRY_END_ROW).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(TASKS_V3.STATUS,true).setAllowInvalid(false).build()
  );
}

function taskLiveStatusFormulaV3_(row) {
  return '=IF(C' + row + '="";"";' +
    'IF(D' + row + '="";"DA FARE";' +
      'IF(D' + row + '=A' + row + ';"IN ATTESA";' +
        'IFERROR(' +
          'IF(INDEX($E$2:$E$' + TASKS_V3.ENTRY_END_ROW + ';MATCH(D' + row + ';$A$2:$A$' + TASKS_V3.ENTRY_END_ROW + ';0))="FATTO";' +
            'IF(AND(J' + row + '="CHECK_CONFERME";OR(F' + row + '="";F' + row + '>TODAY()));"IN ATTESA";"DA FARE");' +
            '"IN ATTESA"' +
          ');' +
          '"IN ATTESA"' +
        ')' +
      ')' +
    ')' +
  ')';
}

function taskAutoNumberFormulaV3_(row) {
  return '=IF(C' + row + '="";"";MAX($A$1:A' + (row - 1) + ')+1)';
}

function prepareTaskEntryRowsV3_(sheet, writtenRows) {
  applyTaskV3Validations_(sheet);
  const firstBlank = Math.max(2, writtenRows + 2);
  if (firstBlank > TASKS_V3.ENTRY_END_ROW) return;

  const numberFormulas = [];
  const statusFormulas = [];
  for (let row=firstBlank; row<=TASKS_V3.ENTRY_END_ROW; row++) {
    numberFormulas.push([taskAutoNumberFormulaV3_(row)]);
    statusFormulas.push([taskLiveStatusFormulaV3_(row)]);
  }
  sheet.getRange(firstBlank,1,numberFormulas.length,1).setFormulas(numberFormulas);
  sheet.getRange(firstBlank,5,statusFormulas.length,1).setFormulas(statusFormulas);
}

function applyLiveDependencyFormulasV3_(sheet, out) {
  out.forEach((r, index)=>{
    const row = index + 2;
    const dependencyNo = Number(r[3]||0);
    const status = normalize_(r[4]);
    const autoBlock = normalize_(r[14]||'');
    if (!dependencyNo || status === 'FATTO' || autoBlock === 'MANUALE') return;
    sheet.getRange(row,5).setFormula(taskLiveStatusFormulaV3_(row));
  });
}

function getTaskPresetChoicesV3_() {
  const out = new Set(TASKS_V3.COMMON_TASKS);
  const rows = sh_(APP.SHEETS.CHECKLIST_CONFIG).getDataRange().getDisplayValues();
  for (let i=1;i<rows.length;i++) {
    const value = String(rows[i][3]||'').trim();
    if (value) out.add(value);
  }
  return Array.from(out).sort((a,b)=>a.localeCompare(b,'it'));
}

function inferTaskAutoKeyV3_(task) {
  const t = normalize_(task);
  const map = {
    'CONVOCAZIONE ATLETI':'CONV_ATLETI',
    'CONVOCAZIONE TECNICO':'CONV_TECNICO',
    'CHECK CONFERMA PRESENZE':'CHECK_CONFERME',
    'RINGRAZIAMENTO CIRCOLO':'RINGRAZIAMENTO_CIRCOLO'
  };
  return map[t] || '';
}

function inferTaskProcessV3_(task, category, autoKey) {
  const key = normalize_(autoKey || task);
  if (key.indexOf('GOMMONE') >= 0) return 'GOMMONE';
  if (key.indexOf('SOGGIORNO') >= 0 || key.indexOf('ALLOGGIO') >= 0 || key.indexOf('HOTEL') >= 0) return 'HOTEL';
  if (key.indexOf('PASTI') >= 0) return 'PASTI';
  if (key.indexOf('VIAGGIO_TECNICO') >= 0 || key.indexOf('VIAGGIO TECNICO') >= 0) return 'VIAGGIO TECNICO';
  if (key.indexOf('CONV_') === 0 || key.indexOf('CONVOCAZ') >= 0 || key === 'CHECK_CONFERME') return 'CONVOCAZIONI';
  if (key.indexOf('CIRCOLO') >= 0 || key.indexOf('OSPITALITA') >= 0 || key.indexOf('RINGRAZIAMENTO') >= 0) return 'CIRCOLO';
  return String(category||'').trim() || 'ALTRO';
}

function seedPresenceCheckTaskV3_(eventId, event, child) {
  const meta = child.getSheetByName(EVENT_SHEET.SHEETS.META);
  const rows = sh_(APP.SHEETS.CHECKLIST).getDataRange().getValues();
  const eventRows = rows.slice(1).filter(r=>String(r[1])===String(eventId));
  const existing = eventRows.find(r=>normalize_(r[9]) === 'CHECK_CONFERME');
  if (existing) {
    if (meta) writeMeta_(meta,{CHECK_CONFERME_SEEDED:'YES'});
    return;
  }
  if (meta && normalize_(readMetaValue_(meta,'CHECK_CONFERME_SEEDED')) === 'YES') return;
  const convocation = eventRows.find(r=>normalize_(r[9]) === 'CONV_ATLETI');
  if (!convocation) return;

  const now = new Date();
  const row = [
    'TASK-' + Utilities.getUuid(), eventId, 90, 'Check conferma presenze', 'CONVOCAZIONI', '',
    'IN ATTESA', '', 'AUTO', 'CHECK_CONFERME',
    'Inserire come SCADENZA la data limite di risposta indicata nella convocazione.', '', now,
    '', String(convocation[0]||''), 'DATA'
  ];
  sh_(APP.SHEETS.CHECKLIST).appendRow(row);
  if (meta) writeMeta_(meta,{CHECK_CONFERME_SEEDED:'YES'});
}

function ensureTaskNumbersAndDefaultDependenciesV3_(eventId, event) {
  ensureChecklistBackendHeadersV3_();
  const sheet = sh_(APP.SHEETS.CHECKLIST);
  const rows = sheet.getDataRange().getValues();
  const items = [];
  for (let i=1;i<rows.length;i++) {
    if (String(rows[i][1]) !== String(eventId)) continue;
    items.push({row:i+1, values:rows[i]});
  }
  items.sort((a,b)=>Number(a.values[2]||0)-Number(b.values[2]||0) || a.row-b.row);

  const used = new Set();
  items.forEach(x=>{
    const n = Number(x.values[13]||0);
    if (n > 0) used.add(n);
  });
  let next = 1;
  items.forEach(x=>{
    let n = Number(x.values[13]||0);
    if (!(n > 0)) {
      while (used.has(next)) next++;
      n = next;
      used.add(n);
      sheet.getRange(x.row,14).setValue(n);
      x.values[13] = n;
    }
  });

  const byKey = {};
  items.forEach(x=>{
    const key = normalize_(x.values[9] || x.values[3]);
    if (key) byKey[key] = x;
  });
  const rules = getChecklistRulesForEvent_(event);
  items.forEach(x=>{
    if (String(x.values[14]||'').trim()) return;
    const key = normalize_(x.values[9] || x.values[3]);
    let dependsOn = '';
    if (key === 'CHECK_CONFERME') dependsOn = 'CONV_ATLETI';
    else if (rules[key] && rules[key].dependsOn) dependsOn = rules[key].dependsOn;
    if (!dependsOn || !byKey[dependsOn]) return;
    sheet.getRange(x.row,15).setValue(String(byKey[dependsOn].values[0]||''));
    x.values[14] = String(byKey[dependsOn].values[0]||'');
  });
}

function recomputeTaskDependenciesV3_(eventId, event) {
  const sheet = sh_(APP.SHEETS.CHECKLIST);
  const rows = sheet.getDataRange().getValues();
  const items = [];
  const byId = {};
  for (let i=1;i<rows.length;i++) {
    if (String(rows[i][1]) !== String(eventId)) continue;
    const item = {row:i+1, values:rows[i]};
    items.push(item);
    if (rows[i][0]) byId[String(rows[i][0])] = item;
  }
  const today = new Date();
  today.setHours(0,0,0,0);
  const ended = eventHasEnded_(event,today);

  items.forEach(item=>{
    let current = normalize_(item.values[6]);
    if (current === 'COMPLETATA') current = 'FATTO';
    let block = normalize_(item.values[15]||'');
    if (current === 'FATTO') {
      if (block) sheet.getRange(item.row,16).clearContent();
      return;
    }

    const depId = String(item.values[14]||'').trim();
    const autoKey = normalize_(item.values[9]);
    const parent = depId ? byId[depId] : null;
    const parentStatus = parent ? normalize_(parent.values[6]) : '';
    const parentDone = !depId || parentStatus === 'FATTO' || parentStatus === 'COMPLETATA';
    let desired = current || 'DA FARE';
    let desiredBlock = block;

    if (!parentDone) {
      desired = 'IN ATTESA';
      desiredBlock = 'DIPENDENZA';
    } else if (autoKey === 'CHECK_CONFERME') {
      const due = item.values[5] instanceof Date ? new Date(item.values[5]) : null;
      if (!due) {
        desired = 'IN ATTESA';
        desiredBlock = 'DATA';
      } else {
        due.setHours(0,0,0,0);
        if (due > today) {
          desired = 'IN ATTESA';
          desiredBlock = 'DATA';
        } else if (block === 'DATA' || block === 'DIPENDENZA') {
          desired = 'DA FARE';
          desiredBlock = '';
        }
      }
    } else if (autoKey === 'RINGRAZIAMENTO_CIRCOLO' && !ended) {
      desired = 'IN ATTESA';
      desiredBlock = 'DATA';
    } else if (autoKey === 'RINGRAZIAMENTO_CIRCOLO' && ended && (block === 'DATA' || block === 'DIPENDENZA')) {
      desired = 'DA FARE';
      desiredBlock = '';
    } else if (depId && parentDone && block === 'DIPENDENZA') {
      desired = 'DA FARE';
      desiredBlock = '';
    }

    if (normalize_(item.values[6]) !== desired || normalize_(item.values[15]||'') !== desiredBlock) {
      sheet.getRange(item.row,7).setValue(desired);
      if (desiredBlock) sheet.getRange(item.row,16).setValue(desiredBlock);
      else sheet.getRange(item.row,16).clearContent();
      sheet.getRange(item.row,13).setValue(new Date());
    }
  });
}

function refreshTasksV3FromBackend_(eventId, event, child) {
  ensureChecklistBackendHeadersV3_();
  seedPresenceCheckTaskV3_(eventId,event,child);
  ensureTaskNumbersAndDefaultDependenciesV3_(eventId,event);
  recomputeTaskDependenciesV3_(eventId,event);

  const sheet = ensureTaskV3Structure_(child);
  const backend = sh_(APP.SHEETS.CHECKLIST).getDataRange().getValues().slice(1)
    .filter(r=>String(r[1])===String(eventId));
  const byId = {};
  backend.forEach(r=>{ if (r[0]) byId[String(r[0])] = r; });
  backend.sort((a,b)=>Number(a[13]||9999)-Number(b[13]||9999) || Number(a[2]||0)-Number(b[2]||0));

  const out = backend.map(r=>{
    const dep = r[14] ? byId[String(r[14])] : null;
    return [
      r[13]||'', inferTaskProcessV3_(r[3],r[4],r[9]), r[3]||'', dep ? dep[13]||'' : '',
      normalize_(r[6]) === 'COMPLETATA' ? 'FATTO' : (r[6]||'DA FARE'), r[5]||'', r[10]||'',
      r[0]||'', r[8]||'', r[9]||'', r[14]||'', r[11]||'', r[12]||'', r[2]||'', r[15]||''
    ];
  });
  const clearRows = Math.max(Math.min(sheet.getLastRow()-1,TASKS_V3.ENTRY_END_ROW-1), out.length, 1);
  sheet.getRange(2,1,clearRows,TASKS_V3.HEADERS.length).clearContent();
  if (out.length) sheet.getRange(2,1,out.length,TASKS_V3.HEADERS.length).setValues(out);
  applyTaskV3Validations_(sheet);
  applyLiveDependencyFormulasV3_(sheet,out);
  prepareTaskEntryRowsV3_(sheet,out.length);
  return out.length;
}

function syncTasksV3ToBackend_(eventId, event, child) {
  ensureChecklistBackendHeadersV3_();
  const sheet = ensureTaskV3Structure_(child);
  const range = sheet.getRange(1,1,TASKS_V3.ENTRY_END_ROW,TASKS_V3.HEADERS.length);
  const values = range.getValues();
  const formulas = range.getFormulas();
  const oldRows = sh_(APP.SHEETS.CHECKLIST).getDataRange().getValues();
  const oldById = {};
  for (let i=1;i<oldRows.length;i++) {
    if (String(oldRows[i][1])===String(eventId) && oldRows[i][0]) oldById[String(oldRows[i][0])] = oldRows[i];
  }

  const draft = [];
  let maxNo = 0;
  for (let i=1;i<values.length;i++) {
    const r = values[i];
    const activity = String(r[2]||'').trim();
    if (!activity) continue;
    const enteredNo = Number(r[0]||0);
    if (enteredNo > maxNo) maxNo = enteredNo;
    draft.push({sheetRow:i+1, values:r, formulas:formulas[i], no:enteredNo});
  }
  const used = new Set(draft.filter(x=>x.no>0).map(x=>x.no));
  draft.forEach(x=>{
    if (x.no>0) return;
    do { maxNo++; } while (used.has(maxNo));
    x.no = maxNo;
    used.add(maxNo);
  });
  const duplicateNos = new Set();
  const seenNos = new Set();
  draft.forEach(x=>{ if (seenNos.has(x.no)) duplicateNos.add(x.no); seenNos.add(x.no); });
  if (duplicateNos.size) throw new Error('Numeri task duplicati: ' + Array.from(duplicateNos).join(', ') + '.');

  const idByNo = {};
  const draftByNo = {};
  draft.forEach(x=>{
    let id = String(x.values[7]||'').trim();
    if (!id) id = 'TASK-' + Utilities.getUuid();
    x.id = id;
    idByNo[x.no] = id;
    draftByNo[x.no] = x;
  });

  const now = new Date();
  const rows = [];
  draft.forEach(x=>{
    const r = x.values;
    const old = oldById[x.id] ? oldById[x.id].slice(0,16) : new Array(16).fill('');
    const depNo = Number(r[3]||0);
    if (depNo && !idByNo[depNo]) throw new Error('La task n. ' + x.no + ' dipende dalla task n. ' + depNo + ', che non esiste.');
    if (depNo === x.no) throw new Error('La task n. ' + x.no + ' non può dipendere da se stessa.');
    const depId = depNo ? idByNo[depNo] : '';
    let status = normalize_(r[4]) || 'DA FARE';
    if (!TASKS_V3.STATUS.includes(status)) status = 'DA FARE';
    const due = r[5] instanceof Date ? r[5] : '';
    const completed = status === 'FATTO' ? (r[11] instanceof Date ? r[11] : (old[11] instanceof Date ? old[11] : now)) : '';
    const source = String(r[8]||'').trim() || (old[8]||'MANUALE');
    const autoKey = String(r[9]||'').trim() || old[9] || inferTaskAutoKeyV3_(r[2]);
    const order = Number(r[13]||0) || x.no * 10;
    const hasLiveStatusFormula = !!String((x.formulas && x.formulas[4]) || '').trim();

    let autoBlock = '';
    if (status !== 'FATTO') {
      if (depNo && hasLiveStatusFormula) {
        const parent = draftByNo[depNo];
        const parentDone = parent && normalize_(parent.values[4]) === 'FATTO';
        if (!parentDone) autoBlock = 'DIPENDENZA';
        else if (normalize_(autoKey) === 'CHECK_CONFERME' && (!(due instanceof Date) || due > now)) autoBlock = 'DATA';
      } else if (depNo && !hasLiveStatusFormula) {
        autoBlock = 'MANUALE';
      } else if (normalize_(autoKey) === 'RINGRAZIAMENTO_CIRCOLO') {
        autoBlock = String(r[14]||'').trim() || old[15] || '';
      }
    }

    rows.push([
      x.id,eventId,order,String(r[2]||'').trim(),String(r[1]||'').trim()||'ALTRO',due,status,
      old[7]||'',source,autoKey,String(r[6]||'').trim(),completed,now,x.no,depId,autoBlock
    ]);
  });

  replaceCentralRowsForEvent_(sh_(APP.SHEETS.CHECKLIST),eventId,2,rows,16);
  recomputeTaskDependenciesV3_(eventId,event);
  return rows.length;
}
