const EVENT_APP = Object.freeze({
  SHEETS: {
    TASKS: 'Attività',
    EXPENSES: 'Spese',
    PARTICIPANTS: 'Partecipanti',
    DOCUMENTS: 'DOCUMENTI',
    META: '_META'
  },
  MASTER_DOCUMENTS: '_DOCUMENTI',
  PARTICIPANTS: { CONV_START: 3, CONV_COUNT: 20, AGG_START: 25, AGG_COUNT: 20 },
  ACTIVITY: {
    HEADER: 1,
    START: 2,
    MAX: 500,
    COLS: 19,
    TYPE: 8,
    OBJECTIVE_ID: 9,
    TASK_ID: 10,
    OBJECTIVE_ORDER: 11,
    STEP_ORDER: 12,
    COLOR: 13,
    AUTO_DUE: 14,
    PREVIOUS_ID: 15,
    COMPLETED_AT: 16,
    PATH: 17,
    DUE_MODE: 18,
    OFFSET: 19
  },
  TEMPLATE_SPECS: [
    { key: 'CONV_ATLETI', label: 'Convocazione atleti', id: '1SpOMrpDwe8aTW9QAyby865WufnD40AsPx6t8zm6Yw3E' },
    { key: 'CONV_TECNICO', label: 'Convocazione tecnico', id: '1Xb61H5TQ0avz8n5_THn-BSd1Xd4YBtmOD3Lu2uwUuJw' },
    { key: 'GOMMONE', label: 'Richiesta gommone alla Zona', id: '1uvK4J8WrWekpMrkPTkpUvsdYTggBhwzCDJaJmyIhUko' },
    { key: 'OSPITALITA', label: 'Richiesta ospitalità circolo', id: '1NuheJqcwtLxp8muTdbnpcZc22Z8HJgA-jbdt_Sq6ruk' },
    { key: 'RINGRAZIAMENTO', label: 'Ringraziamento circolo', id: '11Fqn8U3Hs8MgJPH0tQTKDEozmWDftPtmELoLjgMFg7c' }
  ]
});

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const name = sheet.getName();

  if (name === EVENT_APP.SHEETS.TASKS) {
    eventRepairActivityLayoutV15_(sheet, false);
    eventHandleActivitiesEditV14_(e);
    return;
  }
  if (name === EVENT_APP.SHEETS.EXPENSES) {
    eventHandleExpenseEditV8_(e);
    return;
  }
  if (name === EVENT_APP.SHEETS.PARTICIPANTS) {
    const row = e.range.getRow();
    if ((row >= 3 && row <= 22) || (row >= 25 && row <= 44)) eventRefreshExpenseDashboardV8_();
  }
}

function onSelectionChange(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== EVENT_APP.SHEETS.TASKS || e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < 2) return;

  const type = eventRowTypeV15_(sh, row);
  const aText = String(sh.getRange(row, 1).getDisplayValue() || '').trim();
  const cText = String(sh.getRange(row, 3).getDisplayValue() || '').trim();

  if (col === 1 && (type === 'AGGIUNGI_OBIETTIVO' || /^＋?\s*AGGIUNGI OBIETTIVO/i.test(aText))) {
    eventInsertObjectiveV14_(sh, row);
    return;
  }

  if (col === 3 && type === 'OBIETTIVO' && cText.indexOf('＋') >= 0) {
    eventInsertTaskV14_(sh, row);
  }
}

function eventInitializeV7_() {
  const meta = eventMeta_();
  if (!String(meta.EVENT_ID || '').trim()) return false;

  const activity = SpreadsheetApp.getActive().getSheetByName(EVENT_APP.SHEETS.TASKS);
  if (activity) eventRepairActivityLayoutV15_(activity, true);

  eventApplyCommitmentV8_(meta);
  eventPopulateTechniciansV7_(meta);
  eventRefreshActivitiesV14_();
  eventFormatActivityRowsV13_();
  eventSyncExpenseTasksV8_();
  eventRefreshExpenseDashboardV8_();
  return true;
}

function eventActivityPaletteV14_() {
  return ['#D9EAF7', '#FCE8B2', '#EADCF8', '#D9EAD3', '#F4CCCC', '#D0E0E3', '#FCE5CD', '#D9D2E9', '#CFE2F3', '#E2F0D9'];
}

function eventAddDaysV14_(d, n) {
  if (!(d instanceof Date)) return '';
  const x = new Date(d);
  x.setDate(x.getDate() + Number(n || 0));
  return x;
}

function eventDayV14_(d) {
  if (!(d instanceof Date)) return null;
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}

function eventMetaDateV14_(key) {
  return eventParseIsoDate_(eventMeta_()[key] || '');
}

function eventDefaultObjectiveDueV14_() {
  return eventAddDaysV14_(eventMetaDateV14_('EVENT_START'), -5);
}

function eventActivityRowsV14_(sh) {
  const last = Math.min(EVENT_APP.ACTIVITY.MAX, Math.max(sh.getLastRow(), 2));
  return sh.getRange(2, 1, last - 1, EVENT_APP.ACTIVITY.COLS).getValues().map(function (v, i) {
    return { row: i + 2, v: v };
  });
}

function eventRowTypeV15_(sh, row) {
  const raw = eventNormalizeTextV7_(sh.getRange(row, EVENT_APP.ACTIVITY.TYPE).getDisplayValue());
  if (raw === 'ADD_OBJECTIVE' || raw === 'AGGIUNGI OBIETTIVO') return 'AGGIUNGI_OBIETTIVO';
  if (raw) return raw;

  const a = String(sh.getRange(row, 1).getDisplayValue() || '').trim();
  const c = String(sh.getRange(row, 3).getDisplayValue() || '').trim();
  if (/^＋?\s*AGGIUNGI OBIETTIVO/i.test(a)) return 'AGGIUNGI_OBIETTIVO';
  if (a && c.indexOf('＋') >= 0) return 'OBIETTIVO';
  if (c && c.indexOf('＋') < 0) return 'TASK';
  return '';
}

function eventInferPathV15_(objectiveName, taskName) {
  const o = eventNormalizeTextV7_(objectiveName);
  const t = eventNormalizeTextV7_(taskName);

  if (o.indexOf('VITTO') >= 0 || o.indexOf('ALLOGGIO') >= 0) {
    if (/PAST/.test(t)) return 'PASTI';
    if (/SOGGIORNO|HOTEL|ALLOGGIO|FATTURA/.test(t)) return 'SOGGIORNO';
  }
  if (o.indexOf('CONVOCAZ') >= 0) {
    if (/AGGREGAT/.test(t)) return 'AGGREGATI';
    if (/TECNIC/.test(t)) return 'TECNICI';
    if (/ATLET/.test(t)) return 'ATLETI';
  }
  if (o.indexOf('CIRCOLO') >= 0) return 'CIRCOLO';
  if (o.indexOf('GOMMONE') >= 0) return 'GOMMONE';
  if (o.indexOf('VIAGGIO') >= 0) return 'VIAGGIO';
  if (o.indexOf('ZONA') >= 0) return 'ZONA';
  return 'MANUALE';
}

function eventRepairActivityLayoutV15_(sh, force) {
  if (!sh) return;

  if (sh.getMaxColumns() < EVENT_APP.ACTIVITY.COLS) {
    sh.insertColumnsAfter(sh.getMaxColumns(), EVENT_APP.ACTIVITY.COLS - sh.getMaxColumns());
  }

  sh.getRange(1, 8, 1, 12).setValues([[
    'TIPO RIGA', 'ID OBIETTIVO', 'ID TASK', 'ORDINE OBIETTIVO', 'ORDINE STEP', 'COLORE',
    'SCADENZA AUTOMATICA', 'ID TASK PRECEDENTE', 'DATA COMPLETAMENTO', 'PERCORSO',
    'MODALITA SCADENZA', 'OFFSET GIORNI'
  ]]);

  const maxRow = Math.min(EVENT_APP.ACTIVITY.MAX, Math.max(sh.getLastRow(), 2));
  let currentObjective = null;
  let maxObjectiveOrder = 0;
  let lastDataRow = 1;
  let actionRow = 0;
  let pathLastTask = {};
  let nextStep = 10;

  for (let row = 2; row <= maxRow; row++) {
    const a = String(sh.getRange(row, 1).getDisplayValue() || '').trim();
    const c = String(sh.getRange(row, 3).getDisplayValue() || '').trim();
    let type = eventRowTypeV15_(sh, row);

    if (type === 'AGGIUNGI_OBIETTIVO') {
      actionRow = row;
      sh.getRange(row, 1).setValue('＋ AGGIUNGI OBIETTIVO');
      sh.getRange(row, EVENT_APP.ACTIVITY.TYPE).setValue('AGGIUNGI_OBIETTIVO');
      if (force) sh.getRange(row, 9, 1, 11).clearContent();
      lastDataRow = Math.max(lastDataRow, row);
      continue;
    }

    if (type === 'OBIETTIVO' || (a && c.indexOf('＋') >= 0)) {
      sh.getRange(row, EVENT_APP.ACTIVITY.TYPE).setValue('OBIETTIVO');
      if (!sh.getRange(row, EVENT_APP.ACTIVITY.OBJECTIVE_ID).getValue()) {
        sh.getRange(row, EVENT_APP.ACTIVITY.OBJECTIVE_ID).setValue('OBJ-' + Utilities.getUuid());
      }
      let order = Number(sh.getRange(row, EVENT_APP.ACTIVITY.OBJECTIVE_ORDER).getValue() || 0);
      if (!(order > 0)) {
        order = maxObjectiveOrder + 10;
        sh.getRange(row, EVENT_APP.ACTIVITY.OBJECTIVE_ORDER).setValue(order);
      }
      maxObjectiveOrder = Math.max(maxObjectiveOrder, order);

      let color = String(sh.getRange(row, EVENT_APP.ACTIVITY.COLOR).getDisplayValue() || '').trim();
      const visibleColor = String(sh.getRange(row, 1).getBackground() || '').toLowerCase();
      if (!color) color = visibleColor && visibleColor !== '#ffffff' ? visibleColor : eventActivityPaletteV14_()[Math.max(0, Math.round(order / 10) - 1) % eventActivityPaletteV14_().length];
      sh.getRange(row, EVENT_APP.ACTIVITY.COLOR).setValue(color);
      sh.getRange(row, 3).setValue('＋').setHorizontalAlignment('center');
      sh.getRange(row, 1, 1, 7).setBackground(color).setFontWeight('bold');
      sh.getRange(row, 2).setNumberFormat('dd/MM/yyyy');

      currentObjective = {
        row: row,
        id: String(sh.getRange(row, EVENT_APP.ACTIVITY.OBJECTIVE_ID).getValue() || ''),
        name: a,
        order: order,
        color: color
      };
      pathLastTask = {};
      nextStep = 10;
      lastDataRow = Math.max(lastDataRow, row);
      continue;
    }

    if ((type === 'TASK' || c) && currentObjective && c && c.indexOf('＋') < 0) {
      sh.getRange(row, EVENT_APP.ACTIVITY.TYPE).setValue('TASK');
      sh.getRange(row, EVENT_APP.ACTIVITY.OBJECTIVE_ID).setValue(currentObjective.id);
      if (!sh.getRange(row, EVENT_APP.ACTIVITY.TASK_ID).getValue()) {
        sh.getRange(row, EVENT_APP.ACTIVITY.TASK_ID).setValue('TASK-' + Utilities.getUuid());
      }
      sh.getRange(row, EVENT_APP.ACTIVITY.OBJECTIVE_ORDER).setValue(currentObjective.order);
      if (!sh.getRange(row, EVENT_APP.ACTIVITY.COLOR).getValue()) sh.getRange(row, EVENT_APP.ACTIVITY.COLOR).setValue(currentObjective.color);

      let step = Number(sh.getRange(row, EVENT_APP.ACTIVITY.STEP_ORDER).getValue() || 0);
      if (!(step > 0)) {
        step = nextStep;
        sh.getRange(row, EVENT_APP.ACTIVITY.STEP_ORDER).setValue(step);
      }
      nextStep = Math.max(nextStep, step + 10);

      let path = String(sh.getRange(row, EVENT_APP.ACTIVITY.PATH).getDisplayValue() || '').trim();
      if (!path) {
        path = eventInferPathV15_(currentObjective.name, c);
        sh.getRange(row, EVENT_APP.ACTIVITY.PATH).setValue(path);
      }

      const taskId = String(sh.getRange(row, EVENT_APP.ACTIVITY.TASK_ID).getValue() || '');
      let previousId = String(sh.getRange(row, EVENT_APP.ACTIVITY.PREVIOUS_ID).getDisplayValue() || '').trim();
      if (!previousId && pathLastTask[path]) {
        previousId = pathLastTask[path];
        sh.getRange(row, EVENT_APP.ACTIVITY.PREVIOUS_ID).setValue(previousId);
      }
      pathLastTask[path] = taskId;

      let mode = eventNormalizeTextV7_(sh.getRange(row, EVENT_APP.ACTIVITY.DUE_MODE).getDisplayValue());
      if (!mode) {
        mode = previousId ? 'PRECEDENTE' : 'MANUAL';
        sh.getRange(row, EVENT_APP.ACTIVITY.DUE_MODE).setValue(mode);
      }
      if (sh.getRange(row, EVENT_APP.ACTIVITY.OFFSET).isBlank()) {
        sh.getRange(row, EVENT_APP.ACTIVITY.OFFSET).setValue(previousId ? 2 : 0);
      }

      sh.getRange(row, 4).setNumberFormat('dd/MM/yyyy');
      sh.getRange(row, EVENT_APP.ACTIVITY.COMPLETED_AT).setNumberFormat('dd/MM/yyyy HH:mm');
      const done = sh.getRange(row, 5).getValue() === true;
      sh.getRange(row, 5).insertCheckboxes().setValue(done);
      sh.getRange(row, 1, 1, 7).setBackground('#ffffff').setFontWeight('normal');
      lastDataRow = Math.max(lastDataRow, row);
      continue;
    }
  }

  if (!actionRow) {
    actionRow = Math.min(Math.max(lastDataRow + 1, 2), EVENT_APP.ACTIVITY.MAX);
    if (String(sh.getRange(actionRow, 1, 1, EVENT_APP.ACTIVITY.COLS).getDisplayValues()[0].join('')).trim()) {
      sh.insertRowAfter(lastDataRow);
      actionRow = lastDataRow + 1;
    }
    sh.getRange(actionRow, 1, 1, EVENT_APP.ACTIVITY.COLS).clearContent().clearDataValidations();
    sh.getRange(actionRow, 1).setValue('＋ AGGIUNGI OBIETTIVO');
    sh.getRange(actionRow, EVENT_APP.ACTIVITY.TYPE).setValue('AGGIUNGI_OBIETTIVO');
  }

  sh.getRange(actionRow, 1).setValue('＋ AGGIUNGI OBIETTIVO').setFontWeight('bold');
  sh.getRange(actionRow, EVENT_APP.ACTIVITY.TYPE).setValue('AGGIUNGI_OBIETTIVO');
  sh.getRange(actionRow, 9, 1, 11).clearContent();
  sh.getRange(2, 2, Math.max(1, actionRow - 1), 1).setNumberFormat('dd/MM/yyyy');
  sh.getRange(2, 4, Math.max(1, actionRow - 1), 1).setNumberFormat('dd/MM/yyyy');
  try { sh.hideColumns(8, 12); } catch (err) { console.log('Nascondi colonne tecniche: ' + err.message); }
}

function eventObjectiveTasksV14_(sh, oid) {
  return eventActivityRowsV14_(sh)
    .filter(function (x) { return eventNormalizeTextV7_(x.v[7]) === 'TASK' && String(x.v[8] || '') === String(oid); })
    .sort(function (a, b) { return Number(a.v[11] || 0) - Number(b.v[11] || 0); });
}

function eventFindObjectiveHeaderRowV14_(sh, row) {
  for (let r = row; r >= 2; r--) {
    if (eventRowTypeV15_(sh, r) === 'OBIETTIVO') return r;
  }
  return 0;
}

function eventTaskMapV14_(sh, oid) {
  const out = {};
  eventObjectiveTasksV14_(sh, oid).forEach(function (x) {
    if (x.v[9]) out[String(x.v[9])] = x;
  });
  return out;
}

function eventEnsureObjectiveRowV14_(sh, row) {
  const name = String(sh.getRange(row, 1).getDisplayValue() || '').trim();
  if (!name || /^＋?\s*AGGIUNGI OBIETTIVO/i.test(name)) return;

  sh.getRange(row, EVENT_APP.ACTIVITY.TYPE).setValue('OBIETTIVO');
  if (!sh.getRange(row, EVENT_APP.ACTIVITY.OBJECTIVE_ID).getValue()) sh.getRange(row, EVENT_APP.ACTIVITY.OBJECTIVE_ID).setValue('OBJ-' + Utilities.getUuid());

  if (!(sh.getRange(row, 2).getValue() instanceof Date)) {
    const due = eventDefaultObjectiveDueV14_();
    if (due) sh.getRange(row, 2).setValue(due).setNumberFormat('dd/MM/yyyy');
  }

  if (!(Number(sh.getRange(row, EVENT_APP.ACTIVITY.OBJECTIVE_ORDER).getValue()) > 0)) {
    const vals = eventActivityRowsV14_(sh)
      .filter(function (x) { return eventNormalizeTextV7_(x.v[7]) === 'OBIETTIVO'; })
      .map(function (x) { return Number(x.v[10] || 0); })
      .filter(function (n) { return n > 0; });
    sh.getRange(row, EVENT_APP.ACTIVITY.OBJECTIVE_ORDER).setValue(vals.length ? Math.max.apply(null, vals) + 10 : 10);
  }

  let color = String(sh.getRange(row, 1).getBackground() || '').toLowerCase();
  if (!color || color === '#ffffff') {
    const p = eventActivityPaletteV14_();
    const n = eventActivityRowsV14_(sh).filter(function (x) { return eventNormalizeTextV7_(x.v[7]) === 'OBIETTIVO'; }).length;
    color = p[Math.max(0, n - 1) % p.length];
  }
  sh.getRange(row, EVENT_APP.ACTIVITY.COLOR).setValue(color);
  sh.getRange(row, 3).setValue('＋').setHorizontalAlignment('center');
  sh.getRange(row, 1, 1, 7).setBackground(color).setFontWeight('bold');
}

function eventEnsureTaskRowV14_(sh, row) {
  const task = String(sh.getRange(row, 3).getDisplayValue() || '').trim();
  if (!task) return;

  const header = eventFindObjectiveHeaderRowV14_(sh, row);
  if (!header) throw new Error('Aggiungi prima un obiettivo.');
  eventEnsureObjectiveRowV14_(sh, header);

  const oid = String(sh.getRange(header, EVENT_APP.ACTIVITY.OBJECTIVE_ID).getValue() || '');
  sh.getRange(row, EVENT_APP.ACTIVITY.TYPE).setValue('TASK');
  sh.getRange(row, EVENT_APP.ACTIVITY.OBJECTIVE_ID).setValue(oid);
  if (!sh.getRange(row, EVENT_APP.ACTIVITY.TASK_ID).getValue()) sh.getRange(row, EVENT_APP.ACTIVITY.TASK_ID).setValue('TASK-' + Utilities.getUuid());
  sh.getRange(row, EVENT_APP.ACTIVITY.OBJECTIVE_ORDER).setValue(sh.getRange(header, EVENT_APP.ACTIVITY.OBJECTIVE_ORDER).getValue());

  if (!sh.getRange(row, EVENT_APP.ACTIVITY.PATH).getValue()) {
    sh.getRange(row, EVENT_APP.ACTIVITY.PATH).setValue(eventInferPathV15_(sh.getRange(header, 1).getDisplayValue(), task));
  }

  const tasks = eventObjectiveTasksV14_(sh, oid).filter(function (x) { return x.row !== row; });
  const path = String(sh.getRange(row, EVENT_APP.ACTIVITY.PATH).getValue() || 'MANUALE');
  const same = tasks.filter(function (x) { return String(x.v[16] || 'MANUALE') === path; });
  let step = Number(sh.getRange(row, EVENT_APP.ACTIVITY.STEP_ORDER).getValue() || 0);
  if (!(step > 0)) {
    step = tasks.length ? Math.max.apply(null, tasks.map(function (x) { return Number(x.v[11] || 0); })) + 10 : 10;
    sh.getRange(row, EVENT_APP.ACTIVITY.STEP_ORDER).setValue(step);
  }

  const prev = same.length ? same[same.length - 1] : null;
  sh.getRange(row, EVENT_APP.ACTIVITY.PREVIOUS_ID).setValue(prev ? prev.v[9] : '');
  if (prev && !sh.getRange(row, EVENT_APP.ACTIVITY.DUE_MODE).getValue()) {
    sh.getRange(row, EVENT_APP.ACTIVITY.DUE_MODE).setValue('PRECEDENTE');
    sh.getRange(row, EVENT_APP.ACTIVITY.OFFSET).setValue(2);
  } else if (!prev && !sh.getRange(row, EVENT_APP.ACTIVITY.DUE_MODE).getValue()) {
    sh.getRange(row, EVENT_APP.ACTIVITY.DUE_MODE).setValue('MANUAL');
    sh.getRange(row, EVENT_APP.ACTIVITY.OFFSET).setValue(0);
  }

  const done = sh.getRange(row, 5).getValue() === true;
  sh.getRange(row, 5).insertCheckboxes().setValue(done);
}

function eventHandleActivitiesEditV14_(e) {
  const sh = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < 2 || row > EVENT_APP.ACTIVITY.MAX) return;

  const type = eventRowTypeV15_(sh, row);
  if (type === 'OBIETTIVO') {
    if (col === 1 || col === 2) {
      eventEnsureObjectiveRowV14_(sh, row);
      eventRefreshActivitiesV14_();
    }
    return;
  }

  if (type === 'TASK') {
    if (col === 3) eventEnsureTaskRowV14_(sh, row);
    if (col === 4) {
      sh.getRange(row, EVENT_APP.ACTIVITY.AUTO_DUE).setValue(false);
      sh.getRange(row, EVENT_APP.ACTIVITY.DUE_MODE).setValue('MANUAL');
      sh.getRange(row, EVENT_APP.ACTIVITY.OFFSET).setValue(0);
    }
    if (col === 5) eventToggleTaskV14_(sh, row, e.value === true || String(e.value).toUpperCase() === 'TRUE');
    eventRefreshActivitiesV14_();
  }
}

function eventToggleTaskV14_(sh, row, checked) {
  eventEnsureTaskRowV14_(sh, row);
  const oid = String(sh.getRange(row, EVENT_APP.ACTIVITY.OBJECTIVE_ID).getValue() || '');
  const tid = String(sh.getRange(row, EVENT_APP.ACTIVITY.TASK_ID).getValue() || '');
  const map = eventTaskMapV14_(sh, oid);
  const current = map[tid];
  if (!current) return;

  const prevId = String(current.v[14] || '');
  const prev = prevId ? map[prevId] : null;
  if (checked && prev && !Boolean(prev.v[4])) {
    sh.getRange(row, 5).setValue(false);
    SpreadsheetApp.getUi().alert('Attività ancora bloccata', 'Completa prima l’attività precedente nello stesso percorso.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const descendants = [];
  let cursor = tid;
  while (cursor) {
    const next = Object.keys(map).map(function (key) { return map[key]; }).find(function (x) { return String(x.v[14] || '') === cursor; });
    if (!next) break;
    descendants.push(next);
    cursor = String(next.v[9] || '');
  }

  if (!checked && descendants.some(function (x) { return Boolean(x.v[4]); })) {
    sh.getRange(row, 5).setValue(true);
    SpreadsheetApp.getUi().alert('Sequenza non valida', 'Prima riapri le attività successive già completate nello stesso percorso.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  if (checked) {
    const now = new Date();
    sh.getRange(row, 6).setValue('FATTO');
    sh.getRange(row, EVENT_APP.ACTIVITY.COMPLETED_AT).setValue(now).setNumberFormat('dd/MM/yyyy HH:mm');
    const next = descendants[0];
    if (next && !(next.v[3] instanceof Date) && eventNormalizeTextV7_(next.v[17]) === 'PRECEDENTE') {
      const due = eventAddDaysV14_(now, Number(next.v[18] || 2));
      sh.getRange(next.row, 4).setValue(due).setNumberFormat('dd/MM/yyyy');
      sh.getRange(next.row, EVENT_APP.ACTIVITY.AUTO_DUE).setValue(true);
    }
  } else {
    sh.getRange(row, EVENT_APP.ACTIVITY.COMPLETED_AT).clearContent();
    descendants.forEach(function (x) {
      if (x.v[13] === true && eventNormalizeTextV7_(x.v[17]) === 'PRECEDENTE') {
        sh.getRange(x.row, 4).clearContent();
        sh.getRange(x.row, EVENT_APP.ACTIVITY.AUTO_DUE).setValue(false);
      }
    });
  }
}

function eventRefreshActivitiesV14_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(EVENT_APP.SHEETS.TASKS);
  if (!sh) return;
  eventRepairActivityLayoutV15_(sh, false);

  const rows = eventActivityRowsV14_(sh);
  const today = eventDayV14_(new Date());
  const objectives = {};

  rows.forEach(function (x) {
    if (eventNormalizeTextV7_(x.v[7]) === 'OBIETTIVO' && x.v[8]) objectives[String(x.v[8])] = { row: x.row, due: x.v[1], tasks: [] };
  });
  rows.forEach(function (x) {
    if (eventNormalizeTextV7_(x.v[7]) === 'TASK' && x.v[8] && objectives[String(x.v[8])]) objectives[String(x.v[8])].tasks.push(x);
  });

  Object.keys(objectives).forEach(function (oid) {
    const o = objectives[oid];
    const map = {};
    o.tasks.forEach(function (x) { if (x.v[9]) map[String(x.v[9])] = x; });

    o.tasks.forEach(function (x) {
      const checked = Boolean(sh.getRange(x.row, 5).getValue());
      if (checked) {
        sh.getRange(x.row, 6).setValue('FATTO');
        return;
      }

      const previousId = String(x.v[14] || '');
      const prev = previousId ? map[previousId] : null;
      const prevDone = !prev || Boolean(sh.getRange(prev.row, 5).getValue());
      let due = sh.getRange(x.row, 4).getValue();
      const mode = eventNormalizeTextV7_(x.v[17] || '');
      const offset = Number(x.v[18] || 2);

      if (prevDone && !(due instanceof Date) && mode === 'PRECEDENTE' && prev) {
        const completed = prev.v[15] instanceof Date ? prev.v[15] : new Date();
        due = eventAddDaysV14_(completed, offset);
        sh.getRange(x.row, 4).setValue(due).setNumberFormat('dd/MM/yyyy');
        sh.getRange(x.row, EVENT_APP.ACTIVITY.AUTO_DUE).setValue(true);
      }

      let state = 'IN ATTESA';
      if (prevDone) {
        const d = eventDayV14_(due);
        state = !d || d <= today ? 'DA FARE' : 'IN ATTESA';
      }
      sh.getRange(x.row, 6).setValue(state);
    });

    const done = o.tasks.filter(function (x) { return Boolean(sh.getRange(x.row, 5).getValue()); }).length;
    let state = o.tasks.length && done === o.tasks.length ? 'COMPLETATO' : (done ? 'IN CORSO' : 'DA AVVIARE');
    const od = eventDayV14_(o.due);
    if (state !== 'COMPLETATO' && od && od < today) state = 'IN RITARDO';
    sh.getRange(o.row, 6).setValue(state);
  });
}

function eventInsertObjectiveV14_(sh, actionRow) {
  sh.insertRowBefore(actionRow);
  const row = actionRow;
  sh.getRange(row, 1, 1, EVENT_APP.ACTIVITY.COLS).clearContent().clearDataValidations();
  sh.getRange(row, 1, 1, 7).setBackground('#ffffff').setFontWeight('normal');
  sh.getRange(row, 1).setValue('Nuovo obiettivo');
  sh.getRange(row, 3).setValue('＋');
  sh.getRange(row, EVENT_APP.ACTIVITY.TYPE).setValue('OBIETTIVO');
  eventEnsureObjectiveRowV14_(sh, row);
  eventRepairActivityLayoutV15_(sh, false);
  sh.setActiveRange(sh.getRange(row, 1));
}

function eventInsertTaskV14_(sh, headerRow) {
  eventEnsureObjectiveRowV14_(sh, headerRow);
  const oid = String(sh.getRange(headerRow, EVENT_APP.ACTIVITY.OBJECTIVE_ID).getValue() || '');
  const tasks = eventObjectiveTasksV14_(sh, oid);
  const insertAfter = tasks.length ? tasks[tasks.length - 1].row : headerRow;
  sh.insertRowAfter(insertAfter);

  const row = insertAfter + 1;
  sh.getRange(row, 1, 1, EVENT_APP.ACTIVITY.COLS).clearContent().clearDataValidations();
  sh.getRange(row, 1, 1, 7).setBackground('#ffffff').setFontWeight('normal');
  sh.getRange(row, 3).setValue('Nuova attività');
  sh.getRange(row, EVENT_APP.ACTIVITY.TYPE).setValue('TASK');
  sh.getRange(row, EVENT_APP.ACTIVITY.OBJECTIVE_ID).setValue(oid);
  sh.getRange(row, EVENT_APP.ACTIVITY.PATH).setValue('MANUALE');
  eventEnsureTaskRowV14_(sh, row);
  eventRepairActivityLayoutV15_(sh, false);
  eventRefreshActivitiesV14_();
  sh.setActiveRange(sh.getRange(row, 3));
}

function eventFormatActivityRowsV13_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(EVENT_APP.SHEETS.TASKS);
  if (!sh) return;
  eventRepairActivityLayoutV15_(sh, false);
}

function eventRepairActivities() {
  const sh = SpreadsheetApp.getActive().getSheetByName(EVENT_APP.SHEETS.TASKS);
  if (!sh) throw new Error('Foglio Attività non trovato.');
  eventRepairActivityLayoutV15_(sh, true);
  eventRefreshActivitiesV14_();
  SpreadsheetApp.getActive().toast('Tabella Attività riparata.', 'Scheda evento', 4);
}

function eventHandleActivitiesEditV13_(e) { return eventHandleActivitiesEditV14_(e); }
function eventRefreshActivitiesV13_() { return eventRefreshActivitiesV14_(); }
function eventEnsureObjectiveRowV13_(sh, row) { return eventEnsureObjectiveRowV14_(sh, row); }
function eventEnsureTaskRowV13_(sh, row) { return eventEnsureTaskRowV14_(sh, row); }
function eventToggleTaskV13_(sh, row, checked) { return eventToggleTaskV14_(sh, row, checked); }

function eventNormalizeTextV7_(value) {
  return String(value === null || value === undefined ? '' : value)
    .trim()
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ');
}

function eventMeta_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(EVENT_APP.SHEETS.META);
  if (!sheet) throw new Error('Foglio _META non trovato.');
  const values = sheet.getRange(1, 1, sheet.getLastRow(), 2).getDisplayValues();
  const out = {};
  values.forEach(function (r) { if (r[0]) out[String(r[0]).trim()] = r[1]; });
  return out;
}

function eventWriteLocalMetaV7_(key, value) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(EVENT_APP.SHEETS.META);
  if (!sheet) return;
  const rows = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), 2).getDisplayValues();
  const target = eventNormalizeTextV7_(key);
  for (let i = 0; i < rows.length; i++) {
    if (eventNormalizeTextV7_(rows[i][0]) === target) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

function eventCommitmentMapV7_() {
  return {
    'RIUNIONI/MISSIONI|*': '2-1',
    'ALLENAMENTO|29ER': '28',
    'ALLENAMENTO|420': '29',
    'ALLENAMENTO|NACRA 15': '30',
    'ALLENAMENTO|IQFOIL': '31',
    'ALLENAMENTO|ILCA 6': '32',
    'ALLENAMENTO|ILCA 4': '33-1',
    'ALLENAMENTO INTERZONALE|ILCA 4': '33-2',
    'ALLENAMENTO|KITEFOIL': '34',
    'ALLENAMENTO|WING': '35',
    'TEST FISICI|*': '37',
    'COLLEGIALE|*': '38',
    'REGATA|29ER': '39-1',
    'REGATA|420': '40',
    'REGATA|NACRA 15': '41',
    'REGATA|IQFOIL': '42',
    'REGATA|ILCA 6': '43',
    'REGATA|ILCA 4': '44',
    'REGATA|KITEFOIL': '45',
    'REGATA|WING': '46',
    'OSS.REGATE ITA|*': '36',
    'CAMP. NAZ GIOVANILE SINGOLO|*': '48-1',
    'CAMP. NAZ GIOVANILE DOPPIO|*': '48-1',
    'FOIL ACADEMY|*': '50',
    'ISCRIZIONI REGATE|29ER': '47-1',
    'ISCRIZIONI REGATE|420': '47-2',
    'ISCRIZIONI REGATE|NACRA 15': '47-3',
    'ISCRIZIONI REGATE|IQFOIL': '47-4',
    'ISCRIZIONI REGATE|ILCA 6': '47-5',
    'ISCRIZIONI REGATE|ILCA 4': '47-6',
    'ISCRIZIONI REGATE|KITEFOIL': '47-7',
    'ISCRIZIONI REGATE|WING': '47-8',
    'ISCRIZIONE YWC|TUTTE': '171',
    'MANUTENZIONE ALLENAMENTI|*': '39-3',
    'MANUTENZIONE REGATE|*': '39-4',
    'NOLEGGI ALLENAMENTI|*': '52',
    'NOLEGGI REGATE|*': '51'
  };
}

function eventCanonicalTypeV7_(value) {
  const t = eventNormalizeTextV7_(value);
  const aliases = {
    'RIUNIONE/MISSIONE': 'RIUNIONI/MISSIONI',
    'RIUNIONI / MISSIONI': 'RIUNIONI/MISSIONI',
    'OSSERVAZIONE REGATA': 'OSS.REGATE ITA',
    'ISCRIZIONE REGATA': 'ISCRIZIONI REGATE'
  };
  return aliases[t] || t;
}

function eventCanonicalClassV7_(value) {
  const c = eventNormalizeTextV7_(value);
  return c === 'KITE' ? 'KITEFOIL' : c;
}

function eventResolveCommitmentV7_(meta) {
  const type = eventCanonicalTypeV7_(meta.EVENT_TYPE || '');
  const cls = eventCanonicalClassV7_(meta.EVENT_CLASS || '');
  const map = eventCommitmentMapV7_();
  return map[type + '|' + cls] || map[type + '|*'] || map[type + '|TUTTE'] || '';
}

function eventTechnicianDirectoryV7_() {
  return [
    ['ZAGGIA', 'Leonardo', 'Zaggia'],
    ['CRISI', 'Andrea', 'Crisi'],
    ['RAVEGLIA', 'Matteo', 'Raveglia'],
    ['CARICATO', 'Francesco', 'Caricato'],
    ['PICCIAU', 'Gianluigi', 'Picciau'],
    ['SENSINI', 'Alessandra', 'Sensini'],
    ['NICOLUCCI', 'Matteo', 'Nicolucci'],
    ['CAMBONI', 'Mattia', 'Camboni'],
    ['CANGEMI', 'Antonino', 'Cangemi'],
    ['LOPERFIDO', 'Daniel', 'Loperfido']
  ];
}

function eventPopulateTechniciansV7_(meta) {
  const raw = eventNormalizeTextV7_(meta.EVENT_TECHNICIANS || '');
  if (!raw) return [];
  const found = eventTechnicianDirectoryV7_().filter(function (x) { return raw.indexOf(x[0]) >= 0; });
  if (!found.length) return [];

  const full = found.map(function (x) { return x[1] + ' ' + x[2]; });
  eventWriteLocalMetaV7_('EVENT_TECHNICIANS', full.join(', '));

  const sheet = SpreadsheetApp.getActive().getSheetByName(EVENT_APP.SHEETS.PARTICIPANTS);
  if (!sheet) return full;

  const rows = sheet.getRange(3, 1, 20, 19).getDisplayValues();
  const existing = new Set(rows.map(function (r) { return eventNormalizeTextV7_((r[0] || '') + ' ' + (r[1] || '')); }));
  found.forEach(function (x) {
    const key = eventNormalizeTextV7_(x[1] + ' ' + x[2]);
    if (existing.has(key)) return;
    let target = -1;
    for (let i = 0; i < rows.length; i++) {
      if (!String(rows[i][0] || '').trim() && !String(rows[i][1] || '').trim()) {
        target = i + 3;
        rows[i][0] = x[1];
        rows[i][1] = x[2];
        break;
      }
    }
    if (target < 0) return;
    sheet.getRange(target, 1).setValue(x[1]);
    sheet.getRange(target, 2).setValue(x[2]);
    sheet.getRange(target, 8).setValue('TECNICO');
    sheet.getRange(target, 9).setValue('CONFERMATO');
    existing.add(key);
  });
  return full;
}

function eventDocumentReplacements_(meta) {
  const participants = SpreadsheetApp.getActive().getSheetByName(EVENT_APP.SHEETS.PARTICIPANTS);
  const rows = participants ? participants.getRange(3, 1, 20, 19).getDisplayValues() : [];
  const athletes = rows.filter(function (r) {
    const has = String(r[0] || '').trim() || String(r[1] || '').trim();
    const role = eventNormalizeTextV7_(r[7] || 'ATLETA');
    const status = eventNormalizeTextV7_(r[8]);
    return has && role === 'ATLETA' && status !== 'ASSENTE';
  });
  const athleteLines = athletes.map(function (r) {
    const name = [r[0], r[1]].filter(Boolean).join(' ').trim();
    return r[2] ? name + ' – ' + r[2] : name;
  });
  const start = eventParseIsoDate_(meta.EVENT_START);
  const end = eventParseIsoDate_(meta.EVENT_END);

  return {
    'tipo': eventTitleCase_(meta.EVENT_TYPE || ''),
    'classe': meta.EVENT_CLASS || '',
    'luogo': meta.EVENT_LOCATION || '',
    'localita': meta.EVENT_LOCATION || '',
    'zona': meta.EVENT_ZONE || '',
    'circolo': meta.EVENT_CLUB || '',
    'tecnico': meta.EVENT_TECHNICIANS || '',
    'lista tecnici': meta.EVENT_TECHNICIANS || '',
    'lista atleti': athleteLines.join('\n'),
    'numero atleti': athletes.length ? String(athletes.length) : '',
    'hotel/struttura': meta.EVENT_LODGING || '',
    'data inizio': eventFormatDate_(start),
    'data fine': eventFormatDate_(end),
    'data in': eventFormatDate_(start),
    'data out': eventFormatDate_(end),
    'date': eventFormatDateRange_(start, end),
    'data di oggi': eventFormatDate_(new Date()),
    'data oggi': eventFormatDate_(new Date()),
    'impegno': meta.EVENT_COMMITMENT || ''
  };
}

function eventFillDocument_(docId, replacements, key) {
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  const reps = Object.assign({}, replacements);
  if (key === 'GOMMONE') reps.N = replacements.zona || '';
  if (key === 'OSPITALITA') reps.N = replacements['numero atleti'] || '';

  Object.keys(reps).forEach(function (k) {
    const v = reps[k];
    if (v === '' || v === null || v === undefined) return;
    body.replaceText('(?i)\\{\\{' + eventEscapeRegex_(k) + '\\}\\}', eventSafeReplacement_(v));
  });
  if (reps.classe) body.replaceText('(?i)\\{classe\\}', eventSafeReplacement_(reps.classe));
  doc.saveAndClose();
}

function eventDocumentName_(key, r) {
  const base = {
    CONV_ATLETI: ['Convocazione atleti', r.tipo, r.classe, r.luogo, r.date],
    CONV_TECNICO: ['Convocazione tecnico', r.tecnico, r.tipo, r.classe, r.date],
    GOMMONE: ['Richiesta gommone', r.zona ? r.zona + ' Zona' : '', r.tipo, r.classe, r.date],
    OSPITALITA: ['Richiesta ospitalità', r.circolo, r.tipo, r.classe, r.date],
    RINGRAZIAMENTO: ['Ringraziamento', r.circolo, r.tipo, r.classe, r.date]
  }[key] || [r.tipo, r.classe, r.date];
  return base.filter(Boolean).join(' - ');
}

function eventRegisterDocument_(meta, spec, file) {
  if (!meta.MASTER_SPREADSHEET_ID) return;
  const master = SpreadsheetApp.openById(meta.MASTER_SPREADSHEET_ID);
  const sheet = master.getSheetByName(EVENT_APP.MASTER_DOCUMENTS);
  if (!sheet) return;
  sheet.appendRow(['DOC-' + Utilities.getUuid(), meta.EVENT_ID, spec.key, spec.id, file.getName(), file.getId(), file.getUrl(), 'BOZZA MODIFICABILE', new Date(), '']);
}

function eventParseIsoDate_(s) {
  const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0) : null;
}

function eventFormatDate_(d) {
  if (!(d instanceof Date) || isNaN(d)) return '';
  const months = ['gennaio', 'febbraio', 'marzo', 'aprile', 'maggio', 'giugno', 'luglio', 'agosto', 'settembre', 'ottobre', 'novembre', 'dicembre'];
  return d.getDate() + ' ' + months[d.getMonth()] + ' ' + d.getFullYear();
}

function eventFormatDateRange_(s, e) {
  if (!(s instanceof Date)) return '';
  if (!(e instanceof Date)) return eventFormatDate_(s);
  if (s.getTime() === e.getTime()) return eventFormatDate_(s);
  const months = ['gennaio', 'febbraio', 'marzo', 'aprile', 'maggio', 'giugno', 'luglio', 'agosto', 'settembre', 'ottobre', 'novembre', 'dicembre'];
  if (s.getFullYear() === e.getFullYear() && s.getMonth() === e.getMonth()) return s.getDate() + '-' + e.getDate() + ' ' + months[s.getMonth()] + ' ' + s.getFullYear();
  return eventFormatDate_(s) + ' - ' + eventFormatDate_(e);
}

function eventTitleCase_(s) {
  return String(s || '').toLowerCase().replace(/(^|\s|[-/])([a-zà-öø-ÿ])/g, function (_, p, c) { return p + c.toUpperCase(); });
}

function eventEscapeRegex_(s) {
  return String(s || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function eventSafeReplacement_(s) {
  return String(s || '').replace(/\\/g, '\\\\').replace(/\$/g, '\\$');
}
