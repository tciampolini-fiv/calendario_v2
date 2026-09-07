function onOpen() {
  SpreadsheetApp.getUi().createMenu('Scheda evento')
    .addItem('💾 Salva dati nel Calendario','eventSaveToCalendar')
    .addSeparator()
    .addItem('➕ Importa nuova spesa','eventImportExpense')
    .addItem('🔄 Aggiorna cruscotto spese','eventRefreshExpenseDashboard')
    .addSeparator()
    .addItem('📄 Genera documenti','eventGenerateDocuments')
    .addToUi();
}

function eventSaveToCalendar() {
  const ui = SpreadsheetApp.getUi();
  const child = SpreadsheetApp.getActive();
  const meta = eventMeta_();
  const eventId = String(meta.EVENT_ID || '').trim();
  const masterId = String(meta.MASTER_SPREADSHEET_ID || '').trim();
  if (!eventId) throw new Error('ID EVENTO mancante nel foglio _META.');
  if (!masterId) throw new Error('MASTER_SPREADSHEET_ID mancante nel foglio _META.');

  const master = SpreadsheetApp.openById(masterId);
  const counts = {
    participants:eventSaveParticipants_(master,child,eventId),
    tasks:eventSaveTasks_(master,child,eventId),
    expenses:0
  };
  counts.expenses = eventSaveExpenses_(master,child,eventId,meta);

  eventWriteMetaValue_(child.getSheetByName(EVENT_APP.SHEETS.META),'LAST_SYNC',
    Utilities.formatDate(new Date(),Session.getScriptTimeZone() || 'Europe/Rome','dd/MM/yyyy HH:mm'));
  SpreadsheetApp.flush();

  ui.alert(
    'Calendario aggiornato',
    'I dati della Scheda evento sono stati salvati nel Calendario.\n\n' +
    'Attività: ' + counts.tasks + '\n' +
    'Spese: ' + counts.expenses + '\n' +
    'Partecipanti: ' + counts.participants + '\n\n' +
    'La Scheda evento non è stata ricaricata o sovrascritta.',
    ui.ButtonSet.OK
  );
  return counts;
}

function eventMasterSheet_(master,name) {
  const sheet = master.getSheetByName(name);
  if (!sheet) throw new Error('Foglio ' + name + ' non trovato nel Calendario master.');
  return sheet;
}

function eventNormalize_(value) {
  return String(value === null || value === undefined ? '' : value)
    .trim().toUpperCase().replace(/\s+/g,' ');
}

function eventReplaceRowsForEvent_(sheet,eventId,eventColumn,newRows,width) {
  const values = sheet.getDataRange().getValues();
  const targetRows = [];
  for (let i=1;i<values.length;i++) {
    if (String(values[i][eventColumn-1] || '') === String(eventId)) targetRows.push(i+1);
  }
  const reused = Math.min(targetRows.length,newRows.length);
  for (let i=0;i<reused;i++) sheet.getRange(targetRows[i],1,1,width).setValues([newRows[i]]);
  for (let i=reused;i<targetRows.length;i++) sheet.getRange(targetRows[i],1,1,width).clearContent();
  if (newRows.length > reused) {
    const rest = newRows.slice(reused);
    sheet.getRange(sheet.getLastRow()+1,1,rest.length,width).setValues(rest);
  }
}

function eventSaveTasks_(master,child,eventId) {
  const local = child.getSheetByName(EVENT_APP.SHEETS.TASKS);
  if (!local) throw new Error('Foglio Attività non trovato nella Scheda evento.');
  const backend = eventMasterSheet_(master,'_CHECKLIST');
  const oldRows = backend.getDataRange().getValues();
  const oldByNo = {};
  let maxNo = 0;
  for (let i=1;i<oldRows.length;i++) {
    if (String(oldRows[i][1] || '') !== String(eventId)) continue;
    const n = Number(oldRows[i][13] || 0);
    if (n > 0) oldByNo[n] = oldRows[i];
    if (n > maxNo) maxNo = n;
  }

  const lastRow = Math.min(Math.max(local.getLastRow(),2),500);
  const values = local.getRange(2,1,lastRow-1,6).getValues();
  const draft = [];
  values.forEach((r,i)=>{
    const description = String(r[0] || '').trim();
    if (!description) return;
    let no = Number(r[2] || 0);
    if (!(no > 0)) {
      no = ++maxNo;
      local.getRange(i+2,3).setValue(no);
    }
    maxNo = Math.max(maxNo,no);
    draft.push({row:i+2,values:r,no:no});
  });

  const seen = {};
  draft.forEach(x=>{
    if (seen[x.no]) throw new Error('Numero task duplicato: ' + x.no + '.');
    seen[x.no] = true;
  });
  const idByNo = {};
  draft.forEach(x=>{
    const old = oldByNo[x.no];
    x.old = old ? old.slice(0,16) : new Array(16).fill('');
    x.id = old && old[0] ? String(old[0]) : 'TASK-' + Utilities.getUuid();
    idByNo[x.no] = x.id;
  });

  const now = new Date();
  const rows = draft.map((x,index)=>{
    const r = x.values;
    const old = x.old;
    const depNo = Number(r[3] || 0);
    if (depNo && !idByNo[depNo]) throw new Error('La task n. ' + x.no + ' dipende dalla task n. ' + depNo + ', che non esiste.');
    if (depNo === x.no) throw new Error('La task n. ' + x.no + ' non può dipendere da se stessa.');
    let status = eventNormalize_(r[4]) || 'DA FARE';
    if (!['DA FARE','IN ATTESA','FATTO'].includes(status)) status = 'DA FARE';
    const completed = status === 'FATTO' ? (old[11] instanceof Date ? old[11] : now) : '';
    const description = String(r[0] || '').trim();
    return [
      x.id,eventId,(index+1)*10,description,
      old[4] || eventTaskCategory_(description),
      r[5] instanceof Date ? r[5] : '',status,old[7] || '',old[8] || 'MANUALE',old[9] || '',
      String(r[1] || '').trim(),completed,now,x.no,depNo ? idByNo[depNo] : '',
      status !== 'FATTO' && depNo ? 'DIPENDENZA' : ''
    ];
  });
  eventReplaceRowsForEvent_(backend,eventId,2,rows,16);
  return rows.length;
}

function eventTaskCategory_(description) {
  const d = eventNormalize_(description);
  if (/CONVOC|PRESENZ/.test(d)) return 'CONVOCAZIONI';
  if (/HOTEL|ALLOGG|PAST|OSPITAL|GOMMONE|TRASPORT/.test(d)) return 'LOGISTICA';
  if (/PAG|FATTUR|RIF|AFOR|RIMBORS/.test(d)) return 'AMMINISTRAZIONE';
  if (/CIRCOLO|ZONA|COMUNIC/.test(d)) return 'ORGANIZZAZIONE';
  return 'ALTRO';
}

function eventSaveParticipants_(master,child,eventId) {
  const local = child.getSheetByName(EVENT_APP.SHEETS.PARTICIPANTS);
  if (!local) throw new Error('Foglio Partecipanti non trovato nella Scheda evento.');
  const backend = eventMasterSheet_(master,'_PARTECIPANTI');
  const oldRows = backend.getDataRange().getValues();
  const oldById = {};
  for (let i=1;i<oldRows.length;i++) {
    if (String(oldRows[i][1] || '') === String(eventId) && oldRows[i][0]) oldById[String(oldRows[i][0])] = oldRows[i];
  }

  const now = new Date();
  const rows = [];
  const convTech = [];
  const aggTech = [];
  const conv = local.getRange(3,1,40,16).getValues();
  const agg = local.getRange(46,1,20,16).getValues();

  conv.forEach(r=>{
    const parsed = eventParticipantRow_(r,'CONVOCATO',eventId,oldById,now);
    if (!parsed) { convTech.push(['','','','','','']); return; }
    rows.push(parsed.backend);
    convTech.push(parsed.tech);
  });
  agg.forEach(r=>{
    const parsed = eventParticipantRow_(r,'AGGREGATO',eventId,oldById,now);
    if (!parsed) { aggTech.push(['','','','','','']); return; }
    rows.push(parsed.backend);
    aggTech.push(parsed.tech);
  });

  eventReplaceRowsForEvent_(backend,eventId,2,rows,17);
  local.getRange(3,11,convTech.length,6).setValues(convTech);
  local.getRange(46,11,aggTech.length,6).setValues(aggTech);
  return rows.length;
}

function eventParticipantRow_(r,type,eventId,oldById,now) {
  const name = String(r[0] || '').trim();
  const surname = String(r[1] || '').trim();
  if (!name && !surname) return null;
  let id = String(r[10] || '').trim();
  const old = id && oldById[id] ? oldById[id] : null;
  if (!id) id = 'PAR-' + Utilities.getUuid();
  const personId = String(r[11] || '').trim() || (old ? String(old[2] || '') : '');
  const role = eventNormalize_(r[4]) || 'ATLETA';
  let status = eventNormalize_(r[5]);
  if (type === 'CONVOCATO' && !['DA FARE','MANDATA CONVOCAZIONE','CONFERMATO','ASSENTE'].includes(status)) status = 'DA FARE';
  if (type === 'AGGREGATO' && !['DA AUTORIZZARE','AUTORIZZATO','NON AUTORIZZATO'].includes(status)) status = 'DA AUTORIZZARE';
  const provenance = String(r[13] || '').trim() || (old ? String(old[9] || '') : '') || 'SCHEDA EVENTO';
  const sourceFile = String(r[14] || '').trim() || (old ? String(old[10] || '') : '');
  const createdAt = r[15] instanceof Date ? r[15] : (old && old[12] instanceof Date ? old[12] : now);
  const maxRefund = type === 'CONVOCATO' && r[6] !== '' ? Number(r[6] || 0) : '';
  const passed = type === 'CONVOCATO' && r[7] !== '' ? Number(r[7] || 0) : '';
  const passedDate = passed > 0 ? (r[8] instanceof Date ? r[8] : (old && old[16] instanceof Date ? old[16] : now)) : '';
  const notes = String(type === 'CONVOCATO' ? r[9] : r[6] || '').trim();
  const backend = [
    id,eventId,personId,name,surname,String(r[2] || '').trim(),role,String(r[3] || '').trim(),maxRefund,
    provenance,sourceFile,notes,createdAt,type,status,passed,passedDate
  ];
  return {backend:backend,tech:[id,personId,eventId,provenance,sourceFile,createdAt]};
}

function eventSaveExpenses_(master,child,eventId,meta) {
  const local = child.getSheetByName(EVENT_APP.SHEETS.EXPENSES);
  if (!local) throw new Error('Foglio Spese non trovato nella Scheda evento.');
  const backend = eventMasterSheet_(master,'_SPESE');
  const oldRows = backend.getDataRange().getValues();
  const oldById = {};
  for (let i=1;i<oldRows.length;i++) {
    if (String(oldRows[i][1] || '') === String(eventId) && oldRows[i][0]) oldById[String(oldRows[i][0])] = oldRows[i];
  }

  const values = local.getRange(13,1,488,16).getValues();
  const now = new Date();
  const rows = [];
  values.forEach((r,index)=>{
    const idInput = String(r[10] || '').trim();
    if (idInput.indexOf('RIMBORSO-AUTO-') === 0) return;
    const amount = Number(r[0] || 0);
    const description = String(r[1] || '').trim();
    const hasData = amount || description || String(r[2] || '').trim() || String(r[3] || '').trim() || String(r[4] || '').trim();
    if (!hasData) return;
    if (!description) throw new Error('Descrizione mancante nella riga ' + (13+index) + ' del foglio Spese.');

    const id = idInput || 'SPESA-' + Utilities.getUuid();
    const old = oldById[id] ? oldById[id].slice(0,26) : new Array(26).fill('');
    const category = eventExpenseCategory_(r[2],description);
    const movement = eventExpenseMovement_(r[4]);
    const paid = eventNormalize_(r[5]) === 'PAGATO' && movement !== 'PREVENTIVO';
    const budget = movement === 'PREVENTIVO' ? amount : 0;
    const actual = movement === 'PREVENTIVO' ? 0 : amount;
    const rifCode = String(r[6] || '').trim();
    const notes = '[MOVIMENTO=' + movement + ']' + (String(r[9] || '').trim() ? ' ' + String(r[9] || '').trim() : '');
    const createdAt = r[14] instanceof Date ? r[14] : (old[16] instanceof Date ? old[16] : now);
    const paidDate = paid ? (r[13] instanceof Date ? r[13] : (old[12] instanceof Date ? old[12] : now)) : '';
    const status = eventExpenseBackendStatus_(movement,paid);
    const rifNeeded = Math.max(budget,actual) > 1000;
    let rifStatus = old[18] || '';
    if (rifCode) rifStatus = 'RICEVUTO';
    else if (!rifNeeded) rifStatus = 'NON NECESSARIO';
    else if (!rifStatus || eventNormalize_(rifStatus) === 'NON NECESSARIO') rifStatus = 'DA RICHIEDERE';
    const invoiceReceived = movement === 'FATTURA';
    const closed = paid && ['FATTURA','CARTA DI CREDITO','ALTRO PAGAMENTO'].includes(movement);

    old[0]=id; old[1]=eventId; old[2]='SPESA'; old[3]=category;
    old[4]=eventResolveCeb_(master,meta,category,String(r[11] || '').trim() || old[4]);
    old[5]=description; old[6]=String(r[3] || '').trim(); old[7]=String(r[12] || '').trim();
    old[8]=budget; old[9]=actual; old[10]=''; old[11]=status; old[12]=paidDate; old[13]='';
    old[14]=String(r[8] || '').trim(); old[15]=notes; old[16]=createdAt; old[17]=now;
    old[18]=rifStatus; old[20]=rifCode; old[21]=invoiceReceived;
    old[22]=invoiceReceived ? (old[22] || now) : ''; old[25]=closed ? (old[25] || now) : '';
    rows.push(old);
    if (!idInput) local.getRange(13+index,11).setValue(id);
    if (!String(r[11] || '').trim() && old[4]) local.getRange(13+index,12).setValue(old[4]);
  });

  const participants = child.getSheetByName(EVENT_APP.SHEETS.PARTICIPANTS);
  if (participants) {
    const pRows = participants.getRange(3,1,40,16).getValues();
    pRows.forEach((p,index)=>{
      const name = String(p[0] || '').trim();
      const surname = String(p[1] || '').trim();
      const passed = Number(p[7] || 0);
      if ((!name && !surname) || !(passed > 0)) return;
      const participantId = String(p[10] || '').trim() || ('ROW-' + (index+3));
      const id = 'RIMBORSO-AUTO-' + participantId;
      const old = oldById[id] ? oldById[id].slice(0,26) : new Array(26).fill('');
      const beneficiary = [name,surname].filter(Boolean).join(' ');
      old[0]=id; old[1]=eventId; old[2]='RIMBORSO'; old[3]='RIMBORSO'; old[4]='CEB.002';
      old[5]='Rimborso ' + beneficiary; old[6]=beneficiary; old[7]=String(p[11] || '').trim();
      old[8]=0; old[9]=passed; old[10]=''; old[11]='RIMBORSATO'; old[12]=p[8] instanceof Date ? p[8] : now;
      old[13]=''; old[14]=''; old[15]='[MOVIMENTO=RIMBORSO] [AUTO_RIMBORSO=' + participantId + ']' + (String(p[9] || '').trim() ? ' ' + String(p[9] || '').trim() : '');
      old[16]=old[16] || now; old[17]=now; old[18]='NON NECESSARIO'; old[20]=''; old[21]=false; old[25]=old[25] || now;
      rows.push(old);
    });
  }

  eventReplaceRowsForEvent_(backend,eventId,2,rows,26);
  eventRefreshExpenseDashboard();
  return rows.length;
}

function eventExpenseCategory_(category,description) {
  let c = eventNormalize_(category);
  const d = eventNormalize_(description);
  if (c === 'VITTO / ALLOGGIO') {
    if (/PASTI|VITTO|RISTORANTE|PRANZO|CENA/.test(d)) return 'VITTO';
    return 'ALLOGGIO';
  }
  if (c === 'VIAGGIO') c = 'VIAGGI';
  return c || 'ALTRO';
}

function eventExpenseMovement_(movement) {
  const m = eventNormalize_(movement || 'PREVENTIVO');
  if (m === 'CARTA') return 'CARTA DI CREDITO';
  return ['PREVENTIVO','AFOR','FATTURA','CARTA DI CREDITO','ALTRO PAGAMENTO'].includes(m) ? m : 'PREVENTIVO';
}

function eventExpenseBackendStatus_(movement,paid) {
  if (movement === 'PREVENTIVO') return 'DA DEFINIRE';
  if (!paid) return 'DA PAGARE';
  if (movement === 'AFOR') return 'PAGATO - AFOR';
  if (movement === 'FATTURA') return 'PAGATO - FATTURA';
  if (movement === 'CARTA DI CREDITO') return 'PAGATO CON CC';
  return 'PAGATO';
}

function eventResolveCeb_(master,meta,category,current) {
  if (current) return current;
  const type = eventNormalize_(meta.EVENT_TYPE || '');
  let profile = type;
  const typeCfg = master.getSheetByName('_CONFIG_TIPI_EVENTO');
  if (typeCfg) {
    const rows = typeCfg.getDataRange().getValues();
    for (let i=1;i<rows.length;i++) {
      if (eventNormalize_(rows[i][0]) === type) {
        profile = eventNormalize_(rows[i][4]) || type;
        break;
      }
    }
  }
  if (profile === 'FOIL ACADEMY') return 'CEB.033';
  const cfg = master.getSheetByName('_CONFIG_CEB');
  if (!cfg) return eventNormalize_(category) === 'RIMBORSO' ? 'CEB.002' : '';
  const rows = cfg.getDataRange().getValues();
  const options = [];
  for (let i=1;i<rows.length;i++) {
    const active = rows[i][4] === true || eventNormalize_(rows[i][4]) === 'SI';
    if (!active) continue;
    const allowed = String(rows[i][2] || '').split(';').map(eventNormalize_).filter(Boolean);
    if (!(allowed.includes('TUTTI') || allowed.includes(type) || allowed.includes(profile))) continue;
    if (eventNormalize_(rows[i][3] || 'SPESA') === 'RIMBORSO') continue;
    options.push({code:String(rows[i][0] || ''),cats:String(rows[i][6] || '').split(';').map(eventNormalize_).filter(Boolean)});
  }
  const cat = eventNormalize_(category);
  const exact = options.find(o=>o.cats.includes(cat));
  if (exact) return exact.code;
  const generic = options.find(o=>o.cats.includes('TUTTI'));
  if (generic) return generic.code;
  const travel = options.find(o=>o.code === 'CEB.001');
  return travel ? travel.code : (options[0] ? options[0].code : '');
}

function eventWriteMetaValue_(sheet,key,value) {
  if (!sheet) return;
  const rows = sheet.getRange(1,1,Math.max(sheet.getLastRow(),1),2).getValues();
  const target = eventNormalize_(key);
  for (let i=0;i<rows.length;i++) {
    if (eventNormalize_(rows[i][0]) === target) {
      sheet.getRange(i+1,2).setValue(value);
      return;
    }
  }
  sheet.getRange(sheet.getLastRow()+1,1,1,2).setValues([[key,value]]);
}