const EVENT_MIGRATION_V11 = Object.freeze({
  TEMPLATE_ID: EVENT_SHEET_TEMPLATE_ID_V3,
  TASK_MAX_ROWS: 500,
  V11_CONVOCATI_START: 3,
  V11_CONVOCATI_END: 17,
  V11_AGGREGATI_START: 20,
  V11_AGGREGATI_END: 39,
  PARTICIPANT_COLS: 19,
  HISTORY_START: 20,
  HISTORY_COLS: 18,
  PAYMENT_TYPES: Object.freeze(['AFOR','CARTA DI CREDITO','SALDO FATTURA','ALTRO PAGAMENTO']),
  CATEGORIES: Object.freeze(['VIAGGI','VITTO','ALLOGGIO','NOLEGGI','ISCRIZIONI','ALTRO'])
});

/**
 * Migra la Scheda evento collegata alla riga selezionata al modello V11.
 * Il file originale non viene modificato nei contenuti: resta come backup.
 * La nuova Scheda e una copia completa del modello corrente, quindi eredita
 * anche il progetto Apps Script bound del modello.
 */
function migrateSelectedEventSheetToV11() {
  const ui = SpreadsheetApp.getUi();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const event = selectedEvent_();
    const eventId = ensureEventId_(event);
    const oldChild = getExplicitLinkedEventSheetV3_(event);
    if (!oldChild) throw new Error('La riga selezionata non ha una Scheda evento collegata.');
    if (isEventSheetV11_(oldChild)) {
      ui.alert('Scheda gia aggiornata','La Scheda evento collegata usa gia il layout V11. Non e stata apportata alcuna modifica.',ui.ButtonSet.OK);
      return {skipped:true,id:oldChild.getId(),url:oldChild.getUrl()};
    }

    const result = migrateEventSheetToV11_(eventId,event,oldChild);
    ui.alert(
      'Migrazione V11 completata',
      'La nuova Scheda evento e stata creata e collegata al Calendario.\n\n' +
      'Attivita trasferite: ' + result.tasks + '\n' +
      'Convocati trasferiti: ' + result.convocati + '\n' +
      'Aggregati trasferiti: ' + result.aggregati + '\n' +
      'Movimenti spese trasferiti: ' + result.expenses + '\n\n' +
      'Il vecchio file e stato conservato come BACKUP PRE-V11.',
      ui.ButtonSet.OK
    );
    return result;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Migra tutte le Schede legacy effettivamente collegate a righe del Calendario.
 * Le V11 vengono saltate. Ogni file viene gestito in modo indipendente.
 */
function migrateAllLegacyEventSheetsToV11() {
  const ui = SpreadsheetApp.getUi();
  const cal = sh_(APP.SHEETS.CALENDAR);
  const values = cal.getDataRange().getValues();
  if (values.length < 2) return {migrated:0,skipped:0,errors:[]};
  const headers = values[0], idCol = headers.indexOf(APP.CALENDAR_HEADERS.ID);
  const linkCol = headers.indexOf(APP.CALENDAR_HEADERS.EVENT_SHEET);
  if (idCol < 0 || linkCol < 0) throw new Error('Colonne ID EVENTO / SCHEDA EVENTO non trovate nel Calendario.');

  let migrated = 0, skipped = 0;
  const errors = [];
  for (let i=1;i<values.length;i++) {
    const eventId = String(values[i][idCol]||'').trim();
    if (!eventId) continue;
    const cell = cal.getRange(i+1,linkCol+1), rich = cell.getRichTextValue();
    const url = (rich&&rich.getLinkUrl()) || cell.getDisplayValue();
    const oldId = extractDriveId_(url);
    if (!oldId) continue;
    try {
      const oldChild = SpreadsheetApp.openById(oldId);
      if (isEventSheetV11_(oldChild)) { skipped++; continue; }
      const event = {};
      headers.forEach((h,c)=>event[h]=values[i][c]);
      event._row = i+1;
      migrateEventSheetToV11_(eventId,event,oldChild);
      migrated++;
    } catch (err) {
      errors.push('Riga '+(i+1)+': '+(err.message||String(err)));
    }
  }
  ui.alert(
    'Migrazione Schede evento',
    'Convertite: '+migrated+'\nGia V11 / saltate: '+skipped+'\nErrori: '+errors.length+
    (errors.length?'\n\n'+errors.slice(0,8).join('\n'):''),
    ui.ButtonSet.OK
  );
  return {migrated:migrated,skipped:skipped,errors:errors};
}

function migrateEventSheetToV11_(eventId,event,oldChild) {
  if (!oldChild) throw new Error('Scheda evento sorgente non disponibile.');
  if (isEventSheetV11_(oldChild)) return {skipped:true,id:oldChild.getId(),url:oldChild.getUrl(),tasks:0,convocati:0,aggregati:0,expenses:0};

  const snapshot = migrationExtractSnapshotV11_(oldChild,event);
  migrationValidateSnapshotV11_(snapshot);
  const folderId = migrationResolveFolderIdV11_(oldChild,eventId,event);
  const folder = DriveApp.getFolderById(folderId);
  const oldFile = DriveApp.getFileById(oldChild.getId());
  const originalName = oldFile.getName();
  const stamp = Utilities.formatDate(new Date(),APP.TZ,'yyyyMMdd-HHmmss');
  const tempName = '__MIGRAZIONE V11__ ' + originalName + ' ' + stamp;
  let newFile = null, newChild = null, linkChanged = false;

  try {
    newFile = DriveApp.getFileById(EVENT_MIGRATION_V11.TEMPLATE_ID).makeCopy(tempName,folder);
    newChild = SpreadsheetApp.openById(newFile.getId());
    newChild.setSpreadsheetLocale('it_IT');
    newChild.setSpreadsheetTimeZone(APP.TZ);
    if (!isEventSheetV11_(newChild)) throw new Error('Il modello corrente non risulta V11. Aggiorna prima MODELLO - Scheda evento.');

    migrationApplySnapshotV11_(newChild,snapshot,eventId,event,folderId);
    migrationVerifyTargetV11_(newChild,snapshot);

    writeEventMetaV5_(newChild,eventId,event,folderId);
    const meta = newChild.getSheetByName(EVENT_SHEET.SHEETS.META);
    if (meta) writeMeta_(meta,{
      SYNC_VERSION:'11',
      MIGRATED_FROM_SHEET_ID:oldChild.getId(),
      MIGRATED_AT:Utilities.formatDate(new Date(),APP.TZ,'dd/MM/yyyy HH:mm')
    });
    ensureFastEventTaskTriggerV9_(newChild);
    SpreadsheetApp.flush();

    setEventSheetLink_(event._row,newChild.getUrl());
    SpreadsheetApp.flush();
    linkChanged = true;

    newFile.setName(originalName);
    oldFile.setName('BACKUP PRE-V11 - '+originalName+' - '+stamp);

    return {
      migrated:true,
      id:newChild.getId(),
      url:newChild.getUrl(),
      backupId:oldChild.getId(),
      tasks:snapshot.tasks.values.filter(r=>migrationRowHasDataV11_(r)).length,
      convocati:snapshot.participants.convocati.length,
      aggregati:snapshot.participants.aggregati.length,
      expenses:snapshot.expenses.rows.length,
      expenseVersion:snapshot.expenses.sourceVersion
    };
  } catch (err) {
    if (newFile && !linkChanged) {
      try { newFile.setTrashed(true); } catch (cleanupErr) {}
    }
    throw err;
  }
}

function migrationExtractSnapshotV11_(oldChild,event) {
  return {
    sourceId:oldChild.getId(),
    sourceName:oldChild.getName(),
    tasks:migrationExtractTasksV11_(oldChild),
    participants:migrationExtractParticipantsV11_(oldChild),
    expenses:migrationExtractExpensesV11_(oldChild,event)
  };
}

function migrationValidateSnapshotV11_(snapshot) {
  const convMax = EVENT_MIGRATION_V11.V11_CONVOCATI_END-EVENT_MIGRATION_V11.V11_CONVOCATI_START+1;
  const aggMax = EVENT_MIGRATION_V11.V11_AGGREGATI_END-EVENT_MIGRATION_V11.V11_AGGREGATI_START+1;
  if (snapshot.participants.convocati.length > convMax) {
    throw new Error('La vecchia scheda contiene '+snapshot.participants.convocati.length+' convocati, ma il modello V11 ne contiene '+convMax+'. Nessuna migrazione eseguita.');
  }
  if (snapshot.participants.aggregati.length > aggMax) {
    throw new Error('La vecchia scheda contiene '+snapshot.participants.aggregati.length+' aggregati, ma il modello V11 ne contiene '+aggMax+'. Nessuna migrazione eseguita.');
  }
  if (snapshot.expenses.rows.length > 300) throw new Error('Numero movimenti spesa anomalo: '+snapshot.expenses.rows.length+'. Migrazione interrotta per sicurezza.');
}

function migrationResolveFolderIdV11_(oldChild,eventId,event) {
  const meta = oldChild.getSheetByName(EVENT_SHEET.SHEETS.META);
  const fromMeta = meta ? String(readMetaValue_(meta,'EVENT_FOLDER_ID')||'').trim() : '';
  if (fromMeta) {
    try { DriveApp.getFolderById(fromMeta).getName(); return fromMeta; } catch (e) {}
  }
  try {
    const parents = DriveApp.getFileById(oldChild.getId()).getParents();
    if (parents.hasNext()) return parents.next().getId();
  } catch (e) {}
  return createWorkFolderForEvent_(eventId,event,event._row).folderId;
}

function migrationExtractTasksV11_(child) {
  const sheet = child.getSheetByName(EVENT_SHEET.SHEETS.TASKS);
  if (!sheet) return {values:[],bNotes:[]};
  const last = Math.min(EVENT_MIGRATION_V11.TASK_MAX_ROWS,Math.max(1,sheet.getLastRow()));
  if (last < 2) return {values:[],bNotes:[]};
  const values = sheet.getRange(2,1,last-1,6).getValues();
  const notes = sheet.getRange(2,2,last-1,1).getNotes().map(r=>r[0]||'');
  let end = values.length;
  while (end>0 && !migrationRowHasDataV11_(values[end-1]) && !notes[end-1]) end--;
  return {values:values.slice(0,end),bNotes:notes.slice(0,end)};
}

function migrationExtractParticipantsV11_(child) {
  const sheet = child.getSheetByName('Partecipanti');
  if (!sheet) return {convocati:[],aggregati:[]};
  const aggTitle = migrationFindLabelRowV11_(sheet,'AGGREGATI',1,Math.min(100,sheet.getMaxRows()));
  const header = sheet.getRange(2,1,1,Math.min(19,sheet.getMaxColumns())).getDisplayValues()[0].map(normalize_);
  const modern = header[7] === 'RUOLO';
  const legacy = header[4] === 'RUOLO';
  if (!modern && !legacy) throw new Error('Layout Partecipanti della vecchia scheda non riconosciuto.');
  const convEnd = aggTitle>3 ? aggTitle-1 : Math.min(42,sheet.getLastRow());
  const convRaw = convEnd>=3 ? sheet.getRange(3,1,convEnd-2,Math.min(modern?19:16,sheet.getMaxColumns())).getValues() : [];
  const aggStart = aggTitle>0 ? aggTitle+2 : 46;
  const aggEnd = Math.max(aggStart-1,sheet.getLastRow());
  const aggRaw = aggEnd>=aggStart ? sheet.getRange(aggStart,1,aggEnd-aggStart+1,Math.min(modern?19:16,sheet.getMaxColumns())).getValues() : [];
  const conv = convRaw.filter(r=>String(r[0]||'').trim()||String(r[1]||'').trim()).map(r=>modern?migrationPadRowV11_(r,19):migrationParticipantLegacyToModernV11_(r,true));
  const agg = aggRaw.filter(r=>String(r[0]||'').trim()||String(r[1]||'').trim()).map(r=>modern?migrationPadRowV11_(r,19):migrationParticipantLegacyToModernV11_(r,false));
  return {convocati:conv,aggregati:agg};
}

function migrationParticipantLegacyToModernV11_(r,isConvocato) {
  const out = new Array(19).fill('');
  out[0]=r[0]||''; out[1]=r[1]||''; out[2]=r[2]||'';
  out[6]=r[3]||''; out[7]=r[4]||''; out[8]=r[5]||'';
  if (isConvocato) {
    out[9]=r[6]||''; out[10]=r[7]||''; out[11]=r[8]||''; out[12]=r[9]||'';
  } else out[12]=r[6]||'';
  for (let i=0;i<6;i++) out[13+i]=r[10+i]||'';
  return out;
}

function migrationExtractExpensesV11_(child,event) {
  const sheet = child.getSheetByName(EVENT_SHEET.SHEETS.EXPENSES);
  if (!sheet) return {sourceVersion:'NONE',rows:[],initialBudget:0,imp:''};
  const sourceVersion = migrationExpenseVersionV11_(sheet);
  const initialBudget = Number(sheet.getRange('B2').getValue()||0);
  const meta = child.getSheetByName(EVENT_SHEET.SHEETS.META);
  const metaImp = meta ? String(readMetaValue_(meta,'EVENT_COMMITMENT')||'').trim() : '';
  let rows = [];
  if (sourceVersion === 'FLAT_V4') rows = migrationConvertFlatExpensesV11_(sheet);
  else if (sourceVersion === 'TWO_PAYMENT_V7') rows = migrationConvertTwoPaymentExpensesV11_(sheet);
  else if (sourceVersion === 'GROUPED_V10') rows = migrationConvertGroupedExpensesV11_(sheet);
  else if (sourceVersion === 'V11') rows = migrationConvertGroupedExpensesV11_(sheet);
  else if (sheet.getLastRow()>1) throw new Error('Layout Spese della vecchia scheda non riconosciuto.');

  migrationRecalculateQuoteStatusesArrayV11_(rows);
  rows.sort((a,b)=>{
    const qa=Number(a[0]||0),qb=Number(b[0]||0);if(qa!==qb)return qa-qb;
    const ra=normalize_(a[1])==='PREVENTIVO'?0:1,rb=normalize_(b[1])==='PREVENTIVO'?0:1;if(ra!==rb)return ra-rb;
    const da=a[16] instanceof Date?a[16].getTime():0,db=b[16] instanceof Date?b[16].getTime():0;return da-db;
  });
  const impFromSheet = sourceVersion==='FLAT_V4' ? String(sheet.getRange('H9').getDisplayValue()||'').trim() :
    (sourceVersion==='TWO_PAYMENT_V7' ? String(sheet.getRange('N9').getDisplayValue()||'').trim() : String(sheet.getRange('G9').getDisplayValue()||'').trim());
  const eventImp = String(event&&event[APP.CALENDAR_HEADERS.COMMITMENT]||'').trim();
  return {sourceVersion:sourceVersion,rows:rows,initialBudget:initialBudget,imp:eventImp||metaImp||impFromSheet};
}

function migrationExpenseVersionV11_(sheet) {
  if (isEventSheetV11_({getSheetByName:function(name){return name===EVENT_SHEET.SHEETS.EXPENSES?sheet:null;}})) return 'V11';
  const a8=normalize_(sheet.getRange('A8').getDisplayValue()),e8=normalize_(sheet.getRange('E8').getDisplayValue());
  if (a8==='IMPORTO' && e8==='TIPO') return 'FLAT_V4';
  if (a8==='PREVENTIVO' && e8==='PAGAMENTO 1') return 'TWO_PAYMENT_V7';
  if (normalize_(sheet.getRange('A19').getDisplayValue())==='N. PREVENTIVO' && normalize_(sheet.getRange('B19').getDisplayValue())==='TIPO PAGAMENTO') return 'GROUPED_V10';
  return 'UNKNOWN';
}

function migrationConvertFlatExpensesV11_(sheet) {
  const records = [];
  const last = Math.max(13,sheet.getLastRow());
  [9].concat(Array.from({length:Math.max(0,last-12)},(_,i)=>i+13)).forEach(row=>{
    const r=sheet.getRange(row,1,1,16).getValues()[0];
    const id=String(r[10]||'').trim();
    if (id.indexOf('RIMBORSO-AUTO-')===0) return;
    const amount=Number(r[0]||0),description=String(r[1]||'').trim(),category=String(r[2]||'').trim(),supplier=migrationCleanSupplierV11_(r[3]);
    const movement=migrationPaymentTypeV11_(r[4]) || (normalize_(r[4])==='PREVENTIVO'?'PREVENTIVO':'');
    const status=normalize_(r[5]);
    const hasData=amount>0||description||category||supplier||movement||String(r[6]||'').trim()||String(r[8]||'').trim();
    if(!hasData)return;
    records.push({
      amount:amount,description:description||'Spesa migrata',category:migrationCategoryV11_(category,description),rawCategory:category,supplier:supplier,
      movement:movement||(status?'ALTRO PAGAMENTO':'PREVENTIVO'),status:status,rif:String(r[6]||'').trim(),imp:String(r[7]||'').trim(),
      documents:migrationCellTextOrLinkV11_(sheet,row,9),notes:migrationLegacyNoteV11_(r[9],supplier,category),id:id,ceb:String(r[11]||'').trim(),personId:String(r[12]||'').trim(),
      paymentDate:r[13] instanceof Date?r[13]:'',createdAt:r[14] instanceof Date?r[14]:new Date(),updatedAt:r[15] instanceof Date?r[15]:new Date()
    });
  });
  return migrationBuildGroupedFromFlatV11_(records);
}

function migrationBuildGroupedFromFlatV11_(records) {
  const rows=[],quotes=[];let nextNo=1;
  records.forEach(rec=>{
    if(rec.movement==='PREVENTIVO'){
      const id=rec.id||('PREV-MIG-'+Utilities.getUuid());
      const q={number:nextNo++,id:id,budget:Math.max(0,rec.amount),allocated:0,description:rec.description,category:rec.category,supplier:rec.supplier,rif:rec.rif,imp:rec.imp,documents:rec.documents,notes:rec.notes,ceb:rec.ceb,personId:rec.personId,createdAt:rec.createdAt,updatedAt:rec.updatedAt};
      rows.push([q.number,'PREVENTIVO',q.budget,q.description,'DA PAGARE','','',q.rif,q.imp,q.documents,q.notes,q.category,id,id,q.ceb,q.personId,q.createdAt,q.updatedAt]);
      quotes.push(q);return;
    }
    const amount=Math.max(0,rec.amount);if(!(amount>0))return;
    let q=migrationFindParentQuoteV11_(quotes,rec,amount);
    if(!q){
      const qid='PREV-MIG-'+Utilities.getUuid();
      q={number:nextNo++,id:qid,budget:amount,allocated:0,description:rec.description,category:rec.category,supplier:rec.supplier,rif:rec.rif,imp:rec.imp,documents:rec.documents,notes:rec.notes,ceb:rec.ceb,personId:rec.personId,createdAt:rec.createdAt,updatedAt:rec.updatedAt};
      rows.push([q.number,'PREVENTIVO',q.budget,q.description,'DA PAGARE','','',q.rif,q.imp,q.documents,q.notes,q.category,qid,qid,q.ceb,q.personId,q.createdAt,q.updatedAt]);
      quotes.push(q);
    }
    q.allocated+=amount;
    const paid=rec.status==='PAGATO';
    rows.push([q.number,rec.movement,amount,rec.description,paid?'PAGATO':'DA PAGARE',paid?rec.paymentDate:'','',q.rif||rec.rif,q.imp||rec.imp,q.documents||rec.documents,q.notes||rec.notes,q.category,rec.id||('PAG-MIG-'+Utilities.getUuid()),q.id,rec.ceb||q.ceb,rec.personId||q.personId,rec.createdAt,rec.updatedAt]);
  });
  return rows;
}

function migrationFindParentQuoteV11_(quotes,rec,amount) {
  const open=quotes.filter(q=>q.category===rec.category && q.allocated+amount<=q.budget+0.005);
  if(!open.length)return null;
  if(rec.supplier){
    const exact=open.filter(q=>q.supplier&&normalize_(q.supplier)===normalize_(rec.supplier));
    if(exact.length)return exact[exact.length-1];
  }
  return open.length===1?open[0]:null;
}

function migrationConvertTwoPaymentExpensesV11_(sheet) {
  const rows=[];let quoteNo=1;
  const last=Math.max(13,sheet.getLastRow());
  [9].concat(Array.from({length:Math.max(0,last-12)},(_,i)=>i+13)).forEach(row=>{
    const r=sheet.getRange(row,1,1,18).getValues()[0];
    const budget=Number(r[0]||0),description=String(r[1]||'').trim(),categoryRaw=String(r[2]||'').trim(),supplier=migrationCleanSupplierV11_(r[3]);
    const p1Amount=Number(r[5]||0),p2Amount=Number(r[8]||0),hasData=budget>0||description||categoryRaw||p1Amount>0||p2Amount>0;
    if(!hasData)return;
    const category=migrationCategoryV11_(categoryRaw,description),rif=String(r[12]||'').trim(),imp=String(r[13]||'').trim(),documents=migrationCellTextOrLinkV11_(sheet,row,15),notes=migrationLegacyNoteV11_(r[15],supplier,categoryRaw);
    const created=new Date(),updated=new Date(),qid=String(r[16]||'').trim()||('PREV-MIG-'+Utilities.getUuid()),ceb=String(r[17]||'').trim();
    const payments=[];
    if(p1Amount>0||String(r[4]||'').trim())payments.push({type:migrationPaymentTypeV11_(r[4])||'ALTRO PAGAMENTO',amount:p1Amount,date:r[6] instanceof Date?r[6]:'',due:''});
    if(p2Amount>0||String(r[7]||'').trim())payments.push({type:migrationPaymentTypeV11_(r[7])||'ALTRO PAGAMENTO',amount:p2Amount,date:r[9] instanceof Date?r[9]:'',due:''});
    const allocated=payments.reduce((s,p)=>s+Math.max(0,p.amount),0),finalBudget=Math.max(budget,allocated);
    const due=r[11] instanceof Date?r[11]:'';
    for(let i=payments.length-1;i>=0;i--)if(!payments[i].date&&due){payments[i].due=due;break;}
    rows.push([quoteNo,'PREVENTIVO',finalBudget,description||'Spesa migrata','DA PAGARE','','',rif,imp,documents,notes,category,qid,qid,ceb,'',created,updated]);
    payments.forEach(p=>{if(!(p.amount>0))return;const paid=p.date instanceof Date;rows.push([quoteNo,p.type,p.amount,description||'Spesa migrata',paid?'PAGATO':'DA PAGARE',paid?p.date:'',p.due,rif,imp,documents,notes,category,'PAG-MIG-'+Utilities.getUuid(),qid,ceb,'',created,updated]);});
    quoteNo++;
  });
  return rows;
}

function migrationConvertGroupedExpensesV11_(sheet) {
  const last=sheet.getLastRow();if(last<EVENT_MIGRATION_V11.HISTORY_START)return[];
  const raw=sheet.getRange(EVENT_MIGRATION_V11.HISTORY_START,1,last-EVENT_MIGRATION_V11.HISTORY_START+1,EVENT_MIGRATION_V11.HISTORY_COLS).getValues(),rows=[];
  raw.forEach(r=>{
    const q=Number(r[0]||0),movement=normalize_(r[1]);if(!(q>0)||!movement)return;
    const out=migrationPadRowV11_(r,18),now=new Date();
    out[1]=movement==='FATTURA'?'SALDO FATTURA':movement;
    out[3]=String(out[3]||'').trim()||'Spesa migrata';
    out[4]=normalize_(out[4])||'DA PAGARE';
    out[11]=migrationCategoryV11_(out[11],out[3]);
    out[12]=String(out[12]||'').trim()||((out[1]==='PREVENTIVO'?'PREV-MIG-':'PAG-MIG-')+Utilities.getUuid());
    if(out[1]==='PREVENTIVO')out[13]=String(out[13]||'').trim()||out[12];
    out[16]=out[16] instanceof Date?out[16]:now;out[17]=out[17] instanceof Date?out[17]:now;
    rows.push(out);
  });
  const quoteIds={};rows.forEach(r=>{if(r[1]==='PREVENTIVO')quoteIds[Number(r[0])]=String(r[13]||r[12]);});
  rows.forEach(r=>{if(r[1]!=='PREVENTIVO'&&!String(r[13]||'').trim())r[13]=quoteIds[Number(r[0])]||'';});
  return rows;
}

function migrationApplySnapshotV11_(child,snapshot,eventId,event,folderId) {
  migrationWriteTasksV11_(child,snapshot.tasks);
  migrationWriteParticipantsV11_(child,snapshot.participants);
  migrationWriteExpensesV11_(child,snapshot.expenses,event);
  writeEventMetaV5_(child,eventId,event,folderId);
}

function migrationWriteTasksV11_(child,tasks) {
  const sheet=child.getSheetByName(EVENT_SHEET.SHEETS.TASKS);if(!sheet)throw new Error('Foglio Attivita mancante nel modello V11.');
  const clearRows=Math.min(EVENT_MIGRATION_V11.TASK_MAX_ROWS-1,Math.max(0,sheet.getMaxRows()-1));
  if(clearRows>0){sheet.getRange(2,1,clearRows,6).clearContent();try{sheet.getRange(2,2,clearRows,1).clearNote();}catch(e){}}
  if(tasks.values.length){
    sheet.getRange(2,1,tasks.values.length,6).setValues(tasks.values);
    tasks.bNotes.forEach((note,i)=>{if(note)sheet.getRange(i+2,2).setNote(note);});
    for(let i=0;i<tasks.values.length;i++){
      const row=i+2,status=normalize_(sheet.getRange(row,5).getDisplayValue()),dep=Number(sheet.getRange(row,4).getValue()||0);
      if(dep>0&&status!=='FATTO')sheet.getRange(row,5).setFormula(fastTaskStatusFormulaV9_(row));
    }
  }
}

function migrationWriteParticipantsV11_(child,participants) {
  const sheet=child.getSheetByName('Partecipanti');if(!sheet)throw new Error('Foglio Partecipanti mancante nel modello V11.');
  sheet.getRange(EVENT_MIGRATION_V11.V11_CONVOCATI_START,1,EVENT_MIGRATION_V11.V11_CONVOCATI_END-EVENT_MIGRATION_V11.V11_CONVOCATI_START+1,EVENT_MIGRATION_V11.PARTICIPANT_COLS).clearContent();
  sheet.getRange(EVENT_MIGRATION_V11.V11_AGGREGATI_START,1,EVENT_MIGRATION_V11.V11_AGGREGATI_END-EVENT_MIGRATION_V11.V11_AGGREGATI_START+1,EVENT_MIGRATION_V11.PARTICIPANT_COLS).clearContent();
  if(participants.convocati.length)sheet.getRange(EVENT_MIGRATION_V11.V11_CONVOCATI_START,1,participants.convocati.length,EVENT_MIGRATION_V11.PARTICIPANT_COLS).setValues(participants.convocati);
  if(participants.aggregati.length)sheet.getRange(EVENT_MIGRATION_V11.V11_AGGREGATI_START,1,participants.aggregati.length,EVENT_MIGRATION_V11.PARTICIPANT_COLS).setValues(participants.aggregati);
}

function migrationWriteExpensesV11_(child,expenses,event) {
  const sheet=child.getSheetByName(EVENT_SHEET.SHEETS.EXPENSES);if(!sheet)throw new Error('Foglio Spese mancante nel modello V11.');
  const needed=expenses.rows.length,available=Math.max(0,sheet.getMaxRows()-EVENT_MIGRATION_V11.HISTORY_START+1);
  if(needed>available){
    const add=needed-available+20,oldMax=sheet.getMaxRows();sheet.insertRowsAfter(oldMax,add);
    const templateRow=Math.min(oldMax,Math.max(EVENT_MIGRATION_V11.HISTORY_START,oldMax));
    try{sheet.getRange(templateRow,1,1,18).copyTo(sheet.getRange(oldMax+1,1,add,18),SpreadsheetApp.CopyPasteType.PASTE_FORMAT,false);}catch(e){}
    try{sheet.getRange(templateRow,1,1,18).copyTo(sheet.getRange(oldMax+1,1,add,18),SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION,false);}catch(e){}
  }
  const count=Math.max(0,sheet.getMaxRows()-EVENT_MIGRATION_V11.HISTORY_START+1);if(count)sheet.getRange(EVENT_MIGRATION_V11.HISTORY_START,1,count,18).clearContent();
  if(needed){
    sheet.getRange(EVENT_MIGRATION_V11.HISTORY_START,1,needed,18).setValues(expenses.rows);
    sheet.getRange(EVENT_MIGRATION_V11.HISTORY_START,3,needed,1).setNumberFormat('€ #,##0.00');
    sheet.getRange(EVENT_MIGRATION_V11.HISTORY_START,6,needed,2).setNumberFormat('dd/MM/yyyy');
    sheet.getRange(EVENT_MIGRATION_V11.HISTORY_START,8,needed,2).setNumberFormat('@');
  }
  try{sheet.hideColumns(12,7);}catch(e){}
  sheet.getRange('B2').setValue(Number(expenses.initialBudget||0)).setNumberFormat('€ #,##0.00');
  const imp=String(expenses.imp||'').trim()||resolveEventCommitmentV7_(event)||String(event[APP.CALENDAR_HEADERS.COMMITMENT]||'').trim();
  if(imp)sheet.getRange('G9').setValue(imp).setNumberFormat('@');
  migrationFormatHistoryGroupsV11_(sheet,needed);
  migrationRefreshDashboardV11_(child,expenses.rows);
}

function migrationFormatHistoryGroupsV11_(sheet,rowCount) {
  if(!(rowCount>0))return;
  const vals=sheet.getRange(EVENT_MIGRATION_V11.HISTORY_START,1,rowCount,2).getValues();let start=0;
  while(start<vals.length){
    const quoteNo=Number(vals[start][0]||0);let end=start;while(end+1<vals.length&&Number(vals[end+1][0]||0)===quoteNo)end++;
    sheet.getRange(EVENT_MIGRATION_V11.HISTORY_START+start,1,end-start+1,11).setBorder(true,true,true,true,false,false,'#5f6368',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sheet.getRange(EVENT_MIGRATION_V11.HISTORY_START+start,1,1,11).setFontWeight('bold');
    if(end>start)sheet.getRange(EVENT_MIGRATION_V11.HISTORY_START+start+1,1,end-start,11).setFontWeight('normal');
    start=end+1;
  }
}

function migrationRefreshDashboardV11_(child,rows) {
  const expense=child.getSheetByName(EVENT_SHEET.SHEETS.EXPENSES),participants=child.getSheetByName('Partecipanti');if(!expense)return;
  const buckets={VIAGGI:0,VITTO:0,ALLOGGIO:0,NOLEGGI:0,RIMBORSI:0,ALTRO:0};let forecast=0,paid=0;
  rows.forEach(r=>{const movement=normalize_(r[1]),amount=Number(r[2]||0);if(movement==='PREVENTIVO'){forecast+=amount;const cat=migrationCategoryV11_(r[11],r[3]),bucket=['VIAGGI','VITTO','ALLOGGIO','NOLEGGI'].includes(cat)?cat:'ALTRO';buckets[bucket]+=amount;}else if(normalize_(r[4])==='PAGATO')paid+=amount;});
  if(participants)participants.getRange(3,1,15,19).getValues().forEach(r=>{const name=String(r[0]||'').trim(),surname=String(r[1]||'').trim();if(!name&&!surname)return;if(normalize_(r[7])==='TECNICO')return;const status=normalize_(r[8]),max=Number(r[9]||0),passed=Number(r[10]||0),current=status==='ASSENTE'?0:(passed>0?passed:max);forecast+=current;buckets.RIMBORSI+=current;if(passed>0)paid+=passed;});
  expense.getRange('D2').setValue(forecast).setNumberFormat('€ #,##0.00');expense.getRange('F2').setValue(paid).setNumberFormat('€ #,##0.00');expense.getRange('H2').setValue(Math.max(forecast-paid,0)).setNumberFormat('€ #,##0.00');expense.getRange('A5:F5').setValues([[buckets.VIAGGI,buckets.VITTO,buckets.ALLOGGIO,buckets.NOLEGGI,buckets.RIMBORSI,buckets.ALTRO]]).setNumberFormat('€ #,##0.00');
}

function migrationVerifyTargetV11_(child,snapshot) {
  if(!isEventSheetV11_(child))throw new Error('Verifica finale fallita: la nuova Scheda non risulta V11.');
  const taskSheet=child.getSheetByName(EVENT_SHEET.SHEETS.TASKS),participantSheet=child.getSheetByName('Partecipanti'),expenseSheet=child.getSheetByName(EVENT_SHEET.SHEETS.EXPENSES);
  const expectedTasks=snapshot.tasks.values.filter(r=>migrationRowHasDataV11_(r)).length;
  const actualTasks=taskSheet.getRange(2,1,Math.max(1,snapshot.tasks.values.length),6).getValues().filter(r=>migrationRowHasDataV11_(r)).length;
  if(actualTasks!==expectedTasks)throw new Error('Verifica Attivita fallita: attese '+expectedTasks+', trovate '+actualTasks+'.');
  const actualConv=participantSheet.getRange(3,1,15,2).getValues().filter(r=>String(r[0]||'').trim()||String(r[1]||'').trim()).length;
  const actualAgg=participantSheet.getRange(20,1,20,2).getValues().filter(r=>String(r[0]||'').trim()||String(r[1]||'').trim()).length;
  if(actualConv!==snapshot.participants.convocati.length||actualAgg!==snapshot.participants.aggregati.length)throw new Error('Verifica Partecipanti fallita.');
  const actualExpenses=snapshot.expenses.rows.length?expenseSheet.getRange(20,1,snapshot.expenses.rows.length,2).getValues().filter(r=>Number(r[0]||0)>0&&String(r[1]||'').trim()).length:0;
  if(actualExpenses!==snapshot.expenses.rows.length)throw new Error('Verifica Spese fallita: attesi '+snapshot.expenses.rows.length+', trovati '+actualExpenses+'.');
  const expectedBudget=snapshot.expenses.rows.filter(r=>normalize_(r[1])==='PREVENTIVO').reduce((s,r)=>s+Number(r[2]||0),0);
  const actualRows=snapshot.expenses.rows.length?expenseSheet.getRange(20,1,snapshot.expenses.rows.length,18).getValues():[];
  const actualBudget=actualRows.filter(r=>normalize_(r[1])==='PREVENTIVO').reduce((s,r)=>s+Number(r[2]||0),0);
  if(Math.abs(expectedBudget-actualBudget)>0.01)throw new Error('Verifica importi preventivi fallita.');
}

function migrationRecalculateQuoteStatusesArrayV11_(rows) {
  const quotes={};
  rows.forEach(r=>{if(normalize_(r[1])==='PREVENTIVO'){const n=Number(r[0]||0);if(n>0)quotes[n]={row:r,budget:Number(r[2]||0),paid:0};}});
  rows.forEach(r=>{const n=Number(r[0]||0);if(quotes[n]&&normalize_(r[1])!=='PREVENTIVO'&&normalize_(r[4])==='PAGATO')quotes[n].paid+=Number(r[2]||0);});
  Object.keys(quotes).forEach(k=>{const q=quotes[k];q.row[4]=q.paid<=0.005?'DA PAGARE':(q.paid+0.005>=q.budget?'PAGATO':'PARZIALMENTE PAGATO');});
}

function migrationCategoryV11_(value,description) {
  let c=normalize_(value),d=normalize_(description);
  if(c==='VITTO / ALLOGGIO')c=/(PASTI|VITTO|RISTORANTE|PRANZO|CENA)/.test(d)?'VITTO':'ALLOGGIO';
  if(c==='VIAGGIO'||c==='TRASPORTO')c='VIAGGI';
  return EVENT_MIGRATION_V11.CATEGORIES.includes(c)?c:'ALTRO';
}

function migrationPaymentTypeV11_(value) {
  const t=normalize_(value);
  if(t==='PREVENTIVO')return'PREVENTIVO';
  if(t==='FATTURA'||t==='PAGAMENTO FATTURA')return'SALDO FATTURA';
  if(t==='CARTA'||t==='CC')return'CARTA DI CREDITO';
  return EVENT_MIGRATION_V11.PAYMENT_TYPES.includes(t)?t:'';
}

function migrationLegacyNoteV11_(value,supplier,rawCategory) {
  let note=(value===false||value===null||value===undefined)?'':String(value).trim();
  const parts=[];
  if(note)parts.push(note);
  if(supplier)parts.push('Fornitore: '+supplier);
  const raw=normalize_(rawCategory);if(raw&&migrationCategoryV11_(raw,'')==='ALTRO'&&raw!=='ALTRO')parts.push('Tipologia precedente: '+String(rawCategory).trim());
  return parts.join(' | ');
}

function migrationCleanSupplierV11_(value){const s=String(value||'').trim();return s==='-'?'':s;}
function migrationRowHasDataV11_(row){return (row||[]).some(v=>(v instanceof Date||v===true||v===false)?true:String(v===null||v===undefined?'':v).trim()!=='');}
function migrationPadRowV11_(row,width){const out=(row||[]).slice(0,width);while(out.length<width)out.push('');return out;}

function migrationFindLabelRowV11_(sheet,label,col,maxRows) {
  const n=Math.max(1,Math.min(maxRows||sheet.getMaxRows(),sheet.getMaxRows())),vals=sheet.getRange(1,col,n,1).getDisplayValues();
  for(let i=0;i<vals.length;i++)if(normalize_(vals[i][0])===normalize_(label))return i+1;
  return 0;
}

function migrationCellTextOrLinkV11_(sheet,row,col) {
  const cell=sheet.getRange(row,col),display=String(cell.getDisplayValue()||'').trim(),urls=[];
  try{
    const rich=cell.getRichTextValue();
    if(rich){
      const direct=rich.getLinkUrl();if(direct)urls.push(direct);
      const runs=rich.getRuns();if(runs)runs.forEach(run=>{const u=run.getLinkUrl();if(u&&urls.indexOf(u)<0)urls.push(u);});
    }
  }catch(e){}
  if(urls.length)return urls.join('\n');
  const value=cell.getValue();return value===false?'':String(value||display||'').trim();
}
