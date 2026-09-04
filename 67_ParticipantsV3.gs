const PARTICIPANTS_V3 = Object.freeze({
  SHEET: 'Partecipanti',
  CONVOCATI_TITLE_ROW: 1,
  CONVOCATI_HEADER_ROW: 2,
  CONVOCATI_START_ROW: 3,
  CONVOCATI_END_ROW: 42,
  AGGREGATI_TITLE_ROW: 44,
  AGGREGATI_HEADER_ROW: 45,
  AGGREGATI_START_ROW: 46,
  AGGREGATI_END_ROW: 65,
  VISIBLE_COLS: 10,
  TECH_START_COL: 11,
  TECH_COLS: 6,
  CONVOCATI_HEADERS: Object.freeze([
    'NOME','COGNOME','CIRCOLO','EMAIL','RUOLO','STATO CONVOCAZIONE',
    'MASSIMALE RIMBORSO','IMPORTO PASSATO','DATA','NOTE'
  ]),
  AGGREGATI_HEADERS: Object.freeze([
    'NOME','COGNOME','CIRCOLO','EMAIL','RUOLO','STATO AUTORIZZAZIONE','NOTE'
  ]),
  TECH_HEADERS: Object.freeze([
    'ID PARTECIPAZIONE','ID PERSONA','ID EVENTO','PROVENIENZA','FILE ORIGINE','DATA INSERIMENTO'
  ]),
  ROLES: Object.freeze(['ATLETA','TECNICO','ALTRO']),
  CONVOCATION_STATES: Object.freeze(['DA FARE','MANDATA CONVOCAZIONE','CONFERMATO','ASSENTE']),
  AUTH_STATES: Object.freeze(['DA AUTORIZZARE','AUTORIZZATO','NON AUTORIZZATO'])
});

function ensureParticipantV3Structure_(child,event) {
  let sheet = child.getSheetByName(PARTICIPANTS_V3.SHEET);
  if (!sheet) sheet = child.insertSheet(PARTICIPANTS_V3.SHEET);
  if (sheet.getMaxColumns() < 16) sheet.insertColumnsAfter(sheet.getMaxColumns(),16-sheet.getMaxColumns());
  if (sheet.getMaxRows() < PARTICIPANTS_V3.AGGREGATI_END_ROW) {
    sheet.insertRowsAfter(sheet.getMaxRows(),PARTICIPANTS_V3.AGGREGATI_END_ROW-sheet.getMaxRows());
  }

  // Rimuove le sole unioni delle due righe titolo prima di ricrearle.
  sheet.getRange('A1:J1').breakApart();
  sheet.getRange('A44:J44').breakApart();
  sheet.getRange('A1:J1').merge();
  sheet.getRange('A44:J44').merge();

  const commitment = String(event && event[APP.CALENDAR_HEADERS.COMMITMENT] || '').trim();
  const title = commitment ? 'CONVOCATI  |  IMPEGNO: ' + commitment : 'CONVOCATI';
  sheet.getRange('A1').setValue(title);
  sheet.getRange('A2:J2').setValues([PARTICIPANTS_V3.CONVOCATI_HEADERS]);
  sheet.getRange('A44').setValue('AGGREGATI');
  sheet.getRange('A45:G45').setValues([PARTICIPANTS_V3.AGGREGATI_HEADERS]);
  sheet.getRange('H45:J45').clearContent();
  sheet.getRange('K1:P1').setValues([PARTICIPANTS_V3.TECH_HEADERS]);

  sheet.setFrozenRows(2);
  sheet.showColumns(1,10);
  sheet.hideColumns(11,6);

  sheet.getRange('A1:J1')
    .setFontWeight('bold').setFontSize(13).setBackground('#d9eaf7')
    .setVerticalAlignment('middle').setHorizontalAlignment('left');
  sheet.setRowHeight(1,30);
  sheet.getRange('A44:J44')
    .setFontWeight('bold').setBackground('#d9eaf7').setVerticalAlignment('middle');
  sheet.getRange('A2:J2').setFontWeight('bold').setBackground('#eeeeee').setWrap(true).setVerticalAlignment('middle');
  sheet.getRange('A45:G45').setFontWeight('bold').setBackground('#eeeeee').setWrap(true).setVerticalAlignment('middle');

  [140,140,220,230,105,175,130,125,105,280].forEach((w,i)=>sheet.setColumnWidth(i+1,w));
  sheet.getRange('A3:J42').setVerticalAlignment('top').setWrap(true);
  sheet.getRange('A46:G65').setVerticalAlignment('top').setWrap(true);
  sheet.getRange('G3:H42').setNumberFormat('€ #,##0.00');
  sheet.getRange('I3:I42').setNumberFormat('dd/MM/yyyy');

  sheet.getRange('E3:E42').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(PARTICIPANTS_V3.ROLES,true).setAllowInvalid(false).build()
  );
  sheet.getRange('F3:F42').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(PARTICIPANTS_V3.CONVOCATION_STATES,true).setAllowInvalid(false).build()
  );
  sheet.getRange('E46:E65').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(PARTICIPANTS_V3.ROLES,true).setAllowInvalid(false).build()
  );
  sheet.getRange('F46:F65').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(PARTICIPANTS_V3.AUTH_STATES,true).setAllowInvalid(false).build()
  );

  sheet.getRange('F2').setNote('DA FARE = convocazione non ancora inviata. Poi: MANDATA CONVOCAZIONE, CONFERMATO oppure ASSENTE.');
  sheet.getRange('G2').setNote('Massimale previsto per il rimborso. Viene precompilato ma resta modificabile.');
  sheet.getRange('H2').setNote('Importo effettivamente passato a rimborso. La data viene registrata al successivo aggiornamento.');
  sheet.getRange('F45').setNote('Stati: DA AUTORIZZARE, AUTORIZZATO, NON AUTORIZZATO.');

  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$F3="CONFERMATO"')
      .setBackground('#d9ead3').setRanges([sheet.getRange('A3:J42')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$F3="ASSENTE"')
      .setBackground('#f4cccc').setRanges([sheet.getRange('A3:J42')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$F3="MANDATA CONVOCAZIONE"')
      .setBackground('#fff2cc').setRanges([sheet.getRange('A3:J42')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$F3="DA FARE"')
      .setBackground('#d9eaf7').setRanges([sheet.getRange('A3:J42')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$F46="AUTORIZZATO"')
      .setBackground('#d9ead3').setRanges([sheet.getRange('A46:G65')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$F46="NON AUTORIZZATO"')
      .setBackground('#f4cccc').setRanges([sheet.getRange('A46:G65')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$F46="DA AUTORIZZARE"')
      .setBackground('#fff2cc').setRanges([sheet.getRange('A46:G65')]).build()
  ]);

  const legacy = child.getSheetByName(PARTICIPANTS_V2.LEGACY_SHEET);
  if (legacy) legacy.hideSheet();
  return sheet;
}

function syncParticipantsV3ToBackend_(eventId,event,child) {
  const sheet = ensureParticipantV3Structure_(child,event);
  const backend = sh_(APP.SHEETS.PARTICIPANTS);
  ensureParticipantsBackendHeadersV2_();

  const oldRows = backend.getDataRange().getValues();
  const oldById = {};
  const oldByKey = {};
  for (let i=1;i<oldRows.length;i++) {
    const r = oldRows[i];
    if (String(r[1]||'') !== String(eventId)) continue;
    if (r[0]) oldById[String(r[0])] = r;
    const key = participantKeyV2_(r[3],r[4],r[6]);
    if (key) oldByKey[key] = r;
  }

  const now = new Date();
  const rows = [];
  const convTech = [];
  const aggTech = [];
  const convValues = sheet.getRange(PARTICIPANTS_V3.CONVOCATI_START_ROW,1,
    PARTICIPANTS_V3.CONVOCATI_END_ROW-PARTICIPANTS_V3.CONVOCATI_START_ROW+1,16).getValues();
  const aggValues = sheet.getRange(PARTICIPANTS_V3.AGGREGATI_START_ROW,1,
    PARTICIPANTS_V3.AGGREGATI_END_ROW-PARTICIPANTS_V3.AGGREGATI_START_ROW+1,16).getValues();

  convValues.forEach((v,index)=>{
    const rowNumber = PARTICIPANTS_V3.CONVOCATI_START_ROW + index;
    const parsed = parseParticipantRowV3_(v,'CONVOCATO',eventId,event,oldById,oldByKey,now,rowNumber);
    if (!parsed) { convTech.push(['','','','','','']); return; }
    rows.push(parsed.backendRow);
    convTech.push(parsed.techRow);
  });
  aggValues.forEach((v,index)=>{
    const rowNumber = PARTICIPANTS_V3.AGGREGATI_START_ROW + index;
    const parsed = parseParticipantRowV3_(v,'AGGREGATO',eventId,event,oldById,oldByKey,now,rowNumber);
    if (!parsed) { aggTech.push(['','','','','','']); return; }
    rows.push(parsed.backendRow);
    aggTech.push(parsed.techRow);
  });

  replaceCentralRowsForEvent_(backend,eventId,2,rows,17);
  sheet.getRange(PARTICIPANTS_V3.CONVOCATI_START_ROW,11,convTech.length,6).setValues(convTech);
  sheet.getRange(PARTICIPANTS_V3.AGGREGATI_START_ROW,11,aggTech.length,6).setValues(aggTech);
  writeLegacyParticipantsFromBackendV2_(eventId,child);
  return rows.length;
}

function parseParticipantRowV3_(v,type,eventId,event,oldById,oldByKey,now,rowNumber) {
  const isConvocato = type === 'CONVOCATO';
  const name = String(v[0]||'').trim();
  const surname = String(v[1]||'').trim();
  if (!name && !surname) return null;

  const club = String(v[2]||'').trim();
  const email = String(v[3]||'').trim();
  const role = normalize_(v[4]||'ATLETA') || 'ATLETA';
  const key = participantKeyV2_(name,surname,role);
  let id = String(v[10]||'').trim();
  let old = id && oldById[id] ? oldById[id] : (oldByKey[key] || null);
  if (!id && old && old[0]) id = String(old[0]);
  if (!id) id = 'PAR-' + Utilities.getUuid();

  let personId = String(v[11]||'').trim() || (old ? String(old[2]||'') : '');
  let finalClub = club;
  let finalEmail = email;
  if (!personId) {
    const known = findKnownPerson_(name,surname,'');
    if (known) {
      personId = known.id || '';
      if (!finalClub) finalClub = known.club || '';
      if (!finalEmail) finalEmail = known.email || '';
    }
  }

  let status = normalize_(v[5]);
  if (isConvocato && !PARTICIPANTS_V3.CONVOCATION_STATES.includes(status)) status = 'DA FARE';
  if (!isConvocato && !PARTICIPANTS_V3.AUTH_STATES.includes(status)) status = 'DA AUTORIZZARE';

  let maxRefund = '';
  let passed = '';
  let passedDate = '';
  let notes = '';
  if (isConvocato) {
    const enteredMax = v[6];
    maxRefund = enteredMax === '' || enteredMax === null
      ? defaultRefundForRoleV2_(eventId,event,role)
      : Number(enteredMax||0);
    const enteredPassed = v[7];
    if (enteredPassed !== '' && enteredPassed !== null && Number(enteredPassed||0) >= 0) {
      passed = Number(enteredPassed||0);
      if (passed > 0) {
        passedDate = v[8] instanceof Date ? v[8] : (old && old[16] instanceof Date ? old[16] : now);
      }
    }
    notes = String(v[9]||'').trim();
  } else {
    notes = String(v[6]||'').trim();
  }

  const provenance = String(v[13]||'').trim() || (old ? String(old[9]||'') : '') || 'SCHEDA EVENTO';
  const sourceFile = String(v[14]||'').trim() || (old ? String(old[10]||'') : '');
  const createdAt = v[15] instanceof Date ? v[15] : (old && old[12] instanceof Date ? old[12] : now);

  const backendRow = [
    id,eventId,personId,name,surname,finalClub,role,finalEmail,maxRefund,
    provenance,sourceFile,notes,createdAt,type,status,passed,passedDate
  ];
  return {
    backendRow:backendRow,
    techRow:[id,personId,eventId,provenance,sourceFile,createdAt],
    rowNumber:rowNumber
  };
}

function refreshParticipantsV3FromBackend_(eventId,event,child) {
  const sheet = ensureParticipantV3Structure_(child,event);
  ensureParticipantsBackendHeadersV2_();
  const rows = sh_(APP.SHEETS.PARTICIPANTS).getDataRange().getValues().slice(1)
    .filter(r=>String(r[1]||'')===String(eventId));
  const convocati = rows.filter(r=>normalize_(r[13]||'CONVOCATO')!=='AGGREGATO');
  const aggregati = rows.filter(r=>normalize_(r[13])==='AGGREGATO');

  const convSlots = PARTICIPANTS_V3.CONVOCATI_END_ROW-PARTICIPANTS_V3.CONVOCATI_START_ROW+1;
  const aggSlots = PARTICIPANTS_V3.AGGREGATI_END_ROW-PARTICIPANTS_V3.AGGREGATI_START_ROW+1;
  if (convocati.length > convSlots) throw new Error('Troppi convocati per la scheda evento: ' + convocati.length + '.');
  if (aggregati.length > aggSlots) throw new Error('Troppi aggregati per la scheda evento: ' + aggregati.length + '.');

  sheet.getRange(PARTICIPANTS_V3.CONVOCATI_START_ROW,1,convSlots,16).clearContent();
  sheet.getRange(PARTICIPANTS_V3.AGGREGATI_START_ROW,1,aggSlots,16).clearContent();

  if (convocati.length) {
    const out = convocati.map(r=>{
      const role = normalize_(r[6]||'ATLETA') || 'ATLETA';
      const maxRefund = r[8] === '' || r[8] === null ? defaultRefundForRoleV2_(eventId,event,role) : r[8];
      const status = PARTICIPANTS_V3.CONVOCATION_STATES.includes(normalize_(r[14])) ? normalize_(r[14]) : 'DA FARE';
      return [
        r[3]||'',r[4]||'',r[5]||'',r[7]||'',role,status,maxRefund,r[15]||'',r[16]||'',r[11]||'',
        r[0]||'',r[2]||'',eventId,r[9]||'',r[10]||'',r[12]||''
      ];
    });
    sheet.getRange(PARTICIPANTS_V3.CONVOCATI_START_ROW,1,out.length,16).setValues(out);
  }

  if (aggregati.length) {
    const out = aggregati.map(r=>[
      r[3]||'',r[4]||'',r[5]||'',r[7]||'',r[6]||'ATLETA',r[14]||'DA AUTORIZZARE',r[11]||'','','','',
      r[0]||'',r[2]||'',eventId,r[9]||'',r[10]||'',r[12]||''
    ]);
    sheet.getRange(PARTICIPANTS_V3.AGGREGATI_START_ROW,1,out.length,16).setValues(out);
  }

  sheet.getRange('G3:H42').setNumberFormat('€ #,##0.00');
  sheet.getRange('I3:I42').setNumberFormat('dd/MM/yyyy');
  return rows.length;
}
