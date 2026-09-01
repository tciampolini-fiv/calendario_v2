function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Calendario v2')
    .addItem('Crea / aggiorna scheda evento', 'prepareEventSheetForSelectedEvent')
    .addItem('Aggiorna calendario dalla scheda', 'syncSelectedEventSheetToCalendar')
    .addSeparator()
    .addItem('Crea cartella di lavoro', 'createWorkFolderForSelectedEvent')
    .addItem('Genera documenti evento', 'generateDocumentsForSelectedEvent')
    .addSeparator()
    .addItem('Genera checklist evento', 'generateChecklistForSelectedEvent')
    .addItem('Ripulisci checklist eventi attivi/futuri', 'rebuildCurrentAndFutureStandardChecklists')
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== APP.SHEETS.CALENDAR || e.range.getRow() < 2) return;
  const map = headerMap_(sheet);
  const editedCol = e.range.getColumn();
  const row = e.range.getRow();

  const autoGrowCols = [
    map[APP.CALENDAR_HEADERS.EVENT],
    map[APP.CALENDAR_HEADERS.NOTES]
  ].filter(Boolean);
  if (autoGrowCols.includes(editedCol)) {
    e.range.setWrap(true).setVerticalAlignment('top');
    sheet.autoResizeRows(row, e.range.getNumRows());
    return;
  }

  if (editedCol !== map[APP.CALENDAR_HEADERS.TYPE] && editedCol !== map[APP.CALENDAR_HEADERS.CLASS]) return;

  const commitmentCell = sheet.getRange(row, map[APP.CALENDAR_HEADERS.COMMITMENT]);
  if (String(commitmentCell.getDisplayValue() || '').trim()) return;

  const type = normalize_(sheet.getRange(row, map[APP.CALENDAR_HEADERS.TYPE]).getDisplayValue());
  const cls = normalize_(sheet.getRange(row, map[APP.CALENDAR_HEADERS.CLASS]).getDisplayValue());
  if (!type || !cls) return;

  const code = resolveCommitmentCode_(type, cls);
  if (code) commitmentCell.setNumberFormat('@').setValue(String(code));
}

function resolveCommitmentCode_(type, cls) {
  const rows = sh_(APP.SHEETS.COMMITMENT_CONFIG).getDataRange().getDisplayValues();
  const typeAliases = [normalize_(type)];
  if (typeAliases[0].indexOf('REGATA INT.') === 0) typeAliases.push('REGATA');
  if (typeAliases[0] === 'STAGE') typeAliases.push('ALLENAMENTO');

  for (let i = 1; i < rows.length; i++) {
    const rType = normalize_(rows[i][0]);
    const rClass = normalize_(rows[i][1]);
    const code = String(rows[i][2] || '').trim();
    if (!code) continue;
    if (typeAliases.includes(rType) && (rClass === cls || rClass === '*' || !rClass)) return code;
  }
  return '';
}
