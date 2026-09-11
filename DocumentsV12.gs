const EVENT_DOCUMENTS_V12 = Object.freeze({
  SHEET:'DOCUMENTI',
  TEMPLATE_CELL:'B3',
  INSTRUCTIONS_CELL:'B4',
  LIST_START_ROW:9,
  LIST_MAX_ROWS:190
});

function eventRefreshDocumentListV12() {
  return eventRefreshDocumentListV12_(true);
}

function eventRefreshDocumentListV12_(showToast) {
  const ss=SpreadsheetApp.getActive(),meta=eventMeta_(),sheet=eventEnsureDocumentsSheetV12_();
  const folderId=String(meta.EVENT_FOLDER_ID||'').trim();
  if(!folderId)return 0;
  const folder=DriveApp.getFolderById(folderId),files=folder.getFiles(),rows=[];
  while(files.hasNext()){
    const file=files.next(),name=String(file.getName()||'').trim();
    if(file.getId()===ss.getId())continue;
    if(/^(BACKUP PRE-V11|LEGACY PRE-V11|__MIGRAZIONE)/i.test(name))continue;
    rows.push({name:name,url:file.getUrl()});
  }
  rows.sort((a,b)=>a.name.localeCompare(b.name,'it',{sensitivity:'base'}));
  const clearCount=Math.min(EVENT_DOCUMENTS_V12.LIST_MAX_ROWS,Math.max(sheet.getMaxRows()-EVENT_DOCUMENTS_V12.LIST_START_ROW+1,1));
  sheet.getRange(EVENT_DOCUMENTS_V12.LIST_START_ROW,1,clearCount,2).clearContent();
  if(rows.length){
    const count=Math.min(rows.length,clearCount);
    sheet.getRange(EVENT_DOCUMENTS_V12.LIST_START_ROW,1,count,1).setValues(rows.slice(0,count).map(x=>[x.name]));
    const rich=rows.slice(0,count).map(x=>[SpreadsheetApp.newRichTextValue().setText('↗ APRI').setLinkUrl(x.url).build()]);
    sheet.getRange(EVENT_DOCUMENTS_V12.LIST_START_ROW,2,count,1).setRichTextValues(rich);
  }
  if(showToast)ss.toast(rows.length+' file visualizzati.','Documenti',4);
  return rows.length;
}

function eventCreateDocumentFromDocumentsSheetV12() {
  const ss=SpreadsheetApp.getActive(),sheet=eventEnsureDocumentsSheetV12_(),label=String(sheet.getRange(EVENT_DOCUMENTS_V12.TEMPLATE_CELL).getDisplayValue()||'').trim();
  const instructions=String(sheet.getRange(EVENT_DOCUMENTS_V12.INSTRUCTIONS_CELL).getDisplayValue()||'').trim();
  const spec=EVENT_APP.TEMPLATE_SPECS.find(x=>eventNormalizeTextV7_(x.label)===eventNormalizeTextV7_(label));
  if(!spec)throw new Error('Seleziona un modello valido nel foglio DOCUMENTI.');
  const request=eventBuildDocumentRequestV12_(spec,instructions);
  const file=eventCreateDocumentFromRequestV12_(request);
  eventRefreshDocumentListV12_(false);
  ss.toast('Creato: '+file.getName(),'Documento creato',6);
  return file.getId();
}

function eventBuildDocumentRequestV12_(spec,instructions) {
  const meta=eventMeta_(),replacements=eventDocumentReplacements_(meta);
  return {
    version:1,
    mode:'TEMPLATE',
    aiEnabled:false,
    eventId:String(meta.EVENT_ID||''),
    folderId:String(meta.EVENT_FOLDER_ID||''),
    templateKey:spec.key,
    templateId:spec.id,
    templateLabel:spec.label,
    instructions:String(instructions||'').trim(),
    replacements:replacements,
    outputName:eventDocumentName_(spec.key,replacements)
  };
}

function eventCreateDocumentFromRequestV12_(request) {
  if(!request||!request.folderId||!request.templateId)throw new Error('Richiesta documento incompleta.');
  const folder=DriveApp.getFolderById(request.folderId),meta=eventMeta_();
  const copy=DriveApp.getFileById(request.templateId).makeCopy(request.outputName,folder);
  // Extension point AI: in futuro request.instructions potra essere usato per generare
  // una versione personalizzata del contenuto mantenendo il modello come base.
  eventFillDocument_(copy.getId(),request.replacements,request.templateKey);
  const spec=EVENT_APP.TEMPLATE_SPECS.find(x=>x.key===request.templateKey)||{key:request.templateKey,id:request.templateId};
  eventRegisterDocument_(meta,spec,copy);
  return copy;
}

function eventEnsureDocumentsSheetV12_() {
  const ss=SpreadsheetApp.getActive();let sheet=ss.getSheetByName(EVENT_DOCUMENTS_V12.SHEET);
  if(sheet)return sheet;
  sheet=ss.insertSheet(EVENT_DOCUMENTS_V12.SHEET);
  if(sheet.getMaxColumns()<4)sheet.insertColumnsAfter(sheet.getMaxColumns(),4-sheet.getMaxColumns());
  if(sheet.getMaxRows()<200)sheet.insertRowsAfter(sheet.getMaxRows(),200-sheet.getMaxRows());
  sheet.getRange('A1:D1').merge().setValue('DOCUMENTI EVENTO').setFontWeight('bold').setFontSize(13).setBackground('#d9eaf7');
  sheet.getRange('A3').setValue('MODELLO').setFontWeight('bold').setBackground('#eeeeee');
  sheet.getRange('B3').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(EVENT_APP.TEMPLATE_SPECS.map(x=>x.label),true).setAllowInvalid(false).build());
  sheet.getRange('A4').setValue('DESCRIZIONE / ISTRUZIONI').setFontWeight('bold').setBackground('#eeeeee');
  sheet.getRange('B4:D4').merge();
  sheet.getRange('B4').setNote('Campo predisposto per la futura generazione AI da modello. Per ora non modifica il contenuto standard.');
  sheet.getRange('A5').setValue('CREA DAL MENU').setFontWeight('bold').setBackground('#e8f0fe');
  sheet.getRange('B5:D5').merge().setValue('Scheda evento → Crea documento dal foglio DOCUMENTI');
  sheet.getRange('A7:D7').merge().setValue('FILE NELLA CARTELLA EVENTO').setFontWeight('bold').setBackground('#eeeeee');
  sheet.getRange('A8:B8').setValues([['NOME','LINK']]).setFontWeight('bold').setBackground('#eeeeee');
  sheet.setFrozenRows(8);sheet.setColumnWidth(1,310);sheet.setColumnWidth(2,170);sheet.setColumnWidth(3,170);sheet.setColumnWidth(4,95);
  sheet.getRange('A1:D200').setVerticalAlignment('middle').setWrap(true);
  return sheet;
}
