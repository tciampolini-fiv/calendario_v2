function ensureEventSheetEditTriggerV5_(child) {
  const sourceId = child.getId();
  const exists = ScriptApp.getProjectTriggers().some(t=>{
    if (t.getHandlerFunction() !== 'handleEventSheetEditV5') return false;
    try { return t.getTriggerSourceId() === sourceId; } catch (e) { return false; }
  });
  if (!exists) ScriptApp.newTrigger('handleEventSheetEditV5').forSpreadsheet(sourceId).onEdit().create();
}

function eventFromChildMetaV5_(child) {
  const meta = child.getSheetByName(EVENT_SHEET.SHEETS.META);
  if (!meta) return null;
  const eventId = String(readMetaValue_(meta,'EVENT_ID')||'').trim();
  if (!eventId) return null;
  const found = findCalendarEventById_(eventId);
  return found ? {eventId:eventId,event:found.event} : null;
}

function handleEventSheetEditV5(e) {
  if (!e || !e.range || !e.source) return;
  const sheetName = e.range.getSheet().getName();
  if (sheetName !== EXPENSES_V4.SHEET && sheetName !== PARTICIPANTS_V3.SHEET) return;

  const resolved = eventFromChildMetaV5_(e.source);
  if (!resolved) return;

  if (sheetName === EXPENSES_V4.SHEET) {
    try {
      handleExpenseSheetEditV4_(e,resolved.event);
    } catch (err) {
      try {
        e.range.getSheet().getRange('J7').setValue(false);
        e.source.toast(err.message || String(err),'Spese',6);
      } catch (_) {}
    }
    return;
  }

  // Le modifiche ai rimborsi e allo stato presenza aggiornano subito il cruscotto locale,
  // senza scrivere nel calendario centrale: la sincronizzazione centrale resta manuale.
  const row = e.range.getRow();
  const firstCol = e.range.getColumn();
  const lastCol = firstCol + e.range.getNumColumns() - 1;
  const inConvocati = row >= PARTICIPANTS_V3.CONVOCATI_START_ROW && row <= PARTICIPANTS_V3.CONVOCATI_END_ROW;
  const affectsRefund = firstCol <= 8 && lastCol >= 6;
  if (inConvocati && affectsRefund) {
    const expenseSheet = e.source.getSheetByName(EXPENSES_V4.SHEET);
    if (expenseSheet && normalize_(expenseSheet.getRange('A1').getDisplayValue()) === 'RIEPILOGO SPESE EVENTO') {
      refreshExpenseDashboardV4_(e.source,resolved.event);
    }
  }
}
