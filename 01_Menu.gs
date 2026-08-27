function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Calendario v2')
    .addItem('Apri pannello evento', 'openEventSidebar')
    .addSeparator()
    .addItem('Genera/aggiorna checklist', 'generateChecklistForSelectedEvent')
    .addItem('Registra rimborso', 'openRefundDialog')
    .addSeparator()
    .addItem('Aggiorna riepiloghi', 'refreshSelectedEventSummary')
    .addToUi();
}
