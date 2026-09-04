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
  sheet.getRange('E2:E' + TASKS_V3.ENTRY_END_ROW)
    .setNumberFormat('@')
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(TASKS_V3.STATUS,true)
        .setAllowInvalid(true)
        .build()
    );
}
