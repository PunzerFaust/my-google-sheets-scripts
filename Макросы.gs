function mac() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('22:22').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Архив_ЛИДЫ'), true);
  spreadsheet.getRange('A6').activate();
  spreadsheet.getRange('\'ЛИДЫ\'!22:22').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getRange('D6').activate();
};