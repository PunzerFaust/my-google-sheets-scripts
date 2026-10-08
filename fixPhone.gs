function fixPhoneNumbers() {
  const fixingSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = fixingSheet.getSheetByName("ЛИДЫ");
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  const range = sheet.getRange(2, 3, lastRow - 1, 1);
  
  const formulas = range.getFormulas(); // берём формулы
  const displayValues = range.getDisplayValues(); // берём то, что написано
  const newValues = [];
  
  for (let i = 0; i < formulas.length; i++) {
    let rawText = formulas[i][0] || displayValues[i][0]; // проверяем на на формулу
    
    if (rawText.startsWith("=") || rawText.startsWith("+") || rawText.startsWith("'")) {
      if (rawText.startsWith("=")) rawText = rawText.substring(1); // всё кроме 0 элемента
      if (rawText.startsWith("'")) rawText = rawText.substring(1); 
    }
    
    newValues.push([" "+ rawText]);
  }
  range.clearContent()
  range.setNumberFormat("@");
  range.setValues(newValues);
}
