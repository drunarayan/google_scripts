function copyRowEveryMinute() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sname = 'DAILIES';
  const tname = 'TODAY';
  var rowToCopy = 1;
  const sourceSheet = ss.getSheetByName(sname); // Change to your source tab name
  const targetSheet = ss.getSheetByName(tname); // Change to your target tab name
  if (!sourceSheet) {
    Logger.log(`Source Sheet "${sname}" not found.`);
    return;
  } else dailies();
  if (!targetSheet) {
    Logger.log(`Target Sheet "${tname}" not found.`);
    return;
  }
  
  var lastRow = targetSheet.getLastRow();
  if (lastRow > 1) {
    targetSheet.getRange(targetSheet.getLastRow(),1).setValue(timeNow());
    rowToCopy = 2;
    
    }

  sourceSheet.getRange(sourceSheet.getLastRow(),1).setValue(timeNow());
  
  // Get data from a specific row (e.g., Row 2)

  const range = sourceSheet.getRange(rowToCopy, 1, 1, sourceSheet.getLastColumn());
  const data = range.getValues();
  //data(1,1,1,1).setValue(timeNow());
  // Append data to the bottom of the target sheet
  targetSheet.appendRow(data[0]);
}