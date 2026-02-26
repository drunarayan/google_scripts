/*** @OnlyCurrentDoc */

function deleteColumnsByHeader() {
  const sheetName = 'PREPRC'; // Replace with your sheet name
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`Sheet "${sheetName}" not found.`);
    return;
  }

  // Get all values from the first row (header row)
  const headerRow = sheet.getRange("1:1").getValues()[0];
  const columnsToDelete = [];
  const headersToMatch = ["AS OF", "CUSIP", "FACTOR", "CHANGE IN PRICE", "CHANGE IN VALUE", "PERCENT CHANGE", "PERCENT OF PORTFOLIO"]; // Add the headers you want to delete

  // Loop through the headers to find matches and store their 1-based index
  // Start the loop from the end to the beginning (right to left) to handle deletion correctly
  for (let i = headerRow.length - 1; i >= 0; i--) {
    if (headersToMatch.includes(headerRow[i])) {
      // Columns in Apps Script are 1-based, so add 1 to the 0-based index
      columnsToDelete.push(i + 1);
    }
  }

  // Delete the columns
  // Deleting from right to left ensures that the column indices remain valid as columns are removed
  columnsToDelete.forEach(colIndex => {
    sheet.deleteColumn(colIndex);
  });
}


function copyValuesBasedOnCondition() {
  const sheetName = 'PREPRC'; // Replace with your sheet name
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`Sheet "${sheetName}" not found.`);
    return;
  }
  const SOURCE_HEADER = 'VALUE'; // Update with the header to copy FROM
  const TARGET_HEADER1 = 'TOTAL COST'; // Update with the header1 to copy TO
  const TARGET_HEADER2 = 'QUANTITY'; // Update with the header2 to copy TO
  const TARGET_HEADER3 = 'PRICE'; // Update with the header2 to copy TO
  const TARGET_HEADER4 = 'SYMBOL'; // Update with the header2 to copy TO
  const CONDITION_HEADER = 'CUSIP';   // Update with the header for the condition
  const CONDITION_VALUE = 'N/A';  // Update with the value to check for

  // Get all data including the header row
  const range = sheet.getDataRange();
  const values = range.getValues();
  const headers = values[0]; // First row is the header row

  // Find the column indices for the headers
  const sourceColIndex = headers.indexOf(SOURCE_HEADER);
  const targetColIndex1 = headers.indexOf(TARGET_HEADER1);
  const targetColIndex2 = headers.indexOf(TARGET_HEADER2);
  const targetColIndex3 = headers.indexOf(TARGET_HEADER3);
  const targetColIndex4 = headers.indexOf(TARGET_HEADER4);
  const conditionColIndex = headers.indexOf(CONDITION_HEADER);

  // Check if all headers were found
  if (sourceColIndex === -1 || targetColIndex1 === -1 || targetColIndex2 === -1 || targetColIndex3 === -1 || targetColIndex4 === -1 || conditionColIndex === -1) {
    Logger.log('One or more headers not found. Continuing ...');
  }

  // Loop through rows, starting from the second row (index 1) to skip the header
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const conditionValue = row[conditionColIndex];

    // Check if the condition value matches
    if (conditionValue === CONDITION_VALUE) {
      const valueToCopy = row[sourceColIndex];
      // Set the value in the target column for the current row
      // getRange uses 1-based indexing for row and column numbers
      if (targetColIndex1 >= 0) sheet.getRange(i + 1, targetColIndex1 + 1).setValue(valueToCopy);
      if (targetColIndex2 >= 0) sheet.getRange(i + 1, targetColIndex2 + 1).setValue(valueToCopy);
      if (targetColIndex3 >= 0) sheet.getRange(i + 1, targetColIndex3 + 1).setValue(1.0);
      if (targetColIndex4 >= 0) sheet.getRange(i + 1, targetColIndex4 + 1).setValue('ZZZUBSCASH');
    }
  }
}


function copySecondRowToFirstAndDelete() {
  const sheetName = 'PREPRC'; // Replace with your sheet name
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`Sheet "${sheetName}" not found.`);
    return;
  }
  
  // Get the data from the second row (row 2)
  // The range starts at row 2, column 1, goes for 1 row, and covers all columns
  const lastCol = sheet.getLastColumn();
  const sourceRange = sheet.getRange(2, 1, 1, lastCol);
  const sourceValues = sourceRange.getValues();

  // Set the data into the first row (row 1)
  const targetRange = sheet.getRange(1, 1, 1, lastCol);
  targetRange.setValues(sourceValues);

  // Delete the original second row
  sheet.deleteRow(2);
}

function searchReplaceByHeader() {
  const sheetName = 'PREPRC'; // Replace with your sheet name
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`Sheet "${sheetName}" not found.`);
    return;
  }

  const headerToFind = 'SYMBOL'; // Replace with your column header
  const searchTerm = 'GOOGL'; // The value to search for
  const replaceWith = 'GOOG'; // The value to replace with

  // Get all header values from the first row
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Find the column index (0-based)
  const columnIndex = headers.indexOf(headerToFind);

  if (columnIndex === -1) {
    Logger.log(`Header "${headerToFind}" not found.`);
    return;
  }

  // Column indices in Apps Script are 1-based when using getRange with row/column numbers
  const columnNumber = columnIndex + 1;
  const lastRow = sheet.getLastRow();

  // Define the specific range for the target column, starting from the second row (after header)
  // Ensure we handle cases where lastRow is 1 (only a header row)
  if (lastRow > 1) {
    const columnRange = sheet.getRange(2, columnNumber, lastRow - 1, 1);
    
    // Use TextFinder for efficient search and replace
    columnRange.createTextFinder(searchTerm)
      .replaceAllWith(replaceWith);
      
    Logger.log(`Replaced all occurrences of "${searchTerm}" with "${replaceWith}" in column "${headerToFind}".`);
  } else {
    Logger.log(`No data rows found in column "${headerToFind}".`);
  }
}

function preProcess() {
  copySecondRowToFirstAndDelete();
  copyValuesBasedOnCondition();
  searchReplaceByHeader();
  deleteColumnsByHeader();
}