/*** @OnlyCurrentDoc */

function XNScopySecondRowToFirstAndDelete() {
  const sheetName = 'XNS'; // Replace with your sheet name
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

function XNSrearrangeColumnsByHeader() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  // 1. Define your desired header order
  const targetOrder = ["Date", "Account Name", "Description", "Category", "Tags", "Amount", "Firm Name"];
  
  // 2. Get current headers (Row 1)
  const lastCol = sheet.getLastColumn();
  const currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // 3. Reorder logic
  targetOrder.forEach((headerName, index) => {
    const targetIndex = index + 1; // 1-based index for Apps Script
    
    // Find where the header is currently located
    const currentIndex = currentHeaders.indexOf(headerName) + 1;
    
    // FIX: Only move if the column is found AND not already at the target position
    if (currentIndex > 0 && currentIndex !== targetIndex) {
      const columnToMove = sheet.getRange(1, currentIndex);
      sheet.moveColumns(columnToMove, targetIndex);
      
      // Update local array to track the new positions for the next iteration
      const [movedHeader] = currentHeaders.splice(currentIndex - 1, 1);
      currentHeaders.splice(targetIndex - 1, 0, movedHeader);
    }
  });
}

function XNSPreProcess() {
  XNScopySecondRowToFirstAndDelete();
  XNSrearrangeColumnsByHeader();
}
