## All scripts consolidated

## macros.gs
/** @OnlyCurrentDoc */

function insert10rows() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('18:27').activate();
  spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 10);
  spreadsheet.getActiveRange().offset(0, 0, 10, spreadsheet.getActiveRange().getNumColumns()).activate();
};

## record_nw.gs
/*** @OnlyCurrentDoc */

// Menu for testing your script
function onOpen() {
  var ui = SpreadsheetApp.getUi();
    ui.createMenu('Records')
      .addItem('Records','recordValue')
      .addToUi();
    ui.createMenu('Updates')
      .addItem('Updates','dailyUpdate')
      .addToUi();
    ui.createMenu('Dailies')
      .addItem('Dailies','dailies')
      .addToUi();
    ui.createMenu("Auto Trigger")
      .addItem("Run","runAuto")
      .addToUi();
  dailies();
  recordTime();
}

// Record history from a cell and append to next available row
function recordValue() {
   var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SUMMARY");
    var date1 = new Date();
    var SPX =sheet.getRange("B13").getValue();
    var TOT =sheet.getRange("B8").getValue();
    var UBS =sheet.getRange("B2").getValue();
    var SCH =sheet.getRange("B3").getValue();
    var HSA =sheet.getRange("B5").getValue();
    var CSH =sheet.getRange("B7").getValue();
    var STK =sheet.getRange("B20").getValue();
    var BND =sheet.getRange("B21").getValue();
    var TCS =sheet.getRange("B23").getValue();
    var GLD =sheet.getRange("B24").getValue();
    sheet.appendRow([date1, TOT,0,0,UBS,SCH,HSA,CSH,STK,BND,TCS,GLD,SPX]);
}

// Record history from a cell and append to next available row
function dailies() {
    let tgtsheet  = SpreadsheetApp.getActive().getSheetByName('dailies');
    //var cell =tgtsheet.getRange("AA1"); // Store new date in AA1
    //cell.setValue(date2);
    var OSPX =tgtsheet.getRange("M2").getValue();
    var YSPX =tgtsheet.getRange("M3").getValue();
    var OTOT =tgtsheet.getRange("B2").getValue();
    var YTOT =tgtsheet.getRange("B3").getValue();
    var ODTE =tgtsheet.getRange("A2").getValue();
    let range  = tgtsheet.getRange('2:2');
    let values = range.getValues();

    let sheet  = SpreadsheetApp.getActive().getSheetByName('summary');
    var SPX =sheet.getRange("B13").getValue();
    var TOT =sheet.getRange("B8").getValue();
    var UBS =sheet.getRange("B2").getValue();
    var SCH =sheet.getRange("B3").getValue();
    var HSA =sheet.getRange("B5").getValue();
    var STK =sheet.getRange("B20").getValue();
    var BND =sheet.getRange("B21").getValue();
    var TCS =sheet.getRange("B23").getValue();
    var GLD =sheet.getRange("B24").getValue();
    var CCB =sheet.getRange("B27").getValue();

    let holds  = SpreadsheetApp.getActive().getSheetByName('holdings')
    var MGS =holds.getRange("S164").getValue();
    var CSH =holds.getRange("B7").getValue();

    var replace = false;;
    var date2 = new Date();
    //var date2 = dateOnly(dateX)
    Logger.log(date2);
    Logger.log(ODTE);
        if (date2.getFullYear() === ODTE.getFullYear() &&
      date2.getMonth() === ODTE.getMonth() &&
      date2.getDate() === ODTE.getDate()) replace = true;
    if (replace) {
      Logger.log("Dates are SAME, Replacing Dailies row!");
      rowdata = [[ODTE, TOT,TOT-YTOT,(TOT-YTOT)/YTOT,UBS,SCH,HSA,CSH,STK,BND,TCS,GLD,SPX,(SPX-YSPX)/YSPX,MGS,CCB,0,0,0,0,0]]
    } else {
      tgtsheet.insertRowBefore(2);
      rowdata = [[date2, TOT,TOT-OTOT,(TOT-OTOT)/OTOT,UBS,SCH,HSA,CSH,STK,BND,TCS,GLD,SPX,(SPX-OSPX)/OSPX,MGS,CCB,0,0,0,0,0]]
    }
    range.offset(0,0).setValues(rowdata);
}

function recordTime() {
  var ndate = new Date();

  let sumsheet  = SpreadsheetApp.getActive().getSheetByName('summary');
    var cell =sumsheet.getRange("A15"); // Store new time in A15
    formDate = dateOnly(ndate)
    Logger.log(ndate)
    Logger.log(formDate);
    cell.setValue(ndate);
}

function recordDateOnlySheet(aSheet, aCell) {
  var ndate = new Date();

  let sumsheet  = SpreadsheetApp.getActive().getSheetByName(aSheet);
    var cell =sumsheet.getRange(aCell); // Store new time in aCell
    formDate = dateOnly(ndate)
    Logger.log(formDate);
    cell.setValue(formDate);
}

function dailyUpdate() {
  let sheet  = SpreadsheetApp.getActive().getSheetByName('prices');
  let range  = sheet.getRange('2:2');
  let values = range.getValues();
  sheet.insertRowBefore(3);
  range.offset(1,0).setValues(values);
}

function dateOnly(aDate) {
  const specificDate = new Date(aDate.getFullYear(),aDate.getMonth(),aDate.getDate(),0,1,0)
  return specificDate;
}

function dateTimeOnly(aDate) {
  //usrtz = getUserTimeZoneFromCalendar();
  // Format as US date
  let usDate = Utilities.formatDate(
  aDate,
  //usrtz,
  Session.getScriptTimeZone(),
  "MM/dd/yyyy HH:MM:SS"
  );
  //let usDate = Utilities.formatDate(aDate, 'America/Los_angeles', 'MMMM dd, yyyy HH:mm:ss Z');
  Logger.log("US date format: " + usDate); // "03/15/2023"
  return usDate;
}

function getUserTimeZoneFromCalendar() {
  // Get the default calendar of the authenticated user
  var calendar = CalendarApp.getDefaultCalendar();  

  // Retrieve the time zone of the calendar
  var timeZone = calendar.getTimeZone();
  
  // Log the time zone to the Apps Script Logger
  Logger.log('User Time Zone from Calendar: ' + timeZone);
  
  return timeZone;  // Return the time zone for further processing
} 

function testing() {
  let dsheet  = SpreadsheetApp.getActive().getSheetByName('dailies');
  var rDate =dsheet.getRange("A2").getValue();
  let ssheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SUMMARY");
  var sDate =ssheet.getRange("C7").getValue();  
  Logger.log("Truncated dates are %s and %s",dateOnly(rDate),dateOnly(sDate));
}


## prepholdings.gs
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

## prepxns.gs
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


## invdb.gs
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

## randnums.gs
function timeNow() {
  var d = new Date();
  var currentTime = d.toLocaleTimeString(); // "12:35 PM", for instanc
  return currentTime
}

function refreshUserProps() {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('loopCounter', 0);
  userProperties.setProperty('loopLimit', 480);
  //userProperties.setProperty('loopLimit', 3);
}

function onOpen() { 
  var ui = SpreadsheetApp.getUi();
   
  ui.createMenu("Auto Trigger")
    .addItem("Run","runAuto")
    .addToUi();
}

function runAuto () {
  var today = new Date();
  var day = today.getDay();
  
  var days = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday']
  
  if(day == 6 || day == 0) {
    Logger.log('Today is ' + days[day] + '. So I\'ll do nothing')
  } else {
  // resets the loop counter if it's not 0
    refreshUserProps();
   
  // clear out the sheet
    clearData();
   
  // create trigger to run program automatically
    createTrigger();
  }
}

function clearData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rsheet = ss.getSheetByName('randy');
  var tsheet = ss.getSheetByName('today');
   
  // clear out the matches and output sheets
  var lastRow = rsheet.getLastRow();
  if (lastRow > 1) {
    rsheet.getRange(2,1,lastRow-1,1).clearContent();
  }
  //rsheet.clearContents();

  // clear out the matches and output sheets
  var lastRow = tsheet.getLastRow();
  var lastCol = tsheet.getLastColumn();
  Logger.log("sheet is empty lastRow: "+lastRow+" lastCol: "+lastCol+"No need to clear contents!");
  if (lastRow > 1) {
    tsheet.getRange(1,1,lastRow,lastCol).clearContent();
    //tsheet.clearContents();
  }

  // Log message to confirm loop is started
    rsheet.getRange(1,1).setValue("Setting Trigger");
    Logger.log("Setting Trigger");
}

function createTrigger() {
   
  // Trigger every 1 minute
  ScriptApp.newTrigger('todaysInvData')
      .timeBased()
      .everyMinutes(1)
      .create();
}

function deleteTrigger() {
   
  // Loop over all triggers and delete them
  var allTriggers = ScriptApp.getProjectTriggers();
   
  for (var i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
}

function deletetodaysInvDataTrigger() {
  var Triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < Triggers.length; i++) {
    if (Triggers[i].getHandlerFunction() == "todaysInvData") {
      ScriptApp.deleteTrigger(Triggers[i])
    }
  }
}  

function todaysInvData() {
   
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rsheet = ss.getSheetByName('randy');

  // get the current loop counter
  var userProperties = PropertiesService.getUserProperties();
  var loopCounter = Number(userProperties.getProperty('loopCounter'));
  var loopLimit = Number(userProperties.getProperty('loopLimit'));
   
  // put some limit on the number of loops
  // could be based on a calculation or user input
  // using a static number in this example
  var limit = loopLimit;
   
  // if loop counter < limit number, run the repeatable action
  if (loopCounter < limit) {
     
    // see what the counter value is at the start of the loop
    Logger.log(loopCounter);
     
    // do stuff
    var num = Math.ceil(Math.random()*100);
    //rsheet.getRange(rsheet.getLastRow()+1,1).setValue(num);
    rsheet.getRange(rsheet.getLastRow()+1,1).setValue(timeNow());
    copyRowEveryMinute();
     
    // increment the properties service counter for the loop
    loopCounter +=1;
    userProperties.setProperty('loopCounter', loopCounter);
     
    // see what the counter value is at the end of the loop
    Logger.log(loopCounter);
  }
   
  // if the loop counter is no longer smaller than the limit number
  // run this finishing code instead of the repeatable action block
  else {
    // Log message to confirm loop is finished
    rsheet.getRange(rsheet.getLastRow()+1,1).setValue("Ending Trigger");
    Logger.log("Ending Trigger");
     
    // delete trigger because we've reached the end of the loop
    // this will end the program
    //deleteTrigger();  
    deletetodaysInvDataTrigger()
  }
}





