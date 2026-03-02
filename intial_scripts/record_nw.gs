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
    var GLD =sheet.getRange("B23").getValue();
    var CCB =sheet.getRange("B27").getValue();

    let holds  = SpreadsheetApp.getActive().getSheetByName('holdings')
    var MGS =holds.getRange("S210").getValue();
    var CSH =holds.getRange("Y221").getValue();
    var STK =holds.getRange("Z218").getValue();
    var BND =holds.getRange("Y218").getValue();
    var IFX =holds.getRange("Y222").getValue();
    var IEQ =holds.getRange("Z222").getValue();
    var ITL =holds.getRange("W210").getValue();
    var TIN =holds.getRange("Y223").getValue();
    var TCS =sheet.getRange("Y224").getValue()+CSH;


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
      rowdata = [[ODTE, TOT,TOT-YTOT,(TOT-YTOT)/YTOT,UBS,SCH,HSA,CSH,STK,BND,TCS,GLD,SPX,(SPX-YSPX)/YSPX,MGS,CCB,IEQ,IFX,ITL,(IEQ-ITL)/TIN,IFX/TIN,ITL/TIN,TIN]]
    } else {
      tgtsheet.insertRowBefore(2);
      rowdata = [[date2, TOT,TOT-OTOT,(TOT-OTOT)/OTOT,UBS,SCH,HSA,CSH,STK,BND,TCS,GLD,SPX,(SPX-OSPX)/OSPX,MGS,CCB,IEQ,IFX,ITL,(IEQ-ITL)/TIN,IFX/TIN,ITL/TIN,TIN]]
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

