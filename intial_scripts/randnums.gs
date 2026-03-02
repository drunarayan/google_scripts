function timeNow() {
  var d = new Date();
  var currentTime = d.toLocaleTimeString(); // "12:35 PM", for instanc
  return currentTime
}

function refreshUserProps() {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('loopCounter', 0);
  userProperties.setProperty('loopLimit', 480);
  //userProperties.setProperty('loopLimit', 5);
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

function testDailies() {
  // temp set user properties
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('loopCounter', 0);
  userProperties.setProperty('loopLimit', 5);

  // clear out the sheet
    clearData();
   
  // create trigger to run program automatically
    createTrigger();
}

