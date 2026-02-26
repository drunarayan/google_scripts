/** @OnlyCurrentDoc */

function insert10rows() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('18:27').activate();
  spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 10);
  spreadsheet.getActiveRange().offset(0, 0, 10, spreadsheet.getActiveRange().getNumColumns()).activate();
};