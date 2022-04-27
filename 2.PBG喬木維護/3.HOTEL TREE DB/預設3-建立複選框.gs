//新增複選框
function myCheckBox() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('K2').activate();
  spreadsheet.getActiveRangeList().insertCheckboxes();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('K2:N2'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('K2:N2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('K2:N100'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('K2:N100').activate();
};