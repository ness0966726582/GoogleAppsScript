function myFunction() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('Q2').activate();
  spreadsheet.getCurrentCell().setFormula('=if(DAYS(A2,NOW())>-30,"0",)');
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('Q2:Q100'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('Q2:Q100').activate();
};