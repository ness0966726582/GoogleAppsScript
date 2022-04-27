function myFunction() {
  //指定範圍取得現在時間公式
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A2').activate();
  spreadsheet.getCurrentCell().setFormula('=now()');
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('A2'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('A2').activate();
};
