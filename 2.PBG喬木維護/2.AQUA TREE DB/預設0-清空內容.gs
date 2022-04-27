//清空範圍(包含公式)
function myClean() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A2:P100').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
};