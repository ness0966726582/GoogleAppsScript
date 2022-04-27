//建立下拉式選單-ID
function myChoose1() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B3:B100').activate();
  spreadsheet.getRange('B3:B100').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true)
  .requireValueInRange(spreadsheet.getRange('List_ID'), true)
  .build());
};

//建立下拉式選單-狀態
function myChoose2() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('I3:I100').activate();
  spreadsheet.getRange('I3:I100').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true)
  .requireValueInRange(spreadsheet.getRange('DB_Status'), true)
  .build());
};

//建立下拉式選單-維護者
function myChoose3() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('J3:J100').activate();
  spreadsheet.getRange('J3:J100').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true)
  .requireValueInRange(spreadsheet.getRange('DB_Maintainer'), true)
  .build());
};