//增加公式vlookup取得資訊

//指定範圍取得Location公式
function myFormula1() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C2').activate();
  spreadsheet.getCurrentCell().setFormula('=IFERROR(vlookup(B2,List!$A:$I,2,false),)');
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('C2:C100'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('C2:C100').activate();
};

//指定範圍取得Area公式
function myFormula2() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D2').activate();
  spreadsheet.getCurrentCell().setFormula('=IFERROR(vlookup(B2,List!$A:$I,3,false),)');
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('D2:D100'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('D2:D100').activate();
};

//指定範圍取得Coordinate公式
function myFormula3() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('E2').activate();
  spreadsheet.getCurrentCell().setFormula('=IFERROR(vlookup(B2,List!$A:$I,4,false),)');
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('E2:E100'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('E2:E100').activate();
};

//指定範圍取得Full Name公式
function myFormula4() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('F2').activate();
  spreadsheet.getCurrentCell().setFormula('=IFERROR(vlookup(B2,List!$A:$I,5,false),)');
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('F2:F100'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('F2:F100').activate();
};

//指定範圍取得Name(CH)公式
function myFormula5() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('G2').activate();
  spreadsheet.getCurrentCell().setFormula('=IFERROR(vlookup(F2,DB!$C:$I,2,false),)');
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('G2:G100'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('G2:G100').activate();
};

//指定範圍取得Name(EN)公式
function myFormula6() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('H2').activate();
  spreadsheet.getCurrentCell().setFormula('=IFERROR(vlookup(F2,DB!$C:$I,3,false),)');
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('H2:H100'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('H2:H100').activate();
};