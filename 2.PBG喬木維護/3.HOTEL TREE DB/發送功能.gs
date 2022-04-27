//增加目前時間
function myAdd_Time() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A3').activate();
  spreadsheet.getCurrentCell().setFormula('=now()');
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('A3'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('A3').activate();
}

//發送-第一步判斷資料
function mySend1() {
  
}

//發送-第2步讀取資料+寫入資料
function mySend2() {
  Logger.log("開始檢查更新....")
  var set_url = "-------------------------------輸入"當前" Googlesheet 的URL----------------------------------------";   	//!!!修改!!!!
  var set_name = "Input_area";
  var set_range = "A3:P3";

  var ss = SpreadsheetApp.openByUrl(set_url);      //取得網址
  var sheetName = ss.getSheetByName(set_name);     //分頁名
  var ranges = sheetName.getRange(set_range);                //選取範圍
  var data = ranges.getValues();                            //讀取值
  Logger.log(String(data[0])); //A3~P3
  Logger.log("開始寫入資料...")
  
  //這裡開始修改Datastudio來源顯示的sheet表
  var set_url = "-------------------------------輸入"目標" Googlesheet 的URL----------------------------------------";	//!!!修改!!!!	
  var set_name = "PBG_Tree";
  var set_range = "A2:P2";
  
  var ss = SpreadsheetApp.openByUrl(set_url);      //取得網址
  var sheetName = ss.getSheetByName(set_name);     //分頁名
  sheetName.insertRows(2);   //預做空行

  //寫入指定區域
  var ranges = sheetName.getRange(set_range);      //選取範圍
  Logger.log(data[0]);
  ranges.setValues(data);
  Logger.log("完成!!");
}

//發送-行數+1
function myAdd1() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('10:10').activate();
  spreadsheet.getActiveSheet().insertRowsAfter(spreadsheet.getActiveRange().getLastRow(), 1);
  spreadsheet.getActiveRange().offset(spreadsheet.getActiveRange().getNumRows(), 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
}
//刪除已完成寫入的行數-1
function myDelete() {
  Logger.log("開始刪除舊資料...")
  //這裡讀取指定"範圍"與"取值"
  var set_url = "-------------------------------輸入"當前" Googlesheet 的URL----------------------------------------";   	//!!!修改!!!!		//!!!!!!!!!目前 : Googlesheet 的URL
  var set_name = "Input_area";
  
  var ss = SpreadsheetApp.openByUrl(set_url);      //取得網址
  var sheetName = ss.getSheetByName(set_name);     //分頁名
  
  sheetName.deleteRow(3);
  Logger.log("完成第3行舊資料刪除")
}