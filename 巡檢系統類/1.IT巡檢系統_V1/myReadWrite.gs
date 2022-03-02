//讀取google sheet指定行數+寫入其他google sheet指定行數
function rwSheet() {
  Logger.log("開始讀取資料...")
  //這裡開始讀取APP暫存的sheet表
  var set_url = "-------------------APP 連動的URL--------------------- ";
  var set_name = "轉存區";
  var set_range = "A2:L2";

  var ss = SpreadsheetApp.openByUrl(set_url);      //取得網址
  var sheetName = ss.getSheetByName(set_name);     //分頁名
  var ranges = sheetName.getRange(set_range);      //選取範圍
  var data = ranges.getValues();                   //讀取值
  Logger.log(data[0]);  //檢查資料
  Logger.log("完成!!")
  Logger.log("開始寫入資料...")
  //這裡開始修改Datastudio來源顯示的sheet表
  var set_url = "---------------Data Studio 連動的URL----------------- ";
  var set_name = "log";
  var ss = SpreadsheetApp.openByUrl(set_url);      //取得網址
  var sheetName = ss.getSheetByName(set_name);     //分頁名
  sheetName.insertRows(2);   //預做空行

  //寫入指定區域
  var ranges = sheetName.getRange(set_range);      //選取範圍
  ranges.setValues(data);
  Logger.log("完成!!")
}