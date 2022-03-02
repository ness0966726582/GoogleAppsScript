//刪除已完成寫入的行數
function deleteOld() {
  Logger.log("開始刪除舊資料...")
  //這裡讀取指定"範圍"與"取值"
  var set_url = "-------------------APP 連動的URL--------------------- ";
  var set_name = "轉存區";
  
  var ss = SpreadsheetApp.openByUrl(set_url);      //取得網址
  var sheetName = ss.getSheetByName(set_name);     //分頁名
  
  sheetName.deleteRow(2);
  Logger.log("完成第2行舊資料刪除")
  sheetName.insertRows(10);
  Logger.log("+回刪除行")
  //SpreadsheetApp.flush();
  //Utilities.sleep(100);
}