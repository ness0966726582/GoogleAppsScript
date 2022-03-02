function myschedule(){
  myformula();
  SpreadsheetApp.flush();
  Utilities.sleep(100);
  checkUpdat();   //檢查與更新
  deleteOld();
}

//檢查指定空格是否皆以填上(座標忽略)
function checkUpdat(){
  Logger.log("開始檢查更新....")
  var set_url = "-------------------APP 連動的URL--------------------- ";
  var set_name = "轉存區";
  var set_range = "A2:L2";

  var ss = SpreadsheetApp.openByUrl(set_url);      //取得網址
  var sheetName = ss.getSheetByName(set_name);     //分頁名
  var ranges = sheetName.getRange(set_range);                //選取範圍
  var data = ranges.getValues();                            //讀取值
  //Logger.log(String(data[0][11])); //A2~L2

  if (Boolean(data[0][6]) !== false){ //Status(Auto) 檢查資料 有值=True
      Logger.log(data[0][6]);
      Logger.log("檢查有公式ok !!");
      rwSheet(); 
      }else{
          Logger.log("檢查公式失敗再次檢查是否增加公式");
          myformula();
          }
}