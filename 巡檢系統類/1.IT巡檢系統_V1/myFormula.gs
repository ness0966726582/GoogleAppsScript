//檢查特定欄位公式狀態
function myformula(){
  Logger.log("開始檢查公式...")
  var set_url = "-------------------APP 連動的URL--------------------- ";
  var set_name = "轉存區";
  var set_range = "A2:H2";

  var ss = SpreadsheetApp.openByUrl(set_url);      //取得網址
  var sheetName = ss.getSheetByName(set_name);     //分頁名
  var ranges = sheetName.getRange(set_range);                //選取範圍

  var data = ranges.getValues();                            //讀取值
  //Logger.log(data[0][0] + data[0][1] + data[0][2] +data[0][3]); //A2~D2

  if (Boolean(data[0][0]) !== false){
    var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    ss.getRange("E2").setFormula("=IFERROR(vlookup(D2,'座標對照表'!$A:$D,2,false),)");  //座標
    ss.getRange("F2").setFormula("=IFERROR(vlookup(D2,'座標對照表'!$A:$D,3,false),)");  //位置
    ss.getRange("G2").setFormula("=IFERROR(vlookup(D2,'座標對照表'!$A:$D,4,false),)");  //區域
    ss.getRange("H2").setFormula("=1")     //巡檢一次改變大小或顏色    
    Logger.log("完成寫入公式")
    }else{
      Logger.log("無資料!! 不需要公式")
      }
}