//************** MY排程 *****************
function myschedule(){
  myformula();
  SpreadsheetApp.flush();
  Utilities.sleep(500);
  checkUpdat();
  SpreadsheetApp.flush();
  Utilities.sleep(500);
  deleteOld();
}

//************** 寫入公式 *****************
function myformula(){//檢查特定欄位公式狀態
  Logger.log("開始檢查公式...")
  var set_url = "https://docs.google.com/spreadsheets/d/1lco2glUg4v6n3mohYH1zHOXE9sEqOrmtmg-1GwsjO4U/edit#gid=0";
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
    ss.getRange("H2").setFormula("=0")     //巡檢一次改變大小或顏色    
    Logger.log("完成寫入公式")
    }else{
      Logger.log("無資料!! 不需要公式")
      }
}

//************** 檢查公式+判斷上傳 *****************
function checkUpdat(){//檢查指定空格是否皆以填上(座標忽略)
  Logger.log("開始檢查更新....")
  var set_url = "https://docs.google.com/spreadsheets/d/1lco2glUg4v6n3mohYH1zHOXE9sEqOrmtmg-1GwsjO4U/edit#gid=0";
  var set_name = "轉存區";
  var set_range = "A2:L2";

  var ss = SpreadsheetApp.openByUrl(set_url);      //取得網址
  var sheetName = ss.getSheetByName(set_name);     //分頁名
  var ranges = sheetName.getRange(set_range);                //選取範圍
  var data = ranges.getValues();                            //讀取值
  //Logger.log(String(data[0])); //A2~L2
  
  if (Boolean(data[0][6]) !== false){ //Status(Auto) 檢查資料 有值=True
      Logger.log(data[0][6]);
      Logger.log("檢查有公式ok !!");
      rwSheet(); 
      }else{
          Logger.log("檢查公式失敗再次檢查是否增加公式");
          myformula();
          }
}

//************** 讀取+上傳 *****************
function rwSheet() {//讀取google sheet指定行數+寫入其他google sheet指定行數
  Logger.log("開始讀取資料...")
  //這裡開始讀取APP暫存的sheet表
  var set_url = "https://docs.google.com/spreadsheets/d/1lco2glUg4v6n3mohYH1zHOXE9sEqOrmtmg-1GwsjO4U/edit#gid=0";
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
  var set_url = "https://docs.google.com/spreadsheets/d/1cS05H8lKUJp-DiEXqAO6WpT-lTj9CBOtsPViAkqBDe8/edit#gid=0";
  var set_name = "log";
  var ss = SpreadsheetApp.openByUrl(set_url);      //取得網址
  var sheetName = ss.getSheetByName(set_name);     //分頁名
  sheetName.insertRows(2);   //預做空行

  //寫入指定區域
  var ranges = sheetName.getRange(set_range);      //選取範圍
  ranges.setValues(data);
  Logger.log("完成!!")
}

//************** 刪除行後+回 *****************
function deleteOld() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('2:2').activate();
  spreadsheet.getActiveSheet().deleteRows(spreadsheet.getActiveRange().getRow(), spreadsheet.getActiveRange().getNumRows());
  spreadsheet.getRange('9:9').activate();
  spreadsheet.getActiveSheet().insertRowsAfter(spreadsheet.getActiveRange().getLastRow(), 1);
  spreadsheet.getActiveRange().offset(spreadsheet.getActiveRange().getNumRows(), 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
};