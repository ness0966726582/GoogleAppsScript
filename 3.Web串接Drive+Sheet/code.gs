//1. HTML連結 抓取網頁資訊
function doGet(e) {
  var html = HtmlService.createTemplateFromFile("upload");    //html檔名
  //上傳後的訊息
  html.message = "";
  var check = html.evaluate();
  //避免無法顯示XFrame allow all
  var display = check.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return display;
}

//2. 建立GOOGLE DRIVE 資料夾+上傳
function doPost(e){
  //資料夾的位置
  var folder_id = "-------------------------------改成你的Google資料夾ID-----------------------------------"; //!!!!!!!!!!!!!!!!!!!!!!請複製你的資料夾ID!!!!!!!!!!!!!!!!!!
  var folder = DriveApp.getFolderById(folder_id);

  //建立檔案 data,type,name
  var data = Utilities.base64Decode(e.parameter.fileData);    //抓資料
  var blob = Utilities.newBlob(data, e.parameter.mimeType, e.parameter.fileName);  //抓檔案

  //上傳檔案到資料夾位置
  var fileID = folder.createFile(blob).getId();
  var fileURL = "https://drive.google.com/file/d/"+fileID+"/view";

  //執行-自定義函數 下方3  listRecord()
  listRecord(e.parameter.workID, e.parameter.nickName, e.parameter.description, fileURL);

  //上傳成功訊息
  var htmlOutputer = HtmlService.createTemplateFromFile("upload");
  htmlOutputer.message = "上傳成功!!!"
  return htmlOutputer.evaluate();
}

//HTML與App Script上傳路徑
function getURL(){
  var url = ScriptApp.getService().getUrl();
  return url;
}

//3. 上傳資料到Google 試算表紀錄
function listRecord(workID, nickName, description, fileURL){  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //var sheet = ss.getSheets()[0];              //第一個分頁
  var sheet = ss.getSheetByName("-------------------------------改成你的googlesheet 分頁名-----------------------------------");  //自訂分頁名
  sheet.appendRow([workID, nickName, fileURL, description, new Date()]);
}