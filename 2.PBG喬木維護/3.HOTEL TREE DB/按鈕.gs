//按鈕1還原初始狀態
function myButton1() {
  myClean();                                //清空範圍  
  myFormula1();myFormula2();myFormula3();myFormula4();;myFormula5();myFormula6();   //寫入公式1~3
  myChoose1();myChoose2();myChoose3();      //寫入下拉式1~3
  myCheckBox();                             //建立複選框
}

//按鈕2-單筆發送目標gs,在加回1行
function myButton2() {
  myAdd_Time(); //增加目前時間
  myFormula1();myFormula2();myFormula3();myFormula4();;myFormula5();myFormula6();   //寫入公式1~3
  SpreadsheetApp.flush();
  Utilities.sleep(100);
  mySend1();    //檢查內容(可有可無)
  mySend2();    //"讀取"目前行數資訊+"寫入"指定目標頁面
  myAdd1();     //新增行數+1行
  myDelete();   //刪除已傳送行數-1
  myFormula1();myFormula2();myFormula3();   //寫入公式1~3
}