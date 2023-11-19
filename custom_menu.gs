// 當開啟spreadsheet時，載入功能選單
function onOpen() {
  var ui = SpreadsheetApp.getUi(); // 或者 DocumentApp.getUi()，取決於你的應用
  ui.createMenu('自訂選單名稱')
      .addItem('製作證書', 'myFunction')
      .addItem('寄證書', 'myFunction')
      .addToUi();
}

function myFunction(){
    Logger.log("觸發")
}