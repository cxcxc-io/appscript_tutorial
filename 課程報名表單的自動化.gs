
// 用先前的google form，透過其開啟sheet，確認有一個sheet為催繳名單
function onFormSubmit(e) {
  // 获取表单提交数据
  var submittedData = e.values;

  // 指定工作表 ID（用您的实际工作表 ID 替换下面的字符串）
  var spreadsheetId = '你的Spreadsheet-ID';

  // 指定工作表名称（用您的实际工作表名称替换下面的字符串）
  var sheetName = '催繳名單';

  // 在指定工作表的指定分页上设置 "是否已繳費" 列为 False 并以核取方块的方式呈现
  setFalseWithCheckbox(spreadsheetId, sheetName);
}

// 為新添加的內容增加一個未核銷的欄位
function setFalseWithCheckbox(spreadsheetId, sheetName) {
  // 打开指定工作表
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spreadsheet.getSheetByName(sheetName);

  // 获取表头行
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // 查找 "是否已繳費" 列的索引
  var columnIndex = headers.indexOf("是否已繳費");

  // 如果找到 "是否已繳費" 列
  if (columnIndex !== -1) {
    // 获取新数据的行号
    var lastRow = sheet.getLastRow() + 1;

    // 在 "是否已繳費" 列中设置为 False
    sheet.getRange(lastRow, columnIndex + 1).setValue(false);

    // 在 "是否已繳費" 列中创建核取方块
    sheet.getRange(lastRow, columnIndex + 1).insertCheckboxes();
  }
}
