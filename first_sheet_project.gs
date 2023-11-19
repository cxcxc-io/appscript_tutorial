// 知曉如何指定sheet
function updateCellA3(sheetId, sheetName) {
  // 通过Sheet ID 获取目标表格
  var spreadsheet = SpreadsheetApp.openById(sheetId);
  
  // 通过分页名称获取指定的分页
  var sheet = spreadsheet.getSheetByName(sheetName);
  
  if (sheet) {
    // 获取A3单元格
    var cell = sheet.getRange("A3");
    
    // 将单元格的值设置为5
    cell.setValue(5);
    
    Logger.log("A3单元格已更新为5。");
  } else {
    Logger.log("找不到指定的分页：" + sheetName);
  }
}

// updateCellA3("你的sheet-id", "該sheet的分頁名");

// 大範圍更新內容值
function updateRangeWithCxcxc() {

  // 指定 sheetId
  var sheetId = "你的sheet-id";
  var sheetName = "該sheet的分頁名";

  // 获取指定的Spreadsheet（文档）
  var spreadsheet = SpreadsheetApp.openById(sheetId);
  
  // 获取要操作的工作表（worksheet）（这里假设是第一个工作表）
  var sheet = spreadsheet.getSheetByName(sheetName);
  
  // 获取要操作的范围（A1:B10）
  var range = sheet.getRange("A1:B10");
  
  // 设置范围内所有单元格的值为"cxcxc"
  range.setValue("cxcxc");
}
