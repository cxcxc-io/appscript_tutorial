// 當開啟spreadsheet時，載入功能選單
function onOpen() {
  var ui = SpreadsheetApp.getUi(); // 或者 DocumentApp.getUi()，取決於你的應用
  ui.createMenu('自訂選單名稱')
      .addItem('製作證書', '檢查繳費並打印證書')
      .addItem('寄證書', 'sendBulkEmails')
      .addToUi();
}


function 檢查繳費並打印證書() {

  // 指定Google Sheets工作簿的ID
  var spreadsheetId = "你的Spreadsheet-ID";
  
  // 指定要操作的工作表名稱
  var sheetName = "催繳名單";
  
  // 使用Google Sheets API來打開工作簿
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  
  // 選擇指定的工作表
  var sheet = spreadsheet.getSheetByName(sheetName);
  
  // 獲取工作表的所有資料
  var data = sheet.getDataRange().getValues();
  
  // 獲取欄位名稱行，通常在第一行（列數從0開始）
  var headerRow = data[0];
  
  // 確保"是否已繳費"欄位存在
  var isPaidIndex = headerRow.indexOf("是否已繳費");

  // 載入用戶的真實姓名
  var realNameIndex = headerRow.indexOf("真實姓名");

  if (isPaidIndex === -1) {
    Logger.log("找不到名稱為'是否已繳費'的欄位");
    return;
  }

  // 迭代每一行資料，從第二行開始（因為第一行是欄位名稱）
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    // 獲取"是否已繳費"欄位的值
    var isPaidValue = row[isPaidIndex];

    
    // 檢查是否已繳費
    if (isPaidValue === true || isPaidValue === "TRUE") {
      
      // 取得指定的google folder
      var folder = DriveApp.getFolderById('你放證書的資料夾Id'); 

      // 依照模板檔案複製一份檔案，並放入先前指定的Folder內
      var newFile=DriveApp.getFileById('證書模板的檔案id').makeCopy('temp', folder);

      // 讀取剛複製的副本
      var newDoc = DocumentApp.openById(newFile.getId());
      var newBody = newDoc.getBody();

      // 修改副本內指定的模板詞
      newBody.replaceText('<< 真實姓名 >>', row[realNameIndex]);
      newDoc.saveAndClose();

      // 获取新文档的Blob（二进制数据）
      var newDocBlob = newDoc.getAs('application/pdf');

      // 创建一个Blob文件
      var pdfFile = DriveApp.createFile(newDocBlob);

      // 设置PDF文件的名称
      pdfFile.setName(row[realNameIndex]+ '.pdf');
      
      // 獲取新建立的 Google Sheet 的 File
      var file = DriveApp.getFileById(pdfFile.getId());
      
      // 移動此檔案到指定資料夾
      file.moveTo(folder);

      // 找到該列儲存格的證書檔案位置，並將檔案位置放入該資料列的欄位-證書檔案位置內的cell
      var targetRange = sheet.getRange(i+1,headerRow.indexOf("證書檔案位置")+1);
      var targetCell = targetRange.getCell(1,1);
      targetCell.setValue(file.getUrl());

      // 找到該列儲存格的證書檔案位置，並將檔案位置放入該資料列的欄位-證書檔案位置內的cell
      var certificateIdRange = sheet.getRange(i+1,headerRow.indexOf("證書檔案ID")+1);
      var certificateIdCell = certificateIdRange.getCell(1,1);
      certificateIdCell.setValue(file.getId());

      // 删除新文档和临时PDF文件
      DriveApp.getFileById(newDoc.getId()).setTrashed(true);

      Logger.log("處理完成，已繳費的行: " + (i + 1));
    }
  }
}


// 信件群發
function sendBulkEmails() {
  // 指定Google Sheets工作簿的ID
  var spreadsheetId = "報名表單sheet的id";
  
  // 指定要操作的工作表名稱
  var sheetName = "催繳名單";
  
  // 使用Google Sheets API來打開工作簿
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);

  // 選擇指定的工作表
  var sheet = spreadsheet.getSheetByName(sheetName);

  var data = sheet.getDataRange().getValues();
  
  // 標題欄
  var headerRow = data[0];

  // 確保"是否已繳費"欄位存在
  var isPaidIndex = headerRow.indexOf("是否已繳費");

  // 載入用戶的真實姓名
  var realNameIndex = headerRow.indexOf("真實姓名");

  var recipientEmailIndex = headerRow.indexOf("電子郵件");
  var attachmentIdIndex =  headerRow.indexOf("證書檔案ID");

  var subject = "信件標題"; // 邮件主题，假设在第二列
  var body = "信件內文"; // 邮件正文，假设在第三列

  // 循环遍历电子表格中的每一行数据
  for (var i = 1; i < data.length; i++) { // 从第二行开始，跳过标题行

    var isPaidValue = data[i][isPaidIndex];
        // 檢查是否已繳費
    if (isPaidValue === true || isPaidValue === "TRUE") {
    var recipientEmail = data[i][recipientEmailIndex]; // 收件人的电子邮件地址，假设在第一列

    var attachmentId = data[i][attachmentIdIndex]; // 附件文件的URL，假设在第四列

    // 发送邮件
    var options = {};
    
    // 如果有附件，添加附件选项
    if (attachmentId) {
      var attachment = DriveApp.getFileById(attachmentId);
      options.attachments = [attachment];
    }
    
    MailApp.sendEmail(recipientEmail, subject, body, options);
    
    // 暂停几秒以避免达到 Google Apps Script 的电子邮件发送限制
    Utilities.sleep(2000);
    }
  }
}
