// 嘗試把大型任務拆成多個小任務
// 先確認能爬取
// 確認chatgpt能把json轉成表格
// 確認chatgpt能依照此表格，轉成google sheet儲存

//
function fetchWebsiteData() {
  var url = "https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY?date=20231015&stockNo=2330&response=json";
  var response = UrlFetchApp.fetch(url);
  var content = response.getContentText();
  
  Logger.log(content); // 将爬取的内容打印到日志
  
  // 如果您希望将内容显示在Google表格或其他地方，请根据需求进行相应操作
}


function parseJSONToTable() {

  var jsonData = '{"stat":"OK","date":"20231015","title":"112年10月 2330 台積電           各日成交資訊","fields":["日期","成交股數","成交金額","開盤價","最高價","最低價","收盤價","漲跌價差","成交筆數"],"data":[["112/10/02","26,891,996","14,300,676,237","530.00","534.00","528.00","533.00","+10.00","23,151"],["112/10/03","17,601,440","9,335,110,434","528.00","533.00","528.00","529.00","-4.00","17,430"],["112/10/04","29,808,729","15,519,741,014","521.00","523.00","519.00","520.00","-9.00","49,176"],["112/10/05","25,749,184","13,563,941,939","523.00","529.00","523.00","528.00","+8.00","19,027"],["112/10/06","16,160,314","8,587,684,448","530.00","533.00","529.00","532.00","+4.00","14,047"],["112/10/11","61,670,562","33,477,968,593","542.00","544.00","540.00","544.00","+12.00","43,385"],["112/10/12","36,350,241","19,916,786,786","545.00","550.00","544.00","550.00","+6.00","35,333"],["112/10/13","34,704,924","19,128,100,250","550.00","554.00","548.00","553.00","+3.00","31,159"]],"notes":["符號說明:+/-/X表示漲/跌/不比價","當日統計資訊含一般、零股、盤後定價、鉅額交易，不含拍賣、標購。","ETF證券代號第六碼為K、M、S、C者，表示該ETF以外幣交易。","權證證券代號可重複使用，權證顯示之名稱係目前存續權證之簡稱。"],"total":8}';
  // 将JSON字符串解析为JavaScript对象
  var jsonData = JSON.parse(jsonData);

  // 获取数据部分
  var data = jsonData.data;
  var fields = jsonData.fields;

  // 获取目标电子表格
  var spreadsheet = SpreadsheetApp.openById('你的google-spreadsheet-id'); // 用您的电子表格ID替换 'YOUR_SPREADSHEET_ID'
  var targetSheetName = "工作表1";
  var sheet = spreadsheet.getSheetByName(targetSheetName);

  // 如果找不到指定名称的工作表，则创建一个新的工作表
  if (!sheet) {
    sheet = spreadsheet.insertSheet(targetSheetName);
  }

  // 设置表头
  sheet.getRange(1, 1, 1, fields.length).setValues([fields]);

  // 填充数据
  for (var i = 0; i < data.length; i++) {
    var rowData = data[i];
    sheet.getRange(i + 2, 1, 1, rowData.length).setValues([rowData]);
  }

}


//
//
//
// 完整大合併程式碼
//
//

//把資料解析成json格式
function parseDataFromContent(content) {
  // 在这里编写解析网页内容的代码，将其转化为JSON格式的数据
  // 示例代码中，假设content是一个JSON字符串，您需要根据实际情况解析内容
  var jsonData = JSON.parse(content);

  return jsonData;
}

// 能把json資料寫到指定的資料表
function writeDataToSheet(jsonData, spreadsheetId,targetSheetName) {
  // 获取目标电子表格
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId); // 替换为您的电子表格ID
  var sheet = spreadsheet.getSheetByName(targetSheetName);

  // 如果找不到指定名称的工作表，则创建一个新的工作表
  if (!sheet) {
    sheet = spreadsheet.insertSheet(targetSheetName);
  }

  // 获取数据部分
  var data = jsonData.data;
  var fields = jsonData.fields;

  // 设置表头
  sheet.getRange(1, 1, 1, fields.length).setValues([fields]);

  // 填充数据
  for (var i = 0; i < data.length; i++) {
    var rowData = data[i];
    sheet.getRange(i + 2, 1, 1, rowData.length).setValues([rowData]);
  }
}


function scrapeAndParseData() {
  // 指定要爬取的网址
  var url = "https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY?date=20231015&stockNo=2345&response=json"; // 替换为您要爬取的网址

  // 使用UrlFetchApp爬取网页内容
  var response = UrlFetchApp.fetch(url);
  var content = response.getContentText();

  // 解析爬取的数据
  var jsonData = parseDataFromContent(content);

  // 将数据写入到指定的Google Sheets
  var spreadsheetId = "你的spreadsheet id"
  var targetSheetName = "你想儲存的sheet名"; // 替换为您的目标工作表名称
  writeDataToSheet(jsonData, spreadsheetId, targetSheetName);
}




