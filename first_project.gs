function myFunction() {
  Logger.log("第一次體驗");
}

// 第二個function，讓學員知道在開發時，可以選擇特定function執行。
function secondFunction(){

  Logger.log("第二個任務分派");
  // myFunction();
}

// 有些數字是後續會持續多次調用的，建議可以用變數的方式儲存下來，供多次調用
function thirdFunction(){
  var cxcxc = 10;
  Logger.log(cxcxc+1);
}

function 變數宣告(){

  var 可控變數 = 5;

  Logger.log(3+可控變數);
  Logger.log(可控變數*100);

}

// 帶參數的函數
function fourthFunction(cxcxc){
  Logger.log(cxcxc+1);
}

// fourthFunction(456);


// 使用appscript程式語言，透過Gmail寄信給指定Email
function sendEmail() {
  // 设置收件人的电子邮件地址
  var recipientEmail = "example@example.com";
  
  // 设置邮件主题和正文
  var subject = "这是邮件主题";
  var body = "这是邮件正文。";

  // 发送邮件
  mailApp.sendEmail(recipientEmail, subject, body);
  Logger.log("信件已寄送")
}

// 以迴圈的方式，寄多封信件
function sendEmails() {
  // 定義Email地址陣列，未來可以改用讀取google sheet
  var emailAddresses = ["example@example.com", "example2@example.com"];

  // 寄信迴圈
  for (var i = 0; i < emailAddresses.length; i++) {
    var emailAddress = emailAddresses[i];
    var subject = "主題"; // 設定郵件主題
    var message = "郵件內容"; // 設定郵件內容

    // 使用GmailApp寄信
    GmailApp.sendEmail(emailAddress, subject, message);
  }
}

// 以迴圈的方式，寄多封信件
function sendIfNotGmails() {
  // 定義Email地址陣列，未來可以改用讀取google sheet
  var emailAddresses = ["example@example.com", "example2@example.com"];


  // 寄信迴圈
  for (var i = 0; i < emailAddresses.length; i++) {

    var emailAddress = emailAddresses[i];

    if(emailAddress.endsWith("@gmail.com")){
      var subject = "主題"; // 設定郵件主題
      var message = "郵件內容"; // 設定郵件內容

      // 使用GmailApp寄信
      GmailApp.sendEmail(emailAddress, subject, message);

    }
  }
}