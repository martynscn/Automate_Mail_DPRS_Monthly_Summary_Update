function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2; // First row of data to process
  var numRows = 2;  // Number of rows to process
  // Fetch the range of cells A2:B3
  var dataRange = sheet.getRange(startRow, 1, numRows, 2)
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (i in data) {
    var row = data[i];
    var emailAddress = row[0]; // First column
    var message = row[1];      // Second column
    var subject = "Sending emails from a Spreadsheet";
    MailApp.sendEmail(emailAddress, subject, message);
  }
  
}





// This constant is written in column C for rows for which an email has been sent successfully
var EMAIL_SENT = "EMAIL SENT";

function sendEmails2() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2; // First row of data to process
  var numRows = 2;  // Number of rows to process
  // Fetch the range of cells A2:C3
  var dataRange = sheet.getRange(startRow, 1, numRows, 3)
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var emailAddress = row[0]; // First column
    var message = row[1];      // Second column
    var emailSent = row[2];    // Third column
    if (emailSent != EMAIL_SENT) {  // Prevents sending duplicates
      var subject = "Sending emails from a Spreadsheet";
      MailApp.sendEmail(emailAddress, subject, message);
      sheet.getRange(startRow +i, 3).setValue(EMAIL_SENT);
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  }
  // Get remaining email quota
  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  Logger.log("Remaining email quota: " + emailQuotaRemaining);
  
}

// Third function

function sendEmails3() {
  var e4eLogoUrl = "http://ehealth4everyone.com/wp-content/uploads/2015/08/ehealth_teal_green_70px.png";
  var e4eLogoBlob = UrlFetchApp
                      .fetch(e4eLogoUrl)
                      .getBlob()
                      .setName("eHealth4everyone Blob");
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2; // First row of data to process
  var numRows = 2;  // Number of rows to process
  // Fetch the range of cells A2:C3
  var dataRange = sheet.getRange(startRow, 1, numRows, 3)
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
 for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var emailAddress = row[0]; // First column
    var message = row[1];      // Second column
    var emailSent = row[2];    // Third column
    if (emailSent != EMAIL_SENT) {  // Prevents sending duplicates
      var subject = "Sending emails from a Spreadsheet";
      // MailApp.sendEmail(emailAddress, subject, message);
      MailApp.sendEmail({
        to: emailAddress,
        subject: subject,
        body: message,
        htmlBody: "inline e4e Logo<img src='cid:e4eLogo'>images! <br>" +
                   "inline e4e Logo<img src='cid:e4eLogo'>",
        inlineImages:
          {
            e4eLogo:e4eLogoBlob
          }
        });
      sheet.getRange(startRow +i, 3).setValue(EMAIL_SENT);
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  }
  // Get remaining email quota
  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  Logger.log("Remaining email quota: " + emailQuotaRemaining);
  
}

// Fourth function

// This constant is written in column C for rows for which an email has been sent successfully
var EMAIL_SENT = "EMAIL SENT";



function sendEmails4() {

  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2; // First row of data to process
  var numRows = 2;  // Number of rows to process
  // Fetch the range of cells A2:C3
  var dataRange = sheet.getRange(startRow, 1, numRows, 3)
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var emailAddress = row[0]; // First column
    var message = row[1];      // Second column
    var emailSent = row[2];    // Third column
    // Send an email with two attachments: a file from Google Drive (as a PDF) and an HTML file.
    var file = DriveApp.getFileById("1FAahLvhD_Afu5FjgISVS7W38fPVe6tletAPL1O0DT3I")
    var blob = Utilities.newBlob('Insert any HTML content here', 'text/html','my_document.html');
    if (emailSent != EMAIL_SENT) {  // Prevents sending duplicates
      var subject = "Sending emails from a Spreadsheet";
      MailApp.sendEmail(emailAddress, 'Attachment example', 'Two files are attached.',{
        name: 'eHealth4everyone training team',
        attachments: [blob, file.getAs(MimeType.PDF)]
      });
      sheet.getRange(startRow +i, 3).setValue(EMAIL_SENT);
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  }
  // Get remaining email quota
  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  Logger.log("Remaining email quota: " + emailQuotaRemaining);
  
}

// Fifth function - user defined subject and sender name
function sendEmails5() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Report page');
  var startRow = 6; // First row of data to process
  var col = sheet.getRange('B' + startRow + ':B').getValues();
  var numRows = col.filter(String).length;
  var cc = sheet.getRange('B3').getValue();
  
  var emailaddressSent = [];
  var emailMessagesSent = [];
  var emailSentTime = [];
 
  // Fetch the range of cells A6:C7
  var dataRange = sheet.getRange(startRow, 1, numRows, 3);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var emailAddress = row[0]; // First column
    var message = row[1].toString();      // Second column
    var plainMessage = message.replace(/\<br\/\>/gi,'\n').replace(/(<([^>]+)>)/ig,"");
    var emailSent = row[2];    // Third column
    
    if (emailSent != EMAIL_SENT) {  // Prevents sending duplicates

      Logger.log("Email Address: " + emailAddress);
      var subject = sheet.getRange(1, 2).getValue();
      MailApp.sendEmail(emailAddress, subject,plainMessage,
                        {name: sheet.getRange(2, 2).getValue(),
                         htmlBody:message
//                         ,cc: cc
                        }
                       );
      sheet.getRange(startRow +i, 3).setValue(EMAIL_SENT);
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
      emailaddressSent.push(emailAddress);
      emailMessagesSent.push(plainMessage);
      emailSentTime.push(new Date());
    }
  }
  // Get remaining email quota
  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  Logger.log("Remaining email quota: " + emailQuotaRemaining);
  
  var logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Email logs');
  var logCol = logSheet.getRange('B1:B').getValues();
  var logNumRows = logCol.filter(String).length;
  var logUpdate = transpose([emailSentTime,emailaddressSent,emailMessagesSent]);
  logSheet.getRange(logNumRows + 1, 2, logUpdate.length, logUpdate[0].length).setValues(logUpdate);
  Browser.msgBox("Emails sent");
}

function testFilterData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Report page');
  var startRow = 6; // First row of data to process
  var col = sheet.getRange('B' + startRow + ':B').getValues();
  var numRows = col.filter(String).length;
 
//  var numRows = (sheet.getLastRow() - startRow + 1);

  Logger.log(length);
}

function filterData(data) {
  var outputData = data.filter(function(el) {
    return (el[1] == "");
  })
}

function InputData() {
  var inputData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Retrieve').getRange('A1:G1591').getValues();
  var outputData = inputData.filter(
    function (el) {
      return (el[2] == '39 min')
    }
  );
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('allData').getRange(1,1,outputData.length, outputData[0].length).setValues(outputData);
  
}

function transpose(a)
{
  return Object.keys(a[0]).map(function (c) { 
                                  return a.map(function (r) { 
                                                return r[c];
                                              }); 
                                });
}