// YOU NEED TO ENTER AN EMAIL ADDRESS AND THE ID OF A GOOGLE SHEETS FILE TO USE THIS TEST FUNCTION
function test()
{
  emailPO("YOUR EMAIL ADDRESS HERE", DriveApp.getFileById("ID OF TEST FILE"));
}

function emailPO(recipient, spreadsheetFile)
{
  // Check that we're able to send an email
  if (MailApp.getRemainingDailyQuota() == 0)
  {
    // If we can't send an email, announce this error in Slack and quit
    sendPlainMessage("MailApp hit daily quota. Your Sheets PO " + spreadsheetFile.getName() + " is at " + spreadsheetFile.getUrl());
    return;
  }
  // If we got here, we're able to send messages
  var options = {
    name:"CanSat PO System",
    attachments: [spreadsheetFile.getAs(MimeType.PDF)],
  }
  MailApp.sendEmail(recipient,spreadsheetFile.getName(),"Your PO is attached as a PDF.\nNeed the spreadsheet? Please follow this link: " + spreadsheetFile.getUrl(),options);
}
