function sendTestEmail(recipient)
{
  var prefs = emailGetPrefs();
  var options = {
    name:prefs["email-name"],
    cc:prefs["email-cc"],
  };
  MailApp.sendEmail(recipient,
    "Test Email",
    prefs["email-body"] + ((prefs["email-include-link"] == "true") ? "\nReal email will include PO link" : ""),
    options);
}

function emailPO(recipient, spreadsheetFile)
{
  // Load settings
  var properties = PropertiesService.getDocumentProperties();
  var ssName = spreadsheetFile.getName();
  var ssUrl = spreadsheetFile.getUrl();
  // Check that we're able to send an email
  if (MailApp.getRemainingDailyQuota() == 0)
  {
    // If we can't send an email, announce this error in Slack and quit
    sendPlainMessage("MailApp hit daily quota. Your PO " + ssName + " is at " + ssUrl);
    return;
  }
  // If we got here, we're able to send messages
  var options = {
    name:properties.getProperty('email-name'),
    attachments: [spreadsheetFile.getAs(MimeType.PDF)],
  };
  MailApp.sendEmail(recipient,
    ssName,
    properties.getProperty('email-body') + (properties.getProperty('email-include-link') == "true") ? ssUrl : "",
    options);
}