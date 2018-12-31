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
  var properties = PropertiesService.getScriptProperties();
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

// AUGH a global variable, even worse it's duplicated in emailSidebar.html
var email_string_params = ["email-name", "email-body", "email-cc", "email-include-link"];

function emailGetPrefs()
{
  var scriptProperties = PropertiesService.getScriptProperties();
  var props = {'email-enable':scriptProperties.getProperty('email-enable')};
  for (var i=0; i<email_string_params.length; ++i)
  {
    var propValue = scriptProperties.getProperty(email_string_params[i]);
    if (propValue != null)
    {
      props[email_string_params[i]] = propValue;
    }
  }
  return props;
}
