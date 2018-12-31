// YOU NEED TO ENTER AN EMAIL ADDRESS AND THE ID OF A GOOGLE SHEETS FILE TO USE THIS TEST FUNCTION
function emailTest()
{
  emailPO("YOUR EMAIL ADDRESS HERE", DriveApp.getFileById("ID OF TEST FILE"));
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
  }
  MailApp.sendEmail(recipient,
    ssName,
    properties.getProperty('email-body') + (properties.getProperty('email-include-link') == "true") ? ssUrl : "",
    options);
}

var email_string_params = ["email-name", "email-body"];


function emailGetPrefs()
{
  var scriptProperties = PropertiesService.getScriptProperties();
  var enabled = scriptProperties.getProperty('email-enable');
  if (enabled == null)
  {
    throw('No email preferences saved!');
  }
  var props = {'email-enable':enabled};
  for (var i=0; i<email_string_params.length; ++i)
  {
    props[email_string_params[i]] = scriptProperties.getProperty(email_string_params[i]);
  }
  return props;
}
