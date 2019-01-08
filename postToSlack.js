// This Google Sheets script will post to a slack channel when a user submits data to a Google Forms Spreadsheet

// Source: https://github.com/markfguerra/google-forms-to-slack
// The script has been modified - Emiko Soroka, 07/14/2018

/////////////////////////
// Begin customization //
/////////////////////////

var lineStartColumn = 9; // Zero-indexed column where the line items begin (depends on purchasing sheet format)
var emailColumn = 1; // Zero-indexed column where the submitter's email address is (depends on purchasing sheet format)
var vendorColumn = 4; // Zero-indexed column where the PO vendor is (depends on purchasing sheet format)
var nameColumn = 2; // Zero-indexed position of submitter name (depends on purchasing sheet format)
var lineItemColumns = 6; // Number of columns associated with each line item (depends on sheet format)
// Describe the position of each relevant field (quantity, units of measure, item description, item # and price) in a "line item"
var qtyOffset = 0; // position of Quantity in per-line columns
var unitOffset = 1;
var descriptionOffset = 2;
var itemNumOffset = 3;
var priceOffset = 4;

///////////////////////
// End customization //
///////////////////////

function slackTestPost(url, channel, name, icon, text)
{
  var payload = {
    "channel": channel,
    "username": name,
    "icon_emoji": icon,
    "link_names": 1,
    "text": text
  };

  var options = {
    'method': 'post',
    'payload': JSON.stringify(payload)
  };
  // Send the Slack message
  return response = UrlFetchApp.fetch(url, options);
}

// See here for a description of object e https://developers.google.com/apps-script/guides/triggers/events#form-submit
function submitValuesToSlack(e, message)
{
  var properties = PropertiesService.getDocumentProperties();
  if (properties.getProperty('slack-enable') == "false")
  {
    return;
  }
  var attachments = constructAttachments(e.values, message, properties);
  var payload = {
    "channel": properties.getProperty('slack-channel'),
    "username": properties.getProperty('slack-name'),
    "icon_emoji": properties.getProperty('slack-icon'),
    "link_names": 1,
    "attachments": attachments
  };
  var options = {
    'method': 'post',
    'payload': JSON.stringify(payload)
  };
  // Send the Slack message
  var response = UrlFetchApp.fetch(properties.getProperty('slack-webhook'), options);
}

// Creates Slack message attachments which contain the data from the Google Form
// submission, which is passed in as a parameter
// https://api.slack.com/docs/message-attachments
var constructAttachments = function(values, message, properties) {
  //var fields = makeFields(values);
  var text = makeText(values);
  Logger.log(text);
  for (var i=0; i<values.length; ++i)
  {
    Logger.log(values[i]);
  }
  var attachments = [{
    "fallback" : properties.getProperty('slack-fallback'),
    "pretext" : message,
    "mrkdwn_in" : [properties.getProperty('slack-pretext')],
    "color" : properties.getProperty('slack-color'),
    "text"  : text
    //"fields" : fields
  }]

  return attachments;
}

// Creates an array of Slack fields containing the questions and answers
var makeFields = function(values) {
  var fields = [];

  var columnNames = getColumnNames();

  for (var i = 0; i < columnNames.length; ++i) {
    var colName = columnNames[i];
    var val = values[i];
    if (val != "" && val != undefined)
    {
      fields.push(makeField(colName, val));
    }
  }

  return fields;
}

// Creates a Slack field for your message
// https://api.slack.com/docs/message-attachments#fields
var makeField = function(question, answer) {
  var field = {
    "title" : question,
    "value" : answer,
    "short" : true
  };
  return field;
}

// Extracts the column names from the first row of the spreadsheet
var getColumnNames = function() {
  var sheet = SpreadsheetApp.getActiveSheet();

  // Get the header row using A1 notation
  var headerRow = sheet.getRange("1:1");

  // Extract the values from it
  var headerRowValues = headerRow.getValues()[0];

  return headerRowValues;
}

var makeText = function(values)
{
  text = "*" + values[emailColumn] + " submitted a PO for " + values[vendorColumn]+ "*\n";
  for (var i = lineStartColumn; i<values.length; i=i+lineItemColumns)
  { // Item description of item
    // QTY 3 yards, item # 1234, $19.99 each
    if (values[i+descriptionOffset] != "") // check if there is an item here, because there aren't always 12 line items
    {
      text += "*" + values[i+descriptionOffset]; // put the item description in BOLD
      text += "*\n" // put the quantity and price on the next line
           + "QTY " + values[i+qtyOffset];
      if (values[i+unitOffset] != "") // if there's a unit (such as "each" or "yards")
      {
        text += " " + values[i+unitOffset]; // add the unit
      }
      text += ", item # " + values[i+itemNumOffset] // add the item number
           + ", $" + values[i+priceOffset] + " each.\n"; // and the price
    }
  }
  return text;
}


function sendPlainMessage(message)
{
    var payload = {
    "channel": postChannel,
    "username": postUser,
    "icon_emoji": postIcon,
    "link_names": 1,
    "text": message
  };

  var options = {
    'method': 'post',
    'payload': JSON.stringify(payload)
  };
  // Send the Slack message
  var response = UrlFetchApp.fetch(PropertiesService.getDocumentProperties().getProperty('slack-webhook'), options);
}