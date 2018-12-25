// This Google Sheets script will post to a slack channel when a user submits data to a Google Forms Spreadsheet

// Source: https://github.com/markfguerra/google-forms-to-slack
// The script has been modified - Emiko Soroka, 07/14/2018

/////////////////////////
// Begin customization //
/////////////////////////

// Alter this to match the incoming webhook url provided by Slack
// REPLACE THIS WITH YOUR WEBHOOK URL
var slackIncomingWebhookUrl = '';

// Include # for public channels, omit it for private channels
// REPLACE THIS WITH YOUR SLACK CHANNEL (example: var postChannel = "#purchasing")
var postChannel = "";

var postIcon = ":money_with_wings:";
var postUser = "New Purchase Order";
var postColor = "#0000DD";

var messageFallback = "A teammate submitted a new PO";
var messagePretext = "A teammate submitted a new PO.";

var lineStartColumn = 9; // Zero-indexed column where the line items begin (depends on purchasing sheet format)
var emailColumn = 1; // Zero-indexed column where the submitter's email address is (depends on purchasing sheet format)
var vendorColumn = 2; // Zero-indexed column where the PO vendor is (depends on purchasing sheet format)
var nameColumn = 6; // Zero-indexed position of submitter name (depends on purchasing sheet format)
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

function testSlack()
{
  submitValuesToSlack(undefined, "custom message");
}

// See here for a description of object e https://developers.google.com/apps-script/guides/triggers/events#form-submit
function submitValuesToSlack(e, message) {
  // Test code. uncomment to debug in Google Script editor
   if (typeof e === "undefined") {
     // e = {namedValues: {"Timestamp": ["Today"], "Email Address": ["esoroka@uci.edu"], "Vendor": ["Amazon"], "Vendor Website" : ["amazon.com"],"Lines 1 QTY": ["2"], "Line 1 Description" : ["Crappy hardware"], "Line 1 Item #": ["1234"], "Line 1 Price" : ["19.99"]}};
     e = {values: ["Today", "esoroka@uci.edu", "Download more RAM", "","downloadmoreram.com","Ground","Emiko", "3109196950","","4","GB", "Downloadable RAM","1234","9.99","yes","4","GB", "Extra shiny high-speed downloadable RAM","5678","19.99","no","","","","",""]};
   }
  if (typeof message === undefined)
  {
    message = messagePretext;
  }
  
  var attachments = constructAttachments(e.values, message);
  
  var payload = {
    "channel": postChannel,
    "username": postUser,
    "icon_emoji": postIcon,
    "link_names": 1,
    "attachments": attachments
  };

  var options = {
    'method': 'post',
    'payload': JSON.stringify(payload)
  };
  // Send the Slack message
  var response = UrlFetchApp.fetch(slackIncomingWebhookUrl, options);
}

// Creates Slack message attachments which contain the data from the Google Form
// submission, which is passed in as a parameter
// https://api.slack.com/docs/message-attachments
var constructAttachments = function(values, message) {
  //var fields = makeFields(values);
  var text = makeText(values);
  Logger.log(text);
  for (var i=0; i<values.length; ++i)
  {
    Logger.log(values[i]);
  }
  var attachments = [{
    "fallback" : messageFallback,
    "pretext" : message,
    "mrkdwn_in" : ["pretext"],
    "color" : postColor,
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
  var response = UrlFetchApp.fetch(slackIncomingWebhookUrl, options);
}