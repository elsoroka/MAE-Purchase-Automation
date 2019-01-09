// This script fills in the purchase tracking sheet by appending to the top
// That way the most recent purchases are always easily visible
// Emiko Soroka, 07/14/2018

var recordsSheetName = "Records";

function addRecords(values)
{
  var theSheet = findSheetByName(recordsSheetName);
  if (theSheet==null) // we didn't find the records sheet!!!
  {
    return;
  }
  // Prepare data to write
  var timestamp = values[0]; // timestamp is always first
  var email = values[emailColumn];
  var vendor = values[vendorColumn];
  // for each line item in the PO form response
  for (var i=lineStartColumn; i<values.length; i=i+lineItemColumns)
  {
    if (values[i+descriptionOffset] == "")
    {
      break; // we've reached the end of the populated rows
    }
    var status = PropertiesService.getDocumentProperties().getProperty('po-start-status');
    var theRow = [timestamp, email, vendor, values[i+qtyOffset],
                  values[i+unitOffset],
                  values[i+descriptionOffset],
                  values[i+itemNumOffset],
                  values[i+priceOffset],"=H2*D2", status]; // extended price (quantity * price), then init status
    insertRow(theSheet, theRow, 2);
  }
}

// returns the sheet with the given name, or null if it didn't exist (that should never happen)
var findSheetByName = function(sheetName)
{
    var theSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    theSheets = theSpreadsheet.getSheets();
    for (var i=0; i<theSheets.length; ++i)
    {
      if (theSheets[i].getName() === sheetName)
      {
        return theSheets[i];
      }
    }
    Logger.log("Couldn't find the records sheet!");
    return null; // probably dumb, happens if you don't find the records sheet which shouldn't happen
}

// appends a row to the TOP of the sheet (optionally skipping the first row which may be a header)
// Code stolen from StackOverflow: https://stackoverflow.com/questions/28295056/google-apps-script-appendrow-to-the-top
function insertRow(sheet, rowData, optIndex) { // you can tell it's stolen because the javascript heathens didn't give this poor brace its own line
  var index = optIndex || 1;
  sheet.insertRowBefore(index).getRange(index, 1, 1, rowData.length).setValues([rowData]);
}