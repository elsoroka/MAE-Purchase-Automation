// This is the folder we want to save the POs to
// ENTER A FOLDER ID HERE - this is the long string of numbers when you load that folder in Google Drive
// The POs will be saved to this folder.
var targetFolderID = "";


// ID of template PO file which is a blank PO with formatting in place
// ENTER A TEMPLATE ID HERE - this is the ID for a template PO, which has the correct formatting in place
var templateID = ""
// Name of current PO in this master Purchase Tracking spreadsheet
var poSheetName = "CurrentPO";

function testExport()
{
  var newPO = exportWithTemplate(poSheetName, templateID, "ExportTest2");
  exportToPDF(newPO, "ExportPDFTest2")
}

// Given the form response values
// Constructs a new PO with the name "PO_vendorname_timestamp"
// Returns the File object for this new PO
var makeNewPO = function(values)
{
  var poName = "PO_" + values[vendorColumn] +"_" + values[0]; // timestamp always first
  var newSpreadsheet = exportWithTemplate(poSheetName, templateID, poName);
  return newSpreadsheet
}
// Deprecated in favor of exportWithTemplate
/*
// Export the currentPO sheet to a new Google Sheets document
var exportToSheet = function(sheetName, exportName)
{
  var currentPO = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var folder = DriveApp.getFolderById(targetFolderID);
  var blankFile = DriveApp.getFileById(blankSheetID);
  var exportedPO = SpreadsheetApp.open(blankFile.makeCopy(exportName,folder));
  // Copy the sheet, this brings over the formatting but references will be broken
  currentPO.copyTo(exportedPO);
  // Do some cleanup of the new PO
  // There is an empty sheet "Sheet1" in the new file; we want to get rid of it
  exportedPO.deleteSheet(exportedPO.getSheetByName("Sheet1"));
  // Get the VALUES in the CurrentPO sheet
  var sourceValues  = currentPO.getDataRange().getValues();
  // Copy VALUES to the new PO (overwriting broken references)
  exportedPO.getSheets()[0].getRange(1,1,sourceValues.length, sourceValues[0].length).setValues(sourceValues);
  return exportedPO; // Return file object
}
*/
var exportWithTemplate = function(sheetName, templateID, exportName)
{
  // Get the currentPO sheet
  var currentPO = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  // Open the folder we want to save the POs to
  var folder = DriveApp.getFolderById(targetFolderID);
  
  // Open the template blank PO file and copy it to a new exportedPO
  var template = DriveApp.getFileById(templateID);
  var exportedPO = SpreadsheetApp.open(template.makeCopy(exportName,folder));
  
  // Get the VALUES in the CurrentPO sheet
  var sourceValues  = currentPO.getDataRange().getValues();
  // Copy VALUES to the new PO
  var range = exportedPO.getSheetByName("Sheet1").getRange(1,1,sourceValues.length, sourceValues[0].length)
  range.setValues(sourceValues);
  
  // Don't remove this line. If you do, the next call to emailPO will generate a PDF without the updated values
  var newValues = range.getValues();
  
  return exportedPO; // Return file object
}

// This function isn't currently used
// Export the spreadsheet as PDF and save it in "Generated POs" folder
var exportToPDF = function(spreadsheetFile, exportName)
{ 
  var folder = DriveApp.getFolderById(targetFolderID);
  var sheet = spreadsheetFile.getBlob();
  sheet.setName(exportName);
  var newPO = folder.createFile(sheet);
  return newPO // Return the File object for the new PO PDF
}