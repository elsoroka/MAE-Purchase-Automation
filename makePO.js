// Given the form response values
// Constructs a new PO with the name "PO_vendorname_timestamp"
// Returns the File object for this new PO
var makeNewPO = function(values)
{
  var properties = PropertiesService.getDocumentProperties();
  var targetFolder = DriveApp.getFolderById(properties.getProperty("folder-po"));
  var templateID = properties.getProperty("file-po-template");
  var poName = "PO_" + values[vendorColumn] +"_" + values[0]; // timestamp always first
  var newSpreadsheet = exportWithTemplate("CurrentPO", templateID, poName, targetFolder);
  return newSpreadsheet
}

var exportWithTemplate = function(sheetName, templateID, exportName, targetFolder)
{
  // Get the currentPO sheet
  var currentPO = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (currentPO == null) // the sheet doesn't exist!
  {
    throw("No CurrentPO sheet in Purchase Order spreadsheet!");
  }  
  // Open the template blank PO file and copy it to a new exportedPO
  var template = DriveApp.getFileById(templateID);
  var exportedPO = SpreadsheetApp.open(template.makeCopy(exportName,targetFolder));
  
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
var exportToPDF = function(spreadsheetFile, exportName, targetFolderID)
{ 
  var folder = DriveApp.getFolderById(targetFolderID);
  var sheet = spreadsheetFile.getBlob();
  sheet.setName(exportName);
  var newPO = folder.createFile(sheet);
  return newPO // Return the File object for the new PO PDF
}