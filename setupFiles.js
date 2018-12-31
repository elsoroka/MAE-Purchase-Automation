// The input params to makePoMasterSheet should be an object containing
// string name: name to create the spreadsheet under
// [string] statuses: list of purchase statuses
// string startStatus: the status new POs start with
// fileId poTemplateId: a blank PO spreadsheet.
function makePoMasterSheet(params)
{
  /* WORKING
  // Create the sheet; unfortunately it ends up in the user's root folder and we need to move it
  var poMaster = SpreadsheetApp.create(params.name);
  // Move the file to the right folder
  var userRoot = DriveApp.getRootFolder();
  poMasterFile = DriveApp.getFileById(poMaster.getId()); // File object instead of spreadsheet object
  DriveApp.getFolderById(params.folder).addFile(poMasterFile);
  userRoot.removeFile(poMasterFile);
  */
  var poMaster = SpreadsheetApp.openById("1--XRM9OEyvoljVFIEYUjRYnkPY7PZxviTs5ygcEMlQ4");
  // Build up the PO stuff - WORKING
  /*SpreadsheetApp.setActiveSpreadsheet(poMaster);
  poMaster.setActiveSheet(poMaster.getSheets()[0])
  var recordsSheet = poMaster.renameActiveSheet("Records");
  copyTemplateToMaster(params.poTemplateId, poMaster, "CurrentPO");
  var paramsSheet = poMaster.insertSheet("Params", 2);*/
  
  // Save the new properties
  PropertiesService.getScriptProperties().setProperties({
    "po-start-status":params.startStatus,
    "po-statuses":params.statuses,
    "file-purchase-sheet":poMaster.getId(),
    "file-po-template":params.poTemplateId,
    });
}

function testMakeSheet()
{
  var params = {
    name:"PO Master",
    folder:"1lyeocV7wwwygnJqdJnu3v2pjPxuDzWZm",
    statuses:["In Progress", "Canceled", "Complete"],
    startStatus:"In Progress",
    poTemplateId:"1xiMY3cm6O7rEUCZfH7XZNpuOdsu6NjPzm6xfzdOkVLQ",
    };
    makePoMasterSheet(params);
}

function setupRecords(sheet)
{
  // The columns the system requires to function correctly. More can be added after these.
  // Future versions may add the ability to reorder columns. YOU SHOULD NOT DO IT HERE...UNLESS YOU CHANGE THEM IN postToSlack.gs
  var columns = ["Timestamp","Submitted By","Vendor","Quantity","Unit of Measure","Item Description","Item Number","Unit Price","Ext. Price","Status"];
  sheet.appendRow(columns); // Since the sheet should be empty, this will become the first row.
  sheet.getRange("I2").setFormula("=H2*D2"); // Set up "ext price" column
}

function copyTemplateToMaster(templateID, destination, name)
{
  var source = SpreadsheetApp.openById(templateID);
  var sheet = source.getSheets()[0];
  sheet.copyTo(destination).setName(name);
}

