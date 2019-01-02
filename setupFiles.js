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

function testSetupParams()
{
  ss = SpreadsheetApp.openById("1oTFMIGbdADs5Hk3SA9V1nrZFBDrHjp9jTWLwI3jA-L4"),
  //setupParams(ss, ss.getSheetByName("Params"),12);
  setupRecords(ss.getSheetByName("Records"),["In Progress", "Canceled", "Complete"]); 
}

function setupRecords(sheet, statuses)
{
  // The columns the system requires to function correctly. More can be added after these.
  // Future versions may add the ability to reorder columns. YOU SHOULD NOT DO IT HERE...UNLESS YOU CHANGE THEM IN postToSlack.gs
  var columns = ["Timestamp","Submitted By","Vendor","Quantity","Unit of Measure","Item Description","Item Number","Unit Price","Ext. Price","Status"];
  sheet.getRange("A1:"+colNumToString(columns.length)+"1").setValues([columns]); // Since the sheet should be empty, this will become the first row.
  sheet.getRange("I2").setFormula("=H2*D2"); // Set up "ext price" column
  // Set data validation for "status" column
  if (statuses.length != 0)
  {
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(statuses).build();
    sheet.getRange("J2").setDataValidation(rule);
  }
  sheet.getRange("I2:J2").autoFill(sheet.getRange("I2:J100"),  SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
}

function setupParams(ss, sheet, nLineItems)
{
  var fields = ["Timestamp","EmailAddress","Vendor","VendorContactPhone","Website","Shipping","RequestorName","RequestorPhone","SpecialInstructions"];
  var itemFields = ["Quantity","UnitOfMeasure","Description","ItemNumber","Price"];
  var cells = [
  ["FormName","Form Responses 1"], // BEWARE OF HARDCODED VALUE
  ["NumberLineItems",nLineItems], // Number of line items supported by this system
  ["ColsPerLineItem",itemFields.length+1], // Add +1 because there is 1 more field: "Add another item?"
  ["POCount","=COUNT(INDIRECT(formName&\"!A:A\"),\"<>\")"], // Counts the number of POs submitted
  ];
  var len = fields.length; // Set up the named ranges for these fields
  var offset = cells.length;
  for (var i=0; i<len; ++i)
  {
    // Generate dynamic named ranges referring to the starting columns in the form response
    // These will update based on the NUMBER Of form responses
    // Example: =formName&"!$C$"&(1+POCount)
    cells[i+offset] = [fields[i],"=FormName&\"!$"+(colNumToString(1+i))+"$\"&(1+POCount)"];
  }
  var len = itemFields.length;
  var offset = cells.length;
  var name = sheet.getName();
  for (var i=0; i<len; ++i)
  {
    var letter = colNumToString(3+i); // Offset the column letter by two columns since it will go beside the named ranges
    // Generates formulas referring to ranges "="Params!$C$1:$C$"&LineItems"
    cells[i+offset] = [itemFields[i], "=\""+name+"!$"+letter+"$1:$"+letter+"$\"&NumberLineItems"];
  }
  // Set up named ranges with starting cell at $A$1
  installSingleCellNamedRanges(ss, sheet, cells, ss.getRange(sheet.getName()+"!$A$1"));
  
  // Now we need to generate the PO line item fields refererred to in the previous for-loop We will place these starting at $C$1
  var lineItemCells = [];
  var colStart = fields.length+1; // Start at this column to access line item fields
  cellRow = [];
  for (var i=0; i<itemFields.length; ++i)
  {
    // Generate a row of formulas referring to the line item's fields
    // =OFFSET(INDIRECT(FormName&"!J"&(1+POCount)),0,(ROW()-1)*ColsPerLineItem)
    cellRow[i] = "=OFFSET(INDIRECT(FormName&\"!"+colNumToString(colStart+i)+"\"&(1+POCount)),0,(ROW()-1)*ColsPerLineItem)";
  }
  var sourceRange = sheet.getRange("C1:"+colNumToString(2+itemFields.length)+"1"); // BEWARE OF HARDCODED VALUE
  sourceRange.setValues([cellRow]);
  // The range to fill with values.
  var destination = sheet.getRange("C1:"+colNumToString(2+itemFields.length)+(nLineItems+1)); // BEWARE OF HARDCODED VALUE
  // "Fill downward" so we can replicate the single row we constructed.
  sourceRange.autoFill(destination, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
}

function copyTemplateToMaster(templateID, destination, name)
{
  var source = SpreadsheetApp.openById(templateID);
  var sheet = source.getSheets()[0];
  sheet.copyTo(destination).setName(name);
}

/**
 * Convert a 1-indexed column to a string: 1->A, 2->B, 26->Z, 27->AA
 */
function colNumToString(num)
{
  var result = "";
  for(;num>0;num=Math.floor(num/27))
  {
    var tmp = (num-1)%26;
    result = String.fromCharCode(tmp + 65) + result;
  }
  return result;
}

/**
 * Given a reference to a spreadsheet (ss) and a sheet in the spreadsheet
 * and an Nx2 array of cells, which has the form [[string, string]]
 * and a startCell (which we treat as the "start position" of the Nx2 cell array
 * treat each inner pair [string, string] as [name, value]
 * and set up N named ranges so "name" refers to "value"'s cell
 * and copy the values in the cells array to the sheet at startCell's position
 */
function installSingleCellNamedRanges(ss, sheet, cells, startCell)
{
  // Set up named ranges
  var len = cells.length;
  var name = sheet.getName();
  for (var i=0; i<len; ++i)
  {
    ss.setNamedRange(cells[i][0], startCell.offset(i,1));
  }
  sheet.getRange(startCell.getRow(),startCell.getColumn(),cells.length, 2).setValues(cells);
}
