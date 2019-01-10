/** The input params to setupEverything, makePoMasterSheet, etc. should be an object containing
 * string name: name to create the spreadsheet under
 * [string] statuses: list of purchase statuses
 * string startStatus: the status new POs start with
 * fileId poTemplateId: a blank PO spreadsheet.
 * integer nLineItems: the number of line items to support
 */
function setupEverything(params)
{
  var poMaster = SpreadsheetApp.openById("1oTFMIGbdADs5Hk3SA9V1nrZFBDrHjp9jTWLwI3jA-L4");
  var poForm = FormApp.openById("1HMqExnCDEdHJLmMj6TnH__iyK1hrHqETNrjWY6nn9WU");
  /*var poMaster = makePoMasterSheet(params);
  var poForm = makePoForm(params, poMaster.getId());
  // Install the "on form submit" trigger for poMaster
  installTrigger(poMaster);*/
  // Install default params on the new poMaster spreadsheet
  installDefaultPrefs(poMaster);
  // Save the new properties
  var openFolder = DriveApp.getFileById(params.folder);
  SpreadsheetApp.setActiveSpreadsheet(poMaster);
  PropertiesService.getDocumentProperties().setProperties({
    "po-start-status":params.startStatus,
    "file-purchase-sheet":poMaster.getId(), // Purchase spreadsheet
    "file-po-template":params.poTemplateId, // Template blank PO file
    "folder-root":params.folder, // Root folder
    "folder-po":params.folder, // Folder new POs will be saved to
    "system-enable":true, // Now the system is (mostly) ready!
    "folder-root-url":openFolder.getUrl(),
    "folder-po-url":openFolder.getUrl(),
    "file-purchase-sheet-url":poMaster.getUrl(),
    "file-po-template-url":SpreadsheetApp.openById(params.poTemplateId).getUrl(),
    });
  return {poSheet:poMaster,
          poForm:poForm};
}

/**
 * For a description of params, see setupEverything comments.
 */
function makePoMasterSheet(params)
{
  // Create the sheet; unfortunately it ends up in the user's root folder and we need to move it
  var poMaster = SpreadsheetApp.create(params.name);
  // Move the file to the right folder
  moveFromRootToFolder(poMaster.getId(), params.folder);
  
  // Build up the PO stuff - WORKING
  SpreadsheetApp.setActiveSpreadsheet(poMaster)
  poMaster.setActiveSheet(poMaster.getSheets()[0])
  // Generate the records and params sheets
  poMaster.renameActiveSheet("Records");
  setupRecords(poMaster.getActiveSheet(), params.statuses); // working
  setupParams(poMaster, poMaster.insertSheet("Params", 2), params.nLineItems); // working
  copyTemplateToMaster(params.poTemplateId, poMaster, "CurrentPO"); // working
  
  return poMaster;
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
  // The item order here must agree with makePoForm
  var fields = ["Timestamp","EmailAddress","RequestorName","RequestorPhone","Vendor","VendorContactPhone","Website","SpecialInstructions","Shipping"];
  var itemFields = ["Quantity","UnitOfMeasure","Description","ItemNumber","Price"];
  var cells = [
  ["FormName","Form Responses 1"], // BEWARE OF HARDCODED VALUE
  ["NumberLineItems",nLineItems], // Number of line items supported by this system
  ["ColsPerLineItem",itemFields.length+1], // Add +1 because there is 1 more field: "Add another item?"
  ["POCount","=COUNT(INDIRECT(formName&\"!A:A\"),\"<>\")"], // Counts the number of POs submitted
  ];
  var info =
"These cells are generated by the MAE Purchase System plugin.\nHowever, they are not used by the software.\n\
You may modify formulae and rearrange cells if necessary.\nIf your sheet becomes damaged, you can regenerate it with the plugin.\n\
To access a named range, use =INDIRECT(Name)\nThis is a dynamic range which is required to access the LATEST PO";
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
  var startCell = ss.getRange(sheet.getName()+"!$A$1");
  // Set up named ranges with starting cell at $A$1
  installSingleCellNamedRanges(ss, sheet, cells, startCell);
  // Add info cell and push values to sheet
  cells[cells.length] = [info, ""];
  sheet.getRange(startCell.getRow(),startCell.getColumn(),cells.length, 2).setValues(cells);
  
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

/**
 * Given the ID of a template, and an open destination spreadsheet,
 * copy the first sheet of the template to the destination and name the sheet correctly
 */
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
}

function moveFromRootToFolder(fileId, destinationId)
{
  var userRoot = DriveApp.getRootFolder();
  var file = DriveApp.getFileById(fileId); // File object instead of spreadsheet object
  DriveApp.getFolderById(destinationId).addFile(file);
  userRoot.removeFile(file);
}

/*        PROGRAMMATIC FORM GENERATION         */
/**
 * Programatically create the PO form with given # of line items params.nLineItems in params.folder
 * Link it to the spreadsheet at ID linkToId
 */
function makePoForm(params, linkToId)
{
  var poForm = FormApp.create("New Purchase Order");
  moveFromRootToFolder(poForm.getId(), params.folder);
  poForm.addSectionHeaderItem().setTitle("New Purchase Order");
  
  // Define the text items, and whether they are required (true) or optional.
  var requireNumber = FormApp.createTextValidation().requireNumber().build();
  var textItems = [
    {title:"Your Name", required:true, validation:null, help:""},
    {title:"Your Phone", required:true, validation:requireNumber, help:""},
    {title:"Vendor", required:true, validation:null, help:"You must prepare one PO for each vendor."},
    {title:"Vendor's Phone", required:false, validation:requireNumber, help:"Not all vendors have phone numbers."},
    {title:"Website", required:true, validation:null, help:""},
    {title:"Special Instructions", required:false, validation:null, help:"Coupon codes, shipping instructions, or other notes"},
    ];
  addTextItemsHelper(poForm, textItems, undefined); // don't enumerate these items
  
  // Add the multiple-choice shipping selector
  poForm.addMultipleChoiceItem().setTitle("Shipping")
    .setChoiceValues(["Ground", "2-Day", "3-Day", "Overnight"])
    .showOtherOption(true)
    .setHelpText("If you select Other, enter a date.");
    
  // Set up the N line items
  var lineHelpText = "For each line item, you need Quantity, Description, Item # and Price. Unit of Measure is optional.\n\
  You may enter up to " + params.nLineItems + " items.";
  var lineTextItems = [
    {title:" Quantity", required:true, validation:requireNumber, help:""},
    {title:" Unit of Measure", required:false, validation:null, help:"Unit the quantity is measured in; for example yards or pounds"},
    {title:" Description", required:true, validation:null, help:"Descriptive name of the item. \
If you need to choose a color, finish, or other parameter, do it here."},
    {title:" Item #", required:true, validation:null, help:"When searched on the supplier's website, this should bring up the item. \
For Amazon orders, this is the ASIN number on the item page."},
    {title:" Price", required:true, validation:requireNumber, help:""},
    ]; 
  
  // Generate the line item sections
  for (var i=1; i<=params.nLineItems; ++i)
  {
    poForm.addPageBreakItem().setTitle("Line " + i).setHelpText(lineHelpText);
    addTextItemsHelper(poForm, lineTextItems, i); // enumerate items per line (i)
    if (i != params.nLineItems) // Do not add this to the last line
    {
      var switchItem = poForm.addMultipleChoiceItem().setTitle("Add another line?");
      switchItem.setChoices([
        switchItem.createChoice("Yes", FormApp.PageNavigationType.CONTINUE),
        switchItem.createChoice("No", FormApp.PageNavigationType.SUBMIT),
        ]);
    }
  }
  // Now we have the beginning section and N line items. Make sure the settings are correct and link it to the sheet
  poForm.setCollectEmail(true)
    .setLimitOneResponsePerUser(false)
    .setDestination(FormApp.DestinationType.SPREADSHEET, linkToId);
  return poForm;
}

function addTextItemsHelper(form, textItems, number)
{
  var len = textItems.length;
  for (var i=0; i<len; ++i)
  {
    var item = form.addTextItem()
      .setTitle((number == undefined) ? textItems[i].title : "Line "+i+textItems[i].title) // Optionally enumerate items with provided number
      .setRequired(textItems[i].required)
      .setHelpText(textItems[i].help);
    if (textItems[i].validation != null)
    {
      item.setValidation(textItems[i].validation);
    }
  }
}