/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e)
{
  SpreadsheetApp.getUi().createAddonMenu()
      .addItem('Start', 'showSidebar')
      .addToUi();
}

/**
 * Runs when the add-on is installed.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e)
{
  onOpen(e);
}

/**
 * Install the default preferences on the given spreadsheet
 * ss should be a Spreadsheet object, but it is not assumed to be active.
 */
function installDefaultPrefs(ss)
{
  SpreadsheetApp.setActiveSpreadsheet(ss);
  var defaults = {
   "slack-enable":"true",
   "slack-name":"New Purchase Order",
   "slack-icon":":money_with_wings:",
   "slack-color":"#0000DD",
   "slack-pretext":"A teammate submitted a new PO",
   "slack-fallback":"A teammate submitted a new PO",
   "email-enable":"true",
   "email-name":"PO System",
   "email-body":"Your PO is attached",
   "email-include-link":"false",
   "po-statuses":"In Progress,Complete,Canceled,Submitted",
   "po-start-status":"Submitted",
  }
  PropertiesService.getDocumentProperties().setProperties(defaults);
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 */
function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('mainSidebar')
      .evaluate();
      ui.setTitle('MAE Purchase Automation');
  SpreadsheetApp.getUi().showSidebar(ui);
}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

function includeTemplate(filename) {
  return HtmlService.createTemplateFromFile(filename)
      .evaluate().getContent();
}

/**
 * Displays an HTML-service dialog in Google Sheets that contains client-side
 * JavaScript code for the Google Picker API.
 */
function showPicker(name, type)
{
  var html = HtmlService.createTemplateFromFile('pickerDialog');
  html.paramName = name;
  html.pickerType = type;
  html = html.evaluate()
      .setWidth(600)
      .setHeight(425)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'Select a file ' + type);
}

/**
 * Gets the user's OAuth 2.0 access token so that it can be passed to Picker.
 * This technique keeps Picker from needing to show its own authorization
 * dialog, but is only possible if the OAuth scope that Picker needs is
 * available in Apps Script. In this case, the function includes an unused call
 * to a DriveApp method to ensure that Apps Script requests access to all files
 * in the user's Drive.
 *
 * @return {string} The user's OAuth 2.0 access token.
 */
function getOAuthToken() {
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();
}

/**
 * Save a document property (we need this to call setProperty from a callback in pickerDialog.html)
 */
function saveParam(name, param)
{
  PropertiesService.getDocumentProperties().setProperty(name, param);
}

function clearProperties()
{
  PropertiesService.getDocumentProperties().deleteAllProperties();
}

/**
 * We need this because for some reason you can't access properties from html files
 */
function getSetupPropertiesRunSetup(params)
{
  var props = PropertiesService.getDocumentProperties();
  params.poTemplateId = props.getProperty("file-po-template");
  params.folder = props.getProperty("folder-root");
  // clean up the document properties
  props.deleteAllProperties();
  // Check we have a PO template file
  if (params.poTemplateId == null) // there was no property! Halt
  {
    throw("Please select a PO template.");
  }
  if (params.folder == null) // no root folder;
  {
    throw("Please select a folder to install the PO system");
  }
  var poMaster = setupEverything(params);
  installDefaultPrefs(poMaster);
  return poMaster.getUrl();
}
  
/**
  * Expects: an object with:
  * [string] strings: a list of properties that are loaded into textfields
  * [string] checkboxes: a list of properties that are loaded into checkboxes
  * (Due to limitations of PropertiesService these will be "true" or "false")
  */
function getPrefs(prefs)
{
  var properties = PropertiesService.getDocumentProperties();
  var result = {
    strings:{},
    checkboxes:{}
    };
  result.strings = getPrefsHelper(properties, prefs.strings, result.strings);
  result.checkboxes = getPrefsHelper(properties, prefs.checkboxes, result.checkboxes);
  return result;
}
function getPrefsHelper(properties, prefsList, resultMap)
{
  len = prefsList.length;
  for (var i=0; i<len; ++i)
  {
    var value = properties.getProperty(prefsList[i]);
    if (value != null)
    {
      resultMap[prefsList[i]] = value;
    }
  }
  return resultMap;
}

function saveProps(props)
{
  PropertiesService.getDocumentProperties()
    .setProperties(props.strings)
    .setProperties(props.checkboxes);
}

/**
  * Terrible hack function to copy the "template" files
  * Will also save the file IDs
  */
function copyDefaultFiles()
{
  var templatePurchaseSheetId = "1oTFMIGbdADs5Hk3SA9V1nrZFBDrHjp9jTWLwI3jA-L4";
  var templateBlankPoId = "1ABemUrCoLAOrIYl-c0v2vDywPH2jqmDtAkpfuU0VdKw";
  var templateFormId = "106pVHeDXl-LHUKkr5brWyaBeZFKnasGwXArmUfkNv7Y";
  // Set up to access script properties
  var properties = PropertiesService.getDocumentProperties();
  // Get our system's root folder
  var PoSystemRoot = properties.getProperty('folder-root');
  if (PoSystemRoot == null) // the user has not set a root folder yet!
  {
    throw("Please set a root folder before running setup!");
  }
  PoSystemRoot = DriveApp.getFolderById(PoSystemRoot);
  
  // copySpreadsheet will return the file IDs
  //var purchaseSheetId = copySpreadsheet(templatePurchaseSheetId, "Purchase Master", PoSystemRoot).getId();
  //properties.setProperty('file-purchase-sheet', purchaseSheetId);
  //properties.setProperty('file-po-template', copySpreadsheet(templateBlankPoId, "PO Template", PoSystemRoot).getId());
  // Install the form trigger on the file ID
  //var purchaseSheetId = properties.getProperty('file-purchase-sheet');
  //installTrigger(purchaseSheetId);
  // Copy the form template
  var form = FormApp.openById(templateFormId);
  var formData = DriveApp.getFileById(form.getId());
  formData.makeCopy(formData.getName(), PoSystemRoot);
}

// Given the file ID of a template file, make a copy in the PO system root
// Fails if no PO system root set in scriptProperties
// Returns the Spreadsheet object for the new sheet
function copySpreadsheet(templateId, newName, destination)
{
  // Open the template spreadsheet and make a copy
  var ss = SpreadsheetApp.openById(templateId);
  SpreadsheetApp.setActiveSpreadsheet(ss);
  var ssCopy = ss.copy(newName);
  var newFile = DriveApp.getFileById(ssCopy.getId());
  // Now the copy is in the user root folder, and we want to move it
  var userRoot = DriveApp.getRootFolder();
  destination.addFile(newFile);
  userRoot.removeFile(newFile);
  return ssCopy;
}

// In the Script Editor, run installTrigger() at least once to make your code execute on form submit
// ss should be an open spreadsheet object
function installTrigger(ss) {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i in triggers) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger("purchaseForm")
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
}

// Running the code in initialize() will cause this function to be triggered this on every Form Submit
// See here for a description of object e https://developers.google.com/apps-script/guides/triggers/events#form-submit
function purchaseForm(e)
{
  var newPO = makeNewPO(e.values);
  addRecords(e.values);
  submitValuesToSlack(e, e.values[nameColumn] + ", your PO is at " + newPO.getUrl());
  emailPO(e.values[emailColumn], newPO);
}