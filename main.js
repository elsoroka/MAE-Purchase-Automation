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
 * Show an alert from an HTML file
 */
function showAlert(msg)
{
  var ui = SpreadsheetApp.getUi();
  ui.alert(msg);
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
  SpreadsheetApp.getUi().showModalDialog(html, 'Select a file ');
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
 * Terrible hack implementing a way to pass messages back from the modalWindow to the sidebar,
 * or generally from server to client side, because the sidebar is not the parent of the modalWindow.
 */
function setMessage(element, newMsg)
{
  PropertiesService.getScriptProperties().setProperties({
    "msg":newMsg,
    "element":element,
  });
}
function getMessage()
{
  var props = PropertiesService.getScriptProperties();
  var tmp = props.getProperty("msg");
  var element = props.getProperty("element");
  if (tmp == null)
  {
    tmp = "";
  }
  props.setProperty("msg", "");
  return {
    msg:tmp,
    stop:(tmp != ""),
    element:element,
    };
}

/**
 * Save a document property (we need this to call setProperty from a callback in pickerDialog.html)
 */
function saveParam(name,param)
{
  PropertiesService.getDocumentProperties().setProperty(name, param);
}

function clearPropertiesWithModalConfirm()
{
  // Display a dialog box with a message and "Yes" and "No" buttons. The user can also close the
  // dialog by clicking the close button in its title bar.
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('This will delete all configuration for this document. The system will NOT work. Continue?', ui.ButtonSet.YES_NO);
  // Process the user's response.
  if (response == ui.Button.YES)
  {
    PropertiesService.getDocumentProperties().deleteAllProperties();
    removeTriggers(SpreadsheetApp.getActiveSpreadsheet());
  }
  return (response == ui.button.YES) ? true : false; // true: delete occurred. false: delete was canceled.
}

/**
 * We need this because for some reason you can't access properties from html files
 */
function getSetupPropertiesRunSetup(params)
{
  var props = PropertiesService.getDocumentProperties();
  params.poTemplateId = props.getProperty("file-po-template");
  params.folder = props.getProperty("folder-root");
  // clean up the document properties since this is NOT a working system and may be some random file
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

// In the Script Editor, run installTrigger() at least once to make your code execute on form submit
// ss should be an open spreadsheet object
function installTrigger()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //removeTriggers(ss);
  ScriptApp.newTrigger("purchaseForm") // Run purchaseForm function (taking form response as parameter) on form submit
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
  var userEmail = Session.getActiveUser().getEmail();
  PropertiesService.getDocumentProperties().setProperty("user",userEmail);
  return userEmail;
}

function removeTriggers()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var triggers = ScriptApp.getUserTriggers(ss);
  for (var i in triggers) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  var props = PropertiesService.getDocumentProperties().deleteProperty("user");
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