/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
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
function onInstall(e) {
  installDefaultPrefs();
  onOpen(e);
}

function installDefaultPrefs()
{
  var defaults = {"slack-name":"New Purchase Order",
   "slack-icon":":money_with_wings:",
   "slack-color":"#0000DD",
   "slack-pretext":"A teammate submitted a new PO",
   "slack-fallback":"A teammate submitted a new PO",
  }
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperties(defaults);
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 */
function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('mainSidebar')
      .evaluate();
      ui.setTitle('MAE-Purchase-Automation');
  SpreadsheetApp.getUi().showSidebar(ui);
}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}