// In the Script Editor, run initialize() at least once to make your code execute on form submit
function initialize() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i in triggers) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger("main")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();
}

// Running the code in initialize() will cause this function to be triggered this on every Form Submit
// See here for a description of object e https://developers.google.com/apps-script/guides/triggers/events#form-submit
function main(e)
{
  // Uncomment this to debug by running main without form trigger
  
   if (typeof e === "undefined") {
     // e = {namedValues: {"Timestamp": ["Today"], "Email Address": ["esoroka@uci.edu"], "Vendor": ["Amazon"], "Vendor Website" : ["amazon.com"],"Lines 1 QTY": ["2"], "Line 1 Description" : ["Crappy hardware"], "Line 1 Item #": ["1234"], "Line 1 Price" : ["19.99"]}};
     e = {values: ["Today", "cansatuci@gmail.com", "Download more RAM", "","downloadmoreram.com","Ground","Emiko", "3109196950","","4","GB", "Downloadable RAM","1234","9.99","yes","4","GB", "Extra shiny high-speed downloadable RAM","5678","19.99","no","","","","",""]};
   }
   
  // newPO is the File object for the new PO
  var newPO = makeNewPO(e.values);
  addRecords(e.values);
  submitValuesToSlack(e, e.values[nameColumn] + ", your PO is at " + newPO.getUrl());
  emailPO(e.values[emailColumn], newPO);
}
