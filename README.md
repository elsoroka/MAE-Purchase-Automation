# MAE-Purchase-Automation
Script tools to automate purchase management for UCI Engineering senior design / research teams.

## Why do I need a purchase system?
This system reduces errors in filling out purchase orders and ensures all purchases are recorded in one place.

To make a purchase order with this system:
1. Teammates fill out a PO Google Form.
2. Once the form is submitted, the PO is automatically logged and announced in Slack. A copy of the PO is saved in Google Drive, and a PDF version is emailed to the submitter.
3. Your purchasing manager can review the PO from Slack.
4. Once the PO is approved, the email with the PDF can be easily forwarded to the department.

Benefits of this system:
1. All purchases are timestamped and logged in a central spreadsheet, making it easy to track inventory and project spending.
2. POs cannot be submitted with critical information (such as item numbers or vendor contact information) missing.
3. POs can be posted in Slack to ensure everyone is up-to-date with purchasing activities.

## How can I deploy this system for my team?

You can install the MAE Purchase Automation add-on from https://chrome.google.com/webstore/detail/mae-purchase-automation/fmaokimgoibneecchlmoooeicfpdjklj?authuser=0

If you are in the MAE department, copy a blank purchase order form from:
https://docs.google.com/spreadsheets/d/1JwkjrOoqqr90rCP3AxT5QVSTaxD0cHbC0vtgOLPkmuU/edit#gid=0
If you are in another department, you should provide your own blank purchase order as a Sheets file.

Make sure you have a copy of the blank form in your Google Drive.

Open a Google Sheets document (any will work). If the add-on is installed, you can find it in the Add-Ons menu.
Select "Start" to open the add-on.

You will be guided through set up. The add on will automatically generate a purchase tracking spreadsheet and purchase order form. You wll be provided links to these during the process.

Open the purchase tracking spreadsheet. You should be able to open the MAE Purchase Automation add-on, which will complete the automatic portion of the setup process.

1. Fill in parameters for your Slack workspace (or disable Slack integration entirely).
2. Fill in parameters for your purchase system emails.
3. You shouldn't need anything in Advanced.
4. Send a test Slack message and a test email to confirm your parameters are correct.
5. Fill out a test purchase order using the form. The following things should occur automatically:
    1. Your new purchase order should be posted in Slack (if you have the integration enabled).
    2. You should be emailed a copy of the purchase order as a PDF.
    3. The line items should appear in the "Records" sheet.


## I'm having a problem ##
Please open an issue or contact me directly. Make sure to provide any error messages or screenshots of the problem.

More documentation is coming soon!
