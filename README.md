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
Right now, you can copy the purchasing spreadsheets and form here:
https://drive.google.com/drive/folders/1o56mLraFZ4qXF73F_-e-m_xOMNhXZT64?usp=sharing

**You should complete these steps from a team account or shared Google account, if you have one.**

0. Copy the Purchase_Template spreadsheet. This is your purchase spreadsheet. Also copy the Purchase_Form and Blank_PO.

1. Open the Blank_PO and fill in the PI Name and Account Name for your team. Do the same for the CurrentPO sheet in the Purchase_Template.

2. Link the Purchase_Form and Purchase_Template following these instructions: https://support.google.com/docs/answer/2917686?hl=en

    1. Open the Purchase_Template spreadsheet. In the `Params` sheet, make sure the `Response Sheet Name` is correct. It's usually `Form Responses 1`. Delete the `Dummy Form Response (delete)` sheet.

3. You will need to set up the 5 script files so they are bound to your purchase spreadsheet. Open the Purchase_Template spreadsheet and click "Tools->Script Editor", then paste in the code. Read the comments carefully. There is some information missing you need to fill in.

**Specifically, you need:**

  4. The `targetFolderID` in `makePO.gs`. New POs will be saved to this folder.
  
What's a folder ID?

Given the URL of a folder in Drive (NOT A REAL ID:) https://drive.google.com/drive/u/0/folders/1o56mLraFZ4qXF73F_-e-m_xOMNhXZT65
The folder ID is the part after /u/0/folders/, so this (FAKE) one is: 1o56mLraFZ4qXF73F_-e-m_xOMNhXZT65

  5. The `templateID` in `makePO.gs`. This refers to a blank PO template. You should copy the Blank_PO spreadsheet and use this as your template.

**If you do not want new purchases announced in Slack,** do not complete steps 6, 7, or 8. Instead, locate the line in `main.gs` function `main()`, which calls `postPOToSlack()`, and comment it out.

  6. In `postToSlack.gs` change the `slackIncomingWebhookUrl` to your team's webhook URL.

  7. In `postToSlack.gs` change the `postChannel` to the Slack channel you want to post in.

  8. Optionally, change the `postIcon`, `postUser` and other cosmetic settings.


9. Once you have filled everything in, use the Script Editor to run initialize(). This will install the trigger that makes the script run when users fill out the purchasing form.
**THE SYSTEM WILL NOT WORK UNTIL YOU FINISH THIS STEP**

## I am having a problem.
There are some test functions built into the various files which individually make a Slack post, send an email, or process a fake PO. Try running these from the script editor.

If you still can't figure it out, let me know and I'll do my best to help.

## This is too much work. When is the user-friendly version coming out?
Soon!
