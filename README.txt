# Google Sheets Automation Project
Automates HR workflow for new Employees in a mid-level company

## Description          
This Google Apps Script automates data workflows inside the attached Google Sheet. It eliminates manual data entry, formatting overhead, and human error.

### Key Features
* **Automation 1**: (e.g., Automatically formats raw data imports)
* **Automation 2**: (e.g., Sends automated email alerts based on cell criteria)
* **Auto mation 3**: (e.g., Fetches external data and populates specific rows)

## Project Structure
* `immersion.js`: Core spreadsheet triggers (`onOpen`, `onEdit`) and processing logic.

## Setup Instructions
### Bind Code to Spreadsheet
1. Open your Google Sheet by creating a new empty google sheet.
2. Click **Extensions** > **Apps Script** in the top menu bar.
3. Delete any boilerplate code in the editor window.
4. Copy the contents of `immersion.js` from this repo (https://github.com/jnrops04/Immersion/blob/main/src/immersion.js) and paste it into the editor.
5. Click **Save** (floppy disk icon) or Ctrl+S.
6. Select "setup" function then RUN the selected function.
 -A popup will appear for authorizaion. Click review permission.
 -Select advanced. Click to <untitled project> (unsafe).
 -Click select all checkbox. Then continue.
 -Wait till execution status is complete.
7. Go back to the google sheet. Changes should now be applied by the code.
8. Start adding entries for new employee by providing Name, Id, and Email.
9. Go to your drive to view the generated google sheets once available (Task Folder and Employee Onboarding Logs)

## Process workflow
1.Upon running setup it will create a google sheet "Employee Onboarding Logs" and also adjust the format for the default google sheet.
2.From the google sheet create ENTER the Name, ID, and Email for the new employee.
3.Data will then be validated to check if data is not empty, no duplicates, and in correct format.
4.Welcome email containing tasks will be sent to the employee email.
5.A folder will then be created to contain google sheet task of each valid employee data. Entries will also be added to the task log tab of Logs sheet.
6.Once the employee accessed the task sheet and checked a checkbox, it will update the task log of status of that employee (will take 1-2 minutes interval before the status updates.)


## Triggers Configuration
This project relies on the following execution triggers:
* **Simple Triggers**: `onOpen()` automatically builds the custom sheet menu upon loading.
* **Installable Triggers**: (e.g., Needs a Time-driven trigger to run every night at 12:00 AM). 
  * *To set this up:* Click the **Triggers (Clock Icon)** on the left sidebar > **Add Trigger** > Configure settings > **Save**.


## Limitations for Dev setup (results to manual copy paste)
1. Can't properly setup for devs due to npm command not running due to missing node.js which can't be installed on the pc. Resulting to not access "clasp"
2. Can't use google appscript browser extension since it is only limited to chrome browser.

## Limitations for Automation
1. Can only have status for sent email. Does not provide status if the recipient received the email.