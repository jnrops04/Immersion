# Google Sheets Automation Project

A concise, one-sentence summary of what this automation script does to your spreadsheet data.

## Description

This Google Apps Script automates data workflows inside the attached Google Sheet. It eliminates manual data entry, formatting overhead, and human error.

### Key Features
* **Automation 1**: (e.g., Automatically formats raw data imports)
* **Automation 2**: (e.g., Sends automated email alerts based on cell criteria)
* **Automation 3**: (e.g., Fetches external data and populates specific rows)

## Project Structure

* `Code.js`: Core spreadsheet triggers (`onOpen`, `onEdit`) and processing logic.
* `UI.js`: Code managing custom menus, sidebars, or alert modals.
* `appsscript.json`: Manifest file specifying the required Google OAuth scopes.

## Setup Instructions

### 1. Bind Code to Spreadsheet
1. Open your Google Sheet.
2. Click **Extensions** > **Apps Script** in the top menu bar.
3. Delete any boilerplate code in the editor window.
4. Copy the contents of `Code.js` from this repo and paste it into the editor.
5. Create a new file named `UI.js` if applicable, and paste that code.
6. Click **Save** (floppy disk icon).

### 2. Initial Run & Authorization
1. Refresh your Google Sheet. Look for a new custom menu option in the toolbar (e.g., **Custom Utilities**).
2. Click the menu and select the main run function, or run it directly from the Apps Script editor.
3. An **Authorization Required** prompt will appear.
4. Click **Continue** > select your Google Account > click **Advanced** > click **Go to [Project Name] (unsafe)**.
5. Click **Allow** to grant permission.

## Triggers Configuration

This project relies on the following execution triggers:
* **Simple Triggers**: `onOpen()` automatically builds the custom sheet menu upon loading.
* **Installable Triggers**: (e.g., Needs a Time-driven trigger to run every night at 12:00 AM). 
  * *To set this up:* Click the **Triggers (Clock Icon)** on the left sidebar > **Add Trigger** > Configure settings > **Save**.

## License

This project is licensed under the MIT License.
