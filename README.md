# Google-Sheet-Click-Up-Sync
A Google Apps Script for robust two-way synchronization between a Google Sheet and a Click Up list—handling task creation, updates, custom fields, assignees, tags, and conflict resolution.
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)]

## Features

- 🚀 **Two-way synchronization**  
  - Google Sheets → ClickUp  
  - ClickUp → Google Sheets  
- 🔄 **Automatic detection** of task additions, deletions, and updates  
- ⚔️ **Conflict detection** via snapshot comparison, with safe skipping on conflict  
- 🏷️ **Assignees & tags synchronization**  
- 🔢 **Custom fields support** for any number of fields (example uses 10)  
- 🗑️ **Automatic deletion** of tasks when the title is cleared  
- 🔒 **Script locking** (LockService) to prevent concurrent-run conflicts

- ## Installation

1. **Attach the script to your Google Sheet**  
   - Open the target Google Sheet in your browser.  
   - From the menu, select **Extensions → Apps Script**. This opens the Apps Script IDE bound to that sheet.

2. **Obtain your Sheet ID**  
   - Look at the URL of your sheet:  
     ```
     https://docs.google.com/spreadsheets/d/<SHEET_ID>/edit#gid=0
     ```  
   - Copy the alphanumeric string between `/d/` and `/edit`—that’s your **SHEET_ID**.

3. **Obtain your ClickUp List ID**  
   - In ClickUp, open the List you want to sync.  
   - The URL looks like:  
     ```
     https://app.clickup.com/<WORKSPACE_ID>/v/l/<LIST_ID>
     ```  
   - Copy the `<LIST_ID>` portion—this is your **LIST_ID**.

4. **Generate your ClickUp API Token**  
   - In ClickUp, click your avatar → **Apps** (or **My Settings → Apps**).  
   - Under “Personal API Tokens,” click **Generate New Token**.  
   - Give it a name, then copy the resulting string—this is your **API_TOKEN**.

5. **Configure the script constants**  
   In your Apps Script project’s `Code.gs`, at the top replace the placeholders:
   ```js
   const SHEET_ID = 'YOUR_SHEET_ID';    // from step 2
   const LIST_ID  = 'YOUR_LIST_ID';     // from step 3
   const TOKEN    = 'YOUR_API_TOKEN';   // from step 4
   
6. **Save and Authorize**  

  - Click **Save** in the Apps Script editor.  
  - Run the `syncAll()` function once.  
  - When prompted, grant the script permission to:
     - Access your Google Spreadsheet  
     - Make external HTTP requests  

7. **Set up an Automatic Trigger**

  - In the Apps Script IDE, open **Triggers** (clock icon).  
  - Click **Add Trigger** and configure:  
     - **Function**: `syncAll`  
     - **Event source**: Time-driven  
     - **Type**: Minutes timer  
     - **Interval**: Every 5 minutes (or your preferred cadence)  

Your Google Sheet and ClickUp List are now linked—tasks will flow both ways automatically!

Event source: Time-driven

Type: Minutes timer

Interval: Every 5 minutes (or your preferred cadence)

Your Google Sheet and ClickUp List are now linked—tasks will flow both ways automatically!
