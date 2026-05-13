# Duty Call Automation User Guide

## Introduction
The Duty Call Automation tool is a robust desktop application designed for Nagarkot Forwarders Pvt. Ltd. It automatically parses Bill of Entry (BE) PDFs and Checklist PDFs, reconciles the financial data between them, pushes validated values to Zoho Shakti DSR, and flags critical discrepancies before you file the data.

## How to Use

### 1. Launching the App
Simply double-click the `DutyCallAutomation.exe` file. The application will open in a full-screen, branded window. No installation or command-line experience is required.

### 2. The Workflow (Step-by-Step)
1. **Select BE PDF(s)**: Click this button and choose one or more **Bill of Entry** PDF files. The tool will parse them and build the foundational data tables (organized by branch: CSN, CLC, Pune).
2. **Select Checklists**: Click this button to upload the corresponding **Checklist** PDFs.
   - *Note: The application will automatically attempt to match checklists to BEs using Job Numbers, BE Numbers, or HAWB/BL numbers.*
3. **Review Data**: Navigate through the branch tabs at the bottom. The "Validation" column will instantly show whether the BE and Checklist values match (OK, WARNING, or HIGH RISK).
4. **Edit Missing Data (Optional)**: If a cell is blank or incorrect, double-click the cell directly in the table to manually type the correct value. Press Enter to save.
5. **Push to Shakti**: Once you have reviewed the data, click **Push to Shakti** to sync the BE values (Assessable Value, Total Duty, Forgone Duty) to the Zoho Creator Client DSR form.
   - *Note: The Push button is only enabled when at least one row has both a valid BE Number and Job Number.*
   - *Constraint: Push is disabled for the "General" and "Validation Report" tabs. You must select a Skoda branch tab (CSN, CLC, Pune).*
   - *Inbond BEs are automatically skipped during push.*
6. **Copy Table**: After a successful push, the **Copy Table** button unlocks. Click it to copy the active branch tab's data to your clipboard in an Excel-ready HTML format.
   - *Constraint: You must push to Shakti before copying. This ensures data is synced to the cloud first.*
   - *Constraint: You cannot copy from the "General" tab.*

## Interface Reference

| Control / Input | Description | Expected Format |
| :--- | :--- | :--- |
| **Select BE PDF(s)** | Uploads the primary Bill of Entry documents. | Multiple `.pdf` files |
| **Select Checklists** | Uploads the secondary Checklist documents for reconciliation. | Multiple `.pdf` files |
| **Push to Shakti** | Syncs BE values (Ass. Value, Total Duty, Forgone Duty) to the Zoho Client DSR. Enabled only after checklists are loaded and valid data exists. | N/A |
| **Copy Table** | Copies the active branch tab's data for Excel pasting. Enabled only after a successful Shakti push. | N/A |
| **Clear All** | Resets the application entirely, clearing all uploaded documents. | N/A |
| **Branch Tabs (CSN, CLC, etc.)** | Displays the parsed data specific to that branch. | N/A |
| **Validation Report Tab** | A comprehensive summary of all discrepancies and "HIGH RISK" matches. | N/A |

## Troubleshooting & Validations

If you see an error or warning, check this table:

| Message | What it means | Solution |
| :--- | :--- | :--- |
| `HIGH RISK` | An Assessable Value, Duty, or Foregone Duty mismatch exceeded ₹500. | Double-check the source PDFs. You will be prompted to explicitly confirm before copying the table. |
| `No match` | A BE was uploaded, but no matching Checklist was found. | Upload the missing Checklist, or double-click the row's "Job No" cell to type it manually. |
| `The following BE rows are missing a matching Checklist` | You attempted to click "Copy Table" but there are rows with a blank Job No. | **Hard Block**: You cannot export. You must resolve the missing checklists first. |
| `Missing Data: Cannot copy table` | A mandatory field (like Job No, Supplier, BL No, etc.) is blank. | **Hard Block**: Double-click the specified blank cells to manually fill them in. |
| `Switch to a branch tab` | You clicked "Copy Table" while looking at the "General" or "Validation Report" tab. | Click on a specific branch tab (e.g., CSN, CLC) at the bottom, then click Copy again. |
| `Push Successful` | All records were successfully synced to Zoho Shakti DSR. | No action needed. The Copy Table button is now unlocked. |
| `Push Finished with Errors` | Some records pushed successfully, but others failed. Details are shown in the popup. | Review the listed errors. Common causes: Job Number not found in Zoho, or the DSR record is locked. Copy Table is still unlocked for successfully pushed records. |
| `Push Failed: Critical error` | The entire push operation failed. Zero records were synced. | Check your internet connection. If persistent, verify that `shakti_secret.json` credentials are valid and not expired. |
| `Action Blocked` | You tried to push from the "General" or "Validation Report" tab. | Switch to a Skoda branch tab (CSN, CLC, or Pune) first. |
| `Inbond Records Skipped` | Inbond-type BEs were detected and excluded from the push. | This is expected behavior. Inbond BEs are not pushed to the Client DSR. |
| `Google Sheets Upload Failed: HTML Error Page` | The Webhook sync failed due to Google Apps Script permissions. | Verify that the Google Apps Script deployment's "Who has access" is set to "Anyone". |
| `Closing... Waiting for Google Sheets sync` | You clicked the 'X' to close the app while data was syncing to the cloud. | Do nothing. The app will close itself automatically within a few seconds once the sync is safely completed. |
