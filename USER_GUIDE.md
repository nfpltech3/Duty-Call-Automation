# Duty Call Automation User Guide

## Introduction
The Duty Call Automation tool is a robust desktop application designed for Nagarkot Forwarders Pvt. Ltd. It automatically parses Bill of Entry (BE) PDFs and Checklist PDFs, reconciles the financial data between them, and flags critical discrepancies (like massive duty differences) before you file the data. 

## How to Use

### 1. Launching the App
Simply double-click the `DutyCallAutomation.exe` file located in your `/dist` folder. The application will open in a full-screen, branded window. No installation or command-line experience is required.

### 2. The Workflow (Step-by-Step)
1. **Select BE PDF(s)**: Click this button and choose one or more **Bill of Entry** PDF files. The tool will parse them and build the foundational data tables (organized by branch: CSN, CLC, Pune).
2. **Select Checklists**: Click this button to upload the corresponding **Checklist** PDFs.
   - *Note: The application will automatically attempt to match checklists to BEs using Job Numbers, BE Numbers, or HAWB/BL numbers.*
3. **Review Data**: Navigate through the branch tabs at the bottom. The "Validation" column will instantly show whether the BE and Checklist values match (OK, WARNING, or HIGH RISK).
4. **Edit Missing Data (Optional)**: If a cell is blank or incorrect, double-click the cell directly in the table to manually type the correct value. Press Enter to save.
5. **Copy Table**: Once all "HIGH RISK" rows have been addressed, click **Copy Table**. 
   - *Note: The table will be copied to your clipboard in an Excel-ready HTML format. You can paste it directly into Excel.*
   - *Constraint: You cannot copy from the "General" tab. You must select a specific branch tab (CSN, CLC, Pune) first.*

## Interface Reference

| Control / Input | Description | Expected Format |
| :--- | :--- | :--- |
| **Select BE PDF(s)** | Uploads the primary Bill of Entry documents. | Multiple `.pdf` files |
| **Select Checklists** | Uploads the secondary Checklist documents for reconciliation. | Multiple `.pdf` files |
| **Copy Table** | Copies the active branch tab's data for Excel pasting. | N/A |
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
| `Google Sheets Upload Failed: HTML Error Page` | The Webhook sync failed due to Google Apps Script permissions. | Verify that the Google Apps Script deployment's "Who has access" is set to "Anyone". |
| `Closing... Waiting for Google Sheets sync` | You clicked the 'X' to close the app while data was syncing to the cloud. | Do nothing. The app will close itself automatically within a few seconds once the sync is safely completed. |
