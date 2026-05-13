# Duty Call Automation

Automated Bill of Entry & Checklist Reconciliation tool built for Nagarkot Forwarders Pvt. Ltd.
It parses PDF documents to extract duty information, identifies discrepancies, and pushes validated data to Zoho Creator (Shakti DSR).

> 📖 **Looking for instructions on how to use the application?** 
> Please refer to the [USER_GUIDE.md](./USER_GUIDE.md) for a comprehensive, step-by-step breakdown of the interface, workflows, and troubleshooting.

## Tech Stack
- Python 3.14
- Custom Tkinter GUI
- PDFPlumber for text extraction
- Pandas for data manipulation
- Google Sheets API for automated logging
- Zoho Creator REST API v2 for Shakti DSR sync

---

## Installation

### Clone
```bash
git clone https://github.com/nfpltech3/Duty-Call-Automation.git
cd Duty-Call-Automation
```

---

## Python Setup (MANDATORY)

⚠️ **IMPORTANT:** You must use a virtual environment.

1. Create virtual environment
```bash
python -m venv venv
```

2. Activate (REQUIRED)

Windows:
```cmd
venv\Scripts\activate
```

Mac/Linux:
```bash
source venv/bin/activate
```

3. Install dependencies
```bash
pip install -r requirements.txt
```

4. Run application
```bash
python be_extractor.py
```

---

### Google Sheets Integration
To enable the background validation log uploader:
1. Create a Google Cloud Project with the Sheets and Drive APIs enabled.
2. Setup an OAuth Client ID (Desktop Application type).
3. Download the credentials and save them as `client_secret.json` in the same directory as `be_extractor.py`.
4. (Optional) For headless server operation, you can download a Service Account key instead and save it as `credentials.json`.
5. The Google Sheet must contain a tab explicitly named `"Validation Log"`.

---

### Zoho Shakti DSR Integration
To enable the "Push to Shakti" feature which syncs BE values (Assessable Value, Total Duty, Forgone Duty) to the Zoho Creator Client DSR form:

1. Create a `shakti_secret.json` file in the **same directory** as `be_extractor.py` (or alongside the built `.exe`):
```json
{
    "client_id": "YOUR_ZOHO_CLIENT_ID",
    "client_secret": "YOUR_ZOHO_CLIENT_SECRET",
    "refresh_token": "YOUR_ZOHO_REFRESH_TOKEN"
}
```
2. The Refresh Token must have the scope `ZohoCreator.report.READ,ZohoCreator.report.UPDATE`.
3. This file is **excluded from git** via `.gitignore` — never commit secrets.

---

### Build Executable (For Desktop Apps)

1. Install PyInstaller (Inside venv):
```bash
pip install pyinstaller
```

2. Build using the included Spec file (Ensure you do not run be_extractor.py directly):
```bash
pyinstaller DutyCallAutomation.spec
```

3. Locate Executable:
The application will be generated in the `dist/` folder.

4. Deploy:
Simply share the single `DutyCallAutomation.exe` file. Both `logo.png` and `shakti_secret.json` are bundled inside the executable automatically — no companion files needed.

---

## Notes
- **ALWAYS use virtual environment for Python.**
- Do not commit `venv/`, `dist/`, `build/`, or any secret JSON files.
- Run and test before pushing.
