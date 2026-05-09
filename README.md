# Duty Call Automation

Automated Bill of Entry & Checklist Reconciliation tool built for Nagarkot Forwarders Pvt. Ltd.
It parses PDF documents to extract duty information and identify discrepancies.

> 📖 **Looking for instructions on how to use the application?** 
> Please refer to the [USER_GUIDE.md](./USER_GUIDE.md) for a comprehensive, step-by-step breakdown of the interface, workflows, and troubleshooting.

## Tech Stack
- Python 3.14
- Custom Tkinter GUI
- PDFPlumber for text extraction
- Pandas for data manipulation
- Google Sheets API for automated logging

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

---

## Notes
- **ALWAYS use virtual environment for Python.**
- Do not commit `venv/`.
- Run and test before pushing.
