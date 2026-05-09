import pdfplumber
import glob

files = sorted(glob.glob("*.pdf"))
for f in files:
    print(f"\n================ {f} ================")
    with pdfplumber.open(f) as pdf:
        for idx, page in enumerate(pdf.pages):
            text = page.extract_text(layout=True)
            if "11.DEBIT DUTY" in text or "LICENCE DETAILS" in text:
                print(f"Page {idx} contains licence/debit!")
                # Print first 20 lines
                lines = text.split("\n")
                for l_idx, line in enumerate(lines):
                    if "DEBIT" in line or "LICENCE" in line or l_idx < 30:
                        print(f"  Line {l_idx:02d}: {line}")
