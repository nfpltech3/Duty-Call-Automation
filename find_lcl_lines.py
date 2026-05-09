import pdfplumber
import glob

files = sorted(glob.glob("*.pdf"))
for f in files:
    print(f"\n================ {f} ================")
    with pdfplumber.open(f) as pdf:
        text = pdf.pages[0].extract_text()
        lines = text.split("\n")
        for idx, line in enumerate(lines):
            if "LCL" in line:
                print(f"Line {idx}: {line}")
