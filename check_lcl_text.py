import pdfplumber
import glob

files = sorted(glob.glob("*.pdf"))
for f in files:
    with pdfplumber.open(f) as pdf:
        text = pdf.pages[0].extract_text()
        # count occurrences of LCL (excluding the header 2.LCL/)
        clean_text = text.replace("2.LCL/", "")
        if "LCL" in clean_text or " L " in clean_text:
            print(f"File {f} seems to have LCL in text!")
        else:
            print(f"File {f} is likely FCL only.")
