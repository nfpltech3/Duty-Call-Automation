"""Dump page 0 for BL/HBL and header section."""
import pdfplumber

f = "Import CheckList-IR5459826-27-07-MAY-2026_10_49_AM.pdf"
with pdfplumber.open(f) as pdf:
    ltxt = pdf.pages[0].extract_text(layout=True)
    for ln, line in enumerate(ltxt.split('\n')):
        print(f"[{ln:03d}] {line}")
