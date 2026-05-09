"""Dump checklist PDF text to understand its structure."""
import pdfplumber

f = "Import CheckList-IR5459826-27-07-MAY-2026_10_49_AM.pdf"
with pdfplumber.open(f) as pdf:
    print(f"Total pages: {len(pdf.pages)}")
    for i, page in enumerate(pdf.pages):
        print(f"\n--- PAGE {i} (plain text) ---")
        txt = page.extract_text()
        if txt:
            print(txt[:2000])
        
        print(f"\n--- PAGE {i} (layout text) ---")
        ltxt = page.extract_text(layout=True)
        if ltxt:
            print(ltxt[:2000])
