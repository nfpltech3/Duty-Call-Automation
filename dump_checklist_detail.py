"""Dump pages 0-2 of checklist in detail - focus on duty/assessable/BL/branch/foregone."""
import pdfplumber

f = "Import CheckList-IR5459826-27-07-MAY-2026_10_49_AM.pdf"
with pdfplumber.open(f) as pdf:
    for i in range(min(3, len(pdf.pages))):
        print(f"\n{'='*80}\n--- PAGE {i} (layout text) ---\n{'='*80}")
        ltxt = pdf.pages[i].extract_text(layout=True)
        if ltxt:
            for ln, line in enumerate(ltxt.split('\n')):
                print(f"[{ln:03d}] {line}")
