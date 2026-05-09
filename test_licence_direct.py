import pdfplumber
import glob
import re

files = sorted(glob.glob("*.pdf"))
for f in files:
    print(f"\n================ {f} ================")
    with pdfplumber.open(f) as pdf:
        num_pages = len(pdf.pages)
        lic_tot = 0.0
        lic_found = False
        lic_start_page = -1
        
        # Scan pages from index 2 onwards to find the actual table page
        for p_idx in range(2, num_pages):
            p_text = pdf.pages[p_idx].extract_text()
            if p_text and ("CF. LICENCE DETAILS" in p_text or "11.DEBIT DUTY" in p_text):
                lic_start_page = p_idx
                break
        
        print(f"Direct Scan Licence start page: {lic_start_page}")
        if lic_start_page >= 0:
            for p_idx in range(lic_start_page, min(num_pages, lic_start_page + 4)):
                l_text = pdf.pages[p_idx].extract_text(layout=True)
                if not l_text:
                    continue
                in_lic = False
                for line in l_text.split('\n'):
                    if '11.DEBIT DUTY' in line or 'DEBIT DUTY' in line:
                        in_lic = True
                        print(f"Found DEBIT DUTY header on Page {p_idx}: {line.strip()}")
                        continue
                    if in_lic:
                        if re.match(r'^\s*\d+\s+', line):
                            parts = line.strip().split()
                            val = parts[-1] if parts else ""
                            print(f"  Data line: {line.strip()} -> Last word: {val}")
                            try:
                                lic_tot += float(val.replace(',', ''))
                                lic_found = True
                            except ValueError:
                                pass
                        else:
                            terminators = ('G. RE-IMPORT', 'GLOSSARY', 'Page ', 'H. CERTIFICATE')
                            if any(t in line for t in terminators):
                                print(f"  Hit terminator: {line.strip()}")
                                in_lic = False
                                break
        print(f"Final extracted Licence Total: {lic_tot}")
