import pdfplumber
import glob
import time

files = sorted(glob.glob("*.pdf"))
for f in files:
    t0 = time.time()
    print(f"\n================ {f} ================")
    with pdfplumber.open(f) as pdf:
        num_pages = len(pdf.pages)
        print(f"Total pages: {num_pages}")
        lic_tot = 0.0
        lic_found = False
        lic_start_page = -1
        
        # Scan BACKWARDS for the actual table headers
        for p_idx in range(num_pages - 1, 1, -1):
            p_text = pdf.pages[p_idx].extract_text()
            if p_text and ("11.DEBIT DUTY" in p_text or "CF. LICENCE DETAILS" in p_text):
                lic_start_page = p_idx
                break
                
        print(f"Backward scan found table at page: {lic_start_page}")
        if lic_start_page >= 0:
            l_text = pdf.pages[lic_start_page].extract_text(layout=True)
            in_lic = False
            for line in l_text.split('\n'):
                if '11.DEBIT DUTY' in line or 'DEBIT DUTY' in line:
                    in_lic = True
                    print(f"Found DEBIT DUTY header on Page {lic_start_page}: {line.strip()}")
                    continue
                if in_lic:
                    parts = line.strip().split()
                    if parts and parts[0].isdigit():
                        val = parts[-1]
                        print(f"  Data line: {line.strip()} -> Last word: {val}")
                        try:
                            lic_tot += float(val.replace(',', ''))
                            lic_found = True
                        except ValueError:
                            pass
                    else:
                        terminators = ('G. RE-IMPORT', 'GLOSSARY', 'Page ', 'H. CERTIFICATE')
                        if parts and any(t in line for t in terminators):
                            print(f"  Hit terminator: {line.strip()}")
                            in_lic = False
                            break
        print(f"Final extracted Licence Total: {lic_tot} (took {time.time() - t0:.2f}s)")
