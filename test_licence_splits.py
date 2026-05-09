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
        
        # Cheap scan for licence details page
        lic_start_page = -1
        lic_keywords = {"LICENCE", "DEBIT"}
        for p_idx in range(num_pages - 1, 1, -1):
            page_words = pdf.pages[p_idx].extract_words()
            word_set = {w['text'].upper() for w in page_words}
            if word_set & lic_keywords:
                page_text_words = ' '.join(w['text'].upper() for w in page_words)
                if "LICENCE" in page_text_words and "DEBIT" in page_text_words:
                    lic_start_page = p_idx
                    # Search backwards for actual start
                    for p_back in range(p_idx - 1, max(1, p_idx - 3) - 1, -1):
                        bw = pdf.pages[p_back].extract_words()
                        bt = ' '.join(w['text'].upper() for w in bw)
                        if "LICENCE" in bt:
                            lic_start_page = p_back
                        else:
                            break
                    break
        
        print(f"Licence start page: {lic_start_page}")
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
                            terminators = ('G. RE-IMPORT', 'GLOSSARY', 'Page ')
                            if any(t in line for t in terminators):
                                print(f"  Hit terminator: {line.strip()}")
                                in_lic = False
                                break
        print(f"Final extracted Licence Total: {lic_tot}")
