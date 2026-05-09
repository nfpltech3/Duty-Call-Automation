import pdfplumber
import glob

files = sorted(glob.glob("*.pdf"))
for f in files:
    print(f"\n================ {f} ================")
    with pdfplumber.open(f) as pdf:
        num_pages = len(pdf.pages)
        for p_idx in range(num_pages - 1, 1, -1):
            page_words = pdf.pages[p_idx].extract_words()
            word_set = {w['text'] for w in page_words}
            word_set_upper = {w['text'].upper() for w in page_words}
            
            # Check original set
            orig_match = word_set & {"LICENCE", "DEBIT"}
            upper_match = word_set_upper & {"LICENCE", "DEBIT"}
            
            if upper_match:
                print(f"Page {p_idx}: Upper matches: {upper_match} | Original matches: {orig_match}")
                # Print some text on that page
                text = pdf.pages[p_idx].extract_text()
                lines = text.split('\n')
                for l in lines:
                    if "debit" in l.lower() or "licence" in l.lower() or "debit duty" in l.lower():
                        print(f"  Line: {l}")
