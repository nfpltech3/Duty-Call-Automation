import pdfplumber
import glob
import re

files = sorted(glob.glob("*.pdf"))
for f in files:
    print(f"\n================ {f} ================")
    with pdfplumber.open(f) as pdf:
        p0 = pdf.pages[0]
        words_p0 = p0.extract_words()
        
        # Group words into lines
        lines_p0 = {}
        for w in words_p0:
            top = round(w['top'] / 3) * 3
            if top not in lines_p0:
                lines_p0[top] = []
            lines_p0[top].append(w)
        sorted_tops = sorted(lines_p0.keys())
        
        fcl_lcl = ""
        for i, top in enumerate(sorted_tops):
            lw = sorted(lines_p0[top], key=lambda w: w['x0'])
            lt = ' '.join(w['text'] for w in lw)
            if '2.LCL/' in lt and '1.SNO' in lt:
                lcl_header_x0 = next((w['x0'] for w in lw if '2.LCL/' in w['text']), None)
                if lcl_header_x0 is None:
                    break
                for j in range(i + 1, min(i + 5, len(sorted_tops))):
                    vw = sorted(lines_p0[sorted_tops[j]], key=lambda w: w['x0'])
                    for w in vw:
                        if w['text'] in ('F', 'L') and abs(w['x0'] - lcl_header_x0) < 25:
                            fcl_lcl = w['text']
                            break
                        if w['text'] in ('FCL', 'LCL') and abs(w['x0'] - lcl_header_x0) < 25:
                            fcl_lcl = w['text'][0]  # F or L
                            break
                    if fcl_lcl:
                        break
                break
        print(f"Extracted fcl_lcl: {fcl_lcl}")
        
        # Now let's print the actual text on those lines
        for i, top in enumerate(sorted_tops):
            lw = sorted(lines_p0[top], key=lambda w: w['x0'])
            lt = ' '.join(w['text'] for w in lw)
            if "1.IGM" in lt and "6.MAWB" in lt:
                print(f"Header: {lt}")
                for j in range(i+1, min(i+5, len(sorted_tops))):
                    vw = sorted(lines_p0[sorted_tops[j]], key=lambda w: w['x0'])
                    vt = ' '.join(w['text'] for w in vw)
                    print(f"  Line {j}: {vt}")
