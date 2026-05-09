import pdfplumber
import glob
import re

files = sorted(glob.glob("*.pdf"))
for f in files:
    print(f"\n================ {f} ================")
    with pdfplumber.open(f) as pdf:
        p0 = pdf.pages[0]
        text = p0.extract_text()
        words = p0.extract_words()
        
        # Look for 2.LCL/
        for w in words:
            if "2.LCL" in w['text']:
                print(f"Found '2.LCL' word: {w}")
        
        # Look for FCL/LCL in text
        m_lcl = re.search(r'(FCL|LCL)', text)
        if m_lcl:
            print(f"Regex found FCL/LCL in full text: {m_lcl.group(0)}")
        else:
            print("Regex FCL/LCL NOT found in text.")
            
        # Let's see what is printed around the container section
        lines = text.split('\n')
        for line in lines:
            if 'LCL' in line or 'FCL' in line or 'MAWB' in line or 'HAWB' in line:
                print(f"Line: {line}")
