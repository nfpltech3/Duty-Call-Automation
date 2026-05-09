import pdfplumber

files = ["Processed BE_IR_54285_26-27_9025944.pdf", "Processed BE_IR_54678_9026230.pdf"]
for f in files:
    print(f"\n================ {f} ================")
    with pdfplumber.open(f) as pdf:
        p0 = pdf.pages[0]
        words = p0.extract_words()
        
        # Group words into lines
        lines_p0 = {}
        for w in words:
            top = round(w['top'] / 3) * 3
            if top not in lines_p0:
                lines_p0[top] = []
            lines_p0[top].append(w)
        sorted_tops = sorted(lines_p0.keys())
        
        # Let's see the lines around the container section (near "2.LCL/")
        for i, top in enumerate(sorted_tops):
            lw = sorted(lines_p0[top], key=lambda w: w['x0'])
            lt = ' '.join(w['text'] for w in lw)
            if '2.LCL/' in lt:
                print(f"Header top={top}: {lt}")
                # Print the next 5 lines
                for j in range(i + 1, min(i + 6, len(sorted_tops))):
                    vw = sorted(lines_p0[sorted_tops[j]], key=lambda w: w['x0'])
                    vt = ' '.join(f"[{w['text']} x0={w['x0']:.1f} x1={w['x1']:.1f}]" for w in vw)
                    print(f"  Line {j} top={sorted_tops[j]}: {vt}")
