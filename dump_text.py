import pdfplumber

# Let's extract the full text of Processed BE_IR_54285_26-27_9025944.pdf and write to a txt file
with pdfplumber.open("Processed BE_IR_54285_26-27_9025944.pdf") as pdf:
    with open("text_all.txt", "w", encoding="utf-8") as f:
        for idx, page in enumerate(pdf.pages):
            f.write(f"\n--- PAGE {idx} ---\n")
            f.write(page.extract_text() or "")
            f.write(f"\n--- PAGE {idx} LAYOUT ---\n")
            f.write(page.extract_text(layout=True) or "")
print("Done writing!")
