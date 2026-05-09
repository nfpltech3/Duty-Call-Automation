from be_extractor import process_pdf

files = ["Processed BE_IR_54285_26-27_9025944.pdf", "Processed BE_IR_54678_9026230.pdf"]
for f in files:
    print(f"\n================ {f} ================")
    data, warnings, branch = process_pdf(f, 1)
    print(f"Detected Branch: {branch}")
    print(f"HAWB NO: {data['HAWB NO']}")
    print(f"HAWB DATE: {data['HAWB DT.']}")
    print(f"Warnings: {warnings}")
