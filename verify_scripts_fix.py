from be_extractor import process_pdf

targets = [
    ("Processed BE_IR_54530_9083137.pdf", 1905909.02),
    ("Processed BE_IR_54526_9083312.pdf", 724997.83),
    ("Processed BE_IR_54285_26-27_9025944.pdf", 13033.4),
    ("Processed BE_IR_54678_9026230.pdf", 146553.0),
]

for f, expected in targets:
    data, warnings, branch = process_pdf(f, 1)
    got = data.get("SCRIPTS UTILISED VALUE", 0)
    status = "OK" if abs(got - expected) < 10 else "WRONG"
    print(f"{status} {f}: got={got}, expected={expected}, diff={round(got - expected, 2)}")
