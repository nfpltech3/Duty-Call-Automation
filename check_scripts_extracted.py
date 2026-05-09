from be_extractor import process_pdf
import glob

files = sorted(glob.glob("*.pdf"))
for f in files:
    print(f"\n================ {f} ================")
    data, warnings, branch = process_pdf(f, 1)
    print(f"Branch: {branch}")
    print(f"SCRIPTS UTILISED VALUE: {data.get('SCRIPTS UTILISED VALUE')}")
    print(f"Warnings: {warnings}")
