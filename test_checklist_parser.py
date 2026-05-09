from be_extractor import parse_checklist

f = "Import CheckList-IR5459826-27-07-MAY-2026_10_49_AM.pdf"
data, warnings = parse_checklist(f)

print("=== Parsed Checklist ===")
for k, v in data.items():
    print(f"  {k}: {v}")
print(f"\n=== Warnings ({len(warnings)}) ===")
for w in warnings:
    print(f"  - {w}")

print("\n=== Expected ===")
print("  job_no: 54598")
print("  be_no: 9039377")
print("  branch: CLC")
print("  bl_no: 1073584082")
print("  assessable_value: 2651407.82")
print("  duty_amount: 595771.00")
print("  foregone_duty: 397711.17")
