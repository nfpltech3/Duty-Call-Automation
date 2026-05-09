import pandas as pd
from be_extractor import parse_checklist, validate_be_vs_checklist, clean_point_zero

print("--- Step 1: Testing parse_checklist on the actual checklist PDF ---")
cl_path = "Import CheckList-IR5459826-27-07-MAY-2026_10_49_AM.pdf"
cl_data, cl_warnings = parse_checklist(cl_path)

print(f"Checklist Warnings count: {len(cl_warnings)}")
for w in cl_warnings:
    print(f"  Warning: {w}")

print("Checklist Data extracted:")
for k, v in cl_data.items():
    print(f"  {k}: {v}")

print("\n--- Step 2: Testing validate_be_vs_checklist with perfect match (OK) ---")
# Simulating extracted BE data matching perfectly
mock_be_data_ok = {
    "BOE NO.": "9039377",
    "HAWB NO": "1073584082",
    "ASS. VALUE": "2651407.82",
    "TOTAL DUTY AMT (CASH)": "595771.00",
    "SCRIPTS UTILISED VALUE": "397711.17"
}

res_ok = validate_be_vs_checklist(mock_be_data_ok, cl_data, "CLC")
print(f"Validation Status (Expected OK): {res_ok['status']}")
print(f"Branch match: {res_ok['branch_match']}, BL match: {res_ok['bl_match']}")
print(f"Ass Diff: {res_ok['ass_diff']}, Duty Diff: {res_ok['duty_diff']}, Foregone Diff: {res_ok['foregone_diff']}")
print(f"Remarks: {res_ok['remarks']}")

print("\n--- Step 3: Testing validate_be_vs_checklist with minor verification issue (WARNING) ---")
# Simulating branch mismatch or BL mismatch
mock_be_data_warn = {
    "BOE NO.": "9039377",
    "HAWB NO": "WRONG_BL_123",
    "ASS. VALUE": "2651407.82",
    "TOTAL DUTY AMT (CASH)": "595771.00",
    "SCRIPTS UTILISED VALUE": "397711.17"
}

res_warn = validate_be_vs_checklist(mock_be_data_warn, cl_data, "CSN")
print(f"Validation Status (Expected WARNING): {res_warn['status']}")
print(f"Branch match: {res_warn['branch_match']}, BL match: {res_warn['bl_match']}")
print(f"Remarks: {res_warn['remarks']}")

print("\n--- Step 4: Testing validate_be_vs_checklist with major value differences (HIGH RISK) ---")
# Simulating value differences
mock_be_data_hr = {
    "BOE NO.": "9039377",
    "HAWB NO": "1073584082",
    "ASS. VALUE": "2650000.00", # Differs by -1407.82
    "TOTAL DUTY AMT (CASH)": "595771.00",
    "SCRIPTS UTILISED VALUE": "390000.00" # Differs by -7711.17
}

res_hr = validate_be_vs_checklist(mock_be_data_hr, cl_data, "CLC")
print(f"Validation Status (Expected HIGH RISK): {res_hr['status']}")
print(f"Ass Diff: {res_hr['ass_diff']}, Duty Diff: {res_hr['duty_diff']}, Foregone Diff: {res_hr['foregone_diff']}")
print(f"Remarks: {res_hr['remarks']}")

print("\n--- Step 5: Testing validate_be_vs_checklist with minor value differences within +- 10 (WARNING) ---")
# Simulating value differences within +- 10
mock_be_data_minor = {
    "BOE NO.": "9039377",
    "HAWB NO": "1073584082",
    "ASS. VALUE": "2651405.00", # Differs by -2.82
    "TOTAL DUTY AMT (CASH)": "595775.00", # Differs by +4.00
    "SCRIPTS UTILISED VALUE": "397709.00" # Differs by -2.17
}

res_minor = validate_be_vs_checklist(mock_be_data_minor, cl_data, "CLC")
print(f"Validation Status (Expected WARNING): {res_minor['status']}")
print(f"Ass Diff: {res_minor['ass_diff']}, Duty Diff: {res_minor['duty_diff']}, Foregone Diff: {res_minor['foregone_diff']}")

print("\n--- Step 6: Testing BL / HBL normalization (Expected OK) ---")
# Simulating BL formatting differences: spaces, hyphens, slashes, uppercase
mock_be_data_norm = {
    "BOE NO.": "9039377",
    "HAWB NO": "1073-584-082 ", # Differs by hyphens and spaces
    "ASS. VALUE": "2651407.82",
    "TOTAL DUTY AMT (CASH)": "595771.00",
    "SCRIPTS UTILISED VALUE": "397711.17"
}

res_norm = validate_be_vs_checklist(mock_be_data_norm, cl_data, "CLC")
print(f"Validation Status (Expected OK): {res_norm['status']}")
print(f"BL match: {res_norm['bl_match']}")

print("--- End of Integration Test ---")
