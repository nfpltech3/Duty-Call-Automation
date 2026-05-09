from be_extractor import process_pdf, clean_point_zero
import pandas as pd

f = "Processed BE_IR_54678_9026230.pdf"
data, warnings, branch = process_pdf(f, 1)

# Format and put into a DataFrame list
from be_extractor import format_pune
formatted = format_pune(data)

df = pd.DataFrame([formatted])
print("Original column value:", df["Debit Duty"][0])

# Clean using map
df_cleaned = df.map(clean_point_zero) if hasattr(df, 'map') else df.applymap(clean_point_zero)
print("Cleaned column value :", df_cleaned["Debit Duty"][0])
