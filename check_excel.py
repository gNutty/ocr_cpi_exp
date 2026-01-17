import pandas as pd
import os

file_path = "Vendor_branch.xlsx"
if os.path.exists(file_path):
    try:
        df = pd.read_excel(file_path, nrows=1)
        print("Columns:", list(df.columns))
    except Exception as e:
        print("Error reading excel:", e)
else:
    print("File not found")
