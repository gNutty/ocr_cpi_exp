import pandas as pd
import os
import sys

# Set encoding for output
sys.stdout.reconfigure(encoding='utf-8')

file_path = "Vendor_branch.xlsx"
if os.path.exists(file_path):
    try:
        df = pd.read_excel(file_path, nrows=1)
        columns = list(df.columns)
        print("Columns:", columns)
        
        target = 'เลขประจำตัวผู้เสียภาษี'
        if target in columns:
            print(f"Found '{target}'")
        else:
            print(f"NOT Found '{target}'")
            
        target_branch = 'สาขา'
        if target_branch in columns:
            print(f"Found '{target_branch}'")
        else:
             print(f"NOT Found '{target_branch}'")
             
    except Exception as e:
        print("Error reading excel:", e)
else:
    print("File not found")
