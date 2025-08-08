# debug functionality has been added into main program, 
# made this when I was init testing class functionality.

import os
import pandas as pd
import glob

print("=== DEBUGGING DOCUMENT RENAMER ===")

# Test 1: Check if Excel file exists and can be loaded
excel_path = "vendors.xlsx"
print(f"1. Checking Excel file: {excel_path}")
print(f"   Exists: {os.path.exists(excel_path)}")

if os.path.exists(excel_path):
    try:
        df = pd.read_excel(excel_path, sheet_name="test_spreadsheet", header=1)
        print(f"   Loaded successfully: {len(df)} rows")
        print(f"   Columns: {list(df.columns)}")
    except Exception as e:
        print(f"   Error loading: {e}")

# Test 2: Check directory and files
test_dir = "tests"
print(f"\n2. Checking directory: {test_dir}")
print(f"   Exists: {os.path.exists(test_dir)}")

if os.path.exists(test_dir):
    files = glob.glob(os.path.join(test_dir, "*"))
    print(f"   Files found: {len(files)}")
    for f in files:
        print(f"     - {os.path.basename(f)}")

# Test 3: Test argparse
import sys
print(f"\n3. Command line arguments:")
print(f"   Args: {sys.argv}")

print("\n=== END DEBUG ===")