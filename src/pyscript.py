import os
from datetime import datetime
import pandas as pd
import re
from pypdf import PdfReader
import datefinder
from rapidfuzz import process
from dateutil import parser as date_parser

excel_path = input(r"Enter the path to the Excel file: ")
owner_df = pd.read_excel(excel_path, sheet_name = input("Please enter the name of the spreadsheet: "), header=1)
show_df = owner_df.head()
# print(show_df)

def get_initials(full_name):
    return ''.join([part[0].upper() for part in str(full_name).split() if part])

def get_owner_initials(company_name): 
    row = owner_df[owner_df['Vendor'].str.lower() == company_name.lower()]
    if not row.empty:
        owner_name = row.iloc[0]['Owner']
        return get_initials(owner_name)
    return 'XX'

def get_last_modified_date(filepath):
    timestamp = os.path.getmtime(filepath)
    return datetime.fromtimestamp(timestamp).strftime('%Y%m%d')

def detect_doc_type(filename):  # if loops are in order of specificity to generality
    name = filename.lower()
    if "legal a" in name or "legal_a" in name:
        return "Legal Document A"
    elif "legal" in name or "legal document" in name:
        return "Legal Document"
    # ... Add more specific checks as needed
    else:
        return "Other" # Default case

def guess_vendor_from_filename_or_id(filename, vendor_list, id_list):
    base = os.path.splitext(os.path.basename(filename))[0]
    # Regex for ID: optional apostrophe, 4 digits, optional letter, optional apostrophe, e.g. '5097', '5097A', 5097A, etc.
    id_match = re.search(r"'?\s*(\d{4}[A-Za-z]?)\s*'?", base)
    if id_match:
        id_candidate = id_match.group(1)
        # Try to match ID in DataFrame (case-insensitive, strip spaces)
        row = owner_df[owner_df['Id'].astype(str).str.replace("'", "").str.strip().str.lower() == id_candidate.lower()]
        if not row.empty:
            return row.iloc[0]['Vendor']
    # Fallback: vendor name guessing
    words = re.split(r'[\s\-_\.]+', base)
    ignore = {'inc','ltd','legal', 'Agreement'} # Add common legal/document type words that might skew guessing
    vendor_guess = ' '.join([w for w in words if w.lower() not in ignore and not w.isdigit() and not re.match(r'\d{4}', w)])
    match = process.extractOne(vendor_guess, vendor_list, score_cutoff=60)
    return match[0] if match else None

def build_draft_filename_auto(filepath):
    vendor_names = owner_df['Vendor'].tolist()
    id_list = owner_df['Id'].astype(str).tolist()
    guessed_vendor = guess_vendor_from_filename_or_id(os.path.basename(filepath), vendor_names, id_list)
    if guessed_vendor is None:
        guessed_vendor = "UnknownVendor"
    draft_date = get_last_modified_date(filepath)
    doc_type = detect_doc_type(os.path.basename(filepath))
    owner_initials = get_owner_initials(guessed_vendor)
    return f"{draft_date}-ORG-{guessed_vendor}-{doc_type}-{owner_initials}" # Alter this format as needed.
filepath = input(r"Enter the path to the file: ") #path to your file you'd like to rename
new_name = build_draft_filename_auto(filepath)
print(new_name)
# Drafting works fine, given how the functions are defined for PDF date extraction, that should be fine too
# But I need to determine how best to handle that the program deals extracting dates from PDF's differently
# I also realized that, I don't really make it clear that the distinction between the two is how I have defined 'exectued'.
