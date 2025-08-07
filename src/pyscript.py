import os
from datetime import datetime
import pandas as pd
import re
from pypdf import PdfReader
from rapidfuzz import process
from dateutil import parser as date_parser

#                                               Config

excel_path = r"C:\Users\I1000928\Projects\smart-document-renamer\tests\test_vendor_document.xlsx"
owner_df = pd.read_excel(excel_path, sheet_name="test_spreadsheet", header=1)

#                                              Def & Init

filepath = input(r"Enter the path to the file: ") #path to your file you'd like to rename
extension = os.path.splitext(filepath)
type = input(r"Enter 'Draft' or 'Executed': ")
type_lower = type.lower()

#                                                Util

def get_initials(full_name):
    return ''.join([part[0].upper() for part in str(full_name).split() if part])

def get_owner_initials(company_name): 
    row = owner_df[owner_df['Vendor'].str.lower() == company_name.lower()]
    if not row.empty:
        owner_name = row.iloc[0]['Owner']
        return get_initials(owner_name)
    return 'XX'

def detect_doc_type(filename):  # if loops are in order of specificity to generality
    name = filename.lower()
    if "legal a" in name or "legal_a" in name:
        return "Legal Document A"
    elif "legal" in name or "legal document" in name:
        return "Legal Document"
    elif "agreement" in name:
        return "Agreement"
    # ... Add more specific checks as needed
    else:
        return "Other" # Default case

def guess_vendor_from_filepath_or_id(filepath, filename, vendor_list, id_list):
    base = os.path.splitext(os.path.basename(filename))[0]
    folder = os.path.dirname(filepath) # If folder name is more reliable for vendor guessing (i.e; it is more consistent), use this in the 'words' variable instead of the base filename
    # Regex for ID: optional apostrophe, 4 digits, optional letter, optional apostrophe, e.g. '5097', '5097A', 5097A, etc.
    id_match = re.search(r"'?\s*(\d{4}[A-Za-z]?)\s*'?", base)
    if id_match:
        id_candidate = id_match.group(1)
        # Try to match ID in DataFrame (case-insensitive, strip spaces)
        id_list = owner_df[owner_df['Id'].astype(str).str.replace("'", "").str.strip().str.lower() == id_candidate.lower()]
        if not id_list.empty:
            return id_list.iloc[0]['Vendor']
    # Fallback: vendor name guessing
    words = re.split(r'[\s\-_\.]+', base)
    ignore = {'inc','ltd','legal', 'Agreement'} # Add common legal/document type words that might skew guessing
    vendor_guess = ' '.join([w for w in words if w.lower() not in ignore and not w.isdigit() and not re.match(r'\d{4}', w)])
    match = process.extractOne(vendor_guess, vendor_list, score_cutoff=40)
    return match[0] if match else None

#                                    Draft Logic

if type_lower == 'draft':
    def get_last_modified_date(filepath):
        timestamp = os.path.getmtime(filepath)
        return datetime.fromtimestamp(timestamp).strftime('%Y%m%d')
        
    def build_draft_filename_auto(filepath):
        vendor_names = owner_df['Vendor'].tolist()
        id_list = owner_df['Id'].astype(str).tolist()
        guessed_vendor = guess_vendor_from_filepath_or_id(filepath, os.path.basename(filepath), vendor_names, id_list)
        if guessed_vendor is None:
            guessed_vendor = "UnknownVendor"
        draft_date = get_last_modified_date(filepath)
        doc_type = detect_doc_type(os.path.basename(filepath))
        rm_initials = get_owner_initials(guessed_vendor)
        return f"{draft_date}-ORG-{guessed_vendor}-{doc_type}-{rm_initials}"
    
    new_draft_name = build_draft_filename_auto(filepath)
    print(f"The standardized draft name is:\n{new_draft_name}")

#                                    PDF

elif type_lower == 'executed':
    #Extracting potential date strings with more specific formats
    def extract_date_strings(text):
        patterns = [ 
            r'\b\d{4}[-/]\d{1,2}[-/]\d{1,2}\b',
            r'\b\d{1,2}/\d{1,2}/\d{4}\b',
            r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2}(?:st|nd|rd|th)?,?\s+\d{4}\b',
            r'\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\.?\s+\d{1,2}(?:st|nd|rd|th)?,?\s+\d{4}\b',
            r'\b\d{1,2}(?:st|nd|rd|th)?\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}\b',
        ]

        all_dates = []
        for pattern in patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            all_dates.extend(matches)
        return all_dates

    #Parse dates and filter out obviously wrong ones
    def parse_and_validate_dates(date_strings):
        valid_dates = []
        current_year = datetime.now().year

        for date_str in date_strings:
            try:
                parsed_date = date_parser.parse(date_str, fuzzy=False)
                if 1990 <= parsed_date.year <= current_year + 10:
                    valid_dates.append(parsed_date)
            except (ValueError, TypeError):
                continue
        return valid_dates

    # Look for patterns like "effective as of", "is effective", etc.
    def find_effective_date_context(text):
        effective_patterns = [
            r'effective\s+(?:as\s+of\s+)?([^.]{0,50}(?:\d{1,2}[/-]\d{1,2}[/-]\d{4}|\w+\s+\d{1,2},?\s+\d{4})[^.]{0,20})',
            r'(?:shall\s+be\s+)?effective\s+(?:on\s+)?([^.]{0,50}(?:\d{1,2}[/-]\d{1,2}[/-]\d{4}|\w+\s+\d{1,2},?\s+\d{4})[^.]{0,20})',
            r'(?:agreement|contract|document)\s+(?:is\s+)?effective\s+([^.]{0,50}(?:\d{1,2}[/-]\d{1,2}[/-]\d{4}|\w+\s+\d{1,2},?\s+\d{4})[^.]{0,20})'
        ]

        for pattern in effective_patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                context = match.group(1)
                dates = extract_date_strings(context)
                if dates:
                    parsed = parse_and_validate_dates(dates)
                    if parsed:
                        return max(parsed)  # Return latest date in effective context
        return None

    #Big kahuna
    def select_pdf_date_for_naming(filepath):
        try:
            doc_type = detect_doc_type(os.path.basename(filepath))
            with open(filepath, 'rb') as f:
                reader = PdfReader(f)
                first_page_text = ""
                last_page_text = ""
                if len(reader.pages) > 0:
                    first_page_text = reader.pages[0].extract_text() or ""
                if len(reader.pages) > 1:
                    last_page_text = reader.pages[-1].extract_text() or ""
                else:
                    last_page_text = first_page_text

                # Clean up extracted text
                first_page_clean = re.sub(r'\s+', ' ', first_page_text.strip())
                last_page_clean = re.sub(r'\s+', ' ', last_page_text.strip())
                # Debugging output, helps to know where it pulled the date from without opening the PDF
                print(f"Debug - Doc type: {doc_type}")
                print(f"Debug - First 200 chars: {first_page_clean[:200]}")

                # Priority 1: For Legal documents, prioritize first page
                if doc_type and doc_type.startswith("Legal"):
                    dates = parse_and_validate_dates(extract_date_strings(first_page_clean))
                    if dates:
                        selected_date = min(dates)  # Use earliest date for Legal Doc
                        print(f"Debug - Legal date selected: {selected_date}")
                        return selected_date.strftime('%Y%m%d')
        
                # Priority 2: Look for effective date context
                effective_date = find_effective_date_context(first_page_clean + " " + last_page_clean)
                if effective_date:
                    print(f"Debug - Found effective date: {effective_date}")
                    return effective_date.strftime('%Y%m%d')

                # Priority 3: Look for signature dates (usually on last page)
                signature_context = re.search(r'(?:signed|executed|dated).*?([^.]{0,100})', last_page_clean, re.IGNORECASE)
                if signature_context:
                    sig_dates = parse_and_validate_dates(extract_date_strings(signature_context.group(1)))
                    if sig_dates:
                        selected_date = max(sig_dates)
                        print(f"Debug - Signature date selected: {selected_date}")
                        return selected_date.strftime('%Y%m%d')

                # If all else fails, fallback to all dates, prefer later ones
                all_dates = parse_and_validate_dates(
                    extract_date_strings(first_page_clean) + 
                    extract_date_strings(last_page_clean)
                )
                if all_dates:
                    unique_dates = list(set(all_dates))
                    unique_dates.sort()
                    selected_date = max(unique_dates)
                    print(f"Debug - Fallback date selected: {selected_date}")
                    return selected_date.strftime('%Y%m%d')
                # If no valid dates found, return "Review", debug lets me know the failure was in finding a date
                print("Debug - No valid dates found")
                return "Review"
        # and if any error occurs, return "Review" to avoid breaking the process
        except Exception as e:
            print(f"Error processing PDF {filepath}: {str(e)}")
            return "Review"

    def build_executed_filename_auto(filepath): # pieces together all the functions to build the final filename
        vendor_list = owner_df['Vendor'].tolist()
        id_list = owner_df['Id'].astype(str).tolist()
        guessed_vendor = guess_vendor_from_filepath_or_id(filepath, os.path.basename(filepath), vendor_list, id_list)
        if guessed_vendor is None:
            guessed_vendor = "UnknownVendor"
        eff_date = select_pdf_date_for_naming(filepath)
        doc_type = detect_doc_type(os.path.basename(filepath))
        return f"{eff_date}-ORG-{guessed_vendor}-{doc_type}"

    new_pdf_name = build_executed_filename_auto(filepath)
    print(f"Generated filename: {new_pdf_name}")

else:
    print(f"Need to handle misspellings/inputs more inteligently.")


# Atm, PDF pipeline does not have access to fuzzy matching from util.
# may have to refactor as a class for inheretence to avoid gargantuan code-blocks
# I also realized that, I don't really make it clear that the distinction between the two is how I have defined 'exectued'.
