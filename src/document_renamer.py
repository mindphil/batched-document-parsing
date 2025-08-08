import os
from datetime import datetime
import pandas as pd
import re
from pypdf import PdfReader
from rapidfuzz import process
from dateutil import parser as date_parser
import argparse
import glob

# ================================
# CONFIGURATION SECTION
# ================================
EXCEL_PATH = r"C:\Users\I1000928\Projects\smart-document-renamer\vendors.xlsx"  # Configure your spreadsheet path here
HEADER_ROW = 1  # Row number where headers are located (0-indexed)

# Column names - change these to the column name
VENDOR_COLUMN = 'Vendor'
ID_COLUMN = 'Id'
OWNER_COLUMN = 'Owner'

class DocumentRenamer:
    def __init__(self):
        try:
            # Read the first sheet automatically
            self.owner_df = pd.read_excel(EXCEL_PATH, sheet_name=0, header=HEADER_ROW)
            
            # Verify required columns exist
            required_cols = [VENDOR_COLUMN, ID_COLUMN, OWNER_COLUMN]
            missing_cols = [col for col in required_cols if col not in self.owner_df.columns]
            if missing_cols:
                raise ValueError(f"Missing required columns: {missing_cols}")
            
            # Store vendor names and IDs using configured column names
            self.vendor_names = self.owner_df[VENDOR_COLUMN].tolist()
            self.id_list = self.owner_df[ID_COLUMN].astype(str).tolist()
            
            print(f"✓ Loaded vendor data from {EXCEL_PATH}")
            print(f"  Found {len(self.owner_df)} vendor records")
            
        except FileNotFoundError:
            raise Exception(f"Excel file not found: {EXCEL_PATH}\nPlease update EXCEL_PATH in the configuration section.")
        except Exception as e:
            raise Exception(f"Error loading Excel file: {str(e)}")
    
    # ================================
    # UTILITY METHODS 
    # ================================
    
    def get_initials(self, full_name):
        return ''.join([part[0].upper() for part in str(full_name).split() if part])

    def get_owner_initials(self, company_name):
        row = self.owner_df[self.owner_df[VENDOR_COLUMN].str.lower() == company_name.lower()]
        if not row.empty:
            owner_name = row.iloc[0][OWNER_COLUMN]
            return self.get_initials(owner_name)
        return 'XX'

    def detect_doc_type(self, filename):
        name = filename.lower()
        
        if "amendment" in name or "amdt" in name:
            return "Amendment"
        elif "royalty" in name or "statement" in name:
            return "Royalty Statement"
        elif "reminder" in name or "reminder letter" in name or "annual reminder" in name:
            return "Annual Reminder Letter"
        elif "annual" in name or "annual letter" in name:
            return "Annual Letter"
        elif "annex a" in name or "annex_a" in name:
            return "Annex A"
        elif "annex b" in name or "annex_b" in name:
            return "Annex B"
        elif "annex c" in name or "annex_c" in name:
            return "Annex C"
        elif "annex d" in name or "annex_d" in name:
            return "Annex D"
        elif "annex" in name:
            return "Annex"
        elif "agreement" in name:
            return "Agreement"
        elif "proposal" in name:
            return "Proposal"
        elif "nda" in name or "non disclosure" in name:
            return "NDA"
        else:
            return "Other"

    def guess_vendor_from_filepath_or_id(self, filepath, filename):
        base = os.path.splitext(os.path.basename(filename))[0]
        folder = os.path.dirname(filepath)
        
        # Try to match ID pattern first
        id_match = re.search(r"'?\s*(\d{4}[A-Za-z]?)\s*'?", base)
        if id_match:
            id_candidate = id_match.group(1)
            id_matches = self.owner_df[
                self.owner_df[ID_COLUMN].astype(str).str.replace("'", "").str.strip().str.lower() 
                == id_candidate.lower()
            ]
            if not id_matches.empty:
                return id_matches.iloc[0][VENDOR_COLUMN]
        
        # Fallback: fuzzy matching on vendor names using folder names
        words = re.split(r'[\s\-_\.]+', folder)
        ignore = {
            'verisk', 'strategic', 'alliance', 'agreement', 'moved', 'to', 'database',
            'do', 'not', 'edit', 'or', 'save', 'files', 'here', 'annex', 'draft',
            'revised', 'agreement', 'amendment', 'letter', 'reminder', 'setup', 'memo',
            'final', 'executed', 'signed', 'vsa', 'iso', 'llc', 'inc', 'corp',
            'company', 'co', 'the', 'of', 'and', 'a', 'b', 'c', 'proposal', 'forms',
            'rules', 'loss', 'costs', 'business', 'development', 'tm', 'drafts'
        }
        
        vendor_guess = ' '.join([
            w for w in words 
            if w.lower() not in ignore and not w.isdigit() and not re.match(r'\d{4}', w)
        ])
        
        if vendor_guess.strip():
            match = process.extractOne(vendor_guess, self.vendor_names, score_cutoff=40)
            return match[0] if match else None
        return None

    # ================================
    # DRAFT DOCUMENT METHODS 
    # ================================
    
    def get_last_modified_date(self, filepath):
        timestamp = os.path.getmtime(filepath)
        return datetime.fromtimestamp(timestamp).strftime('%Y%m%d')

    def build_draft_filename(self, filepath):
        guessed_vendor = self.guess_vendor_from_filepath_or_id(
            filepath, os.path.basename(filepath)
        )
        if guessed_vendor is None:
            guessed_vendor = "UnknownVendor"
        
        draft_date = self.get_last_modified_date(filepath)
        doc_type = self.detect_doc_type(os.path.basename(filepath))
        owner_initials = self.get_owner_initials(guessed_vendor)
        
        return f"{draft_date}-ISO-{guessed_vendor}-{doc_type}-{owner_initials}"

    # ================================
    # EXECUTED (PDF) DOCUMENT METHODS 
    # ================================
    
    def extract_date_strings(self, text): # regex patterns
        patterns = [
            r'\b\d{4}[-/]\d{1,2}[-/]\d{1,2}\b',  # ISO format
            r'\b\d{1,2}/\d{1,2}/\d{4}\b',  # US format MM/DD/YYYY
            r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2}(?:st|nd|rd|th)?,?\s+\d{4}\b',
            r'\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\.?\s+\d{1,2}(?:st|nd|rd|th)?,?\s+\d{4}\b',
            r'\b\d{1,2}(?:st|nd|rd|th)?\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}\b',
        ]
        
        all_dates = []
        for pattern in patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            all_dates.extend(matches)
        return all_dates

    def parse_and_validate_dates(self, date_strings):
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

    def find_effective_date_context(self, text): # 'effective' context 
        effective_patterns = [
            r'effective\s+(?:as\s+of\s+)?([^.]{0,50}(?:\d{1,2}[/-]\d{1,2}[/-]\d{4}|\w+\s+\d{1,2},?\s+\d{4})[^.]{0,20})',
            r'(?:shall\s+be\s+)?effective\s+(?:on\s+)?([^.]{0,50}(?:\d{1,2}[/-]\d{1,2}[/-]\d{4}|\w+\s+\d{1,2},?\s+\d{4})[^.]{0,20})',
            r'(?:agreement|contract|document)\s+(?:is\s+)?effective\s+([^.]{0,50}(?:\d{1,2}[/-]\d{1,2}[/-]\d{4}|\w+\s+\d{1,2},?\s+\d{4})[^.]{0,20})'
        ]

        for pattern in effective_patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                context = match.group(1)
                dates = self.extract_date_strings(context)
                if dates:
                    parsed = self.parse_and_validate_dates(dates)
                    if parsed:
                        return max(parsed)
        return None

    def select_pdf_date_for_naming(self, filepath, verbose=False):
        try:
            doc_type = self.detect_doc_type(os.path.basename(filepath))
            
            with open(filepath, 'rb') as f:
                reader = PdfReader(f)
                if len(reader.pages) == 0:
                    return "Review"
                
                first_page_text = reader.pages[0].extract_text() or ""
                last_page_text = reader.pages[-1].extract_text() if len(reader.pages) > 1 else first_page_text

                # Clean up extracted text
                first_page_clean = re.sub(r'\s+', ' ', first_page_text.strip())
                last_page_clean = re.sub(r'\s+', ' ', last_page_text.strip())
                
                # Priority 1: For Legal documents, prioritize first page
                if doc_type and doc_type.startswith("Legal"):
                    dates = self.parse_and_validate_dates(self.extract_date_strings(first_page_clean))
                    if dates:
                        return min(dates).strftime('%Y%m%d')
        
                # Priority 2: Look for effective date context
                effective_date = self.find_effective_date_context(first_page_clean + " " + last_page_clean)
                if effective_date:
                    return effective_date.strftime('%Y%m%d')

                # Priority 3: Look for signature dates (usually on last page)
                signature_context = re.search(
                    r'(?:signed|executed|dated).*?([^.]{0,100})', 
                    last_page_clean, re.IGNORECASE
                )
                if signature_context:
                    sig_dates = self.parse_and_validate_dates(
                        self.extract_date_strings(signature_context.group(1))
                    )
                    if sig_dates:
                        return max(sig_dates).strftime('%Y%m%d')

                # Fallback: all dates, prefer later ones
                all_dates = self.parse_and_validate_dates(
                    self.extract_date_strings(first_page_clean) + 
                    self.extract_date_strings(last_page_clean)
                )
                if all_dates:
                    unique_dates = list(set(all_dates))
                    unique_dates.sort()
                    return max(unique_dates).strftime('%Y%m%d')
                
                return "Review"
                
        except Exception as e:
            if verbose:
                print(f"Error processing PDF {filepath}: {str(e)}")
            return "Review"

    def build_executed_filename(self, filepath):
        guessed_vendor = self.guess_vendor_from_filepath_or_id(
            filepath, os.path.basename(filepath)
        )
        if guessed_vendor is None:
            guessed_vendor = "UnknownVendor"
        
        eff_date = self.select_pdf_date_for_naming(filepath)
        doc_type = self.detect_doc_type(os.path.basename(filepath))
        
        return f"{eff_date}-ISO-{guessed_vendor}-{doc_type}"

    # ================================
    # BATCH PROCESSING METHODS 
    # ================================
    
    def rename_document(self, filepath, doc_status):
        if not os.path.exists(filepath):
            raise FileNotFoundError(f"File not found: {filepath}")
        
        doc_status_lower = doc_status.lower()
        
        if doc_status_lower == 'draft':
            return self.build_draft_filename(filepath)
        elif doc_status_lower == 'executed':
            return self.build_executed_filename(filepath)
        else:
            raise ValueError("Document status must be either 'draft' or 'executed'")
    #batch processing method
    def batch_process(self, directory, doc_status, file_pattern="*", preview=True, rename=False):
        pattern = os.path.join(directory, file_pattern)
        files = glob.glob(pattern)
        
        if not files:
            print(f"No files found matching pattern: {file_pattern} in {directory}")
            return
        
        print(f"\n{'Preview' if preview else 'Processing'} {len(files)} files:")
        print("=" * 80)
        
        results = []
        for filepath in files:
            try:
                new_name = self.rename_document(filepath, doc_status)
                extension = os.path.splitext(filepath)[1]
                old_name = os.path.basename(filepath)
                new_full_name = new_name + extension
                
                results.append({
                    'old': old_name,
                    'new': new_full_name,
                    'status': 'OK'
                })
                
                print(f"✓ {old_name} → {new_full_name}")
                
                if rename and not preview:
                    # Actually rename the file
                    new_path = os.path.join(os.path.dirname(filepath), new_full_name)
                    os.rename(filepath, new_path)
                    
            except Exception as e:
                print(f"✗ {os.path.basename(filepath)} → ERROR: {str(e)}")
                results.append({
                    'old': os.path.basename(filepath),
                    'new': None,
                    'status': f'ERROR: {str(e)}'
                })
        
        # Summary
        print("\n" + "=" * 80)
        successful = sum(1 for r in results if r['status'] == 'OK')
        print(f"Summary: {successful}/{len(files)} files processed successfully")
        
        if rename and not preview:
            print(f"Files have been renamed.")
        elif not preview:
            print(f"Preview complete. Use --rename flag to actually rename files.")
        
        return results

def main():
    parser = argparse.ArgumentParser(
        description='Batch Document Renamer - Automatically rename documents based on content and metadata'
    )
    parser.add_argument('directory', 
                       help='Directory containing files to process')
    parser.add_argument('--status', choices=['draft', 'executed'], required=True,
                       help='Document status: draft (uses modified date) or executed (extracts date from PDF)')
    parser.add_argument('--pattern', default='*',
                       help='File pattern for batch processing (default: * for all files)')
    parser.add_argument('--rename', action='store_true',
                       help='Actually rename files (default is preview only)')

    args = parser.parse_args()

    try:
        # Check if directory exists
        if not os.path.exists(args.directory):
            print(f"Error: Directory not found: {args.directory}")
            return 1
        
        # Initialize the renamer with configured Excel file
        print(f"Initializing Document Renamer...")
        renamer = DocumentRenamer()
        
        # Process files
        renamer.batch_process(
            directory=args.directory,
            doc_status=args.status,
            file_pattern=args.pattern,
            preview=not args.rename,
            rename=args.rename
        )
        
    except Exception as e:
        print(f"Error: {str(e)}")
        return 1

    return 0

if __name__ == "__main__":
    exit(main())