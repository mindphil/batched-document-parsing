import os
from datetime import datetime
import pandas as pd
import re
from pypdf import PdfReader
from rapidfuzz import process
from dateutil import parser as date_parser
import argparse

class DocumentRenamer:
    def __init__(self, excel_path, sheet_name="test_spreadsheet", header=1):
        self.excel_path = excel_path
        try:
            self.owner_df = pd.read_excel(excel_path, sheet_name=sheet_name, header=header)
            self.vendor_names = self.owner_df['Vendor'].tolist()
            self.id_list = self.owner_df['Id'].astype(str).tolist()
        except Exception as e:
            raise Exception(f"Error loading Excel file: {str(e)}")
    
    #                           UTILITY METHODS 
    
    def get_initials(self, full_name):
        return ''.join([part[0].upper() for part in str(full_name).split() if part])

    def get_owner_initials(self, company_name):
        row = self.owner_df[self.owner_df['Vendor'].str.lower() == company_name.lower()]
        if not row.empty:
            owner_name = row.iloc[0]['Owner']
            return self.get_initials(owner_name)
        return 'XX'

    def detect_doc_type(self, filename):
        name = filename.lower()
        if "legal a" in name or "legal_a" in name:
            return "Legal Document A"
        elif "legal" in name or "legal document" in name:
            return "Legal Document"
        elif "amendment" in name or "amdt" in name:
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
                self.owner_df['Id'].astype(str).str.replace("'", "").str.strip().str.lower() 
                == id_candidate.lower()
            ]
            if not id_matches.empty:
                return id_matches.iloc[0]['Vendor']
        
        # Fallback: fuzzy matching on vendor names
        words = re.split(r'[\s\-_\.]+', base)
        ignore = {
            'inc', 'ltd', 'legal', 'agreement', 'draft', 'final', 'executed', 
            'signed', 'company', 'corp', 'llc', 'the', 'of', 'and'
        }
        vendor_guess = ' '.join([
            w for w in words 
            if w.lower() not in ignore and not w.isdigit() and not re.match(r'\d{4}', w)
        ])
        
        if vendor_guess.strip():
            match = process.extractOne(vendor_guess, self.vendor_names, score_cutoff=40)
            return match[0] if match else None
        return None

    #                    DRAFT DOCUMENT METHODS 
    
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
        
        return f"{draft_date}-ORG-{guessed_vendor}-{doc_type}-{owner_initials}"

    #                      EXECUTED (PDF) DOCUMENT METHODS 
    
    def extract_date_strings(self, text):
        """Extract potential date strings using regex patterns."""
        patterns = [
            r'\b\d{4}[-/]\d{1,2}[-/]\d{1,2}\b',  # ISO format
            r'\b\d{1,2}/\d{1,2}/\d{4}\b',  # US format MM/DD/YYYY
            r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2}(?:st|nd|rd|th)?,?\s+\d{4}\b',  # Full month names
            r'\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\.?\s+\d{1,2}(?:st|nd|rd|th)?,?\s+\d{4}\b',  # Abbreviated months
            r'\b\d{1,2}(?:st|nd|rd|th)?\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}\b',  # Day first
        ]
        
        all_dates = []
        for pattern in patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            all_dates.extend(matches)
        return all_dates

    def parse_and_validate_dates(self, date_strings):
        """Parse date strings and filter out invalid ones."""
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

    def find_effective_date_context(self, text):   # some agreement dates follow after "as effective of ..."
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

    def select_pdf_date_for_naming(self, filepath, verbose=True):
        try:
            doc_type = self.detect_doc_type(os.path.basename(filepath))
            
            with open(filepath, 'rb') as f:
                reader = PdfReader(f)
                if len(reader.pages) == 0:
                    if verbose:
                        print("Debug - PDF has no pages")
                    return "Review"
                
                first_page_text = reader.pages[0].extract_text() or ""
                last_page_text = reader.pages[-1].extract_text() if len(reader.pages) > 1 else first_page_text

                # Clean up extracted text
                first_page_clean = re.sub(r'\s+', ' ', first_page_text.strip())
                last_page_clean = re.sub(r'\s+', ' ', last_page_text.strip())
                
                if verbose:
                    print(f"Debug - Doc type: {doc_type}")
                    print(f"Debug - First 200 chars: {first_page_clean[:200]}")

                # Priority 1: For Legal documents, prioritize first page
                if doc_type and doc_type.startswith("Legal"):
                    dates = self.parse_and_validate_dates(self.extract_date_strings(first_page_clean))
                    if dates:
                        selected_date = min(dates)
                        if verbose:
                            print(f"Debug - Legal date selected: {selected_date}")
                        return selected_date.strftime('%Y%m%d')
        
                # Priority 2: Look for effective date context
                effective_date = self.find_effective_date_context(first_page_clean + " " + last_page_clean)
                if effective_date:
                    if verbose:
                        print(f"Debug - Found effective date: {effective_date}")
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
                        selected_date = max(sig_dates)
                        if verbose:
                            print(f"Debug - Signature date selected: {selected_date}")
                        return selected_date.strftime('%Y%m%d')

                # Fallback: all dates, prefer later ones
                all_dates = self.parse_and_validate_dates(
                    self.extract_date_strings(first_page_clean) + 
                    self.extract_date_strings(last_page_clean)
                )
                if all_dates:
                    unique_dates = list(set(all_dates))
                    unique_dates.sort()
                    selected_date = max(unique_dates)
                    if verbose:
                        print(f"Debug - Fallback date selected: {selected_date}")
                    return selected_date.strftime('%Y%m%d')
                
                if verbose:
                    print("Debug - No valid dates found")
                return "Review"
                
        except Exception as e:
            if verbose:
                print(f"Error processing PDF {filepath}: {str(e)}")
            return "Review"

    def build_executed_filename(self, filepath, verbose=True):
        guessed_vendor = self.guess_vendor_from_filepath_or_id(
            filepath, os.path.basename(filepath)
        )
        if guessed_vendor is None:
            guessed_vendor = "UnknownVendor"
        
        eff_date = self.select_pdf_date_for_naming(filepath, verbose)
        doc_type = self.detect_doc_type(os.path.basename(filepath))
        
        return f"{eff_date}-ORG-{guessed_vendor}-{doc_type}"

    #                        MAIN INTERFACE METHODS 
    
    def rename_document(self, filepath, doc_status, verbose=True):
        if not os.path.exists(filepath):
            raise FileNotFoundError(f"File not found: {filepath}")
        
        doc_status_lower = doc_status.lower()
        
        if doc_status_lower == 'draft':
            return self.build_draft_filename(filepath)
        elif doc_status_lower == 'executed':
            return self.build_executed_filename(filepath, verbose)
        else:
            raise ValueError("Document status must be either 'draft' or 'executed'")

    def preview_rename(self, filepath, doc_status): # debug to preview what the name would be
        try:
            new_name = self.rename_document(filepath, doc_status, verbose=True)
            original_name = os.path.basename(filepath)
            extension = os.path.splitext(filepath)[1]
            
            print(f"\nRename Preview:")
            print(f"Original:  {original_name}")
            print(f"New name:  {new_name}{extension}")
            print(f"Status:    {doc_status.title()}")
            
            return new_name + extension
            
        except Exception as e:
            print(f"Error generating preview: {str(e)}")
            return None

    def batch_preview(self, directory, doc_status, file_pattern="*"):
        """Preview renames for multiple files in a directory."""
        import glob
        
        if not os.path.exists(directory):
            print(f"Directory not found: {directory}")
            return
        
        pattern = os.path.join(directory, file_pattern)
        files = glob.glob(pattern)
        
        if not files:
            print(f"No files found matching pattern: {file_pattern}")
            return
        
        print(f"\nBatch Preview for {len(files)} files:")
        print("=" * 80)
        
        for filepath in files:
            try:
                new_name = self.rename_document(filepath, doc_status, verbose=False)
                extension = os.path.splitext(filepath)[1]
                print(f"{os.path.basename(filepath)} → {new_name}{extension}")
            except Exception as e:
                print(f"{os.path.basename(filepath)} → ERROR: {str(e)}")


def main():
    parser = argparse.ArgumentParser(description='Smart Document Renamer')
    parser.add_argument('--excel', required=True, 
                       help='Path to Excel file with vendor data')
    parser.add_argument('--file', 
                       help='Path to single file to rename')
    parser.add_argument('--directory', 
                       help='Directory for batch processing')
    parser.add_argument('--status', choices=['draft', 'executed'], required=True,
                       help='Document status: draft or executed')
    parser.add_argument('--preview', action='store_true',
                       help='Preview mode - show proposed names without renaming')
    parser.add_argument('--pattern', default='*',
                       help='File pattern for batch processing (default: *)')
    parser.add_argument('--sheet', default='test_spreadsheet',
                       help='Excel sheet name (default: test_spreadsheet)')
    parser.add_argument('--header', type=int, default=1,
                       help='Excel header row (default: 1)')

    args = parser.parse_args()

    try:
        # Initialize the renamer
        renamer = DocumentRenamer(args.excel, args.sheet, args.header)
        print(f"Loaded vendor data: {len(renamer.owner_df)} records")

        if args.file:
            # Single file processing
            if args.preview:
                renamer.preview_rename(args.file, args.status)
            else:
                new_name = renamer.rename_document(args.file, args.status)
                extension = os.path.splitext(args.file)[1]
                print(f"Suggested filename: {new_name}{extension}")
        
        elif args.directory:
            # Batch processing
            if args.preview:
                renamer.batch_preview(args.directory, args.status, args.pattern)
            else:
                print("Batch renaming not yet implemented - use --preview to see proposed names")
        
        else:
            # Interactive mode
            filepath = input("Enter the path to the file: ")
            renamer.preview_rename(filepath, args.status)
            
            response = input("\nProceed with this rename? (y/n): ")
            if response.lower() == 'y':
                new_name = renamer.rename_document(filepath, args.status)
                extension = os.path.splitext(filepath)[1]
                print(f"Final suggested name: {new_name}{extension}")

    except Exception as e:
        print(f"Error: {str(e)}")
        return 1

    return 0


if __name__ == "__main__":
    exit(main())