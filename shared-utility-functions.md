`get_initials`
Takes a full name as input (like "Jane Doe Smith") and returns the initials in uppercase (e.g., "JDS"). 
It splits the name by spaces, grabs the first letter of each part, and joins them together.

`get_rm_initials`
Given a company name, this function looks up the corresponding "Relationship Manager" in the vendor DataFrame. 
If it finds a match, it returns that manager’s initials using get_initials. If no match is found, it returns "XX" as a placeholder. 

`get_last_modified_date`
Takes a file path and returns the last modified date of the file in YYYYMMDD format.
It uses the operating system’s file metadata to get the timestamp, then formats it as a string.

`detect_doc_type`
Scans the filename for keywords to guess the document type. 
It checks for terms like "amendment", "royalty", "annex", "proposal", "NDA", etc., and returns a standardized short label (like "Amdt", "Annex A", "Proposal", etc.). 
If no known type is found, it returns "Other".

`guess_vendor_from_filename_or_ird`
Attempts to infer the vendor name from the filename. First, it looks for a 4-digit IRD code (with optional letter) in the filename and tries to match it to the vendor list. 
If that fails, it strips out common words and numbers, then uses fuzzy matching to guess the vendor name from the remaining words. If it can’t find a good match, it returns None.
