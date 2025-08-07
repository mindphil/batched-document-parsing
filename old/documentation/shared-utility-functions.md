`get_initials`
Takes a full name as input (like "Jane Doe Smith") and returns the initials in uppercase (e.g., "JDS"). 
It splits the name by spaces, grabs the first letter of each part, and joins them together.

`get_owner_initials`
Given a company name, this function looks up the corresponding "Owner" in the data frame. 
If it finds a match, it returns that owner’s initials using get_initials. If no match is found, it returns "XX" as a placeholder. 

`get_last_modified_date`
Takes a file path and returns the last modified date of the file in YYYYMMDD format.
It uses the operating system’s file metadata to get the timestamp, then formats it as a string.

`detect_doc_type`
Scans the filename for keywords to guess the document type. 
It checks for key words, preferencing what is detected earlier in the loop and returns your input.
If no known type is found, it returns "Other".

`guess_vendor_from_filename_or_id`
Attempts to infer the vendor name from the filename or id. First, it looks for a 4-digit IRD code (with optional letter) in the filename and tries to match it to the vendor list. 
If that fails, it strips out words you add to the 'ignore' list and numbers, then uses fuzzy matching to guess the vendor name from the remaining words. If it can’t find a good match, it returns None.
