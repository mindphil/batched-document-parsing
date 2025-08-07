`build_draft_filename_auto`
This function builds a standardized filename for draft documents. 
It tries to guess the vendor from the filename, gets the last modified date, detects the document type, and finds the owner's initials.
It then combines all these pieces this format:YYYYMMDD-ORG-VendorName-DocType-OwnerInitials
If any part can’t be determined, it uses a placeholder (like "UnknownVendor" or "XX").
This ensures all draft files follow a consistent, informative naming pattern, making them easier to organize and track.
