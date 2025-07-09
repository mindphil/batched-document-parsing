extract_date_strings;
This function takes a chunk of text (usually from a PDF page) and scans it for anything that looks like a date. 
It uses several regular expressions to catch a wide variety of date formats—everything from work specific-style dates (like 2024-07-09), to US-style (07/09/2024), to written-out dates (like "January 12, 2025" or "12 January 2025"). 
It returns a list of all the date strings it finds, regardless of whether they’re valid or not. 
This is the first step in pulling out any possible date information from the document.

parse_and_validate_dates;
Once you have a list of date strings, this function tries to turn each one into a Python `datetime` object.
It uses `dateutil` to parse the string, and if it can’t be parsed (or if it falls outside a reasonable range—before 1990 or after 10 years from now), it gets ignored. 
The result is a list of valid, plausible dates that can be used for naming or further logic. This helps filter out junk or accidental matches.

find_effective_date_context;
This function looks for phrases in the text that suggest an “effective date”—things like "effective as of", "is effective", or similar.
For each pattern it finds, it grabs a chunk of text around the phrase (50 characters before and 20 after), assuming the actual date will be nearby.
It then runs `extract_date_strings` on that chunk, and if any dates are found, parses and validates them. 
If there are multiple valid dates, it returns the latest one, under the assumption that the most recent date is the one that matters in an effective context. 
If nothing is found, it returns `None`. This is a targeted way to find the “real” date a contract or agreement takes effect.

select_pdf_date_for_naming;
This is the main function that pulls everything together to extract the most appropriate date from a PDF for naming purposes.

- It starts by figuring out the document type from the filename, since some types (like Annexes) have special rules.
- It opens the PDF and extracts text from the first and last pages, since those are where legal and signature dates usually appear.
- The text is cleaned up to make pattern matching easier.
- If the document is an Annex, it looks for all dates on the first page and picks the earliest one.
- Otherwise, it tries to find an “effective date” using `find_effective_date_context` on the combined text of the first and last pages.
- If that doesn’t work, it looks for signature-related phrases on the last page and grabs the latest date found there.
- If all else fails, it collects all dates from both pages, removes duplicates, sorts them, and picks the latest one.
- If no valid dates are found, it returns "Review" to flag the file for manual checking.

Throughout, debug print statements show which strategy was used and what dates were found, making it easier to troubleshoot or refine the logic. 
The whole function is wrapped in a try-except block, so if anything goes wrong (like a corrupt PDF), it won’t crash the workflow. It just returns "Review".

build_executed_filename_auto;
This function creates a standardized filename for executed (finalized) documents.
It tries to guess the vendor from the filename, extracts the most relevant date using `select_pdf_date_for_naming`, detects the document type, and then combines all these pieces into a filename like:  
`YYYYMMDD-ISO-VendorName-DocType`  
If any part can’t be determined, it uses a placeholder (like "UnknownVendor" or "Review").