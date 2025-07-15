## Why This Exists

Renaming legal and vendor documents manually is often tedious and prone to errors, especially when filenames need to incorporate key metadata for compliance and searchability. Initially, I developed this program to reduce some of the tedious work at my job. However, I decided to generalize its functionality as much as possible. This tool automates the renaming process by extracting relevant information such as dates, document types, and vendor names from both the file content and the filenames, and then applies a consistent naming convention.

## File Naming Standard

Renamed files follow the specific pattern I use at work but can be easily amended: 

```
YYYYMMDD-ORG-VendorName-DocType[-OwnerInitials]
```
- `YYYYMMDD`: Key date (last modified, effective, or signature date)
- `ORG`: Organization acronym
- `VendorName`: Inferred from filename or metadata
- `DocType`: Detected document type (e.g., VSAL, Amdt, NDA, etc.)
- `OwnerInitials`: (Optional) Initials of individuals associated with vendor company, for drafts

## Documentation

- [shared-utility-functions](documentation/shared-utility-functions.md): Core utilities for metadata extraction and vendor matching
- [draft-naming-function](documentation/(draft-naming-function.md ): Logic for draft document renaming
- [pdf-naming-functions](documentation/pdf-naming-functions.md): Logic for executed PDF document renaming

## Quick Start

1. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
2. Update `EXCEL_PATH` and sheet/column names in the script to match your vendor Excel file.
3. Paste the file path of the files you'd like to rename into the indicated field
4. Run the notebook

## Notes

Although I tried to generalize this version, it was still designed to meet a specific standard. As I continue to learn more about programming and explore more advanced techniques, I would be open to revisiting this project to find ways to enhance its functionality for a broader range of contexts.

Currently, while the program does save some time, the process is still monotonous. As you'll see, you can only process one file at a time. I plan to evolve the program into a batch processing solution for entire folders/directories once I am confident in its accuracy. See [problems](documentation/problems.md).
