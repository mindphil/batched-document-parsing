## Why This Exists

Manually renaming legal and vendor documents is tedious and error-prone, especially when filenames must encode key metadata for compliance and searchability. This tool automates the process by extracting dates, document types, and vendor names from file content and filenames, then applies a consistent naming convention.

## File Naming Standard

Renamed files follow the specific pattern I use at work but can be easily amended: 

```
YYYYMMDD-ISO-VendorName-DocType[-RMInitials]
```
- `YYYYMMDD`: Key date (last modified, effective, or signature date)
- `ISO`: Organization code
- `VendorName`: Inferred from filename or metadata
- `DocType`: Detected document type (e.g., VSAL, Amdt, NDA, etc.)
- `RMInitials`: (Optional) Relationship Manager initials for drafts

See [shared-utility-functions](shared-utility-functions) for details on how metadata is extracted.

## Documentation

- [shared-utility-functions](shared-utility-functions): Core utilities for metadata extraction and vendor matching
- [draft-naming-function](draft-naming-function): Logic for draft document renaming
- [pdf-naming-functions](pdf-naming-functions): Logic for executed PDF document renaming

## Quick Start

1. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
2. Update `EXCEL_PATH` and sheet/column names in the script to match your vendor Excel file.
3. Run the notebook or import the functions as needed.

## Notes

- Designed for legal/vendor document workflows, but adaptable to other use cases.
- Review and adjust the code for your organization’s data policies and file structures.
