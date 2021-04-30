# xlcompare
- Compare two Excel files where rows have **unique** identifiers
- Ideal for comparing requirements, bill of materials, invoices, etc
- Supports both `.xls` and `.xlsx` files as input files
- Generates `diff.xlsx` containing differences
- Excel autofilter set to show only changed rows and columns

## Excel File Format
- First row is assumed to contain column heading names
- Headings that are common between the two files will be compared
- One column heading must be "ID" (can override with the `--id` option)

## Limitations:
- Only compares first sheet of each Excel file
- Compares as text

# Installation Instructions
```bat
python -m pip install xlcompare
```
