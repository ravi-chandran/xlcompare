# xlcompare
Compare two Excel files (old vs new) where each row has a *unique* ID (identifier). Ideal for comparing requirements, bill of materials, invoices, exported databases, etc. Generates differences Excel file showing changes from old to new (red strikeout for deletions, blue for insertions). Supports both `.xls` and `.xlsx` formats as inputs.

## Easy To Use
Generate differences Excel file `diff.xlsx` with:
```bash
xlcompare old.xls new.xls
```

## Features
- Supports both `.xls` and `.xlsx` file formats for input files
- Generates output Excel file containing differences (default: `diff.xlsx`)
- Output is autofiltered to show differences at a glance
- Changes in each cell are marked with red strikeout for deletions, blue for insertions
- Deleted rows will be at the bottom in red strikeout
- Pure Python (uses `xlrd`, `pylightxl`, `XlsxWriter` packages)

## Excel File Format Assumptions
- First row is assumed to contain column headings
- Columns that are common between the two files will be compared (others are ignored)
- Column containing unique IDs is labeled "ID" (can override with the `--id` option)

## Limitations:
- Only compares first sheet of each Excel file
- Compares cells as text
- Python 3.6 or later

## Installation
```bat
python -m pip install xlcompare
```

## Usage
```bash
usage: xlcompare [-h] [--id ID] [--outfile OUTFILE] [--colwidthmax COLWIDTHMAX] oldfile newfile

Compares Excel .xls or .xlsx files (first sheet only) with headers and unique row IDs; generates diff.xlsx.

positional arguments:
  oldfile               old Excel file
  newfile               new Excel file

optional arguments:
  -h, --help            show this help message and exit
  --id ID               ID column heading (default: ID)
  --outfile OUTFILE, -o OUTFILE
                        output .xlsx file of differences (default: diff.xlsx)
  --colwidthmax COLWIDTHMAX
                        maximum column width in output file (default: 50)
```

## Examples
```bash
xlcompare old.xls new.xls   # Generates diff.xlsx
xlcompare old.xls new.xlsx  # Generates diff.xlsx
xlcompare old.xlsx new.xls  # Generates diff.xlsx
xlcompare old.xls new.xls -o mydiff.xlsx # Generates mydiff.xlsx
xlcompare old.xlsx new.xls --id MYID     # Uses "MYID" as the ID column
```
