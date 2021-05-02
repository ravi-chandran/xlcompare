#!/usr/bin/env python3
import argparse
from collections import OrderedDict
import difflib
from enum import IntEnum
import os
import pylightxl
import sys

# Pypi Packages
import xlrd
import xlsxwriter

DESCRIPTION = 'Compares Excel .xls or .xlsx files (first sheet only) with ' \
              + 'headers and unique row IDs; generates diff.xlsx.'


DEFAULT_COL_WIDTH = 10  # slightly larger than Excel default


class Fmt(IntEnum):
    """Definition for convenience in format strings for xlsxwriter."""
    WRAP = 1
    BOLD = 2
    ITALIC = 2
    WRAPBORDER = 4
    HROW = 5
    H1 = 6
    H2 = 7
    H3 = 8
    BLACKTEXT = 9
    GRAYTEXT = 10
    COURIER = 11
    DEL = 12
    INS = 13
    DEFAULT = 20


FMT = {}  # make the formats global for simplicity


def get_xlsx_formats():
    return FMT, Fmt


def create_xlsx(outfilename):
    """Create the output .xlsx file with the appropriate formats."""
    # create output .xlsx file
    wb = xlsxwriter.Workbook(outfilename)

    # format the output .xlsx file
    FMT[Fmt.WRAP] = wb.add_format({'text_wrap': True})
    # FMT[Fmt.ITALIC] = wb.add_format({'italic': True})
    FMT[Fmt.BOLD] = wb.add_format({'bold': True})

    # Courier New is an equally spaced font, useful for tables
    FMT[Fmt.COURIER] = wb.add_format(
        {'font_name': 'Courier New', 'text_wrap': True})

    FMT[Fmt.WRAPBORDER] = wb.add_format(
        {'text_wrap': True,
         'num_format': '@',
         'border': 1, 'bottom': 1, 'top': 1, 'left': 1, 'right': 1})

    FMT[Fmt.DEFAULT] = FMT[Fmt.WRAPBORDER]

    FMT[Fmt.INS] = wb.add_format(
        {'text_wrap': True,
         'font_color': 'blue',
         'num_format': '@',
         'border': 1, 'bottom': 1, 'top': 1, 'left': 1, 'right': 1})

    FMT[Fmt.DEL] = wb.add_format(
        {'text_wrap': True,
         'font_color': 'red',
         'font_strikeout': True,
         'num_format': '@',
         'border': 1, 'bottom': 1, 'top': 1, 'left': 1, 'right': 1})

    FMT[Fmt.BLACKTEXT] = FMT[Fmt.WRAPBORDER]

    FMT[Fmt.GRAYTEXT] = wb.add_format(
        {'text_wrap': True,
         'font_color': 'gray',
         'num_format': '@',
         'border': 1, 'bottom': 1, 'top': 1, 'left': 1, 'right': 1})

    FMT[Fmt.HROW] = wb.add_format(
        {'bold': True, 'font_color': 'white',
         'num_format': '@',
         'align': 'center', 'bg_color': '#0070c0', 'text_wrap': True,
         'border': 1, 'bottom': 1, 'top': 1, 'left': 1, 'right': 1})

    FMT[Fmt.H1] = wb.add_format(
        {'bold': True, 'font_color': 'white',
         'num_format': '@',
         'font_size': 14,
         'bg_color': '#808080', 'text_wrap': True,
         'border': 1, 'bottom': 1, 'top': 1, 'left': 1, 'right': 1})

    FMT[Fmt.H2] = wb.add_format(
        {'bold': True, 'font_color': 'white',
         'num_format': '@',
         'font_size': 12,
         'bg_color': '#808080', 'text_wrap': True,
         'border': 1, 'bottom': 1, 'top': 1, 'left': 1, 'right': 1})

    FMT[Fmt.H3] = wb.add_format(
        {'bold': True, 'font_color': 'white',
         'num_format': '@',
         'font_size': 11,
         'bg_color': '#808080', 'text_wrap': True,
         'border': 1, 'bottom': 1, 'top': 1, 'left': 1, 'right': 1})

    return wb


def write_header_row_xlsx(wb, hdr2width):
    """Write header row to .xlsx file."""
    ws = wb.add_worksheet()

    # freeze top row
    ws.freeze_panes(1, 0)

    # write header row
    row, col = 0, 0
    for heading in hdr2width:
        s = heading
        ws.write_string(row, col, s)
        ws.set_column(col, col, hdr2width[s])
        col += 1

    ws.write_string(row, col, 'Changed')

    ws.set_row(row, None, FMT[Fmt.HROW])

    return ws


def replace_bullet(s):
    return s.replace('*. ', '\u2022 ')


def compare_celltext(a, b):
    """Compare 2 strings and generate formatted string for output .xlsx."""
    cmp = []  # initialize compare list
    sm = difflib.SequenceMatcher(None, a, b)

    junk = ''
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        junk += '{:7}   a[{}:{}] --> b[{}:{}] {!r:>8} --> {!r}\n'.format(
            tag, i1, i2, j1, j2, a[i1:i2], b[j1:j2])

    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == 'equal' and a[i1:i2]:
            cmp.append(a[i1:i2])
        elif tag == 'delete' and a[i1:i2]:
            cmp.append(FMT[Fmt.DEL])
            cmp.append(a[i1:i2])
        elif tag == 'insert' and b[j1:j2]:
            cmp.append(FMT[Fmt.INS])
            cmp.append(b[j1:j2])
        elif tag == 'replace':
            if a[i1:i2]:
                cmp.append(FMT[Fmt.DEL])
                cmp.append(a[i1:i2])
            if b[j1:j2]:
                cmp.append(FMT[Fmt.INS])
                cmp.append(b[j1:j2])

    return cmp, junk


def write_cell(ws, row, col, list_out):
    """Write rich string to cell."""
    x = ws.write_rich_string(row, col, *list_out, FMT[Fmt.WRAPBORDER])
    if x < 0:
        print('ERROR: write_rich_string returned %d' % x)
        print('row=%d, col=%d' % (row, col))
        print(*list_out)
        sys.exit(1)


def compare_sheets(ws_out, tbl_old, tbl_new, hdr2width, id_column):
    """Compare tables from old and new files."""
    statistics = OrderedDict([
        ('Inserted', 0),
        ('Deleted', 0),
        ('Modified', 0)
        ])
    row, col = 1, 0
    hidden_rows = set()  # set of rows to hide using auto-filter
    visible_cols = {hdr2width[id_column]}  # set of columns to not hide
    blank_d = OrderedDict()
    for s in hdr2width:
        blank_d[s] = ''  # blank to compare to new or deleted objects

    # Create dictionaries from tables
    dct_old = OrderedDict()
    for dct in tbl_old:
        dct_old[dct[id_column]] = dct.copy()

    dct_new = OrderedDict()
    for dct in tbl_new:
        dct_new[dct[id_column]] = dct.copy()

    # Get union of object IDs from old and new
    union_objid = list(dct_new.keys())
    for objid in dct_old:
        if objid not in union_objid:
            union_objid.append(objid)

    # Loop through all objects
    for objid in union_objid:
        bool_diff = False  # flag to indicate difference exists in row
        bool_row_inserted_deleted = False
        if objid in dct_new and objid not in dct_old:  # inserted object
            d_new = dct_new[objid]
            d_old = blank_d
            statistics['Inserted'] += 1
            bool_row_inserted_deleted = True
        elif objid not in dct_new and objid in dct_old:  # deleted object
            d_new = blank_d
            d_old = dct_old[objid]
            statistics['Deleted'] += 1
            bool_row_inserted_deleted = True
        else:
            d_new = dct_new[objid]
            d_old = dct_old[objid]

        # compare columns for current object
        col = 0
        for i, s in enumerate(hdr2width):
            if d_new[s] != d_old[s]:
                bool_diff = True
                visible_cols.add(i)  # mark the column to be visible

            if d_new[s] == '' and d_old[s] == '':
                ws_out.write_blank(row, col, '', FMT[Fmt.WRAPBORDER])
            elif d_new[s].strip() == '' and d_old[s].strip() == '':
                ws_out.write_blank(row, col, '', FMT[Fmt.WRAPBORDER])
            elif d_new[s] == d_old[s]:
                ws_out.write_string(row, col,
                                    replace_bullet(d_old[s]),
                                    FMT[Fmt.WRAPBORDER])
            elif d_new[s] == '':
                ws_out.write_string(row, col,
                                    replace_bullet(d_old[s]),
                                    FMT[Fmt.DEL])
            elif d_old[s] == '':
                ws_out.write_string(row, col,
                                    replace_bullet(d_new[s]),
                                    FMT[Fmt.INS])
            else:
                list_out, junk = compare_celltext(replace_bullet(d_old[s]),
                                                  replace_bullet(d_new[s]))
                ws_out.write_rich_string(row, col,
                                         *list_out,
                                         FMT[Fmt.WRAPBORDER])

            col += 1

        if not bool_diff:
            ws_out.write_string(row, col, 'No', FMT[Fmt.WRAPBORDER])
            hidden_rows.add(row)
        else:
            ws_out.write_string(row, col, 'Yes', FMT[Fmt.WRAPBORDER])
            if not bool_row_inserted_deleted:
                statistics['Modified'] += 1

        row += 1

    # set column widths of output sheet and hide unchanged columns
    for h, width in hdr2width.items():
        if i in visible_cols:
            ws_out.set_column(i, i, width)
        else:
            ws_out.set_column(i, i, width, None, {'hidden': 1})

    # enable auto-filter and filter non-blank entries in "Changed" column
    ws_out.autofilter(0, 0, row-1, len(hdr2width))
    ws_out.filter_column(len(hdr2width), 'x == NonBlanks')
    for i in hidden_rows:
        ws_out.set_row(i, options={'hidden': True})

    # report statistics
    num_changes = 0
    for k, v in statistics.items():
        num_changes += v
    if num_changes == 0:
        print('No differences found.')
    else:
        for k, v in statistics.items():
            if v > 0:
                print(f'{k} rows: {v}')


def compare_headers(hdr2width_old, hdr2width_new, colwidthmax):
    """Compare headers."""
    in_old_but_not_new = set(hdr2width_old.keys()).difference(
                            set(hdr2width_new.keys()))
    in_new_but_not_old = set(hdr2width_new.keys()).difference(
                            set(hdr2width_old.keys()))
    hdr_common = set(hdr2width_old.keys()).intersection(
                            set(hdr2width_new.keys()))

    if in_old_but_not_new:
        print("Columns in old but not new:", in_old_but_not_new)

    if in_new_but_not_old:
        print("Columns in new but not old:", in_new_but_not_old)

    # Rearrange headings in original order
    hdr2width = OrderedDict()
    for s in hdr2width_old:
        if s in hdr2width_new:
            hdr2width[s] = min(hdr2width_old[s], colwidthmax)

    return hdr2width


def cell_to_text(ws, row, col):
    """Convert cell text to string."""
    cell_type = ws.cell_type(row, col)
    if cell_type == xlrd.XL_CELL_TEXT:
        value = ws.cell_value(row, col)
    elif cell_type == xlrd.XL_CELL_BLANK:
        value = ''
    else:
        value = str(ws.cell_value(row, col))
    return value


def integerize_column(tbl, heading):
    """Get rid of decimal points and places in ID field if a number."""
    for dct in tbl:
        s = dct[heading]
        if s.replace('.', '', 1).isdigit():
            dct[heading] = str(int(float(dct[heading])))


def estimate_column_width(text, initial_width):
    """Estimate column width from list of text, and initial width."""
    col_width = 0
    for s in text.splitlines():
        col_width = max(col_width, len(s))

    final_width = max(initial_width, int(1.25 * col_width))

    return final_width


def error_check_id(hdr2width, id_column, filepath):
    """Check whether ID column is present in header."""
    if id_column not in hdr2width:
        _, filename = os.path.split(filepath)
        print(f'ERROR: Column {id_column} not found in {filename}')
        sys.exit(1)


def read_xls(xlsfile, integerize_id=True, id_column='ID'):
    """Read the first sheet of .xls file."""
    wb = xlrd.open_workbook(xlsfile)
    ws = wb.sheet_by_index(0)
    print(f'{xlsfile}: Reading: {ws.name}')
    tbl, hdr2width = read_sheet_xls(ws)

    error_check_id(hdr2width, id_column, xlsfile)

    if integerize_id:
        integerize_column(tbl, id_column)

    return tbl, hdr2width


def read_sheet_xls(ws):
    """Read sheet into dictionary from .xls file."""
    hdr = []     # list of header row elements
    hdr2width = OrderedDict()  # column width of given header
    tbl = []     # list of rows of spreadsheet

    # read header row
    for col in range(ws.ncols):
        h = cell_to_text(ws, 0, col)
        hdr.append(h)
        hdr2width[h] = int(len(h) * 1.25)

    # read data rows
    for row in range(1, ws.nrows):
        d = OrderedDict()
        for col in range(ws.ncols):
            h = hdr[col]
            d[h] = cell_to_text(ws, row, col)
            hdr2width[h] = estimate_column_width(d[h], hdr2width[h])

        tbl.append(d.copy())

    return tbl, hdr2width


def read_xlsx(xlsxfile, integerize_id=True, id_column='ID'):
    """Read the first sheet of .xlsx file."""
    db = pylightxl.readxl(fn=xlsxfile)
    ws_name = db.ws_names[0]
    print(f'{xlsxfile}: Reading: {ws_name}')
    tbl, hdr2width = read_sheet_xlsx(db, ws_name)

    error_check_id(hdr2width, id_column, xlsxfile)

    if integerize_id:
        integerize_column(tbl, id_column)

    return tbl, hdr2width


def read_sheet_xlsx(db, ws_name):
    """Read sheet into dictionary from .xlsx file."""
    hdr = []     # list of header row elements
    hdr2width = OrderedDict()  # column width of given header
    tbl = []     # list of rows of spreadsheet

    # read header row
    header_data = db.ws(ws=ws_name).row(row=1)
    for col in range(len(header_data)):
        h = str(header_data[col])
        hdr.append(h)
        hdr2width[h] = int(len(h) * 1.25)

    # read data rows
    skipped_first_row = False
    for row_data in db.ws(ws=ws_name).rows:
        if not skipped_first_row:
            skipped_first_row = True
            continue
        d = OrderedDict()
        for col in range(len(row_data)):
            h = hdr[col]
            d[h] = str(row_data[col])
            hdr2width[h] = estimate_column_width(d[h], hdr2width[h])

        tbl.append(d.copy())

    return tbl, hdr2width


def get_user_inputs():
    """Get user arguments and open files."""
    # get paths of files to be compared
    parser = argparse.ArgumentParser(
        description=DESCRIPTION,
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument('oldfile', help='old Excel file')
    parser.add_argument('newfile', help='new Excel file')
    parser.add_argument('--id',
                        help='ID column heading',
                        default='ID')
    parser.add_argument('--outfile', '-o',
                        help='output .xlsx file of differences',
                        default='diff.xlsx')
    parser.add_argument('--colwidthmax',
                        help='maximum column width in output file',
                        default=50)
    args = parser.parse_args()

    # Verify that files exist
    if not os.path.isfile(args.oldfile):
        print(f'ERROR: {args.oldfile} not found')
        sys.exit(1)
    if not os.path.isfile(args.newfile):
        print(f'ERROR: {args.newfile} not found')
        sys.exit(1)

    return args


def main():
    args = get_user_inputs()

    # Read data from Excel files
    if args.oldfile.endswith('.xls'):
        tbl_old, hdr2width_old = read_xls(args.oldfile, id_column=args.id)
    else:
        tbl_old, hdr2width_old = read_xlsx(args.oldfile, id_column=args.id)

    if args.newfile.endswith('.xls'):
        tbl_new, hdr2width_new = read_xls(args.newfile, id_column=args.id)
    else:
        tbl_new, hdr2width_new = read_xlsx(args.newfile, id_column=args.id)

    # Compare header rows
    hdr2width = compare_headers(hdr2width_old, hdr2width_new, args.colwidthmax)

    # Create output differences .xlsx file
    wb_out = create_xlsx(args.outfile)
    ws_out = write_header_row_xlsx(wb_out, hdr2width)

    # Compare sheets
    compare_sheets(ws_out, tbl_old, tbl_new, hdr2width, args.id)

    # close and quit
    wb_out.close()

    print('Generated', args.outfile)
    print('Done.')


if __name__ == "__main__":
    main()
