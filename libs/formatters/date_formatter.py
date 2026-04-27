"""
Formatter for date-only columns (Vietnamese locale: dd/MM/yyyy).

Algorithm:
  1. Concatenate token[0] + token[1] to reunite the split year, e.g.:
       "07/01/20" + "23"  -> "07/01/2023"
       "10/01020" + "23"  -> "10/0102023"  (0 is misread /)
  2. Extract by position — ignore whatever character sits at index 2 and 5:
       [0:2] = DD,  [3:5] = MM,  [6:10] = YYYY
"""
from datetime import date


EXCEL_DATE_FORMAT = "DD/MM/YYYY"


def _normalize_date(raw):
    tokens = raw.strip().split()

    # If token[0] already has a 4-digit year (length >= 10), don't consume token[1].
    if len(tokens) >= 2 and len(tokens[0]) < 10:
        date_raw = tokens[0] + tokens[1]
    elif len(tokens) >= 1:
        date_raw = tokens[0]
    else:
        return None

    if len(date_raw) < 10:
        return None

    dd   = date_raw[0:2]
    mm   = date_raw[3:5]
    yyyy = date_raw[6:10]

    if not (dd.isdigit() and mm.isdigit() and yyyy.isdigit()):
        return None

    try:
        return date(int(yyyy), int(mm), int(dd))
    except ValueError:
        return None


def format_date(ws, row_errors, col, **_):
    """
    Parse date text and store as an Excel date value with DD/MM/YYYY format.
    Cells that cannot be parsed are left untouched and logged to row_errors.

    Args:
        col: 1-based column index of the date column.
    """
    DATA_START = 2  # row 1 is the header

    for row in range(DATA_START, ws.max_row + 1):
        cell = ws.cell(row=row, column=col)
        if cell.value is None:
            continue
        parsed = _normalize_date(str(cell.value).strip())
        if parsed:
            cell.value = parsed
            cell.number_format = EXCEL_DATE_FORMAT
        else:
            row_errors.setdefault(row, {})[col] = str(cell.value)
