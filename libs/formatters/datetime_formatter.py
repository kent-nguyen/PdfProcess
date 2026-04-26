"""
Formatter for datetime columns (Vietnamese locale: dd/MM/yyyy HH:mm:ss).
Stores values as proper Excel datetime objects with an explicit DD/MM/YYYY format
so display is locale-independent.

Handles common OCR errors by extracting date and time parts positionally
rather than relying on separators (which OCR often misreads).

Date algorithm:
  1. Concatenate token[0] + token[1] to reunite the split year, e.g.:
       "07/01/20" + "23"  -> "07/01/2023"
       "10/01020" + "23"  -> "10/0102023"  (0 is misread /)
  2. Extract by position — ignore whatever character sits at index 2 and 5:
       [0:2] = DD,  [3:5] = MM,  [6:10] = YYYY

Time algorithm:
  1. Join all remaining tokens with a space, e.g.:
       ["15851:24"]  -> "15851:24"
       ["11558", "24"] -> "11558 24"
  2. Extract by position — ignore whatever character sits at index 2 and 5:
       [0:2] = HH,  [3:5] = MM,  [6:8] = SS
"""
from datetime import datetime


EXCEL_DATETIME_FORMAT = "DD/MM/YYYY HH:MM:SS"


def _extract_date(s):
    """Extract (DD, MM, YYYY) from a concatenated date string by position."""
    if len(s) < 10:
        return None
    dd, mm, yyyy = s[0:2], s[3:5], s[6:10]
    if dd.isdigit() and mm.isdigit() and yyyy.isdigit():
        return dd, mm, yyyy
    return None


def _extract_time(s):
    """
    Extract (HH, MI, SS) from a joined time string.

    - 8 chars: one separator misread as another char (e.g. '8', '5', '.', ' ').
      Use positional extraction — ignore whatever sits at index 2 and 5.
    - 7 chars, 6 digits: one separator missing entirely (e.g. '1544:06', '15:4406').
      Extract the 6 digits and regroup as HH MM SS.
    - 7 chars, 7 digits: second colon was misread as a digit (e.g. '2122452' → '21:22:52').
      Ignore char at index 4 (the misread colon) and extract positionally.
    - 6 chars: both separators missing (e.g. '154406').
      Extract the 6 digits and regroup as HH MM SS.
    """
    if len(s) == 8:
        hh, mi, ss = s[0:2], s[3:5], s[6:8]
    elif len(s) == 7:
        digits = "".join(c for c in s if c.isdigit())
        if len(digits) == 6:
            hh, mi, ss = digits[0:2], digits[2:4], digits[4:6]
        elif len(digits) == 7:
            # All digits — char at index 4 is a misread colon, ignore it
            hh, mi, ss = s[0:2], s[2:4], s[5:7]
        else:
            return None
    elif len(s) == 6:
        digits = "".join(c for c in s if c.isdigit())
        if len(digits) != 6:
            return None
        hh, mi, ss = digits[0:2], digits[2:4], digits[4:6]
    else:
        return None

    if hh.isdigit() and mi.isdigit() and ss.isdigit():
        return hh, mi, ss
    return None


def _normalize(raw):
    tokens = raw.strip().split()

    # If token[0] already has a 4-digit year (length >= 10, e.g. "07/01/2023"),
    # don't consume token[1] as year suffix.
    if len(tokens) >= 2 and len(tokens[0]) < 10:
        date_raw = tokens[0] + tokens[1]
        time_tokens = tokens[2:]
    elif len(tokens) >= 1:
        date_raw = tokens[0]
        time_tokens = tokens[1:]
    else:
        return None

    time_raw = " ".join(time_tokens)

    date_parts = _extract_date(date_raw)
    time_parts = _extract_time(time_raw)

    if not date_parts or not time_parts:
        return None

    dd, mm, yyyy = date_parts
    hh, mi, ss = time_parts
    return f"{dd}/{mm}/{yyyy} {hh}:{mi}:{ss}"


def format_trans_date(ws, row_errors, col):
    """
    Normalize datetime text to "dd/MM/yyyy HH:mm:ss" (Vietnamese locale).
    Cells that cannot be parsed are left untouched and logged to row_errors.

    Args:
        col: 1-based column index of the Trans Date column.
    """
    DATA_START = 2  # row 1 is the header

    for row in range(DATA_START, ws.max_row + 1):
        cell = ws.cell(row=row, column=col)
        if cell.value is None:
            continue
        normalized = _normalize(str(cell.value).strip())
        if normalized:
            cell.value = datetime.strptime(normalized, "%d/%m/%Y %H:%M:%S")
            cell.number_format = EXCEL_DATETIME_FORMAT
        else:
            row_errors.setdefault(row, {})[col] = str(cell.value)
