"""
Formatter for VND currency amount columns.

Raw OCR values use commas as thousands separators and always end with ".00",
e.g. "1,234,567.00". The ".00" suffix may be split to a new line and OCR'd
as a separate token.

Algorithm:
  1. Join all whitespace-separated tokens (reunites the split "00").
  2. Remove all commas (thousands separators) and periods.
  3. Strip the last two characters (always the "00" decimal digits).
  4. Parse the remaining digit string as int.
"""


def _normalize_amount(raw):
    s = "".join(raw.strip().split())      # reunite tokens, drop all spaces
    s = s.replace(",", "").replace(".", "") # remove separators
    s = s[:-2]                             # strip trailing "00"

    if not s or not s.isdigit():
        return None
    return int(s)


def format_amount(ws, row_errors, col):
    """
    Parse VND amount text and store as an integer.
    Empty cells are skipped silently. Cells that cannot be parsed are left
    untouched and logged to row_errors for manual review.

    Args:
        col: 1-based column index of the amount column.
    """
    DATA_START = 2

    for row in range(DATA_START, ws.max_row + 1):
        cell = ws.cell(row=row, column=col)
        if cell.value is None:
            continue
        parsed = _normalize_amount(str(cell.value))
        if parsed is not None:
            cell.value = parsed
            cell.number_format = "#,##0"
        else:
            row_errors.setdefault(row, {})[col] = str(cell.value)
