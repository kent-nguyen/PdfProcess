"""
Formatter for VND currency amount columns.

Raw OCR values use commas as thousands separators and end with ".00",
e.g. "1,234,567.00". The ".00" suffix may be split to a new line and OCR'd
as a separate token (e.g. "1,234,567 00"). OCR may also read only one decimal
digit (".0" instead of ".00").

Algorithm:
  1. Join all whitespace-separated tokens (reunites split decimal tokens).
  2. Remove commas (thousands separators).
  3. If a decimal point is present, keep only the integer part (handles both
     ".00" and ".0" OCR variants).
  4. Otherwise assume the last two characters are the space-joined "00" decimal
     digits and strip them — unless the result would be empty (bare "0").
  5. Parse the remaining digit string as int.
"""


def _normalize_amount(raw):
    s = "".join(raw.strip().split())  # reunite tokens, drop all spaces
    s = s.replace(",", "")            # remove thousands separators
    if "." in s:
        s = s.split(".")[0]           # strip decimal portion (.00 or .0)
    else:
        # space-joined decimal digits appended at end — strip trailing "00"
        stripped = s[:-2]
        if stripped:                  # keep bare "0" as-is
            s = stripped

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
