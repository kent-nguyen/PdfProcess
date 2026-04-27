"""
Formatter for VND currency amount columns.

OCR misreads dots and commas interchangeably, so separators cannot be trusted.
The trailing "00" decimal cents are the one reliable anchor.

Algorithm:
  1. Join all whitespace-separated tokens (reunites tokens split across lines).
  2. Strip the trailing "00" (the fixed decimal cents).
  3. Remove every remaining dot and comma (thousands/decimal separators).
  4. Parse the remaining digit string as int.
"""


def _normalize_amount(raw):
    s = "".join(raw.strip().split())  # reunite tokens, drop all spaces
    if s.endswith("00"):
        s = s[:-2]                    # strip fixed decimal cents
    s = s.replace(".", "").replace(",", "")  # drop all separators

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
