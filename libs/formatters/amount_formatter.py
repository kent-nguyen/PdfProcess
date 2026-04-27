"""
Formatter for VND currency amount columns.

OCR misreads dots and commas interchangeably, so separators cannot be trusted.
The trailing "00" decimal cents are the one reliable anchor.

Algorithm:
  1. Trim leading/trailing whitespace.
  2. Strip the trailing decimal digits: scan from the end collecting digits; if 1 or 2
     digits are found and the character before them is a separator (" ", ".", ","),
     remove that separator + digits (handles ".00", ".07", " 00", ",0", " 0", etc.).
     Valid thousand groups always have 3 digits, so 1-2 trailing digits after a
     separator are unambiguously a decimal suffix — real or OCR-misread.
  3. Fix two-zero thousand groups: any "00" sitting between two separators
     (e.g. ",00," or ",00.") has a dropped zero — restore it to "000".
  4. Fix four-digit groups: a separator misread as a digit gets appended to a
     thousand group, making it 4 digits (e.g. ",0003" where "3" is a misread ".").
     Remove the extra digit by keeping only the first three (e.g. "0003" → "000").
  5. Remove every remaining dot, comma, and space (thousands/decimal separators).
  6. Parse the remaining digit string as int.
"""
import re


def _strip_decimal(s):
    """Remove a trailing decimal suffix of 1-2 digits preceded by a separator.
    Valid thousand groups always have 3 digits, so 1-2 trailing digits after a
    separator can only be a decimal part (real .00 or OCR-misread .07 etc.)."""
    i = len(s)
    while i > 0 and s[i - 1].isdigit():
        i -= 1
    digit_count = len(s) - i
    if digit_count in (1, 2) and i > 0 and s[i - 1] in (' ', '.', ','):
        s = s[:i - 1].rstrip()
    return s


def _normalize_amount(raw):
    """Return (int_value, fix_note_or_None). fix_note is set when a correction was made."""
    s = raw.strip()               # step 1: trim only
    s = _strip_decimal(s)         # step 2: remove trailing decimal zeros
    before_fixes = s
    # Step 3: restore a dropped zero in any thousand group that OCR read as "00".
    s = re.sub(r'(?<=[,\.])00(?=[,\.]|$)', '000', s)
    # Step 4: remove a spurious extra digit from any 4-digit group — the last
    # digit is a separator (. or ,) that OCR misread as a numeral.
    s = re.sub(r'(?<=[,\.])(\d{4})(?=[,\.]|$)', lambda m: m.group(1)[:3], s)
    fix_note = f"Số tiền: đã sửa nhóm chữ số sai (OCR: {raw.strip()!r})" if s != before_fixes else None
    s = s.replace(".", "").replace(",", "").replace(" ", "")  # drop all separators

    if not s or not s.isdigit():
        return None, None
    return int(s), fix_note


def format_amount(ws, row_errors, col, row_fixes=None, **_):
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
        original = str(cell.value)
        parsed, fix_note = _normalize_amount(original)
        if parsed is not None:
            cell.value = parsed
            cell.number_format = "#,##0"
            if fix_note:
                if row_fixes is not None:
                    row_fixes.setdefault(row, []).append(fix_note)
                row_errors.setdefault(row, {})[col] = original
        else:
            row_errors.setdefault(row, {})[col] = str(cell.value)
