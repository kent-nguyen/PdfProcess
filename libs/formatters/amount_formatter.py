"""
Formatter for VND currency amount columns.

OCR misreads dots and commas interchangeably, so separators cannot be trusted.
The trailing "00" decimal cents are the one reliable anchor.

Algorithm:
  1. Join all whitespace-separated tokens (reunites tokens split across lines).
  2. Strip the trailing "00" (the fixed decimal cents).
  3. Fix two-zero thousand groups: any "00" sitting between two separators
     (e.g. ",00," or ",00.") has a dropped zero — restore it to "000".
  4. Remove every remaining dot and comma (thousands/decimal separators).
  5. Parse the remaining digit string as int.
"""
import re


def _normalize_amount(raw):
    """Return (int_value, fix_note_or_None). fix_note is set when a dropped zero was restored."""
    s = "".join(raw.strip().split())  # reunite tokens, drop all spaces
    if s.endswith("00"):
        s = s[:-2]                    # strip fixed decimal cents
    # Restore a dropped zero in any thousand group that OCR read as "00":
    # lookbehind/lookahead don't consume the separators so all groups are
    # found and fixed in a single pass (e.g. "5,00,00." → "5,000,000.").
    fixed = re.sub(r'(?<=[,\.])00(?=[,\.])', '000', s)
    fix_note = None
    if fixed != s:
        fix_note = f"Số tiền: khôi phục chữ số 0 bị thiếu (OCR: {raw.strip()!r})"
    s = fixed.replace(".", "").replace(",", "")  # drop all separators

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
