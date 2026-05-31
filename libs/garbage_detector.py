"""
Detects and removes garbage trailing rows produced by OCR mis-scanning page
footers/watermarks as data rows.

Detection rule: if the last data row has BOTH its transaction-datetime column
(col B) AND its effective-date column (col C) unparseable, the row is almost
certainly footer noise and is deleted before any formatter runs.
"""

from libs.formatters.datetime_formatter import _normalize
from libs.formatters.date_formatter import _normalize_date


def drop_garbage_tail(ws, trans_date_col, date_col=None):
    """
    Delete the last data row if all provided date columns are unparseable.

    Args:
        ws: openpyxl Worksheet (modified in place).
        trans_date_col: 1-based index of the transaction datetime column (or the
            only date column for formats that have just one, e.g. MB bank).
        date_col: 1-based index of the effective date column. Optional — omit for
            bank formats that have a single date column.

    Returns:
        True if a garbage row was removed, False otherwise.
    """
    last_row = ws.max_row
    if last_row < 2:  # nothing beyond the header
        return False

    if date_col is not None:
        trans_raw = ws.cell(row=last_row, column=trans_date_col).value
        date_raw  = ws.cell(row=last_row, column=date_col).value
        trans_ok = trans_raw is not None and _normalize(str(trans_raw).strip()) is not None
        date_ok  = date_raw  is not None and _normalize_date(str(date_raw).strip()) is not None
        should_drop = not trans_ok and not date_ok
    else:
        date_raw = ws.cell(row=last_row, column=trans_date_col).value
        date_ok  = date_raw is not None and _normalize_date(str(date_raw).strip()) is not None
        should_drop = not date_ok

    if should_drop:
        row_preview = [ws.cell(row=last_row, column=c).value for c in range(1, min(5, ws.max_column + 1))]
        print(f"  [GARBAGE] Row {last_row} dropped (both date columns invalid): {row_preview}")
        ws.delete_rows(last_row)
        return True

    return False
