def format_description(ws, row_errors, col, corrections=()):
    """
    Normalize a free-text description column:
      1. Collapse all whitespace runs to a single space.
      2. Apply bank-specific (wrong, correct) substring corrections in order.

    Never marks errors — silently skips empty cells.

    Args:
        col:         1-based column index.
        corrections: sequence of (wrong, correct) string pairs.
    """
    DATA_START = 2

    for row in range(DATA_START, ws.max_row + 1):
        cell = ws.cell(row=row, column=col)
        if cell.value is None:
            continue
        s = " ".join(str(cell.value).split())
        for wrong, correct in corrections:
            s = s.replace(wrong, correct)
        cell.value = s
