def format_int(ws, row_errors, col):
    """
    Convert column values from text to integer, silently skipping failures.

    Args:
        col: 1-based column index.
    """
    DATA_START = 2

    for row in range(DATA_START, ws.max_row + 1):
        cell = ws.cell(row=row, column=col)
        if cell.value is None:
            continue
        try:
            cell.value = int(str(cell.value).strip())
        except (ValueError, TypeError):
            pass
