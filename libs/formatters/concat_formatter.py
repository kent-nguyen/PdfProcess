def format_concat(ws, row_errors, col):
    """
    Concatenate whitespace-separated OCR tokens into a single string.
    Silently skips empty cells. No error marking.

    Args:
        col: 1-based column index.
    """
    DATA_START = 2

    for row in range(DATA_START, ws.max_row + 1):
        cell = ws.cell(row=row, column=col)
        if cell.value is None:
            continue
        joined = "".join(str(cell.value).split())
        try:
            cell.value = int(joined)
        except (ValueError, TypeError):
            cell.value = joined
