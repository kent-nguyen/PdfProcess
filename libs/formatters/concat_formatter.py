def format_concat(ws, row_errors, col, as_text=False, **_):
    """
    Concatenate whitespace-separated OCR tokens into a single string.
    Silently skips empty cells. No error marking.

    Args:
        col: 1-based column index.
        as_text: if True, always store as string with text number format
                 (prevents Excel from reformatting long digit strings as numbers).
    """
    DATA_START = 2

    for row in range(DATA_START, ws.max_row + 1):
        cell = ws.cell(row=row, column=col)
        if cell.value is None:
            continue
        joined = "".join(str(cell.value).split())
        if as_text:
            cell.value = joined
            cell.number_format = "@"
        else:
            try:
                cell.value = int(joined)
            except (ValueError, TypeError):
                cell.value = joined
