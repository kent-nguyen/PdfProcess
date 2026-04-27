"""
Formatter for sequential serial-number (STT) columns.
Converts text values to integers.
Populates row_errors when a value cannot be converted, so fixers can act on it.
"""


def format_stt(ws, row_errors, col, **_):
    """
    Convert all STT column values from text to integer.
    On failure, records the bad value in row_errors and leaves the cell untouched.

    Args:
        col: 1-based column index of the STT column.
    """
    DATA_START = 2  # row 1 is the header

    for row in range(DATA_START, ws.max_row + 1):
        cell = ws.cell(row=row, column=col)
        if cell.value is None:
            continue
        parts = str(cell.value).strip().split()
        if len(parts) >= 2 and all(p.isdigit() for p in parts):
            raise ValueError(
                f"STT cột {col}, dòng {row}: phát hiện hai số '{cell.value}' — "
                f"SplitTableCells có thể đã bỏ sót một đường kẻ ngang. "
                f"Hãy kiểm tra và sửa file raw thủ công."
            )
        try:
            cell.value = int(parts[0])
        except (ValueError, TypeError):
            row_errors.setdefault(row, {})[col] = str(cell.value)
