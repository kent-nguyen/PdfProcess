"""
Fixer for sequential serial-number (STT) columns.
"""


def _int_or_none(value):
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        return None


def fix_stt(ws, row_fixes, row_errors, col):
    """
    Fix empty or malformed cells in a sequential serial-number column.

    Reads row_errors (populated by the formatter) to detect cells that had
    bad values (e.g. 'S2') and includes the original value in fix notes.

    Strategy:
      - Fill by incrementing from the previous known value.
      - If the result disagrees with the next known value, append a warning.
      - If no previous value exists, decrement back from the next known value.
      - If neither neighbour is known, flag for manual review.

    Args:
        col: 1-based column index of the STT column.
    """
    DATA_START = 2  # row 1 is the header

    known = {}
    for row in range(DATA_START, ws.max_row + 1):
        v = _int_or_none(ws.cell(row=row, column=col).value)
        if v is not None:
            known[row] = v

    empty = [r for r in range(DATA_START, ws.max_row + 1) if r not in known]
    if not empty:
        return

    groups = []
    group = [empty[0]]
    for r in empty[1:]:
        if r == group[-1] + 1:
            group.append(r)
        else:
            groups.append(group)
            group = [r]
    groups.append(group)

    for group in groups:
        first, last = group[0], group[-1]
        prev_val = known.get(first - 1)
        next_val = known.get(last + 1)

        if prev_val is not None:
            filled = [prev_val + i + 1 for i in range(len(group))]
            for row, val in zip(group, filled):
                ws.cell(row=row, column=col).value = val
                bad_val = row_errors.get(row, {}).get(col)
                origin = f"bad value '{bad_val}'" if bad_val else "empty"
                note = f"STT: was {origin}, filled as {val} (incremented from {prev_val})"
                if next_val is not None and filled[-1] + 1 != next_val:
                    note += f" [WARNING: sequence break — next known is {next_val}]"
                row_fixes.setdefault(row, []).append(note)

        elif next_val is not None:
            filled = [next_val - (len(group) - i) for i in range(len(group))]
            for row, val in zip(group, filled):
                ws.cell(row=row, column=col).value = val
                bad_val = row_errors.get(row, {}).get(col)
                origin = f"bad value '{bad_val}'" if bad_val else "empty"
                row_fixes.setdefault(row, []).append(
                    f"STT: was {origin}, filled as {val} (decremented from {next_val})"
                )

        else:
            for row in group:
                bad_val = row_errors.get(row, {}).get(col)
                origin = f"bad value '{bad_val}'" if bad_val else "empty"
                row_fixes.setdefault(row, []).append(
                    f"STT: was {origin}, no adjacent values to interpolate — manual review needed"
                )
