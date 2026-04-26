"""
Fixer that validates each row's balance against the accounting equation:

    Balance = Previous Balance - Debit + Credit

When the equation does not hold, the fixer attempts to recover by removing
one digit at a time from the OCR'd balance and checking whether the result
matches the expected value. If a single-digit removal produces a match, the
cell is corrected in place and the row is marked fixed. If no removal
produces a match, the row is marked as an error only.

Validation is skipped for a row if any of the three key values (previous
balance, current debit, current credit) are not integers, which means the
formatter could not parse them.
"""


def _try_remove_one_digit(balance: int, expected: int):
    """Return (corrected_int, position, removed_digit) if removing exactly one
    digit from balance yields expected, otherwise return None."""
    s = str(balance)
    for i, digit in enumerate(s):
        candidate_str = s[:i] + s[i + 1:]
        if not candidate_str:
            continue
        # Avoid leading zeros creating a shorter number unintentionally
        candidate = int(candidate_str)
        if candidate == expected:
            return candidate, i, digit
    return None


def fix_balance(ws, row_fixes, row_errors, debit_col, credit_col, col):
    """
    Args:
        debit_col:  1-based column index of the Debit column.
        credit_col: 1-based column index of the Credit column.
        col:        1-based column index of the Balance column.
    """
    DATA_START = 2

    for row in range(DATA_START + 1, ws.max_row + 1):
        prev_balance = ws.cell(row=row - 1, column=col).value
        debit        = ws.cell(row=row, column=debit_col).value
        credit       = ws.cell(row=row, column=credit_col).value
        balance      = ws.cell(row=row, column=col).value

        # Require prev_balance and balance to be formatted integers
        if not isinstance(prev_balance, int) or not isinstance(balance, int):
            continue
        # Blank OCR cells arrive as "" — treat the same as None (i.e. zero)
        if debit == "":
            debit = None
        if credit == "":
            credit = None
        # If debit or credit failed formatting (non-empty string), skip
        if debit is not None and not isinstance(debit, int):
            continue
        if credit is not None and not isinstance(credit, int):
            continue

        expected = prev_balance - (debit or 0) + (credit or 0)
        if expected == balance:
            continue

        # Try to recover by removing one OCR'd digit
        fix = _try_remove_one_digit(balance, expected)
        if fix:
            corrected, pos, digit = fix
            ws.cell(row=row, column=col).value = corrected
            row_errors.setdefault(row, {})[col] = (
                f"Số dư không khớp: kỳ vọng {expected:,}, OCR đọc được {balance:,}"
            )
            row_fixes.setdefault(row, []).append(
                f"Số dư: đã xóa chữ số thừa '{digit}' tại vị trí {pos} "
                f"(OCR đọc được {balance:,}, đã sửa thành {corrected:,})"
            )
        else:
            ws.cell(row=row, column=col).value = expected
            row_errors.setdefault(row, {})[col] = (
                f"Số dư không khớp: kỳ vọng {expected:,}, OCR đọc được {balance:,}"
            )
            row_fixes.setdefault(row, []).append(
                f"Khôi phục số dư bằng cách ghi đè giá trị kỳ vọng {expected:,} "
                f"(OCR đọc được {balance:,}, đã sửa thành {expected:,})"
            )
