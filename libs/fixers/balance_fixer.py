"""
Fixer that validates each row's balance against the accounting equation:

    Balance = Previous Balance - Debit + Credit

Rows where the equation does not hold are marked in row_errors and a note is
added to row_fixes. No cell values are modified.

Validation is skipped for a row if any of the three key values (previous
balance, current debit, current credit) are not integers, which means the
formatter could not parse them.
"""


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
        # If debit or credit failed formatting (still a string), skip
        if debit is not None and not isinstance(debit, int):
            continue
        if credit is not None and not isinstance(credit, int):
            continue

        expected = prev_balance - (debit or 0) + (credit or 0)
        if expected != balance:
            row_errors.setdefault(row, {})[col] = (
                f"Balance mismatch: expected {expected:,}, got {balance:,}"
            )
