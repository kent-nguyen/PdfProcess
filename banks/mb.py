from functools import partial
from libs.fixers.balance_fixer import fix_balance
from libs.formatters.date_formatter import format_date
from libs.formatters.amount_formatter import format_amount
from libs.formatters.concat_formatter import format_concat
from libs.formatters.description_formatter import format_description
from libs.corrections.bank_names import BANK_NAMES

# Column indices (1-based) for MB bank statement format
DATE_COL = 1        # Column A — Ngày / Date
TRANS_NO_COL = 2    # Column B — Số bút toán / Transaction No
DEBIT_COL = 3       # Column C — Phát sinh nợ / Debit
CREDIT_COL = 4      # Column D — Phát sinh có / Credit
DESC_COL = 5        # Column E — Nội dung / Details
BENEF_COL = 6       # Column F — Đơn vị thụ hưởng / Beneficiary
ACC_COL = 7         # Column G — Tài khoản / Account
BANK_COL = 8        # Column H — Ngân hàng đối tác / Remitter Bank

_ALNUM = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'

FORMATTERS = [
    partial(format_date, col=DATE_COL),
    partial(format_concat, col=TRANS_NO_COL, as_text=True),
    partial(format_amount, col=DEBIT_COL),
    partial(format_amount, col=CREDIT_COL),
    partial(format_description, col=DESC_COL, corrections=BANK_NAMES),
    partial(format_description, col=BENEF_COL, corrections=BANK_NAMES),
    partial(format_concat, col=ACC_COL, as_text=True),
    partial(format_description, col=BANK_COL, corrections=BANK_NAMES),
]

FIXERS = []

# Columns used by the garbage-tail detector (transaction date only).
GARBAGE_DATE_COLS = (DATE_COL,)

# Per-column OCR allowlists (0-based column index from cell image filenames).
COLUMN_ALLOWLISTS = {
    0: '0123456789/ ',        # Date — digits, slash, space
    1: _ALNUM,                # Transaction No — alphanumeric
    2: '0123456789,.',        # Debit amount
    3: '0123456789,.',        # Credit amount
    # 4: Details — unrestricted
    # 5: Beneficiary — unrestricted
    6: '0123456789',          # Account No — digits only
    # 7: Remitter Bank — unrestricted
}
