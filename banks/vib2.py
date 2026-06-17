from functools import partial
from libs.fixers.balance_fixer import fix_balance
from libs.formatters.datetime_formatter import format_trans_date_ampm
from libs.formatters.amount_formatter import format_amount
from libs.formatters.concat_formatter import format_concat
from libs.formatters.description_formatter import format_description
from libs.corrections.bank_names import BANK_NAMES

# Column indices (1-based) for VIB2 statement format
# "BÁO CÁO HOẠT ĐỘNG TÀI KHOẢN GIAO DỊCH ĐỐI ỨNG"
CUSTOMER_CODE_COL = 1   # Column A — Mã khách hàng / Customer Code
CUSTOMER_NAME_COL = 2   # Column B — Tên khách hàng / Customer Name
CURRENCY_COL = 3        # Column C — Loại tiền / Currency
ACCOUNT_COL = 4         # Column D — Số tài khoản / Account No
TRANS_ID_COL = 5        # Column E — Số giao dịch/Id / Transaction ID
TRANS_DATE_COL = 6      # Column F — Ngày giờ giao dịch / Transaction Date/Time (DD/MM/YYYY HH:MM:SS AM/PM)
OPENING_BAL_COL = 7     # Column G — Số đầu / Opening Balance
DEBIT_COL = 8           # Column H — Phát sinh nợ / Debit
CREDIT_COL = 9          # Column I — Phát sinh có / Credit
BALANCE_COL = 10        # Column J — Số dư cuối / Closing Balance
DESC_COL = 11           # Column K — Nội dung giao dịch / Description
CORR_ACC_COL = 12       # Column L — Số TK đối ứng / Corresponding A/C No
CORR_NAME_COL = 13      # Column M — Tên chủ TK đối ứng / Corresponding Name
CORR_BANK_COL = 14      # Column N — Tên TCTD TK đối ứng / Corresponding Bank

_ALNUM = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'

FORMATTERS = [
    partial(format_concat, col=CUSTOMER_CODE_COL, as_text=True),
    partial(format_description, col=CUSTOMER_NAME_COL, corrections=BANK_NAMES),
    partial(format_concat, col=CURRENCY_COL),
    partial(format_concat, col=ACCOUNT_COL, as_text=True),
    partial(format_concat, col=TRANS_ID_COL, as_text=True),
    partial(format_trans_date_ampm, col=TRANS_DATE_COL),
    partial(format_amount, col=OPENING_BAL_COL),
    partial(format_amount, col=DEBIT_COL),
    partial(format_amount, col=CREDIT_COL),
    partial(format_amount, col=BALANCE_COL),
    partial(format_description, col=DESC_COL, corrections=BANK_NAMES),
    partial(format_concat, col=CORR_ACC_COL, as_text=True),
    partial(format_description, col=CORR_NAME_COL, corrections=BANK_NAMES),
    partial(format_description, col=CORR_BANK_COL, corrections=BANK_NAMES),
]

FIXERS = [
    partial(fix_balance, debit_col=DEBIT_COL, credit_col=CREDIT_COL, col=BALANCE_COL, lenient=True),
]

# Columns used by the garbage-tail detector (single datetime column).
GARBAGE_DATE_COLS = (TRANS_DATE_COL,)

# Per-column OCR allowlists (0-based column index from cell image filenames).
COLUMN_ALLOWLISTS = {
    0: _ALNUM,                        # Customer Code — alphanumeric
    # 1: Customer Name — unrestricted
    2: 'ABCDEFGHIJKLMNOPQRSTUVWXYZ',  # Currency — uppercase letters (VND)
    3: '0123456789',                  # Account No — digits only
    4: _ALNUM,                        # Transaction ID — alphanumeric
    5: '0123456789/: APMapm',         # Date/Time — digits, slash, colon, space, AM/PM
    6: '0123456789,.',                # Opening Balance
    7: '0123456789,.',                # Debit amount
    8: '0123456789,.',                # Credit amount
    9: '0123456789,.',                # Closing Balance
    # 10: Description — unrestricted
    11: '0123456789',                 # Corresponding A/C No — digits only
    # 12: Corresponding Name — unrestricted
    # 13: Corresponding Bank — unrestricted
}
