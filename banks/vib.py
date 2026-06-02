from functools import partial
from libs.fixers.balance_fixer import fix_balance
from libs.formatters.datetime_formatter import format_trans_date
from libs.formatters.amount_formatter import format_amount
from libs.formatters.concat_formatter import format_concat
from libs.formatters.description_formatter import format_description
from libs.corrections.bank_names import BANK_NAMES

# Column indices (1-based) for VIB statement format
TRANS_DATE_COL = 1   # Column A — Ngày / Tran Date (DD/MM/YYYY HH:MM:SS)
SEQ_COL = 2          # Column B — Số CT / Seq No
OFFICER_COL = 3      # Column C — Mã NV / Officer ID/Ref
DESC_COL = 4         # Column D — Nội dung / Remarks
TRAN_TYPE_COL = 5    # Column E — MGD / Tran Type (NBCR, NADR, …)
DEBIT_COL = 6        # Column F — PS Nợ / Debit
CREDIT_COL = 7       # Column G — PS Có / Credit
BALANCE_COL = 8      # Column H — Số dư / Balance
CORR_ACC_COL = 9     # Column I — TK đối ứng / Corresponding A/C No
CORR_NAME_COL = 10   # Column J — Tên chủ TK đối ứng / Corresponding Name
CORR_BANK_COL = 11   # Column K — Ngân hàng TK đối ứng / Corresponding Bank

_ALNUM = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'

FORMATTERS = [
    partial(format_trans_date, col=TRANS_DATE_COL),
    partial(format_concat, col=SEQ_COL, as_text=True),
    partial(format_concat, col=OFFICER_COL, as_text=True),
    partial(format_description, col=DESC_COL, corrections=BANK_NAMES),
    partial(format_concat, col=TRAN_TYPE_COL),
    partial(format_amount, col=DEBIT_COL),
    partial(format_amount, col=CREDIT_COL),
    partial(format_amount, col=BALANCE_COL),
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
    0: '0123456789/: ',        # Tran Date — digits, slash, colon, space
    1: '0123456789',           # Seq No — digits only
    2: _ALNUM,                 # Officer ID — alphanumeric
    # 3: Remarks — unrestricted
    4: _ALNUM,                 # Tran Type — alphanumeric (NBCR, NADR, …)
    5: '0123456789,.',         # Debit amount
    6: '0123456789,.',         # Credit amount
    7: '0123456789,.',         # Balance amount
    8: '0123456789',           # Corresponding A/C No — digits only
    # 9:  Corresponding Name — unrestricted
    # 10: Corresponding Bank — unrestricted
}
