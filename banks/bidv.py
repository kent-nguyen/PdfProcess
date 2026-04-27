from functools import partial
from libs.fixers.stt_fixer import fix_stt
from libs.fixers.balance_fixer import fix_balance
from libs.formatters.stt_formatter import format_stt
from libs.formatters.datetime_formatter import format_trans_date
from libs.formatters.date_formatter import format_date
from libs.formatters.int_formatter import format_int
from libs.formatters.amount_formatter import format_amount
from libs.formatters.concat_formatter import format_concat
from libs.formatters.description_formatter import format_description
from libs.corrections.bank_names import BANK_NAMES

# Column indices (1-based) for BIDV statement format
STT_COL = 1        # Column A
TRANS_DATE_COL = 2  # Column B
DATE_COL = 3        # Column C
INT_COL_4 = 4      # Column D
DEBIT_COL = 5      # Column E
CREDIT_COL = 6     # Column F
BALANCE_COL = 7    # Column G
SEQ_COL = 8        # Column H — Sequence No
TELLER_COL = 9     # Column I — Teller ID
BRANCH_COL = 10    # Column J — Branch ID
DESC_COL = 11      # Column K — Txn Description
CORR_ACC_COL = 12  # Column L — Correspondent Account No
CORR_BANK_COL = 14 # Column N — Correspondent Bank Name

_ALNUM = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'

# Common OCR misreadings in BIDV transaction descriptions.
# Longer/more specific patterns are listed first to avoid partial conflicts.
BIDV_DESC_CORRECTIONS = [
    # MBTKThe misreadings
    ("MB TKTne", "MBTKThe"),
    ("MBTKThc",  "MBTKThe"),
    ("MBTKne",   "MBTKThe"),
    ("MbTKInc",  "MBTKThe"),
    ("MBTKTne",  "MBTKThe"),
    ("MBIKIhe",  "MBTKThe"),
    ("MBTKTnc",  "MBTKThe"),
    # Tfr A/c misreadings
    ("Tfr Nc",   "Tfr A/c"),
    ("Tir Nc",   "Tfr A/c"),
    # TKThe misreadings
    ("TKTne",    "TKThe"),
    ("Tkine",    "TKThe"),
    ("Tkinc",    "TKThe"),
    ("IkInc",    "TKThe"),
    ("TKThc",    "TKThe"),
    ("TKTnc",    "TKThe"),
    ("TKIne",    "TKThe"),
]

FORMATTERS = [
    partial(format_stt, col=STT_COL),
    partial(format_trans_date, col=TRANS_DATE_COL),
    partial(format_date, col=DATE_COL),
    partial(format_int, col=INT_COL_4),
    partial(format_amount, col=DEBIT_COL),
    partial(format_amount, col=CREDIT_COL),
    partial(format_amount, col=BALANCE_COL),
    partial(format_concat, col=SEQ_COL, as_text=True),
    partial(format_concat, col=TELLER_COL, as_text=True),
    partial(format_concat, col=BRANCH_COL),
    partial(format_description, col=DESC_COL, corrections=BIDV_DESC_CORRECTIONS + BANK_NAMES),
    partial(format_concat, col=CORR_ACC_COL, as_text=True),
    partial(format_description, col=CORR_BANK_COL, corrections=BIDV_DESC_CORRECTIONS + BANK_NAMES),
]

FIXERS = [
    partial(fix_stt, col=STT_COL),
    partial(fix_balance, debit_col=DEBIT_COL, credit_col=CREDIT_COL, col=BALANCE_COL),
]

# Per-column OCR allowlists (0-based column index from cell image filenames).
# Columns not listed here use unrestricted OCR.
COLUMN_ALLOWLISTS = {
    0: '0123456789',          # STT (No) — integers only
    1: '0123456789/: ',       # Trans Date — digits, slash, colon, space
    2: '0123456789/ ',        # Date — digits, slash, space
    3: '0123456789',          # Column D — integers only
    4: '0123456789,.',        # Debit amount
    5: '0123456789,.',        # Credit amount
    6: '0123456789,.',        # Balance amount
    7: '0123456789',          # Sequence No — digits only
    8: _ALNUM,                # Teller ID — alphanumeric
    9: '0123456789',          # Branch ID — digits only
    10: _ALNUM + '/: ',       # Txn Description — alphanumeric + / : space
    11: '0123456789',          # Correspondent Account No — digits only
}
