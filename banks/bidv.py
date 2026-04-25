from functools import partial
from libs.fixers.stt_fixer import fix_stt
from libs.formatters.stt_formatter import format_stt
from libs.formatters.datetime_formatter import format_trans_date
from libs.formatters.date_formatter import format_date
from libs.formatters.int_formatter import format_int
from libs.formatters.amount_formatter import format_amount

# Column indices (1-based) for BIDV statement format
STT_COL = 1        # Column A
TRANS_DATE_COL = 2  # Column B
DATE_COL = 3        # Column C
INT_COL_4 = 4      # Column D
DEBIT_COL = 5      # Column E
CREDIT_COL = 6     # Column F

FORMATTERS = [
    partial(format_stt, col=STT_COL),
    partial(format_trans_date, col=TRANS_DATE_COL),
    partial(format_date, col=DATE_COL),
    partial(format_int, col=INT_COL_4),
    partial(format_amount, col=DEBIT_COL),
    partial(format_amount, col=CREDIT_COL),
]

FIXERS = [
    partial(fix_stt, col=STT_COL),
]

# Per-column OCR allowlists (0-based column index from cell image filenames).
# Columns not listed here use unrestricted OCR.
COLUMN_ALLOWLISTS = {
    0: '0123456789',        # STT (No) — integers only
    1: '0123456789/: ',     # Trans Date — digits, slash, colon, space
    2: '0123456789/ ',      # Date — digits, slash, space
    3: '0123456789',        # Column D — integers only
    4: '0123456789,.',      # Debit amount
    5: '0123456789,.',      # Credit amount
}
