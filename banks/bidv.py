from functools import partial
from libs.fixers.stt_fixer import fix_stt
from libs.formatters.stt_formatter import format_stt
from libs.formatters.datetime_formatter import format_trans_date

# Column indices (1-based) for BIDV statement format
STT_COL = 1        # Column A
TRANS_DATE_COL = 2  # Column B

FORMATTERS = [
    partial(format_stt, col=STT_COL),
    partial(format_trans_date, col=TRANS_DATE_COL),
]

FIXERS = [
    partial(fix_stt, col=STT_COL),
]

# Per-column OCR allowlists (0-based column index from cell image filenames).
# Columns not listed here use unrestricted OCR.
COLUMN_ALLOWLISTS = {
    0: '0123456789',        # STT (No) — integers only
    1: '0123456789/: ',     # Trans Date — digits, slash, colon, space
}
