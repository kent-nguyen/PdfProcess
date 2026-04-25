from functools import partial
from libs.fixers.stt_fixer import fix_stt
from libs.formatters.stt_formatter import format_stt

# Column indices (1-based) for BIDV statement format
STT_COL = 1  # Column A

FORMATTERS = [
    partial(format_stt, col=STT_COL),
]

FIXERS = [
    partial(fix_stt, col=STT_COL),
]
