#!/usr/bin/env python3
"""Discriminating controls for the numeric-format path that suppresses Excel's printed italic reserve.

Each case is an upright/italic pair of sheets with a numeric cell and a literal text cell showing the same word, so
the numeric-vs-literal reserve and the italic-vs-upright reserve can be separated. The codes vary only the pieces that
could be responsible: a trailing `_)` pad, the same pad quoted, a leading `_(`, a `*` fill, both, an inactive-section
pad, and standard accounting.

    python3 tools/excel-parity/make-print-italic-format-controls.py OUT.xlsx
"""
import importlib.util, os, sys
HERE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location('controls', os.path.join(HERE, 'make-print-right-edge-controls.py'))
c = importlib.util.module_from_spec(spec); spec.loader.exec_module(c)

PAD_ONLY = '0.00_)'
PAD_QUOTED = '0.00"_)"'
STAR_ONLY = '* 0.00'
STAR_PAD = '_(* 0.00_)'
LEAD_PAD = '_(0.00)'
INACTIVE = '0.00;[Red]-0.00_)'
ACCT = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'

NUMFMTS = [(300, PAD_ONLY), (301, PAD_QUOTED), (302, STAR_ONLY), (303, STAR_PAD), (304, LEAD_PAD), (305, INACTIVE), (306, ACCT)]
FMT_ID = {PAD_ONLY: 300, PAD_QUOTED: 301, STAR_ONLY: 302, STAR_PAD: 303, LEAD_PAD: 304, INACTIVE: 305, ACCT: 306}

def rows(fmt, positive_text='0.50', negative_text=None, value=0.5):
    out = [c.num_row(value, positive_text, fmt, 'pos'), c.lit_row(positive_text, fmt, 'pos')]
    if negative_text is not None:
        out += [c.num_row(-value, negative_text, fmt, 'neg'), c.lit_row(negative_text, fmt, 'neg')]
    return out

CASES = [
    ('F01_pad', rows(PAD_ONLY)),
    ('F02_pad_quoted', rows(PAD_QUOTED, '0.50_)')),
    ('F03_star', rows(STAR_ONLY)),
    ('F04_star_pad', rows(STAR_PAD)),
    ('F05_lead_pad', rows(LEAD_PAD)),
    ('F06_inactive_pos', rows(INACTIVE, '0.50', '-0.50')),
    ('F07_accounting', rows(ACCT, '0.50', '(0.50)')),
]
SHEETS = []
for name, rs in CASES:
    SHEETS.append((name + '_upright', 'Aptos Narrow', 14, False, 22, 100, rs))
    SHEETS.append((name + '_italic', 'Aptos Narrow', 14, True, 22, 100, rs))
c.SHEETS = SHEETS
c.NUMFMTS = NUMFMTS
c.FMT_ID = FMT_ID
if __name__ == '__main__':
    c.build(sys.argv[1])
