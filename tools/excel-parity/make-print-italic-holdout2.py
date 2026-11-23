#!/usr/bin/env python3
"""Second blind fixture for the print italic reserve: a fresh face/size/scale set for the revised
`round(usWinAscent x heightDots / upm)` variant, written before its Excel export.
    python3 tools/excel-parity/make-print-italic-holdout2.py OUT.xlsx
"""
import importlib.util, os, sys
spec = importlib.util.spec_from_file_location('holdout1', os.path.join(os.path.dirname(os.path.abspath(__file__)), 'make-print-italic-holdout.py'))
h = importlib.util.module_from_spec(spec); spec.loader.exec_module(h)
SHEETS = [
    ('H15_tahoma16_80_upright', 'Tahoma', 16, False, 10, 'right', 18, [-0.9841, -1, 0.5, 1234567], 80),
    ('H16_tahoma16_80_italic',  'Tahoma', 16, True,  10, 'right', 18, [-0.9841, -1, 0.5, 1234567], 80),
    ('H17_garamond11_115_upright', 'Garamond', 11, False, 204, 'right', 18, [-0.9841, -1, 0.5, 1234567], 115),
    ('H18_garamond11_115_italic',  'Garamond', 11, True,  204, 'right', 18, [-0.9841, -1, 0.5, 1234567], 115),
    ('H19_verdana9_110_upright', 'Verdana', 9, False, 200, 'right', 18, [-0.9841, -1, 0.5, 1234567], 110),
    ('H20_verdana9_110_italic',  'Verdana', 9, True,  200, 'right', 18, [-0.9841, -1, 0.5, 1234567], 110),
    ('H21_calibri10p5_95_upright', 'Calibri', 10.5, False, 10, 'right', 18, [-0.9841, -1, 0.5, 1234567], 95),
    ('H22_calibri10p5_95_italic',  'Calibri', 10.5, True,  10, 'right', 18, [-0.9841, -1, 0.5, 1234567], 95),
]
h.SHEETS = SHEETS
if __name__ == '__main__':
    h.build(sys.argv[1])
