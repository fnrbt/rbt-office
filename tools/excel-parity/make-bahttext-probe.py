#!/usr/bin/env python3
"""Probe cells for BAHTTEXT (FnText.fs), as eval-cells-on-vm.sh takes them.

    python3 tools/excel-parity/make-bahttext-probe.py > bahttext-cells.txt

One line per cell: sheet, address, formula. Whole numbers through the forms of the Thai number words (the
units after ten, eleven and the tens' "one", twenty, the hundreds to the millions and the repeated millions
of the large numbers), the satang (fractions, their rounding at the half satang and below it, the values
whose binary form is just under a half), negatives, zero, the coercions of the argument (text, logicals,
empty) and the errors. Deterministic (fixed seed): bahttext-excel.tsv holds Excel 2019's values;
function-parity-engine.fsx compares the engine.
"""
import random

R = random.Random(20261125)
cells = []


def cell(f):
    cells.append(f)


nums = list(range(0, 130)) + [200, 201, 210, 211, 221, 999, 1000, 1001, 1010, 1011, 1021, 1100, 1111, 10000, 10001, 10011,
                              100000, 100001, 1000000, 1000001, 1000011, 1000021, 1001000, 1010101, 2000000, 10000000,
                              11000000, 21000000, 100000000, 1000000000, 1000000000000, 1000001000000, 1000000000001,
                              11111111111111, 21212121212121, 999999999999999, 123456789012345, 1e15, 1e16, 2e18, 1e20,
                              1.5e21, 1e100, 1e300, 4294967295, 4294967296, 2147483647, 2147483648]
nums += [R.randint(0, 10 ** R.randint(1, 15)) for _ in range(120)]
for n in nums:
    cell('=BAHTTEXT(%r)' % n)
fracs = [0.01, 0.1, 0.11, 0.21, 0.25, 0.5, 0.99, 0.999, 0.995, 0.994, 0.005, 0.004, 0.0049999, 0.015, 0.025, 0.035, 0.045,
         0.105, 0.125, 0.135, 0.145, 0.155, 0.165, 0.175, 0.185, 0.195, 0.205, 0.215, 0.285, 0.565, 0.575, 0.585, 0.675,
         1e-10, 0.333333333333333, 0.666666666666667]
for f in fracs:
    for whole in (0, 1, 21, 100, 1234567, 1000000000):
        cell('=BAHTTEXT(%r)' % (whole + f))
for _ in range(80):
    cell('=BAHTTEXT(%r)' % round(R.uniform(0, 10 ** R.randint(0, 12)), R.randint(1, 4)))
for n in [-1, -0.5, -21.25, -1000000, -0.004, -0.005, -1234.567, -1e15, -0.01]:
    cell('=BAHTTEXT(%r)' % n)
for f in ['BAHTTEXT("12")', 'BAHTTEXT("12.5")', 'BAHTTEXT("x")', 'BAHTTEXT("")', 'BAHTTEXT(TRUE)', 'BAHTTEXT(FALSE)',
          'BAHTTEXT(A1000)', 'BAHTTEXT(1/0)', 'BAHTTEXT(NA())', 'BAHTTEXT("1,000")', 'BAHTTEXT("$5")', 'BAHTTEXT("1e3")',
          'BAHTTEXT({1,2})', 'BAHTTEXT(DATE(2020,1,1))', 'BAHTTEXT(9.99999999999999E+307)', 'BAHTTEXT(0.0000001)',
          'BAHTTEXT(-0)', 'LEN(BAHTTEXT(1E+300))']:
    cell('=' + f)

for i, f in enumerate(cells):
    print('S\tA%d\t%s' % (i + 1, f))
