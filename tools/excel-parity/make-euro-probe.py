#!/usr/bin/env python3
"""Probe cells for EUROCONVERT (FnFinancial.fs), as eval-cells-on-vm.sh takes them.

    python3 tools/excel-parity/make-euro-probe.py > euro-cells.txt

EUROCONVERT is a function of the Euro Currency Tools add-in: the first line opens EUROTOOL.XLAM in the Excel
of the VM (eval-cells.vbs's #ADDIN). Then one line per cell: sheet, address, formula. Every currency the
add-in might know to and from the euro and between them (triangulation), amounts that round at the half
cent and below it, full_precision TRUE and FALSE, triangulation_precision omitted and 2 to 8, the codes in
lower case and with spaces, unknown codes, text and logical amounts, and the errors. Deterministic (fixed
seed): euro-excel.tsv holds Excel 2019's values; function-parity-engine.fsx compares the engine.
"""
import random

R = random.Random(20261125)
cells = []


def cell(f):
    cells.append(f)


CODES = ['BEF', 'LUF', 'DEM', 'ESP', 'FRF', 'IEP', 'ITL', 'NLG', 'ATS', 'PTE', 'FIM', 'GRD', 'SIT', 'CYP', 'MTL', 'SKK',
         'EEK', 'LVL', 'LTL', 'HRK', 'EUR']
AMOUNTS = [1, 100, 1234.5678, 0.005, 0.015, 2.675, 1e9, -55.555]
for c in CODES:
    for amt in AMOUNTS:
        cell('=EUROCONVERT(%r,"%s","EUR")' % (amt, c))
        cell('=EUROCONVERT(%r,"EUR","%s")' % (amt, c))
    cell('=EUROCONVERT(1,"%s","EUR",TRUE)' % c)
    cell('=EUROCONVERT(1,"EUR","%s",TRUE)' % c)
for _ in range(150):
    a, b = R.sample(CODES[:13], 2)
    amt = R.choice([1, 100, 1000, R.uniform(0, 1e5), round(R.uniform(0, 1000), 2), R.uniform(0, 1)])
    full = R.choice(['', ',FALSE', ',TRUE'])
    tri = R.choice(['', '', ',3', ',4', ',5', ',6', ',8', ',2', ',1'])
    if tri and not full:
        full = R.choice([',FALSE', ',TRUE'])
    cell('=EUROCONVERT(%r,"%s","%s"%s%s)' % (amt, a, b, full, tri))
for _ in range(60):
    c = R.choice(CODES[:13])
    amt = R.uniform(0, 1e4)
    for full in ('FALSE', 'TRUE'):
        cell('=EUROCONVERT(%r,"%s","EUR",%s)' % (amt, c, full))
        cell('=EUROCONVERT(%r,"EUR","%s",%s)' % (amt, c, full))
        cell('=EUROCONVERT(%r,"EUR","%s",%s,3)' % (amt, c, full))
        cell('=EUROCONVERT(%r,"%s","EUR",%s,4)' % (amt, c, full))
for f in ['EUROCONVERT(1,"dem","eur")', 'EUROCONVERT(1," DEM","EUR")', 'EUROCONVERT(1,"DEM","USD")', 'EUROCONVERT(1,"XXX","EUR")',
          'EUROCONVERT(1,"EUR","EUR")', 'EUROCONVERT(1,"DEM","DEM")', 'EUROCONVERT(1,"DEM","DEM",TRUE)', 'EUROCONVERT("12","DEM","EUR")',
          'EUROCONVERT("x","DEM","EUR")', 'EUROCONVERT(TRUE,"DEM","EUR")', 'EUROCONVERT(1,"DEM","FRF",TRUE,"x")',
          'EUROCONVERT(1,"DEM","FRF","x")', 'EUROCONVERT(1,"DEM","FRF",1)', 'EUROCONVERT(1,"DEM","FRF",0)',
          'EUROCONVERT(1,"DEM","FRF",TRUE,3.7)', 'EUROCONVERT(1,"DEM","FRF",TRUE,-1)', 'EUROCONVERT(1,"DEM","FRF",TRUE,0)',
          'EUROCONVERT(1,"DEM","FRF",TRUE,15)', 'EUROCONVERT(1,"DEM","FRF",TRUE,100)', 'EUROCONVERT(1/0,"DEM","EUR")',
          'EUROCONVERT(1,1,"EUR")', 'EUROCONVERT(1,"DEM",TRUE)', 'EUROCONVERT(Z1,"DEM","EUR")', 'EUROCONVERT(1,Z1,"EUR")',
          'EUROCONVERT(1,"","EUR")', 'EUROCONVERT(1,"DEM","FRF",,3)', 'EUROCONVERT(1,"DEM","FRF",FALSE,3)',
          'EUROCONVERT(1,"EUR","DEM",FALSE,3)', 'EUROCONVERT(1E+300,"EUR","ITL")', 'EUROCONVERT(-0.004,"DEM","EUR")',
          'EUROCONVERT(1.995583,"DEM","EUR",TRUE,3)', 'EUROCONVERT(1,"DEM","FRF",TRUE,2)', 'EUROCONVERT({1,2},"DEM","EUR")']:
    cell('=' + f)
# Halves: which rounding (Excel's ROUND, decimal; VBA's Round, binary and to even).
for amt in (2.675, 0.125, 1.005, 0.5, 1.5, 2.5, 0.015, 1234.565):
    for pair in (('EUR', 'EUR'), ('DEM', 'DEM'), ('BEF', 'BEF'), ('EUR', 'BEF')):
        cell('=EUROCONVERT(%r,"%s","%s")' % ((amt,) + pair))
for k in (1, 3, 5, 25, 125, 2675):
    for c, rate in (('DEM', 1.95583), ('FRF', 6.55957), ('ITL', 1936.27)):
        cell('=EUROCONVERT(%s*%d/2000,"%s","NLG",TRUE,3)' % (rate, k, c))
        cell('=EUROCONVERT(%s*%d/2000,"%s","EUR",TRUE,3)' % (rate, k, c))
        cell('=EUROCONVERT(%d/2000,"EUR","%s",TRUE,3)' % (k, c))
        cell('=EUROCONVERT(%s*%d/200,"%s","EUR")' % (rate, k, c))

print('#ADDIN\tC:\\Program Files\\Microsoft Office\\root\\Office16\\Library\\EUROTOOL.XLAM')
for i, f in enumerate(cells):
    print('S\tA%d\t%s' % (i + 1, f))
