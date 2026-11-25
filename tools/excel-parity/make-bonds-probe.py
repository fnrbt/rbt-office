#!/usr/bin/env python3
"""Probe cells for the coupon-bond functions (FnBonds.fs), as eval-cells-on-vm.sh takes them.

    python3 tools/excel-parity/make-bonds-probe.py > bonds-cells.txt

One line per cell: sheet, address, formula. The cells cover the COUP* family, PRICE, YIELD, DURATION,
MDURATION, ACCRINT, the ODD* functions and AMORDEGRC/AMORLINC over the five bases and frequencies 1, 2
and 4: settlements on, just before and just after coupon dates, end-of-month maturities (February 28
and 29 among them), maturities on days a short month lacks, and invalid inputs (#NUM!, #VALUE!, the
truncation of frequency and basis). Then the structure sweeps the rules of FnBonds.fs were found on:
COUPDAYS of basis 1 over 1608 settlements, ACCRINT around one first interest date, ODDFPRICE and
ODDLPRICE in every quasi-coupon period of their odd periods (zero rates and zero yields apart), AMORDEGRC
and AMORLINC over lives of 0.8 to 20 years. Deterministic (fixed seeds): bonds-excel.tsv holds Excel
2019's values; function-parity-engine.fsx compares the engine with them.
"""
import calendar
import datetime as dt
import random

R = random.Random(20261124)
cells = []


def d(x):
    return 'DATE(%d,%d,%d)' % (x.year, x.month, x.day)


def add_months(x, n, last=False):
    m = x.month - 1 + n
    y, m = x.year + m // 12, m % 12 + 1
    days = calendar.monthrange(y, m)[1]
    return dt.date(y, m, days if last else min(x.day, days))


def cell(formula):
    cells.append(formula)


MATURITIES = [dt.date(*t) for t in [
    (2030, 2, 28), (2028, 2, 29), (2030, 6, 30), (2030, 8, 31), (2029, 11, 15), (2030, 1, 31),
    (2030, 3, 30), (2030, 5, 31), (2030, 12, 31), (2031, 4, 30), (2030, 8, 30), (2030, 2, 15),
    (2032, 2, 29), (2030, 9, 29), (2030, 11, 30), (2031, 3, 31), (2029, 8, 29), (2030, 5, 30)]]


def settlements(mat):
    """Dates before `mat`: coupon dates (each frequency's), a day either side, month ends, random."""
    out = set()
    last = mat.day == calendar.monthrange(mat.year, mat.month)[1]
    for k in (1, 2, 3, 5, 7, 12, 13, 25, 41):
        c = add_months(mat, -3 * k, last)
        out.update([c, c - dt.timedelta(1), c + dt.timedelta(1)])
    for y in (2027, 2028, 2029):
        out.add(dt.date(y, 2, 28))
        out.add(dt.date(y, 3, 1))
    out.add(dt.date(2028, 2, 29))
    out.add(dt.date(2028, 8, 31))
    out.add(dt.date(2028, 8, 30))
    out.add(mat - dt.timedelta(1))
    for _ in range(4):
        out.add(mat - dt.timedelta(R.randint(2, 3000)))
    return sorted(x for x in out if x < mat and x.year >= 2018)


pairs = []
for mat in MATURITIES:
    s = settlements(mat)
    R.shuffle(s)
    pairs += [(x, mat) for x in s[:7]]

# The COUP* family. The dates and the count do not depend on the basis: all frequencies, a random basis;
# the day counts: all bases, one frequency.
for settle, mat in pairs:
    for f in (1, 2, 4):
        b = R.randint(0, 4)
        for fn in ('COUPNCD', 'COUPPCD', 'COUPNUM'):
            cell('=%s(%s,%s,%d,%d)' % (fn, d(settle), d(mat), f, b))
    for f in R.sample((1, 2, 4), 1):
        for b in range(5):
            for fn in ('COUPDAYBS', 'COUPDAYS', 'COUPDAYSNC'):
                cell('=%s(%s,%s,%d,%d)' % (fn, d(settle), d(mat), f, b))

# PRICE, YIELD, DURATION, MDURATION.
for settle, mat in pairs:
    for _ in range(2):
        f = R.choice((1, 2, 4))
        b = R.randint(0, 4)
        rate = R.choice((0, 0.0275, 0.0575, 0.08, 0.125))
        yld = R.choice((0, 0.01, 0.065, 0.0432, 0.15))
        red = R.choice((100, 100, 105, 98.5))
        cell('=PRICE(%s,%s,%r,%r,%r,%d,%d)' % (d(settle), d(mat), rate, yld, red, f, b))
        pr = R.choice((95.04287, 100, 101.25, 88.3, 112.5))
        cell('=YIELD(%s,%s,%r,%r,%r,%d,%d)' % (d(settle), d(mat), rate, pr, red, f, b))
        cell('=DURATION(%s,%s,%r,%r,%d,%d)' % (d(settle), d(mat), rate or 0.04, yld, f, b))
        cell('=MDURATION(%s,%s,%r,%r,%d,%d)' % (d(settle), d(mat), rate or 0.04, yld, f, b))

# ACCRINT: settlement before, on and after the first interest date; both calculation methods.
for _ in range(110):
    issue = dt.date(2020, 1, 1) + dt.timedelta(R.randint(0, 1500))
    first = add_months(issue, R.choice((1, 2, 3, 4, 6, 7, 11, 12, 14)))
    if R.random() < 0.3:
        first = add_months(first, 0, True)
    pick = R.random()
    if pick < 0.35:
        settle = issue + dt.timedelta(R.randint(1, max(1, (first - issue).days)))
    else:
        settle = first + dt.timedelta(R.randint(-3, 1200))
    f = R.choice((1, 2, 4))
    b = R.randint(0, 4)
    rate = R.choice((0.05, 0.1, 0.0375))
    par = R.choice((1000, 100, 5000))
    method = R.choice(('', ',TRUE', ',FALSE'))
    cell('=ACCRINT(%s,%s,%s,%r,%r,%d,%d%s)' % (d(issue), d(first), d(settle), rate, par, f, b, method))
    if method == '':
        cell('=ACCRINT(%s,%s,%s,%r,%r,%d,%d,FALSE)' % (d(issue), d(first), d(settle), rate, par, f, b))

# ODDFPRICE / ODDFYIELD: short and long first periods; ODDLPRICE / ODDLYIELD: short and long last periods.
for _ in range(90):
    f = R.choice((1, 2, 4))
    b = R.randint(0, 4)
    mat = R.choice(MATURITIES)
    last = mat.day == calendar.monthrange(mat.year, mat.month)[1]
    n = R.randint(1, 14)
    first = add_months(mat, -n * 12 // f, last)
    span = R.choice((1, 2, 3, 5, 8, 11, 14, 20, 30)) * 30 + R.randint(-10, 10)
    issue = first - dt.timedelta(span)
    settle = issue + dt.timedelta(R.randint(1, max(1, span - 1)))
    rate = R.choice((0.0785, 0.05, 0.1, 0.0))
    yld = R.choice((0.0625, 0.04, 0.12, 0.0))
    red = R.choice((100, 100, 102))
    cell('=ODDFPRICE(%s,%s,%s,%s,%r,%r,%r,%d,%d)' % (d(settle), d(mat), d(issue), d(first), rate, yld, red, f, b))
    pr = R.choice((113.6, 98.75, 100, 91.5))
    cell('=ODDFYIELD(%s,%s,%s,%s,%r,%r,%r,%d,%d)' % (d(settle), d(mat), d(issue), d(first), rate or 0.05, pr, red, f, b))
for _ in range(90):
    f = R.choice((1, 2, 4))
    b = R.randint(0, 4)
    lastint = dt.date(2020, 1, 1) + dt.timedelta(R.randint(0, 2500))
    if R.random() < 0.3:
        lastint = add_months(lastint, 0, True)
    mat = lastint + dt.timedelta(R.choice((1, 2, 4, 5, 7, 11, 13, 20)) * 30 + R.randint(-10, 10))
    settle = lastint + dt.timedelta(R.randint(0, max(0, (mat - lastint).days - 1)))
    rate = R.choice((0.0375, 0.05, 0.1, 0.0))
    yld = R.choice((0.0405, 0.04, 0.12, 0.0))
    red = R.choice((100, 100, 102))
    cell('=ODDLPRICE(%s,%s,%s,%r,%r,%r,%d,%d)' % (d(settle), d(mat), d(lastint), rate, yld, red, f, b))
    pr = R.choice((99.875, 98.75, 100, 101.5))
    cell('=ODDLYIELD(%s,%s,%s,%r,%r,%r,%d,%d)' % (d(settle), d(mat), d(lastint), rate, pr, red, f, b))

# AMORDEGRC / AMORLINC.
for _ in range(80):
    cost = R.choice((2400, 10000, 1500.5, 100000))
    bought = dt.date(2018, 1, 1) + dt.timedelta(R.randint(0, 2000))
    first = R.choice((dt.date(bought.year, 12, 31), dt.date(bought.year + 1, 12, 31), dt.date(bought.year, 6, 30)))
    if first <= bought:
        first = dt.date(bought.year + 1, 12, 31)
    salvage = R.choice((0, 300, 100, cost * 0.1))
    rate = R.choice((0.15, 0.1, 0.25, 0.3, 0.4, 0.5, 0.2, 0.18, 0.35, 0.05))
    b = R.choice((0, 1, 3, 4, 2))
    for period in R.sample(range(0, 12), 4):
        for fn in ('AMORDEGRC', 'AMORLINC'):
            cell('=%s(%r,%s,%s,%r,%d,%r,%d)' % (fn, cost, d(bought), d(first), salvage, period, rate, b))

# Invalid inputs and argument handling.
S, M = 'DATE(2008,2,15)', 'DATE(2017,11,15)'
for f in ('0', '3', '2.9', '4.5', '1.1', '12', '-2'):
    for fn in ('COUPDAYS', 'COUPNUM', 'COUPNCD'):
        cell('=%s(%s,%s,%s,1)' % (fn, S, M, f))
    cell('=PRICE(%s,%s,0.0575,0.065,100,%s,0)' % (S, M, f))
for b in ('-1', '5', '4.9', '0.5', '"1"', 'TRUE', '""', '"x"'):
    for fn in ('COUPDAYS', 'COUPDAYBS'):
        cell('=%s(%s,%s,2,%s)' % (fn, S, M, b))
    cell('=PRICE(%s,%s,0.0575,0.065,100,2,%s)' % (S, M, b))
    cell('=AMORLINC(2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15,%s)' % b)
    cell('=AMORDEGRC(2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15,%s)' % b)
for fn in ('COUPDAYS', 'COUPPCD', 'COUPNUM'):
    cell('=%s(%s,%s,2,0)' % (fn, M, S))
    cell('=%s(%s,%s,2,0)' % (fn, S, S))
    cell('=%s("2008-02-15","2017-11-15",2,0)' % fn)
    cell('=%s("abc",%s,2,0)' % (fn, M))
    cell('=%s(%s+0.75,%s+0.25,2,0)' % (fn, S, M))
    cell('=%s(-1,%s,2,0)' % (fn, M))
    cell('=%s(0,%s,2,0)' % (fn, M))
    cell('=%s(%s,%s,2)' % (fn, S, M))
    cell('=%s(%s,3000000,2)' % (fn, S))
    cell('=%s(%s,%s,2,)' % (fn, S, M))
    cell('=%s(%s,%s,,0)' % (fn, S, M))
    cell('=%s(%s,%s,#N/A,0)' % (fn, S, M))
    cell('=%s(TRUE,%s,2,0)' % (fn, M))
for args in ('0.0575,-0.065,100', '-0.0575,0.065,100', '0.0575,0.065,0', '0.0575,0.065,-1', '0.0575,"x",100'):
    cell('=PRICE(%s,%s,%s,2,0)' % (S, M, args))
for args in ('0.0575,95,100', '-0.0575,95,100', '0.0575,0,100', '0.0575,95,0', '0.0575,-95,100', '0,95,100',
             '0.0575,1000,100', '0.0575,1,100', '0.0575,95,1', '5,20,100'):
    cell('=YIELD(%s,%s,%s,2,0)' % (S, M, args))
    cell('=YIELD(%s,DATE(2008,6,1),%s,2,0)' % (S, args))
for args in ('0.08,0.09', '-0.08,0.09', '0.08,-0.09', '0,0', '0.08,0'):
    cell('=DURATION(DATE(2018,7,1),DATE(2048,1,1),%s,2,1)' % args)
    cell('=MDURATION(DATE(2008,1,1),DATE(2016,1,1),%s,2,1)' % args)
cell('=DURATION(DATE(2048,1,1),DATE(2018,7,1),0.08,0.09,2,1)')
for args in ('DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,0',
             'DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,0,FALSE',
             'DATE(2008,3,5),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,0,TRUE',
             'DATE(2008,3,1),DATE(2008,8,31),DATE(2008,3,1),0.1,1000,2,0',
             'DATE(2008,3,1),DATE(2008,8,31),DATE(2008,2,1),0.1,1000,2,0',
             'DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0,1000,2,0',
             'DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0.1,0,2,0',
             'DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0.1,,2,0',
             'DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,3,0',
             'DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,5',
             'DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,0,"x"',
             'DATE(2008,3,1),DATE(2008,2,1),DATE(2008,5,1),0.1,1000,2,0',
             'DATE(2008,3,1),DATE(2008,3,1),DATE(2008,5,1),0.1,1000,2,0',
             'DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),-0.1,1000,2,0'):
    cell('=ACCRINT(%s)' % args)
odd = ['DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1',
       'DATE(2008,11,11),DATE(2021,3,1),DATE(2008,11,11),DATE(2009,3,1),0.0785,0.0625,100,2,1',
       'DATE(2009,3,1),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1',
       'DATE(2008,11,11),DATE(2009,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1',
       'DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),-0.0785,0.0625,100,2,1',
       'DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,-0.0625,100,2,1',
       'DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,0,2,1',
       'DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,3,1',
       'DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,5',
       'DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,15),0.0785,0.0625,100,2,1',
       'DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,9,1),0.0785,0.0625,100,2,1',
       'DATE(2008,11,11),DATE(2021,3,1),DATE(2007,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1',
       'DATE(2008,11,11),DATE(2021,3,1),DATE(2005,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1',
       'DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2']
for args in odd:
    cell('=ODDFPRICE(%s)' % args)
    cell('=ODDFYIELD(%s)' % args.replace(',0.0625,', ',98.5,').replace(',-0.0625,', ',-98.5,'))
oddl = ['DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.0405,100,2,0',
        'DATE(2008,2,7),DATE(2008,6,15),DATE(2008,2,7),0.0375,0.0405,100,2,0',
        'DATE(2008,2,7),DATE(2008,6,15),DATE(2008,3,7),0.0375,0.0405,100,2,0',
        'DATE(2008,6,15),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.0405,100,2,0',
        'DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),-0.0375,0.0405,100,2,0',
        'DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,-0.0405,100,2,0',
        'DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.0405,0,2,0',
        'DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.0405,100,5,0',
        'DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.0405,100,2,-1',
        'DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.0405,100,2',
        'DATE(2008,2,7),DATE(2010,6,15),DATE(2007,10,15),0.0375,0.0405,100,2,0',
        'DATE(2008,2,7),DATE(2008,6,15),DATE(2005,10,15),0.0375,0.0405,100,1,0']
for args in oddl:
    cell('=ODDLPRICE(%s)' % args)
    cell('=ODDLYIELD(%s)' % args.replace(',0.0405,', ',99.875,').replace(',-0.0405,', ',-99.875,'))
for args in ('2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15,1', '2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15',
             '-2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15,1', '2400,DATE(2008,8,19),DATE(2008,12,31),-300,1,0.15,1',
             '2400,DATE(2008,8,19),DATE(2008,12,31),300,-1,0.15,1', '2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0,1',
             '2400,DATE(2008,8,19),DATE(2008,12,31),300,1,-0.15,1', '2400,DATE(2008,12,31),DATE(2008,8,19),300,1,0.15,1',
             '2400,DATE(2008,8,19),DATE(2008,8,19),300,1,0.15,1', '2400,DATE(2008,8,19),DATE(2008,12,31),2400,1,0.15,1',
             '2400,DATE(2008,8,19),DATE(2008,12,31),3000,1,0.15,1', '2400,DATE(2008,8,19),DATE(2008,12,31),300,1.7,0.15,1',
             '2400,DATE(2008,8,19),DATE(2008,12,31),300,0,0.15,1', '2400,DATE(2008,8,19),DATE(2008,12,31),300,0,0.15,0',
             '2400,DATE(2008,1,1),DATE(2008,12,31),300,0,0.15,1', '2400,DATE(2007,1,1),DATE(2007,12,31),300,0,0.15,1',
             '2400,DATE(2008,8,19),DATE(2008,12,31),300,1,1.5,1', '2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.5,1',
             '2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.34,1', '2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.2,1',
             '2400,DATE(2008,8,19),DATE(2008,12,31),300,1,1/6,1', '2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.25,1',
             '2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.4,1', '2400,DATE(2008,8,19),DATE(2008,12,31),300,1,1/3,1'):
    cell('=AMORDEGRC(%s)' % args)
    cell('=AMORLINC(%s)' % args)
# Microsoft's examples.
cell('=PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0)')
cell('=YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,95.04287,100,2,0)')
cell('=DURATION(DATE(2018,7,1),DATE(2048,1,1),0.08,0.09,2,1)')
cell('=MDURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,1)')
cell('=ACCRINT(DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,0)')
cell('=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1)')
cell('=ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0575,84.5,100,2,0)')
cell('=ODDLPRICE(DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.0405,100,2,0)')
cell('=ODDLYIELD(DATE(2008,4,20),DATE(2008,6,15),DATE(2007,12,24),0.0375,99.875,100,2,0)')

# --- Structure sweeps: what the rules in FnBonds.fs were found on. ---------------------------------
# Each is deterministic on its own seed, so the cells above stay as they were.

# COUPDAYS of basis 1 over settlements every 9 days for six maturities (its period is not COUPPCD's
# to COUPNCD's).
for mat in [dt.date(2030, 3, 30), dt.date(2030, 2, 28), dt.date(2030, 8, 31), dt.date(2030, 5, 29),
            dt.date(2030, 3, 31), dt.date(2030, 1, 30)]:
    x = dt.date(2027, 1, 3)
    while x < mat:
        for f in (2, 4):
            cell('=COUPDAYS(%s,%s,%d,1)' % (d(x), d(mat), f))
        x += dt.timedelta(9)

# ACCRINT: issues and settlements around the quasi-coupon dates of one first interest date.
fi = dt.date(2023, 3, 23)
issues = [dt.date(*t) for t in [(2022, 4, 23), (2022, 4, 10), (2022, 5, 1), (2022, 6, 23), (2022, 7, 5), (2022, 9, 30),
                                (2022, 12, 24), (2023, 1, 10), (2021, 11, 23)]]
setts = [dt.date(*t) for t in [(2022, 11, 7), (2022, 12, 1), (2023, 1, 15), (2023, 3, 22), (2023, 3, 23), (2023, 3, 24),
                               (2023, 5, 8), (2023, 8, 1), (2024, 1, 15)]]
for b, fs in ((2, (4, 2, 1)), (1, (4,)), (0, (4,))):
    for f in fs:
        for i in issues:
            for x in setts:
                if x <= i:
                    continue
                for m in ('TRUE', 'FALSE'):
                    cell('=ACCRINT(%s,%s,%s,0.1,1000,%d,%d,%s)' % (d(i), d(fi), d(x), f, b, m))

# ODDFPRICE: long and short first periods, settlement in each of their quasi-coupon periods; a zero
# rate shows the discounting, a zero yield the day counts.
R2 = random.Random(7)
for fcd in [dt.date(2029, 4, 30), dt.date(2029, 3, 30), dt.date(2029, 3, 15), dt.date(2029, 8, 31), dt.date(2029, 2, 28),
            dt.date(2029, 6, 10)]:
    eom = fcd.day == calendar.monthrange(fcd.year, fcd.month)[1]
    for f in (1, 2, 4):
        m = 12 // f
        mat = add_months(fcd, 6 * m, eom)
        for nc in (1, 2, 3, 4):
            p_hi = add_months(fcd, -(nc - 1) * m, eom)
            p_lo = add_months(fcd, -nc * m, eom)
            span = (p_hi - p_lo).days
            for frac in (0.1, 0.6):
                issue = p_lo + dt.timedelta(max(1, int(span * frac)))
                for sk in range(nc):
                    s_hi = add_months(fcd, -sk * m, eom)
                    s_lo = add_months(fcd, -(sk + 1) * m, eom)
                    x = s_lo + dt.timedelta(max(1, int((s_hi - s_lo).days * R2.choice((0.3, 0.8)))))
                    if x <= issue:
                        x = issue + dt.timedelta(1)
                    if x >= fcd:
                        continue
                    b = R2.randint(0, 4)
                    for args in ('0,0.08', '0.1,0', '0.1,0.08'):
                        cell('=ODDFPRICE(%s,%s,%s,%s,%s,100,%d,%d)' % (d(x), d(mat), d(issue), d(fcd), args, f, b))

# ODDLPRICE / ODDLYIELD the same way over the periods after the last coupon; AMORDEGRC / AMORLINC
# over lives from 0.8 to 20 years.
R3 = random.Random(11)
for li in [dt.date(2028, 4, 30), dt.date(2028, 3, 30), dt.date(2028, 3, 15), dt.date(2028, 8, 31), dt.date(2028, 2, 29),
           dt.date(2028, 6, 10), dt.date(2028, 11, 30)]:
    eom = li.day == calendar.monthrange(li.year, li.month)[1]
    for f in (1, 2, 4):
        m = 12 // f
        for nc in (1, 2, 3):
            for frac in (0.15, 0.55, 0.95):
                lo = add_months(li, (nc - 1) * m, eom)
                hi = add_months(li, nc * m, eom)
                mat = lo + dt.timedelta(max(1, int((hi - lo).days * frac)))
                for sk in range(nc):
                    s_lo = add_months(li, sk * m, eom)
                    s_hi = add_months(li, (sk + 1) * m, eom)
                    x = s_lo + dt.timedelta(max(1, int((s_hi - s_lo).days * R3.choice((0.2, 0.7)))))
                    if x >= mat:
                        x = mat - dt.timedelta(1)
                    if x <= li:
                        continue
                    b = R3.randint(0, 4)
                    for args in ('0,0.08', '0.1,0', '0.1,0.08'):
                        cell('=ODDLPRICE(%s,%s,%s,%s,100,%d,%d)' % (d(x), d(mat), d(li), args, f, b))
                    cell('=ODDLYIELD(%s,%s,%s,0.1,98,100,%d,%d)' % (d(x), d(mat), d(li), f, b))
for rate in (0.05, 0.1, 1 / 7.5, 0.15, 1 / 6, 0.18, 0.2, 0.21, 0.22, 0.24, 0.25, 0.26, 0.3, 1 / 3, 0.34, 0.35, 0.4, 0.45,
             0.49, 0.5, 0.55, 0.7, 0.99, 1.0, 1.2):
    for bought, first in ((dt.date(2008, 8, 19), dt.date(2008, 12, 31)), (dt.date(2008, 1, 1), dt.date(2008, 12, 31)),
                          (dt.date(2009, 3, 1), dt.date(2009, 12, 31)), (dt.date(2010, 12, 31), dt.date(2010, 12, 31)),
                          (dt.date(2010, 6, 30), dt.date(2011, 6, 30))):
        b = R3.choice((0, 1, 3, 4))
        for period in range(0, 9):
            for fn in ('AMORDEGRC', 'AMORLINC'):
                cell('=%s(10000,%s,%s,1000,%d,%r,%d)' % (fn, d(bought), d(first), period, rate, b))


for i, f in enumerate(cells):
    print('S\tA%d\t%s' % (i + 1, f))
