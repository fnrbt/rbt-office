#!/usr/bin/env python3
"""Probe cells for the normal distribution (Formula.fs's NORM.* and its legacy names, GAUSS, PHI, Z.TEST,
CONFIDENCE.NORM; FnStatDist.fs's LOGNORM.*), as eval-cells-on-vm.sh takes them.

    python3 tools/excel-parity/make-normal-probe.py > normal-cells.txt

One line per cell: sheet, address, formula. The distribution function over z = -40 .. 40 (finely through
the far lower tail, where Excel's results leave the normal doubles, and around the upper tail where 1 - p
rounds to 1), the density there, the quantile of p = 10^-k down to the least normal doubles (Excel does not
take a subnormal typed in) and of p near 1, the
general forms with a mean and deviation, the lognormal ones, and the domain errors. Deterministic (fixed
seed): normal-excel.tsv holds Excel 2019's values; function-parity-engine.fsx compares the engine.
"""
import random

R = random.Random(20261125)
cells = []


def cell(f):
    cells.append(f)


zs = [x / 4 for x in range(-160, 161)]
zs += [-37 - x / 40 for x in range(0, 61)]
zs += [-8.5 + x / 10 for x in range(0, 15)]
zs += [R.uniform(-39, 9) for _ in range(120)]
zs += [1e-10, -1e-10, 1e-300, 0.1, -0.1, 0.01]
for z in zs:
    cell('=NORM.S.DIST(%r,TRUE)' % z)
for z in zs[::3]:
    cell('=NORMSDIST(%r)' % z)
for z in [x / 2 for x in range(-80, 81)] + [37.5 + x / 40 for x in range(0, 61)] + [R.uniform(-40, 40) for _ in range(30)]:
    cell('=NORM.S.DIST(%r,FALSE)' % z)
ps = [10.0 ** -k for k in range(1, 309)] + [3e-308, 2.3e-308]
ps += [R.uniform(0, 1) for _ in range(80)] + [R.uniform(0, 1) ** 8 for _ in range(30)]
ps += [1 - 10.0 ** -k for k in range(1, 17)] + [0.5, 0.5 + 1e-15, 0.5 - 1e-12, 0.975, 0.025, 0.02425, 0.97575]
for p in ps:
    cell('=NORM.S.INV(%r)' % p)
for p in ps[::4]:
    cell('=NORMSINV(%r)' % p)
for _ in range(60):
    x, m, s = R.uniform(-50, 50), R.uniform(-20, 20), R.choice([0.5, 1, 2.5, 10, R.uniform(0.01, 30)])
    cell('=NORM.DIST(%r,%r,%r,TRUE)' % (x, m, s))
    cell('=NORM.DIST(%r,%r,%r,FALSE)' % (x, m, s))
    cell('=NORMDIST(%r,%r,%r,%s)' % (x, m, s, R.choice(['TRUE', 'FALSE'])))
    p = R.choice([R.uniform(0, 1), 10.0 ** -R.randint(1, 300), 1 - 10.0 ** -R.randint(1, 15)])
    cell('=NORM.INV(%r,%r,%r)' % (p, m, s))
    cell('=NORMINV(%r,%r,%r)' % (p, m, s))
for z in [x / 2 for x in range(-80, 81)] + [R.uniform(-40, 40) for _ in range(20)]:
    cell('=GAUSS(%r)' % z)
    cell('=PHI(%r)' % z)
for _ in range(40):
    n = R.randint(1, 8)
    xs = ','.join('%r' % round(R.uniform(-5, 5), 3) for _ in range(n))
    x = R.choice([R.uniform(-3, 3), R.uniform(-60, 60)])
    cell('=Z.TEST({%s},%r)' % (xs, x))
    cell('=Z.TEST({%s},%r,%r)' % (xs, x, R.uniform(0.05, 3)))
    cell('=ZTEST({%s},%r)' % (xs, x))
for _ in range(40):
    a = R.choice([R.uniform(0, 1), 10.0 ** -R.randint(1, 300), 1 - 10.0 ** -R.randint(1, 15)])
    cell('=CONFIDENCE.NORM(%r,%r,%r)' % (a, R.uniform(0.1, 10), R.choice([1, 2, 7.9, 100])))
    cell('=CONFIDENCE(%r,%r,%r)' % (a, R.uniform(0.1, 10), R.choice([1, 2, 7.9, 100])))
for _ in range(60):
    x, m, s = R.choice([R.uniform(0, 5), 10.0 ** -R.randint(1, 300), R.uniform(1, 1e6)]), R.uniform(-3, 3), R.uniform(0.1, 3)
    cell('=LOGNORM.DIST(%r,%r,%r,TRUE)' % (x, m, s))
    cell('=LOGNORM.DIST(%r,%r,%r,FALSE)' % (x, m, s))
    cell('=LOGNORMDIST(%r,%r,%r)' % (x, m, s))
    p = R.choice([R.uniform(0, 1), 10.0 ** -R.randint(1, 300), 1 - 10.0 ** -R.randint(1, 15)])
    cell('=LOGNORM.INV(%r,%r,%r)' % (p, m, s))
    cell('=LOGINV(%r,%r,%r)' % (p, m, s))
for f in ['NORM.S.INV(0)', 'NORM.S.INV(1)', 'NORM.S.INV(-0.1)', 'NORM.S.INV("x")', 'NORM.S.INV("0.3")',
          'NORM.S.DIST("x",TRUE)', 'NORM.S.DIST(1,"x")', 'NORM.S.DIST(1,2)', 'NORM.S.DIST(1,0)',
          'NORM.DIST(1,0,0,TRUE)', 'NORM.DIST(1,0,-1,TRUE)', 'NORM.INV(0.5,0,0)',
          'NORM.INV(0,0,1)', 'NORM.INV(1,0,1)', 'NORMSDIST(TRUE)', 'NORM.S.DIST(1E+300,TRUE)', 'NORM.S.DIST(-1E+300,TRUE)',
          'NORM.S.DIST(1E+300,FALSE)', 'GAUSS(1E+300)', 'GAUSS(-1E+300)', 'PHI(1E+300)', 'Z.TEST({1},1)', 'Z.TEST({1,1},1)',
          'Z.TEST({1,2},1,0)', 'CONFIDENCE.NORM(0,1,1)', 'CONFIDENCE.NORM(1,1,1)', 'CONFIDENCE.NORM(0.05,1,0.5)',
          'CONFIDENCE.NORM(0.05,0,1)', 'LOGNORM.DIST(0,0,1,TRUE)', 'LOGNORM.DIST(1,0,0,TRUE)', 'LOGNORM.INV(0,0,1)',
          'LOGNORM.INV(0.5,0,0)', 'LOGNORM.INV(1E-300,0,3)', 'LOGNORM.INV(1-1E-16,700,1)', 'NORM.INV(1E-300,0,1E+300)']:
    cell('=' + f)

for i, f in enumerate(cells):
    print('S\tA%d\t%s' % (i + 1, f))
