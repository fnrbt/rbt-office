#!/usr/bin/env python3
"""Generate the probe cells of the statistical-distribution functions (FnStatDist.fs).

    python3 tools/excel-parity/make-statdist-probe.py OUTDIR
    for f in OUTDIR/statdist-cells-*.tsv; do
        tools/excel-parity/eval-cells-on-vm.sh "$f" "${f/cells/excel}"; done
    cat OUTDIR/statdist-excel-*.tsv > tools/excel-parity/statdist-excel-results.tsv

Batch k is one sheet "Sk": the data block H1:N20 (the same in every batch,
typed first) and the probe formulas in column A, one per row. The engine side
(statdist-compare.fsx) types the same cells into a model and evaluates them.
Per function the probe covers typical values, the tails, truncation of the
integer arguments, the boundaries of the domain and the coercion of each
argument (text, a text number, a logical, a blank cell, an omitted argument,
an error).
"""
import itertools
import os
import sys

SHEET = 'S'
BATCH = 900

# The data block: H y, I x1, J x2, K x3 = 2*x1 (collinear), L mixed (text,
# blank, logical), M blank, N y > 0 for LOGEST/GROWTH.
Y = [3.1, 4.9, 7.2, 8.8, 11.3, 12.7, 15.2, 16.9, 19.4, 20.8]
X1 = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
X2 = [2.5, 1.0, 4.2, 3.3, 6.1, 5.0, 7.7, 6.4, 9.9, 8.2]
L = ['4.5', 'text', None, 'TRUE', '6', '7.5', None, '2', '9', '1']
NY = [1.2, 1.9, 3.1, 4.8, 7.9, 12.1, 19.8, 30.5, 50.2, 79.1]
DATA = []
for r in range(10):
    DATA.append((f'H{r+1}', repr(Y[r])))
    DATA.append((f'I{r+1}', repr(X1[r])))
    DATA.append((f'J{r+1}', repr(X2[r])))
    DATA.append((f'K{r+1}', repr(2 * X1[r])))
    if L[r] is not None:
        DATA.append((f'L{r+1}', L[r]))
    DATA.append((f'N{r+1}', repr(NY[r])))
# A row of y (H12:Q12) and x (H13:Q13) for row-oriented regression.
for c in range(10):
    col = chr(ord('H') + c)
    DATA.append((f'{col}12', repr(Y[c])))
    DATA.append((f'{col}13', repr(X1[c])))
# Contingency tables for CHISQ.TEST: H15:J16 observed, H18:J19 expected.
OBS = [[58, 35, 7], [11, 25, 23]]
EXP = [[45.35, 47.65, 10.0], [17.56, 18.44, 23.0]]
for r in range(2):
    for c in range(3):
        DATA.append((f'{chr(ord("H")+c)}{15+r}', repr(OBS[r][c])))
        DATA.append((f'{chr(ord("H")+c)}{18+r}', repr(EXP[r][c])))

BLANK = 'M1'

formulas = []


def add(f):
    formulas.append('=' + f)


def call(name, args):
    return f'{name}({",".join(args)})'


def grid(name, *axes):
    for combo in itertools.product(*axes):
        add(call(name, [str(a) for a in combo]))


def sweep(name, base, numeric_positions=None, bool_positions=()):
    """Each argument of `base` in turn replaced by the coercion cases."""
    n = len(base)
    for i in range(n):
        variants = ['"abc"', f'"{base[i]}"' if base[i] not in ('TRUE', 'FALSE') else '"TRUE"', BLANK, '#N/A', '']
        if i not in bool_positions:
            variants += ['TRUE', 'FALSE']
        else:
            variants += ['0', '2', '-1', '0.5']
        for v in variants:
            args = list(base)
            args[i] = v
            add(call(name, args))
    # (Too few or too many arguments are not probed: Excel refuses such a formula.)


XS_T = [-40, -6, -2.5, -1, -0.3, 0, 1e-8, 0.5, 1, 1.96, 3, 10, 50, 1e5]
DFS = [1, 1.9, 2, 2.5, 3, 5, 10, 30, 100, 1000, 1e6, 1e10]
PS = [1e-300, 1e-12, 1e-6, 0.001, 0.025, 0.05, 0.1, 0.3, 0.5, 0.7, 0.9, 0.95, 0.975, 0.999, 1 - 1e-9, 0.9999999999999]

# ----- Student's t
grid('T.DIST', XS_T, [1, 2.5, 3, 10, 100, 1e6], ['TRUE', 'FALSE'])
grid('T.DIST', [0.7], [0.5, 0.99, -1, 0, 1e10, 1e11], ['TRUE'])
sweep('T.DIST', ['1.5', '4', 'TRUE'], bool_positions=(2,))
grid('T.DIST.2T', [0, 1e-8, 0.5, 1, 1.96, 3, 10, 50, 1e5, -1], [1, 2.9, 3, 10, 1000, 1e6])
grid('T.DIST.2T', [1], [0.5, 0.999, 1e10, 1e11])
sweep('T.DIST.2T', ['1.5', '4'])
grid('T.DIST.RT', [-10, -1, 0, 0.5, 1.96, 3, 10, 50], [1, 2.5, 3, 30, 1e6])
sweep('T.DIST.RT', ['1.5', '4'])
grid('T.INV', PS + [0, 1, -0.1, 1.1], [1, 2, 2.7, 5, 30, 1000, 1e6])
grid('T.INV', [0.9], [0.5, 0.99, 1e10, 1e11])
sweep('T.INV', ['0.9', '4'])
grid('T.INV.2T', PS + [0, 1, -0.1], [1, 2, 2.7, 5, 30, 1000, 1e6])
sweep('T.INV.2T', ['0.1', '4'])
grid('TDIST', [0, 0.5, 1, 1.96, 3, 10, -1], [1, 2.5, 10, 1e6], [1, 2])
grid('TDIST', [1.5], [4], [0, 1.5, 2.9, 3, -1])
sweep('TDIST', ['1.5', '4', '2'])
grid('TINV', [1e-12, 0.001, 0.05, 0.1, 0.5, 0.9, 1, 0, 1.1], [1, 2.5, 10, 1e6])
sweep('TINV', ['0.1', '4'])

# ----- F
XS_F = [0, 1e-10, 0.1, 0.5, 1, 1.5, 3, 10, 100, 1e5, -1]
DF_F = [(1, 1), (1, 10), (2, 2), (2.5, 7.9), (5, 3), (10, 20), (30, 1000), (1e6, 1e6), (3, 1e10)]
for x in XS_F:
    for d1, d2 in DF_F:
        for c in ['TRUE', 'FALSE']:
            add(call('F.DIST', [str(x), str(d1), str(d2), c]))
grid('F.DIST', [1], [0.5, 1e10, 1e11], [5], ['TRUE'])
grid('F.DIST', [1], [5], [0.5, 1e10, 1e11], ['TRUE'])
sweep('F.DIST', ['1.5', '4', '6', 'TRUE'], bool_positions=(3,))
for x in XS_F:
    for d1, d2 in DF_F:
        add(call('F.DIST.RT', [str(x), str(d1), str(d2)]))
        add(call('FDIST', [str(x), str(d1), str(d2)]))
sweep('F.DIST.RT', ['1.5', '4', '6'])
sweep('FDIST', ['1.5', '4', '6'])
for p in PS + [0, 1, 1.1]:
    for d1, d2 in [(1, 1), (2, 5), (4.5, 6), (10, 20), (100, 3), (1e6, 1e6)]:
        add(call('F.INV', [str(p), str(d1), str(d2)]))
        add(call('F.INV.RT', [str(p), str(d1), str(d2)]))
        add(call('FINV', [str(p), str(d1), str(d2)]))
sweep('F.INV', ['0.3', '4', '6'])
sweep('F.INV.RT', ['0.3', '4', '6'])
sweep('FINV', ['0.3', '4', '6'])

# ----- chi-squared
XS_C = [0, 1e-10, 0.01, 0.5, 1, 2, 3.84, 10, 50, 100, 1000, 1e5, -1]
DF_C = [1, 1.9, 2, 3, 5.5, 10, 100, 1e4, 1e10, 1e11]
grid('CHISQ.DIST', XS_C, DF_C[:-2], ['TRUE', 'FALSE'])
grid('CHISQ.DIST', [1], [0.5, 1e10, 1e11], ['TRUE', 'FALSE'])
sweep('CHISQ.DIST', ['2.5', '4', 'TRUE'], bool_positions=(2,))
grid('CHISQ.DIST.RT', XS_C, DF_C)
grid('CHIDIST', XS_C, DF_C)
sweep('CHISQ.DIST.RT', ['2.5', '4'])
sweep('CHIDIST', ['2.5', '4'])
grid('CHISQ.INV', PS + [0, 1], [1, 2, 3.5, 10, 100, 1e4])
grid('CHISQ.INV.RT', PS + [0, 1], [1, 2, 3.5, 10, 100, 1e4])
grid('CHIINV', PS + [0, 1], [1, 2, 3.5, 10, 100, 1e4])
grid('CHISQ.INV', [0.5], [0.5, 1e10, 1e11])
sweep('CHISQ.INV', ['0.3', '4'])
sweep('CHISQ.INV.RT', ['0.3', '4'])
sweep('CHIINV', ['0.3', '4'])

# ----- beta
XS_B = [0, 1e-10, 0.01, 0.2, 0.5, 0.8, 0.99, 1, -0.1, 1.1]
AB = [(1, 1), (0.5, 0.5), (2, 3), (0.1, 5), (5, 0.1), (10, 10), (100, 200), (1e4, 3e4), (3, 1e6)]
for x in XS_B:
    for a, b in AB:
        for c in ['TRUE', 'FALSE']:
            add(call('BETA.DIST', [str(x), str(a), str(b), c]))
        add(call('BETADIST', [str(x), str(a), str(b)]))
for args in [('3', '2', '5', 'TRUE', '1', '4'), ('3', '2', '5', 'FALSE', '1', '4'), ('1', '2', '5', 'TRUE', '1', '4'),
             ('4', '2', '5', 'TRUE', '1', '4'), ('4', '2', '5', 'FALSE', '1', '4'), ('0.5', '2', '5', 'TRUE', '1', '4'),
             ('3', '2', '5', 'TRUE', '4', '1'), ('3', '2', '5', 'TRUE', '3', '3'), ('3', '2', '5', 'TRUE', '1'),
             ('-3', '2', '5', 'TRUE', '-4', '-1'), ('2', '0', '5', 'TRUE'), ('0.5', '2', '0', 'TRUE'), ('0.5', '-1', '2', 'TRUE')]:
    add(call('BETA.DIST', list(args)))
for args in [('3', '2', '5', '1', '4'), ('1', '2', '5', '1', '4'), ('4', '2', '5', '1', '4'), ('3', '2', '5', '1'),
             ('3', '2', '5', '4', '1'), ('0.5', '0', '5'), ('0.5', '2', '-1'), ('3', '2', '5', '3', '3')]:
    add(call('BETADIST', list(args)))
sweep('BETA.DIST', ['0.4', '2', '3', 'TRUE', '0', '1'], bool_positions=(3,))
sweep('BETADIST', ['0.4', '2', '3', '0', '1'])
for p in PS + [0, 1]:
    for a, b in [(1, 1), (0.5, 0.5), (2, 3), (0.1, 5), (10, 10), (100, 200)]:
        add(call('BETA.INV', [str(p), str(a), str(b)]))
        add(call('BETAINV', [str(p), str(a), str(b)]))
for args in [('0.3', '2', '5', '1', '4'), ('0.3', '2', '5', '1'), ('0.3', '2', '5', '4', '1'), ('0.3', '2', '5', '3', '3'),
             ('0.3', '0', '5'), ('0.3', '2', '0')]:
    add(call('BETA.INV', list(args)))
    add(call('BETAINV', list(args)))
sweep('BETA.INV', ['0.3', '2', '3', '0', '1'])
sweep('BETAINV', ['0.3', '2', '3', '0', '1'])

# ----- gamma
for x in [1, 2, 3, 0.5, 1.5, 2.5, 0.1, 1e-5, 1e-300, 4.7, 10, 10.5, 20, 33.3, 100, 150.5, 170, 171, 171.6, 172, 0, -1, -2,
          -0.5, -1.5, -2.5, -3.7, -10.2, -100.5, -170.5, -171.5, -180.5, -1e-5]:
    add(call('GAMMA', [str(x)]))
for x in [1, 2, 3, 0.5, 1.5, 2.5, 0.1, 1e-5, 1e-300, 0.999, 1.001, 1.999, 2.001, 4.7, 10, 100, 1e5, 1e10, 1e300, 0, -1, -0.5]:
    add(call('GAMMALN', [str(x)]))
    add(call('GAMMALN.PRECISE', [str(x)]))
for f in ['GAMMA', 'GAMMALN', 'GAMMALN.PRECISE']:
    for v in ['"abc"', '"2.5"', 'TRUE', 'FALSE', BLANK, '#N/A']:
        add(call(f, [v]))
XS_G = [0, 1e-10, 0.1, 0.5, 1, 2, 5, 10, 30, 100, 1000, -1]
AB_G = [(0.1, 1), (0.5, 2), (1, 1), (2, 0.5), (3, 2), (9, 1), (50, 2), (1000, 1), (1e5, 0.01)]
for x in XS_G:
    for a, b in AB_G:
        for c in ['TRUE', 'FALSE']:
            add(call('GAMMA.DIST', [str(x), str(a), str(b), c]))
        add(call('GAMMADIST', [str(x), str(a), str(b), 'TRUE']))
grid('GAMMA.DIST', [1], [0, -1], [2], ['TRUE'])
grid('GAMMA.DIST', [1], [2], [0, -1], ['TRUE'])
sweep('GAMMA.DIST', ['2.5', '3', '2', 'TRUE'], bool_positions=(3,))
sweep('GAMMADIST', ['2.5', '3', '2', 'FALSE'], bool_positions=(3,))
for p in PS + [0, 1]:
    for a, b in [(0.1, 1), (0.5, 2), (1, 1), (3, 2), (50, 2), (1000, 1)]:
        add(call('GAMMA.INV', [str(p), str(a), str(b)]))
        add(call('GAMMAINV', [str(p), str(a), str(b)]))
grid('GAMMA.INV', [0.5], [0, -1, 2], [0, -1])
sweep('GAMMA.INV', ['0.3', '3', '2'])
sweep('GAMMAINV', ['0.3', '3', '2'])

# ----- binomial
for s, n, p in [(0, 10, 0.5), (3, 10, 0.5), (5, 10, 0.5), (10, 10, 0.5), (6, 10, 0.3), (6.9, 10.9, 0.3), (0, 0, 0.5),
                (0, 1, 0), (1, 1, 0), (0, 1, 1), (1, 1, 1), (11, 10, 0.5), (-1, 10, 0.5), (3, 10, -0.1), (3, 10, 1.1),
                (50, 100, 0.5), (500, 1000, 0.5), (30, 1000, 0.01), (1, 1000, 0.999), (999, 1000, 0.999), (500, 1000, 0.1),
                (5000, 10000, 0.5), (12, 100000, 1e-4), (5e5, 1e6, 0.5), (3, 1e9, 1e-9)]:
    for c in ['TRUE', 'FALSE']:
        add(call('BINOM.DIST', [str(s), str(n), str(p), c]))
        add(call('BINOMDIST', [str(s), str(n), str(p), c]))
sweep('BINOM.DIST', ['3', '10', '0.4', 'TRUE'], bool_positions=(3,))
sweep('BINOMDIST', ['3', '10', '0.4', 'FALSE'], bool_positions=(3,))
for args in [('60', '0.75', '48'), ('60', '0.75', '45', '50'), ('60', '0.75', '50', '45'), ('60', '0.75', '0', '60'),
             ('60', '0.75', '61'), ('60', '0.75', '-1'), ('60.7', '0.75', '45.9', '50.9'), ('10', '0', '0'), ('10', '1', '10'),
             ('10', '0', '1'), ('10', '1.1', '1'), ('-1', '0.5', '0'), ('1000', '0.3', '250', '350'), ('0', '0.5', '0')]:
    add(call('BINOM.DIST.RANGE', list(args)))
sweep('BINOM.DIST.RANGE', ['10', '0.4', '3', '5'])
for n, p, alpha in [(6, 0.5, 0.75), (10, 0.3, 0.5), (10, 0.3, 0), (10, 0.3, 1), (10, 0.3, 1e-12), (10, 0.3, 0.999999),
                    (100, 0.5, 0.5), (100, 0.5, 0.95), (1000, 0.01, 0.9), (10.9, 0.3, 0.5), (0, 0.5, 0.5), (10, 0, 0.5),
                    (10, 1, 0.5), (10, 1.1, 0.5), (10, 0.3, 1.1), (10, 0.3, -0.1), (-1, 0.5, 0.5), (1e6, 0.5, 0.5),
                    (10, 0.5, 0.623046875), (10, 0.5, 0.376953125)]:
    add(call('BINOM.INV', [str(n), str(p), str(alpha)]))
    add(call('CRITBINOM', [str(n), str(p), str(alpha)]))
sweep('BINOM.INV', ['10', '0.4', '0.5'])
sweep('CRITBINOM', ['10', '0.4', '0.5'])

# ----- Poisson
for x, m in [(0, 0), (1, 0), (0, 1), (2, 5), (5, 5), (10, 5), (2.9, 5), (0, 1e-10), (100, 100), (1000, 900), (3, 1000),
             (1e5, 1e5), (-1, 5), (2, -1), (50, 0.5), (1e6, 999000)]:
    for c in ['TRUE', 'FALSE']:
        add(call('POISSON.DIST', [str(x), str(m), c]))
        add(call('POISSON', [str(x), str(m), c]))
sweep('POISSON.DIST', ['2', '5', 'TRUE'], bool_positions=(2,))
sweep('POISSON', ['2', '5', 'FALSE'], bool_positions=(2,))

# ----- Weibull, exponential, lognormal
for x, a, b in [(105, 20, 100), (0, 2, 1), (0, 1, 1), (0, 0.5, 1), (1, 0.5, 1), (1e-10, 3, 2), (50, 1, 10), (3, 5, 1),
                (-1, 2, 1), (1, 0, 1), (1, 2, 0), (1, -1, 1), (1, 1e3, 1), (0.999, 1e3, 1)]:
    for c in ['TRUE', 'FALSE']:
        add(call('WEIBULL.DIST', [str(x), str(a), str(b), c]))
        add(call('WEIBULL', [str(x), str(a), str(b), c]))
sweep('WEIBULL.DIST', ['2', '3', '4', 'TRUE'], bool_positions=(3,))
sweep('WEIBULL', ['2', '3', '4', 'FALSE'], bool_positions=(3,))
for x, l in [(0.2, 10), (0, 1), (0, 0.5), (1e-10, 1), (1, 1), (50, 1), (800, 1), (-1, 1), (1, 0), (1, -1), (3, 1e-10)]:
    for c in ['TRUE', 'FALSE']:
        add(call('EXPON.DIST', [str(x), str(l), c]))
        add(call('EXPONDIST', [str(x), str(l), c]))
sweep('EXPON.DIST', ['2', '0.5', 'TRUE'], bool_positions=(2,))
sweep('EXPONDIST', ['2', '0.5', 'FALSE'], bool_positions=(2,))
for x, m, s in [(4, 3.5, 1.2), (1, 0, 1), (0.01, 0, 1), (1e5, 0, 1), (1e-10, 0, 1), (0, 0, 1), (-1, 0, 1), (2, 0, 0),
                (2, 0, -1), (2, -3, 0.1), (1e100, 200, 5)]:
    for c in ['TRUE', 'FALSE']:
        add(call('LOGNORM.DIST', [str(x), str(m), str(s), c]))
    add(call('LOGNORMDIST', [str(x), str(m), str(s)]))
sweep('LOGNORM.DIST', ['2', '0.5', '1.5', 'TRUE'], bool_positions=(3,))
sweep('LOGNORMDIST', ['2', '0.5', '1.5'])
for p in PS + [0, 1]:
    for m, s in [(3.5, 1.2), (0, 1), (-2, 0.1)]:
        add(call('LOGNORM.INV', [str(p), str(m), str(s)]))
        add(call('LOGINV', [str(p), str(m), str(s)]))
grid('LOGNORM.INV', [0.5], [0], [0, -1])
sweep('LOGNORM.INV', ['0.3', '0.5', '1.5'])
sweep('LOGINV', ['0.3', '0.5', '1.5'])

# ----- hypergeometric, negative binomial
for args in [(1, 4, 8, 20), (0, 4, 8, 20), (4, 4, 8, 20), (5, 4, 8, 20), (1.9, 4.9, 8.9, 20.9), (3, 10, 5, 20), (0, 15, 5, 20),
             (5, 15, 5, 20), (4, 15, 5, 20), (0, 0, 5, 20), (1, 4, 8, 4), (1, 21, 8, 20), (1, 4, 21, 20), (-1, 4, 8, 20),
             (1, 0, 8, 20), (1, 4, 0, 20), (1, 4, 8, 0), (20, 50, 400, 1000), (300, 1000, 30000, 100000), (5, 4, 3, 20),
             (2, 4, 8, 8)]:
    for c in ['TRUE', 'FALSE']:
        add(call('HYPGEOM.DIST', [str(a) for a in args] + [c]))
    add(call('HYPGEOMDIST', [str(a) for a in args]))
sweep('HYPGEOM.DIST', ['1', '4', '8', '20', 'TRUE'], bool_positions=(4,))
sweep('HYPGEOMDIST', ['1', '4', '8', '20'])
for f, s, p in [(10, 5, 0.25), (0, 5, 0.25), (0, 1, 0.5), (3, 1, 0.5), (10.9, 5.9, 0.25), (10, 5, 0), (10, 5, 1), (0, 5, 1),
                (10, 0, 0.25), (10, 0.5, 0.25), (-1, 5, 0.25), (100, 50, 0.5), (1000, 3, 0.001), (5, 1e4, 0.999),
                (10, 5, 1.1), (10, 5, -0.1)]:
    for c in ['TRUE', 'FALSE']:
        add(call('NEGBINOM.DIST', [str(f), str(s), str(p), c]))
    add(call('NEGBINOMDIST', [str(f), str(s), str(p)]))
sweep('NEGBINOM.DIST', ['10', '5', '0.25', 'TRUE'], bool_positions=(3,))
sweep('NEGBINOMDIST', ['10', '5', '0.25'])

# ----- CONFIDENCE.T
for alpha, sd, n in [(0.05, 1, 50), (0.05, 2.5, 50), (0.01, 1, 2), (0.5, 3, 10.9), (0.05, 1, 1), (0.05, 1, 0.9),
                     (0, 1, 10), (1, 1, 10), (0.05, 0, 10), (0.05, -1, 10), (0.05, 1, 1e6), (1e-10, 1, 3)]:
    add(call('CONFIDENCE.T', [str(alpha), str(sd), str(n)]))
sweep('CONFIDENCE.T', ['0.05', '2', '20'])

# ----- tests
A1 = '{3,4,5,8,9,1,2,4,5}'
A2 = '{6,19,3,2,14,4,5,17,1}'
A3 = '{1.5,2.5,3.5,2.25}'
for a, b in [(A1, A2), (A2, A1), (A1, A3), ('H1:H10', 'I1:I10'), ('H1:H10', 'J1:J10'), ('L1:L10', 'J1:J10'), ('{1,2}', '{1,3}'),
             ('{1}', '{2,3}'), ('{1,1,1}', '{2,2,2}'), ('{1,2,3}', '{1,2,3}'), ('{1,2,3}', '{2,3,4.5}')]:
    for tails in [1, 2]:
        for typ in [1, 2, 3]:
            add(call('T.TEST', [a, b, str(tails), str(typ)]))
            add(call('TTEST', [a, b, str(tails), str(typ)]))
for args in [(A1, A2, '1.9', '1.9'), (A1, A2, '0', '1'), (A1, A2, '3', '1'), (A1, A2, '1', '0'), (A1, A2, '1', '4'),
             (A1, A2, '1', '3.9'), (A1, A2, '"abc"', '1'), (A1, A2, '1', '"2"'), (A1, A2, 'TRUE', 'TRUE'),
             (A1, '{1,#N/A,3}', '1', '1'), (A1, A2, BLANK, '1')]:
    add(call('T.TEST', list(args)))
for a, b in [(A1, A2), (A2, A1), (A1, A3), ('H1:H10', 'I1:I10'), ('L1:L10', 'J1:J10'), ('{1,2}', '{1,3}'), ('{1}', '{2,3}'),
             ('{1,1,1}', '{2,3,4}'), ('{1,2,3}', '{1,2,3}'), ('{1,1}', '{1,1}'), ('{1,2,"a"}', '{2,4,7}'),
             ('{1,2,#N/A}', '{2,4,7}'), ('H1:H10', BLANK)]:
    add(call('F.TEST', [a, b]))
    add(call('FTEST', [a, b]))
for a, b in [('H15:J16', 'H18:J19'), ('H15:J15', 'H18:J18'), ('H15:H16', 'H18:H19'), ('{58,35;11,25}', '{45,40;20,24}'),
             ('H15:J16', 'H18:I19'), ('{1}', '{1}'), ('{1,2}', '{0,2}'), ('{1,2}', '{-1,2}'), ('{1,"a",3}', '{2,2,2}'),
             ('{5,5,5}', '{5,5,5}'), ('{1;2;3}', '{1.5;2;2.5}'), ('{1,2;3,4}', '{1.5,1.5;3.5,3.5}'), ('{1,2,3}', '{1.5;2;2.5}')]:
    add(call('CHISQ.TEST', [a, b]))
    add(call('CHITEST', [a, b]))

# ----- regression: every element of the result through INDEX, and its shape
def regression(name, args, rows, cols):
    f = call(name, args)
    add(f'ROWS({f})')
    add(f'COLUMNS({f})')
    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            add(f'INDEX({f},{r},{c})')


for name, yr in [('LINEST', 'H1:H10'), ('LOGEST', 'N1:N10')]:
    regression(name, [yr], 1, 2)
    regression(name, [yr, 'I1:I10'], 1, 2)
    regression(name, [yr, 'I1:I10', 'TRUE', 'TRUE'], 5, 2)
    regression(name, [yr, 'I1:I10', 'FALSE', 'TRUE'], 5, 2)
    regression(name, [yr, 'I1:J10', 'TRUE', 'TRUE'], 5, 3)
    regression(name, [yr, 'I1:J10', 'FALSE', 'TRUE'], 5, 3)
    regression(name, [yr, 'I1:K10', 'TRUE', 'TRUE'], 5, 4)
    regression(name, [yr, 'I1:K10', 'FALSE', 'TRUE'], 5, 4)
    regression(name, [yr, 'I1:K10', 'TRUE', 'FALSE'], 1, 4)
    regression(name, [yr, 'I1:I10', '0', '1'], 5, 2)
    regression(name, [yr, '', 'TRUE', 'TRUE'], 5, 2)
    regression(name, [yr, 'I1:I10', '', 'TRUE'], 5, 2)
    regression(name, [yr, 'I1:I3'], 1, 2)
    regression(name, [yr, 'L1:L10'], 1, 2)
    regression(name, [yr, 'I1:I10', '"abc"'], 1, 2)
    regression(name, [yr, 'I1:I10', '"TRUE"', 'TRUE'], 5, 2)
    regression(name, ['{1,2,3}', '{2,4,6}', 'TRUE', 'TRUE'], 5, 2)
    regression(name, ['{1,2}', '{1,2}', 'TRUE', 'TRUE'], 5, 2)
    regression(name, ['{1}', '{1}', 'TRUE', 'TRUE'], 5, 2)
    regression(name, ['{1,2,3,4}', '{1,1,1,1}', 'TRUE', 'TRUE'], 5, 2)
    regression(name, ['{1,2,3,4}', '{1,1,1,1}', 'FALSE', 'TRUE'], 5, 2)
    regression(name, ['{2,4,6,8}', '{1,2,3,4}', 'TRUE', 'TRUE'], 5, 2)
    regression(name, ['{1,2,3,4,5}', '{1,4;2,3;3,7;4,1;5,9}', 'TRUE', 'TRUE'], 5, 3)
    regression(name, ['{1;2;3;4;5}', '{1,4;2,3;3,7;4,1;5,9}', 'TRUE', 'TRUE'], 5, 3)
    regression(name, ['H12:Q12', 'H13:Q13', 'TRUE', 'TRUE'], 5, 2)
    regression(name, ['H12:Q12', '', 'TRUE', 'TRUE'], 5, 2)
    regression(name, ['{1,2,"a"}', '{1,2,3}'], 1, 2)
    regression(name, ['{1,2,3}', '{1,#N/A,3}'], 1, 2)
add('LOGEST({1,-2,3},{1,2,3})')
add('LOGEST({1,0,3},{1,2,3})')
add('INDEX(LINEST({1,2,3,4},{1,2,3,4}^{1;2}),1,1)')
for name, yr in [('TREND', 'H1:H10'), ('GROWTH', 'N1:N10')]:
    regression(name, [yr], 10, 1)
    regression(name, [yr, 'I1:I10'], 10, 1)
    regression(name, [yr, 'I1:I10', '{11;12;0.5}'], 3, 1)
    regression(name, [yr, 'I1:I10', '{11,12,0.5}'], 1, 3)
    regression(name, [yr, 'I1:I10', '{11;12}', 'FALSE'], 2, 1)
    regression(name, [yr, 'I1:J10', '{11,12;13,14}'], 2, 1)
    regression(name, [yr, 'I1:J10', '{11,12;13,14}', 'FALSE'], 2, 1)
    regression(name, [yr, 'I1:K10', '{11,12,22}'], 1, 1)
    regression(name, [yr, 'I1:J10', '{11,12,13}'], 1, 1)
    regression(name, [yr, 'I1:I10', '15'], 1, 1)
    regression(name, [yr, '', '{11;12}'], 2, 1)
    regression(name, [yr, 'I1:I10', '"abc"'], 1, 1)
    regression(name, [yr, 'I1:I10', '{1;"a"}'], 2, 1)
    regression(name, [yr, 'I1:I3'], 1, 1)
    regression(name, ['{1,2,3}', '{1,2,3}', '{4,5}', '"abc"'], 1, 1)
    regression(name, ['H12:Q12', 'H13:Q13', '{11,12}'], 1, 2)
    regression(name, ['{5}', '{2}', '{3}'], 1, 1)

cells = [(addr, value) for addr, value in DATA]
os.makedirs(sys.argv[1], exist_ok=True)
batches = [formulas[i:i + BATCH] for i in range(0, len(formulas), BATCH)]
for b, batch in enumerate(batches, 1):
    with open(os.path.join(sys.argv[1], f'statdist-cells-{b}.tsv'), 'w', encoding='utf-8', newline='\n') as out:
        for addr, value in DATA:
            out.write(f'{SHEET}{b}\t{addr}\t{value}\n')
        for row, f in enumerate(batch, 1):
            out.write(f'{SHEET}{b}\tA{row}\t{f}\n')
print(f'{len(formulas)} probe formulas in {len(batches)} batches')
