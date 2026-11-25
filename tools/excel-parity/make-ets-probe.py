#!/usr/bin/env python3
"""Probe cells for FORECAST.ETS and its kin (FnStatistical.fs), as eval-cells-on-vm.sh takes them.

    python3 tools/excel-parity/make-ets-probe.py > ets-cells.txt

One line per cell: sheet, address, formula. Series written in as array constants (values and a timeline):
without a season (seasonality 0), with a given season length and detected (seasonality 1 or omitted); for
each FORECAST.ETS.STAT's eight statistics (alpha, beta, gamma, MASE, SMAPE, MAE, RMSE, step), forecasts
one to several steps ahead and between the steps, FORECAST.ETS.CONFINT at confidence levels, and
FORECAST.ETS.SEASONALITY. Then the timeline rules: gaps (data_completion 1 and 0), repeated dates
(aggregation 1 to 7), an unsorted timeline, steps that are not constant, targets before the end, the
argument errors. Deterministic (fixed seed): ets-excel.tsv holds Excel 2019's values; function-parity-engine.fsx
compares the engine.
"""
import math
import random

R = random.Random(20261126)
cells = []


def cell(f):
    cells.append(f)


def arr(xs):
    return '{' + ','.join(fmt(x) for x in xs) + '}'


def fmt(x):
    if isinstance(x, int):
        return str(x)
    s = '%.6g' % x
    return s


def series(kind, n, m=0):
    level = R.uniform(-50, 200)
    slope = R.choice([0, R.uniform(-3, 3)])
    amp = R.uniform(1, 30)
    noise = R.choice([0.1, 1, 5, 15])
    phase = R.random()
    out = []
    walk = 0
    for t in range(n):
        walk += R.gauss(0, noise) if kind == 'walk' else 0
        s = amp * math.sin(2 * math.pi * (t / m + phase)) if m else 0
        out.append(round(level + slope * t + s + walk + R.gauss(0, noise), 2))
    return out


def probe(ys, ts, seas=None, more=True):
    n = len(ys)
    ya, ta = arr(ys), arr(ts)
    tail = '' if seas is None else ',%d' % seas
    for k in range(1, 9):
        cell('=FORECAST.ETS.STAT(%s,%s,%d%s)' % (ya, ta, k, tail))
    step = ts[-1] - ts[-2]
    for h in ([1, 2, 5] if more else [1]):
        cell('=FORECAST.ETS(%s,%s,%s%s)' % (fmt(ts[-1] + h * step), ya, ta, tail))
    if more:
        cell('=FORECAST.ETS(%s,%s,%s%s)' % (fmt(ts[-1] + 1.5 * step), ya, ta, tail))
        cell('=FORECAST.ETS.CONFINT(%s,%s,%s,0.95%s)' % (fmt(ts[-1] + step), ya, ta, tail))
        cell('=FORECAST.ETS.CONFINT(%s,%s,%s,0.8%s)' % (fmt(ts[-1] + 3 * step), ya, ta, tail))
    cell('=FORECAST.ETS.SEASONALITY(%s,%s%s)' % (ya, ta, tail))


# No season.
for i in range(30):
    n = R.choice([4, 5, 8, 12, 20, 30, 45])
    ys = series(R.choice(['walk', 'trend']), n)
    probe(ys, list(range(1, n + 1)), 0)
# Seasons given.
for i in range(30):
    m = R.choice([2, 3, 4, 6, 7, 12])
    n = R.choice([2 * m + 1, 3 * m, 4 * m + 2, 5 * m])
    ys = series(R.choice(['walk', 'trend']), n, m)
    probe(ys, list(range(1, n + 1)), m)
# Seasons detected.
for i in range(40):
    m = R.choice([0, 3, 4, 5, 7, 12])
    n = R.choice([12, 24, 36, 48]) if m else R.choice([10, 20, 40])
    ys = series(R.choice(['walk', 'trend']), n, m)
    probe(ys, list(range(1, n + 1)), R.choice([None, 1]))
# Timelines: dates, a step other than 1, gaps, repeats, unsorted, uneven.
base = [12, 15, 14, 18, 21, 19, 25, 27, 26, 30, 33, 31, 36, 38, 37, 41]
d0 = 43831
probe(base, [d0 + 7 * k for k in range(16)], 0)
probe(base, [10 * k for k in range(16)], 0)
probe(base, [0.5 * k for k in range(16)], 0)
gap = [k for k in range(1, 17) if k not in (5, 9, 10)]
for dc in (1, 0):
    ys = [base[k - 1] for k in gap]
    for s in range(1, 9):
        cell('=FORECAST.ETS.STAT(%s,%s,%d,0,%d)' % (arr(ys), arr(gap), s, dc))
    cell('=FORECAST.ETS(17,%s,%s,0,%d)' % (arr(ys), arr(gap), dc))
rep = [1, 2, 2, 3, 4, 4, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]
for agg in range(1, 8):
    for s in (1, 6, 8):
        cell('=FORECAST.ETS.STAT(%s,%s,%d,0,1,%d)' % (arr(base), arr(rep), s, agg))
    cell('=FORECAST.ETS(14,%s,%s,0,1,%d)' % (arr(base), arr(rep), agg))
perm = list(range(16))
R.shuffle(perm)
probe([base[p] for p in perm], [p + 1 for p in perm], 0, more=False)
for f in ['FORECAST.ETS(20,{1,2,3,4,5,6},{1,2,3,4,5,7})', 'FORECAST.ETS(20,{1,2,3,4,5,6},{1,2,4,5,7,8})',
          'FORECAST.ETS(20,{1,2,3,4,5,6},{1,3,5,7,9,10})', 'FORECAST.ETS(20,{1,2,3,4,5,6},{1,2,3,4,5,6.5})',
          'FORECAST.ETS(3,{1,2,3,4,5,6},{1,2,3,4,5,6})', 'FORECAST.ETS(6,{1,2,3,4,5,6},{1,2,3,4,5,6})',
          'FORECAST.ETS(0,{1,2,3,4,5,6},{1,2,3,4,5,6})', 'FORECAST.ETS(7,{1,2,3,4,5,6},{1,2,3,4,5})',
          'FORECAST.ETS(7,{1,2,3},{1,2,3})', 'FORECAST.ETS(7,{1,2},{1,2})', 'FORECAST.ETS(7,{1,2,3,4},{1,2,3,4})',
          'FORECAST.ETS(7,{1,2,3,4,5,6},{1,2,3,4,5,6},-1)', 'FORECAST.ETS(7,{1,2,3,4,5,6},{1,2,3,4,5,6},8761)',
          'FORECAST.ETS(7,{1,2,3,4,5,6},{1,2,3,4,5,6},4)', 'FORECAST.ETS(7,{1,2,3,4,5,6},{1,2,3,4,5,6},3)',
          'FORECAST.ETS(7,{1,2,3,4,5,6},{1,2,3,4,5,6},2.7)', 'FORECAST.ETS(7,{1,2,3,4,5,6},{1,2,3,4,5,6},0,2)',
          'FORECAST.ETS(7,{1,2,3,4,5,6},{1,2,3,4,5,6},0,1,0)', 'FORECAST.ETS(7,{1,2,3,4,5,6},{1,2,3,4,5,6},0,1,8)',
          'FORECAST.ETS(7,{1,2,3,4,5,6},{1,2,3,4,5,6},"x")', 'FORECAST.ETS("x",{1,2,3,4,5,6},{1,2,3,4,5,6})',
          'FORECAST.ETS(7,{1,2,"a",4,5,6},{1,2,3,4,5,6})', 'FORECAST.ETS(7,{1,2,3,4,5,6},{1,2,"a",4,5,6})',
          'FORECAST.ETS(7,{5,5,5,5,5,5},{1,2,3,4,5,6})', 'FORECAST.ETS.STAT({5,5,5,5,5,5},{1,2,3,4,5,6},1)',
          'FORECAST.ETS.STAT({5,5,5,5,5,5},{1,2,3,4,5,6},4)', 'FORECAST.ETS.STAT({1,2,3,4,5,6},{1,2,3,4,5,6},9)',
          'FORECAST.ETS.STAT({1,2,3,4,5,6},{1,2,3,4,5,6},0)', 'FORECAST.ETS.CONFINT(7,{1,2,3,4,5,6},{1,2,3,4,5,6},1)',
          'FORECAST.ETS.CONFINT(7,{1,2,3,4,5,6},{1,2,3,4,5,6},0)', 'FORECAST.ETS.CONFINT(7,{1,2,3,4,5,6},{1,2,3,4,5,6})',
          'FORECAST.ETS.CONFINT(7,{1,2,4,4,5,7},{1,2,3,4,5,6},0.5)', 'FORECAST.ETS.SEASONALITY({1,2,3,4,5,6},{1,2,3,4,5,6})',
          'FORECAST.ETS(1000,{1,2,4,4,5,7},{1,2,3,4,5,6})', 'FORECAST.ETS(7,{1,2,4,4,5,7},{6,5,4,3,2,1})',
          'FORECAST.ETS(7,{1;2;4;4;5;7},{1;2;3;4;5;6})', 'FORECAST.ETS(7,{1,2;4,4;5,7},{1,2;3,4;5,6})',
          'FORECAST.ETS(7,{1,2,4,4,5,7},{1;2;3;4;5;6})', 'FORECAST.ETS(7.5,{1,2,4,4,5,7},{1,2,3,4,5,6})',
          'FORECAST.ETS(-5,{1,2,4,4,5,7},{-12,-11,-10,-9,-8,-7})']:
    cell('=' + f)

for i, f in enumerate(cells):
    print('S\tA%d\t%s' % (i + 1, f))
