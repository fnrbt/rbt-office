#!/usr/bin/env python3
"""Generate an XIRR six-decimal threshold probe, its pre-measurement predictions,
and (after Excel and the engine have run) the combined comparison table.

    python3 tools/excel-parity/make-xirr-threshold-probe.py cells /tmp/xirr-threshold-cells.tsv
    python3 tools/excel-parity/make-xirr-threshold-probe.py predictions /tmp/xirr-threshold-predictions.tsv
    tools/excel-parity/eval-cells-on-vm.sh /tmp/xirr-threshold-cells.tsv /tmp/xirr-threshold-excel.tsv
    dotnet fsi tools/excel-parity/xirr-cells-engine.fsx /tmp/xirr-threshold-cells.tsv > /tmp/xirr-threshold-engine.tsv
    python3 tools/excel-parity/make-xirr-threshold-probe.py results \
        --excel /tmp/xirr-threshold-excel.tsv --engine /tmp/xirr-threshold-engine.tsv

The cases are public synthetic cash flows, not workbook inputs. Exact roots are
solved in decimal with 80 digits; predictions were derived from those roots and
the observed Excel stopping step before measurement and are kept in CASES so
the predictions cannot drift with the measurements. `predictions` prints them.

Two kinds of case separate stopping from root selection:

* single-root cases put the root at 1e-7, 1e-8, 3e-9, 1e-8 or 1e-7 either side of
  the six-decimal half-way value 0.1000005 (and its negative mirror), plus an
  irregular three-flow and a monthly waterfall with the same root;
* multi-root cases put one root at 0.100000503 (two positive roots) or
  -0.100000503 (opposite-signed roots). Excel rejects the first and, without a
  guess that converges to the positive root, the second; the engine still finds
  a valid root.

Results: xirr-threshold-results.tsv.
"""
import sys
from decimal import Decimal, getcontext, ROUND_HALF_UP

getcontext().prec = 80
DAY0 = 43831  # 2020-01-01, as in the other XIRR probes


def D(x):
    return Decimal(str(x))


def short(x):
    return format(x, '.15g')


def fmt(x, digits=12):
    s = format(x, f'.{digits}f')
    if '.' in s:
        s = s.rstrip('0').rstrip('.')
    return s or '0'


def exact_roots(flows, days):
    """All roots above -1, located by a float scan and refined in decimal.

    A float scan alone can bracket the wrong side of a root when the present
    value cancels large flows (the values are ~1e-8 against 1e9 inputs), so each
    candidate is refined with decimal Newton, whose derivative is analytic.
    """
    import math

    fv = [float(v) for v in flows]
    te = [(t - days[0]) / 365.0 for t in days]

    def f(r):
        l = math.log1p(r)
        return sum(v * math.exp(-e * l) for v, e in zip(fv, te))

    def refine(start):
        r = D(start)
        for _ in range(120):
            l = (1 + r).ln()
            pv = Decimal(0)
            dpv = Decimal(0)
            for v, e in zip(flows, te):
                term = D(v) / (D(e) * l).exp()
                pv += term
                dpv -= D(e) * term / (1 + r)
            if dpv == 0:
                return r
            step = pv / dpv
            nxt = r - step
            if nxt <= -1:
                nxt = (r - 1) / 2
            if abs(step) <= D('1e-40') * max(Decimal(1), abs(nxt)):
                return nxt
            r = nxt
        return r

    xs = set()
    for i in range(1500):
        xs.add(-0.999999999 + i * (0.999999999 / 1500))
    for i in range(3000):
        xs.add((i / 3000.0) ** 3 * 50.0)
    xs.add(0.0)
    xs = sorted(xs)
    ys = [f(x) for x in xs]
    found = []
    for i in range(len(xs) - 1):
        a, b = ys[i], ys[i + 1]
        if a == 0.0:
            lo = hi = xs[i]
        elif a * b < 0.0:
            lo, hi = xs[i], xs[i + 1]
            for _ in range(90):
                mid = (lo + hi) / 2.0
                if f(lo) * f(mid) <= 0.0:
                    hi = mid
                else:
                    lo = mid
        else:
            continue
        # Decimal refinement of the float bracket. The float sign can be wrong
        # when the present value is tiny against the inputs; Newton does not
        # depend on that sign.
        root = refine((lo + hi) / 2.0)
        if all(abs(root - q) > D('1e-20') * max(Decimal(1), abs(root)) for q in found):
            found.append(root)
    return sorted(found)


def excel_round6(v):
    v = D(v)
    return (v * Decimal(1000000)).quantize(Decimal(1), rounding=ROUND_HALF_UP) / Decimal(1000000)


def dist_to_halfway(v):
    s = D(v) * Decimal(1000000)
    return (Decimal('0.5') - abs(s - s.to_integral_value(rounding=ROUND_HALF_UP))) / Decimal(1000000)


def two_flow(r0, v0='-1000', days=(0, 365)):
    return [v0, fmt(-D(v0) * (1 + D(r0)), 10)], list(days)


def annual_pair(r_a, r_b, v0='-1000'):
    """Annual flows v0, a, b with roots r_a and r_b (x = 1/(1+r))."""
    xa, xb = 1 / (1 + D(r_a)), 1 / (1 + D(r_b))
    b = D(v0) / (xa * xb)
    a = -b * (xa + xb)
    return [v0, fmt(a, 12), fmt(b, 12)], [0, 365, 730]


def solve_last(flows, days, r0):
    """Replace the last flow so that r0 is a root; earlier flows stay."""
    r0 = D(r0)
    total = sum(D(v) / (((D(t) - D(days[0])) / D(365)) * (1 + r0).ln()).exp()
                for v, t in zip(flows[:-1], days[:-1]))
    last = -total * (((D(days[-1]) - D(days[0])) / D(365)) * (1 + r0).ln()).exp()
    return flows[:-1] + [format(last, '.15g')]


# name, flows, day offsets, guess, target root, predicted Excel ROUND6,
# predicted engine ROUND6, note.
# Predictions: one-year Excel stopping errors measured in the impact probe were
# -5.3e-9 (root 0.1000005) and -3.35e-9 (root -0.1); the predicted ROUND6 below
# adds that bias to the exact root. A "?" is deliberately unregistered.
CASES = [
    ('T1-1e-7-below', *two_flow('0.10000040'), 'default', '0.10000040', '0.1', '0.1',
     'single root 1e-7 below the half-way value; bias rule predicts down'),
    ('T2-1e-8-below', *two_flow('0.10000049'), 'default', '0.10000049', '0.1', '0.1',
     'single root 1e-8 below; bias rule predicts down'),
    ('T3-at-boundary', *two_flow('0.10000050'), 'default', '0.10000050', '0.1', '0.100001',
     'single root on the half-way value; bias rule predicts down, engine exact up'),
    ('T4-3e-9-above', *two_flow('0.100000503'), 'default', '0.100000503', '0.1', '0.100001',
     'single root 3e-9 above; bias rule still predicts down'),
    ('T4-guess0.5', *two_flow('0.100000503'), '0.5', '0.100000503', '0.1', '0.100001',
     'same root as T4 from guess 0.5; impact boundary guesses stopped below, so predict down'),
    ('T5-1e-8-above', *two_flow('0.10000051'), 'default', '0.10000051', '0.100001', '0.100001',
     'single root 1e-8 above; bias rule predicts up'),
    ('T6-1e-7-above', *two_flow('0.10000060'), 'default', '0.10000060', '0.100001', '0.100001',
     'single root 1e-7 above; bias rule predicts up'),
    ('T7-neg-at-boundary', *two_flow('-0.10000050'), 'default', '-0.10000050', '-0.100001', '-0.100001',
     'negative half-way value; bias rule predicts away from zero'),
    ('T8-neg-1e-8-above', *two_flow('-0.10000049'), 'default', '-0.10000049', '-0.1', '-0.1',
     'negative root 1e-8 toward zero; bias rule predicts toward zero'),
    ('T9-neg-1e-8-below', *two_flow('-0.10000051'), 'default', '-0.10000051', '-0.100001', '-0.100001',
     'negative root 1e-8 away from zero'),
    ('T10-large-at-boundary', *two_flow('0.1000005', v0='-1000000000'), 'default', '0.1000005', '0.1', '0.100001',
     'same boundary at 1e6 scale'),
    ('M1-two-positive', *annual_pair('0.100000503', '0.3'), 'default', '0.100000503', '#NUM!', '0.100001',
     'two positive roots; Excel rejects, engine finds the near-boundary one'),
    ('M2-opposite-sign', *annual_pair('-0.100000503', '0.3'), 'default', '-0.100000503', '0.3', '-0.100001',
     'opposite-signed roots at the default guess; prediction from the domain probe was negative for the engine'),
    ('M2-opposite-sign-guess0', *annual_pair('-0.100000503', '0.3'), '0', '-0.100000503', '0.3', '-0.100001',
     'same flows at guess 0; the engine brackets from just above -1 and picks the negative root'),
    ('M2-opposite-sign-guess0.5', *annual_pair('-0.100000503', '0.3'), '0.5', '-0.100000503', '0.3', '0.3',
     'same flows with a positive guess; both pick the positive root'),
    ('M3-irregular-single', solve_last(['-1000', '300', '400'], [0, 200, 500], '0.100000503'), [0, 200, 500],
     'default', '0.100000503', '0.1', '0.100001',
     'irregular three-flow single root above the boundary; bias rule predicts down'),
    ('M4-monthly-single', solve_last(['-1000'] + ['10'] * 12, [0, 31, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335, 366], '0.100000503'),
     [0, 31, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335, 366], 'default', '0.100000503', '?', '0.100001',
     'monthly waterfall with the last flow solved to that root; no Excel prediction'),
    ('M5-monthly-below', solve_last(['-1000'] + ['10'] * 12, [0, 31, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335, 366], '0.10000049'),
     [0, 31, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335, 366], 'default', '0.10000049', '?', '0.1',
     'monthly waterfall 1e-8 below; no Excel prediction'),
]

HEADER = ['case', 'guess', 'exact_roots', 'target_root', 'dist_to_halfway', 'exact_round6',
          'pred_excel_round6', 'pred_engine_round6', 'note']


def formula(values, days, guess):
    v = '{' + ','.join(values) + '}'
    d = '{' + ','.join(str(DAY0 + int(x)) for x in days) + '}'
    g = '' if guess == 'default' else ',' + guess
    return f'=XIRR({v},{d}{g})'


def write_cells(path):
    with open(path, 'w', encoding='utf-8', newline='\n') as out:
        for row, (name, flows, days, guess, target, _pe, _pn, _note) in enumerate(CASES, 1):
            f = formula(flows, days, guess)
            out.write(f'Thr\tB{row}\t{f}\n')
            out.write(f'Thr\tC{row}\t=ROUND({f[1:]},6)\n')
            out.write(f'Thr\tD{row}\t{name}\n')


def write_predictions(path):
    rows = []
    for name, flows, days, guess, target, pe, pn, note in CASES:
        roots = exact_roots(flows, days)
        target_d = D(target)
        actual = min(roots, key=lambda q: abs(q - target_d))
        rows.append([name, guess, ';'.join(short(r) for r in roots), short(actual),
                     f'{dist_to_halfway(actual):.3E}', str(excel_round6(actual)), pe, pn, note])
    with open(path, 'w', encoding='utf-8', newline='\n') as out:
        out.write('\t'.join(HEADER) + '\n')
        for r in rows:
            out.write('\t'.join(r) + '\n')


def parse_exact(text):
    sign = ''
    if text.startswith('-'):
        sign, text = '-', text[1:]
    hi, lo, e = (int(x) for x in text.split(':'))
    return D(sign + str(hi * 67108864 + lo)) * (D(2) ** e)


def read_excel(path):
    out = {}
    for line in open(path, encoding='utf-8'):
        p = line.rstrip('\n').split('\t')
        if len(p) < 5 or p[0] != 'Thr':
            continue
        kind = p[2]
        if kind == 'n':
            out[p[1]] = ('n', parse_exact(p[5]))
        elif kind == 'e':
            out[p[1]] = ('e', p[3])
        else:
            out[p[1]] = (kind, p[3])
    return out


def read_engine(path):
    out = {}
    for line in open(path, encoding='utf-8'):
        p = line.rstrip('\n').split('\t')
        if len(p) == 3 and p[0] == 'Thr':
            out[p[1]] = p[2]
    return out


def canon6(text):
    """A six-decimal value as the shortest decimal string, or an error."""
    if isinstance(text, str) and text.startswith('#'):
        return text
    return fmt(excel_round6(text), 6)


def write_results(excel_path, engine_path):
    excel = read_excel(excel_path)
    engine = read_engine(engine_path)
    cols = ['case', 'guess', 'exact_roots', 'target_root', 'excel_value', 'excel_round6',
            'engine_value', 'engine_round6', 'excel_root', 'engine_root', 'same_root',
            'round6_agrees', 'pred_excel_round6', 'pred_engine_round6', 'pred_excel_ok',
            'pred_engine_ok', 'cause', 'note']
    print('\t'.join(cols))
    for row, (name, flows, days, guess, target, pe, pn, note) in enumerate(CASES, 1):
        roots = exact_roots(flows, days)
        e_direct = excel[f'B{row}']
        e_round = excel[f'C{row}']
        n_direct = engine[f'B{row}']
        n_round = engine[f'C{row}']
        have_e = e_direct[0] == 'n'
        have_n = not n_direct.startswith('#')
        e_val = e_direct[1]
        n_val = n_direct

        def nearest(v):
            if not isinstance(v, Decimal):
                return None
            return min(roots, key=lambda q: abs(q - v))

        er = nearest(e_val) if have_e else None
        nr = nearest(D(n_val)) if have_n else None
        same = 'yes' if er is not None and nr is not None and abs(er - nr) < D('1e-6') else 'no'
        e_r6 = canon6(e_round[1]) if e_round[0] == 'n' else e_round[1]
        n_r6 = canon6(D(n_round)) if have_n else n_round
        if have_e and have_n:
            agrees = 'yes' if e_r6 == n_r6 else 'no'
        else:
            agrees = 'n/a'
        if not have_e:
            cause = 'Excel rejects (engine keeps a valid root)'
        elif not have_n:
            cause = 'engine rejects (not observed)'
        elif same == 'no':
            cause = 'root selection'
        elif agrees == 'no':
            cause = 'stopping, same root'
        else:
            cause = 'agree'
        pe_ok = '-' if pe == '?' else ('yes' if e_r6 == canon6(pe) else 'no')
        pn_ok = '-' if pn == '?' else ('yes' if n_r6 == canon6(pn) else 'no')
        vals = [name, guess, ';'.join(short(r) for r in roots), target,
                fmt(e_val, 12) if have_e else e_val, e_r6,
                n_val, n_r6,
                short(er) if er is not None else '', short(nr) if nr is not None else '',
                same, agrees, pe, pn, pe_ok, pn_ok, cause, note]
        print('\t'.join(vals))


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print(__doc__, file=sys.stderr)
        sys.exit(2)
    mode = sys.argv[1]
    if mode == 'cells':
        write_cells(sys.argv[2])
    elif mode == 'predictions':
        write_predictions(sys.argv[2])
    elif mode == 'results':
        opts = dict(zip(sys.argv[2::2], sys.argv[3::2]))
        write_results(opts['--excel'], opts['--engine'])
    else:
        print(__doc__, file=sys.stderr)
        sys.exit(2)
