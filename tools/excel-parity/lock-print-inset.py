#!/usr/bin/env python3
"""Lock or check plain-text print inset predictions (make-print-inset-fixture.py).

    python3 tools/excel-parity/lock-print-inset.py lock  CELLS.tsv GDI.tsv PRED.tsv
    python3 tools/excel-parity/lock-print-inset.py check PRED.tsv MEASURED.tsv

GDI.tsv (face size scale ppem digit w_pos w_neg w_zero) is written by `print-inset-gdi.fsx` from the
engine's hinted-advance tables. The rule, from the training export (Normal font Calibri 11, cell
fonts Aptos Narrow/Arial/Calibri/Georgia at 9/11/14 pt, 60 % and 100 %), all in 600-dpi device dots:

  pad    P = ceil(d / 4) + 4, d = the CELL font's widest digit at ppem round(size x scale x 600/72)
  left   x0 - cell_left = P
  right  cell_right - x0 - W = P + 1              W = the text's hinted GDI width at that ppem
  centre x0 - cell_left - (cell_width - W) / 2 = 3

Tolerances: left 0.1 dot, right [0.8, 1.3], centre [2.8, 3.7] (cell edges in the PDF are not all on
the dot grid). Check exits 1 on any failure or on a key-set mismatch.
"""
import csv
import math
import sys

D = 600.0 / 72.0
WIDTH_KEY = {'pos': 'w_pos', 'text': 'w_pos', 'neg': 'w_neg', 'zero': 'w_zero'}
BOUNDS = {'left': (-0.1, 0.1), 'right': (0.8, 1.3), 'center': (2.8, 3.7)}


def lock(cells, gdi, pred):
    g = {(r['face'], float(r['size']), float(r['scale'])): r
         for r in csv.DictReader(open(gdi, encoding='utf-8'), delimiter='\t')}
    with open(pred, 'w', encoding='utf-8') as out:
        out.write('sheet\trow\twidth\talign\tface\tsize\tscale\tkind\tppem\tdigit\tpad\tw_gdi\n')
        for r in csv.DictReader(open(cells, encoding='utf-8'), delimiter='\t'):
            m = g[(r['face'], float(r['size']), float(r['scale']))]
            d = int(m['digit'])
            pad = math.ceil(d / 4) + 4
            for w in r['widths'].split(','):
                out.write('\t'.join([r['sheet'], r['row'], w, r['align'], r['face'], r['size'], r['scale'],
                                     r['kind'], m['ppem'], m['digit'], str(pad), m[WIDTH_KEY[r['kind']]]]) + '\n')
    print(pred)


def check(pred, measured):
    p = {(r['sheet'], r['row'], r['width']): r for r in csv.DictReader(open(pred, encoding='utf-8'), delimiter='\t')}
    m = {(r['sheet'], r['row'], r['width']): r for r in csv.DictReader(open(measured, encoding='utf-8'), delimiter='\t')}
    missing, extra = sorted(set(p) - set(m)), sorted(set(m) - set(p))
    fails = []
    worst = {}
    for k in sorted(set(p) & set(m)):
        a, b = p[k], m[k]
        pad, w = float(a['pad']), float(a['w_gdi'])
        x0 = float(b['left_inset']) * D
        cw = float(b['cell_width']) * D
        res = {'left': x0 - pad, 'right': cw - x0 - w - pad, 'center': x0 - (cw - w) / 2}[a['align']]
        lo, hi = BOUNDS[a['align']]
        key = (a['align'], a['face'])
        worst.setdefault(key, [res, res])
        worst[key] = [min(worst[key][0], res), max(worst[key][1], res)]
        if not lo <= res <= hi:
            fails.append((k, a['align'], a['face'], a['size'], a['scale'], a['kind'], round(res, 3)))
    for (align, face), (lo, hi) in sorted(worst.items()):
        print(f'{align:6s} {face:14s} residual [{lo:+.3f}, {hi:+.3f}] dots')
    print(f'cells {len(p)} measured {len(m)} missing {len(missing)} extra {len(extra)} failures {len(fails)}')
    for f in fails[:40]:
        print('  FAIL', *f)
    ok = not fails and not missing and not extra
    print('PASS' if ok else 'FAIL')
    sys.exit(0 if ok else 1)


if __name__ == '__main__':
    if sys.argv[1] == 'lock':
        lock(*sys.argv[2:5])
    else:
        check(*sys.argv[2:4])
