#!/usr/bin/env python3
"""Evaluate the locked border/fill models against the measured discriminating fixture.

    python3 tools/excel-parity/analyze-print-border-fill.py OUT_PREFIX MEASURED.tsv [MEASURED.tsv ...]

Each MEASURED.tsv is measure-print-border-fill.py's long table. Prints the model verdicts, writes
`OUT_PREFIX-cases.tsv` (one row per non-hash left-aligned fill cell with every model's prediction and
signed error) and `OUT_PREFIX-summary.tsv`.

Models from print-border-fill-predictions.tsv (no fitted per-face values):
  M_HALF   c = s_right / 2   (s_right = physical right border strip width; 0 on `none`)
  M_FIXED  c = 0.24
  M_NONE   c = 0
  count = floor((W - L - R + c - w) / a), with w/a the 600-dpi predicted advances. The measured
  number width (Excel amount_x1 - amount_x0) is reported as a second basis that removes the text
  width model from the boundary measurement.
"""
import collections
import importlib.util
import os
import statistics
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location('countpredict', os.path.join(HERE, 'predict-print-count.py'))
cp = importlib.util.module_from_spec(spec)
spec.loader.exec_module(cp)

FACES = {
    'arial': ('Arial', 11.0, 60),
    'aptos': ('Aptos Narrow', 14.0, 100),
    'calibri': ('Calibri', 11.0, 100),
}
STYLES = ['none', 'thin', 'medium', 'thick']
CHARS = ' .-0123,45'


def advances(face, size, scale):
    g, _, _ = cp.font_metrics(os.path.join(cp.FONT_DIR, cp.FONT_FILES[(face, False)]))
    dots = round(size * scale / 100.0 * 600.0 / 72.0)
    return {c: round(g[c] * dots) * 72.0 / 600.0 for c in CHARS}


def load(path):
    rows = [l.rstrip('\n').split('\t') for l in open(path, encoding='utf-8') if l.strip()]
    h = rows[0]
    return [dict(zip(h, r)) for r in rows[1:]]


def fnum(r, k):
    v = r.get(k, '')
    return None if v == '' else float(v)


def main():
    prefix, measured = sys.argv[1], sys.argv[2:]
    cases = []
    interval_rows = []
    inv_rows = []
    trans_rows = []
    for path in measured:
        rows = load(path)
        if not rows:
            continue
        key = rows[0]['face']
        face, size, scale = FACES[key]
        adv = advances(face, size, scale)
        n1 = {(r['sheet'], r['value'], r['width']): r
              for r in rows if r['format'] == 'N1_plain_0' and r['amount']}
        # per-style control insets and text positions
        by_key = collections.defaultdict(dict)
        for r in rows:
            by_key[(r['align'], r['width'], r['value'], r['format'])][r['border']] = r
        inv = collections.defaultdict(list)
        for k, d in by_key.items():
            base = d.get('none')
            if not base:
                continue
            for style in STYLES[1:]:
                r = d.get(style)
                if not r:
                    continue
                for f in ('amount_x0', 'amount_x1', 'amount_gap_left', 'amount_gap_right'):
                    a, b = fnum(base, f), fnum(r, f)
                    if a is not None and b is not None:
                        inv[(style, 'text' if f.startswith('amount_x') else 'inset', f)].append(abs(a - b))
        for (style, kind, f), v in inv.items():
            inv_rows.append((face, style, kind, f, max(v), statistics.mean(v)))

        cells = []
        for r in rows:
            if r['align'] != 'left' or r['format'] not in ('S2_space_after', 'D2_dot_after', 'F8_suffix_empty'):
                continue
            if r['fill_char'] not in adv or not r['amount'] or r['hash'] == 'yes':
                continue
            lr = n1.get((r['sheet'], r['value'], r['width']))
            rr = n1.get((r['sheet'].replace('_left', '_right'), r['value'], r['width']))
            if not lr or not rr:
                continue
            L = float(lr['amount_x0']) - float(lr['cell_left'])
            R = float(rr['cell_right']) - float(rr['amount_x1'])
            X = float(r['cell_width']) - L - R
            n = int(r['count'])
            a = adv[r['fill_char']]
            s = fnum(r, 'right_strip_w') or 0.0
            w_pred = sum(adv[c] for c in r['amount'])
            w_meas = float(r['amount_x1']) - float(r['amount_x0'])
            preds = {
                'HALF': max(0, int((X + s / 2.0 - w_pred) // a)),
                'FIXED': max(0, int((X + 0.24 - w_pred) // a)),
                'NONE': max(0, int((X - w_pred) // a)),
            }
            preds_meas = {
                'HALF': max(0, int((X + s / 2.0 - w_meas) // a)),
                'FIXED': max(0, int((X + 0.24 - w_meas) // a)),
                'NONE': max(0, int((X - w_meas) // a)),
            }
            cells.append({'face': face, 'style': r['border'], 'align': r['align'],
                          'width': r['width'], 'value': r['value'], 'format': r['format'],
                          'count': n, 'X': X, 'w_pred': w_pred, 'w_meas': w_meas, 'a': a, 's': s,
                          'preds': preds, 'preds_meas': preds_meas,
                          'x0': fnum(r, 'amount_x0'), 'x1': fnum(r, 'amount_x1'),
                          'composite': fnum(r, 'composite_x1'), 'residual': fnum(r, 'residual')})
        cases.extend(cells)
        # feasible correction intervals, both width bases
        for style in STYLES:
            sc = [c for c in cells if c['style'] == style]
            if not sc:
                continue
            for basis in ('pred', 'meas'):
                lo, hi = -99.0, 99.0
                for c in sc:
                    w = c['w_pred'] if basis == 'pred' else c['w_meas']
                    lo = max(lo, c['count'] * c['a'] + w - c['X'])
                    hi = min(hi, (c['count'] + 1) * c['a'] + w - c['X'])
                interval_rows.append((face, style, basis, lo, hi, len(sc)))
            # residual and transition evidence
            resid = [c['residual'] for c in sc if c['residual'] is not None]
            inv_rows.append((face, style, 'residual', 'min', min(resid), statistics.mean(resid)))
        # transition content width per fill series (X where count first reaches n+1)
        series = collections.defaultdict(list)
        for c in cells:
            series[(c['style'], c['value'], c['format'])].append(c)
        trans = {}
        for (style, value, fmt), rr in series.items():
            rr.sort(key=lambda c: c['X'])
            for i in range(len(rr) - 1):
                a, b = rr[i], rr[i + 1]
                if b['count'] == a['count'] + 1:
                    t = a['X'] + (a['count'] + 1 - a['count']) / (b['count'] - a['count']) * (b['X'] - a['X'])
                    trans[(value, fmt, a['count'])] = t
        for k2, t in trans.items():
            trans_rows.append((face, style, k2[0], k2[1], k2[2], t))

    # cases TSV
    with open(prefix + '-cases.tsv', 'w', encoding='utf-8') as f:
        f.write('face\tstyle\twidth\tvalue\tformat\tcount\tX\tw_pred\tw_meas\ta\ts_right\t'
                'pred_HALF\terr_HALF\tpred_FIXED\terr_FIXED\tpred_NONE\terr_NONE\t'
                'mpred_HALF\tmerr_HALF\tmpred_FIXED\tmerr_FIXED\tmpred_NONE\tmerr_NONE\t'
                'x0\tx1\tcomposite\tresidual\n')
        for c in cases:
            f.write('\t'.join([
                c['face'], c['style'], c['width'], c['value'], c['format'], str(c['count']),
                f'{c["X"]:.3f}', f'{c["w_pred"]:.3f}', f'{c["w_meas"]:.3f}', f'{c["a"]:.3f}',
                f'{c["s"]:.3f}'] + [
                f'{c["preds"][m]}\t{c["preds"][m]-c["count"]:+d}' for m in ('HALF', 'FIXED', 'NONE')] + [
                f'{c["preds_meas"][m]}\t{c["preds_meas"][m]-c["count"]:+d}' for m in ('HALF', 'FIXED', 'NONE')] + [
                '' if c['x0'] is None else f'{c["x0"]:.3f}',
                '' if c['x1'] is None else f'{c["x1"]:.3f}',
                '' if c['composite'] is None else f'{c["composite"]:.3f}',
                '' if c['residual'] is None else f'{c["residual"]:.3f}']) + '\n')
    # summary TSV
    with open(prefix + '-summary.tsv', 'w', encoding='utf-8') as f:
        f.write('face\tmodel\tbasis\texact\tover\tunder\tmax_abs_err\n')
        for face in FACES.values():
            fc = [c for c in cases if c['face'] == face[0]]
            if not fc:
                continue
            for basis, field in (('locked', 'preds'), ('measured_w', 'preds_meas')):
                for model in ('HALF', 'FIXED', 'NONE'):
                    errs = [c[field][model] - c['count'] for c in fc]
                    f.write(f'{face[0]}\t{model}\t{basis}\t{sum(e == 0 for e in errs)}\t'
                            f'{sum(e > 0 for e in errs)}\t{sum(e < 0 for e in errs)}\t{max(map(abs, errs))}\n')

    print('== correction interval c in count=floor((X+c-w)/a)')
    print('   identical across none/thin/medium/thick means border weight does not enter')
    for face, style, basis, lo, hi, n in interval_rows:
        print(f'   {face:13s} {style:6s} {basis:4s} n={n:3d} c=[{lo:+.3f},{hi:+.3f}]'
              + ('  EMPTY' if hi < lo else ''))
    by_face = collections.defaultdict(dict)
    for face, style, basis, lo, hi, n in interval_rows:
        by_face[(face, basis)][style] = (lo, hi)
    for (face, basis), d in sorted(by_face.items()):
        lo = max(v[0] for v in d.values())
        hi = min(v[1] for v in d.values())
        print(f'   INTERSECT {face:13s} {basis:4s} [{lo:+.3f},{hi:+.3f}]' + ('  EMPTY' if hi < lo else ''))
    print('\n== model counts')
    counts = collections.defaultdict(lambda: [0, 0, 0, 0])
    for c in cases:
        for field, basis in (('preds', 'pred_w'), ('preds_meas', 'meas_w')):
            for m in ('HALF', 'FIXED', 'NONE'):
                e = c[field][m] - c['count']
                counts[(c['face'], m, basis)][0 if e == 0 else (1 if e > 0 else 2)] += 1
                counts[(c['face'], m, basis)][3] = max(counts[(c['face'], m, basis)][3], abs(e))
    for k, v in sorted(counts.items()):
        print(f'   {k[0]:13s} {k[1]:5s} {k[2]:7s} exact {v[0]:3d} over {v[1]:3d} under {v[2]:3d} max {v[3]}')
    print('\n== invariance vs none (max pt)')
    for face, style, kind, fld, mx, mean in inv_rows:
        if kind == 'residual':
            print(f'   {face:13s} {style:6s} residual min {mx:.3f} mean {mean:.3f}')
        else:
            print(f'   {face:13s} {style:6s} {kind:5s} {fld:18s} max {mx:.3f}')
    print('\n== transition content-width shift across weights (pt)')
    tr = collections.defaultdict(dict)
    for face, style, value, fmt, n, t in trans_rows:
        tr[(face, value, fmt, n)][style] = t
    worst = 0.0
    for k, d in sorted(tr.items()):
        if 'none' not in d:
            continue
        shifts = {s: d[s] - d['none'] for s in d if s != 'none'}
        if shifts:
            m = max(abs(v) for v in shifts.values())
            worst = max(worst, m)
            print(f'   {k[0]:13s} {k[1]:6s} {k[2]:16s} n={k[3]} shifts ' +
                  ' '.join(f'{s}={v:+.3f}' for s, v in sorted(shifts.items())))
    print(f'   WORST shift {worst:.3f} pt (FIXED/NONE predict 0; HALF predicts |ds/2|)')
    print('\n== constant correction sweep (same c for every face/style/weight)')
    for basis, field in (('pred_w', 'w_pred'), ('meas_w', 'w_meas')):
        for ctest in (0.24, 0.25, 0.26, 0.28, 0.30, 0.32, 0.33, 0.36):
            bad = collections.Counter()
            tot = 0
            for c in cases:
                tot += 1
                pred = max(0, int((c['X'] + ctest - c[field]) // c['a']))
                if pred != c['count']:
                    bad[c['face']] += 1
            print(f'   c={ctest:.2f} {basis}: ' + ' '.join(f'{f}={n}' for f, n in sorted(bad.items()))
                  + (f' (0/{tot} failures)' if not bad else f' total {sum(bad.values())}/{tot}'))
    print('\n== VERDICTS')
    for face in FACES.values():
        fc = [c for c in cases if c['face'] == face[0]]
        if not fc:
            continue
        for field, basis in (('preds', 'pred_w'), ('preds_meas', 'meas_w')):
            for m in ('HALF', 'FIXED', 'NONE'):
                bad = sum(1 for c in fc if c[field][m] != c['count'])
                print(f'   {face[0]:13s} {basis:7s} {m:5s} '
                      f'{"SUPPORTED" if bad == 0 else "FALSIFIED"} ({len(fc) - bad}/{len(fc)} exact)')
    if worst > 0.12:
        print('   border-weight transition shift > 0.12 pt: border-proportional model falsified')
    print('   note: the same interval is reported on `none`; if it is non-zero there, the constant is '
          'not border presence')


if __name__ == '__main__':
    main()
