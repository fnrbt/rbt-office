#!/usr/bin/env python3
"""Summarises measure-print-count-transitions.py rows: per case pitch, physical-width slope,
count-transition intercept intervals, transition-spacing drift, and the hash->full-number fade-in.

    python3 tools/excel-parity/analyze-print-count-transitions.py META RESULTS OUT.tsv

The pitch used to turn a boundary into an intercept is the design `#` advance quantised to the
printed 600-dpi dot grid at the effective font size (size x print scale), which agrees with the
PDF word-box run advance to +-0.005 pt on the training faces. For a face whose local font file is
not the one Excel printed, the measured run advance (`a_run`) is used instead and marked in `note`.
"""
import collections
import importlib.util
import os
import statistics
import sys

HERE = os.path.dirname(__file__)
spec = importlib.util.spec_from_file_location('countpredict', os.path.join(HERE, 'predict-print-count.py'))
countpredict = importlib.util.module_from_spec(spec)
spec.loader.exec_module(countpredict)


def main():
    meta_path, results_path, out_path = sys.argv[1:4]
    meta = [l.rstrip('\n').split('\t') for l in open(meta_path, encoding='utf-8')]
    mhead, meta = meta[0], meta[1:]
    meta = {(d['sheet'], d['column']): d for d in (dict(zip(mhead, r)) for r in meta)}
    rows = [l.rstrip('\n').split('\t') for l in open(results_path, encoding='utf-8')]
    head, rows = rows[0], rows[1:]
    by_sheet = collections.defaultdict(list)
    for r in (dict(zip(head, x)) for x in rows):
        if r['row'] == '1' and r['nHashes'] != '':
            by_sheet[r['sheet']].append(r)

    cache = {}
    out = ['sheet\tface\tsize_pt\titalic\tscale\tcolumns\tn_min\tn_max\tslots\t'
           'a_design_pt\ta_run_pt\ta_minus_run\t'
           'c_lo\tc_hi\tc_mid\tc_formula\tc_formula_minus_mid\t'
           'transition_residual_min\ttransition_residual_max\tboundary_slope_pt\tboundary_drift_pt\t'
           'full_last_hash_W\tfull_first_W\tfull_text_run\tc_text_lo\tc_text_hi\t'
           'W_slope\tW_intercept\tnote']
    for sheet, rr in sorted(by_sheet.items()):
        rr.sort(key=lambda r: float(r['width_chars']))
        face, italic = rr[0]['face'], rr[0]['italic'] == '1'
        size, scale = float(rr[0]['size_pt']), int(rr[0]['scale'])
        key = (face, italic)
        note = []
        a_run = statistics.mean(float(r['runWidth']) / int(r['nHashes']) for r in rr if int(r['nHashes']) > 0)
        a_design = None
        if key in countpredict.FONT_FILES and not (face == 'Georgia' and scale != 100):
            try:
                g, upm, _ = countpredict.font_metrics(f'{countpredict.FONT_DIR}/{countpredict.FONT_FILES[key]}')
                dots = round(size * scale / 100.0 * 600.0 / 72.0)
                a_design = round(g['#'] * dots) * 72.0 / 600.0
            except OSError:
                note.append('font-file-missing')
        if a_design is None:
            a_design = a_run
            note.append('a_design=a_run')
        a = a_design

        seq = [(float(r['width_chars']), float(r['cellWidth']), int(r['nHashes']), float(r['runWidth'])) for r in rr]
        nz = [s for s in seq if s[2] > 0]
        intervals = []
        bounds = []
        for u, v in zip(seq, seq[1:]):
            n, m = u[2], v[2]
            if n > 0 and m == n + 1:
                Wl, Wf = u[1], v[1]
                clo = Wl - (n + 1) * a
                chi = min(Wl - n * a, Wf - (n + 1) * a)
                intervals.append((n, clo, chi))
                bounds.append((n, (Wl + Wf) / 2.0))
        c_lo = max((i[1] for i in intervals), default=float('nan'))
        c_hi = min((i[2] for i in intervals), default=float('nan'))
        c_mid = (c_lo + c_hi) / 2.0 if intervals else float('nan')
        c_formula = float('nan')
        if key in countpredict.FONT_FILES:
            try:
                g, upm, win = countpredict.font_metrics(f'{countpredict.FONT_DIR}/{countpredict.FONT_FILES[key]}')
                c_formula = g['0'] * size * scale / 100.0 / 2.0 + 0.875
                if italic:
                    c_formula += countpredict.italic_shift(win, upm, size, scale / 100.0)
            except OSError:
                pass
        fmt = next((m.get('format', '') for k, m in meta.items() if k[0] == sheet), '')
        slots = 1 if fmt.startswith('_("$"') else 0
        if c_formula == c_formula:
            c_formula += slots * a
        formula_delta = c_formula - c_mid if intervals else float('nan')
        residuals = [w - (c_mid + (n + 1) * a) for n, w in bounds]
        slope = float('nan')
        if len(bounds) >= 2:
            mx = statistics.mean(b[0] for b in bounds)
            my = statistics.mean(b[1] for b in bounds)
            den = sum((b[0] - mx) ** 2 for b in bounds)
            slope = sum((b[0] - mx) * (b[1] - my) for b in bounds) / den if den else float('nan')
        # Hash -> full-number fade-in.
        full = [i for i, s in enumerate(seq) if s[2] == 0]
        full_last = full_first = text_run = ct_lo = ct_hi = float('nan')
        if full and nz:
            i = full[0]
            if i > 0 and seq[i - 1][2] > 0:
                last, first = seq[i - 1], seq[i]
                full_last, full_first, text_run = last[1], first[1], first[3]
                ct_lo, ct_hi = last[1] - text_run, first[1] - text_run
        xs = [s[0] for s in seq]
        ys = [s[1] for s in seq]
        mx, my = statistics.mean(xs), statistics.mean(ys)
        wslope = sum((x - mx) * (y - my) for x, y in zip(xs, ys)) / sum((x - mx) ** 2 for x in xs)
        winter = my - wslope * mx
        out.append('\t'.join([
            sheet, face, rr[0]['size_pt'], rr[0]['italic'], rr[0]['scale'], str(len(seq)),
            str(min(s[2] for s in nz)), str(max(s[2] for s in nz)), str(slots),
            f'{a_design:.4f}', f'{a_run:.4f}', f'{a_design - a_run:+.4f}',
            f'{c_lo:.4f}' if intervals else '', f'{c_hi:.4f}' if intervals else '',
            f'{c_mid:.4f}' if intervals else '', f'{c_formula:.4f}' if c_formula == c_formula else '',
            f'{formula_delta:+.4f}' if formula_delta == formula_delta else '',
            f'{min(residuals):+.4f}' if residuals else '', f'{max(residuals):+.4f}' if residuals else '',
            f'{slope:.4f}' if slope == slope else '', f'{slope - a:+.4f}' if slope == slope else '',
            f'{full_last:.4f}' if full_last == full_last else '', f'{full_first:.4f}' if full_first == full_first else '',
            f'{text_run:.4f}' if text_run == text_run else '',
            f'{ct_lo:.4f}' if ct_lo == ct_lo else '', f'{ct_hi:.4f}' if ct_hi == ct_hi else '',
            f'{wslope:.4f}', f'{winter:.4f}', ';'.join(note) or 'ok']))
    with open(out_path, 'w', encoding='utf-8', newline='\n') as f:
        f.write('\n'.join(out) + '\n')
    print(f'wrote {len(out) - 1} case rows to {out_path}')


if __name__ == '__main__':
    main()
