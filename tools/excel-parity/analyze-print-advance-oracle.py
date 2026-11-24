#!/usr/bin/env python3
"""Analyses the printed-advance oracle: slope/intercept, W, and cutoffs.

Reads `measure-print-advance-oracle.py`'s results for the training workbook and
reports, in this order:

  agreement      PDF positioning advance vs Poppler bbox width (should be ~0);
  slope          advance = a * (repeated digit count) + b regressions for
                 Verdana 13 at 75 % and 100 %, against the engine's rounded-height
                 model and the exact-height candidate;
  column width   measured fill width vs the engine's W(width_chars) model, by
                 sheet (the Georgia column-width residual);
  cutoffs        for every tight sweep, the last hashed and first full physical
                 widths, the implied `c = W - advance` bracket, and candidate /
                 baseline decision misses;
  punctuation    literal strings' measured advances against both models.

    python3 tools/excel-parity/analyze-print-advance-oracle.py META.tsv RESULTS.tsv OUT_PREFIX
"""
import collections
import csv
import importlib.util
import math
import os
import sys

_here = os.path.dirname(os.path.abspath(__file__))
_spec = importlib.util.spec_from_file_location('gen', os.path.join(_here, 'make-print-advance-oracle.py'))
gen = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(gen)


def load(path):
    with open(path, encoding='utf-8') as f:
        return list(csv.DictReader(f, delimiter='\t'))


def models(face, size, scale, text):
    em = gen.glyph_ems(os.path.join(gen.ROOT, gen.FONT_FILES[(face, False)]))
    c = em['0'] * size * scale / 2.0 + 0.875
    return (gen.model_text(text, em, size, scale, True),
            gen.model_text(text, em, size, scale, False), c)


def fmt(x, n=4):
    return f'{x:.{n}f}'


def main():
    if len(sys.argv) != 4:
        raise SystemExit(__doc__)
    meta_path, results_path, prefix = sys.argv[1:]
    meta = load(meta_path)
    rows = load(results_path)
    row1 = [d for d in rows if d['row'] == '1']
    out = []

    # ---- 1. PDF pen advance vs extracted word bbox ----
    diffs = [abs(float(d['adv_minus_bbox'])) for d in rows if d['adv_minus_bbox']]
    out.append('## advance agreement')
    out.append(f'rows={len(rows)}  max|pdfAdv-bbox|={max(diffs):.6f} pt  '
               f'over0.001={sum(1 for x in diffs if x > 0.001)}')

    # ---- 2. slope/intercept of repeated same digits ----
    out.append('')
    out.append('## advance = slope * digits + intercept (repeated 1s, format 0.00%)')
    out.append('sheet\tdigit_count\tmeasured_adv\tengine_adv\texactheight_adv\t'
               'measured_slope\tmeasured_intercept\tengine_slope\tengine_intercept\t'
               'exact_slope\texact_intercept\tmax_residual')
    table = []
    for sheet in ('V13up_pct75', 'V13up_pct100', 'H_V13up_pct140', 'H_Tahoma15up_pct90'):
        pts = []
        for d in row1:
            t = d['text_run']
            if d['sheet'] == sheet and d['kind'] == 'wide' and d['role'] == 'num' and t.endswith('.00%') and t[:-4] == '1' * len(t[:-4]):
                n = len(t[:-4])
                ex, eng, _c = models(d['face'], float(d['size_pt']), int(d['scale']) / 100.0, t)
                pts.append((n, float(d['pdf_adv_pt']), eng, ex))
        if len(pts) < 2:
            continue
        pts.sort()

        def fit(idx):
            xs = [p[0] for p in pts]
            ys = [p[idx] for p in pts]
            mx = sum(xs) / len(xs)
            my = sum(ys) / len(ys)
            a = sum((x - mx) * (y - my) for x, y in zip(xs, ys)) / sum((x - mx) ** 2 for x in xs)
            b = my - a * mx
            resid = max(abs(y - (a * x + b)) for x, y in zip(xs, ys))
            return a, b, resid
        am, bm, rm = fit(1)
        ae, be, re = fit(2)
        ax, bx, rx = fit(3)
        table.append((sheet, pts))
        out.append(f'{sheet}\t{",".join(str(p[0]) for p in pts)}\t'
                   f'a={am:.5f} b={bm:.5f}\tengine a={ae:.5f} b={be:.5f}\texact a={ax:.5f} b={bx:.5f}\t'
                   f'measured_resid={rm:.4f} engine_resid={re:.4f} exact_resid={rx:.4f}')
        for n, adv, eng, ex in pts:
            out.append(f'{sheet}\t{n}\t{adv:.5f}\t{eng:.5f}\t{ex:.5f}\t\t\t\t\t\t\t'
                       f'{eng - adv:+.4f}/{ex - adv:+.4f}')
    # per-digit differences at the two Verdana scales
    out.append('')
    out.append('## implied per-digit advance by PDF operator strategy')
    out.append('sheet\ttc\tdigit_advance\tstrings')
    for sheet in ('V13up_pct75', 'V13up_pct100'):
        groups = collections.defaultdict(list)
        for d in row1:
            t = d['text_run']
            if d['sheet'] == sheet and d['kind'] == 'wide' and d['role'] == 'num' and t.endswith('.00%') and t[:-4] == '1' * len(t[:-4]):
                groups[d['tc']].append((len(t[:-4]), float(d['pdf_adv_pt']), t))
        for tc, pts in sorted(groups.items()):
            pts.sort()
            slopes = [(pts[i + 1][1] - pts[i][1]) / (pts[i + 1][0] - pts[i][0]) for i in range(len(pts) - 1)]
            out.append(f'{sheet}\t{tc}\t{sum(slopes) / len(slopes):.5f}\t' +
                       ', '.join(f'n{n}' for n, _, _ in pts))

    # ---- 3. measured column width vs the engine's W(width) ----
    out.append('')
    out.append('## physical column width minus W(width_chars) model')
    out.append('sheet\tkind\tn\tmean\tmin\tmax\tp95')
    for sheet in sorted({d['sheet'] for d in rows}):
        for kind in ('wide', 'dense'):
            rs = []
            for d in rows:
                if d['sheet'] != sheet or d['row'] != '1' or d['kind'] != kind:
                    continue
                w = gen.physical_width(float(d['width_chars']), int(d['scale']) / 100.0)
                rs.append(float(d['cellWidth']) - w)
            if not rs:
                continue
            rs.sort()
            out.append(f'{sheet}\t{kind}\t{len(rs)}\t{sum(rs) / len(rs):+.4f}\t{rs[0]:+.4f}\t{rs[-1]:+.4f}\t'
                       f'{rs[int(0.95 * (len(rs) - 1))]:+.4f}')

    # ---- 4. cutoff sweeps ----
    out.append('')
    out.append('## tight-cell cutoffs (W from the PDF fill rectangle)')
    out.append('sheet\ttext\tcols\tlast_hash_W\tfirst_full_W\tc_bracket(measuredW,measuredAdv)\t'
               'c_bracket(modelW,engineAdv)\tmodel_c\tinside_measured\tinside_modelW\t'
               'engine_wrong\texact_wrong\tengine_misses\texact_misses')
    cluster = collections.defaultdict(list)
    for d in row1:
        if d['kind'] == 'dense':
            cluster[(d['sheet'], d['text_expected'])].append(d)
    for (sheet, text), ds in sorted(cluster.items()):
        ds.sort(key=lambda d: float(d['width_chars']))
        face = ds[0]['face']
        size = float(ds[0]['size_pt'])
        scale = int(ds[0]['scale']) / 100.0
        exact, engine, c = models(face, size, scale, text)
        # measured advance from the identical wide string (bbox/pos advance)
        wide = [d for d in row1 if d['sheet'] == sheet and d['kind'] == 'wide'
                and d['text_run'] == text and d['nHashes'] == '0']
        adv = float(wide[0]['pdf_adv_pt']) if wide else None
        full = [d for d in ds if int(d['nHashes']) == 0]
        hashed = [d for d in ds if int(d['nHashes']) > 0]
        if not full or not hashed:
            out.append(f'{sheet}\t{text}\t{len(ds)}\t-[0m\t-\t-\t{c:.4f}\t-\t-\t-\t-\t-')
            continue
        first_full = min(full, key=lambda d: float(d['width_chars']))
        last_hash = max(hashed, key=lambda d: float(d['width_chars']))
        wf, wh = float(first_full['cellWidth']), float(last_hash['cellWidth'])
        lo, hi = wh - adv, wf - adv
        wfm = gen.physical_width(float(first_full['width_chars']), scale)
        whm = gen.physical_width(float(last_hash['width_chars']), scale)
        lo_m, hi_m = whm - engine, wfm - engine
        eng_wrong = ex_wrong = 0
        eng_miss, ex_miss = [], []
        for d in ds:
            Wm = gen.physical_width(float(d['width_chars']), scale)
            excel_full = int(d['nHashes']) == 0
            pe = engine <= Wm - c + 1e-9
            px = exact <= Wm - c + 1e-9
            if pe != excel_full:
                eng_wrong += 1
                eng_miss.append(d['column'])
            if px != excel_full:
                ex_wrong += 1
                ex_miss.append(d['column'])
        out.append(f'{sheet}\t{text}\t{len(ds)}\t{wh:.4f}\t{wf:.4f}\t({lo:.4f}, {hi:.4f}]\t'
                   f'({lo_m:.4f}, {hi_m:.4f}]\t{c:.4f}\t'
                   f'{"yes" if lo < c <= hi else "no"}\t{"yes" if lo_m < c <= hi_m else "no"}\t'
                   f'{eng_wrong}\t{ex_wrong}\t{",".join(eng_miss)}\t{",".join(ex_miss)}')

    # ---- 5. punctuation recipients ----
    out.append('')
    out.append('## literal punctuation advances (wide cells)')
    out.append('sheet\ttext\tmeasured\tengine\texact\tdelta_exact')
    for d in row1:
        if d['kind'] != 'wide' or d['role'] != 'lit':
            continue
        if d['sheet'] not in ('V13up_pct75', 'V13up_pct100'):
            continue
        t = d['text_run']
        if any(x in t for x in (',', '(', ')', '$', '%', '-')) and t not in ('1234567.89%',):
            ex, eng, _ = models(d['face'], float(d['size_pt']), int(d['scale']) / 100.0, t)
            out.append(f"{d['sheet']}\t{t}\t{float(d['pdf_adv_pt']):.5f}\t{eng:.5f}\t{ex:.5f}\t{ex - float(d['pdf_adv_pt']):+.4f}")

    with open(prefix + '.analysis.tsv', 'w', encoding='utf-8', newline='\n') as f:
        f.write('\n'.join(out) + '\n')
    print('\n'.join(out))


if __name__ == '__main__':
    main()
