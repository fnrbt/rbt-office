#!/usr/bin/env python3
"""Compares the locked holdout decisions and advances with Excel's PDF.

The locks are written before the export:
  ENGINE_LOCK  the engine agent's lock (baseline + hinted candidate at the same ppem);
  EXACT_LOCK   this experiment's exact-height candidate lock.
The script never changes a prediction; it only joins them to
measure-print-advance-oracle.py's results.

    python3 tools/excel-parity/compare-print-advance-holdout.py \
        ENGINE_LOCK.tsv EXACT_LOCK.tsv RESULTS.tsv OUT.tsv
"""
import collections
import csv
import sys


def load(path):
    with open(path, encoding='utf-8') as f:
        return list(csv.DictReader((l for l in f if not l.startswith('#')), delimiter='\t'))


def main():
    if len(sys.argv) != 5:
        raise SystemExit(__doc__)
    engine_lock, exact_lock, results_path, out_path = sys.argv[1:]
    E = {(d['sheet'], d['column'], d['row']): d for d in load(engine_lock)}
    X = {(d['sheet'], d['column'], d['row']): d for d in load(exact_lock)}
    R = {(d['sheet'], d['column'], d['row']): d for d in load(results_path)}
    out = ['section\tsheet\ttext\tcells\texcel_full\tengine_full\tengine_cand_full\texact_full\t'
           'engine_wrong\tengine_cand_wrong\texact_wrong\tengine_misses\tengine_cand_misses\texact_misses']
    clusters = collections.defaultdict(lambda: dict(cells=0, excel=0, ef=0, cf=0, xf=0,
                                                    ew=0, cw=0, xw=0, em=[], cm=[], xm=[]))
    wide = collections.defaultdict(lambda: [0, 0.0, 0.0, 0.0, 0.0])
    for key, e in E.items():
        if key[2] != '1' or key not in R or key not in X:
            continue
        r, x = R[key], X[key]
        excel_full = int(r['nHashes']) == 0
        ef = e['base_full'] == 'yes'
        cf = e['cand_full'] == 'yes'
        xf = x['exact_full'] == 'yes'
        if e['kind'] == 'wide':
            if r['nHashes'] == '0':
                w = wide[(e['sheet'], e['text'])]
                w[0] += 1
                w[1] += float(r['pdf_adv_pt'])
                w[2] += float(e['base_run_pt'])
                w[3] += float(e['cand_run_pt'])
                w[4] += float(x['exact_text_pt'])
            continue
        c = clusters[(e['sheet'], e['text'])]
        c['cells'] += 1
        c['excel'] += excel_full
        c['ef'] += ef
        c['cf'] += cf
        c['xf'] += xf
        for flag, name, miss in ((ef, 'ew', 'em'), (cf, 'cw', 'cm'), (xf, 'xw', 'xm')):
            if flag != excel_full:
                c[name] += 1
                c[miss].append(r['column'])
    te = tc = tx = 0
    for (sheet, text), c in sorted(clusters.items()):
        out.append(f"cutoff\t{sheet}\t{text}\t{c['cells']}\t{c['excel']}\t{c['ef']}\t{c['cf']}\t{c['xf']}\t"
                   f"{c['ew']}\t{c['cw']}\t{c['xw']}\t{','.join(c['em'])}\t{','.join(c['cm'])}\t{','.join(c['xm'])}")
        te += c['ew']
        tc += c['cw']
        tx += c['xw']
    out += ['', 'section\tsheet\ttext\tcells\tmeasured\tengine\texact\t'
            'engine-minus-measured\texact-minus-measured']
    for (sheet, text), w in sorted(wide.items()):
        n = w[0] or 1
        out.append(f'advance\t{sheet}\t{text}\t{w[0]}\t{w[1] / n:.5f}\t{w[2] / n:.5f}\t'
                   f'{w[3] / n:.5f}\t{w[4] / n:.5f}\t{(w[2] - w[1]) / n:+.5f}\t'
                   f'{(w[4] - w[1]) / n:+.5f}')
    out += ['', 'section\tcells\tengine_wrong\tengine_cand_wrong\texact_wrong',
            f'dense\t{sum(c["cells"] for c in clusters.values())}\t{te}\t{tc}\t{tx}']
    with open(out_path, 'w', encoding='utf-8', newline='\n') as f:
        f.write('\n'.join(out) + '\n')
    print('\n'.join(out[-8:]))


if __name__ == '__main__':
    main()
