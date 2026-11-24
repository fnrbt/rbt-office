#!/usr/bin/env python3
"""Check the printed fill-count law on Excel's measured fill cells (print-inset-oracle.md, "Fill count").

    python3 tools/excel-parity/check-print-fill-count.py [NUMFILL_ORACLE_DIR]   (default /var/tmp/numfill-oracle)

count = max(0, floor((cw - 2P + 2 - w) / a)) in 600-dpi dots: cw the true column width (fill rectangle less
the 1 dot Excel overpaints), P = ceil(widest hinted digit / 4) + 4 of the cell font, w the hinted width of the
number run, a the hinted advance of the fill character (print-fill-count-gdi.tsv). Exits 1 on any miss.
"""
import collections
import csv
import importlib.util
import math
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location('numfill', os.path.join(HERE, 'measure-print-numfill.py'))
nf = importlib.util.module_from_spec(spec)
spec.loader.exec_module(nf)
D = 600.0 / 72.0
FILL = {'space': ' ', 'dot': '.', 'dash': '-'}
adv = collections.defaultdict(dict)
for r in csv.DictReader(open(os.path.join(HERE, 'print-fill-count-gdi.tsv'), encoding='utf-8'), delimiter='\t'):
    adv[r['face']][r['char']] = int(r['adv'])


def pad(face):
    return math.ceil(max(adv[face][c] for c in '0123456789') / 4) + 4


def predict(face, cw, amount, fill):
    w = sum(adv[face][c] for c in amount)
    return max(0, math.floor((cw + 2 - 2 * pad(face) - w) / adv[face][FILL.get(fill, fill)]))


def fill_widths(directory, stem):
    """True column widths per sheet, left to right, from the tall fill rectangles."""
    out = {}
    for sheet, _, name in nf.read_pdf_table(os.path.join(directory, stem + '.tsv')):
        widths = []
        for page in nf.parse_pdf(os.path.join(directory, name)):
            spans = sorted({(round(x * D, 2), round((x + w) * D, 2)) for x, y, w, h in page['rects'] if w > 5 and h > 30})
            widths += [round(b - a) - 1 for a, b in spans]
        out[sheet] = widths
    return out


def main():
    root = sys.argv[1] if len(sys.argv) > 1 else '/var/tmp/numfill-oracle'
    total_bad = 0
    for key, face in (('aptos', 'Aptos Narrow'), ('arial', 'Arial'), ('calibri', 'Calibri')):
        stem = f'synthetic-print-border-fill-{key}'
        widths = fill_widths(os.path.join(root, f'border-{key}'), stem)
        rows = list(csv.DictReader(open(os.path.join(root, f'border-{key}-measured.tsv'), encoding='utf-8'), delimiter='\t'))
        labels = collections.defaultdict(set)
        for r in rows:
            labels[r['sheet']].add(float(r['width']))
        n = bad = 0
        for r in rows:
            if not r['fill_char'] or not r['amount']:
                continue
            cw = widths[r['sheet']][sorted(labels[r['sheet']]).index(float(r['width']))]
            n += 1
            bad += predict(face, cw, r['amount'], r['fill_char']) != int(r['count'])
        print(f'border-{key}: {n - bad}/{n} exact')
        total_bad += bad
    for set_name, stem, face in (('trail-train', 'synthetic-print-numfill-trail', 'Aptos Narrow'),
                                 ('trail-holdout', 'synthetic-print-numfill-trail-holdout', 'Arial')):
        widths = fill_widths(os.path.join(root, set_name), stem)
        rows = [r for r in csv.DictReader(open(os.path.join(root, f'{set_name}-comparison.tsv'), encoding='utf-8'),
                                          delimiter='\t') if r['side'] == 'excel']
        labels = collections.defaultdict(set)
        for r in rows:
            labels[r['sheet']].add(float(r['width']))
        n = bad = 0
        for r in rows:
            if not r['fill_char'] or not r['amount']:
                continue
            cw = widths[r['sheet']][sorted(labels[r['sheet']]).index(float(r['width']))]
            count = int(r['fill_before'] or 0) + int(r['fill_after'] or 0)
            n += 1
            bad += predict(face, cw, r['amount'], r['fill_char']) != count
        print(f'{set_name}: {n - bad}/{n} exact')
        total_bad += bad
    print('PASS' if total_bad == 0 else f'FAIL ({total_bad} misses)')
    sys.exit(0 if total_bad == 0 else 1)


if __name__ == '__main__':
    main()
