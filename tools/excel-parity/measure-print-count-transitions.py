#!/usr/bin/env python3
"""Measures printed hash counts from a direct Excel PDF export of
make-print-count-fixture.py's workbook.

Reuses measure-print-hash-geometry.py's PDF readers (fill rectangle -> physical
cell edges, Poppler word box -> printed run) rather than duplicating them. Emits
one row per printed word with the metadata needed to fit the count rule.

    python3 tools/excel-parity/measure-print-count-transitions.py META.tsv EXCELDIR OUT.tsv
"""
import collections
import importlib.util
import os
import sys

spec = importlib.util.spec_from_file_location(
    'hashmeasure', os.path.join(os.path.dirname(__file__), 'measure-print-hash-geometry.py'))
hashmeasure = importlib.util.module_from_spec(spec)
spec.loader.exec_module(hashmeasure)


def main():
    meta_path, excel_dir, out_path = sys.argv[1:4]
    meta = [l.rstrip('\n').split('\t') for l in open(meta_path, encoding='utf-8')]
    head, meta = meta[0], meta[1:]
    tsv = os.path.join(excel_dir, sorted(f for f in os.listdir(excel_dir) if f.endswith('.tsv'))[0])
    pdf_of = {}
    for line in open(tsv, encoding='utf-8'):
        p = line.rstrip('\n').split('\t')
        if len(p) >= 4 and p[2] == 'pdf' and p[3].endswith('.pdf'):
            pdf_of[p[0]] = p[3]

    by_sheet = collections.defaultdict(list)
    for m in meta:
        d = dict(zip(head, m))
        by_sheet[d['sheet']].append(d)

    rows = ['sheet\tcolumn\twidth_chars\tface\tsize_pt\titalic\tscale\tformat\tvalue\trow\ttext\tnHashes\t'
            'xMin\txMax\trunWidth\tcellLeft\tcellRight\tcellWidth\txMinLeft\txMaxLeft\trightGap\t'
            'diagnostics\tpage\tpdf_fonts']
    for sheet, items in by_sheet.items():
        pdf = os.path.join(excel_dir, pdf_of[sheet])
        pages = hashmeasure.pdf_pages(pdf)
        for page in range(1, pages + 1):
            fills = hashmeasure.svg_fills(pdf, page)
            words = hashmeasure.page_words(pdf, page)
            fonts = hashmeasure.pdf_fonts(pdf, page)
            for d in items:
                col = int(d['column'])
                rect = fills.get(hashmeasure.fill_color(col - 2))
                if rect is None:
                    continue
                left, right = rect
                inside = sorted([w for w in words if left - 2 <= (w[0] + w[2]) / 2 <= right + 2],
                                key=lambda w: w[1])
                # Group the cell's words into printed lines (row 1, row 2); a padded accounting
                # value can arrive as several words ("$" and "1,234,567").
                clusters = []
                for w in inside:
                    if clusters and abs(w[1] - clusters[-1][0][1]) <= 2.0:
                        clusters[-1].append(w)
                    else:
                        clusters.append([w])
                diags = []
                if not clusters:
                    diags.append('missing-word')
                elif len(clusters) > 2:
                    diags.append(f'extra-line:{len(clusters)}')
                for ri, group in enumerate(clusters[:2]):
                    group = sorted(group, key=lambda w: w[0])
                    x0 = min(w[0] for w in group)
                    x1 = max(w[2] for w in group)
                    text = ' '.join(w[4] for w in group)
                    n = len(text) if text and set(text) == {'#'} else 0
                    row = '1' if ri == 0 else '2'
                    if row == '1' and n == 0:
                        diags.append('full-numeric' if text and not set(text) <= {'#'} else 'no-word')
                    rows.append('\t'.join([
                        sheet, d['column'], d['width_chars'], d['face'], d['size_pt'], d['italic'],
                        d['scale'], d.get('format', ''), d.get('value', ''), row, text, str(n),
                        f'{x0:.3f}', f'{x1:.3f}', f'{x1 - x0:.3f}', f'{left:.3f}', f'{right:.3f}',
                        f'{right - left:.3f}', f'{x0 - left:.3f}', f'{x1 - left:.3f}',
                        f'{right - x1:.3f}', ';'.join(sorted(set(diags))) or 'ok', str(page), fonts]))
    with open(out_path, 'w', encoding='utf-8', newline='\n') as f:
        f.write('\n'.join(rows) + '\n')
    print(f'wrote {len(rows) - 1} measurements to {out_path}')


if __name__ == '__main__':
    main()
