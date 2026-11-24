#!/usr/bin/env python3
"""Measures the full<->hash anchor / holdout workbooks from a direct Excel PDF export.

Reuses measure-print-hash-geometry.py's readers (solid fill rectangle -> physical cell
edges; Poppler `pdftotext -bbox` -> printed word box; `pdffonts` -> embedded faces).
For each sheet column both rows are read: row 1 is the test cell, row 2 the `-98.41`
non-overflow control of known text.

Per printed line it emits the word-box span (`runWidth` = printed advance incl. side
bearings, blank `_x` pads excluded), both physical cell edges and the left/right gaps,
the printed hash count and the printed text.  A wide-cell row therefore gives the
printed advance of the full formatted string and the right gap against the physical
cell edge; a dense-column row gives the physical width at the full->hash cutoff.

    python3 tools/excel-parity/measure-print-fullhash-anchor.py META.tsv EXCELDIR OUT.tsv
"""
import collections
import importlib.util
import os
import sys

_here = os.path.dirname(os.path.abspath(__file__))
_spec = importlib.util.spec_from_file_location(
    'hashmeasure', os.path.join(_here, 'measure-print-hash-geometry.py'))
hashmeasure = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(hashmeasure)


def main():
    argv = sys.argv[1:]
    map_path = argv[argv.index('--map') + 1] if '--map' in argv else None
    argv = [a for i, a in enumerate(argv) if a != '--map' and (i == 0 or argv[i - 1] != '--map')]
    meta_path, excel_dir, out_path = argv[:3]
    meta = [l.rstrip('\n').split('\t') for l in open(meta_path, encoding='utf-8')]
    head, meta = meta[0], meta[1:]
    tsv = map_path or os.path.join(excel_dir, sorted(
        f for f in os.listdir(excel_dir)
        if f.endswith('.tsv') and not f.startswith(('results', 'analysis')))[0])
    pdf_of = {}
    for line in open(tsv, encoding='utf-8'):
        p = line.rstrip('\n').split('\t')
        if len(p) >= 4 and p[2] == 'pdf' and p[3].endswith('.pdf'):
            pdf_of[p[0]] = p[3]

    by_sheet = collections.defaultdict(list)
    for m in meta:
        d = dict(zip(head, m))
        by_sheet[d['sheet']].append(d)

    rows = ['sheet\tcolumn\trow\twidth_chars\tkind\tface\tsize_pt\titalic\tscale\tformat\tvalue\t'
            'text_expected\ttext_printed\tnHashes\txMin\txMax\trunWidth\tcellLeft\tcellRight\tcellWidth\t'
            'leftGap\trightGap\txMinRight\tnWords\tlastWordText\tlastWordWidth\t'
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
                # Only words that start and end at this cell: an adjacent accounting
                # column's leading "$" sits exactly on the shared edge.
                inside = sorted([w for w in words
                                 if w[0] >= left - 1.0 and w[2] <= right + 1.0],
                                key=lambda w: w[1])
                clusters = []
                for w in inside:
                    if clusters and abs(w[1] - clusters[-1][0][1]) <= 2.0:
                        clusters[-1].append(w)
                    else:
                        clusters.append([w])
                diags = []
                want_row = d['row']
                # Top line is row 1, next is row 2; the control is always the second.
                idx = 0 if want_row == '1' else 1
                if len(clusters) > 2:
                    diags.append(f'extra-line:{len(clusters)}')
                if len(clusters) <= idx:
                    diags.append('missing-word')
                    rows.append('\t'.join([
                        sheet, str(col), want_row, d['width_chars'], d['kind'], d['face'],
                        d['size_pt'], d['italic'], d['scale'], d['format'], d['value'],
                        d['text'], '', '0', '', '', '', f'{left:.3f}', f'{right:.3f}',
                        f'{right - left:.3f}', '', '', '', '', '', '', ';'.join(diags), str(page), fonts]))
                    continue
                group = sorted(clusters[idx], key=lambda w: w[0])
                x0 = min(w[0] for w in group)
                x1 = max(w[2] for w in group)
                text = ' '.join(w[4] for w in group)
                n = sum(w[4].count('#') for w in group)
                if n and text.strip('# ') != '':
                    diags.append('mixed-hash-word')
                if x0 < left - 0.5 or x1 > right + 0.5:
                    diags.append('outside-cell')
                rows.append('\t'.join([
                    sheet, str(col), want_row, d['width_chars'], d['kind'], d['face'],
                    d['size_pt'], d['italic'], d['scale'], d['format'], d['value'], d['text'],
                    text, str(n), f'{x0:.3f}', f'{x1:.3f}', f'{x1 - x0:.3f}',
                    f'{left:.3f}', f'{right:.3f}', f'{right - left:.3f}',
                    f'{x0 - left:.3f}', f'{right - x1:.3f}', f'{right - x0:.3f}',
                    str(len(group)), group[-1][4], f'{group[-1][2] - group[-1][0]:.3f}',
                    ';'.join(sorted(set(diags))) or 'ok', str(page), fonts]))
    with open(out_path, 'w', encoding='utf-8', newline='\n') as f:
        f.write('\n'.join(rows) + '\n')
    print(f'wrote {len(rows) - 1} measurements to {out_path}')


if __name__ == '__main__':
    main()
