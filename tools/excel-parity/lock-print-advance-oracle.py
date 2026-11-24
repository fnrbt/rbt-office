#!/usr/bin/env python3
"""Locks the printed-advance oracle's holdout predictions before its export.

The holdout workbook reserves an unseen scale (Verdana 13 upright at 140 %) and
an unseen face (Tahoma 15 upright at 90 %).  For every row-1 cell this writes:

  * the engine's current rounded-height advance and full/hash decision;
  * the exact-height candidate's advance and decision;
  * the pre-registered W(width_chars) and the measured-W decision columns.

It also records the SHA-256 of the workbook and metadata.  No PDF may be
produced from the holdout workbook before this file exists.

    python3 tools/excel-parity/lock-print-advance-oracle.py META.tsv WORKBOOK.xlsx OUT.tsv
"""
import csv
import hashlib
import importlib.util
import os
import sys

_here = os.path.dirname(os.path.abspath(__file__))
_spec = importlib.util.spec_from_file_location('gen', os.path.join(_here, 'make-print-advance-oracle.py'))
gen = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(gen)


def main():
    if len(sys.argv) != 4:
        raise SystemExit(__doc__)
    meta_path, workbook, out_path = sys.argv[1:]
    with open(meta_path, encoding='utf-8') as f:
        meta = list(csv.DictReader(f, delimiter='\t'))
    header = ['sheet', 'column', 'row', 'width_chars', 'kind', 'role', 'face', 'size_pt', 'scale',
              'format', 'text', 'W_model_pt', 'engine_text_pt', 'exact_text_pt', 'c_pt',
              'engine_full', 'exact_full', 'engine_margin_pt', 'exact_margin_pt']
    out = ['# locked from the shipping rounded-height model and the exact-height candidate; '
           'no holdout PDF existed when this file was written',
           '# workbook_sha256=' + hashlib.sha256(open(workbook, 'rb').read()).hexdigest(),
           '# meta_sha256=' + hashlib.sha256(open(meta_path, 'rb').read()).hexdigest()] + ['\t'.join(header)]
    counts = {'engine_full': 0, 'exact_full': 0, 'n': 0}
    for d in meta:
        if d['row'] != '1':
            continue
        em = gen.glyph_ems(os.path.join(gen.ROOT, gen.FONT_FILES[(d['face'], d['italic'] == '1')]))
        size, scale = float(d['size_pt']), int(d['scale']) / 100.0
        c = em['0'] * size * scale / 2.0 + 0.875
        exact = gen.model_text(d['text'], em, size, scale, True)
        engine = gen.model_text(d['text'], em, size, scale, False)
        W = gen.physical_width(float(d['width_chars']), scale)
        ef = engine <= W - c + 1e-9
        xf = exact <= W - c + 1e-9
        counts['n'] += 1
        counts['engine_full'] += ef
        counts['exact_full'] += xf
        out.append('\t'.join([d['sheet'], d['column'], d['row'], d['width_chars'], d['kind'], d['role'],
                              d['face'], d['size_pt'], d['scale'], d['format'], d['text'],
                              f'{W:.4f}', f'{engine:.4f}', f'{exact:.4f}', f'{c:.4f}',
                              'yes' if ef else 'no', 'yes' if xf else 'no',
                              f'{W - c - engine:+.4f}', f'{W - c - exact:+.4f}']))
    with open(out_path, 'w', encoding='utf-8', newline='\n') as f:
        f.write('\n'.join(out) + '\n')
    print(f'locked {counts["n"]} row-1 cells: engine full {counts["engine_full"]}, '
          f'exact full {counts["exact_full"]} -> {out_path}')


if __name__ == '__main__':
    main()
