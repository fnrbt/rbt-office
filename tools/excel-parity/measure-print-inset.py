#!/usr/bin/env python3
"""Measure the plain-number inset fixture (make-print-inset-fixture.py) from PDF operators.

    python3 tools/excel-parity/measure-print-inset.py CELLS.tsv PDF_DIR SHEET_TSV OUT.tsv

SHEET_TSV maps sheets to PDFs (`sheet index pdf file` rows, as excel-print-pages.sh writes; the
browser/native proofs write the same shape). Cell edges are the per-column solid fill rectangles
(column B = first fill column, C = second; pages of a vertically paginated sheet share x). Text
origin x0 and advance end x1 come from Tm/TJ positions and the embedded /Widths, over every run in
the cell's row band (label in column A, nearest baseline).

Columns: sheet index row label scale align face size kind width cell_left cell_right cell_width
  text font fs x0 x1 advance left_inset right_inset
Exits 1 when the measured key set (sheet,row,width) differs from the manifest.
"""
import collections
import importlib.util
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location('numfill', os.path.join(HERE, 'measure-print-numfill.py'))
nf = importlib.util.module_from_spec(spec)
spec.loader.exec_module(nf)


def columns(page):
    """x ranges of the fill columns, left to right."""
    spans = collections.defaultdict(list)
    for x, y, w, h in page['rects']:
        if w > 5 and h > 3:
            spans[(round(x, 1), round(x + w, 1))].append((x, x + w))
    cols = []
    for k in sorted(spans):
        xs = spans[k]
        l = sum(a for a, _ in xs) / len(xs)
        r = sum(b for _, b in xs) / len(xs)
        if cols and abs(cols[-1][0] - l) < 0.2:
            continue
        cols.append((l, r))
    return cols


def f3(v):
    return '' if v is None else (f'{v:.3f}' if isinstance(v, float) else str(v))


def main():
    cells_path, pdf_dir, sheet_tsv, out_path = sys.argv[1:5]
    manifest = {}
    widths_of = {}
    for line in open(cells_path, encoding='utf-8'):
        if line.startswith('sheet\t') or not line.strip():
            continue
        p = line.rstrip('\n').split('\t')
        sheet, index, row, label, scale, align, face, size, kind, value, widths = p
        manifest[(sheet, label)] = (int(index), int(row), int(scale), align, face, size, kind)
        widths_of[sheet] = widths.split(',')
    expected = {(s, str(v[1]), w) for (s, _), v in manifest.items() for w in widths_of[s]}
    written = collections.Counter()
    out = open(out_path, 'w', encoding='utf-8')
    out.write('sheet\tindex\trow\tlabel\tscale\talign\tface\tsize\tkind\twidth\tcell_left\tcell_right\t'
              'cell_width\ttext\tfont\tfs\tx0\tx1\tadvance\tleft_inset\tright_inset\n')
    for sheet, index, name in nf.read_pdf_table(sheet_tsv):
        if sheet not in widths_of:
            continue
        for page in nf.parse_pdf(os.path.join(pdf_dir, name)):
            cols = columns(page)
            if len(cols) < len(widths_of[sheet]):
                continue
            cols = cols[-len(widths_of[sheet]):]
            labels = [r for r in page['runs'] if re.fullmatch(r'r\d{3}', r['text'].strip())
                      and r['x'] < cols[0][0]]
            ys = sorted(l['y'] for l in labels)
            for lab in labels:
                key = (sheet, lab['text'].strip())
                if key not in manifest:
                    continue
                idx, row, scale, align, face, size, kind = manifest[key]
                for (cl, cr), w in zip(cols, widths_of[sheet]):
                    rs = [r for r in page['runs'] if cl - 1.0 <= r['x'] <= cr + 0.5 and r['text'].strip()
                          and abs(r['y'] - lab['y']) < 5.0
                          and min(ys, key=lambda y: abs(y - r['y'])) == lab['y']]
                    if not rs:
                        continue
                    rs.sort(key=lambda r: r['x'])
                    text = ''.join(r['text'] for r in rs)
                    x0 = rs[0]['x']
                    x1 = max(r['end'] for r in rs)
                    written[(sheet, str(row), w)] += 1
                    out.write('\t'.join(map(f3, [
                        sheet, idx, row, lab['text'].strip(), scale, align, face, size, kind, w,
                        cl, cr, cr - cl, text, rs[0]['font'], rs[0]['size'], x0, x1, x1 - x0,
                        x0 - cl, cr - x1])) + '\n')
    out.close()
    keys = set(written)
    missing = sorted(expected - keys)
    extra = sorted(keys - expected)
    dup = sorted(k for k, n in written.items() if n > 1)
    ok = not missing and not extra and not dup
    print(f'{out_path}: keys expected {len(expected)} measured {len(keys)} missing {len(missing)} '
          f'extra {len(extra)} duplicate {len(dup)} {"PASS" if ok else "FAIL"}')
    for kind, items in (('missing', missing), ('extra', extra), ('duplicate', dup)):
        for k in items[:10]:
            print(f'  {kind}: {k}')
    sys.exit(0 if ok else 1)


if __name__ == '__main__':
    main()
