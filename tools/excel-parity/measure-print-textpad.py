#!/usr/bin/env python3
"""Measure text/num word boxes against the adjacent solid-fill cell edges in an Excel/ours PDF pair.

Metric: pdftotext -bbox word box against the yellow (#FFF2CC, text) and pale-blue (#DDEBF7, numeric
control) cell fills from pdftocairo SVG. Both fills are physical cell rectangles, so xMin/gap comparisons
do not depend on page geometry.

    python3 tools/excel-parity/measure-print-textpad.py CELLS.tsv EXCEL_DIR EXCEL_TSV OURS_DIR OUT.tsv

EXCEL_DIR holds the fresh direct-export PDFs and EXCEL_TSV is excel-print-pages.sh's sheet/index/pdf table.
OURS_DIR holds ours-<index>.pdf from tools/spreadsheet-pdf-runner with
RB_WORD_FONT_DIR=artifacts/fonts/windows:artifacts/fonts/windows-extra.
"""
import re, subprocess, sys, os, tempfile, html, collections

FILLS = {'FFFFF2CC': (100.0, 94.898987, 79.998779), 'FFDDEBF7': (86.7, 92.2, 96.9)}


def svg_boxes(pdf, page=1):
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, 'p.svg')
        subprocess.run(['pdftocairo', '-f', str(page), '-l', str(page), '-svg', pdf, out], check=True,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        svg = out + '.svg' if os.path.exists(out + '.svg') else out
        text = open(svg, encoding='utf-8').read()
    boxes = {}
    for m in re.finditer(r'<path\s+([^>]*?)/>', text):
        attrs = m.group(1)
        fm = re.search(r'fill="rgb\(([^)]*)\)"', attrs)
        if not fm:
            continue
        try:
            rgb = tuple(float(x.strip().rstrip('%')) for x in fm.group(1).split(','))
        except ValueError:
            continue
        name = None
        for key, want in FILLS.items():
            if all(abs(a - b) < 0.6 for a, b in zip(rgb, want)):
                name = key
                break
        if name is None:
            continue
        d = re.search(r'd="([^"]*)"', attrs)
        if not d:
            continue
        nums = [float(v) for v in re.findall(r'-?\d+\.?\d*', d.group(1))]
        xs, ys = nums[0::2], nums[1::2]
        box = (min(xs), min(ys), max(xs), max(ys))
        if name in boxes:
            b = boxes[name]
            boxes[name] = (min(b[0], box[0]), min(b[1], box[1]), max(b[2], box[2]), max(b[3], box[3]))
        else:
            boxes[name] = box
    return boxes


def words(pdf, page=1):
    out = subprocess.run(['pdftotext', '-bbox', '-f', str(page), '-l', str(page), pdf, '-'],
                         capture_output=True, text=True).stdout
    return [(float(a), float(b), float(x), float(y2), html.unescape(t))
            for a, b, x, y2, t in re.findall(
                r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">([^<]*)</word>', out)]


def rows_of(ws):
    rows = []
    for w in sorted(ws, key=lambda w: w[1]):
        if rows and abs(w[3] - rows[-1][0]) < 4.0:
            rows[-1][1].append(w)
        else:
            rows.append([w[3], [w]])
    return [sorted(p, key=lambda w: w[0]) for _, p in rows]


def measure(pdf, page, nrows):
    boxes = svg_boxes(pdf, page)
    bt, bn = boxes['FFFFF2CC'], boxes['FFDDEBF7']
    rows = []
    for part in rows_of(words(pdf, page)):
        split = bn[0] - 10.0
        b = [w for w in part if w[0] > bt[0] - 20.0 and w[0] < split]
        c = [w for w in part if w[0] >= split]
        def span(ps):
            return None if not ps else (' '.join(w[4] for w in ps), min(w[0] for w in ps), max(w[2] for w in ps))
        rows.append((span(b), span(c)))
    if len(rows) != nrows:
        raise SystemExit(f'{pdf}: expected {nrows} rows, got {len(rows)}')
    return bt, bn, rows


def main():
    cells, excel_dir, excel_tsv, ours_dir, out_path = sys.argv[1:6]
    pdf_by_index = {}
    for line in open(excel_tsv, encoding='utf-8'):
        p = line.rstrip('\n').split('\t')
        if len(p) >= 4 and p[2] == 'pdf' and p[3].endswith('.pdf'):
            pdf_by_index[int(p[1])] = os.path.join(excel_dir, p[3])
    manifest = [l.rstrip('\n').split('\t') for l in open(cells)][1:]
    bysheet = collections.OrderedDict()
    for p in manifest:
        bysheet.setdefault(p[0], []).append(p)
    with open(out_path, 'w') as f:
        f.write('sheet\tindex\trow\tkind\tformat_name\tformat\tvalue\tside\tedge_left\tedge_right\ttext\txMin\txMax\tgap_left\tgap_right\n')
        for sheet, rows in bysheet.items():
            index = int(rows[0][1])
            nrows = len({int(r[2]) for r in rows})
            # Excel column A label ends well before column B; B left is > 190 pt.
            eb, ec, erows = measure(pdf_by_index[index], 1, nrows)
            ob, oc, orows = measure(os.path.join(ours_dir, f'ours-{index}.pdf'), 1, nrows)
            for ri, (meta, e, o) in enumerate(zip(rows[0::2], erows, orows)):
                for kind, se, so, be, bo in (('text', e[0], o[0], eb, ob), ('num', e[1], o[1], ec, oc)):
                    if se is None or so is None:
                        continue
                    te, xe0, xe1 = se
                    to, xo0, xo1 = so
                    for side, span, box in (('excel', se, be), ('ours', so, bo)):
                        t, x0, x1 = span
                        f.write('\t'.join(map(str, [sheet, index, meta[2], kind, meta[4], meta[5], meta[6],
                            side, f'{box[0]:.3f}', f'{box[2]:.3f}', t, f'{x0:.3f}', f'{x1:.3f}',
                            f'{x0 - box[0]:.3f}', f'{box[2] - x1:.3f}'])) + '\n')
    print(out_path)


if __name__ == '__main__':
    main()
