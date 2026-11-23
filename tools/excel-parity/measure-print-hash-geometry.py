#!/usr/bin/env python3
"""Measures Excel's printed hash-overflow geometry from a fresh direct PDF export of
make-print-hash-geometry.py's workbook, and compares it with predict-print-hash-geometry.py's locked predictions.

For every test column the solid fill rectangle gives BOTH physical cell edges in PDF points (left/right). Poppler's
`pdftotext -bbox` gives the word box (xMin/xMax) of the printed run; that is an advance box, not literal raster ink,
so side bearings and trailing pads are inside it. Reported quantities:

  nHashes      number of '#' characters in the printed word (0 = it printed the full number, etc.)
  runWidth     xMax - xMin (advance-box width of the rigid run)
  xMinLeft     xMin - cellLeft      xMaxLeft    xMax - cellLeft
  rightGap     cellRight - xMax     xMinRight   xMin - cellRight

Diagnostics flag the cases the model must not silently swallow: 'full-numeric' (the column got wide enough to
print digits instead of hashes), 'outside-cell' (word box crosses the physical cell edge, i.e. clipping risk),
'missing-word', 'extra-word', 'no-fill' and 'width-mismatch' (measured physical width vs the pre-registered
5.25*w+3.75 formula). Fonts actually embedded on each page are recorded as 'pdf-fonts'.

    python3 tools/excel-parity/measure-print-hash-geometry.py META.tsv PREDICTIONS.tsv EXCELDIR OUT.tsv
"""
import collections
import os
import re
import subprocess
import sys
import tempfile


def fill_color(i):
    return (0xE0 + (i % 16), 0x20 + ((i * 37) % 0xC0), 0x30 + ((i * 71) % 0xA0))


def svg_fills(pdf, page):
    with tempfile.TemporaryDirectory() as tmp:
        base = os.path.join(tmp, 'p.svg')
        subprocess.run(['pdftocairo', '-f', str(page), '-l', str(page), '-svg', pdf, base],
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        text = open(base if os.path.exists(base) else base[:-4], encoding='utf-8').read()
    out = collections.defaultdict(lambda: [float('inf'), -float('inf')])
    for m in re.finditer(r'<path\b([^>]*)/>', text):
        a = m.group(1)
        c = re.search(r'fill="rgb\(([^)]*)\)"', a)
        d = re.search(r'\bd="([^"]*)"', a)
        if not c or not d:
            continue
        nums = [float(v) for v in re.findall(r'-?\d+\.?\d*', d.group(1))]
        if len(nums) < 8:
            continue
        xs = [nums[0], nums[2], nums[4], nums[6]]
        rgb = tuple(round(float(v.strip().rstrip('%')) * 255 / 100) for v in c.group(1).split(','))
        lo, hi = out[rgb]
        out[rgb] = [min(lo, min(xs)), max(hi, max(xs))]
    return out


def page_words(pdf, page):
    out = subprocess.run(['pdftotext', '-bbox', '-f', str(page), '-l', str(page), pdf, '-'],
                         capture_output=True, text=True).stdout
    return [(float(a), float(b), float(x), float(y), t) for a, b, x, y, t in re.findall(
        r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">([^<]*)</word>', out)]


def pdf_pages(pdf):
    out = subprocess.run(['pdfinfo', pdf], capture_output=True, text=True).stdout
    return int(re.search(r'^Pages:\s+(\d+)', out, re.M).group(1))


def pdf_fonts(pdf, page):
    out = subprocess.run(['pdffonts', '-f', str(page), '-l', str(page), pdf], capture_output=True, text=True).stdout
    return ';'.join(l.split()[0] for l in out.splitlines()[2:] if l.strip())


def main():
    meta_path, pred_path, excel_dir, out_path = sys.argv[1:5]
    meta = [l.rstrip('\n').split('\t') for l in open(meta_path, encoding='utf-8')]
    mhead, meta = meta[0], meta[1:]
    pred = {}
    for l in open(pred_path, encoding='utf-8'):
        if l.startswith('#'):
            continue
        p = l.rstrip('\n').split('\t')
        if p[0] == 'sheet':
            continue
        pred[(p[0], p[1])] = dict(zip(['pred_W_pt', 'pred_advance_pt', 'pred_n'], p[9:12]))

    # sheet name -> PDF from the export TSV (*.tsv next to the PDFs)
    stem = os.path.basename(out_path).replace('.tsv', '')
    tsv = os.path.join(excel_dir, sorted(f for f in os.listdir(excel_dir) if f.endswith('.tsv'))[0])
    pdf_of = {}
    for l in open(tsv, encoding='utf-8'):
        p = l.rstrip('\n').split('\t')
        if len(p) >= 4 and p[2] == 'pdf' and p[3].endswith('.pdf'):
            pdf_of[p[0]] = p[3]

    by_sheet = collections.defaultdict(list)
    for m in meta:
        d = dict(zip(mhead, m))
        by_sheet[d['sheet']].append(d)

    rows = ['sheet\tcolumn\twidth_chars\talign\tface\tsize_pt\titalic\tscale\tkind\trow\ttext\tnHashes\t'
            'xMin\txMax\trunWidth\tcellLeft\tcellRight\tcellWidth\txMinLeft\txMaxLeft\trightGap\txMinRight\t'
            'pred_W_pt\tpred_n\tdiagnostics\tpage\tpdf_fonts']
    for sheet in by_sheet:
        pdf = os.path.join(excel_dir, pdf_of[sheet])
        pages = pdf_pages(pdf)
        for page in range(1, pages + 1):
            fills = svg_fills(pdf, page)
            words = page_words(pdf, page)
            fonts = pdf_fonts(pdf, page)
            for d in by_sheet[sheet]:
                col = int(d['column'])
                rgb = fill_color(col - 2)
                rect = fills.get(rgb)
                if rect is None:
                    continue
                left, right = rect
                inside = sorted([w for w in words if left - 2 <= (w[0] + w[2]) / 2 <= right + 2], key=lambda w: w[1])
                column_diags = []
                if len(inside) == 0:
                    column_diags.append('missing-word')
                elif len(inside) > 2:
                    column_diags.append(f'extra-word:{len(inside)}')
                pr = pred[(sheet, d['column'])]
                w_meas = right - left
                if abs(w_meas - float(pr['pred_W_pt'])) > 0.15:
                    column_diags.append('width-mismatch')
                for ri, w in enumerate(inside[:2]):
                    x0, y0, x1, y1, text = w
                    n = len(text) if text and set(text) == {'#'} else 0
                    row = '1' if ri == 0 else '2'
                    diags = list(column_diags)
                    if row == '1' and d['kind'] == 'overflow' and n == 0:
                        diags.append('full-numeric' if re.fullmatch(r'[-+()\d.,%$ ]+', text or '') else 'unexpected-text')
                    if x0 < left - 0.6 or x1 > right + 0.6:
                        diags.append('outside-cell')
                    if d['kind'] == 'literal' and row == '2' and (x1 - x0) > (right - left - 7.5) + 0.2:
                        diags.append('clip-risk')
                    if d['kind'] == 'overflow' and row == '1' and n and int(pr['pred_n']) != n:
                        diags.append(f"count-vs-pred:{pr['pred_n']}->{n}")
                    rows.append('\t'.join([
                        sheet, d['column'], d['width_chars'], d['align'], d['face'], d['size_pt'], d['italic'],
                        d['scale'], d['kind'], row, text, str(n),
                        f'{x0:.3f}', f'{x1:.3f}', f'{x1 - x0:.3f}', f'{left:.3f}', f'{right:.3f}', f'{right - left:.3f}',
                        f'{x0 - left:.3f}', f'{x1 - left:.3f}', f'{right - x1:.3f}', f'{x0 - right:.3f}',
                        pr['pred_W_pt'], pr['pred_n'], ';'.join(diags) or 'ok', str(page), fonts]))
    with open(out_path, 'w', encoding='utf-8', newline='\n') as f:
        f.write('\n'.join(rows) + '\n')
    print(f'wrote {len(rows) - 1} measurements to {out_path}')


if __name__ == '__main__':
    main()
