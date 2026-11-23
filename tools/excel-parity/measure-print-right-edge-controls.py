#!/usr/bin/env python3
"""Measures the literal-vs-generated right-edge controls and writes the reproducible oracle table.

Metric: Poppler word box `pdftotext -bbox` xMax of the last word to the printed #FFF2CC fill's right edge (the
cell's physical right border). Not raster ink. The full displayed string, its font, and its advance width are
recovered from the same PDFs so that a numeric/literal pair can be checked for identity before its gap is compared.

    python3 tools/excel-parity/measure-print-right-edge-controls.py \
        tools/excel-parity/print-right-edge-controls-cells.tsv \
        EXCEL_DIR EXCEL_TSV OURS_DIR tools/excel-parity/print-right-edge-controls-reference.tsv

EXCEL_DIR holds the fresh direct-export PDFs and EXCEL_TSV is excel-print-pages.sh's sheet/index/pdf table.
OURS_DIR holds ours-<index>.pdf (1-based) from tools/spreadsheet-pdf-runner with
RB_WORD_FONT_DIR=artifacts/fonts/windows:artifacts/fonts/windows-extra.
"""
import html
import os
import re
import subprocess
import sys
import tempfile
from collections import defaultdict

FILL_HINTS = ('94.8', '94.9')  # pdftocairo's rgb(...) for #FFF2CC


def fill_rect(pdf):
    """Right edge of the printed #FFF2CC rectangle, in points."""
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, 'page.svg')
        subprocess.run(['pdftocairo', '-f', '1', '-l', '1', '-svg', pdf, out], check=True,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        svg = out + '.svg' if os.path.exists(out + '.svg') else out
        text = open(svg, encoding='utf-8').read()
    rects = []
    for m in re.finditer(r'<path\s+([^>]*?)/>', text):
        attrs = m.group(1)
        if not any(h in attrs for h in FILL_HINTS):
            continue
        d = re.search(r'd="([^"]*)"', attrs)
        if not d:
            continue
        nums = [float(v) for v in re.findall(r'-?\d+\.?\d*', d.group(1))]
        if len(nums) >= 8:
            xs = [nums[0], nums[2], nums[4], nums[6]]
            rects.append((min(xs), max(xs)))
    return max(r[1] for r in rects) if rects else None


def words(pdf):
    out = subprocess.run(['pdftotext', '-bbox', pdf, '-'], capture_output=True, text=True).stdout
    return [(float(a), float(b), float(x), float(y2), t) for a, b, x, y2, t in
            re.findall(r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">([^<]*)</word>', out)]


def text_elements(pdf):
    """(top, left, width, font_family, text) from pdftohtml's XML, in page pixels."""
    out = subprocess.run(['pdftohtml', '-xml', '-f', '1', '-l', '1', '-stdout', pdf],
                         capture_output=True, text=True, check=True).stdout
    fonts = {m.group(1): html.unescape(m.group(2))
             for m in re.finditer(r'<fontspec id="(\d+)"[^>]*family="([^"]*)"', out)}
    page = re.search(r'<page[^>]*width="(\d+)"[^>]*height="(\d+)"', out)
    page_w = int(page.group(1)) if page else None
    elements = []
    for m in re.finditer(r'<text top="(\d+)" left="(\d+)" width="(\d+)" height="(\d+)" font="(\d+)"[^>]*>(.*?)</text>', out):
        top, left, width, _height, fid, raw = m.groups()
        value = html.unescape(re.sub(r'<[^>]+>', '', raw))
        elements.append((int(top), int(left), int(width), fonts.get(fid, fid), value))
    return elements, page_w


def group_lines(items, y_index, tolerance):
    items = sorted(items, key=lambda v: v[y_index])
    lines = []
    for item in items:
        y = item[y_index]
        if not lines or abs(y - lines[-1][0]) > tolerance:
            lines.append([y, [item]])
        else:
            lines[-1][1].append(item)
    return [line[1] for line in lines]


def measure_pdf(pdf, n_rows):
    border = fill_rect(pdf)
    if border is None:
        raise RuntimeError(f'no #FFF2CC fill found: {pdf}')
    ws = words(pdf)
    elements, _page_px = text_elements(pdf)
    word_lines = group_lines(ws, 1, 3.0)
    elem_lines = group_lines(elements, 0, 4.0)
    if len(word_lines) != n_rows or len(elem_lines) != n_rows:
        raise RuntimeError(f'{pdf}: expected {n_rows} lines, got {len(word_lines)} word lines / {len(elem_lines)} text lines')
    rows = []
    for wline, eline in zip(word_lines, elem_lines):
        right = max(wline, key=lambda w: w[2])
        # The rightmost text element is the displayed value; accounting's separate "$" is farther left.
        elem = max(eline, key=lambda e: e[1] + e[2])
        # The displayed value is the line's words except accounting's separate currency symbol. Joining them
        # reconstructs strings that Poppler splits at spaces (e.g. `0.50 kg`), so the identity check covers the
        # whole rendered value, not only the last word whose xMax is the metric.
        value_words = sorted((w for w in wline if w[4] != '$'), key=lambda w: w[0]) or [right]
        rows.append({
            'text': ' '.join(w[4] for w in value_words),
            'last_word': right[4],
            'xMin': value_words[0][0],
            'xMax': right[2],
            'width': right[2] - value_words[0][0],
            'border': border,
            'gap': border - right[2],
            'font': elem[3],
        })
    return rows


def main():
    manifest_path, excel_dir, excel_tsv, ours_dir, out_path = sys.argv[1:6]
    pdf_by_index = {}
    for line in open(excel_tsv, encoding='utf-8'):
        parts = line.rstrip('\n').split('\t')
        if len(parts) >= 4 and parts[2] == 'pdf' and parts[3].endswith('.pdf'):
            pdf_by_index[int(parts[1])] = (parts[0], os.path.join(excel_dir, parts[3]))
    cells = defaultdict(list)
    with open(manifest_path, encoding='utf-8') as f:
        header = f.readline().rstrip('\n').split('\t')
        for line in f:
            if not line.strip():
                continue
            cells[line.rstrip('\n').split('\t')[0]].append(dict(zip(header, line.rstrip('\n').split('\t'))))
    out = open(out_path, 'w', encoding='utf-8')
    cols = ('sheet', 'index', 'row', 'kind', 'pair', 'expected_text', 'face', 'size', 'style', 'scale', 'format',
            'excel_text', 'excel_last_word', 'excel_xMin', 'excel_xMax', 'excel_width', 'excel_border', 'excel_gap',
            'excel_font', 'ours_text', 'ours_last_word', 'ours_xMin', 'ours_xMax', 'ours_width', 'ours_border',
            'ours_gap', 'ours_font', 'gap_delta')
    out.write('\t'.join(cols) + '\n')
    for sheet, rows in cells.items():
        rows.sort(key=lambda r: int(r['row']))
        index = int(rows[0]['index'])
        if index not in pdf_by_index:
            print(f'missing Excel PDF for sheet {sheet} index {index}', file=sys.stderr)
            continue
        excel_pdf = pdf_by_index[index][1]
        ours_pdf = os.path.join(ours_dir, f'ours-{index}.pdf')
        if not os.path.exists(ours_pdf):
            print(f'missing ours PDF {ours_pdf}', file=sys.stderr)
            continue
        try:
            erows = measure_pdf(excel_pdf, len(rows))
            orows = measure_pdf(ours_pdf, len(rows))
        except RuntimeError as exc:
            print(exc, file=sys.stderr)
            continue
        for cell, e, o in zip(rows, erows, orows):
            vals = [sheet, cell['index'], cell['row'], cell['kind'], cell['pair'], cell['expected_text'],
                    cell['face'], cell['size'], cell['style'], cell['scale'], cell['format'],
                    e['text'], e['last_word'], f"{e['xMin']:.3f}", f"{e['xMax']:.3f}", f"{e['width']:.3f}",
                    f"{e['border']:.3f}", f"{e['gap']:.3f}", e['font'],
                    o['text'], o['last_word'], f"{o['xMin']:.3f}", f"{o['xMax']:.3f}", f"{o['width']:.3f}",
                    f"{o['border']:.3f}", f"{o['gap']:.3f}", o['font'], f"{o['gap'] - e['gap']:+.3f}"]
            out.write('\t'.join(str(v) for v in vals) + '\n')
    out.close()


if __name__ == '__main__':
    main()
