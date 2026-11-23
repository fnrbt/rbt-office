#!/usr/bin/env python3
"""Measures how far a printed number's word box stands from its cell's right border, in a fresh Excel 2019 direct
PDF export and in tools/spreadsheet-pdf-runner's PDF of the same sheet.

The metric is Poppler's word box from `pdftotext -bbox` (xMin/xMax), not the literal raster outline: xMax is where
Poppler ends the word, so the gap includes the font's right side bearing and any trailing format pad. A difference
between two word-box gaps is therefore an observed placement difference, not a directly measured glyph outline.

The fixture (make-print-right-edge-fixture.py) fills the value cell #FFF2CC, whose printed rectangle's right edge is
the cell's right border. Each PDF is converted to SVG with pdftocairo only to read that rectangle; the words come
from pdftotext -bbox.

    python3 tools/excel-parity/measure-print-right-edge.py TSV EXCELDIR OURSDIR [--words]

TSV is excel-print-pages.sh's output (sheet <tab> n <tab> pdf <tab> name); EXCELDIR holds its PDFs; OURSDIR holds
ours-<n>.pdf for sheet n (1-based). Print/run the runner with
RB_WORD_FONT_DIR=artifacts/fonts/windows:artifacts/fonts/windows-extra so the fixture's Aptos Narrow faces are
really used (without them the runner falls back to Calibri and the oracle does not apply).
Prints one line per printed number: sheet, text, excel word-box gap, ours word-box gap, their delta.
"""
import os
import re
import subprocess
import sys
import tempfile


def fill_rect(pdf):
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, 'page.svg')
        subprocess.run(['pdftocairo', '-f', '1', '-l', '1', '-svg', pdf, out], check=True,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        svg = out + '.svg' if os.path.exists(out + '.svg') else out
        text = open(svg, encoding='utf-8').read()
    rects = []
    for m in re.finditer(r'<path\s+([^>]*?)/>', text):
        attrs = m.group(1)
        if '94.8' not in attrs and '94.9' not in attrs:
            continue  # the #FFF2CC fixture fill
        d = re.search(r'd="([^"]*)"', attrs)
        if not d:
            continue
        nums = [float(v) for v in re.findall(r'-?\d+\.?\d*', d.group(1))]
        if len(nums) >= 8:
            xs = [nums[0], nums[2], nums[4], nums[6]]
            rects.append((min(xs), max(xs)))
    # The four rows of the value cell share this fill; its right edge is the cell's right border.
    return (min(r[0] for r in rects), max(r[1] for r in rects)) if rects else None


def words(pdf):
    out = subprocess.run(['pdftotext', '-bbox', pdf, '-'], capture_output=True, text=True).stdout
    return [(float(a), float(b), float(x), t) for a, b, x, t in
            re.findall(r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="[\d.]+">([^<]*)</word>', out)]


def main():
    tsv, excel_dir, ours_dir = sys.argv[1:4]
    show_all = '--words' in sys.argv
    rows = []
    for line in open(tsv, encoding='utf-8'):
        parts = line.rstrip('\n').split('\t')
        if len(parts) < 4 or parts[2] != 'pdf' or not parts[3].endswith('.pdf'):
            continue
        sheet, index, name = parts[0], int(parts[1]), parts[3]
        excel_pdf = os.path.join(excel_dir, name)
        ours_pdf = os.path.join(ours_dir, f'ours-{index}.pdf')
        if not (os.path.exists(excel_pdf) and os.path.exists(ours_pdf)):
            continue
        left_e, be = fill_rect(excel_pdf)
        left_o, bo = fill_rect(ours_pdf)
        we, wo = words(excel_pdf), words(ours_pdf)
        for text, ex, ou in matched(we, wo, left_e):
            if not show_all and not re.search(r'[\d#]', text):
                continue
            eg, og = be - ex, bo - ou
            print(f'{sheet}\t{index}\t{text}\t{ex:.3f}\t{be:.3f}\t{eg:.3f}\t{ou:.3f}\t{bo:.3f}\t{og:.3f}\t{og - eg:+.3f}')


def matched(excel_words, ours_words, column_left):
    """Pairs words of the same text nearest in y, inside the filled value column (not the labels)."""
    used = set()
    for x0, y0, x1, text in excel_words:
        if x0 < column_left - 1:
            continue
        best = None
        for j, (u0, v0, u1, same) in enumerate(ours_words):
            if same != text or j in used:
                continue
            d = abs(v0 - y0)
            if best is None or d < best[0]:
                best = (d, j, u1)
        if best is not None and best[0] < 3:
            used.add(best[1])
            yield text, x1, best[2]


if __name__ == '__main__':
    main()
