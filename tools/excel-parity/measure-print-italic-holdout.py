#!/usr/bin/env python3
"""Measures the printed italic reserve of the blind holdouts in Excel 2019's direct PDFs and in
tools/spreadsheet-pdf-runner's PDFs: per pair, the italic word-box gap less the upright gap. The metric is Poppler's
`pdftotext -bbox` xMax to the printed #FFF2CC fill's right edge (`measure-print-right-edge.py`).

    python3 tools/excel-parity/measure-print-italic-holdout.py EXCELDIR OURSDIR [--second]

EXCELDIR holds synthetic-print-italic-holdout[-2]-<n>.pdf (from excel-print-pages.sh); OURSDIR holds ours-<n>.pdf for
sheet n (1-based) for the same workbook. Ours must be rendered with RB_WORD_FONT_DIR including artifacts/fonts/windows
and windows-extra (Aptos Narrow, Tahoma, Garamond, etc.).
"""
import importlib.util
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location('right_edge', os.path.join(HERE, 'measure-print-right-edge.py'))
right_edge = importlib.util.module_from_spec(spec)
spec.loader.exec_module(right_edge)

FIRST = [
    ('H01_verdana13_85', 'H02_verdana13_85_italic', 1, 2, 2.160),
    ('H03_georgia15', 'H04_georgia15_italic', 3, 4, 2.760),
    ('H05_segoe10p5_90', 'H06_segoe10p5_90_italic', 5, 6, 2.040),
    ('H07_verdana13_85_hash', 'H08_verdana13_85_hash_italic', 7, 8, 2.160),
    ('H09_aptos14_acct', 'H10_aptos14_acct_italic', 9, 10, 0.000),
    ('H11_aptos14_acctlit', 'H12_aptos14_acctlit_italic', 11, 12, 2.760),
    ('H13_aptos14_overflow', 'H14_aptos14_overflow_italic', 13, 14, None),
]
SECOND = [
    ('H15_tahoma16_80', 'H16_tahoma16_80_italic', 1, 2, 2.520),
    ('H17_garamond11_115', 'H18_garamond11_115_italic', 3, 4, 2.160),
    ('H19_verdana9_110', 'H20_verdana9_110_italic', 5, 6, 1.920),
    ('H21_calibri10p5_95', 'H22_calibri10p5_95_italic', 7, 8, 1.800),
]


def page_columns(pdf):
    left, right = right_edge.fill_rect(pdf)
    return right, [w for w in right_edge.words(pdf) if w[0] >= left - 1.0]


def reserve_of(pdf_u, pdf_i):
    right_u, words_u = page_columns(pdf_u)
    right_i, words_i = page_columns(pdf_i)
    used = set()
    values = []
    for x0, y0, x1, text in words_u:
        best = None
        for j, (a, b, c, same) in enumerate(words_i):
            if same != text or j in used:
                continue
            d = abs(b - y0)
            if best is None or d < best[0]:
                best = (d, j, c)
        if best and best[0] < 3:
            used.add(best[1])
            values.append((right_i - best[2]) - (right_u - x1))
    return values


def main():
    excel_dir, ours_dir = sys.argv[1:3]
    second = '--second' in sys.argv
    stem = 'synthetic-print-italic-holdout2' if second else 'synthetic-print-italic-holdout'
    configs = SECOND if second else FIRST
    print('pair\tcase\tpredicted\tobserved\tresidual\tn\tours')
    for upright, italic, nu, ni, predicted in configs:
        excel_u = os.path.join(excel_dir, f'{stem}-{nu}.pdf')
        excel_i = os.path.join(excel_dir, f'{stem}-{ni}.pdf')
        ours_u = os.path.join(ours_dir, f'ours-{nu}.pdf')
        ours_i = os.path.join(ours_dir, f'ours-{ni}.pdf')
        if not all(os.path.exists(p) for p in (excel_u, excel_i, ours_u, ours_i)):
            continue
        excel_values = reserve_of(excel_u, excel_i)
        ours_values = reserve_of(ours_u, ours_i)
        n = min(len(excel_values), len(ours_values))
        if not n:
            continue
        observed = sum(excel_values[:n]) / n
        our = sum(ours_values[:n]) / n
        residual = '' if predicted is None else f'{observed - predicted:+.4f}'
        print(f'{upright}/{italic}\t{stem}\t{"" if predicted is None else f"{predicted:.3f}"}\t{observed:.4f}\t{residual}\t{n}\tours={our:.4f}')


if __name__ == '__main__':
    main()
