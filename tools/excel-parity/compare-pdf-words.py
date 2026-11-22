#!/usr/bin/env python3
"""Compares the words of two PDFs of a sheet page by page: Excel 2019's (excel-print-pages.sh) and ours
(tools/spreadsheet-pdf-runner). Each word of a page of one is paired with the nearest word of the same text on the same
page of the other; per page: words paired, the median and the largest move of a paired word (x, y in points), and the
words of either PDF not found in the other.

    python3 tools/excel-parity/compare-pdf-words.py EXCEL.pdf OURS.pdf [--words]
"""
import re
import subprocess
import sys
from statistics import median


def pages(path):
    out = subprocess.run(['pdftotext', '-bbox', path, '-'], capture_output=True, text=True).stdout
    result = []
    for page in out.split('<page ')[1:]:
        words = []
        for m in re.finditer(r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">([^<]*)</word>', page):
            words.append((m.group(5), float(m.group(1)), float(m.group(2)), float(m.group(3)), float(m.group(4))))
        result.append(words)
    return result


def compare(excel, ours, show=False):
    a, b = pages(excel), pages(ours)
    print(f'pages: excel {len(a)} ours {len(b)}')
    for n in range(max(len(a), len(b))):
        wa = a[n] if n < len(a) else []
        wb = b[n] if n < len(b) else []
        used = set()
        moves = []
        missing = []
        for text, x0, y0, x1, y1 in wa:
            best = None
            for j, (t, u0, v0, u1, v1) in enumerate(wb):
                if t != text or j in used:
                    continue
                d = abs(u0 - x0) + abs(v0 - y0)
                if best is None or d < best[0]:
                    best = (d, j, u0 - x0, v0 - y0)
            if best is None or best[0] > 60:
                missing.append(text)
            else:
                used.add(best[1])
                moves.append((best[2], best[3], text))
        extra = [wb[j][0] for j in range(len(wb)) if j not in used]
        if moves:
            dx = median(m[0] for m in moves)
            dy = median(m[1] for m in moves)
            worst = max(moves, key=lambda m: abs(m[0]) + abs(m[1]))
            print(f'page {n + 1}: paired {len(moves)}, median dx {dx:+.2f} dy {dy:+.2f}, largest {worst[0]:+.2f} {worst[1]:+.2f} '
                  f'"{worst[2]}", excel only {len(missing)}, ours only {len(extra)}')
        else:
            print(f'page {n + 1}: paired 0, excel only {len(missing)}, ours only {len(extra)}')
        if show:
            print('  excel only:', ' '.join(missing[:30]))
            print('  ours only:', ' '.join(extra[:30]))


if __name__ == '__main__':
    compare(sys.argv[1], sys.argv[2], '--words' in sys.argv)
