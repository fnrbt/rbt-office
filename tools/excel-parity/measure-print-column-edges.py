#!/usr/bin/env python3
"""Measures printed column geometry from vector objects in a direct Excel PDF.

A solid-fill rectangle can be expanded for coverage, so the cell edge is read
from the thin black grid border as well as from the fill:

  * fill rectangles: `r g b rg  x y w h re f*`
  * grid borders: black thin filled rectangles or black stroked paths
    (`re f`, `m/l S`, with the `w` line width) = the border centreline.

For every test column: fill edges/width, border centreline edges/width, adjacent
fill overlap, word box and hash count.  Derived:

    fill_width - border_width          total fill expansion
    fill_overlap with next column      same expansion, measured independently
    border_right - border_left         true printed cell width vs layout width

All comparisons are page-local, so page x origins are removed.

    python3 tools/excel-parity/measure-print-column-edges.py META.tsv PRED.tsv EXCELDIR OUT.tsv
"""
import collections
import os
import re
import subprocess
import sys
import zlib


def pdf_streams(path):
    data = open(path, 'rb').read()
    out = []
    for m in re.finditer(rb'stream\r?\n', data):
        start = m.end()
        end = data.find(b'endstream', start)
        if end < 0:
            continue
        try:
            dec = zlib.decompress(data[start:end])
        except zlib.error:
            continue
        if b' re' in dec or b'BT' in dec:
            out.append(dec.decode('latin1'))
    return out


def pdf_pages(path):
    text = subprocess.run(['pdfinfo', path], capture_output=True, text=True).stdout
    return int(re.search(r'^Pages:\s+(\d+)', text, re.M).group(1))


def page_words(path, page):
    text = subprocess.run(['pdftotext', '-bbox', '-f', str(page), '-l', str(page), path, '-'],
                          capture_output=True, text=True).stdout
    return [(float(a), float(b), float(c), float(d), t) for a, b, c, d, t in re.findall(
        r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">([^<]*)</word>', text)]


NUM = r'-?\d+(?:\.\d+)?'
OP_RE = re.compile(
    rf'({NUM})\s+({NUM})\s+({NUM})\s+({NUM})\s+re\b'
    rf'|({NUM})\s+({NUM})\s+({NUM})\s+rg\b'
    rf'|({NUM})\s+g\b'
    rf'|({NUM})\s+({NUM})\s+({NUM})\s+RG\b'
    rf'|({NUM})\s+G\b'
    rf'|({NUM})\s+w\b'
    rf'|\b(W\*?)\s+n\b'
    rf'|\b(f\*?|S|s|B\*?|b\*?)\b'
    rf'|({NUM})\s+({NUM})\s+m\b'
    rf'|({NUM})\s+({NUM})\s+l\b')


def parse_content(text):
    fills, strokes = [], []
    fill = stroke = (0.0, 0.0, 0.0)
    lw = 1.0
    rects, segments = [], []
    for m in OP_RE.finditer(text):
        g = m.groups()
        if g[0] is not None:
            rects.append((float(g[0]), float(g[1]), float(g[2]), float(g[3])))
        elif g[4] is not None:
            fill = (float(g[4]), float(g[5]), float(g[6]))
        elif g[7] is not None:
            fill = (float(g[7]),) * 3
        elif g[8] is not None:
            stroke = (float(g[8]), float(g[9]), float(g[10]))
        elif g[11] is not None:
            stroke = (float(g[11]),) * 3
        elif g[12] is not None:
            lw = float(g[12])
        elif g[13] is not None:
            rects, segments = [], []
        elif g[14] is not None:
            op = g[14]
            if op.startswith('f'):
                for r in rects:
                    fills.append((*r, fill))
            elif op in ('S', 's'):
                if segments:
                    strokes.append((segments, stroke, lw))
            elif op[:1] in ('B', 'b'):
                for r in rects:
                    fills.append((*r, fill))
                if segments:
                    strokes.append((segments, stroke, lw))
            rects, segments = [], []
        elif g[15] is not None:
            segments.append((float(g[15]), float(g[16])))
        elif g[17] is not None:
            segments.append((float(g[17]), float(g[18])))
    return fills, strokes


def layout_width(width_chars, scale_pct):
    scale = scale_pct / 100.0
    screen_px = int(((256.0 * width_chars + 128 // 7) / 256.0) * 7.0)
    return int(screen_px * 47.0 / 7.0 + 0.5) * 72.0 / 600.0 * scale


def rgb255(c):
    return tuple(round(v * 255) for v in c)


def tsv(path):
    rows = [l.rstrip('\n').split('\t') for l in open(path, encoding='utf-8') if not l.startswith('#')]
    return rows[0], rows[1:]


def main():
    meta_path, pred_path, excel_dir, out_path = sys.argv[1:5]
    mh, mr = tsv(meta_path)
    meta = {(d['sheet'], int(d['column'])): d for d in (dict(zip(mh, r)) for r in mr)}
    ph, pr = tsv(pred_path)
    pred = {(d['sheet'], d['column']): d for d in (dict(zip(ph, r)) for r in pr)}

    export_tsv = os.path.join(excel_dir, sorted(f for f in os.listdir(excel_dir) if f.endswith('.tsv'))[0])
    pdf_of = {}
    for line in open(export_tsv, encoding='utf-8'):
        p = line.rstrip('\n').split('\t')
        if len(p) >= 4 and p[2] == 'pdf' and p[3].endswith('.pdf'):
            pdf_of[p[0]] = p[3]

    by_sheet = collections.defaultdict(list)
    for (sheet, col), d in meta.items():
        by_sheet[sheet].append((col, d))

    rows = ['sheet\tpage\tcolumn\tblock\tblock_index\twidth_chars\tscale\t'
            'fill_left\tfill_right\tfill_width\tborder_left\tborder_right\tborder_width\t'
            'fill_minus_border\tfill_overlap_next\tborder_gap_next\tx_origin\t'
            'text\ttext_xmin\ttext_xmax\ttext_width\tnHashes\tdecision\t'
            'pred_W_layout\tpred_fill_width\tpred_border_width\tpred_fill_overlap\t'
            'fill_width_minus_pred\tborder_width_minus_pred\tpred_fill_decision\tdiagnostics']
    for sheet, cols in by_sheet.items():
        pdf = os.path.join(excel_dir, pdf_of[sheet])
        pages = pdf_pages(pdf)
        page_streams = [s for s in pdf_streams(pdf) if '/Artifact BMC' in s or s.lstrip().startswith('q')]
        for page in range(1, pages + 1):
            fills, strokes = parse_content(page_streams[page - 1])
            words = page_words(pdf, page)
            present = []
            for col, d in cols:
                want = rgb255(tuple(int(d['fill_rgb'][i:i + 2], 16) / 255.0 for i in (0, 2, 4)))
                cands = [f for f in fills if rgb255(f[4]) == want]
                if cands:
                    present.append((col, d, max(cands, key=lambda r: r[3] * r[2])))
            present.sort(key=lambda t: t[2][0])
            vlines = []
            for x, y, w, h, c in fills:
                if max(c) <= 0.12 and h >= 5 and w <= 3.0:
                    vlines.append(x + w / 2.0)
            for pts, c, lw in strokes:
                if max(c) > 0.12:
                    continue
                for (x0, y0), (x1, y1) in zip(pts, pts[1:]):
                    if abs(x1 - x0) <= 0.2 and abs(y1 - y0) >= 5:
                        vlines.append((x0 + x1) / 2.0)
            vlines = sorted(vlines)
            merged = []
            for x in vlines:
                if not merged or x - merged[-1] > 0.3:
                    merged.append(x)
                else:
                    merged[-1] = (merged[-1] + x) / 2.0
            vlines = merged

            def border_pair(fl, fr):
                left = min((x for x in vlines if abs(x - fl) <= 1.2), key=lambda x: abs(x - fl), default=None)
                right = min((x for x in vlines if abs(x - fr) <= 1.2), key=lambda x: abs(x - fr), default=None)
                return left, right

            origin = present[0][2][0] if present else 0.0
            for i, (col, d, (x, y, w, h, c)) in enumerate(present):
                fl, fr = x, x + w
                bl, br = border_pair(fl, fr)
                nxt = present[i + 1][2][0] if i + 1 < len(present) else None
                prev_right = present[i - 1][2][0] + present[i - 1][2][2] if i > 0 else None
                inside = sorted([wd for wd in words if fl - 1 <= (wd[0] + wd[2]) / 2 <= fr + 1],
                                key=lambda wd: wd[1])
                text = text_xmin = text_xmax = text_width = ''
                n_hashes = decision = ''
                if inside:
                    top = inside[0]
                    text = top[4]
                    text_xmin, text_xmax = top[0], top[2]
                    text_width = f'{top[2] - top[0]:.3f}'
                    n_hashes = str(len(text)) if text and set(text) == {'#'} else '0'
                    decision = 'hash' if n_hashes != '0' else 'full'
                prd = pred[(sheet, d['column'])]
                diags = []
                if bl is None or br is None:
                    diags.append('missing-border')
                rows.append('\t'.join([
                    sheet, str(page), d['column'], d['block'], d['block_index'], d['width_chars'],
                    d['scale'], f'{fl:.3f}', f'{fr:.3f}', f'{w:.3f}',
                    f'{bl:.3f}' if bl is not None else '',
                    f'{br:.3f}' if br is not None else '',
                    f'{br - bl:.3f}' if bl is not None and br is not None else '',
                    f'{(fr - fl) - (br - bl):.3f}' if bl is not None and br is not None else '',
                    f'{fr - nxt:.3f}' if nxt is not None else '',
                    f'{bl - prev_right:.3f}' if (bl is not None and prev_right is not None) else '',
                    f'{origin:.3f}', text,
                    f'{text_xmin:.3f}' if text_xmin != '' else '',
                    f'{text_xmax:.3f}' if text_xmax != '' else '', text_width,
                    n_hashes, decision,
                    prd['pred_W_layout_pt'], prd['pred_fill_width_pt'], prd['pred_border_width_pt'],
                    prd['pred_fill_overlap_pt'],
                    f'{w - float(prd["pred_fill_width_pt"]):.3f}',
                    f'{(br - bl) - float(prd["pred_border_width_pt"]):.3f}' if bl is not None and br is not None else '',
                    prd['pred_fill_decision'], ';'.join(diags) or 'ok']))
    with open(out_path, 'w', encoding='utf-8', newline='\n') as f:
        f.write('\n'.join(rows) + '\n')
    print(f'wrote {len(rows) - 1} measurements to {out_path}')


if __name__ == '__main__':
    main()
