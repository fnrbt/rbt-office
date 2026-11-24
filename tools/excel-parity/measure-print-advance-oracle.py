#!/usr/bin/env python3
"""Measures the printed-advance oracle workbooks from a fresh direct Excel PDF export.

For every cell it reads three independent descriptions of the same printed run:

  1. the PDF text positioning operators (`Tf` size, `Tm` origin, `Tc`, `TJ`
     strings and adjustments) and the embedded font's `/Widths`, giving the pen
     advance Excel actually used for drawing;
  2. Poppler's `pdftotext -bbox` word box, the extracted run bbox used by the
     earlier fit work;
  3. the solid fill rectangle, giving the physical cell edges (W).

`pdfAdvance = bboxWidth` on every run is what makes (1) and (2) interchangeable;
`cellWidth` is the independent W.  The script also records the printed decision
(full number or `###` run) and both right-edge gaps.

    python3 tools/excel-parity/measure-print-advance-oracle.py META.tsv EXCELDIR OUT.tsv
"""
import collections
import importlib.util
import os
import re
import subprocess
import sys
import tempfile
import zlib

_here = os.path.dirname(os.path.abspath(__file__))
_spec = importlib.util.spec_from_file_location(
    'hashmeasure', os.path.join(_here, 'measure-print-hash-geometry.py'))
hashmeasure = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(hashmeasure)


# ---------------- PDF object / content-stream reader ----------------
def objects_of(data):
    return {int(m.group(1)): data[m.end():data.find(b'endobj', m.end())]
            for m in re.finditer(rb'(\d+) 0 obj\b', data)}


def stream_of(body):
    m = re.search(rb'stream\r?\n', body)
    if not m:
        return None
    raw = body[m.end():body.find(b'endstream', m.end())]
    if raw.endswith(b'\r\n'):
        raw = raw[:-2]
    elif raw.endswith(b'\n'):
        raw = raw[:-1]
    try:
        return zlib.decompress(raw)
    except zlib.error:
        return raw


def page_objects(objs):
    pages = [n for n, b in objs.items()
             if re.search(rb'/Type\s*/Page\b', b) and not re.search(rb'/Type\s*/Pages\b', b)]
    for n, b in objs.items():
        if re.search(rb'/Type\s*/Pages\b', b):
            km = re.search(rb'/Kids\s*\[(.*?)\]', b, re.S)
            if km:
                kids = [int(x) for x in re.findall(rb'(\d+) 0 R', km.group(1))]
                if kids and all(k in pages for k in kids):
                    return kids
    return sorted(pages)


def font_refs(objs, page_body, seen=0):
    m = re.search(rb'/Font\s*<<(.*?)>>', page_body, re.S)
    if not m:
        rm = re.search(rb'/Resources\s+(\d+) 0 R', page_body)
        if rm and seen < 4:
            return font_refs(objs, objs[int(rm.group(1))], seen + 1)
        return {}
    return {nm.decode(): int(ref) for nm, ref in re.findall(rb'/(\w+)\s+(\d+) 0 R', m.group(1))}


def font_widths(objs, ref):
    b = objs[ref]
    wm = re.search(rb'/Widths\s+(\d+) 0 R', b)
    fc = re.search(rb'/FirstChar\s+(\d+)', b)
    if wm and fc:
        arr = objs[int(wm.group(1))]
        nums = [int(x) for x in re.findall(rb'-?\d+', arr.split(b'[')[1].split(b']')[0])]
        first = int(fc.group(1))
        return {first + i: v / 1000.0 for i, v in enumerate(nums)}
    dm = re.search(rb'/DescendantFonts\s*\[\s*(\d+) 0 R', b)
    if dm:
        db = objs[int(dm.group(1))]
        wm = re.search(rb'/W\s*\[(.*?)\]', db, re.S)
        if wm:
            toks = re.findall(rb'\[[^\]]*\]|-?\d+', wm.group(1))
            vals = [('arr', [int(x) for x in re.findall(rb'-?\d+', t)]) if t.startswith(b'[')
                    else ('n', int(t)) for t in toks]
            out = {}
            i = 0
            while i < len(vals):
                if vals[i][0] == 'n' and i + 1 < len(vals) and vals[i + 1][0] == 'arr':
                    c = vals[i][1]
                    for w in vals[i + 1][1]:
                        out[c] = w / 1000.0
                        c += 1
                    i += 2
                elif i + 2 < len(vals) and all(v[0] == 'n' for v in vals[i:i + 3]):
                    c0, c1, w = vals[i][1], vals[i + 1][1], vals[i + 2][1]
                    for c in range(c0, c1 + 1):
                        out[c] = w / 1000.0
                    i += 3
                else:
                    i += 1
            return out
    return {}


def unescape(b):
    out = bytearray()
    i = 0
    while i < len(b):
        c = b[i]
        if c == 0x5C and i + 1 < len(b):
            i += 1
            c = b[i]
            if c in b'nrtbf':
                out.append({ord('n'): 10, ord('r'): 13, ord('t'): 9, ord('b'): 8, ord('f'): 12}[c])
                i += 1
            elif c in b'()\\':
                out.append(c)
                i += 1
            elif 0x30 <= c <= 0x37:
                oct_s = bytes([c])
                i += 1
                while i < len(b) and len(oct_s) < 3 and 0x30 <= b[i] <= 0x37:
                    oct_s += bytes([b[i]])
                    i += 1
                out.append(int(oct_s, 8) & 0xFF)
            elif c in b'\r\n':
                if c == 13 and i < len(b) and b[i] == 10:
                    i += 1
                i += 1
            else:
                out.append(c)
                i += 1
        else:
            out.append(c)
            i += 1
    return bytes(out)


def tj_items(b):
    items = []
    i = 0
    while i < len(b):
        c = b[i]
        if c in b' \r\n\t':
            i += 1
        elif c == 0x28:
            depth, j = 1, i + 1
            while j < len(b) and depth:
                if b[j] == 0x5C:
                    j += 2
                    continue
                if b[j] == 0x28:
                    depth += 1
                elif b[j] == 0x29:
                    depth -= 1
                j += 1
            items.append(('s', unescape(b[i + 1:j - 1])))
            i = j
        elif c == 0x3C:
            j = b.find(b'>', i + 1)
            items.append(('s', bytes.fromhex(re.sub(rb'\s', b'', b[i + 1:j]).decode())))
            i = j + 1
        else:
            m = re.match(rb'-?\d+\.?\d*', b[i:])
            if not m:
                i += 1
                continue
            items.append(('n', float(m.group(0))))
            i += m.end()
    return items


def text_runs(content):
    """Runs one BT block at a time while tracking the graphics/text state.

    Excel wraps some pages' cells in a single q ... Q, so `Tf`, `Tc` and the
    font size persist from one BT block to the next; elsewhere each cell has
    its own q ... Q.  q/Q save and restore the text state (it is part of the
    graphics state), which is exactly what this loop models.
    """
    runs = []
    state = {'font': None, 'fs': None, 'tc': 0.0}
    stack = []
    for m in re.finditer(rb'BT(.*?)ET|(?<![A-Za-z])q(?![A-Za-z])|(?<![A-Za-z])Q(?![A-Za-z])', content, re.S):
        tok = m.group(0)
        if tok == b'q':
            stack.append(dict(state))
            continue
        if tok == b'Q':
            if stack:
                state = stack.pop()
            continue
        blk = m.group(1)
        fm = re.search(rb'/(\w+)\s+([\d.]+)\s+Tf', blk)
        if fm:
            state['font'], state['fs'] = fm.group(1).decode(), float(fm.group(2))
        tcm = re.findall(rb'([-\d.]+)\s+Tc', blk)
        if tcm:
            state['tc'] = float(tcm[-1])
        tm = re.search(rb'([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+Tm', blk)
        if not tm or state['font'] is None:
            continue
        text, adjs = b'', []
        for show in re.finditer(rb'\[(.*?)\]\s*TJ|\((?:[^()\\]|\\.)*\)\s*Tj', blk, re.S):
            seg = show.group(0)
            if seg.rstrip().endswith(b'TJ'):
                items = tj_items(re.match(rb'\[(.*)\]\s*TJ', seg, re.S).group(1))
            else:
                lit = re.match(rb'(\((?:[^()\\]|\\.)*\))\s*Tj', seg, re.S).group(1)
                items = [('s', unescape(lit[1:-1]))]
            for kind, val in items:
                if kind == 's':
                    text += val
                else:
                    adjs.append(val)
        runs.append(dict(font=state['font'], fs=state['fs'],
                         x=float(tm.group(5)), y=float(tm.group(6)),
                         tc=state['tc'], text=text, adjs=adjs))
    return runs

def pdf_pages_ops(pdf):
    data = open(pdf, 'rb').read()
    objs = objects_of(data)
    out = []
    for pn in page_objects(objs):
        body = objs[pn]
        cm = re.search(rb'/Contents\s+(\d+) 0 R', body)
        if not cm:
            continue
        content = stream_of(objs[int(cm.group(1))]) or b''
        fonts = font_refs(objs, body)
        widths = {k: font_widths(objs, v) for k, v in fonts.items()}
        out.append((content, widths))
    return out


def advance_of(run, widths):
    w = widths.get(run['font'], {})
    n = len(run['text'])
    base = sum(w.get(c, 0) for c in run['text']) * run['fs']
    adv = base - sum(run['adjs']) * run['fs'] / 1000.0 + (n - 1) * run['tc']
    return base, adv, n


def main():
    if len(sys.argv) != 4:
        raise SystemExit(__doc__)
    meta_path, excel_dir, out_path = sys.argv[1:]
    meta = [l.rstrip('\n').split('\t') for l in open(meta_path, encoding='utf-8')]
    head, meta = meta[0], meta[1:]
    tsv = os.path.join(excel_dir, sorted(f for f in os.listdir(excel_dir) if f.endswith('.tsv'))[0])
    pdf_of = {}
    for line in open(tsv, encoding='utf-8'):
        p = line.rstrip('\n').split('\t')
        if len(p) >= 4 and p[2] == 'pdf' and p[3].endswith('.pdf'):
            pdf_of[p[0]] = p[3]
    by_sheet = collections.defaultdict(list)
    for m in meta:
        d = dict(zip(head, m))
        by_sheet[d['sheet']].append(d)

    rows = ['sheet\tcolumn\trow\twidth_chars\tkind\trole\tface\tsize_pt\titalic\tscale\tformat\tvalue\t'
            'text_expected\ttext_run\tnHashes\tpdf_font\tpdf_size\ttc\tn_glyphs\ttm_x\ttm_y\t'
            'base_pt\tpdf_adv_pt\tpdf_adv_trail_pt\tbbox_xMin\tbbox_xMax\tbbox_width\tadv_minus_bbox\t'
            'cellLeft\tcellRight\tcellWidth\tleftGap\trightGap\tdiagnostics\tpage\tpdf_fonts']
    for sheet, items in by_sheet.items():
        if sheet not in pdf_of:
            continue
        pdf = os.path.join(excel_dir, pdf_of[sheet])
        pages = hashmeasure.pdf_pages(pdf)
        ops = pdf_pages_ops(pdf)
        seen = {}
        for page in range(1, pages + 1):
            fills = hashmeasure.svg_fills(pdf, page)
            words = hashmeasure.page_words(pdf, page)
            fonts = hashmeasure.pdf_fonts(pdf, page)
            content, widths = ops[page - 1]
            runs = text_runs(content)
            for d in items:
                col = int(d['column'])
                rect = fills.get(hashmeasure.fill_color(col - 2))
                if rect is None:
                    continue
                left, right = rect
                inside = sorted([r for r in runs if left - 1.5 <= r['x'] <= right + 1.5],
                                key=lambda r: -r['y'])
                if not inside:
                    continue
                idx = 0 if d['row'] == '1' else 1
                if len(inside) <= idx:
                    continue
                r = inside[idx]
                base, adv, n = advance_of(r, widths)
                text = r['text'].decode('latin1')
                cands = [w for w in words if abs(w[0] - r['x']) < 0.05]
                # two rows can start at the same x; prefer the word whose text is the run
                match = [w for w in cands if w[4] == text] or cands
                if match:
                    w0, y0, x1, y1, wtext = match[0]
                    bw = x1 - w0
                else:
                    w0 = x1 = bw = None
                    wtext = ''
                n_hash = text.count('#')
                diags = []
                if r['text'].decode('latin1') != wtext and text.strip('#') != '':
                    diags.append('bbox-text-differs')
                if bw is not None and abs(adv - bw) > 0.01:
                    diags.append('adv-bbox-diff')
                if r['x'] < left - 0.5 or (x1 is not None and x1 > right + 0.5):
                    diags.append('outside-cell')
                if n_hash and text.strip('# ') != '':
                    diags.append('mixed-hash-word')
                rows.append('\t'.join([
                    sheet, str(col), d['row'], d['width_chars'], d['kind'], d['role'], d['face'],
                    d['size_pt'], d['italic'], d['scale'], d['format'], d['value'],
                    d['text'], text, str(n_hash), r['font'], f"{r['fs']:g}", f"{r['tc']:g}", str(n),
                    f"{r['x']:.5f}", f"{r['y']:.3f}", f'{base:.5f}', f'{adv:.5f}', f'{adv + r["tc"]:.5f}',
                    '' if w0 is None else f'{w0:.5f}', '' if x1 is None else f'{x1:.5f}',
                    '' if bw is None else f'{bw:.5f}', '' if bw is None else f'{adv - bw:.6f}',
                    f'{left:.4f}', f'{right:.4f}', f'{right - left:.4f}',
                    f"{r['x'] - left:.4f}", f"{right - (r['x'] + adv):.4f}",
                    ';'.join(sorted(set(diags))) or 'ok', str(page), fonts]))
    with open(out_path, 'w', encoding='utf-8', newline='\n') as f:
        f.write('\n'.join(rows) + '\n')
    print(f'wrote {len(rows) - 1} measurements to {out_path}')


if __name__ == '__main__':
    main()
