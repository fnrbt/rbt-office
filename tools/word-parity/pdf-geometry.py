#!/usr/bin/env python3
"""Where a PDF puts each line of text, across and down, and each picture: for checking how text goes
round floating objects (Word's PDFs, or ours from tools/pdf-runner) - see fixtures/float-wrap.py.

    python3 tools/word-parity/pdf-geometry.py <pdf> [<pdf> ...] > out.tsv

Rows (points from the page's top-left):
  <doc> TAB line TAB <page> TAB <top> TAB <bottom> TAB <left> TAB <right> TAB <first word> TAB <last word>
  <doc> TAB image TAB <page> TAB <left> TAB <top> TAB <width> TAB <height>
A line is a run of words pdftotext puts on one baseline with nothing wide between them (a line of
text on each side of a picture is two rows). Pictures: each image XObject drawn (`cm` ... `Do`), with
the transformation in force. Needs pdftotext (poppler); no Python packages.
"""
import os, re, subprocess, sys, zlib


def line_rows(pdf):
    out = subprocess.run(['pdftotext', '-bbox', pdf, '-'], capture_output=True, text=True, check=True).stdout
    page, words = 0, []
    pages = []
    for line in out.splitlines():
        if '<page ' in line:
            page += 1
            pages.append([])
        m = re.search(r'xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">(.*)</word>', line)
        if m:
            x0, y0, x1, y1 = map(float, m.groups()[:4])
            pages[-1].append((y1, x0, y0, x1, m.group(5)))
    for number, ws in enumerate(pages, 1):
        # Words on one baseline (bottoms within 0.5pt), left to right; a gap wider than 12pt (or a word
        # left of the one before: a table's cell beside a line) starts a fragment.
        ws.sort(key=lambda w: (round(w[0] * 2) / 2, w[1]))
        rows = []
        for bottom, x0, top, x1, text in ws:
            last = rows[-1] if rows else None
            if last and abs(last['bottom'] - bottom) <= 0.5 and -1 <= x0 - last['right'] <= 12:
                last['right'] = x1; last['last'] = text; last['top'] = min(last['top'], top)
            else:
                rows.append({'bottom': bottom, 'top': top, 'left': x0, 'right': x1, 'first': text, 'last': text})
        for r in sorted(rows, key=lambda r: (r['bottom'], r['left'])):
            yield number, round(r['top'], 2), round(r['bottom'], 2), round(r['left'], 2), round(r['right'], 2), r['first'], r['last']


def objects(data):
    return {int(m.group(1)): m.start() for m in re.finditer(rb'(\d+) 0 obj', data)}


def image_rows(pdf):
    data = open(pdf, 'rb').read()
    objs = objects(data)
    def body(num):
        start = objs[num]
        return data[start:data.find(b'endobj', start)]
    def stream(num):
        b = body(num)
        s = b.find(b'stream') + len(b'stream')
        s += 2 if b[s:s + 2] == b'\r\n' else 1
        raw = b[s:b.rfind(b'endstream')]
        try:
            return zlib.decompress(raw).decode('latin1')
        except zlib.error:
            return raw.decode('latin1')
    def kids(num):
        b = body(num)
        if re.search(rb'/Type\s*/Page[^s]', b):
            yield num
        m = re.search(rb'/Kids\s*\[([^\]]*)\]', b)
        if m:
            for k in re.findall(rb'(\d+) 0 R', m.group(1)):
                yield from kids(int(k))
    def page_height(num):
        m = re.search(rb'/MediaBox\s*\[\s*([\d.\-]+)\s+([\d.\-]+)\s+([\d.\-]+)\s+([\d.\-]+)', body(num))
        return float(m.group(4)) if m else 792.0
    def xobject_is_image(num, name):
        b = body(num)
        res = re.search(rb'/Resources\s*(\d+) 0 R', b)
        text = body(int(res.group(1))) if res else b
        m = re.search(rb'/' + name.encode() + rb'\s+(\d+) 0 R', text)
        if not m:
            # Resources may be nested under the page tree: search everywhere.
            m = re.search(rb'/' + name.encode() + rb'\s+(\d+) 0 R', data)
        return bool(m) and b'/Image' in body(int(m.group(1)))[:400]
    catalog = next((num for num in objs if re.search(rb'/Type\s*/Catalog', body(num))), None)
    root = catalog and re.search(rb'/Pages\s+(\d+) 0 R', body(catalog))
    if not root:
        return
    token = re.compile(r'\[[^\]]*\]|\([^)]*\)|/[^\s/\[\]()<>]+|[-+]?\d*\.?\d+|[A-Za-z\'"*]+|<<|>>|<[^>]*>')
    for page, num in enumerate(kids(int(root.group(1))), 1):
        height = page_height(num)
        m = re.search(rb'/Contents\s*(\[[^\]]*\]|\d+ 0 R)', body(num))
        if not m:
            continue
        content = "\n".join(stream(int(ref)) for ref in re.findall(rb'(\d+) 0 R', m.group(1)))
        ctm = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
        stack, operands = [], []
        for t in token.findall(content):
            if re.fullmatch(r'[-+]?\d*\.?\d+', t) or t.startswith('/') or t.startswith('[') or t.startswith('(') or t.startswith('<'):
                operands.append(t)
                continue
            if t == 'q':
                stack.append(ctm)
            elif t == 'Q':
                ctm = stack.pop() if stack else ctm
            elif t == 'cm' and len(operands) >= 6:
                a, b, c, d, e, f = map(float, operands[-6:])
                A, B, C, D, E, F = ctm
                ctm = (a * A + b * C, a * B + b * D, c * A + d * C, c * B + d * D, e * A + f * C + E, e * B + f * D + F)
            elif t == 'Do' and operands and xobject_is_image(num, operands[-1][1:]):
                a, b, c, d, e, f = ctm
                xs = [e, e + a, e + c, e + a + c]
                ys = [f, f + b, f + d, f + b + d]
                yield page, round(min(xs), 2), round(height - max(ys), 2), round(max(xs) - min(xs), 2), round(max(ys) - min(ys), 2)
            operands = []


for pdf in sys.argv[1:]:
    doc = os.path.splitext(os.path.basename(pdf))[0]
    for page, top, bottom, left, right, first, last in line_rows(pdf):
        print(f'{doc}\tline\t{page}\t{top}\t{bottom}\t{left}\t{right}\t{first}\t{last}')
    for page, left, top, width, height in image_rows(pdf):
        print(f'{doc}\timage\t{page}\t{left}\t{top}\t{width}\t{height}')
