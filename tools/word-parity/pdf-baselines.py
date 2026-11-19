#!/usr/bin/env python3
"""Where Word's PDF puts each line's text and each filled rule, for checking layout rules to 0.12pt
(Word's PDFs place things in 1/600in steps) where Information(6) gives whole pixels.

    python3 tools/word-parity/pdf-baselines.py <pdf> [<pdf> ...] > out.tsv

Rows (y in points from the page's top):
  <doc> TAB text TAB <page> TAB <baseline> TAB <size> TAB <x> TAB <first word>
  <doc> TAB rule TAB <page> TAB <top> TAB <thickness> TAB <x> TAB <width>
The baseline is pdftotext's word box bottom less a quarter of its size (Calibri's descent in the
font descriptor Word writes: exact for Calibri, approximate for other fonts); the size is the box's
height over 1.00364. Rules are the content streams' filled rectangles (`re f`), page by page.
Needs pdftotext (poppler); no Python packages.
"""
import os, re, subprocess, sys, zlib


def text_rows(pdf):
    out = subprocess.run(['pdftotext', '-bbox', pdf, '-'], capture_output=True, text=True, check=True).stdout
    page, seen = 0, set()
    for line in out.splitlines():
        if '<page ' in line:
            page += 1
        m = re.search(r'xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">(.*)</word>', line)
        if not m:
            continue
        x0, y0, _, y1 = map(float, m.groups()[:4])
        size = (y1 - y0) / 1.00364
        baseline = round(y1 - 0.2497 * size, 3)
        if (page, baseline) in seen:
            continue
        seen.add((page, baseline))
        yield page, baseline, round(size, 2), round(x0, 2), m.group(5)


def rule_rows(pdf, height=792.0):
    data = open(pdf, 'rb').read()
    # Page content streams in page order: the /Contents of each /Type /Page object.
    objects = {int(m.group(1)): m.start() for m in re.finditer(rb'(\d+) 0 obj', data)}
    def body(num):
        start = objects[num]
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
    root = re.search(rb'/Type\s*/Catalog.*?/Pages (\d+) 0 R', data, re.S)
    if not root:
        return
    for page, num in enumerate(kids(int(root.group(1))), 1):
        m = re.search(rb'/Contents\s*(\[[^\]]*\]|\d+ 0 R)', body(num))
        if not m:
            continue
        for ref in re.findall(rb'(\d+) 0 R', m.group(1)):
            for r in re.finditer(r'([\d.\-]+) ([\d.\-]+) ([\d.\-]+) ([\d.\-]+) re\s*f\*?', stream(int(ref))):
                x, y, w, h = map(float, r.groups())
                yield page, round(height - y - h, 3), round(h, 3), x, w


for pdf in sys.argv[1:]:
    doc = os.path.splitext(os.path.basename(pdf))[0]
    for page, baseline, size, x, word in text_rows(pdf):
        print(f'{doc}\ttext\t{page}\t{baseline}\t{size}\t{x}\t{word}')
    for page, top, thickness, x, width in rule_rows(pdf):
        print(f'{doc}\trule\t{page}\t{top}\t{thickness}\t{x}\t{width}')
