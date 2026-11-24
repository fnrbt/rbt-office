#!/usr/bin/env python3
"""Measure numeric-fill cells from PDF text operators against their printed cell borders.

    python3 tools/excel-parity/measure-print-numfill.py CELLS.tsv EXCEL_DIR EXCEL_TSV OURS_DIR OUT.tsv

EXCEL_DIR holds the fresh direct Excel 2019 PDFs (one per sheet, as Excel's PDF export produces);
EXCEL_TSV is excel-print-pages.sh's sheet/index/pdf table. OURS_DIR holds ours-<index>.pdf from
tools/spreadsheet-pdf-runner or the Chromium proof. Both sides are measured from the content stream:
the text matrix / TJ operand positions of the amount and the `$` sign, not from Poppler word boxes
(which drop fills and whitespace). Cell edges are the thin printed border centrelines. Gaps are
cell-relative, so page origins and the native/browser x shift cancel.

Each row of the long table is one test cell and side:
  sheet index row width format_name value side cell_left cell_right text hash
  amount amount_x0 amount_x1 amount_gap_left amount_gap_right
  dollar dollar_x0 dollar_x1 dollar_gap_left fill_char fill_before fill_after count_steps
A missing amount (a printed hash run) leaves the amount fields empty. `fill_before`/`fill_after`
count the format's active fill character that the operators place left/right of the amount.
"""
import html, os, re, sys, zlib
from collections import OrderedDict

# ---------- PDF object / content parser ----------

def load_objects(data):
    return {int(m.group(1)): m.group(2)
            for m in re.finditer(rb'(?m)^(\d+) 0 obj\s*(.*?)endobj', data, re.S)}


def stream_of(body):
    m = re.search(rb'stream\r?\n(.*?)\r?\nendstream', body, re.S)
    if not m:
        return b''
    raw = m.group(1)
    if b'/FlateDecode' in body:
        try:
            return zlib.decompress(raw)
        except Exception:
            return b''
    return raw


def basefont(body):
    m = re.search(rb'/BaseFont\s*/([^\s/<>]+)', body)
    return m.group(1).decode('latin1') if m else '?'


def parse_fonts(objs, page):
    fonts = {}
    rm = re.search(rb'/Font\s*<<(.*?)>>', page, re.S)
    if not rm:
        return fonts
    for fm in re.finditer(rb'/(\w+)\s+(\d+) 0 R', rm.group(1)):
        name, fo = fm.group(1).decode(), int(fm.group(2))
        fb = objs.get(fo, b'')
        if b'/Type0' in fb:
            dm = re.search(rb'/DescendantFonts\s*\[\s*(\d+) 0 R', fb)
            db = objs.get(int(dm.group(1)) if dm else fo, b'')
            widths = {}
            wm = re.search(rb'/W\s*\[(.*?)\]\s*(?:/|>>)', db, re.S)
            if wm:
                ws = wm.group(1).decode('latin1')
                for mm in re.finditer(r'(\d+)\s*\[([^\]]*)\]|(\d+)\s+(\d+)\s+(-?\d+)', ws):
                    if mm.group(1):
                        c = int(mm.group(1))
                        for k, w in enumerate(int(x) for x in re.findall(r'-?\d+', mm.group(2))):
                            widths[c + k] = w
                    else:
                        c, d, w = int(mm.group(3)), int(mm.group(4)), int(mm.group(5))
                        for k in range(c, d + 1):
                            widths[k] = w
            tu = {}
            tm = re.search(rb'/ToUnicode\s+(\d+) 0 R', fb)
            if tm:
                tube = stream_of(objs.get(int(tm.group(1)), b''))
                for bm in re.finditer(rb'<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>', tube):
                    try:
                        tu[int(bm.group(1), 16)] = bytes.fromhex(bm.group(2).decode()).decode('utf-16-be', 'ignore')
                    except Exception:
                        pass
            fonts[name] = {'kind': 'type0', 'base': basefont(db), 'widths': widths, 'tu': tu}
        else:
            fc = re.search(rb'/FirstChar\s+(\d+)', fb)
            widths = {}
            wm = re.search(rb'/Widths\s+(\d+) 0 R', fb)
            if wm:
                wb = objs.get(int(wm.group(1)), b'')
                if b'[' in wb:
                    arr = [int(x) for x in re.findall(rb'\d+', wb.split(b'[', 1)[1].split(b']')[0])]
                    first = int(fc.group(1)) if fc else 0
                    for i, w in enumerate(arr):
                        widths[first + i] = w
            fonts[name] = {'kind': 'simple', 'base': basefont(fb), 'widths': widths}
    return fonts


def pdf_unescape(s):
    out, i = [], 0
    while i < len(s):
        c = s[i]
        if c == '\\' and i + 1 < len(s):
            n = s[i + 1]
            mp = {'n': '\n', 'r': '\r', 't': '\t', 'b': '\b', 'f': '\f', '(': '(', ')': ')', '\\': '\\'}
            if n in mp:
                out.append(mp[n]); i += 2
            elif n.isdigit():
                j = i + 1; oct = ''
                while j < len(s) and len(oct) < 3 and s[j] in '01234567':
                    oct += s[j]; j += 1
                out.append(chr(int(oct, 8))); i = j
            elif n == '\n':
                i += 2
            else:
                out.append(n); i += 2
        else:
            out.append(c); i += 1
    return ''.join(out)


def shown_tokens(token, font):
    """[(glyph code, unicode text)] for a PDF literal/hex shown string."""
    if token.startswith('<'):
        hx = re.sub(r'\s', '', token[1:-1])
        if font and font['kind'] == 'type0':
            return [(int(hx[i:i + 4], 16), font['tu'].get(int(hx[i:i + 4], 16), ''))
                    for i in range(0, len(hx), 4)]
        raw = bytes.fromhex(hx).decode('latin1')
        return [(ord(c), c) for c in raw]
    raw = pdf_unescape(token[1:-1])
    if font and font['kind'] == 'type0':
        return [(ord(c), font['tu'].get(ord(c), '')) for c in raw]
    return [(ord(c), c) for c in raw]



def line_tokens(line):
    toks, i = [], 0
    while i < len(line):
        c = line[i]
        if c.isspace():
            i += 1; continue
        if c == '(':
            j, depth = i + 1, 1
            while j < len(line) and depth:
                if line[j] == '\\':
                    j += 2; continue
                if line[j] == '(':
                    depth += 1
                elif line[j] == ')':
                    depth -= 1
                    if depth == 0:
                        j += 1; break
                j += 1
            toks.append(line[i:j]); i = j
        elif c == '<':
            j = line.find('>', i + 1)
            j = len(line) if j < 0 else j + 1
            toks.append(line[i:j]); i = j
        elif c == '[':
            j, depth, quoted = i + 1, 1, False
            while j < len(line) and depth:
                ch = line[j]
                if quoted:
                    if ch == '\\':
                        j += 2; continue
                    if ch == ')':
                        quoted = False
                elif ch == '(':
                    quoted = True
                elif ch == '[':
                    depth += 1
                elif ch == ']':
                    depth -= 1
                    if depth == 0:
                        j += 1; break
                j += 1
            toks.append(line[i:j]); i = j
        else:
            j = i
            while j < len(line) and not line[j].isspace() and line[j] not in '([<':
                j += 1
            if j == i:
                j += 1
            toks.append(line[i:j]); i = j
    return toks


def array_pieces(arr):
    inner = arr[1:-1]
    pieces, i = [], 0
    while i < len(inner):
        c = inner[i]
        if c.isspace():
            i += 1
        elif c == '(':
            j, depth = i + 1, 1
            while j < len(inner) and depth:
                if inner[j] == '\\':
                    j += 2; continue
                if inner[j] == '(':
                    depth += 1
                elif inner[j] == ')':
                    depth -= 1
                j += 1
            pieces.append(inner[i:j]); i = j
        elif c == '<':
            j = inner.find('>', i + 1) + 1
            pieces.append(inner[i:j]); i = j
        else:
            j = i
            while j < len(inner) and not inner[j].isspace() and inner[j] not in '(<[':
                j += 1
            if j == i:
                j += 1
            pieces.append(inner[i:j]); i = j
    return pieces


def parse_pdf(path):
    data = open(path, 'rb').read()
    objs = load_objects(data)
    pages = []
    for n, b in objs.items():
        if re.search(rb'/Type\s*/Page\b', b) and b'/Contents' in b:
            pages.append((n, b))
    pages.sort()
    result = []
    for pn, pb in pages:
        fonts = parse_fonts(objs, pb)
        cm = re.search(rb'/Contents\s+(\d+) 0 R', pb)
        if cm:
            cids = [int(cm.group(1))]
        else:
            cm = re.search(rb'/Contents\s*\[(.*?)\]', pb, re.S)
            cids = [int(x) for x in re.findall(rb'(\d+) 0 R', cm.group(1))] if cm else []
        content = '\n'.join(stream_of(objs.get(c, b'')).decode('latin1') for c in cids)
        runs, segs, rects, pending = [], [], [], []
        font = None; fs = 0.0; tm = (1, 0, 0, 1, 0, 0); charsp = 0.0; wordsp = 0.0
        cur_m = None
        for line in content.split('\n'):
            line = line.strip()
            if not line:
                continue
            toks = line_tokens(line)
            if not toks:
                continue
            if re.fullmatch(r"[A-Za-z*']+", toks[-1]) and not (len(toks) == 1 and toks[0].startswith('/')):
                op, args = toks[-1], toks[:-1]
            else:
                continue
            if op == 'Tf' and len(args) >= 2:
                font = fonts.get(args[-2].lstrip('/')); fs = float(args[-1])
            elif op == 'Tc' and args:
                try: charsp = float(args[-1])
                except ValueError: pass
            elif op == 'Tw' and args:
                try: wordsp = float(args[-1])
                except ValueError: pass
            elif op == 'Tm' and len(args) >= 6:
                tm = tuple(float(v) for v in args[-6:])
            elif op == 're' and len(args) >= 4:
                try:
                    pending.append((float(args[-4]), float(args[-3]), float(args[-2]), float(args[-1])))
                except ValueError:
                    pass
            elif op in ('f', 'f*'):
                rects.extend(pending); pending = []
            elif op in ('n', 'W', 'W*'):
                pending = []
            elif op == 'm' and len(args) >= 2:
                try: cur_m = (float(args[-2]), float(args[-1]))
                except ValueError: cur_m = None
            elif op == 'l' and len(args) >= 2 and cur_m is not None:
                try:
                    segs.append((cur_m[0], cur_m[1], float(args[-2]), float(args[-1])))
                except ValueError:
                    pass
            elif op in ('TJ', 'Tj'):
                if op == 'Tj':
                    pieces = [next((a for a in reversed(args) if a.startswith(('(', '<'))), args[-1])]
                else:
                    arr = next((a for a in reversed(args) if a.startswith('[')), '[]')
                    pieces = array_pieces(arr)
                x = tm[4]; start = x; text = ''; chars = []
                for p in pieces:
                    if p.startswith('(') or p.startswith('<'):
                        toks = shown_tokens(p, font)
                        text += ''.join(ch for _, ch in toks)
                        for g, ch in toks:
                            gw = font['widths'].get(g, 0) if font else 0
                            adv = gw / 1000.0 * fs + charsp + (wordsp if ch == ' ' else 0)
                            chars.append((ch, x, x + adv)); x += adv
                    else:
                        try: x -= float(p) / 1000.0 * fs
                        except ValueError: pass
                runs.append({'x': start, 'y': tm[5], 'end': x, 'text': text,
                             'chars': chars, 'font': font['base'] if font else '?', 'size': fs})
        result.append({'page': pn, 'runs': runs, 'segs': segs, 'rects': rects})
    return result



def split_sections(code):
    out, cur, i, quoted = [], [], 0, False
    while i < len(code):
        c = code[i]
        if quoted:
            cur.append(c)
            if c == '"':
                quoted = False
        elif c == '"':
            quoted = True; cur.append(c)
        elif c == '\\' and i + 1 < len(code):
            cur.append(c); cur.append(code[i + 1]); i += 1
        elif c == '[':
            j = code.find(']', i + 1)
            if j > i:
                cur.append(code[i:j + 1]); i = j
            else:
                cur.append(c)
        elif c == ';':
            out.append(''.join(cur)); cur = []
        else:
            cur.append(c)
        i += 1
    out.append(''.join(cur))
    return out


def active_section(code, value):
    sections = split_sections(code)
    if len(sections) >= 3 and value == 0:
        return sections[2] if sections[2] else sections[0]
    if len(sections) >= 2 and value < 0:
        return sections[1] if sections[1] else sections[0]
    return sections[0] if sections else code


def active_fill(code, value):
    """The active section's unquoted/unescaped `*x` fill character, or ''."""
    sec = active_section(code, value)
    i, quoted = 0, False
    while i < len(sec):
        c = sec[i]
        if quoted:
            if c == '"':
                quoted = False
            i += 1
        elif c == '"':
            quoted = True; i += 1
        elif c == '\\' and i + 1 < len(sec):
            i += 2
        elif c == '[':
            j = sec.find(']', i + 1)
            i = j + 1 if j > i else i + 1
        elif c == '*' and i + 1 < len(sec):
            return sec[i + 1]
        else:
            i += 1
    return ''


# ---------- cell geometry and measurements ----------

def grid(pg):
    verts = sorted({round(s[0], 2) for s in pg['segs'] if abs(s[0] - s[2]) < 0.01})
    hors = sorted({round(s[1], 2) for s in pg['segs'] if abs(s[1] - s[3]) < 0.01}, reverse=True)
    return verts, hors


def runs_in_row(pg, top, bottom):
    ys = [r['y'] for r in pg['runs'] if bottom - 0.5 <= r['y'] <= top + 0.5]
    if not ys:
        return []
    y = round(sorted(ys)[len(ys) // 2], 2)
    return [r for r in pg['runs'] if abs(r['y'] - y) < 0.75]


def runs_in_cell(pg, left, right, top, bottom):
    rs = runs_in_row(pg, top, bottom)
    return sorted([r for r in rs if r['x'] >= left - 1.0 and r['x'] <= right + 0.5], key=lambda r: r['x'])


def cell_text(rs):
    return ''.join(r['text'] for r in rs)


def find_amount(rs):
    """Return (text, x0, x1, run) for the leftmost formatted amount."""
    for r in sorted(rs, key=lambda r: r['x']):
        m = re.search(r'-?[\d,]+\.\d+', r['text'])
        if not m:
            continue
        chars = r['chars']
        # account for any literal prefix in the same run
        idx = m.start()
        return m.group(0), chars[idx][1], chars[m.end() - 1][2], r
    return None, None, None, None


def find_dollar(rs):
    xs = []
    for r in rs:
        for ch, a, b in r['chars']:
            if ch == '$':
                xs.append((a, b))
    return (min(a for a, _ in xs), max(b for _, b in xs)) if xs else (None, None)


def fill_counts(rs, amount, fill_char):
    """Count fill glyphs before/after the amount in emitted text order.

    Glyph origins need not sit inside the previous glyph's advance box: Excel's first trailing fill
    glyph commonly starts a few hundredths of a point before the number's /Widths advance end (and
    the last one may overlap the next glyph's origin). An x threshold therefore drops that glyph and
    undercounts by one. The emitted text order is unambiguous: amount match first, fill chars after.
    """
    if not fill_char or amount is None:
        return 0, 0
    amount = amount[0]
    if not amount:
        return 0, 0
    text = ''.join(r['text'] for r in rs)
    i = text.find(amount)
    if i < 0:
        return 0, 0
    return text.count(fill_char, 0, i), text.count(fill_char, i + len(amount))


def fill_step(rs, amount, fill_char):
    """Median origin-to-origin advance of the trailing fill glyphs, else one glyph advance."""
    if not fill_char or amount is None:
        return None
    amount = amount[0]
    if not amount:
        return None
    text = ''.join(r['text'] for r in rs)
    i = text.find(amount)
    if i < 0:
        return None
    flat = [(ch, a, b) for r in rs for ch, a, b in r['chars']]
    ordered = flat
    start = i + len(amount)
    tail = [b - a for ch, a, b in ordered[start:] if ch == fill_char]
    if not tail:
        return None
    origins = [a for ch, a, b in ordered[start:] if ch == fill_char]
    steps = sorted(round(y - x, 4) for x, y in zip(origins, origins[1:]) if y - x > 0.01)
    if steps:
        return steps[len(steps) // 2]
    return round(tail[0], 4)


def measure_side(pdfs, cells):
    """Measure one side. pdfs: [(sheet,index,path)]; cells: manifest rows for those sheets."""
    by_sheet = OrderedDict()
    for c in cells:
        by_sheet.setdefault(c[0], []).append(c)
    measured = {}
    for sheet, index, path in pdfs:
        want = by_sheet.get(sheet, [])
        if not want:
            continue
        for pg in parse_pdf(path):
            verts, hors = grid(pg)
            # Georgia/Tahoma direct exports omit the thin border path; their tall solid fills still
            # give the physical column rectangles, so use those when no border centrelines were drawn.
            if len(verts) < 2:
                fills = sorted({(round(r[0], 2), round(r[2], 2)) for r in pg['rects']
                                if r[2] > 10 and r[3] > 15})
                if len(fills) >= 2:
                    verts = [f[0] for f in fills] + [fills[-1][0] + fills[-1][1]]
                    # adjacent inflated fills overlap; use each fill's left edge and the next fill's left edge
                    verts = sorted(set(round(v, 2) for v in verts))
            if len(verts) < 2:
                continue
            # rows are identified by the A-column text, not by horizontal rules (which may be absent).
            label_runs = [r for r in pg['runs'] if r['x'] < verts[0] - 1.0]
            row_labels = []
            for r in label_runs:
                if re.match(r'\S+ \S+ (pos|neg|zero)$', r['text'].strip()):
                    row_labels.append(r)
            if not row_labels:
                continue
            header_y = max(r['y'] for r in row_labels) + 5.0
            header_runs = [r for r in pg['runs'] if r['y'] > header_y]
            width_at = {}
            for i in range(len(verts) - 1):
                l, r = verts[i], verts[i + 1]
                label = ''.join(x['text'] for x in header_runs if l <= x['x'] < r).strip()
                if re.fullmatch(r'\d+(\.\d+)?', label):
                    width_at[i] = label
            for lab in row_labels:
                m = re.match(r'(\S+) (\S+) (pos|neg|zero)$', lab['text'].strip())
                if not m or m.group(1) != sheet:
                    continue
                fname, vname = m.group(2), m.group(3)
                value = {'pos': '1234.5', 'neg': '-1234.5', 'zero': '0'}[vname]
                row = next((c[1] for c in want
                            if c[3] == fname and abs(float(c[5]) - float(value)) < 1e-9), None)
                if row is None:
                    continue
                for ci, wlabel in width_at.items():
                    if ci + 1 >= len(verts):
                        continue
                    left, right = verts[ci], verts[ci + 1]
                    rs = sorted([r for r in pg['runs']
                                 if abs(r['y'] - lab['y']) < 0.75
                                 and r['x'] >= left - 1.0 and r['x'] <= right + 0.5], key=lambda r: r['x'])
                    text = ''.join(r['text'] for r in rs)
                    amount, ax0, ax1, _ = find_amount(rs)
                    dx0, dx1 = find_dollar(rs)
                    fill = next((c[7] for c in want if c[1] == row and c[2] == wlabel), '')
                    fb, fa = fill_counts(rs, (amount, ax0, ax1, _) if amount else None, fill)
                    fstep = fill_step(rs, (amount, ax0, ax1, _) if amount else None, fill)
                    composite = max((b for r in rs for _, _, b in r['chars']), default=None)
                    measured[(sheet, row, wlabel)] = {
                        'cell_left': left, 'cell_right': right, 'text': text,
                        'hash': bool(amount is None and '#' in text),
                        'amount': amount or '', 'amount_x0': ax0, 'amount_x1': ax1,
                        'amount_gap_left': None if ax0 is None else ax0 - left,
                        'amount_gap_right': None if ax1 is None else right - ax1,
                        'dollar': '$' if dx0 is not None else '', 'dollar_x0': dx0, 'dollar_x1': dx1,
                        'dollar_gap_left': None if dx0 is None else dx0 - left,
                        'fill_char': fill or '', 'fill_before': fb, 'fill_after': fa,
                        'composite_x1': composite, 'fill_adv': fstep,
                    }
    return measured


def read_pdf_table(excel_tsv):
    out = []
    for line in open(excel_tsv, encoding='utf-8'):
        p = line.rstrip('\n').split('\t')
        if len(p) >= 4 and p[2] == 'pdf' and p[3].endswith('.pdf'):
            out.append((p[0], int(p[1]), p[3]))
    return out


def main():
    cells_path, excel_dir, excel_tsv, ours_dir, out_path = sys.argv[1:6]
    manifest = []
    for line in open(cells_path, encoding='utf-8'):
        if line.startswith('sheet\t') or not line.strip():
            continue
        sheet, index, row, fname, code, value, widths, align = line.rstrip('\n').split('\t')
        for w in widths.split(','):
            manifest.append((sheet, int(row), w, fname, code, float(value), align,
                             active_fill(code, float(value))))
    pdfs = read_pdf_table(excel_tsv)
    excel = measure_side([(s, i, os.path.join(excel_dir, name)) for s, i, name in pdfs], manifest)
    ours_paths = [(s, i, os.path.join(ours_dir, f'ours-{i}.pdf')) for s, i, _ in pdfs]
    ours = measure_side([p for p in ours_paths if os.path.exists(p[2])], manifest)
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write('sheet\tindex\trow\twidth\tformat_name\tformat\tvalue\talign\tside\tcell_left\tcell_right\t'
                'text\thash\tamount\tamount_x0\tamount_x1\tamount_gap_left\tamount_gap_right\t'
                'dollar\tdollar_x0\tdollar_x1\tdollar_gap_left\tfill_char\tfill_before\tfill_after\t'
                'composite_x1\tfill_adv\n')
        for sheet, index, _ in pdfs:
            for c in [x for x in manifest if x[0] == sheet]:
                sname, row, width, fname, code, value, align, fill = c
                for side, table in (('excel', excel), ('ours', ours)):
                    m = table.get((sheet, row, width))
                    if not m:
                        continue
                    def fmt(v):
                        return '' if v is None else f'{v:.3f}'
                    f.write('\t'.join(map(str, [
                        sheet, index, row, width, fname, code, value, align, side,
                        fmt(m['cell_left']), fmt(m['cell_right']), m['text'], 'yes' if m['hash'] else 'no',
                        m['amount'], fmt(m['amount_x0']), fmt(m['amount_x1']),
                        fmt(m['amount_gap_left']), fmt(m['amount_gap_right']),
                        m['dollar'], fmt(m['dollar_x0']), fmt(m['dollar_x1']), fmt(m['dollar_gap_left']),
                        m['fill_char'], m['fill_before'], m['fill_after'],
                        fmt(m['composite_x1']), fmt(m['fill_adv'])])) + '\n')
    print(out_path)


if __name__ == '__main__':
    main()
