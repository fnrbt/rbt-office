#!/usr/bin/env python3
"""Measure the border/fill discriminating fixture from Excel PDF operators.

    python3 tools/excel-parity/measure-print-border-fill.py CELLS.tsv EXCEL_DIR EXCEL_TSV OUT.tsv

Cell geometry comes from the tall solid cell-fill rectangles, which Excel draws from the cell edge in
every border style (verified against the border strip centres). This keeps W, L and R identical
across weights and independent of the border-convention shift. For each boundary the tool also
records the physical border ink:

  strip   narrow (~<4 pt) tall filled rectangle: [left, right], width S; centre
  stroke  vertical stroked path: x, width T from the `w` operator (thin Aptos/Calibri only)

Amount and fill come from Tm/TJ content-stream positions. Fill glyphs are counted in emitted text
order after the matched amount. Output is one row per cell (Excel side only).

Columns: face sheet border index row width align format value cell_left cell_right cell_width
  left_strip_w right_strip_w left_strip_l left_strip_r right_strip_l right_strip_r
  left_stroke_x left_stroke_w right_stroke_x right_stroke_w
  text hash amount amount_x0 amount_x1 amount_gap_left amount_gap_right
  fill_char fill_before fill_after count composite_x1 residual fill_adv
"""
import collections
import importlib.util
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location('numfill', os.path.join(HERE, 'measure-print-numfill.py'))
nf = importlib.util.module_from_spec(spec)
spec.loader.exec_module(nf)

BORDER_ORDER = ['none', 'thin', 'medium', 'thick']


def border_of(sheet):
    for b in BORDER_ORDER:
        if sheet.startswith('B' + b + '_'):
            return b
    return '?'


def tall_fills(pg):
    """Solid cell-fill rectangles (the cell rectangle in every border style)."""
    out = []
    for x, y, w, h in pg['rects']:
        if w > 5 and h > 30:
            out.append((x, y, w, h))
    return out


def narrow_strips(pg):
    out = []
    for x, y, w, h in pg['rects']:
        if 0.05 < w < 4 and h > 30:
            out.append((x, x + w, w))
    return out


def vertical_strokes(pg):
    out = []
    for a, b, c, d, lw in pg.get('ssegs', []):
        if abs(a - c) < 0.01 and abs(b - d) > 10:
            out.append((a, lw))
    return out


def boundaries(fills):
    xs = []
    for x, y, w, h in fills:
        xs.append(round(x, 3))
        xs.append(round(x + w, 3))
    xs = sorted(set(xs))
    out = []
    for x in xs:
        if not out or x - out[-1] > 0.15:
            out.append(x)
    return out


def append(v):
    return '' if v is None else (f'{v:.3f}' if isinstance(v, float) else str(v))


def expected_keys(cells_path):
    expected = set()
    for line in open(cells_path, encoding='utf-8'):
        if line.startswith('sheet\t') or not line.strip():
            continue
        p = line.rstrip('\n').split('\t')
        for w in p[6].split(','):
            expected.add((p[0], p[2], w))
    return expected


def check_key_sets(expected, actual, label):
    """Mandatory completeness gate: missing, extra and duplicate keys all fail."""
    keys = set(actual)
    missing = sorted(expected - keys)
    extra = sorted(keys - expected)
    dup = sorted(k for k, n in actual.items() if n > 1)
    ok = not missing and not extra and not dup
    print(f'{label}: keys expected {len(expected)} measured {len(keys)} '
          f'missing {len(missing)} extra {len(extra)} duplicate {len(dup)} '
          f'{"PASS" if ok else "FAIL"}')
    for kind, items in (('missing', missing), ('extra', extra), ('duplicate', dup)):
        for k in items[:10]:
            print(f'  {kind}: {k[0]} row {k[1]} width {k[2]}')
    return ok


def main():
    cells_path, excel_dir, excel_tsv, out_path = sys.argv[1:5]
    manifest = {}
    for line in open(cells_path, encoding='utf-8'):
        if line.startswith('sheet\t') or not line.strip():
            continue
        sheet, index, row, fname, code, value, widths, align = line.rstrip('\n').split('\t')
        for w in widths.split(','):
            manifest[(sheet, int(row), w)] = (fname, code, float(value), align,
                                              nf.active_fill(code, float(value)), int(index))
    pdfs = nf.read_pdf_table(excel_tsv)
    written = collections.Counter()
    out = open(out_path, 'w', encoding='utf-8')
    out.write('face\tsheet\tborder\tindex\trow\twidth\talign\tformat\tvalue\t'
              'cell_left\tcell_right\tcell_width\t'
              'left_strip_w\tright_strip_w\tleft_strip_l\tleft_strip_r\tright_strip_l\tright_strip_r\t'
              'left_stroke_x\tleft_stroke_w\tright_stroke_x\tright_stroke_w\t'
              'text\thash\tamount\tamount_x0\tamount_x1\tamount_gap_left\tamount_gap_right\t'
              'fill_char\tfill_before\tfill_after\tcount\tcomposite_x1\tresidual\tfill_adv\n')
    for sheet, index, name in pdfs:
        pages = nf.parse_pdf(os.path.join(excel_dir, name))
        # A wide sheet can paginate horizontally. Page 0 carries the row labels; later pages have the
        # same row y positions, so labels transfer by y. Every page carries its own width header.
        shared_labels = None
        for page in pages:
            fills = tall_fills(page)
            if not fills:
                continue
            bnd = boundaries(fills)
            if len(bnd) < 2:
                continue
            strips = narrow_strips(page)
            strokes = vertical_strokes(page)
            page_labels = [r for r in page['runs']
                           if re.match(r'\S+ \S+ (pos|neg|zero)$', r['text'].strip())
                           and r['x'] < bnd[0] - 1.0]
            if page_labels and shared_labels is None:
                shared_labels = page_labels
            if shared_labels is None:
                continue
            if page_labels:
                header_y = max(r['y'] for r in page_labels) + 5.0
            else:
                header_y = max(r['y'] for r in shared_labels) + 5.0
            width_at = {}
            for i in range(len(bnd) - 1):
                l, r = bnd[i], bnd[i + 1]
                label = ''.join(x['text'] for x in page['runs']
                                if x['y'] > header_y and l <= x['x'] < r).strip()
                if re.fullmatch(r'\d+(\.\d+)?', label):
                    width_at[i] = label

            def strip_at(x):
                hit = [s for s in strips if s[0] - 0.4 <= x <= s[1] + 0.6]
                if not hit:
                    return None
                return min(hit, key=lambda s: abs((s[0] + s[1]) / 2.0 - x))

            def stroke_at(x, strip):
                cand = [s for s in strokes if abs(s[0] - x) < 0.6]
                if strip is not None:
                    cand += [s for s in strokes if abs(s[0] - strip[0]) < 0.2 or abs(s[0] - strip[1]) < 0.2]
                return min(cand, key=lambda s: abs(s[0] - x)) if cand else None

            for lab in shared_labels:
                m = re.match(r'(\S+) (\S+) (pos|neg|zero)$', lab['text'].strip())
                if not m or m.group(1) != sheet:
                    continue
                fname, vname = m.group(2), m.group(3)
                value = {'pos': '1234.5', 'neg': '-1234.5', 'zero': '0'}[vname]
                for ci, wlabel in width_at.items():
                    hit = [k for k, v in manifest.items()
                           if k[0] == sheet and k[2] == wlabel and v[0] == fname and v[2] == float(value)]
                    if not hit:
                        continue
                    fname2, code, val2, align, fill, idx = manifest[hit[0]]
                    row_no = hit[0][1]
                    l, r = bnd[ci], bnd[ci + 1]
                    rs = sorted([x for x in page['runs']
                                 if abs(x['y'] - lab['y']) < 0.75
                                 and l - 1.0 <= x['x'] <= r + 0.5], key=lambda x: x['x'])
                    text = ''.join(x['text'] for x in rs)
                    amount, ax0, ax1, _ = nf.find_amount(rs)
                    fb, fa = nf.fill_counts(rs, (amount, ax0, ax1, _) if amount else None, fill)
                    step = nf.fill_step(rs, (amount, ax0, ax1, _) if amount else None, fill)
                    composite = max((b for x in rs for _, _, b in x['chars']), default=None)
                    sl, sr = strip_at(l), strip_at(r)
                    kl, kr = stroke_at(l, sl), stroke_at(r, sr)
                    written[(sheet, str(row_no), wlabel)] += 1
                    out.write('\t'.join(map(append, [
                        os.path.basename(excel_dir).replace('border-', ''),
                        sheet, border_of(sheet), idx, row_no, wlabel, align,
                        fname2, val2, l, r, r - l,
                        None if sl is None else sl[2], None if sr is None else sr[2],
                        None if sl is None else sl[0], None if sl is None else sl[1],
                        None if sr is None else sr[0], None if sr is None else sr[1],
                        None if kl is None else kl[0], None if kl is None else kl[1],
                        None if kr is None else kr[0], None if kr is None else kr[1],
                        text, 'yes' if (amount is None and '#' in text) else 'no',
                        amount or '', ax0, ax1,
                        None if ax0 is None else ax0 - l, None if ax1 is None else r - ax1,
                        fill or '', fb, fa, fb + fa, composite,
                        None if composite is None else r - composite, step,
                    ])) + '\n')
    out.close()
    print(out_path)
    if not check_key_sets(expected_keys(cells_path), written, out_path):
        sys.exit(1)


if __name__ == '__main__':
    main()
