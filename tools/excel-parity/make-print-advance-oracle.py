#!/usr/bin/env python3
"""Builds the printed-advance slope/intercept oracle workbooks.

Purpose: separate a *printed text advance* (the pen movement from the PDF
content stream, which is also exactly what Poppler's word box measures) from the
shipping model's integer-dot glyph advance, and from the physical cell width W.
Each sheet is one face/size/print-scale combination:

  * wide cells (28 characters) hold repeated same digits at varying lengths,
    punctuation strings, numeric percent values and their literal twins, so the
    advance can be regressed on length: slope = per-digit advance, intercept =
    the constant suffix/punctuation part;
  * tight cells sweep column widths across the full->hash cutoff at 1/32
    character around the model-predicted transition, so the cutoff can be
    verified from the PDF fill rectangle without the model.

`Verdana 13` upright at 75 % and 100 % are the measured pair.  Georgia 15
upright at 90 % is included only to separate its physical column-width residual
from the advance residual.  The holdout workbook reserves an unseen scale
(Verdana 13 at 140 %) and an unseen face (Tahoma 15 at 90 %); its predictions
are locked before that workbook is exported.

Cells carry unique solid fills (same palette as the other print fixtures), so a
column's physical edges are read from the PDF page.

    python3 tools/excel-parity/make-print-advance-oracle.py OUT.xlsx META.tsv train|holdout
"""
import importlib.util
import math
import os
import struct
import sys
import zipfile
from xml.sax.saxutils import escape

_here = os.path.dirname(os.path.abspath(__file__))
_spec = importlib.util.spec_from_file_location('countfixture', os.path.join(_here, 'make-print-count-fixture.py'))
countfixture = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(countfixture)

DECL = countfixture.DECL
MAIN = countfixture.MAIN
REL = countfixture.REL
PKG = countfixture.PKG
CT = countfixture.CT
fill_color = countfixture.fill_color
col_letter = countfixture.col_letter
attr = countfixture.attr

ROOT = os.path.dirname(os.path.dirname(_here))
FONT_FILES = {
    ('Verdana', False): 'artifacts/fonts/windows/verdana.ttf',
    ('Georgia', False): 'artifacts/fonts/windows-extra/georgia.ttf',
    ('Tahoma', False): 'artifacts/fonts/windows-extra/tahoma.ttf',
}
FMT_ID = {'0.00%': 166, '#,##0.00': 167}
WIDE = 28.0
STEP = 0.03125

# name, face, size, italic, scale, format, split
SHEETS = [
    ('V13up_pct75', 'Verdana', 13, False, 75, '0.00%', 'train'),
    ('V13up_pct100', 'Verdana', 13, False, 100, '0.00%', 'train'),
    ('G15up_grp90', 'Georgia', 15, False, 90, '#,##0.00', 'train'),
    ('H_V13up_pct140', 'Verdana', 13, False, 140, '0.00%', 'holdout'),
    ('H_Tahoma15up_pct90', 'Tahoma', 15, False, 90, '0.00%', 'holdout'),
]

# Per-sheet wide strings.  ('num', value) uses the sheet format; ('lit', text).
def wide_columns(fmt, face):
    cols = []
    if fmt == '0.00%':
        for n in (1, 2, 3, 4, 6, 8, 10, 12):
            cols.append(('num', int('1' * n) / 100))
            cols.append(('num', int('8' * n) / 100))
        for v in (12345.6789, 11111.11, 88888.88, 9999.99, 123.4567, 1.2345, 1234.5678):
            cols.append(('num', v))
        for t in ('1234567.89%', '1111111.00%', '8888888.00%', '999999.00%',
                  '12345.67%', '123.45%', '123456.78%'):
            cols.append(('lit', t))
        for n in (4, 8, 12):
            for d in '18':
                cols.append(('lit', d * n))
                cols.append(('lit', d * n + '.00%'))
        for t in ('1111111', '111.111', '1,111,111', '-1111111', '(1111111)',
                  '1111111.89%', '1111111%', '111%111', '1.1.1.1.1', '$1111111',
                  '111,111.11%', '1111111-', '---1111111', '1,1,1,1,1,1,1'):
            cols.append(('lit', t))
    else:  # #,##0.00
        for v in (1234567, 1111111.11, 8888888.88, 111111.11, 888888.88,
                  11111.11, 88888.88, 1111.11, 8888.88, -1234567):
            cols.append(('num', v))
        for t in ('1,234,567.00', '1,111,111.11', '8,888,888.88',
                  '111,111.11', '888,888.88', '11,111.11', '88,888.88'):
            cols.append(('lit', t))
        for t in ('1,234,567', '1234567', '-1,234,567', '(1,234,567)',
                  '$1,234,567', '1,234,567%', '1,2,3,4,5,6,7'):
            cols.append(('lit', t))
    return cols


# Tight sweeps: (kind, value or text, display text used to centre the window).
def dense_columns(fmt):
    if fmt == '0.00%':
        return [('num', 12345.6789), ('num', 11111111.11)]
    return [('num', 1234567), ('num', 1111111.11)]


def fmt_text(v, fmt):
    if fmt == '0.00%':
        return f'{v * 100:.2f}%'
    return f'{v:,.2f}'


# ---- font metrics: same integer-dot model the shipping engine uses ----
def u16(d, o):
    return struct.unpack_from('>H', d, o)[0]


def font_dir(path):
    d = open(path, 'rb').read()
    base = 0
    if d[:4] == b'ttcf':
        base = struct.unpack_from('>I', d, 12)[0]
    n = u16(d, base + 4)
    t = {d[base + 12 + 16 * i:base + 16 + 16 * i].decode('latin1'):
         struct.unpack_from('>II', d, base + 20 + 16 * i) for i in range(n)}
    return d, t


def glyph_ems(path):
    d, t = font_dir(path)
    upm = u16(d, t['head'][0] + 18)
    hhea = t['hhea'][0]
    nhm = u16(d, hhea + 34)
    hm = t['hmtx'][0]
    cmap = t['cmap'][0]
    sub = None
    for i in range(u16(d, cmap + 2)):
        pid, eid = u16(d, cmap + 4 + 8 * i), u16(d, cmap + 6 + 8 * i)
        so = struct.unpack_from('>I', d, cmap + 8 + 8 * i)[0]
        if u16(d, cmap + so) == 4 and (pid, eid) in ((3, 1), (0, 3), (0, 4)):
            sub = cmap + so
            break
    seg = u16(d, sub + 6) // 2
    end = sub + 14
    start = end + 2 * seg + 2
    delta = start + 2 * seg
    ro = delta + 2 * seg

    def lookup(code):
        for i in range(seg):
            s, e = u16(d, start + 2 * i), u16(d, end + 2 * i)
            if s <= code <= e:
                de = struct.unpack_from('>h', d, delta + 2 * i)[0]
                r = u16(d, ro + 2 * i)
                if r == 0:
                    return (code + de) & 0xFFFF
                g = u16(d, ro + 2 * i + r + 2 * (code - s))
                return (g + de) & 0xFFFF if g else 0
        return 0

    out = {}
    for ch in '$(),.% 0123456789#-':
        g = lookup(ord(ch))
        if g:
            out[ch] = u16(d, hm + 4 * min(g, nhm - 1)) / upm
    return out


def physical_width(w, scale):
    screen_px = int(((256.0 * w + 128 // 7) / 256.0) * 7.0)
    return int(screen_px * 47.0 / 7.0 + 0.5) * 72.0 / 600.0 * scale


def model_text(text, glyphs, size, scale, exact_height):
    h = size * scale * 600.0 / 72.0
    dots = h if exact_height else math.floor(h + 0.5)
    return sum(math.floor(glyphs[c] * dots + 0.5) * 72.0 / 600.0 for c in text)


def dense_widths(text, face, size, scale, exact_height, half=0.5):
    glyphs = glyph_ems(os.path.join(ROOT, FONT_FILES[(face, False)]))
    c = glyphs['0'] * size * scale / 2.0 + 0.875
    text_pt = model_text(text, glyphs, size, scale, exact_height)
    target = text_pt + c
    lo, hi = 1.0, 60.0
    for _ in range(60):
        mid = (lo + hi) / 2.0
        if physical_width(mid, scale) < target:
            lo = mid
        else:
            hi = mid
    center = (lo + hi) / 2.0
    n = int(round(2 * half / STEP))
    return [round(center - half + i * STEP, 4) for i in range(n + 1)]


def build(out, meta_path, split):
    xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
           '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">'
           '<alignment horizontal="left"/></xf>']
    fonts = ['<sz val="11"/><name val="Calibri"/><family val="2"/>']
    used_fmts = sorted({FMT_ID[fmt] for _, _, _, _, _, fmt, sp in SHEETS if sp == split})
    numfmts = ''.join(f'<numFmt numFmtId="{FMT_ID[code]}" formatCode="{attr(code)}"/>'
                      for code in FMT_ID if FMT_ID[code] in used_fmts)
    meta = ['sheet\tcolumn\trow\twidth_chars\talign\tface\tsize_pt\titalic\tscale\tkind\trole\tformat\tvalue\ttext']
    sheet_files = []
    for name, face, size, italic, scale, fmt, sp in SHEETS:
        if sp != split:
            continue
        font_id = len(fonts)
        fonts.append(f'{"<i/>" if italic else ""}<sz val="{size}"/><name val="{attr(face)}"/><family val="2"/>')
        cols = []  # (width, kind, role, value, text)
        for role, v in wide_columns(fmt, face):
            if role == 'num':
                cols.append((WIDE, 'wide', 'num', str(v), fmt_text(v, fmt)))
            else:
                cols.append((WIDE, 'wide', 'lit', '', v))
        for role, v in dense_columns(fmt):
            text = fmt_text(v, fmt)
            for w in dense_widths(text, face, size, scale / 100.0, exact_height=True):
                cols.append((w, 'dense', role, str(v), text))
        col_xml = ''.join(f'<col min="{2 + j}" max="{2 + j}" width="{w}" customWidth="1"/>'
                          for j, (w, *_rest) in enumerate(cols))
        style_of, control_of = [], []
        for j, (w, kind, role, v, text) in enumerate(cols):
            style_of.append(len(xfs))
            xfs.append(f'<xf numFmtId="{FMT_ID[fmt]}" fontId="{font_id}" fillId="{2 + j}" borderId="0" '
                       f'xfId="0" applyFont="1" applyFill="1" applyNumberFormat="1" applyAlignment="1">'
                       f'<alignment horizontal="right"/></xf>')
            control_of.append(len(xfs))
            xfs.append(f'<xf numFmtId="2" fontId="{font_id}" fillId="{2 + j}" borderId="0" '
                       f'xfId="0" applyFont="1" applyFill="1" applyNumberFormat="1" applyAlignment="1">'
                       f'<alignment horizontal="right"/></xf>')
        cells = [f'<row r="1"><c r="A1" s="1" t="inlineStr"><is><t>{escape(name)}</t></is></c>']
        for j, (w, kind, role, v, text) in enumerate(cols):
            ref = f'{col_letter(2 + j)}1'
            if role == 'lit':
                cells.append(f'<c r="{ref}" s="{style_of[j]}" t="inlineStr"><is><t>{escape(text)}</t></is></c>')
                value = ''
            else:
                cells.append(f'<c r="{ref}" s="{style_of[j]}"><v>{v}</v></c>')
                value = v
            meta.append('\t'.join([name, str(2 + j), '1', f'{w:g}', 'right', face, str(size),
                                   '1' if italic else '0', str(scale), kind, role, fmt, value, text]))
        cells.append('</row>')
        cells.append(f'<row r="2"><c r="A2" s="1" t="inlineStr"><is><t>control</t></is></c>')
        for j, (w, kind, role, v, text) in enumerate(cols):
            cells.append(f'<c r="{col_letter(2 + j)}2" s="{control_of[j]}"><v>-98.41</v></c>')
            meta.append('\t'.join([name, str(2 + j), '2', f'{w:g}', 'right', face, str(size),
                                   '1' if italic else '0', str(scale), 'control', 'num', '0.00', '-98.41', '-98.41']))
        cells.append('</row>')
        xml = DECL + (
            f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}">'
            f'<sheetViews><sheetView{" tabSelected=\"1\"" if not sheet_files else ""} workbookViewId="0"/></sheetViews>'
            '<sheetFormatPr defaultRowHeight="15"/>'
            f'<cols><col min="1" max="1" width="24" customWidth="1"/>{col_xml}</cols>'
            f'<sheetData>{"".join(cells)}</sheetData>'
            '<pageMargins left="0.4" right="0.4" top="0.4" bottom="0.4" header="0.2" footer="0.2"/>'
            f'<pageSetup paperSize="1" orientation="landscape" scale="{scale}"/></worksheet>')
        sheet_files.append((name, xml))

    fills = ['<fill><patternFill patternType="none"/></fill>',
             '<fill><patternFill patternType="gray125"/></fill>']
    # one fill per column of the widest sheet (palette repeats only past 960)
    max_cols = 0
    for name, face, size, italic, scale, fmt, sp in SHEETS:
        if sp != split:
            continue
        n = len(wide_columns(fmt, face))
        for _, v in dense_columns(fmt):
            n += len(dense_widths(fmt_text(v, fmt), face, size, scale / 100.0, exact_height=True))
        max_cols = max(max_cols, n)
    max_cols = max(max_cols, 2)
    for i in range(max_cols):
        r, g, b = fill_color(i)
        fills.append(f'<fill><patternFill patternType="solid"><fgColor rgb="FF{r:02X}{g:02X}{b:02X}"/>'
                     f'<bgColor indexed="64"/></patternFill></fill>')

    styles = DECL + (
        f'<styleSheet xmlns="{MAIN}">'
        f'<numFmts count="{len(used_fmts)}">{numfmts}</numFmts>'
        f'<fonts count="{len(fonts)}">' + ''.join(f'<font>{f}</font>' for f in fonts) + '</fonts>'
        f'<fills count="{len(fills)}">' + ''.join(fills) + '</fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        f'<cellXfs count="{len(xfs)}">' + ''.join(xfs) + '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')

    over = ''.join(f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" ContentType="{CT}.worksheet+xml"/>'
                   for i in range(len(sheet_files)))
    files = {
        '[Content_Types].xml': DECL + (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f'<Override PartName="/xl/workbook.xml" ContentType="{CT}.sheet.main+xml"/>'
            f'<Override PartName="/xl/styles.xml" ContentType="{CT}.styles+xml"/>{over}</Types>'),
        '_rels/.rels': DECL + f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml': DECL + (
            f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><bookViews><workbookView/></bookViews><sheets>'
            + ''.join(f'<sheet name="{name}" sheetId="{i + 1}" r:id="rId{i + 1}"/>'
                      for i, (name, _) in enumerate(sheet_files))
            + '</sheets></workbook>'),
        'xl/_rels/workbook.xml.rels': DECL + (
            f'<Relationships xmlns="{PKG}">'
            + ''.join(f'<Relationship Id="rId{i + 1}" Type="{REL}/worksheet" Target="worksheets/sheet{i + 1}.xml"/>'
                      for i in range(len(sheet_files)))
            + f'<Relationship Id="rId{len(sheet_files) + 1}" Type="{REL}/styles" Target="styles.xml"/></Relationships>'),
        'xl/styles.xml': styles,
    }
    for i, (_, xml) in enumerate(sheet_files):
        files[f'xl/worksheets/sheet{i + 1}.xml'] = xml
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, text in files.items():
            info = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, text)
    with open(meta_path, 'w', encoding='utf-8', newline='\n') as f:
        f.write('\n'.join(meta) + '\n')
    print(f'wrote {out} ({len(sheet_files)} sheets, {len(meta) - 1} cells) and {meta_path}')


if __name__ == '__main__':
    if len(sys.argv) != 4 or sys.argv[3] not in ('train', 'holdout'):
        raise SystemExit(__doc__)
    build(sys.argv[1], sys.argv[2], sys.argv[3])
