#!/usr/bin/env python3
"""Writes src/Spreadsheet/GdiAdvances.fs: the advance widths Excel 2019 (GDI) draws a cell's text with, in whole pixels.

Excel draws a cell's text at a whole pixel size, round(points x 96/72 x zoom) (80 % and 85 % of 11 pt are both 12 px),
with the advance widths GDI hints the font to at that size; under 8 points on screen (points x zoom) not in the cell's
font at all but in Windows' raster "Small Fonts" (bold: a pixel wider a character). Checked in Excel 2019 on generated cells
(docs/excel-parity-tracker.md, "Cell text at a zoom"): FreeType's v35 TrueType interpreter (the one that hints
across, as GDI does) gives Excel's advance of every character of every font and size measured.

What this writes, per font face (a family with bold / italic) and for the characters U+0020-007E, U+00A0-00FF:
- the design advance in font units (with the face's units per em), so that a width can be had without the font;
- where the hinted advance at 11-36 px differs from the design advance rounded, the difference.
And the advance widths of Small Fonts' faces by their em height in pixels (the 96-dpi and 120-dpi files both
installed: Excel's 10 px is the 120-dpi 6-pt face, 10.5 px its 7-pt face of 11 px).

Needs freetype-py, the Windows fonts (artifacts/fonts/windows, windows-extra) and Windows' smalle.fon / smallf.fon
(C:\\Windows\\Fonts, in artifacts/fonts/raster):

    FREETYPE_PROPERTIES=truetype:interpreter-version=35 python3 tools/excel-parity/make-gdi-advances.py
"""
import os
import struct
import sys

if 'interpreter-version=35' not in os.environ.get('FREETYPE_PROPERTIES', ''):
    os.environ['FREETYPE_PROPERTIES'] = 'truetype:interpreter-version=35'
    os.execv(sys.executable, [sys.executable] + sys.argv)
import freetype  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.normpath(os.path.join(HERE, '..', '..'))
FONTS = os.path.join(REPO, 'artifacts', 'fonts')
OUT = os.path.join(REPO, 'src', 'Spreadsheet', 'GdiAdvances.fs')
CHARS = list(range(0x20, 0x7F)) + list(range(0xA0, 0x100))
PPEMS = range(11, 37)
# (family, bold, italic, file, face index in a collection)
FACES = [
    ('Calibri', 0, 0, 'windows/calibri.ttf', 0), ('Calibri', 1, 0, 'windows/calibrib.ttf', 0),
    ('Calibri', 0, 1, 'windows/calibrii.ttf', 0), ('Calibri', 1, 1, 'windows/calibriz.ttf', 0),
    ('Calibri Light', 0, 0, 'windows/calibril.ttf', 0), ('Calibri Light', 0, 1, 'windows/calibrili.ttf', 0),
    ('Arial', 0, 0, 'windows/arial.ttf', 0), ('Arial', 1, 0, 'windows/arialbd.ttf', 0),
    ('Arial', 0, 1, 'windows/ariali.ttf', 0), ('Arial', 1, 1, 'windows/arialbi.ttf', 0),
    ('Arial Narrow', 0, 0, 'windows/ARIALN.TTF', 0), ('Arial Narrow', 1, 0, 'windows/ARIALNB.TTF', 0),
    ('Arial Narrow', 0, 1, 'windows/ARIALNI.TTF', 0), ('Arial Narrow', 1, 1, 'windows/ARIALNBI.TTF', 0),
    ('Arial Black', 0, 0, 'windows-extra/ariblk.ttf', 0),
    ('Times New Roman', 0, 0, 'windows/times.ttf', 0), ('Times New Roman', 1, 0, 'windows/timesbd.ttf', 0),
    ('Times New Roman', 0, 1, 'windows/timesi.ttf', 0), ('Times New Roman', 1, 1, 'windows/timesbi.ttf', 0),
    ('Cambria', 0, 0, 'windows/cambria.ttc', 0), ('Cambria', 1, 0, 'windows/cambriab.ttf', 0),
    ('Cambria', 0, 1, 'windows/cambriai.ttf', 0), ('Cambria', 1, 1, 'windows/cambriaz.ttf', 0),
    ('Verdana', 0, 0, 'windows/verdana.ttf', 0), ('Verdana', 1, 0, 'windows/verdanab.ttf', 0),
    ('Verdana', 0, 1, 'windows/verdanai.ttf', 0), ('Verdana', 1, 1, 'windows/verdanaz.ttf', 0),
    ('Tahoma', 0, 0, 'windows-extra/tahoma.ttf', 0), ('Tahoma', 1, 0, 'windows-extra/tahomabd.ttf', 0),
    ('Segoe UI', 0, 0, 'windows/segoeui.ttf', 0), ('Segoe UI', 1, 0, 'windows/segoeuib.ttf', 0),
    ('Segoe UI', 0, 1, 'windows/segoeuii.ttf', 0), ('Segoe UI', 1, 1, 'windows/segoeuiz.ttf', 0),
    ('Segoe UI Black', 0, 0, 'windows-extra/seguibl.ttf', 0), ('Segoe UI Black', 0, 1, 'windows-extra/seguibli.ttf', 0),
    ('Garamond', 0, 0, 'windows/GARA.TTF', 0), ('Garamond', 1, 0, 'windows/GARABD.TTF', 0), ('Garamond', 0, 1, 'windows/GARAIT.TTF', 0),
    ('Georgia', 0, 0, 'windows-extra/georgia.ttf', 0), ('Georgia', 1, 0, 'windows-extra/georgiab.ttf', 0),
    ('Georgia', 0, 1, 'windows-extra/georgiai.ttf', 0), ('Georgia', 1, 1, 'windows-extra/georgiaz.ttf', 0),
    ('Trebuchet MS', 0, 0, 'windows-extra/trebuc.ttf', 0), ('Trebuchet MS', 1, 0, 'windows-extra/trebucbd.ttf', 0),
    ('Trebuchet MS', 0, 1, 'windows-extra/trebucit.ttf', 0), ('Trebuchet MS', 1, 1, 'windows-extra/trebucbi.ttf', 0),
    ('Courier New', 0, 0, 'windows-extra/cour.ttf', 0), ('Courier New', 1, 0, 'windows-extra/courbd.ttf', 0),
    ('Courier New', 0, 1, 'windows-extra/couri.ttf', 0), ('Courier New', 1, 1, 'windows-extra/courbi.ttf', 0),
    ('Consolas', 0, 0, 'windows-extra/consola.ttf', 0), ('Consolas', 1, 0, 'windows-extra/consolab.ttf', 0),
    ('Consolas', 0, 1, 'windows-extra/consolai.ttf', 0), ('Consolas', 1, 1, 'windows-extra/consolaz.ttf', 0),
    ('Lucida Console', 0, 0, 'windows-extra/lucon.ttf', 0),
    ('Aptos Narrow', 0, 0, 'windows-extra/aptosnarrow-27970306811.ttf', 0), ('Aptos Narrow', 1, 0, 'windows-extra/aptosnarrow-25065980124.ttf', 0),
    ('Aptos Narrow', 0, 1, 'windows-extra/aptosnarrow-28903934312.ttf', 0),
    ('Tw Cen MT', 0, 0, 'windows-extra/TCM_____.TTF', 0), ('Tw Cen MT', 1, 0, 'windows-extra/TCB_____.TTF', 0),
    ('Tw Cen MT', 0, 1, 'windows-extra/TCMI____.TTF', 0), ('Tw Cen MT', 1, 1, 'windows-extra/TCBI____.TTF', 0),
]


def base36(n):
    """Three base-36 digits ('zzz': no glyph)."""
    if n < 0:
        return 'zzz'
    assert n < 36 ** 3 - 1
    digits = '0123456789abcdefghijklmnopqrstuvwxyz'
    return digits[n // 1296] + digits[n // 36 % 36] + digits[n % 36]


def rounded(x):
    return int(x + 0.5) if x >= 0 else -int(-x + 0.5)


def face_table(path, index):
    face = freetype.Face(path, index)
    upm = face.units_per_EM
    design = []
    for code in CHARS:
        if face.get_char_index(code) == 0:
            design.append(-1)
            continue
        face.set_pixel_sizes(0, 64)
        face.load_char(chr(code), freetype.FT_LOAD_NO_SCALE)
        design.append(face.glyph.metrics.horiAdvance)
    segments = []
    for ppem in PPEMS:
        face.set_pixel_sizes(0, ppem)
        seg = ''
        for i, code in enumerate(CHARS):
            if design[i] < 0:
                continue
            face.load_char(chr(code), freetype.FT_LOAD_DEFAULT)
            hinted = face.glyph.advance.x // 64
            delta = hinted - rounded(design[i] * ppem / upm)
            if delta:
                assert -9 <= delta <= 9, (path, ppem, code, delta)
                seg += '%02x%s' % (i, 'abcdefghijklmnopqrs'[delta + 9])
        segments.append(seg)
    return upm, design, segments


def raster_faces(path):
    """The FNT resources of a .fon: (em height, cell height, widths of 0x20-0xFF)."""
    d = open(path, 'rb').read()
    ne = struct.unpack_from('<H', d, 0x3C)[0]
    rt = ne + struct.unpack_from('<H', d, ne + 0x24)[0]
    shift = struct.unpack_from('<H', d, rt)[0]
    p = rt + 2
    out = []
    while True:
        typ = struct.unpack_from('<H', d, p)[0]
        if typ == 0:
            break
        count = struct.unpack_from('<H', d, p + 2)[0]
        p += 8
        for _ in range(count):
            off, _length = struct.unpack_from('<HH', d, p)
            p += 12
            if typ == 0x8008:
                f = d[off << shift:]
                version = struct.unpack_from('<H', f, 0)[0]
                internal_leading = struct.unpack_from('<H', f, 0x4C)[0]
                height = struct.unpack_from('<H', f, 0x58)[0]
                charset = f[0x55]
                first, last = f[0x5F], f[0x60]
                base, step = (0x76, 4) if version == 0x200 else (0x94, 6)
                widths = {first + k: struct.unpack_from('<H', f, base + k * step)[0] for k in range(last - first + 1)}
                out.append((height - internal_leading, height, charset, [widths.get(c, widths.get(f[0x61] + first, 0)) for c in range(0x20, 0x100)]))
    return out


def win_extent(path, index):
    """A line of the face as GDI spaces lines, in font units, and its units per em: tmHeight (OS/2 usWinAscent +
    usWinDescent) and tmExternalLeading (what the hhea line - ascender, descender and line gap - has more)."""
    data = open(path, 'rb').read()
    base = 0
    if data[:4] == b'ttcf':
        base = struct.unpack('>I', data[12 + 4 * index:16 + 4 * index])[0]
    tables = {}
    for i in range(struct.unpack('>H', data[base + 4:base + 6])[0]):
        tag, _, offset, _ = struct.unpack('>4sIII', data[base + 12 + 16 * i:base + 28 + 16 * i])
        tables[tag] = offset
    upm = struct.unpack('>H', data[tables[b'head'] + 18:tables[b'head'] + 20])[0]
    ascent, descent = struct.unpack('>HH', data[tables[b'OS/2'] + 74:tables[b'OS/2'] + 78])
    above, below, gap = struct.unpack('>hhh', data[tables[b'hhea'] + 4:tables[b'hhea'] + 10])
    return max(ascent + descent, above - below + gap), upm


def main():
    lines = ['/// Generated by tools/excel-parity/make-gdi-advances.py - do not edit.',
             '/// The advance widths Excel 2019 (GDI) draws a cell\'s text with (`ExcelText`): per font face, the design advances',
             '/// of U+0020-007E and U+00A0-00FF in font units (three base-36 digits each, \'zzz\' none) and, at 11-36 px,',
             '/// where GDI\'s hinted advance differs from the design advance rounded, the difference (two hex digits: the',
             '/// character\'s index; a letter: the difference, \'j\' 0); Small Fonts\' advances by em height (a hex digit',
             '/// a character, U+0020-00FF).',
             'module Rbt.Spreadsheet.GdiAdvances', '',
             '/// The characters a face\'s tables cover, by index.',
             'let charIndex (code: int) = if code >= 0x20 && code < 0x7F then code - 0x20 elif code >= 0xA0 && code <= 0xFF then code - 0xA0 + 95 else -1', '',
             '/// The sizes of the hinted differences (a face\'s segments are 11, 12 ... 36 px).',
             f'let firstHintedPx = {PPEMS.start}',
             f'let lastHintedPx = {PPEMS.stop - 1}', '',
             '/// (family lower case, bold, italic, units per em, design advances, hinted differences per size)',
             'let faces : (string * bool * bool * int * string * string[])[] = [|']
    for family, bold, italic, rel, index in FACES:
        path = os.path.join(FONTS, rel)
        if not os.path.exists(path):
            print('missing', path, file=sys.stderr)
            continue
        upm, design, segments = face_table(path, index)
        lines.append(f'    ("{family.lower()}", {"true" if bold else "false"}, {"true" if italic else "false"}, {upm},')
        lines.append('     "' + ''.join(base36(x) for x in design) + '",')
        lines.append('     [| ' + '; '.join(f'"{s}"' for s in segments) + ' |])')
    lines.append('|]')
    lines.append('')
    lines.append('/// A line of each family (its regular face) as GDI spaces lines, tmHeight + tmExternalLeading: (family lower case, the line in font units, units per em).')
    lines.append('let lineExtents : (string * int * int)[] = [|')
    for family, bold, italic, rel, index in FACES:
        path = os.path.join(FONTS, rel)
        if bold or italic or not os.path.exists(path):
            continue
        extent, upm = win_extent(path, index)
        lines.append(f'    ("{family.lower()}", {extent}, {upm})')
    lines.append('|]')
    lines.append('')
    raster = []
    for name, dpi in (('smalle.fon', 96), ('smallf.fon', 120)):
        seen = set()
        for em, height, charset, widths in raster_faces(os.path.join(FONTS, 'raster', name)):
            if height in seen:
                continue
            seen.add(height)
            raster.append((em, dpi, height, widths))
    chosen = {}
    for em, dpi, height, widths in sorted(raster, key=lambda r: (r[0], r[1])):
        chosen.setdefault(em, (dpi, height, widths))
    lines.append('/// Small Fonts by em height in pixels (the 96-dpi face where both have one): (em, advances of U+0020-00FF).')
    lines.append('let smallFonts : (int * string)[] = [|')
    for em in sorted(chosen):
        dpi, height, widths = chosen[em]
        assert max(widths) < 16
        lines.append(f'    ({em}, "{"".join("%x" % w for w in widths)}") // {dpi} dpi, {height} px high')
    lines.append('|]')
    open(OUT, 'w').write('\n'.join(lines) + '\n')
    print(OUT, os.path.getsize(OUT), 'bytes')


main()
