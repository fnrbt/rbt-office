#!/usr/bin/env python3
"""Locks the second (hint-discriminating) challenge before its Excel export.

For each row-1 cell this writes the design-rounded baseline advance and the
FreeType-v35 hinted advance at ppem = floor(size*scale*600/72+0.5), the fit
reserve c, the model W and both full/hash decisions.  FreeType is called through
libfreetype with interpreter-version=35, the same GDI-compatible interpreter
`make-gdi-advances.py` uses.

    python3 tools/excel-parity/lock-print-hint-oracle.py META.tsv WORKBOOK.xlsx OUT.tsv
"""
import ctypes
import csv
import hashlib
import importlib.util
import math
import os
import struct
import sys

_here = os.path.dirname(os.path.abspath(__file__))
_spec = importlib.util.spec_from_file_location('gen', os.path.join(_here, 'make-print-advance-oracle.py'))
gen = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(gen)

_lib = ctypes.CDLL('libfreetype.so.6')
for _n, _a, _r in [
        ('FT_Init_FreeType', [ctypes.POINTER(ctypes.c_void_p)], ctypes.c_int),
        ('FT_New_Face', [ctypes.c_void_p, ctypes.c_char_p, ctypes.c_long, ctypes.POINTER(ctypes.c_void_p)], ctypes.c_int),
        ('FT_Set_Pixel_Sizes', [ctypes.c_void_p, ctypes.c_uint, ctypes.c_uint], ctypes.c_int),
        ('FT_Get_Char_Index', [ctypes.c_void_p, ctypes.c_ulong], ctypes.c_uint),
        ('FT_Get_Advance', [ctypes.c_void_p, ctypes.c_uint, ctypes.c_int32, ctypes.POINTER(ctypes.c_long)], ctypes.c_int),
        ('FT_Property_Set', [ctypes.c_void_p, ctypes.c_char_p, ctypes.c_char_p, ctypes.c_void_p], ctypes.c_int)]:
    _f = getattr(_lib, _n)
    _f.argtypes = _a
    _f.restype = _r
_lp = ctypes.c_void_p()
assert _lib.FT_Init_FreeType(ctypes.byref(_lp)) == 0
assert _lib.FT_Property_Set(_lp, b'truetype', b'interpreter-version', ctypes.byref(ctypes.c_uint(35))) == 0
_faces = {}


def face(path):
    if path not in _faces:
        f = ctypes.c_void_p()
        assert _lib.FT_New_Face(_lp, path.encode(), 0, ctypes.byref(f)) == 0
        _faces[path] = f
    return _faces[path]


def hinted_px(path, ppem, ch):
    f = face(path)
    assert _lib.FT_Set_Pixel_Sizes(f, 0, ppem) == 0
    gi = _lib.FT_Get_Char_Index(f, ord(ch))
    adv = ctypes.c_long()
    assert _lib.FT_Get_Advance(f, gi, 0, ctypes.byref(adv)) == 0
    return adv.value / 65536.0


def font_path(face_name):
    return os.path.join(gen.ROOT, gen.FONT_FILES[(face_name, False)])


def design_metrics(path):
    # design advances for the characters this experiment uses
    d = open(path, 'rb').read()
    base = 0
    if d[:4] == b'ttcf':
        base = struct.unpack_from('>I', d, 12)[0]
    n = struct.unpack_from('>H', d, base + 4)[0]
    t = {d[base + 12 + 16 * i:base + 16 + 16 * i].decode('latin1'):
         struct.unpack_from('>II', d, base + 20 + 16 * i) for i in range(n)}
    upm = struct.unpack_from('>H', d, t['head'][0] + 18)[0]
    hhea = t['hhea'][0]
    nhm = struct.unpack_from('>H', d, hhea + 34)[0]
    hm = t['hmtx'][0]
    cmap = t['cmap'][0]
    sub = None
    for i in range(struct.unpack_from('>H', d, cmap + 2)[0]):
        pid, eid = struct.unpack_from('>H', d, cmap + 4 + 8 * i)[0], struct.unpack_from('>H', d, cmap + 6 + 8 * i)[0]
        so = struct.unpack_from('>I', d, cmap + 8 + 8 * i)[0]
        if struct.unpack_from('>H', d, cmap + so)[0] == 4 and (pid, eid) in ((3, 1), (0, 3), (0, 4)):
            sub = cmap + so
            break
    seg = struct.unpack_from('>H', d, sub + 6)[0] // 2
    end = sub + 14
    start = end + 2 * seg + 2
    delta = start + 2 * seg
    ro = delta + 2 * seg

    def lookup(code):
        for i in range(seg):
            s, e = struct.unpack_from('>H', d, start + 2 * i)[0], struct.unpack_from('>H', d, end + 2 * i)[0]
            if s <= code <= e:
                de = struct.unpack_from('>h', d, delta + 2 * i)[0]
                r = struct.unpack_from('>H', d, ro + 2 * i)[0]
                if r == 0:
                    return (code + de) & 0xFFFF
                g = struct.unpack_from('>H', d, ro + 2 * i + r + 2 * (code - s))[0]
                return (g + de) & 0xFFFF if g else 0
        return 0

    out = {}
    for ch in '$(),.% 0123456789#-':
        g = lookup(ord(ch))
        if g:
            out[ch] = struct.unpack_from('>H', d, hm + 4 * min(g, nhm - 1))[0] / upm
    return out


def main():
    if len(sys.argv) != 4:
        raise SystemExit(__doc__)
    meta_path, workbook, out_path = sys.argv[1:]
    with open(meta_path, encoding='utf-8') as f:
        meta = list(csv.DictReader(f, delimiter='\t'))
    head = ['sheet', 'column', 'row', 'width_chars', 'kind', 'role', 'face', 'size_pt', 'scale',
            'format', 'text', 'ppem', 'W_model_pt', 'c_pt', 'baseline_text_pt', 'hinted_text_pt',
            'baseline_full', 'hinted_full', 'baseline_margin_pt', 'hinted_margin_pt']
    out = ['# hint challenge lock: baseline = design advances rounded at ppem; '
           'hinted = FreeType v35 at the same ppem; no challenge PDF existed when written',
           '# workbook_sha256=' + hashlib.sha256(open(workbook, 'rb').read()).hexdigest(),
           '# meta_sha256=' + hashlib.sha256(open(meta_path, 'rb').read()).hexdigest(),
           '\t'.join(head)]
    counts = {'base_full': 0, 'hint_full': 0, 'n': 0}
    cache = {}
    for d in meta:
        if d['row'] != '1':
            continue
        face_name = d['face']
        if face_name not in cache:
            path = font_path(face_name)
            cache[face_name] = (path, design_metrics(path))
        path, ems = cache[face_name]
        size, scale = float(d['size_pt']), int(d['scale']) / 100.0
        H = size * scale * 600.0 / 72.0
        ppem = math.floor(H + 0.5)
        c = ems['0'] * size * scale / 2.0 + 0.875
        text = d['text']
        base_dots = sum(math.floor(ems[ch] * ppem + 0.5) for ch in text)
        hint_dots = sum(hinted_px(path, ppem, ch) for ch in text)
        base = base_dots * 72.0 / 600.0
        hint = hint_dots * 72.0 / 600.0
        W = gen.physical_width(float(d['width_chars']), scale)
        bf = base <= W - c + 1e-9
        hf = hint <= W - c + 1e-9
        counts['n'] += 1
        counts['base_full'] += bf
        counts['hint_full'] += hf
        out.append('\t'.join([d['sheet'], d['column'], d['row'], d['width_chars'], d['kind'], d['role'],
                              face_name, d['size_pt'], d['scale'], d['format'], text, str(ppem),
                              f'{W:.4f}', f'{c:.4f}', f'{base:.4f}', f'{hint:.4f}',
                              'yes' if bf else 'no', 'yes' if hf else 'no',
                              f'{W - c - base:+.4f}', f'{W - c - hint:+.4f}']))
    with open(out_path, 'w', encoding='utf-8', newline='\n') as f:
        f.write('\n'.join(out) + '\n')
    print(f'locked {counts["n"]} row-1 cells: baseline full {counts["base_full"]}, '
          f'hinted full {counts["hint_full"]} -> {out_path}')


if __name__ == '__main__':
    main()
