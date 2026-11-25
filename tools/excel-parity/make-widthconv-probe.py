#!/usr/bin/env python3
"""Probe cells for ASC, DBCS, JIS and PHONETIC (FnText.fs), as eval-cells-on-vm.sh takes them.

    python3 tools/excel-parity/make-widthconv-probe.py > widthconv-cells.txt

One line per cell: sheet, address, formula. The first cells are hand-picked cases (katakana with voiced
and semi-voiced marks, punctuation, the yen sign and quotes, argument types); then UNICODE(ASC(CHAR))
and UNICODE(DBCS(...)) of every code point of the blocks the conversions touch (ASCII, Latin-1, the
CJK symbols, hiragana, katakana, the half-width and full-width forms), so that the tables of FnText.fs
are Excel's. widthconv-excel.tsv holds Excel 2019's values; function-parity-engine.fsx compares.
"""
cells = [
    '=ASC("ＡＢＣ１２３")', '=ASC("アイウ")', '=ASC("ガギグ")', '=ASC("ｶﾞ")', '=ASC("abc")', '=ASC("　")', '=ASC("パ")',
    '=DBCS("ABC123")', '=DBCS("ｶﾞｷﾞ")', '=DBCS("abc")', '=DBCS(" ")', '=DBCS("ﾊﾟ")', '=DBCS("ｳﾞ")',
    '=JIS("ABC")', '=ASC(123)', '=DBCS(123)', '=ASC(TRUE)', '=DBCS(1.5)', '=ASC(#N/A)', '=DBCS(A1)', '=ASC("")',
    '=PHONETIC(A1)', '=PHONETIC(B1:B2)', '=PHONETIC(A1:A2 B1:B2)', '=PHONETIC(INDIRECT("A1"))',
    '=LEN(DBCS("ABC"))', '=LEN(DBCS("ｶﾞ"))', '=LEN(ASC("ガ"))', '=LEN(ASC("ヴ"))', '=LEN(DBCS("ｳﾞ"))',
    '=LENB("ガ")', '=LENB("ｶ")', '=DBCS("ﾊﾞﾋﾟ")', '=DBCS("ﾞ")', '=DBCS("aﾞ")', '=DBCS("ﾜﾞ")', '=DBCS("ｦﾞ")',
    '=ASC({"Ａ","Ｂ"})', '=ROWS(DBCS({"a";"b"}))',
]
# Every code point of the blocks, one at a time.
blocks = list(range(0x20, 0x7F)) + list(range(0xA0, 0x100)) + list(range(0x2010, 0x2040)) + [0x2116, 0x2121, 0x2160, 0x2190, 0x2191, 0x2192, 0x2193, 0x2212, 0x2225, 0x2502, 0x25A0, 0x25CB] \
    + list(range(0x3000, 0x3020)) + list(range(0x3040, 0x3100)) + list(range(0x3130, 0x3190)) + list(range(0xFF00, 0xFFF0))
for cp in blocks:
    cells.append('=UNICODE(ASC(UNICHAR(%d)))&","&LEN(ASC(UNICHAR(%d)))&","&IFERROR(UNICODE(MID(ASC(UNICHAR(%d)),2,1)),"")' % (cp, cp, cp))
    cells.append('=UNICODE(DBCS(UNICHAR(%d)))&","&LEN(DBCS(UNICHAR(%d)))' % (cp, cp))
# Half-width katakana followed by the (semi-)voiced mark, for DBCS's joining.
for cp in range(0xFF66, 0xFF9E):
    for mark in (0xFF9E, 0xFF9F):
        cells.append('=UNICODE(DBCS(UNICHAR(%d)&UNICHAR(%d)))&","&LEN(DBCS(UNICHAR(%d)&UNICHAR(%d)))' % (cp, mark, cp, mark))
for i, c in enumerate(cells):
    print('S\tD%d\t%s' % (i + 1, c))
