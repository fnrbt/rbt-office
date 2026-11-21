#!/usr/bin/env python3
"""Writes tests/test-documents-excel/synthetic-grid-parity.xlsx, a synthetic workbook (not saved by Excel) with
one case of each of the rules the browser grid follows to draw a sheet as Excel 2019 does (found on real-world
real-estate models, checked in Excel; tools/xlsx-browser check `grid-excel-parity`):

- "Grid", Arial 10 in a sheet of 12.75-pt rows:
  - A1: a bold cell of rich text, "Note: " without run properties (bold, as the cell) and a run with its own
    (italic only): Excel draws that run upright weight, not bold.
  - row 2: ht="15" without customHeight: Excel fits the row to its cells when it opens the file - 12.75 pt
    (17 px), a line of Arial 10 - whatever `ht` says (make-sizing-fixtures.py).
  - B3:E5: thin lines round every cell; C4 has a conditional format whose dxf names every side without a
    style (`<left/><right/>...`): no line on any of its edges, not even its neighbours'; E4 one of a dashed
    blue left and a thick red top, in place of its neighbours' thin lines.
  - A7: blue text, a conditional format of `<font><color auto="1"/></font>` on a black fill: black on black.
  - B9: wrapped text, set in 3 px from the cell's sides.
  - A13: left-aligned text with an indent, longer than its column: it runs over the empty cells beside it.
  - F:XFD have a grey default style (`<col min="6" max="16384" style=...>`): every cell of them is grey.
  - Drawings: a red rectangle marked hidden (`cNvPr hidden="1"`), a text box with no fill, no line and no
    text, a hidden picture - none of them drawn by Excel -, and a linked picture (Excel's camera,
    `a14:cameraTool`, in `mc:AlternateContent` as Excel writes it) of A10:B11.
- "Zoom" at 70 %: A1 a medium bottom line (2 screen pixels at any zoom), A3:C3 yellow cells that meet.

Deterministic: run it again and the bytes are the same.

    python3 tools/excel-parity/make-grid-parity-fixture.py
"""
import os
import struct
import zipfile
import zlib

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.normpath(os.path.join(HERE, '..', '..'))
OUT = os.path.join(REPO, 'tests', 'test-documents-excel', 'synthetic-grid-parity.xlsx')
DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
A14 = 'http://schemas.microsoft.com/office/drawing/2010/main'


def png(width, height, rgb):
    """A PNG of one colour."""
    def chunk(kind, data):
        return struct.pack('>I', len(data)) + kind + data + struct.pack('>I', zlib.crc32(kind + data) & 0xffffffff)
    rows = b''.join(b'\x00' + bytes(rgb) * width for _ in range(height))
    return (b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', width, height, 8, 2, 0, 0, 0))
            + chunk(b'IDAT', zlib.compress(rows, 9)) + chunk(b'IEND', b''))


PNG = png(4, 4, (128, 128, 128))

STYLES = DECL + (
    f'<styleSheet xmlns="{MAIN}">'
    '<numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0"/></numFmts>'
    '<fonts count="3">'
    '<font><sz val="10"/><name val="Arial"/><family val="2"/></font>'
    '<font><b/><sz val="10"/><name val="Arial"/><family val="2"/></font>'
    '<font><sz val="10"/><color rgb="FF0000FF"/><name val="Arial"/><family val="2"/></font>'
    '</fonts>'
    '<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FFD9D9D9"/><bgColor indexed="64"/></patternFill></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill></fills>'
    '<borders count="3"><border><left/><right/><top/><bottom/><diagonal/></border>'
    '<border><left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right>'
    '<top style="thin"><color indexed="64"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/></border>'
    '<border><left/><right/><top/><bottom style="medium"><color indexed="64"/></bottom><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="9">'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'                                   # 0
    '<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>'                     # 1 grey
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'                     # 2 bold
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>'                   # 3 thin box
    '<xf numFmtId="164" fontId="2" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>'  # 4 blue
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>'  # 5 wrap
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left" indent="2"/></xf>'  # 6 indent
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1"/>'                   # 7 medium bottom
    '<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>'                     # 8 yellow
    '</cellXfs>'
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    '<dxfs count="3">'
    '<dxf><border><left/><right/><top/><bottom/><vertical/><horizontal/></border></dxf>'
    '<dxf><border><left style="dashed"><color rgb="FF0000FF"/></left><top style="thick"><color rgb="FFFF0000"/></top></border></dxf>'
    '<dxf><font><color auto="1"/></font><fill><patternFill><bgColor rgb="FF000000"/></patternFill></fill></dxf>'
    '</dxfs></styleSheet>')


def inline(ref, style, text):
    return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'


GRID_ROWS = [
    '<row r="1"><c r="A1" s="2" t="inlineStr"><is><r><t xml:space="preserve">Note: </t></r>'
    '<r><rPr><i/><sz val="10"/><rFont val="Arial"/><family val="2"/></rPr><t>italic only</t></r></is></c></row>',
    '<row r="2" ht="15">' + inline('A2', 0, 'a row of 15 pt') + '</row>',
    '<row r="3">' + ''.join(f'<c r="{c}3" s="3"/>' for c in 'BCDE') + '</row>',
    '<row r="4"><c r="B4" s="3"/>' + inline('C4', 3, 'x') + '<c r="D4" s="3"/>' + inline('E4', 3, 'y') + '</row>',
    '<row r="5">' + ''.join(f'<c r="{c}5" s="3"/>' for c in 'BCDE') + '</row>',
    '<row r="7"><c r="A7" s="4"><v>4400000</v></c></row>',
    '<row r="9">' + inline('B9', 5, 'Releasing Cost per Unit') + '</row>',
    '<row r="10">' + inline('A10', 0, 'Linked') + '<c r="B10" s="0"><v>7.27</v></c></row>',
    '<row r="11">' + inline('A11', 2, 'Picture') + '<c r="B11" s="0"><v>12</v></c></row>',
    '<row r="13">' + inline('A13', 6, 'Mortgage Information and more') + '</row>',
]
GRID = DECL + (
    f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}">'
    '<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>'
    '<sheetFormatPr defaultRowHeight="12.75"/>'
    '<cols><col min="1" max="5" width="12.7109375" customWidth="1"/><col min="6" max="16384" width="9.140625" style="1"/></cols>'
    '<sheetData>' + ''.join(GRID_ROWS) + '</sheetData>'
    '<conditionalFormatting sqref="C4"><cfRule type="expression" dxfId="0" priority="1"><formula>TRUE</formula></cfRule></conditionalFormatting>'
    '<conditionalFormatting sqref="E4"><cfRule type="expression" dxfId="1" priority="2"><formula>TRUE</formula></cfRule></conditionalFormatting>'
    '<conditionalFormatting sqref="A7"><cfRule type="expression" dxfId="2" priority="3"><formula>$A$7&gt;0</formula></cfRule></conditionalFormatting>'
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
    '<drawing r:id="rId1"/></worksheet>')
ZOOM = DECL + (
    f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}">'
    '<sheetViews><sheetView zoomScale="70" zoomScaleNormal="70" workbookViewId="0"/></sheetViews>'
    '<sheetFormatPr defaultRowHeight="12.75"/>'
    '<sheetData><row r="1">' + inline('A1', 7, 'medium line below') + '</row>'
    '<row r="3">' + ''.join(f'<c r="{c}3" s="8"/>' for c in 'ABC') + '</row></sheetData>'
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>')


def anchor(col0, row0, col1, row1, body, edit=''):
    return (f'<xdr:twoCellAnchor{edit}><xdr:from><xdr:col>{col0}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row0}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
            f'<xdr:to><xdr:col>{col1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
            + body + '<xdr:clientData/></xdr:twoCellAnchor>')


def shape(id_, name, extra_props, fill, text, hidden=False):
    return (f'<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{id_}" name="{name}"{" hidden=\"1\"" if hidden else ""}/>'
            f'<xdr:cNvSpPr{extra_props}/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="571500"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{fill}</xdr:spPr>'
            '<xdr:txBody><a:bodyPr vertOverflow="clip" wrap="square" rtlCol="0" anchor="t"/><a:lstStyle/>'
            f'<a:p>{text}<a:endParaRPr lang="en-US" sz="1100"/></a:p></xdr:txBody></xdr:sp>')


def picture(id_, name, hidden=False, camera=None):
    link = ''
    if camera:
        link = (f'<a:extLst><a:ext uri="{{84589F7E-364E-4C9E-8A38-B11213B215E9}}"><a14:cameraTool cellRange="{camera}" spid="_x0000_s1025"/>'
                '</a:ext></a:extLst>')
    return (f'<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{id_}" name="{name}"{" hidden=\"1\"" if hidden else ""}/>'
            f'<xdr:cNvPicPr><a:picLocks noChangeAspect="1" noChangeArrowheads="1"/>{link}</xdr:cNvPicPr></xdr:nvPicPr>'
            f'<xdr:blipFill><a:blip xmlns:r="{REL}" r:embed="rId1"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
            '<xdr:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="1219200" cy="381000"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></xdr:spPr></xdr:pic>')


CAMERA = "'Grid'!$A$10:$B$11"
DRAWING = DECL + (
    f'<xdr:wsDr xmlns:xdr="{XDR}" xmlns:a="{A}">'
    + anchor(7, 1, 10, 4, shape(2, 'Hidden Rectangle', '', '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>',
                                 '<a:r><a:rPr lang="en-US" sz="1100"/><a:t>HIDDEN</a:t></a:r>', hidden=True))
    + anchor(7, 5, 10, 8, shape(3, 'Blank Box', ' txBox="1"', '<a:noFill/><a:ln w="9525" cmpd="sng"><a:noFill/></a:ln>', ''))
    + f'<mc:AlternateContent xmlns:mc="{MC}"><mc:Choice xmlns:a14="{A14}" Requires="a14">'
    + anchor(7, 9, 10, 12, picture(4, 'Picture 3', camera=CAMERA), ' editAs="oneCell"')
    + '</mc:Choice><mc:Fallback>' + anchor(7, 9, 10, 12, picture(4, 'Picture 3'), ' editAs="oneCell"')
    + '</mc:Fallback></mc:AlternateContent>'
    + anchor(7, 14, 9, 16, picture(5, 'Hidden Picture', hidden=True), ' editAs="oneCell"')
    + '</xdr:wsDr>')

FILES = {
    '[Content_Types].xml': DECL + (
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
        '</Types>'),
    '_rels/.rels': DECL + f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    'xl/workbook.xml': DECL + (f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><bookViews><workbookView/></bookViews><sheets>'
                               '<sheet name="Grid" sheetId="1" r:id="rId1"/><sheet name="Zoom" sheetId="2" r:id="rId2"/></sheets></workbook>'),
    'xl/_rels/workbook.xml.rels': DECL + (
        f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/worksheet" Target="worksheets/sheet1.xml"/>'
        f'<Relationship Id="rId2" Type="{REL}/worksheet" Target="worksheets/sheet2.xml"/>'
        f'<Relationship Id="rId3" Type="{REL}/styles" Target="styles.xml"/></Relationships>'),
    'xl/styles.xml': STYLES,
    'xl/worksheets/sheet1.xml': GRID,
    'xl/worksheets/_rels/sheet1.xml.rels': DECL + f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/drawing" Target="../drawings/drawing1.xml"/></Relationships>',
    'xl/worksheets/sheet2.xml': ZOOM,
    'xl/drawings/drawing1.xml': DRAWING,
    'xl/drawings/_rels/drawing1.xml.rels': DECL + f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/image" Target="../media/image1.png"/></Relationships>',
}

with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
    for name, text in FILES.items():
        info = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
        info.compress_type = zipfile.ZIP_DEFLATED
        z.writestr(info, text)
    info = zipfile.ZipInfo('xl/media/image1.png', (1980, 1, 1, 0, 0, 0))
    z.writestr(info, PNG)
print(OUT)
