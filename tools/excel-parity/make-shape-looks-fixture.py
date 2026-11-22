#!/usr/bin/env python3
"""Writes tests/test-documents-excel/synthetic-shape-looks.xlsx, a synthetic workbook (not saved by Excel) of the
things a dashboard's cards are made of, each checked against Excel 2019's picture of it (tools/xlsx-browser check
`shape-looks-excel-parity`):

- "Links": cells in number formats with a locale tag (`[$$-1009]#,##0.00`, `[$$-409]`, `[$-4009]`, `[$-439]`) and
  shapes linked to them (`textlink`, a `<a:fld type="TxLink">` run) whose saved text is stale - grouped by lakhs,
  as a workbook last saved on a machine whose system locale is en-IN has it ("$22,97,200.86"). Excel shows the
  cell's value in its format, grouped by thousands; also for a cell of another sheet ('Data'!B2), and for a
  cell too narrow for its number (the shape shows the number, not `####`).
- "Colors": a column of cells, each a block of U+2588 in a colour of the number format: the eight colour names
  ([Black] .. [Cyan]) and [Color1] .. [Color56]; and a KPI-like `[Green]\\▲0.00%;[Red]\\▼0.00%;` at 20.62 % and
  -5.15 %.
- "Rounds": rounded rectangles of several sizes and adjustments (`adj` 3572 .. 50000, none) with a 1-pt line,
  rectangles with lines of 0.75 to 3 pt, a group with an outer shadow (as the dashboard's KPI cards: blur 4 pt,
  3 pt down-right at 45 degrees, black at 40 %) and a single shape with the same shadow; horizontal lines (no
  height) with arrowheads of each kind and size, a slanting one; an ellipse and a flow chart connector with a
  3-pt light line; a shape with the theme's third effect (a shadow) and one whose empty effect list overrides it;
  a shape of no fill, no line and no text (nothing is drawn).

Deterministic: run it again and the bytes are the same.

    python3 tools/excel-parity/make-shape-looks-fixture.py
"""
import os
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.normpath(os.path.join(HERE, '..', '..'))
OUT = os.path.join(REPO, 'tests', 'test-documents-excel', 'synthetic-shape-looks.xlsx')
DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
EMU = 9525  # a pixel at 96 dpi
COL, ROW = 64, 20  # the default column (8.43 of Calibri 11) and row (15 pt) in pixels


def esc(text):
    return text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


# Number formats: the locale-tagged currencies, the colours.
LINK_FORMATS = ['[$$-1009]#,##0.00', '[$$-1009]#,##0.00;\\-[$$-1009]#,##0.00', '[$$-409]#,##0.00', '[$-4009]#,##0.00',
                '[$-439]#,##0.00', '#,##0.00']
NAMES = ['Black', 'White', 'Red', 'Green', 'Blue', 'Yellow', 'Magenta', 'Cyan']
COLOR_FORMATS = [f'[{name}]"████"' for name in NAMES] + [f'[Color{n}]"████"' for n in range(1, 57)]
KPI_FORMAT = '[Green]\\▲0.00%;[Red]\\▼0.00%;'
FORMATS = LINK_FORMATS + COLOR_FORMATS + [KPI_FORMAT]
FMT_ID = {code: 164 + i for i, code in enumerate(FORMATS)}

# cellXfs: 0 Normal; 1.. one per format in Calibri 11; the colour cells in Arial 14 (font 1).
XFS = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>']
XF = {}
for code in FORMATS:
    XF[code] = len(XFS)
    font = 1 if code in COLOR_FORMATS or code == KPI_FORMAT else 0
    XFS.append(f'<xf numFmtId="{FMT_ID[code]}" fontId="{font}" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>')

STYLES = DECL + (
    f'<styleSheet xmlns="{MAIN}">'
    f'<numFmts count="{len(FORMATS)}">' + ''.join(f'<numFmt numFmtId="{FMT_ID[c]}" formatCode="{esc(c)}"/>' for c in FORMATS) + '</numFmts>'
    '<fonts count="2"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>'
    '<font><b/><sz val="14"/><color theme="1"/><name val="Arial"/><family val="2"/></font></fonts>'
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    f'<cellXfs count="{len(XFS)}">' + ''.join(XFS) + '</cellXfs>'
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    '<dxfs count="0"/></styleSheet>')


def text_cell(ref, text, style=0):
    return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{esc(text)}</t></is></c>'


def num_cell(ref, value, style):
    return f'<c r="{ref}" s="{style}"><v>{value}</v></c>'


def sheet(rows, cols='', drawing=False, view=''):
    return DECL + (
        f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}"><sheetViews><sheetView {view}workbookViewId="0"/></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>' + (f'<cols>{cols}</cols>' if cols else '')
        + '<sheetData>' + ''.join(rows) + '</sheetData>'
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
        + ('<drawing r:id="rId1"/>' if drawing else '') + '</worksheet>')


def anchor(x, y, w, h, body):
    """A two-cell anchor of a rectangle in pixels of a sheet of default columns and rows."""
    def at(px, size):
        return px // size, (px % size) * EMU
    c0, cx0 = at(x, COL)
    r0, ry0 = at(y, ROW)
    c1, cx1 = at(x + w, COL)
    r1, ry1 = at(y + h, ROW)
    return (f'<xdr:twoCellAnchor><xdr:from><xdr:col>{c0}</xdr:col><xdr:colOff>{cx0}</xdr:colOff><xdr:row>{r0}</xdr:row><xdr:rowOff>{ry0}</xdr:rowOff></xdr:from>'
            f'<xdr:to><xdr:col>{c1}</xdr:col><xdr:colOff>{cx1}</xdr:colOff><xdr:row>{r1}</xdr:row><xdr:rowOff>{ry1}</xdr:rowOff></xdr:to>'
            + body + '<xdr:clientData/></xdr:twoCellAnchor>')


def xfrm(x, y, w, h, tag='a:xfrm', child=None):
    inner = f'<a:off x="{x * EMU}" y="{y * EMU}"/><a:ext cx="{w * EMU}" cy="{h * EMU}"/>'
    if child:
        cx, cy, cw, ch = child
        inner += f'<a:chOff x="{cx * EMU}" y="{cy * EMU}"/><a:chExt cx="{cw * EMU}" cy="{ch * EMU}"/>'
    return f'<{tag}>{inner}</{tag}>'


def shape(id_, name, x, y, w, h, prst='rect', adj=None, line_w=12700, line='3E8853', fill='FFFFFF', text='', link=None, effects=''):
    av = f'<a:gd name="adj" fmla="val {adj}"/>' if adj is not None else ''
    # (line '': no <a:ln> at all - and no style: no line.)
    ln = f'<a:ln w="{line_w}"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>' if line else ('' if line == '' else '<a:ln><a:noFill/></a:ln>')
    fl = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>' if fill else '<a:noFill/>'
    rpr = '<a:rPr lang="en-US" sz="1500" b="1"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr>'
    if link:
        run = f'<a:fld id="{{5236D0F9-347D-48DD-9819-7AEA6394{id_:04X}}}" type="TxLink">{rpr}<a:pPr algn="ctr"/><a:t>{esc(text)}</a:t></a:fld>'
    elif text:
        run = f'<a:r>{rpr}<a:t>{esc(text)}</a:t></a:r>'
    else:
        run = ''
    return (f'<xdr:sp macro="" textlink="{esc(link or "")}"><xdr:nvSpPr><xdr:cNvPr id="{id_}" name="{name}"/><xdr:cNvSpPr/></xdr:nvSpPr>'
            f'<xdr:spPr>{xfrm(x, y, w, h)}<a:prstGeom prst="{prst}"><a:avLst>{av}</a:avLst></a:prstGeom>{fl}{ln}{effects}</xdr:spPr>'
            '<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" rtlCol="0" anchor="ctr"/><a:lstStyle/>'
            f'<a:p><a:pPr algn="ctr"/>{run}<a:endParaRPr lang="en-US" sz="1500" b="1"/></a:p></xdr:txBody></xdr:sp>')


SHADOW = ('<a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl" rotWithShape="0">'
          '<a:prstClr val="black"><a:alpha val="40000"/></a:prstClr></a:outerShdw></a:effectLst>')

# "Links": the values in B, a narrow cell in D, the linked shapes from column F.
LINK_ROWS = [f'<row r="1">{text_cell("A1", "Format")}{text_cell("B1", "Value")}</row>']
LINKS = []


def linked(id_, name, x, y, w, h, text, link):
    LINKS.append(anchor(x, y, w, h, shape(id_, name, x, y, w, h, text=text, link=link)))


for i, code in enumerate(LINK_FORMATS):
    r = i + 2
    value = '2297200.8603000008' if i != 1 else '-98765.4'
    LINK_ROWS.append(f'<row r="{r}">{text_cell(f"A{r}", code)}{num_cell(f"B{r}", value, XF[code])}'
                     + (num_cell(f'D{r}', '1234567.891', XF['#,##0.00']) if i == 0 else '') + '</row>')
    linked(10 + i, f'Linked {r}', 6 * COL, (i + 1) * 40, 200, 34, '$22,97,200.86', f'Links!$B${r}')
linked(20, 'Linked narrow', 6 * COL + 220, 40, 160, 34, 'stale', 'Links!$D$2')
linked(21, 'Linked other sheet', 6 * COL + 220, 80, 160, 34, '$2,86,397.02', "'Data'!$B$2")
LINKS_SHEET = sheet(LINK_ROWS, '<col min="1" max="1" width="30" customWidth="1"/><col min="2" max="2" width="18" customWidth="1"/>'
                    '<col min="4" max="4" width="3.7109375" customWidth="1"/>', drawing=True, view='tabSelected="1" ')
LINKS_DRAWING = DECL + f'<xdr:wsDr xmlns:xdr="{XDR}" xmlns:a="{A}">' + ''.join(LINKS) + '</xdr:wsDr>'

DATA_SHEET = sheet([f'<row r="2">{num_cell("B2", "286397.02169999998", XF["[$$-1009]#,##0.00"])}</row>'])

# "Colors": the code in A, the block in B; the KPI cells in D.
COLOR_ROWS = []
for i, code in enumerate(COLOR_FORMATS):
    r = i + 1
    extra = ''
    if i == 0:
        extra = num_cell('D1', '0.2062', XF[KPI_FORMAT])
    elif i == 1:
        extra = num_cell('D2', '-0.0515', XF[KPI_FORMAT])
    COLOR_ROWS.append(f'<row r="{r}" ht="21" customHeight="1">{text_cell(f"A{r}", code)}{num_cell(f"B{r}", "1", XF[code])}{extra}</row>')
COLORS_SHEET = sheet(COLOR_ROWS, '<col min="1" max="1" width="24" customWidth="1"/><col min="2" max="2" width="12" customWidth="1"/>'
                     '<col min="4" max="4" width="14" customWidth="1"/>')

# "Rounds": rounded rectangles (x, y, w, h, adj), line widths, shadows.
ROUNDS = [(20, 20, 200, 100, None), (240, 20, 200, 100, 6482), (460, 20, 200, 100, 3572), (680, 20, 100, 200, 10000),
          (800, 20, 300, 60, 25000), (800, 100, 120, 120, 50000), (20, 140, 220, 108, 6482), (260, 140, 110, 50, 16667)]
ROUND_SHAPES = [anchor(x, y, w, h, shape(30 + i, f'Round {i}', x, y, w, h, prst='roundRect', adj=adj)) for i, (x, y, w, h, adj) in enumerate(ROUNDS)]
WIDTHS = [9525, 12700, 15875, 19050, 25400, 38100]
WIDTH_SHAPES = [anchor(20 + i * 110, 280, 90, 60, shape(50 + i, f'Line {w}', 20 + i * 110, 280, 90, 60, line_w=w, line='000000')) for i, w in enumerate(WIDTHS)]
# The group: two cards with the shadow on the group, as the dashboard's (its own child space = its rectangle).
GX, GY, GW, GH = 20, 380, 460, 110
GROUP = anchor(GX, GY, GW, GH,
               f'<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="60" name="Shadow group"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>'
               f'<xdr:grpSpPr>{xfrm(GX, GY, GW, GH, child=(GX, GY, GW, GH))}{SHADOW}</xdr:grpSpPr>'
               + shape(61, 'Card 1', GX, GY, 220, GH, prst='roundRect', adj=6482, line_w=15875)
               + shape(62, 'Card 2', GX + 240, GY, 220, GH, prst='roundRect', adj=6482, line_w=15875)
               + '</xdr:grpSp>')
SINGLE = anchor(520, 380, 220, 110, shape(63, 'Shadowed card', 520, 380, 220, 110, prst='roundRect', adj=6482, line_w=15875, effects=SHADOW))
# "Rounds" too: lines (connectors) with arrowheads, as wide as given, none high; ellipses (an ellipse and a flow chart
# connector) with a wide light line, and a shape with the theme's effect (`effectRef idx="3"`: Office's default theme
# casts a shadow from its third effect style) and one whose own empty effect list overrides it.
def connector(id_, name, x, y, w, h, line_w, head=None, tail=None):
    ends = ''
    if head:
        ends += f'<a:headEnd type="{head[0]}" w="{head[1]}" len="{head[2]}"/>'
    if tail:
        ends += f'<a:tailEnd type="{tail[0]}" w="{tail[1]}" len="{tail[2]}"/>'
    return anchor(x, y, w, h,
                  f'<xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="{id_}" name="{name}"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr>'
                  f'<xdr:spPr>{xfrm(x, y, w, h)}<a:prstGeom prst="line"><a:avLst/></a:prstGeom>'
                  f'<a:ln w="{line_w}"><a:solidFill><a:srgbClr val="7F7F7F"/></a:solidFill>{ends}</a:ln></xdr:spPr></xdr:cxnSp>')


ARROWS = [(12700, None, ('triangle', 'med', 'lg')), (12700, ('triangle', 'sm', 'sm'), ('triangle', 'lg', 'lg')),
          (28575, None, ('triangle', 'med', 'med')), (9525, ('arrow', 'med', 'med'), ('stealth', 'med', 'med')),
          (12700, ('oval', 'med', 'med'), ('diamond', 'med', 'med'))]
LINE_SHAPES = [connector(70 + i, f'Arrow {i}', 820, 280 + i * 30, 240, 0, w, head, tail) for i, (w, head, tail) in enumerate(ARROWS)]
LINE_SHAPES.append(connector(76, 'Slant', 1100, 280, 100, 60, 19050, None, ('triangle', 'med', 'med')))


def ellipse(id_, name, x, y, w, h, prst, effects='', style=''):
    return anchor(x, y, w, h,
                  f'<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{id_}" name="{name}"/><xdr:cNvSpPr/></xdr:nvSpPr>'
                  f'<xdr:spPr>{xfrm(x, y, w, h)}<a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="398CCC"/></a:solidFill>'
                  f'<a:ln w="38100"><a:solidFill><a:srgbClr val="F2F2F2"/></a:solidFill></a:ln>{effects}</xdr:spPr>{style}</xdr:sp>')


THEME_STYLE = ('<xdr:style><a:lnRef idx="1"><a:schemeClr val="dk1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="dk1"/></a:fillRef>'
               '<a:effectRef idx="3"><a:schemeClr val="dk1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef></xdr:style>')
LINE_SHAPES += [ellipse(80, 'Ellipse', 820, 440, 40, 40, 'ellipse'), ellipse(81, 'Flow connector', 880, 440, 40, 40, 'flowChartConnector'),
                ellipse(82, 'Theme effect', 940, 440, 60, 40, 'rect', style=THEME_STYLE),
                ellipse(83, 'Theme effect off', 1020, 440, 60, 40, 'rect', effects='<a:effectLst/>', style=THEME_STYLE)]

# A shape of no fill, no line (none given, no style) and no text, as house-flipping's clickable rectangles: nothing.
LINE_SHAPES.append(anchor(1100, 380, 100, 60, shape(84, 'Empty', 1100, 380, 100, 60, line='', fill=None)))

ROUNDS_DRAWING = DECL + f'<xdr:wsDr xmlns:xdr="{XDR}" xmlns:a="{A}">' + ''.join(ROUND_SHAPES + WIDTH_SHAPES + [GROUP, SINGLE] + LINE_SHAPES) + '</xdr:wsDr>'
ROUNDS_SHEET = sheet([], drawing=True)

SHEETS = [('Links', LINKS_SHEET, LINKS_DRAWING), ('Colors', COLORS_SHEET, None), ('Rounds', ROUNDS_SHEET, ROUNDS_DRAWING), ('Data', DATA_SHEET, None)]
CT = ('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')
FILES = {}
drawings = 0
for i, (name, xml, drawing) in enumerate(SHEETS, 1):
    CT += f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    FILES[f'xl/worksheets/sheet{i}.xml'] = xml
    if drawing:
        drawings += 1
        CT += f'<Override PartName="/xl/drawings/drawing{drawings}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
        FILES[f'xl/worksheets/_rels/sheet{i}.xml.rels'] = DECL + (
            f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/drawing" Target="../drawings/drawing{drawings}.xml"/></Relationships>')
        FILES[f'xl/drawings/drawing{drawings}.xml'] = drawing
CT += '</Types>'
FILES = {
    '[Content_Types].xml': DECL + CT,
    '_rels/.rels': DECL + f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    'xl/workbook.xml': DECL + (f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><bookViews><workbookView/></bookViews><sheets>'
                               + ''.join(f'<sheet name="{name}" sheetId="{i}" r:id="rId{i}"/>' for i, (name, _, _) in enumerate(SHEETS, 1))
                               + '</sheets></workbook>'),
    'xl/_rels/workbook.xml.rels': DECL + (
        f'<Relationships xmlns="{PKG}">'
        + ''.join(f'<Relationship Id="rId{i}" Type="{REL}/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1, len(SHEETS) + 1))
        + f'<Relationship Id="rId{len(SHEETS) + 1}" Type="{REL}/styles" Target="styles.xml"/></Relationships>'),
    'xl/styles.xml': STYLES,
    **FILES,
}

with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
    for name, text in FILES.items():
        info = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
        info.compress_type = zipfile.ZIP_DEFLATED
        z.writestr(info, text.encode('utf-8'))
print(OUT)
