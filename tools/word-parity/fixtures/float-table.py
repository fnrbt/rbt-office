#!/usr/bin/env python3
"""Fixture family: how Word places a floating table (w:tblpPr) and wraps text round it.

    python3 tools/word-parity/fixtures/float-table.py [out dir]   (default tests/word/fixtures/float-table)

then, to ask Word (winvm) and compare with ours (Word's pages and PDFs in <out dir>/word):

    WORD_VM_DIR=C:/Users/olivia/RbtExcelParity/<own dir> \
      node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/float-table --refresh-word
    python3 tools/word-parity/pdf-geometry.py tests/word/fixtures/float-table/word/*.pdf > tests/word/fixtures/float-table/word/geometry.tsv

tests/word/float-table.fsx checks ours against geometry.tsv (each line's pieces: page, top, left,
right, first and last word - the table's cells' lines among them, "t11" ... in row 1, column 1).

Hypotheses (docs/layout-parity-flow.md ledger): a floating table is placed from the paragraph after
it (its anchor) as a picture is from its paragraph - vertAnchor text: that paragraph's top plus tblpY,
margin / page: from there; horzAnchor (text, margin, page) plus tblpX or tblpXSpec - and the text goes
round it as round a square-wrapped picture on both sides: every line whose box reaches into the
table's box widened by its from-text distances goes beside it, in a band a quarter inch wide or more.

Every document: Letter, 1in margins, Calibri 11pt single-spaced (13.43pt lines), nothing before or
after, justified paragraphs of numbered words: A (2 lines, "a01 ..."), the table, B (its anchor,
~14 lines, "b001 ..."), C ("c01 ..."). The table: 2 columns of 1in (tcW 1440, autofit), 4 rows of one
line ("t11" "t12" / "t21" ...), single 0.5pt borders, Word's cell margins (5.4pt left and right). Its
w:tblpPr has only what the name says (vertAnchor text, tblpY 0 unless it says):

  base, base-c14                no other attribute (horzAnchor, tblpX, distances omitted)
  cal, cal-c14                  calibre-demo's: rightFromText 187, bottomFromText 72, tblpY 1, tblOverlap never
  y<k>, ym<k>                   tblpY k pt (ym: -k)
  h<text|margin|page>-x<k>      horzAnchor with tblpX k pt
  ind-h<text|margin>            B (the anchor) indented 1in, horzAnchor text / margin, tblpX 0
  xs-<left|center|right|inside|outside>[-page]   tblpXSpec, horzAnchor margin (page)
  vm-y<k>, vp-y<k>              vertAnchor margin / page, tblpY k pt
  ys-<top|center|bottom>[-text] tblpYSpec, vertAnchor margin (text)
  d-<l|r|t|b><k>                that from-text distance k pt (d-l*: the table at the right, tblpXSpec right;
                                d-t*: tblpY 3 lines)
  gap-<g>, gapr-<g>             horzAnchor margin; the table's left edge g pt right of the margin (gapr: its
                                right edge g pt left of the right margin) - the least band
  ov-<never|overlap>            A, a 6-row floating table T1 anchored on a short paragraph P, a second table T2
                                anchored on B, both at tblpY 0, tblpX 0: tblOverlap never / overlap on T2
  pb, pb-big                    B's first line 4 lines above the page's foot (44 one-line paragraphs before);
                                pb-big: 30 rows
  short, empty                  B 20 words (short) / an empty paragraph (empty), then C: C beside it too?
  sb12                          B with 12pt space before
  wide                          a 5in table (2 x 2.5in)
  cell                          the floating table and B in a full-width 1 x 1 table's cell
  nested                        the floating table in a floating table's cell (vertAnchor text, its anchor
                                paragraph after it in the cell), the outer one at the margin, 4in wide
  fl-b, fl-normal               a 0.3in first-line indent on B / on Normal (every paragraph, the cells' too)
  fl-normal-inline              Normal's first-line indent over an inline (not floating) table
  fl-normal-cal-c14             calibre-demo's case: Normal's first-line indent, cal's tblpPr, Word 2010's layout
  gapil-<g>                     the table's left edge g pt in from the margin, B of 5pt words ("il")
  <name>-c14                    hmargin-x72, xs-right, xs-center, d-r18 in Word 2010's layout
  sz12, sz12-c14                1.5pt borders
  vm-y0                         vertAnchor margin, tblpY 0: above A
  hmargin-xm36                  horzAnchor margin, tblpX -36pt: into the left margin
  ts-<fl|fl-both|after|after-only|line|font|font-only>   an inline table with a table style "TS": Normal's
                                first-line indent, space after, line spacing or font size against TS's
  ts-fl-float                   ts-fl's table floating
  va-none-y100                  no vertAnchor, tblpY 100pt
No dependencies: the package is written as XML.
"""
import os, sys, zipfile

out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), "../../../tests/word/fixtures/float-table")
os.makedirs(out, exist_ok=True)
NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
LINE = 13.4307
SINGLE = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
IN = 72.0


def tw(pt):
    return str(round(pt * 20))


def words(prefix, count, digits):
    return " ".join(f"{prefix}{k:0{digits}d}" for k in range(1, count + 1))


def run(text):
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def para(content, ppr="", jc="both"):
    return f'<w:p><w:pPr>{ppr}{SINGLE}<w:jc w:val="{jc}"/></w:pPr>{content}</w:p>'


def borders(sz=4):
    return ('<w:tblBorders>' + "".join(f'<w:{s} w:val="single" w:sz="{sz}" w:space="0" w:color="000000"/>'
                                       for s in ("top", "left", "bottom", "right", "insideH", "insideV")) + '</w:tblBorders>')


def table(pp=None, rows=4, cols=2, width=IN, prefix="t", overlap=None, cell=None, sz=4, style=None):
    """pp: the w:tblpPr's attributes (None: not floating); cell(r, c): a cell's content."""
    tblpp = ("<w:tblpPr " + " ".join(f'w:{k}="{v}"' for k, v in pp.items()) + "/>") if pp is not None else ""
    ov = f'<w:tblOverlap w:val="{overlap}"/>' if overlap else ""
    grid = "".join(f'<w:gridCol w:w="{tw(width)}"/>' for _ in range(cols))
    body = ""
    for r in range(1, rows + 1):
        body += "<w:tr>"
        for c in range(1, cols + 1):
            content = cell(r, c) if cell else f'<w:p><w:pPr>{SINGLE}</w:pPr>{run(f"{prefix}{r}{c}")}</w:p>'
            body += f'<w:tc><w:tcPr><w:tcW w:w="{tw(width)}" w:type="dxa"/></w:tcPr>{content}</w:tc>'
        body += "</w:tr>"
    ts = f'<w:tblStyle w:val="{style}"/>' if style else ""
    return (f'<w:tbl><w:tblPr>{ts}{tblpp}{ov}<w:tblW w:w="0" w:type="auto"/>{borders(sz)}<w:tblLook w:val="0000"/></w:tblPr>'
            f'<w:tblGrid>{grid}</w:tblGrid>{body}</w:tbl>')


def ft(**attrs):
    """A floating table, vertAnchor text, tblpY 0 unless given; values in pt (tw) or strings."""
    rows = attrs.pop("rows", 4)
    width = attrs.pop("width", IN)
    overlap = attrs.pop("overlap", None)
    prefix = attrs.pop("prefix", "t")
    sz = attrs.pop("sz", 4)
    pp = {"vertAnchor": "text", "tblpY": "0"}
    for k, v in attrs.items():
        pp[k] = tw(v) if isinstance(v, (int, float)) else v
    pp = {k: v for k, v in pp.items() if v is not None}
    return table(pp, rows=rows, width=width, overlap=overlap, prefix=prefix, sz=sz)


A = para(run(words("a", 40, 2)))
B = para(run(words("b", 200, 3)))
C = para(run(words("c", 60, 2)))


def lines_of(n, prefix="a"):
    return "".join(para(run(f"{prefix}{k:03d}")) for k in range(1, n + 1))


def document(body):
    sect = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
            'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{body}{sect}</w:body></w:document>'


def styles(normal_ppr="", normal_rpr="", ts_ppr=None, ts_rpr=""):
    """Normal's pPr / rPr; ts_ppr: a table style "TS" (based on TableNormal) with that pPr and rPr."""
    xml = STYLES.replace('<w:name w:val="Normal"/></w:style>',
                         f'<w:name w:val="Normal"/><w:pPr>{normal_ppr}</w:pPr><w:rPr>{normal_rpr}</w:rPr></w:style>')
    if ts_ppr is not None:
        xml = xml.replace('</w:styles>', '<w:style w:type="table" w:styleId="TS"><w:name w:val="TS"/><w:basedOn w:val="TableNormal"/>'
                          f'<w:pPr>{ts_ppr}</w:pPr><w:rPr>{ts_rpr}</w:rPr><w:tblPr/></w:style></w:styles>')
    return xml


STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>'
          '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>' + SINGLE + '</w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
          '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>'
          '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/>'
          '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>'
          '<w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>')


def write(name, body, compat=15, style_xml=None):
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
             '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
    compat_xml = (f'<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{compat}"/></w:compat>'
                  if compat else "")
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                f'{compat_xml}</w:settings>')
    with zipfile.ZipFile(os.path.join(out, name + ".docx"), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", rels)
        z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/document.xml", document(body))
        z.writestr("word/styles.xml", style_xml or STYLES); z.writestr("word/settings.xml", settings)


CAL = dict(rightFromText="187", bottomFromText="72", tblpY="1", overlap="never")

for compat in (15, 14):
    sfx = "" if compat == 15 else f"-c{compat}"
    write("base" + sfx, A + ft() + B + C, compat)
    write("cal" + sfx, A + ft(**CAL) + B + C, compat)

for k in (20, 100):
    write(f"y{k}", A + ft(tblpY=k) + B + C)
write("ym20", A + ft(tblpY=-20) + B + C)

for anchor, k in (("text", 72), ("margin", 72), ("page", 144)):
    write(f"h{anchor}-x{k}", A + ft(horzAnchor=anchor, tblpX=k) + B + C)
for anchor in ("text", "margin"):
    write(f"ind-h{anchor}", A + ft(horzAnchor=anchor, tblpX=0) + para(run(words("b", 200, 3)), '<w:ind w:left="1440"/>') + C)

for spec in ("left", "center", "right", "inside", "outside"):
    write(f"xs-{spec}", A + ft(horzAnchor="margin", tblpXSpec=spec) + B + C)
for spec in ("center", "right"):
    write(f"xs-{spec}-page", A + ft(horzAnchor="page", tblpXSpec=spec) + B + C)

write("vm-y100", A + ft(vertAnchor="margin", tblpY=100) + B + C)
write("vp-y200", A + ft(vertAnchor="page", tblpY=200) + B + C)
for spec in ("top", "center", "bottom"):
    write(f"ys-{spec}", A + ft(vertAnchor="margin", tblpYSpec=spec) + B + C)
    write(f"ys-{spec}-text", A + ft(tblpYSpec=spec) + B + C)

for side, attr in (("l", "leftFromText"), ("r", "rightFromText"), ("t", "topFromText"), ("b", "bottomFromText")):
    extra = {}
    if side == "l":
        extra = dict(horzAnchor="margin", tblpXSpec="right")
    if side == "t":
        extra = dict(tblpY=3 * LINE)
    write(f"d-{side}18", A + ft(**{attr: 18}, **extra) + B + C)

# The least band: the table's left edge g pt in from the margin (its border; its cells' text 5.4pt further).
for g in (12, 16, 18, 20, 24, 36):
    write(f"gap-{g}", A + ft(horzAnchor="margin", tblpX=g) + B + C)
    write(f"gapr-{g}", A + ft(horzAnchor="margin", tblpX=6.5 * IN - 2 * IN - g) + B + C)

P = para(run(words("p", 10, 2)))
for ov in ("never", "overlap"):
    write(f"ov-{ov}", A + ft(rows=6, prefix="u") + P + ft(overlap=ov) + B + C)

write("pb", lines_of(44) + ft(rows=6) + B + C)
write("pb-big", lines_of(44) + ft(rows=30) + B + C)
write("short", A + ft(rows=6) + para(run(words("b", 20, 3))) + C)
write("empty", A + ft(rows=6) + para("") + C)
SB12 = ('<w:p><w:pPr><w:spacing w:before="240" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr>'
        + run(words("b", 200, 3)) + '</w:p>')
write("sb12", A + ft() + SB12 + C)
write("wide", A + ft(width=2.5 * IN) + B + C)

CELLB = para(run(words("b", 120, 3)))
cell_body = lambda r, c: ft() + CELLB
write("cell", A + table(None, rows=1, cols=1, width=6.5 * IN, cell=cell_body) + C)
inner = lambda r, c: ft(prefix="n") + para(run(words("i", 40, 2)))
write("nested", A + table({"vertAnchor": "text", "tblpY": "0"}, rows=1, cols=1, width=4 * IN, cell=inner) + B + C)

# Batch 2. A first-line indent (calibre-demo's Normal has 0.3in): the paragraph beside the table, every
# paragraph (the cells' too), an inline table under it; the least band in 5pt words ("il"); Word 2010's
# layout across; thicker borders; the table above its paragraph (margin anchor) and left of the margin.
FL = '<w:ind w:firstLine="432"/>'
write("fl-b", A + ft() + para(run(words("b", 200, 3)), FL) + C)
write("fl-normal", A + ft() + B + C, style_xml=styles(FL))
write("fl-normal-inline", A + table(None) + B + C, style_xml=styles(FL))
write("fl-normal-cal-c14", A + ft(**CAL) + B + C, compat=14, style_xml=styles(FL))
IL = para(run(" ".join(["il"] * 700)))
for g in (16, 18, 24):
    write(f"gapil-{g}", A + ft(horzAnchor="margin", tblpX=g) + IL + C)
for name, kw in (("hmargin-x72", dict(horzAnchor="margin", tblpX=72)), ("xs-right", dict(horzAnchor="margin", tblpXSpec="right")),
                 ("xs-center", dict(horzAnchor="margin", tblpXSpec="center")), ("d-r18", dict(rightFromText=18))):
    write(f"{name}-c14", A + ft(**kw) + B + C, compat=14)
write("sz12", A + ft(sz=12) + B + C)
write("sz12-c14", A + ft(sz=12) + B + C, compat=14)
write("vm-y0", A + ft(vertAnchor="margin", tblpY=0) + B + C)
write("hmargin-xm36", A + ft(horzAnchor="margin", tblpX=-36) + B + C)

# Batch 3. A cell's paragraph without a style in a table with a table style: Normal's properties or the
# table style's (calibre-demo: Normal's first-line indent in its cells); no vertAnchor.
SP = lambda after=0, line=240: f'<w:spacing w:before="0" w:after="{after}" w:line="{line}" w:lineRule="auto"/>'
TS_TABLE = table(None, style="TS")
write("ts-fl", A + TS_TABLE + B + C, style_xml=styles(FL, ts_ppr=SP()))
write("ts-fl-both", A + TS_TABLE + B + C, style_xml=styles(FL, ts_ppr=SP() + '<w:ind w:firstLine="864"/>'))
write("ts-after", A + TS_TABLE + B + C, style_xml=styles(SP(240), ts_ppr=SP(0)))
write("ts-after-only", A + TS_TABLE + B + C, style_xml=styles(SP(240), ts_ppr=""))
write("ts-line", A + TS_TABLE + B + C, style_xml=styles(SP(0, 480), ts_ppr=SP(0, 240)))
write("ts-font", A + TS_TABLE + B + C, style_xml=styles(normal_rpr='<w:sz w:val="32"/>', ts_ppr="", ts_rpr='<w:sz w:val="16"/>'))
write("ts-font-only", A + TS_TABLE + B + C, style_xml=styles(normal_rpr='<w:sz w:val="32"/>', ts_ppr=""))
write("ts-fl-float", A + table({"vertAnchor": "text", "tblpY": "0"}, style="TS") + B + C, style_xml=styles(FL, ts_ppr=SP()))
write("va-none-y100", A + table({"tblpY": tw(100)}) + B + C)
print(out)
