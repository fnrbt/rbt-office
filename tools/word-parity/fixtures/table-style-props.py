#!/usr/bin/env python3
"""Fixture family: which properties a table cell's paragraph takes from the table style, which from Normal.

    python3 tools/word-parity/fixtures/table-style-props.py [out dir]   (default tests/word/fixtures/table-style-props)

then, to ask Word (winvm) and compare with ours (Word's pages and PDFs in <out dir>/word):

    WORD_VM_DIR=C:/Users/olivia/RbtExcelParity/<own dir> \
      node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/table-style-props --refresh-word
    python3 tools/word-parity/pdf-geometry.py tests/word/fixtures/table-style-props/word/*.pdf > tests/word/fixtures/table-style-props/word/geometry.tsv

Hypotheses (docs/layout-parity-flow.md ledger): the style hierarchy is document defaults, table style,
paragraph style (Normal), direct formatting - so Normal's properties beat the table style's - except
that, unless the document's w:compatSetting overrideTableStyleFontSizeAndJustification is on, the
table style's font size (its own or inherited, down to the document defaults) beats Normal's when
Normal's is 11 or 12pt, and its justification beats Normal's when Normal's is left.

Every document: Letter, 1in margins, Calibri (docDefaults' rFonts) single-spaced, nothing before or
after, unless the name says; a paragraph "p01 ...", the table, a paragraph "q01 ...". The table: one
4in column, two rows, each cell two paragraphs of 30 words ("a01" ..., "b01" ... in row 1, "d01" ...,
"e01" ... in row 2) with no pPr or rPr of their own (unless the name says a paragraph style "P1"),
a table style "TS" (based on TableNormal), tblLook 0000 (no conditional formats) unless "cond".
Names: d<...> what docDefaults set, n<...> Normal, t<...> TS, cond<...> TS's firstRow (tblLook 0020),
s<...> P1 (based on Normal, the cells' paragraphs' style); sz<k> half-points / 2 ("sz12": 12pt),
fl<tw> firstLine, after<tw>, line<240ths>, jc<value>, ft<face>, b bold; -ov the compat setting
overrideTableStyleFontSizeAndJustification = 1; -c14 / -c12 compatibilityMode 14 / 12, -cnone no
compatibilityMode setting (else 15).
No dependencies: the package is written as XML.
"""
import os, sys, zipfile

out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), "../../../tests/word/fixtures/table-style-props")
os.makedirs(out, exist_ok=True)
NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
SINGLE = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'


def words(prefix, count):
    return " ".join(f"{prefix}{k:02d}" for k in range(1, count + 1))


def run(text):
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def para(text, ppr=""):
    return f'<w:p>{"<w:pPr>" + ppr + "</w:pPr>" if ppr else ""}{run(text)}</w:p>'


def table(cell_ppr="", look="0000"):
    def cell(a, b):
        return (f'<w:tc><w:tcPr><w:tcW w:w="5760" w:type="dxa"/></w:tcPr>'
                f'{para(words(a, 30), cell_ppr)}{para(words(b, 30), cell_ppr)}</w:tc>')
    borders = ('<w:tblBorders>' + "".join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                                         for s in ("top", "left", "bottom", "right", "insideH", "insideV")) + '</w:tblBorders>')
    return (f'<w:tbl><w:tblPr><w:tblStyle w:val="TS"/><w:tblW w:w="0" w:type="auto"/>{borders}<w:tblLook w:val="{look}"/></w:tblPr>'
            f'<w:tblGrid><w:gridCol w:w="5760"/></w:tblGrid><w:tr>{cell("a", "b")}</w:tr><w:tr>{cell("d", "e")}</w:tr></w:tbl>')


def document(body):
    sect = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
            'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{body}{sect}</w:body></w:document>'


def styles(d_ppr=SINGLE, d_rpr='<w:sz w:val="22"/><w:szCs w:val="22"/>', n_ppr="", n_rpr="", t_ppr="", t_rpr="",
           c_ppr=None, c_rpr="", s_ppr=None, s_rpr="", t_tblpr="", c_raw=""):
    fonts = '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>'
    cond = (f'<w:tblStylePr w:type="firstRow"><w:pPr>{c_ppr}</w:pPr><w:rPr>{c_rpr}</w:rPr><w:tblPr/></w:tblStylePr>'
            if c_ppr is not None else "")
    p1 = (f'<w:style w:type="paragraph" w:styleId="P1"><w:name w:val="P1"/><w:basedOn w:val="Normal"/>'
          f'<w:pPr>{s_ppr}</w:pPr><w:rPr>{s_rpr}</w:rPr></w:style>' if s_ppr is not None else "")
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + '>'
            f'<w:docDefaults><w:rPrDefault><w:rPr>{fonts}{d_rpr}</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>{d_ppr}</w:pPr></w:pPrDefault></w:docDefaults>'
            f'<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr>{n_ppr}</w:pPr><w:rPr>{n_rpr}</w:rPr></w:style>'
            '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>'
            '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/>'
            '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>'
            '<w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
            f'<w:style w:type="table" w:styleId="TS"><w:name w:val="TS"/><w:basedOn w:val="TableNormal"/>'
            + ('' if t_ppr is None else f'<w:pPr>{t_ppr}</w:pPr>') + ('' if t_rpr is None else f'<w:rPr>{t_rpr}</w:rPr>')
            + f'<w:tblPr>{t_tblpr}</w:tblPr>{cond}{c_raw}</w:style>'
            f'{p1}</w:styles>')


def write(name, compat=15, override=False, cell_ppr="", look="0000", **style_args):
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
    ov = ('<w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
          if override else "")
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + NS + '><w:compat>' +
                (f'<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{compat}"/>' if compat else '')
                + f'{ov}</w:compat></w:settings>')
    body = para(words("p", 20)) + table(cell_ppr, look) + para(words("q", 20))
    with zipfile.ZipFile(os.path.join(out, name + ".docx"), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", rels)
        z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/document.xml", document(body))
        z.writestr("word/styles.xml", styles(**style_args)); z.writestr("word/settings.xml", settings)


def sz(points):
    return f'<w:sz w:val="{round(points * 2)}"/>'


def dsz(points):
    return f'<w:sz w:val="{round(points * 2)}"/><w:szCs w:val="{round(points * 2)}"/>'


def jc(value):
    return f'<w:jc w:val="{value}"/>'


def fl(tw):
    return f'<w:ind w:firstLine="{tw}"/>'


def after(tw):
    return f'<w:spacing w:after="{tw}"/>'


def line(v):
    return f'<w:spacing w:line="{v}" w:lineRule="auto"/>'


def face(name):
    return f'<w:rFonts w:ascii="{name}" w:hAnsi="{name}"/>'


S1 = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
P1 = '<w:pStyle w:val="P1"/>'

# Font size: Normal's 11 / 12pt against the table style's (or its inherited docDefaults'), other sizes.
write("nsz12-t", n_rpr=sz(12))                                      # calibre's case: TS no size
write("nsz12-tsz8", n_rpr=sz(12), t_rpr=sz(8))
write("nsz11-tsz8", n_rpr=sz(11), t_rpr=sz(8))
write("dsz10-nsz11-tsz8", d_rpr=dsz(10), n_rpr=sz(11), t_rpr=sz(8))
write("dsz10-nsz12-t", d_rpr=dsz(10), n_rpr=sz(12))
write("dsz10-nsz12-tsz14", d_rpr=dsz(10), n_rpr=sz(12), t_rpr=sz(14))
write("nsz16-tsz8", n_rpr=sz(16), t_rpr=sz(8))
write("dsz16-nsz16-tsz8", d_rpr=dsz(16), n_rpr=sz(16), t_rpr=sz(8))
write("nsz10.5-tsz8", n_rpr=sz(10.5), t_rpr=sz(8))
write("nsz11.5-tsz8", n_rpr=sz(11.5), t_rpr=sz(8))
write("nsz13-tsz8", n_rpr=sz(13), t_rpr=sz(8))
write("nsz12-tsz8-ov", override=True, n_rpr=sz(12), t_rpr=sz(8))
write("nsz12-t-ov", override=True, n_rpr=sz(12))
write("nsz12-tsz8-c14", compat=14, n_rpr=sz(12), t_rpr=sz(8))
write("nsz12-tsz8-ov-c14", compat=14, override=True, n_rpr=sz(12), t_rpr=sz(8))
write("nsz12-condsz8", look="0020", n_rpr=sz(12), c_ppr="", c_rpr=sz(8))
write("nsz16-condsz8", look="0020", n_rpr=sz(16), c_ppr="", c_rpr=sz(8))
write("nsz12-tsz14-condsz8", look="0020", n_rpr=sz(12), t_rpr=sz(14), c_ppr="", c_rpr=sz(8))
write("ssz12-tsz8", cell_ppr=P1, s_ppr="", s_rpr=sz(12), t_rpr=sz(8))
write("nsz12-s-tsz8", cell_ppr=P1, s_ppr="", n_rpr=sz(12), t_rpr=sz(8))
write("ssz16-nsz12-tsz8", cell_ppr=P1, s_ppr="", s_rpr=sz(16), n_rpr=sz(12), t_rpr=sz(8))
write("nszcs12-tsz8", n_rpr='<w:szCs w:val="24"/>', t_rpr=sz(8))
# Justification.
write("njcleft-tjccenter", n_ppr=jc("left"), t_ppr=jc("center"))
write("njcboth-tjccenter", n_ppr=jc("both"), t_ppr=jc("center"))
write("njcright-tjccenter", n_ppr=jc("right"), t_ppr=jc("center"))
write("njcstart-tjccenter", n_ppr=jc("start"), t_ppr=jc("center"))
write("njcleft-tjccenter-ov", override=True, n_ppr=jc("left"), t_ppr=jc("center"))
write("njcleft-tjccenter-c14", compat=14, n_ppr=jc("left"), t_ppr=jc("center"))
write("djcright-njcleft-t", d_ppr=SINGLE + jc("right"), n_ppr=jc("left"))
write("njcleft-condjccenter", look="0020", n_ppr=jc("left"), c_ppr=jc("center"))
# First-line indent, spacing, line spacing.
write("nfl432-tfl864", n_ppr=fl(432), t_ppr=fl(864))
write("dfl432-nfl432-tfl864", d_ppr=SINGLE + fl(432), n_ppr=fl(432), t_ppr=fl(864))
write("nfl432-tfl864-c14", compat=14, n_ppr=fl(432), t_ppr=fl(864))
write("nfl0-tfl864", n_ppr=fl(0), t_ppr=fl(864))
write("nfl432-condfl864", look="0020", n_ppr=fl(432), c_ppr=fl(864))
write("nafter240-tafter0", n_ppr=after(240), t_ppr=after(0))
write("nafter0-tafter240", n_ppr=after(0), t_ppr=after(240))
write("nline480-tline240", n_ppr=line(480), t_ppr=line(240))
write("nline240-tline480", n_ppr=line(240), t_ppr=line(480))
write("dline276-tline240", d_ppr='<w:spacing w:before="0" w:after="0" w:line="276" w:lineRule="auto"/>', t_ppr=line(240))
# Run properties other than the size.
write("nfttimes-tftarial", n_rpr=face("Times New Roman"), t_rpr=face("Arial"))
write("nftcalibri-tftarial", n_rpr=face("Calibri"), t_rpr=face("Arial"))
write("tb", t_rpr="<w:b/>")
write("nb0-tb", n_rpr='<w:b w:val="0"/>', t_rpr="<w:b/>")
# calibre-demo's own: Normal firstLine 0.3in, Ubuntu-less (Calibri) 12pt; docDefaults 1.15 lines; TS single.
write("calibre", compat=14, d_ppr='<w:spacing w:line="276" w:lineRule="auto"/>', n_ppr=fl(432), n_rpr=sz(12), t_ppr=line(240))

# Batch 2: Word 2010's layout (compatibilityMode 14) and earlier - which sizes and justifications of
# Normal lose to the table style's, from where; the other properties; a paragraph style other than Normal.
C14 = dict(compat=14)
write("nsz11-tsz8-c14", **C14, n_rpr=sz(11), t_rpr=sz(8))
write("nsz11.5-tsz8-c14", **C14, n_rpr=sz(11.5), t_rpr=sz(8))
write("nsz13-tsz8-c14", **C14, n_rpr=sz(13), t_rpr=sz(8))
write("nsz10.5-tsz8-c14", **C14, n_rpr=sz(10.5), t_rpr=sz(8))
write("nsz10-tsz8-c14", **C14, n_rpr=sz(10), t_rpr=sz(8))
write("nsz12-t-c14", **C14, n_rpr=sz(12))
write("dsz10-nsz12-t-c14", **C14, d_rpr=dsz(10), n_rpr=sz(12))
write("dsz10-nsz11-tsz8-c14", **C14, d_rpr=dsz(10), n_rpr=sz(11), t_rpr=sz(8))
write("dsz12-nsz11-t-c14", **C14, d_rpr=dsz(12), n_rpr=sz(11))
write("dsz12-t-tsz8-c14", **C14, d_rpr=dsz(12), t_rpr=sz(8))
write("nsz12-condsz8-c14", **C14, look="0020", n_rpr=sz(12), c_ppr="", c_rpr=sz(8))
write("nsz12-tsz14-condsz8-c14", **C14, look="0020", n_rpr=sz(12), t_rpr=sz(14), c_ppr="", c_rpr=sz(8))
write("nsz16-condsz8-c14", **C14, look="0020", n_rpr=sz(16), c_ppr="", c_rpr=sz(8))
write("ssz12-tsz8-c14", **C14, cell_ppr=P1, s_ppr="", s_rpr=sz(12), t_rpr=sz(8))
write("nsz12-s-tsz8-c14", **C14, cell_ppr=P1, s_ppr="", n_rpr=sz(12), t_rpr=sz(8))
write("nszcs12-tsz8-c14", **C14, n_rpr='<w:szCs w:val="24"/>', t_rpr=sz(8))
write("njcboth-tjccenter-c14", **C14, n_ppr=jc("both"), t_ppr=jc("center"))
write("njcright-tjccenter-c14", **C14, n_ppr=jc("right"), t_ppr=jc("center"))
write("njcstart-tjccenter-c14", **C14, n_ppr=jc("start"), t_ppr=jc("center"))
write("djcright-njcleft-t-c14", **C14, d_ppr=SINGLE + jc("right"), n_ppr=jc("left"))
write("njcleft-condjccenter-c14", **C14, look="0020", n_ppr=jc("left"), c_ppr=jc("center"))
write("sjcleft-tjccenter-c14", **C14, cell_ppr=P1, s_ppr=jc("left"), t_ppr=jc("center"))
write("njcleft-tjccenter-ov-c14", **C14, override=True, n_ppr=jc("left"), t_ppr=jc("center"))
write("nline480-tline240-c14", **C14, n_ppr=line(480), t_ppr=line(240))
write("nafter240-tafter0-c14", **C14, n_ppr=after(240), t_ppr=after(0))
write("nfttimes-tftarial-c14", **C14, n_rpr=face("Times New Roman"), t_rpr=face("Arial"))
write("nsz12-tsz8-c12", compat=12, n_rpr=sz(12), t_rpr=sz(8))
write("njcleft-tjccenter-c12", compat=12, n_ppr=jc("left"), t_ppr=jc("center"))
write("nsz12-tsz8-cnone", compat=None, n_rpr=sz(12), t_rpr=sz(8))
write("njcleft-tjccenter-cnone", compat=None, n_ppr=jc("left"), t_ppr=jc("center"))

# Batch 3: Normal's 12pt in Word 2010's layout - lost when the table style has properties of any kind
# (calibre: a line spacing only), kept when it has none (nsz12-t-c14)?
TLINE = line(240)
write("nsz12-tline-c14", **C14, n_rpr=sz(12), t_ppr=TLINE)
write("nsz12-tline", n_rpr=sz(12), t_ppr=TLINE)
write("nsz12-tline-ov-c14", **C14, override=True, n_rpr=sz(12), t_ppr=TLINE)
write("dsz10-nsz12-tline-c14", **C14, d_rpr=dsz(10), n_rpr=sz(12), t_ppr=TLINE)
write("nsz12-tb-c14", **C14, n_rpr=sz(12), t_rpr="<w:b/>")
write("nsz12-tcolor-c14", **C14, n_rpr=sz(12), t_rpr='<w:color w:val="FF0000"/>')
write("nsz12-tfl-c14", **C14, n_rpr=sz(12), t_ppr=fl(0))
write("nsz12-tnone-c14", **C14, n_rpr=sz(12), t_ppr=None, t_rpr=None)
write("dline276-nsz12-t-c14", **C14, d_ppr='<w:spacing w:line="276" w:lineRule="auto"/>', n_rpr=sz(12))
write("ssz12-tline-c14", **C14, cell_ppr=P1, s_ppr="", s_rpr=sz(12), t_ppr=TLINE)
write("nsz12-s-tline-c14", **C14, cell_ppr=P1, s_ppr="", n_rpr=sz(12), t_ppr=TLINE)
write("nsz11-tline-c14", **C14, d_rpr=dsz(10), n_rpr=sz(11), t_ppr=TLINE)
write("djcright-njcleft-tline-c14", **C14, d_ppr=SINGLE + jc("right"), n_ppr=jc("left"), t_ppr=TLINE)
write("djcright-njcleft-tline", d_ppr=SINGLE + jc("right"), n_ppr=jc("left"), t_ppr=TLINE)
write("djccenter-njcleft-tb-c14", **C14, d_ppr=SINGLE + jc("center"), n_ppr=jc("left"), t_rpr="<w:b/>")

# Batch 4: what Normal's 12pt gives way to with document defaults other than 11pt.
for d, t in ((10, 8), (10, 10), (10, 11), (9, None), (14, None), (10.5, None), (12, None)):
    tname = f"tsz{t:g}" if t else "tline"
    write(f"dsz{d:g}-nsz12-{tname}-c14", **C14, d_rpr=dsz(d), n_rpr=sz(12), t_ppr=TLINE, t_rpr=sz(t) if t else "")
write("njcleft-s-tjccenter-c14", **C14, cell_ppr=P1, s_ppr="", n_ppr=jc("left"), t_ppr=jc("center"))

# Batch 5: a table style with table or cell properties only - does Normal's 12pt give way?
write("nsz12-ttblonly-c14", **C14, n_rpr=sz(12), t_ppr=None, t_rpr=None,
      t_tblpr='<w:tblBorders><w:top w:val="single" w:sz="12" w:space="0" w:color="FF0000"/></w:tblBorders>')
write("nsz12-tcondshd-c14", **C14, look="0020", n_rpr=sz(12), t_ppr=None, t_rpr=None,
      c_raw='<w:tblStylePr w:type="firstRow"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/></w:tcPr></w:tblStylePr>')
write("nsz12-condsz8-nolook-c14", **C14, look="0200", n_rpr=sz(12), c_ppr="", c_rpr=sz(8))
write("nsz12-condlastshd-c14", **C14, look="0240", n_rpr=sz(12), t_ppr=None, t_rpr=None,
      c_raw='<w:tblStylePr w:type="lastRow"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/></w:tcPr></w:tblStylePr>')
write("djccenter-njcleft-tcondshd-c14", **C14, look="0020", d_ppr=SINGLE + jc("center"), n_ppr=jc("left"), t_ppr=None, t_rpr=None,
      c_raw='<w:tblStylePr w:type="firstRow"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/></w:tcPr></w:tblStylePr>')
print(out)
