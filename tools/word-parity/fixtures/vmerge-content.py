#!/usr/bin/env python3
"""Fixture family: what a vertically merged cell's continuation cells (w:vMerge without restart) hold, and
the rows' heights (annual-results' Pipeline table: each continuation cell holds an empty w:p of Normal's
spacing - 6pt after, 1.15 lines - and ours made those rows taller than Word's).

    python3 tools/word-parity/fixtures/vmerge-content.py [out dir]   (default tests/word/fixtures/vmerge-content)

then, to ask Word (winvm) and read where its PDFs put each line and border:

    WORD_VM_DIR=C:/Users/olivia/RbtExcelParity/<own dir> \\
      node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/vmerge-content --refresh-word
    python3 tools/word-parity/pdf-geometry.py --paths tests/word/fixtures/vmerge-content/word/*.pdf > tests/word/fixtures/vmerge-content/word/geometry.tsv

tests/word/vmerge-content.fsx checks ours against geometry.tsv.

Every document: Letter, 1in margins, Calibri 11pt; Normal 10pt after, 1.15 lines (docDefaults single, none).
a01, a 3 x 3 table (columns 1.5in, 2in, 1.5in, single 0.5pt borders, Word's cell margins), c01. Column 1 is
one cell merged down the three rows ("m11", restart) unless the name says; columns 2 and 3 "t21" ... in
paragraphs with nothing after, single-spaced. The continuation cells (rows 2 and 3 of column 1):

  empty, empty-center, empty-c14   an empty w:p (Normal's spacing); -center: the merged cell vAlign center
  none-spacing                     an empty w:p with no spacing (Normal's spacing removed for it)
  text                             "x21" / "x31" in a w:p without spacing: drawn? where?
  tall                             row 2's: 5 paragraphs "x21" ... "x25" without spacing - does row 2 grow?
  restart-tall                     the merged cell 6 lines, continuation cells empty w:p without spacing
  restart-tall-empty               restart-tall with the continuation cells' empty w:p of Normal's spacing
  two-merged                       columns 1 and 3 both merged down, empty Normal w:p in both continuations
"""
import os, sys, zipfile

out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), "../../../tests/word/fixtures/vmerge-content")
os.makedirs(out, exist_ok=True)
NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
TIGHT = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
WIDTHS = (2160, 2880, 2160)


def p(text="", ppr=TIGHT):
    r = f'<w:r><w:t>{text}</w:t></w:r>' if text else ""
    return f'<w:p><w:pPr>{ppr}</w:pPr>{r}</w:p>' if ppr else f'<w:p>{r}</w:p>'


def tc(width, content, merge=None, valign=None):
    vm = '' if merge is None else ('<w:vMerge w:val="restart"/>' if merge == "restart" else '<w:vMerge/>')
    va = f'<w:vAlign w:val="{valign}"/>' if valign else ""
    return f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{vm}{va}</w:tcPr>{content}</w:tc>'


def table(first, cont, valign=None, third_merged=False, third_cont=None):
    """first: the merged cell's content; cont(r): row r's continuation cell's content (r = 2, 3)."""
    rows = ""
    for r in (1, 2, 3):
        c1 = tc(WIDTHS[0], first if r == 1 else cont(r), "restart" if r == 1 else "cont", valign)
        c2 = tc(WIDTHS[1], p(f"t{r}2"))
        if third_merged:
            c3 = tc(WIDTHS[2], p("m13") if r == 1 else third_cont(r), "restart" if r == 1 else "cont")
        else:
            c3 = tc(WIDTHS[2], p(f"t{r}3"))
        rows += f"<w:tr>{c1}{c2}{c3}</w:tr>"
    borders = "".join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>' for s in ("top", "left", "bottom", "right", "insideH", "insideV"))
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in WIDTHS)
    return (f'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>{borders}</w:tblBorders><w:tblLayout w:type="fixed"/>'
            f'<w:tblLook w:val="0000"/></w:tblPr><w:tblGrid>{grid}</w:tblGrid>{rows}</w:tbl>')


STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + '>'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>'
          '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>' + TIGHT + '</w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:style>'
          '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/>'
          '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>'
          '<w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>')


def write(name, body, compat=15):
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
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + NS + '>'
                f'<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{compat}"/></w:compat></w:settings>')
    sect = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
            'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')
    doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{p("a01")}{body}{p("c01")}{sect}</w:body></w:document>'
    with zipfile.ZipFile(os.path.join(out, name + ".docx"), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", rels)
        z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/document.xml", doc)
        z.writestr("word/styles.xml", STYLES); z.writestr("word/settings.xml", settings)


NORMAL_EMPTY = lambda r: "<w:p/>"
TIGHT_EMPTY = lambda r: p()
M11 = p("m11")
write("empty", table(M11, NORMAL_EMPTY))
write("empty-center", table(M11, NORMAL_EMPTY, "center"))
write("empty-c14", table(M11, NORMAL_EMPTY), 14)
write("none-spacing", table(M11, TIGHT_EMPTY))
write("text", table(M11, lambda r: p(f"x{r}1")))
write("tall", table(M11, lambda r: "".join(p(f"x2{k}") for k in range(1, 6)) if r == 2 else p()))
TALL = "".join(p(f"m1{k}") for k in range(1, 7))
write("restart-tall", table(TALL, TIGHT_EMPTY))
write("restart-tall-empty", table(TALL, NORMAL_EMPTY))
write("two-merged", table(M11, NORMAL_EMPTY, third_merged=True, third_cont=NORMAL_EMPTY))
print(out)
