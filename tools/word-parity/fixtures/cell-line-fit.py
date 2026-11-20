#!/usr/bin/env python3
"""Fixture family: a table cell's line at the page's foot, at single, 1.15 and 1.5 line spacing: does it stay when
its text at single spacing fits (as a body line does) or only when the whole line does - with the cell's bottom
margin and the border under the row?

    python3 tools/word-parity/fixtures/cell-line-fit.py tests/word/fixtures/cell-line-fit
    node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/cell-line-fit --refresh-word
    python3 tools/word-parity/pdf-geometry.py tests/word/fixtures/cell-line-fit/word/*.pdf > tests/word/fixtures/cell-line-fit/word/geometry.tsv

docx4j-getting-started's switch table (1.15 lines, a 43-twip bottom margin, 0.5pt borders) breaks its "\b field-argument"
row after its first line in Word, where the second line's text at single spacing fits above the page's foot by 1.9pt and
the whole line (with the margin and border) does not by 0.14. Each document: Letter, 1in margins, Calibri 11pt; a spacer
paragraph of an exact height, then a fixed one-cell table (6.5in) whose paragraph has 8 lines "P1 line" ... "P8 line"
(w:br-broken, no space after), then "C 1". The spacer puts line 5's foot - the whole line, the cell's bottom margin and
the border under it - o points below the page's text (negative: above). Name: <sp>-<setup>-o<o>[-m14]
  sp     s100, s115, s150: the cell paragraph's line spacing (240, 276, 360 auto)
  setup  bare: no cell margins at top and bottom, no borders; mb: 43-twip top and bottom margins, single 0.5pt borders
  -after the cell's paragraph has 5 lines and 10pt after (line 5 its last)
  -after2 as -after, a second paragraph of 3 lines "Q1 line" ... under it in the cell
  body-  the same 8 lines as a body paragraph (no table)
  -m14   compatibility mode 14 (else 15)
No dependencies: the package is written as XML.
"""
import os, sys, zipfile

out = sys.argv[1] if len(sys.argv) > 1 else "cell-line-fit"
os.makedirs(out, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
SECT = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/></w:sectPr>')
SINGLE = 11 * 1.2207
SPACING = {"s100": 240, "s115": 276, "s150": 360}


def para(text, ppr=''):
    return f'<w:p><w:pPr>{ppr}</w:pPr><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def document(sp, setup, o, after=False, body=False, second=False):
    line = SPACING[sp]
    pitch = SINGLE * line / 240
    n, k = (5, 5) if after else (8, 5)
    bordered = setup == "mb"
    bt = bb = 0.5 if bordered and not body else 0.0
    mt = mb = 43 / 20 if bordered and not body else 0.0
    spacer = 720 + o - (bt + mt + k * pitch + mb + bb) - 72
    spacer_tw = round(spacer * 20)
    body_xml = [f'<w:p><w:pPr><w:spacing w:line="{spacer_tw}" w:lineRule="exact"/></w:pPr><w:r><w:t>S</w:t></w:r></w:p>']
    runs = '<w:r><w:br/></w:r>'.join(f'<w:r><w:t xml:space="preserve">P{j} line</w:t></w:r>' for j in range(1, n + 1))
    ppr = f'<w:spacing w:after="{200 if after else 0}" w:line="{line}" w:lineRule="auto"/>'
    cell_p = f'<w:p><w:pPr>{ppr}</w:pPr>{runs}</w:p>'
    if second:
        q_runs = '<w:r><w:br/></w:r>'.join(f'<w:r><w:t xml:space="preserve">Q{j} line</w:t></w:r>' for j in range(1, 4))
        cell_p += f'<w:p><w:pPr><w:spacing w:after="0" w:line="{line}" w:lineRule="auto"/></w:pPr>{q_runs}</w:p>'
    if body:
        body_xml.append(cell_p)
    else:
        borders = ''.join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>' for s in ["top", "left", "bottom", "right", "insideH", "insideV"]) if bordered else ''
        mar = f'<w:tblCellMar><w:top w:w="{43 if bordered else 0}" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="{43 if bordered else 0}" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>'
        body_xml.append(f'<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
                        f'{"<w:tblBorders>" + borders + "</w:tblBorders>" if borders else ""}{mar}</w:tblPr>'
                        f'<w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>{cell_p}</w:tc></w:tr></w:tbl>')
    body_xml.append(para("C 1"))
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{"".join(body_xml)}{SECT}</w:body></w:document>'


def styles(widow=True):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr>'
            '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
            '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
            '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>'
            '</w:styles>')


def settings(mode):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}>'
            f'<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/></w:compat></w:settings>')


def write(name, doc, mode, widow=True):
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
          '</Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
             '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
             '</Relationships>')
    with zipfile.ZipFile(os.path.join(out, name + ".docx"), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", rels)
        z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/document.xml", doc)
        z.writestr("word/styles.xml", styles(widow)); z.writestr("word/settings.xml", settings(mode))


count = 0
OFFSETS = {"s100": [-0.8, 0.8], "s115": [-1.0, 0.8, 1.6, 2.8], "s150": [-1.0, 1.0, 3.0, 6.0, 7.5]}
oname = lambda o: ("m" if o < 0 else "") + f"{abs(o):g}".replace(".", "_")
for mode in [15, 14]:
    suffix = "-m14" if mode == 14 else ""
    for sp, offsets in OFFSETS.items():
        for setup in ["bare", "mb"]:
            for o in offsets:
                write(f"{sp}-{setup}-o{oname(o)}{suffix}", document(sp, setup, o), mode); count += 1
    for o in [-1.0, 0.8]:
        write(f"s115-mb-after-o{oname(o)}{suffix}", document("s115", "mb", o, after=True), mode); count += 1
        write(f"s115-mb-after2-o{oname(o)}{suffix}", document("s115", "mb", o, after=True, second=True), mode); count += 1
    write(f"body-s115-o0_8{suffix}", document("s115", "bare", 0.8, body=True), mode); count += 1
print(out, count)
