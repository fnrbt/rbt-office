#!/usr/bin/env python3
"""Fixture family: how close to the page's foot a table cell's line may end and stay - what of the cell's bottom
margin, the table's borders and the line's spacing has to fit above the bottom margin.

    python3 tools/word-parity/fixtures/cell-foot.py tests/word/fixtures/cell-foot
    node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/cell-foot --refresh-word
    python3 tools/word-parity/pdf-geometry.py tests/word/fixtures/cell-foot/word/*.pdf > tests/word/fixtures/cell-foot/word/geometry.tsv

docx4j-getting-started p36: "\\b field-argument" (two lines in a bordered cell, 0.5pt borders, 43 twips of bottom
margin, 1.15-spaced Calibri 11pt) - Word breaks the row after the cell's first line, ours kept both: the second
line's bottom and the cell's margin 0.36pt above the bottom margin. Each document: Letter, 1in margins (the
body's foot at 720pt), Calibri 11pt; 30 paragraphs "A k" of exactly 20pt lines (to 672pt), the last with a space
after that puts the cell's second line where wanted; then a one-column table of one row, its cell three
one-line paragraphs "R1" "R2" "R3" (no space after); "C 1" after the table:
  <b>-m<m>-<s>-g<g>[-m15]
  b: none, sz4 (0.5pt single borders all round and inside), sz12 (1.5pt)
  m: the cell's bottom margin in twips (tblCellMar), 0 or 43
  s: s100 single lines (13.43pt), s115 1.15 (15.44pt: 2pt of it above the text's single line)
  g: what is left between the bottom margin and R2's line's bottom less the cell's bottom margin, in
     hundredths of a point (n: negative) - its top at the top border's width below the row's top
in compatibility mode 14 (and, for sz4-m43, mode 15).
No dependencies: the package is written as XML.
"""
import os, sys, zipfile

out = sys.argv[1] if len(sys.argv) > 1 else "cell-foot"
os.makedirs(out, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
SECT = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/></w:sectPr>')
LINE = {"s100": (240, 2500 / 2048 * 11), "s115": (276, 2500 / 2048 * 11 * 1.15)}
BORDER = {"none": 0, "sz4": 4, "sz12": 12}

def para(text, ppr=''):
    return f'<w:p><w:pPr>{ppr}</w:pPr><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'

def document(border, margin, spacing, gap):
    line_twips, h = LINE[spacing]
    sz = BORDER[border]
    top = sz / 8
    # R2's bottom = 672 + after + top + 2h = 720 - gap - margin
    after = 48 - gap - margin / 20 - top - 2 * h
    exact = '<w:spacing w:after="0" w:line="400" w:lineRule="exact"/>'
    body = [para(f"A {k}", exact) for k in range(1, 30)]
    body.append(para("A 30", f'<w:spacing w:after="{round(after * 20)}" w:line="400" w:lineRule="exact"/>'))
    borders = ''
    if sz:
        edge = f'w:val="single" w:sz="{sz}" w:space="0" w:color="000000"'
        borders = '<w:tblBorders>' + ''.join(f'<w:{e} {edge}/>' for e in ("top", "left", "bottom", "right", "insideH", "insideV")) + '</w:tblBorders>'
    cellmar = f'<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="{margin}" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>'
    cell = "".join(para(f"R{k}", f'<w:spacing w:after="0" w:line="{line_twips}" w:lineRule="auto"/>') for k in range(1, 4))
    body.append(f'<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblLayout w:type="fixed"/>{borders}{cellmar}</w:tblPr>'
                f'<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>{cell}</w:tc></w:tr></w:tbl>')
    body.append(para("C 1"))
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{"".join(body)}{SECT}</w:body></w:document>'

STYLES = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
          '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>'
          '</w:styles>')

def settings(mode):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}>'
            f'<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/></w:compat></w:settings>')

def write(name, doc, mode):
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
        z.writestr("word/styles.xml", STYLES); z.writestr("word/settings.xml", settings(mode))

GAPS = [2.0, 1.25, 0.75, 0.25, -0.25, -1.0, -1.75]
count = 0
for border in BORDER:
    for margin in [0, 43]:
        for spacing in LINE:
            for gap in GAPS:
                g = f"g{'n' if gap < 0 else ''}{round(abs(gap) * 100)}"
                name = f"{border}-m{margin}-{spacing}-{g}"
                write(name, document(border, margin, spacing, gap), 14); count += 1
                if border == "sz4" and margin == 43:
                    write(name + "-m15", document(border, margin, spacing, gap), 15); count += 1
print(out, count)
