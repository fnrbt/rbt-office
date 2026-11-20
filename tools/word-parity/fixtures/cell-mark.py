#!/usr/bin/env python3
"""Fixture family: a table row at a page's foot whose cells' paragraphs hold a w:lastRenderedPageBreak (where
Word last broke the page) - at a paragraph's start, or at the start of one of its lines.

    python3 tools/word-parity/fixtures/cell-mark.py tests/word/fixtures/cell-mark
    node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/cell-mark --refresh-word
    python3 tools/word-parity/pdf-geometry.py tests/word/fixtures/cell-mark/word/*.pdf > tests/word/fixtures/cell-mark/word/geometry.tsv

docx4j-headers-footers p30 ("top (Top Margin " / "Spacing)") and docx4j-getting-started p36 ("\\b field-" /
"argument"): Word breaks a two-cell row at the page's foot inside both cells, ours moved the row whole - both
rows' second cell starts with a stale mark. Each document: Letter, 1in margins, Courier New 10pt single-spaced
(11.33pt lines, 57 to the page), no space after; F filler lines "A k", then a borderless two-column table
(4680 twips each) of one row: its cells' paragraphs four lines of 30 characters ("<cell>L<k> aaaa bbbb cccc
dddd eeee"), a w:br ending each but the last (or, `wrap-*`, one run of 35 words that wraps), then "C 1" after the table:
  base-f<F>           no marks (F 55: two of the cell's lines fit, 56: one)
  c2first-f<F>        cell 2's paragraph starts with the mark
  c1first-f<F>        cell 1's
  bothfirst-f<F>      both
  c1line2-f<F>        cell 1's line 2 starts with it (a line that fits at F 55)
  c1line3-f<F>        cell 1's line 3 (the first that does not fit at F 55)
  hf-f<F>             as docx4j-headers-footers: cell 1's line 2, cell 2's lines 1 and 2
  mid-c2first         the row in the middle of page 1 (F 10)
  wrap-c2first-f<F>   wrapped text, cell 2's run starts with the mark
  base-f54, c2first-f54, c1first-f54, c1line2-f54   three of the four lines fit (widow control's case in -m15)
  kl-base-f55, kl-c1first-f55   cell 1's paragraph keepLines, without the mark and with it at its start
in compatibility mode 14 and 15 (-m15: widow control on, from no w:widowControl).
No dependencies: the package is written as XML.
"""
import os, sys, zipfile

out = sys.argv[1] if len(sys.argv) > 1 else "cell-mark"
os.makedirs(out, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
SECT = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/></w:sectPr>')
MARK = '<w:lastRenderedPageBreak/>'

def para(text):
    return f'<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'

def cell_para(cell, marks, wrap, keep=False):
    """marks: the line numbers (1-based) whose run starts with the mark."""
    if wrap:
        text = " ".join(f"{cell}w{k}" for k in range(1, 36))
        return f'<w:p><w:r>{MARK if 1 in marks else ""}<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
    ppr = '<w:pPr><w:keepLines/></w:pPr>' if keep else ''
    runs = []
    for k in range(1, 5):
        br = '<w:br/>' if k < 4 else ''
        runs.append(f'<w:r>{MARK if k in marks else ""}<w:t xml:space="preserve">{cell}L{k} aaaa bbbb cccc dddd eeee</w:t>{br}</w:r>')
    return f'<w:p>{ppr}{"".join(runs)}</w:p>'

def document(fillers, marks1, marks2, wrap=False, keep=False):
    body = [para(f"A {k}") for k in range(1, fillers + 1)]
    cells = "".join(f'<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>{cell_para(c, m, wrap, keep and c == "P")}</w:tc>'
                    for c, m in (("P", marks1), ("Q", marks2)))
    body.append('<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>'
                f'<w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid><w:tr>{cells}</w:tr></w:tbl>')
    body.append(para("C 1"))
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{"".join(body)}{SECT}</w:body></w:document>'

STYLES = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:eastAsia="Courier New" w:cs="Courier New"/><w:sz w:val="20"/></w:rPr></w:rPrDefault>'
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

count = 0
for mode in [14, 15]:
    suffix = "-m15" if mode == 15 else ""
    for f in [55, 56]:
        for name, m1, m2 in [("base", [], []), ("c2first", [], [1]), ("c1first", [1], []), ("bothfirst", [1], [1]),
                             ("c1line2", [2], []), ("c1line3", [3], []), ("hf", [2], [1, 2])]:
            write(f"{name}-f{f}{suffix}", document(f, m1, m2), mode); count += 1
        write(f"wrap-c2first-f{f}{suffix}", document(f, [], [1], True), mode); count += 1
    write(f"mid-c2first{suffix}", document(10, [], [1]), mode); count += 1
    for name, m1, m2 in [("base", [], []), ("c2first", [], [1]), ("c1first", [1], []), ("c1line2", [2], [])]:
        write(f"{name}-f54{suffix}", document(54, m1, m2), mode); count += 1
    for name, m1 in [("kl-base", []), ("kl-c1first", [1])]:
        write(f"{name}-f55{suffix}", document(55, m1, [], keep=True), mode); count += 1
print(out, count)
