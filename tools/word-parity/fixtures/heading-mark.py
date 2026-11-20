#!/usr/bin/env python3
"""Fixture family: a Heading 1 that starts with a w:lastRenderedPageBreak (where Word last broke the page),
at the top of a page it came to by the flow of the text, or in the middle of one.

    python3 tools/word-parity/fixtures/heading-mark.py tests/word/fixtures/heading-mark
    node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/heading-mark --refresh-word
    python3 tools/word-parity/pdf-geometry.py tests/word/fixtures/heading-mark/word/*.pdf > tests/word/fixtures/heading-mark/word/geometry.tsv

docx4j-repeating-section's "RepeatingSection positives" (Heading 1, 24pt before, keepNext and keepLines, after a
10pt-after paragraph; its text starts with the mark) goes to page 7 with the paragraphs it keeps with: Word sets
it at the top margin, ours 14pt lower (the reader makes the mark of a Heading 1 a page break before it, and in
Word 2010's layout a page break before keeps the space before, less the space after above it). Each document:
Letter, 1in margins, Calibri 11pt single-spaced, Normal 10pt after; filler paragraphs "A k" (no space after),
"Z" (10pt after), then "H-probe" (Heading 1: Calibri 14pt bold, 24pt before, 0 after, keepNext, keepLines),
"K" (keepNext), a four-line paragraph "L ...", "C 1":
  keep-<m>    the heading fits at the page's foot but the paragraphs kept with it do not (42 fillers)
  line-<m>    the heading itself does not fit (44 fillers)
  stale-<m>   the heading in the middle of the first page (10 fillers)
  m: mark (the heading's text starts with w:lastRenderedPageBreak), nomark
  stale-plain-mark  as stale-mark, the probe a Normal paragraph
in compatibility mode 15 and 14 (-m14).
No dependencies: the package is written as XML.
"""
import os, sys, zipfile

out = sys.argv[1] if len(sys.argv) > 1 else "heading-mark"
os.makedirs(out, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
SECT = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/></w:sectPr>')

def para(text, ppr=''):
    return f'<w:p><w:pPr>{ppr}</w:pPr><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'

LONG = " ".join(f"L{k} word" for k in range(1, 60))

def document(fillers, mark, plain=False):
    body = [para(f"A {k}", '<w:spacing w:after="0"/>') for k in range(1, fillers + 1)]
    body.append(para("Z"))
    rendered = '<w:lastRenderedPageBreak/>' if mark else ''
    style = '' if plain else '<w:pStyle w:val="Heading1"/>'
    body.append(f'<w:p><w:pPr>{style}</w:pPr><w:r>{rendered}<w:t>H-probe</w:t></w:r></w:p>')
    body.append(para("K", '<w:keepNext/>'))
    body.append(para(LONG))
    body.append(para("C 1"))
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{"".join(body)}{SECT}</w:body></w:document>'

STYLES = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="200" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
          '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>'
          '<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="480" w:after="0"/><w:outlineLvl w:val="0"/></w:pPr>'
          '<w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>'
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
for mode in [15, 14]:
    suffix = "-m14" if mode == 14 else ""
    for name, fillers in [("keep", 42), ("line", 44), ("stale", 10)]:
        for mark in [True, False]:
            write(f"{name}-{'mark' if mark else 'nomark'}{suffix}", document(fillers, mark), mode); count += 1
    write(f"stale-plain-mark{suffix}", document(10, True, True), mode); count += 1
print(out, count)
