#!/usr/bin/env python3
"""Fixture family: which section's top and bottom margins a page takes when a continuous section break
starts a section part-way down it.

    python3 tools/word-parity/fixtures/section-margins.py tests/word/fixtures/section-margins
    node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/section-margins --refresh-word

Hypothesis (docs/layout-parity-flow.md ledger): a page's top and bottom margins are those of the section the
page starts in; a section that starts part-way down a page (w:type continuous) takes its own from the next
page on (its left and right margins and columns at once). Each document: Calibri 11pt single-spaced lines
(13.43pt), nothing before or after, a Letter page (792pt high), 1in left and right; the sections' paragraphs
"A k", "B k", "C k" one line each, a section's last paragraph holding its w:sectPr. Names give each section's
lines, top / bottom margins in twips and columns:
  b<B1>-<B2>[-c2|-c14]  section A 10 lines, bottom B1; section B (continuous) 90 lines, bottom B2, one column
                    or two (-c2); -c14: compatibility mode 14 (the others 15)
  t<T2>             section A 10 lines, top / bottom 1440; section B 90 lines, top T2: page 2's top
  late-b<B1>-<B2>   section A 60 lines (on to page 2), bottom B1; section B 60 lines, bottom B2: page 2's foot
  three-<B1>-<B2>-<B3>  A 10 lines, B (continuous) 10 lines, C (continuous) 90 lines, bottoms B1, B2, B3
  next-b<B1>-<B2>   as b<B1>-<B2> with B a w:type nextPage section (the control: its own margins at once)
With 72pt top, a bottom of 720 twips leaves 684pt (50 lines), 1440 648pt (48 lines), 2160 612pt (45 lines).
No dependencies: the package is written as XML.
"""
import os, sys, zipfile

out = sys.argv[1] if len(sys.argv) > 1 else "section-margins"
os.makedirs(out, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

def sect_pr(top, bottom, cols=1, kind=None):
    t = f'<w:type w:val="{kind}"/>' if kind else ''
    c = f'<w:cols w:num="{cols}" w:space="720"/>' if cols > 1 else '<w:cols w:space="720"/>'
    return (f'<w:sectPr>{t}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="{top}" w:right="1440" w:bottom="{bottom}" '
            f'w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>{c}</w:sectPr>')

def document(sections):
    """sections: (letter, lines, top, bottom, cols, kind) - kind the w:type of the section's own break."""
    body = []
    for s, (letter, lines, top, bottom, cols, kind) in enumerate(sections):
        last_section = s == len(sections) - 1
        for k in range(1, lines + 1):
            ppr = '' if last_section or k < lines else '<w:pPr>' + sect_pr(top, bottom, cols, kind) + '</w:pPr>'
            body.append(f'<w:p>{ppr}<w:r><w:t>{letter} {k}</w:t></w:r></w:p>')
    letter, lines, top, bottom, cols, kind = sections[-1]
    body.append(sect_pr(top, bottom, cols, kind))
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{"".join(body)}</w:body></w:document>'

STYLES = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
          '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>'
          '</w:styles>')

def settings(mode):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}>'
            f'<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/></w:compat></w:settings>')

def write(name, doc, mode=15):
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

C = "continuous"
count = 0
for b1, b2 in [(720, 1440), (1440, 720), (720, 2160), (2160, 720)]:
    # (A section's w:type is how it starts: B's, in the body's w:sectPr.)
    write(f"b{b1}-{b2}", document([("A", 10, 1440, b1, 1, None), ("B", 90, 1440, b2, 1, C)])); count += 1
    write(f"b{b1}-{b2}-c2", document([("A", 10, 1440, b1, 1, None), ("B", 90, 1440, b2, 2, C)])); count += 1
    write(f"b{b1}-{b2}-c14", document([("A", 10, 1440, b1, 1, None), ("B", 90, 1440, b2, 1, C)]), mode=14); count += 1
    write(f"late-b{b1}-{b2}", document([("A", 60, 1440, b1, 1, None), ("B", 60, 1440, b2, 1, C)])); count += 1
    write(f"next-b{b1}-{b2}", document([("A", 10, 1440, b1, 1, None), ("B", 90, 1440, b2, 1, "nextPage")])); count += 1
for t2 in [720, 2160]:
    write(f"t{t2}", document([("A", 10, 1440, 1440, 1, None), ("B", 90, t2, 1440, 1, C)])); count += 1
for b in [(720, 1440, 2160), (2160, 1440, 720), (1440, 720, 2160)]:
    write(f"three-{b[0]}-{b[1]}-{b[2]}", document([("A", 10, 1440, b[0], 1, None), ("B", 10, 1440, b[1], 1, C), ("C", 90, 1440, b[2], 1, C)])); count += 1
print(out, count)
