#!/usr/bin/env python3
"""Fixture family: does a run that holds only a tab or a space make its line as tall as its font?

    python3 tools/word-parity/fixtures/blank-run-height.py tests/word/fixtures/blank-run-height
    node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/blank-run-height --refresh-word
    python3 tools/word-parity/fixtures/blank-run-height.py --table tests/word/fixtures/blank-run-height/word/*.pdf \
        > tests/word/fixtures/blank-run-height/word/heights.tsv

docx4j-getting-started's code samples: a paragraph of Consolas 10pt text after a run of one tab in Consolas 12pt
(p38's first line). Ours made the line 14.05pt (Consolas 12pt), Word's next baseline is 11.7 below (Consolas 10pt)
- every line of the page 2.3pt low, its last one over the margin. Each document: Letter, 1in margins, nothing
before or after any paragraph, single spacing; "Topq", ten probe paragraphs, "Botq" (Calibri 11pt, 13.43pt
lines). A probe's line: (Botq's top - Topq's top - 13.43) / 10. The probe: "Probe" in 10pt (12.21pt lines in
Calibri) with a 20pt run (24.41 in Calibri):
  ctl            no 20pt run                       lead-x       "x" before the text (the control: counts)
  lead-tab       a tab before the text             lead-space   a space before the text
  lead-2tab      two tabs                          lead-tabx    a tab and "x" in one run
  mid-tab        between two runs of text          mid-space    a space between them
  end-tab        after the text                    end-space    a space after it
  only-tab       the tab alone (the mark 10pt)     only-space   the space alone
  lead-tab-cons  Consolas: a 12pt tab, 10pt text (as docx4j-getting-started)
Suffixes: -c14 (Word 2010's layout), -cell (the probes in a one-cell table), -wrap (the probe's text two lines,
the 20pt run on the first: does the second line take it?).
`--table <pdf> ...` writes <doc> TAB <probe line height> from Word's PDFs (pdftotext's word boxes).
No dependencies: the package is written as XML.
"""
import os, re, subprocess, sys, zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
CAL = 'w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"'
CON = 'w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"'


def run(content, size, font=CAL):
    return f'<w:r><w:rPr><w:rFonts {font}/><w:sz w:val="{size}"/><w:szCs w:val="{size}"/></w:rPr>{content}</w:r>'


T = lambda s: f'<w:t xml:space="preserve">{s}</w:t>'
TAB = '<w:tab/>'
LONG = "Probe " + "words to wrap the line " * 5
PROBES = {
    "ctl": lambda: run(T("Probe"), 20),
    "lead-x": lambda: run(T("x"), 40) + run(T("Probe"), 20),
    "lead-tab": lambda: run(TAB, 40) + run(T("Probe"), 20),
    "lead-space": lambda: run(T(" "), 40) + run(T("Probe"), 20),
    "lead-2tab": lambda: run(TAB + TAB, 40) + run(T("Probe"), 20),
    "lead-tabx": lambda: run(TAB + T("x"), 40) + run(T("Probe"), 20),
    "mid-tab": lambda: run(T("Pro"), 20) + run(TAB, 40) + run(T("be"), 20),
    "mid-space": lambda: run(T("Pro"), 20) + run(T(" "), 40) + run(T("be"), 20),
    "end-tab": lambda: run(T("Probe"), 20) + run(TAB, 40),
    "end-space": lambda: run(T("Probe"), 20) + run(T(" "), 40),
    "only-tab": lambda: run(TAB, 40),
    "only-space": lambda: run(T(" "), 40),
    "lead-tab-cons": lambda: run(TAB, 24, CON) + run(T("Probe"), 20, CON),
}


def write(out, name, probe, mode=15, cell=False, wrap=False):
    text = lambda t: f'<w:p><w:r><w:t>{t}</w:t></w:r></w:p>'
    body = PROBES[probe]()
    if wrap:
        body = body.replace(T("Probe"), T(LONG)).replace(T("be"), T("be " + LONG))
    para = f'<w:p><w:pPr><w:rPr><w:rFonts {CAL}/><w:sz w:val="20"/></w:rPr></w:pPr>{body}</w:p>'
    probes = para * 10
    if cell:
        probes = ('<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>'
                  f'<w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>{probes}</w:tc></w:tr></w:tbl>')
    doc = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{text("Topq")}{probes}{text("Botq")}'
           '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
    styles = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts {CAL} w:eastAsia="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
              '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
              '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>'
              '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/>'
              '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
              '</w:styles>')
    settings = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}>'
                f'<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/></w:compat></w:settings>')
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             f'<Relationship Id="rId1" Type="{R}styles" Target="styles.xml"/><Relationship Id="rId2" Type="{R}settings" Target="settings.xml"/></Relationships>')
    with zipfile.ZipFile(os.path.join(out, name + ".docx"), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", rels); z.writestr("word/_rels/document.xml.rels", drels)
        z.writestr("word/document.xml", doc); z.writestr("word/styles.xml", styles); z.writestr("word/settings.xml", settings)


def variants():
    for probe in PROBES:
        yield probe, probe, 15, False, False
    for probe in ("lead-tab", "lead-space", "end-space", "only-tab"):
        yield f"{probe}-c14", probe, 14, False, False
    for probe in ("lead-tab", "lead-x"):
        yield f"{probe}-cell", probe, 15, True, False
        yield f"{probe}-wrap", probe, 15, False, True


def table(pdfs):
    for pdf in sorted(pdfs):
        tops = {}
        rows = subprocess.run(['pdftotext', '-bbox', pdf, '-'], capture_output=True, text=True, check=True).stdout
        for m in re.finditer(r'yMin="([\d.]+)" xMax="[\d.]+" yMax="[\d.]+">(Topq|Botq)</word>', rows):
            tops.setdefault(m.group(2), float(m.group(1)))
        name = os.path.basename(pdf)[:-4]
        print(f"{name}\t{(tops['Botq'] - tops['Topq'] - 13.43) / 10:.3f}" if len(tops) == 2 else f"{name}\t?")


if __name__ == "__main__":
    if sys.argv[1:2] == ["--table"]:
        table(sys.argv[2:])
    else:
        out = sys.argv[1] if len(sys.argv) > 1 else "blank-run-height"
        os.makedirs(out, exist_ok=True)
        for name, *rest in variants():
            write(out, name, *rest)
