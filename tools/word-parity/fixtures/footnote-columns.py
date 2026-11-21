#!/usr/bin/env python3
"""Fixture family: where Word puts the footnotes of a section with columns.

    python3 tools/word-parity/fixtures/footnote-columns.py [out dir]   (default tests/word/fixtures/footnote-columns)

then, to ask Word (winvm) and read its PDFs:

    WORD_VM_DIR=C:/Users/olivia/RbtExcelParity/fncol \\
      node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/footnote-columns --refresh-word
    python3 tools/word-parity/pdf-geometry.py --paths tests/word/fixtures/footnote-columns/word/*.pdf \\
      > tests/word/fixtures/footnote-columns/word/geometry.tsv

tests/word/footnote-columns.fsx checks ours against those.

Hypotheses (docs/layout-parity-flow.md ledger), from the showcase magazine-feature (two columns, its notes
under the columns their references are in, one separator at the first column): which column a note goes
under - its reference's, or the notes flow through the columns like text -, where the separator is and how
long, how far each column's text runs down, a reference near a column's foot, balanced columns.

Every document: Letter, 1in margins, Calibri 11pt single-spaced lines (13.43pt) "Body k" (48 to a full
column), nothing before or after; FootnoteText Calibri 10pt single, nothing after; the separators as Word
writes them; two columns 3in wide, 0.5in apart (column 2 from 4.5in = 324pt), compatibility mode 15 unless
the name starts m14-. Notes "Note n line k" are made of line breaks; "wrap" notes are one long sentence.
Only what the name says varies (cK: the reference in column K; Body 1-48 are column 1 without notes):

  one-c1, one-c2               a 3-line note referenced in Body 5 / Body 60
  two-c1c2                     a 3-line note in Body 5, a 1-line one in Body 60
  two-c1c2-tall2               a 1-line note in Body 5, a 5-line one in Body 60 (does column 1 run down further?)
  two-c2c1-p2                  as two-c1c2 on page 2 (Body 150 and Body 105 ... referenced in that order: 105, 150)
  three-c1                     three 3-line notes, Body 5, 10, 15 (stacked under column 1, or spilling into 2?)
  long-c1                      a 12-line note in Body 5
  long-c1-30                   a 30-line note in Body 5 (more than half the column)
  wrap-c1, wrap-c2             a long sentence note in Body 5 / Body 60 (its width)
  both-wrap                    wrap notes in Body 5 and Body 60
  bot-c1-k<k>                  a 3-line note in Body k = 42 ... 47 (near column 1's foot)
  bot-c2-k<k>                  a 3-line note in Body k = 90 ... 95 (near column 2's foot)
  bal-c1, bal-c2, bal-both     a two-column section of 30 lines ending in a continuous break, then 5 lines in
                               one column: the note (3 lines) in Body 5 / Body 25 / both (3 and 1 line)
  bal-long                     the same with a 12-line note in Body 5
  three-col-c2, three-col-c3   three columns (2in, 0.25in apart), a 3-line note in Body 60 / Body 110
  uneq-c1, uneq-c2             unequal columns (2in and 4in, 0.5in apart), a wrap note in Body 5 / Body 60
  mixed-top                    10 lines in one column, then (continuous) a two-column section of 120 lines;
                               the note (wrap) referenced in the one-column part's Body 3
  mixed-col                    the same, the note referenced in the two-column part's Body 5
  four-col-c1                  four columns (1.4in, 0.25in apart), a wrap note in Body 5 (is the 2in separator cut?)
  mixed-both                   mixed-top with a second wrap note in the two-column part's Body 5
  huge-c1                      a 120-line note in Body 5 (on over the page)
  balp-<n>                     a two-column section of n = 120, 130, 200 lines ending in a continuous break (then 5
                               lines in one column): are its last page's columns balanced?
  balp-120-fn2, balp-120-fn1   balp-120 with a 3-line note referenced on page 2 (Body 100) / page 1 (Body 5)
  sep-text-c1c2                two-c1c2 with a word after each separator's mark (where its paragraph is)
  m14-one-c2, m14-three-c1, m14-two-c1c2-tall2, m14-wrap-c1, m14-bal-c2   compatibility mode 14
No dependencies: the package is written as XML.
"""
import os, sys, zipfile

out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), "../../../tests/word/fixtures/footnote-columns")
os.makedirs(out, exist_ok=True)
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
SINGLE = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
WRAP = ("This note is one long sentence that runs on over several lines of the note area, so that where its "
        "lines end shows how wide Word sets the notes of a section with columns.")

def cols_xml(cols):
    if cols == 1: return '<w:cols w:space="720"/>'
    if cols == "uneq":
        return ('<w:cols w:num="2" w:equalWidth="0"><w:col w:w="2880" w:space="720"/><w:col w:w="5760"/></w:cols>')
    if cols in (3, 4): return f'<w:cols w:num="{cols}" w:space="360"/>'
    return f'<w:cols w:num="{cols}" w:space="720"/>'

def sect(cols=2, kind=None):
    t = f'<w:type w:val="{kind}"/>' if kind else ''
    return (f'<w:sectPr>{t}<w:footnotePr><w:numFmt w:val="decimal"/></w:footnotePr><w:pgSz w:w="12240" w:h="15840"/>'
            f'<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>{cols_xml(cols)}</w:sectPr>')

STYLES = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {NS}><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
          f'<w:pPrDefault><w:pPr>{SINGLE}</w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
          '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>'
          '<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/>'
          f'<w:pPr>{SINGLE}</w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>'
          '<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:basedOn w:val="DefaultParagraphFont"/>'
          '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style></w:styles>')

def separators(text=False):
    s = '<w:r><w:t>Sepword</w:t></w:r>' if text else ''
    c = '<w:r><w:t>Contword</w:t></w:r>' if text else ''
    return (f'<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr>{SINGLE}</w:pPr><w:r><w:separator/></w:r>{s}</w:p></w:footnote>'
            f'<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr>{SINGLE}</w:pPr><w:r><w:continuationSeparator/></w:r>{c}</w:p></w:footnote>')

MARK = '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r>'

def footnote(note, count):
    """A note of `count` lines (line breaks), or the long sentence when count is "wrap"."""
    if count == "wrap":
        body = f'<w:r><w:t xml:space="preserve">Note {note}: {WRAP}</w:t></w:r>'
    else:
        parts = []
        for k in range(1, count + 1):
            if k > 1: parts.append('<w:r><w:br/></w:r>')
            parts.append(f'<w:r><w:t xml:space="preserve">Note {note} line {k}</w:t></w:r>')
        body = "".join(parts)
    return f'<w:footnote w:id="{note}"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>{MARK}{body}</w:p></w:footnote>'

def ref(note):
    return f'<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="{note}"/></w:r>'

def para(text, extra="", ppr=""):
    return f'<w:p><w:pPr>{SINGLE}{ppr}</w:pPr><w:r><w:t>{text}</w:t></w:r>{extra}</w:p>'

def document(sections):
    """sections: [(label, lines, cols, refs {line: note}, kind of the section's own w:type)]; a section's
    last paragraph holds its w:sectPr, the last section's is the body's."""
    ps = []
    for s, (label, lines, cols, refs, kind) in enumerate(sections):
        last = s == len(sections) - 1
        for k in range(1, lines + 1):
            extra = ref(refs[k]) if k in refs else ""
            ppr = sect(cols, kind) if (not last and k == lines) else ""
            ps.append(para(f"{label} {k}", extra, ppr))
    label, lines, cols, refs, kind = sections[-1]
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{"".join(ps)}{sect(cols, kind)}</w:body></w:document>')

def simple(lines=150, cols=2, refs=None):
    return document([("Body", lines, cols, refs or {}, None)])

def write(name, doc, notes, mode=15, seps=None):
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
          '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
             '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
             '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>')
    settings = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {NS}><w:compat>'
                f'<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/></w:compat>'
                '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr></w:settings>')
    footnotes = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes {NS}>{seps or separators()}{"".join(notes)}</w:footnotes>'
    with zipfile.ZipFile(os.path.join(out, name + ".docx"), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", rels)
        z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/document.xml", doc)
        z.writestr("word/styles.xml", STYLES); z.writestr("word/settings.xml", settings); z.writestr("word/footnotes.xml", footnotes)

def family(prefix="", mode=15, only=None):
    def w(name, doc, notes, **kw):
        if only is None or name in only: write(prefix + name, doc, notes, mode=mode, **kw)
    w("one-c1", simple(refs={5: 1}), [footnote(1, 3)])
    w("one-c2", simple(refs={60: 1}), [footnote(1, 3)])
    w("two-c1c2", simple(refs={5: 1, 60: 2}), [footnote(1, 3), footnote(2, 1)])
    w("two-c1c2-tall2", simple(refs={5: 1, 60: 2}), [footnote(1, 1), footnote(2, 5)])
    w("two-c2c1-p2", simple(lines=200, refs={105: 1, 150: 2}), [footnote(1, 3), footnote(2, 1)])
    w("three-c1", simple(refs={5: 1, 10: 2, 15: 3}), [footnote(1, 3), footnote(2, 3), footnote(3, 3)])
    w("long-c1", simple(refs={5: 1}), [footnote(1, 12)])
    w("long-c1-30", simple(refs={5: 1}), [footnote(1, 30)])
    w("wrap-c1", simple(refs={5: 1}), [footnote(1, "wrap")])
    w("wrap-c2", simple(refs={60: 1}), [footnote(1, "wrap")])
    w("both-wrap", simple(refs={5: 1, 60: 2}), [footnote(1, "wrap"), footnote(2, "wrap")])
    for k in range(42, 48):
        w(f"bot-c1-k{k}", simple(refs={k: 1}), [footnote(1, 3)])
    for k in range(90, 96):
        w(f"bot-c2-k{k}", simple(refs={k: 1}), [footnote(1, 3)])
    def bal(refs):
        return document([("Body", 30, 2, refs, None), ("After", 5, 1, {}, "continuous")])
    w("bal-c1", bal({5: 1}), [footnote(1, 3)])
    w("bal-c2", bal({25: 1}), [footnote(1, 3)])
    w("bal-both", bal({5: 1, 25: 2}), [footnote(1, 3), footnote(2, 1)])
    w("bal-long", bal({5: 1}), [footnote(1, 12)])
    w("three-col-c2", simple(lines=250, cols=3, refs={60: 1}), [footnote(1, 3)])
    w("three-col-c3", simple(lines=250, cols=3, refs={110: 1}), [footnote(1, 3)])
    w("uneq-c1", simple(cols="uneq", refs={5: 1}), [footnote(1, "wrap")])
    w("uneq-c2", simple(cols="uneq", refs={60: 1}), [footnote(1, "wrap")])
    def mixed(refs_top, refs_cols):
        return document([("Top", 10, 1, refs_top, None), ("Body", 120, 2, refs_cols, "continuous")])
    w("mixed-top", mixed({3: 1}, {}), [footnote(1, "wrap")])
    w("mixed-col", mixed({}, {5: 1}), [footnote(1, "wrap")])
    w("four-col-c1", simple(lines=250, cols=4, refs={5: 1}), [footnote(1, "wrap")])
    w("mixed-both", mixed({3: 1}, {5: 2}), [footnote(1, "wrap"), footnote(2, "wrap")])
    w("huge-c1", simple(lines=300, refs={5: 1}), [footnote(1, 120)])
    # A section of columns over more than a page, ending in a continuous break: its last page's columns.
    def balp(lines, refs):
        return document([("Body", lines, 2, refs, None), ("After", 5, 1, {}, "continuous")])
    w("balp-120", balp(120, {}), [])
    w("balp-130", balp(130, {}), [])
    w("balp-200", balp(200, {}), [])
    w("balp-120-fn2", balp(120, {100: 1}), [footnote(1, 3)])
    w("balp-120-fn1", balp(120, {5: 1}), [footnote(1, 3)])
    w("sep-text-c1c2", simple(refs={5: 1, 60: 2}), [footnote(1, 3), footnote(2, 1)], seps=separators(text=True))

family()
family("m14-", 14, only={"one-c2", "three-c1", "two-c1c2-tall2", "wrap-c1", "bal-c2"})
print(out)
