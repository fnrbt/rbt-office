#!/usr/bin/env python3
"""Fixture family: where Word puts the footnote block and how far down the body may run above it.

    python3 tools/word-parity/fixtures/footnote-block.py [out dir]   (default tests/word/fixtures/footnote-block)

then, to ask Word (winvm) and compare with ours (Word's pages and PDFs in <out dir>/word):

    node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/footnote-block --refresh-word
    python3 tools/word-parity/pdf-baselines.py tests/word/fixtures/footnote-block/word/*.pdf > tests/word/fixtures/footnote-block/word/pdf-lines.tsv
    cscript //nologo para-positions.vbs <the .docx on the VM> <out>   -> tests/word/fixtures/footnote-block/word-positions/*.tsv (UTF-8)

tests/word/footnote-block.fsx checks ours against those.

Hypothesis (docs/layout-parity-flow.md ledger): the footnote block is stacked up from the bottom margin
(its last line's bottom is the page height less the bottom margin); the body may run down to the
block's top less the separator; a footnote line is as tall as its tallest content, the superscript
reference mark included.

Every document: Letter, 1in margins (the body ends at 720pt), Calibri; body paragraphs "Body <k>" at
single spacing (Calibri 11pt, 13.43pt) with nothing before or after; the footnotes in FootnoteText
(10pt, single, nothing after, as Word writes it) with their lines made by line breaks; the separator
and continuation separator as Word writes them (a paragraph single-spaced, nothing after). Only what
the name says varies:

  one-1, one-2, one-3, one-10   one footnote of 1, 2, 3, 10 lines (its reference in Body 1)
  two-1                         two one-line footnotes (Body 1, Body 2)
  two-para                      one footnote of two one-line paragraphs
  super-10                      a 10-line footnote with a superscript letter on every line
  nosuper-1                     the reference mark not raised (FootnoteReference without superscript)
  custom-1                      a custom mark "*" (customMarkFollows), plain text in the note
  dflt115-1                     w:pPrDefault of 1.15 and 8pt after (FootnoteText keeps its own single, 0 after)
  dflt115-inherit-2             the same, FootnoteText without spacing of its own: it inherits 1.15 / 8pt after
  dflt115-inherit-two           the same with two one-line footnotes
  fnspaced-2                    FootnoteText 6pt before, 12pt after, 1.5 lines
  sep-double-1                  the separator paragraph at double spacing
  sep-big-1                     the separator paragraph's mark and run in 20pt
  cont-40                       a 40-line footnote referenced in Body 35: it runs over to page 2
  cont-k<k>                     a 10-line footnote referenced in Body k = 44 ... 47: room on page 1 for 3, 2, 1, 0
                                of its lines under the separator - how much of it has to stay with its reference?
  cont-k45-nowidow, cont-k46-nowidow   the same with FootnoteText's widow control off
  cont11-k38[-nowidow]         an 11-line note with room for 10: does widow control keep 9 + 2 (10 + 1 without)?
  two-split-k<k>                a 3-line note (Body 30) and a 10-line one in Body k = 38, 43, 44 (room for 7, 1, 0 of it)
  cont-new                      cont-40 and a 3-line note in Body 40 (on page 2 under the continued one?)
  cont-new-sep24, cont-sep24    the same and cont-40 with the continuation separator 24pt before (which separator is used)
  cont-new-sepdouble, cont-sepdouble   ... with the continuation separator double-spaced
  cont-septext                  cont-40 with a word after each separator's mark (which paragraph is shown where)
  para-ref2, para-ref5[-nowidow]   39 body lines, then a 6-line paragraph with a 10-line note's reference at the end
                                of its line 2 (room for 6 note lines under it) or 5 (room for 3: the line after it alone on page 2)
  cont-notice-empty-k44         cont-k44 with an empty continuation notice paragraph (does it take a line?)
  cont-sep1double               cont-40 with the separator double-spaced, the continuation separator single
  cont-100                      a 100-line note in Body 35: on over two page breaks
  cont-notice-40, -100, -k44..46   cont-40, cont-100 and cont-k44..46 with a continuation notice (id 1; the note id 2)
  fnbefore-1, fnbefore-two      FootnoteText 12pt before (single, 0 after): one, two one-line footnotes
  body-super                    no footnotes: 40 body lines with a 11pt superscript "x" each (does it raise the line?)
  limit-<d>                     a body that ends <d>pt from where the hypothesis puts its limit above a
                                one-line footnote (the first paragraph is an exact spacer holding the reference,
                                so the lines after it are where they are; Body 45's bottom): d = -1.5 ... +1.5
  limit-nofn-<d>                the same without a footnote (Body 47's bottom at 720 + d): the limit is the bottom margin
No dependencies: the package is written as XML.
"""
import os, sys, zipfile

out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), "../../../tests/word/fixtures/footnote-block")
os.makedirs(out, exist_ok=True)
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
SINGLE = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
SECT = ('<w:sectPr><w:footnotePr><w:numFmt w:val="decimal"/></w:footnotePr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')
CALIBRI11 = 13.427   # Calibri's single line per point (1.2207) x 11
LINE_STEP = CALIBRI11

def styles(ppr_default="<w:pPr>" + SINGLE + "</w:pPr>", fn_ppr=SINGLE, ref_rpr='<w:vertAlign w:val="superscript"/>'):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {NS}><w:docDefaults><w:rPrDefault><w:rPr>'
            '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
            f'<w:pPrDefault>{ppr_default}</w:pPrDefault></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
            '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>'
            '<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/>'
            f'<w:pPr>{fn_ppr}</w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>'
            '<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:basedOn w:val="DefaultParagraphFont"/>'
            f'<w:rPr>{ref_rpr}</w:rPr></w:style></w:styles>')

def separators(sep_ppr=SINGLE, sep_rpr="", cont_ppr=None, notice=None):
    """`cont_ppr`: the continuation separator's own paragraph properties (default: the separator's);
    `notice`: the text of a continuation notice (w:type="continuationNotice", id 1, in FootnoteText)."""
    rpr = f'<w:rPr>{sep_rpr}</w:rPr>' if sep_rpr else ''
    cont_ppr = sep_ppr if cont_ppr is None else cont_ppr
    notice_run = f'<w:r><w:t>{notice}</w:t></w:r>' if notice else ''
    notice_xml = (f'<w:footnote w:type="continuationNotice" w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>'
                  f'{notice_run}</w:p></w:footnote>') if notice is not None else ''
    return (f'<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr>{sep_ppr}{rpr}</w:pPr><w:r>{rpr}<w:separator/></w:r></w:p></w:footnote>'
            f'<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr>{cont_ppr}{rpr}</w:pPr><w:r>{rpr}<w:continuationSeparator/></w:r></w:p></w:footnote>'
            + notice_xml)

def lines_text(note, count, super_each=False):
    parts = []
    for k in range(1, count + 1):
        if k > 1: parts.append('<w:r><w:br/></w:r>')
        parts.append(f'<w:r><w:t xml:space="preserve">Note {note} line {k}</w:t></w:r>')
        if super_each: parts.append('<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>x</w:t></w:r>')
    return "".join(parts)

def footnote(note, count, custom=False, super_each=False):
    mark = '<w:r><w:t>*</w:t></w:r>' if custom else '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
    return (f'<w:footnote w:id="{note}"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>{mark}'
            f'<w:r><w:t xml:space="preserve"> </w:t></w:r>{lines_text(note, count, super_each)}</w:p></w:footnote>')

def footnote_paras(note, paras):
    ps = []
    for k in range(1, paras + 1):
        mark = '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r>' if k == 1 else ''
        ps.append(f'<w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>{mark}<w:r><w:t>Note {note} paragraph {k}</w:t></w:r></w:p>')
    return f'<w:footnote w:id="{note}">{"".join(ps)}</w:footnote>'

def ref(note, custom=False):
    if custom: return f'<w:r><w:footnoteReference w:customMarkFollows="1" w:id="{note}"/><w:t>*</w:t></w:r>'
    return f'<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="{note}"/></w:r>'

def body_para(k, extra=""):
    return f'<w:p><w:pPr>{SINGLE}</w:pPr><w:r><w:t>Body {k}</w:t></w:r>{extra}</w:p>'

def body(count=60, refs=None, spacer=None):
    """`refs`: {body line: run xml}; `spacer`: (exact height in pt, run xml) as a first paragraph."""
    refs = refs or {}
    ps = []
    if spacer:
        height, run = spacer
        ps.append(f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="{round(height * 20)}" w:lineRule="exact"/></w:pPr>'
                  f'<w:r><w:t>Spacer</w:t></w:r>{run}</w:p>')
    for k in range(1, count + 1):
        ps.append(body_para(k, refs.get(k, "")))
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{"".join(ps)}{SECT}</w:body></w:document>')

def write(name, document, notes, styles_xml=None, seps=None, notice=False):
    styles_xml = styles_xml or styles()
    seps = seps if seps is not None else separators()
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
                '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
                '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/>' + ('<w:footnote w:id="1"/>' if notice else '') +
                '</w:footnotePr></w:settings>')
    footnotes = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes {NS}>{seps}{"".join(notes)}</w:footnotes>'
    with zipfile.ZipFile(os.path.join(out, name + ".docx"), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", rels)
        z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/document.xml", document)
        z.writestr("word/styles.xml", styles_xml); z.writestr("word/settings.xml", settings); z.writestr("word/footnotes.xml", footnotes)

for n in (1, 2, 3, 10):
    write(f"one-{n}", body(refs={1: ref(1)}), [footnote(1, n)])
write("two-1", body(refs={1: ref(1), 2: ref(2)}), [footnote(1, 1), footnote(2, 1)])
write("two-para", body(refs={1: ref(1)}), [footnote_paras(1, 2)])
write("super-10", body(refs={1: ref(1)}), [footnote(1, 10, super_each=True)])
write("nosuper-1", body(refs={1: ref(1)}), [footnote(1, 1)], styles(ref_rpr=""))
write("custom-1", body(refs={1: ref(1, custom=True)}), [footnote(1, 1, custom=True)])
DFLT115 = '<w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/></w:pPr>'
write("dflt115-1", body(refs={1: ref(1)}), [footnote(1, 1)], styles(DFLT115))
write("dflt115-inherit-2", body(refs={1: ref(1)}), [footnote(1, 2)], styles(DFLT115, fn_ppr=""))
write("dflt115-inherit-two", body(refs={1: ref(1), 2: ref(2)}), [footnote(1, 1), footnote(2, 1)], styles(DFLT115, fn_ppr=""))
write("fnspaced-2", body(refs={1: ref(1)}), [footnote(1, 2)],
      styles(fn_ppr='<w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="auto"/>'))
write("sep-double-1", body(refs={1: ref(1)}), [footnote(1, 1)], seps=separators('<w:spacing w:before="0" w:after="0" w:line="480" w:lineRule="auto"/>'))
write("sep-big-1", body(refs={1: ref(1)}), [footnote(1, 1)], seps=separators(sep_rpr='<w:sz w:val="40"/><w:szCs w:val="40"/>'))
write("cont-40", body(refs={35: ref(1)}), [footnote(1, 40)])
for k in (44, 45, 46, 47):
    write(f"cont-k{k}", body(refs={k: ref(1)}), [footnote(1, 10)])
for k in (45, 46):
    write(f"cont-k{k}-nowidow", body(refs={k: ref(1)}), [footnote(1, 10)], styles(fn_ppr='<w:widowControl w:val="0"/>' + SINGLE))
# A note broken across pages (the rule the cont-* above confirmed), at its edges:
write("cont11-k38", body(refs={38: ref(1)}), [footnote(1, 11)])
write("cont11-k38-nowidow", body(refs={38: ref(1)}), [footnote(1, 11)], styles(fn_ppr='<w:widowControl w:val="0"/>' + SINGLE))
for k in (38, 43, 44):
    write(f"two-split-k{k}", body(refs={30: ref(1), k: ref(2)}), [footnote(1, 3), footnote(2, 10)])
write("cont-new", body(refs={35: ref(1), 40: ref(2)}), [footnote(1, 40), footnote(2, 3)])
BEFORE24 = '<w:spacing w:before="480" w:after="0" w:line="240" w:lineRule="auto"/>'
DOUBLE = '<w:spacing w:before="0" w:after="0" w:line="480" w:lineRule="auto"/>'
write("cont-new-sep24", body(refs={35: ref(1), 40: ref(2)}), [footnote(1, 40), footnote(2, 3)], seps=separators(cont_ppr=BEFORE24))
write("cont-sep24", body(refs={35: ref(1)}), [footnote(1, 40)], seps=separators(cont_ppr=BEFORE24))
write("cont-new-sepdouble", body(refs={35: ref(1), 40: ref(2)}), [footnote(1, 40), footnote(2, 3)], seps=separators(cont_ppr=DOUBLE))
write("cont-sepdouble", body(refs={35: ref(1)}), [footnote(1, 40)], seps=separators(cont_ppr=DOUBLE))
TEXTSEPS = separators().replace('<w:separator/></w:r>', '<w:separator/></w:r><w:r><w:t>Sepword</w:t></w:r>').replace(
    '<w:continuationSeparator/></w:r>', '<w:continuationSeparator/></w:r><w:r><w:t>Contword</w:t></w:r>')
write("cont-septext", body(refs={35: ref(1)}), [footnote(1, 40)], seps=TEXTSEPS)
def body_with_para(before, lines, ref_line, note, widow=True, after=20):
    """`before` single-line paragraphs, then one paragraph of `lines` lines (line breaks) with the
    reference to `note` at the end of line `ref_line`, then `after` more."""
    runs = []
    for k in range(1, lines + 1):
        if k > 1: runs.append('<w:r><w:br/></w:r>')
        runs.append(f'<w:r><w:t xml:space="preserve">Para line {k}</w:t></w:r>')
        if k == ref_line: runs.append(ref(note))
    ppr = SINGLE + ('' if widow else '<w:widowControl w:val="0"/>')
    ps = [body_para(k) for k in range(1, before + 1)]
    ps.append(f'<w:p><w:pPr>{ppr}</w:pPr>{"".join(runs)}</w:p>')
    ps += [body_para(k) for k in range(before + 1, before + after + 1)]
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{"".join(ps)}{SECT}</w:body></w:document>')
write("para-ref2", body_with_para(39, 6, 2, 1), [footnote(1, 10)])
write("para-ref5", body_with_para(39, 6, 5, 1), [footnote(1, 10)])
write("para-ref5-nowidow", body_with_para(39, 6, 5, 1, widow=False), [footnote(1, 10)])
write("cont-notice-empty-k44", body(refs={44: ref(2)}), [footnote(2, 10)], seps=separators(notice=""), notice=True)
write("cont-sep1double", body(refs={35: ref(1)}), [footnote(1, 40)], seps=separators(sep_ppr=DOUBLE, cont_ppr=SINGLE))
write("cont-100", body(count=120, refs={35: ref(1)}), [footnote(1, 100)])
write("cont-notice-40", body(refs={35: ref(2)}), [footnote(2, 40)], seps=separators(notice="Continued on the next page"), notice=True)
for k in (44, 45, 46):
    write(f"cont-notice-k{k}", body(refs={k: ref(2)}), [footnote(2, 10)], seps=separators(notice="Continued on the next page"), notice=True)
write("cont-notice-100", body(count=120, refs={35: ref(2)}), [footnote(2, 100)], seps=separators(notice="Continued on the next page"), notice=True)
BEFORE12 = '<w:spacing w:before="240" w:after="0" w:line="240" w:lineRule="auto"/>'
write("fnbefore-1", body(refs={1: ref(1)}), [footnote(1, 1)], styles(fn_ppr=BEFORE12))
write("fnbefore-two", body(refs={1: ref(1), 2: ref(2)}), [footnote(1, 1), footnote(2, 1)], styles(fn_ppr=BEFORE12))
write("body-super", body(count=40, refs={k: '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>x</w:t></w:r>' for k in range(1, 41)}), [])

# The limit: under the hypothesis, with a one-line footnote (10pt Calibri, 12.21pt, its mark raised) and
# the separator a single-spaced Calibri 11pt paragraph (13.43pt), the body may run down to
# 720 - 12.21 - 13.43 = 694.36pt. The spacer (exact, its top at 72pt) puts the bottom of Body 46 at
# limit + d: 72 + spacer + 45 x 13.427 = 694.36 + d.
PREDICTED_LIMIT = 720 - 10 * 1.2207 - CALIBRI11
LINES = 45
for d in (-1.5, -1.0, -0.5, 0.0, 0.5, 1.0, 1.5):
    spacer = PREDICTED_LIMIT + d - 72 - LINES * LINE_STEP
    tag = ("m" if d < 0 else "p") + f"{abs(d):.1f}".replace(".", "")
    write(f"limit-{tag}", body(count=LINES + 4, spacer=(spacer, ref(1))), [footnote(1, 1)])
for d in (-1.0, 0.0, 1.0):
    # No footnote: Body 47 ends at 720 + d.
    spacer = 720 + d - 72 - 47 * LINE_STEP
    tag = ("m" if d < 0 else "p") + f"{abs(d):.1f}".replace(".", "")
    write(f"limit-nofn-{tag}", body(count=52, spacer=(spacer, "")), [])
print(out)
