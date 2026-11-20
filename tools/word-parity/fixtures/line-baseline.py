#!/usr/bin/env python3
"""Fixture family: where a line's baseline is inside the line, by line spacing, font, size and story.

    python3 tools/word-parity/fixtures/line-baseline.py tests/word/fixtures/line-baseline
    node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/line-baseline --refresh-word
    python3 tools/word-parity/fixtures/line-baseline.py --table tests/word/fixtures/line-baseline/word/*.pdf \
        > tests/word/fixtures/line-baseline/word/baselines.tsv

The showcase handbook's running header (8.5pt Calibri, its mark 11pt) is 1.56pt lower in ours than in Word's
PDF. Each document: Letter, 1in margins, header and footer 0.5in from the edge, nothing before or after any
paragraph, the same spacing, font and size on every paragraph of every story:
  header    "Hdrq"                          its line's top at 36pt
  body      "Bodya", "Bodyb"                the first line's top at 72pt
  cell      "Cella", "Cellb"                a one-cell table under them (no borders, no top margin)
  body      "Refp" with a footnote reference
  footnote  "Notea", "Noteb"                the block's last line's bottom at 720pt
  footer    "Ftrq"                          its line's bottom at 756pt
Names: <font>-<size in half-points>-<spacing>[-mk<mark size>][-c14]. Spacings: single (240 auto), m115 (276),
m15 (360), double (480), exact20 (400 exact), exact10 (200 exact), least20 (400 atLeast), least8 (160 atLeast).
`-mk<n>`: the paragraph marks at another size than the text (the handbook's header: text 17, mark 22).
`-fa12`, `-fb12`, `-f2`: the footer's paragraph 12pt after, 12pt before, or under another one ("Ftrp").
`--table <pdf> ...` writes each probe word's baseline (from the content stream's text matrix: exact) as
<doc> TAB <word> TAB <page> TAB <baseline from the page's top> TAB <size>.
No dependencies: the package is written as XML.
"""
import os, re, sys, zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
SPACINGS = {"single": (240, "auto"), "m115": (276, "auto"), "m15": (360, "auto"), "double": (480, "auto"),
            "exact20": (400, "exact"), "exact10": (200, "exact"), "least20": (400, "atLeast"), "least8": (160, "atLeast")}
FONTS = {"cal": "Calibri", "tnr": "Times New Roman", "ari": "Arial"}
WORDS = ["Hdrq", "Bodya", "Bodyb", "Cella", "Cellb", "Refp", "Notea", "Noteb", "Ftrp", "Ftrq"]


def para(text, spacing, size, mark, extra=""):
    line, rule = SPACINGS[spacing]
    return (f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="{line}" w:lineRule="{rule}"/>'
            f'<w:rPr><w:sz w:val="{mark}"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="{size}"/></w:rPr><w:t>{text}</w:t></w:r>{extra}</w:p>')


def document(spacing, size, mark):
    p = lambda t, extra="": para(t, spacing, size, mark, extra)
    ref = f'<w:r><w:rPr><w:sz w:val="{size}"/><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>'
    table = ('<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
             '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>'
             '</w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>'
             f'{p("Cella")}{p("Cellb")}</w:tc></w:tr></w:tbl>')
    sect = ('<w:sectPr><w:headerReference w:type="default" r:id="rId4"/><w:footerReference w:type="default" r:id="rId5"/>'
            '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>'
            f'{p("Bodya")}{p("Bodyb")}{table}{p("Refp", ref)}{sect}</w:body></w:document>')


def footer_xml(kind, p, spacing, size, mark):
    """The footer's paragraphs: "Ftrq" alone, or with 12pt after it (fa12) or before it (fb12), or under
    another paragraph "Ftrp" (f2)."""
    if kind == "fa12":
        return para("Ftrq", spacing, size, mark).replace('w:after="0"', 'w:after="240"')
    if kind == "fb12":
        return para("Ftrq", spacing, size, mark).replace('w:before="0"', 'w:before="240"')
    if kind == "f2":
        return p("Ftrp") + p("Ftrq")
    return p("Ftrq")


def write(out, name, font, spacing, size, mark, mode, footer=None):
    p = lambda t: para(t, spacing, size, mark)
    styles = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr>'
              f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:eastAsia="{font}" w:cs="{font}"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
              '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
              '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>'
              '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/>'
              '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
              '</w:styles>')
    settings = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>'
                f'<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/></w:compat></w:settings>')
    sep = '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p>'
    csep = '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p>'
    footnotes = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes {W}>'
                 f'<w:footnote w:type="separator" w:id="-1">{sep}</w:footnote><w:footnote w:type="continuationSeparator" w:id="0">{csep}</w:footnote>'
                 f'<w:footnote w:id="1">{p("Notea")}{p("Noteb")}</w:footnote></w:footnotes>')
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
          '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
          '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
          '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             f'<Relationship Id="rId1" Type="{R}styles" Target="styles.xml"/><Relationship Id="rId2" Type="{R}settings" Target="settings.xml"/>'
             f'<Relationship Id="rId3" Type="{R}footnotes" Target="footnotes.xml"/><Relationship Id="rId4" Type="{R}header" Target="header1.xml"/>'
             f'<Relationship Id="rId5" Type="{R}footer" Target="footer1.xml"/></Relationships>')
    with zipfile.ZipFile(os.path.join(out, name + ".docx"), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", rels)
        z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/document.xml", document(spacing, size, mark))
        z.writestr("word/styles.xml", styles); z.writestr("word/settings.xml", settings); z.writestr("word/footnotes.xml", footnotes)
        z.writestr("word/header1.xml", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {W}>{p("Hdrq")}</w:hdr>')
        z.writestr("word/footer1.xml", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr {W}>{footer_xml(footer, p, spacing, size, mark)}</w:ftr>')


def variants():
    for font in ("cal", "tnr"):
        for size in (17, 22, 40):
            for spacing in SPACINGS:
                yield f"{font}-{size}-{spacing}", font, spacing, size, size, 15
    for spacing in SPACINGS:
        yield f"ari-22-{spacing}", "ari", spacing, 22, 22, 15
    for spacing in ("single", "m115", "double"):
        yield f"cal-17-{spacing}-mk22", "cal", spacing, 17, 22, 15
        yield f"cal-22-{spacing}-mk17", "cal", spacing, 22, 17, 15
    for size in (17, 22):
        for spacing in SPACINGS:
            yield f"cal-{size}-{spacing}-c14", "cal", spacing, size, size, 14
    yield "cal-17-m115-mk22-c14", "cal", "m115", 17, 22, 14


# The footer's own spacing: does its last line or its space after end at the footer distance?
FOOTER_VARIANTS = [(f"cal-22-{spacing}-{kind}", "cal", spacing, 22, 22, mode, kind)
                   for spacing in ("single", "m115") for kind in ("fa12", "fb12", "f2") for mode in (15,)] + \
                  [("cal-22-m115-fa12-c14", "cal", "m115", 22, 22, 14, "fa12")]


# --- reading the PDFs: each text show's origin from the content stream ---

def page_streams(pdf):
    import zlib
    data = open(pdf, 'rb').read()
    objects = {int(m.group(1)): m.start() for m in re.finditer(rb'(\d+) 0 obj', data)}
    def body(num):
        start = objects[num]
        return data[start:data.find(b'endobj', start)]
    def stream(num):
        b = body(num)
        s = b.find(b'stream') + len(b'stream')
        s += 2 if b[s:s + 2] == b'\r\n' else 1
        raw = b[s:b.rfind(b'endstream')]
        try:
            return zlib.decompress(raw).decode('latin1')
        except zlib.error:
            return raw.decode('latin1')
    def kids(num):
        b = body(num)
        if re.search(rb'/Type\s*/Page[^s]', b):
            yield num
        m = re.search(rb'/Kids\s*\[([^\]]*)\]', b)
        if m:
            for k in re.findall(rb'(\d+) 0 R', m.group(1)):
                yield from kids(int(k))
    root = (re.search(rb'/Type\s*/Catalog.*?/Pages (\d+) 0 R', data, re.S)
            or re.search(rb'/Pages (\d+) 0 R\s*/Type\s*/Catalog', data, re.S))
    for page, num in enumerate(kids(int(root.group(1))), 1):
        b = body(num)
        m = re.search(rb'/Contents\s*(\[[^\]]*\]|\d+ 0 R)', b)
        box = re.search(rb'/MediaBox\s*\[\s*([\d.\-]+)\s+([\d.\-]+)\s+([\d.\-]+)\s+([\d.\-]+)', b)
        height = float(box.group(4)) if box else 792.0
        yield page, height, ''.join(stream(int(r)) for r in re.findall(rb'(\d+) 0 R', m.group(1)))


TOK = re.compile(r'\[(?:[^\]\\]|\\.)*\]|\((?:[^)\\]|\\.)*\)|<[0-9A-Fa-f\s]*>|/[^\s/\[\]()<>]+|[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?|[A-Za-z*\'"]+')


def mul(a, b):
    return [a[0]*b[0]+a[1]*b[2], a[0]*b[1]+a[1]*b[3], a[2]*b[0]+a[3]*b[2], a[2]*b[1]+a[3]*b[3],
            a[4]*b[0]+a[5]*b[2]+b[4], a[4]*b[1]+a[5]*b[3]+b[5]]


def text_origins(stream):
    """(x, y, size) of each text show: the text matrix under the CTM (Tm / Td / TD / T* / cm / q / Q)."""
    ctm, stack, tm, tlm, size, ops, tl = [1, 0, 0, 1, 0, 0], [], [1, 0, 0, 1, 0, 0], [1, 0, 0, 1, 0, 0], 0.0, [], 0.0
    for t in TOK.findall(stream):
        if t[0] in '[(</' or re.match(r'^[-+]?\d*\.?\d', t):
            ops.append(t); continue
        nums = []
        for x in ops:
            try: nums.append(float(x))
            except ValueError: pass
        if t == 'q': stack.append(ctm[:])
        elif t == 'Q': ctm = stack.pop() if stack else ctm
        elif t == 'cm' and len(nums) >= 6: ctm = mul(nums[-6:], ctm)
        elif t == 'BT': tm = [1, 0, 0, 1, 0, 0]; tlm = tm[:]
        elif t == 'Tm' and len(nums) >= 6: tm = nums[-6:]; tlm = tm[:]
        elif t in ('Td', 'TD') and len(nums) >= 2:
            if t == 'TD': tl = -nums[-1]
            tlm = mul([1, 0, 0, 1, nums[-2], nums[-1]], tlm); tm = tlm[:]
        elif t == 'TL' and nums: tl = nums[-1]
        elif t == 'T*': tlm = mul([1, 0, 0, 1, 0, -tl], tlm); tm = tlm[:]
        elif t == 'Tf' and nums: size = nums[-1]
        elif t in ('Tj', 'TJ', "'", '"'):
            m = mul(tm, ctm)
            yield m[4], m[5], size * abs(m[3])
        ops = []


def word_boxes(pdf):
    import subprocess
    out = subprocess.run(['pdftotext', '-bbox', pdf, '-'], capture_output=True, text=True, check=True).stdout
    page, res = 0, {}
    for line in out.splitlines():
        if '<page ' in line:
            page += 1; res[page] = []
        m = re.search(r'xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">(.*)</word>', line)
        if m:
            res[page].append((float(m.group(1)), float(m.group(2)), float(m.group(4)), m.group(5)))
    return res


def baselines(pdf, wanted=None):
    """(page, word, baseline from the page's top, size) of each text show that starts a pdftotext word."""
    boxes = word_boxes(pdf)
    for page, height, stream in page_streams(pdf):
        seen = set()
        for x, y, size in text_origins(stream):
            top = height - y
            for wx, y0, y1, text in boxes.get(page, []):
                if abs(wx - x) < 1.0 and y0 - 1 <= top <= y1 + 1 and (wanted is None or text in wanted) and (text, round(top, 1)) not in seen:
                    seen.add((text, round(top, 1)))
                    yield page, text, round(top, 3), round(size, 2)
                    break


if __name__ == "__main__":
    if sys.argv[1:2] == ["--table"]:
        for pdf in sys.argv[2:]:
            doc = os.path.splitext(os.path.basename(pdf))[0]
            for page, text, top, size in baselines(pdf, set(WORDS)):
                print(f"{doc}\t{text}\t{page}\t{top}\t{size}")
        sys.exit(0)
    out = sys.argv[1] if len(sys.argv) > 1 else "line-baseline"
    os.makedirs(out, exist_ok=True)
    count = 0
    for name, font, spacing, size, mark, mode in variants():
        write(out, name, FONTS[font], spacing, size, mark, mode); count += 1
    for name, font, spacing, size, mark, mode, footer in FOOTER_VARIANTS:
        write(out, name, FONTS[font], spacing, size, mark, mode, footer); count += 1
    print(out, count)
