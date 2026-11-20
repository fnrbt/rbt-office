#!/usr/bin/env python3
"""Showcase documents: designed .docx files that exercise the Word layout and rendering engine - merged
and nested tables, borders and shading, columns, footnotes, lists, landscape sections - for the demo, for
looking at, and for the parity loop (docs/layout-parity-flow.md: they are corpus documents too).

    python3 -m venv .venv && .venv/bin/pip install python-docx
    .venv/bin/python tools/showcase-docs/make.py [out dir]      (default tests/test-documents-showcase)

All content is invented (Northwind Energy and the rest are fictional). Fonts: Georgia, Calibri, Segoe UI,
Cambria, Arial - the faces the demo has free stand-ins for and Word has on Windows.
"""
import copy, os, sys
from xml.sax.saxutils import escape as esc
from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
from docx.shared import Pt, Cm, RGBColor, Emu

OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(__file__), "..", "..", "tests", "test-documents-showcase")
os.makedirs(OUT, exist_ok=True)

NAVY, TEAL, SKY, MIST, INK, GREY, SAND, CORAL, MOSS, PLUM = "13294B", "0F7B8A", "9CC9E8", "EEF4F7", "1F2933", "5B6770", "F6EFE3", "D9644A", "5E8C4A", "6B4C8A"

# ---------------------------------------------------------------- helpers

def base(body_font="Calibri", heading_font="Georgia", size=10.5):
    doc = Document()
    st = doc.styles
    normal = st["Normal"]
    normal.font.name = body_font
    normal.font.size = Pt(size)
    normal.font.color.rgb = RGBColor.from_string(INK)
    rpr = normal.element.get_or_add_rPr()
    fonts = rpr.find(qn("w:rFonts"))
    for k in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        fonts.set(qn(k), body_font)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.15
    for level, (sz, color) in {1: (20, NAVY), 2: (14, TEAL), 3: (11.5, NAVY)}.items():
        h = st["Heading %d" % level]
        h.font.name = heading_font
        h.font.size = Pt(sz)
        h.font.bold = level != 1
        h.font.color.rgb = RGBColor.from_string(color)
        hf = h.element.get_or_add_rPr().find(qn("w:rFonts"))
        if hf is None:
            hf = OxmlElement("w:rFonts"); h.element.get_or_add_rPr().append(hf)
        for k in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            hf.set(qn(k), heading_font)
        for k in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
            if hf.get(qn(k)) is not None:
                del hf.attrib[qn(k)]
        h.paragraph_format.space_before = Pt(14 if level == 1 else 10)
        h.paragraph_format.space_after = Pt(4)
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2.2)
    sec.top_margin, sec.bottom_margin = Cm(2.0), Cm(2.2)
    return doc

def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:shd")):
        tcPr.remove(old)
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{fill}"/>'))

def borders(cell, **sides):
    """sides: top/bottom/left/right/insideH/insideV = (val, size eighths, color)."""
    tcPr = cell._tc.get_or_add_tcPr()
    b = tcPr.find(qn("w:tcBorders"))
    if b is None:
        b = OxmlElement("w:tcBorders"); tcPr.append(b)
    for side, (val, sz, color) in sides.items():
        e = OxmlElement("w:" + side)
        e.set(qn("w:val"), val); e.set(qn("w:sz"), str(sz)); e.set(qn("w:space"), "0"); e.set(qn("w:color"), color)
        b.append(e)

def table_borders(table, val="single", sz=4, color="BFC8D0", inside=True):
    tblPr = table._tbl.tblPr
    for old in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(old)
    sides = ["top", "left", "bottom", "right"] + (["insideH", "insideV"] if inside else [])
    xml = "".join(f'<w:{s} w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>' for s in sides)
    tblPr.append(parse_xml(f'<w:tblBorders {nsdecls("w")}>{xml}</w:tblBorders>'))

def no_borders(table):
    table_borders(table, val="nil", sz=0, color="auto")

def cell_margins(table, top=50, bottom=50, left=100, right=100):
    tblPr = table._tbl.tblPr
    tblPr.append(parse_xml(f'<w:tblCellMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/>'
                           f'<w:bottom w:w="{bottom}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tblCellMar>'))

def widths(table, cms):
    table.autofit = False
    for row in table.rows:
        for cell, w in zip(row.cells, cms):
            cell.width = Cm(w)

def header_row(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

def cant_split(row):
    row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

def text(cell, s, bold=False, color=None, size=None, align=None, font=None, italic=False, space_after=0):
    p = cell.paragraphs[0] if cell.paragraphs and not cell.paragraphs[0].text and len(cell.paragraphs[0].runs) == 0 else cell.add_paragraph()
    r = p.add_run(s)
    r.bold, r.italic = bold, italic
    if color: r.font.color.rgb = RGBColor.from_string(color)
    if size: r.font.size = Pt(size)
    if font: r.font.name = font
    if align is not None: p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.0
    return p

def vtext(cell, direction="btLr"):
    cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:textDirection {nsdecls("w")} w:val="{direction}"/>'))

def para(doc, s="", style=None, size=None, color=None, bold=False, italic=False, align=None, after=None, before=None, font=None):
    p = doc.add_paragraph(style=style)
    if s:
        r = p.add_run(s)
        r.bold, r.italic = bold, italic
        if size: r.font.size = Pt(size)
        if color: r.font.color.rgb = RGBColor.from_string(color)
        if font: r.font.name = font
    if align is not None: p.alignment = align
    if after is not None: p.paragraph_format.space_after = Pt(after)
    if before is not None: p.paragraph_format.space_before = Pt(before)
    return p

def rule(doc, color=NAVY, sz=12, after=6):
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pPr.append(parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="{sz}" w:space="1" w:color="{color}"/></w:pBdr>'))
    p.paragraph_format.space_after = Pt(after)
    return p

def columns(section, num, space_cm=0.8, sep=False):
    sectPr = section._sectPr
    for old in sectPr.findall(qn("w:cols")):
        sectPr.remove(old)
    sectPr.append(parse_xml(f'<w:cols {nsdecls("w")} w:num="{num}" w:space="{int(space_cm * 567)}"' + (' w:sep="1"' if sep else '') + '/>'))

class Footnotes:
    """A footnotes part for python-docx (which has none): separator, continuation separator, notes."""
    def __init__(self, doc):
        self.doc, self.next = doc, 1
        self.root = parse_xml(
            f'<w:footnotes {nsdecls("w", "r")}>'
            '<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>'
            '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
            '</w:footnotes>')
        styles = doc.styles.element
        styles.append(parse_xml(f'<w:style {nsdecls("w")} w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/>'
                                '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="17"/></w:rPr></w:style>'))
        styles.append(parse_xml(f'<w:style {nsdecls("w")} w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/>'
                                '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>'))
    def add(self, paragraph, note):
        n = self.next; self.next += 1
        paragraph._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="{n}"/></w:r>'))
        self.root.append(parse_xml(
            f'<w:footnote {nsdecls("w")} w:id="{n}"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>'
            f'<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> {esc(note)}</w:t></w:r></w:p></w:footnote>'))
    def save(self):
        from lxml import etree
        part = Part(PackURI("/word/footnotes.xml"), "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
                    etree.tostring(self.root, xml_declaration=True, encoding="UTF-8", standalone=True), self.doc.part.package)
        self.doc.part.relate_to(part, RT.FOOTNOTES)

def save(doc, name):
    path = os.path.join(OUT, name)
    doc.save(path)
    print(path)

# ---------------------------------------------------------------- 1. annual results

def annual_results():
    doc = base("Calibri", "Georgia")
    notes = Footnotes(doc)
    para(doc, "NORTHWIND ENERGY", size=9, color=TEAL, bold=True, after=0)
    para(doc, "Annual results 2026", "Title" if False else None, size=30, color=NAVY, font="Georgia", after=2)
    para(doc, "Generation, capacity and finances across our wind and solar sites", size=12, color=GREY, italic=True, after=10)
    rule(doc)
    # Key figures: four tiles, a borderless table of shaded cells.
    tiles = doc.add_table(rows=2, cols=4)
    no_borders(tiles); cell_margins(tiles, 120, 120, 140, 140)
    figures = [("11,950", "GWh generated"), ("1,284", "MW installed"), ("94.2%", "availability"), ("€412m", "revenue")]
    for i, (big, small) in enumerate(figures):
        top, bottom = tiles.cell(0, i), tiles.cell(1, i)
        shade(top, MIST); shade(bottom, MIST)
        text(top, big, bold=True, size=20, color=NAVY, font="Georgia")
        text(bottom, small, size=9, color=GREY)
    widths(tiles, [4.15] * 4)
    para(doc, after=4)
    doc.add_heading("Generation by site", 1)
    p = para(doc, "Output rose at every site with a full year of operation. Northgate Offshore, connected in June, "
                  "contributed seven months; its availability figure covers that period only.")
    notes.add(p, "Availability is the share of hours a turbine could generate, whatever the wind. Planned maintenance counts against it.")
    # Financial table: a two-level header - "2026" spans the four quarters - and a vertically merged "Site".
    t = doc.add_table(rows=8, cols=7)
    table_borders(t, sz=4, color="D3DAE0"); cell_margins(t, 40, 40, 90, 90)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    header_row(t.rows[0]); header_row(t.rows[1])
    site = t.cell(0, 0).merge(t.cell(1, 0)); text(site, "Site", bold=True, color="FFFFFF")
    year = t.cell(0, 1).merge(t.cell(0, 4)); text(year, "2026, GWh by quarter", bold=True, color="FFFFFF", align=WD_ALIGN_PARAGRAPH.CENTER)
    tot = t.cell(0, 5).merge(t.cell(1, 5)); text(tot, "Year", bold=True, color="FFFFFF", align=WD_ALIGN_PARAGRAPH.RIGHT)
    chg = t.cell(0, 6).merge(t.cell(1, 6)); text(chg, "vs 2025", bold=True, color="FFFFFF", align=WD_ALIGN_PARAGRAPH.RIGHT)
    for c in range(7):
        shade(t.cell(0, c), NAVY); shade(t.cell(1, c), NAVY)
    for q in range(4):
        text(t.cell(1, 1 + q), "Q%d" % (q + 1), bold=True, color=SKY, align=WD_ALIGN_PARAGRAPH.RIGHT)
    rows = [("Ridge Park", [812, 640, 598, 905], "+4.1%"), ("Harbor Point", [700, 551, 530, 874], "+2.7%"),
            ("Millbrook Flats", [455, 380, 362, 613], "+6.0%"), ("Cedar Mesa (solar)", [402, 688, 731, 359], "+11.3%"),
            ("Northgate Offshore", [0, 0, 1180, 1610], "new")]
    for r, (name, qs, d) in enumerate(rows, start=2):
        text(t.cell(r, 0), name, bold=True)
        for q, v in enumerate(qs):
            text(t.cell(r, 1 + q), f"{v:,}" if v else "–", align=WD_ALIGN_PARAGRAPH.RIGHT)
        text(t.cell(r, 5), f"{sum(qs):,}", bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
        text(t.cell(r, 6), d, color=MOSS if d.startswith("+") else GREY, align=WD_ALIGN_PARAGRAPH.RIGHT)
        if r % 2:
            for c in range(7): shade(t.cell(r, c), "F7F9FA")
    last = 7
    text(t.cell(last, 0), "All sites", bold=True, color=NAVY)
    totals = [sum(r[1][q] for r in rows) for q in range(4)]
    for q, v in enumerate(totals):
        text(t.cell(last, 1 + q), f"{v:,}", bold=True, color=NAVY, align=WD_ALIGN_PARAGRAPH.RIGHT)
    text(t.cell(last, 5), f"{sum(totals):,}", bold=True, color=NAVY, align=WD_ALIGN_PARAGRAPH.RIGHT)
    text(t.cell(last, 6), "+5.9%", bold=True, color=NAVY, align=WD_ALIGN_PARAGRAPH.RIGHT)
    for c in range(7):
        borders(t.cell(last, c), top=("double", 6, NAVY))
    widths(t, [4.4, 1.8, 1.8, 1.8, 1.8, 2.0, 2.0])
    p = para(doc, "Figures are metered at the grid connection, after losses in the site's own transformers.", size=8.5, color=GREY, italic=True, before=4)
    notes.add(p, "Cedar Mesa's meter was replaced in March; the quarter's figure includes eleven days estimated from the inverters' own counts.")
    doc.add_heading("Capacity", 1)
    para(doc, "Installed capacity grew by 312 MW. Two projects reached financial close and are listed below with their "
              "expected connection dates; neither is included in the year's figures.")
    # A plain list, then a small two-column table of projects with a merged "Pipeline" label.
    for s in ["Northgate Offshore phase II: 240 MW, foundations complete", "Cedar Mesa storage: 60 MW / 240 MWh batteries",
              "Repowering at Ridge Park: 14 older turbines replaced by 6 larger ones"]:
        doc.add_paragraph(s, style="List Bullet")
    pt = doc.add_table(rows=3, cols=3)
    table_borders(pt, val="single", sz=8, color=TEAL, inside=False); cell_margins(pt, 60, 60, 110, 110)
    lab = pt.cell(0, 0).merge(pt.cell(2, 0)); shade(lab, TEAL); text(lab, "Pipeline", bold=True, color="FFFFFF", size=12, font="Georgia")
    lab.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    for r, (n, d) in enumerate([("Northgate II", "Q3 2027"), ("Cedar Mesa storage", "Q1 2027"), ("Ridge Park repowering", "Q4 2027")]):
        text(pt.cell(r, 1), n, bold=True); text(pt.cell(r, 2), d, align=WD_ALIGN_PARAGRAPH.RIGHT, color=GREY)
        for c in (1, 2):
            if r < 2: borders(pt.cell(r, c), bottom=("dotted", 4, TEAL))
    widths(pt, [3.5, 8.0, 4.0])
    doc.add_heading("Outlook", 1)
    p = para(doc, "We expect generation of 14,500 to 15,200 GWh in 2027, the range set by the wind more than by anything we do. "
                  "Prices for a third of that output are fixed by long-term contracts.")
    notes.add(p, "Contracts with three industrial buyers and one utility, for terms of ten to fifteen years.")
    notes.save()
    save(doc, "annual-results.docx")

# ---------------------------------------------------------------- 2. conference programme

def conference_programme():
    doc = base("Segoe UI", "Georgia", size=9.5)
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)
    sec.left_margin = sec.right_margin = Cm(1.6)
    sec.top_margin, sec.bottom_margin = Cm(1.5), Cm(1.5)
    para(doc, "WINDWORKS 2026  ·  ROTTERDAM  ·  14–15 OCTOBER", size=9, color=CORAL, bold=True, after=0)
    para(doc, "Programme", size=28, color=NAVY, font="Georgia", after=4)
    tracks = [("Main hall", NAVY), ("Room A · Offshore", TEAL), ("Room B · Grids", PLUM)]
    legend = doc.add_table(rows=1, cols=3); no_borders(legend); cell_margins(legend, 30, 30, 80, 80)
    for i, (name, col) in enumerate(tracks):
        shade(legend.cell(0, i), col); text(legend.cell(0, i), name, bold=True, color="FFFFFF", size=9)
    widths(legend, [5.0, 5.0, 5.0])
    para(doc, after=2)
    days = [
        ("Tuesday 14 October", [
            ("08:30", "Registration and coffee", "all", 1),
            ("09:30", "Opening keynote: The next gigawatt — Mara Lindqvist", "all", 1),
            ("10:30", ["Floating foundations at scale", "Workshop: wake losses, measured", "Grid codes after 2030"], "tracks", 2),
            ("11:30", ["Jack-up vessels: who builds the next fleet", None, "Storage as a grid asset"], "tracks", 1),
            ("12:30", "Lunch on the quay", "all", 1),
            ("13:30", ["Corrosion in the splash zone", "Cable faults: finding them faster", "Workshop: build a balancing market"], "tracks", 2),
            ("14:30", ["Decommissioning the first parks", "Inspection by drone", None], "tracks", 1),
            ("15:30", "Plenary: questions to the regulators", "all", 1)]),
        ("Wednesday 15 October", [
            ("09:00", "Keynote: Ports are the bottleneck — Ade Okonkwo", "all", 1),
            ("10:00", ["Workshop: blade recycling, hands on", "Monopile noise and porpoises", "Forecasting for traders"], "tracks", 2),
            ("11:00", [None, "Hydrogen at the turbine", "Panel: who pays for the wires?"], "tracks", 1),
            ("12:00", "Closing panel and farewell lunch", "all", 1)]),
    ]
    for day, slots in days:
        doc.add_heading(day, 2)
        t = doc.add_table(rows=1, cols=4)
        table_borders(t, sz=4, color="FFFFFF"); cell_margins(t, 60, 60, 90, 90)
        hdr = t.rows[0]; header_row(hdr)
        text(hdr.cells[0], "Time", bold=True, color=GREY, size=8.5)
        for i, (name, col) in enumerate(tracks):
            text(hdr.cells[1 + i], name.upper(), bold=True, color=col, size=8.5)
        for c in hdr.cells:
            borders(c, bottom=("single", 12, NAVY))
        # Rows first, merges after (a session of two hours spans two rows: a vertical merge).
        made = []
        for time, what, kind, span in slots:
            if what is None: continue
            row = t.add_row(); cant_split(row)
            made.append((row, time, what, kind, span))
        rows = list(t.rows)
        index = {id(r._tr): i for i, r in enumerate(rows)}
        for row, time, what, kind, span in made:
            r = index[id(row._tr)]
            text(row.cells[0], time, bold=True, color=NAVY)
            if kind == "all":
                cell = t.cell(r, 1).merge(t.cell(r, 3))
                shade(cell, SAND if "Lunch" in what or "coffee" in what else MIST)
                text(cell, what, bold="Keynote" in what or "keynote" in what or "Plenary" in what, color=INK, align=WD_ALIGN_PARAGRAPH.CENTER)
            else:
                for i, s in enumerate(what):
                    cell = t.cell(r, 1 + i)
                    if s is None:
                        shade(cell, "FAFAFA"); text(cell, "—", color="C0C6CC", align=WD_ALIGN_PARAGRAPH.CENTER); continue
                    col = tracks[i][1]
                    shade(cell, {NAVY: "E3E8F0", TEAL: "DDF0F1", PLUM: "ECE4F3"}[col])
                    borders(cell, left=("single", 24, col))
                    text(cell, s, bold=True, size=9.5)
                    text(cell, ["Hall", "Room A", "Room B"][i] + (" · 2 hours" if "Workshop" in s else " · 50 min"), color=GREY, size=8)
        # Two-hour workshops: merge the cell below into the workshop's cell (vertical merge).
        for row, time, what, kind, span in made:
            r = index[id(row._tr)]
            following = [m for m in made if index[id(m[0]._tr)] == r + 1]
            if kind != "all" and span == 2 and following and following[0][3] != "all":
                for i, s in enumerate(what):
                    if s and "Workshop" in s and following[0][2][i] is None:
                        below = t.cell(r + 1, 1 + i)
                        for p in list(below.paragraphs):
                            p._p.getparent().remove(p._p)
                        below._tc.append(parse_xml(f'<w:p {nsdecls("w")}/>'))
                        t.cell(r, 1 + i).merge(below)
        widths(t, [2.0, 7.9, 7.9, 7.9])
    doc.add_heading("Practical", 2)
    para(doc, "Badges are collected at the registration desk in the quay building. The main hall seats 600; "
              "rooms A and B seat 180 each. Workshops are limited to 40 and need a ticket from the desk.")
    save(doc, "conference-programme.docx")

# ---------------------------------------------------------------- 3. field guide

def field_guide():
    doc = base("Cambria", "Georgia", size=10.5)
    notes = Footnotes(doc)
    para(doc, "COASTAL SURVEY · FIELD GUIDE", size=9, color=MOSS, bold=True, after=0)
    para(doc, "Seabirds of the outer banks", size=28, color=NAVY, font="Georgia", after=2)
    para(doc, "A pocket guide for survey crews: what to look for, where, and how to count it.", italic=True, color=GREY, size=12, after=8)
    rule(doc, MOSS)
    p = para(doc, "The banks hold the largest breeding colonies on this coast. Counts are made from the boat at a fixed "
                  "speed along the transect lines, and every bird within 300 metres is recorded with its behaviour.")
    notes.add(p, "Distances are estimated with a rangefinder card held at arm's length; see the crew handbook for its use.")
    # From here: two columns (a continuous section break).
    new = doc.add_section(WD_SECTION.CONTINUOUS)
    columns(new, 2, 0.9, sep=True)
    species = [
        ("Northern gannet", "Morus bassanus", "Large white seabird with black wingtips and a yellow head; plunge-dives from height.",
         [("Length", "87–100 cm"), ("Wingspan", "165–180 cm"), ("Breeds", "Apr–Sep"), ("Status", "Increasing")]),
        ("Atlantic puffin", "Fratercula arctica", "Small auk with a triangular, banded bill in summer; whirring flight low over the water.",
         [("Length", "26–29 cm"), ("Wingspan", "47–63 cm"), ("Breeds", "Apr–Aug"), ("Status", "Declining")]),
        ("Black-legged kittiwake", "Rissa tridactyla", "Gull with dipped-in-ink wingtips and black legs; nests on narrow cliff ledges.",
         [("Length", "37–42 cm"), ("Wingspan", "93–105 cm"), ("Breeds", "May–Aug"), ("Status", "Declining")]),
        ("Great skua", "Stercorarius skua", "Heavy, dark brown seabird with white wing flashes; chases other birds for their catch.",
         [("Length", "50–58 cm"), ("Wingspan", "125–140 cm"), ("Breeds", "May–Aug"), ("Status", "Stable")]),
    ]
    doc.add_heading("Species cards", 2)
    for name, latin, desc, facts in species:
        # A card: an outer table, the description beside a nested table of facts.
        card = doc.add_table(rows=2, cols=1)
        table_borders(card, val="single", sz=6, color="C9D6C2", inside=False); cell_margins(card, 80, 80, 110, 110)
        cant_split(card.rows[0]); cant_split(card.rows[1])
        top = card.cell(0, 0); shade(top, "EEF3EA")
        text(top, name, bold=True, size=12, color=NAVY, font="Georgia")
        text(top, latin, italic=True, color=GREY, size=9)
        body = card.cell(1, 0)
        text(body, desc, size=9.5, space_after=4)
        inner = body.add_table(rows=len(facts), cols=2)
        table_borders(inner, val="single", sz=2, color="DCE3D6"); cell_margins(inner, 20, 20, 70, 70)
        for i, (k, v) in enumerate(facts):
            text(inner.cell(i, 0), k, color=GREY, size=8.5); text(inner.cell(i, 1), v, bold=True, size=8.5,
                                                                      color=CORAL if v == "Declining" else (MOSS if v == "Increasing" else INK))
        widths(inner, [2.2, 3.4])
        body.add_paragraph().paragraph_format.space_after = Pt(0)
        para(doc, after=2)
    doc.add_heading("Counting method", 2)
    p = para(doc, "Count in blocks of one minute. Birds on the water are counted continuously; birds in flight are "
                  "counted in a snapshot at the start of each minute, so that a bird circling the boat is not counted twice.")
    notes.add(p, "The snapshot method follows the standard of the regional seabird programme, revised in 2024.")
    # A shaded sidebar: a one-cell table with a thick left rule.
    side = doc.add_table(rows=1, cols=1); no_borders(side); cell_margins(side, 100, 100, 160, 120)
    c = side.cell(0, 0); shade(c, SAND); borders(c, left=("single", 36, CORAL))
    text(c, "On the boat", bold=True, color=CORAL, size=10.5, space_after=3)
    for s in ["Keep the transect speed at 10 knots.", "Record sea state every 30 minutes.",
              "Stop counting in fog below 1 km visibility.", "Note every vessel within 2 km."]:
        q = c.add_paragraph(s, style="List Bullet"); q.paragraph_format.space_after = Pt(1)
    para(doc, after=2)
    for s in ["Behaviour codes are entered with each count: S sitting, F flying, D diving, C carrying food.",
              "Unidentified auks are recorded as such, never guessed.",
              "Photographs are taken of any ringed bird, with the ring's position noted."]:
        doc.add_paragraph(s, style="List Number")
    notes.save()
    save(doc, "field-guide.docx")

# ---------------------------------------------------------------- 4. product specification

def product_spec():
    doc = base("Arial", "Arial", size=9.5)
    para(doc, "TECHNICAL DATA SHEET", size=8.5, color=GREY, bold=True, after=0)
    para(doc, "NW-6.2 wind turbine", size=24, color=NAVY, bold=True, after=0)
    para(doc, "6.2 MW · 170 m rotor · IEC class S", size=11, color=TEAL, after=8)
    # An outer two-column layout table: the left cell holds nested spec tables, the right a summary panel.
    layout = doc.add_table(rows=1, cols=2); no_borders(layout); cell_margins(layout, 0, 0, 0, 160)
    left, right = layout.cell(0, 0), layout.cell(0, 1)
    sections = [
        ("Rotor", [("Diameter", "170 m"), ("Swept area", "22,698 m²"), ("Blades", "3, carbon spar caps"), ("Speed range", "4.9–10.6 rpm")]),
        ("Operating data", [("Cut-in wind speed", "3 m/s"), ("Rated wind speed", "11.5 m/s"), ("Cut-out wind speed", "25 m/s"), ("Re-cut-in", "22 m/s")]),
        ("Electrical", [("Generator", "Permanent magnet, medium speed"), ("Converter", "Full scale"), ("Voltage", "690 V / 66 kV"), ("Frequency", "50 / 60 Hz")]),
    ]
    first = True
    for title, rows in sections:
        p = left.paragraphs[0] if first else left.add_paragraph(); first = False
        r = p.add_run(title.upper()); r.bold = True; r.font.color.rgb = RGBColor.from_string(TEAL); r.font.size = Pt(9)
        p.paragraph_format.space_after = Pt(2); p.paragraph_format.space_before = Pt(6)
        t = left.add_table(rows=len(rows), cols=2)
        table_borders(t, val="single", sz=4, color="E1E6EA", inside=True); cell_margins(t, 30, 30, 80, 80)
        for i, (k, v) in enumerate(rows):
            text(t.cell(i, 0), k, color=GREY); text(t.cell(i, 1), v, bold=True)
            if i % 2 == 0:
                shade(t.cell(i, 0), "F4F6F8"); shade(t.cell(i, 1), "F4F6F8")
        widths(t, [4.4, 5.4])
    left.add_paragraph().paragraph_format.space_after = Pt(0)
    shade(right, NAVY)
    text(right, "At a glance", bold=True, color="FFFFFF", size=12, space_after=6)
    for big, small in [("6.2 MW", "rated power"), ("170 m", "rotor diameter"), ("25 yr", "design life"), ("52%", "capacity factor at 9 m/s")]:
        text(right, big, bold=True, color=SKY, size=18)
        text(right, small, color="DCE6F0", size=8.5, space_after=6)
    right.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    widths(layout, [10.4, 6.2])
    para(doc, after=4)
    doc.add_heading("Power curve", 2)
    # Vertical header text, double outer border, thick rule under the header.
    pc = doc.add_table(rows=3, cols=11)
    table_borders(pc, val="double", sz=4, color=NAVY, inside=False); cell_margins(pc, 30, 30, 50, 50)
    lab = pc.cell(0, 0).merge(pc.cell(2, 0)); vtext(lab); shade(lab, MIST)
    text(lab, "Power curve", bold=True, color=NAVY, size=8.5)
    for i, v in enumerate(range(3, 23, 2)):
        text(pc.cell(0, 1 + i), f"{v} m/s", bold=True, size=8, align=WD_ALIGN_PARAGRAPH.CENTER)
        borders(pc.cell(0, 1 + i), bottom=("single", 12, NAVY))
        kw = min(6200, int(6200 * max(0, (v - 3) / 8.5) ** 2.2)) if v < 25 else 0
        text(pc.cell(1, 1 + i), f"{kw:,}", size=8.5, align=WD_ALIGN_PARAGRAPH.CENTER)
        text(pc.cell(2, 1 + i), f"{min(0.49, 0.18 + 0.03 * i):.2f}", size=8.5, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER)
    widths(pc, [1.0] + [1.56] * 10)
    para(doc, "Row 1: electrical power in kW; row 2: power coefficient Cp. Air density 1.225 kg/m³.", size=8, color=GREY, italic=True, before=3)
    doc.add_heading("Options", 2)
    opt = doc.add_table(rows=4, cols=3)
    table_borders(opt, sz=4, color="D3DAE0"); cell_margins(opt, 50, 50, 90, 90)
    for c, h in enumerate(["Option", "Includes", "Availability"]):
        shade(opt.cell(0, c), TEAL); text(opt.cell(0, c), h, bold=True, color="FFFFFF")
    options = [("Cold climate", ["Blade heating", "Low-temperature steel", "Heated nacelle"], "All markets"),
               ("Low noise", ["Serrated trailing edges", "Noise-reduced modes 1–4"], "Europe"),
               ("Grid support", ["Fault ride-through to 0 V", "Synthetic inertia", "Black start"], "On request")]
    for r, (name, inc, avail) in enumerate(options, start=1):
        text(opt.cell(r, 0), name, bold=True)
        for i, s in enumerate(inc):
            p = opt.cell(r, 1).paragraphs[0] if i == 0 else opt.cell(r, 1).add_paragraph()
            p.style = doc.styles["List Bullet"]; p.add_run(s); p.paragraph_format.space_after = Pt(0)
        text(opt.cell(r, 2), avail, color=GREY)
    widths(opt, [3.8, 8.8, 4.0])
    save(doc, "product-spec.docx")

# ---------------------------------------------------------------- 5. invoice

def invoice():
    doc = base("Segoe UI", "Segoe UI", size=9.5)
    head = doc.add_table(rows=1, cols=2); no_borders(head)
    l, r = head.cell(0, 0), head.cell(0, 1)
    text(l, "Harbor Point Services", bold=True, size=16, color=NAVY)
    text(l, "Quayside 4, 3011 Rotterdam", color=GREY); text(l, "VAT NL 8123 4567 B01", color=GREY)
    text(r, "INVOICE", bold=True, size=22, color=CORAL, align=WD_ALIGN_PARAGRAPH.RIGHT)
    # Invoice details as a nested table, right-aligned in its cell.
    det = r.add_table(rows=3, cols=2); no_borders(det)
    for i, (k, v) in enumerate([("Number", "2026-0417"), ("Date", "18 November 2026"), ("Due", "18 December 2026")]):
        text(det.cell(i, 0), k, color=GREY, align=WD_ALIGN_PARAGRAPH.RIGHT); text(det.cell(i, 1), v, bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
    widths(det, [2.6, 4.2])
    r.add_paragraph().paragraph_format.space_after = Pt(0)
    widths(head, [9.0, 7.6])
    para(doc, after=10)
    to = doc.add_table(rows=1, cols=2); no_borders(to); cell_margins(to, 80, 80, 120, 120)
    a, b = to.cell(0, 0), to.cell(0, 1)
    shade(a, MIST); shade(b, MIST)
    text(a, "BILL TO", bold=True, color=TEAL, size=8); text(a, "Northwind Energy B.V.", bold=True)
    text(a, "Accounts payable"); text(a, "Windstraat 12, 1012 Amsterdam")
    text(b, "SITE", bold=True, color=TEAL, size=8); text(b, "Northgate Offshore", bold=True)
    text(b, "Purchase order 77-3301"); text(b, "Crew transfer, October")
    widths(to, [8.3, 8.3])
    para(doc, after=8)
    items = [("Crew transfer vessel, day rate", 19, 2450.00), ("Fuel surcharge", 19, 310.00),
             ("Standby, weather days", 4, 1225.00), ("Additional technicians on board", 22, 85.00),
             ("Port fees, Eemshaven", 1, 1840.00)]
    t = doc.add_table(rows=1, cols=4)
    no_borders(t); cell_margins(t, 60, 60, 90, 90)
    for c, h in enumerate(["Description", "Qty", "Unit price", "Amount"]):
        text(t.rows[0].cells[c], h, bold=True, color=GREY, size=8.5, align=None if c == 0 else WD_ALIGN_PARAGRAPH.RIGHT)
        borders(t.rows[0].cells[c], bottom=("single", 12, NAVY))
    total = 0
    for i, (d, q, u) in enumerate(items):
        row = t.add_row().cells
        amount = q * u; total += amount
        text(row[0], d); text(row[1], str(q), align=WD_ALIGN_PARAGRAPH.RIGHT)
        text(row[2], f"€{u:,.2f}", align=WD_ALIGN_PARAGRAPH.RIGHT); text(row[3], f"€{amount:,.2f}", align=WD_ALIGN_PARAGRAPH.RIGHT)
        for c in row:
            borders(c, bottom=("single", 2, "E1E6EA"))
    vat = total * 0.21
    for label, value, strong in [("Subtotal", total, False), ("VAT 21%", vat, False), ("Total due", total + vat, True)]:
        row = t.add_row()
        lab = row.cells[0].merge(row.cells[2])
        text(lab, label, bold=strong, color=NAVY if strong else GREY, align=WD_ALIGN_PARAGRAPH.RIGHT, size=11 if strong else None)
        text(row.cells[3], f"€{value:,.2f}", bold=strong, color=NAVY if strong else None, align=WD_ALIGN_PARAGRAPH.RIGHT, size=11 if strong else None)
        if strong:
            shade(row.cells[3], MIST); borders(row.cells[3], top=("single", 12, NAVY))
    widths(t, [9.0, 1.6, 3.0, 3.0])
    para(doc, after=14)
    para(doc, "Payment to IBAN NL91 ABNA 0417 1643 00 within 30 days, quoting the invoice number.", color=GREY, size=9)
    para(doc, "Thank you for your business.", italic=True, color=TEAL)
    save(doc, "invoice.docx")


# ---------------------------------------------------------------- drawing helpers (DrawingML written out)

ASSETS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
DML = ('xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
       'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
       'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
       'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
       'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
       'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"')
_ids = [100]
def _id():
    _ids[0] += 1
    return _ids[0]

def _anchor(inner_graphic, cx, cy, h, v, wrap, behind=False, dist=114300):
    """A floating drawing: h / v are (relativeFrom, align or offset EMU)."""
    def pos(tag, rel, val):
        body = f"<wp:align>{val}</wp:align>" if isinstance(val, str) else f"<wp:posOffset>{int(val)}</wp:posOffset>"
        return f'<wp:position{tag} relativeFrom="{rel}">{body}</wp:position{tag}>'
    wraps = {"square": '<wp:wrapSquare wrapText="bothSides"/>', "topAndBottom": "<wp:wrapTopAndBottom/>", "none": "<wp:wrapNone/>"}
    i = _id()
    return (f'<wp:anchor {DML} distT="{dist}" distB="{dist}" distL="{dist}" distR="{dist}" simplePos="0" relativeHeight="{251658240 + i}" '
            f'behindDoc="{1 if behind else 0}" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
            f'{pos("H", *h)}{pos("V", *v)}<wp:extent cx="{int(cx)}" cy="{int(cy)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'{wraps[wrap]}<wp:docPr id="{i}" name="Drawing {i}"/><wp:cNvGraphicFramePr/>{inner_graphic}</wp:anchor>')

def float_picture(doc, paragraph, path, width_cm, h=("column", "right"), v=("paragraph", 0), wrap="square"):
    """A picture anchored in `paragraph`, floating with text wrapped round it."""
    run = paragraph.add_run()
    inline = run.add_picture(path, width=Cm(width_cm))._inline
    cx, cy = inline.extent.cx, inline.extent.cy
    graphic = inline.find(qn("a:graphic"))
    from lxml import etree
    g = etree.tostring(graphic).decode()
    drawing = inline.getparent()
    drawing.remove(inline)
    drawing.append(parse_xml(_anchor(g, cx, cy, h, v, wrap)))
    return run

def shape(paragraph, geometry, w_cm, h_cm, x_cm, y_cm, fill, line=None, lines=(), color="FFFFFF", size=10, bold=True,
          wrap="none", rel=("column", "paragraph"), align="center", radius=None):
    """A preset shape (rect, roundRect, rightArrow, ellipse, wedgeRectCallout, ...) with text, floating."""
    cx, cy = Cm(w_cm), Cm(h_cm)
    ln = f'<a:ln w="{int(line[1] * 12700)}"><a:solidFill><a:srgbClr val="{line[0]}"/></a:solidFill></a:ln>' if line else '<a:ln><a:noFill/></a:ln>'
    fill_xml = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>' if fill else '<a:noFill/>'
    jc = {"center": "center", "left": "left", "right": "right"}[align]
    paras = "".join(
        f'<w:p><w:pPr><w:jc w:val="{jc}"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
        f'<w:r><w:rPr>{"<w:b/>" if bold and k == 0 else ""}<w:color w:val="{color}"/><w:sz w:val="{int((size if k == 0 else size - 1.5) * 2)}"/></w:rPr>'
        f'<w:t xml:space="preserve">{esc(t)}</w:t></w:r></w:p>' for k, t in enumerate(lines))
    txbx = f'<wps:txbx><w:txbxContent>{paras}</w:txbxContent></wps:txbx>' if lines else ""
    graphic = (f'<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>'
               f'<wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{int(cx)}" cy="{int(cy)}"/></a:xfrm>'
               f'<a:prstGeom prst="{geometry}"><a:avLst/></a:prstGeom>{fill_xml}{ln}</wps:spPr>{txbx}'
               f'<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="72000" tIns="36000" rIns="72000" bIns="36000" anchor="ctr" anchorCtr="0"><a:noAutofit/></wps:bodyPr>'
               f'</wps:wsp></a:graphicData></a:graphic>')
    run = paragraph.add_run()
    run._r.append(parse_xml(f'<w:drawing {DML}>{_anchor(graphic, cx, cy, (rel[0], Cm(x_cm)), (rel[1], Cm(y_cm)), wrap, dist=0)}</w:drawing>'))
    return run

def drop_cap(paragraph, lines=3, font="Georgia", color=NAVY):
    """Word's drop cap: the paragraph's first letter in a frame of its own, `lines` lines tall."""
    first, rest = paragraph.runs[0].text[0], paragraph.runs[0].text[1:]
    paragraph.runs[0].text = rest
    cap = paragraph._p.getparent()
    letter = parse_xml(f'<w:p {nsdecls("w")}><w:pPr><w:keepNext/><w:framePr w:dropCap="drop" w:lines="{lines}" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/>'
                       f'<w:spacing w:after="0" w:line="{lines * 252}" w:lineRule="exact"/><w:textAlignment w:val="baseline"/></w:pPr>'
                       f'<w:r><w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}"/><w:color w:val="{color}"/><w:position w:val="-7"/><w:sz w:val="{lines * 28 + 8}"/></w:rPr><w:t>{first}</w:t></w:r></w:p>')
    paragraph._p.addprevious(letter)

def field(paragraph, instr, result="1"):
    for xml in (f'<w:r {nsdecls("w")}><w:fldChar w:fldCharType="begin"/></w:r>',
                f'<w:r {nsdecls("w")}><w:instrText xml:space="preserve"> {instr} </w:instrText></w:r>',
                f'<w:r {nsdecls("w")}><w:fldChar w:fldCharType="separate"/></w:r>',
                f'<w:r {nsdecls("w")}><w:t>{esc(result)}</w:t></w:r>',
                f'<w:r {nsdecls("w")}><w:fldChar w:fldCharType="end"/></w:r>'):
        paragraph._p.append(parse_xml(xml))

def justify_with_hyphenation(doc):
    doc.settings.element.append(parse_xml(f'<w:autoHyphenation {nsdecls("w")}/>'))

LOREM = {
    "wind": ["The first turbines on the ridge went up in 1998, nine machines of 600 kilowatts each, and for most of the next decade they were the only thing on the skyline that moved. Farmers leased the land a hectare at a time; the rent paid for fences and, in one case, a new barn roof.",
             "What changed was the size. A modern turbine sweeps an area bigger than two football pitches, and a single one of them produces more in a year than the whole of the original park. Repowering the ridge means taking fourteen of the old machines down and putting six up in their place.",
             "The engineers talk about capacity factors and wake losses; the people who live under the ridge talk about the noise at night and the shadow flicker in the kitchen in late afternoon. Both conversations end up at the same place: where exactly each tower stands.",
             "Moving a turbine fifty metres is a decision about wind, about birds, about the line of sight from the village church and about the cost of the road that has to reach it. The planning file for the repowering runs to eleven hundred pages, most of it maps.",
             "There is a quieter argument about what happens to the old blades. Steel towers are cut up and sold; blades are glass fibre and resin, and until recently the only place for them was a landfill. The new contract requires every blade to be recycled, ground into filler for cement."],
}

# ---------------------------------------------------------------- 6. magazine feature

def magazine_feature():
    doc = base("Georgia", "Georgia", size=10)
    justify_with_hyphenation(doc)
    para(doc, "FEATURE  ·  ISSUE 15  ·  WINTER 2026", size=8.5, color=CORAL, bold=True, after=2, font="Segoe UI")
    para(doc, "Taller, fewer, quieter", size=32, color=NAVY, after=2)
    para(doc, "Repowering the ridge: what it takes to replace a wind farm with a bigger one, and who gets a say.", size=13, color=GREY, italic=True, after=8)
    pic = para(doc, after=2)
    pic.add_run().add_picture(os.path.join(ASSETS, "turbines.jpg"), width=Cm(16.6))
    para(doc, "Evening on the ridge: the old machines, due to come down next spring.", size=8, color=GREY, italic=True, after=10, font="Segoe UI")
    sec = doc.add_section(WD_SECTION.CONTINUOUS)
    columns(sec, 2, 0.8)
    paragraphs = LOREM["wind"] * 2
    for i, t in enumerate(paragraphs):
        p = para(doc, t, after=6, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
        p.paragraph_format.first_line_indent = Cm(0 if i == 0 else 0.5)
        if i == 0:
            drop_cap(p)
        if i == 2:
            # A pull quote: a floating text box in the column, text above and below it.
            shape(p, "rect", 7.6, 2.6, 0, 0.2, "F6EFE3", line=(CORAL, 0), lines=["“Every blade will be ground into filler for cement.”", "— the new contract"],
                  color=NAVY, size=13, bold=False, wrap="topAndBottom")
        if i == 4:
            float_picture(doc, p, os.path.join(ASSETS, "otter.jpg"), 4.2, h=("column", "right"), v=("paragraph", Cm(0.2)), wrap="square")
        if i == 6:
            doc.add_heading("Who decides", 3)
    save(doc, "magazine-feature.docx")

# ---------------------------------------------------------------- 7. diagrams

def diagrams():
    doc = base("Segoe UI", "Georgia", size=10)
    para(doc, "HOW A TURBINE GETS BUILT", size=8.5, color=TEAL, bold=True, after=0)
    para(doc, "From permit to power", size=26, color=NAVY, font="Georgia", after=10)
    para(doc, "Five stages, each with its own sign-off. The arrows are the hand-overs where most delays happen.", after=6)
    # A row of stage boxes joined by arrows, anchored to an empty paragraph tall enough to hold them.
    holder = para(doc, after=0)
    holder.paragraph_format.space_after = Pt(0)
    stages = [("Permit", "12–24 months", NAVY), ("Foundations", "6 weeks", TEAL), ("Tower", "3 days", PLUM), ("Nacelle & rotor", "2 days", CORAL), ("Grid test", "2 weeks", MOSS)]
    x = 0.0
    for i, (name, time, col) in enumerate(stages):
        shape(holder, "roundRect", 2.6, 1.8, x, 0.2, col, lines=[name, time], size=10)
        x += 2.6
        if i < len(stages) - 1:
            shape(holder, "rightArrow", 0.9, 0.7, x + 0.1, 0.75, "C9D2DA")
            x += 1.1
    for _ in range(5):
        para(doc, after=0)
    doc.add_heading("Who signs what", 2)
    para(doc, "A callout marks the one stage that can stop everything after it.", after=4)
    holder2 = para(doc, after=0)
    shape(holder2, "ellipse", 4.0, 2.2, 0.5, 0.3, MIST, line=(NAVY, 1), lines=["Grid operator", "connection offer"], color=NAVY, size=10)
    shape(holder2, "wedgeRectCallout", 6.5, 2.0, 6.2, 0.1, "FFF4D6", line=("E0B84A", 1), lines=["No offer, no build:", "the queue is now three years long."], color=INK, size=10, bold=True, align="left")
    shape(holder2, "rect", 4.0, 2.2, 13.2, 0.3, None, line=(GREY, 1.5), lines=["Local authority", "planning consent"], color=INK, size=10)
    for _ in range(5):
        para(doc, after=0)
    doc.add_heading("On site", 2)
    p = para(doc, "The crane that lifts the nacelle is the largest piece of equipment on site, and it is booked months ahead. "
                  "If the wind is above 12 metres per second at hub height the lift waits; a week of bad weather can push the "
                  "whole schedule into the next month, because the crane has another site to go to. Project managers keep a "
                  "spreadsheet of forecast wind speeds for every lift window, and the most experienced of them say the "
                  "spreadsheet is the most important document on the project.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    float_picture(doc, p, os.path.join(ASSETS, "turbines.jpg"), 6.0, h=("margin", "left"), v=("paragraph", 0), wrap="square")
    para(doc, "Behind the text: a shape placed behind it, the text running over it, as a watermark does.", after=4)
    wm = para(doc, "DRAFT — the sequence above is simplified; every site has its own.", italic=True, color=GREY, after=6)
    shape(wm, "rect", 16.6, 1.0, 0, -0.15, "EEF4F7", lines=[], wrap="none")
    wm.runs[-1]._r.find(qn("w:drawing")).find(qn("wp:anchor")).set("behindDoc", "1")
    save(doc, "diagrams.docx")

# ---------------------------------------------------------------- 8. handbook (long: contents, headers, sections, endnotes)

def handbook():
    doc = base("Calibri", "Cambria", size=11)
    sec = doc.sections[0]
    sec.different_first_page_header_footer = True
    hp = sec.header.paragraphs[0]; hp.text = "Field Operations Handbook"; hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hp.runs[0].font.size = Pt(8.5); hp.runs[0].font.color.rgb = RGBColor.from_string(GREY)
    fp = sec.footer.paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = fp.add_run("Page "); r.font.size = Pt(9); r.font.color.rgb = RGBColor.from_string(GREY)
    field(fp, "PAGE", "1")
    para(doc, after=120)
    para(doc, "Field Operations", size=36, color=NAVY, font="Cambria", after=0)
    para(doc, "Handbook", size=36, color=TEAL, font="Cambria", after=12)
    para(doc, "Northwind Energy · edition 4 · 2026", size=12, color=GREY, after=0)
    doc.add_page_break()
    para(doc, "Contents", size=20, color=NAVY, font="Cambria", after=8)
    toc = para(doc)
    field(toc, 'TOC \\o "1-2" \\h \\z \\u', "Update the table of contents (right-click, Update field) to list the chapters.")
    chapters = ["Before the visit", "On site", "Working at height", "Electrical safety", "Faults and incidents", "After the visit"]
    body = LOREM["wind"]
    for n, title in enumerate(chapters, start=1):
        h = doc.add_heading(f"{n}. {title}", 1)
        h.paragraph_format.page_break_before = True
        for k in range(3):
            doc.add_heading(f"{n}.{k + 1} " + ["Planning", "Checks", "Records"][k], 2)
            for j in range(3):
                para(doc, body[(n + k + j) % len(body)], align=WD_ALIGN_PARAGRAPH.JUSTIFY)
            if k == 1:
                for s in ["Confirm the permit is valid for the date.", "Check the weather window twice: the day before and at the yard.", "Brief the crew on the site's own hazards."]:
                    doc.add_paragraph(s, style="List Number")
        if n == 3:
            # A landscape section for a wide table, then back to portrait.
            wide = doc.add_section(WD_SECTION.NEW_PAGE)
            wide.orientation = WD_ORIENT.LANDSCAPE
            wide.page_width, wide.page_height = Cm(29.7), Cm(21.0)
            doc.add_heading("Rescue kit, by site", 2)
            t = doc.add_table(rows=6, cols=7)
            table_borders(t, sz=4, color="D3DAE0"); cell_margins(t, 40, 40, 80, 80)
            heads = ["Site", "Rescue kit", "Last inspected", "Harnesses", "Descender", "Stretcher", "Notes"]
            for c, hd in enumerate(heads):
                shade(t.cell(0, c), NAVY); text(t.cell(0, c), hd, bold=True, color="FFFFFF")
            header_row(t.rows[0])
            sites = ["Ridge Park", "Harbor Point", "Millbrook Flats", "Cedar Mesa", "Northgate Offshore"]
            for r, site in enumerate(sites, start=1):
                vals = [site, "Type B", f"{r + 3} Sep 2026", str(4 + r), "Yes", "Yes" if r % 2 else "On order", "Kit stored at the base of tower 1" if r == 1 else ""]
                for c, v in enumerate(vals):
                    text(t.cell(r, c), v, bold=(c == 0))
            widths(t, [3.6, 2.6, 3.0, 2.4, 2.4, 2.6, 8.4])
            back = doc.add_section(WD_SECTION.NEW_PAGE)
            back.orientation = WD_ORIENT.PORTRAIT
            back.page_width, back.page_height = Cm(21.0), Cm(29.7)
    save(doc, "handbook.docx")


# ---------------------------------------------------------------- embedded fonts (Word's "Embed fonts in the file")

import re, uuid, zipfile
FONTS = os.path.join(ASSETS, "fonts")

def _obfuscate(data, key):
    """ECMA-376 Part 2 font obfuscation (its own inverse): the first 32 bytes XOR the key GUID's 16 bytes, last byte first."""
    k = bytes.fromhex(key.replace("-", ""))[::-1]
    head = bytes(b ^ k[i % 16] for i, b in enumerate(data[:32]))
    return head + data[32:]

def embed_fonts(path, families):
    """Embed TrueType faces in a saved .docx as Word does: families = {"Family": {"Regular": "file.ttf", "Bold": ...}}."""
    src = zipfile.ZipFile(path)
    items = {n: src.read(n) for n in src.namelist()}
    src.close()
    table = items["word/fontTable.xml"].decode()
    rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    n = 0
    for family, faces in families.items():
        embeds = ""
        for face, file in faces.items():
            n += 1
            key = str(uuid.uuid4()).upper()
            items[f"word/fonts/font{n}.odttf"] = _obfuscate(open(os.path.join(FONTS, file), "rb").read(), key)
            rels += f'<Relationship Id="rId{n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font{n}.odttf"/>'
            embeds += f'<w:embed{face} r:id="rId{n}" w:fontKey="{{{key}}}"/>'
        entry = (f'<w:font w:name="{family}"><w:panose1 w:val="00000000000000000000"/><w:charset w:val="00"/>'
                 f'<w:family w:val="auto"/><w:pitch w:val="variable"/>{embeds}</w:font>')
        table = re.sub(rf'<w:font w:name="{re.escape(family)}">.*?</w:font>', "", table, flags=re.S)
        table = table.replace("</w:fonts>", entry + "</w:fonts>")
    rels += "</Relationships>"
    if 'xmlns:r=' not in table[:400]:
        table = table.replace("<w:fonts ", '<w:fonts xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ', 1)
    items["word/fontTable.xml"] = table.encode()
    items["word/_rels/fontTable.xml.rels"] = rels.encode()
    ct = items["[Content_Types].xml"].decode()
    if 'Extension="odttf"' not in ct:
        ct = ct.replace("<Default ", '<Default Extension="odttf" ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont"/><Default ', 1)
    items["[Content_Types].xml"] = ct.encode()
    settings = items["word/settings.xml"].decode()
    if "embedTrueTypeFonts" not in settings:
        settings = re.sub(r"(<w:settings[^>]*>)", r"\1<w:embedTrueTypeFonts/>", settings, count=1)
    items["word/settings.xml"] = settings.encode()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in items.items():
            z.writestr(name, data)

# ---------------------------------------------------------------- 9. exhibition leaflet (embedded fonts)

def exhibition_leaflet():
    SERIF, SANS = "Instrument Serif", "Poppins"
    doc = base(SANS, SERIF, size=9.5)
    for level in (1, 2, 3):
        doc.styles["Heading %d" % level].font.bold = False
    doc.styles["Heading 2"].font.name = SANS
    doc.styles["Heading 2"].font.bold = True
    hf = doc.styles["Heading 2"].element.get_or_add_rPr().find(qn("w:rFonts"))
    for k in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        hf.set(qn(k), SANS)
    para(doc, "THE WEATHER MUSEUM  ·  SPRING EXHIBITION", size=8.5, color=CORAL, bold=True, after=0)
    p = para(doc, "Weather", size=64, color=NAVY, font=SERIF, after=0)
    p.paragraph_format.line_spacing = 0.9
    r = p.add_run(" you can see"); r.italic = True; r.font.size = Pt(64); r.font.name = SERIF; r.font.color.rgb = RGBColor.from_string(TEAL)
    para(doc, "Clouds, storms and the people who drew them, 1800 to now. Eleven rooms, two hundred works, one very large barometer.",
         size=13, color=GREY, font=SERIF, italic=True, after=10)
    rule(doc, CORAL, sz=8)
    intro = doc.add_table(rows=1, cols=3); no_borders(intro); cell_margins(intro, 60, 60, 0, 160)
    blocks = [("Where", "West wing, rooms 4–14", "Lift to the first floor; step-free throughout"),
              ("When", "14 March – 2 August", "Daily 10:00–17:30, Thursdays until 21:00"),
              ("Tickets", "€14, concessions €9", "Free for members and under-18s")]
    for i, (h, big, small) in enumerate(blocks):
        c = intro.cell(0, i)
        text(c, h.upper(), bold=True, size=7.5, color=CORAL)
        text(c, big, size=15, font=SERIF, color=NAVY)
        text(c, small, size=8.5, color=GREY)
    widths(intro, [5.5, 5.5, 5.6])
    para(doc, after=6)
    doc.add_heading("The rooms", 2)
    rooms = [("4", "Before the names", "Clouds in landscape painting before Luke Howard classified them in 1803."),
             ("5", "Cirrus, cumulus, stratus", "Howard's own watercolours, and the letters Goethe wrote him about them."),
             ("6–7", "The storm", "Shipwrecks, squalls and a hurricane drawn from the deck of the ship that survived it."),
             ("8", "Instruments", "Barometers, hygrometers and the eleven-metre water barometer, working."),
             ("9–10", "The forecast", "The first weather maps, telegraphed in from fifty stations at eight each morning."),
             ("11–14", "Seen from above", "Satellite images, from the first fuzzy frame in 1960 to this morning's.")]
    t = doc.add_table(rows=len(rooms), cols=3)
    table_borders(t, val="nil", sz=0, color="auto"); cell_margins(t, 70, 70, 0, 120)
    for i, (num, name, what) in enumerate(rooms):
        text(t.cell(i, 0), num, font=SERIF, size=22, color=TEAL)
        text(t.cell(i, 1), name, bold=True, size=10)
        text(t.cell(i, 2), what, size=9, color=INK)
        for c in range(3):
            borders(t.cell(i, c), bottom=("single", 4, "D9DEE3"))
    widths(t, [1.8, 4.6, 10.2])
    para(doc, after=6)
    q = doc.add_table(rows=1, cols=1); no_borders(q); cell_margins(q, 160, 160, 300, 300)
    c = q.cell(0, 0); shade(c, SAND)
    text(c, "“I have attempted to give names to the clouds, which before had none.”", font=SERIF, size=20, italic=True, color=NAVY, space_after=4)
    text(c, "LUKE HOWARD, ON THE MODIFICATIONS OF CLOUDS, 1803", size=7.5, bold=True, color=CORAL)
    para(doc, after=6)
    doc.add_heading("Talks and walks", 2)
    for when, what in [("Thursdays, 19:00", "Late talks with the curators, in the barometer room"),
                       ("Saturdays, 11:00", "Cloud-spotting walk on the roof terrace, weather permitting"),
                       ("Every day, 15:00", "The water barometer, read aloud")]:
        p = doc.add_paragraph(style="List Bullet")
        r = p.add_run(when + "  "); r.bold = True
        p.add_run(what)
    para(doc, "Set in Instrument Serif and Poppins, both under the SIL Open Font License, embedded in this file.",
         size=7.5, color=GREY, italic=True, before=10)
    save(doc, "exhibition-leaflet.docx")
    embed_fonts(os.path.join(OUT, "exhibition-leaflet.docx"), {
        SERIF: {"Regular": "InstrumentSerif-Regular.ttf", "Italic": "InstrumentSerif-Italic.ttf"},
        SANS: {"Regular": "Poppins-Regular.ttf", "Bold": "Poppins-Bold.ttf"}})

for make in (annual_results, conference_programme, field_guide, product_spec, invoice, magazine_feature, diagrams, handbook, exhibition_leaflet):
    make()
