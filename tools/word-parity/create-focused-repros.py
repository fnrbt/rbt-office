#!/usr/bin/env python3
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path


WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def esc(text):
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def paragraph(text="", style=None, spacing_after=None, list_id=None, list_level=None):
    props = []
    if style:
        props.append(f'<w:pStyle w:val="{esc(style)}"/>')
    if spacing_after is not None:
        props.append(f'<w:spacing w:after="{spacing_after}"/>')
    if list_id is not None:
        level = 0 if list_level is None else list_level
        props.append(
            "<w:numPr>"
            f'<w:ilvl w:val="{level}"/>'
            f'<w:numId w:val="{list_id}"/>'
            "</w:numPr>"
        )
    ppr = f"<w:pPr>{''.join(props)}</w:pPr>" if props else ""
    return (
        f"<w:p>{ppr}<w:r><w:t xml:space=\"preserve\">{esc(text)}</w:t></w:r></w:p>"
    )


def cell(text, width=2400, spacing_after=None):
    return (
        "<w:tc>"
        f'<w:tcPr><w:tcW w:w="{width}" w:type="dxa"/></w:tcPr>'
        f"{paragraph(text, spacing_after=spacing_after)}"
        "</w:tc>"
    )


def row(values, spacing_after=None):
    return "<w:tr>" + "".join(cell(value, spacing_after=spacing_after) for value in values) + "</w:tr>"


def table(rows, width_pct=5000):
    borders = "".join(
        f'<w:{name} w:val="single" w:sz="4" w:color="000000"/>'
        for name in ["top", "left", "bottom", "right", "insideH", "insideV"]
    )
    return (
        "<w:tbl>"
        f'<w:tblPr><w:tblW w:w="{width_pct}" w:type="pct"/><w:tblBorders>{borders}</w:tblBorders></w:tblPr>'
        + "".join(row_values for row_values in rows)
        + "</w:tbl>"
    )


def section_properties(include_header_footer=False):
    refs = (
        '<w:headerReference w:type="default" r:id="rId3"/>'
        '<w:footerReference w:type="default" r:id="rId4"/>'
        if include_header_footer
        else ""
    )
    return f'<w:sectPr>{refs}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>'


def document(body_xml, leading_section_properties=False, include_header_footer=False):
    sectpr = section_properties(include_header_footer)
    if leading_section_properties:
        body = f"{sectpr}{body_xml}"
    else:
        body = f"{body_xml}{sectpr}"
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}"><w:body>'
        f"{body}"
        "</w:body></w:document>"
    )


def styles():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}">'
        "<w:docDefaults>"
        '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
        "</w:docDefaults>"
        '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="240"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="Heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="240" w:after="160"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/></w:rPr></w:style>'
        "</w:styles>"
    )


def content_types(include_numbering=False, include_header_footer=False):
    numbering_override = (
        '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
        if include_numbering
        else ""
    )
    running_overrides = (
        '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
        '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
        if include_header_footer
        else ""
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Types xmlns="{CONTENT_TYPES_NS}">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        f"{numbering_override}"
        f"{running_overrides}"
        "</Types>"
    )


def root_rels():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_REL_NS}">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        "</Relationships>"
    )


def document_rels(include_numbering=False, include_header_footer=False):
    numbering_rel = (
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
        if include_numbering
        else ""
    )
    running_rels = (
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
        '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
        if include_header_footer
        else ""
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_REL_NS}">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        f"{numbering_rel}"
        f"{running_rels}"
        "</Relationships>"
    )


def numbering():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:numbering xmlns:w="{WORD_NS}">'
        '<w:abstractNum w:abstractNumId="0">'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl>'
        '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/><w:lvlJc w:val="left"/></w:lvl>'
        '<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%3."/><w:lvlJc w:val="left"/></w:lvl>'
        '</w:abstractNum>'
        '<w:abstractNum w:abstractNumId="1">'
        '<w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/></w:lvl>'
        '<w:lvl w:ilvl="1"><w:numFmt w:val="bullet"/><w:lvlText w:val="○"/><w:lvlJc w:val="left"/></w:lvl>'
        '<w:lvl w:ilvl="2"><w:numFmt w:val="bullet"/><w:lvlText w:val="■"/><w:lvlJc w:val="left"/></w:lvl>'
        '</w:abstractNum>'
        '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
        '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>'
        "</w:numbering>"
    )


def header_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:hdr xmlns:w="{WORD_NS}">'
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Focused Header</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>Page 1</w:t></w:r></w:p>'
        '</w:hdr>'
    )


def footer_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:ftr xmlns:w="{WORD_NS}">'
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>Focused Footer</w:t></w:r></w:p>'
        '</w:ftr>'
    )


def write_docx(path, body_xml, include_numbering=False, leading_section_properties=False, include_header_footer=False):
    path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        (root / "_rels").mkdir()
        (root / "word" / "_rels").mkdir(parents=True)
        (root / "[Content_Types].xml").write_text(content_types(include_numbering, include_header_footer), encoding="utf-8")
        (root / "_rels" / ".rels").write_text(root_rels(), encoding="utf-8")
        (root / "word" / "document.xml").write_text(document(body_xml, leading_section_properties, include_header_footer), encoding="utf-8")
        (root / "word" / "styles.xml").write_text(styles(), encoding="utf-8")
        if include_numbering:
            (root / "word" / "numbering.xml").write_text(numbering(), encoding="utf-8")
        if include_header_footer:
            (root / "word" / "header1.xml").write_text(header_xml(), encoding="utf-8")
            (root / "word" / "footer1.xml").write_text(footer_xml(), encoding="utf-8")
        (root / "word" / "_rels" / "document.xml.rels").write_text(document_rels(include_numbering, include_header_footer), encoding="utf-8")
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
            for file in sorted(root.rglob("*")):
                if file.is_file():
                    zf.write(file, file.relative_to(root).as_posix())


def main():
    out_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("tests/repros/word")
    repros = {
        "table-import-preserves-structure.docx": (False, (
            paragraph("Table import repro", "Heading1")
            + table([row(["A1", "B1"]), row(["A2", "B2"])])
            + paragraph("After table")
        )),
        "table-cell-default-after-spacing.docx": (False, (
            paragraph("Table cell spacing repro", "Heading1")
            + table([row(["Alpha", "Beta"]), row(["Gamma", "Delta"])])
            + paragraph("After table")
        )),
        "two-tables-border-position.docx": (False, (
            paragraph("Two table border position repro", "Heading1")
            + paragraph("First table", "Heading2")
            + table([row(["A1", "B1", "C1"]), row(["A2", "B2", "C2"])])
            + paragraph("\n")
            + paragraph("Second table", "Heading2")
            + table([row(["Product", "Price", "Quantity", "Total"]), row(["Widget A", "$10.00", "5", "$50.00"]), row(["Total", "", "", "$50.00"])])
            + paragraph("After second table")
        )),
        "normal-paragraph-line-pitch.docx": (False, (
            paragraph("Normal paragraph pitch repro", "Heading1")
            + paragraph("First normal line")
            + paragraph("Second normal line")
            + paragraph("Third normal line")
            + paragraph("Wrapped line pitch - Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.")
        )),
        "list-default-indent.docx": (True, False, (
            paragraph("List default indent repro", "Heading1")
            + paragraph("Top bullet", list_id=2, list_level=0)
            + paragraph("Nested bullet", list_id=2, list_level=1)
            + paragraph("Deep bullet", list_id=2, list_level=2)
            + paragraph("Top number", list_id=1, list_level=0)
            + paragraph("Nested number", list_id=1, list_level=1)
            + paragraph("Deep number", list_id=1, list_level=2)
            + paragraph("After list")
        )),
        "running-header-body-offset.docx": (False, True, True, (
            paragraph("Running header body offset repro", "Heading1")
            + paragraph("First body paragraph")
            + paragraph("Second body paragraph")
        )),
    }
    for name, value in repros.items():
        if len(value) == 2:
            include_numbering, body = value
            leading_section_properties = False
            include_header_footer = False
        elif len(value) == 3:
            include_numbering, leading_section_properties, body = value
            include_header_footer = False
        else:
            include_numbering, leading_section_properties, include_header_footer, body = value
        path = out_dir / name
        write_docx(path, body, include_numbering, leading_section_properties, include_header_footer)
        print(path)


if __name__ == "__main__":
    main()
