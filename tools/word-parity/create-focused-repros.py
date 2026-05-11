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


def paragraph(
    text="",
    style=None,
    alignment=None,
    spacing_before=None,
    spacing_after=None,
    list_id=None,
    list_level=None,
    indent_left=None,
    hanging=None,
    first_line=None,
):
    props = []
    if style:
        props.append(f'<w:pStyle w:val="{esc(style)}"/>')
    if alignment:
        props.append(f'<w:jc w:val="{esc(alignment)}"/>')
    spacing_attrs = []
    if spacing_before is not None:
        spacing_attrs.append(f'w:before="{spacing_before}"')
    if spacing_after is not None:
        spacing_attrs.append(f'w:after="{spacing_after}"')
    if spacing_attrs:
        props.append(f"<w:spacing {' '.join(spacing_attrs)}/>")
    if list_id is not None:
        level = 0 if list_level is None else list_level
        props.append(
            "<w:numPr>"
            f'<w:ilvl w:val="{level}"/>'
            f'<w:numId w:val="{list_id}"/>'
            "</w:numPr>"
        )
    indent_attrs = []
    if indent_left is not None:
        indent_attrs.append(f'w:left="{indent_left}"')
    if hanging is not None:
        indent_attrs.append(f'w:hanging="{hanging}"')
    if first_line is not None:
        indent_attrs.append(f'w:firstLine="{first_line}"')
    if indent_attrs:
        props.append(f"<w:ind {' '.join(indent_attrs)}/>")
    ppr = f"<w:pPr>{''.join(props)}</w:pPr>" if props else ""
    return (
        f"<w:p>{ppr}<w:r><w:t xml:space=\"preserve\">{esc(text)}</w:t></w:r></w:p>"
    )


def empty_paragraph():
    return "<w:p/>"


def cell(text, width=2400, spacing_after=None, padding_left=None, alignment=None, valign=None, shading=None):
    tc_props = [f'<w:tcW w:w="{width}" w:type="dxa"/>']
    if shading:
        tc_props.append(f'<w:shd w:val="{esc(shading)}" w:color="auto" w:fill="auto"/>')
    if valign:
        tc_props.append(f'<w:vAlign w:val="{esc(valign)}"/>')
    if padding_left is not None:
        tc_props.append(f'<w:tcMar><w:left w:w="{padding_left}" w:type="dxa"/></w:tcMar>')
    return (
        "<w:tc>"
        f"<w:tcPr>{''.join(tc_props)}</w:tcPr>"
        f"{paragraph(text, alignment=alignment, spacing_after=spacing_after)}"
        "</w:tc>"
    )


def cell_paragraphs(paragraphs, width=2400, padding_left=None):
    tc_props = [f'<w:tcW w:w="{width}" w:type="dxa"/>']
    if padding_left is not None:
        tc_props.append(f'<w:tcMar><w:left w:w="{padding_left}" w:type="dxa"/></w:tcMar>')
    return (
        "<w:tc>"
        f"<w:tcPr>{''.join(tc_props)}</w:tcPr>"
        f"{''.join(paragraphs)}"
        "</w:tc>"
    )


def row(values, spacing_after=None, height=None):
    tr_props = f'<w:trPr><w:trHeight w:val="{height}"/></w:trPr>' if height is not None else ""
    return "<w:tr>" + tr_props + "".join(cell(value, spacing_after=spacing_after) for value in values) + "</w:tr>"


def row_cells(cells):
    return "<w:tr>" + "".join(cells) + "</w:tr>"


def row_cells_props(cells, height=None):
    tr_props = f'<w:trPr><w:trHeight w:val="{height}"/></w:trPr>' if height is not None else ""
    return "<w:tr>" + tr_props + "".join(cells) + "</w:tr>"


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


def table_with_grid(rows, grid_widths, table_width="0", width_type="auto"):
    borders = "".join(
        f'<w:{name} w:val="single" w:sz="4" w:color="000000"/>'
        for name in ["top", "left", "bottom", "right", "insideH", "insideV"]
    )
    grid = "".join(f'<w:gridCol w:w="{grid_width}"/>' for grid_width in grid_widths)
    return (
        "<w:tbl>"
        f'<w:tblPr><w:tblW w:w="{table_width}" w:type="{width_type}"/><w:tblBorders>{borders}</w:tblBorders></w:tblPr>'
        f"<w:tblGrid>{grid}</w:tblGrid>"
        + "".join(row_values for row_values in rows)
        + "</w:tbl>"
    )


def table_grid_style_table(rows, grid_widths, table_width="0", width_type="auto"):
    grid = "".join(f'<w:gridCol w:w="{grid_width}"/>' for grid_width in grid_widths)
    return (
        "<w:tbl>"
        f'<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="{table_width}" w:type="{width_type}"/>'
        '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
        "</w:tblPr>"
        f"<w:tblGrid>{grid}</w:tblGrid>"
        + "".join(row_values for row_values in rows)
        + "</w:tbl>"
    )


def repeated_long_paragraphs(count):
    text = (
        "Lorem ipsum dolor sit amet, consectetur adipiscing elit. "
        "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. "
        "Lorem ipsum dolor sit amet, consectetur adipiscing elit. "
        "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua."
    )
    return "".join(paragraph(text) for _ in range(count))


def long_paragraph(spacing_before=None):
    text = (
        "Lorem ipsum dolor sit amet, consectetur adipiscing elit. "
        "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. "
        "Lorem ipsum dolor sit amet, consectetur adipiscing elit. "
        "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua."
    )
    return paragraph(text, spacing_before=spacing_before)


def section_properties(include_header_footer=False, margin_bottom=1440, include_doc_grid=False):
    refs = (
        '<w:headerReference w:type="default" r:id="rId3"/>'
        '<w:footerReference w:type="default" r:id="rId4"/>'
        if include_header_footer
        else ""
    )
    doc_grid = '<w:docGrid w:type="default" w:linePitch="360" w:charSpace="4294965247"/>' if include_doc_grid else ""
    return f'<w:sectPr>{refs}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="{margin_bottom}" w:left="1440" w:header="720" w:footer="720"/>{doc_grid}</w:sectPr>'


def document(body_xml, leading_section_properties=False, include_header_footer=False, margin_bottom=1440, include_doc_grid=False):
    sectpr = section_properties(include_header_footer, margin_bottom, include_doc_grid)
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


def styles(include_heading_spacing=True, include_doc_defaults=True):
    heading1_spacing = '<w:spacing w:after="240"/>' if include_heading_spacing else ""
    heading2_spacing = '<w:spacing w:before="240" w:after="160"/>' if include_heading_spacing else ""
    doc_defaults = (
        "<w:docDefaults>"
        '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
        "</w:docDefaults>"
        if include_doc_defaults
        else ""
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}">'
        f"{doc_defaults}"
        '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/></w:style>'
        f'<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:basedOn w:val="Normal"/><w:pPr>{heading1_spacing}<w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>'
        f'<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="Heading 2"/><w:basedOn w:val="Normal"/><w:pPr>{heading2_spacing}<w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/></w:rPr></w:style>'
        "</w:styles>"
    )


def table_grid_styles():
    borders = "".join(
        f'<w:{name} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        for name in ["top", "left", "bottom", "right", "insideH", "insideV"]
    )
    cell_margin = (
        '<w:tblCellMar>'
        '<w:top w:w="0" w:type="dxa"/>'
        '<w:left w:w="0" w:type="dxa"/>'
        '<w:bottom w:w="0" w:type="dxa"/>'
        '<w:right w:w="0" w:type="dxa"/>'
        '</w:tblCellMar>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}">'
        "<w:docDefaults>"
        '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
        "</w:docDefaults>"
        '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="240"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>'
        f'<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr>{cell_margin}</w:tblPr></w:style>'
        f'<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblBorders>{borders}</w:tblBorders>{cell_margin}</w:tblPr></w:style>'
        "</w:styles>"
    )


def table_grid_paragraph_spacing_styles():
    borders = "".join(
        f'<w:{name} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        for name in ["top", "left", "bottom", "right", "insideH", "insideV"]
    )
    cell_margin = (
        '<w:tblCellMar>'
        '<w:top w:w="0" w:type="dxa"/>'
        '<w:left w:w="108" w:type="dxa"/>'
        '<w:bottom w:w="0" w:type="dxa"/>'
        '<w:right w:w="108" w:type="dxa"/>'
        '</w:tblCellMar>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}">'
        "<w:docDefaults>"
        '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
        "</w:docDefaults>"
        '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/><w:rPr><w:sz w:val="24"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="240"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>'
        f'<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr>{cell_margin}</w:tblPr></w:style>'
        f'<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblBorders>{borders}</w:tblBorders>{cell_margin}</w:tblPr></w:style>'
        "</w:styles>"
    )


def top_heading_before_spacing_styles():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}">'
        "<w:docDefaults>"
        '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
        "</w:docDefaults>"
        '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading1">'
        '<w:name w:val="Heading 1"/><w:basedOn w:val="Normal"/>'
        '<w:pPr><w:spacing w:before="480" w:after="0"/><w:outlineLvl w:val="0"/></w:pPr>'
        '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>'
        "</w:style>"
        "</w:styles>"
    )


def cambria_heading_wrap_styles():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}">'
        "<w:docDefaults>"
        '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
        "</w:docDefaults>"
        '<w:style w:type="paragraph" w:styleId="Normal" w:default="1">'
        '<w:name w:val="Normal"/>'
        '<w:pPr><w:widowControl/><w:spacing w:lineRule="auto" w:line="276" w:before="100" w:after="100"/><w:ind w:left="567" w:hanging="0"/><w:jc w:val="left"/></w:pPr>'
        '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>'
        "</w:style>"
        '<w:style w:type="paragraph" w:styleId="Heading1">'
        '<w:name w:val="Heading 1"/><w:basedOn w:val="Normal"/>'
        '<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="480" w:after="0"/><w:ind w:left="567" w:hanging="567"/><w:outlineLvl w:val="0"/></w:pPr>'
        '<w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria" w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/><w:b/><w:color w:val="365F91"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>'
        "</w:style>"
        "</w:styles>"
    )


def content_types(include_numbering=False, include_header_footer=False, include_office_theme=False):
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
    theme_overrides = (
        '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
        '<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>'
        '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
        if include_office_theme
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
        f"{theme_overrides}"
        "</Types>"
    )


def root_rels():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_REL_NS}">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        "</Relationships>"
    )


def document_rels(include_numbering=False, include_header_footer=False, include_office_theme=False):
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
    theme_rels = (
        '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>'
        '<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
        '<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
        if include_office_theme
        else ""
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_REL_NS}">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        f"{numbering_rel}"
        f"{running_rels}"
        f"{theme_rels}"
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


def word_settings_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:settings xmlns:w="{WORD_NS}">'
        '<w:zoom w:percent="100"/>'
        '<w:defaultTabStop w:val="720"/>'
        '<w:compat>'
        '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/>'
        '<w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
        '<w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
        '<w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
        '</w:compat>'
        '<w:themeFontLang w:val="en-US" w:eastAsia="en-US" w:bidi=""/>'
        '</w:settings>'
    )


def font_table_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:fonts xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}">'
        '<w:font w:name="Calibri"><w:charset w:val="01"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font>'
        '<w:font w:name="Cambria"><w:charset w:val="01"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font>'
        '<w:font w:name="Arial"><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/></w:font>'
        '<w:font w:name="Times New Roman"><w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font>'
        "</w:fonts>"
    )


def theme_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">'
        '<a:themeElements>'
        '<a:clrScheme name="Office">'
        '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
        '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
        '<a:dk2><a:srgbClr val="1F497D"/></a:dk2>'
        '<a:lt2><a:srgbClr val="EEECE1"/></a:lt2>'
        '<a:accent1><a:srgbClr val="4F81BD"/></a:accent1>'
        '<a:accent2><a:srgbClr val="C0504D"/></a:accent2>'
        '<a:accent3><a:srgbClr val="9BBB59"/></a:accent3>'
        '<a:accent4><a:srgbClr val="8064A2"/></a:accent4>'
        '<a:accent5><a:srgbClr val="4BACC6"/></a:accent5>'
        '<a:accent6><a:srgbClr val="F79646"/></a:accent6>'
        '<a:hlink><a:srgbClr val="0000FF"/></a:hlink>'
        '<a:folHlink><a:srgbClr val="800080"/></a:folHlink>'
        '</a:clrScheme>'
        '<a:fontScheme name="Office">'
        '<a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
        '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>'
        '</a:fontScheme>'
        '<a:fmtScheme name="Office">'
        '<a:fillStyleLst>'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"/></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"/></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill>'
        '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"/></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"/></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill>'
        '</a:fillStyleLst>'
        '<a:lnStyleLst>'
        '<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>'
        '<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>'
        '<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>'
        '</a:lnStyleLst>'
        '<a:effectStyleLst>'
        '<a:effectStyle><a:effectLst/></a:effectStyle>'
        '<a:effectStyle><a:effectLst/></a:effectStyle>'
        '<a:effectStyle><a:effectLst/></a:effectStyle>'
        '</a:effectStyleLst>'
        '<a:bgFillStyleLst>'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '</a:bgFillStyleLst>'
        '</a:fmtScheme>'
        '</a:themeElements>'
        '<a:objectDefaults/>'
        '<a:extraClrSchemeLst/>'
        '</a:theme>'
    )


def write_docx(path, body_xml, include_numbering=False, leading_section_properties=False, include_header_footer=False, margin_bottom=1440, include_heading_spacing=True, include_doc_defaults=True, styles_xml=None, include_office_theme=False, include_doc_grid=False):
    path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        (root / "_rels").mkdir()
        (root / "word" / "_rels").mkdir(parents=True)
        (root / "[Content_Types].xml").write_text(content_types(include_numbering, include_header_footer, include_office_theme), encoding="utf-8")
        (root / "_rels" / ".rels").write_text(root_rels(), encoding="utf-8")
        (root / "word" / "document.xml").write_text(document(body_xml, leading_section_properties, include_header_footer, margin_bottom, include_doc_grid), encoding="utf-8")
        (root / "word" / "styles.xml").write_text(styles_xml or styles(include_heading_spacing, include_doc_defaults), encoding="utf-8")
        if include_numbering:
            (root / "word" / "numbering.xml").write_text(numbering(), encoding="utf-8")
        if include_header_footer:
            (root / "word" / "header1.xml").write_text(header_xml(), encoding="utf-8")
            (root / "word" / "footer1.xml").write_text(footer_xml(), encoding="utf-8")
        if include_office_theme:
            (root / "word" / "theme").mkdir()
            (root / "word" / "settings.xml").write_text(word_settings_xml(), encoding="utf-8")
            (root / "word" / "fontTable.xml").write_text(font_table_xml(), encoding="utf-8")
            (root / "word" / "theme" / "theme1.xml").write_text(theme_xml(), encoding="utf-8")
        (root / "word" / "_rels" / "document.xml.rels").write_text(document_rels(include_numbering, include_header_footer, include_office_theme), encoding="utf-8")
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
            for file in sorted(root.rglob("*")):
                if file.is_file():
                    info = zipfile.ZipInfo(file.relative_to(root).as_posix(), (1980, 1, 1, 0, 0, 0))
                    info.compress_type = zipfile.ZIP_DEFLATED
                    zf.writestr(info, file.read_bytes())


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
        "table-cell-explicit-left-padding.docx": (False, (
            paragraph("Table cell explicit left padding repro", "Heading1")
            + table([
                row_cells([
                    cell("Name", padding_left=108),
                    cell("Value", padding_left=108),
                ]),
                row_cells([
                    cell("Country", padding_left=108),
                    cell("AU", padding_left=108),
                ]),
            ])
            + paragraph("After padded table")
        )),
        "table-dxa-grid-alignment-row-height.docx": (False, (
            paragraph("Table DXA grid alignment row height repro", "Heading1")
            + table_with_grid([
                row_cells_props([
                    cell("A", width=4606),
                    cell("B", width=4606, alignment="right"),
                ], height=1134),
                row_cells_props([
                    cell("C", width=4606),
                    cell("D", width=4606, alignment="right"),
                ], height=1134),
            ], [4606, 4606])
            + paragraph("After table")
        )),
        "table-grid-style-borders.docx": (False, False, False, (
            paragraph("TableGrid style border repro", "Heading1")
            + table_grid_style_table([
                row_cells([
                    cell("A", width=3000),
                    cell("B", width=3000),
                ]),
                row_cells([
                    cell("C", width=3000),
                    cell("D", width=3000),
                ]),
            ], [3000, 3000])
            + paragraph("After styled table")
        ), True, True, table_grid_styles()),
        "table-style-paragraph-spacing.docx": (False, False, False, (
            paragraph("Table style paragraph spacing repro", "Heading1")
            + table_grid_style_table([
                row_cells([cell("row 01 a", width=2303), cell("row 01 b", width=2303), cell("row 01 c", width=2303), cell("row 01 d", width=2303)]),
                row_cells([cell("row 02 a", width=2303), cell("row 02 b", width=2303), cell("row 02 c", width=2303), cell("row 02 d", width=2303)]),
                row_cells([cell("row 03 a", width=2303), cell("row 03 b", width=2303), cell("row 03 c", width=2303), cell("row 03 d", width=2303)]),
                row_cells([cell("row 04 a", width=2303), cell("row 04 b", width=2303), cell("row 04 c", width=2303), cell("row 04 d", width=2303)]),
            ], [2303, 2303, 2303, 2303])
            + paragraph("After table")
        ), True, True, table_grid_paragraph_spacing_styles()),
        "table-cell-explicit-before-spacing.docx": (False, (
            paragraph("Table cell before-spacing repro", "Heading1")
            + table([
                row_cells([
                    cell_paragraphs([paragraph("First", spacing_before=100, spacing_after=0)]),
                    cell_paragraphs([paragraph("", spacing_before=100, spacing_after=0)]),
                ]),
                row_cells([
                    cell_paragraphs([paragraph("Second", spacing_before=100, spacing_after=0)]),
                    cell_paragraphs([
                        paragraph("", spacing_before=100, spacing_after=0),
                        paragraph("", spacing_before=100, spacing_after=0),
                        paragraph("", spacing_before=100, spacing_after=0),
                        paragraph("", spacing_before=100, spacing_after=0),
                    ]),
                ]),
                row_cells([
                    cell_paragraphs([paragraph("Third", spacing_before=100, spacing_after=0)]),
                    cell_paragraphs([paragraph("", spacing_before=100, spacing_after=0)]),
                ]),
            ])
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
        "empty-normal-paragraph-mark-size.docx": (False, False, False, (
            paragraph("Empty paragraph mark size repro")
            + empty_paragraph()
            + paragraph("Second line after an empty Normal paragraph")
            + empty_paragraph()
            + paragraph("Third line after another empty Normal paragraph")
        ), True, False),
        "heading-style-no-spacing.docx": (False, False, False, (
            paragraph("Heading style no spacing repro", "Heading1")
            + paragraph("")
            + paragraph("Section A", "Heading2")
            + paragraph("First normal line")
            + paragraph("Second normal line")
            + paragraph("")
            + paragraph("Section B", "Heading2")
            + paragraph("Third normal line")
        ), False),
        "top-heading-before-spacing.docx": (False, False, False, (
            paragraph("Top heading before-spacing repro", "Heading1")
            + paragraph("Second heading with the same imported before-spacing", "Heading1")
            + paragraph("After heading")
        ), True, True, top_heading_before_spacing_styles()),
        "heading-cambria-font-wrap.docx": (True, False, False, (
            paragraph("Heading Cambria font wrap repro", spacing_before=0, spacing_after=0, indent_left=0, hanging=0)
            + paragraph("Subject to the terms and conditions of this Agreement, You hereby grant to the Sponsor and to recipients of software distributed by the Sponsor a perpetual, worldwide, non-exclusive, no-charge, royalty-free, irrevocable (except as stated in this section) patent license to make, have made, use, offer to sell, sell, import, and otherwise transfer the Work, where such license applies only to those patent claims licensable by You that are necessarily infringed by Your Contribution(s) alone or by combination of Your Contribution(s) with the Work to which such Contribution(s) was submitted. If any entity institutes patent litigation against You or any other entity (including a cross-claim or counterclaim in a lawsuit) alleging that your Contribution, or the Work to which you have contributed, constitutes direct or contributory patent infringement, then any patent licenses granted to that entity under this Agreement for that Contribution or Work shall terminate as of the date such litigation is filed. ")
            + paragraph("You represent that you are legally entitled to grant the above license. ", "Heading1", list_id=1, list_level=0, indent_left=567, hanging=567)
            + paragraph("If your employer(s) has rights to intellectual property that you create that includes your Contributions, you represent that you have received permission to make Contributions on behalf of that employer, that your employer has waived such rights for your Contributions to the Sponsor, or that your employer has executed a separate Corporate CLA with the Sponsor. ")
        ), True, True, cambria_heading_wrap_styles(), True, True),
        "wrapped-paragraph-line-pitch.docx": (False, (
            paragraph("Wrapped paragraph line pitch repro", "Heading1")
            + repeated_long_paragraphs(18)
            + paragraph("After wrapped paragraphs")
        )),
        "widow-control-last-line.docx": (False, (
            paragraph("Widow control last line repro", "Heading1")
            + repeated_long_paragraphs(12)
            + long_paragraph()
            + paragraph("After widow probe")
        ), 1740),
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
        margin_bottom = 1440
        include_heading_spacing = True
        include_doc_defaults = True
        include_office_theme = False
        include_doc_grid = False
        if len(value) == 2:
            include_numbering, body = value
            leading_section_properties = False
            include_header_footer = False
            styles_xml = None
        elif len(value) == 3:
            styles_xml = None
            if isinstance(value[2], int):
                include_numbering, body, margin_bottom = value
                leading_section_properties = False
                include_header_footer = False
            else:
                include_numbering, leading_section_properties, body = value
                include_header_footer = False
        else:
            include_numbering, leading_section_properties, include_header_footer, body = value[:4]
            styles_xml = None
            if len(value) > 4:
                include_heading_spacing = value[4]
            if len(value) > 5:
                include_doc_defaults = value[5]
            if len(value) > 6:
                styles_xml = value[6]
            if len(value) > 7:
                include_office_theme = value[7]
            if len(value) > 8:
                include_doc_grid = value[8]
        path = out_dir / name
        write_docx(path, body, include_numbering, leading_section_properties, include_header_footer, margin_bottom, include_heading_spacing, include_doc_defaults, styles_xml, include_office_theme, include_doc_grid)
        print(path)


if __name__ == "__main__":
    main()
