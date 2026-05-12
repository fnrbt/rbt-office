#!/usr/bin/env python3
import subprocess
import sys
import xml.etree.ElementTree as ET
from pathlib import Path


def local_name(tag):
    return tag.rsplit("}", 1)[-1]


def attr_float(node, name):
    try:
        return float(node.attrib.get(name, "nan"))
    except ValueError:
        return float("nan")


def clean_cell(value):
    return str(value).replace("\t", " ").replace("\r", " ").replace("\n", " ")


def write_tsv(path, rows):
    headers = [
        "doc",
        "page",
        "paragraph",
        "line",
        "start",
        "end",
        "x",
        "y",
        "lastX",
        "lastY",
        "adjustedPage",
        "section",
        "pageCount",
        "textBoundaryX",
        "textBoundaryY",
        "firstCharColumn",
        "startRow",
        "endRow",
        "maxRows",
        "startColumn",
        "endColumn",
        "maxColumns",
        "atEndOfRowMarker",
        "inContentControl",
        "inFieldCode",
        "inFieldResult",
        "rangeLines",
        "rangePages",
        "inTable",
        "table",
        "row",
        "column",
        "style",
        "text",
    ]
    with open(path, "w", encoding="utf-8", newline="\n") as out:
        out.write("\t".join(headers) + "\n")
        for row in rows:
            out.write("\t".join(clean_cell(row.get(header, "")) for header in headers) + "\n")


def extract_rows(pdf_path):
    xml = subprocess.check_output(
        ["pdftotext", "-bbox-layout", str(pdf_path), "-"],
        text=True,
        stderr=subprocess.PIPE,
    )
    root = ET.fromstring(xml)
    rows = []
    pages = [node for node in root.iter() if local_name(node.tag) == "page"]
    for page_index, page in enumerate(pages, start=1):
        for line in page.iter():
            if local_name(line.tag) != "line":
                continue
            words = [word for word in line if local_name(word.tag) == "word"]
            text = " ".join("".join(word.itertext()) for word in words).strip()
            if not text:
                continue
            rows.append(
                {
                    "doc": str(pdf_path),
                    "page": page_index,
                    "paragraph": 0,
                    "line": 0,
                    "start": 0,
                    "end": len(text),
                    "x": f"{attr_float(line, 'xMin'):.6f}",
                    "y": f"{attr_float(line, 'yMin'):.6f}",
                    "lastX": f"{attr_float(line, 'xMax'):.6f}",
                    "lastY": f"{attr_float(line, 'yMax'):.6f}",
                    "inTable": "",
                    "table": "",
                    "row": "",
                    "column": "",
                    "style": "pdf",
                    "text": text,
                }
            )
    rows.sort(key=lambda row: (int(row["page"]), round(float(row["y"]), 3), round(float(row["x"]), 3)))
    page_line_counts = {}
    for index, row in enumerate(rows, start=1):
        page = int(row["page"])
        page_line_counts[page] = page_line_counts.get(page, 0) + 1
        row["paragraph"] = index
        row["line"] = page_line_counts[page]
    return rows


def main():
    if len(sys.argv) < 3:
        print("usage: extract-pdf-text-layout.py INPUT.pdf OUTPUT.tsv", file=sys.stderr)
        return 2
    pdf_path = Path(sys.argv[1]).resolve()
    output_path = Path(sys.argv[2]).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    write_tsv(output_path, extract_rows(pdf_path))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
