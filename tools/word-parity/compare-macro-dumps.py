#!/usr/bin/env python3
"""compare-macro-dumps.py <word.tsv> <ours.tsv>

Compares a macro's result as Word dumped it (run-word-macro.vbs) with ours (run-macro-native.fsx): the
counts, each paragraph's text and formatting, the tables' cells. Booleans are compared as Word's
numbers (True = -1). Prints the differing fields and a summary line; exits 1 when anything differs.
"""
import sys

P_FIELDS = ["kind", "index", "start", "end", "text", "style", "alignment", "spaceBefore", "spaceAfter",
            "leftIndent", "inTable", "font", "size", "bold", "italic", "underline", "color", "highlight",
            "rangeBold", "rangeSize", "rangeColor", "listType", "shading"]
COUNT_FIELDS = ["kind", "paragraphs", "words", "characters", "tables", "contentEnd", "statWords", "statPages"]
TABLE_FIELDS = ["kind", "table", "rows", "columns", "start", "end", "borders"]
CELL_FIELDS = ["kind", "table", "row", "column", "text", "shading", "bold", "color"]


def norm(value):
    v = value.strip()
    if v == "True":
        return "-1"
    if v == "False":
        return "0"
    try:
        f = float(v)
        return repr(round(f, 2))
    except ValueError:
        return v


def load(path):
    with open(path, encoding="utf-8-sig") as f:
        return [line.rstrip("\r\n").split("\t") for line in f if line.strip() and not line.startswith("log\t")]


def fields_for(kind):
    return {"p": P_FIELDS, "counts": COUNT_FIELDS, "table": TABLE_FIELDS, "cell": CELL_FIELDS}.get(kind)


def main():
    word, ours = load(sys.argv[1]), load(sys.argv[2])
    differences = 0
    for i in range(max(len(word), len(ours))):
        w = word[i] if i < len(word) else []
        o = ours[i] if i < len(ours) else []
        if not w or not o or w[0] != o[0]:
            print(f"line {i + 1}: word {w[:5]} / ours {o[:5]}")
            differences += 1
            continue
        names = fields_for(w[0]) or [f"f{k}" for k in range(max(len(w), len(o)))]
        in_table = w[0] == "p" and len(w) > 10 and norm(w[10]) == "-1"
        for k in range(1, max(len(w), len(o))):
            # (Word's Paragraph.Shading in a table cell is not the cell's: an empty cell says -1, an
            # end-of-row mark the colour of a cell of its row. Not compared.)
            if in_table and k < len(names) and names[k] == "shading":
                continue
            # (A theme colour is a large negative number in Word, 0xD...; ours are resolved RGB.)
            if k < len(names) and names[k] in ("color", "rangeColor", "shading") and k < len(w) and w[k].lstrip("-").isdigit() and int(w[k]) < -16777216:
                continue
            a = norm(w[k]) if k < len(w) else ""
            b = norm(o[k]) if k < len(o) else ""
            if a != b:
                name = names[k] if k < len(names) else f"f{k}"
                print(f"line {i + 1} {w[0]} {w[1] if len(w) > 1 else ''} {name}: word {w[k] if k < len(w) else ''!r} ours {o[k] if k < len(o) else ''!r}")
                differences += 1
    print(f"{differences} difference(s) in {len(word)} Word line(s)")
    sys.exit(1 if differences else 0)


if __name__ == "__main__":
    main()
