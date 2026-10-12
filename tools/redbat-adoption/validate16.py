# Structural validation of OOXML packages. usage: python3 -I validate.py file... ; prints one line per file, issues below it.
import sys, zipfile, posixpath, re, collections
import xml.etree.ElementTree as ET
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
def check(path):
    issues, facts = [], {}
    z = zipfile.ZipFile(path)
    bad = z.testzip()
    if bad: issues.append(f"zip: bad CRC in {bad}")
    names = [n for n in z.namelist() if not n.endswith("/")]
    dup = [n for n, c in collections.Counter(names).items() if c > 1]
    if dup: issues.append(f"zip: duplicate entries {dup}")
    nameset = set(names); trees = {}
    for n in names:
        if re.search(r"\.(xml|rels|vml)$", n, re.I) or n == "[Content_Types].xml":
            try: trees[n] = ET.fromstring(z.read(n))
            except ET.ParseError as e:
                if n.endswith(".vml"): facts.setdefault("vml not well-formed XML", []).append(n)
                else: issues.append(f"not well-formed: {n}: {e}")
    # content types
    ct = trees.get("[Content_Types].xml")
    if ct is None: issues.append("no [Content_Types].xml"); return issues, facts
    defaults = {e.get("Extension").lower(): e.get("ContentType") for e in ct if e.tag.endswith("Default")}
    overrides = {e.get("PartName"): e.get("ContentType") for e in ct if e.tag.endswith("Override")}
    ovl = [e.get("PartName") for e in ct if e.tag.endswith("Override")]
    for p, c in collections.Counter(ovl).items():
        if c > 1: issues.append(f"content types: {p} overridden {c} times")
    for p in overrides:
        if p.lstrip("/") not in nameset: issues.append(f"content types: Override for a part that is not there: {p}")
    for n in names:
        if n == "[Content_Types].xml": continue
        ext = n.rsplit(".", 1)[-1].lower() if "." in n else ""
        if "/" + n not in overrides and ext not in defaults: issues.append(f"no content type for {n}")
    # relationships
    relmap = {}
    for n in names:
        if not n.endswith(".rels") or n not in trees: continue
        d = posixpath.dirname(posixpath.dirname(n)); src = posixpath.join(d, posixpath.basename(n)[:-5]) if posixpath.basename(n) != ".rels" else ""
        if src and src not in nameset: issues.append(f"{n}: relationships of a part that is not there ({src})")
        ids = [r.get("Id") for r in trees[n]]
        for i, c in collections.Counter(ids).items():
            if c > 1: issues.append(f"{n}: relationship id {i} {c} times")
        m = {}
        for r in trees[n]:
            t = r.get("Target"); m[r.get("Id")] = (r.get("Type"), t, r.get("TargetMode"))
            if r.get("TargetMode") == "External": continue
            target = t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join(d, t))
            target = target.split("#")[0]
            if target not in nameset: issues.append(f"{n}: {r.get('Id')} -> {t}: no such part")
        relmap[src] = m
    # r:id / r:embed ... attributes that name a relationship the part does not have
    for n, tree in trees.items():
        if n.endswith(".rels") or n == "[Content_Types].xml": continue
        used = set()
        for e in tree.iter():
            for k, v in e.attrib.items():
                if k.startswith(R): used.add(v)
        missing = [u for u in used if u and u not in relmap.get(n, {})]
        if missing: issues.append(f"{n}: refers to relationships it does not have: {sorted(missing)[:6]}")
    # xlsx
    if "xl/workbook.xml" in trees:
        ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        wb = trees["xl/workbook.xml"]; sheets = wb.findall("m:sheets/m:sheet", ns)
        for attr, label in ((R + "id", "r:id"), ("sheetId", "sheetId"), ("name", "name")):
            vals = [s.get(attr) for s in sheets]
            d = [v for v, c in collections.Counter(vals).items() if c > 1]
            if d: issues.append(f"workbook: duplicate sheet {label} {d}")
        st = trees.get("xl/styles.xml"); nxf = None
        if st is not None:
            for coll, child in (("fonts", "font"), ("fills", "fill"), ("borders", "border"), ("cellStyleXfs", "xf"), ("cellXfs", "xf"), ("cellStyles", "cellStyle"), ("dxfs", "dxf"), ("numFmts", "numFmt")):
                c = st.find("m:" + coll, ns)
                if c is None: continue
                real = len(c.findall("m:" + child, ns)); said = c.get("count")
                if coll == "cellXfs": nxf = real
                if said is not None and int(said) != real: issues.append(f"styles: <{coll} count=\"{said}\"> but {real} <{child}>")
            cx = st.find("m:cellXfs", ns)
            if cx is not None:
                counts = {k: len(st.find("m:" + k, ns) or []) for k in ("fonts", "fills", "borders")}
                for i, xf in enumerate(cx):
                    for a, k in (("fontId", "fonts"), ("fillId", "fills"), ("borderId", "borders")):
                        if xf.get(a) is not None and int(xf.get(a)) >= counts[k]: issues.append(f"styles: cellXfs[{i}] {a}={xf.get(a)} of {counts[k]}"); break
                sig = collections.Counter(ET.tostring(x) for x in cx); facts["cellXfs"] = f"{len(cx)} ({len(sig)} distinct)"
        ss = trees.get("xl/sharedStrings.xml"); nsi = None; tcount = 0
        if ss is not None:
            nsi = len(ss.findall("m:si", ns))
            if ss.get("uniqueCount") is not None and int(ss.get("uniqueCount")) != nsi: issues.append(f"sharedStrings: uniqueCount=\"{ss.get('uniqueCount')}\" but {nsi} <si>")
            texts = collections.Counter(ET.tostring(si) for si in ss.findall("m:si", ns)); facts["shared strings"] = f"{nsi} ({nsi - len(texts)} duplicates)"
        for n, tree in trees.items():
            if not re.match(r"xl/worksheets/sheet\d+\.xml$", n): continue
            refs = set(); lastrow = 0
            for row in tree.findall("m:sheetData/m:row", ns):
                r = int(row.get("r") or 0)
                if r and r <= lastrow: issues.append(f"{n}: row {r} after row {lastrow}"); break
                lastrow = r or lastrow
                for c in row.findall("m:c", ns):
                    ref = c.get("r")
                    if ref in refs: issues.append(f"{n}: cell {ref} twice"); break
                    refs.add(ref)
                    if c.get("s") is not None and nxf is not None and int(c.get("s")) >= nxf: issues.append(f"{n}: cell {ref} s={c.get('s')} of {nxf} cellXfs"); break
                    if c.get("t") == "s":
                        tcount += 1; v = c.find("m:v", ns)
                        if v is None or nsi is None or int(v.text) >= nsi: issues.append(f"{n}: cell {ref} names shared string {v.text if v is not None else None} of {nsi}"); break
                if row.get("s") is not None and nxf is not None and int(row.get("s")) >= nxf: issues.append(f"{n}: row {r} s={row.get('s')} of {nxf}")
        if ss is not None and ss.get("count") is not None: facts["sst count"] = f"says {ss.get('count')}, cells use {tcount}"
        if ss is not None and ss.get("count") is not None and int(ss.get("count")) != tcount: facts["sst count mismatch"] = True
        cms = [n for n in names if re.match(r"xl/comments\d+\.xml$", n)]
        for n in cms:
            t = trees[n]; na = len(t.findall("m:authors/m:author", ns))
            for c in t.findall("m:commentList/m:comment", ns):
                if int(c.get("authorId")) >= na: issues.append(f"{n}: comment {c.get('ref')} authorId {c.get('authorId')} of {na}")
    # docx
    if "word/document.xml" in trees:
        W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"; W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"; WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
        doc = trees["word/document.xml"]
        def dups(vals): return [v for v, c in collections.Counter(vals).items() if c > 1 and v is not None]
        facts["dup paraId"] = len(dups([p.get(W14 + "paraId") for n, t in trees.items() if n.startswith("word/") for p in t.iter(W + "p")]))
        facts["dup bookmark id"] = len(dups([b.get(W + "id") for b in doc.iter(W + "bookmarkStart")]))
        facts["dup bookmark name"] = len(dups([b.get(W + "name") for b in doc.iter(W + "bookmarkStart")]))
        facts["dup docPr id"] = len(dups([b.get("id") for b in doc.iter(WP + "docPr")]))
        starts = collections.Counter(b.get(W + "id") for b in doc.iter(W + "bookmarkStart")); ends = collections.Counter(b.get(W + "id") for b in doc.iter(W + "bookmarkEnd"))
        facts["unbalanced bookmarks"] = sum(((starts - ends) + (ends - starts)).values())
        for kind, part in (("footnoteReference", "word/footnotes.xml"), ("endnoteReference", "word/endnotes.xml"), ("commentReference", "word/comments.xml")):
            refs = [r.get(W + "id") for r in doc.iter(W + kind)]
            if refs:
                have = set(e.get(W + "id") for e in trees[part]) if part in trees else set()
                miss = [r for r in refs if r not in have]
                if miss: issues.append(f"document: {kind} {miss[:5]} not in {part}")
        # every part: each range (bookmark, comment, permission, moved text) opens once and closes once, in that order
        ranges = (("bookmarkStart", "bookmarkEnd", "bookmark"), ("commentRangeStart", "commentRangeEnd", "comment range"), ("permStart", "permEnd", "permission"), ("moveFromRangeStart", "moveFromRangeEnd", "moveFrom"), ("moveToRangeStart", "moveToRangeEnd", "moveTo"))
        bad_ranges = 0; range_total = 0
        for n, t in trees.items():
            if not re.match(r"word/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$", n): continue
            order = {}
            for i, el in enumerate(t.iter()):
                for s_, e_, label in ranges:
                    if el.tag == W + s_: order.setdefault((label, el.get(W + "id")), []).append(("s", i))
                    elif el.tag == W + e_: order.setdefault((label, el.get(W + "id")), []).append(("e", i))
            for (label, rid), marks in order.items():
                range_total += 1
                kinds = "".join(k for k, _ in marks)
                # (a comment's range may be absent altogether: only its reference is required)
                if kinds != "se":
                    bad_ranges += 1
                    if bad_ranges <= 6: issues.append(f"{n}: {label} {rid} marks are '{kinds}' (wanted one start, then one end)")
            names_ = [el.get(W + "name") for el in t.iter(W + "bookmarkStart")]
            d = dups(names_)
            if d: issues.append(f"{n}: bookmark names more than once: {d[:5]}")
        facts["ranges"] = range_total; facts["bad ranges"] = bad_ranges
        if "word/comments.xml" in trees:
            cids = [c.get(W + "id") for c in trees["word/comments.xml"].iter(W + "comment")]
            if dups(cids): issues.append(f"comments: ids more than once {dups(cids)[:5]}")
            refd = set(r.get(W + "id") for n, t in trees.items() if n.startswith("word/") for r in t.iter(W + "commentReference"))
            facts["comments"] = f"{len(cids)} ({len(set(cids) - refd)} without a reference)"
        st = trees.get("word/styles.xml")
        if st is not None:
            facts["styles"] = f"{sum(1 for _ in st.iter(W + 'style'))} in {len(z.read('word/styles.xml'))} bytes"
            have = set(s.get(W + "styleId") for s in st.iter(W + "style"))
            used = set(e.get(W + "val") for tag in ("pStyle", "rStyle", "tblStyle") for e in doc.iter(W + tag))
            facts["styles used but not defined"] = len(used - have)
        nm = trees.get("word/numbering.xml")
        nums = set(n.get(W + "numId") for n in nm.iter(W + "num")) if nm is not None else set()
        usedn = set(e.get(W + "val") for e in doc.iter(W + "numId")) - {"0"}
        facts["numIds used but not defined"] = len(usedn - nums)
        # a table row must have a cell, a cell a paragraph at its end
        for tc in doc.iter(W + "tc"):
            kids = [k for k in tc if k.tag != W + "tcPr"]
            if not kids or kids[-1].tag != W + "p": issues.append("document: a table cell does not end with a paragraph"); break
        for tr in doc.iter(W + "tr"):
            if tr.find(W + "tc") is None and tr.find(W + "sdt") is None: issues.append("document: a table row without a cell"); break
    return issues, facts
base = {}
for path in sys.argv[1:]:
    try: issues, facts = check(path)
    except Exception as e: issues, facts = [f"could not be checked: {type(e).__name__}: {e}"], {}
    print(f"{'OK  ' if not issues else 'BAD '} {path.split('/')[-1]}: {len(issues)} issues; " + "; ".join(f"{k}={v}" for k, v in facts.items()))
    for i in issues[:12]: print("      " + i)
