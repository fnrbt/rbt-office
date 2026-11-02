# One save path

Status: design, 2026-11-02. Replaces the two-path Word save (and, after it, the Excel one).

## What is wrong now

A Word save has two paths:

- **Patched.** `Docx.paragraphEdit` diffs the paragraph as it was read with the paragraph as it is, and if the
  difference is "one stretch of plain text" or "formatting of characters that were there", the writer splices
  that into the paragraph's XML. Everything else in the XML stays as it was.
- **Written from the model.** Any other change (a link, a field, a picture, a content control, a tracked change,
  a table row insert, a document made from scratch) writes the whole paragraph or table from the model. What the
  model has no field for is lost: content controls, smart tags, custom XML, unknown wrappers, properties we do
  not model.

So every loss is found one at a time (in Word, in LibreOffice, in a customer's file) and fixed by adding one more
field through reader, buffer, export and writer. The patch path is a heuristic diff, so whether a change is kept
depends on its shape, not on what was changed. The spreadsheet has the same split (cells patched by span,
structural edits written from the model).

## What it should be

1. **One path.** Every save goes through one writer. Nothing is "regenerated"; a node that changed is written
   from its source with the change applied, a node that is new is written from the model.
2. **Byte splice.** A part nothing changed in is its original bytes. A node nothing changed in is its original
   bytes. An edit changes the bytes of what was edited and nothing else.
3. **Nothing unmodelled is lost, by construction.** What the model does not understand is kept as a span of the
   source at its position and written back as it was. Losing an element should be impossible, not tested for.
4. **Type safe.** Whether a node changed is decided by comparing the typed value read with the typed value now,
   not by diffing text. A property change is a typed edit (`set Bold true`) applied to the source element; the
   writer knows the schema order of what it inserts.

## Design

### Source tree with ranges

One XML parser for every read (it replaces the DOM, `XDocument` and scanner parsers of `XmlDsl`): it scans the
part's original string, without normalising it, into a tree of

```
SourceElement = { Name; Attrs; Start; StartTagEnd; EndTagStart; End; Children }   // offsets into the part's text
SourceText    = { Start; End; Value }                                              // decoded value, raw range
```

`slice part node` is the node's original text. The reader reads values from this tree as it reads from
`XmlNode` now (an adapter keeps `XmlNode` working while the reader moves over).

### Model nodes keep their source

- A paragraph keeps a reference to its `w:p` and to its `w:pPr`, and the property values it read from it.
- A run (a stretch of characters of one format in the text buffer) keeps a reference to the `w:r` it came from,
  its `w:rPr`, the values read, and its **wrapper chain**: the elements between the paragraph and the run
  (`w:hyperlink`, `w:sdt`, `w:smartTag`, `w:customXml`, `w:ins`, `w:del`, `w:fldSimple`, `w:bdo`, `w:dir`,
  unknown ones), each a source reference. Runs with a common chain prefix are inside the same wrapper.
- Children of a paragraph or run that the model does not hold (`w:proofErr`, `w:permStart`, unknown elements,
  `w:sdtPr` of a wrapper) are **kept anchors**: a source reference at a character offset, carried through edits
  like a bookmark is now (deleted with the text around them, never invented).
- Tables, rows, cells, sections, headers, notes, comments the same way.

The text buffer stays the editing model. These references ride on it the way bookmarks and revision marks do now;
they are data, not behaviour: an edit never has to know about them except to keep or drop them.

### One writer

```
write node =
  match node.Source with
  | Some src when node.Value = src.ValueRead -> slice src                         // untouched: the bytes
  | Some src -> startTag src (attribute edits) + children + endTag src             // changed: splice
  | None     -> canonical node                                                     // new
children = kept anchors and modelled children in document order, each written by `write`
properties = source pPr/rPr with typed edits for the properties whose value differs from the value read
```

- Text of a run that changed rewrites only its `w:t` content (and `xml:space`).
- Runs are grouped by wrapper chain; a wrapper is written by its source start and end tags around its runs, with
  its own non-run children (an `sdtPr`) copied.
- A run split by an edit is two runs with the same source `rPr`; runs joined keep the first one's.
- A part with no changed node is its original bytes (already true for unedited parts).

The spreadsheet's per-cell spans (`XlsxWriter` patch plan) become the same mechanism: cell, row, sheet element.

### Measuring it

- **Identity:** every corpus document, imported and written with no edit, is byte-identical part by part.
- **Census:** every corpus document with scripted edits (type, format, link, insert row, insert control) keeps
  every element and attribute of the original that the edit did not touch (`xmlCensusDiff`, extended).
- **Word and Excel:** saves of the corpus open in Word and Excel with the original's pages and text
  (`tools/word-parity/open-all.vbs`, `tools/excel-parity/open-all.vbs`).

## Order of work

1. Range parser in `XmlDsl` with the identity test (slice of every node = its source; speed no worse than the
   scanner); an `XmlNode` view of it so the reader can switch parser without other changes.
2. The reader on the range parser everywhere (the three parsers go); paragraphs and runs record their source
   references and read values.
3. Source references, wrapper chains and kept anchors in the text buffer, kept through edits (`textbuffer.fsx`).
4. The one writer for paragraphs; `paragraphEdit` and the regenerate path go. Identity and census tests over the
   corpus; Word open-all.
5. Tables, sections, headers, footers, notes, comments.
6. The spreadsheet.
7. Delete what is left of the old writer.

The `fix/sdt` branch (content controls threaded through as fields) is superseded by step 3; its fixture
(`synthetic/inline-controls`) and its Word check stay as tests of the new path.
