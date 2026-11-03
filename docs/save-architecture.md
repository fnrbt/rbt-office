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

### XmlPatch: the only way the writers change XML that was read

`src/Office.Formats/src/XmlPatch.fs` (`tests/office-formats/xml-patch.fsx`). Every change to XML that was read is
a typed edit of a node of the span tree, made in its text; what no edit names stays its bytes:

- `ElementEdit`: `SetAttribute` / `RemoveAttribute`; `SetChild` / `SetChildren` / `RemoveChild` / `EditChild`
  (by name, with the element to put in where there is none); `AddChild`, `SetChildWhere` / `RemoveChildWhere` /
  `EditChildWhere` (a child chosen by an attribute: the `w:num` with this `w:numId`); `EditNth` (the n-th of a
  name); `SetText` (with `xml:space` where the text needs it); `Removed`, `InsertBefore` / `InsertAfter`,
  `Prepend` / `Append`, `Replaced`, `SetAllAttributes`; `editsAt` reaches any element of the tree (a chain of
  `EditNth`), `EditAt` a child by where it starts. A child put in goes to its place in the schema order
  (`XmlPatch.Word.order`, `XmlPatch.Spreadsheet.order`; `*` for the names it does not list). Fragments
  (`readFragmentRoot` / `applyFragment`) and big parts in outline (`readPartExcept`) are edited the same way.
- `OpcXml`: relationships parts and `[Content_Types].xml` read, edited and merged as entries; the relationship
  ids XML refers to, by the relationships namespace. `PartPath`: package paths by segment. `ModelText`: the
  values the editor hands over as text (bookmark tokens, a run's tabs and breaks, link targets, data URLs).
- `apply` / `applyPart` make the edits; `diff was now` gives the edits that make what the model wrote before into
  what it writes now (attributes, children by name, repeated children in their places, text) - so a merge is
  `apply (diff was now)` on the element that was read: the model's change, and nothing else.
- `layout` / `writeLayout`: the children of a table, row, cell or paragraph as items (rows, cells, paragraphs,
  runs) in the wrappers they are in (content controls, custom XML, links, tracked changes), among what the model
  does not hold (marks, unknown elements, white space); written back with the items the model has, each `AtItem`
  (in its wrappers), `InItem` (a further part of a split item, in its wrappers), `WithPrevious` (new, in the
  wrappers of the one before) or `Outside`.
- `tests/xml-text-ratchet.sh` (in `tests/run-all.sh`): the lines of the writers and readers with text
  operations (`Substring`, `IndexOf`, `Regex`, ...) may only go down: DocumentWriter 244 -> 0, XlsxWriter
  235 -> 0, DocxReader 96 -> 0, XlsxReader 98 -> 0.

### The one scanner: how the readers read XML

`XmlDsl` holds the only code that takes XML text apart: `XmlLexer`, a pull lexer over a range of a part's
text (`Next` gives a start tag, an end tag, text, CDATA, a comment, a processing instruction; the name,
attributes - as written or decoded, into an array in one pass (`ReadAttributes`) - and text of the token are
read off it without making strings of what is not asked for; end tags must close the open element, unless
`Nesting` is off). On it:

- `tryScan` / `tryScanAt`: a part (a range of one) as a span tree.
- `tryReadElement` / `tryReadElementAt`: the element the lexer stands on, read to its end tag, the lexer then
  past it - a large part read element by element (a document's body: `DocxReader.openProgressiveFromSource`;
  a cell with an inline string). `elementEndAt`, `SkipElement` (lexed) and `SkipContent` (to the end tag of
  its name, what is between not read: a sheet's rows passed over) for what is not read. `parseElementOf`: an
  element the scanner does not read, by the platform's parser in the namespaces of the part's root.
- `rootStartTag`: the root's start tag and the lexer after it; `startTagsNamed`: the start tags of some local
  names, tag after tag, whatever they nest like (the section properties, header references and moves of a
  main part, before its body is read; the client data of a VML part, which need not be well formed);
  `SeekStartTag`: on to the next start tag of a name (a sheet's formulas, for a what-if data table).
- `tryScanOutline`: a part with the elements of some names without what they hold (a sheet without its
  rows); `tryScanKeeping`: with the text of some elements kept as written (a drawing's `a:t`).
- `endTagBefore` / `elementBefore`: the last elements of a part found from its end, read forward again by
  the scanner (the section properties that end a body, read before the body).

XlsxReader reads a sheet's rows token by token on the lexer (`scanSheetRows`, `scanRowHeaders`,
`scanFormulaCells`) and the rest of the sheet - kept elements and extensions, the auto filter, the printer
settings, whether the fast native reader may read it - in outline. Values that are text but not XML (VML
style declarations and lengths, field instructions, colours, cell references, number formats, paths) are read
by `ValueText` (with `ModelText`, `PartPath`, `OpcXml`).

### Done

- Spans: `XmlNode.Element` carries its `Span`; one scanner, the part's text unchanged (`xml-source.fsx`); the
  reader's paragraphs, runs (of every kind: text, link, picture), tables, rows and cells carry their source.
- Origins (`Document.SourceOrigin`): paragraphs, rows and cells of the model keep `ReadAt` the start of their
  element; one made like another (a row or column put in, the second half of a split) is `LikeAt` it. The writer
  finds each element by its origin in one span tree of the part (`sourceElementAt`) - no indexes of blocks and
  paragraphs, no rescanning.
- Tables: a table is its properties and grid, its rows in their wrappers (a content control around rows, around
  cells, around the table), a row its cells, a cell its paragraphs; a copy of a row or cell is its origin's
  element without its own ids. Row and column insert, cell shading, split: nothing lost (census).
- Paragraphs: one path. A paragraph that changed is its source paragraph (its start tag, its properties merged,
  `mergedProperties` on XmlPatch) with its runs: a run as it was read is its bytes; a run whose formatting or text
  changed is its source `w:r` with the model's change made in it (`diff` of the model's run before and after:
  properties merged, text `SetText` in its `w:t`); a new run next to a run that was read (typed with its style) is
  a copy of that run's element without its session ids; a field that is as it was is its runs from begin mark to
  end mark, one whose result was formatted has the result's runs changed in place. Links and tracked changes are
  source wrappers while the model's wrapping of a run is what it was. Comment and permission ranges stay where
  they stood; the bookmarks the model writes are its. The old text patcher (`scanSourceParagraphs`,
  `planParagraphChange`, `sourceRunsOf`, `runPiece`, cut and join) is deleted.
- Numbering: definitions and lists added and levels changed by edits of the part's tree.
- The rest of DocumentWriter: the body plan (the body's elements, field marks and the marks between blocks from
  the part's tree), carried and balanced range marks (a piece of the source is a range of its tree; a piece of
  the model is edited by `editsAt`), shapes moved or resized (anchor, extent and VML style edited), settings,
  note marks, content controls, relationships and content types. Its own tokenizer is gone.
- XlsxWriter: cells, rows, sheet sections, relationships, comments, VML, tables, workbook, styles, shared
  strings - all by edits of the parts' trees.

### Left

- The model's XML written by a typed builder (it is text made by `element` / `selfClosing` now).
- Bookmarks handed over typed, not as tokens in a string (`ModelText` reads them at the writer's end).
- `XmlPatch.readPartExcept` still has a tokenizer of its own (XmlPatch is changed only by additions):
  `XmlDsl.tryScanOutline` reads the same outline.
- Headers, footers, notes and comments through the same path (origins of their own part).
- Joins: the paragraph joined to another keeps nothing of its own properties (as Word).

### The original plan

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
