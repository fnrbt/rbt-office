// Incremental layout must produce exactly what the cold layout produces.
//
// Renderer.layoutDocument uses the incremental entry point; PDF export uses the
// cold one. If they disagree, what is on screen differs from what is exported
// and from where clicks land. Every scenario here uses a realistic page (US
// Letter with 1in margins, default paragraph spacing) so page boundaries,
// numbering, spacing collapse, table splitting and notes are all in play, and
// compares the two full results (every line of every block) after each edit.
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"

open Rbt.Word.TextBuffer
module Doc = Rbt.Word.Document
module Layout = Rbt.Word.Layout

let mutable pass = 0
let mutable fail = 0

let check label got expected =
    if got = expected then
        pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-42s expected %A got %A" label expected got

let checkTrue label value =
    if value then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %s" label

let settings : Layout.LayoutSettings =
    { Width = 816.0
      PageHeight = 1056.0
      FontSize = 16.0
      LineHeight = 24.0
      PaddingLeft = 96.0
      PaddingTop = 96.0
      MarginRight = 96.0
      MarginBottom = 96.0
      InTableCell = false
      TableContinuationHeaderHeight = 0.0
      LongTrailingSpaceFitSafety = 3.5; DefaultTabStop = 48.0; CompatibilityMode = 12 }

let pageSetup =
    { Doc.defaultPageSetup with
        PageWidth = settings.Width
        PageHeight = settings.PageHeight
        MarginLeft = settings.PaddingLeft
        MarginRight = settings.MarginRight
        MarginTop = settings.PaddingTop
        MarginBottom = settings.MarginBottom }

let measureText (text: string) (style: TextStyle) =
    float text.Length * style.FontSize * 0.5

let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText

let words prefix count =
    String.concat " " [ for j in 1 .. count -> sprintf "%s%d" prefix j ]

let paraWith id metadata text : Doc.ParagraphBlock =
    { Id = id; Buffer = ofString text; Metadata = metadata }

let para id text = paraWith id Doc.defaultMetadata text

let document blocks =
    { Doc.empty with Blocks = blocks; PageSetup = pageSetup; Styles = []; Selection = None }

let pageOf (y: float) = int (floor (y / (settings.PageHeight + Layout.pageGap)))

let pagesOfBlock id (layouts: Layout.BlockLayout list) =
    layouts
    |> List.filter (fun layout -> layout.BlockId = id)
    |> List.collect (fun layout -> layout.Lines |> List.map (fun line -> pageOf line.Y))
    |> List.distinct

/// First difference between two results, for a readable failure.
let describeDifference (inc: Layout.BlockLayout list) (cold: Layout.BlockLayout list) =
    if inc.Length <> cold.Length then
        sprintf "layout count %d vs %d" inc.Length cold.Length
    else
        List.zip inc cold
        |> List.tryFind (fun (a, b) -> a <> b)
        |> Option.map (fun (a, b) ->
            sprintf "block %s: incremental y=%.2f h=%.2f lines=%d prefix=%A | cold %s y=%.2f h=%.2f lines=%d prefix=%A"
                a.BlockId a.Y a.Height a.Lines.Length a.Prefix b.BlockId b.Y b.Height b.Lines.Length b.Prefix)
        |> Option.defaultValue "no difference"

let compare label state doc =
    let coldLayouts, coldPages = Layout.layoutDocumentPages settings measureText measureRange doc
    let incLayouts, incPages = Layout.layoutDocumentPagesIncremental state settings measureText measureRange doc
    if incLayouts = coldLayouts then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %s: incremental differs from cold: %s" label (describeDifference incLayouts coldLayouts)
    check (label + " pages") incPages coldPages
    coldLayouts

/// The start of a document laid out before the rest (IncrementalLayoutState.BlockLimit):
/// with the limit raised in steps - and the edits made in between - every
/// result is the layout of that many blocks on their own, and the last one,
/// with the limit at the whole document, is the layout of the whole document.
let mutable progressiveSteps = 0
let assertProgressive label (doc: Doc.Document) (edits: (string * (Doc.Document -> Doc.Document)) list) =
    let state = Layout.createIncrementalLayoutState ()
    let mutable seed = 17 + String.length label * 31
    let next bound =
        seed <- (seed * 1103515245 + 12345) &&& 0x3fffffff
        1 + (seed >>> 8) % bound
    let comparePart name limit (current: Doc.Document) =
        progressiveSteps <- progressiveSteps + 1
        state.BlockLimit <- limit
        let coldLayouts, coldPages = Layout.layoutDocumentPagesUpTo limit settings measureText measureRange current
        let incLayouts, incPages = Layout.layoutDocumentPagesIncremental state settings measureText measureRange current
        if incLayouts = coldLayouts && incPages = coldPages then pass <- pass + 1
        else
            fail <- fail + 1
            printfn "FAIL %s: the first %d blocks %s differ from those blocks laid out on their own: %s" label limit name (describeDifference incLayouts coldLayouts)
        // (Asked again, the result is handed back as it is.)
        let againLayouts, _ = Layout.layoutDocumentPagesIncremental state settings measureText measureRange current
        checkTrue (sprintf "%s: the first %d blocks %s, asked again" label limit name) (obj.ReferenceEquals(againLayouts, incLayouts))
    let mutable current = doc
    let mutable limit = 0
    let step name =
        let count = current.Blocks.Length
        limit <- min (limit + next (max 1 (count / 4 + 1))) (count + 2)
        comparePart name limit current
        checkTrue (sprintf "%s: complete only with every block %s (%d of %d)" label name limit count) (Layout.incrementalLayoutIsWhole state = (limit >= count))
    step "at first"
    step "then"
    for name, edit in edits do
        current <- edit current
        // (The edit at the limit the state had, then a step further.)
        if not (Layout.incrementalLayoutIsWhole state) then comparePart ("after " + name) limit current
        step ("past " + name)
    while limit < current.Blocks.Length do step "further"
    // The whole document: what a pass over all of it makes.
    state.BlockLimit <- System.Int32.MaxValue
    let coldLayouts, coldPages = Layout.layoutDocumentPages settings measureText measureRange current
    let incLayouts, incPages = Layout.layoutDocumentPagesIncremental state settings measureText measureRange current
    if incLayouts = coldLayouts && incPages = coldPages then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %s: laid out in parts, the whole differs from the whole laid out at once: %s" label (describeDifference incLayouts coldLayouts)
    checkTrue (label + ": complete") (Layout.incrementalLayoutIsWhole state)

/// Lay the document out cold and incrementally, then again after every edit
/// (the incremental state is carried across edits, as in the editor).
let assertEquivalent label doc (edits: (string * (Doc.Document -> Doc.Document)) list) =
    let state = Layout.createIncrementalLayoutState ()
    let initial = compare (label + " initial") state doc
    let mutable current = doc
    for name, edit in edits do
        current <- edit current
        compare (sprintf "%s after %s" label name) state current |> ignore
    assertProgressive label doc edits
    assertProgressive (label + " (edited)") current []
    state, initial, current

let typeIn blockId offset text doc =
    doc |> Doc.setCursor blockId offset |> Doc.insertText text

let mapBlocks f (doc: Doc.Document) = { doc with Blocks = f doc.Blocks }

let removeBlock id =
    mapBlocks (List.filter (function
        | Doc.Paragraph p | Doc.Heading (_, p) -> p.Id <> id
        | _ -> true))

let insertAfter id block =
    mapBlocks (List.collect (fun existing ->
        match existing with
        | Doc.Paragraph p | Doc.Heading (_, p) when p.Id = id -> [ existing; block ]
        | _ -> [ existing ]))

// ---------------------------------------------------------------------------
// Plain multi-page text, including one paragraph long enough to cross a page.
// ---------------------------------------------------------------------------
let plainDoc =
    [ for i in 1 .. 60 ->
        if i = 9 then Doc.Paragraph (para "long" (words "long" 900))
        else Doc.Paragraph (para (sprintf "p%d" i) (words (sprintf "w%d_" i) (15 + (i * 7) % 60))) ]
    |> document

let plainState, plainInitial, _ =
    assertEquivalent "plain" plainDoc
        [ "typing a character", typeIn "p30" 4 "X"
          "typing that adds a line near the top", typeIn "p2" 3 (words "added" 40)
          "deleting a paragraph", removeBlock "p5"
          "inserting a paragraph", insertAfter "p20" (Doc.Paragraph (para "inserted" (words "ins" 50)))
          "typing in the long paragraph", typeIn "long" 10 (words "more" 30) ]

checkTrue "plain document spans several pages" ((Layout.uniformPageInfos settings plainInitial).Length >= 4)
checkTrue "long paragraph is split across pages" ((pagesOfBlock "long" plainInitial).Length >= 2)

// A single keystroke re-measures only the paragraph that changed.
let keystrokeState = Layout.createIncrementalLayoutState ()
Layout.layoutDocumentIncrementalWithRange keystrokeState settings measureText measureRange plainDoc |> ignore
check "first pass measures every paragraph" (Layout.incrementalLayoutMeasuredBlockCount keystrokeState) 60
Layout.resetIncrementalLayoutMeasurements keystrokeState
Layout.layoutDocumentIncrementalWithRange keystrokeState settings measureText measureRange (typeIn "p30" 4 "X" plainDoc) |> ignore
check "keystroke re-measures one paragraph" (Layout.incrementalLayoutMeasuredBlockCount keystrokeState) 1
Layout.resetIncrementalLayoutMeasurements keystrokeState
let sameDoc = typeIn "p30" 4 "X" plainDoc
let firstResult = Layout.layoutDocumentIncrementalWithRange keystrokeState settings measureText measureRange sameDoc
let secondResult = Layout.layoutDocumentIncrementalWithRange keystrokeState settings measureText measureRange sameDoc
checkTrue "unchanged document reuses the layout" (obj.ReferenceEquals(firstResult, secondResult))

// ---------------------------------------------------------------------------
// Numbered lists: markers must count up (not "1." everywhere) and renumber
// after an edit, exactly as the cold layout does.
// ---------------------------------------------------------------------------
let numbered listId level =
    { Doc.defaultMetadata with ListStyle = Doc.Numbered; ListLevel = level; ListId = Some listId }

let listDoc =
    [ yield Doc.Paragraph (para "intro" (words "intro" 30))
      for i in 1 .. 45 -> Doc.Paragraph (paraWith (sprintf "n%d" i) (numbered "list-a" 1) (words (sprintf "item%d_" i) (5 + i % 30)))
      yield Doc.Paragraph (para "between" "between the lists")
      for i in 1 .. 6 -> Doc.Paragraph (paraWith (sprintf "m%d" i) (numbered "list-b" (if i % 3 = 0 then 2 else 1)) (sprintf "second list %d" i)) ]
    |> document

let _, listInitial, listEdited =
    assertEquivalent "numbered list" listDoc
        [ "typing in an item", typeIn "n20" 2 "X"
          "removing an item", removeBlock "n3"
          "adding an item", insertAfter "n10" (Doc.Paragraph (paraWith "n-new" (numbered "list-a" 1) "new item")) ]

let prefixOf id (layouts: Layout.BlockLayout list) =
    layouts |> List.find (fun layout -> layout.BlockId = id) |> fun layout -> layout.Prefix.Trim()

let listIncremental =
    Layout.layoutDocumentIncrementalWithRange (Layout.createIncrementalLayoutState ()) settings measureText measureRange listDoc
check "incremental numbering item 1" (prefixOf "n1" listIncremental) "1."
check "incremental numbering item 2" (prefixOf "n2" listIncremental) "2."
check "incremental numbering item 45" (prefixOf "n45" listIncremental) "45."
check "incremental numbering second list restarts" (prefixOf "m1" listIncremental) "1."
checkTrue "numbered list spans pages" ((Layout.uniformPageInfos settings listInitial).Length >= 2)

// ---------------------------------------------------------------------------
// A table that crosses a page boundary, with text before and after.
// ---------------------------------------------------------------------------
let templateTable =
    (Doc.insertTable 1 2 Doc.empty).Blocks
    |> List.pick (function Doc.Table table -> Some table | _ -> None)
let templateRow = templateTable.Rows.Head
let templateCell = templateRow.Cells.Head

let cell id text = { templateCell with Content = [ para id text ] }

let table id rowCount =
    Doc.Table
        { templateTable with
            Id = id
            Rows =
                [ for r in 1 .. rowCount ->
                    { templateRow with
                        Cells =
                            [ cell (sprintf "%s-r%d-a" id r) (words (sprintf "r%da" r) (3 + r % 12))
                              cell (sprintf "%s-r%d-b" id r) (words (sprintf "r%db" r) (2 + (r * 5) % 25)) ] } ] }

let tableDoc =
    [ for i in 1 .. 6 -> Doc.Paragraph (para (sprintf "tb%d" i) (words "before" 60))
      yield table "t1" 40
      for i in 1 .. 4 -> Doc.Paragraph (para (sprintf "ta%d" i) (words "after" 40)) ]
    |> document

let _, tableInitial, _ =
    assertEquivalent "table across pages" tableDoc
        [ "typing before the table", typeIn "tb1" 3 (words "push" 45)
          "typing in a cell", typeIn "t1-r20-b" 2 (words "cell" 30)
          "typing after the table", typeIn "ta2" 3 "X" ]

let tablePages =
    tableInitial
    |> List.filter (fun layout -> layout.BlockId.StartsWith "t1-")
    |> List.map (fun layout -> pageOf layout.Y)
    |> List.distinct
checkTrue "table rows land on more than one page" (tablePages.Length >= 2)

// A vertically merged cell that is split with the rows it spans, and rows
// kept with the next: what an edit lays out again is what a layout from
// nothing makes (the merged content is placed once the last of its rows is).
let mergedTable id rowCount =
    Doc.Table
        { templateTable with
            Id = id
            Rows =
                [ for r in 1 .. rowCount ->
                    { templateRow with
                        IsHeader = (r = 1)
                        Cells =
                            [ (if r = 1 then cell (sprintf "%s-h" id) "head"
                               elif r = 2 then { templateCell with VerticalMerge = Some Doc.RestartMerge; Content = [ for i in 1 .. 14 -> para (sprintf "%s-m%d" id i) (words (sprintf "m%d" i) 30) ] }
                               else { templateCell with VerticalMerge = Some Doc.ContinueMerge; Content = [ para (sprintf "%s-c%d" id r) "" ] })
                              { templateCell with
                                    Content = [ paraWith (sprintf "%s-r%d" id r) (if r % 7 = 0 then { Doc.defaultMetadata with KeepNext = Some true } else Doc.defaultMetadata) (words (sprintf "r%d" r) (2 + (r * 5) % 9)) ] } ] } ] }

let mergedDoc =
    [ for i in 1 .. 7 -> Doc.Paragraph (para (sprintf "mb%d" i) (words "before" 60))
      yield mergedTable "mt" 34
      for i in 1 .. 3 -> Doc.Paragraph (para (sprintf "ma%d" i) (words "after" 40)) ]
    |> document

let _, mergedInitial, _ =
    assertEquivalent "merged cell across pages" mergedDoc
        [ "typing before the table", typeIn "mb1" 3 (words "push" 45)
          "typing in the merged cell", typeIn "mt-m3" 2 (words "more" 60)
          "typing in a row beside it", typeIn "mt-r20" 2 (words "cell" 30)
          "typing after the table", typeIn "ma2" 3 "X" ]
checkTrue "the merged cell's content is on more than one page"
    ((mergedInitial |> List.filter (fun layout -> layout.BlockId.StartsWith "mt-m") |> List.collect (fun layout -> layout.Lines) |> List.map (fun line -> pageOf line.Y) |> List.distinct).Length >= 2)

// ---------------------------------------------------------------------------
// Spacing before/after collapses between paragraphs (max, not sum), headings
// keep with the next paragraph, and widow control moves short tails.
// ---------------------------------------------------------------------------
let spaced before after =
    { Doc.defaultMetadata with SpacingBefore = Some before; SpacingAfter = Some after }

let spacingDoc =
    [ for i in 1 .. 70 ->
        if i % 9 = 0 then Doc.Heading (1, paraWith (sprintf "h%d" i) (spaced 24.0 6.0) (sprintf "Heading %d" i))
        else Doc.Paragraph (paraWith (sprintf "s%d" i) (spaced (float (i % 4) * 6.0) (float (i % 5) * 4.0)) (words "sp" (20 + (i * 11) % 70))) ]
    |> document

let _, spacingInitial, _ =
    assertEquivalent "spacing and headings" spacingDoc
        [ "typing", typeIn "s10" 5 "X"
          "typing that reflows", typeIn "s1" 5 (words "reflow" 80)
          "removing a heading", removeBlock "h9" ]

let spacingLayout id = spacingInitial |> List.find (fun layout -> layout.BlockId = id)
let s1, s2 = spacingLayout "s1", spacingLayout "s2"
// s1 has 4pt after, s2 has 12pt before: the gap is the larger (12pt = 16px), not the sum.
check "spacing between paragraphs collapses"
    (System.Math.Round((s2.Lines.Head.Y - (s1.Lines |> List.last |> fun line -> line.Y + line.Height)), 2))
    16.0

// ---------------------------------------------------------------------------
// Footnotes and endnotes.
// ---------------------------------------------------------------------------
let noted id noteType noteId text noteText =
    let buffer = ofString text |> setNote (text.Length - 1) 1 (Some noteId) (Some noteType) (Some noteText)
    Doc.Paragraph { Id = id; Buffer = buffer; Metadata = Doc.defaultMetadata }

let notesDoc =
    [ for i in 1 .. 40 ->
        if i % 7 = 0 then noted (sprintf "fn%d" i) "footnote" (string (i / 7)) (words "noted" 30 + " x") (words "footnote" 25)
        elif i % 11 = 0 then noted (sprintf "en%d" i) "endnote" (string (i / 11)) (words "ended" 30 + " x") (words "endnote" 25)
        else Doc.Paragraph (para (sprintf "np%d" i) (words "body" (20 + (i * 13) % 60))) ]
    |> document

let _, notesInitial, _ =
    assertEquivalent "footnotes and endnotes" notesDoc
        [ "typing", typeIn "np1" 3 "X"
          "typing that moves a footnote", typeIn "np2" 3 (words "move" 120) ]

checkTrue "footnotes are laid out" (notesInitial |> List.exists (fun layout -> layout.BlockId.StartsWith Layout.footnoteBlockIdPrefix))
checkTrue "endnotes are laid out" (notesInitial |> List.exists (fun layout -> layout.BlockId.StartsWith Layout.endnoteBlockIdPrefix))
let notesIncremental =
    Layout.layoutDocumentIncrementalWithRange (Layout.createIncrementalLayoutState ()) settings measureText measureRange notesDoc
checkTrue "incremental keeps endnotes" (notesIncremental |> List.exists (fun layout -> layout.BlockId.StartsWith Layout.endnoteBlockIdPrefix))

// ---------------------------------------------------------------------------
// A lastRenderedPageBreak hint inside a paragraph must not send lines a
// million pixels down the canvas.
// ---------------------------------------------------------------------------
let hintedBuffer =
    let text = "first line of the paragraph second line after the hint"
    ofString text |> applyStyle 28 (text.Length - 28) (fun style -> { style with RenderedPageBreakBefore = true })

let hintDoc =
    document
        [ Doc.Paragraph (para "hint-before" "before")
          Doc.Paragraph { Id = "hint"; Buffer = hintedBuffer; Metadata = Doc.defaultMetadata }
          Doc.Paragraph (para "hint-after" "after") ]

let _, hintInitial, _ = assertEquivalent "rendered page break hint" hintDoc [ "typing", typeIn "hint-after" 2 "X" ]
checkTrue "hinted paragraph stays within two pages" ((Layout.uniformPageInfos settings hintInitial).Length <= 2)

// ---------------------------------------------------------------------------
// Floating image, sections and columns go through the same pass.
// ---------------------------------------------------------------------------
let floatingBuffer =
    ofString (words "wrap" 80)
    |> insertImageWithMetadata 6 "memory://float" false (Some 120.0) (Some 90.0) None None None true (Some "square") (Some 200.0) (Some 10.0) None None None None None None None None None None None None None None None None None None

let floatDoc =
    [ yield Doc.Paragraph (para "fl-before" (words "lead" 40))
      yield Doc.Paragraph { Id = "fl"; Buffer = floatingBuffer; Metadata = Doc.defaultMetadata }
      for i in 1 .. 12 -> Doc.Paragraph (para (sprintf "fl%d" i) (words "tail" 50)) ]
    |> document

assertEquivalent "floating image" floatDoc
    [ "typing before the float", typeIn "fl-before" 2 (words "shift" 30)
      "typing after the float", typeIn "fl5" 2 "X" ]
|> ignore

let columns : Doc.SectionColumns = { Count = 2; Gap = 24.0; Separator = true; Widths = [] }

let columnDoc =
    { document [ for i in 1 .. 60 -> Doc.Paragraph (para (sprintf "mc%d" i) (words "col" (10 + (i * 3) % 40))) ] with
        PageSetup = { pageSetup with Columns = Some columns } }

assertEquivalent "two columns" columnDoc
    [ "typing", typeIn "mc6" 2 "X"
      "typing that reflows", typeIn "mc2" 2 (words "reflow" 60) ]
|> ignore

let sectionDoc =
    [ for i in 1 .. 30 ->
        let metadata =
            if i = 10 then { Doc.defaultMetadata with SectionBreak = Some "nextPage" }
            elif i = 20 then { Doc.defaultMetadata with SectionBreak = Some "continuous"; SectionColumns = Some columns }
            else Doc.defaultMetadata
        Doc.Paragraph (paraWith (sprintf "sec%d" i) metadata (words "sec" (15 + (i * 7) % 50))) ]
    |> document

assertEquivalent "sections" sectionDoc
    [ "typing in the first section", typeIn "sec3" 2 (words "grow" 40)
      "typing in the last section", typeIn "sec25" 2 "X" ]
|> ignore

// ---------------------------------------------------------------------------
// Style changes invalidate reused measurements.
// ---------------------------------------------------------------------------
let styleState, _, styledDoc = assertEquivalent "style change" plainDoc []
let withMetadataChange =
    styledDoc
    |> mapBlocks (List.map (function
        | Doc.Paragraph p when p.Id = "p12" -> Doc.Paragraph { p with Metadata = { p.Metadata with IndentLeft = Some 80.0; LineSpacing = Some 2.0 } }
        | block -> block))
compare "style change after paragraph formatting" styleState withMetadataChange |> ignore
let withBold =
    withMetadataChange
    |> mapBlocks (List.map (function
        | Doc.Paragraph p when p.Id = "p14" -> Doc.Paragraph { p with Buffer = p.Buffer |> applyStyle 0 20 (fun style -> { style with Bold = true; FontSize = 30.0 }) }
        | block -> block))
compare "style change after run formatting" styleState withBold |> ignore

// ---------------------------------------------------------------------------
// No layout state is shared between documents or threads.
//
// The library is used server-side, where several documents are laid out at
// once. Two different documents (each with numbered lists inside a table, the
// case that used to go through a process-wide cache) are laid out on one
// thread interleaved, and concurrently on many threads; every result must
// equal the single-threaded result for that document.
// ---------------------------------------------------------------------------
let listCell id listId (texts: string list) =
    { templateCell with
        Content = texts |> List.mapi (fun i text -> paraWith (sprintf "%s-%d" id i) (numbered listId 1) text) }

let listTable id listId rowCount itemsPerCell =
    Doc.Table
        { templateTable with
            Id = id
            Rows =
                [ for r in 1 .. rowCount ->
                    { templateRow with
                        Cells =
                            [ listCell (sprintf "%s-r%d-a" id r) listId [ for i in 1 .. itemsPerCell -> words (sprintf "a%d_" i) (3 + (r + i) % 9) ]
                              listCell (sprintf "%s-r%d-b" id r) listId [ for i in 1 .. itemsPerCell -> words (sprintf "b%d_" i) (2 + (r * i) % 14) ] ] } ] }

// The two documents reuse the same block ids on purpose: a cache keyed by id
// alone would hand one document the other's measurements.
let sharedIdDoc variant =
    let listId, rows, items, wordCount = if variant = 0 then "list-x", 30, 2, 25 else "list-y", 18, 3, 55
    [ for i in 1 .. 12 -> Doc.Paragraph (paraWith (sprintf "shared-n%d" i) (numbered listId 1) (words (sprintf "v%d_" variant) (wordCount + i)))
      yield listTable "shared-table" listId rows items
      for i in 1 .. 25 -> Doc.Paragraph (para (sprintf "shared-p%d" i) (words (sprintf "tail%d_" variant) (wordCount + (i * 7) % 40))) ]
    |> document

let docA0, docB0 = sharedIdDoc 0, sharedIdDoc 1
let docA1, docB1 = typeIn "shared-p3" 2 (words "edit" 30) docA0, typeIn "shared-table-r2-a-0" 1 "X" docB0
let coldOf doc = Layout.layoutDocumentPages settings measureText measureRange doc
let expectedA0, expectedB0, expectedA1, expectedB1 = coldOf docA0, coldOf docB0, coldOf docA1, coldOf docB1
checkTrue "the two documents lay out differently" (fst expectedA0 <> fst expectedB0)
let cellPrefix id (layouts: Layout.BlockLayout list) = (layouts |> List.find (fun layout -> layout.BlockId = id)).Prefix.Trim()
// Cell numbering restarts at the top of each table and runs through its cells.
// The table continues the list of the 12 numbered body paragraphs before it.
check "document A numbers its table cells" (cellPrefix "shared-table-r1-a-0" (fst expectedA0), cellPrefix "shared-table-r2-a-0" (fst expectedA0)) ("13.", "17.")
check "document B numbers its table cells" (cellPrefix "shared-table-r1-a-0" (fst expectedB0), cellPrefix "shared-table-r2-a-0" (fst expectedB0)) ("13.", "19.")

// Interleaved on one thread, each document with its own state.
let stateA, stateB = Layout.createIncrementalLayoutState (), Layout.createIncrementalLayoutState ()
let incrementalOf state doc = Layout.layoutDocumentPagesIncremental state settings measureText measureRange doc
check "interleaved A first pass" (incrementalOf stateA docA0) expectedA0
check "interleaved B first pass" (incrementalOf stateB docB0) expectedB0
check "interleaved A after edit" (incrementalOf stateA docA1) expectedA1
check "interleaved B after edit" (incrementalOf stateB docB1) expectedB1
check "interleaved A back to original" (incrementalOf stateA docA0) expectedA0
check "interleaved B back to original" (incrementalOf stateB docB0) expectedB0

// Interleaved through ONE state: nothing measured for one document may be
// reused for the other, even though they share block ids.
let sharedState = Layout.createIncrementalLayoutState ()
for round in 1 .. 3 do
    check (sprintf "shared state round %d A" round) (incrementalOf sharedState docA0) expectedA0
    check (sprintf "shared state round %d B" round) (incrementalOf sharedState docB0) expectedB0
    check (sprintf "shared state round %d A edited" round) (incrementalOf sharedState docA1) expectedA1
    check (sprintf "shared state round %d B edited" round) (incrementalOf sharedState docB1) expectedB1
    // Cold layouts in between must not disturb or be disturbed by the state.
    check (sprintf "shared state round %d cold B" round) (coldOf docB0) expectedB0

// Concurrently on several threads: cold layouts, and incremental layouts with
// a state per task (a state is one editor's and is not itself thread-safe).
let concurrentFailures = System.Collections.Concurrent.ConcurrentQueue<string>()
let concurrentThreads = System.Collections.Concurrent.ConcurrentDictionary<int, bool>()
System.Threading.Tasks.Parallel.For(
    0, 64,
    System.Threading.Tasks.ParallelOptions(MaxDegreeOfParallelism = 8),
    fun iteration ->
        concurrentThreads.[System.Threading.Thread.CurrentThread.ManagedThreadId] <- true
        let original, edited, expectedOriginal, expectedEdited =
            if iteration % 2 = 0 then docA0, docA1, expectedA0, expectedA1 else docB0, docB1, expectedB0, expectedB1
        try
            if coldOf original <> expectedOriginal then concurrentFailures.Enqueue(sprintf "iteration %d: cold original" iteration)
            let state = Layout.createIncrementalLayoutState ()
            if incrementalOf state original <> expectedOriginal then concurrentFailures.Enqueue(sprintf "iteration %d: incremental original" iteration)
            if incrementalOf state edited <> expectedEdited then concurrentFailures.Enqueue(sprintf "iteration %d: incremental edited" iteration)
            if coldOf edited <> expectedEdited then concurrentFailures.Enqueue(sprintf "iteration %d: cold edited" iteration)
        with ex -> concurrentFailures.Enqueue(sprintf "iteration %d: %s" iteration ex.Message))
|> ignore
check "concurrent layouts match the single-threaded results" (List.ofSeq concurrentFailures) []
checkTrue "concurrent layouts ran on more than one thread" (concurrentThreads.Count > 1)

// ---------------------------------------------------------------------------
// What is kept between passes of a long document: the rest of a paragraph
// that breaks over a page (in a document with sections it is laid out as a
// paragraph of its own, with the paragraph's id), the plans of tables, and
// the paint commands of paragraphs.
// ---------------------------------------------------------------------------
let brokenDoc =
    [ for i in 1 .. 40 ->
        let metadata = if i = 30 then { Doc.defaultMetadata with SectionBreak = Some "nextPage" } else Doc.defaultMetadata
        // Every fifth paragraph is long enough to run over a page.
        Doc.Paragraph (paraWith (sprintf "br%d" i) metadata (words (sprintf "b%d_" i) (if i % 5 = 0 then 700 else 20 + (i * 7) % 40))) ]
    |> document
let brokenState, brokenInitial, _ =
    assertEquivalent "paragraphs broken over pages" brokenDoc
        [ "typing a character", typeIn "br12" 3 "X"
          "typing in a broken paragraph", typeIn "br10" 5 "Y"
          "typing that moves everything", typeIn "br2" 3 (words "push" 70) ]
checkTrue "several paragraphs break over a page"
    ([ for i in 5 .. 5 .. 25 -> (pagesOfBlock (sprintf "br%d" i) brokenInitial).Length ] |> List.filter (fun pages -> pages >= 2) |> List.length >= 3)
let brokenKeystroke = Layout.createIncrementalLayoutState ()
Layout.layoutDocumentIncrementalWithRange brokenKeystroke settings measureText measureRange brokenDoc |> ignore
Layout.resetIncrementalLayoutMeasurements brokenKeystroke
Layout.layoutDocumentIncrementalWithRange brokenKeystroke settings measureText measureRange (typeIn "br12" 3 "X" brokenDoc) |> ignore
check "a keystroke measures one paragraph, not the ones broken over pages" (Layout.incrementalLayoutMeasuredBlockCount brokenKeystroke) 1
Layout.resetIncrementalLayoutMeasurements brokenKeystroke
Layout.layoutDocumentIncrementalWithRange brokenKeystroke settings measureText measureRange (typeIn "br12" 3 "XY" brokenDoc) |> ignore
check "and so does the next" (Layout.incrementalLayoutMeasuredBlockCount brokenKeystroke) 1

// Table plans are kept within a scope: a keystroke outside the tables measures no cell again.
let mutable measured = 0
let countingMeasure (text: string) (style: TextStyle) =
    measured <- measured + 1
    measureText text style
let countingRange = Layout.rangeMeasurerFromTextMeasurer countingMeasure
let planDoc =
    [ yield Doc.Paragraph (para "pl-first" (words "first" 30))
      for t in 1 .. 4 do
          yield table (sprintf "pt%d" t) 12
          yield Doc.Paragraph (para (sprintf "pl%d" t) (words "between" 25))
      yield Doc.Paragraph (para "pl-last" (words "last" 30)) ]
    |> document
let scope = obj ()
let planState = Layout.createIncrementalLayoutState ()
let inScope (run: unit -> 'T) = Layout.withTablePlanScope scope run
let coldPlan = Layout.layoutDocumentPages settings measureText measureRange planDoc
let firstPlan = inScope (fun () -> Layout.layoutDocumentPagesIncremental planState settings countingMeasure countingRange planDoc)
check "tables in a scope: the cold layout" firstPlan coldPlan
let firstMeasured = measured
measured <- 0
let typedLast = typeIn "pl-last" 2 "X" planDoc
let secondPlan = inScope (fun () -> Layout.layoutDocumentPagesIncremental planState settings countingMeasure countingRange typedLast)
check "after a keystroke below the tables: the cold layout" secondPlan (Layout.layoutDocumentPages settings measureText measureRange typedLast)
checkTrue (sprintf "and no table is measured again (%d measurements, %d at first)" measured firstMeasured) (measured * 20 < firstMeasured)
measured <- 0
let typedCell = typeIn "pt2-r3-a" 1 (words "cell" 12) typedLast
let thirdPlan = inScope (fun () -> Layout.layoutDocumentPagesIncremental planState settings countingMeasure countingRange typedCell)
check "after typing in a cell (the tables below move): the cold layout" thirdPlan (Layout.layoutDocumentPages settings measureText measureRange typedCell)
let typedFirst = typeIn "pl-first" 2 (words "push" 60) typedCell
check "after typing above every table: the cold layout"
    (inScope (fun () -> Layout.layoutDocumentPagesIncremental planState settings countingMeasure countingRange typedFirst))
    (Layout.layoutDocumentPages settings measureText measureRange typedFirst)
// Another scope (other measurers) does not see these plans.
let wide (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.7
let wideRange = Layout.rangeMeasurerFromTextMeasurer wide
check "another scope with other measurers lays out with those"
    (Layout.withTablePlanScope (obj ()) (fun () -> Layout.layoutDocumentPagesIncremental (Layout.createIncrementalLayoutState ()) settings wide wideRange typedFirst))
    (Layout.layoutDocumentPages settings wide wideRange typedFirst)
check "and the first scope still has its own"
    (inScope (fun () -> Layout.layoutDocumentPagesIncremental planState settings countingMeasure countingRange typedFirst))
    (Layout.layoutDocumentPages settings measureText measureRange typedFirst)

// Paint commands kept with paragraph layouts: what a fresh collection gives.
module Paint = Rbt.Word.Paint
let commandsOf (doc: Doc.Document) (layouts: Layout.BlockLayout list) (pages: Layout.PageInfo list) =
    let _, commands, _ = Paint.collectDocumentCommandsWithLayoutsAndPages settings "Calibri" measureText measureRange doc layouts pages
    commands
let paintState = Layout.createIncrementalLayoutState ()
let paintOnce (doc: Doc.Document) =
    let layouts, pages = Layout.withTablePlanScope scope (fun () -> Layout.layoutDocumentPagesIncremental paintState settings measureText measureRange doc)
    Layout.withTablePlanScope scope (fun () -> commandsOf doc layouts pages)
let freshCommands (doc: Doc.Document) =
    let layouts, pages = Layout.layoutDocumentPages settings measureText measureRange doc
    commandsOf doc layouts pages
for label, doc in [ "paragraphs broken over pages", brokenDoc; "tables", planDoc; "sections", sectionDoc; "plain", plainDoc ] do
    check (label + ": commands of the incremental layout") (paintOnce doc) (freshCommands doc)
    check (label + ": collected again") (paintOnce doc) (freshCommands doc)
    let firstParagraph = doc.Blocks |> List.pick (function Doc.Paragraph p -> Some p.Id | _ -> None)
    let edited = typeIn firstParagraph 2 "Z" doc
    check (label + ": after a keystroke") (paintOnce edited) (freshCommands edited)

// ---------------------------------------------------------------------------
// Long random edit sequences. A pass over an edited document starts at the
// first block that differs, from the state the last pass had there, and stops
// where its state is the last pass's again; the layouts before and after are
// the last pass's. Whatever the edits - typing, Enter, a merge, blocks removed
// and inserted, a list item, a page break, keep-with-next, a footnote, a
// section break, in body text and in table cells - the result is the cold
// layout's, and so are the paint commands collected from it.
// ---------------------------------------------------------------------------
let randomDocument (random: System.Random) (sections: bool) (twoColumns: bool) =
    let blocks =
        [ for i in 1 .. 90 ->
            let id = sprintf "r%d" i
            let length = 4 + random.Next 70
            match random.Next 24 with
            | 0 -> table (sprintf "rt%d" i) (2 + random.Next 14)
            | 1 -> noted id "footnote" (sprintf "f%d" i) (words "noted" length + " x") (words "footnote" (5 + random.Next 30))
            | 2 when i % 2 = 0 -> noted id "endnote" (sprintf "e%d" i) (words "ended" length + " x") (words "endnote" (5 + random.Next 20))
            | 3 | 4 | 5 -> Doc.Paragraph (paraWith id (numbered (if i % 2 = 0 then "ra" else "rb") (1 + random.Next 2)) (words "item" (2 + random.Next 30)))
            | 6 -> Doc.Paragraph (paraWith id { Doc.defaultMetadata with KeepNext = Some true } (words "heading" (2 + random.Next 6)))
            | 7 -> Doc.Paragraph (paraWith id { Doc.defaultMetadata with PageBreakBefore = random.Next 3 = 0; SpacingBefore = Some 18.0; SpacingAfter = Some 4.0 } (words "spaced" length))
            | 8 -> Doc.Paragraph (para id "")
            | 9 when sections ->
                let metadata =
                    if random.Next 2 = 0 then { Doc.defaultMetadata with SectionBreak = Some "nextPage" }
                    else { Doc.defaultMetadata with SectionBreak = Some "continuous"; SectionColumns = (if random.Next 2 = 0 then Some { columns with Separator = false } else None) }
                Doc.Paragraph (paraWith id metadata (words "sectionend" (3 + random.Next 20)))
            | 10 -> Doc.Paragraph (para id (words "long" (250 + random.Next 500)))
            | _ -> Doc.Paragraph (para id (words (sprintf "w%d_" i) length)) ]
    let doc = document blocks
    if twoColumns then { doc with PageSetup = { pageSetup with Columns = Some { columns with Separator = false } } } else doc

let paragraphIds (doc: Doc.Document) =
    [ for block in doc.Blocks do
        match block with
        | Doc.Paragraph p | Doc.Heading (_, p) -> yield p.Id, Rbt.Word.TextBuffer.length p.Buffer, true
        | Doc.Table t ->
            for row in t.Rows do
                for c in row.Cells do
                    for p in c.Content -> p.Id, Rbt.Word.TextBuffer.length p.Buffer, false
        | _ -> () ]

let mutable freshId = 0
let randomEdit (random: System.Random) (doc: Doc.Document) : string * Doc.Document =
    let ids = paragraphIds doc
    let id, length, topLevel = ids.[random.Next ids.Length]
    let at = if length = 0 then 0 else random.Next (length + 1)
    let mapParagraph (change: Doc.ParagraphBlock -> Doc.ParagraphBlock) =
        doc |> mapBlocks (List.map (function
            | Doc.Paragraph p when p.Id = id -> Doc.Paragraph (change p)
            | other -> other))
    match random.Next 16 with
    | 0 | 1 | 2 | 3 -> "a character", typeIn id at "x" doc
    | 4 | 5 -> "some words", typeIn id at (" " + words "typed" (1 + random.Next 40)) doc
    | 6 -> "Enter", doc |> Doc.setCursor id at |> Doc.splitBlock
    | 7 -> "Backspace at the start", doc |> Doc.setCursor id 0 |> Doc.deleteBackward
    | 8 when topLevel && ids.Length > 30 -> "a block removed", removeBlock id doc
    | 9 when topLevel ->
        freshId <- freshId + 1
        "a paragraph inserted", insertAfter id (Doc.Paragraph (para (sprintf "new%d" freshId) (words "new" (1 + random.Next 80)))) doc
    | 10 when topLevel ->
        freshId <- freshId + 1
        "a table inserted", insertAfter id (table (sprintf "newt%d" freshId) (1 + random.Next 6)) doc
    | 11 when topLevel -> "keep with next toggled", mapParagraph (fun p -> { p with Metadata = { p.Metadata with KeepNext = (if p.Metadata.KeepNext = Some true then None else Some true) } })
    | 12 when topLevel -> "a list item toggled", mapParagraph (fun p -> { p with Metadata = (if p.Metadata.ListStyle = Doc.NoList then { p.Metadata with ListStyle = Doc.Numbered; ListLevel = 1; ListId = Some "ra" } else { p.Metadata with ListStyle = Doc.NoList; ListLevel = 0; ListId = None }) })
    | 13 when topLevel -> "a page break toggled", mapParagraph (fun p -> { p with Metadata = { p.Metadata with PageBreakBefore = not p.Metadata.PageBreakBefore } })
    | 14 when topLevel && length > 2 ->
        freshId <- freshId + 1
        "a footnote added", mapParagraph (fun p -> { p with Buffer = p.Buffer |> setNote (length - 1) 1 (Some (sprintf "nf%d" freshId)) (Some "footnote") (Some (words "late" (3 + random.Next 20))) })
    | 15 when topLevel && length > 0 -> "text deleted", mapParagraph (fun p -> { p with Buffer = Rbt.Word.TextBuffer.delete 0 (max 1 (length / 2)) p.Buffer })
    | _ -> "a character", typeIn id at "y" doc

let paintScope = obj ()
let mutable keptPasses = 0
let mutable chunkedPaints = 0
let mutable blocksPainted = 0
let mutable blocksCollected = 0
let mutable deferredChunks = 0
for seed, sections, twoColumns, steps in [ 11, false, false, 140; 12, true, false, 140; 13, true, false, 140; 14, false, true, 60; 15, false, false, 140; 16, true, false, 100 ] do
    let random = System.Random seed
    let state = Layout.createIncrementalLayoutState ()
    let mutable doc = randomDocument random sections twoColumns
    let mutable failed = false
    let paintCache = Paint.createPaintCache ()
    let mutable previousChunks : Paint.PaintChunk array option = None
    let mutable previousCommands : Paint.PaintCommand list = []
    let compareAll label =
        if not failed then
            let coldLayouts, coldPages = Layout.layoutDocumentPages settings measureText measureRange doc
            let incLayouts, incPages = Layout.withTablePlanScope paintScope (fun () -> Layout.layoutDocumentPagesIncremental state settings measureText measureRange doc)
            if incLayouts <> coldLayouts || incPages <> coldPages then
                failed <- true
                fail <- fail + 1
                printfn "FAIL random %d %s: incremental differs from cold: %s" seed label (describeDifference incLayouts coldLayouts)
            else
                pass <- pass + 1
                let coldCommands = commandsOf doc coldLayouts coldPages
                let incCommands = Layout.withTablePlanScope paintScope (fun () -> commandsOf doc incLayouts incPages)
                if incCommands <> coldCommands then
                    failed <- true
                    fail <- fail + 1
                    printfn "FAIL random %d %s: paint commands differ" seed label
                // The commands in chunks, the chunks of what did not change kept from the paint before.
                match state.LastSlices with
                | Some slices ->
                    match Layout.withTablePlanScope paintScope (fun () -> Paint.collectDocumentChunks paintCache settings "Calibri" measureText measureRange doc slices incPages) with
                    | Some chunks ->
                        chunkedPaints <- chunkedPaints + 1
                        // A chunk whose commands are not collected yet stands for a band: they lie in it.
                        for chunk in chunks do
                            if chunk.Pending.IsSome then
                                deferredChunks <- deferredChunks + 1
                                let top, bottom = chunk.Top, chunk.Bottom
                                let forced = Paint.forceChunk chunk
                                if forced.Commands.Length > 0 && (forced.Top < top || forced.Bottom > bottom) then
                                    failed <- true
                                    fail <- fail + 1
                                    printfn "FAIL random %d %s: a paragraph's commands (%.1f .. %.1f) lie outside the band it stood for (%.1f .. %.1f)" seed label forced.Top forced.Bottom top bottom
                        blocksPainted <- blocksPainted + doc.Blocks.Length
                        blocksCollected <- blocksCollected + Paint.paintCacheCollectedBlocks paintCache
                        let chunkCommands = Paint.commandsOfChunks chunks
                        if chunkCommands <> coldCommands then
                            failed <- true
                            fail <- fail + 1
                            let at = List.zip (List.truncate (min chunkCommands.Length coldCommands.Length) chunkCommands) (List.truncate (min chunkCommands.Length coldCommands.Length) coldCommands) |> List.tryFindIndex (fun (a, b) -> a <> b)
                            printfn "FAIL random %d %s: the commands in chunks differ (%d vs %d commands, first difference at %A)" seed label chunkCommands.Length coldCommands.Length at
                        else
                            pass <- pass + 1
                            // What the chunks' extents say is where they differ from the paint before.
                            match previousChunks with
                            | Some before ->
                                match Paint.chunkDifferenceBand before chunks with
                                | Choice1Of2 None -> if previousCommands <> coldCommands then (failed <- true; fail <- fail + 1; printfn "FAIL random %d %s: no difference band although the commands differ" seed label)
                                | Choice1Of2 (Some (top, bottom)) ->
                                    let outside commands = commands |> List.filter (fun command -> match Paint.commandYExtent command with Some (t, b) -> b < top || t > bottom | None -> false)
                                    if outside previousCommands <> outside coldCommands then
                                        failed <- true
                                        fail <- fail + 1
                                        printfn "FAIL random %d %s: commands outside the difference band %.1f..%.1f differ" seed label top bottom
                                | Choice2Of2 () -> ()
                            | None -> ()
                        previousChunks <- Some chunks
                        previousCommands <- coldCommands
                    | None -> previousChunks <- None
                | None -> previousChunks <- None
    compareAll "initial"
    for step in 1 .. steps do
        let name, edited = randomEdit random doc
        doc <- edited
        compareAll (sprintf "step %d (%s)" step name)
        if state.Memo.MemoKeptPass then keptPasses <- keptPasses + 1
// The same, with the document laid out and painted in parts: the limit is
// raised between the edits (as a host raises it in idle slices while the
// user already types). Every result is the layout of that many blocks on
// their own, its chunks hold the commands of that layout, and the whole is
// the whole laid out at once.
let mutable partPaints = 0
for seed, sections, twoColumns, steps in [ 21, false, false, 60; 22, true, false, 60; 23, true, false, 60; 24, false, true, 30; 25, true, false, 60 ] do
    let random = System.Random seed
    let state = Layout.createIncrementalLayoutState ()
    let mutable doc = randomDocument random sections twoColumns
    let mutable failed = false
    let mutable limit = 0
    let paintCache = Paint.createPaintCache ()
    let compareAll label =
        if not failed then
            state.BlockLimit <- limit
            let coldLayouts, coldPages = Layout.layoutDocumentPagesUpTo limit settings measureText measureRange doc
            let incLayouts, incPages = Layout.withTablePlanScope paintScope (fun () -> Layout.layoutDocumentPagesIncremental state settings measureText measureRange doc)
            if incLayouts <> coldLayouts || incPages <> coldPages then
                failed <- true
                fail <- fail + 1
                printfn "FAIL random %d %s (the first %d of %d blocks): in parts differs from those blocks on their own: %s" seed label limit doc.Blocks.Length (describeDifference incLayouts coldLayouts)
            else
                pass <- pass + 1
                // (The commands of the blocks that are laid out: of the document cut there.)
                // (With the paragraph that ends the section the cut falls in: it says what columns the
                // section has - a table of it is painted on its column - and has no layout to paint.)
                let cutSectionEnd =
                    doc.Blocks |> List.skip (min limit doc.Blocks.Length)
                    |> List.tryFind (fun block -> match Doc.tryGetParagraph block with Some para -> para.Metadata.SectionBreak.IsSome | None -> false)
                let coldCommands = commandsOf { doc with Blocks = List.truncate limit doc.Blocks @ Option.toList cutSectionEnd } coldLayouts coldPages
                match state.LastSlices with
                | Some slices ->
                    match Layout.withTablePlanScope paintScope (fun () -> Paint.collectDocumentChunks paintCache settings "Calibri" measureText measureRange doc slices incPages) with
                    | Some chunks ->
                        partPaints <- partPaints + 1
                        let chunkCommands = Paint.commandsOfChunks chunks
                        if chunkCommands <> coldCommands then
                            failed <- true
                            fail <- fail + 1
                            let firstDifference = List.zip (List.truncate (min chunkCommands.Length coldCommands.Length) chunkCommands) (List.truncate (min chunkCommands.Length coldCommands.Length) coldCommands) |> List.tryFind (fun (a, b) -> a <> b)
                            printfn "FAIL random %d %s (the first %d of %d blocks): the commands in chunks differ (%d vs %d commands, the first that differ: %A)" seed label limit doc.Blocks.Length chunkCommands.Length coldCommands.Length firstDifference
                        else pass <- pass + 1
                    | None -> ()
                | None -> ()
    let further () = limit <- min (limit + 1 + random.Next (doc.Blocks.Length / 6 + 1)) (doc.Blocks.Length + 3)
    further ()
    compareAll "initial"
    for step in 1 .. steps do
        let name, edited = randomEdit random doc
        doc <- edited
        compareAll (sprintf "step %d (%s)" step name)
        if random.Next 3 = 0 then
            further ()
            compareAll (sprintf "step %d, further" step)
    while limit < doc.Blocks.Length do
        further ()
        compareAll "to the end"
    limit <- System.Int32.MaxValue
    compareAll "the whole"
    if not failed then
        let wholeLayouts, wholePages = Layout.layoutDocumentPages settings measureText measureRange doc
        let incLayouts, incPages = Layout.layoutDocumentPagesIncremental state settings measureText measureRange doc
        check (sprintf "random %d: laid out in parts, the whole is the whole laid out at once" seed) (incLayouts = wholeLayouts, incPages = wholePages, Layout.incrementalLayoutIsWhole state) (true, true, true)
checkTrue (sprintf "the documents laid out in parts were painted in chunks (%d)" partPaints) (partPaints > 200)

checkTrue (sprintf "most passes of the random sequences kept layouts of the pass before (%d)" keptPasses) (keptPasses > 400)
checkTrue (sprintf "the paints in chunks (%d) collected a fraction of the blocks anew (%d of %d)" chunkedPaints blocksCollected blocksPainted) (chunkedPaints > 400 && blocksCollected * 3 < blocksPainted)
checkTrue (sprintf "paragraphs' commands were collected when wanted (%d chunks), within the band they stood for" deferredChunks) (deferredChunks > 1000)
printfn "random sequences: %d passes kept layouts; %d paints in chunks collected %d of %d blocks" keptPasses chunkedPaints blocksCollected blocksPainted

// What a block's place depends on apart from the state the pass reaches it
// with: the blocks after it (a paragraph kept with the next, an empty
// paragraph that looks for the next paragraph with text) and the block before
// it (an empty paragraph after a table). An edit to those is an edit to it.
let setMetadata id (change: Doc.BlockMetadata -> Doc.BlockMetadata) =
    mapBlocks (List.map (function
        | Doc.Paragraph p when p.Id = id -> Doc.Paragraph { p with Metadata = change p.Metadata }
        | other -> other))
let pageOfBlock id (layouts: Layout.BlockLayout list) =
    layouts |> List.tryFind (fun layout -> layout.BlockId = id) |> Option.map (fun layout -> pageOf layout.Y)
for sectioned in [ false; true ] do
    let name = if sectioned then "sections" else "one section"
    // A heading at the foot of a page is kept with the paragraph after it; when that paragraph
    // is kept with the next one too, the group no longer fits and the heading moves.
    let keepDoc k =
        [ if sectioned then yield Doc.Paragraph (paraWith "k-sec" { Doc.defaultMetadata with SectionBreak = Some "continuous" } "a section ends here")
          for i in 1 .. 40 ->
            if i = k then Doc.Paragraph (paraWith "k-head" { Doc.defaultMetadata with KeepNext = Some true } "Heading")
            elif i = k + 1 then Doc.Paragraph (para "k-next" (words "next" 45))
            else Doc.Paragraph (para (sprintf "k%d" i) (words "body" 38)) ]
        |> document
    let toggle = setMetadata "k-next" (fun metadata -> { metadata with KeepNext = Some true })
    let moving =
        [ 3 .. 30 ] |> List.tryFind (fun k ->
            let before = Layout.layoutDocumentPages settings measureText measureRange (keepDoc k) |> fst
            let after = Layout.layoutDocumentPages settings measureText measureRange (toggle (keepDoc k)) |> fst
            pageOfBlock "k-head" before <> pageOfBlock "k-head" after)
    checkTrue (name + ": a heading that moves when the paragraph after it is kept with the next") moving.IsSome
    match moving with
    | Some k -> assertEquivalent (name + ": keep with next") (keepDoc k) [ "the next paragraph is kept too", toggle; "and not", setMetadata "k-next" (fun metadata -> { metadata with KeepNext = None }) ] |> ignore
    | None -> ()

let emptyTable id =
    Doc.Table { templateTable with Id = id; Rows = [ for r in 1 .. 2 -> { templateRow with Cells = [ cell (sprintf "%s-r%d-a" id r) ""; cell (sprintf "%s-r%d-b" id r) "" ] } ] }
let afterTableDoc =
    [ yield Doc.Paragraph (para "at1" (words "before" 30))
      yield emptyTable "at-table"
      yield Doc.Paragraph (para "at-empty" "")
      for i in 1 .. 20 -> Doc.Paragraph (para (sprintf "at-p%d" i) (words "after" 30)) ]
    |> document
let _, afterTableInitial, afterTableEdited =
    assertEquivalent "an empty paragraph after a table" afterTableDoc
        [ "text typed into the table", typeIn "at-table-r1-a" 0 "x"
          "the table removed", mapBlocks (List.filter (function Doc.Table _ -> false | _ -> true)) ]
let heightOf id (layouts: Layout.BlockLayout list) = (layouts |> List.find (fun layout -> layout.BlockId = id)).Height
let coldLayoutsOf doc = Layout.layoutDocumentPages settings measureText measureRange doc |> fst
checkTrue "the empty paragraph after an empty table is taller than after text or a table with text"
    (heightOf "at-empty" afterTableInitial > heightOf "at-empty" (coldLayoutsOf afterTableEdited))

let lookAheadDoc =
    [ yield Doc.Paragraph (paraWith "la-toc" { Doc.defaultMetadata with Style = Some "TOC1" } "Contents entry")
      yield Doc.Paragraph (paraWith "la-empty" { Doc.defaultMetadata with SpacingAfter = Some 12.0 } "")
      yield table "la-table" 2
      yield Doc.Paragraph (paraWith "la-web" { Doc.defaultMetadata with Style = Some "NormalWeb" } (words "web" 20))
      for i in 1 .. 20 -> Doc.Paragraph (para (sprintf "la-p%d" i) (words "after" 30)) ]
    |> document
let _, lookAheadInitial, lookAheadEdited =
    assertEquivalent "an empty paragraph that looks ahead" lookAheadDoc
        [ "the paragraph it looks at changes style", setMetadata "la-web" (fun metadata -> { metadata with Style = None })
          "and back", setMetadata "la-web" (fun metadata -> { metadata with Style = Some "NormalWeb" })
          "a paragraph inserted between", insertAfter "la-empty" (Doc.Paragraph (para "la-between" "between")) ]
// (The empty paragraph's height no longer depends on the paragraph after the table: the layout
// took its spacing after away between a TOC1 and a NormalWeb paragraph, which Word does not - see
// tests/word/line-boxes.fsx. The edits above are still laid out as the document laid out at once.)
ignore lookAheadInitial

// ---- the highlight of a long selection, for a band of the document ----------
// A selection over many blocks is highlighted from the blocks' places and
// their layouts in the band on screen: the rectangles are those of the walk
// over the whole selection, cut to the band.
for label, doc, anchor, focus in [ "plain", plainDoc, ("p3", 2), ("p40", 5); "backwards", plainDoc, ("p50", 4), ("p2", 1); "tables", tableDoc, ("tb2", 3), ("ta3", 6); "from a cell", tableDoc, ("t1-r3-a", 1), ("ta2", 2) ] do
    let selected = Doc.setSelection { BlockId = fst anchor; Offset = snd anchor } { BlockId = fst focus; Offset = snd focus } doc
    let state = Layout.createIncrementalLayoutState ()
    let layouts = Layout.layoutDocumentIncrementalWithRange state settings measureText measureRange selected
    let whole = Paint.selectionCommandsWithRange settings measureRange selected layouts |> List.sort
    match state.LastSlices with
    | Some slices ->
        let everywhere = Paint.selectionCommandsInBand settings measureRange selected slices System.Double.NegativeInfinity System.Double.PositiveInfinity
        check (label + ": the banded highlight over the whole document") (everywhere |> Option.map List.sort) (Some whole)
        checkTrue (label + ": there is something to highlight") (whole.Length > 20)
        let top, bottom = 1500.0, 2300.0
        let inBand = whole |> List.filter (fun command -> match Paint.commandYExtent command with Some (t, b) -> b >= top && t <= bottom | None -> true)
        let banded = Paint.selectionCommandsInBand settings measureRange selected slices top bottom |> Option.defaultValue []
        checkTrue (label + ": a band holds the rectangles that reach into it, and no others far from it")
            (inBand |> List.forall (fun command -> List.contains command banded) && banded.Length < whole.Length
             && banded |> List.forall (fun command -> List.contains command whole))
        // A block's layout by the block's place.
        let lookup = Paint.sliceLayoutLookup selected slices
        check (label + ": a layout found by its block's place") (lookup (fst focus) |> Option.map (fun layout -> layout.BlockId, layout.Y)) (layouts |> List.tryFind (fun layout -> layout.BlockId = fst focus) |> Option.map (fun layout -> layout.BlockId, layout.Y))
    | None -> checkTrue (label + ": the pass hands its layouts over block by block") false

// Blocks are found by the place they were last seen at, wherever edits moved them.
do
    let mutable doc = plainDoc
    let find id = Doc.findParagraph doc id |> Option.map (fun para -> para.Id)
    check "found" (find "p30", find "p60", find "nothing") (Some "p30", Some "p60", None)
    doc <- removeBlock "p5" doc |> removeBlock "p6" |> removeBlock "p7"
    check "after blocks before it were removed" (find "p30", Doc.findBlockIndex doc "p30") (Some "p30", Some 26)
    for i in 1 .. 20 do doc <- insertAfter "p2" (Doc.Paragraph (para (sprintf "added%d" i) "added")) doc
    check "after many were inserted before it" (find "p30", Doc.findBlockIndex doc "p30", find "added7") (Some "p30", Some 46, Some "added7")
    let replaced = Doc.replaceParagraph doc "p30" (para "p30" "replaced")
    check "replaced in place: the blocks around it are the same objects"
        (List.length replaced.Blocks, obj.ReferenceEquals(List.item 45 replaced.Blocks, List.item 45 doc.Blocks), obj.ReferenceEquals(List.skip 47 replaced.Blocks, List.skip 47 doc.Blocks),
         Doc.findParagraph replaced "p30" |> Option.map (fun p -> toString p.Buffer))
        (List.length doc.Blocks, true, true, Some "replaced")
    check "an id no block has leaves the document as it is" (obj.ReferenceEquals(Doc.replaceParagraph doc "none" (para "none" "x"), doc)) true
    check "a paragraph in a table cell" ((Doc.replaceParagraph tableDoc "t1-r3-b" (para "t1-r3-b" "cell") |> fun d -> Doc.findParagraph d "t1-r3-b") |> Option.map (fun p -> toString p.Buffer)) (Some "cell")

// ---------------------------------------------------------------------------
// An edit that adds a page to a section moves the sections after it: their
// passes stand (no block of theirs is laid out again), their layouts are the
// layouts they had, one page further down - and what results is the layout
// of the whole document made from nothing. Page sizes that are not round
// numbers (A4) and a landscape section are in the document on purpose.
// ---------------------------------------------------------------------------
do
    let a4 : Doc.SectionPageSetup option =
        Some { PageWidth = 793.7; PageHeight = 1122.52; MarginLeft = 94.49; MarginRight = 94.49; MarginTop = 94.49; MarginBottom = 94.49; HeaderDistance = 47.0; FooterDistance = 47.0; DifferentFirstPage = false; EvenAndOddHeaders = false }
    let landscape : Doc.SectionPageSetup option =
        Some { PageWidth = 1122.52; PageHeight = 793.7; MarginLeft = 75.6; MarginRight = 75.6; MarginTop = 94.49; MarginBottom = 94.49; HeaderDistance = 47.0; FooterDistance = 47.0; DifferentFirstPage = false; EvenAndOddHeaders = false }
    let sectionOf name count (geometry: Doc.SectionPageSetup option) breakKind =
        [ for i in 1 .. count -> Doc.Paragraph (para (sprintf "%s-p%d" name i) (words (sprintf "%s%d_" name i) (20 + (i * 11) % 70)))
          yield Doc.Paragraph (paraWith (name + "-end") { Doc.defaultMetadata with SectionBreak = Some breakKind; SectionPageSetup = geometry } (words "end" 6)) ]
    let movedDoc =
        [ yield! sectionOf "one" 40 a4 "nextPage"
          yield! sectionOf "two" 60 landscape "nextPage"
          yield! sectionOf "three" 50 a4 "nextPage"
          yield noted "moved-note" "footnote" "mf1" (words "noted" 30 + " x") (words "footnote" 12)
          yield! sectionOf "four" 30 a4 "continuous"
          for i in 1 .. 40 -> Doc.Paragraph (para (sprintf "last-p%d" i) (words "last" (15 + (i * 5) % 50))) ]
        |> document
    let state = Layout.createIncrementalLayoutState ()
    let kept (doc: Doc.Document) = Layout.layoutDocumentPagesIncremental state settings measureText measureRange doc |> fst
    compare "sections moved: initial" state movedDoc |> ignore
    let before = kept movedDoc
    let pageCountOf (doc: Doc.Document) = (Layout.layoutDocumentPages settings measureText measureRange doc |> snd).Length
    // A paragraph long enough to add a page to the first section.
    let grown = insertAfter "one-p20" (Doc.Paragraph (para "one-grown" (words "grown" 700))) movedDoc
    checkTrue "sections moved: the edit adds a page" (pageCountOf grown > pageCountOf movedDoc)
    compare "sections moved: after the first section grew by a page" state grown |> ignore
    let laidOut = Layout.incrementalLayoutBlocksLaidOut state
    let after = kept grown
    checkTrue (sprintf "sections moved: only blocks of the edited section were laid out (%d of %d)" laidOut grown.Blocks.Length) (laidOut <= 41)
    let layoutOf id (layouts: Layout.BlockLayout list) = layouts |> List.find (fun layout -> layout.BlockId = id)
    let shift = (layoutOf "three-p10" after).Y - (layoutOf "three-p10" before).Y
    checkTrue "sections moved: a later section stands further down" (shift > 500.0)
    check "sections moved: by the same distance, block for block and line for line"
        ([ for id in [ "two-p3"; "two-p59"; "three-p1"; "three-p49"; "last-p40" ] ->
            let a, b = layoutOf id before, layoutOf id after
            abs (b.Y - a.Y - shift) < 1e-6 && List.forall2 (fun (x: Layout.MeasuredLine) (y: Layout.MeasuredLine) -> abs (y.Y - x.Y - shift) < 1e-6 && x.Segments = y.Segments) a.Lines b.Lines ])
        [ true; true; true; true; true ]
    // The same edit taken back: the sections are where they were.
    let shrunk = removeBlock "one-grown" grown
    compare "sections moved: after the paragraph was removed again" state shrunk |> ignore
    checkTrue (sprintf "sections moved: again only the edited section was laid out (%d)" (Layout.incrementalLayoutBlocksLaidOut state)) (Layout.incrementalLayoutBlocksLaidOut state <= 41)
    checkTrue "sections moved: back where they were" (kept shrunk = before)
    // Typing in a moved section afterwards is an edit of that section alone.
    compare "sections moved: typing in a later section" state (typeIn "three-p7" 3 "typed " grown) |> ignore
    compare "sections moved: and in the first again" state (typeIn "one-p2" 3 "typed " grown) |> ignore
    assertProgressive "sections moved" movedDoc [ "the first section grows", insertAfter "one-p20" (Doc.Paragraph (para "one-grown" (words "grown" 700))) ]

// A change of every paragraph, undone: undo brings back the very paragraphs
// that were laid out before the change, and their lines are still known - no
// line is broken (and no text measured) again. Redo likewise.
do
    let plain = document [ for i in 1 .. 300 -> Doc.Paragraph (para (sprintf "all-%d" i) (words (sprintf "w%d-" i) (20 + i % 40))) ]
    let selected = Doc.selectAll plain
    let bold = Doc.toggleBold selected
    let state = Layout.createIncrementalLayoutState ()
    compare "whole-document format: before" state selected |> ignore
    Layout.resetIncrementalLayoutMeasurements state
    compare "whole-document format: made bold" state bold |> ignore
    check "whole-document format: every paragraph's lines are broken for the change" (Layout.incrementalLayoutMeasuredBlockCount state) 300
    Layout.resetIncrementalLayoutMeasurements state
    compare "whole-document format: undone" state selected |> ignore
    check "whole-document format: undo breaks no line again" (Layout.incrementalLayoutMeasuredBlockCount state) 0
    compare "whole-document format: redone" state bold |> ignore
    check "whole-document format: redo breaks no line again" (Layout.incrementalLayoutMeasuredBlockCount state) 0
    // (What is kept for it goes with the next sweep: typing on does not pile entries up.)
    let mutable current = selected
    for i in 1 .. 700 do
        current <- { current with Blocks = current.Blocks |> List.mapi (fun index block -> if index = i % 300 then Doc.Paragraph (para (sprintf "all-%d" (index + 1)) (words (sprintf "t%d-" i) 12)) else block) }
        Layout.layoutDocumentPagesIncremental state settings measureText measureRange current |> ignore
    compare "whole-document format: after typing on" state current |> ignore
    checkTrue (sprintf "whole-document format: the entries kept do not pile up (%d for 300 paragraphs)" state.Memo.MemoEntries.Count) (state.Memo.MemoEntries.Count <= 2 * 300 + 600)

printfn "incremental-layout: %d passed, %d failed" pass fail
if fail > 0 then failwithf "%d incremental layout checks failed" fail
