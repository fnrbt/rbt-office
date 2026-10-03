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
      LongTrailingSpaceFitSafety = 3.5 }

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

/// Lay the document out cold and incrementally, then again after every edit
/// (the incremental state is carried across edits, as in the editor).
let assertEquivalent label doc (edits: (string * (Doc.Document -> Doc.Document)) list) =
    let state = Layout.createIncrementalLayoutState ()
    let initial = compare (label + " initial") state doc
    let mutable current = doc
    for name, edit in edits do
        current <- edit current
        compare (sprintf "%s after %s" label name) state current |> ignore
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
check "document A numbers its table cells" (cellPrefix "shared-table-r1-a-0" (fst expectedA0), cellPrefix "shared-table-r2-a-0" (fst expectedA0)) ("1.", "5.")
check "document B numbers its table cells" (cellPrefix "shared-table-r1-a-0" (fst expectedB0), cellPrefix "shared-table-r2-a-0" (fst expectedB0)) ("1.", "7.")

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

printfn "incremental-layout: %d passed, %d failed" pass fail
if fail > 0 then failwithf "%d incremental layout checks failed" fail
