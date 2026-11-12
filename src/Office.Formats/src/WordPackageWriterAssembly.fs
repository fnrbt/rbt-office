module DocumentWriter

open System
open XmlGenerator
open WordEditorPackage
open WordPackageWriterModel
open ZipWriter
#if FABLE_COMPILER
open Fable.Core
open Fable.Core.JsInterop
#endif
open WordPackageWriterSplice

type private PackageContext = {
    Model: WriterContext
    mutable SourceBodyXml: string
    mutable SourceChunkStarts: int array
    mutable SourceChunkElementStarts: int array
    mutable SourceChunkEnds: int array
    mutable SourceItemChunks: int array
    mutable SourceStoryCounts: int array
    mutable SourceElements: System.Collections.Generic.Dictionary<int, XmlDsl.XmlNode> option
    mutable WritingBody: bool
    mutable ReportInline: int
    mutable ReportModelBlocks: int
    mutable ReportModelParagraphs: int
    mutable ReportSourceBlocks: int
    ReportReasons: ResizeArray<string> }

let private createPackageContext () =
    { Model = createWriterContext (); SourceBodyXml = ""; SourceChunkStarts = [||]
      SourceChunkElementStarts = [||]; SourceChunkEnds = [||]; SourceItemChunks = [||]
      SourceStoryCounts = [||]; SourceElements = None; WritingBody = false
      ReportInline = 0; ReportModelBlocks = 0; ReportModelParagraphs = 0
      ReportSourceBlocks = 0; ReportReasons = ResizeArray() }

let private noteReason (ctx: PackageContext) (reason: string) =
    if ctx.ReportReasons.Count < 400 then ctx.ReportReasons.Add reason

#if FABLE_COMPILER
type private SaveReport = {
    Reasons: string list; Derived: int; Generated: int; Spliced: int; Patched: int
    Inline: int; ModelBlocks: int; ModelParagraphs: int; SourceBlocks: int }
let mutable private lastCompletedReport =
    { Reasons = []; Derived = 0; Generated = 0; Spliced = 0; Patched = 0
      Inline = 0; ModelBlocks = 0; ModelParagraphs = 0; SourceBlocks = 0 }
let lastSaveReport () : obj =
    let report = lastCompletedReport
    createObj [
        "Reasons" ==> List.toArray report.Reasons
        "DerivedParagraphs" ==> report.Derived
        "GeneratedParagraphs" ==> report.Generated
        "SplicedParagraphs" ==> report.Spliced
        "PatchedParagraphs" ==> report.Patched
        "ModelParagraphsInTables" ==> report.Inline
        "ModelBlocks" ==> report.ModelBlocks
        "ModelBlockParagraphs" ==> report.ModelParagraphs
        "SourceBlocks" ==> report.SourceBlocks ]
#endif

// ---- a paragraph that was split, and paragraphs that were joined ---------------
//
// A paragraph split in two (Enter) is, in the file, two paragraphs made of
// the one that was read: each has the paragraph's properties and the runs,
// bookmarks, fields and links of its part of the text - the run the split is
// in is split. Paragraphs joined (Backspace at the start of the second) are
// the first with the content of the second at its end. Neither is written
// from the model.

/// The paragraphs, tables, rows and cells of a span tree by where each starts.
let private indexOfElements (root: XmlDsl.XmlNode) =
    let index = System.Collections.Generic.Dictionary<int, XmlDsl.XmlNode>()
    for node in XmlPatch.descendantsWhere (fun node -> List.contains (XmlPatch.localName node) [ "p"; "tbl"; "tr"; "tc" ]) root do
        if XmlDsl.Span.isKnown (XmlPatch.span node) then index.[(XmlPatch.span node).Start] <- node
    index
/// Origins name elements of the body part: they are looked up only while the body is written.
let private sourceElementAt (ctx: PackageContext) (at: int) : XmlDsl.XmlNode option =
    if not ctx.WritingBody || ctx.SourceBodyXml = "" || at < 0 then None
    else
        let index =
            match ctx.SourceElements with
            | Some index -> index
            | None ->
                let index = XmlDsl.tryScan ctx.SourceBodyXml (System.Collections.Generic.HashSet<string>()) |> Option.map indexOfElements |> Option.defaultValue (System.Collections.Generic.Dictionary())
                ctx.SourceElements <- Some index
                index
        match index.TryGetValue at with
        | true, node -> Some node
        | _ -> None
/// The element a paragraph of the model is of (its SourceAt), as it was read.
let private sourceParagraphOf (ctx: PackageContext) (para: EditorParagraph) : XmlDsl.XmlNode option =
    let at = match para.Source.Origin with Some (ReadAt at | LikeAt at) -> Some at | None -> None
    match at |> Option.bind (sourceElementAt ctx) with
    | Some (XmlDsl.Element (name, _, _, _) as node) when name = "w:p" -> Some node
    | _ -> None
/// Whether a paragraph as it was read is tied to what is around it, so that it cannot be written on its
/// own: a shape is anchored in it (written with its block), or a field begins in it and ends in another.
let private paragraphIsBound (node: XmlDsl.XmlNode) : bool =
    let marks = XmlPatch.descendants "fldChar" node |> List.choose (fun mark -> XmlPatch.attribute mark "w:fldCharType")
    holdsShape node || (marks |> List.filter ((=) "begin")).Length <> (marks |> List.filter ((=) "end")).Length
/// What is known of the part being written as it was read (see planBody): a header's plan is made with
/// the body's, and each is put back while its part is written.
type private SourceState = {
    Xml: string; Elements: System.Collections.Generic.Dictionary<int, XmlDsl.XmlNode> option
    ChunkStarts: int array; ChunkElementStarts: int array; ChunkEnds: int array; ItemChunks: int array; StoryCounts: int array }
let private sourceState (ctx: PackageContext) () : SourceState =
    { Xml = ctx.SourceBodyXml; Elements = ctx.SourceElements; ChunkStarts = ctx.SourceChunkStarts; ChunkElementStarts = ctx.SourceChunkElementStarts
      ChunkEnds = ctx.SourceChunkEnds; ItemChunks = ctx.SourceItemChunks; StoryCounts = ctx.SourceStoryCounts }
let private useSourceState (ctx: PackageContext) (state: SourceState) =
    ctx.SourceBodyXml <- state.Xml
    ctx.SourceElements <- state.Elements
    ctx.SourceChunkStarts <- state.ChunkStarts
    ctx.SourceChunkElementStarts <- state.ChunkElementStarts
    ctx.SourceChunkEnds <- state.ChunkEnds
    ctx.SourceItemChunks <- state.ItemChunks
    ctx.SourceStoryCounts <- state.StoryCounts
let private forgetSourceBody (ctx: PackageContext) () =
    ctx.SourceElements <- None
    ctx.SourceBodyXml <- ""
    ctx.SourceChunkStarts <- [||]
    ctx.SourceChunkElementStarts <- [||]
    ctx.SourceChunkEnds <- [||]
    ctx.SourceItemChunks <- [||]
    ctx.SourceStoryCounts <- [||]
/// The w:tbl of an element of a body: the table itself, or the one element of a content control or custom XML that is.
let rec private tableElementIn (node: XmlDsl.XmlNode) : XmlDsl.XmlNode option =
    let kidsOf (node: XmlDsl.XmlNode) = XmlPatch.elements node
    match XmlPatch.localName node with
    | "tbl" -> Some node
    | "sdt" ->
        (match kidsOf node |> List.filter (fun kid -> XmlPatch.localName kid = "sdtContent") with
         | [ content ] -> (match kidsOf content with [ one ] -> tableElementIn one | _ -> None)
         | _ -> None)
    | "customXml" -> (match kidsOf node |> List.filter (fun kid -> XmlPatch.localName kid <> "customXmlPr") with [ one ] -> tableElementIn one | _ -> None)
    | _ -> None

/// A w:tbl of `xml` taken apart (see TableSource), with what stands around it in its block (`wrapStart`, `wrapEnd`).
let private sourceTableOf (xml: string) (tableNode: XmlDsl.XmlNode) (wrapStart: string) (wrapEnd: string) : SourceTableText option =
    let local (node: XmlDsl.XmlNode) = match node with XmlDsl.Element (name, _, _, _) -> XmlDsl.Ns.stripNs name | _ -> ""
    let spanOf (node: XmlDsl.XmlNode) = match node with XmlDsl.Element (_, _, _, sp) -> sp | _ -> XmlDsl.Span.none
    let kidsOf (node: XmlDsl.XmlNode) = match node with XmlDsl.Element (_, _, kids, _) -> kids |> List.filter (fun kid -> local kid <> "") | _ -> []
    let slice = XmlPatch.slice xml
    match tableNode with
    | XmlDsl.Element (name, _, children, span) when XmlDsl.Span.isKnown span && name = "w:tbl" ->
        let rows = System.Collections.Generic.Dictionary<int, SourceRowText>()
        let cells = System.Collections.Generic.Dictionary<int, SourceCellText>()
        let mutable sound = true
        let partOf (nodes: XmlDsl.XmlNode list) (name: string) =
            nodes |> List.tryFind (fun kid -> local kid = name) |> Option.map (XmlPatch.bytesOf xml) |> Option.defaultValue ""
        // Where the leading children named `names` of an element end (its start tag's end when it has none).
        let headEndOf (sp: XmlDsl.Span) (nodes: XmlDsl.XmlNode list) (names: string list) =
            kidsOf (XmlDsl.Element ("", Map.empty, nodes, XmlDsl.Span.none))
            |> List.takeWhile (fun kid -> List.contains (local kid) names)
            |> List.fold (fun at kid -> let kidSpan = spanOf kid in if XmlDsl.Span.isKnown kidSpan then kidSpan.End else at) sp.StartTagEnd
        let laid (isItem: string -> bool) (take: XmlDsl.XmlNode -> unit) (node: XmlDsl.XmlNode) (from: int) =
            match XmlPatch.layout (XmlPatch.localName >> isItem) (fun _ -> false) (fun _ -> false) take xml node from with
            | Some found -> found
            | None ->
                sound <- false
                { XmlPatch.Layout.Text = xml; Wrappers = System.Collections.Generic.Dictionary(); Children = []; Leads = System.Collections.Generic.Dictionary(); KeptEnds = System.Collections.Generic.Dictionary(); Tail = (0, 0) }
        let copyTag (node: XmlDsl.XmlNode) = XmlPatch.startTag xml node [ XmlPatch.RemoveAttribute "w14:paraId"; XmlPatch.RemoveAttribute "w14:textId" ]
        let takeCell (node: XmlDsl.XmlNode) =
            let sp = spanOf node
            if sp.StartTagEnd >= sp.End then sound <- false
            else
                let kids = match node with XmlDsl.Element (_, _, kids, _) -> kids | _ -> []
                let headEnd = headEndOf sp kids [ "tcPr" ]
                cells.[sp.Start] <- { CellXml = XmlPatch.bytesOf xml node
                                      CellStartTag = slice sp.Start sp.StartTagEnd
                                      CopyStartTag = copyTag node
                                      CellProps = partOf kids "tcPr"
                                      CellHead = slice sp.Start headEnd
                                      CellChildren = laid (fun name -> name = "p" || name = "tbl") ignore node headEnd }
        let takeRow (node: XmlDsl.XmlNode) =
            let sp = spanOf node
            if sp.StartTagEnd >= sp.End then sound <- false
            else
                let kids = match node with XmlDsl.Element (_, _, kids, _) -> kids | _ -> []
                let headEnd = headEndOf sp kids [ "tblPrEx"; "trPr" ]
                let inner = laid (fun name -> name = "tc") takeCell node headEnd
                rows.[sp.Start] <- { RowXml = XmlPatch.bytesOf xml node
                                     RowStartTag = slice sp.Start sp.StartTagEnd
                                     CopyStartTag = copyTag node
                                     RowExceptions = partOf kids "tblPrEx"
                                     RowProps = partOf kids "trPr"
                                     RowHead = slice sp.Start headEnd
                                     RowChildren = inner }
        let headEnd = headEndOf span children [ "tblPr"; "tblGrid" ]
        let tableChildren = laid (fun name -> name = "tr") takeRow tableNode headEnd
        if not sound then None
        else Some { WrapStart = wrapStart; WrapEnd = wrapEnd
                    TableHead = slice span.Start headEnd; TableStartTag = slice span.Start span.StartTagEnd; TableProps = partOf children "tblPr"; TableGrid = partOf children "tblGrid"
                    TableRows = rows; TableCells = cells; TableChildren = tableChildren; TableNode = tableNode }
    | _ -> None

// The table a block came from, taken apart (see TableSource): a table whose element is the block's alone,
// or the one element in a content control or custom XML that is.
let private configureSources (ctx: PackageContext) =
    ctx.Model.ParagraphSource <-
        Some { new ParagraphSource with
            // (Nothing was done to it: the bytes it was read as.)
            member _.Of (para: EditorParagraph) =
                if not (para.SourceParagraph |> Option.exists (fun source -> source.Verbatim)) then None
                else sourceParagraphOf ctx para |> Option.filter (paragraphIsBound >> not) |> Option.map (XmlPatch.bytesOf ctx.SourceBodyXml)
            member _.SpanOf (para: EditorParagraph) =
                sourceParagraphOf ctx para
                |> Option.filter (paragraphIsBound >> not)
                |> Option.map (fun node -> match node with XmlDsl.Element (_, _, _, span) -> span.Start, span.End | _ -> -1, -1) }

    ctx.Model.ParagraphOrigin <-
        Some { new ParagraphOrigin with
            member _.Of (para: EditorParagraph) = sourceParagraphOf ctx para |> Option.map (XmlPatch.bytesOf ctx.SourceBodyXml) }

    ctx.Model.TableSource <-
        Some { new TableSource with
            member _.Of (spec: WordEditorPackage.TableSource) =
                let block = spec.Block |> Option.defaultValue -1
                let chunk = if block >= 0 && block < ctx.SourceItemChunks.Length then ctx.SourceItemChunks.[block] else -1
                if chunk < 0 || chunk >= ctx.SourceChunkStarts.Length || chunk >= ctx.SourceChunkElementStarts.Length || (ctx.SourceItemChunks |> Array.filter (fun other -> other = chunk)).Length <> 1 then None
                else
                    let xml = ctx.SourceBodyXml
                    let elementStart = ctx.SourceChunkElementStarts.[chunk]
                    // (What stands before the block's element - white space alone here: marks there are the plan's.)
                    let lead = if elementStart >= ctx.SourceChunkStarts.[chunk] then XmlPatch.slice xml ctx.SourceChunkStarts.[chunk] elementStart else "x"
                    if lead.Trim() <> "" || elementStart >= ctx.SourceChunkEnds.[chunk] then None
                    else
                        match XmlDsl.tryScanAt xml elementStart ctx.SourceChunkEnds.[chunk] (System.Collections.Generic.HashSet<string>()) |> Option.bind tableElementIn with
                        | Some tableNode when XmlDsl.Span.isKnown (XmlPatch.span tableNode) ->
                            let span = XmlPatch.span tableNode
                            sourceTableOf xml tableNode (XmlPatch.slice xml ctx.SourceChunkStarts.[chunk] span.Start) (XmlPatch.slice xml span.End ctx.SourceChunkEnds.[chunk])
                        | _ -> None
            member _.At (tableNode: XmlDsl.XmlNode) = sourceTableOf ctx.SourceBodyXml tableNode "" ""
            member _.Merge (container: string, source: string, was: string list, now: string list) =
                mergedProperties container source was now }

    WordPackageWriterSplice.configure ctx.Model

// ---- a table whose rows or columns came or went --------------------------------
//
// A table that was read and changed in more than its paragraphs - a row or a column put in or taken out, a
// cell's properties - is its element as it was read with those changes made in it (XmlPatch): a row or a cell
// that is gone is taken out, the marks of ranges in it (bookmarks, comment and permission ranges) standing
// where it stood; one that is new is put in next to the one before it (made like the one it was made like:
// see generateTableRow, generateTableCell); one that changed is written from its source; the grid has the
// columns put in or taken out; every other row and cell is the bytes it was read as.

/// The marks of ranges (bookmarks, comment and permission ranges) in an element that is taken out: they stay
/// where it stood. (Not the range of a comment the document has no more.)
let private marksOutOf (ctx: PackageContext) (xml: string) (node: XmlDsl.XmlNode) : string =
    rangeMarksIn carriedRangeKinds node
    |> List.filter (fun mark -> mark.Kind <> "commentRange" || Set.contains mark.Id ctx.Model.LiveCommentIds)
    |> List.map (fun mark -> XmlPatch.bytesOf xml mark.Node)
    |> String.concat ""

/// The edit that takes an element out: the marks of ranges in it in its place, or nothing.
let private takenOut (ctx: PackageContext) (xml: string) (node: XmlDsl.XmlNode) : XmlPatch.ElementEdit =
    match marksOutOf ctx xml node with
    | "" -> XmlPatch.Removed
    | marks -> XmlPatch.Replaced marks

/// The items of a layout (where each starts) and the wrappers each is in.
let private layoutItems (laid: XmlPatch.Layout) : (int * int list) list =
    laid.Children |> List.choose (function XmlPatch.LaidItem (at, wrappers) -> Some (at, wrappers) | XmlPatch.LaidKept _ -> None)

/// The rows of a table of the model that are rows of the table as it was read (not copies): where each starts.
let private sourceRowsOf (known: SourceTableText) (table: EditorTable) : int list =
    table.Rows
    |> Array.toList
    |> List.choose (fun row -> match sourceOfRow row.SourceRow with Some (at, false, _) when known.TableRows.ContainsKey at -> Some at | _ -> None)

/// Whether a table of the model can be written as its element as read with its changes made in it: some of its
/// rows are rows that were read, in the order they were read in.
let private tableIsPatchable (known: SourceTableText) (table: EditorTable) : bool =
    let rows = sourceRowsOf known table
    not rows.IsEmpty && (rows |> List.pairwise |> List.forall (fun (a, b) -> a < b))

/// What is taken out of a layout (of rows, of cells) where the items `kept` stay: each item that does not stay, or the
/// outermost of its wrappers that holds none that stays (a content control around rows: a repeating section's item) -
/// where each starts. (`known`: the wrappers that may be taken out - a content control, custom XML.)
let private takenOutOf (laid: XmlPatch.Layout) (known: int -> bool) (kept: int -> bool) : int list =
    let items = layoutItems laid
    items
    |> List.filter (fun (at, _) -> not (kept at))
    |> List.map (fun (at, wrappers) ->
        wrappers
        |> List.tryFind (fun wrapper -> known wrapper && items |> List.forall (fun (other, around) -> not (List.contains wrapper around) || not (kept other)))
        |> Option.defaultValue at)
    |> List.distinct

/// The table written as its element as it was read with the model's changes made in it (see tableIsPatchable).
let private patchedTableXml (ctx: PackageContext) (hyperlinks: Map<string, string>) (images: ImageScope) (known: SourceTableText) (table: EditorTable) : string =
    let xml = ctx.SourceBodyXml
    let tableNode = known.TableNode
    // (The rows and cells of the table and the content controls and custom XML around them, by where each starts.)
    let nodes = System.Collections.Generic.Dictionary<int, XmlDsl.XmlNode>()
    for node in XmlPatch.descendantsWhere (fun node -> List.contains (XmlPatch.localName node) [ "tr"; "tc"; "sdt"; "customXml" ]) tableNode do
        nodes.[(XmlPatch.span node).Start] <- node
    let targets = System.Collections.Generic.Dictionary<int, XmlDsl.XmlNode * ResizeArray<XmlPatch.ElementEdit>>()
    let edit (node: XmlDsl.XmlNode) (change: XmlPatch.ElementEdit) =
        let at = (XmlPatch.span node).Start
        match targets.TryGetValue at with
        | true, (_, edits) -> edits.Add change
        | _ -> targets.[at] <- (node, ResizeArray [ change ])
    // (A child of an element by its local name, with the name it has there.)
    let childNamed (node: XmlDsl.XmlNode) (local: string) = XmlPatch.elements node |> List.tryFind (fun kid -> XmlPatch.localName kid = local)
    let qualified (node: XmlDsl.XmlNode) (local: string) = childNamed node local |> Option.map XmlPatch.name |> Option.defaultValue ("w:" + local)
    // A properties element that is now `now` where it was `was` (as read): set, taken out, or left.
    let setProperties (node: XmlDsl.XmlNode) (local: string) (was: string) (now: string) : XmlPatch.ElementEdit list =
        if now = was then []
        elif now = "" then [ XmlPatch.RemoveChild (qualified node local) ]
        else [ XmlPatch.SetChild (qualified node local, now) ]
    // Restore the enclosing table context after rendering this table.
    let mutable outerContext = ctx.Model.ParagraphsInContext
    let mutable outerTable = ctx.Model.CurrentSourceTable
    ctx.Model.ParagraphsInContext <- true
    ctx.Model.CurrentSourceTable <- Some known
    // The table's properties: those it was read with, changed where the model's changed.
    match table.SourceTable |> Option.bind (fun source -> source.Before) with
    | Some before ->
        let now = tablePropertyElements table
        match ctx.Model.TableSource.Value.Merge ("w:tblPr", known.TableProps, tablePropertyElements before, now) with
        | Some props ->
            for change in setProperties tableNode "tblPr" known.TableProps (if props = "" then "<w:tblPr/>" else props) do edit tableNode change
        | None -> edit tableNode (XmlPatch.SetChild (qualified tableNode "tblPr", element "w:tblPr" [] now))
    | None -> ()
    // The grid: a column put in like its neighbour, one taken out; otherwise the model's widths.
    if not (sourceGridIs known.TableGrid table) then
        let now = table.GridColumns |> Array.map roundToInt
        match childNamed tableNode "tblGrid" with
        | None -> edit tableNode (XmlPatch.SetChild ("w:tblGrid", generateTableGrid table))
        | Some grid ->
            let columns = XmlPatch.elements grid |> List.filter (fun node -> XmlPatch.localName node = "gridCol") |> List.toArray
            let widths = columns |> Array.map (fun column -> XmlPatch.attribute column "w:w" |> Option.bind (fun value -> match System.Double.TryParse value with | true, width -> Some (roundToInt width) | _ -> None))
            let near (a: int) (b: int) = abs (a - b) <= 1
            let read = widths |> Array.choose id
            // (The place k at which `longer` has a column `shorter` has not.)
            let extraAt (longer: int array) (shorter: int array) =
                [ 0 .. shorter.Length ]
                |> List.tryFind (fun k -> [ 0 .. shorter.Length - 1 ] |> List.forall (fun i -> near shorter.[i] longer.[if i < k then i else i + 1]))
            let width (column: XmlDsl.XmlNode) (value: int) =
                XmlPatch.apply XmlPatch.Word.order { Text = xml; Node = column } [ XmlPatch.SetAttribute ("w:w", string value) ]
            let whole () = edit grid (XmlPatch.Replaced (generateTableGrid table))
            if read.Length <> widths.Length || columns.Length = 0 || now.Length = 0 then whole ()
            elif now.Length = read.Length + 1 then
                match extraAt now read with
                | Some k when k > 0 -> edit columns.[k - 1] (XmlPatch.InsertAfter (if near read.[k - 1] now.[k] then XmlPatch.bytesOf xml columns.[k - 1] else width columns.[k - 1] now.[k]))
                | Some _ -> edit columns.[0] (XmlPatch.InsertBefore (if near read.[0] now.[0] then XmlPatch.bytesOf xml columns.[0] else width columns.[0] now.[0]))
                | None -> whole ()
            elif now.Length = read.Length - 1 then
                match extraAt read now with
                | Some k -> edit columns.[k] XmlPatch.Removed
                | None -> whole ()
            elif now.Length = read.Length then
                for i in 0 .. read.Length - 1 do
                    if not (near read.[i] now.[i]) then edit columns.[i] (XmlPatch.SetAttribute ("w:w", string now.[i]))
            else whole ()
    // The items of a row or of the table in the model's order - each one that was read (where it starts) or a new
    // one (its XML, and where the one it was made like starts) -: the new ones put in among those that were read.
    // A run of new ones goes before the item that follows it where each was made like that one (a row put in
    // above it), else after the item before it (in the wrappers of that one), else before the first.
    let placeNew (entries: Choice<int, string * int option> list) =
        let rec walk (previous: int option) (entries: Choice<int, string * int option> list) =
            match entries with
            | [] -> ()
            | Choice1Of2 at :: rest -> walk (Some at) rest
            | Choice2Of2 _ :: _ ->
                let run = entries |> List.takeWhile (function Choice2Of2 _ -> true | Choice1Of2 _ -> false) |> List.choose (function Choice2Of2 entry -> Some entry | _ -> None)
                let rest = entries |> List.skip run.Length
                let next = match rest with Choice1Of2 at :: _ -> Some at | _ -> None
                let text = run |> List.map fst |> String.concat ""
                match previous, next with
                | None, Some at -> edit nodes.[at] (XmlPatch.InsertBefore text)
                | Some _, Some at when run |> List.forall (fun (_, like) -> like = Some at) -> edit nodes.[at] (XmlPatch.InsertBefore text)
                | Some at, _ -> edit nodes.[at] (XmlPatch.InsertAfter text)
                | None, None -> ()
                walk previous rest
        walk None entries
    // The rows.
    let keptRows = Set.ofList (sourceRowsOf known table)
    for at in takenOutOf known.TableChildren nodes.ContainsKey keptRows.Contains do
        match nodes.TryGetValue at with
        | true, node -> edit node (takenOut ctx xml node)
        | _ -> ()
    [ for row in table.Rows do
        match sourceOfRow row.SourceRow with
        | Some (at, false, same) when keptRows.Contains at && nodes.ContainsKey at ->
            let rowNode = nodes.[at]
            if not same then
                let knownRow = known.TableRows.[at]
                // Its properties, changed where the model's changed.
                let head =
                    match row.SourceRow |> Option.bind (fun source -> source.Before) with
                    | None -> Some []
                    | Some rowBefore ->
                        let exceptions, props = tableRowProperties row
                        let exceptionsBefore, propsBefore = tableRowProperties rowBefore
                        ctx.Model.TableSource.Value.Merge ("w:trPr", knownRow.RowProps, propsBefore, props)
                        |> Option.map (fun merged ->
                            setProperties rowNode "trPr" knownRow.RowProps merged
                            @ (if exceptionsBefore = exceptions then [] else setProperties rowNode "tblPrEx" knownRow.RowExceptions exceptions))
                // Its cells: those that were read in the order they were read in, each where it stood.
                let cells = row.Cells
                let keptCells =
                    cells |> Array.toList |> List.choose (fun cell ->
                        match sourceOfCell cell.SourceCell with
                        | Some (cellAt, false, _) when XmlPatch.isItemOf knownRow.RowChildren cellAt && nodes.ContainsKey cellAt -> Some cellAt
                        | _ -> None)
                let cellsInPlace = not keptCells.IsEmpty && (keptCells |> List.pairwise |> List.forall (fun (a, b) -> a < b))
                match head with
                | Some headEdits when cellsInPlace ->
                    for change in headEdits do edit rowNode change
                    for cellAt in takenOutOf knownRow.RowChildren nodes.ContainsKey (fun cellAt -> List.contains cellAt keptCells) do
                        match nodes.TryGetValue cellAt with
                        | true, node -> edit node (takenOut ctx xml node)
                        | _ -> ()
                    placeNew
                        [ for cell in cells do
                            match sourceOfCell cell.SourceCell with
                            | Some (cellAt, false, cellSame) when List.contains cellAt keptCells ->
                                if not cellSame then edit nodes.[cellAt] (XmlPatch.Replaced (generateTableCell ctx.Model hyperlinks images cell))
                                yield Choice1Of2 cellAt
                            | _ -> yield Choice2Of2 (generateTableCell ctx.Model hyperlinks images cell, (sourceOfCell cell.SourceCell |> Option.map (fun (at, _, _) -> at))) ]
                | _ -> edit rowNode (XmlPatch.Replaced (generateTableRow ctx.Model hyperlinks images row))
            yield Choice1Of2 at
        | _ ->
            // A row that is new.
            yield Choice2Of2 (generateTableRow ctx.Model hyperlinks images row, (sourceOfRow row.SourceRow |> Option.map (fun (at, _, _) -> at))) ]
    |> placeNew
    ctx.Model.ParagraphsInContext <- outerContext
    ctx.Model.CurrentSourceTable <- outerTable
    let edits = XmlPatch.editsAt tableNode [ for KeyValue (_, (node, changes)) in targets -> node, List.ofSeq changes ]
    XmlPatch.apply XmlPatch.Word.order { Text = xml; Node = tableNode } edits

/// The paragraphs of the items of an element of the body (from `start` to `finish` in the part) that were
/// changed, to be written from their source (see splicedRuns) in the place of their elements (their
/// SourceAt): where each element starts and ends, and the paragraph. None when one of them cannot be: its
/// element is not known, is not in this one, or is bound to what is around it (a shape anchored in it, a
/// field that begins or ends in another), or it cannot be spliced - in a paragraph of the body itself the
/// whole element is then written from the model; in a table, the table.
let private blockChanges (item: EditorBlock) =
    match item with Paragraph p -> p.SourceTextChanges | Table t -> t.SourceTextChanges | TextBox t -> t.SourceTextChanges
let private blockIndex (item: EditorBlock) =
    match item with Paragraph p -> p.SourceIndex | Table t -> t.SourceIndex | TextBox t -> t.SourceIndex
let private blockUnchanged (item: EditorBlock) =
    match item with Paragraph p -> p.SourceUnchanged | Table t -> t.SourceUnchanged | TextBox t -> t.SourceUnchanged
let private blockWhy (item: EditorBlock) =
    match item with Paragraph p -> p.SourceWhy | Table t -> t.SourceWhy | TextBox t -> t.SourceWhy

let private patchSourceChunk (ctx: PackageContext) (start: int) (finish: int) (items: EditorBlock array) : (int * int * EditorParagraph) list option =
    let inlined = ResizeArray<int * int * EditorParagraph>()
    let mutable ok = true
    for item in items do
        match blockChanges item with
        | Some changes when ok ->
            for change in changes.Edits do
                if ok then
                    match change.After |> Option.bind (sourceParagraphOf ctx), change.After with
                    | Some (XmlDsl.Element (_, _, _, span) as node), Some after when span.Start >= start && span.End <= finish ->
                        let otherShape = XmlPatch.descendantsWhere (fun inner -> XmlPatch.name inner = "w:r" && holdsShape inner && not (holdsShapeItem inner)) node
                        if not otherShape.IsEmpty then
                            noteReason ctx (sprintf "paragraph %d of the block: a shape the model does not hold as an item is in it: the block is written from the model" change.Paragraph)
                            ok <- false
                        elif after.SpliceOriginal.IsNone ||
                             (after.SpliceOriginal |> Option.exists (fun original ->
                                 let was, now = original.Formatting, after.Formatting
                                 was.ContentControlId <> now.ContentControlId || was.ContentControlTag <> now.ContentControlTag ||
                                 was.ContentControlAlias <> now.ContentControlAlias || was.ContentControlXml <> now.ContentControlXml ||
                                 was.ContentControlGroup <> now.ContentControlGroup || was.ContentControlAncestors <> now.ContentControlAncestors)) then
                            noteReason ctx (sprintf "paragraph %d of the block: %s" change.Paragraph (change.Why |> Option.defaultValue "changed, and not of a paragraph that was read"))
                            ok <- false
                        else
                            let updated = { after with SpliceFrom = Some (span.Start, span.End) } |> setSpliceCut change
                            ctx.Model.CarriedSources.[updated] <- XmlPatch.bytesOf ctx.SourceBodyXml node
                            inlined.Add (span.Start, span.End, updated)
                    | _ ->
                        noteReason ctx (sprintf "paragraph %d of the block: its element as it was read is not known" change.Paragraph)
                        ok <- false
        | _ -> ()
    let ordered = inlined |> List.ofSeq |> List.distinctBy (fun (start, finish, _) -> start, finish) |> List.sortBy (fun (start, _, _) -> start)
    if ordered.Length < inlined.Count then noteReason ctx "two paragraphs of the block are of one element as it was read: the first is written in its place"
    let overlapping = ordered |> List.pairwise |> List.exists (fun ((_, firstEnd, _), (nextStart, _, _)) -> nextStart < firstEnd)
    if overlapping then noteReason ctx "paragraphs of the block are of elements one inside the other: the block is written from the model"
    if ok && not overlapping then Some ordered else None

type private BodySegment =
    /// The part's text from `start` to `finish`, as it was read but for the paragraphs in it written from
    /// their source: in the place of each element (where it starts and ends), its paragraph.
    | SourceText of start: int * finish: int * paragraphs: (int * int * EditorParagraph) list
    | ModelBlocks of EditorBlock array
    /// The part's text from `start` to `finish` (a block that is a table, and what stands before it) as it was read
    /// but for the table, which is its element as it was read with the model's changes made in it (patchedTableXml).
    | PatchedTable of start: int * finish: int * table: SourceTableText * item: EditorTable
    /// The marks of ranges that stood between two blocks, before a block that is written from
    /// the model (or is not there any more): they stay where they stood.
    | CarriedMarks of RangeMark list

type private BodyPlan = {
    /// The part as it was read.
    Original: string
    /// Its root element (w:document).
    Root: XmlDsl.XmlNode
    /// The element that holds the blocks (w:body; the root of a header; a w:footnote, w:endnote, w:comment).
    Container: XmlDsl.XmlNode
    /// Where what w:body holds starts and ends.
    ContentStart: int
    ContentEnd: int
    Segments: BodySegment list
    /// What stands between the last block and the body's section properties (where it starts and ends).
    BeforeSection: int * int
    /// The body's own w:sectPr as it was read (where it is), when the page setup is what it was.
    Section: (int * int) option
    /// Source section even when its modeled properties changed.
    OriginalSection: (int * int) option
    AfterSection: int * int
    /// The items written from the model.
    Regenerated: EditorBlock array
}

/// Where the blocks of a part are (see planBody).
type private BlockContainer =
    /// The document's: in w:body, with its section properties.
    | DocumentBody
    /// A header's or a footer's: in the part's root.
    | PartRoot
    /// A note's or a comment's: in its element (w:footnote, w:endnote, w:comment) of this w:id.
    | StoryElement of name: string * id: string

let private sourceIndexOf item = blockIndex item |> Option.defaultValue -1

type private ElementRanges = { Starts: int array; Ends: int array; Nodes: XmlDsl.XmlNode array; CloseAt: int }

/// Which of the body's elements go back as they were read. None when the
/// text cannot be matched with what the reader made of it (the body is then
/// written from the model, as a new document is), or when nothing is as it was.
/// `container`: where the blocks are - the document's in w:body; a header's or a footer's in its root, a
/// note's or a comment's in its element (neither has section properties of its own).
let private planBody (ctx: PackageContext) (container: BlockContainer) (xml: string) (itemChunks: int array) (storyCounts: int array) (chunkCount: int) (items: EditorBlock array) (pageSetupUnchanged: bool) : BodyPlan option =
    let isBody = (container = DocumentBody)
    // The part's root, its w:body, and the elements the body holds (the blocks, as the reader counted them).
    let bodyOfRoot =
        XmlPatch.readPart xml
        |> Option.bind (fun part ->
            (match container with
             | DocumentBody -> XmlPatch.elements part.Node |> List.tryFind (fun node -> XmlPatch.localName node = "body")
             | PartRoot -> Some part.Node
             | StoryElement (name, id) -> XmlPatch.childrenNamed part name |> List.tryFind (fun node -> XmlPatch.attribute node "w:id" = Some id))
            |> Option.filter (fun body -> (XmlPatch.span body).StartTagEnd < (XmlPatch.span body).End)
            |> Option.map (fun body ->
                let kids = XmlPatch.elements body |> List.toArray
                part.Node, body, { Starts = kids |> Array.map (fun kid -> (XmlPatch.span kid).Start)
                                   Ends = kids |> Array.map (fun kid -> (XmlPatch.span kid).End)
                                   Nodes = kids
                                   CloseAt = (XmlPatch.span body).EndTagStart }))
    match bodyOfRoot with
    | Some (root, bodyNode, body) when body.Starts.Length = chunkCount && chunkCount > 0 ->
        let contentStart = (XmlPatch.span bodyNode).StartTagEnd
        let contentEnd = body.CloseAt
        let isSection index = isBody && XmlPatch.localName body.Nodes.[index] = "sectPr"
        let sectIndex = if isSection (chunkCount - 1) then chunkCount - 1 else -1
        let strayFinalSection = [ 0 .. chunkCount - 2 ] |> List.exists isSection
        // How many items each element gave when it was read.
        let originalCount = Array.zeroCreate<int> chunkCount
        let mutable mapped = true
        for chunk in itemChunks do
            if chunk >= chunkCount then mapped <- false
            elif chunk >= 0 then originalCount.[chunk] <- originalCount.[chunk] + 1
        if strayFinalSection || not mapped || (sectIndex >= 0 && originalCount.[sectIndex] > 0) then None
        else
            // An element's text starts where the block element before it ends: what
            // stands between two blocks (a bookmark end, white space) goes with the second.
            let textStart = Array.zeroCreate<int> chunkCount
            let mutable previousEnd = contentStart
            for chunk in 0 .. chunkCount - 1 do
                textStart.[chunk] <- previousEnd
                if originalCount.[chunk] > 0 then previousEnd <- body.Ends.[chunk]
            let tailStart = previousEnd
            // (The paragraphs that were read, for those that are written of them: see sourceParagraphXml.)
            ctx.SourceBodyXml <- xml
            ctx.SourceElements <- Some (indexOfElements root)
            ctx.SourceChunkStarts <- textStart
            ctx.SourceChunkElementStarts <- Array.copy body.Starts
            ctx.SourceChunkEnds <- Array.init chunkCount (fun chunk -> body.Ends.[chunk])
            ctx.SourceItemChunks <- itemChunks
            ctx.SourceStoryCounts <- storyCounts
                    // The items there are now, by the element they came from.
            let chunkOfItem =
                items
                |> Array.map (fun item ->
                    let index = sourceIndexOf item
                    if index >= 0 && index < itemChunks.Length then itemChunks.[index] else -1)
            let isTextBox item = match item with TextBox _ -> true | _ -> false
            let keep = Array.init chunkCount (fun chunk -> originalCount.[chunk] > 0)
            let firstAt = Array.create chunkCount (-1)
            let lastAt = Array.create chunkCount (-1)
            let found = Array.zeroCreate<int> chunkCount
            let lastSource = Array.create chunkCount (-1)
            let patched = Array.zeroCreate<bool> chunkCount
            let patchedText = Array.create<(int * int * EditorParagraph) list option> chunkCount None
            let mutable anchorChunk = -1
            items
            |> Array.iteri (fun position item ->
                let chunk = chunkOfItem.[position]
                if chunk >= 0 then
                    if firstAt.[chunk] < 0 then firstAt.[chunk] <- position
                    // (All of an element's items stand together, in the order they were read in.)
                    if lastAt.[chunk] >= 0 && lastAt.[chunk] <> position - 1 then keep.[chunk] <- false
                    let index = sourceIndexOf item
                    if index <= lastSource.[chunk] then keep.[chunk] <- false
                    lastSource.[chunk] <- index
                    lastAt.[chunk] <- position
                    found.[chunk] <- found.[chunk] + 1
                    if not (blockUnchanged item) then
                        // (Changed in its text alone: the change is made in the element's XML.)
                        if (blockChanges item).IsNone then
                            if keep.[chunk] then noteReason ctx (sprintf "block %d: %s" position (blockWhy item |> Option.defaultValue "changed"))
                            keep.[chunk] <- false
                        else patched.[chunk] <- true
                if isTextBox item then
                    // A shape is written into the block it is anchored in: one that is new,
                    // or that now follows another block, makes that block's XML.
                    if chunk <> anchorChunk then
                        if chunk >= 0 then keep.[chunk] <- false
                        if anchorChunk >= 0 then keep.[anchorChunk] <- false
                else anchorChunk <- chunk)
            ctx.WritingBody <- true
            for chunk in 0 .. chunkCount - 1 do
                if found.[chunk] <> originalCount.[chunk] then keep.[chunk] <- false
                if keep.[chunk] && patched.[chunk] then
                    patchedText.[chunk] <- patchSourceChunk ctx textStart.[chunk] body.Ends.[chunk] (Array.sub items firstAt.[chunk] originalCount.[chunk])
                    if patchedText.[chunk].IsNone then keep.[chunk] <- false
            ctx.WritingBody <- false
            // A field that begins in one element and ends in a later one (a table of
            // contents): those elements go back together or not at all.
            let blockChunks = [| for chunk in 0 .. chunkCount - 1 do if originalCount.[chunk] > 0 then yield chunk |]
            let mutable depth = 0
            let group = ResizeArray<int>()
            // An element in which every field that ends also begins holds none of the marks of
            // the field around it (an entry of a table of contents): it can be written from
            // the model while the others stay the text they were.
            // (The marks of the fields of an element's text - begin, separate, end - in order.)
            let fieldMarks =
                Array.init chunkCount (fun chunk ->
                    body.Nodes
                    |> Array.filter (fun node -> (XmlPatch.span node).Start >= textStart.[chunk] && (XmlPatch.span node).End <= body.Ends.[chunk])
                    |> Array.toList
                    |> List.collect (XmlPatch.descendants "fldChar")
                    |> List.choose (fun mark -> XmlPatch.attribute mark "w:fldCharType"))
            let neutral (chunk: int) =
                let depths = fieldMarks.[chunk] |> List.scan (fun depth mark -> match mark with "begin" -> depth + 1 | "end" -> depth - 1 | _ -> depth) 0
                List.forall (fun depth -> depth >= 0) depths && List.last depths = 0
            let inGroup = System.Collections.Generic.HashSet<int>()
            let closeGroup () =
                if group.Count > 1 then
                    for chunk in group do inGroup.Add chunk |> ignore
                    let together =
                        group |> Seq.forall (fun chunk -> keep.[chunk] || (firstAt.[chunk] >= 0 && neutral chunk))
                        && (Seq.pairwise group |> Seq.forall (fun (a, b) -> firstAt.[b] = lastAt.[a] + 1))
                    if not together then
                        let kept = group |> Seq.filter (fun chunk -> keep.[chunk]) |> Seq.length
                        if kept > 0 then noteReason ctx (sprintf "%d blocks: a field begins in one block and ends in another (a table of contents), and one of its %d blocks is written from the model" kept group.Count)
                        for chunk in group do keep.[chunk] <- false
                group.Clear()
            for chunk in blockChunks do
                let begins = fieldMarks.[chunk] |> List.filter ((=) "begin") |> List.length
                let finishes = fieldMarks.[chunk] |> List.filter ((=) "end") |> List.length
                if depth > 0 || begins > finishes then
                    group.Add chunk
                    depth <- max 0 (depth + begins - finishes)
                    if depth = 0 then closeGroup ()
            closeGroup ()
            // A table whose rows or columns came or went (or that changed in more than its paragraphs), that is
            // its element's alone and not one of several held together by a field: its element as it was read with
            // the changes made in it (patchedTableXml), what stands before it as it was read.
            let patchedTables = Array.create<SourceTableText option> chunkCount None
            for chunk in 0 .. chunkCount - 1 do
                if originalCount.[chunk] = 1 && found.[chunk] = 1 && not keep.[chunk] && not (inGroup.Contains chunk) && firstAt.[chunk] >= 0 then
                    let item = items.[firstAt.[chunk]]
                    let next = firstAt.[chunk] + 1
                    match item with
                    | Table table when table.SourceTable.IsSome && table.SourceWhy.IsSome
                                       && not (next < items.Length && isTextBox items.[next]) && body.Starts.[chunk] >= textStart.[chunk] ->
                        match tableElementIn body.Nodes.[chunk] |> Option.bind ctx.Model.TableSource.Value.At with
                        | Some known when tableIsPatchable known table -> patchedTables.[chunk] <- Some known
                        | _ -> ()
                    | _ -> ()
            // In the order the items have now.
            let segments = ResizeArray<BodySegment>()
            let pending = ResizeArray<EditorBlock>()
            let regenerated = ResizeArray<EditorBlock>()
            let flush () =
                if pending.Count > 0 then
                    segments.Add (ModelBlocks (pending.ToArray()))
                    pending.Clear()
            // The marks of ranges that stand between two blocks (in the text before an element: see
            // textStart) stay where they stood when the element is written from the model, or is
            // not there any more: before the element's first item, or before what comes next.
            // (The elements of paragraphs joined to another: what stood before each goes with its text - see the splice.)
            let joinedChunks =
                let found = System.Collections.Generic.HashSet<int>()
                for item in items do
                    match blockChanges item with
                    | Some changes ->
                        for change in changes.Edits do
                            for join in change.Join do
                                let block = join.Block
                                if block >= 0 && block < itemChunks.Length && itemChunks.[block] >= 0 then found.Add itemChunks.[block] |> ignore
                    | None -> ()
                found
            let carriedBefore = System.Collections.Generic.Dictionary<int, ResizeArray<RangeMark>>()
            let carriedAtEnd = ResizeArray<RangeMark>()
            for chunk in 0 .. chunkCount - 1 do
                if originalCount.[chunk] > 0 && not keep.[chunk] && patchedTables.[chunk].IsNone && not (joinedChunks.Contains chunk) && body.Starts.[chunk] > textStart.[chunk] then
                    let marks =
                        body.Nodes
                        |> Array.filter (fun node -> (XmlPatch.span node).Start >= textStart.[chunk] && (XmlPatch.span node).End <= body.Starts.[chunk])
                        |> Array.toList
                        |> List.collect (rangeMarksIn carriedRangeKinds)
                    if not marks.IsEmpty then
                            // (Where: the first item of the element, or of the next element that still has one.)
                            let mutable target = -1
                            let mutable next = chunk
                            while target < 0 && next < chunkCount do
                                if firstAt.[next] >= 0 then target <- firstAt.[next]
                                next <- next + 1
                            if target < 0 then carriedAtEnd.AddRange marks
                            else
                                if not (carriedBefore.ContainsKey target) then carriedBefore.[target] <- ResizeArray<RangeMark>()
                                carriedBefore.[target].AddRange marks
            let mutable position = 0
            let mutable anySource = false
            while position < items.Length do
                let chunk = chunkOfItem.[position]
                match carriedBefore.TryGetValue position with
                | true, marks ->
                    flush ()
                    segments.Add (CarriedMarks (List.ofSeq marks))
                | _ -> ()
                if chunk >= 0 && keep.[chunk] then
                    flush ()
                    match (if patched.[chunk] then patchedText.[chunk] else None) with
                    | Some paragraphs ->
                        segments.Add (SourceText (textStart.[chunk], body.Ends.[chunk], paragraphs))
                        for (_, _, para) in paragraphs do regenerated.Add (Paragraph para)
                    | None -> segments.Add (SourceText (textStart.[chunk], body.Ends.[chunk], []))
                    anySource <- true
                    ctx.ReportSourceBlocks <- ctx.ReportSourceBlocks + originalCount.[chunk]
                    position <- position + originalCount.[chunk]
                elif chunk >= 0 && patchedTables.[chunk].IsSome then
                    flush ()
                    match items.[position] with
                    | Table table -> segments.Add (PatchedTable (textStart.[chunk], body.Ends.[chunk], patchedTables.[chunk].Value, table))
                    | _ -> invalidOp "Patched table segment is not a table"
                    // (What of it is written from the model - a row, a cell - takes its pictures and links from it.)
                    regenerated.Add items.[position]
                    anySource <- true
                    ctx.ReportSourceBlocks <- ctx.ReportSourceBlocks + 1
                    position <- position + 1
                else
                    pending.Add items.[position]
                    regenerated.Add items.[position]
                    ctx.ReportModelBlocks <- ctx.ReportModelBlocks + 1
                    ctx.ReportModelParagraphs <- ctx.ReportModelParagraphs + storyParagraphCount items.[position]
                    position <- position + 1
            flush ()
            if carriedAtEnd.Count > 0 then segments.Add (CarriedMarks (List.ofSeq carriedAtEnd))
            // (Nothing as it was read: a new body. Around it the document stays the text it
            // was - the root's start tag, the section properties - when the page setup is.)
            if not anySource && (not pageSetupUnchanged || not isBody) then None
            else
                let sectionStart = if sectIndex >= 0 then body.Starts.[sectIndex] else contentEnd
                let sectionEnd = if sectIndex >= 0 then body.Ends.[sectIndex] else contentEnd
                Some { Original = xml
                       Root = root
                       Container = bodyNode
                       ContentStart = contentStart
                       ContentEnd = contentEnd
                       Segments = List.ofSeq segments
                       BeforeSection = (tailStart, max tailStart sectionStart)
                       Section = if pageSetupUnchanged || not isBody then Some (sectionStart, sectionEnd) else None
                       OriginalSection = if isBody && sectIndex >= 0 then Some (sectionStart, sectionEnd) else None
                       AfterSection = (sectionEnd, contentEnd)
                       Regenerated = regenerated.ToArray() }
    | _ -> None

/// The namespace declarations the root of a part needs for what was written from the model (`written`):
/// of the prefixes it uses (w, r, m) those the root does not declare. The root element of a document stays
/// as it was read unless something needs more.
let private writerNamespaceEdits (root: XmlDsl.XmlNode) (written: string list) : XmlPatch.ElementEdit list =
    let used =
        written
        |> List.choose XmlPatch.readFragmentRoot
        |> List.collect (fun fragment ->
            let nodes = XmlPatch.descendantsWhere (fun _ -> true) fragment.Node
            (nodes |> List.map XmlPatch.prefix)
            @ (nodes |> List.collect (fun node -> XmlPatch.attributeNames node) |> List.map ModelText.prefixOf))
        |> Set.ofList
    [ "w", Namespaces.word; "r", Namespaces.relationships; "m", Namespaces.math ]
    |> List.filter (fun (prefix, _) -> used.Contains prefix && (XmlPatch.attribute root ("xmlns:" + prefix)).IsNone)
    |> List.map (fun (prefix, uri) -> XmlPatch.SetAttribute ("xmlns:" + prefix, uri))

/// A piece of a part that is put together: a range of the text it was read as, or XML the model wrote.
type private PartPiece =
    | FromSource of start: int * finish: int
    | FromModel of xml: string

/// The ranges of a part that is put together of text as it was read and text
/// written from the model - bookmarks, comment ranges, permission ranges, the
/// ranges of moved text - opened once and closed once each. `parts`: the
/// pieces in order, and whether each was written from the model. Only what
/// was written from the model is changed: a start an earlier piece of the
/// model wrote (an end a later one writes) is not written again; an end of a
/// range nothing opens is left out; a range the model
/// opens and nothing closes - its end stood where the model does not keep
/// one: between two paragraphs, before the first run of another - is closed
/// at the end of its paragraph; one the file's text opens and nothing closes
/// any more is closed at the start of the next piece of the model - or where it
/// is opened, when none follows; an end in the file's text of a range nothing
/// opens any more (the block its start stood in is gone) is left out. What the
/// part had so when it was read (`source`: its text and root) is left as it is.
let private balancedRanges (source: (string * XmlDsl.XmlNode) option) (pieces: PartPiece array) : string array =
    let sourceText, sourceMarks =
        match source with
        | Some (text, root) -> text, rangeMarksIn allRangeKinds root
        | None -> "", []
    let original = source |> Option.map (fun _ -> sourceMarks)
    // (A piece of the source is a range of its text - not always well formed on its own: its marks are those
    // of the tree of the part in that range. A piece of the model is XML of its own.)
    let parts = pieces |> Array.map (function FromSource (start, finish) -> XmlPatch.slice sourceText start finish, false | FromModel text -> text, true)
    let trees = pieces |> Array.map (function FromModel text -> XmlPatch.readFragmentRoot text | FromSource _ -> None)
    let occurrences =
        pieces
        |> Array.mapi (fun index piece ->
            match piece, trees.[index] with
            | FromSource (start, finish), _ ->
                sourceMarks |> List.filter (fun mark -> (XmlPatch.span mark.Node).Start >= start && (XmlPatch.span mark.Node).End <= finish)
            | FromModel _, Some fragment -> rangeMarksIn allRangeKinds fragment.Node
            | FromModel _, None -> [])
    if occurrences |> Array.forall List.isEmpty then parts |> Array.map fst
    else
        let sourceStarts = System.Collections.Generic.Dictionary<string, int>()   // key -> the first piece that has it
        let sourceEnds = System.Collections.Generic.HashSet<string>()
        let modelStarts = System.Collections.Generic.Dictionary<string, int>()    // key -> how many
        let modelEnds = System.Collections.Generic.Dictionary<string, int>()
        let count (table: System.Collections.Generic.Dictionary<string, int>) key = match table.TryGetValue key with | true, n -> n | false, _ -> 0
        let kindOf = System.Collections.Generic.Dictionary<string, string * string>()
        let mutable modelHasAny = false
        for index in 0 .. parts.Length - 1 do
            let isModel = snd parts.[index]
            for mark in occurrences.[index] do
                let key = mark.Key
                kindOf.[key] <- (mark.Kind, mark.Id)
                if isModel then
                    modelHasAny <- true
                    if mark.Opens then modelStarts.[key] <- count modelStarts key + 1 else modelEnds.[key] <- count modelEnds key + 1
                elif mark.Opens then (if not (sourceStarts.ContainsKey key) then sourceStarts.[key] <- index)
                else sourceEnds.Add key |> ignore
        // A range the file's text opens and nothing closes any more: closed at the start of the next piece of the model.
        let closeAt = System.Collections.Generic.Dictionary<int, ResizeArray<string>>()
        // What the file's own text has of a range whose other mark is nowhere any more (the block
        // it stood in was deleted, or the model does not hold it): an end nothing opens, and a
        // start nothing closes that no piece of the model follows. Not what the file had so itself.
        let strayEnds = System.Collections.Generic.HashSet<string>()
        let strayStarts = System.Collections.Generic.HashSet<string>()
        for KeyValue (key, at) in sourceStarts do
            if not (sourceEnds.Contains key) && count modelEnds key = 0 then
                let mutable next = -1
                for index in at + 1 .. parts.Length - 1 do
                    if next < 0 && snd parts.[index] then next <- index
                if next >= 0 then
                    if not (closeAt.ContainsKey next) then closeAt.[next] <- ResizeArray<string>()
                    closeAt.[next].Add key
                else strayStarts.Add key |> ignore
        for key in sourceEnds do
            if not (sourceStarts.ContainsKey key) && count modelStarts key = 0 then strayEnds.Add key |> ignore
        if strayEnds.Count > 0 || strayStarts.Count > 0 || closeAt.Count > 0 then
            // (As the file had it: left as it is.)
            match original with
            | Some marks ->
                let starts = marks |> List.filter (fun mark -> mark.Opens) |> List.map (fun mark -> mark.Key) |> Set.ofList
                let ends = marks |> List.filter (fun mark -> not mark.Opens) |> List.map (fun mark -> mark.Key) |> Set.ofList
                for key in Seq.toArray strayEnds do if not (starts.Contains key) then strayEnds.Remove key |> ignore
                for key in Seq.toArray strayStarts do if not (ends.Contains key) then strayStarts.Remove key |> ignore
                for KeyValue (_, keys) in closeAt do keys.RemoveAll (fun key -> not (ends.Contains key)) |> ignore
            | None ->
                strayEnds.Clear()
                strayStarts.Clear()
        let anyClose = closeAt.Values |> Seq.exists (fun keys -> keys.Count > 0)
        if not modelHasAny && not anyClose && strayEnds.Count = 0 && strayStarts.Count = 0 then parts |> Array.map fst
        else
            let endTag (key: string) = let (kind, id) = kindOf.[key] in rangeEndXml kind id
            let seenStarts = System.Collections.Generic.HashSet<string>()
            let seenEnds = System.Collections.Generic.Dictionary<string, int>()
            // (A part with edits made in it; one that is not well formed stays as it is.)
            let edited (index: int) (edits: (XmlDsl.XmlNode * XmlPatch.ElementEdit list) list) =
                match trees.[index] with
                | Some fragment when not edits.IsEmpty -> XmlPatch.applyFragment XmlPatch.Word.order fragment (XmlPatch.editsAt fragment.Node edits)
                | _ -> fst parts.[index]
            parts
            |> Array.mapi (fun index (text, isModel) ->
                if not isModel then
                    match pieces.[index] with
                    | FromSource (start, finish) when strayEnds.Count > 0 || strayStarts.Count > 0 ->
                        // (A start nothing closes: closed where it is opened. An end nothing opens: left out.)
                        XmlPatch.sliceWith sourceText start finish
                            [ for mark in occurrences.[index] do
                                  let span = XmlPatch.span mark.Node
                                  if mark.Opens && strayStarts.Contains mark.Key then yield span.Start, span.End, XmlPatch.bytesOf sourceText mark.Node + endTag mark.Key
                                  elif not mark.Opens && strayEnds.Contains mark.Key then yield span.Start, span.End, "" ]
                    | _ -> text
                else
                    let opening = match closeAt.TryGetValue index with | true, keys -> keys |> Seq.map endTag |> String.concat "" | false, _ -> ""
                    // (The paragraphs of the part by where they end: the first that ends after a mark is the one it is in, or the next.)
                    let paragraphs =
                        match trees.[index] with
                        | Some fragment -> XmlPatch.descendantsWhere (fun node -> XmlPatch.name node = "w:p") fragment.Node |> List.sortBy (fun para -> (XmlPatch.span para).EndTagStart)
                        | None -> []
                    let atEnd = System.Text.StringBuilder()
                    let edits =
                        [ for mark in occurrences.[index] do
                              let key = mark.Key
                              let keep =
                                  if mark.Opens then
                                      // (The first the model writes. What the file's own text has is not looked at:
                                      // a file can use an id twice, and its text is left as it is.)
                                      seenStarts.Add key
                                  else
                                      let seen = count seenEnds key + 1
                                      seenEnds.[key] <- seen
                                      // (The last one the model writes, of a range that is opened.)
                                      seen = count modelEnds key && (sourceStarts.ContainsKey key || count modelStarts key > 0)
                              if not keep then yield mark.Node, [ XmlPatch.Removed ]
                              elif mark.Opens && not (sourceEnds.Contains key) && count modelEnds key = 0 then
                                  // (A range the model opens and nothing closes: closed at the end of the paragraph it is
                                  // in - or of the first after it -, or at the end of the part.)
                                  match paragraphs |> List.tryFind (fun para -> (XmlPatch.span para).EndTagStart >= (XmlPatch.span mark.Node).End) with
                                  | Some para -> yield para, [ XmlPatch.Append (endTag key) ]
                                  | None -> atEnd.Append(endTag key) |> ignore ]
                    opening + edited index edits + atEnd.ToString())

/// A part that is written from the model whole - the body of a new document, a header or a
/// footer, the notes, the comments - with its ranges opened once and closed once each.
let private balancedPart (xml: string) : string =
    // (The part's root holds what is balanced.)
    match XmlPatch.readPart xml with
    | Some part ->
        let span = XmlPatch.span part.Node
        if span.StartTagEnd >= span.End then xml
        else
            let content = XmlPatch.slice xml span.StartTagEnd span.EndTagStart
            let balanced = (balancedRanges None [| FromModel content |]).[0]
            if balanced = content then xml else XmlPatch.slice xml 0 span.StartTagEnd + balanced + XmlPatch.slice xml span.EndTagStart xml.Length
    | None -> xml

/// What a plan's container holds up to its section properties, written: its blocks (the ranges in them opened
/// and closed once each) and what stands after the last of them. What is written from the model is added to
/// `written`.
let private planContentParts (ctx: PackageContext) (plan: BodyPlan) (setup: PageSetupInfo) (images: ImageScope) (hyperlinks: Map<string, string>) (written: ResizeArray<string>) : ResizeArray<string> =
    let slice (start: int, finish: int) = XmlPatch.slice plan.Original start finish
    let pieces = ResizeArray<PartPiece>()
    // (Paragraphs written from the model splice the part's text: see splicedRuns. The origins of what is
    // written are of this part: see sourceElementAt.)
    ctx.Model.SpliceSourceText <- plan.Original
    ctx.WritingBody <- true
    for segment in plan.Segments do
        match segment with
        | SourceText (start, finish, paragraphs) ->
            // The paragraphs written from their source: each in the place of its element.
            let mutable at = start
            for (paraStart, paraEnd, para) in paragraphs |> List.sortBy (fun (paraStart, _, _) -> paraStart) do
                pieces.Add (FromSource (at, paraStart))
                let mutable splicedBefore = ctx.Model.ReportSpliced
                let paragraph = generateParagraphWithHyperlinksAndImagesWithPageSetup ctx.Model setup hyperlinks images para
                // (The paragraphs of a block kept as it was read that are not written from their source.)
                if ctx.Model.ReportSpliced = splicedBefore then ctx.ReportInline <- ctx.ReportInline + 1
                written.Add paragraph
                pieces.Add (FromModel paragraph)
                at <- paraEnd
            pieces.Add (FromSource (at, finish))
        | PatchedTable (start, finish, table, item) ->
            let span = XmlPatch.span table.TableNode
            pieces.Add (FromSource (start, span.Start))
            let text = patchedTableXml ctx hyperlinks images table item
            written.Add text
            pieces.Add (FromModel text)
            pieces.Add (FromSource (span.End, finish))
        | ModelBlocks items ->
            let text = items |> generateBlocks (generateDocumentItemWithImagesAndPageSetup ctx.Model setup hyperlinks images) (generateTextBoxRun ctx.Model setup hyperlinks images)
            written.Add text
            pieces.Add (FromModel text)
        | CarriedMarks marks ->
            // (As what the model writes: a mark of a range that is not there any more is left out.)
            // (Nor the range of a comment the document has no more.)
            let live =
                marks
                |> List.filter (fun mark -> mark.Kind <> "commentRange" || Set.contains mark.Id ctx.Model.LiveCommentIds)
                |> List.map (fun mark -> XmlPatch.bytesOf plan.Original mark.Node)
                |> String.concat ""
            pieces.Add (FromModel live)
    ctx.WritingBody <- false
    let parts = ResizeArray<string>(balancedRanges (Some (plan.Original, plan.Root)) (pieces.ToArray()))
    parts.Add (slice plan.BeforeSection)
    parts

let private generateDocumentXmlFromPlan (ctx: PackageContext) (plan: BodyPlan) (images: ImageScope) (hyperlinks: Map<string, string>) (pageSetup: PageSetupInfo option) (originalPageSetup: PageSetupInfo option) (headers: HeaderFooterPartInfo array) (footers: HeaderFooterPartInfo array) : string =
    let setup = pageSetup |> Option.defaultValue defaultPageSetupInfo
    let slice (start: int, finish: int) = XmlPatch.slice plan.Original start finish
    let written = ResizeArray<string>()
    let parts = planContentParts ctx plan setup images hyperlinks written
    match plan.Section with
    | Some range -> parts.Add (slice range)
    | None ->
        let globalHeaders = headers |> Array.filter (fun part -> part.IsGlobal)
        let globalFooters = footers |> Array.filter (fun part -> part.IsGlobal)
        let generated = generateSectionProperties pageSetup globalHeaders globalFooters
        let text =
            match plan.OriginalSection, originalPageSetup with
            | Some range, Some original ->
                let before = generateSectionProperties (Some original) globalHeaders globalFooters
                XmlPatch.merge XmlPatch.Word.order "<w:sectPr/>" (slice range) before generated
                |> Option.defaultValue generated
            | _ -> generated
        written.Add text
        parts.Add text
    parts.Add (slice plan.AfterSection)
    // The part: its root's start tag with the namespaces what was written needs, the body's start tag as it
    // was read, what it holds now, and the rest as it was read.
    let rootSpan = XmlPatch.span plan.Root
    let namespaces = if written.Count > 0 then writerNamespaceEdits plan.Root (List.ofSeq written) else []
    XmlPatch.slice plan.Original 0 rootSpan.Start
    + XmlPatch.startTag plan.Original plan.Root namespaces
    + XmlPatch.slice plan.Original rootSpan.StartTagEnd plan.ContentStart
    + String.concat "" parts
    + XmlPatch.slice plan.Original plan.ContentEnd plan.Original.Length

/// The element of a note or a comment written from its plan (see planBody): its start and end tags as they were
/// read, what it holds as `planContentParts` writes it. The run with the note's or comment's own mark
/// (w:footnoteRef, w:endnoteRef, w:annotationRef) - which the model does not hold as a run - is what it was:
/// where what was written has none, the one that was read is put at the start of the first paragraph.
let private generateStoryXmlFromPlan (ctx: PackageContext) (plan: BodyPlan) (images: ImageScope) (hyperlinks: Map<string, string>) : string =
    let written = ResizeArray<string>()
    let parts = planContentParts ctx plan defaultPageSetupInfo images hyperlinks written
    let span = XmlPatch.span plan.Container
    let text =
        XmlPatch.slice plan.Original span.Start plan.ContentStart
        + String.concat "" parts
        + XmlPatch.slice plan.Original plan.ContentEnd span.End
    let isMark (node: XmlDsl.XmlNode) = List.contains (XmlPatch.localName node) [ "footnoteRef"; "endnoteRef"; "annotationRef" ]
    let markRunOf (node: XmlDsl.XmlNode) =
        XmlPatch.descendantsWhere (fun run -> XmlPatch.localName run = "r" && (XmlPatch.elements run |> List.exists isMark)) node |> List.tryHead
    match markRunOf plan.Container, XmlPatch.readFragmentRoot text with
    | Some mark, Some fragment ->
        match XmlPatch.elements fragment.Node with
        | [ story ] when (markRunOf story).IsNone ->
            match XmlPatch.descendantsWhere (fun node -> XmlPatch.localName node = "p") story |> List.tryHead with
            | Some first -> XmlPatch.applyFragment XmlPatch.Word.order fragment (XmlPatch.editsAt fragment.Node [ first, [ XmlPatch.Prepend (XmlPatch.bytesOf plan.Original mark) ] ])
            | None -> text
        | _ -> text
    | _ -> text


let private generateDocumentXmlPackage (ctx: PackageContext) (paragraphs: EditorBlock array) (images: ImageScope) (pageSetup: PageSetupInfo option) (headers: HeaderFooterPartInfo array) (footers: HeaderFooterPartInfo array) : string =
    let hyperlinks = hyperlinkMap ctx.Model paragraphs
    let setup = pageSetup |> Option.defaultValue defaultPageSetupInfo
    let paraElements =
        paragraphs
        |> generateBlocks (generateDocumentItemWithImagesAndPageSetup ctx.Model setup hyperlinks images) (generateTextBoxRun ctx.Model setup hyperlinks images)
    let globalHeaders = headers |> Array.filter (fun part -> part.IsGlobal)
    let globalFooters = footers |> Array.filter (fun part -> part.IsGlobal)
    let body = element "w:body" [] [paraElements; generateSectionProperties pageSetup globalHeaders globalFooters]
    let doc = element "w:document" [("xmlns:w", Namespaces.word); ("xmlns:r", Namespaces.relationships); ("xmlns:m", Namespaces.math)] [body]
    document doc

/// Generate [Content_Types].xml
let generateContentTypes (hasComments: bool) (hasFootnotes: bool) (hasEndnotes: bool) : string =
    let children = [
        yield selfClosing "Default" [("Extension", "rels"); ("ContentType", "application/vnd.openxmlformats-package.relationships+xml")]
        yield selfClosing "Default" [("Extension", "xml"); ("ContentType", "application/xml")]
        yield selfClosing "Override" [("PartName", "/word/document.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")]
        yield selfClosing "Override" [("PartName", "/word/styles.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml")]
        yield selfClosing "Override" [("PartName", "/word/numbering.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml")]
        if hasComments then
            yield selfClosing "Override" [("PartName", "/word/comments.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml")]
        if hasFootnotes then
            yield selfClosing "Override" [("PartName", "/word/footnotes.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml")]
        if hasEndnotes then
            yield selfClosing "Override" [("PartName", "/word/endnotes.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml")]
    ]
    let root = element "Types" [("xmlns", Namespaces.contentTypes)] children
    document root


/// [Content_Types].xml for a package that carries kept parts: the original
/// Default entries, the writer's own, and the Override entries of the kept parts.
let private generateContentTypesPreserved (preserved: PreservedPackage) hasNumbering hasSettings hasComments hasFootnotes hasEndnotes (headers: HeaderFooterPartInfo array) (footers: HeaderFooterPartInfo array) (imageAssets: ImageAsset array) =
    let own =
        [ yield "/word/document.xml", preserved.MainContentType |> Option.defaultValue "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
          yield "/word/styles.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
          if hasNumbering then yield "/word/numbering.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
          if hasSettings then yield "/word/settings.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
          for header in headers do yield header.PartName, "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
          for footer in footers do yield footer.PartName, "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
          if hasComments then yield "/word/comments.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
          if hasFootnotes then yield "/word/footnotes.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
          if hasEndnotes then yield "/word/endnotes.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml" ]
    let ownNames = own |> List.map (fun (name, _) -> name.ToLowerInvariant()) |> Set.ofList
    let keptNames = preserved.Parts |> List.map (fun part -> ("/" + part.Path).ToLowerInvariant()) |> Set.ofList
    let keptOverrides =
        preserved.Overrides
        |> List.filter (fun (name, _) -> Set.contains (name.ToLowerInvariant()) keptNames && not (Set.contains (name.ToLowerInvariant()) ownNames))
    let defaults =
        [ yield "rels", "application/vnd.openxmlformats-package.relationships+xml"
          yield "xml", "application/xml"
          yield! preserved.Defaults
          for image in imageAssets do
              if not image.Linked then yield image.Extension, image.ContentType ]
        |> List.distinctBy (fun (extension, _) -> extension.ToLowerInvariant())
    let children =
        [ for extension, contentType in defaults do yield selfClosing "Default" [("Extension", extension); ("ContentType", contentType)]
          for name, contentType in own @ keptOverrides do yield selfClosing "Override" [("PartName", name); ("ContentType", contentType)] ]
    document (element "Types" [("xmlns", Namespaces.contentTypes)] children)

let generateContentTypesPackage hasSettings hasComments hasFootnotes hasEndnotes (headers: HeaderFooterPartInfo array) (footers: HeaderFooterPartInfo array) (imageAssets: ImageAsset array) =
    let children = [
        yield selfClosing "Default" [("Extension", "rels"); ("ContentType", "application/vnd.openxmlformats-package.relationships+xml")]
        yield selfClosing "Default" [("Extension", "xml"); ("ContentType", "application/xml")]
        for extension, contentType in imageAssets |> Array.filter (fun image -> not image.Linked) |> Array.map (fun image -> image.Extension, image.ContentType) |> Array.distinct do
            yield selfClosing "Default" [("Extension", extension); ("ContentType", contentType)]
        yield selfClosing "Override" [("PartName", "/word/document.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")]
        yield selfClosing "Override" [("PartName", "/word/styles.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml")]
        yield selfClosing "Override" [("PartName", "/word/numbering.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml")]
        if hasSettings then yield selfClosing "Override" [("PartName", "/word/settings.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml")]
        for header in headers do
            yield selfClosing "Override" [("PartName", header.PartName); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")]
        for footer in footers do
            yield selfClosing "Override" [("PartName", footer.PartName); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")]
        if hasComments then yield selfClosing "Override" [("PartName", "/word/comments.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml")]
        if hasFootnotes then yield selfClosing "Override" [("PartName", "/word/footnotes.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml")]
        if hasEndnotes then yield selfClosing "Override" [("PartName", "/word/endnotes.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml")]
    ]
    document (element "Types" [("xmlns", Namespaces.contentTypes)] children)


/// Generate word/_rels/document.xml.rels
/// The writer's own relationship ids are named, not numbered: kept
/// relationships of the opened document have ids like rId1, rId2.
let private documentRelationshipElementsWith (ctx: PackageContext) (hasNumbering: bool) (hyperlinks: Map<string, string>) (hasComments: bool) (hasFootnotes: bool) (hasEndnotes: bool) =
    let baseRels = [
        yield selfClosing "Relationship" [("Id", freshRelationshipId ctx.Model "rIdStyles"); ("Type", RelationshipTypes.styles); ("Target", "styles.xml")]
        if hasNumbering then
            yield selfClosing "Relationship" [("Id", freshRelationshipId ctx.Model "rIdNumbering"); ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"); ("Target", "numbering.xml")]
        if hasComments then
            yield selfClosing "Relationship" [("Id", freshRelationshipId ctx.Model "rIdComments"); ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"); ("Target", "comments.xml")]
        if hasFootnotes then
            yield selfClosing "Relationship" [("Id", freshRelationshipId ctx.Model "rIdFootnotes"); ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"); ("Target", "footnotes.xml")]
        if hasEndnotes then
            yield selfClosing "Relationship" [("Id", freshRelationshipId ctx.Model "rIdEndnotes"); ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"); ("Target", "endnotes.xml")]
    ]
    let hyperlinkRels =
        hyperlinks
        |> Map.toList
        |> List.map (fun (url, relId) ->
            selfClosing "Relationship" [("Id", relId); ("Type", RelationshipTypes.hyperlink); ("Target", url); ("TargetMode", "External")])
    baseRels @ hyperlinkRels

let private documentRelationshipElements (ctx: PackageContext) (hyperlinks: Map<string, string>) (hasComments: bool) (hasFootnotes: bool) (hasEndnotes: bool) =
    documentRelationshipElementsWith ctx true hyperlinks hasComments hasFootnotes hasEndnotes

let private generateDocumentRels (ctx: PackageContext) (hyperlinks: Map<string, string>) (hasComments: bool) (hasFootnotes: bool) (hasEndnotes: bool) : string =
    let rels = documentRelationshipElements ctx hyperlinks hasComments hasFootnotes hasEndnotes
    let root = element "Relationships" [("xmlns", Namespaces.relationshipsPackage)] rels
    document root

let private generateImageRelationships (assets: ImageAsset array) =
    assets
    |> Array.map (fun asset ->
        if asset.Linked then
            selfClosing "Relationship" [("Id", asset.RelId); ("Type", RelationshipTypes.image); ("Target", asset.Src); ("TargetMode", "External")]
        else
            selfClosing "Relationship" [("Id", asset.RelId); ("Type", RelationshipTypes.image); ("Target", sprintf "media/%s" asset.FileName)])
    |> Array.toList

let private generateHyperlinkRelationships (hyperlinks: Map<string, string>) =
    hyperlinks
    |> Map.toList
    |> List.map (fun (url, relId) ->
        selfClosing "Relationship" [("Id", relId); ("Type", RelationshipTypes.hyperlink); ("Target", url); ("TargetMode", "External")])


let private generateDocumentRelsPackage (ctx: PackageContext) hasSettings hyperlinks hasComments hasFootnotes hasEndnotes (headers: HeaderFooterPartInfo array) (footers: HeaderFooterPartInfo array) (images: ImageScope) =
    let headerFooterRels =
        [ if hasSettings then yield selfClosing "Relationship" [("Id", freshRelationshipId ctx.Model "rIdSettings"); ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"); ("Target", "settings.xml")]
          for header in headers -> selfClosing "Relationship" [("Id", header.RelId); ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"); ("Target", header.Target)]
          for footer in footers -> selfClosing "Relationship" [("Id", footer.RelId); ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"); ("Target", footer.Target)] ]
    let rels = documentRelationshipElements ctx hyperlinks hasComments hasFootnotes hasEndnotes @ headerFooterRels @ generateImageRelationships images.Assets
    let root = element "Relationships" [("xmlns", Namespaces.relationshipsPackage)] rels
    document root

/// The relationships part of the main document: the writer's own, then the kept ones.
let private generateDocumentRelsPreserved (ctx: PackageContext) (kept: PreservedRelationshipInfo list) hasNumbering hasSettings hyperlinks hasComments hasFootnotes hasEndnotes (headers: HeaderFooterPartInfo array) (footers: HeaderFooterPartInfo array) (images: ImageScope) =
    let headerFooterRels =
        [ if hasSettings then yield selfClosing "Relationship" [("Id", freshRelationshipId ctx.Model "rIdSettings"); ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"); ("Target", "settings.xml")]
          for header in headers -> selfClosing "Relationship" [("Id", header.RelId); ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"); ("Target", header.Target)]
          for footer in footers -> selfClosing "Relationship" [("Id", footer.RelId); ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"); ("Target", footer.Target)] ]
    let rels =
        documentRelationshipElementsWith ctx hasNumbering hyperlinks hasComments hasFootnotes hasEndnotes
        @ headerFooterRels
        @ generateImageRelationships images.Assets
        @ (kept |> List.map preservedRelationshipElement)
    document (element "Relationships" [("xmlns", Namespaces.relationshipsPackage)] rels)

let private generateRootRelsPreserved (preserved: PreservedPackage) : string =
    let main = selfClosing "Relationship" [("Id", preserved.DocumentRelationshipId); ("Type", RelationshipTypes.officeDocument); ("Target", "word/document.xml")]
    let kept =
        preserved.RootRelationships
        |> List.filter (fun rel -> rel.Id <> preserved.DocumentRelationshipId)
        |> List.map preservedRelationshipElement
    document (element "Relationships" [("xmlns", Namespaces.relationshipsPackage)] (main :: kept))

let private generatePartRelsWithKept (kept: PreservedRelationshipInfo list) (hyperlinks: Map<string, string>) (images: ImageScope) =
    let rels = generateImageRelationships images.Assets @ generateHyperlinkRelationships hyperlinks @ (kept |> List.map preservedRelationshipElement)
    document (element "Relationships" [("xmlns", Namespaces.relationshipsPackage)] rels)

let private generatePartRels (hyperlinks: Map<string, string>) (images: ImageScope) =
    generatePartRelsWithKept [] hyperlinks images


/// A footnote / endnote with its body as blocks (package.Footnotes /
/// package.Endnotes), or with plain text only (a reference run's NoteText, for
/// callers that do not deliver bodies).
type NotePart = {
    Id: string
    /// w:type of a special note: separator, continuationSeparator, continuationNotice.
    Kind: string option
    HasMark: bool
    MarkStyle: string option
    Blocks: EditorBlock array
    PlainText: string option
    /// The note is what it was when it was read.
    Unchanged: bool
    /// Of one that changed: how many paragraphs each of its blocks had when it was read (its blocks say which
    /// they were then, as the body's do - see planBody).
    SourceStoryCounts: int array
}

type CommentPart = {
    Id: string
    Author: string option
    Date: string option
    Initials: string option
    Blocks: EditorBlock array
    PlainText: string option
    Unchanged: bool
    SourceStoryCounts: int array
}

let private notePartsFromPackage (notes: Story array) : NotePart list =
    notes
    |> Array.toList
    |> List.choose (fun note ->
        if String.IsNullOrEmpty note.Id then None
        else Some {
            Id = note.Id; Kind = note.Kind; HasMark = note.HasMark |> Option.defaultValue true
            MarkStyle = note.MarkStyle; Blocks = note.Blocks; PlainText = None
            Unchanged = note.Unchanged; SourceStoryCounts = note.SourceStoryCounts })

let private commentPartsFromPackage (comments: WordEditorPackage.CommentPart array) : CommentPart list =
    comments
    |> Array.toList
    |> List.choose (fun comment ->
        if String.IsNullOrEmpty comment.Id then None
        else Some {
            Id = comment.Id; Author = comment.Author; Date = comment.Date
            Initials = comment.Initials; Blocks = comment.Blocks; PlainText = None
            Unchanged = comment.Unchanged; SourceStoryCounts = comment.SourceStoryCounts })

/// Bodies from the package first; a note that is only known from a reference
/// run (NoteText) is added as plain text.
let private mergeNoteParts (fromPackage: NotePart list) (fromRuns: NoteRef list) : NotePart list =
    let known = fromPackage |> List.map (fun note -> note.Id) |> Set.ofList
    fromPackage
    @ (fromRuns
       |> List.filter (fun note -> not (Set.contains note.Id known))
       |> List.map (fun note -> { Id = note.Id; Kind = None; HasMark = false; MarkStyle = None; Blocks = [||]; PlainText = Some note.Text; Unchanged = false; SourceStoryCounts = [||] }))

let private mergeCommentParts (fromPackage: CommentPart list) (fromRuns: CommentRef list) : CommentPart list =
    let known = fromPackage |> List.map (fun comment -> comment.Id) |> Set.ofList
    fromPackage
    @ (fromRuns
       |> List.filter (fun comment -> not (Set.contains comment.Id known))
       |> List.map (fun comment -> { Id = comment.Id; Author = comment.Author; Date = comment.Date; Initials = None; Blocks = [||]; PlainText = Some comment.Text; Unchanged = false; SourceStoryCounts = [||] }))

let private partRootAttributes =
    [("xmlns:w", Namespaces.word); ("xmlns:r", Namespaces.relationships); ("xmlns:m", Namespaces.math)]

/// footnotes.xml / endnotes.xml with the note bodies as blocks. The reader
/// used to hand over the text of a note only, and this wrote one plain run per
/// note, so a table, a hyperlink or bold text in a footnote was gone after a save.
let private noteElement (noteElementName: string) (markElementName: string) (generateBody: EditorBlock array -> string) (note: NotePart) : string =
    let markRun (note: NotePart) =
        match note.Kind with
        | Some "separator" -> element "w:r" [] [selfClosing "w:separator" []]
        | Some "continuationSeparator" -> element "w:r" [] [selfClosing "w:continuationSeparator" []]
        | _ ->
            let rPr =
                match note.MarkStyle with
                | Some styleId when styleId <> "" -> [selfClosing "w:rStyle" [("w:val", styleId)]]
                | _ -> [selfClosing "w:vertAlign" [("w:val", "superscript")]]
            element "w:r" [] [element "w:rPr" [] rPr; selfClosing markElementName []]
    // The mark (or the separator line of a separator note) is the first run of
    // the first paragraph.
    let withMark (note: NotePart) (body: string) =
        let isSeparator = note.Kind = Some "separator" || note.Kind = Some "continuationSeparator"
        if not note.HasMark && not isSeparator then body
        else
            match XmlPatch.readFragmentRoot body with
            | Some fragment ->
                match XmlPatch.elements fragment.Node with
                | first :: _ when XmlPatch.name first = "w:p" ->
                    XmlPatch.applyFragment XmlPatch.Word.order fragment (XmlPatch.editsAt fragment.Node [ first, [ XmlPatch.Prepend (markRun note) ] ])
                | _ -> body
            | None -> body
    let attrs =
        [ match note.Kind with
          | Some kind -> yield ("w:type", kind)
          | None -> ()
          yield ("w:id", note.Id) ]
    let body =
        match note.PlainText with
        | Some text -> element "w:p" [] [runElement [] false text]
        | None ->
            let body = generateBody note.Blocks
            // A note must hold at least one paragraph.
            withMark note (if body = "" then "<w:p></w:p>" else body)
    element noteElementName attrs [body]

let private generateNotesPart (rootName: string) (noteElementName: string) (markElementName: string) (generateBody: EditorBlock array -> string) (notes: NotePart list) : string =
    document (element rootName partRootAttributes (notes |> List.map (noteElement noteElementName markElementName generateBody)))

let private commentElement (generateBody: EditorBlock array -> string) (comment: CommentPart) : string =
    let attrs =
        [ yield ("w:id", comment.Id)
          yield ("w:author", comment.Author |> Option.defaultValue "RedBat")
          match comment.Date with
          | Some date -> yield ("w:date", date)
          | None -> ()
          match comment.Initials with
          | Some initials -> yield ("w:initials", initials)
          | None -> () ]
    let body =
        match comment.PlainText with
        | Some text -> element "w:p" [] [runElement [] false text]
        | None ->
            let body = generateBody comment.Blocks
            if body = "" then element "w:p" [] [] else body
    element "w:comment" attrs [body]

let private generateCommentsPart (generateBody: EditorBlock array -> string) (comments: CommentPart list) : string =
    document (element "w:comments" partRootAttributes (comments |> List.map (commentElement generateBody)))

/// A part of notes or comments (`elementName`: w:footnote, w:endnote, w:comment) as it was read with what
/// changed made in it: the element of each that is what it was (`unchanged`, by id) stays its text, one that
/// changed is what the model writes of it (`write`; `ids`: those the model has, in order), one the model has
/// no more goes, a new one is put in at the end; the root keeps its declarations and gets those what the model
/// wrote needs. None when the part does not read.
let private patchedStoriesPart (xml: string) (elementName: string) (ids: string list) (unchanged: Set<string>) (write: string -> string) : string option =
    XmlPatch.readPart xml
    |> Option.map (fun part ->
        let readIdsFirst = XmlPatch.childrenNamed part elementName |> List.choose (fun node -> XmlPatch.attribute node "w:id") |> Set.ofList
        let written = ids |> List.filter (fun id -> not (Set.contains id unchanged && Set.contains id readIdsFirst)) |> List.map (fun id -> id, write id)
        let byId = Map.ofList written
        let present = Set.ofList ids
        let read = XmlPatch.childrenNamed part elementName |> List.choose (fun node -> XmlPatch.attribute node "w:id" |> Option.map (fun id -> id, node))
        let readIds = read |> List.map fst |> Set.ofList
        let changes =
            [ for id, node in read do
                  let at = (XmlPatch.span node).Start
                  match Map.tryFind id byId with
                  | Some text -> yield XmlPatch.EditAt (at, [ XmlPatch.Replaced text ])
                  | None when not (Set.contains id present) -> yield XmlPatch.EditAt (at, [ XmlPatch.Removed ])
                  | None -> ()
              for id, text in written do
                  if not (Set.contains id readIds) then yield XmlPatch.Append text ]
        let fromModel = written |> List.map snd
        let namespaces = if fromModel.IsEmpty then [] else writerNamespaceEdits part.Node fromModel
        if changes.IsEmpty && namespaces.IsEmpty then xml
        else XmlPatch.applyPart (fun _ -> []) part (namespaces @ changes))

type NumberingInfo = {
    NumId: string
    Level: int
    Format: string
    LevelText: string
    Start: int option
    HasOverride: bool
    SourceId: string option
    FormatEdited: bool
    StartOverrides: (int * int) list
    /// The level's own indents (px; a negative first-line indent is a hanging
    /// one): what its paragraphs have without stating an indent themselves.
    IndentLeft: float option
    FirstLineIndent: float option
    /// A list id of the file that the file's numbering part does not define
    /// (the reader then has no format and no level text for it). It stays
    /// undefined: a definition made up for it gave the paragraphs numbers
    /// they do not have in Word. The id is still one that is taken.
    Undefined: bool
}

let private numericListId (id: string option) =
    id |> Option.bind (fun value ->
        match System.Int32.TryParse value with
        | true, parsed -> Some (string parsed)
        | _ -> None)

let private nonEmptyString = Option.bind (fun value -> if String.IsNullOrEmpty value then None else Some value)

/// One entry per (list id, level) used in the document. The list id, number
/// format, level text and start value all come from the paragraph: without
/// them every numbered paragraph was written with numId 2, which merged
/// separate lists into one decimal list and lost letters / roman numerals.
let private collectNumberingInfoOf (undefinedLists: bool) (items: EditorBlock array) =
    items
    |> Array.toList
    |> List.collect collectItemParagraphs
    |> List.choose (fun para ->
        match para.List with
        | None -> None
        | Some list when list.Style = "none" -> None
        | Some list ->
            numericListId list.Id
            |> Option.filter (fun numId -> numId <> "0")
            |> Option.map (fun numId ->
                let level = max 0 (list.Level - 1)
                let format = nonEmptyString list.NumberFormat |> Option.defaultValue (if list.Style = "bullet" then "bullet" else "decimal")
                let levelText = nonEmptyString list.LevelText |> Option.defaultValue (if format = "bullet" then "\u2022" else "%" + string (level + 1) + ".")
                { NumId = numId; Level = level; Format = format; LevelText = levelText
                  Start = list.Start; HasOverride = list.Restart
                  SourceId = numericListId list.SourceId; FormatEdited = list.FormatEdited
                  StartOverrides = list.StartOverrides |> List.choose (fun (level, start) ->
                      if level >= 1 && level <= 9 then Some (level - 1, start) else None)
                  IndentLeft = list.ListLevelIndents |> Option.bind fst
                  FirstLineIndent = list.ListLevelIndents |> Option.bind snd
                  Undefined = undefinedLists && (nonEmptyString list.NumberFormat).IsNone && (nonEmptyString list.LevelText).IsNone }))
    |> List.distinctBy (fun info -> info.NumId, info.Level)

let private numberingStarts (infos: NumberingInfo list) =
    (Map.empty, infos)
    ||> List.fold (fun starts info ->
        let starts = if info.HasOverride then Map.add info.Level (info.Start |> Option.defaultValue 1) starts else starts
        (starts, info.StartOverrides) ||> List.fold (fun result (level, value) -> Map.add level value result))


/// Generate minimal styles.xml
type StyleInfo = {
    StyleId: string
    Name: string option
    BasedOn: string option
    Next: string option
    Linked: string option
    IsQuickStyle: bool
    IsDefault: bool
    Aliases: string option
    UiPriority: int option
    SemiHidden: bool
    UnhideWhenUsed: bool
    StyleType: string
    Bold: bool option
    Italic: bool option
    Underline: bool option
    AllCaps: bool option
    SmallCaps: bool option
    Strike: bool option
    VertAlign: string option
    FontSize: float option
    CharacterSpacing: float option
    FontFamily: string option
    Color: string option
    Highlight: string option
    SpacingBefore: float option
    SpacingAfter: float option
    LineSpacing: float option
    LineRule: string option
    Alignment: string option
    IndentLeft: float option
    IndentRight: float option
    FirstLineIndent: float option
    BackgroundColor: string option
    OutlineLevel: int option
    KeepNext: bool option
    KeepLines: bool option
    PageBreakBefore: bool option
    WidowControl: bool option
}

let private emptyStyleInfo styleId styleType = {
    StyleId = styleId
    Name = None
    BasedOn = None
    Next = None
    Linked = None
    IsQuickStyle = false
    IsDefault = false
    Aliases = None
    UiPriority = None
    SemiHidden = false
    UnhideWhenUsed = false
    StyleType = styleType
    Bold = None
    Italic = None
    Underline = None
    AllCaps = None
    SmallCaps = None
    Strike = None
    VertAlign = None
    FontSize = None
    CharacterSpacing = None
    FontFamily = None
    Color = None
    Highlight = None
    SpacingBefore = None
    SpacingAfter = None
    LineSpacing = None
    LineRule = None
    Alignment = None
    IndentLeft = None
    IndentRight = None
    FirstLineIndent = None
    BackgroundColor = None
    OutlineLevel = None
    KeepNext = None
    KeepLines = None
    PageBreakBefore = None
    WidowControl = None
}

let private styleInfoFromInference styleId styleType (style: InferredStyle) =
    { emptyStyleInfo styleId styleType with
        Name = style.Name
        BasedOn = style.BasedOn
        Next = style.Next
        Linked = style.Linked
        IsQuickStyle = style.IsQuickStyle |> Option.defaultValue false
        IsDefault = style.IsDefault |> Option.defaultValue false
        Aliases = style.Aliases
        UiPriority = style.UiPriority
        SemiHidden = style.SemiHidden |> Option.defaultValue false
        UnhideWhenUsed = style.UnhideWhenUsed |> Option.defaultValue false
        Bold = style.Bold
        Italic = style.Italic
        Underline = style.Underline
        AllCaps = style.AllCaps
        SmallCaps = style.SmallCaps
        Strike = style.Strike
        VertAlign = style.VertAlign
        FontSize = style.FontSize
        CharacterSpacing = style.CharacterSpacing
        FontFamily = style.FontFamily
        Color = style.Color
        Highlight = style.Highlight
        SpacingBefore = style.SpacingBefore
        SpacingAfter = style.SpacingAfter
        LineSpacing = style.LineSpacing
        LineRule = style.LineRule
        Alignment = style.Alignment
        IndentLeft = style.IndentLeft
        IndentRight = style.IndentRight
        FirstLineIndent = style.FirstLineIndent
        BackgroundColor = style.BackgroundColor
        OutlineLevel = style.OutlineLevel
        KeepNext = style.KeepNext
        KeepLines = style.KeepLines
        PageBreakBefore = style.PageBreakBefore
        WidowControl = style.WidowControl
    }

let private styleInfoFromCatalogFields styleId styleType (style: WordEditorPackage.Style) =
    { emptyStyleInfo styleId styleType with
        Name = style.Name
        BasedOn = style.BasedOn
        Next = style.Next
        Linked = style.Linked
        IsQuickStyle = style.IsQuickStyle |> Option.defaultValue false
        IsDefault = style.IsDefault
        Aliases = style.Aliases
        UiPriority = style.UiPriority
        SemiHidden = style.SemiHidden |> Option.defaultValue false
        UnhideWhenUsed = style.UnhideWhenUsed |> Option.defaultValue false
        Bold = style.Run |> Option.bind (fun run -> run.Bold)
        Italic = style.Run |> Option.bind (fun run -> run.Italic)
        Underline = style.Run |> Option.bind (fun run -> run.Underline)
        AllCaps = style.Run |> Option.bind (fun run -> run.AllCaps)
        SmallCaps = style.Run |> Option.bind (fun run -> run.SmallCaps)
        Strike = style.Run |> Option.bind (fun run -> run.Strike)
        VertAlign = style.Run |> Option.bind (fun run -> run.VertAlign)
        FontSize = style.Run |> Option.bind (fun run -> run.FontSize)
        CharacterSpacing = style.Run |> Option.bind (fun run -> run.CharacterSpacing)
        FontFamily = style.Run |> Option.bind (fun run -> run.FontFamily)
        Color = style.Run |> Option.bind (fun run -> run.Color)
        Highlight = style.Run |> Option.bind (fun run -> run.Highlight)
        SpacingBefore = style.Paragraph |> Option.bind (fun paragraph -> paragraph.SpacingBefore)
        SpacingAfter = style.Paragraph |> Option.bind (fun paragraph -> paragraph.SpacingAfter)
        LineSpacing = style.Paragraph |> Option.bind (fun paragraph -> paragraph.LineSpacing)
        LineRule = style.Paragraph |> Option.bind (fun paragraph -> paragraph.LineSpacingRule)
        Alignment = style.Paragraph |> Option.bind (fun paragraph -> paragraph.Alignment)
        IndentLeft = style.Paragraph |> Option.bind (fun paragraph -> paragraph.IndentLeft)
        IndentRight = style.Paragraph |> Option.bind (fun paragraph -> paragraph.IndentRight)
        FirstLineIndent = style.Paragraph |> Option.bind (fun paragraph -> paragraph.FirstLineIndent)
        BackgroundColor = style.Paragraph |> Option.bind (fun paragraph -> paragraph.BackgroundColor)
        OutlineLevel = style.Paragraph |> Option.bind (fun paragraph -> paragraph.OutlineLevel)
        KeepNext = style.Paragraph |> Option.bind (fun paragraph -> paragraph.KeepNext)
        KeepLines = style.Paragraph |> Option.bind (fun paragraph -> paragraph.KeepLines)
        PageBreakBefore = style.Paragraph |> Option.bind (fun paragraph -> paragraph.PageBreakBefore)
        WidowControl = style.Paragraph |> Option.bind (fun paragraph -> paragraph.WidowControl)
    }

let private styleInfoFromCatalog (style: WordEditorPackage.Style) : StyleInfo option =
    if String.IsNullOrEmpty style.Id then None
    else
        let baseInfo = styleInfoFromCatalogFields style.Id style.StyleType style
        Some baseInfo

let private styleIdFromParagraph (para: EditorParagraph) =
    para.Formatting.Style
    |> nonEmptyString
    |> Option.filter (fun id -> id <> documentDefaultsStyleId)
    |> Option.map (fun id ->
        match para.Formatting.StyleInfo with
        | Some info -> styleInfoFromInference id "paragraph" info
        | None -> emptyStyleInfo id "paragraph")

let private styleIdFromRun (run: EditorRun) =
    run.Formatting.Style
    |> nonEmptyString
    |> Option.map (fun id ->
        match run.Formatting.StyleInfo with
        | Some info -> styleInfoFromInference id "character" info
        | None -> emptyStyleInfo id "character")

let private runStylesFromParagraph (para: EditorParagraph) =
    para.Runs |> Array.toList |> List.choose styleIdFromRun

let private collectItemParagraphStyles (item: EditorBlock) =
    collectItemParagraphs item
    |> List.collect (fun para -> (styleIdFromParagraph para |> Option.toList) @ runStylesFromParagraph para)

let private collectStyles (items: EditorBlock array) =
    items
    |> Array.toList
    |> List.collect collectItemParagraphStyles
    |> List.distinctBy (fun style -> style.StyleType, style.StyleId)

let private collectPackageStyles (styles: WordEditorPackage.Style array) =
    styles
    |> Array.toList
    |> List.choose styleInfoFromCatalog

/// Style run properties, in CT_RPr order (see orderedRunProps). A style only
/// states what it sets, so absent properties are left out rather than written off.
let private styleRunProps style =
    let styleToggle name value =
        match value with
        | Some isOn -> [toggle name isOn]
        | None -> []
    [
        match style.FontFamily with
        | Some family when family <> "" -> yield selfClosing "w:rFonts" [("w:ascii", family); ("w:hAnsi", family); ("w:cs", family)]
        | _ -> ()
        yield! styleToggle "w:b" style.Bold
        yield! styleToggle "w:i" style.Italic
        yield! styleToggle "w:caps" style.AllCaps
        yield! styleToggle "w:smallCaps" style.SmallCaps
        yield! styleToggle "w:strike" style.Strike
        match InverseConversion.cssColorToWord style.Color with
        | Some color -> yield selfClosing "w:color" [("w:val", color)]
        | None -> ()
        match style.CharacterSpacing with
        | Some spacing when abs spacing >= 0.001 -> yield selfClosing "w:spacing" [("w:val", InverseConversion.pixelsToTwips spacing)]
        | _ -> ()
        match style.FontSize with
        | Some size ->
            let value = InverseConversion.pixelsToHalfPoints size
            yield selfClosing "w:sz" [("w:val", value)]
            yield selfClosing "w:szCs" [("w:val", value)]
        | _ -> ()
        match style.Highlight with
        | Some color when color <> "" -> yield selfClosing "w:highlight" [("w:val", color)]
        | _ -> ()
        match style.Underline with
        | Some isOn -> yield selfClosing "w:u" [("w:val", (if isOn then "single" else "none"))]
        | None -> ()
        match style.VertAlign with | Some value when value = "superscript" || value = "subscript" -> yield selfClosing "w:vertAlign" [("w:val", value)] | _ -> ()
    ]

/// Style paragraph properties in CT_PPr order: keepNext, keepLines,
/// pageBreakBefore, widowControl, shd, spacing, ind, jc, outlineLvl.
let private styleParagraphProps style =
    let styleToggle name value =
        match value with
        | Some isOn -> [toggle name isOn]
        | None -> []
    [
        yield! styleToggle "w:keepNext" style.KeepNext
        yield! styleToggle "w:keepLines" style.KeepLines
        yield! styleToggle "w:pageBreakBefore" style.PageBreakBefore
        yield! styleToggle "w:widowControl" style.WidowControl
        match InverseConversion.cssColorToWord style.BackgroundColor with
        | Some color -> yield selfClosing "w:shd" [("w:val", "clear"); ("w:color", "auto"); ("w:fill", color)]
        | None -> ()
        let spacingAttrs =
            [
                match style.SpacingBefore with | Some value -> yield ("w:before", InverseConversion.pointsToTwips value) | None -> ()
                match style.SpacingAfter with | Some value -> yield ("w:after", InverseConversion.pointsToTwips value) | None -> ()
                // Styles.styleToObj reads exact/at-least as px (twips / 15) and auto as a multiple (line / 240).
                match style.LineSpacing, style.LineRule with
                | Some value, Some rule when rule = "exact" || rule = "atLeast" ->
                    yield ("w:line", InverseConversion.pixelsToTwips value)
                    yield ("w:lineRule", rule)
                | Some value, _ -> yield ("w:line", InverseConversion.lineSpacingToWord value)
                | None, _ -> ()
            ]
        if not spacingAttrs.IsEmpty then yield selfClosing "w:spacing" spacingAttrs
        let indentAttrs =
            [
                match style.IndentLeft with | Some value -> yield ("w:left", InverseConversion.pixelsToTwips value) | None -> ()
                match style.IndentRight with | Some value -> yield ("w:right", InverseConversion.pixelsToTwips value) | None -> ()
                match style.FirstLineIndent with
                | Some value when value < 0.0 -> yield ("w:hanging", InverseConversion.pixelsToTwips -value)
                | Some value -> yield ("w:firstLine", InverseConversion.pixelsToTwips value)
                | None -> ()
            ]
        if not indentAttrs.IsEmpty then yield selfClosing "w:ind" indentAttrs
        match style.Alignment with
        | Some value when value <> "" -> yield selfClosing "w:jc" [("w:val", InverseConversion.alignmentToWord value)]
        | _ -> ()
        match style.OutlineLevel with | Some value -> yield selfClosing "w:outlineLvl" [("w:val", string value)] | None -> ()
    ]

let private defaultStyleId = documentDefaultsStyleId

let private docDefaultsFromStyle (style: StyleInfo option) =
    let defaultRunStyle =
        style
        |> Option.defaultValue { emptyStyleInfo defaultStyleId "defaults" with FontFamily = Some "Calibri"; FontSize = Some defaultFontSizePixels }
    let runProps = styleRunProps defaultRunStyle
    let rPr = element "w:rPr" [] runProps
    let rPrDefault = element "w:rPrDefault" [] [rPr]
    let pPr = style |> Option.map styleParagraphProps |> Option.defaultValue []
    let pPrDefault = if pPr.IsEmpty then element "w:pPrDefault" [] [] else element "w:pPrDefault" [] [element "w:pPr" [] pPr]
    element "w:docDefaults" [] [rPrDefault; pPrDefault]

let private styleElement (style: StyleInfo) : string =
            let children = [
                yield selfClosing "w:name" [("w:val", style.Name |> Option.defaultValue style.StyleId)]
                match style.Aliases with
                | Some aliases when aliases <> "" -> yield selfClosing "w:aliases" [("w:val", aliases)]
                | _ -> ()
                match style.BasedOn with
                | Some basedOn when basedOn <> "" -> yield selfClosing "w:basedOn" [("w:val", basedOn)]
                | _ -> ()
                match style.Next with
                | Some next when next <> "" -> yield selfClosing "w:next" [("w:val", next)]
                | _ -> ()
                match style.Linked with
                | Some linked when linked <> "" -> yield selfClosing "w:link" [("w:val", linked)]
                | _ -> ()
                match style.UiPriority with
                | Some priority -> yield selfClosing "w:uiPriority" [("w:val", string priority)]
                | None -> ()
                if style.SemiHidden then yield selfClosing "w:semiHidden" []
                if style.UnhideWhenUsed then yield selfClosing "w:unhideWhenUsed" []
                if style.IsQuickStyle then yield selfClosing "w:qFormat" []
                let pPr = styleParagraphProps style
                if not pPr.IsEmpty then yield element "w:pPr" [] pPr
                let rPr = styleRunProps style
                if not rPr.IsEmpty then yield element "w:rPr" [] rPr
            ]
            let attrs = [("w:type", style.StyleType); ("w:styleId", style.StyleId)] @ (if style.IsDefault then [("w:default", "1")] else [])
            element "w:style" attrs children

let generateDefaultStylesWithStyles (styles: StyleInfo list) : string =
    let docDefaults = styles |> List.tryFind (fun style -> style.StyleType = "defaults") |> docDefaultsFromStyle
    let customStyles =
        styles
        |> List.filter (fun style -> style.StyleType <> "defaults")
        |> List.map styleElement
    let root = element "w:styles" [("xmlns:w", Namespaces.word); ("xmlns:r", Namespaces.relationships)] (docDefaults :: customStyles)
    document root

let private rootStartTag (name: string) (attrs: Map<string, string>) =
    let attributes = attrs |> Map.toList |> List.map (fun (key, value) -> sprintf " %s=\"%s\"" key (escapeXmlAttribute value)) |> String.concat ""
    "<" + name + attributes + ">"

/// styles.xml of a document that was opened from a file: the original part
/// with the styles the editor changed (or added) replaced by their current
/// definition. Everything the editor has no model for stays as it was: table
/// styles, latent styles, numbering-linked styles, and whatever an unchanged
/// style says beyond the properties the editor knows (the style part used to be
/// rewritten from the editor's catalog, which reduced every style to those).
/// `changed` holds "type:id" keys; None = the caller does not know, all
/// catalog styles are taken as changed.
let private mergeStyles (originalXml: string) (changed: Set<string> option) (styles: StyleInfo list) : string =
    match XmlDsl.parse originalXml with
    | XmlDsl.Element (rootName, rootAttrs, children, _) ->
        let isChanged (style: StyleInfo) =
            match changed with
            | Some keys -> Set.contains (style.StyleType + ":" + style.StyleId) keys
            | None -> true
        let localName node = XmlDsl.Xml.name node |> Option.map XmlDsl.Ns.stripNs
        let styleKey node =
            let attr name = XmlDsl.Xml.attr ("w:" + name) node |> Option.orElse (XmlDsl.Xml.attr name node)
            (attr "type" |> Option.defaultValue "paragraph"), (attr "styleId" |> Option.defaultValue "")
        let present =
            children
            |> List.choose (fun child -> if localName child = Some "style" then Some (styleKey child) else None)
            |> Set.ofList
        let hasDefaults = children |> List.exists (fun child -> localName child = Some "docDefaults")
        let defaults = styles |> List.tryFind (fun style -> style.StyleType = "defaults")
        // Nothing was changed and nothing is new: the part as it was read (writing
        // it out again from its nodes reordered the attributes of every style).
        let untouched =
            (changed |> Option.exists Set.isEmpty)
            && (hasDefaults || defaults.IsNone)
            && (styles |> List.forall (fun style -> style.StyleType = "defaults" || Set.contains (style.StyleType, style.StyleId) present))
        if untouched then originalXml else
        let body =
            [ // No docDefaults in the original: the catalog's, if it has any.
              if not hasDefaults then
                  match defaults with
                  | Some _ -> yield docDefaultsFromStyle defaults
                  | None -> ()
              for child in children do
                  match localName child with
                  | Some "docDefaults" ->
                      match defaults with
                      | Some style when isChanged style -> yield docDefaultsFromStyle defaults
                      | _ -> yield DocxReader.Docx.xmlToString child
                  | Some "style" ->
                      let styleType, styleId = styleKey child
                      match styles |> List.tryFind (fun style -> style.StyleType = styleType && style.StyleId = styleId) with
                      | Some style when isChanged style -> yield styleElement style
                      | _ -> yield DocxReader.Docx.xmlToString child
                  | Some _ -> yield DocxReader.Docx.xmlToString child
                  | None -> ()
              for style in styles do
                  if style.StyleType <> "defaults" && not (Set.contains (style.StyleType, style.StyleId) present) then
                      yield styleElement style ]
        document (rootStartTag rootName rootAttrs + String.concat "" body + "</" + rootName + ">")
    | _ -> generateDefaultStylesWithStyles styles

let generateDefaultStyles () : string =
    generateDefaultStylesWithStyles []

/// Bullets from Word are usually private-use glyphs of a symbol font; the
/// editor model has no slot for the level font, so name the font the glyph
/// belongs to or Word shows an empty box.
let private bulletFont (levelText: string) =
    if levelText = "o" then Some "Courier New"
    elif levelText.Length = 1 && int levelText.[0] >= 0xF000 && int levelText.[0] <= 0xF0FF then
        if int levelText.[0] = 0xF0B7 then Some "Symbol" else Some "Wingdings"
    else None

/// `withIndent` is for the built-in lists only. Paragraphs of an imported
/// list already carry the level indent as direct formatting (the reader folds
/// it into IndentLeft), so repeating a guessed one here would add an indent to
/// lists whose definition has none.
let private numberingLevelWithIndents (withIndent: bool) (indentLeft: float option) (firstLineIndent: float option) (level: int) (start: int option) (format: string) (levelText: string) =
    let levelIndents =
        [ match indentLeft with
          | Some value -> yield ("w:left", InverseConversion.pixelsToTwips value)
          | None -> ()
          match firstLineIndent with
          | Some value when value < 0.0 -> yield ("w:hanging", InverseConversion.pixelsToTwips -value)
          | Some value -> yield ("w:firstLine", InverseConversion.pixelsToTwips value)
          | None -> () ]
    element "w:lvl" [("w:ilvl", string level)] [
        match start with
        | Some value -> yield selfClosing "w:start" [("w:val", string value)]
        | None -> ()
        yield selfClosing "w:numFmt" [("w:val", format)]
        yield selfClosing "w:lvlText" [("w:val", levelText)]
        yield selfClosing "w:lvlJc" [("w:val", "left")]
        if withIndent then
            yield element "w:pPr" [] [
                selfClosing "w:ind" [("w:left", string (720 * (level + 1))); ("w:hanging", "360")]
            ]
        elif not levelIndents.IsEmpty then
            yield element "w:pPr" [] [selfClosing "w:ind" levelIndents]
        match (if format = "bullet" then bulletFont levelText else None) with
        | Some font -> yield element "w:rPr" [] [selfClosing "w:rFonts" [("w:ascii", font); ("w:hAnsi", font); ("w:hint", "default")]]
        | None -> ()
    ]

let private numberingLevel (withIndent: bool) (level: int) (start: int option) (format: string) (levelText: string) =
    numberingLevelWithIndents withIndent None None level start format levelText

let private numberingLevels (infos: NumberingInfo list) =
    let stated = infos |> List.map (fun info -> info.Level) |> Set.ofList
    [ yield! infos |> List.map (fun info -> info.Level, numberingLevelWithIndents false info.IndentLeft info.FirstLineIndent info.Level info.Start info.Format info.LevelText)
      for KeyValue (level, _) in numberingStarts infos do
          if not (Set.contains level stated) then
              yield level, numberingLevel false level (Some 1) "decimal" ("%" + string (level + 1) + ".") ]
    |> List.sortBy fst |> List.map snd

let private numberingOverrideXml (infos: NumberingInfo list) =
    [ for KeyValue (level, value) in numberingStarts infos ->
        element "w:lvlOverride" [ "w:ilvl", string level ] [ selfClosing "w:startOverride" [ "w:val", string value ] ] ]

let private generateNumbering (ctx: PackageContext) (customInfos: NumberingInfo list) : string =
    let usedIds =
        customInfos
        |> List.choose (fun info -> match System.Int32.TryParse info.NumId with | true, value -> Some value | _ -> None)
    let nextFree = (if usedIds.IsEmpty then 0 else List.max usedIds) + 1
    // Defaults for lists made in the editor, on ids no imported list uses.
    let bulletId, numberedId =
        if usedIds |> List.exists (fun id -> id = 1 || id = 2) then string nextFree, string (nextFree + 1) else "1", "2"
    ctx.Model.DefaultListNumIds <- (bulletId, numberedId)
    let defaultLevels format text =
        [0..8] |> List.map (fun level -> numberingLevel true level (Some 1) format (if format = "bullet" then text else "%" + string (level + 1) + "."))

    let groups = customInfos |> List.filter (fun info -> not info.Undefined) |> List.groupBy (fun info -> info.NumId)
    let abstracts =
        [ yield element "w:abstractNum" [("w:abstractNumId", bulletId)] (defaultLevels "bullet" "\u2022")
          yield element "w:abstractNum" [("w:abstractNumId", numberedId)] (defaultLevels "decimal" "")
          for numId, infos in groups do
              let levels = numberingLevels infos
              yield element "w:abstractNum" [("w:abstractNumId", numId)] levels ]
    let nums =
        [ yield element "w:num" [("w:numId", bulletId)] [selfClosing "w:abstractNumId" [("w:val", bulletId)]]
          yield element "w:num" [("w:numId", numberedId)] [selfClosing "w:abstractNumId" [("w:val", numberedId)]]
          for numId, infos in groups do
              let overrides = numberingOverrideXml infos
              yield element "w:num" [("w:numId", numId)] (selfClosing "w:abstractNumId" [("w:val", numId)] :: overrides) ]

    document (element "w:numbering" [("xmlns:w", Namespaces.word)] (abstracts @ nums))

/// The numbering part as it was read, with definitions added for the lists
/// that were made in the editor (and, when a paragraph written from the model
/// has a list without an id, the two default lists). The definitions of the
/// file keep what the model does not hold of them: the bullet fonts, the
/// alignment and the tab of the number, picture bullets, the levels no
/// paragraph uses yet, the links to numbering styles. None when the part is
/// not what this expects (it is then generated as for a new document).
let private mergeNumbering (ctx: PackageContext) (original: string) (customInfos: NumberingInfo list) (needsDefaults: bool) : string option =
    match XmlPatch.readPart original with
    | Some part when (match part.Node with XmlDsl.Element (name, _, _, _) -> name = "w:numbering" | _ -> false) ->
        let lists = XmlPatch.childrenNamed part "w:num"
        let known = lists |> List.choose (fun num -> XmlPatch.attribute num "w:numId") |> Set.ofList
        // (The highest number any list or definition has, or refers to.)
        let highest =
            [ for num in lists do yield! Option.toList (XmlPatch.attribute num "w:numId")
              for definition in XmlPatch.childrenNamed part "w:abstractNum" do yield! Option.toList (XmlPatch.attribute definition "w:abstractNumId")
              for reference in XmlPatch.descendants "abstractNumId" part.Node do yield! Option.toList (XmlPatch.attribute reference "w:val")
              for reference in XmlPatch.descendants "numId" part.Node do yield! Option.toList (XmlPatch.attribute reference "w:val") ]
            |> List.choose (fun text -> match System.Int32.TryParse text with | true, value -> Some value | _ -> None)
            |> List.fold max 0
        let unknown = customInfos |> List.filter (fun info -> not (Set.contains info.NumId known))
        let added = unknown |> List.filter (fun info -> not info.Undefined) |> List.groupBy (fun info -> info.NumId)
        let addedIds = unknown |> List.groupBy (fun info -> info.NumId) |> List.choose (fun (numId, _) -> match System.Int32.TryParse numId with | true, value -> Some value | _ -> None)
        let mutable nextFree = (List.fold max highest addedIds) + 1
        let fresh () =
            let id = string nextFree
            nextFree <- nextFree + 1
            id
        let abstracts = ResizeArray<string>()
        let nums = ResizeArray<string>()
        if needsDefaults then
            let bulletId, numberedId = fresh (), fresh ()
            ctx.Model.DefaultListNumIds <- (bulletId, numberedId)
            let defaultLevels format text =
                [0..8] |> List.map (fun level -> numberingLevel true level (Some 1) format (if format = "bullet" then text else "%" + string (level + 1) + "."))
            abstracts.Add (element "w:abstractNum" [("w:abstractNumId", bulletId)] (defaultLevels "bullet" "\u2022"))
            abstracts.Add (element "w:abstractNum" [("w:abstractNumId", numberedId)] (defaultLevels "decimal" ""))
            nums.Add (element "w:num" [("w:numId", bulletId)] [selfClosing "w:abstractNumId" [("w:val", bulletId)]])
            nums.Add (element "w:num" [("w:numId", numberedId)] [selfClosing "w:abstractNumId" [("w:val", numberedId)]])
        for numId, infos in added do
            let abstractId = fresh ()
            let source =
                infos |> List.tryPick (fun info -> info.SourceId)
                |> Option.bind (fun sourceId -> lists |> List.tryFind (fun num -> XmlPatch.attribute num "w:numId" = Some sourceId))
            let definition =
                source
                |> Option.bind (XmlPatch.child "abstractNumId")
                |> Option.bind (fun reference -> XmlPatch.attribute reference "w:val")
                |> Option.bind (fun id -> XmlPatch.childrenNamed part "w:abstractNum" |> List.tryFind (fun node -> XmlPatch.attribute node "w:abstractNumId" = Some id))
            match source, definition with
            | Some source, Some definition ->
                let markerEdits info =
                    [ XmlPatch.EditChild ("w:numFmt", [ XmlPatch.SetAttribute ("w:val", info.Format) ], Some (selfClosing "w:numFmt" [ "w:val", info.Format ]))
                      XmlPatch.EditChild ("w:lvlText", [ XmlPatch.SetAttribute ("w:val", info.LevelText) ], Some (selfClosing "w:lvlText" [ "w:val", info.LevelText ]))
                      if info.Format = "bullet" then XmlPatch.EditChild ("w:rPr", [ XmlPatch.RemoveChild "w:rFonts" ], None) ]
                let levelEdits =
                    [ for info in infos do
                        if info.FormatEdited then
                            let level = string info.Level
                            let exists = XmlPatch.childrenNamed { XmlPatch.SourceElement.Text = part.Text; XmlPatch.SourceElement.Node = definition } "w:lvl" |> List.exists (fun node -> XmlPatch.attribute node "w:ilvl" = Some level)
                            if exists then yield XmlPatch.EditChildWhere ("w:lvl", "w:ilvl", level, markerEdits info)
                            else yield XmlPatch.SetChildWhere ("w:lvl", "w:ilvl", level, numberingLevelWithIndents false info.IndentLeft info.FirstLineIndent info.Level info.Start info.Format info.LevelText) ]
                let defXml = XmlPatch.apply XmlPatch.Word.order { XmlPatch.SourceElement.Text = part.Text; XmlPatch.SourceElement.Node = definition }
                                ([ XmlPatch.SetAttribute ("w:abstractNumId", abstractId); XmlPatch.RemoveChild "w:nsid"; XmlPatch.RemoveChild "w:tmpl" ] @ levelEdits)
                let overrides = XmlPatch.childrenNamed { XmlPatch.SourceElement.Text = part.Text; XmlPatch.SourceElement.Node = source } "w:lvlOverride"
                let listEdits =
                    [ yield XmlPatch.SetAttribute ("w:numId", numId)
                      yield XmlPatch.EditChild ("w:abstractNumId", [ XmlPatch.SetAttribute ("w:val", abstractId) ], None)
                      let starts = numberingStarts infos
                      let levels =
                          Set.union (starts |> Map.toSeq |> Seq.map fst |> Set.ofSeq)
                                    (infos |> List.filter (fun info -> info.FormatEdited) |> List.map (fun info -> info.Level) |> Set.ofList)
                      for levelNumber in levels do
                          let level = string levelNumber
                          let existing = overrides |> List.tryFind (fun node -> XmlPatch.attribute node "w:ilvl" = Some level)
                          let format = infos |> List.tryFind (fun info -> info.Level = levelNumber && info.FormatEdited)
                          let start = Map.tryFind levelNumber starts
                          let edits =
                              [ match start with
                                | Some value -> yield XmlPatch.EditChild ("w:startOverride", [ XmlPatch.SetAttribute ("w:val", string value) ], Some (selfClosing "w:startOverride" [ "w:val", string value ]))
                                | None -> ()
                                match format with
                                | Some info when (existing |> Option.bind (XmlPatch.child "lvl")).IsSome -> yield XmlPatch.EditChild ("w:lvl", markerEdits info, None)
                                | _ -> () ]
                          if not edits.IsEmpty then
                              match existing, start with
                              | Some _, _ -> yield XmlPatch.EditChildWhere ("w:lvlOverride", "w:ilvl", level, edits)
                              | None, Some value -> yield XmlPatch.AddChild ("w:lvlOverride", element "w:lvlOverride" [ "w:ilvl", level ] [ selfClosing "w:startOverride" [ "w:val", string value ] ])
                              | None, None -> () ]
                abstracts.Add defXml
                nums.Add (XmlPatch.apply XmlPatch.Word.order { XmlPatch.SourceElement.Text = part.Text; XmlPatch.SourceElement.Node = source } listEdits)
            | _ ->
                let levels = numberingLevels infos
                let overrides = numberingOverrideXml infos
                abstracts.Add (element "w:abstractNum" [ "w:abstractNumId", abstractId ] levels)
                nums.Add (element "w:num" [ "w:numId", numId ] (selfClosing "w:abstractNumId" [ "w:val", abstractId ] :: overrides))
        if abstracts.Count = 0 then Some original
        else
            // (CT_Numbering: the definitions, then the lists - each put in after those of its kind.)
            Some (XmlPatch.applyPart XmlPatch.Word.order part
                    ([ for xml in abstracts -> XmlPatch.AddChild ("w:abstractNum", xml) ] @ [ for xml in nums -> XmlPatch.AddChild ("w:num", xml) ]))
    | _ -> if not needsDefaults && customInfos.IsEmpty then Some original else None

/// The numbering part as it was read with the levels the editor changed
/// ("list id/level") written anew: in the list's own abstract definition (a
/// copy of it, made for the list, when other lists share the definition), the
/// other levels, lists and definitions staying the text they were. None when
/// a list or its definition is not found in the text.
let private patchNumberingLevels (original: string) (customInfos: NumberingInfo list) (changed: string array) : string option =
    let mutable text = Some original
    for key in changed do
        match text |> Option.bind XmlPatch.readPart with
        | None -> text <- None
        | Some part ->
            // (The editor counts levels from one, the file from zero.)
            let numId, level = ModelText.numberingLevel key
            let lists = XmlPatch.childrenNamed part "w:num"
            let list = lists |> List.tryFind (fun num -> XmlPatch.attribute num "w:numId" = Some numId)
            let abstractId = list |> Option.bind (XmlPatch.child "abstractNumId") |> Option.bind (fun reference -> XmlPatch.attribute reference "w:val")
            let definition = abstractId |> Option.bind (fun id -> XmlPatch.childrenNamed part "w:abstractNum" |> List.tryFind (fun node -> XmlPatch.attribute node "w:abstractNumId" = Some id))
            match customInfos |> List.tryFind (fun info -> info.NumId = numId && info.Level = level && not info.Undefined), abstractId, definition with
            | Some info, Some abstractId, Some definition ->

                let shared = (lists |> List.filter (fun num -> (XmlPatch.child "abstractNumId" num |> Option.bind (fun reference -> XmlPatch.attribute reference "w:val")) = Some abstractId)).Length > 1
                let highest =
                    [ for num in lists do yield! Option.toList (XmlPatch.attribute num "w:numId")
                      for node in XmlPatch.childrenNamed part "w:abstractNum" do yield! Option.toList (XmlPatch.attribute node "w:abstractNumId") ]
                    |> List.choose (fun value -> match System.Int32.TryParse value with | true, number -> Some number | _ -> None)
                    |> List.fold max 0
                let ownId = if shared then string (highest + 1) else abstractId
                // Change only the marker fields, retaining unknown level data and overrides.
                let markerEdits =
                    [ XmlPatch.EditChild ("w:numFmt", [ XmlPatch.SetAttribute ("w:val", info.Format) ], Some (selfClosing "w:numFmt" [ "w:val", info.Format ]))
                      XmlPatch.EditChild ("w:lvlText", [ XmlPatch.SetAttribute ("w:val", info.LevelText) ], Some (selfClosing "w:lvlText" [ "w:val", info.LevelText ])) ]
                let sourceLevels = XmlPatch.childrenNamed { XmlPatch.SourceElement.Text = part.Text; XmlPatch.SourceElement.Node = definition } "w:lvl"
                let levelEdit =
                    if sourceLevels |> List.exists (fun node -> XmlPatch.attribute node "w:ilvl" = Some (string level)) then
                        XmlPatch.EditChildWhere ("w:lvl", "w:ilvl", string level, markerEdits)
                    else
                        XmlPatch.SetChildWhere ("w:lvl", "w:ilvl", string level, numberingLevelWithIndents false info.IndentLeft info.FirstLineIndent info.Level info.Start info.Format info.LevelText)
                let listEdits =
                    [ if shared then yield XmlPatch.EditChild ("w:abstractNumId", [ XmlPatch.SetAttribute ("w:val", ownId) ], None)
                      match list with
                      | Some source ->
                          let overrides = XmlPatch.childrenNamed { XmlPatch.SourceElement.Text = part.Text; XmlPatch.SourceElement.Node = source } "w:lvlOverride"
                          if overrides |> List.exists (fun node -> XmlPatch.attribute node "w:ilvl" = Some (string level) && (XmlPatch.child "lvl" node).IsSome) then
                              yield XmlPatch.EditChildWhere ("w:lvlOverride", "w:ilvl", string level, [ XmlPatch.EditChild ("w:lvl", markerEdits, None) ])
                      | None -> () ]
                let definitionEdits =
                    if shared then
                        // A definition of its own for this list: with an id and a signature of its own.
                        let own = XmlPatch.apply XmlPatch.Word.order { XmlPatch.SourceElement.Text = part.Text; XmlPatch.SourceElement.Node = definition } [ XmlPatch.SetAttribute ("w:abstractNumId", ownId); XmlPatch.RemoveChild "w:nsid"; XmlPatch.RemoveChild "w:tmpl"; levelEdit ]
                        [ XmlPatch.AddChild ("w:abstractNum", own) ]
                    else [ XmlPatch.EditChildWhere ("w:abstractNum", "w:abstractNumId", abstractId, [ levelEdit ]) ]
                text <- Some (XmlPatch.applyPart XmlPatch.Word.order part (definitionEdits @ [ XmlPatch.EditChildWhere ("w:num", "w:numId", numId, listEdits) ]))
            | _ -> text <- None
    text

let private generateSettings (pageSetup: PageSetupInfo option) =
    let children = [
        match pageSetup with
        | Some setup when setup.EvenAndOddHeaders -> yield selfClosing "w:evenAndOddHeaders" []
        | _ -> ()
    ]
    document (element "w:settings" [("xmlns:w", Namespaces.word)] children)

/// Children of w:settings that come before w:evenAndOddHeaders in CT_Settings; every other comes after it.
let private settingsOrder : XmlPatch.SchemaOrder = fun _ ->
    [ "writeProtection"; "view"; "zoom"; "removePersonalInformation"; "removeDateAndTime"; "doNotDisplayPageBoundaries"
      "displayBackgroundShape"; "printPostScriptOverText"; "printFractionalCharacterWidth"; "printFormsData"; "embedTrueTypeFonts"
      "embedSystemFonts"; "saveSubsetFonts"; "saveFormsData"; "mirrorMargins"; "alignBordersAndEdges"; "bordersDoNotSurroundHeader"
      "bordersDoNotSurroundFooter"; "gutterAtTop"; "hideSpellingErrors"; "hideGrammaticalErrors"; "activeWritingStyle"; "proofState"
      "formsDesign"; "attachedTemplate"; "linkStyles"; "stylePaneFormatFilter"; "stylePaneSortMethod"; "documentType"; "mailMerge"
      "revisionView"; "trackRevisions"; "doNotTrackMoves"; "doNotTrackFormatting"; "documentProtection"; "autoFormatOverride"
      "styleLockTheme"; "styleLockQFSet"; "defaultTabStop"; "autoHyphenation"; "consecutiveHyphenLimit"; "hyphenationZone"
      "doNotHyphenateCaps"; "showEnvelope"; "summaryLength"; "clickAndTypeStyle"; "defaultTableStyle"; "evenAndOddHeaders"; "*" ]

/// The original settings part with the one setting the editor owns, even and
/// odd headers, brought up to date. Everything else (compatibility options,
/// default tab stop, track-revisions flag, footnote properties, rsids, ...)
/// stays: the part used to be dropped, or replaced by one holding this flag only.
let private patchSettings (originalXml: string) (pageSetup: PageSetupInfo option) : string =
    match XmlPatch.readPart originalXml with
    | Some part ->
        let evenAndOdd = pageSetup |> Option.exists (fun setup -> setup.EvenAndOddHeaders)
        let setting = XmlPatch.elements part.Node |> List.tryFind (fun node -> XmlPatch.localName node = "evenAndOddHeaders")
        // The setting is what the part says already: the part as it was read.
        let hadEvenAndOdd =
            setting |> Option.exists (fun node -> match XmlPatch.attribute node "w:val" with Some "0" | Some "false" | Some "off" -> false | _ -> true)
        if hadEvenAndOdd = evenAndOdd then originalXml
        else
            let name = match setting with Some node -> XmlPatch.name node | None -> XmlDsl.Ns.withPrefixOf (XmlPatch.name part.Node) "evenAndOddHeaders"
            XmlPatch.applyPart settingsOrder part
                [ if evenAndOdd then XmlPatch.SetChild (name, selfClosing name []) else XmlPatch.RemoveChild name ]
    | None -> generateSettings pageSetup


// ---- relationship parts and [Content_Types].xml as they were read -------------
// The writer makes these parts anew for every save: its own order, its own
// ids for the parts nothing refers to by id (the styles, the lists, the
// settings, the notes, the comments), absolute targets. A part that was read
// and says what the new one says is written back as the text it was.

/// Relationship types of the parts a document has one of, which nothing refers to by id.
let private isImplicitRelationshipType (relType: string) =
    [ "styles"; "numbering"; "settings"; "comments"; "footnotes"; "endnotes" ] |> List.contains (OpcXml.relationshipKind relType)

/// What a relationships part says: one entry per relationship (its id - but
/// for the parts nothing refers to by id -, type, mode and target, a target
/// in the package as its path), sorted. None when it is not well formed.
let private relationshipEntries (ownerDirectory: string) (xml: string) : string list option =
    OpcXml.relationships xml
    |> Option.map (fun rels ->
        rels
        |> List.map (fun rel ->
            (if isImplicitRelationshipType rel.Type then "" else rel.Id) + "\n" + rel.Type + "\n" + (if rel.External then "external" else "internal") + "\n" + OpcXml.targetPath ownerDirectory rel)
        |> List.sort)

/// Whether the [Content_Types].xml that was read is right for the package
/// that is written: every part has the type the writer gives it, and it names
/// no part that is not written.
let private contentTypesStillRight (original: string) (generated: string) (writtenPaths: string list) =
    let originalType, originalOverrides = OpcXml.contentTypeLookup original
    let generatedType, _ = OpcXml.contentTypeLookup generated
    let written = writtenPaths |> List.map (fun path -> ("/" + path).ToLowerInvariant()) |> Set.ofList
    writtenPaths |> List.forall (fun path -> path = "[Content_Types].xml" || (match originalType path with Some contentType -> generatedType path = Some contentType | None -> false))
    && originalOverrides |> List.forall (fun name -> Set.contains name written)


/// Generate all parts with source state scoped to this save.
let private preparePackage (ctx: PackageContext) (package: EditorPackage) =
    configureSources ctx
    ctx.Model.NextShapeId <- 1000
    // What was kept of the opened package: written back around the parts the
    // editor owns. A package without it (a new document) is written from scratch.
    let preserved = package.Preserved |> Option.defaultValue emptyPreservedPackage
    // RegenerateAll: every part is written from the model, as it was before
    // unchanged content was written back as it was read (the round-trip
    // checks use it to see what the model alone gives).
    let regenerateAll = package.RegenerateAll
    ctx.Model.ReportPatched <- 0
    ctx.ReportInline <- 0
    ctx.ReportModelBlocks <- 0
    ctx.ReportModelParagraphs <- 0
    ctx.ReportSourceBlocks <- 0
    ctx.ReportReasons.Clear()
    ctx.Model.ReportGenerated <- 0
    ctx.Model.ReportDerived <- 0
    ctx.Model.ReportSpliced <- 0
    ctx.Model.LiveCommentIds <- Set.empty
    forgetSourceBody ctx ()
    ctx.Model.ShapesFromModel <- regenerateAll
    let sourceOf (key: string) = if regenerateAll then None else Map.tryFind key preserved.SourceXml
    let paragraphs = package.Blocks
    // The body's elements that go back as they were read (see planBody).
    let bodyPlan =
        match sourceOf "document" with
        | Some xml when preserved.BodyItemChunks.Length > 0 ->
            planBody ctx DocumentBody xml preserved.BodyItemChunks package.SourceStoryCounts preserved.BodyChunkCount paragraphs package.PageSetupUnchanged
        | _ -> None
    let bodyState = sourceState ctx ()
    // A header or footer that was changed: what of it is as it was read goes back as it was, as the body's
    // blocks do (its plan, and what is known of its text while it is written).
    let partPlans =
        Array.append package.HeaderParts package.FooterParts
        |> Array.choose (fun part ->
            match part.RelationshipId with
            | Some relId when not part.SourceUnchanged ->
                match sourceOf relId, Map.tryFind relId preserved.PartItemChunks with
                | Some xml, Some (chunks, count) when chunks.Length > 0 ->
                    forgetSourceBody ctx ()
                    let counts = part.SourceStoryCounts
                    planBody ctx PartRoot xml chunks counts count part.Blocks false
                    |> Option.map (fun plan -> relId, (plan, sourceState ctx ()))
                | _ -> None
            | _ -> None)
        |> Map.ofArray
    forgetSourceBody ctx ()
    useSourceState ctx bodyState
    /// The parts that are written back as they were read keep all the
    /// relationships they had; a part written from the model keeps those the
    /// writer does not make itself.
    let sourceParts = System.Collections.Generic.HashSet<string>()
    if bodyPlan.IsSome then sourceParts.Add "document" |> ignore
    for KeyValue (relId, _) in partPlans do sourceParts.Add relId |> ignore
    for part in Array.append package.HeaderParts package.FooterParts do
        match part.RelationshipId with
        | Some relId when part.SourceUnchanged && (sourceOf relId).IsSome -> sourceParts.Add relId |> ignore
        | _ -> ()
    // (A part of notes or comments that was read keeps them whether it is as it was or not: the notes in it
    // that are what they were are written back as they were - see patchedStoriesPart.)
    for key in [ "footnotes"; "endnotes"; "comments" ] do
        if (sourceOf key).IsSome then sourceParts.Add key |> ignore
    let keptFor (key: string) =
        let ofMap (map: Map<string, PreservedRelationshipInfo list>) = map |> Map.tryFind key |> Option.defaultValue []
        if sourceParts.Contains key then
            let all = ofMap preserved.SourceRelationships
            let known = all |> List.map (fun rel -> rel.Id) |> Set.ofList
            all @ (ofMap preserved.PartRelationships |> List.filter (fun rel -> not (Set.contains rel.Id known)))
        else ofMap preserved.PartRelationships
    // (The ids and the pictures of what is written back as it was read are taken.)
    ctx.Model.ReservedRelationshipIds <-
        (Map.toList preserved.PartRelationships @ (Map.toList preserved.SourceRelationships |> List.filter (fun (key, _) -> sourceParts.Contains key)))
        |> List.collect (fun (_, rels) -> rels |> List.map (fun rel -> rel.Id))
        |> Set.ofList
    ctx.Model.ReservedPartPaths <- (preserved.Parts @ (if sourceParts.Count > 0 then preserved.SourceParts else [])) |> List.map (fun part -> part.Path) |> Set.ofList
    ctx.Model.NextMoveRangeId <- 900000
    let withRelationshipsOf (key: string) (blocks: EditorBlock array) (generate: unit -> string) =
        ctx.Model.CurrentPartRelationshipIds <- keptFor key |> List.map (fun rel -> rel.Id) |> Set.ofList
        // The runs of each tracked move in this part, for its range markers.
        blocks
        |> Seq.collect collectItemParagraphs
        |> Seq.collect (fun para -> para.Runs)
        |> resetMoveRanges ctx.Model
        let xml = generate ()
        ctx.Model.CurrentPartRelationshipIds <- Set.empty
        xml
    let pageSetup = optionalPageSetup package.PageSetup
    let hasSettings = preserved.SettingsXml.IsSome || (pageSetup |> Option.exists (fun setup -> setup.EvenAndOddHeaders))
    // Relationship ids some section refers to; the parts behind the others are
    // attached to the final section.
    let sectionReferenceIds =
        let finalIds =
            pageSetup
            |> Option.map (fun setup -> (setup.HeaderReferences @ setup.FooterReferences) |> List.map snd)
            |> Option.defaultValue []
        Set.union (collectSectionReferenceIds paragraphs) (Set.ofList finalIds)
    let headerParts = package.HeaderParts |> headerFooterInfos ctx.Model preserved.PartPaths sectionReferenceIds "header" 1
    let footerParts = package.FooterParts |> headerFooterInfos ctx.Model preserved.PartPaths sectionReferenceIds "footer" 1
    // A header / footer that is what it was when it was read goes back as its text.
    let partSource (part: HeaderFooterPartInfo) = if part.Unchanged then sourceOf part.RelId else None
    // The blocks of the body that are written from the model: their pictures
    // and hyperlinks are the ones the writer makes relationships for.
    let bodyModelBlocks = bodyPlan |> Option.map (fun plan -> plan.Regenerated) |> Option.defaultValue paragraphs
    let bodyImages, nextAfterBody = buildImageScope ctx.Model 1 bodyModelBlocks
    let mutable nextMediaIndex = nextAfterBody
    let partPackage (part: HeaderFooterPartInfo) =
        match partSource part, Map.tryFind part.RelId partPlans with
        | Some _, _ -> { Info = part; Images = emptyImageScope; Hyperlinks = Map.empty }
        | None, Some (plan, _) ->
            // (Its pictures and links are those of what is written from the model; the others keep theirs.)
            let images, next = buildImageScope ctx.Model nextMediaIndex plan.Regenerated
            nextMediaIndex <- next
            { Info = part; Images = images; Hyperlinks = hyperlinkMapWith ctx.Model (keptFor part.RelId) plan.Regenerated }
        | None, None ->
            let images, next = buildImageScope ctx.Model nextMediaIndex part.Blocks
            nextMediaIndex <- next
            { Info = part; Images = images; Hyperlinks = hyperlinkMap ctx.Model part.Blocks }
    let headerPackages = headerParts |> Array.map partPackage
    let footerPackages = footerParts |> Array.map partPackage
    let contentBlocks =
        [| yield! paragraphs
           for part in headerParts do yield! part.Blocks
           for part in footerParts do yield! part.Blocks |]
    // (A link the body had a relationship for keeps it: see hyperlinkMapWith.)
    let hyperlinks = hyperlinkMapWith ctx.Model (if sourceParts.Contains "document" then keptFor "document" else []) bodyModelBlocks
    // The links the relationships part gets an entry of its own for: the others are among the kept ones.
    let newBodyHyperlinks =
        let keptIds = keptFor "document" |> List.map (fun rel -> rel.Id) |> Set.ofList
        hyperlinks |> Map.filter (fun _ id -> not (Set.contains id keptIds))
    // Note and comment bodies come with the package; a reference whose body
    // is not there is written from the text on the reference run.
    let comments = mergeCommentParts (commentPartsFromPackage package.Comments) (collectComments contentBlocks)
    ctx.Model.LiveCommentIds <- comments |> List.map (fun comment -> comment.Id) |> Set.ofList
    ctx.Model.DeletedCommentIds <-
        match sourceOf "comments" |> Option.bind XmlPatch.readPart with
        | Some part -> XmlPatch.childrenNamed part "w:comment" |> List.choose (fun comment -> XmlPatch.attribute comment "w:id") |> Set.ofList |> fun read -> Set.difference read ctx.Model.LiveCommentIds
        | None -> Set.empty
    let footnotes = mergeNoteParts (notePartsFromPackage package.Footnotes) (collectNotes "footnote" contentBlocks)
    let endnotes = mergeNoteParts (notePartsFromPackage package.Endnotes) (collectNotes "endnote" contentBlocks)
    // A note or comment that was changed: what of it is as it was read goes back as it was, as the body's blocks
    // do (its plan, and what is known of its part while it is written), by its part and id ("footnotes:2").
    let notePlans =
        let planOf (key: string) (elementName: string) (id: string) (unchanged: bool) (blocks: EditorBlock array) (counts: int array) =
            if unchanged || blocks.Length = 0 then None
            else
                match sourceOf key, Map.tryFind (key + ":" + id) preserved.NoteItemChunks with
                | Some xml, Some (chunks, count) when chunks.Length > 0 ->
                    forgetSourceBody ctx ()
                    planBody ctx (StoryElement (elementName, id)) xml chunks counts count blocks false
                    |> Option.map (fun plan -> key + ":" + id, (plan, sourceState ctx ()))
                | _ -> None
        [ for note in footnotes do yield planOf "footnotes" "w:footnote" note.Id note.Unchanged note.Blocks note.SourceStoryCounts
          for note in endnotes do yield planOf "endnotes" "w:endnote" note.Id note.Unchanged note.Blocks note.SourceStoryCounts
          for comment in comments do yield planOf "comments" "w:comment" comment.Id comment.Unchanged comment.Blocks comment.SourceStoryCounts ]
        |> List.choose id
        |> Map.ofList
    forgetSourceBody ctx ()
    useSourceState ctx bodyState
    // (Of a note written from its plan, the blocks written from the model.)
    let storyBlocks (key: string) (id: string) (blocks: EditorBlock array) =
        match Map.tryFind (key + ":" + id) notePlans with
        | Some (plan, _) -> plan.Regenerated
        | None -> blocks
    // Each of these parts has its own relationships: its images and hyperlinks.
    let bodyPart (key: string) (blocks: EditorBlock array) =
        let images, next = buildImageScope ctx.Model nextMediaIndex blocks
        nextMediaIndex <- next
        images, hyperlinkMapWith ctx.Model (keptFor key) blocks
    // (Of a part that was read, those of the notes that changed: the others are written back as they were.)
    let readFor (key: string) = (sourceOf key).IsSome
    let footnoteBlocks = footnotes |> List.filter (fun note -> not (readFor "footnotes" && note.Unchanged)) |> List.toArray |> Array.collect (fun note -> storyBlocks "footnotes" note.Id note.Blocks)
    let endnoteBlocks = endnotes |> List.filter (fun note -> not (readFor "endnotes" && note.Unchanged)) |> List.toArray |> Array.collect (fun note -> storyBlocks "endnotes" note.Id note.Blocks)
    let commentBlocks = comments |> List.filter (fun comment -> not (readFor "comments" && comment.Unchanged)) |> List.toArray |> Array.collect (fun comment -> storyBlocks "comments" comment.Id comment.Blocks)
    // The notes and the comments as they were read, when nothing of them was changed.
    let footnotesSource = if package.FootnotesUnchanged then sourceOf "footnotes" else None
    let endnotesSource = if package.EndnotesUnchanged then sourceOf "endnotes" else None
    let commentsSource = if package.CommentsUnchanged then sourceOf "comments" else None
    let footnoteImages, footnoteLinks = if footnotesSource.IsSome then emptyImageScope, Map.empty else bodyPart "footnotes" footnoteBlocks
    let endnoteImages, endnoteLinks = if endnotesSource.IsSome then emptyImageScope, Map.empty else bodyPart "endnotes" endnoteBlocks
    let commentImages, commentLinks = if commentsSource.IsSome then emptyImageScope, Map.empty else bodyPart "comments" commentBlocks
    let partBody (images: ImageScope) (links: Map<string, string>) (blocks: EditorBlock array) =
        blocks |> generateBlocks (generateDocumentItemWithImages ctx.Model links images) (generateTextBoxRun ctx.Model defaultPageSetupInfo links images)
    let allImages =
        [| yield! bodyImages.Assets
           for part in headerPackages do yield! part.Images.Assets
           for part in footerPackages do yield! part.Images.Assets
           yield! footnoteImages.Assets
           yield! endnoteImages.Assets
           yield! commentImages.Assets |]
    let allBlocks =
        [| yield! contentBlocks
           // (All of them: the lists and styles of those written back as they were are the document's too.)
           for note in footnotes do yield! note.Blocks
           for note in endnotes do yield! note.Blocks
           for comment in comments do yield! comment.Blocks |]
    // Rendered first: it picks the default list ids the paragraph writer uses.
    let numberingXml =
        let infos = collectNumberingInfoOf package.Preserved.IsSome allBlocks
        // Blocks written from the model with a list that has no id of the file.
        let modelBlocks =
            [| yield! bodyModelBlocks
               for part in Array.append headerParts footerParts do
                   match partSource part, Map.tryFind part.RelId partPlans with
                   | Some _, _ -> ()
                   | None, Some (plan, _) -> yield! plan.Regenerated
                   | None, None -> yield! part.Blocks
               if footnotesSource.IsNone then yield! footnoteBlocks
               if endnotesSource.IsNone then yield! endnoteBlocks
               if commentsSource.IsNone then yield! commentBlocks |]
        let needsDefaults =
            modelBlocks
            |> Array.toList
            |> List.collect collectItemParagraphs
            |> List.exists (fun para ->
                para.List |> Option.exists (fun list ->
                    (list.Style = "bullet" || list.Style = "numbered") && list.Id.IsNone))
        let merged =
            // (Also with RegenerateAll: the lists are not blocks, and they are as they were.)
            if package.ListsUnchanged then
                Map.tryFind "numbering" preserved.SourceXml |> Option.bind (fun original -> mergeNumbering ctx original infos needsDefaults)
            else
                // Lists whose format the editor changed: those levels alone are written anew.
                let changed = package.ChangedListLevels
                if changed.Length = 0 then None
                else
                    Map.tryFind "numbering" preserved.SourceXml
                    |> Option.bind (fun original -> mergeNumbering ctx original infos needsDefaults)
                    |> Option.bind (fun merged -> patchNumberingLevels merged infos changed)
        // (Rendered also when it is not written: it picks the default list ids.)
        let generated = lazy (generateNumbering ctx infos)
        match merged with
        | Some text -> Some text
        | None ->
            let text = generated.Force()
            // A document that was opened without a part for its lists, and has no list, is saved without one.
            // (Nor is one made for lists the document refers to and never had a definition of.)
            if package.Preserved.IsSome && not (Map.containsKey "numbering" preserved.SourceXml)
               && (infos |> List.forall (fun info -> info.Undefined)) && not needsDefaults then None
            else Some text
    let hasNumbering = numberingXml.IsSome
    let styles =
        (collectPackageStyles package.Styles @ collectStyles allBlocks)
        |> List.distinctBy (fun style -> style.StyleType, style.StyleId)
        // The document defaults are not a style: a style of that name (inferred by
        // the editor from table cell paragraphs) would take their place when read.
        |> List.filter (fun style -> style.StyleId <> documentDefaultsStyleId || style.StyleType = "defaults")
    let hasComments = not comments.IsEmpty || commentsSource.IsSome
    let hasFootnotes = not footnotes.IsEmpty || footnotesSource.IsSome
    let hasEndnotes = not endnotes.IsEmpty || endnotesSource.IsSome
    // Pictures and other parts that only content written back as it was read refers to.
    let preserved = if sourceParts.Count > 0 then { preserved with Parts = preserved.Parts @ preserved.SourceParts } else preserved
    // A header or footer: from its plan (what is as it was read stays its text), or from the model.
    let writeHeaderFooter (rootName: string) (part: HeaderFooterPartPackage) =
        match Map.tryFind part.Info.RelId partPlans with
        | Some (plan, state) ->
            useSourceState ctx state
            let text = withRelationshipsOf part.Info.RelId plan.Regenerated (fun () -> generateDocumentXmlFromPlan ctx plan part.Images part.Hyperlinks None None [||] [||])
            forgetSourceBody ctx ()
            text
        | None -> withRelationshipsOf part.Info.RelId part.Info.Blocks (fun () -> balancedPart (generateHeaderFooterXml ctx.Model rootName part.Hyperlinks part.Images part.Info.Blocks))
    // (The links of a part that kept the relationships it had: those it had are among the kept ones.)
    let newLinks (key: string) (links: Map<string, string>) =
        let keptIds = keptFor key |> List.map (fun rel -> rel.Id) |> Set.ofList
        links |> Map.filter (fun _ id -> not (Set.contains id keptIds))
    // A part of notes or comments that was read, with what changed made in it (see patchedStoriesPart): each
    // that changed from its plan (what is as it was read stays its text), or the element the model writes with
    // its ranges opened and closed in it.
    let stories (key: string) (elementName: string) (images: ImageScope) (links: Map<string, string>) (items: (string * bool) list) (write: string -> string) =
        sourceOf key
        |> Option.bind (fun xml ->
            let unchanged = items |> List.filter snd |> List.map fst |> Set.ofList
            patchedStoriesPart xml elementName (List.map fst items) unchanged (fun id ->
                match Map.tryFind (key + ":" + id) notePlans with
                | Some (plan, state) ->
                    useSourceState ctx state
                    let text = generateStoryXmlFromPlan ctx plan images links
                    forgetSourceBody ctx ()
                    text
                | None -> (balancedRanges None [| FromModel (write id) |]).[0]))
    let partRelsNeeded (key: string) (images: ImageScope) (links: Map<string, string>) =
        images.Assets.Length > 0 || not links.IsEmpty || not (keptFor key).IsEmpty

    let files = [
        yield { Path = "[Content_Types].xml"; Content = generateContentTypesPreserved preserved hasNumbering hasSettings hasComments hasFootnotes hasEndnotes headerParts footerParts allImages }
        yield { Path = "_rels/.rels"; Content = generateRootRelsPreserved preserved }
        yield { Path = "word/document.xml"
                Content =
                    withRelationshipsOf "document" bodyModelBlocks (fun () ->
                        match bodyPlan with
                        | Some plan -> generateDocumentXmlFromPlan ctx plan bodyImages hyperlinks pageSetup (optionalPageSetup package.OriginalPageSetup) headerParts footerParts
                        | None -> balancedPart (generateDocumentXmlPackage ctx paragraphs bodyImages pageSetup headerParts footerParts)) }
        yield { Path = "word/_rels/document.xml.rels"; Content = generateDocumentRelsPreserved ctx (keptFor "document") hasNumbering hasSettings newBodyHyperlinks hasComments hasFootnotes hasEndnotes headerParts footerParts bodyImages }
        yield { Path = "word/styles.xml"
                Content =
                    match preserved.StylesXml with
                    | Some original ->
                        let changed = package.ChangedStyles |> Option.map Set.ofArray
                        mergeStyles original changed styles
                    | None -> generateDefaultStylesWithStyles styles }
        if not (keptFor "styles").IsEmpty then
            yield { Path = "word/_rels/styles.xml.rels"; Content = generatePartRelsWithKept (keptFor "styles") Map.empty emptyImageScope }
        match numberingXml with
        | Some text ->
            yield { Path = "word/numbering.xml"; Content = text }
            // (What the lists refer to: the picture of a picture bullet.)
            if not (keptFor "numbering").IsEmpty then
                yield { Path = "word/_rels/numbering.xml.rels"; Content = generatePartRelsWithKept (keptFor "numbering") Map.empty emptyImageScope }
        | None -> ()
        if hasSettings then
            yield { Path = "word/settings.xml"
                    Content =
                        match preserved.SettingsXml with
                        | Some original -> patchSettings original pageSetup
                        | None -> generateSettings pageSetup }
            if not (keptFor "settings").IsEmpty then
                yield { Path = "word/_rels/settings.xml.rels"; Content = generatePartRelsWithKept (keptFor "settings") Map.empty emptyImageScope }
        for header in headerPackages do
            yield { Path = "word/" + header.Info.Target
                    Content =
                        match partSource header.Info with
                        | Some text -> text
                        | None -> writeHeaderFooter "w:hdr" header }
            if partRelsNeeded header.Info.RelId header.Images header.Hyperlinks then
                yield { Path = sprintf "word/_rels/%s.rels" header.Info.Target; Content = generatePartRelsWithKept (keptFor header.Info.RelId) (newLinks header.Info.RelId header.Hyperlinks) header.Images }
        for footer in footerPackages do
            yield { Path = "word/" + footer.Info.Target
                    Content =
                        match partSource footer.Info with
                        | Some text -> text
                        | None -> writeHeaderFooter "w:ftr" footer }
            if partRelsNeeded footer.Info.RelId footer.Images footer.Hyperlinks then
                yield { Path = sprintf "word/_rels/%s.rels" footer.Info.Target; Content = generatePartRelsWithKept (keptFor footer.Info.RelId) (newLinks footer.Info.RelId footer.Hyperlinks) footer.Images }
        if hasComments then
            yield { Path = "word/comments.xml"
                    Content =
                        match commentsSource with
                        | Some text -> text
                        | None ->
                            withRelationshipsOf "comments" commentBlocks (fun () ->
                                let write = commentElement (partBody commentImages commentLinks)
                                stories "comments" "w:comment" commentImages commentLinks (comments |> List.map (fun comment -> comment.Id, comment.Unchanged)) (fun id -> write (comments |> List.find (fun comment -> comment.Id = id)))
                                |> Option.defaultWith (fun () -> balancedPart (generateCommentsPart (partBody commentImages commentLinks) comments))) }
            if partRelsNeeded "comments" commentImages commentLinks then
                yield { Path = "word/_rels/comments.xml.rels"; Content = generatePartRelsWithKept (keptFor "comments") (newLinks "comments" commentLinks) commentImages }
        if hasFootnotes then
            yield { Path = "word/footnotes.xml"
                    Content =
                        match footnotesSource with
                        | Some text -> text
                        | None ->
                            withRelationshipsOf "footnotes" footnoteBlocks (fun () ->
                                let write = noteElement "w:footnote" "w:footnoteRef" (partBody footnoteImages footnoteLinks)
                                stories "footnotes" "w:footnote" footnoteImages footnoteLinks (footnotes |> List.map (fun note -> note.Id, note.Unchanged)) (fun id -> write (footnotes |> List.find (fun note -> note.Id = id)))
                                |> Option.defaultWith (fun () -> balancedPart (generateNotesPart "w:footnotes" "w:footnote" "w:footnoteRef" (partBody footnoteImages footnoteLinks) footnotes))) }
            if partRelsNeeded "footnotes" footnoteImages footnoteLinks then
                yield { Path = "word/_rels/footnotes.xml.rels"; Content = generatePartRelsWithKept (keptFor "footnotes") (newLinks "footnotes" footnoteLinks) footnoteImages }
        if hasEndnotes then
            yield { Path = "word/endnotes.xml"
                    Content =
                        match endnotesSource with
                        | Some text -> text
                        | None ->
                            withRelationshipsOf "endnotes" endnoteBlocks (fun () ->
                                let write = noteElement "w:endnote" "w:endnoteRef" (partBody endnoteImages endnoteLinks)
                                stories "endnotes" "w:endnote" endnoteImages endnoteLinks (endnotes |> List.map (fun note -> note.Id, note.Unchanged)) (fun id -> write (endnotes |> List.find (fun note -> note.Id = id)))
                                |> Option.defaultWith (fun () -> balancedPart (generateNotesPart "w:endnotes" "w:endnote" "w:endnoteRef" (partBody endnoteImages endnoteLinks) endnotes))) }
            if partRelsNeeded "endnotes" endnoteImages endnoteLinks then
                yield { Path = "word/_rels/endnotes.xml.rels"; Content = generatePartRelsWithKept (keptFor "endnotes") (newLinks "endnotes" endnoteLinks) endnoteImages }
    ]

    // Kept parts go back as they were, except where the writer produced a part
    // of the same name.
    let written = files |> List.map (fun file -> file.Path) |> Set.ofList
    let binaryFiles =
        (allImages
         |> Array.choose (fun image -> image.Data |> Option.map (fun data -> sprintf "word/media/%s" image.FileName, data))
         |> Array.toList)
        @ (preserved.Parts |> List.choose (fun part -> if Set.contains part.Path written then None else Some (part.Path, part.Data)))

    // Pictures nothing refers to any more do not stay in the package (a block
    // that was deleted, or that is now written from the model with a picture
    // file of its own, leaves the relationship and the file it had): a picture
    // relationship of a part whose XML does not use its id goes, and then every
    // file under word/media that no relationship part names. The same of the
    // relationship of a hyperlink (a paragraph with a link that was deleted, or
    // whose link was taken away): it goes when the part's XML does not use its
    // id any more - one the part did not use when it was read stays, the
    // file's own.
    let files =
        let contentOf = files |> List.map (fun file -> file.Path, file.Content) |> Map.ofList
        let sourceOfPath =
            [ yield "word/document.xml", "document"
              yield "word/footnotes.xml", "footnotes"
              yield "word/endnotes.xml", "endnotes"
              yield "word/comments.xml", "comments"
              for part in Array.append headerParts footerParts do yield "word/" + part.Target, part.RelId ]
            |> List.choose (fun (path, key) ->
                Map.tryFind key preserved.SourceXml
                |> Option.map (fun xml ->
                    let ids (map: Map<string, PreservedRelationshipInfo list>) = map |> Map.tryFind key |> Option.defaultValue [] |> List.map (fun rel -> rel.Id)
                    path, (xml, Set.ofList (ids preserved.SourceRelationships @ ids preserved.PartRelationships))))
            |> Map.ofList
        files
        |> List.map (fun file ->
            match PartPath.ofRelationships file.Path with
            | Some partPath when PartPath.directory file.Path = "word/_rels" ->
                match Map.tryFind partPath contentOf with
                | None -> file
                | Some part ->
                    let usedNow = OpcXml.relationshipIdsOfPart part
                    let source = Map.tryFind partPath sourceOfPath |> Option.map (fun (source, ids) -> OpcXml.relationshipIdsOfPart source, ids)
                    let originalIds =
                        Map.tryFind file.Path preserved.OriginalText |> Option.map (OpcXml.relationshipsOrEmpty >> List.map (fun rel -> rel.Id) >> Set.ofList)
                    let unused (rel: OpcXml.Relationship) =
                        // (A link's relationship the part had and did not use when it was read: the file's own.)
                        let inherited =
                            match source with
                            | Some (usedThen, ids) -> Set.contains rel.Id ids && not (Set.contains rel.Id usedThen)
                            | None -> false
                        // (And a picture's: one the file's relationships named and the part it was read as did not
                        // use - a picture a tool left behind - is the file's own too. Only what a change made
                        // unused goes.)
                        let leftBehind =
                            match source, originalIds with
                            | Some (usedThen, _), Some ids -> not (Set.contains rel.Id usedThen) && Set.contains rel.Id ids
                            | _ -> false
                        let kind = OpcXml.relationshipKind rel.Type
                        let picture = kind = "image" && not rel.External && not leftBehind
                        let link = kind = "hyperlink" && not inherited
                        (picture || link) && not (Set.contains rel.Id usedNow)
                    { file with Content = OpcXml.withoutRelationships unused file.Content }
            | _ -> file)
    let binaryFiles =
        // (Every part a relationship of the package names, as its path.)
        let named =
            (files |> List.map (fun file -> file.Path, file.Content))
            @ (binaryFiles |> List.map (fun (path, data) -> path, System.Text.Encoding.UTF8.GetString data))
            |> List.filter (fun (path, _) -> PartPath.ofRelationships path |> Option.isSome)
            |> List.collect (fun (path, text) ->
                OpcXml.relationshipsOrEmpty text |> List.filter (fun rel -> not rel.External) |> List.map (OpcXml.targetPath (OpcXml.ownerDirectory path)))
            |> Set.ofList
        binaryFiles
        |> List.filter (fun (path, _) -> PartPath.directory path <> "word/media" || Set.contains path named)

    // A document that was read without a styles part gets none while nothing needs one: no
    // style was made or changed, and no part names one. Nor the relationship and the content
    // type of one; and the body's relationships part, which the file did not have, is not
    // written when nothing is left in it. (A shape moved in a document of one part was saved
    // with a styles part of the writer's defaults and a relationships part to name it.)
    let files =
        let stylesChanged = package.ChangedStyles |> Option.exists (Array.isEmpty >> not)
        let wasRead = Map.containsKey "[Content_Types].xml" preserved.OriginalText
        let namesStyle (file: PackageFile) =
            PartPath.directory file.Path = "word" && PartPath.extension file.Path = "xml" && file.Path <> "word/styles.xml"
            && (match XmlPatch.readPart file.Content with
                | Some part -> [ "pStyle"; "rStyle"; "tblStyle" ] |> List.exists (fun name -> not (XmlPatch.descendants name part.Node).IsEmpty)
                | None -> false)
        if regenerateAll || not wasRead || preserved.StylesXml.IsSome || stylesChanged || not (keptFor "styles").IsEmpty || List.exists namesStyle files then files
        else
            files
            |> List.filter (fun file -> file.Path <> "word/styles.xml")
            |> List.map (fun file ->
                if file.Path = "word/_rels/document.xml.rels" then
                    { file with Content = file.Content |> OpcXml.withoutRelationships (fun rel -> OpcXml.relationshipKind rel.Type = "styles") }
                elif file.Path = "[Content_Types].xml" then
                    let styles = OpcXml.contentTypeEntries file.Content |> Option.defaultValue [] |> List.filter (function OpcXml.Override (name, _) -> name = "/word/styles.xml" | _ -> false)
                    { file with Content = file.Content |> OpcXml.editContentTypes styles [] }
                else file)
            |> List.filter (fun file ->
                not (file.Path = "word/_rels/document.xml.rels"
                     && not (Map.containsKey file.Path preserved.OriginalText)
                     && OpcXml.relationshipsOrEmpty file.Content |> List.isEmpty))

    // The relationship parts and [Content_Types].xml that say what they said when they were read go back as that text.
    let files =
        if regenerateAll || preserved.OriginalText.IsEmpty then files
        else
            let writtenPaths = (files |> List.map (fun file -> file.Path)) @ (binaryFiles |> List.map fst)
            files
            |> List.map (fun file ->
                match Map.tryFind file.Path preserved.OriginalText with
                | None -> file
                | Some original ->
                    if file.Path = "[Content_Types].xml" then
                        if contentTypesStillRight original file.Content writtenPaths then { file with Content = original } else file
                    elif PartPath.ofRelationships file.Path |> Option.isSome then
                        let directory = OpcXml.ownerDirectory file.Path
                        match relationshipEntries directory original, relationshipEntries directory file.Content with
                        | Some before, Some after when before = after -> { file with Content = original }
                        | _ -> file
                    else file)

    ctx.Model.ReservedRelationshipIds <- Set.empty
    ctx.Model.ReservedPartPaths <- Set.empty
    forgetSourceBody ctx ()
    files, binaryFiles

let private unchangedBytes (package: EditorPackage) =
    if package.Unchanged && not package.RegenerateAll then
        package.Preserved |> Option.bind (fun preserved -> preserved.OpenedBytes)
    else None

#if FABLE_COMPILER
[<Emit("Promise.resolve(new Blob([$0.slice()], { type: 'application/zip' }))")>]
let private blobOfCopy (_bytes: byte[]) : Fable.Core.JS.Promise<obj> = jsNative

let writeTypedPackage (package: EditorPackage) : Fable.Core.JS.Promise<obj> =
    match unchangedBytes package with
    | Some bytes -> blobOfCopy bytes
    | None ->
        let ctx = createPackageContext ()
        let files, binaryFiles = preparePackage ctx package
        let report =
            { Reasons = List.ofSeq ctx.ReportReasons; Derived = ctx.Model.ReportDerived
              Generated = ctx.Model.ReportGenerated; Spliced = ctx.Model.ReportSpliced; Patched = ctx.Model.ReportPatched
              Inline = ctx.ReportInline; ModelBlocks = ctx.ReportModelBlocks
              ModelParagraphs = ctx.ReportModelParagraphs; SourceBlocks = ctx.ReportSourceBlocks }
        ZipInterop.promise {
            let! blob = createPackageWithBinary files binaryFiles
            lastCompletedReport <- report
            return blob
        }

let writeFromEditorPackage (package: obj) : Fable.Core.JS.Promise<obj> =
    package |> WordEditorPackage.packageOfEditorObject |> writeTypedPackage
#else
let writeTypedPackageBytes (package: EditorPackage) : byte[] =
    match unchangedBytes package with
    | Some bytes -> Array.copy bytes
    | None ->
        let ctx = createPackageContext ()
        let files, binaryFiles = preparePackage ctx package
        createPackageBytesWithBinary files binaryFiles
#endif
