module RedBat.Word.Commands

open System.Collections.Generic
open RedBat.Word.Document
open RedBat.Word.TextBuffer

/// Command interface for undo/redo
type ICommand =
    abstract Execute: Document -> Document
    abstract Undo: Document -> Document
    abstract Description: string
    abstract RetainedSourceTextCharCount: int

/// Command history for undo/redo
type CommandHistory = {
    Undos: ICommand list
    Redos: ICommand list
    MaxHistory: int
}

/// Create empty command history
let empty = {
    Undos = []
    Redos = []
    MaxHistory = 100
}

let private commandHistoryMaxTotalRetainedSourceTextChars = 20000000
let private commandHistoryMaxRetainedSourceTraversalItems = 200000

let private addCapped limit value total =
    if total > limit then
        total
    else
        let value = max 0 value
        let remaining = limit - total
        if value > remaining then limit + 1
        else total + value

let private commandHistorySourceTextOverLimit = commandHistoryMaxTotalRetainedSourceTextChars + 1

let private failClosedForTraversalCount visitedItems total =
    if visitedItems > commandHistoryMaxRetainedSourceTraversalItems then
        commandHistorySourceTextOverLimit
    else
        total

let private optionStringCharCount (value: string option) =
    match value with
    | Some value -> value.Length
    | None -> 0

let private sectionColumnsStringCharCount (columns: SectionColumns option) =
    match columns with
    | Some columns ->
        let mutable total = 0
        let mutable seen = 0
        let mutable widths = columns.Widths
        while total <= commandHistoryMaxTotalRetainedSourceTextChars
              && seen <= commandHistoryMaxRetainedSourceTraversalItems
              && not widths.IsEmpty do
            total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars 1 total
            seen <- seen + 1
            widths <- widths.Tail
        failClosedForTraversalCount seen total
    | None -> 0

let private textStyleStringCharCount (style: RedBat.Word.TextBuffer.TextStyle) =
    let mutable total = 0
    let add (value: string option) =
        total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount value) total
    add style.Style
    add style.UnderlineStyle
    add style.UnderlineColor
    add style.EmphasisMark
    add style.FontFamily
    add style.Color
    add style.Highlight
    add style.Hyperlink
    add style.CommentId
    add style.CommentText
    add style.CommentAuthor
    add style.CommentDate
    add style.NoteId
    add style.NoteType
    add style.NoteText
    add style.BookmarkStartId
    add style.BookmarkName
    add style.BookmarkEndId
    add style.FieldCode
    add style.FieldResult
    add style.EquationXml
    add style.RevisionId
    add style.RevisionType
    add style.RevisionAuthor
    add style.RevisionDate
    add style.VertAlign
    total

let private imageStringCharCount (image: RedBat.Word.TextBuffer.ImageRun) =
    let mutable total = image.Src.Length
    let add (value: string option) =
        total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount value) total
    add image.Alt
    add image.Name
    add image.Title
    add image.Wrap
    add image.PositionHRelativeFrom
    add image.PositionVRelativeFrom
    add image.RelativeHeight
    add image.CropLeft
    add image.CropTop
    add image.CropRight
    add image.CropBottom
    total

let private runStringCharCount (run: RedBat.Word.TextBuffer.Run) =
    match run with
    | RedBat.Word.TextBuffer.Text textRun -> textStyleStringCharCount textRun.Style
    | RedBat.Word.TextBuffer.Image image -> imageStringCharCount image

let rec private paragraphSourceTextCharCount (paragraph: ParagraphBlock) =
    let limit = commandHistoryMaxTotalRetainedSourceTextChars
    let mutable total = 0
    total <- addCapped limit paragraph.Id.Length total
    total <- addCapped limit (RedBat.Word.TextBuffer.length paragraph.Buffer) total
    let mutable seen = 0
    let mutable runs = paragraph.Buffer.Runs
    while total <= limit
          && seen <= commandHistoryMaxRetainedSourceTraversalItems
          && not runs.IsEmpty do
        total <- addCapped limit (runStringCharCount runs.Head) total
        seen <- seen + 1
        runs <- runs.Tail
    if total <= limit then
        total <- addCapped limit (paragraphMetadataStringCharCount paragraph.Metadata) total
    failClosedForTraversalCount seen total

and private paragraphBorderStringCharCount (border: ParagraphBorder option) =
    match border with
    | Some border ->
        let mutable total = 0
        total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount border.Style) total
        total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount border.Color) total
        total
    | None -> 0

and private paragraphBordersStringCharCount (borders: ParagraphBorders option) =
    match borders with
    | Some borders ->
        let mutable total = 0
        total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (paragraphBorderStringCharCount borders.Top) total
        total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (paragraphBorderStringCharCount borders.Right) total
        total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (paragraphBorderStringCharCount borders.Bottom) total
        total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (paragraphBorderStringCharCount borders.Left) total
        total
    | None -> 0

and private tabStopStringCharCount (tabStop: TabStop) =
    addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount tabStop.Leader) tabStop.Alignment.Length

and private paragraphMetadataStringCharCount (metadata: BlockMetadata) =
    let mutable total = 0
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount metadata.Style) total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount metadata.ContentControlId) total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount metadata.ContentControlTag) total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount metadata.ContentControlAlias) total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount metadata.ListId) total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount metadata.ListNumberFormat) total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount metadata.ListLevelText) total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount metadata.RightTabLeader) total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount metadata.LineSpacingRule) total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount metadata.BackgroundColor) total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (paragraphBordersStringCharCount metadata.ParagraphBorders) total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount metadata.SectionBreak) total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (sectionColumnsStringCharCount metadata.SectionColumns) total
    let mutable seen = 0
    let mutable tabStops = metadata.TabStops
    while total <= commandHistoryMaxTotalRetainedSourceTextChars
          && seen <= commandHistoryMaxRetainedSourceTraversalItems
          && not tabStops.IsEmpty do
        total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (tabStopStringCharCount tabStops.Head) total
        seen <- seen + 1
        tabStops <- tabStops.Tail
    let mutable headerReferences = metadata.HeaderReferences
    while total <= commandHistoryMaxTotalRetainedSourceTextChars
          && seen <= commandHistoryMaxRetainedSourceTraversalItems
          && not headerReferences.IsEmpty do
        let kind, relationshipId = headerReferences.Head
        total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars kind.Length total
        total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars relationshipId.Length total
        seen <- seen + 1
        headerReferences <- headerReferences.Tail
    let mutable footerReferences = metadata.FooterReferences
    while total <= commandHistoryMaxTotalRetainedSourceTextChars
          && seen <= commandHistoryMaxRetainedSourceTraversalItems
          && not footerReferences.IsEmpty do
        let kind, relationshipId = footerReferences.Head
        total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars kind.Length total
        total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars relationshipId.Length total
        seen <- seen + 1
        footerReferences <- footerReferences.Tail
    failClosedForTraversalCount seen total

let private tableBorderStringCharCount (border: TableBorder option) =
    match border with
    | Some border ->
        let mutable total = 0
        total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount border.Style) total
        total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount border.Color) total
        total
    | None -> 0

let private tableBordersStringCharCount (borders: TableBorders option) =
    match borders with
    | Some borders ->
        let mutable total = 0
        total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (tableBorderStringCharCount borders.Top) total
        total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (tableBorderStringCharCount borders.Right) total
        total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (tableBorderStringCharCount borders.Bottom) total
        total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (tableBorderStringCharCount borders.Left) total
        total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (tableBorderStringCharCount borders.InsideH) total
        total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (tableBorderStringCharCount borders.InsideV) total
        total
    | None -> 0

let private shadingStringCharCount (shading: Shading option) =
    match shading with
    | Some shading ->
        let mutable total = 0
        total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount shading.Pattern) total
        total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount shading.ForegroundColor) total
        total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount shading.BackgroundColor) total
        total
    | None -> 0

let private tableStringCharCount (table: TableBlock) =
    let mutable total = table.Id.Length
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount table.Alignment) total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount table.BackgroundColor) total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (shadingStringCharCount table.Shading) total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (tableBordersStringCharCount table.Borders) total
    total

let private tableRowStringCharCount (row: TableRow) =
    let mutable total = 0
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount row.HeightRule) total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (tableBordersStringCharCount row.BordersOverride) total
    total

let private tableCellStringCharCount (cell: TableCell) =
    let mutable total = 0
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount cell.BackgroundColor) total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (shadingStringCharCount cell.Shading) total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount cell.BorderColor) total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (tableBordersStringCharCount cell.Borders) total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount cell.VerticalAlignment) total
    total

let private pathSegmentsStringCharCount (segments: PathSeg list) =
    let mutable total = 0
    let mutable seen = 0
    let mutable remaining = segments
    while total <= commandHistoryMaxTotalRetainedSourceTextChars
          && seen <= commandHistoryMaxRetainedSourceTraversalItems
          && not remaining.IsEmpty do
        total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars 1 total
        seen <- seen + 1
        remaining <- remaining.Tail
    failClosedForTraversalCount seen total

let private shapeGeometryStringCharCount = function
    | GeoRect -> 0
    | GeoPreset preset -> preset.Length
    | GeoCustom (segments, _, _) -> pathSegmentsStringCharCount segments

let private gradientStringCharCount (gradient: GradientSpec option) =
    match gradient with
    | Some gradient ->
        let mutable total = 0
        let mutable seen = 0
        let mutable stops = gradient.Stops
        while total <= commandHistoryMaxTotalRetainedSourceTextChars
              && seen <= commandHistoryMaxRetainedSourceTraversalItems
              && not stops.IsEmpty do
            total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars stops.Head.Color.Length total
            seen <- seen + 1
            stops <- stops.Tail
        failClosedForTraversalCount seen total
    | None -> 0

let private textBoxStringCharCount (textBox: TextBoxBlock) =
    let mutable total = textBox.Id.Length
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount textBox.PositionHRelativeFrom) total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount textBox.PositionVRelativeFrom) total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount textBox.FillColor) total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount textBox.BorderColor) total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount textBox.RelativeHeight) total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (shapeGeometryStringCharCount textBox.Geometry) total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (gradientStringCharCount textBox.Gradient) total
    total

let private blocksSourceTextCharCount (blocks: Block list) =
    let mutable total = 0
    let mutable visitedItems = 0
    let canContinue () =
        total <= commandHistoryMaxTotalRetainedSourceTextChars
        && visitedItems <= commandHistoryMaxRetainedSourceTraversalItems
    let visitItem () =
        visitedItems <- visitedItems + 1
    let addParagraph paragraph =
        if canContinue () then
            visitItem ()
            total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (paragraphSourceTextCharCount paragraph) total
    let addTable table =
        let tables = Stack<TableBlock>()
        tables.Push table
        while tables.Count > 0 && canContinue () do
            let table = tables.Pop()
            visitItem ()
            total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (tableStringCharCount table) total
            let mutable rows = table.Rows
            while canContinue () && not rows.IsEmpty do
                let row = rows.Head
                visitItem ()
                total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (tableRowStringCharCount row) total
                let mutable cells = row.Cells
                while canContinue () && not cells.IsEmpty do
                    let cell = cells.Head
                    visitItem ()
                    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (tableCellStringCharCount cell) total
                    let mutable paragraphs = cell.Content
                    while canContinue () && not paragraphs.IsEmpty do
                        addParagraph paragraphs.Head
                        paragraphs <- paragraphs.Tail
                    let mutable nestedTables = cell.NestedTables
                    while canContinue () && not nestedTables.IsEmpty do
                        tables.Push nestedTables.Head
                        nestedTables <- nestedTables.Tail
                    cells <- cells.Tail
                rows <- rows.Tail
    let mutable remaining = blocks
    while canContinue () && not remaining.IsEmpty do
        match remaining.Head with
        | Paragraph paragraph
        | Heading (_, paragraph) -> addParagraph paragraph
        | Table table -> addTable table
        | TextBox textBox ->
            visitItem ()
            total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (textBoxStringCharCount textBox) total
            let mutable paragraphs = textBox.Content
            while canContinue () && not paragraphs.IsEmpty do
                addParagraph paragraphs.Head
                paragraphs <- paragraphs.Tail
        remaining <- remaining.Tail
    failClosedForTraversalCount visitedItems total

let private paragraphListSourceTextCharCount (paragraphs: ParagraphBlock list) =
    let mutable total = 0
    let mutable seen = 0
    let mutable remaining = paragraphs
    while total <= commandHistoryMaxTotalRetainedSourceTextChars
          && seen <= commandHistoryMaxRetainedSourceTraversalItems
          && not remaining.IsEmpty do
        total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (paragraphSourceTextCharCount remaining.Head) total
        seen <- seen + 1
        remaining <- remaining.Tail
    failClosedForTraversalCount seen total

let private headerFooterPartsMetadataStringCharCount (parts: HeaderFooterPart list) =
    let limit = commandHistoryMaxTotalRetainedSourceTextChars
    let mutable total = 0
    let mutable seen = 0
    let mutable remaining = parts
    while total <= limit
          && seen <= commandHistoryMaxRetainedSourceTraversalItems
          && not remaining.IsEmpty do
        let part = remaining.Head
        total <- addCapped limit part.Kind.Length total
        if total <= limit then
            total <- addCapped limit (optionStringCharCount part.RelationshipId) total
        seen <- seen + 1
        remaining <- remaining.Tail
    failClosedForTraversalCount seen total

let private headerFooterPartsBlocksSourceTextCharCount (parts: HeaderFooterPart list) =
    let limit = commandHistoryMaxTotalRetainedSourceTextChars
    let mutable total = 0
    let mutable seen = 0
    let mutable remaining = parts
    while total <= limit
          && seen <= commandHistoryMaxRetainedSourceTraversalItems
          && not remaining.IsEmpty do
        total <- addCapped limit (blocksSourceTextCharCount remaining.Head.Blocks) total
        seen <- seen + 1
        remaining <- remaining.Tail
    failClosedForTraversalCount seen total

let private namedStyleStringCharCount (style: NamedStyle) =
    let mutable total = style.Id.Length
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount style.Name) total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars style.StyleType.Length total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount style.BasedOn) total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount style.Next) total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount style.Linked) total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount style.Aliases) total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount style.LineSpacingRule) total
    total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (optionStringCharCount style.BackgroundColor) total
    match style.Text with
    | Some textStyle ->
        total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (textStyleStringCharCount textStyle) total
    | None -> ()
    total

let private stylesSourceTextCharCount (styles: NamedStyle list) =
    let mutable total = 0
    let mutable seen = 0
    let mutable remaining = styles
    while total <= commandHistoryMaxTotalRetainedSourceTextChars
          && seen <= commandHistoryMaxRetainedSourceTraversalItems
          && not remaining.IsEmpty do
        total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (namedStyleStringCharCount remaining.Head) total
        seen <- seen + 1
        remaining <- remaining.Tail
    failClosedForTraversalCount seen total

let private selectionSourceTextCharCount (selection: Selection option) =
    match selection with
    | Some selection ->
        addCapped commandHistoryMaxTotalRetainedSourceTextChars selection.Focus.BlockId.Length selection.Anchor.BlockId.Length
    | None -> 0

let private pageSetupStringCharCount (setup: PageSetup) =
    addCapped
        commandHistoryMaxTotalRetainedSourceTextChars
        (sectionColumnsStringCharCount setup.Columns)
        (optionStringCharCount setup.SectionStart)

let private documentSourceTextCharCount (doc: Document) =
    let limit = commandHistoryMaxTotalRetainedSourceTextChars
    let mutable total = 0
    if total <= limit then
        total <- addCapped limit (selectionSourceTextCharCount doc.Selection) total
    if total <= limit then
        total <- addCapped limit (pageSetupStringCharCount doc.PageSetup) total
    if total <= limit then
        total <- addCapped limit (stylesSourceTextCharCount doc.Styles) total
    if total <= limit then
        total <- addCapped limit (headerFooterPartsMetadataStringCharCount doc.HeaderParts) total
    if total <= limit then
        total <- addCapped limit (headerFooterPartsMetadataStringCharCount doc.FooterParts) total
    if total <= limit then
        total <- addCapped limit (blocksSourceTextCharCount doc.Blocks) total
    if total <= limit then
        total <- addCapped limit (paragraphListSourceTextCharCount doc.Headers) total
    if total <= limit then
        total <- addCapped limit (paragraphListSourceTextCharCount doc.Footers) total
    if total <= limit then
        total <- addCapped limit (headerFooterPartsBlocksSourceTextCharCount doc.HeaderParts) total
    if total <= limit then
        total <- addCapped limit (headerFooterPartsBlocksSourceTextCharCount doc.FooterParts) total
    total

let private retainedSourceTextCharCountForDocs (docs: Document list) =
    let mutable total = 0
    let mutable remaining = docs
    while total <= commandHistoryMaxTotalRetainedSourceTextChars && not remaining.IsEmpty do
        total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (documentSourceTextCharCount remaining.Head) total
        remaining <- remaining.Tail
    total

let private retainedSourceTextCharCountForDocsWithInitial initialChars (docs: Document list) =
    let mutable total = addCapped commandHistoryMaxTotalRetainedSourceTextChars initialChars 0
    let mutable remaining = docs
    while total <= commandHistoryMaxTotalRetainedSourceTextChars && not remaining.IsEmpty do
        total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars (documentSourceTextCharCount remaining.Head) total
        remaining <- remaining.Tail
    total

let private trim maxHistory (commands: ICommand list) =
    let rec loop (remaining: ICommand list) (kept: ICommand list) count totalRetainedChars =
        match remaining with
        | [] -> List.rev kept
        | cmd :: rest ->
            let retainedChars = max 1 cmd.RetainedSourceTextCharCount
            let nextRetainedChars = totalRetainedChars + retainedChars
            if count < maxHistory
               && retainedChars <= commandHistoryMaxTotalRetainedSourceTextChars
               && nextRetainedChars <= commandHistoryMaxTotalRetainedSourceTextChars then
                loop rest (cmd :: kept) (count + 1) nextRetainedChars
            else
                loop rest kept count totalRetainedChars
    loop commands [] 0 0

/// Execute a command and add to history
let execute (cmd: ICommand) (doc: Document) (history: CommandHistory) =
    let newDoc = cmd.Execute doc
    if obj.ReferenceEquals(newDoc, doc) then
        doc, history
    else
        let newHistory = {
            Undos = (cmd :: history.Undos) |> trim history.MaxHistory
            Redos = []  // Clear redo stack when new command is executed
            MaxHistory = history.MaxHistory
        }
        (newDoc, newHistory)

/// Undo last command
let undo (doc: Document) (history: CommandHistory) =
    match history.Undos with
    | [] -> (doc, history)  // Nothing to undo
    | cmd :: rest ->
        let newDoc = cmd.Undo doc
        let newHistory = {
            Undos = rest
            Redos = (cmd :: history.Redos) |> trim history.MaxHistory
            MaxHistory = history.MaxHistory
        }
        (newDoc, newHistory)

/// Redo last undone command
let redo (doc: Document) (history: CommandHistory) =
    match history.Redos with
    | [] -> (doc, history)  // Nothing to redo
    | cmd :: rest ->
        let newDoc = cmd.Execute doc
        let newHistory = {
            Undos = (cmd :: history.Undos) |> trim history.MaxHistory
            Redos = rest
            MaxHistory = history.MaxHistory
        }
        (newDoc, newHistory)

/// Insert text command
type InsertTextCommand(text: string, position: Position, beforeDoc: Document) =
    let retainedSourceTextCharCount =
        retainedSourceTextCharCountForDocsWithInitial text.Length [ beforeDoc ]

    interface ICommand with
        member _.Execute doc =
            // Set cursor to position and insert text
            let doc' = setCursor position.BlockId position.Offset doc
            insertText text doc'

        member _.Undo doc =
            beforeDoc

        member _.Description = sprintf "Insert '%s'" text
        member _.RetainedSourceTextCharCount = retainedSourceTextCharCount

/// Delete text command (backspace/delete)
type DeleteTextCommand(beforeDoc: Document, afterDoc: Document, description: string) =
    let retainedSourceTextCharCount =
        retainedSourceTextCharCountForDocs [ beforeDoc; afterDoc ]

    interface ICommand with
        member _.Execute doc = afterDoc
        member _.Undo doc = beforeDoc
        member _.Description = description
        member _.RetainedSourceTextCharCount = retainedSourceTextCharCount

/// Split block command (Enter key)
type SplitBlockCommand(beforeDoc: Document, afterDoc: Document) =
    let retainedSourceTextCharCount =
        retainedSourceTextCharCountForDocs [ beforeDoc; afterDoc ]

    interface ICommand with
        member _.Execute doc = afterDoc
        member _.Undo doc = beforeDoc
        member _.Description = "Split paragraph"
        member _.RetainedSourceTextCharCount = retainedSourceTextCharCount

/// Format text command (bold/italic/underline)
type FormatTextCommand(beforeDoc: Document, afterDoc: Document, formatType: string) =
    let retainedSourceTextCharCount =
        retainedSourceTextCharCountForDocs [ beforeDoc; afterDoc ]

    interface ICommand with
        member _.Execute doc = afterDoc
        member _.Undo doc = beforeDoc
        member _.Description = sprintf "Format: %s" formatType
        member _.RetainedSourceTextCharCount = retainedSourceTextCharCount

/// Composite command for grouping multiple commands
type CompositeCommand(commands: ICommand list, description: string) =
    let retainedSourceTextCharCount =
        let mutable total = 0
        let mutable remaining = commands
        let mutable seen = 0
        while total <= commandHistoryMaxTotalRetainedSourceTextChars
              && seen <= commandHistoryMaxRetainedSourceTraversalItems
              && not remaining.IsEmpty do
            total <- addCapped commandHistoryMaxTotalRetainedSourceTextChars remaining.Head.RetainedSourceTextCharCount total
            seen <- seen + 1
            remaining <- remaining.Tail
        failClosedForTraversalCount seen total

    interface ICommand with
        member _.Execute doc =
            commands |> List.fold (fun d cmd -> cmd.Execute d) doc

        member _.Undo doc =
            commands |> List.rev |> List.fold (fun d cmd -> cmd.Undo d) doc

        member _.Description = description
        member _.RetainedSourceTextCharCount = retainedSourceTextCharCount
