module RedBat.Word.Document

open RedBat.Word.TextBuffer

/// Block ID for tracking blocks
type BlockId = string

/// Counter for generating unique block IDs
let mutable private blockIdCounter = 0

/// Generate next block ID
let private nextBlockId () =
    blockIdCounter <- blockIdCounter + 1
    sprintf "block-%d" blockIdCounter

/// Text style for runs
type TextStyle = {
    Bold: bool
    Italic: bool
    Underline: bool
    FontSize: float option
    Color: string option
}

/// Default text style
let defaultStyle = {
    Bold = false
    Italic = false
    Underline = false
    FontSize = None
    Color = None
}

/// Text run - contiguous text with same formatting
type TextRun = {
    Text: string
    Style: TextStyle
}

/// List style
type ListStyle =
    | NoList
    | Bullet
    | Numbered

/// Text alignment
type Alignment =
    | Left
    | Center
    | Right
    | Justify

type SectionPageSetup = {
    PageWidth: float
    PageHeight: float
    MarginLeft: float
    MarginRight: float
    MarginTop: float
    MarginBottom: float
    HeaderDistance: float
    FooterDistance: float
    DifferentFirstPage: bool
    EvenAndOddHeaders: bool
}

type ParagraphBorder = {
    Style: string option
    Color: string option
    Width: float option
    Space: float option
}

type ParagraphBorders = {
    Top: ParagraphBorder option
    Right: ParagraphBorder option
    Bottom: ParagraphBorder option
    Left: ParagraphBorder option
}

/// Block metadata
type BlockMetadata = {
    Style: string option
    ContentControlId: string option
    ContentControlTag: string option
    ContentControlAlias: string option
    ListStyle: ListStyle
    ListLevel: int  // 0 = not in list, 1-9 = nesting level
    ListId: string option
    ListNumberFormat: string option
    ListLevelText: string option
    ListStart: int option
    Alignment: Alignment
    IndentLeft: float option
    IndentRight: float option
    FirstLineIndent: float option
    RightTabPosition: float option
    SpacingBefore: float option
    SpacingAfter: float option
    ParagraphMarkFontSize: float option
    ContextualSpacing: bool option
    LineSpacing: float option
    LineSpacingRule: string option
    WidowControl: bool option
    BackgroundColor: string option
    ParagraphBorders: ParagraphBorders option
    PageBreakBefore: bool
    PageBreakAfter: bool
    SectionBreak: string option
    SectionPageNumberStart: int option
    RenderedPageBreakBefore: int option
    SectionPageSetup: SectionPageSetup option
    HeaderReferences: (string * string) list
    FooterReferences: (string * string) list
}

/// Default block metadata
let defaultMetadata = {
    Style = None
    ContentControlId = None
    ContentControlTag = None
    ContentControlAlias = None
    ListStyle = NoList
    ListLevel = 0
    ListId = None
    ListNumberFormat = None
    ListLevelText = None
    ListStart = None
    Alignment = Left
    IndentLeft = None
    IndentRight = None
    FirstLineIndent = None
    RightTabPosition = None
    SpacingBefore = None
    SpacingAfter = None
    ParagraphMarkFontSize = None
    ContextualSpacing = None
    LineSpacing = None
    LineSpacingRule = None
    WidowControl = None
    BackgroundColor = None
    ParagraphBorders = None
    PageBreakBefore = false
    PageBreakAfter = false
    SectionBreak = None
    SectionPageNumberStart = None
    RenderedPageBreakBefore = None
    SectionPageSetup = None
    HeaderReferences = []
    FooterReferences = []
}

/// Block types
type Block =
    | Paragraph of ParagraphBlock
    | Heading of level: int * ParagraphBlock
    | Table of TableBlock
    | TextBox of TextBoxBlock

and ParagraphBlock = {
    Id: BlockId
    Buffer: TextBuffer
    Metadata: BlockMetadata
}

and TableBlock = {
    Id: BlockId
    Rows: TableRow list
    PreferredWidth: TableWidth option
    LayoutMode: TableLayoutMode option
    GridColumns: TableWidth list
    Borders: TableBorders option
    BackgroundColor: string option
    CellSpacing: float option
    PaddingTop: float option
    PaddingRight: float option
    PaddingBottom: float option
    PaddingLeft: float option
    Floating: bool
}

and TextBoxBlock = {
    Id: BlockId
    Content: ParagraphBlock list
    X: float option
    Y: float option
    PositionHRelativeFrom: string option
    PositionVRelativeFrom: string option
    Width: float
    Height: float
    FillColor: string option
    BorderColor: string option
    BorderWidth: float option
    RelativeHeight: string option
}

and TableRow = {
    Cells: TableCell list
    IsHeader: bool
    Height: float option
    HeightRule: string option
    CanSplit: bool option
    CellSpacing: float option
    PaddingTop: float option
    PaddingRight: float option
    PaddingBottom: float option
    PaddingLeft: float option
}

and TableCell = {
    Content: ParagraphBlock list  // Cells can contain multiple paragraphs
    ColSpan: int
    BackgroundColor: string option
    Width: float option
    PreferredWidth: TableWidth option
    BorderColor: string option
    BorderWidth: float option
    Borders: TableBorders option
    VerticalAlignment: string option
    VerticalMerge: TableVerticalMerge option
    PaddingTop: float option
    PaddingRight: float option
    PaddingBottom: float option
    PaddingLeft: float option
}

and TableWidthKind =
    | AutoWidth
    | DxaWidth
    | PercentWidth
    | PointWidth

and TableWidth = {
    Kind: TableWidthKind
    Value: float option
}

and TableLayoutMode =
    | FixedTableLayout
    | AutoFitTableLayout

and TableBorder = {
    Style: string option
    Color: string option
    Width: float option
}

and TableBorders = {
    Top: TableBorder option
    Right: TableBorder option
    Bottom: TableBorder option
    Left: TableBorder option
    InsideH: TableBorder option
    InsideV: TableBorder option
}

and TableVerticalMerge =
    | RestartMerge
    | ContinueMerge

/// Cursor position
type Position = {
    BlockId: BlockId
    Offset: int  // Character offset in block
}

/// Selection (anchor = start, focus = end/cursor)
type Selection = {
    Anchor: Position
    Focus: Position
}
with
    /// Check if selection is collapsed (cursor only)
    member this.IsCollapsed =
        this.Anchor.BlockId = this.Focus.BlockId &&
        this.Anchor.Offset = this.Focus.Offset

    /// Get normalized range (start, end) accounting for direction - single block only
    member this.Range =
        if this.Anchor.BlockId = this.Focus.BlockId then
            if this.Anchor.Offset <= this.Focus.Offset then
                (this.Anchor, this.Focus)
            else
                (this.Focus, this.Anchor)
        else
            // For now, assume single block
            (this.Anchor, this.Focus)

type BookmarkInfo = {
    Id: string
    Name: string
    BlockId: BlockId
    StartOffset: int
    EndOffset: int
}

type ContentControlInfo = {
    Id: string
    Tag: string
    Alias: string
    BlockId: BlockId
}

/// Document
type HeaderFooterPart = {
    Kind: string
    RelationshipId: string option
    Blocks: Block list
}

type PageSetup = {
    PageWidth: float
    PageHeight: float
    MarginLeft: float
    MarginRight: float
    MarginTop: float
    MarginBottom: float
    HeaderDistance: float
    FooterDistance: float
    DifferentFirstPage: bool
    EvenAndOddHeaders: bool
}

type NamedStyle = {
    Id: string
    Name: string option
    StyleType: string
    IsDefault: bool
    BasedOn: string option
    Next: string option
    Linked: string option
    IsQuickStyle: bool
    Aliases: string option
    UiPriority: int option
    SemiHidden: bool
    UnhideWhenUsed: bool
    Text: TextBuffer.TextStyle option
    SpacingBefore: float option
    SpacingAfter: float option
    ContextualSpacing: bool option
    LineSpacing: float option
    LineSpacingRule: string option
    Alignment: Alignment option
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

let defaultPageSetup = {
    PageWidth = 816.0
    PageHeight = 1056.0
    MarginLeft = 96.0
    MarginRight = 96.0
    MarginTop = 96.0
    MarginBottom = 96.0
    HeaderDistance = 48.0
    FooterDistance = 48.0
    DifferentFirstPage = false
    EvenAndOddHeaders = false
}

type Document = {
    Blocks: Block list
    Selection: Selection option
    PageSetup: PageSetup
    Styles: NamedStyle list
    Headers: ParagraphBlock list
    Footers: ParagraphBlock list
    HeaderParts: HeaderFooterPart list
    FooterParts: HeaderFooterPart list
}

type FindMatch = {
    BlockId: BlockId
    Start: int
    End: int
    Text: string
}

type RevisionInfo = {
    BlockId: BlockId
    Start: int
    End: int
    RevisionType: string
    Author: string option
    Date: string option
    Text: string
}

type PageBlockAnchor = {
    PageBlockId: BlockId
    PageNumber: int
}

type PageFieldAnchor = {
    PageFieldBlockId: BlockId
    PageFieldStartOffset: int
    PageFieldCode: string
    PageNumber: int
}

type PageBookmarkAnchor = {
    PageBookmarkName: string
    PageBookmarkBlockId: BlockId
    PageBookmarkStartOffset: int
    PageNumber: int
}

type PageMap = {
    PageCount: int
    BlockAnchors: PageBlockAnchor list
    Fields: PageFieldAnchor list
    Bookmarks: PageBookmarkAnchor list
}

/// Create empty document
let empty = {
    Blocks = [
        Paragraph {
            Id = "block-0"
            Buffer = TextBuffer.empty
            Metadata = defaultMetadata
        }
    ]
    Selection = Some {
        Anchor = { BlockId = "block-0"; Offset = 0 }
        Focus = { BlockId = "block-0"; Offset = 0 }
    }
    PageSetup = defaultPageSetup
    Styles = []
    Headers = []
    Footers = []
    HeaderParts = []
    FooterParts = []
}

/// Find block by ID
let findBlock doc blockId =
    doc.Blocks
    |> List.tryFind (function
        | Paragraph p -> p.Id = blockId
        | Heading (_, p) -> p.Id = blockId
        | Table t -> t.Id = blockId
        | TextBox t -> t.Id = blockId)

/// Try to get the editable paragraph payload from a block.
let tryGetParagraph = function
    | Paragraph p -> Some p
    | Heading (_, p) -> Some p
    | Table _
    | TextBox _ -> None

/// Get paragraph from block. Callers that can encounter tables should use tryGetParagraph.
let getParagraph = function
    | Paragraph p -> p
    | Heading (_, p) -> p
    | Table _
    | TextBox _ -> invalidOp "This block does not have a paragraph payload"

let private getBlockId = function
    | Paragraph p -> p.Id
    | Heading (_, p) -> p.Id
    | Table t -> t.Id
    | TextBox t -> t.Id

let private withParagraphBuffer block para =
    match block with
    | Paragraph _ -> Some (Paragraph para)
    | Heading (level, _) -> Some (Heading (level, para))
    | Table _
    | TextBox _ -> None

/// Replace block in document
let replaceBlock (doc: Document) blockId newBlock =
    { doc with
        Blocks = doc.Blocks |> List.map (fun b ->
            if getBlockId b = blockId then newBlock else b) }

/// Find an editable paragraph anywhere in the document, including table cells.
let findParagraph doc paragraphId =
    doc.Blocks
    |> List.tryPick (function
        | Paragraph p
        | Heading (_, p) when p.Id = paragraphId -> Some p
        | Table t ->
            t.Rows
            |> List.tryPick (fun row ->
                row.Cells
                |> List.tryPick (fun cell ->
                    cell.Content |> List.tryFind (fun p -> p.Id = paragraphId)))
        | TextBox t ->
            t.Content |> List.tryFind (fun p -> p.Id = paragraphId)
        | _ -> None)

/// Replace an editable paragraph anywhere in the document, preserving its container.
let replaceParagraph (doc: Document) paragraphId newPara =
    let replaceInCell (cell: TableCell) =
        { cell with
            Content =
                cell.Content
                |> List.map (fun p -> if p.Id = paragraphId then newPara else p) }

    { doc with
        Blocks =
            doc.Blocks
            |> List.map (function
                | Paragraph p when p.Id = paragraphId -> Paragraph newPara
                | Heading (level, p) when p.Id = paragraphId -> Heading (level, newPara)
                | Table t ->
                    Table {
                        t with
                            Rows =
                                t.Rows
                                |> List.map (fun row ->
                                    { row with Cells = row.Cells |> List.map replaceInCell })
                    }
                | TextBox t ->
                    TextBox { t with Content = t.Content |> List.map (fun p -> if p.Id = paragraphId then newPara else p) }
                | block -> block) }

let private editableParagraphsInOrder doc =
    doc.Blocks
    |> List.collect (function
        | Paragraph p
        | Heading (_, p) -> [p]
        | Table t ->
            t.Rows
            |> List.collect (fun row ->
                row.Cells
                |> List.collect (fun cell -> cell.Content))
        | TextBox t -> t.Content)

let private blockParagraphs (block: Block) =
    match block with
    | Paragraph p
    | Heading (_, p) -> [p]
    | Table table ->
        table.Rows
        |> List.collect (fun row -> row.Cells)
        |> List.collect (fun cell -> cell.Content)
    | TextBox textBox -> textBox.Content

let private blockId (block: Block) =
    match block with
    | Paragraph p
    | Heading (_, p) -> p.Id
    | Table table -> table.Id
    | TextBox textBox -> textBox.Id

let private paragraphPageBreakBefore (doc: Document) (para: ParagraphBlock) =
    let styleBreak =
        para.Metadata.Style
        |> Option.bind (fun styleId ->
            doc.Styles
            |> List.tryFind (fun style -> style.StyleType = "paragraph" && style.Id = styleId)
            |> Option.bind (fun style -> style.PageBreakBefore))

    let defaultBreak =
        doc.Styles
        |> List.tryFind (fun style -> style.StyleType = "defaults")
        |> Option.bind (fun style -> style.PageBreakBefore)

    styleBreak |> Option.orElse defaultBreak |> Option.defaultValue false

let buildPageMap (doc: Document) =
    let mutable physicalPageNumber = 1
    let mutable displayPageStart = 1
    let mutable displayPageStartPhysical = 1
    let mutable blockAnchors: PageBlockAnchor list = []
    let mutable fieldAnchors: PageFieldAnchor list = []
    let mutable bookmarkAnchors: PageBookmarkAnchor list = []

    let displayPageNumber () = displayPageStart + physicalPageNumber - displayPageStartPhysical

    let addParagraphAnchors para =
        let pageNumber = displayPageNumber ()
        para.Buffer.Runs
        |> List.iter (function
            | Text run ->
                run.Style.FieldCode
                |> Option.iter (fun code ->
                    fieldAnchors <- fieldAnchors @ [{ PageFieldBlockId = para.Id; PageFieldStartOffset = run.Start; PageFieldCode = code; PageNumber = pageNumber }])
                run.Style.BookmarkName
                |> Option.iter (fun name ->
                    bookmarkAnchors <- bookmarkAnchors @ [{ PageBookmarkName = name; PageBookmarkBlockId = para.Id; PageBookmarkStartOffset = run.Start; PageNumber = pageNumber }])
            | Image _ -> ())

    doc.Blocks
    |> List.iter (fun block ->
        let paragraphs = blockParagraphs block
        if blockAnchors.Length > 0 && paragraphs |> List.exists (paragraphPageBreakBefore doc) then
            physicalPageNumber <- physicalPageNumber + 1
        paragraphs
        |> List.choose (fun para -> para.Metadata.SectionPageNumberStart)
        |> List.tryLast
        |> Option.iter (fun start ->
            displayPageStart <- max 1 start
            displayPageStartPhysical <- physicalPageNumber)
        let pageNumber = displayPageNumber ()
        blockAnchors <- blockAnchors @ [{ PageBlockId = blockId block; PageNumber = pageNumber }]
        paragraphs |> List.iter addParagraphAnchors
        if paragraphs |> List.exists (fun para -> para.Metadata.PageBreakAfter || para.Metadata.SectionBreak = Some "nextPage") then
            physicalPageNumber <- physicalPageNumber + 1)

    {
        PageCount = max 1 physicalPageNumber
        BlockAnchors = blockAnchors
        Fields = fieldAnchors
        Bookmarks = bookmarkAnchors |> List.distinctBy (fun anchor -> anchor.PageBookmarkName)
    }

let bookmarkInfos (doc: Document) =
    let rawInfos =
        editableParagraphsInOrder doc
        |> List.collect (fun para ->
            para.Buffer.Runs
            |> List.choose (function
                | Text run ->
                    match run.Style.BookmarkStartId, run.Style.BookmarkName with
                    | Some bookmarkId, Some bookmarkName ->
                        Some {
                            Id = bookmarkId
                            Name = bookmarkName
                            BlockId = para.Id
                            StartOffset = run.Start
                            EndOffset = run.Start + run.Length
                        }
                    | _ -> None
                | Image _ -> None))
    rawInfos
    |> List.groupBy (fun info -> info.Name)
    |> List.map (fun (bookmarkName, infos) ->
        let firstInfo = infos |> List.head
        let blockInfos = infos |> List.filter (fun info -> info.BlockId = firstInfo.BlockId)
        {
            Id = firstInfo.Id
            Name = bookmarkName
            BlockId = firstInfo.BlockId
            StartOffset = blockInfos |> List.map (fun info -> info.StartOffset) |> List.min
            EndOffset = blockInfos |> List.map (fun info -> info.EndOffset) |> List.max
        })

let contentControlInfos (doc: Document) =
    editableParagraphsInOrder doc
    |> List.choose (fun para ->
        para.Metadata.ContentControlId
        |> Option.map (fun controlId -> {
            Id = controlId
            Tag = para.Metadata.ContentControlTag |> Option.defaultValue "content-control"
            Alias = para.Metadata.ContentControlAlias |> Option.defaultValue (para.Metadata.ContentControlTag |> Option.defaultValue controlId)
            BlockId = para.Id
        }))
    |> List.distinctBy (fun info -> info.Id, info.BlockId)

let selectBookmark bookmarkId doc =
    match bookmarkInfos doc |> List.tryFind (fun info -> info.Id = bookmarkId || info.Name = bookmarkId) with
    | None -> doc
    | Some info ->
        { doc with
            Selection = Some {
                Anchor = { BlockId = info.BlockId; Offset = info.StartOffset }
                Focus = { BlockId = info.BlockId; Offset = info.EndOffset }
            } }

let selectContentControl controlId doc =
    match contentControlInfos doc |> List.tryFind (fun info -> info.Id = controlId || info.Tag = controlId) with
    | None -> doc
    | Some info ->
        match findParagraph doc info.BlockId with
        | None -> doc
        | Some para ->
            { doc with
                Selection = Some {
                    Anchor = { BlockId = para.Id; Offset = 0 }
                    Focus = { BlockId = para.Id; Offset = TextBuffer.length para.Buffer }
                } }

let removeBookmark bookmarkId doc =
    let clearBookmarkStyle style =
        if style.BookmarkStartId = Some bookmarkId || style.BookmarkName = Some bookmarkId then
            { style with BookmarkStartId = None; BookmarkName = None; BookmarkEndId = None }
        else style

    (doc, editableParagraphsInOrder doc)
    ||> List.fold (fun currentDoc para ->
        let newRuns =
            para.Buffer.Runs
            |> List.map (function
                | Text run -> Text { run with Style = clearBookmarkStyle run.Style }
                | Image run -> Image run)
        replaceParagraph currentDoc para.Id { para with Buffer = { para.Buffer with Runs = newRuns } })

let renameBookmark bookmarkId bookmarkName doc =
    let normalizedName =
        if System.String.IsNullOrWhiteSpace bookmarkName then "Bookmark"
        else bookmarkName.Trim().Replace(" ", "_")

    let renameBookmarkStyle style =
        if style.BookmarkStartId = Some bookmarkId || style.BookmarkName = Some bookmarkId then
            { style with BookmarkName = Some normalizedName }
        else style

    (doc, editableParagraphsInOrder doc)
    ||> List.fold (fun currentDoc para ->
        let newRuns =
            para.Buffer.Runs
            |> List.map (function
                | Text run -> Text { run with Style = renameBookmarkStyle run.Style }
                | Image run -> Image run)
        replaceParagraph currentDoc para.Id { para with Buffer = { para.Buffer with Runs = newRuns } })

let removeContentControl controlId doc =
    (doc, contentControlInfos doc)
    ||> List.fold (fun currentDoc info ->
        if info.Id = controlId || info.Tag = controlId then
            match findParagraph currentDoc info.BlockId with
            | None -> currentDoc
            | Some para ->
                replaceParagraph currentDoc para.Id {
                    para with
                        Metadata = {
                            para.Metadata with
                                ContentControlId = None
                                ContentControlTag = None
                                ContentControlAlias = None
                        }
                }
        else currentDoc)

let updateContentControl controlId newControlId tag alias doc =
    let normalize fallback (value: string) =
        if System.String.IsNullOrWhiteSpace value then fallback else value.Trim()
    let newControlId = normalize controlId newControlId
    let tag = normalize "content-control" tag
    let alias = normalize tag alias

    (doc, contentControlInfos doc)
    ||> List.fold (fun currentDoc info ->
        if info.Id = controlId || info.Tag = controlId then
            match findParagraph currentDoc info.BlockId with
            | None -> currentDoc
            | Some para ->
                replaceParagraph currentDoc para.Id {
                    para with
                        Metadata = {
                            para.Metadata with
                                ContentControlId = Some newControlId
                                ContentControlTag = Some tag
                                ContentControlAlias = Some alias
                        }
                }
        else currentDoc)

let private emptyTableCell () =
    {
        Content = [
            {
                Id = nextBlockId ()
                Buffer = TextBuffer.empty
                Metadata = defaultMetadata
            }
        ]
        ColSpan = 1
        BackgroundColor = None
        Width = None
        PreferredWidth = None
        BorderColor = None
        BorderWidth = None
        Borders = None
        VerticalAlignment = None
        VerticalMerge = None
        PaddingTop = None
        PaddingRight = None
        PaddingBottom = None
        PaddingLeft = None
    }

let private tryFindTableCellByParagraph doc paragraphId =
    doc.Blocks
    |> List.tryPick (function
        | Table table ->
            table.Rows
            |> List.mapi (fun rowIndex row -> rowIndex, row)
            |> List.tryPick (fun (rowIndex, row) ->
                row.Cells
                |> List.mapi (fun colIndex cell -> colIndex, cell)
                |> List.tryPick (fun (colIndex, cell) ->
                    if cell.Content |> List.exists (fun para -> para.Id = paragraphId) then
                        Some (table.Id, rowIndex, colIndex)
                    else
                        None))
        | _ -> None)

let private updateTable (doc: Document) tableId update =
    { doc with
        Blocks =
            doc.Blocks
            |> List.map (function
                | Table table when table.Id = tableId -> Table (update table)
                | block -> block) }

let private currentTableCell doc =
    match doc.Selection with
    | Some sel -> tryFindTableCellByParagraph doc sel.Focus.BlockId
    | None -> None

let private tryGetTableCellParagraph doc tableId rowIndex colIndex =
    doc.Blocks
    |> List.tryPick (function
        | Table table when table.Id = tableId ->
            table.Rows
            |> List.tryItem rowIndex
            |> Option.bind (fun row -> row.Cells |> List.tryItem colIndex)
            |> Option.bind (fun cell -> cell.Content |> List.tryHead)
        | _ -> None)

/// Find block index by ID
let findBlockIndex doc blockId =
    doc.Blocks
    |> List.tryFindIndex (fun b -> getBlockId b = blockId)

/// Get normalized range with block IDs for multi-block selection
let getNormalizedRange (sel: Selection) (doc: Document) =
    let paragraphs = editableParagraphsInOrder doc
    let anchorIdx = paragraphs |> List.tryFindIndex (fun p -> p.Id = sel.Anchor.BlockId)
    let focusIdx = paragraphs |> List.tryFindIndex (fun p -> p.Id = sel.Focus.BlockId)

    match anchorIdx, focusIdx with
    | Some aIdx, Some fIdx ->
        // Determine selection direction
        let (startIdx, startPos, endIdx, endPos) =
            if aIdx < fIdx then
                (aIdx, sel.Anchor, fIdx, sel.Focus)
            elif aIdx > fIdx then
                (fIdx, sel.Focus, aIdx, sel.Anchor)
            else
                // Same block, compare offsets
                if sel.Anchor.Offset <= sel.Focus.Offset then
                    (aIdx, sel.Anchor, fIdx, sel.Focus)
                else
                    (fIdx, sel.Focus, aIdx, sel.Anchor)

        let blockIds =
            paragraphs
            |> List.mapi (fun i p -> (i, p))
            |> List.filter (fun (i, _) -> i >= startIdx && i <= endIdx)
            |> List.map (fun (_, p) -> p.Id)

        (startPos, endPos, blockIds)
    | _ ->
        // Fallback if blocks not found
        (sel.Anchor, sel.Focus, [])

/// Split block at cursor, creating two blocks
let splitBlock doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        match findBlock doc sel.Focus.BlockId with
        | None -> doc
        | Some block ->
            let para = getParagraph block
            let offset = sel.Focus.Offset

            // Split buffer at offset
            let beforeText = TextBuffer.slice 0 offset para.Buffer
            let afterText = TextBuffer.slice offset (TextBuffer.length para.Buffer - offset) para.Buffer

            // Create two buffers
            let beforeBuffer = TextBuffer.ofString beforeText
            let afterBuffer = TextBuffer.ofString afterText

            // Create new block for the second part (carry over metadata)
            let newBlockId = nextBlockId()
            let newPara = { Id = newBlockId; Buffer = afterBuffer; Metadata = para.Metadata }
            let newBlock =
                match block with
                | Paragraph _ -> Paragraph newPara
                | Heading (level, _) -> Heading (level, newPara)
                | Table t -> Table t
                | TextBox t -> TextBox t

            // Update first block with before text (keep existing metadata)
            let updatedPara = { para with Buffer = beforeBuffer }
            let updatedBlock =
                match block with
                | Paragraph _ -> Paragraph updatedPara
                | Heading (level, _) -> Heading (level, updatedPara)
                | Table t -> Table t
                | TextBox t -> TextBox t

            // Insert new block after current block
            match findBlockIndex doc para.Id with
            | None -> doc
            | Some idx ->
                let newBlocks =
                    doc.Blocks
                    |> List.mapi (fun i b ->
                        if i = idx then [updatedBlock; newBlock]
                        else [b])
                    |> List.concat

                // Move cursor to start of new block
                let newSel = {
                    Anchor = { BlockId = newBlockId; Offset = 0 }
                    Focus = { BlockId = newBlockId; Offset = 0 }
                }
                { doc with Blocks = newBlocks; Selection = Some newSel }

/// Merge block with previous block (for Backspace at start)
let mergeWithPrevious doc blockId =
    match findBlockIndex doc blockId with
    | None -> doc
    | Some idx when idx = 0 -> doc  // First block, can't merge
    | Some idx ->
        match findBlock doc blockId with
        | None -> doc
        | Some currentBlock ->
            let currentPara = getParagraph currentBlock
            let prevBlock = doc.Blocks.[idx - 1]
            let prevPara = getParagraph prevBlock

            // Merge buffers: previous + current
            let prevText = TextBuffer.toString prevPara.Buffer
            let currentText = TextBuffer.toString currentPara.Buffer
            let mergedBuffer = TextBuffer.ofString (prevText + currentText)

            // Update previous block with merged buffer
            let mergedPara = { prevPara with Buffer = mergedBuffer }
            let mergedBlock =
                match prevBlock with
                | Paragraph _ -> Paragraph mergedPara
                | Heading (level, _) -> Heading (level, mergedPara)
                | Table t -> Table t
                | TextBox t -> TextBox t

            // Remove current block
            let newBlocks =
                doc.Blocks
                |> List.mapi (fun i b ->
                    if i = idx - 1 then Some mergedBlock
                    elif i = idx then None
                    else Some b)
                |> List.choose id

            // Move cursor to junction point
            let newSel = {
                Anchor = { BlockId = prevPara.Id; Offset = prevText.Length }
                Focus = { BlockId = prevPara.Id; Offset = prevText.Length }
            }
            { doc with Blocks = newBlocks; Selection = Some newSel }

/// Merge block with next block (for Delete at end)
let mergeWithNext doc blockId =
    match findBlockIndex doc blockId with
    | None -> doc
    | Some idx when idx >= doc.Blocks.Length - 1 -> doc  // Last block, can't merge
    | Some idx ->
        match findBlock doc blockId with
        | None -> doc
        | Some currentBlock ->
            let currentPara = getParagraph currentBlock
            let nextBlock = doc.Blocks.[idx + 1]
            let nextPara = getParagraph nextBlock

            // Merge buffers: current + next
            let currentText = TextBuffer.toString currentPara.Buffer
            let nextText = TextBuffer.toString nextPara.Buffer
            let mergedBuffer = TextBuffer.ofString (currentText + nextText)

            // Update current block with merged buffer
            let mergedPara = { currentPara with Buffer = mergedBuffer }
            let mergedBlock =
                match currentBlock with
                | Paragraph _ -> Paragraph mergedPara
                | Heading (level, _) -> Heading (level, mergedPara)
                | Table t -> Table t
                | TextBox t -> TextBox t

            // Remove next block
            let newBlocks =
                doc.Blocks
                |> List.mapi (fun i b ->
                    if i = idx then Some mergedBlock
                    elif i = idx + 1 then None
                    else Some b)
                |> List.choose id

            // Keep cursor at current position
            { doc with Blocks = newBlocks }

/// Delete selected text across blocks
let deleteSelection doc =
    match doc.Selection with
    | None -> doc
    | Some sel when sel.IsCollapsed -> doc  // No selection to delete
    | Some sel ->
        let (startPos, endPos, blockIds) = getNormalizedRange sel doc

        match blockIds with
        | [] -> doc
        | [singleBlock] when singleBlock = startPos.BlockId && singleBlock = endPos.BlockId ->
            // Single block selection
            match findParagraph doc singleBlock with
            | None -> doc
            | Some para ->
                let len = endPos.Offset - startPos.Offset
                let newBuffer = TextBuffer.delete startPos.Offset len para.Buffer
                let newPara = { para with Buffer = newBuffer }
                let newDoc = replaceParagraph doc para.Id newPara

                // Move cursor to start of selection
                let newSel = {
                    Anchor = startPos
                    Focus = startPos
                }
                { newDoc with Selection = Some newSel }
        | _ ->
            // Multi-block selection
            // Keep text before selection start in first block
            // Keep text after selection end in last block
            // Delete middle blocks

            if blockIds.IsEmpty then doc
            else
                let firstBlockId = List.head blockIds
                let lastBlockId = List.last blockIds

                match findBlock doc firstBlockId, findBlock doc lastBlockId with
                | Some firstBlock, Some lastBlock ->
                    let firstPara = getParagraph firstBlock
                    let lastPara = getParagraph lastBlock

                    // Get text before selection start
                    let beforeText = TextBuffer.slice 0 startPos.Offset firstPara.Buffer

                    // Get text after selection end
                    let lastBufferLen = TextBuffer.length lastPara.Buffer
                    let afterText = TextBuffer.slice endPos.Offset (lastBufferLen - endPos.Offset) lastPara.Buffer

                    // Merge into single buffer
                    let mergedBuffer = TextBuffer.ofString (beforeText + afterText)

                    // Update first block with merged content
                    let mergedPara = { firstPara with Buffer = mergedBuffer }
                    let mergedBlock =
                        match firstBlock with
                        | Paragraph _ -> Paragraph mergedPara
                        | Heading (level, _) -> Heading (level, mergedPara)
                        | Table t -> Table t
                        | TextBox t -> TextBox t

                    // Remove all selected editable paragraphs and replace the first with merged content.
                    // Table-cell paragraphs are not top-level blocks, so filter them inside table cells.
                    let removeIds = Set.ofList (List.tail blockIds)
                    let newBlocks =
                        doc.Blocks
                        |> List.choose (function
                            | Paragraph p when Set.contains p.Id removeIds -> None
                            | Heading (_, p) when Set.contains p.Id removeIds -> None
                            | Paragraph p when p.Id = firstBlockId -> Some mergedBlock
                            | Heading (_, p) when p.Id = firstBlockId -> Some mergedBlock
                            | Table t ->
                                let rows =
                                    t.Rows
                                    |> List.map (fun row ->
                                        { row with
                                            Cells =
                                                row.Cells
                                                |> List.map (fun cell ->
                                                    { cell with Content = cell.Content |> List.filter (fun p -> not (Set.contains p.Id removeIds)) }) })
                                Some (Table { t with Rows = rows })
                            | block -> Some block)

                    // Move cursor to merge point
                    let newSel = {
                        Anchor = { BlockId = firstBlockId; Offset = String.length beforeText }
                        Focus = { BlockId = firstBlockId; Offset = String.length beforeText }
                    }

                    { doc with Blocks = newBlocks; Selection = Some newSel }
                | _ -> doc

/// Delete forward (Delete key)
let deleteForward doc =
    match doc.Selection with
    | None -> doc
    | Some sel when not sel.IsCollapsed ->
        // Delete selection if active
        deleteSelection doc
    | Some sel ->
        match findParagraph doc sel.Focus.BlockId with
        | None -> doc
        | Some para ->
            let maxOffset = TextBuffer.length para.Buffer

            if sel.Focus.Offset >= maxOffset then
                mergeWithNext doc sel.Focus.BlockId
            else
                let newBuffer = TextBuffer.delete sel.Focus.Offset 1 para.Buffer
                let newPara = { para with Buffer = newBuffer }
                replaceParagraph doc para.Id newPara

/// Move cursor up to previous block
let moveCursorUp doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        match findBlockIndex doc sel.Focus.BlockId with
        | None -> doc
        | Some idx when idx = 0 -> doc  // Already at first block
        | Some idx ->
            let prevBlock = doc.Blocks.[idx - 1]
            let prevPara = getParagraph prevBlock
            let prevLen = TextBuffer.length prevPara.Buffer

            // Move to end of previous block
            let newSel = {
                Anchor = { BlockId = prevPara.Id; Offset = prevLen }
                Focus = { BlockId = prevPara.Id; Offset = prevLen }
            }
            { doc with Selection = Some newSel }

/// Move cursor down to next block
let moveCursorDown doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        match findBlockIndex doc sel.Focus.BlockId with
        | None -> doc
        | Some idx when idx >= doc.Blocks.Length - 1 -> doc  // Already at last block
        | Some idx ->
            let nextBlock = doc.Blocks.[idx + 1]
            let nextPara = getParagraph nextBlock

            // Move to start of next block
            let newSel = {
                Anchor = { BlockId = nextPara.Id; Offset = 0 }
                Focus = { BlockId = nextPara.Id; Offset = 0 }
            }
            { doc with Selection = Some newSel }

/// Move cursor to start of current block (Home key)
let moveCursorToLineStart doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        let newSel = {
            Anchor = { sel.Focus with Offset = 0 }
            Focus = { sel.Focus with Offset = 0 }
        }
        { doc with Selection = Some newSel }

/// Move cursor to end of current block (End key)
let moveCursorToLineEnd doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        match findParagraph doc sel.Focus.BlockId with
        | None -> doc
        | Some para ->
            let maxOffset = TextBuffer.length para.Buffer
            let newSel = {
                Anchor = { sel.Focus with Offset = maxOffset }
                Focus = { sel.Focus with Offset = maxOffset }
            }
            { doc with Selection = Some newSel }


/// Insert text at cursor
let insertText text doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        // Delete selection if not collapsed
        let doc' = if sel.IsCollapsed then doc else deleteSelection doc

        // Insert text at cursor (which is now at the selection start after deletion)
        match doc'.Selection with
        | None -> doc'
        | Some sel' ->
            match findParagraph doc' sel'.Focus.BlockId with
            | None -> doc'
            | Some para ->
                let newBuffer = TextBuffer.insert sel'.Focus.Offset text para.Buffer
                let newPara = { para with Buffer = newBuffer }

                let newDoc = replaceParagraph doc' para.Id newPara

                // Move cursor forward
                let newSel = {
                    Anchor = { sel'.Focus with Offset = sel'.Focus.Offset + text.Length }
                    Focus = { sel'.Focus with Offset = sel'.Focus.Offset + text.Length }
                }
                { newDoc with Selection = Some newSel }

let insertField fieldCode result doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        let doc' = insertText result doc
        match doc'.Selection with
        | None -> doc'
        | Some sel' ->
            let startOffset = max 0 (sel'.Focus.Offset - result.Length)
            match findParagraph doc' sel'.Focus.BlockId with
            | None -> doc'
            | Some para ->
                let newBuffer =
                    TextBuffer.applyStyle startOffset result.Length (fun style ->
                        { style with FieldCode = Some fieldCode; FieldResult = Some result }) para.Buffer
                let newPara = { para with Buffer = newBuffer }
                replaceParagraph doc' para.Id newPara

let private escapeXml (text: string) =
    text
        .Replace("&", "&amp;")
        .Replace("<", "&lt;")
        .Replace(">", "&gt;")
        .Replace("\"", "&quot;")
        .Replace("'", "&apos;")

let insertEquation equationText doc =
    let text = "[Equation]"
    let equationXml = sprintf "<m:oMath><m:r><m:t>%s</m:t></m:r></m:oMath>" (escapeXml equationText)
    match doc.Selection with
    | None -> doc
    | Some sel ->
        let doc' = insertText text doc
        match doc'.Selection with
        | None -> doc'
        | Some sel' ->
            let startOffset = max 0 (sel'.Focus.Offset - text.Length)
            match findParagraph doc' sel'.Focus.BlockId with
            | None -> doc'
            | Some para ->
                let newBuffer =
                    TextBuffer.applyStyle startOffset text.Length (fun style ->
                        { style with EquationXml = Some equationXml; FontFamily = Some "Cambria Math" }) para.Buffer
                let newPara = { para with Buffer = newBuffer }
                replaceParagraph doc' para.Id newPara

/// Insert image at cursor position
let insertImageAtCursor src width height alt doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        // Delete selection if not collapsed
        let doc' = if sel.IsCollapsed then doc else deleteSelection doc

        // Insert image at cursor
        match doc'.Selection with
        | None -> doc'
        | Some sel' ->
            match findParagraph doc' sel'.Focus.BlockId with
            | None -> doc'
            | Some para ->
                let newBuffer = TextBuffer.insertImage sel'.Focus.Offset src width height alt para.Buffer
                let newPara = { para with Buffer = newBuffer }

                let newDoc = replaceParagraph doc' para.Id newPara

                // Move cursor forward by 1 (image occupies 1 character position)
                let newSel = {
                    Anchor = { sel'.Focus with Offset = sel'.Focus.Offset + 1 }
                    Focus = { sel'.Focus with Offset = sel'.Focus.Offset + 1 }
                }
                { newDoc with Selection = Some newSel }

/// Resize image at specified position
let resizeImage blockId offset newWidth newHeight doc =
    match findParagraph doc blockId with
    | None -> doc
    | Some para ->
        let newBuffer = TextBuffer.resizeImage offset newWidth newHeight para.Buffer
        let newPara = { para with Buffer = newBuffer }

        replaceParagraph doc para.Id newPara

let resizeImageNearCursor newWidth newHeight doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        match findParagraph doc sel.Focus.BlockId with
        | None -> doc
        | Some para ->
            let imageOffset =
                para.Buffer.Runs
                |> List.choose (function
                    | TextBuffer.Image img when img.Start <= sel.Focus.Offset -> Some img.Start
                    | _ -> None)
                |> List.sortDescending
                |> List.tryHead
            match imageOffset with
            | None -> doc
            | Some offset -> resizeImage para.Id offset newWidth newHeight doc

/// Move an image run to a new document position.
let moveImage blockId offset targetBlockId targetOffset doc =
    match findParagraph doc blockId with
    | None -> doc
    | Some sourcePara ->
        match sourcePara.Buffer.Runs |> List.tryPick (function TextBuffer.Image img when img.Start = offset -> Some img | _ -> None) with
        | None -> doc
        | Some img ->
            let sourceBuffer = TextBuffer.delete offset 1 sourcePara.Buffer
            let docWithoutImage = replaceParagraph doc sourcePara.Id { sourcePara with Buffer = sourceBuffer }
            match findParagraph docWithoutImage targetBlockId with
            | None -> docWithoutImage
            | Some targetPara ->
                let insertOffset =
                    if blockId = targetBlockId && targetOffset > offset then targetOffset - 1
                    else targetOffset
                let boundedOffset = max 0 (min (TextBuffer.length targetPara.Buffer) insertOffset)
                let targetBuffer = TextBuffer.insertImageWithMetadata boundedOffset img.Src img.Linked img.Width img.Height img.Alt img.Name img.Title img.Floating img.Wrap img.PositionX img.PositionY img.PositionHRelativeFrom img.PositionVRelativeFrom img.BehindDoc img.AllowOverlap img.LayoutInCell img.DistanceTop img.DistanceBottom img.DistanceLeft img.DistanceRight img.RelativeHeight img.Locked img.SimplePosition img.SimplePositionX img.SimplePositionY img.CropLeft img.CropTop img.CropRight img.CropBottom targetPara.Buffer
                let nextDoc = replaceParagraph docWithoutImage targetPara.Id { targetPara with Buffer = targetBuffer }
                { nextDoc with Selection = Some { Anchor = { BlockId = targetBlockId; Offset = boundedOffset + 1 }; Focus = { BlockId = targetBlockId; Offset = boundedOffset + 1 } } }

/// Delete backward (backspace)
let deleteBackward doc =
    match doc.Selection with
    | None -> doc
    | Some sel when not sel.IsCollapsed ->
        // Delete selection if active
        deleteSelection doc
    | Some sel when sel.Focus.Offset = 0 ->
        // At start of block, merge with previous
        mergeWithPrevious doc sel.Focus.BlockId
    | Some sel ->
        match findParagraph doc sel.Focus.BlockId with
        | None -> doc
        | Some para ->
            let newBuffer = TextBuffer.delete (sel.Focus.Offset - 1) 1 para.Buffer
            let newPara = { para with Buffer = newBuffer }

            let newDoc = replaceParagraph doc para.Id newPara

            // Move cursor back
            let newSel = {
                Anchor = { sel.Focus with Offset = sel.Focus.Offset - 1 }
                Focus = { sel.Focus with Offset = sel.Focus.Offset - 1 }
            }
            { newDoc with Selection = Some newSel }

/// Set cursor position
let setCursor blockId offset doc =
    { doc with
        Selection = Some {
            Anchor = { BlockId = blockId; Offset = offset }
            Focus = { BlockId = blockId; Offset = offset }
        } }

let insertTableRowAfterSelection doc =
    match currentTableCell doc with
    | None -> doc
    | Some (tableId, rowIndex, _) ->
        updateTable doc tableId (fun table ->
            let columnCount =
                table.Rows
                |> List.tryItem rowIndex
                |> Option.map (fun row -> row.Cells.Length)
                |> Option.defaultValue 1
                |> max 1
            let newRow = { Cells = [ for _ in 1 .. columnCount -> emptyTableCell () ]; IsHeader = false; Height = None; HeightRule = None; CanSplit = None; CellSpacing = None; PaddingTop = None; PaddingRight = None; PaddingBottom = None; PaddingLeft = None }
            { table with
                Rows =
                    table.Rows
                    |> List.mapi (fun i row -> if i = rowIndex then [row; newRow] else [row])
                    |> List.concat })

let insertTableColumnAfterSelection doc =
    match currentTableCell doc with
    | None -> doc
    | Some (tableId, _, colIndex) ->
        updateTable doc tableId (fun table ->
            { table with
                Rows =
                    table.Rows
                    |> List.map (fun row ->
                        { row with
                            Cells =
                                row.Cells
                                |> List.mapi (fun i cell -> if i = colIndex then [cell; emptyTableCell ()] else [cell])
                                |> List.concat }) })

let deleteSelectedTableRow doc =
    match currentTableCell doc with
    | None -> doc
    | Some (tableId, rowIndex, colIndex) ->
        let updatedDoc = updateTable doc tableId (fun table ->
            if table.Rows.Length <= 1 then table
            else { table with Rows = table.Rows |> List.mapi (fun i row -> i, row) |> List.choose (fun (i, row) -> if i = rowIndex then None else Some row) })

        match doc.Blocks |> List.tryPick (function Table table when table.Id = tableId -> Some table.Rows.Length | _ -> None) with
        | Some oldLength when oldLength > 1 ->
            let targetRow = min rowIndex (oldLength - 2)
            match tryGetTableCellParagraph updatedDoc tableId targetRow colIndex with
            | Some para -> setCursor para.Id 0 updatedDoc
            | None -> updatedDoc
        | _ -> updatedDoc

let deleteSelectedTableColumn doc =
    match currentTableCell doc with
    | None -> doc
    | Some (tableId, _, colIndex) ->
        let updatedDoc = updateTable doc tableId (fun table ->
            let maxColumns = table.Rows |> List.map (fun row -> row.Cells.Length) |> List.fold max 0
            if maxColumns <= 1 then table
            else
                { table with
                    Rows =
                        table.Rows
                        |> List.map (fun row ->
                            { row with
                                Cells =
                                    row.Cells
                                    |> List.mapi (fun i cell -> i, cell)
                                    |> List.choose (fun (i, cell) -> if i = colIndex then None else Some cell) }) })

        match currentTableCell updatedDoc with
        | Some _ -> updatedDoc
        | None ->
            match currentTableCell doc with
            | Some (_, rowIndex, _) ->
                let targetCol = max 0 (colIndex - 1)
                match tryGetTableCellParagraph updatedDoc tableId rowIndex targetCol with
                | Some para -> setCursor para.Id 0 updatedDoc
                | None -> updatedDoc
            | None -> updatedDoc

let insertTable rows columns doc =
    let rowCount = max 1 rows
    let columnCount = max 1 columns
    let table = {
        Id = nextBlockId ()
        Rows = [ for _ in 1 .. rowCount -> { Cells = [ for _ in 1 .. columnCount -> emptyTableCell () ]; IsHeader = false; Height = None; HeightRule = None; CanSplit = None; CellSpacing = None; PaddingTop = None; PaddingRight = None; PaddingBottom = None; PaddingLeft = None } ]
        PreferredWidth = None
        LayoutMode = None
        GridColumns = []
        Borders = None
        BackgroundColor = None
        CellSpacing = None
        PaddingTop = None
        PaddingRight = None
        PaddingBottom = None
        PaddingLeft = None
        Floating = false
    }
    let firstCellPara = table.Rows.Head.Cells.Head.Content.Head
    let tableBlock = Table table
    let blocks =
        match doc.Selection with
        | Some sel ->
            let insertAfterIndex =
                doc.Blocks
                |> List.tryFindIndex (function
                    | Paragraph p -> p.Id = sel.Focus.BlockId
                    | Heading (_, p) -> p.Id = sel.Focus.BlockId
                    | Table t -> t.Id = sel.Focus.BlockId
                    | TextBox t -> t.Id = sel.Focus.BlockId)
            match insertAfterIndex with
            | Some index ->
                let before = doc.Blocks |> List.take (index + 1)
                let after = doc.Blocks |> List.skip (index + 1)
                before @ [tableBlock] @ after
            | None -> doc.Blocks @ [tableBlock]
        | None -> doc.Blocks @ [tableBlock]
    { doc with Blocks = blocks } |> setCursor firstCellPara.Id 0

let private insertTopLevelBlockAfterSelection block doc =
    let blocks =
        match doc.Selection with
        | Some sel ->
            let insertAfterIndex =
                doc.Blocks
                |> List.tryFindIndex (function
                    | Paragraph p -> p.Id = sel.Focus.BlockId
                    | Heading (_, p) -> p.Id = sel.Focus.BlockId
                    | Table t -> t.Id = sel.Focus.BlockId
                    | TextBox t -> t.Id = sel.Focus.BlockId)
            match insertAfterIndex with
            | Some index ->
                let before = doc.Blocks |> List.take (index + 1)
                let after = doc.Blocks |> List.skip (index + 1)
                before @ [block] @ after
            | None -> doc.Blocks @ [block]
        | None -> doc.Blocks @ [block]
    { doc with Blocks = blocks }

let insertTextBox text doc =
    let para = {
        Id = nextBlockId ()
        Buffer = TextBuffer.ofString text
        Metadata = defaultMetadata
    }
    let textBox = TextBox {
        Id = nextBlockId ()
        Content = [para]
        X = Some 96.0
        Y = Some 96.0
        PositionHRelativeFrom = None
        PositionVRelativeFrom = None
        Width = 240.0
        Height = 96.0
        FillColor = Some "#FFFFFF"
        BorderColor = Some "#808080"
        BorderWidth = Some 1.0
        RelativeHeight = None
    }
    insertTopLevelBlockAfterSelection textBox doc

let private textRunsInParagraph para =
    para.Buffer.Runs
    |> List.choose (function
        | TextBuffer.Text run -> Some run
        | TextBuffer.Image _ -> None)

let private textRunsInBlock block =
    match block with
    | Paragraph para
    | Heading (_, para) -> textRunsInParagraph para
    | Table table ->
        table.Rows
        |> List.collect (fun row -> row.Cells)
        |> List.collect (fun cell -> cell.Content)
        |> List.collect textRunsInParagraph
    | TextBox textBox ->
        textBox.Content |> List.collect textRunsInParagraph

let private nextSequenceResult label doc =
    let prefix = sprintf " SEQ %s " label
    let maxExisting =
        doc.Blocks
        |> List.collect textRunsInBlock
        |> List.choose (fun run ->
            match run.Style.FieldCode, run.Style.FieldResult with
            | Some code, Some result when code.StartsWith(prefix) ->
                match System.Int32.TryParse(result) with
                | true, value -> Some value
                | false, _ -> None
            | _ -> None)
        |> function
            | [] -> 0
            | values -> List.max values
    string (maxExisting + 1)

let private nextBookmarkId doc =
    let maxExisting =
        doc.Blocks
        |> List.collect textRunsInBlock
        |> List.collect (fun run -> [run.Style.BookmarkStartId; run.Style.BookmarkEndId])
        |> List.choose id
        |> List.choose (fun value ->
            match System.Int32.TryParse(value) with
            | true, parsed -> Some parsed
            | false, _ -> None)
        |> function
            | [] -> 0
            | values -> List.max values
    string (maxExisting + 1)

let private nextRevisionId doc =
    let maxExisting =
        doc.Blocks
        |> List.collect textRunsInBlock
        |> List.choose (fun run -> run.Style.RevisionId)
        |> List.choose (fun value ->
            match System.Int32.TryParse(value) with
            | true, parsed -> Some parsed
            | false, _ -> None)
        |> function
            | [] -> 0
            | values -> List.max values
    string (maxExisting + 1)

let insertCaption label captionText doc =
    let result = nextSequenceResult label doc
    let bookmarkId = nextBookmarkId doc
    let bookmarkName = sprintf "_Ref%s" bookmarkId
    let prefix = label + " "
    let suffix = ": " + captionText
    let fullText = prefix + result + suffix
    let buffer =
        TextBuffer.ofString fullText
        |> TextBuffer.applyStyle prefix.Length result.Length (fun style ->
            { style with
                BookmarkStartId = Some bookmarkId
                BookmarkName = Some bookmarkName
                BookmarkEndId = Some bookmarkId
                FieldCode = Some (sprintf " SEQ %s \\* ARABIC " label)
                FieldResult = Some result })
    let para = Paragraph {
        Id = nextBlockId ()
        Buffer = buffer
        Metadata = { defaultMetadata with Style = Some "Caption" }
    }
    insertTopLevelBlockAfterSelection para doc

let insertCrossReferenceToLatestCaption label doc =
    let prefix = sprintf " SEQ %s " label
    let latest =
        doc.Blocks
        |> List.collect textRunsInBlock
        |> List.choose (fun run ->
            match run.Style.FieldCode, run.Style.FieldResult, run.Style.BookmarkName with
            | Some code, Some result, Some bookmarkName when code.StartsWith(prefix) ->
                match System.Int32.TryParse(result) with
                | true, number -> Some (number, result, bookmarkName)
                | false, _ -> None
            | _ -> None)
        |> List.sortBy (fun (number, _, _) -> number)
        |> List.tryLast
    match latest with
    | None -> doc
    | Some (_, result, bookmarkName) ->
        insertField (sprintf " REF %s \\h " bookmarkName) result doc

let private fieldCodeParts (code: string) =
    code.Trim().Split([| ' '; '\t'; '\r'; '\n' |], System.StringSplitOptions.RemoveEmptyEntries)
    |> Array.toList

let private dateFieldFormat (code: string) =
    let formatSwitch = code.IndexOf("\\@")
    if formatSwitch < 0 then "yyyy-MM-dd"
    else
        let firstQuote = code.IndexOf('"', formatSwitch)
        if firstQuote < 0 then "yyyy-MM-dd"
        else
            let secondQuote = code.IndexOf('"', firstQuote + 1)
            if secondQuote <= firstQuote then "yyyy-MM-dd"
            else code.Substring(firstQuote + 1, secondQuote - firstQuote - 1)

let private fieldSwitchValue (switchName: string) (code: string) =
    let switchIndex = code.IndexOf(switchName)
    if switchIndex < 0 then None
    else
        let firstQuote = code.IndexOf('"', switchIndex)
        if firstQuote < 0 then None
        else
            let secondQuote = code.IndexOf('"', firstQuote + 1)
            if secondQuote <= firstQuote then None
            else Some (code.Substring(firstQuote + 1, secondQuote - firstQuote - 1))

let private paragraphText para = Rope.toString para.Buffer.Rope

let private tocHeadingRange (code: string) =
    let parseBound fallback (text: string) =
        match System.Int32.TryParse(text: string) with
        | true, value -> max 1 (min 9 value)
        | false, _ -> fallback

    match fieldSwitchValue "\\o" code with
    | None -> 1, 3
    | Some value ->
        match value.Split('-') with
        | parts when parts.Length = 2 ->
            let first = parseBound 1 parts[0]
            let last = parseBound 3 parts[1]
            min first last, max first last
        | _ -> 1, 3

let private tocHeadingLines code pageMap doc =
    let minLevel, maxLevel = tocHeadingRange code
    doc.Blocks
    |> List.choose (function
        | Heading (level, para) when level >= minLevel && level <= maxLevel ->
            let pageNumber =
                pageMap.BlockAnchors
                |> List.tryFind (fun anchor -> anchor.PageBlockId = para.Id)
                |> Option.map (fun anchor -> anchor.PageNumber)
                |> Option.defaultValue 1
            Some (String.replicate (max 0 (level - 1)) "  " + paragraphText para + "\t" + string pageNumber)
        | _ -> None)
    |> String.concat "\n"

let private tocCaptionLines label doc =
    let hasSequenceField para =
        para.Buffer.Runs
        |> List.exists (function
            | Text run ->
                match run.Style.FieldCode with
                | Some code -> (fieldCodeParts code) |> List.exists ((=) label)
                | None -> false
            | Image _ -> false)

    doc.Blocks
    |> List.choose (function
        | Paragraph para
        | Heading (_, para) when hasSequenceField para -> Some (paragraphText para)
        | _ -> None)
    |> String.concat "\n"

let private collectSequenceRefreshTargets doc =
    let mutable counters: Map<string, int> = Map.empty
    let mutable sequences: Map<string, string list> = Map.empty
    let mutable bookmarks: Map<string, string> = Map.empty

    let addSequence label bookmarkName =
        let next = (counters |> Map.tryFind label |> Option.defaultValue 0) + 1
        counters <- counters |> Map.add label next
        let result = string next
        sequences <- sequences |> Map.add label ((sequences |> Map.tryFind label |> Option.defaultValue []) @ [result])
        bookmarkName |> Option.iter (fun name -> bookmarks <- bookmarks |> Map.add name result)

    editableParagraphsInOrder doc
    |> List.iter (fun para ->
        para.Buffer.Runs
        |> List.iter (function
            | Text run ->
                match run.Style.FieldCode |> Option.map fieldCodeParts with
                | Some ("SEQ" :: label :: _) -> addSequence label run.Style.BookmarkName
                | _ -> ()
            | Image _ -> ()))
    sequences, bookmarks

let refreshFieldsWithPageMap pageMap doc =
    let allSequenceResults, allBookmarkResults = collectSequenceRefreshTargets doc
    let mutable sequenceCounters: Map<string, int> = Map.empty
    let mutable sequenceResults: Map<string, string list> = Map.empty

    let nextSeqResult label =
        let next = (sequenceCounters |> Map.tryFind label |> Option.defaultValue 0) + 1
        sequenceCounters <- sequenceCounters |> Map.add label next
        let result = string next
        let existing = sequenceResults |> Map.tryFind label |> Option.defaultValue []
        sequenceResults <- sequenceResults |> Map.add label (existing @ [result])
        result

    let fieldPage blockId startOffset =
        pageMap.Fields
        |> List.tryFind (fun anchor -> anchor.PageFieldBlockId = blockId && anchor.PageFieldStartOffset = startOffset)
        |> Option.map (fun anchor -> anchor.PageNumber)
        |> Option.defaultValue 1

    let bookmarkPage bookmarkName =
        pageMap.Bookmarks
        |> List.tryFind (fun anchor -> anchor.PageBookmarkName = bookmarkName)
        |> Option.map (fun anchor -> anchor.PageNumber)
        |> Option.defaultValue 1

    let sequencePageResults =
        let mutable counters: Map<string, int> = Map.empty
        let mutable results: Map<string, string list> = Map.empty
        editableParagraphsInOrder doc
        |> List.iter (fun para ->
            para.Buffer.Runs
            |> List.iter (function
                | Text run ->
                    match run.Style.FieldCode |> Option.map fieldCodeParts with
                    | Some ("SEQ" :: label :: _) ->
                        let next = (counters |> Map.tryFind label |> Option.defaultValue 0) + 1
                        counters <- counters |> Map.add label next
                        let pageNumber = run.Style.BookmarkName |> Option.map bookmarkPage |> Option.defaultValue (fieldPage para.Id run.Start)
                        let existing = results |> Map.tryFind label |> Option.defaultValue []
                        results <- results |> Map.add label (existing @ [sprintf "%d\t%d" next pageNumber])
                    | _ -> ()
                | Image _ -> ()))
        results

    let resultForField blockId startOffset style text =
        match style.FieldCode |> Option.map (fun code -> code, fieldCodeParts code) with
        | Some (code, "DATE" :: _) -> System.DateTime.Now.ToString(dateFieldFormat code)
        | Some (code, "TOC" :: _) ->
            match fieldSwitchValue "\\c" code with
            | Some label -> sequencePageResults |> Map.tryFind label |> Option.defaultValue [] |> String.concat "\n"
            | None -> tocHeadingLines code pageMap doc
        | Some (_, "SEQ" :: label :: _) ->
            nextSeqResult label
        | Some (_, "REF" :: bookmarkName :: _) ->
            allBookmarkResults |> Map.tryFind bookmarkName |> Option.defaultValue text
        | Some (_, "PAGE" :: _) -> string (fieldPage blockId startOffset)
        | Some (_, "PAGEREF" :: bookmarkName :: _) -> string (bookmarkPage bookmarkName)
        | Some (_, "NUMPAGES" :: _) -> string pageMap.PageCount
        | _ -> text

    (doc, editableParagraphsInOrder doc)
    ||> List.fold (fun currentDoc para ->
        let mutable buffer = para.Buffer
        let mutable shift = 0
        for run in para.Buffer.Runs do
            match run with
            | Image _ -> ()
            | Text textRun ->
                match textRun.Style.FieldCode with
                | None -> ()
                | Some _ ->
                    let start = textRun.Start + shift
                    let oldText = TextBuffer.slice start textRun.Length buffer
                    let newText = resultForField para.Id textRun.Start textRun.Style oldText
                    if newText <> oldText then
                        let newStyle = { textRun.Style with FieldResult = Some newText }
                        buffer <- buffer |> TextBuffer.delete start textRun.Length |> TextBuffer.insert start newText |> TextBuffer.applyStyle start newText.Length (fun _ -> newStyle)
                        shift <- shift + newText.Length - textRun.Length
                    else
                        buffer <- buffer |> TextBuffer.applyStyle start textRun.Length (fun style -> { style with FieldResult = Some oldText })
        replaceParagraph currentDoc para.Id { para with Buffer = buffer })

let refreshFields doc =
    refreshFieldsWithPageMap (buildPageMap doc) doc

let setCurrentTextBoxStyle x y width height fillColor borderColor borderWidth doc =
    let selectedTextBoxId =
        doc.Selection
        |> Option.bind (fun sel ->
            doc.Blocks
            |> List.tryPick (function
                | TextBox textBox when textBox.Id = sel.Focus.BlockId -> Some textBox.Id
                | _ -> None))
    let fallbackTextBoxId =
        doc.Blocks
        |> List.rev
        |> List.tryPick (function
            | TextBox textBox -> Some textBox.Id
            | _ -> None)
    match selectedTextBoxId |> Option.orElse fallbackTextBoxId with
    | None -> doc
    | Some textBoxId ->
        { doc with
            Blocks =
                doc.Blocks
                |> List.map (function
                    | TextBox textBox when textBox.Id = textBoxId ->
                        TextBox { textBox with
                                    X = Some x
                                    Y = Some y
                                    Width = width
                                    Height = height
                                    FillColor = if System.String.IsNullOrWhiteSpace fillColor then None else Some fillColor
                                    BorderColor = if System.String.IsNullOrWhiteSpace borderColor then None else Some borderColor
                                    BorderWidth = if borderWidth > 0.0 then Some borderWidth else None }
                    | block -> block) }

let setSelectedTableCellStyle backgroundColor width borderColor borderWidth verticalAlignment paddingTop paddingRight paddingBottom paddingLeft doc =
    match currentTableCell doc with
    | None -> doc
    | Some (tableId, rowIndex, colIndex) ->
        updateTable doc tableId (fun table ->
            { table with
                Rows =
                    table.Rows
                    |> List.mapi (fun r row ->
                        if r <> rowIndex then row
                        else
                            { row with
                                Cells =
                                    row.Cells
                                    |> List.mapi (fun c cell ->
                                        if c <> colIndex then cell
                                        else
                                            { cell with
                                                BackgroundColor = backgroundColor
                                                Width = width
                                                BorderColor = borderColor
                                                BorderWidth = borderWidth
                                                VerticalAlignment = verticalAlignment
                                                PaddingTop = paddingTop
                                                PaddingRight = paddingRight
                                                PaddingBottom = paddingBottom
                                                PaddingLeft = paddingLeft }) }) })

let private headingFontSize level =
    match level with
    | 1 -> 24.0
    | 2 -> 18.0
    | 3 -> 14.0
    | _ -> 11.0

let private styleHeadingBuffer level buffer =
    let len = TextBuffer.length buffer
    if len = 0 then buffer
    else
        TextBuffer.applyStyle 0 len (fun style ->
            { style with
                Bold = level >= 1 && level <= 3
                FontSize = headingFontSize level }) buffer

let setHeadingLevel level doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        match findBlock doc sel.Focus.BlockId with
        | Some (Paragraph para)
        | Some (Heading (_, para)) ->
            let boundedLevel = max 1 (min 3 level)
            let styledPara = { para with Buffer = styleHeadingBuffer boundedLevel para.Buffer }
            replaceBlock doc para.Id (Heading (boundedLevel, styledPara))
        | _ -> doc

let clearHeading doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        match findBlock doc sel.Focus.BlockId with
        | Some (Heading (_, para)) -> replaceBlock doc para.Id (Paragraph para)
        | _ -> doc

/// Move cursor left
let moveCursorLeft doc =
    match doc.Selection with
    | None -> doc
    | Some sel when sel.Focus.Offset = 0 -> doc
    | Some sel ->
        let newSel = {
            Anchor = { sel.Focus with Offset = sel.Focus.Offset - 1 }
            Focus = { sel.Focus with Offset = sel.Focus.Offset - 1 }
        }
        { doc with Selection = Some newSel }

/// Move cursor right
let moveCursorRight doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        match findParagraph doc sel.Focus.BlockId with
        | None -> doc
        | Some para ->
            let maxOffset = TextBuffer.length para.Buffer
            if sel.Focus.Offset >= maxOffset then doc
            else
                let newSel = {
                    Anchor = { sel.Focus with Offset = sel.Focus.Offset + 1 }
                    Focus = { sel.Focus with Offset = sel.Focus.Offset + 1 }
                }
                { doc with Selection = Some newSel }

/// Extend selection left (Shift+ArrowLeft)
let extendSelectionLeft doc =
    match doc.Selection with
    | None -> doc
    | Some sel when sel.Focus.Offset = 0 -> doc
    | Some sel ->
        // Keep anchor, move focus left
        let newFocus = { sel.Focus with Offset = sel.Focus.Offset - 1 }
        { doc with Selection = Some { sel with Focus = newFocus } }

/// Extend selection right (Shift+ArrowRight)
let extendSelectionRight doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        match findParagraph doc sel.Focus.BlockId with
        | None -> doc
        | Some para ->
            let maxOffset = TextBuffer.length para.Buffer
            if sel.Focus.Offset >= maxOffset then doc
            else
                // Keep anchor, move focus right
                let newFocus = { sel.Focus with Offset = sel.Focus.Offset + 1 }
                { doc with Selection = Some { sel with Focus = newFocus } }

/// Get document text (for testing/debugging)
let getText doc =
    editableParagraphsInOrder doc
    |> List.map (fun para -> TextBuffer.toString para.Buffer)
    |> String.concat "\n"

let private textForSearch matchCase (text: string) =
    if matchCase then text else text.ToLowerInvariant()

let private findMatchesInText matchCase (query: string) (text: string) =
    if query = "" then
        []
    else
        let source = textForSearch matchCase text
        let needle = textForSearch matchCase query
        let mutable offset = 0
        let mutable matches = []
        let mutable searching = true

        while searching do
            let index = source.IndexOf(needle, offset)
            if index < 0 then
                searching <- false
            else
                matches <- matches @ [{ BlockId = ""; Start = index; End = index + query.Length; Text = text.Substring(index, query.Length) }]
                offset <- index + max 1 query.Length

        matches

let findText query matchCase doc =
    editableParagraphsInOrder doc
    |> List.collect (fun para ->
        let text = TextBuffer.toString para.Buffer
        findMatchesInText matchCase query text
        |> List.map (fun matchInfo -> { matchInfo with BlockId = para.Id }))

let selectFindMatch (matchInfo: FindMatch) doc =
    { doc with
        Selection = Some {
            Anchor = { BlockId = matchInfo.BlockId; Offset = matchInfo.Start }
            Focus = { BlockId = matchInfo.BlockId; Offset = matchInfo.End }
        } }

let replaceAllText query replacement matchCase doc =
    if query = "" then
        (doc, 0)
    else
        let mutable updatedDoc = doc
        let mutable count = 0

        for para in editableParagraphsInOrder doc do
            let text = TextBuffer.toString para.Buffer
            let matches = findMatchesInText matchCase query text

            if not matches.IsEmpty then
                let mutable buffer = para.Buffer
                for matchInfo in matches |> List.rev do
                    buffer <- buffer |> TextBuffer.delete matchInfo.Start (matchInfo.End - matchInfo.Start)
                    buffer <- buffer |> TextBuffer.insert matchInfo.Start replacement
                    count <- count + 1

                updatedDoc <- replaceParagraph updatedDoc para.Id { para with Buffer = buffer }

        (updatedDoc, count)

/// Set selection range
let setSelection anchor focus doc =
    { doc with Selection = Some { Anchor = anchor; Focus = focus } }

/// Select range in current block
let selectRange startOffset endOffset doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        { doc with
            Selection = Some {
                Anchor = { sel.Focus with Offset = startOffset }
                Focus = { sel.Focus with Offset = endOffset }
            } }

/// Select all text in document (all blocks)
let selectAll doc =
    match editableParagraphsInOrder doc with
    | [] -> doc
    | firstPara :: rest ->
        let lastPara =
            match rest with
            | [] -> firstPara
            | _ -> List.last rest
        let lastLen = TextBuffer.length lastPara.Buffer
        { doc with
            Selection = Some {
                Anchor = { BlockId = firstPara.Id; Offset = 0 }
                Focus = { BlockId = lastPara.Id; Offset = lastLen }
            } }

/// Apply formatting to selection
let applyFormattingToSelection formatFn doc =
    match doc.Selection with
    | None -> doc
    | Some sel when sel.IsCollapsed -> doc  // No selection
    | Some sel ->
        let (startPos, endPos, blockIds) = getNormalizedRange sel doc

        // Apply formatting to each block in selection
        let mutable updatedDoc = doc
        for (i, blockId) in (blockIds |> List.mapi (fun i id -> (i, id))) do
            match findParagraph updatedDoc blockId with
            | None -> ()
            | Some para ->
                let bufferLen = TextBuffer.length para.Buffer

                // Determine range to format within this block
                let (formatStart, formatLen) =
                    if blockIds.Length = 1 then
                        // Single block selection
                        (startPos.Offset, endPos.Offset - startPos.Offset)
                    elif i = 0 then
                        // First block: from startOffset to end
                        (startPos.Offset, bufferLen - startPos.Offset)
                    elif i = blockIds.Length - 1 then
                        // Last block: from 0 to endOffset
                        (0, endPos.Offset)
                    else
                        // Middle block: format entire block
                        (0, bufferLen)

                // Apply formatting
                if formatLen > 0 then
                    let newBuffer = formatFn formatStart formatLen para.Buffer
                    let newPara = { para with Buffer = newBuffer }
                    updatedDoc <- replaceParagraph updatedDoc para.Id newPara

        updatedDoc

/// Toggle bold on selection
let toggleBold doc =
    applyFormattingToSelection TextBuffer.toggleBold doc

/// Toggle italic on selection
let toggleItalic doc =
    applyFormattingToSelection TextBuffer.toggleItalic doc

/// Toggle underline on selection
let toggleUnderline doc =
    applyFormattingToSelection TextBuffer.toggleUnderline doc

/// Toggle all caps on selection
let toggleAllCaps doc =
    applyFormattingToSelection TextBuffer.toggleAllCaps doc

/// Toggle small caps on selection
let toggleSmallCaps doc =
    applyFormattingToSelection TextBuffer.toggleSmallCaps doc

/// Toggle strikethrough on selection
let toggleStrike doc =
    applyFormattingToSelection TextBuffer.toggleStrike doc

/// Clear direct character formatting on selection.
let clearFormatting doc =
    applyFormattingToSelection TextBuffer.clearFormatting doc

/// Set vertical alignment on selection
let setVertAlign align doc =
    applyFormattingToSelection (fun start len buf -> TextBuffer.setVertAlign start len align buf) doc

/// Set text color on selection
let setColor color doc =
    applyFormattingToSelection (fun start len buf -> TextBuffer.setColor start len color buf) doc

/// Set highlight color on selection
let setHighlight color doc =
    applyFormattingToSelection (fun start len buf -> TextBuffer.setHighlight start len color buf) doc

/// Set hyperlink on selection
let setHyperlink target doc =
    applyFormattingToSelection (fun start len buf -> TextBuffer.setHyperlink start len target buf) doc

/// Set comment metadata on selection
let setComment commentId commentText author date doc =
    applyFormattingToSelection (fun start len buf -> TextBuffer.setComment start len commentId commentText author date buf) doc

/// Set footnote/endnote metadata on selection
let setNote noteId noteType noteText doc =
    applyFormattingToSelection (fun start len buf -> TextBuffer.setNote start len noteId noteType noteText buf) doc

/// Set bookmark markers on selection
let setBookmark bookmarkName doc =
    let normalizedName =
        if System.String.IsNullOrWhiteSpace bookmarkName then "Bookmark"
        else bookmarkName.Trim().Replace(" ", "_")
    let bookmarkId = nextBookmarkId doc
    applyFormattingToSelection (fun start len buf -> TextBuffer.setBookmark start len bookmarkId normalizedName buf) doc

/// Set content control metadata on selected paragraphs
let setContentControl controlId tag alias doc =
    let normalize fallback (value: string) =
        if System.String.IsNullOrWhiteSpace value then fallback else value.Trim()
    let controlId = normalize "1" controlId
    let tag = normalize "content-control" tag
    let alias = normalize tag alias
    let blockIds =
        match doc.Selection with
        | None -> []
        | Some sel when sel.IsCollapsed -> [sel.Focus.BlockId]
        | Some sel ->
            let (_, _, ids) = getNormalizedRange sel doc
            ids
    (doc, blockIds)
    ||> List.fold (fun currentDoc blockId ->
        match findParagraph currentDoc blockId with
        | None -> currentDoc
        | Some para ->
            replaceParagraph currentDoc para.Id {
                para with
                    Metadata = {
                        para.Metadata with
                            ContentControlId = Some controlId
                            ContentControlTag = Some tag
                            ContentControlAlias = Some alias
                    }
            })

/// Set named paragraph style on selected paragraphs
let setParagraphStyle styleId doc =
    let normalizedStyle =
        if System.String.IsNullOrWhiteSpace styleId then None else Some (styleId.Trim())
    let blockIds =
        match doc.Selection with
        | None -> []
        | Some sel when sel.IsCollapsed -> [sel.Focus.BlockId]
        | Some sel ->
            let (_, _, ids) = getNormalizedRange sel doc
            ids
    (doc, blockIds)
    ||> List.fold (fun currentDoc blockId ->
        match findParagraph currentDoc blockId with
        | None -> currentDoc
        | Some para ->
            replaceParagraph currentDoc para.Id {
                para with Metadata = { para.Metadata with Style = normalizedStyle }
            })

/// Set named character style on selection
let setCharacterStyle styleId doc =
    let normalizedStyle =
        if System.String.IsNullOrWhiteSpace styleId then None else Some (styleId.Trim())
    applyFormattingToSelection (fun start len buf -> TextBuffer.setCharacterStyle start len normalizedStyle buf) doc

/// Set revision metadata on selection
let setRevision revisionType revisionId author date doc =
    let normalizedType = if revisionType = "delete" then "delete" else "insert"
    let revisionId = if System.String.IsNullOrWhiteSpace revisionId then "1" else revisionId
    let author = if System.String.IsNullOrWhiteSpace author then None else Some author
    let date = if System.String.IsNullOrWhiteSpace date then None else Some date
    applyFormattingToSelection (fun start len buf -> TextBuffer.setRevision start len normalizedType revisionId author date buf) doc

let private setRevisionWithOptions revisionType revisionId author date doc =
    applyFormattingToSelection (fun start len buf -> TextBuffer.setRevision start len revisionType revisionId author date buf) doc

let private adjacentInsertRevisionId doc (pos: Position) author date =
    findParagraph doc pos.BlockId
    |> Option.bind (fun para ->
        para.Buffer.Runs
        |> List.tryPick (function
            | TextBuffer.Text tr when tr.Start + tr.Length = pos.Offset && tr.Style.RevisionType = Some "insert" && tr.Style.RevisionAuthor = author && tr.Style.RevisionDate = date -> tr.Style.RevisionId
            | _ -> None))

/// Insert text while recording the inserted run as a revision.
let insertTrackedText author date text doc =
    let author = if System.String.IsNullOrWhiteSpace author then None else Some author
    let date = if System.String.IsNullOrWhiteSpace date then None else Some date
    let deletionRevisionId = nextRevisionId doc
    let docForInsert =
        match doc.Selection with
        | Some sel when not sel.IsCollapsed ->
            let startPos, _, _ = getNormalizedRange sel doc
            let marked = setRevisionWithOptions "delete" deletionRevisionId author date doc
            { marked with Selection = Some { Anchor = startPos; Focus = startPos } }
        | _ -> doc

    let insertionRevisionId =
        match docForInsert.Selection with
        | Some sel -> adjacentInsertRevisionId docForInsert sel.Focus author date |> Option.defaultValue (nextRevisionId docForInsert)
        | None -> nextRevisionId docForInsert
    let inserted = insertText text docForInsert
    match inserted.Selection with
    | None -> inserted
    | Some sel when text.Length > 0 ->
        let startOffset = max 0 (sel.Focus.Offset - text.Length)
        let revisionSelection = {
            Anchor = { sel.Focus with Offset = startOffset }
            Focus = sel.Focus
        }
        let styled = setRevisionWithOptions "insert" insertionRevisionId author date { inserted with Selection = Some revisionSelection }
        { styled with Selection = Some sel }
    | _ -> inserted

/// Mark the selected text as deleted instead of physically removing it.
let trackDeleteSelection author date doc =
    match doc.Selection with
    | None -> doc
    | Some sel when sel.IsCollapsed -> doc
    | Some sel ->
        let startPos, _, _ = getNormalizedRange sel doc
        let author = if System.String.IsNullOrWhiteSpace author then None else Some author
        let date = if System.String.IsNullOrWhiteSpace date then None else Some date
        let marked = setRevisionWithOptions "delete" (nextRevisionId doc) author date doc
        { marked with Selection = Some { Anchor = startPos; Focus = startPos } }

let private isDeletedRun (run: TextBuffer.TextRun) =
    run.Style.RevisionType = Some "delete"

let private nextTrackDeletableOffset para offset =
    para.Buffer.Runs
    |> List.choose (function
        | TextBuffer.Text tr when not (isDeletedRun tr) && tr.Start + tr.Length > offset ->
            Some (max offset tr.Start)
        | _ -> None)
    |> List.tryHead

let private previousTrackDeletableOffset para offset =
    para.Buffer.Runs
    |> List.choose (function
        | TextBuffer.Text tr when not (isDeletedRun tr) && tr.Start < offset ->
            Some ((min offset (tr.Start + tr.Length)) - 1)
        | _ -> None)
    |> List.filter (fun value -> value >= 0)
    |> List.tryLast

let trackDeleteForward author date doc =
    match doc.Selection with
    | None -> doc
    | Some sel when not sel.IsCollapsed -> trackDeleteSelection author date doc
    | Some sel ->
        match findParagraph doc sel.Focus.BlockId with
        | None -> doc
        | Some para ->
            match nextTrackDeletableOffset para sel.Focus.Offset with
            | None -> doc
            | Some offset ->
                let startPos = { sel.Focus with Offset = offset }
                let selection = { Anchor = startPos; Focus = { startPos with Offset = offset + 1 } }
                trackDeleteSelection author date { doc with Selection = Some selection }

let trackDeleteBackward author date doc =
    match doc.Selection with
    | None -> doc
    | Some sel when not sel.IsCollapsed -> trackDeleteSelection author date doc
    | Some sel ->
        match findParagraph doc sel.Focus.BlockId with
        | None -> doc
        | Some para ->
            match previousTrackDeletableOffset para sel.Focus.Offset with
            | None -> doc
            | Some offset ->
                let startPos = { sel.Focus with Offset = offset }
                let selection = { Anchor = startPos; Focus = { startPos with Offset = offset + 1 } }
                trackDeleteSelection author date { doc with Selection = Some selection }

type private RevisionEdit = {
    Start: int
    Length: int
    RevisionType: string
}

let private selectionRangeForBlock blockId doc =
    match doc.Selection with
    | None -> None
    | Some sel when sel.IsCollapsed ->
        if sel.Focus.BlockId = blockId then Some (sel.Focus.Offset, sel.Focus.Offset, true) else None
    | Some sel ->
        let startPos, endPos, blockIds = getNormalizedRange sel doc
        if blockIds |> List.contains blockId then
            match findParagraph doc blockId with
            | None -> None
            | Some para ->
                let bufferLen = TextBuffer.length para.Buffer
                let startOffset = if blockId = startPos.BlockId then startPos.Offset else 0
                let endOffset = if blockId = endPos.BlockId then endPos.Offset else bufferLen
                Some (startOffset, endOffset, false)
        else None

let private selectedRevisionEdits para selectionRange =
    let selStart, selEnd, isCollapsed = selectionRange
    let intersects start endPos selStart selEnd isCollapsed =
        if isCollapsed then selStart >= start && selStart <= endPos
        else start < selEnd && endPos > selStart

    para.Buffer.Runs
    |> List.choose (function
        | Text tr ->
            match tr.Style.RevisionType with
            | Some revisionType ->
                let runEnd = tr.Start + tr.Length
                if intersects tr.Start runEnd selStart selEnd isCollapsed then
                    Some { Start = tr.Start; Length = tr.Length; RevisionType = revisionType }
                else None
            | None -> None
        | Image _ -> None)

let revisionInfos doc =
    editableParagraphsInOrder doc
    |> List.collect (fun para ->
        para.Buffer.Runs
        |> List.choose (function
            | Text tr ->
                tr.Style.RevisionType
                |> Option.map (fun revisionType -> {
                    BlockId = para.Id
                    Start = tr.Start
                    End = tr.Start + tr.Length
                    RevisionType = revisionType
                    Author = tr.Style.RevisionAuthor
                    Date = tr.Style.RevisionDate
                    Text = TextBuffer.slice tr.Start tr.Length para.Buffer
                })
            | Image _ -> None))

let private applyRevisionDecision accept doc =
    let selectedBlockIds =
        match doc.Selection with
        | None -> []
        | Some sel when sel.IsCollapsed -> [sel.Focus.BlockId]
        | Some sel ->
            let _, _, blockIds = getNormalizedRange sel doc
            blockIds

    let mutable updatedDoc = doc
    let mutable changed = false
    let mutable firstSelection: Position option = None

    for blockId in selectedBlockIds do
        match findParagraph updatedDoc blockId, selectionRangeForBlock blockId updatedDoc with
        | Some para, Some selectionRange ->
            let edits = selectedRevisionEdits para selectionRange
            if not edits.IsEmpty then
                let mutable buffer = para.Buffer
                for edit in edits |> List.sortByDescending (fun edit -> edit.Start) do
                    let shouldDelete =
                        match accept, edit.RevisionType with
                        | true, "delete" -> true
                        | false, "insert" -> true
                        | _ -> false
                    if shouldDelete then
                        buffer <- TextBuffer.delete edit.Start edit.Length buffer
                    else
                        buffer <- TextBuffer.clearRevision edit.Start edit.Length buffer
                let newPara = { para with Buffer = buffer }
                updatedDoc <- replaceParagraph updatedDoc para.Id newPara
                changed <- true
                match firstSelection with
                | None -> firstSelection <- Some { BlockId = para.Id; Offset = edits |> List.map _.Start |> List.min }
                | Some _ -> ()
        | _ -> ()

    if changed then
        let selection = firstSelection |> Option.map (fun pos -> { Anchor = pos; Focus = pos })
        { updatedDoc with Selection = selection |> Option.orElse updatedDoc.Selection }
    else updatedDoc

/// Accept selected revision runs. Insertions become normal text; deletions are removed.
let acceptSelectedRevision doc =
    applyRevisionDecision true doc

/// Reject selected revision runs. Insertions are removed; deletions become normal text.
let rejectSelectedRevision doc =
    applyRevisionDecision false doc

let private allRevisionSelection doc =
    match revisionInfos doc with
    | [] -> None
    | revisions ->
        let first = revisions |> List.head
        let last = revisions |> List.last
        Some {
            Anchor = { BlockId = first.BlockId; Offset = first.Start }
            Focus = { BlockId = last.BlockId; Offset = last.End }
        }

/// Accept all revision runs in document order.
let acceptAllRevisions doc =
    match allRevisionSelection doc with
    | None -> doc
    | Some selection -> applyRevisionDecision true { doc with Selection = Some selection }

/// Reject all revision runs in document order.
let rejectAllRevisions doc =
    match allRevisionSelection doc with
    | None -> doc
    | Some selection -> applyRevisionDecision false { doc with Selection = Some selection }

let private revisionSortKey doc revision =
    let paragraphs = editableParagraphsInOrder doc
    let blockIndex = paragraphs |> List.tryFindIndex (fun para -> para.Id = revision.BlockId) |> Option.defaultValue System.Int32.MaxValue
    blockIndex, revision.Start

let private currentRevisionCursor useSelectionStart doc =
    let paragraphs = editableParagraphsInOrder doc
    match doc.Selection with
    | Some sel ->
        let pos =
            if sel.Anchor.BlockId = sel.Focus.BlockId then
                let offset = if useSelectionStart then min sel.Anchor.Offset sel.Focus.Offset else max sel.Anchor.Offset sel.Focus.Offset
                { sel.Focus with Offset = offset }
            elif useSelectionStart then
                let startPos, _, _ = getNormalizedRange sel doc
                startPos
            else
                let _, endPos, _ = getNormalizedRange sel doc
                endPos
        let blockIndex = paragraphs |> List.tryFindIndex (fun para -> para.Id = pos.BlockId) |> Option.defaultValue 0
        blockIndex, pos.Offset
    | None -> 0, 0

let private selectRevision revision doc =
    { doc with
        Selection = Some {
            Anchor = { BlockId = revision.BlockId; Offset = revision.Start }
            Focus = { BlockId = revision.BlockId; Offset = revision.End }
        } }

/// Select next revision, wrapping to the first revision when needed.
let selectNextRevision doc =
    let revisions = revisionInfos doc
    match revisions with
    | [] -> doc
    | _ ->
        let currentBlock, currentOffset = currentRevisionCursor false doc
        let next =
            revisions
            |> List.tryFind (fun revision ->
                let blockIndex, start = revisionSortKey doc revision
                blockIndex > currentBlock || (blockIndex = currentBlock && start > currentOffset))
            |> Option.defaultValue (List.head revisions)
        selectRevision next doc

/// Select previous revision, wrapping to the last revision when needed.
let selectPreviousRevision doc =
    let revisions = revisionInfos doc
    match revisions with
    | [] -> doc
    | _ ->
        let currentBlock, currentOffset = currentRevisionCursor true doc
        let previous =
            revisions
            |> List.filter (fun revision ->
                let blockIndex, start = revisionSortKey doc revision
                blockIndex < currentBlock || (blockIndex = currentBlock && start < currentOffset))
            |> List.tryLast
            |> Option.defaultValue (List.last revisions)
        selectRevision previous doc

/// Set font size on selection
let setFontSize size doc =
    applyFormattingToSelection (fun start len buf -> TextBuffer.setFontSize start len size buf) doc

/// Set font family on selection
let setFontFamily family doc =
    applyFormattingToSelection (fun start len buf -> TextBuffer.setFontFamily start len family buf) doc

/// Get selected text as plain text
let getSelectionText doc =
    match doc.Selection with
    | None -> ""
    | Some sel when sel.Anchor = sel.Focus -> ""  // No selection
    | Some sel ->
        let (startPos, endPos, blockIds) = getNormalizedRange sel doc

        blockIds
        |> List.mapi (fun i blockId ->
            let para = findParagraph doc blockId
            match para with
            | None -> ""
            | Some para ->
                let bufferLen = TextBuffer.length para.Buffer

                // Determine slice range for this block
                let sliceStart = if i = 0 then startPos.Offset else 0
                let sliceEnd = if i = blockIds.Length - 1 then endPos.Offset else bufferLen
                let sliceLen = sliceEnd - sliceStart

                if sliceLen <= 0 then ""
                else TextBuffer.slice sliceStart sliceLen para.Buffer
        )
        |> String.concat "\n"  // Separate blocks with newlines

/// Toggle bullet list on current block
let toggleBulletList doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        match findBlock doc sel.Focus.BlockId with
        | None -> doc
        | Some block ->
            let para = getParagraph block
            let newMetadata =
                if para.Metadata.ListStyle = Bullet then
                    { para.Metadata with ListStyle = NoList; ListLevel = 0 }
                else
                    { para.Metadata with ListStyle = Bullet; ListLevel = max 1 para.Metadata.ListLevel }
            let newPara = { para with Metadata = newMetadata }
            let newBlock = match block with
                           | Paragraph _ -> Paragraph newPara
                           | Heading (level, _) -> Heading (level, newPara)
                           | Table t -> Table t
                           | TextBox t -> TextBox t
            replaceBlock doc para.Id newBlock

/// Toggle numbered list on current block
let toggleNumberedList doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        match findBlock doc sel.Focus.BlockId with
        | None -> doc
        | Some block ->
            let para = getParagraph block
            let newMetadata =
                if para.Metadata.ListStyle = Numbered then
                    { para.Metadata with ListStyle = NoList; ListLevel = 0 }
                else
                    { para.Metadata with ListStyle = Numbered; ListLevel = max 1 para.Metadata.ListLevel }
            let newPara = { para with Metadata = newMetadata }
            let newBlock = match block with
                           | Paragraph _ -> Paragraph newPara
                           | Heading (level, _) -> Heading (level, newPara)
                           | Table t -> Table t
                           | TextBox t -> TextBox t
            replaceBlock doc para.Id newBlock

/// Apply a numbering marker format to the current block and ensure it is a numbered list.
let setNumberFormat format doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        match findBlock doc sel.Focus.BlockId with
        | None -> doc
        | Some block ->
            let para = getParagraph block
            let normalizedFormat =
                match format with
                | "lowerLetter" | "upperLetter" | "lowerRoman" | "upperRoman" -> format
                | _ -> "decimal"
            let newMetadata = {
                para.Metadata with
                    ListStyle = Numbered
                    ListLevel = max 1 para.Metadata.ListLevel
                    ListNumberFormat = Some normalizedFormat
                    ListLevelText = None
            }
            let newPara = { para with Metadata = newMetadata }
            let newBlock = match block with
                           | Paragraph _ -> Paragraph newPara
                           | Heading (level, _) -> Heading (level, newPara)
                           | Table t -> Table t
                           | TextBox t -> TextBox t
            replaceBlock doc para.Id newBlock

/// Increase list indentation (Tab)
let increaseIndent doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        match findBlock doc sel.Focus.BlockId with
        | None -> doc
        | Some block ->
            let para = getParagraph block
            // Only increase if in a list and not at max level
            if para.Metadata.ListStyle <> NoList && para.Metadata.ListLevel < 9 then
                let newMetadata = { para.Metadata with ListLevel = para.Metadata.ListLevel + 1 }
                let newPara = { para with Metadata = newMetadata }
                let newBlock = match block with
                               | Paragraph _ -> Paragraph newPara
                               | Heading (level, _) -> Heading (level, newPara)
                               | Table t -> Table t
                               | TextBox t -> TextBox t
                replaceBlock doc para.Id newBlock
            else
                doc

/// Decrease list indentation (Shift+Tab)
let decreaseIndent doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        match findBlock doc sel.Focus.BlockId with
        | None -> doc
        | Some block ->
            let para = getParagraph block
            // Only decrease if in a list
            if para.Metadata.ListStyle <> NoList && para.Metadata.ListLevel > 0 then
                let newLevel = para.Metadata.ListLevel - 1
                let newMetadata =
                    if newLevel = 0 then
                        // Exit list when reaching level 0
                        { para.Metadata with ListStyle = NoList; ListLevel = 0 }
                    else
                        { para.Metadata with ListLevel = newLevel }
                let newPara = { para with Metadata = newMetadata }
                let newBlock = match block with
                               | Paragraph _ -> Paragraph newPara
                               | Heading (level, _) -> Heading (level, newPara)
                               | Table t -> Table t
                               | TextBox t -> TextBox t
                replaceBlock doc para.Id newBlock
            else
                doc


/// Set alignment to left
let alignLeft doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        match findBlock doc sel.Focus.BlockId with
        | None -> doc
        | Some block ->
            let para = getParagraph block
            let newMetadata = { para.Metadata with Alignment = Left }
            let newPara = { para with Metadata = newMetadata }
            let newBlock = match block with
                           | Paragraph _ -> Paragraph newPara
                           | Heading (level, _) -> Heading (level, newPara)
                           | Table t -> Table t
                           | TextBox t -> TextBox t
            replaceBlock doc para.Id newBlock

/// Set alignment to center
let alignCenter doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        match findBlock doc sel.Focus.BlockId with
        | None -> doc
        | Some block ->
            let para = getParagraph block
            let newMetadata = { para.Metadata with Alignment = Center }
            let newPara = { para with Metadata = newMetadata }
            let newBlock = match block with
                           | Paragraph _ -> Paragraph newPara
                           | Heading (level, _) -> Heading (level, newPara)
                           | Table t -> Table t
                           | TextBox t -> TextBox t
            replaceBlock doc para.Id newBlock

/// Set alignment to right
let alignRight doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        match findBlock doc sel.Focus.BlockId with
        | None -> doc
        | Some block ->
            let para = getParagraph block
            let newMetadata = { para.Metadata with Alignment = Right }
            let newPara = { para with Metadata = newMetadata }
            let newBlock = match block with
                           | Paragraph _ -> Paragraph newPara
                           | Heading (level, _) -> Heading (level, newPara)
                           | Table t -> Table t
                           | TextBox t -> TextBox t
            replaceBlock doc para.Id newBlock

/// Set alignment to justify
let alignJustify doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        match findBlock doc sel.Focus.BlockId with
        | None -> doc
        | Some block ->
            let para = getParagraph block
            let newMetadata = { para.Metadata with Alignment = Justify }
            let newPara = { para with Metadata = newMetadata }
            let newBlock = match block with
                           | Paragraph _ -> Paragraph newPara
                           | Heading (level, _) -> Heading (level, newPara)
                           | Table t -> Table t
                           | TextBox t -> TextBox t
            replaceBlock doc para.Id newBlock

let setCurrentParagraphStyle alignment indentLeft indentRight spacingBefore spacingAfter lineSpacing firstLineIndent backgroundColor doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        match findBlock doc sel.Focus.BlockId with
        | None -> doc
        | Some block ->
            let para = getParagraph block
            let newMetadata = {
                para.Metadata with
                    Alignment = alignment
                    IndentLeft = indentLeft
                    IndentRight = indentRight
                    SpacingBefore = spacingBefore
                    SpacingAfter = spacingAfter
                    LineSpacing = lineSpacing
                    LineSpacingRule = Some "auto"
                    FirstLineIndent = firstLineIndent
                    BackgroundColor = backgroundColor
            }
            let newPara = { para with Metadata = newMetadata }
            let newBlock =
                match block with
                | Paragraph _ -> Paragraph newPara
                | Heading (level, _) -> Heading (level, newPara)
                | Table table -> Table table
                | TextBox textBox -> TextBox textBox
            replaceBlock doc para.Id newBlock

let setCurrentSectionBreak sectionBreak doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        match findBlock doc sel.Focus.BlockId with
        | None -> doc
        | Some block ->
            let para = getParagraph block
            let newMetadata = { para.Metadata with SectionBreak = sectionBreak }
            let newPara = { para with Metadata = newMetadata }
            let newBlock =
                match block with
                | Paragraph _ -> Paragraph newPara
                | Heading (level, _) -> Heading (level, newPara)
                | Table table -> Table table
                | TextBox textBox -> TextBox textBox
            replaceBlock doc para.Id newBlock

let setCurrentHeaderFooterReferences headerReferences footerReferences doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        match findBlock doc sel.Focus.BlockId with
        | None -> doc
        | Some block ->
            let para = getParagraph block
            let newMetadata = { para.Metadata with HeaderReferences = headerReferences; FooterReferences = footerReferences }
            let newPara = { para with Metadata = newMetadata }
            let newBlock =
                match block with
                | Paragraph _ -> Paragraph newPara
                | Heading (level, _) -> Heading (level, newPara)
                | Table table -> Table table
                | TextBox textBox -> TextBox textBox
            replaceBlock doc para.Id newBlock
