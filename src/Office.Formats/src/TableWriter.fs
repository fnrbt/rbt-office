module TableWriter

open XmlGenerator
open DocumentProcessors.EditorProcessor

/// Table border style
type BorderStyle =
    | NoBorder
    | Single
    | Double
    | Dashed
    | Dotted

/// Table border
type TableBorder = {
    Style: BorderStyle
    Width: int  // in eighths of a point
    Color: string option
}

/// Cell vertical alignment
type VerticalAlignment =
    | Top
    | Center
    | Bottom

/// Table cell properties
type CellProperties = {
    Width: int option  // in twips (1/1440 inch)
    BackgroundColor: string option
    VerticalAlignment: VerticalAlignment option
    Borders: Map<string, TableBorder> option  // top, bottom, left, right
}

/// Table cell
type TableCell = {
    Paragraphs: EditorParagraph list
    Properties: CellProperties option
}

/// Table row
type TableRow = {
    Cells: TableCell list
    Height: int option  // in twips
}

/// Table properties
type TableProperties = {
    Width: int option  // in twips
    Borders: Map<string, TableBorder> option
    CellSpacing: int option
    Alignment: string option
}

/// Table definition
type Table = {
    Rows: TableRow list
    Properties: TableProperties option
}

/// Convert border style to Word format
let borderStyleToWord (style: BorderStyle) : string =
    match style with
    | NoBorder -> "none"
    | Single -> "single"
    | Double -> "double"
    | Dashed -> "dashed"
    | Dotted -> "dotted"

/// Generate border XML
let generateBorder (name: string) (border: TableBorder) : string =
    let style = borderStyleToWord border.Style
    let width = string border.Width
    let color = defaultArg border.Color "auto"

    selfClosing (sprintf "w:%s" name) [
        ("w:val", style)
        ("w:sz", width)
        ("w:color", color)
    ]

/// Generate table borders XML
let generateTableBorders (borders: Map<string, TableBorder>) : string =
    let borderElements =
        borders
        |> Map.toList
        |> List.map (fun (name, border) -> generateBorder name border)
        |> String.concat ""

    element "w:tblBorders" [] [borderElements]

/// Generate table properties XML
let generateTableProperties (props: TableProperties option) : string =
    match props with
    | None -> ""
    | Some p ->
        let mutable children = []

        match p.Width with
        | Some w -> children <- (selfClosing "w:tblW" [("w:w", string w); ("w:type", "dxa")]) :: children
        | None -> ()

        match p.Borders with
        | Some borders -> children <- (generateTableBorders borders) :: children
        | None -> ()

        match p.CellSpacing with
        | Some spacing -> children <- (selfClosing "w:tblCellSpacing" [("w:w", string spacing); ("w:type", "dxa")]) :: children
        | None -> ()

        match p.Alignment with
        | Some align -> children <- (selfClosing "w:jc" [("w:val", align)]) :: children
        | None -> ()

        if List.isEmpty children then
            ""
        else
            element "w:tblPr" [] (List.rev children)

/// Generate vertical alignment
let verticalAlignmentToWord (valign: VerticalAlignment) : string =
    match valign with
    | Top -> "top"
    | Center -> "center"
    | Bottom -> "bottom"

/// Generate cell properties XML
let generateCellProperties (props: CellProperties option) : string =
    match props with
    | None -> ""
    | Some p ->
        let mutable children = []

        match p.Width with
        | Some w -> children <- (selfClosing "w:tcW" [("w:w", string w); ("w:type", "dxa")]) :: children
        | None -> ()

        match p.BackgroundColor with
        | Some color -> children <- (selfClosing "w:shd" [("w:val", "clear"); ("w:color", "auto"); ("w:fill", color)]) :: children
        | None -> ()

        match p.VerticalAlignment with
        | Some valign -> children <- (selfClosing "w:vAlign" [("w:val", verticalAlignmentToWord valign)]) :: children
        | None -> ()

        match p.Borders with
        | Some borders ->
            let borderElements =
                borders
                |> Map.toList
                |> List.map (fun (name, border) -> generateBorder name border)
                |> String.concat ""
            children <- (element "w:tcBorders" [] [borderElements]) :: children
        | None -> ()

        if List.isEmpty children then
            ""
        else
            element "w:tcPr" [] (List.rev children)

/// Adapt the earlier table-builder API to the same typed paragraph writer used
/// by editor package saves. This is a public builder edge, not a second XML writer.
let private editorRun (run: EditorRun) : WordEditorPackage.EditorRun =
    let content =
        if run.IsImage then
            WordEditorPackage.Picture {
                Source = run.ImageSrc; Linked = run.ImageLinked
                Width = run.ImageWidth; Height = run.ImageHeight
                Alt = run.ImageAlt; Name = run.ImageName; Title = run.ImageTitle
                Floating = run.ImageFloating; Wrap = run.ImageWrap
                PositionX = run.ImagePositionX; PositionY = run.ImagePositionY
                PositionHRelativeFrom = run.ImagePositionHRelativeFrom
                PositionVRelativeFrom = run.ImagePositionVRelativeFrom
                PositionHAlign = None; PositionVAlign = None
                BehindDoc = run.ImageBehindDoc; AllowOverlap = run.ImageAllowOverlap
                LayoutInCell = run.ImageLayoutInCell
                DistanceTop = run.ImageDistanceTop; DistanceBottom = run.ImageDistanceBottom
                DistanceLeft = run.ImageDistanceLeft; DistanceRight = run.ImageDistanceRight
                RelativeHeight = run.ImageRelativeHeight; Locked = run.ImageLocked
                SimplePosition = run.ImageSimplePosition
                SimplePositionX = run.ImageSimplePositionX; SimplePositionY = run.ImageSimplePositionY
                CropLeft = run.ImageCropLeft; CropTop = run.ImageCropTop
                CropRight = run.ImageCropRight; CropBottom = run.ImageCropBottom
                WrapText = None; WrapPolygon = []
                FlipH = run.ImageFlipH; FlipV = run.ImageFlipV; Opaque = None }
        else
            match run.EquationXml with
            | Some xml -> WordEditorPackage.Equation (run.Text, xml)
            | None -> WordEditorPackage.Text run.Text
    { Content = content
      Formatting = {
        Style = None; StyleInfo = None
        Bold = run.Bold; Italic = run.Italic; Underline = run.Underline
        UnderlineStyle = None; UnderlineColor = None
        AllCaps = false; SmallCaps = false; Strike = run.Strike; DoubleStrike = false
        Hidden = false; Position = 0.0; EmphasisMark = None
        FontSize = run.FontSize; CharacterSpacing = run.CharacterSpacing
        FontFamily = run.FontFamily; FontFamilyInherited = None
        Color = run.Color; Highlight = run.Highlight; VertAlign = run.VertAlign
        HasExplicit = false; ExplicitBold = None; ExplicitItalic = None
        ExplicitUnderline = None; ExplicitAllCaps = None; ExplicitSmallCaps = None
        ExplicitStrike = None; ExplicitDoubleStrike = None; ExplicitHidden = None
        ExplicitFontSize = None; ExplicitColorAuto = false }
      Hyperlink = run.Hyperlink; InactiveHyperlink = None
      CommentId = run.CommentId; CommentText = run.CommentText
      CommentAuthor = run.CommentAuthor; CommentDate = run.CommentDate
      NoteId = run.NoteId; NoteType = run.NoteType; NoteText = run.NoteText
      BookmarkStarts = run.BookmarkStarts; BookmarkEnds = run.BookmarkEnds
      FieldCode = run.FieldCode; FieldResult = run.FieldResult
      Revision = None; RenderedPageBreakBefore = run.RenderedPageBreakBefore
      Checkbox = None; ContentControls = []; SourceRun = None; NoProps = false }

let private editorParagraph (para: EditorParagraph) : WordEditorPackage.EditorParagraph =
    { Runs = para.Runs |> List.map editorRun |> List.toArray
      HeadingLevel = None
      Formatting = {
        Style = None; StyleInfo = None; Alignment = para.Alignment
        HasExplicit = false; ExplicitAlignment = None
        IndentLeft = para.IndentLeft; IndentRight = para.IndentRight
        FirstLineIndent = para.FirstLineIndent
        SpacingBefore = para.SpacingBefore; SpacingAfter = para.SpacingAfter
        LineSpacing = para.LineSpacing; LineSpacingRule = None
        RightTabPosition = para.RightTabPosition; RightTabLeader = para.RightTabLeader
        TabStops = []; Rtl = false; ContextualSpacing = None
        KeepNext = None; KeepLines = None; WidowControl = None
        PageBreakBefore = false; PageBreakBeforeDerived = false
        PageBreakAtStart = false; PageBreakAfter = false; ColumnBreakAfter = false
        RenderedPageBreakBefore = None; BackgroundColor = None
        ShadingPattern = None; ShadingForeground = None; Borders = None
        EffectiveBorders = None
        OutlineLevel = None; ExplicitParagraphFormat = false
        ExplicitIndentLeft = None; ExplicitIndentRight = None; ExplicitFirstLineIndent = None
        ExplicitSpacingBefore = None; ExplicitSpacingAfter = None; ExplicitLineSpacing = None
        ExplicitContextualSpacing = None; ExplicitBackgroundColor = None
        ExplicitShadingPattern = None; ExplicitShadingForeground = None; ExplicitTabs = false
        HasMarkProperties = false; ExplicitMarkFontSize = None; ParagraphMarkFontSize = None
        MarkRevision = None; DropCapKind = None; DropCapLines = None; DropCapDistance = None
        ContentControlId = None; ContentControlTag = None; ContentControlAlias = None
        ContentControlXml = None; ContentControlGroup = None; ContentControlAncestors = []; EnclosingCellControls = [] }
      List = None; SectionBreak = None; SectionPageSetup = None
      SectionColumns = None; SectionVerticalAlign = None
      SectionPageNumberStart = None; SectionPageNumberFormat = None
      HeaderReferences = []; FooterReferences = []
      Source = { Origin = None; Xml = None; Copy = false; Same = false }
      SourceParagraph = None; SpliceOriginal = None; SpliceFrom = None
      SpliceLeading = true; SpliceTrailing = true; SpliceJoined = [||]
      SourceIndex = None; SourceUnchanged = false; SourceWhy = None; SourceTextChanges = None }

let generateCellParagraph (para: EditorParagraph) : string =
    let typed = editorParagraph para
    let context = WordPackageWriterModel.createWriterContext ()
    let imageScope, _ = WordPackageWriterModel.buildImageScope context 1 [| WordEditorPackage.Paragraph typed |]
    WordPackageWriterModel.generateParagraphWithHyperlinksAndImagesWithPageSetup
        context WordPackageWriterModel.defaultPageSetupInfo
        (WordPackageWriterModel.hyperlinkMap context [| WordEditorPackage.Paragraph typed |])
        imageScope typed

/// Generate cell XML
let generateCell (cell: TableCell) : string =
    let tcPr = generateCellProperties cell.Properties
    let paragraphs =
        if List.isEmpty cell.Paragraphs then
            // Empty cell needs at least one empty paragraph
            [element "w:p" [] []]
        else
            cell.Paragraphs |> List.map generateCellParagraph

    let content = String.concat "" paragraphs

    element "w:tc" [] [tcPr; content]

/// Generate row properties XML
let generateRowProperties (height: int option) : string =
    match height with
    | None -> ""
    | Some h ->
        element "w:trPr" [] [
            selfClosing "w:trHeight" [("w:val", string h); ("w:hRule", "atLeast")]
        ]

/// Generate row XML
let generateRow (row: TableRow) : string =
    let trPr = generateRowProperties row.Height
    let cells =
        row.Cells
        |> List.map generateCell
        |> String.concat ""

    element "w:tr" [] [trPr; cells]

/// Generate table XML
let generateTable (table: Table) : string =
    let tblPr = generateTableProperties table.Properties
    let rows =
        table.Rows
        |> List.map generateRow
        |> String.concat ""

    element "w:tbl" [] [tblPr; rows]

/// Create a simple table from 2D array of text
let createSimpleTable (data: string list list) : Table =
    let rows =
        data
        |> List.map (fun rowData ->
            let cells =
                rowData
                |> List.map (fun cellText ->
                    let para = {
                        Runs = [{
                            Text = cellText
                            Bold = false
                            Italic = false
                            Underline = false
                            Strike = false
                            FontSize = 11.0
                            CharacterSpacing = 0.0
                            FontFamily = "Calibri"
                            Color = None
                            Highlight = None
                            VertAlign = None
                            CommentId = None
                            CommentText = None
                            CommentAuthor = None
                            CommentDate = None
                            NoteId = None
                            NoteType = None
                            NoteText = None
                            Hyperlink = None
                            BookmarkStarts = []
                            BookmarkEnds = []
                            FieldCode = None
                            FieldResult = None
                            EquationXml = None
                            RenderedPageBreakBefore = false
                            IsImage = false
                            ImageSrc = None
                            ImageLinked = false
                            ImageWidth = None
                            ImageHeight = None
                            ImageAlt = None
                            ImageName = None
                            ImageTitle = None
                            ImageFloating = false
                            ImageWrap = None
                            ImagePositionX = None
                            ImagePositionY = None
                            ImagePositionHRelativeFrom = None
                            ImagePositionVRelativeFrom = None
                            ImageBehindDoc = None
                            ImageAllowOverlap = None
                            ImageLayoutInCell = None
                            ImageDistanceTop = None
                            ImageDistanceBottom = None
                            ImageDistanceLeft = None
                            ImageDistanceRight = None
                            ImageRelativeHeight = None
                            ImageLocked = None
                            ImageSimplePosition = None
                            ImageSimplePositionX = None
                            ImageSimplePositionY = None
                            ImageCropLeft = None
                            ImageCropTop = None
                            ImageCropRight = None
                            ImageCropBottom = None
                            ImageFlipH = false
                            ImageFlipV = false
                        }]
                        Alignment = "left"
                        IndentLeft = None
                        IndentRight = None
                        FirstLineIndent = None
                        RightTabPosition = None
                        RightTabLeader = None
                        SpacingBefore = None
                        SpacingAfter = None
                        LineSpacing = None
                    }
                    { Paragraphs = [para]; Properties = None }
                )
            { Cells = cells; Height = None }
        )

    { Rows = rows; Properties = None }

/// Create default table border
let defaultBorder () : TableBorder = {
    Style = Single
    Width = 4  // Half point
    Color = Some "000000"
}

/// Create table with borders
let createTableWithBorders (data: string list list) : Table =
    let table = createSimpleTable data

    let borders = Map.ofList [
        ("top", defaultBorder())
        ("bottom", defaultBorder())
        ("left", defaultBorder())
        ("right", defaultBorder())
        ("insideH", defaultBorder())
        ("insideV", defaultBorder())
    ]

    { table with Properties = Some { Width = None; Borders = Some borders; CellSpacing = None; Alignment = Some "left" } }
