module NativeRunner

open System
open System.IO
open Renderer
open RedBat.Word
open RedBat.Word.Document
open RedBat.Word.TextBuffer
open RedBat.Word.Rope

module EP = DocumentProcessors.EditorProcessor
module TC = DocumentProcessors.TypeConversion

let private usage () =
    eprintfn "Usage: dotnet run --project tools/native-runner -- <input.docx> <output.png> [width] [height]"
    eprintfn "Defaults: width=900 height=1600"

let private fullPath (path: string) =
    if Path.IsPathRooted path then path
    else Path.GetFullPath(path, Directory.GetCurrentDirectory())

let private existingFile path =
    if File.Exists path then Some path else None

let private findWordFont () =
    [
        Environment.GetEnvironmentVariable("RB_WORD_FONT_DIR")
        Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "artifacts/fonts/windows"))
        Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "../../artifacts/fonts/windows"))
    ]
    |> List.choose (fun dir ->
        if String.IsNullOrWhiteSpace dir then None
        else Path.Combine(dir, "calibri.ttf") |> existingFile)
    |> List.tryHead

let private initWordFont () =
    findWordFont ()
    |> Option.iter Composit.Ui.Layout.initFont

let private parseFloat (value: string option) =
    value
    |> Option.bind (fun text ->
        match Double.TryParse(text, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture) with
        | true, parsed -> Some parsed
        | false, _ -> None)

let private twipsToCanvasPixels value = parseFloat value |> Option.map (fun v -> v / 15.0)

let private twipsToPoints value = parseFloat value |> Option.map (fun v -> v / 20.0)

let private intOpt (value: string option) =
    value
    |> Option.bind (fun text ->
        match Int32.TryParse(text) with
        | true, parsed -> Some parsed
        | false, _ -> None)

let private color value = TC.normalizeColor value

let private alignment = function
    | "center" -> Center
    | "right" -> Right
    | "justify"
    | "both"
    | "distribute" -> Justify
    | _ -> Left

let private listStyle = function
    | Some "bullet" -> Bullet
    | Some "decimal"
    | Some "lowerLetter"
    | Some "upperLetter"
    | Some "lowerRoman"
    | Some "upperRoman" -> Numbered
    | _ -> NoList

let private border (value: DocxReader.BorderStyle option) : TableBorder option =
    value |> Option.map (fun b -> {
        Style = b.BorderType
        Color = color b.Color
        Width = parseFloat b.Size |> Option.map (fun size -> size / 8.0)
    })

let private tableBorders (value: DocxReader.CellBorders option) : TableBorders option =
    value |> Option.map (fun b -> {
        Top = border b.Top
        Right = border b.Right
        Bottom = border b.Bottom
        Left = border b.Left
        InsideH = border b.InsideH
        InsideV = border b.InsideV
    })

let private mergeCellBorders (baseBorders: DocxReader.CellBorders option) (borders: DocxReader.CellBorders option) =
    match baseBorders, borders with
    | None, value -> value
    | value, None -> value
    | Some baseValue, Some value ->
        Some {
            Top = value.Top |> Option.orElse baseValue.Top
            Bottom = value.Bottom |> Option.orElse baseValue.Bottom
            Left = value.Left |> Option.orElse baseValue.Left
            Right = value.Right |> Option.orElse baseValue.Right
            InsideH = value.InsideH |> Option.orElse baseValue.InsideH
            InsideV = value.InsideV |> Option.orElse baseValue.InsideV
        }

let private mergeTableStyleProperties (baseProps: DocxReader.TableStyleProperties) (props: DocxReader.TableStyleProperties) : DocxReader.TableStyleProperties = {
    Borders = mergeCellBorders baseProps.Borders props.Borders
    Shading = props.Shading |> Option.orElse baseProps.Shading
    PaddingTop = props.PaddingTop |> Option.orElse baseProps.PaddingTop
    PaddingRight = props.PaddingRight |> Option.orElse baseProps.PaddingRight
    PaddingBottom = props.PaddingBottom |> Option.orElse baseProps.PaddingBottom
    PaddingLeft = props.PaddingLeft |> Option.orElse baseProps.PaddingLeft
}

let private resolveTableStyleProperties (styleMap: Map<string, DocxReader.StyleDefinition>) styleId =
    let rec resolve visited id =
        if Set.contains id visited then None
        else
            match Map.tryFind id styleMap with
            | None -> None
            | Some style ->
                let baseProps =
                    match style.BasedOn with
                    | Some basedOn when basedOn <> id -> resolve (Set.add id visited) basedOn
                    | _ -> None
                match baseProps, style.TableProperties with
                | Some baseValue, Some value -> Some (mergeTableStyleProperties baseValue value)
                | None, Some value -> Some value
                | Some value, None -> Some value
                | None, None -> None

    styleId |> Option.bind (resolve Set.empty)

let private paraBorder (value: DocxReader.ParagraphBorder option) : RedBat.Word.Document.ParagraphBorder option =
    value |> Option.map (fun b -> {
        Style = b.Style
        Color = color b.Color
        Width = parseFloat b.Width |> Option.map (fun size -> size / 8.0)
        Space = parseFloat b.Space
    })

let private paraBorders (value: DocxReader.ParagraphBorders option) : RedBat.Word.Document.ParagraphBorders option =
    value |> Option.map (fun b -> {
        Top = paraBorder b.Top
        Right = paraBorder b.Right
        Bottom = paraBorder b.Bottom
        Left = paraBorder b.Left
    })

type private NumberingResolver = {
    ById: Map<string * int, DocxReader.NumberingDefinition>
    ByStyle: Map<string, DocxReader.NumberingDefinition>
}

let private numberingResolver (definitions: DocxReader.NumberingDefinition list) =
    {
        ById = definitions |> List.map (fun def -> (def.NumId, def.Level), def) |> Map.ofList
        ByStyle =
            definitions
            |> List.choose (fun def -> def.ParagraphStyle |> Option.map (fun styleId -> styleId, def))
            |> Map.ofList
    }

let private resolvedNumberingDefinition resolver (props: DocxReader.ParagraphProperties) =
    let fromDirectList =
        props.ListInfo
        |> Option.bind (fun info -> resolver.ById |> Map.tryFind (info.ListId, info.Level))

    fromDirectList
    |> Option.orElseWith (fun () ->
        props.Style
        |> Option.bind (fun styleId -> resolver.ByStyle |> Map.tryFind styleId))

let private tableWidth width widthType =
    match widthType, parseFloat width with
    | Some "auto", _
    | Some "nil", _ -> Some { Kind = AutoWidth; Value = None }
    | Some "pct", Some value -> Some { Kind = PercentWidth; Value = Some (value / 50.0) }
    | Some "dxa", Some value -> Some { Kind = DxaWidth; Value = Some value }
    | _, Some value -> Some { Kind = PointWidth; Value = Some value }
    | _ -> None

let private lineSpacingValue value rule =
    match rule with
    | Some "exact"
    | Some "atLeast" -> twipsToPoints value
    | _ -> TC.lineSpacingToMultiple value

let private textStyle (run: EP.EditorRun) : RedBat.Word.TextBuffer.TextStyle =
    {
        defaultStyle with
            Bold = run.Bold
            Italic = run.Italic
            Underline = run.Underline
            Strike = run.Strike
            FontSize = run.FontSize
            FontFamily = Some run.FontFamily
            Color = run.Color
            Highlight = run.Highlight
            CommentId = run.CommentId
            CommentText = run.CommentText
            CommentAuthor = run.CommentAuthor
            CommentDate = run.CommentDate
            NoteId = run.NoteId
            NoteType = run.NoteType
            NoteText = run.NoteText
            VertAlign = run.VertAlign
    }

let private editorRunsToBuffer (images: Map<string, string>) (runs: EP.EditorRun list) =
    let mutable offset = 0
    let parts = ResizeArray<string>()
    let bufferRuns = ResizeArray<RedBat.Word.TextBuffer.Run>()

    for run in runs do
        if run.IsImage then
            let src = run.ImageSrc |> Option.orElseWith (fun () -> Map.tryFind run.Text images) |> Option.defaultValue ""
            parts.Add("\uFFFC")
            bufferRuns.Add(Image {
                Start = offset
                Src = src
                Linked = run.ImageLinked
                Width = run.ImageWidth
                Height = run.ImageHeight
                Alt = run.ImageAlt
                Name = run.ImageName
                Title = run.ImageTitle
                Floating = run.ImageFloating
                Wrap = run.ImageWrap
                PositionX = run.ImagePositionX
                PositionY = run.ImagePositionY
                PositionHRelativeFrom = run.ImagePositionHRelativeFrom
                PositionVRelativeFrom = run.ImagePositionVRelativeFrom
                BehindDoc = run.ImageBehindDoc
                AllowOverlap = run.ImageAllowOverlap
                LayoutInCell = run.ImageLayoutInCell
                DistanceTop = run.ImageDistanceTop
                DistanceBottom = run.ImageDistanceBottom
                DistanceLeft = run.ImageDistanceLeft
                DistanceRight = run.ImageDistanceRight
                RelativeHeight = run.ImageRelativeHeight
                Locked = run.ImageLocked
                SimplePosition = run.ImageSimplePosition
                SimplePositionX = run.ImageSimplePositionX
                SimplePositionY = run.ImageSimplePositionY
                CropLeft = run.ImageCropLeft
                CropTop = run.ImageCropTop
                CropRight = run.ImageCropRight
                CropBottom = run.ImageCropBottom
            })
            offset <- offset + 1
        elif not (String.IsNullOrEmpty run.Text) then
            parts.Add(run.Text)
            bufferRuns.Add(Text { Start = offset; Length = run.Text.Length; Style = textStyle run })
            offset <- offset + run.Text.Length

    { Rope = Rope.ofString (String.Concat parts); Runs = bufferRuns |> Seq.toList }

let private paraMetadata resolver (para: DocxReader.Paragraph) (editorPara: EP.EditorParagraph) =
    let props = para.Properties
    let numberingDef = resolvedNumberingDefinition resolver props
    let numberingIndentLeft = numberingDef |> Option.bind (fun def -> twipsToPoints def.IndentLeft)
    let numberingRightTab = numberingDef |> Option.bind (fun def -> twipsToPoints def.RightTabPosition)
    let numberingFirstLineIndent =
        numberingDef
        |> Option.bind (fun def ->
            match twipsToPoints def.FirstLineIndent, twipsToPoints def.HangingIndent with
            | Some firstLine, _ -> Some firstLine
            | None, Some hanging -> Some -hanging
            | None, None -> None)
    {
        defaultMetadata with
            Style = props.Style
            ContentControlId = props.ContentControlId
            ContentControlTag = props.ContentControlTag
            ContentControlAlias = props.ContentControlAlias
            Alignment = alignment editorPara.Alignment
            IndentLeft = editorPara.IndentLeft |> Option.orElse numberingIndentLeft
            IndentRight = editorPara.IndentRight
            FirstLineIndent =
                match twipsToPoints props.HangingIndent with
                | Some hanging -> Some -hanging
                | None -> editorPara.FirstLineIndent |> Option.orElse numberingFirstLineIndent
            RightTabPosition = twipsToPoints props.RightTabPosition |> Option.orElse numberingRightTab
            SpacingBefore = editorPara.SpacingBefore
            SpacingAfter = editorPara.SpacingAfter
            ContextualSpacing = props.ContextualSpacing
            LineSpacing = editorPara.LineSpacing
            LineSpacingRule = props.LineRule
            WidowControl = props.WidowControl
            BackgroundColor = color props.BackgroundColor
            ParagraphBorders = paraBorders props.ParagraphBorders
            PageBreakBefore = props.PageBreakBefore |> Option.defaultValue false
            SectionBreak = props.SectionBreak
            SectionPageNumberStart = props.SectionPageNumberStart
            RenderedPageBreakBefore = props.RenderedPageBreakBefore
            HeaderReferences = props.HeaderReferences |> List.map (fun ref -> ref.Kind, ref.RelationshipId)
            FooterReferences = props.FooterReferences |> List.map (fun ref -> ref.Kind, ref.RelationshipId)
            ListStyle =
                match numberingDef with
                | Some def when def.NumberFormat = "bullet" -> Bullet
                | Some _ -> Numbered
                | None -> props.ListInfo |> Option.map (fun li -> li.NumberFormat) |> Option.defaultValue None |> listStyle
            ListLevel =
                match numberingDef with
                | Some def -> def.Level + 1
                | None -> props.ListInfo |> Option.map (fun li -> li.Level + 1) |> Option.defaultValue 0
            ListId = numberingDef |> Option.map _.NumId |> Option.orElse (props.ListInfo |> Option.map _.ListId)
            ListNumberFormat = numberingDef |> Option.map _.NumberFormat |> Option.orElse (props.ListInfo |> Option.bind _.NumberFormat)
            ListLevelText = numberingDef |> Option.bind _.LevelText |> Option.orElse (props.ListInfo |> Option.bind _.LevelText)
            ListStart = numberingDef |> Option.bind _.Start |> Option.orElse (props.ListInfo |> Option.bind _.Start)
            ParagraphMarkFontSize = props.ParagraphRunProperties |> Option.map (fun rPr -> TC.fontSizeToPoints rPr.FontSize)
    }

let private mkPara resolver images styleMap comments notes para =
    let editorPara = EP.toEditorParagraph None styleMap comments notes para
    {
        Id = Guid.NewGuid().ToString("N")
        Buffer = editorRunsToBuffer images editorPara.Runs
        Metadata = paraMetadata resolver para editorPara
    }

let private tableCell resolver images styleMap comments notes (cell: DocxReader.TableCell) =
    let props = cell.Properties
    let content = cell.Paragraphs |> List.map (mkPara resolver images styleMap comments notes)
    {
        Content = if content.IsEmpty then [mkPara resolver images styleMap comments notes { Runs = []; Properties = DocxReader.Docx.paragraphProps XmlDsl.Empty }] else content
        ColSpan = props |> Option.bind (fun p -> p.ColSpan) |> Option.defaultValue 1
        BackgroundColor = props |> Option.bind (fun p -> p.Shading) |> Option.bind (fun s -> color s.Fill)
        Width = props |> Option.bind (fun p -> twipsToCanvasPixels p.Width)
        PreferredWidth = props |> Option.bind (fun p -> tableWidth p.Width p.WidthType)
        BorderColor = None
        BorderWidth = None
        Borders = props |> Option.bind (fun p -> tableBorders p.Borders)
        VerticalAlignment = props |> Option.bind (fun p -> p.VerticalAlignment)
        VerticalMerge =
            props
            |> Option.bind (fun p -> p.VerticalMerge)
            |> Option.bind (function | "restart" -> Some RestartMerge | "continue" -> Some ContinueMerge | _ -> None)
        PaddingTop = props |> Option.bind (fun p -> twipsToCanvasPixels p.PaddingTop)
        PaddingRight = props |> Option.bind (fun p -> twipsToCanvasPixels p.PaddingRight)
        PaddingBottom = props |> Option.bind (fun p -> twipsToCanvasPixels p.PaddingBottom)
        PaddingLeft = props |> Option.bind (fun p -> twipsToCanvasPixels p.PaddingLeft)
    }

let private tableBlock resolver images styleMap tableStyleMap comments notes (table: DocxReader.Table) =
    let styleProperties = resolveTableStyleProperties tableStyleMap table.StyleId
    let tablePadding selector =
        table.Properties
        |> Option.bind selector
        |> Option.orElseWith (fun () -> styleProperties |> Option.bind selector)
        |> twipsToCanvasPixels
    let tableShading =
        table.Properties
        |> Option.bind (fun p -> p.Shading)
        |> Option.orElseWith (fun () -> styleProperties |> Option.bind (fun p -> p.Shading))
    let tableBorderProperties =
        table.Borders
        |> Option.orElseWith (fun () -> table.Properties |> Option.bind (fun p -> p.Borders))
        |> Option.orElseWith (fun () -> styleProperties |> Option.bind (fun p -> p.Borders))
    Table {
        Id = Guid.NewGuid().ToString("N")
        Rows =
            table.Rows
            |> List.map (fun row -> {
                Cells = row.Cells |> List.map (tableCell resolver images styleMap comments notes)
                IsHeader = row.IsHeader
                Height = twipsToCanvasPixels row.Height
                HeightRule = row.HeightRule
                CanSplit = row.CanSplit
                CellSpacing = twipsToCanvasPixels row.CellSpacing
                PaddingTop = twipsToCanvasPixels row.PaddingTop
                PaddingRight = twipsToCanvasPixels row.PaddingRight
                PaddingBottom = twipsToCanvasPixels row.PaddingBottom
                PaddingLeft = twipsToCanvasPixels row.PaddingLeft
            })
        PreferredWidth = tableWidth table.Width table.WidthType
        LayoutMode =
            match table.LayoutMode with
            | Some "fixed" -> Some FixedTableLayout
            | Some "autofit"
            | Some "autoFit" -> Some AutoFitTableLayout
            | _ -> None
        GridColumns = table.GridColumns |> List.choose (fun width -> tableWidth (Some width) (Some "dxa"))
        Borders = tableBorders tableBorderProperties
        BackgroundColor = tableShading |> Option.bind (fun s -> color s.Fill)
        CellSpacing = twipsToCanvasPixels table.CellSpacing
        PaddingTop = tablePadding (fun p -> p.PaddingTop)
        PaddingRight = tablePadding (fun p -> p.PaddingRight)
        PaddingBottom = tablePadding (fun p -> p.PaddingBottom)
        PaddingLeft = tablePadding (fun p -> p.PaddingLeft)
        Floating = table.Floating
    }

let private namedStyle (style: DocxReader.StyleDefinition) : NamedStyle =
    let run = style.RunProperties
    let para = style.ParagraphProperties
    {
        Id = style.Id
        Name = style.Name
        StyleType = style.StyleType
        IsDefault = style.IsDefault
        BasedOn = style.BasedOn
        Next = style.Next
        Linked = style.Linked
        IsQuickStyle = style.IsQuickStyle
        Aliases = style.Aliases
        UiPriority = intOpt style.UiPriority
        SemiHidden = style.SemiHidden
        UnhideWhenUsed = style.UnhideWhenUsed
        Text =
            run |> Option.map (fun r -> {
                defaultStyle with
                    Bold = r.Bold |> Option.defaultValue false
                    Italic = r.Italic |> Option.defaultValue false
                    Underline = r.Underline |> Option.defaultValue false
                    AllCaps = r.AllCaps |> Option.defaultValue false
                    SmallCaps = r.SmallCaps |> Option.defaultValue false
                    Strike = r.Strike |> Option.defaultValue false
                    FontSize = TC.fontSizeToPoints r.FontSize
                    FontFamily = r.FontFamily
                    Color = color r.Color
                    Highlight = r.Highlight
                    VertAlign = r.VertAlign
            })
        SpacingBefore = para |> Option.bind (fun p -> twipsToPoints p.SpacingBefore)
        SpacingAfter = para |> Option.bind (fun p -> twipsToPoints p.SpacingAfter)
        ContextualSpacing = para |> Option.bind (fun p -> p.ContextualSpacing)
        LineSpacing = para |> Option.bind (fun p -> lineSpacingValue p.LineSpacing p.LineRule)
        LineSpacingRule = para |> Option.bind (fun p -> p.LineRule)
        Alignment = para |> Option.bind (fun p -> p.Alignment) |> Option.map alignment
        IndentLeft = para |> Option.bind (fun p -> twipsToPoints p.IndentLeft)
        IndentRight = para |> Option.bind (fun p -> twipsToPoints p.IndentRight)
        FirstLineIndent = para |> Option.bind (fun p -> twipsToPoints p.FirstLineIndent)
        BackgroundColor = para |> Option.bind (fun p -> color p.BackgroundColor)
        OutlineLevel = para |> Option.bind (fun p -> intOpt p.OutlineLevel)
        KeepNext = para |> Option.bind (fun p -> p.KeepNext)
        KeepLines = para |> Option.bind (fun p -> p.KeepLines)
        PageBreakBefore = para |> Option.bind (fun p -> p.PageBreakBefore)
        WidowControl = para |> Option.bind (fun p -> p.WidowControl)
    }

let private pageSetup (doc: DocxReader.WordDocument) =
    match doc.PageSetup with
    | None -> defaultPageSetup
    | Some setup ->
        {
            PageWidth = setup.PageWidth |> Option.defaultValue defaultPageSetup.PageWidth
            PageHeight = setup.PageHeight |> Option.defaultValue defaultPageSetup.PageHeight
            MarginLeft = setup.MarginLeft |> Option.defaultValue defaultPageSetup.MarginLeft
            MarginRight = setup.MarginRight |> Option.defaultValue defaultPageSetup.MarginRight
            MarginTop = setup.MarginTop |> Option.defaultValue defaultPageSetup.MarginTop
            MarginBottom = setup.MarginBottom |> Option.defaultValue defaultPageSetup.MarginBottom
            HeaderDistance = setup.HeaderDistance |> Option.defaultValue defaultPageSetup.HeaderDistance
            FooterDistance = setup.FooterDistance |> Option.defaultValue defaultPageSetup.FooterDistance
            DifferentFirstPage = setup.DifferentFirstPage |> Option.defaultValue defaultPageSetup.DifferentFirstPage
            EvenAndOddHeaders = setup.EvenAndOddHeaders |> Option.defaultValue defaultPageSetup.EvenAndOddHeaders
        }

let private headerFooterPart resolver images styleMap comments notes (item: DocxReader.HeaderFooter) =
    {
        Kind = item.Type
        RelationshipId = item.RelationshipId
        Blocks = item.Content |> List.map (mkPara resolver images styleMap comments notes >> Paragraph)
    }

let importDoc (doc: DocxReader.WordDocument) =
    let styleMap = DocumentProcessors.StyleResolution.buildStyleMapWithTheme doc.Theme doc.Styles
    let tableStyleMap = doc.Styles |> List.map (fun style -> style.Id, style) |> Map.ofList
    let resolver = numberingResolver doc.NumberingDefinitions
    let comments = doc.Comments |> List.map (fun c -> c.Id, c) |> Map.ofList
    let notes =
        let noteText (note: DocxReader.Note) =
            note.Content
            |> List.collect (fun para -> para.Runs)
            |> List.map (fun run -> EP.extractRunText Map.empty run.Content)
            |> String.concat ""
        (doc.Footnotes |> List.map (fun n -> ("footnote", n.Id), noteText n)) @
        (doc.Endnotes |> List.map (fun n -> ("endnote", n.Id), noteText n))
        |> Map.ofList

    let blocks =
        doc.Content
        |> List.map (function
            | DocxReader.Para para -> Paragraph (mkPara resolver doc.Images styleMap comments notes para)
            | DocxReader.Tbl table -> tableBlock resolver doc.Images styleMap tableStyleMap comments notes table)

    let firstParagraph =
        blocks
        |> List.tryPick (function
            | Paragraph para
            | Heading (_, para) -> Some para
            | Table table -> table.Rows |> List.tryPick (fun row -> row.Cells |> List.tryPick (fun cell -> cell.Content |> List.tryHead))
            | TextBox textBox -> textBox.Content |> List.tryHead)

    let headerParts = doc.Headers |> List.map (headerFooterPart resolver doc.Images styleMap comments notes)
    let footerParts = doc.Footers |> List.map (headerFooterPart resolver doc.Images styleMap comments notes)

    {
        Blocks = blocks
        Selection = firstParagraph |> Option.map (fun para -> { Anchor = { BlockId = para.Id; Offset = 0 }; Focus = { BlockId = para.Id; Offset = 0 } })
        PageSetup = pageSetup doc
        Styles = doc.Styles |> List.map namedStyle
        Headers = headerParts |> List.tryHead |> Option.map (fun part -> part.Blocks |> List.choose (function | Paragraph p -> Some p | Heading (_, p) -> Some p | _ -> None)) |> Option.defaultValue []
        Footers = footerParts |> List.tryHead |> Option.map (fun part -> part.Blocks |> List.choose (function | Paragraph p -> Some p | Heading (_, p) -> Some p | _ -> None)) |> Option.defaultValue []
        HeaderParts = headerParts
        FooterParts = footerParts
    }

#if NATIVE_RUNNER_MAIN
let private renderDocx (inputPath: string) (outputPath: string) (width: int) (height: int) =
    initWordFont ()
    let bytes = File.ReadAllBytes inputPath
    let doc =
        match DocxReader.loadBytesNative bytes with
        | Some parsed -> parsed
        | None -> failwithf "Could not parse DOCX: %s" inputPath

    let wordDoc = { importDoc doc with Selection = None }
    let widget = Composit.Ui.WordEditor.scrollEditor (float width) (float height) wordDoc
    let backend = Raster.create()
    let framebuffer = Composit.Ui.Render.renderWithBg Color.white backend width height widget
    let png = PngEncode.encode framebuffer

    Directory.CreateDirectory(Path.GetDirectoryName outputPath) |> ignore
    File.WriteAllBytes(outputPath, png)
    printfn "native runner ok: %s -> %s (%dx%d)" inputPath outputPath width height

[<EntryPoint>]
let main argv =
    try
        match argv with
        | [| input; output |] ->
            renderDocx (fullPath input) (fullPath output) 900 1600
            0
        | [| input; output; width |] ->
            renderDocx (fullPath input) (fullPath output) (Int32.Parse width) 1600
            0
        | [| input; output; width; height |] ->
            renderDocx (fullPath input) (fullPath output) (Int32.Parse width) (Int32.Parse height)
            0
        | _ ->
            usage()
            2
    with ex ->
        eprintfn "%O" ex
        1
#endif
