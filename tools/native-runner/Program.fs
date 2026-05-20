module NativeRunner

open System
open System.IO
open Renderer
open RedBat.Word
open RedBat.Word.Document
open RedBat.Word.TextBuffer
open RedBat.Word.Rope

module EP = DocumentProcessors.EditorProcessor
module SR = DocumentProcessors.StyleResolution
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

let private borderColor value =
    match value with
    | Some ""
    | Some "auto" -> Some "#000000"
    | _ -> color value

let private shadingColorValue value =
    match value with
    | Some "" -> None
    | Some "auto" -> Some "auto"
    | _ -> color value

let private shading (value: DocxReader.Shading) : Shading = {
    Pattern = value.Pattern
    ForegroundColor = shadingColorValue value.Color
    BackgroundColor = shadingColorValue value.Fill
}

let private directShadingBackgroundColor (shading: DocxReader.Shading) =
    match shading.Pattern with
    | None
    | Some "clear" -> color shading.Fill
    | _ -> None

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
        Color = borderColor b.Color
        Width = parseFloat b.Size |> Option.map (fun size -> size / 6.0)
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

let private mergeRunProperties (baseProps: DocxReader.RunProperties) (props: DocxReader.RunProperties) : DocxReader.RunProperties = {
    Style = props.Style |> Option.orElse baseProps.Style
    Bold = props.Bold |> Option.orElse baseProps.Bold
    Italic = props.Italic |> Option.orElse baseProps.Italic
    Underline = props.Underline |> Option.orElse baseProps.Underline
    AllCaps = props.AllCaps |> Option.orElse baseProps.AllCaps
    SmallCaps = props.SmallCaps |> Option.orElse baseProps.SmallCaps
    FontSize = props.FontSize |> Option.orElse baseProps.FontSize
    CharacterSpacing = props.CharacterSpacing |> Option.orElse baseProps.CharacterSpacing
    FontFamily =
        match props.FontFamily, props.ThemeFont with
        | Some family, _ -> Some family
        | None, Some _ -> None
        | None, None -> baseProps.FontFamily
    ThemeFont =
        match props.FontFamily, props.ThemeFont with
        | Some _, None -> None
        | _, Some themeFont -> Some themeFont
        | None, None -> baseProps.ThemeFont
    Color = props.Color |> Option.orElse baseProps.Color
    ThemeColor = props.ThemeColor |> Option.orElse baseProps.ThemeColor
    ThemeTint = props.ThemeTint |> Option.orElse baseProps.ThemeTint
    ThemeShade = props.ThemeShade |> Option.orElse baseProps.ThemeShade
    Highlight = props.Highlight |> Option.orElse baseProps.Highlight
    Strike = props.Strike |> Option.orElse baseProps.Strike
    VertAlign = props.VertAlign |> Option.orElse baseProps.VertAlign
}

let private mergeParagraphProperties (baseProps: DocxReader.ParagraphProperties) (props: DocxReader.ParagraphProperties) : DocxReader.ParagraphProperties = {
    Style = props.Style |> Option.orElse baseProps.Style
    ContentControlId = props.ContentControlId |> Option.orElse baseProps.ContentControlId
    ContentControlTag = props.ContentControlTag |> Option.orElse baseProps.ContentControlTag
    ContentControlAlias = props.ContentControlAlias |> Option.orElse baseProps.ContentControlAlias
    Alignment = props.Alignment |> Option.orElse baseProps.Alignment
    IndentLeft = props.IndentLeft |> Option.orElse baseProps.IndentLeft
    IndentRight = props.IndentRight |> Option.orElse baseProps.IndentRight
    FirstLineIndent = props.FirstLineIndent |> Option.orElse baseProps.FirstLineIndent
    HangingIndent = props.HangingIndent |> Option.orElse baseProps.HangingIndent
    RightTabPosition = props.RightTabPosition |> Option.orElse baseProps.RightTabPosition
    RightTabLeader = props.RightTabLeader |> Option.orElse baseProps.RightTabLeader
    SpacingBefore = props.SpacingBefore |> Option.orElse baseProps.SpacingBefore
    SpacingBeforeAuto = props.SpacingBeforeAuto |> Option.orElse baseProps.SpacingBeforeAuto
    SpacingAfter = props.SpacingAfter |> Option.orElse baseProps.SpacingAfter
    SpacingAfterAuto = props.SpacingAfterAuto |> Option.orElse baseProps.SpacingAfterAuto
    LineSpacing = props.LineSpacing |> Option.orElse baseProps.LineSpacing
    LineRule = props.LineRule |> Option.orElse baseProps.LineRule
    OutlineLevel = props.OutlineLevel |> Option.orElse baseProps.OutlineLevel
    KeepNext = props.KeepNext |> Option.orElse baseProps.KeepNext
    KeepLines = props.KeepLines |> Option.orElse baseProps.KeepLines
    PageBreakBefore = props.PageBreakBefore |> Option.orElse baseProps.PageBreakBefore
    WidowControl = props.WidowControl |> Option.orElse baseProps.WidowControl
    ContextualSpacing = props.ContextualSpacing |> Option.orElse baseProps.ContextualSpacing
    BackgroundColor = props.BackgroundColor |> Option.orElse baseProps.BackgroundColor
    ParagraphRunProperties =
        match baseProps.ParagraphRunProperties, props.ParagraphRunProperties with
        | Some baseRun, Some run -> Some (mergeRunProperties baseRun run)
        | None, run -> run
        | run, None -> run
    ParagraphBorders = props.ParagraphBorders |> Option.orElse baseProps.ParagraphBorders
    ListInfo = props.ListInfo |> Option.orElse baseProps.ListInfo
    SectionBreak = props.SectionBreak |> Option.orElse baseProps.SectionBreak
    SectionPageNumberStart = props.SectionPageNumberStart |> Option.orElse baseProps.SectionPageNumberStart
    RenderedPageBreakBefore = props.RenderedPageBreakBefore |> Option.orElse baseProps.RenderedPageBreakBefore
    HeaderReferences = if props.HeaderReferences.IsEmpty then baseProps.HeaderReferences else props.HeaderReferences
    FooterReferences = if props.FooterReferences.IsEmpty then baseProps.FooterReferences else props.FooterReferences
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

let private mergeParagraphPropertyOptions baseProps props =
    match baseProps, props with
    | Some baseValue, Some value -> Some (mergeParagraphProperties baseValue value)
    | None, value -> value
    | value, None -> value

let private resolveTableParagraphProperties selector (styleMap: Map<string, DocxReader.StyleDefinition>) styleId =
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
                mergeParagraphPropertyOptions baseProps (selector style)

    styleId |> Option.bind (resolve Set.empty)

let private paraBorder (value: DocxReader.ParagraphBorder option) : RedBat.Word.Document.ParagraphBorder option =
    value |> Option.map (fun b -> {
        Style = b.Style
        Color = borderColor b.Color
        Width = parseFloat b.Width |> Option.map (fun size -> size / 6.0)
        Space = parseFloat b.Space
    })

let private paraBorders (value: DocxReader.ParagraphBorders option) : RedBat.Word.Document.ParagraphBorders option =
    value |> Option.map (fun b -> {
        Top = paraBorder b.Top
        Right = paraBorder b.Right
        Bottom = paraBorder b.Bottom
        Left = paraBorder b.Left
    })

let private paragraphBorderAdvanceToLayoutPx (value: DocxReader.ParagraphBorder option) =
    value
    |> Option.bind (fun border ->
        match border.Style |> Option.map (fun style -> style.ToLowerInvariant()) with
        | Some "none"
        | Some "nil" -> None
        | _ ->
            let widthPt = parseFloat border.Width |> Option.map (fun size -> size / 8.0) |> Option.defaultValue 0.5
            let spacePt = parseFloat border.Space |> Option.defaultValue 0.0
            Some ((spacePt + widthPt * 2.0) * 96.0 / 72.0))

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

let private wordAutoSpacingPoints = 14.0

let private spacingValue value auto =
    if auto = Some true then Some wordAutoSpacingPoints else twipsToPoints value

let private textStyle (run: EP.EditorRun) : RedBat.Word.TextBuffer.TextStyle =
    {
        defaultStyle with
            Bold = run.Bold
            Italic = run.Italic
            Underline = run.Underline
            Strike = run.Strike
            FontSize = run.FontSize
            CharacterSpacing = run.CharacterSpacing
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
            Hyperlink = run.Hyperlink
            BookmarkStartId = run.BookmarkStartId
            BookmarkName = run.BookmarkName
            BookmarkEndId = run.BookmarkEndId
            FieldCode = run.FieldCode
            FieldResult = run.FieldResult
            RenderedPageBreakBefore = run.RenderedPageBreakBefore
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
            let style = textStyle run
            parts.Add(run.Text)
            if style.RenderedPageBreakBefore && run.Text.Length > 1 then
                bufferRuns.Add(Text { Start = offset; Length = 1; Style = style })
                bufferRuns.Add(Text { Start = offset + 1; Length = run.Text.Length - 1; Style = { style with RenderedPageBreakBefore = false } })
            else
                bufferRuns.Add(Text { Start = offset; Length = run.Text.Length; Style = style })
            offset <- offset + run.Text.Length

    { Rope = Rope.ofString (String.Concat parts); Runs = bufferRuns |> Seq.toList }

let private hasDirectParagraphFormatting (props: DocxReader.ParagraphProperties) =
    props.Style.IsSome
    || props.Alignment.IsSome
    || props.IndentLeft.IsSome
    || props.IndentRight.IsSome
    || props.FirstLineIndent.IsSome
    || props.HangingIndent.IsSome
    || props.RightTabPosition.IsSome
    || props.RightTabLeader.IsSome
    || props.SpacingBefore.IsSome
    || props.SpacingBeforeAuto.IsSome
    || props.SpacingAfter.IsSome
    || props.SpacingAfterAuto.IsSome
    || props.LineSpacing.IsSome
    || props.LineRule.IsSome
    || props.OutlineLevel.IsSome
    || props.KeepNext.IsSome
    || props.KeepLines.IsSome
    || props.PageBreakBefore.IsSome
    || props.WidowControl.IsSome
    || props.ContextualSpacing.IsSome
    || props.BackgroundColor.IsSome
    || props.ParagraphRunProperties.IsSome
    || props.ParagraphBorders.IsSome
    || props.ListInfo.IsSome
    || props.SectionBreak.IsSome
    || props.SectionPageNumberStart.IsSome
    || props.RenderedPageBreakBefore.IsSome
    || not props.HeaderReferences.IsEmpty
    || not props.FooterReferences.IsEmpty

let private isPureEmptyParagraph (para: DocxReader.Paragraph) =
    para.Runs.IsEmpty && not (hasDirectParagraphFormatting para.Properties)

let private documentDefaultRunFontSize (styleMap: Map<string, SR.ResolvedRunProperties * SR.ResolvedParagraphProperties>) =
    Map.tryFind SR.documentDefaultsStyleId styleMap
    |> Option.map (fun (runProps, _) -> runProps.FontSize)

let private paraMetadata styleMap resolver useDocumentMarkForPureEmpty defaultMissingSpacingToZero (para: DocxReader.Paragraph) (editorPara: EP.EditorParagraph) =
    let props = para.Properties
    let resolvedParaProps = DocumentProcessors.StyleResolution.resolveParagraphProps styleMap props
    let numberingDef = resolvedNumberingDefinition resolver props
    let numberingIndentLeft = numberingDef |> Option.bind (fun def -> twipsToPoints def.IndentLeft)
    let numberingRightTab = numberingDef |> Option.bind (fun def -> twipsToCanvasPixels def.RightTabPosition)
    let numberingFirstLineIndent =
        numberingDef
        |> Option.bind (fun def ->
            match twipsToPoints def.FirstLineIndent, twipsToPoints def.HangingIndent with
            | Some firstLine, _ -> Some firstLine
            | None, Some hanging -> Some -hanging
            | None, None -> None)
    let missingSpacingFallback = if defaultMissingSpacingToZero then Some 0.0 else None
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
            RightTabPosition = editorPara.RightTabPosition |> Option.orElse numberingRightTab
            RightTabLeader = editorPara.RightTabLeader
            SpacingBefore = if props.SpacingBeforeAuto = Some true then Some wordAutoSpacingPoints else editorPara.SpacingBefore |> Option.orElse missingSpacingFallback
            SpacingAfter = if props.SpacingAfterAuto = Some true then Some wordAutoSpacingPoints else editorPara.SpacingAfter |> Option.orElse missingSpacingFallback
            ContextualSpacing = props.ContextualSpacing
            LineSpacing = editorPara.LineSpacing
            LineSpacingRule = props.LineRule
            WidowControl = props.WidowControl
            BackgroundColor = color props.BackgroundColor
            ParagraphBorders = paraBorders props.ParagraphBorders
            ParagraphBorderBefore = resolvedParaProps.ParagraphBorders |> Option.bind (fun borders -> paragraphBorderAdvanceToLayoutPx borders.Top)
            ParagraphBorderAfter = resolvedParaProps.ParagraphBorders |> Option.bind (fun borders -> paragraphBorderAdvanceToLayoutPx borders.Bottom)
            PageBreakBefore = props.PageBreakBefore |> Option.defaultValue false
            PageBreakAfter = para.Runs |> List.exists (fun run -> match run.Content with | DocxReader.PageBreak -> true | _ -> false)
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
            ParagraphMarkFontSize =
                props.ParagraphRunProperties
                |> Option.map (fun rPr -> TC.fontSizeToPoints rPr.FontSize)
                |> Option.orElseWith (fun () ->
                    if useDocumentMarkForPureEmpty && isPureEmptyParagraph para then documentDefaultRunFontSize styleMap else None)
    }

let private directParagraphFormattingResetsMissingSpacing (props: DocxReader.ParagraphProperties) =
    props.IndentLeft.IsSome
    || props.IndentRight.IsSome
    || props.FirstLineIndent.IsSome
    || props.HangingIndent.IsSome
    || props.RightTabPosition.IsSome
    || props.RightTabLeader.IsSome
    || props.ParagraphRunProperties.IsSome
    || props.BackgroundColor.IsSome
    || props.ParagraphBorders.IsSome
    || props.ContextualSpacing.IsSome

let private tableStyleParagraphResetsMissingSpacing (props: DocxReader.ParagraphProperties) =
    props.LineSpacing.IsSome || props.LineRule.IsSome

let private mkPara resolver images styleMap comments notes useDocumentMarkForPureEmpty defaultMissingSpacingToZero para =
    let editorPara = EP.toEditorParagraph None styleMap comments notes para
    {
        Id = Guid.NewGuid().ToString("N")
        Buffer = editorRunsToBuffer images editorPara.Runs
        Metadata = paraMetadata styleMap resolver useDocumentMarkForPureEmpty defaultMissingSpacingToZero para editorPara
    }

let private tableCell resolver images styleMap comments notes (paragraphProperties: DocxReader.ParagraphProperties option) (cell: DocxReader.TableCell) =
    let props = cell.Properties
    let applyTableParagraphProperties (para: DocxReader.Paragraph) =
        match paragraphProperties with
        | Some styleProps ->
            let properties = mergeParagraphProperties styleProps para.Properties
            let properties =
                if para.Properties.Style.IsNone && styleProps.Style.IsNone then
                    { properties with Style = Some DocxReader.Styles.defaultStyleId }
                else
                    properties
            { para with Properties = properties }
        | None -> para
    let emptyPara : DocxReader.Paragraph = { Runs = []; Properties = DocxReader.Docx.paragraphProps XmlDsl.Empty }
    let defaultMissingSpacingToZero =
        paragraphProperties |> Option.exists tableStyleParagraphResetsMissingSpacing
    let content =
        cell.Paragraphs
        |> List.map applyTableParagraphProperties
        |> List.map (mkPara resolver images styleMap comments notes false defaultMissingSpacingToZero)
    {
        Content = if content.IsEmpty then [mkPara resolver images styleMap comments notes false defaultMissingSpacingToZero (applyTableParagraphProperties emptyPara)] else content
        ColSpan = props |> Option.bind (fun p -> p.ColSpan) |> Option.defaultValue 1
        BackgroundColor = props |> Option.bind (fun p -> p.Shading) |> Option.bind directShadingBackgroundColor
        Shading = props |> Option.bind (fun p -> p.Shading) |> Option.map shading
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
    let tableParagraphProperties =
        resolveTableParagraphProperties (fun style -> style.ParagraphProperties) tableStyleMap table.StyleId
    let firstRowParagraphProperties =
        resolveTableParagraphProperties (fun style -> style.FirstRowParagraphProperties) tableStyleMap table.StyleId
    let paragraphPropertiesForRow rowIndex =
        let firstRowEnabled = table.Look |> Option.map _.FirstRow |> Option.defaultValue false
        let conditionalProperties = if rowIndex = 0 && firstRowEnabled then firstRowParagraphProperties else None
        mergeParagraphPropertyOptions tableParagraphProperties conditionalProperties
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
            |> List.mapi (fun rowIndex row -> {
                Cells = row.Cells |> List.map (tableCell resolver images styleMap comments notes (paragraphPropertiesForRow rowIndex))
                IsHeader = row.IsHeader
                Height = parseFloat row.Height
                HeightRule = row.HeightRule
                CanSplit = row.CanSplit
                CellSpacing = twipsToCanvasPixels row.CellSpacing
                PaddingTop = twipsToCanvasPixels row.PaddingTop
                PaddingRight = twipsToCanvasPixels row.PaddingRight
                PaddingBottom = twipsToCanvasPixels row.PaddingBottom
                PaddingLeft = twipsToCanvasPixels row.PaddingLeft
                GridBefore = row.GridBefore |> Option.defaultValue 0 |> max 0
                GridAfter = row.GridAfter |> Option.defaultValue 0 |> max 0
                WidthBefore = twipsToCanvasPixels row.WidthBefore
                WidthAfter = twipsToCanvasPixels row.WidthAfter
                BordersOverride = tableBorders row.BordersOverride
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
        BackgroundColor = tableShading |> Option.bind directShadingBackgroundColor
        Shading = tableShading |> Option.map shading
        CellSpacing = twipsToCanvasPixels table.CellSpacing
        PaddingTop = tablePadding (fun p -> p.PaddingTop)
        PaddingRight = tablePadding (fun p -> p.PaddingRight)
        PaddingBottom = tablePadding (fun p -> p.PaddingBottom)
        PaddingLeft = tablePadding (fun p -> p.PaddingLeft)
        Floating = table.Floating
    }

let private resolveThemeFontFamily (theme: DocxReader.Theme option) (themeFont: string option) =
    // Mirror DocxReader.resolveThemeFont: <w:asciiTheme="majorHAnsi"> picks
    // the document theme's major font (Cambria by default), "minorHAnsi"
    // picks the minor font (Calibri).  Without this resolution, styles that
    // express their face only via the asciiTheme attribute (e.g. the default
    // Office Heading1 style which has no explicit w:ascii) lose their font
    // and fall back to the hard-coded "Calibri" in WordStyles.headingText.
    match theme, themeFont with
    | Some t, Some value when value.StartsWith("major") -> t.MajorFont
    | Some t, Some value when value.StartsWith("minor") -> t.MinorFont
    | _, Some value when value.StartsWith("major") -> Some "Cambria"
    | _, Some value when value.StartsWith("minor") -> Some "Calibri"
    | _ -> None

let private namedStyle (theme: DocxReader.Theme option) (style: DocxReader.StyleDefinition) : NamedStyle =
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
                    CharacterSpacing = twipsToCanvasPixels r.CharacterSpacing |> Option.defaultValue 0.0
                    FontFamily =
                        r.FontFamily
                        |> Option.orElseWith (fun () -> resolveThemeFontFamily theme r.ThemeFont)
                    Color = color r.Color
                    Highlight = r.Highlight
                    VertAlign = r.VertAlign
            })
        TextBold = run |> Option.bind (fun r -> r.Bold)
        TextItalic = run |> Option.bind (fun r -> r.Italic)
        TextUnderline = run |> Option.bind (fun r -> r.Underline)
        TextAllCaps = run |> Option.bind (fun r -> r.AllCaps)
        TextSmallCaps = run |> Option.bind (fun r -> r.SmallCaps)
        TextStrike = run |> Option.bind (fun r -> r.Strike)
        SpacingBefore = para |> Option.bind (fun p -> spacingValue p.SpacingBefore p.SpacingBeforeAuto)
        SpacingAfter = para |> Option.bind (fun p -> spacingValue p.SpacingAfter p.SpacingAfterAuto)
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
        Blocks = item.Content |> List.map (mkPara resolver images styleMap comments notes true false >> Paragraph)
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
            | DocxReader.Para para -> Paragraph (mkPara resolver doc.Images styleMap comments notes true (directParagraphFormattingResetsMissingSpacing para.Properties) para)
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
        Styles = doc.Styles |> List.map (namedStyle doc.Theme)
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
