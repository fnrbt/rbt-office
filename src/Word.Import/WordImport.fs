module Rbt.Word.Import

open System
open Rbt.Word
open Rbt.Word.Document
open Rbt.Word.TextBuffer
open Rbt.Word.Rope

module EP = DocumentProcessors.EditorProcessor
module SR = DocumentProcessors.StyleResolution
module TC = DocumentProcessors.TypeConversion

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
    UnderlineStyle = props.UnderlineStyle |> Option.orElse baseProps.UnderlineStyle
    UnderlineColor = props.UnderlineColor |> Option.orElse baseProps.UnderlineColor
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
    DoubleStrike = props.DoubleStrike |> Option.orElse baseProps.DoubleStrike
    Hidden = props.Hidden |> Option.orElse baseProps.Hidden
    Position = props.Position |> Option.orElse baseProps.Position
    EmphasisMark = props.EmphasisMark |> Option.orElse baseProps.EmphasisMark
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
    FirstLineIndent = if props.FirstLineIndent.IsSome || props.HangingIndent.IsSome then props.FirstLineIndent else baseProps.FirstLineIndent
    HangingIndent = if props.FirstLineIndent.IsSome || props.HangingIndent.IsSome then props.HangingIndent else baseProps.HangingIndent
    RightTabPosition = props.RightTabPosition |> Option.orElse baseProps.RightTabPosition
    RightTabLeader = props.RightTabLeader |> Option.orElse baseProps.RightTabLeader
    TabStops = (if List.isEmpty props.TabStops then baseProps.TabStops else props.TabStops)
    Rtl = props.Rtl |> Option.orElse baseProps.Rtl
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
    SectionColumns = props.SectionColumns |> Option.orElse baseProps.SectionColumns
    SectionGeometry = props.SectionGeometry |> Option.orElse baseProps.SectionGeometry
    SectionPageNumberStart = props.SectionPageNumberStart |> Option.orElse baseProps.SectionPageNumberStart
    SectionPageNumberFormat = props.SectionPageNumberFormat |> Option.orElse baseProps.SectionPageNumberFormat
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

let private mergeRunPropertyOptions (baseProps: DocxReader.RunProperties option) (props: DocxReader.RunProperties option) =
    match baseProps, props with
    | Some baseValue, Some value -> Some (mergeRunProperties baseValue value)
    | None, value -> value
    | value, None -> value

let private paragraphPropertiesWithRunProperties (runProps: DocxReader.RunProperties option) (props: DocxReader.ParagraphProperties option) : DocxReader.ParagraphProperties option =
    match runProps, props with
    | None, value -> value
    | Some runProps, Some props ->
        let paragraphRunProperties = mergeRunPropertyOptions (Some runProps) props.ParagraphRunProperties
        Some { props with ParagraphRunProperties = paragraphRunProperties }
    | Some runProps, None ->
        Some { DocxReader.Docx.paragraphProps XmlDsl.Empty with ParagraphRunProperties = Some runProps }

let private resolveTableRunProperties selector (styleMap: Map<string, DocxReader.StyleDefinition>) styleId =
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
                mergeRunPropertyOptions baseProps (selector style)

    styleId |> Option.bind (resolve Set.empty)

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

let private paraBorder (value: DocxReader.ParagraphBorder option) : Rbt.Word.Document.ParagraphBorder option =
    value |> Option.map (fun b -> {
        Style = b.Style
        Color = borderColor b.Color
        Width = parseFloat b.Width |> Option.map (fun size -> size / 6.0)
        Space = parseFloat b.Space
    })

let private paraBorders (value: DocxReader.ParagraphBorders option) : Rbt.Word.Document.ParagraphBorders option =
    value |> Option.map (fun b -> {
        Top = paraBorder b.Top
        Right = paraBorder b.Right
        Bottom = paraBorder b.Bottom
        Left = paraBorder b.Left
    })

let private styleIsNone (style: string option) =
    match style with
    | Some value ->
        String.Equals(value, "none", StringComparison.OrdinalIgnoreCase)
        || String.Equals(value, "nil", StringComparison.OrdinalIgnoreCase)
    | None -> false

let private paragraphBorderAdvanceToLayoutPx (value: DocxReader.ParagraphBorder option) =
    value
    |> Option.bind (fun border ->
        if styleIsNone border.Style then
            None
        else
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
    // Style-level exact / at-least is px (twips / 15), as in the browser style catalog.
    | Some "exact"
    | Some "atLeast" -> twipsToCanvasPixels value
    | _ -> TC.lineSpacingToMultiple value

let private wordAutoSpacingPoints = 14.0

let private spacingValue value auto =
    if auto = Some true then Some wordAutoSpacingPoints else twipsToPoints value

let private textStyle (run: EP.EditorRun) : Rbt.Word.TextBuffer.TextStyle =
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
            // Only a target that may be followed is a link (see DocxReader.isActiveHyperlinkTarget).
            Hyperlink = run.Hyperlink |> Option.filter DocxReader.isActiveHyperlinkTarget
            InactiveHyperlink = run.Hyperlink |> Option.filter (fun target -> not (DocxReader.isActiveHyperlinkTarget target))
            BookmarkStartId = run.BookmarkStartId
            BookmarkName = run.BookmarkName
            BookmarkEndId = run.BookmarkEndId
            FieldCode = run.FieldCode
            FieldResult = run.FieldResult
            EquationXml = run.EquationXml
            RenderedPageBreakBefore = run.RenderedPageBreakBefore
            VertAlign = run.VertAlign
    }

/// The raw style definitions of the document being imported. The typed
/// DocumentProcessors model (EditorRun / ResolvedParagraphProperties) has no
/// slot for hidden text, caps, tab stops, RTL or the line rule, so those are
/// resolved here from the raw DocxReader properties, the way the browser
/// importer (DocxReader.toEditorParagraphs) does. Thread-static: importDoc can
/// run on several threads at once.
type private RawStyles() =
    [<ThreadStatic; DefaultValue>]
    static val mutable private slot : Map<string, DocxReader.StyleDefinition>
    static member Current
        with get () = if isNull (box RawStyles.slot) then Map.empty else RawStyles.slot
        and set (value: Map<string, DocxReader.StyleDefinition>) = RawStyles.slot <- value

let private rawStyles () = RawStyles.Current

let private emptyRawRunProps : DocxReader.RunProperties = DocxReader.Docx.runProps XmlDsl.Empty
let private emptyRawParagraphProps : DocxReader.ParagraphProperties = DocxReader.Docx.paragraphProps XmlDsl.Empty

/// Run and paragraph properties of a style with its basedOn chain applied.
let rec private rawStyleProps (visited: Set<string>) (styleId: string) : DocxReader.RunProperties * DocxReader.ParagraphProperties =
    match Map.tryFind styleId (rawStyles ()) with
    | Some style when not (visited.Contains styleId) ->
        let baseRun, basePara =
            match style.BasedOn with
            | Some basedOn -> rawStyleProps (visited.Add styleId) basedOn
            | None -> emptyRawRunProps, emptyRawParagraphProps
        (style.RunProperties |> Option.map (mergeRunProperties baseRun) |> Option.defaultValue baseRun),
        (style.ParagraphProperties |> Option.map (mergeParagraphProperties basePara) |> Option.defaultValue basePara)
    | _ -> emptyRawRunProps, emptyRawParagraphProps

let private rawParagraphStyleId (props: DocxReader.ParagraphProperties) =
    props.Style
    |> Option.orElseWith (fun () ->
        rawStyles () |> Map.tryPick (fun id style -> if style.StyleType = "paragraph" && style.IsDefault then Some id else None))

/// Document defaults -> paragraph style chain -> direct formatting.
let private effectiveRawParagraphProps (props: DocxReader.ParagraphProperties) =
    let _, defaults = rawStyleProps Set.empty DocxReader.Styles.defaultStyleId
    let _, styled = rawParagraphStyleId props |> Option.map (rawStyleProps Set.empty) |> Option.defaultValue (emptyRawRunProps, emptyRawParagraphProps)
    mergeParagraphProperties (mergeParagraphProperties defaults styled) props

/// Document defaults -> paragraph style -> paragraph mark -> character style -> direct formatting.
let private effectiveRawRunProps (paraProps: DocxReader.ParagraphProperties) (runProps: DocxReader.RunProperties) =
    let defaults, _ = rawStyleProps Set.empty DocxReader.Styles.defaultStyleId
    let styled, _ = rawParagraphStyleId paraProps |> Option.map (rawStyleProps Set.empty) |> Option.defaultValue (emptyRawRunProps, emptyRawParagraphProps)
    let mark = paraProps.ParagraphRunProperties |> Option.defaultValue emptyRawRunProps
    let character, _ = runProps.Style |> Option.map (rawStyleProps Set.empty) |> Option.defaultValue (emptyRawRunProps, emptyRawParagraphProps)
    mergeRunProperties (mergeRunProperties (mergeRunProperties (mergeRunProperties defaults styled) mark) character) runProps

/// What EditorRun cannot carry, taken from the source run: the browser
/// importer delivers all of these, the native one used to drop them.
let private withSourceRunProps (paraProps: DocxReader.ParagraphProperties) (run: DocxReader.Run) (style: Rbt.Word.TextBuffer.TextStyle) : Rbt.Word.TextBuffer.TextStyle =
    let sourceProps =
        match run.Content with
        | DocxReader.Link link -> link.Properties
        | _ -> run.Properties
    let raw = effectiveRawRunProps paraProps sourceProps
    // Properties set directly on the run override the styles even when the
    // value is the default (w:b w:val="0" under a bold style, an explicit size).
    // The paragraph mark's run properties count as run-level too: this importer
    // applies them to the paragraph's runs (see effectiveRawRunProps), and they
    // are not something a named style would produce again.
    let mark = paraProps.ParagraphRunProperties |> Option.defaultValue emptyRawRunProps
    let explicitFlags =
        let flag (direct: 'a option) (fromMark: 'a option) (bit: int) = if direct.IsSome || fromMark.IsSome then bit else 0
        flag sourceProps.Bold mark.Bold ExplicitFormat.Bold
        ||| flag sourceProps.Italic mark.Italic ExplicitFormat.Italic
        ||| flag sourceProps.Underline mark.Underline ExplicitFormat.Underline
        ||| flag sourceProps.AllCaps mark.AllCaps ExplicitFormat.AllCaps
        ||| flag sourceProps.SmallCaps mark.SmallCaps ExplicitFormat.SmallCaps
        ||| flag sourceProps.Strike mark.Strike ExplicitFormat.Strike
        ||| flag sourceProps.DoubleStrike mark.DoubleStrike ExplicitFormat.DoubleStrike
        ||| flag sourceProps.Hidden mark.Hidden ExplicitFormat.Hidden
        ||| flag sourceProps.FontSize mark.FontSize ExplicitFormat.FontSize
        ||| (if (sourceProps.Color |> Option.orElse mark.Color) = Some "auto" then ExplicitFormat.ColorAuto else 0)
    { style with
        Style = sourceProps.Style
        Explicit = explicitFlags
        Bold = sourceProps.Bold |> Option.defaultValue style.Bold
        Italic = sourceProps.Italic |> Option.defaultValue style.Italic
        Strike = sourceProps.Strike |> Option.defaultValue style.Strike
        AllCaps = raw.AllCaps |> Option.defaultValue false
        SmallCaps = raw.SmallCaps |> Option.defaultValue false
        DoubleStrike = raw.DoubleStrike |> Option.defaultValue false
        Hidden = raw.Hidden |> Option.defaultValue false
        UnderlineStyle = raw.UnderlineStyle
        UnderlineColor = raw.UnderlineColor
        // w:position is half-points; the model holds points.
        Position = parseFloat raw.Position |> Option.map (fun halfPoints -> halfPoints / 2.0) |> Option.defaultValue 0.0
        EmphasisMark = raw.EmphasisMark
        RevisionId = run.RevisionId
        RevisionType = run.RevisionType
        RevisionAuthor = run.RevisionAuthor
        RevisionDate = run.RevisionDate
        DeletedInsertion = run.InsertedRevision |> Option.map (fun mark -> ({ Id = mark.Id; Author = mark.Author; Date = mark.Date }: Rbt.Word.TextBuffer.RevisionMark))
        MoveName = run.MoveName }

let private editorRunsToBufferWithImages (decorate: int -> Rbt.Word.TextBuffer.TextStyle -> Rbt.Word.TextBuffer.TextStyle) (decorateImageRun: int -> Rbt.Word.TextBuffer.ImageRun -> Rbt.Word.TextBuffer.ImageRun) (images: Map<string, string>) (runs: EP.EditorRun list) =
    let decorateImage index (run: Rbt.Word.TextBuffer.Run) =
        match run with
        | Image image -> Image (decorateImageRun index image)
        | other -> other
    let mutable runIndex = -1
    let mutable offset = 0
    let parts = ResizeArray<string>()
    let bufferRuns = ResizeArray<Rbt.Word.TextBuffer.Run>()

    for run in runs do
        runIndex <- runIndex + 1
        if run.IsImage then
            let src = run.ImageSrc |> Option.orElseWith (fun () -> Map.tryFind run.Text images) |> Option.defaultValue ""
            parts.Add("\uFFFC")
            bufferRuns.Add(Image {
                Start = offset
                Src = src
                // A picture whose source is an address rather than package data is linked.
                Linked = run.ImageLinked || DocxReader.isExternalImageSource src
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
                PositionHAlign = None
                PositionVAlign = None
                WrapText = None
                WrapPolygon = []
                Opaque = None
            } |> decorateImage runIndex)
            offset <- offset + 1
        elif not (String.IsNullOrEmpty run.Text) then
            let style = decorate runIndex (textStyle run)
            parts.Add(run.Text)
            if style.RenderedPageBreakBefore && run.Text.Length > 1 then
                bufferRuns.Add(Text { Start = offset; Length = 1; Style = style })
                bufferRuns.Add(Text { Start = offset + 1; Length = run.Text.Length - 1; Style = { style with RenderedPageBreakBefore = false } })
            else
                bufferRuns.Add(Text { Start = offset; Length = run.Text.Length; Style = style })
            offset <- offset + run.Text.Length

    { Rope = Rope.ofString (String.Concat parts); Runs = bufferRuns |> Seq.toList }

let private editorRunsToBufferWith decorate (images: Map<string, string>) (runs: EP.EditorRun list) =
    editorRunsToBufferWithImages decorate (fun _ image -> image) images runs

let private editorRunsToBuffer (images: Map<string, string>) (runs: EP.EditorRun list) =
    editorRunsToBufferWith (fun _ style -> style) images runs

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
    || props.SectionPageNumberFormat.IsSome
    || props.RenderedPageBreakBefore.IsSome
    || not props.HeaderReferences.IsEmpty
    || not props.FooterReferences.IsEmpty

let private isPureEmptyParagraph (para: DocxReader.Paragraph) =
    para.Runs.IsEmpty && not (hasDirectParagraphFormatting para.Properties)

let private documentDefaultRunFontSize (styleMap: Map<string, SR.ResolvedRunProperties * SR.ResolvedParagraphProperties>) =
    Map.tryFind SR.documentDefaultsStyleId styleMap
    |> Option.map (fun (runProps, _) -> runProps.FontSize)

let private documentDefaultParagraphResetsMissingSpacing (styleMap: Map<string, SR.ResolvedRunProperties * SR.ResolvedParagraphProperties>) =
    Map.tryFind SR.documentDefaultsStyleId styleMap
    |> Option.exists (fun (_, props) -> props.LineSpacing.IsSome)

let private paraMetadata styleMap resolver useDocumentMarkForPureEmpty defaultMissingSpacingToZero (para: DocxReader.Paragraph) (editorPara: EP.EditorParagraph) =
    let props = para.Properties
    let resolvedParaProps = DocumentProcessors.StyleResolution.resolveParagraphProps styleMap props
    let numberingDef = resolvedNumberingDefinition resolver props
    // Indents are canvas px (twips / 15), as the browser importer delivers them and
    // as layout reads them (its default list indent is 48 = 720 twips). They used
    // to be twips / 20 here, so natively imported text was indented 25% too little.
    // EditorParagraph carries them as twips / 20; pointsToPixels brings those over.
    let pointsToPixels (value: float option) = value |> Option.map (fun points -> points * 20.0 / 15.0)
    let rawProps = effectiveRawParagraphProps props
    let numberingIndentLeft = numberingDef |> Option.bind (fun def -> twipsToCanvasPixels def.IndentLeft)
    let numberingRightTab = numberingDef |> Option.bind (fun def -> twipsToCanvasPixels def.RightTabPosition)
    let numberingFirstLineIndent =
        numberingDef
        |> Option.bind (fun def ->
            match twipsToCanvasPixels def.FirstLineIndent, twipsToCanvasPixels def.HangingIndent with
            | Some firstLine, _ -> Some firstLine
            | None, Some hanging -> Some -hanging
            | None, None -> None)
    // Exact / at-least line spacing, in the DocxReader -> Docx.fs convention that
    // layout interprets: line height in px divided by the paragraph font size in
    // px, with that font size in ParagraphMarkFontSize. (It used to arrive here as
    // twips / 240, which made a 14pt-exact paragraph several times too tall.)
    let fixedLineSpacing =
        match rawProps.LineRule, parseFloat rawProps.LineSpacing with
        | Some rule, Some twips when rule = "exact" || rule = "atLeast" ->
            let markSize = TC.fontSizeToPoints (effectiveRawRunProps props emptyRawRunProps).FontSize
            let fontSize =
                editorPara.Runs
                |> List.filter (fun run -> not run.IsImage && not (String.IsNullOrEmpty run.Text))
                |> List.fold (fun size run -> max size run.FontSize) markSize
            Some ((twips / 15.0) / max 1.0 fontSize, fontSize)
        | _ -> None
    let missingSpacingFallback = if defaultMissingSpacingToZero then Some 0.0 else None
    {
        defaultMetadata with
            Style = props.Style
            ContentControlId = props.ContentControlId
            ContentControlTag = props.ContentControlTag
            ContentControlAlias = props.ContentControlAlias
            Alignment = alignment editorPara.Alignment
            ExplicitAlignment = props.Alignment.IsSome
            IndentLeft = pointsToPixels editorPara.IndentLeft |> Option.orElse numberingIndentLeft
            IndentRight = pointsToPixels editorPara.IndentRight
            FirstLineIndent =
                match twipsToCanvasPixels props.HangingIndent with
                | Some hanging -> Some -hanging
                | None -> pointsToPixels editorPara.FirstLineIndent |> Option.orElse numberingFirstLineIndent
            TabStops =
                rawProps.TabStops
                |> List.choose (fun (position, tabAlignment, leader) ->
                    twipsToCanvasPixels (Some position)
                    |> Option.map (fun px -> ({ Position = px; Alignment = tabAlignment; Leader = leader }: TabStop)))
            Rtl = rawProps.Rtl |> Option.defaultValue false
            RightTabPosition = editorPara.RightTabPosition |> Option.orElse numberingRightTab
            RightTabLeader = editorPara.RightTabLeader
            SpacingBefore = if props.SpacingBeforeAuto = Some true then Some wordAutoSpacingPoints else editorPara.SpacingBefore |> Option.orElse missingSpacingFallback
            SpacingAfter = if props.SpacingAfterAuto = Some true then Some wordAutoSpacingPoints else editorPara.SpacingAfter |> Option.orElse missingSpacingFallback
            ContextualSpacing = props.ContextualSpacing
            LineSpacing = fixedLineSpacing |> Option.map fst |> Option.orElse editorPara.LineSpacing
            LineSpacingRule = rawProps.LineRule
            WidowControl = props.WidowControl
            KeepNext = props.KeepNext
            KeepLines = props.KeepLines
            BackgroundColor = color props.BackgroundColor
            ParagraphBorders = paraBorders props.ParagraphBorders
            ParagraphBorderBefore = resolvedParaProps.ParagraphBorders |> Option.bind (fun borders -> paragraphBorderAdvanceToLayoutPx borders.Top)
            ParagraphBorderAfter = resolvedParaProps.ParagraphBorders |> Option.bind (fun borders -> paragraphBorderAdvanceToLayoutPx borders.Bottom)
            PageBreakBefore = props.PageBreakBefore |> Option.defaultValue false
            PageBreakAfter = para.Runs |> List.exists (fun run -> match run.Content with | DocxReader.PageBreak -> true | _ -> false)
            SectionBreak = props.SectionBreak
            SectionColumns = props.SectionColumns |> Option.map (fun c -> ({ Count = c.Count; Gap = c.Gap; Separator = c.Separator; Widths = c.Widths }: Rbt.Word.Document.SectionColumns))
            SectionPageSetup = props.SectionGeometry |> Option.map (fun g -> ({ PageWidth = g.PageWidth; PageHeight = g.PageHeight; MarginLeft = g.MarginLeft; MarginRight = g.MarginRight; MarginTop = g.MarginTop; MarginBottom = g.MarginBottom; HeaderDistance = g.HeaderDistance; FooterDistance = g.FooterDistance; DifferentFirstPage = g.DifferentFirstPage; EvenAndOddHeaders = false }: Rbt.Word.Document.SectionPageSetup))
            SectionPageNumberStart = props.SectionPageNumberStart
            SectionPageNumberFormat = props.SectionPageNumberFormat
            RenderedPageBreakBefore = props.RenderedPageBreakBefore
            HeaderReferences = props.HeaderReferences |> List.map (fun ref -> ref.Kind, ref.RelationshipId)
            FooterReferences = props.FooterReferences |> List.map (fun ref -> ref.Kind, ref.RelationshipId)
            // What the paragraph's own w:pPr states (the values above are effective).
            Explicit =
                let flag (isSet: bool) (bit: int) = if isSet then bit else 0
                flag (props.SpacingBefore.IsSome || props.SpacingBeforeAuto.IsSome) ParagraphExplicit.SpacingBefore
                ||| flag (props.SpacingAfter.IsSome || props.SpacingAfterAuto.IsSome) ParagraphExplicit.SpacingAfter
                ||| flag props.LineSpacing.IsSome ParagraphExplicit.LineSpacing
                ||| flag props.IndentLeft.IsSome ParagraphExplicit.IndentLeft
                ||| flag props.IndentRight.IsSome ParagraphExplicit.IndentRight
                ||| flag (props.FirstLineIndent.IsSome || props.HangingIndent.IsSome) ParagraphExplicit.FirstLineIndent
                ||| flag props.ContextualSpacing.IsSome ParagraphExplicit.ContextualSpacing
                ||| flag props.BackgroundColor.IsSome ParagraphExplicit.BackgroundColor
                ||| flag (not props.TabStops.IsEmpty) ParagraphExplicit.Tabs
            HasMarkProperties = props.ParagraphRunProperties.IsSome
            ExplicitMarkFontSize = props.ParagraphRunProperties |> Option.bind (fun mark -> mark.FontSize |> Option.map (fun _ -> TC.fontSizeToPoints mark.FontSize))
            ListLevelIndents = numberingDef |> Option.map (fun _ -> numberingIndentLeft, numberingFirstLineIndent)
            MarkRevision =
                para.MarkRevision
                |> Option.map (fun (revisionType, mark) ->
                    ({ RevisionType = revisionType; Mark = { Id = mark.Id; Author = mark.Author; Date = mark.Date } }: ParagraphMarkRevision))
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
                fixedLineSpacing
                |> Option.map snd
                |> Option.orElseWith (fun () -> props.ParagraphRunProperties |> Option.map (fun rPr -> TC.fontSizeToPoints rPr.FontSize))
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
    // toEditorParagraph keeps a source run or drops it by its content alone, so the
    // kept source runs line up with editorPara.Runs one to one.
    let sourceRuns =
        para.Runs
        |> List.filter (fun run -> (EP.toEditorRun None styleMap SR.defaultRunProps comments notes run).IsSome)
        |> List.toArray
    let decorate index style =
        if sourceRuns.Length = editorPara.Runs.Length && index < sourceRuns.Length then
            withSourceRunProps para.Properties sourceRuns.[index] style
        else style
    // What EditorRun does not carry about an image, from the source run.
    let decorateImage index (image: Rbt.Word.TextBuffer.ImageRun) =
        if sourceRuns.Length = editorPara.Runs.Length && index < sourceRuns.Length then
            match sourceRuns.[index].Content with
            | DocxReader.Image source ->
                { image with
                    PositionHAlign = source.PositionHAlign
                    PositionVAlign = source.PositionVAlign
                    WrapText = source.WrapText
                    WrapPolygon = source.WrapPolygon
                    // An opaque object shows its fallback picture, or a placeholder frame.
                    Src = (if source.Opaque.IsSome then DocxReader.editorImageSource images source else image.Src)
                    Opaque = source.Opaque |> Option.map (fun opaque -> ({ Xml = opaque.Xml; Kind = opaque.Kind; ParagraphLevel = opaque.ParagraphLevel }: Rbt.Word.TextBuffer.OpaqueObject)) }
            | _ -> image
        else image
    {
        Id = Guid.NewGuid().ToString("N")
        Buffer = editorRunsToBufferWithImages decorate decorateImage images editorPara.Runs
        Metadata = paraMetadata styleMap resolver useDocumentMarkForPureEmpty defaultMissingSpacingToZero para editorPara
    }

let rec private tableCell resolver images styleMap tableStyleMap comments notes (paragraphProperties: DocxReader.ParagraphProperties option) (cell: DocxReader.TableCell) =
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
    let emptyPara : DocxReader.Paragraph = { Runs = []; Properties = DocxReader.Docx.paragraphProps XmlDsl.Empty; MarkRevision = None }
    let defaultMissingSpacingToZero =
        paragraphProperties |> Option.exists tableStyleParagraphResetsMissingSpacing
    let content =
        cell.Paragraphs
        |> List.map applyTableParagraphProperties
        |> List.map (mkPara resolver images styleMap comments notes false defaultMissingSpacingToZero)
    {
        Content = if content.IsEmpty then [mkPara resolver images styleMap comments notes false defaultMissingSpacingToZero (applyTableParagraphProperties emptyPara)] else content
        // (In a table inside a cell a merged cell's paragraphs are still handed out to its rows.)
        NestedTables =
            cell.NestedTables
            |> List.map (tableBlock resolver images styleMap tableStyleMap comments notes)
            |> List.map (fun table -> { table with Rows = WordLayout.Table.distributeVerticalMergeContent table.Rows })
        NestedTableAnchors = cell.NestedTableAnchors
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

and private tableBlock resolver images styleMap tableStyleMap comments notes (table: DocxReader.Table) =
    let styleProperties = resolveTableStyleProperties tableStyleMap table.StyleId
    let tableRunProperties =
        resolveTableRunProperties (fun style -> style.RunProperties) tableStyleMap table.StyleId
    let tableParagraphProperties =
        resolveTableParagraphProperties (fun style -> style.ParagraphProperties) tableStyleMap table.StyleId
        |> paragraphPropertiesWithRunProperties tableRunProperties
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
    let rows =
        table.Rows
        |> List.mapi (fun rowIndex row -> {
            Cells = row.Cells |> List.map (tableCell resolver images styleMap tableStyleMap comments notes (paragraphPropertiesForRow rowIndex))
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
    {
        Id = Guid.NewGuid().ToString("N")
        Rows = rows
        PreferredWidth = tableWidth table.Width table.WidthType
        Alignment = table.Alignment
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
        FloatingXml = table.FloatingXml
        StyleId = table.StyleId
        Look =
            table.Look
            |> Option.map (fun look ->
                ({ FirstRow = look.FirstRow; LastRow = look.LastRow; FirstColumn = look.FirstColumn; LastColumn = look.LastColumn
                   HorizontalBanding = look.HorizontalBanding; VerticalBanding = look.VerticalBanding }: Rbt.Word.Document.TableLook))
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
        IndentLeft = para |> Option.bind (fun p -> twipsToCanvasPixels p.IndentLeft)
        IndentRight = para |> Option.bind (fun p -> twipsToCanvasPixels p.IndentRight)
        FirstLineIndent =
            para
            |> Option.bind (fun p ->
                twipsToCanvasPixels p.FirstLineIndent
                |> Option.orElse (twipsToCanvasPixels p.HangingIndent |> Option.map (fun hanging -> -hanging)))
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
            Columns = setup.Columns |> Option.map (fun c -> ({ Count = c.Count; Gap = c.Gap; Separator = c.Separator; Widths = c.Widths }: Rbt.Word.Document.SectionColumns))
            SectionStart = setup.SectionStart
            HeaderReferences = setup.HeaderReferences |> List.map (fun ref -> ref.Kind, ref.RelationshipId)
            FooterReferences = setup.FooterReferences |> List.map (fun ref -> ref.Kind, ref.RelationshipId)
            PageNumberStart = setup.PageNumberStart
            PageNumberFormat = setup.PageNumberFormat
        }

/// A header / footer part with its blocks in order. It used to be built from
/// HeaderFooter.Content, the part's paragraphs with every table flattened, so
/// a header laid out with a table (logo left, address right) came out as a
/// column of loose paragraphs.
let private headerFooterPart resolver images styleMap tableStyleMap comments notes (item: DocxReader.HeaderFooter) =
    {
        Kind = item.Type
        RelationshipId = item.RelationshipId
        Blocks =
            item.Blocks
            |> List.map (function
                | DocxReader.Para para -> Paragraph (mkPara resolver images styleMap comments notes true false para)
                | DocxReader.Tbl table -> Table (tableBlock resolver images styleMap tableStyleMap comments notes table))
    }

let importDoc (doc: DocxReader.WordDocument) =
    RawStyles.Current <- doc.Styles |> List.map (fun style -> style.Id, style) |> Map.ofList
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

    // Body-style blocks: paragraphs and tables in order. Used for the body and
    // for the bodies of footnotes, endnotes and comments.
    let contentBlocks (content: DocxReader.DocumentContent list) =
        let documentDefaultMissingSpacingToZero = documentDefaultParagraphResetsMissingSpacing styleMap
        content
        |> List.map (function
            | DocxReader.Para para ->
                let defaultMissingSpacingToZero =
                    documentDefaultMissingSpacingToZero
                    || directParagraphFormattingResetsMissingSpacing para.Properties
                Paragraph (mkPara resolver doc.Images styleMap comments notes true defaultMissingSpacingToZero para)
            | DocxReader.Tbl table -> Table (tableBlock resolver doc.Images styleMap tableStyleMap comments notes table))

    let blocks = contentBlocks doc.Content

    let noteBodies noteType (items: DocxReader.Note list) : NoteBody list =
        items
        |> List.map (fun note ->
            { NoteType = noteType
              Id = note.Id
              Kind = note.Kind
              HasMark = note.HasMark
              MarkStyle = note.MarkStyle
              Text = notes |> Map.tryFind (noteType, note.Id) |> Option.defaultValue ""
              Blocks = contentBlocks note.Blocks })

    let commentBodies : CommentBody list =
        doc.Comments
        |> List.map (fun comment ->
            { Id = comment.Id
              Author = comment.Author
              Date = comment.Date
              Initials = comment.Initials
              Text = comment.Text
              Blocks = contentBlocks comment.Blocks })

    // Drawing shapes / autoshapes -> floating TextBox blocks (preset/custom
    // geometry + gradient), mirroring the Fable bridge in Office.Formats Main.
    let shapeBlocks =
        doc.Shapes
        |> List.filter (fun s -> not s.Content.IsEmpty || s.Preset.IsSome || not s.CustomPath.IsEmpty)
        |> List.mapi (fun i (shape: DocxReader.Shape) ->
            let content = shape.Content |> List.map (mkPara resolver doc.Images styleMap comments notes true false)
            let geometry =
                if not shape.CustomPath.IsEmpty then
                    let segs =
                        shape.CustomPath
                        |> List.choose (fun seg ->
                            let p k = seg.Pts |> List.tryItem k |> Option.defaultValue 0.0
                            match seg.Op with
                            | "M" -> Some (PMove(p 0, p 1))
                            | "L" -> Some (PLine(p 0, p 1))
                            | "C" -> Some (PCubic(p 0, p 1, p 2, p 3, p 4, p 5))
                            | "Q" -> Some (PQuad(p 0, p 1, p 2, p 3))
                            | "Z" -> Some PClose
                            | _ -> None)
                    GeoCustom(segs, shape.CustomW |> Option.defaultValue 0.0, shape.CustomH |> Option.defaultValue 0.0)
                else
                    match shape.Preset with
                    | Some preset -> GeoPreset preset
                    | None -> GeoRect
            let gradient =
                shape.Gradient
                |> Option.map (fun g ->
                    { Stops = g.Stops |> List.map (fun s -> { Offset = s.Offset; Color = s.Color }); Angle = g.Angle; Radial = g.Radial })
            let anchorId =
                shape.AnchorBlock
                |> Option.bind (fun index -> List.tryItem index blocks)
                |> Option.bind (function
                    | Paragraph para
                    | Heading (_, para) -> Some para.Id
                    | Table table -> Some table.Id
                    | TextBox _ -> None)
            TextBox {
                defaultTextBox with
                    Id = sprintf "shape-%d" i
                    Content = content
                    X = shape.Position |> Option.map fst
                    Y = shape.Position |> Option.map snd
                    PositionHRelativeFrom = shape.PositionHRelativeFrom
                    PositionVRelativeFrom = shape.PositionVRelativeFrom
                    Width = shape.Width |> Option.defaultValue 220.0
                    Height = shape.Height |> Option.defaultValue 72.0
                    FillColor = if shape.NoFill then None else shape.FillColor |> Option.orElse (Some "#FFFFFF")
                    BorderColor = if shape.NoLine then None else shape.BorderColor |> Option.orElse (Some "#808080")
                    BorderWidth = shape.BorderWidth |> Option.orElse (Some 1.0)
                    RelativeHeight = shape.RelativeHeight
                    Geometry = geometry
                    Gradient = gradient
                    Rotation = shape.Rotation
                    Inline = shape.Inline
                    Wrap = shape.Wrap
                    PositionHAlign = shape.PositionHAlign
                    PositionVAlign = shape.PositionVAlign
                    BehindDoc = shape.BehindDoc
                    AllowOverlap = shape.AllowOverlap
                    LayoutInCell = shape.LayoutInCell
                    DistanceTop = shape.DistanceTop
                    DistanceBottom = shape.DistanceBottom
                    DistanceLeft = shape.DistanceLeft
                    DistanceRight = shape.DistanceRight
                    BorderDash = shape.LineDash
                    Name = shape.Name
                    AnchorBlockId = anchorId
            })

    let blocks = blocks @ shapeBlocks

    let firstParagraph =
        blocks
        |> List.tryPick (function
            | Paragraph para
            | Heading (_, para) -> Some para
            | Table table -> table.Rows |> List.tryPick (fun row -> row.Cells |> List.tryPick (fun cell -> cell.Content |> List.tryHead))
            | TextBox textBox -> textBox.Content |> List.tryHead)

    let styles = doc.Styles |> List.map (namedStyle doc.Theme)
    let headerParts = doc.Headers |> List.map (headerFooterPart resolver doc.Images styleMap tableStyleMap comments notes)
    let footerParts = doc.Footers |> List.map (headerFooterPart resolver doc.Images styleMap tableStyleMap comments notes)

    {
        Blocks = blocks
        Selection = firstParagraph |> Option.map (fun para -> { Anchor = { BlockId = para.Id; Offset = 0 }; Focus = { BlockId = para.Id; Offset = 0 } })
        PageSetup = pageSetup doc
        Styles = styles
        Headers = headerParts |> List.tryHead |> Option.map (fun part -> part.Blocks |> List.choose (function | Paragraph p -> Some p | Heading (_, p) -> Some p | _ -> None)) |> Option.defaultValue []
        Footers = footerParts |> List.tryHead |> Option.map (fun part -> part.Blocks |> List.choose (function | Paragraph p -> Some p | Heading (_, p) -> Some p | _ -> None)) |> Option.defaultValue []
        HeaderParts = headerParts
        FooterParts = footerParts
        Package =
            { emptyPackageData with
                Notes = noteBodies "footnote" doc.Footnotes @ noteBodies "endnote" doc.Endnotes
                Comments = commentBodies
                Preserved = box doc.Preserved
                ImportedStyles = styles }
    }
