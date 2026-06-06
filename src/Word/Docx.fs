module RedBat.Word.Docx

open Fable.Core
open Fable.Core.JsInterop
open RedBat.Word.Document
open RedBat.Word.TextBuffer
open RedBat.Word.Rope

/// Block ID counter for imported documents
let mutable private importedBlockCounter = 1000

/// Generate next imported block ID
let private nextImportedBlockId () =
    importedBlockCounter <- importedBlockCounter + 1
    sprintf "imported-block-%d" importedBlockCounter

/// Get JS array element (avoid F# Array.item for JS arrays)
[<Emit("$0[$1]")>]
let private jsArrayGet (arr: obj) (index: int) : obj = jsNative

[<Emit("$0 === null || $0 === undefined")>]
let private isNullOrUndefined (_value: obj) : bool = jsNative

[<Emit("$0.length")>]
let private jsArrayLength (_arr: obj) : int = jsNative

let private optionalString (value: obj) =
    if isNullOrUndefined value then None
    else
        match unbox<string> value with
        | null | "" -> None
        | s -> Some s

let private optionalFloat (value: obj) =
    if isNullOrUndefined value then None
    else Some (unbox<float> value)

let private optionalBool (value: obj) =
    if isNullOrUndefined value then None
    else Some (unbox<bool> value)

let private optionalInt (value: obj) =
    if isNullOrUndefined value then None
    else Some (unbox<int> value)

let private optionalIntAsFloat (value: obj) =
    if isNullOrUndefined value then None
    else Some (float (unbox<int> value))

let private optionalDxaAsCanvasPixels (value: obj) =
    optionalIntAsFloat value |> Option.map (fun dxa -> dxa / WordLayout.canvasScale.DxaPerUnit)

let private optionalDxaAsPoints (value: obj) =
    optionalIntAsFloat value |> Option.map (fun dxa -> dxa / WordLayout.printScale.DxaPerUnit)

let private optionalNumberAsFloat (value: obj) =
    if isNullOrUndefined value then None
    else
        try Some (unbox<float> value)
        with _ ->
            try Some (float (unbox<int> value))
            with _ ->
                match optionalString value with
                | Some text ->
                    match System.Double.TryParse(text) with
                    | true, parsed -> Some parsed
                    | false, _ -> None
                | None -> None

let private boolOrDefault defaultValue (value: obj) =
    if isNullOrUndefined value then defaultValue
    else unbox<bool> value

let private floatOrDefault defaultValue (value: obj) =
    if isNullOrUndefined value then defaultValue
    else unbox<float> value

let private stringOrDefault defaultValue (value: obj) =
    if isNullOrUndefined value then defaultValue
    else
        match unbox<string> value with
        | null | "" -> defaultValue
        | s -> s

let private intOrDefault defaultValue (value: obj) =
    if isNullOrUndefined value then defaultValue
    else unbox<int> value

// --- Shape geometry bridge (Office.Formats emits these on a TextBox/shape item) ---

/// Reads a CustomPath array of { Op: "M"|"L"|"C"|"Q"|"Z", P: number[] } segments.
let private pathSegsFromEditor (value: obj) : PathSeg list =
    if isNullOrUndefined value then []
    else
        [ for i in 0 .. jsArrayLength value - 1 ->
            let seg = jsArrayGet value i
            let op = stringOrDefault "" (seg?Op)
            let p = seg?P
            let n k = if isNullOrUndefined p then 0.0 else floatOrDefault 0.0 (jsArrayGet p k)
            match op with
            | "M" -> Some (PMove (n 0, n 1))
            | "L" -> Some (PLine (n 0, n 1))
            | "C" -> Some (PCubic (n 0, n 1, n 2, n 3, n 4, n 5))
            | "Q" -> Some (PQuad (n 0, n 1, n 2, n 3))
            | "Z" -> Some PClose
            | _ -> None ]
        |> List.choose id

let private geometryFromEditor (item: obj) : ShapeGeometry =
    match optionalString (item?Geometry) with
    | Some "custom" ->
        let segs = pathSegsFromEditor (item?CustomPath)
        let cw = optionalNumberAsFloat (item?CustomW) |> Option.defaultValue 0.0
        let ch = optionalNumberAsFloat (item?CustomH) |> Option.defaultValue 0.0
        if segs.IsEmpty || cw <= 0.0 || ch <= 0.0 then GeoRect else GeoCustom (segs, cw, ch)
    | Some "rect" | Some "" | None -> GeoRect
    | Some preset -> GeoPreset preset

let private gradientFromEditor (value: obj) : GradientSpec option =
    if isNullOrUndefined value then None
    else
        let stopsRaw = value?Stops
        let stops =
            if isNullOrUndefined stopsRaw then []
            else
                [ for i in 0 .. jsArrayLength stopsRaw - 1 ->
                    let s = jsArrayGet stopsRaw i
                    { Offset = floatOrDefault 0.0 (s?Offset); Color = stringOrDefault "#000000" (s?Color) } ]
        if stops.IsEmpty then None
        else Some { Stops = stops; Angle = floatOrDefault 0.0 (value?Angle); Radial = boolOrDefault false (value?Radial) }

let private sectionPageSetupFromEditor (value: obj) : SectionPageSetup option =
    if isNullOrUndefined value then None
    else
        Some ({
            PageWidth = optionalFloat (value?PageWidth) |> Option.defaultValue defaultPageSetup.PageWidth
            PageHeight = optionalFloat (value?PageHeight) |> Option.defaultValue defaultPageSetup.PageHeight
            MarginLeft = optionalFloat (value?MarginLeft) |> Option.defaultValue defaultPageSetup.MarginLeft
            MarginRight = optionalFloat (value?MarginRight) |> Option.defaultValue defaultPageSetup.MarginRight
            MarginTop = optionalFloat (value?MarginTop) |> Option.defaultValue defaultPageSetup.MarginTop
            MarginBottom = optionalFloat (value?MarginBottom) |> Option.defaultValue defaultPageSetup.MarginBottom
            HeaderDistance = optionalFloat (value?HeaderDistance) |> Option.defaultValue defaultPageSetup.HeaderDistance
            FooterDistance = optionalFloat (value?FooterDistance) |> Option.defaultValue defaultPageSetup.FooterDistance
            DifferentFirstPage = value?DifferentFirstPage |> boolOrDefault defaultPageSetup.DifferentFirstPage
            EvenAndOddHeaders = value?EvenAndOddHeaders |> boolOrDefault defaultPageSetup.EvenAndOddHeaders
        }: SectionPageSetup)

let private sectionPageSetupToEditor (setup: SectionPageSetup) =
    createObj [
        "PageWidth" ==> setup.PageWidth
        "PageHeight" ==> setup.PageHeight
        "MarginLeft" ==> setup.MarginLeft
        "MarginRight" ==> setup.MarginRight
        "MarginTop" ==> setup.MarginTop
        "MarginBottom" ==> setup.MarginBottom
        "HeaderDistance" ==> setup.HeaderDistance
        "FooterDistance" ==> setup.FooterDistance
        "DifferentFirstPage" ==> setup.DifferentFirstPage
        "EvenAndOddHeaders" ==> setup.EvenAndOddHeaders
    ]

let private shadingFromEditor (value: obj) : Shading option =
    if isNullOrUndefined value then None
    else
        let shading = {
            Pattern = optionalString (value?Pattern)
            ForegroundColor = optionalString (value?ForegroundColor)
            BackgroundColor = optionalString (value?BackgroundColor)
        }
        if shading.Pattern.IsNone && shading.ForegroundColor.IsNone && shading.BackgroundColor.IsNone then None
        else Some shading

let private shadingFromEditorFields (value: obj) =
    shadingFromEditor (value?Shading)
    |> Option.orElseWith (fun () ->
        let shading = {
            Pattern = optionalString (value?ShadingPattern)
            ForegroundColor = optionalString (value?ShadingForegroundColor)
            BackgroundColor = optionalString (value?ShadingBackgroundColor)
        }
        if shading.Pattern.IsNone && shading.ForegroundColor.IsNone && shading.BackgroundColor.IsNone then None
        else Some shading)

let private shadingToEditor (shading: Shading option) =
    shading
    |> Option.map (fun value ->
        createObj [
            "Pattern" ==> value.Pattern
            "ForegroundColor" ==> value.ForegroundColor
            "BackgroundColor" ==> value.BackgroundColor
        ])

/// Convert alignment string to Alignment type
let private alignmentFromString (align: string) : Alignment =
    match align.ToLower() with
    | "center" -> Center
    | "right" -> Right
    | "justify" -> Justify
    | _ -> Left

/// Convert Alignment type to string
let private alignmentToString (align: Alignment) : string =
    match align with
    | Left -> "left"
    | Center -> "center"
    | Right -> "right"
    | Justify -> "justify"

let private listStyleFromString (style: string) =
    match style.ToLower() with
    | "bullet" -> Bullet
    | "numbered" -> Numbered
    | _ -> NoList

let private listStyleToString style =
    match style with
    | Bullet -> "bullet"
    | Numbered -> "numbered"
    | NoList -> "none"

let private tableWidthKindFromString = function
    | Some "dxa" -> DxaWidth
    | Some "pct" -> PercentWidth
    | Some "auto" -> AutoWidth
    | Some "nil" -> AutoWidth
    | _ -> PointWidth

let private tableWidthKindToString = function
    | DxaWidth -> "dxa"
    | PercentWidth -> "pct"
    | AutoWidth -> "auto"
    | PointWidth -> "pt"

let private tableWidthFromEditor value typeValue =
    match optionalNumberAsFloat value with
    | Some amount ->
        let kind = tableWidthKindFromString (optionalString typeValue)
        let normalizedAmount =
            match kind with
            | PercentWidth when amount > 100.0 -> amount / 50.0
            | _ -> amount
        Some { Kind = kind; Value = Some normalizedAmount }
    | None ->
        match optionalString typeValue with
        | Some "auto" | Some "nil" -> Some { Kind = AutoWidth; Value = None }
        | _ -> None

let private tableWidthToEditor (width: TableWidth option) =
    width |> Option.map (fun value -> value.Value |> Option.defaultValue 0.0)

let private tableWidthTypeToEditor (width: TableWidth option) =
    width |> Option.map (fun value -> tableWidthKindToString value.Kind)

let private tableLayoutModeFromString = function
    | Some "fixed" -> Some FixedTableLayout
    | Some "autofit" -> Some AutoFitTableLayout
    | Some "autoFit" -> Some AutoFitTableLayout
    | _ -> None

let private tableLayoutModeToString = function
    | Some FixedTableLayout -> Some "fixed"
    | Some AutoFitTableLayout -> Some "autofit"
    | None -> None

let private tableVerticalMergeFromString = function
    | Some "restart" -> Some RestartMerge
    | Some "continue" -> Some ContinueMerge
    | _ -> None

let private tableVerticalMergeToString = function
    | Some RestartMerge -> Some "restart"
    | Some ContinueMerge -> Some "continue"
    | None -> None

let private tableBorderFromEditor (value: obj) =
    if isNullOrUndefined value then None
    else
        Some {
            Style = optionalString (value?Style)
            Color = optionalString (value?Color)
            Width = optionalNumberAsFloat (value?Width)
        }

let private tableBordersFromEditor (value: obj) =
    if isNullOrUndefined value then None
    else
        Some {
            Top = tableBorderFromEditor (value?Top)
            Right = tableBorderFromEditor (value?Right)
            Bottom = tableBorderFromEditor (value?Bottom)
            Left = tableBorderFromEditor (value?Left)
            InsideH = tableBorderFromEditor (value?InsideH)
            InsideV = tableBorderFromEditor (value?InsideV)
        }

let private tableBorderToEditor (border: TableBorder option) =
    border |> Option.map (fun value -> createObj [
        "Style" ==> value.Style
        "Color" ==> value.Color
        "Width" ==> value.Width
    ])

let private tableBordersToEditor (borders: TableBorders option) =
    borders |> Option.map (fun value -> createObj [
        "Top" ==> tableBorderToEditor value.Top
        "Right" ==> tableBorderToEditor value.Right
        "Bottom" ==> tableBorderToEditor value.Bottom
        "Left" ==> tableBorderToEditor value.Left
        "InsideH" ==> tableBorderToEditor value.InsideH
        "InsideV" ==> tableBorderToEditor value.InsideV
    ])

let private paragraphBorderFromEditor (value: obj) =
    if isNullOrUndefined value then None
    else
        Some {
            Style = optionalString (value?Style)
            Color = optionalString (value?Color)
            Width = optionalNumberAsFloat (value?Width)
            Space = optionalNumberAsFloat (value?Space)
        }

let private paragraphBordersFromEditor (value: obj) =
    if isNullOrUndefined value then None
    else
        Some {
            Top = paragraphBorderFromEditor (value?Top)
            Right = paragraphBorderFromEditor (value?Right)
            Bottom = paragraphBorderFromEditor (value?Bottom)
            Left = paragraphBorderFromEditor (value?Left)
        }

let private paragraphBorderToEditor (border: ParagraphBorder option) =
    border |> Option.map (fun value -> createObj [
        "Style" ==> value.Style
        "Color" ==> value.Color
        "Width" ==> value.Width
        "Space" ==> value.Space
    ])

let private paragraphBordersToEditor (borders: ParagraphBorders option) =
    borders |> Option.map (fun value -> createObj [
        "Top" ==> paragraphBorderToEditor value.Top
        "Right" ==> paragraphBorderToEditor value.Right
        "Bottom" ==> paragraphBorderToEditor value.Bottom
        "Left" ==> paragraphBorderToEditor value.Left
    ])

let private refsFromEditor (value: obj) =
    if isNullOrUndefined value then []
    else
        value
        |> unbox<obj array>
        |> Array.choose (fun item ->
            let kind = item?Kind |> stringOrDefault "default"
            optionalString (item?RelationshipId) |> Option.map (fun id -> (kind, id)))
        |> Array.toList

let private refsToEditor refs =
    refs
    |> List.map (fun (kind, relationshipId) -> createObj [ "Kind" ==> kind; "RelationshipId" ==> relationshipId ])
    |> List.toArray

let private metadataFromEditor (item: obj) =
    let alignment = item?Alignment |> stringOrDefault "left" |> alignmentFromString
    let style = optionalString (item?Style)
    let contentControlId = optionalString (item?ContentControlId)
    let contentControlTag = optionalString (item?ContentControlTag)
    let contentControlAlias = optionalString (item?ContentControlAlias)
    let listStyle = item?ListStyle |> stringOrDefault "none" |> listStyleFromString
    let listLevel = item?ListLevel |> intOrDefault 0
    let listId = optionalString (item?ListId)
    let listNumberFormat = optionalString (item?ListNumberFormat)
    let listLevelText = optionalString (item?ListLevelText)
    let listStart = if isNullOrUndefined (item?ListStart) then None else Some (unbox<int> item?ListStart)
    let indentLeft = optionalFloat (item?IndentLeft)
    let indentRight = optionalFloat (item?IndentRight)
    let firstLineIndent = optionalFloat (item?FirstLineIndent)
    let rightTabPosition = optionalFloat (item?RightTabPosition)
    let rightTabLeader = optionalString (item?RightTabLeader)
    let tabStops : TabStop list =
        let arr = item?TabStops
        if isNullOrUndefined arr then []
        else
            [ for i in 0 .. jsArrayLength arr - 1 do
                let t = jsArrayGet arr i
                yield { Position = optionalFloat (t?Position) |> Option.defaultValue 0.0
                        Alignment = optionalString (t?Alignment) |> Option.defaultValue "left"
                        Leader = optionalString (t?Leader) } ]
    let rtl = item?Rtl |> boolOrDefault false
    let spacingBefore = optionalFloat (item?SpacingBefore)
    let spacingAfter = optionalFloat (item?SpacingAfter)
    let paragraphMarkFontSize = optionalFloat (item?ParagraphMarkFontSize)
    let lineSpacing = optionalFloat (item?LineSpacing)
    let lineSpacingRule = optionalString (item?LineSpacingRule)
    let widowControl = optionalBool (item?StyleWidowControl) |> Option.orElse (optionalBool (item?WidowControl))
    let backgroundColor = optionalString (item?BackgroundColor)
    let paragraphBorders = paragraphBordersFromEditor (item?ParagraphBorders)
    let paragraphBorderBefore = optionalFloat (item?ParagraphBorderBefore)
    let paragraphBorderAfter = optionalFloat (item?ParagraphBorderAfter)
    let pageBreakBefore = item?PageBreakBefore |> boolOrDefault false
    let pageBreakAfter = item?PageBreakAfter |> boolOrDefault false
    let sectionBreak = optionalString (item?SectionBreak)
    let sectionPageNumberStart = optionalInt (item?SectionPageNumberStart)
    let renderedPageBreakBefore = optionalInt (item?RenderedPageBreakBefore)
    let sectionPageSetup = sectionPageSetupFromEditor (item?SectionPageSetup)
    let headerReferences = refsFromEditor (item?HeaderReferences)
    let footerReferences = refsFromEditor (item?FooterReferences)
    {
        defaultMetadata with
            Style = style
            ContentControlId = contentControlId
            ContentControlTag = contentControlTag
            ContentControlAlias = contentControlAlias
            Alignment = alignment
            ListStyle = listStyle
            ListLevel = if listStyle = NoList then 0 else max 1 listLevel
            ListId = listId
            ListNumberFormat = listNumberFormat
            ListLevelText = listLevelText
            ListStart = listStart
            IndentLeft = indentLeft
            IndentRight = indentRight
            FirstLineIndent = firstLineIndent
            RightTabPosition = rightTabPosition
            RightTabLeader = rightTabLeader
            TabStops = tabStops
            Rtl = rtl
            SpacingBefore = spacingBefore
            SpacingAfter = spacingAfter
            ParagraphMarkFontSize = paragraphMarkFontSize
            ContextualSpacing = optionalBool (item?StyleContextualSpacing)
            LineSpacing = lineSpacing
            LineSpacingRule = lineSpacingRule
            WidowControl = widowControl
            BackgroundColor = backgroundColor
            ParagraphBorders = paragraphBorders
            ParagraphBorderBefore = paragraphBorderBefore
            ParagraphBorderAfter = paragraphBorderAfter
            PageBreakBefore = pageBreakBefore
            PageBreakAfter = pageBreakAfter
            SectionBreak = sectionBreak
            SectionPageNumberStart = sectionPageNumberStart
            RenderedPageBreakBefore = renderedPageBreakBefore
            SectionPageSetup = sectionPageSetup
            HeaderReferences = headerReferences
            FooterReferences = footerReferences
    }

/// Convert EditorRun array to TextBuffer
let private editorRunsToBuffer (runs: obj array) : TextBuffer =
    // Build full text and create runs
    let mutable currentOffset = 0
    let mutable textParts = []
    let mutable bufferRuns = []

    for i in 0 .. runs.Length - 1 do
        let run = jsArrayGet runs i  // Use JS array access
        let isImage = run?IsImage |> boolOrDefault false

        if isImage then
            // Image run
            let src = run?ImageSrc |> stringOrDefault ""
            let linked = run?ImageLinked |> boolOrDefault false
            let width = optionalFloat (run?ImageWidth)
            let height = optionalFloat (run?ImageHeight)
            let alt = optionalString (run?ImageAlt)
            let name = optionalString (run?ImageName)
            let title = optionalString (run?ImageTitle)
            let floating = run?ImageFloating |> boolOrDefault false
            let wrap = optionalString (run?ImageWrap)
            let positionX = optionalFloat (run?ImagePositionX)
            let positionY = optionalFloat (run?ImagePositionY)
            let positionHRelativeFrom = optionalString (run?ImagePositionHRelativeFrom)
            let positionVRelativeFrom = optionalString (run?ImagePositionVRelativeFrom)
            let behindDoc = optionalBool (run?ImageBehindDoc)
            let allowOverlap = optionalBool (run?ImageAllowOverlap)
            let layoutInCell = optionalBool (run?ImageLayoutInCell)
            let distanceTop = optionalFloat (run?ImageDistanceTop)
            let distanceBottom = optionalFloat (run?ImageDistanceBottom)
            let distanceLeft = optionalFloat (run?ImageDistanceLeft)
            let distanceRight = optionalFloat (run?ImageDistanceRight)
            let relativeHeight = optionalString (run?ImageRelativeHeight)
            let locked = optionalBool (run?ImageLocked)
            let simplePosition = optionalBool (run?ImageSimplePosition)
            let simplePositionX = optionalFloat (run?ImageSimplePositionX)
            let simplePositionY = optionalFloat (run?ImageSimplePositionY)
            let cropLeft = optionalString (run?ImageCropLeft)
            let cropTop = optionalString (run?ImageCropTop)
            let cropRight = optionalString (run?ImageCropRight)
            let cropBottom = optionalString (run?ImageCropBottom)

            // Add placeholder character to text
            textParts <- textParts @ [TextBuffer.objectReplacementChar]

            // Create image run
            bufferRuns <- bufferRuns @ [Image {
                Start = currentOffset
                Src = src
                Linked = linked
                Width = width
                Height = height
                Alt = alt
                Name = name
                Title = title
                Floating = floating
                Wrap = wrap
                PositionX = positionX
                PositionY = positionY
                PositionHRelativeFrom = positionHRelativeFrom
                PositionVRelativeFrom = positionVRelativeFrom
                BehindDoc = behindDoc
                AllowOverlap = allowOverlap
                LayoutInCell = layoutInCell
                DistanceTop = distanceTop
                DistanceBottom = distanceBottom
                DistanceLeft = distanceLeft
                DistanceRight = distanceRight
                RelativeHeight = relativeHeight
                Locked = locked
                SimplePosition = simplePosition
                SimplePositionX = simplePositionX
                SimplePositionY = simplePositionY
                CropLeft = cropLeft
                CropTop = cropTop
                CropRight = cropRight
                CropBottom = cropBottom
            }]

            currentOffset <- currentOffset + 1
        else
            // Text run
            let text = run?Text |> unbox<string>

            // Don't skip whitespace-only runs - they're important for spacing!
            if not (System.String.IsNullOrEmpty(text)) then
                let bold = run?Bold |> boolOrDefault false
                let italic = run?Italic |> boolOrDefault false
                let underline = run?Underline |> boolOrDefault false
                let underlineStyle = optionalString (run?UnderlineStyle)
                let underlineColor = optionalString (run?UnderlineColor)
                let allCaps = run?AllCaps |> boolOrDefault false
                let smallCaps = run?SmallCaps |> boolOrDefault false
                let strike = run?Strike |> boolOrDefault false
                let doubleStrike = run?DoubleStrike |> boolOrDefault false
                let hidden = run?Hidden |> boolOrDefault false
                let position = optionalFloat (run?Position) |> Option.defaultValue 0.0
                let emphasisMark = optionalString (run?EmphasisMark)
                let fontSize = run?FontSize |> floatOrDefault TextBuffer.defaultStyle.FontSize
                let characterSpacing = optionalFloat (run?CharacterSpacing) |> Option.defaultValue TextBuffer.defaultStyle.CharacterSpacing
                let fontFamily = optionalString (run?FontFamily)
                let color = optionalString (run?Color)
                let highlight = optionalString (run?Highlight)
                let hyperlink = optionalString (run?Hyperlink)
                let commentId = optionalString (run?CommentId)
                let commentText = optionalString (run?CommentText)
                let commentAuthor = optionalString (run?CommentAuthor)
                let commentDate = optionalString (run?CommentDate)
                let noteId = optionalString (run?NoteId)
                let noteType = optionalString (run?NoteType)
                let noteText = optionalString (run?NoteText)
                let bookmarkStartId = optionalString (run?BookmarkStartId)
                let bookmarkName = optionalString (run?BookmarkName)
                let bookmarkEndId = optionalString (run?BookmarkEndId)
                let fieldCode = optionalString (run?FieldCode)
                let fieldResult = optionalString (run?FieldResult)
                let renderedPageBreakBefore = run?RenderedPageBreakBefore |> boolOrDefault false
                let equationXml = optionalString (run?EquationXml)
                let revisionId = optionalString (run?RevisionId)
                let revisionType = optionalString (run?RevisionType)
                let revisionAuthor = optionalString (run?RevisionAuthor)
                let revisionDate = optionalString (run?RevisionDate)
                let vertAlign = optionalString (run?VertAlign)
                let styleId = optionalString (run?Style)

                let style = {
                    Style = styleId
                    Bold = bold
                    Italic = italic
                    Underline = underline || hyperlink.IsSome
                    DoubleStrike = doubleStrike
                    Hidden = hidden
                    Position = position
                    EmphasisMark = emphasisMark
                    UnderlineStyle = underlineStyle
                    UnderlineColor = underlineColor
                    AllCaps = allCaps
                    SmallCaps = smallCaps
                    Strike = strike
                    FontSize = fontSize
                    CharacterSpacing = characterSpacing
                    FontFamily = fontFamily
                    Color = color
                    Highlight = highlight
                    Hyperlink = hyperlink
                    CommentId = commentId
                    CommentText = commentText
                    CommentAuthor = commentAuthor
                    CommentDate = commentDate
                    NoteId = noteId
                    NoteType = noteType
                    NoteText = noteText
                    BookmarkStartId = bookmarkStartId
                    BookmarkName = bookmarkName
                    BookmarkEndId = bookmarkEndId
                    FieldCode = fieldCode
                    FieldResult = fieldResult
                    RenderedPageBreakBefore = renderedPageBreakBefore
                    EquationXml = equationXml
                    RevisionId = revisionId
                    RevisionType = revisionType
                    RevisionAuthor = revisionAuthor
                    RevisionDate = revisionDate
                    VertAlign = vertAlign
                }

                textParts <- textParts @ [text]
                if style.RenderedPageBreakBefore && text.Length > 1 then
                    bufferRuns <- bufferRuns @ [
                        Text { Start = currentOffset; Length = 1; Style = style }
                        Text { Start = currentOffset + 1; Length = text.Length - 1; Style = { style with RenderedPageBreakBefore = false } }
                    ]
                else
                    bufferRuns <- bufferRuns @ [Text {
                        Start = currentOffset
                        Length = text.Length
                        Style = style
                    }]

                currentOffset <- currentOffset + text.Length

    let fullText = String.concat "" textParts
    let rope = Rope.ofString fullText

    { Rope = rope; Runs = bufferRuns }

let private paragraphFromEditor (item: obj) =
    let runs = item?Runs |> unbox<obj array>
    {
        Id = nextImportedBlockId ()
        Buffer = editorRunsToBuffer runs
        Metadata = metadataFromEditor item
    }

let private styleTextFromEditor (value: obj) =
    let hasAny =
        not (isNullOrUndefined value?StyleBold)
        || not (isNullOrUndefined value?StyleItalic)
        || not (isNullOrUndefined value?StyleUnderline)
        || not (isNullOrUndefined value?StyleAllCaps)
        || not (isNullOrUndefined value?StyleSmallCaps)
        || not (isNullOrUndefined value?StyleStrike)
        || not (isNullOrUndefined value?StyleVertAlign)
        || not (isNullOrUndefined value?StyleFontSize)
        || not (isNullOrUndefined value?StyleCharacterSpacing)
        || not (isNullOrUndefined value?StyleFontFamily)
        || not (isNullOrUndefined value?StyleColor)
        || not (isNullOrUndefined value?StyleHighlight)
    if not hasAny then None
    else
        Some {
            TextBuffer.defaultStyle with
                Bold = boolOrDefault false (value?StyleBold)
                Italic = boolOrDefault false (value?StyleItalic)
                Underline = boolOrDefault false (value?StyleUnderline)
                UnderlineStyle = optionalString (value?StyleUnderlineStyle)
                UnderlineColor = optionalString (value?StyleUnderlineColor)
                AllCaps = boolOrDefault false (value?StyleAllCaps)
                SmallCaps = boolOrDefault false (value?StyleSmallCaps)
                Strike = boolOrDefault false (value?StyleStrike)
                DoubleStrike = boolOrDefault false (value?StyleDoubleStrike)
                Hidden = boolOrDefault false (value?StyleHidden)
                Position = optionalFloat (value?StylePosition) |> Option.defaultValue 0.0
                EmphasisMark = optionalString (value?StyleEmphasisMark)
                VertAlign = optionalString (value?StyleVertAlign)
                FontSize = optionalFloat (value?StyleFontSize) |> Option.defaultValue TextBuffer.defaultStyle.FontSize
                CharacterSpacing = optionalFloat (value?StyleCharacterSpacing) |> Option.defaultValue TextBuffer.defaultStyle.CharacterSpacing
                FontFamily = optionalString (value?StyleFontFamily)
                Color = optionalString (value?StyleColor)
                Highlight = optionalString (value?StyleHighlight)
        }

let private styleTextBoldFromEditor (value: obj) = optionalBool (value?StyleBold)
let private styleTextItalicFromEditor (value: obj) = optionalBool (value?StyleItalic)
let private styleTextUnderlineFromEditor (value: obj) = optionalBool (value?StyleUnderline)
let private styleTextAllCapsFromEditor (value: obj) = optionalBool (value?StyleAllCaps)
let private styleTextSmallCapsFromEditor (value: obj) = optionalBool (value?StyleSmallCaps)
let private styleTextStrikeFromEditor (value: obj) = optionalBool (value?StyleStrike)

let private paragraphStyleFromEditor (item: obj) =
    optionalString (item?Style)
    |> Option.map (fun styleId -> {
        Id = styleId
        Name = optionalString (item?StyleName)
        StyleType = "paragraph"
        IsDefault = optionalBool (item?IsDefaultStyle) |> Option.orElse (optionalBool (item?IsDefault)) |> Option.defaultValue false
        BasedOn = optionalString (item?BasedOnStyle)
        Next = optionalString (item?NextStyle)
        Linked = optionalString (item?LinkedStyle)
        IsQuickStyle = optionalBool (item?IsQuickStyle) |> Option.defaultValue false
        Aliases = optionalString (item?StyleAliases) |> Option.orElse (optionalString (item?Aliases))
        UiPriority = optionalInt (item?StyleUiPriority)
        SemiHidden = optionalBool (item?StyleSemiHidden) |> Option.defaultValue false
        UnhideWhenUsed = optionalBool (item?StyleUnhideWhenUsed) |> Option.defaultValue false
        Text = styleTextFromEditor item
        TextBold = styleTextBoldFromEditor item
        TextItalic = styleTextItalicFromEditor item
        TextUnderline = styleTextUnderlineFromEditor item
        TextAllCaps = styleTextAllCapsFromEditor item
        TextSmallCaps = styleTextSmallCapsFromEditor item
        TextStrike = styleTextStrikeFromEditor item
        SpacingBefore = optionalFloat (item?StyleSpacingBefore)
        SpacingAfter = optionalFloat (item?StyleSpacingAfter)
        ContextualSpacing = optionalBool (item?StyleContextualSpacing)
        LineSpacing = optionalFloat (item?StyleLineSpacing)
        LineSpacingRule = optionalString (item?StyleLineRule) |> Option.orElse (optionalString (item?StyleLineSpacingRule))
        Alignment = optionalString (item?StyleAlignment) |> Option.map alignmentFromString
        IndentLeft = optionalFloat (item?StyleIndentLeft)
        IndentRight = optionalFloat (item?StyleIndentRight)
        FirstLineIndent = optionalFloat (item?StyleFirstLineIndent)
        BackgroundColor = optionalString (item?StyleBackgroundColor)
        OutlineLevel = optionalInt (item?StyleOutlineLevel)
        KeepNext = optionalBool (item?StyleKeepNext)
        KeepLines = optionalBool (item?StyleKeepLines)
        PageBreakBefore = optionalBool (item?StylePageBreakBefore)
        WidowControl = optionalBool (item?StyleWidowControl)
    })

let private runStyleFromEditor (run: obj) =
    optionalString (run?Style)
    |> Option.map (fun styleId -> {
        Id = styleId
        Name = optionalString (run?StyleName)
        StyleType = "character"
        IsDefault = optionalBool (run?IsDefaultStyle) |> Option.orElse (optionalBool (run?IsDefault)) |> Option.defaultValue false
        BasedOn = optionalString (run?BasedOnStyle)
        Next = optionalString (run?NextStyle)
        Linked = optionalString (run?LinkedStyle)
        IsQuickStyle = optionalBool (run?IsQuickStyle) |> Option.defaultValue false
        Aliases = optionalString (run?StyleAliases) |> Option.orElse (optionalString (run?Aliases))
        UiPriority = optionalInt (run?StyleUiPriority)
        SemiHidden = optionalBool (run?StyleSemiHidden) |> Option.defaultValue false
        UnhideWhenUsed = optionalBool (run?StyleUnhideWhenUsed) |> Option.defaultValue false
        Text = styleTextFromEditor run
        TextBold = styleTextBoldFromEditor run
        TextItalic = styleTextItalicFromEditor run
        TextUnderline = styleTextUnderlineFromEditor run
        TextAllCaps = styleTextAllCapsFromEditor run
        TextSmallCaps = styleTextSmallCapsFromEditor run
        TextStrike = styleTextStrikeFromEditor run
        SpacingBefore = None
        SpacingAfter = None
        ContextualSpacing = None
        LineSpacing = None
        LineSpacingRule = None
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
    })

let rec private stylesFromItem (item: obj) =
    match item?Type |> stringOrDefault "Paragraph" with
    | "TextBox" when not (isNullOrUndefined item?Paragraphs) ->
        item?Paragraphs |> unbox<obj array> |> Array.toList |> List.collect stylesFromItem
    | "Table" when not (isNullOrUndefined item?Rows) ->
        item?Rows
        |> unbox<obj array>
        |> Array.toList
        |> List.collect (fun row ->
            if isNullOrUndefined row?Cells then []
            else
                row?Cells
                |> unbox<obj array>
                |> Array.toList
                |> List.collect (fun cell ->
                    let paragraphStyles =
                        if isNullOrUndefined cell?Paragraphs then []
                        else cell?Paragraphs |> unbox<obj array> |> Array.toList |> List.collect stylesFromItem
                    let nestedStyles =
                        if isNullOrUndefined cell?NestedTables then []
                        else cell?NestedTables |> unbox<obj array> |> Array.toList |> List.collect stylesFromItem
                    paragraphStyles @ nestedStyles))
    | _ ->
        let runStyles =
            if isNullOrUndefined item?Runs then []
            else item?Runs |> unbox<obj array> |> Array.toList |> List.choose runStyleFromEditor
        (paragraphStyleFromEditor item |> Option.toList) @ runStyles

let private stylesFromEditor (items: obj array) =
    items |> Array.toList |> List.collect stylesFromItem |> List.distinctBy (fun style -> style.StyleType, style.Id)

let private stylesFromCatalog (items: obj array) =
    items
    |> Array.toList
    |> List.choose (fun item ->
        let styleId = optionalString (item?StyleId) |> Option.orElse (optionalString (item?Id)) |> Option.orElse (optionalString (item?Style))
        styleId
        |> Option.map (fun id -> {
            Id = id
            Name = optionalString (item?Name) |> Option.orElse (optionalString (item?StyleName))
            StyleType = optionalString (item?StyleType) |> Option.defaultValue "paragraph"
            IsDefault = optionalBool (item?IsDefaultStyle) |> Option.orElse (optionalBool (item?IsDefault)) |> Option.defaultValue false
            BasedOn = optionalString (item?BasedOn) |> Option.orElse (optionalString (item?BasedOnStyle))
            Next = optionalString (item?Next) |> Option.orElse (optionalString (item?NextStyle))
            Linked = optionalString (item?Linked) |> Option.orElse (optionalString (item?LinkedStyle))
            IsQuickStyle = optionalBool (item?IsQuickStyle) |> Option.defaultValue false
            Aliases = optionalString (item?StyleAliases) |> Option.orElse (optionalString (item?Aliases))
            UiPriority = optionalInt (item?StyleUiPriority)
            SemiHidden = optionalBool (item?StyleSemiHidden) |> Option.defaultValue false
            UnhideWhenUsed = optionalBool (item?StyleUnhideWhenUsed) |> Option.defaultValue false
            Text = styleTextFromEditor item
            TextBold = styleTextBoldFromEditor item
            TextItalic = styleTextItalicFromEditor item
            TextUnderline = styleTextUnderlineFromEditor item
            TextAllCaps = styleTextAllCapsFromEditor item
            TextSmallCaps = styleTextSmallCapsFromEditor item
            TextStrike = styleTextStrikeFromEditor item
            SpacingBefore = optionalFloat (item?StyleSpacingBefore)
            SpacingAfter = optionalFloat (item?StyleSpacingAfter)
            ContextualSpacing = optionalBool (item?StyleContextualSpacing)
            LineSpacing = optionalFloat (item?StyleLineSpacing)
            LineSpacingRule = optionalString (item?StyleLineRule) |> Option.orElse (optionalString (item?StyleLineSpacingRule))
            Alignment = optionalString (item?StyleAlignment) |> Option.map alignmentFromString
            IndentLeft = optionalFloat (item?StyleIndentLeft)
            IndentRight = optionalFloat (item?StyleIndentRight)
            FirstLineIndent = optionalFloat (item?StyleFirstLineIndent)
            BackgroundColor = optionalString (item?StyleBackgroundColor)
            OutlineLevel = optionalInt (item?StyleOutlineLevel)
            KeepNext = optionalBool (item?StyleKeepNext)
            KeepLines = optionalBool (item?StyleKeepLines)
            PageBreakBefore = optionalBool (item?StylePageBreakBefore)
            WidowControl = optionalBool (item?StyleWidowControl)
        }))

let private styleToEditor (style: NamedStyle) =
    createObj [
        "StyleId" ==> style.Id
        "Name" ==> style.Name
        "StyleType" ==> style.StyleType
        "IsDefaultStyle" ==> style.IsDefault
        "BasedOn" ==> style.BasedOn
        "NextStyle" ==> style.Next
        "LinkedStyle" ==> style.Linked
        "IsQuickStyle" ==> style.IsQuickStyle
        "StyleAliases" ==> style.Aliases
        "StyleUiPriority" ==> style.UiPriority
        "StyleSemiHidden" ==> style.SemiHidden
        "StyleUnhideWhenUsed" ==> style.UnhideWhenUsed
        "StyleBold" ==> (style.TextBold |> Option.orElse (style.Text |> Option.map (fun text -> text.Bold)))
        "StyleItalic" ==> (style.TextItalic |> Option.orElse (style.Text |> Option.map (fun text -> text.Italic)))
        "StyleUnderline" ==> (style.TextUnderline |> Option.orElse (style.Text |> Option.map (fun text -> text.Underline)))
        "StyleAllCaps" ==> (style.TextAllCaps |> Option.orElse (style.Text |> Option.map (fun text -> text.AllCaps)))
        "StyleSmallCaps" ==> (style.TextSmallCaps |> Option.orElse (style.Text |> Option.map (fun text -> text.SmallCaps)))
        "StyleStrike" ==> (style.TextStrike |> Option.orElse (style.Text |> Option.map (fun text -> text.Strike)))
        "StyleVertAlign" ==> (style.Text |> Option.bind (fun text -> text.VertAlign))
        "StyleFontSize" ==> (style.Text |> Option.map (fun text -> text.FontSize))
        "StyleCharacterSpacing" ==> (style.Text |> Option.map (fun text -> text.CharacterSpacing))
        "StyleFontFamily" ==> (style.Text |> Option.bind (fun text -> text.FontFamily))
        "StyleColor" ==> (style.Text |> Option.bind (fun text -> text.Color))
        "StyleHighlight" ==> (style.Text |> Option.bind (fun text -> text.Highlight))
        "StyleSpacingBefore" ==> style.SpacingBefore
        "StyleSpacingAfter" ==> style.SpacingAfter
        "StyleContextualSpacing" ==> style.ContextualSpacing
        "StyleLineSpacing" ==> style.LineSpacing
        "StyleLineRule" ==> style.LineSpacingRule
        "StyleAlignment" ==> (style.Alignment |> Option.map alignmentToString)
        "StyleIndentLeft" ==> style.IndentLeft
        "StyleIndentRight" ==> style.IndentRight
        "StyleFirstLineIndent" ==> style.FirstLineIndent
        "StyleBackgroundColor" ==> style.BackgroundColor
        "StyleOutlineLevel" ==> style.OutlineLevel
        "StyleKeepNext" ==> style.KeepNext
        "StyleKeepLines" ==> style.KeepLines
        "StylePageBreakBefore" ==> style.PageBreakBefore
        "StyleWidowControl" ==> style.WidowControl
    ]

let exportStylesToEditor (styles: NamedStyle list) =
    styles |> List.map styleToEditor |> List.toArray

let private paragraphsFromEditor (items: obj array) =
    items
    |> Array.choose (fun item ->
        let itemType = item?Type |> stringOrDefault "Paragraph"
        if itemType = "Table" then None else Some (paragraphFromEditor item))
    |> Array.toList

let rec private blockFromEditor (item: obj) =
    let itemType = item?Type |> stringOrDefault "Paragraph"

    match itemType with
    | "TextBox" ->
        let paragraphs = item?Paragraphs |> unbox<obj array>
        let cellParagraphs = paragraphs |> Array.map paragraphFromEditor |> Array.toList
        TextBox {
            Id = nextImportedBlockId ()
            Content = cellParagraphs
            X = optionalNumberAsFloat (item?X)
            Y = optionalNumberAsFloat (item?Y)
            PositionHRelativeFrom = optionalString (item?PositionHRelativeFrom) |> Option.orElse (optionalString (item?TextBoxPositionHRelativeFrom))
            PositionVRelativeFrom = optionalString (item?PositionVRelativeFrom) |> Option.orElse (optionalString (item?TextBoxPositionVRelativeFrom))
            Width = optionalNumberAsFloat (item?Width) |> Option.defaultValue 220.0
            Height = optionalNumberAsFloat (item?Height) |> Option.defaultValue 72.0
            FillColor = optionalString (item?FillColor) |> Option.orElse (Some "#FFFFFF")
            BorderColor = optionalString (item?BorderColor) |> Option.orElse (Some "#808080")
            BorderWidth = optionalNumberAsFloat (item?BorderWidth) |> Option.orElse (Some 1.0)
            RelativeHeight = optionalString (item?RelativeHeight) |> Option.orElse (optionalString (item?TextBoxRelativeHeight))
            Geometry = geometryFromEditor item
            Gradient = gradientFromEditor (item?Gradient)
            Rotation = optionalNumberAsFloat (item?Rotation) |> Option.defaultValue 0.0
        }

    | "Table" ->
        let rows = item?Rows |> unbox<obj array>
        let tableRows =
            rows
            |> Array.map (fun row ->
                let cells = row?Cells |> unbox<obj array>
                let tableCells =
                    cells
                    |> Array.map (fun cell ->
                        let paragraphs = cell?Paragraphs |> unbox<obj array>
                        let cellParagraphs = paragraphs |> Array.map paragraphFromEditor |> Array.toList
                        let nestedTables =
                            if isNullOrUndefined cell?NestedTables then []
                            else
                                cell?NestedTables
                                |> unbox<obj array>
                                |> Array.toList
                                |> List.choose (fun item ->
                                    match blockFromEditor item with
                                    | Table table -> Some table
                                    | _ -> None)
                        let nestedTableAnchors =
                            if isNullOrUndefined cell?NestedTableAnchors then []
                            else cell?NestedTableAnchors |> unbox<int array> |> Array.toList
                        {
                            Content = cellParagraphs
                            NestedTables = nestedTables
                            NestedTableAnchors = nestedTableAnchors
                            ColSpan = optionalInt (cell?ColSpan) |> Option.defaultValue 1 |> max 1
                            BackgroundColor = optionalString (cell?BackgroundColor)
                            Shading = shadingFromEditorFields cell
                            Width = optionalIntAsFloat (cell?Width)
                            PreferredWidth = tableWidthFromEditor (cell?PreferredWidth) (cell?PreferredWidthType)
                            BorderColor = optionalString (cell?BorderColor)
                            BorderWidth = optionalFloat (cell?BorderWidth)
                            Borders = tableBordersFromEditor (cell?Borders)
                            VerticalAlignment = optionalString (cell?VerticalAlignment)
                            VerticalMerge = tableVerticalMergeFromString (optionalString (cell?VerticalMerge))
                            PaddingTop = optionalDxaAsCanvasPixels (cell?PaddingTop)
                            PaddingRight = optionalDxaAsCanvasPixels (cell?PaddingRight)
                            PaddingBottom = optionalDxaAsCanvasPixels (cell?PaddingBottom)
                            PaddingLeft = optionalDxaAsCanvasPixels (cell?PaddingLeft)
                        })
                    |> Array.toList
                let canSplit = if isNullOrUndefined (row?CanSplit) then None else Some (unbox<bool> row?CanSplit)
                { Cells = tableCells
                  IsHeader = boolOrDefault false (row?IsHeader)
                  Height = optionalNumberAsFloat (row?Height)
                  HeightRule = optionalString (row?HeightRule)
                  CanSplit = canSplit
                  CellSpacing = optionalDxaAsCanvasPixels (row?CellSpacing)
                  PaddingTop = optionalDxaAsCanvasPixels (row?PaddingTop)
                  PaddingRight = optionalDxaAsCanvasPixels (row?PaddingRight)
                  PaddingBottom = optionalDxaAsCanvasPixels (row?PaddingBottom)
                  PaddingLeft = optionalDxaAsCanvasPixels (row?PaddingLeft)
                  GridBefore = optionalInt (row?GridBefore) |> Option.defaultValue 0 |> max 0
                  GridAfter = optionalInt (row?GridAfter) |> Option.defaultValue 0 |> max 0
                  WidthBefore = optionalDxaAsCanvasPixels (row?WidthBefore)
                  WidthAfter = optionalDxaAsCanvasPixels (row?WidthAfter)
                  BordersOverride = None })
            |> Array.toList

        let gridColumns =
            if isNullOrUndefined (item?GridColumns) then []
            else
                item?GridColumns
                |> unbox<obj array>
                |> Array.choose (fun value -> tableWidthFromEditor value (box "dxa"))
                |> Array.toList

        Table {
            Id = nextImportedBlockId ()
            Rows = WordLayout.Table.distributeVerticalMergeContent tableRows
            PreferredWidth = tableWidthFromEditor (item?PreferredWidth) (item?PreferredWidthType)
            Alignment = optionalString (item?TableAlignment) |> Option.orElse (optionalString (item?Alignment))
            LayoutMode = tableLayoutModeFromString (optionalString (item?LayoutMode))
            GridColumns = gridColumns
            Borders = tableBordersFromEditor (item?Borders)
            BackgroundColor = optionalString (item?BackgroundColor)
            Shading = shadingFromEditorFields item
            CellSpacing = optionalDxaAsCanvasPixels (item?CellSpacing)
            PaddingTop = optionalDxaAsCanvasPixels (item?PaddingTop)
            PaddingRight = optionalDxaAsCanvasPixels (item?PaddingRight)
            PaddingBottom = optionalDxaAsCanvasPixels (item?PaddingBottom)
            PaddingLeft = optionalDxaAsCanvasPixels (item?PaddingLeft)
            Floating = boolOrDefault false (item?Floating)
        }

    | "Heading" ->
        let level = item?HeadingLevel |> intOrDefault 1 |> max 1 |> min 3
        Heading (level, paragraphFromEditor item)

    | _ ->
        Paragraph (paragraphFromEditor item)

let private blocksFromEditor (items: obj array) =
    items |> Array.map blockFromEditor |> Array.toList

let private isDiscardableTrailingParagraph para =
    System.String.IsNullOrWhiteSpace(TextBuffer.toString para.Buffer)
    && not para.Metadata.PageBreakAfter
    && para.Metadata.SectionBreak.IsNone
    && para.Metadata.RenderedPageBreakBefore.IsNone
    && para.Metadata.SectionPageSetup.IsNone
    && para.Metadata.HeaderReferences.IsEmpty
    && para.Metadata.FooterReferences.IsEmpty
    && para.Metadata.ParagraphBorders.IsNone
    && para.Metadata.ParagraphBorderBefore.IsNone
    && para.Metadata.ParagraphBorderAfter.IsNone

let private trimTrailingEmptyParagraphs blocks =
    let rec trim = function
        | (Paragraph para) :: rest when isDiscardableTrailingParagraph para -> trim rest
        | (Heading (_, para)) :: rest when isDiscardableTrailingParagraph para -> trim rest
        | rest -> rest
    blocks |> List.rev |> trim |> List.rev

let private paragraphsFromBlocks blocks =
    blocks
    |> List.choose (function
        | Paragraph para
        | Heading (_, para) -> Some para
        | Table _
        | TextBox _ -> None)

let private headerFooterPartsFromEditor (items: obj array) =
    items
    |> Array.map (fun item ->
        let kind = item?Type |> stringOrDefault "default"
        let relationshipId = optionalString (item?RelationshipId)
        let blocks = item?Paragraphs |> unbox<obj array> |> blocksFromEditor
        { Kind = kind; RelationshipId = relationshipId; Blocks = blocks })
    |> Array.toList

let private pageSetupFromEditor (value: obj) =
    {
        PageWidth = optionalFloat (value?PageWidth) |> Option.defaultValue defaultPageSetup.PageWidth
        PageHeight = optionalFloat (value?PageHeight) |> Option.defaultValue defaultPageSetup.PageHeight
        MarginLeft = optionalFloat (value?MarginLeft) |> Option.defaultValue defaultPageSetup.MarginLeft
        MarginRight = optionalFloat (value?MarginRight) |> Option.defaultValue defaultPageSetup.MarginRight
        MarginTop = optionalFloat (value?MarginTop) |> Option.defaultValue defaultPageSetup.MarginTop
        MarginBottom = optionalFloat (value?MarginBottom) |> Option.defaultValue defaultPageSetup.MarginBottom
        HeaderDistance = optionalFloat (value?HeaderDistance) |> Option.defaultValue defaultPageSetup.HeaderDistance
        FooterDistance = optionalFloat (value?FooterDistance) |> Option.defaultValue defaultPageSetup.FooterDistance
        DifferentFirstPage = optionalBool (value?DifferentFirstPage) |> Option.defaultValue defaultPageSetup.DifferentFirstPage
        EvenAndOddHeaders = optionalBool (value?EvenAndOddHeaders) |> Option.defaultValue defaultPageSetup.EvenAndOddHeaders
        Columns =
            match optionalFloat (value?ColumnCount) with
            | Some count when count >= 2.0 ->
                Some ({ Count = int count
                        Gap = optionalFloat (value?ColumnSpace) |> Option.defaultValue 0.0
                        Separator = optionalBool (value?ColumnSeparator) |> Option.defaultValue false
                        Widths =
                            match value?ColumnWidths with
                            | w when isNullOrUndefined w -> []
                            | w -> [ for i in 0 .. jsArrayLength w - 1 -> let o = jsArrayGet w i in (optionalFloat (o?W) |> Option.defaultValue 0.0, optionalFloat (o?S) |> Option.defaultValue 0.0) ] }: SectionColumns)
            | _ -> None
        SectionStart = optionalString (value?SectionStart)
    }

/// Import EditorParagraphs from Office.Formats into Word document
let importFromEditor (editorParagraphs: obj array) : Document =
    let blocks = blocksFromEditor editorParagraphs |> trimTrailingEmptyParagraphs
    let styles = stylesFromEditor editorParagraphs

    // Create document with all blocks
    let doc = { Blocks = blocks; Selection = None; PageSetup = defaultPageSetup; Styles = styles; Headers = []; Footers = []; HeaderParts = []; FooterParts = [] }

    // Set cursor at start of first block (paragraph only, skip tables for now)
    match blocks |> List.tryFind (function | Paragraph _ -> true | _ -> false) with
    | Some block ->
        let para = getParagraph block
        { doc with Selection = Some { Anchor = { BlockId = para.Id; Offset = 0 }; Focus = { BlockId = para.Id; Offset = 0 } } }
    | None ->
        doc

let importFromEditorWithStyles (editorParagraphs: obj array) (styleCatalog: obj array) : Document =
    let doc = importFromEditor editorParagraphs
    let styles =
        (stylesFromCatalog styleCatalog @ doc.Styles)
        |> List.distinctBy (fun style -> style.StyleType, style.Id)
    { doc with Styles = styles }

let importFromEditorWithHeaderFooters (editorParagraphs: obj array) (headerParagraphs: obj array) (footerParagraphs: obj array) : Document =
    let doc = importFromEditor editorParagraphs
    { doc with
        Headers = paragraphsFromEditor headerParagraphs
        Footers = paragraphsFromEditor footerParagraphs }

let importFromEditorWithHeaderFooterParts (editorParagraphs: obj array) (headerParts: obj array) (footerParts: obj array) : Document =
    let doc = importFromEditor editorParagraphs
    let headers = headerFooterPartsFromEditor headerParts
    let footers = headerFooterPartsFromEditor footerParts
    { doc with
        HeaderParts = headers
        FooterParts = footers
        Headers = headers |> List.tryHead |> Option.map (fun part -> paragraphsFromBlocks part.Blocks) |> Option.defaultValue []
        Footers = footers |> List.tryHead |> Option.map (fun part -> paragraphsFromBlocks part.Blocks) |> Option.defaultValue [] }

let importFromEditorWithPageSetupAndHeaderFooterParts (editorParagraphs: obj array) (pageSetup: obj) (headerParts: obj array) (footerParts: obj array) : Document =
    let doc = importFromEditorWithHeaderFooterParts editorParagraphs headerParts footerParts
    { doc with PageSetup = pageSetupFromEditor pageSetup }

let private optionalArray (value: obj) =
    if isNullOrUndefined value then [||] else unbox<obj array> value

let importFromEditorPackage (package: obj) : Document =
    let paragraphs = optionalArray (package?Paragraphs)
    let styles = optionalArray (package?Styles)
    let headerParts = optionalArray (package?HeaderParts)
    let footerParts = optionalArray (package?FooterParts)
    let doc = importFromEditorWithHeaderFooterParts paragraphs headerParts footerParts
    let doc =
        let importedStyles =
            (stylesFromCatalog styles @ doc.Styles)
            |> List.distinctBy (fun style -> style.StyleType, style.Id)
        { doc with Styles = importedStyles }
    if isNullOrUndefined package?PageSetup then doc else { doc with PageSetup = pageSetupFromEditor package?PageSetup }

/// Export Word document to EditorParagraphs for Office.Formats
let private paragraphToEditor itemType headingLevel (para: ParagraphBlock) =
    let buffer = para.Buffer

    let runs =
        buffer.Runs
        |> List.choose (function
            | Text tr ->
                let text = Rope.slice tr.Start tr.Length buffer.Rope
                Some (createObj [
                    "Text" ==> text
                    "Style" ==> tr.Style.Style
                    "Bold" ==> tr.Style.Bold
                    "Italic" ==> tr.Style.Italic
                    "Underline" ==> tr.Style.Underline
                    "AllCaps" ==> tr.Style.AllCaps
                    "SmallCaps" ==> tr.Style.SmallCaps
                    "Strike" ==> tr.Style.Strike
                    "FontSize" ==> tr.Style.FontSize
                    "CharacterSpacing" ==> tr.Style.CharacterSpacing
                    "FontFamily" ==> (tr.Style.FontFamily |> Option.defaultValue "Calibri")
                    "Color" ==> tr.Style.Color
                    "Highlight" ==> tr.Style.Highlight
                    "Hyperlink" ==> tr.Style.Hyperlink
                    "CommentId" ==> tr.Style.CommentId
                    "CommentText" ==> tr.Style.CommentText
                    "CommentAuthor" ==> tr.Style.CommentAuthor
                    "CommentDate" ==> tr.Style.CommentDate
                    "NoteId" ==> tr.Style.NoteId
                    "NoteType" ==> tr.Style.NoteType
                    "NoteText" ==> tr.Style.NoteText
                    "BookmarkStartId" ==> tr.Style.BookmarkStartId
                    "BookmarkName" ==> tr.Style.BookmarkName
                    "BookmarkEndId" ==> tr.Style.BookmarkEndId
                    "FieldCode" ==> tr.Style.FieldCode
                    "FieldResult" ==> tr.Style.FieldResult
                    "RenderedPageBreakBefore" ==> tr.Style.RenderedPageBreakBefore
                    "IsEquation" ==> tr.Style.EquationXml.IsSome
                    "EquationXml" ==> tr.Style.EquationXml
                    "RevisionId" ==> tr.Style.RevisionId
                    "RevisionType" ==> tr.Style.RevisionType
                    "RevisionAuthor" ==> tr.Style.RevisionAuthor
                    "RevisionDate" ==> tr.Style.RevisionDate
                    "VertAlign" ==> tr.Style.VertAlign
                    "IsImage" ==> false
                ])
            | Image img ->
                Some (createObj [
                    "Text" ==> ""
                    "IsImage" ==> true
                    "ImageSrc" ==> img.Src
                    "ImageLinked" ==> img.Linked
                    "ImageWidth" ==> img.Width
                    "ImageHeight" ==> img.Height
                    "ImageAlt" ==> img.Alt
                    "ImageName" ==> img.Name
                    "ImageTitle" ==> img.Title
                    "ImageFloating" ==> img.Floating
                    "ImageWrap" ==> img.Wrap
                    "ImagePositionX" ==> img.PositionX
                    "ImagePositionY" ==> img.PositionY
                    "ImagePositionHRelativeFrom" ==> img.PositionHRelativeFrom
                    "ImagePositionVRelativeFrom" ==> img.PositionVRelativeFrom
                    "ImageBehindDoc" ==> img.BehindDoc
                    "ImageAllowOverlap" ==> img.AllowOverlap
                    "ImageLayoutInCell" ==> img.LayoutInCell
                    "ImageDistanceTop" ==> img.DistanceTop
                    "ImageDistanceBottom" ==> img.DistanceBottom
                    "ImageDistanceLeft" ==> img.DistanceLeft
                    "ImageDistanceRight" ==> img.DistanceRight
                    "ImageRelativeHeight" ==> img.RelativeHeight
                    "ImageLocked" ==> img.Locked
                    "ImageSimplePosition" ==> img.SimplePosition
                    "ImageSimplePositionX" ==> img.SimplePositionX
                    "ImageSimplePositionY" ==> img.SimplePositionY
                    "ImageCropLeft" ==> img.CropLeft
                    "ImageCropTop" ==> img.CropTop
                    "ImageCropRight" ==> img.CropRight
                    "ImageCropBottom" ==> img.CropBottom
                ])
        )

    createObj [
        "Type" ==> itemType
        "HeadingLevel" ==> headingLevel
        "Runs" ==> (runs |> List.toArray)
        "Style" ==> para.Metadata.Style
        "ContentControlId" ==> para.Metadata.ContentControlId
        "ContentControlTag" ==> para.Metadata.ContentControlTag
        "ContentControlAlias" ==> para.Metadata.ContentControlAlias
        "Alignment" ==> alignmentToString para.Metadata.Alignment
        "ListStyle" ==> listStyleToString para.Metadata.ListStyle
        "ListLevel" ==> para.Metadata.ListLevel
        "IndentLeft" ==> para.Metadata.IndentLeft
        "IndentRight" ==> para.Metadata.IndentRight
        "FirstLineIndent" ==> para.Metadata.FirstLineIndent
        "SpacingBefore" ==> para.Metadata.SpacingBefore
        "SpacingAfter" ==> para.Metadata.SpacingAfter
        "StyleContextualSpacing" ==> para.Metadata.ContextualSpacing
        "LineSpacing" ==> para.Metadata.LineSpacing
        "LineSpacingRule" ==> para.Metadata.LineSpacingRule
        "WidowControl" ==> para.Metadata.WidowControl
        "BackgroundColor" ==> para.Metadata.BackgroundColor
        "ParagraphBorders" ==> paragraphBordersToEditor para.Metadata.ParagraphBorders
        "ParagraphBorderBefore" ==> para.Metadata.ParagraphBorderBefore
        "ParagraphBorderAfter" ==> para.Metadata.ParagraphBorderAfter
        "PageBreakBefore" ==> para.Metadata.PageBreakBefore
        "PageBreakAfter" ==> para.Metadata.PageBreakAfter
        "SectionBreak" ==> para.Metadata.SectionBreak
        "SectionPageNumberStart" ==> para.Metadata.SectionPageNumberStart
        "RenderedPageBreakBefore" ==> para.Metadata.RenderedPageBreakBefore
        "SectionPageSetup" ==> (para.Metadata.SectionPageSetup |> Option.map sectionPageSetupToEditor |> Option.toObj)
        "HeaderReferences" ==> refsToEditor para.Metadata.HeaderReferences
        "FooterReferences" ==> refsToEditor para.Metadata.FooterReferences
    ]

let rec private blockToEditor block =
    match block with
    | Paragraph para -> paragraphToEditor "Paragraph" 0 para
    | Heading (level, para) -> paragraphToEditor "Heading" level para
    | TextBox textBox ->
        createObj [
            "Type" ==> "TextBox"
            "X" ==> textBox.X
            "Y" ==> textBox.Y
            "PositionHRelativeFrom" ==> textBox.PositionHRelativeFrom
            "PositionVRelativeFrom" ==> textBox.PositionVRelativeFrom
            "TextBoxPositionHRelativeFrom" ==> textBox.PositionHRelativeFrom
            "TextBoxPositionVRelativeFrom" ==> textBox.PositionVRelativeFrom
            "Width" ==> textBox.Width
            "Height" ==> textBox.Height
            "FillColor" ==> textBox.FillColor
            "BorderColor" ==> textBox.BorderColor
            "BorderWidth" ==> textBox.BorderWidth
            "RelativeHeight" ==> textBox.RelativeHeight
            "TextBoxRelativeHeight" ==> textBox.RelativeHeight
            "Paragraphs" ==> (textBox.Content |> List.map (paragraphToEditor "Paragraph" 0) |> List.toArray)
        ]
    | Table table ->
            let rows =
                table.Rows
                |> List.map (fun row ->
                    let cells =
                        row.Cells
                        |> List.map (fun cell ->
                            createObj [
                                "Paragraphs" ==> (cell.Content |> List.map (paragraphToEditor "Paragraph" 0) |> List.toArray)
                                "NestedTables" ==> (cell.NestedTables |> List.map (Table >> blockToEditor) |> List.toArray)
                                "NestedTableAnchors" ==> (cell.NestedTableAnchors |> List.toArray)
                                "ColSpan" ==> cell.ColSpan
                                "BackgroundColor" ==> cell.BackgroundColor
                                "Shading" ==> shadingToEditor cell.Shading
                                "Width" ==> cell.Width
                                "PreferredWidth" ==> tableWidthToEditor cell.PreferredWidth
                                "PreferredWidthType" ==> tableWidthTypeToEditor cell.PreferredWidth
                                "BorderColor" ==> cell.BorderColor
                                "BorderWidth" ==> cell.BorderWidth
                                "Borders" ==> tableBordersToEditor cell.Borders
                                "VerticalAlignment" ==> cell.VerticalAlignment
                                "VerticalMerge" ==> tableVerticalMergeToString cell.VerticalMerge
                                "PaddingTop" ==> cell.PaddingTop
                                "PaddingRight" ==> cell.PaddingRight
                                "PaddingBottom" ==> cell.PaddingBottom
                                "PaddingLeft" ==> cell.PaddingLeft
                            ])
                        |> List.toArray

                    createObj [
                        "Cells" ==> cells
                        "IsHeader" ==> row.IsHeader
                        "Height" ==> row.Height
                        "HeightRule" ==> row.HeightRule
                        "CanSplit" ==> row.CanSplit
                        "CellSpacing" ==> row.CellSpacing
                        "PaddingTop" ==> row.PaddingTop
                        "PaddingRight" ==> row.PaddingRight
                        "PaddingBottom" ==> row.PaddingBottom
                        "PaddingLeft" ==> row.PaddingLeft
                    ])
                |> List.toArray

            createObj [
                "Type" ==> "Table"
                "Rows" ==> rows
                "PreferredWidth" ==> tableWidthToEditor table.PreferredWidth
                "PreferredWidthType" ==> tableWidthTypeToEditor table.PreferredWidth
                "TableAlignment" ==> table.Alignment
                "LayoutMode" ==> tableLayoutModeToString table.LayoutMode
                "GridColumns" ==> (table.GridColumns |> List.map (fun width -> width.Value |> Option.defaultValue 0.0) |> List.toArray)
                "Borders" ==> tableBordersToEditor table.Borders
                "BackgroundColor" ==> table.BackgroundColor
                "Shading" ==> shadingToEditor table.Shading
                "CellSpacing" ==> table.CellSpacing
                "PaddingTop" ==> table.PaddingTop
                "PaddingRight" ==> table.PaddingRight
                "PaddingBottom" ==> table.PaddingBottom
                "PaddingLeft" ==> table.PaddingLeft
            ]

let exportToEditor (doc: Document) : obj array =
    doc.Blocks
    |> List.map blockToEditor
    |> List.toArray

let exportHeaderFooterPartsToEditor (parts: HeaderFooterPart list) : obj array =
    parts
    |> List.map (fun part ->
        createObj [
            "Type" ==> part.Kind
            "Kind" ==> part.Kind
            "RelationshipId" ==> part.RelationshipId
            "Paragraphs" ==> (part.Blocks |> List.map blockToEditor |> List.toArray)
            "Blocks" ==> (part.Blocks |> List.map blockToEditor |> List.toArray)
        ])
    |> List.toArray
