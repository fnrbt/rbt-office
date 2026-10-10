module Rbt.Word.Docx

open System.Text
open System.Collections.Generic
open Fable.Core
open Fable.Core.JsInterop
open Rbt.Word.Document
open Rbt.Word.TextBuffer
open Rbt.Word.Rope

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

/// Inverse of optionalDxaAsCanvasPixels, for the export: the editor package
/// carries table paddings and cell spacing in dxa (that is what the reader
/// emits and what the writer puts in w:w), the document model in canvas px.
/// Exporting the px value unchanged shrank every padding 15x per save.
let private canvasPixelsToDxa (value: float option) =
    value |> Option.map (fun pixels -> System.Math.Floor (pixels * WordLayout.canvasScale.DxaPerUnit + 0.5))

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
    // The reader passes w:jc through verbatim, so the OOXML spellings arrive here too.
    match align.ToLower() with
    | "center" -> Center
    | "right" | "end" -> Right
    | "justify" | "both" | "distribute" -> Justify
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
    let sectionPageNumberFormat = optionalString (item?SectionPageNumberFormat)
    let renderedPageBreakBefore = optionalInt (item?RenderedPageBreakBefore)
    let sectionPageSetup = sectionPageSetupFromEditor (item?SectionPageSetup)
    let sectionColumns : SectionColumns option =
        let value = item?SectionColumns
        if isNullOrUndefined value then None
        else
            match optionalFloat (value?Count) with
            | Some count when count >= 2.0 ->
                let widths = value?Widths
                Some { Count = int count
                       Gap = optionalFloat (value?Gap) |> Option.defaultValue 0.0
                       Separator = value?Separator |> boolOrDefault false
                       Widths =
                           if isNullOrUndefined widths then []
                           else [ for i in 0 .. jsArrayLength widths - 1 ->
                                    let column = jsArrayGet widths i
                                    (optionalFloat (column?W) |> Option.defaultValue 0.0, optionalFloat (column?S) |> Option.defaultValue 0.0) ] }
            | _ -> None
    let headerReferences = refsFromEditor (item?HeaderReferences)
    let footerReferences = refsFromEditor (item?FooterReferences)
    {
        defaultMetadata with
            Style = style
            ContentControlId = contentControlId
            ContentControlTag = contentControlTag
            ContentControlAlias = contentControlAlias
            Alignment = alignment
            ExplicitAlignment = (optionalString (item?ExplicitAlignment)).IsSome
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
            KeepNext = optionalBool (item?KeepNext)
            KeepLines = optionalBool (item?KeepLines)
            BackgroundColor = backgroundColor
            ParagraphBorders = paragraphBorders
            ParagraphBorderBefore = paragraphBorderBefore
            ParagraphBorderAfter = paragraphBorderAfter
            PageBreakBefore = pageBreakBefore
            PageBreakAfter = pageBreakAfter
            SectionBreak = sectionBreak
            SectionPageNumberStart = sectionPageNumberStart
            SectionPageNumberFormat = sectionPageNumberFormat
            RenderedPageBreakBefore = renderedPageBreakBefore
            SectionPageSetup = sectionPageSetup
            SectionColumns = sectionColumns
            HeaderReferences = headerReferences
            FooterReferences = footerReferences
            // What the paragraph itself states. An object that does not say
            // (ExplicitParagraphFormat missing) only has effective values, and all
            // of them count as its own.
            Explicit =
                if boolOrDefault false (item?ExplicitParagraphFormat) then
                    let flag (value: obj) (bit: int) = if isNullOrUndefined value then 0 else bit
                    flag (item?ExplicitSpacingBefore) ParagraphExplicit.SpacingBefore
                    ||| flag (item?ExplicitSpacingAfter) ParagraphExplicit.SpacingAfter
                    ||| flag (item?ExplicitLineSpacing) ParagraphExplicit.LineSpacing
                    ||| flag (item?ExplicitIndentLeft) ParagraphExplicit.IndentLeft
                    ||| flag (item?ExplicitIndentRight) ParagraphExplicit.IndentRight
                    ||| flag (item?ExplicitFirstLineIndent) ParagraphExplicit.FirstLineIndent
                    ||| flag (item?ExplicitContextualSpacing) ParagraphExplicit.ContextualSpacing
                    ||| flag (item?ExplicitBackgroundColor) ParagraphExplicit.BackgroundColor
                    ||| (if boolOrDefault false (item?ExplicitTabs) then ParagraphExplicit.Tabs else 0)
                else
                    let flag (isSet: bool) (bit: int) = if isSet then bit else 0
                    flag spacingBefore.IsSome ParagraphExplicit.SpacingBefore
                    ||| flag spacingAfter.IsSome ParagraphExplicit.SpacingAfter
                    ||| flag lineSpacing.IsSome ParagraphExplicit.LineSpacing
                    ||| flag indentLeft.IsSome ParagraphExplicit.IndentLeft
                    ||| flag indentRight.IsSome ParagraphExplicit.IndentRight
                    ||| flag firstLineIndent.IsSome ParagraphExplicit.FirstLineIndent
                    ||| flag (optionalBool (item?StyleContextualSpacing)).IsSome ParagraphExplicit.ContextualSpacing
                    ||| flag backgroundColor.IsSome ParagraphExplicit.BackgroundColor
                    ||| flag (not tabStops.IsEmpty || rightTabPosition.IsSome) ParagraphExplicit.Tabs
                |> fun flags -> if pageBreakBefore && boolOrDefault false (item?PageBreakBeforeDerived) then flags ||| ParagraphExplicit.DerivedPageBreakBefore else flags
            HasMarkProperties = boolOrDefault false (item?HasMarkProperties)
            ExplicitMarkFontSize = optionalFloat (item?ExplicitMarkFontSize)
            MarkRevision =
                optionalString (item?MarkRevisionType)
                |> Option.map (fun revisionType ->
                    { RevisionType = revisionType
                      Mark = { Id = optionalString (item?MarkRevisionId); Author = optionalString (item?MarkRevisionAuthor); Date = optionalString (item?MarkRevisionDate) } })
            ListLevelIndents =
                if boolOrDefault false (item?HasListLevelIndents) then Some (optionalFloat (item?ListLevelIndentLeft), optionalFloat (item?ListLevelFirstLineIndent))
                else None
    }

/// Convert EditorRun array to TextBuffer
let private editorRunsToBuffer (runs: obj array) : TextBuffer =
    // Build full text and create runs
    let mutable currentOffset = 0
    let textParts = StringBuilder()
    let bufferRuns = ResizeArray<Run>()

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
            textParts.Append(TextBuffer.objectReplacementChar) |> ignore

            // Create image run
            bufferRuns.Add(Image {
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
                PositionHAlign = optionalString (run?ImagePositionHAlign)
                PositionVAlign = optionalString (run?ImagePositionVAlign)
                WrapText = optionalString (run?ImageWrapText)
                WrapPolygon =
                    (let points = run?ImageWrapPolygon
                     if isNullOrUndefined points then []
                     else
                         let values = unbox<float array> points
                         [ for i in 0 .. values.Length / 2 - 1 -> values.[i * 2], values.[i * 2 + 1] ])
                Opaque =
                    optionalString (run?OpaqueXml)
                    |> Option.map (fun xml ->
                        { Xml = xml
                          Kind = run?OpaqueKind |> stringOrDefault "drawing"
                          ParagraphLevel = run?OpaqueParagraphLevel |> boolOrDefault false })
            })

            currentOffset <- currentOffset + 1
        else
            // Text run
            let text = run?Text |> unbox<string>

            // Don't skip whitespace-only runs - they're important for spacing!
            if not (System.String.IsNullOrEmpty(text)) then
                let bold = run?Bold |> boolOrDefault false
                let italic = run?Italic |> boolOrDefault false
                let underline = run?Underline |> boolOrDefault false
                // (A single line is what an underline is when nothing else is said: text underlined in
                // the editor has no line style, and text read with w:u w:val="single" is the same text.)
                let underlineStyle = optionalString (run?UnderlineStyle) |> Option.filter (fun style -> style <> "single")
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
                // Only a target that may be followed is a link; a package object that
                // offers anything else as one (or says so itself) has it kept inactive.
                let offeredHyperlink = optionalString (run?Hyperlink)
                let hyperlink = offeredHyperlink |> Option.filter TextBuffer.isActiveHyperlinkTarget
                let inactiveHyperlink =
                    optionalString (run?InactiveHyperlink)
                    |> Option.orElse (offeredHyperlink |> Option.filter (fun target -> not (TextBuffer.isActiveHyperlinkTarget target)))
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
                // Set directly on the run (ExplicitX = true/false) as opposed to
                // inherited (null): the value then overrides the styles.
                let mutable explicitFlags = 0
                let direct (value: obj) flag (effective: bool) =
                    match optionalBool value with
                    | Some value ->
                        explicitFlags <- explicitFlags ||| flag
                        value
                    | None -> effective
                let bold = direct (run?ExplicitBold) ExplicitFormat.Bold bold
                let italic = direct (run?ExplicitItalic) ExplicitFormat.Italic italic
                let underline = direct (run?ExplicitUnderline) ExplicitFormat.Underline underline
                let allCaps = direct (run?ExplicitAllCaps) ExplicitFormat.AllCaps allCaps
                let smallCaps = direct (run?ExplicitSmallCaps) ExplicitFormat.SmallCaps smallCaps
                let strike = direct (run?ExplicitStrike) ExplicitFormat.Strike strike
                let doubleStrike = direct (run?ExplicitDoubleStrike) ExplicitFormat.DoubleStrike doubleStrike
                let hidden = direct (run?ExplicitHidden) ExplicitFormat.Hidden hidden
                let fontSize =
                    match optionalFloat (run?ExplicitFontSize) with
                    | Some value ->
                        explicitFlags <- explicitFlags ||| ExplicitFormat.FontSize
                        value
                    | None -> fontSize
                if boolOrDefault false (run?ExplicitColorAuto) && color.IsNone then
                    explicitFlags <- explicitFlags ||| ExplicitFormat.ColorAuto

                let style = {
                    Style = styleId
                    Bold = bold
                    Italic = italic
                    Underline = underline || hyperlink.IsSome || inactiveHyperlink.IsSome
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
                    InactiveHyperlink = inactiveHyperlink
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
                    // A deletion inside a tracked insertion, and a tracked move.
                    DeletedInsertion =
                        if isNullOrUndefined run?InsertedRevisionId && isNullOrUndefined run?InsertedRevisionAuthor && isNullOrUndefined run?InsertedRevisionDate then None
                        else Some { Id = optionalString (run?InsertedRevisionId); Author = optionalString (run?InsertedRevisionAuthor); Date = optionalString (run?InsertedRevisionDate) }
                    MoveName = optionalString (run?RevisionMoveName)
                    VertAlign = vertAlign
                    Explicit = explicitFlags
                }

                textParts.Append(text) |> ignore
                if style.RenderedPageBreakBefore && text.Length > 1 then
                    bufferRuns.Add(Text { Start = currentOffset; Length = 1; Style = style })
                    bufferRuns.Add(Text { Start = currentOffset + 1; Length = text.Length - 1; Style = { style with RenderedPageBreakBefore = false } })
                else
                    bufferRuns.Add(Text {
                        Start = currentOffset
                        Length = text.Length
                        Style = style
                    })

                currentOffset <- currentOffset + text.Length

    let fullText = textParts.ToString()
    let rope = Rope.ofString fullText

    { Rope = rope; Runs = List.ofSeq bufferRuns }

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
                UnderlineStyle = optionalString (value?StyleUnderlineStyle) |> Option.filter (fun style -> style <> "single")
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
                Explicit = (if isNullOrUndefined value?StyleFontSize then 0 else ExplicitFormat.FontSize)
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

let private addStyleIfNew (styles: ResizeArray<NamedStyle>) (seen: Dictionary<string, Dictionary<string, unit>>) style =
    let mutable ids = Unchecked.defaultof<Dictionary<string, unit>>
    if not (seen.TryGetValue(style.StyleType, &ids)) then
        ids <- Dictionary<string, unit>()
        seen.[style.StyleType] <- ids
    if not (ids.ContainsKey style.Id) then
        ids.[style.Id] <- ()
        styles.Add style

let rec private collectStylesFromItem (styles: ResizeArray<NamedStyle>) (seen: Dictionary<string, Dictionary<string, unit>>) (item: obj) =
    match item?Type |> stringOrDefault "Paragraph" with
    | "TextBox" when not (isNullOrUndefined item?Paragraphs) ->
        let paragraphs = item?Paragraphs |> unbox<obj array>
        for paragraph in paragraphs do
            collectStylesFromItem styles seen paragraph
    | "Table" when not (isNullOrUndefined item?Rows) ->
        let rows = item?Rows |> unbox<obj array>
        for row in rows do
            if not (isNullOrUndefined row?Cells) then
                let cells = row?Cells |> unbox<obj array>
                for cell in cells do
                    if not (isNullOrUndefined cell?Paragraphs) then
                        let paragraphs = cell?Paragraphs |> unbox<obj array>
                        for paragraph in paragraphs do
                            collectStylesFromItem styles seen paragraph
                    if not (isNullOrUndefined cell?NestedTables) then
                        let nestedTables = cell?NestedTables |> unbox<obj array>
                        for nestedTable in nestedTables do
                            collectStylesFromItem styles seen nestedTable
    | _ ->
        paragraphStyleFromEditor item |> Option.iter (addStyleIfNew styles seen)
        if not (isNullOrUndefined item?Runs) then
            let runs = item?Runs |> unbox<obj array>
            for run in runs do
                runStyleFromEditor run |> Option.iter (addStyleIfNew styles seen)

let private stylesFromEditor (items: obj array) =
    let styles = ResizeArray<NamedStyle>()
    let seen = Dictionary<string, Dictionary<string, unit>>()
    for item in items do
        collectStylesFromItem styles seen item
    List.ofSeq styles

let private styleFromCatalog item =
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
    })

let private collectStylesFromCatalog (styles: ResizeArray<NamedStyle>) (seen: Dictionary<string, Dictionary<string, unit>>) (items: obj array) =
    for item in items do
        styleFromCatalog item |> Option.iter (addStyleIfNew styles seen)

let private mergeCatalogAndInferredStyles (catalog: obj array) (inferred: NamedStyle list) =
    let styles = ResizeArray<NamedStyle>()
    let seen = Dictionary<string, Dictionary<string, unit>>()
    collectStylesFromCatalog styles seen catalog
    for style in inferred do
        addStyleIfNew styles seen style
    List.ofSeq styles

/// A style states only what it sets. NamedStyle.Text holds plain values where
/// "off" and the default size also mean "not set", so only "on" (or the
/// TextX option, which records an explicit off) is exported: writing the unset
/// values out made every style switch bold/italic/... off and pin 8.5pt, which
/// overrode the style it is based on and the paragraph style of a character style.
let private styleTextOn (style: NamedStyle) (pick: TextStyle -> bool) : bool option =
    style.Text |> Option.bind (fun text -> if pick text then Some true else None)

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
        "StyleBold" ==> (style.TextBold |> Option.orElse (styleTextOn style (fun text -> text.Bold)))
        "StyleItalic" ==> (style.TextItalic |> Option.orElse (styleTextOn style (fun text -> text.Italic)))
        "StyleUnderline" ==> (style.TextUnderline |> Option.orElse (styleTextOn style (fun text -> text.Underline)))
        "StyleAllCaps" ==> (style.TextAllCaps |> Option.orElse (styleTextOn style (fun text -> text.AllCaps)))
        "StyleSmallCaps" ==> (style.TextSmallCaps |> Option.orElse (styleTextOn style (fun text -> text.SmallCaps)))
        "StyleStrike" ==> (style.TextStrike |> Option.orElse (styleTextOn style (fun text -> text.Strike)))
        "StyleVertAlign" ==> (style.Text |> Option.bind (fun text -> text.VertAlign))
        "StyleFontSize" ==> (style.Text
                            |> Option.bind (fun text ->
                                if ExplicitFormat.has ExplicitFormat.FontSize text || text.FontSize <> TextBuffer.defaultStyle.FontSize then Some text.FontSize
                                else None))
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
        // A shape without a fill / line says so (NoFill / NoLine); a text box that
        // says nothing about them is one made in the editor, white with a grey outline.
        let noFill = boolOrDefault false (item?NoFill)
        let noLine = boolOrDefault false (item?NoLine)
        TextBox {
            defaultTextBox with
                Id = nextImportedBlockId ()
                Content = cellParagraphs
                X = optionalNumberAsFloat (item?X)
                Y = optionalNumberAsFloat (item?Y)
                PositionHRelativeFrom = optionalString (item?PositionHRelativeFrom) |> Option.orElse (optionalString (item?TextBoxPositionHRelativeFrom))
                PositionVRelativeFrom = optionalString (item?PositionVRelativeFrom) |> Option.orElse (optionalString (item?TextBoxPositionVRelativeFrom))
                Width = optionalNumberAsFloat (item?Width) |> Option.defaultValue 220.0
                Height = optionalNumberAsFloat (item?Height) |> Option.defaultValue 72.0
                FillColor = if noFill then None else optionalString (item?FillColor) |> Option.orElse (Some "#FFFFFF")
                BorderColor = if noLine then None else optionalString (item?BorderColor) |> Option.orElse (Some "#808080")
                BorderWidth = optionalNumberAsFloat (item?BorderWidth) |> Option.orElse (Some 1.0)
                RelativeHeight = optionalString (item?RelativeHeight) |> Option.orElse (optionalString (item?TextBoxRelativeHeight))
                Geometry = geometryFromEditor item
                Gradient = gradientFromEditor (item?Gradient)
                Rotation = optionalNumberAsFloat (item?Rotation) |> Option.defaultValue 0.0
                Inline = boolOrDefault false (item?Inline)
                Wrap = optionalString (item?Wrap)
                PositionHAlign = optionalString (item?PositionHAlign)
                PositionVAlign = optionalString (item?PositionVAlign)
                BehindDoc = optionalBool (item?BehindDoc)
                AllowOverlap = optionalBool (item?AllowOverlap)
                LayoutInCell = optionalBool (item?LayoutInCell)
                DistanceTop = optionalNumberAsFloat (item?DistanceTop)
                DistanceBottom = optionalNumberAsFloat (item?DistanceBottom)
                DistanceLeft = optionalNumberAsFloat (item?DistanceLeft)
                DistanceRight = optionalNumberAsFloat (item?DistanceRight)
                BorderDash = optionalString (item?BorderDash)
                Name = optionalString (item?Name)
                InsetLeft = optionalNumberAsFloat (item?InsetLeft)
                InsetTop = optionalNumberAsFloat (item?InsetTop)
                InsetRight = optionalNumberAsFloat (item?InsetRight)
                InsetBottom = optionalNumberAsFloat (item?InsetBottom)
                VerticalAnchor = optionalString (item?VerticalAnchor)
                AutoFit = boolOrDefault false (item?AutoFit)
                NoTextWrap = boolOrDefault false (item?NoTextWrap)
                FitText = boolOrDefault false (item?FitText)
                Opacity = optionalNumberAsFloat (item?Opacity) |> Option.defaultValue 1.0
                SourceXml = optionalString (item?SourceXml)
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
                                    // (In a table inside a cell a merged cell's paragraphs are still handed
                                    // out to its rows: such a table is placed and painted row by row with
                                    // the cell that holds it.)
                                    | Table table -> Some { table with Rows = WordLayout.Table.distributeVerticalMergeContent table.Rows }
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
            // (The content of a vertically merged cell stays in the cell that starts the merge, as
            // in the file: the layout lays it over the rows the cell spans - and splits it with them
            // at a page break. It used to be handed out to the rows, a paragraph each.)
            Rows = tableRows
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
            FloatingXml = optionalString (item?FloatingXml)
            StyleId = optionalString (item?TableStyleId)
            Look =
                (let look = item?TableLook
                 if isNullOrUndefined look then None
                 else
                     Some { FirstRow = boolOrDefault false (look?FirstRow)
                            LastRow = boolOrDefault false (look?LastRow)
                            FirstColumn = boolOrDefault false (look?FirstColumn)
                            LastColumn = boolOrDefault false (look?LastColumn)
                            HorizontalBanding = boolOrDefault false (look?HorizontalBanding)
                            VerticalBanding = boolOrDefault false (look?VerticalBanding) })
        }

    | "Heading" ->
        let level = item?HeadingLevel |> intOrDefault 1 |> max 1 |> min 3
        Heading (level, paragraphFromEditor item)

    | _ ->
        Paragraph (paragraphFromEditor item)

/// The id of the block a shape is anchored in. The reader gives the anchor as
/// AnchorBlock, an index into the same item array.
let private blocksFromEditor (items: obj array) =
    let blocks = items |> Array.map blockFromEditor
    let idAt index =
        if index < 0 || index >= blocks.Length then None
        else
            match blocks.[index] with
            | Paragraph para
            | Heading (_, para) -> Some para.Id
            | Table table -> Some table.Id
            | TextBox _ -> None
    blocks
    |> Array.mapi (fun index block ->
        match block with
        | TextBox textBox ->
            match optionalInt (items.[index]?AnchorBlock) |> Option.bind idAt with
            | Some anchorId -> TextBox { textBox with AnchorBlockId = Some anchorId }
            | None -> block
        | _ -> block)
    |> Array.toList

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
        HeaderReferences = refsFromEditor (value?HeaderReferences)
        FooterReferences = refsFromEditor (value?FooterReferences)
        PageNumberStart = optionalInt (value?PageNumberStart)
        PageNumberFormat = optionalString (value?PageNumberFormat)
    }

/// Import EditorParagraphs from Office.Formats into Word document
let importFromEditor (editorParagraphs: obj array) : Document =
    // Every paragraph of the file is kept, empty ones at the end of the body
    // included: they are content (Word lays them out, and writes them back).
    let blocks = blocksFromEditor editorParagraphs
    let styles = stylesFromEditor editorParagraphs

    // Create document with all blocks
    let doc = { Blocks = blocks; Selection = None; PageSetup = defaultPageSetup; Styles = styles; Headers = []; Footers = []; HeaderParts = []; FooterParts = []; Package = emptyPackageData }

    // Set cursor at start of first block (paragraph only, skip tables for now)
    match blocks |> List.tryFind (function | Paragraph _ -> true | _ -> false) with
    | Some block ->
        let para = getParagraph block
        { doc with Selection = Some { Anchor = { BlockId = para.Id; Offset = 0 }; Focus = { BlockId = para.Id; Offset = 0 } } }
    | None ->
        doc

let importFromEditorWithStyles (editorParagraphs: obj array) (styleCatalog: obj array) : Document =
    let doc = importFromEditor editorParagraphs
    { doc with Styles = mergeCatalogAndInferredStyles styleCatalog doc.Styles }

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

let private noteBodiesFromEditor (noteType: string) (items: obj array) : NoteBody list =
    items
    |> Array.choose (fun item ->
        optionalString (item?Id)
        |> Option.map (fun id ->
            { NoteType = noteType
              Id = id
              Kind = optionalString (item?Kind)
              HasMark = item?HasMark |> boolOrDefault true
              MarkStyle = optionalString (item?MarkStyle)
              Text = item?Text |> stringOrDefault ""
              Blocks = blocksFromEditor (optionalArray (item?Blocks)) }))
    |> Array.toList

let private commentBodiesFromEditor (items: obj array) : CommentBody list =
    items
    |> Array.choose (fun item ->
        optionalString (item?Id)
        |> Option.map (fun id ->
            { Id = id
              Author = optionalString (item?Author)
              Date = optionalString (item?Date)
              Initials = optionalString (item?Initials)
              Text = item?Text |> stringOrDefault ""
              Blocks = blocksFromEditor (optionalArray (item?Blocks)) }))
    |> Array.toList

/// The reader's comment reference run only carries the comment id; give it the
/// text, author and date of the comment so the comments pane can show them.
let private withCommentDetails (comments: CommentBody list) (doc: Document) : Document =
    if comments.IsEmpty then doc
    else
        let byId = comments |> List.map (fun comment -> comment.Id, comment) |> Map.ofList
        let fill (para: ParagraphBlock) =
            let needsFill =
                para.Buffer.Runs
                |> List.exists (function
                    | Text run -> run.Style.CommentId.IsSome && run.Style.CommentText.IsNone
                    | Image _ -> false)
            if not needsFill then para
            else
                let runs =
                    para.Buffer.Runs
                    |> List.map (function
                        | Text run when run.Style.CommentId.IsSome && run.Style.CommentText.IsNone ->
                            match run.Style.CommentId |> Option.bind (fun id -> Map.tryFind id byId) with
                            | Some comment ->
                                Text { run with
                                        Style = { run.Style with
                                                    CommentText = Some comment.Text
                                                    CommentAuthor = run.Style.CommentAuthor |> Option.orElse comment.Author
                                                    CommentDate = run.Style.CommentDate |> Option.orElse comment.Date } }
                            | None -> Text run
                        | other -> other)
                { para with Buffer = { para.Buffer with Runs = runs } }
        let mapPart (part: HeaderFooterPart) = { part with Blocks = mapParagraphsOfBlocks fill part.Blocks }
        { doc with
            Blocks = mapParagraphsOfBlocks fill doc.Blocks
            HeaderParts = doc.HeaderParts |> List.map mapPart
            FooterParts = doc.FooterParts |> List.map mapPart }

let private convertEditorPackage (package: obj) : Document =
    let paragraphs = optionalArray (package?Paragraphs)
    let styles = optionalArray (package?Styles)
    let headerParts = optionalArray (package?HeaderParts)
    let footerParts = optionalArray (package?FooterParts)
    let doc = importFromEditorWithHeaderFooterParts paragraphs headerParts footerParts
    let doc =
        { doc with Styles = mergeCatalogAndInferredStyles styles doc.Styles }
    let comments = commentBodiesFromEditor (optionalArray (package?Comments))
    let doc =
        { doc with
            Package =
                { doc.Package with
                    Notes =
                        noteBodiesFromEditor "footnote" (optionalArray (package?Footnotes))
                        @ noteBodiesFromEditor "endnote" (optionalArray (package?Endnotes))
                    Comments = comments
                    Preserved = (if isNullOrUndefined package?Preserved then null else package?Preserved)
                    ImportedStyles = doc.Styles } }
        |> withCommentDetails comments
    let doc = if isNullOrUndefined package?PageSetup then doc else { doc with PageSetup = pageSetupFromEditor package?PageSetup }
    // What the document was when it was read (see PackageData.Imported).
    if isNull doc.Package.Preserved then doc else { doc with Package = { doc.Package with Imported = box doc } }

[<Emit("new WeakMap()")>]
let private newPackageMap () : obj = jsNative
[<Emit("((typeof $1 === 'object' && $1 !== null) ? $0.get($1) : undefined)")>]
let private packageMapGet (_map: obj) (_package: obj) : Document = jsNative
[<Emit("((typeof $1 === 'object' && $1 !== null) ? $0.set($1, $2) : undefined)")>]
let private packageMapSet (_map: obj) (_package: obj) (_doc: Document) : unit = jsNative
let private convertedPackages = newPackageMap ()

// ---- a package that is still being read ---------------------------------------
// Office.Formats' openProgressive hands a package over with the first blocks
// of the body in Paragraphs and reads the rest in slices (its Progressive
// member; the text boxes come last, when the body is read). Its document is
// made as the items arrive: the styles they bring, the comments' details,
// the anchors of the text boxes - what convertEditorPackage does at once -
// and, the package read to its end, it is the document convertEditorPackage
// makes of the complete package.

/// The document as it was imported (PackageData.Imported) of a package that
/// is still being read: what was read so far, the whole of it at the end.
type ImportedDocument(initial: Document) =
    member val Document : Document = initial with get, set

/// What PackageData.Imported holds: the document that was imported.
let importedDocumentOf (value: obj) : Document option =
    if isNull value then None
    else
        match value with
        | :? ImportedDocument as imported -> Some imported.Document
        | _ -> Some (unbox<Document> value)

/// More of a package was read.
type ImportStep = {
    /// The blocks that were read, in their order (at the end: the text boxes anchored in the body).
    Appended: Block list
    /// What was read, given to a document made of what was read before - the
    /// imported one, or one edited since: the blocks after its blocks, the
    /// styles they brought after its styles. The same blocks (and styles)
    /// always get the same extended list, so documents that shared their
    /// content before share it after.
    Extend: Document -> Document
    /// The package is read to its end.
    Complete: bool
    /// The package cannot be read to its end (Office.Formats' Progressive.Failed):
    /// why. This is its last step - the document is what was read.
    Failed: string option
    /// The font faces of the blocks that were read, once something asked for them (null until
    /// then; see Viewer.facesOfStep): they are found once for all who follow the package.
    mutable Faces: obj
}

/// The document of a package that is being read (see above).
type ProgressiveImport = {
    Package: obj
    /// What is read so far, as a document.
    mutable Document: Document
    mutable Complete: bool
    /// Why the package could not be read to its end, when it could not: the document stays
    /// what was read (and cannot be saved, see PackageData.Imported).
    mutable Failed: string option
    /// Told after each step: those who look at what the step brought before anything is done
    /// with it (the font faces it needs, see Viewer.onFontsDiscovered), then those who show it.
    FirstListeners: ResizeArray<ImportStep -> unit>
    Listeners: ResizeArray<ImportStep -> unit>
    // (Kept between the steps.)
    StyleList: ResizeArray<NamedStyle>
    StylesSeen: Dictionary<string, Dictionary<string, unit>>
    /// The id of the block of each item of the body ("": an item that has none).
    ItemBlockIds: ResizeArray<string>
    CommentBodies: CommentBody list
    Imported: ImportedDocument
    mutable Consumed: int
    /// The package's count of the changes of its page setup, as last looked at.
    mutable PageSetupVersion: int
    /// How many editors and viewers show the document while it is read (see holdProgressiveImport).
    mutable Holders: int
}

[<Emit("($0 != null && typeof $0 === 'object' && $0.Progressive != null)")>]
let private hasProgressive (_package: obj) : bool = jsNative
[<Emit("$0.get($1)")>]
let private importMapGet (_map: obj) (_package: obj) : ProgressiveImport = jsNative
[<Emit("$0.set($1, $2)")>]
let private importMapSet (_map: obj) (_package: obj) (_state: ProgressiveImport) : unit = jsNative
let private progressiveImports = newPackageMap ()

/// The details of the comments for the comment references of blocks (see withCommentDetails).
let private withCommentDetailsOfBlocks (comments: CommentBody list) (blocks: Block list) : Block list =
    if comments.IsEmpty then blocks
    else ({ Document.empty with Blocks = blocks; HeaderParts = []; FooterParts = [] } |> withCommentDetails comments).Blocks

/// What a step gives a document (ImportStep.Extend).
let private extension (appended: Block list) (stylesBefore: NamedStyle list) (stylesAfter: NamedStyle list)
                      (pageSetupBefore: PageSetup) (pageSetupAfter: PageSetup) (selection: Selection option) : Document -> Document =
    let addedStyles =
        if obj.ReferenceEquals(stylesBefore, stylesAfter) then []
        else stylesAfter |> List.skip (List.length stylesBefore)
    let knownBlocks = ResizeArray<Block list * Block list>()
    let knownStyles = ResizeArray<NamedStyle list * NamedStyle list>()
    let knownPackages = ResizeArray<PackageData * PackageData>()
    // (Used here as well as in the function below: a list that is named once, inside the
    // function, is made anew by the JavaScript build at every call of it - and nothing
    // that was extended would be found again.)
    knownBlocks.Clear()
    knownStyles.Clear()
    knownPackages.Clear()
    let extended (known: ResizeArray<'T * 'T>) (value: 'T) (make: unit -> 'T) : 'T =
        let mutable found = Unchecked.defaultof<'T>
        let mutable has = false
        let mutable i = 0
        while not has && i < known.Count do
            let before, after = known.[i]
            // (What was extended, or the extended one itself: a document is given a step once.)
            if obj.ReferenceEquals(before, value) || obj.ReferenceEquals(after, value) then
                found <- after
                has <- true
            i <- i + 1
        if has then found
        else
            let made = make ()
            known.Add ((value, made))
            made
    fun (doc: Document) ->
        let blocks = if appended.IsEmpty then doc.Blocks else extended knownBlocks doc.Blocks (fun () -> doc.Blocks @ appended)
        let styles =
            if addedStyles.IsEmpty then doc.Styles
            elif obj.ReferenceEquals(doc.Styles, stylesBefore) then stylesAfter
            else extended knownStyles doc.Styles (fun () -> doc.Styles @ addedStyles)
        // (The style catalog as imported: with the styles the blocks brought.)
        let package =
            if addedStyles.IsEmpty then doc.Package
            else extended knownPackages doc.Package (fun () -> { doc.Package with ImportedStyles = stylesAfter })
        let pageSetup = if obj.ReferenceEquals(doc.PageSetup, pageSetupBefore) then pageSetupAfter else doc.PageSetup
        let selection = match doc.Selection, selection with | None, Some _ -> selection | current, _ -> current
        if obj.ReferenceEquals(blocks, doc.Blocks) && obj.ReferenceEquals(styles, doc.Styles) && obj.ReferenceEquals(package, doc.Package)
           && obj.ReferenceEquals(pageSetup, doc.PageSetup) && obj.ReferenceEquals(selection, doc.Selection) then doc
        else { doc with Blocks = blocks; Styles = styles; Package = package; PageSetup = pageSetup; Selection = selection }

/// Take the items of the package that the document does not have yet.
let private consumeProgressive (state: ProgressiveImport) : unit =
    let items = optionalArray (state.Package?Paragraphs)
    let complete = unbox<bool> (state.Package?Progressive?IsComplete)
    let failed =
        let flag : obj = state.Package?Progressive?Failed
        if state.Failed.IsNone && not complete && not (isNullOrUndefined flag) && unbox<bool> flag then
            let error : obj = state.Package?Progressive?Error
            Some (if isNullOrUndefined error then "The document could not be read to its end" else string error)
        else None
    if items.Length > state.Consumed || (complete && not state.Complete) || failed.IsSome || (state.Package?Progressive?PageSetupVersion |> intOrDefault 0) <> state.PageSetupVersion then
        let stylesBefore = state.Document.Styles
        let styleCount = state.StyleList.Count
        let appended = ResizeArray<Block>()
        for index in state.Consumed .. items.Length - 1 do
            let item = items.[index]
            collectStylesFromItem state.StyleList state.StylesSeen item
            match blockFromEditor item with
            | TextBox textBox ->
                // (The reader gives the anchor as an index into the items.)
                let anchor =
                    optionalInt (item?AnchorBlock)
                    |> Option.bind (fun at -> if at >= 0 && at < state.ItemBlockIds.Count && state.ItemBlockIds.[at] <> "" then Some state.ItemBlockIds.[at] else None)
                state.ItemBlockIds.Add ""
                appended.Add (match anchor with Some anchorId -> TextBox { textBox with AnchorBlockId = Some anchorId } | None -> TextBox textBox)
            | Paragraph para
            | Heading (_, para) as block ->
                state.ItemBlockIds.Add para.Id
                appended.Add block
            | Table table as block ->
                state.ItemBlockIds.Add table.Id
                appended.Add block
        state.Consumed <- items.Length
        let blocks = withCommentDetailsOfBlocks state.CommentBodies (List.ofSeq appended)
        let stylesAfter = if state.StyleList.Count = styleCount then stylesBefore else List.ofSeq state.StyleList
        // The caret a document is imported with: at the start of its first paragraph.
        let selection =
            if state.Document.Selection.IsSome then None
            else
                blocks
                |> List.tryPick (function Paragraph para -> Some para | _ -> None)
                |> Option.map (fun para -> { Anchor = { BlockId = para.Id; Offset = 0 }; Focus = { BlockId = para.Id; Offset = 0 } })
        let pageSetupBefore = state.Document.PageSetup
        // (Until the body is read the package states the page setup of the section the blocks
        // read so far end in; it says when that is another one.)
        let pageSetupVersion = state.Package?Progressive?PageSetupVersion |> intOrDefault 0
        let pageSetupAfter =
            if pageSetupVersion <> state.PageSetupVersion && not (isNullOrUndefined state.Package?PageSetup) then
                state.PageSetupVersion <- pageSetupVersion
                let read = pageSetupFromEditor state.Package?PageSetup
                if read = pageSetupBefore then pageSetupBefore else read
            else pageSetupBefore
        let extend = extension blocks stylesBefore stylesAfter pageSetupBefore pageSetupAfter selection
        let document = extend state.Document
        state.Document <- document
        state.Complete <- complete
        if failed.IsSome then state.Failed <- failed
        // What the document was when it was read (see PackageData.Imported).
        state.Imported.Document <- { document with Package = { document.Package with Imported = null } }
        let step = { Appended = blocks; Extend = extend; Complete = complete; Failed = failed; Faces = null }
        for listener in state.FirstListeners.ToArray() do
            if state.FirstListeners.Contains listener then listener step
        for listener in state.Listeners.ToArray() do
            if state.Listeners.Contains listener then listener step

let private beginProgressive (package: obj) : ProgressiveImport =
    let headers = headerFooterPartsFromEditor (optionalArray (package?HeaderParts))
    let footers = headerFooterPartsFromEditor (optionalArray (package?FooterParts))
    let comments = commentBodiesFromEditor (optionalArray (package?Comments))
    let styleList = ResizeArray<NamedStyle>()
    let stylesSeen = Dictionary<string, Dictionary<string, unit>>()
    collectStylesFromCatalog styleList stylesSeen (optionalArray (package?Styles))
    let styles = List.ofSeq styleList
    let imported = ImportedDocument(Document.empty)
    let doc =
        { Blocks = []
          Selection = None
          PageSetup = (if isNullOrUndefined package?PageSetup then defaultPageSetup else pageSetupFromEditor package?PageSetup)
          Styles = styles
          Headers = headers |> List.tryHead |> Option.map (fun part -> paragraphsFromBlocks part.Blocks) |> Option.defaultValue []
          Footers = footers |> List.tryHead |> Option.map (fun part -> paragraphsFromBlocks part.Blocks) |> Option.defaultValue []
          HeaderParts = headers
          FooterParts = footers
          Package =
            { emptyPackageData with
                Notes =
                    noteBodiesFromEditor "footnote" (optionalArray (package?Footnotes))
                    @ noteBodiesFromEditor "endnote" (optionalArray (package?Endnotes))
                Comments = comments
                Preserved = (if isNullOrUndefined package?Preserved then null else package?Preserved)
                ImportedStyles = styles
                Imported = box imported } }
        |> withCommentDetails comments
    let state =
        { Package = package
          Document = doc
          Complete = false
          Failed = None
          FirstListeners = ResizeArray()
          Listeners = ResizeArray()
          StyleList = styleList
          StylesSeen = stylesSeen
          ItemBlockIds = ResizeArray()
          CommentBodies = comments
          Imported = imported
          Consumed = 0
          PageSetupVersion = package?Progressive?PageSetupVersion |> intOrDefault 0
          Holders = 0 }
    consumeProgressive state
    package?Progressive?subscribe (System.Func<int, int, bool, unit>(fun _ _ _ -> consumeProgressive state)) |> ignore
    state

/// The document of a package that is still being read: None for a package
/// that is read (forEditorPackage's, or one opened front first that was read
/// to its end before anything asked for its document).
let progressiveImportOf (package: obj) : ProgressiveImport option =
    if not (hasProgressive package) then None
    else
        let known = importMapGet progressiveImports package
        if not (isNull (box known)) then Some known
        elif unbox<bool> (package?Progressive?IsComplete) then None
        else
            let state = beginProgressive package
            importMapSet progressiveImports package state
            Some state

/// Read the rest of the package now. (One that cannot be read to its end: as far as it can
/// be - `state.Failed` says so afterwards.)
let completeProgressiveImport (state: ProgressiveImport) : unit =
    if not state.Complete && state.Failed.IsNone then
        state.Package?Progressive?complete () |> ignore
        consumeProgressive state

/// Stop reading the package: its document stays what was read, as that of a
/// package that failed (Office.Formats' Progressive.cancel; the listeners are
/// told with ImportStep.Failed). False when there was nothing to cancel.
let cancelProgressiveImport (state: ProgressiveImport) : bool =
    if state.Complete || state.Failed.IsSome then false
    else
        let cancel : obj = state.Package?Progressive?cancel
        if isNullOrUndefined cancel then false
        else
            let cancelled = unbox<bool> (state.Package?Progressive?cancel ())
            consumeProgressive state
            cancelled

/// An editor or a viewer shows the document while it is read ...
let holdProgressiveImport (state: ProgressiveImport) : unit =
    state.Holders <- state.Holders + 1

/// ... and is done with it (another document was loaded, it was disposed of):
/// when nothing shows a package that is still being read, the reading is
/// cancelled - no slices for a document nobody looks at, and the package lets
/// go of the file it reads.
let releaseProgressiveImport (state: ProgressiveImport) : unit =
    if state.Holders > 0 then
        state.Holders <- state.Holders - 1
        if state.Holders = 0 then cancelProgressiveImport state |> ignore

/// The document of an editor package. A package object is converted once:
/// a host that asks for a package's fonts (getRequiredFonts) and then loads
/// it into an editor or a viewer pays for one conversion, not two. (A
/// Document is immutable; the package object is not expected to change.)
/// For a package that is still being read (Office.Formats' openProgressive)
/// it is the document of what is read so far - see progressiveImportOf.
let importFromEditorPackage (package: obj) : Document =
    match progressiveImportOf package with
    | Some state -> state.Document
    | None ->
        let known = packageMapGet convertedPackages package
        if not (isNull (box known)) then known
        else
            let doc = convertEditorPackage package
            packageMapSet convertedPackages package doc
            doc

/// Export Word document to EditorParagraphs for Office.Formats
/// Reader px for a run with no size anywhere (11pt), which the model holds as
/// the unset size TextBuffer.defaultStyle.FontSize.
let private unsetFontSizePixels = 11.0 * 96.0 / 72.0

/// `context` is the style context of the document being exported; with it,
/// formatting a run merely inherits from its paragraph / character style is
/// left out of the file (ExplicitX = null) instead of being written on every run.
let private paragraphToEditor (context: WordStyles.DocumentStyleContext option) itemType headingLevel (para: ParagraphBlock) =
    let buffer = para.Buffer
    let metadata = para.Metadata
    // What the document's own styles give this paragraph (not the editor's
    // built-in look, which the file does not hold).
    let namedStyle =
        context
        |> Option.bind (fun ctx ->
            metadata.Style
            |> Option.orElse ctx.DefaultParagraphStyleId
            |> Option.bind (fun styleId -> Map.tryFind styleId ctx.ResolvedParagraphStyles))
    let fromStyles (pick: NamedStyle -> float option) =
        namedStyle
        |> Option.bind pick
        |> Option.orElseWith (fun () -> context |> Option.bind (fun ctx -> ctx.Defaults) |> Option.bind pick)
    // Direct formatting to write: a value the paragraph states, or one that
    // neither the styles nor the list level produce (set in the editor, or
    // resolved by the reader from something the model has no style for). An
    // unset value and zero are the same paragraph.
    let explicitMeasure (flag: int) (value: float option) (inherited: float option) : float option =
        match value with
        | None -> None
        | Some v when ParagraphExplicit.has flag metadata -> Some v
        | Some v ->
            match inherited with
            | Some i when abs (i - v) < 0.01 -> None
            | None when abs v < 0.01 -> None
            | _ -> Some v
    let levelIndents = if metadata.ListStyle <> NoList then metadata.ListLevelIndents else None
    let explicitIndent (flag: int) (value: float option) (styled: float option) (fromLevel: float option option) =
        match value, fromLevel with
        // An indent the paragraph does not state and its list level has is the level's.
        | Some v, Some (Some level) when not (ParagraphExplicit.has flag metadata) && abs (v - level) < 0.01 -> None
        | _ -> explicitMeasure flag value styled
    // The style the paragraph resolves to. An editor-made heading without a
    // style id is styled by the built-in heading look, which the file's
    // "HeadingN" style need not match, so nothing is treated as inherited there.
    let paragraphStyle =
        match context with
        | Some ctx when not (itemType = "Heading" && para.Metadata.Style.IsNone) ->
            let block = if itemType = "Heading" then Heading (headingLevel, para) else Paragraph para
            Some (ctx, WordStyles.forBlockInContext ctx block)
        | _ -> None
    let inheritedStyle (runStyle: TextStyle) : TextStyle option =
        paragraphStyle
        |> Option.map (fun (ctx, style) -> WordStyles.effectiveTextStyleInContext ctx style { TextBuffer.defaultStyle with Style = runStyle.Style })
    // What to write as direct formatting: a property set on the run (even
    // "off"), or a value that the styles do not produce (formatting the reader
    // resolved from something the model has no style for, e.g. a table style).
    let explicitToggle flag (pick: TextStyle -> bool) (inherited: TextStyle option) (style: TextStyle) : bool option =
        if ExplicitFormat.has flag style then Some (pick style)
        elif pick style && not (inherited |> Option.exists pick) then Some true
        else None
    // With exact / at-least line spacing the writer pins the paragraph font
    // size on the paragraph mark, and the reader hands the mark's size to runs
    // without one; every run then states its size so it cannot pick that up.
    // (The writer pins the mark size only for paragraph objects that do not
    // tell their own formatting from the inherited; the ones exported here do.)
    let pinsMarkSize = false
    let explicitFontSize (inherited: TextStyle option) (style: TextStyle) : float option =
        if ExplicitFormat.has ExplicitFormat.FontSize style then Some style.FontSize
        elif style.FontSize = TextBuffer.defaultStyle.FontSize then
            if not pinsMarkSize then None
            else
                match inherited with
                | Some inheritedStyle when inheritedStyle.FontSize <> TextBuffer.defaultStyle.FontSize -> Some inheritedStyle.FontSize
                | _ -> Some unsetFontSizePixels
        elif pinsMarkSize then Some style.FontSize
        else
            match inherited with
            | Some inheritedStyle ->
                let inheritedSize =
                    if inheritedStyle.FontSize = TextBuffer.defaultStyle.FontSize then unsetFontSizePixels else inheritedStyle.FontSize
                if abs (inheritedSize - style.FontSize) < 0.01 then None else Some style.FontSize
            | None -> Some style.FontSize
    let explicitAlignment =
        let metadata = para.Metadata
        if metadata.ExplicitAlignment then Some metadata.Alignment
        elif metadata.Alignment = Left then None
        else
            match paragraphStyle with
            | Some (_, style) when style.Alignment = Some metadata.Alignment -> None
            | _ -> Some metadata.Alignment

    let runs =
        buffer.Runs
        |> List.choose (function
            | Text tr ->
                let text = Rope.slice tr.Start tr.Length buffer.Rope
                let inherited = inheritedStyle tr.Style
                Some (createObj [
                    "Text" ==> text
                    "HasExplicit" ==> true
                    "ExplicitBold" ==> explicitToggle ExplicitFormat.Bold (fun style -> style.Bold) inherited tr.Style
                    "ExplicitItalic" ==> explicitToggle ExplicitFormat.Italic (fun style -> style.Italic) inherited tr.Style
                    "ExplicitUnderline" ==> explicitToggle ExplicitFormat.Underline (fun style -> style.Underline) inherited tr.Style
                    "ExplicitAllCaps" ==> explicitToggle ExplicitFormat.AllCaps (fun style -> style.AllCaps) inherited tr.Style
                    "ExplicitSmallCaps" ==> explicitToggle ExplicitFormat.SmallCaps (fun style -> style.SmallCaps) inherited tr.Style
                    "ExplicitStrike" ==> explicitToggle ExplicitFormat.Strike (fun style -> style.Strike) inherited tr.Style
                    "ExplicitDoubleStrike" ==> explicitToggle ExplicitFormat.DoubleStrike (fun style -> style.DoubleStrike) inherited tr.Style
                    "ExplicitHidden" ==> explicitToggle ExplicitFormat.Hidden (fun style -> style.Hidden) inherited tr.Style
                    "ExplicitFontSize" ==> explicitFontSize inherited tr.Style
                    "ExplicitColorAuto" ==> (ExplicitFormat.has ExplicitFormat.ColorAuto tr.Style && tr.Style.Color.IsNone)
                    // The font is what the paragraph's and the run's styles give: not direct formatting.
                    "FontFamilyInherited" ==> (tr.Style.FontFamily.IsSome && inherited |> Option.exists (fun style -> style.FontFamily = tr.Style.FontFamily))
                    "Style" ==> tr.Style.Style
                    "Bold" ==> tr.Style.Bold
                    "Italic" ==> tr.Style.Italic
                    "Underline" ==> tr.Style.Underline
                    "AllCaps" ==> tr.Style.AllCaps
                    "SmallCaps" ==> tr.Style.SmallCaps
                    "Strike" ==> tr.Style.Strike
                    "DoubleStrike" ==> tr.Style.DoubleStrike
                    "Hidden" ==> tr.Style.Hidden
                    "UnderlineStyle" ==> tr.Style.UnderlineStyle
                    "UnderlineColor" ==> tr.Style.UnderlineColor
                    "Position" ==> tr.Style.Position
                    "EmphasisMark" ==> tr.Style.EmphasisMark
                    "FontSize" ==> tr.Style.FontSize
                    "CharacterSpacing" ==> tr.Style.CharacterSpacing
                    "FontFamily" ==> (tr.Style.FontFamily |> Option.defaultValue "Calibri")
                    "Color" ==> tr.Style.Color
                    "Highlight" ==> tr.Style.Highlight
                    "Hyperlink" ==> tr.Style.Hyperlink
                    "InactiveHyperlink" ==> tr.Style.InactiveHyperlink
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
                    "InsertedRevisionId" ==> (tr.Style.DeletedInsertion |> Option.bind (fun mark -> mark.Id))
                    "InsertedRevisionAuthor" ==> (tr.Style.DeletedInsertion |> Option.bind (fun mark -> mark.Author))
                    "InsertedRevisionDate" ==> (tr.Style.DeletedInsertion |> Option.bind (fun mark -> mark.Date))
                    "IsDeletedInsertion" ==> tr.Style.DeletedInsertion.IsSome
                    "RevisionMoveName" ==> tr.Style.MoveName
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
                    "ImagePositionHAlign" ==> img.PositionHAlign
                    "ImagePositionVAlign" ==> img.PositionVAlign
                    "ImageWrapText" ==> img.WrapText
                    "ImageWrapPolygon" ==> (if img.WrapPolygon.IsEmpty then null else box (img.WrapPolygon |> List.collect (fun (x, y) -> [ x; y ]) |> List.toArray))
                    "OpaqueXml" ==> (img.Opaque |> Option.map (fun opaque -> opaque.Xml))
                    "OpaqueKind" ==> (img.Opaque |> Option.map (fun opaque -> opaque.Kind))
                    "OpaqueParagraphLevel" ==> (img.Opaque |> Option.map (fun opaque -> opaque.ParagraphLevel))
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
        "HasExplicit" ==> true
        "ExplicitAlignment" ==> (explicitAlignment |> Option.map alignmentToString)
        "ListStyle" ==> listStyleToString para.Metadata.ListStyle
        "ListLevel" ==> para.Metadata.ListLevel
        // Without the list identity every numbered paragraph is saved into one shared decimal list.
        "ListId" ==> para.Metadata.ListId
        "ListNumberFormat" ==> para.Metadata.ListNumberFormat
        "ListLevelText" ==> para.Metadata.ListLevelText
        "ListStart" ==> para.Metadata.ListStart
        "RightTabPosition" ==> para.Metadata.RightTabPosition
        "RightTabLeader" ==> para.Metadata.RightTabLeader
        "TabStops" ==> (para.Metadata.TabStops
                        |> List.map (fun stop -> createObj [ "Position" ==> stop.Position; "Alignment" ==> stop.Alignment; "Leader" ==> stop.Leader ])
                        |> List.toArray)
        "Rtl" ==> para.Metadata.Rtl
        "IndentLeft" ==> para.Metadata.IndentLeft
        "IndentRight" ==> para.Metadata.IndentRight
        "FirstLineIndent" ==> para.Metadata.FirstLineIndent
        "SpacingBefore" ==> para.Metadata.SpacingBefore
        "SpacingAfter" ==> para.Metadata.SpacingAfter
        "ExplicitParagraphFormat" ==> true
        "ExplicitSpacingBefore" ==> explicitMeasure ParagraphExplicit.SpacingBefore metadata.SpacingBefore (fromStyles (fun style -> style.SpacingBefore))
        "ExplicitSpacingAfter" ==> explicitMeasure ParagraphExplicit.SpacingAfter metadata.SpacingAfter (fromStyles (fun style -> style.SpacingAfter))
        // Exact / at-least spacing is held relative to the font size here and in px on a
        // style; unless the paragraph states it, only "auto" spacing is compared.
        "ExplicitLineSpacing" ==>
            (match metadata.LineSpacing with
             | Some value when ParagraphExplicit.has ParagraphExplicit.LineSpacing metadata -> Some value
             | Some value ->
                 let styleRule = namedStyle |> Option.bind (fun style -> style.LineSpacingRule)
                 let isFixed (rule: string option) = rule = Some "exact" || rule = Some "atLeast"
                 match fromStyles (fun style -> style.LineSpacing) with
                 | Some inherited when not (isFixed metadata.LineSpacingRule) && not (isFixed styleRule) && abs (inherited - value) < 0.001 -> None
                 | Some _ when isFixed metadata.LineSpacingRule && metadata.LineSpacingRule = styleRule -> None
                 | _ -> Some value
             | None -> None)
        "ExplicitIndentLeft" ==> explicitIndent ParagraphExplicit.IndentLeft metadata.IndentLeft (fromStyles (fun style -> style.IndentLeft)) (levelIndents |> Option.map fst)
        "ExplicitIndentRight" ==> explicitMeasure ParagraphExplicit.IndentRight metadata.IndentRight (fromStyles (fun style -> style.IndentRight))
        "ExplicitFirstLineIndent" ==> explicitIndent ParagraphExplicit.FirstLineIndent metadata.FirstLineIndent (fromStyles (fun style -> style.FirstLineIndent)) (levelIndents |> Option.map snd)
        "ExplicitContextualSpacing" ==>
            (match metadata.ContextualSpacing with
             | Some value when ParagraphExplicit.has ParagraphExplicit.ContextualSpacing metadata -> Some value
             | _ -> None)
        "ExplicitBackgroundColor" ==>
            (match metadata.BackgroundColor with
             | Some color when ParagraphExplicit.has ParagraphExplicit.BackgroundColor metadata -> Some color
             | Some color when (namedStyle |> Option.bind (fun style -> style.BackgroundColor)) <> Some color -> Some color
             | _ -> None)
        "ExplicitTabs" ==> (ParagraphExplicit.has ParagraphExplicit.Tabs metadata)
        "HasMarkProperties" ==> metadata.HasMarkProperties
        "ExplicitMarkFontSize" ==> metadata.ExplicitMarkFontSize
        "MarkRevisionType" ==> (metadata.MarkRevision |> Option.map (fun revision -> revision.RevisionType))
        "MarkRevisionId" ==> (metadata.MarkRevision |> Option.bind (fun revision -> revision.Mark.Id))
        "MarkRevisionAuthor" ==> (metadata.MarkRevision |> Option.bind (fun revision -> revision.Mark.Author))
        "MarkRevisionDate" ==> (metadata.MarkRevision |> Option.bind (fun revision -> revision.Mark.Date))
        // The indents of this paragraph's list level, for the numbering definition.
        "ListLevelIndentLeft" ==> (levelIndents |> Option.bind fst)
        "ListLevelFirstLineIndent" ==> (levelIndents |> Option.bind snd)
        "HasListLevelIndents" ==> levelIndents.IsSome
        "StyleContextualSpacing" ==> para.Metadata.ContextualSpacing
        "LineSpacing" ==> para.Metadata.LineSpacing
        "LineSpacingRule" ==> para.Metadata.LineSpacingRule
        // Exact/at-least line spacing is stored relative to this size; the writer needs it to get the twips back.
        "ParagraphMarkFontSize" ==> para.Metadata.ParagraphMarkFontSize
        "WidowControl" ==> para.Metadata.WidowControl
        "KeepNext" ==> para.Metadata.KeepNext
        "KeepLines" ==> para.Metadata.KeepLines
        "BackgroundColor" ==> para.Metadata.BackgroundColor
        "ParagraphBorders" ==> paragraphBordersToEditor para.Metadata.ParagraphBorders
        "ParagraphBorderBefore" ==> para.Metadata.ParagraphBorderBefore
        "ParagraphBorderAfter" ==> para.Metadata.ParagraphBorderAfter
        "PageBreakBefore" ==> para.Metadata.PageBreakBefore
        "PageBreakBeforeDerived" ==> (para.Metadata.PageBreakBefore && ParagraphExplicit.has ParagraphExplicit.DerivedPageBreakBefore para.Metadata)
        "PageBreakAfter" ==> para.Metadata.PageBreakAfter
        "SectionBreak" ==> para.Metadata.SectionBreak
        "SectionPageNumberStart" ==> para.Metadata.SectionPageNumberStart
        "SectionPageNumberFormat" ==> para.Metadata.SectionPageNumberFormat
        "RenderedPageBreakBefore" ==> para.Metadata.RenderedPageBreakBefore
        "SectionPageSetup" ==> (para.Metadata.SectionPageSetup |> Option.map sectionPageSetupToEditor |> Option.toObj)
        "SectionColumns" ==> (para.Metadata.SectionColumns
                              |> Option.map (fun columns ->
                                  createObj [
                                      "Count" ==> columns.Count
                                      "Gap" ==> columns.Gap
                                      "Separator" ==> columns.Separator
                                      "Widths" ==> (columns.Widths |> List.map (fun (width, space) -> createObj [ "W" ==> width; "S" ==> space ]) |> List.toArray)
                                  ])
                              |> Option.toObj)
        "HeaderReferences" ==> refsToEditor para.Metadata.HeaderReferences
        "FooterReferences" ==> refsToEditor para.Metadata.FooterReferences
    ]

/// Inverse of WordLayout.Table.distributeVerticalMergeContent, which on import
/// spreads the paragraphs of a vertically merged cell over its continuation
/// cells so each row can be laid out. In the file all of that content belongs
/// to the w:vMerge="restart" cell: Word does not show what a continuation
/// cell contains, so exporting the distributed form made every paragraph after
/// the first disappear from a merged cell once the saved file was opened in Word.
let private gatherVerticalMergeContent (rows: TableRow list) : TableRow list =
    let cells = rows |> List.map (fun row -> List.toArray row.Cells) |> List.toArray
    let startColumns =
        rows
        |> List.map (fun row ->
            let mutable column = row.GridBefore
            row.Cells
            |> List.map (fun cell ->
                let start = column
                column <- column + max 1 cell.ColSpan
                start)
            |> List.toArray)
        |> List.toArray
    let hasMerge = cells |> Array.exists (Array.exists (fun cell -> cell.VerticalMerge = Some ContinueMerge))
    if not hasMerge then rows
    else
        let isBlank (para: ParagraphBlock) = TextBuffer.isWhiteSpace para.Buffer
        for rowIndex in 0 .. cells.Length - 1 do
            for cellIndex in 0 .. cells.[rowIndex].Length - 1 do
                let cell = cells.[rowIndex].[cellIndex]
                if cell.VerticalMerge = Some RestartMerge then
                    let column = startColumns.[rowIndex].[cellIndex]
                    let gathered = ResizeArray<ParagraphBlock>()
                    let mutable nextRow = rowIndex + 1
                    let mutable keepGoing = true
                    while keepGoing && nextRow < cells.Length do
                        match startColumns.[nextRow] |> Array.tryFindIndex (fun start -> start = column) with
                        | Some index when cells.[nextRow].[index].VerticalMerge = Some ContinueMerge ->
                            let continuation = cells.[nextRow].[index]
                            gathered.AddRange continuation.Content
                            let placeholder =
                                match continuation.Content with
                                | first :: _ -> [ { first with Buffer = TextBuffer.empty } ]
                                | [] -> []
                            cells.[nextRow].[index] <- { continuation with Content = placeholder }
                            nextRow <- nextRow + 1
                        | _ -> keepGoing <- false
                    // Blank paragraphs at the end are just the continuation cells' own
                    // (mandatory) empty paragraph; blank ones in between are real spacing.
                    while gathered.Count > 0 && isBlank gathered.[gathered.Count - 1] do
                        gathered.RemoveAt (gathered.Count - 1)
                    if gathered.Count > 0 then
                        cells.[rowIndex].[cellIndex] <- { cell with Content = cell.Content @ List.ofSeq gathered }
        rows |> List.mapi (fun rowIndex row -> { row with Cells = List.ofArray cells.[rowIndex] })

/// The shapes of the document that was read, by their ids, while a document
/// is exported for a save (null otherwise): a shape that is still what was
/// read is written as the element it was read from - its DrawingML and its
/// VML fallback as they were, whatever of them the model does not hold.
let mutable private importedTextBoxes : Dictionary<string, TextBoxBlock> = null

let private textBoxIsAsImported (textBox: TextBoxBlock) : bool =
    textBox.SourceXml.IsSome
    && not (isNull importedTextBoxes)
    && (match importedTextBoxes.TryGetValue textBox.Id with
        | true, original ->
            obj.ReferenceEquals(original, textBox)
            || (obj.ReferenceEquals(original.Content, textBox.Content) && { original with Content = [] } = { textBox with Content = [] })
        | false, _ -> false)

let rec private blockToEditor (context: WordStyles.DocumentStyleContext option) block =
    match block with
    | Paragraph para -> paragraphToEditor context "Paragraph" 0 para
    | Heading (level, para) -> paragraphToEditor context "Heading" level para
    | TextBox textBox ->
        let geometry, customPath, customW, customH =
            match textBox.Geometry with
            | GeoRect -> "rect", [||], 0.0, 0.0
            | GeoPreset preset -> preset, [||], 0.0, 0.0
            | GeoCustom (segs, coordW, coordH) ->
                let seg (op: string) (points: float list) = createObj [ "Op" ==> op; "P" ==> List.toArray points ]
                let path =
                    segs
                    |> List.map (function
                        | PMove (x, y) -> seg "M" [ x; y ]
                        | PLine (x, y) -> seg "L" [ x; y ]
                        | PCubic (c1x, c1y, c2x, c2y, x, y) -> seg "C" [ c1x; c1y; c2x; c2y; x; y ]
                        | PQuad (cx, cy, x, y) -> seg "Q" [ cx; cy; x; y ]
                        | PClose -> seg "Z" [])
                    |> List.toArray
                "custom", path, coordW, coordH
        createObj [
            "Type" ==> "TextBox"
            "Geometry" ==> geometry
            "CustomPath" ==> customPath
            "CustomW" ==> customW
            "CustomH" ==> customH
            "Gradient" ==> (textBox.Gradient
                            |> Option.map (fun gradient ->
                                createObj [
                                    "Stops" ==> (gradient.Stops |> List.map (fun stop -> createObj [ "Offset" ==> stop.Offset; "Color" ==> stop.Color ]) |> List.toArray)
                                    "Angle" ==> gradient.Angle
                                    "Radial" ==> gradient.Radial ])
                            |> Option.toObj)
            "Rotation" ==> textBox.Rotation
            "NoFill" ==> (textBox.FillColor.IsNone && textBox.Gradient.IsNone)
            "NoLine" ==> textBox.BorderColor.IsNone
            "Inline" ==> textBox.Inline
            "Wrap" ==> textBox.Wrap
            "PositionHAlign" ==> textBox.PositionHAlign
            "PositionVAlign" ==> textBox.PositionVAlign
            "BehindDoc" ==> textBox.BehindDoc
            "AllowOverlap" ==> textBox.AllowOverlap
            "LayoutInCell" ==> textBox.LayoutInCell
            "DistanceTop" ==> textBox.DistanceTop
            "DistanceBottom" ==> textBox.DistanceBottom
            "DistanceLeft" ==> textBox.DistanceLeft
            "DistanceRight" ==> textBox.DistanceRight
            "BorderDash" ==> textBox.BorderDash
            "Name" ==> textBox.Name
            "InsetLeft" ==> textBox.InsetLeft
            "InsetTop" ==> textBox.InsetTop
            "InsetRight" ==> textBox.InsetRight
            "InsetBottom" ==> textBox.InsetBottom
            "VerticalAnchor" ==> textBox.VerticalAnchor
            "AutoFit" ==> textBox.AutoFit
            "NoTextWrap" ==> textBox.NoTextWrap
            "FitText" ==> textBox.FitText
            "Opacity" ==> (if textBox.Opacity < 0.999 then Some textBox.Opacity else None)
            // (Its element as it was read, while the shape is what was read: see importedTextBoxes.)
            "SourceXml" ==> (if textBoxIsAsImported textBox then textBox.SourceXml else None)
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
            "Paragraphs" ==> (textBox.Content |> List.map (paragraphToEditor context "Paragraph" 0) |> List.toArray)
        ]
    | Table table ->
            let rows =
                gatherVerticalMergeContent table.Rows
                |> List.map (fun row ->
                    let cells =
                        row.Cells
                        |> List.map (fun cell ->
                            createObj [
                                "Paragraphs" ==> (cell.Content |> List.map (paragraphToEditor context "Paragraph" 0) |> List.toArray)
                                "NestedTables" ==> (cell.NestedTables |> List.map (Table >> blockToEditor context) |> List.toArray)
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
                                "PaddingTop" ==> canvasPixelsToDxa cell.PaddingTop
                                "PaddingRight" ==> canvasPixelsToDxa cell.PaddingRight
                                "PaddingBottom" ==> canvasPixelsToDxa cell.PaddingBottom
                                "PaddingLeft" ==> canvasPixelsToDxa cell.PaddingLeft
                            ])
                        |> List.toArray

                    createObj [
                        "Cells" ==> cells
                        "IsHeader" ==> row.IsHeader
                        "Height" ==> row.Height
                        "HeightRule" ==> row.HeightRule
                        "CanSplit" ==> row.CanSplit
                        "CellSpacing" ==> canvasPixelsToDxa row.CellSpacing
                        "PaddingTop" ==> canvasPixelsToDxa row.PaddingTop
                        "PaddingRight" ==> canvasPixelsToDxa row.PaddingRight
                        "PaddingBottom" ==> canvasPixelsToDxa row.PaddingBottom
                        "PaddingLeft" ==> canvasPixelsToDxa row.PaddingLeft
                    ])
                |> List.toArray

            createObj [
                "Type" ==> "Table"
                "Rows" ==> rows
                "TableStyleId" ==> table.StyleId
                "TableLook" ==> (table.Look
                                 |> Option.map (fun look ->
                                     createObj [
                                         "FirstRow" ==> look.FirstRow
                                         "LastRow" ==> look.LastRow
                                         "FirstColumn" ==> look.FirstColumn
                                         "LastColumn" ==> look.LastColumn
                                         "HorizontalBanding" ==> look.HorizontalBanding
                                         "VerticalBanding" ==> look.VerticalBanding ])
                                 |> Option.toObj)
                "PreferredWidth" ==> tableWidthToEditor table.PreferredWidth
                "PreferredWidthType" ==> tableWidthTypeToEditor table.PreferredWidth
                "TableAlignment" ==> table.Alignment
                "LayoutMode" ==> tableLayoutModeToString table.LayoutMode
                "GridColumns" ==> (table.GridColumns |> List.map (fun width -> width.Value |> Option.defaultValue 0.0) |> List.toArray)
                "Floating" ==> table.Floating
                "FloatingXml" ==> table.FloatingXml
                "Borders" ==> tableBordersToEditor table.Borders
                "BackgroundColor" ==> table.BackgroundColor
                "Shading" ==> shadingToEditor table.Shading
                "CellSpacing" ==> canvasPixelsToDxa table.CellSpacing
                "PaddingTop" ==> canvasPixelsToDxa table.PaddingTop
                "PaddingRight" ==> canvasPixelsToDxa table.PaddingRight
                "PaddingBottom" ==> canvasPixelsToDxa table.PaddingBottom
                "PaddingLeft" ==> canvasPixelsToDxa table.PaddingLeft
            ]

/// Blocks in the order they are written: a shape follows the block it is
/// anchored in (the writer puts it into that block), wherever it sits in the
/// block list. Imported shapes are kept at the end of the list, and writing
/// them there moved every shape of a document into its last paragraph.
let private blocksInAnchorOrder (blocks: Block list) : Block list =
    let idOf block =
        match block with
        | Paragraph para
        | Heading (_, para) -> Some para.Id
        | Table table -> Some table.Id
        | TextBox _ -> None
    let anchorIds = blocks |> List.choose idOf |> Set.ofList
    let isAnchored (textBox: TextBoxBlock) = textBox.AnchorBlockId |> Option.exists (fun id -> Set.contains id anchorIds)
    let anchored =
        blocks |> List.choose (function TextBox textBox when isAnchored textBox -> Some textBox | _ -> None)
    if anchored.IsEmpty then blocks
    else
        blocks
        |> List.collect (fun block ->
            match block with
            | TextBox textBox when isAnchored textBox -> []
            | _ ->
                match idOf block with
                | Some id -> block :: (anchored |> List.filter (fun textBox -> textBox.AnchorBlockId = Some id) |> List.map TextBox)
                | None -> [ block ])

let exportToEditor (doc: Document) : obj array =
    let context = Some (WordStyles.styleContext doc)
    blocksInAnchorOrder doc.Blocks
    |> List.map (blockToEditor context)
    |> List.toArray

let private headerFooterPartsToEditor context (parts: HeaderFooterPart list) : obj array =
    parts
    |> List.map (fun part ->
        createObj [
            "Type" ==> part.Kind
            "Kind" ==> part.Kind
            "RelationshipId" ==> part.RelationshipId
            "Paragraphs" ==> (blocksInAnchorOrder part.Blocks |> List.map (blockToEditor context) |> List.toArray)
            "Blocks" ==> (blocksInAnchorOrder part.Blocks |> List.map (blockToEditor context) |> List.toArray)
        ])
    |> List.toArray

/// Header / footer parts without their document: nothing is known to be
/// inherited from a style, so all formatting is written as direct formatting.
let exportHeaderFooterPartsToEditor (parts: HeaderFooterPart list) : obj array =
    headerFooterPartsToEditor None parts

/// Page setup for the editor package, including the section columns and
/// section start that the reader delivers (toEditorPageSetup) and
/// pageSetupFromEditor imports. Same field names and units (px) as the reader.
let exportPageSetupToEditor (setup: PageSetup) : obj =
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
        "ColumnCount" ==> (setup.Columns |> Option.map (fun columns -> columns.Count) |> Option.defaultValue 1)
        "ColumnSpace" ==> (setup.Columns |> Option.map (fun columns -> columns.Gap) |> Option.defaultValue 0.0)
        "ColumnSeparator" ==> (setup.Columns |> Option.map (fun columns -> columns.Separator) |> Option.defaultValue false)
        "ColumnWidths" ==> (setup.Columns
                            |> Option.map (fun columns -> columns.Widths |> List.map (fun (width, space) -> createObj [ "W" ==> width; "S" ==> space ]) |> List.toArray)
                            |> Option.defaultValue [||])
        "SectionStart" ==> setup.SectionStart
        "HeaderReferences" ==> refsToEditor setup.HeaderReferences
        "FooterReferences" ==> refsToEditor setup.FooterReferences
        "PageNumberStart" ==> setup.PageNumberStart
        "PageNumberFormat" ==> setup.PageNumberFormat
    ]

/// A note / comment body made of plain text: what the editor has for a note
/// it created, or for one whose text was changed after import.
let private plainBodyBlocks context (styleId: string option) (text: string) : obj array =
    let para = { Id = nextImportedBlockId (); Buffer = TextBuffer.ofString text; Metadata = { defaultMetadata with Style = styleId } }
    [| paragraphToEditor context "Paragraph" 0 para |]

/// Footnotes or endnotes for the editor package. A note is written from its
/// imported body (formatting, hyperlinks, tables) as long as the reference in
/// the text still carries the text it was imported with; otherwise from the
/// text on the reference. Notes nothing refers to any more are left out, the
/// separator notes are kept.
let private notesToEditor context (noteType: string) (doc: Document) : obj array =
    let bodies = doc.Package.Notes |> List.filter (fun note -> note.NoteType = noteType)
    let referenced =
        allParagraphs doc
        |> List.collect (fun para ->
            para.Buffer.Runs
            |> List.choose (function
                | Text run ->
                    match run.Style.NoteId with
                    | Some id when (run.Style.NoteType |> Option.defaultValue "footnote") = noteType -> Some (id, run.Style.NoteText)
                    | _ -> None
                | Image _ -> None))
        |> List.distinctBy fst
    let textStyle =
        let styleId = if noteType = "endnote" then "EndnoteText" else "FootnoteText"
        if doc.Styles |> List.exists (fun style -> style.Id = styleId && style.StyleType = "paragraph") then Some styleId else None
    let noteObject (id: string) (kind: string option) hasMark (markStyle: string option) (text: string) (blocks: obj array) =
        createObj [
            "Id" ==> id
            "Kind" ==> kind
            "HasMark" ==> hasMark
            "MarkStyle" ==> markStyle
            "Text" ==> text
            "Blocks" ==> blocks
        ]
    let rich (body: NoteBody) =
        noteObject body.Id body.Kind body.HasMark body.MarkStyle body.Text (body.Blocks |> List.map (blockToEditor context) |> List.toArray)
    [| for body in bodies do
           if body.Kind.IsSome then yield rich body
       for id, text in referenced do
           match bodies |> List.tryFind (fun body -> body.Id = id && body.Kind.IsNone) with
           | Some body when text.IsNone || text = Some body.Text -> yield rich body
           | Some body ->
               let text = text |> Option.defaultValue ""
               yield noteObject id None body.HasMark body.MarkStyle text (plainBodyBlocks context textStyle text)
           | None ->
               let text = text |> Option.defaultValue ""
               yield noteObject id None true None text (plainBodyBlocks context textStyle text) |]

let private commentsToEditor context (doc: Document) : obj array =
    let referenced =
        allParagraphs doc
        |> List.collect (fun para ->
            para.Buffer.Runs
            |> List.choose (function
                | Text run -> run.Style.CommentId |> Option.map (fun id -> id, run.Style)
                | Image _ -> None))
        |> List.distinctBy fst
    let commentObject (id: string) (author: string option) (date: string option) (initials: string option) (text: string) (blocks: obj array) =
        createObj [
            "Id" ==> id
            "Author" ==> author
            "Date" ==> date
            "Initials" ==> initials
            "Text" ==> text
            "Blocks" ==> blocks
        ]
    [| for id, style in referenced do
           match doc.Package.Comments |> List.tryFind (fun body -> body.Id = id) with
           | Some body when style.CommentText.IsNone || style.CommentText = Some body.Text ->
               yield commentObject id (style.CommentAuthor |> Option.orElse body.Author) (style.CommentDate |> Option.orElse body.Date) body.Initials body.Text
                         (body.Blocks |> List.map (blockToEditor context) |> List.toArray)
           | body ->
               let text = style.CommentText |> Option.defaultValue ""
               yield commentObject id style.CommentAuthor style.CommentDate (body |> Option.bind (fun value -> value.Initials)) text (plainBodyBlocks context None text) |]

/// The whole document as the package DocumentWriter.writeFromEditorPackage
/// takes: the inverse of importFromEditorPackage.
[<Emit("JSON.stringify($0)")>]
let private jsonOf (_value: obj) : string = jsNative

let private sameBlockObject (a: Block) (b: Block) =
    match a, b with
    | Paragraph x, Paragraph y -> obj.ReferenceEquals(x, y)
    | Heading (levelX, x), Heading (levelY, y) -> levelX = levelY && obj.ReferenceEquals(x, y)
    | Table x, Table y -> obj.ReferenceEquals(x, y)
    | TextBox x, TextBox y -> obj.ReferenceEquals(x, y)
    | _ -> false

let private blockIdOf (block: Block) =
    match block with
    | Paragraph para
    | Heading (_, para) -> para.Id
    | Table table -> table.Id
    | TextBox textBox -> textBox.Id

// ---- a changed block as a change of its text ----------------------------------
//
// A block of which only the text of some paragraphs was changed (typing,
// deleting, replacing, formatting a piece) is not written from the model: the
// writer makes the change in the XML the block was read from, so that
// whatever the model does not hold of it (bookmarks, complex fields, revision
// ids, properties and elements of later formats) stays as it was. Here the
// change is found: per paragraph, the one stretch of characters that is not
// what it was - a character being what it was when the character and its
// run's formatting are.

/// The paragraphs of a block that are paragraphs of the body's XML, in the
/// order they have there (a vertically merged cell's content in its first
/// cell, a nested table's where the table is). None for a text box.
let private storyParagraphs (block: Block) : ParagraphBlock list option =
    let rec ofTable (table: TableBlock) : ParagraphBlock list =
        gatherVerticalMergeContent table.Rows
        |> List.collect (fun row ->
            row.Cells
            |> List.collect (fun cell ->
                if cell.NestedTables.IsEmpty then cell.Content
                else
                    let count = cell.Content.Length
                    let anchors = cell.NestedTableAnchors |> List.toArray
                    let nested =
                        cell.NestedTables
                        |> List.mapi (fun i inner -> inner, (if i < anchors.Length then anchors.[i] else count) |> max 0 |> min count)
                        |> List.sortBy snd
                    let before index = nested |> List.filter (fun (_, anchor) -> anchor = index) |> List.collect (fun (inner, _) -> ofTable inner)
                    (cell.Content |> List.mapi (fun index para -> before index @ [ para ]) |> List.concat) @ before count))
    match block with
    | Paragraph para
    | Heading (_, para) -> Some [ para ]
    | Table table -> Some (ofTable table)
    | TextBox _ -> None

/// One entry per character: the run it is in.
let private runsByCharacter (buffer: TextBuffer.TextBuffer) : TextBuffer.Run array =
    let result = ResizeArray<TextBuffer.Run>()
    for run in buffer.Runs do
        match run with
        | TextBuffer.Text text -> for _ in 1 .. text.Length do result.Add run
        | TextBuffer.Image _ -> result.Add run
    result.ToArray()

/// The stretch of a paragraph that was changed, for the writer: where it
/// starts in the text that was read, how many characters of that text it
/// replaces, the text now there, whether it goes into the run before or
/// after it, and otherwise where a run is that is formatted as it is. None when the paragraph was changed in another way (its properties,
/// a picture), or not at all.
let private paragraphTextChange (context: WordStyles.DocumentStyleContext option) (index: int) (before: ParagraphBlock) (after: ParagraphBlock) : obj option =
    if obj.ReferenceEquals(before, after) || not (before.Metadata = after.Metadata) then None
    else
        let oldText = TextBuffer.toString before.Buffer
        let newText = TextBuffer.toString after.Buffer
        let oldRuns = runsByCharacter before.Buffer
        let newRuns = runsByCharacter after.Buffer
        if oldRuns.Length <> oldText.Length || newRuns.Length <> newText.Length then None
        else
            // (Characters of one run are compared once.)
            let mutable lastA = Unchecked.defaultof<TextBuffer.Run>
            let mutable lastB = Unchecked.defaultof<TextBuffer.Run>
            let mutable lastSame = false
            let sameRun (a: TextBuffer.Run) (b: TextBuffer.Run) =
                if obj.ReferenceEquals(a, b) then true
                elif obj.ReferenceEquals(a, lastA) && obj.ReferenceEquals(b, lastB) then lastSame
                else
                    let same =
                        match a, b with
                        // (The mark of a rendered page break aside: where the marks are is compared below.)
                        | TextBuffer.Text x, TextBuffer.Text y ->
                            obj.ReferenceEquals(x.Style, y.Style) || x.Style = y.Style
                            || (x.Style.RenderedPageBreakBefore <> y.Style.RenderedPageBreakBefore && { x.Style with RenderedPageBreakBefore = y.Style.RenderedPageBreakBefore } = y.Style)
                        | TextBuffer.Image x, TextBuffer.Image y -> { x with Start = 0 } = { y with Start = 0 }
                        | _ -> false
                    lastA <- a
                    lastB <- b
                    lastSame <- same
                    same
            let same (i: int) (j: int) = oldText.[i] = newText.[j] && sameRun oldRuns.[i] newRuns.[j]
            let limit = min oldText.Length newText.Length
            let mutable prefix = 0
            while prefix < limit && same prefix prefix do prefix <- prefix + 1
            let mutable suffix = 0
            while suffix < limit - prefix && same (oldText.Length - 1 - suffix) (newText.Length - 1 - suffix) do suffix <- suffix + 1
            let deleted = oldText.Length - prefix - suffix
            let insertedLength = newText.Length - prefix - suffix
            let isText (run: TextBuffer.Run) = match run with TextBuffer.Text _ -> true | TextBuffer.Image _ -> false
            let allText (runs: TextBuffer.Run array) (start: int) (count: int) =
                let mutable ok = true
                for i in start .. start + count - 1 do
                    if not (isText runs.[i]) then ok <- false
                ok
            // The marks of page breaks as they were rendered (each before the first character of
            // its run) are where they were: in the XML a mark stays before the text it is before.
            // Text typed just before a marked run has the mark instead of the run - it goes into
            // the run's w:t, after the mark.
            let marks (buffer: TextBuffer.TextBuffer) =
                buffer.Runs |> List.choose (function TextBuffer.Text run when run.Style.RenderedPageBreakBefore && run.Length > 0 -> Some run.Start | _ -> None)
            let marksKept =
                let before = marks before.Buffer
                let after = marks after.Buffer
                if before.IsEmpty && after.IsEmpty then true
                else
                    let moved = before |> List.filter (fun at -> at < prefix || at >= prefix + deleted) |> List.map (fun at -> if at < prefix then at else at - deleted + insertedLength)
                    let inside = before |> List.exists (fun at -> at >= prefix && at < prefix + deleted)
                    not inside
                    && (moved = after
                        || (insertedLength > 0 && List.contains (prefix + deleted) before
                            && (moved |> List.map (fun at -> if at = prefix + insertedLength then prefix else at)) = after))
            if deleted = 0 && insertedLength = 0 then None
            elif not marksKept then None
            elif not (allText oldRuns prefix deleted) || not (allText newRuns prefix insertedLength) then None
            else
                // Text typed into a run: it has the formatting of the character before or after it.
                let oneRun =
                    insertedLength > 0
                    && (let mutable one = true
                        for i in prefix + 1 .. prefix + insertedLength - 1 do
                            if not (sameRun newRuns.[prefix] newRuns.[i]) then one <- false
                        one)
                let joinsBefore = oneRun && prefix > 0 && sameRun oldRuns.[prefix - 1] newRuns.[prefix]
                let joinsAfter = oneRun && suffix > 0 && sameRun oldRuns.[prefix + deleted] newRuns.[prefix]
                // A run of the text that was read that is formatted as the new text is (links, fields and
                // marks aside): the properties a run of its own is written with.
                let plain =
                    oneRun
                    && (match newRuns.[prefix] with
                        | TextBuffer.Text run -> not (TextBuffer.hasAnnotationIdentity run.Style)
                        | TextBuffer.Image _ -> false)
                let witnesses =
                    if not plain then [||]
                    else
                        let formatted (run: TextBuffer.Run) =
                            match run, newRuns.[prefix] with
                            | TextBuffer.Text candidate, TextBuffer.Text typed ->
                                obj.ReferenceEquals(candidate.Style, typed.Style) || TextBuffer.withoutAnnotations candidate.Style = typed.Style
                            | _ -> false
                        // (A character of the run that is one of a w:t: not a tab, a break, the place of an object.)
                        let ordinary (c: char) = c >= ' ' && c <> '\uFFFC' && c <> '\u00AD' && c <> '\u2011' && c <> '\u2028'
                        let found = ResizeArray<int>()
                        let mutable lastRun = Unchecked.defaultof<TextBuffer.Run>
                        let mutable lastFormatted = false
                        let mutable taken = false
                        // (From the change on, then before it; one character of each run, a few runs.)
                        let order = seq { yield! seq { prefix + deleted .. oldRuns.Length - 1 }; yield! seq { prefix - 1 .. -1 .. 0 } }
                        for i in order do
                            if found.Count < 6 then
                                if not (obj.ReferenceEquals(oldRuns.[i], lastRun)) then
                                    lastRun <- oldRuns.[i]
                                    lastFormatted <- formatted oldRuns.[i]
                                    taken <- false
                                if lastFormatted && not taken && ordinary oldText.[i] then
                                    found.Add i
                                    taken <- true
                        found.ToArray()
                // A paragraph without text: what is typed into it has the paragraph mark's formatting.
                let markStyle = plain && witnesses.Length = 0 && not (oldRuns |> Array.exists isText)
                Some (createObj [
                        "Paragraph" ==> index
                        "Offset" ==> prefix
                        "Deleted" ==> deleted
                        "Inserted" ==> newText.Substring(prefix, insertedLength)
                        "JoinBefore" ==> joinsBefore
                        "JoinAfter" ==> joinsAfter
                        "RunsPlain" ==> plain
                        "Witnesses" ==> witnesses
                        "MarkStyle" ==> markStyle
                        "NewText" ==> newText ])

/// A paragraph whose text is what it was and whose formatting was changed as
/// a whole - every character of it made bold (or not bold), italic, struck
/// through or underlined, by a selection that covered it; its alignment - and
/// in nothing else: what was set, for the writer to set in the properties of
/// the paragraph's own XML (its runs' w:rPr, its w:pPr) instead of writing
/// the paragraph from the model. None when anything else of it changed, or a
/// part of its text only.
let private paragraphFormatChange (context: WordStyles.DocumentStyleContext option) (index: int) (before: ParagraphBlock) (after: ParagraphBlock) : obj option =
    if TextBuffer.toString before.Buffer <> TextBuffer.toString after.Buffer then None
    else
        let textRuns (buffer: TextBuffer.TextBuffer) =
            buffer.Runs |> List.choose (function TextBuffer.Text run when run.Length > 0 -> Some run | _ -> None)
        let beforeRuns = textRuns before.Buffer
        let afterRuns = textRuns after.Buffer
        // A property every character of the paragraph states, with one value, and did not all state so before.
        let stated (flag: int) (pick: TextBuffer.TextStyle -> bool) : bool option =
            match afterRuns with
            | [] -> None
            | first :: _ ->
                let value = pick first.Style
                let all (runs: TextBuffer.TextRun list) = runs |> List.forall (fun run -> run.Style.Explicit &&& flag <> 0 && pick run.Style = value)
                if all afterRuns && not (all beforeRuns) then Some value else None
        let bold = stated TextBuffer.ExplicitFormat.Bold (fun style -> style.Bold)
        let italic = stated TextBuffer.ExplicitFormat.Italic (fun style -> style.Italic)
        let strike = stated TextBuffer.ExplicitFormat.Strike (fun style -> style.Strike)
        // (An underline of the plain kind: one with a style or a colour of its own is written from the model.)
        let underline =
            stated TextBuffer.ExplicitFormat.Underline (fun style -> style.Underline)
            |> Option.filter (fun _ ->
                afterRuns |> List.forall (fun run -> (run.Style.UnderlineStyle.IsNone || run.Style.UnderlineStyle = Some "single") && run.Style.UnderlineColor.IsNone))
        let alignmentChanged = before.Metadata.Alignment <> after.Metadata.Alignment || before.Metadata.ExplicitAlignment <> after.Metadata.ExplicitAlignment
        if bold.IsNone && italic.IsNone && strike.IsNone && underline.IsNone && not alignmentChanged then None
        else
            // The paragraph as it was, with those properties set as they are now, is the paragraph as it is.
            let set (style: TextBuffer.TextStyle) =
                let style = match bold with Some on -> TextBuffer.withExplicitBold on style | None -> style
                let style = match italic with Some on -> TextBuffer.withExplicitItalic on style | None -> style
                let style = match strike with Some on -> TextBuffer.withExplicitStrike on style | None -> style
                match underline with Some on -> TextBuffer.withExplicitUnderline on style | None -> style
            // (Runs that are formatted alike afterwards are one run, as they are in the paragraph the editor has.)
            let merged (runs: TextBuffer.Run list) =
                List.foldBack (fun (run: TextBuffer.Run) (rest: TextBuffer.Run list) ->
                    match run, rest with
                    | TextBuffer.Text current, TextBuffer.Text next :: others when current.Start + current.Length = next.Start && current.Style = next.Style ->
                        TextBuffer.Text { current with Length = current.Length + next.Length } :: others
                    | _ -> run :: rest) runs []
            let withSet (para: ParagraphBlock) =
                { para with
                    Buffer = { para.Buffer with Runs = para.Buffer.Runs |> List.map (function TextBuffer.Text run -> TextBuffer.Text { run with Style = set run.Style } | other -> other) |> merged }
                    Metadata = { para.Metadata with Alignment = after.Metadata.Alignment; ExplicitAlignment = after.Metadata.ExplicitAlignment } }
            let exported (para: ParagraphBlock) = paragraphToEditor context "Paragraph" 0 para
            let now = exported after
            if jsonOf (exported (withSet before)) <> jsonOf now then None
            else
                let toggle (value: bool option) : obj = match value with Some on -> box on | None -> null
                Some (createObj [
                        "Paragraph" ==> index
                        "Format" ==> true
                        "Bold" ==> toggle bold
                        "Italic" ==> toggle italic
                        "Strike" ==> toggle strike
                        "Underline" ==> toggle underline
                        "AlignmentChanged" ==> alignmentChanged
                        // (What the paragraph states of its alignment now, as the writer reads it of a paragraph: null - its style's.)
                        "Alignment" ==> now?ExplicitAlignment ])

/// A changed block as changes of the text of its paragraphs (see above):
/// the text every paragraph of the block had when it was read, in the order
/// of the XML, and the changes. None when anything else of the block is not
/// what it was.
let private blockTextChanges (context: WordStyles.DocumentStyleContext option) (original: Block) (current: Block) : obj option =
    match storyParagraphs original, storyParagraphs current with
    | Some before, Some after when before.Length = after.Length && List.forall2 (fun (a: ParagraphBlock) (b: ParagraphBlock) -> a.Id = b.Id) before after ->
        let changed = List.zip before after |> List.mapi (fun index pair -> index, pair) |> List.filter (fun (_, (a, b)) -> not (obj.ReferenceEquals(a, b)))
        // With the paragraphs that were changed as they were, the block is what it was.
        let originals = Dictionary<string, ParagraphBlock>()
        for (_, (a, _)) in changed do originals.[a.Id] <- a
        let restored =
            [ current ]
            |> mapParagraphsOfBlocks (fun para -> match originals.TryGetValue para.Id with | true, was -> was | _ -> para)
            |> List.head
        if changed.IsEmpty || jsonOf (blockToEditor context restored) <> jsonOf (blockToEditor context original) then None
        else
            let unchangedText (a: ParagraphBlock) (b: ParagraphBlock) = jsonOf (paragraphToEditor context "Paragraph" 0 a) = jsonOf (paragraphToEditor context "Paragraph" 0 b)
            // (A paragraph whose text is what it was: a change of its formatting as a whole, see paragraphFormatChange.)
            let edits =
                changed
                |> List.map (fun (index, (a, b)) ->
                    if unchangedText a b then Some null
                    else
                        // (First: to the comparison of characters, text made bold is text typed again.)
                        paragraphFormatChange context index a b
                        |> Option.map box
                        |> Option.orElseWith (fun () -> paragraphTextChange context index a b |> Option.map box))
            if edits |> List.exists Option.isNone then None
            else
                Some (createObj [
                        "Texts" ==> (before |> List.map (fun para -> TextBuffer.toString para.Buffer) |> List.toArray)
                        "Edits" ==> (edits |> List.choose id |> List.filter (fun edit -> not (isNull edit)) |> List.toArray) ])
    | _ -> None

/// The body for the editor package. A block that was imported says which
/// item of the imported package it was (SourceIndex) and whether it is still
/// what it was then (SourceUnchanged: the same object, or one that exports to
/// the same item - a character typed and deleted again): the writer puts the
/// XML of such a block back as it was in the file.
let private exportBodyWithSource context (imported: Document option) (doc: Document) : obj array =
    match imported with
    | None -> exportToEditor doc
    | Some original ->
        let known = Dictionary<string, int * Block>()
        original.Blocks |> List.iteri (fun index block -> known.[blockIdOf block] <- (index, block))
        blocksInAnchorOrder doc.Blocks
        |> List.map (fun block ->
            let item = blockToEditor context block
            match known.TryGetValue (blockIdOf block) with
            | true, (index, originalBlock) ->
                let unchanged = sameBlockObject block originalBlock || jsonOf item = jsonOf (blockToEditor context originalBlock)
                item?SourceIndex <- index
                item?SourceUnchanged <- unchanged
                if not unchanged then
                    match blockTextChanges context originalBlock block with
                    | Some changes -> item?SourceTextChanges <- changes
                    | None -> ()
            | _ -> ()
            item)
        |> List.toArray

/// (list id, level) -> how the paragraphs of that list level are numbered.
let private listSignatures (doc: Document) =
    let signatures = Dictionary<string, string>()
    let blocks =
        doc.Blocks
        @ (doc.HeaderParts @ doc.FooterParts |> List.collect (fun part -> part.Blocks))
        @ (doc.Package.Notes |> List.collect (fun note -> note.Blocks))
        @ (doc.Package.Comments |> List.collect (fun comment -> comment.Blocks))
    for para in paragraphsOfBlocks blocks do
        let metadata = para.Metadata
        match metadata.ListId with
        | Some listId when metadata.ListStyle <> NoList ->
            let key = listId + "/" + string metadata.ListLevel
            let value =
                String.concat "\u0000" [
                    listStyleToString metadata.ListStyle
                    metadata.ListNumberFormat |> Option.defaultValue ""
                    metadata.ListLevelText |> Option.defaultValue ""
                    (match metadata.ListStart with Some start -> string start | None -> "") ]
            // (The first paragraph of a level says what the level is.)
            if not (signatures.ContainsKey key) then signatures.[key] <- value
        | _ -> ()
    signatures

let private exportDocumentPackageOf (doc: Document) (imported: Document option) : obj =
    let context = Some (WordStyles.styleContext doc)
    let footnotes = notesToEditor context "footnote" doc
    let endnotes = notesToEditor context "endnote" doc
    let comments = commentsToEditor context doc
    // A header / footer that is what it was when the document was read.
    let partsWithSource (parts: HeaderFooterPart list) (originalParts: HeaderFooterPart list) =
        let items = headerFooterPartsToEditor context parts
        parts
        |> List.iteri (fun index part ->
            match part.RelationshipId |> Option.bind (fun id -> originalParts |> List.tryFind (fun candidate -> candidate.RelationshipId = Some id)) with
            | Some originalPart ->
                let unchanged =
                    obj.ReferenceEquals(part, originalPart)
                    || obj.ReferenceEquals(part.Blocks, originalPart.Blocks) && part.Kind = originalPart.Kind
                    || jsonOf items.[index] = jsonOf (headerFooterPartsToEditor context [ originalPart ]).[0]
                items.[index]?SourceUnchanged <- unchanged
            | None -> ())
        items
    let sameAsImported (current: obj array) (ofOriginal: Document -> obj array) =
        match imported with
        | Some original -> jsonOf current = jsonOf (ofOriginal original)
        | None -> false
    // The levels of lists of the file ("list id/level") that are not numbered as they were.
    let changedListLevels =
        match imported with
        | Some original ->
            let before = listSignatures original
            let after = listSignatures doc
            after
            |> Seq.filter (fun entry ->
                match before.TryGetValue entry.Key with
                | true, value -> value <> entry.Value
                | _ -> false)
            |> Seq.map (fun entry -> entry.Key)
            |> Seq.toArray
        | None -> [||]
    let listsUnchanged = imported.IsSome && changedListLevels.Length = 0
    let partIds (parts: HeaderFooterPart list) = parts |> List.map (fun part -> part.Kind, part.RelationshipId)
    createObj [
        "Footnotes" ==> footnotes
        "Endnotes" ==> endnotes
        "Comments" ==> comments
        // What is still as it was read (the writer then writes it back from the file's own XML).
        "FootnotesUnchanged" ==> sameAsImported footnotes (notesToEditor context "footnote")
        "EndnotesUnchanged" ==> sameAsImported endnotes (notesToEditor context "endnote")
        "CommentsUnchanged" ==> sameAsImported comments (commentsToEditor context)
        "ListsUnchanged" ==> listsUnchanged
        // (The writer then writes those levels anew and leaves the rest of the numbering as it was read.)
        "ChangedListLevels" ==> changedListLevels
        "PageSetupUnchanged" ==> (match imported with
                                  | Some original ->
                                      doc.PageSetup = original.PageSetup
                                      && partIds doc.HeaderParts = partIds original.HeaderParts
                                      && partIds doc.FooterParts = partIds original.FooterParts
                                  | None -> false)
        "Preserved" ==> doc.Package.Preserved
        // Styles changed in the editor (or new): the others keep their original definition.
        "ChangedStyles" ==> (doc.Styles
                             |> List.filter (fun style -> not (doc.Package.ImportedStyles |> List.exists (fun imported -> imported = style)))
                             |> List.map (fun style -> style.StyleType + ":" + style.Id)
                             |> List.toArray)
        "Paragraphs" ==> exportBodyWithSource context imported doc
        "PageSetup" ==> exportPageSetupToEditor doc.PageSetup
        "Styles" ==> exportStylesToEditor doc.Styles
        "HeaderParts" ==> partsWithSource doc.HeaderParts (imported |> Option.map (fun original -> original.HeaderParts) |> Option.defaultValue [])
        "FooterParts" ==> partsWithSource doc.FooterParts (imported |> Option.map (fun original -> original.FooterParts) |> Option.defaultValue [])
    ]

let exportDocumentPackage (doc: Document) : obj =
    let imported : Document option = importedDocumentOf doc.Package.Imported
    // (The shapes that were read, for the ones that are written as they were: see importedTextBoxes.)
    let shapes = Dictionary<string, TextBoxBlock>()
    match imported with
    | Some original ->
        let note (blocks: Block list) =
            for block in blocks do
                match block with
                | TextBox textBox when textBox.SourceXml.IsSome -> shapes.[textBox.Id] <- textBox
                | _ -> ()
        note original.Blocks
        for part in original.HeaderParts do note part.Blocks
        for part in original.FooterParts do note part.Blocks
    | None -> ()
    importedTextBoxes <- (if shapes.Count > 0 then shapes else null)
    try exportDocumentPackageOf doc imported
    finally importedTextBoxes <- null
