module RedBat.Word.WordLayout

open System
open RedBat.Word.Document
open RedBat.Word.TextBuffer

type Scale = {
    DxaPerUnit: float
    DefaultCellPadding: float
    DefaultColumnWidth: float
}

let canvasScale = {
    DxaPerUnit = 15.0
    DefaultCellPadding = 0.0
    DefaultColumnWidth = 140.0
}

let printScale = {
    DxaPerUnit = 20.0
    DefaultCellPadding = 0.0
    DefaultColumnWidth = 140.0
}

module Page =
    let gap = 28.0

    type PrintGeometry = {
        Width: float
        Height: float
        MarginTop: float
        MarginRight: float
        MarginBottom: float
        MarginLeft: float
        HeaderDistance: float
        FooterDistance: float
    }

    let printGeometry (setup: PageSetup) = {
        Width = max 1.0 setup.PageWidth
        Height = max 1.0 setup.PageHeight
        MarginTop = max 0.0 setup.MarginTop
        MarginRight = max 0.0 setup.MarginRight
        MarginBottom = max 0.0 setup.MarginBottom
        MarginLeft = max 0.0 setup.MarginLeft
        HeaderDistance = max 0.0 setup.HeaderDistance
        FooterDistance = max 0.0 setup.FooterDistance
    }

    let contentWidth (geometry: PrintGeometry) =
        max 1.0 (geometry.Width - geometry.MarginLeft - geometry.MarginRight)

    let private paragraphsFromBlocks blocks =
        blocks
        |> List.choose (function
            | Paragraph para
            | Heading (_, para) -> Some para
            | Table _
            | TextBox _ -> None)

    let private partBlocks kind (parts: HeaderFooterPart list) =
        parts
        |> List.tryFind (fun part -> part.Kind = kind)
        |> Option.map _.Blocks

    let private partBlocksForReferences kind refs (parts: HeaderFooterPart list) =
        refs
        |> List.tryPick (fun (refKind, relationshipId) ->
            if refKind = kind then
                parts |> List.tryFind (fun part -> part.RelationshipId = Some relationshipId)
            else None)
        |> Option.map _.Blocks

    let private partParagraphs kind parts =
        partBlocks kind parts |> Option.map paragraphsFromBlocks

    let private partParagraphsForReferences kind refs parts =
        partBlocksForReferences kind refs parts |> Option.map paragraphsFromBlocks

    let private defaultPartParagraphs fallback (parts: HeaderFooterPart list) =
        partParagraphs "default" parts
        |> Option.defaultValue fallback

    let defaultHeaderParagraphs (doc: Document) =
        defaultPartParagraphs doc.Headers doc.HeaderParts

    let defaultFooterParagraphs (doc: Document) =
        defaultPartParagraphs doc.Footers doc.FooterParts

    let firstHeaderParagraphs (doc: Document) =
        partParagraphs "first" doc.HeaderParts |> Option.defaultValue []

    let firstFooterParagraphs (doc: Document) =
        partParagraphs "first" doc.FooterParts |> Option.defaultValue []

    let evenHeaderParagraphs (doc: Document) =
        partParagraphs "even" doc.HeaderParts |> Option.defaultValue []

    let evenFooterParagraphs (doc: Document) =
        partParagraphs "even" doc.FooterParts |> Option.defaultValue []

    let referencedHeaderParagraphs kind refs (doc: Document) =
        partParagraphsForReferences kind refs doc.HeaderParts

    let referencedFooterParagraphs kind refs (doc: Document) =
        partParagraphsForReferences kind refs doc.FooterParts

    let defaultHeaderBlocks (doc: Document) =
        partBlocks "default" doc.HeaderParts |> Option.defaultValue (doc.Headers |> List.map Paragraph)

    let defaultFooterBlocks (doc: Document) =
        partBlocks "default" doc.FooterParts |> Option.defaultValue (doc.Footers |> List.map Paragraph)

    let firstHeaderBlocks (doc: Document) =
        if doc.PageSetup.DifferentFirstPage then partBlocks "first" doc.HeaderParts |> Option.defaultValue [] else []

    let firstFooterBlocks (doc: Document) =
        if doc.PageSetup.DifferentFirstPage then partBlocks "first" doc.FooterParts |> Option.defaultValue [] else []

    let evenHeaderBlocks (doc: Document) =
        if doc.PageSetup.EvenAndOddHeaders then partBlocks "even" doc.HeaderParts |> Option.defaultValue [] else []

    let evenFooterBlocks (doc: Document) =
        if doc.PageSetup.EvenAndOddHeaders then partBlocks "even" doc.FooterParts |> Option.defaultValue [] else []

    let referencedHeaderBlocks kind refs (doc: Document) =
        partBlocksForReferences kind refs doc.HeaderParts

    let referencedFooterBlocks kind refs (doc: Document) =
        partBlocksForReferences kind refs doc.FooterParts

    let pageTopForY pageHeight y =
        let stride = pageHeight + gap
        floor (max 0.0 y / stride) * stride

    let contentTop paddingTop pageTop =
        pageTop + paddingTop

    let contentBottom pageHeight marginBottom pageTop =
        pageTop + pageHeight - marginBottom

    let minBodyTop paddingTop fontSize pageTop =
        pageTop + max paddingTop (fontSize * 2.0)

    let private blockParagraphCount = function
        | Paragraph _
        | Heading _ -> 1
        | Table table ->
            table.Rows
            |> List.sumBy (fun row -> row.Cells |> List.sumBy (fun cell -> cell.Content.Length))
        | TextBox textBox -> textBox.Content.Length

    let private firstPageHeaderBlocks (doc: Document) =
        if doc.PageSetup.DifferentFirstPage then
            let first = firstHeaderBlocks doc
            if first.IsEmpty then defaultHeaderBlocks doc else first
        else
            defaultHeaderBlocks doc

    let bodyTopWithRunningHeader (doc: Document) =
        let headers = firstPageHeaderBlocks doc
        if headers.IsEmpty then
            doc.PageSetup.MarginTop
        else
            let headerLinePitch = 30.0
            let headerDepth = headers |> List.sumBy blockParagraphCount |> float |> (*) headerLinePitch
            max doc.PageSetup.MarginTop (doc.PageSetup.HeaderDistance + headerDepth)

    let advanceToNextPage pageHeight paddingTop y =
        let pageTop = pageTopForY pageHeight y
        contentTop paddingTop (pageTop + pageHeight + gap)

    let paginateBlock pageHeight paddingTop marginBottom y blockHeight =
        let pageTop = pageTopForY pageHeight y
        if y > contentTop paddingTop pageTop && y + blockHeight > contentBottom pageHeight marginBottom pageTop then
            advanceToNextPage pageHeight paddingTop y
        else
            y

module Floating =
    type BoxKind =
        | ImageBox of ImageRun
        | TextBox of TextBoxBlock
        | TableBox of TableBlock

    type Box = {
        Kind: BoxKind
        X: float
        Y: float
        Width: float
        Height: float
        Wrap: string option
        DistanceTop: float
        DistanceRight: float
        DistanceBottom: float
        DistanceLeft: float
    }

    let private isPageRelative = function
        | Some "page" -> true
        | _ -> false

    let canvasPosition (margin: float) relativeFrom (position: float) =
        if isPageRelative relativeFrom then position else margin + position

    let lineRelativePosition (margin: float) (current: float) relativeFrom (position: float) =
        if isPageRelative relativeFrom then position - margin else current + position

    let printPosition (margin: float) relativeFrom (position: float) =
        if isPageRelative relativeFrom then position else margin + position

    let printYFromTop (marginTop: float) relativeFrom (position: float) =
        printPosition marginTop relativeFrom position

    let printY (pageHeight: float) (marginTop: float) relativeFrom (position: float) (height: float) =
        pageHeight - printYFromTop marginTop relativeFrom position - height

    let obstacleWidth (availableWidth: float) (x: float) (width: float) (distanceLeft: float) (distanceRight: float) =
        let obstacleX = max 0.0 (x - distanceLeft)
        let obstacleRight = min availableWidth (x + width + distanceRight)
        obstacleX, max 0.0 (obstacleRight - obstacleX)

    let obstacleForBox availableWidth (box: Box) =
        let obstacleX, obstacleWidth = obstacleWidth availableWidth box.X box.Width box.DistanceLeft box.DistanceRight
        let obstacleTop = box.Y - box.DistanceTop
        let obstacleHeight = box.Height + box.DistanceTop + box.DistanceBottom
        obstacleX, obstacleTop, obstacleWidth, obstacleHeight

    let relativeObstacleForBox contentLeft availableWidth (box: Box) =
        obstacleForBox availableWidth { box with X = box.X - contentLeft }

    let wrapsText = function
        | Some "square"
        | Some "tight"
        | Some "through" -> true
        | _ -> false

    let wrapsTopAndBottom = function
        | Some "topAndBottom" -> true
        | _ -> false

    let participatesInWrap (box: Box) =
        match box.Kind with
        | ImageBox img when img.BehindDoc = Some true -> false
        | _ -> wrapsText box.Wrap || wrapsTopAndBottom box.Wrap

    let private overlaps (a: Box) (b: Box) =
        a.X < b.X + b.Width && a.X + a.Width > b.X && a.Y < b.Y + b.Height && a.Y + a.Height > b.Y

    let avoidOverlap (placed: Box list) (box: Box) =
        let mutable adjusted = box
        let mutable changed = true
        while changed do
            changed <- false
            for existing in placed do
                if overlaps adjusted existing then
                    adjusted <- { adjusted with Y = existing.Y + existing.Height }
                    changed <- true
        adjusted

    let imageCanvasBox marginLeft contentY currentX lineY paddingTop width height (img: ImageRun) =
        let positionX = img.PositionX |> Option.defaultValue 0.0
        let positionY = img.PositionY |> Option.defaultValue 0.0
        let x = lineRelativePosition marginLeft currentX img.PositionHRelativeFrom positionX
        let yOffset = lineRelativePosition contentY 0.0 img.PositionVRelativeFrom positionY
        {
            Kind = ImageBox img
            X = x
            Y = contentY + (lineY - paddingTop) + yOffset
            Width = width
            Height = height
            Wrap = img.Wrap
            DistanceTop = img.DistanceTop |> Option.defaultValue 0.0
            DistanceRight = img.DistanceRight |> Option.defaultValue 0.0
            DistanceBottom = img.DistanceBottom |> Option.defaultValue 0.0
            DistanceLeft = img.DistanceLeft |> Option.defaultValue 0.0
        }

    let textBoxCanvasBox marginLeft marginTop (textBox: TextBoxBlock) =
        {
            Kind = TextBox textBox
            X = canvasPosition marginLeft textBox.PositionHRelativeFrom (textBox.X |> Option.defaultValue 0.0)
            Y = canvasPosition marginTop textBox.PositionVRelativeFrom (textBox.Y |> Option.defaultValue 0.0)
            Width = textBox.Width
            Height = textBox.Height
            Wrap = Some "square"
            DistanceTop = 0.0
            DistanceRight = 0.0
            DistanceBottom = 0.0
            DistanceLeft = 0.0
        }

    let topLevelTextBoxes (blocks: Block list) =
        blocks
        |> List.choose (function
            | RedBat.Word.Document.TextBox textBox -> Some textBox
            | _ -> None)

    let topLevelTextBoxCanvasBoxes marginLeft marginTop blocks =
        topLevelTextBoxes blocks |> List.map (textBoxCanvasBox marginLeft marginTop)

    let imagePrintBox marginLeft marginTop pageHeight positionX positionY width height (img: ImageRun) =
        {
            Kind = ImageBox img
            X = printPosition marginLeft img.PositionHRelativeFrom positionX
            Y = printY pageHeight marginTop img.PositionVRelativeFrom positionY height
            Width = width
            Height = height
            Wrap = img.Wrap
            DistanceTop = img.DistanceTop |> Option.defaultValue 0.0
            DistanceRight = img.DistanceRight |> Option.defaultValue 0.0
            DistanceBottom = img.DistanceBottom |> Option.defaultValue 0.0
            DistanceLeft = img.DistanceLeft |> Option.defaultValue 0.0
        }

    let textBoxPrintBox marginLeft marginTop pageHeight positionX positionY width height (textBox: TextBoxBlock) =
        {
            Kind = TextBox textBox
            X = printPosition marginLeft textBox.PositionHRelativeFrom positionX
            Y = printY pageHeight marginTop textBox.PositionVRelativeFrom positionY height
            Width = width
            Height = height
            Wrap = Some "square"
            DistanceTop = 0.0
            DistanceRight = 0.0
            DistanceBottom = 0.0
            DistanceLeft = 0.0
        }

    let topLevelTextBoxPrintBoxes marginLeft marginTop pageHeight pxToPt blocks =
        topLevelTextBoxes blocks
        |> List.map (fun textBox ->
            let positionX = textBox.X |> Option.defaultValue 0.0 |> pxToPt
            let positionY = textBox.Y |> Option.defaultValue 0.0 |> pxToPt
            let width = pxToPt textBox.Width
            let height = pxToPt textBox.Height
            textBoxPrintBox marginLeft marginTop pageHeight positionX positionY width height textBox)

module Image =
    let defaultCanvasWidth = 100.0
    let defaultCanvasHeight = 100.0
    let defaultPrintWidth = 96.0
    let defaultPrintHeight = 72.0

    let pxToPt px = px * 72.0 / 96.0

    let canvasSize width height =
        (width |> Option.defaultValue defaultCanvasWidth,
         height |> Option.defaultValue defaultCanvasHeight)

    let printSize width height =
        (width |> Option.map pxToPt |> Option.defaultValue defaultPrintWidth,
         height |> Option.map pxToPt |> Option.defaultValue defaultPrintHeight)

    let assetId index = sprintf "word-image-%d" index

    let assetKindForMime (mime: string) =
        match mime.Split(';').[0].Trim().ToLowerInvariant() with
        | "image/png" -> Some 1
        | "image/jpeg" -> Some 2
        | "image/gif" -> Some 5
        | _ -> None

    let tryDataUrlBytes (decodeBase64: string -> string) (src: string) =
        let marker = ";base64,"
        let markerIndex = src.IndexOf(marker, StringComparison.OrdinalIgnoreCase)
        if markerIndex < 0 then None
        else
            let mime = src.Substring(5, markerIndex - 5)
            let data = src.Substring(markerIndex + marker.Length)
            let decoded = decodeBase64 data
            let bytes = Array.init decoded.Length (fun i -> byte (int decoded.[i]))
            assetKindForMime mime |> Option.map (fun kind -> kind, bytes)

    let assetIdForSource (assets: (string * string * int * byte[]) list) src =
        assets
        |> List.tryPick (fun (existing, id, _, _) -> if existing = src then Some id else None)
        |> Option.defaultValue ""

module Text =
    let isInsertedRevision (style: TextStyle) =
        style.RevisionType |> Option.exists (fun value -> value = "insert")

    let isDeletedRevision (style: TextStyle) =
        style.RevisionType |> Option.exists (fun value -> value = "delete")

    let revisionColor (style: TextStyle) =
        if isDeletedRevision style then Some "#C00000"
        elif isInsertedRevision style then Some "#008000"
        else None

    let fallbackColor (style: TextStyle) =
        match revisionColor style, style.Color with
        | Some color, _ -> color
        | None, Some color -> color
        | None, None ->
            if style.Hyperlink.IsSome then "#0563C1"
            elif style.CommentId.IsSome || style.NoteId.IsSome then "#7E22CE"
            else "#000000"

    let printColor (style: TextStyle) =
        match revisionColor style, style.Color, style.Hyperlink, style.CommentId, style.NoteId with
        | Some color, _, _, _, _ -> Some color
        | None, Some color, _, _, _ -> Some color
        | None, None, Some _, _, _ -> Some "#0563C1"
        | None, None, _, Some _, _
        | None, None, _, _, Some _ -> Some "#7E22CE"
        | _ -> None

    let printDecoration (style: TextStyle) =
        let underline = style.Underline || Option.isSome style.Hyperlink || isInsertedRevision style
        let strike = style.Strike || isDeletedRevision style
        match underline, strike with
        | false, false -> 0
        | true, false -> 1
        | false, true -> 2
        | true, true -> 3

    let displayText (style: TextStyle) (text: string) =
        if style.AllCaps || style.SmallCaps then text.ToUpperInvariant() else text

    let displayFontSize (style: TextStyle) =
        if style.SmallCaps then style.FontSize * 0.85 else style.FontSize

    let displayStyle (style: TextStyle) =
        if style.SmallCaps then { style with FontSize = style.FontSize * 0.85 } else style

    let printFontSize (style: TextStyle) =
        let baseSize =
            match style.VertAlign with
            | Some "superscript"
            | Some "subscript" -> style.FontSize * 0.7
            | _ -> style.FontSize
        if style.SmallCaps then baseSize * 0.85 else baseSize

    let baselineFontSize (style: TextStyle) =
        match style.VertAlign with
        | Some "superscript"
        | Some "subscript" -> style.FontSize * 0.7
        | _ -> style.FontSize

    let printBaselineShift (style: TextStyle) =
        match style.VertAlign with
        | Some "superscript" -> baselineFontSize style * 0.35
        | Some "subscript" -> -(baselineFontSize style * 0.15)
        | _ -> 0.0

    let underlineOffset (_style: TextStyle) = -2.0
    let underlineThickness (style: TextStyle) = max 0.5 (printFontSize style * 0.05)
    let strikethroughOffset (style: TextStyle) = printFontSize style * 0.35
    let strikethroughThickness (style: TextStyle) = max 0.5 (printFontSize style * 0.05)

    let canvasUnderlineOffset = 2.0
    let canvasStrikethroughOffset (style: TextStyle) = printFontSize style / 3.0

module Font =
    let normalizeFamily (fontFamily: string option) =
        match fontFamily with
        | Some family when family.Trim() <> "" -> family.Trim()
        | _ -> "Calibri"

    let exactFace (style: TextStyle) =
        let family = normalizeFamily style.FontFamily
        match style.Bold, style.Italic with
        | true, true -> family + "-BoldItalic"
        | true, false -> family + "-Bold"
        | false, true -> family + "-Italic"
        | false, false -> family

    let knownWordFamily (family: string) =
        match family.Trim().ToLowerInvariant() with
        | "aptos" | "aptos display" | "aptos body" | "arial" | "arialmt" | "calibri" | "calibri light" | "helvetica" | "inter" | "segoe ui"
        | "cambria" | "georgia" | "times" | "times new roman"
        | "consolas" | "courier" | "courier new" | "symbol" | "wingdings" -> true
        | _ -> false

    let stripSuffix (suffix: string) (value: string) =
        if value.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) then Some (value.Substring(0, value.Length - suffix.Length)) else None

    let splitFace (face: string) =
        match stripSuffix "-BoldItalic" face |> Option.orElseWith (fun () -> stripSuffix "-BoldOblique" face) with
        | Some family -> family, "builtin:inter-bolditalic"
        | None ->
            match stripSuffix "-Italic" face |> Option.orElseWith (fun () -> stripSuffix "-Oblique" face) with
            | Some family -> family, "builtin:inter-italic"
            | None ->
                match stripSuffix "-Bold" face with
                | Some family -> family, "builtin:inter-bold"
                | None -> face, "builtin:inter-regular"

    let serverFontRefForFace face =
        let family, reference = splitFace face
        if knownWordFamily family then Some reference else None

module Paragraph =
    let alignmentTag = function
        | Left -> 0
        | Center -> 1
        | Right -> 2
        | Justify -> 3

    let printBackgroundColor (metadata: BlockMetadata) =
        metadata.BackgroundColor |> Option.bind (fun value -> if value = "" then None else Some value)

    let firstTextStyle (paragraphStyle: WordStyles.ParagraphStyle) (para: ParagraphBlock) =
        para.Buffer.Runs
        |> List.tryPick (function
            | Text tr -> Some (WordStyles.effectiveTextStyle paragraphStyle tr.Style)
            | Image _ -> None)
        |> Option.defaultValue paragraphStyle.Text

    let inlineCount marker (para: ParagraphBlock) =
        para.Buffer.Runs.Length + if Option.isSome marker then 1 else 0

    let spacingBefore (paragraphStyle: WordStyles.ParagraphStyle) (metadata: BlockMetadata) =
        WordStyles.spacingBefore paragraphStyle metadata

    let spacingAfter (paragraphStyle: WordStyles.ParagraphStyle) (metadata: BlockMetadata) =
        WordStyles.spacingAfter paragraphStyle metadata

    let lineSpacing (paragraphStyle: WordStyles.ParagraphStyle) (metadata: BlockMetadata) =
        WordStyles.lineSpacing paragraphStyle metadata

    let leftIndent (metadata: BlockMetadata) =
        metadata.IndentLeft
        |> Option.defaultWith (fun () ->
            match metadata.ListStyle with
            | NoList -> 0.0
            | Bullet
            | Numbered -> 48.0)

    let rightIndent (metadata: BlockMetadata) =
        metadata.IndentRight |> Option.defaultValue 0.0

    let firstLineIndent (metadata: BlockMetadata) =
        metadata.FirstLineIndent
        |> Option.defaultWith (fun () ->
            match metadata.ListStyle with
            | NoList -> 0.0
            | Bullet
            | Numbered -> -48.0)

module Numbering =
    let private markerFontFamily (metadata: BlockMetadata) marker =
        match metadata.ListStyle, metadata.ListLevelText with
        | Bullet, Some text when text.Contains("\uF0B7") -> Some "Symbol"
        | Bullet, Some text when text.Contains("\uF0FC") || text.Contains("\uF0A7") -> Some "Wingdings"
        | Bullet, Some text when text = "o" -> Some "Courier New"
        | _ -> None

    let markerTextStyle (metadata: BlockMetadata) marker (paragraphTextStyle: TextStyle) =
        match markerFontFamily metadata marker with
        | Some family -> { paragraphTextStyle with FontFamily = Some family }
        | None -> paragraphTextStyle

    let displayMarkerText (marker: string) =
        marker
            .Replace("\uF0B7", "•")
            .Replace("\uF0FC", "✓")
            .Replace("\uF0A7", "▪")

    let markerLineBoxFactor (style: TextStyle) =
        match style.FontFamily |> Option.map (fun family -> family.Trim().ToLowerInvariant()) with
        | Some "symbol" | Some "wingdings" -> Some 1.27
        | _ -> None

    let romanNumeral value =
        let table =
            [ 1000, "M"; 900, "CM"; 500, "D"; 400, "CD"; 100, "C"; 90, "XC"; 50, "L"; 40, "XL"; 10, "X"; 9, "IX"; 5, "V"; 4, "IV"; 1, "I" ]
        let mutable remaining = max 1 value
        let mutable parts = []
        for number, marker in table do
            while remaining >= number do
                parts <- marker :: parts
                remaining <- remaining - number
        parts |> List.rev |> String.concat ""

    let alphaNumber upper value =
        let baseChar = if upper then int 'A' else int 'a'
        let mutable n = max 1 value
        let mutable chars = []
        while n > 0 do
            let offset = (n - 1) % 26
            chars <- char (baseChar + offset) :: chars
            n <- (n - 1) / 26
        chars |> Array.ofList |> System.String

    let formatNumber (format: string option) value =
        match format |> Option.map (fun text -> text.ToLowerInvariant()) with
        | Some "lowerletter" -> alphaNumber false value
        | Some "upperletter" -> alphaNumber true value
        | Some "lowerroman" -> romanNumeral value |> fun text -> text.ToLowerInvariant()
        | Some "upperroman" -> romanNumeral value
        | _ -> string value

    let counterKey (metadata: BlockMetadata) =
        (metadata.ListId |> Option.defaultValue "", metadata.ListLevel)

    let formatMarker (metadata: BlockMetadata) (counters: Map<string * int, int>) =
        match metadata.ListLevelText with
        | Some template when template.Contains("%") ->
            let mutable text = template
            for level in 1 .. 9 do
                let key = (metadata.ListId |> Option.defaultValue "", level)
                let value = counters |> Map.tryFind key |> Option.defaultValue 1
                let replacement =
                    if level = metadata.ListLevel then formatNumber metadata.ListNumberFormat value
                    else string value
                text <- text.Replace("%" + string level, replacement)
            text + " "
        | _ -> sprintf "%s. " (formatNumber metadata.ListNumberFormat (counters |> Map.tryFind (counterKey metadata) |> Option.defaultValue 1))

    let markerAndCounters (metadata: BlockMetadata) (counters: Map<string * int, int>) =
        match metadata.ListStyle with
        | NoList -> None, counters
        | Bullet -> Some ((metadata.ListLevelText |> Option.defaultValue "•") + " "), counters
        | Numbered ->
            let level = max 1 metadata.ListLevel
            let key = counterKey metadata
            let next =
                match counters |> Map.tryFind key with
                | Some value -> value + 1
                | None -> metadata.ListStart |> Option.defaultValue 1
            let mutable updated = counters |> Map.add key next
            for childLevel in level + 1 .. 9 do
                updated <- updated |> Map.remove (metadata.ListId |> Option.defaultValue "", childLevel)
            Some (formatMarker metadata updated), updated

    let markerWidth measureText marker textStyle =
        if marker = "" then 0.0
        else max 20.0 (measureText marker textStyle + 8.0)

module Table =
    type CellPadding = {
        Top: float
        Right: float
        Bottom: float
        Left: float
    }

    type PositionedCell = {
        Cell: TableCell
        StartColumn: int
        Span: int
    }

    type PrintStroke = {
        Color: string
        Width: float
        Style: string option
        DashStyle: int
        DashOffset: float
    }

    type CellPrintBorders = {
        Top: PrintStroke option
        Right: PrintStroke option
        Bottom: PrintStroke option
        Left: PrintStroke option
        InsideH: PrintStroke option
        InsideV: PrintStroke option
    }

    type ResolvedCell = {
        Cell: TableCell
        StartColumn: int
        ColSpan: int
        RowSpan: int
        Borders: CellPrintBorders
    }

    type ResolvedRow = {
        Source: TableRow
        Cells: ResolvedCell list
        IsHeader: bool
        CanSplit: bool
    }

    type ResolvedTable = {
        ColumnWidths: float list
        HeaderRows: ResolvedRow list
        BodyRows: ResolvedRow list
    }

    type PreferredWidth =
        | AutoWidth
        | PointWidth of float
        | DxaWidth of float
        | PercentWidth of float

    let normalizeLength scale value =
        if value > 100.0 then value / scale.DxaPerUnit else value

    let preferredWidthFromLegacyValue value =
        if value > 100.0 then DxaWidth value else PointWidth value

    let preferredWidthFromDocument (width: RedBat.Word.Document.TableWidth) =
        match width.Kind with
        | RedBat.Word.Document.AutoWidth -> AutoWidth
        | RedBat.Word.Document.DxaWidth -> DxaWidth (width.Value |> Option.defaultValue 0.0)
        | RedBat.Word.Document.PercentWidth -> PercentWidth (width.Value |> Option.defaultValue 0.0)
        | RedBat.Word.Document.PointWidth -> PointWidth (width.Value |> Option.defaultValue 0.0)

    let resolvePreferredWidth scale availableWidth width =
        match width with
        | AutoWidth -> None
        | PointWidth points -> Some (max 0.0 points)
        | DxaWidth dxa -> Some (max 0.0 (dxa / scale.DxaPerUnit))
        | PercentWidth percent -> availableWidth |> Option.map (fun width -> width * max 0.0 percent / 100.0)

    let optionalLength scale fallback value =
        match value with
        | Some amount when amount > 100.0 -> amount / scale.DxaPerUnit
        | Some amount when amount >= 0.0 -> amount
        | _ -> fallback

    let cellPadding scale (cell: TableCell) : CellPadding = {
        Top = optionalLength scale 0.0 cell.PaddingTop
        Right = optionalLength scale scale.DefaultCellPadding cell.PaddingRight
        Bottom = optionalLength scale 0.0 cell.PaddingBottom
        Left = optionalLength scale scale.DefaultCellPadding cell.PaddingLeft
    }

    let cellPaddingForTable scale (table: TableBlock) (cell: TableCell) : CellPadding =
        let tableTop = optionalLength scale 0.0 table.PaddingTop
        let tableRight = optionalLength scale scale.DefaultCellPadding table.PaddingRight
        let tableBottom = optionalLength scale 0.0 table.PaddingBottom
        let tableLeft = optionalLength scale scale.DefaultCellPadding table.PaddingLeft
        {
            Top = optionalLength scale tableTop cell.PaddingTop
            Right = optionalLength scale tableRight cell.PaddingRight
            Bottom = optionalLength scale tableBottom cell.PaddingBottom
            Left = optionalLength scale tableLeft cell.PaddingLeft
        }

    let rowCellPaddingDefaults scale (table: TableBlock) (row: TableRow) : CellPadding =
        let tableTop = optionalLength scale 0.0 table.PaddingTop
        let tableRight = optionalLength scale scale.DefaultCellPadding table.PaddingRight
        let tableBottom = optionalLength scale 0.0 table.PaddingBottom
        let tableLeft = optionalLength scale scale.DefaultCellPadding table.PaddingLeft
        let rowTop = optionalLength scale tableTop row.PaddingTop
        let rowRight = optionalLength scale tableRight row.PaddingRight
        let rowBottom = optionalLength scale tableBottom row.PaddingBottom
        let rowLeft = optionalLength scale tableLeft row.PaddingLeft
        {
            Top = rowTop
            Right = rowRight
            Bottom = rowBottom
            Left = rowLeft
        }

    let cellPaddingForRow scale (table: TableBlock) (row: TableRow) (cell: TableCell) : CellPadding =
        let rowPadding = rowCellPaddingDefaults scale table row
        {
            Top = optionalLength scale rowPadding.Top cell.PaddingTop
            Right = optionalLength scale rowPadding.Right cell.PaddingRight
            Bottom = optionalLength scale rowPadding.Bottom cell.PaddingBottom
            Left = optionalLength scale rowPadding.Left cell.PaddingLeft
        }

    let cellSpacing scale (table: TableBlock) =
        table.CellSpacing
        |> Option.map (normalizeLength scale)
        |> Option.defaultValue 0.0
        |> max 0.0

    let rowCellSpacing scale (table: TableBlock) (row: TableRow) =
        row.CellSpacing
        |> Option.map (normalizeLength scale)
        |> Option.defaultValue (cellSpacing scale table)
        |> max 0.0

    let rowSpacingAdvance spacing =
        // Word keeps a spacing band above, between, and below spaced rows.
        if spacing > 0.0 then spacing * 3.0 else 0.0

    let internalSpacingTextOffset spacing columnIndex =
        // Internal cell text starts one spacing unit before the spaced cell border box.
        if spacing > 0.0 && columnIndex > 0 then spacing else 0.0

    let cellSpan (cell: TableCell) = max 1 cell.ColSpan

    let positionedCells (row: TableRow) =
        let mutable currentColumn = 0
        row.Cells
        |> List.map (fun cell ->
            let span = cellSpan cell
            let positioned = {
                Cell = cell
                StartColumn = currentColumn
                Span = span
            }
            currentColumn <- currentColumn + span
            positioned)

    let outerCellInsets scale (table: TableBlock) =
        table.Rows
        |> List.tryPick (fun row ->
            match positionedCells row with
            | [] -> None
            | _ ->
                let rowPadding = rowCellPaddingDefaults scale table row
                Some (rowPadding.Left, rowPadding.Right))
        |> Option.defaultValue (0.0, 0.0)

    let rowColumnCount (row: TableRow) =
        positionedCells row
        |> List.tryLast
        |> Option.map (fun positioned -> positioned.StartColumn + positioned.Span)
        |> Option.defaultValue 0

    let columnCount (table: TableBlock) =
        table.Rows
        |> List.map rowColumnCount
        |> List.fold max 0

    let cellPreferredWidth scale availableWidth (cell: TableCell) =
        match cell.PreferredWidth with
        | Some width -> preferredWidthFromDocument width |> resolvePreferredWidth scale availableWidth
        | None -> cell.Width |> Option.bind (preferredWidthFromLegacyValue >> resolvePreferredWidth scale availableWidth)

    let private explicitColumnWidths scale maxWidth columnCount table =
        [0 .. columnCount - 1]
        |> List.map (fun columnIndex ->
            table.Rows
            |> List.tryPick (fun row ->
                positionedCells row
                |> List.tryPick (fun positioned ->
                    if columnIndex >= positioned.StartColumn && columnIndex < positioned.StartColumn + positioned.Span then
                        positioned.Cell |> cellPreferredWidth scale maxWidth |> Option.map (fun width -> width / float positioned.Span)
                    else
                        None)))

    let tableColumnWidths scale maxWidth table =
        let columns = columnCount table

        if columns = 0 then
            []
        else
            let targetWidth =
                match table.PreferredWidth |> Option.map preferredWidthFromDocument with
                | Some AutoWidth -> None
                | Some width -> resolvePreferredWidth scale maxWidth width
                | None -> maxWidth
                |> Option.map (fun width ->
                    match maxWidth with
                    | Some maxWidth -> min maxWidth width
                    | None -> width)
            let gridWidths =
                table.GridColumns
                |> List.choose (preferredWidthFromDocument >> resolvePreferredWidth scale maxWidth)
            let explicitWidths =
                if gridWidths.Length >= columns then
                    gridWidths |> List.truncate columns |> List.map Some
                else
                    explicitColumnWidths scale maxWidth columns table
            let explicitTotal = explicitWidths |> List.choose id |> List.sum

            if explicitTotal <= 0.0 then
                match targetWidth with
                | Some width -> List.replicate columns (width / float columns)
                | None -> List.replicate columns scale.DefaultColumnWidth
            else
                let missingCount = explicitWidths |> List.filter Option.isNone |> List.length
                let missingWidth =
                    if missingCount = 0 then 0.0
                    else
                        match targetWidth with
                        | Some width -> max 24.0 ((width - explicitTotal) / float missingCount)
                        | None -> scale.DefaultColumnWidth

                let widths = explicitWidths |> List.map (Option.defaultValue missingWidth)
                match targetWidth with
                | Some width ->
                    let total = widths |> List.sum
                    if total > 0.0 && abs (total - width) > 0.01 then widths |> List.map (fun columnWidth -> columnWidth * width / total)
                    else widths
                | None -> widths

    let spannedWidth (columnWidths: float list) start span fallback =
        [start .. start + span - 1]
        |> List.sumBy (fun index -> columnWidths |> List.tryItem index |> Option.defaultValue fallback)

    let startOffset (columnWidths: float list) columnIndex =
        columnWidths |> List.take columnIndex |> List.sum

    let spacedStartOffset (columnWidths: float list) spacing columnIndex =
        startOffset columnWidths columnIndex + spacing

    let spacedSpannedWidth (columnWidths: float list) spacing start span fallback =
        max 0.0 (spannedWidth columnWidths start span fallback - spacing * 2.0)

    let positionedCellWidth columnWidths fallback (positioned: PositionedCell) =
        spannedWidth columnWidths positioned.StartColumn positioned.Span fallback

    let positionedCellX columnWidths (positioned: PositionedCell) =
        startOffset columnWidths positioned.StartColumn

    let positionedSpacedCellWidth columnWidths spacing fallback (positioned: PositionedCell) =
        spacedSpannedWidth columnWidths spacing positioned.StartColumn positioned.Span fallback

    let positionedSpacedCellX columnWidths spacing (positioned: PositionedCell) =
        spacedStartOffset columnWidths spacing positioned.StartColumn

    let tableWidthWithSpacing (columnWidths: float list) spacing =
        columnWidths |> List.sum

    let afterSpacing = 0.0

    let tableAfterSpacing scale (table: TableBlock) =
        // The last row leaves the table-level spacing band before following content.
        cellSpacing scale table

    let defaultRowHeight (_lineHeight: float) =
        0.0

    let rowMinHeight scale lineHeight (row: TableRow) =
        match row.HeightRule |> Option.map (fun rule -> rule.ToLowerInvariant()) with
        | Some "atleast"
        | Some "exact" ->
            row.Height
            |> Option.map (normalizeLength scale)
            |> Option.defaultValue (defaultRowHeight lineHeight)
        | _ ->
            row.Height
            |> Option.map (normalizeLength scale)
            |> Option.defaultValue (defaultRowHeight lineHeight)

    let emptyTableAdvance (lineHeight: float) =
        lineHeight + afterSpacing

    let cellContentX columnWidths (positioned: PositionedCell) (padding: CellPadding) =
        positionedCellX columnWidths positioned + padding.Left

    let cellContentWidth (cellWidth: float) (padding: CellPadding) =
        cellWidth - padding.Left - padding.Right

    let cellContentStartY (rowY: float) (padding: CellPadding) =
        rowY + padding.Top

    let cellContentHeight (rowY: float) (paddingTop: float) (paraY: float) =
        paraY - rowY - paddingTop

    let rowHeightForContent (currentRowHeight: float) (padding: CellPadding) (contentHeight: float) =
        max currentRowHeight (contentHeight + padding.Top + padding.Bottom)

    let verticalContentOffset verticalAlignment rowHeight paddingTop paddingBottom contentHeight =
        let available = max 0.0 (rowHeight - paddingTop - paddingBottom - contentHeight)
        match verticalAlignment with
        | Some "center" -> available / 2.0
        | Some "bottom" -> available
        | _ -> 0.0

    let paintBorderColor (cell: TableCell) =
        cell.BorderColor |> Option.defaultValue "#cccccc"

    let paintBorderWidth (cell: TableCell) =
        cell.BorderWidth |> Option.defaultValue 1.0

    let private dashStyleForBorderStyle style =
        match RedBat.BorderGrid.styleKey style with
        | "dotted" -> 1
        | "dashed"
        | "dashsmallgap" -> 2
        | "dotdash"
        | "dotdotdash"
        | "dashdotstroked" -> 3
        | _ -> 0

    let printBorderStrokeWithStyle style color width =
        color |> Option.map (fun borderColor -> {
            Color = borderColor
            Width = width |> Option.defaultValue 1.0
            Style = style
            DashStyle = dashStyleForBorderStyle style
            DashOffset = 0.0
        })

    let printBorderStroke color width =
        printBorderStrokeWithStyle (Some "single") color width

    let printBorderStrokeFromBorder (border: TableBorder option) fallbackColor fallbackWidth =
        match border with
        | Some value when RedBat.BorderGrid.styleKey value.Style = "nil" || RedBat.BorderGrid.styleKey value.Style = "none" -> None
        | Some value -> printBorderStrokeWithStyle value.Style (value.Color |> Option.orElse fallbackColor) (value.Width |> Option.orElse fallbackWidth)
        | None -> printBorderStroke fallbackColor fallbackWidth

    let cellPrintBorderStroke (cell: TableCell) =
        printBorderStroke cell.BorderColor cell.BorderWidth

    let cellPrintBorders (cell: TableCell) : CellPrintBorders =
        match cell.Borders with
        | None ->
            let stroke = cellPrintBorderStroke cell
            { Top = stroke; Right = stroke; Bottom = stroke; Left = stroke; InsideH = None; InsideV = None }
        | Some borders ->
            { Top = printBorderStrokeFromBorder borders.Top cell.BorderColor cell.BorderWidth
              Right = printBorderStrokeFromBorder borders.Right cell.BorderColor cell.BorderWidth
              Bottom = printBorderStrokeFromBorder borders.Bottom cell.BorderColor cell.BorderWidth
              Left = printBorderStrokeFromBorder borders.Left cell.BorderColor cell.BorderWidth
              InsideH = printBorderStrokeFromBorder borders.InsideH cell.BorderColor cell.BorderWidth
              InsideV = printBorderStrokeFromBorder borders.InsideV cell.BorderColor cell.BorderWidth }

    let private tablePrintBorders (borders: TableBorders option) : CellPrintBorders =
        match borders with
        | None -> { Top = None; Right = None; Bottom = None; Left = None; InsideH = None; InsideV = None }
        | Some borders ->
            { Top = printBorderStrokeFromBorder borders.Top None None
              Right = printBorderStrokeFromBorder borders.Right None None
              Bottom = printBorderStrokeFromBorder borders.Bottom None None
              Left = printBorderStrokeFromBorder borders.Left None None
              InsideH = printBorderStrokeFromBorder borders.InsideH None None
              InsideV = printBorderStrokeFromBorder borders.InsideV None None }

    let private cellPrintBordersWithTableFallback tableBorders rowIndex rowCount columnCount (positioned: PositionedCell) cell =
        let cellBorders = cellPrintBorders cell
        let tableBorders = tablePrintBorders tableBorders
        { Top = cellBorders.Top |> Option.orElse (if rowIndex = 0 then tableBorders.Top else tableBorders.InsideH)
          Right = cellBorders.Right |> Option.orElse (if positioned.StartColumn + positioned.Span >= columnCount then tableBorders.Right else tableBorders.InsideV)
          Bottom = cellBorders.Bottom |> Option.orElse (if rowIndex = rowCount - 1 then tableBorders.Bottom else tableBorders.InsideH)
          Left = cellBorders.Left |> Option.orElse (if positioned.StartColumn = 0 then tableBorders.Left else tableBorders.InsideV)
          InsideH = None
          InsideV = None }

    let private borderWeight = function
        | None -> 0.0
        | Some border -> border.Width

    let resolveBorderConflict first second =
        match first, second with
        | None, winner
        | winner, None -> winner
        | Some firstBorder, Some secondBorder ->
            if borderWeight (Some secondBorder) > borderWeight (Some firstBorder) then Some secondBorder
            else Some firstBorder

    let private setCellBorders rowIndex cellIndex borders (rows: ResizeArray<ResolvedRow>) =
        let row = rows.[rowIndex]
        let cells = row.Cells |> List.toArray
        cells.[cellIndex] <- { cells.[cellIndex] with Borders = borders }
        rows.[rowIndex] <- { row with Cells = cells |> Array.toList }

    let private collapseResolvedBorders (rows: ResolvedRow list) =
        let mutable mutableRows = rows |> ResizeArray
        let overlaps firstStart firstSpan secondStart secondSpan =
            firstStart < secondStart + secondSpan && secondStart < firstStart + firstSpan

        for rowIndex in 0 .. mutableRows.Count - 1 do
            let cells = mutableRows.[rowIndex].Cells

            for leftIndex in 0 .. cells.Length - 1 do
                let left = mutableRows.[rowIndex].Cells.[leftIndex]
                for rightIndex in 0 .. cells.Length - 1 do
                    let right = mutableRows.[rowIndex].Cells.[rightIndex]
                    if left.StartColumn + left.ColSpan = right.StartColumn then
                        let currentLeft = mutableRows.[rowIndex].Cells.[leftIndex]
                        let currentRight = mutableRows.[rowIndex].Cells.[rightIndex]
                        let winner = resolveBorderConflict currentLeft.Borders.Right currentRight.Borders.Left
                        setCellBorders rowIndex leftIndex { currentLeft.Borders with Right = winner; InsideH = None; InsideV = None } mutableRows
                        setCellBorders rowIndex rightIndex { currentRight.Borders with Left = None; InsideH = None; InsideV = None } mutableRows

            for upperIndex in 0 .. mutableRows.[rowIndex].Cells.Length - 1 do
                let upper = mutableRows.[rowIndex].Cells.[upperIndex]
                let lowerRowIndex = rowIndex + upper.RowSpan
                if lowerRowIndex < mutableRows.Count then
                    for lowerIndex in 0 .. mutableRows.[lowerRowIndex].Cells.Length - 1 do
                        let lower = mutableRows.[lowerRowIndex].Cells.[lowerIndex]
                        if overlaps upper.StartColumn upper.ColSpan lower.StartColumn lower.ColSpan then
                            let currentUpper = mutableRows.[rowIndex].Cells.[upperIndex]
                            let currentLower = mutableRows.[lowerRowIndex].Cells.[lowerIndex]
                            let winner = resolveBorderConflict currentUpper.Borders.Bottom currentLower.Borders.Top
                            setCellBorders rowIndex upperIndex { currentUpper.Borders with Bottom = winner; InsideH = None; InsideV = None } mutableRows
                            setCellBorders lowerRowIndex lowerIndex { currentLower.Borders with Top = None; InsideH = None; InsideV = None } mutableRows

        mutableRows |> Seq.toList

    let private setResolvedCellRowSpan rowIndex cellIndex rowSpan (rows: ResizeArray<ResolvedRow>) =
        let row = rows.[rowIndex]
        let cells = row.Cells |> List.toArray
        cells.[cellIndex] <- { cells.[cellIndex] with RowSpan = rowSpan }
        rows.[rowIndex] <- { row with Cells = cells |> Array.toList }

    let private resolveRows (table: TableBlock) =
        let rows = table.Rows
        let rowCount = rows.Length
        let columns = columnCount table
        let resolvedRows = ResizeArray<ResolvedRow>()
        let mutable activeMerges = Map.empty<int, int * int * int>

        rows
        |> List.iteri (fun rowIndex row ->
            let resolvedCells = ResizeArray<ResolvedCell>()

            positionedCells row
            |> List.iter (fun positioned ->
                let cell = positioned.Cell
                let borders = cellPrintBordersWithTableFallback table.Borders rowIndex rowCount columns positioned cell
                match cell.VerticalMerge with
                | Some ContinueMerge ->
                    match activeMerges |> Map.tryFind positioned.StartColumn with
                    | Some (rowIndex, cellIndex, rowSpan) ->
                        let nextSpan = rowSpan + 1
                        setResolvedCellRowSpan rowIndex cellIndex nextSpan resolvedRows
                        activeMerges <- activeMerges |> Map.add positioned.StartColumn (rowIndex, cellIndex, nextSpan)
                    | None ->
                        resolvedCells.Add({ Cell = cell; StartColumn = positioned.StartColumn; ColSpan = positioned.Span; RowSpan = 1; Borders = borders })
                | _ ->
                    let resolvedCell = { Cell = cell; StartColumn = positioned.StartColumn; ColSpan = positioned.Span; RowSpan = 1; Borders = borders }
                    let cellIndex = resolvedCells.Count
                    resolvedCells.Add(resolvedCell)
                    match cell.VerticalMerge with
                    | Some RestartMerge -> activeMerges <- activeMerges |> Map.add positioned.StartColumn (resolvedRows.Count, cellIndex, 1)
                    | _ -> activeMerges <- activeMerges |> Map.remove positioned.StartColumn)

            resolvedRows.Add({ Source = row; Cells = resolvedCells |> Seq.toList; IsHeader = row.IsHeader; CanSplit = row.CanSplit |> Option.defaultValue true }))

        resolvedRows |> Seq.toList |> collapseResolvedBorders

    let resolveTable scale maxWidth (table: TableBlock) =
        let rows = resolveRows table
        { ColumnWidths = tableColumnWidths scale maxWidth table
          HeaderRows = rows |> List.takeWhile _.IsHeader
          BodyRows = rows }

    let verticalAlignmentTag = function
        | Some "center" -> 1
        | Some "bottom" -> 2
        | _ -> 0

    let headerRows (table: TableBlock) =
        table.Rows |> List.takeWhile _.IsHeader

    let bodyRows (table: TableBlock) =
        table.Rows
