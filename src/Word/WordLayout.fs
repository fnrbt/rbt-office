module Rbt.Word.WordLayout

open System
open System.Collections.Generic
open Rbt.Word.Document
open Rbt.Word.TextBuffer

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

let private equalsIgnoreCase (left: string) (right: string) =
    String.Equals(left, right, StringComparison.OrdinalIgnoreCase)

let private optionEqualsIgnoreCase expected value =
    match value with
    | Some actual -> equalsIgnoreCase actual expected
    | None -> false

let private resizeArrayToList (items: ResizeArray<'T>) =
    let mutable result = []
    for i = items.Count - 1 downto 0 do
        result <- items.[i] :: result
    result

let private arrayToList (items: 'T[]) =
    let mutable result = []
    for i = items.Length - 1 downto 0 do
        result <- items.[i] :: result
    result

let private listToIndexedArray (items: 'T list) =
    let result = ResizeArray<int * 'T>()
    let mutable index = 0
    let mutable remaining = items
    while not remaining.IsEmpty do
        match remaining with
        | item :: tail ->
            result.Add((index, item))
            index <- index + 1
            remaining <- tail
        | [] -> ()
    result.ToArray()

let private trimOrDefault fallback (value: string) =
    if isNull value then fallback
    else
        let mutable start = 0
        let mutable finish = value.Length - 1
        while start <= finish && Char.IsWhiteSpace value.[start] do
            start <- start + 1
        while finish >= start && Char.IsWhiteSpace value.[finish] do
            finish <- finish - 1
        if start > finish then fallback
        elif start = 0 && finish = value.Length - 1 then value
        else value.Substring(start, finish - start + 1)

let private mimeEquals (expected: string) (mime: string) =
    if isNull mime then false
    else
        let mutable finish = 0
        let mutable stop = false
        while not stop && finish < mime.Length do
            if mime.[finish] = ';' then stop <- true
            else finish <- finish + 1

        let mutable start = 0
        let mutable last = finish - 1
        while start <= last && Char.IsWhiteSpace mime.[start] do
            start <- start + 1
        while last >= start && Char.IsWhiteSpace mime.[last] do
            last <- last - 1

        let length = last - start + 1
        if length <> expected.Length then false
        else
            let mutable index = 0
            let mutable equal = true
            while equal && index < length do
                let left = mime.[start + index]
                let right = expected.[index]
                if left <> right && Char.ToUpperInvariant left <> Char.ToUpperInvariant right then
                    equal <- false
                index <- index + 1
            equal

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

    let private headerParagraphCountLimit = 100000

    let private firstPartBlocksByKind kind (parts: HeaderFooterPart list) =
        let mutable result = None
        let mutable remaining = parts
        while result.IsNone && not remaining.IsEmpty do
            let part = remaining.Head
            if part.Kind = kind then
                result <- Some part.Blocks
            remaining <- remaining.Tail
        result

    let private firstPartBlocksByRelationshipId relationshipId (parts: HeaderFooterPart list) =
        let mutable result = None
        let mutable remaining = parts
        while result.IsNone && not remaining.IsEmpty do
            let part = remaining.Head
            match part.RelationshipId with
            | Some value when value = relationshipId ->
                result <- Some part.Blocks
            | _ -> ()
            remaining <- remaining.Tail
        result

    let private partBlocks kind (parts: HeaderFooterPart list) =
        firstPartBlocksByKind kind parts

    let private partBlocksForReferences kind (refs: (string * string) list) (parts: HeaderFooterPart list) =
        let mutable result = None
        let mutable remaining = refs
        while result.IsNone && not remaining.IsEmpty do
            match remaining with
            | (refKind, relationshipId) :: tail ->
                if refKind = kind then
                    result <- firstPartBlocksByRelationshipId relationshipId parts
                remaining <- tail
            | [] -> ()
        result

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
        match partBlocks "default" doc.HeaderParts with
        | Some blocks -> blocks
        | None -> doc.Headers |> List.map Paragraph

    let defaultFooterBlocks (doc: Document) =
        match partBlocks "default" doc.FooterParts with
        | Some blocks -> blocks
        | None -> doc.Footers |> List.map Paragraph

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

    let private addHeaderParagraphCount total value =
        if total > headerParagraphCountLimit then total
        else min (headerParagraphCountLimit + 1) (total + max 0 value)

    let private paragraphListCount (paragraphs: ParagraphBlock list) =
        let mutable count = 0
        let mutable remaining = paragraphs
        while count <= headerParagraphCountLimit && not remaining.IsEmpty do
            count <- count + 1
            remaining <- remaining.Tail
        min count (headerParagraphCountLimit + 1)

    let private blockParagraphCount = function
        | Paragraph _
        | Heading _ -> 1
        | Table table ->
            let mutable count = 0
            let mutable rows = table.Rows
            while count <= headerParagraphCountLimit && not rows.IsEmpty do
                let mutable cells = rows.Head.Cells
                while count <= headerParagraphCountLimit && not cells.IsEmpty do
                    count <- addHeaderParagraphCount count (paragraphListCount cells.Head.Content)
                    cells <- cells.Tail
                rows <- rows.Tail
            min count (headerParagraphCountLimit + 1)
        | TextBox textBox ->
            paragraphListCount textBox.Content

    let private blockListParagraphCount (blocks: Block list) =
        let mutable count = 0
        let mutable remaining = blocks
        while count <= headerParagraphCountLimit && not remaining.IsEmpty do
            count <- addHeaderParagraphCount count (blockParagraphCount remaining.Head)
            remaining <- remaining.Tail
        min count (headerParagraphCountLimit + 1)

    let private defaultHeaderParagraphCount (doc: Document) =
        match partBlocks "default" doc.HeaderParts with
        | Some blocks -> blockListParagraphCount blocks
        | None -> paragraphListCount doc.Headers

    let private firstPageHeaderParagraphCount (doc: Document) =
        if doc.PageSetup.DifferentFirstPage then
            match partBlocks "first" doc.HeaderParts with
            | Some first when not first.IsEmpty -> blockListParagraphCount first
            | _ -> defaultHeaderParagraphCount doc
        else
            defaultHeaderParagraphCount doc

    let bodyTopWithRunningHeader (doc: Document) =
        let paragraphCount = firstPageHeaderParagraphCount doc
        if paragraphCount = 0 then
            doc.PageSetup.MarginTop
        else
            let headerLinePitch = 30.0
            let headerDepth = float paragraphCount * headerLinePitch
            max doc.PageSetup.MarginTop (doc.PageSetup.HeaderDistance + headerDepth)

    /// Where the body text of a section starts below its header: at the top
    /// margin, or below the header's paragraphs when they reach further down
    /// (as `bodyTopWithRunningHeader` for a document without sections).
    let bodyTopBelowHeader (marginTop: float) (headerDistance: float) (headerBlocks: Block list) =
        let paragraphCount = blockListParagraphCount headerBlocks
        if paragraphCount = 0 then marginTop
        else max marginTop (headerDistance + float paragraphCount * 30.0)

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
        // Page-relative posV anchors the image to an absolute page-Y; body-relative
        // (margin/paragraph/line) anchors it to an offset down from the current line.
        let y =
            if isPageRelative img.PositionVRelativeFrom then positionY
            else contentY + (lineY - paddingTop) + positionY
        {
            Kind = ImageBox img
            X = x
            Y = y
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
            | Rbt.Word.Document.TextBox textBox -> Some textBox
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
        if mimeEquals "image/png" mime then Some 1
        elif mimeEquals "image/jpeg" mime then Some 2
        elif mimeEquals "image/gif" mime then Some 5
        else None

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
        TextBuffer.isAdditionRevision style

    let isDeletedRevision (style: TextStyle) =
        TextBuffer.isRemovalRevision style

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
            elif style.CommentId.IsSome then "#7E22CE"
            else "#000000"

    let printColor (style: TextStyle) =
        match revisionColor style, style.Color, style.Hyperlink, style.CommentId, style.NoteId with
        | Some color, _, _, _, _ -> Some color
        | None, Some color, _, _, _ -> Some color
        | None, None, Some _, _, _ -> Some "#0563C1"
        // Comment marks stay highlighted in print; footnote/endnote reference
        // marks render in the normal text colour (Word prints them black).
        | None, None, _, Some _, _ -> Some "#7E22CE"
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

let private trimmedEqualsIgnoreCase (value: string) (expected: string) =
    if isNull value then false
    else
        let mutable start = 0
        let mutable finish = value.Length - 1
        while start <= finish && Char.IsWhiteSpace value.[start] do
            start <- start + 1
        while finish >= start && Char.IsWhiteSpace value.[finish] do
            finish <- finish - 1
        let length = finish - start + 1
        if length <> expected.Length then false
        else
            let mutable index = 0
            let mutable equal = true
            while equal && index < length do
                let left = value.[start + index]
                let right = expected.[index]
                if left <> right && Char.ToUpperInvariant left <> Char.ToUpperInvariant right then
                    equal <- false
                index <- index + 1
            equal

module Font =
    let normalizeFamily (fontFamily: string option) =
        match fontFamily with
        | Some family -> trimOrDefault "Calibri" family
        | _ -> "Calibri"

    let exactFace (style: TextStyle) =
        let family = normalizeFamily style.FontFamily
        match style.Bold, style.Italic with
        | true, true -> family + "-BoldItalic"
        | true, false -> family + "-Bold"
        | false, true -> family + "-Italic"
        | false, false -> family

    let knownWordFamily (family: string) =
        trimmedEqualsIgnoreCase family "aptos"
        || trimmedEqualsIgnoreCase family "aptos display"
        || trimmedEqualsIgnoreCase family "aptos body"
        || trimmedEqualsIgnoreCase family "arial"
        || trimmedEqualsIgnoreCase family "arialmt"
        || trimmedEqualsIgnoreCase family "calibri"
        || trimmedEqualsIgnoreCase family "calibri light"
        || trimmedEqualsIgnoreCase family "helvetica"
        || trimmedEqualsIgnoreCase family "inter"
        || trimmedEqualsIgnoreCase family "segoe ui"
        || trimmedEqualsIgnoreCase family "cambria"
        || trimmedEqualsIgnoreCase family "georgia"
        || trimmedEqualsIgnoreCase family "times"
        || trimmedEqualsIgnoreCase family "times new roman"
        || trimmedEqualsIgnoreCase family "consolas"
        || trimmedEqualsIgnoreCase family "courier"
        || trimmedEqualsIgnoreCase family "courier new"
        || trimmedEqualsIgnoreCase family "symbol"
        || trimmedEqualsIgnoreCase family "wingdings"

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
        match style.FontFamily with
        | Some family when trimmedEqualsIgnoreCase family "symbol" || trimmedEqualsIgnoreCase family "wingdings" -> Some 1.27
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
        if optionEqualsIgnoreCase "lowerletter" format then alphaNumber false value
        elif optionEqualsIgnoreCase "upperletter" format then alphaNumber true value
        elif optionEqualsIgnoreCase "lowerroman" format then romanNumeral value |> fun text -> text.ToLowerInvariant()
        elif optionEqualsIgnoreCase "upperroman" format then romanNumeral value
        else string value

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

    let private emptyCellPrintBorders =
        { Top = None
          Right = None
          Bottom = None
          Left = None
          InsideH = None
          InsideV = None }

    type ResolvedCell = {
        Cell: TableCell
        SourceCellIndex: int
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

    type private ResolvedRowWork = {
        Source: TableRow
        Cells: ResolvedCell[]
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

    let preferredWidthFromDocument (width: Rbt.Word.Document.TableWidth) =
        match width.Kind with
        | Rbt.Word.Document.AutoWidth -> AutoWidth
        | Rbt.Word.Document.DxaWidth -> DxaWidth (width.Value |> Option.defaultValue 0.0)
        | Rbt.Word.Document.PercentWidth -> PercentWidth (width.Value |> Option.defaultValue 0.0)
        | Rbt.Word.Document.PointWidth -> PointWidth (width.Value |> Option.defaultValue 0.0)

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
        // Word keeps the cell border box separate from the inter-cell spacing bands.
        // The measured advance is slightly larger than three spacing units because
        // the following row's top band participates in the preceding row advance.
        if spacing > 0.0 then spacing * 3.25 else 0.0

    let spacedCellTextTopOffset spacing paddingTop =
        // Spaced table text starts from the visual cell box, plus Word's internal
        // top inset.  Direct cell tcMar can change row height, but not this origin.
        if spacing > 0.0 then spacing + paddingTop * (2.0 / 3.0) else 0.0

    let spacedCellPaintTopOffset spacing =
        if spacing > 0.0 then spacing * 2.0 else spacing

    let spacedCellPaintHeight rowHeight spacing =
        if spacing > 0.0 then max 0.0 (rowHeight - rowSpacingAdvance spacing) else max 0.0 (rowHeight - spacing)

    let spacedCellPaintTopAdjustment spacing (textPadding: CellPadding) (cellPadding: CellPadding) =
        if spacing > 0.0 then max 0.0 (cellPadding.Top - textPadding.Top) else 0.0

    let spacedCellPaintXOffset spacing columnIndex =
        if spacing > 0.0 && columnIndex > 0 then spacing else 0.0

    let spacedCellPaintWidthAdjustment spacing columnCount start span =
        if spacing <= 0.0 then
            0.0
        else
            let left = if start = 0 then spacing else 0.0
            let right = if start + span >= columnCount then spacing else 0.0
            left + right

    let spacedVerticalMarginExcess (textPadding: CellPadding) (cellPadding: CellPadding) =
        max 0.0 (max (cellPadding.Top - textPadding.Top) (cellPadding.Bottom - textPadding.Bottom))

    let internalSpacingTextOffset spacing columnIndex =
        // Internal cell text starts one spacing unit before the spaced cell border box.
        if spacing > 0.0 && columnIndex > 0 then spacing else 0.0

    let cellSpan (cell: TableCell) = max 1 cell.ColSpan

    let positionedCells (row: TableRow) =
        let mutable currentColumn = max 0 row.GridBefore
        let positioned = ResizeArray<PositionedCell>()
        for cell in row.Cells do
            let span = cellSpan cell
            positioned.Add {
                Cell = cell
                StartColumn = currentColumn
                Span = span
            }
            currentColumn <- currentColumn + span
        resizeArrayToList positioned

    let private indexedPositionedCells (row: TableRow) =
        let mutable currentColumn = max 0 row.GridBefore
        let positioned = ResizeArray<int * PositionedCell>()
        let mutable sourceCellIndex = 0
        for cell in row.Cells do
            let span = cellSpan cell
            positioned.Add(
                sourceCellIndex,
                {
                    Cell = cell
                    StartColumn = currentColumn
                    Span = span
                })
            currentColumn <- currentColumn + span
            sourceCellIndex <- sourceCellIndex + 1
        positioned.ToArray()

    let private cellContentIsEmpty (content: ParagraphBlock list) =
        let mutable empty = true
        let mutable remaining = content
        while empty && not remaining.IsEmpty do
            match remaining with
            | para :: tail ->
                if not (TextBuffer.isWhiteSpace para.Buffer) then
                    empty <- false
                remaining <- tail
            | [] -> ()
        empty

    /// The paragraphs of a vertically merged cell handed out to the cells it
    /// continues in, one each. (How merged cells were imported before the
    /// layout laid a merged cell's content over the rows it spans; the
    /// importers no longer do this - Docx.gatherVerticalMergeContent still
    /// undoes it for documents that hold content in continuation cells.)
    let distributeVerticalMergeContent (rows: TableRow list) =
        let rowArray = rows |> List.toArray
        let cellRows = Array.zeroCreate<TableCell[]> rowArray.Length
        let positionedRows = Array.zeroCreate<(int * PositionedCell)[]> rowArray.Length
        let continueMergeByStart = Array.zeroCreate<System.Collections.Generic.Dictionary<int, int>> rowArray.Length
        for rowIndex = 0 to rowArray.Length - 1 do
            let row = rowArray.[rowIndex]
            cellRows.[rowIndex] <- row.Cells |> List.toArray
            let positioned = indexedPositionedCells row
            positionedRows.[rowIndex] <- positioned
            let continuations = System.Collections.Generic.Dictionary<int, int>(positioned.Length)
            for sourceCellIndex, positionedCell in positioned do
                if positionedCell.Cell.VerticalMerge = Some ContinueMerge then
                    continuations.[positionedCell.StartColumn] <- sourceCellIndex
            continueMergeByStart.[rowIndex] <- continuations

        let setCellContent rowIndex sourceCellIndex content =
            let cell = cellRows.[rowIndex].[sourceCellIndex]
            cellRows.[rowIndex].[sourceCellIndex] <- { cell with Content = content }

        for rowIndex in 0 .. positionedRows.Length - 1 do
            for sourceCellIndex, positioned in positionedRows.[rowIndex] do
                let cell = cellRows.[rowIndex].[sourceCellIndex]
                if cell.VerticalMerge = Some RestartMerge
                   && (match cell.Content with | _ :: _ :: _ -> true | _ -> false) then
                    let continuationCells = ResizeArray<int * int>()
                    let mutable lastContinuation = ValueNone
                    let mutable nextRowIndex = rowIndex + 1
                    let mutable keepGoing = true
                    while keepGoing && nextRowIndex < positionedRows.Length do
                        let mutable foundSourceCellIndex = 0
                        if continueMergeByStart.[nextRowIndex].TryGetValue(positioned.StartColumn, &foundSourceCellIndex) then
                            let continuation = nextRowIndex, foundSourceCellIndex
                            continuationCells.Add continuation
                            lastContinuation <- ValueSome continuation
                            nextRowIndex <- nextRowIndex + 1
                        else
                            keepGoing <- false

                    match cell.Content with
                    | first :: rest when continuationCells.Count > 0 ->
                        setCellContent rowIndex sourceCellIndex [first]
                        let mutable remaining = rest
                        for nextRowIndex, nextSourceCellIndex in continuationCells do
                            match remaining with
                            | para :: tail ->
                                let nextCell = cellRows.[nextRowIndex].[nextSourceCellIndex]
                                let nextContent =
                                    if cellContentIsEmpty nextCell.Content then [para]
                                    else nextCell.Content @ [para]
                                setCellContent nextRowIndex nextSourceCellIndex nextContent
                                remaining <- tail
                            | [] -> ()
                        match remaining, lastContinuation with
                        | [], _ -> ()
                        | extra, ValueSome (lastRowIndex, lastSourceCellIndex) ->
                            let lastCell = cellRows.[lastRowIndex].[lastSourceCellIndex]
                            setCellContent lastRowIndex lastSourceCellIndex (lastCell.Content @ extra)
                        | extra, ValueNone ->
                            let currentCell = cellRows.[rowIndex].[sourceCellIndex]
                            setCellContent rowIndex sourceCellIndex (currentCell.Content @ extra)
                    | _ -> ()

        let resolved = ResizeArray<TableRow>(rowArray.Length)
        for rowIndex = 0 to rowArray.Length - 1 do
            resolved.Add({ rowArray.[rowIndex] with Cells = arrayToList cellRows.[rowIndex] })
        resizeArrayToList resolved

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
        let mutable currentColumn = max 0 row.GridBefore
        for cell in row.Cells do
            currentColumn <- currentColumn + cellSpan cell
        currentColumn + (max 0 row.GridAfter)

    let columnCount (table: TableBlock) =
        let mutable columns = 0
        for row in table.Rows do
            columns <- max columns (rowColumnCount row)
        columns

    let cellPreferredWidth scale availableWidth (cell: TableCell) =
        match cell.PreferredWidth with
        | Some width -> preferredWidthFromDocument width |> resolvePreferredWidth scale availableWidth
        | None -> cell.Width |> Option.bind (preferredWidthFromLegacyValue >> resolvePreferredWidth scale availableWidth)

    let private explicitColumnWidths scale maxWidth columnCount table =
        let widths: float option array = Array.create columnCount None
        for row in table.Rows do
            let mutable currentColumn = max 0 row.GridBefore
            for cell in row.Cells do
                let span = cellSpan cell
                match cellPreferredWidth scale maxWidth cell with
                | Some width ->
                    let perColumnWidth = width / float span
                    let lastColumn = min columnCount (currentColumn + span)
                    let mutable columnIndex = max 0 currentColumn
                    while columnIndex < lastColumn do
                        if widths.[columnIndex].IsNone then
                            widths.[columnIndex] <- Some perColumnWidth
                        columnIndex <- columnIndex + 1
                | None -> ()
                currentColumn <- currentColumn + span
        arrayToList widths

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
            let explicitWidthArray = explicitWidths |> List.toArray
            let mutable explicitTotal = 0.0
            let mutable missingCount = 0
            for width in explicitWidthArray do
                match width with
                | Some value -> explicitTotal <- explicitTotal + value
                | None -> missingCount <- missingCount + 1

            if explicitTotal <= 0.0 then
                match targetWidth with
                | Some width -> List.replicate columns (width / float columns)
                | None -> List.replicate columns scale.DefaultColumnWidth
            else
                let missingWidth =
                    if missingCount = 0 then 0.0
                    else
                        match targetWidth with
                        | Some width -> max 24.0 ((width - explicitTotal) / float missingCount)
                        | None -> scale.DefaultColumnWidth

                let widths = ResizeArray<float>(columns)
                let mutable total = 0.0
                for width in explicitWidthArray do
                    let resolved = defaultArg width missingWidth
                    widths.Add resolved
                    total <- total + resolved
                match targetWidth with
                | Some width ->
                    if total > 0.0 && abs (total - width) > 0.01 then
                        let scaleFactor = width / total
                        for i = 0 to widths.Count - 1 do
                            widths.[i] <- widths.[i] * scaleFactor
                    resizeArrayToList widths
                | None -> resizeArrayToList widths

    let spannedWidth (columnWidths: float list) start span fallback =
        let mutable total = 0.0
        let mutable remaining = span
        let mutable index = start
        let mutable widths = columnWidths

        while remaining > 0 && index < 0 do
            total <- total + fallback
            index <- index + 1
            remaining <- remaining - 1

        while index > 0 do
            match widths with
            | _ :: rest ->
                widths <- rest
                index <- index - 1
            | [] -> index <- 0

        while remaining > 0 do
            match widths with
            | width :: rest ->
                total <- total + width
                widths <- rest
            | [] ->
                total <- total + fallback
            remaining <- remaining - 1

        total

    let startOffset (columnWidths: float list) columnIndex =
        let mutable total = 0.0
        let mutable remaining = columnIndex
        let mutable widths = columnWidths
        while remaining > 0 do
            match widths with
            | width :: rest ->
                total <- total + width
                widths <- rest
                remaining <- remaining - 1
            | [] -> remaining <- 0
        total

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

    let positionedSpacedCellPaintWidth columnCount columnWidths spacing fallback (positioned: PositionedCell) =
        positionedSpacedCellWidth columnWidths spacing fallback positioned
        - spacedCellPaintWidthAdjustment spacing columnCount positioned.StartColumn positioned.Span
        |> max 0.0

    let tableWidthWithSpacing (columnWidths: float list) spacing =
        columnWidths |> List.sum

    let afterSpacing = 0.0

    let tableAfterSpacing scale (table: TableBlock) =
        let spacing = cellSpacing scale table
        if spacing > 0.0 then
            // The table's bottom spacing band participates in the following block
            // advance.  Word includes a small contribution from the table top cell
            // margin here for spaced tables.
            spacing + optionalLength scale 0.0 table.PaddingTop * 0.5
        else
            0.0

    let defaultRowHeight (_lineHeight: float) =
        0.0

    let rowMinHeight scale lineHeight (row: TableRow) =
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

    let private visibleBorderWidth (border: TableBorder option) =
        match border with
        | Some value ->
            match RedBat.BorderGrid.styleKey value.Style with
            | "nil"
            | "none"
            | "hidden" -> 0.0
            | _ -> value.Width |> Option.defaultValue 1.0
        | None -> 0.0

    let cellVerticalBorderExtent (cell: TableCell) =
        match cell.Borders with
        | Some borders ->
            let top = visibleBorderWidth borders.Top
            let bottom = visibleBorderWidth borders.Bottom
            let continuationBottomAllowance =
                match cell.VerticalMerge with
                // Word gives continuation rows a small extra reserve when the
                // spanning cell restates only the bottom border.
                | Some ContinueMerge when top = 0.0 && bottom > 0.0 -> bottom * 0.75
                | _ -> 0.0
            max top bottom + continuationBottomAllowance
        | None -> 0.0

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

    let private printExplicitBorderStrokeFromBorder (border: TableBorder option) fallbackColor fallbackWidth =
        match border with
        | Some value when RedBat.BorderGrid.styleKey value.Style = "nil" || RedBat.BorderGrid.styleKey value.Style = "none" -> None
        | Some value -> printBorderStrokeWithStyle value.Style (value.Color |> Option.orElse fallbackColor) (value.Width |> Option.orElse fallbackWidth)
        | None -> None

    let cellPrintBorderStroke (cell: TableCell) =
        printBorderStroke cell.BorderColor cell.BorderWidth

    let cellPrintBorders (cell: TableCell) : CellPrintBorders =
        match cell.Borders with
        | None ->
            let stroke = cellPrintBorderStroke cell
            { Top = stroke; Right = stroke; Bottom = stroke; Left = stroke; InsideH = None; InsideV = None }
        | Some borders ->
            { Top = printExplicitBorderStrokeFromBorder borders.Top cell.BorderColor cell.BorderWidth
              Right = printExplicitBorderStrokeFromBorder borders.Right cell.BorderColor cell.BorderWidth
              Bottom = printExplicitBorderStrokeFromBorder borders.Bottom cell.BorderColor cell.BorderWidth
              Left = printExplicitBorderStrokeFromBorder borders.Left cell.BorderColor cell.BorderWidth
              InsideH = printExplicitBorderStrokeFromBorder borders.InsideH cell.BorderColor cell.BorderWidth
              InsideV = printExplicitBorderStrokeFromBorder borders.InsideV cell.BorderColor cell.BorderWidth }

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

    let collapsedHorizontalBorderExtent (table: TableBlock) (row: TableRow) =
        let borders = row.BordersOverride |> Option.orElse table.Borders |> tablePrintBorders
        let mutable width = 0.0
        match borders.Top with
        | Some border -> width <- max width border.Width
        | None -> ()
        match borders.Bottom with
        | Some border -> width <- max width border.Width
        | None -> ()
        match borders.InsideH with
        | Some border -> width <- max width border.Width
        | None -> ()
        width

    let private cellPrintBordersWithTableFallback tableBorders rowIndex rowCount columnCount (positioned: PositionedCell) cell =
        let cellBorders = cellPrintBorders cell
        let tableBorders = tablePrintBorders tableBorders
        { Top = cellBorders.Top |> Option.orElse (if rowIndex = 0 then tableBorders.Top else tableBorders.InsideH)
          Right = cellBorders.Right |> Option.orElse (if positioned.StartColumn + positioned.Span >= columnCount then tableBorders.Right else tableBorders.InsideV)
          Bottom = cellBorders.Bottom |> Option.orElse (if rowIndex = rowCount - 1 then tableBorders.Bottom else tableBorders.InsideH)
          Left = cellBorders.Left |> Option.orElse (if positioned.StartColumn = 0 then tableBorders.Left else tableBorders.InsideV)
          InsideH = None
          InsideV = None }

    let private tableCellContentIsEmpty (cell: TableCell) =
        cellContentIsEmpty cell.Content

    let private cellHasPrintBorderInput (cell: TableCell) =
        cell.BorderColor.IsSome
        || cell.BorderWidth.IsSome
        || cell.Borders.IsSome

    let private tableHasPrintBorderInput (table: TableBlock) =
        let mutable found = table.Borders.IsSome
        let mutable remainingRows = table.Rows
        while not found && not remainingRows.IsEmpty do
            match remainingRows with
            | row :: rowTail ->
                if row.BordersOverride.IsSome then
                    found <- true
                else
                    let mutable remainingCells = row.Cells
                    while not found && not remainingCells.IsEmpty do
                        match remainingCells with
                        | cell :: cellTail ->
                            if cellHasPrintBorderInput cell then
                                found <- true
                            remainingCells <- cellTail
                        | [] -> ()
                    remainingRows <- rowTail
            | [] -> ()
        found

    let private tableHasVerticalMerge (table: TableBlock) =
        let mutable found = false
        let mutable remainingRows = table.Rows
        while not found && not remainingRows.IsEmpty do
            match remainingRows with
            | row :: rowTail ->
                let mutable remainingCells = row.Cells
                while not found && not remainingCells.IsEmpty do
                    match remainingCells with
                    | cell :: cellTail ->
                        if cell.VerticalMerge.IsSome then
                            found <- true
                        remainingCells <- cellTail
                    | [] -> ()
                remainingRows <- rowTail
            | [] -> ()
        found

    let private suppressEmptyVerticalMergeRestartTopBorder rowIndex (cell: TableCell) (borders: CellPrintBorders) =
        match cell.VerticalMerge with
        | Some RestartMerge when rowIndex > 0 && tableCellContentIsEmpty cell -> { borders with Top = None }
        | _ -> borders

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

    let private resolveBorderConflictWithWinner first second =
        match first, second with
        | None, winner -> winner, true
        | winner, None -> winner, false
        | Some firstBorder, Some secondBorder ->
            if borderWeight (Some secondBorder) > borderWeight (Some firstBorder) then Some secondBorder, true
            else Some firstBorder, false

    let private collapseResolvedBorders (rows: ResolvedRow list) =
        let rowArray = rows |> List.toArray
        let cellRows = Array.zeroCreate<ResolvedCell[]> rowArray.Length
        for rowIndex = 0 to rowArray.Length - 1 do
            cellRows.[rowIndex] <- rowArray.[rowIndex].Cells |> List.toArray

        let setCellBorders rowIndex cellIndex borders =
            let cells = cellRows.[rowIndex]
            cells.[cellIndex] <- { cells.[cellIndex] with Borders = borders }

        let cellIndexByStart (cells: ResolvedCell[]) =
            let index = Dictionary<int, int>(cells.Length)
            for cellIndex = 0 to cells.Length - 1 do
                index.[cells.[cellIndex].StartColumn] <- cellIndex
            index

        let cellIndexByExactSpan (cells: ResolvedCell[]) =
            let index = Dictionary<struct (int * int), int>(cells.Length)
            for cellIndex = 0 to cells.Length - 1 do
                let cell = cells.[cellIndex]
                index.[struct (cell.StartColumn, cell.ColSpan)] <- cellIndex
            index

        let exactSpanIndexes = Array.zeroCreate<Dictionary<struct (int * int), int>> cellRows.Length
        let exactSpanIndexFor rowIndex =
            let existing = exactSpanIndexes.[rowIndex]
            if obj.ReferenceEquals(existing, null) then
                let created = cellIndexByExactSpan cellRows.[rowIndex]
                exactSpanIndexes.[rowIndex] <- created
                created
            else
                existing

        for rowIndex in 0 .. cellRows.Length - 1 do
            let cells = cellRows.[rowIndex]
            let rightIndexByStart = cellIndexByStart cells

            for leftIndex in 0 .. cells.Length - 1 do
                let left = cells.[leftIndex]
                let mutable rightIndex = 0
                if rightIndexByStart.TryGetValue(left.StartColumn + left.ColSpan, &rightIndex) then
                    let currentLeft = cells.[leftIndex]
                    let currentRight = cells.[rightIndex]
                    let winner, rightWins = resolveBorderConflictWithWinner currentLeft.Borders.Right currentRight.Borders.Left
                    if rightWins then
                        setCellBorders rowIndex leftIndex { currentLeft.Borders with Right = None; InsideH = None; InsideV = None }
                        setCellBorders rowIndex rightIndex { currentRight.Borders with Left = winner; InsideH = None; InsideV = None }
                    else
                        setCellBorders rowIndex leftIndex { currentLeft.Borders with Right = winner; InsideH = None; InsideV = None }
                        setCellBorders rowIndex rightIndex { currentRight.Borders with Left = None; InsideH = None; InsideV = None }

            for upperIndex in 0 .. cells.Length - 1 do
                let upper = cells.[upperIndex]
                let lowerRowIndex = rowIndex + upper.RowSpan
                if lowerRowIndex < cellRows.Length then
                    let lowerIndexBySpan = exactSpanIndexFor lowerRowIndex
                    let mutable lowerIndex = 0
                    if lowerIndexBySpan.TryGetValue(struct (upper.StartColumn, upper.ColSpan), &lowerIndex) then
                        let currentUpper = cells.[upperIndex]
                        let currentLower = cellRows.[lowerRowIndex].[lowerIndex]
                        let winner, lowerWins = resolveBorderConflictWithWinner currentUpper.Borders.Bottom currentLower.Borders.Top
                        if lowerWins then
                            setCellBorders rowIndex upperIndex { currentUpper.Borders with Bottom = None; InsideH = None; InsideV = None }
                            setCellBorders lowerRowIndex lowerIndex { currentLower.Borders with Top = winner; InsideH = None; InsideV = None }
                        else
                            setCellBorders rowIndex upperIndex { currentUpper.Borders with Bottom = winner; InsideH = None; InsideV = None }
                            setCellBorders lowerRowIndex lowerIndex { currentLower.Borders with Top = None; InsideH = None; InsideV = None }

        let resolved = ResizeArray<ResolvedRow>(rowArray.Length)
        for rowIndex = 0 to rowArray.Length - 1 do
            resolved.Add({ rowArray.[rowIndex] with Cells = arrayToList cellRows.[rowIndex] })
        resizeArrayToList resolved

    let private hasResolvedBorders (rows: ResolvedRow list) =
        let mutable found = false
        let mutable remainingRows = rows
        while not found && not remainingRows.IsEmpty do
            match remainingRows with
            | row :: rowTail ->
                let mutable remainingCells = row.Cells
                while not found && not remainingCells.IsEmpty do
                    match remainingCells with
                    | cell :: cellTail ->
                        let borders = cell.Borders
                        if borders.Top.IsSome
                           || borders.Right.IsSome
                           || borders.Bottom.IsSome
                           || borders.Left.IsSome
                           || borders.InsideH.IsSome
                           || borders.InsideV.IsSome then
                            found <- true
                        remainingCells <- cellTail
                    | [] -> ()
                remainingRows <- rowTail
            | [] -> ()
        found

    let private setResolvedCellRowSpan rowIndex cellIndex rowSpan (rows: ResizeArray<ResolvedRowWork>) =
        let cells = rows.[rowIndex].Cells
        cells.[cellIndex] <- { cells.[cellIndex] with RowSpan = rowSpan }

    let private resolveRows scale (table: TableBlock) =
        let rows = table.Rows
        let rowCount = rows.Length
        let hasPrintBorderInput = tableHasPrintBorderInput table
        let hasVerticalMerge = tableHasVerticalMerge table

        let resolveBorderlessRows () =
            let resolvedRows = ResizeArray<ResolvedRow>(rowCount)
            for row in rows do
                let resolvedCells = ResizeArray<ResolvedCell>()
                let mutable sourceCellIndex = 0
                let mutable currentColumn = max 0 row.GridBefore
                for cell in row.Cells do
                    let span = cellSpan cell
                    resolvedCells.Add(
                        { Cell = cell
                          SourceCellIndex = sourceCellIndex
                          StartColumn = currentColumn
                          ColSpan = span
                          RowSpan = 1
                          Borders = emptyCellPrintBorders })
                    currentColumn <- currentColumn + span
                    sourceCellIndex <- sourceCellIndex + 1
                resolvedRows.Add(
                    { Source = row
                      Cells = resizeArrayToList resolvedCells
                      IsHeader = row.IsHeader
                      CanSplit = row.CanSplit |> Option.defaultValue true })
            resizeArrayToList resolvedRows

        let resolveGeneralRows () =
            let columns = if hasPrintBorderInput then columnCount table else 0
            let resolvedRows = ResizeArray<ResolvedRowWork>()
            let activeMerges = Dictionary<int, struct (int * int * int)>()

            rows
            |> List.iteri (fun rowIndex row ->
                let resolvedCells = ResizeArray<ResolvedCell>()
                let effectiveTableBorders = row.BordersOverride |> Option.orElse table.Borders

                let mutable sourceCellIndex = 0
                let mutable currentColumn = max 0 row.GridBefore
                for cell in row.Cells do
                    let span = cellSpan cell
                    let positioned =
                        {
                            Cell = cell
                            StartColumn = currentColumn
                            Span = span
                        }
                    let borders =
                        if hasPrintBorderInput then
                            cellPrintBordersWithTableFallback effectiveTableBorders rowIndex rowCount columns positioned cell
                            |> suppressEmptyVerticalMergeRestartTopBorder rowIndex cell
                        else
                            emptyCellPrintBorders
                    match cell.VerticalMerge with
                    | Some ContinueMerge ->
                        let mutable activeMerge = Unchecked.defaultof<struct (int * int * int)>
                        if activeMerges.TryGetValue(positioned.StartColumn, &activeMerge) then
                            let struct (restartRowIndex, cellIndex, rowSpan) = activeMerge
                            let nextSpan = rowSpan + 1
                            setResolvedCellRowSpan restartRowIndex cellIndex nextSpan resolvedRows
                            activeMerges.[positioned.StartColumn] <- struct (restartRowIndex, cellIndex, nextSpan)
                        else
                            resolvedCells.Add({ Cell = cell; SourceCellIndex = sourceCellIndex; StartColumn = positioned.StartColumn; ColSpan = positioned.Span; RowSpan = 1; Borders = borders })
                    | _ ->
                        let resolvedCell = { Cell = cell; SourceCellIndex = sourceCellIndex; StartColumn = positioned.StartColumn; ColSpan = positioned.Span; RowSpan = 1; Borders = borders }
                        let cellIndex = resolvedCells.Count
                        resolvedCells.Add(resolvedCell)
                        match cell.VerticalMerge with
                        | Some RestartMerge -> activeMerges.[positioned.StartColumn] <- struct (resolvedRows.Count, cellIndex, 1)
                        | _ -> activeMerges.Remove positioned.StartColumn |> ignore
                    currentColumn <- currentColumn + span
                    sourceCellIndex <- sourceCellIndex + 1

                resolvedRows.Add({ Source = row; Cells = resolvedCells.ToArray(); IsHeader = row.IsHeader; CanSplit = row.CanSplit |> Option.defaultValue true }))

            let rows =
                let result = ResizeArray<ResolvedRow>(resolvedRows.Count)
                for row in resolvedRows do
                    result.Add({ Source = row.Source; Cells = arrayToList row.Cells; IsHeader = row.IsHeader; CanSplit = row.CanSplit })
                resizeArrayToList result
            let mutable hasCellSpacing = false
            let mutable remaining = table.Rows
            while not hasCellSpacing && not remaining.IsEmpty do
                match remaining with
                | row :: tail ->
                    if rowCellSpacing scale table row > 0.0 then
                        hasCellSpacing <- true
                    remaining <- tail
                | [] -> ()
            if hasCellSpacing || not hasPrintBorderInput || not (hasResolvedBorders rows) then rows else collapseResolvedBorders rows

        if not hasPrintBorderInput && not hasVerticalMerge then
            resolveBorderlessRows ()
        else
            resolveGeneralRows ()

    let resolveTable scale maxWidth (table: TableBlock) =
        let rows = resolveRows scale table
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
