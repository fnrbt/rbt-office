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

    /// Whether a blank page goes before a section: one that starts on an even
    /// page ("evenPage") or an odd one ("oddPage") when the page that comes
    /// next would have a number - `number` - of the other kind. (The blank
    /// page has that number and is the section's: it counts, and it has the
    /// header and footer the section's rules give a page of its number; the
    /// first page of the section's text is the one after it.)
    let sectionStartsAfterBlankPage (startType: string option) (number: int) : bool =
        match startType with
        | Some "evenPage" -> number % 2 <> 0
        | Some "oddPage" -> number % 2 = 0
        | _ -> false

    /// Whether a section's text is moved on its pages (w:vAlign other than top).
    let sectionIsVerticallyAligned (vAlign: string option) : bool =
        match vAlign with
        | Some "center" | Some "both" | Some "bottom" -> true
        | _ -> false

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

    /// Where a paragraph's first line starts: at the body's top, whatever the size of its text (it was
    /// at least twice the text's size down: a title larger than half the top margin went that much further
    /// down - sample-2's 55 pt titles 71 pt - and ran its page over).
    let minBodyTop (paddingTop: float) (_fontSize: float) (pageTop: float) =
        pageTop + paddingTop

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

    /// Where the body starts below the header, estimated without measuring (30px a
    /// paragraph): what documentSettings gives before a layout; the layout puts it
    /// below the header as it is laid out (Layout.withHeaderFooterExtents).
    let bodyTopWithRunningHeader (doc: Document) =
        let paragraphCount = firstPageHeaderParagraphCount doc
        if paragraphCount = 0 then
            doc.PageSetup.MarginTop
        else
            let headerLinePitch = 30.0
            let headerDepth = float paragraphCount * headerLinePitch
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
        /// A floating table, with where its top is (its box may be the part of it on the page it ends on).
        | TableBox of TableBlock * top: float

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

    /// Where a floating table (w:tblpPr) stands and the room it keeps free of text, canvas px.
    /// HorizontalAnchor: "text" (the column), "margin" or "page"; VerticalAnchor: "text" (the
    /// paragraph after the table, its anchor), "margin" or "page"; XSpec / YSpec: an alignment
    /// that takes the place of X / Y.
    type TablePosition = {
        HorizontalAnchor: string
        VerticalAnchor: string
        X: float
        XSpec: string option
        Y: float
        YSpec: string option
        FromLeft: float
        FromRight: float
        FromTop: float
        FromBottom: float
    }

    let private tablePositions = Dictionary<string, TablePosition>()

    /// The position a floating table's w:tblpPr (TableBlock.FloatingXml) gives, None for a table
    /// in the text flow. Twips to px (15 to the px); an attribute left out: 0, the anchors
    /// "text" across, "margin" down (ECMA-376 17.4.58).
    let tablePosition (table: TableBlock) : TablePosition option =
        if not table.Floating then None
        else
            let xml = table.FloatingXml |> Option.defaultValue ""
            match tablePositions.TryGetValue xml with
            | true, position -> Some position
            | _ ->
                let attributes = Dictionary<string, string>()
                for m in System.Text.RegularExpressions.Regex.Matches(xml, "w:([A-Za-z]+)=\"([^\"]*)\"") do
                    attributes.[m.Groups.[1].Value] <- m.Groups.[2].Value
                let text name = match attributes.TryGetValue name with | true, value when value <> "" -> Some value | _ -> None
                let twips name =
                    match text name |> Option.bind (fun value -> match Double.TryParse(value, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture) with | true, v -> Some v | _ -> None) with
                    | Some value -> value / 15.0
                    | None -> 0.0
                let position =
                    { HorizontalAnchor = text "horzAnchor" |> Option.defaultValue "text"
                      VerticalAnchor = text "vertAnchor" |> Option.defaultValue "margin"
                      X = twips "tblpX"; XSpec = text "tblpXSpec"
                      Y = twips "tblpY"; YSpec = text "tblpYSpec"
                      FromLeft = twips "leftFromText"; FromRight = twips "rightFromText"
                      FromTop = twips "topFromText"; FromBottom = twips "bottomFromText" }
                tablePositions.[xml] <- position
                Some position

    /// A floating table's top (canvas px, down the document): from its anchor paragraph's top
    /// (`anchorY`, vertAnchor text), the page's margin or the page itself - tblpY, or tblpYSpec
    /// in that room (top, center, bottom; with vertAnchor text, its top).
    let tableTop (pageTop: float) (contentTop: float) (contentBottom: float) (pageHeight: float) (anchorY: float) (height: float) (position: TablePosition) =
        let top, room =
            match position.VerticalAnchor with
            | "page" -> pageTop, pageHeight
            | "margin" -> contentTop, contentBottom - contentTop
            | _ -> anchorY, 0.0
        match position.YSpec with
        | Some ("center") when room > 0.0 -> top + (room - height) / 2.0
        | Some ("bottom" | "outside") when room > 0.0 -> top + room - height
        | Some _ -> top
        | None -> top + position.Y

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

    /// Text is not put into a band beside a floating object that is narrower
    /// than this (layout px: a quarter of an inch). Word 2019 puts a word in a
    /// band of 18pt or more, none in one of 16pt (fixtures float-wrap gapil-*:
    /// 5pt words; gap-36: one 22.6pt word in a 36pt band).
    let minimumWrapBand = 24.0

    /// The side of a picture its text may stand on (wp:wrapSquare / wrapTight /
    /// wrapThrough w:wrapText): "left", "right", "largest" (the wider of the
    /// two), or None for both sides.
    let private wrapSide (box: Box) : string option =
        match box.Kind, box.Wrap with
        | ImageBox image, (Some "square" | Some "tight" | Some "through") ->
            match image.WrapText with
            | Some "left" | Some "right" | Some "largest" -> image.WrapText
            | _ -> None
        | _ -> None

    /// What a floating object keeps free of text, across and down (x from
    /// the left edge of the text, y down the document): its box with its
    /// distances from text (distL / distR / distT / distB) - reaching to the
    /// edge of the text on the side its text may not stand on (wrapText).
    let obstacleForBox availableWidth (box: Box) =
        let obstacleX, obstacleWidth = obstacleWidth availableWidth box.X box.Width box.DistanceLeft box.DistanceRight
        let obstacleX, obstacleWidth =
            let right = obstacleX + obstacleWidth
            let toRightEdge = obstacleX, max 0.0 (availableWidth - obstacleX)
            let toLeftEdge = 0.0, max 0.0 right
            match wrapSide box with
            | Some "left" -> toRightEdge
            | Some "right" -> toLeftEdge
            | Some "largest" -> if obstacleX >= availableWidth - right then toRightEdge else toLeftEdge
            | _ -> obstacleX, obstacleWidth
        let obstacleTop = box.Y - box.DistanceTop
        let obstacleHeight = box.Height + box.DistanceTop + box.DistanceBottom
        obstacleX, obstacleTop, obstacleWidth, obstacleHeight

    let relativeObstacleForBox contentLeft availableWidth (box: Box) =
        obstacleForBox availableWidth { box with X = box.X - contentLeft }

    /// The obstacles of a picture with "tight" or "through" wrapping and an
    /// outline (wp:wrapPolygon): the outline in slabs a few pixels high, each
    /// as wide as the outline is there - with the picture's distance from
    /// text to its left and right. Text beside the picture follows the
    /// outline instead of standing clear of the picture's whole box. None
    /// for any other picture (its box is its obstacle, see obstacleForBox).
    let outlineObstacles (availableWidth: float) (box: Box) : (float * float * float * float) list option =
        match box.Kind, box.Wrap with
        | ImageBox img, (Some "tight" | Some "through") when img.WrapPolygon.Length >= 3 && box.Width > 0.0 && box.Height > 0.0 ->
            let points = img.WrapPolygon |> List.map (fun (px, py) -> box.X + px / 21600.0 * box.Width, box.Y + py / 21600.0 * box.Height) |> List.toArray
            let top = points |> Array.map snd |> Array.min
            let bottom = points |> Array.map snd |> Array.max
            let slab = 6.0
            let obstacles = ResizeArray<float * float * float * float>()
            let mutable y = top
            while y < bottom - 0.01 do
                let slabBottom = min bottom (y + slab)
                // The outline's extent between y and slabBottom: its corners there, and where its edges cross the two lines.
                let mutable left = System.Double.PositiveInfinity
                let mutable right = System.Double.NegativeInfinity
                let take (x: float) =
                    if x < left then left <- x
                    if x > right then right <- x
                for i in 0 .. points.Length - 1 do
                    let (x1, y1) = points.[i]
                    let (x2, y2) = points.[(i + 1) % points.Length]
                    if y1 >= y && y1 <= slabBottom then take x1
                    if abs (y2 - y1) > 0.0001 then
                        for edge in [ y; slabBottom ] do
                            let t = (edge - y1) / (y2 - y1)
                            if t >= 0.0 && t <= 1.0 then take (x1 + t * (x2 - x1))
                if right > left then
                    let obstacleX = max 0.0 (left - box.DistanceLeft)
                    let obstacleRight = min availableWidth (right + box.DistanceRight)
                    if obstacleRight > obstacleX then obstacles.Add((obstacleX, y, obstacleRight - obstacleX, slabBottom - y))
                y <- slabBottom
            if obstacles.Count = 0 then None else Some (List.ofSeq obstacles)
        | _ -> None

    /// Whether a point of a picture's box is on the picture as far as text
    /// is concerned: inside its outline when text wraps along one ("tight" /
    /// "through" with a wp:wrapPolygon) - text stands in the rest of the box,
    /// and a click there is a click on that text - and anywhere in the box
    /// otherwise. The box: `x`, `y`, `width`, `height`.
    let pointOnPicture (img: ImageRun) (x: float) (y: float) (width: float) (height: float) (pointX: float) (pointY: float) : bool =
        match img.Wrap with
        | Some "tight" | Some "through" when img.Floating && img.WrapPolygon.Length >= 3 && width > 0.0 && height > 0.0 ->
            let points = img.WrapPolygon |> List.map (fun (px, py) -> x + px / 21600.0 * width, y + py / 21600.0 * height) |> List.toArray
            let mutable inside = false
            for i in 0 .. points.Length - 1 do
                let (x1, y1) = points.[i]
                let (x2, y2) = points.[(i + 1) % points.Length]
                if (y1 > pointY) <> (y2 > pointY) then
                    let crossing = x1 + (pointY - y1) / (y2 - y1) * (x2 - x1)
                    if pointX < crossing then inside <- not inside
            inside
        | _ -> true

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
        | TextBox textBox when textBox.BehindDoc = Some true -> false
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

    /// A shape that may not overlap other objects (wp:anchor allowOverlap="0")
    /// where Word puts it: moved down below each shape placed before it that it
    /// would overlap and that may not overlap either (Word moves it when both
    /// say so; it stays across where it was). Any other box as it is.
    let keptClearOfShapes (placed: Box list) (box: Box) : Box =
        match box.Kind with
        | TextBox textBox when textBox.AllowOverlap = Some false && not textBox.Inline ->
            let others =
                placed
                |> List.filter (fun other ->
                    match other.Kind with
                    | TextBox shape -> shape.AllowOverlap = Some false && not shape.Inline && shape.Id <> textBox.Id
                    | _ -> false)
            if others.IsEmpty then box else avoidOverlap others box
        | _ -> box

    /// The page an anchored object is placed on, as the paragraph that holds
    /// its anchor is laid out: the page's size and margins (layout pixels).
    type AnchorPage = {
        PageWidth: float
        PageHeight: float
        MarginLeft: float
        MarginRight: float
        MarginTop: float
        MarginBottom: float
        /// The column the anchor's paragraph is laid out in: its left edge from
        /// the left margin, and its width (0: the text's whole width - no columns).
        ColumnLeft: float
        ColumnWidth: float
    }

    /// Where an object of `size` is along a stretch that starts at `start`
    /// and is `extent` long: by its alignment (wp:align: left / center /
    /// right, top / center / bottom; inside and outside as left and right on
    /// an odd page), else `offset` from the start (wp:posOffset).
    let alignedPosition (align: string option) (start: float) (extent: float) (size: float) (offset: float) =
        match align with
        | Some "center" -> start + (extent - size) / 2.0
        | Some "right" | Some "bottom" | Some "outside" -> start + extent - size
        | Some "left" | Some "top" | Some "inside" -> start
        | _ -> start + offset

    /// The stretch a horizontal position is relative to (wp:positionH
    /// relativeFrom), from the left edge of the anchor's column: its start
    /// and its length. `currentX`: where the anchor's character is. In a
    /// section with columns `column` is the anchor's column and `margin` the
    /// page's margins (Word 2019: float-wrap col-colright, col-coloff - from the
    /// second column's edges; col-marginoff, col-marginright, col-pageoff - from
    /// the page's).
    let horizontalReference (page: AnchorPage) (currentX: float) (relativeFrom: string option) : float * float =
        let textWidth = page.PageWidth - page.MarginLeft - page.MarginRight
        let columnWidth = if page.ColumnWidth > 0.0 then page.ColumnWidth else textWidth
        let margin = -page.ColumnLeft
        match relativeFrom with
        | Some "page" -> margin - page.MarginLeft, page.PageWidth
        | Some "margin" -> margin, textWidth
        | Some "column" -> 0.0, columnWidth
        | Some "leftMargin" | Some "insideMargin" -> margin - page.MarginLeft, page.MarginLeft
        | Some "rightMargin" | Some "outsideMargin" -> margin + textWidth, page.MarginRight
        // (character, and an object that says nothing: from where its anchor is in the line.)
        | _ -> currentX, 0.0

    /// The stretch a vertical position is relative to (wp:positionV
    /// relativeFrom), in layout pixels down the canvas: its start and its
    /// length. `pageTop`: the top of the page the anchor's line is on;
    /// `paragraphTop` / `lineTop`: of the anchor's paragraph and line.
    let verticalReference (page: AnchorPage) (pageTop: float) (paragraphTop: float) (lineTop: float) (relativeFrom: string option) : float * float =
        match relativeFrom with
        | Some "page" -> pageTop, page.PageHeight
        | Some "margin" -> pageTop + page.MarginTop, page.PageHeight - page.MarginTop - page.MarginBottom
        | Some "topMargin" | Some "insideMargin" -> pageTop, page.MarginTop
        | Some "bottomMargin" | Some "outsideMargin" -> pageTop + page.PageHeight - page.MarginBottom, page.MarginBottom
        | Some "paragraph" -> paragraphTop, 0.0
        // (line, and an object that says nothing: from its anchor's line.)
        | _ -> lineTop, 0.0

    /// An anchored picture where its anchor puts it: relative to the page,
    /// the margins, the column, the anchor's paragraph, line or character,
    /// by an alignment or an offset (wp:positionH / wp:positionV). X is from
    /// the left edge of the text, Y down the canvas. (It was placed from the
    /// anchor's character and line unless page-relative - and page-relative
    /// meant the first page; the alignments were not read.)
    let imageCanvasBox (page: AnchorPage) contentY currentX lineY paddingTop width height (img: ImageRun) =
        let positionX = img.PositionX |> Option.defaultValue 0.0
        let positionY = img.PositionY |> Option.defaultValue 0.0
        let lineTop = contentY + (lineY - paddingTop)
        // (On a page without an end - a strip that is placed afterwards - the page is where the paragraph is.)
        let paged = page.PageHeight < 50000.0
        let pageTop = if paged then Page.pageTopForY page.PageHeight lineTop else 0.0
        let x =
            let start, extent = horizontalReference page currentX img.PositionHRelativeFrom
            alignedPosition img.PositionHAlign start extent width positionX
        let y =
            let relativeFrom =
                match img.PositionVRelativeFrom with
                | Some "page" | Some "margin" | Some "topMargin" | Some "bottomMargin" | Some "insideMargin" | Some "outsideMargin" when not paged ->
                    // (No page to go by: an offset from the page's top is one from the strip's top, as before.)
                    None
                | given -> given
            match img.PositionVRelativeFrom, relativeFrom with
            | Some "page", None -> positionY
            | _ ->
                let start, extent = verticalReference page pageTop contentY lineTop relativeFrom
                alignedPosition img.PositionVAlign start extent height positionY
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

    /// The room between a shape's edge and its text: left, top, right, bottom
    /// (wps:bodyPr lIns / tIns / rIns / bIns; 0.1 inch and 0.05 inch when the
    /// shape does not say).
    let textBoxInsets (textBox: TextBoxBlock) : float * float * float * float =
        textBox.InsetLeft |> Option.defaultValue 9.6,
        textBox.InsetTop |> Option.defaultValue 4.8,
        textBox.InsetRight |> Option.defaultValue 9.6,
        textBox.InsetBottom |> Option.defaultValue 4.8

    /// How a shape's anchor wraps the text (none when it does not say; a
    /// shape in the line stands on lines of its own).
    let textBoxWrap (textBox: TextBoxBlock) : string option =
        if textBox.Inline then Some "topAndBottom" else textBox.Wrap |> Option.orElse (Some "none")

    /// A text box or shape where its anchor puts it, as `imageCanvasBox` puts
    /// an anchored picture: relative to the page, the margins, the column or
    /// the anchor's paragraph, by an alignment or an offset. X is from the
    /// left edge of the text, Y down the canvas. `pageTop`: the top of the
    /// page the anchor's paragraph starts on; `paragraphTop`: of that
    /// paragraph; `height`: of the shape (of its text when it follows it).
    /// (Every shape was placed on the first page, by its offset alone.)
    let anchoredTextBoxBox (page: AnchorPage) (pageTop: float) (paragraphTop: float) (height: float) (textBox: TextBoxBlock) : Box =
        let x, y =
            if textBox.Inline then 0.0, paragraphTop
            else
                let startX, extentX = horizontalReference page 0.0 (textBox.PositionHRelativeFrom |> Option.orElse (Some "column"))
                let startY, extentY = verticalReference page pageTop paragraphTop paragraphTop (textBox.PositionVRelativeFrom |> Option.orElse (Some "paragraph"))
                alignedPosition textBox.PositionHAlign startX extentX textBox.Width (textBox.X |> Option.defaultValue 0.0),
                alignedPosition textBox.PositionVAlign startY extentY height (textBox.Y |> Option.defaultValue 0.0)
        {
            Kind = TextBox textBox
            X = x
            Y = y
            Width = textBox.Width
            Height = height
            Wrap = textBoxWrap textBox
            DistanceTop = textBox.DistanceTop |> Option.defaultValue 0.0
            DistanceRight = textBox.DistanceRight |> Option.defaultValue 0.0
            DistanceBottom = textBox.DistanceBottom |> Option.defaultValue 0.0
            DistanceLeft = textBox.DistanceLeft |> Option.defaultValue 0.0
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
    /// The level of the paragraph's list the file defines (none: a list made in the editor).
    let levelDefinition (metadata: BlockMetadata) (level: int) : ListLevelDefinition option =
        if metadata.ListLevels.IsEmpty then None
        else metadata.ListLevels |> List.tryFind (fun definition -> definition.Level = level)

    let private ownLevel (metadata: BlockMetadata) : ListLevelDefinition option =
        // (The level the paragraph was read at, or was moved to: its format is the definition's.)
        levelDefinition metadata metadata.ListLevel
        |> Option.filter (fun definition -> metadata.ListNumberFormat.IsNone || metadata.ListNumberFormat = Some definition.NumberFormat)

    let private markerFontFamily (metadata: BlockMetadata) marker =
        match metadata.ListStyle, metadata.ListLevelText with
        | Bullet, Some text when text.Contains("\uF0B7") -> Some "Symbol"
        | Bullet, Some text when text.Contains("\uF0FC") || text.Contains("\uF0A7") -> Some "Wingdings"
        | Bullet, Some text when text = "o" -> Some "Courier New"
        | _ -> None

    /// The style the number or bullet is set in: the paragraph's, with what
    /// the list level says of its number (face, bold, italic, colour, size).
    /// (The paragraph mark's own size is the marker's under the level's: fixtures marker-ascent `*-sym`, a
    /// Symbol bullet at the 10pt or 12pt of the mark's w:rPr where the style is 11pt.)
    let markerTextStyle (metadata: BlockMetadata) marker (paragraphTextStyle: TextStyle) =
        let paragraphTextStyle =
            match metadata.ExplicitMarkFontSize with
            | Some size when size > 0.0 -> { paragraphTextStyle with FontSize = size }
            | _ -> paragraphTextStyle
        let styled =
            match ownLevel metadata with
            | None -> paragraphTextStyle
            | Some definition ->
                let style = paragraphTextStyle
                let style = match definition.MarkBold with Some bold -> { style with Bold = bold } | None -> style
                let style = match definition.MarkItalic with Some italic -> { style with Italic = italic } | None -> style
                let style = match definition.MarkColor with Some color -> { style with Color = Some color } | None -> style
                let style = match definition.MarkSize with Some size when size > 0.0 -> { style with FontSize = size } | _ -> style
                match definition.MarkFont with
                | Some family when family.Trim() <> "" -> { style with FontFamily = Some family }
                | _ -> style
        match markerFontFamily metadata marker with
        | Some family -> { styled with FontFamily = Some family }
        | None -> styled

    /// The characters of Wingdings and of Symbol a bullet is commonly made of
    /// (the private-use code U+F0xx Word writes for them), as the Unicode
    /// character they look like: a host has no Wingdings or Symbol face, and
    /// no other face has those codes (they showed as missing-glyph boxes).
    let private wingdingsCharacters =
        dict [
            0x6C, "●"; 0x6D, "❍"; 0x6E, "■"; 0x6F, "□"; 0x70, "◻"; 0x71, "❑"; 0x72, "❒"; 0x73, "⬧"
            0x74, "⧫"; 0x75, "◆"; 0x76, "❖"; 0x77, "⬥"; 0x78, "⌧"; 0x7A, "⌘"; 0x7B, "❀"; 0x7C, "✿"
            0x9E, "·"; 0x9F, "•"; 0xA0, "▪"; 0xA1, "○"; 0xA4, "◉"; 0xA5, "◎"; 0xA7, "▪"; 0xA8, "◻"
            0xAA, "✦"; 0xAB, "★"; 0xAC, "✶"; 0xAD, "✴"; 0xAE, "✹"; 0xAF, "✵"; 0xD8, "➢"; 0xE8, "➔"
            0xF0, "⇨"; 0xFB, "✗"; 0xFC, "✓"; 0xFD, "☒"; 0xFE, "☑" ]
    let private symbolCharacters =
        dict [ 0xB7, "•"; 0xA7, "♣"; 0xA8, "♦"; 0xA9, "♥"; 0xAA, "♠"; 0xB0, "°"; 0x2D, "−"; 0xAE, "→"; 0xDE, "⇒"; 0xE0, "◊" ]

    /// A list marker as it is shown: the private-use codes of Wingdings / Symbol
    /// (`family`: the marker's face) as the Unicode characters they look like.
    let displayMarkerText (family: string option) (marker: string) =
        let table =
            match family with
            | Some family when trimmedEqualsIgnoreCase family "symbol" -> Some symbolCharacters
            | Some family when trimmedEqualsIgnoreCase family "wingdings" -> Some wingdingsCharacters
            | _ -> None
        let mapped =
            match table with
            | Some table when marker |> Seq.exists (fun c -> int c >= 0xF020 && int c <= 0xF0FF) ->
                marker
                |> Seq.map (fun c ->
                    let code = int c
                    match table.TryGetValue (code - 0xF000) with
                    | true, shown when code >= 0xF020 && code <= 0xF0FF -> shown
                    | _ -> string c)
                |> String.concat ""
            | _ -> marker
        mapped
            .Replace("\uF0B7", "•")
            .Replace("\uF0FC", "✓")
            .Replace("\uF0A7", "▪")

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
        elif optionEqualsIgnoreCase "decimalzero" format then (if value < 10 then "0" + string value else string value)
        elif optionEqualsIgnoreCase "none" format then ""
        else string value

    let counterKey (metadata: BlockMetadata) =
        (metadata.ListId |> Option.defaultValue "", metadata.ListLevel)

    /// What follows the number: a tab (shown as a space) unless the level says a space or nothing.
    let private markerSuffix (metadata: BlockMetadata) =
        match ownLevel metadata |> Option.bind (fun definition -> definition.Suffix) with
        | Some "nothing" -> ""
        | _ -> " "

    let formatMarker (metadata: BlockMetadata) (counters: Map<string * int, int>) =
        match metadata.ListLevelText with
        | Some template when template.Contains("%") ->
            // (w:isLgl: the numbers of all the levels are decimal in this level's text.)
            let legal = ownLevel metadata |> Option.exists (fun definition -> definition.Legal)
            let mutable text = template
            for level in 1 .. 9 do
                if text.Contains("%" + string level) then
                    let key = (metadata.ListId |> Option.defaultValue "", level)
                    let definition = levelDefinition metadata level
                    let value =
                        match counters |> Map.tryFind key with
                        | Some value -> value
                        // (A level above that has no paragraph yet counts from its start.)
                        | None -> definition |> Option.bind (fun definition -> definition.Start) |> Option.defaultValue 1
                    let replacement =
                        if legal then string value
                        elif level = metadata.ListLevel then formatNumber metadata.ListNumberFormat value
                        else
                            // (The number of another level is written as that level writes it.)
                            match definition with
                            | Some definition -> formatNumber (Some definition.NumberFormat) value
                            | None -> string value
                    text <- text.Replace("%" + string level, replacement)
            text + markerSuffix metadata
        | Some template when metadata.ListStyle = Numbered && not metadata.ListLevels.IsEmpty ->
            // (A level whose text has no number: the text as it is - "Chapter", or nothing.)
            if template = "" then "" else template + markerSuffix metadata
        | _ -> sprintf "%s. " (formatNumber metadata.ListNumberFormat (counters |> Map.tryFind (counterKey metadata) |> Option.defaultValue 1))

    let markerAndCounters (metadata: BlockMetadata) (counters: Map<string * int, int>) =
        match metadata.ListStyle with
        | NoList -> None, counters
        | Bullet -> Some ((metadata.ListLevelText |> Option.defaultValue "•") + " "), counters
        | Numbered ->
            let updated = Rbt.Word.Document.advanceListCounters metadata counters
            Some (formatMarker metadata updated), updated

    /// How the number stands at its position ("left", "center", "right").
    let markerJustification (metadata: BlockMetadata) : string =
        match metadata.ListStyle with
        | NoList -> "left"
        | _ -> ownLevel metadata |> Option.bind (fun definition -> definition.Justification) |> Option.defaultValue "left"

    /// What follows the number (w:suff): "tab" (the default), "space" or "nothing".
    let markerSuffixKind (metadata: BlockMetadata) =
        match ownLevel metadata |> Option.bind (fun definition -> definition.Suffix) with
        | Some "space" -> "space"
        | Some "nothing" -> "nothing"
        | _ -> "tab"

    /// The number or bullet without what follows it (formatMarker shows a tab as a space).
    let markerBody (metadata: BlockMetadata) (marker: string) =
        if markerSuffixKind metadata <> "nothing" && marker.EndsWith " " then marker.Substring(0, marker.Length - 1)
        else marker

    /// Where the text of a numbered paragraph's first line starts, from the
    /// left indent, as Word places it. The number stands at the first-line
    /// indent (w:lvlJc: its start, middle or end there). A tab after it goes
    /// to the nearest of: a custom tab stop past the number (the paragraph's,
    /// or the level's own), the left indent when the paragraph hangs and the
    /// number ends before it; failing those, the next default tab stop. A
    /// space after it is a space; nothing, nothing. A number wider than its
    /// hanging indent so pushes the text on instead of running into it.
    let textStart (measureText: string -> TextStyle -> float) (defaultTabStop: float) (metadata: BlockMetadata) (indent: float) (firstLineIndent: float) (marker: string) (markerStyle: TextStyle) =
        let body = displayMarkerText markerStyle.FontFamily (markerBody metadata marker)
        let bodyWidth = if body = "" then 0.0 else measureText body markerStyle
        let markerEnd =
            match markerJustification metadata with
            | "right" | "end" -> firstLineIndent
            | "center" -> firstLineIndent + bodyWidth / 2.0
            | _ -> firstLineIndent + bodyWidth
        match markerSuffixKind metadata with
        | "nothing" -> markerEnd
        | "space" -> markerEnd + measureText " " markerStyle
        | _ ->
            // Tab stops are measured from the margin, not from the indent.
            let endAtMargin = indent + markerEnd
            let mutable best = System.Double.MaxValue
            let consider position = if position > endAtMargin + 0.01 && position < best then best <- position
            for stop in metadata.TabStops do
                if stop.Alignment <> "bar" then consider stop.Position
            metadata.RightTabPosition |> Option.iter consider
            if firstLineIndent < 0.0 then consider indent
            let target =
                if best < System.Double.MaxValue then best
                elif defaultTabStop > 0.0 then (floor (endAtMargin / defaultTabStop) + 1.0) * defaultTabStop
                else endAtMargin
            target - indent

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

    /// An indent applies to left-aligned tables only, in canvas pixels.
    let tableIndent (table: TableBlock) =
        match table.Alignment with
        | Some alignment when alignment.ToLowerInvariant() = "center" || alignment.ToLowerInvariant() = "right" || alignment.ToLowerInvariant() = "end" -> 0.0
        | _ -> table.Indent |> Option.defaultValue 0.0

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
    /// continues in, one each. (With the old import the layout placed merged
    /// cell content over every row it spanned; the exporter undoes this for
    /// documents that still hold content in continuation cells.)
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

    /// How much wider than the text a table's room is (a percentage's base, an autofit table's
    /// most): in compatibility mode 14 and older its outer cells' margins, a centred or
    /// right-aligned table's too (fixtures nested-width calibre-c14: a centred 70% table 70% of
    /// the text and margins); none from mode 15 on.
    let outerCellRoom scale (compatibilityMode: int) (table: TableBlock) =
        if compatibilityMode >= 15 then (0.0, 0.0)
        else
            table.Rows
            |> List.tryPick (fun row ->
                match positionedCells row with
                | [] -> None
                | first :: _ ->
                    let rowPadding = rowCellPaddingDefaults scale table row
                    let firstPadding = cellPaddingForRow scale table row first.Cell
                    Some (firstPadding.Left, rowPadding.Right))
            |> Option.defaultValue (0.0, 0.0)

    /// How far a table reaches past its indent on the left and right. In a document of
    /// compatibilityMode 14 or older Word puts the text of the first cell where the
    /// paragraphs' text is: the table starts that cell's left margin further left (and a
    /// table as wide as the text reaches as far past it on the right, by the row's margin).
    /// From mode 15 on the table's edge is at the indent. A centred or right-aligned
    /// table is placed by its edges in every mode.
    let outerCellInsets scale (compatibilityMode: int) (table: TableBlock) =
        let aligned =
            match table.Alignment with
            | Some alignment -> (match alignment.ToLowerInvariant() with "center" | "right" | "end" -> true | _ -> false)
            | None -> false
        if aligned then (0.0, 0.0) else outerCellRoom scale compatibilityMode table

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

    /// The widths of a table's columns, as Word works them out (fixtures table-width,
    /// tools/word-parity/fixtures/table-width.py; docs/layout-parity-flow.md). `maxWidth` is the
    /// room the table has (the text's or the column's or the cell's width, less the indent; in
    /// compatibility mode 14 and older with the outer cells' margins it reaches past the text by);
    /// `content` gives a cell's content in an autofit table: its longest word and its width set on
    /// one line (cell margins not included).
    /// - Each column is as wide as the widest of its cells' widths (w:tcW), a cell that spans
    ///   columns giving the last of them what they lack of it; the continuation of a vertical
    ///   merge counts for nothing. The grid (w:tblGrid) stands in only where a cell has no width,
    ///   in a fixed table; an autofit table takes such a cell's content on one line instead.
    /// - The table is as wide as its w:tblW (a percentage of `maxWidth` and the indent); with none (auto) a fixed
    ///   table is as wide as its columns, even past the margin, and an autofit one too as far as
    ///   `maxWidth`.
    /// - Between the two widths the columns give way past their least widths in proportion to what
    ///   they have above them, and grow in proportion to their widths. The least width of an autofit
    ///   column is its cells' longest word and margins (a column is never narrower), of a fixed
    ///   one its margins and one pixel.
    /// A table with no widths at all (one the editor has just made) fills `maxWidth` evenly.
    /// `fixedLayout`: whether the table is laid out fixed - a nested table is laid out as the
    /// table it is in (fixtures nested-width nest-fixed, nest-*-outfixed). Also gives the table's
    /// least width: its columns' least widths, or its w:tblW where that is more and not a percentage.
    let tableColumnWidthsAndLeast scale (fixedLayout: bool) (maxWidth: float option) (content: (TableCell -> struct (float * float)) option) (table: TableBlock) : float list * float =
        let columns = columnCount table
        if columns = 0 then [], 0.0 else
        // (A percentage is of the whole width the table could have, its indent included:
        // docx4j-getting-started's 78.76% table, indented 745 twips, is its grid's 7291.)
        let tableWidth =
            match table.PreferredWidth |> Option.map preferredWidthFromDocument with
            | None | Some AutoWidth -> None
            | Some width -> resolvePreferredWidth scale (maxWidth |> Option.map (fun room -> room + tableIndent table)) width
        let grid = table.GridColumns |> List.map (preferredWidthFromDocument >> resolvePreferredWidth scale (tableWidth |> Option.orElse maxWidth)) |> List.toArray
        let gridAt column = if column < grid.Length then grid.[column] else None
        let measured = Dictionary<TableCell, struct (float * float)>(HashIdentity.Reference)
        let measure (cell: TableCell) =
            match content with
            | Some measure when not fixedLayout ->
                match measured.TryGetValue cell with
                | true, value -> value
                | _ ->
                    let value = measure cell
                    measured.[cell] <- value
                    value
            | _ -> struct (0.0, 0.0)
        let cellLeast row (cell: TableCell) =
            let padding = cellPaddingForRow scale table row cell
            let struct (longestWord, _) = measure cell
            if fixedLayout then padding.Left + padding.Right + 1.0 else longestWord + padding.Left + padding.Right
        // Percentages of a table of no width of its own (auto) are of the narrowest table that
        // gives each such cell its least width (fixtures tcwpct-*), as far as `maxWidth`.
        let percentBase =
            match tableWidth with
            | Some width -> Some width
            | None ->
                let mutable narrowest = 0.0
                for row in table.Rows do
                    for cell in row.Cells do
                        match cell.PreferredWidth with
                        | Some { Kind = Rbt.Word.Document.PercentWidth; Value = Some percent } when percent > 0.0 && cell.VerticalMerge <> Some ContinueMerge ->
                            narrowest <- max narrowest (cellLeast row cell * 100.0 / percent)
                        | _ -> ()
                if narrowest > 0.0 then Some (match maxWidth with Some room -> min room narrowest | None -> narrowest)
                else maxWidth
        let hasCellWidths =
            table.Rows |> List.exists (fun row -> row.Cells |> List.exists (fun cell -> (cellPreferredWidth scale percentBase cell).IsSome))
        if table.PreferredWidth.IsNone && table.GridColumns.IsEmpty && not hasCellWidths then
            List.replicate columns ((maxWidth |> Option.map (fun width -> width / float columns)) |> Option.defaultValue scale.DefaultColumnWidth), 0.0
        else
        let preferred: float option array = Array.create columns None
        let least = Array.create columns 0.0
        let widen (slot: float option array) column value =
            slot.[column] <- Some (match slot.[column] with Some current -> max current value | None -> value)
        let spanning = ResizeArray<int * int * float option * float>()
        for row in table.Rows do
            let mutable column = max 0 row.GridBefore
            for cell in row.Cells do
                let span = cellSpan cell
                if cell.VerticalMerge <> Some ContinueMerge && column < columns then
                    let padding = cellPaddingForRow scale table row cell
                    let margins = padding.Left + padding.Right
                    let struct (_, oneLine) = measure cell
                    let cellLeast = cellLeast row cell
                    let width =
                        match cellPreferredWidth scale percentBase cell with
                        | Some width -> Some width
                        | None when fixedLayout || content.IsNone -> None
                        | None -> Some (oneLine + margins)
                    if span = 1 then
                        least.[column] <- max least.[column] cellLeast
                        match width |> Option.orElse (gridAt column) with
                        | Some width -> widen preferred column width
                        | None -> ()
                    else
                        spanning.Add((column, min columns (column + span) - 1, width, cellLeast))
                elif span = 1 && column < columns then
                    // (A continuation's width counts for nothing, but the merge's paragraphs the
                    // import hands out to it - distributeVerticalMergeContent - are the merge's
                    // content: the calibre family's nested "Three".)
                    least.[column] <- max least.[column] (cellLeast row cell)
                column <- column + span
        for first, last, width, cellLeast in spanning do
            let mutable known = 0.0
            let mutable unknown = 0
            let mutable leastSum = 0.0
            for column in first .. last do
                leastSum <- leastSum + least.[column]
                match preferred.[column] with
                | Some value -> known <- known + value
                | None -> unknown <- unknown + 1
            if leastSum < cellLeast then
                least.[last] <- least.[last] + cellLeast - leastSum
            match width with
            | Some width when unknown > 0 ->
                let share = max 0.0 (width - known) / float unknown
                for column in first .. last do
                    if preferred.[column].IsNone then preferred.[column] <- Some share
            | Some width when known < width -> preferred.[last] <- Some (preferred.[last].Value + width - known)
            | _ -> ()
        let widths =
            Array.init columns (fun column ->
                let value = preferred.[column] |> Option.orElse (gridAt column) |> Option.defaultValue least.[column]
                if fixedLayout then value else max value least.[column])
        let total = Array.sum widths
        let leastTotal = Array.sum least
        let target =
            match tableWidth with
            | Some width -> if fixedLayout then width else max width leastTotal
            | None when fixedLayout -> total
            | None ->
                match maxWidth with
                | Some room -> max leastTotal (min total room)
                | None -> total
        let tableLeast =
            // (A percentage resolves to nothing without a room: fixtures nested-width nest-pct.)
            match table.PreferredWidth |> Option.map preferredWidthFromDocument |> Option.bind (resolvePreferredWidth scale None) with
            | Some width -> max width leastTotal
            | None -> leastTotal
        let columnWidths =
            if abs (target - total) <= 0.01 || total <= 0.0 then
                arrayToList widths
            elif target < total && target >= leastTotal && total - leastTotal > 0.0 then
                let factor = (target - leastTotal) / (total - leastTotal)
                widths |> Array.mapi (fun column width -> least.[column] + (width - least.[column]) * factor) |> arrayToList
            else
                widths |> Array.map (fun width -> width * target / total) |> arrayToList
        columnWidths, tableLeast

    let tableColumnWidthsWith scale (maxWidth: float option) (content: (TableCell -> struct (float * float)) option) (table: TableBlock) =
        tableColumnWidthsAndLeast scale (table.LayoutMode = Some FixedTableLayout) maxWidth content table |> fst

    let tableColumnWidths scale maxWidth table = tableColumnWidthsWith scale maxWidth None table

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
        |> Option.map (fun twips -> max 0.0 (twips / scale.DxaPerUnit))
        |> Option.defaultValue (defaultRowHeight lineHeight)

    /// Whether a row may break across pages: not a cantSplit one, nor one of an exact height (Word
    /// 2019: fixtures row-break height-exact100_f27 - a 100pt row of 20 lines on 100pt of room stays
    /// whole, its lines past the row drawn over the margin - and _f28-30, moved whole).
    let rowHeightIsExact (row: TableRow) =
        row.HeightRule = Some "exact" && (row.Height |> Option.exists (fun height -> height > 0.0))

    let rowCanSplit (row: TableRow) =
        row.CanSplit <> Some false && not (rowHeightIsExact row)

    /// The height a row has to have room for on a page to start there (w:trHeight, at-least or exact;
    /// not hRule auto): Word moves a row whole when it does not (fixtures row-break height-least50_f30,
    /// height-least100_f28-30, height-least100-short_f28-30; _f27 of 100pt room keeps it).
    let rowHeightToStart scale lineHeight (row: TableRow) =
        if row.HeightRule = Some "auto" then 0.0 else rowMinHeight scale lineHeight row

    /// The least a row takes with the borders above and below it (rowBorderBands): an exact
    /// height holds the border above, an at-least height is the content's and the borders come
    /// on top of it (Word: fixtures table-border, exact-sz24, atleast-sz24).
    let rowMinHeightWithBands scale lineHeight (row: TableRow) (top: float, bottom: float) =
        let height = rowMinHeight scale lineHeight row
        if height <= 0.0 then 0.0
        elif row.HeightRule = Some "exact" then max height top + bottom
        else height + top + bottom

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
            match Rbt.BorderGrid.styleKey value.Style with
            | "nil"
            | "none"
            | "hidden" -> 0.0
            | _ -> value.Width |> Option.defaultValue 1.0
        | None -> 0.0

    /// The widths of a table's left and right outer borders (its first and last cells' of its
    /// first row, or the table's own), canvas px.
    let outerSideBorderWidths (table: TableBlock) =
        let tableBorders = table.Borders
        match table.Rows with
        | row :: _ when not row.Cells.IsEmpty ->
            let first = List.head row.Cells
            let last = List.last row.Cells
            let side (cell: TableCell) pick tablePick =
                match cell.Borders |> Option.bind pick with
                | Some border -> visibleBorderWidth (Some border)
                | None ->
                    match tableBorders |> Option.bind tablePick with
                    | Some border -> visibleBorderWidth (Some border)
                    | None -> cell.BorderWidth |> Option.defaultValue 0.0
            side first (fun (b: TableBorders) -> b.Left) (fun (b: TableBorders) -> b.Left),
            side last (fun (b: TableBorders) -> b.Right) (fun (b: TableBorders) -> b.Right)
        | _ -> 0.0, 0.0

    /// What a horizontal border takes in the flow of rows: its width - three widths for a
    /// double line, five for a triple (Word: fixtures table-border).
    let horizontalBorderAdvance (border: TableBorder option) =
        let width = visibleBorderWidth border
        match border |> Option.map (fun value -> Rbt.BorderGrid.styleKey value.Style) with
        | Some "double" -> 3.0 * width
        | Some "triple" -> 5.0 * width
        | _ -> width

    /// The rows of one table by reference: a row's index, and the rows in order.
    type private RowIndex = { IndexedTable: TableBlock; Indexes: Dictionary<TableRow, int>; Rows: TableRow array }

    // (The index of the table last asked about. Layouts run on several threads: an index is never
    // changed once made, a new one replaces it whole, and a reader keeps the one it took.)
    let mutable private lastRowIndex : RowIndex option = None

    /// Whether the row is the table's last, and the row before it (None: the first, or not one of the table's rows).
    let private rowPlaceIn (table: TableBlock) (row: TableRow) =
        match table.Rows with
        | [] | [ _ ] -> true, None
        | first :: _ when obj.ReferenceEquals(first, row) -> false, None
        | _ ->
            let index =
                match lastRowIndex with
                | Some known when obj.ReferenceEquals(known.IndexedTable, table) -> known
                | _ ->
                    let rows = List.toArray table.Rows
                    let indexes = Dictionary<TableRow, int>(HashIdentity.Reference)
                    rows |> Array.iteri (fun position tableRow -> indexes.[tableRow] <- position)
                    let made = { IndexedTable = table; Indexes = indexes; Rows = rows }
                    lastRowIndex <- Some made
                    made
            match index.Indexes.TryGetValue row with
            | true, position -> position = index.Rows.Length - 1, (if position > 0 then Some index.Rows.[position - 1] else None)
            | _ -> false, None

    /// The borders a row's content is below and the table's last row above (Word: fixtures
    /// table-border): above a row, the widest horizontal border on its top edge - the table's
    /// top border (first row) or inside one, or a cell's own top border, or the bottom border
    /// of a cell of the row before; below the last row, the widest bottom border. A row that
    /// is not one of the table's (a header row's copy) is taken as its first. Tables with
    /// cell spacing draw each cell's own frame: none.
    let rowBorderBands (table: TableBlock) (row: TableRow) =
        let tableBorders = row.BordersOverride |> Option.orElse table.Borders
        // The widest of the cells' borders on one side, a cell without its own taking the table's.
        let edge (cells: TableCell list) (skipContinued: bool) (own: TableBorders -> TableBorder option) (fallback: TableBorders -> TableBorder option) =
            let fallbackWidth = match tableBorders with Some borders -> horizontalBorderAdvance (fallback borders) | None -> 0.0
            let mutable widest = 0.0
            for cell in cells do
                if not (skipContinued && cell.VerticalMerge = Some ContinueMerge) then
                    let width =
                        match cell.Borders with
                        | Some borders ->
                            match own borders with
                            | Some _ as border -> horizontalBorderAdvance border
                            | None -> fallbackWidth
                        | None -> fallbackWidth
                    if width > widest then widest <- width
            widest
        let isLast, above = rowPlaceIn table row
        let top =
            match above with
            | None -> edge row.Cells true _.Top _.Top
            | Some above -> max (edge row.Cells true _.Top _.InsideH) (edge above.Cells false _.Bottom _.InsideH)
        let bottom = if isLast then edge row.Cells false _.Bottom _.Bottom else 0.0
        top, bottom

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
        match Rbt.BorderGrid.styleKey style with
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
        | Some value when Rbt.BorderGrid.styleKey value.Style = "nil" || Rbt.BorderGrid.styleKey value.Style = "none" -> None
        | Some value -> printBorderStrokeWithStyle value.Style (value.Color |> Option.orElse fallbackColor) (value.Width |> Option.orElse fallbackWidth)
        | None -> printBorderStroke fallbackColor fallbackWidth

    let private printExplicitBorderStrokeFromBorder (border: TableBorder option) fallbackColor fallbackWidth =
        match border with
        | Some value when Rbt.BorderGrid.styleKey value.Style = "nil" || Rbt.BorderGrid.styleKey value.Style = "none" -> None
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
