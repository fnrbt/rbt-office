module DocumentWriter

open Fable.Core
open Fable.Core.JsInterop
open XmlGenerator
open ZipWriter
open ZipInterop
open DocumentProcessors.EditorProcessor

[<Fable.Core.Emit("$0 === null || $0 === undefined")>]
let private isNullOrUndefined (_value: obj) : bool = jsNative

let private stringOrDefault defaultValue (value: obj) =
    if isNullOrUndefined value then defaultValue
    else
        match unbox<string> value with
        | null | "" -> defaultValue
        | value -> value

let private optionalString (value: obj) =
    if isNullOrUndefined value then None
    else
        match unbox<string> value with
        | null | "" -> None
        | value -> Some value

let private optionalArray (value: obj) : obj array =
    if isNullOrUndefined value then [||]
    else unbox<obj array> value

let private intOrDefault defaultValue (value: obj) =
    if isNullOrUndefined value then defaultValue
    else unbox<int> value

let private optionalInt (value: obj) =
    if isNullOrUndefined value then None
    else Some (unbox<int> value)

let private boolOrDefault defaultValue (value: obj) =
    if isNullOrUndefined value then defaultValue
    else unbox<bool> value

let private optionalBool (value: obj) =
    if isNullOrUndefined value then None
    else Some (unbox<bool> value)

let private optionalIntString (value: obj) =
    optionalString value
    |> Option.bind (fun text ->
        match System.Int32.TryParse(text) with
        | true, parsed -> Some (string parsed)
        | false, _ -> None)

let private optionalFloat (value: obj) =
    if isNullOrUndefined value then None
    else Some (unbox<float> value)

let private floatOrDefault defaultValue (value: obj) =
    optionalFloat value |> Option.defaultValue defaultValue

let private pixelsToTwips (value: float) =
    int (value * 15.0) |> string

let private pixelsToPoints (value: float) =
    value * 72.0 / 96.0

type PageSetupInfo = {
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

let private defaultPageSetupInfo = {
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

let private optionalPageSetup (value: obj) =
    if isNullOrUndefined value then None
    else
        Some {
            PageWidth = floatOrDefault defaultPageSetupInfo.PageWidth (value?PageWidth)
            PageHeight = floatOrDefault defaultPageSetupInfo.PageHeight (value?PageHeight)
            MarginLeft = floatOrDefault defaultPageSetupInfo.MarginLeft (value?MarginLeft)
            MarginRight = floatOrDefault defaultPageSetupInfo.MarginRight (value?MarginRight)
            MarginTop = floatOrDefault defaultPageSetupInfo.MarginTop (value?MarginTop)
            MarginBottom = floatOrDefault defaultPageSetupInfo.MarginBottom (value?MarginBottom)
            HeaderDistance = floatOrDefault defaultPageSetupInfo.HeaderDistance (value?HeaderDistance)
            FooterDistance = floatOrDefault defaultPageSetupInfo.FooterDistance (value?FooterDistance)
            DifferentFirstPage = boolOrDefault defaultPageSetupInfo.DifferentFirstPage (value?DifferentFirstPage)
            EvenAndOddHeaders = boolOrDefault defaultPageSetupInfo.EvenAndOddHeaders (value?EvenAndOddHeaders)
        }

let private sectionPageSetupChildren (setup: PageSetupInfo) =
    [ yield selfClosing "w:pgSz" [("w:w", pixelsToTwips setup.PageWidth); ("w:h", pixelsToTwips setup.PageHeight)]
      yield selfClosing "w:pgMar" [
          ("w:top", pixelsToTwips setup.MarginTop)
          ("w:right", pixelsToTwips setup.MarginRight)
          ("w:bottom", pixelsToTwips setup.MarginBottom)
          ("w:left", pixelsToTwips setup.MarginLeft)
          ("w:header", pixelsToTwips setup.HeaderDistance)
          ("w:footer", pixelsToTwips setup.FooterDistance)
          ("w:gutter", "0")
      ]
      if setup.DifferentFirstPage then yield selfClosing "w:titlePg" [] ]

/// Type conversion helpers (inverse of TypeConversion in DocumentProcessors)
module InverseConversion =
    /// Convert points to Word half-points (Word uses half-points for font sizes)
    let pointsToHalfPoints (points: float) : string =
        string (int (points * 2.0))

    /// Convert CSS color (#RRGGBB) to Word format (RRGGBB)
    let cssColorToWord (color: string option) : string option =
        color
        |> Option.map (fun c ->
            if c.StartsWith("#") then
                c.Substring(1).ToUpper()
            else
                c.ToUpper()
        )

    /// Convert web alignment to Word alignment
    let alignmentToWord (alignment: string) : string =
        match alignment.ToLower() with
        | "center" -> "center"
        | "right" -> "right"
        | "justify" -> "both"
        | _ -> "left"

    let pointsToTwips (points: float) : string =
        string (int (points * 20.0))

    let lineSpacingToWord (spacing: float) : string =
        string (int (spacing * 240.0))

/// Generate run properties XML
let generateRunProps (run: EditorRun) : string list =
    let props = []

    let props = if run.Bold then (selfClosing "w:b" []) :: props else props
    let props = if run.Italic then (selfClosing "w:i" []) :: props else props
    let props = if run.Underline then (selfClosing "w:u" [("w:val", "single")]) :: props else props
    let props = if run.Strike then (selfClosing "w:strike" []) :: props else props
    let props =
        match run.Highlight with
        | Some color -> (selfClosing "w:highlight" [("w:val", color)]) :: props
        | None -> props

    let props =
        match run.VertAlign with
        | Some value when value = "superscript" || value = "subscript" -> (selfClosing "w:vertAlign" [("w:val", value)]) :: props
        | _ -> props

    let props =
        let sz = InverseConversion.pointsToHalfPoints run.FontSize
        (selfClosing "w:sz" [("w:val", sz)]) ::
        (selfClosing "w:szCs" [("w:val", sz)]) :: props

    let props =
        if abs run.CharacterSpacing >= 0.001 then
            (selfClosing "w:spacing" [("w:val", string (int (run.CharacterSpacing * 15.0)))]) :: props
        else props

    let props =
        (selfClosing "w:rFonts" [("w:ascii", run.FontFamily); ("w:hAnsi", run.FontFamily); ("w:cs", run.FontFamily)]) :: props

    let props =
        match InverseConversion.cssColorToWord run.Color with
        | Some color -> (selfClosing "w:color" [("w:val", color)]) :: props
        | None -> props

    List.rev props

/// Generate a run element with formatting
let generateRun (run: EditorRun) : string =
    let rPr = generateRunProps run
    let rPrElement = if List.isEmpty rPr then "" else element "w:rPr" [] rPr
    let tElement = elementWithText "w:t" [("xml:space", "preserve")] run.Text

    element "w:r" [] [rPrElement; tElement]

let private generateRunWithObjectProps (runObj: obj) (run: EditorRun) : string =
    let rPr =
        generateRunProps run
        @ [ match optionalString (runObj?Style) |> Option.orElse (optionalString (runObj?RunStyle)) |> Option.orElse (optionalString (runObj?CharacterStyle)) with
            | Some styleId when styleId <> "" -> yield selfClosing "w:rStyle" [("w:val", styleId)]
            | _ -> ()
            if boolOrDefault false (runObj?AllCaps) then yield selfClosing "w:caps" []
            if boolOrDefault false (runObj?SmallCaps) then yield selfClosing "w:smallCaps" [] ]
    let rPrElement = if List.isEmpty rPr then "" else element "w:rPr" [] rPr
    let tElement = elementWithText "w:t" [("xml:space", "preserve")] run.Text
    element "w:r" [] [rPrElement; tElement]

let private generateFieldRun (runObj: obj) (fallbackRun: EditorRun) (fieldCode: string) =
    let result = optionalString (runObj?FieldResult) |> Option.defaultValue fallbackRun.Text
    let resultRun = generateRunWithObjectProps runObj { fallbackRun with Text = result }
    element "w:fldSimple" [("w:instr", fieldCode)] [resultRun]

let private bookmarkStartXml (runObj: obj) =
    optionalString (runObj?BookmarkStartId)
    |> Option.orElse (optionalString (runObj?BookmarkId))
    |> Option.map (fun id ->
        let name = optionalString (runObj?BookmarkName) |> Option.defaultValue ("bookmark" + id)
        selfClosing "w:bookmarkStart" [("w:id", id); ("w:name", name)])

let private bookmarkEndXml (runObj: obj) =
    let explicitEnd =
        optionalString (runObj?BookmarkEndId)
        |> Option.orElse (optionalString (runObj?BookmarkId))
    let fieldWrappedEnd =
        optionalString (runObj?FieldCode)
        |> Option.bind (fun _ -> optionalString (runObj?BookmarkStartId))
    explicitEnd
    |> Option.orElse fieldWrappedEnd
    |> Option.map (fun id -> selfClosing "w:bookmarkEnd" [("w:id", id)])

let private generateReferenceRun (runObj: obj) (fallbackRun: EditorRun) =
    let inner =
        if boolOrDefault false (runObj?IsEquation) then
            optionalString (runObj?EquationXml) |> Option.defaultValue ""
        else
            match optionalString (runObj?FieldCode) with
            | Some fieldCode when fieldCode <> "" -> generateFieldRun runObj fallbackRun fieldCode
            | _ ->
                match optionalString (runObj?CommentId), optionalString (runObj?NoteId) with
                | Some commentId, _ ->
                    element "w:r" [] [selfClosing "w:commentReference" [("w:id", commentId)]]
                | None, Some noteId ->
                    let noteType = optionalString (runObj?NoteType) |> Option.defaultValue "footnote"
                    let elementName = if noteType = "endnote" then "w:endnoteReference" else "w:footnoteReference"
                    element "w:r" [] [selfClosing elementName [("w:id", noteId)]]
                | None, None ->
                    if fallbackRun.Text = "" && bookmarkStartXml runObj |> Option.isSome || fallbackRun.Text = "" && bookmarkEndXml runObj |> Option.isSome then ""
                    else generateRunWithObjectProps runObj fallbackRun
    [ bookmarkStartXml runObj; (if inner <> "" then Some inner else None); bookmarkEndXml runObj ]
    |> List.choose id
    |> String.concat ""

let private revisionWrappedRun (runObj: obj) (runXml: string) =
    match optionalString (runObj?RevisionType) with
    | Some revisionType when revisionType = "insert" || revisionType = "delete" ->
        let elementName = if revisionType = "delete" then "w:del" else "w:ins"
        let attrs = [
            yield ("w:id", optionalString (runObj?RevisionId) |> Option.defaultValue "1")
            match optionalString (runObj?RevisionAuthor) with
            | Some author -> yield ("w:author", author)
            | None -> ()
            match optionalString (runObj?RevisionDate) with
            | Some date -> yield ("w:date", date)
            | None -> ()
        ]
        element elementName attrs [runXml]
    | _ -> runXml

let private collectParagraphHyperlinks (para: obj) =
    if isNullOrUndefined para?Runs then
        []
    else
        para?Runs
        |> unbox<obj array>
        |> Array.toList
        |> List.choose (fun run -> optionalString (run?Hyperlink))

let rec private collectItemHyperlinks (item: obj) =
    match item?Type |> stringOrDefault "Paragraph" with
    | "TextBox" when not (isNullOrUndefined item?Paragraphs) ->
        item?Paragraphs |> unbox<obj array> |> Array.toList |> List.collect collectParagraphHyperlinks
    | "Table" when not (isNullOrUndefined item?Rows) ->
        item?Rows
        |> unbox<obj array>
        |> Array.toList
        |> List.collect (fun row ->
            if isNullOrUndefined row?Cells then
                []
            else
                row?Cells
                |> unbox<obj array>
                |> Array.toList
                |> List.collect (fun cell ->
                    if isNullOrUndefined cell?Paragraphs then
                        []
                    else
                        cell?Paragraphs
                        |> unbox<obj array>
                        |> Array.toList
                        |> List.collect collectParagraphHyperlinks))
    | _ -> collectParagraphHyperlinks item

let private hyperlinkMap (items: obj array) =
    items
    |> Array.toList
    |> List.collect collectItemHyperlinks
    |> List.distinct
    |> List.mapi (fun index url -> (url, sprintf "rId%d" (index + 3)))
    |> Map.ofList

type CommentRef = { Id: string; Text: string; Author: string option; Date: string option }
type NoteRef = { Id: string; NoteType: string; Text: string }

type ImageAsset = {
    Src: string
    RelId: string
    FileName: string
    Extension: string
    ContentType: string
    Data: byte[] option
    Linked: bool
}

type ImageScope = {
    Assets: ImageAsset array
    AssetBySrc: Map<string, ImageAsset>
}

let private emptyImageScope = { Assets = [||]; AssetBySrc = Map.empty }

[<Emit("typeof Buffer !== 'undefined' ? Buffer.from($0, 'base64') : Uint8Array.from(atob($0), c => c.charCodeAt(0))")>]
let private base64ToBytes (_base64: string) : byte[] = jsNative

let private mimeToExtensionAndContentType (mime: string) =
    match mime.ToLowerInvariant() with
    | "image/jpeg" | "image/jpg" -> Some ("jpeg", "image/jpeg")
    | "image/png" -> Some ("png", "image/png")
    | "image/gif" -> Some ("gif", "image/gif")
    | "image/bmp" -> Some ("bmp", "image/bmp")
    | "image/svg+xml" -> Some ("svg", "image/svg+xml")
    | _ -> None

let private parseDataUrl (src: string) =
    if src.StartsWith("data:") then
        let commaIndex = src.IndexOf(',')
        if commaIndex > 0 then
            let meta = src.Substring(5, commaIndex - 5)
            let data = src.Substring(commaIndex + 1)
            let parts = meta.Split(';')
            let mime = if parts.Length > 0 then parts.[0] else ""
            if meta.ToLowerInvariant().Contains("base64") then
                mimeToExtensionAndContentType mime
                |> Option.map (fun (extension, contentType) -> extension, contentType, base64ToBytes data)
            else None
        else None
    else None

let private markerText defaultValue (value: obj) =
    optionalString value |> Option.defaultValue defaultValue

let private collectParagraphComments (para: obj) =
    if isNullOrUndefined para?Runs then
        []
    else
        para?Runs
        |> unbox<obj array>
        |> Array.toList
        |> List.choose (fun run ->
            optionalString (run?CommentId)
            |> Option.map (fun id -> {
                Id = id
                Text = markerText (sprintf "Comment %s" id) (run?CommentText)
                Author = optionalString (run?CommentAuthor)
                Date = optionalString (run?CommentDate)
            }))

let private collectParagraphNotes (para: obj) =
    if isNullOrUndefined para?Runs then
        []
    else
        para?Runs
        |> unbox<obj array>
        |> Array.toList
        |> List.choose (fun run ->
            optionalString (run?NoteId)
            |> Option.map (fun id ->
                let noteType = markerText "footnote" (run?NoteType)
                { Id = id; NoteType = noteType; Text = markerText (sprintf "%s %s" noteType id) (run?NoteText) }))

let rec private collectItemParagraphs (item: obj) =
    match item?Type |> stringOrDefault "Paragraph" with
    | "TextBox" when not (isNullOrUndefined item?Paragraphs) ->
        item?Paragraphs |> unbox<obj array> |> Array.toList
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
                    if isNullOrUndefined cell?Paragraphs then []
                    else cell?Paragraphs |> unbox<obj array> |> Array.toList))
    | _ -> [item]

let private collectComments (items: obj array) =
    items
    |> Array.toList
    |> List.collect collectItemParagraphs
    |> List.collect collectParagraphComments
    |> List.distinctBy (fun comment -> comment.Id)

let private collectNotes (noteType: string) (items: obj array) =
    items
    |> Array.toList
    |> List.collect collectItemParagraphs
    |> List.collect collectParagraphNotes
    |> List.filter (fun note -> note.NoteType = noteType)
    |> List.distinctBy (fun note -> note.Id)

let private collectSectionReferenceIds (items: obj array) =
    let idsFromRefs (value: obj) =
        if isNullOrUndefined value then []
        else
            value
            |> optionalArray
            |> Array.toList
            |> List.choose (fun item -> optionalString (item?RelationshipId))

    items
    |> Array.toList
    |> List.collect collectItemParagraphs
    |> List.collect (fun para -> idsFromRefs (para?HeaderReferences) @ idsFromRefs (para?FooterReferences))
    |> Set.ofList

let private collectParagraphImageSources (para: obj) =
    if isNullOrUndefined para?Runs then []
    else
        para?Runs
        |> unbox<obj array>
        |> Array.toList
        |> List.choose (fun run ->
            if boolOrDefault false (run?IsImage) then optionalString (run?ImageSrc)
            else None)

let rec private collectItemImageSources (item: obj) =
    match item?Type |> stringOrDefault "Paragraph" with
    | "TextBox" when not (isNullOrUndefined item?Paragraphs) ->
        item?Paragraphs |> unbox<obj array> |> Array.toList |> List.collect collectParagraphImageSources
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
                    if isNullOrUndefined cell?Paragraphs then []
                    else
                        cell?Paragraphs
                        |> unbox<obj array>
                        |> Array.toList
                        |> List.collect collectParagraphImageSources))
    | _ -> collectParagraphImageSources item

let private buildImageScope (mediaStartIndex: int) (items: obj array) =
    let mutable nextMediaIndex = mediaStartIndex
    let assets =
        items
        |> Array.toList
        |> List.collect collectItemImageSources
        |> List.distinct
        |> List.choose (fun src ->
            parseDataUrl src
            |> Option.map (fun (extension, contentType, data) ->
                let fileName = sprintf "image%d.%s" nextMediaIndex extension
                let relId = sprintf "rIdImage%d" nextMediaIndex
                nextMediaIndex <- nextMediaIndex + 1
                { Src = src; RelId = relId; FileName = fileName; Extension = extension; ContentType = contentType; Data = Some data; Linked = false })
            |> Option.orElseWith (fun () ->
                if src.StartsWith("http://") || src.StartsWith("https://") then
                    let relId = sprintf "rIdImage%d" nextMediaIndex
                    nextMediaIndex <- nextMediaIndex + 1
                    Some { Src = src; RelId = relId; FileName = ""; Extension = ""; ContentType = ""; Data = None; Linked = true }
                else None))
        |> List.toArray
    { Assets = assets; AssetBySrc = assets |> Array.map (fun asset -> asset.Src, asset) |> Map.ofArray }, nextMediaIndex

let private pixelsToEmus (pixels: float) =
    int (pixels * 9525.0)

let private optionalImageSize defaultValue (value: obj) =
    optionalFloat value |> Option.defaultValue defaultValue |> pixelsToEmus |> string

let private optionalImagePosition (value: obj) =
    optionalFloat value |> Option.defaultValue 0.0 |> pixelsToEmus |> string

let private optionalImageDistance defaultValue (value: obj) =
    optionalFloat value |> Option.defaultValue defaultValue |> pixelsToEmus |> string

let private boolFlag defaultValue (value: obj) =
    if boolOrDefault defaultValue value then "1" else "0"

let private generateImageRun (asset: ImageAsset) (runObj: obj) =
    let width = optionalImageSize 96.0 (runObj?ImageWidth)
    let height = optionalImageSize 48.0 (runObj?ImageHeight)
    let alt = optionalString (runObj?ImageAlt) |> Option.defaultValue "Image"
    let imageName = optionalString (runObj?ImageName) |> Option.defaultValue asset.FileName
    let titleAttrs = optionalString (runObj?ImageTitle) |> Option.map (fun title -> [("title", title)]) |> Option.defaultValue []
    let positionX = optionalImagePosition (runObj?ImagePositionX)
    let positionY = optionalImagePosition (runObj?ImagePositionY)
    let positionHRelativeFrom = optionalString (runObj?ImagePositionHRelativeFrom) |> Option.defaultValue "column"
    let positionVRelativeFrom = optionalString (runObj?ImagePositionVRelativeFrom) |> Option.defaultValue "paragraph"
    let distanceTop = optionalImageDistance 0.0 (runObj?ImageDistanceTop)
    let distanceBottom = optionalImageDistance 0.0 (runObj?ImageDistanceBottom)
    let distanceLeft = optionalImageDistance 12.0 (runObj?ImageDistanceLeft)
    let distanceRight = optionalImageDistance 12.0 (runObj?ImageDistanceRight)
    let behindDoc = boolFlag false (runObj?ImageBehindDoc)
    let allowOverlap = boolFlag true (runObj?ImageAllowOverlap)
    let layoutInCell = boolFlag true (runObj?ImageLayoutInCell)
    let relativeHeight = optionalString (runObj?ImageRelativeHeight) |> Option.defaultValue "0"
    let locked = boolFlag false (runObj?ImageLocked)
    let simplePosition = boolFlag false (runObj?ImageSimplePosition)
    let simplePositionX = optionalImagePosition (runObj?ImageSimplePositionX)
    let simplePositionY = optionalImagePosition (runObj?ImageSimplePositionY)
    let cropAttrs =
        [ ("l", optionalString (runObj?ImageCropLeft))
          ("t", optionalString (runObj?ImageCropTop))
          ("r", optionalString (runObj?ImageCropRight))
          ("b", optionalString (runObj?ImageCropBottom)) ]
        |> List.choose (fun (name, value) -> value |> Option.map (fun text -> name, text))
    let drawingContainer =
        if boolOrDefault false (runObj?ImageFloating) then
            let wrap =
                match optionalString (runObj?ImageWrap) with
                | Some "none" -> selfClosing "wp:wrapNone" []
                | Some "topAndBottom" -> selfClosing "wp:wrapTopAndBottom" []
                | Some "tight" -> selfClosing "wp:wrapTight" [("wrapText", "bothSides")]
                | Some "through" -> selfClosing "wp:wrapThrough" [("wrapText", "bothSides")]
                | _ -> selfClosing "wp:wrapSquare" [("wrapText", "bothSides")]
            fun children ->
                element "wp:anchor" [
                    ("xmlns:wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")
                    ("distT", distanceTop); ("distB", distanceBottom); ("distL", distanceLeft); ("distR", distanceRight)
                    ("simplePos", simplePosition); ("relativeHeight", relativeHeight); ("behindDoc", behindDoc)
                    ("locked", locked); ("layoutInCell", layoutInCell); ("allowOverlap", allowOverlap)
                ] ([
                    selfClosing "wp:simplePos" [("x", simplePositionX); ("y", simplePositionY)]
                    element "wp:positionH" [("relativeFrom", positionHRelativeFrom)] [elementWithText "wp:posOffset" [] positionX]
                    element "wp:positionV" [("relativeFrom", positionVRelativeFrom)] [elementWithText "wp:posOffset" [] positionY]
                    selfClosing "wp:extent" [("cx", width); ("cy", height)]
                    wrap
                ] @ children)
        else
            fun children ->
                element "wp:inline" [("xmlns:wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")] (
                    [selfClosing "wp:extent" [("cx", width); ("cy", height)]] @ children)
    element "w:r" [] [
        element "w:drawing" [] [
            drawingContainer [
                selfClosing "wp:docPr" ([("id", asset.RelId.Replace("rIdImage", "")); ("name", imageName); ("descr", alt)] @ titleAttrs)
                element "a:graphic" [("xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main")] [
                    element "a:graphicData" [("uri", "http://schemas.openxmlformats.org/drawingml/2006/picture")] [
                        element "pic:pic" [("xmlns:pic", "http://schemas.openxmlformats.org/drawingml/2006/picture")] [
                            element "pic:nvPicPr" [] [
                                selfClosing "pic:cNvPr" ([("id", asset.RelId.Replace("rIdImage", "")); ("name", imageName); ("descr", alt)] @ titleAttrs)
                                selfClosing "pic:cNvPicPr" []
                            ]
                            element "pic:blipFill" [] [
                                selfClosing "a:blip" [((if asset.Linked then "r:link" else "r:embed"), asset.RelId)]
                                if not cropAttrs.IsEmpty then selfClosing "a:srcRect" cropAttrs
                                element "a:stretch" [] [selfClosing "a:fillRect" []]
                            ]
                            element "pic:spPr" [] [
                                element "a:xfrm" [] [
                                    selfClosing "a:off" [("x", "0"); ("y", "0")]
                                    selfClosing "a:ext" [("cx", width); ("cy", height)]
                                ]
                                selfClosing "a:prstGeom" [("prst", "rect")]
                            ]
                        ]
                    ]
                ]
            ]
        ]
    ]

let private includeHeaderFooterKind differentFirstPage evenAndOddHeaders kind =
    kind = "default"
    || (kind = "first" && differentFirstPage)
    || (kind = "even" && evenAndOddHeaders)

let private sectionReferences elementName differentFirstPage evenAndOddHeaders (value: obj) =
    if isNullOrUndefined value then []
    else
        value
        |> optionalArray
        |> Array.toList
        |> List.choose (fun item ->
            let kind = item?Kind |> stringOrDefault (item?Type |> stringOrDefault "default")
            if includeHeaderFooterKind differentFirstPage evenAndOddHeaders kind then
                optionalString (item?RelationshipId)
                |> Option.map (fun relationshipId ->
                    selfClosing elementName [("w:type", kind); ("r:id", relationshipId)])
            else None)

let private paragraphSectionProperties (packageSetup: PageSetupInfo) (para: obj) =
    let sectionBreak = optionalString (para?SectionBreak)
    let paraSetup =
        match optionalPageSetup (para?SectionPageSetup) with
        | Some setup -> Some setup
        | None -> optionalPageSetup (para?PageSetup)
    let effectiveSetup = paraSetup |> Option.defaultValue packageSetup
    let refs =
        sectionReferences "w:headerReference" effectiveSetup.DifferentFirstPage effectiveSetup.EvenAndOddHeaders (para?HeaderReferences)
        @ sectionReferences "w:footerReference" effectiveSetup.DifferentFirstPage effectiveSetup.EvenAndOddHeaders (para?FooterReferences)
    let pageNumberStart = optionalInt (para?SectionPageNumberStart)
    let children = [
        yield! refs
        match paraSetup with
        | Some setup -> yield! sectionPageSetupChildren setup
        | None -> ()
        match pageNumberStart with
        | Some start -> yield selfClosing "w:pgNumType" [("w:start", string (max 1 start))]
        | None -> ()
        match sectionBreak with
        | Some breakType when breakType = "continuous" || breakType = "nextPage" ->
            yield selfClosing "w:type" [("w:val", breakType)]
        | _ -> ()
    ]
    if children.IsEmpty then None else Some (element "w:sectPr" [] children)

let private generateParaPropsWithPageSetup (packageSetup: PageSetupInfo) (para: obj) : string =
    let alignment = para?Alignment |> stringOrDefault "left" |> InverseConversion.alignmentToWord
    let jcElement = selfClosing "w:jc" [("w:val", alignment)]
    let spacingAttributes = [
        match optionalFloat (para?SpacingBefore) with
        | Some value -> yield ("w:before", InverseConversion.pointsToTwips value)
        | None -> ()
        match optionalFloat (para?SpacingAfter) with
        | Some value -> yield ("w:after", InverseConversion.pointsToTwips value)
        | None -> ()
        match optionalFloat (para?LineSpacing) with
        | Some value ->
            yield ("w:line", InverseConversion.lineSpacingToWord value)
            yield ("w:lineRule", "auto")
        | None -> ()
    ]
    let spacingElement =
        if spacingAttributes.IsEmpty then None
        else Some (selfClosing "w:spacing" spacingAttributes)
    let indentAttributes = [
        match optionalFloat (para?IndentLeft) with
        | Some value -> yield ("w:left", InverseConversion.pointsToTwips value)
        | None -> ()
        match optionalFloat (para?IndentRight) with
        | Some value -> yield ("w:right", InverseConversion.pointsToTwips value)
        | None -> ()
        match optionalFloat (para?FirstLineIndent) with
        | Some value when value > 0.0 -> yield ("w:firstLine", InverseConversion.pointsToTwips value)
        | Some value when value < 0.0 -> yield ("w:hanging", InverseConversion.pointsToTwips -value)
        | _ -> ()
    ]
    let indentElement =
        if indentAttributes.IsEmpty then None
        else Some (selfClosing "w:ind" indentAttributes)
    let listStyle = para?ListStyle |> stringOrDefault "none"
    let listLevel = para?ListLevel |> intOrDefault 0
    let itemType = para?Type |> stringOrDefault "Paragraph"
    let headingLevel = para?HeadingLevel |> intOrDefault 0
    let styleElement =
        match optionalString (para?Style) with
        | Some styleId when styleId <> "" -> Some (selfClosing "w:pStyle" [("w:val", styleId)])
        | _ when itemType = "Heading" && headingLevel >= 1 && headingLevel <= 3 ->
            Some (selfClosing "w:pStyle" [("w:val", sprintf "Heading%d" headingLevel)])
        | _ -> None
    let backgroundElement =
        optionalString (para?BackgroundColor)
        |> Option.map (fun color ->
            let wordColor = InverseConversion.cssColorToWord (Some color) |> Option.defaultValue color
            selfClosing "w:shd" [("w:val", "clear"); ("w:color", "auto"); ("w:fill", wordColor)])

    let children =
        if listStyle = "bullet" || listStyle = "numbered" then
            let numId = optionalIntString (para?ListId) |> Option.defaultValue (if listStyle = "bullet" then "1" else "2")
            let ilvl = string (max 0 (listLevel - 1))
            [
                element "w:numPr" [] [
                    selfClosing "w:ilvl" [("w:val", ilvl)]
                    selfClosing "w:numId" [("w:val", numId)]
                ]
                jcElement
            ]
        else
            [jcElement]
        |> fun children ->
            match spacingElement with
            | Some spacing -> spacing :: children
            | None -> children
        |> fun children ->
            match indentElement with
            | Some indent -> indent :: children
            | None -> children
        |> fun children ->
            match styleElement with
            | Some style -> style :: children
            | None -> children
        |> fun children ->
            match backgroundElement with
            | Some background -> background :: children
            | None -> children
        |> fun children ->
            match paragraphSectionProperties packageSetup para with
            | Some sectPr -> children @ [sectPr]
            | None -> children

    element "w:pPr" [] children

/// Generate paragraph properties XML (handles JS objects)
let generateParaProps (para: obj) : string =
    generateParaPropsWithPageSetup defaultPageSetupInfo para

/// Generate a paragraph element with formatting (handles JS objects)
let private generateParagraphWithHyperlinksAndImagesWithPageSetup (packageSetup: PageSetupInfo) (hyperlinks: Map<string, string>) (images: ImageScope) (para: obj) : string =
    let pPr = generateParaPropsWithPageSetup packageSetup para
    let runsArray = para?Runs |> unbox<obj array>
    let runs =
        runsArray
        |> Array.toList
        |> List.map (fun runObj ->
            let runXml =
                if boolOrDefault false (runObj?IsImage) then
                    match optionalString (runObj?ImageSrc) |> Option.bind (fun src -> Map.tryFind src images.AssetBySrc) with
                    | Some asset -> generateImageRun asset runObj
                    | None ->
                        let alt = optionalString (runObj?ImageAlt) |> Option.defaultValue "Image"
                        element "w:r" [] [elementWithText "w:t" [("xml:space", "preserve")] alt]
                else
                    let run: EditorRun = {
                        Text = runObj?Text |> unbox<string>
                        Bold = boolOrDefault false (runObj?Bold)
                        Italic = boolOrDefault false (runObj?Italic)
                        Underline = boolOrDefault false (runObj?Underline)
                        Strike = boolOrDefault false (runObj?Strike)
                        FontSize = if isNullOrUndefined runObj?FontSize then 11.0 else unbox<float> runObj?FontSize
                        CharacterSpacing = optionalFloat (runObj?CharacterSpacing) |> Option.defaultValue 0.0
                        FontFamily = runObj?FontFamily |> stringOrDefault "Calibri"
                        Color = if isNullOrUndefined runObj?Color then None else unbox<string option> runObj?Color
                        Highlight = optionalString (runObj?Highlight)
                        VertAlign = optionalString (runObj?VertAlign)
                        CommentId = optionalString (runObj?CommentId)
                        CommentText = optionalString (runObj?CommentText)
                        CommentAuthor = optionalString (runObj?CommentAuthor)
                        CommentDate = optionalString (runObj?CommentDate)
                        NoteId = optionalString (runObj?NoteId)
                        NoteType = optionalString (runObj?NoteType)
                        NoteText = optionalString (runObj?NoteText)
                        Hyperlink = optionalString (runObj?Hyperlink)
                        BookmarkStartId = optionalString (runObj?BookmarkStartId)
                        BookmarkName = optionalString (runObj?BookmarkName)
                        BookmarkEndId = optionalString (runObj?BookmarkEndId)
                        FieldCode = optionalString (runObj?FieldCode)
                        FieldResult = optionalString (runObj?FieldResult)
                        RenderedPageBreakBefore = boolOrDefault false (runObj?RenderedPageBreakBefore)
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
                    }
                    generateReferenceRun runObj run
            let runXml = revisionWrappedRun runObj runXml
            match optionalString (runObj?Hyperlink) |> Option.bind (fun url -> Map.tryFind url hyperlinks) with
            | Some relId -> element "w:hyperlink" [("r:id", relId); ("w:history", "1")] [runXml]
            | None -> runXml
        )
    let runs =
        if boolOrDefault false (para?PageBreakAfter) then
            runs @ [element "w:r" [] [selfClosing "w:br" [("w:type", "page")]]]
        else runs

    element "w:p" [] (pPr :: runs)

let private generateParagraphWithHyperlinksAndImages (hyperlinks: Map<string, string>) (images: ImageScope) (para: obj) : string =
    generateParagraphWithHyperlinksAndImagesWithPageSetup defaultPageSetupInfo hyperlinks images para

let private generateParagraphWithHyperlinks (hyperlinks: Map<string, string>) (para: obj) : string =
    generateParagraphWithHyperlinksAndImages hyperlinks emptyImageScope para

let private defaultTableBorders () =
    element "w:tblBorders" [] [
        selfClosing "w:top" [("w:val", "single"); ("w:sz", "4"); ("w:color", "auto")]
        selfClosing "w:left" [("w:val", "single"); ("w:sz", "4"); ("w:color", "auto")]
        selfClosing "w:bottom" [("w:val", "single"); ("w:sz", "4"); ("w:color", "auto")]
        selfClosing "w:right" [("w:val", "single"); ("w:sz", "4"); ("w:color", "auto")]
        selfClosing "w:insideH" [("w:val", "single"); ("w:sz", "4"); ("w:color", "auto")]
        selfClosing "w:insideV" [("w:val", "single"); ("w:sz", "4"); ("w:color", "auto")]
    ]

let private borderElement name (value: obj) =
    if isNullOrUndefined value then None
    else
        let style = optionalString (value?Style) |> Option.defaultValue "single"
        let color =
            optionalString (value?Color)
            |> InverseConversion.cssColorToWord
            |> Option.defaultValue "auto"
        let width =
            match optionalFloat (value?Width) with
            | Some amount -> string (int (amount * 8.0))
            | None -> if style = "none" || style = "nil" then "0" else "4"
        Some (selfClosing name [("w:val", if style = "none" then "nil" else style); ("w:sz", width); ("w:color", color)])

let private bordersElement elementName (value: obj) =
    if isNullOrUndefined value then None
    else
        let sides =
            [ borderElement "w:top" (value?Top)
              borderElement "w:right" (value?Right)
              borderElement "w:bottom" (value?Bottom)
              borderElement "w:left" (value?Left) ]
            |> List.choose id
        if sides.IsEmpty then None else Some (element elementName [] sides)

let private tableWidthValue widthType width =
    match widthType with
    | "pct" -> int (width * 50.0) |> string
    | "pt" -> InverseConversion.pointsToTwips width
    | "auto" | "nil" -> "0"
    | _ -> int width |> string

let private optionalTableWidth elementName value typeValue =
    match optionalFloat value, optionalString typeValue with
    | Some width, Some widthType -> Some (selfClosing elementName [("w:w", tableWidthValue widthType width); ("w:type", widthType)])
    | Some width, None -> Some (selfClosing elementName [("w:w", string (int width)); ("w:type", "dxa")])
    | None, Some widthType when widthType = "auto" || widthType = "nil" -> Some (selfClosing elementName [("w:w", "0"); ("w:type", widthType)])
    | _ -> None

let private wordColorOrAuto value =
    optionalString value
    |> Option.map (fun color ->
        if color = "auto" then "auto"
        else InverseConversion.cssColorToWord (Some color) |> Option.defaultValue color)
    |> Option.defaultValue "auto"

let private shadingElement (value: obj) =
    if isNullOrUndefined value then None
    else
        let pattern = optionalString (value?Pattern) |> Option.defaultValue "clear"
        let foreground = wordColorOrAuto (value?ForegroundColor)
        let background = wordColorOrAuto (value?BackgroundColor)
        Some (selfClosing "w:shd" [("w:val", pattern); ("w:color", foreground); ("w:fill", background)])

let private backgroundShadingElement (value: obj) =
    optionalString value
    |> Option.map (fun color ->
        let wordColor = InverseConversion.cssColorToWord (Some color) |> Option.defaultValue color
        selfClosing "w:shd" [("w:val", "clear"); ("w:color", "auto"); ("w:fill", wordColor)])

let private generateTableCell (hyperlinks: Map<string, string>) (images: ImageScope) (cell: obj) : string =
    let paragraphs =
        if isNullOrUndefined cell?Paragraphs then
            [element "w:p" [] []]
        else
            let paraArray = cell?Paragraphs |> unbox<obj array>
            if paraArray.Length = 0 then
                [element "w:p" [] []]
            else
                paraArray |> Array.toList |> List.map (generateParagraphWithHyperlinksAndImages hyperlinks images)

    let tcPrChildren = [
        match optionalTableWidth "w:tcW" (cell?PreferredWidth) (cell?PreferredWidthType) with
        | Some width -> yield width
        | None ->
            if not (isNullOrUndefined cell?Width) then
                let width = cell?Width |> unbox<float> |> int |> string
                yield selfClosing "w:tcW" [("w:w", width); ("w:type", "dxa")]

        match shadingElement (cell?Shading) |> Option.orElseWith (fun () -> backgroundShadingElement (cell?BackgroundColor)) with
        | Some shading -> yield shading
        | None -> ()

        match bordersElement "w:tcBorders" (cell?Borders) with
        | Some borders -> yield borders
        | None -> ()

        if isNullOrUndefined cell?Borders && (not (isNullOrUndefined cell?BorderColor) || not (isNullOrUndefined cell?BorderWidth)) then
            let color =
                if isNullOrUndefined cell?BorderColor then "auto"
                else
                    match unbox<string option> cell?BorderColor with
                    | Some value when value <> "" -> InverseConversion.cssColorToWord (Some value) |> Option.defaultValue value
                    | _ -> "auto"
            let width =
                if isNullOrUndefined cell?BorderWidth then "8"
                else cell?BorderWidth |> unbox<float> |> (*) 8.0 |> int |> string
            let border name = selfClosing name [("w:val", "single"); ("w:sz", width); ("w:color", color)]
            yield element "w:tcBorders" [] [
                border "w:top"
                border "w:left"
                border "w:bottom"
                border "w:right"
            ]

        if not (isNullOrUndefined cell?VerticalAlignment) then
            match unbox<string option> cell?VerticalAlignment with
            | Some value when value = "top" || value = "center" || value = "bottom" ->
                yield selfClosing "w:vAlign" [("w:val", value)]
            | _ -> ()

        let padding name (value: obj) =
            match optionalFloat value with
            | Some amount when amount >= 0.0 -> Some (selfClosing name [("w:w", string (int amount)); ("w:type", "dxa")])
            | _ -> None
        let paddingChildren =
            [ padding "w:top" cell?PaddingTop
              padding "w:right" cell?PaddingRight
              padding "w:bottom" cell?PaddingBottom
              padding "w:left" cell?PaddingLeft ]
            |> List.choose id
        if not paddingChildren.IsEmpty then
            yield element "w:tcMar" [] paddingChildren

        if not (isNullOrUndefined cell?ColSpan) then
            let span = cell?ColSpan |> unbox<int>
            if span > 1 then
                yield selfClosing "w:gridSpan" [("w:val", string span)]

        match optionalString (cell?VerticalMerge) with
        | Some value when value = "restart" || value = "continue" ->
            yield selfClosing "w:vMerge" [("w:val", value)]
        | _ -> ()
    ]

    element "w:tc" [] [
        element "w:tcPr" [] tcPrChildren
        paragraphs |> String.concat ""
    ]

let private generateTableRow (hyperlinks: Map<string, string>) (images: ImageScope) (row: obj) : string =
    let cells =
        if isNullOrUndefined row?Cells then
            ""
        else
            row?Cells
            |> unbox<obj array>
            |> Array.map (generateTableCell hyperlinks images)
            |> String.concat ""

    let rowPropChildren = [
        if not (isNullOrUndefined row?IsHeader) && unbox<bool> row?IsHeader then
            yield selfClosing "w:tblHeader" [("w:val", "1")]
        match optionalFloat (row?Height) with
        | Some height ->
            let attrs =
                [("w:val", string (int height))]
                @ (optionalString (row?HeightRule) |> Option.map (fun rule -> [("w:hRule", rule)]) |> Option.defaultValue [])
            yield selfClosing "w:trHeight" attrs
        | None -> ()
        if not (isNullOrUndefined row?CanSplit) && not (unbox<bool> row?CanSplit) then
            yield selfClosing "w:cantSplit" []
    ]
    let rowProps = if rowPropChildren.IsEmpty then "" else element "w:trPr" [] rowPropChildren

    element "w:tr" [] [rowProps; cells]

let private generateTableGrid (table: obj) =
    if isNullOrUndefined table?GridColumns then ""
    else
        let columns =
            table?GridColumns
            |> unbox<obj array>
            |> Array.toList
            |> List.choose (fun value -> optionalFloat value)
            |> List.map (fun width -> selfClosing "w:gridCol" [("w:w", string (int width))])
        if columns.IsEmpty then "" else element "w:tblGrid" [] columns

let private generateTable (hyperlinks: Map<string, string>) (images: ImageScope) (table: obj) : string =
    let rows =
        if isNullOrUndefined table?Rows then
            ""
        else
            table?Rows
            |> unbox<obj array>
            |> Array.map (generateTableRow hyperlinks images)
            |> String.concat ""

    let tblPrChildren = [
        match optionalTableWidth "w:tblW" (table?PreferredWidth) (table?PreferredWidthType) with
        | Some width -> yield width
        | None -> ()
        match optionalString (table?LayoutMode) with
        | Some mode when mode = "fixed" || mode = "autofit" || mode = "autoFit" ->
            let docxMode = if mode = "fixed" then "fixed" else "autofit"
            yield selfClosing "w:tblLayout" [("w:type", docxMode)]
        | _ -> ()
        match shadingElement (table?Shading) |> Option.orElseWith (fun () -> backgroundShadingElement (table?BackgroundColor)) with
        | Some shading -> yield shading
        | None -> ()
        match bordersElement "w:tblBorders" (table?Borders) with
        | Some borders -> yield borders
        | None -> yield defaultTableBorders ()
    ]

    element "w:tbl" [] [
        element "w:tblPr" [] tblPrChildren
        generateTableGrid table
        rows
    ]

let private generateTextBox (packageSetup: PageSetupInfo) (hyperlinks: Map<string, string>) (images: ImageScope) (item: obj) : string =
    let width = floatOrDefault 220.0 (item?Width) |> pixelsToPoints
    let height = floatOrDefault 72.0 (item?Height) |> pixelsToPoints
    let x = optionalFloat (item?X) |> Option.map pixelsToPoints |> Option.defaultValue 0.0
    let y = optionalFloat (item?Y) |> Option.map pixelsToPoints |> Option.defaultValue 0.0
    let fill = optionalString (item?FillColor) |> Option.defaultValue "#FFFFFF"
    let stroke = optionalString (item?BorderColor) |> Option.defaultValue "#808080"
    let strokeWeight = optionalFloat (item?BorderWidth) |> Option.map pixelsToPoints |> Option.defaultValue 0.75
    let shapeId = optionalString (item?ShapeId) |> Option.defaultValue ("textbox" + string (abs (hash (width, height, x, y))))
    let paragraphs =
        if isNullOrUndefined item?Paragraphs then [||] else item?Paragraphs |> unbox<obj array>
    let content =
        if paragraphs.Length = 0 then generateParagraphWithHyperlinksAndImagesWithPageSetup packageSetup hyperlinks images (createObj [ "Runs" ==> [||]; "Alignment" ==> "left" ])
        else paragraphs |> Array.map (generateParagraphWithHyperlinksAndImagesWithPageSetup packageSetup hyperlinks images) |> String.concat ""
    let style = sprintf "position:absolute;margin-left:%.2fpt;margin-top:%.2fpt;width:%.2fpt;height:%.2fpt;z-index:1" x y width height
    element "w:p" [] [
        element "w:r" [] [
            element "w:pict" [] [
                element "v:shape" [
                    ("xmlns:v", "urn:schemas-microsoft-com:vml")
                    ("id", shapeId)
                    ("type", "#_x0000_t202")
                    ("style", style)
                    ("fillcolor", fill)
                    ("strokecolor", stroke)
                    ("strokeweight", sprintf "%.2fpt" strokeWeight)
                ] [
                    element "v:textbox" [ ("style", "mso-fit-shape-to-text:t") ] [
                        element "w:txbxContent" [] [content]
                    ]
                ]
            ]
        ]
    ]

let private wrapContentControl (item: obj) (content: string) =
    match optionalString (item?ContentControlId), optionalString (item?ContentControlTag), optionalString (item?ContentControlAlias) with
    | None, None, None -> content
    | id, tag, alias ->
        let props = [
            match id with
            | Some value when value <> "" -> yield selfClosing "w:id" [("w:val", value)]
            | _ -> ()
            match tag with
            | Some value when value <> "" -> yield selfClosing "w:tag" [("w:val", value)]
            | _ -> ()
            match alias with
            | Some value when value <> "" -> yield selfClosing "w:alias" [("w:val", value)]
            | _ -> ()
        ]
        element "w:sdt" [] [element "w:sdtPr" [] props; element "w:sdtContent" [] [content]]

let private generateDocumentItemWithImages (hyperlinks: Map<string, string>) (images: ImageScope) (item: obj) : string =
    let content =
        match item?Type |> stringOrDefault "Paragraph" with
        | "Table" -> generateTable hyperlinks images item
        | "TextBox" -> generateTextBox defaultPageSetupInfo hyperlinks images item
        | _ -> generateParagraphWithHyperlinksAndImagesWithPageSetup defaultPageSetupInfo hyperlinks images item
    wrapContentControl item content

let private generateDocumentItemWithImagesAndPageSetup (packageSetup: PageSetupInfo) (hyperlinks: Map<string, string>) (images: ImageScope) (item: obj) : string =
    let content =
        match item?Type |> stringOrDefault "Paragraph" with
        | "Table" -> generateTable hyperlinks images item
        | "TextBox" -> generateTextBox packageSetup hyperlinks images item
        | _ -> generateParagraphWithHyperlinksAndImagesWithPageSetup packageSetup hyperlinks images item
    wrapContentControl item content

let private generateDocumentItem (hyperlinks: Map<string, string>) (item: obj) : string =
    generateDocumentItemWithImages hyperlinks emptyImageScope item

let generateParagraph (para: obj) : string =
    generateParagraphWithHyperlinks Map.empty para

type HeaderFooterPartInfo = { Kind: string; RelId: string; Target: string; PartName: string; Blocks: obj array; IsGlobal: bool }
type HeaderFooterPartPackage = { Info: HeaderFooterPartInfo; Images: ImageScope; Hyperlinks: Map<string, string> }

let private partBlocks (part: obj) =
    if not (isNullOrUndefined part?Blocks) then optionalArray (part?Blocks)
    elif not (isNullOrUndefined part?Paragraphs) then optionalArray (part?Paragraphs)
    else [||]

let private headerFooterInfos scopedRelationshipIds prefix startIndex (parts: obj array) =
    parts
    |> Array.mapi (fun index part ->
        let number = startIndex + index
        let explicitRelId = optionalString (part?RelationshipId)
        let relId = explicitRelId |> Option.defaultValue (sprintf "rId%s%d" prefix number)
        let fileName = sprintf "%s%d.xml" prefix number
        { Kind = part?Kind |> stringOrDefault (part?Type |> stringOrDefault "default")
          RelId = relId
          Target = fileName
          PartName = sprintf "/word/%s" fileName
          Blocks = partBlocks part
          IsGlobal = not (Set.contains relId scopedRelationshipIds) })

let private generateHeaderFooterXml rootName (hyperlinks: Map<string, string>) (images: ImageScope) (blocks: obj array) =
    let content = blocks |> Array.map (generateDocumentItemWithImages hyperlinks images) |> String.concat ""
    document (element rootName [("xmlns:w", Namespaces.word); ("xmlns:r", Namespaces.relationships); ("xmlns:m", Namespaces.math)] [content])

let private generateSectionProperties (pageSetup: PageSetupInfo option) (headers: HeaderFooterPartInfo array) (footers: HeaderFooterPartInfo array) =
    let setup = pageSetup |> Option.defaultValue defaultPageSetupInfo
    let refs name (items: HeaderFooterPartInfo array) =
        items
        |> Array.filter (fun item -> includeHeaderFooterKind setup.DifferentFirstPage setup.EvenAndOddHeaders item.Kind)
        |> Array.map (fun item -> selfClosing name [("w:type", item.Kind); ("r:id", item.RelId)])
        |> Array.toList
    let pageSetupChildren =
        match pageSetup with
        | None -> []
        | Some setup -> sectionPageSetupChildren setup
    let children = refs "w:headerReference" headers @ refs "w:footerReference" footers @ pageSetupChildren
    if children.IsEmpty then "" else element "w:sectPr" [] children

/// Generate document.xml with formatting (handles JS arrays)
let generateDocumentXmlWithFormatting (paragraphs: obj array) : string =
    let hyperlinks = hyperlinkMap paragraphs
    let paraElements =
        paragraphs
        |> Array.map (generateDocumentItem hyperlinks)
        |> String.concat ""

    let body = element "w:body" [] [paraElements]

    let doc =
        element "w:document"
            [("xmlns:w", Namespaces.word); ("xmlns:r", Namespaces.relationships); ("xmlns:m", Namespaces.math)]
            [body]

    document doc

let private generateDocumentXmlWithFormattingAndImages (paragraphs: obj array) (images: ImageScope) : string =
    let hyperlinks = hyperlinkMap paragraphs
    let paraElements =
        paragraphs
        |> Array.map (generateDocumentItemWithImages hyperlinks images)
        |> String.concat ""

    let body = element "w:body" [] [paraElements]

    let doc =
        element "w:document"
            [("xmlns:w", Namespaces.word); ("xmlns:r", Namespaces.relationships); ("xmlns:m", Namespaces.math)]
            [body]

    document doc

let generateDocumentXmlPackage (paragraphs: obj array) (images: ImageScope) (pageSetup: PageSetupInfo option) (headers: HeaderFooterPartInfo array) (footers: HeaderFooterPartInfo array) : string =
    let hyperlinks = hyperlinkMap paragraphs
    let setup = pageSetup |> Option.defaultValue defaultPageSetupInfo
    let paraElements =
        paragraphs
        |> Array.map (generateDocumentItemWithImagesAndPageSetup setup hyperlinks images)
        |> String.concat ""
    let globalHeaders = headers |> Array.filter (fun part -> part.IsGlobal)
    let globalFooters = footers |> Array.filter (fun part -> part.IsGlobal)
    let body = element "w:body" [] [paraElements; generateSectionProperties pageSetup globalHeaders globalFooters]
    let doc = element "w:document" [("xmlns:w", Namespaces.word); ("xmlns:r", Namespaces.relationships); ("xmlns:m", Namespaces.math)] [body]
    document doc

/// Generate [Content_Types].xml
let generateContentTypes (hasComments: bool) (hasFootnotes: bool) (hasEndnotes: bool) : string =
    let children = [
        yield selfClosing "Default" [("Extension", "rels"); ("ContentType", "application/vnd.openxmlformats-package.relationships+xml")]
        yield selfClosing "Default" [("Extension", "xml"); ("ContentType", "application/xml")]
        yield selfClosing "Override" [("PartName", "/word/document.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")]
        yield selfClosing "Override" [("PartName", "/word/styles.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml")]
        yield selfClosing "Override" [("PartName", "/word/numbering.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml")]
        if hasComments then
            yield selfClosing "Override" [("PartName", "/word/comments.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml")]
        if hasFootnotes then
            yield selfClosing "Override" [("PartName", "/word/footnotes.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml")]
        if hasEndnotes then
            yield selfClosing "Override" [("PartName", "/word/endnotes.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml")]
    ]
    let root = element "Types" [("xmlns", Namespaces.contentTypes)] children
    document root

let private generateContentTypesWithImages hasComments hasFootnotes hasEndnotes (imageAssets: ImageAsset array) =
    let children = [
        yield selfClosing "Default" [("Extension", "rels"); ("ContentType", "application/vnd.openxmlformats-package.relationships+xml")]
        yield selfClosing "Default" [("Extension", "xml"); ("ContentType", "application/xml")]
        for extension, contentType in imageAssets |> Array.filter (fun image -> not image.Linked) |> Array.map (fun image -> image.Extension, image.ContentType) |> Array.distinct do
            yield selfClosing "Default" [("Extension", extension); ("ContentType", contentType)]
        yield selfClosing "Override" [("PartName", "/word/document.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")]
        yield selfClosing "Override" [("PartName", "/word/styles.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml")]
        yield selfClosing "Override" [("PartName", "/word/numbering.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml")]
        if hasComments then yield selfClosing "Override" [("PartName", "/word/comments.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml")]
        if hasFootnotes then yield selfClosing "Override" [("PartName", "/word/footnotes.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml")]
        if hasEndnotes then yield selfClosing "Override" [("PartName", "/word/endnotes.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml")]
    ]
    document (element "Types" [("xmlns", Namespaces.contentTypes)] children)

let generateContentTypesPackage hasSettings hasComments hasFootnotes hasEndnotes (headers: HeaderFooterPartInfo array) (footers: HeaderFooterPartInfo array) (imageAssets: ImageAsset array) =
    let children = [
        yield selfClosing "Default" [("Extension", "rels"); ("ContentType", "application/vnd.openxmlformats-package.relationships+xml")]
        yield selfClosing "Default" [("Extension", "xml"); ("ContentType", "application/xml")]
        for extension, contentType in imageAssets |> Array.filter (fun image -> not image.Linked) |> Array.map (fun image -> image.Extension, image.ContentType) |> Array.distinct do
            yield selfClosing "Default" [("Extension", extension); ("ContentType", contentType)]
        yield selfClosing "Override" [("PartName", "/word/document.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")]
        yield selfClosing "Override" [("PartName", "/word/styles.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml")]
        yield selfClosing "Override" [("PartName", "/word/numbering.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml")]
        if hasSettings then yield selfClosing "Override" [("PartName", "/word/settings.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml")]
        for header in headers do
            yield selfClosing "Override" [("PartName", header.PartName); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")]
        for footer in footers do
            yield selfClosing "Override" [("PartName", footer.PartName); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")]
        if hasComments then yield selfClosing "Override" [("PartName", "/word/comments.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml")]
        if hasFootnotes then yield selfClosing "Override" [("PartName", "/word/footnotes.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml")]
        if hasEndnotes then yield selfClosing "Override" [("PartName", "/word/endnotes.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml")]
    ]
    document (element "Types" [("xmlns", Namespaces.contentTypes)] children)

/// Generate _rels/.rels
let generateRootRels () : string =
    let rel = selfClosing "Relationship" [("Id", "rId1"); ("Type", RelationshipTypes.officeDocument); ("Target", "word/document.xml")]
    let root = element "Relationships" [("xmlns", Namespaces.relationshipsPackage)] [rel]
    document root

/// Generate word/_rels/document.xml.rels
let private documentRelationshipElements (hyperlinks: Map<string, string>) (hasComments: bool) (hasFootnotes: bool) (hasEndnotes: bool) =
    let baseRels = [
        yield selfClosing "Relationship" [("Id", "rId1"); ("Type", RelationshipTypes.styles); ("Target", "styles.xml")]
        yield selfClosing "Relationship" [("Id", "rId2"); ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"); ("Target", "numbering.xml")]
        if hasComments then
            yield selfClosing "Relationship" [("Id", "rIdComments"); ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"); ("Target", "comments.xml")]
        if hasFootnotes then
            yield selfClosing "Relationship" [("Id", "rIdFootnotes"); ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"); ("Target", "footnotes.xml")]
        if hasEndnotes then
            yield selfClosing "Relationship" [("Id", "rIdEndnotes"); ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"); ("Target", "endnotes.xml")]
    ]
    let hyperlinkRels =
        hyperlinks
        |> Map.toList
        |> List.map (fun (url, relId) ->
            selfClosing "Relationship" [("Id", relId); ("Type", RelationshipTypes.hyperlink); ("Target", url); ("TargetMode", "External")])
    baseRels @ hyperlinkRels

let generateDocumentRels (hyperlinks: Map<string, string>) (hasComments: bool) (hasFootnotes: bool) (hasEndnotes: bool) : string =
    let rels = documentRelationshipElements hyperlinks hasComments hasFootnotes hasEndnotes
    let root = element "Relationships" [("xmlns", Namespaces.relationshipsPackage)] rels
    document root

let private generateImageRelationships (assets: ImageAsset array) =
    assets
    |> Array.map (fun asset ->
        if asset.Linked then
            selfClosing "Relationship" [("Id", asset.RelId); ("Type", RelationshipTypes.image); ("Target", asset.Src); ("TargetMode", "External")]
        else
            selfClosing "Relationship" [("Id", asset.RelId); ("Type", RelationshipTypes.image); ("Target", sprintf "media/%s" asset.FileName)])
    |> Array.toList

let private generateHyperlinkRelationships (hyperlinks: Map<string, string>) =
    hyperlinks
    |> Map.toList
    |> List.map (fun (url, relId) ->
        selfClosing "Relationship" [("Id", relId); ("Type", RelationshipTypes.hyperlink); ("Target", url); ("TargetMode", "External")])

let private generateDocumentRelsWithImages hyperlinks hasComments hasFootnotes hasEndnotes (images: ImageScope) =
    let rels = documentRelationshipElements hyperlinks hasComments hasFootnotes hasEndnotes @ generateImageRelationships images.Assets
    let root = element "Relationships" [("xmlns", Namespaces.relationshipsPackage)] rels
    document root

let generateDocumentRelsPackage hasSettings hyperlinks hasComments hasFootnotes hasEndnotes (headers: HeaderFooterPartInfo array) (footers: HeaderFooterPartInfo array) (images: ImageScope) =
    let headerFooterRels =
        [ if hasSettings then yield selfClosing "Relationship" [("Id", "rIdSettings"); ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"); ("Target", "settings.xml")]
          for header in headers -> selfClosing "Relationship" [("Id", header.RelId); ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"); ("Target", header.Target)]
          for footer in footers -> selfClosing "Relationship" [("Id", footer.RelId); ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"); ("Target", footer.Target)] ]
    let rels = documentRelationshipElements hyperlinks hasComments hasFootnotes hasEndnotes @ headerFooterRels @ generateImageRelationships images.Assets
    let root = element "Relationships" [("xmlns", Namespaces.relationshipsPackage)] rels
    document root

let private generatePartRels (hyperlinks: Map<string, string>) (images: ImageScope) =
    let rels = generateImageRelationships images.Assets @ generateHyperlinkRelationships hyperlinks
    document (element "Relationships" [("xmlns", Namespaces.relationshipsPackage)] rels)

let generateComments (comments: CommentRef list) : string =
    let children =
        comments
        |> List.map (fun comment ->
            let attrs =
                [("w:id", comment.Id); ("w:author", comment.Author |> Option.defaultValue "RedBat")]
                @ (comment.Date |> Option.map (fun date -> [("w:date", date)]) |> Option.defaultValue [])
            element "w:comment" attrs [
                element "w:p" [] [
                    element "w:r" [] [elementWithText "w:t" [("xml:space", "preserve")] comment.Text]
                ]
            ])

    document (element "w:comments" [("xmlns:w", Namespaces.word)] children)

let generateNotes (rootName: string) (noteElementName: string) (notes: NoteRef list) : string =
    let children =
        notes
        |> List.map (fun note ->
            element noteElementName [("w:id", note.Id)] [
                element "w:p" [] [
                    element "w:r" [] [elementWithText "w:t" [("xml:space", "preserve")] note.Text]
                ]
            ])

    document (element rootName [("xmlns:w", Namespaces.word)] children)

type NumberingInfo = { NumId: string; Level: int; Format: string; LevelText: string; Start: int; HasCustomDefinition: bool; HasOverride: bool }

let private collectNumberingInfo (items: obj array) =
    items
    |> Array.toList
    |> List.collect collectItemParagraphs
    |> List.choose (fun para ->
        let listStyle = para?ListStyle |> stringOrDefault "none"
        if listStyle = "none" then None
        else
            optionalIntString (para?ListId)
            |> Option.bind (fun numId ->
                let hasCustomDefinition = numId <> "1" && numId <> "2"
                let hasOverride = boolOrDefault false (para?ListRestart) || boolOrDefault false (para?RestartNumbering) || not (isNullOrUndefined para?ListStart)
                if not hasCustomDefinition && not hasOverride then None
                else
                    let level = max 0 ((para?ListLevel |> intOrDefault 1) - 1)
                    let format = optionalString (para?ListNumberFormat) |> Option.defaultValue (if listStyle = "bullet" then "bullet" else "decimal")
                    let levelText = optionalString (para?ListLevelText) |> Option.defaultValue (if format = "bullet" then "•" else "%" + string (level + 1) + ".")
                    let start = if isNullOrUndefined para?ListStart then 1 else unbox<int> para?ListStart
                    Some { NumId = numId; Level = level; Format = format; LevelText = levelText; Start = start; HasCustomDefinition = hasCustomDefinition; HasOverride = hasOverride }))
    |> List.distinctBy (fun info -> info.NumId, info.Level)

/// Generate minimal styles.xml
type StyleInfo = {
    StyleId: string
    Name: string option
    BasedOn: string option
    Next: string option
    Linked: string option
    IsQuickStyle: bool
    IsDefault: bool
    Aliases: string option
    UiPriority: int option
    SemiHidden: bool
    UnhideWhenUsed: bool
    StyleType: string
    Bold: bool option
    Italic: bool option
    Underline: bool option
    AllCaps: bool option
    SmallCaps: bool option
    Strike: bool option
    VertAlign: string option
    FontSize: float option
    CharacterSpacing: float option
    FontFamily: string option
    Color: string option
    Highlight: string option
    SpacingBefore: float option
    SpacingAfter: float option
    LineSpacing: float option
    Alignment: string option
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

let private emptyStyleInfo styleId styleType = {
    StyleId = styleId
    Name = None
    BasedOn = None
    Next = None
    Linked = None
    IsQuickStyle = false
    IsDefault = false
    Aliases = None
    UiPriority = None
    SemiHidden = false
    UnhideWhenUsed = false
    StyleType = styleType
    Bold = None
    Italic = None
    Underline = None
    AllCaps = None
    SmallCaps = None
    Strike = None
    VertAlign = None
    FontSize = None
    CharacterSpacing = None
    FontFamily = None
    Color = None
    Highlight = None
    SpacingBefore = None
    SpacingAfter = None
    LineSpacing = None
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
}

let private styleInfoFromObject defaultType (value: obj) =
    let styleId =
        optionalString (value?StyleId)
        |> Option.orElse (optionalString (value?Id))
        |> Option.orElse (optionalString (value?Style))
    styleId
    |> Option.bind (fun id ->
        if id = "" then None
        else
            Some {
                emptyStyleInfo id (optionalString (value?StyleType) |> Option.defaultValue defaultType) with
                    Name = optionalString (value?Name) |> Option.orElse (optionalString (value?StyleName))
                    BasedOn = optionalString (value?BasedOn) |> Option.orElse (optionalString (value?BasedOnStyle))
                    Next = optionalString (value?Next) |> Option.orElse (optionalString (value?NextStyle))
                    Linked = optionalString (value?Linked) |> Option.orElse (optionalString (value?LinkedStyle))
                    IsQuickStyle = optionalBool (value?IsQuickStyle) |> Option.defaultValue false
                    IsDefault = optionalBool (value?IsDefaultStyle) |> Option.orElse (optionalBool (value?IsDefault)) |> Option.defaultValue false
                    Aliases = optionalString (value?StyleAliases) |> Option.orElse (optionalString (value?Aliases))
                    UiPriority = if isNullOrUndefined value?StyleUiPriority then None else Some (unbox<int> value?StyleUiPriority)
                    SemiHidden = optionalBool (value?StyleSemiHidden) |> Option.defaultValue false
                    UnhideWhenUsed = optionalBool (value?StyleUnhideWhenUsed) |> Option.defaultValue false
                    Bold = optionalBool (value?StyleBold)
                    Italic = optionalBool (value?StyleItalic)
                    Underline = optionalBool (value?StyleUnderline)
                    AllCaps = optionalBool (value?StyleAllCaps)
                    SmallCaps = optionalBool (value?StyleSmallCaps)
                    Strike = optionalBool (value?StyleStrike)
                    VertAlign = optionalString (value?StyleVertAlign)
                    FontSize = optionalFloat (value?StyleFontSize)
                    CharacterSpacing = optionalFloat (value?StyleCharacterSpacing)
                    FontFamily = optionalString (value?StyleFontFamily)
                    Color = optionalString (value?StyleColor)
                    Highlight = optionalString (value?StyleHighlight)
                    SpacingBefore = optionalFloat (value?StyleSpacingBefore)
                    SpacingAfter = optionalFloat (value?StyleSpacingAfter)
                    LineSpacing = optionalFloat (value?StyleLineSpacing)
                    Alignment = optionalString (value?StyleAlignment)
                    IndentLeft = optionalFloat (value?StyleIndentLeft)
                    IndentRight = optionalFloat (value?StyleIndentRight)
                    FirstLineIndent = optionalFloat (value?StyleFirstLineIndent)
                    BackgroundColor = optionalString (value?StyleBackgroundColor)
                    OutlineLevel = if isNullOrUndefined value?StyleOutlineLevel then None else Some (unbox<int> value?StyleOutlineLevel)
                    KeepNext = optionalBool (value?StyleKeepNext)
                    KeepLines = optionalBool (value?StyleKeepLines)
                    PageBreakBefore = optionalBool (value?StylePageBreakBefore)
                    WidowControl = optionalBool (value?StyleWidowControl)
            })

let private styleIdFromParagraph (para: obj) =
    optionalString (para?Style)
    |> Option.bind (fun styleId ->
        if styleId = "" then None
        else Some {
            emptyStyleInfo styleId "paragraph" with
                Name = optionalString (para?StyleName)
                BasedOn = optionalString (para?BasedOnStyle)
                Next = optionalString (para?NextStyle)
                Linked = optionalString (para?LinkedStyle)
                IsQuickStyle = optionalBool (para?IsQuickStyle) |> Option.defaultValue false
                IsDefault = optionalBool (para?IsDefaultStyle) |> Option.orElse (optionalBool (para?IsDefault)) |> Option.defaultValue false
                Aliases = optionalString (para?StyleAliases) |> Option.orElse (optionalString (para?Aliases))
                UiPriority = if isNullOrUndefined para?StyleUiPriority then None else Some (unbox<int> para?StyleUiPriority)
                SemiHidden = optionalBool (para?StyleSemiHidden) |> Option.defaultValue false
                UnhideWhenUsed = optionalBool (para?StyleUnhideWhenUsed) |> Option.defaultValue false
                Bold = optionalBool (para?StyleBold)
                Italic = optionalBool (para?StyleItalic)
                Underline = optionalBool (para?StyleUnderline)
                AllCaps = optionalBool (para?StyleAllCaps)
                SmallCaps = optionalBool (para?StyleSmallCaps)
                Strike = optionalBool (para?StyleStrike)
                VertAlign = optionalString (para?StyleVertAlign)
                FontSize = optionalFloat (para?StyleFontSize)
                CharacterSpacing = optionalFloat (para?StyleCharacterSpacing)
                FontFamily = optionalString (para?StyleFontFamily)
                Color = optionalString (para?StyleColor)
                Highlight = optionalString (para?StyleHighlight)
                SpacingBefore = optionalFloat (para?StyleSpacingBefore)
                SpacingAfter = optionalFloat (para?StyleSpacingAfter)
                LineSpacing = optionalFloat (para?StyleLineSpacing)
                Alignment = optionalString (para?StyleAlignment)
                IndentLeft = optionalFloat (para?StyleIndentLeft)
                IndentRight = optionalFloat (para?StyleIndentRight)
                FirstLineIndent = optionalFloat (para?StyleFirstLineIndent)
                BackgroundColor = optionalString (para?StyleBackgroundColor)
                OutlineLevel = if isNullOrUndefined para?StyleOutlineLevel then None else Some (unbox<int> para?StyleOutlineLevel)
                KeepNext = optionalBool (para?StyleKeepNext)
                KeepLines = optionalBool (para?StyleKeepLines)
                PageBreakBefore = optionalBool (para?StylePageBreakBefore)
                WidowControl = optionalBool (para?StyleWidowControl)
        })

let private styleIdFromRun (run: obj) =
    optionalString (run?Style)
    |> Option.orElse (optionalString (run?RunStyle))
    |> Option.orElse (optionalString (run?CharacterStyle))
    |> Option.bind (fun styleId ->
        if styleId = "" then None
        else Some {
            emptyStyleInfo styleId "character" with
                Name = optionalString (run?StyleName) |> Option.orElse (optionalString (run?RunStyleName)) |> Option.orElse (optionalString (run?CharacterStyleName))
                BasedOn = optionalString (run?BasedOnStyle)
                Next = optionalString (run?NextStyle)
                Linked = optionalString (run?LinkedStyle)
                IsQuickStyle = optionalBool (run?IsQuickStyle) |> Option.defaultValue false
                IsDefault = optionalBool (run?IsDefaultStyle) |> Option.orElse (optionalBool (run?IsDefault)) |> Option.defaultValue false
                Aliases = optionalString (run?StyleAliases) |> Option.orElse (optionalString (run?Aliases))
                UiPriority = if isNullOrUndefined run?StyleUiPriority then None else Some (unbox<int> run?StyleUiPriority)
                SemiHidden = optionalBool (run?StyleSemiHidden) |> Option.defaultValue false
                UnhideWhenUsed = optionalBool (run?StyleUnhideWhenUsed) |> Option.defaultValue false
                Bold = optionalBool (run?StyleBold)
                Italic = optionalBool (run?StyleItalic)
                Underline = optionalBool (run?StyleUnderline)
                AllCaps = optionalBool (run?StyleAllCaps)
                SmallCaps = optionalBool (run?StyleSmallCaps)
                Strike = optionalBool (run?StyleStrike)
                VertAlign = optionalString (run?StyleVertAlign)
                FontSize = optionalFloat (run?StyleFontSize)
                CharacterSpacing = optionalFloat (run?StyleCharacterSpacing)
                FontFamily = optionalString (run?StyleFontFamily)
                Color = optionalString (run?StyleColor)
                Highlight = optionalString (run?StyleHighlight)
        })

let private runStylesFromParagraph (para: obj) =
    if isNullOrUndefined para?Runs then []
    else para?Runs |> unbox<obj array> |> Array.toList |> List.choose styleIdFromRun

let rec private collectItemParagraphStyles (item: obj) =
    match item?Type |> stringOrDefault "Paragraph" with
    | "TextBox" when not (isNullOrUndefined item?Paragraphs) ->
        item?Paragraphs |> unbox<obj array> |> Array.toList |> List.collect collectItemParagraphStyles
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
                    if isNullOrUndefined cell?Paragraphs then []
                    else cell?Paragraphs |> unbox<obj array> |> Array.toList |> List.collect collectItemParagraphStyles))
    | _ -> (styleIdFromParagraph item |> Option.toList) @ runStylesFromParagraph item

let private collectStyles (items: obj array) =
    items
    |> Array.toList
    |> List.collect collectItemParagraphStyles
    |> List.distinctBy (fun style -> style.StyleType, style.StyleId)

let private collectPackageStyles (styles: obj array) =
    styles
    |> Array.toList
    |> List.choose (styleInfoFromObject "paragraph")

let private styleRunProps style =
    [
        match style.Bold with | Some true -> yield selfClosing "w:b" [] | _ -> ()
        match style.Italic with | Some true -> yield selfClosing "w:i" [] | _ -> ()
        match style.Underline with | Some true -> yield selfClosing "w:u" [("w:val", "single")] | _ -> ()
        match style.AllCaps with | Some true -> yield selfClosing "w:caps" [] | _ -> ()
        match style.SmallCaps with | Some true -> yield selfClosing "w:smallCaps" [] | _ -> ()
        match style.Strike with | Some true -> yield selfClosing "w:strike" [] | _ -> ()
        match style.VertAlign with | Some value when value = "superscript" || value = "subscript" -> yield selfClosing "w:vertAlign" [("w:val", value)] | _ -> ()
        match style.FontFamily with
        | Some family when family <> "" -> yield selfClosing "w:rFonts" [("w:ascii", family); ("w:hAnsi", family); ("w:cs", family)]
        | _ -> ()
        match style.FontSize with
        | Some size ->
            let value = InverseConversion.pointsToHalfPoints size
            yield selfClosing "w:sz" [("w:val", value)]
            yield selfClosing "w:szCs" [("w:val", value)]
        | _ -> ()
        match style.CharacterSpacing with
        | Some spacing when abs spacing >= 0.001 -> yield selfClosing "w:spacing" [("w:val", string (int (spacing * 15.0)))]
        | _ -> ()
        match InverseConversion.cssColorToWord style.Color with
        | Some color -> yield selfClosing "w:color" [("w:val", color)]
        | None -> ()
        match style.Highlight with
        | Some color when color <> "" -> yield selfClosing "w:highlight" [("w:val", color)]
        | _ -> ()
    ]

let private styleParagraphProps style =
    [
        let spacingAttrs =
            [
                match style.SpacingBefore with | Some value -> yield ("w:before", InverseConversion.pointsToTwips value) | None -> ()
                match style.SpacingAfter with | Some value -> yield ("w:after", InverseConversion.pointsToTwips value) | None -> ()
                match style.LineSpacing with | Some value -> yield ("w:line", InverseConversion.lineSpacingToWord value) | None -> ()
            ]
        if not spacingAttrs.IsEmpty then yield selfClosing "w:spacing" spacingAttrs
        match style.Alignment with
        | Some value when value <> "" -> yield selfClosing "w:jc" [("w:val", InverseConversion.alignmentToWord value)]
        | _ -> ()
        let indentAttrs =
            [
                match style.IndentLeft with | Some value -> yield ("w:left", InverseConversion.pointsToTwips value) | None -> ()
                match style.IndentRight with | Some value -> yield ("w:right", InverseConversion.pointsToTwips value) | None -> ()
                match style.FirstLineIndent with
                | Some value when value > 0.0 -> yield ("w:firstLine", InverseConversion.pointsToTwips value)
                | Some value when value < 0.0 -> yield ("w:hanging", InverseConversion.pointsToTwips -value)
                | _ -> ()
            ]
        if not indentAttrs.IsEmpty then yield selfClosing "w:ind" indentAttrs
        match InverseConversion.cssColorToWord style.BackgroundColor with
        | Some color -> yield selfClosing "w:shd" [("w:val", "clear"); ("w:fill", color)]
        | None -> ()
        match style.OutlineLevel with | Some value -> yield selfClosing "w:outlineLvl" [("w:val", string value)] | None -> ()
        match style.KeepNext with | Some true -> yield selfClosing "w:keepNext" [] | _ -> ()
        match style.KeepLines with | Some true -> yield selfClosing "w:keepLines" [] | _ -> ()
        match style.PageBreakBefore with | Some true -> yield selfClosing "w:pageBreakBefore" [] | _ -> ()
        match style.WidowControl with | Some true -> yield selfClosing "w:widowControl" [] | _ -> ()
    ]

let private defaultStyleId = "__docDefaults"

let private docDefaultsFromStyle (style: StyleInfo option) =
    let defaultRunStyle =
        style
        |> Option.defaultValue { emptyStyleInfo defaultStyleId "defaults" with FontFamily = Some "Calibri"; FontSize = Some 11.0 }
    let runProps = styleRunProps defaultRunStyle
    let rPr = element "w:rPr" [] runProps
    let rPrDefault = element "w:rPrDefault" [] [rPr]
    let pPr = style |> Option.map styleParagraphProps |> Option.defaultValue []
    let pPrDefault = if pPr.IsEmpty then element "w:pPrDefault" [] [] else element "w:pPrDefault" [] [element "w:pPr" [] pPr]
    element "w:docDefaults" [] [rPrDefault; pPrDefault]

let generateDefaultStylesWithStyles (styles: StyleInfo list) : string =
    let docDefaults = styles |> List.tryFind (fun style -> style.StyleType = "defaults") |> docDefaultsFromStyle
    let customStyles =
        styles
        |> List.filter (fun style -> style.StyleType <> "defaults")
        |> List.map (fun style ->
            let children = [
                yield selfClosing "w:name" [("w:val", style.Name |> Option.defaultValue style.StyleId)]
                match style.BasedOn with
                | Some basedOn when basedOn <> "" -> yield selfClosing "w:basedOn" [("w:val", basedOn)]
                | _ -> ()
                match style.Next with
                | Some next when next <> "" -> yield selfClosing "w:next" [("w:val", next)]
                | _ -> ()
                match style.Linked with
                | Some linked when linked <> "" -> yield selfClosing "w:link" [("w:val", linked)]
                | _ -> ()
                match style.Aliases with
                | Some aliases when aliases <> "" -> yield selfClosing "w:aliases" [("w:val", aliases)]
                | _ -> ()
                match style.UiPriority with
                | Some priority -> yield selfClosing "w:uiPriority" [("w:val", string priority)]
                | None -> ()
                if style.SemiHidden then yield selfClosing "w:semiHidden" []
                if style.UnhideWhenUsed then yield selfClosing "w:unhideWhenUsed" []
                if style.IsQuickStyle then yield selfClosing "w:qFormat" []
                let pPr = styleParagraphProps style
                if not pPr.IsEmpty then yield element "w:pPr" [] pPr
                let rPr = styleRunProps style
                if not rPr.IsEmpty then yield element "w:rPr" [] rPr
            ]
            let attrs = [("w:type", style.StyleType); ("w:styleId", style.StyleId)] @ (if style.IsDefault then [("w:default", "1")] else [])
            element "w:style" attrs children)
    let root = element "w:styles" [("xmlns:w", Namespaces.word); ("xmlns:r", Namespaces.relationships)] (docDefaults :: customStyles)
    document root

let generateDefaultStyles () : string =
    generateDefaultStylesWithStyles []

let generateNumbering (customInfos: NumberingInfo list) : string =
    let levels numFmt text =
        [0..8]
        |> List.map (fun level ->
            element "w:lvl" [("w:ilvl", string level)] [
                selfClosing "w:start" [("w:val", "1")]
                selfClosing "w:numFmt" [("w:val", numFmt)]
                selfClosing "w:lvlText" [("w:val", text)]
                element "w:pPr" [] [
                    selfClosing "w:ind" [("w:left", string (720 * (level + 1))); ("w:hanging", "360")]
                ]
            ])

    let bulletAbstract =
        element "w:abstractNum" [("w:abstractNumId", "1")] (levels "bullet" "•")
    let numberedAbstract =
        element "w:abstractNum" [("w:abstractNumId", "2")] (levels "decimal" "%1.")
    let hasNumberingInfo numId = customInfos |> List.exists (fun info -> info.NumId = numId)
    let defaultNums =
        [ if not (hasNumberingInfo "1") then yield element "w:num" [("w:numId", "1")] [selfClosing "w:abstractNumId" [("w:val", "1")]]
          if not (hasNumberingInfo "2") then yield element "w:num" [("w:numId", "2")] [selfClosing "w:abstractNumId" [("w:val", "2")]] ]

    let customAbstracts =
        customInfos
        |> List.filter (fun info -> info.HasCustomDefinition)
        |> List.groupBy (fun info -> info.NumId)
        |> List.map (fun (numId, infos) ->
            let levels =
                infos
                |> List.sortBy (fun info -> info.Level)
                |> List.map (fun info ->
                    element "w:lvl" [("w:ilvl", string info.Level)] [
                        selfClosing "w:start" [("w:val", string info.Start)]
                        selfClosing "w:numFmt" [("w:val", info.Format)]
                        selfClosing "w:lvlText" [("w:val", info.LevelText)]
                        element "w:pPr" [] [
                            selfClosing "w:ind" [("w:left", string (720 * (info.Level + 1))); ("w:hanging", "360")]
                        ]
                    ])
            element "w:abstractNum" [("w:abstractNumId", numId)] levels)

    let customNums =
        customInfos
        |> List.groupBy (fun info -> info.NumId)
        |> List.map (fun (numId, infos) ->
            let abstractNumId =
                if infos |> List.exists (fun info -> info.HasCustomDefinition) then numId
                elif numId = "1" then "1"
                else "2"
            let overrides =
                infos
                |> List.filter (fun info -> info.HasOverride)
                |> List.sortBy (fun info -> info.Level)
                |> List.map (fun info ->
                    element "w:lvlOverride" [("w:ilvl", string info.Level)] [
                        selfClosing "w:startOverride" [("w:val", string info.Start)]
                    ])
            element "w:num" [("w:numId", numId)] (selfClosing "w:abstractNumId" [("w:val", abstractNumId)] :: overrides))

    let root =
        element "w:numbering"
            [("xmlns:w", Namespaces.word)]
            ([bulletAbstract; numberedAbstract] @ defaultNums @ customAbstracts @ customNums)

    document root

let private generateSettings (pageSetup: PageSetupInfo option) =
    let children = [
        match pageSetup with
        | Some setup when setup.EvenAndOddHeaders -> yield selfClosing "w:evenAndOddHeaders" []
        | _ -> ()
    ]
    document (element "w:settings" [("xmlns:w", Namespaces.word)] children)

/// Write DOCX from EditorFormat (Phase 2)
let writeFromEditor (paragraphs: obj array) : Fable.Core.JS.Promise<obj> =
    let hyperlinks = hyperlinkMap paragraphs
    let bodyImages, _ = buildImageScope 1 paragraphs
    let comments = collectComments paragraphs
    let footnotes = collectNotes "footnote" paragraphs
    let endnotes = collectNotes "endnote" paragraphs
    let numbering = collectNumberingInfo paragraphs
    let styles = collectStyles paragraphs
    let hasComments = not comments.IsEmpty
    let hasFootnotes = not footnotes.IsEmpty
    let hasEndnotes = not endnotes.IsEmpty
    let files = [
        yield { Path = "[Content_Types].xml"; Content = generateContentTypesWithImages hasComments hasFootnotes hasEndnotes bodyImages.Assets }
        yield { Path = "_rels/.rels"; Content = generateRootRels() }
        yield { Path = "word/document.xml"; Content = generateDocumentXmlWithFormattingAndImages paragraphs bodyImages }
        yield { Path = "word/_rels/document.xml.rels"; Content = generateDocumentRelsWithImages hyperlinks hasComments hasFootnotes hasEndnotes bodyImages }
        yield { Path = "word/styles.xml"; Content = generateDefaultStylesWithStyles styles }
        yield { Path = "word/numbering.xml"; Content = generateNumbering numbering }
        if hasComments then
            yield { Path = "word/comments.xml"; Content = generateComments comments }
        if hasFootnotes then
            yield { Path = "word/footnotes.xml"; Content = generateNotes "w:footnotes" "w:footnote" footnotes }
        if hasEndnotes then
            yield { Path = "word/endnotes.xml"; Content = generateNotes "w:endnotes" "w:endnote" endnotes }
    ]

    let binaryFiles =
        bodyImages.Assets
        |> Array.choose (fun image -> image.Data |> Option.map (fun data -> sprintf "word/media/%s" image.FileName, data))
        |> Array.toList

    createPackageWithBinary files binaryFiles

let writeFromEditorPackage (package: obj) : Fable.Core.JS.Promise<obj> =
    let paragraphs = optionalArray (package?Paragraphs)
    let pageSetup = optionalPageSetup (package?PageSetup)
    let hasSettings = pageSetup |> Option.exists (fun setup -> setup.EvenAndOddHeaders)
    let sectionReferenceIds = collectSectionReferenceIds paragraphs
    let headerParts = optionalArray (package?HeaderParts) |> headerFooterInfos sectionReferenceIds "header" 1
    let footerParts = optionalArray (package?FooterParts) |> headerFooterInfos sectionReferenceIds "footer" 1
    let bodyImages, nextAfterBody = buildImageScope 1 paragraphs
    let mutable nextMediaIndex = nextAfterBody
    let headerPackages =
        headerParts
        |> Array.map (fun part ->
            let images, next = buildImageScope nextMediaIndex part.Blocks
            nextMediaIndex <- next
            { Info = part; Images = images; Hyperlinks = hyperlinkMap part.Blocks })
    let footerPackages =
        footerParts
        |> Array.map (fun part ->
            let images, next = buildImageScope nextMediaIndex part.Blocks
            nextMediaIndex <- next
            { Info = part; Images = images; Hyperlinks = hyperlinkMap part.Blocks })
    let allImages =
        [| yield! bodyImages.Assets
           for part in headerPackages do yield! part.Images.Assets
           for part in footerPackages do yield! part.Images.Assets |]
    let allBlocks =
        [| yield! paragraphs
           for part in headerParts do yield! part.Blocks
           for part in footerParts do yield! part.Blocks |]
    let hyperlinks = hyperlinkMap paragraphs
    let comments = collectComments allBlocks
    let footnotes = collectNotes "footnote" allBlocks
    let endnotes = collectNotes "endnote" allBlocks
    let numbering = collectNumberingInfo allBlocks
    let styles =
        (collectPackageStyles (optionalArray (package?Styles)) @ collectStyles allBlocks)
        |> List.distinctBy (fun style -> style.StyleType, style.StyleId)
    let hasComments = not comments.IsEmpty
    let hasFootnotes = not footnotes.IsEmpty
    let hasEndnotes = not endnotes.IsEmpty

    let files = [
        yield { Path = "[Content_Types].xml"; Content = generateContentTypesPackage hasSettings hasComments hasFootnotes hasEndnotes headerParts footerParts allImages }
        yield { Path = "_rels/.rels"; Content = generateRootRels() }
        yield { Path = "word/document.xml"; Content = generateDocumentXmlPackage paragraphs bodyImages pageSetup headerParts footerParts }
        yield { Path = "word/_rels/document.xml.rels"; Content = generateDocumentRelsPackage hasSettings hyperlinks hasComments hasFootnotes hasEndnotes headerParts footerParts bodyImages }
        yield { Path = "word/styles.xml"; Content = generateDefaultStylesWithStyles styles }
        yield { Path = "word/numbering.xml"; Content = generateNumbering numbering }
        if hasSettings then
            yield { Path = "word/settings.xml"; Content = generateSettings pageSetup }
        for header in headerPackages do
            yield { Path = "word/" + header.Info.Target; Content = generateHeaderFooterXml "w:hdr" header.Hyperlinks header.Images header.Info.Blocks }
            if header.Images.Assets.Length > 0 || not header.Hyperlinks.IsEmpty then
                yield { Path = sprintf "word/_rels/%s.rels" header.Info.Target; Content = generatePartRels header.Hyperlinks header.Images }
        for footer in footerPackages do
            yield { Path = "word/" + footer.Info.Target; Content = generateHeaderFooterXml "w:ftr" footer.Hyperlinks footer.Images footer.Info.Blocks }
            if footer.Images.Assets.Length > 0 || not footer.Hyperlinks.IsEmpty then
                yield { Path = sprintf "word/_rels/%s.rels" footer.Info.Target; Content = generatePartRels footer.Hyperlinks footer.Images }
        if hasComments then
            yield { Path = "word/comments.xml"; Content = generateComments comments }
        if hasFootnotes then
            yield { Path = "word/footnotes.xml"; Content = generateNotes "w:footnotes" "w:footnote" footnotes }
        if hasEndnotes then
            yield { Path = "word/endnotes.xml"; Content = generateNotes "w:endnotes" "w:endnote" endnotes }
    ]

    let binaryFiles =
        allImages
        |> Array.choose (fun image -> image.Data |> Option.map (fun data -> sprintf "word/media/%s" image.FileName, data))
        |> Array.toList

    createPackageWithBinary files binaryFiles
