module Composit.Ui.WordEditor

open System
open System.Collections.Generic
open System.IO
open Renderer
open Composit.Ui.Reactive
open Composit.Ui.Widget

open RedBat.Word.Document

module WordLayout = RedBat.Word.Layout
module WordPaint = RedBat.Word.Paint
module WordTextBuffer = RedBat.Word.TextBuffer

let inline private px (v: float) : float<px> = LanguagePrimitives.FloatWithMeasure v

let private resizeArrayToList (items: ResizeArray<'T>) =
    let mutable result = []
    for i = items.Count - 1 downto 0 do
        result <- items.[i] :: result
    result

let private prependResizeArray (items: ResizeArray<'T>) tail =
    let mutable result = tail
    for i = items.Count - 1 downto 0 do
        result <- items.[i] :: result
    result

let private rendererFont () = Composit.Ui.Layout.font()

let private pointsToPixels points = points * 96.0 / 72.0

let private wordFontCache = Dictionary<string, Renderer.Text.FontData option>()

let private existingDir path =
    if Directory.Exists path then Some path else None

let private wordFontDirs () =
    let dirs = ResizeArray<string>()
    for path in [
        Environment.GetEnvironmentVariable("RB_WORD_FONT_DIR")
        Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "artifacts/fonts/windows"))
        Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "../../artifacts/fonts/windows"))
        Path.GetFullPath(Path.Combine(__SOURCE_DIRECTORY__, "../artifacts/fonts/windows"))
    ] do
        if not (String.IsNullOrWhiteSpace path) then
            match existingDir path with
            | Some dir -> dirs.Add dir
            | None -> ()
    resizeArrayToList dirs

let private wordFontFamilyKey (family: string option) =
    let mutable start = 0
    let mutable finish = -1
    match family with
    | Some value ->
        finish <- value.Length - 1
        while start <= finish && Char.IsWhiteSpace value.[start] do
            start <- start + 1
        while finish >= start && Char.IsWhiteSpace value.[finish] do
            finish <- finish - 1
    | None -> ()

    let familyEquals expected =
        match family with
        | None -> false
        | Some value ->
            let length = finish - start + 1
            if length <> String.length expected then
                false
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

    if familyEquals "cambria" then Some "cambria"
    elif familyEquals "calibri" then Some "calibri"
    elif familyEquals "arial" then Some "arial"
    elif familyEquals "times new roman" then Some "times new roman"
    else None

let private fontFilesForKey familyKey (bold: bool) (italic: bool) =
    if familyKey = "cambria" then
        if bold then ["cambriab.ttf"; "cambria.ttc"]
        elif italic then ["cambriai.ttf"; "cambria.ttc"]
        else ["cambria.ttc"]
    elif familyKey = "calibri" then
        if bold && italic then ["calibriz.ttf"; "calibrib.ttf"; "calibri.ttf"]
        elif bold then ["calibrib.ttf"; "calibri.ttf"]
        elif italic then ["calibrii.ttf"; "calibri.ttf"]
        else ["calibri.ttf"]
    elif familyKey = "arial" then
        if bold && italic then ["arialbi.ttf"; "arialbd.ttf"; "arial.ttf"]
        elif bold then ["arialbd.ttf"; "arial.ttf"]
        elif italic then ["ariali.ttf"; "arial.ttf"]
        else ["arial.ttf"]
    elif familyKey = "times new roman" then
        if bold && italic then ["timesbi.ttf"; "timesbd.ttf"; "times.ttf"]
        elif bold then ["timesbd.ttf"; "times.ttf"]
        elif italic then ["timesi.ttf"; "times.ttf"]
        else ["times.ttf"]
    else []

let private tryLoadWordFont (family: string option) (bold: bool) (italic: bool) =
    match wordFontFamilyKey family with
    | None -> None
    | Some familyKey ->
        let key = sprintf "%s|%b|%b" familyKey bold italic
        match wordFontCache.TryGetValue key with
        | true, cached -> cached
        | _ ->
            let loaded =
                fontFilesForKey familyKey bold italic
                |> List.tryPick (fun fileName ->
                    wordFontDirs ()
                    |> List.tryPick (fun dir ->
                        let path = Path.Combine(dir, fileName)
                        if File.Exists path then
                            try Some (Renderer.Text.loadFontFile path)
                            with _ -> None
                        else None))
            wordFontCache.[key] <- loaded
            loaded

let private fontForTextStyle (style: RedBat.Word.TextBuffer.TextStyle) =
    tryLoadWordFont style.FontFamily style.Bold style.Italic
    |> Option.defaultWith rendererFont

let private fontForPaintStyle (style: WordPaint.TextPaintStyle) =
    tryLoadWordFont style.FontFamily style.Bold style.Italic
    |> Option.defaultWith rendererFont

[<Struct>]
type private WordMeasureCacheKey = {
    Text: string
    FontSize: float
    FontFamilyKey: string
    Bold: bool
    Italic: bool
    CharacterSpacing: float
}

let private wordMeasureCache = Dictionary<WordMeasureCacheKey, float>()
let private wordMeasureCacheOrder = Queue<WordMeasureCacheKey>()
let private wordMeasureCacheMaxEntries = 200000
let private wordMeasureCacheMaxTotalTextChars = 5000000
let private wordMeasureCacheMaxSingleTextChars = 32768
let mutable private wordMeasureCacheTotalTextChars = 0

let private wordMeasureCacheKeyChars key =
    key.Text.Length + key.FontFamilyKey.Length

let private trimWordMeasureCache () =
    while (wordMeasureCache.Count > wordMeasureCacheMaxEntries
           || wordMeasureCacheTotalTextChars > wordMeasureCacheMaxTotalTextChars)
          && wordMeasureCacheOrder.Count > 0 do
        let oldest = wordMeasureCacheOrder.Dequeue()
        if wordMeasureCache.Remove oldest then
            wordMeasureCacheTotalTextChars <- max 0 (wordMeasureCacheTotalTextChars - wordMeasureCacheKeyChars oldest)

let private rememberWordMeasure key width =
    let entryChars = wordMeasureCacheKeyChars key
    if key.Text.Length <= wordMeasureCacheMaxSingleTextChars
       && entryChars <= wordMeasureCacheMaxTotalTextChars then
        if not (wordMeasureCache.ContainsKey key) then
            wordMeasureCacheOrder.Enqueue key
            wordMeasureCacheTotalTextChars <- wordMeasureCacheTotalTextChars + entryChars
        wordMeasureCache.[key] <- width
        trimWordMeasureCache ()

let private wordMeasureCacheKey text (style: RedBat.Word.TextBuffer.TextStyle) = {
    Text = text
    FontSize = style.FontSize
    FontFamilyKey = wordFontFamilyKey style.FontFamily |> Option.defaultValue ""
    Bold = style.Bold
    Italic = style.Italic
    CharacterSpacing = style.CharacterSpacing
}

let private measureTextUncached (text: string) (style: RedBat.Word.TextBuffer.TextStyle) =
    let font = fontForTextStyle style
    Renderer.Text.measureWidth font (pointsToPixels style.FontSize) text + RedBat.Word.TextBuffer.characterSpacingAdvance text style

let private measureText (text: string) (style: RedBat.Word.TextBuffer.TextStyle) =
    let key = wordMeasureCacheKey text style
    match wordMeasureCache.TryGetValue key with
    | true, width -> width
    | false, _ ->
        let width = measureTextUncached text style
        rememberWordMeasure key width
        width

[<Struct>]
type private WordGlyphLayoutCacheKey = {
    Text: string
    FontSize: float
    FontFamilyKey: string
    Bold: bool
    Italic: bool
    CharacterSpacing: float
}

let private wordGlyphLayoutCache = Dictionary<WordGlyphLayoutCacheKey, Renderer.Text.GlyphRun list>()
let private wordGlyphLayoutCacheOrder = Queue<WordGlyphLayoutCacheKey>()
let private wordGlyphLayoutCacheMaxEntries = 200000
let private wordGlyphLayoutCacheMaxTotalTextChars = 5000000
let private wordGlyphLayoutCacheMaxSingleTextChars = 32768
let mutable private wordGlyphLayoutCacheTotalTextChars = 0

let private wordGlyphLayoutCacheKeyChars key =
    key.Text.Length + key.FontFamilyKey.Length

let private trimWordGlyphLayoutCache () =
    while (wordGlyphLayoutCache.Count > wordGlyphLayoutCacheMaxEntries
           || wordGlyphLayoutCacheTotalTextChars > wordGlyphLayoutCacheMaxTotalTextChars)
          && wordGlyphLayoutCacheOrder.Count > 0 do
        let oldest = wordGlyphLayoutCacheOrder.Dequeue()
        if wordGlyphLayoutCache.Remove oldest then
            wordGlyphLayoutCacheTotalTextChars <- max 0 (wordGlyphLayoutCacheTotalTextChars - wordGlyphLayoutCacheKeyChars oldest)

let private rememberWordGlyphLayout key layout =
    let entryChars = wordGlyphLayoutCacheKeyChars key
    if key.Text.Length <= wordGlyphLayoutCacheMaxSingleTextChars
       && entryChars <= wordGlyphLayoutCacheMaxTotalTextChars then
        if not (wordGlyphLayoutCache.ContainsKey key) then
            wordGlyphLayoutCacheOrder.Enqueue key
            wordGlyphLayoutCacheTotalTextChars <- wordGlyphLayoutCacheTotalTextChars + entryChars
        wordGlyphLayoutCache.[key] <- layout
        trimWordGlyphLayoutCache ()

let private wordGlyphLayoutCacheKey text (style: WordPaint.TextPaintStyle) = {
    Text = text
    FontSize = style.FontSize
    FontFamilyKey = wordFontFamilyKey style.FontFamily |> Option.defaultValue ""
    Bold = style.Bold
    Italic = style.Italic
    CharacterSpacing = style.CharacterSpacing
}

let private glyphLayoutForPaintText text style =
    let key = wordGlyphLayoutCacheKey text style
    match wordGlyphLayoutCache.TryGetValue key with
    | true, layout -> layout
    | false, _ ->
        let font = fontForPaintStyle style
        let layout = Renderer.Text.layout font (pointsToPixels style.FontSize) text
        rememberWordGlyphLayout key layout
        layout

let private parseHexColor (value: string) =
    let mutable start = 0
    let mutable finish = value.Length - 1
    while start <= finish && Char.IsWhiteSpace value.[start] do
        start <- start + 1
    while finish >= start && Char.IsWhiteSpace value.[finish] do
        finish <- finish - 1

    let spanLength () = finish - start + 1

    let equalsIgnoreCase expected =
        let length = spanLength ()
        if length <> String.length expected then
            false
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

    let hexValue ch =
        if ch >= '0' && ch <= '9' then Some (int ch - int '0')
        elif ch >= 'a' && ch <= 'f' then Some (10 + int ch - int 'a')
        elif ch >= 'A' && ch <= 'F' then Some (10 + int ch - int 'A')
        else None

    let hexByte index =
        match hexValue value.[index], hexValue value.[index + 1] with
        | Some hi, Some lo -> Some (byte (hi * 16 + lo))
        | _ -> None

    let doubledHexByte index =
        match hexValue value.[index] with
        | Some v -> Some (byte (v * 16 + v))
        | None -> None

    let parseDecimalByte partStart partEnd =
        let mutable i = partStart
        let mutable j = partEnd
        while i <= j && Char.IsWhiteSpace value.[i] do
            i <- i + 1
        while j >= i && Char.IsWhiteSpace value.[j] do
            j <- j - 1
        let mutable n = 0
        let mutable valid = i <= j
        while valid && i <= j do
            let ch = value.[i]
            if ch >= '0' && ch <= '9' then
                n <- n * 10 + int ch - int '0'
                if n > 255 then valid <- false
                i <- i + 1
            else
                valid <- false
        if valid then Some (byte n) else None

    let parseAlpha partStart partEnd =
        let mutable i = partStart
        let mutable j = partEnd
        while i <= j && Char.IsWhiteSpace value.[i] do
            i <- i + 1
        while j >= i && Char.IsWhiteSpace value.[j] do
            j <- j - 1
        let mutable whole = 0.0
        let mutable fraction = 0.0
        let mutable scale = 1.0
        let mutable afterDot = false
        let mutable sawDigit = false
        let mutable valid = i <= j
        while valid && i <= j do
            let ch = value.[i]
            if ch >= '0' && ch <= '9' then
                sawDigit <- true
                let digit = float (int ch - int '0')
                if afterDot then
                    scale <- scale * 10.0
                    fraction <- fraction + digit / scale
                else
                    whole <- whole * 10.0 + digit
                i <- i + 1
            elif ch = '.' && not afterDot then
                afterDot <- true
                i <- i + 1
            else
                valid <- false
        if valid && sawDigit then Some (whole + fraction) else None

    let parseRgba () =
        let prefix = "rgba("
        if spanLength () < prefix.Length + 1 || finish < start || value.[finish] <> ')' then
            None
        else
            let mutable index = 0
            let mutable matchesPrefix = true
            while matchesPrefix && index < prefix.Length do
                let left = value.[start + index]
                let right = prefix.[index]
                if left <> right && Char.ToUpperInvariant left <> Char.ToUpperInvariant right then
                    matchesPrefix <- false
                index <- index + 1
            if not matchesPrefix then
                None
            else
                let partStart = start + prefix.Length
                let partFinish = finish - 1
                let mutable partCount = 0
                let mutable r0, r1 = 0, 0
                let mutable g0, g1 = 0, 0
                let mutable b0, b1 = 0, 0
                let mutable a0, a1 = 0, 0
                let capture spanStart spanEnd =
                    match partCount with
                    | 0 -> r0 <- spanStart; r1 <- spanEnd
                    | 1 -> g0 <- spanStart; g1 <- spanEnd
                    | 2 -> b0 <- spanStart; b1 <- spanEnd
                    | 3 -> a0 <- spanStart; a1 <- spanEnd
                    | _ -> ()
                    partCount <- partCount + 1
                let mutable segmentStart = partStart
                let mutable i = partStart
                while i <= partFinish do
                    if value.[i] = ',' then
                        capture segmentStart (i - 1)
                        segmentStart <- i + 1
                    i <- i + 1
                capture segmentStart partFinish
                if partCount <> 4 then
                    None
                else
                    match parseDecimalByte r0 r1, parseDecimalByte g0 g1, parseDecimalByte b0 b1, parseAlpha a0 a1 with
                    | Some r, Some g, Some b, Some a ->
                        Some (Color.rgba r g b (byte (max 0.0 (min 255.0 (a * 255.0)))))
                    | _ -> None

    if start > finish then Color.black
    elif equalsIgnoreCase "black" then Color.rgb 0uy 0uy 0uy
    elif equalsIgnoreCase "blue" then Color.rgb 0uy 0uy 255uy
    elif equalsIgnoreCase "cyan" then Color.rgb 0uy 255uy 255uy
    elif equalsIgnoreCase "green" then Color.rgb 0uy 255uy 0uy
    elif equalsIgnoreCase "magenta" then Color.rgb 255uy 0uy 255uy
    elif equalsIgnoreCase "red" then Color.rgb 255uy 0uy 0uy
    elif equalsIgnoreCase "yellow" then Color.rgb 255uy 255uy 0uy
    elif equalsIgnoreCase "white" then Color.rgb 255uy 255uy 255uy
    elif equalsIgnoreCase "darkblue" then Color.rgb 0uy 0uy 128uy
    elif equalsIgnoreCase "darkcyan" then Color.rgb 0uy 128uy 128uy
    elif equalsIgnoreCase "darkgreen" then Color.rgb 0uy 128uy 0uy
    elif equalsIgnoreCase "darkmagenta" then Color.rgb 128uy 0uy 128uy
    elif equalsIgnoreCase "darkred" then Color.rgb 128uy 0uy 0uy
    elif equalsIgnoreCase "darkyellow" then Color.rgb 128uy 128uy 0uy
    elif equalsIgnoreCase "darkgray" then Color.rgb 128uy 128uy 128uy
    elif equalsIgnoreCase "lightgray" then Color.rgb 192uy 192uy 192uy
    else
        let hexStart = if value.[start] = '#' then start + 1 else start
        let hexLength = finish - hexStart + 1
        match hexLength with
        | 3 ->
            match doubledHexByte hexStart, doubledHexByte (hexStart + 1), doubledHexByte (hexStart + 2) with
            | Some r, Some g, Some b -> Color.rgb r g b
            | _ -> Color.black
        | 6 ->
            match hexByte hexStart, hexByte (hexStart + 2), hexByte (hexStart + 4) with
            | Some r, Some g, Some b -> Color.rgb r g b
            | _ -> Color.black
        | 8 ->
            match hexByte hexStart, hexByte (hexStart + 2), hexByte (hexStart + 4), hexByte (hexStart + 6) with
            | Some r, Some g, Some b, Some a -> Color.rgba r g b a
            | _ -> Color.black
        | _ ->
            parseRgba () |> Option.defaultValue Color.black

let private textFont style = fontForPaintStyle style

type WordEditorState = {
    Document: Signal<Document>
    HoveredImage: Signal<(BlockId * int) option>
    ScrollY: Signal<float>
    DragAnchor: Signal<Position option>
    WordDragAnchor: Signal<(Position * Position) option>
    ParagraphDragAnchor: Signal<(Position * Position) option>
    ResizingImage: Signal<(BlockId * int * float * float * float * float * Document) option>
    DraggingImage: Signal<(BlockId * int * float * float * Document) option>
    LastClick: Signal<(float * float * float * int) option>
    LastCursorRect: Signal<Rect option>
    PendingStyle: Signal<WordTextBuffer.TextStyle option>
    UndoStack: Signal<Document list>
    RedoStack: Signal<Document list>
    LayoutCache: WordLayoutCache
    FocusId: int
}

and WordLayoutCache = {
    mutable Width: float
    mutable CachedDocument: Document option
    mutable CachedLayoutDocument: Document option
    mutable Settings: WordLayout.LayoutSettings option
    mutable Layouts: WordLayout.BlockLayout list option
    mutable LayoutArray: WordLayout.BlockLayout array option
    mutable LayoutById: Dictionary<BlockId, WordLayout.BlockLayout> option
    mutable ContentHeight: float option
    mutable RendererContentPaintCommands: WordPaint.PaintCommand list option
    mutable RendererContentCommands: Cmd list option
    mutable RendererSelectionLayouts: WordLayout.BlockLayout list option
    mutable RendererSelectionValue: Selection option
    mutable RendererSelectionCommands: WordPaint.PaintCommand list option
    mutable RendererFrameContentCommands: Cmd list option
    mutable RendererFrameSelectionCommands: WordPaint.PaintCommand list option
    mutable RendererFrameHoverCommands: WordPaint.PaintCommand list option
    mutable RendererFrameCommands: Cmd list option
    mutable ViewportSupportDocument: Document option
    mutable ViewportSupportsContentCommands: bool option
}

let private maxUndoDepth = 50
let private historyMaxTotalRetainedSourceTextChars = 20000000
let private historyMaxRetainedSourceTraversalItems = 200000
let private historySourceTextOverLimit = historyMaxTotalRetainedSourceTextChars + 1

let private historyOptionStringCharCount (value: string option) =
    match value with
    | Some value -> value.Length
    | None -> 0

let private historyAddCapped limit value total =
    if total > limit then
        total
    else
        let value = max 0 value
        let remaining = limit - total
        if value > remaining then limit + 1
        else total + value

let private historyFailClosedForTraversalCount seen total =
    if seen > historyMaxRetainedSourceTraversalItems then historySourceTextOverLimit
    else total

let private historySectionColumnsStringCharCount (columns: SectionColumns option) =
    match columns with
    | Some columns ->
        let mutable total = 0
        let mutable seen = 0
        let mutable widths = columns.Widths
        while total <= historyMaxTotalRetainedSourceTextChars
              && seen <= historyMaxRetainedSourceTraversalItems
              && not widths.IsEmpty do
            total <- historyAddCapped historyMaxTotalRetainedSourceTextChars 1 total
            seen <- seen + 1
            widths <- widths.Tail
        historyFailClosedForTraversalCount seen total
    | None -> 0

let private historyTextStyleStringCharCount (style: WordTextBuffer.TextStyle) =
    let mutable total = 0
    let add value =
        total <- historyAddCapped historyMaxTotalRetainedSourceTextChars (historyOptionStringCharCount value) total
    add style.Style
    add style.UnderlineStyle
    add style.UnderlineColor
    add style.EmphasisMark
    add style.FontFamily
    add style.Color
    add style.Highlight
    add style.Hyperlink
    add style.CommentId
    add style.CommentText
    add style.CommentAuthor
    add style.CommentDate
    add style.NoteId
    add style.NoteType
    add style.NoteText
    add style.BookmarkStartId
    add style.BookmarkName
    add style.BookmarkEndId
    add style.FieldCode
    add style.FieldResult
    add style.EquationXml
    add style.RevisionId
    add style.RevisionType
    add style.RevisionAuthor
    add style.RevisionDate
    add style.VertAlign
    total

let private historyImageStringCharCount (image: WordTextBuffer.ImageRun) =
    image.Src.Length
    + historyOptionStringCharCount image.Alt
    + historyOptionStringCharCount image.Name
    + historyOptionStringCharCount image.Title
    + historyOptionStringCharCount image.Wrap
    + historyOptionStringCharCount image.PositionHRelativeFrom
    + historyOptionStringCharCount image.PositionVRelativeFrom
    + historyOptionStringCharCount image.RelativeHeight
    + historyOptionStringCharCount image.CropLeft
    + historyOptionStringCharCount image.CropTop
    + historyOptionStringCharCount image.CropRight
    + historyOptionStringCharCount image.CropBottom

let private historyRunStringCharCount = function
    | WordTextBuffer.Text textRun -> historyTextStyleStringCharCount textRun.Style
    | WordTextBuffer.Image image -> historyImageStringCharCount image

let private historyParagraphMetadataStringCharCount (metadata: BlockMetadata) =
    let limit = historyMaxTotalRetainedSourceTextChars
    let mutable total = 0
    let add value =
        total <- historyAddCapped limit (historyOptionStringCharCount value) total
    add metadata.Style
    add metadata.ContentControlId
    add metadata.ContentControlTag
    add metadata.ContentControlAlias
    add metadata.ListId
    add metadata.ListNumberFormat
    add metadata.ListLevelText
    add metadata.RightTabLeader
    add metadata.LineSpacingRule
    add metadata.BackgroundColor
    add metadata.SectionBreak
    total <- historyAddCapped limit (historySectionColumnsStringCharCount metadata.SectionColumns) total
    let mutable seen = 0
    let mutable tabStops = metadata.TabStops
    while total <= limit
          && seen <= historyMaxRetainedSourceTraversalItems
          && not tabStops.IsEmpty do
        let tabStop = tabStops.Head
        total <- historyAddCapped limit tabStop.Alignment.Length total
        total <- historyAddCapped limit (historyOptionStringCharCount tabStop.Leader) total
        seen <- seen + 1
        tabStops <- tabStops.Tail
    let mutable headerReferences = metadata.HeaderReferences
    while total <= limit
          && seen <= historyMaxRetainedSourceTraversalItems
          && not headerReferences.IsEmpty do
        let kind, relationshipId = headerReferences.Head
        total <- historyAddCapped limit kind.Length total
        total <- historyAddCapped limit relationshipId.Length total
        seen <- seen + 1
        headerReferences <- headerReferences.Tail
    let mutable footerReferences = metadata.FooterReferences
    while total <= limit
          && seen <= historyMaxRetainedSourceTraversalItems
          && not footerReferences.IsEmpty do
        let kind, relationshipId = footerReferences.Head
        total <- historyAddCapped limit kind.Length total
        total <- historyAddCapped limit relationshipId.Length total
        seen <- seen + 1
        footerReferences <- footerReferences.Tail
    historyFailClosedForTraversalCount seen total

let private historyParagraphSourceTextCharCount (para: ParagraphBlock) =
    let limit = historyMaxTotalRetainedSourceTextChars
    let mutable total = para.Id.Length
    total <- historyAddCapped limit (WordTextBuffer.length para.Buffer) total
    let mutable seen = 0
    let mutable runs = para.Buffer.Runs
    while total <= limit
          && seen <= historyMaxRetainedSourceTraversalItems
          && not runs.IsEmpty do
        total <- historyAddCapped limit (historyRunStringCharCount runs.Head) total
        seen <- seen + 1
        runs <- runs.Tail
    total <- historyFailClosedForTraversalCount seen total
    if total <= limit then
        total <- historyAddCapped limit (historyParagraphMetadataStringCharCount para.Metadata) total
    total

let private historyTableBorderStringCharCount (border: TableBorder option) =
    match border with
    | Some border ->
        let limit = historyMaxTotalRetainedSourceTextChars
        let mutable total = 0
        total <- historyAddCapped limit (historyOptionStringCharCount border.Style) total
        total <- historyAddCapped limit (historyOptionStringCharCount border.Color) total
        total
    | None -> 0

let private historyTableBordersStringCharCount (borders: TableBorders option) =
    match borders with
    | Some borders ->
        let limit = historyMaxTotalRetainedSourceTextChars
        let mutable total = 0
        total <- historyAddCapped limit (historyTableBorderStringCharCount borders.Top) total
        total <- historyAddCapped limit (historyTableBorderStringCharCount borders.Right) total
        total <- historyAddCapped limit (historyTableBorderStringCharCount borders.Bottom) total
        total <- historyAddCapped limit (historyTableBorderStringCharCount borders.Left) total
        total <- historyAddCapped limit (historyTableBorderStringCharCount borders.InsideH) total
        total <- historyAddCapped limit (historyTableBorderStringCharCount borders.InsideV) total
        total
    | None -> 0

let private historyShadingStringCharCount (shading: Shading option) =
    match shading with
    | Some shading ->
        let limit = historyMaxTotalRetainedSourceTextChars
        let mutable total = 0
        total <- historyAddCapped limit (historyOptionStringCharCount shading.Pattern) total
        total <- historyAddCapped limit (historyOptionStringCharCount shading.ForegroundColor) total
        total <- historyAddCapped limit (historyOptionStringCharCount shading.BackgroundColor) total
        total
    | None -> 0

let private historyTableStringCharCount (table: TableBlock) =
    let limit = historyMaxTotalRetainedSourceTextChars
    let mutable total = table.Id.Length
    total <- historyAddCapped limit (historyOptionStringCharCount table.Alignment) total
    total <- historyAddCapped limit (historyOptionStringCharCount table.BackgroundColor) total
    total <- historyAddCapped limit (historyShadingStringCharCount table.Shading) total
    total <- historyAddCapped limit (historyTableBordersStringCharCount table.Borders) total
    total

let private historyTableRowStringCharCount (row: TableRow) =
    let limit = historyMaxTotalRetainedSourceTextChars
    let mutable total = 0
    total <- historyAddCapped limit (historyOptionStringCharCount row.HeightRule) total
    total <- historyAddCapped limit (historyTableBordersStringCharCount row.BordersOverride) total
    total

let private historyTableCellStringCharCount (cell: TableCell) =
    let limit = historyMaxTotalRetainedSourceTextChars
    let mutable total = 0
    total <- historyAddCapped limit (historyOptionStringCharCount cell.BackgroundColor) total
    total <- historyAddCapped limit (historyShadingStringCharCount cell.Shading) total
    total <- historyAddCapped limit (historyOptionStringCharCount cell.BorderColor) total
    total <- historyAddCapped limit (historyTableBordersStringCharCount cell.Borders) total
    total <- historyAddCapped limit (historyOptionStringCharCount cell.VerticalAlignment) total
    total

let private historyTableSourceTextCharCount (table: TableBlock) =
    let limit = historyMaxTotalRetainedSourceTextChars
    let mutable total = 0
    let mutable visitedItems = 0
    let canContinue () =
        total <= limit && visitedItems <= historyMaxRetainedSourceTraversalItems
    let visitItem () =
        visitedItems <- visitedItems + 1
    let tables = System.Collections.Generic.Stack<TableBlock>()
    tables.Push table
    while tables.Count > 0 && canContinue () do
        let table = tables.Pop()
        visitItem ()
        total <- historyAddCapped limit (historyTableStringCharCount table) total
        let mutable rows = table.Rows
        while canContinue () && not rows.IsEmpty do
            let row = rows.Head
            visitItem ()
            total <- historyAddCapped limit (historyTableRowStringCharCount row) total
            let mutable cells = row.Cells
            while canContinue () && not cells.IsEmpty do
                let cell = cells.Head
                visitItem ()
                total <- historyAddCapped limit (historyTableCellStringCharCount cell) total
                let mutable paragraphs = cell.Content
                while canContinue () && not paragraphs.IsEmpty do
                    total <- historyAddCapped limit (historyParagraphSourceTextCharCount paragraphs.Head) total
                    paragraphs <- paragraphs.Tail
                let mutable nestedTables = cell.NestedTables
                while canContinue () && not nestedTables.IsEmpty do
                    tables.Push nestedTables.Head
                    nestedTables <- nestedTables.Tail
                cells <- cells.Tail
            rows <- rows.Tail
    historyFailClosedForTraversalCount visitedItems total

let private historyPathSegmentsStringCharCount (segments: PathSeg list) =
    let limit = historyMaxTotalRetainedSourceTextChars
    let mutable total = 0
    let mutable seen = 0
    let mutable remaining = segments
    while total <= limit
          && seen <= historyMaxRetainedSourceTraversalItems
          && not remaining.IsEmpty do
        total <- historyAddCapped limit 1 total
        seen <- seen + 1
        remaining <- remaining.Tail
    historyFailClosedForTraversalCount seen total

let private historyShapeGeometryStringCharCount = function
    | GeoRect -> 0
    | GeoPreset preset -> preset.Length
    | GeoCustom (segments, _, _) -> historyPathSegmentsStringCharCount segments

let private historyGradientStringCharCount = function
    | Some gradient ->
        let limit = historyMaxTotalRetainedSourceTextChars
        let mutable total = 0
        let mutable seen = 0
        let mutable stops = gradient.Stops
        while total <= limit
              && seen <= historyMaxRetainedSourceTraversalItems
              && not stops.IsEmpty do
            total <- historyAddCapped limit stops.Head.Color.Length total
            seen <- seen + 1
            stops <- stops.Tail
        historyFailClosedForTraversalCount seen total
    | None -> 0

let private historyTextBoxSourceTextCharCount (textBox: TextBoxBlock) =
    let limit = historyMaxTotalRetainedSourceTextChars
    let mutable total = textBox.Id.Length
    total <- historyAddCapped limit (historyOptionStringCharCount textBox.PositionHRelativeFrom) total
    total <- historyAddCapped limit (historyOptionStringCharCount textBox.PositionVRelativeFrom) total
    total <- historyAddCapped limit (historyOptionStringCharCount textBox.FillColor) total
    total <- historyAddCapped limit (historyOptionStringCharCount textBox.BorderColor) total
    total <- historyAddCapped limit (historyOptionStringCharCount textBox.RelativeHeight) total
    total <- historyAddCapped limit (historyShapeGeometryStringCharCount textBox.Geometry) total
    total <- historyAddCapped limit (historyGradientStringCharCount textBox.Gradient) total
    let mutable seen = 0
    let mutable paragraphs = textBox.Content
    while total <= limit
          && seen <= historyMaxRetainedSourceTraversalItems
          && not paragraphs.IsEmpty do
        total <- historyAddCapped limit (historyParagraphSourceTextCharCount paragraphs.Head) total
        seen <- seen + 1
        paragraphs <- paragraphs.Tail
    historyFailClosedForTraversalCount seen total

let private historyBlockSourceTextCharCount = function
    | Paragraph para
    | Heading (_, para) -> historyParagraphSourceTextCharCount para
    | Table table -> historyTableSourceTextCharCount table
    | TextBox textBox -> historyTextBoxSourceTextCharCount textBox

let private historyBlocksSourceTextCharCount (blocks: Block list) =
    let limit = historyMaxTotalRetainedSourceTextChars
    let mutable total = 0
    let mutable seen = 0
    let mutable remaining = blocks
    while total <= limit
          && seen <= historyMaxRetainedSourceTraversalItems
          && not remaining.IsEmpty do
        total <- historyAddCapped limit (historyBlockSourceTextCharCount remaining.Head) total
        seen <- seen + 1
        remaining <- remaining.Tail
    historyFailClosedForTraversalCount seen total

let private historyParagraphListSourceTextCharCount (paragraphs: ParagraphBlock list) =
    let limit = historyMaxTotalRetainedSourceTextChars
    let mutable total = 0
    let mutable seen = 0
    let mutable remaining = paragraphs
    while total <= limit
          && seen <= historyMaxRetainedSourceTraversalItems
          && not remaining.IsEmpty do
        total <- historyAddCapped limit (historyParagraphSourceTextCharCount remaining.Head) total
        seen <- seen + 1
        remaining <- remaining.Tail
    historyFailClosedForTraversalCount seen total

let private historyHeaderFooterPartsMetadataStringCharCount (parts: HeaderFooterPart list) =
    let limit = historyMaxTotalRetainedSourceTextChars
    let mutable total = 0
    let mutable seen = 0
    let mutable remaining = parts
    while total <= limit
          && seen <= historyMaxRetainedSourceTraversalItems
          && not remaining.IsEmpty do
        let part = remaining.Head
        total <- historyAddCapped limit part.Kind.Length total
        total <- historyAddCapped limit (historyOptionStringCharCount part.RelationshipId) total
        seen <- seen + 1
        remaining <- remaining.Tail
    historyFailClosedForTraversalCount seen total

let private historyHeaderFooterPartsBlocksSourceTextCharCount (parts: HeaderFooterPart list) =
    let limit = historyMaxTotalRetainedSourceTextChars
    let mutable total = 0
    let mutable seen = 0
    let mutable remaining = parts
    while total <= limit
          && seen <= historyMaxRetainedSourceTraversalItems
          && not remaining.IsEmpty do
        total <- historyAddCapped limit (historyBlocksSourceTextCharCount remaining.Head.Blocks) total
        seen <- seen + 1
        remaining <- remaining.Tail
    historyFailClosedForTraversalCount seen total

let private historyNamedStyleStringCharCount (style: NamedStyle) =
    let limit = historyMaxTotalRetainedSourceTextChars
    let mutable total = style.Id.Length
    total <- historyAddCapped limit style.StyleType.Length total
    total <- historyAddCapped limit (historyOptionStringCharCount style.Name) total
    total <- historyAddCapped limit (historyOptionStringCharCount style.BasedOn) total
    total <- historyAddCapped limit (historyOptionStringCharCount style.Next) total
    total <- historyAddCapped limit (historyOptionStringCharCount style.Linked) total
    total <- historyAddCapped limit (historyOptionStringCharCount style.Aliases) total
    total <- historyAddCapped limit (historyOptionStringCharCount style.LineSpacingRule) total
    total <- historyAddCapped limit (historyOptionStringCharCount style.BackgroundColor) total
    match style.Text with
    | Some textStyle -> historyAddCapped limit (historyTextStyleStringCharCount textStyle) total
    | None -> total

let private historyStylesSourceTextCharCount (styles: NamedStyle list) =
    let limit = historyMaxTotalRetainedSourceTextChars
    let mutable total = 0
    let mutable seen = 0
    let mutable remaining = styles
    while total <= limit
          && seen <= historyMaxRetainedSourceTraversalItems
          && not remaining.IsEmpty do
        total <- historyAddCapped limit (historyNamedStyleStringCharCount remaining.Head) total
        seen <- seen + 1
        remaining <- remaining.Tail
    historyFailClosedForTraversalCount seen total

let private historySelectionSourceTextCharCount = function
    | Some selection ->
        historyAddCapped
            historyMaxTotalRetainedSourceTextChars
            selection.Focus.BlockId.Length
            selection.Anchor.BlockId.Length
    | None -> 0

let private historyDocumentSourceTextCharCount (doc: Document) =
    let limit = historyMaxTotalRetainedSourceTextChars
    let mutable total = historySelectionSourceTextCharCount doc.Selection
    if total <= limit then
        total <- historyAddCapped limit (historyOptionStringCharCount doc.PageSetup.SectionStart) total
    if total <= limit then
        total <- historyAddCapped limit (historySectionColumnsStringCharCount doc.PageSetup.Columns) total
    if total <= limit then
        total <- historyAddCapped limit (historyStylesSourceTextCharCount doc.Styles) total
    if total <= limit then
        total <- historyAddCapped limit (historyHeaderFooterPartsMetadataStringCharCount doc.HeaderParts) total
    if total <= limit then
        total <- historyAddCapped limit (historyHeaderFooterPartsMetadataStringCharCount doc.FooterParts) total
    if total <= limit then
        total <- historyAddCapped limit (historyBlocksSourceTextCharCount doc.Blocks) total
    if total <= limit then
        total <- historyAddCapped limit (historyParagraphListSourceTextCharCount doc.Headers) total
    if total <= limit then
        total <- historyAddCapped limit (historyParagraphListSourceTextCharCount doc.Footers) total
    if total <= limit then
        total <- historyAddCapped limit (historyHeaderFooterPartsBlocksSourceTextCharCount doc.HeaderParts) total
    if total <= limit then
        total <- historyAddCapped limit (historyHeaderFooterPartsBlocksSourceTextCharCount doc.FooterParts) total
    total

let private trimHistory (stack: Document list) =
    let rec trim remaining count totalRetainedChars kept =
        match remaining with
        | [] -> List.rev kept
        | item :: rest ->
            let retainedChars = max 1 (historyDocumentSourceTextCharCount item)
            let nextRetainedChars = totalRetainedChars + retainedChars
            if count < maxUndoDepth
               && retainedChars <= historyMaxTotalRetainedSourceTextChars
               && nextRetainedChars <= historyMaxTotalRetainedSourceTextChars then
                trim rest (count + 1) nextRetainedChars (item :: kept)
            else
                trim rest count totalRetainedChars kept
    trim stack 0 0 []

let private sameListOrEqual left right =
    obj.ReferenceEquals(left, right) || left = right

let private documentsEqualForChange (left: Document) (right: Document) =
    obj.ReferenceEquals(left, right)
    || (left.Selection = right.Selection
        && left.PageSetup = right.PageSetup
        && sameListOrEqual left.Blocks right.Blocks
        && sameListOrEqual left.Styles right.Styles
        && sameListOrEqual left.Headers right.Headers
        && sameListOrEqual left.Footers right.Footers
        && sameListOrEqual left.HeaderParts right.HeaderParts
        && sameListOrEqual left.FooterParts right.FooterParts)

let private createLayoutCache () = {
    Width = nan
    CachedDocument = None
    CachedLayoutDocument = None
    Settings = None
    Layouts = None
    LayoutArray = None
    LayoutById = None
    ContentHeight = None
    RendererContentPaintCommands = None
    RendererContentCommands = None
    RendererSelectionLayouts = None
    RendererSelectionValue = None
    RendererSelectionCommands = None
    RendererFrameContentCommands = None
    RendererFrameSelectionCommands = None
    RendererFrameHoverCommands = None
    RendererFrameCommands = None
    ViewportSupportDocument = None
    ViewportSupportsContentCommands = None
}

let private layoutDictionary (layouts: WordLayout.BlockLayout list) =
    let dict = Dictionary<BlockId, WordLayout.BlockLayout>()
    for layout in layouts do
        if not (dict.ContainsKey layout.BlockId) then
            dict.[layout.BlockId] <- layout
    dict

let createState doc = {
    Document = signal doc
    HoveredImage = signal None
    ScrollY = signal 0.0
    DragAnchor = signal None
    WordDragAnchor = signal None
    ParagraphDragAnchor = signal None
    ResizingImage = signal None
    DraggingImage = signal None
    LastClick = signal None
    LastCursorRect = signal None
    PendingStyle = signal None
    UndoStack = signal []
    RedoStack = signal []
    LayoutCache = createLayoutCache()
    FocusId = Focus.allocId()
}

let private layoutSettings width (doc: Document) : WordLayout.LayoutSettings =
    let sourceWidth = max 1.0 doc.PageSetup.PageWidth
    let scale = width / sourceWidth
    let lineHeight = 24.0
    let hasFooter = not doc.Footers.IsEmpty || not doc.FooterParts.IsEmpty
    let footerReserve =
        if hasFooter then doc.PageSetup.FooterDistance * scale + lineHeight * 1.5
        else doc.PageSetup.MarginBottom * scale
    {
        Width = width
        PageHeight = max 1.0 (doc.PageSetup.PageHeight * scale)
        FontSize = 16.0
        LineHeight = lineHeight
        PaddingLeft = doc.PageSetup.MarginLeft * scale
        PaddingTop = RedBat.Word.WordLayout.Page.bodyTopWithRunningHeader doc * scale
        MarginRight = doc.PageSetup.MarginRight * scale
        MarginBottom = max (doc.PageSetup.MarginBottom * scale) footerReserve
        InTableCell = false
        TableContinuationHeaderHeight = 0.0
        LongTrailingSpaceFitSafety = 0.25
    }

let private layoutInputsEqual (left: Document) (right: Document) =
    Object.ReferenceEquals(left.Blocks, right.Blocks)
    && left.PageSetup = right.PageSetup
    && Object.ReferenceEquals(left.Styles, right.Styles)
    && Object.ReferenceEquals(left.Headers, right.Headers)
    && Object.ReferenceEquals(left.Footers, right.Footers)
    && Object.ReferenceEquals(left.HeaderParts, right.HeaderParts)
    && Object.ReferenceEquals(left.FooterParts, right.FooterParts)

let private cachedLayouts width (state: WordEditorState) doc =
    let cache = state.LayoutCache
    let matches =
        cache.Width = width &&
        match cache.CachedDocument with
        | Some cachedDoc when Object.ReferenceEquals(cachedDoc, doc) -> true
        | _ ->
            match cache.CachedLayoutDocument with
            | Some cachedLayoutDoc -> layoutInputsEqual cachedLayoutDoc doc
            | None -> false

    match matches, cache.Settings, cache.Layouts with
    | true, Some settings, Some layouts ->
        cache.CachedDocument <- Some doc
        settings, layouts
    | _ ->
        let settings = layoutSettings width doc
        let layouts = WordLayout.layoutDocument settings measureText doc
        cache.Width <- width
        cache.CachedDocument <- Some doc
        cache.CachedLayoutDocument <- Some doc
        cache.Settings <- Some settings
        cache.Layouts <- Some layouts
        cache.LayoutArray <- Some (Array.ofList layouts)
        cache.LayoutById <- Some (layoutDictionary layouts)
        cache.ContentHeight <- None
        cache.RendererContentPaintCommands <- None
        cache.RendererContentCommands <- None
        cache.RendererSelectionLayouts <- None
        cache.RendererSelectionValue <- None
        cache.RendererSelectionCommands <- None
        cache.RendererFrameContentCommands <- None
        cache.RendererFrameSelectionCommands <- None
        cache.RendererFrameHoverCommands <- None
        cache.RendererFrameCommands <- None
        cache.ViewportSupportDocument <- None
        cache.ViewportSupportsContentCommands <- None
        settings, layouts

let private cachedLayoutArray width (state: WordEditorState) doc =
    let settings, layouts = cachedLayouts width state doc
    match state.LayoutCache.LayoutArray with
    | Some layoutArray ->
        settings, layoutArray
    | _ ->
        let layoutArray = Array.ofList layouts
        state.LayoutCache.LayoutArray <- Some layoutArray
        settings, layoutArray

let private cachedLayoutLookup width (state: WordEditorState) doc =
    let settings, layouts = cachedLayouts width state doc
    let dict =
        match state.LayoutCache.LayoutById with
        | Some dict ->
            dict
        | None ->
            let dict = layoutDictionary layouts
            state.LayoutCache.LayoutById <- Some dict
            dict
    let lookup blockId =
        let mutable layout = Unchecked.defaultof<WordLayout.BlockLayout>
        if dict.TryGetValue(blockId, &layout) then Some layout else None
    settings, lookup

let hitTest width doc x y =
    WordLayout.hitTest (layoutSettings width doc) measureText doc x y

let private hitTestCached width (state: WordEditorState) doc x y =
    let settings, layouts = cachedLayoutArray width state doc
    WordLayout.hitTestLayoutArrayWithRange (WordLayout.rangeMeasurerFromTextMeasurer measureText) settings.PaddingLeft layouts x y

let documentHeight width doc =
    let settings = layoutSettings width doc
    let layouts = WordLayout.layoutDocument settings measureText doc
    match layouts with
    | [] -> settings.PaddingTop * 2.0 + settings.LineHeight
    | _ ->
        let mutable bottom = 0.0
        for layout in layouts do
            bottom <- max bottom (layout.Y + layout.Height)
        bottom + settings.PaddingTop

let private documentHeightCached width (state: WordEditorState) doc =
    let settings, layouts = cachedLayouts width state doc
    match state.LayoutCache.ContentHeight with
    | Some height -> height
    | None ->
        let height =
            match layouts with
            | [] -> settings.PaddingTop * 2.0 + settings.LineHeight
            | _ ->
                let mutable bottom = 0.0
                for layout in layouts do
                    bottom <- max bottom (layout.Y + layout.Height)
                bottom + settings.PaddingTop
        state.LayoutCache.ContentHeight <- Some height
        height

let setCursorAt width x y (state: WordEditorState) =
    match hitTestCached width state state.Document.Value x y with
    | None -> false
    | Some (blockId, offset) ->
        state.Document.Value <- setCursor blockId offset state.Document.Value
        true

let private positionAt width x y (state: WordEditorState) =
    hitTestCached width state state.Document.Value x y
    |> Option.map (fun (blockId, offset) -> { BlockId = blockId; Offset = offset })

let private imageAt width x y (state: WordEditorState) =
    let settings, layouts = cachedLayoutArray width state state.Document.Value
    WordLayout.hitTestForImageLayoutArray settings.PaddingLeft layouts x y

let private isResizeHandle x y imageX imageY imageWidth imageHeight =
    let handleSize = 8.0
    let half = handleSize / 2.0
    let nearCorner cx cy =
        x >= cx - half && x <= cx + half && y >= cy - half && y <= cy + half

    nearCorner imageX imageY
    || nearCorner (imageX + imageWidth) imageY
    || nearCorner imageX (imageY + imageHeight)
    || nearCorner (imageX + imageWidth) (imageY + imageHeight)

let private isWordChar (c: char) = System.Char.IsLetterOrDigit(c) || c = '_'

let private wordRangeAt (position: Position) doc =
    match findParagraph doc position.BlockId with
    | None -> None
    | Some para ->
        let text = WordTextBuffer.toString para.Buffer
        if text = "" then Some (position, position)
        else
            let index = max 0 (min (text.Length - 1) position.Offset)
            let isSelectedClass = isWordChar text.[index]
            let mutable startOffset = index
            let mutable endOffset = index + 1

            while startOffset > 0 && isWordChar text.[startOffset - 1] = isSelectedClass do
                startOffset <- startOffset - 1

            while endOffset < text.Length && isWordChar text.[endOffset] = isSelectedClass do
                endOffset <- endOffset + 1

            Some ({ position with Offset = startOffset }, { position with Offset = endOffset })

let private selectWordAt (position: Position) doc =
    match wordRangeAt position doc with
    | Some (startPosition, endPosition) -> setSelection startPosition endPosition doc
    | None -> doc

let private selectParagraphAt (position: Position) doc =
    match findParagraph doc position.BlockId with
    | None -> doc
    | Some para ->
        let length = WordTextBuffer.length para.Buffer
        setSelection { position with Offset = 0 } { position with Offset = length } doc

let private paragraphRangeAt (position: Position) doc =
    match findParagraph doc position.BlockId with
    | None -> None
    | Some para ->
        let length = WordTextBuffer.length para.Buffer
        Some ({ position with Offset = 0 }, { position with Offset = length })

let private endOfParagraph blockId doc =
    match findParagraph doc blockId with
    | Some para -> { BlockId = blockId; Offset = WordTextBuffer.length para.Buffer }
    | None -> { BlockId = blockId; Offset = 0 }

let private imageMoveTarget width x y blockId sourceOffset state doc =
    let target =
        match hitTestCached width state doc x y with
        | Some (targetBlockId, targetOffset) -> { BlockId = targetBlockId; Offset = targetOffset }
        | None -> endOfParagraph blockId doc

    if target.BlockId = blockId && target.Offset <= sourceOffset then endOfParagraph blockId doc
    else target

let private documentStart doc : Position option =
    let mutable found = None
    for block in doc.Blocks do
        if found.IsNone then
            match block with
            | Paragraph p
            | Heading (_, p) -> found <- Some { BlockId = p.Id; Offset = 0 }
            | Table table ->
                for row in table.Rows do
                    if found.IsNone then
                        for cell in row.Cells do
                            if found.IsNone then
                                match cell.Content with
                                | para :: _ -> found <- Some { BlockId = para.Id; Offset = 0 }
                                | [] -> ()
            | TextBox textBox ->
                match textBox.Content with
                | para :: _ -> found <- Some { BlockId = para.Id; Offset = 0 }
                | [] -> ()
    found

let private documentEnd doc : Position option =
    let mutable last: Position option = None
    let setLast (para: ParagraphBlock) =
        last <- Some { BlockId = para.Id; Offset = WordTextBuffer.length para.Buffer }
    for block in doc.Blocks do
        match block with
        | Paragraph p
        | Heading (_, p) -> setLast p
        | Table table ->
            for row in table.Rows do
                for cell in row.Cells do
                    for para in cell.Content do
                        setLast para
        | TextBox textBox ->
            for para in textBox.Content do
                setLast para
    last

let private moveToDocumentBoundary (boundary: Document -> Position option) doc =
    match boundary doc with
    | Some position -> setCursor position.BlockId position.Offset doc
    | None -> doc

let private extendToDocumentBoundary (boundary: Document -> Position option) doc =
    match doc.Selection, boundary doc with
    | Some selection, Some position -> { doc with Selection = Some { selection with Focus = position } }
    | _ -> doc

let private previousWordOffset (text: string) offset =
    let mutable index = max 0 (min text.Length offset)
    while index > 0 && not (isWordChar text.[index - 1]) do
        index <- index - 1
    while index > 0 && isWordChar text.[index - 1] do
        index <- index - 1
    index

let private nextWordOffset (text: string) offset =
    let mutable index = max 0 (min text.Length offset)
    while index < text.Length && not (isWordChar text.[index]) do
        index <- index + 1
    while index < text.Length && isWordChar text.[index] do
        index <- index + 1
    index

let private moveCursorByWord direction doc =
    match doc.Selection with
    | None -> doc
    | Some selection ->
        match findParagraph doc selection.Focus.BlockId with
        | None -> doc
        | Some para ->
            let text = WordTextBuffer.toString para.Buffer
            let offset =
                if direction < 0 then previousWordOffset text selection.Focus.Offset
                else nextWordOffset text selection.Focus.Offset
            setCursor selection.Focus.BlockId offset doc

let private extendSelectionByWord direction doc =
    match doc.Selection with
    | None -> doc
    | Some selection ->
        match findParagraph doc selection.Focus.BlockId with
        | None -> doc
        | Some para ->
            let text = WordTextBuffer.toString para.Buffer
            let offset =
                if direction < 0 then previousWordOffset text selection.Focus.Offset
                else nextWordOffset text selection.Focus.Offset
            { doc with Selection = Some { selection with Focus = { selection.Focus with Offset = offset } } }

let private deleteWord direction doc =
    match doc.Selection with
    | None -> doc
    | Some selection when not selection.IsCollapsed -> deleteSelection doc
    | Some selection ->
        match findParagraph doc selection.Focus.BlockId with
        | None -> doc
        | Some para ->
            let text = WordTextBuffer.toString para.Buffer
            let offset =
                if direction < 0 then previousWordOffset text selection.Focus.Offset
                else nextWordOffset text selection.Focus.Offset

            if offset = selection.Focus.Offset then doc
            else
                let rangeStart = min offset selection.Focus.Offset
                let rangeEnd = max offset selection.Focus.Offset
                { doc with
                    Selection = Some {
                        Anchor = { selection.Focus with Offset = rangeStart }
                        Focus = { selection.Focus with Offset = rangeEnd }
                    } }
                |> deleteSelection

let private styleAtCursor doc =
    match doc.Selection with
    | Some selection ->
        match findParagraph doc selection.Focus.BlockId with
        | Some para ->
            let length = WordTextBuffer.length para.Buffer
            let offset = if selection.Focus.Offset > 0 && selection.Focus.Offset >= length then selection.Focus.Offset - 1 else selection.Focus.Offset
            WordTextBuffer.getStyleAt offset para.Buffer
        | None -> WordTextBuffer.defaultStyle
    | None -> WordTextBuffer.defaultStyle

let private togglePendingStyle transform state =
    let before = state.PendingStyle.Value |> Option.defaultValue (styleAtCursor state.Document.Value)
    state.PendingStyle.Value <- Some (transform before)
    Platform.requestRedraw()
    true

let private hasCollapsedSelection doc =
    match doc.Selection with
    | Some selection -> selection.IsCollapsed
    | None -> false

let private clickCount x y state =
    let now = Platform.now()
    let count =
        match state.LastClick.Value with
        | Some (lastTime, lastX, lastY, lastCount) ->
            let dx = x - lastX
            let dy = y - lastY
            if now > lastTime && now - lastTime <= 0.45 && dx * dx + dy * dy <= 36.0 then min 3 (lastCount + 1)
            else 1
        | None -> 1

    state.LastClick.Value <- Some (now, x, y, count)
    count

let private caretRect width (state: WordEditorState) doc =
    match doc.Selection with
    | Some selection when selection.IsCollapsed ->
        let settings, lookupLayout = cachedLayoutLookup width state doc
        match lookupLayout selection.Focus.BlockId with
        | None -> None
        | Some layout ->
            let line = WordLayout.lineForOffset layout selection.Focus.Offset
            let x = WordLayout.xForOffset measureText line selection.Focus.Offset
            let lineAlignmentOffset = WordLayout.alignmentOffset layout line
            Some {
                X = settings.PaddingLeft + x + lineAlignmentOffset
                Y = line.Y - state.ScrollY.Value
                Width = 2.0
                Height = line.Height
            }
    | _ -> None

let private updateCursorRect width state doc =
    let nextRect = if Focus.isFocused state.FocusId then caretRect width state doc else None
    if nextRect <> state.LastCursorRect.Value then
        match nextRect with
        | Some rect ->
            Platform.onTextInputFocus rect
            Platform.onCursorRectChange rect
        | None -> Platform.onTextInputBlur()

        state.LastCursorRect.Value <- nextRect

let private applyDocumentChange (state: WordEditorState) transform =
    let before = state.Document.Value
    let after = transform before
    if not (documentsEqualForChange after before) then
        state.UndoStack.Value <- (before :: state.UndoStack.Value) |> trimHistory
        state.RedoStack.Value <- []
        state.Document.Value <- after
        Platform.requestRedraw()
        true
    else
        false

let private applyDocumentChangeNoHistory (state: WordEditorState) transform =
    let before = state.Document.Value
    let after = transform before
    if not (documentsEqualForChange after before) then
        state.Document.Value <- after
        Platform.requestRedraw()
        true
    else
        false

let private insertPlainText text doc =
    if System.String.IsNullOrEmpty text then doc
    else
        let mutable nextDoc = doc
        let mutable segmentStart = 0
        let mutable i = 0
        let flushSegment segmentEnd =
            nextDoc <- insertText (text.Substring(segmentStart, segmentEnd - segmentStart)) nextDoc

        while i < text.Length do
            let ch = text.[i]
            if ch = '\r' || ch = '\n' then
                flushSegment i
                nextDoc <- splitBlock nextDoc
                if ch = '\r' && i + 1 < text.Length && text.[i + 1] = '\n' then
                    i <- i + 1
                segmentStart <- i + 1
            i <- i + 1

        flushSegment text.Length

        nextDoc

let private insertStyledText text style doc =
    match doc.Selection with
    | None -> doc
    | Some selection ->
        let docBeforeInsert = if selection.IsCollapsed then doc else deleteSelection doc
        match docBeforeInsert.Selection with
        | Some cursor ->
            match findParagraph docBeforeInsert cursor.Focus.BlockId with
            | Some para ->
                let startOffset = cursor.Focus.Offset
                let insertedBuffer = WordTextBuffer.insert startOffset text para.Buffer
                let styledBuffer = WordTextBuffer.applyStyle startOffset text.Length (fun _ -> style) insertedBuffer
                let nextDoc = replaceParagraph docBeforeInsert para.Id { para with Buffer = styledBuffer }
                { nextDoc with
                    Selection = Some {
                        Anchor = { cursor.Focus with Offset = startOffset + text.Length }
                        Focus = { cursor.Focus with Offset = startOffset + text.Length }
                    } }
            | None -> docBeforeInsert
        | None -> docBeforeInsert

let private insertPlainTextWithPendingStyle state text doc =
    match state.PendingStyle.Value with
    | None -> insertPlainText text doc
    | Some style when not (text.Contains("\n") || text.Contains("\r")) -> insertStyledText text style doc
    | Some _ -> insertPlainText text doc

let private autoListAtCursor doc =
    match doc.Selection with
    | Some selection when selection.IsCollapsed ->
        match findParagraph doc selection.Focus.BlockId with
        | Some para ->
            let beforeCursor = WordTextBuffer.slice 0 selection.Focus.Offset para.Buffer
            let marker =
                match beforeCursor with
                | "-" | "*" -> Some (1, Bullet)
                | "1." -> Some (2, Numbered)
                | _ -> None

            match marker with
            | Some (markerLength, listStyle) ->
                let nextBuffer = WordTextBuffer.delete 0 markerLength para.Buffer
                let nextPara = {
                    para with
                        Buffer = nextBuffer
                        Metadata = { para.Metadata with ListStyle = listStyle; ListLevel = max 1 para.Metadata.ListLevel }
                }
                let nextDoc = replaceParagraph doc para.Id nextPara
                { nextDoc with
                    Selection = Some {
                        Anchor = { selection.Focus with Offset = 0 }
                        Focus = { selection.Focus with Offset = 0 }
                    } }
            | None -> insertPlainText " " doc
        | None -> insertPlainText " " doc
    | _ -> insertPlainText " " doc

let private clipRuns startOffset endOffset shift (buffer: WordTextBuffer.TextBuffer) =
    let runs = ResizeArray<WordTextBuffer.Run>()
    for runValue in buffer.Runs do
        match runValue with
        | WordTextBuffer.Text run ->
            let runStart = run.Start
            let runEnd = run.Start + run.Length
            let clippedStart = max runStart startOffset
            let clippedEnd = min runEnd endOffset
            if clippedStart < clippedEnd then
                runs.Add(WordTextBuffer.Text {
                    run with
                        Start = clippedStart + shift
                        Length = clippedEnd - clippedStart
                })
        | WordTextBuffer.Image image ->
            if image.Start >= startOffset && image.Start < endOffset then
                runs.Add(WordTextBuffer.Image { image with Start = image.Start + shift })
    resizeArrayToList runs

let private splitBufferPreservingRuns offset buffer =
    let length = WordTextBuffer.length buffer
    let beforeText = WordTextBuffer.slice 0 offset buffer
    let afterText = WordTextBuffer.slice offset (length - offset) buffer
    let beforeBuffer = WordTextBuffer.ofString beforeText
    let afterBuffer = WordTextBuffer.ofString afterText
    ({ beforeBuffer with Runs = clipRuns 0 offset 0 buffer },
     { afterBuffer with Runs = clipRuns offset length -offset buffer })

let private splitBlockPreservingStyles doc =
    match doc.Selection with
    | None -> doc
    | Some selection ->
        let addSplit para makeBeforeBlock makeAfterBlock =
            let beforeBuffer, afterBuffer = splitBufferPreservingRuns selection.Focus.Offset para.Buffer
            let newBlockId = "editor-block-" + System.Guid.NewGuid().ToString("N")
            let beforePara = { para with Buffer = beforeBuffer }
            let afterPara = { para with Id = newBlockId; Buffer = afterBuffer }
            (makeBeforeBlock beforePara,
             makeAfterBlock afterPara,
             {
                 Anchor = { BlockId = newBlockId; Offset = 0 }
                 Focus = { BlockId = newBlockId; Offset = 0 }
             })
        let prefix = ResizeArray<Block>()
        let mutable remaining = doc.Blocks
        let mutable result: Document option = None
        while result.IsNone && not remaining.IsEmpty do
            match remaining with
            | Paragraph para :: tail when para.Id = selection.Focus.BlockId ->
                let beforeBlock, afterBlock, nextSelection = addSplit para Paragraph Paragraph
                result <-
                    Some {
                        doc with
                            Blocks = prependResizeArray prefix (beforeBlock :: afterBlock :: tail)
                            Selection = Some nextSelection
                    }
            | Heading (level, para) :: tail when para.Id = selection.Focus.BlockId ->
                let beforeBlock, afterBlock, nextSelection = addSplit para (fun p -> Heading (level, p)) (fun p -> Heading (level, p))
                result <-
                    Some {
                        doc with
                            Blocks = prependResizeArray prefix (beforeBlock :: afterBlock :: tail)
                            Selection = Some nextSelection
                    }
            | block :: tail ->
                prefix.Add block
                remaining <- tail
            | [] -> ()
        result |> Option.defaultWith (fun () -> splitBlock doc)

let private autoScrollForPointer viewportHeight contentHeight y state =
    let margin = 32.0
    let step = 48.0
    let top = state.ScrollY.Value
    let bottom = top + viewportHeight
    let maxScrollY = max 0.0 (contentHeight - viewportHeight)
    let nextScrollY =
        if y < top + margin then max 0.0 (top - step)
        elif y > bottom - margin then min maxScrollY (top + step)
        else top

    if nextScrollY <> top then
        state.ScrollY.Value <- nextScrollY
        Platform.requestRedraw()

let private enterAtCursor doc =
    match doc.Selection with
    | Some selection ->
        match findParagraph doc selection.Focus.BlockId with
        | Some para when WordTextBuffer.length para.Buffer = 0 && para.Metadata.ListStyle <> NoList -> decreaseIndent doc
        | _ -> splitBlockPreservingStyles doc
    | None -> doc

let private undo state =
    match state.UndoStack.Value with
    | previous :: rest ->
        state.RedoStack.Value <- (state.Document.Value :: state.RedoStack.Value) |> trimHistory
        state.UndoStack.Value <- rest
        state.Document.Value <- previous
        Platform.requestRedraw()
        true
    | [] -> true

let private redo state =
    match state.RedoStack.Value with
    | next :: rest ->
        state.UndoStack.Value <- (state.Document.Value :: state.UndoStack.Value) |> trimHistory
        state.RedoStack.Value <- rest
        state.Document.Value <- next
        Platform.requestRedraw()
        true
    | [] -> true

let private cancelSelection state =
    state.DragAnchor.Value <- None
    state.WordDragAnchor.Value <- None
    state.ParagraphDragAnchor.Value <- None
    state.ResizingImage.Value <- None
    state.DraggingImage.Value <- None

    match state.Document.Value.Selection with
    | Some selection when not selection.IsCollapsed ->
        state.Document.Value <- setCursor selection.Focus.BlockId selection.Focus.Offset state.Document.Value
        Platform.requestRedraw()
        true
    | _ -> true

let private extendSelectionBy move doc =
    match doc.Selection with
    | None -> doc
    | Some selection ->
        let moved = move doc
        match moved.Selection with
        | Some movedSelection -> { moved with Selection = Some { selection with Focus = movedSelection.Focus } }
        | None -> doc

let private caretDocumentPoint width state doc =
    match doc.Selection with
    | Some selection ->
        let settings, lookupLayout = cachedLayoutLookup width state doc
        match lookupLayout selection.Focus.BlockId with
        | Some layout ->
            let line = WordLayout.lineForOffset layout selection.Focus.Offset
            let x = WordLayout.xForOffset measureText line selection.Focus.Offset
            let lineAlignmentOffset = WordLayout.alignmentOffset layout line
            Some (settings.PaddingLeft + x + lineAlignmentOffset, line.Y + line.Height / 2.0)
        | None -> None
    | None -> None

let private movePage width height direction extendSelection state =
    let doc = state.Document.Value
    let amount = max 24.0 (height * 0.8)
    let targetPosition =
        match caretDocumentPoint width state doc with
        | Some (x, y) ->
            let targetY = y + amount * float direction
            hitTestCached width state doc x targetY
            |> Option.map (fun (blockId, offset) -> { BlockId = blockId; Offset = offset })
        | None -> None

    let nextDoc =
        match doc.Selection, targetPosition with
        | Some selection, Some position when extendSelection -> { doc with Selection = Some { selection with Focus = position } }
        | _, Some position -> setCursor position.BlockId position.Offset doc
        | Some selection, None when direction < 0 && extendSelection -> { doc with Selection = Some { selection with Focus = documentStart doc |> Option.defaultValue selection.Focus } }
        | Some selection, None when direction > 0 && extendSelection -> { doc with Selection = Some { selection with Focus = documentEnd doc |> Option.defaultValue selection.Focus } }
        | _, None when direction < 0 -> moveToDocumentBoundary documentStart doc
        | _, None -> moveToDocumentBoundary documentEnd doc

    state.ScrollY.Value <- max 0.0 (state.ScrollY.Value + amount * float direction)
    if not (documentsEqualForChange nextDoc doc) then state.Document.Value <- nextDoc
    Platform.requestRedraw()
    true

let private moveVisualLine width direction extendSelection state doc =
    let targetPosition =
        match caretDocumentPoint width state doc with
        | Some (x, y) ->
            let settings = layoutSettings width doc
            let targetY = y + settings.LineHeight * float direction
            hitTestCached width state doc x targetY
            |> Option.map (fun (blockId, offset) -> { BlockId = blockId; Offset = offset })
        | None -> None

    match doc.Selection, targetPosition with
    | Some selection, Some position when extendSelection -> { doc with Selection = Some { selection with Focus = position } }
    | _, Some position -> setCursor position.BlockId position.Offset doc
    | Some selection, None when direction < 0 && extendSelection -> { doc with Selection = Some { selection with Focus = documentStart doc |> Option.defaultValue selection.Focus } }
    | Some selection, None when direction > 0 && extendSelection -> { doc with Selection = Some { selection with Focus = documentEnd doc |> Option.defaultValue selection.Focus } }
    | _, None when direction < 0 -> moveToDocumentBoundary documentStart doc
    | _, None -> moveToDocumentBoundary documentEnd doc

let private moveToVisualLineBoundary width toEnd extendSelection state doc =
    let visualLineForBoundary toEnd offset (layout: WordLayout.BlockLayout) =
        let mutable firstPass = None
        let mutable secondPass = None
        let mutable remaining = layout.Lines
        while (firstPass.IsNone || secondPass.IsNone) && not remaining.IsEmpty do
            match remaining with
            | line :: tail ->
                if toEnd then
                    if firstPass.IsNone && offset >= line.StartOffset && offset < line.EndOffset then
                        firstPass <- Some line
                    if secondPass.IsNone && offset = line.EndOffset then
                        secondPass <- Some line
                else
                    if firstPass.IsNone && offset = line.StartOffset then
                        firstPass <- Some line
                    if secondPass.IsNone && offset > line.StartOffset && offset <= line.EndOffset then
                        secondPass <- Some line
                remaining <- tail
            | [] -> ()
        firstPass
        |> Option.orElse secondPass
        |> Option.defaultValue (WordLayout.lineForOffset layout offset)

    let targetPosition =
        match doc.Selection with
        | Some selection ->
            let _, lookupLayout = cachedLayoutLookup width state doc
            match lookupLayout selection.Focus.BlockId with
            | Some layout ->
                let line = visualLineForBoundary toEnd selection.Focus.Offset layout
                Some { BlockId = selection.Focus.BlockId; Offset = if toEnd then line.EndOffset else line.StartOffset }
            | None -> None
        | None -> None

    match doc.Selection, targetPosition with
    | Some selection, Some position when extendSelection -> { doc with Selection = Some { selection with Focus = position } }
    | _, Some position -> setCursor position.BlockId position.Offset doc
    | _ -> doc

let private handleKey width height state (ev: KeyEvent) =
    match ev with
    | { Key = "Escape" } -> cancelSelection state
    | { Ctrl = true; Key = key } when key = "a" || key = "A" -> applyDocumentChangeNoHistory state selectAll
    | { Ctrl = true; Shift = false; Key = key } when key = "z" || key = "Z" -> undo state
    | { Ctrl = true; Key = key } when key = "y" || key = "Y" -> redo state
    | { Ctrl = true; Shift = true; Key = key } when key = "z" || key = "Z" -> redo state
    | { Ctrl = true; Key = key } when key = "c" || key = "C" ->
        let text = getSelectionText state.Document.Value
        if text <> "" then Platform.clipboard.SetText text
        true
    | { Ctrl = true; Key = key } when key = "x" || key = "X" ->
        let text = getSelectionText state.Document.Value
        if text <> "" then
            Platform.clipboard.SetText text
            applyDocumentChange state deleteSelection
        else
            true
    | { Ctrl = true; Key = key } when key = "v" || key = "V" ->
        match Platform.clipboard.GetText() with
        | Some text when text <> "" -> applyDocumentChange state (insertPlainTextWithPendingStyle state text)
        | _ -> true
    | { Ctrl = true; Key = key } when key = "b" || key = "B" ->
        if hasCollapsedSelection state.Document.Value then togglePendingStyle (fun style -> { style with Bold = not style.Bold }) state
        else applyDocumentChange state toggleBold
    | { Ctrl = true; Key = key } when key = "i" || key = "I" ->
        if hasCollapsedSelection state.Document.Value then togglePendingStyle (fun style -> { style with Italic = not style.Italic }) state
        else applyDocumentChange state toggleItalic
    | { Ctrl = true; Key = key } when key = "u" || key = "U" ->
        if hasCollapsedSelection state.Document.Value then togglePendingStyle (fun style -> { style with Underline = not style.Underline }) state
        else applyDocumentChange state toggleUnderline
    | { Ctrl = true; Shift = true; Key = "8" } -> applyDocumentChange state toggleBulletList
    | { Ctrl = true; Shift = true; Key = "7" } -> applyDocumentChange state toggleNumberedList
    | { Ctrl = true; Shift = true; Key = key } when key = "l" || key = "L" -> applyDocumentChange state alignLeft
    | { Ctrl = true; Shift = true; Key = key } when key = "e" || key = "E" -> applyDocumentChange state alignCenter
    | { Ctrl = true; Shift = true; Key = key } when key = "r" || key = "R" -> applyDocumentChange state alignRight
    | { Ctrl = true; Shift = true; Key = key } when key = "j" || key = "J" -> applyDocumentChange state alignJustify
    | { Ctrl = true; Shift = true; Key = "ArrowLeft" } -> applyDocumentChangeNoHistory state (extendSelectionByWord -1)
    | { Ctrl = true; Shift = true; Key = "ArrowRight" } -> applyDocumentChangeNoHistory state (extendSelectionByWord 1)
    | { Ctrl = true; Key = "ArrowLeft" } -> applyDocumentChangeNoHistory state (moveCursorByWord -1)
    | { Ctrl = true; Key = "ArrowRight" } -> applyDocumentChangeNoHistory state (moveCursorByWord 1)
    | { Ctrl = true; Key = "Backspace" } -> applyDocumentChange state (deleteWord -1)
    | { Ctrl = true; Key = "Delete" } -> applyDocumentChange state (deleteWord 1)
    | { Ctrl = true; Shift = true; Key = "Home" } -> applyDocumentChangeNoHistory state (extendToDocumentBoundary documentStart)
    | { Ctrl = true; Shift = true; Key = "End" } -> applyDocumentChangeNoHistory state (extendToDocumentBoundary documentEnd)
    | { Ctrl = true; Key = "Home" } -> applyDocumentChangeNoHistory state (moveToDocumentBoundary documentStart)
    | { Ctrl = true; Key = "End" } -> applyDocumentChangeNoHistory state (moveToDocumentBoundary documentEnd)
    | { Key = "Backspace" } -> applyDocumentChange state deleteBackward
    | { Key = "Delete" } -> applyDocumentChange state deleteForward
    | { Key = "ArrowLeft"; Shift = true } -> applyDocumentChangeNoHistory state extendSelectionLeft
    | { Key = "ArrowRight"; Shift = true } -> applyDocumentChangeNoHistory state extendSelectionRight
    | { Key = "ArrowUp"; Shift = true } -> applyDocumentChangeNoHistory state (moveVisualLine width -1 true state)
    | { Key = "ArrowDown"; Shift = true } -> applyDocumentChangeNoHistory state (moveVisualLine width 1 true state)
    | { Key = "Home"; Shift = true } -> applyDocumentChangeNoHistory state (moveToVisualLineBoundary width false true state)
    | { Key = "End"; Shift = true } -> applyDocumentChangeNoHistory state (moveToVisualLineBoundary width true true state)
    | { Key = "PageUp"; Shift = true } -> movePage width height -1 true state
    | { Key = "PageDown"; Shift = true } -> movePage width height 1 true state
    | { Key = "PageUp" } -> movePage width height -1 false state
    | { Key = "PageDown" } -> movePage width height 1 false state
    | { Key = "ArrowLeft" } -> applyDocumentChangeNoHistory state moveCursorLeft
    | { Key = "ArrowRight" } -> applyDocumentChangeNoHistory state moveCursorRight
    | { Key = "ArrowUp" } -> applyDocumentChangeNoHistory state (moveVisualLine width -1 false state)
    | { Key = "ArrowDown" } -> applyDocumentChangeNoHistory state (moveVisualLine width 1 false state)
    | { Key = "Home" } -> applyDocumentChangeNoHistory state (moveToVisualLineBoundary width false false state)
    | { Key = "End" } -> applyDocumentChangeNoHistory state (moveToVisualLineBoundary width true false state)
    | { Key = "Enter" } -> applyDocumentChange state enterAtCursor
    | { Key = "Tab"; Shift = true } -> applyDocumentChange state decreaseIndent
    | { Key = "Tab" } -> applyDocumentChange state increaseIndent
    | { Ctrl = false; Alt = false; Key = " " } -> applyDocumentChange state autoListAtCursor
    | { Ctrl = false; Alt = false; Key = key } when key.Length = 1 -> applyDocumentChange state (insertPlainTextWithPendingStyle state key)
    | _ -> false

let private toRendererCommand (resolveImage: string -> Framebuffer option) = function
    | WordPaint.DrawText (text, style, x, baseline, _) ->
        [ Cmd.DrawGlyphs(vec2 x baseline, glyphLayoutForPaintText text style, Solid (parseHexColor style.Color)) ]

    | WordPaint.DrawUnderline (color, _lineStyle, x, y, width) ->
        [ Cmd.Line(vec2 x y, vec2 (x + width) y, parseHexColor color, px 1.0) ]

    | WordPaint.DrawImage (src, x, y, width, height, _, _, _, _) ->
        match resolveImage src with
        | Some fb ->
            [ Cmd.DrawImage(fb, rect 0.0 0.0 (float fb.W) (float fb.H), rect x y width height, true) ]
        | None ->
            // The browser path can draw data URLs directly. Widgets need a decoded framebuffer.
            [ Cmd.StrokeRect(rect x y width height, Color.rgb 180uy 180uy 180uy, px 1.0) ]

    | WordPaint.FillRect (color, x, y, width, height) ->
        [ Cmd.FillRect(rect x y width height, parseHexColor color) ]

    | WordPaint.StrokeRect (color, lineWidth, x, y, width, height) ->
        [ Cmd.StrokeRect(rect x y width height, parseHexColor color, px lineWidth) ]

    | WordPaint.StrokeLine (color, lineWidth, dashStyle, dashOffset, x1, y1, x2, y2) ->
        let dashArray =
            match dashStyle with
            | 1 -> [px lineWidth; px lineWidth]
            | 2 -> [px (lineWidth * 3.0); px (lineWidth * 1.5)]
            | 3 -> [px (lineWidth * 3.0); px (lineWidth * 1.5); px lineWidth; px (lineWidth * 1.5)]
            | _ -> []
        let style = {
            Width = px lineWidth
            Cap = Butt
            Join = Miter
            Paint = Solid (parseHexColor color)
            DashArray = dashArray
            DashOffset = px dashOffset
            MiterLimit = 4.0
        }
        [ Cmd.StrokePath([MoveTo (vec2 x1 y1); LineTo (vec2 x2 y2)], style) ]

    | WordPaint.DrawShape (fill, stroke, strokeWidth, (bx, by, bw, bh), segs) ->
        let path = ResizeArray()
        for segment in segs do
            path.Add(
                match segment with
                | PMove (x, y) -> MoveTo (vec2 x y)
                | PLine (x, y) -> LineTo (vec2 x y)
                | PQuad (cx, cy, x, y) -> QuadTo (vec2 cx cy, vec2 x y)
                | PCubic (a, b, c, d, x, y) -> CubicTo (vec2 a b, vec2 c d, vec2 x y)
                | PClose -> Close)
        let rpath = resizeArrayToList path
        let fillCmds =
            match fill with
            | WordPaint.FillNone -> []
            | WordPaint.FillSolid c -> [ Cmd.FillPath(rpath, Solid (parseHexColor c), NonZero) ]
            | WordPaint.FillGradient g ->
                let stopBuffer = ResizeArray()
                for stop in g.Stops do
                    stopBuffer.Add { Offset = stop.Offset; StopColor = parseHexColor stop.Color }
                let stops = resizeArrayToList stopBuffer
                let cx, cy = bx + bw / 2.0, by + bh / 2.0
                let paint =
                    if g.Radial then RadialGradient(vec2 cx cy, px (max bw bh / 2.0), stops, Pad)
                    else
                        let rad = g.Angle * System.Math.PI / 180.0
                        let dx, dy = cos rad * bw / 2.0, sin rad * bh / 2.0
                        LinearGradient(vec2 (cx - dx) (cy - dy), vec2 (cx + dx) (cy + dy), stops, Pad)
                [ Cmd.FillPath(rpath, paint, NonZero) ]
        let strokeCmds =
            match stroke with
            | Some c ->
                let style = { Width = px strokeWidth; Cap = Butt; Join = Miter; Paint = Solid (parseHexColor c); DashArray = []; DashOffset = px 0.0; MiterLimit = 4.0 }
                [ Cmd.StrokePath(rpath, style) ]
            | None -> []
        fillCmds @ strokeCmds

let private clearRendererFrameCommands state =
    state.LayoutCache.RendererFrameContentCommands <- None
    state.LayoutCache.RendererFrameSelectionCommands <- None
    state.LayoutCache.RendererFrameHoverCommands <- None
    state.LayoutCache.RendererFrameCommands <- None

let private rendererSelectionCommandsForState state layouts settings measureRange doc layoutLookup =
    match
        state.LayoutCache.RendererSelectionLayouts,
        state.LayoutCache.RendererSelectionValue,
        state.LayoutCache.RendererSelectionCommands
    with
    | Some cachedLayouts, cachedSelection, Some cachedCommands
        when obj.ReferenceEquals(cachedLayouts, layouts)
             && cachedSelection = doc.Selection ->
        cachedCommands
    | _ ->
        let commands = WordPaint.selectionCommandsWithRangeLookup settings measureRange doc layoutLookup
        state.LayoutCache.RendererSelectionLayouts <- Some layouts
        state.LayoutCache.RendererSelectionValue <- doc.Selection
        state.LayoutCache.RendererSelectionCommands <- Some commands
        commands

let private rendererContentCommandsCacheable (paintCommands: WordPaint.PaintCommand list) =
    let mutable remaining = paintCommands
    let mutable cacheable = true
    while cacheable && not remaining.IsEmpty do
        match remaining.Head with
        | WordPaint.DrawImage _ -> cacheable <- false
        | _ -> remaining <- remaining.Tail
    cacheable

let private rendererCommandsForContent resolveImage state contentCommands =
    match state.LayoutCache.RendererContentPaintCommands, state.LayoutCache.RendererContentCommands with
    | Some cachedPaintCommands, Some cachedRendererCommands when obj.ReferenceEquals(cachedPaintCommands, contentCommands) ->
        cachedRendererCommands
    | _ ->
        let rendererCommands = ResizeArray<Cmd>()
        for paintCommand in contentCommands do
            rendererCommands.AddRange(toRendererCommand resolveImage paintCommand)
        let rendered = resizeArrayToList rendererCommands
        if rendererContentCommandsCacheable contentCommands then
            state.LayoutCache.RendererContentPaintCommands <- Some contentCommands
            state.LayoutCache.RendererContentCommands <- Some rendered
        else
            state.LayoutCache.RendererContentPaintCommands <- None
            state.LayoutCache.RendererContentCommands <- None
            clearRendererFrameCommands state
        rendered

let private samePaintCommands left right =
    obj.ReferenceEquals(left, right) || left = right

let private rendererCommandsForFrame
    resolveImage
    state
    (contentRendererCommands: Cmd list)
    (selectionCommands: WordPaint.PaintCommand list)
    (hoverCommands: WordPaint.PaintCommand list) =
    match
        state.LayoutCache.RendererFrameContentCommands,
        state.LayoutCache.RendererFrameSelectionCommands,
        state.LayoutCache.RendererFrameHoverCommands,
        state.LayoutCache.RendererFrameCommands
    with
    | Some cachedContentCommands, Some cachedSelectionCommands, Some cachedHoverCommands, Some cachedFrameCommands
        when obj.ReferenceEquals(cachedContentCommands, contentRendererCommands)
             && samePaintCommands cachedSelectionCommands selectionCommands
             && samePaintCommands cachedHoverCommands hoverCommands ->
        cachedFrameCommands
    | _ ->
        let overlayCommands = ResizeArray<Cmd>()
        for paintCommand in selectionCommands do
            overlayCommands.AddRange(toRendererCommand resolveImage paintCommand)
        for paintCommand in hoverCommands do
            overlayCommands.AddRange(toRendererCommand resolveImage paintCommand)
        let rendered =
            let overlay = resizeArrayToList overlayCommands
            if contentRendererCommands.IsEmpty then
                overlay
            else
                Cmd.Group(Matrix3x2.identity, 1.0, contentRendererCommands) :: overlay
        state.LayoutCache.RendererFrameContentCommands <- Some contentRendererCommands
        state.LayoutCache.RendererFrameSelectionCommands <- Some selectionCommands
        state.LayoutCache.RendererFrameHoverCommands <- Some hoverCommands
        state.LayoutCache.RendererFrameCommands <- Some rendered
        rendered

let private supportsViewportContentCommands (doc: Document) =
    doc.Headers.IsEmpty
    && doc.Footers.IsEmpty
    && doc.HeaderParts.IsEmpty
    && doc.FooterParts.IsEmpty
    && doc.PageSetup.Columns.IsNone
    && (doc.Blocks |> List.forall (RedBat.Word.Document.tryGetParagraph >> Option.isSome))

let private supportsViewportContentCommandsCached (state: WordEditorState) (doc: Document) =
    match state.LayoutCache.ViewportSupportDocument, state.LayoutCache.ViewportSupportsContentCommands with
    | Some cachedDoc, Some cached when layoutInputsEqual cachedDoc doc ->
        cached
    | _ ->
        let supported = supportsViewportContentCommands doc
        state.LayoutCache.ViewportSupportDocument <- Some doc
        state.LayoutCache.ViewportSupportsContentCommands <- Some supported
        supported

let private viewportContentPaintCommands
    (settings: WordLayout.LayoutSettings)
    (measureRange: WordLayout.TextRangeMeasurer)
    (layouts: WordLayout.BlockLayout array)
    (scrollY: float)
    (viewportHeight: float) =
    let top = max 0.0 (scrollY - settings.LineHeight * 2.0)
    let bottom = scrollY + viewportHeight + settings.LineHeight * 2.0
    let commands = ResizeArray<WordPaint.PaintCommand>()
    let mutable lo = 0
    let mutable hi = layouts.Length
    while lo < hi do
        let mid = lo + (hi - lo) / 2
        let layout = layouts.[mid]
        if layout.Y + layout.Height < top then
            lo <- mid + 1
        else
            hi <- mid

    let mutable index = lo
    while index < layouts.Length do
        let layout = layouts.[index]
        if layout.Y <= bottom then
            commands.AddRange(WordPaint.paragraphCommands settings "Calibri" measureText measureRange layout)
        else
            index <- layouts.Length - 1
        index <- index + 1
    resizeArrayToList commands

let commands (width: float) (height: float) (resolveImage: string -> Framebuffer option) (hoveredImage: (BlockId * int) option) state (doc: Document) =
    let settings, cached = cachedLayouts width state doc
    let _, layoutLookup = cachedLayoutLookup width state doc
    let measureRange = WordLayout.rangeMeasurerFromTextMeasurer measureText
    let layouts, contentCommands, _ =
        WordPaint.collectDocumentCommandsWithLayouts settings "Calibri" measureText measureRange doc (Some cached)
    let contentRendererCommands = rendererCommandsForContent resolveImage state contentCommands
    let selectionCommands = rendererSelectionCommandsForState state cached settings measureRange doc layoutLookup
    let hoverCommands = WordPaint.hoverCommandsLookup settings layoutLookup hoveredImage

    if selectionCommands.IsEmpty && hoverCommands.IsEmpty then
        clearRendererFrameCommands state
        contentRendererCommands
    else
        rendererCommandsForFrame resolveImage state contentRendererCommands selectionCommands hoverCommands

let private commandsInViewport
    (width: float)
    (viewportHeight: float)
    (scrollY: float)
    (resolveImage: string -> Framebuffer option)
    (hoveredImage: (BlockId * int) option)
    state
    (doc: Document) =
    if not (supportsViewportContentCommandsCached state doc) then
        commands width viewportHeight resolveImage hoveredImage state doc
    else
        let settings, layoutArray = cachedLayoutArray width state doc
        let _, layoutLookup = cachedLayoutLookup width state doc
        let measureRange = WordLayout.rangeMeasurerFromTextMeasurer measureText
        let contentPaintCommands = viewportContentPaintCommands settings measureRange layoutArray scrollY viewportHeight
        let contentRendererCommands =
            let rendererCommands = ResizeArray<Cmd>()
            for paintCommand in contentPaintCommands do
                rendererCommands.AddRange(toRendererCommand resolveImage paintCommand)
            resizeArrayToList rendererCommands
        let cachedLayouts = state.LayoutCache.Layouts |> Option.defaultValue []
        let selectionCommands = rendererSelectionCommandsForState state cachedLayouts settings measureRange doc layoutLookup
        let hoverCommands = WordPaint.hoverCommandsLookup settings layoutLookup hoveredImage

        if selectionCommands.IsEmpty && hoverCommands.IsEmpty then
            contentRendererCommands
        else
            rendererCommandsForFrame resolveImage state contentRendererCommands selectionCommands hoverCommands

let widget width height resolveImage hoveredImage state doc =
    canvas width height (fun w h -> commands w h resolveImage hoveredImage state doc)

let private viewportWidget width contentHeight viewportHeight resolveImage hoveredImage state doc =
    canvas width contentHeight (fun w _ -> commandsInViewport w viewportHeight state.ScrollY.Value resolveImage hoveredImage state doc)

let editor width height doc =
    let state = createState doc
    widget width height (fun _ -> None) None state doc

let private interactiveWidgetInViewport width height viewportHeight resolveImage state =
    dynamic (fun () ->
        let doc = state.Document.Value
        updateCursorRect width state doc
        let hoveredImage = state.HoveredImage.Value
        let handleText = function
            | TextCommit text ->
                applyDocumentChange state (insertPlainTextWithPendingStyle state text)
            | CompositionStart | CompositionUpdate _ | CompositionEnd ->
                // Hosts own native IME/preedit UI; committed text arrives as TextCommit.
                true
            | TextKeyPress ev -> handleKey width height state ev

        let handlePointer = function
            | PointerDown (x, y) ->
                Focus.request state.FocusId
                match imageAt width x y state with
                | Some (blockId, offset, imageWidth, imageHeight, imageX, imageY) when isResizeHandle x y imageX imageY imageWidth imageHeight ->
                    state.ResizingImage.Value <- Some (blockId, offset, x, y, imageWidth, imageHeight, state.Document.Value)
                    state.HoveredImage.Value <- Some (blockId, offset)
                    Platform.requestRedraw()
                | Some (blockId, offset, _, _, _, _) ->
                    state.DraggingImage.Value <- Some (blockId, offset, x, y, state.Document.Value)
                    state.HoveredImage.Value <- Some (blockId, offset)
                    Platform.requestRedraw()
                | _ ->
                    let count = clickCount x y state
                    match positionAt width x y state with
                    | Some pos ->
                        if count >= 3 then
                            state.ParagraphDragAnchor.Value <- paragraphRangeAt pos state.Document.Value
                            state.WordDragAnchor.Value <- None
                            applyDocumentChangeNoHistory state (selectParagraphAt pos) |> ignore
                        elif count = 2 then
                            state.WordDragAnchor.Value <- wordRangeAt pos state.Document.Value
                            state.ParagraphDragAnchor.Value <- None
                            applyDocumentChangeNoHistory state (selectWordAt pos) |> ignore
                        else
                            state.WordDragAnchor.Value <- None
                            state.ParagraphDragAnchor.Value <- None
                            state.DragAnchor.Value <- Some pos
                            applyDocumentChangeNoHistory state (setCursor pos.BlockId pos.Offset) |> ignore
                    | None -> ()
            | PointerMove (x, y) ->
                if state.DragAnchor.Value.IsSome || state.WordDragAnchor.Value.IsSome || state.ParagraphDragAnchor.Value.IsSome || state.DraggingImage.Value.IsSome then
                    autoScrollForPointer viewportHeight height y state

                match state.ResizingImage.Value with
                | Some (blockId, offset, startX, startY, startWidth, startHeight, _) ->
                    let deltaX = x - startX
                    let deltaY = y - startY
                    let delta = if abs deltaX > abs deltaY then deltaX else deltaY
                    let newWidth = max 20.0 (startWidth + delta)
                    let newHeight = max 20.0 (startHeight + delta)
                    applyDocumentChangeNoHistory state (resizeImage blockId offset newWidth newHeight) |> ignore
                | None ->
                    match state.DraggingImage.Value with
                    | Some _ -> ()
                    | None ->
                        match state.ParagraphDragAnchor.Value, positionAt width x y state with
                        | Some (anchorStart, anchorEnd), Some focus ->
                            match paragraphRangeAt focus state.Document.Value with
                            | Some (focusStart, focusEnd) ->
                                let nextDoc =
                                    let normalized = getNormalizedRange { Anchor = anchorStart; Focus = focusEnd } state.Document.Value
                                    let startPosition, _, _ = normalized
                                    if startPosition.BlockId = anchorStart.BlockId && startPosition.Offset = anchorStart.Offset then
                                        setSelection anchorStart focusEnd state.Document.Value
                                    else
                                        setSelection anchorEnd focusStart state.Document.Value
                                applyDocumentChangeNoHistory state (fun _ -> nextDoc) |> ignore
                            | None -> ()
                        | None, _ ->
                            match state.WordDragAnchor.Value, positionAt width x y state with
                            | Some (anchorStart, anchorEnd), Some focus ->
                                match wordRangeAt focus state.Document.Value with
                                | Some (focusStart, focusEnd) ->
                                    let nextDoc =
                                        let normalized = getNormalizedRange { Anchor = anchorStart; Focus = focusEnd } state.Document.Value
                                        let startPosition, _, _ = normalized
                                        if startPosition.BlockId = anchorStart.BlockId && startPosition.Offset = anchorStart.Offset then
                                            setSelection anchorStart focusEnd state.Document.Value
                                        else
                                            setSelection anchorEnd focusStart state.Document.Value
                                    applyDocumentChangeNoHistory state (fun _ -> nextDoc) |> ignore
                                | None -> ()
                            | None, _ ->
                                match state.DragAnchor.Value, positionAt width x y state with
                                | Some anchor, Some focus ->
                                    applyDocumentChangeNoHistory state (setSelection anchor focus) |> ignore
                                | _ -> ()
                            | _, None -> ()
                        | _, None -> ()
                        let nextHover =
                            imageAt width x y state
                            |> Option.map (fun (blockId, offset, _, _, _, _) -> (blockId, offset))
                        if nextHover <> state.HoveredImage.Value then
                            state.HoveredImage.Value <- nextHover
                            Platform.requestRedraw()
            | PointerUp (x, y) ->
                match state.DraggingImage.Value with
                | Some (blockId, offset, startX, startY, originalDocument) ->
                    let dx = x - startX
                    let dy = y - startY
                    if abs dx > 3.0 || abs dy > 3.0 then
                        let target = imageMoveTarget width x y blockId offset state state.Document.Value
                        let nextDocument = moveImage blockId offset target.BlockId target.Offset state.Document.Value
                        if not (documentsEqualForChange nextDocument state.Document.Value) then
                            state.UndoStack.Value <- (originalDocument :: state.UndoStack.Value) |> trimHistory
                            state.RedoStack.Value <- []
                            state.Document.Value <- nextDocument
                            Platform.requestRedraw()
                | None -> ()

                match state.ResizingImage.Value with
                | Some (_, _, _, _, _, _, originalDocument) when not (documentsEqualForChange state.Document.Value originalDocument) ->
                    state.UndoStack.Value <- (originalDocument :: state.UndoStack.Value) |> trimHistory
                    state.RedoStack.Value <- []
                | _ -> ()
                state.DragAnchor.Value <- None
                state.WordDragAnchor.Value <- None
                state.ParagraphDragAnchor.Value <- None
                state.ResizingImage.Value <- None
                state.DraggingImage.Value <- None

        Focusable(state.FocusId, onKey (handleKey width height state) (onText handleText (pointerRegion handlePointer (viewportWidget width height viewportHeight resolveImage hoveredImage state doc)))))

let interactiveWidget width height resolveImage state =
    interactiveWidgetInViewport width height height resolveImage state

let interactiveEditor width height doc =
    interactiveWidget width height (fun _ -> None) (createState doc)

let scrollWidget width height resolveImage state =
    dynamic (fun () ->
        let contentHeight = max height (documentHeightCached width state state.Document.Value)
        scrollWith false true (signal 0.0) state.ScrollY (interactiveWidgetInViewport width contentHeight height resolveImage state))

let scrollEditor width height doc =
    scrollWidget width height (fun _ -> None) (createState doc)
