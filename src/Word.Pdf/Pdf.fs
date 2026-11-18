module Rbt.Word.Pdf

open System
open System.Text
#if FABLE_COMPILER
open Fable.Core
#endif
open Rbt.Word.Document
open Rbt.Word.TextBuffer

module WordCanvasLayout = Rbt.Word.Layout
module WordStyles = Rbt.Word.WordStyles
module WordPaint = Rbt.Word.Paint
module PdfDoc = Rbt.Pdf.Doc
module PdfDraw = Rbt.Pdf.Draw
module PdfColor = Rbt.Pdf.Color
module PdfFont = Rbt.Pdf.Font
module PdfImage = Rbt.Pdf.Image

let private textCacheLimit = 100000
let private pdfPointScale = 72.0 / 96.0

/// Unit of a paint plan's page geometry and drawing coordinates.
type PageUnit =
    /// CSS pixels at 96 DPI. Every Document is in this unit (both importers
    /// convert twips to px), so every plan collected from a document is too.
    | CssPixels
    /// PDF points, for plans built directly in points.
    | Points

// The unit comes from the plan, never from the page size: guessing "wide
// pages are pixels" exported A5 (559px wide) a third too large and would
// shrink a landscape Letter plan built in points (792pt wide).
let private pdfScaleForUnit unit =
    match unit with
    | CssPixels -> pdfPointScale
    | Points -> 1.0

#if FABLE_COMPILER
[<Emit("try { if (new URLSearchParams(window.location.search).has('rbPdfExportTrace') || window.localStorage.getItem('rb.word.pdfExportTrace') === '1') console.log(new Date().toISOString() + ' word-pdf:' + $0); } catch (_) {}")>]
let private trace (_label: string) : unit = jsNative

[<Emit("(function(){ try { return new URLSearchParams(window.location.search).has('rbPdfExportTrace') || window.localStorage.getItem('rb.word.pdfExportTrace') === '1'; } catch (_) { return false; } })()")>]
let private traceTimingEnabled : bool = jsNative

[<Emit("performance.now()")>]
let private nowMs () : float = jsNative
#else
let private trace (_label: string) : unit = ()
let private traceTimingEnabled = false

let private nowMs () =
    float (System.Diagnostics.Stopwatch.GetTimestamp()) * 1000.0 / float System.Diagnostics.Stopwatch.Frequency
#endif

let private timingStart () = if traceTimingEnabled then nowMs () else 0.0

let private traceElapsed label start =
    if traceTimingEnabled then trace (sprintf "%s %.1fms" label (nowMs () - start))

let private resizeArrayToList (items: ResizeArray<'T>) =
    let mutable result = []
    for i = items.Count - 1 downto 0 do
        result <- items.[i] :: result
    result

type FontFace = {
    FaceName: string
    Family: string
    Bold: bool
    Italic: bool
    Data: byte[]
}

type ImageResource = {
    Source: string
    Width: int
    Height: int
    ColorSpace: string
    BitsPerComponent: int
    Data: byte[]
    Filter: string option
}

type ImageUse = {
    Source: string
    Width: float
    Height: float
}

type RenderOptions = {
    DefaultFontFamily: string
    Fonts: FontFace list
    MeasureText: WordCanvasLayout.TextMeasurer option
    ShapeText: (string -> TextStyle -> (Rbt.Text.Engine.TextFont * Rbt.Text.Engine.ShapedText) option) option
    Images: ImageResource list
}

type PaintPlan = {
    Settings: WordCanvasLayout.LayoutSettings
    Commands: WordPaint.PaintCommand list
    Pages: WordCanvasLayout.PageInfo list
    ImageSources: string list
    ImageUses: ImageUse list
    PageUnit: PageUnit
}

[<Struct>]
type private TextStyleCacheKey = {
    FontFamily: string
    FontSize: float
    Bold: bool
    Italic: bool
    SmallCaps: bool
    AllCaps: bool
    CharacterSpacing: float
}

[<Struct>]
type private TextCacheKey = {
    Style: TextStyleCacheKey
    Text: string
}

[<Struct>]
type private PdfRunCacheKey = {
    FaceName: string
    Style: TextStyleCacheKey
    WidthKey: int64
    Text: string
}

[<Struct>]
type private RemappedPairCacheKey = {
    FontName: string
    Pairs: (int * int)[]
}

let private remappedPairCacheKeyComparer =
    { new System.Collections.Generic.IEqualityComparer<RemappedPairCacheKey> with
        member _.Equals(left, right) =
            left.FontName = right.FontName
            && obj.ReferenceEquals(left.Pairs, right.Pairs)

        member _.GetHashCode(key) =
            (hash key.FontName * 397)
            ^^^ System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(key.Pairs) }

type private TextContext = {
    FontSet: Rbt.Word.ShaperText.FontSet
    mutable FontBytes: Map<string, byte[]>
    FontNamesByBytes: System.Collections.Generic.Dictionary<byte[], string>
    PdfFontCache: System.Collections.Generic.Dictionary<string, PdfFont.TrueTypeFont>
    mutable PdfFonts: Map<string, PdfFont.TrueTypeFont>
    GlyphText: System.Collections.Generic.Dictionary<string, System.Collections.Generic.Dictionary<int, string>>
    ShapeCache: System.Collections.Generic.Dictionary<TextCacheKey, (Rbt.Text.Engine.TextFont * Rbt.Text.Engine.ShapedText) option>
    PdfRunCache: System.Collections.Generic.Dictionary<PdfRunCacheKey, (float * float * (int * int)[] * (int * string)[]) list>
    GlyphUsagePairCache: System.Collections.Generic.Dictionary<RemappedPairCacheKey, unit>
    ColorCache: System.Collections.Generic.Dictionary<string, PdfColor.Color>
    ExternalShapeText: (string -> TextStyle -> (Rbt.Text.Engine.TextFont * Rbt.Text.Engine.ShapedText) option) option
    mutable TextOpCount: int
    mutable TextShapeMs: float
    mutable TextRunMs: float
    mutable TextAllocMs: float
}

type private GlyphUsage = System.Collections.Generic.Dictionary<string, System.Collections.Generic.HashSet<int>>

let private textStyleCacheKey (style: TextStyle) = {
    FontFamily = defaultArg style.FontFamily ""
    FontSize = style.FontSize
    Bold = style.Bold
    Italic = style.Italic
    SmallCaps = style.SmallCaps
    AllCaps = style.AllCaps
    CharacterSpacing = style.CharacterSpacing
}

let private textCacheKey text style = {
    Style = textStyleCacheKey style
    Text = text
}

let private pdfWidthCacheKey width =
    int64 (Math.Round(width * 10000.0))

let private pdfRunCacheKey faceName text width style = {
    FaceName = faceName
    Style = textStyleCacheKey style
    WidthKey = pdfWidthCacheKey width
    Text = text
}

let private paintStyleToTextStyle (style: WordPaint.TextPaintStyle) : TextStyle =
    { TextBuffer.defaultStyle with
        Bold = style.Bold
        Italic = style.Italic
        Underline = style.Underline
        AllCaps = style.AllCaps
        SmallCaps = style.SmallCaps
        Strike = style.Strike
        FontSize = style.FontSize
        CharacterSpacing = style.CharacterSpacing
        FontFamily = style.FontFamily
        Color = Some style.Color
        Highlight = style.Highlight
        VertAlign = style.VertAlign }

let private createTextContext (options: RenderOptions) =
    let shaperFaces =
        options.Fonts
        |> List.map (fun face ->
            ({ Family = face.Family
               Bold = face.Bold
               Italic = face.Italic
               Data = face.Data }: Rbt.Word.ShaperText.FontFace))
    let fontBytes = options.Fonts |> List.map (fun face -> face.FaceName, face.Data) |> Map.ofList
    let bytesComparer =
        { new System.Collections.Generic.IEqualityComparer<byte[]> with
            member _.Equals(left, right) = obj.ReferenceEquals(left, right)
            member _.GetHashCode(bytes) = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(bytes) }
    let fontNamesByBytes = System.Collections.Generic.Dictionary<byte[], string>(bytesComparer)
    for KeyValue(name, bytes) in fontBytes do
        if not (fontNamesByBytes.ContainsKey bytes) then fontNamesByBytes.[bytes] <- name
    {
        FontSet = Rbt.Word.ShaperText.createFontSet options.DefaultFontFamily shaperFaces
        FontBytes = fontBytes
        FontNamesByBytes = fontNamesByBytes
        PdfFontCache = System.Collections.Generic.Dictionary<string, PdfFont.TrueTypeFont>()
        PdfFonts = Map.empty
        GlyphText = System.Collections.Generic.Dictionary<string, System.Collections.Generic.Dictionary<int, string>>()
        ShapeCache = System.Collections.Generic.Dictionary<TextCacheKey, (Rbt.Text.Engine.TextFont * Rbt.Text.Engine.ShapedText) option>()
        PdfRunCache = System.Collections.Generic.Dictionary<PdfRunCacheKey, (float * float * (int * int)[] * (int * string)[]) list>()
        GlyphUsagePairCache = System.Collections.Generic.Dictionary<RemappedPairCacheKey, unit>(remappedPairCacheKeyComparer)
        ColorCache = System.Collections.Generic.Dictionary<string, PdfColor.Color>()
        ExternalShapeText = options.ShapeText
        TextOpCount = 0
        TextShapeMs = 0.0
        TextRunMs = 0.0
        TextAllocMs = 0.0
    }
let private tryGetPdfFont (textContext: TextContext) faceName =
    match Map.tryFind faceName textContext.PdfFonts with
    | Some font -> Some font
    | None ->
        let resolved =
            lock textContext.PdfFontCache (fun () ->
                let mutable cachedFont = Unchecked.defaultof<PdfFont.TrueTypeFont>
                if textContext.PdfFontCache.TryGetValue(faceName, &cachedFont) then
                    Some cachedFont
                else
                    match Map.tryFind faceName textContext.FontBytes with
                    | None -> None
                    | Some bytes ->
                        let font = PdfFont.parseTrueType bytes
                        textContext.PdfFontCache.[faceName] <- font
                        Some font)
        match resolved with
        | None -> None
        | Some font ->
            textContext.PdfFonts <- textContext.PdfFonts |> Map.add faceName font
            Some font

// A shaper can return a face not listed in options.Fonts. Its parsed font retains
// the original bytes, so use those bytes both for PDF widths and embedding.
let private shapedFaceName (textContext: TextContext) (font: Rbt.Text.Engine.TextFont) =
    let bytes = font.ShaperFont.Data
    match Portable.Dict.tryFind bytes textContext.FontNamesByBytes with
    | ValueSome name -> name
    | ValueNone ->
        let mutable index = textContext.FontNamesByBytes.Count
        let mutable name = "RbtShapedFont" + string index
        while Map.containsKey name textContext.FontBytes do
            index <- index + 1
            name <- "RbtShapedFont" + string index
        textContext.FontBytes <- Map.add name bytes textContext.FontBytes
        textContext.FontNamesByBytes.[bytes] <- name
        name

let private shapeTextCached (textContext: TextContext) text style =
    let key = textCacheKey text style
    match Portable.Dict.tryFind key textContext.ShapeCache with
    | ValueSome value -> value
    | ValueNone ->
        if textContext.ShapeCache.Count > textCacheLimit then textContext.ShapeCache.Clear()
        let value =
            match textContext.ExternalShapeText with
            | Some shapeText -> shapeText text style
            | None -> Rbt.Word.ShaperText.shapeText textContext.FontSet text style
        textContext.ShapeCache.[key] <- value
        value

let private createContextMeasurer (textContext: TextContext) : WordCanvasLayout.TextMeasurer =
    let cache = System.Collections.Generic.Dictionary<TextCacheKey, float>()
    fun text style ->
        // Measure superscript/subscript at the effective rendered size
        // (FontSize*0.7, what printFontSize/embeddedTextOps draw) so the layout
        // width matches the shaped advance and horizontalScale stays 1.0 instead
        // of stretching the mark glyph.
        let style =
            match style.VertAlign with
            | Some "superscript" | Some "subscript" -> { style with FontSize = style.FontSize * 0.7 }
            | _ -> style
        let key = textCacheKey text style
        match Portable.Dict.tryFind key cache with
        | ValueSome value -> value
        | ValueNone ->
            if cache.Count > textCacheLimit then cache.Clear()
            let value =
                match shapeTextCached textContext text style with
                | Some (_, shapedText) -> shapedText.TotalAdvance + TextBuffer.characterSpacingAdvance text style
                | None -> 0.0
            cache.[key] <- value
            value

let createMeasurer (options: RenderOptions) : WordCanvasLayout.TextMeasurer =
    match options.MeasureText with
    | Some measureText -> measureText
    | None -> createTextContext options |> createContextMeasurer

let private createContextRangeMeasurer (textContext: TextContext) : WordCanvasLayout.TextRangeMeasurer =
    // Match the browser renderer (Rbt.Renderer.fs uses rangeMeasurerFromTextMeasurer)
    // and our own PDF text painter: measure substrings by shaping them in
    // isolation, not by indexing into the full text's shape.
    //
    // The PDF render path (embeddedTextOps) always shapes the segment text
    // standalone and uses that TotalAdvance to size the text matrix. If
    // measureRange instead returned in-string advances (via Engine.measureRange
    // on a ShapedText of the full run text), Layout would size a segment with
    // kerning context, the PDF renderer would shape the same substring without
    // it, and the horizontalScale (= layoutWidth / standaloneShape.TotalAdvance)
    // would squeeze or stretch the glyphs to fill the layout-assigned width.
    // Visible drift is small for word-boundary breaks (no cross-character kern
    // at spaces) but real for mid-word splits.  Aligning Layout's measurement
    // to the standalone-substring shape keeps horizontalScale = 1.0 on every
    // segment and removes the artifact.
    let measureText = createContextMeasurer textContext
    fun text style start length ->
        let startIndex = max 0 (min text.Length start)
        let endIndex = max startIndex (min text.Length (start + length))
        if endIndex = startIndex then 0.0
        elif startIndex = 0 && endIndex = text.Length then measureText text style
        else measureText (text.Substring(startIndex, endIndex - startIndex)) style

let private controlWhitespaceAt (text: string) index =
    if index < 0 || index >= text.Length then false
    else
        match text.[index] with
        | '\t' | '\n' | '\r' -> true
        | _ -> false

let private clusterTextMappings (text: string) (glyphs: Rbt.Text.Engine.PositionedGlyph[]) =
    let clusterSet = System.Collections.Generic.HashSet<int>()
    let clusters = System.Collections.Generic.List<int>()
    for glyph in glyphs do
        let cluster = max 0 (min text.Length glyph.Cluster)
        if not (controlWhitespaceAt text cluster) && clusterSet.Add cluster then
            clusters.Add cluster
    clusters.Sort()
    let clusterEnds = System.Collections.Generic.Dictionary<int, int>(clusters.Count)
    for i = 0 to clusters.Count - 1 do
        let finish = if i + 1 < clusters.Count then clusters.[i + 1] else text.Length
        clusterEnds.[clusters.[i]] <- finish
    let seen = System.Collections.Generic.HashSet<int>()
    let mappings = System.Collections.Generic.List<int * string>(clusters.Count)
    for glyph in glyphs do
        let cluster = max 0 (min text.Length glyph.Cluster)
        if not (controlWhitespaceAt text cluster) && seen.Add cluster then
            match Portable.Dict.tryFind cluster clusterEnds with
            | ValueSome finish when finish > cluster ->
                mappings.Add(glyph.GlyphId, text.Substring(cluster, finish - cluster))
            | _ -> ()
    mappings.ToArray()

let private addGlyphTextMappings (textContext: TextContext) faceName (mappings: (int * string)[]) =
    if mappings.Length > 0 then
        if not (textContext.GlyphText.ContainsKey faceName) then
            textContext.GlyphText.[faceName] <- System.Collections.Generic.Dictionary<int, string>()
        let fontMappings = textContext.GlyphText.[faceName]
        for gid, text in mappings do
            match Portable.Dict.tryFind gid fontMappings with
            | ValueSome existing when existing.Length >= text.Length -> ()
            | _ -> fontMappings.[gid] <- text

let private equalsOrdinalIgnoreCase (left: string) (right: string) =
    String.Equals(left, right, StringComparison.OrdinalIgnoreCase)

let private containsOrdinalIgnoreCase (value: string) (needle: string) =
    value.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0

let private hexValue (ch: char) =
    if ch >= '0' && ch <= '9' then int ch - int '0'
    elif ch >= 'A' && ch <= 'F' then int ch - int 'A' + 10
    elif ch >= 'a' && ch <= 'f' then int ch - int 'a' + 10
    else -1

let private parseHexByteAt (text: string) index =
    if index + 1 >= text.Length then 0.0
    else
        let hi = hexValue text.[index]
        let lo = hexValue text.[index + 1]
        if hi < 0 || lo < 0 then 0.0 else float ((hi <<< 4) ||| lo) / 255.0

let private parseHex4At (text: string) index =
    let mutable value = 0
    let mutable valid = true
    let mutable offset = 0
    while valid && offset < 4 do
        let digit = hexValue text.[index + offset]
        if digit < 0 then valid <- false
        else value <- (value <<< 4) ||| digit
        offset <- offset + 1
    if valid then value else Convert.ToInt32(text.Substring(index, 4), 16)

let private hexDigits = "0123456789ABCDEF".ToCharArray()

let private appendHex4 (sb: StringBuilder) value =
    if value >= 0 && value <= 0xFFFF then
        sb.Append(hexDigits.[(value >>> 12) &&& 0xF]) |> ignore
        sb.Append(hexDigits.[(value >>> 8) &&& 0xF]) |> ignore
        sb.Append(hexDigits.[(value >>> 4) &&& 0xF]) |> ignore
        sb.Append(hexDigits.[value &&& 0xF]) |> ignore
    else
        sb.Append(value.ToString("X")) |> ignore

let private skipWhitespace (text: string) finish index =
    let mutable i = index
    while i < finish && Char.IsWhiteSpace text.[i] do
        i <- i + 1
    i

let private parseCssNumberAt (text: string) start finish =
    let mutable i = skipWhitespace text finish start
    let mutable sign = 1.0
    if i < finish && text.[i] = '-' then
        sign <- -1.0
        i <- i + 1
    elif i < finish && text.[i] = '+' then
        i <- i + 1
    let mutable whole = 0.0
    let mutable fraction = 0.0
    let mutable scale = 1.0
    let mutable sawDigit = false
    while i < finish && text.[i] >= '0' && text.[i] <= '9' do
        sawDigit <- true
        whole <- whole * 10.0 + float (int text.[i] - int '0')
        i <- i + 1
    if i < finish && text.[i] = '.' then
        i <- i + 1
        while i < finish && text.[i] >= '0' && text.[i] <= '9' do
            sawDigit <- true
            scale <- scale * 10.0
            fraction <- fraction + float (int text.[i] - int '0') / scale
            i <- i + 1
    struct (sign * (whole + fraction), i, sawDigit)

let private tryRgbColor (value: string) =
    let start = value.IndexOf('(')
    let finish = value.IndexOf(')')
    if start < 0 || finish <= start then
        None
    else
        let mutable i = start + 1
        let readChannel requireComma =
            if requireComma then
                i <- skipWhitespace value finish i
                if i < finish && value.[i] = ',' then
                    i <- i + 1
                    true
                else
                    false
            else
                true
        if not (readChannel false) then
            None
        else
            let struct (r, nextR, okR) = parseCssNumberAt value i finish
            i <- nextR
            if not okR || not (readChannel true) then
                None
            else
                let struct (g, nextG, okG) = parseCssNumberAt value i finish
                i <- nextG
                if not okG || not (readChannel true) then
                    None
                else
                    let struct (b, nextB, okB) = parseCssNumberAt value i finish
                    i <- skipWhitespace value finish nextB
                    if okB && i = finish then
                        Some (PdfColor.rgb (max 0.0 (min 1.0 (r / 255.0))) (max 0.0 (min 1.0 (g / 255.0))) (max 0.0 (min 1.0 (b / 255.0))))
                    else
                        None

let private pdfColor (value: string) =
    let value =
        if String.IsNullOrWhiteSpace(value) then "#000000"
        else
            let trimmed = value.Trim()
            if equalsOrdinalIgnoreCase trimmed "black" then "#000000"
            elif equalsOrdinalIgnoreCase trimmed "blue" then "#0000FF"
            elif equalsOrdinalIgnoreCase trimmed "cyan" then "#00FFFF"
            elif equalsOrdinalIgnoreCase trimmed "green" then "#00FF00"
            elif equalsOrdinalIgnoreCase trimmed "magenta" then "#FF00FF"
            elif equalsOrdinalIgnoreCase trimmed "red" then "#FF0000"
            elif equalsOrdinalIgnoreCase trimmed "yellow" then "#FFFF00"
            elif equalsOrdinalIgnoreCase trimmed "white" then "#FFFFFF"
            elif equalsOrdinalIgnoreCase trimmed "darkblue" then "#000080"
            elif equalsOrdinalIgnoreCase trimmed "darkcyan" then "#008080"
            elif equalsOrdinalIgnoreCase trimmed "darkgreen" then "#008000"
            elif equalsOrdinalIgnoreCase trimmed "darkmagenta" then "#800080"
            elif equalsOrdinalIgnoreCase trimmed "darkred" then "#800000"
            elif equalsOrdinalIgnoreCase trimmed "darkyellow" then "#808000"
            elif equalsOrdinalIgnoreCase trimmed "darkgray" then "#808080"
            elif equalsOrdinalIgnoreCase trimmed "lightgray" then "#C0C0C0"
            else trimmed
    if value.StartsWith("#") && value.Length >= 7 then
        PdfColor.rgb (parseHexByteAt value 1) (parseHexByteAt value 3) (parseHexByteAt value 5)
    elif value.StartsWith("rgb", StringComparison.OrdinalIgnoreCase) then
        tryRgbColor value |> Option.defaultValue PdfColor.black
    else PdfColor.black

let private cachedPdfColor (textContext: TextContext) (value: string) =
    let key = if isNull value then "" else value
    let mutable color = Unchecked.defaultof<PdfColor.Color>
    if textContext.ColorCache.TryGetValue(key, &color) then
        color
    else
        color <- pdfColor value
        textContext.ColorCache.[key] <- color
        color

/// Build PDF path-construction ops for a shape path. `yPdf` flips to PDF space.
/// Quadratics are promoted to cubics (PDF has no quad operator).
let private pdfShapeOps (yPdf: float -> float) (segs: PathSeg list) =
    let ops = ResizeArray<PdfDraw.DrawOp>()
    let mutable cx = 0.0
    let mutable cy = 0.0
    for seg in segs do
        match seg with
        | PMove(x, y) -> ops.Add(PdfDraw.MoveTo(x, yPdf y)); cx <- x; cy <- y
        | PLine(x, y) -> ops.Add(PdfDraw.LineTo(x, yPdf y)); cx <- x; cy <- y
        | PCubic(a, b, c, d, x, y) -> ops.Add(PdfDraw.CurveTo(a, yPdf b, c, yPdf d, x, yPdf y)); cx <- x; cy <- y
        | PQuad(qx, qy, x, y) ->
            let c1x = cx + 2.0 / 3.0 * (qx - cx)
            let c1y = cy + 2.0 / 3.0 * (qy - cy)
            let c2x = x + 2.0 / 3.0 * (qx - x)
            let c2y = y + 2.0 / 3.0 * (qy - y)
            ops.Add(PdfDraw.CurveTo(c1x, yPdf c1y, c2x, yPdf c2y, x, yPdf y)); cx <- x; cy <- y
        | PClose -> ops.Add(PdfDraw.ClosePath)
    resizeArrayToList ops

/// FsPdf has no axial-shading op exposed here, so a gradient is approximated by
/// the stop nearest the midpoint (a solid). True gradients are a PDF follow-up.
let private gradientSolid (g: GradientSpec) =
    g.Stops |> List.sortBy (fun s -> abs (s.Offset - 0.5)) |> List.tryHead |> Option.map (fun s -> s.Color)

let private pdfFontName (style: WordPaint.TextPaintStyle) =
    match style.Bold, style.Italic with
    | true, true -> "Helvetica-BoldOblique"
    | true, false -> "Helvetica-Bold"
    | false, true -> "Helvetica-Oblique"
    | false, false -> "Helvetica"

let private glyphKernPairs (text: string) (font: PdfFont.TrueTypeFont) (fontSize: float) characterSpacing (glyphs: Rbt.Text.Engine.PositionedGlyph[]) =
    let mutable count = 0
    let mutable firstX = 0.0
    let mutable firstY = 0.0
    for glyph in glyphs do
        if not (controlWhitespaceAt text glyph.Cluster) then
            if count = 0 then
                firstX <- glyph.X
                firstY <- glyph.Y
            count <- count + 1
    if count = 0 then None else
    let pairs = Array.zeroCreate<int * int> count
    let mutable writeIndex = 0
    for glyph in glyphs do
        if not (controlWhitespaceAt text glyph.Cluster) then
            let baseWidthFu =
                if glyph.GlyphId >= 0 && glyph.GlyphId < font.GlyphWidths.Length then font.GlyphWidths.[glyph.GlyphId]
                else font.GlyphWidths.[0]
            let baseWidthPt = float baseWidthFu / float font.UnitsPerEm * fontSize
            let actualAdvancePt = glyph.XAdv + if writeIndex < count - 1 then characterSpacing else 0.0
            let diff = baseWidthPt - actualAdvancePt
            let kern = int (Math.Round(diff / fontSize * 1000.0))
            pairs.[writeIndex] <- glyph.GlyphId, kern
            writeIndex <- writeIndex + 1
    Some (firstX, firstY, pairs)

let private recordFontUse (usage: GlyphUsage) name =
    if not (usage.ContainsKey name) then
        usage.[name] <- System.Collections.Generic.HashSet<int>()

let private recordGlyphPairs (textContext: TextContext) (usage: GlyphUsage) name (pairs: (int * int)[]) =
    let key = { FontName = name; Pairs = pairs }
    if not (textContext.GlyphUsagePairCache.ContainsKey key) then
        textContext.GlyphUsagePairCache.[key] <- ()
        recordFontUse usage name
        let set = usage.[name]
        for gid, _ in pairs do
            set.Add gid |> ignore

let private addEmbeddedTextOps (ops: ResizeArray<PdfDraw.DrawOp>) (textContext: TextContext) (usage: GlyphUsage) pageTop pageHeight text (style: WordPaint.TextPaintStyle) x baseline width =
    let yPdf y = pageHeight - (y - pageTop)
    let textStyle = paintStyleToTextStyle style
    let shapeStart = timingStart ()
    let shaped = shapeTextCached textContext text textStyle
    let resolved =
        shaped |> Option.bind (fun (font, shapedText) ->
            let faceName = shapedFaceName textContext font
            tryGetPdfFont textContext faceName
            |> Option.map (fun pdfFont -> faceName, pdfFont, shapedText))
    if traceTimingEnabled then textContext.TextShapeMs <- textContext.TextShapeMs + (nowMs () - shapeStart)
    match resolved with
    | Some (faceName, pdfFont, shapedText) ->
        let spacedAdvance = shapedText.TotalAdvance + TextBuffer.characterSpacingAdvance text textStyle
        let horizontalScale = if width > 0.0 && spacedAdvance > 0.0 then width / spacedAdvance else 1.0
        let cacheKey = pdfRunCacheKey faceName text width textStyle
        let runStart = timingStart ()
        let runs =
            match Portable.Dict.tryFind cacheKey textContext.PdfRunCache with
            | ValueSome value -> value
            | ValueNone ->
                if textContext.PdfRunCache.Count > textCacheLimit then textContext.PdfRunCache.Clear()
                let value =
                    shapedText.Runs
                    |> List.choose (fun run ->
                        match glyphKernPairs text pdfFont style.FontSize style.CharacterSpacing run.Glyphs with
                        | Some (firstX, firstY, pairs) ->
                            Some (firstX, firstY, pairs, clusterTextMappings text run.Glyphs)
                        | None -> None)
                textContext.PdfRunCache.[cacheKey] <- value
                value
        if traceTimingEnabled then textContext.TextRunMs <- textContext.TextRunMs + (nowMs () - runStart)
        let allocStart = timingStart ()
        for runDx, runDy, pairs, mappings in runs do
            addGlyphTextMappings textContext faceName mappings
            recordGlyphPairs textContext usage faceName pairs
            ops.Add PdfDraw.BeginText
            ops.Add(PdfDraw.SetFillColor (cachedPdfColor textContext style.Color))
            ops.Add(PdfDraw.SetFont (faceName, style.FontSize))
            ops.Add(PdfDraw.TextMatrix (horizontalScale, 0.0, 0.0, 1.0, x + runDx, yPdf (baseline - runDy)))
            ops.Add(PdfDraw.ShowTextHexWithKern pairs)
            ops.Add PdfDraw.EndText
        if traceTimingEnabled then
            textContext.TextOpCount <- textContext.TextOpCount + 1
            textContext.TextAllocMs <- textContext.TextAllocMs + (nowMs () - allocStart)
    | _ ->
        recordFontUse usage (pdfFontName style)
        ops.Add PdfDraw.BeginText
        ops.Add(PdfDraw.SetFillColor (cachedPdfColor textContext style.Color))
        ops.Add(PdfDraw.SetFont (pdfFontName style, style.FontSize))
        ops.Add(PdfDraw.TextMatrix (1.0, 0.0, 0.0, 1.0, x, yPdf baseline))
        ops.Add(PdfDraw.ShowText text)
        ops.Add PdfDraw.EndText

let private embeddedTextOps (textContext: TextContext) (usage: GlyphUsage) pageTop pageHeight text (style: WordPaint.TextPaintStyle) x baseline width =
    let ops = ResizeArray<PdfDraw.DrawOp>(6)
    addEmbeddedTextOps ops textContext usage pageTop pageHeight text style x baseline width
    resizeArrayToList ops

let private commandBounds = function
    | WordPaint.DrawText (_, _, x, baseline, width) -> Some (x, baseline - 24.0, x + max 1.0 width, baseline + 4.0)
    | WordPaint.DrawUnderline (_, _, x, y, width) -> Some (x, y - 2.0, x + width, y + 3.0)
    | WordPaint.DrawImage (_, x, y, width, height, _, _, _, _) -> Some (x, y, x + width, y + height)
    | WordPaint.FillRect (_, x, y, width, height)
    | WordPaint.StrokeRect (_, _, x, y, width, height) -> Some (x, y, x + width, y + height)
    | WordPaint.StrokeLine (_, lineWidth, _, _, x1, y1, x2, y2) ->
        Some (min x1 x2 - lineWidth, min y1 y2 - lineWidth, max x1 x2 + lineWidth, max y1 y2 + lineWidth)
    | WordPaint.DrawShape (_, _, sw, (x, y, w, h), _) -> Some (x - sw, y - sw, x + w + sw, y + h + sw)
    | WordPaint.DrawGroup (_, _, _, _, _, (x, y, w, h), _) -> Some (x, y, x + w, y + h)
    | WordPaint.DrawMirrored (_, _, _, _, (x, y, w, h), _) -> Some (x, y, x + w, y + h)

let private intersectsPage pageTop pageHeight command =
    match commandBounds command with
    | None -> false
    | Some (_, y1, _, y2) -> y2 >= pageTop && y1 <= pageTop + pageHeight

// Index of the page whose [Top, Top+Height) band contains y (clamped).
let private pageIndexForY (pages: WordCanvasLayout.PageInfo[]) y =
    if pages.Length = 0 then 0 else
    let target = y + 0.01
    let mutable lo = 0
    let mutable hi = pages.Length
    while lo < hi do
        let mid = lo + (hi - lo) / 2
        if pages.[mid].Top <= target then
            lo <- mid + 1
        else
            hi <- mid
    max 0 (lo - 1)

let private commandPageRange (pages: WordCanvasLayout.PageInfo[]) command =
    match commandBounds command with
    | None -> None
    | Some (_, y1, _, y2) ->
        let firstPage = pageIndexForY pages (max 0.0 y1)
        let lastPage = pageIndexForY pages (max 0.0 y2)
        if lastPage < firstPage then None else Some (firstPage, lastPage)

let private bucketCommandsByPage (pages: WordCanvasLayout.PageInfo[]) (commands: WordPaint.PaintCommand list) =
    let bucketCapacity =
        if pages.Length = 0 then 0
        else max 4 (commands.Length / pages.Length + 1)
    let buckets = Array.init pages.Length (fun _ -> ResizeArray<WordPaint.PaintCommand>(bucketCapacity))
    for command in commands do
        match commandPageRange pages command with
        | None -> ()
        | Some (firstPage, lastPage) ->
            for pageIndex in firstPage .. lastPage do
                buckets.[pageIndex].Add(command)
    buckets

let private nearlyEqual a b = abs (a - b) < 0.01

let private mergePaintCommands commands =
    let merged = ResizeArray<WordPaint.PaintCommand>()
    for command in commands do
        if merged.Count = 0 then
            merged.Add(command)
        else
            let lastIndex = merged.Count - 1
            match merged.[lastIndex], command with
            | WordPaint.DrawText (priorText, priorStyle, priorX, priorBaseline, priorWidth), WordPaint.DrawText (text, style, x, baseline, width)
                when priorStyle = style && nearlyEqual priorBaseline baseline && nearlyEqual (priorX + priorWidth) x ->
                merged.[lastIndex] <- WordPaint.DrawText (priorText + text, priorStyle, priorX, priorBaseline, priorWidth + width)
            | _ -> merged.Add(command)
    resizeArrayToList merged

let private dashPattern lineWidth = function
    | 1 -> [lineWidth; lineWidth]
    | 2 -> [lineWidth * 3.0; lineWidth * 1.5]
    | 3 -> [lineWidth * 3.0; lineWidth * 1.5; lineWidth; lineWidth * 1.5]
    | _ -> []

let private cropValue (value: string option) =
    match value with
    | Some text ->
        match Double.TryParse(text, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture) with
        | true, parsed -> max 0.0 (min 100000.0 parsed)
        | false, _ -> 0.0
    | None -> 0.0

let rec private pdfOpsForCommand (textContext: TextContext) (usage: GlyphUsage) (imageNames: Map<string, string>) pageTop pageHeight =
    let yPdf y = pageHeight - (y - pageTop)
    function
    | WordPaint.DrawText (text, style, x, baseline, width) ->
        embeddedTextOps textContext usage pageTop pageHeight text style x baseline width
    | WordPaint.DrawUnderline (color, lineStyle, x, y, width) ->
        let heavy = containsOrdinalIgnoreCase lineStyle "heavy" || equalsOrdinalIgnoreCase lineStyle "thick"
        let flat dy = [ PdfDraw.MoveTo (x, yPdf (y + dy)); PdfDraw.LineTo (x + width, yPdf (y + dy)); PdfDraw.Stroke ]
        let dashed pattern = PdfDraw.SetDash (pattern, 0.0) :: flat 0.0 @ [ PdfDraw.SetDash ([], 0.0) ]
        let wavy dy =
            let ops = ResizeArray<PdfDraw.DrawOp>()
            ops.Add(PdfDraw.MoveTo (x, yPdf (y + dy)))
            let mutable cx = x
            let mutable up = true
            while cx < x + width do
                let nx = min (x + width) (cx + 2.0)
                ops.Add(PdfDraw.LineTo (nx, yPdf (y + dy + (if up then -1.3 else 1.3))))
                cx <- nx
                up <- not up
            ops.Add(PdfDraw.Stroke)
            resizeArrayToList ops
        let body =
            if containsOrdinalIgnoreCase lineStyle "wav" then wavy 0.0 @ (if containsOrdinalIgnoreCase lineStyle "double" then wavy 2.6 else [])
            elif containsOrdinalIgnoreCase lineStyle "dotdotdash" || containsOrdinalIgnoreCase lineStyle "dashdotdot" then dashed [4.0;2.0;1.0;2.0;1.0;2.0]
            elif containsOrdinalIgnoreCase lineStyle "dotdash" || containsOrdinalIgnoreCase lineStyle "dashdot" then dashed [4.0;2.0;1.0;2.0]
            elif containsOrdinalIgnoreCase lineStyle "dash" then dashed [4.0;2.0]
            elif containsOrdinalIgnoreCase lineStyle "dot" then dashed [1.0;2.0]
            elif equalsOrdinalIgnoreCase lineStyle "double" then flat 0.0 @ flat 2.0
            else flat 0.0
        [ PdfDraw.SetStrokeColor (cachedPdfColor textContext color)
          PdfDraw.SetLineWidth (if heavy then 2.0 else 1.0)
          PdfDraw.SetDash ([], 0.0) ] @ body
    | WordPaint.FillRect (color, x, y, width, height) ->
        [ PdfDraw.SetFillColor (cachedPdfColor textContext color)
          PdfDraw.Rect (x, yPdf (y + height), width, height)
          PdfDraw.Fill ]
    | WordPaint.StrokeRect (color, lineWidth, x, y, width, height) ->
        [ PdfDraw.SetStrokeColor (cachedPdfColor textContext color)
          PdfDraw.SetLineWidth lineWidth
          PdfDraw.SetDash ([], 0.0)
          PdfDraw.Rect (x, yPdf (y + height), width, height)
          PdfDraw.Stroke ]
    | WordPaint.StrokeLine (color, lineWidth, dashStyle, dashOffset, x1, y1, x2, y2) ->
        [ PdfDraw.SetStrokeColor (cachedPdfColor textContext color)
          PdfDraw.SetLineWidth lineWidth
          PdfDraw.SetDash (dashPattern lineWidth dashStyle, dashOffset)
          PdfDraw.MoveTo (x1, yPdf y1)
          PdfDraw.LineTo (x2, yPdf y2)
          PdfDraw.Stroke
          PdfDraw.SetDash ([], 0.0) ]
    | WordPaint.DrawGroup (clip, rotation, centerX, centerY, opacity, _, commands) ->
        // A text box or shape with its text: clipped, turned about its centre (the page's y runs up), translucent.
        let ops = ResizeArray<PdfDraw.DrawOp>()
        ops.Add PdfDraw.SaveState
        if opacity < 0.999 then
            ops.Add (PdfDraw.SetFillAlpha (max 0.0 opacity))
            ops.Add (PdfDraw.SetStrokeAlpha (max 0.0 opacity))
        if abs rotation >= 0.01 then
            let angle = -rotation * Math.PI / 180.0
            let c, s = cos angle, sin angle
            let cx, cy = centerX, yPdf centerY
            ops.Add (PdfDraw.Transform (c, s, -s, c, cx - cx * c + cy * s, cy - cx * s - cy * c))
        match clip with
        | Some (x, y, width, height) ->
            ops.Add (PdfDraw.Rect (x, yPdf (y + height), width, height))
            ops.Add PdfDraw.ClipPath
            ops.Add PdfDraw.EndPath
        | None -> ()
        for command in commands do
            ops.AddRange (pdfOpsForCommand textContext usage imageNames pageTop pageHeight command)
        ops.Add PdfDraw.RestoreState
        List.ofSeq ops
    | WordPaint.DrawMirrored (flipH, flipV, centerX, centerY, _, commands) ->
        // A flipped picture: mirrored about its middle (x -> 2cx - x, y -> 2cy - y on the page).
        let sx, sy = (if flipH then -1.0 else 1.0), (if flipV then -1.0 else 1.0)
        let cx, cy = centerX, yPdf centerY
        let ops = ResizeArray<PdfDraw.DrawOp>()
        ops.Add PdfDraw.SaveState
        ops.Add (PdfDraw.Transform (sx, 0.0, 0.0, sy, cx - sx * cx, cy - sy * cy))
        for command in commands do
            ops.AddRange (pdfOpsForCommand textContext usage imageNames pageTop pageHeight command)
        ops.Add PdfDraw.RestoreState
        List.ofSeq ops
    | WordPaint.DrawShape (fill, stroke, strokeWidth, _, segs) ->
        let fillColor =
            match fill with
            | WordPaint.FillNone -> None
            | WordPaint.FillSolid c -> Some c
            | WordPaint.FillGradient g -> gradientSolid g
        let colorOps =
            (fillColor |> Option.map (fun c -> [ PdfDraw.SetFillColor (cachedPdfColor textContext c) ]) |> Option.defaultValue [])
            @ (stroke |> Option.map (fun c -> [ PdfDraw.SetStrokeColor (cachedPdfColor textContext c); PdfDraw.SetLineWidth strokeWidth; PdfDraw.SetDash ([], 0.0) ]) |> Option.defaultValue [])
        let paintOp =
            match fillColor.IsSome, stroke.IsSome with
            | true, true -> [ PdfDraw.FillAndStroke ]
            | true, false -> [ PdfDraw.Fill ]
            | false, true -> [ PdfDraw.Stroke ]
            | false, false -> [ PdfDraw.EndPath ]
        colorOps @ pdfShapeOps yPdf segs @ paintOp
    | WordPaint.DrawImage (src, x, y, width, height, cropLeft, cropTop, cropRight, cropBottom) ->
        match Map.tryFind src imageNames with
        | None -> []
        | Some name ->
            let left = cropValue cropLeft
            let top = cropValue cropTop
            let right = cropValue cropRight
            let bottom = cropValue cropBottom
            if left + right >= 100000.0 || top + bottom >= 100000.0 then []
            elif left = 0.0 && top = 0.0 && right = 0.0 && bottom = 0.0 then
                [ PdfDraw.SaveState
                  PdfDraw.Transform (width, 0.0, 0.0, height, x, yPdf (y + height))
                  PdfDraw.DrawXObject name
                  PdfDraw.RestoreState ]
            else
                let fullWidth = width * 100000.0 / (100000.0 - left - right)
                let fullHeight = height * 100000.0 / (100000.0 - top - bottom)
                let fullX = x - fullWidth * left / 100000.0
                let fullY = yPdf (y + height) - fullHeight * bottom / 100000.0
                [ PdfDraw.SaveState
                  PdfDraw.Rect (x, yPdf (y + height), width, height)
                  PdfDraw.ClipPath
                  PdfDraw.EndPath
                  PdfDraw.Transform (fullWidth, 0.0, 0.0, fullHeight, fullX, fullY)
                  PdfDraw.DrawXObject name
                  PdfDraw.RestoreState ]

let private addImages (builder: PdfDoc.DocBuilder) (images: ImageResource list) =
    images
    |> List.mapi (fun index image ->
        let name = sprintf "Im%d" (index + 1)
        let info: PdfImage.ImageInfo = {
            Width = max 1 image.Width
            Height = max 1 image.Height
            ColorSpace = if String.IsNullOrWhiteSpace image.ColorSpace then "DeviceRGB" else image.ColorSpace
            BitsPerComponent = if image.BitsPerComponent <= 0 then 8 else image.BitsPerComponent
            Data = image.Data
            Filter = image.Filter
        }
        let imageRef = PdfImage.toXObject builder.Doc info
        builder.Images <- builder.Images |> Map.add name imageRef
        image.Source, name)
    |> Map.ofList

let private glyphUsageToMap (usage: GlyphUsage) =
    let mutable result = Map.empty
    for KeyValue(name, gids) in usage do
        let mutable gidSet = Set.empty
        for gid in gids do
            gidSet <- Set.add gid gidSet
        result <- Map.add name gidSet result
    result

let private remapAndScalePageOps
    (remappedPairCache: System.Collections.Generic.Dictionary<RemappedPairCacheKey, (int * int)[]>)
    (remaps: Map<string, Map<int, int>>)
    scale
    (ops: PdfDraw.DrawOp list) =
    let remapHex remap (hex: string) =
        let sb = StringBuilder(hex.Length)
        let mutable i = 0
        while i + 4 <= hex.Length do
            let oldGid = parseHex4At hex i
            let newGid =
                match Map.tryFind oldGid remap with
                | Some gid -> gid
                | None -> 0
            appendHex4 sb newGid
            i <- i + 4
        sb.ToString()

    let remapPairs remap (pairs: (int * int)[]) =
        let mapped = Array.zeroCreate<int * int> pairs.Length
        for i = 0 to pairs.Length - 1 do
            let gid, kern = pairs.[i]
            let mappedGid =
                match Map.tryFind gid remap with
                | Some value -> value
                | None -> 0
            mapped.[i] <- mappedGid, kern
        mapped

    let mappedOps = ResizeArray<PdfDraw.DrawOp>(List.length ops + 2)
    mappedOps.Add PdfDraw.SaveState
    mappedOps.Add(PdfDraw.Transform (scale, 0.0, 0.0, scale, 0.0, 0.0))
    let mutable currentFontName = ""
    let mutable currentRemap = Map.empty<int, int>
    let mutable hasRemap = false
    for op in ops do
        match op with
        | PdfDraw.SetFont (name, _) ->
            currentFontName <- name
            currentRemap <- Map.tryFind name remaps |> Option.defaultValue Map.empty
            hasRemap <- not (Map.isEmpty currentRemap)
            mappedOps.Add op
        | PdfDraw.ShowTextHex hex when hasRemap ->
            mappedOps.Add(PdfDraw.ShowTextHex (remapHex currentRemap hex))
        | PdfDraw.ShowTextHexWithKern pairs when hasRemap ->
            let key = { FontName = currentFontName; Pairs = pairs }
            let mutable mappedPairs = Unchecked.defaultof<(int * int)[]>
            if not (remappedPairCache.TryGetValue(key, &mappedPairs)) then
                mappedPairs <- remapPairs currentRemap pairs
                remappedPairCache.[key] <- mappedPairs
            mappedOps.Add(PdfDraw.ShowTextHexWithKern mappedPairs)
        | _ ->
            mappedOps.Add op
    mappedOps.Add PdfDraw.RestoreState
    resizeArrayToList mappedOps

let collectPaintCommandsWithLayouts defaultFontFamily (measureText: WordCanvasLayout.TextMeasurer) (measureRange: WordCanvasLayout.TextRangeMeasurer) (doc: Document) (cachedLayouts: WordCanvasLayout.BlockLayout list option) =
    let settings = WordCanvasLayout.documentSettings 14.0 21.0 0.25 doc
    let layoutStart = timingStart ()
    let layouts, paintCommands, pages = WordPaint.collectDocumentCommandsWithLayouts settings defaultFontFamily measureText measureRange doc cachedLayouts
    let layoutLabel = if cachedLayouts.IsSome then "collect:cachedLayouts:done" else "collect:layoutDocument:done"
    traceElapsed (sprintf "%s layouts=%d commands=%d pages=%d" layoutLabel layouts.Length paintCommands.Length pages.Length) layoutStart
    settings, paintCommands, pages

let collectPaintCommands defaultFontFamily (measureText: WordCanvasLayout.TextMeasurer) (measureRange: WordCanvasLayout.TextRangeMeasurer) (doc: Document) =
    collectPaintCommandsWithLayouts defaultFontFamily measureText measureRange doc None

let private imageSourcesFromCommands commands =
    let sources = ResizeArray<string>()
    let seen = System.Collections.Generic.HashSet<string>()
    for command in commands do
        match command with
        | WordPaint.DrawImage (src, _, _, _, _, _, _, _, _) when not (String.IsNullOrWhiteSpace src) ->
            if seen.Add src then sources.Add src
        | _ -> ()
    resizeArrayToList sources

let private imageUsesFromCommands commands =
    commands
    |> List.choose (function
        | WordPaint.DrawImage (src, _, _, width, height, _, _, _, _) when not (String.IsNullOrWhiteSpace src) ->
            Some { Source = src; Width = max 1.0 width; Height = max 1.0 height }
        | _ -> None)

let private usedImageSourcesForPlan planSources commands =
    let used = System.Collections.Generic.HashSet<string>()
    let add source =
        if not (String.IsNullOrWhiteSpace source) then
            used.Add source |> ignore
    for source in planSources do
        add source
    for command in commands do
        match command with
        | WordPaint.DrawImage (src, _, _, _, _, _, _, _, _) -> add src
        | _ -> ()
    used

let private usedImagesForPlan (usedSources: System.Collections.Generic.HashSet<string>) (images: ImageResource list) =
    let seen = System.Collections.Generic.HashSet<string>()
    let mutable kept = []
    for image in List.rev images do
        if not (String.IsNullOrWhiteSpace image.Source)
           && usedSources.Contains image.Source
           && seen.Add image.Source then
            kept <- image :: kept
    kept

let private collectPaintPlanWithCachedLayouts (cachedLayouts: WordCanvasLayout.BlockLayout list option) (options: RenderOptions) (doc: Document) =
    let measureText, measureRange =
        match options.MeasureText with
        | Some measureText ->
            measureText, WordCanvasLayout.rangeMeasurerFromTextMeasurer measureText
        | None ->
            let textContext = createTextContext options
            createContextMeasurer textContext, createContextRangeMeasurer textContext
    let settings, commands, pages = collectPaintCommandsWithLayouts options.DefaultFontFamily measureText measureRange doc cachedLayouts
    { Settings = settings
      Commands = commands
      Pages = pages
      ImageSources = imageSourcesFromCommands (WordPaint.flattenedCommands commands)
      ImageUses = imageUsesFromCommands (WordPaint.flattenedCommands commands)
      PageUnit = CssPixels }

let collectPaintPlan (options: RenderOptions) (doc: Document) =
    collectPaintPlanWithCachedLayouts None options doc

let collectPaintPlanWithLayouts (options: RenderOptions) (doc: Document) (layouts: WordCanvasLayout.BlockLayout list) =
    collectPaintPlanWithCachedLayouts (Some layouts) options doc

let collectImageSourcesWithLayouts (options: RenderOptions) (doc: Document) (layouts: WordCanvasLayout.BlockLayout list) =
    (collectPaintPlanWithLayouts options doc layouts).ImageSources

let collectImageSources (options: RenderOptions) (doc: Document) =
    (collectPaintPlan options doc).ImageSources

let toBytesWithPaintPlan (options: RenderOptions) (plan: PaintPlan) =
    trace "toBytes:createTextContext:start"
    let textContext = createTextContext options
    trace "toBytes:createTextContext:done"
    let settings = plan.Settings
    let commands = plan.Commands
    let pages = plan.Pages |> Array.ofList
    let pageCount = pages.Length
    trace (sprintf "toBytes:paintPlan commands=%d pages=%d" commands.Length pageCount)
    let builder = PdfDoc.create()
    let usedImageSources = usedImageSourcesForPlan plan.ImageSources (WordPaint.flattenedCommands commands)
    let images = usedImagesForPlan usedImageSources options.Images
    let imageNames = addImages builder images
    let commandBuckets = bucketCommandsByPage pages commands
    let glyphUsage = GlyphUsage()
    trace "toBytes:pageOps:start"
    let pageOps =
        [ for pageIndex in 0 .. pageCount - 1 do
            let info = pages.[pageIndex]
            let commandsOnPage = commandBuckets.[pageIndex]
            let ops = ResizeArray<PdfDraw.DrawOp>(max 16 (3 + commandsOnPage.Count * 8))
            ops.Add(PdfDraw.SetFillColor PdfColor.white)
            ops.Add(PdfDraw.Rect (0.0, 0.0, info.Width, info.Height))
            ops.Add(PdfDraw.Fill)
            for command in commandsOnPage do
                if intersectsPage info.Top info.Height command then
                    match command with
                    | WordPaint.DrawText (text, style, x, baseline, width) ->
                        addEmbeddedTextOps ops textContext glyphUsage info.Top info.Height text style x baseline width
                    | _ ->
                        ops.AddRange(pdfOpsForCommand textContext glyphUsage imageNames info.Top info.Height command)
            yield resizeArrayToList ops ]
    trace (sprintf "toBytes:pageOps:done pages=%d" pageOps.Length)
    if traceTimingEnabled then
        trace (sprintf "toBytes:textOps:summary commands=%d shapeMs=%.1f runMs=%.1f allocMs=%.1f runCache=%d shapeCache=%d" textContext.TextOpCount textContext.TextShapeMs textContext.TextRunMs textContext.TextAllocMs textContext.PdfRunCache.Count textContext.ShapeCache.Count)

    trace "toBytes:glyphUsage:start"
    let usage = glyphUsageToMap glyphUsage
    trace (sprintf "toBytes:glyphUsage:done fonts=%d" (Map.count usage))
    let mutable remaps = Map.empty<string, Map<int, int>>
    trace "toBytes:subset:start"
    let usedEmbeddedFontNames = ResizeArray<string>()
    for KeyValue(fontName, gids) in usage do
        match Map.tryFind fontName textContext.FontBytes with
        | Some bytes when not (Set.isEmpty gids) ->
            let remap =
                try
                    PdfDoc.loadFontBytesSubset builder fontName bytes gids
                with _ ->
                    PdfDoc.loadFontBytes builder fontName bytes |> ignore
                    Map.empty
            remaps <- remaps |> Map.add fontName remap
            usedEmbeddedFontNames.Add fontName
        | _ -> ()
    trace (sprintf "toBytes:subset:done fonts=%d" usedEmbeddedFontNames.Count)

    let pageFontBuffer = ResizeArray<string>(Map.count usage)
    pageFontBuffer.AddRange usedEmbeddedFontNames
    for KeyValue(fontName, _) in usage do
        if not (Map.containsKey fontName textContext.FontBytes) then
            pageFontBuffer.Add fontName
    let pageFonts = resizeArrayToList pageFontBuffer

    trace "toBytes:addPages:start"
    let usedImageBuffer = ResizeArray<string>(Map.count imageNames)
    for KeyValue(_, name) in imageNames do
        usedImageBuffer.Add name
    let usedImages = resizeArrayToList usedImageBuffer
    let remappedPairCache = System.Collections.Generic.Dictionary<RemappedPairCacheKey, (int * int)[]>(remappedPairCacheKeyComparer)
    pageOps
    |> List.iteri (fun pageIndex ops ->
        let info = pages.[pageIndex]
        let pageScale = pdfScaleForUnit plan.PageUnit
        let pageSize = info.Width * pageScale, info.Height * pageScale
        PdfDoc.addPage builder pageSize pageFonts usedImages (remapAndScalePageOps remappedPairCache remaps pageScale ops) |> ignore)
    trace "toBytes:addPages:done"

    trace "toBytes:serialize:start"
    let bytes = PdfDoc.toBytes builder
    trace (sprintf "toBytes:serialize:done bytes=%d" bytes.Length)
    bytes

let private toBytesWithCachedLayouts (cachedLayouts: WordCanvasLayout.BlockLayout list option) (options: RenderOptions) (doc: Document) =
    trace "toBytes:createPaintPlan:start"
    let plan = collectPaintPlanWithCachedLayouts cachedLayouts options doc
    trace (sprintf "toBytes:createPaintPlan:done commands=%d pages=%d" plan.Commands.Length plan.Pages.Length)
    toBytesWithPaintPlan options plan

let toBytes (options: RenderOptions) (doc: Document) =
    toBytesWithCachedLayouts None options doc

let toBytesWithLayouts (options: RenderOptions) (doc: Document) (layouts: WordCanvasLayout.BlockLayout list) =
    toBytesWithCachedLayouts (Some layouts) options doc
