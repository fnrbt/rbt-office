module RedBat.Word.Pdf

open System
open System.Text
#if FABLE_COMPILER
open Fable.Core
#endif
open RedBat.Word.Document
open RedBat.Word.TextBuffer

module WordCanvasLayout = RedBat.Word.Layout
module WordModelLayout = RedBat.Word.WordLayout
module WordStyles = RedBat.Word.WordStyles
module WordPaint = RedBat.Word.Paint
module PdfDoc = FsPdf.Doc
module PdfDraw = FsPdf.Draw
module PdfColor = FsPdf.Color
module PdfFont = FsPdf.Font
module PdfImage = FsPdf.Image

let private textCacheLimit = 100000
let private pdfPointScale = 72.0 / 96.0

let private pdfScaleForPageWidth width =
    // Browser-imported Word documents use CSS pixels; native diagnostics use points.
    if width > 700.0 then pdfPointScale else 1.0

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
    ShapeText: (string -> TextStyle -> (Text.Engine.TextFont * Text.Engine.ShapedText) option) option
    Images: ImageResource list
}

type PaintPlan = {
    Settings: WordCanvasLayout.LayoutSettings
    Commands: WordPaint.PaintCommand list
    Pages: WordCanvasLayout.PageInfo list
    ImageSources: string list
    ImageUses: ImageUse list
}

type private TextContext = {
    FontSet: RedBat.Word.ShaperText.FontSet
    FontBytes: Map<string, byte[]>
    DefaultFamily: string
    mutable PdfFonts: Map<string, PdfFont.TrueTypeFont>
    GlyphText: System.Collections.Generic.Dictionary<string, System.Collections.Generic.Dictionary<int, string>>
    ShapeCache: System.Collections.Generic.Dictionary<string, (Text.Engine.TextFont * Text.Engine.ShapedText) option>
    PdfRunCache: System.Collections.Generic.Dictionary<string, (float * float * (int * int)[] * (int * string)[]) list>
    ExternalShapeText: (string -> TextStyle -> (Text.Engine.TextFont * Text.Engine.ShapedText) option) option
    mutable TextOpCount: int
    mutable TextShapeMs: float
    mutable TextRunMs: float
    mutable TextAllocMs: float
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
               Data = face.Data }: RedBat.Word.ShaperText.FontFace))
    {
        FontSet = RedBat.Word.ShaperText.createFontSet options.DefaultFontFamily shaperFaces
        FontBytes = options.Fonts |> List.map (fun face -> face.FaceName, face.Data) |> Map.ofList
        DefaultFamily = options.DefaultFontFamily
        PdfFonts = Map.empty
        GlyphText = System.Collections.Generic.Dictionary<string, System.Collections.Generic.Dictionary<int, string>>()
        ShapeCache = System.Collections.Generic.Dictionary<string, (Text.Engine.TextFont * Text.Engine.ShapedText) option>()
        PdfRunCache = System.Collections.Generic.Dictionary<string, (float * float * (int * int)[] * (int * string)[]) list>()
        ExternalShapeText = options.ShapeText
        TextOpCount = 0
        TextShapeMs = 0.0
        TextRunMs = 0.0
        TextAllocMs = 0.0
    }

let private fallbackFaceName (textContext: TextContext) (style: WordPaint.TextPaintStyle) =
    // Mirror WordModelLayout.Font.exactFace's naming scheme, but anchored on
    // the FontSet's DefaultFamily.  Used when the requested face (e.g.
    // "Cambria") isn't registered: the shaper falls back to the default
    // family internally, so we follow the same chain for the PDF font lookup
    // — keeps shape and embedded-font face in agreement and avoids dropping
    // into PDF's Helvetica simple path (which paints spaces at Helvetica
    // metrics while layout positioned glyphs from the fallback shape).
    let defaultFamily = textContext.DefaultFamily
    match style.Bold, style.Italic with
    | true, true -> defaultFamily + "-BoldItalic"
    | true, false -> defaultFamily + "-Bold"
    | false, true -> defaultFamily + "-Italic"
    | false, false -> defaultFamily

let private tryGetPdfFont (textContext: TextContext) faceName =
    match Map.tryFind faceName textContext.PdfFonts with
    | Some font -> Some font
    | None ->
        match Map.tryFind faceName textContext.FontBytes with
        | None -> None
        | Some bytes ->
            let font = PdfFont.parseTrueType bytes
            textContext.PdfFonts <- textContext.PdfFonts |> Map.add faceName font
            Some font

let private tryGetPdfFontWithFallback (textContext: TextContext) (style: WordPaint.TextPaintStyle) faceName =
    match tryGetPdfFont textContext faceName with
    | Some font -> Some (faceName, font)
    | None ->
        let fallback = fallbackFaceName textContext style
        if fallback = faceName then None
        else
            tryGetPdfFont textContext fallback
            |> Option.map (fun font -> fallback, font)

let private textStyleKey (style: TextStyle) =
    String.concat "|" [
        defaultArg style.FontFamily ""
        string style.FontSize
        if style.Bold then "b" else ""
        if style.Italic then "i" else ""
        if style.SmallCaps then "sc" else ""
        if style.AllCaps then "ac" else ""
        string style.CharacterSpacing
    ]

let private shapeTextCached (textContext: TextContext) text style =
    let key = textStyleKey style + "|" + text
    match textContext.ShapeCache.TryGetValue key with
    | true, value -> value
    | false, _ ->
        if textContext.ShapeCache.Count > textCacheLimit then textContext.ShapeCache.Clear()
        let value =
            match textContext.ExternalShapeText with
            | Some shapeText -> shapeText text style
            | None -> RedBat.Word.ShaperText.shapeText textContext.FontSet text style
        textContext.ShapeCache.[key] <- value
        value

let createMeasurer (options: RenderOptions) : WordCanvasLayout.TextMeasurer =
    match options.MeasureText with
    | Some measureText -> measureText
    | None ->
        let fontSet =
            options.Fonts
            |> List.map (fun face ->
                ({ Family = face.Family
                   Bold = face.Bold
                   Italic = face.Italic
                   Data = face.Data }: RedBat.Word.ShaperText.FontFace))
            |> RedBat.Word.ShaperText.createFontSet options.DefaultFontFamily
        let measure = RedBat.Word.ShaperText.createMeasurer fontSet
        let cache = System.Collections.Generic.Dictionary<string, float>()
        fun text style ->
            // Measure at the effective rendered size: superscript/subscript draw
            // at FontSize*0.7 (the render uses printFontSize), so the layout width
            // must match or the PDF horizontalScale stretches the glyphs.
            let style =
                match style.VertAlign with
                | Some "superscript" | Some "subscript" -> { style with FontSize = style.FontSize * 0.7 }
                | _ -> style
            let key =
                String.concat "|" [
                    defaultArg style.FontFamily ""
                    string style.FontSize
                    if style.Bold then "b" else ""
                    if style.Italic then "i" else ""
                    if style.SmallCaps then "sc" else ""
                    if style.AllCaps then "ac" else ""
                    string style.CharacterSpacing
                    text
                ]
            match cache.TryGetValue key with
            | true, value -> value
            | false, _ ->
                if cache.Count > textCacheLimit then cache.Clear()
                let value = measure text style
                cache.[key] <- value
                value

let private createContextMeasurer (textContext: TextContext) : WordCanvasLayout.TextMeasurer =
    let cache = System.Collections.Generic.Dictionary<string, float>()
    fun text style ->
        // Measure superscript/subscript at the effective rendered size
        // (FontSize*0.7, what printFontSize/embeddedTextOps draw) so the layout
        // width matches the shaped advance and horizontalScale stays 1.0 instead
        // of stretching the mark glyph.
        let style =
            match style.VertAlign with
            | Some "superscript" | Some "subscript" -> { style with FontSize = style.FontSize * 0.7 }
            | _ -> style
        let key = textStyleKey style + "|" + text
        match cache.TryGetValue key with
        | true, value -> value
        | false, _ ->
            if cache.Count > textCacheLimit then cache.Clear()
            let value =
                match shapeTextCached textContext text style with
                | Some (_, shapedText) -> shapedText.TotalAdvance + TextBuffer.characterSpacingAdvance text style
                | None -> 0.0
            cache.[key] <- value
            value

let private createContextRangeMeasurer (textContext: TextContext) : WordCanvasLayout.TextRangeMeasurer =
    // Match the browser renderer (Renderer.fs uses rangeMeasurerFromTextMeasurer)
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
        else measureText (text.Substring(startIndex, endIndex - startIndex)) style

let private controlWhitespaceAt (text: string) index =
    if index < 0 || index >= text.Length then false
    else
        match text.[index] with
        | '\t' | '\n' | '\r' -> true
        | _ -> false

let private clusterTextMappings (text: string) (glyphs: Text.Engine.PositionedGlyph[]) =
    let clusters =
        glyphs
        |> Array.map (fun glyph -> max 0 (min text.Length glyph.Cluster))
        |> Array.distinct
        |> Array.sort
    let clusterEnd cluster =
        clusters
        |> Array.tryFind (fun next -> next > cluster)
        |> Option.defaultValue text.Length
    let seen = System.Collections.Generic.HashSet<int>()
    [| for glyph in glyphs do
        let cluster = max 0 (min text.Length glyph.Cluster)
        if seen.Add cluster then
            let finish = clusterEnd cluster
            if finish > cluster && not (controlWhitespaceAt text cluster) then
                yield glyph.GlyphId, text.Substring(cluster, finish - cluster) |]

let private addGlyphTextMappings (textContext: TextContext) faceName (mappings: (int * string)[]) =
    if mappings.Length > 0 then
        if not (textContext.GlyphText.ContainsKey faceName) then
            textContext.GlyphText.[faceName] <- System.Collections.Generic.Dictionary<int, string>()
        let fontMappings = textContext.GlyphText.[faceName]
        for gid, text in mappings do
            match fontMappings.TryGetValue gid with
            | true, existing when existing.Length >= text.Length -> ()
            | _ -> fontMappings.[gid] <- text

let private rendererPdfFaceName (style: WordPaint.TextPaintStyle) =
    paintStyleToTextStyle style |> WordModelLayout.Font.exactFace

let private pdfColor (value: string) =
    let value =
        if String.IsNullOrWhiteSpace(value) then "#000000"
        else
            match value.Trim().ToLowerInvariant() with
            | "black" -> "#000000"
            | "blue" -> "#0000FF"
            | "cyan" -> "#00FFFF"
            | "green" -> "#00FF00"
            | "magenta" -> "#FF00FF"
            | "red" -> "#FF0000"
            | "yellow" -> "#FFFF00"
            | "white" -> "#FFFFFF"
            | "darkblue" -> "#000080"
            | "darkcyan" -> "#008080"
            | "darkgreen" -> "#008000"
            | "darkmagenta" -> "#800080"
            | "darkred" -> "#800000"
            | "darkyellow" -> "#808000"
            | "darkgray" -> "#808080"
            | "lightgray" -> "#C0C0C0"
            | _ -> value.Trim()
    let parseHexByte (text: string) =
        match Byte.TryParse(text, Globalization.NumberStyles.HexNumber, Globalization.CultureInfo.InvariantCulture) with
        | true, b -> float b / 255.0
        | false, _ -> 0.0
    if value.StartsWith("#") && value.Length >= 7 then
        PdfColor.rgb (parseHexByte value.[1..2]) (parseHexByte value.[3..4]) (parseHexByte value.[5..6])
    elif value.StartsWith("rgb") then
        let start = value.IndexOf('(')
        let finish = value.IndexOf(')')
        if start >= 0 && finish > start then
            let parts = value.Substring(start + 1, finish - start - 1).Split(',') |> Array.map (fun s -> s.Trim())
            if parts.Length >= 3 then
                let channel i =
                    match Double.TryParse(parts.[i], Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture) with
                    | true, n -> max 0.0 (min 1.0 (n / 255.0))
                    | false, _ -> 0.0
                PdfColor.rgb (channel 0) (channel 1) (channel 2)
            else PdfColor.black
        else PdfColor.black
    else PdfColor.black

let private pdfFontName (style: WordPaint.TextPaintStyle) =
    match style.Bold, style.Italic with
    | true, true -> "Helvetica-BoldOblique"
    | true, false -> "Helvetica-Bold"
    | false, true -> "Helvetica-Oblique"
    | false, false -> "Helvetica"

let private glyphKernPairs (font: PdfFont.TrueTypeFont) (fontSize: float) characterSpacing (glyphs: Text.Engine.PositionedGlyph[]) =
    glyphs
    |> Array.mapi (fun index glyph ->
        let baseWidthFu =
            if glyph.GlyphId >= 0 && glyph.GlyphId < font.GlyphWidths.Length then font.GlyphWidths.[glyph.GlyphId]
            else font.GlyphWidths.[0]
        let baseWidthPt = float baseWidthFu / float font.UnitsPerEm * fontSize
        let actualAdvancePt = glyph.XAdv + if index < glyphs.Length - 1 then characterSpacing else 0.0
        let diff = baseWidthPt - actualAdvancePt
        let kern = int (Math.Round(diff / fontSize * 1000.0))
        glyph.GlyphId, kern)

let private embeddedTextOps (textContext: TextContext) pageTop pageHeight text (style: WordPaint.TextPaintStyle) x baseline width =
    let yPdf y = pageHeight - (y - pageTop)
    let textStyle = paintStyleToTextStyle style
    let requestedFace = rendererPdfFaceName style
    let shapeStart = timingStart ()
    let pdfFontResolved = tryGetPdfFontWithFallback textContext style requestedFace
    let shaped = shapeTextCached textContext text textStyle
    if traceTimingEnabled then textContext.TextShapeMs <- textContext.TextShapeMs + (nowMs () - shapeStart)
    match pdfFontResolved, shaped with
    | Some (faceName, pdfFont), Some (_, shapedText) ->
        let spacedAdvance = shapedText.TotalAdvance + TextBuffer.characterSpacingAdvance text textStyle
        let horizontalScale = if width > 0.0 && spacedAdvance > 0.0 then width / spacedAdvance else 1.0
        let cacheKey = sprintf "%s|%s|%.4f|%s" faceName (textStyleKey textStyle) width text
        let runStart = timingStart ()
        let runs =
            match textContext.PdfRunCache.TryGetValue cacheKey with
            | true, value -> value
            | false, _ ->
                if textContext.PdfRunCache.Count > textCacheLimit then textContext.PdfRunCache.Clear()
                let value =
                    shapedText.Runs
                    |> List.choose (fun run ->
                        let glyphs = run.Glyphs |> Array.filter (fun glyph -> not (controlWhitespaceAt text glyph.Cluster))
                        if glyphs.Length = 0 then None
                        else Some (glyphs.[0].X, glyphs.[0].Y, glyphKernPairs pdfFont style.FontSize style.CharacterSpacing glyphs, clusterTextMappings text glyphs))
                textContext.PdfRunCache.[cacheKey] <- value
                value
        if traceTimingEnabled then textContext.TextRunMs <- textContext.TextRunMs + (nowMs () - runStart)
        let allocStart = timingStart ()
        let textRuns =
            runs
            |> List.collect (fun (runDx, runDy, pairs, mappings) ->
                addGlyphTextMappings textContext faceName mappings
                [ PdfDraw.BeginText
                  PdfDraw.SetFillColor (pdfColor style.Color)
                  PdfDraw.SetFont (faceName, style.FontSize)
                  PdfDraw.TextMatrix (horizontalScale, 0.0, 0.0, 1.0, x + runDx, yPdf (baseline - runDy))
                  PdfDraw.ShowTextHexWithKern pairs
                  PdfDraw.EndText ])
        if traceTimingEnabled then
            textContext.TextOpCount <- textContext.TextOpCount + 1
            textContext.TextAllocMs <- textContext.TextAllocMs + (nowMs () - allocStart)
        if textRuns.IsEmpty then [] else textRuns
    | _ ->
        [ PdfDraw.BeginText
          PdfDraw.SetFillColor (pdfColor style.Color)
          PdfDraw.SetFont (pdfFontName style, style.FontSize)
          PdfDraw.TextMatrix (1.0, 0.0, 0.0, 1.0, x, yPdf baseline)
          PdfDraw.ShowText text
          PdfDraw.EndText ]

let private commandBounds = function
    | WordPaint.DrawText (_, _, x, baseline, width) -> Some (x, baseline - 24.0, x + max 1.0 width, baseline + 4.0)
    | WordPaint.DrawUnderline (_, x, y, width) -> Some (x, y - 1.0, x + width, y + 1.0)
    | WordPaint.DrawImage (_, x, y, width, height, _, _, _, _) -> Some (x, y, x + width, y + height)
    | WordPaint.FillRect (_, x, y, width, height)
    | WordPaint.StrokeRect (_, _, x, y, width, height) -> Some (x, y, x + width, y + height)
    | WordPaint.StrokeLine (_, lineWidth, _, _, x1, y1, x2, y2) ->
        Some (min x1 x2 - lineWidth, min y1 y2 - lineWidth, max x1 x2 + lineWidth, max y1 y2 + lineWidth)

let private intersectsPage pageTop pageHeight command =
    match commandBounds command with
    | None -> false
    | Some (_, y1, _, y2) -> y2 >= pageTop && y1 <= pageTop + pageHeight

// Index of the page whose [Top, Top+Height) band contains y (clamped).
let private pageIndexForY (pages: WordCanvasLayout.PageInfo[]) y =
    let mutable idx = 0
    for p in 0 .. pages.Length - 1 do
        if y >= pages.[p].Top - 0.01 then idx <- p
    idx

let private commandPageRange (pages: WordCanvasLayout.PageInfo[]) command =
    match commandBounds command with
    | None -> None
    | Some (_, y1, _, y2) ->
        let firstPage = pageIndexForY pages (max 0.0 y1)
        let lastPage = pageIndexForY pages (max 0.0 y2)
        if lastPage < firstPage then None else Some (firstPage, lastPage)

let private bucketCommandsByPage (pages: WordCanvasLayout.PageInfo[]) commands =
    let buckets = Array.init pages.Length (fun _ -> ResizeArray<WordPaint.PaintCommand>())
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
    List.ofSeq merged

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

let private pdfOpsForCommand (textContext: TextContext) (imageNames: Map<string, string>) pageTop pageHeight =
    let yPdf y = pageHeight - (y - pageTop)
    function
    | WordPaint.DrawText (text, style, x, baseline, width) ->
        embeddedTextOps textContext pageTop pageHeight text style x baseline width
    | WordPaint.DrawUnderline (color, x, y, width) ->
        [ PdfDraw.SetStrokeColor (pdfColor color)
          PdfDraw.SetLineWidth 1.0
          PdfDraw.SetDash ([], 0.0)
          PdfDraw.MoveTo (x, yPdf y)
          PdfDraw.LineTo (x + width, yPdf y)
          PdfDraw.Stroke ]
    | WordPaint.FillRect (color, x, y, width, height) ->
        [ PdfDraw.SetFillColor (pdfColor color)
          PdfDraw.Rect (x, yPdf (y + height), width, height)
          PdfDraw.Fill ]
    | WordPaint.StrokeRect (color, lineWidth, x, y, width, height) ->
        [ PdfDraw.SetStrokeColor (pdfColor color)
          PdfDraw.SetLineWidth lineWidth
          PdfDraw.SetDash ([], 0.0)
          PdfDraw.Rect (x, yPdf (y + height), width, height)
          PdfDraw.Stroke ]
    | WordPaint.StrokeLine (color, lineWidth, dashStyle, dashOffset, x1, y1, x2, y2) ->
        [ PdfDraw.SetStrokeColor (pdfColor color)
          PdfDraw.SetLineWidth lineWidth
          PdfDraw.SetDash (dashPattern lineWidth dashStyle, dashOffset)
          PdfDraw.MoveTo (x1, yPdf y1)
          PdfDraw.LineTo (x2, yPdf y2)
          PdfDraw.Stroke
          PdfDraw.SetDash ([], 0.0) ]
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

let private collectPdfGlyphUsage (opsByPage: PdfDraw.DrawOp list list) =
    let usage = System.Collections.Generic.Dictionary<string, System.Collections.Generic.HashSet<int>>()
    let ensure name =
        if not (usage.ContainsKey name) then usage.[name] <- System.Collections.Generic.HashSet<int>()
    for ops in opsByPage do
        let mutable currentFont = ""
        for op in ops do
            match op with
            | PdfDraw.SetFont (name, _) ->
                currentFont <- name
                ensure name
            | PdfDraw.ShowTextHex hex when currentFont <> "" ->
                ensure currentFont
                let set = usage.[currentFont]
                let mutable i = 0
                while i + 4 <= hex.Length do
                    set.Add(Convert.ToInt32(hex.Substring(i, 4), 16)) |> ignore
                    i <- i + 4
            | PdfDraw.ShowTextHexWithKern pairs when currentFont <> "" ->
                ensure currentFont
                let set = usage.[currentFont]
                for gid, _ in pairs do set.Add(gid) |> ignore
            | _ -> ()
    usage |> Seq.map (fun kv -> kv.Key, Set.ofSeq kv.Value) |> Map.ofSeq

let private remapPdfGlyphOps (remaps: Map<string, Map<int, int>>) (ops: PdfDraw.DrawOp list) =
    let remapHex remap (hex: string) =
        let sb = StringBuilder(hex.Length)
        let mutable i = 0
        while i + 4 <= hex.Length do
            let oldGid = Convert.ToInt32(hex.Substring(i, 4), 16)
            let newGid = Map.tryFind oldGid remap |> Option.defaultValue 0
            sb.AppendFormat("{0:X4}", newGid) |> ignore
            i <- i + 4
        sb.ToString()

    let mutable currentRemap = Map.empty<int, int>
    ops
    |> List.map (fun op ->
        match op with
        | PdfDraw.SetFont (name, _) ->
            currentRemap <- Map.tryFind name remaps |> Option.defaultValue Map.empty
            op
        | PdfDraw.ShowTextHex hex when not (Map.isEmpty currentRemap) ->
            PdfDraw.ShowTextHex (remapHex currentRemap hex)
        | PdfDraw.ShowTextHexWithKern pairs when not (Map.isEmpty currentRemap) ->
            pairs
            |> Array.map (fun (gid, kern) -> Map.tryFind gid currentRemap |> Option.defaultValue 0, kern)
            |> PdfDraw.ShowTextHexWithKern
        | _ -> op)

let private scalePageOps scale ops =
    [ yield PdfDraw.SaveState
      yield PdfDraw.Transform (scale, 0.0, 0.0, scale, 0.0, 0.0)
      yield! ops
      yield PdfDraw.RestoreState ]

let collectPaintCommandsWithLayouts defaultFontFamily (measureText: WordCanvasLayout.TextMeasurer) (measureRange: WordCanvasLayout.TextRangeMeasurer) (doc: Document) (cachedLayouts: WordCanvasLayout.BlockLayout list option) =
    let lineHeight = 21.0
    let footerReserve =
        if not doc.Footers.IsEmpty || not doc.FooterParts.IsEmpty then doc.PageSetup.FooterDistance + lineHeight * 1.5
        else doc.PageSetup.MarginBottom
    let settings: WordCanvasLayout.LayoutSettings = {
        Width = doc.PageSetup.PageWidth
        PageHeight = doc.PageSetup.PageHeight
        FontSize = 14.0
        LineHeight = lineHeight
        PaddingLeft = doc.PageSetup.MarginLeft
        PaddingTop = WordModelLayout.Page.bodyTopWithRunningHeader doc
        MarginRight = doc.PageSetup.MarginRight
        MarginBottom = max doc.PageSetup.MarginBottom footerReserve
        InTableCell = false
        TableContinuationHeaderHeight = 0.0
        LongTrailingSpaceFitSafety = 0.25
    }
    let layoutStart = timingStart ()
    let layouts, paintCommands, pages = WordPaint.collectDocumentCommandsWithLayouts settings defaultFontFamily measureText measureRange doc cachedLayouts
    let layoutLabel = if cachedLayouts.IsSome then "collect:cachedLayouts:done" else "collect:layoutDocument:done"
    traceElapsed (sprintf "%s layouts=%d commands=%d pages=%d" layoutLabel layouts.Length paintCommands.Length pages.Length) layoutStart
    settings, paintCommands, pages

let collectPaintCommands defaultFontFamily (measureText: WordCanvasLayout.TextMeasurer) (measureRange: WordCanvasLayout.TextRangeMeasurer) (doc: Document) =
    collectPaintCommandsWithLayouts defaultFontFamily measureText measureRange doc None

let private imageSourcesFromCommands commands =
    commands
    |> List.choose (function
        | WordPaint.DrawImage (src, _, _, _, _, _, _, _, _) when not (String.IsNullOrWhiteSpace src) -> Some src
        | _ -> None)
    |> List.distinct

let private imageUsesFromCommands commands =
    commands
    |> List.choose (function
        | WordPaint.DrawImage (src, _, _, width, height, _, _, _, _) when not (String.IsNullOrWhiteSpace src) ->
            Some { Source = src; Width = max 1.0 width; Height = max 1.0 height }
        | _ -> None)

let private collectPaintPlanWithCachedLayouts (cachedLayouts: WordCanvasLayout.BlockLayout list option) (options: RenderOptions) (doc: Document) =
    let textContext = createTextContext options
    let measureText = createContextMeasurer textContext
    let measureRange = createContextRangeMeasurer textContext
    let settings, commands, pages = collectPaintCommandsWithLayouts options.DefaultFontFamily measureText measureRange doc cachedLayouts
    { Settings = settings
      Commands = commands
      Pages = pages
      ImageSources = imageSourcesFromCommands commands
      ImageUses = imageUsesFromCommands commands }

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
    let imageNames = addImages builder options.Images
    let commandBuckets = bucketCommandsByPage pages commands
    trace "toBytes:pageOps:start"
    let pageOps =
        [ for pageIndex in 0 .. pageCount - 1 do
            let info = pages.[pageIndex]
            let ops = ResizeArray<PdfDraw.DrawOp>()
            ops.Add(PdfDraw.SetFillColor PdfColor.white)
            ops.Add(PdfDraw.Rect (0.0, 0.0, info.Width, info.Height))
            ops.Add(PdfDraw.Fill)
            for command in commandBuckets.[pageIndex] do
                if intersectsPage info.Top info.Height command then
                    ops.AddRange(pdfOpsForCommand textContext imageNames info.Top info.Height command)
            yield List.ofSeq ops ]
    trace (sprintf "toBytes:pageOps:done pages=%d" pageOps.Length)
    if traceTimingEnabled then
        trace (sprintf "toBytes:textOps:summary commands=%d shapeMs=%.1f runMs=%.1f allocMs=%.1f runCache=%d shapeCache=%d" textContext.TextOpCount textContext.TextShapeMs textContext.TextRunMs textContext.TextAllocMs textContext.PdfRunCache.Count textContext.ShapeCache.Count)

    trace "toBytes:glyphUsage:start"
    let usage = collectPdfGlyphUsage pageOps
    trace (sprintf "toBytes:glyphUsage:done fonts=%d" (Map.count usage))
    let mutable remaps = Map.empty<string, Map<int, int>>
    trace "toBytes:subset:start"
    let usedEmbeddedFontNames =
        usage
        |> Map.toList
        |> List.choose (fun (fontName, gids) ->
            match Map.tryFind fontName textContext.FontBytes with
            | Some bytes when not (Set.isEmpty gids) ->
                let remap =
                    try
                        PdfDoc.loadFontBytesSubset builder fontName bytes gids
                    with _ ->
                        PdfDoc.loadFontBytes builder fontName bytes |> ignore
                        Map.empty
                remaps <- remaps |> Map.add fontName remap
                Some fontName
            | _ -> None)
    trace (sprintf "toBytes:subset:done fonts=%d" usedEmbeddedFontNames.Length)

    let pageFonts =
        [ yield! usedEmbeddedFontNames
          yield!
              usage
              |> Map.keys
              |> Seq.filter (fun name -> not (Map.containsKey name textContext.FontBytes)) ]
        |> List.distinct

    trace "toBytes:addPages:start"
    let usedImages = imageNames |> Map.toList |> List.map snd
    pageOps
    |> List.iteri (fun pageIndex ops ->
        let info = pages.[pageIndex]
        let pageScale = pdfScaleForPageWidth info.Width
        let pageSize = info.Width * pageScale, info.Height * pageScale
        PdfDoc.addPage builder pageSize pageFonts usedImages (ops |> remapPdfGlyphOps remaps |> scalePageOps pageScale) |> ignore)
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
