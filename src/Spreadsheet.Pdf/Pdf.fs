module RedBat.Spreadsheet.Pdf.Render

open System
open RedBat.Spreadsheet.Model
open RedBat.Spreadsheet.Layout
open RedBat.Spreadsheet.Paint

module PdfDoc = FsPdf.Doc
module PdfDraw = FsPdf.Draw
module PdfColor = FsPdf.Color
module PdfFont = FsPdf.Font

/// One loadable TTF face the runner passes to us. Family is the abstract
/// family name from the cell ("Calibri", "Arial", ...); Bold/Italic mark the
/// variant.
type FontFace = {
    FaceName: string
    Family: string
    Bold: bool
    Italic: bool
    Data: byte[]
}

type RenderOptions = {
    /// Default family to use when a cell didn't request one.
    DefaultFontFamily: string
    Fonts: FontFace list
    /// Fallback faces for codepoints the primary font doesn't cover. Mirrors
    /// the renderer's GDI font-link chain: Yu Gothic (Japanese) → MS Gothic →
    /// Microsoft YaHei (Simplified Chinese) → JhengHei (Traditional) → Malgun
    /// (Korean). Order matters: first face whose cmap covers the codepoint
    /// wins. The PDF emitter picks bold vs regular based on the run's Bold
    /// flag; italic is ignored (these families don't ship italic variants).
    CjkFallback: FontFace list
    /// Page size in CSS pixels (we convert to points internally).
    PageWidth: float
    PageHeight: float
}

/// Convert "#RRGGBB" / "RRGGBB" to a PDF RGB color in 0..1.
let private parseHexColor (color: string) : PdfColor.Color =
    let s = if color.StartsWith "#" then color.Substring 1 else color
    let clean = if s.Length = 8 then s.Substring 2 else s
    if clean.Length <> 6 then PdfColor.black
    else
        let byteOf (h: string) =
            try float (Convert.ToByte(h, 16)) / 255.0 with _ -> 0.0
        PdfColor.rgb (byteOf (clean.Substring(0, 2))) (byteOf (clean.Substring(2, 2))) (byteOf (clean.Substring(4, 2)))

/// CSS pixels → PDF points. Excel widths flow in CSS px (1 px = 72/96 pt).
let private pxToPt = 72.0 / 96.0

let private px (v: float) = v * pxToPt

/// Pick the loaded TTF for the cell's (family, bold, italic) triple. Falls
/// back to the family's regular weight, then to the default family's regular.
let private pickFace (options: RenderOptions) (familyName: string option) (bold: bool) (italic: bool) : FontFace option =
    let family = familyName |> Option.defaultValue options.DefaultFontFamily
    let exact = options.Fonts |> List.tryFind (fun f -> f.Family = family && f.Bold = bold && f.Italic = italic)
    let weighted = options.Fonts |> List.tryFind (fun f -> f.Family = family && f.Bold = bold && not f.Italic)
    let regular = options.Fonts |> List.tryFind (fun f -> f.Family = family && not f.Bold && not f.Italic)
    let fallback = options.Fonts |> List.tryFind (fun f -> f.Family = options.DefaultFontFamily && not f.Bold && not f.Italic)
    exact
    |> Option.orElseWith (fun () -> weighted)
    |> Option.orElseWith (fun () -> regular)
    |> Option.orElseWith (fun () -> fallback)
    |> Option.orElseWith (fun () -> options.Fonts |> List.tryHead)

/// Walk `chain` and return the first face whose loaded cmap covers `cp`.
/// `faceByName` carries the parsed TrueTypeFont for each loaded face.
let private pickCjkFallback
    (chain: FontFace list)
    (faceByName: Map<string, PdfFont.TrueTypeFont>)
    (bold: bool)
    (cp: int)
    : (FontFace * PdfFont.TrueTypeFont) option =
    chain
    |> List.filter (fun f -> f.Bold = bold || not (chain |> List.exists (fun g -> g.Bold = bold)))
    |> List.tryPick (fun f ->
        match Map.tryFind f.FaceName faceByName with
        | Some ttf when ttf.Cmap.ContainsKey cp -> Some (f, ttf)
        | _ -> None)

/// Split `text` into contiguous (font, substring) segments where each
/// substring's codepoints are covered by `font`'s cmap. Falls back through
/// `cjk` for missing glyphs; uses `primary` for anything `primary` itself
/// covers (and as a last resort when no fallback covers a codepoint).
/// Mirrors SpreadsheetGrid.toRendererCommands so PDF and PNG outputs draw
/// the same glyph stream for mixed-script text.
let private segmentByFont
    (primary: FontFace * PdfFont.TrueTypeFont)
    (cjk: FontFace list)
    (faceByName: Map<string, PdfFont.TrueTypeFont>)
    (bold: bool)
    (text: string)
    : ((FontFace * PdfFont.TrueTypeFont) * string) list =
    let primaryTtf = snd primary
    let fontFor cp =
        if primaryTtf.Cmap.ContainsKey cp then primary
        else
            match pickCjkFallback cjk faceByName bold cp with
            | Some hit -> hit
            | None -> primary
    let result = System.Collections.Generic.List<(FontFace * PdfFont.TrueTypeFont) * string>()
    let buf = System.Text.StringBuilder()
    let mutable currentFace : (FontFace * PdfFont.TrueTypeFont) voption = ValueNone
    let mutable i = 0
    while i < text.Length do
        let ch = text.[i]
        let cp, width =
            if System.Char.IsHighSurrogate ch && i + 1 < text.Length && System.Char.IsLowSurrogate text.[i + 1] then
                System.Char.ConvertToUtf32(ch, text.[i + 1]), 2
            else int ch, 1
        let face = fontFor cp
        match currentFace with
        | ValueNone -> currentFace <- ValueSome face
        | ValueSome cf when System.Object.ReferenceEquals(snd cf, snd face) -> ()
        | ValueSome cf ->
            result.Add(cf, buf.ToString())
            buf.Clear() |> ignore
            currentFace <- ValueSome face
        for k in 0 .. width - 1 do buf.Append(text.[i + k]) |> ignore
        i <- i + width
    match currentFace with
    | ValueSome f when buf.Length > 0 -> result.Add(f, buf.ToString())
    | _ -> ()
    List.ofSeq result

/// Re-encode a substring as hex glyph IDs, applying a per-face subset
/// remap when present. With an empty remap (face was full-embedded), the
/// output matches PdfFont.toGlyphHex. With a remap, original GIDs are
/// translated to the post-subset numbering; missing entries collapse to
/// GID 0 (.notdef), which prints as a missing-glyph box.
let private toRemappedHex (ttf: PdfFont.TrueTypeFont) (remap: Map<int, int>) (text: string) =
    let sb = System.Text.StringBuilder()
    let isEmpty = Map.isEmpty remap
    for ch in text do
        let origGid = ttf.Cmap.TryFind (int ch) |> Option.defaultValue 0
        let gid =
            if isEmpty then origGid
            else remap.TryFind origGid |> Option.defaultValue 0
        sb.AppendFormat("{0:X4}", gid) |> ignore
    sb.ToString()

/// Convert one Paint item into a list of PDF draw ops. Coordinates flip from
/// our top-left CSS-pixel space into PDF's bottom-left point space.
let private translateItem (options: RenderOptions) (faceByName: Map<string, PdfFont.TrueTypeFont>) (faceResource: Map<string, string>) (faceRemap: Map<string, Map<int, int>>) (pageHeightPt: float) (item: PaintItem) : PdfDraw.DrawOp list =
    let yPdf y = pageHeightPt - px y
    match item with
    | FillRect (rect, color) ->
        [ PdfDraw.SetFillColor (parseHexColor color)
          PdfDraw.Rect (px rect.X, yPdf (rect.Y + rect.Height), px rect.Width, px rect.Height)
          PdfDraw.Fill ]
    | StrokeRect (rect, color, width) ->
        [ PdfDraw.SetStrokeColor (parseHexColor color)
          PdfDraw.SetLineWidth (px width)
          PdfDraw.SetDash ([], 0.0)
          PdfDraw.Rect (px rect.X, yPdf (rect.Y + rect.Height), px rect.Width, px rect.Height)
          PdfDraw.Stroke ]
    | GridLine (x1, y1, x2, y2, color, width) ->
        [ PdfDraw.SetStrokeColor (parseHexColor color)
          PdfDraw.SetLineWidth (px width)
          PdfDraw.SetDash ([], 0.0)
          PdfDraw.MoveTo (px x1, yPdf y1)
          PdfDraw.LineTo (px x2, yPdf y2)
          PdfDraw.Stroke ]
    | DrawText t ->
        match pickFace options t.FontName t.Bold t.Italic with
        | None -> []
        | Some primaryFace ->
            match Map.tryFind primaryFace.FaceName faceByName with
            | Some primaryTtf ->
                let fontSizePt = px t.FontSize
                let segments = segmentByFont (primaryFace, primaryTtf) options.CjkFallback faceByName t.Bold t.Text
                // Width in pt for each segment (uses its own font's UPM).
                let segWidthPt (ttf: PdfFont.TrueTypeFont) (s: string) =
                    float (PdfFont.getTextWidth ttf s) / float ttf.UnitsPerEm * fontSizePt
                let segmentsWithWidth =
                    segments |> List.map (fun (face, s) -> face, s, segWidthPt (snd face) s)
                let totalWidthPt = segmentsWithWidth |> List.sumBy (fun (_, _, w) -> w)
                let rectXPt = px t.Rect.X
                let rectWidthPt = px t.Rect.Width
                let alignedXPt =
                    match t.Align with
                    | AlignLeft -> rectXPt
                    | AlignCenter -> rectXPt + (rectWidthPt - totalWidthPt) / 2.0
                    | AlignRight -> rectXPt + rectWidthPt - totalWidthPt
                // Baseline from the primary font's win-metrics so a Calibri
                // line with embedded CJK glyphs doesn't jitter when fallback
                // ascents differ. Mirrors SpreadsheetGrid.toRendererCommands.
                let scaleY = t.FontSize / float primaryTtf.UnitsPerEm
                let ascent = float primaryTtf.WinAscent * scaleY
                let descent = float primaryTtf.WinDescent * scaleY
                let lineHeight = ascent + descent
                let blockHeight = lineHeight * float t.LineCount
                let blockTopY =
                    match t.VerticalAlign with
                    | VAlignTop -> t.Rect.Y
                    | VAlignMiddle -> t.Rect.Y + (t.Rect.Height - blockHeight) / 2.0
                    | VAlignBottom -> t.Rect.Y + t.Rect.Height - blockHeight
                // Super/subscript shifts the baseline relative to the
                // line's normal baseline. BaselineShift comes in pt; the
                // PDF coordinate space is also in pt so it applies
                // directly without unit conversion.
                let baselineY = blockTopY + float t.LineIndex * lineHeight + ascent + t.BaselineShift
                let textColor = parseHexColor t.Color
                let glyphCmds =
                    [ yield PdfDraw.BeginText
                      yield PdfDraw.SetFillColor textColor
                      let mutable cursorXPt = alignedXPt
                      for (face, s, w) in segmentsWithWidth do
                          let faceName = (fst face).FaceName
                          let _, ttf = face
                          let remap = faceRemap.TryFind faceName |> Option.defaultValue Map.empty
                          let hex = toRemappedHex ttf remap s
                          yield PdfDraw.SetFont (faceName, fontSizePt)
                          yield PdfDraw.TextMatrix (1.0, 0.0, 0.0, 1.0, cursorXPt, yPdf baselineY)
                          yield PdfDraw.ShowTextHex hex
                          cursorXPt <- cursorXPt + w
                      yield PdfDraw.EndText ]
                let underlineCmds =
                    if t.Underline && totalWidthPt > 0.0 then
                        let underlineY = baselineY + max 1.0 (t.FontSize * 0.08)
                        let single =
                            [ PdfDraw.SetStrokeColor textColor
                              PdfDraw.SetLineWidth (px 1.0)
                              PdfDraw.SetDash ([], 0.0)
                              PdfDraw.MoveTo (alignedXPt, yPdf underlineY)
                              PdfDraw.LineTo (alignedXPt + totalWidthPt, yPdf underlineY)
                              PdfDraw.Stroke ]
                        if t.UnderlineDouble then
                            // Second parallel stroke ~2 pt below the first.
                            let underlineY2 = underlineY + max 2.0 (t.FontSize * 0.12)
                            single @
                            [ PdfDraw.MoveTo (alignedXPt, yPdf underlineY2)
                              PdfDraw.LineTo (alignedXPt + totalWidthPt, yPdf underlineY2)
                              PdfDraw.Stroke ]
                        else single
                    else []
                let strikeCmds =
                    if t.Strikethrough && totalWidthPt > 0.0 then
                        let strikeY = baselineY - t.FontSize * 0.28
                        [ PdfDraw.SetStrokeColor textColor
                          PdfDraw.SetLineWidth (px 1.0)
                          PdfDraw.SetDash ([], 0.0)
                          PdfDraw.MoveTo (alignedXPt, yPdf strikeY)
                          PdfDraw.LineTo (alignedXPt + totalWidthPt, yPdf strikeY)
                          PdfDraw.Stroke ]
                    else []
                glyphCmds @ underlineCmds @ strikeCmds
            | _ -> []

/// Render a workbook page to PDF bytes. Currently emits a single page sized
/// to `options.PageWidth/PageHeight` containing the whole used range; no
/// pagination yet.
///
/// Embedding strategy: parse every face up-front to drive segmentation and
/// width-measurement, then subset-embed each face using just the glyphs
/// actually used by the page's text. Without subsetting a workbook that
/// touches a few kanji would pull in ~14 MB of YuGothR / YuGothB / MS
/// Gothic per CJK font — the resulting PDF balloons past 100 MB. The
/// subset path produces ~30 KB embeds for the same content.
let toBytes (options: RenderOptions) (model: SpreadsheetModel) : byte[] =
    let builder = PdfDoc.create ()

    // First pass: parse face data without embedding. We need TrueTypeFont
    // values for segmentation (Cmap lookups), measurement (GlyphWidths),
    // and to compute used-GID sets for subset embed.
    let parsed = System.Collections.Generic.Dictionary<string, PdfFont.TrueTypeFont>()
    let parseFace (face: FontFace) =
        try parsed.[face.FaceName] <- PdfFont.parseTrueType face.Data
        with _ -> ()  // unparseable face — skip silently rather than abort
    for face in options.Fonts do parseFace face
    for face in options.CjkFallback do parseFace face

    let parsedMap = parsed |> Seq.map (fun kv -> kv.Key, kv.Value) |> Map.ofSeq

    // Measurer mirrors the renderer: walks the same segmentation so wrap
    // widths agree with what we ultimately draw.
    let measurePdf (text: string) (fontPt: float) (fontName: string option) (bold: bool) (italic: bool) : float =
        let approx () = float text.Length * 7.0
        match pickFace options fontName bold italic with
        | None -> approx ()
        | Some primaryFace ->
            match parsed.TryGetValue primaryFace.FaceName with
            | true, primaryTtf ->
                let segments = segmentByFont (primaryFace, primaryTtf) options.CjkFallback parsedMap bold text
                segments
                |> List.sumBy (fun (face, s) ->
                    let _, ttf = face
                    let widthUnits = PdfFont.getTextWidth ttf s
                    float widthUnits / float ttf.UnitsPerEm * fontPt)
            | _ -> approx ()

    let viewport = {
        ScrollX = 0.0
        ScrollY = 0.0
        Width = options.PageWidth
        Height = options.PageHeight
        HeaderHeight = 22.0
        RowHeaderWidth = 48.0
        Measurer = Some measurePdf
    }

    let pageWidthPt = px options.PageWidth
    let pageHeightPt = px options.PageHeight

    let items = buildPaintItems model viewport

    // Collect used GIDs per face by replaying the same segmentation we'll
    // use during op emission. The subset closes over composite-glyph refs
    // itself; we only need the directly-referenced GIDs from cmap lookup.
    let usedGids = System.Collections.Generic.Dictionary<string, System.Collections.Generic.HashSet<int>>()
    let recordGids (faceName: string) (ttf: PdfFont.TrueTypeFont) (text: string) =
        let set =
            match usedGids.TryGetValue faceName with
            | true, s -> s
            | _ ->
                let s = System.Collections.Generic.HashSet<int>()
                usedGids.[faceName] <- s
                s
        for ch in text do
            let gid = ttf.Cmap.TryFind (int ch) |> Option.defaultValue 0
            set.Add gid |> ignore
    for item in items do
        match item with
        | DrawText t ->
            match pickFace options t.FontName t.Bold t.Italic with
            | Some primaryFace ->
                match parsed.TryGetValue primaryFace.FaceName with
                | true, primaryTtf ->
                    let segments = segmentByFont (primaryFace, primaryTtf) options.CjkFallback parsedMap t.Bold t.Text
                    for (face, s) in segments do
                        recordGids (fst face).FaceName (snd face) s
                | _ -> ()
            | None -> ()
        | _ -> ()

    // Subset-embed only the faces referenced by drawn text. Skipping
    // unreferenced faces drops a non-CJK page from ~18 MB to ~250 KB by
    // dropping the unused Cambria/Times/Verdana/CJK weights the runner
    // optimistically loaded "just in case". Faces that the subsetter
    // rejects (CFF, unusual layouts) fall back to a full embed so we
    // still draw rather than show tofu.
    let faceRemap = System.Collections.Generic.Dictionary<string, Map<int, int>>()
    let embedFace (face: FontFace) =
        match parsed.TryGetValue face.FaceName with
        | false, _ -> ()
        | true, _ ->
            let used =
                match usedGids.TryGetValue face.FaceName with
                | true, s -> Set.ofSeq s
                | _ -> Set.empty
            if Set.isEmpty used then () else
            try
                let remap = PdfDoc.loadFontBytesSubset builder face.FaceName face.Data used
                if Map.isEmpty remap then
                    // subsetCompact returned (data, empty) — likely CFF; fall
                    // back to full embed so the face is at least usable.
                    PdfDoc.loadFontBytes builder face.FaceName face.Data |> ignore
                    faceRemap.[face.FaceName] <- Map.empty
                else
                    faceRemap.[face.FaceName] <- remap
            with _ ->
                try
                    PdfDoc.loadFontBytes builder face.FaceName face.Data |> ignore
                    faceRemap.[face.FaceName] <- Map.empty
                with _ -> ()
    for face in options.Fonts do embedFace face
    for face in options.CjkFallback do embedFace face

    // Refresh TrueTypeFont references — subset embed swaps the embedded
    // font's GlyphWidths/Cmap to the post-subset numbering. We keep the
    // ORIGINAL parsed font for segmentation/measurement (subsetting doesn't
    // alter glyph metrics for kept glyphs, just their indices) so the only
    // place the remap matters is hex emission.
    let faceTtfMap = parsedMap
    let faceResourceMap =
        [ for face in options.Fonts @ options.CjkFallback do
            match PdfDoc.getEmbeddedFont builder face.FaceName with
            | Some ef -> yield face.FaceName, ef.Name
            | None -> () ]
        |> Map.ofList
    let faceRemapMap = faceRemap |> Seq.map (fun kv -> kv.Key, kv.Value) |> Map.ofSeq

    let ops = items |> List.collect (translateItem options faceTtfMap faceResourceMap faceRemapMap pageHeightPt)

    // addPage's first font list controls which entries make it into the
    // page's `/Font` resources dict — without it the SetFont (F#, …) ops we
    // already emitted reference unknown resource names and the reader
    // refuses to draw any text. Filter to faces that actually embedded.
    let loadedFaceNames =
        (options.Fonts @ options.CjkFallback)
        |> List.map (fun f -> f.FaceName)
        |> List.filter faceResourceMap.ContainsKey
    PdfDoc.addPage builder (pageWidthPt, pageHeightPt) loadedFaceNames [] ops |> ignore
    PdfDoc.toBytes builder
