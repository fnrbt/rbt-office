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
    /// Sheet name surfaced as `&A` in headers / footers.
    SheetName: string
    /// File name surfaced as `&F` in headers / footers (bare leaf, no
    /// directory). Empty when the source isn't a file.
    FileName: string
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

/// Page placement: where this page's content sits on the paper, what the
/// page is sized at, and how content is scaled. translateItem bakes all
/// three into every coordinate it emits.
type private Placement = {
    /// PDF points from the paper's left edge to where page-local X=0 lands.
    LeftPtOffset: float
    /// PDF points from the paper's *top* edge to where page-local Y=0 lands
    /// (the top of the printable area).
    TopPtOffset: float
    /// PDF points for the paper's full height — needed for the Y-flip.
    PageHeightPt: float
    /// Multiplier applied to every size + offset (fit-to-page / manual scale).
    Scale: float
}

/// Convert one Paint item into a list of PDF draw ops. Coordinates flip from
/// our top-left CSS-pixel space into PDF's bottom-left point space, with the
/// `Placement`'s margin offsets and scale baked in.
let private translateItem (options: RenderOptions) (faceByName: Map<string, PdfFont.TrueTypeFont>) (faceResource: Map<string, string>) (faceRemap: Map<string, Map<int, int>>) (placement: Placement) (item: PaintItem) : PdfDraw.DrawOp list =
    let sz v = v * pxToPt * placement.Scale          // size in pt
    let xPt x = placement.LeftPtOffset + sz x
    let yPdf y = placement.PageHeightPt - placement.TopPtOffset - sz y
    match item with
    | FillRect (rect, color) ->
        [ PdfDraw.SetFillColor (parseHexColor color)
          PdfDraw.Rect (xPt rect.X, yPdf (rect.Y + rect.Height), sz rect.Width, sz rect.Height)
          PdfDraw.Fill ]
    | StrokeRect (rect, color, width) ->
        [ PdfDraw.SetStrokeColor (parseHexColor color)
          PdfDraw.SetLineWidth (sz width)
          PdfDraw.SetDash ([], 0.0)
          PdfDraw.Rect (xPt rect.X, yPdf (rect.Y + rect.Height), sz rect.Width, sz rect.Height)
          PdfDraw.Stroke ]
    | GridLine (x1, y1, x2, y2, color, width) ->
        [ PdfDraw.SetStrokeColor (parseHexColor color)
          PdfDraw.SetLineWidth (sz width)
          PdfDraw.SetDash ([], 0.0)
          PdfDraw.MoveTo (xPt x1, yPdf y1)
          PdfDraw.LineTo (xPt x2, yPdf y2)
          PdfDraw.Stroke ]
    | DrawText t ->
        match pickFace options t.FontName t.Bold t.Italic with
        | None -> []
        | Some primaryFace ->
            match Map.tryFind primaryFace.FaceName faceByName with
            | Some primaryTtf ->
                let fontSizePt = sz t.FontSize
                let segments = segmentByFont (primaryFace, primaryTtf) options.CjkFallback faceByName t.Bold t.Text
                // Width in pt for each segment (uses its own font's UPM).
                let segWidthPt (ttf: PdfFont.TrueTypeFont) (s: string) =
                    float (PdfFont.getTextWidth ttf s) / float ttf.UnitsPerEm * fontSizePt
                let segmentsWithWidth =
                    segments |> List.map (fun (face, s) -> face, s, segWidthPt (snd face) s)
                let totalWidthPt = segmentsWithWidth |> List.sumBy (fun (_, _, w) -> w)
                let rectXPt = xPt t.Rect.X
                let rectWidthPt = sz t.Rect.Width
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
                              PdfDraw.SetLineWidth (sz 1.0)
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
                          PdfDraw.SetLineWidth (sz 1.0)
                          PdfDraw.SetDash ([], 0.0)
                          PdfDraw.MoveTo (alignedXPt, yPdf strikeY)
                          PdfDraw.LineTo (alignedXPt + totalWidthPt, yPdf strikeY)
                          PdfDraw.Stroke ]
                    else []
                glyphCmds @ underlineCmds @ strikeCmds
            | _ -> []

/// Sum of column widths in an inclusive range.
let private colSumPx (model: SpreadsheetModel) (startInc: int) (endInc: int) =
    let mutable s = 0.0
    for c in startInc .. endInc do
        s <- s + (Map.tryFind c model.ColumnWidths |> Option.defaultValue 100.0)
    s

/// Sum of row heights in an inclusive range.
let private rowSumPx (model: SpreadsheetModel) (startInc: int) (endInc: int) =
    let mutable s = 0.0
    for r in startInc .. endInc do
        s <- s + (Map.tryFind r model.RowHeights |> Option.defaultValue 25.0)
    s

/// True when a paint item has degenerate (zero-size) geometry. With the
/// print viewport's HeaderHeight / RowHeaderWidth both set to 0, the
/// column-letter / row-number chrome items collapse to W=0 or H=0; they'd
/// otherwise draw faint baseline marks at the top of every page.
let private isDegenerate (item: PaintItem) : bool =
    match item with
    | FillRect (r, _) | StrokeRect (r, _, _) -> r.Width < 0.5 || r.Height < 0.5
    | DrawText t -> t.Rect.Width < 0.5 || t.Rect.Height < 0.5
    | GridLine (x1, y1, x2, y2, _, _) ->
        // A real gridline has length in exactly one axis; only drop the
        // genuinely zero-length point.
        abs (x2 - x1) < 0.5 && abs (y2 - y1) < 0.5

/// Shift every coordinate of a paint item by (dx, dy) CSS px. Used to map
/// items from their absolute viewport position into page-local space.
let private shiftItem (dx: float) (dy: float) (item: PaintItem) : PaintItem =
    let shiftRect (r: Rect) = { r with X = r.X + dx; Y = r.Y + dy }
    match item with
    | FillRect (r, c) -> FillRect (shiftRect r, c)
    | StrokeRect (r, c, w) -> StrokeRect (shiftRect r, c, w)
    | DrawText t -> DrawText { t with Rect = shiftRect t.Rect }
    | GridLine (x1, y1, x2, y2, c, w) -> GridLine (x1 + dx, y1 + dy, x2 + dx, y2 + dy, c, w)

/// Build the paint items for a contiguous (rowStart..rowEnd × colStart..colEnd)
/// slice of the model, shifted so the slice's top-left lands at page-local
/// (offsetX, offsetY). Setting HeaderHeight / RowHeaderWidth to 0 in the
/// viewport keeps the column-letter / row-number chrome out — we only want
/// the body for print.
let private sliceItems
    (model: SpreadsheetModel)
    (measurer: TextMeasurer)
    (rowStart: int) (rowEnd: int)
    (colStart: int) (colEnd: int)
    (offsetX: float) (offsetY: float)
    : PaintItem list =
    // Scroll position is bumped past the strict-less-than threshold that
    // visibleRange's start-row loop uses, then height/width are inflated
    // by 1 so the end loop's strict-less-than admits the slice's last
    // row/col. Together they make the visible range land *exactly* on
    // [rowStart..rowEnd] × [colStart..colEnd]; without these nudges
    // visibleRange either pulls in the previous row (scroll-boundary
    // overlap) or drops the last one (end-boundary exclusion).
    let scrollX = colSumPx model 0 (colStart - 1) + 0.5
    let scrollY = rowSumPx model 0 (rowStart - 1) + 0.5
    let width = colSumPx model colStart colEnd + 1.0
    let height = rowSumPx model rowStart rowEnd + 1.0
    let viewport =
        { ScrollX = scrollX
          ScrollY = scrollY
          Width = width
          Height = height
          HeaderHeight = 0.0
          RowHeaderWidth = 0.0
          Measurer = Some measurer }
    buildPaintItems model viewport
    |> List.filter (isDegenerate >> not)
    |> List.map (shiftItem (offsetX - scrollX) (offsetY - scrollY))

/// Emit ops for one of the (left, center, right) sections of a header or
/// footer. Y is the PDF-space baseline (already accounting for the band).
/// `xAlign` is one of `AlignLeft` / `AlignCenter` / `AlignRight`, applied
/// against the printable-area span (leftPt → rightPt).
let private sectionOps
    (options: RenderOptions)
    (faceByName: Map<string, PdfFont.TrueTypeFont>)
    (faceResource: Map<string, string>)
    (faceRemap: Map<string, Map<int, int>>)
    (leftPt: float) (rightPt: float)
    (text: string) (fontSizePt: float) (baselineYPt: float)
    (xAlign: TextAlign) : PdfDraw.DrawOp list =
    if String.IsNullOrEmpty text then [] else
    match pickFace options None false false with
    | None -> []
    | Some primaryFace ->
        match Map.tryFind primaryFace.FaceName faceByName with
        | None -> []
        | Some primaryTtf ->
            let segments = segmentByFont (primaryFace, primaryTtf) options.CjkFallback faceByName false text
            let segWidthPt (ttf: PdfFont.TrueTypeFont) (s: string) =
                float (PdfFont.getTextWidth ttf s) / float ttf.UnitsPerEm * fontSizePt
            let segmentsWithWidth =
                segments |> List.map (fun (face, s) -> face, s, segWidthPt (snd face) s)
            let totalWidthPt = segmentsWithWidth |> List.sumBy (fun (_, _, w) -> w)
            let startXPt =
                match xAlign with
                | AlignLeft -> leftPt
                | AlignCenter -> leftPt + ((rightPt - leftPt) - totalWidthPt) / 2.0
                | AlignRight -> rightPt - totalWidthPt
            [ yield PdfDraw.BeginText
              yield PdfDraw.SetFillColor PdfColor.black
              let mutable cursorXPt = startXPt
              for (face, s, w) in segmentsWithWidth do
                  let faceName = (fst face).FaceName
                  let _, ttf = face
                  let remap = faceRemap.TryFind faceName |> Option.defaultValue Map.empty
                  let hex = toRemappedHex ttf remap s
                  yield PdfDraw.SetFont (faceName, fontSizePt)
                  yield PdfDraw.TextMatrix (1.0, 0.0, 0.0, 1.0, cursorXPt, baselineYPt)
                  yield PdfDraw.ShowTextHex hex
                  cursorXPt <- cursorXPt + w
              yield PdfDraw.EndText ]

/// Render a workbook page to PDF bytes. Drives 2D pagination off the
/// model's PageSetup: each page slice is built into a viewport, shifted
/// onto the printable area (between the page margins) at the configured
/// scale, with repeat rows / cols and headers / footers added per page.
///
/// Embedding strategy: parse every face up-front, walk every page's items
/// to collect used glyph IDs, then subset-embed each face with just those
/// glyphs. Without subsetting a workbook that touches a few kanji would
/// pull in ~14 MB of YuGothR / YuGothB / MS Gothic per CJK font and the
/// resulting PDF balloons past 100 MB. The subset path produces ~30 KB
/// embeds for the same content.
let toBytes (options: RenderOptions) (model: SpreadsheetModel) : byte[] =
    let builder = PdfDoc.create ()

    let parsed = System.Collections.Generic.Dictionary<string, PdfFont.TrueTypeFont>()
    let parseFace (face: FontFace) =
        try parsed.[face.FaceName] <- PdfFont.parseTrueType face.Data
        with _ -> ()
    for face in options.Fonts do parseFace face
    for face in options.CjkFallback do parseFace face
    let parsedMap = parsed |> Seq.map (fun kv -> kv.Key, kv.Value) |> Map.ofSeq

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

    // Compute pagination from PageSetup. Empty workbook → one blank page.
    let pagination = Pagination.paginate model
    let ps = model.PageSetup
    let pageWidthPt = ps.PageWidth
    let pageHeightPt = ps.PageHeight
    let leftMarginPt = ps.Margins.Left * 72.0
    let rightMarginPt = ps.Margins.Right * 72.0
    let topMarginPt = ps.Margins.Top * 72.0
    let bottomMarginPt = ps.Margins.Bottom * 72.0
    let headerMarginPt = ps.Margins.Header * 72.0
    let footerMarginPt = ps.Margins.Footer * 72.0

    let totalPages = max 1 (List.length pagination.Pages)
    let repeatColsWidthPx =
        pagination.RepeatCols
        |> Option.map (fun (a, b) -> colSumPx model a b)
        |> Option.defaultValue 0.0
    let repeatRowsHeightPx =
        pagination.RepeatRows
        |> Option.map (fun (a, b) -> rowSumPx model a b)
        |> Option.defaultValue 0.0

    // Build every page's paint items up-front. We need them to drive font
    // subsetting before we can emit ops.
    let pageItems : PaintItem list array =
        let pages = pagination.Pages
        Array.init (max 1 (List.length pages)) (fun i ->
            if i >= List.length pages then [] else
            let slice = pages.[i]
            // Body: rows × cols of the slice, offset by repeat band sizes.
            let body =
                sliceItems model measurePdf
                    slice.RowStart slice.RowEnd
                    slice.ColStart slice.ColEnd
                    repeatColsWidthPx repeatRowsHeightPx
            // Repeat-rows band (above body, full width): repeat rows × slice cols.
            let repeatRows =
                match pagination.RepeatRows with
                | Some (rs, re) when re >= rs ->
                    sliceItems model measurePdf
                        rs re
                        slice.ColStart slice.ColEnd
                        repeatColsWidthPx 0.0
                | _ -> []
            // Repeat-cols band (left of body, full height): slice rows × repeat cols.
            let repeatCols =
                match pagination.RepeatCols with
                | Some (cs, ce) when ce >= cs ->
                    sliceItems model measurePdf
                        slice.RowStart slice.RowEnd
                        cs ce
                        0.0 repeatRowsHeightPx
                | _ -> []
            // Repeat corner (top-left): repeat rows × repeat cols.
            let repeatCorner =
                match pagination.RepeatRows, pagination.RepeatCols with
                | Some (rs, re), Some (cs, ce) when re >= rs && ce >= cs ->
                    sliceItems model measurePdf rs re cs ce 0.0 0.0
                | _ -> []
            body @ repeatRows @ repeatCols @ repeatCorner)

    // Collect used GIDs across every page's items, plus header / footer
    // sections (so their glyphs land in the subset too).
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
    let recordText (text: string) (bold: bool) =
        match pickFace options None bold false with
        | None -> ()
        | Some primaryFace ->
            match parsed.TryGetValue primaryFace.FaceName with
            | true, primaryTtf ->
                let segments = segmentByFont (primaryFace, primaryTtf) options.CjkFallback parsedMap bold text
                for (face, s) in segments do
                    recordGids (fst face).FaceName (snd face) s
            | _ -> ()
    for items in pageItems do
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

    // Build header/footer page contexts and pre-resolve all text so we
    // can register its glyphs for subsetting before embedding fonts.
    let resolvedHFs =
        [ for pageIdx in 0 .. totalPages - 1 ->
            let ctx : HeaderFooter.PageContext =
                { PageNumber = pageIdx + 1
                  TotalPages = totalPages
                  SheetName = options.SheetName
                  FileName = options.FileName
                  FilePath = ""
                  Now = DateTime.Now }
            let header =
                ps.HeaderFooter.OddHeader
                |> Option.map (HeaderFooter.resolve ctx)
            let footer =
                ps.HeaderFooter.OddFooter
                |> Option.map (HeaderFooter.resolve ctx)
            header, footer ]
    for (header, footer) in resolvedHFs do
        for section in [ header; footer ] |> List.choose id do
            recordText section.Left false
            recordText section.Center false
            recordText section.Right false

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

    let faceTtfMap = parsedMap
    let faceResourceMap =
        [ for face in options.Fonts @ options.CjkFallback do
            match PdfDoc.getEmbeddedFont builder face.FaceName with
            | Some ef -> yield face.FaceName, ef.Name
            | None -> () ]
        |> Map.ofList
    let faceRemapMap = faceRemap |> Seq.map (fun kv -> kv.Key, kv.Value) |> Map.ofSeq

    let loadedFaceNames =
        (options.Fonts @ options.CjkFallback)
        |> List.map (fun f -> f.FaceName)
        |> List.filter faceResourceMap.ContainsKey

    // Header/footer ops live in PDF coords directly (no scale), since the
    // header band sits in the page margins outside the scaled content area.
    let hfFontSizePt = 10.0
    let printableLeftPt = leftMarginPt
    let printableRightPt = pageWidthPt - rightMarginPt
    let headerBaselineYPt = pageHeightPt - headerMarginPt - hfFontSizePt
    let footerBaselineYPt = footerMarginPt
    let renderSection (section: HeaderFooter.Section) (baselineYPt: float) =
        sectionOps options faceTtfMap faceResourceMap faceRemapMap
            printableLeftPt printableRightPt
            section.Left hfFontSizePt baselineYPt AlignLeft
        @ sectionOps options faceTtfMap faceResourceMap faceRemapMap
            printableLeftPt printableRightPt
            section.Center hfFontSizePt baselineYPt AlignCenter
        @ sectionOps options faceTtfMap faceResourceMap faceRemapMap
            printableLeftPt printableRightPt
            section.Right hfFontSizePt baselineYPt AlignRight

    // Emit one PDF page per page slice (or one blank page when the
    // pagination is empty).
    for pageIdx in 0 .. totalPages - 1 do
        let items = pageItems.[pageIdx]
        let placement =
            { LeftPtOffset = leftMarginPt
              TopPtOffset = topMarginPt
              PageHeightPt = pageHeightPt
              Scale = pagination.Scale }
        let contentOps =
            items |> List.collect (translateItem options faceTtfMap faceResourceMap faceRemapMap placement)
        let header, footer = resolvedHFs.[pageIdx]
        let headerOps =
            header |> Option.map (fun s -> renderSection s headerBaselineYPt) |> Option.defaultValue []
        let footerOps =
            footer |> Option.map (fun s -> renderSection s footerBaselineYPt) |> Option.defaultValue []
        let ops = contentOps @ headerOps @ footerOps
        PdfDoc.addPage builder (pageWidthPt, pageHeightPt) loadedFaceNames [] ops |> ignore

    PdfDoc.toBytes builder
