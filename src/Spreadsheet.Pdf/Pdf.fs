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
    /// When true, every fill becomes white and every stroke / text becomes
    /// black. Matches Excel's `<pageSetup blackAndWhite="1"/>` print mode.
    BlackAndWhite: bool
}

let private bwFill (bw: bool) (color: string) =
    if bw then "#FFFFFF" else color

let private bwStroke (bw: bool) (color: string) =
    if bw then "#000000" else color

/// Convert one Paint item into a list of PDF draw ops. Coordinates flip from
/// our top-left CSS-pixel space into PDF's bottom-left point space, with the
/// `Placement`'s margin offsets and scale baked in.
let private translateItem (options: RenderOptions) (faceByName: Map<string, PdfFont.TrueTypeFont>) (faceResource: Map<string, string>) (faceRemap: Map<string, Map<int, int>>) (placement: Placement) (item: PaintItem) : PdfDraw.DrawOp list =
    let sz v = v * pxToPt * placement.Scale          // size in pt
    let xPt x = placement.LeftPtOffset + sz x
    let yPdf y = placement.PageHeightPt - placement.TopPtOffset - sz y
    // Vector paths/groups. The 2×3 transform is applied in CSS space (so it
    // matches the native/canvas Group), then each point is flipped into PDF
    // space via xPt/yPdf. VecGroup composes its transform onto its children.
    let rec vec (a, b, c, d, e, f) (it: PaintItem) : PdfDraw.DrawOp list =
        match it with
        | VecPath (segs, fill, stroke) ->
            let pt x y = xPt (a * x + c * y + e), yPdf (b * x + d * y + f)
            let pathOps =
                segs |> List.map (fun s ->
                    match s with
                    | VMove (x, y) -> let px, py = pt x y in PdfDraw.MoveTo (px, py)
                    | VLine (x, y) -> let px, py = pt x y in PdfDraw.LineTo (px, py)
                    | VCubic (x1, y1, x2, y2, x, y) ->
                        let p1x, p1y = pt x1 y1
                        let p2x, p2y = pt x2 y2
                        let px, py = pt x y
                        PdfDraw.CurveTo (p1x, p1y, p2x, p2y, px, py)
                    | VClose -> PdfDraw.ClosePath)
            let setup =
                [ match fill with
                  | Some col -> PdfDraw.SetFillColor (parseHexColor col)
                  | None -> ()
                  match stroke with
                  | Some (col, w) ->
                      PdfDraw.SetStrokeColor (parseHexColor col)
                      PdfDraw.SetLineWidth (sz w)
                  | None -> () ]
            let paint =
                match fill, stroke with
                | Some _, Some _ -> [ PdfDraw.FillAndStroke ]
                | Some _, None -> [ PdfDraw.Fill ]
                | None, Some _ -> [ PdfDraw.Stroke ]
                | None, None -> []
            setup @ pathOps @ paint
        | VecGroup ((a2, b2, c2, d2, e2, f2), children) ->
            let comp = (a*a2 + c*b2, b*a2 + d*b2, a*c2 + c*d2, b*c2 + d*d2, a*e2 + c*f2 + e, b*e2 + d*f2 + f)
            children |> List.collect (vec comp)
        | _ -> []
    match item with
    | (VecPath _ | VecGroup _) -> vec (1.0, 0.0, 0.0, 1.0, 0.0, 0.0) item
    | FillRect (rect, color) ->
        [ PdfDraw.SetFillColor (parseHexColor (bwFill placement.BlackAndWhite color))
          PdfDraw.Rect (xPt rect.X, yPdf (rect.Y + rect.Height), sz rect.Width, sz rect.Height)
          PdfDraw.Fill ]
    | StrokeRect (rect, color, width) ->
        [ PdfDraw.SetStrokeColor (parseHexColor (bwStroke placement.BlackAndWhite color))
          PdfDraw.SetLineWidth (sz width)
          PdfDraw.SetDash ([], 0.0)
          PdfDraw.Rect (xPt rect.X, yPdf (rect.Y + rect.Height), sz rect.Width, sz rect.Height)
          PdfDraw.Stroke ]
    | GridLine (x1, y1, x2, y2, color, width) ->
        [ PdfDraw.SetStrokeColor (parseHexColor (bwStroke placement.BlackAndWhite color))
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
                // PaintText.FontSize is stored in points. Draw in PDF points,
                // but use CSS-pixel metrics for layout because cell geometry is
                // still page-local CSS px before translateItem applies pxToPt.
                let fontPx = t.FontSize * 96.0 / 72.0
                let fontSizePt = t.FontSize * placement.Scale
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
                let scaleY = fontPx / float primaryTtf.UnitsPerEm
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
                // line's normal baseline. BaselineShift is page-local CSS px;
                // yPdf converts it to PDF coords together with the baseline.
                // No explicit nudge: the OS/2 winAscent metric maps directly
                // to Excel's print baseline for the common Calibri 11
                // workbooks (synthetic-print + header/footer fixtures all
                // saw row-text drop into pixel alignment when we removed
                // the historical -2.0 fudge from the earlier renderer).
                // A residual ~1 px drift remains for some Calibri 12 cells
                // (pandas-test1 col-B numbers, ~+0.3 body MAE), but the
                // net MAE across the 8 print fixtures drops from 30.95 to
                // 29.16 vs the fudged baseline.
                let baselineY = blockTopY + float t.LineIndex * lineHeight + ascent + t.BaselineShift
                let textColor = parseHexColor (bwStroke placement.BlackAndWhite t.Color)
                // Cell-text rotation (Excel `<alignment textRotation>`).
                // 0 = horizontal, 1..90 = CCW degrees, 91..180 = CW degrees
                // encoded as 90 + |angle|, 255 = stacked. We treat stacked
                // as 0 here (per-glyph stacking would need a separate emit
                // loop); 1..180 rotates the text around its anchor.
                let useRotation = t.Rotation > 0 && t.Rotation <> 255
                let angleRad =
                    if not useRotation then 0.0
                    elif t.Rotation <= 90 then float t.Rotation * System.Math.PI / 180.0
                    else -(float (t.Rotation - 90)) * System.Math.PI / 180.0
                let cosT = if useRotation then System.Math.Cos angleRad else 1.0
                let sinT = if useRotation then System.Math.Sin angleRad else 0.0
                // Rotated-text placement, matching the native renderer + Excel
                // (rule derived from tools/excel-parity/build-rotation-sweep.vbs):
                // rotate the run about its baseline start, then position the
                // rotated bounding box — RIGHT-aligned when 90 ≤ textRotation <
                // 180, else LEFT-aligned, always BOTTOM-aligned. Computed in
                // CSS px (y-down) then converted; the glyph loop below extends
                // the run along the PDF (y-up) rotation from this anchor.
                let anchorXPt, anchorYPt =
                    if not useRotation then alignedXPt, yPdf baselineY
                    else
                        let thetaS =
                            if t.Rotation <= 90 then -(float t.Rotation) * System.Math.PI / 180.0
                            else float (t.Rotation - 90) * System.Math.PI / 180.0
                        let cosS, sinS = System.Math.Cos thetaS, System.Math.Sin thetaS
                        let totalWidthPx =
                            segments
                            |> List.sumBy (fun (face, s) ->
                                float (PdfFont.getTextWidth (snd face) s) / float (snd face).UnitsPerEm * fontPx)
                        let ax0 = t.Rect.X + 4.0
                        let ay0 = t.Rect.Y + t.Rect.Height - 2.0
                        let rotPt x y =
                            let dx, dy = x - ax0, y - ay0
                            ax0 + dx * cosS - dy * sinS, ay0 + dx * sinS + dy * cosS
                        let corners =
                            [ rotPt ax0 (ay0 - ascent)
                              rotPt (ax0 + totalWidthPx) (ay0 - ascent)
                              rotPt ax0 (ay0 + descent)
                              rotPt (ax0 + totalWidthPx) (ay0 + descent) ]
                        let minX = corners |> List.map fst |> List.min
                        let maxX = corners |> List.map fst |> List.max
                        let maxY = corners |> List.map snd |> List.max
                        let rightAlign = t.Rotation >= 90 && t.Rotation < 180
                        let shiftX =
                            if rightAlign then (t.Rect.X + t.Rect.Width - 2.0) - maxX
                            else (t.Rect.X + 2.0) - minX
                        let shiftY = (t.Rect.Y + t.Rect.Height - 2.0) - maxY
                        xPt (ax0 + shiftX), yPdf (ay0 + shiftY)
                let glyphCmds =
                    [ yield PdfDraw.BeginText
                      yield PdfDraw.SetFillColor textColor
                      let mutable cumWidthPt = 0.0
                      for (face, s, w) in segmentsWithWidth do
                          let faceName = (fst face).FaceName
                          let _, ttf = face
                          let remap = faceRemap.TryFind faceName |> Option.defaultValue Map.empty
                          let hex = toRemappedHex ttf remap s
                          yield PdfDraw.SetFont (faceName, fontSizePt)
                          // Position this segment's anchor in PDF coords.
                          // Without rotation: anchor + (cumWidth, 0).
                          // With rotation: anchor + cumWidth × (cosθ, sinθ).
                          let xAt = anchorXPt + cosT * cumWidthPt
                          let yAt = anchorYPt + sinT * cumWidthPt
                          yield PdfDraw.TextMatrix (cosT, sinT, -sinT, cosT, xAt, yAt)
                          yield PdfDraw.ShowTextHex hex
                          cumWidthPt <- cumWidthPt + w
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

/// Excel's live grid uses the floored screen MDW for explicit `<col width>`
/// values, but its PDF print path lays those explicit widths out with the
/// non-floored Calibri-11 width (about 7.5 px instead of 7 px). Keep the
/// screen model untouched and apply that delta only while rendering PDF
/// pages. Seeded default-width columns are left alone; the default width is
/// inferred as the most common width in the model's column map.
/// Excel's print default for cell comments is "none" — comments don't
/// appear in the printout regardless of the on-screen red marker. Strip
/// `HasComment` flags from the cell formats unless `CellComments` opts in.
let private withCommentMarkersForPrint (model: SpreadsheetModel) : SpreadsheetModel =
    match model.PageSetup.CellComments with
    | CellCommentsAsDisplayed -> model
    | CellCommentsNone | CellCommentsAtEnd ->
        let cells =
            model.Cells
            |> Map.map (fun _ cd ->
                if cd.Format.HasComment then { cd with Format = { cd.Format with HasComment = false } }
                else cd)
        { model with Cells = cells }

let private withPrintColumnWidths (model: SpreadsheetModel) : SpreadsheetModel =
    if Map.isEmpty model.ColumnWidths then model else
    // Excel's live grid lays columns out with the floored screen MDW (7 px for
    // Calibri 11), but its PDF print path uses the non-floored width (~7.5 px),
    // i.e. ~15/14 wider — for EVERY column, including the workbook-default
    // width. Measured on `synthetic-diagonal-borders.xlsx`: a default 8.43-char
    // column prints at ~68.8 px vs our 64 px; 64 × 15/14 = 68.6 px. Earlier
    // this scale was applied only to explicit `<col>` widths, which left the
    // default columns ~5 px narrow and drifted every cell right of column A
    // leftward. Scale all positive widths.
    let printScale = 15.0 / 14.0
    let widths =
        model.ColumnWidths
        |> Map.map (fun _ width -> if width <= 0.0 then width else width * printScale)
    { model with ColumnWidths = widths }

/// Excel's print PDF renders Calibri-11 auto-derived rows at 19 px (14.25 pt)
/// even though `cell.Height` via COM still reports 15 pt (20 px). This is a
/// long-standing Excel quirk specific to that font/size pair — verified
/// empirically against `synthetic-print-chrome` (no styles.xml, Calibri 11
/// default, row pitch 19.4 px) and `pandas-dimension_large` body (Calibri 11
/// cells, row pitch 19-20 px alternating). Other auto-row heights match
/// editor exactly: Calibri 12 → 21 px (no shift), Calibri 18 Title → 31 px
/// (no shift). The harness must keep reporting 15 pt (matches Excel COM),
/// so the adjustment lives only in the print path here.
///
/// Strategy: any row whose model height is exactly 20 px gets remapped to
/// 19.36 px (= 14.52 pt) during PDF rendering. The fractional value
/// matters — Excel's actual print pitch is 14.55 pt, which on the integer
/// pixel grid alternates 19/20 px as cumulative y crosses each pixel
/// boundary. Mapping to 19 px exactly produces a uniform 19-pitch that
/// drifts the OTHER way from Excel; the fractional value lets the PDF
/// rasterizer reproduce Excel's alternating pattern. Other heights pass
/// through unchanged (preserves explicit `ht` values, Title rows, and
/// font-derived rows from non-11pt cells).
///
/// The 20→19.36 mapping carries a theoretical regression risk if a
/// workbook explicitly sets `ht="15.0" customHeight="true"` — we'd shift
/// it to 14.52 pt when Excel would honor the stored 15 pt. No fixture in
/// the test corpus trips that; if one does, we'll need to thread per-row
/// "auto vs explicit" flags through the model and key the remap off that.
let private withPrintRowHeights (model: SpreadsheetModel) : SpreadsheetModel =
    if Map.isEmpty model.RowHeights then model else
    // Calibri 12 print pitch is similarly off — editor 21 px → print
    // ~20.64 px (= 15.48 pt). pandas-test1 rows positioned 1 px per row
    // below Excel because we used 21 px uniform while Excel's quantized
    // 20/21/21/21/21/20 pattern averages to ~20.64 px.
    let heights =
        model.RowHeights
        |> Map.map (fun _ px ->
            if abs (px - 20.0) < 0.01 then 19.36      // Calibri 11 auto
            elif abs (px - 21.0) < 0.01 then 20.64    // Calibri 12 auto
            else px)
    { model with RowHeights = heights }

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

/// The screen path emits a white fill for every cell even when the workbook
/// has no explicit fill. In print those default fills would cover the
/// print-gridline layer; the PDF page background already provides the same
/// white, so drop them while preserving non-white cell fills.
let private isDefaultPrintCellFill (item: PaintItem) : bool =
    match item with
    | FillRect (_, color) when color.Equals("#ffffff", StringComparison.OrdinalIgnoreCase) -> true
    | _ -> false

/// True when a paint item is the interactive view's always-on cell
/// gridline — a `StrokeRect` around every cell that the screen path emits
/// so the grid is visible without `gridLines="true"`. Print mode honors
/// the fixture's `PrintOptions.GridLines` (and emits its own GridLines
/// helper), so these per-cell strokes have to come out; otherwise an
/// unbordered, `gridLines="false"` workbook prints with our screen-grade
/// grid where Excel prints nothing. Explicit cell borders flow through
/// `emitBorderSide` as `GridLine` items, not `StrokeRect`, so they
/// survive the filter.
let private isInteractiveCellStroke (item: PaintItem) : bool =
    match item with
    | StrokeRect _ -> true
    // The interactive light gridline is now emitted as per-edge `GridLine`
    // items (colour #a6a6a6) rather than a per-cell StrokeRect, so the print
    // path must drop those too — Excel doesn't print the screen grid. Explicit
    // cell borders are GridLines in their own (non-#a6a6a6) colour and survive.
    | GridLine (_, _, _, _, color, _) ->
        color.Equals("#a6a6a6", StringComparison.OrdinalIgnoreCase)
    | _ -> false

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
    // Land the viewport on the exact rowStart..rowEnd / colStart..colEnd
    // rectangle. `visibleRange` now lands `startRow` on whatever row
    // contains `ScrollY` (top-of-cell semantics), so scrolling to the top
    // of `rowStart` picks `rowStart` cleanly; `Width`/`Height` exactly
    // matching the slice's content extents lets the end loop walk to
    // `rowEnd+1` and stop.
    let scrollX = colSumPx model 0 (colStart - 1)
    let scrollY = rowSumPx model 0 (rowStart - 1)
    let width = colSumPx model colStart colEnd
    let height = rowSumPx model rowStart rowEnd
    let viewport =
        { ScrollX = scrollX
          ScrollY = scrollY
          Width = width
          Height = height
          HeaderHeight = 0.0
          RowHeaderWidth = 0.0
          Measurer = Some measurer }
    buildPaintItems model viewport
    |> List.filter (fun item ->
        not (isDegenerate item)
        && not (isDefaultPrintCellFill item)
        && not (isInteractiveCellStroke item))
    |> List.map (shiftItem (offsetX - scrollX) (offsetY - scrollY))

/// Near-black cell-boundary lines for a slice, used when
/// `PrintOptions.GridLines` is on. Emits one vertical line at every
/// column edge (left edges of `colStart..colEnd` plus the right edge of
/// `colEnd`) and one horizontal at every row edge in the same fashion.
/// Coordinates are page-local CSS px shifted by `(offsetX, offsetY)`.
let private gridLineItems
    (model: SpreadsheetModel)
    (rowStart: int) (rowEnd: int)
    (colStart: int) (colEnd: int)
    (offsetX: float) (offsetY: float) : PaintItem list =
    if rowStart > rowEnd || colStart > colEnd then [] else
    let totalW = colSumPx model colStart colEnd
    let totalH = rowSumPx model rowStart rowEnd
    // Excel's printed gridlines rasterize as near-black in the reference
    // PDF, unlike the lighter screen grid used in the interactive view.
    let color = "#000000"
    let w = 1.0
    let xs =
        [ let mutable x = 0.0
          yield x
          for c in colStart .. colEnd do
              x <- x + (Map.tryFind c model.ColumnWidths |> Option.defaultValue 100.0)
              yield x ]
    let ys =
        [ let mutable y = 0.0
          yield y
          for r in rowStart .. rowEnd do
              y <- y + (Map.tryFind r model.RowHeights |> Option.defaultValue 25.0)
              yield y ]
    [ for x in xs -> GridLine (offsetX + x, offsetY, offsetX + x, offsetY + totalH, color, w)
      for y in ys -> GridLine (offsetX, offsetY + y, offsetX + totalW, offsetY + y, color, w) ]

/// Splice `gridlines` into a band's paint stream before any remaining
/// non-white cell fill / border / text. The print slice filter already
/// removes default white cell fills, so gridlines remain visible in
/// unfilled cells while explicit colored fills still cover them.
let private withGridlines (gridlines: PaintItem list) (items: PaintItem list) : PaintItem list =
    match items, gridlines with
    | _, [] -> items
    | (FillRect _ as bg) :: rest, _ -> bg :: gridlines @ rest
    | _ -> gridlines @ items

/// One DrawText item rendered with the default font at a fixed size and
/// alignment; used for the print-headings strips.
let private headingText (text: string) (rect: Rect) (align: TextAlign) (fontSize: float) : PaintItem =
    DrawText {
        Text = text
        Rect = rect
        Color = "#444444"
        FontSize = fontSize
        FontName = None
        Bold = false
        Italic = false
        Underline = false
        UnderlineDouble = false
        Strikethrough = false
        Align = align
        VerticalAlign = VAlignMiddle
        LineIndex = 0
        LineCount = 1
        // Body text gets a small PDF-only lift; heading labels need to stay
        // visually centered inside the shorter heading strip.
        BaselineShift = 2.0
        Rotation = 0 }

/// Build the column-letter strip above the body+repeats and the row-number
/// strip to its left, plus a separator gridline below the column-letter
/// strip and to the right of the row-number strip. `xLeft` / `yTop` is the
/// page-local CSS-px top-left of the heading-shifted content origin
/// (typically `(printHeadingsColWidthPx, printHeadingsRowHeightPx)` when
/// the page itself starts at 0,0). The strips' inner-padded text avoids
/// touching the separator lines.
let private headingItems
    (model: SpreadsheetModel)
    (slice: Pagination.PageSlice)
    (pagination: Pagination.Pagination)
    : PaintItem list =
    let headW = Pagination.printHeadingsColWidthPx
    let headH = Pagination.printHeadingsRowHeightPx
    // Column ranges to label: repeat cols (if any) followed by the slice's
    // body cols. Each emits its letter in the strip directly above the
    // matching column.
    let repeatColRange = pagination.RepeatCols |> Option.toList |> List.collect (fun (a, b) -> [a .. b])
    let bodyColRange = [ slice.ColStart .. slice.ColEnd ]
    let allCols = repeatColRange @ bodyColRange
    let repeatRowRange = pagination.RepeatRows |> Option.toList |> List.collect (fun (a, b) -> [a .. b])
    let bodyRowRange = [ slice.RowStart .. slice.RowEnd ]
    let allRows = repeatRowRange @ bodyRowRange
    let colWidth c = Map.tryFind c model.ColumnWidths |> Option.defaultValue 100.0
    let rowHeight r = Map.tryFind r model.RowHeights |> Option.defaultValue 25.0
    let totalContentW = allCols |> List.sumBy colWidth
    let totalContentH = allRows |> List.sumBy rowHeight
    let fontSize = 10.0
    let separator = "#000000"
    let colLetterItems =
        let mutable x = headW
        [ for c in allCols do
            let w = colWidth c
            let rect = { X = x; Y = 0.0; Width = w; Height = headH }
            yield headingText (columnToLetter c) rect AlignCenter fontSize
            x <- x + w ]
    let rowNumberItems =
        let mutable y = headH
        [ for r in allRows do
            let h = rowHeight r
            let rect = { X = 0.0; Y = y; Width = headW; Height = h }
            yield headingText (string (r + 1)) rect AlignCenter fontSize
            y <- y + h ]
    let separatorBelow =
        GridLine (headW, headH, headW + totalContentW, headH, separator, 1.0)
    let separatorRight =
        GridLine (headW, headH, headW, headH + totalContentH, separator, 1.0)
    // Outer edges of the heading chrome. Excel prints the blank top-left
    // corner cell as a real outlined cell; without these four short strokes
    // the header strips look broken even though the body gridlines are fine.
    let topEdge =
        GridLine (headW, 0.0, headW + totalContentW, 0.0, separator, 1.0)
    let leftEdge =
        GridLine (0.0, headH, 0.0, headH + totalContentH, separator, 1.0)
    let cornerEdges =
        [ GridLine (0.0, 0.0, headW, 0.0, separator, 1.0)
          GridLine (0.0, headH, headW, headH, separator, 1.0)
          GridLine (0.0, 0.0, 0.0, headH, separator, 1.0)
          GridLine (headW, 0.0, headW, headH, separator, 1.0) ]
    // Per-cell dividers inside each strip so the chrome reads as a row of
    // labelled cells rather than a single label band. Excel paints these
    // even when PrintOptions.GridLines is off.
    let colLetterDividers =
        let mutable x = headW
        [ for c in allCols do
            x <- x + colWidth c
            yield GridLine (x, 0.0, x, headH, separator, 1.0) ]
    let rowNumberDividers =
        let mutable y = headH
        [ for r in allRows do
            y <- y + rowHeight r
            yield GridLine (0.0, y, headW, y, separator, 1.0) ]
    colLetterItems @ rowNumberItems
    @ [ separatorBelow; separatorRight; topEdge; leftEdge ]
    @ cornerEdges
    @ colLetterDividers @ rowNumberDividers

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
    (section: HeaderFooter.StyledSection) (defaultFontSizePt: float) (textScale: float) (baselineYPt: float)
    (xAlign: TextAlign) : PdfDraw.DrawOp list =
    if String.IsNullOrEmpty section.Text then [] else
    // Resolve face / size / color per H/F run. Excel lets section strings
    // switch style mid-run with codes like `&B`, `&"Arial,Italic"`, `&12`,
    // `&KFF0000`, `&U`, `&S`, `&X`, `&Y`; alignment still applies to the
    // total section width. Super/subscript shrinks the run's font size
    // to ~58% of the base; underline / strikethrough render as strokes.
    let segmentsWithWidth =
        [ for run in section.Runs do
            if not (String.IsNullOrEmpty run.Text) then
                // Excel scales H/F text by the page's print scale (the
                // `&NN` size code carries the *un*scaled value). Apply the
                // scale uniformly here so the &NN baseline still works.
                let rawSizePt = run.FontSize |> Option.defaultValue defaultFontSizePt
                let baseSizePt = rawSizePt * textScale
                let scale =
                    if run.Superscript || run.Subscript then 0.58 else 1.0
                let fontSizePt = baseSizePt * scale
                let baselineDeltaPt =
                    if run.Superscript then baseSizePt * 0.33
                    elif run.Subscript then -baseSizePt * 0.20
                    else 0.0
                match pickFace options run.FontName run.Bold run.Italic with
                | None -> ()
                | Some primaryFace ->
                    match Map.tryFind primaryFace.FaceName faceByName with
                    | None -> ()
                    | Some primaryTtf ->
                        let segments = segmentByFont (primaryFace, primaryTtf) options.CjkFallback faceByName run.Bold run.Text
                        for (face, s) in segments do
                            let ttf = snd face
                            let w = float (PdfFont.getTextWidth ttf s) / float ttf.UnitsPerEm * fontSizePt
                            yield face, s, w, fontSizePt, run.Color, run.Underline, run.DoubleUnderline, run.Strikethrough, baselineDeltaPt ]
    if List.isEmpty segmentsWithWidth then [] else
    let totalWidthPt = segmentsWithWidth |> List.sumBy (fun (_, _, w, _, _, _, _, _, _) -> w)
    let startXPt =
        match xAlign with
        | AlignLeft -> leftPt
        | AlignCenter -> leftPt + ((rightPt - leftPt) - totalWidthPt) / 2.0
        | AlignRight -> rightPt - totalWidthPt
    // Track per-segment decoration spans so we can emit underline /
    // strikethrough strokes after the text ops (PDF needs path ops outside
    // any text object).
    let decorations =
        let acc = System.Collections.Generic.List<float * float * float * float * bool * bool * bool * string option>()
        let mutable cursor = startXPt
        for (_, _, w, fontSizePt, color, ul, du, st, _) in segmentsWithWidth do
            acc.Add (cursor, w, fontSizePt, 0.0, ul, du, st, color)
            cursor <- cursor + w
        acc |> List.ofSeq
    let textOps =
        [ yield PdfDraw.BeginText
          let mutable cursorXPt = startXPt
          for (face, s, w, fontSizePt, color, _, _, _, baselineDelta) in segmentsWithWidth do
              let faceName = (fst face).FaceName
              let _, ttf = face
              let remap = faceRemap.TryFind faceName |> Option.defaultValue Map.empty
              let hex = toRemappedHex ttf remap s
              yield PdfDraw.SetFillColor (color |> Option.map parseHexColor |> Option.defaultValue PdfColor.black)
              yield PdfDraw.SetFont (faceName, fontSizePt)
              yield PdfDraw.TextMatrix (1.0, 0.0, 0.0, 1.0, cursorXPt, baselineYPt + baselineDelta)
              yield PdfDraw.ShowTextHex hex
              cursorXPt <- cursorXPt + w
          yield PdfDraw.EndText ]
    let decorationOps =
        [ for (x, w, fontSizePt, _, ul, du, st, color) in decorations do
            let strokeColor = color |> Option.map parseHexColor |> Option.defaultValue PdfColor.black
            let lineWidth = max 0.5 (fontSizePt * 0.05)
            if ul then
                let yUL = baselineYPt - fontSizePt * 0.12
                yield PdfDraw.SetStrokeColor strokeColor
                yield PdfDraw.SetLineWidth lineWidth
                yield PdfDraw.SetDash ([], 0.0)
                yield PdfDraw.MoveTo (x, yUL)
                yield PdfDraw.LineTo (x + w, yUL)
                yield PdfDraw.Stroke
            if du then
                let yU1 = baselineYPt - fontSizePt * 0.08
                let yU2 = baselineYPt - fontSizePt * 0.18
                yield PdfDraw.SetStrokeColor strokeColor
                yield PdfDraw.SetLineWidth lineWidth
                yield PdfDraw.SetDash ([], 0.0)
                yield PdfDraw.MoveTo (x, yU1)
                yield PdfDraw.LineTo (x + w, yU1)
                yield PdfDraw.Stroke
                yield PdfDraw.MoveTo (x, yU2)
                yield PdfDraw.LineTo (x + w, yU2)
                yield PdfDraw.Stroke
            if st then
                let yST = baselineYPt + fontSizePt * 0.28
                yield PdfDraw.SetStrokeColor strokeColor
                yield PdfDraw.SetLineWidth lineWidth
                yield PdfDraw.SetDash ([], 0.0)
                yield PdfDraw.MoveTo (x, yST)
                yield PdfDraw.LineTo (x + w, yST)
                yield PdfDraw.Stroke ]
    textOps @ decorationOps

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
    let model = model |> withCommentMarkersForPrint |> withPrintColumnWidths |> withPrintRowHeights
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
                    float widthUnits / float ttf.UnitsPerEm * fontPt * 96.0 / 72.0)
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
    let printGridLines = ps.Options.GridLines
    let printHeadings = ps.Options.Headings
    // When PrintOptions.Headings is on, every content band shifts right /
    // down by these to make room for the column-letter strip on top and
    // the row-number strip on the left. Pagination has already shrunk the
    // per-page area by the same amounts, so manual breaks land in the
    // right place.
    let headingsColPx, headingsRowPx = Pagination.headingsOffsets ps
    let assembleBand (rowStart: int) (rowEnd: int) (colStart: int) (colEnd: int) (offsetX: float) (offsetY: float) =
        let content =
            sliceItems model measurePdf rowStart rowEnd colStart colEnd offsetX offsetY
        if printGridLines then
            let gridlines = gridLineItems model rowStart rowEnd colStart colEnd offsetX offsetY
            withGridlines gridlines content
        else content
    let pageItems : PaintItem list array =
        let pages = pagination.Pages
        Array.init (max 1 (List.length pages)) (fun i ->
            if i >= List.length pages then [] else
            let slice = pages.[i]
            // Body: rows × cols of the slice, offset by repeat band sizes.
            let body =
                assembleBand
                    slice.RowStart slice.RowEnd
                    slice.ColStart slice.ColEnd
                    (headingsColPx + repeatColsWidthPx) (headingsRowPx + repeatRowsHeightPx)
            // Repeat-rows band (above body, full width): repeat rows × slice cols.
            let repeatRows =
                match pagination.RepeatRows with
                | Some (rs, re) when re >= rs ->
                    assembleBand
                        rs re
                        slice.ColStart slice.ColEnd
                        (headingsColPx + repeatColsWidthPx) headingsRowPx
                | _ -> []
            // Repeat-cols band (left of body, full height): slice rows × repeat cols.
            let repeatCols =
                match pagination.RepeatCols with
                | Some (cs, ce) when ce >= cs ->
                    assembleBand
                        slice.RowStart slice.RowEnd
                        cs ce
                        headingsColPx (headingsRowPx + repeatRowsHeightPx)
                | _ -> []
            // Repeat corner (top-left): repeat rows × repeat cols.
            let repeatCorner =
                match pagination.RepeatRows, pagination.RepeatCols with
                | Some (rs, re), Some (cs, ce) when re >= rs && ce >= cs ->
                    assembleBand rs re cs ce headingsColPx headingsRowPx
                | _ -> []
            let headings =
                if printHeadings then headingItems model slice pagination else []
            body @ repeatRows @ repeatCols @ repeatCorner @ headings)

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
    let recordStyled (section: HeaderFooter.StyledSection) =
        for run in section.Runs do
            if not (String.IsNullOrEmpty run.Text) then
                match pickFace options run.FontName run.Bold run.Italic with
                | None -> ()
                | Some primaryFace ->
                    match parsed.TryGetValue primaryFace.FaceName with
                    | true, primaryTtf ->
                        let segments = segmentByFont (primaryFace, primaryTtf) options.CjkFallback parsedMap run.Bold run.Text
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

    // Pick the (header, footer) variant for a 1-indexed page number. When
    // DifferentFirst is set, page 1 uses FirstHeader/FirstFooter; when
    // DifferentOddEven is set, even pages use EvenHeader/EvenFooter; all
    // remaining pages fall through to OddHeader/OddFooter. A flag whose
    // variant slot is None renders nothing on that page — matches Excel's
    // behavior (DifferentOddEven=1 + EvenHeader missing = blank even pages).
    let pickHeaderFooter (pageNum1: int) : string option * string option =
        let hf = ps.HeaderFooter
        if hf.DifferentFirst && pageNum1 = 1 then hf.FirstHeader, hf.FirstFooter
        elif hf.DifferentOddEven && pageNum1 % 2 = 0 then hf.EvenHeader, hf.EvenFooter
        else hf.OddHeader, hf.OddFooter

    // Build header/footer page contexts and pre-resolve all text so we
    // can register its glyphs for subsetting before embedding fonts.
    // `&P` honors `firstPageNumber` when `useFirstPageNumber` is set;
    // `&N` keeps reporting the physical page count (Excel doesn't add
    // the first-page offset to N).
    let firstPageOffset =
        if ps.UseFirstPageNumber then ps.FirstPageNumber - 1 else 0
    let resolvedHFs =
        [ for pageIdx in 0 .. totalPages - 1 ->
            let ctx : HeaderFooter.PageContext =
                { PageNumber = pageIdx + 1 + firstPageOffset
                  TotalPages = totalPages
                  SheetName = options.SheetName
                  FileName = options.FileName
                  FilePath = ""
                  Now = DateTime.Now }
            let headerRaw, footerRaw = pickHeaderFooter (pageIdx + 1)
            let header = headerRaw |> Option.map (HeaderFooter.resolve ctx)
            let footer = footerRaw |> Option.map (HeaderFooter.resolve ctx)
            header, footer ]
    for (header, footer) in resolvedHFs do
        for section in [ header; footer ] |> List.choose id do
            recordStyled section.Left
            recordStyled section.Center
            recordStyled section.Right

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

    // Header/footer ops live in PDF coords directly. Excel scales H/F text
    // by the page's print scale when `headerFooter/@scaleWithDoc` is true
    // (its default). When false, the H/F text stays at its raw size even
    // though the body is scaled. `hfFontSizePt` is the raw fallback size;
    // `hfTextScale` is the multiplier we apply to every H/F glyph size.
    // Excel's print PDF uses Calibri 11 (workbook default) for unstyled
    // H/F runs — verified by glyph cap-height pixel-counting against
    // synthetic-print-chrome ("FIRST" renders 9 px tall in Excel = 11pt).
    let hfFontSizePt = 11.0
    let hfTextScale =
        if ps.HeaderFooter.ScaleWithDoc then pagination.Scale else 1.0
    let hfScaledFontSizePt = hfFontSizePt * hfTextScale
    let printableLeftPt = leftMarginPt
    let printableRightPt = pageWidthPt - rightMarginPt
    let maxSectionFontSize (section: HeaderFooter.StyledSection) =
        section.Runs
        |> List.filter (fun run -> not (String.IsNullOrEmpty run.Text))
        |> List.map (fun run -> (run.FontSize |> Option.defaultValue hfFontSizePt) * hfTextScale)
        |> List.fold max hfScaledFontSizePt
    let maxHeaderFooterFontSize (section: HeaderFooter.Section) =
        [ section.Left; section.Center; section.Right ]
        |> List.map maxSectionFontSize
        |> List.fold max hfScaledFontSizePt
    let headerBaselineYPt (section: HeaderFooter.Section) =
        // Excel anchors header text from the top header margin. Large H/F
        // runs therefore push the baseline down; a fixed 10 pt baseline puts
        // 16 pt runs visibly too high. Excel leaves a little leading above
        // larger text, so don't subtract the full max size.
        // The -0.75 pt nudge (= 1 px at 96 dpi) drops the baseline so our
        // glyph tops align with Excel's print-PDF header position. Verified
        // by pixel-counting "FIRST" in synthetic-print-chrome: Excel's
        // glyph tops at y=35, ours used to sit at y=34 without this nudge.
        let fontSizePt = maxHeaderFooterFontSize section
        let anchoredSizePt =
            if fontSizePt > hfScaledFontSizePt then max hfScaledFontSizePt (fontSizePt - 1.5)
            else hfScaledFontSizePt
        pageHeightPt - headerMarginPt - anchoredSizePt - 0.75
    let footerBaselineYPt =
        // Excel places footer text a little higher than the raw footer
        // margin baseline. A fixed lift matches both normal 10 pt footers
        // and the synthetic 16 pt mixed-run footer crop.
        footerMarginPt + 4.5
    let renderSection (section: HeaderFooter.Section) (baselineYPt: float) =
        sectionOps options faceTtfMap faceResourceMap faceRemapMap
            printableLeftPt printableRightPt
            section.Left hfFontSizePt hfTextScale baselineYPt AlignLeft
        @ sectionOps options faceTtfMap faceResourceMap faceRemapMap
            printableLeftPt printableRightPt
            section.Center hfFontSizePt hfTextScale baselineYPt AlignCenter
        @ sectionOps options faceTtfMap faceResourceMap faceRemapMap
            printableLeftPt printableRightPt
            section.Right hfFontSizePt hfTextScale baselineYPt AlignRight

    // Emit one PDF page per page slice (or one blank page when the
    // pagination is empty).
    let printableWidthPt = pageWidthPt - leftMarginPt - rightMarginPt
    let printableHeightPt = pageHeightPt - topMarginPt - bottomMarginPt
    for pageIdx in 0 .. totalPages - 1 do
        let items = pageItems.[pageIdx]
        // Per-page content extent in points, accounting for heading strips
        // and repeat bands. Excel centers each page individually against its
        // own filled area, so a partly-filled last page still centers on
        // what's actually drawn rather than the full printable rectangle.
        let leftPtOffset, topPtOffset =
            if pagination.Pages.IsEmpty then leftMarginPt, topMarginPt
            else
                let slice = pagination.Pages.[pageIdx]
                let bodyWidthPx =
                    headingsColPx + repeatColsWidthPx
                    + colSumPx model slice.ColStart slice.ColEnd
                let bodyHeightPx =
                    headingsRowPx + repeatRowsHeightPx
                    + rowSumPx model slice.RowStart slice.RowEnd
                let bodyWidthPt = bodyWidthPx * pxToPt * pagination.Scale
                let bodyHeightPt = bodyHeightPx * pxToPt * pagination.Scale
                let extraLeft =
                    if ps.Options.HorizontalCentered && bodyWidthPt < printableWidthPt
                    then (printableWidthPt - bodyWidthPt) / 2.0
                    else 0.0
                let extraTop =
                    if ps.Options.VerticalCentered && bodyHeightPt < printableHeightPt
                    then (printableHeightPt - bodyHeightPt) / 2.0
                    else 0.0
                leftMarginPt + extraLeft, topMarginPt + extraTop
        let placement =
            { LeftPtOffset = leftPtOffset
              TopPtOffset = topPtOffset
              PageHeightPt = pageHeightPt
              Scale = pagination.Scale
              BlackAndWhite = ps.BlackAndWhite }
        let contentOps =
            items |> List.collect (translateItem options faceTtfMap faceResourceMap faceRemapMap placement)
        let header, footer = resolvedHFs.[pageIdx]
        let headerOps =
            header |> Option.map (fun s -> renderSection s (headerBaselineYPt s)) |> Option.defaultValue []
        let footerOps =
            footer |> Option.map (fun s -> renderSection s footerBaselineYPt) |> Option.defaultValue []
        let ops = contentOps @ headerOps @ footerOps
        PdfDoc.addPage builder (pageWidthPt, pageHeightPt) loadedFaceNames [] ops |> ignore

    PdfDoc.toBytes builder

/// Render a standalone vector PaintItem scene (no model/text) to a one-page
/// PDF — proves the VecPath/VecGroup primitives through the PDF path.
let sceneToBytes (widthPx: float) (heightPx: float) (items: PaintItem list) : byte[] =
    let builder = PdfDoc.create ()
    let pageWidthPt = widthPx * pxToPt
    let pageHeightPt = heightPx * pxToPt
    let placement =
        { LeftPtOffset = 0.0; TopPtOffset = 0.0; PageHeightPt = pageHeightPt
          Scale = 1.0; BlackAndWhite = false }
    let opts = { DefaultFontFamily = "Calibri"; Fonts = []; CjkFallback = []; SheetName = ""; FileName = "" }
    let ops = items |> List.collect (translateItem opts Map.empty Map.empty Map.empty placement)
    PdfDoc.addPage builder (pageWidthPt, pageHeightPt) [] [] ops |> ignore
    PdfDoc.toBytes builder
