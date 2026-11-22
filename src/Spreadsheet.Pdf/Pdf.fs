module Rbt.Spreadsheet.Pdf.Render

open System
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Layout
open Rbt.Spreadsheet.Paint

module PdfDoc = Rbt.Pdf.Doc
module PdfDraw = Rbt.Pdf.Draw
module PdfColor = Rbt.Pdf.Color
module PdfFont = Rbt.Pdf.Font

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
    let hexDigit ch =
        if ch >= '0' && ch <= '9' then int ch - int '0'
        elif ch >= 'A' && ch <= 'F' then int ch - int 'A' + 10
        elif ch >= 'a' && ch <= 'f' then int ch - int 'a' + 10
        else -1
    let byteAt index =
        let hi = hexDigit color.[index]
        let lo = hexDigit color.[index + 1]
        if hi < 0 || lo < 0 then ValueNone else ValueSome ((hi <<< 4) ||| lo)
    let start = if color.Length > 0 && color.[0] = '#' then 1 else 0
    let length = color.Length - start
    let first = if length = 8 then start + 2 else start
    if length <> 6 && length <> 8 then PdfColor.black
    else
        match byteAt first, byteAt (first + 2), byteAt (first + 4) with
        | ValueSome r, ValueSome g, ValueSome b -> PdfColor.rgb (float r / 255.0) (float g / 255.0) (float b / 255.0)
        | _ -> PdfColor.black

/// CSS pixels → PDF points. Excel widths flow in CSS px (1 px = 72/96 pt).
let private pxToPt = 72.0 / 96.0

let private px (v: float) = v * pxToPt

let private resizeArrayToList (items: ResizeArray<'T>) =
    let mutable result = []
    for i = items.Count - 1 downto 0 do
        result <- items.[i] :: result
    result

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
    (tryGetFace: FontFace -> PdfFont.TrueTypeFont option)
    (bold: bool)
    (cp: int)
    : (FontFace * PdfFont.TrueTypeFont) option =
    let hasRequestedWeight = chain |> List.exists (fun f -> f.Bold = bold)
    let rec loop remaining =
        match remaining with
        | [] -> None
        | f :: rest ->
            if hasRequestedWeight && f.Bold <> bold then loop rest
            else
                match tryGetFace f with
                | Some ttf when ttf.Cmap.ContainsKey cp -> Some (f, ttf)
                | _ -> loop rest
    loop chain

let private fontCoversAllText (ttf: PdfFont.TrueTypeFont) (text: string) =
    let mutable covered = true
    let mutable i = 0
    while covered && i < text.Length do
        let ch = text.[i]
        let cp, width =
            if System.Char.IsHighSurrogate ch && i + 1 < text.Length && System.Char.IsLowSurrogate text.[i + 1] then
                0x10000 + ((int ch - 0xD800) <<< 10) + (int text.[i + 1] - 0xDC00), 2
            else int ch, 1
        if not (ttf.Cmap.ContainsKey cp) then
            covered <- false
        i <- i + width
    covered

let private segmentMixedByFont
    (primary: FontFace * PdfFont.TrueTypeFont)
    (cjk: FontFace list)
    (tryGetFace: FontFace -> PdfFont.TrueTypeFont option)
    (bold: bool)
    (text: string)
    : ((FontFace * PdfFont.TrueTypeFont) * string) list =
    let primaryTtf = snd primary
    let fontFor cp =
        if primaryTtf.Cmap.ContainsKey cp then primary
        else
            match pickCjkFallback cjk tryGetFace bold cp with
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
                0x10000 + ((int ch - 0xD800) <<< 10) + (int text.[i + 1] - 0xDC00), 2
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
    resizeArrayToList result

/// Split `text` into contiguous (font, substring) segments where each
/// substring's codepoints are covered by `font`'s cmap. Falls back through
/// `cjk` for missing glyphs; uses `primary` for anything `primary` itself
/// covers (and as a last resort when no fallback covers a codepoint).
/// Mirrors SpreadsheetGrid.toRendererCommands so PDF and PNG outputs draw
/// the same glyph stream for mixed-script text.
let private segmentByFont
    (primary: FontFace * PdfFont.TrueTypeFont)
    (cjk: FontFace list)
    (tryGetFace: FontFace -> PdfFont.TrueTypeFont option)
    (bold: bool)
    (text: string)
    : ((FontFace * PdfFont.TrueTypeFont) * string) list =
    if text.Length = 0 then []
    elif List.isEmpty cjk || fontCoversAllText (snd primary) text then [ primary, text ]
    else segmentMixedByFont primary cjk tryGetFace bold text

let private hexDigits = "0123456789ABCDEF".ToCharArray()

let private appendHex4 (sb: System.Text.StringBuilder) value =
    if value >= 0 && value <= 0xFFFF then
        sb.Append(hexDigits.[(value >>> 12) &&& 0xF]) |> ignore
        sb.Append(hexDigits.[(value >>> 8) &&& 0xF]) |> ignore
        sb.Append(hexDigits.[(value >>> 4) &&& 0xF]) |> ignore
        sb.Append(hexDigits.[value &&& 0xF]) |> ignore
    else
        sb.Append(value.ToString("X")) |> ignore

/// Re-encode a substring as hex glyph IDs, applying a per-face subset
/// remap when present. With an empty remap (face was full-embedded), the
/// output matches PdfFont.toGlyphHex. With a remap, original GIDs are
/// translated to the post-subset numbering; missing entries collapse to
/// GID 0 (.notdef), which prints as a missing-glyph box.
let private toRemappedHex (ttf: PdfFont.TrueTypeFont) (remap: Map<int, int>) (text: string) =
    let sb = System.Text.StringBuilder(text.Length * 4)
    let isEmpty = Map.isEmpty remap
    for ch in text do
        let origGid =
            match ttf.Cmap.TryFind (int ch) with
            | Some gid -> gid
            | None -> 0
        let gid =
            if isEmpty then origGid
            else
                match remap.TryFind origGid with
                | Some mapped -> mapped
                | None -> 0
        appendHex4 sb gid
    sb.ToString()

let private tryGetMappedFace (faceByName: Map<string, PdfFont.TrueTypeFont>) (face: FontFace) =
    Map.tryFind face.FaceName faceByName

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
    /// The name of a picture's image object in the document (embedding it at
    /// first use), or None for a picture that cannot be embedded: a labelled
    /// frame is drawn for that one.
    ImageName: PaintImage -> string option
}

/// usWinAscent / usWinDescent from the font's OS/2 table, in font units.
/// Rbt.Pdf's TrueTypeFont carries only the hhea metrics, so they are read from
/// the font data here; a font without a usable OS/2 table falls back to hhea.
let private winMetrics (ttf: PdfFont.TrueTypeFont) : float * float =
    let readUInt16 offset = (int ttf.Data.[offset] <<< 8) ||| int ttf.Data.[offset + 1]
    match Map.tryFind "OS/2" ttf.Tables with
    | Some table when table.Length >= 78u && int table.Offset + 78 <= ttf.Data.Length ->
        let start = int table.Offset
        float (readUInt16 (start + 74)), float (readUInt16 (start + 76))
    | _ -> float ttf.Ascender, float (abs ttf.Descender)

let private bwFill (bw: bool) (color: string) =
    if bw then "#FFFFFF" else color

let private bwStroke (bw: bool) (color: string) =
    if bw then "#000000" else color

[<Struct>]
type private HeaderFooterSegment = {
    Face: FontFace
    Ttf: PdfFont.TrueTypeFont
    Text: string
    WidthPt: float
    FontSizePt: float
    Color: string option
    Underline: bool
    DoubleUnderline: bool
    Strikethrough: bool
    BaselineDeltaPt: float
}

let private appendDrawTextOps
    (options: RenderOptions)
    (faceByName: Map<string, PdfFont.TrueTypeFont>)
    (faceRemap: Map<string, Map<int, int>>)
    (placement: Placement)
    (ops: ResizeArray<PdfDraw.DrawOp>)
    (t: PaintText) =
    match pickFace options t.FontName t.Bold t.Italic with
    | None -> ()
    | Some primaryFace ->
        match Map.tryFind primaryFace.FaceName faceByName with
        | None -> ()
        | Some primaryTtf ->
            let sz v = v * pxToPt * placement.Scale
            let xPt x = placement.LeftPtOffset + sz x
            let yPdf y = placement.PageHeightPt - placement.TopPtOffset - sz y
            // PaintText.FontSize is stored in points. Draw in PDF points,
            // but use CSS-pixel metrics for layout because cell geometry is
            // still page-local CSS px before translateItem applies pxToPt.
            let fontPx = t.FontSize * 96.0 / 72.0
            let fontSizePt = t.FontSize * placement.Scale
            let singlePrimarySegment =
                t.Text.Length > 0
                && (List.isEmpty options.CjkFallback || fontCoversAllText primaryTtf t.Text)
            let segments =
                if singlePrimarySegment then []
                else segmentMixedByFont (primaryFace, primaryTtf) options.CjkFallback (tryGetMappedFace faceByName) t.Bold t.Text
            let segmentWidthScale (ttf: PdfFont.TrueTypeFont) s =
                float (PdfFont.getTextWidth ttf s) / float ttf.UnitsPerEm
            let mutable totalWidthPt = 0.0
            let mutable totalWidthPx = 0.0
            if singlePrimarySegment then
                let widthScale = segmentWidthScale primaryTtf t.Text
                totalWidthPt <- totalWidthPt + widthScale * fontSizePt
                totalWidthPx <- totalWidthPx + widthScale * fontPx
            else
                for (face, s) in segments do
                    let ttf = snd face
                    let widthScale = segmentWidthScale ttf s
                    totalWidthPt <- totalWidthPt + widthScale * fontSizePt
                    totalWidthPx <- totalWidthPx + widthScale * fontPx
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
            let winAscent, winDescent = winMetrics primaryTtf
            let ascent = winAscent * scaleY
            let descent = winDescent * scaleY
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
            // rotated bounding box - RIGHT-aligned when 90 <= textRotation <
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
                    let ax0 = t.Rect.X + 4.0
                    let ay0 = t.Rect.Y + t.Rect.Height - 2.0
                    let rotPt x y =
                        let dx, dy = x - ax0, y - ay0
                        ax0 + dx * cosS - dy * sinS, ay0 + dx * sinS + dy * cosS
                    let c1x, c1y = rotPt ax0 (ay0 - ascent)
                    let c2x, c2y = rotPt (ax0 + totalWidthPx) (ay0 - ascent)
                    let c3x, c3y = rotPt ax0 (ay0 + descent)
                    let c4x, c4y = rotPt (ax0 + totalWidthPx) (ay0 + descent)
                    let minX = min (min c1x c2x) (min c3x c4x)
                    let maxX = max (max c1x c2x) (max c3x c4x)
                    let maxY = max (max c1y c2y) (max c3y c4y)
                    let rightAlign = t.Rotation >= 90 && t.Rotation < 180
                    let shiftX =
                        if rightAlign then (t.Rect.X + t.Rect.Width - 2.0) - maxX
                        else (t.Rect.X + 2.0) - minX
                    let shiftY = (t.Rect.Y + t.Rect.Height - 2.0) - maxY
                    xPt (ax0 + shiftX), yPdf (ay0 + shiftY)
            ops.Add PdfDraw.BeginText
            ops.Add(PdfDraw.SetFillColor textColor)
            let emitTextSegment faceName ttf s cumWidthPt =
                let remap =
                    match faceRemap.TryFind faceName with
                    | Some found -> found
                    | None -> Map.empty
                let hex = toRemappedHex ttf remap s
                ops.Add(PdfDraw.SetFont (faceName, fontSizePt))
                // Position this segment's anchor in PDF coords.
                // Without rotation: anchor + (cumWidth, 0).
                // With rotation: anchor + cumWidth * (cos theta, sin theta).
                let xAt = anchorXPt + cosT * cumWidthPt
                let yAt = anchorYPt + sinT * cumWidthPt
                ops.Add(PdfDraw.TextMatrix (cosT, sinT, -sinT, cosT, xAt, yAt))
                ops.Add(PdfDraw.ShowTextHex hex)
            let mutable cumWidthPt = 0.0
            if singlePrimarySegment then
                emitTextSegment primaryFace.FaceName primaryTtf t.Text 0.0
            else
                let mutable remainingSegments = segments
                while not (List.isEmpty remainingSegments) do
                    let face, s, nextSegments =
                        match remainingSegments with
                        | (face, s) :: rest -> face, s, rest
                        | [] -> failwith "unreachable"
                    let faceInfo, ttf = face
                    emitTextSegment faceInfo.FaceName ttf s cumWidthPt
                    if not (List.isEmpty nextSegments) then
                        cumWidthPt <- cumWidthPt + segmentWidthScale ttf s * fontSizePt
                    remainingSegments <- nextSegments
            ops.Add PdfDraw.EndText
            if t.Underline && totalWidthPt > 0.0 then
                let underlineY = baselineY + max 1.0 (t.FontSize * 0.08)
                ops.Add(PdfDraw.SetStrokeColor textColor)
                ops.Add(PdfDraw.SetLineWidth (sz 1.0))
                ops.Add(PdfDraw.SetDash ([], 0.0))
                ops.Add(PdfDraw.MoveTo (alignedXPt, yPdf underlineY))
                ops.Add(PdfDraw.LineTo (alignedXPt + totalWidthPt, yPdf underlineY))
                ops.Add PdfDraw.Stroke
                if t.UnderlineDouble then
                    let underlineY2 = underlineY + max 2.0 (t.FontSize * 0.12)
                    ops.Add(PdfDraw.MoveTo (alignedXPt, yPdf underlineY2))
                    ops.Add(PdfDraw.LineTo (alignedXPt + totalWidthPt, yPdf underlineY2))
                    ops.Add PdfDraw.Stroke
            if t.Strikethrough && totalWidthPt > 0.0 then
                let strikeY = baselineY - t.FontSize * 0.28
                ops.Add(PdfDraw.SetStrokeColor textColor)
                ops.Add(PdfDraw.SetLineWidth (sz 1.0))
                ops.Add(PdfDraw.SetDash ([], 0.0))
                ops.Add(PdfDraw.MoveTo (alignedXPt, yPdf strikeY))
                ops.Add(PdfDraw.LineTo (alignedXPt + totalWidthPt, yPdf strikeY))
                ops.Add PdfDraw.Stroke

/// A text item, cut off at its clip rectangle when it has one (text wider
/// than its cell: Excel's text overflow stops at the first filled cell).
let private appendClippedTextOps
    (options: RenderOptions)
    (faceByName: Map<string, PdfFont.TrueTypeFont>)
    (faceRemap: Map<string, Map<int, int>>)
    (placement: Placement)
    (ops: ResizeArray<PdfDraw.DrawOp>)
    (t: PaintText) =
    match t.Clip with
    | Some clip when clip.Width > 0.0 && clip.Height > 0.0 ->
        let sz v = v * pxToPt * placement.Scale
        let xPt x = placement.LeftPtOffset + sz x
        let yPdf y = placement.PageHeightPt - placement.TopPtOffset - sz y
        ops.Add PdfDraw.SaveState
        ops.Add(PdfDraw.Rect (xPt clip.X, yPdf (clip.Y + clip.Height), sz clip.Width, sz clip.Height))
        ops.Add PdfDraw.ClipPath
        ops.Add PdfDraw.EndPath
        appendDrawTextOps options faceByName faceRemap placement ops t
        ops.Add PdfDraw.RestoreState
    | _ -> appendDrawTextOps options faceByName faceRemap placement ops t

/// Convert one Paint item into a list of PDF draw ops. Coordinates flip from
/// our top-left CSS-pixel space into PDF's bottom-left point space, with the
/// `Placement`'s margin offsets and scale baked in.
let rec private translateItem (options: RenderOptions) (faceByName: Map<string, PdfFont.TrueTypeFont>) (faceResource: Map<string, string>) (faceRemap: Map<string, Map<int, int>>) (placement: Placement) (item: PaintItem) : PdfDraw.DrawOp list =
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
    | Clipped (rect, children) ->
        [ PdfDraw.SaveState
          PdfDraw.Rect (xPt rect.X, yPdf (rect.Y + rect.Height), sz rect.Width, sz rect.Height)
          PdfDraw.ClipPath
          PdfDraw.EndPath ]
        @ (children |> List.collect (translateItem options faceByName faceResource faceRemap placement))
        @ [ PdfDraw.RestoreState ]
    // A picture (PNG or JPEG) is embedded and stretched over its rectangle; a
    // labelled frame stands for one in a format the PDF cannot hold.
    | DrawImage (rect, image) ->
        match (if rect.Width > 0.0 && rect.Height > 0.0 then placement.ImageName image else None) with
        | Some name ->
            [ PdfDraw.SaveState
              // An image fills the unit square: scale it to the rectangle, its lower-left corner at the rectangle's.
              PdfDraw.Transform (sz rect.Width, 0.0, 0.0, sz rect.Height, xPt rect.X, yPdf (rect.Y + rect.Height))
              PdfDraw.DrawXObject name
              PdfDraw.RestoreState ]
        | None ->
            drawingFrameItems rect image.Label |> List.collect (translateItem options faceByName faceResource faceRemap placement)
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
        let ops = ResizeArray<PdfDraw.DrawOp>()
        appendClippedTextOps options faceByName faceRemap placement ops t
        resizeArrayToList ops

let private appendTranslatedItem (options: RenderOptions) (faceByName: Map<string, PdfFont.TrueTypeFont>) (faceResource: Map<string, string>) (faceRemap: Map<string, Map<int, int>>) (placement: Placement) (ops: ResizeArray<PdfDraw.DrawOp>) (item: PaintItem) =
    let sz v = v * pxToPt * placement.Scale
    let xPt x = placement.LeftPtOffset + sz x
    let yPdf y = placement.PageHeightPt - placement.TopPtOffset - sz y
    match item with
    | FillRect (rect, color) ->
        ops.Add(PdfDraw.SetFillColor (parseHexColor (bwFill placement.BlackAndWhite color)))
        ops.Add(PdfDraw.Rect (xPt rect.X, yPdf (rect.Y + rect.Height), sz rect.Width, sz rect.Height))
        ops.Add PdfDraw.Fill
    | StrokeRect (rect, color, width) ->
        ops.Add(PdfDraw.SetStrokeColor (parseHexColor (bwStroke placement.BlackAndWhite color)))
        ops.Add(PdfDraw.SetLineWidth (sz width))
        ops.Add(PdfDraw.SetDash ([], 0.0))
        ops.Add(PdfDraw.Rect (xPt rect.X, yPdf (rect.Y + rect.Height), sz rect.Width, sz rect.Height))
        ops.Add PdfDraw.Stroke
    | GridLine (x1, y1, x2, y2, color, width) ->
        ops.Add(PdfDraw.SetStrokeColor (parseHexColor (bwStroke placement.BlackAndWhite color)))
        ops.Add(PdfDraw.SetLineWidth (sz width))
        ops.Add(PdfDraw.SetDash ([], 0.0))
        ops.Add(PdfDraw.MoveTo (xPt x1, yPdf y1))
        ops.Add(PdfDraw.LineTo (xPt x2, yPdf y2))
        ops.Add PdfDraw.Stroke
    | DrawText t ->
        appendClippedTextOps options faceByName faceRemap placement ops t
    | _ ->
        for op in translateItem options faceByName faceResource faceRemap placement item do
            ops.Add op

/// The pictures embedded in a PDF document: each picture's data becomes one
/// image object, whatever the number of pages it is drawn on.
type private EmbeddedImages(builder: PdfDoc.DocBuilder) =
    let names = System.Collections.Generic.Dictionary<byte[], string option>(HashIdentity.Reference)
    let onPage = System.Collections.Generic.HashSet<string>()
    /// The image object of a picture (PNG and JPEG can be embedded).
    member _.NameOf (image: PaintImage) : string option =
        let name =
            match Portable.Dict.tryFind image.Data names with
            | ValueSome known -> known
            | ValueNone ->
                let isPng = image.Data.Length > 8 && image.Data.[0] = 0x89uy && image.Data.[1] = 0x50uy && image.Data.[2] = 0x4Euy && image.Data.[3] = 0x47uy
                let isJpeg = image.Data.Length > 3 && image.Data.[0] = 0xFFuy && image.Data.[1] = 0xD8uy
                let parsed =
                    try
                        // (Result.Error: the sheet model has an `Error` case of its own in scope.)
                        if isPng then (match Rbt.Pdf.Image.fromPng image.Data with Result.Ok info -> Some info | Result.Error _ -> None)
                        elif isJpeg then (match Rbt.Pdf.Image.fromJpeg image.Data with Result.Ok info -> Some info | Result.Error _ -> None)
                        else None
                    with _ -> None
                let made =
                    parsed
                    |> Option.map (fun info ->
                        let name = sprintf "Im%d" (names.Count + 1)
                        builder.Images <- builder.Images |> Map.add name (Rbt.Pdf.Image.toXObject builder.Doc info)
                        name)
                names.[image.Data] <- made
                made
        match name with
        | Some used -> onPage.Add used |> ignore
        | None -> ()
        name
    /// The images drawn since the last call: the page's image resources.
    member _.TakePageNames () : string list =
        let used = List.ofSeq onPage
        onPage.Clear()
        used

type private AxisGeometry = {
    Size: int -> float
    Sum: int -> int -> float
}

type private PrintGeometry = {
    Col: AxisGeometry
    Row: AxisGeometry
}

let private buildAxisGeometry defaultSize (sizes: Map<int, float>) maxIndex =
    let upper = max 0 maxIndex
    let prefix = Array.zeroCreate<float> (upper + 2)
    for i in 0 .. upper do
        prefix.[i + 1] <- prefix.[i] + (Map.tryFind i sizes |> Option.defaultValue defaultSize)
    let size i =
        if i >= 0 && i <= upper then prefix.[i + 1] - prefix.[i]
        else Map.tryFind i sizes |> Option.defaultValue defaultSize
    let sum startInc endInc =
        if startInc > endInc || endInc < 0 then 0.0
        elif startInc >= 0 && endInc <= upper then prefix.[endInc + 1] - prefix.[startInc]
        else
            let mutable s = 0.0
            for i in max 0 startInc .. endInc do
                s <- s + size i
            s
    { Size = size; Sum = sum }

let private buildPrintGeometry (model: SpreadsheetModel) maxRow maxCol =
    { Col = buildAxisGeometry model.DefaultColumnWidth model.ColumnWidths maxCol
      Row = buildAxisGeometry model.DefaultRowHeight model.RowHeights maxRow }

/// Excel's live grid uses the floored screen MDW for explicit `<col width>`
/// values, but its PDF print path lays those explicit widths out with the
/// non-floored Calibri-11 width (about 7.5 px instead of 7 px). Keep the
/// screen model untouched and apply that delta only while rendering PDF
/// pages. Seeded default-width columns are left alone; the default width is
/// inferred as the most common width in the model's column map.
/// Excel's print default for cell comments is "none" — comments don't
/// appear in the printout regardless of the on-screen red marker. Keep the
/// model immutable here; the print slice emitter drops the marker strokes
/// instead of cloning large cell maps to clear `HasComment` flags.
let private withCommentMarkersForPrint (model: SpreadsheetModel) : SpreadsheetModel =
    model

let private directPrintDimensionRemapLimit = 100000

// (The sizes rows and columns print at are Pagination's: a host is told the same pages.)
let private printColumnWidthPx mdw width = Pagination.printColumnWidthPx mdw width

let private printRowHeightPx px = Pagination.printRowHeightPx px

let private printDimensionBounds
    (model: SpreadsheetModel)
    (overlays: (int * int * float * float * float * float * PaintItem list) list) =
    let mutable maxCol = 0
    let mutable maxRow = 0
    match printAreasOf model.PageSetup with
    | _ :: _ as areas ->
        for printRange in areas do
            maxCol <- max maxCol (max printRange.From.Col printRange.To.Col)
            maxRow <- max maxRow (max printRange.From.Row printRange.To.Row)
    | [] ->
        for KeyValue(addr, _) in model.Cells do
            if addr.Col > maxCol then maxCol <- addr.Col
            if addr.Row > maxRow then maxRow <- addr.Row
    match model.PageSetup.RepeatCols with
    | Some (a, b) -> maxCol <- max maxCol (max a b)
    | None -> ()
    match model.PageSetup.RepeatRows with
    | Some (a, b) -> maxRow <- max maxRow (max a b)
    | None -> ()
    for row, col, _, _, _, _, _ in overlays do
        maxRow <- max maxRow row
        maxCol <- max maxCol col
    Some maxCol, Some maxRow

let private remapPrintDimensionDirect maxIndex remap (values: Map<int, float>) =
    let mutable mapped = Map.empty
    for index in 0 .. max 0 maxIndex do
        match Map.tryFind index values with
        | Some value -> mapped <- Map.add index (remap value) mapped
        | None -> ()
    mapped

let private remapPrintColumnWidthsDirect mdw maxCol values =
    remapPrintDimensionDirect maxCol (printColumnWidthPx mdw) values

let private remapPrintRowHeightsDirect maxRow values =
    remapPrintDimensionDirect maxRow printRowHeightPx values

let private withPrintColumnWidths maxPrintCol (model: SpreadsheetModel) : SpreadsheetModel =
    if Map.isEmpty model.ColumnWidths then model else
    // Excel's live grid lays columns out with the floored screen MDW (7 px for
    // Calibri 11), but its PDF print path uses the non-floored width (~7.5 px),
    // i.e. ~15/14 wider — for EVERY column, including the workbook-default
    // width. Measured on `synthetic-diagonal-borders.xlsx`: a default 8.43-char
    // column prints at ~68.8 px vs our 64 px; 64 × 15/14 = 68.6 px. Earlier
    // this scale was applied only to explicit `<col>` widths, which left the
    // default columns ~5 px narrow and drifted every cell right of column A
    // leftward. Scale all positive widths.
    let widths =
        match maxPrintCol with
        | Some maxCol when maxCol >= 0 && maxCol <= directPrintDimensionRemapLimit ->
            remapPrintColumnWidthsDirect model.MaxDigitWidth maxCol model.ColumnWidths
        | _ ->
            model.ColumnWidths |> Map.map (fun _ width -> printColumnWidthPx model.MaxDigitWidth width)
    { model with ColumnWidths = widths; DefaultColumnWidth = printColumnWidthPx model.MaxDigitWidth model.DefaultColumnWidth }

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
let private withPrintRowHeights maxPrintRow (model: SpreadsheetModel) : SpreadsheetModel =
    // Calibri 12 print pitch is similarly off — editor 21 px → print
    // ~20.64 px (= 15.48 pt). pandas-test1 rows positioned 1 px per row
    // below Excel because we used 21 px uniform while Excel's quantized
    // 20/21/21/21/21/20 pattern averages to ~20.64 px.
    let heights =
        match maxPrintRow with
        | Some maxRow when maxRow >= 0 && maxRow <= directPrintDimensionRemapLimit ->
            remapPrintRowHeightsDirect maxRow model.RowHeights
        | _ ->
            model.RowHeights |> Map.map (fun _ px -> printRowHeightPx px)
    { model with RowHeights = heights; DefaultRowHeight = printRowHeightPx model.DefaultRowHeight }

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
    | VecPath _ | VecGroup _ | Clipped _ -> false
    | DrawImage (r, _) -> r.Width < 0.5 || r.Height < 0.5

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

let private commentMarkerStrokeKey (item: PaintItem) =
    match item with
    | GridLine (x1, y1, x2, y2, color, width) ->
        if color.Equals("#C00000", StringComparison.OrdinalIgnoreCase)
           && abs (width - 1.0) < 0.001
           && abs (y2 - y1) < 0.01
           && x2 >= x1
           && x2 - x1 <= 5.1 then
            Some (struct (
                int64 (Math.Round(x1 * 1000.0)),
                int64 (Math.Round(y1 * 1000.0)),
                int64 (Math.Round(x2 * 1000.0)),
                int64 (Math.Round(y2 * 1000.0))))
        else
            None
    | _ -> None

let private commentMarkerStrokeKeysForSlice
    (geometry: PrintGeometry)
    (model: SpreadsheetModel)
    (rowStart: int) (rowEnd: int)
    (colStart: int) (colEnd: int) =
    match model.PageSetup.CellComments with
    | CellCommentsAsDisplayed -> None
    | CellCommentsNone | CellCommentsAtEnd ->
        let keys = System.Collections.Generic.HashSet<struct (int64 * int64 * int64 * int64)>()
        for row in max 0 rowStart .. rowEnd do
            for col in max 0 colStart .. colEnd do
                let addr: CellAddress = { Row = row; Col = col }
                match Map.tryFind addr model.Cells with
                | Some cell when cell.Format.HasComment ->
                    let rightX = geometry.Col.Sum 0 col
                    let topY = geometry.Row.Sum 0 (row - 1)
                    let triSize = 5.0
                    for k in 0 .. int triSize - 1 do
                        let y = topY + float k + 0.5
                        let leftX = rightX - (triSize - float k)
                        keys.Add (struct (
                            int64 (Math.Round(leftX * 1000.0)),
                            int64 (Math.Round(y * 1000.0)),
                            int64 (Math.Round(rightX * 1000.0)),
                            int64 (Math.Round(y * 1000.0)))) |> ignore
                | _ -> ()
        Some keys

let private shouldDropCommentMarkerForPrint
    (commentMarkerKeys: System.Collections.Generic.HashSet<struct (int64 * int64 * int64 * int64)> option)
    (item: PaintItem) =
    match commentMarkerKeys, commentMarkerStrokeKey item with
    | Some keys, Some key -> keys.Contains key
    | _ -> false

/// Shift every coordinate of a paint item by (dx, dy) CSS px. Used to map
/// items from their absolute viewport position into page-local space.
let rec private shiftItem (dx: float) (dy: float) (item: PaintItem) : PaintItem =
    let shiftRect (r: Rect) = { r with X = r.X + dx; Y = r.Y + dy }
    let shiftSeg seg =
        let mv x y = x + dx, y + dy
        match seg with
        | VMove (x, y) -> let a, b = mv x y in VMove (a, b)
        | VLine (x, y) -> let a, b = mv x y in VLine (a, b)
        | VCubic (x1, y1, x2, y2, x, y) ->
            let a1, b1 = mv x1 y1
            let a2, b2 = mv x2 y2
            let a, b = mv x y
            VCubic (a1, b1, a2, b2, a, b)
        | VClose -> VClose
    match item with
    | FillRect (r, c) -> FillRect (shiftRect r, c)
    | StrokeRect (r, c, w) -> StrokeRect (shiftRect r, c, w)
    | DrawText t -> DrawText { t with Rect = shiftRect t.Rect; Clip = t.Clip |> Option.map shiftRect }
    | GridLine (x1, y1, x2, y2, c, w) -> GridLine (x1 + dx, y1 + dy, x2 + dx, y2 + dy, c, w)
    | VecPath (segs, fill, stroke) ->
        let shifted = ResizeArray<VecSeg>(segs.Length)
        for seg in segs do
            shifted.Add(shiftSeg seg)
        VecPath (resizeArrayToList shifted, fill, stroke)
    | VecGroup ((a, b, c, d, e, f), children) ->
        VecGroup ((a, b, c, d, e + dx, f + dy), children)
    | DrawImage (r, image) -> DrawImage (shiftRect r, image)
    | Clipped (r, children) -> Clipped (shiftRect r, children |> List.map (shiftItem dx dy))

/// Iterate the paint items for a contiguous (rowStart..rowEnd × colStart..colEnd)
/// slice of the model, shifted so the slice's top-left lands at page-local
/// (offsetX, offsetY). Setting HeaderHeight / RowHeaderWidth to 0 in the
/// viewport keeps the column-letter / row-number chrome out — we only want
/// the body for print.
let private iterSliceItems
    (geometry: PrintGeometry)
    (model: SpreadsheetModel)
    (measurer: TextMeasurer)
    (includeText: bool)
    (rowStart: int) (rowEnd: int)
    (colStart: int) (colEnd: int)
    (offsetX: float) (offsetY: float)
    (emit: PaintItem -> unit) : unit =
    // Land the viewport on the exact rowStart..rowEnd / colStart..colEnd
    // rectangle. `visibleRange` now lands `startRow` on whatever row
    // contains `ScrollY` (top-of-cell semantics), so scrolling to the top
    // of `rowStart` picks `rowStart` cleanly; `Width`/`Height` exactly
    // matching the slice's content extents lets the end loop walk to
    // `rowEnd+1` and stop.
    let scrollX = geometry.Col.Sum 0 (colStart - 1)
    let scrollY = geometry.Row.Sum 0 (rowStart - 1)
    let width = geometry.Col.Sum colStart colEnd
    let height = geometry.Row.Sum rowStart rowEnd
    let viewport =
        { ScrollX = scrollX
          ScrollY = scrollY
          Width = width
          Height = height
          HeaderHeight = 0.0
          RowHeaderWidth = 0.0
          Measurer = Some measurer }
    let dx = offsetX - scrollX
    let dy = offsetY - scrollY
    let commentMarkerKeys = commentMarkerStrokeKeysForSlice geometry model rowStart rowEnd colStart colEnd
    for item in buildPaintItems model viewport do
        if not (isDegenerate item)
           && not (isDefaultPrintCellFill item)
           && not (isInteractiveCellStroke item)
           && not (shouldDropCommentMarkerForPrint commentMarkerKeys item)
           && (includeText || match item with DrawText _ -> false | _ -> true) then
            emit (shiftItem dx dy item)

/// Build the paint items for a print slice.
let private sliceItems
    (geometry: PrintGeometry)
    (model: SpreadsheetModel)
    (measurer: TextMeasurer)
    (rowStart: int) (rowEnd: int)
    (colStart: int) (colEnd: int)
    (offsetX: float) (offsetY: float)
    : PaintItem list =
    let rowCount = max 0 (rowEnd - rowStart + 1)
    let colCount = max 0 (colEnd - colStart + 1)
    let shifted = ResizeArray<PaintItem>(max 16 (rowCount * colCount + rowCount + colCount + 2))
    iterSliceItems geometry model measurer true rowStart rowEnd colStart colEnd offsetX offsetY shifted.Add
    resizeArrayToList shifted

/// Near-black cell-boundary lines for a slice, used when
/// `PrintOptions.GridLines` is on. Emits one vertical line at every
/// column edge (left edges of `colStart..colEnd` plus the right edge of
/// `colEnd`) and one horizontal at every row edge in the same fashion.
/// Coordinates are page-local CSS px shifted by `(offsetX, offsetY)`.
let private iterGridLineItems
    (geometry: PrintGeometry)
    (model: SpreadsheetModel)
    (rowStart: int) (rowEnd: int)
    (colStart: int) (colEnd: int)
    (offsetX: float) (offsetY: float)
    (emit: PaintItem -> unit) : unit =
    if rowStart <= rowEnd && colStart <= colEnd then
        let totalW = geometry.Col.Sum colStart colEnd
        let totalH = geometry.Row.Sum rowStart rowEnd
        // Excel's printed gridlines rasterize as near-black in the reference
        // PDF, unlike the lighter screen grid used in the interactive view.
        let color = "#000000"
        let w = 1.0
        let mutable x = 0.0
        emit (GridLine (offsetX + x, offsetY, offsetX + x, offsetY + totalH, color, w))
        for c in colStart .. colEnd do
            x <- x + geometry.Col.Size c
            emit (GridLine (offsetX + x, offsetY, offsetX + x, offsetY + totalH, color, w))
        let mutable y = 0.0
        emit (GridLine (offsetX, offsetY + y, offsetX + totalW, offsetY + y, color, w))
        for r in rowStart .. rowEnd do
            y <- y + geometry.Row.Size r
            emit (GridLine (offsetX, offsetY + y, offsetX + totalW, offsetY + y, color, w))

let private gridLineItems
    (geometry: PrintGeometry)
    (model: SpreadsheetModel)
    (rowStart: int) (rowEnd: int)
    (colStart: int) (colEnd: int)
    (offsetX: float) (offsetY: float) : PaintItem list =
    if rowStart > rowEnd || colStart > colEnd then [] else
    let rowCount = max 0 (rowEnd - rowStart + 1)
    let colCount = max 0 (colEnd - colStart + 1)
    let items = ResizeArray<PaintItem>(rowCount + colCount + 2)
    iterGridLineItems geometry model rowStart rowEnd colStart colEnd offsetX offsetY items.Add
    resizeArrayToList items

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
        Rotation = 0
        Clip = None }

/// Build the column-letter strip above the body+repeats and the row-number
/// strip to its left, plus a separator gridline below the column-letter
/// strip and to the right of the row-number strip. `xLeft` / `yTop` is the
/// page-local CSS-px top-left of the heading-shifted content origin
/// (typically `(printHeadingsColWidthPx, printHeadingsRowHeightPx)` when
/// the page itself starts at 0,0). The strips' inner-padded text avoids
/// touching the separator lines.
let private headingItems
    (geometry: PrintGeometry)
    (model: SpreadsheetModel)
    (slice: Pagination.PageSlice)
    (pagination: Pagination.Pagination)
    : PaintItem list =
    let headW = Pagination.printHeadingsColWidthPx
    let headH = Pagination.printHeadingsRowHeightPx
    // Column ranges to label: repeat cols (if any) followed by the slice's
    // body cols. Each emits its letter in the strip directly above the
    // matching column.
    let allCols = ResizeArray<int>()
    match pagination.RepeatCols with
    | Some (a, b) when b >= a ->
        for c in a .. b do allCols.Add c
    | _ -> ()
    for c in slice.ColStart .. slice.ColEnd do
        allCols.Add c
    let allRows = ResizeArray<int>()
    match pagination.RepeatRows with
    | Some (a, b) when b >= a ->
        for r in a .. b do allRows.Add r
    | _ -> ()
    for r in slice.RowStart .. slice.RowEnd do
        allRows.Add r
    let colWidth c = geometry.Col.Size c
    let rowHeight r = geometry.Row.Size r
    let mutable totalContentW = 0.0
    for c in allCols do
        totalContentW <- totalContentW + colWidth c
    let mutable totalContentH = 0.0
    for r in allRows do
        totalContentH <- totalContentH + rowHeight r
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
let private appendSectionOps
    (ops: ResizeArray<PdfDraw.DrawOp>)
    (options: RenderOptions)
    (faceByName: Map<string, PdfFont.TrueTypeFont>)
    (faceResource: Map<string, string>)
    (faceRemap: Map<string, Map<int, int>>)
    (leftPt: float) (rightPt: float)
    (section: HeaderFooter.StyledSection) (defaultFontSizePt: float) (textScale: float) (baselineYPt: float)
    (xAlign: TextAlign) =
    // Resolve face / size / color per H/F run. Excel lets section strings
    // switch style mid-run with codes like `&B`, `&"Arial,Italic"`, `&12`,
    // `&KFF0000`, `&U`, `&S`, `&X`, `&Y`; alignment still applies to the
    // total section width. Super/subscript shrinks the run's font size
    // to ~58% of the base; underline / strikethrough render as strokes.
    let runSizeAndBaseline (run: HeaderFooter.StyledRun) =
        let rawSizePt = run.FontSize |> Option.defaultValue defaultFontSizePt
        let baseSizePt = rawSizePt * textScale
        let scale =
            if run.Superscript || run.Subscript then 0.58 else 1.0
        let fontSizePt = baseSizePt * scale
        let baselineDeltaPt =
            if run.Superscript then baseSizePt * 0.33
            elif run.Subscript then -baseSizePt * 0.20
            else 0.0
        fontSizePt, baselineDeltaPt
    let appendTextRun faceName ttf text fontSizePt baselineDeltaPt color xPt =
        let remap =
            match faceRemap.TryFind faceName with
            | Some found -> found
            | None -> Map.empty
        let hex = toRemappedHex ttf remap text
        let fillColor =
            match color with
            | Some color -> parseHexColor color
            | None -> PdfColor.black
        ops.Add PdfDraw.BeginText
        ops.Add(PdfDraw.SetFillColor fillColor)
        ops.Add(PdfDraw.SetFont (faceName, fontSizePt))
        ops.Add(PdfDraw.TextMatrix (1.0, 0.0, 0.0, 1.0, xPt, baselineYPt + baselineDeltaPt))
        ops.Add(PdfDraw.ShowTextHex hex)
        ops.Add PdfDraw.EndText
    let startX totalWidthPt =
        match xAlign with
        | AlignLeft -> leftPt
        | AlignCenter -> leftPt + ((rightPt - leftPt) - totalWidthPt) / 2.0
        | AlignRight -> rightPt - totalWidthPt
    let tryAppendSingleRun () =
        match section.Runs with
        | [ run ] when
            not (String.IsNullOrEmpty run.Text)
            && not run.Underline
            && not run.DoubleUnderline
            && not run.Strikethrough ->
            match pickFace options run.FontName run.Bold run.Italic with
            | None -> false
            | Some primaryFace ->
                match Map.tryFind primaryFace.FaceName faceByName with
                | None -> false
                | Some primaryTtf ->
                    if List.isEmpty options.CjkFallback || fontCoversAllText primaryTtf run.Text then
                        let fontSizePt, baselineDeltaPt = runSizeAndBaseline run
                        let widthPt = float (PdfFont.getTextWidth primaryTtf run.Text) / float primaryTtf.UnitsPerEm * fontSizePt
                        appendTextRun primaryFace.FaceName primaryTtf run.Text fontSizePt baselineDeltaPt run.Color (startX widthPt)
                        true
                    else false
        | _ -> false
    if not (String.IsNullOrEmpty section.Text) && not (tryAppendSingleRun ()) then
        let segmentsWithWidth = ResizeArray<HeaderFooterSegment>()
        let mutable totalWidthPt = 0.0
        let addSegment face ttf text fontSizePt color underline doubleUnderline strikethrough baselineDeltaPt =
            let widthPt = float (PdfFont.getTextWidth ttf text) / float ttf.UnitsPerEm * fontSizePt
            segmentsWithWidth.Add {
                Face = face
                Ttf = ttf
                Text = text
                WidthPt = widthPt
                FontSizePt = fontSizePt
                Color = color
                Underline = underline
                DoubleUnderline = doubleUnderline
                Strikethrough = strikethrough
                BaselineDeltaPt = baselineDeltaPt
            }
            totalWidthPt <- totalWidthPt + widthPt
        for run in section.Runs do
            if not (String.IsNullOrEmpty run.Text) then
                // Excel scales H/F text by the page's print scale (the
                // `&NN` size code carries the *un*scaled value). Apply the
                // scale uniformly here so the &NN baseline still works.
                let fontSizePt, baselineDeltaPt = runSizeAndBaseline run
                match pickFace options run.FontName run.Bold run.Italic with
                | None -> ()
                | Some primaryFace ->
                    match Map.tryFind primaryFace.FaceName faceByName with
                    | None -> ()
                    | Some primaryTtf ->
                        if List.isEmpty options.CjkFallback || fontCoversAllText primaryTtf run.Text then
                            addSegment primaryFace primaryTtf run.Text fontSizePt run.Color run.Underline run.DoubleUnderline run.Strikethrough baselineDeltaPt
                        else
                            for (face, text) in segmentMixedByFont (primaryFace, primaryTtf) options.CjkFallback (tryGetMappedFace faceByName) run.Bold run.Text do
                                addSegment (fst face) (snd face) text fontSizePt run.Color run.Underline run.DoubleUnderline run.Strikethrough baselineDeltaPt
        if segmentsWithWidth.Count > 0 then
            let startXPt = startX totalWidthPt
            ops.Add PdfDraw.BeginText
            let mutable cursorXPt = startXPt
            for segment in segmentsWithWidth do
                let faceName = segment.Face.FaceName
                let remap =
                    match faceRemap.TryFind faceName with
                    | Some found -> found
                    | None -> Map.empty
                let hex = toRemappedHex segment.Ttf remap segment.Text
                let fillColor =
                    match segment.Color with
                    | Some color -> parseHexColor color
                    | None -> PdfColor.black
                ops.Add(PdfDraw.SetFillColor fillColor)
                ops.Add(PdfDraw.SetFont (faceName, segment.FontSizePt))
                ops.Add(PdfDraw.TextMatrix (1.0, 0.0, 0.0, 1.0, cursorXPt, baselineYPt + segment.BaselineDeltaPt))
                ops.Add(PdfDraw.ShowTextHex hex)
                cursorXPt <- cursorXPt + segment.WidthPt
            ops.Add PdfDraw.EndText

            // Decorations must be emitted after EndText; PDF path ops cannot appear
            // inside a text object.
            let mutable decorationXPt = startXPt
            for segment in segmentsWithWidth do
                if segment.Underline || segment.DoubleUnderline || segment.Strikethrough then
                    let strokeColor =
                        match segment.Color with
                        | Some color -> parseHexColor color
                        | None -> PdfColor.black
                    let lineWidth = max 0.5 (segment.FontSizePt * 0.05)
                    if segment.Underline then
                        let yUL = baselineYPt - segment.FontSizePt * 0.12
                        ops.Add(PdfDraw.SetStrokeColor strokeColor)
                        ops.Add(PdfDraw.SetLineWidth lineWidth)
                        ops.Add(PdfDraw.SetDash ([], 0.0))
                        ops.Add(PdfDraw.MoveTo (decorationXPt, yUL))
                        ops.Add(PdfDraw.LineTo (decorationXPt + segment.WidthPt, yUL))
                        ops.Add PdfDraw.Stroke
                    if segment.DoubleUnderline then
                        let yU1 = baselineYPt - segment.FontSizePt * 0.08
                        let yU2 = baselineYPt - segment.FontSizePt * 0.18
                        ops.Add(PdfDraw.SetStrokeColor strokeColor)
                        ops.Add(PdfDraw.SetLineWidth lineWidth)
                        ops.Add(PdfDraw.SetDash ([], 0.0))
                        ops.Add(PdfDraw.MoveTo (decorationXPt, yU1))
                        ops.Add(PdfDraw.LineTo (decorationXPt + segment.WidthPt, yU1))
                        ops.Add PdfDraw.Stroke
                        ops.Add(PdfDraw.MoveTo (decorationXPt, yU2))
                        ops.Add(PdfDraw.LineTo (decorationXPt + segment.WidthPt, yU2))
                        ops.Add PdfDraw.Stroke
                    if segment.Strikethrough then
                        let yST = baselineYPt + segment.FontSizePt * 0.28
                        ops.Add(PdfDraw.SetStrokeColor strokeColor)
                        ops.Add(PdfDraw.SetLineWidth lineWidth)
                        ops.Add(PdfDraw.SetDash ([], 0.0))
                        ops.Add(PdfDraw.MoveTo (decorationXPt, yST))
                        ops.Add(PdfDraw.LineTo (decorationXPt + segment.WidthPt, yST))
                        ops.Add PdfDraw.Stroke
                decorationXPt <- decorationXPt + segment.WidthPt

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
/// Render to PDF with anchored vector overlays (charts/sparklines). Each
/// overlay is `(fromRow, fromCol, fromColOffPx, fromRowOffPx, items)`: a cell
/// anchor + pixel offset within that cell, and the overlay's paint items in
/// local (0,0)-origin space. The overlay is placed on the page whose slice
/// contains the anchor, offset into that page's content band — so it paginates
/// with the grid. Positions use the print model's own widths, keeping them
/// consistent with the cell content around them.
/// One sheet of a PDF: the sheet (its page setup is what it prints), its name (what `&A` prints) and the
/// charts and shapes drawn over its cells (`Paint.drawingOverlays`).
type PdfSheet = {
    Model: SpreadsheetModel
    SheetName: string
    Overlays: (int * int * float * float * float * float * PaintItem list) list
}

/// The PDF of sheets, one after the other in one document (a workbook printed whole): each sheet
/// laid out on its pages by its own page setup, the fonts embedded once for all of them.
let toBytesOfSheets (options: RenderOptions) (sheets: PdfSheet list) : byte[] =
    let builder = PdfDoc.create ()
    let imageNames = EmbeddedImages builder

    let parsed = System.Collections.Generic.Dictionary<string, PdfFont.TrueTypeFont>()
    let failedFaces = System.Collections.Generic.HashSet<string>()
    let tryGetParsedFace (face: FontFace) =
        let mutable cached = Unchecked.defaultof<PdfFont.TrueTypeFont>
        if parsed.TryGetValue(face.FaceName, &cached) then
            Some cached
        elif failedFaces.Contains face.FaceName then
            None
        else
            try
                let ttf = PdfFont.parseTrueType face.Data
                parsed.[face.FaceName] <- ttf
                Some ttf
            with _ ->
                failedFaces.Add face.FaceName |> ignore
                None
    for face in options.Fonts do
        tryGetParsedFace face |> ignore
    let canDrawText =
        options.Fonts
        |> List.exists (fun face -> parsed.ContainsKey face.FaceName)

    let measurePdf (text: string) (fontPt: float) (fontName: string option) (bold: bool) (italic: bool) : float =
        let approx () = float text.Length * 7.0
        match pickFace options fontName bold italic with
        | None -> approx ()
        | Some primaryFace ->
            match Portable.Dict.tryFind primaryFace.FaceName parsed with
            | ValueSome primaryTtf ->
                let scaleWidth (ttf: PdfFont.TrueTypeFont) s =
                    let widthUnits = PdfFont.getTextWidth ttf s
                    float widthUnits / float ttf.UnitsPerEm * fontPt * 96.0 / 72.0
                if text.Length = 0 then 0.0
                elif List.isEmpty options.CjkFallback || fontCoversAllText primaryTtf text then
                    // Excel prints at 600 dpi: the font as high as a whole number of dots (11 pt is 92, 11.04 pt in its
                    // PDF), each character a whole number of dots wide (make-print-text-fixture.py: Calibri 11's 10
                    // zeros 56.36 pt, 470 dots). Surrogate pairs and the like as a whole.
                    let dots = System.Math.Round (fontPt * 600.0 / 72.0)
                    if text |> Seq.exists System.Char.IsSurrogate then scaleWidth primaryTtf text
                    else
                        let mutable total = 0.0
                        for c in text do
                            total <- total + System.Math.Round (float (PdfFont.getTextWidth primaryTtf (string c)) / float primaryTtf.UnitsPerEm * dots)
                        total * 96.0 / 600.0
                else
                    let mutable total = 0.0
                    for (face, s) in segmentMixedByFont (primaryFace, primaryTtf) options.CjkFallback tryGetParsedFace bold text do
                        total <- total + scaleWidth (snd face) s
                    total
            | _ -> approx ()

    // Collect used GIDs across every page's items, plus header / footer
    // sections (so their glyphs land in the subset too).
    let usedGids = System.Collections.Generic.Dictionary<string, System.Collections.Generic.HashSet<int>>()
    let recordGids (faceName: string) (ttf: PdfFont.TrueTypeFont) (text: string) =
        let set =
            match Portable.Dict.tryFind faceName usedGids with
            | ValueSome s -> s
            | _ ->
                let s = System.Collections.Generic.HashSet<int>()
                usedGids.[faceName] <- s
                s
        for ch in text do
            let gid =
                match ttf.Cmap.TryFind (int ch) with
                | Some found -> found
                | None -> 0
            set.Add gid |> ignore
    let recordTextGids bold italic fontName text =
        if not (String.IsNullOrEmpty text) then
            match pickFace options fontName bold italic with
            | None -> ()
            | Some primaryFace ->
                match Portable.Dict.tryFind primaryFace.FaceName parsed with
                | ValueSome primaryTtf ->
                    if List.isEmpty options.CjkFallback || fontCoversAllText primaryTtf text then
                        recordGids primaryFace.FaceName primaryTtf text
                    else
                        for (face, s) in segmentMixedByFont (primaryFace, primaryTtf) options.CjkFallback tryGetParsedFace bold text do
                            recordGids (fst face).FaceName (snd face) s
                | _ -> ()
    let recordStyled (section: HeaderFooter.StyledSection) =
        for run in section.Runs do
            if not (String.IsNullOrEmpty run.Text) then
                match pickFace options run.FontName run.Bold run.Italic with
                | None -> ()
                | Some primaryFace ->
                    match Portable.Dict.tryFind primaryFace.FaceName parsed with
                    | ValueSome primaryTtf ->
                        if List.isEmpty options.CjkFallback || fontCoversAllText primaryTtf run.Text then
                            recordGids primaryFace.FaceName primaryTtf run.Text
                        else
                            for (face, s) in segmentMixedByFont (primaryFace, primaryTtf) options.CjkFallback tryGetParsedFace run.Bold run.Text do
                                recordGids (fst face).FaceName (snd face) s
                    | _ -> ()
    let recordPaintTextGids (t: PaintText) =
        recordTextGids t.Bold t.Italic t.FontName t.Text
    let rec recordPaintItemGids item =
        match item with
        | DrawText t -> recordPaintTextGids t
        | VecGroup (_, children) | Clipped (_, children) ->
            for child in children do
                recordPaintItemGids child
        | _ -> ()
    let planSheet (sheet: PdfSheet) =
        let options = { options with SheetName = sheet.SheetName }
        let model = sheet.Model
        let overlays = sheet.Overlays
        let maxPrintCol, maxPrintRow = printDimensionBounds model overlays
        let model =
            model
            |> withCommentMarkersForPrint
            |> withPrintColumnWidths maxPrintCol
            |> withPrintRowHeights maxPrintRow
        // Compute pagination from PageSetup. Empty workbook → one blank page.
        let pagination = Pagination.paginate model
        let ps = model.PageSetup
        // (A landscape page is the wider way round, however the setup holds the paper's size.)
        let pageWidthPt, pageHeightPt = Pagination.orientedPageSize ps
        let leftMarginPt = ps.Margins.Left * 72.0
        let rightMarginPt = ps.Margins.Right * 72.0
        let topMarginPt = ps.Margins.Top * 72.0
        let bottomMarginPt = ps.Margins.Bottom * 72.0
        let headerMarginPt = ps.Margins.Header * 72.0
        let footerMarginPt = ps.Margins.Footer * 72.0

        let pages = pagination.Pages |> List.toArray
        let totalPages = max 1 pages.Length
        let maxColIndex =
            let mutable maxCol = 0
            for page in pages do
                maxCol <- max maxCol page.ColEnd
            match pagination.RepeatCols with
            | Some (_, b) -> maxCol <- max maxCol b
            | None -> ()
            for _, fc, _, _, _, _, _ in overlays do
                maxCol <- max maxCol fc
            maxCol
        let maxRowIndex =
            let mutable maxRow = 0
            for page in pages do
                maxRow <- max maxRow page.RowEnd
            match pagination.RepeatRows with
            | Some (_, b) -> maxRow <- max maxRow b
            | None -> ()
            for fr, _, _, _, _, _, _ in overlays do
                maxRow <- max maxRow fr
            maxRow
        let geometry = buildPrintGeometry model maxRowIndex maxColIndex
        let repeatColsWidthPx =
            pagination.RepeatCols
            |> Option.map (fun (a, b) -> geometry.Col.Sum a b)
            |> Option.defaultValue 0.0
        let repeatRowsHeightPx =
            pagination.RepeatRows
            |> Option.map (fun (a, b) -> geometry.Row.Sum a b)
            |> Option.defaultValue 0.0

        // Page items are streamed. When fonts are available, glyph collection
        // makes a first pass over the pages so font subsetting can run before
        // PDF pages are emitted, but we avoid retaining every page's PaintItem
        // list at once.
        let printGridLines = ps.Options.GridLines
        let printHeadings = ps.Options.Headings
        // When PrintOptions.Headings is on, every content band shifts right /
        // down by these to make room for the column-letter strip on top and
        // the row-number strip on the left. Pagination has already shrunk the
        // per-page area by the same amounts, so manual breaks land in the
        // right place.
        let headingsColPx, headingsRowPx = Pagination.headingsOffsets ps
        let overlaysByPage =
            if List.isEmpty overlays || pages.Length = 0 then
                [||]
            else
                // A drawing belongs to every page it reaches on (cut off at the page's edge); one beside a
                // print area to none; without any page that reaches it, to the first (a model without its drawings).
                let spanOf (sum: int -> int -> float) (first: int) (last: int) = sum 0 (first - 1), sum 0 last
                let byPage = Array.init pages.Length (fun _ -> ResizeArray<_>())
                for overlay in overlays do
                    let fr, fc, fox, foy, width, height, _ = overlay
                    let left = geometry.Col.Sum 0 (fc - 1) + fox
                    let top = geometry.Row.Sum 0 (fr - 1) + foy
                    let mutable placed = false
                    pages |> Array.iteri (fun index slice ->
                        let colStart, colEnd = spanOf geometry.Col.Sum slice.ColStart slice.ColEnd
                        let rowStart, rowEnd = spanOf geometry.Row.Sum slice.RowStart slice.RowEnd
                        if left < colEnd && left + max width 1.0 > colStart && top < rowEnd && top + max height 1.0 > rowStart then
                            byPage.[index].Add overlay
                            placed <- true)
                    if not placed && List.isEmpty (printAreasOf model.PageSetup) then byPage.[0].Add overlay
                byPage

        // (What the overlays draw is not painted again as the frames of the cell slices.)
        let paintedModel =
            if List.isEmpty overlays then model
            else { model with Drawings = model.Drawings |> List.filter (fun drawing -> match drawing.Content with KeptObject _ -> false | _ -> true) }

        let appendBand (emit: PaintItem -> unit)
                       (rowStart: int) (rowEnd: int)
                       (colStart: int) (colEnd: int)
                       (offsetX: float) (offsetY: float) =
            if printGridLines then
                let mutable sawContent = false
                let mutable addedGrid = false
                let addGrid () =
                    if not addedGrid then
                        iterGridLineItems geometry model rowStart rowEnd colStart colEnd offsetX offsetY emit
                        addedGrid <- true
                iterSliceItems geometry paintedModel measurePdf canDrawText rowStart rowEnd colStart colEnd offsetX offsetY (fun item ->
                    if not sawContent then
                        sawContent <- true
                        match item with
                        | FillRect _ ->
                            emit item
                            addGrid ()
                        | _ ->
                            addGrid ()
                            emit item
                    else
                        emit item)
                if not sawContent then
                    addGrid ()
            else
                iterSliceItems geometry paintedModel measurePdf canDrawText rowStart rowEnd colStart colEnd offsetX offsetY emit

        let emitPageItems (i: int) (emit: PaintItem -> unit) =
            if i < pages.Length then
                let slice = pages.[i]
                // Body: rows × cols of the slice, offset by repeat band sizes.
                appendBand
                    emit
                    slice.RowStart slice.RowEnd
                    slice.ColStart slice.ColEnd
                    (headingsColPx + repeatColsWidthPx) (headingsRowPx + repeatRowsHeightPx)
                // Repeat-rows band (above body, full width): repeat rows × slice cols.
                match pagination.RepeatRows with
                | Some (rs, re) when re >= rs ->
                    appendBand
                        emit
                        rs re
                        slice.ColStart slice.ColEnd
                        (headingsColPx + repeatColsWidthPx) headingsRowPx
                | _ -> ()
                // Repeat-cols band (left of body, full height): slice rows × repeat cols.
                match pagination.RepeatCols with
                | Some (cs, ce) when ce >= cs ->
                    appendBand
                        emit
                        slice.RowStart slice.RowEnd
                        cs ce
                        headingsColPx (headingsRowPx + repeatRowsHeightPx)
                | _ -> ()
                // Repeat corner (top-left): repeat rows × repeat cols.
                match pagination.RepeatRows, pagination.RepeatCols with
                | Some (rs, re), Some (cs, ce) when re >= rs && ce >= cs ->
                    appendBand emit rs re cs ce headingsColPx headingsRowPx
                | _ -> ()
                if printHeadings then
                    for item in headingItems geometry model slice pagination do
                        emit item
                // Inject anchored overlays (charts/sparklines) onto the page whose slice
                // contains each anchor cell, offset into that page's content band. Done
                // before GID collection so overlay text fonts get subset too.
                if i < overlaysByPage.Length && overlaysByPage.[i].Count > 0 then
                    let baseX = headingsColPx + repeatColsWidthPx
                    let baseY = headingsRowPx + repeatRowsHeightPx
                    let colStartPx = geometry.Col.Sum 0 (slice.ColStart - 1)
                    let rowStartPx = geometry.Row.Sum 0 (slice.RowStart - 1)
                    let clip =
                        { X = baseX; Y = baseY
                          Width = geometry.Col.Sum slice.ColStart slice.ColEnd
                          Height = geometry.Row.Sum slice.RowStart slice.RowEnd }
                    for fr, fc, fox, foy, _, _, locItems in overlaysByPage.[i] do
                        let gx = geometry.Col.Sum 0 (fc - 1) + fox
                        let gy = geometry.Row.Sum 0 (fr - 1) + foy
                        let tx = baseX + (gx - colStartPx)
                        let ty = baseY + (gy - rowStartPx)
                        emit (Clipped (clip, locItems |> List.map (shiftItem tx ty)))

        if canDrawText then
            for pageIdx in 0 .. totalPages - 1 do
                emitPageItems pageIdx recordPaintItemGids

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

        let headerFooterTemplateCache = System.Collections.Generic.Dictionary<string, HeaderFooter.Template>()
        let parsedHeaderFooter raw =
            match Portable.Dict.tryFind raw headerFooterTemplateCache with
            | ValueSome template -> template
            | ValueNone ->
                let template = HeaderFooter.parse raw
                headerFooterTemplateCache.[raw] <- template
                template
        let resolveHeaderFooter ctx raw =
            HeaderFooter.resolveTemplate ctx (parsedHeaderFooter raw)

        // Build header/footer page contexts and pre-resolve all text so we
        // can register its glyphs for subsetting before embedding fonts.
        // `&P` honors `firstPageNumber` when `useFirstPageNumber` is set;
        // `&N` keeps reporting the physical page count (Excel doesn't add
        // the first-page offset to N).
        let firstPageOffset =
            if ps.UseFirstPageNumber then ps.FirstPageNumber - 1 else 0
        let resolvedHFs =
            Array.init totalPages (fun pageIdx ->
                let ctx : HeaderFooter.PageContext =
                    { PageNumber = pageIdx + 1 + firstPageOffset
                      TotalPages = totalPages
                      SheetName = options.SheetName
                      FileName = options.FileName
                      FilePath = ""
                      Now = DateTime.Now }
                let headerRaw, footerRaw = pickHeaderFooter (pageIdx + 1)
                let header = headerRaw |> Option.map (resolveHeaderFooter ctx)
                let footer = footerRaw |> Option.map (resolveHeaderFooter ctx)
                header, footer)
        for (header, footer) in resolvedHFs do
            match header with
            | Some section ->
                recordStyled section.Left
                recordStyled section.Center
                recordStyled section.Right
            | None -> ()
            match footer with
            | Some section ->
                recordStyled section.Left
                recordStyled section.Center
                recordStyled section.Right
            | None -> ()

        // (The pages are added once the fonts of every sheet are embedded.)
        fun (faceTtfMap: Map<string, PdfFont.TrueTypeFont>) (faceResourceMap: Map<string, string>) (faceRemapMap: Map<string, Map<int, int>>) (loadedFaceNames: string list) ->
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
                let mutable maxFontSize = hfScaledFontSizePt
                for run in section.Runs do
                    if not (String.IsNullOrEmpty run.Text) then
                        let fontSize = (run.FontSize |> Option.defaultValue hfFontSizePt) * hfTextScale
                        if fontSize > maxFontSize then
                            maxFontSize <- fontSize
                maxFontSize
            let maxHeaderFooterFontSize (section: HeaderFooter.Section) =
                max (max (maxSectionFontSize section.Left) (maxSectionFontSize section.Center)) (maxSectionFontSize section.Right)
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
            let appendStyledSectionOps (ops: ResizeArray<PdfDraw.DrawOp>) styled baselineYPt xAlign =
                appendSectionOps ops options faceTtfMap faceResourceMap faceRemapMap
                    printableLeftPt printableRightPt
                    styled hfFontSizePt hfTextScale baselineYPt xAlign
            let appendHeaderFooterOps ops (section: HeaderFooter.Section) baselineYPt =
                appendStyledSectionOps ops section.Left baselineYPt AlignLeft
                appendStyledSectionOps ops section.Center baselineYPt AlignCenter
                appendStyledSectionOps ops section.Right baselineYPt AlignRight

            // Emit one PDF page per page slice (or one blank page when the
            // pagination is empty).
            let printableWidthPt = pageWidthPt - leftMarginPt - rightMarginPt
            let printableHeightPt = pageHeightPt - topMarginPt - bottomMarginPt
            for pageIdx in 0 .. totalPages - 1 do
                // Per-page content extent in points, accounting for heading strips
                // and repeat bands. Excel centers each page individually against its
                // own filled area, so a partly-filled last page still centers on
                // what's actually drawn rather than the full printable rectangle.
                let leftPtOffset, topPtOffset =
                    if pages.Length = 0 then leftMarginPt, topMarginPt
                    else
                        let slice = pages.[pageIdx]
                        let bodyWidthPx =
                            headingsColPx + repeatColsWidthPx
                            + geometry.Col.Sum slice.ColStart slice.ColEnd
                        let bodyHeightPx =
                            headingsRowPx + repeatRowsHeightPx
                            + geometry.Row.Sum slice.RowStart slice.RowEnd
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
                      BlackAndWhite = ps.BlackAndWhite
                      ImageName = imageNames.NameOf }
                let ops = System.Collections.Generic.List<PdfDraw.DrawOp>()
                let sz v = v * pxToPt * placement.Scale
                let xPt x = placement.LeftPtOffset + sz x
                let yPdf y = placement.PageHeightPt - placement.TopPtOffset - sz y
                let mutable gridBatchActive = false
                let mutable gridBatchColor = ""
                let mutable gridBatchWidth = 0.0
                let flushGridBatch () =
                    if gridBatchActive then
                        ops.Add PdfDraw.Stroke
                        gridBatchActive <- false
                let appendGridLine x1 y1 x2 y2 color width =
                    let color = bwStroke placement.BlackAndWhite color
                    if not gridBatchActive || gridBatchColor <> color || gridBatchWidth <> width then
                        flushGridBatch ()
                        ops.Add(PdfDraw.SetStrokeColor (parseHexColor color))
                        ops.Add(PdfDraw.SetLineWidth (sz width))
                        ops.Add(PdfDraw.SetDash ([], 0.0))
                        gridBatchActive <- true
                        gridBatchColor <- color
                        gridBatchWidth <- width
                    ops.Add(PdfDraw.MoveTo (xPt x1, yPdf y1))
                    ops.Add(PdfDraw.LineTo (xPt x2, yPdf y2))
                let appendItem item =
                    match item with
                    | GridLine (x1, y1, x2, y2, color, width) ->
                        appendGridLine x1 y1 x2 y2 color width
                    | _ ->
                        flushGridBatch ()
                        appendTranslatedItem options faceTtfMap faceResourceMap faceRemapMap placement ops item
                emitPageItems pageIdx appendItem
                flushGridBatch ()
                let header, footer = resolvedHFs.[pageIdx]
                match header with
                | Some section ->
                    appendHeaderFooterOps ops section (headerBaselineYPt section)
                | None -> ()
                match footer with
                | Some section ->
                    appendHeaderFooterOps ops section footerBaselineYPt
                | None -> ()
                PdfDoc.addPage builder (pageWidthPt, pageHeightPt) loadedFaceNames (imageNames.TakePageNames ()) (resizeArrayToList ops) |> ignore

    let addSheets = sheets |> List.map planSheet
    // An empty list of sheets is a document of one blank page.
    let faceRemap = System.Collections.Generic.Dictionary<string, Map<int, int>>()
    let embedFace (face: FontFace) =
        match Portable.Dict.tryFind face.FaceName parsed with
        | ValueNone -> ()
        | ValueSome _ ->
            let used =
                match Portable.Dict.tryFind face.FaceName usedGids with
                | ValueSome s ->
                    let mutable gids = Set.empty
                    for gid in s do
                        gids <- Set.add gid gids
                    gids
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

    let mutable faceTtfMap = Map.empty
    for KeyValue(name, ttf) in parsed do
        faceTtfMap <- Map.add name ttf faceTtfMap
    let mutable faceResourceMap = Map.empty
    for face in options.Fonts do
        match PdfDoc.getEmbeddedFont builder face.FaceName with
        | Some ef -> faceResourceMap <- Map.add face.FaceName ef.Name faceResourceMap
        | None -> ()
    for face in options.CjkFallback do
        match PdfDoc.getEmbeddedFont builder face.FaceName with
        | Some ef -> faceResourceMap <- Map.add face.FaceName ef.Name faceResourceMap
        | None -> ()
    let mutable faceRemapMap = Map.empty
    for KeyValue(name, remap) in faceRemap do
        faceRemapMap <- Map.add name remap faceRemapMap

    let loadedFaceNames =
        let result = ResizeArray<string>()
        for face in options.Fonts do
            if faceResourceMap.ContainsKey face.FaceName then
                result.Add face.FaceName
        for face in options.CjkFallback do
            if faceResourceMap.ContainsKey face.FaceName then
                result.Add face.FaceName
        resizeArrayToList result

    if List.isEmpty addSheets then
        PdfDoc.addPage builder (612.0, 792.0) [] [] [] |> ignore
    for addPages in addSheets do
        addPages faceTtfMap faceResourceMap faceRemapMap loadedFaceNames
    PdfDoc.toBytes builder

/// Render one sheet to PDF with its overlays (charts and shapes drawn over its cells).
let toBytesWith (options: RenderOptions) (model: SpreadsheetModel)
                (overlays: (int * int * float * float * float * float * PaintItem list) list) : byte[] =
    toBytesOfSheets options [ { Model = model; SheetName = options.SheetName; Overlays = overlays } ]

/// Render to PDF with no overlays (the common case).
let toBytes (options: RenderOptions) (model: SpreadsheetModel) : byte[] =
    toBytesWith options model []

/// Render a standalone vector PaintItem scene (no model/text) to a one-page
/// PDF — proves the VecPath/VecGroup primitives through the PDF path.
let sceneToBytes (widthPx: float) (heightPx: float) (items: PaintItem list) : byte[] =
    let builder = PdfDoc.create ()
    let imageNames = EmbeddedImages builder
    let pageWidthPt = widthPx * pxToPt
    let pageHeightPt = heightPx * pxToPt
    let placement =
        { LeftPtOffset = 0.0; TopPtOffset = 0.0; PageHeightPt = pageHeightPt
          Scale = 1.0; BlackAndWhite = false; ImageName = imageNames.NameOf }
    let opts = { DefaultFontFamily = "Calibri"; Fonts = []; CjkFallback = []; SheetName = ""; FileName = "" }
    let ops = System.Collections.Generic.List<PdfDraw.DrawOp>()
    for item in items do
        for op in translateItem opts Map.empty Map.empty Map.empty placement item do
            ops.Add op
    PdfDoc.addPage builder (pageWidthPt, pageHeightPt) [] (imageNames.TakePageNames ()) (resizeArrayToList ops) |> ignore
    PdfDoc.toBytes builder

// ---- pictures for a writer that cannot decode PNG ----------------------------------------------------
// Rbt.Pdf embeds a JPEG as it is and decodes a PNG to embed its pixels - on .NET; its browser build
// cannot (no inflate of PNG data there). The browser component hands the PNG pictures of the sheets it
// exports to the browser, which draws them as JPEG (`pngPictures`, `withPictures`).

let private isPngData (data: byte[]) =
    data.Length > 8 && data.[0] = 0x89uy && data.[1] = 0x50uy && data.[2] = 0x4Euy && data.[3] = 0x47uy

let rec private itemPictures (add: byte[] -> unit) (item: PaintItem) =
    match item with
    | DrawImage (_, image) when isPngData image.Data -> add image.Data
    | VecGroup (_, children) | Clipped (_, children) -> children |> List.iter (itemPictures add)
    | _ -> ()

/// The PNG pictures of the sheets (their pictures and those in their overlays), each once.
let pngPictures (sheets: PdfSheet list) : byte[] array =
    let found = ResizeArray<byte[]>()
    let add (data: byte[]) = if not (found.Exists (fun known -> obj.ReferenceEquals(known, data))) then found.Add data
    for sheet in sheets do
        for drawing in sheet.Model.Drawings do
            match drawing.Content with
            | Picture (_, _, data) when isPngData data -> add data
            | _ -> ()
        for _, _, _, _, _, _, items in sheet.Overlays do
            items |> List.iter (itemPictures add)
    found.ToArray()

/// The sheets with each picture whose data `replace` gives other data for (by reference: the arrays of
/// `pngPictures`) drawn from that instead.
let withPictures (replace: byte[] -> (string * byte[]) option) (sheets: PdfSheet list) : PdfSheet list =
    let rec item (paint: PaintItem) =
        match paint with
        | DrawImage (rect, image) ->
            match replace image.Data with
            | Some (contentType, data) -> DrawImage (rect, { image with ContentType = contentType; Data = data })
            | None -> paint
        | VecGroup (transform, children) -> VecGroup (transform, children |> List.map item)
        | Clipped (rect, children) -> Clipped (rect, children |> List.map item)
        | _ -> paint
    sheets |> List.map (fun sheet ->
        { sheet with
            Model =
                { sheet.Model with
                    Drawings =
                        sheet.Model.Drawings |> List.map (fun drawing ->
                            match drawing.Content with
                            | Picture (_, extension, data) ->
                                match replace data with
                                | Some (contentType, other) -> { drawing with Content = Picture (contentType, extension, other) }
                                | None -> drawing
                            | _ -> drawing) }
            Overlays = sheet.Overlays |> List.map (fun (r, c, x, y, w, h, items) -> r, c, x, y, w, h, items |> List.map item) })

/// A face from what a host hands (the record of Word.Pdf's `FontFace` has the same fields, in the same
/// order: a host makes its faces once for both).
let fontFace (faceName: string) (family: string) (bold: bool) (italic: bool) (data: byte[]) : FontFace =
    { FaceName = faceName; Family = family; Bold = bold; Italic = italic; Data = data }

/// For the browser component: the PDF of `sheets` (each the sheet, its name, its overlays) with `fonts`.
let exportSheets (defaultFamily: string) (fileName: string) (fonts: FontFace array) (sheets: PdfSheet array) : byte[] =
    let options = { DefaultFontFamily = defaultFamily; Fonts = List.ofArray fonts; CjkFallback = []; SheetName = ""; FileName = fileName }
    toBytesOfSheets options (List.ofArray sheets)

/// The faces (family, bold, italic) the text of `sheets` is set in: the cells' formats and their rich-text
/// runs (cells with a value), the text of their charts and shapes, and the default family's regular and bold (headers, footers
/// and headings). Families are as the workbook names them (`defaultFamily` for a cell that names none).
let requiredFaces (defaultFamily: string) (sheets: PdfSheet array) : (string * bool * bool) array =
    let faces = System.Collections.Generic.HashSet<string * bool * bool>()
    let add (family: string option) bold italic =
        faces.Add((family |> Option.filter (String.IsNullOrWhiteSpace >> not) |> Option.defaultValue defaultFamily), bold, italic) |> ignore
    add None false false
    add None true false
    let rec addItem item =
        match item with
        | DrawText t -> add t.FontName t.Bold t.Italic
        | VecGroup (_, children) | Clipped (_, children) -> children |> List.iter addItem
        | _ -> ()
    for sheet in sheets do
        for KeyValue(address, cell) in sheet.Model.Cells do
            // A cell with a format but nothing in it shows no text.
            if unwrapFormulaValue cell.Value <> CellValue.Empty then
                let format = cell.Format
                add format.FontName format.Bold format.Italic
                match Map.tryFind address sheet.Model.RichRuns with
                | Some runs ->
                    for _, run in runs do
                        add (run.FontName |> Option.orElse format.FontName) (defaultArg run.Bold format.Bold) (defaultArg run.Italic format.Italic)
                | None -> ()
        for _, _, _, _, _, _, items in sheet.Overlays do
            items |> List.iter addItem
    Seq.toArray faces

#if FABLE_COMPILER
// The package's own faces (fonts/ beside this bundle), for text the host has supplied no face for.
[<Fable.Core.Emit("new URL('../fonts/', import.meta.url).href")>]
let private defaultFontsUrl () : string = Fable.Core.Util.jsNative

[<Fable.Core.Emit("$0.then($1)")>]
let private thenMap (_: Fable.Core.JS.Promise<'T>) (_: 'T -> 'U) : Fable.Core.JS.Promise<'U> = Fable.Core.Util.jsNative

/// As `exportSheets`, with the package's own face (DefaultFonts) for each face the sheets' text needs that
/// `fonts` has not got: a host that supplies no fonts gets a PDF with all its text.
let exportSheetsWithDefaultFonts (defaultFamily: string) (fileName: string) (fonts: FontFace array) (sheets: PdfSheet array) : Fable.Core.JS.Promise<byte[]> =
    let has = DefaultFonts.hasFace (fonts |> Array.map (fun face -> face.Family, face.Bold, face.Italic))
    let planned = DefaultFonts.plan has defaultFamily (requiredFaces defaultFamily sheets)
    thenMap (DefaultFonts.fetchFaces (defaultFontsUrl ()) planned) (fun fetched ->
        let added = fetched |> Array.map (fun (face, data) -> fontFace face.FaceName face.Family face.Bold face.Italic data)
        exportSheets defaultFamily fileName (Array.append fonts added) sheets)
#endif

/// A sheet of `exportSheets`.
let pdfSheet (model: SpreadsheetModel) (sheetName: string) (overlays: (int * int * float * float * float * float * PaintItem list) list) : PdfSheet =
    { Model = model; SheetName = sheetName; Overlays = overlays }
