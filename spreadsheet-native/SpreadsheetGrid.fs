module Composit.Ui.SpreadsheetGrid

open Renderer
open Composit.Ui.Theme
open Composit.Ui.Widget

open RedBat.Spreadsheet.Model
open RedBat.Spreadsheet.Layout
open RedBat.Spreadsheet.Paint

let inline private px (v: float) : float<px> = LanguagePrimitives.FloatWithMeasure v

let private defaultFont () = Composit.Ui.Layout.font()

/// Cache of font variants keyed by (familyKey, bold, italic). familyKey is the
/// normalized lower-case font name we use to pick a TTF set ("calibri",
/// "arial", "cambria", "times"); unknown families fall back to "calibri".
let private fontCache = System.Collections.Generic.Dictionary<string * bool * bool, Renderer.Text.FontData option>()

let private fontDirs () =
    let cwd = System.IO.Directory.GetCurrentDirectory()
    [
        System.Environment.GetEnvironmentVariable("RB_WORD_FONT_DIR")
        System.IO.Path.GetFullPath(System.IO.Path.Combine(cwd, "artifacts/fonts/windows"))
        System.IO.Path.GetFullPath(System.IO.Path.Combine(cwd, "../../artifacts/fonts/windows"))
        // Sibling office repo's font cache, when running from office-spreadsheet/.
        System.IO.Path.GetFullPath(System.IO.Path.Combine(cwd, "../office/artifacts/fonts/windows"))
    ]
    |> List.choose (fun dir ->
        if System.String.IsNullOrWhiteSpace dir then None
        elif System.IO.Directory.Exists dir then Some dir
        else None)

/// Map an XLSX font `name val="..."` to a normalized family key. Anything we
/// don't have a TTF for falls back to "calibri" — the renderer keeps working
/// even when we encounter Asian/script fonts we can't substitute properly.
let private familyKey (name: string option) =
    match name with
    | None -> "calibri"
    | Some n ->
        let lower = n.Trim().ToLowerInvariant()
        if lower.StartsWith "arial" then "arial"
        elif lower.StartsWith "calibri" then "calibri"
        elif lower.StartsWith "cambria" then "cambria"
        elif lower.StartsWith "times" then "times"
        elif lower.StartsWith "verdana" then "verdana"
        elif lower.StartsWith "garamond" then "garamond"
        elif lower.StartsWith "segoe" then "segoeui"
        // Localized "Body Text Font" names from non-English Excel installs.
        // OOXML stores the localized display name verbatim; Windows resolves
        // it through the OS locale. For our purposes Arial is the closest
        // Excel substitute on a system that doesn't have the localized
        // font (verified against pandas-testdateoverflow's Polish
        // "Czcionka tekstu podstawowego" — column widths and row heights
        // match Excel's reported values with Arial metrics, not Calibri's).
        elif lower = "czcionka tekstu podstawowego" then "arial"     // Polish
        elif lower = "schrift des textkörpers" then "arial"          // German
        elif lower = "police du corps" then "arial"                  // French
        elif lower = "carattere del corpo" then "arial"              // Italian
        elif lower = "fuente del cuerpo" then "arial"                // Spanish
        elif lower = "本文のフォント" then "arial"                    // Japanese
        elif lower = "正文字体" then "arial"                          // Chinese (Simplified)
        elif lower = "본문 글꼴" then "arial"                         // Korean
        else "calibri"

let private fontFiles family bold italic =
    match family with
    | "arial" ->
        if bold && italic then ["arialbi.ttf"; "arialbd.ttf"; "arial.ttf"]
        elif bold then ["arialbd.ttf"; "arial.ttf"]
        elif italic then ["ariali.ttf"; "arial.ttf"]
        else ["arial.ttf"]
    | "cambria" ->
        // Cambria ships only regular/bold/italic — no bold-italic TTF.
        if bold && italic then ["cambriab.ttf"; "cambriai.ttf"; "cambria.ttc"]
        elif bold then ["cambriab.ttf"; "cambria.ttc"]
        elif italic then ["cambriai.ttf"; "cambria.ttc"]
        else ["cambria.ttc"]
    | "times" ->
        if bold && italic then ["timesbi.ttf"; "timesbd.ttf"; "times.ttf"]
        elif bold then ["timesbd.ttf"; "times.ttf"]
        elif italic then ["timesi.ttf"; "times.ttf"]
        else ["times.ttf"]
    | "verdana" ->
        if bold && italic then ["verdanaz.ttf"; "verdanab.ttf"; "verdana.ttf"]
        elif bold then ["verdanab.ttf"; "verdana.ttf"]
        elif italic then ["verdanai.ttf"; "verdana.ttf"]
        else ["verdana.ttf"]
    | "garamond" ->
        // Monotype Garamond (ships with Office) — no dedicated bold-italic.
        if bold && italic then ["GARABD.TTF"; "GARAIT.TTF"; "GARA.TTF"]
        elif bold then ["GARABD.TTF"; "GARA.TTF"]
        elif italic then ["GARAIT.TTF"; "GARA.TTF"]
        else ["GARA.TTF"]
    | "segoeui" ->
        if bold && italic then ["segoeuiz.ttf"; "segoeuib.ttf"; "segoeui.ttf"]
        elif bold then ["segoeuib.ttf"; "segoeui.ttf"]
        elif italic then ["segoeuii.ttf"; "segoeui.ttf"]
        else ["segoeui.ttf"]
    | _ ->
        if bold && italic then ["calibriz.ttf"; "calibrib.ttf"; "calibri.ttf"]
        elif bold then ["calibrib.ttf"; "calibri.ttf"]
        elif italic then ["calibrii.ttf"; "calibri.ttf"]
        else ["calibri.ttf"]

let private loadVariant family bold italic =
    let key = (family, bold, italic)
    match fontCache.TryGetValue key with
    | true, cached -> cached
    | _ ->
        let loaded =
            fontFiles family bold italic
            |> List.tryPick (fun name ->
                fontDirs ()
                |> List.tryPick (fun dir ->
                    let path = System.IO.Path.Combine(dir, name)
                    if System.IO.File.Exists path then
                        try Some (Renderer.Text.loadFontFile path) with _ -> None
                    else None))
        fontCache.[key] <- loaded
        loaded

let private rendererFont (fontName: string option) (bold: bool) (italic: bool) =
    let family = familyKey fontName
    match loadVariant family bold italic with
    | Some f -> f
    | None ->
        // Family lookup failed — try Calibri before falling back to the
        // built-in default so we don't lose bold/italic for Arial/Times etc.
        loadVariant "calibri" bold italic
        |> Option.defaultWith defaultFont

/// Public mirror of `rendererFont` so external callers (notably the native
/// runner, which needs to measure the workbook's default font for MDW)
/// resolve fonts through the same family map and cache as the renderer.
let rendererFontFor (fontName: string option) (bold: bool) (italic: bool) =
    rendererFont fontName bold italic

/// Load a CJK fallback by file name. Mirrors loadVariant's caching and
/// fontDirs lookup. Returns None when the file is absent (older Windows
/// images don't ship every face).
let private cjkFontCache = System.Collections.Generic.Dictionary<string, Renderer.Text.FontData option>()
let private loadCjkFile (fileName: string) : Renderer.Text.FontData option =
    match cjkFontCache.TryGetValue fileName with
    | true, cached -> cached
    | _ ->
        let loaded =
            fontDirs ()
            |> List.tryPick (fun dir ->
                let path = System.IO.Path.Combine(dir, fileName)
                if System.IO.File.Exists path then
                    try Some (Renderer.Text.loadFontFile path) with _ -> None
                else None)
        cjkFontCache.[fileName] <- loaded
        loaded

/// CJK fallback chain in the order Windows GDI / Excel resolves missing
/// glyphs when the primary font (Calibri, Arial, etc.) doesn't cover a
/// codepoint. Yu Gothic is the modern Japanese default; Microsoft YaHei
/// covers Simplified Chinese; JhengHei Traditional; Malgun Korean. MS
/// Gothic stays as a legacy fallback for fixtures with older glyph data.
let private cjkFallbackChain (bold: bool) : string list =
    if bold then
        [ "YuGothB.ttc"; "msgothic.ttc"; "msyhbd.ttc"; "msjhbd.ttc"; "malgunbd.ttf" ]
    else
        [ "YuGothR.ttc"; "msgothic.ttc"; "msyh.ttc"; "msjh.ttc"; "malgun.ttf" ]

/// For a codepoint not covered by the primary font, find the first font in
/// the fallback chain whose cmap maps it. Returns None when no chain font
/// is installed or none covers the codepoint (will render .notdef = box).
let private fallbackFontFor (codepoint: int) (bold: bool) : Renderer.Text.FontData option =
    cjkFallbackChain bold
    |> List.tryPick (fun name ->
        loadCjkFile name
        |> Option.bind (fun f -> if Renderer.Text.hasGlyph f codepoint then Some f else None))

let private parseHexColor (value: string) =
    let value = value.Trim()
    let parseByte (s: string) =
        match System.Byte.TryParse(s, System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture) with
        | true, value -> Some value
        | false, _ -> None
    let s = if value.StartsWith("#") then value.Substring(1) else value
    match s.Length with
    | 3 ->
        let expand (c: char) = string c + string c
        match parseByte (expand s.[0]), parseByte (expand s.[1]), parseByte (expand s.[2]) with
        | Some r, Some g, Some b -> Color.rgb r g b
        | _ -> Color.black
    | 6 ->
        match parseByte s.[0..1], parseByte s.[2..3], parseByte s.[4..5] with
        | Some r, Some g, Some b -> Color.rgb r g b
        | _ -> Color.black
    | 8 ->
        match parseByte s.[0..1], parseByte s.[2..3], parseByte s.[4..5], parseByte s.[6..7] with
        | Some r, Some g, Some b, Some a -> Color.rgba r g b a
        | _ -> Color.black
    | _ -> Color.black

/// Some fonts set OS/2 fsSelection bit 7 (USE_TYPO_METRICS) to opt into
/// typographic ascender/descender for line stacking. When that bit is
/// set we honor it; otherwise we use the GDI win metrics Excel anchors
/// its cell-text box layout to.
let private useTypoMetrics (font: Renderer.Text.FontData) =
    match font.OtFont.OS2 with
    | Some os2 -> (os2.FsSelection &&& 128us) <> 0us
    | None -> false

/// OS/2 winAscent scaled to pixels — the "WIN ascent" GDI/Excel uses as
/// the line top. Derived here from the (public) FontData.OtFont rather
/// than in Renderer.Text so the general-purpose renderer doesn't carry a
/// GDI-specific opinion. USE_TYPO_METRICS fonts get the typo ascender +
/// half the typo line gap instead; with no OS/2 table at all we fall
/// back to the hhea ascender.
let private winAscent (font: Renderer.Text.FontData) (fontSize: float) : float =
    let unitsPerEm = float font.OtFont.Head.UnitsPerEm
    let units =
        match font.OtFont.OS2 with
        | Some os2 when not (useTypoMetrics font) -> float os2.WinAscent
        | Some os2 -> float os2.TypoAscender + float os2.TypoLineGap / 2.0
        | None -> float font.OtFont.Hhea.Ascender
    units * fontSize / unitsPerEm

/// OS/2 winDescent as a positive pixel distance (baseline → em bottom).
/// Mirrors winAscent's USE_TYPO_METRICS / no-OS/2 fallbacks.
let private winDescent (font: Renderer.Text.FontData) (fontSize: float) : float =
    let unitsPerEm = float font.OtFont.Head.UnitsPerEm
    let units =
        match font.OtFont.OS2 with
        | Some os2 when not (useTypoMetrics font) -> float os2.WinDescent
        | Some os2 -> float -os2.TypoDescender + float os2.TypoLineGap / 2.0
        | None -> - (float font.OtFont.Hhea.Descender)
    units * fontSize / unitsPerEm

/// winAscent + winDescent — the GDI line-box height.
let private winLineHeight (font: Renderer.Text.FontData) (fontSize: float) : float =
    winAscent font fontSize + winDescent font fontSize

let private toRendererCommands (text: PaintText) =
    // Excel font sizes (XLSX <sz val="11"/>) are in points; the Renderer takes
    // pixel units. Convert pt→px so a Calibri 11-pt cell renders at Excel's
    // glyph height. Row/column heights are pre-converted via 96/72 upstream.
    let primary = rendererFont text.FontName text.Bold text.Italic
    let pxSize = text.FontSize * 96.0 / 72.0

    // Font-fallback for CJK / other glyphs not covered by the primary font.
    // Walk codepoints, partition into contiguous segments by which font
    // covers each, then layout/measure per segment. Excel uses GDI font
    // linking with similar fallback chains; we mirror the common case
    // (Latin primary → CJK fallback) so 切手/横浜/東京 etc. render rather
    // than printing .notdef boxes. Surrogate pairs are folded into one
    // segment with their lead surrogate's font choice.
    let fontFor (cp: int) : Renderer.Text.FontData =
        if Renderer.Text.hasGlyph primary cp then primary
        else
            match fallbackFontFor cp text.Bold with
            | Some fb -> fb
            | None -> primary  // no fallback installed — keep using primary, render as .notdef
    let segments : (Renderer.Text.FontData * string) list =
        let txt = text.Text
        let result = System.Collections.Generic.List<Renderer.Text.FontData * string>()
        let buf = System.Text.StringBuilder()
        let mutable currentFont : Renderer.Text.FontData voption = ValueNone
        let mutable i = 0
        while i < txt.Length do
            let ch = txt.[i]
            let cp, width =
                if System.Char.IsHighSurrogate ch && i + 1 < txt.Length && System.Char.IsLowSurrogate txt.[i+1] then
                    System.Char.ConvertToUtf32(ch, txt.[i+1]), 2
                else int ch, 1
            let f = fontFor cp
            match currentFont with
            | ValueNone -> currentFont <- ValueSome f
            | ValueSome cf when obj.ReferenceEquals(cf, f) -> ()
            | ValueSome cf ->
                result.Add(cf, buf.ToString())
                buf.Clear() |> ignore
                currentFont <- ValueSome f
            for k in 0 .. width - 1 do buf.Append(txt.[i + k]) |> ignore
            i <- i + width
        match currentFont with
        | ValueSome f when buf.Length > 0 -> result.Add(f, buf.ToString())
        | _ -> ()
        List.ofSeq result

    let segmentLayouts =
        segments
        |> List.map (fun (f, s) ->
            let l = Renderer.Text.layout f pxSize s
            let w = Renderer.Text.measureWidth f pxSize s
            f, l, w)
    let totalWidth = segmentLayouts |> List.sumBy (fun (_, _, w) -> w)
    // Cell-text rotation (Excel `<alignment textRotation>`): 1..90 = degrees
    // CCW, 91..180 = (value-90) degrees CW, 255 = stacked (handled as flat
    // here). The PDF renderer already rotates; the native/interactive path
    // laid every cell flat. We lay the text out horizontally from a leading-
    // edge anchor, then rotate the whole glyph stream around that anchor via
    // a Cmd.Group transform — mirroring the (Excel-validated) PDF path: anchor
    // at left + 4 px, baseline at the cell bottom for CCW (text rises) or at
    // the cell top + ascent for CW (text descends, staying inside the cell).
    let isRotated = text.Rotation > 0 && text.Rotation <> 255
    let alignedX =
        if isRotated then text.Rect.X + 4.0
        else
            match text.Align with
            | AlignLeft -> text.Rect.X
            | AlignCenter -> text.Rect.X + (text.Rect.Width - totalWidth) / 2.0
            | AlignRight -> text.Rect.X + text.Rect.Width - totalWidth
    // Baseline from OS/2 winAscent/winDescent of the PRIMARY font. Mixing
    // ascents across fallback fonts would jitter the baseline; sticking to
    // the primary's ascent matches Excel/GDI's font-linking behavior, which
    // keeps the baseline anchored to the primary font's metrics.
    let ascent = winAscent primary pxSize
    let lineHeight = winLineHeight primary pxSize
    let blockHeight = lineHeight * float text.LineCount
    // Excel quirk: vertical="top" is ignored on rows that aren't taller than
    // about one line-height. pandas/openpyxl emit `<alignment vertical="top"/>`
    // on every header cell, but Excel renders default-height rows as bottom
    // (Excel "General") regardless. Mirror that — when the cell has at most
    // ~half a line of vertical slack, collapse top → bottom. Tall wrap-text
    // rows still respect the hint and place text near the cell top.
    let slack = text.Rect.Height - blockHeight
    let effectiveVAlign =
        match text.VerticalAlign with
        | VAlignTop when slack < lineHeight * 0.5 -> VAlignBottom
        | va -> va
    let blockTopY =
        match effectiveVAlign with
        | VAlignTop -> text.Rect.Y
        | VAlignMiddle -> text.Rect.Y + (text.Rect.Height - blockHeight) / 2.0
        | VAlignBottom -> text.Rect.Y + text.Rect.Height - blockHeight
    // Super/subscript shifts the baseline relative to the line's normal
    // baseline. PaintText.BaselineShift is in pt at the cell's base size;
    // convert to px the same way pt → px conversion happens elsewhere.
    let baselineShiftPx = text.BaselineShift * 96.0 / 72.0
    // Rotated text is laid out at a provisional bottom-left position; the
    // Group transform below rotates it and then aligns its rotated bounding
    // box within the cell, so this exact baseline only needs to be consistent.
    let baseline =
        if isRotated then text.Rect.Y + text.Rect.Height - 2.0
        else blockTopY + float text.LineIndex * lineHeight + ascent + baselineShiftPx
    let color = Solid (parseHexColor text.Color)
    let mutable cursor = alignedX
    let inner = [
        for (_, layout, width) in segmentLayouts do
            yield Cmd.DrawGlyphs(vec2 cursor baseline, layout, color)
            cursor <- cursor + width
        let lineStyle = {
            Width = px 1.0
            Cap = Butt
            Join = Miter
            Paint = color
            DashArray = []
            DashOffset = px 0.0
            MiterLimit = 4.0
        }
        if text.Underline && totalWidth > 0.0 then
            // Single underline stroke spans all glyph segments. Position at
            // ~1 px below the baseline (matches Excel/GDI; OpenType's
            // OS/2.sxUnderlinePosition isn't surfaced by our font loader,
            // so use a Calibri-tuned constant). Stroke thickness 1 px.
            let underlineY = baseline + max 1.0 (pxSize * 0.08)
            yield Cmd.StrokePath(
                [ MoveTo (vec2 alignedX underlineY)
                  LineTo (vec2 (alignedX + totalWidth) underlineY) ],
                lineStyle)
            if text.UnderlineDouble then
                // Double underline: a second stroke ~2 px below the first.
                let underlineY2 = underlineY + max 2.0 (pxSize * 0.12)
                yield Cmd.StrokePath(
                    [ MoveTo (vec2 alignedX underlineY2)
                      LineTo (vec2 (alignedX + totalWidth) underlineY2) ],
                    lineStyle)
        if text.Strikethrough && totalWidth > 0.0 then
            // Strikethrough at ~30% above baseline — visually centers on
            // the lowercase x-height. (OpenType yStrikeoutPosition would
            // be more accurate; using a Calibri-tuned constant to match
            // Excel/GDI.)
            let strikeY = baseline - pxSize * 0.28
            yield Cmd.StrokePath(
                [ MoveTo (vec2 alignedX strikeY)
                  LineTo (vec2 (alignedX + totalWidth) strikeY) ],
                lineStyle)
    ]
    if not isRotated then inner
    else
        // Screen space is y-down so Matrix3x2.rotate(θ) turns clockwise;
        // Excel's 1..90 is CCW (θ negative) and 91..180 is CW (θ positive).
        let theta =
            if text.Rotation <= 90 then -(float text.Rotation) * System.Math.PI / 180.0
            else float (text.Rotation - 90) * System.Math.PI / 180.0
        // Rotate about the run's anchor, then translate so the rotated bounding
        // box is aligned within the cell the way Excel does (derived from a
        // 13-angle × 3-length sweep, tools/excel-parity/build-rotation-sweep.vbs):
        // RIGHT-align the box when 90 ≤ textRotation < 180 (CCW-vertical and
        // CW-angled), LEFT-align otherwise; always bottom-align. Bottom-align
        // also keeps vertical placement matching (rot 90 at y 88–122 vs 88–121).
        let descent = winDescent primary pxSize
        let cosT, sinT = cos theta, sin theta
        let rotPt x y =
            let dx, dy = x - alignedX, y - baseline
            alignedX + dx * cosT - dy * sinT, baseline + dx * sinT + dy * cosT
        let corners =
            [ rotPt alignedX (baseline - ascent)
              rotPt (alignedX + totalWidth) (baseline - ascent)
              rotPt alignedX (baseline + descent)
              rotPt (alignedX + totalWidth) (baseline + descent) ]
        let minX = corners |> List.map fst |> List.min
        let maxX = corners |> List.map fst |> List.max
        let maxY = corners |> List.map snd |> List.max
        let rightAlign = text.Rotation >= 90 && text.Rotation < 180
        let shiftX =
            if rightAlign then (text.Rect.X + text.Rect.Width - 2.0) - maxX
            else (text.Rect.X + 2.0) - minX
        let shiftY = (text.Rect.Y + text.Rect.Height - 2.0) - maxY
        let rot =
            Matrix3x2.multiply
                (Matrix3x2.multiply (Matrix3x2.translate -alignedX -baseline) (Matrix3x2.rotate theta))
                (Matrix3x2.translate alignedX baseline)
        let xform = Matrix3x2.multiply rot (Matrix3x2.translate shiftX shiftY)
        [ Cmd.Group(xform, 1.0, inner) ]

let rec private renderPaintItem item =
    match item with
    | FillRect (rect, color) ->
        [ Cmd.FillRect(Renderer.Types.rect rect.X rect.Y rect.Width rect.Height, parseHexColor color) ]
    | StrokeRect (rect, color, width) ->
        [ Cmd.StrokeRect(Renderer.Types.rect rect.X rect.Y rect.Width rect.Height, parseHexColor color, px width) ]
    | DrawText text -> toRendererCommands text
    | GridLine (x1, y1, x2, y2, color, width) ->
        // Pixel-snap axis-aligned strokes so they cover WHOLE pixels — like
        // Excel's GDI — instead of straddling the grid. A stroke centred on an
        // integer coordinate with an odd width spreads half a pixel either
        // side, so its edges anti-alias to ~50% grey (the "half the line is
        // lighter" artefact). Rounding the width and centring odd widths on a
        // pixel centre (n+0.5), even widths on a boundary (n), keeps the line
        // crisp. Diagonals (neither axis constant) are left unsnapped.
        let w = max 1.0 (System.Math.Round width)
        let snap (c: float) =
            if int w % 2 = 1 then floor c + 0.5 else System.Math.Round c
        let sx1, sy1, sx2, sy2 =
            if abs (y1 - y2) < 0.01 then let sy = snap y1 in x1, sy, x2, sy
            elif abs (x1 - x2) < 0.01 then let sx = snap x1 in sx, y1, sx, y2
            else x1, y1, x2, y2
        let style = {
            // Butt caps: with pixel-snapped strokes the box sides meet crisply
            // at the shared corner pixel. Square caps were tried but their
            // half-width extension lands off the pixel grid and anti-aliases to
            // a grey blob at the corner ("sticks out like a ball").
            Width = px w
            Cap = Butt
            Join = Miter
            Paint = Solid (parseHexColor color)
            DashArray = []
            DashOffset = px 0.0
            MiterLimit = 4.0
        }
        [ Cmd.StrokePath([MoveTo (vec2 sx1 sy1); LineTo (vec2 sx2 sy2)], style) ]
    | VecPath (segs, fill, stroke) ->
        let toSeg = function
            | VMove (x, y) -> MoveTo (vec2 x y)
            | VLine (x, y) -> LineTo (vec2 x y)
            | VCubic (x1, y1, x2, y2, x, y) -> CubicTo (vec2 x1 y1, vec2 x2 y2, vec2 x y)
            | VClose -> Close
        let ps = segs |> List.map toSeg
        [ match fill with
          | Some c -> yield Cmd.FillPath(ps, Solid (parseHexColor c), NonZero)
          | None -> ()
          match stroke with
          | Some (c, w) ->
              yield Cmd.StrokePath(ps, { Width = px w; Cap = Butt; Join = Miter
                                         Paint = Solid (parseHexColor c)
                                         DashArray = []; DashOffset = px 0.0; MiterLimit = 4.0 })
          | None -> () ]
    | VecGroup ((a, b, c, d, e, f), children) ->
        let xform = { M11 = a; M12 = b; M21 = c; M22 = d; DX = e; DY = f }
        [ Cmd.Group(xform, 1.0, children |> List.collect renderPaintItem) ]

/// Measure a text run in pixels using the real loaded font. Plugged into
/// Viewport.Measurer so Paint.wrapLines can break lines on actual glyph
/// width instead of mdwPx × charCount. Critical for non-Calibri-narrow
/// fonts where the approximation drifts (e.g. Times Roman fits more
/// characters per line than mdwPx assumes).
/// Split `text` into (font, substring) segments by primary-vs-fallback
/// coverage. Same logic the draw path uses in `toRendererCommands`, lifted
/// so the measurer can produce widths that match what we actually paint.
/// Without this, a run containing kanji measured against Times-only would
/// come back too narrow, and Paint.fs's rich-text cursor would step the
/// next run on top of the trailing characters.
let private segmentByFontCoverage
    (primary: Renderer.Text.FontData) (bold: bool) (text: string)
    : (Renderer.Text.FontData * string) list =
    let fontFor (cp: int) : Renderer.Text.FontData =
        if Renderer.Text.hasGlyph primary cp then primary
        else
            match fallbackFontFor cp bold with
            | Some fb -> fb
            | None -> primary
    let result = System.Collections.Generic.List<Renderer.Text.FontData * string>()
    let buf = System.Text.StringBuilder()
    let mutable currentFont : Renderer.Text.FontData voption = ValueNone
    let mutable i = 0
    while i < text.Length do
        let ch = text.[i]
        let cp, width =
            if System.Char.IsHighSurrogate ch && i + 1 < text.Length && System.Char.IsLowSurrogate text.[i + 1] then
                System.Char.ConvertToUtf32(ch, text.[i + 1]), 2
            else int ch, 1
        let f = fontFor cp
        match currentFont with
        | ValueNone -> currentFont <- ValueSome f
        | ValueSome cf when obj.ReferenceEquals(cf, f) -> ()
        | ValueSome cf ->
            result.Add(cf, buf.ToString())
            buf.Clear() |> ignore
            currentFont <- ValueSome f
        for k in 0 .. width - 1 do buf.Append(text.[i + k]) |> ignore
        i <- i + width
    match currentFont with
    | ValueSome f when buf.Length > 0 -> result.Add(f, buf.ToString())
    | _ -> ()
    List.ofSeq result

let private measureText
    (text: string) (fontPt: float)
    (fontName: string option) (bold: bool) (italic: bool) : float =
    let font = rendererFont fontName bold italic
    let pxSize = fontPt * 96.0 / 72.0
    // Walk codepoints with the same primary-then-fallback policy the draw
    // path uses. Otherwise a run like " (切手) " measured under Times comes
    // back as ~2 px of Latin chars (kanji counted as zero/.notdef width),
    // and Paint.fs's rich-text cursor positions the next run on top of the
    // trailing space — visible in phpsp-26template row 1 as "(切手)& "
    // instead of "(切手) &".
    segmentByFontCoverage font bold text
    |> List.sumBy (fun (f, s) -> Renderer.Text.measureWidth f pxSize s)

let viewportOf (model: SpreadsheetModel) width height : Viewport =
    // Match Excel's column/row header pixel sizes at 96 DPI / 100% zoom so
    // parity screenshots align cell-for-cell with our PNG renders. Measured
    // against artifacts/excel/pandas-test1-sheet1-headings.png: the column
    // header strip is 20 px tall (Calibri 11), the row header strip is 28 px
    // wide for two-digit row numbers, and each row in the default 15-pt
    // height renders at 21 source px.
    let headerHeight = 20.0
    let rowHeaderWidth = 28.0
    { ScrollX = 0.0
      ScrollY = 0.0
      Width = width
      Height = height
      HeaderHeight = headerHeight
      RowHeaderWidth = rowHeaderWidth
      Measurer = Some measureText }

let private cellCommands (model: SpreadsheetModel) width height =
    let viewport = viewportOf model width height
    buildPaintItems model viewport
    |> List.collect renderPaintItem

let widget width height (model: SpreadsheetModel) =
    canvas width height (fun w h -> cellCommands model w h)

/// Render the sheet with extra paint items (e.g. charts) overlaid on top, in
/// the same viewport pixel space as the cells.
let widgetWith width height (model: SpreadsheetModel) (extra: PaintItem list) =
    canvas width height (fun w h -> (cellCommands model w h) @ (extra |> List.collect renderPaintItem))

/// Render a standalone vector PaintItem scene (no model) on a white
/// background — used to prove the VecPath/VecGroup primitives end-to-end.
let renderScene width height (items: PaintItem list) =
    canvas width height (fun w h ->
        Cmd.FillRect(Renderer.Types.rect 0.0 0.0 w h, parseHexColor "#FFFFFF")
        :: (items |> List.collect renderPaintItem))
