module Composit.Ui.SpreadsheetGrid

open Renderer
open Composit.Ui.Theme
open Composit.Ui.Widget

open RedBat.Spreadsheet.Model
open RedBat.Spreadsheet.Layout
open RedBat.Spreadsheet.Paint

open System.Collections.Generic
open System.Runtime.CompilerServices

let inline private px (v: float) : float<px> = LanguagePrimitives.FloatWithMeasure v

let private resizeArrayToList (items: System.Collections.Generic.List<'T>) =
    let mutable result = []
    for i = items.Count - 1 downto 0 do
        result <- items.[i] :: result
    result

let private clearArray (items: 'T array) =
    System.Array.Clear(items, 0, items.Length)

let private defaultFont () = Composit.Ui.Layout.font()

/// Cache of font variants keyed by (familyKey, bold, italic). familyKey is the
/// normalized lower-case font name we use to pick a TTF set ("calibri",
/// "arial", "cambria", "times"); unknown families fall back to "calibri".
let private fontCache = System.Collections.Generic.Dictionary<string * bool * bool, Renderer.Text.FontData option>()

let private trimBounds (value: string) =
    let mutable start = 0
    let mutable last = value.Length - 1
    while start <= last && System.Char.IsWhiteSpace value.[start] do
        start <- start + 1
    while last >= start && System.Char.IsWhiteSpace value.[last] do
        last <- last - 1
    start, last - start + 1

let private startsWithAtIgnoreCase (value: string) start length (prefix: string) =
    length >= prefix.Length
    && System.String.Compare(value, start, prefix, 0, prefix.Length, System.StringComparison.OrdinalIgnoreCase) = 0

let private equalsAtIgnoreCase (value: string) start length (expected: string) =
    length = expected.Length
    && System.String.Compare(value, start, expected, 0, length, System.StringComparison.OrdinalIgnoreCase) = 0

let private fontDirs () =
    let cwd = System.IO.Directory.GetCurrentDirectory()
    let dirs = System.Collections.Generic.List<string>()
    for dir in [
        System.Environment.GetEnvironmentVariable("RB_WORD_FONT_DIR")
        System.IO.Path.GetFullPath(System.IO.Path.Combine(cwd, "artifacts/fonts/windows"))
        System.IO.Path.GetFullPath(System.IO.Path.Combine(cwd, "../../artifacts/fonts/windows"))
        // Sibling office repo's font cache, when running from office-spreadsheet/.
        System.IO.Path.GetFullPath(System.IO.Path.Combine(cwd, "../office/artifacts/fonts/windows"))
    ] do
        if not (System.String.IsNullOrWhiteSpace dir) && System.IO.Directory.Exists dir then
            dirs.Add dir
    resizeArrayToList dirs

/// Map an XLSX font `name val="..."` to a normalized family key. Anything we
/// don't have a TTF for falls back to "calibri" — the renderer keeps working
/// even when we encounter Asian/script fonts we can't substitute properly.
let private familyKey (name: string option) =
    match name with
    | None -> "calibri"
    | Some n ->
        let start, length = trimBounds n
        if length = 0 then "calibri"
        elif startsWithAtIgnoreCase n start length "arial" then "arial"
        elif startsWithAtIgnoreCase n start length "calibri" then "calibri"
        elif startsWithAtIgnoreCase n start length "cambria" then "cambria"
        elif startsWithAtIgnoreCase n start length "times" then "times"
        elif startsWithAtIgnoreCase n start length "verdana" then "verdana"
        elif startsWithAtIgnoreCase n start length "garamond" then "garamond"
        elif startsWithAtIgnoreCase n start length "segoe" then "segoeui"
        // Localized "Body Text Font" names from non-English Excel installs.
        // OOXML stores the localized display name verbatim; Windows resolves
        // it through the OS locale. For our purposes Arial is the closest
        // Excel substitute on a system that doesn't have the localized
        // font (verified against pandas-testdateoverflow's Polish
        // "Czcionka tekstu podstawowego" — column widths and row heights
        // match Excel's reported values with Arial metrics, not Calibri's).
        elif equalsAtIgnoreCase n start length "czcionka tekstu podstawowego" then "arial"     // Polish
        elif equalsAtIgnoreCase n start length "schrift des textkörpers" then "arial"          // German
        elif equalsAtIgnoreCase n start length "police du corps" then "arial"                  // French
        elif equalsAtIgnoreCase n start length "carattere del corpo" then "arial"              // Italian
        elif equalsAtIgnoreCase n start length "fuente del cuerpo" then "arial"                // Spanish
        elif equalsAtIgnoreCase n start length "本文のフォント" then "arial"                    // Japanese
        elif equalsAtIgnoreCase n start length "正文字体" then "arial"                          // Chinese (Simplified)
        elif equalsAtIgnoreCase n start length "본문 글꼴" then "arial"                         // Korean
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
    let font =
        match loadVariant family bold italic with
        | Some f -> f
        | None ->
            // Family lookup failed — try Calibri before falling back to the
            // built-in default so we don't lose bold/italic for Arial/Times etc.
            loadVariant "calibri" bold italic
            |> Option.defaultWith defaultFont
    font

let private rendererFontForFamily family (bold: bool) (italic: bool) =
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
    let hexValue (c: char) =
        if c >= '0' && c <= '9' then int c - int '0'
        elif c >= 'a' && c <= 'f' then int c - int 'a' + 10
        elif c >= 'A' && c <= 'F' then int c - int 'A' + 10
        else -1
    let parseByteAt offset =
        let hi = hexValue value.[offset]
        let lo = hexValue value.[offset + 1]
        if hi < 0 || lo < 0 then None else Some (byte ((hi <<< 4) ||| lo))
    let parseShortAt offset =
        let v = hexValue value.[offset]
        if v < 0 then None else Some (byte ((v <<< 4) ||| v))
    let start0, length0 = trimBounds value
    let start, length =
        if length0 > 0 && value.[start0] = '#' then start0 + 1, length0 - 1
        else start0, length0
    match length with
    | 3 ->
        match parseShortAt start, parseShortAt (start + 1), parseShortAt (start + 2) with
        | Some r, Some g, Some b -> Color.rgb r g b
        | _ -> Color.black
    | 6 ->
        match parseByteAt start, parseByteAt (start + 2), parseByteAt (start + 4) with
        | Some r, Some g, Some b -> Color.rgb r g b
        | _ -> Color.black
    | 8 ->
        match parseByteAt start, parseByteAt (start + 2), parseByteAt (start + 4), parseByteAt (start + 6) with
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
    let segments =
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
        result

    let segmentLayouts = System.Collections.Generic.List<_>(segments.Count)
    let mutable totalWidth = 0.0
    for f, s in segments do
        let layout = Renderer.Text.layout f pxSize s
        let width = Renderer.Text.measureWidth f pxSize s
        segmentLayouts.Add(f, layout, width)
        totalWidth <- totalWidth + width
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
    let innerCommands = System.Collections.Generic.List<Cmd>(segmentLayouts.Count + 3)
    for _, layout, width in segmentLayouts do
        innerCommands.Add(Cmd.DrawGlyphs(vec2 cursor baseline, layout, color))
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
        innerCommands.Add(Cmd.StrokePath(
            [ MoveTo (vec2 alignedX underlineY)
              LineTo (vec2 (alignedX + totalWidth) underlineY) ],
            lineStyle))
        if text.UnderlineDouble then
            // Double underline: a second stroke ~2 px below the first.
            let underlineY2 = underlineY + max 2.0 (pxSize * 0.12)
            innerCommands.Add(Cmd.StrokePath(
                [ MoveTo (vec2 alignedX underlineY2)
                  LineTo (vec2 (alignedX + totalWidth) underlineY2) ],
                lineStyle))
    if text.Strikethrough && totalWidth > 0.0 then
        // Strikethrough at ~30% above baseline — visually centers on
        // the lowercase x-height. (OpenType yStrikeoutPosition would
        // be more accurate; using a Calibri-tuned constant to match
        // Excel/GDI.)
        let strikeY = baseline - pxSize * 0.28
        innerCommands.Add(Cmd.StrokePath(
            [ MoveTo (vec2 alignedX strikeY)
              LineTo (vec2 (alignedX + totalWidth) strikeY) ],
            lineStyle))
    let inner = resizeArrayToList innerCommands
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
        let x1, y1 = rotPt alignedX (baseline - ascent)
        let x2, y2 = rotPt (alignedX + totalWidth) (baseline - ascent)
        let x3, y3 = rotPt alignedX (baseline + descent)
        let x4, y4 = rotPt (alignedX + totalWidth) (baseline + descent)
        let minX = min (min x1 x2) (min x3 x4)
        let maxX = max (max x1 x2) (max x3 x4)
        let maxY = max (max y1 y2) (max y3 y4)
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

let rec private appendPaintItem (commands: System.Collections.Generic.List<Cmd>) item =
    match item with
    | FillRect (rect, color) ->
        commands.Add(Cmd.FillRect(Renderer.Types.rect rect.X rect.Y rect.Width rect.Height, parseHexColor color))
    | StrokeRect (rect, color, width) ->
        commands.Add(Cmd.StrokeRect(Renderer.Types.rect rect.X rect.Y rect.Width rect.Height, parseHexColor color, px width))
    | DrawText text -> commands.AddRange(toRendererCommands text)
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
        commands.Add(Cmd.StrokePath([MoveTo (vec2 sx1 sy1); LineTo (vec2 sx2 sy2)], style))
    | VecPath (segs, fill, stroke) ->
        let pathSegments = System.Collections.Generic.List<_>()
        for seg in segs do
            match seg with
            | VMove (x, y) -> pathSegments.Add(MoveTo (vec2 x y))
            | VLine (x, y) -> pathSegments.Add(LineTo (vec2 x y))
            | VCubic (x1, y1, x2, y2, x, y) -> pathSegments.Add(CubicTo (vec2 x1 y1, vec2 x2 y2, vec2 x y))
            | VClose -> pathSegments.Add Close
        let ps = resizeArrayToList pathSegments
        match fill with
        | Some c -> commands.Add(Cmd.FillPath(ps, Solid (parseHexColor c), NonZero))
        | None -> ()
        match stroke with
        | Some (c, w) ->
            commands.Add(Cmd.StrokePath(ps, { Width = px w; Cap = Butt; Join = Miter
                                              Paint = Solid (parseHexColor c)
                                              DashArray = []; DashOffset = px 0.0; MiterLimit = 4.0 }))
        | None -> ()
    | VecGroup ((a, b, c, d, e, f), children) ->
        let xform = { M11 = a; M12 = b; M21 = c; M22 = d; DX = e; DY = f }
        let childCommands = System.Collections.Generic.List<Cmd>()
        for child in children do
            appendPaintItem childCommands child
        commands.Add(Cmd.Group(xform, 1.0, resizeArrayToList childCommands))

let private renderPaintItem item =
    let commands = System.Collections.Generic.List<Cmd>()
    appendPaintItem commands item
    resizeArrayToList commands

type private RendererPaintItemsCacheEntry = {
    PaintItems: PaintItem list
    Commands: Cmd list
    CommandCount: int
    TextCharCount: int
}

let mutable private cellRendererCommandsCache: RendererPaintItemsCacheEntry list = []
let private paintItemsReferenceComparer =
    { new IEqualityComparer<PaintItem list> with
        member _.Equals(left, right) = obj.ReferenceEquals(left, right)
        member _.GetHashCode(value) = RuntimeHelpers.GetHashCode(value) }
let private cellRendererCommandsCacheByItems =
    Dictionary<PaintItem list, RendererPaintItemsCacheEntry>(paintItemsReferenceComparer)
let private cellRendererCommandsFastSlots = Array.zeroCreate<RendererPaintItemsCacheEntry> 256
let private cellRendererCommandsCacheMaxEntries = 256
let private cellRendererCommandsCacheMaxCommands = 200000
let private cellRendererCommandsCacheMaxTextChars = 5000000

let private rendererCommandsPaintItemsSlot paintItems slotCount =
    RuntimeHelpers.GetHashCode(paintItems) &&& (slotCount - 1)

let private addCapped limit value total =
    if total > limit then
        total
    else
        let value = max 0 value
        let remaining = limit - total
        if value > remaining then limit + 1 else total + value

let rec private paintItemTextCharCount limit total item =
    match item with
    | DrawText text -> addCapped limit text.Text.Length total
    | VecGroup (_, children) ->
        let mutable next = total
        let mutable remaining = children
        while next <= limit && not remaining.IsEmpty do
            next <- paintItemTextCharCount limit next remaining.Head
            remaining <- remaining.Tail
        next
    | _ -> total

let private paintItemsTextCharCount limit (items: PaintItem list) =
    let mutable total = 0
    let mutable remaining = items
    while total <= limit && not remaining.IsEmpty do
        total <- paintItemTextCharCount limit total remaining.Head
        remaining <- remaining.Tail
    total

let private rendererCommandsCacheEntryMatches paintItems entry =
    obj.ReferenceEquals(entry.PaintItems, paintItems)

let private setRendererCommandsCache entries =
    cellRendererCommandsCache <- entries
    cellRendererCommandsCacheByItems.Clear()
    clearArray cellRendererCommandsFastSlots
    for entry in entries do
        cellRendererCommandsCacheByItems.[entry.PaintItems] <- entry
        cellRendererCommandsFastSlots.[rendererCommandsPaintItemsSlot entry.PaintItems cellRendererCommandsFastSlots.Length] <- entry

let private tryGetRendererCommandsForPaintItems paintItems =
    let slot = rendererCommandsPaintItemsSlot paintItems cellRendererCommandsFastSlots.Length
    let fast = cellRendererCommandsFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && obj.ReferenceEquals(fast.PaintItems, paintItems) then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<RendererPaintItemsCacheEntry>
        if cellRendererCommandsCacheByItems.TryGetValue(paintItems, &cached) then
            cellRendererCommandsFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private rememberRendererCommandsForPaintItems paintItems commands commandCount textCharCount =
    let rec removeMatching acc remaining =
        match remaining with
        | [] -> List.rev acc
        | candidate :: tail when rendererCommandsCacheEntryMatches paintItems candidate ->
            List.rev acc @ tail
        | candidate :: tail ->
            removeMatching (candidate :: acc) tail
    if commandCount > cellRendererCommandsCacheMaxCommands
       || textCharCount > cellRendererCommandsCacheMaxTextChars then
        setRendererCommandsCache (removeMatching [] cellRendererCommandsCache)
    else
        let entry = {
            PaintItems = paintItems
            Commands = commands
            CommandCount = commandCount
            TextCharCount = textCharCount
        }
        let trim entries =
            let rec loop keptCount keptCommands keptTextChars acc remaining =
                match remaining with
                | [] -> List.rev acc
                | candidate :: tail ->
                    let candidateCommands = max 1 candidate.CommandCount
                    let candidateTextChars = max 1 candidate.TextCharCount
                    if keptCount < cellRendererCommandsCacheMaxEntries
                       && keptCommands + candidateCommands <= cellRendererCommandsCacheMaxCommands
                       && keptTextChars + candidateTextChars <= cellRendererCommandsCacheMaxTextChars then
                        loop
                            (keptCount + 1)
                            (keptCommands + candidateCommands)
                            (keptTextChars + candidateTextChars)
                            (candidate :: acc)
                            tail
                    else
                        loop keptCount keptCommands keptTextChars acc tail
            loop 0 0 0 [] entries
        (entry :: removeMatching [] cellRendererCommandsCache)
        |> trim
        |> setRendererCommandsCache

let private rendererCommandsForPaintItems (paintItems: PaintItem list) =
    match tryGetRendererCommandsForPaintItems paintItems with
    | ValueSome cached ->
        cached.Commands
    | ValueNone ->
        let commands = System.Collections.Generic.List<Cmd>()
        for item in paintItems do
            appendPaintItem commands item
        let rendered = resizeArrayToList commands
        let commandCount = commands.Count
        let textCharCount =
            if commandCount <= cellRendererCommandsCacheMaxCommands then
                paintItemsTextCharCount cellRendererCommandsCacheMaxTextChars paintItems
            else
                cellRendererCommandsCacheMaxTextChars + 1
        rememberRendererCommandsForPaintItems paintItems rendered commandCount textCharCount
        rendered

type private ViewportFrameCommandsCacheEntry = {
    Commands: Cmd list
    ScrollX: float
    ScrollY: float
    Width: float
    Height: float
    Result: Cmd list
    CommandCount: int
}

[<Struct>]
type private ViewportFrameCommandsKey = {
    Commands: Cmd list
    ScrollX: float
    ScrollY: float
    Width: float
    Height: float
}

let mutable private viewportFrameCommandsCache: ViewportFrameCommandsCacheEntry list = []
let private viewportFrameCommandsKeyComparer =
    { new IEqualityComparer<ViewportFrameCommandsKey> with
        member _.Equals(left, right) =
            obj.ReferenceEquals(left.Commands, right.Commands)
            && left.ScrollX = right.ScrollX
            && left.ScrollY = right.ScrollY
            && left.Width = right.Width
            && left.Height = right.Height
        member _.GetHashCode(value) =
            let mutable hashCode = RuntimeHelpers.GetHashCode(value.Commands)
            hashCode <- (hashCode * 397) ^^^ value.ScrollX.GetHashCode()
            hashCode <- (hashCode * 397) ^^^ value.ScrollY.GetHashCode()
            hashCode <- (hashCode * 397) ^^^ value.Width.GetHashCode()
            (hashCode * 397) ^^^ value.Height.GetHashCode() }
let private viewportFrameCommandsCacheByKey =
    Dictionary<ViewportFrameCommandsKey, ViewportFrameCommandsCacheEntry>(viewportFrameCommandsKeyComparer)
let private viewportFrameCommandsFastSlots = Array.zeroCreate<ViewportFrameCommandsCacheEntry> 256
let private viewportFrameCommandsCacheMaxEntries = 256
let private viewportFrameCommandsCacheMaxCommands = 200000

let private viewportFrameCommandsKey commands scrollX scrollY width height : ViewportFrameCommandsKey = {
    Commands = commands
    ScrollX = scrollX
    ScrollY = scrollY
    Width = width
    Height = height
}

let private viewportFrameCommandsSlot commands scrollX scrollY width height slotCount =
    let mutable hashCode = RuntimeHelpers.GetHashCode(commands)
    hashCode <- (hashCode * 397) ^^^ scrollX.GetHashCode()
    hashCode <- (hashCode * 397) ^^^ scrollY.GetHashCode()
    hashCode <- (hashCode * 397) ^^^ width.GetHashCode()
    hashCode <- (hashCode * 397) ^^^ height.GetHashCode()
    hashCode &&& (slotCount - 1)

let private viewportFrameCommandsEntryMatches commands scrollX scrollY width height (entry: ViewportFrameCommandsCacheEntry) =
    obj.ReferenceEquals(entry.Commands, commands)
    && entry.ScrollX = scrollX
    && entry.ScrollY = scrollY
    && entry.Width = width
    && entry.Height = height

let private setViewportFrameCommandsCache entries =
    viewportFrameCommandsCache <- entries
    viewportFrameCommandsCacheByKey.Clear()
    clearArray viewportFrameCommandsFastSlots
    for entry in entries do
        let key =
            viewportFrameCommandsKey
                entry.Commands
                entry.ScrollX
                entry.ScrollY
                entry.Width
                entry.Height
        viewportFrameCommandsCacheByKey.[key] <- entry
        let slot =
            viewportFrameCommandsSlot
                entry.Commands
                entry.ScrollX
                entry.ScrollY
                entry.Width
                entry.Height
                viewportFrameCommandsFastSlots.Length
        viewportFrameCommandsFastSlots.[slot] <- entry

let private viewportFrameCommandCount (commands: Cmd list) =
    let mutable count = 0
    let mutable remaining = commands
    while count <= viewportFrameCommandsCacheMaxCommands && not remaining.IsEmpty do
        count <- count + 1
        remaining <- remaining.Tail
    count

let private tryGetViewportFrameCommands (commands: Cmd list) scrollX scrollY width height =
    let slot =
        viewportFrameCommandsSlot
            commands
            scrollX
            scrollY
            width
            height
            viewportFrameCommandsFastSlots.Length
    let fast = viewportFrameCommandsFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && viewportFrameCommandsEntryMatches commands scrollX scrollY width height fast then
        ValueSome fast.Result
    else
        let key = viewportFrameCommandsKey commands scrollX scrollY width height
        let mutable cached = Unchecked.defaultof<ViewportFrameCommandsCacheEntry>
        if viewportFrameCommandsCacheByKey.TryGetValue(key, &cached) then
            viewportFrameCommandsFastSlots.[slot] <- cached
            ValueSome cached.Result
        else
            ValueNone

let private rememberViewportFrameCommands (commands: Cmd list) scrollX scrollY width height result commandCount =
    if commandCount <= viewportFrameCommandsCacheMaxCommands then
        let entry: ViewportFrameCommandsCacheEntry = {
            Commands = commands
            ScrollX = scrollX
            ScrollY = scrollY
            Width = width
            Height = height
            Result = result
            CommandCount = commandCount
        }
        let rec removeMatching acc remaining =
            match remaining with
            | [] -> List.rev acc
            | candidate :: tail when viewportFrameCommandsEntryMatches commands scrollX scrollY width height candidate ->
                List.rev acc @ tail
            | candidate :: tail ->
                removeMatching (candidate :: acc) tail
        let trim entries =
            let rec loop keptCount keptCommands acc remaining =
                match remaining with
                | [] -> List.rev acc
                | candidate :: tail ->
                    let commandCount = max 1 candidate.CommandCount
                    if keptCount < viewportFrameCommandsCacheMaxEntries
                       && keptCommands + commandCount <= viewportFrameCommandsCacheMaxCommands then
                        loop (keptCount + 1) (keptCommands + commandCount) (candidate :: acc) tail
                    else
                        loop keptCount keptCommands acc tail
            loop 0 0 [] entries
        (entry :: removeMatching [] viewportFrameCommandsCache)
        |> trim
        |> setViewportFrameCommandsCache

let private commandsWithViewportFrame scrollX scrollY viewportWidth viewportHeight (commands: Cmd list) =
    match tryGetViewportFrameCommands commands scrollX scrollY viewportWidth viewportHeight with
    | ValueSome result -> result
    | ValueNone ->
        let commandCount = viewportFrameCommandCount commands
        let result =
            Cmd.FillRect(
                Renderer.Types.rect scrollX scrollY viewportWidth viewportHeight,
                parseHexColor "#FFFFFF")
            :: commands
        rememberViewportFrameCommands commands scrollX scrollY viewportWidth viewportHeight result (commandCount + 1)
        result

type private SparseDefaultRendererFrameCacheEntry = {
    Parts: SparseDefaultPaintParts
    PrefixItems: PaintItem list
    SelectionFills: PaintItem list
    GridItems: PaintItem list
    SelectionStrokes: PaintItem list
    TextItems: PaintItem list
    SuffixItems: PaintItem list
    Commands: Cmd list
    CommandCount: int
}

let mutable private sparseDefaultRendererFrameCache: SparseDefaultRendererFrameCacheEntry list = []
let private sparseDefaultPartsReferenceComparer =
    { new IEqualityComparer<SparseDefaultPaintParts> with
        member _.Equals(left, right) = obj.ReferenceEquals(left, right)
        member _.GetHashCode(value) = RuntimeHelpers.GetHashCode(value) }
let private sparseDefaultRendererFrameCacheByParts =
    Dictionary<SparseDefaultPaintParts, SparseDefaultRendererFrameCacheEntry>(sparseDefaultPartsReferenceComparer)
let private sparseDefaultRendererFrameFastSlots = Array.zeroCreate<SparseDefaultRendererFrameCacheEntry> 256
let private sparseDefaultRendererFrameCacheMaxEntries = 256
let private sparseDefaultRendererFrameCacheMaxCommands = 200000

let private sparseDefaultPartsSlot parts slotCount =
    RuntimeHelpers.GetHashCode(parts) &&& (slotCount - 1)

let private sparseDefaultRendererFrameMatches (parts: SparseDefaultPaintParts) (entry: SparseDefaultRendererFrameCacheEntry) =
    obj.ReferenceEquals(entry.Parts, parts)
    || (entry.PrefixItems = parts.PrefixItems
        && obj.ReferenceEquals(entry.GridItems, parts.GridItems)
        && obj.ReferenceEquals(entry.TextItems, parts.TextItems)
        && entry.SelectionFills = parts.SelectionFills
        && entry.SelectionStrokes = parts.SelectionStrokes
        && entry.SuffixItems = parts.SuffixItems)

let private setSparseDefaultRendererFrameCache entries =
    sparseDefaultRendererFrameCache <- entries
    sparseDefaultRendererFrameCacheByParts.Clear()
    clearArray sparseDefaultRendererFrameFastSlots
    for entry in entries do
        sparseDefaultRendererFrameCacheByParts.[entry.Parts] <- entry
        sparseDefaultRendererFrameFastSlots.[sparseDefaultPartsSlot entry.Parts sparseDefaultRendererFrameFastSlots.Length] <- entry

let private tryGetSparseDefaultRendererFrame parts =
    let slot = sparseDefaultPartsSlot parts sparseDefaultRendererFrameFastSlots.Length
    let fast = sparseDefaultRendererFrameFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && sparseDefaultRendererFrameMatches parts fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<SparseDefaultRendererFrameCacheEntry>
        if sparseDefaultRendererFrameCacheByParts.TryGetValue(parts, &cached) then
            sparseDefaultRendererFrameFastSlots.[slot] <- cached
            ValueSome cached
        else
            match sparseDefaultRendererFrameCache |> List.tryFind (sparseDefaultRendererFrameMatches parts) with
            | Some cached ->
                sparseDefaultRendererFrameFastSlots.[slot] <- cached
                ValueSome cached
            | None -> ValueNone

let private rememberSparseDefaultRendererFrame (parts: SparseDefaultPaintParts) commands commandCount =
    if commandCount > sparseDefaultRendererFrameCacheMaxCommands then
        ()
    else
        let entry = {
            Parts = parts
            PrefixItems = parts.PrefixItems
            SelectionFills = parts.SelectionFills
            GridItems = parts.GridItems
            SelectionStrokes = parts.SelectionStrokes
            TextItems = parts.TextItems
            SuffixItems = parts.SuffixItems
            Commands = commands
            CommandCount = commandCount
        }
        let rec removeMatching acc remaining =
            match remaining with
            | [] -> List.rev acc
            | candidate :: tail when sparseDefaultRendererFrameMatches parts candidate ->
                List.rev acc @ tail
            | candidate :: tail ->
                removeMatching (candidate :: acc) tail
        let trim entries =
            let rec loop keptCount keptCommands acc remaining =
                match remaining with
                | [] -> List.rev acc
                | candidate :: tail ->
                    let commandCount = max 1 candidate.CommandCount
                    if keptCount < sparseDefaultRendererFrameCacheMaxEntries
                       && keptCommands + commandCount <= sparseDefaultRendererFrameCacheMaxCommands then
                        loop
                            (keptCount + 1)
                            (keptCommands + commandCount)
                            (candidate :: acc)
                            tail
                    else
                        loop keptCount keptCommands acc tail
            loop 0 0 [] entries
        (entry :: removeMatching [] sparseDefaultRendererFrameCache)
        |> trim
        |> setSparseDefaultRendererFrameCache

let private addRendererCommands (target: System.Collections.Generic.List<Cmd>) commands =
    for command in commands do
        target.Add command

let private groupedRendererCommands (commands: Cmd list) =
    if commands.IsEmpty then []
    else [ Cmd.Group(Matrix3x2.identity, 1.0, commands) ]

let private appendPaintItemsAndCount (target: System.Collections.Generic.List<Cmd>) (paintItems: PaintItem list) =
    let before = target.Count
    for item in paintItems do
        appendPaintItem target item
    target.Count - before

let private rendererCommandsForSparseDefaultParts (parts: SparseDefaultPaintParts) =
    match tryGetSparseDefaultRendererFrame parts with
    | ValueSome cached -> cached.Commands
    | ValueNone ->
        let prefixCommands = rendererCommandsForPaintItems parts.PrefixItems
        let gridCommands = rendererCommandsForPaintItems parts.GridItems
        let textCommands = rendererCommandsForPaintItems parts.TextItems
        let suffixCommands = rendererCommandsForPaintItems parts.SuffixItems
        let renderedCommands = System.Collections.Generic.List<Cmd>()
        addRendererCommands renderedCommands (groupedRendererCommands prefixCommands)
        let fillCommandCount = appendPaintItemsAndCount renderedCommands parts.SelectionFills
        addRendererCommands renderedCommands (groupedRendererCommands gridCommands)
        let strokeCommandCount = appendPaintItemsAndCount renderedCommands parts.SelectionStrokes
        addRendererCommands renderedCommands (groupedRendererCommands textCommands)
        addRendererCommands renderedCommands (groupedRendererCommands suffixCommands)
        let rendered = resizeArrayToList renderedCommands
        let commandCount =
            prefixCommands.Length
            + fillCommandCount
            + gridCommands.Length
            + strokeCommandCount
            + textCommands.Length
            + suffixCommands.Length
        rememberSparseDefaultRendererFrame parts rendered commandCount
        rendered

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
    : System.Collections.Generic.List<Renderer.Text.FontData * string> =
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
    result

let private measureTextUncached
    (text: string) (fontPt: float)
    family (bold: bool) (italic: bool) : float =
    let font = rendererFontForFamily family bold italic
    let pxSize = fontPt * 96.0 / 72.0
    // Walk codepoints with the same primary-then-fallback policy the draw
    // path uses. Otherwise a run like " (切手) " measured under Times comes
    // back as ~2 px of Latin chars (kanji counted as zero/.notdef width),
    // and Paint.fs's rich-text cursor positions the next run on top of the
    // trailing space — visible in phpsp-26template row 1 as "(切手)& "
    // instead of "(切手) &".
    let mutable width = 0.0
    for f, s in segmentByFontCoverage font bold text do
        width <- width + Renderer.Text.measureWidth f pxSize s
    width

[<Struct>]
type private TextMeasureCacheKey = {
    Text: string
    FontPt: float
    Family: string
    Bold: bool
    Italic: bool
}

let private textMeasureCache = System.Collections.Generic.Dictionary<TextMeasureCacheKey, float>()
let private textMeasureCacheOrder = System.Collections.Generic.Queue<TextMeasureCacheKey>()
let private textMeasureCacheMaxEntries = 200000
let private textMeasureCacheMaxTotalTextChars = 5000000
let private textMeasureCacheMaxSingleTextChars = 32768
let private textMeasureFastFontNameMaxChars = 256
let mutable private textMeasureCacheTotalTextChars = 0
let mutable private textMeasureLastSet = false
let mutable private textMeasureLastText = ""
let mutable private textMeasureLastFontPt = 0.0
let mutable private textMeasureLastFontName: string option = None
let mutable private textMeasureLastBold = false
let mutable private textMeasureLastItalic = false
let mutable private textMeasureLastWidth = 0.0

let private textMeasureCacheKeyChars key =
    key.Text.Length + key.Family.Length

let private sameMeasureString (left: string) (right: string) =
    obj.ReferenceEquals(left, right) || left = right

let private sameMeasureStringOption (left: string option) (right: string option) =
    match left with
    | None ->
        match right with
        | None -> true
        | Some _ -> false
    | Some left ->
        match right with
        | Some right -> sameMeasureString left right
        | None -> false

let private measureFontNameRetainable (fontName: string option) =
    match fontName with
    | Some value -> value.Length <= textMeasureFastFontNameMaxChars
    | None -> true

let private textMeasureFastKeyRetainable (text: string) (fontName: string option) =
    text.Length <= textMeasureCacheMaxSingleTextChars
    && measureFontNameRetainable fontName

let private tryLastTextMeasure (text: string) fontPt (fontName: string option) bold italic =
    if textMeasureLastSet
       && textMeasureLastFontPt = fontPt
       && textMeasureLastBold = bold
       && textMeasureLastItalic = italic
       && sameMeasureString textMeasureLastText text
       && sameMeasureStringOption textMeasureLastFontName fontName then
        ValueSome textMeasureLastWidth
    else
        ValueNone

let private rememberLastTextMeasure (text: string) fontPt (fontName: string option) bold italic width =
    if textMeasureFastKeyRetainable text fontName then
        textMeasureLastSet <- true
        textMeasureLastText <- text
        textMeasureLastFontPt <- fontPt
        textMeasureLastFontName <- fontName
        textMeasureLastBold <- bold
        textMeasureLastItalic <- italic
        textMeasureLastWidth <- width
    elif textMeasureLastSet
         && sameMeasureString textMeasureLastText text
         && sameMeasureStringOption textMeasureLastFontName fontName then
        textMeasureLastSet <- false

let private trimTextMeasureCache () =
    while (textMeasureCache.Count > textMeasureCacheMaxEntries
           || textMeasureCacheTotalTextChars > textMeasureCacheMaxTotalTextChars)
          && textMeasureCacheOrder.Count > 0 do
        let oldest = textMeasureCacheOrder.Dequeue()
        if textMeasureCache.Remove oldest then
            textMeasureCacheTotalTextChars <- max 0 (textMeasureCacheTotalTextChars - textMeasureCacheKeyChars oldest)

let private rememberTextMeasure key width =
    let entryChars = textMeasureCacheKeyChars key
    if key.Text.Length <= textMeasureCacheMaxSingleTextChars
       && entryChars <= textMeasureCacheMaxTotalTextChars then
        if not (textMeasureCache.ContainsKey key) then
            textMeasureCacheOrder.Enqueue key
            textMeasureCacheTotalTextChars <- textMeasureCacheTotalTextChars + entryChars
        textMeasureCache.[key] <- width
        trimTextMeasureCache ()

let private measureText
    (text: string) (fontPt: float)
    (fontName: string option) (bold: bool) (italic: bool) : float =
    match tryLastTextMeasure text fontPt fontName bold italic with
    | ValueSome width -> width
    | ValueNone ->
        let family = familyKey fontName
        let key = { Text = text; FontPt = fontPt; Family = family; Bold = bold; Italic = italic }
        let width =
            match textMeasureCache.TryGetValue key with
            | true, width -> width
            | false, _ ->
                let width = measureTextUncached text fontPt family bold italic
                rememberTextMeasure key width
                width
        rememberLastTextMeasure text fontPt fontName bold italic width
        width

let private viewportTextMeasurer = Some measureText

let private viewportOfCore scrollX scrollY width height : Viewport =
    // Match Excel's column/row header pixel sizes at 96 DPI / 100% zoom so
    // parity screenshots align cell-for-cell with our PNG renders. Measured
    // against artifacts/excel/pandas-test1-sheet1-headings.png: the column
    // header strip is 20 px tall (Calibri 11), the row header strip is 28 px
    // wide for two-digit row numbers, and each row in the default 15-pt
    // height renders at 21 source px.
    let headerHeight = 20.0
    let rowHeaderWidth = 28.0
    { ScrollX = scrollX
      ScrollY = scrollY
      Width = width
      Height = height
      HeaderHeight = headerHeight
      RowHeaderWidth = rowHeaderWidth
      Measurer = viewportTextMeasurer }

let viewportOf (model: SpreadsheetModel) width height : Viewport =
    viewportOfCore 0.0 0.0 width height

[<Struct>]
type private ViewportCommandsKey = {
    Model: SpreadsheetModel
    ScrollX: float
    ScrollY: float
    Width: float
    Height: float
}

type private ViewportCommandsCacheEntry = {
    Key: ViewportCommandsKey
    Commands: Cmd list
    CommandCount: int
}

let mutable private viewportCommandsCache: ViewportCommandsCacheEntry list = []
let private viewportCommandsKeyComparer =
    { new IEqualityComparer<ViewportCommandsKey> with
        member _.Equals(left, right) =
            obj.ReferenceEquals(left.Model, right.Model)
            && left.ScrollX = right.ScrollX
            && left.ScrollY = right.ScrollY
            && left.Width = right.Width
            && left.Height = right.Height
        member _.GetHashCode(value) =
            let mutable hashCode = RuntimeHelpers.GetHashCode(value.Model)
            hashCode <- (hashCode * 397) ^^^ value.ScrollX.GetHashCode()
            hashCode <- (hashCode * 397) ^^^ value.ScrollY.GetHashCode()
            hashCode <- (hashCode * 397) ^^^ value.Width.GetHashCode()
            (hashCode * 397) ^^^ value.Height.GetHashCode() }
let private viewportCommandsCacheByKey =
    Dictionary<ViewportCommandsKey, ViewportCommandsCacheEntry>(viewportCommandsKeyComparer)
let private viewportCommandsFastSlots = Array.zeroCreate<ViewportCommandsCacheEntry> 256
let private viewportCommandsCacheMaxEntries = 256
let private viewportCommandsCacheMaxCommands = 200000

let private viewportCommandsKey model scrollX scrollY width height = {
    Model = model
    ScrollX = scrollX
    ScrollY = scrollY
    Width = width
    Height = height
}

let private viewportCommandsSlot (key: ViewportCommandsKey) slotCount =
    viewportCommandsKeyComparer.GetHashCode key &&& (slotCount - 1)

let private viewportCommandsEntryMatches key (entry: ViewportCommandsCacheEntry) =
    viewportCommandsKeyComparer.Equals(key, entry.Key)

let private setViewportCommandsCache entries =
    viewportCommandsCache <- entries
    viewportCommandsCacheByKey.Clear()
    clearArray viewportCommandsFastSlots
    for entry in entries do
        viewportCommandsCacheByKey.[entry.Key] <- entry
        viewportCommandsFastSlots.[viewportCommandsSlot entry.Key viewportCommandsFastSlots.Length] <- entry

let private tryGetViewportCommands model scrollX scrollY width height =
    let key = viewportCommandsKey model scrollX scrollY width height
    let slot = viewportCommandsSlot key viewportCommandsFastSlots.Length
    let fast = viewportCommandsFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null)) && viewportCommandsEntryMatches key fast then
        ValueSome fast.Commands
    else
        let mutable cached = Unchecked.defaultof<ViewportCommandsCacheEntry>
        if viewportCommandsCacheByKey.TryGetValue(key, &cached) then
            viewportCommandsFastSlots.[slot] <- cached
            ValueSome cached.Commands
        else
            ValueNone

let private rememberViewportCommands model scrollX scrollY width height commands commandCount =
    if commandCount <= viewportCommandsCacheMaxCommands then
        let key = viewportCommandsKey model scrollX scrollY width height
        let entry = {
            Key = key
            Commands = commands
            CommandCount = commandCount
        }
        let rec removeMatching acc remaining =
            match remaining with
            | [] -> List.rev acc
            | candidate :: tail when viewportCommandsEntryMatches key candidate ->
                List.rev acc @ tail
            | candidate :: tail ->
                removeMatching (candidate :: acc) tail
        let trim entries =
            let rec loop keptCount keptCommands acc remaining =
                match remaining with
                | [] -> List.rev acc
                | candidate :: tail ->
                    let commandCount = max 1 candidate.CommandCount
                    if keptCount < viewportCommandsCacheMaxEntries
                       && keptCommands + commandCount <= viewportCommandsCacheMaxCommands then
                        loop (keptCount + 1) (keptCommands + commandCount) (candidate :: acc) tail
                    else
                        loop keptCount keptCommands acc tail
            loop 0 0 [] entries
        (entry :: removeMatching [] viewportCommandsCache)
        |> trim
        |> setViewportCommandsCache

[<Struct>]
type private CellCommandsKey = {
    Model: SpreadsheetModel
    Width: float
    Height: float
}

type private CellCommandsCacheEntry = {
    Key: CellCommandsKey
    Commands: Cmd list
    CommandCount: int
}

let mutable private cellCommandsCache: CellCommandsCacheEntry list = []
let private cellCommandsKeyComparer =
    { new IEqualityComparer<CellCommandsKey> with
        member _.Equals(left, right) =
            obj.ReferenceEquals(left.Model, right.Model)
            && left.Width = right.Width
            && left.Height = right.Height
        member _.GetHashCode(value) =
            let mutable hashCode = RuntimeHelpers.GetHashCode(value.Model)
            hashCode <- (hashCode * 397) ^^^ value.Width.GetHashCode()
            (hashCode * 397) ^^^ value.Height.GetHashCode() }
let private cellCommandsCacheByKey =
    Dictionary<CellCommandsKey, CellCommandsCacheEntry>(cellCommandsKeyComparer)
let private cellCommandsFastSlots = Array.zeroCreate<CellCommandsCacheEntry> 256
let private cellCommandsCacheMaxEntries = 256
let private cellCommandsCacheMaxCommands = 200000
let mutable private cellCommandsLastEntrySet = false
let mutable private cellCommandsLastEntry = Unchecked.defaultof<CellCommandsCacheEntry>
let mutable private cellCommandsPreviousEntrySet = false
let mutable private cellCommandsPreviousEntry = Unchecked.defaultof<CellCommandsCacheEntry>

let private cellCommandsKey model width height : CellCommandsKey = {
    Model = model
    Width = width
    Height = height
}

let private cellCommandsSlot (key: CellCommandsKey) slotCount =
    cellCommandsKeyComparer.GetHashCode key &&& (slotCount - 1)

let private cellCommandsEntryMatches key (entry: CellCommandsCacheEntry) =
    cellCommandsKeyComparer.Equals(key, entry.Key)

let private cellCommandsEntryMatchesRaw model width height (entry: CellCommandsCacheEntry) =
    obj.ReferenceEquals(entry.Key.Model, model)
    && entry.Key.Width = width
    && entry.Key.Height = height

let private rememberCellCommandsLastEntry entry =
    if cellCommandsLastEntrySet && obj.ReferenceEquals(cellCommandsLastEntry, entry) then
        ()
    elif cellCommandsPreviousEntrySet && obj.ReferenceEquals(cellCommandsPreviousEntry, entry) then
        let current = cellCommandsLastEntry
        cellCommandsLastEntry <- entry
        cellCommandsPreviousEntry <- current
        cellCommandsPreviousEntrySet <- cellCommandsLastEntrySet
        cellCommandsLastEntrySet <- true
    else
        if cellCommandsLastEntrySet then
            cellCommandsPreviousEntry <- cellCommandsLastEntry
            cellCommandsPreviousEntrySet <- true
        cellCommandsLastEntry <- entry
        cellCommandsLastEntrySet <- true

let private clearCellCommandsLastEntries () =
    cellCommandsLastEntrySet <- false
    cellCommandsLastEntry <- Unchecked.defaultof<CellCommandsCacheEntry>
    cellCommandsPreviousEntrySet <- false
    cellCommandsPreviousEntry <- Unchecked.defaultof<CellCommandsCacheEntry>

let private setCellCommandsCache entries =
    cellCommandsCache <- entries
    cellCommandsCacheByKey.Clear()
    clearArray cellCommandsFastSlots
    clearCellCommandsLastEntries ()
    for entry in entries do
        cellCommandsCacheByKey.[entry.Key] <- entry
        cellCommandsFastSlots.[cellCommandsSlot entry.Key cellCommandsFastSlots.Length] <- entry
        if not cellCommandsLastEntrySet
           || not cellCommandsPreviousEntrySet then
            rememberCellCommandsLastEntry entry

let private tryGetCellCommands model width height =
    if cellCommandsLastEntrySet
       && cellCommandsEntryMatchesRaw model width height cellCommandsLastEntry then
        ValueSome cellCommandsLastEntry.Commands
    elif cellCommandsPreviousEntrySet
         && cellCommandsEntryMatchesRaw model width height cellCommandsPreviousEntry then
        let entry = cellCommandsPreviousEntry
        rememberCellCommandsLastEntry entry
        ValueSome entry.Commands
    else
        let key = cellCommandsKey model width height
        let slot = cellCommandsSlot key cellCommandsFastSlots.Length
        let fast = cellCommandsFastSlots.[slot]
        if not (obj.ReferenceEquals(fast, null)) && cellCommandsEntryMatches key fast then
            rememberCellCommandsLastEntry fast
            ValueSome fast.Commands
        else
            let mutable cached = Unchecked.defaultof<CellCommandsCacheEntry>
            if cellCommandsCacheByKey.TryGetValue(key, &cached) then
                cellCommandsFastSlots.[slot] <- cached
                rememberCellCommandsLastEntry cached
                ValueSome cached.Commands
            else
                ValueNone

let private rememberCellCommands model width height commands commandCount =
    if commandCount <= cellCommandsCacheMaxCommands then
        let key = cellCommandsKey model width height
        let entry = {
            Key = key
            Commands = commands
            CommandCount = commandCount
        }
        let rec removeMatching acc remaining =
            match remaining with
            | [] -> List.rev acc
            | candidate :: tail when cellCommandsEntryMatches key candidate ->
                List.rev acc @ tail
            | candidate :: tail ->
                removeMatching (candidate :: acc) tail
        let trim entries =
            let rec loop keptCount keptCommands acc remaining =
                match remaining with
                | [] -> List.rev acc
                | candidate :: tail ->
                    let commandCount = max 1 candidate.CommandCount
                    if keptCount < cellCommandsCacheMaxEntries
                       && keptCommands + commandCount <= cellCommandsCacheMaxCommands then
                        loop (keptCount + 1) (keptCommands + commandCount) (candidate :: acc) tail
                    else
                        loop keptCount keptCommands acc tail
            loop 0 0 [] entries
        (entry :: removeMatching [] cellCommandsCache)
        |> trim
        |> setCellCommandsCache

let private cellCommands (model: SpreadsheetModel) width height =
    match tryGetCellCommands model width height with
    | ValueSome commands -> commands
    | ValueNone ->
        let viewport = viewportOf model width height
        let commands =
            match tryBuildSparseDefaultPaintPartsValue model viewport with
            | ValueSome parts -> rendererCommandsForSparseDefaultParts parts
            | ValueNone ->
                buildPaintItems model viewport
                |> rendererCommandsForPaintItems
        rememberCellCommands model width height commands commands.Length
        commands

let private cellCommandsInViewport (model: SpreadsheetModel) scrollX scrollY viewportWidth viewportHeight =
    match tryGetViewportCommands model scrollX scrollY viewportWidth viewportHeight with
    | ValueSome commands -> commands
    | ValueNone ->
        let viewport = viewportOfCore scrollX scrollY viewportWidth viewportHeight
        let commands, cacheable =
            match tryBuildSparseDefaultPaintPartsValue model viewport with
            | ValueSome parts -> rendererCommandsForSparseDefaultParts parts, true
            | ValueNone ->
                buildPaintItems model viewport
                |> rendererCommandsForPaintItems,
                false
        let framed = commandsWithViewportFrame scrollX scrollY viewportWidth viewportHeight commands
        if cacheable then
            rememberViewportCommands model scrollX scrollY viewportWidth viewportHeight framed framed.Length
        framed

let widget width height (model: SpreadsheetModel) =
    canvas width height (fun w h -> cellCommands model w h)

let viewportWidget contentWidth contentHeight viewportWidth viewportHeight scrollX scrollY (model: SpreadsheetModel) =
    canvas contentWidth contentHeight (fun _ _ ->
        cellCommandsInViewport model scrollX scrollY viewportWidth viewportHeight)

/// Render the sheet with extra paint items (e.g. charts) overlaid on top, in
/// the same viewport pixel space as the cells.
let widgetWith width height (model: SpreadsheetModel) (extra: PaintItem list) =
    canvas width height (fun w h ->
        let commands = System.Collections.Generic.List<Cmd>()
        commands.AddRange(cellCommands model w h)
        for item in extra do
            appendPaintItem commands item
        resizeArrayToList commands)

/// Render a standalone vector PaintItem scene (no model) on a white
/// background — used to prove the VecPath/VecGroup primitives end-to-end.
let renderScene width height (items: PaintItem list) =
    canvas width height (fun w h ->
        let commands = System.Collections.Generic.List<Cmd>()
        commands.Add(Cmd.FillRect(Renderer.Types.rect 0.0 0.0 w h, parseHexColor "#FFFFFF"))
        for item in items do
            appendPaintItem commands item
        resizeArrayToList commands)
