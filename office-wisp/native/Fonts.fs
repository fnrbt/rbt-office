module Rbt.OfficeNative.Fonts

let private resizeArrayToList (items: System.Collections.Generic.List<'T>) =
    let mutable result = []
    for i = items.Count - 1 downto 0 do
        result <- items.[i] :: result
    result

let private defaultFont = lazy (
    [ "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"
      "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf"
      "/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf"
      "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
      System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Fonts), "arial.ttf") ]
    |> List.tryFind System.IO.File.Exists
    |> Option.map Rbt.Renderer.Text.loadFontFile
    |> Option.defaultWith (fun () -> failwith "No default font found; set RB_WORD_FONT_DIR to a directory containing calibri.ttf"))

let private fallbackDefaultFont () = defaultFont.Value

/// Cache of font variants keyed by (familyKey, bold, italic). familyKey is the
/// normalized lower-case font name we use to pick a TTF set ("calibri",
/// "arial", "cambria", "times"); unknown families fall back to "calibri".
let private fontCache = System.Collections.Generic.Dictionary<string * bool * bool, Rbt.Renderer.Text.FontData option>()

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
        elif startsWithAtIgnoreCase n start length "calibri light" then "calibri-light"
        elif startsWithAtIgnoreCase n start length "calibri" then "calibri"
        elif startsWithAtIgnoreCase n start length "cambria" then "cambria"
        elif startsWithAtIgnoreCase n start length "times" then "times"
        elif startsWithAtIgnoreCase n start length "verdana" then "verdana"
        elif startsWithAtIgnoreCase n start length "garamond" then "garamond"
        elif startsWithAtIgnoreCase n start length "segoe" then "segoeui"
        elif startsWithAtIgnoreCase n start length "helvetica" then "arial"
        elif startsWithAtIgnoreCase n start length "georgia" then "times"
        elif startsWithAtIgnoreCase n start length "consolas" then "consolas"
        elif startsWithAtIgnoreCase n start length "courier" then "courier"
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
        if bold && italic then ["cambriaz.ttf"; "cambriab.ttf"; "cambriai.ttf"; "cambria.ttc"]
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
    | "consolas" ->
        if bold && italic then ["consolaz.ttf"; "consolab.ttf"; "consola.ttf"]
        elif bold then ["consolab.ttf"; "consola.ttf"]
        elif italic then ["consolai.ttf"; "consola.ttf"]
        else ["consola.ttf"]
    | "courier" ->
        if bold && italic then ["courbi.ttf"; "courbd.ttf"; "cour.ttf"]
        elif bold then ["courbd.ttf"; "cour.ttf"]
        elif italic then ["couri.ttf"; "cour.ttf"]
        else ["cour.ttf"]
    | "calibri-light" ->
        if bold && italic then ["calibriz.ttf"; "calibrili.ttf"; "calibril.ttf"]
        elif bold then ["calibrib.ttf"; "calibril.ttf"]
        elif italic then ["calibrili.ttf"; "calibrii.ttf"; "calibril.ttf"]
        else ["calibril.ttf"; "calibri.ttf"]
    | _ ->
        if bold && italic then ["calibriz.ttf"; "calibrib.ttf"; "calibri.ttf"]
        elif bold then ["calibrib.ttf"; "calibri.ttf"]
        elif italic then ["calibrii.ttf"; "calibri.ttf"]
        else ["calibri.ttf"]

let private loadVariantUnlocked family bold italic =
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
                        try Some (Rbt.Renderer.Text.loadFontFile path) with _ -> None
                    else None))
        let loaded =
            match loaded with
            | Some _ -> loaded
            | None ->
                let baseName =
                    match family with
                    | "times" | "cambria" | "garamond" -> "LiberationSerif"
                    | "consolas" | "courier" -> "LiberationMono"
                    | _ -> "LiberationSans"
                let style =
                    if bold && italic then "BoldItalic"
                    elif bold then "Bold"
                    elif italic then "Italic"
                    else "Regular"
                [ "/usr/share/fonts/truetype/liberation"
                  "/usr/share/fonts/truetype/liberation2" ]
                |> List.tryPick (fun dir ->
                    let path = System.IO.Path.Combine(dir, baseName + "-" + style + ".ttf")
                    if System.IO.File.Exists path then
                        try Some (Rbt.Renderer.Text.loadFontFile path) with _ -> None
                    else None)
        fontCache.[key] <- loaded
        loaded
let private loadVariant family bold italic =
    lock fontCache (fun () -> loadVariantUnlocked family bold italic)


let resolve (fontName: string option) (bold: bool) (italic: bool) =
    let family = familyKey fontName
    let font =
        match loadVariant family bold italic with
        | Some f -> f
        | None ->
            // Family lookup failed — try Calibri before falling back to the
            // built-in default so we don't lose bold/italic for Arial/Times etc.
            loadVariant "calibri" bold italic
            |> Option.defaultWith fallbackDefaultFont
    font

/// Load a CJK fallback by file name. Mirrors loadVariant's caching and
/// fontDirs lookup. Returns None when the file is absent (older Windows
/// images don't ship every face).
let private cjkFontCache = System.Collections.Generic.Dictionary<string, Rbt.Renderer.Text.FontData option>()
let private loadCjkFileUnlocked (fileName: string) : Rbt.Renderer.Text.FontData option =
    match cjkFontCache.TryGetValue fileName with
    | true, cached -> cached
    | _ ->
        let loaded =
            fontDirs ()
            |> List.tryPick (fun dir ->
                let path = System.IO.Path.Combine(dir, fileName)
                if System.IO.File.Exists path then
                    try Some (Rbt.Renderer.Text.loadFontFile path) with _ -> None
                else None)
        cjkFontCache.[fileName] <- loaded
        loaded
let private loadCjkFile fileName =
    lock cjkFontCache (fun () -> loadCjkFileUnlocked fileName)


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

/// The installed fallback faces, in the same order used by native text painting.
/// PDF adapters reuse their loaded font bytes instead of maintaining another lookup.
let fallbackFaces bold =
    cjkFallbackChain bold
    |> List.choose (fun name -> loadCjkFile name |> Option.map (fun font -> name, font))

/// For a codepoint not covered by the primary font, find the first font in
/// the fallback chain whose cmap maps it. Returns None when no chain font
/// is installed or none covers the codepoint (will render .notdef = box).
let private fallbackFontFor (codepoint: int) (bold: bool) : Rbt.Renderer.Text.FontData option =
    cjkFallbackChain bold
    |> List.tryPick (fun name ->
        loadCjkFile name
        |> Option.bind (fun f -> if Rbt.Renderer.Text.hasGlyph f codepoint then Some f else None))

let segmentByFontCoverage
    (primary: Rbt.Renderer.Text.FontData) (bold: bool) (text: string)
    : System.Collections.Generic.List<Rbt.Renderer.Text.FontData * string> =
    let fontFor (cp: int) : Rbt.Renderer.Text.FontData =
        if Rbt.Renderer.Text.hasGlyph primary cp then primary
        else
            match fallbackFontFor cp bold with
            | Some fb -> fb
            | None -> primary
    let result = System.Collections.Generic.List<Rbt.Renderer.Text.FontData * string>()
    let buf = System.Text.StringBuilder()
    let mutable currentFont : Rbt.Renderer.Text.FontData voption = ValueNone
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

/// Measure with the same primary and fallback faces used for drawing; size is pixels.
let measure (text: string) (sizePx: float) (family: string option) (bold: bool) (italic: bool) =
    let font = resolve family bold italic
    let mutable width = 0.0
    for face, run in segmentByFontCoverage font bold text do
        width <- width + Rbt.Renderer.Text.measureWidth face sizePx run
    width
