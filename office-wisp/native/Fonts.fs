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

/// Map an XLSX font `name val="..."` to a normalized family key: None for a
/// family we have no TTF set for (see familyKey).
let private knownFamilyKey (name: string option) =
    match name with
    | None -> Some "calibri"
    | Some n ->
        let start, length = trimBounds n
        let known key = Some key
        if length = 0 then known "calibri"
        elif startsWithAtIgnoreCase n start length "arial" then known "arial"
        elif startsWithAtIgnoreCase n start length "calibri light" then known "calibri-light"
        elif startsWithAtIgnoreCase n start length "calibri" then known "calibri"
        elif startsWithAtIgnoreCase n start length "cambria" then known "cambria"
        elif startsWithAtIgnoreCase n start length "times" then known "times"
        elif startsWithAtIgnoreCase n start length "verdana" then known "verdana"
        elif startsWithAtIgnoreCase n start length "garamond" then known "garamond"
        elif startsWithAtIgnoreCase n start length "segoe" then known "segoeui"
        elif startsWithAtIgnoreCase n start length "helvetica" then known "arial"
        elif startsWithAtIgnoreCase n start length "georgia" then known "times"
        elif startsWithAtIgnoreCase n start length "consolas" then known "consolas"
        elif startsWithAtIgnoreCase n start length "courier" then known "courier"
        // Localized "Body Text Font" names from non-English Excel installs.
        // OOXML stores the localized display name verbatim; Windows resolves
        // it through the OS locale. For our purposes Arial is the closest
        // Excel substitute on a system that doesn't have the localized
        // font (verified against pandas-testdateoverflow's Polish
        // "Czcionka tekstu podstawowego" — column widths and row heights
        // match Excel's reported values with Arial metrics, not Calibri's).
        elif equalsAtIgnoreCase n start length "czcionka tekstu podstawowego" then known "arial"     // Polish
        elif equalsAtIgnoreCase n start length "schrift des textkörpers" then known "arial"          // German
        elif equalsAtIgnoreCase n start length "police du corps" then known "arial"                  // French
        elif equalsAtIgnoreCase n start length "carattere del corpo" then known "arial"              // Italian
        elif equalsAtIgnoreCase n start length "fuente del cuerpo" then known "arial"                // Spanish
        elif equalsAtIgnoreCase n start length "本文のフォント" then known "arial"                    // Japanese
        elif equalsAtIgnoreCase n start length "正文字体" then known "arial"                          // Chinese (Simplified)
        elif equalsAtIgnoreCase n start length "본문 글꼴" then known "arial"                         // Korean
        else None

/// Anything we don't have a TTF for falls back to "calibri" — the renderer
/// keeps working even when we encounter Asian/script fonts we can't
/// substitute properly.
let private familyKey (name: string option) =
    knownFamilyKey name |> Option.defaultValue "calibri"

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

/// Whether a family is one this machine sets text in its own font of (a family of the known sets
/// whose files - or their Liberation stand-ins - are here); text in another family falls back.
let hasFamily (fontName: string) =
    match knownFamilyKey (Some fontName) with
    | Some family -> (loadVariant family false false).IsSome
    | None -> false

// The fonts of a document's embedded faces, by the face's bytes (read once).
let private embeddedFonts = System.Runtime.CompilerServices.ConditionalWeakTable<byte[], Rbt.Renderer.Text.FontData option>()

let private embeddedFont (data: byte[]) =
    lock embeddedFonts (fun () ->
        match embeddedFonts.TryGetValue data with
        | true, font -> font
        | _ ->
            let font = try Some (Rbt.Renderer.Text.loadFont data) with _ -> None
            embeddedFonts.Add(data, font)
            font)

/// As `resolve`, with the faces a document embeds ((family, bold, italic, font file) - the document's
/// EmbeddedFonts): text in a family this machine has no font of (hasFamily) is set in an embedded face
/// of it - the face of its style, else another of the family's (as Word: fixtures embedded-fonts) -
/// and text in any other family as `resolve` sets it.
let resolveWith (embedded: (string * bool * bool * byte[]) list) (fontName: string option) (bold: bool) (italic: bool) =
    let norm (name: string) = name.Trim().ToLowerInvariant()
    match fontName with
    | Some name when not embedded.IsEmpty && not (hasFamily name) ->
        let faces = embedded |> List.filter (fun (family, _, _, _) -> norm family = norm name)
        [ bold, italic; bold, false; false, italic; false, false; not bold, italic; bold, not italic; not bold, not italic ]
        |> List.tryPick (fun (b, i) ->
            faces |> List.tryPick (fun (_, faceBold, faceItalic, data) -> if faceBold = b && faceItalic = i then embeddedFont data else None))
        |> Option.defaultWith (fun () -> resolve fontName bold italic)
    | _ -> resolve fontName bold italic

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

/// The runs of `text` and the fonts that set them, with a document's embedded faces (see resolveWith):
/// a character the embedded face of the family has no glyph of is set in Calibri, as Word sets it -
/// that character alone when the family embeds one face, the rest of the text from it when it embeds
/// more (fixtures embedded-fonts subset*); any other text as segmentByFontCoverage sets it.
let segmentWith (embedded: (string * bool * bool * byte[]) list) (family: string option) (bold: bool) (italic: bool) (text: string) =
    let norm (name: string) = name.Trim().ToLowerInvariant()
    let primary = resolveWith embedded family bold italic
    let faces =
        match family with
        | Some name when not embedded.IsEmpty && not (hasFamily name) -> embedded |> List.filter (fun (f, _, _, _) -> norm f = norm name) |> List.length
        | _ -> 0
    if faces = 0 then segmentByFontCoverage primary bold text
    else
        let calibri = resolve (Some "Calibri") bold italic
        let result = System.Collections.Generic.List<Rbt.Renderer.Text.FontData * string>()
        let add (font: Rbt.Renderer.Text.FontData) (part: string) =
            if part <> "" then
                if result.Count > 0 && obj.ReferenceEquals(fst result.[result.Count - 1], font) then
                    result.[result.Count - 1] <- (font, snd result.[result.Count - 1] + part)
                else result.Add((font, part))
        let mutable i = 0
        let mutable rest = false
        while i < text.Length do
            let width = if System.Char.IsHighSurrogate text.[i] && i + 1 < text.Length then 2 else 1
            let part = text.Substring(i, width)
            let cp = System.Char.ConvertToUtf32(part, 0)
            if rest || not (Rbt.Renderer.Text.hasGlyph primary cp) then
                add calibri part
                if faces > 1 then rest <- true
            else add primary part
            i <- i + width
        result

/// As `measure`, with a document's embedded faces (see resolveWith, segmentWith).
let measureWith (embedded: (string * bool * bool * byte[]) list) (text: string) (sizePx: float) (family: string option) (bold: bool) (italic: bool) =
    let mutable width = 0.0
    for face, run in segmentWith embedded family bold italic text do
        width <- width + Rbt.Renderer.Text.measureWidth face sizePx run
    width

/// Measure with the same primary and fallback faces used for drawing; size is pixels.
let measure (text: string) (sizePx: float) (family: string option) (bold: bool) (italic: bool) =
    let font = resolve family bold italic
    let mutable width = 0.0
    for face, run in segmentByFontCoverage font bold text do
        width <- width + Rbt.Renderer.Text.measureWidth face sizePx run
    width
