module Rbt.Word.ShaperText

open Rbt.Word.TextBuffer

type FontFace = {
    Family: string
    Bold: bool
    Italic: bool
    Data: byte[]
}

type FontFamilyData = {
    Family: string
    Regular: byte[]
    Bold: byte[] option
    Italic: byte[] option
    BoldItalic: byte[] option
}

type FontKey = string * bool * bool

type FontMetrics = {
    UnitsPerEm: int
    Ascender: float
    Descender: float
    LineGap: float
    LineHeight: float
}

type GlyphPathCommand =
    | PathMoveTo of x: float * y: float
    | PathLineTo of x: float * y: float
    | PathQuadTo of cx: float * cy: float * x: float * y: float
    | PathCubicTo of cx1: float * cy1: float * cx2: float * cy2: float * x: float * y: float
    | PathClose

/// A font file as it is used: read once for shaping, and - when a glyph of
/// it is first drawn - once for its outlines.
type private ParsedFace = {
    Data: byte[]
    Font: Rbt.Text.Engine.TextFont
    mutable Outline: Rbt.OpenType.Font.Font option
}

type private FontEntry = {
    Family: string
    Bold: bool
    Italic: bool
    Face: ParsedFace
}

type FontSet = private {
    DefaultFamily: string
    Fonts: Map<FontKey, FontEntry>
}

let private normalizeFamily (value: string) =
    value.Trim().ToLowerInvariant()

let private fontKey family bold italic = normalizeFamily family, bold, italic

// A font file is megabytes and reading it takes tens of milliseconds, most of
// it for the outlines (Calibri: 2 ms for shaping, 15 - 60 ms for the outlines).
// A host builds a font set of the faces it has loaded, and another one when a
// face arrives; it gives a file several times, once per name the family goes
// by. So: a file is read once for as long as its bytes are around - the bytes
// are known by the array that holds them, not by their content - and its
// outlines are read when the first glyph of it is drawn (or ahead of that, see
// prepareOutline), not when a set is made.
let mutable private shapingParses = 0
let mutable private outlineParses = 0

#if FABLE_COMPILER
[<Fable.Core.Emit("new WeakMap()")>]
let private newFaceCache () : obj = Fable.Core.Util.jsNative
[<Fable.Core.Emit("$0.get($1)")>]
let private faceCacheGet (_cache: obj) (_data: byte[]) : ParsedFace = Fable.Core.Util.jsNative
[<Fable.Core.Emit("$0.set($1, $2)")>]
let private faceCacheSet (_cache: obj) (_data: byte[]) (_face: ParsedFace) : unit = Fable.Core.Util.jsNative
let private faceCache = newFaceCache ()
let private tryCachedFace (data: byte[]) : ParsedFace option =
    let known = faceCacheGet faceCache data
    if isNull (box known) then None else Some known
let private cacheFace (data: byte[]) (face: ParsedFace) = faceCacheSet faceCache data face
#else
let private faceCache = System.Runtime.CompilerServices.ConditionalWeakTable<byte[], ParsedFace>()
let private tryCachedFace (data: byte[]) : ParsedFace option =
    match faceCache.TryGetValue data with
    | true, known -> Some known
    | false, _ -> None
let private cacheFace (data: byte[]) (face: ParsedFace) =
    lock faceCache (fun () ->
        match faceCache.TryGetValue data with
        | true, _ -> ()
        | false, _ -> faceCache.Add (data, face))
#endif

/// The face of a font file: the one made of these bytes before, or a new one.
let private parsedFace (data: byte[]) : ParsedFace =
    match tryCachedFace data with
    | Some known -> known
    | None ->
        shapingParses <- shapingParses + 1
        let face = { Data = data; Font = Rbt.Text.Engine.loadFont data; Outline = None }
        cacheFace data face
        face

let private outlineFont (entry: FontEntry) : Rbt.OpenType.Font.Font =
    match entry.Face.Outline with
    | Some font -> font
    | None ->
        outlineParses <- outlineParses + 1
        let font = Rbt.OpenType.Font.load entry.Face.Data
        entry.Face.Outline <- Some font
        font

// ---- shaping as Word sets text ------------------------------------------------
//
// Word sets a run without kerning and without ligatures unless the run asks for them (w:kern at
// its size and up, w14:ligatures): Word 2019's PDFs of tests/word/fixtures/kerning give every word
// the sum of its glyphs' advances - "AVAVAVAVAV" 63.04pt of Calibri 11pt, kerned 58.73pt only under
// w:kern - and "office" its own letters, fused only under w14:ligatures. The engine's shaping
// applies the fonts' kerning and ligatures, and a line laid out with it held text Word's does not
// (docx4j-headers-footers' "Reference", 45.68pt kerned and 46.24pt as Word measures it, in a 46.25pt
// cell; test-long-document's lines a word longer). So text of the scripts that need neither (Latin,
// Greek, Cyrillic and the punctuation and symbols round them) is shaped here with the features
// that compose and place marks only; other text as the engine shapes it.
let private wordFeatures =
    Some (Set.ofList [ Rbt.Shaper.OpenType.Types.Features.ccmp; Rbt.Shaper.OpenType.Types.Features.locl
                       Rbt.Shaper.OpenType.Types.Features.rlig; Rbt.Shaper.OpenType.Types.Features.mark
                       Rbt.Shaper.OpenType.Types.Features.mkmk ])
let private wordConfig : Rbt.Shaper.Shaper.ShaperConfig = { Features = wordFeatures; Language = None; Script = None }

/// Text that left-to-right shaping with no features but composition and marks sets as Word does:
/// below U+0590 (Latin, IPA, combining marks, Greek, Cyrillic, Armenian) or general punctuation,
/// letterlike symbols, arrows and the like (U+2000 - U+2BFF).
let private isSimpleText (text: string) =
    let mutable simple = true
    let mutable i = 0
    while simple && i < text.Length do
        let code = int text.[i]
        simple <- code < 0x0590 || (code >= 0x2000 && code < 0x2C00)
        i <- i + 1
    simple

let private wordShapeCache = System.Collections.Generic.Dictionary<struct (int * string), Rbt.Text.Engine.ShapedText>()
let private wordShapeCacheLimit = 4096
let private wordCacheLock = obj ()
// (The shaper's own caches are not safe to share between threads: shaping is one at a time.)
let private wordShapeLock = obj ()

/// Each glyph's advance spread over the characters it stands for (as the engine does).
let private prefixAdvances (text: string) (glyphs: Rbt.Text.Engine.PositionedGlyph[]) =
    let prefixes = Array.zeroCreate (text.Length + 1)
    let mutable total = 0.0
    for index in 0 .. glyphs.Length - 1 do
        let glyph = glyphs.[index]
        let nextCluster = if index + 1 < glyphs.Length then glyphs.[index + 1].Cluster else text.Length
        let startCluster = max 0 (min text.Length glyph.Cluster)
        let endCluster = max startCluster (min text.Length nextCluster)
        let span = endCluster - startCluster
        if span <= 1 then
            total <- total + glyph.XAdv
            prefixes.[endCluster] <- max prefixes.[endCluster] total
        else
            let perStep = glyph.XAdv / float span
            for step in 1 .. span do
                let pos = startCluster + step
                prefixes.[pos] <- max prefixes.[pos] (total + perStep * float step)
            total <- total + glyph.XAdv
    let mutable current = 0.0
    for i in 0 .. prefixes.Length - 1 do
        current <- max current prefixes.[i]
        prefixes.[i] <- current
    prefixes, total

let private shapeSimple (font: Rbt.Text.Engine.TextFont) (fontSize: float) (text: string) : Rbt.Text.Engine.ShapedText =
    let scale = fontSize / float font.ShaperFont.UnitsPerEm
    let codepoints = Array.init text.Length (fun i -> int text.[i])
    let shaped = lock wordShapeLock (fun () -> Rbt.Shaper.Shaper.shapeCodepoints font.ShaperFont wordConfig codepoints Rbt.Shaper.Shaper.LeftToRight)
    let mutable x = 0.0
    let glyphs =
        [| for g in shaped.Glyphs do
            let adv = float g.XAdvance * scale
            yield ({ GlyphId = g.GlyphId; X = x + float g.XOffset * scale; Y = float g.YOffset * scale; XAdv = adv; Cluster = g.Cluster }: Rbt.Text.Engine.PositionedGlyph)
            x <- x + adv |]
    let prefixes, total = prefixAdvances text glyphs
    { Text = text; FontSize = fontSize
      Runs = [ { FontSize = fontSize; Glyphs = glyphs; Direction = Rbt.Shaper.Shaper.LeftToRight } ]
      PrefixAdvances = prefixes; TotalAdvance = total }

let private scaled (fontSize: float) (shaped: Rbt.Text.Engine.ShapedText) : Rbt.Text.Engine.ShapedText =
    if shaped.FontSize = fontSize then shaped
    else
        let ratio = fontSize / shaped.FontSize
        { shaped with
            FontSize = fontSize
            Runs = shaped.Runs |> List.map (fun run ->
                { run with
                    FontSize = fontSize
                    Glyphs = run.Glyphs |> Array.map (fun g -> { g with X = g.X * ratio; Y = g.Y * ratio; XAdv = g.XAdv * ratio }) })
            PrefixAdvances = shaped.PrefixAdvances |> Array.map ((*) ratio)
            TotalAdvance = shaped.TotalAdvance * ratio }

/// `text` shaped as Word sets it (see wordFeatures).
let private shapeAsWord (font: Rbt.Text.Engine.TextFont) (fontSize: float) (text: string) : Rbt.Text.Engine.ShapedText =
    if System.String.IsNullOrEmpty text || not (isSimpleText text) then
        // (Under the same lock: the engine shapes with the same shaper.)
        lock wordShapeLock (fun () -> Rbt.Text.Engine.shapeText font fontSize text)
    else
        let key = struct (LanguagePrimitives.PhysicalHash font.ShaperFont, text)
        let cached = lock wordCacheLock (fun () -> match wordShapeCache.TryGetValue key with | true, c -> Some c | _ -> None)
        match cached with
        | Some shaped -> scaled fontSize shaped
        | None ->
            let shaped = shapeSimple font fontSize text
            lock wordCacheLock (fun () ->
                if wordShapeCache.Count >= wordShapeCacheLimit then wordShapeCache.Clear()
                wordShapeCache.[key] <- shaped)
            shaped

/// How many font files were read since the module was loaded: for shaping
/// (once per distinct array of bytes given in a FontFace) and for their
/// outlines (once per file a glyph of which was drawn).
let fontSetStats () = {| ShapingParses = shapingParses; OutlineParses = outlineParses |}

/// A font set with more faces: `fontSet`'s and `faces` (a face of a family,
/// weight and slant the set has is replaced). The set given is not changed -
/// what was measured and shaped with it stays valid - and the new one shares
/// its fonts: only files that no set was made of before are read. A host
/// whose font set gains a face builds its measurers and its painter on the
/// new set and tells the editor which faces came (Word's
/// setTextMeasurersForFaces), so that only text in them is laid out again.
let addFaces (fontSet: FontSet) (faces: FontFace seq) : FontSet =
    let mutable fonts = fontSet.Fonts
    for face in faces do
        fonts <-
            fonts
            |> Map.add (fontKey face.Family face.Bold face.Italic)
                       { Family = face.Family; Bold = face.Bold; Italic = face.Italic; Face = parsedFace face.Data }
    { fontSet with Fonts = fonts }

/// A font set of `faces`. Text in a family, weight or slant it has no face of
/// is set in the nearest face of the family, then of `defaultFamily`. The same
/// array of bytes given again (under another family name, in another set) is
/// the same font: it is read once.
let createFontSet defaultFamily (faces: FontFace seq) =
    addFaces { DefaultFamily = defaultFamily; Fonts = Map.empty } faces

/// Read the outlines of one file of the set that has not been drawn from yet
/// (they are read anyway when its first glyph is drawn: this moves the tens
/// of milliseconds that takes to a moment the host picks, an idle one). True
/// when there are more such files: call it again.
let prepareOutline (fontSet: FontSet) : bool =
    let pending = fontSet.Fonts |> Map.toSeq |> Seq.map snd |> Seq.filter (fun entry -> entry.Face.Outline.IsNone) |> Seq.toList
    match pending with
    | [] -> false
    | entry :: _ ->
        outlineFont entry |> ignore
        pending |> List.exists (fun other -> other.Face.Outline.IsNone)

let fontFacesFromFamilies (families: FontFamilyData seq) =
    families
    |> Seq.collect (fun (family: FontFamilyData) ->
        seq {
            yield ({ Family = family.Family; Bold = false; Italic = false; Data = family.Regular }: FontFace)
            match family.Bold with
            | Some data -> yield ({ Family = family.Family; Bold = true; Italic = false; Data = data }: FontFace)
            | None -> ()
            match family.Italic with
            | Some data -> yield ({ Family = family.Family; Bold = false; Italic = true; Data = data }: FontFace)
            | None -> ()
            match family.BoldItalic with
            | Some data -> yield ({ Family = family.Family; Bold = true; Italic = true; Data = data }: FontFace)
            | None -> ()
        })

let createFontSetFromFamilies defaultFamily families =
    fontFacesFromFamilies families |> createFontSet defaultFamily

let private tryFindFont (fontSet: FontSet) family bold italic =
    let requested = family |> Option.defaultValue fontSet.DefaultFamily
    let candidates = [
        fontKey requested bold italic
        fontKey requested bold false
        fontKey requested false italic
        fontKey requested false false
        fontKey fontSet.DefaultFamily bold italic
        fontKey fontSet.DefaultFamily bold false
        fontKey fontSet.DefaultFamily false italic
        fontKey fontSet.DefaultFamily false false
    ]
    candidates |> List.tryPick (fun key -> fontSet.Fonts |> Map.tryFind key)

let measureText (fontSet: FontSet) (text: string) (style: TextStyle) =
    if System.String.IsNullOrEmpty text then 0.0
    else
        match tryFindFont fontSet style.FontFamily style.Bold style.Italic with
        | None -> 0.0
        | Some entry -> (shapeAsWord entry.Face.Font style.FontSize text).TotalAdvance + TextBuffer.characterSpacingAdvance text style

let tryMeasureText (fontSet: FontSet) (text: string) (style: TextStyle) =
    if System.String.IsNullOrEmpty text then Some 0.0
    else
        match tryFindFont fontSet style.FontFamily style.Bold style.Italic with
        | None -> None
        | Some entry -> Some ((shapeAsWord entry.Face.Font style.FontSize text).TotalAdvance + TextBuffer.characterSpacingAdvance text style)

let tryLineMetrics (fontSet: FontSet) (style: TextStyle) =
    match tryFindFont fontSet style.FontFamily style.Bold style.Italic with
    | None -> None
    | Some entry ->
        let metrics = Rbt.Text.Engine.metrics entry.Face.Font
        let scale = style.FontSize / float metrics.UnitsPerEm
        let ascender = float metrics.Ascender * scale
        let descender = float metrics.Descender * scale
        let lineGap = float metrics.LineGap * scale
        Some {
            UnitsPerEm = metrics.UnitsPerEm
            Ascender = ascender
            Descender = descender
            LineGap = lineGap
            LineHeight = ascender - descender + lineGap
        }

let createMeasurer fontSet : Rbt.Word.Layout.TextMeasurer =
    fun text style -> measureText fontSet text style

let shapeText (fontSet: FontSet) (text: string) (style: TextStyle) =
    match tryFindFont fontSet style.FontFamily style.Bold style.Italic with
    | None -> None
    | Some entry -> Some (entry.Face.Font, shapeAsWord entry.Face.Font style.FontSize text)

let unitsPerEm (fontSet: FontSet) (style: TextStyle) =
    match tryFindFont fontSet style.FontFamily style.Bold style.Italic with
    | None -> 1000
    | Some entry -> (Rbt.Text.Engine.metrics entry.Face.Font).UnitsPerEm

let glyphPath (fontSet: FontSet) (style: TextStyle) glyphId =
    match tryFindFont fontSet style.FontFamily style.Bold style.Italic with
    | None -> [||]
    | Some entry ->
        Rbt.OpenType.Font.outline (outlineFont entry) glyphId
        |> Array.map (function
            | Rbt.OpenType.Outline.MoveTo (x, y) -> PathMoveTo (x, y)
            | Rbt.OpenType.Outline.LineTo (x, y) -> PathLineTo (x, y)
            | Rbt.OpenType.Outline.QuadTo (cx, cy, x, y) -> PathQuadTo (cx, cy, x, y)
            | Rbt.OpenType.Outline.CubicTo (cx1, cy1, cx2, cy2, x, y) -> PathCubicTo (cx1, cy1, cx2, cy2, x, y)
            | Rbt.OpenType.Outline.Close -> PathClose)

let createRangeMeasurer fontSet : Rbt.Word.Layout.TextRangeMeasurer =
    fun text style start length ->
        match shapeText fontSet text style with
        | Some (_, shaped) -> Rbt.Text.Engine.measureRange shaped start length
        | None ->
            let startIndex = max 0 (min text.Length start)
            let endIndex = max startIndex (min text.Length (start + length))
            if endIndex = startIndex then 0.0
            else measureText fontSet (text.Substring(startIndex, endIndex - startIndex)) style

let advanceAt shapedText offset = Rbt.Text.Engine.advanceAt shapedText offset

let measureRange shapedText start length = Rbt.Text.Engine.measureRange shapedText start length
