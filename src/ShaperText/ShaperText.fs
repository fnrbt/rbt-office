module RedBat.Word.ShaperText

open RedBat.Word.TextBuffer

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

type private FontEntry = {
    Family: string
    Bold: bool
    Italic: bool
    Font: Text.Engine.TextFont
    OutlineFont: OpenType.Font.Font
}

type FontSet = private {
    DefaultFamily: string
    Fonts: Map<FontKey, FontEntry>
}

let private normalizeFamily (value: string) =
    value.Trim().ToLowerInvariant()

let private fontKey family bold italic = normalizeFamily family, bold, italic

let createFontSet defaultFamily (faces: FontFace seq) =
    let fonts =
        faces
        |> Seq.map (fun (face: FontFace) ->
            let entry = {
                Family = face.Family
                Bold = face.Bold
                Italic = face.Italic
                Font = Text.Engine.loadFont face.Data
                OutlineFont = OpenType.Font.load face.Data
            }
            fontKey face.Family face.Bold face.Italic, entry)
        |> Map.ofSeq
    { DefaultFamily = defaultFamily; Fonts = fonts }

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
        | Some entry -> Text.Engine.measureWidth entry.Font style.FontSize text

let tryMeasureText (fontSet: FontSet) (text: string) (style: TextStyle) =
    if System.String.IsNullOrEmpty text then Some 0.0
    else
        match tryFindFont fontSet style.FontFamily style.Bold style.Italic with
        | None -> None
        | Some entry -> Some (Text.Engine.measureWidth entry.Font style.FontSize text)

let tryLineMetrics (fontSet: FontSet) (style: TextStyle) =
    match tryFindFont fontSet style.FontFamily style.Bold style.Italic with
    | None -> None
    | Some entry ->
        let metrics = Text.Engine.metrics entry.Font
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

let createMeasurer fontSet : RedBat.Word.Layout.TextMeasurer =
    fun text style -> measureText fontSet text style

let shapeText (fontSet: FontSet) (text: string) (style: TextStyle) =
    match tryFindFont fontSet style.FontFamily style.Bold style.Italic with
    | None -> None
    | Some entry -> Some (entry.Font, Text.Engine.shapeText entry.Font style.FontSize text)

let unitsPerEm (fontSet: FontSet) (style: TextStyle) =
    match tryFindFont fontSet style.FontFamily style.Bold style.Italic with
    | None -> 1000
    | Some entry -> entry.OutlineFont.Head.UnitsPerEm

let glyphPath (fontSet: FontSet) (style: TextStyle) glyphId =
    match tryFindFont fontSet style.FontFamily style.Bold style.Italic with
    | None -> [||]
    | Some entry ->
        OpenType.Font.outline entry.OutlineFont glyphId
        |> Array.map (function
            | OpenType.Outline.MoveTo (x, y) -> PathMoveTo (x, y)
            | OpenType.Outline.LineTo (x, y) -> PathLineTo (x, y)
            | OpenType.Outline.QuadTo (cx, cy, x, y) -> PathQuadTo (cx, cy, x, y)
            | OpenType.Outline.CubicTo (cx1, cy1, cx2, cy2, x, y) -> PathCubicTo (cx1, cy1, cx2, cy2, x, y)
            | OpenType.Outline.Close -> PathClose)

let createRangeMeasurer fontSet : RedBat.Word.Layout.TextRangeMeasurer =
    fun text style start length ->
        match shapeText fontSet text style with
        | Some (_, shaped) -> Text.Engine.measureRange shaped start length
        | None ->
            let startIndex = max 0 (min text.Length start)
            let endIndex = max startIndex (min text.Length (start + length))
            if endIndex = startIndex then 0.0
            else measureText fontSet (text.Substring(startIndex, endIndex - startIndex)) style

let advanceAt shapedText offset = Text.Engine.advanceAt shapedText offset

let measureRange shapedText start length = Text.Engine.measureRange shapedText start length
