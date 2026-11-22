/// The width of a cell's text as Excel 2019 draws it on Windows (GDI), at a zoom.
///
/// Checked in Excel on generated cells (docs/excel-parity-tracker.md, "Cell text at a zoom"):
/// - Excel draws a cell's text at a whole pixel size: its points x 96/72 x the zoom, rounded
///   (`pixelSize`): 11 pt is 15 px at 100 %, 12 px at both 80 % and 85 %; Arial 10 is 13 px.
/// - From 11 px up, in the cell's font, every character as wide as GDI hints it at that size: a
///   whole number of pixels, which is the design width rounded or a pixel either side of it
///   (Calibri's "0" is 7 px at 15 px, not 8; its "n" 7 px at 12 px, not 6). `GdiAdvances` has
///   these widths for the usual Office fonts.
/// - Under 8 points on screen (points x zoom, to a twentieth), whatever the font, in Windows' raster Small Fonts (the
///   face as high as the pixel size, or the next lower), a pixel wider a character when bold -
///   Calibri 11 at 67 % or 70 %, Arial 7.5 at 100 %.
/// - Without kerning.
/// A font the tables do not have is measured by the browser, a character at a time, each rounded.
module Rbt.Spreadsheet.ExcelText

open System.Collections.Generic

/// The whole pixel size Excel draws text of `points` at `zoom` (1.0 is 100 %), rounded to the
/// nearest (a half up: Arial 13.5 at 75 % is drawn at 14 px).
let pixelSize (points: float) (zoom: float) : int =
    max 1 (int (floor (System.Math.Round (points * 96.0 / 72.0 * zoom, 9) + 0.5)))

/// Whether Excel draws text of `points` at `zoom` in Small Fonts: under 8 points on screen, the size
/// at the zoom taken to the nearest twentieth of a point (synthetic-small-fonts: Calibri and Arial 7 -
/// 16 pt at 23 zooms) - Arial 7.5 at 100 %, Calibri 11 at 70 % and 72 % (7.92), Calibri 15 at 53 %
/// (7.95), Arial 10.5 at 75 % (7.875: 10.5 px, Small Fonts' 11-px face); not Calibri 14 at 57 % or
/// 10.5 at 76 % (7.98, which is 8), Calibri 8 at 100 %, Arial 10 at 80 %.
let isSmallFonts (points: float) (zoom: float) = System.Math.Floor (System.Math.Round (points * zoom * 20.0, 9) + 0.5) < 160.0

/// What Excel 2019 leaves out of a number's room in its cell beyond its padding for the slant or weight GDI adds to a
/// font (`CellText.NumberFit.Overhang`), at pixel size `px`: italic in a TrueType font round(px x 0.16) - 2 px at
/// 11-15 px, 3 at 16-21, 4 at 24, 5 at 29 -, in Small Fonts (`smallFonts`, its italic made up) half the size,
/// cut down - 3 px at 7, 4 at 8, 5 at 10 -, and bold in Small Fonts a pixel more (checked in Excel on generated
/// cells, make-number-room-fixture.py: Aptos Narrow, Calibri, Arial, Times New Roman, upright, italic and bold at
/// 8-22 pt and 50-100 %: the widths at which -100.00%, 50.00%, 1,234,567 and -1,234.50 show, and the #s before;
/// Calibri 16 pt italic, 21 px, has 4).
let overhang (px: int) (smallFonts: bool) (bold: bool) (italic: bool) : float =
    let slant = if not italic then 0.0 elif smallFonts then floor (float px / 2.0) else floor (float px * 0.16 + 0.5)
    slant + (if bold && smallFonts then 1.0 else 0.0)

let private base36 (c: char) = if c <= '9' then int c - int '0' else int c - int 'a' + 10

type private Face =
    { UnitsPerEm: float
      Design: int[]
      /// Per size from `GdiAdvances.firstHintedPx`, the hinted differences by character index (lazily).
      Segments: string[]
      Hinted: Dictionary<int, int[]> }

let private faces =
    lazy
        (let table = Dictionary<string, Face>()
         for (family, bold, italic, upm, design, segments) in GdiAdvances.faces do
             let advances =
                 Array.init (design.Length / 3) (fun i ->
                     let s = design.Substring(i * 3, 3)
                     if s = "zzz" then -1 else base36 s.[0] * 1296 + base36 s.[1] * 36 + base36 s.[2])
             table.[family + (if bold then "|b" else "|") + (if italic then "i" else "")] <-
                 { UnitsPerEm = float upm; Design = advances; Segments = segments; Hinted = Dictionary() }
         table)

let private smallFontFaces =
    lazy
        (GdiAdvances.smallFonts
         |> Array.map (fun (em, widths) -> em, widths |> Seq.map (fun c -> base36 c) |> Array.ofSeq))

let private faceKey (family: string) (bold: bool) (italic: bool) =
    family.Trim().ToLowerInvariant() + (if bold then "|b" else "|") + (if italic then "i" else "")

/// The face of the tables for a family and style, and whether its bold is GDI's made-up one (a
/// family without a bold face: a pixel wider a character).
let private findFace (family: string) (bold: bool) (italic: bool) : (Face * bool) option =
    let table = faces.Force()
    let get b i = match table.TryGetValue(faceKey family b i) with | true, face -> Some face | _ -> None
    match get bold italic with
    | Some face -> Some (face, false)
    | None ->
        match (if italic then get bold false else None) with
        | Some face -> Some (face, false)
        | None ->
            if not bold then None
            else
                match get false italic |> Option.orElse (get false false) with
                | Some face -> Some (face, true)
                | None -> None

/// Whether the tables know a family (its text is measured without the browser).
let knowsFamily (family: string) = findFace family false false |> Option.isSome

let private rounded (value: float) = int (floor (value + 0.5))

let private hintedDifferences (face: Face) (px: int) : int[] =
    match face.Hinted.TryGetValue px with
    | true, row -> row
    | _ ->
        let row = Array.zeroCreate face.Design.Length
        let segment = face.Segments.[px - GdiAdvances.firstHintedPx]
        let mutable i = 0
        while i + 2 < segment.Length do
            let index = System.Convert.ToInt32(segment.Substring(i, 2), 16)
            row.[index] <- int segment.[i + 2] - int 'j'
            i <- i + 3
        face.Hinted.[px] <- row
        row

/// The advance of one character (a UTF-16 code unit) in whole pixels as Excel draws it at pixel size
/// `px` (in Small Fonts where `smallFonts`, `isSmallFonts`; else the face's hinted advance), where
/// the tables know it.
let advance (family: string) (bold: bool) (italic: bool) (px: int) (smallFonts: bool) (code: int) : int voption =
    if smallFonts then
        if code < 0x20 || code > 0xFF then ValueNone
        else
            let faces = smallFontFaces.Force()
            let mutable chosen = snd faces.[0]
            for (em, widths) in faces do
                if em <= px then chosen <- widths
            ValueSome (chosen.[code - 0x20] + (if bold then 1 else 0))
    else
        let index = GdiAdvances.charIndex code
        if index < 0 then ValueNone
        else
            match findFace family bold italic with
            | None -> ValueNone
            | Some (face, madeBold) ->
                let design = face.Design.[index]
                if design < 0 then ValueNone
                else
                    let plain = rounded (float design * float px / face.UnitsPerEm)
                    let hinted =
                        if px >= GdiAdvances.firstHintedPx && px <= GdiAdvances.lastHintedPx then plain + (hintedDifferences face px).[index]
                        else plain
                    ValueSome (hinted + (if madeBold then 1 else 0))

/// The width in screen pixels of `text` as Excel draws it in a family / style at pixel size `px`
/// (`smallFonts`: in Small Fonts): the sum of its characters' advances; `measure` gives the
/// browser's width of a character the tables do not know at that size (rounded here, as GDI's are
/// whole pixels).
let textWidth (measure: string -> float) (family: string) (bold: bool) (italic: bool) (px: int) (smallFonts: bool) (text: string) : float =
    let mutable total = 0
    for i in 0 .. text.Length - 1 do
        let code = int text.[i]
        match advance family bold italic px smallFonts code with
        | ValueSome w -> total <- total + w
        | ValueNone ->
            // (A surrogate pair is measured as its character, on its first half.)
            if code >= 0xDC00 && code <= 0xDFFF && i > 0 && int text.[i - 1] >= 0xD800 && int text.[i - 1] <= 0xDBFF then ()
            elif code >= 0xD800 && code <= 0xDBFF && i + 1 < text.Length then total <- total + rounded (measure (text.Substring(i, 2)))
            else total <- total + rounded (measure (string text.[i]))
    float total

/// The widest digit of a font (the Normal font of a workbook) as Excel 2019 prints it, in dots of 600 dpi: each digit's
/// advance at the font's 600-dpi pixel size (points x 600 / 72, rounded) - 47 for Calibri 11, 46 for Arial 10, 51 for
/// Calibri 12, 56 for Arial 12 ("Arial MT" as Arial). (make-print-pages-fixture.py "widths": the PDF's columns.)
let printDigitWidth (family: string) (points: float) : float =
    let family = let f = family.Trim() in if f.EndsWith " MT" then f.Substring(0, f.Length - 3) else f
    let dots = int (System.Math.Floor (points * 600.0 / 72.0 + 0.5))
    let mutable widest = 0
    for code in int '0' .. int '9' do
        match advance family false false dots false code with
        | ValueSome w when w > widest -> widest <- w
        | _ -> ()
    if widest > 0 then float widest else FontMetrics.maxDigitWidth (Some family) (points * 6.25)

/// The row of a Normal font (family, points) as Excel 2019 prints it, in dots of 600 dpi: a row of the font's
/// default height in its PDF (make-print-pages-fixture.py "heights" and "rowunit-*": 32 fonts and sizes).
let private printedRows : (string * float * float) list =
    [ "calibri", 9.0, 99.0; "calibri", 10.0, 109.0; "calibri", 11.0, 121.0; "calibri", 12.0, 130.0; "calibri", 14.0, 150.0
      "calibri light", 11.0, 121.0; "arial", 8.0, 85.0; "arial", 9.0, 93.0; "arial", 10.0, 103.0; "arial", 11.0, 114.0
      "arial", 12.0, 123.0; "arial narrow", 10.0, 106.0; "aptos", 11.0, 127.0; "aptos", 12.0, 137.0
      "aptos narrow", 10.0, 116.0; "aptos narrow", 11.0, 127.0; "aptos narrow", 12.0, 137.0; "cambria", 14.0, 145.0
      "tahoma", 8.0, 89.0; "tahoma", 9.0, 98.0; "tahoma", 10.0, 108.0; "tahoma", 11.0, 119.0; "verdana", 10.0, 108.0
      "segoe ui", 9.0, 114.0; "segoe ui", 10.0, 124.0; "segoe ui", 11.0, 138.0; "times new roman", 10.0, 106.0
      "times new roman", 12.0, 126.0; "book antiqua", 10.0, 110.0; "century gothic", 10.0, 108.0; "garamond", 11.0, 111.0
      "georgia", 10.0, 104.0; "trebuchet ms", 10.0, 113.0 ]

/// Dots of 600 dpi a pixel of a row's height prints at, as Excel 2019 prints a sheet whose Normal font is (family,
/// points): the font's own row in print (`printedRows`: 121 dots for Calibri 11, 103 for Arial 10, 127 for Aptos
/// Narrow 11, 108 for Tahoma 10) over its row on the screen (`rowPx`: 20, 17, 20, 17 px) - rows of 7.5 to 409.5 pt
/// in 9 of them within a dot of floor(points x 4/3 x scale + 0.14) (151 of 153). A font not measured: 6.2, about
/// the middle of those measured (5.55 for Garamond 11 to 7.1 for Segoe UI 9).
let printRowScale (family: string) (points: float) (rowPx: float) : float =
    let family = let f = family.Trim().ToLowerInvariant() in if f.EndsWith " mt" then f.Substring(0, f.Length - 3) else f
    match printedRows |> List.tryFind (fun (f, p, _) -> f = family && p = points) with
    | Some (_, _, dots) when rowPx > 0.0 -> dots / rowPx
    | _ -> 6.2
