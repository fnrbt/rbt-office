/// What Excel 2019 makes of a font at 100 % on Windows (96 dpi): its maximum digit width - the unit of
/// column widths - and the height of a row of one line of it. Measured in Excel 2019 with each font as the
/// workbook's Normal font (`tools/excel-parity/make-font-metrics.py`, docs/excel-parity-tracker.md):
/// a column of ColumnWidth 10 and 20 (10 x mdw + padding, 20 x mdw + padding) and the sheet's StandardHeight,
/// by the font's pixel size (fractional sizes too: 10.5 pt is 14 px).
module Rbt.Spreadsheet.FontMetrics

open System

/// The pixel size Excel sets a font of `points` at (96 dpi): round(points x 4/3), a half up - 10.5 pt is 14 px,
/// 8 and 8.5 pt are both 11 px and measure the same.
let pixelSize (points: float) : int = int (Math.Floor (Math.Round (points * 4.0 / 3.0, 9) + 0.5))

let private parsed =
    lazy (
        FontMetricsData.regular
        |> List.map (fun (face, text) ->
            face,
            text.Split(' ')
            |> Array.map (fun entry ->
                let parts = entry.Split(':')
                float parts.[0], float parts.[1], float parts.[2]))
        |> dict)

/// The entries of a face; a weight of a family measured is the family's; a face Excel does not have is drawn in
/// its substitute, Arial ("Arial MT", Helvetica and a name no font has measure as Arial in Excel 2019).
let private entriesOf (face: string option) =
    let key = (defaultArg face "Calibri").Trim().ToLowerInvariant()
    // ("Tw Cen MT" is a face of its own; "Arial MT" Arial.)
    let key = if key.EndsWith " mt" && not (parsed.Value.ContainsKey key) then key.Substring(0, key.Length - 3) else key
    match parsed.Value.TryGetValue key with
    | true, entries -> entries
    | false, _ ->
        // (A weight or width of a family - "Segoe UI Semibold", "Calibri Bold" - measures as its family.)
        match parsed.Value.Keys |> Seq.filter (fun family -> key.StartsWith (family + " ")) |> Seq.sortByDescending (fun family -> family.Length) |> Seq.tryHead with
        | Some family -> parsed.Value.[family]
        | None -> parsed.Value.["arial"]

/// A size's value, by its pixel size: the measured one, between two measured pixel sizes in proportion, past the
/// table in proportion to its end.
let private lookup (pick: float * float * float -> float) (face: string option) (points: float) : float =
    let entries = entriesOf face
    let size = float (pixelSize points)
    let px (s, _, _) = s
    match entries |> Array.tryFind (fun entry -> px entry = size) with
    | Some entry -> pick entry
    | None ->
        let first, last = entries.[0], entries.[entries.Length - 1]
        if size <= px first then max 1.0 (Math.Round (pick first * size / px first))
        elif size >= px last then Math.Round (pick last * size / px last)
        else
            let upper = entries |> Array.findIndex (fun entry -> px entry > size)
            let a, b = entries.[upper - 1], entries.[upper]
            Math.Round (pick a + (pick b - pick a) * (size - px a) / (px b - px a))

/// The maximum digit width of a font (face, points) in pixels: 7 for Calibri 11 and Arial 10, 9 for Arial 12.
let maxDigitWidth (face: string option) (points: float) : float =
    lookup (fun (_, mdw, _) -> mdw) face points

/// The height of a row of one line of a font, in pixels: 20 for Calibri 11 (15 pt), 17 for Arial 10 (12.75 pt).
/// A row of n lines of it is n times that.
let rowHeightPx (face: string option) (points: float) : float =
    lookup (fun (_, _, px) -> px) face points

/// The face of the tables a face is measured as (`entriesOf`'s): its own, its family's, Arial.
let private tableFace (face: string option) =
    let key = (defaultArg face "Calibri").Trim().ToLowerInvariant()
    let key = if key.EndsWith " mt" && not (parsed.Value.ContainsKey key) then key.Substring(0, key.Length - 3) else key
    if parsed.Value.ContainsKey key then key
    else
        match parsed.Value.Keys |> Seq.filter (fun family -> key.StartsWith (family + " ")) |> Seq.sortByDescending (fun family -> family.Length) |> Seq.tryHead with
        | Some family -> family
        | None -> "arial"

let private styleKey (bold: bool) (italic: bool) = (if bold then "b" else "") + (if italic then "i" else "")

let private styledDeltas =
    lazy (
        FontMetricsData.styled
        |> List.map (fun (key, text) ->
            key, text.Split(' ') |> Array.map (fun entry -> let parts = entry.Split(':') in int parts.[0], float parts.[1]) |> dict)
        |> dict)

let private descentTable = lazy (FontMetricsData.descents |> dict)

let private base36 (c: char) = if c <= '9' then int c - int '0' else int c - int 'a' + 10

/// GDI's descent of a font (tmDescent), in pixels: under 8 points Windows' Small Fonts' (as Excel draws it); a
/// face GDI does not have, Arial's.
let descentPx (face: string option) (points: float) (bold: bool) (italic: bool) : float =
    let px = pixelSize points
    let key =
        if points < 8.0 then "small fonts|"
        else
            let face = tableFace face
            let own = face + "|" + styleKey bold italic
            if descentTable.Value.ContainsKey own then own
            elif descentTable.Value.ContainsKey (face + "|") then face + "|"
            else "arial|" + styleKey bold italic
    match descentTable.Value.TryGetValue key with
    | true, text ->
        if px < 6 then float (base36 text.[0])
        elif px <= 60 then float (base36 text.[px - 6])
        else Math.Round (float (base36 text.[54]) * float px / 60.0)
    | false, _ -> Math.Round (float px * 0.21)
// Printer baseline levels from GDI TEXTMETRIC at 96 dpi. Only the six face/style families checked against
// Excel's print PDFs are used here; screen and AutoFit continue to use tmDescent alone. Each character is
// max(tmDescent, tmInternalLeading) at pixel heights 6..60 (gdi-text-metrics.tsv).
let private printLevels =
    lazy (
        [ "arial|", "1123333333344445556566667777788899999aaaaaabbbccccddddd"
          "arial|b", "112223333444445555666777777789999999abbbbbbbcccdddddeee"
          "arial|i", "11333333333455556666666777888889999aaaabbbbbbbdddddddee"
          "calibri|", "222233334445555666677788889999aaaabbbccccddddeeeffffggg"
          "calibri|b", "222233334445555666677788889999aaaabbbccccddddeeeffffggg"
          "calibri|i", "222233334445555666677788889999aaaabbbccccddddeeeffffggg"
          "segoe ui|", "2333223455567689888aaaabbbdccccdddeefggggfhhhhiiijijjjl"
          "segoe ui|b", "2333223455567689888aaaabbbdccccdddeefggggfhhhhiiijijjjl"
          "segoe ui|i", "2333223455567689888aaaabbbdccccdddeefggggfhhhhiiijijjjl"
          "segoe ui black|", "2333223455567689888aaaabbbdccccdddeefggggfhhhhiiijijjjl"
          "segoe ui black|b", "2333223455567689888aaaabbbdccccdddeefggggfhhhhiiijijjjl"
          "segoe ui black|i", "2333223455567689888aaaabbbdccccdddeefggggfhhhhiiijijjjl"
          "times new roman|", "1223233334444555566667777788889999aaaabbbbccccddddeeeee"
          "times new roman|b", "12332233334445555555666677777888899999aaaabbbbccccccddd"
          "times new roman|i", "1222233334444555566667777788889999aaaabbbbccccddddeeeee"
          "trebuchet ms|", "1256556545665566666778888889a99abaabbccccddddefeeeggffh"
          "trebuchet ms|b", "1256556545665566666778888889a99abaabbccccddddefeeeggffh"
          "trebuchet ms|i", "1256556545665566666778888889a99abaabbccccddddefeeeggffh" ]
        |> dict)

/// GDI's printer baseline level (larger of descent and internal leading), in 96-dpi pixels.
/// Unmeasured families retain the existing descent rather than guessing their print positioning.
let printLevelPx (face: string option) (points: float) (bold: bool) (italic: bool) =
    let px = pixelSize points
    let family = tableFace face
    let key = family + "|" + styleKey bold italic
    match printLevels.Value.TryGetValue key with
    | true, text when px >= 6 && px <= 60 -> float (base36 text.[px - 6])
    | _ when family = "aptos" || family = "aptos narrow" || family = "aptos display" || family = "lato" ->
        // These Office/cloud fonts are not installed on the GDI host; Excel uses their Windows ascent/descent.
        let ascent, descent, em = if family = "lato" then 2233.0, 599.0, 2000.0 else 2068.0, 563.0, 2048.0
        let rounded units = Math.Floor (units * float px / em + 0.5)
        let a, d = rounded ascent, rounded descent
        max d (a + d - float px)
    | _ -> descentPx face points bold italic

/// How far Excel 2019's printed layout sets an italic run left of where its PDF font's advances put it, in points,
/// from the chosen face's Windows ascent: `floor(round(usWinAscent x heightDots / unitsPerEm) / 5) x 72/600`, where
/// `heightDots = round(points x scale x 600 / 72)` is GDI's font height at 600 dpi. The printer DC measurement and
/// the blind holdouts are in docs/excel-parity-tracker.md ("Print italic reserve"): Aptos Narrow 14 23 dots
/// (2.76 pt), Calibri 11 17 (2.04), Arial 14 21 (2.52), Georgia 15 23 (2.76), Tahoma 16 at 80 % 21 (2.52).
/// The ascent is usWinAscent scaled and rounded; GDI's own tmAscent differs by one dot for a few face/size pairs
/// (Arial 14 at 75 % is 79 dots, where this gives 80 and 1.92 pt against Excel's 1.80; errors are at most
/// 0.12 pt). Excel applies it to right-aligned text only; centred text takes about half, left-aligned none, and
/// numeric formats with an active `*` fill (accounting) or `###` overflow use their own layout.
let printItalicShiftPt (winAscent: float) (unitsPerEm: float) (points: float) (scale: float) : float =
    if unitsPerEm <= 0.0 || points <= 0.0 || scale <= 0.0 then 0.0
    else
        let heightDots = Math.Floor (points * scale * 600.0 / 72.0 + 0.5)
        let ascentDots = Math.Floor (winAscent * heightDots / unitsPerEm + 0.5)
        Math.Floor (ascentDots / 5.0) * 72.0 / 600.0

/// These are *printer* metrics, not row-height scales (the latter do not predict printed baselines).
let private printNormalLevels =
    dict [ ("calibri", 15), (4.0, 25.0); ("calibri", 27), (7.0, 45.0); ("calibri", 11), (3.0, 18.0)
           ("segoe ui", 13), (4.0, 27.0); ("segoe ui", 12), (3.0, 25.0)
           ("arial", 11), (3.0, 15.0); ("arial", 13), (3.0, 18.0); ("arial", 16), (3.0, 22.0)
           ("times new roman", 16), (4.0, 24.0); ("cambria", 19), (4.0, 26.0)
           ("tahoma", 13), (3.0, 17.0)
           ("aptos narrow", 13), (4.0, 24.0); ("aptos narrow", 15), (4.0, 26.0)
           ("aptos narrow", 16), (4.0, 28.0); ("aptos", 15), (4.0, 26.0)
           ("aptos display", 15), (4.0, 26.0); ("lato", 15), (6.0, 39.0) ]

let printNormalLevelsOf (face: string) (points: float) : (float * float) option =
    let key = face.Trim().ToLowerInvariant(), pixelSize points
    match printNormalLevels.TryGetValue key with
    | true, levels -> Some levels
    | _ -> None


/// A font's line in a row and its descent, in pixels: the row of a line of it (a bold or italic one's own where
/// Excel's differs: Arial 12 bold 21 px, regular 20) and how much of that is below the baseline. The fonts of a
/// row meet at their baseline (`SheetOps.fitRowsOnOpen`).
let lineOf (face: string option) (points: float) (bold: bool) (italic: bool) : struct (float * float) =
    let plain = rowHeightPx face points
    let line =
        if not bold && not italic then plain
        else
            match styledDeltas.Value.TryGetValue (tableFace face + "|" + styleKey bold italic) with
            | true, deltas ->
                (match deltas.TryGetValue (pixelSize points) with
                 | true, delta -> plain + delta
                 | false, _ -> plain)
            | false, _ -> plain
    struct (line, descentPx face points bold italic)

/// The padding a column has besides its characters, in pixels: 5 for a maximum digit width of 7 or 8, 7 for
/// 9 to 12 (twice a quarter of it, rounded up, and one).
let columnPaddingPx (mdw: float) : float = 2.0 * Math.Ceiling (mdw / 4.0) + 1.0

/// The pixels of a column of a file's width (`<col width>`, characters of the maximum digit width with the
/// padding in them): trunc(((256 w + trunc(128 / mdw)) / 256) x mdw) - 8.8984375 is 80 px at 9, 9.140625 64 at 7.
let columnWidthPx (mdw: float) (width: float) : float =
    if width <= 0.0 then 0.0
    else Math.Floor (((256.0 * width + Math.Floor (128.0 / mdw)) / 256.0) * mdw + 1e-9)

/// The file's width of a column of these pixels (what `columnWidthPx` reads back as them).
let columnWidthOfPx (mdw: float) (px: float) : float = px / mdw

/// The width of a column of no width of its own: `baseColWidth` characters (8) and the padding, up to a
/// multiple of 8 px - 64 at 7, 80 at 9.
let baseColumnWidthPx (mdw: float) (characters: float) : float =
    Math.Ceiling ((characters * mdw + columnPaddingPx mdw) / 8.0) * 8.0

/// Pixels of a height in points as Excel 2019 shows it (a row's `ht`, a sheet's own default height): whole
/// pixels of 15 twentieths of a point, cut down - but a height a twentieth of a point short of a pixel is that
/// pixel (synthetic-row-heights "Custom": 15.65 pt is 20 px, 15.7 is 21; 20.15 is 26, 20.2 is 27; 20 is 26).
let pointsToRowPx (points: float) : float = Math.Floor ((Math.Floor (points * 20.0 + 0.5) + 1.0) / 15.0)
