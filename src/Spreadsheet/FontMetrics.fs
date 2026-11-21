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
    let key = if key.EndsWith " mt" then key.Substring(0, key.Length - 3) else key
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

/// Pixels of a height in points as Excel shows it: whole pixels, cut down (20 pt is 26 px, 15.7 is 20).
let pointsToRowPx (points: float) : float = Math.Floor (points * 4.0 / 3.0 + 1e-6)
