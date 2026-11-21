/// What Excel 2019 makes of a font at 100 % on Windows (96 dpi): its maximum digit width - the unit of
/// column widths - and the height of a row of one line of it. Measured in Excel 2019 with each font as the
/// workbook's Normal font (`tools/excel-parity/make-sizing-fixtures.py`, docs/excel-parity-tracker.md):
/// a column of ColumnWidth 10 and 20 (10 x mdw + padding, 20 x mdw + padding) and the sheet's StandardHeight.
module Rbt.Spreadsheet.FontMetrics

open System

/// Face -> "size:mdw:rowPx" (points, pixels, pixels).
let private table : (string * string) list =
    [ "calibri", "6:5:11 7:5:12 8:6:15 9:6:16 10:7:17 10.5:7:19 11:7:20 12:8:21 13:9:23 14:10:25 16:11:28 18:12:31 20:14:35 24:16:42 28:19:48 36:24:62"
      "arial", "6:5:11 7:5:12 8:6:15 9:7:16 10:7:17 10.5:8:18 11:8:19 12:9:20 13:9:22 14:11:24 16:12:27 18:13:31 20:15:34 24:18:40 28:21:46 36:27:59"
      "times new roman", "6:5:11 7:5:12 8:6:15 9:6:16 10:6:17 10.5:7:18 11:7:20 12:8:21 13:9:22 14:9:25 16:11:27 18:12:31 20:14:35 24:16:41 28:19:47 36:24:61"
      "cambria", "6:5:11 7:5:12 8:6:14 9:7:16 10:7:17 10.5:8:18 11:8:19 12:9:21 13:9:22 14:11:24 16:12:27 18:13:30 20:15:34 24:18:40 28:20:46 36:27:60"
      "verdana", "6:5:11 7:5:12 8:7:14 9:8:15 10:8:17 10.5:9:18 11:10:19 12:10:20 13:11:21 14:12:24 16:13:26 18:15:30 20:17:33 24:20:39 28:24:47 36:31:61"
      "tahoma", "6:5:11 7:5:12 8:6:14 9:7:15 10:7:17 10.5:8:18 11:8:19 12:9:20 13:9:22 14:10:24 16:11:26 18:13:30 20:15:34 24:17:40 28:20:46 36:26:59"
      "segoe ui", "6:5:11 7:5:12 8:6:14 9:6:16 10:7:19 10.5:8:21 11:8:22 12:9:23 13:9:25 14:10:27 16:11:34 18:13:35 20:15:41 24:17:50 28:20:54 36:26:70"
      "courier new", "6:5:11 7:5:12 8:7:15 9:7:16 10:8:18 10.5:8:19 11:9:20 12:10:21 13:10:23 14:11:25 16:13:28 18:14:32 20:16:35 24:19:42 28:22:48 36:29:62"
      "georgia", "6:5:11 7:5:12 8:8:15 9:8:16 10:9:17 10.5:9:18 11:9:19 12:10:20 13:10:22 14:12:24 16:13:27 18:15:31 20:17:34 24:20:40 28:23:46 36:29:60"
      "calibri light", "6:5:11 7:5:12 8:6:15 9:6:16 10:7:17 10.5:7:19 11:8:20 12:8:21 13:9:23 14:10:25 16:11:28 18:12:31 20:14:35 24:16:42 28:19:48 36:24:62"
      "garamond", "6:5:11 7:5:12 8:5:15 9:6:16 10:6:17 10.5:7:19 11:7:20 12:8:21 13:8:22 14:9:25 16:10:28 18:11:31 20:13:35 24:15:41 28:17:48 36:23:62"
      "century gothic", "6:5:11 7:5:12 8:6:18 9:7:19 10:7:18 10.5:8:21 11:8:22 12:9:23 13:9:22 14:11:24 16:12:26 18:13:32 20:15:35 24:18:41 28:21:46 36:27:61"
      "aptos", "6:5:11 7:5:12 8:6:15 9:6:16 10:7:18 10.5:7:19 11:8:20 12:9:21 13:9:23 14:10:25 16:11:28 18:13:32 20:14:35 24:17:42 28:20:48 36:26:62"
      "trebuchet ms", "6:5:11 7:5:12 8:6:18 9:6:20 10:7:20 10.5:7:19 11:8:22 12:8:24 13:9:24 14:10:25 16:11:28 18:13:31 20:14:37 24:17:41 28:19:48 36:25:62"
      "ms sans serif", "6:5:11 7:5:12 8:6:14 9:6:14 10:7:17 10.5:7:17 11:7:17 12:9:21 13:9:21 14:10:26 16:11:26 18:13:31 20:14:33 24:18:41 28:20:51 36:26:61" ]

let private parsed =
    lazy (
        table
        |> List.map (fun (face, text) ->
            face,
            text.Split(' ')
            |> Array.map (fun entry ->
                let parts = entry.Split(':')
                Double.Parse(parts.[0], Globalization.CultureInfo.InvariantCulture), float parts.[1], float parts.[2]))
        |> dict)

/// The entries of a face; a face Excel does not have is drawn in its substitute, Arial ("Arial MT" is Arial).
let private entriesOf (face: string option) =
    let key = (defaultArg face "Calibri").Trim().ToLowerInvariant()
    let key = if key.EndsWith " mt" then key.Substring(0, key.Length - 3) else key
    match parsed.Value.TryGetValue key with
    | true, entries -> entries
    | false, _ -> parsed.Value.["arial"]

/// A size's value: the measured one, between two sizes in proportion, past the table in proportion to its end.
let private lookup (pick: float * float * float -> float) (face: string option) (points: float) : float =
    let entries = entriesOf face
    let size (s, _, _) = s
    match entries |> Array.tryFind (fun entry -> size entry = points) with
    | Some entry -> pick entry
    | None ->
        let first, last = entries.[0], entries.[entries.Length - 1]
        if points <= size first then max 1.0 (Math.Round (pick first * points / size first))
        elif points >= size last then Math.Round (pick last * points / size last)
        else
            let upper = entries |> Array.findIndex (fun entry -> size entry > points)
            let a, b = entries.[upper - 1], entries.[upper]
            Math.Round (pick a + (pick b - pick a) * (points - size a) / (size b - size a))

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
