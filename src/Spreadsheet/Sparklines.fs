/// Sparklines as Excel 2019 draws them in their cells (tools/excel-parity/build-sparklines.vbs:
/// synthetic-sparkline-looks.xlsx, one case a row in cells of 200 x 40 px, shot in Excel):
/// - the plot is the cell from 3 px right of its left grid line to 2 px left of its right one, and from 5 px below its
///   top grid line to 4 px above its bottom one (rows 5 .. 36 of a 40-px row: 31 px from the highest to the lowest);
/// - the n points stand at the middles of n equal slots across the plot; a column is its slot less 15 % at each side,
///   its edges rounded to pixels;
/// - the scale runs from the lowest value to the highest (of the sparkline, of its group, or the group's own): a
///   column rises from 0 (its top the pixel row its value falls in, an eighth of a pixel rounded down: ColMany 7 of 1..12
///   at row 19.09 is drawn from 19), or from the scale's end nearer to 0 when 0 is outside it - the lowest column of all-positive
///   values is a 1-px line; all values alike stand at the top;
/// - win / loss: a win the plot's upper half, a loss its lower, 0 nothing;
/// - the line is its weight in points rounded, in pixels (0.75 pt: 1 px; 2.25 pt: 2 px), empty cells break it ("gap"), are 0 ("zero") or are left out ("span");
/// - markers 4 px round, in the markers' colour, the first, last, highest, lowest and negative points in theirs;
///   columns of those points take that colour;
/// - the axis (when shown and 0 is inside the scale) a 1-px line at 0 across the plot and a pixel past it.
module Rbt.Spreadsheet.Sparklines

open System
open Rbt.Spreadsheet.Model

/// What a sparkline draws, in pixels from its cell's top left grid line.
type SparkMark =
    | SparkPath of points: (float * float) list * color: string * width: float
    | SparkBar of x: float * y: float * w: float * h: float * color: string
    | SparkDot of x: float * y: float * r: float * color: string

/// The lowest and highest known values, without allocating an intermediate array.
let range (values: float option array) : (float * float) option =
    let mutable low, high = Double.PositiveInfinity, Double.NegativeInfinity
    for value in values do
        match value with
        | Some number ->
            low <- min low number
            high <- max high number
        | None -> ()
    if low = Double.PositiveInfinity then None else Some (low, high)

/// Resolve a source sheet once per formula, then look up its cells without
/// allocating an intermediate rectangular block.
type Source = string option -> (int -> int -> CellData option) option

let values (source: Source) (formula: string) : float option array =
    let bang = formula.LastIndexOf '!'
    let sheet, reference =
        if bang < 0 then None, formula
        else
            let name = formula.Substring(0, bang).Trim()
            let name =
                if name.Length >= 2 && name.[0] = '\'' && name.[name.Length - 1] = '\''
                then name.Substring(1, name.Length - 2).Replace("''", "'")
                else name
            Some name, formula.Substring(bang + 1)
    match Rbt.Spreadsheet.SheetOps.tryParseRange (reference.Replace("$", "")) with
    | Some range when float (range.To.Row - range.From.Row + 1) * float (range.To.Col - range.From.Col + 1) <= 4096.0 ->
        match source sheet with
        | None -> [||]
        | Some cellAt ->
            [| for row in range.From.Row .. range.To.Row do
                   for col in range.From.Col .. range.To.Col do
                       match cellAt row col |> Option.map (fun c -> unwrapFormulaValue c.Value) with
                       | Some (Number v) -> yield Some v
                       | _ -> yield None |]
    | _ -> [||]

/// A group item with its data and vertical scale, including group-wide and
/// manually specified bounds. Only the requested item reads its data for
/// individually scaled groups.
let inCell (source: Source) (groups: SparklineGroup list) (addr: CellAddress) =
    groups
    |> List.tryPick (fun group ->
        group.Items
        |> List.tryFind (fun (at, _) -> at = addr)
        |> Option.bind (fun (_, formula) ->
            let points = values source formula
            let own = range points
            let mutable groupBounds = None
            let all () =
                match groupBounds with
                | Some bounds -> bounds
                | None ->
                    let bounds = group.Items |> List.choose (fun (_, f) -> range (values source f))
                    groupBounds <- Some bounds
                    bounds
            let pick kind manual ownValue groupValue =
                match kind, manual with
                | "custom", Some v -> Some v
                | "group", _ -> match all () with [] -> None | bounds -> Some (groupValue bounds)
                | _ -> own |> Option.map ownValue
            match pick group.MinAxis group.ManualMin fst (List.map fst >> List.min),
                  pick group.MaxAxis group.ManualMax snd (List.map snd >> List.max) with
            | Some low, Some high when points.Length > 0 -> Some (group, points, (low, high))
            | _ -> None))

let modelSource (model: SpreadsheetModel) : Source =
    function
    | Some _ -> None
    | None -> Some (fun row col -> getCell { Row = row; Col = col } model)

/// The marks of a sparkline of `group` over `values` (one a cell of its data, None an empty one) in a cell `width` x
/// `height` px (grid line to grid line); `scale` the lowest and highest value of its vertical scale.
let marks (group: SparklineGroup) (values: float option array) (scale: float * float) (width: float) (height: float) : SparkMark list =
    let n = values.Length
    if n = 0 || width < 8.0 || height < 8.0 then []
    else
        let values =
            if group.EmptyAs = "zero" then values |> Array.map (fun v -> match v with None -> Some 0.0 | some -> some) else values
        let values = if group.RightToLeft then Array.rev values else values
        let left, right = 3.0, width - 2.0
        let top, bottom = 5.0, height - 4.0
        let slot = (right - left) / float n
        // (A point's middle, on a pixel's middle as Excel's line falls: 16.5 for the first of 7 in 200 px.)
        let middle (i: int) = left + (float i + 0.5) * slot - 0.5
        let low, high = scale
        let yOf (v: float) = if high > low then bottom - (v - low) / (high - low) * (bottom - top) else top
        let known = values |> Array.choose id
        let firstAt = values |> Array.tryFindIndex Option.isSome
        let lastAt = values |> Array.tryFindIndexBack Option.isSome
        let highest = if known.Length > 0 then Array.max known else 0.0
        let lowest = if known.Length > 0 then Array.min known else 0.0
        // (The colour a point is marked in, if any: the last of the marks that hold for it.)
        let special (i: int) (v: float) =
            let firstHere = (if group.RightToLeft then lastAt else firstAt) = Some i
            let lastHere = (if group.RightToLeft then firstAt else lastAt) = Some i
            [ if group.Negative && v < 0.0 then group.NegativeColor
              if group.Low && v = lowest then group.LowColor
              if group.High && v = highest then group.HighColor
              if group.First && firstHere then group.FirstColor
              if group.Last && lastHere then group.LastColor ]
            |> List.tryLast
        let axis =
            if group.DisplayXAxis && low < 0.0 && high > 0.0 || (group.DisplayXAxis && group.Kind = "stacked") then
                let y = if group.Kind = "stacked" then Math.Floor((top + bottom) / 2.0) else Math.Floor(yOf 0.0)
                [ SparkBar (left, y, right - left + 1.0, 1.0, group.AxisColor) ]
            else []
        match group.Kind with
        | "column" | "stacked" ->
            let bars =
                [ for i in 0 .. n - 1 do
                    match values.[i] with
                    | Some v ->
                        let x0 = Math.Round(left + float i * slot + 0.15 * slot)
                        let x1 = Math.Round(left + float (i + 1) * slot - 0.15 * slot)
                        let span =
                            if group.Kind = "stacked" then
                                let mid = Math.Floor((top + bottom) / 2.0)
                                if v > 0.0 then Some (top, mid + 1.0) elif v < 0.0 then Some (mid + 1.0, bottom + 1.0) else None
                            else
                                let baseValue = max low (min high 0.0)
                                let b = yOf baseValue
                                let y = yOf v
                                if v >= baseValue then Some (Math.Ceiling (y - 0.125), Math.Floor b + 1.0)
                                else Some (Math.Ceiling b, Math.Floor y + 1.0)
                        match span with
                        | Some (y0, y1) when x1 > x0 ->
                            let y0, y1 = (if y1 <= y0 then y0, y0 + 1.0 else y0, y1)
                            let color = special i v |> Option.defaultValue group.SeriesColor
                            yield SparkBar (x0, y0, x1 - x0, y1 - y0, color)
                        | _ -> ()
                    | None -> () ]
            bars @ axis
        | _ ->
            // The line through the points, broken at empty cells unless they are spanned.
            let weight = max 1.0 (Math.Round group.LineWeight)
            let segments = ResizeArray<(float * float) list>()
            let current = ResizeArray<float * float>()
            for i in 0 .. n - 1 do
                match values.[i] with
                | Some v -> current.Add((middle i, yOf v + 0.5))
                | None ->
                    if group.EmptyAs <> "span" && current.Count > 0 then
                        segments.Add(List.ofSeq current)
                        current.Clear()
            if current.Count > 0 then segments.Add(List.ofSeq current)
            let lines = [ for points in segments do if points.Length > 1 then yield SparkPath (points, group.SeriesColor, weight) ]
            let dots =
                [ for i in 0 .. n - 1 do
                    match values.[i] with
                    | Some v ->
                        match special i v with
                        | Some color -> yield SparkDot (middle i, yOf v + 0.5, 2.0, color)
                        | None -> if group.Markers then yield SparkDot (middle i, yOf v + 0.5, 2.0, group.MarkersColor)
                    | None -> () ]
            axis @ lines @ dots

/// The marks as SVG of the cell's size.
let svg (width: float) (height: float) (items: SparkMark list) : string =
    let num (v: float) = (Math.Round(v, 2)).ToString(Globalization.CultureInfo.InvariantCulture)
    let body =
        items
        |> List.map (fun item ->
            match item with
            | SparkPath (points, color, w) ->
                let d = points |> List.mapi (fun i (x, y) -> (if i = 0 then "M" else "L") + num x + " " + num y) |> String.concat " "
                "<path d=\"" + d + "\" fill=\"none\" stroke=\"" + color + "\" stroke-width=\"" + num w + "\" stroke-linejoin=\"round\" stroke-linecap=\"round\"/>"
            | SparkBar (x, y, w, h, color) ->
                "<rect x=\"" + num x + "\" y=\"" + num y + "\" width=\"" + num w + "\" height=\"" + num h + "\" fill=\"" + color + "\" shape-rendering=\"crispEdges\"/>"
            | SparkDot (x, y, r, color) ->
                "<circle cx=\"" + num x + "\" cy=\"" + num y + "\" r=\"" + num r + "\" fill=\"" + color + "\"/>")
        |> String.concat ""
    "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"" + num width + "\" height=\"" + num height + "\" viewBox=\"0 0 " + num width + " " + num height + "\">" + body + "</svg>"
