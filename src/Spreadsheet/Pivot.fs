/// A minimal pivot-table engine: group source rows by a row field, optionally
/// cross-tab by a column field, aggregating a value field. Pure and
/// renderer-agnostic — it turns a 2-D source (row 0 = headers) into a 2-D
/// pivot grid (headers + Grand Total row/column) that a caller can place into
/// a model for rendering.
module RedBat.Spreadsheet.Pivot

open RedBat.Spreadsheet.Model

type Agg =
    | Sum
    | Count
    | Average
    | Max
    | Min

let aggLabel =
    function
    | Sum -> "Sum"
    | Count -> "Count"
    | Average -> "Average"
    | Max -> "Max"
    | Min -> "Min"

let rec private textOf (v: CellValue) : string =
    match v with
    | String s -> s
    | Number n -> n.ToString("R", System.Globalization.CultureInfo.InvariantCulture)
    | Boolean b -> if b then "TRUE" else "FALSE"
    | Formula (_, cached) -> textOf cached
    | _ -> ""

let rec private num (v: CellValue) : float option =
    match v with
    | Number n -> Some n
    | Boolean b -> Some(if b then 1.0 else 0.0)
    | Formula (_, cached) -> num cached
    | _ -> None

/// Grouping key — case-insensitive, matching how the rest of the engine
/// compares text.
let private key (v: CellValue) : string = (textOf v).ToLowerInvariant()

let private aggregate (agg: Agg) (vals: CellValue list) : CellValue =
    let nums = vals |> List.choose num
    match agg with
    | Count -> CellValue.Number(float (List.length vals))
    | Sum -> CellValue.Number(List.sum nums)
    | Average -> if List.isEmpty nums then CellValue.Empty else CellValue.Number(List.average nums)
    | Max -> if List.isEmpty nums then CellValue.Empty else CellValue.Number(List.max nums)
    | Min -> if List.isEmpty nums then CellValue.Empty else CellValue.Number(List.min nums)

/// Compute a pivot. `source.[0]` is the header row; `rowField`/`colField`/
/// `valueField` are 0-based column indices into the source. Returns a 2-D grid
/// (header row, one row per distinct row value, then a Grand Total row;
/// cross-tabbed output adds a column per distinct col value plus Grand Total).
let computePivot (source: CellValue[][]) (rowField: int) (colField: int option) (valueField: int) (agg: Agg) : CellValue[][] =
    if source.Length < 1 then [||]
    else
        let headers = source.[0]
        let data = if source.Length > 1 then source.[1..] else [||]
        let at (row: CellValue[]) i = if i >= 0 && i < row.Length then row.[i] else CellValue.Empty
        let hdr i = at headers i
        let valuesWhere (pred: CellValue[] -> bool) =
            data |> Array.filter pred |> Array.toList |> List.map (fun r -> at r valueField)
        let distinct field =
            data
            |> Array.map (fun r -> at r field)
            |> Array.toList
            |> List.distinctBy key
            |> List.sortWith compareCellValues
        let rowKeys = distinct rowField
        let corner = CellValue.String("Row Labels")
        let grandLabel = CellValue.String("Grand Total")
        match colField with
        | None ->
            let header = [| hdr rowField; CellValue.String(aggLabel agg + " of " + textOf (hdr valueField)) |]
            let body =
                rowKeys
                |> List.map (fun rk ->
                    [| rk; aggregate agg (valuesWhere (fun r -> key (at r rowField) = key rk)) |])
            let grand = [| grandLabel; aggregate agg (valuesWhere (fun _ -> true)) |]
            Array.ofList (header :: body @ [ grand ])
        | Some cf ->
            let colKeys = distinct cf
            let header =
                Array.concat [ [| corner |]; Array.ofList colKeys; [| grandLabel |] ]
            let body =
                rowKeys
                |> List.map (fun rk ->
                    let inRow r = key (at r rowField) = key rk
                    let cells =
                        colKeys
                        |> List.map (fun ck -> aggregate agg (valuesWhere (fun r -> inRow r && key (at r cf) = key ck)))
                    let total = aggregate agg (valuesWhere inRow)
                    Array.concat [ [| rk |]; Array.ofList cells; [| total |] ])
            let totalRow =
                let cells =
                    colKeys
                    |> List.map (fun ck -> aggregate agg (valuesWhere (fun r -> key (at r cf) = key ck)))
                Array.concat [ [| grandLabel |]; Array.ofList cells; [| aggregate agg (valuesWhere (fun _ -> true)) |] ]
            Array.ofList (header :: body @ [ totalRow ])
