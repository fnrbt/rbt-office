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

/// Parse an aggregation name (from the pivot dialog's dropdown).
let parseAgg (s: string) : Agg option =
    match s.Trim().ToLowerInvariant() with
    | "sum" -> Some Sum
    | "count" -> Some Count
    | "average" | "avg" -> Some Average
    | "max" -> Some Max
    | "min" -> Some Min
    | _ -> None

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

/// A richer pivot: NESTED row fields (tabular layout, one column per level)
/// with a subtotal row after each non-innermost group, plus MULTIPLE value
/// fields (one measure column each), and a Grand Total row. This is the common
/// "grouped list with subtotals" pivot. (`rowFields` outer→inner.)
let computePivotN (source: CellValue[][]) (rowFields: int list) (values: (int * Agg) list) : CellValue[][] =
    if source.Length < 1 || List.isEmpty rowFields || List.isEmpty values then [||]
    else
        let headers = source.[0]
        let data = if source.Length > 1 then source.[1..] else [||]
        let at (row: CellValue[]) i = if i >= 0 && i < row.Length then row.[i] else CellValue.Empty
        let hdr i = at headers i
        let nf = List.length rowFields
        let measures (rows: CellValue[][]) =
            values |> List.map (fun (v, a) -> aggregate a (rows |> Array.toList |> List.map (fun r -> at r v)))
        let distinctSorted (rows: CellValue[][]) field =
            rows
            |> Array.map (fun r -> at r field)
            |> Array.toList
            |> List.distinctBy key
            |> List.sortWith compareCellValues
        // Build body rows recursively: leaf level emits a data row per value;
        // each non-leaf level emits its children then a "<value> Total" subtotal.
        let rec build (level: int) (rows: CellValue[][]) (labels: CellValue list) : CellValue list list =
            let field = List.item level rowFields
            [ for kv in distinctSorted rows field do
                let subset = rows |> Array.filter (fun r -> key (at r field) = key kv)
                let labels' = labels @ [ kv ]
                if level = nf - 1 then
                    let rowCells = [ for i in 0 .. nf - 1 -> if i < labels'.Length then labels'.[i] else CellValue.Empty ]
                    yield rowCells @ measures subset
                else
                    yield! build (level + 1) subset labels'
                    let subCells =
                        [ for i in 0 .. nf - 1 ->
                            if i < level then labels'.[i]
                            elif i = level then CellValue.String(textOf kv + " Total")
                            else CellValue.Empty ]
                    yield subCells @ measures subset ]
        let header =
            (rowFields |> List.map hdr)
            @ (values |> List.map (fun (v, a) -> CellValue.String(aggLabel a + " of " + textOf (hdr v))))
        let grand =
            (CellValue.String "Grand Total" :: List.replicate (nf - 1) CellValue.Empty) @ measures data
        (header :: build 0 data [] @ [ grand ])
        |> List.map Array.ofList
        |> Array.ofList

/// GROUPBY as a spilling function operates on *data-only* parallel ranges
/// (no header row): `rowM` = the row-field columns, `valueM` = the measure
/// columns. Returns nested groups with subtotals + a Grand Total, one measure
/// column per value column, WITHOUT a field-name header row.
let groupBy (rowM: CellValue[][]) (valueM: CellValue[][]) (agg: Agg) : CellValue[][] =
    if rowM.Length = 0 then [||]
    else
        let r = rowM.[0].Length
        let v = if valueM.Length > 0 then valueM.[0].Length else 0
        let n = min rowM.Length valueM.Length
        let header = Array.append (Array.init r (fun i -> CellValue.String(sprintf "_r%d" i))) (Array.init v (fun j -> CellValue.String(sprintf "_v%d" j)))
        let body = [| for i in 0 .. n - 1 -> Array.append rowM.[i] valueM.[i] |]
        let source = Array.append [| header |] body
        let grid = computePivotN source [ 0 .. r - 1 ] [ for j in 0 .. v - 1 -> (r + j, agg) ]
        if grid.Length > 1 then grid.[1..] else grid

/// PIVOTBY as a spilling function over data-only parallel ranges: single row
/// field × one or more (NESTED) column fields, one measure. `colM` may have
/// several columns — each becomes a column level, producing a multi-row column
/// header. Keeps per-row and grand totals.
let pivotBy (rowM: CellValue[][]) (colM: CellValue[][]) (valueM: CellValue[][]) (agg: Agg) : CellValue[][] =
    let n = [ rowM.Length; colM.Length; valueM.Length ] |> List.min
    let levels = if colM.Length > 0 then colM.[0].Length else 0
    if n = 0 || levels = 0 then [||]
    else
        let rowOf i = if rowM.[i].Length > 0 then rowM.[i].[0] else CellValue.Empty
        let valOf i = if valueM.[i].Length > 0 then valueM.[i].[0] else CellValue.Empty
        let comboOf i = colM.[i]
        let rkey (v: CellValue) = (textOf v).ToLowerInvariant()
        let ckey (a: CellValue[]) = a |> Array.map (fun v -> (textOf v).ToLowerInvariant()) |> String.concat ""
        let idx = [ 0 .. n - 1 ]
        let rowKeys = idx |> List.map rowOf |> List.distinctBy rkey |> List.sortWith compareCellValues
        let colCombos =
            idx |> List.map comboOf |> List.distinctBy ckey
            |> List.sortWith (fun a b ->
                let rec cmp k = if k >= levels then 0 else (match compareCellValues a.[k] b.[k] with 0 -> cmp (k + 1) | c -> c)
                cmp 0)
        let aggWhere pred = idx |> List.filter pred |> List.map valOf |> aggregate agg
        let headers =
            [ for lvl in 0 .. levels - 1 ->
                let corner = if lvl = 0 then CellValue.String "Row Labels" else CellValue.Empty
                let grand = if lvl = 0 then CellValue.String "Grand Total" else CellValue.Empty
                Array.concat [ [| corner |]; (colCombos |> List.map (fun c -> c.[lvl]) |> Array.ofList); [| grand |] ] ]
        let body =
            rowKeys
            |> List.map (fun rk ->
                let cells = colCombos |> List.map (fun combo -> aggWhere (fun i -> rkey (rowOf i) = rkey rk && ckey (comboOf i) = ckey combo))
                Array.concat [ [| rk |]; Array.ofList cells; [| aggWhere (fun i -> rkey (rowOf i) = rkey rk) |] ])
        let totalRow =
            let cells = colCombos |> List.map (fun combo -> aggWhere (fun i -> ckey (comboOf i) = ckey combo))
            Array.concat [ [| CellValue.String "Grand Total" |]; Array.ofList cells; [| aggWhere (fun _ -> true) |] ]
        Array.ofList (headers @ body @ [ totalRow ])

/// Find a field's 0-based column by matching the header row (case-insensitive).
let fieldIndex (headers: CellValue[]) (name: string) : int option =
    let t = name.Trim().ToLowerInvariant()
    headers |> Array.tryFindIndex (fun h -> (textOf h).Trim().ToLowerInvariant() = t)

/// Build a LIVE pivot formula (`=GROUPBY(...)` / `=PIVOTBY(...)`) from the
/// dialog's field names + source bounds. The ranges are the *data* rows of
/// each field (header row excluded). This is what makes our pivot better than
/// Excel's: the dialog writes a formula the user owns — live, composable,
/// diffable — instead of a static object. `headers` is the source's header row
/// (length ec-sc+1); column indices are absolute (offset by `sc`).
let buildFormula (sr: int) (sc: int) (er: int) (_ec: int) (headers: CellValue[])
                 (rowFieldNames: string list) (colFieldName: string option)
                 (valueFieldName: string) (aggName: string) : Result<string, string> =
    let resolve name = fieldIndex headers name |> Option.map (fun li -> sc + li)
    let rowNames = rowFieldNames |> List.map (fun s -> s.Trim()) |> List.filter (fun s -> s <> "")
    if List.isEmpty rowNames then Result.Error "Choose at least one row field."
    else
        match rowNames |> List.map (fun n -> n, resolve n) |> List.tryFind (snd >> Option.isNone) with
        | Some (missing, _) -> Result.Error(sprintf "Row field not found: %s" missing)
        | None ->
            let cols = rowNames |> List.map (resolve >> Option.get)
            match resolve valueFieldName with
            | None -> Result.Error(sprintf "Value field not found: %s" valueFieldName)
            | Some vi ->
                let dataTop = sr + 1
                let rng c0 c1 = sprintf "%s:%s" (formatCellRef dataTop c0) (formatCellRef er c1)
                let aggStr = "\"" + aggName.Trim().ToLowerInvariant() + "\""
                let rowRange = rng (List.min cols) (List.max cols)
                match colFieldName |> Option.map (fun s -> s.Trim()) |> Option.filter (fun s -> s <> "") with
                | Some cn ->
                    match resolve cn with
                    | None -> Result.Error(sprintf "Column field not found: %s" cn)
                    | Some ci -> Result.Ok(sprintf "=PIVOTBY(%s,%s,%s,%s)" rowRange (rng ci ci) (rng vi vi) aggStr)
                | None -> Result.Ok(sprintf "=GROUPBY(%s,%s,%s)" rowRange (rng vi vi) aggStr)

/// Resolve field NAMES against the source's header row and compute the pivot —
/// the entry point for the Insert-Pivot dialog. Nested row fields use
/// `computePivotN`; a column field uses the single-row cross-tab
/// `computePivot`. Returns the pivot grid or a human-readable error.
let pivotFromNames (source: CellValue[][]) (rowFieldNames: string list) (colFieldName: string option) (valueFieldName: string) (agg: Agg) : Result<CellValue[][], string> =
    if source.Length < 1 then Result.Error "Source range has no header row."
    else
        let headers = source.[0]
        let rowNames = rowFieldNames |> List.map (fun s -> s.Trim()) |> List.filter (fun s -> s <> "")
        if List.isEmpty rowNames then Result.Error "Choose at least one row field."
        else
            let rowIdxs = rowNames |> List.map (fun n -> n, fieldIndex headers n)
            match rowIdxs |> List.tryFind (snd >> Option.isNone) with
            | Some (missing, _) -> Result.Error(sprintf "Row field not found: %s" missing)
            | None ->
                let rows = rowIdxs |> List.map (snd >> Option.get)
                match fieldIndex headers valueFieldName with
                | None -> Result.Error(sprintf "Value field not found: %s" valueFieldName)
                | Some vi ->
                    match colFieldName |> Option.map (fun s -> s.Trim()) |> Option.filter (fun s -> s <> "") with
                    | Some cn ->
                        match fieldIndex headers cn with
                        | None -> Result.Error(sprintf "Column field not found: %s" cn)
                        | Some ci -> Result.Ok(computePivot source (List.head rows) (Some ci) vi agg)
                    | None -> Result.Ok(computePivotN source rows [ (vi, agg) ])
