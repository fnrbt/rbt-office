/// A minimal pivot-table engine: group source rows by a row field, optionally
/// cross-tab by a column field, aggregating a value field. Pure and
/// renderer-agnostic — it turns a 2-D source (row 0 = headers) into a 2-D
/// pivot grid (headers + Grand Total row/column) that a caller can place into
/// a model for rendering.
module RedBat.Spreadsheet.Pivot

open System
open System.Collections.Generic
open RedBat.Spreadsheet.Model

type Agg =
    | Sum
    | Count
    | Average
    | Max
    | Min

let private charEqualsIgnoreCase (a: char) (b: char) =
    a = b || Char.ToUpperInvariant a = Char.ToUpperInvariant b

let private trimmedEqualsIgnoreCase (left: string) (right: string) =
    let mutable li = 0
    let mutable lj = left.Length - 1
    while li <= lj && Char.IsWhiteSpace left.[li] do
        li <- li + 1
    while lj >= li && Char.IsWhiteSpace left.[lj] do
        lj <- lj - 1

    let mutable ri = 0
    let mutable rj = right.Length - 1
    while ri <= rj && Char.IsWhiteSpace right.[ri] do
        ri <- ri + 1
    while rj >= ri && Char.IsWhiteSpace right.[rj] do
        rj <- rj - 1

    let leftLen = lj - li + 1
    let rightLen = rj - ri + 1
    if leftLen <> rightLen then
        false
    else
        let mutable offset = 0
        let mutable equal = true
        while equal && offset < leftLen do
            if not (charEqualsIgnoreCase left.[li + offset] right.[ri + offset]) then
                equal <- false
            offset <- offset + 1
        equal

let aggLabel =
    function
    | Sum -> "Sum"
    | Count -> "Count"
    | Average -> "Average"
    | Max -> "Max"
    | Min -> "Min"

/// Parse an aggregation name (from the pivot dialog's dropdown).
let parseAgg (s: string) : Agg option =
    if trimmedEqualsIgnoreCase s "sum" then Some Sum
    elif trimmedEqualsIgnoreCase s "count" then Some Count
    elif trimmedEqualsIgnoreCase s "average" || trimmedEqualsIgnoreCase s "avg" then Some Average
    elif trimmedEqualsIgnoreCase s "max" then Some Max
    elif trimmedEqualsIgnoreCase s "min" then Some Min
    else None

let private textOf (v: CellValue) : string =
    match unwrapFormulaValue v with
    | String s -> s
    | Number n -> n.ToString("R", System.Globalization.CultureInfo.InvariantCulture)
    | Boolean b -> if b then "TRUE" else "FALSE"
    | _ -> ""

let private num (v: CellValue) : float option =
    match unwrapFormulaValue v with
    | Number n -> Some n
    | Boolean b -> Some(if b then 1.0 else 0.0)
    | _ -> None

/// Grouping key — case-insensitive, matching how the rest of the engine
/// compares text.
let private key (v: CellValue) : string =
#if FABLE_COMPILER
    (textOf v).ToLowerInvariant()
#else
    textOf v
#endif

let private stringKeyDict<'V> () =
#if FABLE_COMPILER
    Dictionary<string, 'V>()
#else
    Dictionary<string, 'V>(System.StringComparer.OrdinalIgnoreCase)
#endif

type private AggState() =
    member val Count = 0 with get, set
    member val NumericCount = 0 with get, set
    member val Sum = 0.0 with get, set
    member val MaxValue = System.Double.NegativeInfinity with get, set
    member val MinValue = System.Double.PositiveInfinity with get, set

let private addAggValue (state: AggState) (value: CellValue) : unit =
    state.Count <- state.Count + 1
    match num value with
    | Some n ->
        state.NumericCount <- state.NumericCount + 1
        state.Sum <- state.Sum + n
        if n > state.MaxValue then state.MaxValue <- n
        if n < state.MinValue then state.MinValue <- n
    | None -> ()

let private finishAggState (agg: Agg) (state: AggState) : CellValue =
    match agg with
    | Count -> CellValue.Number(float state.Count)
    | Sum -> CellValue.Number(state.Sum)
    | Average -> if state.NumericCount = 0 then CellValue.Empty else CellValue.Number(state.Sum / float state.NumericCount)
    | Max -> if state.NumericCount = 0 then CellValue.Empty else CellValue.Number(state.MaxValue)
    | Min -> if state.NumericCount = 0 then CellValue.Empty else CellValue.Number(state.MinValue)

let private finishAggOption (agg: Agg) (state: AggState option) : CellValue =
    match state with
    | Some s -> finishAggState agg s
    | None -> finishAggState agg (AggState())

let private aggregateRows (agg: Agg) (rows: CellValue[][]) valueField =
    let state = AggState()
    for row in rows do
        addAggValue state (if valueField >= 0 && valueField < row.Length then row.[valueField] else CellValue.Empty)
    finishAggState agg state

let private getOrAdd (dict: Dictionary<'K, 'V>) (key: 'K) (create: unit -> 'V) : 'V =
    let mutable value = Unchecked.defaultof<'V>
    if dict.TryGetValue(key, &value) then
        value
    else
        let value = create ()
        dict.[key] <- value
        value

let private tryFind (dict: Dictionary<'K, 'V>) (key: 'K) : 'V option =
    let mutable value = Unchecked.defaultof<'V>
    if dict.TryGetValue(key, &value) then Some value else None

let private addDistinct (dict: Dictionary<'K, 'V>) key value =
    let mutable existing = Unchecked.defaultof<'V>
    if not (dict.TryGetValue(key, &existing)) then
        dict.[key] <- value

let private arrayToList (items: 'T[]) =
    let mutable result = []
    for i = items.Length - 1 downto 0 do
        result <- items.[i] :: result
    result

let private resizeArrayToList (items: ResizeArray<'T>) =
    let mutable result = []
    for i = items.Count - 1 downto 0 do
        result <- items.[i] :: result
    result

let private sortedValuesByKeyArray (valuesByKey: Dictionary<string, CellValue>) : CellValue[] =
    let values = ResizeArray<CellValue>(valuesByKey.Count)
    for pair in valuesByKey do
        values.Add pair.Value
    let result = values.ToArray()
    Array.sortInPlaceWith compareCellValues result
    result

let private sortedValuesByKey (valuesByKey: Dictionary<string, CellValue>) : CellValue list =
    let result = sortedValuesByKeyArray valuesByKey
    arrayToList result

let private prefixedArray (header: CellValue[]) (rows: CellValue[][]) =
    let result = Array.zeroCreate<CellValue[]> (rows.Length + 1)
    result.[0] <- header
    Array.Copy(rows, 0, result, 1, rows.Length)
    result

let private comboKey (values: CellValue[]) : string =
    match values.Length with
    | 0 -> ""
    | 1 -> key values.[0]
    | 2 -> key values.[0] + "\u0001" + key values.[1]
    | _ ->
        let sb = System.Text.StringBuilder()
        for i = 0 to values.Length - 1 do
            if i > 0 then
                sb.Append('\u0001') |> ignore
            sb.Append(key values.[i]) |> ignore
        sb.ToString()

let private comboKeyFromRecord (pick: CellValue[] -> int -> CellValue) (indexes: int[]) (row: CellValue[]) : string =
    match indexes.Length with
    | 0 -> ""
    | 1 -> key (pick row indexes.[0])
    | 2 -> key (pick row indexes.[0]) + "\u0001" + key (pick row indexes.[1])
    | _ ->
        let sb = System.Text.StringBuilder()
        for i = 0 to indexes.Length - 1 do
            if i > 0 then
                sb.Append('\u0001') |> ignore
            sb.Append(key (pick row indexes.[i])) |> ignore
        sb.ToString()

let private comboFromRecord (pick: CellValue[] -> int -> CellValue) (indexes: int[]) (row: CellValue[]) =
    let values = Array.zeroCreate<CellValue> indexes.Length
    for i = 0 to indexes.Length - 1 do
        values.[i] <- pick row indexes.[i]
    values

let private sortedCombosByKeyArray (valuesByKey: Dictionary<string, CellValue[]>) (levels: int) : CellValue[][] =
    let values = ResizeArray<CellValue[]>(valuesByKey.Count)
    for pair in valuesByKey do
        values.Add pair.Value
    let result = values.ToArray()
    Array.sortInPlaceWith (fun (a: CellValue[]) (b: CellValue[]) ->
        let rec cmp k =
            if k >= levels then 0
            else
                match compareCellValues a.[k] b.[k] with
                | 0 -> cmp (k + 1)
                | c -> c
        cmp 0) result
    result

let private pivotByProjected (records: CellValue[][]) (rowIdx: int) (colIdxs: int[]) (dataIdx: int) (agg: Agg) =
    let levels = colIdxs.Length
    if records.Length = 0 || levels = 0 then [||]
    else
        let pick (r: CellValue[]) i = if i >= 0 && i < r.Length then r.[i] else CellValue.Empty
        let rowValuesByKey = stringKeyDict<CellValue>()
        let comboValuesByKey = stringKeyDict<CellValue[]>()
        let rowTotals = stringKeyDict<AggState>()
        let comboTotals = stringKeyDict<AggState>()
        let cells = stringKeyDict<Dictionary<string, AggState>>()
        let grandTotal = AggState()
        for row in records do
            let rowValue = pick row rowIdx
            let rowKey = key rowValue
            let colKey = comboKeyFromRecord pick colIdxs row
            addDistinct rowValuesByKey rowKey rowValue
            if not (comboValuesByKey.ContainsKey colKey) then
                comboValuesByKey.[colKey] <- comboFromRecord pick colIdxs row
            let value = pick row dataIdx
            addAggValue (getOrAdd rowTotals rowKey AggState) value
            addAggValue (getOrAdd comboTotals colKey AggState) value
            let rowCells = getOrAdd cells rowKey stringKeyDict<AggState>
            addAggValue (getOrAdd rowCells colKey AggState) value
            addAggValue grandTotal value
        let rowKeys = sortedValuesByKey rowValuesByKey
        let colCombos = sortedCombosByKeyArray comboValuesByKey levels
        let colComboKeys = Array.zeroCreate<string> colCombos.Length
        for i in 0 .. colCombos.Length - 1 do
            colComboKeys.[i] <- comboKey colCombos.[i]
        let result = ResizeArray<CellValue[]>(levels + rowKeys.Length + 1)
        for lvl in 0 .. levels - 1 do
            let header = Array.create (colCombos.Length + 2) CellValue.Empty
            header.[0] <- if lvl = 0 then CellValue.String "Row Labels" else CellValue.Empty
            for i in 0 .. colCombos.Length - 1 do
                header.[i + 1] <- colCombos.[i].[lvl]
            header.[colCombos.Length + 1] <- if lvl = 0 then CellValue.String "Grand Total" else CellValue.Empty
            result.Add header
        for rk in rowKeys do
            let rowKey = key rk
            let rowCells = tryFind cells rowKey
            let row = Array.create (colCombos.Length + 2) CellValue.Empty
            row.[0] <- rk
            for i in 0 .. colCombos.Length - 1 do
                let cellState =
                    match rowCells with
                    | Some rowDict -> tryFind rowDict colComboKeys.[i]
                    | None -> None
                row.[i + 1] <- finishAggOption agg cellState
            row.[colCombos.Length + 1] <- finishAggOption agg (tryFind rowTotals rowKey)
            result.Add row
        let totalRow = Array.create (colCombos.Length + 2) CellValue.Empty
        totalRow.[0] <- CellValue.String "Grand Total"
        for i in 0 .. colCombos.Length - 1 do
            totalRow.[i + 1] <- finishAggOption agg (tryFind comboTotals colComboKeys.[i])
        totalRow.[colCombos.Length + 1] <- finishAggState agg grandTotal
        result.Add totalRow
        result.ToArray()

/// Compute a pivot. `source.[0]` is the header row; `rowField`/`colField`/
/// `valueField` are 0-based column indices into the source. Returns a 2-D grid
/// (header row, one row per distinct row value, then a Grand Total row;
/// cross-tabbed output adds a column per distinct col value plus Grand Total).
let computePivot (source: CellValue[][]) (rowField: int) (colField: int option) (valueField: int) (agg: Agg) : CellValue[][] =
    if source.Length < 1 then [||]
    else
        let headers = source.[0]
        let at (row: CellValue[]) i = if i >= 0 && i < row.Length then row.[i] else CellValue.Empty
        let hdr i = at headers i
        let rowValuesByKey = stringKeyDict<CellValue>()
        let rowTotals = stringKeyDict<AggState>()
        let grandTotal = AggState()
        let corner = CellValue.String("Row Labels")
        let grandLabel = CellValue.String("Grand Total")
        match colField with
        | None ->
            for rowIndex = 1 to source.Length - 1 do
                let row = source.[rowIndex]
                let rowValue = at row rowField
                let rowKey = key rowValue
                addDistinct rowValuesByKey rowKey rowValue
                let value = at row valueField
                addAggValue (getOrAdd rowTotals rowKey AggState) value
                addAggValue grandTotal value
            let rowKeys = sortedValuesByKey rowValuesByKey
            let header = [| hdr rowField; CellValue.String(aggLabel agg + " of " + textOf (hdr valueField)) |]
            let grand = [| grandLabel; finishAggState agg grandTotal |]
            let result = ResizeArray<CellValue[]>(rowKeys.Length + 2)
            result.Add header
            for rk in rowKeys do
                let rowKey = key rk
                result.Add [| rk; finishAggOption agg (tryFind rowTotals rowKey) |]
            result.Add grand
            result.ToArray()
        | Some cf ->
            let colValuesByKey = stringKeyDict<CellValue>()
            let colTotals = stringKeyDict<AggState>()
            let cells = stringKeyDict<Dictionary<string, AggState>>()
            for rowIndex = 1 to source.Length - 1 do
                let row = source.[rowIndex]
                let rowValue = at row rowField
                let colValue = at row cf
                let rowKey = key rowValue
                let colKey = key colValue
                addDistinct rowValuesByKey rowKey rowValue
                addDistinct colValuesByKey colKey colValue
                let value = at row valueField
                addAggValue (getOrAdd rowTotals rowKey AggState) value
                addAggValue (getOrAdd colTotals colKey AggState) value
                let rowCells = getOrAdd cells rowKey stringKeyDict<AggState>
                addAggValue (getOrAdd rowCells colKey AggState) value
                addAggValue grandTotal value
            let rowKeys = sortedValuesByKey rowValuesByKey
            let colKeys = sortedValuesByKeyArray colValuesByKey
            let colKeyStrings = Array.zeroCreate<string> colKeys.Length
            for i in 0 .. colKeys.Length - 1 do
                colKeyStrings.[i] <- key colKeys.[i]
            let header = Array.create (colKeys.Length + 2) CellValue.Empty
            header.[0] <- corner
            for i in 0 .. colKeys.Length - 1 do
                header.[i + 1] <- colKeys.[i]
            header.[colKeys.Length + 1] <- grandLabel
            let result = ResizeArray<CellValue[]>(rowKeys.Length + 2)
            result.Add header
            for rk in rowKeys do
                let rowKey = key rk
                let rowCells = tryFind cells rowKey
                let row = Array.create (colKeys.Length + 2) CellValue.Empty
                row.[0] <- rk
                for i in 0 .. colKeys.Length - 1 do
                    let cellState =
                        match rowCells with
                        | Some rowDict -> tryFind rowDict colKeyStrings.[i]
                        | None -> None
                    row.[i + 1] <- finishAggOption agg cellState
                row.[colKeys.Length + 1] <- finishAggOption agg (tryFind rowTotals rowKey)
                result.Add row
            let totalRow = Array.create (colKeys.Length + 2) CellValue.Empty
            totalRow.[0] <- grandLabel
            for i in 0 .. colKeys.Length - 1 do
                totalRow.[i + 1] <- finishAggOption agg (tryFind colTotals colKeyStrings.[i])
            totalRow.[colKeys.Length + 1] <- finishAggState agg grandTotal
            result.Add totalRow
            result.ToArray()

let private computePivotNCore (headers: CellValue[]) (data: CellValue[][]) (rowFields: int list) (values: (int * Agg) list) : CellValue[][] =
    if List.isEmpty rowFields || List.isEmpty values then [||]
    else
        let at (row: CellValue[]) i = if i >= 0 && i < row.Length then row.[i] else CellValue.Empty
        let hdr i = at headers i
        let rowFields = rowFields |> List.toArray
        let values = values |> List.toArray
        let nf = rowFields.Length
        let measures (rows: CellValue[][]) =
            let result = Array.zeroCreate<CellValue> values.Length
            for i in 0 .. values.Length - 1 do
                let v, a = values.[i]
                result.[i] <- aggregateRows a rows v
            result
        let groupedRows (rows: CellValue[][]) field =
            let valuesByKey = stringKeyDict<CellValue>()
            let rowsByKey = stringKeyDict<ResizeArray<CellValue[]>>()
            for row in rows do
                let value = at row field
                let groupKey = key value
                addDistinct valuesByKey groupKey value
                (getOrAdd rowsByKey groupKey (fun () -> ResizeArray<CellValue[]>())).Add row
            let values = sortedValuesByKeyArray valuesByKey
            let result = Array.zeroCreate<CellValue * CellValue[][]> values.Length
            for i in 0 .. values.Length - 1 do
                let value = values.[i]
                let groupKey = key value
                result.[i] <- value, rowsByKey.[groupKey].ToArray()
            result
        // Build body rows recursively: leaf level emits a data row per value;
        // each non-leaf level emits its children then a "<value> Total" subtotal.
        let result = ResizeArray<CellValue[]>()
        let labels = Array.create nf CellValue.Empty
        let addRow prefixCount labelAtLevel (measureRows: CellValue[][]) =
            let row = Array.create (nf + values.Length) CellValue.Empty
            for i in 0 .. prefixCount - 1 do
                row.[i] <- labels.[i]
            match labelAtLevel with
            | Some level, value -> row.[level] <- value
            | None, _ -> ()
            let measureValues = measures measureRows
            System.Array.Copy(measureValues, 0, row, nf, measureValues.Length)
            result.Add row
        let rec build (level: int) (rows: CellValue[][]) : unit =
            let field = rowFields.[level]
            for kv, subset in groupedRows rows field do
                labels.[level] <- kv
                if level = nf - 1 then
                    addRow nf (None, CellValue.Empty) subset
                else
                    build (level + 1) subset
                    addRow level (Some level, CellValue.String(textOf kv + " Total")) subset
        let header = Array.create (nf + values.Length) CellValue.Empty
        for i in 0 .. nf - 1 do
            header.[i] <- hdr rowFields.[i]
        for i in 0 .. values.Length - 1 do
            let v, a = values.[i]
            header.[nf + i] <- CellValue.String(aggLabel a + " of " + textOf (hdr v))
        result.Add header
        build 0 data
        let grand = Array.create (nf + values.Length) CellValue.Empty
        grand.[0] <- CellValue.String "Grand Total"
        let grandMeasures = measures data
        System.Array.Copy(grandMeasures, 0, grand, nf, grandMeasures.Length)
        result.Add grand
        result.ToArray()

let computePivotN (source: CellValue[][]) (rowFields: int list) (values: (int * Agg) list) : CellValue[][] =
    if source.Length < 1 then [||]
    else
        let headers = source.[0]
        let data = if source.Length > 1 then source.[1..] else [||]
        computePivotNCore headers data rowFields values

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
        if r = 0 || v = 0 || n = 0 then [||]
        else
            let rowValue rowIndex field =
                if field >= 0 && field < rowM.[rowIndex].Length then rowM.[rowIndex].[field] else CellValue.Empty
            let measureValue rowIndex valueIndex =
                if valueIndex >= 0 && valueIndex < valueM.[rowIndex].Length then valueM.[rowIndex].[valueIndex] else CellValue.Empty
            let measures (indexes: ResizeArray<int>) =
                let result = Array.zeroCreate<CellValue> v
                for valueIndex = 0 to v - 1 do
                    let state = AggState()
                    for rowIndex in indexes do
                        addAggValue state (measureValue rowIndex valueIndex)
                    result.[valueIndex] <- finishAggState agg state
                result
            let groupedRows (indexes: ResizeArray<int>) field =
                let valuesByKey = stringKeyDict<CellValue>()
                let rowsByKey = stringKeyDict<ResizeArray<int>>()
                for rowIndex in indexes do
                    let value = rowValue rowIndex field
                    let groupKey = key value
                    addDistinct valuesByKey groupKey value
                    (getOrAdd rowsByKey groupKey (fun () -> ResizeArray<int>())).Add rowIndex
                let values = sortedValuesByKeyArray valuesByKey
                let result = Array.zeroCreate<CellValue * ResizeArray<int>> values.Length
                for i = 0 to values.Length - 1 do
                    let value = values.[i]
                    result.[i] <- value, rowsByKey.[key value]
                result
            let result = ResizeArray<CellValue[]>()
            let labels = Array.create r CellValue.Empty
            let addRow prefixCount labelAtLevel (measureRows: ResizeArray<int>) =
                let row = Array.create (r + v) CellValue.Empty
                for i = 0 to prefixCount - 1 do
                    row.[i] <- labels.[i]
                match labelAtLevel with
                | Some level, value -> row.[level] <- value
                | None, _ -> ()
                let measureValues = measures measureRows
                Array.Copy(measureValues, 0, row, r, measureValues.Length)
                result.Add row
            let rec build level indexes =
                for kv, subset in groupedRows indexes level do
                    labels.[level] <- kv
                    if level = r - 1 then
                        addRow r (None, CellValue.Empty) subset
                    else
                        build (level + 1) subset
                        addRow level (Some level, CellValue.String(textOf kv + " Total")) subset
            let allIndexes = ResizeArray<int>(n)
            for i = 0 to n - 1 do
                allIndexes.Add i
            build 0 allIndexes
            let grand = Array.create (r + v) CellValue.Empty
            grand.[0] <- CellValue.String "Grand Total"
            let grandMeasures = measures allIndexes
            Array.Copy(grandMeasures, 0, grand, r, grandMeasures.Length)
            result.Add grand
            result.ToArray()

/// PIVOTBY as a spilling function over data-only parallel ranges: single row
/// field × one or more (NESTED) column fields, one measure. `colM` may have
/// several columns — each becomes a column level, producing a multi-row column
/// header. Keeps per-row and grand totals.
let pivotBy (rowM: CellValue[][]) (colM: CellValue[][]) (valueM: CellValue[][]) (agg: Agg) : CellValue[][] =
    let n = min rowM.Length (min colM.Length valueM.Length)
    let levels = if colM.Length > 0 then colM.[0].Length else 0
    if n = 0 || levels = 0 then [||]
    else
        let rowOf i = if rowM.[i].Length > 0 then rowM.[i].[0] else CellValue.Empty
        let valOf i = if valueM.[i].Length > 0 then valueM.[i].[0] else CellValue.Empty
        let comboOf i = colM.[i]
        let rowValuesByKey = stringKeyDict<CellValue>()
        let comboValuesByKey = stringKeyDict<CellValue[]>()
        let rowTotals = stringKeyDict<AggState>()
        let comboTotals = stringKeyDict<AggState>()
        let cells = stringKeyDict<Dictionary<string, AggState>>()
        let grandTotal = AggState()
        for i in 0 .. n - 1 do
            let rowValue = rowOf i
            let rowKey = key rowValue
            let combo = comboOf i
            let colKey = comboKey combo
            addDistinct rowValuesByKey rowKey rowValue
            addDistinct comboValuesByKey colKey combo
            let value = valOf i
            addAggValue (getOrAdd rowTotals rowKey AggState) value
            addAggValue (getOrAdd comboTotals colKey AggState) value
            let rowCells = getOrAdd cells rowKey stringKeyDict<AggState>
            addAggValue (getOrAdd rowCells colKey AggState) value
            addAggValue grandTotal value
        let rowKeys = sortedValuesByKey rowValuesByKey
        let colCombos = sortedCombosByKeyArray comboValuesByKey levels
        let colComboKeys = Array.zeroCreate<string> colCombos.Length
        for i in 0 .. colCombos.Length - 1 do
            colComboKeys.[i] <- comboKey colCombos.[i]
        let result = ResizeArray<CellValue[]>(levels + rowKeys.Length + 1)
        for lvl in 0 .. levels - 1 do
            let header = Array.create (colCombos.Length + 2) CellValue.Empty
            header.[0] <- if lvl = 0 then CellValue.String "Row Labels" else CellValue.Empty
            for i in 0 .. colCombos.Length - 1 do
                header.[i + 1] <- colCombos.[i].[lvl]
            header.[colCombos.Length + 1] <- if lvl = 0 then CellValue.String "Grand Total" else CellValue.Empty
            result.Add header
        for rk in rowKeys do
            let rowKey = key rk
            let rowCells = tryFind cells rowKey
            let row = Array.create (colCombos.Length + 2) CellValue.Empty
            row.[0] <- rk
            for i in 0 .. colCombos.Length - 1 do
                let cellState =
                    match rowCells with
                    | Some rowDict -> tryFind rowDict colComboKeys.[i]
                    | None -> None
                row.[i + 1] <- finishAggOption agg cellState
            row.[colCombos.Length + 1] <- finishAggOption agg (tryFind rowTotals rowKey)
            result.Add row
        let totalRow = Array.create (colCombos.Length + 2) CellValue.Empty
        totalRow.[0] <- CellValue.String "Grand Total"
        for i in 0 .. colCombos.Length - 1 do
            totalRow.[i + 1] <- finishAggOption agg (tryFind comboTotals colComboKeys.[i])
        totalRow.[colCombos.Length + 1] <- finishAggState agg grandTotal
        result.Add totalRow
        result.ToArray()

/// Recompute a pivot from a parsed cache: `headers` are the cache field names,
/// `records` the resolved data rows (shared items already dereferenced), and
/// the field indices + aggregation come from the pivotTableDefinition. No
/// column fields → nested-row pivot (`computePivotN`); with column fields →
/// cross-tab (`pivotBy`, single row field). This is the import-recompute entry.
let fromCache (headers: string[]) (records: CellValue[][]) (rowIdxs: int list) (colIdxs: int list) (dataIdx: int) (agg: Agg) : CellValue[][] =
    if List.isEmpty rowIdxs || records.Length = 0 then [||]
    else
        let pick (r: CellValue[]) i = if i >= 0 && i < r.Length then r.[i] else CellValue.Empty
        match colIdxs with
        | [] ->
            computePivotNCore (headers |> Array.map CellValue.String) records rowIdxs [ (dataIdx, agg) ]
        | _ ->
            let rowIdxs = rowIdxs |> List.toArray
            let colIdxs = colIdxs |> List.toArray
            pivotByProjected records rowIdxs.[0] colIdxs dataIdx agg

/// Find a field's 0-based column by matching the header row (case-insensitive).
let fieldIndex (headers: CellValue[]) (name: string) : int option =
    let mutable found = -1
    let mutable i = 0
    while found < 0 && i < headers.Length do
        if trimmedEqualsIgnoreCase (textOf headers.[i]) name then
            found <- i
        i <- i + 1
    if found < 0 then None else Some found

let private trimNonEmptyFields (fields: string list) =
    let result = ResizeArray<string>()
    for field in fields do
        let trimmed = field.Trim()
        if trimmed <> "" then
            result.Add trimmed
    result

let private trimOptionalField (field: string option) =
    match field with
    | Some value ->
        let trimmed = value.Trim()
        if trimmed = "" then None else Some trimmed
    | None -> None

let private resolveFieldIndexes (headers: CellValue[]) (fields: ResizeArray<string>) offset =
    let indexes = ResizeArray<int>(fields.Count)
    let mutable missing = None
    for field in fields do
        match missing with
        | Some _ -> ()
        | None ->
            match fieldIndex headers field with
            | Some index -> indexes.Add(index + offset)
            | None -> missing <- Some field
    match missing with
    | Some field -> Result.Error field
    | None -> Result.Ok indexes

/// Build a LIVE pivot formula (`=GROUPBY(...)` / `=PIVOTBY(...)`) from the
/// dialog's field names + source bounds. The ranges are the *data* rows of
/// each field (header row excluded). This is what makes our pivot better than
/// Excel's: the dialog writes a formula the user owns — live, composable,
/// diffable — instead of a static object. `headers` is the source's header row
/// (length ec-sc+1); column indices are absolute (offset by `sc`).
let buildFormula (sr: int) (sc: int) (er: int) (_ec: int) (headers: CellValue[])
                 (rowFieldNames: string list) (colFieldName: string option)
                 (valueFieldName: string) (aggName: string) : Result<string, string> =
    let resolve name =
        match fieldIndex headers name with
        | Some li -> Some(sc + li)
        | None -> None
    let rowNames = trimNonEmptyFields rowFieldNames
    if rowNames.Count = 0 then Result.Error "Choose at least one row field."
    else
        match resolveFieldIndexes headers rowNames sc with
        | Result.Error missing -> Result.Error(sprintf "Row field not found: %s" missing)
        | Result.Ok cols ->
            match resolve valueFieldName with
            | None -> Result.Error(sprintf "Value field not found: %s" valueFieldName)
            | Some vi ->
                let dataTop = sr + 1
                let rng c0 c1 = sprintf "%s:%s" (formatCellRef dataTop c0) (formatCellRef er c1)
                let aggStr = "\"" + aggName.Trim().ToLowerInvariant() + "\""
                let mutable minCol = cols.[0]
                let mutable maxCol = cols.[0]
                for i = 1 to cols.Count - 1 do
                    let col = cols.[i]
                    if col < minCol then minCol <- col
                    if col > maxCol then maxCol <- col
                let rowRange = rng minCol maxCol
                match trimOptionalField colFieldName with
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
        let rowNames = trimNonEmptyFields rowFieldNames
        if rowNames.Count = 0 then Result.Error "Choose at least one row field."
        else
            match resolveFieldIndexes headers rowNames 0 with
            | Result.Error missing -> Result.Error(sprintf "Row field not found: %s" missing)
            | Result.Ok rows ->
                match fieldIndex headers valueFieldName with
                | None -> Result.Error(sprintf "Value field not found: %s" valueFieldName)
                | Some vi ->
                    match trimOptionalField colFieldName with
                    | Some cn ->
                        match fieldIndex headers cn with
                        | None -> Result.Error(sprintf "Column field not found: %s" cn)
                        | Some ci -> Result.Ok(computePivot source rows.[0] (Some ci) vi agg)
                    | None -> Result.Ok(computePivotN source (resizeArrayToList rows) [ (vi, agg) ])
