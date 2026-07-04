module RedBat.Spreadsheet.Vba

open Fable.Core
open Fable.Core.JsInterop
open Values
open Ast
open RedBat.Spreadsheet.Model

[<Emit("$0[$1]")>]
let private getProp (_source: obj) (_name: string): obj = jsNative

let private eq (left: string) (right: string) =
    if left.Length <> right.Length then
        false
    else
        let mutable index = 0
        let mutable equal = true
        while equal && index < left.Length do
            let l = left.[index]
            let r = right.[index]
            if l <> r && System.Char.ToUpperInvariant l <> System.Char.ToUpperInvariant r then
                equal <- false
            index <- index + 1
        equal

type private HostObject(typeName: string, getMember: string -> Value option, setMember: string -> Value -> bool, callMethod: string -> Value array -> Value option) =
    interface IHostObject with
        member _.TypeName = typeName
        member _.GetMember name = getMember name
        member _.SetMember name value = setMember name value
        member _.CallMethod name args = callMethod name args

let private host (typeName: string) (getMember: string -> Value option) (callMethod: string -> Value array -> Value option) =
    VHostObject (HostObject(typeName, getMember, (fun _ _ -> false), callMethod) :> IHostObject)

let private isDefaultMemberName (name: string) =
    name = "" || eq name "_Default" || eq name "Item"

let private valueToInt = function
    | VInteger i -> i
    | VDouble d -> int d
    | VBoolean b -> if b then 1 else 0
    | VString s -> match System.Int32.TryParse(s) with true, i -> i | _ -> 0
    | _ -> 0

let private valueToString = function
    | VString s -> s
    | VInteger i -> string i
    | VDouble d -> string d
    | VBoolean b -> if b then "TRUE" else "FALSE"
    | _ -> ""

let private valueToBool = function
    | VBoolean b -> b
    | VInteger i -> i <> 0
    | VDouble d -> d <> 0.0
    | VString s -> not (System.String.IsNullOrWhiteSpace s) && not (eq s "False")
    | _ -> false

let private jsValue (value: obj) =
    if isNull value then VEmpty
    else
        match value with
        | :? bool as b -> VBoolean b
        | :? int as i -> VInteger i
        | :? float as f ->
            if System.Double.IsNaN f then VEmpty
            elif abs (f - float (int f)) < 0.000001 then VInteger (int f)
            else VDouble f
        | :? string as s -> VString s
        | _ -> VString (string value)

let private propValue name (source: obj) =
    getProp source name |> jsValue

let private propFloat name (source: obj) =
    match propValue name source with
    | VInteger i -> float i
    | VDouble d -> d
    | VString s -> match System.Double.TryParse(s) with true, d -> d | _ -> 0.0
    | _ -> 0.0

let private propInt name (source: obj) =
    match propValue name source with
    | VInteger i -> i
    | VDouble d -> int d
    | VString s -> match System.Int32.TryParse(s) with true, i -> i | _ -> 0
    | _ -> 0

let private parseAddressSpan (address: string) startIndex endExclusive =
    if startIndex >= endExclusive then
        None
    else
        let mutable sheetEnd = -1
        let mutable sheetIndex = endExclusive - 1
        while sheetEnd < 0 && sheetIndex >= startIndex do
            if address.[sheetIndex] = '!' then
                sheetEnd <- sheetIndex
            sheetIndex <- sheetIndex - 1
        let mutable index = if sheetEnd >= startIndex then sheetEnd + 1 else startIndex
        let mutable col = 0
        let mutable row = 0
        let mutable hasCol = false
        let mutable hasRow = false
        let mutable inRow = false
        let mutable valid = true
        while valid && index < endExclusive do
            let c = address.[index]
            if c = '$' then
                ()
            elif System.Char.IsDigit c then
                inRow <- true
                hasRow <- true
                row <- row * 10 + int c - int '0'
            elif System.Char.IsLetter c && not inRow then
                hasCol <- true
                let upper = System.Char.ToUpperInvariant c
                col <- col * 26 + int upper - int 'A' + 1
            else
                valid <- false
            index <- index + 1
        if valid && hasCol && hasRow && row > 0 && col > 0 then
            Some (row - 1, col - 1)
        else
            None

let private parseRangeRef (reference: string) =
    let colon = reference.IndexOf(':')
    if colon < 0 then
        parseAddressSpan reference 0 reference.Length
        |> Option.map (fun (r, c) -> r, c, r, c)
    elif colon = reference.LastIndexOf(':') then
        match parseAddressSpan reference 0 colon, parseAddressSpan reference (colon + 1) reference.Length with
        | Some (r1, c1), Some (r2, c2) -> Some (min r1 r2, min c1 c2, max r1 r2, max c1 c2)
        | _ -> None
    else
        None

let private formatCellAddress row col rowAbsolute columnAbsolute =
    sprintf "%s%s%s%d"
        (if columnAbsolute then "$" else "")
        (columnToLetter (col - 1))
        (if rowAbsolute then "$" else "")
        row

type CellSnapshot = {
    Row: int
    Column: int
    Address: string
    X: float
    Y: float
    Width: float
    Height: float
    MergedRows: int
    MergedCols: int
    RowHeight: float
    ColumnWidth: float
    WrapText: bool
    HAlign: string
    VAlign: string
    FontName: string
    FontSize: float
    Bold: bool
    Italic: bool
    NumberFormat: string
    Formula: string
    Value: string
    Text: string
}

type private SnapshotCells = {
    All: unit -> CellSnapshot array
    TryCell: int -> int -> CellSnapshot option
}

let private pxToPt (value: float) = value * 72.0 / 96.0

let private cellKey row col =
    ((int64 row) <<< 32) ||| int64 col

let private cellValueRepr = function
    | CellValue.Empty -> ""
    | CellValue.Number n -> string n
    | CellValue.String s -> s
    | CellValue.Boolean b -> if b then "TRUE" else "FALSE"
    | CellValue.Error e -> "#" + e
    | CellValue.Formula (expr, _) -> "=" + expr

let private cellFormulaRepr = function
    | CellValue.Formula (expr, _) -> "=" + expr
    | _ -> ""

let private modelSnapshotBounds (model: SpreadsheetModel) =
    match usedBounds model with
    | Some (row, col) -> min (model.RowCount - 1) row, min (model.ColCount - 1) col
    | None -> 0, 0

let private findMergeAt row col (model: SpreadsheetModel) =
    let mutable result = ValueNone
    let mutable remaining = model.MergedCells
    while result.IsNone && not remaining.IsEmpty do
        let merge = remaining.Head
        if row >= merge.StartRow && row <= merge.EndRow &&
           col >= merge.StartCol && col <= merge.EndCol then
            result <- ValueSome merge
        remaining <- remaining.Tail
    result

let private emitModelSnapshotCell (model: SpreadsheetModel) (viewport: RedBat.Spreadsheet.Layout.Viewport) row col merge =
    let addr = { Row = row; Col = col }
    let rect =
        match merge with
        | ValueSome m -> RedBat.Spreadsheet.Layout.mergedRect model viewport m
        | ValueNone -> RedBat.Spreadsheet.Layout.cellRect model viewport row col
    let cellData = getCell addr model
    let value = getCellValue addr model
    let format = cellData |> Option.map (fun c -> c.Format) |> Option.defaultValue defaultFormat
    let fontPt = format.FontSize |> Option.map float |> Option.defaultValue 11.0
    let displayText = RedBat.Spreadsheet.CellText.cellDisplayText None rect.Width format value
    let mergedRows, mergedCols =
        match merge with
        | ValueSome m -> m.EndRow - m.StartRow + 1, m.EndCol - m.StartCol + 1
        | ValueNone -> 1, 1
    {
        Row = row + 1
        Column = col + 1
        Address = columnToLetter col + string (row + 1)
        X = pxToPt (rect.X - viewport.RowHeaderWidth)
        Y = pxToPt (rect.Y - viewport.HeaderHeight)
        Width = pxToPt rect.Width
        Height = pxToPt rect.Height
        MergedRows = mergedRows
        MergedCols = mergedCols
        RowHeight = pxToPt (getRowHeight row model)
        ColumnWidth = pxToPt (getColumnWidth col model)
        WrapText = format.WrapText
        HAlign = format.HorizontalAlign |> Option.defaultValue "general"
        VAlign = format.VerticalAlign |> Option.defaultValue "bottom"
        FontName = format.FontName |> Option.defaultValue "Calibri"
        FontSize = fontPt
        Bold = format.Bold
        Italic = format.Italic
        NumberFormat = format.NumberFormat |> Option.defaultValue "General"
        Formula = cellFormulaRepr value
        Value = cellValueRepr value
        Text = displayText
    }

let private snapshotFromModelWithViewport (model: SpreadsheetModel) (viewport: RedBat.Spreadsheet.Layout.Viewport) =
    let maxRow, maxCol = modelSnapshotBounds model

    if List.isEmpty model.MergedCells then
        [|
            for row in 0 .. maxRow do
                for col in 0 .. maxCol do
                    yield emitModelSnapshotCell model viewport row col ValueNone
        |]
    else
        let mergedByCell = System.Collections.Generic.Dictionary<int64, MergedCell>()
        for merge in model.MergedCells do
            let r0 = max 0 merge.StartRow
            let c0 = max 0 merge.StartCol
            let r1 = min maxRow merge.EndRow
            let c1 = min maxCol merge.EndCol
            for row in r0 .. r1 do
                for col in c0 .. c1 do
                    mergedByCell.[cellKey row col] <- merge

        [|
            for row in 0 .. maxRow do
                for col in 0 .. maxCol do
                    match mergedByCell.TryGetValue(cellKey row col) with
                    | true, m when row = m.StartRow && col = m.StartCol ->
                        yield emitModelSnapshotCell model viewport row col (ValueSome m)
                    | true, _ ->
                        ()
                    | false, _ ->
                        yield emitModelSnapshotCell model viewport row col ValueNone
    |]

let snapshotFromModel (model: SpreadsheetModel) =
    snapshotFromModelWithViewport model (RedBat.Spreadsheet.Layout.defaultViewport 100000.0 100000.0)

let private snapshotCellsFromModel (model: SpreadsheetModel) =
    let viewport = RedBat.Spreadsheet.Layout.defaultViewport 100000.0 100000.0
    let all = lazy (snapshotFromModelWithViewport model viewport)
    let maxRow, maxCol = modelSnapshotBounds model
    {
        All = fun () -> all.Force()
        TryCell =
            fun row col ->
                if row < 0 || col < 0 || row > maxRow || col > maxCol then
                    None
                else
                    match findMergeAt row col model with
                    | ValueSome merge when row = merge.StartRow && col = merge.StartCol ->
                        Some (emitModelSnapshotCell model viewport row col (ValueSome merge))
                    | ValueSome _ ->
                        None
                    | ValueNone ->
                        Some (emitModelSnapshotCell model viewport row col ValueNone)
    }

let private snapshot (spreadsheet: Spreadsheet.Spreadsheet) =
    spreadsheet.GetVbaLayoutCells()
    |> Array.map box

let private cellsInRange (spreadsheet: Spreadsheet.Spreadsheet) startRow startCol endRow endCol =
    snapshot spreadsheet
    |> Array.filter (fun cell ->
        let row = propInt "row" cell
        let col = propInt "column" cell
        row >= startRow + 1 && row <= endRow + 1 && col >= startCol + 1 && col <= endCol + 1)

let rec private rangeObject (spreadsheet: Spreadsheet.Spreadsheet) startRow startCol endRow endCol =
    let cellsCache = lazy (cellsInRange spreadsheet startRow startCol endRow endCol)
    let cells () = cellsCache.Force()
    let firstCellCache = lazy (cells () |> Array.tryHead)
    let firstCell () = firstCellCache.Force()

    let boundsCache = lazy (
        let cells = cells ()
        if cells.Length = 0 then 0.0, 0.0, 0.0, 0.0
        else
            let first = cells.[0]
            let mutable left = propFloat "x" first
            let mutable top = propFloat "y" first
            let mutable right = left + propFloat "width" first
            let mutable bottom = top + propFloat "height" first
            for i = 1 to cells.Length - 1 do
                let cell = cells.[i]
                let x = propFloat "x" cell
                let y = propFloat "y" cell
                let r = x + propFloat "width" cell
                let b = y + propFloat "height" cell
                if x < left then left <- x
                if y < top then top <- y
                if r > right then right <- r
                if b > bottom then bottom <- b
            left, top, right - left, bottom - top
    )
    let aggregateBounds () = boundsCache.Force()

    let formatAddress rowAbsolute columnAbsolute =
        let first = formatCellAddress (startRow + 1) (startCol + 1) rowAbsolute columnAbsolute
        let last = formatCellAddress (endRow + 1) (endCol + 1) rowAbsolute columnAbsolute
        if startRow = endRow && startCol = endCol then
            first
        else
            sprintf "%s:%s" first last

    let address = formatAddress true true

    let getMember (name: string) =
        if eq name "Address" then Some (VString address)
        elif eq name "Row" then Some (VInteger (startRow + 1))
        elif eq name "Column" then Some (VInteger (startCol + 1))
        elif eq name "Count" then Some (VInteger ((endRow - startRow + 1) * (endCol - startCol + 1)))
        elif eq name "Left" then let l, _, _, _ = aggregateBounds () in Some (VDouble l)
        elif eq name "Top" then let _, t, _, _ = aggregateBounds () in Some (VDouble t)
        elif eq name "Width" then let _, _, w, _ = aggregateBounds () in Some (VDouble w)
        elif eq name "Height" then let _, _, _, h = aggregateBounds () in Some (VDouble h)
        elif eq name "RowHeight" then
            firstCell () |> Option.map (propValue "rowHeight") |> Option.orElse (Some VEmpty)
        elif eq name "ColumnWidth" then
            firstCell () |> Option.map (propValue "columnWidth") |> Option.orElse (Some VEmpty)
        elif eq name "WrapText" then
            firstCell () |> Option.map (propValue "wrapText") |> Option.orElse (Some VEmpty)
        elif eq name "HorizontalAlignment" then
            firstCell () |> Option.map (propValue "hAlign") |> Option.orElse (Some VEmpty)
        elif eq name "VerticalAlignment" then
            firstCell () |> Option.map (propValue "vAlign") |> Option.orElse (Some VEmpty)
        elif eq name "Value" || eq name "Value2" then
            firstCell () |> Option.map (propValue "value") |> Option.orElse (Some VEmpty)
        elif eq name "Text" then
            firstCell () |> Option.map (propValue "text") |> Option.orElse (Some VEmpty)
        elif eq name "Formula" then
            firstCell () |> Option.map (propValue "formula") |> Option.orElse (Some VEmpty)
        elif eq name "NumberFormat" then
            firstCell () |> Option.map (propValue "numberFormat") |> Option.orElse (Some VEmpty)
        elif eq name "MergeCells" then
            firstCell ()
            |> Option.map (fun c -> VBoolean (propInt "mergedRows" c > 1 || propInt "mergedCols" c > 1))
            |> Option.orElse (Some (VBoolean false))
        elif eq name "MergeArea" then Some (rangeObject spreadsheet startRow startCol endRow endCol)
        elif eq name "Rows" then
            Some (host "Excel.Rows"
                (fun (field: string) -> if eq field "Count" then Some (VInteger (endRow - startRow + 1)) else None)
                (fun (_: string) (_: Value array) -> None))
        elif eq name "Columns" then
            Some (host "Excel.Columns"
                (fun (field: string) -> if eq field "Count" then Some (VInteger (endCol - startCol + 1)) else None)
                (fun (_: string) (_: Value array) -> None))
        elif eq name "Cells" then
            Some (host "Excel.Cells"
                (fun (field: string) -> if eq field "Count" then Some (VInteger ((endRow - startRow + 1) * (endCol - startCol + 1))) else None)
                (fun (methodName: string) (args: Value array) ->
                    if eq methodName "Item" && args.Length >= 2 then
                        let row = startRow + valueToInt args.[0] - 1
                        let col = startCol + valueToInt args.[1] - 1
                        Some (rangeObject spreadsheet row col row col)
                    else None))
        else None

    let callMethod (name: string) (args: Value array) =
        if eq name "Address" then
            let rowAbsolute = if args.Length >= 1 then valueToBool args.[0] else true
            let columnAbsolute = if args.Length >= 2 then valueToBool args.[1] else true
            Some (VString (formatAddress rowAbsolute columnAbsolute))
        elif eq name "Cells" && args.Length >= 2 then
            let row = startRow + valueToInt args.[0] - 1
            let col = startCol + valueToInt args.[1] - 1
            Some (rangeObject spreadsheet row col row col)
        elif eq name "Offset" then
            let rowOffset = if args.Length >= 1 then valueToInt args.[0] else 0
            let colOffset = if args.Length >= 2 then valueToInt args.[1] else 0
            Some (rangeObject spreadsheet (startRow + rowOffset) (startCol + colOffset) (endRow + rowOffset) (endCol + colOffset))
        elif eq name "Resize" then
            let rows = if args.Length >= 1 then max 1 (valueToInt args.[0]) else endRow - startRow + 1
            let cols = if args.Length >= 2 then max 1 (valueToInt args.[1]) else endCol - startCol + 1
            Some (rangeObject spreadsheet startRow startCol (startRow + rows - 1) (startCol + cols - 1))
        else None

    host "Excel.Range" getMember callMethod

let private worksheetObject (spreadsheet: Spreadsheet.Spreadsheet) =
    host "Excel.Worksheet"
        (fun (name: string) ->
            if eq name "Name" then
                spreadsheet.SheetNames()
                |> Array.tryItem spreadsheet.ActiveSheetIndex
                |> Option.map VString
            elif eq name "UsedRange" then
                let cells = snapshot spreadsheet
                if cells.Length = 0 then
                    Some (rangeObject spreadsheet 0 0 0 0)
                else
                    let first = cells.[0]
                    let mutable minRow = propInt "row" first
                    let mutable maxRow = minRow
                    let mutable minCol = propInt "column" first
                    let mutable maxCol = minCol
                    for i = 1 to cells.Length - 1 do
                        let cell = cells.[i]
                        let row = propInt "row" cell
                        let col = propInt "column" cell
                        if row < minRow then minRow <- row
                        if row > maxRow then maxRow <- row
                        if col < minCol then minCol <- col
                        if col > maxCol then maxCol <- col
                    Some (rangeObject spreadsheet (minRow - 1) (minCol - 1) (maxRow - 1) (maxCol - 1))
            elif eq name "Cells" then
                Some (rangeObject spreadsheet 0 0 1048575 16383)
            else None)
        (fun (name: string) (args: Value array) ->
            if eq name "Range" && args.Length >= 1 then
                match parseRangeRef (valueToString args.[0]) with
                | Some (r1, c1, r2, c2) -> Some (rangeObject spreadsheet r1 c1 r2 c2)
                | None -> Some VEmpty
            elif eq name "Cells" && args.Length >= 2 then
                let row = valueToInt args.[0]
                let col = valueToInt args.[1]
                Some (rangeObject spreadsheet (row - 1) (col - 1) (row - 1) (col - 1))
            else None)

let private workbookObject (spreadsheet: Spreadsheet.Spreadsheet) =
    host "Excel.Workbook"
        (fun (name: string) ->
            if eq name "ActiveSheet" then Some (worksheetObject spreadsheet)
            else None)
        (fun (_: string) (_: Value array) -> None)

let private applicationObject (spreadsheet: Spreadsheet.Spreadsheet) =
    host "Excel.Application"
        (fun (name: string) ->
            if eq name "ActiveWorkbook" then Some (workbookObject spreadsheet)
            elif eq name "ActiveSheet" then Some (worksheetObject spreadsheet)
            elif eq name "Version" then Some (VString "15.0")
            else None)
        (fun (name: string) (args: Value array) ->
            if eq name "Range" && args.Length >= 1 then
                match parseRangeRef (valueToString args.[0]) with
                | Some (r1, c1, r2, c2) -> Some (rangeObject spreadsheet r1 c1 r2 c2)
                | None -> Some VEmpty
            elif eq name "Cells" && args.Length >= 2 then
                let row = valueToInt args.[0]
                let col = valueToInt args.[1]
                Some (rangeObject spreadsheet (row - 1) (col - 1) (row - 1) (col - 1))
            else None)

let private rangeFunction (spreadsheet: Spreadsheet.Spreadsheet) =
    host "Excel.RangeFunction"
        (fun (_: string) -> None)
        (fun (name: string) (args: Value array) ->
            if isDefaultMemberName name && args.Length >= 1 then
                match parseRangeRef (valueToString args.[0]) with
                | Some (r1, c1, r2, c2) -> Some (rangeObject spreadsheet r1 c1 r2 c2)
                | None -> Some VEmpty
            else None)

let private cellsFunction (spreadsheet: Spreadsheet.Spreadsheet) =
    host "Excel.CellsFunction"
        (fun (_: string) -> None)
        (fun (name: string) (args: Value array) ->
            if isDefaultMemberName name && args.Length >= 2 then
                let row = valueToInt args.[0]
                let col = valueToInt args.[1]
                Some (rangeObject spreadsheet (row - 1) (col - 1) (row - 1) (col - 1))
            else None)

let private createObjectFunction (spreadsheet: Spreadsheet.Spreadsheet) =
    host "VBScript.CreateObject"
        (fun (_: string) -> None)
        (fun (name: string) (args: Value array) ->
            if isDefaultMemberName name && args.Length >= 1 then
                let progId = valueToString args.[0]
                if eq progId "Excel.Application" then Some (applicationObject spreadsheet)
                else Some VNothing
            else None)

let hostGlobals (spreadsheet: Spreadsheet.Spreadsheet) =
    Map.ofList [
        "Application", applicationObject spreadsheet
        "ActiveWorkbook", workbookObject spreadsheet
        "ActiveSheet", worksheetObject spreadsheet
        "Range", rangeFunction spreadsheet
        "Cells", cellsFunction spreadsheet
    ]

let hostGlobalsVbs (spreadsheet: Spreadsheet.Spreadsheet) =
    hostGlobals spreadsheet
    |> Map.add "CreateObject" (createObjectFunction spreadsheet)

let run (spreadsheet: Spreadsheet.Spreadsheet) (source: string) =
    let options =
        FsVb.RunOptions.vba
        |> FsVb.RunOptions.withHostGlobals (hostGlobals spreadsheet)
    FsVb.runSource options source

let runVbs (spreadsheet: Spreadsheet.Spreadsheet) (source: string) =
    let options =
        FsVb.RunOptions.vbScript
        |> FsVb.RunOptions.withHostGlobals (hostGlobalsVbs spreadsheet)
    FsVb.runSource options source

let private snapshotCells (snapshot: unit -> CellSnapshot array) =
    let mutable allCache: CellSnapshot array voption = ValueNone
    let all () =
        match allCache with
        | ValueSome cells -> cells
        | ValueNone ->
            let cells = snapshot ()
            allCache <- ValueSome cells
            cells
    let buildIndex () =
        let index = System.Collections.Generic.Dictionary<int64, CellSnapshot>()
        for cell in all () do
            index.[cellKey (cell.Row - 1) (cell.Column - 1)] <- cell
        index
    let mutable lookupCount = 0
    let mutable byCell: System.Collections.Generic.Dictionary<int64, CellSnapshot> voption = ValueNone
    let linearFind row col =
        let cells = all ()
        let row1 = row + 1
        let col1 = col + 1
        let mutable found = ValueNone
        let mutable index = 0
        while found.IsNone && index < cells.Length do
            let cell = cells.[index]
            if cell.Row = row1 && cell.Column = col1 then
                found <- ValueSome cell
            index <- index + 1
        match found with
        | ValueSome cell -> Some cell
        | ValueNone -> None
    {
        All = all
        TryCell =
            fun row col ->
                match byCell with
                | ValueSome index ->
                    match index.TryGetValue(cellKey row col) with
                    | true, cell -> Some cell
                    | false, _ -> None
                | ValueNone ->
                    lookupCount <- lookupCount + 1
                    if lookupCount >= 8 && (all ()).Length >= 32 then
                        let index = buildIndex ()
                        byCell <- ValueSome index
                        match index.TryGetValue(cellKey row col) with
                        | true, cell -> Some cell
                        | false, _ -> None
                    else
                        linearFind row col
    }

let private cellsInSnapshot (snapshot: SnapshotCells) startRow startCol endRow endCol =
    if startRow = endRow && startCol = endCol then
        match snapshot.TryCell startRow startCol with
        | Some cell -> [| cell |]
        | None -> [||]
    else
        let cells = snapshot.All ()
        let startRow1 = startRow + 1
        let startCol1 = startCol + 1
        let endRow1 = endRow + 1
        let endCol1 = endCol + 1
        let result = ResizeArray<CellSnapshot>()
        for cell in cells do
            if cell.Row >= startRow1 && cell.Row <= endRow1 &&
               cell.Column >= startCol1 && cell.Column <= endCol1 then
                result.Add cell
        result.ToArray()

let rec private rangeObjectForSnapshot (sheetName: string) (snapshot: SnapshotCells) startRow startCol endRow endCol =
    let isSingleCell = startRow = endRow && startCol = endCol
    let cellsCache = lazy (cellsInSnapshot snapshot startRow startCol endRow endCol)
    let cells () = cellsCache.Force()
    let firstCellCache = lazy (
        if isSingleCell then snapshot.TryCell startRow startCol
        else cells () |> Array.tryHead)
    let firstCell () = firstCellCache.Force()

    let boundsCache = lazy (
        if isSingleCell then
            match firstCell () with
            | Some cell -> cell.X, cell.Y, cell.Width, cell.Height
            | None -> 0.0, 0.0, 0.0, 0.0
        else
            let cells = cells ()
            if cells.Length = 0 then 0.0, 0.0, 0.0, 0.0
            else
                let first = cells.[0]
                let mutable left = first.X
                let mutable top = first.Y
                let mutable right = first.X + first.Width
                let mutable bottom = first.Y + first.Height
                for i = 1 to cells.Length - 1 do
                    let cell = cells.[i]
                    let r = cell.X + cell.Width
                    let b = cell.Y + cell.Height
                    if cell.X < left then left <- cell.X
                    if cell.Y < top then top <- cell.Y
                    if r > right then right <- r
                    if b > bottom then bottom <- b
                left, top, right - left, bottom - top
    )
    let aggregateBounds () = boundsCache.Force()

    let formatAddress rowAbsolute columnAbsolute =
        let first = formatCellAddress (startRow + 1) (startCol + 1) rowAbsolute columnAbsolute
        let last = formatCellAddress (endRow + 1) (endCol + 1) rowAbsolute columnAbsolute
        if startRow = endRow && startCol = endCol then first else sprintf "%s:%s" first last

    let address = formatAddress true true

    let rowsObject () =
        host "Excel.Rows"
            (fun (field: string) ->
                if eq field "Count" then Some (VInteger (endRow - startRow + 1))
                else None)
            (fun (_: string) (_: Value array) -> None)

    let columnsObject () =
        host "Excel.Columns"
            (fun (field: string) ->
                if eq field "Count" then Some (VInteger (endCol - startCol + 1))
                else None)
            (fun (_: string) (_: Value array) -> None)

    let cellsObject () =
        host "Excel.Cells"
            (fun (field: string) ->
                if eq field "Count" then Some (VInteger ((endRow - startRow + 1) * (endCol - startCol + 1)))
                else None)
            (fun (methodName: string) (args: Value array) ->
                if eq methodName "Item" && args.Length >= 2 then
                    let row = startRow + valueToInt args.[0] - 1
                    let col = startCol + valueToInt args.[1] - 1
                    Some (rangeObjectForSnapshot sheetName snapshot row col row col)
                else None)

    host "Excel.Range"
        (fun (name: string) ->
            if eq name "Address" then Some (VString address)
            elif eq name "Worksheet" then Some (worksheetObjectForSnapshot sheetName snapshot)
            elif eq name "Row" then Some (VInteger (startRow + 1))
            elif eq name "Column" then Some (VInteger (startCol + 1))
            elif eq name "Count" then Some (VInteger ((endRow - startRow + 1) * (endCol - startCol + 1)))
            elif eq name "Left" then let l, _, _, _ = aggregateBounds () in Some (VDouble l)
            elif eq name "Top" then let _, t, _, _ = aggregateBounds () in Some (VDouble t)
            elif eq name "Width" then let _, _, w, _ = aggregateBounds () in Some (VDouble w)
            elif eq name "Height" then let _, _, _, h = aggregateBounds () in Some (VDouble h)
            elif eq name "RowHeight" then firstCell () |> Option.map (fun c -> VDouble c.RowHeight) |> Option.orElse (Some VEmpty)
            elif eq name "ColumnWidth" then firstCell () |> Option.map (fun c -> VDouble c.ColumnWidth) |> Option.orElse (Some VEmpty)
            elif eq name "WrapText" then firstCell () |> Option.map (fun c -> VBoolean c.WrapText) |> Option.orElse (Some VEmpty)
            elif eq name "HorizontalAlignment" then firstCell () |> Option.map (fun c -> VString c.HAlign) |> Option.orElse (Some VEmpty)
            elif eq name "VerticalAlignment" then firstCell () |> Option.map (fun c -> VString c.VAlign) |> Option.orElse (Some VEmpty)
            elif eq name "Value" || eq name "Value2" then firstCell () |> Option.map (fun c -> VString c.Value) |> Option.orElse (Some VEmpty)
            elif eq name "Text" then firstCell () |> Option.map (fun c -> VString c.Text) |> Option.orElse (Some VEmpty)
            elif eq name "Formula" then firstCell () |> Option.map (fun c -> VString c.Formula) |> Option.orElse (Some VEmpty)
            elif eq name "NumberFormat" then firstCell () |> Option.map (fun c -> VString c.NumberFormat) |> Option.orElse (Some VEmpty)
            elif eq name "MergeCells" then firstCell () |> Option.map (fun c -> VBoolean (c.MergedRows > 1 || c.MergedCols > 1)) |> Option.orElse (Some (VBoolean false))
            elif eq name "MergeArea" then Some (rangeObjectForSnapshot sheetName snapshot startRow startCol endRow endCol)
            elif eq name "Rows" then Some (rowsObject ())
            elif eq name "Columns" then Some (columnsObject ())
            elif eq name "Cells" then Some (cellsObject ())
            else None)
        (fun (name: string) (args: Value array) ->
            if eq name "Address" then
                let rowAbsolute = if args.Length >= 1 then valueToBool args.[0] else true
                let columnAbsolute = if args.Length >= 2 then valueToBool args.[1] else true
                Some (VString (formatAddress rowAbsolute columnAbsolute))
            elif eq name "Cells" && args.Length >= 2 then
                let row = startRow + valueToInt args.[0] - 1
                let col = startCol + valueToInt args.[1] - 1
                Some (rangeObjectForSnapshot sheetName snapshot row col row col)
            elif eq name "Offset" then
                let rowOffset = if args.Length >= 1 then valueToInt args.[0] else 0
                let colOffset = if args.Length >= 2 then valueToInt args.[1] else 0
                Some (rangeObjectForSnapshot sheetName snapshot (startRow + rowOffset) (startCol + colOffset) (endRow + rowOffset) (endCol + colOffset))
            elif eq name "Resize" then
                let rows = if args.Length >= 1 then max 1 (valueToInt args.[0]) else endRow - startRow + 1
                let cols = if args.Length >= 2 then max 1 (valueToInt args.[1]) else endCol - startCol + 1
                Some (rangeObjectForSnapshot sheetName snapshot startRow startCol (startRow + rows - 1) (startCol + cols - 1))
            else None)

and private usedRangeObjectForSnapshot (sheetName: string) (snapshot: SnapshotCells) =
    let cells = snapshot.All ()
    if cells.Length = 0 then
        rangeObjectForSnapshot sheetName snapshot 0 0 0 0
    else
        let first = cells.[0]
        let mutable minRow = first.Row
        let mutable maxRow = first.Row
        let mutable minCol = first.Column
        let mutable maxCol = first.Column
        for i = 1 to cells.Length - 1 do
            let cell = cells.[i]
            if cell.Row < minRow then minRow <- cell.Row
            if cell.Row > maxRow then maxRow <- cell.Row
            if cell.Column < minCol then minCol <- cell.Column
            if cell.Column > maxCol then maxCol <- cell.Column
        rangeObjectForSnapshot sheetName snapshot (minRow - 1) (minCol - 1) (maxRow - 1) (maxCol - 1)

and private worksheetObjectForSnapshot (sheetName: string) (snapshot: SnapshotCells) =
    host "Excel.Worksheet"
        (fun (name: string) ->
            if eq name "Name" then Some (VString sheetName)
            elif eq name "UsedRange" then Some (usedRangeObjectForSnapshot sheetName snapshot)
            elif eq name "Cells" then Some (rangeObjectForSnapshot sheetName snapshot 0 0 1048575 16383)
            else None)
        (fun (name: string) (args: Value array) ->
            if eq name "Range" && args.Length >= 1 then
                match parseRangeRef (valueToString args.[0]) with
                | Some (r1, c1, r2, c2) -> Some (rangeObjectForSnapshot sheetName snapshot r1 c1 r2 c2)
                | None -> Some VEmpty
            elif eq name "Cells" && args.Length >= 2 then
                let row = valueToInt args.[0]
                let col = valueToInt args.[1]
                Some (rangeObjectForSnapshot sheetName snapshot (row - 1) (col - 1) (row - 1) (col - 1))
            else None)

let private workbookObjectForSnapshot sheetName snapshot =
    host "Excel.Workbook"
        (fun (name: string) ->
            if eq name "ActiveSheet" then Some (worksheetObjectForSnapshot sheetName snapshot)
            else None)
        (fun (_: string) (_: Value array) -> None)

let private applicationObjectForSnapshot sheetName snapshot =
    host "Excel.Application"
        (fun (name: string) ->
            if eq name "ActiveWorkbook" then Some (workbookObjectForSnapshot sheetName snapshot)
            elif eq name "ActiveSheet" then Some (worksheetObjectForSnapshot sheetName snapshot)
            elif eq name "Version" then Some (VString "15.0")
            else None)
        (fun (name: string) (args: Value array) ->
            if eq name "Range" && args.Length >= 1 then
                match parseRangeRef (valueToString args.[0]) with
                | Some (r1, c1, r2, c2) -> Some (rangeObjectForSnapshot sheetName snapshot r1 c1 r2 c2)
                | None -> Some VEmpty
            elif eq name "Cells" && args.Length >= 2 then
                let row = valueToInt args.[0]
                let col = valueToInt args.[1]
                Some (rangeObjectForSnapshot sheetName snapshot (row - 1) (col - 1) (row - 1) (col - 1))
            else None)

let private rangeFunctionForSnapshot sheetName snapshot =
    host "Excel.RangeFunction"
        (fun (_: string) -> None)
        (fun (name: string) (args: Value array) ->
            if isDefaultMemberName name && args.Length >= 1 then
                match parseRangeRef (valueToString args.[0]) with
                | Some (r1, c1, r2, c2) -> Some (rangeObjectForSnapshot sheetName snapshot r1 c1 r2 c2)
                | None -> Some VEmpty
            else None)

let private cellsFunctionForSnapshot sheetName snapshot =
    host "Excel.CellsFunction"
        (fun (_: string) -> None)
        (fun (name: string) (args: Value array) ->
            if isDefaultMemberName name && args.Length >= 2 then
                let row = valueToInt args.[0]
                let col = valueToInt args.[1]
                Some (rangeObjectForSnapshot sheetName snapshot (row - 1) (col - 1) (row - 1) (col - 1))
            else None)

let private createObjectFunctionForSnapshot sheetName snapshot =
    host "VBScript.CreateObject"
        (fun (_: string) -> None)
        (fun (name: string) (args: Value array) ->
            if isDefaultMemberName name && args.Length >= 1 then
                let progId = valueToString args.[0]
                if eq progId "Excel.Application" then Some (applicationObjectForSnapshot sheetName snapshot)
                else Some VNothing
            else None)

let private hostGlobalsForSnapshotCells sheetName snapshot =
    Map.ofList [
        "Application", applicationObjectForSnapshot sheetName snapshot
        "ActiveWorkbook", workbookObjectForSnapshot sheetName snapshot
        "ActiveSheet", worksheetObjectForSnapshot sheetName snapshot
        "Range", rangeFunctionForSnapshot sheetName snapshot
        "Cells", cellsFunctionForSnapshot sheetName snapshot
    ]

let hostGlobalsForSnapshot sheetName snapshot =
    snapshotCells snapshot
    |> hostGlobalsForSnapshotCells sheetName

let private hostGlobalsForSnapshotCellsVbs sheetName source =
    hostGlobalsForSnapshotCells sheetName source
    |> Map.add "CreateObject" (createObjectFunctionForSnapshot sheetName source)

let hostGlobalsForSnapshotVbs sheetName snapshot =
    snapshotCells snapshot
    |> hostGlobalsForSnapshotCellsVbs sheetName

// ---- Writable model-backed host (cross-platform: native + web) -------------
// `runOnModel` runs a macro with READ+WRITE Range/Cells/Worksheet/Application
// over a SpreadsheetModel. Reads fall through to the model; writes are tracked
// and folded back after the run, then recomputed. Unlike the read-only
// snapshot host this lets a macro change cells — and it needs no DOM, so it
// works in the native editor as well as the web shell. Covers the common
// surface: Range/Cells .Value/.Value2/.Formula (get + set), .Address/.Row/
// .Column/.Count/.Cells/.Offset/.Resize, Worksheet.Range/.Cells/.Name.

let private hostRW typeName getMember setMember callMethod =
    VHostObject (HostObject(typeName, getMember, setMember, callMethod) :> IHostObject)

/// Mutable cell view over a model: reads prefer pending writes, then the model;
/// writes accumulate (0-indexed row/col → raw value string).
type private ModelCells(model: SpreadsheetModel) =
    let writes = System.Collections.Generic.Dictionary<struct (int * int), string>()
    member _.Read (r: int) (c: int) : string =
        match writes.TryGetValue(struct (r, c)) with
        | true, v -> v
        | _ -> displayValue (getCellValue { Row = r; Col = c } model)
    member _.Write (r: int) (c: int) (v: string) = writes.[struct (r, c)] <- v
    member _.Writes = writes

let rec private rangeRW (cells: ModelCells) (sheetName: string) r0 c0 r1 c1 =
    let formatAddr ra ca =
        let f = formatCellAddress (r0 + 1) (c0 + 1) ra ca
        let l = formatCellAddress (r1 + 1) (c1 + 1) ra ca
        if r0 = r1 && c0 = c1 then f else sprintf "%s:%s" f l
    let setAll (v: string) = for r in r0 .. r1 do for c in c0 .. c1 do cells.Write r c v
    hostRW "Excel.Range"
        (fun name ->
            if eq name "Value" || eq name "Value2" || eq name "Text" || eq name "Formula" then Some (VString (cells.Read r0 c0))
            elif eq name "Address" then Some (VString (formatAddr true true))
            elif eq name "Row" then Some (VInteger (r0 + 1))
            elif eq name "Column" then Some (VInteger (c0 + 1))
            elif eq name "Count" then Some (VInteger ((r1 - r0 + 1) * (c1 - c0 + 1)))
            elif eq name "Worksheet" then Some (worksheetRW cells sheetName)
            elif eq name "Cells" then Some (rangeRW cells sheetName r0 c0 r1 c1)
            else None)
        (fun name value ->
            if eq name "Value" || eq name "Value2" || eq name "Formula" then setAll (valueToString value); true
            else false)
        (fun name args ->
            if (eq name "Cells" || eq name "Item") && args.Length >= 2 then
                let r = r0 + valueToInt args.[0] - 1
                let c = c0 + valueToInt args.[1] - 1
                Some (rangeRW cells sheetName r c r c)
            elif eq name "Offset" then
                let ro = if args.Length >= 1 then valueToInt args.[0] else 0
                let co = if args.Length >= 2 then valueToInt args.[1] else 0
                Some (rangeRW cells sheetName (r0 + ro) (c0 + co) (r1 + ro) (c1 + co))
            elif eq name "Resize" then
                let rs = if args.Length >= 1 then max 1 (valueToInt args.[0]) else r1 - r0 + 1
                let cs = if args.Length >= 2 then max 1 (valueToInt args.[1]) else c1 - c0 + 1
                Some (rangeRW cells sheetName r0 c0 (r0 + rs - 1) (c0 + cs - 1))
            elif eq name "Address" then Some (VString (formatAddr true true))
            else None)

and private worksheetRW (cells: ModelCells) (sheetName: string) =
    hostRW "Excel.Worksheet"
        (fun name ->
            if eq name "Name" then Some (VString sheetName)
            elif eq name "Cells" then Some (rangeRW cells sheetName 0 0 1048575 16383)
            else None)
        (fun _ _ -> false)
        (fun name args ->
            if eq name "Range" && args.Length >= 1 then
                match parseRangeRef (valueToString args.[0]) with
                | Some (a, b, c, d) -> Some (rangeRW cells sheetName a b c d)
                | None -> Some VEmpty
            elif eq name "Cells" && args.Length >= 2 then
                let r = valueToInt args.[0] - 1
                let c = valueToInt args.[1] - 1
                Some (rangeRW cells sheetName r c r c)
            else None)

let private workbookRW cells sheetName =
    hostRW "Excel.Workbook"
        (fun name -> if eq name "ActiveSheet" then Some (worksheetRW cells sheetName) else None)
        (fun _ _ -> false)
        (fun _ _ -> None)

let private applicationRW cells sheetName =
    hostRW "Excel.Application"
        (fun name ->
            if eq name "ActiveWorkbook" then Some (workbookRW cells sheetName)
            elif eq name "ActiveSheet" then Some (worksheetRW cells sheetName)
            elif eq name "Version" then Some (VString "15.0")
            else None)
        (fun _ _ -> false)
        (fun name args ->
            if eq name "Range" && args.Length >= 1 then
                match parseRangeRef (valueToString args.[0]) with
                | Some (a, b, c, d) -> Some (rangeRW cells sheetName a b c d)
                | None -> Some VEmpty
            elif eq name "Cells" && args.Length >= 2 then
                Some (rangeRW cells sheetName (valueToInt args.[0] - 1) (valueToInt args.[1] - 1) (valueToInt args.[0] - 1) (valueToInt args.[1] - 1))
            else None)

let private rangeFnRW cells sheetName =
    hostRW "Excel.RangeFunction" (fun _ -> None) (fun _ _ -> false)
        (fun name args ->
            if isDefaultMemberName name && args.Length >= 1 then
                match parseRangeRef (valueToString args.[0]) with
                | Some (a, b, c, d) -> Some (rangeRW cells sheetName a b c d)
                | None -> Some VEmpty
            else None)

let private cellsFnRW cells sheetName =
    hostRW "Excel.CellsFunction" (fun _ -> None) (fun _ _ -> false)
        (fun name args ->
            if isDefaultMemberName name && args.Length >= 2 then
                Some (rangeRW cells sheetName (valueToInt args.[0] - 1) (valueToInt args.[1] - 1) (valueToInt args.[0] - 1) (valueToInt args.[1] - 1))
            else None)

let private hostGlobalsRW cells sheetName =
    Map.ofList [
        "Application", applicationRW cells sheetName
        "ActiveWorkbook", workbookRW cells sheetName
        "ActiveSheet", worksheetRW cells sheetName
        "Range", rangeFnRW cells sheetName
        "Cells", cellsFnRW cells sheetName
    ]

// The model-backed host always binds the same global NAMES, so a module
// compiles to the same bytecode regardless of which model it runs against;
// only the bound host VALUES change per run.
let private vbaHostNames = [ "Application"; "ActiveWorkbook"; "ActiveSheet"; "Range"; "Cells" ]
let private vbaParseOpts = FsVb.ParseOptions.withHostGlobals vbaHostNames FsVb.ParseOptions.vba

let private bytecodeCache = System.Collections.Generic.Dictionary<string, Bytecode.BytecodeProgram>()

/// Compile a macro module to bytecode, caching the result by source text.
let compileCached (source: string) : Result<Bytecode.BytecodeProgram, string> =
    match bytecodeCache.TryGetValue source with
    | true, bc -> Result.Ok bc
    | _ ->
        match FsVb.compileSource vbaParseOpts source with
        | Result.Ok (_, bc) -> bytecodeCache.[source] <- bc; Result.Ok bc
        | Result.Error e -> Result.Error (sprintf "%A" e)

/// Execute an already-compiled program against a model with read+write host
/// objects, fold the writes back in (a `=…` value becomes a formula), and
/// recompute. Shared by the compile path and the run-from-bytecode path.
let private runProgramOnModel (sheetName: string) (model: SpreadsheetModel) (bytecode: Bytecode.BytecodeProgram) : Result<SpreadsheetModel, string> =
    let cells = ModelCells(model)
    let runErr =
        try FsVb.executeWithHostGlobals (hostGlobalsRW cells sheetName |> Map.toSeq) bytecode |> ignore; None
        with ex -> Some ex.Message
    match runErr with
    | Some e -> Result.Error e
    | None ->
        let mutable m = model
        for kv in cells.Writes do
            let struct (r, c) = kv.Key
            let text = kv.Value
            let cv =
                if text.StartsWith "=" then CellValue.Formula (text.Substring 1, CellValue.Empty)
                else
                    match System.Double.TryParse(text, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture) with
                    | true, n -> CellValue.Number n
                    | _ -> CellValue.String text
            let fmt = match getCell { Row = r; Col = c } m with Some cd -> cd.Format | None -> defaultFormat
            m <- setCell { Row = r; Col = c } { Value = cv; Format = fmt; Comment = None } m
        Result.Ok (RedBat.Spreadsheet.Formula.recalcAndSpill None (Some sheetName) m)

/// Run a macro against a model with read+write host objects, fold the writes
/// back in (a `=…` value becomes a formula), and recompute. Cross-platform.
let runOnModelWithSheetName (sheetName: string) (model: SpreadsheetModel) (source: string) : Result<SpreadsheetModel, string> =
    match compileCached source with
    | Result.Error e -> Result.Error e
    | Result.Ok bytecode -> runProgramOnModel sheetName model bytecode

let runOnModel (model: SpreadsheetModel) (source: string) : Result<SpreadsheetModel, string> =
    runOnModelWithSheetName "Sheet1" model source

/// Compile a macro module to a portable bytecode blob (our "p-code"): save it,
/// then re-run later with `runBytesOnModel` without re-parsing the source.
/// Native-only (serialization uses System.IO).
let compileToBytes (source: string) : Result<byte[], string> =
    compileCached source |> Result.map BytecodeIO.serialize

/// Run a previously-serialized bytecode blob against a model — skips lex/parse/
/// compile entirely. Mirrors how Office runs persisted p-code.
let runBytesOnModelWithSheetName (sheetName: string) (model: SpreadsheetModel) (blob: byte[]) : Result<SpreadsheetModel, string> =
    match (try Result.Ok (BytecodeIO.deserialize blob) with ex -> Result.Error ex.Message) with
    | Result.Error e -> Result.Error e
    | Result.Ok bytecode -> runProgramOnModel sheetName model bytecode

let runBytesOnModel (model: SpreadsheetModel) (blob: byte[]) : Result<SpreadsheetModel, string> =
    runBytesOnModelWithSheetName "Sheet1" model blob

/// Does the module source define a `Sub name` (optionally Private/Public)?
let hasSub (name: string) (source: string) : bool =
    not (System.String.IsNullOrWhiteSpace source)
    && System.Text.RegularExpressions.Regex.IsMatch(
        source,
        @"(?im)^\s*(Private\s+|Public\s+)?Sub\s+" + System.Text.RegularExpressions.Regex.Escape(name) + @"\b")

/// Fire a workbook/worksheet event: if `model.Macros` defines `Sub eventName`,
/// run the module plus a top-level call to it (with optional VBA argument text,
/// e.g. `Range("B2")` for Worksheet_Change), and return the updated model.
/// A missing sub or a run error leaves the model unchanged. Writes are applied
/// directly (not through the editor), so firing on a cell edit can't recurse.
let fireEvent (eventName: string) (callArgs: string) (model: SpreadsheetModel) : SpreadsheetModel =
    if not (hasSub eventName model.Macros) then model
    else
        // A bare `Foo` parses as an expression, not a call; `Foo()` / `Foo arg`
        // invoke the sub.
        let call = if System.String.IsNullOrEmpty callArgs then eventName + "()" else eventName + " " + callArgs
        match runOnModel model (model.Macros + "\n" + call) with
        | Result.Ok m -> m
        | Result.Error _ -> model

let runSnapshot sheetName snapshot source =
    let options =
        FsVb.RunOptions.vba
        |> FsVb.RunOptions.withHostGlobals (hostGlobalsForSnapshot sheetName snapshot)
    FsVb.runSource options source

let runSnapshotVbs sheetName snapshot source =
    let options =
        FsVb.RunOptions.vbScript
        |> FsVb.RunOptions.withHostGlobals (hostGlobalsForSnapshotVbs sheetName snapshot)
    FsVb.runSource options source

let private runSnapshotCells sheetName snapshot source =
    let options =
        FsVb.RunOptions.vba
        |> FsVb.RunOptions.withHostGlobals (hostGlobalsForSnapshotCells sheetName snapshot)
    FsVb.runSource options source

let private runSnapshotCellsVbs sheetName snapshot source =
    let options =
        FsVb.RunOptions.vbScript
        |> FsVb.RunOptions.withHostGlobals (hostGlobalsForSnapshotCellsVbs sheetName snapshot)
    FsVb.runSource options source

let runModelWithSheetName sheetName (model: SpreadsheetModel) source =
    runSnapshotCells sheetName (snapshotCellsFromModel model) source

let runModel (model: SpreadsheetModel) source =
    runModelWithSheetName "Sheet1" model source

let runModelWithSheetNameVbs sheetName (model: SpreadsheetModel) source =
    runSnapshotCellsVbs sheetName (snapshotCellsFromModel model) source

let runModelVbs (model: SpreadsheetModel) source =
    runModelWithSheetNameVbs "Sheet1" model source

let rec private valueToObj (value: Value) =
    match value with
    | VInteger i -> box i
    | VDouble d -> box d
    | VString s -> box s
    | VBoolean b -> box b
    | VArray values -> values |> Array.map valueToObj |> box
    | VRef (values, idx) when idx >= 0 && idx < values.Length -> valueToObj values.[idx]
    | VRef _ -> null
    | VEmpty
    | VNull
    | VNothing
    | VUndefined -> null
    | VObject o -> box o.ClassName
    | VHostObject o -> box o.TypeName

let runForJs (spreadsheet: Spreadsheet.Spreadsheet) (source: string) (resultName: string) =
    match run spreadsheet source with
    | Result.Ok result ->
        let value =
            if isNull (box resultName) || resultName = "" then null
            else FsVb.tryGetGlobal resultName result.State |> Option.map valueToObj |> Option.defaultValue null
        createObj [
            "ok" ==> true
            "result" ==> value
        ]
    | Result.Error err ->
        createObj [
            "ok" ==> false
            "error" ==> (sprintf "%A" err)
        ]

let runVbsForJs (spreadsheet: Spreadsheet.Spreadsheet) (source: string) (resultName: string) =
    match runVbs spreadsheet source with
    | Result.Ok result ->
        let value =
            if isNull (box resultName) || resultName = "" then null
            else FsVb.tryGetGlobal resultName result.State |> Option.map valueToObj |> Option.defaultValue null
        createObj [
            "ok" ==> true
            "result" ==> value
        ]
    | Result.Error err ->
        createObj [
            "ok" ==> false
            "error" ==> (sprintf "%A" err)
        ]
