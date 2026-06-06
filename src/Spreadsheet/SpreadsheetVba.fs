module RedBat.Spreadsheet.Vba

open Fable.Core
open Fable.Core.JsInterop
open Values
open Ast
open RedBat.Spreadsheet.Model

[<Emit("$0[$1]")>]
let private getProp (_source: obj) (_name: string): obj = jsNative

let private eq (left: string) (right: string) =
    left.ToLowerInvariant() = right.ToLowerInvariant()

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

let private normalizeAddress (address: string) =
    let noSheet =
        match address.LastIndexOf("!") with
        | -1 -> address
        | idx -> address.Substring(idx + 1)
    noSheet.Replace("$", "").ToUpperInvariant()

let private parseRangeRef (reference: string) =
    let parts = reference.Split(':')
    let parse part = parseCellRef (normalizeAddress part)
    match parts with
    | [| one |] ->
        parse one |> Option.map (fun (r, c) -> r, c, r, c)
    | [| first; last |] ->
        match parse first, parse last with
        | Some (r1, c1), Some (r2, c2) -> Some (min r1 r2, min c1 c2, max r1 r2, max c1 c2)
        | _ -> None
    | _ -> None

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

let private pxToPt (value: float) = value * 72.0 / 96.0

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

let private snapshotFromModelWithViewport (model: SpreadsheetModel) (viewport: RedBat.Spreadsheet.Layout.Viewport) =
    let maxRow, maxCol =
        let mutable mr = 0
        let mutable mc = 0
        for KeyValue(addr, _) in model.Cells do
            mr <- max mr addr.Row
            mc <- max mc addr.Col
        for merge in model.MergedCells do
            mr <- max mr merge.EndRow
            mc <- max mc merge.EndCol
        min (model.RowCount - 1) mr, min (model.ColCount - 1) mc

    [|
        for row in 0 .. maxRow do
            for col in 0 .. maxCol do
                let isMergedTrailing =
                    model.MergedCells
                    |> List.exists (fun m ->
                        row >= m.StartRow && row <= m.EndRow &&
                        col >= m.StartCol && col <= m.EndCol &&
                        not (row = m.StartRow && col = m.StartCol))
                if not isMergedTrailing then
                    let addr = { Row = row; Col = col }
                    let merge =
                        model.MergedCells
                        |> List.tryFind (fun m ->
                            row >= m.StartRow && row <= m.EndRow &&
                            col >= m.StartCol && col <= m.EndCol)
                    let rect =
                        match merge with
                        | Some m when row = m.StartRow && col = m.StartCol -> RedBat.Spreadsheet.Layout.mergedRect model viewport m
                        | _ -> RedBat.Spreadsheet.Layout.cellRect model viewport row col
                    let cellData = getCell addr model
                    let value = getCellValue addr model
                    let format = cellData |> Option.map (fun c -> c.Format) |> Option.defaultValue defaultFormat
                    let fontPt = format.FontSize |> Option.map float |> Option.defaultValue 11.0
                    let displayText = RedBat.Spreadsheet.CellText.cellDisplayText None rect.Width format value
                    let mergedRows, mergedCols =
                        match merge with
                        | Some m -> m.EndRow - m.StartRow + 1, m.EndCol - m.StartCol + 1
                        | None -> 1, 1
                    yield {
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
    |]

let snapshotFromModel (model: SpreadsheetModel) =
    snapshotFromModelWithViewport model (RedBat.Spreadsheet.Layout.defaultViewport 100000.0 100000.0)

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
    let cells () = cellsInRange spreadsheet startRow startCol endRow endCol
    let firstCell () = cells () |> Array.tryHead

    let aggregateBounds () =
        let cells = cells ()
        if cells.Length = 0 then 0.0, 0.0, 0.0, 0.0
        else
            let left = cells |> Array.map (propFloat "x") |> Array.min
            let top = cells |> Array.map (propFloat "y") |> Array.min
            let right = cells |> Array.map (fun c -> propFloat "x" c + propFloat "width" c) |> Array.max
            let bottom = cells |> Array.map (fun c -> propFloat "y" c + propFloat "height" c) |> Array.max
            left, top, right - left, bottom - top

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
                    let rows = cells |> Array.map (propInt "row")
                    let cols = cells |> Array.map (propInt "column")
                    Some (rangeObject spreadsheet ((Array.min rows) - 1) ((Array.min cols) - 1) ((Array.max rows) - 1) ((Array.max cols) - 1))
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

let private cellsInSnapshot (snapshot: unit -> CellSnapshot array) startRow startCol endRow endCol =
    snapshot ()
    |> Array.filter (fun cell ->
        cell.Row >= startRow + 1 && cell.Row <= endRow + 1 &&
        cell.Column >= startCol + 1 && cell.Column <= endCol + 1)

let rec private rangeObjectForSnapshot (sheetName: string) (snapshot: unit -> CellSnapshot array) startRow startCol endRow endCol =
    let cells () = cellsInSnapshot snapshot startRow startCol endRow endCol
    let firstCell () = cells () |> Array.tryHead

    let aggregateBounds () =
        let cells = cells ()
        if cells.Length = 0 then 0.0, 0.0, 0.0, 0.0
        else
            let left = cells |> Array.map _.X |> Array.min
            let top = cells |> Array.map _.Y |> Array.min
            let right = cells |> Array.map (fun c -> c.X + c.Width) |> Array.max
            let bottom = cells |> Array.map (fun c -> c.Y + c.Height) |> Array.max
            left, top, right - left, bottom - top

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

and private usedRangeObjectForSnapshot (sheetName: string) (snapshot: unit -> CellSnapshot array) =
    let cells = snapshot ()
    if cells.Length = 0 then
        rangeObjectForSnapshot sheetName snapshot 0 0 0 0
    else
        rangeObjectForSnapshot sheetName snapshot
            ((cells |> Array.map _.Row |> Array.min) - 1)
            ((cells |> Array.map _.Column |> Array.min) - 1)
            ((cells |> Array.map _.Row |> Array.max) - 1)
            ((cells |> Array.map _.Column |> Array.max) - 1)

and private worksheetObjectForSnapshot (sheetName: string) (snapshot: unit -> CellSnapshot array) =
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

let hostGlobalsForSnapshot sheetName snapshot =
    Map.ofList [
        "Application", applicationObjectForSnapshot sheetName snapshot
        "ActiveWorkbook", workbookObjectForSnapshot sheetName snapshot
        "ActiveSheet", worksheetObjectForSnapshot sheetName snapshot
        "Range", rangeFunctionForSnapshot sheetName snapshot
        "Cells", cellsFunctionForSnapshot sheetName snapshot
    ]

let hostGlobalsForSnapshotVbs sheetName snapshot =
    hostGlobalsForSnapshot sheetName snapshot
    |> Map.add "CreateObject" (createObjectFunctionForSnapshot sheetName snapshot)

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

let runModelWithSheetName sheetName (model: SpreadsheetModel) source =
    let cells = lazy (snapshotFromModel model)
    runSnapshot sheetName (fun () -> cells.Force()) source

let runModel (model: SpreadsheetModel) source =
    runModelWithSheetName "Sheet1" model source

let runModelWithSheetNameVbs sheetName (model: SpreadsheetModel) source =
    let cells = lazy (snapshotFromModel model)
    runSnapshotVbs sheetName (fun () -> cells.Force()) source

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
