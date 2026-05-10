module RedBat.Spreadsheet.Formula

open System
open System.Globalization
open RedBat.Spreadsheet.Model

/// Formula expression AST
type Expr =
    | Number of float
    | CellRef of row: int * col: int
    | RangeRef of startRow: int * startCol: int * endRow: int * endCol: int
    | TableRef of tableName: string * columnName: string
    | BinOp of left: Expr * op: string * right: Expr
    | Function of name: string * args: Expr list

/// Tokenizer
type Token =
    | TNumber of float
    | TIdent of string
    | TPlus | TMinus | TStar | TSlash
    | TLParen | TRParen
    | TLBracket | TRBracket
    | TComma | TColon
    | TEnd

let tokenize (input: string): Token list =
    let mutable tokens = []
    let mutable i = 0
    let chars = input.ToCharArray()

    while i < chars.Length do
        let c = chars.[i]

        if Char.IsWhiteSpace(c) then
            i <- i + 1
        elif Char.IsDigit(c) || c = '.' then
            // Parse number
            let mutable num = ""
            while i < chars.Length && (Char.IsDigit(chars.[i]) || chars.[i] = '.') do
                num <- num + string chars.[i]
                i <- i + 1
            tokens <- tokens @ [TNumber (Double.Parse(num, CultureInfo.InvariantCulture))]
        elif Char.IsLetter(c) then
            // Parse identifier (cell ref or function name)
            let mutable ident = ""
            while i < chars.Length && (Char.IsLetterOrDigit(chars.[i])) do
                ident <- ident + string chars.[i]
                i <- i + 1
            tokens <- tokens @ [TIdent ident]
        elif c = '+' then
            tokens <- tokens @ [TPlus]
            i <- i + 1
        elif c = '-' then
            tokens <- tokens @ [TMinus]
            i <- i + 1
        elif c = '*' then
            tokens <- tokens @ [TStar]
            i <- i + 1
        elif c = '/' then
            tokens <- tokens @ [TSlash]
            i <- i + 1
        elif c = '(' then
            tokens <- tokens @ [TLParen]
            i <- i + 1
        elif c = ')' then
            tokens <- tokens @ [TRParen]
            i <- i + 1
        elif c = '[' then
            tokens <- tokens @ [TLBracket]
            i <- i + 1
        elif c = ']' then
            tokens <- tokens @ [TRBracket]
            i <- i + 1
        elif c = ',' then
            tokens <- tokens @ [TComma]
            i <- i + 1
        elif c = ':' then
            tokens <- tokens @ [TColon]
            i <- i + 1
        else
            i <- i + 1  // Skip unknown characters

    tokens @ [TEnd]

/// Parser
type Parser = {
    Tokens: Token list
    Position: int
}

let peek (parser: Parser): Token =
    if parser.Position < parser.Tokens.Length then
        parser.Tokens.[parser.Position]
    else
        TEnd

let advance (parser: Parser): Parser =
    { parser with Position = parser.Position + 1 }

let expect (token: Token) (parser: Parser): Parser =
    if peek parser = token then
        advance parser
    else
        failwith (sprintf "Expected %A but got %A" token (peek parser))

/// Parse primary expression (number, cell ref, function call, parenthesized expr)
let rec parsePrimary (parser: Parser): Expr * Parser =
    match peek parser with
    | TNumber n ->
        (Number n, advance parser)
    | TIdent ident ->
        // Could be cell ref, table ref, or function call
        let parser' = advance parser
        match peek parser' with
        | TLParen ->
            // Function call
            let parser'' = advance parser'
            let args, parser''' = parseArgs parser''
            let parser'''' = expect TRParen parser'''
            (Function(ident, args), parser'''')
        | TLBracket ->
            // Table reference: TableName[ColumnName]
            let parser'' = advance parser'
            match peek parser'' with
            | TIdent columnName ->
                let parser''' = advance parser''
                let parser'''' = expect TRBracket parser'''
                (TableRef(ident, columnName), parser'''')
            | _ ->
                failwith "Expected column name in table reference"
        | _ ->
            // Cell reference (potentially a range like A1:B10)
            match parseCellRef ident with
            | Some (row, col) ->
                // Check if followed by colon (range syntax)
                match peek parser' with
                | TColon ->
                    let parser'' = advance parser'
                    // Parse end cell reference
                    match peek parser'' with
                    | TIdent endIdent ->
                        let parser''' = advance parser''
                        match parseCellRef endIdent with
                        | Some (endRow, endCol) ->
                            (RangeRef(row, col, endRow, endCol), parser''')
                        | None -> failwith (sprintf "Invalid cell reference: %s" endIdent)
                    | _ -> failwith "Expected cell reference after colon"
                | _ ->
                    // Just a single cell reference
                    (CellRef(row, col), parser')
            | None -> failwith (sprintf "Invalid cell reference: %s" ident)
    | TLParen ->
        let parser' = advance parser
        let expr, parser'' = parseExpr parser'
        let parser''' = expect TRParen parser''
        (expr, parser''')
    | _ ->
        failwith (sprintf "Unexpected token: %A" (peek parser))

/// Parse function arguments
and parseArgs (parser: Parser): Expr list * Parser =
    if peek parser = TRParen then
        ([], parser)
    else
        let expr, parser' = parseExpr parser
        match peek parser' with
        | TComma ->
            let parser'' = advance parser'
            let rest, parser''' = parseArgs parser''
            (expr :: rest, parser''')
        | _ ->
            ([expr], parser')

/// Parse multiplication and division
and parseTerm (parser: Parser): Expr * Parser =
    let mutable left, parser' = parsePrimary parser

    let rec loop (left: Expr) (parser: Parser): Expr * Parser =
        match peek parser with
        | TStar ->
            let parser' = advance parser
            let right, parser'' = parsePrimary parser'
            loop (BinOp(left, "*", right)) parser''
        | TSlash ->
            let parser' = advance parser
            let right, parser'' = parsePrimary parser'
            loop (BinOp(left, "/", right)) parser''
        | _ ->
            (left, parser)

    loop left parser'

/// Parse addition and subtraction
and parseExpr (parser: Parser): Expr * Parser =
    let mutable left, parser' = parseTerm parser

    let rec loop (left: Expr) (parser: Parser): Expr * Parser =
        match peek parser with
        | TPlus ->
            let parser' = advance parser
            let right, parser'' = parseTerm parser'
            loop (BinOp(left, "+", right)) parser''
        | TMinus ->
            let parser' = advance parser
            let right, parser'' = parseTerm parser'
            loop (BinOp(left, "-", right)) parser''
        | _ ->
            (left, parser)

    loop left parser'

/// Parse formula (expects leading "=")
let parseFormula (input: string): Result<Expr, string> =
    try
        let formula = if input.StartsWith("=") then input.Substring(1) else input
        let tokens = tokenize formula
        let parser = { Tokens = tokens; Position = 0 }
        let expr, parser' = parseExpr parser
        match peek parser' with
        | TEnd -> Ok expr
        | token -> Result.Error (sprintf "Unexpected token: %A" token)
    with
    | ex -> Result.Error ex.Message

/// Unwrap a CellValue to get the actual computed value
/// Formulas store their cached result, so we need to extract it
let rec unwrapValue (value: CellValue): CellValue =
    match value with
    | CellValue.Formula (_, cached) -> unwrapValue cached
    | _ -> value

let private normalizedRange startRow startCol endRow endCol =
    (min startRow endRow, min startCol endCol, max startRow endRow, max startCol endCol)

let private referencesInExpr (expr: Expr): Set<CellAddress> =
    let rec loop expr =
        match expr with
        | Expr.Number _ -> Set.empty
        | Expr.CellRef (row, col) -> Set.singleton { Row = row; Col = col }
        | Expr.RangeRef (startRow, startCol, endRow, endCol) ->
            let sr, sc, er, ec = normalizedRange startRow startCol endRow endCol
            [
                for row in sr .. er do
                    for col in sc .. ec do
                        yield { Row = row; Col = col }
            ]
            |> Set.ofList
        | Expr.TableRef _ -> Set.empty
        | Expr.BinOp (left, _, right) -> Set.union (loop left) (loop right)
        | Expr.Function (_, args) ->
            args
            |> List.map loop
            |> List.fold Set.union Set.empty
    loop expr

/// Expand range reference to list of cell values
let expandRangeRef (startRow: int) (startCol: int) (endRow: int) (endCol: int) (model: SpreadsheetModel): CellValue list =
    let sr, sc, er, ec = normalizedRange startRow startCol endRow endCol
    [sr .. er]
    |> List.collect (fun row ->
        [sc .. ec]
        |> List.map (fun col ->
            let addr = { Row = row; Col = col }
            let value = getCellValue addr model
            unwrapValue value
        )
    )

/// Expand table reference to list of cell values
let expandTableRef (tableName: string) (columnName: string) (workbook: RedBat.Spreadsheet.Workbook.Workbook option) (model: SpreadsheetModel): CellValue list =
    match workbook with
    | Some wb ->
        match RedBat.Spreadsheet.Workbook.getTableColumn tableName columnName wb with
        | Some (sheetId, startAddr, endAddr) ->
            // Get all values in the column range
            [startAddr.Row .. endAddr.Row]
            |> List.map (fun row ->
                let addr = { Row = row; Col = startAddr.Col }
                let value = getCellValue addr model
                unwrapValue value
            )
        | None ->
            [CellValue.Error "REF"]  // Table or column not found
    | None ->
        [CellValue.Error "REF"]  // No workbook context

/// Evaluate expression
let rec eval (expr: Expr) (model: SpreadsheetModel) (workbook: RedBat.Spreadsheet.Workbook.Workbook option): CellValue =
    match expr with
    | Expr.Number n ->
        CellValue.Number n

    | Expr.CellRef (row, col) ->
        let addr = { Row = row; Col = col }
        let value = getCellValue addr model
        unwrapValue value

    | Expr.RangeRef (startRow, startCol, endRow, endCol) ->
        // Range references in scalar context return first value
        let values = expandRangeRef startRow startCol endRow endCol model
        match values with
        | [] -> CellValue.Error "REF"
        | v :: _ -> v

    | Expr.TableRef (tableName, columnName) ->
        // Table references in scalar context return first value
        let values = expandTableRef tableName columnName workbook model
        match values with
        | [] -> CellValue.Error "REF"
        | v :: _ -> v

    | Expr.BinOp (left, op, right) ->
        let leftVal = eval left model workbook
        let rightVal = eval right model workbook

        match leftVal, rightVal with
        | CellValue.Number l, CellValue.Number r ->
            match op with
            | "+" -> CellValue.Number (l + r)
            | "-" -> CellValue.Number (l - r)
            | "*" -> CellValue.Number (l * r)
            | "/" -> if r = 0.0 then CellValue.Error "DIV0" else CellValue.Number (l / r)
            | _ -> CellValue.Error "INVALID_OP"
        | _ -> CellValue.Error "VALUE"

    | Expr.Function (name, args) ->
        evalFunction name args model workbook

/// Evaluate function
and evalFunction (name: string) (args: Expr list) (model: SpreadsheetModel) (workbook: RedBat.Spreadsheet.Workbook.Workbook option): CellValue =
    match name.ToUpper() with
    | "SUM" ->
        // Collect all values, expanding ranges and table references
        let allValues =
            args |> List.collect (fun arg ->
                match arg with
                | Expr.RangeRef (startRow, startCol, endRow, endCol) ->
                    expandRangeRef startRow startCol endRow endCol model
                | Expr.TableRef (tableName, columnName) ->
                    expandTableRef tableName columnName workbook model
                | _ ->
                    [eval arg model workbook]
            )
        let numbers = allValues |> List.choose (fun v ->
            let unwrapped = unwrapValue v
            match unwrapped with
            | CellValue.Number n -> Some n
            | _ -> None
        )
        CellValue.Number (List.sum numbers)

    | "AVERAGE" ->
        // Collect all values, expanding ranges and table references
        let allValues =
            args |> List.collect (fun arg ->
                match arg with
                | Expr.RangeRef (startRow, startCol, endRow, endCol) ->
                    expandRangeRef startRow startCol endRow endCol model
                | Expr.TableRef (tableName, columnName) ->
                    expandTableRef tableName columnName workbook model
                | _ ->
                    [eval arg model workbook]
            )
        let numbers = allValues |> List.choose (fun v ->
            let unwrapped = unwrapValue v
            match unwrapped with
            | CellValue.Number n -> Some n
            | _ -> None
        )
        if numbers.IsEmpty then
            CellValue.Error "DIV0"
        else
            CellValue.Number (List.average numbers)

    | "COUNT" ->
        // Count non-empty cells in args
        let allValues =
            args |> List.collect (fun arg ->
                match arg with
                | Expr.RangeRef (startRow, startCol, endRow, endCol) ->
                    expandRangeRef startRow startCol endRow endCol model
                | Expr.TableRef (tableName, columnName) ->
                    expandTableRef tableName columnName workbook model
                | _ ->
                    [eval arg model workbook]
            )
        let count =
            allValues
            |> List.filter (fun v ->
                let unwrapped = unwrapValue v
                match unwrapped with
                | CellValue.Empty -> false
                | CellValue.Error _ -> false
                | _ -> true)
            |> List.length
        CellValue.Number (float count)

    | "IF" ->
        // IF(condition, trueValue, falseValue)
        match args with
        | [condExpr; trueExpr; falseExpr] ->
            let condValRaw = eval condExpr model workbook
            let condVal = unwrapValue condValRaw
            match condVal with
            | CellValue.Boolean true -> eval trueExpr model workbook
            | CellValue.Boolean false -> eval falseExpr model workbook
            | CellValue.Number n ->
                // Excel treats non-zero as true, zero as false
                if n <> 0.0 then eval trueExpr model workbook
                else eval falseExpr model workbook
            | CellValue.String s ->
                // Non-empty strings are true
                if not (System.String.IsNullOrEmpty(s)) then eval trueExpr model workbook
                else eval falseExpr model workbook
            | CellValue.Empty -> eval falseExpr model workbook
            | CellValue.Error _ -> condVal
            | _ -> CellValue.Error "VALUE"
        | _ -> CellValue.Error "ARGS"

    | "MIN" ->
        // Find minimum value
        let allValues =
            args |> List.collect (fun arg ->
                match arg with
                | Expr.RangeRef (startRow, startCol, endRow, endCol) ->
                    expandRangeRef startRow startCol endRow endCol model
                | Expr.TableRef (tableName, columnName) ->
                    expandTableRef tableName columnName workbook model
                | _ ->
                    [eval arg model workbook]
            )
        let numbers = allValues |> List.choose (fun v ->
            let unwrapped = unwrapValue v
            match unwrapped with
            | CellValue.Number n -> Some n
            | _ -> None
        )
        if numbers.IsEmpty then
            CellValue.Error "VALUE"
        else
            CellValue.Number (List.min numbers)

    | "MAX" ->
        // Find maximum value
        let allValues =
            args |> List.collect (fun arg ->
                match arg with
                | Expr.RangeRef (startRow, startCol, endRow, endCol) ->
                    expandRangeRef startRow startCol endRow endCol model
                | Expr.TableRef (tableName, columnName) ->
                    expandTableRef tableName columnName workbook model
                | _ ->
                    [eval arg model workbook]
            )
        let numbers = allValues |> List.choose (fun v ->
            let unwrapped = unwrapValue v
            match unwrapped with
            | CellValue.Number n -> Some n
            | _ -> None
        )
        if numbers.IsEmpty then
            CellValue.Error "VALUE"
        else
            CellValue.Number (List.max numbers)

    | _ ->
        CellValue.Error "NAME"

/// Evaluate formula and return result
let evaluateFormula (formula: string) (model: SpreadsheetModel) (workbook: RedBat.Spreadsheet.Workbook.Workbook option): CellValue =
    match parseFormula formula with
    | Ok expr -> eval expr model workbook
    | Result.Error msg -> CellValue.Error msg

/// Recalculate all cached formula values. This intentionally stays simple for the
/// shared model; Calc.fs owns dependency-aware recalculation.
let recalculateAllFormulas (workbook: RedBat.Spreadsheet.Workbook.Workbook option) (model: SpreadsheetModel): SpreadsheetModel =
    let formulaCells =
        model.Cells
        |> Map.toList
        |> List.choose (fun (addr, cell) ->
            match cell.Value with
            | CellValue.Formula (formula, _) ->
                match parseFormula formula with
                | Ok expr -> Some (addr, cell, formula, expr, None)
                | Result.Error error -> Some (addr, cell, formula, Expr.Number 0.0, Some error)
            | _ -> None)

    let formulaByAddress =
        formulaCells
        |> List.map (fun (addr, cell, formula, expr, parseError) -> addr, (cell, formula, expr, parseError))
        |> Map.ofList

    let dependencies =
        formulaCells
        |> List.map (fun (addr, _, _, expr, parseError) ->
            let refs =
                match parseError with
                | Some _ -> Set.empty
                | None ->
                    referencesInExpr expr
                    |> Set.filter (fun refAddr -> Map.containsKey refAddr formulaByAddress)
            addr, refs)
        |> Map.ofList

    let mutable visited = Set.empty
    let mutable visiting = Set.empty
    let mutable order = []
    let mutable cyclic = Set.empty

    let rec visit path addr =
        if Set.contains addr visiting then
            let cycleNodes =
                path
                |> List.takeWhile ((<>) addr)
                |> fun nodes -> addr :: nodes
            cyclic <- Set.union cyclic (Set.ofList cycleNodes)
        elif not (Set.contains addr visited) then
            visiting <- Set.add addr visiting
            let deps = Map.tryFind addr dependencies |> Option.defaultValue Set.empty
            for dep in deps do
                visit (addr :: path) dep
            visiting <- Set.remove addr visiting
            visited <- Set.add addr visited
            order <- addr :: order

    formulaCells
    |> List.iter (fun (addr, _, _, _, _) -> visit [] addr)

    let mutable updated = model
    for addr in List.rev order do
        match Map.tryFind addr formulaByAddress with
        | Some (cell, formula, _, Some parseError) ->
            updated <- setCell addr { cell with Value = CellValue.Formula (formula, CellValue.Error parseError) } updated
        | Some (cell, formula, _, None) when Set.contains addr cyclic ->
            updated <- setCell addr { cell with Value = CellValue.Formula (formula, CellValue.Error "CYCLE") } updated
        | Some (cell, formula, _, None) ->
            let result = evaluateFormula formula updated workbook
            updated <- setCell addr { cell with Value = CellValue.Formula (formula, result) } updated
        | None -> ()
    updated
