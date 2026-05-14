module RedBat.Spreadsheet.Formula

open System
open System.Globalization
open RedBat.Spreadsheet.Model

// ---------------------------------------------------------------------------
// Excel error values. Stored as the bare code (no leading '#'); the display
// layer prepends the '#'. These are the seven Excel error literals plus the
// internal CYCLE sentinel for circular references.
// ---------------------------------------------------------------------------
module Err =
    let div0 = "DIV/0!"
    let value = "VALUE!"
    let ref = "REF!"
    let name = "NAME?"
    let num = "NUM!"
    let na = "N/A"
    let nul = "NULL!"
    let cycle = "CYCLE"
    /// Every recognized error literal, for the lexer.
    let all = [ "#DIV/0!"; "#VALUE!"; "#REF!"; "#NAME?"; "#NUM!"; "#N/A"; "#NULL!" ]

// ---------------------------------------------------------------------------
// AST
// ---------------------------------------------------------------------------

/// A parsed cell reference. Row/Col are the 0-indexed resolved position.
/// RowAbs/ColAbs record whether the source used a `$` lock (kept for
/// copy/fill + conditional-format re-anchoring, ignored by plain eval).
/// Sheet is the qualifying sheet name when the ref was `Sheet!A1`.
type RefAst = {
    Sheet: string option
    Row: int
    Col: int
    RowAbs: bool
    ColAbs: bool
}

type Expr =
    | Number of float
    | Str of string
    | Bool of bool
    /// Error literal written directly in the formula (`#REF!` etc). Carries
    /// the bare code.
    | ErrorLit of string
    | CellRef of RefAst
    | RangeRef of RefAst * RefAst
    | TableRef of tableName: string * columnName: string
    /// Named range / defined name — resolved at eval time against the
    /// workbook's NamedRanges (S3); for now evaluates to #NAME?.
    | NameRef of string
    /// Prefix `-`/`+` or postfix `%` (modeled as unary with op "%").
    | UnaryOp of op: string * operand: Expr
    | BinOp of left: Expr * op: string * right: Expr
    | Function of name: string * args: Expr list

// ---------------------------------------------------------------------------
// Lexer
// ---------------------------------------------------------------------------

type Token =
    | TNumber of float
    | TString of string
    /// Cell ref / function name / defined name / TRUE / FALSE — disambiguated
    /// by the parser. May contain `$` (e.g. "$A$1").
    | TIdent of string
    /// Quoted sheet name from `'My Sheet'!A1`.
    | TSheet of string
    /// Error literal, bare code (no '#').
    | TError of string
    | TPlus | TMinus | TStar | TSlash | TCaret | TPercent | TAmp
    | TEq | TNe | TLt | TGt | TLe | TGe
    | TLParen | TRParen
    | TLBracket | TRBracket
    | TComma | TColon | TBang
    | TEnd

let tokenize (input: string) : Token list =
    let chars = input.ToCharArray()
    let n = chars.Length
    let tokens = System.Collections.Generic.List<Token>()
    let mutable i = 0
    while i < n do
        let c = chars.[i]
        if Char.IsWhiteSpace c then
            i <- i + 1
        elif Char.IsDigit c || (c = '.' && i + 1 < n && Char.IsDigit chars.[i + 1]) then
            // Number, with optional fraction and scientific exponent.
            let start = i
            while i < n && Char.IsDigit chars.[i] do i <- i + 1
            if i < n && chars.[i] = '.' then
                i <- i + 1
                while i < n && Char.IsDigit chars.[i] do i <- i + 1
            if i < n && (chars.[i] = 'e' || chars.[i] = 'E') then
                let save = i
                i <- i + 1
                if i < n && (chars.[i] = '+' || chars.[i] = '-') then i <- i + 1
                if i < n && Char.IsDigit chars.[i] then
                    while i < n && Char.IsDigit chars.[i] do i <- i + 1
                else
                    // 'E' not followed by an exponent — not part of the number.
                    i <- save
            let text = System.String(chars, start, i - start)
            tokens.Add(TNumber (Double.Parse(text, CultureInfo.InvariantCulture)))
        elif c = '"' then
            // String literal. "" is an escaped quote.
            i <- i + 1
            let sb = System.Text.StringBuilder()
            let mutable closed = false
            while i < n && not closed do
                if chars.[i] = '"' then
                    if i + 1 < n && chars.[i + 1] = '"' then
                        sb.Append('"') |> ignore
                        i <- i + 2
                    else
                        closed <- true
                        i <- i + 1
                else
                    sb.Append(chars.[i]) |> ignore
                    i <- i + 1
            tokens.Add(TString (sb.ToString()))
        elif c = '\'' then
            // Quoted sheet name. '' is an escaped apostrophe.
            i <- i + 1
            let sb = System.Text.StringBuilder()
            let mutable closed = false
            while i < n && not closed do
                if chars.[i] = '\'' then
                    if i + 1 < n && chars.[i + 1] = '\'' then
                        sb.Append('\'') |> ignore
                        i <- i + 2
                    else
                        closed <- true
                        i <- i + 1
                else
                    sb.Append(chars.[i]) |> ignore
                    i <- i + 1
            tokens.Add(TSheet (sb.ToString()))
        elif c = '#' then
            // Error literal: longest match against the known set.
            let rest = System.String(chars, i, n - i)
            match Err.all |> List.tryFind (fun e -> rest.StartsWith(e, StringComparison.OrdinalIgnoreCase)) with
            | Some lit ->
                tokens.Add(TError (lit.Substring 1))
                i <- i + lit.Length
            | None ->
                i <- i + 1  // stray '#', skip
        elif c = '$' || Char.IsLetter c || c = '_' then
            // Reference / name / function name. Allow '$', letters, digits,
            // '_', '.' so "$A$1", "Sheet_2", "_xlfn.FOO" all lex as one token.
            let start = i
            while i < n && (chars.[i] = '$' || Char.IsLetterOrDigit chars.[i] || chars.[i] = '_' || chars.[i] = '.') do
                i <- i + 1
            tokens.Add(TIdent (System.String(chars, start, i - start)))
        else
            // Operators / punctuation. Two-char operators checked first.
            let two = if i + 1 < n then System.String(chars, i, 2) else ""
            match two with
            | "<=" -> tokens.Add(TLe); i <- i + 2
            | ">=" -> tokens.Add(TGe); i <- i + 2
            | "<>" -> tokens.Add(TNe); i <- i + 2
            | _ ->
                match c with
                | '+' -> tokens.Add(TPlus); i <- i + 1
                | '-' -> tokens.Add(TMinus); i <- i + 1
                | '*' -> tokens.Add(TStar); i <- i + 1
                | '/' -> tokens.Add(TSlash); i <- i + 1
                | '^' -> tokens.Add(TCaret); i <- i + 1
                | '%' -> tokens.Add(TPercent); i <- i + 1
                | '&' -> tokens.Add(TAmp); i <- i + 1
                | '=' -> tokens.Add(TEq); i <- i + 1
                | '<' -> tokens.Add(TLt); i <- i + 1
                | '>' -> tokens.Add(TGt); i <- i + 1
                | '(' -> tokens.Add(TLParen); i <- i + 1
                | ')' -> tokens.Add(TRParen); i <- i + 1
                | '[' -> tokens.Add(TLBracket); i <- i + 1
                | ']' -> tokens.Add(TRBracket); i <- i + 1
                | ',' -> tokens.Add(TComma); i <- i + 1
                | ':' -> tokens.Add(TColon); i <- i + 1
                | '!' -> tokens.Add(TBang); i <- i + 1
                | _ -> i <- i + 1  // skip unknown
    tokens.Add(TEnd)
    List.ofSeq tokens

// ---------------------------------------------------------------------------
// Reference parsing
// ---------------------------------------------------------------------------

/// Parse an A1-style reference token (possibly with `$` locks) into a RefAst
/// with no sheet. Returns None when the token isn't a well-formed cell ref.
let parseRefToken (token: string) : RefAst option =
    // [ $ ] letters [ $ ] digits
    let mutable idx = 0
    let len = token.Length
    let colAbs = idx < len && token.[idx] = '$'
    if colAbs then idx <- idx + 1
    let colStart = idx
    while idx < len && Char.IsLetter token.[idx] do idx <- idx + 1
    let colPart = token.Substring(colStart, idx - colStart)
    let rowAbs = idx < len && token.[idx] = '$'
    if rowAbs then idx <- idx + 1
    let rowStart = idx
    while idx < len && Char.IsDigit token.[idx] do idx <- idx + 1
    let rowPart = token.Substring(rowStart, idx - rowStart)
    if idx <> len || colPart = "" || rowPart = "" then None
    else
        match Int32.TryParse rowPart with
        | true, r when r >= 1 ->
            Some {
                Sheet = None
                Row = r - 1
                Col = letterToColumn colPart
                RowAbs = rowAbs
                ColAbs = colAbs
            }
        | _ -> None

// ---------------------------------------------------------------------------
// Parser — recursive descent, Excel operator precedence.
//   loosest → tightest:
//     comparison (= <> < > <= >=)
//     concat (&)
//     add/sub (+ -)
//     mul/div (* /)
//     exponent (^)
//     percent (postfix %)
//     unary (prefix - +)
//     range (a : b)
//     primary
//   (Excel binds unary minus tighter than ^, so `-2^2` = 4.)
// ---------------------------------------------------------------------------

type private Parser = { Tokens: Token[]; mutable Pos: int }

let private peekP (p: Parser) = if p.Pos < p.Tokens.Length then p.Tokens.[p.Pos] else TEnd
let private nextP (p: Parser) = let t = peekP p in p.Pos <- p.Pos + 1; t
let private expectP (p: Parser) (t: Token) =
    if peekP p = t then p.Pos <- p.Pos + 1
    else failwithf "Expected %A but got %A" t (peekP p)

let rec private parsePrimary (p: Parser) : Expr =
    match peekP p with
    | TNumber n -> p.Pos <- p.Pos + 1; Number n
    | TString s -> p.Pos <- p.Pos + 1; Str s
    | TError e -> p.Pos <- p.Pos + 1; ErrorLit e
    | TMinus -> p.Pos <- p.Pos + 1; UnaryOp("-", parseUnary p)
    | TPlus -> p.Pos <- p.Pos + 1; UnaryOp("+", parseUnary p)
    | TLParen ->
        p.Pos <- p.Pos + 1
        let e = parseComparison p
        expectP p TRParen
        e
    | TSheet sheetName ->
        p.Pos <- p.Pos + 1
        expectP p TBang
        parseSheetQualified p (Some sheetName)
    | TIdent ident ->
        p.Pos <- p.Pos + 1
        match peekP p with
        | TLParen ->
            // Function call.
            p.Pos <- p.Pos + 1
            let args = parseArgs p
            expectP p TRParen
            Function(ident, args)
        | TLBracket ->
            // Structured table reference: Table[Column].
            p.Pos <- p.Pos + 1
            match nextP p with
            | TIdent col ->
                expectP p TRBracket
                TableRef(ident, col)
            | other -> failwithf "Expected column name in table ref, got %A" other
        | TBang ->
            // Unquoted sheet-qualified ref: Sheet1!A1.
            p.Pos <- p.Pos + 1
            parseSheetQualified p (Some ident)
        | _ ->
            // Cell ref, defined name, or boolean literal.
            match ident.ToUpperInvariant() with
            | "TRUE" -> Bool true
            | "FALSE" -> Bool false
            | _ ->
                match parseRefToken ident with
                | Some refAst -> CellRef refAst
                | None -> NameRef ident
    | other -> failwithf "Unexpected token %A" other

/// Parse the reference portion that follows `Sheet!`. May itself be a range.
and private parseSheetQualified (p: Parser) (sheet: string option) : Expr =
    match nextP p with
    | TIdent ident ->
        match parseRefToken ident with
        | Some refAst ->
            let withSheet = { refAst with Sheet = sheet }
            match peekP p with
            | TColon ->
                p.Pos <- p.Pos + 1
                match nextP p with
                | TIdent endIdent ->
                    match parseRefToken endIdent with
                    | Some endRef -> RangeRef(withSheet, { endRef with Sheet = sheet })
                    | None -> failwithf "Invalid cell ref after ':': %s" endIdent
                | other -> failwithf "Expected cell ref after ':', got %A" other
            | _ -> CellRef withSheet
        | None ->
            // Sheet-qualified defined name (rare). Treat as NameRef.
            NameRef ident
    | other -> failwithf "Expected reference after sheet '!', got %A" other

and private parseArgs (p: Parser) : Expr list =
    if peekP p = TRParen then []
    else
        let rec loop acc =
            let e = parseComparison p
            match peekP p with
            | TComma -> p.Pos <- p.Pos + 1; loop (e :: acc)
            | _ -> List.rev (e :: acc)
        loop []

and private parseRange (p: Parser) : Expr =
    let left = parsePrimary p
    match peekP p, left with
    | TColon, CellRef a ->
        p.Pos <- p.Pos + 1
        match parsePrimary p with
        | CellRef b -> RangeRef(a, b)
        | _ -> failwith "Expected cell reference after ':'"
    | _ -> left

and private parseUnary (p: Parser) : Expr =
    match peekP p with
    | TMinus -> p.Pos <- p.Pos + 1; UnaryOp("-", parseUnary p)
    | TPlus -> p.Pos <- p.Pos + 1; UnaryOp("+", parseUnary p)
    | _ -> parsePercent p

and private parsePercent (p: Parser) : Expr =
    let mutable e = parseRange p
    while peekP p = TPercent do
        p.Pos <- p.Pos + 1
        e <- UnaryOp("%", e)
    e

and private parseExponent (p: Parser) : Expr =
    let mutable e = parseUnary p
    while peekP p = TCaret do
        p.Pos <- p.Pos + 1
        e <- BinOp(e, "^", parseUnary p)
    e

and private parseMulDiv (p: Parser) : Expr =
    let mutable e = parseExponent p
    let mutable go = true
    while go do
        match peekP p with
        | TStar -> p.Pos <- p.Pos + 1; e <- BinOp(e, "*", parseExponent p)
        | TSlash -> p.Pos <- p.Pos + 1; e <- BinOp(e, "/", parseExponent p)
        | _ -> go <- false
    e

and private parseAddSub (p: Parser) : Expr =
    let mutable e = parseMulDiv p
    let mutable go = true
    while go do
        match peekP p with
        | TPlus -> p.Pos <- p.Pos + 1; e <- BinOp(e, "+", parseMulDiv p)
        | TMinus -> p.Pos <- p.Pos + 1; e <- BinOp(e, "-", parseMulDiv p)
        | _ -> go <- false
    e

and private parseConcat (p: Parser) : Expr =
    let mutable e = parseAddSub p
    while peekP p = TAmp do
        p.Pos <- p.Pos + 1
        e <- BinOp(e, "&", parseAddSub p)
    e

and private parseComparison (p: Parser) : Expr =
    let mutable e = parseConcat p
    let mutable go = true
    while go do
        let op =
            match peekP p with
            | TEq -> Some "=" | TNe -> Some "<>"
            | TLt -> Some "<" | TGt -> Some ">"
            | TLe -> Some "<=" | TGe -> Some ">="
            | _ -> None
        match op with
        | Some o -> p.Pos <- p.Pos + 1; e <- BinOp(e, o, parseConcat p)
        | None -> go <- false
    e

/// Parse a formula string (leading `=` optional). Returns the AST or an error
/// message describing the parse failure.
let parseFormula (input: string) : Result<Expr, string> =
    try
        let body = if input.StartsWith "=" then input.Substring 1 else input
        let tokens = tokenize body |> Array.ofList
        let p = { Tokens = tokens; Pos = 0 }
        let e = parseComparison p
        match peekP p with
        | TEnd -> Ok e
        | t -> Result.Error (sprintf "Unexpected trailing token %A" t)
    with ex -> Result.Error ex.Message

// ---------------------------------------------------------------------------
// Value helpers + coercion
// ---------------------------------------------------------------------------

/// Unwrap a cached formula value to its underlying scalar.
let rec unwrapValue (value: CellValue) : CellValue =
    match value with
    | CellValue.Formula (_, cached) -> unwrapValue cached
    | other -> other

let private err code : CellValue = CellValue.Error code

/// First error among the operands, if any — for error propagation.
let private firstError (vs: CellValue list) : CellValue option =
    vs |> List.tryPick (fun v -> match unwrapValue v with CellValue.Error _ as e -> Some e | _ -> None)

/// Format a number the way Excel coerces it to text for `&` / TEXT-less
/// contexts: integers without a decimal, otherwise the shortest round-trip.
let numberToText (n: float) : string =
    if Double.IsNaN n || Double.IsInfinity n then "#NUM!"
    elif n = Math.Floor n && abs n < 1e15 then
        (int64 n).ToString(CultureInfo.InvariantCulture)
    else
        n.ToString("G15", CultureInfo.InvariantCulture)

/// Coerce a value to a number using Excel's rules. Blank → 0, TRUE → 1,
/// FALSE → 0, numeric text → its value, non-numeric text → #VALUE!,
/// error → that error.
let coerceNumber (value: CellValue) : Result<float, string> =
    match unwrapValue value with
    | CellValue.Number n -> Ok n
    | CellValue.Boolean b -> Ok (if b then 1.0 else 0.0)
    | CellValue.Empty -> Ok 0.0
    | CellValue.Error e -> Result.Error e
    | CellValue.String s ->
        let t = s.Trim()
        match Double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture) with
        | true, n -> Ok n
        | _ ->
            // Excel also coerces "TRUE"/"FALSE" text in arithmetic? No — it
            // does not; only the boolean *value* coerces. Text TRUE → #VALUE!.
            Result.Error Err.value
    | CellValue.Formula _ -> Result.Error Err.value  // unreachable after unwrap

/// Coerce a value to text using Excel's rules.
let coerceText (value: CellValue) : Result<string, string> =
    match unwrapValue value with
    | CellValue.String s -> Ok s
    | CellValue.Number n -> Ok (numberToText n)
    | CellValue.Boolean b -> Ok (if b then "TRUE" else "FALSE")
    | CellValue.Empty -> Ok ""
    | CellValue.Error e -> Result.Error e
    | CellValue.Formula _ -> Result.Error Err.value

/// Coerce a value to a boolean for IF-style truthiness. Blank → false,
/// 0 → false, non-zero → true, "" → false, "TRUE"/"FALSE" text parsed.
let coerceBool (value: CellValue) : Result<bool, string> =
    match unwrapValue value with
    | CellValue.Boolean b -> Ok b
    | CellValue.Number n -> Ok (n <> 0.0)
    | CellValue.Empty -> Ok false
    | CellValue.Error e -> Result.Error e
    | CellValue.String s ->
        match s.Trim().ToUpperInvariant() with
        | "TRUE" -> Ok true
        | "FALSE" -> Ok false
        | "" -> Ok false
        | other ->
            match Double.TryParse(other, NumberStyles.Float, CultureInfo.InvariantCulture) with
            | true, n -> Ok (n <> 0.0)
            | _ -> Result.Error Err.value
    | CellValue.Formula _ -> Result.Error Err.value

// ---------------------------------------------------------------------------
// Reference resolution + evaluation
// ---------------------------------------------------------------------------

let private normalizeRange (a: RefAst) (b: RefAst) =
    (min a.Row b.Row, min a.Col b.Col, max a.Row b.Row, max a.Col b.Col)

/// Static set of cell addresses an expression references (same-sheet only;
/// cross-sheet deps handled by the calc engine in S3). Used by the
/// dependency tracker.
let rec referencesInExpr (expr: Expr) : Set<CellAddress> =
    match expr with
    | Number _ | Str _ | Bool _ | ErrorLit _ | NameRef _ | TableRef _ -> Set.empty
    | CellRef r -> Set.singleton { Row = r.Row; Col = r.Col }
    | RangeRef (a, b) ->
        let sr, sc, er, ec = normalizeRange a b
        seq { for row in sr .. er do for col in sc .. ec -> { Row = row; Col = col } }
        |> Set.ofSeq
    | UnaryOp (_, e) -> referencesInExpr e
    | BinOp (l, _, r) -> Set.union (referencesInExpr l) (referencesInExpr r)
    | Function (_, args) -> args |> List.fold (fun acc a -> Set.union acc (referencesInExpr a)) Set.empty

/// Read a single cell's value (same-sheet; cross-sheet resolution is S3).
let private readCell (r: RefAst) (model: SpreadsheetModel) : CellValue =
    getCellValue { Row = r.Row; Col = r.Col } model |> unwrapValue

/// Expand a RangeRef to the flat list of unwrapped cell values, row-major.
let expandRangeRef (a: RefAst) (b: RefAst) (model: SpreadsheetModel) : CellValue list =
    let sr, sc, er, ec = normalizeRange a b
    [ for row in sr .. er do
        for col in sc .. ec ->
            getCellValue { Row = row; Col = col } model |> unwrapValue ]

/// Expand a structured table reference to its column's values.
let expandTableRef (tableName: string) (columnName: string)
                   (workbook: RedBat.Spreadsheet.Workbook.Workbook option)
                   (model: SpreadsheetModel) : CellValue list =
    match workbook with
    | Some wb ->
        match RedBat.Spreadsheet.Workbook.getTableColumn tableName columnName wb with
        | Some (_, startAddr, endAddr) ->
            [ for row in startAddr.Row .. endAddr.Row ->
                getCellValue { Row = row; Col = startAddr.Col } model |> unwrapValue ]
        | None -> [ err Err.ref ]
    | None -> [ err Err.ref ]

/// Compare two scalar values for the comparison operators. Excel orders
/// number < text < FALSE < TRUE; numbers compare numerically, text
/// case-insensitively. Returns -1 / 0 / 1, or an Error value to propagate.
let private compareValues (l: CellValue) (r: CellValue) : Result<int, string> =
    let rank v =
        match v with
        | CellValue.Number _ -> 0
        | CellValue.String _ -> 1
        | CellValue.Boolean _ -> 2
        | CellValue.Empty -> 0  // blank behaves as 0 / "" depending on the other side
        | _ -> 3
    match unwrapValue l, unwrapValue r with
    | CellValue.Error e, _ -> Result.Error e
    | _, CellValue.Error e -> Result.Error e
    | a, b ->
        match a, b with
        | CellValue.Number x, CellValue.Number y -> Ok (compare x y)
        | CellValue.String x, CellValue.String y ->
            Ok (String.Compare(x, y, StringComparison.OrdinalIgnoreCase))
        | CellValue.Boolean x, CellValue.Boolean y -> Ok (compare x y)
        | CellValue.Empty, CellValue.Empty -> Ok 0
        // Blank vs number/text: blank acts as 0 or "".
        | CellValue.Empty, CellValue.Number y -> Ok (compare 0.0 y)
        | CellValue.Number x, CellValue.Empty -> Ok (compare x 0.0)
        | CellValue.Empty, CellValue.String y -> Ok (String.Compare("", y, StringComparison.OrdinalIgnoreCase))
        | CellValue.String x, CellValue.Empty -> Ok (String.Compare(x, "", StringComparison.OrdinalIgnoreCase))
        // Cross-type: order by rank.
        | _ -> Ok (compare (rank a) (rank b))

/// Evaluate an expression to a scalar CellValue.
let rec eval (expr: Expr) (model: SpreadsheetModel)
             (workbook: RedBat.Spreadsheet.Workbook.Workbook option) : CellValue =
    match expr with
    | Number n -> CellValue.Number n
    | Str s -> CellValue.String s
    | Bool b -> CellValue.Boolean b
    | ErrorLit e -> CellValue.Error e
    | NameRef _ -> err Err.name        // resolved against workbook names in S3
    | CellRef r -> readCell r model
    | RangeRef (a, b) ->
        // A bare range in scalar context yields its first cell (Excel's
        // implicit-intersection collapse is close enough for our use).
        match expandRangeRef a b model with
        | v :: _ -> v
        | [] -> err Err.ref
    | TableRef (t, c) ->
        match expandTableRef t c workbook model with
        | v :: _ -> v
        | [] -> err Err.ref
    | UnaryOp (op, operand) ->
        let v = eval operand model workbook
        match op with
        | "-" ->
            match coerceNumber v with
            | Ok n -> CellValue.Number (-n)
            | Result.Error e -> err e
        | "+" ->
            match coerceNumber v with
            | Ok n -> CellValue.Number n
            | Result.Error e -> err e
        | "%" ->
            match coerceNumber v with
            | Ok n -> CellValue.Number (n / 100.0)
            | Result.Error e -> err e
        | _ -> err Err.value
    | BinOp (left, op, right) ->
        let lv = eval left model workbook
        let rv = eval right model workbook
        match firstError [ lv; rv ] with
        | Some e -> e
        | None ->
            match op with
            | "&" ->
                match coerceText lv, coerceText rv with
                | Ok a, Ok b -> CellValue.String (a + b)
                | Result.Error e, _ | _, Result.Error e -> err e
            | "=" | "<>" | "<" | ">" | "<=" | ">=" ->
                match compareValues lv rv with
                | Result.Error e -> err e
                | Ok cmp ->
                    let result =
                        match op with
                        | "=" -> cmp = 0
                        | "<>" -> cmp <> 0
                        | "<" -> cmp < 0
                        | ">" -> cmp > 0
                        | "<=" -> cmp <= 0
                        | ">=" -> cmp >= 0
                        | _ -> false
                    CellValue.Boolean result
            | _ ->
                // Arithmetic — coerce both operands to numbers.
                match coerceNumber lv, coerceNumber rv with
                | Result.Error e, _ -> err e
                | _, Result.Error e -> err e
                | Ok a, Ok b ->
                    match op with
                    | "+" -> CellValue.Number (a + b)
                    | "-" -> CellValue.Number (a - b)
                    | "*" -> CellValue.Number (a * b)
                    | "/" -> if b = 0.0 then err Err.div0 else CellValue.Number (a / b)
                    | "^" ->
                        let p = a ** b
                        if Double.IsNaN p || Double.IsInfinity p then err Err.num
                        else CellValue.Number p
                    | _ -> err Err.value
    | Function (name, args) -> evalFunction name args model workbook

/// Evaluate every argument, flattening ranges/tables into their cell values
/// (the form most aggregate functions want).
and flattenArgs (args: Expr list) (model: SpreadsheetModel)
                (workbook: RedBat.Spreadsheet.Workbook.Workbook option) : CellValue list =
    args |> List.collect (fun arg ->
        match arg with
        | RangeRef (a, b) -> expandRangeRef a b model
        | TableRef (t, c) -> expandTableRef t c workbook model
        | _ -> [ eval arg model workbook ])

/// Numeric values from a flattened argument list — text and blanks dropped,
/// matching how SUM/AVERAGE/etc. ignore non-numeric cells. The first error
/// encountered is returned via the error channel so it can propagate.
and numbersOf (values: CellValue list) : Result<float list, string> =
    match firstError values with
    | Some (CellValue.Error e) -> Result.Error e
    | _ ->
        values
        |> List.choose (fun v ->
            match unwrapValue v with
            | CellValue.Number n -> Some n
            | CellValue.Boolean b -> Some (if b then 1.0 else 0.0)
            | _ -> None)
        |> Ok

and evalFunction (name: string) (args: Expr list) (model: SpreadsheetModel)
                 (workbook: RedBat.Spreadsheet.Workbook.Workbook option) : CellValue =
    let evalArg i = eval (List.item i args) model workbook
    let aggregateNumbers () = numbersOf (flattenArgs args model workbook)
    match name.ToUpperInvariant() with
    | "SUM" ->
        match aggregateNumbers () with
        | Ok ns -> CellValue.Number (List.sum ns)
        | Result.Error e -> err e
    | "AVERAGE" ->
        match aggregateNumbers () with
        | Ok [] -> err Err.div0
        | Ok ns -> CellValue.Number (List.average ns)
        | Result.Error e -> err e
    | "COUNT" ->
        // Counts numeric values only.
        let vals = flattenArgs args model workbook
        match firstError vals with
        | Some (CellValue.Error e) -> err e
        | _ ->
            let c =
                vals |> List.sumBy (fun v ->
                    match unwrapValue v with CellValue.Number _ -> 1 | _ -> 0)
            CellValue.Number (float c)
    | "MIN" ->
        match aggregateNumbers () with
        | Ok [] -> CellValue.Number 0.0
        | Ok ns -> CellValue.Number (List.min ns)
        | Result.Error e -> err e
    | "MAX" ->
        match aggregateNumbers () with
        | Ok [] -> CellValue.Number 0.0
        | Ok ns -> CellValue.Number (List.max ns)
        | Result.Error e -> err e
    | "IF" ->
        match args with
        | [ condExpr; trueExpr ] ->
            match coerceBool (eval condExpr model workbook) with
            | Ok true -> eval trueExpr model workbook
            | Ok false -> CellValue.Boolean false
            | Result.Error e -> err e
        | [ condExpr; trueExpr; falseExpr ] ->
            match coerceBool (eval condExpr model workbook) with
            | Ok true -> eval trueExpr model workbook
            | Ok false -> eval falseExpr model workbook
            | Result.Error e -> err e
        | _ -> err Err.value
    | _ ->
        // Unknown function — S2 expands this table.
        err Err.name

/// Parse + evaluate a formula string in one shot.
let evaluateFormula (formula: string) (model: SpreadsheetModel)
                    (workbook: RedBat.Spreadsheet.Workbook.Workbook option) : CellValue =
    match parseFormula formula with
    | Ok expr -> eval expr model workbook
    | Result.Error msg -> CellValue.Error msg

// ---------------------------------------------------------------------------
// Whole-model recalculation (dependency-ordered). Calc.fs owns the
// incremental engine; this is the simple full pass the shared model uses.
// ---------------------------------------------------------------------------

let recalculateAllFormulas (workbook: RedBat.Spreadsheet.Workbook.Workbook option)
                           (model: SpreadsheetModel) : SpreadsheetModel =
    let formulaCells =
        model.Cells
        |> Map.toList
        |> List.choose (fun (addr, cell) ->
            match cell.Value with
            | CellValue.Formula (formula, _) ->
                match parseFormula formula with
                | Ok expr -> Some (addr, cell, formula, expr, None)
                | Result.Error error -> Some (addr, cell, formula, Number 0.0, Some error)
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

    formulaCells |> List.iter (fun (addr, _, _, _, _) -> visit [] addr)

    let mutable updated = model
    for addr in List.rev order do
        match Map.tryFind addr formulaByAddress with
        | Some (cell, formula, _, Some parseError) ->
            updated <- setCell addr { cell with Value = CellValue.Formula (formula, CellValue.Error parseError) } updated
        | Some (cell, formula, _, None) when Set.contains addr cyclic ->
            updated <- setCell addr { cell with Value = CellValue.Formula (formula, CellValue.Error Err.cycle) } updated
        | Some (cell, formula, _, None) ->
            let result = evaluateFormula formula updated workbook
            updated <- setCell addr { cell with Value = CellValue.Formula (formula, result) } updated
        | None -> ()
    updated
