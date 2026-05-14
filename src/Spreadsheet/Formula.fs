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

/// Resolve a defined name to its final parsed expression, scoped to the
/// current sheet (a sheet-local name shadows a workbook-global one). A
/// name's refers-to text is itself a formula — usually a range, sometimes
/// a constant or another name; name→name hops are followed with a cycle
/// guard. Returns None when the name is unknown, won't parse, or cycles.
let resolveNameExpr (name: string) (currentSheet: RedBat.Spreadsheet.Workbook.SheetId option)
                    (workbook: RedBat.Spreadsheet.Workbook.Workbook) : Expr option =
    let rec resolve (n: string) (seen: Set<string>) (hops: int) =
        let key = n.ToUpperInvariant()
        if hops > 16 || Set.contains key seen then None
        else
            match RedBat.Spreadsheet.Workbook.getName currentSheet n workbook with
            | None -> None
            | Some refersTo ->
                let body = if refersTo.StartsWith "=" then refersTo else "=" + refersTo
                match parseFormula body with
                | Result.Error _ -> None
                | Ok (NameRef inner) -> resolve inner (Set.add key seen) (hops + 1)
                | Ok expr -> Some expr
    resolve name Set.empty 0

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

/// Set of cell addresses an expression references *within the current
/// model*. Cross-sheet refs (`r.Sheet = Some _`) are deliberately excluded
/// — single-model recalc ordering can't span sheets, so they're treated as
/// external constants. A `NameRef` is resolved through the workbook (when
/// one is supplied) and its target cells contribute to the set, so
/// `=SUM(MyRange)` registers a dependency on the range's members. Defined
/// names resolve at workbook-global scope here: a missed sheet-local name
/// only affects recalc *ordering*, never the computed values, since `eval`
/// resolves names with the correct scope. Used by the dependency tracker.
let referencesInExpr (workbook: RedBat.Spreadsheet.Workbook.Workbook option)
                     (expr: Expr) : Set<CellAddress> =
    let rec go (seen: Set<string>) (e: Expr) : Set<CellAddress> =
        match e with
        | Number _ | Str _ | Bool _ | ErrorLit _ | TableRef _ -> Set.empty
        | NameRef name ->
            let key = name.ToUpperInvariant()
            if Set.contains key seen then Set.empty
            else
                match workbook |> Option.bind (resolveNameExpr name None) with
                | Some resolved -> go (Set.add key seen) resolved
                | None -> Set.empty
        | CellRef r when r.Sheet.IsSome -> Set.empty
        | CellRef r -> Set.singleton { Row = r.Row; Col = r.Col }
        | RangeRef (a, _) when a.Sheet.IsSome -> Set.empty
        | RangeRef (a, b) ->
            let sr, sc, er, ec = normalizeRange a b
            seq { for row in sr .. er do for col in sc .. ec -> { Row = row; Col = col } }
            |> Set.ofSeq
        | UnaryOp (_, inner) -> go seen inner
        | BinOp (l, _, r) -> Set.union (go seen l) (go seen r)
        | Function (_, args) -> args |> List.fold (fun acc a -> Set.union acc (go seen a)) Set.empty
    go Set.empty expr

/// Cross-sheet-aware reference extraction. Every reference is qualified to
/// an absolute (sheet, cell) pair: a ref's own sheet name when it has one,
/// otherwise `homeSheet` (the sheet the formula lives on). `NameRef`s
/// resolve through the workbook with a cycle guard. A ref whose sheet name
/// can't be resolved is dropped. This is what the incremental calc engine
/// keys its dependency graph on, so dependencies — and dirty propagation —
/// span sheet boundaries.
let referencesQualified (workbook: RedBat.Spreadsheet.Workbook.Workbook option)
                        (homeSheet: RedBat.Spreadsheet.Workbook.SheetId)
                        (expr: Expr) : Set<RedBat.Spreadsheet.Workbook.SheetCell> =
    let sheetIdOf (name: string option) =
        match name with
        | None -> Some homeSheet
        | Some n -> workbook |> Option.bind (RedBat.Spreadsheet.Workbook.resolveSheetName n)
    let rec go (seen: Set<string>) (e: Expr) : Set<RedBat.Spreadsheet.Workbook.SheetCell> =
        match e with
        | Number _ | Str _ | Bool _ | ErrorLit _ | TableRef _ -> Set.empty
        | NameRef name ->
            let key = name.ToUpperInvariant()
            if Set.contains key seen then Set.empty
            else
                match workbook |> Option.bind (resolveNameExpr name (Some homeSheet)) with
                | Some resolved -> go (Set.add key seen) resolved
                | None -> Set.empty
        | CellRef r ->
            match sheetIdOf r.Sheet with
            | Some sid ->
                Set.singleton
                    ({ Sheet = sid; Cell = { Row = r.Row; Col = r.Col } }
                     : RedBat.Spreadsheet.Workbook.SheetCell)
            | None -> Set.empty
        | RangeRef (a, b) ->
            match sheetIdOf a.Sheet with
            | Some sid ->
                let sr, sc, er, ec = normalizeRange a b
                seq { for row in sr .. er do
                        for col in sc .. ec ->
                            ({ Sheet = sid; Cell = { Row = row; Col = col } }
                             : RedBat.Spreadsheet.Workbook.SheetCell) }
                |> Set.ofSeq
            | None -> Set.empty
        | UnaryOp (_, inner) -> go seen inner
        | BinOp (l, _, r) -> Set.union (go seen l) (go seen r)
        | Function (_, args) -> args |> List.fold (fun acc a -> Set.union acc (go seen a)) Set.empty
    go Set.empty expr

/// Shift every *relative* reference in an expression by (dr, dc). Absolute
/// components (`$`-locked row/col) are left untouched. This is how a
/// formula written for one anchor cell is re-pointed for another — used by
/// fill/copy and by conditional-format `expression` rules, which store one
/// formula anchored at the range's top-left.
let rec shiftExpr (dr: int) (dc: int) (expr: Expr) : Expr =
    let shiftRef (r: RefAst) =
        { r with
            Row = if r.RowAbs then r.Row else r.Row + dr
            Col = if r.ColAbs then r.Col else r.Col + dc }
    match expr with
    | Number _ | Str _ | Bool _ | ErrorLit _ | NameRef _ | TableRef _ -> expr
    | CellRef r -> CellRef (shiftRef r)
    | RangeRef (a, b) -> RangeRef (shiftRef a, shiftRef b)
    | UnaryOp (op, e) -> UnaryOp (op, shiftExpr dr dc e)
    | BinOp (l, op, r) -> BinOp (shiftExpr dr dc l, op, shiftExpr dr dc r)
    | Function (name, args) -> Function (name, args |> List.map (shiftExpr dr dc))

/// True if the expression contains a volatile function (RAND/RANDBETWEEN/
/// NOW/TODAY/OFFSET/INDIRECT) anywhere in its tree. Such cells must be
/// recalculated on every pass regardless of dependency state.
let rec containsVolatile (expr: Expr) : bool =
    match expr with
    | Function (name, args) ->
        let n = name.ToUpperInvariant()
        n = "RAND" || n = "RANDBETWEEN" || n = "NOW" || n = "TODAY"
        || n = "OFFSET" || n = "INDIRECT"
        || List.exists containsVolatile args
    | UnaryOp (_, e) -> containsVolatile e
    | BinOp (l, _, r) -> containsVolatile l || containsVolatile r
    | _ -> false

/// Resolve the model a reference targets. Unqualified refs use the current
/// model; sheet-qualified refs resolve through the workbook. Returns None
/// when a sheet name can't be resolved (caller surfaces #REF!).
let private resolveModel (sheet: string option) (model: SpreadsheetModel)
                         (workbook: RedBat.Spreadsheet.Workbook.Workbook option)
                         : SpreadsheetModel option =
    match sheet with
    | None -> Some model
    | Some name ->
        match workbook with
        | None -> None
        | Some wb ->
            RedBat.Spreadsheet.Workbook.resolveSheetName name wb
            |> Option.bind (fun sid -> Map.tryFind sid wb.Sheets)
            |> Option.map (fun s -> s.Model)

/// Read a single cell's value, resolving cross-sheet references.
let private readCell (r: RefAst) (model: SpreadsheetModel)
                     (workbook: RedBat.Spreadsheet.Workbook.Workbook option) : CellValue =
    match resolveModel r.Sheet model workbook with
    | Some m -> getCellValue { Row = r.Row; Col = r.Col } m |> unwrapValue
    | None -> err Err.ref

/// Expand a RangeRef to the flat list of unwrapped cell values, row-major.
/// The sheet qualifier on the start ref selects the target sheet.
let expandRangeRef (a: RefAst) (b: RefAst) (model: SpreadsheetModel)
                   (workbook: RedBat.Spreadsheet.Workbook.Workbook option) : CellValue list =
    match resolveModel a.Sheet model workbook with
    | Some m ->
        let sr, sc, er, ec = normalizeRange a b
        [ for row in sr .. er do
            for col in sc .. ec ->
                getCellValue { Row = row; Col = col } m |> unwrapValue ]
    | None -> [ err Err.ref ]

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

// --- Date helpers -----------------------------------------------------------
// Excel serial dates: serial 1 = 1900-01-01. Excel's phantom 1900-02-29 (the
// 1900-leap-year bug) makes the simple offset exact for every date from
// 1900-03-01 (serial 61) onward, which covers every realistic spreadsheet.
let private excelEpoch = DateTime(1899, 12, 30)
let private serialToDate (s: float) : DateTime = excelEpoch.AddDays(s)
let private dateToSerial (d: DateTime) : float = (d - excelEpoch).TotalDays

/// Shared RNG for the volatile RAND / RANDBETWEEN functions.
let private rng = Random()

// --- Criteria parsing (SUMIF / COUNTIF / …) ---------------------------------
/// Translate Excel's `*` / `?` wildcards into a case-insensitive regex.
let private wildcardRegex (pattern: string) =
    let sb = System.Text.StringBuilder()
    sb.Append('^') |> ignore
    let mutable i = 0
    while i < pattern.Length do
        match pattern.[i] with
        | '~' when i + 1 < pattern.Length ->
            // ~* / ~? escape the wildcard.
            sb.Append(System.Text.RegularExpressions.Regex.Escape(string pattern.[i + 1])) |> ignore
            i <- i + 2
        | '*' -> sb.Append(".*") |> ignore; i <- i + 1
        | '?' -> sb.Append('.') |> ignore; i <- i + 1
        | c -> sb.Append(System.Text.RegularExpressions.Regex.Escape(string c)) |> ignore; i <- i + 1
    sb.Append('$') |> ignore
    System.Text.RegularExpressions.Regex(
        sb.ToString(), System.Text.RegularExpressions.RegexOptions.IgnoreCase)

/// Compile a criteria value (`">5"`, `"<>0"`, `"app*"`, `42`, …) into a
/// predicate over candidate cell values. Used by the *IF / *IFS family.
let private compileCriteria (criteria: CellValue) : CellValue -> bool =
    match unwrapValue criteria with
    | CellValue.Number cn ->
        fun v -> match coerceNumber v with Ok x -> x = cn | _ -> false
    | CellValue.Boolean cb ->
        fun v -> match coerceBool v with Ok x -> x = cb | _ -> false
    | CellValue.String raw ->
        let ops = [ ">="; "<="; "<>"; ">"; "<"; "=" ]
        let op, rest =
            match ops |> List.tryFind (fun o -> raw.StartsWith o) with
            | Some o -> o, raw.Substring(o.Length)
            | None -> "=", raw
        match Double.TryParse(rest, NumberStyles.Float, CultureInfo.InvariantCulture) with
        | true, num ->
            fun v ->
                match coerceNumber v with
                | Ok x ->
                    match op with
                    | ">=" -> x >= num | "<=" -> x <= num | ">" -> x > num
                    | "<" -> x < num | "<>" -> x <> num | _ -> x = num
                | _ -> false
        | _ ->
            match op with
            | ">" | "<" | ">=" | "<=" ->
                fun v ->
                    match coerceText v with
                    | Ok t ->
                        let c = String.Compare(t, rest, StringComparison.OrdinalIgnoreCase)
                        match op with
                        | ">" -> c > 0 | "<" -> c < 0 | ">=" -> c >= 0 | _ -> c <= 0
                    | _ -> false
            | "<>" ->
                let rx = wildcardRegex rest
                fun v -> match coerceText v with Ok t -> not (rx.IsMatch t) | _ -> true
            | _ ->
                let rx = wildcardRegex rest
                fun v -> match coerceText v with Ok t -> rx.IsMatch t | _ -> false
    | CellValue.Empty ->
        fun v -> match unwrapValue v with CellValue.Empty -> true | _ -> false
    | _ -> fun _ -> false

/// Static 0-indexed bounds of a range-shaped expression (RangeRef or a
/// single CellRef). None for anything else.
let private rangeBounds (e: Expr) : (int * int * int * int) option =
    match e with
    | RangeRef (a, b) -> Some (min a.Row b.Row, min a.Col b.Col, max a.Row b.Row, max a.Col b.Col)
    | CellRef r -> Some (r.Row, r.Col, r.Row, r.Col)
    | _ -> None

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

/// Everything an evaluation needs beyond the expression itself: the sheet
/// model being evaluated, the workbook (for cross-sheet refs and defined
/// names), and which sheet is "current" (for scoping sheet-local names).
type EvalContext = {
    Model: SpreadsheetModel
    Workbook: RedBat.Spreadsheet.Workbook.Workbook option
    Sheet: RedBat.Spreadsheet.Workbook.SheetId option
}

/// Build a context with no workbook — same-sheet formulas only, no defined
/// names. Convenient for callers that just have a bare model.
let contextOf (model: SpreadsheetModel) : EvalContext =
    { Model = model; Workbook = None; Sheet = None }

/// Evaluate an expression to a scalar CellValue.
let rec eval (expr: Expr) (ctx: EvalContext) : CellValue =
    let model = ctx.Model
    let workbook = ctx.Workbook
    match expr with
    | Number n -> CellValue.Number n
    | Str s -> CellValue.String s
    | Bool b -> CellValue.Boolean b
    | ErrorLit e -> CellValue.Error e
    | NameRef name ->
        match workbook with
        | None -> err Err.name
        | Some wb ->
            match resolveNameExpr name ctx.Sheet wb with
            | Some e -> eval e ctx
            | None -> err Err.name
    | CellRef r -> readCell r model workbook
    | RangeRef (a, b) ->
        // A bare range in scalar context yields its first cell (Excel's
        // implicit-intersection collapse is close enough for our use).
        match expandRangeRef a b model workbook with
        | v :: _ -> v
        | [] -> err Err.ref
    | TableRef (t, c) ->
        match expandTableRef t c workbook model with
        | v :: _ -> v
        | [] -> err Err.ref
    | UnaryOp (op, operand) ->
        let v = eval operand ctx
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
        let lv = eval left ctx
        let rv = eval right ctx
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
    | Function (name, args) -> evalFunction name args ctx

/// Evaluate every argument, flattening ranges/tables into their cell values
/// (the form most aggregate functions want).
and flattenArgs (args: Expr list) (ctx: EvalContext) : CellValue list =
    let model = ctx.Model
    let workbook = ctx.Workbook
    args |> List.collect (fun arg ->
        match arg with
        | RangeRef (a, b) -> expandRangeRef a b model workbook
        | TableRef (t, c) -> expandTableRef t c workbook model
        | NameRef name ->
            // A name standing in for a range must flatten like the range
            // would — otherwise SUM(MyRange) would see only its first cell.
            match workbook |> Option.bind (resolveNameExpr name ctx.Sheet) with
            | Some (RangeRef (a, b)) -> expandRangeRef a b model workbook
            | Some (TableRef (t, c)) -> expandTableRef t c workbook model
            | Some e -> [ eval e ctx ]
            | None -> [ err Err.name ]
        | _ -> [ eval arg ctx ])

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

/// Read a cell straight from the model (used by lookup functions that
/// index into a range rectangle).
and private gridCell (model: SpreadsheetModel) row col : CellValue =
    getCellValue { Row = row; Col = col } model |> unwrapValue

and evalFunction (name: string) (args: Expr list) (ctx: EvalContext) : CellValue =
    let model = ctx.Model
    let workbook = ctx.Workbook
    let nargs = List.length args
    let argRaw i = eval (List.item i args) ctx
    let argNum i = coerceNumber (argRaw i)
    let argText i = coerceText (argRaw i)
    let argBool i = coerceBool (argRaw i)
    let aggregateNumbers () = numbersOf (flattenArgs args ctx)
    // Run f over the aggregate numeric values; common shape for math
    // reductions. Empty list → emptyResult.
    let reduceNums emptyResult f =
        match aggregateNumbers () with
        | Result.Error e -> err e
        | Ok [] -> emptyResult
        | Ok ns -> f ns
    // Apply a unary numeric function to arg 0 with error propagation.
    let unaryNum f =
        match argNum 0 with Ok x -> f x | Result.Error e -> err e
    // (criteriaRangeExpr, criteriaValue) → the predicate-matched subset of
    // a parallel range, as a flat row-major list. Used by *IF / *IFS.
    let rangeValues (e: Expr) =
        match rangeBounds e with
        | Some (sr, sc, er, ec) ->
            [ for r in sr .. er do for c in sc .. ec -> gridCell model r c ]
        | None -> [ eval e ctx ]
    // Median of a sorted float list.
    let median (ns: float list) =
        let a = ns |> List.sort |> Array.ofList
        let n = a.Length
        if n = 0 then nan
        elif n % 2 = 1 then a.[n / 2]
        else (a.[n / 2 - 1] + a.[n / 2]) / 2.0
    let isErr v = match unwrapValue v with CellValue.Error _ -> true | _ -> false

    match name.ToUpperInvariant() with
    // ----- aggregates -------------------------------------------------------
    | "SUM" -> reduceNums (CellValue.Number 0.0) (fun ns -> CellValue.Number (List.sum ns))
    | "PRODUCT" -> reduceNums (CellValue.Number 0.0) (fun ns -> CellValue.Number (List.fold (*) 1.0 ns))
    | "AVERAGE" -> reduceNums (err Err.div0) (fun ns -> CellValue.Number (List.average ns))
    | "MIN" -> reduceNums (CellValue.Number 0.0) (fun ns -> CellValue.Number (List.min ns))
    | "MAX" -> reduceNums (CellValue.Number 0.0) (fun ns -> CellValue.Number (List.max ns))
    | "MEDIAN" -> reduceNums (err Err.num) (fun ns -> CellValue.Number (median ns))
    | "COUNT" ->
        let vals = flattenArgs args ctx
        match firstError vals with
        | Some (CellValue.Error e) -> err e
        | _ -> CellValue.Number (float (vals |> List.sumBy (fun v -> match unwrapValue v with CellValue.Number _ -> 1 | _ -> 0)))
    | "COUNTA" ->
        let vals = flattenArgs args ctx
        CellValue.Number (float (vals |> List.sumBy (fun v -> match unwrapValue v with CellValue.Empty -> 0 | _ -> 1)))
    | "COUNTBLANK" ->
        let vals = flattenArgs args ctx
        CellValue.Number (float (vals |> List.sumBy (fun v ->
            match unwrapValue v with
            | CellValue.Empty -> 1
            | CellValue.String "" -> 1
            | _ -> 0)))
    | "STDEV" | "STDEVP" ->
        reduceNums (err Err.div0) (fun ns ->
            let n = List.length ns
            if (name.ToUpperInvariant() = "STDEV" && n < 2) then err Err.div0
            else
                let mean = List.average ns
                let ss = ns |> List.sumBy (fun x -> (x - mean) ** 2.0)
                let denom = if name.ToUpperInvariant() = "STDEV" then float (n - 1) else float n
                CellValue.Number (sqrt (ss / denom)))
    | "VAR" ->
        reduceNums (err Err.div0) (fun ns ->
            let n = List.length ns
            if n < 2 then err Err.div0
            else
                let mean = List.average ns
                CellValue.Number ((ns |> List.sumBy (fun x -> (x - mean) ** 2.0)) / float (n - 1)))
    | "LARGE" | "SMALL" ->
        if nargs <> 2 then err Err.value else
        match numbersOf (flattenArgs [ List.item 0 args ] ctx), argNum 1 with
        | Result.Error e, _ -> err e
        | _, Result.Error e -> err e
        | Ok ns, Ok k ->
            let sorted =
                if name.ToUpperInvariant() = "LARGE" then List.sortDescending ns
                else List.sort ns
            let ki = int k - 1
            if ki < 0 || ki >= List.length sorted then err Err.num
            else CellValue.Number sorted.[ki]
    // ----- criteria aggregates ---------------------------------------------
    | "COUNTIF" ->
        if nargs <> 2 then err Err.value else
        let pred = compileCriteria (argRaw 1)
        CellValue.Number (float (rangeValues (List.item 0 args) |> List.filter pred |> List.length))
    | "SUMIF" ->
        // SUMIF(range, criteria, [sum_range])
        if nargs < 2 then err Err.value else
        let critRange = rangeValues (List.item 0 args)
        let sumRange =
            if nargs >= 3 then rangeValues (List.item 2 args) else critRange
        let pred = compileCriteria (argRaw 1)
        let total =
            List.zip
                (critRange |> List.truncate (List.length sumRange))
                (sumRange |> List.truncate (List.length critRange))
            |> List.sumBy (fun (c, s) ->
                if pred c then (match coerceNumber s with Ok n -> n | _ -> 0.0) else 0.0)
        CellValue.Number total
    | "AVERAGEIF" ->
        if nargs < 2 then err Err.value else
        let critRange = rangeValues (List.item 0 args)
        let avgRange = if nargs >= 3 then rangeValues (List.item 2 args) else critRange
        let pred = compileCriteria (argRaw 1)
        let matched =
            List.zip
                (critRange |> List.truncate (List.length avgRange))
                (avgRange |> List.truncate (List.length critRange))
            |> List.choose (fun (c, v) ->
                if pred c then (match coerceNumber v with Ok n -> Some n | _ -> None) else None)
        if List.isEmpty matched then err Err.div0
        else CellValue.Number (List.average matched)
    | "SUMIFS" | "COUNTIFS" ->
        // SUMIFS(sum_range, crit_range1, crit1, crit_range2, crit2, …)
        // COUNTIFS(crit_range1, crit1, …)  — no sum range.
        let isSum = name.ToUpperInvariant() = "SUMIFS"
        let pairsStart = if isSum then 1 else 0
        let pairArgs = args |> List.skip pairsStart
        if List.length pairArgs % 2 <> 0 || List.isEmpty pairArgs then err Err.value else
        let pairs =
            pairArgs
            |> List.chunkBySize 2
            |> List.map (fun chunk ->
                match chunk with
                | [ rangeExpr; critExpr ] ->
                    rangeValues rangeExpr, compileCriteria (eval critExpr ctx)
                | _ -> [], (fun _ -> false))
        let count = pairs |> List.map (fst >> List.length) |> List.min
        let sumRange = if isSum then rangeValues (List.item 0 args) else []
        let mutable total = 0.0
        let mutable hits = 0
        for i in 0 .. count - 1 do
            let allMatch = pairs |> List.forall (fun (vals, pred) -> pred (List.item i vals))
            if allMatch then
                hits <- hits + 1
                if isSum && i < List.length sumRange then
                    match coerceNumber (List.item i sumRange) with
                    | Ok n -> total <- total + n
                    | _ -> ()
        if isSum then CellValue.Number total else CellValue.Number (float hits)
    | "SUMPRODUCT" ->
        // Element-wise product of parallel ranges, then sum.
        let cols = args |> List.map (fun a -> numbersOf (rangeValues a))
        match cols |> List.tryPick (function Result.Error e -> Some e | _ -> None) with
        | Some e -> err e
        | None ->
            let arrs = cols |> List.choose (function Ok ns -> Some ns | _ -> None)
            match arrs with
            | [] -> CellValue.Number 0.0
            | first :: _ ->
                let len = arrs |> List.map List.length |> List.min
                let mutable total = 0.0
                for i in 0 .. len - 1 do
                    total <- total + (arrs |> List.fold (fun acc a -> acc * List.item i a) 1.0)
                CellValue.Number total
    // ----- math -------------------------------------------------------------
    | "ABS" -> unaryNum (abs >> CellValue.Number)
    | "SQRT" -> unaryNum (fun x -> if x < 0.0 then err Err.num else CellValue.Number (sqrt x))
    | "SIGN" -> unaryNum (fun x -> CellValue.Number (float (sign x)))
    | "INT" -> unaryNum (fun x -> CellValue.Number (floor x))
    | "TRUNC" ->
        match argNum 0 with
        | Result.Error e -> err e
        | Ok x ->
            let digits = if nargs >= 2 then (match argNum 1 with Ok d -> int d | _ -> 0) else 0
            let factor = 10.0 ** float digits
            CellValue.Number (Math.Truncate(x * factor) / factor)
    | "ROUND" | "ROUNDUP" | "ROUNDDOWN" ->
        if nargs < 1 then err Err.value else
        match argNum 0, (if nargs >= 2 then argNum 1 else Ok 0.0) with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok x, Ok d ->
            let factor = 10.0 ** d
            let scaled = x * factor
            let rounded =
                match name.ToUpperInvariant() with
                | "ROUNDUP" -> (if scaled < 0.0 then floor scaled else ceil scaled)
                | "ROUNDDOWN" -> Math.Truncate scaled
                | _ -> Math.Round(scaled, MidpointRounding.AwayFromZero)
            CellValue.Number (rounded / factor)
    | "MOD" ->
        if nargs <> 2 then err Err.value else
        match argNum 0, argNum 1 with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok _, Ok d when d = 0.0 -> err Err.div0
        | Ok a, Ok d ->
            // Excel MOD takes the sign of the divisor.
            let r = a - d * floor (a / d)
            CellValue.Number r
    | "POWER" ->
        if nargs <> 2 then err Err.value else
        match argNum 0, argNum 1 with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok b, Ok p ->
            let r = b ** p
            if Double.IsNaN r || Double.IsInfinity r then err Err.num else CellValue.Number r
    // ----- logical ----------------------------------------------------------
    | "IF" ->
        match args with
        | [ c; t ] ->
            match coerceBool (eval c ctx) with
            | Ok true -> eval t ctx
            | Ok false -> CellValue.Boolean false
            | Result.Error e -> err e
        | [ c; t; f ] ->
            match coerceBool (eval c ctx) with
            | Ok true -> eval t ctx
            | Ok false -> eval f ctx
            | Result.Error e -> err e
        | _ -> err Err.value
    | "AND" | "OR" | "XOR" ->
        let bools =
            flattenArgs args ctx
            |> List.choose (fun v ->
                match unwrapValue v with
                | CellValue.Empty -> None
                | other -> Some (coerceBool other))
        match bools |> List.tryPick (function Result.Error e -> Some e | _ -> None) with
        | Some e -> err e
        | None ->
            let bs = bools |> List.choose (function Ok b -> Some b | _ -> None)
            if List.isEmpty bs then err Err.value
            else
                match name.ToUpperInvariant() with
                | "AND" -> CellValue.Boolean (List.forall id bs)
                | "OR" -> CellValue.Boolean (List.exists id bs)
                | _ -> CellValue.Boolean (bs |> List.filter id |> List.length |> fun n -> n % 2 = 1)
    | "NOT" ->
        if nargs <> 1 then err Err.value else
        match argBool 0 with Ok b -> CellValue.Boolean (not b) | Result.Error e -> err e
    | "IFERROR" ->
        if nargs <> 2 then err Err.value else
        let v = argRaw 0
        if isErr v then argRaw 1 else v
    | "IFNA" ->
        if nargs <> 2 then err Err.value else
        match unwrapValue (argRaw 0) with
        | CellValue.Error e when e = Err.na -> argRaw 1
        | v -> v
    | "IFS" ->
        // IFS(cond1, val1, cond2, val2, …)
        if nargs < 2 || nargs % 2 <> 0 then err Err.value else
        let rec loop i =
            if i >= nargs then err Err.na
            else
                match coerceBool (argRaw i) with
                | Result.Error e -> err e
                | Ok true -> argRaw (i + 1)
                | Ok false -> loop (i + 2)
        loop 0
    | "SWITCH" ->
        // SWITCH(expr, case1, val1, …, [default])
        if nargs < 3 then err Err.value else
        let target = argRaw 0
        let rec loop i =
            if i + 1 >= nargs then
                // odd trailing arg = default
                if i < nargs then argRaw i else err Err.na
            else
                match compareValues target (argRaw i) with
                | Ok 0 -> argRaw (i + 1)
                | _ -> loop (i + 2)
        loop 1
    | "TRUE" -> CellValue.Boolean true
    | "FALSE" -> CellValue.Boolean false
    | "NA" -> err Err.na
    // ----- info -------------------------------------------------------------
    | "ISBLANK" ->
        if nargs <> 1 then err Err.value
        else CellValue.Boolean (match unwrapValue (argRaw 0) with CellValue.Empty -> true | _ -> false)
    | "ISERROR" ->
        if nargs <> 1 then err Err.value else CellValue.Boolean (isErr (argRaw 0))
    | "ISNA" ->
        if nargs <> 1 then err Err.value
        else CellValue.Boolean (match unwrapValue (argRaw 0) with CellValue.Error e when e = Err.na -> true | _ -> false)
    | "ISNUMBER" ->
        if nargs <> 1 then err Err.value
        else CellValue.Boolean (match unwrapValue (argRaw 0) with CellValue.Number _ -> true | _ -> false)
    | "ISTEXT" ->
        if nargs <> 1 then err Err.value
        else CellValue.Boolean (match unwrapValue (argRaw 0) with CellValue.String _ -> true | _ -> false)
    | "ISLOGICAL" ->
        if nargs <> 1 then err Err.value
        else CellValue.Boolean (match unwrapValue (argRaw 0) with CellValue.Boolean _ -> true | _ -> false)
    // ----- text -------------------------------------------------------------
    | "CONCAT" | "CONCATENATE" ->
        let parts = flattenArgs args ctx |> List.map coerceText
        match parts |> List.tryPick (function Result.Error e -> Some e | _ -> None) with
        | Some e -> err e
        | None -> CellValue.String (parts |> List.choose (function Ok s -> Some s | _ -> None) |> String.concat "")
    | "TEXTJOIN" ->
        // TEXTJOIN(delimiter, ignore_empty, text…)
        if nargs < 3 then err Err.value else
        match argText 0, argBool 1 with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok delim, Ok ignoreEmpty ->
            let parts =
                flattenArgs (args |> List.skip 2) ctx
                |> List.choose (fun v ->
                    match unwrapValue v with
                    | CellValue.Empty when ignoreEmpty -> None
                    | other -> match coerceText other with Ok s -> Some s | _ -> Some "")
            CellValue.String (String.Join(delim, parts))
    | "LEN" ->
        if nargs <> 1 then err Err.value
        else match argText 0 with Ok s -> CellValue.Number (float s.Length) | Result.Error e -> err e
    | "TRIM" ->
        if nargs <> 1 then err Err.value
        else match argText 0 with
             | Ok s ->
                // Excel TRIM collapses runs of interior spaces to one.
                let collapsed = System.Text.RegularExpressions.Regex.Replace(s.Trim(), " +", " ")
                CellValue.String collapsed
             | Result.Error e -> err e
    | "UPPER" ->
        if nargs <> 1 then err Err.value
        else match argText 0 with Ok s -> CellValue.String (s.ToUpperInvariant()) | Result.Error e -> err e
    | "LOWER" ->
        if nargs <> 1 then err Err.value
        else match argText 0 with Ok s -> CellValue.String (s.ToLowerInvariant()) | Result.Error e -> err e
    | "PROPER" ->
        if nargs <> 1 then err Err.value
        else match argText 0 with
             | Ok s -> CellValue.String (CultureInfo.InvariantCulture.TextInfo.ToTitleCase(s.ToLowerInvariant()))
             | Result.Error e -> err e
    | "LEFT" ->
        if nargs < 1 then err Err.value else
        match argText 0 with
        | Result.Error e -> err e
        | Ok s ->
            let cnt = if nargs >= 2 then (match argNum 1 with Ok n -> int n | _ -> 1) else 1
            CellValue.String (s.Substring(0, max 0 (min s.Length cnt)))
    | "RIGHT" ->
        if nargs < 1 then err Err.value else
        match argText 0 with
        | Result.Error e -> err e
        | Ok s ->
            let cnt = if nargs >= 2 then (match argNum 1 with Ok n -> int n | _ -> 1) else 1
            let c = max 0 (min s.Length cnt)
            CellValue.String (s.Substring(s.Length - c, c))
    | "MID" ->
        if nargs <> 3 then err Err.value else
        match argText 0, argNum 1, argNum 2 with
        | Result.Error e, _, _ | _, Result.Error e, _ | _, _, Result.Error e -> err e
        | Ok s, Ok startN, Ok lenN ->
            let start0 = int startN - 1   // Excel MID is 1-indexed
            if start0 < 0 then err Err.value
            elif start0 >= s.Length then CellValue.String ""
            else
                let take = max 0 (min (s.Length - start0) (int lenN))
                CellValue.String (s.Substring(start0, take))
    | "REPT" ->
        if nargs <> 2 then err Err.value else
        match argText 0, argNum 1 with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok s, Ok cnt -> CellValue.String (String.replicate (max 0 (int cnt)) s)
    | "FIND" | "SEARCH" ->
        // FIND is case-sensitive, SEARCH is not. Both 1-indexed.
        if nargs < 2 then err Err.value else
        match argText 0, argText 1 with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok needle, Ok hay ->
            let startAt = if nargs >= 3 then (match argNum 2 with Ok n -> int n - 1 | _ -> 0) else 0
            if startAt < 0 || startAt > hay.Length then err Err.value
            else
                let cmp =
                    if name.ToUpperInvariant() = "FIND" then StringComparison.Ordinal
                    else StringComparison.OrdinalIgnoreCase
                let idx = hay.IndexOf(needle, startAt, cmp)
                if idx < 0 then err Err.value else CellValue.Number (float (idx + 1))
    | "SUBSTITUTE" ->
        // SUBSTITUTE(text, old, new, [instance])
        if nargs < 3 then err Err.value else
        match argText 0, argText 1, argText 2 with
        | Result.Error e, _, _ | _, Result.Error e, _ | _, _, Result.Error e -> err e
        | Ok s, Ok oldT, Ok newT ->
            if oldT = "" then CellValue.String s
            elif nargs >= 4 then
                match argNum 3 with
                | Result.Error e -> err e
                | Ok inst ->
                    // Replace only the inst-th occurrence.
                    let target = int inst
                    let mutable count = 0
                    let mutable idx = 0
                    let sb = System.Text.StringBuilder()
                    while idx < s.Length do
                        if idx + oldT.Length <= s.Length && s.Substring(idx, oldT.Length) = oldT then
                            count <- count + 1
                            if count = target then sb.Append(newT) |> ignore
                            else sb.Append(oldT) |> ignore
                            idx <- idx + oldT.Length
                        else
                            sb.Append(s.[idx]) |> ignore
                            idx <- idx + 1
                    CellValue.String (sb.ToString())
            else
                CellValue.String (s.Replace(oldT, newT))
    | "REPLACE" ->
        // REPLACE(text, start, num_chars, new_text) — 1-indexed start.
        if nargs <> 4 then err Err.value else
        match argText 0, argNum 1, argNum 2, argText 3 with
        | Result.Error e, _, _, _ | _, Result.Error e, _, _
        | _, _, Result.Error e, _ | _, _, _, Result.Error e -> err e
        | Ok s, Ok startN, Ok lenN, Ok newT ->
            let start0 = max 0 (int startN - 1)
            let cnt = max 0 (int lenN)
            if start0 > s.Length then CellValue.String (s + newT)
            else
                let before = s.Substring(0, start0)
                let after = if start0 + cnt < s.Length then s.Substring(start0 + cnt) else ""
                CellValue.String (before + newT + after)
    | "TEXT" ->
        // Minimal TEXT — passes the format through .NET. Common Excel codes
        // (0, 0.00, #,##0, 0%) are .NET-compatible enough for the usual case.
        if nargs <> 2 then err Err.value else
        match argNum 0, argText 1 with
        | Result.Error e, _ -> err e
        | _, Result.Error e -> err e
        | Ok n, Ok fmt ->
            try CellValue.String (n.ToString(fmt, CultureInfo.InvariantCulture))
            with _ -> CellValue.String (numberToText n)
    // ----- lookup -----------------------------------------------------------
    | "CHOOSE" ->
        if nargs < 2 then err Err.value else
        match argNum 0 with
        | Result.Error e -> err e
        | Ok idx ->
            let i = int idx
            if i < 1 || i >= nargs then err Err.value else argRaw i
    | "INDEX" ->
        // INDEX(range, row_num, [col_num]) — 1-indexed; 0 = whole row/col,
        // collapsed here to the first cell of that line.
        if nargs < 2 then err Err.value else
        match rangeBounds (List.item 0 args) with
        | None -> err Err.ref
        | Some (sr, sc, er, ec) ->
            match argNum 1 with
            | Result.Error e -> err e
            | Ok rowN ->
                let colN = if nargs >= 3 then (match argNum 2 with Ok c -> int c | _ -> 1) else 1
                let rows = er - sr + 1
                let csz = ec - sc + 1
                // Single-row or single-col range: the lone index walks it.
                let r, c =
                    if rows = 1 then sr, sc + (max 1 (int rowN)) - 1
                    elif csz = 1 then sr + (max 1 (int rowN)) - 1, sc
                    else sr + (max 1 (int rowN)) - 1, sc + (max 1 colN) - 1
                if r < sr || r > er || c < sc || c > ec then err Err.ref
                else gridCell model r c
    | "MATCH" ->
        // MATCH(lookup, range, [match_type]) — 1=≤(sorted asc), 0=exact,
        // -1=≥(sorted desc). Defaults to 1.
        if nargs < 2 then err Err.value else
        match rangeBounds (List.item 1 args) with
        | None -> err Err.ref
        | Some (sr, sc, er, ec) ->
            let target = argRaw 0
            let matchType = if nargs >= 3 then (match argNum 2 with Ok m -> int m | _ -> 1) else 1
            let cells =
                [ for r in sr .. er do for c in sc .. ec -> gridCell model r c ]
            let result =
                if matchType = 0 then
                    cells |> List.tryFindIndex (fun v -> match compareValues target v with Ok 0 -> true | _ -> false)
                elif matchType = 1 then
                    // largest value ≤ target, range assumed ascending
                    cells
                    |> List.mapi (fun i v -> i, v)
                    |> List.filter (fun (_, v) -> match compareValues v target with Ok c -> c <= 0 | _ -> false)
                    |> List.tryLast
                    |> Option.map fst
                else
                    // smallest value ≥ target, range assumed descending
                    cells
                    |> List.mapi (fun i v -> i, v)
                    |> List.filter (fun (_, v) -> match compareValues v target with Ok c -> c >= 0 | _ -> false)
                    |> List.tryLast
                    |> Option.map fst
            match result with
            | Some i -> CellValue.Number (float (i + 1))
            | None -> err Err.na
    | "VLOOKUP" | "HLOOKUP" ->
        // VLOOKUP(lookup, table, col_index, [range_lookup])
        // HLOOKUP(lookup, table, row_index, [range_lookup])
        if nargs < 3 then err Err.value else
        match rangeBounds (List.item 1 args) with
        | None -> err Err.ref
        | Some (sr, sc, er, ec) ->
            let isV = name.ToUpperInvariant() = "VLOOKUP"
            let target = argRaw 0
            match argNum 2 with
            | Result.Error e -> err e
            | Ok idxN ->
                let idx = int idxN
                let approx = if nargs >= 4 then (match argBool 3 with Ok b -> b | _ -> true) else true
                // search line: first column (V) or first row (H)
                let searchLine =
                    if isV then [ for r in sr .. er -> r, gridCell model r sc ]
                    else [ for c in sc .. ec -> c, gridCell model sr c ]
                let hitLine =
                    if approx then
                        // largest value ≤ target (assumes sorted ascending)
                        searchLine
                        |> List.filter (fun (_, v) -> match compareValues v target with Ok c -> c <= 0 | _ -> false)
                        |> List.tryLast
                    else
                        searchLine
                        |> List.tryFind (fun (_, v) -> match compareValues target v with Ok 0 -> true | _ -> false)
                match hitLine with
                | None -> err Err.na
                | Some (lineIdx, _) ->
                    if isV then
                        let col = sc + idx - 1
                        if col < sc || col > ec then err Err.ref else gridCell model lineIdx col
                    else
                        let row = sr + idx - 1
                        if row < sr || row > er then err Err.ref else gridCell model row lineIdx
    // ----- date -------------------------------------------------------------
    | "TODAY" -> CellValue.Number (dateToSerial DateTime.Today)
    | "NOW" -> CellValue.Number (dateToSerial DateTime.Now)
    | "DATE" ->
        if nargs <> 3 then err Err.value else
        match argNum 0, argNum 1, argNum 2 with
        | Result.Error e, _, _ | _, Result.Error e, _ | _, _, Result.Error e -> err e
        | Ok y, Ok m, Ok d ->
            try
                let baseDate = DateTime(int y, 1, 1)
                let dt = baseDate.AddMonths(int m - 1).AddDays(d - 1.0)
                CellValue.Number (dateToSerial dt)
            with _ -> err Err.num
    | "YEAR" -> unaryNum (fun s -> CellValue.Number (float (serialToDate s).Year))
    | "MONTH" -> unaryNum (fun s -> CellValue.Number (float (serialToDate s).Month))
    | "DAY" -> unaryNum (fun s -> CellValue.Number (float (serialToDate s).Day))
    | "WEEKDAY" ->
        match argNum 0 with
        | Result.Error e -> err e
        | Ok s ->
            // Default type 1: Sunday = 1 … Saturday = 7.
            let dow = int (serialToDate s).DayOfWeek  // Sunday = 0
            CellValue.Number (float (dow + 1))
    | "EOMONTH" ->
        if nargs <> 2 then err Err.value else
        match argNum 0, argNum 1 with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok s, Ok months ->
            let d = (serialToDate s).AddMonths(int months)
            let eom = DateTime(d.Year, d.Month, DateTime.DaysInMonth(d.Year, d.Month))
            CellValue.Number (dateToSerial eom)
    // ----- volatile --------------------------------------------------------
    | "RAND" -> CellValue.Number (rng.NextDouble())
    | "RANDBETWEEN" ->
        if nargs <> 2 then err Err.value else
        match argNum 0, argNum 1 with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok lo, Ok hi ->
            let l, h = int (ceil lo), int (floor hi)
            if l > h then err Err.num
            else CellValue.Number (float (rng.Next(l, h + 1)))
    | "OFFSET" ->
        // OFFSET(reference, rows, cols, [height], [width]) — scalar context
        // collapses to the top-left cell of the resulting block.
        if nargs < 3 then err Err.value else
        let anchor =
            match List.item 0 args with
            | CellRef r -> Some (r.Sheet, r.Row, r.Col)
            | RangeRef (a, _) -> Some (a.Sheet, a.Row, a.Col)
            | _ -> None
        match anchor with
        | None -> err Err.ref
        | Some (sheet, br, bc) ->
            match argNum 1, argNum 2 with
            | Result.Error e, _ | _, Result.Error e -> err e
            | Ok dr, Ok dc ->
                let r = br + int dr
                let c = bc + int dc
                if r < 0 || c < 0 then err Err.ref
                else
                    match resolveModel sheet model workbook with
                    | Some m -> getCellValue { Row = r; Col = c } m |> unwrapValue
                    | None -> err Err.ref
    | "INDIRECT" ->
        // INDIRECT(ref_text) — parse the text as an A1 reference and read it.
        if nargs < 1 then err Err.value else
        match argText 0 with
        | Result.Error e -> err e
        | Ok refText ->
            match parseFormula ("=" + refText) with
            | Ok (CellRef _ as e) | Ok (RangeRef _ as e) -> eval e ctx
            | _ -> err Err.ref
    | _ ->
        // Unknown / unsupported function.
        err Err.name

/// Parse + evaluate a formula string in one shot. Defined names resolve at
/// the workbook's active sheet — the right scope for an interactive
/// edit-buffer preview, which is what this entry point serves.
let evaluateFormula (formula: string) (model: SpreadsheetModel)
                    (workbook: RedBat.Spreadsheet.Workbook.Workbook option) : CellValue =
    match parseFormula formula with
    | Ok expr ->
        let ctx =
            { Model = model
              Workbook = workbook
              Sheet = workbook |> Option.map (fun w -> w.ActiveSheet) }
        eval expr ctx
    | Result.Error msg -> CellValue.Error msg

// ---------------------------------------------------------------------------
// Whole-model recalculation (dependency-ordered). Calc.fs owns the
// incremental engine; this is the simple full pass the shared model uses.
// ---------------------------------------------------------------------------

/// Recalculate every formula cell in a model, in dependency order.
/// `currentSheet` scopes defined-name resolution to the sheet being
/// recalculated (the caller knows which sheet `model` belongs to).
let recalculateAllFormulas (workbook: RedBat.Spreadsheet.Workbook.Workbook option)
                           (currentSheet: RedBat.Spreadsheet.Workbook.SheetId option)
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
                    referencesInExpr workbook expr
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
        | Some (cell, formula, expr, None) ->
            let ctx = { Model = updated; Workbook = workbook; Sheet = currentSheet }
            let result = eval expr ctx
            updated <- setCell addr { cell with Value = CellValue.Formula (formula, result) } updated
        | None -> ()
    updated
