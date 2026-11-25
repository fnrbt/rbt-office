/// The Web functions of the function table: FILTERXML, on an XML parser and
/// an XPath 1.0 evaluator written here (System.Xml is not in the browser
/// build), so that the .NET and browser builds agree.
///
/// Excel 2019's rules, measured with tools/excel-parity/make-filterxml-probe.py
/// (Excel reads the XML with MSXML):
///  - The XML must be one well-formed element; a DOCTYPE, a second root, a
///    duplicate attribute, an unknown entity are #VALUE!. Whitespace-only text
///    is dropped as MSXML drops it. Names are matched as written, prefix and
///    all: no namespaces are resolved (`//b` finds the b of xmlns='urn:x',
///    `//p:b` the p:b).
///  - The expression is evaluated with the document as the context node (so
///    `r/a` finds the a children of the root r). Only a node-set is a result:
///    a number, string or boolean (count(//a), name(/a)) is #VALUE!, and so is
///    an empty node-set.
///  - A node's value is, for an element, its own text children joined (not
///    its descendants': <a>1<b>2</b>3</a> is 13), for an attribute its value,
///    trimmed of XML whitespace; for any other node, or an empty value,
///    #VALUE!. The value is then read as if typed into a cell: numbers,
///    percentages, currency, dates, TRUE (but #N/A, an error's text, stays
///    text).
///  - The nodes are a column in document order. (Excel 2019 gives one node
///    as a single value rather than a 1x1 array, which only ROWS, COLUMNS
///    and INDEX of a text or logical tell apart: #VALUE! there.)
module Rbt.Spreadsheet.Functions.FnWeb

open System
open System.Collections.Generic
open Rbt.Spreadsheet
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
open Rbt.Spreadsheet.Functions.Kit

let moduleName = "Web"

// ===== the XML tree ==========================================================

type NodeKind =
    | Document = 0
    | Element = 1
    | Attribute = 2
    | Text = 3
    | Comment = 4
    | Instruction = 5

[<AllowNullLiteral>]
type XNode(kind: NodeKind, name: string, value: string, parent: XNode) =
    member _.Kind = kind
    member _.Name = name
    member _.Value = value
    member _.Parent = parent
    member val Children = ResizeArray<XNode>()
    member val Attributes = ResizeArray<XNode>()
    /// The position in document order.
    member val Order = 0 with get, set

exception XmlError of string

let private isXmlSpace (c: char) = c = ' ' || c = '\t' || c = '\n' || c = '\r'
let private isNameStart (c: char) = Char.IsLetter c || c = '_' || c = ':' || int c > 127
let private isNameChar (c: char) = isNameStart c || Char.IsDigit c || c = '-' || c = '.'
let private isName (s: string) = s.Length > 0 && isNameStart s.[0] && Seq.forall isNameChar s

/// The document node of `text`, or XmlError. The text is taken apart by the package readers' lexer
/// (XmlDsl.XmlLexer); what it lets through that MSXML refuses is refused here: a name that is none, an
/// attribute twice, text outside the root element, an XML declaration anywhere but first, a DOCTYPE.
let parseXml (source: string) : XNode =
    let fail () = raise (XmlError "malformed")
    let lexer = XmlDsl.XmlLexer(source, 0, source.Length)
    let doc = XNode(NodeKind.Document, "", "", null)
    let mutable current = doc
    let mutable rootSeen = false
    let mutable first = true
    let pending = Text.StringBuilder()
    let flush () =
        if pending.Length > 0 then
            let text = pending.ToString()
            pending.Clear() |> ignore
            // MSXML drops the text that is only whitespace.
            if not (Seq.forall isXmlSpace text) then current.Children.Add(XNode(NodeKind.Text, "", text, current))
    let raw (opening: int) (closing: int) = source.Substring(lexer.Start + opening, lexer.End - lexer.Start - opening - closing)
    let mutable go = true
    try
        while go do
            let token = lexer.Next ()
            let atFirst = first
            first <- false
            match token with
            | XmlDsl.XmlToken.StartTag ->
                flush ()
                if obj.ReferenceEquals(current, doc) && rootSeen then fail ()
                let name = lexer.Name
                if not (isName name) then fail ()
                let node = XNode(NodeKind.Element, name, "", current)
                current.Children.Add node
                let seen = HashSet<string>()
                for a in 0 .. lexer.AttributeCount - 1 do
                    let attr = lexer.AttributeName a
                    if not (isName attr) || not (seen.Add attr) then fail ()
                    node.Attributes.Add(XNode(NodeKind.Attribute, attr, lexer.AttributeValue a, node))
                if obj.ReferenceEquals(current, doc) then rootSeen <- true
                if not lexer.SelfClosing then current <- node
            | XmlDsl.XmlToken.EndTag ->
                flush ()
                current <- current.Parent
            | XmlDsl.XmlToken.Text ->
                if obj.ReferenceEquals(current, doc) then (if not lexer.TextIsBlank then fail ())
                else pending.Append(lexer.TextValue ()) |> ignore
            | XmlDsl.XmlToken.CData ->
                if obj.ReferenceEquals(current, doc) then fail ()
                flush ()
                current.Children.Add(XNode(NodeKind.Text, "", lexer.TextValue (), current))
            | XmlDsl.XmlToken.Comment ->
                flush ()
                current.Children.Add(XNode(NodeKind.Comment, "", (raw 4 3).Replace("\r\n", "\n").Replace('\r', '\n'), current))
            | XmlDsl.XmlToken.Instruction ->
                flush ()
                let body = (raw 2 2).Replace("\r\n", "\n").Replace('\r', '\n')
                let mutable k = 0
                while k < body.Length && not (isXmlSpace body.[k]) do k <- k + 1
                let target = body.Substring(0, k)
                if not (isName target) then fail ()
                if target.Equals("xml", StringComparison.OrdinalIgnoreCase) then
                    // The XML declaration: first in the text, and nowhere else.
                    if not (atFirst && lexer.Start = 0 && target = "xml") then fail ()
                else current.Children.Add(XNode(NodeKind.Instruction, target, body.Substring(k).TrimStart(), current))
            | XmlDsl.XmlToken.EndOfText -> go <- false
            | _ -> fail ()
    with
    | XmlError _ -> reraise ()
    | _ -> fail ()
    if not (obj.ReferenceEquals(current, doc)) || not rootSeen then fail ()
    // Document order.
    let mutable order = 0
    let rec number (x: XNode) =
        x.Order <- order
        order <- order + 1
        for a in x.Attributes do
            a.Order <- order
            order <- order + 1
        for c in x.Children do number c
    number doc
    doc

// ===== XPath 1.0 ===============================================================

exception XPathError of string

type Axis =
    | Child | Descendant | Parent | Ancestor | FollowingSibling | PrecedingSibling | Following | Preceding
    | AttributeAxis | Namespace | Self | DescendantOrSelf | AncestorOrSelf

type NodeTest =
    | AnyName
    | PrefixAny of string
    | QName of string
    | AnyNode
    | TextTest
    | CommentTest
    | InstructionTest of string option

type Expr =
    | Or of Expr * Expr
    | And of Expr * Expr
    | Compare of string * Expr * Expr
    | Arith of string * Expr * Expr
    | Negate of Expr
    | Union of Expr * Expr
    | Num of float
    | Lit of string
    | Call of string * Expr list
    /// A location path: from the root (true) or the context node, its steps.
    | Path of bool * Step list
    /// A filter expression with its predicates, then steps after it.
    | Filtered of Expr * Expr list * Step list

and Step = { Axis: Axis; Test: NodeTest; Predicates: Expr list }

type private Token =
    | TNum of float
    | TLit of string
    | TName of string
    | TOp of string

/// The tokens after which a name or * is a name test, not an operator.
let private operatorTokens =
    set [ "@"; "::"; "("; "["; ","; "/"; "//"; "|"; "+"; "-"; "="; "!="; "<"; "<="; ">"; ">="; "and"; "or"; "mod"; "div"; "*m" ]

let private tokenize (text: string) : Token list =
    let tokens = ResizeArray<Token>()
    let n = text.Length
    let mutable i = 0
    let fail () = raise (XPathError "syntax")
    // Whether the token before makes a name or * an operator.
    let operatorContext () =
        if tokens.Count = 0 then false
        else
            match tokens.[tokens.Count - 1] with
            | TOp op -> not (operatorTokens.Contains op)
            | _ -> true
    while i < n do
        let c = text.[i]
        if isXmlSpace c then i <- i + 1
        elif Char.IsDigit c || (c = '.' && i + 1 < n && Char.IsDigit text.[i + 1]) then
            let start = i
            while i < n && Char.IsDigit text.[i] do i <- i + 1
            if i < n && text.[i] = '.' then
                i <- i + 1
                while i < n && Char.IsDigit text.[i] do i <- i + 1
            tokens.Add(TNum (Double.Parse(text.Substring(start, i - start), Globalization.CultureInfo.InvariantCulture)))
        elif c = '"' || c = '\'' then
            let j = text.IndexOf(c, i + 1)
            if j < 0 then fail ()
            tokens.Add(TLit (text.Substring(i + 1, j - i - 1)))
            i <- j + 1
        elif isNameStart c && c <> ':' then
            let start = i
            while i < n && isNameChar text.[i] && text.[i] <> ':' do i <- i + 1
            // prefix:local or prefix:*
            if i + 1 < n && text.[i] = ':' && text.[i + 1] <> ':' then
                i <- i + 1
                if text.[i] = '*' then i <- i + 1
                elif isNameStart text.[i] then
                    while i < n && isNameChar text.[i] && text.[i] <> ':' do i <- i + 1
                else fail ()
            let name = text.Substring(start, i - start)
            if operatorContext () && (name = "and" || name = "or" || name = "mod" || name = "div") then tokens.Add(TOp name)
            else tokens.Add(TName name)
        else
            let two = if i + 1 < n then text.Substring(i, 2) else ""
            match two with
            | "//" | "::" | "!=" | "<=" | ">=" | ".." -> tokens.Add(TOp two); i <- i + 2
            | _ ->
                match c with
                | '*' -> tokens.Add(TOp (if operatorContext () then "*m" else "*")); i <- i + 1
                | '/' | '(' | ')' | '[' | ']' | '.' | '@' | ',' | '|' | '+' | '-' | '=' | '<' | '>' ->
                    tokens.Add(TOp (string c)); i <- i + 1
                | _ -> fail ()
    List.ofSeq tokens

let private axisOf (name: string) =
    match name with
    | "child" -> Child | "descendant" -> Descendant | "parent" -> Parent | "ancestor" -> Ancestor
    | "following-sibling" -> FollowingSibling | "preceding-sibling" -> PrecedingSibling
    | "following" -> Following | "preceding" -> Preceding | "attribute" -> AttributeAxis
    | "namespace" -> Namespace | "self" -> Self | "descendant-or-self" -> DescendantOrSelf
    | "ancestor-or-self" -> AncestorOrSelf
    | _ -> raise (XPathError "axis")

/// The expression of an XPath text, or XPathError.
let parseXPath (text: string) : Expr =
    let mutable toks = tokenize text
    let fail () = raise (XPathError "syntax")
    let peekTok () = match toks with t :: _ -> Some t | [] -> None
    let peek2 () = match toks with _ :: t :: _ -> Some t | _ -> None
    let advance () = match toks with _ :: rest -> toks <- rest | [] -> fail ()
    let isOp (o: string) = peekTok () = Some (TOp o)
    let expectOp (o: string) = if isOp o then advance () else fail ()
    let descendantStep = { Axis = DescendantOrSelf; Test = AnyNode; Predicates = [] }
    let rec orExpr () : Expr =
        let mutable left = andExpr ()
        while isOp "or" do (advance (); left <- Or (left, andExpr ()))
        left
    and andExpr () =
        let mutable left = equality ()
        while isOp "and" do (advance (); left <- And (left, equality ()))
        left
    and equality () =
        let mutable left = relational ()
        while isOp "=" || isOp "!=" do
            let op = match peekTok () with Some (TOp o) -> o | _ -> ""
            advance ()
            left <- Compare (op, left, relational ())
        left
    and relational () =
        let mutable left = additive ()
        while isOp "<" || isOp "<=" || isOp ">" || isOp ">=" do
            let op = match peekTok () with Some (TOp o) -> o | _ -> ""
            advance ()
            left <- Compare (op, left, additive ())
        left
    and additive () =
        let mutable left = multiplicative ()
        while isOp "+" || isOp "-" do
            let op = match peekTok () with Some (TOp o) -> o | _ -> ""
            advance ()
            left <- Arith (op, left, multiplicative ())
        left
    and multiplicative () =
        let mutable left = unary ()
        while isOp "*m" || isOp "div" || isOp "mod" do
            let op = match peekTok () with Some (TOp "*m") -> "*" | Some (TOp o) -> o | _ -> ""
            advance ()
            left <- Arith (op, left, unary ())
        left
    and unary () =
        if isOp "-" then (advance (); Negate (unary ()))
        else union ()
    and union () =
        let mutable left = pathExpr ()
        while isOp "|" do (advance (); left <- Union (left, pathExpr ()))
        left
    and pathExpr () =
        let startsFilter =
            match peekTok () with
            | Some (TNum _) | Some (TLit _) | Some (TOp "(") -> true
            | Some (TName name) ->
                peek2 () = Some (TOp "(") && not (name = "node" || name = "text" || name = "comment" || name = "processing-instruction")
            | _ -> false
        if startsFilter then
            let primary = primaryExpr ()
            let preds = predicates ()
            let steps =
                if isOp "/" then (advance (); relativePath ())
                elif isOp "//" then (advance (); descendantStep :: relativePath ())
                else []
            if preds.IsEmpty && steps.IsEmpty then primary else Filtered (primary, preds, steps)
        elif isOp "/" then
            advance ()
            let canStep =
                match peekTok () with
                | Some (TName _) | Some (TOp "*") | Some (TOp ".") | Some (TOp "..") | Some (TOp "@") -> true
                | _ -> false
            Path (true, (if canStep then relativePath () else []))
        elif isOp "//" then
            advance ()
            Path (true, descendantStep :: relativePath ())
        else Path (false, relativePath ())
    and relativePath () : Step list =
        let first = step ()
        let steps = ResizeArray [ first ]
        let mutable go = true
        while go do
            if isOp "/" then (advance (); steps.Add(step ()))
            elif isOp "//" then (advance (); steps.Add descendantStep; steps.Add(step ()))
            else go <- false
        List.ofSeq steps
    and step () : Step =
        if isOp "." then (advance (); { Axis = Self; Test = AnyNode; Predicates = [] })
        elif isOp ".." then (advance (); { Axis = Parent; Test = AnyNode; Predicates = [] })
        else
            let axis =
                if isOp "@" then (advance (); AttributeAxis)
                else
                    match peekTok (), peek2 () with
                    | Some (TName name), Some (TOp "::") -> advance (); advance (); axisOf name
                    | _ -> Child
            let test =
                match peekTok () with
                | Some (TOp "*") -> advance (); AnyName
                | Some (TName name) ->
                    advance ()
                    if isOp "(" && (name = "node" || name = "text" || name = "comment" || name = "processing-instruction") then
                        advance ()
                        let t =
                            match name with
                            | "node" -> AnyNode
                            | "text" -> TextTest
                            | "comment" -> CommentTest
                            | _ ->
                                match peekTok () with
                                | Some (TLit l) -> advance (); InstructionTest (Some l)
                                | _ -> InstructionTest None
                        expectOp ")"
                        t
                    elif name.EndsWith ":*" then PrefixAny (name.Substring(0, name.Length - 2))
                    else QName name
                | _ -> fail ()
            { Axis = axis; Test = test; Predicates = predicates () }
    and predicates () : Expr list =
        let preds = ResizeArray<Expr>()
        while isOp "[" do
            advance ()
            preds.Add(orExpr ())
            expectOp "]"
        List.ofSeq preds
    and primaryExpr () : Expr =
        match peekTok () with
        | Some (TNum x) -> advance (); Num x
        | Some (TLit l) -> advance (); Lit l
        | Some (TOp "(") ->
            advance ()
            let e = orExpr ()
            expectOp ")"
            e
        | Some (TName name) ->
            advance ()
            expectOp "("
            let args = ResizeArray<Expr>()
            if not (isOp ")") then
                args.Add(orExpr ())
                while isOp "," do (advance (); args.Add(orExpr ()))
            expectOp ")"
            Call (name, List.ofSeq args)
        | _ -> fail ()
    let e = orExpr ()
    if not toks.IsEmpty then fail ()
    e

type XValue =
    | NodeSet of XNode list
    | XNumber of float
    | XString of string
    | XBool of bool

let rec private stringValue (x: XNode) : string =
    match x.Kind with
    | NodeKind.Document | NodeKind.Element ->
        let sb = Text.StringBuilder()
        let rec walk (y: XNode) =
            for c in y.Children do
                if c.Kind = NodeKind.Text then sb.Append c.Value |> ignore
                elif c.Kind = NodeKind.Element then walk c
        walk x
        sb.ToString()
    | _ -> x.Value

/// XPath's number(string): optional whitespace, an optional minus, digits with
/// an optional fraction; anything else NaN.
let private stringToNumber (s: string) =
    let t = s.Trim([| ' '; '\t'; '\n'; '\r' |])
    let body = if t.StartsWith "-" then t.Substring 1 else t
    let valid =
        body.Length > 0 && body <> "."
        && Seq.forall (fun c -> Char.IsDigit c || c = '.') body
        && (body |> Seq.filter ((=) '.') |> Seq.length) <= 1
    if valid then Double.Parse(t, Globalization.CultureInfo.InvariantCulture) else nan

/// XPath's string(number): integers without a point, others in plain decimals.
let private numberToString (x: float) =
    if Double.IsNaN x then "NaN"
    elif Double.IsPositiveInfinity x then "Infinity"
    elif Double.IsNegativeInfinity x then "-Infinity"
    elif x = 0.0 then "0"
    else
        // The shortest of 15 to 17 significant digits that reads back as x.
        let mutable text = ""
        let mutable p = 15
        while text = "" do
            let candidate = NumberText.numberToString ("E" + string (p - 1)) x
            if p = 17 || Double.Parse(candidate, Globalization.CultureInfo.InvariantCulture) = x then text <- candidate
            p <- p + 1
        let neg = text.StartsWith "-"
        let body = text.TrimStart('-')
        let e = body.IndexOf 'E'
        let digits = body.Substring(0, e).Replace(".", "").TrimEnd('0')
        let digits = if digits = "" then "0" else digits
        let exponent = int (body.Substring(e + 1))
        let plain =
            if exponent < 0 then "0." + System.String('0', -exponent - 1) + digits
            elif exponent + 1 >= digits.Length then digits + System.String('0', exponent + 1 - digits.Length)
            else digits.Substring(0, exponent + 1) + "." + digits.Substring(exponent + 1)
        (if neg then "-" else "") + plain

let private toStr (v: XValue) =
    match v with
    | NodeSet [] -> ""
    | NodeSet (x :: _) -> stringValue x
    | XNumber x -> numberToString x
    | XString s -> s
    | XBool b -> if b then "true" else "false"

let private toNum (v: XValue) =
    match v with
    | XNumber x -> x
    | XBool b -> if b then 1.0 else 0.0
    | other -> stringToNumber (toStr other)

let private toBool (v: XValue) =
    match v with
    | NodeSet nodes -> not nodes.IsEmpty
    | XNumber x -> not (x = 0.0 || Double.IsNaN x)
    | XString s -> s.Length > 0
    | XBool b -> b

let private sortDocument (nodes: seq<XNode>) =
    let seen = HashSet<int>()
    nodes |> Seq.filter (fun x -> seen.Add x.Order) |> Seq.sortBy (fun x -> x.Order) |> List.ofSeq

/// The nodes of an axis from x, in the axis's order (reverse axes nearest first).
let private axisNodes (axis: Axis) (x: XNode) : XNode list =
    let rec descendants (y: XNode) = [ for c in y.Children do yield c; yield! descendants c ]
    let rec ancestors (y: XNode) = if isNull y.Parent then [] else y.Parent :: ancestors y.Parent
    let siblings () =
        if isNull x.Parent || x.Kind = NodeKind.Attribute then [], []
        else
            let all = List.ofSeq x.Parent.Children
            let i = all |> List.findIndex (fun c -> Object.ReferenceEquals(c, x))
            List.rev (List.take i all), List.skip (i + 1) all
    match axis with
    | Child -> List.ofSeq x.Children
    | Descendant -> descendants x
    | DescendantOrSelf -> x :: descendants x
    | Parent -> if isNull x.Parent then [] else [ x.Parent ]
    | Ancestor -> ancestors x
    | AncestorOrSelf -> x :: ancestors x
    | Self -> [ x ]
    | AttributeAxis -> if x.Kind = NodeKind.Element then List.ofSeq x.Attributes else []
    | Namespace -> []
    | FollowingSibling -> snd (siblings ())
    | PrecedingSibling -> fst (siblings ())
    | Following ->
        // After x in document order, not its descendants.
        let start = if x.Kind = NodeKind.Attribute then x.Parent else x
        let rec after (y: XNode) =
            if isNull y.Parent then []
            else
                let sibs = List.ofSeq y.Parent.Children
                let i = sibs |> List.findIndex (fun c -> Object.ReferenceEquals(c, y))
                [ for s in List.skip (i + 1) sibs do yield s; yield! descendants s ] @ after y.Parent
        let found = after start
        if x.Kind = NodeKind.Attribute then descendants x.Parent @ found else found
    | Preceding ->
        // Before x in document order, not its ancestors; nearest first.
        let start = if x.Kind = NodeKind.Attribute then x.Parent else x
        let ancestorsOf = HashSet<int>(ancestors start |> List.map (fun a -> a.Order))
        let rec all (y: XNode) = [ yield y; for c in y.Children do yield! all c ]
        let root = List.last (start :: ancestors start)
        all root
        |> List.filter (fun y -> y.Order < start.Order && not (ancestorsOf.Contains y.Order))
        |> List.rev

let private matches (axis: Axis) (test: NodeTest) (x: XNode) =
    let principal = if axis = AttributeAxis then NodeKind.Attribute else NodeKind.Element
    match test with
    | AnyName -> x.Kind = principal
    | PrefixAny prefix -> x.Kind = principal && x.Name.StartsWith(prefix + ":")
    | QName name -> x.Kind = principal && x.Name = name
    | AnyNode -> true
    | TextTest -> x.Kind = NodeKind.Text
    | CommentTest -> x.Kind = NodeKind.Comment
    | InstructionTest target -> x.Kind = NodeKind.Instruction && (target.IsNone || target = Some x.Name)

let private localName (name: string) =
    let i = name.IndexOf ':'
    if i >= 0 then name.Substring(i + 1) else name

/// XPath's round: halves up, -0 for -0.5 <= x < 0.
let private xround (x: float) =
    if Double.IsNaN x || Double.IsInfinity x then x
    elif x < 0.0 && x >= -0.5 then -0.0
    else Math.Floor (x + 0.5)

let rec private eval (e: Expr) (node: XNode) (pos: int) (size: int) : XValue =
    match e with
    | Num x -> XNumber x
    | Lit s -> XString s
    | Or (a, b) -> XBool (toBool (eval a node pos size) || toBool (eval b node pos size))
    | And (a, b) -> XBool (toBool (eval a node pos size) && toBool (eval b node pos size))
    | Negate a -> XNumber (-(toNum (eval a node pos size)))
    | Arith (op, a, b) ->
        let x = toNum (eval a node pos size)
        let y = toNum (eval b node pos size)
        XNumber (match op with
                 | "+" -> x + y
                 | "-" -> x - y
                 | "*" -> x * y
                 | "div" -> x / y
                 | _ -> if y = 0.0 || Double.IsInfinity x then nan else x - y * Math.Truncate (x / y))
    | Compare (op, a, b) -> XBool (compare op (eval a node pos size) (eval b node pos size))
    | Union (a, b) ->
        match eval a node pos size, eval b node pos size with
        | NodeSet xs, NodeSet ys -> NodeSet (sortDocument (xs @ ys))
        | _ -> raise (XPathError "union")
    | Path (fromRoot, steps) ->
        let start =
            if fromRoot then
                let mutable r = node
                while not (isNull r.Parent) do r <- r.Parent
                r
            else node
        NodeSet (applySteps [ start ] steps)
    | Filtered (primary, preds, steps) ->
        match eval primary node pos size with
        | NodeSet nodes ->
            let filtered = preds |> List.fold (fun acc p -> filter acc p) nodes
            NodeSet (if steps.IsEmpty then filtered else applySteps filtered steps)
        | _ when preds.IsEmpty && steps.IsEmpty -> eval primary node pos size
        | _ -> raise (XPathError "filter")
    | Call (name, args) -> call name args node pos size

and private filter (nodes: XNode list) (pred: Expr) : XNode list =
    let size = nodes.Length
    nodes
    |> List.indexed
    |> List.filter (fun (i, x) ->
        match eval pred x (i + 1) size with
        | XNumber n -> n = float (i + 1)
        | v -> toBool v)
    |> List.map snd

and private applySteps (start: XNode list) (steps: Step list) : XNode list =
    steps
    |> List.fold (fun (context: XNode list) step ->
        context
        |> List.collect (fun x ->
            let candidates = axisNodes step.Axis x |> List.filter (matches step.Axis step.Test)
            step.Predicates |> List.fold filter candidates)
        |> sortDocument) start

and private compare (op: string) (a: XValue) (b: XValue) : bool =
    let numeric (x: float) (y: float) =
        match op with
        | "=" -> x = y
        | "!=" -> x <> y
        | "<" -> x < y
        | "<=" -> x <= y
        | ">" -> x > y
        | _ -> x >= y
    let atoms (x: XValue) (y: XValue) =
        if op = "=" || op = "!=" then
            match x, y with
            | XBool _, _ | _, XBool _ -> (if op = "=" then (=) else (<>)) (toBool x) (toBool y)
            | XNumber _, _ | _, XNumber _ -> numeric (toNum x) (toNum y)
            | _ -> (if op = "=" then (=) else (<>)) (toStr x) (toStr y)
        else numeric (toNum x) (toNum y)
    match a, b with
    | NodeSet xs, NodeSet ys ->
        xs |> List.exists (fun x -> ys |> List.exists (fun y -> atoms (XString (stringValue x)) (XString (stringValue y))))
    | NodeSet xs, XBool _ -> atoms (XBool (not xs.IsEmpty)) b
    | XBool _, NodeSet ys -> atoms a (XBool (not ys.IsEmpty))
    | NodeSet xs, other -> xs |> List.exists (fun x -> atoms (match other with XNumber _ -> XNumber (stringToNumber (stringValue x)) | _ -> XString (stringValue x)) other)
    | other, NodeSet ys -> ys |> List.exists (fun y -> atoms other (match other with XNumber _ -> XNumber (stringToNumber (stringValue y)) | _ -> XString (stringValue y)))
    | _ -> atoms a b

and private call (name: string) (args: Expr list) (node: XNode) (pos: int) (size: int) : XValue =
    let arg i = eval (List.item i args) node pos size
    let count = args.Length
    let need (lo: int) (hi: int) = if count < lo || count > hi then raise (XPathError ("arguments of " + name))
    let str i = toStr (arg i)
    let strOrContext () = if count = 0 then stringValue node else str 0
    let nodesOf i = match arg i with NodeSet ns -> ns | _ -> raise (XPathError "node-set expected")
    let firstNode () =
        if count = 0 then Some node
        else match nodesOf 0 with x :: _ -> Some x | [] -> None
    match name with
    | "last" -> need 0 0; XNumber (float size)
    | "position" -> need 0 0; XNumber (float pos)
    | "count" -> need 1 1; XNumber (float (nodesOf 0).Length)
    | "id" -> need 1 1; NodeSet []
    | "local-name" -> need 0 1; XString (match firstNode () with Some x -> localName x.Name | None -> "")
    | "namespace-uri" -> need 0 1; XString ""
    | "name" -> need 0 1; XString (match firstNode () with Some x -> x.Name | None -> "")
    | "string" -> need 0 1; XString (strOrContext ())
    | "concat" -> (if count < 2 then raise (XPathError "concat")); XString (System.String.Join("", [ for i in 0 .. count - 1 -> str i ]))
    | "starts-with" -> need 2 2; XBool ((str 0).StartsWith(str 1, StringComparison.Ordinal))
    | "contains" -> need 2 2; XBool ((str 0).Contains(str 1))
    | "substring-before" ->
        need 2 2
        let s, t = str 0, str 1
        let i = s.IndexOf(t, StringComparison.Ordinal)
        XString (if i < 0 then "" else s.Substring(0, i))
    | "substring-after" ->
        need 2 2
        let s, t = str 0, str 1
        let i = s.IndexOf(t, StringComparison.Ordinal)
        XString (if i < 0 then "" else s.Substring(i + t.Length))
    | "substring" ->
        need 2 3
        let s = str 0
        let start = xround (toNum (arg 1))
        let finish = if count = 3 then start + xround (toNum (arg 2)) else infinity
        let chars = [ for i in 1 .. s.Length do if float i >= start && float i < finish then yield s.[i - 1] ]
        XString (System.String(Array.ofList chars))
    | "string-length" -> need 0 1; XNumber (float (strOrContext ()).Length)
    | "normalize-space" ->
        need 0 1
        let parts = (strOrContext ()).Split([| ' '; '\t'; '\n'; '\r' |], StringSplitOptions.RemoveEmptyEntries)
        XString (System.String.Join(" ", parts))
    | "translate" ->
        need 3 3
        let s, from, into = str 0, str 1, str 2
        let sb = Text.StringBuilder()
        for c in s do
            let i = from.IndexOf c
            if i < 0 then sb.Append c |> ignore
            elif i < into.Length then sb.Append into.[i] |> ignore
        XString (sb.ToString())
    | "boolean" -> need 1 1; XBool (toBool (arg 0))
    | "not" -> need 1 1; XBool (not (toBool (arg 0)))
    | "true" -> need 0 0; XBool true
    | "false" -> need 0 0; XBool false
    | "lang" ->
        need 1 1
        let wanted = (str 0).ToLowerInvariant()
        let mutable x = node
        let mutable found = None
        while found.IsNone && not (isNull x) do
            match x.Attributes |> Seq.tryFind (fun a -> a.Name = "xml:lang") with
            | Some a -> found <- Some (a.Value.ToLowerInvariant())
            | None -> x <- x.Parent
        XBool (match found with Some l -> l = wanted || l.StartsWith(wanted + "-") | None -> false)
    | "number" -> need 0 1; XNumber (if count = 0 then stringToNumber (stringValue node) else toNum (arg 0))
    | "sum" -> need 1 1; XNumber (nodesOf 0 |> List.sumBy (fun x -> stringToNumber (stringValue x)))
    | "floor" -> need 1 1; XNumber (Math.Floor (toNum (arg 0)))
    | "ceiling" -> need 1 1; XNumber (Math.Ceiling (toNum (arg 0)))
    | "round" -> need 1 1; XNumber (xround (toNum (arg 0)))
    | _ -> raise (XPathError ("unknown function " + name))

/// The value of an XPath expression with the document as the context node.
let evaluate (doc: XNode) (expr: Expr) : XValue = eval expr doc 1 1

// ===== FILTERXML ===============================================================

let private trimXml (s: string) = s.Trim([| ' '; '\t'; '\n'; '\r' |])

let private errorTexts = set [ "#DIV/0!"; "#VALUE!"; "#REF!"; "#NAME?"; "#NUM!"; "#N/A"; "#NULL!"; "#SPILL!"; "#CALC!" ]

/// A selected node as a cell value (see the module notes).
let private cellOfNode (x: XNode) : CellValue =
    let text =
        match x.Kind with
        | NodeKind.Element ->
            System.String.Join("", x.Children |> Seq.filter (fun c -> c.Kind = NodeKind.Text) |> Seq.map (fun c -> c.Value)) |> trimXml
        | NodeKind.Attribute -> trimXml x.Value
        | _ -> ""
    if text = "" then CellValue.Error Err.value
    elif text.StartsWith "=" || text.StartsWith "'" || errorTexts.Contains (text.ToUpperInvariant()) then CellValue.String text
    else
        match (Input.parseInputIn DateTime.UtcNow.Year text).Value with
        | CellValue.Number _ | CellValue.Boolean _ as v -> v
        | _ -> CellValue.String text

/// FILTERXML(xml, xpath): the nodes of the document xpath selects.
let filterXml (xml: string) (xpath: string) : Result<CellValue[][], string> =
    try
        let doc = parseXml xml
        match evaluate doc (parseXPath xpath) with
        | NodeSet [] -> Result.Error Err.value
        | NodeSet nodes -> Ok (nodes |> List.map (fun x -> [| cellOfNode x |]) |> Array.ofList)
        | _ -> Result.Error Err.value
    with
    | XmlError _ | XPathError _ -> Result.Error Err.value

/// One text of an argument: the first value of an array or range (Excel 2019
/// takes the xml and xpath of FILTERXML({...}) from its first element).
let private textArg (a: Args) (i: int) : Result<string, string> =
    let v = (a.Array i).[0].[0]
    match v with
    | CellValue.Error e -> Result.Error e
    | v -> coerceText v

let functions : FunctionDef list =
    [ arrayed "FILTERXML" 2 2 "" (fun a -> res {
        let! xml = textArg a 0
        let! xpath = textArg a 1
        // No text of Excel's is longer.
        do! check (xml.Length <= 32767 && xpath.Length <= 32767) Err.value
        return! filterXml xml xpath }) ]

let install () = registerFunctions moduleName functions
