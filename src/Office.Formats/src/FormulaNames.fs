/// The names of a formula as the file format stores them and as the model holds them: functions newer
/// than Excel 2007 under their prefix (`_xlfn.XLOOKUP`), the names a LET or LAMBDA binds under `_xlpm.`.
module FormulaNames

open System.Text

/// Functions newer than Excel 2007 are stored under a prefix (`_xlfn.XLOOKUP`,
/// `_xlfn._xlws.FILTER`), and the names a LET or LAMBDA binds under `_xlpm.`.
/// A file that names one of them bare is refused by an Excel that knows it.
let private futureFunctions =
    let plain =
        [ "ACOT"; "ACOTH"; "AGGREGATE"; "ARABIC"; "BASE"; "BETA.DIST"; "BETA.INV"; "BINOM.DIST"; "BINOM.DIST.RANGE"; "BINOM.INV"
          "BITAND"; "BITLSHIFT"; "BITOR"; "BITRSHIFT"; "BITXOR"; "CEILING.MATH"; "CEILING.PRECISE"; "CHISQ.DIST"; "CHISQ.DIST.RT"
          "CHISQ.INV"; "CHISQ.INV.RT"; "CHISQ.TEST"; "COMBINA"; "CONCAT"; "CONFIDENCE.NORM"; "CONFIDENCE.T"; "COT"; "COTH"
          "COVARIANCE.P"; "COVARIANCE.S"; "CSC"; "CSCH"; "DAYS"; "DECIMAL"; "ECMA.CEILING"; "ENCODEURL"; "ERF.PRECISE"; "ERFC.PRECISE"
          "EXPON.DIST"; "F.DIST"; "F.DIST.RT"; "F.INV"; "F.INV.RT"; "F.TEST"; "FILTERXML"; "FLOOR.MATH"; "FLOOR.PRECISE"
          "FORECAST.ETS"; "FORECAST.ETS.CONFINT"; "FORECAST.ETS.SEASONALITY"; "FORECAST.ETS.STAT"; "FORECAST.LINEAR"; "FORMULATEXT"
          "GAMMA"; "GAMMA.DIST"; "GAMMA.INV"; "GAMMALN.PRECISE"; "GAUSS"; "HYPGEOM.DIST"; "IFNA"; "IFS"; "IMCOSH"; "IMCOT"; "IMCSC"
          "IMCSCH"; "IMSEC"; "IMSECH"; "IMSINH"; "IMTAN"; "ISFORMULA"; "ISO.CEILING"; "ISOWEEKNUM"; "LOGNORM.DIST"; "LOGNORM.INV"
          "MAXIFS"; "MINIFS"; "MODE.MULT"; "MODE.SNGL"; "MUNIT"; "NEGBINOM.DIST"; "NETWORKDAYS.INTL"; "NORM.DIST"; "NORM.INV"
          "NORM.S.DIST"; "NORM.S.INV"; "NUMBERVALUE"; "PDURATION"; "PERCENTILE.EXC"; "PERCENTILE.INC"; "PERCENTRANK.EXC"
          "PERCENTRANK.INC"; "PERMUTATIONA"; "PHI"; "POISSON.DIST"; "QUARTILE.EXC"; "QUARTILE.INC"; "RANK.AVG"; "RANK.EQ"; "RRI"
          "SEC"; "SECH"; "SHEET"; "SHEETS"; "SKEW.P"; "STDEV.P"; "STDEV.S"; "SWITCH"; "T.DIST"; "T.DIST.2T"; "T.DIST.RT"; "T.INV"
          "T.INV.2T"; "T.TEST"; "TEXTJOIN"; "UNICHAR"; "UNICODE"; "VAR.P"; "VAR.S"; "WEBSERVICE"; "WEIBULL.DIST"; "WORKDAY.INTL"
          "XOR"; "Z.TEST"
          // Dynamic arrays and after.
          "ANCHORARRAY"; "SINGLE"; "SEQUENCE"; "SORTBY"; "UNIQUE"; "RANDARRAY"; "XLOOKUP"; "XMATCH"; "LET"; "LAMBDA"; "MAP"; "REDUCE"
          "SCAN"; "MAKEARRAY"; "BYROW"; "BYCOL"; "ISOMITTED"; "CHOOSECOLS"; "CHOOSEROWS"; "DROP"; "TAKE"; "EXPAND"; "HSTACK"; "VSTACK"
          "TOCOL"; "TOROW"; "WRAPCOLS"; "WRAPROWS"; "TEXTBEFORE"; "TEXTAFTER"; "TEXTSPLIT"; "VALUETOTEXT"; "ARRAYTOTEXT"
          "STOCKHISTORY"; "FIELDVALUE"; "IMAGE"; "GROUPBY"; "PIVOTBY"; "PERCENTOF"; "REGEXTEST"; "REGEXEXTRACT"; "REGEXREPLACE"
          "TRIMRANGE"; "DETECTLANGUAGE"; "TRANSLATE" ]
    let table = System.Collections.Generic.Dictionary<string, string>()
    for name in plain do table.[name] <- "_xlfn."
    for name in [ "FILTER"; "SORT" ] do table.[name] <- "_xlfn._xlws."
    table

let private isNameStart (c: char) = System.Char.IsLetter c || c = '_' || c = '\\'
let private isNameChar (c: char) = System.Char.IsLetterOrDigit c || c = '_' || c = '.' || c = '\\'

/// The trim references of a formula (`A1.:E10`, `A1:.E10`, `A1.:.E10`, `A:.A`) written as the
/// TRIMRANGE calls they stand for (`TRIMRANGE(A1:E10,1,1)`), for the file: what Excel stores for
/// them is not documented, TRIMRANGE is.
let private trimRefsAsCalls (formula: string) : string =
    if formula.IndexOf(".:", System.StringComparison.Ordinal) < 0 && formula.IndexOf(":.", System.StringComparison.Ordinal) < 0 then formula
    else
        let n = formula.Length
        let isRefChar (c: char) = System.Char.IsLetterOrDigit c || c = '$' || c = '_' || c = '.'
        let sb = StringBuilder(n + 24)
        let mutable copied = 0
        let mutable i = 0
        while i < n do
            let c = formula.[i]
            if c = '"' then
                i <- i + 1
                while i < n && formula.[i] <> '"' do i <- i + 1
                i <- i + 1
            elif c = '\'' then
                i <- i + 1
                while i < n && not (formula.[i] = '\'' && not (i + 1 < n && formula.[i + 1] = '\'')) do
                    i <- if formula.[i] = '\'' then i + 2 else i + 1
                i <- i + 1
            elif c = ':' then
                let leading = i > 0 && formula.[i - 1] = '.'
                let trailing = i + 1 < n && formula.[i + 1] = '.'
                if leading || trailing then
                    // the left operand, with its sheet
                    let mutable start = if leading then i - 1 else i
                    while start > copied && isRefChar formula.[start - 1] do start <- start - 1
                    if start > copied && formula.[start - 1] = '!' then
                        start <- start - 1
                        if start > copied && formula.[start - 1] = '\'' then
                            start <- start - 1
                            while start > copied && not (formula.[start - 1] = '\'' && not (start - 2 >= copied && formula.[start - 2] = '\'')) do start <- start - 1
                            if start > copied then start <- start - 1
                        else
                            while start > copied && isRefChar formula.[start - 1] do start <- start - 1
                    let mutable finish = if trailing then i + 2 else i + 1
                    while finish < n && (System.Char.IsLetterOrDigit formula.[finish] || formula.[finish] = '$') do finish <- finish + 1
                    let left = formula.Substring(start, (if leading then i - 1 else i) - start)
                    let right = formula.Substring((if trailing then i + 2 else i + 1), finish - (if trailing then i + 2 else i + 1))
                    if left <> "" && right <> "" then
                        let mode = (if leading then 1 else 0) + (if trailing then 2 else 0)
                        sb.Append(formula.Substring(copied, start - copied)).Append("TRIMRANGE(").Append(left).Append(':').Append(right)
                          .Append(',').Append(mode).Append(',').Append(mode).Append(')') |> ignore
                        copied <- finish
                    i <- finish
                else i <- i + 1
            else i <- i + 1
        sb.Append(formula.Substring copied).ToString()

/// A formula as the file format stores it: prefixes added where they are missing.
let fileFormula (formula: string) : string =
    let formula = trimRefsAsCalls formula
    // Fast path: no `(` after a name of the table, nothing to do.
    let n = formula.Length
    let names = System.Collections.Generic.List<int * int * bool>()   // start, length, followed by '('
    let mutable i = 0
    while i < n do
        let c = formula.[i]
        if c = '"' then
            i <- i + 1
            while i < n && not (formula.[i] = '"' && not (i + 1 < n && formula.[i + 1] = '"')) do
                i <- if formula.[i] = '"' then i + 2 else i + 1
            i <- i + 1
        elif c = '\'' then
            i <- i + 1
            while i < n && not (formula.[i] = '\'' && not (i + 1 < n && formula.[i + 1] = '\'')) do
                i <- if formula.[i] = '\'' then i + 2 else i + 1
            i <- i + 1
        elif c = '[' then
            // A structured reference's column names are not formula names.
            let mutable depth = 0
            let mutable stop = false
            while i < n && not stop do
                if formula.[i] = '[' then depth <- depth + 1
                elif formula.[i] = ']' then
                    depth <- depth - 1
                    if depth = 0 then stop <- true
                i <- i + 1
        elif isNameStart c && (i = 0 || not (isNameChar formula.[i - 1]) && formula.[i - 1] <> '!' && formula.[i - 1] <> '$') then
            let start = i
            while i < n && isNameChar formula.[i] do i <- i + 1
            names.Add((start, i - start, i < n && formula.[i] = '('))
        else i <- i + 1
    let upper (start, length, _) = formula.Substring(start, length).ToUpperInvariant()
    let needs = names |> Seq.exists (fun ((_, _, call) as name) -> call && futureFunctions.ContainsKey (upper name))
    if not needs then formula
    else
        // The names a LET or LAMBDA binds: LAMBDA's arguments but the last, LET's odd ones.
        let parameters = System.Collections.Generic.HashSet<string>()
        for ((start, length, call) as name) in names do
            let fn = upper name
            if call && (fn = "LET" || fn = "LAMBDA") then
                let mutable j = start + length + 1
                let mutable depth = 0
                let mutable argument = 0
                let mutable argStart = j
                let args = System.Collections.Generic.List<string>()
                let mutable stop = false
                while j < n && not stop do
                    match formula.[j] with
                    | '"' ->
                        j <- j + 1
                        while j < n && formula.[j] <> '"' do j <- j + 1
                    | '(' | '{' -> depth <- depth + 1
                    | ')' | '}' when depth > 0 -> depth <- depth - 1
                    | ')' ->
                        args.Add(formula.Substring(argStart, j - argStart).Trim())
                        stop <- true
                    | ',' when depth = 0 ->
                        args.Add(formula.Substring(argStart, j - argStart).Trim())
                        argument <- argument + 1
                        argStart <- j + 1
                    | _ -> ()
                    j <- j + 1
                let bound =
                    if fn = "LAMBDA" then args |> Seq.take (max 0 (args.Count - 1))
                    else args |> Seq.indexed |> Seq.filter (fun (k, _) -> k % 2 = 0 && k < args.Count - 1) |> Seq.map snd
                for name in bound do
                    let name = if name.StartsWith("_xlpm.", System.StringComparison.OrdinalIgnoreCase) then name.Substring 6 else name
                    if name <> "" then parameters.Add(name.ToUpperInvariant()) |> ignore
        let sb = StringBuilder(n + 16)
        let mutable at = 0
        for ((start, length, call) as name) in names do
            let word = upper name
            let prefix =
                if word.StartsWith "_XL" then ""
                elif call && futureFunctions.ContainsKey word then futureFunctions.[word]
                elif parameters.Contains word then "_xlpm."
                else ""
            if prefix <> "" then
                sb.Append(formula.Substring(at, start - at)).Append(prefix) |> ignore
                at <- start
        sb.Append(formula.Substring at).ToString()

/// A formula as the model holds it and Excel shows it: the file format's
/// prefixes taken off where `fileFormula` puts them back (`_xlfn.XLOOKUP` is
/// `XLOOKUP`, `_xlpm.x` is `x`). A prefixed name the table does not know keeps
/// its prefix, so it is written as it was read.
let modelFormula (formula: string) : string =
    if formula.IndexOf("_xl", System.StringComparison.OrdinalIgnoreCase) < 0 then formula
    else
        let n = formula.Length
        let sb = StringBuilder(n)
        let mutable i = 0
        while i < n do
            let c = formula.[i]
            if c = '"' || c = '\'' then
                let quote = c
                let start = i
                i <- i + 1
                while i < n && not (formula.[i] = quote && not (i + 1 < n && formula.[i + 1] = quote)) do
                    i <- if formula.[i] = quote then i + 2 else i + 1
                i <- min n (i + 1)
                sb.Append(formula.Substring(start, i - start)) |> ignore
            elif isNameStart c && (i = 0 || not (isNameChar formula.[i - 1])) then
                let start = i
                while i < n && isNameChar formula.[i] do i <- i + 1
                let word = formula.Substring(start, i - start)
                let upper = word.ToUpperInvariant()
                let call = i < n && formula.[i] = '('
                let strip (prefix: string) =
                    if upper.StartsWith (prefix.ToUpperInvariant()) then
                        let bare = upper.Substring prefix.Length
                        if prefix = "_xlpm." then Some (word.Substring prefix.Length)
                        elif call && futureFunctions.ContainsKey bare && futureFunctions.[bare] = prefix then Some (word.Substring prefix.Length)
                        else None
                    else None
                match [ "_xlfn._xlws."; "_xlfn."; "_xlpm." ] |> List.tryPick strip with
                | Some bare -> sb.Append bare |> ignore
                | None -> sb.Append word |> ignore
            else
                sb.Append c |> ignore
                i <- i + 1
        sb.ToString()

/// An error value of a cell with its "#" ("N/A" -> "#N/A").
let errorValue (text: string) : string = if text.StartsWith "#" then text else "#" + text
