/// Values the readers find in attributes and parts that are text but not XML, read into what they say: VML's
/// style declarations, lengths and colours, field instructions, theme fonts and style names, link targets and
/// picture sources, media types, lists of words, number format codes. (XML is not among them: that is the
/// scanner's, XmlDsl.)
module ValueText

/// A style attribute's value as VML writes it ("position:absolute;margin-left:12pt;rotation:90fd"): its
/// declarations, and the lengths and numbers in them.
module Css =
    /// The value (trimmed) of the first declaration whose property (trimmed) is `name`, in any case.
    let valueIgnoringCase (name: string) (style: string) : string option =
        style.Split(';')
        |> Array.tryPick (fun part ->
            let pieces = part.Split([| ':' |], 2)
            if pieces.Length = 2 && pieces.[0].Trim().Equals(name, System.StringComparison.OrdinalIgnoreCase) then Some (pieces.[1].Trim())
            else None)

    /// The value (trimmed) of the first declaration that is written "name:..." (trimmed; the name as it is).
    let valueOf (name: string) (style: string) : string option =
        style.Split(';')
        |> Array.tryPick (fun part ->
            let text = part.Trim()
            if text.StartsWith(name + ":") then Some (text.Substring(name.Length + 1).Trim()) else None)

    /// The value (trimmed) of the first declaration that is written "name:..." in any case (`name` lower case).
    let valueOfLower (name: string) (style: string) : string option =
        style.Split(';')
        |> Array.tryPick (fun part ->
            let text = part.Trim()
            if text.ToLowerInvariant().StartsWith(name + ":") then Some (text.Substring(name.Length + 1).Trim()) else None)

    /// A number with one of these units after it: (the number's text, the unit); None for a value with none
    /// of them.
    let withUnit (units: string list) (value: string) : (string * string) option =
        units |> List.tryFind (fun unit -> value.EndsWith unit) |> Option.map (fun unit -> value.Substring(0, value.Length - unit.Length), unit)

    /// A number of points ("12pt", "12PT", "12"): the number's text.
    let withoutPoints (value: string) : string =
        if value.EndsWith("pt", System.StringComparison.OrdinalIgnoreCase) then value.Substring(0, value.Length - 2) else value

    /// The items of a list of values ("1pt,2pt,1pt,2pt").
    let items (separator: char) (value: string) : string array = value.Split(separator)

/// Colours as the parts write them.
module Color =
    /// "#rrggbb" of "rrggbb" or "#rrggbb".
    let withHash (value: string) : string = if value.StartsWith("#") then value else "#" + value

    let isHash (value: string) : bool = value.StartsWith("#")

    /// "#RRGGBB" of the last six digits of "rrggbb" or "aarrggbb".
    let ofLastSix (value: string) : string = "#" + value.Substring(value.Length - 6).ToUpperInvariant()

    /// The digits of "rrggbb" or "aarrggbb" without the alpha.
    let withoutAlpha (value: string) : string = if value.Length = 8 then value.Substring(2) else value

    /// The byte of the two hex digits at `at` ("1F497D", 2 -> 0x49).
    let hexByte (hex: string) (at: int) : int = System.Convert.ToInt32(hex.Substring(at, 2), 16)

    /// Word's "aarrggbb" (8 digits) as "#RRGGBB".
    let ofArgb (value: string) : string = "#" + value.Substring(2).ToUpper()

    let private namedColors =
        Map.ofList [ "black", "#000000"; "white", "#FFFFFF"; "silver", "#C0C0C0"; "gray", "#808080"; "grey", "#808080"; "red", "#FF0000"; "maroon", "#800000"
                     "yellow", "#FFFF00"; "olive", "#808000"; "lime", "#00FF00"; "green", "#008000"; "aqua", "#00FFFF"; "teal", "#008080"; "blue", "#0000FF"
                     "navy", "#000080"; "fuchsia", "#FF00FF"; "purple", "#800080"; "orange", "#FFA500" ]

    /// A colour as VML writes it: "#rgb", "#rrggbb", a name, each with or without a palette entry after it ("silver [3209]").
    let ofVml (value: string option) : string option =
        match value |> Option.map (fun text -> (text.Trim().Split(' ')).[0]) with
        | Some "" | None -> None
        | Some color when color.StartsWith("#") && color.Length = 4 ->
            Some (sprintf "#%c%c%c%c%c%c" color.[1] color.[1] color.[2] color.[2] color.[3] color.[3])
        | Some color when color.StartsWith("#") -> Some color
        | Some color ->
            match Map.tryFind (color.ToLowerInvariant()) namedColors with
            | Some hex -> Some hex
            | None -> Some ("#" + color)

/// A field's instruction ("PAGE \* MERGEFORMAT", "REF _Ref123 \h").
module FieldCode =
    /// Whether the instruction (upper case) begins with the keyword.
    let startsWith (keyword: string) (upper: string) : bool = upper.StartsWith(keyword)

    /// PAGE, NUMPAGES, SECTIONPAGES (the instruction upper case): a field whose result the layout computes.
    let isComputed (upper: string) : bool =
        upper.StartsWith "PAGE" || upper.StartsWith "NUMPAGES" || upper.StartsWith "SECTIONPAGES"

    /// What follows the keyword "REF " (trimmed) in an instruction (trimmed).
    let refTarget (instruction: string) : string = instruction.Substring(4).Trim()

/// Names of styles and of theme fonts.
module StyleName =
    /// A table of contents' style ("TOC 1", "toc 2").
    let isTableOfContents (style: string) : bool = style.StartsWith("TOC", System.StringComparison.OrdinalIgnoreCase)

    /// "Heading 2" -> Some "2" (what follows "Heading"); None for a style that is not a heading.
    let headingSuffix (style: string) : string option =
        if style.StartsWith("Heading") then Some (style.Substring("Heading".Length)) else None

    /// A theme font's name: major (headings) or minor (body) - Some true / Some false.
    let isMajorThemeFont (value: string) : bool option =
        if value.StartsWith("major") then Some true
        elif value.StartsWith("minor") then Some false
        else None

/// Link targets and picture sources.
module Link =
    /// A target a reader follows: http(s), mail, telephone, a place in the document (`lower`: lower case).
    let isActive (lower: string) : bool =
        lower.StartsWith "http://" || lower.StartsWith "https://" || lower.StartsWith "mailto:" || lower.StartsWith "tel:"
        || (lower.StartsWith "#" && lower.Length > 1)

    let isDataUrl (source: string) : bool = source.StartsWith "data:"

    /// A data URL of a picture a browser draws (not EMF, WMF or TIFF).
    let isDrawnPicture (source: string) : bool =
        source.StartsWith "data:image/" && not (source.StartsWith "data:image/x-emf") && not (source.StartsWith "data:image/x-wmf") && not (source.StartsWith "data:image/tiff")

/// The media type of a part by the extension of its name.
let mediaTypeOf (path: string) : string =
    let ext = path.ToLower()
    if ext.EndsWith(".png") then "image/png"
    elif ext.EndsWith(".jpg") || ext.EndsWith(".jpeg") then "image/jpeg"
    elif ext.EndsWith(".gif") then "image/gif"
    elif ext.EndsWith(".bmp") then "image/bmp"
    elif ext.EndsWith(".svg") then "image/svg+xml"
    elif ext.EndsWith(".webp") then "image/webp"
    elif ext.EndsWith(".emf") then "image/x-emf"
    elif ext.EndsWith(".wmf") then "image/x-wmf"
    elif ext.EndsWith(".tif") || ext.EndsWith(".tiff") then "image/tiff"
    else "application/octet-stream"

/// The words of a value that is a list (mc:Ignorable="w14 wp14", Requires="wps").
let words (value: string) : string list =
    value.Split([| ' '; '\t'; '\n'; '\r' |], System.StringSplitOptions.RemoveEmptyEntries) |> Array.toList

/// Whether a relationship's type is of that kind ("/header", "/image": what it ends with).
let isRelationshipType (suffix: string) (relType: string) : bool = relType.EndsWith suffix

/// The value without that prefix, where it has it ("#DIV/0!" -> "DIV/0!").
let withoutPrefix (prefix: string) (value: string) : string = if value.StartsWith prefix then value.Substring prefix.Length else value

/// The number of `count` hex digits at `at` (the character of an escape "_x000D_").
let hexNumber (text: string) (at: int) (count: int) : int = System.Convert.ToInt32(text.Substring(at, count), 16)

/// The extension of a path (lower case, after its last dot; "" for none).
let extensionOf (lowerPath: string) : string =
    match lowerPath.LastIndexOf '.' with
    | -1 -> ""
    | index -> lowerPath.Substring(index + 1)

/// Spreadsheet references and the values of a sheet: cell references, ranges, defined names, number formats,
/// pattern fills.
module Sheet =
    /// The corners of a range ("A1:F5"), or the columns of a row's spans ("1:12").
    let corners (range: string) : string array = range.Split(':')

    /// The column letters of a cell reference from `start` to `finish` ("$AB$12").
    let letters (reference: string) (start: int) (finish: int) : string =
        if finish > start then reference.Substring(start, finish - start) else ""

    /// What follows the sheet of a reference ("'Sheet One'!$A$1" -> "$A$1").
    let withoutSheet (reference: string) : string =
        match reference.IndexOf '!' with
        | -1 -> reference
        | i -> reference.Substring(i + 1)

    /// The areas of a reference to several (`A,B`): split at the commas that are not inside a quoted sheet name.
    let areas (text: string) : string list =
        let parts = ResizeArray<string>()
        let mutable start = 0
        let mutable quoted = false
        for i in 0 .. text.Length - 1 do
            if text.[i] = '\'' then quoted <- not quoted
            elif text.[i] = ',' && not quoted then
                parts.Add (text.Substring(start, i - start))
                start <- i + 1
        parts.Add (text.Substring start)
        List.ofSeq parts

    /// The parts of a reference at every comma.
    let commaParts (text: string) : string array = text.Split(',')

    /// A built-in name ("_xlnm.Print_Area").
    let isBuiltInName (name: string) : bool = name.StartsWith "_xlnm."

    /// What a number format's `[...]` at `at` holds (lower case; "" for none) and where it closes (-1: it does not).
    let bracketAt (code: string) (at: int) : string * int =
        let close = code.IndexOf(']', at)
        (if close > at then code.Substring(at + 1, close - at - 1).ToLowerInvariant() else ""), close

    /// A number format with the quotes around a lone currency symbol taken away (`"$"` -> `$`).
    let withoutCurrencyQuotes (isCurrencySymbol: char -> bool) (code: string) : string =
        System.Text.RegularExpressions.Regex.Replace(
            code,
            "\"(.)\"",
            fun m ->
                let inner = m.Groups.[1].Value.[0]
                if isCurrencySymbol inner then string inner else m.Value)

    /// How dense a pattern fill is by its name: dark-* 0.7, light-* 0.3, others 1.
    let patternDensity (patternName: string) : float =
        if patternName.StartsWith "dark" then 0.7
        elif patternName.StartsWith "light" then 0.3
        else 1.0

    /// A sheet part or the shared strings: the large parts of a workbook (`lower`: the path lower case).
    let isLargeTextPart (lower: string) : bool =
        (lower.StartsWith "xl/worksheets/" && not (lower.Contains "/_rels/")) || lower.EndsWith "sharedstrings.xml"
