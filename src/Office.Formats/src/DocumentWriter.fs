module DocumentWriter

open Fable.Core
open Fable.Core.JsInterop
open XmlGenerator
open ZipWriter
open ZipInterop
open DocumentProcessors.EditorProcessor

[<Fable.Core.Emit("$0 === null || $0 === undefined")>]
let private isNullOrUndefined (_value: obj) : bool = jsNative

let private stringOrDefault defaultValue (value: obj) =
    if isNullOrUndefined value then defaultValue
    else
        match unbox<string> value with
        | null | "" -> defaultValue
        | value -> value

let private optionalString (value: obj) =
    if isNullOrUndefined value then None
    else
        match unbox<string> value with
        | null | "" -> None
        | value -> Some value

let private optionalArray (value: obj) : obj array =
    if isNullOrUndefined value then [||]
    else unbox<obj array> value

let private intOrDefault defaultValue (value: obj) =
    if isNullOrUndefined value then defaultValue
    else unbox<int> value

let private optionalInt (value: obj) =
    if isNullOrUndefined value then None
    else Some (unbox<int> value)

let private boolOrDefault defaultValue (value: obj) =
    if isNullOrUndefined value then defaultValue
    else unbox<bool> value

let private optionalBool (value: obj) =
    if isNullOrUndefined value then None
    else Some (unbox<bool> value)

let private optionalIntString (value: obj) =
    optionalString value
    |> Option.bind (fun text ->
        match System.Int32.TryParse(text) with
        | true, parsed -> Some (string parsed)
        | false, _ -> None)

[<Fable.Core.Emit("Number($0)")>]
let private jsNumber (_value: obj) : float = jsNative

/// Numbers arrive as JS numbers from the editor, but some reader fields
/// (table widths) are passed through as the raw attribute string.
let private optionalFloat (value: obj) =
    if isNullOrUndefined value then None
    else
        let number = jsNumber value
        if System.Double.IsNaN number || System.Double.IsInfinity number then None else Some number

let private floatOrDefault defaultValue (value: obj) =
    optionalFloat value |> Option.defaultValue defaultValue

/// Round half away from zero. `int` truncates (11pt -> 14.667px -> 29
/// half-points -> 14.5pt), and Math.Round's MidpointRounding overloads are not
/// dependable under Fable, so every unit conversion below goes through this.
let private roundToInt (value: float) : int =
    if System.Double.IsNaN value || System.Double.IsInfinity value then 0
    elif value >= 0.0 then int (System.Math.Floor (value + 0.5))
    else -(int (System.Math.Floor (0.5 - value)))

let private pixelsToTwips (value: float) =
    roundToInt (value * 15.0) |> string

let private pixelsToPoints (value: float) =
    value * 72.0 / 96.0

type PageSetupInfo = {
    PageWidth: float
    PageHeight: float
    MarginLeft: float
    MarginRight: float
    MarginTop: float
    MarginBottom: float
    HeaderDistance: float
    FooterDistance: float
    DifferentFirstPage: bool
    EvenAndOddHeaders: bool
    /// 0 or 1 = single column. Space and widths are px (reader: twips / 15).
    ColumnCount: int
    ColumnSpace: float
    ColumnSeparator: bool
    ColumnWidths: (float * float) list
    SectionStart: string option
    /// The final section's header / footer references ((kind, relationship id))
    /// and page number start: what the body-level w:sectPr holds.
    HeaderReferences: (string * string) list
    FooterReferences: (string * string) list
    PageNumberStart: int option
    PageNumberFormat: string option
}

let private defaultPageSetupInfo = {
    PageWidth = 816.0
    PageHeight = 1056.0
    MarginLeft = 96.0
    MarginRight = 96.0
    MarginTop = 96.0
    MarginBottom = 96.0
    HeaderDistance = 48.0
    FooterDistance = 48.0
    DifferentFirstPage = false
    EvenAndOddHeaders = false
    ColumnCount = 1
    ColumnSpace = 0.0
    ColumnSeparator = false
    ColumnWidths = []
    SectionStart = None
    HeaderReferences = []
    FooterReferences = []
    PageNumberStart = None
    PageNumberFormat = None
}

let private referencePairs (value: obj) : (string * string) list =
    optionalArray value
    |> Array.toList
    |> List.choose (fun item ->
        optionalString (item?RelationshipId)
        |> Option.map (fun relId -> item?Kind |> stringOrDefault (item?Type |> stringOrDefault "default"), relId))

let private optionalPageSetup (value: obj) =
    if isNullOrUndefined value then None
    else
        Some {
            PageWidth = floatOrDefault defaultPageSetupInfo.PageWidth (value?PageWidth)
            PageHeight = floatOrDefault defaultPageSetupInfo.PageHeight (value?PageHeight)
            MarginLeft = floatOrDefault defaultPageSetupInfo.MarginLeft (value?MarginLeft)
            MarginRight = floatOrDefault defaultPageSetupInfo.MarginRight (value?MarginRight)
            MarginTop = floatOrDefault defaultPageSetupInfo.MarginTop (value?MarginTop)
            MarginBottom = floatOrDefault defaultPageSetupInfo.MarginBottom (value?MarginBottom)
            HeaderDistance = floatOrDefault defaultPageSetupInfo.HeaderDistance (value?HeaderDistance)
            FooterDistance = floatOrDefault defaultPageSetupInfo.FooterDistance (value?FooterDistance)
            DifferentFirstPage = boolOrDefault defaultPageSetupInfo.DifferentFirstPage (value?DifferentFirstPage)
            EvenAndOddHeaders = boolOrDefault defaultPageSetupInfo.EvenAndOddHeaders (value?EvenAndOddHeaders)
            ColumnCount = optionalFloat (value?ColumnCount) |> Option.map roundToInt |> Option.defaultValue 1
            ColumnSpace = floatOrDefault 0.0 (value?ColumnSpace)
            ColumnSeparator = boolOrDefault false (value?ColumnSeparator)
            ColumnWidths =
                optionalArray (value?ColumnWidths)
                |> Array.toList
                |> List.map (fun column -> floatOrDefault 0.0 (column?W), floatOrDefault 0.0 (column?S))
            SectionStart = optionalString (value?SectionStart)
            HeaderReferences = referencePairs (value?HeaderReferences)
            FooterReferences = referencePairs (value?FooterReferences)
            PageNumberStart = optionalInt (value?PageNumberStart)
            PageNumberFormat = optionalString (value?PageNumberFormat)
        }

let private sectionColumnsElement (setup: PageSetupInfo) =
    if setup.ColumnCount < 2 then None
    else
        let attrs =
            [ yield ("w:num", string setup.ColumnCount)
              if setup.ColumnSeparator then yield ("w:sep", "1")
              yield ("w:space", pixelsToTwips setup.ColumnSpace)
              if not setup.ColumnWidths.IsEmpty then yield ("w:equalWidth", "0") ]
        let columns =
            setup.ColumnWidths
            |> List.map (fun (width, space) -> selfClosing "w:col" [("w:w", pixelsToTwips width); ("w:space", pixelsToTwips space)])
        Some (element "w:cols" attrs columns)

/// Children of w:sectPr in CT_SectPr sequence order: header/footer references,
/// type, pgSz, pgMar, pgNumType, cols, titlePg.
let private sectionPropertiesChildren (references: string list) (sectionType: string option) (setup: PageSetupInfo option) (pageNumberStart: int option) (pageNumberFormat: string option) =
    [ yield! references
      match sectionType with
      | Some value when value = "continuous" || value = "nextPage" || value = "evenPage" || value = "oddPage" || value = "nextColumn" ->
          yield selfClosing "w:type" [("w:val", value)]
      | _ -> ()
      match setup with
      | Some setup ->
          yield selfClosing "w:pgSz" [("w:w", pixelsToTwips setup.PageWidth); ("w:h", pixelsToTwips setup.PageHeight)]
          yield selfClosing "w:pgMar" [
              ("w:top", pixelsToTwips setup.MarginTop)
              ("w:right", pixelsToTwips setup.MarginRight)
              ("w:bottom", pixelsToTwips setup.MarginBottom)
              ("w:left", pixelsToTwips setup.MarginLeft)
              ("w:header", pixelsToTwips setup.HeaderDistance)
              ("w:footer", pixelsToTwips setup.FooterDistance)
              ("w:gutter", "0")
          ]
      | None -> ()
      match pageNumberStart, pageNumberFormat with
      | None, None -> ()
      | start, format ->
          yield selfClosing "w:pgNumType" [
              match format with
              | Some value when value <> "" -> yield ("w:fmt", value)
              | _ -> ()
              match start with
              | Some start -> yield ("w:start", string (max 0 start))
              | None -> () ]
      match setup |> Option.bind sectionColumnsElement with
      | Some columns -> yield columns
      | None -> ()
      match setup with
      | Some setup when setup.DifferentFirstPage -> yield selfClosing "w:titlePg" []
      | _ -> () ]

/// Type conversion helpers (inverse of TypeConversion in DocumentProcessors)
///
/// The editor model is in CSS pixels at 96 dpi except where noted, because that
/// is what DocxReader.toEditorParagraphs produces. Each function here is the
/// exact inverse of the reader conversion named in its comment; keep the two
/// in step.
module InverseConversion =
    /// Font size: px -> half-points. Reader: wordFontSizeToPoints = halfPoints / 2 * 96 / 72.
    let pixelsToHalfPoints (pixels: float) : string =
        string (max 1 (roundToInt (pixels * 1.5)))

    /// w:position: points -> half-points. Reader: wordPositionToPoints = halfPoints / 2.
    let pointsToHalfPoints (points: float) : string =
        string (roundToInt (points * 2.0))

    /// Indents, tab positions, character spacing, page geometry: px -> twips. Reader: twips / 15.
    let pixelsToTwips (pixels: float) : string =
        string (roundToInt (pixels * 15.0))

    /// Border width: px -> eighths of a point. Reader: borderWidthToPoints = sz / 6.
    let pixelsToEighthPoints (pixels: float) : string =
        string (max 0 (roundToInt (pixels * 6.0)))

    /// Convert CSS color (#RRGGBB) to Word format (RRGGBB)
    let cssColorToWord (color: string option) : string option =
        color
        |> Option.map (fun c ->
            if c.StartsWith("#") then
                c.Substring(1).ToUpper()
            else
                c.ToUpper()
        )

    /// Convert web alignment to Word alignment
    let alignmentToWord (alignment: string) : string =
        match alignment.ToLower() with
        | "center" -> "center"
        | "right" | "end" -> "right"
        // Style alignments reach the writer in the reader's (OOXML) spelling.
        | "justify" | "both" -> "both"
        | "distribute" -> "distribute"
        | _ -> "left"

    /// Paragraph spacing before/after: points -> twips. Reader: wordSpacingTwipsToPoints = twips / 20.
    let pointsToTwips (points: float) : string =
        string (roundToInt (points * 20.0))

    /// Line spacing as a multiple of single spacing -> 240ths of a line. Reader: line / 240.
    let lineSpacingToWord (spacing: float) : string =
        string (roundToInt (spacing * 240.0))

    /// Exact / at-least line spacing. The reader stores it as a multiple of the
    /// paragraph's font size in px: (twips / 15) / fontSize, so the inverse needs
    /// the same font size (ParagraphMarkFontSize).
    let fixedLineSpacingToTwips (fontSizePixels: float) (spacing: float) : string =
        string (roundToInt (spacing * (max 1.0 fontSizePixels) * 15.0))

/// Run properties the typed EditorRun record does not carry; read off the
/// untyped editor run object.
type private RunExtras = {
    StyleId: string option
    AllCaps: bool
    SmallCaps: bool
    DoubleStrike: bool
    Hidden: bool
    UnderlineStyle: string option
    UnderlineColor: string option
    /// Baseline shift in points (reader: w:position half-points / 2).
    Position: float
    EmphasisMark: string option
    /// The run object says which on/off properties and size were set directly
    /// on the run (ExplicitX: null = inherited, true/false = set). Without it
    /// (callers that only have effective formatting) everything is written.
    HasExplicit: bool
    ExplicitBold: bool option
    ExplicitItalic: bool option
    ExplicitUnderline: bool option
    ExplicitAllCaps: bool option
    ExplicitSmallCaps: bool option
    ExplicitStrike: bool option
    ExplicitDoubleStrike: bool option
    ExplicitHidden: bool option
    ExplicitFontSize: float option
    /// w:color w:val="auto": the run has no colour and takes none from a style.
    ExplicitColorAuto: bool
    /// The run's font is the one its styles give it: nothing to write.
    FontInherited: bool
}

let private noRunExtras =
    { StyleId = None; AllCaps = false; SmallCaps = false; DoubleStrike = false; Hidden = false
      UnderlineStyle = None; UnderlineColor = None; Position = 0.0; EmphasisMark = None
      HasExplicit = false; ExplicitBold = None; ExplicitItalic = None; ExplicitUnderline = None
      ExplicitAllCaps = None; ExplicitSmallCaps = None; ExplicitStrike = None; ExplicitDoubleStrike = None
      ExplicitHidden = None; ExplicitFontSize = None; ExplicitColorAuto = false; FontInherited = false }

/// An on/off property written as direct formatting. "Off" has to be an
/// explicit w:val="0": leaving the element out hands the property back to the
/// paragraph style and e.g. makes a non-bold run of a bold heading style bold.
let private toggle name (isOn: bool) =
    if isOn then selfClosing name [] else selfClosing name [("w:val", "0")]

/// A run object that says what was set directly on the run (HasExplicit) gets
/// exactly that written: on, w:val="0" for explicit off, nothing for a property
/// it inherits. A run object with effective formatting only (style and
/// document defaults already merged in, as the reader used to deliver alone)
/// gets every property spelled out, since "not bold" may be an override there.
let private runToggle (extras: RunExtras) name (explicitValue: bool option) (effective: bool) =
    if extras.HasExplicit then
        match explicitValue with
        | Some isOn -> [toggle name isOn]
        | None -> []
    else [toggle name effective]

/// Run properties in CT_RPr (EG_RPrBase) sequence order. Word rejects or
/// silently repairs w:rPr whose children are out of order, so the order of the
/// yields below is significant: rStyle, rFonts, b, i, caps, smallCaps, strike,
/// dstrike, vanish, color, spacing, position, sz, szCs, highlight, u, vertAlign.
let private orderedRunProps (extras: RunExtras) (run: EditorRun) : string list =
    [
        match extras.StyleId with
        | Some styleId when styleId <> "" -> yield selfClosing "w:rStyle" [("w:val", styleId)]
        | _ -> ()
        // (A run that states the font its styles give it anyway would keep that
        // font when the style's is changed; every run of a saved document did.)
        if not (extras.HasExplicit && extras.FontInherited) then
            yield selfClosing "w:rFonts" [("w:ascii", run.FontFamily); ("w:hAnsi", run.FontFamily); ("w:cs", run.FontFamily)]
        yield! runToggle extras "w:b" extras.ExplicitBold run.Bold
        yield! runToggle extras "w:i" extras.ExplicitItalic run.Italic
        yield! runToggle extras "w:caps" extras.ExplicitAllCaps extras.AllCaps
        yield! runToggle extras "w:smallCaps" extras.ExplicitSmallCaps extras.SmallCaps
        yield! runToggle extras "w:strike" extras.ExplicitStrike run.Strike
        yield! runToggle extras "w:dstrike" extras.ExplicitDoubleStrike extras.DoubleStrike
        yield! runToggle extras "w:vanish" extras.ExplicitHidden extras.Hidden
        match InverseConversion.cssColorToWord run.Color with
        | Some color -> yield selfClosing "w:color" [("w:val", color)]
        | None -> if extras.ExplicitColorAuto then yield selfClosing "w:color" [("w:val", "auto")]
        if abs run.CharacterSpacing >= 0.001 then
            yield selfClosing "w:spacing" [("w:val", InverseConversion.pixelsToTwips run.CharacterSpacing)]
        if abs extras.Position >= 0.001 then
            yield selfClosing "w:position" [("w:val", InverseConversion.pointsToHalfPoints extras.Position)]
        match (if extras.HasExplicit then extras.ExplicitFontSize else Some run.FontSize) with
        | Some size ->
            let sz = InverseConversion.pixelsToHalfPoints size
            yield selfClosing "w:sz" [("w:val", sz)]
            yield selfClosing "w:szCs" [("w:val", sz)]
        | None -> ()
        match run.Highlight with
        | Some color -> yield selfClosing "w:highlight" [("w:val", color)]
        | None -> ()
        let underline = if extras.HasExplicit then extras.ExplicitUnderline else Some run.Underline
        if underline = Some true then
            let style =
                match extras.UnderlineStyle with
                | Some value when value <> "" && value <> "none" -> value
                | _ -> "single"
            let color =
                match InverseConversion.cssColorToWord extras.UnderlineColor with
                | Some value when value <> "AUTO" -> [("w:color", value)]
                | _ -> []
            yield selfClosing "w:u" (("w:val", style) :: color)
        elif underline = Some false then
            yield selfClosing "w:u" [("w:val", "none")]
        match run.VertAlign with
        | Some value when value = "superscript" || value = "subscript" -> yield selfClosing "w:vertAlign" [("w:val", value)]
        | _ -> ()
        match extras.EmphasisMark with
        | Some value when value <> "" && value <> "none" -> yield selfClosing "w:em" [("w:val", value)]
        | _ -> ()
    ]

/// Generate run properties XML
let generateRunProps (run: EditorRun) : string list =
    orderedRunProps noRunExtras run

/// The content of a run. Inside w:del the schema only allows w:delText;
/// a w:t there makes Word report the file as corrupt.
///
/// The reader turns <w:tab/> into "\t" and <w:br/> / <w:cr/> into "\n" inside
/// the run text. They have to go back out as elements: a literal tab or line
/// feed inside w:t is plain whitespace to Word, so tabbed columns collapsed and
/// manual line breaks disappeared after one save.
let private runTextElements (deleted: bool) (text: string) : string list =
    let textName = if deleted then "w:delText" else "w:t"
    let textElement (value: string) = elementWithText textName [("xml:space", "preserve")] value
    if isNull text || text = "" then [ textElement "" ]
    elif text.IndexOf('\t') < 0 && text.IndexOf('\n') < 0 && text.IndexOf('\r') < 0 then [ textElement text ]
    else
        let parts = ResizeArray<string>()
        let mutable segmentStart = 0
        let flush upTo =
            if upTo > segmentStart then parts.Add (textElement (text.Substring(segmentStart, upTo - segmentStart)))
        let mutable i = 0
        while i < text.Length do
            match text.[i] with
            | '\t' ->
                flush i
                parts.Add (selfClosing "w:tab" [])
                segmentStart <- i + 1
            | '\n' ->
                flush i
                parts.Add (selfClosing "w:br" [])
                segmentStart <- i + 1
            | '\r' ->
                flush i
                parts.Add (selfClosing "w:br" [])
                if i + 1 < text.Length && text.[i + 1] = '\n' then i <- i + 1
                segmentStart <- i + 1
            | _ -> ()
            i <- i + 1
        flush text.Length
        List.ofSeq parts

let private runElement (rPr: string list) (deleted: bool) (text: string) : string =
    let rPrElement = if List.isEmpty rPr then [] else [element "w:rPr" [] rPr]
    element "w:r" [] (rPrElement @ runTextElements deleted text)

/// Generate a run element with formatting
let generateRun (run: EditorRun) : string =
    runElement (generateRunProps run) false run.Text

let private runExtrasFromObject (runObj: obj) = {
    StyleId = optionalString (runObj?Style) |> Option.orElse (optionalString (runObj?RunStyle)) |> Option.orElse (optionalString (runObj?CharacterStyle))
    AllCaps = boolOrDefault false (runObj?AllCaps)
    SmallCaps = boolOrDefault false (runObj?SmallCaps)
    DoubleStrike = boolOrDefault false (runObj?DoubleStrike)
    Hidden = boolOrDefault false (runObj?Hidden)
    UnderlineStyle = optionalString (runObj?UnderlineStyle)
    UnderlineColor = optionalString (runObj?UnderlineColor)
    Position = floatOrDefault 0.0 (runObj?Position)
    EmphasisMark = optionalString (runObj?EmphasisMark)
    HasExplicit = boolOrDefault false (runObj?HasExplicit)
    ExplicitBold = optionalBool (runObj?ExplicitBold)
    ExplicitItalic = optionalBool (runObj?ExplicitItalic)
    ExplicitUnderline = optionalBool (runObj?ExplicitUnderline)
    ExplicitAllCaps = optionalBool (runObj?ExplicitAllCaps)
    ExplicitSmallCaps = optionalBool (runObj?ExplicitSmallCaps)
    ExplicitStrike = optionalBool (runObj?ExplicitStrike)
    ExplicitDoubleStrike = optionalBool (runObj?ExplicitDoubleStrike)
    ExplicitHidden = optionalBool (runObj?ExplicitHidden)
    ExplicitFontSize = optionalFloat (runObj?ExplicitFontSize)
    ExplicitColorAuto = boolOrDefault false (runObj?ExplicitColorAuto)
    FontInherited = boolOrDefault false (runObj?FontFamilyInherited)
}

let private isDeletedRun (runObj: obj) =
    optionalString (runObj?RevisionType) = Some "delete"

let private generateRunWithObjectProps (runObj: obj) (run: EditorRun) : string =
    let rPr = orderedRunProps (runExtrasFromObject runObj) run
    let deleted = isDeletedRun runObj
    // Where the page ended when the document was last laid out by its author's
    // application: the layout keeps those page ends, so a save that dropped the
    // hint repaginated the document (75 pages came back as 53).
    if boolOrDefault false (runObj?RenderedPageBreakBefore) && not deleted then
        let rPrElement = if List.isEmpty rPr then [] else [element "w:rPr" [] rPr]
        element "w:r" [] (rPrElement @ [selfClosing "w:lastRenderedPageBreak" []] @ runTextElements deleted run.Text)
    else
        runElement rPr deleted run.Text

/// The further runs of the result of the field whose first run is being written: a field's
/// result that the editor holds as several runs (a part of it was formatted) is one field
/// with those runs, each with its own text. (Set by the paragraph writer.)
let mutable private fieldResultParts : (obj * EditorRun) list = []

let private generateFieldRun (runObj: obj) (fallbackRun: EditorRun) (fieldCode: string) =
    match fieldResultParts with
    | [] ->
        // (A field without a result is shown as "[Field:<code>]": that is not its result.)
        let placeholder = fallbackRun.Text.StartsWith "[Field:" && fallbackRun.Text.EndsWith "]"
        let result = optionalString (runObj?FieldResult) |> Option.defaultValue (if placeholder then "" else fallbackRun.Text)
        let resultRun = generateRunWithObjectProps runObj { fallbackRun with Text = result }
        element "w:fldSimple" [("w:instr", fieldCode)] [resultRun]
    | parts ->
        fieldResultParts <- []
        element "w:fldSimple" [("w:instr", fieldCode)] (generateRunWithObjectProps runObj fallbackRun :: (parts |> List.map (fun (part, run) -> generateRunWithObjectProps part run)))

let private bookmarkStartXml (runObj: obj) =
    optionalString (runObj?BookmarkStartId)
    |> Option.orElse (optionalString (runObj?BookmarkId))
    |> Option.map (fun id ->
        let name = optionalString (runObj?BookmarkName) |> Option.defaultValue ("bookmark" + id)
        selfClosing "w:bookmarkStart" [("w:id", id); ("w:name", name)])

let private bookmarkEndXml (runObj: obj) =
    let explicitEnd =
        optionalString (runObj?BookmarkEndId)
        |> Option.orElse (optionalString (runObj?BookmarkId))
    let fieldWrappedEnd =
        optionalString (runObj?FieldCode)
        |> Option.bind (fun _ -> optionalString (runObj?BookmarkStartId))
    explicitEnd
    |> Option.orElse fieldWrappedEnd
    |> Option.map (fun id -> selfClosing "w:bookmarkEnd" [("w:id", id)])

let private generateReferenceRun (runObj: obj) (fallbackRun: EditorRun) =
    let inner =
        if boolOrDefault false (runObj?IsEquation) then
            optionalString (runObj?EquationXml) |> Option.defaultValue ""
        else
            match optionalString (runObj?FieldCode) with
            | Some fieldCode when fieldCode <> "" -> generateFieldRun runObj fallbackRun fieldCode
            | _ ->
                match optionalString (runObj?CommentId), optionalString (runObj?NoteId) with
                | Some commentId, _ ->
                    // The mark with the formatting of its run - but for the colour the editor shows a
                    // mark without one in, which is not the file's.
                    let ownColor = fallbackRun.Color |> Option.filter (fun color -> color.ToLowerInvariant() <> DocxReader.commentMarkDisplayColor)
                    let rPr = orderedRunProps (runExtrasFromObject runObj) { fallbackRun with Color = ownColor }
                    element "w:r" [] ((if List.isEmpty rPr then [] else [element "w:rPr" [] rPr]) @ [selfClosing "w:commentReference" [("w:id", commentId)]])
                | None, Some noteId ->
                    let noteType = optionalString (runObj?NoteType) |> Option.defaultValue "footnote"
                    let elementName = if noteType = "endnote" then "w:endnoteReference" else "w:footnoteReference"
                    // The mark keeps the formatting of its run (the reference style, a colour,
                    // a size): written bare it took the paragraph's.
                    let rPr = orderedRunProps (runExtrasFromObject runObj) fallbackRun
                    element "w:r" [] ((if List.isEmpty rPr then [] else [element "w:rPr" [] rPr]) @ [selfClosing elementName [("w:id", noteId)]])
                | None, None ->
                    if fallbackRun.Text = "" && bookmarkStartXml runObj |> Option.isSome || fallbackRun.Text = "" && bookmarkEndXml runObj |> Option.isSome then ""
                    else generateRunWithObjectProps runObj fallbackRun
    [ bookmarkStartXml runObj; (if inner <> "" then Some inner else None); bookmarkEndXml runObj ]
    |> List.choose id
    |> String.concat ""

let private revisionAttributes (id: string option) (author: string option) (date: string option) =
    [ yield ("w:id", id |> Option.defaultValue "1")
      // w:author is required on a tracked change.
      yield ("w:author", author |> Option.defaultValue "")
      match date with
      | Some value -> yield ("w:date", value)
      | None -> () ]

/// How many runs of each tracked move (kind + name) of the part being written
/// are still to come, and which moves have been opened: the range markers go
/// before the first run and after the last run of a move, wherever those are.
let private moveRunsLeft = System.Collections.Generic.Dictionary<string, int>()
let private moveRangeIds = System.Collections.Generic.Dictionary<string, string>()
let mutable private nextMoveRangeId = 900000

let private moveKey (runObj: obj) =
    match optionalString (runObj?RevisionType), optionalString (runObj?RevisionMoveName) with
    | Some kind, Some name when kind = "moveFrom" || kind = "moveTo" -> Some (kind + "\u0000" + name, kind, name)
    | _ -> None

/// The tracked change around a run: w:ins, w:del, a deletion inside an
/// insertion (w:del in w:ins: another author's insertion that was deleted),
/// or a move with its range markers.
let private revisionWrappedRun (runObj: obj) (runXml: string) =
    let attrs = revisionAttributes (optionalString (runObj?RevisionId)) (optionalString (runObj?RevisionAuthor)) (optionalString (runObj?RevisionDate))
    match optionalString (runObj?RevisionType) with
    | Some "insert" -> element "w:ins" attrs [runXml]
    | Some "delete" ->
        let deletion = element "w:del" attrs [runXml]
        let insertedId, insertedAuthor, insertedDate =
            optionalString (runObj?InsertedRevisionId), optionalString (runObj?InsertedRevisionAuthor), optionalString (runObj?InsertedRevisionDate)
        if insertedId.IsSome || insertedAuthor.IsSome || insertedDate.IsSome then
            element "w:ins" (revisionAttributes insertedId insertedAuthor insertedDate) [deletion]
        else deletion
    | Some kind when kind = "moveFrom" || kind = "moveTo" ->
        let wrapped = element ("w:" + kind) attrs [runXml]
        match moveKey runObj with
        | Some (key, _, name) ->
            let rangeStart =
                if moveRangeIds.ContainsKey key then ""
                else
                    let rangeId = string nextMoveRangeId
                    nextMoveRangeId <- nextMoveRangeId + 1
                    moveRangeIds.[key] <- rangeId
                    selfClosing ("w:" + kind + "RangeStart") (revisionAttributes (Some rangeId) (optionalString (runObj?RevisionAuthor)) (optionalString (runObj?RevisionDate)) @ [("w:name", name)])
            let left = (match moveRunsLeft.TryGetValue key with | true, count -> count | _ -> 1) - 1
            moveRunsLeft.[key] <- left
            let rangeEnd = if left <= 0 then selfClosing ("w:" + kind + "RangeEnd") [("w:id", moveRangeIds.[key])] else ""
            rangeStart + wrapped + rangeEnd
        | None -> wrapped
    | _ -> runXml

// ---- parts kept from the opened package -----------------------------------

type PreservedRelationshipInfo = { Id: string; Type: string; Target: string; External: bool }

/// The "Preserved" object of the editor package (DocxReader.toEditorPreserved).
type PreservedPackage = {
    /// (path, data)
    Parts: (string * byte[]) list
    RootRelationships: PreservedRelationshipInfo list
    DocumentRelationshipId: string
    PartRelationships: Map<string, PreservedRelationshipInfo list>
    Defaults: (string * string) list
    Overrides: (string * string) list
    MainContentType: string option
    StylesXml: string option
    SettingsXml: string option
    /// The regenerated parts as they were read, by key ("document", a header /
    /// footer relationship id, "footnotes", "endnotes", "comments", "numbering").
    SourceXml: Map<string, string>
    /// All the relationships those parts had (but those to parts the writer makes).
    SourceRelationships: Map<string, PreservedRelationshipInfo list>
    /// Parts only SourceRelationships lead to.
    SourceParts: (string * byte[]) list
    /// The child element of w:body each item of the imported body came from.
    BodyItemChunks: int array
    BodyChunkCount: int
    /// The package path each header / footer part had, by its relationship id.
    PartPaths: Map<string, string>
    /// [Content_Types].xml and relationship parts as they were read, by path.
    OriginalText: Map<string, string>
}

let private emptyPreservedPackage = {
    SourceXml = Map.empty
    SourceRelationships = Map.empty
    SourceParts = []
    BodyItemChunks = [||]
    BodyChunkCount = 0
    PartPaths = Map.empty
    OriginalText = Map.empty
    Parts = []
    RootRelationships = []
    DocumentRelationshipId = "rId1"
    PartRelationships = Map.empty
    Defaults = []
    Overrides = []
    MainContentType = None
    StylesXml = None
    SettingsXml = None
}

let private preservedRelationshipsFromObject (value: obj) : PreservedRelationshipInfo list =
    optionalArray value
    |> Array.toList
    |> List.choose (fun rel ->
        match optionalString (rel?Id), optionalString (rel?Type), optionalString (rel?Target) with
        | Some id, Some relType, Some target -> Some { Id = id; Type = relType; Target = target; External = boolOrDefault false (rel?External) }
        | _ -> None)

let private preservedPackageFromObject (value: obj) : PreservedPackage =
    if isNullOrUndefined value then emptyPreservedPackage
    else
        let partList (parts: obj) =
            optionalArray parts
            |> Array.toList
            |> List.choose (fun part ->
                if isNullOrUndefined part?Data then None
                else optionalString (part?Path) |> Option.map (fun path -> path, unbox<byte[]> part?Data))
        { SourceXml =
            optionalArray (value?SourceXml)
            |> Array.toList
            |> List.choose (fun entry -> Option.map2 (fun key text -> key, text) (optionalString (entry?Part)) (optionalString (entry?Xml)))
            |> Map.ofList
          SourceRelationships =
            optionalArray (value?SourceRelationships)
            |> Array.toList
            |> List.choose (fun entry -> optionalString (entry?Part) |> Option.map (fun key -> key, preservedRelationshipsFromObject (entry?Relationships)))
            |> Map.ofList
          SourceParts = partList (value?SourceParts)
          BodyItemChunks = optionalArray (value?BodyItemChunks) |> Array.map (fun chunk -> intOrDefault (-1) chunk)
          BodyChunkCount = intOrDefault 0 (value?BodyChunkCount)
          PartPaths =
            optionalArray (value?PartPaths)
            |> Array.toList
            |> List.choose (fun entry -> Option.map2 (fun key path -> key, path) (optionalString (entry?Part)) (optionalString (entry?Path)))
            |> Map.ofList
          OriginalText =
            optionalArray (value?OriginalText)
            |> Array.toList
            |> List.choose (fun entry -> Option.map2 (fun path text -> path, text) (optionalString (entry?Path)) (optionalString (entry?Text)))
            |> Map.ofList
          Parts =
            optionalArray (value?Parts)
            |> Array.toList
            |> List.choose (fun part ->
                if isNullOrUndefined part?Data then None
                else optionalString (part?Path) |> Option.map (fun path -> path, unbox<byte[]> part?Data))
          RootRelationships = preservedRelationshipsFromObject (value?RootRelationships)
          DocumentRelationshipId = optionalString (value?DocumentRelationshipId) |> Option.defaultValue "rId1"
          PartRelationships =
            optionalArray (value?PartRelationships)
            |> Array.toList
            |> List.choose (fun entry -> optionalString (entry?Part) |> Option.map (fun key -> key, preservedRelationshipsFromObject (entry?Relationships)))
            |> Map.ofList
          Defaults =
            optionalArray (value?Defaults)
            |> Array.toList
            |> List.choose (fun entry -> Option.map2 (fun extension contentType -> extension, contentType) (optionalString (entry?Extension)) (optionalString (entry?ContentType)))
          Overrides =
            optionalArray (value?Overrides)
            |> Array.toList
            |> List.choose (fun entry -> Option.map2 (fun partName contentType -> partName, contentType) (optionalString (entry?PartName)) (optionalString (entry?ContentType)))
          MainContentType = optionalString (value?MainContentType)
          StylesXml = optionalString (value?StylesXml)
          SettingsXml = optionalString (value?SettingsXml) }

/// Relationship ids and part paths of the kept package: what the writer makes
/// up itself (image / hyperlink ids, media file names) must not collide with them.
let mutable private reservedRelationshipIds : Set<string> = Set.empty
let mutable private reservedPartPaths : Set<string> = Set.empty
/// Kept relationship ids of the part being written: an opaque object is only
/// written when everything it refers to is there.
let mutable private currentPartRelationshipIds : Set<string> = Set.empty

let private freshRelationshipId (candidate: string) =
    let mutable id = candidate
    while Set.contains id reservedRelationshipIds do
        id <- id + "x"
    id

let private preservedRelationshipElement (rel: PreservedRelationshipInfo) =
    selfClosing "Relationship" [
        yield ("Id", rel.Id)
        yield ("Type", rel.Type)
        yield ("Target", rel.Target)
        if rel.External then yield ("TargetMode", "External")
    ]

let private relationshipIdPattern = System.Text.RegularExpressions.Regex("\\s(?:r:[A-Za-z]+|o:relid)=\"([^\"]*)\"")

/// Every paragraph object reachable from a block: text box content, table
/// cells and the tables nested inside cells. The hyperlink, image, comment,
/// note, numbering and style collectors all go through this, so content in a
/// nested table gets its relationships and definitions like any other.
let rec private collectItemParagraphs (item: obj) : obj list =
    match item?Type |> stringOrDefault "Paragraph" with
    | "TextBox" -> optionalArray (item?Paragraphs) |> Array.toList
    | "Table" ->
        optionalArray (item?Rows)
        |> Array.toList
        |> List.collect (fun row ->
            optionalArray (row?Cells)
            |> Array.toList
            |> List.collect (fun cell ->
                (optionalArray (cell?Paragraphs) |> Array.toList |> List.collect collectItemParagraphs)
                @ (optionalArray (cell?NestedTables) |> Array.toList |> List.collect collectItemParagraphs)))
    | _ -> [item]

/// The hyperlink target of a run: the active one, or a target the reader found
/// but does not let anyone follow (InactiveHyperlink). Both are written back.
let private runHyperlinkTarget (run: obj) =
    optionalString (run?Hyperlink) |> Option.orElse (optionalString (run?InactiveHyperlink))

/// Targets that need a relationship: everything but links to a bookmark of the
/// document itself ("#name"), which are written as w:anchor.
let private collectParagraphHyperlinks (para: obj) =
    optionalArray (para?Runs)
    |> Array.toList
    |> List.choose runHyperlinkTarget
    |> List.filter (fun target -> not (target.StartsWith "#"))

let private collectItemHyperlinks (item: obj) =
    collectItemParagraphs item |> List.collect collectParagraphHyperlinks

/// The relationship id of every hyperlink target of the items. A target the
/// part already has a relationship for (`existing`: the relationships it was
/// read with, when they are written back) keeps that relationship's id, so a
/// block written from the model refers to the link the file had and leaves no
/// relationship behind that nothing refers to.
let private hyperlinkMapWith (existing: PreservedRelationshipInfo list) (items: obj array) =
    let known =
        existing
        |> List.filter (fun rel -> rel.External && rel.Type.EndsWith "/hyperlink")
    items
    |> Array.toList
    |> List.collect collectItemHyperlinks
    |> List.distinct
    // Not "rId<n>": header/footer parts keep the relationship ids they were read
    // with (rId7, ...), and a hyperlink numbered into the same range pointed at
    // header1.xml after saving.
    |> List.mapi (fun index url ->
        match known |> List.tryFind (fun rel -> rel.Target = url) with
        | Some rel -> (url, rel.Id)
        | None -> (url, freshRelationshipId (sprintf "rIdLink%d" (index + 1))))
    |> Map.ofList

let private hyperlinkMap (items: obj array) = hyperlinkMapWith [] items

type CommentRef = { Id: string; Text: string; Author: string option; Date: string option }
type NoteRef = { Id: string; NoteType: string; Text: string }

type ImageAsset = {
    Src: string
    RelId: string
    FileName: string
    Extension: string
    ContentType: string
    Data: byte[] option
    Linked: bool
}

type ImageScope = {
    Assets: ImageAsset array
    AssetBySrc: Map<string, ImageAsset>
}

let private emptyImageScope = { Assets = [||]; AssetBySrc = Map.empty }

[<Emit("typeof Buffer !== 'undefined' ? Buffer.from($0, 'base64') : Uint8Array.from(atob($0), c => c.charCodeAt(0))")>]
let private base64ToBytes (_base64: string) : byte[] = jsNative

let private mimeToExtensionAndContentType (mime: string) =
    match mime.ToLowerInvariant() with
    | "image/jpeg" | "image/jpg" -> Some ("jpeg", "image/jpeg")
    | "image/png" -> Some ("png", "image/png")
    | "image/gif" -> Some ("gif", "image/gif")
    | "image/bmp" -> Some ("bmp", "image/bmp")
    | "image/svg+xml" -> Some ("svg", "image/svg+xml")
    | "image/webp" -> Some ("webp", "image/webp")
    | "image/x-emf" | "image/emf" -> Some ("emf", "image/x-emf")
    | "image/x-wmf" | "image/wmf" -> Some ("wmf", "image/x-wmf")
    | "image/tiff" -> Some ("tiff", "image/tiff")
    | _ -> None

let private parseDataUrl (src: string) =
    if src.StartsWith("data:") then
        let commaIndex = src.IndexOf(',')
        if commaIndex > 0 then
            let meta = src.Substring(5, commaIndex - 5)
            let data = src.Substring(commaIndex + 1)
            let parts = meta.Split(';')
            let mime = if parts.Length > 0 then parts.[0] else ""
            if meta.ToLowerInvariant().Contains("base64") then
                mimeToExtensionAndContentType mime
                |> Option.map (fun (extension, contentType) -> extension, contentType, base64ToBytes data)
            else None
        else None
    else None

let private markerText defaultValue (value: obj) =
    optionalString value |> Option.defaultValue defaultValue

let private collectParagraphComments (para: obj) =
    if isNullOrUndefined para?Runs then
        []
    else
        para?Runs
        |> unbox<obj array>
        |> Array.toList
        |> List.choose (fun run ->
            optionalString (run?CommentId)
            |> Option.map (fun id -> {
                Id = id
                Text = markerText (sprintf "Comment %s" id) (run?CommentText)
                Author = optionalString (run?CommentAuthor)
                Date = optionalString (run?CommentDate)
            }))

let private collectParagraphNotes (para: obj) =
    if isNullOrUndefined para?Runs then
        []
    else
        para?Runs
        |> unbox<obj array>
        |> Array.toList
        |> List.choose (fun run ->
            optionalString (run?NoteId)
            |> Option.map (fun id ->
                let noteType = markerText "footnote" (run?NoteType)
                { Id = id; NoteType = noteType; Text = markerText (sprintf "%s %s" noteType id) (run?NoteText) }))

let private collectComments (items: obj array) =
    items
    |> Array.toList
    |> List.collect collectItemParagraphs
    |> List.collect collectParagraphComments
    |> List.distinctBy (fun comment -> comment.Id)

let private collectNotes (noteType: string) (items: obj array) =
    items
    |> Array.toList
    |> List.collect collectItemParagraphs
    |> List.collect collectParagraphNotes
    |> List.filter (fun note -> note.NoteType = noteType)
    |> List.distinctBy (fun note -> note.Id)

let private collectSectionReferenceIds (items: obj array) =
    let idsFromRefs (value: obj) =
        if isNullOrUndefined value then []
        else
            value
            |> optionalArray
            |> Array.toList
            |> List.choose (fun item -> optionalString (item?RelationshipId))

    items
    |> Array.toList
    |> List.collect collectItemParagraphs
    |> List.collect (fun para -> idsFromRefs (para?HeaderReferences) @ idsFromRefs (para?FooterReferences))
    |> Set.ofList

let private collectParagraphImageSources (para: obj) =
    if isNullOrUndefined para?Runs then []
    else
        para?Runs
        |> unbox<obj array>
        |> Array.toList
        |> List.choose (fun run ->
            // An opaque object's picture is one of the kept parts, not a new image.
            if boolOrDefault false (run?IsImage) && isNullOrUndefined run?OpaqueXml then optionalString (run?ImageSrc)
            else None)

let private collectItemImageSources (item: obj) =
    collectItemParagraphs item |> List.collect collectParagraphImageSources

let private buildImageScope (mediaStartIndex: int) (items: obj array) =
    let mutable nextMediaIndex = mediaStartIndex
    let assets =
        items
        |> Array.toList
        |> List.collect collectItemImageSources
        |> List.distinct
        |> List.choose (fun src ->
            parseDataUrl src
            |> Option.map (fun (extension, contentType, data) ->
                // Not the name of a kept part (a fallback picture of an opaque object).
                while Set.contains (sprintf "word/media/image%d.%s" nextMediaIndex extension) reservedPartPaths do
                    nextMediaIndex <- nextMediaIndex + 1
                let fileName = sprintf "image%d.%s" nextMediaIndex extension
                let relId = freshRelationshipId (sprintf "rIdImage%d" nextMediaIndex)
                nextMediaIndex <- nextMediaIndex + 1
                { Src = src; RelId = relId; FileName = fileName; Extension = extension; ContentType = contentType; Data = Some data; Linked = false })
            |> Option.orElseWith (fun () ->
                if src.Contains "://" && not (src.StartsWith "data:") && not (src.StartsWith "blob:") then
                    let relId = freshRelationshipId (sprintf "rIdImage%d" nextMediaIndex)
                    nextMediaIndex <- nextMediaIndex + 1
                    Some { Src = src; RelId = relId; FileName = ""; Extension = ""; ContentType = ""; Data = None; Linked = true }
                else None))
        |> List.toArray
    { Assets = assets; AssetBySrc = assets |> Array.map (fun asset -> asset.Src, asset) |> Map.ofArray }, nextMediaIndex

/// Reader: emuToPixels = emu / 9525.
let private pixelsToEmus (pixels: float) =
    roundToInt (pixels * 9525.0)

let private optionalImageSize defaultValue (value: obj) =
    optionalFloat value |> Option.defaultValue defaultValue |> pixelsToEmus |> string

let private optionalImagePosition (value: obj) =
    optionalFloat value |> Option.defaultValue 0.0 |> pixelsToEmus |> string

let private optionalImageDistance defaultValue (value: obj) =
    optionalFloat value |> Option.defaultValue defaultValue |> pixelsToEmus |> string

let private boolFlag defaultValue (value: obj) =
    if boolOrDefault defaultValue value then "1" else "0"

/// wp:wrapNone / wrapSquare / wrapTight / wrapThrough / wrapTopAndBottom of an anchor.
/// Tight and through have a wrap polygon: the object's own (`points`: x, y, x, y, ... in
/// the file's units, wp:start first), or the bounding box when it has none.
let private anchorWrapElement (wrap: string option) (wrapText: string option) (points: float array) =
    let side = wrapText |> Option.defaultValue "bothSides"
    let number (value: float) = string (int (System.Math.Round value))
    let polygon =
        if points.Length >= 6 then
            element "wp:wrapPolygon" [("edited", "0")] [
                for i in 0 .. points.Length / 2 - 1 ->
                    selfClosing (if i = 0 then "wp:start" else "wp:lineTo") [("x", number points.[i * 2]); ("y", number points.[i * 2 + 1])] ]
        else
        element "wp:wrapPolygon" [("edited", "0")] [
            selfClosing "wp:start" [("x", "0"); ("y", "0")]
            selfClosing "wp:lineTo" [("x", "0"); ("y", "21600")]
            selfClosing "wp:lineTo" [("x", "21600"); ("y", "21600")]
            selfClosing "wp:lineTo" [("x", "21600"); ("y", "0")]
            selfClosing "wp:lineTo" [("x", "0"); ("y", "0")]
        ]
    match wrap with
    | Some "none" -> selfClosing "wp:wrapNone" []
    | Some "topAndBottom" -> selfClosing "wp:wrapTopAndBottom" []
    | Some "tight" -> element "wp:wrapTight" [("wrapText", side)] [polygon]
    | Some "through" -> element "wp:wrapThrough" [("wrapText", side)] [polygon]
    | _ -> selfClosing "wp:wrapSquare" [("wrapText", side)]

/// wp:positionH / wp:positionV: by alignment when the object has one, else by offset.
let private anchorPositionElement (name: string) (relativeFrom: string) (align: string option) (offset: string) =
    match align with
    | Some value when value <> "" -> element name [("relativeFrom", relativeFrom)] [elementWithText "wp:align" [] value]
    | _ -> element name [("relativeFrom", relativeFrom)] [elementWithText "wp:posOffset" [] offset]

let private generateImageRun (asset: ImageAsset) (runObj: obj) =
    let width = optionalImageSize 96.0 (runObj?ImageWidth)
    let height = optionalImageSize 48.0 (runObj?ImageHeight)
    let altAttrs = optionalString (runObj?ImageAlt) |> Option.map (fun alt -> [("descr", alt)]) |> Option.defaultValue []
    let imageName = optionalString (runObj?ImageName) |> Option.defaultValue asset.FileName
    let titleAttrs = optionalString (runObj?ImageTitle) |> Option.map (fun title -> [("title", title)]) |> Option.defaultValue []
    let positionX = optionalImagePosition (runObj?ImagePositionX)
    let positionY = optionalImagePosition (runObj?ImagePositionY)
    let positionHRelativeFrom = optionalString (runObj?ImagePositionHRelativeFrom) |> Option.defaultValue "column"
    let positionVRelativeFrom = optionalString (runObj?ImagePositionVRelativeFrom) |> Option.defaultValue "paragraph"
    let distanceTop = optionalImageDistance 0.0 (runObj?ImageDistanceTop)
    let distanceBottom = optionalImageDistance 0.0 (runObj?ImageDistanceBottom)
    let distanceLeft = optionalImageDistance 12.0 (runObj?ImageDistanceLeft)
    let distanceRight = optionalImageDistance 12.0 (runObj?ImageDistanceRight)
    let behindDoc = boolFlag false (runObj?ImageBehindDoc)
    let allowOverlap = boolFlag true (runObj?ImageAllowOverlap)
    let layoutInCell = boolFlag true (runObj?ImageLayoutInCell)
    let relativeHeight = optionalString (runObj?ImageRelativeHeight) |> Option.defaultValue "0"
    let locked = boolFlag false (runObj?ImageLocked)
    let simplePosition = boolFlag false (runObj?ImageSimplePosition)
    let simplePositionX = optionalImagePosition (runObj?ImageSimplePositionX)
    let simplePositionY = optionalImagePosition (runObj?ImageSimplePositionY)
    let cropAttrs =
        [ ("l", optionalString (runObj?ImageCropLeft))
          ("t", optionalString (runObj?ImageCropTop))
          ("r", optionalString (runObj?ImageCropRight))
          ("b", optionalString (runObj?ImageCropBottom)) ]
        |> List.choose (fun (name, value) -> value |> Option.map (fun text -> name, text))
    let drawingContainer =
        if boolOrDefault false (runObj?ImageFloating) then
            let wrapPoints = (let points = runObj?ImageWrapPolygon in if isNullOrUndefined points then [||] else unbox<float array> points)
            let wrap = anchorWrapElement (optionalString (runObj?ImageWrap)) (optionalString (runObj?ImageWrapText)) wrapPoints
            fun children ->
                element "wp:anchor" [
                    ("xmlns:wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")
                    ("distT", distanceTop); ("distB", distanceBottom); ("distL", distanceLeft); ("distR", distanceRight)
                    ("simplePos", simplePosition); ("relativeHeight", relativeHeight); ("behindDoc", behindDoc)
                    ("locked", locked); ("layoutInCell", layoutInCell); ("allowOverlap", allowOverlap)
                ] ([
                    selfClosing "wp:simplePos" [("x", simplePositionX); ("y", simplePositionY)]
                    anchorPositionElement "wp:positionH" positionHRelativeFrom (optionalString (runObj?ImagePositionHAlign)) positionX
                    anchorPositionElement "wp:positionV" positionVRelativeFrom (optionalString (runObj?ImagePositionVAlign)) positionY
                    selfClosing "wp:extent" [("cx", width); ("cy", height)]
                    wrap
                ] @ children)
        else
            fun children ->
                element "wp:inline" [("xmlns:wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")] (
                    [selfClosing "wp:extent" [("cx", width); ("cy", height)]] @ children)
    element "w:r" [] [
        element "w:drawing" [] [
            drawingContainer [
                selfClosing "wp:docPr" ([("id", (asset.RelId.Replace("rIdImage", "").TrimEnd('x'))); ("name", imageName)] @ altAttrs @ titleAttrs)
                element "a:graphic" [("xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main")] [
                    element "a:graphicData" [("uri", "http://schemas.openxmlformats.org/drawingml/2006/picture")] [
                        element "pic:pic" [("xmlns:pic", "http://schemas.openxmlformats.org/drawingml/2006/picture")] [
                            element "pic:nvPicPr" [] [
                                selfClosing "pic:cNvPr" ([("id", (asset.RelId.Replace("rIdImage", "").TrimEnd('x'))); ("name", imageName)] @ altAttrs @ titleAttrs)
                                selfClosing "pic:cNvPicPr" []
                            ]
                            element "pic:blipFill" [] [
                                selfClosing "a:blip" [((if asset.Linked then "r:link" else "r:embed"), asset.RelId)]
                                if not cropAttrs.IsEmpty then selfClosing "a:srcRect" cropAttrs
                                element "a:stretch" [] [selfClosing "a:fillRect" []]
                            ]
                            element "pic:spPr" [] [
                                element "a:xfrm" [] [
                                    selfClosing "a:off" [("x", "0"); ("y", "0")]
                                    selfClosing "a:ext" [("cx", width); ("cy", height)]
                                ]
                                selfClosing "a:prstGeom" [("prst", "rect")]
                            ]
                        ]
                    ]
                ]
            ]
        ]
    ]

/// A first-page or even-page header that is switched off (no w:titlePg / no
/// evenAndOddHeaders) is still part of the section in the file; Word keeps the
/// reference so the header comes back when the option is turned on again.
/// Dropping the reference orphaned the part and it reopened as a second
/// "default" header.
let private includeHeaderFooterKind (_differentFirstPage: bool) (_evenAndOddHeaders: bool) (kind: string) =
    kind = "default" || kind = "first" || kind = "even"

let private sectionReferences elementName differentFirstPage evenAndOddHeaders (value: obj) =
    if isNullOrUndefined value then []
    else
        value
        |> optionalArray
        |> Array.toList
        |> List.distinctBy (fun item -> item?Kind |> stringOrDefault (item?Type |> stringOrDefault "default"))
        |> List.choose (fun item ->
            let kind = item?Kind |> stringOrDefault (item?Type |> stringOrDefault "default")
            if includeHeaderFooterKind differentFirstPage evenAndOddHeaders kind then
                optionalString (item?RelationshipId)
                |> Option.map (fun relationshipId ->
                    selfClosing elementName [("w:type", kind); ("r:id", relationshipId)])
            else None)

/// The w:sectPr of a paragraph that ends a section: one with a section break,
/// section geometry or section header / footer references. It describes the
/// section that ends with the paragraph.
let private paragraphSectionProperties (packageSetup: PageSetupInfo) (para: obj) =
    // The reader clears SectionBreak on an empty next-page paragraph that is
    // followed by a continuous one, but that paragraph still ends a section and
    // still carries its geometry; it is written back with the default type.
    let sectionBreak = optionalString (para?SectionBreak)
    let hasReferences = (optionalArray (para?HeaderReferences)).Length > 0 || (optionalArray (para?FooterReferences)).Length > 0
    if sectionBreak.IsNone && isNullOrUndefined para?SectionPageSetup && not hasReferences then None
    else
        let paraSetup =
            optionalPageSetup (para?SectionPageSetup)
            |> Option.orElseWith (fun () -> optionalPageSetup (para?PageSetup))
            // A sectPr without pgSz/pgMar is US Letter with default margins to Word,
            // so a section with unknown geometry takes the document's.
            |> Option.defaultValue { packageSetup with DifferentFirstPage = false; ColumnCount = 1; ColumnWidths = []; SectionStart = None }
            |> fun setup -> { setup with HeaderReferences = []; FooterReferences = []; PageNumberStart = None; PageNumberFormat = None }
        let setup =
            let columns = para?SectionColumns
            if isNullOrUndefined columns then paraSetup
            else
                { paraSetup with
                    ColumnCount = optionalFloat (columns?Count) |> Option.map roundToInt |> Option.defaultValue 1
                    ColumnSpace = floatOrDefault 0.0 (columns?Gap)
                    ColumnSeparator = boolOrDefault false (columns?Separator)
                    ColumnWidths =
                        optionalArray (columns?Widths)
                        |> Array.toList
                        |> List.map (fun column -> floatOrDefault 0.0 (column?W), floatOrDefault 0.0 (column?S)) }
        let refs =
            sectionReferences "w:headerReference" setup.DifferentFirstPage setup.EvenAndOddHeaders (para?HeaderReferences)
            @ sectionReferences "w:footerReference" setup.DifferentFirstPage setup.EvenAndOddHeaders (para?FooterReferences)
        let pageNumberStart = optionalInt (para?SectionPageNumberStart)
        Some (element "w:sectPr" [] (sectionPropertiesChildren refs sectionBreak (Some setup) pageNumberStart (optionalString (para?SectionPageNumberFormat))))

let private defaultFontSizePixels = 11.0 * 96.0 / 72.0

/// The reader's name for the document defaults (w:docDefaults) in its style
/// list, and the style id it gives a table cell paragraph that only has those.
/// It is not a style of the file: never written as w:pStyle or as a w:style.
let private documentDefaultsStyleId = "__docDefaults"

/// numIds used for list paragraphs that carry no ListId (lists made in the
/// editor): (bullet, numbered). Chosen per document in numberingPlan so they
/// never collide with a list id that came from the file.
let mutable private defaultListNumIds = ("1", "2")

/// The font size the reader divided exact/at-least line spacing by: the
/// paragraph mark size it reported, else the largest text run.
let private paragraphFontSize (para: obj) : float =
    match optionalFloat (para?ParagraphMarkFontSize) with
    | Some size when size > 0.0 -> size
    | _ ->
        optionalArray (para?Runs)
        |> Array.choose (fun run ->
            if boolOrDefault false (run?IsImage) then None else optionalFloat (run?FontSize))
        |> fun sizes -> if sizes.Length = 0 then defaultFontSizePixels else Array.max sizes

let private paragraphBorderElement name (value: obj) =
    if isNullOrUndefined value then None
    else
        let style = optionalString (value?Style) |> Option.defaultValue "single"
        let hidden = style = "none" || style = "nil"
        let width =
            match optionalFloat (value?Width) with
            | Some amount -> InverseConversion.pixelsToEighthPoints amount
            | None -> if hidden then "0" else "4"
        // w:space is in points on both sides (reader: paragraphBorderSpaceToPoints).
        let space = optionalFloat (value?Space) |> Option.map (fun amount -> string (max 0 (roundToInt amount))) |> Option.defaultValue "0"
        let color = optionalString (value?Color) |> InverseConversion.cssColorToWord |> Option.defaultValue "auto"
        Some (selfClosing name [("w:val", style); ("w:sz", width); ("w:space", space); ("w:color", color)])

let private paragraphBordersElement (value: obj) =
    if isNullOrUndefined value then None
    else
        // CT_PBdr order: top, left, bottom, right.
        let sides =
            [ paragraphBorderElement "w:top" (value?Top)
              paragraphBorderElement "w:left" (value?Left)
              paragraphBorderElement "w:bottom" (value?Bottom)
              paragraphBorderElement "w:right" (value?Right) ]
            |> List.choose id
        if sides.IsEmpty then None else Some (element "w:pBdr" [] sides)

let private tabStopsElement (para: obj) =
    let tab (alignment: string) (leader: string option) (position: float) =
        selfClosing "w:tab" [
            yield ("w:val", (if alignment = "" then "left" else alignment))
            match leader with
            | Some value when value <> "" && value <> "none" -> yield ("w:leader", value)
            | _ -> ()
            yield ("w:pos", InverseConversion.pixelsToTwips position)
        ]
    let stops =
        optionalArray (para?TabStops)
        |> Array.toList
        |> List.choose (fun stop ->
            optionalFloat (stop?Position)
            |> Option.map (fun position -> tab (stop?Alignment |> stringOrDefault "left") (optionalString (stop?Leader)) position))
    let stops =
        // RightTabPosition is the reader's shorthand for the last right tab; only
        // needed when the full list is not there (older exports, editor-made TOC lines).
        if not stops.IsEmpty then stops
        else
            match optionalFloat (para?RightTabPosition) with
            | Some position -> [ tab "right" (optionalString (para?RightTabLeader)) position ]
            | None -> []
    if stops.IsEmpty then None else Some (element "w:tabs" [] stops)

/// A paragraph object that says which of its properties are its own
/// (ExplicitParagraphFormat, with ExplicitSpacingBefore, ExplicitIndentLeft,
/// ...): only those are written. An object without it has effective values
/// only, and all of them are written (callers that predate the distinction).
let private statesExplicitFormat (para: obj) = boolOrDefault false (para?ExplicitParagraphFormat)

/// The value of a measure to write: the explicit one, or for an object that
/// only has effective values, the effective one.
let private writtenMeasure (para: obj) (explicitValue: obj) (effective: obj) : float option =
    if statesExplicitFormat para then optionalFloat explicitValue else optionalFloat effective

let private paragraphMarkRunProps (para: obj) =
    // Exact / at-least line spacing is stored relative to the paragraph font
    // size, so pin that size on the paragraph mark: the reader then divides by
    // the same number next time and the value is a fixed point.
    // An object that tells its own formatting from the inherited keeps its run
    // sizes as they are, so the reader divides by the same size again without a pin
    // (a pin would give the paragraph mark formatting it did not have).
    let pinned =
        match optionalString (para?LineSpacingRule), writtenMeasure para (para?ExplicitLineSpacing) (para?LineSpacing) with
        | Some rule, Some _ when (rule = "exact" || rule = "atLeast") && not (statesExplicitFormat para) -> Some (paragraphFontSize para)
        | _ -> None
    // A paragraph mark that had run properties keeps a w:rPr (with the size it
    // stated): the mark's size sets the height of an empty paragraph.
    let stated = if boolOrDefault false (para?HasMarkProperties) then Some (optionalFloat (para?ExplicitMarkFontSize)) else None
    // A tracked paragraph mark (Enter, or two paragraphs joined, under track
    // changes): w:ins / w:del come first in the mark's run properties.
    let tracked =
        match optionalString (para?MarkRevisionType) with
        | Some revisionType when revisionType = "insert" || revisionType = "delete" ->
            [ selfClosing (if revisionType = "delete" then "w:del" else "w:ins")
                  (revisionAttributes (optionalString (para?MarkRevisionId)) (optionalString (para?MarkRevisionAuthor)) (optionalString (para?MarkRevisionDate))) ]
        | _ -> []
    let size (value: float) =
        let sz = InverseConversion.pixelsToHalfPoints value
        [selfClosing "w:sz" [("w:val", sz)]; selfClosing "w:szCs" [("w:val", sz)]]
    match pinned, stated with
    | Some value, _
    | None, Some (Some value) -> Some (element "w:rPr" [] (tracked @ size value))
    | None, Some None -> Some (element "w:rPr" [] tracked)
    | None, None -> if tracked.IsEmpty then None else Some (element "w:rPr" [] tracked)

/// Paragraph properties in CT_PPr sequence order: pStyle, keepNext, keepLines,
/// pageBreakBefore, widowControl, numPr, pBdr, shd, tabs, bidi, spacing, ind,
/// contextualSpacing, jc, rPr, sectPr. Like w:rPr this is an xsd:sequence.
let private generateParaPropsWithPageSetup (packageSetup: PageSetupInfo) (para: obj) : string =
    // HasExplicit: the paragraph says whether its alignment is direct
    // formatting (ExplicitAlignment) or inherited from the style (null).
    let alignment =
        if boolOrDefault false (para?HasExplicit) then
            optionalString (para?ExplicitAlignment) |> Option.map InverseConversion.alignmentToWord
        else
            Some (para?Alignment |> stringOrDefault "left" |> InverseConversion.alignmentToWord)
    // Spacing stated on the paragraph is a measure, never "automatic": a style
    // with beforeAutospacing / afterAutospacing would otherwise override the value.
    let spacingAttributes = [
        match writtenMeasure para (para?ExplicitSpacingBefore) (para?SpacingBefore) with
        | Some value ->
            yield ("w:before", InverseConversion.pointsToTwips value)
            yield ("w:beforeAutospacing", "0")
        | None -> ()
        match writtenMeasure para (para?ExplicitSpacingAfter) (para?SpacingAfter) with
        | Some value ->
            yield ("w:after", InverseConversion.pointsToTwips value)
            yield ("w:afterAutospacing", "0")
        | None -> ()
        match writtenMeasure para (para?ExplicitLineSpacing) (para?LineSpacing) with
        | Some value ->
            match optionalString (para?LineSpacingRule) with
            | Some rule when rule = "exact" || rule = "atLeast" ->
                yield ("w:line", InverseConversion.fixedLineSpacingToTwips (paragraphFontSize para) value)
                yield ("w:lineRule", rule)
            | _ ->
                yield ("w:line", InverseConversion.lineSpacingToWord value)
                yield ("w:lineRule", "auto")
        | None -> ()
    ]
    let indentAttributes = [
        match writtenMeasure para (para?ExplicitIndentLeft) (para?IndentLeft) with
        | Some value -> yield ("w:left", InverseConversion.pixelsToTwips value)
        | None -> ()
        match writtenMeasure para (para?ExplicitIndentRight) (para?IndentRight) with
        | Some value -> yield ("w:right", InverseConversion.pixelsToTwips value)
        | None -> ()
        match writtenMeasure para (para?ExplicitFirstLineIndent) (para?FirstLineIndent) with
        | Some value when value < 0.0 -> yield ("w:hanging", InverseConversion.pixelsToTwips -value)
        // An explicit 0 matters: it overrides a first-line indent from the style.
        | Some value -> yield ("w:firstLine", InverseConversion.pixelsToTwips value)
        | None -> ()
    ]
    let listStyle = para?ListStyle |> stringOrDefault "none"
    let listLevel = para?ListLevel |> intOrDefault 0
    let itemType = para?Type |> stringOrDefault "Paragraph"
    let headingLevel = para?HeadingLevel |> intOrDefault 0
    let onOff name (value: obj) =
        match optionalBool value with
        | Some isOn -> [toggle name isOn]
        | None -> []

    let children = [
        match optionalString (para?Style) with
        | Some styleId when styleId <> "" && styleId <> documentDefaultsStyleId -> yield selfClosing "w:pStyle" [("w:val", styleId)]
        | Some styleId when styleId = documentDefaultsStyleId -> ()
        | _ when itemType = "Heading" && headingLevel >= 1 && headingLevel <= 3 ->
            yield selfClosing "w:pStyle" [("w:val", sprintf "Heading%d" headingLevel)]
        | _ -> ()
        // The paragraph's own keepNext / keepLines. A paragraph object without
        // direct-formatting information (HasExplicit) only has the effective
        // value, under StyleKeepNext / StyleKeepLines.
        let keepValue (direct: obj) (effective: obj) =
            if boolOrDefault false (para?HasExplicit) || not (isNullOrUndefined direct) then direct else effective
        yield! onOff "w:keepNext" (keepValue (para?KeepNext) (para?StyleKeepNext))
        yield! onOff "w:keepLines" (keepValue (para?KeepLines) (para?StyleKeepLines))
        // (Not the page break of the paragraph's style, nor the one the reader
        // gives a heading that started a page: written, either would be the
        // paragraph's own from then on.)
        if boolOrDefault false (para?PageBreakBefore) && not (boolOrDefault false (para?PageBreakBeforeDerived)) then yield selfClosing "w:pageBreakBefore" []
        yield! onOff "w:widowControl" (para?WidowControl)
        if listStyle = "bullet" || listStyle = "numbered" then
            let numId = optionalIntString (para?ListId) |> Option.defaultValue (if listStyle = "bullet" then fst defaultListNumIds else snd defaultListNumIds)
            yield element "w:numPr" [] [
                selfClosing "w:ilvl" [("w:val", string (max 0 (listLevel - 1)))]
                selfClosing "w:numId" [("w:val", numId)]
            ]
        elif optionalString (para?ListId) = Some "0" then
            // numId 0 switches off numbering inherited from the paragraph style.
            yield element "w:numPr" [] [selfClosing "w:numId" [("w:val", "0")]]
        match paragraphBordersElement (para?ParagraphBorders) with
        | Some borders -> yield borders
        | None -> ()
        match (if statesExplicitFormat para then optionalString (para?ExplicitBackgroundColor) else optionalString (para?BackgroundColor)) with
        | Some color ->
            let wordColor = InverseConversion.cssColorToWord (Some color) |> Option.defaultValue color
            yield selfClosing "w:shd" [("w:val", "clear"); ("w:color", "auto"); ("w:fill", wordColor)]
        | None -> ()
        // Tab stops a paragraph inherits from its style stay in the style.
        match (if statesExplicitFormat para && not (boolOrDefault false (para?ExplicitTabs)) then None else tabStopsElement para) with
        | Some tabs -> yield tabs
        | None -> ()
        if boolOrDefault false (para?Rtl) then yield selfClosing "w:bidi" []
        if not spacingAttributes.IsEmpty then yield selfClosing "w:spacing" spacingAttributes
        if not indentAttributes.IsEmpty then yield selfClosing "w:ind" indentAttributes
        yield! onOff "w:contextualSpacing" (if statesExplicitFormat para then para?ExplicitContextualSpacing else para?StyleContextualSpacing)
        match alignment with
        | Some value -> yield selfClosing "w:jc" [("w:val", value)]
        | None -> ()
        match paragraphMarkRunProps para with
        | Some rPr -> yield rPr
        | None -> ()
        match paragraphSectionProperties packageSetup para with
        | Some sectPr -> yield sectPr
        | None -> ()
    ]

    element "w:pPr" [] children

/// Generate paragraph properties XML (handles JS objects)
let generateParaProps (para: obj) : string =
    generateParaPropsWithPageSetup defaultPageSetupInfo para

/// Generate a paragraph element with formatting (handles JS objects)
/// The run of the editor's package as the record the run writers take.
let private editorRunOfObject (runObj: obj) : EditorRun =
    {
      Text = runObj?Text |> unbox<string>
      Bold = boolOrDefault false (runObj?Bold)
      Italic = boolOrDefault false (runObj?Italic)
      Underline = boolOrDefault false (runObj?Underline)
      Strike = boolOrDefault false (runObj?Strike)
      FontSize = optionalFloat (runObj?FontSize) |> Option.defaultValue defaultFontSizePixels
      CharacterSpacing = optionalFloat (runObj?CharacterSpacing) |> Option.defaultValue 0.0
      FontFamily = runObj?FontFamily |> stringOrDefault "Calibri"
      Color = if isNullOrUndefined runObj?Color then None else unbox<string option> runObj?Color
      Highlight = optionalString (runObj?Highlight)
      VertAlign = optionalString (runObj?VertAlign)
      CommentId = optionalString (runObj?CommentId)
      CommentText = optionalString (runObj?CommentText)
      CommentAuthor = optionalString (runObj?CommentAuthor)
      CommentDate = optionalString (runObj?CommentDate)
      NoteId = optionalString (runObj?NoteId)
      NoteType = optionalString (runObj?NoteType)
      NoteText = optionalString (runObj?NoteText)
      Hyperlink = optionalString (runObj?Hyperlink)
      BookmarkStartId = optionalString (runObj?BookmarkStartId)
      BookmarkName = optionalString (runObj?BookmarkName)
      BookmarkEndId = optionalString (runObj?BookmarkEndId)
      FieldCode = optionalString (runObj?FieldCode)
      FieldResult = optionalString (runObj?FieldResult)
      EquationXml = optionalString (runObj?EquationXml)
      RenderedPageBreakBefore = boolOrDefault false (runObj?RenderedPageBreakBefore)
      IsImage = false
      ImageSrc = None
      ImageLinked = false
      ImageWidth = None
      ImageHeight = None
      ImageAlt = None
      ImageName = None
      ImageTitle = None
      ImageFloating = false
      ImageWrap = None
      ImagePositionX = None
      ImagePositionY = None
      ImagePositionHRelativeFrom = None
      ImagePositionVRelativeFrom = None
      ImageBehindDoc = None
      ImageAllowOverlap = None
      ImageLayoutInCell = None
      ImageDistanceTop = None
      ImageDistanceBottom = None
      ImageDistanceLeft = None
      ImageDistanceRight = None
      ImageRelativeHeight = None
      ImageLocked = None
      ImageSimplePosition = None
      ImageSimplePositionX = None
      ImageSimplePositionY = None
      ImageCropLeft = None
      ImageCropTop = None
      ImageCropRight = None
      ImageCropBottom = None
    }

/// The properties the model writes for a run of the editor's package (none for a picture).
let private modelRunProperties (runObj: obj) : string list =
    if boolOrDefault false (runObj?IsImage) || not (isNullOrUndefined runObj?OpaqueXml) then []
    else orderedRunProps (runExtrasFromObject runObj) (editorRunOfObject runObj)

let private generateParagraphWithHyperlinksAndImagesWithPageSetup (packageSetup: PageSetupInfo) (hyperlinks: Map<string, string>) (images: ImageScope) (para: obj) : string =
    let pPr = generateParaPropsWithPageSetup packageSetup para
    // The mark of a comment or a note is one element: where the editor holds it as two runs
    // (a part of the mark was formatted) it is still one reference, written with the first.
    let runsArray =
        let all = para?Runs |> unbox<obj array>
        let sameMark (a: obj) (b: obj) =
            match optionalString (a?CommentId), optionalString (b?CommentId) with
            | Some x, Some y -> x = y
            | None, None ->
                (match optionalString (a?NoteId), optionalString (b?NoteId) with
                 | Some x, Some y -> x = y && optionalString (a?NoteType) = optionalString (b?NoteType)
                 | _ -> false)
            | _ -> false
        all |> Array.mapi (fun index run -> index, run) |> Array.filter (fun (index, run) -> index = 0 || not (sameMark all.[index - 1] run)) |> Array.map snd
    // The result of a field held as several runs (a part of it was formatted): one field, see fieldResultParts.
    let fieldParts = System.Collections.Generic.Dictionary<int, (obj * EditorRun) list>()
    let fieldPartRuns = System.Collections.Generic.HashSet<int>()
    (let isPlainText (run: obj) = not (boolOrDefault false (run?IsImage)) && isNullOrUndefined run?OpaqueXml && not (boolOrDefault false (run?IsEquation))
     let wrapper (run: obj) = (runHyperlinkTarget run, optionalString (run?RevisionType), optionalString (run?RevisionId), optionalString (run?CommentId), optionalString (run?NoteId))
     let marked (run: obj) = (optionalString (run?BookmarkStartId)).IsSome || (optionalString (run?BookmarkEndId)).IsSome || (optionalString (run?BookmarkId)).IsSome
     let mutable index = 0
     while index < runsArray.Length do
         let head = runsArray.[index]
         let mutable next = index + 1
         match (if isPlainText head then optionalString (head?FieldCode) else None), optionalString (head?FieldResult) with
         | Some code, Some whole when code <> "" ->
             let mutable text = stringOrDefault "" (head?Text)
             while next < runsArray.Length
                   && isPlainText runsArray.[next]
                   && optionalString (runsArray.[next]?FieldCode) = Some code
                   && optionalString (runsArray.[next]?FieldResult) = Some whole
                   && wrapper runsArray.[next] = wrapper head
                   && not (marked runsArray.[next])
                   && text.Length < whole.Length do
                 text <- text + stringOrDefault "" (runsArray.[next]?Text)
                 next <- next + 1
             if next > index + 1 && text = whole then
                 fieldParts.[index] <- [ for part in index + 1 .. next - 1 -> runsArray.[part], editorRunOfObject runsArray.[part] ]
                 for part in index + 1 .. next - 1 do fieldPartRuns.Add part |> ignore
             else next <- index + 1
         | _ -> ()
         index <- next)
    let runs =
        runsArray
        |> Array.toList
        |> List.mapi (fun runIndex runObj -> runIndex, runObj)
        |> List.filter (fun (runIndex, _) -> not (fieldPartRuns.Contains runIndex))
        |> List.map (fun (runIndex, runObj) ->
            fieldResultParts <- (match fieldParts.TryGetValue runIndex with | true, parts -> parts | _ -> [])
            let runXml =
                if not (isNullOrUndefined runObj?OpaqueXml) then
                    // An object kept from the file (chart, SmartArt, OLE object, ...):
                    // its element goes back as it was read. If a relationship it
                    // needs is not in this part (the run was pasted from another
                    // document), only its description is written.
                    let xml = runObj?OpaqueXml |> unbox<string>
                    let referenced = [ for m in relationshipIdPattern.Matches xml -> m.Groups.[1].Value ]
                    if referenced |> List.forall (fun id -> Set.contains id currentPartRelationshipIds) then
                        if boolOrDefault false (runObj?OpaqueParagraphLevel) then xml else "<w:r>" + xml + "</w:r>"
                    else
                        runElement [] (isDeletedRun runObj) (optionalString (runObj?ImageAlt) |> Option.defaultValue "")
                elif boolOrDefault false (runObj?IsImage) then
                    match optionalString (runObj?ImageSrc) |> Option.bind (fun src -> Map.tryFind src images.AssetBySrc) with
                    | Some asset -> generateImageRun asset runObj
                    | None ->
                        let alt = optionalString (runObj?ImageAlt) |> Option.defaultValue "Image"
                        runElement [] (isDeletedRun runObj) alt
                else
                    let run = editorRunOfObject runObj
                    generateReferenceRun runObj run
            let runXml = revisionWrappedRun runObj runXml
            match runHyperlinkTarget runObj with
            | Some target when target.StartsWith "#" && target.Length > 1 -> element "w:hyperlink" [("w:anchor", target.Substring 1); ("w:history", "1")] [runXml]
            | Some target ->
                match Map.tryFind target hyperlinks with
                | Some relId -> element "w:hyperlink" [("r:id", relId); ("w:history", "1")] [runXml]
                | None -> runXml
            | None -> runXml
        )
    let runs =
        // A next-page section break already starts a new page (the reader reports it
        // as PageBreakAfter too); adding w:br there would leave a blank page in Word.
        let sectionBreaksPage =
            match optionalString (para?SectionBreak) with
            | Some "nextPage" | Some "evenPage" | Some "oddPage" -> true
            | _ -> false
        if boolOrDefault false (para?PageBreakAfter) && not sectionBreaksPage then
            runs @ [element "w:r" [] [selfClosing "w:br" [("w:type", "page")]]]
        else runs

    // (A paragraph that states nothing of its own has no properties element.)
    // (With a start and an end tag also when it is empty: a shape anchored in it goes before its end tag.)
    match (if pPr = "<w:pPr/>" then runs else pPr :: runs) with
    | [] -> "<w:p></w:p>"
    | children -> element "w:p" [] children

let private generateParagraphWithHyperlinksAndImages (hyperlinks: Map<string, string>) (images: ImageScope) (para: obj) : string =
    generateParagraphWithHyperlinksAndImagesWithPageSetup defaultPageSetupInfo hyperlinks images para

let private generateParagraphWithHyperlinks (hyperlinks: Map<string, string>) (para: obj) : string =
    generateParagraphWithHyperlinksAndImages hyperlinks emptyImageScope para

let private defaultTableBorders () =
    element "w:tblBorders" [] [
        selfClosing "w:top" [("w:val", "single"); ("w:sz", "4"); ("w:color", "auto")]
        selfClosing "w:left" [("w:val", "single"); ("w:sz", "4"); ("w:color", "auto")]
        selfClosing "w:bottom" [("w:val", "single"); ("w:sz", "4"); ("w:color", "auto")]
        selfClosing "w:right" [("w:val", "single"); ("w:sz", "4"); ("w:color", "auto")]
        selfClosing "w:insideH" [("w:val", "single"); ("w:sz", "4"); ("w:color", "auto")]
        selfClosing "w:insideV" [("w:val", "single"); ("w:sz", "4"); ("w:color", "auto")]
    ]

let private borderElement name (value: obj) =
    if isNullOrUndefined value then None
    else
        let style = optionalString (value?Style) |> Option.defaultValue "single"
        let hidden = style = "none" || style = "nil"
        let color =
            optionalString (value?Color)
            |> InverseConversion.cssColorToWord
            |> Option.defaultValue "auto"
        // A nil/none border is written bare: sz/color on it would come back as
        // an explicit zero-width black border.
        let attrs =
            [ yield ("w:val", style)
              match optionalFloat (value?Width) with
              | Some amount -> yield ("w:sz", InverseConversion.pixelsToEighthPoints amount)
              | None -> if not hidden then yield ("w:sz", "4")
              if not hidden || (optionalString (value?Color)).IsSome then yield ("w:color", color) ]
        Some (selfClosing name attrs)

/// CT_TblBorders / CT_TcBorders order: top, left, bottom, right, insideH, insideV.
let private bordersElement elementName (value: obj) =
    if isNullOrUndefined value then None
    else
        let sides =
            [ borderElement "w:top" (value?Top)
              borderElement "w:left" (value?Left)
              borderElement "w:bottom" (value?Bottom)
              borderElement "w:right" (value?Right)
              borderElement "w:insideH" (value?InsideH)
              borderElement "w:insideV" (value?InsideV) ]
            |> List.choose id
        if sides.IsEmpty then None else Some (element elementName [] sides)

let private tableWidthValue widthType width =
    match widthType with
    | "pct" -> roundToInt (width * 50.0) |> string
    // Editor-made widths ("pt" kind) are px; DOCX has no such unit, so they go out as dxa.
    | "pt" -> InverseConversion.pixelsToTwips width
    | "auto" | "nil" -> "0"
    | _ -> roundToInt width |> string

let private optionalTableWidth elementName value typeValue =
    match optionalFloat value, optionalString typeValue with
    | Some width, Some widthType -> Some (selfClosing elementName [("w:w", tableWidthValue widthType width); ("w:type", (if widthType = "pt" then "dxa" else widthType))])
    | Some width, None -> Some (selfClosing elementName [("w:w", string (roundToInt width)); ("w:type", "dxa")])
    | None, Some widthType when widthType = "auto" || widthType = "nil" -> Some (selfClosing elementName [("w:w", "0"); ("w:type", widthType)])
    | _ -> None

let private wordColorOrAuto value =
    optionalString value
    |> Option.map (fun color ->
        if color = "auto" then "auto"
        else InverseConversion.cssColorToWord (Some color) |> Option.defaultValue color)
    |> Option.defaultValue "auto"

let private shadingElement (value: obj) =
    if isNullOrUndefined value then None
    else
        let pattern = optionalString (value?Pattern) |> Option.defaultValue "clear"
        let foreground = wordColorOrAuto (value?ForegroundColor)
        let background = wordColorOrAuto (value?BackgroundColor)
        Some (selfClosing "w:shd" [("w:val", pattern); ("w:color", foreground); ("w:fill", background)])

let private backgroundShadingElement (value: obj) =
    optionalString value
    |> Option.map (fun color ->
        let wordColor = InverseConversion.cssColorToWord (Some color) |> Option.defaultValue color
        selfClosing "w:shd" [("w:val", "clear"); ("w:color", "auto"); ("w:fill", wordColor)])

/// Cell margins (tcMar / tblCellMar) from PaddingTop/Left/Bottom/Right, which
/// the editor package carries in dxa. CT_TcMar order: top, left, bottom, right.
let private cellMarginsElement elementName (owner: obj) =
    let margin name (value: obj) =
        match optionalFloat value with
        | Some amount when amount >= 0.0 -> Some (selfClosing name [("w:w", string (roundToInt amount)); ("w:type", "dxa")])
        | _ -> None
    let children =
        [ margin "w:top" owner?PaddingTop
          margin "w:left" owner?PaddingLeft
          margin "w:bottom" owner?PaddingBottom
          margin "w:right" owner?PaddingRight ]
        |> List.choose id
    if children.IsEmpty then None else Some (element elementName [] children)

let private cellSpacingElement (owner: obj) =
    match optionalFloat (owner?CellSpacing) with
    | Some amount when amount > 0.0 -> Some (selfClosing "w:tblCellSpacing" [("w:w", string (roundToInt amount)); ("w:type", "dxa")])
    | _ -> None

let private generateTableGrid (table: obj) =
    let columns =
        optionalArray (table?GridColumns)
        |> Array.toList
        |> List.choose (fun value -> optionalFloat value)
        |> List.map (fun width -> selfClosing "w:gridCol" [("w:w", string (roundToInt width))])
    if columns.IsEmpty then "" else element "w:tblGrid" [] columns

let rec private generateTableCell (hyperlinks: Map<string, string>) (images: ImageScope) (cell: obj) : string =
    let paragraphs =
        optionalArray (cell?Paragraphs)
        |> Array.map (generateParagraphWithHyperlinksAndImages hyperlinks images)
    let nestedTables = optionalArray (cell?NestedTables)
    // NestedTableAnchors.[i] = how many of the cell's paragraphs precede nested
    // table i. The writer used to read only Paragraphs, so a nested table and
    // everything in it was deleted on save.
    let anchors = optionalArray (cell?NestedTableAnchors) |> Array.map (fun value -> jsNumber value |> roundToInt)
    let content = ResizeArray<string>()
    let mutable lastIsTable = false
    let emitTablesAt index =
        nestedTables
        |> Array.iteri (fun tableIndex nested ->
            let anchor = if tableIndex < anchors.Length then anchors.[tableIndex] else paragraphs.Length
            if (min anchor paragraphs.Length) = index then
                content.Add (generateTable hyperlinks images nested)
                lastIsTable <- true)
    for index in 0 .. paragraphs.Length - 1 do
        emitTablesAt index
        content.Add paragraphs.[index]
        lastIsTable <- false
    emitTablesAt paragraphs.Length
    // A cell must contain, and end with, a paragraph.
    if content.Count = 0 || lastIsTable then content.Add (element "w:p" [] [])

    // CT_TcPr order: tcW, gridSpan, vMerge, tcBorders, shd, tcMar, vAlign.
    let tcPrChildren = [
        match optionalTableWidth "w:tcW" (cell?PreferredWidth) (cell?PreferredWidthType) with
        | Some width -> yield width
        | None ->
            match optionalFloat (cell?Width) with
            | Some width -> yield selfClosing "w:tcW" [("w:w", string (roundToInt width)); ("w:type", "dxa")]
            | None -> ()

        match optionalFloat (cell?ColSpan) with
        | Some span when span > 1.0 -> yield selfClosing "w:gridSpan" [("w:val", string (roundToInt span))]
        | _ -> ()

        match optionalString (cell?VerticalMerge) with
        | Some "restart" -> yield selfClosing "w:vMerge" [("w:val", "restart")]
        | Some "continue" -> yield selfClosing "w:vMerge" []
        | _ -> ()

        match bordersElement "w:tcBorders" (cell?Borders) with
        | Some borders -> yield borders
        | None ->
            // Legacy single-border fields from editor-made cells.
            if isNullOrUndefined cell?Borders && ((optionalString (cell?BorderColor)).IsSome || (optionalFloat (cell?BorderWidth)).IsSome) then
                let color = optionalString (cell?BorderColor) |> InverseConversion.cssColorToWord |> Option.defaultValue "auto"
                let width = optionalFloat (cell?BorderWidth) |> Option.map InverseConversion.pixelsToEighthPoints |> Option.defaultValue "8"
                let border name = selfClosing name [("w:val", "single"); ("w:sz", width); ("w:color", color)]
                yield element "w:tcBorders" [] [ border "w:top"; border "w:left"; border "w:bottom"; border "w:right" ]

        match shadingElement (cell?Shading) |> Option.orElseWith (fun () -> backgroundShadingElement (cell?BackgroundColor)) with
        | Some shading -> yield shading
        | None -> ()

        match cellMarginsElement "w:tcMar" cell with
        | Some margins -> yield margins
        | None -> ()

        match optionalString (cell?VerticalAlignment) with
        | Some value when value = "top" || value = "center" || value = "bottom" || value = "both" ->
            yield selfClosing "w:vAlign" [("w:val", value)]
        | _ -> ()
    ]

    element "w:tc" [] ((if tcPrChildren.IsEmpty then [] else [element "w:tcPr" [] tcPrChildren]) @ List.ofSeq content)

and private generateTableRow (hyperlinks: Map<string, string>) (images: ImageScope) (row: obj) : string =
    let cells =
        optionalArray (row?Cells)
        |> Array.map (generateTableCell hyperlinks images)
        |> String.concat ""

    // Row-level cell margins live in w:tblPrEx, which precedes w:trPr.
    let rowExceptions =
        match cellMarginsElement "w:tblCellMar" row with
        | Some margins -> element "w:tblPrEx" [] [margins]
        | None -> ""
    let rowPropChildren = [
        if boolOrDefault true (row?CanSplit) = false then yield selfClosing "w:cantSplit" []
        match optionalFloat (row?Height) with
        | Some height ->
            let attrs =
                [("w:val", string (roundToInt height))]
                @ (optionalString (row?HeightRule) |> Option.map (fun rule -> [("w:hRule", rule)]) |> Option.defaultValue [])
            yield selfClosing "w:trHeight" attrs
        | None -> ()
        if boolOrDefault false (row?IsHeader) then yield selfClosing "w:tblHeader" []
        match cellSpacingElement row with
        | Some spacing -> yield spacing
        | None -> ()
    ]
    let rowProps = if rowPropChildren.IsEmpty then "" else element "w:trPr" [] rowPropChildren

    element "w:tr" [] [rowExceptions; rowProps; cells]

and private generateTable (hyperlinks: Map<string, string>) (images: ImageScope) (table: obj) : string =
    let rows =
        optionalArray (table?Rows)
        |> Array.map (generateTableRow hyperlinks images)
        |> String.concat ""

    // CT_TblPr order: tblStyle, tblW, jc, tblCellSpacing, tblBorders, shd, tblLayout, tblCellMar, tblLook.
    let tblPrChildren = [
        // The table keeps its style: the reader resolves the style's formatting
        // into the cells from the reference (and the style itself is kept in
        // styles.xml), so dropping it changed how the saved table read back.
        match optionalString (table?TableStyleId) with
        | Some styleId -> yield selfClosing "w:tblStyle" [("w:val", styleId)]
        | None -> ()
        // A floating table stays one: without its w:tblpPr the table came back in
        // the text flow and the paragraphs that ran beside it below it.
        match optionalString (table?FloatingXml) with
        | Some floatingXml when floatingXml.StartsWith "<w:tblpPr" -> yield floatingXml
        | _ -> if boolOrDefault false (table?Floating) then yield selfClosing "w:tblpPr" [("w:leftFromText", "180"); ("w:rightFromText", "180"); ("w:vertAnchor", "text"); ("w:tblpY", "1")]
        match optionalTableWidth "w:tblW" (table?PreferredWidth) (table?PreferredWidthType) with
        | Some width -> yield width
        | None -> ()
        match optionalString (table?TableAlignment) with
        | Some value when value <> "" -> yield selfClosing "w:jc" [("w:val", value)]
        | _ -> ()
        match cellSpacingElement table with
        | Some spacing -> yield spacing
        | None -> ()
        match bordersElement "w:tblBorders" (table?Borders) with
        | Some borders -> yield borders
        | None ->
            // Only a table made in the editor (no grid, no width: it never came
            // from a file) gets the default grid. An imported table without
            // borders is borderless and must stay that way.
            if (optionalArray (table?GridColumns)).Length = 0 && isNullOrUndefined table?PreferredWidthType then
                yield defaultTableBorders ()
        match shadingElement (table?Shading) |> Option.orElseWith (fun () -> backgroundShadingElement (table?BackgroundColor)) with
        | Some shading -> yield shading
        | None -> ()
        match optionalString (table?LayoutMode) with
        | Some mode when mode = "fixed" || mode = "autofit" || mode = "autoFit" ->
            yield selfClosing "w:tblLayout" [("w:type", (if mode = "fixed" then "fixed" else "autofit"))]
        | _ -> ()
        match cellMarginsElement "w:tblCellMar" table with
        | Some margins -> yield margins
        | None -> ()
        let look = table?TableLook
        if not (isNullOrUndefined look) then
            let flag (value: obj) = boolOrDefault false value
            let bits =
                (if flag (look?FirstRow) then 0x0020 else 0) ||| (if flag (look?LastRow) then 0x0040 else 0)
                ||| (if flag (look?FirstColumn) then 0x0080 else 0) ||| (if flag (look?LastColumn) then 0x0100 else 0)
                ||| (if flag (look?HorizontalBanding) then 0 else 0x0200) ||| (if flag (look?VerticalBanding) then 0 else 0x0400)
            let hexDigits = "0123456789ABCDEF"
            let hex = System.String([| for shift in [ 12; 8; 4; 0 ] -> hexDigits.[(bits >>> shift) &&& 0xF] |])
            let onOffAttr (isOn: bool) = if isOn then "1" else "0"
            yield selfClosing "w:tblLook" [
                ("w:val", hex)
                ("w:firstRow", onOffAttr (flag (look?FirstRow))); ("w:lastRow", onOffAttr (flag (look?LastRow)))
                ("w:firstColumn", onOffAttr (flag (look?FirstColumn))); ("w:lastColumn", onOffAttr (flag (look?LastColumn)))
                ("w:noHBand", onOffAttr (not (flag (look?HorizontalBanding)))); ("w:noVBand", onOffAttr (not (flag (look?VerticalBanding))))
            ]
    ]

    element "w:tbl" [] [
        element "w:tblPr" [] tblPrChildren
        generateTableGrid table
        rows
    ]

/// Shape ids must be unique in the document; they are handed out per save.
let mutable private nextShapeId = 1000
/// Every shape is written from the model, none as the element it was read
/// from (a package that asks for everything from the model: RegenerateAll).
let mutable private shapesFromModel = false

let private drawingNamespaces =
    {| Wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
       A = "http://schemas.openxmlformats.org/drawingml/2006/main"
       Wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
       Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006" |}

let private srgbColor (color: string) =
    selfClosing "a:srgbClr" [("val", InverseConversion.cssColorToWord (Some color) |> Option.defaultValue "000000")]

/// a:prstGeom or a:custGeom of a shape item (Geometry / CustomPath / CustomW / CustomH).
let private shapeGeometryElement (item: obj) =
    let geometry = optionalString (item?Geometry) |> Option.defaultValue "rect"
    let coordinate (value: float) = string (roundToInt value)
    let point (points: obj) index = selfClosing "a:pt" [("x", coordinate (floatOrDefault 0.0 (points?(index)))); ("y", coordinate (floatOrDefault 0.0 (points?(index + 1))))]
    let segments =
        optionalArray (item?CustomPath)
        |> Array.toList
        |> List.choose (fun segment ->
            let points = segment?P
            match optionalString (segment?Op) with
            | Some "M" -> Some (element "a:moveTo" [] [point points 0])
            | Some "L" -> Some (element "a:lnTo" [] [point points 0])
            | Some "C" -> Some (element "a:cubicBezTo" [] [point points 0; point points 2; point points 4])
            | Some "Q" -> Some (element "a:quadBezTo" [] [point points 0; point points 2])
            | Some "Z" -> Some (selfClosing "a:close" [])
            | _ -> None)
    let coordW = floatOrDefault 0.0 (item?CustomW)
    let coordH = floatOrDefault 0.0 (item?CustomH)
    if geometry = "custom" && not segments.IsEmpty && coordW > 0.0 && coordH > 0.0 then
        element "a:custGeom" [] [
            selfClosing "a:avLst" []
            selfClosing "a:gdLst" []
            selfClosing "a:ahLst" []
            selfClosing "a:cxnLst" []
            selfClosing "a:rect" [("l", "0"); ("t", "0"); ("r", "r"); ("b", "b")]
            element "a:pathLst" [] [element "a:path" [("w", coordinate coordW); ("h", coordinate coordH)] segments]
        ]
    else
        element "a:prstGeom" [("prst", (if geometry = "custom" || geometry = "" then "rect" else geometry))] [selfClosing "a:avLst" []]

/// The fill of a shape: gradient, solid colour or none.
let private shapeFillElement (item: obj) =
    let gradient = item?Gradient
    let stops = if isNullOrUndefined gradient then [||] else optionalArray (gradient?Stops)
    if stops.Length > 0 then
        let stopElements =
            stops
            |> Array.toList
            |> List.map (fun stop ->
                element "a:gs" [("pos", string (roundToInt (floatOrDefault 0.0 (stop?Offset) * 100000.0)))] [srgbColor (stop?Color |> stringOrDefault "#000000")])
        let direction =
            if boolOrDefault false (gradient?Radial) then
                element "a:path" [("path", "circle")] [selfClosing "a:fillToRect" [("l", "50000"); ("t", "50000"); ("r", "50000"); ("b", "50000")]]
            else
                selfClosing "a:lin" [("ang", string (roundToInt (floatOrDefault 0.0 (gradient?Angle) * 60000.0))); ("scaled", "0")]
        element "a:gradFill" [("rotWithShape", "1")] [element "a:gsLst" [] stopElements; direction]
    elif boolOrDefault false (item?NoFill) then selfClosing "a:noFill" []
    else element "a:solidFill" [] [srgbColor (optionalString (item?FillColor) |> Option.defaultValue "#FFFFFF")]

/// a:ln: width (px -> EMU), colour or none, preset dash.
let private shapeLineElement (item: obj) =
    let width = optionalFloat (item?BorderWidth) |> Option.defaultValue 1.0 |> pixelsToEmus
    let children =
        [ if boolOrDefault false (item?NoLine) then yield selfClosing "a:noFill" []
          else yield element "a:solidFill" [] [srgbColor (optionalString (item?BorderColor) |> Option.defaultValue "#808080")]
          match optionalString (item?BorderDash) with
          | Some dash when dash <> "" && dash <> "solid" -> yield selfClosing "a:prstDash" [("val", dash)]
          | _ -> () ]
    element "a:ln" [("w", string (max 0 width))] children

/// The VML form of a shape, for readers without DrawingML shapes (the
/// mc:Fallback Word writes too): a text box rectangle at the same place.
let private generateVmlShape (shapeId: string) (item: obj) (content: string) : string =
    let width = floatOrDefault 220.0 (item?Width) |> pixelsToPoints
    let height = floatOrDefault 72.0 (item?Height) |> pixelsToPoints
    let x = optionalFloat (item?X) |> Option.map pixelsToPoints |> Option.defaultValue 0.0
    let y = optionalFloat (item?Y) |> Option.map pixelsToPoints |> Option.defaultValue 0.0
    let strokeWeight = optionalFloat (item?BorderWidth) |> Option.map pixelsToPoints |> Option.defaultValue 0.75
    let rotation = floatOrDefault 0.0 (item?Rotation)
    let style =
        if boolOrDefault false (item?Inline) then sprintf "width:%.4fpt;height:%.4fpt" width height
        else sprintf "position:absolute;margin-left:%.4fpt;margin-top:%.4fpt;width:%.4fpt;height:%.4fpt;z-index:1" x y width height
    let style = if abs rotation >= 0.01 then style + sprintf ";rotation:%.2f" rotation else style
    element "w:pict" [] [
        element "v:shape" [
            yield ("xmlns:v", "urn:schemas-microsoft-com:vml")
            yield ("id", shapeId)
            yield ("type", "#_x0000_t202")
            yield ("style", style)
            if boolOrDefault false (item?NoFill) then yield ("filled", "f")
            else yield ("fillcolor", optionalString (item?FillColor) |> Option.defaultValue "#FFFFFF")
            if boolOrDefault false (item?NoLine) then yield ("stroked", "f")
            else
                yield ("strokecolor", optionalString (item?BorderColor) |> Option.defaultValue "#808080")
                yield ("strokeweight", sprintf "%.2fpt" strokeWeight)
        ] [
            element "v:textbox" [
                // (The same insets and fit as the DrawingML choice says.)
                let inset (value: obj) (fallback: float) = optionalFloat value |> Option.map pixelsToPoints |> Option.defaultValue fallback
                if not (isNullOrUndefined item?InsetLeft && isNullOrUndefined item?InsetTop && isNullOrUndefined item?InsetRight && isNullOrUndefined item?InsetBottom) then
                    yield ("inset", sprintf "%.2fpt,%.2fpt,%.2fpt,%.2fpt" (inset (item?InsetLeft) 7.2) (inset (item?InsetTop) 3.6) (inset (item?InsetRight) 7.2) (inset (item?InsetBottom) 3.6))
                if boolOrDefault false (item?AutoFit) then yield ("style", "mso-fit-shape-to-text:t")
            ] [
                element "w:txbxContent" [] [content]
            ]
        ]
    ]

/// A text box / shape as Word writes it: DrawingML (wp:anchor or wp:inline
/// around a wps:wsp) in an mc:AlternateContent with a VML fallback. Geometry,
/// rotation, solid / gradient / no fill, line colour, width and dash, wrap
/// mode and position are written from the model. (It used to be a bare VML
/// text box rectangle, so an ellipse or a rotated arrow came back as a white box.)
let rec private generateTextBoxRun (packageSetup: PageSetupInfo) (hyperlinks: Map<string, string>) (images: ImageScope) (item: obj) : string =
    // A shape that is what was read goes back as the element it was read from
    // (its DrawingML and its VML fallback, with what the model does not hold
    // of them) - when everything that element refers to is in this part.
    let source =
        (if shapesFromModel then None else optionalString (item?SourceXml))
        |> Option.filter (fun xml ->
            [ for m in relationshipIdPattern.Matches xml -> m.Groups.[1].Value ]
            |> List.forall (fun id -> Set.contains id currentPartRelationshipIds))
    match source with
    | Some xml -> "<w:r>" + xml + "</w:r>"
    | None when boolOrDefault false (item?FitText) -> generateWordArtRun item
    | None -> generateTextBoxRunFromModel packageSetup hyperlinks images item

/// Text drawn as large as its shape (WordArt, a watermark) from the model: a
/// VML shape with a v:textpath, as Word writes a watermark - where it stands
/// (aligned or by its margins, relative to the margins, the page or the
/// text), turned, behind the text, the colour and opacity of its letters,
/// their face. (Written as a text box it came back as small text in a frame.)
and private generateWordArtRun (item: obj) : string =
    let shapeNumber = nextShapeId
    nextShapeId <- nextShapeId + 1
    let paragraphs = if isNullOrUndefined item?Paragraphs then [||] else item?Paragraphs |> unbox<obj array>
    let runs = paragraphs |> Array.collect (fun paragraph -> optionalArray (paragraph?Runs))
    let text = runs |> Array.map (fun run -> optionalString (run?Text) |> Option.defaultValue "") |> String.concat ""
    let first = runs |> Array.tryHead
    let face = first |> Option.bind (fun run -> optionalString (run?FontFamily)) |> Option.defaultValue "Calibri"
    let color = first |> Option.bind (fun run -> optionalString (run?Color)) |> Option.defaultValue "#C0C0C0"
    let bold = first |> Option.exists (fun run -> boolOrDefault false (run?Bold))
    let italic = first |> Option.exists (fun run -> boolOrDefault false (run?Italic))
    let points (value: obj) (fallback: float) = optionalFloat value |> Option.map pixelsToPoints |> Option.defaultValue fallback
    let rotation = floatOrDefault 0.0 (item?Rotation)
    let relative (value: obj) (ofText: string) =
        match optionalString value with
        | Some "margin" -> Some "margin"
        | Some "page" -> Some "page"
        | Some "column" | Some "paragraph" -> Some ofText
        | Some "character" -> Some "char"
        | Some "line" -> Some "line"
        | _ -> None
    let style =
        [ if not (boolOrDefault false (item?Inline)) then yield "position:absolute"
          yield sprintf "margin-left:%.2fpt" (points (item?X) 0.0)
          yield sprintf "margin-top:%.2fpt" (points (item?Y) 0.0)
          yield sprintf "width:%.2fpt" (points (item?Width) 400.0)
          yield sprintf "height:%.2fpt" (points (item?Height) 100.0)
          if abs rotation >= 0.01 then yield sprintf "rotation:%.2f" rotation
          yield (if boolOrDefault false (item?BehindDoc) then "z-index:-251658752" else "z-index:251658752")
          match optionalString (item?PositionHAlign) with
          | Some align -> yield "mso-position-horizontal:" + align
          | None -> ()
          match relative (item?PositionHRelativeFrom) "text" with
          | Some name -> yield "mso-position-horizontal-relative:" + name
          | None -> ()
          match optionalString (item?PositionVAlign) with
          | Some align -> yield "mso-position-vertical:" + align
          | None -> ()
          match relative (item?PositionVRelativeFrom) "text" with
          | Some name -> yield "mso-position-vertical-relative:" + name
          | None -> () ]
        |> String.concat ";"
    let pathStyle =
        sprintf "font-family:\"%s\";font-size:1pt" face
        + (if bold then ";font-weight:bold" else "")
        + (if italic then ";font-style:italic" else "")
    let shapeType =
        "<v:shapetype id=\"_x0000_t136\" coordsize=\"21600,21600\" o:spt=\"136\" adj=\"10800\" path=\"m@7,l@8,m@5,21600l@6,21600e\">"
        + "<v:formulas><v:f eqn=\"sum #0 0 10800\"/><v:f eqn=\"prod #0 2 1\"/><v:f eqn=\"sum 21600 0 @1\"/><v:f eqn=\"sum 0 0 @2\"/><v:f eqn=\"sum 21600 0 @3\"/><v:f eqn=\"if @0 @3 0\"/><v:f eqn=\"if @0 21600 @1\"/>"
        + "<v:f eqn=\"if @0 0 @2\"/><v:f eqn=\"if @0 @4 21600\"/><v:f eqn=\"mid @5 @6\"/><v:f eqn=\"mid @8 @5\"/><v:f eqn=\"mid @7 @8\"/><v:f eqn=\"mid @6 @7\"/><v:f eqn=\"sum @6 0 @5\"/></v:formulas>"
        + "<v:path textpathok=\"t\" o:connecttype=\"custom\" o:connectlocs=\"@9,0;@10,10800;@11,21600;@12,10800\" o:connectangles=\"270,180,90,0\"/><v:textpath on=\"t\" fitshape=\"t\"/>"
        + "<v:handles><v:h position=\"#0,bottomRight\" xrange=\"6629,14971\"/></v:handles><o:lock v:ext=\"edit\" text=\"t\" shapetype=\"t\"/></v:shapetype>"
    let opacity = optionalFloat (item?Opacity) |> Option.filter (fun value -> value >= 0.0 && value < 0.999)
    element "w:r" [] [
        element "w:pict" [("xmlns:v", "urn:schemas-microsoft-com:vml"); ("xmlns:o", "urn:schemas-microsoft-com:office:office")] [
            shapeType
            element "v:shape" [
                ("id", optionalString (item?Name) |> Option.defaultValue (sprintf "WordArt%d" shapeNumber)); ("o:spid", sprintf "_x0000_s%d" shapeNumber); ("type", "#_x0000_t136")
                ("style", style); ("o:allowincell", "f"); ("fillcolor", color); ("stroked", "f")
            ] [
                match opacity with
                | Some value -> yield selfClosing "v:fill" [("opacity", sprintf "%.3f" value)]
                | None -> ()
                yield selfClosing "v:textpath" [("style", pathStyle); ("string", text)]
            ]
        ]
    ]

and private generateTextBoxRunFromModel (packageSetup: PageSetupInfo) (hyperlinks: Map<string, string>) (images: ImageScope) (item: obj) : string =
    let widthEmu = floatOrDefault 220.0 (item?Width) |> pixelsToEmus |> max 0 |> string
    let heightEmu = floatOrDefault 72.0 (item?Height) |> pixelsToEmus |> max 0 |> string
    let shapeNumber = nextShapeId
    nextShapeId <- nextShapeId + 1
    let shapeName = optionalString (item?Name) |> Option.defaultValue (sprintf "Shape %d" shapeNumber)
    let paragraphs =
        if isNullOrUndefined item?Paragraphs then [||] else item?Paragraphs |> unbox<obj array>
    let content =
        if paragraphs.Length = 0 then ""
        else paragraphs |> Array.map (generateParagraphWithHyperlinksAndImagesWithPageSetup packageSetup hyperlinks images) |> String.concat ""
    let emptyParagraph = element "w:p" [] []
    let rotation = floatOrDefault 0.0 (item?Rotation)
    let shape =
        element "wps:wsp" [] [
            yield selfClosing "wps:cNvSpPr" (if paragraphs.Length > 0 then [("txBox", "1")] else [])
            yield element "wps:spPr" [] [
                element "a:xfrm" (if abs rotation >= 0.0001 then [("rot", string (roundToInt (rotation * 60000.0)))] else []) [
                    selfClosing "a:off" [("x", "0"); ("y", "0")]
                    selfClosing "a:ext" [("cx", widthEmu); ("cy", heightEmu)]
                ]
                shapeGeometryElement item
                shapeFillElement item
                shapeLineElement item
            ]
            if paragraphs.Length > 0 then
                yield element "wps:txbx" [] [element "w:txbxContent" [] [content]]
            // The room around the text, where it stands in the shape's height, whether the shape follows it.
            let insetEmu (value: obj) (fallback: int) = optionalFloat value |> Option.map (fun px -> pixelsToEmus px |> max 0) |> Option.defaultValue fallback |> string
            let bodyAttributes =
                [ ("rot", "0"); ("vert", "horz"); ("wrap", (if boolOrDefault false (item?NoTextWrap) then "none" else "square"))
                  ("lIns", insetEmu (item?InsetLeft) 91440); ("tIns", insetEmu (item?InsetTop) 45720); ("rIns", insetEmu (item?InsetRight) 91440); ("bIns", insetEmu (item?InsetBottom) 45720)
                  ("anchor", optionalString (item?VerticalAnchor) |> Option.defaultValue "t"); ("anchorCtr", "0") ]
            if boolOrDefault false (item?AutoFit) then yield element "wps:bodyPr" bodyAttributes [ selfClosing "a:spAutoFit" [] ]
            else yield selfClosing "wps:bodyPr" bodyAttributes
        ]
    let graphic =
        element "a:graphic" [("xmlns:a", drawingNamespaces.A)] [
            element "a:graphicData" [("uri", drawingNamespaces.Wps)] [shape]
        ]
    let docPr = selfClosing "wp:docPr" [("id", string shapeNumber); ("name", shapeName)]
    let distance (value: obj) = optionalFloat value |> Option.defaultValue 0.0 |> pixelsToEmus |> max 0 |> string
    let drawing =
        if boolOrDefault false (item?Inline) then
            element "wp:inline" [("xmlns:wp", drawingNamespaces.Wp); ("distT", distance (item?DistanceTop)); ("distB", distance (item?DistanceBottom)); ("distL", distance (item?DistanceLeft)); ("distR", distance (item?DistanceRight))] [
                selfClosing "wp:extent" [("cx", widthEmu); ("cy", heightEmu)]
                docPr
                selfClosing "wp:cNvGraphicFramePr" []
                graphic
            ]
        else
            let offset (value: obj) = optionalFloat value |> Option.defaultValue 0.0 |> pixelsToEmus |> string
            let relativeFrom (value: obj) (legacy: obj) fallback =
                optionalString value |> Option.orElse (optionalString legacy) |> Option.defaultValue fallback
            element "wp:anchor" [
                ("xmlns:wp", drawingNamespaces.Wp)
                ("distT", distance (item?DistanceTop)); ("distB", distance (item?DistanceBottom)); ("distL", distance (item?DistanceLeft)); ("distR", distance (item?DistanceRight))
                ("simplePos", "0")
                ("relativeHeight", optionalString (item?RelativeHeight) |> Option.orElse (optionalString (item?TextBoxRelativeHeight)) |> Option.defaultValue "251658240")
                ("behindDoc", boolFlag false (item?BehindDoc)); ("locked", "0")
                ("layoutInCell", boolFlag true (item?LayoutInCell)); ("allowOverlap", boolFlag true (item?AllowOverlap))
            ] [
                selfClosing "wp:simplePos" [("x", "0"); ("y", "0")]
                anchorPositionElement "wp:positionH" (relativeFrom (item?PositionHRelativeFrom) (item?TextBoxPositionHRelativeFrom) "column") (optionalString (item?PositionHAlign)) (offset (item?X))
                anchorPositionElement "wp:positionV" (relativeFrom (item?PositionVRelativeFrom) (item?TextBoxPositionVRelativeFrom) "paragraph") (optionalString (item?PositionVAlign)) (offset (item?Y))
                selfClosing "wp:extent" [("cx", widthEmu); ("cy", heightEmu)]
                selfClosing "wp:effectExtent" [("l", "0"); ("t", "0"); ("r", "0"); ("b", "0")]
                // A text box made in the editor has no wrap mode; it floats over the text.
                anchorWrapElement (optionalString (item?Wrap) |> Option.orElse (Some "none")) None [||]
                docPr
                selfClosing "wp:cNvGraphicFramePr" []
                graphic
            ]
    element "w:r" [] [
        element "mc:AlternateContent" [("xmlns:mc", drawingNamespaces.Mc); ("xmlns:wps", drawingNamespaces.Wps)] [
            element "mc:Choice" [("Requires", "wps")] [element "w:drawing" [] [drawing]]
            element "mc:Fallback" [] [generateVmlShape (sprintf "_x0000_s%d" shapeNumber) item (if content = "" then emptyParagraph else content)]
        ]
    ]

let private generateTextBox (packageSetup: PageSetupInfo) (hyperlinks: Map<string, string>) (images: ImageScope) (item: obj) : string =
    element "w:p" [] [generateTextBoxRun packageSetup hyperlinks images item]

/// Blocks to XML. A text box is absolutely positioned and the reader lists
/// shapes separately from the paragraph that anchors them, so giving each one
/// a paragraph of its own added an empty paragraph per shape on every save.
/// It is anchored in the preceding paragraph instead when there is one.
let private generateBlocks (generateItem: obj -> string) (generateTextBoxRun: obj -> string) (items: obj array) : string =
    let parts = ResizeArray<string>()
    let paragraphEnd = "</w:p>"
    for item in items do
        let isTextBox = (item?Type |> stringOrDefault "Paragraph") = "TextBox"
        if isTextBox && parts.Count > 0 && parts.[parts.Count - 1].EndsWith paragraphEnd then
            let previous = parts.[parts.Count - 1]
            parts.[parts.Count - 1] <- previous.Substring(0, previous.Length - paragraphEnd.Length) + generateTextBoxRun item + paragraphEnd
        elif isTextBox && parts.Count > 0 && parts.[parts.Count - 1].EndsWith "</w:tbl>" && parts.[parts.Count - 1].LastIndexOf paragraphEnd >= 0 then
            // Anchored inside a table: stays in the table, in its last paragraph.
            let previous = parts.[parts.Count - 1]
            let index = previous.LastIndexOf paragraphEnd
            parts.[parts.Count - 1] <- previous.Substring(0, index) + generateTextBoxRun item + previous.Substring index
        else
            parts.Add (generateItem item)
    String.concat "" parts

let private wrapContentControl (item: obj) (content: string) =
    match optionalString (item?ContentControlId), optionalString (item?ContentControlTag), optionalString (item?ContentControlAlias) with
    | None, None, None -> content
    | id, tag, alias ->
        let props = [
            match id with
            | Some value when value <> "" -> yield selfClosing "w:id" [("w:val", value)]
            | _ -> ()
            match tag with
            | Some value when value <> "" -> yield selfClosing "w:tag" [("w:val", value)]
            | _ -> ()
            match alias with
            | Some value when value <> "" -> yield selfClosing "w:alias" [("w:val", value)]
            | _ -> ()
        ]
        element "w:sdt" [] [element "w:sdtPr" [] props; element "w:sdtContent" [] [content]]

let private generateDocumentItemWithImages (hyperlinks: Map<string, string>) (images: ImageScope) (item: obj) : string =
    let content =
        match item?Type |> stringOrDefault "Paragraph" with
        | "Table" -> generateTable hyperlinks images item
        | "TextBox" -> generateTextBox defaultPageSetupInfo hyperlinks images item
        | _ -> generateParagraphWithHyperlinksAndImagesWithPageSetup defaultPageSetupInfo hyperlinks images item
    wrapContentControl item content

let private generateDocumentItemWithImagesAndPageSetup (packageSetup: PageSetupInfo) (hyperlinks: Map<string, string>) (images: ImageScope) (item: obj) : string =
    let content =
        match item?Type |> stringOrDefault "Paragraph" with
        | "Table" -> generateTable hyperlinks images item
        | "TextBox" -> generateTextBox packageSetup hyperlinks images item
        | _ -> generateParagraphWithHyperlinksAndImagesWithPageSetup packageSetup hyperlinks images item
    wrapContentControl item content

let private generateDocumentItem (hyperlinks: Map<string, string>) (item: obj) : string =
    generateDocumentItemWithImages hyperlinks emptyImageScope item

let generateParagraph (para: obj) : string =
    generateParagraphWithHyperlinks Map.empty para

type HeaderFooterPartInfo = {
    Kind: string; RelId: string; Target: string; PartName: string; Blocks: obj array; IsGlobal: bool
    /// The part says it is what it was when the document was read.
    Unchanged: bool }
type HeaderFooterPartPackage = { Info: HeaderFooterPartInfo; Images: ImageScope; Hyperlinks: Map<string, string> }

let private partBlocks (part: obj) =
    if not (isNullOrUndefined part?Blocks) then optionalArray (part?Blocks)
    elif not (isNullOrUndefined part?Paragraphs) then optionalArray (part?Paragraphs)
    else [||]

/// `originalPaths`: the package path a part was read from, by its relationship
/// id. A part that was directly in word/ keeps its name (the headers of a
/// document are not renamed by a save); the others get the first free one.
let private headerFooterInfos (originalPaths: Map<string, string>) scopedRelationshipIds prefix startIndex (parts: obj array) =
    let originalName (part: obj) =
        optionalString (part?RelationshipId)
        |> Option.bind (fun relId -> Map.tryFind relId originalPaths)
        |> Option.bind (fun path ->
            if path.StartsWith "word/" && path.IndexOf('/', 5) < 0 && path.EndsWith ".xml" then Some (path.Substring 5) else None)
    let taken = System.Collections.Generic.HashSet<string>()
    for KeyValue (_, path) in originalPaths do
        if path.StartsWith "word/" then taken.Add (path.Substring(5).ToLowerInvariant()) |> ignore
    let given = System.Collections.Generic.HashSet<string>()
    let names =
        parts
        |> Array.map (fun part ->
            match originalName part with
            | Some name when given.Add (name.ToLowerInvariant()) -> Some name
            | _ -> None)
    let mutable nextNumber = startIndex
    let freshName () =
        let mutable name = sprintf "%s%d.xml" prefix nextNumber
        while taken.Contains name || given.Contains name do
            nextNumber <- nextNumber + 1
            name <- sprintf "%s%d.xml" prefix nextNumber
        nextNumber <- nextNumber + 1
        given.Add name |> ignore
        name
    parts
    |> Array.mapi (fun index part ->
        let number = startIndex + index
        let explicitRelId = optionalString (part?RelationshipId)
        let relId = explicitRelId |> Option.defaultValue (freshRelationshipId (sprintf "rId%s%d" prefix number))
        let fileName = match names.[index] with Some name -> name | None -> freshName ()
        { Kind = part?Kind |> stringOrDefault (part?Type |> stringOrDefault "default")
          RelId = relId
          Target = fileName
          PartName = sprintf "/word/%s" fileName
          Blocks = partBlocks part
          IsGlobal = not (Set.contains relId scopedRelationshipIds)
          Unchanged = explicitRelId.IsSome && boolOrDefault false (part?SourceUnchanged) })

let private generateHeaderFooterXml rootName (hyperlinks: Map<string, string>) (images: ImageScope) (blocks: obj array) =
    let content = blocks |> generateBlocks (generateDocumentItemWithImages hyperlinks images) (generateTextBoxRun defaultPageSetupInfo hyperlinks images)
    // w:hdr / w:ftr must contain at least one block.
    let content = if content = "" then element "w:p" [] [] else content
    document (element rootName [("xmlns:w", Namespaces.word); ("xmlns:r", Namespaces.relationships); ("xmlns:m", Namespaces.math)] [content])

/// The body-level w:sectPr: the final section. Its header / footer references
/// and page number start come from the page setup; a header / footer part
/// that no section refers to (one made in the editor) is attached here.
let private generateSectionProperties (pageSetup: PageSetupInfo option) (headers: HeaderFooterPartInfo array) (footers: HeaderFooterPartInfo array) =
    let setup = pageSetup |> Option.defaultValue defaultPageSetupInfo
    let refs name (explicit: (string * string) list) (parts: HeaderFooterPartInfo array) =
        let fromParts = parts |> Array.toList |> List.map (fun part -> part.Kind, part.RelId)
        // CT_SectPr allows one reference per type; the section's own references win.
        explicit @ fromParts
        |> List.filter (fun (kind, _) -> includeHeaderFooterKind setup.DifferentFirstPage setup.EvenAndOddHeaders kind)
        |> List.distinctBy fst
        |> List.map (fun (kind, relId) -> selfClosing name [("w:type", kind); ("r:id", relId)])
    let children =
        sectionPropertiesChildren
            (refs "w:headerReference" setup.HeaderReferences headers @ refs "w:footerReference" setup.FooterReferences footers)
            (pageSetup |> Option.bind (fun setup -> setup.SectionStart))
            pageSetup
            setup.PageNumberStart
            setup.PageNumberFormat
    if children.IsEmpty then "" else element "w:sectPr" [] children

/// Generate document.xml with formatting (handles JS arrays)
let generateDocumentXmlWithFormatting (paragraphs: obj array) : string =
    let hyperlinks = hyperlinkMap paragraphs
    let paraElements =
        paragraphs
        |> Array.map (generateDocumentItem hyperlinks)
        |> String.concat ""

    let body = element "w:body" [] [paraElements]

    let doc =
        element "w:document"
            [("xmlns:w", Namespaces.word); ("xmlns:r", Namespaces.relationships); ("xmlns:m", Namespaces.math)]
            [body]

    document doc

let private generateDocumentXmlWithFormattingAndImages (paragraphs: obj array) (images: ImageScope) : string =
    let hyperlinks = hyperlinkMap paragraphs
    let paraElements =
        paragraphs
        |> generateBlocks (generateDocumentItemWithImages hyperlinks images) (generateTextBoxRun defaultPageSetupInfo hyperlinks images)

    let body = element "w:body" [] [paraElements]

    let doc =
        element "w:document"
            [("xmlns:w", Namespaces.word); ("xmlns:r", Namespaces.relationships); ("xmlns:m", Namespaces.math)]
            [body]

    document doc

// ---- the body as it was read ---------------------------------------------------
//
// A document that was opened is not written from the model alone: a child
// element of w:body whose blocks the editor did not change goes back into
// document.xml as the text it was read from, with everything in it the model
// does not hold (revision ids, proofing marks, complex fields, content
// control properties, attributes and elements of later formats). Only the
// blocks that were changed, added or moved next to a change are written from
// the model.

type private ElementRanges = { Starts: int array; Ends: int array; Names: string array; CloseAt: int }

/// Index of the '>' that ends the tag starting at `lt` (an attribute value may hold one).
let private tagEndIndex (xml: string) (lt: int) : int =
    let n = xml.Length
    let mutable i = lt + 1
    let mutable inQuote = false
    let mutable quote = '"'
    let mutable found = -1
    while found < 0 && i < n do
        let c = xml.[i]
        if inQuote then
            if c = quote then inQuote <- false
        elif c = '"' || c = '\'' then
            inQuote <- true
            quote <- c
        elif c = '>' then found <- i
        i <- i + 1
    found

let private tagNameAt (xml: string) (lt: int) : string =
    let mutable i = lt + 1
    let mutable fin = false
    while not fin && i < xml.Length do
        let c = xml.[i]
        if c = ' ' || c = '>' || c = '/' || c = '\n' || c = '\t' || c = '\r' then fin <- true else i <- i + 1
    xml.Substring(lt + 1, i - lt - 1)

let private localNameOf (name: string) =
    let colon = name.IndexOf ':'
    if colon >= 0 then name.Substring(colon + 1) else name

/// Where the child elements of an element are in the text: `from` is just
/// after its start tag; CloseAt is where its end tag starts. None for text
/// that is not well formed.
let private childElementRanges (xml: string) (from: int) : ElementRanges option =
    let n = xml.Length
    let starts = ResizeArray<int>()
    let ends = ResizeArray<int>()
    let names = ResizeArray<string>()
    let mutable pos = from
    let mutable depth = 0
    let mutable closeAt = -1
    let mutable broken = false
    // (Character by character: String.CompareOrdinal with offsets does not compare in the JavaScript build.)
    let startsWithAt (at: int) (value: string) =
        at + value.Length <= n
        && (let mutable same = true
            let mutable i = 0
            while same && i < value.Length do
                if xml.[at + i] <> value.[i] then same <- false
                i <- i + 1
            same)
    let skipTo (at: int) (terminator: string) =
        let found = xml.IndexOf(terminator, at, System.StringComparison.Ordinal)
        if found < 0 then broken <- true else pos <- found + terminator.Length
    while closeAt < 0 && not broken do
        let lt = if pos < n then xml.IndexOf('<', pos) else -1
        if lt < 0 || lt + 1 >= n then broken <- true
        elif startsWithAt lt "<!--" then skipTo lt "-->"
        elif startsWithAt lt "<![CDATA[" then skipTo lt "]]>"
        elif xml.[lt + 1] = '?' then skipTo lt "?>"
        elif xml.[lt + 1] = '/' then
            let gt = xml.IndexOf('>', lt)
            if gt < 0 then broken <- true
            elif depth = 0 then closeAt <- lt
            else
                depth <- depth - 1
                if depth = 0 then ends.Add(gt + 1)
                pos <- gt + 1
        else
            let gt = tagEndIndex xml lt
            if gt < 0 then broken <- true
            else
                let selfClosing = xml.[gt - 1] = '/'
                if depth = 0 then
                    starts.Add lt
                    names.Add (tagNameAt xml lt)
                    if selfClosing then ends.Add(gt + 1)
                if not selfClosing then depth <- depth + 1
                pos <- gt + 1
    if broken || starts.Count <> ends.Count then None
    else Some { Starts = starts.ToArray(); Ends = ends.ToArray(); Names = names.ToArray(); CloseAt = closeAt }

/// Where the root element's start tag begins.
let private rootTagStart (xml: string) : int =
    let mutable pos = 0
    let mutable found = -1
    while found < 0 && pos >= 0 && pos < xml.Length do
        let lt = xml.IndexOf('<', pos)
        if lt < 0 || lt + 1 >= xml.Length then pos <- -1
        elif xml.[lt + 1] = '?' then
            let fin = xml.IndexOf("?>", lt, System.StringComparison.Ordinal)
            pos <- if fin < 0 then -1 else fin + 2
        elif xml.[lt + 1] = '!' then
            let fin = xml.IndexOf('>', lt)
            pos <- if fin < 0 then -1 else fin + 1
        else found <- lt
    found

let private countOccurrences (text: string) (from: int) (until: int) (value: string) : int =
    let mutable count = 0
    let mutable pos = from
    let mutable fin = false
    while not fin do
        let found = text.IndexOf(value, pos, System.StringComparison.Ordinal)
        if found < 0 || found >= until then fin <- true
        else
            count <- count + 1
            pos <- found + value.Length
    count

// ---- a change of text made in the XML a block was read from ---------------------
//
// A block of which only the text of paragraphs was changed (the editor says
// which characters: SourceTextChanges) is not written from the model. The
// change is made in its XML: characters go into and out of the w:t elements
// they belong to, a run whose text is all deleted goes, text that does not
// go into the run next to it comes as a run of its own with the properties
// of a run of the paragraph that is formatted the same (of the paragraph
// mark in a paragraph without text) - at the start or the end of the
// paragraph, between two runs, or inside a run, which is split, both halves
// with its properties. Everything else of the block stays the text it
// was: bookmarks, fields, revision ids, properties the model does not know.
// Whenever the XML cannot be matched with the text the editor had (a field
// whose result was recomputed, an element the reader turned into text in a
// way not known here), the block is written from the model as before.

type private SourceTextSegment = {
    /// The start tag of the w:t (w:delText, m:t) and the text between the tags.
    TagStart: int
    ContentStart: int
    ContentEnd: int
    /// A w:t of a run of the body's own text.
    Editable: bool
    /// The w:r it is in (-1: none).
    RunStart: int
    mutable RunEnd: int
    /// The run holds its properties and this one w:t.
    mutable RunSimple: bool
    /// The run is a child of the paragraph (not of a hyperlink, a tracked change, a content control).
    RunTopLevel: bool
    /// The run is one text is typed into: not one of a content control or a field shown by its own element.
    RunTyped: bool
    /// The run's w:rPr (-1: none).
    RunPropsStart: int
    RunPropsEnd: int
    mutable Text: string
    /// The text without the white space around it when the w:t does not say xml:space="preserve"
    /// (what the reader makes of such an element).
    TrimmedText: string
    /// An element that may stand for text of any kind (a note reference, a symbol) comes before it.
    WideGapBefore: bool
    /// Where its text starts in the paragraph's text.
    mutable At: int
}

type private SourceParagraph = {
    Start: int
    /// The '>' of its start tag.
    TagEnd: int
    SelfClosed: bool
    /// Where its end tag starts, and where the element ends.
    mutable CloseAt: int
    mutable End: int
    Segments: ResizeArray<SourceTextSegment>
    /// Its child elements: where each starts and ends, and its name.
    Children: ResizeArray<int * int * string>
    mutable WideGapAtEnd: bool
    /// Text was found in a form that is not understood here.
    mutable Unknown: bool
}

let private unescapeXmlText (raw: string) : string =
    if raw.IndexOf '&' < 0 then raw
    else
        let sb = System.Text.StringBuilder()
        let mutable i = 0
        while i < raw.Length do
            let c = raw.[i]
            let semi = if c = '&' then raw.IndexOf(';', i) else -1
            if semi > i && semi - i <= 10 then
                let name = raw.Substring(i + 1, semi - i - 1)
                let code =
                    match name with
                    | "amp" -> 38
                    | "lt" -> 60
                    | "gt" -> 62
                    | "quot" -> 34
                    | "apos" -> 39
                    | _ when name.StartsWith "#x" || name.StartsWith "#X" ->
                        (try System.Convert.ToInt32(name.Substring 2, 16) with _ -> -1)
                    | _ when name.StartsWith "#" ->
                        (match System.Int32.TryParse(name.Substring 1) with
                         | true, value -> value
                         | _ -> -1)
                    | _ -> -1
                if code < 0 then
                    sb.Append c |> ignore
                    i <- i + 1
                else
                    if code >= 0x10000 then
                        let v = code - 0x10000
                        sb.Append(char (0xD800 + (v >>> 10))).Append(char (0xDC00 + (v &&& 0x3FF))) |> ignore
                    else sb.Append(char code) |> ignore
                    i <- semi + 1
            else
                sb.Append c |> ignore
                i <- i + 1
        sb.ToString()

let private escapeXmlText (text: string) : string =
    text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")

/// Text for an element whose content was `raw`: quotation marks and
/// apostrophes written as references when the content had them so. (A w:t
/// that a character is typed into is written whole: the characters that were
/// not touched stay the text they were - `&quot;Contribution&quot;` did not.)
let private escapeXmlTextAs (raw: string) (text: string) : string =
    let escaped = escapeXmlText text
    let escaped = if raw.IndexOf "&quot;" >= 0 && raw.IndexOf '"' < 0 then escaped.Replace("\"", "&quot;") else escaped
    if raw.IndexOf "&apos;" >= 0 && raw.IndexOf '\'' < 0 then escaped.Replace("'", "&apos;") else escaped

/// The paragraphs of the body's own text in a piece of document.xml (those
/// of tables, of tables in cells and of block content controls too; not the
/// ones inside drawings and pictures), with where their text is.
let private scanSourceParagraphs (xml: string) : SourceParagraph list option =
    let n = xml.Length
    let paragraphs = ResizeArray<SourceParagraph>()
    // The open elements: name and where each started.
    let names = ResizeArray<string>()
    let starts = ResizeArray<int>()
    let mutable skipDepth = -1
    let mutable current : SourceParagraph option = None
    let mutable paragraphDepth = -1
    let mutable runStart = -1
    let mutable runDepth = -1
    let mutable runTopLevel = false
    let mutable runTyped = false
    let mutable runPropsStart = -1
    let mutable runPropsEnd = -1
    let mutable runFirstSegment = 0
    let mutable runTexts = 0
    let mutable runOthers = 0
    let mutable wideGap = false
    let mutable pos = 0
    let mutable broken = false
    let startsWithAt (at: int) (value: string) =
        at + value.Length <= n
        && (let mutable same = true
            let mutable i = 0
            while same && i < value.Length do
                if xml.[at + i] <> value.[i] then same <- false
                i <- i + 1
            same)
    let isSkipped (name: string) = name = "w:drawing" || name = "w:pict" || name = "mc:AlternateContent" || name = "w:object"
    let isWide (name: string) =
        match name with
        | "w:tab" | "w:br" | "w:cr" | "w:noBreakHyphen" | "w:softHyphen" | "w:ptab" | "w:rPr" | "w:lastRenderedPageBreak" | "w:fldChar" | "w:instrText" -> false
        | _ -> true
    while not broken && pos < n do
        let lt = xml.IndexOf('<', pos)
        if lt < 0 then pos <- n
        elif startsWithAt lt "<!--" then
            let fin = xml.IndexOf("-->", lt, System.StringComparison.Ordinal)
            if fin < 0 then broken <- true else pos <- fin + 3
        elif startsWithAt lt "<![CDATA[" then
            current |> Option.iter (fun para -> para.Unknown <- true)
            let fin = xml.IndexOf("]]>", lt, System.StringComparison.Ordinal)
            if fin < 0 then broken <- true else pos <- fin + 3
        elif lt + 1 < n && xml.[lt + 1] = '?' then
            let fin = xml.IndexOf("?>", lt, System.StringComparison.Ordinal)
            if fin < 0 then broken <- true else pos <- fin + 2
        elif lt + 1 < n && xml.[lt + 1] = '/' then
            let gt = xml.IndexOf('>', lt)
            if gt < 0 || names.Count = 0 then
                // (The end tag of an element that began before this piece: nothing of a paragraph.)
                if gt < 0 then broken <- true else pos <- gt + 1
            else
                let depth = names.Count
                let name = names.[depth - 1]
                names.RemoveAt(depth - 1)
                starts.RemoveAt(depth - 1)
                if skipDepth = depth then skipDepth <- -1
                if skipDepth < 0 then
                    match current with
                    | Some para ->
                        if depth = paragraphDepth then
                            para.CloseAt <- lt
                            para.End <- gt + 1
                            para.WideGapAtEnd <- wideGap
                            paragraphs.Add para
                            current <- None
                            paragraphDepth <- -1
                        elif depth = paragraphDepth + 1 then
                            let (childStart, _, childName) = para.Children.[para.Children.Count - 1]
                            para.Children.[para.Children.Count - 1] <- (childStart, gt + 1, childName)
                        if runDepth >= 0 && depth = runDepth + 1 && name = "w:rPr" then runPropsEnd <- gt + 1
                        if depth = runDepth then
                            for i in runFirstSegment .. para.Segments.Count - 1 do
                                para.Segments.[i].RunEnd <- gt + 1
                                para.Segments.[i].RunSimple <- runTexts = 1 && runOthers = 0
                            runDepth <- -1
                            runStart <- -1
                    | None -> ()
                pos <- gt + 1
        else
            let gt = tagEndIndex xml lt
            if gt < 0 then broken <- true
            else
                let selfClosing = xml.[gt - 1] = '/'
                let name = tagNameAt xml lt
                let depth = names.Count
                if skipDepth < 0 then
                    match current with
                    | None ->
                        if name = "w:p" then
                            let para =
                                { Start = lt; TagEnd = gt; SelfClosed = selfClosing; CloseAt = (if selfClosing then gt + 1 else -1); End = (if selfClosing then gt + 1 else -1)
                                  Segments = ResizeArray(); Children = ResizeArray(); WideGapAtEnd = false; Unknown = false }
                            wideGap <- false
                            if selfClosing then paragraphs.Add para
                            else
                                current <- Some para
                                paragraphDepth <- depth + 1
                    | Some para ->
                        if depth = paragraphDepth then para.Children.Add((lt, (if selfClosing then gt + 1 else -1), name))
                        if name = "w:p" then para.Unknown <- true
                        if isSkipped name then
                            wideGap <- true
                        elif name = "w:r" && not selfClosing then
                            runStart <- lt
                            runDepth <- depth + 1
                            runTopLevel <- depth = paragraphDepth
                            runTyped <-
                                (let mutable plain = true
                                 for i in paragraphDepth .. depth - 1 do
                                     match names.[i] with
                                     | "w:hyperlink" | "w:ins" | "w:smartTag" | "w:moveTo" -> ()
                                     | _ -> plain <- false
                                 plain)
                            runPropsStart <- -1
                            runPropsEnd <- -1
                            runFirstSegment <- para.Segments.Count
                            runTexts <- 0
                            runOthers <- 0
                        elif name = "w:t" || name = "w:delText" || name = "m:t" then
                            let inRun = runDepth >= 0 && depth = runDepth
                            let contentStart = gt + 1
                            let contentEnd =
                                if selfClosing then contentStart
                                else
                                    let next = xml.IndexOf('<', contentStart)
                                    if next < 0 || not (startsWithAt next ("</" + name + ">")) then
                                        para.Unknown <- true
                                        contentStart
                                    else next
                            if inRun then runTexts <- runTexts + 1
                            para.Segments.Add
                                { TagStart = lt; ContentStart = contentStart; ContentEnd = contentEnd
                                  Editable = name = "w:t" && inRun && not selfClosing
                                  RunStart = (if inRun then runStart else -1); RunEnd = -1; RunSimple = false; RunTopLevel = inRun && runTopLevel
                                  RunTyped = inRun && runTyped
                                  RunPropsStart = (if inRun && runPropsEnd > runPropsStart then runPropsStart else -1)
                                  RunPropsEnd = (if inRun && runPropsEnd > runPropsStart then runPropsEnd else -1)
                                  Text = unescapeXmlText (xml.Substring(contentStart, contentEnd - contentStart))
                                  TrimmedText =
                                    (let text = unescapeXmlText (xml.Substring(contentStart, contentEnd - contentStart))
                                     if xml.Substring(lt, gt - lt).IndexOf "xml:space=\"preserve\"" >= 0 then text else text.Trim())
                                  WideGapBefore = wideGap; At = -1 }
                            wideGap <- false
                        else
                            if runDepth >= 0 && depth = runDepth && name = "w:rPr" then
                                runPropsStart <- lt
                                runPropsEnd <- if selfClosing then gt + 1 else -1
                            if runDepth >= 0 && depth = runDepth && name <> "w:rPr" then runOthers <- runOthers + 1
                            if runDepth >= 0 && depth = runDepth && isWide name then wideGap <- true
                            // (Text of an equation, a field shown by its own element, ...: of any kind.)
                            if name = "w:fldSimple" || name = "m:oMath" || name = "m:oMathPara" || name = "w:ruby" then wideGap <- true
                if not selfClosing then
                    names.Add name
                    starts.Add lt
                    if skipDepth < 0 && isSkipped name then skipDepth <- depth + 1
                pos <- gt + 1
    if broken || current.IsSome then None else Some (List.ofSeq paragraphs)

/// Characters the reader makes of elements that are not w:t (a tab, a break,
/// a hyphen, the place of an object).
let private isPlaceholderCharacter (c: char) =
    c < ' ' || c = '￼' || c = '­' || c = '‑' || c = ' '

/// Where each piece of text of a paragraph's XML is in the paragraph's text
/// as the editor had it. False when the two do not match.
let rec private alignSourceParagraph (para: SourceParagraph) (text: string) : bool =
    if para.Unknown then false
    else
        alignSourceSegments para text
        || (// As the reader has the text of elements that do not preserve their white space.
            para.Segments |> Seq.exists (fun segment -> segment.Text <> segment.TrimmedText)
            && (for segment in para.Segments do segment.Text <- segment.TrimmedText
                alignSourceSegments para text))

and private alignSourceSegments (para: SourceParagraph) (text: string) : bool =
        let mutable cursor = 0
        let mutable ok = true
        let gapAllowed (from: int) (until: int) (wide: bool) =
            wide
            || (let mutable all = true
                for i in from .. until - 1 do
                    if not (isPlaceholderCharacter text.[i]) then all <- false
                all)
        for segment in para.Segments do
            if ok then
                if segment.Text.Length = 0 then segment.At <- cursor
                else
                    let found = text.IndexOf(segment.Text, cursor, System.StringComparison.Ordinal)
                    if found < 0 || not (gapAllowed cursor found segment.WideGapBefore) then ok <- false
                    else
                        segment.At <- found
                        cursor <- found + segment.Text.Length
        ok && gapAllowed cursor text.Length para.WideGapAtEnd

let private needsSpacePreserved (text: string) =
    text.Length > 0 && (System.Char.IsWhiteSpace text.[0] || System.Char.IsWhiteSpace text.[text.Length - 1] || text.IndexOf "  " >= 0)

/// The replacements (from, until, text) that make one change of text in a
/// paragraph's XML.
/// `modelProps`: the properties the model writes for the text that was typed (a w:rPr, or ""), for a
/// run of its own where no run of the paragraph is formatted as it is.
let private planParagraphChange (xml: string) (para: SourceParagraph) (oldText: string) (change: obj) (modelProps: string option) : (int * int * string) list option =
    let offset = intOrDefault 0 (change?Offset)
    let deleted = intOrDefault 0 (change?Deleted)
    let inserted = stringOrDefault "" (change?Inserted)
    let joinBefore = boolOrDefault false (change?JoinBefore)
    let joinAfter = boolOrDefault false (change?JoinAfter)
    let runsPlain = boolOrDefault false (change?RunsPlain)
    let witnesses = optionalArray (change?Witnesses) |> Array.map (fun at -> intOrDefault -1 at)
    let markStyle = boolOrDefault false (change?MarkStyle)
    let deleteEnd = offset + deleted
    if offset < 0 || deleteEnd > oldText.Length || not (alignSourceParagraph para oldText) then None
    else
        let segments = para.Segments |> Seq.filter (fun segment -> segment.Text.Length > 0) |> Seq.toArray
        let segmentAt (position: int) =
            segments |> Array.tryFind (fun segment -> segment.At <= position && position < segment.At + segment.Text.Length)
        // Every deleted character is one of a w:t.
        let mutable covered = 0
        let mutable deletable = true
        for segment in segments do
            let from = max offset segment.At
            let until = min deleteEnd (segment.At + segment.Text.Length)
            if until > from then
                covered <- covered + (until - from)
                if not segment.Editable then deletable <- false
        let typable = inserted |> Seq.forall (fun c -> c >= ' ' && c <> '￼')
        let joinTarget =
            if inserted.Length = 0 || not typable then None
            else
                let typedInto (segment: SourceTextSegment) = segment.Editable && segment.RunTyped
                let before = if joinBefore && offset > 0 then segmentAt (offset - 1) |> Option.filter typedInto else None
                match before with
                | Some _ -> before
                | None -> if joinAfter && deleteEnd < oldText.Length then segmentAt deleteEnd |> Option.filter typedInto else None
        // A run of its own for the text: with the properties of a run that is formatted as it is,
        // or of the paragraph mark when the paragraph has no text.
        let newRun : string option =
            if inserted.Length = 0 || not runsPlain || not typable then None
            else
                let props =
                    if witnesses.Length > 0 then
                        witnesses
                        |> Array.tryPick (fun witness -> segmentAt witness |> Option.filter (fun segment -> segment.Editable && segment.RunTyped))
                        |> Option.map (fun segment -> if segment.RunPropsStart >= 0 then xml.Substring(segment.RunPropsStart, segment.RunPropsEnd - segment.RunPropsStart) else "")
                    elif joinAfter && offset = 0 && deleted = 0 && not para.SelfClosed then
                        // Typed before a tab or a break the paragraph begins with, and formatted as it is (no
                        // run of text is): the properties of the run that holds the tab - the paragraph's
                        // first run, when that is what it holds.
                        para.Children
                        |> Seq.tryFind (fun (_, _, name) -> name <> "w:pPr")
                        |> Option.bind (fun (start, fin, name) ->
                            let run = xml.Substring(start, fin - start)
                            let gt = tagEndIndex run 0
                            let holds (tag: string) = run.IndexOf tag >= 0
                            if name <> "w:r" || gt < 0 || run.[gt - 1] = '/' then None
                            elif not (holds "<w:tab/>" || holds "<w:br/>" || holds "<w:br ") then None
                            elif holds "fldChar" || holds "instrText" || holds "rPrChange" || holds "<w:drawing" || holds "<w:pict" || holds "<w:object" then None
                            elif run.IndexOf("<w:rPr/>", gt + 1) = gt + 1 then Some ""
                            elif run.IndexOf("<w:rPr>", gt + 1) = gt + 1 then
                                let close = run.IndexOf("</w:rPr>", gt + 1)
                                if close < 0 then None else Some (run.Substring(gt + 1, close + 8 - (gt + 1)))
                            else Some "")
                    elif markStyle && not para.SelfClosed && para.Children.Count > 0 then
                        let (start, fin, name) = para.Children.[0]
                        if name <> "w:pPr" then Some ""
                        else
                            let tagEnd = tagEndIndex xml start
                            if tagEnd < 0 || xml.[tagEnd - 1] = '/' then Some ""
                            else
                                childElementRanges xml (tagEnd + 1)
                                |> Option.map (fun children ->
                                    match children.Names |> Array.tryFindIndex (fun child -> child = "w:rPr") with
                                    | Some index when children.Ends.[index] <= fin -> xml.Substring(children.Starts.[index], children.Ends.[index] - children.Starts.[index])
                                    | _ -> "")
                    elif markStyle then Some ""
                    else None
                props
                |> Option.orElse modelProps
                // (Not the properties of a tracked change.)
                |> Option.filter (fun rPr -> rPr.IndexOf "Change" < 0 && rPr.IndexOf "<w:ins" < 0 && rPr.IndexOf "<w:del" < 0 && rPr.IndexOf "<w:move" < 0)
                |> Option.map (fun rPr -> "<w:r>" + rPr + (if needsSpacePreserved inserted then "<w:t xml:space=\"preserve\">" else "<w:t>") + escapeXmlText inserted + "</w:t></w:r>")
        if not deletable || covered <> deleted then None
        elif inserted.Length > 0 && joinTarget.IsNone && newRun.IsNone then None
        else
            let runsToken = defaultArg newRun ""
            let replacements = ResizeArray<int * int * string>()
            let setText (segment: SourceTextSegment) (text: string) =
                let tagEnd = segment.ContentStart
                let tag = xml.Substring(segment.TagStart, tagEnd - segment.TagStart)
                let tag = if needsSpacePreserved text && tag.IndexOf "xml:space" < 0 then tag.Substring(0, tag.Length - 1) + " xml:space=\"preserve\">" else tag
                replacements.Add((segment.TagStart, segment.ContentEnd, tag + escapeXmlTextAs (xml.Substring(segment.ContentStart, segment.ContentEnd - segment.ContentStart)) text))
            let remaining (segment: SourceTextSegment) =
                let from = max offset segment.At - segment.At
                let until = min deleteEnd (segment.At + segment.Text.Length) - segment.At
                if until > from then segment.Text.Substring(0, from), segment.Text.Substring until else segment.Text, ""
            let isTarget (segment: SourceTextSegment) = match joinTarget with Some target -> obj.ReferenceEquals(target, segment) | None -> false
            // The segment a new run would split: one w:t holds the character before and the one after the change.
            let splitSegment =
                if inserted.Length = 0 || joinTarget.IsSome || offset = 0 || deleteEnd = oldText.Length then None
                else
                    match segmentAt (offset - 1), segmentAt deleteEnd with
                    | Some a, Some b when obj.ReferenceEquals(a, b) -> Some a
                    | _ -> None
            let mutable ok = true
            for segment in segments do
                let touched = min deleteEnd (segment.At + segment.Text.Length) > max offset segment.At
                let split = match splitSegment with Some target -> obj.ReferenceEquals(target, segment) | None -> false
                if isTarget segment then
                    let at = max 0 (min segment.Text.Length (offset - segment.At))
                    let left, right = if touched then remaining segment else segment.Text.Substring(0, at), segment.Text.Substring at
                    setText segment (left + inserted + right)
                elif split then
                    if not (segment.RunSimple && segment.RunTopLevel && segment.Editable) then ok <- false
                    else
                        let left, right = if touched then remaining segment else (let at = offset - segment.At in segment.Text.Substring(0, at), segment.Text.Substring at)
                        // The run's start tag and properties: both halves have them.
                        let head = xml.Substring(segment.RunStart, segment.TagStart - segment.RunStart)
                        let raw = xml.Substring(segment.ContentStart, segment.ContentEnd - segment.ContentStart)
                        let half (text: string) = head + "<w:t xml:space=\"preserve\">" + escapeXmlTextAs raw text + "</w:t></w:r>"
                        replacements.Add((segment.RunStart, segment.RunEnd, half left + runsToken + half right))
                elif touched then
                    let left, right = remaining segment
                    if left.Length = 0 && right.Length = 0 && segment.RunSimple then replacements.Add((segment.RunStart, segment.RunEnd, ""))
                    else setText segment (left + right)
            // Runs of their own that split no run: at an end of the paragraph, or between two of its runs.
            if ok && inserted.Length > 0 && joinTarget.IsNone && splitSegment.IsNone then
                if offset = 0 then
                    if para.SelfClosed then
                        replacements.Add((para.Start, para.End, xml.Substring(para.Start, para.TagEnd - 1 - para.Start).TrimEnd() + ">" + runsToken + "</w:p>"))
                    else
                        let mutable at = para.TagEnd + 1
                        let mutable leading = true
                        para.Children
                        |> Seq.iteri (fun index (start, fin, name) ->
                            if leading then
                                // After the properties, and after a page break the paragraph begins with
                                // (the reader makes it a property of the paragraph: the text comes after it).
                                let pageBreak =
                                    name = "w:r" && fin > start
                                    && (let run = xml.Substring(start, fin - start)
                                        run.IndexOf "w:type=\"page\"" >= 0 && run.IndexOf "<w:t" < 0 && run.IndexOf "<w:drawing" < 0)
                                if (index = 0 && name = "w:pPr") || pageBreak then at <- fin
                                else leading <- false)
                        replacements.Add((at, at, runsToken))
                elif deleteEnd = oldText.Length then replacements.Add((para.CloseAt, para.CloseAt, runsToken))
                else
                    match segmentAt (offset - 1), segmentAt deleteEnd with
                    | Some a, Some b when a.RunTopLevel && b.RunTopLevel ->
                        let childIndex (segment: SourceTextSegment) = para.Children |> Seq.tryFindIndex (fun (start, _, _) -> start = segment.RunStart)
                        match childIndex a, childIndex b with
                        | Some first, Some last when first < last ->
                            // A field that begins or ends between the two: the new text is outside it.
                            let mutable begins = 0
                            let mutable finishes = 0
                            for i in first + 1 .. last - 1 do
                                let (start, fin, _) = para.Children.[i]
                                begins <- begins + countOccurrences xml start fin "w:fldCharType=\"begin\""
                                finishes <- finishes + countOccurrences xml start fin "w:fldCharType=\"end\""
                            if finishes = 0 then replacements.Add((a.RunEnd, a.RunEnd, runsToken))
                            elif begins = 0 then replacements.Add((b.RunStart, b.RunStart, runsToken))
                            else ok <- false
                        | _ -> ok <- false
                    | _ -> ok <- false
            if ok then Some (List.ofSeq replacements) else None

/// How many paragraphs of the body's own text an item of the editor's package holds.
let rec private storyParagraphCount (item: obj) : int =
    match item?Type |> stringOrDefault "Paragraph" with
    | "TextBox" -> 0
    | "Table" ->
        optionalArray (item?Rows)
        |> Array.sumBy (fun row ->
            optionalArray (row?Cells)
            |> Array.sumBy (fun cell -> (optionalArray (cell?Paragraphs)).Length + (optionalArray (cell?NestedTables) |> Array.sumBy storyParagraphCount)))
    | _ -> 1

// ---- a change of formatting made in the XML a paragraph was read from ---------
//
// The editor says of a paragraph whose text is what it was what it was
// before its formatting was changed (the paragraph as the model writes it,
// Before) and what it is now (After). Neither is written: what differs
// between the properties the model writes for the two - of the paragraph, and
// of the run each character is in - is set in the properties the paragraph's
// own XML has. An element the model writes differently replaces the one of
// that name (attribute by attribute where both are a single tag: an attribute
// the model does not write stays, and the order of the attributes), one it
// writes no more goes, a new one stands where the schema puts it. Where the
// formatting of only a part of a run changed the run is split, each part with
// the run's start tag (its revision ids) and properties, the run's children
// shared out by the characters they stand for. Everything else of the
// paragraph stays the text it was read from: bookmarks, fields, hyperlinks
// and content controls around the runs, comment ranges, proofing marks,
// properties and elements the model does not hold. Where the XML cannot be
// matched with the characters the editor shows (a run that stands for text in
// a way not known here, where the change does not cover all it may stand
// for), or holds a tracked change of the properties themselves, the paragraph
// is written from the model.

/// The schema's order of the children of w:rPr and of w:pPr (local names).
let private runPropertyOrder =
    [ "rStyle"; "rFonts"; "b"; "bCs"; "i"; "iCs"; "caps"; "smallCaps"; "strike"; "dstrike"; "outline"; "shadow"; "emboss"; "imprint"
      "noProof"; "snapToGrid"; "vanish"; "webHidden"; "color"; "spacing"; "w"; "kern"; "position"; "sz"; "szCs"; "highlight"; "u"
      "effect"; "bdr"; "shd"; "fitText"; "vertAlign"; "rtl"; "cs"; "em"; "lang"; "eastAsianLayout"; "specVanish"; "oMath" ]
let private paragraphPropertyOrder =
    [ "pStyle"; "keepNext"; "keepLines"; "pageBreakBefore"; "framePr"; "widowControl"; "numPr"; "suppressLineNumbers"; "pBdr"; "shd"
      "tabs"; "suppressAutoHyphens"; "kinsoku"; "wordWrap"; "overflowPunct"; "topLinePunct"; "autoSpaceDE"; "autoSpaceDN"; "bidi"
      "adjustRightInd"; "snapToGrid"; "spacing"; "ind"; "contextualSpacing"; "mirrorIndents"; "suppressOverlap"; "jc"; "textDirection"
      "textAlignment"; "textboxTightWrap"; "outlineLvl"; "divId"; "cnfStyle"; "rPr"; "sectPr"; "pPrChange" ]

/// The child elements of the element a text is of (`<w:x ...> ... </w:x>` or
/// `<w:x .../>`): (name, where it starts, where it ends) of each, and where the
/// element's content ends (-1: the element has none, it is closed in its start
/// tag). None when the text is not that of one well-formed element.
let private childElementsOf (element: string) : ((string * int * int) list * int) option =
    if element.Length < 4 || element.[0] <> '<' then None
    else
        let gt = tagEndIndex element 0
        if gt < 0 then None
        elif element.[gt - 1] = '/' then (if gt = element.Length - 1 then Some ([], -1) else None)
        else
            let found = ResizeArray<string * int * int>()
            let mutable pos = gt + 1
            let mutable depth = 0
            let mutable childStart = -1
            let mutable childName = ""
            let mutable contentEnd = -1
            let mutable ok = true
            while ok && contentEnd < 0 do
                let lt = element.IndexOf('<', pos)
                if lt < 0 || lt + 1 >= element.Length then ok <- false
                elif element.[lt + 1] = '!' || element.[lt + 1] = '?' then ok <- false
                elif element.[lt + 1] = '/' then
                    let fin = element.IndexOf('>', lt)
                    if fin < 0 then ok <- false
                    elif depth = 0 then
                        if fin = element.Length - 1 then contentEnd <- lt else ok <- false
                    else
                        depth <- depth - 1
                        if depth = 0 then found.Add ((childName, childStart, fin + 1))
                        pos <- fin + 1
                else
                    let fin = tagEndIndex element lt
                    if fin < 0 then ok <- false
                    else
                        if depth = 0 then
                            childStart <- lt
                            childName <- tagNameAt element lt
                        if element.[fin - 1] = '/' then
                            if depth = 0 then found.Add ((childName, childStart, fin + 1))
                        else depth <- depth + 1
                        pos <- fin + 1
            if ok then Some (List.ofSeq found, contentEnd) else None

/// The attributes of a start tag, in their order: (name, value as it is written).
let private attributesOfTag (tag: string) : (string * string) list =
    let result = ResizeArray<string * string>()
    let n = tag.Length
    let isSpace (c: char) = c = ' ' || c = '\t' || c = '\n' || c = '\r'
    // (After the name.)
    let mutable i = 1
    while i < n && not (isSpace tag.[i]) && tag.[i] <> '/' && tag.[i] <> '>' do i <- i + 1
    let mutable fin = false
    while not fin do
        while i < n && isSpace tag.[i] do i <- i + 1
        if i >= n || tag.[i] = '/' || tag.[i] = '>' then fin <- true
        else
            let eq = tag.IndexOf('=', i)
            if eq < 0 then fin <- true
            else
                let name = tag.Substring(i, eq - i).Trim()
                let mutable q = eq + 1
                while q < n && tag.[q] <> '"' && tag.[q] <> '\'' do q <- q + 1
                if q >= n then fin <- true
                else
                    let close = tag.IndexOf(tag.[q], q + 1)
                    if close < 0 then fin <- true
                    else
                        let value = tag.Substring(q + 1, close - q - 1)
                        result.Add ((name, (if tag.[q] = '\'' then value.Replace("\"", "&quot;") else value)))
                        i <- close + 1
    List.ofSeq result

/// An element of properties as it is to be: the one the XML has (None: it has
/// none) with what the model changed of it - `was` and `now` are the element
/// as the model writes it before and after the change (None: not at all).
/// None: the element goes.
let private mergedPropertyElement (source: string option) (was: string option) (now: string option) : string option =
    match source, now with
    | _, None -> None
    | None, Some _ -> now
    | Some sourceText, Some nowText ->
        // (A single tag: nothing in it but attributes.)
        let simple (text: string) = text.EndsWith "/>" && text.IndexOf('<', 1) < 0
        let name = tagNameAt sourceText 0
        if not (simple sourceText && simple nowText && (match was with Some text -> simple text | None -> true)) || tagNameAt nowText 0 <> name then now
        else
            let wasAttributes = match was with Some text -> attributesOfTag text | None -> []
            let nowAttributes = attributesOfTag nowText
            let valueIn (attributes: (string * string) list) (key: string) = attributes |> List.tryFind (fun (k, _) -> k = key) |> Option.map snd
            // What the model writes now is what the element says (also where the XML said it
            // another way than the model did before: a spacing that was automatic); an attribute
            // the model wrote and writes no more goes; one it never writes stays - but for those
            // that would overrule what is set: the same through the document's theme
            // (w:asciiTheme beside w:ascii, w:themeColor beside w:val), or in other units
            // (w:afterLines beside w:after, w:leftChars beside w:left).
            let removed = wasAttributes |> List.map fst |> List.filter (fun key -> (valueIn nowAttributes key).IsNone)
            let overrules (key: string) =
                let lower = key.ToLowerInvariant()
                (valueIn nowAttributes key).IsNone
                && (if name = "w:rFonts" then lower.Contains "theme" && nowAttributes |> List.exists (fun (k, _) -> lower = k.ToLowerInvariant() + "theme")
                    else lower.Contains "theme" || key.EndsWith "Lines" || key.EndsWith "Chars")
            let kept =
                attributesOfTag sourceText
                |> List.filter (fun (key, _) -> not (List.contains key removed) && not (overrules key))
                |> List.map (fun (key, value) -> key, (valueIn nowAttributes key |> Option.defaultValue value))
            let added = nowAttributes |> List.filter (fun (key, _) -> (valueIn kept key).IsNone)
            Some ("<" + name + (kept @ added |> List.map (fun (key, value) -> " " + key + "=\"" + value + "\"") |> String.concat "") + "/>")

/// What differs between two lists of property elements as the model writes
/// them: (local name, the element before, the element after).
let private propertyChanges (was: string list) (now: string list) : (string * string option * string option) list =
    let named (items: string list) = items |> List.map (fun text -> localNameOf (tagNameAt text 0), text)
    let before = named was
    let after = named now
    [ for (name, text) in after do
          match before |> List.tryFind (fun (known, _) -> known = name) with
          | Some (_, old) when old = text -> ()
          | Some (_, old) -> yield name, Some old, Some text
          | None -> yield name, None, Some text
      for (name, old) in before do
          if not (after |> List.exists (fun (known, _) -> known = name)) then yield name, Some old, None ]

/// The children of a properties element as the model wrote it (`<w:pPr>...</w:pPr>`, `<w:pPr/>`).
let private propertyElementsOf (element: string) : string list =
    match childElementsOf element with
    | Some (children, _) -> children |> List.map (fun (_, start, fin) -> element.Substring(start, fin - start))
    | None -> []

/// The text of a properties element (w:rPr, w:pPr; "" - there is none yet)
/// with the changes made in it (see mergedPropertyElement); "" when nothing is
/// left in it. None when the element is not one this can be done in.
let rec private withPropertiesMerged (container: string) (order: string list) (properties: string) (changes: (string * string option * string option) list) : string option =
    let rankOf (name: string) = order |> List.tryFindIndex (fun known -> "w:" + known = name)
    if changes.IsEmpty then Some properties
    elif changes |> List.exists (fun (name, _, _) -> not (List.contains name order)) then None
    else
        let parsed =
            if properties = "" then Some ([], "<" + container + ">")
            else
                match childElementsOf properties with
                | Some (children, contentEnd) when tagNameAt properties 0 = container ->
                    let gt = tagEndIndex properties 0
                    let startTag = if contentEnd < 0 then properties.Substring(0, gt - 1).TrimEnd() + ">" else properties.Substring(0, gt + 1)
                    Some (children |> List.map (fun (name, start, fin) -> name, properties.Substring(start, fin - start)), startTag)
                | _ -> None
        match parsed with
        | None -> None
        // (A tracked change of the properties holds the properties as they were: not touched here.)
        | Some (children, _) when children |> List.exists (fun (name, _) -> name.EndsWith "Change") -> None
        | Some (children, startTag) ->
            // What is there keeps its place; an element the schema does not name (of a later
            // format) stands after the one before it.
            let result = ResizeArray<int * string * string>()
            let mutable lastRank = -1
            for (name, text) in children do
                let rank = match rankOf name with Some rank -> rank | None -> lastRank
                lastRank <- max lastRank rank
                result.Add ((rank, name, text))
            // (The complex-script twin of bold and italic follows them where the XML states it.)
            let changes =
                changes
                |> List.collect (fun (name, was, now) ->
                    let twin = name + "Cs"
                    if (name = "b" || name = "i") && children |> List.exists (fun (child, _) -> child = "w:" + twin) && not (changes |> List.exists (fun (other, _, _) -> other = twin)) then
                        let asTwin (text: string option) = text |> Option.map (fun text -> "<w:" + twin + text.Substring(("<w:" + name).Length))
                        [ name, was, now; twin, asTwin was, asTwin now ]
                    else [ name, was, now ])
            let mutable ok = true
            for (name, was, now) in changes do
                let full = "w:" + name
                let at = result |> Seq.tryFindIndex (fun (_, child, _) -> child = full)
                let source = at |> Option.map (fun index -> let (_, _, text) = result.[index] in text)
                let merged =
                    if name = "rPr" && container = "w:pPr" then
                        // The properties of the paragraph mark: changed child by child, as a run's.
                        let inner = propertyChanges (was |> Option.map propertyElementsOf |> Option.defaultValue []) (now |> Option.map propertyElementsOf |> Option.defaultValue [])
                        match withPropertiesMerged "w:rPr" runPropertyOrder (defaultArg source "") inner with
                        | Some "" -> None
                        | Some text -> Some text
                        | None ->
                            ok <- false
                            None
                    else mergedPropertyElement source was now
                match at, merged with
                | Some index, Some text -> let (rank, child, _) = result.[index] in result.[index] <- (rank, child, text)
                | Some index, None -> result.RemoveAt index
                | None, Some text ->
                    let rank = (rankOf full).Value
                    let mutable position = 0
                    while position < result.Count && (let (known, _, _) = result.[position] in known <= rank) do position <- position + 1
                    result.Insert (position, (rank, full, text))
                | None, None -> ()
            if not ok then None
            elif result.Count = 0 then Some ""
            else Some (startTag + (result |> Seq.map (fun (_, _, text) -> text) |> String.concat "") + "</" + container + ">")

/// A run of a paragraph's XML and the characters of the paragraph's text it stands for.
type private SourceRunInfo = {
    RunStart: int
    RunEnd: int
    /// Just after its start tag.
    ContentStart: int
    /// Its w:rPr (-1: none).
    PropsStart: int
    PropsEnd: int
    /// Its other children: name, where each starts and ends.
    Parts: (string * int * int) array
    /// How many characters each stands for (-1: not known).
    PartChars: int array
    /// The w:t (w:delText) of each part that is one (null: another element).
    PartSegments: SourceTextSegment option array
    /// The characters it stands for, where they are known: [From, Until). -1: not known.
    mutable From: int
    mutable Until: int
    /// Every character of [From, Until) is the one of a part.
    mutable Exact: bool
    /// Where the characters may be that are not known to be its own.
    mutable ReachFrom: int
    mutable ReachUntil: int
    /// The field whose result it is a run of (a number of the paragraph's; -1: of none).
    mutable Field: int
}

/// The runs of a paragraph's own text (not those of drawings and text boxes in it), in their order.
/// The paragraph's segments are aligned with `text`. None: the XML is not understood.
let private sourceRunsOf (xml: string) (para: SourceParagraph) (text: string) : SourceRunInfo array option =
    let runs = ResizeArray<SourceRunInfo>()
    let segmentAt = System.Collections.Generic.Dictionary<int, SourceTextSegment>()
    for segment in para.Segments do segmentAt.[segment.TagStart] <- segment
    let isSkipped (name: string) = name = "w:drawing" || name = "w:pict" || name = "mc:AlternateContent" || name = "w:object"
    let mutable pos = para.TagEnd + 1
    let mutable ok = not para.SelfClosed
    let mutable depth = 0
    let mutable skipDepth = -1
    // The fields that are open: of w:fldSimple (the depth it is at), and of w:fldChar (whether its result has begun).
    let simpleFields = ResizeArray<int * int>()
    let complexFields = ResizeArray<int * bool>()
    let mutable nextField = 1
    let fin = para.CloseAt
    while ok && pos < fin do
        let lt = xml.IndexOf('<', pos)
        if lt < 0 || lt >= fin then pos <- fin
        elif lt + 1 < xml.Length && (xml.[lt + 1] = '!' || xml.[lt + 1] = '?') then
            let gt = xml.IndexOf('>', lt)
            if gt < 0 then ok <- false else pos <- gt + 1
        elif xml.[lt + 1] = '/' then
            let gt = xml.IndexOf('>', lt)
            if gt < 0 then ok <- false
            else
                if skipDepth = depth then skipDepth <- -1
                if simpleFields.Count > 0 && fst simpleFields.[simpleFields.Count - 1] = depth then simpleFields.RemoveAt(simpleFields.Count - 1)
                depth <- depth - 1
                pos <- gt + 1
        else
            let gt = tagEndIndex xml lt
            if gt < 0 then ok <- false
            else
                let selfClosing = xml.[gt - 1] = '/'
                let name = tagNameAt xml lt
                if skipDepth < 0 && name = "w:fldSimple" && not selfClosing then
                    simpleFields.Add ((depth + 1, nextField))
                    nextField <- nextField + 1
                if skipDepth < 0 && name = "w:r" && not selfClosing then
                    // The whole run: to the end tag at its own depth.
                    let mutable inner = 1
                    let mutable at = gt + 1
                    let mutable runEnd = -1
                    while ok && runEnd < 0 do
                        let next = xml.IndexOf('<', at)
                        if next < 0 || next >= fin then ok <- false
                        elif xml.[next + 1] = '/' then
                            let close = xml.IndexOf('>', next)
                            inner <- inner - 1
                            if inner = 0 then runEnd <- close + 1
                            at <- close + 1
                        elif xml.[next + 1] = '!' || xml.[next + 1] = '?' then ok <- false
                        else
                            let close = tagEndIndex xml next
                            if close < 0 then ok <- false
                            else
                                if xml.[close - 1] <> '/' then inner <- inner + 1
                                at <- close + 1
                    if ok then
                        match childElementsOf (xml.Substring(lt, runEnd - lt)) with
                        | None -> ok <- false
                        | Some (children, _) ->
                            let children = children |> List.map (fun (child, start, close) -> child, lt + start, lt + close)
                            let props = children |> List.tryFind (fun (child, _, _) -> child = "w:rPr")
                            let parts = children |> List.filter (fun (child, _, _) -> child <> "w:rPr") |> List.toArray
                            let segments = parts |> Array.map (fun (child, start, _) -> if child = "w:t" || child = "w:delText" then (match segmentAt.TryGetValue start with | true, segment -> Some segment | _ -> None) else None)
                            let chars =
                                parts
                                |> Array.mapi (fun index (child, start, close) ->
                                    match child with
                                    | "w:t" | "w:delText" -> (match segments.[index] with Some segment -> segment.Text.Length | None -> -1)
                                    | "w:tab" | "w:cr" | "w:noBreakHyphen" | "w:softHyphen" -> 1
                                    | "w:br" -> if xml.Substring(start, close - start).IndexOf "w:type=\"" >= 0 && xml.Substring(start, close - start).IndexOf "w:type=\"textWrapping\"" < 0 then -1 else 1
                                    | "w:lastRenderedPageBreak" -> 0
                                    | _ -> -1)
                            // The field it is a run of the result of; the marks of a field open and close it.
                            let mutable marks = false
                            for (child, start, close) in parts do
                                if child = "w:fldChar" then
                                    marks <- true
                                    let tag = xml.Substring(start, close - start)
                                    if tag.IndexOf "\"begin\"" >= 0 then
                                        complexFields.Add ((nextField, false))
                                        nextField <- nextField + 1
                                    elif tag.IndexOf "\"separate\"" >= 0 then
                                        if complexFields.Count > 0 then complexFields.[complexFields.Count - 1] <- (fst complexFields.[complexFields.Count - 1], true)
                                    elif tag.IndexOf "\"end\"" >= 0 then
                                        if complexFields.Count > 0 then complexFields.RemoveAt(complexFields.Count - 1)
                                        else
                                            // (A field that began in a paragraph before: what came so far is of its result.)
                                            for known in runs do
                                                if known.Field < 0 then known.Field <- 0
                                elif child = "w:instrText" then marks <- true
                            let field =
                                if marks then -1
                                elif simpleFields.Count > 0 then snd simpleFields.[simpleFields.Count - 1]
                                else
                                    match complexFields |> Seq.tryFindBack snd with
                                    | Some (id, _) -> id
                                    | None -> -1
                            runs.Add
                                { RunStart = lt; RunEnd = runEnd; ContentStart = gt + 1
                                  Field = field
                                  PropsStart = (match props with Some (_, start, _) -> start | None -> -1)
                                  PropsEnd = (match props with Some (_, _, close) -> close | None -> -1)
                                  Parts = parts; PartChars = chars; PartSegments = segments
                                  From = -1; Until = -1; Exact = false; ReachFrom = -1; ReachUntil = -1 }
                            pos <- runEnd
                else
                    if not selfClosing then
                        depth <- depth + 1
                        if skipDepth < 0 && isSkipped name then skipDepth <- depth
                    pos <- gt + 1
    if not ok then None
    else
        // The character a part of one character stands for.
        let standsFor (name: string) (c: char) =
            match name with
            | "w:tab" -> c = '\t'
            | "w:br" | "w:cr" -> c = '\n' || c = ' ' || c = '\u000B'
            | "w:noBreakHyphen" -> c = '‑'
            | "w:softHyphen" -> c = '­'
            | _ -> true
        // Each run with text: from the text it holds, and what its other parts are known to stand for.
        for run in runs do
            let anchors = [ for index in 0 .. run.Parts.Length - 1 do match run.PartSegments.[index] with Some segment when segment.Text.Length > 0 && segment.At >= 0 -> yield index, segment | _ -> () ]
            match anchors with
            | [] -> ()
            | (firstIndex, firstSegment) :: _ ->
                let (lastIndex, lastSegment) = List.last anchors
                let mutable exact = true
                let mutable from = firstSegment.At
                for index in firstIndex - 1 .. -1 .. 0 do
                    if run.PartChars.[index] < 0 then exact <- false
                    elif exact then from <- from - run.PartChars.[index]
                let mutable until = lastSegment.At + lastSegment.Text.Length
                let mutable trailing = true
                for index in lastIndex + 1 .. run.Parts.Length - 1 do
                    if run.PartChars.[index] < 0 then
                        exact <- false
                        trailing <- false
                    elif trailing then until <- until + run.PartChars.[index]
                // Between its texts: what the parts there stand for is what the text has there.
                let mutable at = from
                if exact then
                    for index in 0 .. run.Parts.Length - 1 do
                        let (name, _, _) = run.Parts.[index]
                        match run.PartSegments.[index] with
                        | Some segment when segment.Text.Length > 0 -> if segment.At <> at then exact <- false
                        | _ -> ()
                        if exact && run.PartChars.[index] = 1 && run.PartSegments.[index].IsNone then
                            if at < 0 || at >= text.Length || not (standsFor name text.[at]) then exact <- false
                        at <- at + max 0 run.PartChars.[index]
                    if at <> until then exact <- false
                if exact && from >= 0 && until <= text.Length then
                    run.From <- from
                    run.Until <- until
                    run.Exact <- true
                else
                    run.From <- firstSegment.At
                    run.Until <- lastSegment.At + lastSegment.Text.Length
                    run.Exact <- false
        // The runs without text between them: one after the other where what each stands for is
        // known and fills the gap; otherwise they share the gap.
        let mutable cursor = 0
        let mutable index = 0
        let mutable sound = true
        while sound && index < runs.Count do
            if runs.[index].From >= 0 then
                if runs.[index].From < cursor then sound <- false
                else
                    runs.[index].ReachFrom <- (if runs.[index].Exact then runs.[index].From else cursor)
                    cursor <- runs.[index].Until
                    index <- index + 1
            else
                let first = index
                while index < runs.Count && runs.[index].From < 0 do index <- index + 1
                let gapEnd = if index < runs.Count then runs.[index].From else text.Length
                if gapEnd < cursor then sound <- false
                else
                    let total (run: SourceRunInfo) = if run.PartChars |> Array.exists (fun count -> count < 0) then -1 else Array.sum run.PartChars
                    let known = seq { first .. index - 1 } |> Seq.forall (fun i -> total runs.[i] >= 0)
                    let sum = if known then seq { first .. index - 1 } |> Seq.sumBy (fun i -> total runs.[i]) else -1
                    let mutable fits = known && sum = gapEnd - cursor
                    if fits then
                        // (And each part stands for the character that is there.)
                        let mutable at = cursor
                        for i in first .. index - 1 do
                            for part in 0 .. runs.[i].Parts.Length - 1 do
                                let (name, _, _) = runs.[i].Parts.[part]
                                if runs.[i].PartChars.[part] = 1 && not (standsFor name text.[at]) then fits <- false
                                at <- at + runs.[i].PartChars.[part]
                    if fits then
                        let mutable at = cursor
                        for i in first .. index - 1 do
                            runs.[i].From <- at
                            runs.[i].Until <- at + total runs.[i]
                            runs.[i].Exact <- true
                            runs.[i].ReachFrom <- at
                            runs.[i].ReachUntil <- runs.[i].Until
                            at <- runs.[i].Until
                    else
                        for i in first .. index - 1 do
                            runs.[i].ReachFrom <- cursor
                            runs.[i].ReachUntil <- gapEnd
                    cursor <- gapEnd
        if not sound then None
        else
            // A run with text of which not every part is known reaches to the texts next to it.
            for i in 0 .. runs.Count - 1 do
                let run = runs.[i]
                if run.From >= 0 && run.ReachUntil < 0 then
                    if run.Exact then run.ReachUntil <- run.Until
                    else
                        let next = seq { i + 1 .. runs.Count - 1 } |> Seq.tryFind (fun j -> runs.[j].From >= 0)
                        run.ReachUntil <- (match next with Some j -> runs.[j].From | None -> text.Length)
            Some (runs.ToArray())

/// How many characters of a paragraph's text a run of the editor's package stands for.
let private runObjectLength (run: obj) : int =
    if boolOrDefault false (run?IsImage) then 1 else (stringOrDefault "" (run?Text)).Length

/// The change of a paragraph's formatting (see above) as replacements in the
/// XML: (from, until, text). `text` is the paragraph's text, which `before`
/// and `after` both have. None: not in this XML.
let private planParagraphProperties (xml: string) (para: SourceParagraph) (text: string) (before: obj) (after: obj) : (int * int * string) list option =
    let runsOf (item: obj) = optionalArray (item?Runs)
    let byCharacter (item: obj) =
        let result = ResizeArray<int>()
        runsOf item |> Array.iteri (fun index run -> for _ in 1 .. runObjectLength run do result.Add index)
        result.ToArray()
    let oldRuns = runsOf before
    let newRuns = runsOf after
    let oldAt = byCharacter before
    let newAt = byCharacter after
    // (An item of a list the file does not have: the list's number is known when the lists are written.)
    let ofNewList = (let style = after?ListStyle |> stringOrDefault "none" in style = "bullet" || style = "numbered") && (optionalIntString (after?ListId)).IsNone
    if para.SelfClosed || ofNewList || oldAt.Length <> text.Length || newAt.Length <> text.Length || not (alignSourceParagraph para text) then None
    else
        // What changed of the properties of the run each character is in: by the pair of runs.
        let known = System.Collections.Generic.Dictionary<int, (string * string option * string option) list * string>()
        let changesAt (i: int) : ((string * string option * string option) list * string) option =
            let oldRun = oldRuns.[oldAt.[i]]
            let newRun = newRuns.[newAt.[i]]
            // (A picture has no properties of text.)
            if boolOrDefault false (oldRun?IsImage) || boolOrDefault false (newRun?IsImage) then None
            else
                let key = oldAt.[i] * (newRuns.Length + 1) + newAt.[i]
                match known.TryGetValue key with
                | true, found -> Some found
                | _ ->
                    let changes = propertyChanges (modelRunProperties oldRun) (modelRunProperties newRun)
                    let signature = changes |> List.map (fun (name, was, now) -> name + "\u0001" + defaultArg was "" + "\u0001" + defaultArg now "") |> String.concat "\u0002"
                    known.[key] <- (changes, signature)
                    Some (changes, signature)
        let anyRunChange = seq { 0 .. text.Length - 1 } |> Seq.exists (fun i -> match changesAt i with Some (changes, _) -> not changes.IsEmpty | None -> false)
        let replacements = ResizeArray<int * int * string>()
        let mutable ok = true
        if anyRunChange then
            // (A check box of a content control is shown by the control's state, not by the run that holds its mark.)
            if xml.Substring(para.Start, para.End - para.Start).Contains "checkbox" then ok <- false
            else
                match sourceRunsOf xml para text with
                | None -> ok <- false
                | Some runs ->
                    let covered = Array.zeroCreate<bool> text.Length
                    let propsOf (run: SourceRunInfo) = if run.PropsStart >= 0 then xml.Substring(run.PropsStart, run.PropsEnd - run.PropsStart) else ""
                    // (A run that shows nothing: the marks of a field, its code.)
                    let shows (run: SourceRunInfo) =
                        run.Parts |> Array.exists (fun (name, _, _) -> name <> "w:fldChar" && name <> "w:instrText" && name <> "w:lastRenderedPageBreak" && name <> "w:annotationRef")
                    // The result of a field is read as one run (DocxReader): its runs are changed alike, or not here.
                    let ofField = System.Collections.Generic.Dictionary<int, string>()
                    for run in runs do
                        if ok then
                            let from = if run.Exact then run.From else run.ReachFrom
                            let until = if run.Exact then run.Until else run.ReachUntil
                            if run.Field >= 0 then
                                for i in from .. until - 1 do
                                    match changesAt i with
                                    | Some (_, signature) ->
                                        match ofField.TryGetValue run.Field with
                                        | true, known -> if known <> signature then ok <- false
                                        | _ -> ofField.[run.Field] <- signature
                                    | None -> ()
                            // The changes of the characters it stands for, or may stand for.
                            let distinct = ResizeArray<string>()
                            let mutable first : (string * string option * string option) list = []
                            for i in from .. until - 1 do
                                match changesAt i with
                                | Some (changes, signature) ->
                                    if not (distinct.Contains signature) then
                                        if distinct.Count = 0 then first <- changes
                                        distinct.Add signature
                                | None -> ()
                            if distinct.Count = 1 then
                                if not first.IsEmpty && shows run then
                                    match withPropertiesMerged "w:rPr" runPropertyOrder (propsOf run) first with
                                    | Some properties ->
                                        if run.PropsStart >= 0 then replacements.Add ((run.PropsStart, run.PropsEnd, properties))
                                        else replacements.Add ((run.ContentStart, run.ContentStart, properties))
                                        for i in from .. until - 1 do covered.[i] <- true
                                    | None -> ok <- false
                                elif first.IsEmpty then
                                    for i in from .. until - 1 do covered.[i] <- true
                            elif distinct.Count > 1 then
                                if not run.Exact then ok <- false
                                else
                                    // Split where the change changes: each piece the run's start tag, its
                                    // properties with the piece's change, and the parts of the piece.
                                    let head = xml.Substring(run.RunStart, run.ContentStart - run.RunStart)
                                    let props = propsOf run
                                    let signatureAt (i: int) = match changesAt i with Some (_, signature) -> signature | None -> ""
                                    let pieces = ResizeArray<int * int>()
                                    let mutable pieceStart = run.From
                                    for i in run.From + 1 .. run.Until do
                                        if i = run.Until || signatureAt i <> signatureAt (i - 1) then
                                            pieces.Add ((pieceStart, i))
                                            pieceStart <- i
                                    let built = System.Text.StringBuilder()
                                    for (pieceFrom, pieceUntil) in pieces do
                                        if ok then
                                            let changes = match changesAt pieceFrom with Some (changes, _) -> changes | None -> []
                                            match withPropertiesMerged "w:rPr" runPropertyOrder props changes with
                                            | None -> ok <- false
                                            | Some properties ->
                                                built.Append(head).Append(properties) |> ignore
                                                let last = pieceUntil = run.Until
                                                let mutable at = run.From
                                                for index in 0 .. run.Parts.Length - 1 do
                                                    let (_, start, close) = run.Parts.[index]
                                                    let count = run.PartChars.[index]
                                                    if count = 0 then
                                                        // (What stands for no character goes with the character after it.)
                                                        if (at >= pieceFrom && at < pieceUntil) || (last && at = run.Until) then built.Append(xml.Substring(start, close - start)) |> ignore
                                                    elif at >= pieceFrom && at + count <= pieceUntil then built.Append(xml.Substring(start, close - start)) |> ignore
                                                    elif at < pieceUntil && at + count > pieceFrom then
                                                        // A text that the piece has a part of.
                                                        match run.PartSegments.[index] with
                                                        | Some segment ->
                                                            let cutFrom = max pieceFrom at - at
                                                            let cutUntil = min pieceUntil (at + count) - at
                                                            let piece = segment.Text.Substring(cutFrom, cutUntil - cutFrom)
                                                            let tag = xml.Substring(segment.TagStart, segment.ContentStart - segment.TagStart)
                                                            let tag = if needsSpacePreserved piece && tag.IndexOf "xml:space" < 0 then tag.Substring(0, tag.Length - 1) + " xml:space=\"preserve\">" else tag
                                                            let name = tagNameAt tag 0
                                                            built.Append(tag).Append(escapeXmlTextAs (xml.Substring(segment.ContentStart, segment.ContentEnd - segment.ContentStart)) piece).Append("</" + name + ">") |> ignore
                                                        | None -> ok <- false
                                                    at <- at + count
                                                built.Append("</w:r>") |> ignore
                                    if ok then
                                        replacements.Add ((run.RunStart, run.RunEnd, built.ToString()))
                                        for i in run.From .. run.Until - 1 do covered.[i] <- true
                    // Every character whose formatting changed is one of a run that was changed.
                    if ok then
                        for i in 0 .. text.Length - 1 do
                            match changesAt i with
                            | Some (changes, _) when not changes.IsEmpty && not covered.[i] -> ok <- false
                            | _ -> ()
        // The properties of the paragraph.
        if ok then
            let was = propertyElementsOf (generateParaPropsWithPageSetup defaultPageSetupInfo before)
            let now = propertyElementsOf (generateParaPropsWithPageSetup defaultPageSetupInfo after)
            let changes = propertyChanges was now
            if changes |> List.exists (fun (name, _, _) -> name = "sectPr") then ok <- false
            elif not changes.IsEmpty then
                match para.Children |> Seq.tryFind (fun (_, _, name) -> name = "w:pPr") with
                | Some (start, fin, _) ->
                    match withPropertiesMerged "w:pPr" paragraphPropertyOrder (xml.Substring(start, fin - start)) changes with
                    | Some properties -> replacements.Add ((start, fin, properties))
                    | None -> ok <- false
                | None ->
                    match withPropertiesMerged "w:pPr" paragraphPropertyOrder "" changes with
                    | Some properties -> if properties <> "" then replacements.Add ((para.TagEnd + 1, para.TagEnd + 1, properties))
                    | None -> ok <- false
        if ok then Some (List.ofSeq replacements) else None

/// A text with replacements (from, until, text) made in it. None when two of them overlap.
let private withReplacements (xml: string) (replacements: (int * int * string) list) : string option =
    let ordered = replacements |> List.sortByDescending (fun (from, _, _) -> from) |> List.toArray
    let mutable text = xml
    let mutable limit = xml.Length
    let mutable ok = true
    for (from, until, replacement) in ordered do
        if until > limit || from > until then ok <- false
        elif ok then
            text <- text.Substring(0, from) + replacement + text.Substring until
            limit <- from
    if ok then Some text else None

/// What the last save made of the body (see Main's getLastSaveReport).
let mutable private reportPatched = 0
let mutable private reportInline = 0
let mutable private reportModelBlocks = 0
let mutable private reportModelParagraphs = 0
let mutable private reportSourceBlocks = 0
/// Why blocks and paragraphs were written from the model (the first few).
let private reportReasons = ResizeArray<string>()
let private noteReason (reason: string) = if reportReasons.Count < 24 then reportReasons.Add reason
let lastSaveReport () : obj =
    createObj [
        "Reasons" ==> reportReasons.ToArray()
        "PatchedParagraphs" ==> reportPatched
        "ModelParagraphsInTables" ==> reportInline
        "ModelBlocks" ==> reportModelBlocks
        "ModelBlockParagraphs" ==> reportModelParagraphs
        "SourceBlocks" ==> reportSourceBlocks ]

/// Where a paragraph of a table that is written from the model goes in the
/// table's XML: its number between two of these (see BodyPlan.Inline).
let private inlineMarkCharacter = '\u0001'
let private inlineMark (number: int) = string inlineMarkCharacter + string number + string inlineMarkCharacter
let mutable private nextInlineNumber = 0

/// The XML of one paragraph with what was done to the paragraph (the editor's
/// account of it: Docx.paragraphEdit) done in it - the change of its text,
/// then that of its formatting. None when the XML does not allow it.
let private patchSourceParagraph (paragraphXml: string) (oldText: string) (edit: obj) : string option =
    if boolOrDefault false (edit?Regenerate) then None
    else
        let scanned (xml: string) =
            match scanSourceParagraphs xml with
            | Some [ para ] when para.Start = 0 && para.End = xml.Length -> Some para
            | _ -> None
        let textChange = edit?Text
        let withText =
            if isNullOrUndefined textChange then Some paragraphXml
            else
                // The properties the model writes for what was typed (the run the first new character is in).
                let modelProps =
                    let offset = intOrDefault 0 (textChange?Offset)
                    let mutable at = 0
                    let mutable found : string option = None
                    if not (isNullOrUndefined (edit?After)) && (stringOrDefault "" (textChange?Inserted)).Length > 0 then
                        for run in optionalArray (edit?After?Runs) do
                            let length = runObjectLength run
                            if found.IsNone && offset >= at && offset < at + length && not (boolOrDefault false (run?IsImage)) then
                                found <- Some (match modelRunProperties run with [] -> "" | properties -> element "w:rPr" [] properties)
                            at <- at + length
                    found
                scanned paragraphXml
                |> Option.bind (fun para -> planParagraphChange paragraphXml para oldText textChange modelProps)
                |> Option.bind (withReplacements paragraphXml)
        withText
        |> Option.bind (fun xml ->
            let before = edit?Before
            if isNullOrUndefined before then Some xml
            else
                let text = if isNullOrUndefined (edit?NewText) then oldText else stringOrDefault "" (edit?NewText)
                scanned xml
                |> Option.bind (fun para -> planParagraphProperties xml para text before (edit?After))
                |> Option.bind (withReplacements xml))

/// The XML of an element of the body with what was done to the paragraphs of
/// its items done in it (see patchSourceParagraph). A paragraph of a table
/// whose XML does not allow that is marked to be written from the model, the
/// one paragraph (the marks and their paragraphs are the second of the
/// pair); in a paragraph of the body itself that fails the whole element is
/// (None), as when the XML does not hold the paragraphs the editor had.
let private patchSourceChunk (xml: string) (items: obj array) : (string * (int * obj) list * int) option =
    match scanSourceParagraphs xml with
    | None -> None
    | Some paragraphs ->
        let paragraphs = List.toArray paragraphs
        if (items |> Array.sumBy storyParagraphCount) <> paragraphs.Length then
            noteReason (sprintf "the XML of the block holds %d paragraphs, the editor %d" paragraphs.Length (items |> Array.sumBy storyParagraphCount))
            None
        else
            let replacements = ResizeArray<int * int * string>()
            let inlined = ResizeArray<int * obj>()
            let mutable ok = true
            let mutable first = 0
            for item in items do
                let count = storyParagraphCount item
                let changes = item?SourceTextChanges
                let inTable = (item?Type |> stringOrDefault "Paragraph") = "Table"
                if ok && not (isNullOrUndefined changes) then
                    let texts = optionalArray (changes?Texts) |> Array.map (fun text -> stringOrDefault "" text)
                    if texts.Length <> count then ok <- false
                    else
                        let edited = System.Collections.Generic.HashSet<int>()
                        for change in optionalArray (changes?Edits) do
                            let index = intOrDefault -1 (change?Paragraph)
                            if ok then
                                if index < 0 || index >= count || not (edited.Add index) then ok <- false
                                else
                                    let para = paragraphs.[first + index]
                                    let own = xml.Substring(para.Start, para.End - para.Start)
                                    match patchSourceParagraph own texts.[index] change with
                                    | Some patched -> replacements.Add ((para.Start, para.End, patched))
                                    | None ->
                                        // (A shape anchored in the paragraph is not the paragraph's in the model: with it the table is written.)
                                        let holdsShape = own.IndexOf "<mc:AlternateContent" >= 0 || own.IndexOf "<w:pict" >= 0 || own.IndexOf "<wps:" >= 0 || own.IndexOf "<wpg:" >= 0
                                        // (A field that begins in the paragraph and ends in another: the paragraph is not on its own.)
                                        let fieldOpen = countOccurrences own 0 own.Length "w:fldCharType=\"begin\"" <> countOccurrences own 0 own.Length "w:fldCharType=\"end\""
                                        let why = if boolOrDefault false (change?Regenerate) then "changed in more than its text and formatting" else "the change cannot be made in its XML"
                                        if not inTable || holdsShape || fieldOpen || isNullOrUndefined (change?After) then
                                            noteReason (sprintf "paragraph %d of the block: %s%s" index why (if not inTable then "" elif holdsShape then "; a shape is anchored in it: the table is written" elif fieldOpen then "; a field begins or ends in it: the table is written" else ""))
                                            ok <- false
                                        else
                                            noteReason (sprintf "paragraph %d of a table: %s" index why)
                                            nextInlineNumber <- nextInlineNumber + 1
                                            inlined.Add ((nextInlineNumber, change?After))
                                            replacements.Add ((para.Start, para.End, inlineMark nextInlineNumber))
                        // The others are the paragraphs the editor had there: the first text of each is in its text.
                        for index in 0 .. count - 1 do
                            if ok && not (edited.Contains index) then
                                let para = paragraphs.[first + index]
                                match para.Segments |> Seq.tryFind (fun segment -> segment.Editable && segment.Text.Length > 0) with
                                | Some segment when texts.[index].IndexOf(segment.Text, System.StringComparison.Ordinal) < 0 ->
                                    noteReason (sprintf "paragraph %d of the block is not the paragraph the editor had there" index)
                                    ok <- false
                                | _ -> ()
                first <- first + count
            if not ok then None
            else withReplacements xml (List.ofSeq replacements) |> Option.map (fun text -> text, List.ofSeq inlined, replacements.Count - inlined.Count)

type private BodySegment =
    | SourceText of string
    | ModelBlocks of obj array

type private BodyPlan = {
    /// The text up to and including the start tag of w:body.
    Prefix: string
    Segments: BodySegment list
    /// What stands between the last block and the body's section properties.
    BeforeSection: string
    /// The body's own w:sectPr as it was read, when the page setup is what it was.
    Section: string option
    AfterSection: string
    /// From the end tag of w:body on.
    Suffix: string
    /// The items written from the model.
    Regenerated: obj array
    /// The paragraphs of tables that are written from the model into the table's own XML, by
    /// the number of the mark that stands where each goes (see patchSourceChunk).
    Inline: Map<int, obj>
}

[<Fable.Core.Emit("(typeof $0.SourceIndex === 'number' ? $0.SourceIndex : -1)")>]
let private sourceIndexOf (_item: obj) : int = jsNative

/// Which of the body's elements go back as they were read. None when the
/// text cannot be matched with what the reader made of it (the body is then
/// written from the model, as a new document is), or when nothing is as it was.
let private planBody (xml: string) (itemChunks: int array) (chunkCount: int) (items: obj array) (pageSetupUnchanged: bool) : BodyPlan option =
    let rootAt = rootTagStart xml
    let rootEnd = if rootAt < 0 then -1 else tagEndIndex xml rootAt
    let bodyOfRoot =
        if rootEnd < 0 || xml.[rootEnd - 1] = '/' then None
        else
            childElementRanges xml (rootEnd + 1)
            |> Option.bind (fun children ->
                children.Names
                |> Array.tryFindIndex (fun name -> localNameOf name = "body")
                |> Option.bind (fun index ->
                    let bodyTagEnd = tagEndIndex xml children.Starts.[index]
                    if bodyTagEnd < 0 || xml.[bodyTagEnd - 1] = '/' then None
                    else childElementRanges xml (bodyTagEnd + 1) |> Option.map (fun body -> bodyTagEnd + 1, body)))
    match bodyOfRoot with
    | Some (contentStart, body) when body.Starts.Length = chunkCount && chunkCount > 0 ->
        let contentEnd = body.CloseAt
        let isSection index = localNameOf body.Names.[index] = "sectPr"
        let sectIndex = if isSection (chunkCount - 1) then chunkCount - 1 else -1
        let strayFinalSection = [ 0 .. chunkCount - 2 ] |> List.exists isSection
        // How many items each element gave when it was read.
        let originalCount = Array.zeroCreate<int> chunkCount
        let mutable mapped = true
        for chunk in itemChunks do
            if chunk >= chunkCount then mapped <- false
            elif chunk >= 0 then originalCount.[chunk] <- originalCount.[chunk] + 1
        if strayFinalSection || not mapped || (sectIndex >= 0 && originalCount.[sectIndex] > 0) then None
        else
            // An element's text starts where the block element before it ends: what
            // stands between two blocks (a bookmark end, white space) goes with the second.
            let textStart = Array.zeroCreate<int> chunkCount
            let mutable previousEnd = contentStart
            for chunk in 0 .. chunkCount - 1 do
                textStart.[chunk] <- previousEnd
                if originalCount.[chunk] > 0 then previousEnd <- body.Ends.[chunk]
            let tailStart = previousEnd
            // The items there are now, by the element they came from.
            let chunkOfItem =
                items
                |> Array.map (fun item ->
                    let index = sourceIndexOf item
                    if index >= 0 && index < itemChunks.Length then itemChunks.[index] else -1)
            let isTextBox (item: obj) = (item?Type |> stringOrDefault "Paragraph") = "TextBox"
            let keep = Array.init chunkCount (fun chunk -> originalCount.[chunk] > 0)
            let firstAt = Array.create chunkCount (-1)
            let lastAt = Array.create chunkCount (-1)
            let found = Array.zeroCreate<int> chunkCount
            let lastSource = Array.create chunkCount (-1)
            let patched = Array.zeroCreate<bool> chunkCount
            let patchedText = Array.create<(string * (int * obj) list * int) option> chunkCount None
            let mutable anchorChunk = -1
            items
            |> Array.iteri (fun position item ->
                let chunk = chunkOfItem.[position]
                if chunk >= 0 then
                    if firstAt.[chunk] < 0 then firstAt.[chunk] <- position
                    // (All of an element's items stand together, in the order they were read in.)
                    if lastAt.[chunk] >= 0 && lastAt.[chunk] <> position - 1 then keep.[chunk] <- false
                    let index = sourceIndexOf item
                    if index <= lastSource.[chunk] then keep.[chunk] <- false
                    lastSource.[chunk] <- index
                    lastAt.[chunk] <- position
                    found.[chunk] <- found.[chunk] + 1
                    if not (boolOrDefault false (item?SourceUnchanged)) then
                        // (Changed in its text alone: the change is made in the element's XML.)
                        if isNullOrUndefined (item?SourceTextChanges) then
                            if keep.[chunk] then noteReason (sprintf "block %d: %s" position (optionalString (item?SourceWhy) |> Option.defaultValue "changed"))
                            keep.[chunk] <- false
                        else patched.[chunk] <- true
                if isTextBox item then
                    // A shape is written into the block it is anchored in: one that is new,
                    // or that now follows another block, makes that block's XML.
                    if chunk <> anchorChunk then
                        if chunk >= 0 then keep.[chunk] <- false
                        if anchorChunk >= 0 then keep.[anchorChunk] <- false
                else anchorChunk <- chunk)
            for chunk in 0 .. chunkCount - 1 do
                if found.[chunk] <> originalCount.[chunk] then keep.[chunk] <- false
                if keep.[chunk] && patched.[chunk] then
                    let text = xml.Substring(textStart.[chunk], body.Ends.[chunk] - textStart.[chunk])
                    patchedText.[chunk] <- patchSourceChunk text (Array.sub items firstAt.[chunk] originalCount.[chunk])
                    if patchedText.[chunk].IsNone then keep.[chunk] <- false
            // A field that begins in one element and ends in a later one (a table of
            // contents): those elements go back together or not at all.
            let blockChunks = [| for chunk in 0 .. chunkCount - 1 do if originalCount.[chunk] > 0 then yield chunk |]
            let mutable depth = 0
            let group = ResizeArray<int>()
            let closeGroup () =
                if group.Count > 1 then
                    let together =
                        group |> Seq.forall (fun chunk -> keep.[chunk])
                        && (Seq.pairwise group |> Seq.forall (fun (a, b) -> firstAt.[b] = lastAt.[a] + 1))
                    if not together then
                        for chunk in group do keep.[chunk] <- false
                group.Clear()
            for chunk in blockChunks do
                let begins = countOccurrences xml textStart.[chunk] body.Ends.[chunk] "w:fldCharType=\"begin\""
                let finishes = countOccurrences xml textStart.[chunk] body.Ends.[chunk] "w:fldCharType=\"end\""
                if depth > 0 || begins > finishes then
                    group.Add chunk
                    depth <- max 0 (depth + begins - finishes)
                    if depth = 0 then closeGroup ()
            closeGroup ()
            // In the order the items have now.
            let segments = ResizeArray<BodySegment>()
            let pending = ResizeArray<obj>()
            let regenerated = ResizeArray<obj>()
            let inlined = ResizeArray<int * obj>()
            let flush () =
                if pending.Count > 0 then
                    segments.Add (ModelBlocks (pending.ToArray()))
                    pending.Clear()
            let mutable position = 0
            let mutable anySource = false
            while position < items.Length do
                let chunk = chunkOfItem.[position]
                if chunk >= 0 && keep.[chunk] then
                    flush ()
                    match (if patched.[chunk] then patchedText.[chunk] else None) with
                    | Some (text, inlineParagraphs, patchedCount) ->
                        segments.Add (SourceText text)
                        reportPatched <- reportPatched + patchedCount
                        for (number, para) in inlineParagraphs do
                            inlined.Add ((number, para))
                            regenerated.Add para
                            reportInline <- reportInline + 1
                    | None -> segments.Add (SourceText (xml.Substring(textStart.[chunk], body.Ends.[chunk] - textStart.[chunk])))
                    anySource <- true
                    reportSourceBlocks <- reportSourceBlocks + originalCount.[chunk]
                    position <- position + originalCount.[chunk]
                else
                    pending.Add items.[position]
                    regenerated.Add items.[position]
                    reportModelBlocks <- reportModelBlocks + 1
                    reportModelParagraphs <- reportModelParagraphs + storyParagraphCount items.[position]
                    position <- position + 1
            flush ()
            // (Nothing as it was read: a new body. Around it the document stays the text it
            // was - the root's start tag, the section properties - when the page setup is.)
            if not anySource && not pageSetupUnchanged then None
            else
                let sectionStart = if sectIndex >= 0 then body.Starts.[sectIndex] else contentEnd
                let sectionEnd = if sectIndex >= 0 then body.Ends.[sectIndex] else contentEnd
                Some { Prefix = xml.Substring(0, contentStart)
                       Segments = List.ofSeq segments
                       BeforeSection = xml.Substring(tailStart, max 0 (sectionStart - tailStart))
                       Section = if pageSetupUnchanged then Some (xml.Substring(sectionStart, sectionEnd - sectionStart)) else None
                       AfterSection = xml.Substring(sectionEnd, contentEnd - sectionEnd)
                       Suffix = xml.Substring contentEnd
                       Regenerated = regenerated.ToArray()
                       Inline = Map.ofSeq inlined }
    | _ -> None

/// The root start tag with the namespaces the writer's own elements use.
/// The root tag of a part with the namespaces declared that what was written
/// from the model (`written`) uses and the root does not declare: the root
/// element of a document stays as it was read unless something needs more.
let private withWriterNamespaces (prefix: string) (written: string list) : string =
    let rootAt = rootTagStart prefix
    if rootAt < 0 then prefix
    else
        let name = tagNameAt prefix rootAt
        let insertAt = rootAt + 1 + name.Length
        let uses (namespacePrefix: string) =
            written
            |> List.exists (fun text ->
                text.IndexOf("<" + namespacePrefix + ":", System.StringComparison.Ordinal) >= 0
                || text.IndexOf(" " + namespacePrefix + ":", System.StringComparison.Ordinal) >= 0)
        let missing =
            [ "xmlns:w", "w", Namespaces.word; "xmlns:r", "r", Namespaces.relationships; "xmlns:m", "m", Namespaces.math ]
            |> List.filter (fun (attribute, namespacePrefix, _) -> uses namespacePrefix && prefix.IndexOf(" " + attribute + "=", System.StringComparison.Ordinal) < 0)
            |> List.map (fun (attribute, _, uri) -> attribute, uri)
            |> List.map (fun (attribute, uri) -> sprintf " %s=\"%s\"" attribute uri)
            |> String.concat ""
        prefix.Substring(0, insertAt) + missing + prefix.Substring insertAt

/// document.xml of a document that was opened: the elements of the plan that
/// are as they were read, the others from the model.
let private generateDocumentXmlFromPlan (plan: BodyPlan) (images: ImageScope) (hyperlinks: Map<string, string>) (pageSetup: PageSetupInfo option) (headers: HeaderFooterPartInfo array) (footers: HeaderFooterPartInfo array) : string =
    let setup = pageSetup |> Option.defaultValue defaultPageSetupInfo
    let parts = ResizeArray<string>()
    let written = ResizeArray<string>()
    for segment in plan.Segments do
        match segment with
        | SourceText text when text.IndexOf inlineMarkCharacter < 0 -> parts.Add text
        | SourceText text ->
            // The paragraphs of a table that are written from the model: each where its mark is.
            let pieces = text.Split inlineMarkCharacter
            pieces
            |> Array.iteri (fun index piece ->
                if index % 2 = 0 then parts.Add piece
                else
                    match System.Int32.TryParse piece with
                    | true, number when plan.Inline.ContainsKey number ->
                        let paragraph = generateParagraphWithHyperlinksAndImagesWithPageSetup setup hyperlinks images plan.Inline.[number]
                        written.Add paragraph
                        parts.Add paragraph
                    | _ -> ())
        | ModelBlocks items ->
            let text = items |> generateBlocks (generateDocumentItemWithImagesAndPageSetup setup hyperlinks images) (generateTextBoxRun setup hyperlinks images)
            written.Add text
            parts.Add text
    parts.Add plan.BeforeSection
    match plan.Section with
    | Some text -> parts.Add text
    | None ->
        let globalHeaders = headers |> Array.filter (fun part -> part.IsGlobal)
        let globalFooters = footers |> Array.filter (fun part -> part.IsGlobal)
        let text = generateSectionProperties pageSetup globalHeaders globalFooters
        written.Add text
        parts.Add text
    parts.Add plan.AfterSection
    parts.Add plan.Suffix
    (if written.Count > 0 then withWriterNamespaces plan.Prefix (List.ofSeq written) else plan.Prefix) + String.concat "" parts

let private numberAttributePattern = System.Text.RegularExpressions.Regex("w:(?:numId|abstractNumId)=\"(\\d+)\"|<w:(?:numId|abstractNumId) w:val=\"(\\d+)\"")
let private numElementPattern = System.Text.RegularExpressions.Regex("<w:num(?:\\s[^>]*)?\\sw:numId=\"(\\d+)\"")
let private firstNumPattern = System.Text.RegularExpressions.Regex("<w:num[\\s>]")

let generateDocumentXmlPackage (paragraphs: obj array) (images: ImageScope) (pageSetup: PageSetupInfo option) (headers: HeaderFooterPartInfo array) (footers: HeaderFooterPartInfo array) : string =
    let hyperlinks = hyperlinkMap paragraphs
    let setup = pageSetup |> Option.defaultValue defaultPageSetupInfo
    let paraElements =
        paragraphs
        |> generateBlocks (generateDocumentItemWithImagesAndPageSetup setup hyperlinks images) (generateTextBoxRun setup hyperlinks images)
    let globalHeaders = headers |> Array.filter (fun part -> part.IsGlobal)
    let globalFooters = footers |> Array.filter (fun part -> part.IsGlobal)
    let body = element "w:body" [] [paraElements; generateSectionProperties pageSetup globalHeaders globalFooters]
    let doc = element "w:document" [("xmlns:w", Namespaces.word); ("xmlns:r", Namespaces.relationships); ("xmlns:m", Namespaces.math)] [body]
    document doc

/// Generate [Content_Types].xml
let generateContentTypes (hasComments: bool) (hasFootnotes: bool) (hasEndnotes: bool) : string =
    let children = [
        yield selfClosing "Default" [("Extension", "rels"); ("ContentType", "application/vnd.openxmlformats-package.relationships+xml")]
        yield selfClosing "Default" [("Extension", "xml"); ("ContentType", "application/xml")]
        yield selfClosing "Override" [("PartName", "/word/document.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")]
        yield selfClosing "Override" [("PartName", "/word/styles.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml")]
        yield selfClosing "Override" [("PartName", "/word/numbering.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml")]
        if hasComments then
            yield selfClosing "Override" [("PartName", "/word/comments.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml")]
        if hasFootnotes then
            yield selfClosing "Override" [("PartName", "/word/footnotes.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml")]
        if hasEndnotes then
            yield selfClosing "Override" [("PartName", "/word/endnotes.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml")]
    ]
    let root = element "Types" [("xmlns", Namespaces.contentTypes)] children
    document root

let private generateContentTypesWithImages hasComments hasFootnotes hasEndnotes (imageAssets: ImageAsset array) =
    let children = [
        yield selfClosing "Default" [("Extension", "rels"); ("ContentType", "application/vnd.openxmlformats-package.relationships+xml")]
        yield selfClosing "Default" [("Extension", "xml"); ("ContentType", "application/xml")]
        for extension, contentType in imageAssets |> Array.filter (fun image -> not image.Linked) |> Array.map (fun image -> image.Extension, image.ContentType) |> Array.distinct do
            yield selfClosing "Default" [("Extension", extension); ("ContentType", contentType)]
        yield selfClosing "Override" [("PartName", "/word/document.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")]
        yield selfClosing "Override" [("PartName", "/word/styles.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml")]
        yield selfClosing "Override" [("PartName", "/word/numbering.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml")]
        if hasComments then yield selfClosing "Override" [("PartName", "/word/comments.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml")]
        if hasFootnotes then yield selfClosing "Override" [("PartName", "/word/footnotes.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml")]
        if hasEndnotes then yield selfClosing "Override" [("PartName", "/word/endnotes.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml")]
    ]
    document (element "Types" [("xmlns", Namespaces.contentTypes)] children)

/// [Content_Types].xml for a package that carries kept parts: the original
/// Default entries, the writer's own, and the Override entries of the kept parts.
let private generateContentTypesPreserved (preserved: PreservedPackage) hasNumbering hasSettings hasComments hasFootnotes hasEndnotes (headers: HeaderFooterPartInfo array) (footers: HeaderFooterPartInfo array) (imageAssets: ImageAsset array) =
    let own =
        [ yield "/word/document.xml", preserved.MainContentType |> Option.defaultValue "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
          yield "/word/styles.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
          if hasNumbering then yield "/word/numbering.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
          if hasSettings then yield "/word/settings.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
          for header in headers do yield header.PartName, "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
          for footer in footers do yield footer.PartName, "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
          if hasComments then yield "/word/comments.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
          if hasFootnotes then yield "/word/footnotes.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
          if hasEndnotes then yield "/word/endnotes.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml" ]
    let ownNames = own |> List.map (fun (name, _) -> name.ToLowerInvariant()) |> Set.ofList
    let keptNames = preserved.Parts |> List.map (fun (path, _) -> ("/" + path).ToLowerInvariant()) |> Set.ofList
    let keptOverrides =
        preserved.Overrides
        |> List.filter (fun (name, _) -> Set.contains (name.ToLowerInvariant()) keptNames && not (Set.contains (name.ToLowerInvariant()) ownNames))
    let defaults =
        [ yield "rels", "application/vnd.openxmlformats-package.relationships+xml"
          yield "xml", "application/xml"
          yield! preserved.Defaults
          for image in imageAssets do
              if not image.Linked then yield image.Extension, image.ContentType ]
        |> List.distinctBy (fun (extension, _) -> extension.ToLowerInvariant())
    let children =
        [ for extension, contentType in defaults do yield selfClosing "Default" [("Extension", extension); ("ContentType", contentType)]
          for name, contentType in own @ keptOverrides do yield selfClosing "Override" [("PartName", name); ("ContentType", contentType)] ]
    document (element "Types" [("xmlns", Namespaces.contentTypes)] children)

let generateContentTypesPackage hasSettings hasComments hasFootnotes hasEndnotes (headers: HeaderFooterPartInfo array) (footers: HeaderFooterPartInfo array) (imageAssets: ImageAsset array) =
    let children = [
        yield selfClosing "Default" [("Extension", "rels"); ("ContentType", "application/vnd.openxmlformats-package.relationships+xml")]
        yield selfClosing "Default" [("Extension", "xml"); ("ContentType", "application/xml")]
        for extension, contentType in imageAssets |> Array.filter (fun image -> not image.Linked) |> Array.map (fun image -> image.Extension, image.ContentType) |> Array.distinct do
            yield selfClosing "Default" [("Extension", extension); ("ContentType", contentType)]
        yield selfClosing "Override" [("PartName", "/word/document.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")]
        yield selfClosing "Override" [("PartName", "/word/styles.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml")]
        yield selfClosing "Override" [("PartName", "/word/numbering.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml")]
        if hasSettings then yield selfClosing "Override" [("PartName", "/word/settings.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml")]
        for header in headers do
            yield selfClosing "Override" [("PartName", header.PartName); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")]
        for footer in footers do
            yield selfClosing "Override" [("PartName", footer.PartName); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")]
        if hasComments then yield selfClosing "Override" [("PartName", "/word/comments.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml")]
        if hasFootnotes then yield selfClosing "Override" [("PartName", "/word/footnotes.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml")]
        if hasEndnotes then yield selfClosing "Override" [("PartName", "/word/endnotes.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml")]
    ]
    document (element "Types" [("xmlns", Namespaces.contentTypes)] children)

/// Generate _rels/.rels
let generateRootRels () : string =
    let rel = selfClosing "Relationship" [("Id", "rId1"); ("Type", RelationshipTypes.officeDocument); ("Target", "word/document.xml")]
    let root = element "Relationships" [("xmlns", Namespaces.relationshipsPackage)] [rel]
    document root

/// Generate word/_rels/document.xml.rels
/// The writer's own relationship ids are named, not numbered: kept
/// relationships of the opened document have ids like rId1, rId2.
let private documentRelationshipElementsWith (hasNumbering: bool) (hyperlinks: Map<string, string>) (hasComments: bool) (hasFootnotes: bool) (hasEndnotes: bool) =
    let baseRels = [
        yield selfClosing "Relationship" [("Id", freshRelationshipId "rIdStyles"); ("Type", RelationshipTypes.styles); ("Target", "styles.xml")]
        if hasNumbering then
            yield selfClosing "Relationship" [("Id", freshRelationshipId "rIdNumbering"); ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"); ("Target", "numbering.xml")]
        if hasComments then
            yield selfClosing "Relationship" [("Id", freshRelationshipId "rIdComments"); ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"); ("Target", "comments.xml")]
        if hasFootnotes then
            yield selfClosing "Relationship" [("Id", freshRelationshipId "rIdFootnotes"); ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"); ("Target", "footnotes.xml")]
        if hasEndnotes then
            yield selfClosing "Relationship" [("Id", freshRelationshipId "rIdEndnotes"); ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"); ("Target", "endnotes.xml")]
    ]
    let hyperlinkRels =
        hyperlinks
        |> Map.toList
        |> List.map (fun (url, relId) ->
            selfClosing "Relationship" [("Id", relId); ("Type", RelationshipTypes.hyperlink); ("Target", url); ("TargetMode", "External")])
    baseRels @ hyperlinkRels

let private documentRelationshipElements (hyperlinks: Map<string, string>) (hasComments: bool) (hasFootnotes: bool) (hasEndnotes: bool) =
    documentRelationshipElementsWith true hyperlinks hasComments hasFootnotes hasEndnotes

let generateDocumentRels (hyperlinks: Map<string, string>) (hasComments: bool) (hasFootnotes: bool) (hasEndnotes: bool) : string =
    let rels = documentRelationshipElements hyperlinks hasComments hasFootnotes hasEndnotes
    let root = element "Relationships" [("xmlns", Namespaces.relationshipsPackage)] rels
    document root

let private generateImageRelationships (assets: ImageAsset array) =
    assets
    |> Array.map (fun asset ->
        if asset.Linked then
            selfClosing "Relationship" [("Id", asset.RelId); ("Type", RelationshipTypes.image); ("Target", asset.Src); ("TargetMode", "External")]
        else
            selfClosing "Relationship" [("Id", asset.RelId); ("Type", RelationshipTypes.image); ("Target", sprintf "media/%s" asset.FileName)])
    |> Array.toList

let private generateHyperlinkRelationships (hyperlinks: Map<string, string>) =
    hyperlinks
    |> Map.toList
    |> List.map (fun (url, relId) ->
        selfClosing "Relationship" [("Id", relId); ("Type", RelationshipTypes.hyperlink); ("Target", url); ("TargetMode", "External")])

let private generateDocumentRelsWithImages hyperlinks hasComments hasFootnotes hasEndnotes (images: ImageScope) =
    let rels = documentRelationshipElements hyperlinks hasComments hasFootnotes hasEndnotes @ generateImageRelationships images.Assets
    let root = element "Relationships" [("xmlns", Namespaces.relationshipsPackage)] rels
    document root

let generateDocumentRelsPackage hasSettings hyperlinks hasComments hasFootnotes hasEndnotes (headers: HeaderFooterPartInfo array) (footers: HeaderFooterPartInfo array) (images: ImageScope) =
    let headerFooterRels =
        [ if hasSettings then yield selfClosing "Relationship" [("Id", freshRelationshipId "rIdSettings"); ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"); ("Target", "settings.xml")]
          for header in headers -> selfClosing "Relationship" [("Id", header.RelId); ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"); ("Target", header.Target)]
          for footer in footers -> selfClosing "Relationship" [("Id", footer.RelId); ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"); ("Target", footer.Target)] ]
    let rels = documentRelationshipElements hyperlinks hasComments hasFootnotes hasEndnotes @ headerFooterRels @ generateImageRelationships images.Assets
    let root = element "Relationships" [("xmlns", Namespaces.relationshipsPackage)] rels
    document root

/// The relationships part of the main document: the writer's own, then the kept ones.
let private generateDocumentRelsPreserved (kept: PreservedRelationshipInfo list) hasNumbering hasSettings hyperlinks hasComments hasFootnotes hasEndnotes (headers: HeaderFooterPartInfo array) (footers: HeaderFooterPartInfo array) (images: ImageScope) =
    let headerFooterRels =
        [ if hasSettings then yield selfClosing "Relationship" [("Id", freshRelationshipId "rIdSettings"); ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"); ("Target", "settings.xml")]
          for header in headers -> selfClosing "Relationship" [("Id", header.RelId); ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"); ("Target", header.Target)]
          for footer in footers -> selfClosing "Relationship" [("Id", footer.RelId); ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"); ("Target", footer.Target)] ]
    let rels =
        documentRelationshipElementsWith hasNumbering hyperlinks hasComments hasFootnotes hasEndnotes
        @ headerFooterRels
        @ generateImageRelationships images.Assets
        @ (kept |> List.map preservedRelationshipElement)
    document (element "Relationships" [("xmlns", Namespaces.relationshipsPackage)] rels)

let private generateRootRelsPreserved (preserved: PreservedPackage) : string =
    let main = selfClosing "Relationship" [("Id", preserved.DocumentRelationshipId); ("Type", RelationshipTypes.officeDocument); ("Target", "word/document.xml")]
    let kept =
        preserved.RootRelationships
        |> List.filter (fun rel -> rel.Id <> preserved.DocumentRelationshipId)
        |> List.map preservedRelationshipElement
    document (element "Relationships" [("xmlns", Namespaces.relationshipsPackage)] (main :: kept))

let private generatePartRelsWithKept (kept: PreservedRelationshipInfo list) (hyperlinks: Map<string, string>) (images: ImageScope) =
    let rels = generateImageRelationships images.Assets @ generateHyperlinkRelationships hyperlinks @ (kept |> List.map preservedRelationshipElement)
    document (element "Relationships" [("xmlns", Namespaces.relationshipsPackage)] rels)

let private generatePartRels (hyperlinks: Map<string, string>) (images: ImageScope) =
    generatePartRelsWithKept [] hyperlinks images

let generateComments (comments: CommentRef list) : string =
    let children =
        comments
        |> List.map (fun comment ->
            let attrs =
                [("w:id", comment.Id); ("w:author", comment.Author |> Option.defaultValue "RedBat")]
                @ (comment.Date |> Option.map (fun date -> [("w:date", date)]) |> Option.defaultValue [])
            element "w:comment" attrs [
                element "w:p" [] [
                    runElement [] false comment.Text
                ]
            ])

    document (element "w:comments" [("xmlns:w", Namespaces.word)] children)

let generateNotes (rootName: string) (noteElementName: string) (notes: NoteRef list) : string =
    let children =
        notes
        |> List.map (fun note ->
            element noteElementName [("w:id", note.Id)] [
                element "w:p" [] [
                    runElement [] false note.Text
                ]
            ])

    document (element rootName [("xmlns:w", Namespaces.word)] children)

/// A footnote / endnote with its body as blocks (package.Footnotes /
/// package.Endnotes), or with plain text only (a reference run's NoteText, for
/// callers that do not deliver bodies).
type NotePart = {
    Id: string
    /// w:type of a special note: separator, continuationSeparator, continuationNotice.
    Kind: string option
    HasMark: bool
    MarkStyle: string option
    Blocks: obj array
    PlainText: string option
}

type CommentPart = {
    Id: string
    Author: string option
    Date: string option
    Initials: string option
    Blocks: obj array
    PlainText: string option
}

let private notePartsFromPackage (value: obj) : NotePart list =
    optionalArray value
    |> Array.toList
    |> List.choose (fun item ->
        optionalString (item?Id)
        |> Option.map (fun id ->
            { Id = id
              Kind = optionalString (item?Kind)
              HasMark = boolOrDefault true (item?HasMark)
              MarkStyle = optionalString (item?MarkStyle)
              Blocks = optionalArray (item?Blocks)
              PlainText = None }))

let private commentPartsFromPackage (value: obj) : CommentPart list =
    optionalArray value
    |> Array.toList
    |> List.choose (fun item ->
        optionalString (item?Id)
        |> Option.map (fun id ->
            { Id = id
              Author = optionalString (item?Author)
              Date = optionalString (item?Date)
              Initials = optionalString (item?Initials)
              Blocks = optionalArray (item?Blocks)
              PlainText = None }))

/// Bodies from the package first; a note that is only known from a reference
/// run (NoteText) is added as plain text.
let private mergeNoteParts (fromPackage: NotePart list) (fromRuns: NoteRef list) : NotePart list =
    let known = fromPackage |> List.map (fun note -> note.Id) |> Set.ofList
    fromPackage
    @ (fromRuns
       |> List.filter (fun note -> not (Set.contains note.Id known))
       |> List.map (fun note -> { Id = note.Id; Kind = None; HasMark = false; MarkStyle = None; Blocks = [||]; PlainText = Some note.Text }))

let private mergeCommentParts (fromPackage: CommentPart list) (fromRuns: CommentRef list) : CommentPart list =
    let known = fromPackage |> List.map (fun comment -> comment.Id) |> Set.ofList
    fromPackage
    @ (fromRuns
       |> List.filter (fun comment -> not (Set.contains comment.Id known))
       |> List.map (fun comment -> { Id = comment.Id; Author = comment.Author; Date = comment.Date; Initials = None; Blocks = [||]; PlainText = Some comment.Text }))

let private partRootAttributes =
    [("xmlns:w", Namespaces.word); ("xmlns:r", Namespaces.relationships); ("xmlns:m", Namespaces.math)]

/// footnotes.xml / endnotes.xml with the note bodies as blocks. The reader
/// used to hand over the text of a note only, and this wrote one plain run per
/// note, so a table, a hyperlink or bold text in a footnote was gone after a save.
let private generateNotesPart (rootName: string) (noteElementName: string) (markElementName: string) (generateBody: obj array -> string) (notes: NotePart list) : string =
    let markRun (note: NotePart) =
        match note.Kind with
        | Some "separator" -> element "w:r" [] [selfClosing "w:separator" []]
        | Some "continuationSeparator" -> element "w:r" [] [selfClosing "w:continuationSeparator" []]
        | _ ->
            let rPr =
                match note.MarkStyle with
                | Some styleId when styleId <> "" -> [selfClosing "w:rStyle" [("w:val", styleId)]]
                | _ -> [selfClosing "w:vertAlign" [("w:val", "superscript")]]
            element "w:r" [] [element "w:rPr" [] rPr; selfClosing markElementName []]
    // The mark (or the separator line of a separator note) is the first run of
    // the first paragraph.
    let withMark (note: NotePart) (body: string) =
        let emptyProps = "<w:p><w:pPr/>"
        let propsStart = "<w:p><w:pPr>"
        let propsEnd = "</w:pPr>"
        let isSeparator = note.Kind = Some "separator" || note.Kind = Some "continuationSeparator"
        if not note.HasMark && not isSeparator then body
        elif body.StartsWith emptyProps then emptyProps + markRun note + body.Substring emptyProps.Length
        elif body.StartsWith propsStart then
            let index = body.IndexOf propsEnd
            if index < 0 then body
            else body.Substring(0, index + propsEnd.Length) + markRun note + body.Substring(index + propsEnd.Length)
        elif body.StartsWith "<w:p>" then "<w:p>" + markRun note + body.Substring 5
        elif body.StartsWith "<w:p/>" then "<w:p>" + markRun note + "</w:p>" + body.Substring 6
        else body
    let children =
        notes
        |> List.map (fun note ->
            let attrs =
                [ match note.Kind with
                  | Some kind -> yield ("w:type", kind)
                  | None -> ()
                  yield ("w:id", note.Id) ]
            let body =
                match note.PlainText with
                | Some text -> element "w:p" [] [runElement [] false text]
                | None ->
                    let body = generateBody note.Blocks
                    // A note must hold at least one paragraph.
                    withMark note (if body = "" then "<w:p></w:p>" else body)
            element noteElementName attrs [body])
    document (element rootName partRootAttributes children)

let private generateCommentsPart (generateBody: obj array -> string) (comments: CommentPart list) : string =
    let children =
        comments
        |> List.map (fun comment ->
            let attrs =
                [ yield ("w:id", comment.Id)
                  yield ("w:author", comment.Author |> Option.defaultValue "RedBat")
                  match comment.Date with
                  | Some date -> yield ("w:date", date)
                  | None -> ()
                  match comment.Initials with
                  | Some initials -> yield ("w:initials", initials)
                  | None -> () ]
            let body =
                match comment.PlainText with
                | Some text -> element "w:p" [] [runElement [] false text]
                | None ->
                    let body = generateBody comment.Blocks
                    if body = "" then element "w:p" [] [] else body
            element "w:comment" attrs [body])
    document (element "w:comments" partRootAttributes children)

type NumberingInfo = {
    NumId: string
    Level: int
    Format: string
    LevelText: string
    Start: int option
    HasOverride: bool
    /// The level's own indents (px; a negative first-line indent is a hanging
    /// one): what its paragraphs have without stating an indent themselves.
    IndentLeft: float option
    FirstLineIndent: float option
    /// A list id of the file that the file's numbering part does not define
    /// (the reader then has no format and no level text for it). It stays
    /// undefined: a definition made up for it gave the paragraphs numbers
    /// they do not have in Word. The id is still one that is taken.
    Undefined: bool
}

/// One entry per (list id, level) used in the document. The list id, number
/// format, level text and start value all come from the paragraph: without
/// them every numbered paragraph was written with numId 2, which merged
/// separate lists into one decimal list and lost letters / roman numerals.
let private collectNumberingInfoOf (undefinedLists: bool) (items: obj array) =
    items
    |> Array.toList
    |> List.collect collectItemParagraphs
    |> List.choose (fun para ->
        let listStyle = para?ListStyle |> stringOrDefault "none"
        if listStyle = "none" then None
        else
            optionalIntString (para?ListId)
            |> Option.filter (fun numId -> numId <> "0")
            |> Option.map (fun numId ->
                let level = max 0 ((para?ListLevel |> intOrDefault 1) - 1)
                let format = optionalString (para?ListNumberFormat) |> Option.defaultValue (if listStyle = "bullet" then "bullet" else "decimal")
                let levelText = optionalString (para?ListLevelText) |> Option.defaultValue (if format = "bullet" then "\u2022" else "%" + string (level + 1) + ".")
                let start = optionalFloat (para?ListStart) |> Option.map roundToInt
                let restart = boolOrDefault false (para?ListRestart) || boolOrDefault false (para?RestartNumbering)
                // Indents the paragraph has from its list level go into the level
                // definition. A paragraph object that does not tell them apart from
                // its own carries them as direct formatting instead.
                let levelIndent (value: obj) = if boolOrDefault false (para?HasListLevelIndents) then optionalFloat value else None
                { NumId = numId; Level = level; Format = format; LevelText = levelText; Start = start; HasOverride = restart
                  IndentLeft = levelIndent (para?ListLevelIndentLeft); FirstLineIndent = levelIndent (para?ListLevelFirstLineIndent)
                  Undefined = undefinedLists && (optionalString (para?ListNumberFormat)).IsNone && (optionalString (para?ListLevelText)).IsNone }))
    |> List.distinctBy (fun info -> info.NumId, info.Level)

let private collectNumberingInfo (items: obj array) = collectNumberingInfoOf false items

/// Generate minimal styles.xml
type StyleInfo = {
    StyleId: string
    Name: string option
    BasedOn: string option
    Next: string option
    Linked: string option
    IsQuickStyle: bool
    IsDefault: bool
    Aliases: string option
    UiPriority: int option
    SemiHidden: bool
    UnhideWhenUsed: bool
    StyleType: string
    Bold: bool option
    Italic: bool option
    Underline: bool option
    AllCaps: bool option
    SmallCaps: bool option
    Strike: bool option
    VertAlign: string option
    FontSize: float option
    CharacterSpacing: float option
    FontFamily: string option
    Color: string option
    Highlight: string option
    SpacingBefore: float option
    SpacingAfter: float option
    LineSpacing: float option
    LineRule: string option
    Alignment: string option
    IndentLeft: float option
    IndentRight: float option
    FirstLineIndent: float option
    BackgroundColor: string option
    OutlineLevel: int option
    KeepNext: bool option
    KeepLines: bool option
    PageBreakBefore: bool option
    WidowControl: bool option
}

let private emptyStyleInfo styleId styleType = {
    StyleId = styleId
    Name = None
    BasedOn = None
    Next = None
    Linked = None
    IsQuickStyle = false
    IsDefault = false
    Aliases = None
    UiPriority = None
    SemiHidden = false
    UnhideWhenUsed = false
    StyleType = styleType
    Bold = None
    Italic = None
    Underline = None
    AllCaps = None
    SmallCaps = None
    Strike = None
    VertAlign = None
    FontSize = None
    CharacterSpacing = None
    FontFamily = None
    Color = None
    Highlight = None
    SpacingBefore = None
    SpacingAfter = None
    LineSpacing = None
    LineRule = None
    Alignment = None
    IndentLeft = None
    IndentRight = None
    FirstLineIndent = None
    BackgroundColor = None
    OutlineLevel = None
    KeepNext = None
    KeepLines = None
    PageBreakBefore = None
    WidowControl = None
}

let private styleInfoFromObject defaultType (value: obj) =
    let styleId =
        optionalString (value?StyleId)
        |> Option.orElse (optionalString (value?Id))
        |> Option.orElse (optionalString (value?Style))
    styleId
    |> Option.bind (fun id ->
        if id = "" then None
        else
            Some {
                emptyStyleInfo id (optionalString (value?StyleType) |> Option.defaultValue defaultType) with
                    Name = optionalString (value?Name) |> Option.orElse (optionalString (value?StyleName))
                    BasedOn = optionalString (value?BasedOn) |> Option.orElse (optionalString (value?BasedOnStyle))
                    Next = optionalString (value?Next) |> Option.orElse (optionalString (value?NextStyle))
                    Linked = optionalString (value?Linked) |> Option.orElse (optionalString (value?LinkedStyle))
                    IsQuickStyle = optionalBool (value?IsQuickStyle) |> Option.defaultValue false
                    IsDefault = optionalBool (value?IsDefaultStyle) |> Option.orElse (optionalBool (value?IsDefault)) |> Option.defaultValue false
                    Aliases = optionalString (value?StyleAliases) |> Option.orElse (optionalString (value?Aliases))
                    UiPriority = if isNullOrUndefined value?StyleUiPriority then None else Some (unbox<int> value?StyleUiPriority)
                    SemiHidden = optionalBool (value?StyleSemiHidden) |> Option.defaultValue false
                    UnhideWhenUsed = optionalBool (value?StyleUnhideWhenUsed) |> Option.defaultValue false
                    Bold = optionalBool (value?StyleBold)
                    Italic = optionalBool (value?StyleItalic)
                    Underline = optionalBool (value?StyleUnderline)
                    AllCaps = optionalBool (value?StyleAllCaps)
                    SmallCaps = optionalBool (value?StyleSmallCaps)
                    Strike = optionalBool (value?StyleStrike)
                    VertAlign = optionalString (value?StyleVertAlign)
                    FontSize = optionalFloat (value?StyleFontSize)
                    CharacterSpacing = optionalFloat (value?StyleCharacterSpacing)
                    FontFamily = optionalString (value?StyleFontFamily)
                    Color = optionalString (value?StyleColor)
                    Highlight = optionalString (value?StyleHighlight)
                    SpacingBefore = optionalFloat (value?StyleSpacingBefore)
                    SpacingAfter = optionalFloat (value?StyleSpacingAfter)
                    LineSpacing = optionalFloat (value?StyleLineSpacing)
                    LineRule = optionalString (value?StyleLineRule) |> Option.orElse (optionalString (value?StyleLineSpacingRule))
                    Alignment = optionalString (value?StyleAlignment)
                    IndentLeft = optionalFloat (value?StyleIndentLeft)
                    IndentRight = optionalFloat (value?StyleIndentRight)
                    FirstLineIndent = optionalFloat (value?StyleFirstLineIndent)
                    BackgroundColor = optionalString (value?StyleBackgroundColor)
                    OutlineLevel = if isNullOrUndefined value?StyleOutlineLevel then None else Some (unbox<int> value?StyleOutlineLevel)
                    KeepNext = optionalBool (value?StyleKeepNext)
                    KeepLines = optionalBool (value?StyleKeepLines)
                    PageBreakBefore = optionalBool (value?StylePageBreakBefore)
                    WidowControl = optionalBool (value?StyleWidowControl)
            })

let private styleIdFromParagraph (para: obj) =
    optionalString (para?Style)
    |> Option.bind (fun styleId ->
        if styleId = "" || styleId = documentDefaultsStyleId then None
        else Some {
            emptyStyleInfo styleId "paragraph" with
                Name = optionalString (para?StyleName)
                BasedOn = optionalString (para?BasedOnStyle)
                Next = optionalString (para?NextStyle)
                Linked = optionalString (para?LinkedStyle)
                IsQuickStyle = optionalBool (para?IsQuickStyle) |> Option.defaultValue false
                IsDefault = optionalBool (para?IsDefaultStyle) |> Option.orElse (optionalBool (para?IsDefault)) |> Option.defaultValue false
                Aliases = optionalString (para?StyleAliases) |> Option.orElse (optionalString (para?Aliases))
                UiPriority = if isNullOrUndefined para?StyleUiPriority then None else Some (unbox<int> para?StyleUiPriority)
                SemiHidden = optionalBool (para?StyleSemiHidden) |> Option.defaultValue false
                UnhideWhenUsed = optionalBool (para?StyleUnhideWhenUsed) |> Option.defaultValue false
                Bold = optionalBool (para?StyleBold)
                Italic = optionalBool (para?StyleItalic)
                Underline = optionalBool (para?StyleUnderline)
                AllCaps = optionalBool (para?StyleAllCaps)
                SmallCaps = optionalBool (para?StyleSmallCaps)
                Strike = optionalBool (para?StyleStrike)
                VertAlign = optionalString (para?StyleVertAlign)
                FontSize = optionalFloat (para?StyleFontSize)
                CharacterSpacing = optionalFloat (para?StyleCharacterSpacing)
                FontFamily = optionalString (para?StyleFontFamily)
                Color = optionalString (para?StyleColor)
                Highlight = optionalString (para?StyleHighlight)
                SpacingBefore = optionalFloat (para?StyleSpacingBefore)
                SpacingAfter = optionalFloat (para?StyleSpacingAfter)
                LineSpacing = optionalFloat (para?StyleLineSpacing)
                LineRule = optionalString (para?StyleLineRule) |> Option.orElse (optionalString (para?StyleLineSpacingRule))
                Alignment = optionalString (para?StyleAlignment)
                IndentLeft = optionalFloat (para?StyleIndentLeft)
                IndentRight = optionalFloat (para?StyleIndentRight)
                FirstLineIndent = optionalFloat (para?StyleFirstLineIndent)
                BackgroundColor = optionalString (para?StyleBackgroundColor)
                OutlineLevel = if isNullOrUndefined para?StyleOutlineLevel then None else Some (unbox<int> para?StyleOutlineLevel)
                KeepNext = optionalBool (para?StyleKeepNext)
                KeepLines = optionalBool (para?StyleKeepLines)
                PageBreakBefore = optionalBool (para?StylePageBreakBefore)
                WidowControl = optionalBool (para?StyleWidowControl)
        })

let private styleIdFromRun (run: obj) =
    optionalString (run?Style)
    |> Option.orElse (optionalString (run?RunStyle))
    |> Option.orElse (optionalString (run?CharacterStyle))
    |> Option.bind (fun styleId ->
        if styleId = "" then None
        else Some {
            emptyStyleInfo styleId "character" with
                Name = optionalString (run?StyleName) |> Option.orElse (optionalString (run?RunStyleName)) |> Option.orElse (optionalString (run?CharacterStyleName))
                BasedOn = optionalString (run?BasedOnStyle)
                Next = optionalString (run?NextStyle)
                Linked = optionalString (run?LinkedStyle)
                IsQuickStyle = optionalBool (run?IsQuickStyle) |> Option.defaultValue false
                IsDefault = optionalBool (run?IsDefaultStyle) |> Option.orElse (optionalBool (run?IsDefault)) |> Option.defaultValue false
                Aliases = optionalString (run?StyleAliases) |> Option.orElse (optionalString (run?Aliases))
                UiPriority = if isNullOrUndefined run?StyleUiPriority then None else Some (unbox<int> run?StyleUiPriority)
                SemiHidden = optionalBool (run?StyleSemiHidden) |> Option.defaultValue false
                UnhideWhenUsed = optionalBool (run?StyleUnhideWhenUsed) |> Option.defaultValue false
                Bold = optionalBool (run?StyleBold)
                Italic = optionalBool (run?StyleItalic)
                Underline = optionalBool (run?StyleUnderline)
                AllCaps = optionalBool (run?StyleAllCaps)
                SmallCaps = optionalBool (run?StyleSmallCaps)
                Strike = optionalBool (run?StyleStrike)
                VertAlign = optionalString (run?StyleVertAlign)
                FontSize = optionalFloat (run?StyleFontSize)
                CharacterSpacing = optionalFloat (run?StyleCharacterSpacing)
                FontFamily = optionalString (run?StyleFontFamily)
                Color = optionalString (run?StyleColor)
                Highlight = optionalString (run?StyleHighlight)
        })

let private runStylesFromParagraph (para: obj) =
    if isNullOrUndefined para?Runs then []
    else para?Runs |> unbox<obj array> |> Array.toList |> List.choose styleIdFromRun

let private collectItemParagraphStyles (item: obj) =
    collectItemParagraphs item
    |> List.collect (fun para -> (styleIdFromParagraph para |> Option.toList) @ runStylesFromParagraph para)

let private collectStyles (items: obj array) =
    items
    |> Array.toList
    |> List.collect collectItemParagraphStyles
    |> List.distinctBy (fun style -> style.StyleType, style.StyleId)

let private collectPackageStyles (styles: obj array) =
    styles
    |> Array.toList
    |> List.choose (styleInfoFromObject "paragraph")

/// Style run properties, in CT_RPr order (see orderedRunProps). A style only
/// states what it sets, so absent properties are left out rather than written off.
let private styleRunProps style =
    let styleToggle name value =
        match value with
        | Some isOn -> [toggle name isOn]
        | None -> []
    [
        match style.FontFamily with
        | Some family when family <> "" -> yield selfClosing "w:rFonts" [("w:ascii", family); ("w:hAnsi", family); ("w:cs", family)]
        | _ -> ()
        yield! styleToggle "w:b" style.Bold
        yield! styleToggle "w:i" style.Italic
        yield! styleToggle "w:caps" style.AllCaps
        yield! styleToggle "w:smallCaps" style.SmallCaps
        yield! styleToggle "w:strike" style.Strike
        match InverseConversion.cssColorToWord style.Color with
        | Some color -> yield selfClosing "w:color" [("w:val", color)]
        | None -> ()
        match style.CharacterSpacing with
        | Some spacing when abs spacing >= 0.001 -> yield selfClosing "w:spacing" [("w:val", InverseConversion.pixelsToTwips spacing)]
        | _ -> ()
        match style.FontSize with
        | Some size ->
            let value = InverseConversion.pixelsToHalfPoints size
            yield selfClosing "w:sz" [("w:val", value)]
            yield selfClosing "w:szCs" [("w:val", value)]
        | _ -> ()
        match style.Highlight with
        | Some color when color <> "" -> yield selfClosing "w:highlight" [("w:val", color)]
        | _ -> ()
        match style.Underline with
        | Some isOn -> yield selfClosing "w:u" [("w:val", (if isOn then "single" else "none"))]
        | None -> ()
        match style.VertAlign with | Some value when value = "superscript" || value = "subscript" -> yield selfClosing "w:vertAlign" [("w:val", value)] | _ -> ()
    ]

/// Style paragraph properties in CT_PPr order: keepNext, keepLines,
/// pageBreakBefore, widowControl, shd, spacing, ind, jc, outlineLvl.
let private styleParagraphProps style =
    let styleToggle name value =
        match value with
        | Some isOn -> [toggle name isOn]
        | None -> []
    [
        yield! styleToggle "w:keepNext" style.KeepNext
        yield! styleToggle "w:keepLines" style.KeepLines
        yield! styleToggle "w:pageBreakBefore" style.PageBreakBefore
        yield! styleToggle "w:widowControl" style.WidowControl
        match InverseConversion.cssColorToWord style.BackgroundColor with
        | Some color -> yield selfClosing "w:shd" [("w:val", "clear"); ("w:color", "auto"); ("w:fill", color)]
        | None -> ()
        let spacingAttrs =
            [
                match style.SpacingBefore with | Some value -> yield ("w:before", InverseConversion.pointsToTwips value) | None -> ()
                match style.SpacingAfter with | Some value -> yield ("w:after", InverseConversion.pointsToTwips value) | None -> ()
                // Styles.styleToObj reads exact/at-least as px (twips / 15) and auto as a multiple (line / 240).
                match style.LineSpacing, style.LineRule with
                | Some value, Some rule when rule = "exact" || rule = "atLeast" ->
                    yield ("w:line", InverseConversion.pixelsToTwips value)
                    yield ("w:lineRule", rule)
                | Some value, _ -> yield ("w:line", InverseConversion.lineSpacingToWord value)
                | None, _ -> ()
            ]
        if not spacingAttrs.IsEmpty then yield selfClosing "w:spacing" spacingAttrs
        let indentAttrs =
            [
                match style.IndentLeft with | Some value -> yield ("w:left", InverseConversion.pixelsToTwips value) | None -> ()
                match style.IndentRight with | Some value -> yield ("w:right", InverseConversion.pixelsToTwips value) | None -> ()
                match style.FirstLineIndent with
                | Some value when value < 0.0 -> yield ("w:hanging", InverseConversion.pixelsToTwips -value)
                | Some value -> yield ("w:firstLine", InverseConversion.pixelsToTwips value)
                | None -> ()
            ]
        if not indentAttrs.IsEmpty then yield selfClosing "w:ind" indentAttrs
        match style.Alignment with
        | Some value when value <> "" -> yield selfClosing "w:jc" [("w:val", InverseConversion.alignmentToWord value)]
        | _ -> ()
        match style.OutlineLevel with | Some value -> yield selfClosing "w:outlineLvl" [("w:val", string value)] | None -> ()
    ]

let private defaultStyleId = documentDefaultsStyleId

let private docDefaultsFromStyle (style: StyleInfo option) =
    let defaultRunStyle =
        style
        |> Option.defaultValue { emptyStyleInfo defaultStyleId "defaults" with FontFamily = Some "Calibri"; FontSize = Some defaultFontSizePixels }
    let runProps = styleRunProps defaultRunStyle
    let rPr = element "w:rPr" [] runProps
    let rPrDefault = element "w:rPrDefault" [] [rPr]
    let pPr = style |> Option.map styleParagraphProps |> Option.defaultValue []
    let pPrDefault = if pPr.IsEmpty then element "w:pPrDefault" [] [] else element "w:pPrDefault" [] [element "w:pPr" [] pPr]
    element "w:docDefaults" [] [rPrDefault; pPrDefault]

let private styleElement (style: StyleInfo) : string =
            let children = [
                yield selfClosing "w:name" [("w:val", style.Name |> Option.defaultValue style.StyleId)]
                match style.Aliases with
                | Some aliases when aliases <> "" -> yield selfClosing "w:aliases" [("w:val", aliases)]
                | _ -> ()
                match style.BasedOn with
                | Some basedOn when basedOn <> "" -> yield selfClosing "w:basedOn" [("w:val", basedOn)]
                | _ -> ()
                match style.Next with
                | Some next when next <> "" -> yield selfClosing "w:next" [("w:val", next)]
                | _ -> ()
                match style.Linked with
                | Some linked when linked <> "" -> yield selfClosing "w:link" [("w:val", linked)]
                | _ -> ()
                match style.UiPriority with
                | Some priority -> yield selfClosing "w:uiPriority" [("w:val", string priority)]
                | None -> ()
                if style.SemiHidden then yield selfClosing "w:semiHidden" []
                if style.UnhideWhenUsed then yield selfClosing "w:unhideWhenUsed" []
                if style.IsQuickStyle then yield selfClosing "w:qFormat" []
                let pPr = styleParagraphProps style
                if not pPr.IsEmpty then yield element "w:pPr" [] pPr
                let rPr = styleRunProps style
                if not rPr.IsEmpty then yield element "w:rPr" [] rPr
            ]
            let attrs = [("w:type", style.StyleType); ("w:styleId", style.StyleId)] @ (if style.IsDefault then [("w:default", "1")] else [])
            element "w:style" attrs children

let generateDefaultStylesWithStyles (styles: StyleInfo list) : string =
    let docDefaults = styles |> List.tryFind (fun style -> style.StyleType = "defaults") |> docDefaultsFromStyle
    let customStyles =
        styles
        |> List.filter (fun style -> style.StyleType <> "defaults")
        |> List.map styleElement
    let root = element "w:styles" [("xmlns:w", Namespaces.word); ("xmlns:r", Namespaces.relationships)] (docDefaults :: customStyles)
    document root

let private rootStartTag (name: string) (attrs: Map<string, string>) =
    let attributes = attrs |> Map.toList |> List.map (fun (key, value) -> sprintf " %s=\"%s\"" key (escapeXmlAttribute value)) |> String.concat ""
    "<" + name + attributes + ">"

/// styles.xml of a document that was opened from a file: the original part
/// with the styles the editor changed (or added) replaced by their current
/// definition. Everything the editor has no model for stays as it was: table
/// styles, latent styles, numbering-linked styles, and whatever an unchanged
/// style says beyond the properties the editor knows (the style part used to be
/// rewritten from the editor's catalog, which reduced every style to those).
/// `changed` holds "type:id" keys; None = the caller does not know, all
/// catalog styles are taken as changed.
let private mergeStyles (originalXml: string) (changed: Set<string> option) (styles: StyleInfo list) : string =
    match XmlDsl.parse originalXml with
    | XmlDsl.Element (rootName, rootAttrs, children) ->
        let isChanged (style: StyleInfo) =
            match changed with
            | Some keys -> Set.contains (style.StyleType + ":" + style.StyleId) keys
            | None -> true
        let localName node = XmlDsl.Xml.name node |> Option.map XmlDsl.Ns.stripNs
        let styleKey node =
            let attr name = XmlDsl.Xml.attr ("w:" + name) node |> Option.orElse (XmlDsl.Xml.attr name node)
            (attr "type" |> Option.defaultValue "paragraph"), (attr "styleId" |> Option.defaultValue "")
        let present =
            children
            |> List.choose (fun child -> if localName child = Some "style" then Some (styleKey child) else None)
            |> Set.ofList
        let hasDefaults = children |> List.exists (fun child -> localName child = Some "docDefaults")
        let defaults = styles |> List.tryFind (fun style -> style.StyleType = "defaults")
        // Nothing was changed and nothing is new: the part as it was read (writing
        // it out again from its nodes reordered the attributes of every style).
        let untouched =
            (changed |> Option.exists Set.isEmpty)
            && (hasDefaults || defaults.IsNone)
            && (styles |> List.forall (fun style -> style.StyleType = "defaults" || Set.contains (style.StyleType, style.StyleId) present))
        if untouched then originalXml else
        let body =
            [ // No docDefaults in the original: the catalog's, if it has any.
              if not hasDefaults then
                  match defaults with
                  | Some _ -> yield docDefaultsFromStyle defaults
                  | None -> ()
              for child in children do
                  match localName child with
                  | Some "docDefaults" ->
                      match defaults with
                      | Some style when isChanged style -> yield docDefaultsFromStyle defaults
                      | _ -> yield DocxReader.Docx.xmlToString child
                  | Some "style" ->
                      let styleType, styleId = styleKey child
                      match styles |> List.tryFind (fun style -> style.StyleType = styleType && style.StyleId = styleId) with
                      | Some style when isChanged style -> yield styleElement style
                      | _ -> yield DocxReader.Docx.xmlToString child
                  | Some _ -> yield DocxReader.Docx.xmlToString child
                  | None -> ()
              for style in styles do
                  if style.StyleType <> "defaults" && not (Set.contains (style.StyleType, style.StyleId) present) then
                      yield styleElement style ]
        document (rootStartTag rootName rootAttrs + String.concat "" body + "</" + rootName + ">")
    | _ -> generateDefaultStylesWithStyles styles

let generateDefaultStyles () : string =
    generateDefaultStylesWithStyles []

/// Bullets from Word are usually private-use glyphs of a symbol font; the
/// editor model has no slot for the level font, so name the font the glyph
/// belongs to or Word shows an empty box.
let private bulletFont (levelText: string) =
    if levelText = "o" then Some "Courier New"
    elif levelText.Length = 1 && int levelText.[0] >= 0xF000 && int levelText.[0] <= 0xF0FF then
        if int levelText.[0] = 0xF0B7 then Some "Symbol" else Some "Wingdings"
    else None

/// `withIndent` is for the built-in lists only. Paragraphs of an imported
/// list already carry the level indent as direct formatting (the reader folds
/// it into IndentLeft), so repeating a guessed one here would add an indent to
/// lists whose definition has none.
let private numberingLevelWithIndents (withIndent: bool) (indentLeft: float option) (firstLineIndent: float option) (level: int) (start: int option) (format: string) (levelText: string) =
    let levelIndents =
        [ match indentLeft with
          | Some value -> yield ("w:left", InverseConversion.pixelsToTwips value)
          | None -> ()
          match firstLineIndent with
          | Some value when value < 0.0 -> yield ("w:hanging", InverseConversion.pixelsToTwips -value)
          | Some value -> yield ("w:firstLine", InverseConversion.pixelsToTwips value)
          | None -> () ]
    element "w:lvl" [("w:ilvl", string level)] [
        match start with
        | Some value -> yield selfClosing "w:start" [("w:val", string value)]
        | None -> ()
        yield selfClosing "w:numFmt" [("w:val", format)]
        yield selfClosing "w:lvlText" [("w:val", levelText)]
        yield selfClosing "w:lvlJc" [("w:val", "left")]
        if withIndent then
            yield element "w:pPr" [] [
                selfClosing "w:ind" [("w:left", string (720 * (level + 1))); ("w:hanging", "360")]
            ]
        elif not levelIndents.IsEmpty then
            yield element "w:pPr" [] [selfClosing "w:ind" levelIndents]
        match (if format = "bullet" then bulletFont levelText else None) with
        | Some font -> yield element "w:rPr" [] [selfClosing "w:rFonts" [("w:ascii", font); ("w:hAnsi", font); ("w:hint", "default")]]
        | None -> ()
    ]

let private numberingLevel (withIndent: bool) (level: int) (start: int option) (format: string) (levelText: string) =
    numberingLevelWithIndents withIndent None None level start format levelText

let generateNumbering (customInfos: NumberingInfo list) : string =
    let usedIds =
        customInfos
        |> List.choose (fun info -> match System.Int32.TryParse info.NumId with | true, value -> Some value | _ -> None)
    let nextFree = (if usedIds.IsEmpty then 0 else List.max usedIds) + 1
    // Defaults for lists made in the editor, on ids no imported list uses.
    let bulletId, numberedId =
        if usedIds |> List.exists (fun id -> id = 1 || id = 2) then string nextFree, string (nextFree + 1) else "1", "2"
    defaultListNumIds <- (bulletId, numberedId)
    let defaultLevels format text =
        [0..8] |> List.map (fun level -> numberingLevel true level (Some 1) format (if format = "bullet" then text else "%" + string (level + 1) + "."))

    let groups = customInfos |> List.filter (fun info -> not info.Undefined) |> List.groupBy (fun info -> info.NumId)
    let abstracts =
        [ yield element "w:abstractNum" [("w:abstractNumId", bulletId)] (defaultLevels "bullet" "\u2022")
          yield element "w:abstractNum" [("w:abstractNumId", numberedId)] (defaultLevels "decimal" "")
          for numId, infos in groups do
              let levels =
                  infos
                  |> List.sortBy (fun info -> info.Level)
                  |> List.map (fun info -> numberingLevelWithIndents false info.IndentLeft info.FirstLineIndent info.Level info.Start info.Format info.LevelText)
              yield element "w:abstractNum" [("w:abstractNumId", numId)] levels ]
    let nums =
        [ yield element "w:num" [("w:numId", bulletId)] [selfClosing "w:abstractNumId" [("w:val", bulletId)]]
          yield element "w:num" [("w:numId", numberedId)] [selfClosing "w:abstractNumId" [("w:val", numberedId)]]
          for numId, infos in groups do
              let overrides =
                  infos
                  |> List.filter (fun info -> info.HasOverride)
                  |> List.sortBy (fun info -> info.Level)
                  |> List.map (fun info ->
                      element "w:lvlOverride" [("w:ilvl", string info.Level)] [
                          selfClosing "w:startOverride" [("w:val", string (info.Start |> Option.defaultValue 1))]
                      ])
              yield element "w:num" [("w:numId", numId)] (selfClosing "w:abstractNumId" [("w:val", numId)] :: overrides) ]

    document (element "w:numbering" [("xmlns:w", Namespaces.word)] (abstracts @ nums))

/// The numbering part as it was read, with definitions added for the lists
/// that were made in the editor (and, when a paragraph written from the model
/// has a list without an id, the two default lists). The definitions of the
/// file keep what the model does not hold of them: the bullet fonts, the
/// alignment and the tab of the number, picture bullets, the levels no
/// paragraph uses yet, the links to numbering styles. None when the part is
/// not what this expects (it is then generated as for a new document).
let private mergeNumbering (original: string) (customInfos: NumberingInfo list) (needsDefaults: bool) : string option =
    let closing = "</w:numbering>"
    // (A part without a list is an empty element.)
    let original =
        let trimmed = original.TrimEnd()
        if original.IndexOf(closing, System.StringComparison.Ordinal) < 0 && original.IndexOf("<w:numbering", System.StringComparison.Ordinal) >= 0 && trimmed.EndsWith "/>"
           && (needsDefaults || not customInfos.IsEmpty) then
            trimmed.Substring(0, trimmed.Length - 2) + ">" + closing
        else original
    let closeAt = original.LastIndexOf(closing, System.StringComparison.Ordinal)
    if closeAt < 0 then (if not needsDefaults && customInfos.IsEmpty then Some original else None)
    else
        let known = [ for m in numElementPattern.Matches original -> m.Groups.[1].Value ] |> Set.ofList
        let highest =
            [ for m in numberAttributePattern.Matches original do
                let text = if m.Groups.[1].Success then m.Groups.[1].Value else m.Groups.[2].Value
                match System.Int32.TryParse text with
                | true, value -> yield value
                | _ -> () ]
            |> List.fold max 0
        let unknown = customInfos |> List.filter (fun info -> not (Set.contains info.NumId known))
        let added = unknown |> List.filter (fun info -> not info.Undefined) |> List.groupBy (fun info -> info.NumId)
        let addedIds = unknown |> List.groupBy (fun info -> info.NumId) |> List.choose (fun (numId, _) -> match System.Int32.TryParse numId with | true, value -> Some value | _ -> None)
        let mutable nextFree = (List.fold max highest addedIds) + 1
        let fresh () =
            let id = string nextFree
            nextFree <- nextFree + 1
            id
        let abstracts = ResizeArray<string>()
        let nums = ResizeArray<string>()
        if needsDefaults then
            let bulletId, numberedId = fresh (), fresh ()
            defaultListNumIds <- (bulletId, numberedId)
            let defaultLevels format text =
                [0..8] |> List.map (fun level -> numberingLevel true level (Some 1) format (if format = "bullet" then text else "%" + string (level + 1) + "."))
            abstracts.Add (element "w:abstractNum" [("w:abstractNumId", bulletId)] (defaultLevels "bullet" "\u2022"))
            abstracts.Add (element "w:abstractNum" [("w:abstractNumId", numberedId)] (defaultLevels "decimal" ""))
            nums.Add (element "w:num" [("w:numId", bulletId)] [selfClosing "w:abstractNumId" [("w:val", bulletId)]])
            nums.Add (element "w:num" [("w:numId", numberedId)] [selfClosing "w:abstractNumId" [("w:val", numberedId)]])
        for numId, infos in added do
            let abstractId = fresh ()
            let levels =
                infos
                |> List.sortBy (fun info -> info.Level)
                |> List.map (fun info -> numberingLevelWithIndents false info.IndentLeft info.FirstLineIndent info.Level info.Start info.Format info.LevelText)
            let overrides =
                infos
                |> List.filter (fun info -> info.HasOverride)
                |> List.sortBy (fun info -> info.Level)
                |> List.map (fun info ->
                    element "w:lvlOverride" [("w:ilvl", string info.Level)] [
                        selfClosing "w:startOverride" [("w:val", string (info.Start |> Option.defaultValue 1))]
                    ])
            abstracts.Add (element "w:abstractNum" [("w:abstractNumId", abstractId)] levels)
            nums.Add (element "w:num" [("w:numId", numId)] (selfClosing "w:abstractNumId" [("w:val", abstractId)] :: overrides))
        if abstracts.Count = 0 then Some original
        else
            // CT_Numbering: the abstract definitions, then the lists, then numIdMacAtCleanup.
            let cleanupAt = original.IndexOf("<w:numIdMacAtCleanup", System.StringComparison.Ordinal)
            let numsAt = if cleanupAt >= 0 then cleanupAt else closeAt
            let firstNum = firstNumPattern.Match original
            let abstractsAt = if firstNum.Success then firstNum.Index else numsAt
            Some (original.Substring(0, abstractsAt)
                  + String.concat "" abstracts
                  + original.Substring(abstractsAt, numsAt - abstractsAt)
                  + String.concat "" nums
                  + original.Substring numsAt)

/// The numbering part as it was read with the levels the editor changed
/// ("list id/level") written anew: in the list's own abstract definition (a
/// copy of it, made for the list, when other lists share the definition), the
/// other levels, lists and definitions staying the text they were. None when
/// a list or its definition is not found in the text.
let private patchNumberingLevels (original: string) (customInfos: NumberingInfo list) (changed: string array) : string option =
    let mutable text = original
    let mutable ok = true
    let elementRange (from: int) (name: string) =
        // (The element that starts at `from`: where it ends.)
        let tagEnd = tagEndIndex text from
        if tagEnd < 0 then -1
        elif text.[tagEnd - 1] = '/' then tagEnd + 1
        else
            match childElementRanges text (tagEnd + 1) with
            | Some children ->
                let close = text.IndexOf('>', children.CloseAt)
                if close < 0 then -1 else close + 1
            | None -> -1
    let findStart (pattern: string) =
        let m = System.Text.RegularExpressions.Regex.Match(text, pattern)
        if m.Success then m.Index else -1
    let escape (value: string) = System.Text.RegularExpressions.Regex.Escape value
    let highest () =
        [ for m in numberAttributePattern.Matches text do
            let value = if m.Groups.[1].Success then m.Groups.[1].Value else m.Groups.[2].Value
            match System.Int32.TryParse value with
            | true, number -> yield number
            | _ -> () ]
        |> List.fold max 0
    for key in changed do
        let slash = key.LastIndexOf '/'
        let numId = if slash > 0 then key.Substring(0, slash) else key
        // (The editor counts levels from one, the file from zero.)
        let level = if slash > 0 then (match System.Int32.TryParse(key.Substring(slash + 1)) with | true, value -> value - 1 | _ -> -1) else -1
        match customInfos |> List.tryFind (fun info -> info.NumId = numId && info.Level = level && not info.Undefined) with
        | None -> ok <- false
        | Some info when ok ->
            let numAt = findStart ("<w:num(?:\\s[^>]*)?\\sw:numId=\"" + escape numId + "\"")
            let numEnd = if numAt < 0 then -1 else elementRange numAt "w:num"
            if numAt < 0 || numEnd < 0 then ok <- false
            else
                let num = text.Substring(numAt, numEnd - numAt)
                let reference = System.Text.RegularExpressions.Regex.Match(num, "<w:abstractNumId w:val=\"(\\d+)\"\\s*/>")
                if not reference.Success then ok <- false
                else
                    let abstractId = reference.Groups.[1].Value
                    let abstractPattern = "<w:abstractNum(?:\\s[^>]*)?\\sw:abstractNumId=\"" + escape abstractId + "\""
                    let abstractAt = findStart abstractPattern
                    let abstractEnd = if abstractAt < 0 then -1 else elementRange abstractAt "w:abstractNum"
                    if abstractAt < 0 || abstractEnd < 0 then ok <- false
                    else
                        let shared = System.Text.RegularExpressions.Regex.Matches(text, "<w:abstractNumId w:val=\"" + escape abstractId + "\"\\s*/>").Count > 1
                        let definition = text.Substring(abstractAt, abstractEnd - abstractAt)
                        let ownId = if shared then string (highest () + 1) else abstractId
                        // The level written anew, in the place of the one that is there.
                        let levelXml = numberingLevelWithIndents false info.IndentLeft info.FirstLineIndent info.Level info.Start info.Format info.LevelText
                        let withLevel (definition: string) =
                            let m = System.Text.RegularExpressions.Regex.Match(definition, "<w:lvl(?:\\s[^>]*)?\\sw:ilvl=\"" + string level + "\"[^>]*>")
                            if m.Success then
                                let close = definition.IndexOf("</w:lvl>", m.Index, System.StringComparison.Ordinal)
                                if close < 0 then None else Some (definition.Substring(0, m.Index) + levelXml + definition.Substring(close + 8))
                            else
                                let close = definition.LastIndexOf("</w:abstractNum>", System.StringComparison.Ordinal)
                                if close < 0 then None else Some (definition.Substring(0, close) + levelXml + definition.Substring close)
                        match withLevel definition with
                        | None -> ok <- false
                        | Some definition ->
                            let definition =
                                if not shared then definition
                                else
                                    // A definition of its own for this list: with an id and a signature of its own.
                                    let attribute = "w:abstractNumId=\"" + abstractId + "\""
                                    let at = definition.IndexOf(attribute, System.StringComparison.Ordinal)
                                    let renamed = definition.Substring(0, at) + "w:abstractNumId=\"" + ownId + "\"" + definition.Substring(at + attribute.Length)
                                    System.Text.RegularExpressions.Regex.Replace(renamed, "<w:nsid [^>]*/>|<w:tmpl [^>]*/>", "")
                            // The list: its definition, and no override of the level but the editor's own start.
                            let num =
                                let repointed = if shared then num.Replace(reference.Value, "<w:abstractNumId w:val=\"" + ownId + "\"/>") else num
                                let without = System.Text.RegularExpressions.Regex.Replace(repointed, "<w:lvlOverride w:ilvl=\"" + string level + "\"(?:/>|>[\\s\\S]*?</w:lvlOverride>)", "")
                                if not info.HasOverride then without
                                else
                                    let startOverride = element "w:lvlOverride" [("w:ilvl", string level)] [ selfClosing "w:startOverride" [("w:val", string (info.Start |> Option.defaultValue 1))] ]
                                    if without.EndsWith "/>" then without.Substring(0, without.Length - 2) + ">" + startOverride + "</w:num>"
                                    else
                                        let close = without.LastIndexOf("</w:num>", System.StringComparison.Ordinal)
                                        without.Substring(0, close) + startOverride + without.Substring close
                            // (The list comes after the definitions: it is replaced first.)
                            if numAt > abstractAt then
                                text <- text.Substring(0, numAt) + num + text.Substring numEnd
                                text <-
                                    if shared then text.Substring(0, abstractEnd) + definition + text.Substring abstractEnd
                                    else text.Substring(0, abstractAt) + definition + text.Substring abstractEnd
                            else ok <- false
        | Some _ -> ()
    if ok then Some text else None

let private generateSettings (pageSetup: PageSetupInfo option) =
    let children = [
        match pageSetup with
        | Some setup when setup.EvenAndOddHeaders -> yield selfClosing "w:evenAndOddHeaders" []
        | _ -> ()
    ]
    document (element "w:settings" [("xmlns:w", Namespaces.word)] children)

/// Children of w:settings that come before w:evenAndOddHeaders in CT_Settings.
let private settingsBeforeEvenAndOdd =
    [ "writeProtection"; "view"; "zoom"; "removePersonalInformation"; "removeDateAndTime"; "doNotDisplayPageBoundaries"
      "displayBackgroundShape"; "printPostScriptOverText"; "printFractionalCharacterWidth"; "printFormsData"; "embedTrueTypeFonts"
      "embedSystemFonts"; "saveSubsetFonts"; "saveFormsData"; "mirrorMargins"; "alignBordersAndEdges"; "bordersDoNotSurroundHeader"
      "bordersDoNotSurroundFooter"; "gutterAtTop"; "hideSpellingErrors"; "hideGrammaticalErrors"; "activeWritingStyle"; "proofState"
      "formsDesign"; "attachedTemplate"; "linkStyles"; "stylePaneFormatFilter"; "stylePaneSortMethod"; "documentType"; "mailMerge"
      "revisionView"; "trackRevisions"; "doNotTrackMoves"; "doNotTrackFormatting"; "documentProtection"; "autoFormatOverride"
      "styleLockTheme"; "styleLockQFSet"; "defaultTabStop"; "autoHyphenation"; "consecutiveHyphenLimit"; "hyphenationZone"
      "doNotHyphenateCaps"; "showEnvelope"; "summaryLength"; "clickAndTypeStyle"; "defaultTableStyle" ]
    |> Set.ofList

/// The original settings part with the one setting the editor owns, even and
/// odd headers, brought up to date. Everything else (compatibility options,
/// default tab stop, track-revisions flag, footnote properties, rsids, ...)
/// stays: the part used to be dropped, or replaced by one holding this flag only.
let private patchSettings (originalXml: string) (pageSetup: PageSetupInfo option) : string =
    match XmlDsl.parse originalXml with
    | XmlDsl.Element (rootName, rootAttrs, children) ->
        let localName node = XmlDsl.Xml.name node |> Option.map XmlDsl.Ns.stripNs
        let others = children |> List.filter (fun child -> localName child <> Some "evenAndOddHeaders")
        let evenAndOdd = pageSetup |> Option.exists (fun setup -> setup.EvenAndOddHeaders)
        // The setting is what the part says already: the part as it was read.
        let hadEvenAndOdd =
            children
            |> List.exists (fun child ->
                localName child = Some "evenAndOddHeaders"
                && (match XmlDsl.Xml.attr "w:val" child with
                    | Some "0" | Some "false" | Some "off" -> false
                    | _ -> true))
        if hadEvenAndOdd = evenAndOdd then originalXml else
        let insertAt =
            others
            |> List.mapi (fun index child -> index, child)
            |> List.filter (fun (_, child) -> localName child |> Option.exists (fun name -> Set.contains name settingsBeforeEvenAndOdd))
            |> List.map fst
            |> function
               | [] -> 0
               | indexes -> List.max indexes + 1
        let prefix = if rootName.Contains ":" then rootName.Substring(0, rootName.IndexOf ':' + 1) else ""
        let serialised = others |> List.map DocxReader.Docx.xmlToString
        let body =
            if evenAndOdd then (List.truncate insertAt serialised) @ [ selfClosing (prefix + "evenAndOddHeaders") [] ] @ (List.skip insertAt serialised)
            else serialised
        document (rootStartTag rootName rootAttrs + String.concat "" body + "</" + rootName + ">")
    | _ -> generateSettings pageSetup

// ---- relationship parts and [Content_Types].xml as they were read -------------
// The writer makes these parts anew for every save: its own order, its own
// ids for the parts nothing refers to by id (the styles, the lists, the
// settings, the notes, the comments), absolute targets. A part that was read
// and says what the new one says is written back as the text it was.

let private relationshipTagPattern = System.Text.RegularExpressions.Regex("<Relationship\\s[^>]*>")
let private attributeOfTag (tag: string) (name: string) : string option =
    let m = System.Text.RegularExpressions.Regex.Match(tag, "\\s" + name + "\\s*=\\s*(?:\"([^\"]*)\"|'([^']*)')")
    if not m.Success then None
    elif m.Groups.[1].Success then Some m.Groups.[1].Value
    else Some m.Groups.[2].Value

/// Relationship types of the parts a document has one of, which nothing refers to by id.
let private isImplicitRelationshipType (relType: string) =
    [ "/styles"; "/numbering"; "/settings"; "/comments"; "/footnotes"; "/endnotes" ] |> List.exists (fun suffix -> relType.EndsWith suffix)

/// What a relationships part says: one entry per relationship (its id - but
/// for the parts nothing refers to by id -, type, mode and target, a target
/// in the package as its path), sorted. None when it has something else.
let private relationshipEntries (ownerDirectory: string) (xml: string) : string list option =
    let tags = [ for m in relationshipTagPattern.Matches xml -> m.Value ]
    let entries =
        tags
        |> List.map (fun tag ->
            match attributeOfTag tag "Id", attributeOfTag tag "Type", attributeOfTag tag "Target" with
            | Some id, Some relType, Some target ->
                let isExternal = attributeOfTag tag "TargetMode" = Some "External"
                let target = if isExternal then target else ZipInterop.resolveTarget ownerDirectory target
                Some ((if isImplicitRelationshipType relType then "" else id) + "\n" + relType + "\n" + (if isExternal then "external" else "internal") + "\n" + target)
            | _ -> None)
    if entries |> List.exists Option.isNone then None else Some (entries |> List.choose id |> List.sort)

/// The directory of the part a relationships part belongs to ("word/_rels/document.xml.rels" -> "word").
let private relationshipsOwnerDirectory (relsPath: string) =
    let at = relsPath.LastIndexOf "_rels/"
    if at <= 0 then "" else relsPath.Substring(0, at).TrimEnd('/')

let private contentTypeTagPattern = System.Text.RegularExpressions.Regex("<(Default|Override)\\s[^>]*>")

/// The content type [Content_Types].xml gives each part: by its name, or by its extension.
let private contentTypeLookup (xml: string) : (string -> string option) * string list =
    let defaults = System.Collections.Generic.Dictionary<string, string>()
    let overrides = System.Collections.Generic.Dictionary<string, string>()
    for m in contentTypeTagPattern.Matches xml do
        match m.Groups.[1].Value, attributeOfTag m.Value "ContentType" with
        | "Default", Some contentType ->
            match attributeOfTag m.Value "Extension" with
            | Some extension -> defaults.[extension.ToLowerInvariant()] <- contentType
            | None -> ()
        | "Override", Some contentType ->
            match attributeOfTag m.Value "PartName" with
            | Some name -> overrides.[name.ToLowerInvariant()] <- contentType
            | None -> ()
        | _ -> ()
    let lookup (path: string) =
        match overrides.TryGetValue (("/" + path).ToLowerInvariant()) with
        | true, contentType -> Some contentType
        | false, _ ->
            let dot = path.LastIndexOf '.'
            if dot < 0 then None
            else
                match defaults.TryGetValue (path.Substring(dot + 1).ToLowerInvariant()) with
                | true, contentType -> Some contentType
                | false, _ -> None
    lookup, [ for KeyValue (name, _) in overrides -> name ]

/// Whether the [Content_Types].xml that was read is right for the package
/// that is written: every part has the type the writer gives it, and it names
/// no part that is not written.
let private contentTypesStillRight (original: string) (generated: string) (writtenPaths: string list) =
    let originalType, originalOverrides = contentTypeLookup original
    let generatedType, _ = contentTypeLookup generated
    let written = writtenPaths |> List.map (fun path -> ("/" + path).ToLowerInvariant()) |> Set.ofList
    writtenPaths |> List.forall (fun path -> path = "[Content_Types].xml" || (match originalType path with Some contentType -> generatedType path = Some contentType | None -> false))
    && originalOverrides |> List.forall (fun name -> Set.contains name written)

/// Write DOCX from EditorFormat (Phase 2)
let writeFromEditor (paragraphs: obj array) : Fable.Core.JS.Promise<obj> =
    nextShapeId <- 1000
    let hyperlinks = hyperlinkMap paragraphs
    let bodyImages, _ = buildImageScope 1 paragraphs
    let comments = collectComments paragraphs
    let footnotes = collectNotes "footnote" paragraphs
    let endnotes = collectNotes "endnote" paragraphs
    // Rendered first: it picks the default list ids the paragraph writer uses.
    let numberingXml = generateNumbering (collectNumberingInfo paragraphs)
    let styles = collectStyles paragraphs
    let hasComments = not comments.IsEmpty
    let hasFootnotes = not footnotes.IsEmpty
    let hasEndnotes = not endnotes.IsEmpty
    let files = [
        yield { Path = "[Content_Types].xml"; Content = generateContentTypesWithImages hasComments hasFootnotes hasEndnotes bodyImages.Assets }
        yield { Path = "_rels/.rels"; Content = generateRootRels() }
        yield { Path = "word/document.xml"; Content = generateDocumentXmlWithFormattingAndImages paragraphs bodyImages }
        yield { Path = "word/_rels/document.xml.rels"; Content = generateDocumentRelsWithImages hyperlinks hasComments hasFootnotes hasEndnotes bodyImages }
        yield { Path = "word/styles.xml"; Content = generateDefaultStylesWithStyles styles }
        yield { Path = "word/numbering.xml"; Content = numberingXml }
        if hasComments then
            yield { Path = "word/comments.xml"; Content = generateComments comments }
        if hasFootnotes then
            yield { Path = "word/footnotes.xml"; Content = generateNotes "w:footnotes" "w:footnote" footnotes }
        if hasEndnotes then
            yield { Path = "word/endnotes.xml"; Content = generateNotes "w:endnotes" "w:endnote" endnotes }
    ]

    let binaryFiles =
        bodyImages.Assets
        |> Array.choose (fun image -> image.Data |> Option.map (fun data -> sprintf "word/media/%s" image.FileName, data))
        |> Array.toList

    createPackageWithBinary files binaryFiles

[<Emit("(($0 != null && $0.Unchanged === true && $0.RegenerateAll !== true && $0.Preserved != null && $0.Preserved.OpenedBytes instanceof Uint8Array) ? $0.Preserved.OpenedBytes : null)")>]
let private openedBytesOfUnchanged (_package: obj) : obj = jsNative
[<Emit("Promise.resolve(new Blob([$0.slice()], { type: 'application/zip' }))")>]
let private blobOfCopy (_bytes: obj) : Fable.Core.JS.Promise<obj> = jsNative

/// Write an editor package (Word's exportDocumentPackage) as a DOCX: a Blob.
///
/// A package that says `Unchanged: true` - the editor sets it when the
/// document is what it was when it was opened: nothing was edited, or
/// everything was undone - and that was opened from bytes (forEditorPackage /
/// openProgressive keep a copy of them: Preserved.OpenedBytes) is written as
/// the file that was opened, byte for byte: the same container, not its
/// parts zipped again. A host that changes an exported package before it
/// writes it sets Unchanged to false (or deletes it).
let writeFromEditorPackage (package: obj) : Fable.Core.JS.Promise<obj> =
  let opened = openedBytesOfUnchanged package
  if not (isNull opened) then blobOfCopy opened
  else
    nextShapeId <- 1000
    // What was kept of the opened package: written back around the parts the
    // editor owns. A package without it (a new document) is written from scratch.
    let preserved = preservedPackageFromObject (package?Preserved)
    // RegenerateAll: every part is written from the model, as it was before
    // unchanged content was written back as it was read (the round-trip
    // checks use it to see what the model alone gives).
    let regenerateAll = boolOrDefault false (package?RegenerateAll)
    reportPatched <- 0
    reportInline <- 0
    reportModelBlocks <- 0
    reportModelParagraphs <- 0
    reportSourceBlocks <- 0
    reportReasons.Clear()
    shapesFromModel <- regenerateAll
    let sourceOf (key: string) = if regenerateAll then None else Map.tryFind key preserved.SourceXml
    let paragraphs = optionalArray (package?Paragraphs)
    // The body's elements that go back as they were read (see planBody).
    let bodyPlan =
        match sourceOf "document" with
        | Some xml when preserved.BodyItemChunks.Length > 0 ->
            planBody xml preserved.BodyItemChunks preserved.BodyChunkCount paragraphs (boolOrDefault false (package?PageSetupUnchanged))
        | _ -> None
    /// The parts that are written back as they were read keep all the
    /// relationships they had; a part written from the model keeps those the
    /// writer does not make itself.
    let sourceParts = System.Collections.Generic.HashSet<string>()
    if bodyPlan.IsSome then sourceParts.Add "document" |> ignore
    for part in Array.append (optionalArray (package?HeaderParts)) (optionalArray (package?FooterParts)) do
        match optionalString (part?RelationshipId) with
        | Some relId when boolOrDefault false (part?SourceUnchanged) && (sourceOf relId).IsSome -> sourceParts.Add relId |> ignore
        | _ -> ()
    for key, flag in [ "footnotes", package?FootnotesUnchanged; "endnotes", package?EndnotesUnchanged; "comments", package?CommentsUnchanged ] do
        if boolOrDefault false flag && (sourceOf key).IsSome then sourceParts.Add key |> ignore
    let keptFor (key: string) =
        let ofMap (map: Map<string, PreservedRelationshipInfo list>) = map |> Map.tryFind key |> Option.defaultValue []
        if sourceParts.Contains key then
            let all = ofMap preserved.SourceRelationships
            let known = all |> List.map (fun rel -> rel.Id) |> Set.ofList
            all @ (ofMap preserved.PartRelationships |> List.filter (fun rel -> not (Set.contains rel.Id known)))
        else ofMap preserved.PartRelationships
    // (The ids and the pictures of what is written back as it was read are taken.)
    reservedRelationshipIds <-
        (Map.toList preserved.PartRelationships @ (Map.toList preserved.SourceRelationships |> List.filter (fun (key, _) -> sourceParts.Contains key)))
        |> List.collect (fun (_, rels) -> rels |> List.map (fun rel -> rel.Id))
        |> Set.ofList
    reservedPartPaths <- (preserved.Parts @ (if sourceParts.Count > 0 then preserved.SourceParts else [])) |> List.map fst |> Set.ofList
    nextMoveRangeId <- 900000
    let withRelationshipsOf (key: string) (blocks: obj array) (generate: unit -> string) =
        currentPartRelationshipIds <- keptFor key |> List.map (fun rel -> rel.Id) |> Set.ofList
        // The runs of each tracked move in this part, for its range markers.
        moveRunsLeft.Clear()
        moveRangeIds.Clear()
        for para in blocks |> Array.toList |> List.collect collectItemParagraphs do
            for run in optionalArray (para?Runs) do
                match moveKey run with
                | Some (moveKeyValue, _, _) ->
                    moveRunsLeft.[moveKeyValue] <- (match moveRunsLeft.TryGetValue moveKeyValue with | true, count -> count | _ -> 0) + 1
                | None -> ()
        let xml = generate ()
        currentPartRelationshipIds <- Set.empty
        xml
    let pageSetup = optionalPageSetup (package?PageSetup)
    let hasSettings = preserved.SettingsXml.IsSome || (pageSetup |> Option.exists (fun setup -> setup.EvenAndOddHeaders))
    // Relationship ids some section refers to; the parts behind the others are
    // attached to the final section.
    let sectionReferenceIds =
        let finalIds =
            pageSetup
            |> Option.map (fun setup -> (setup.HeaderReferences @ setup.FooterReferences) |> List.map snd)
            |> Option.defaultValue []
        Set.union (collectSectionReferenceIds paragraphs) (Set.ofList finalIds)
    let headerParts = optionalArray (package?HeaderParts) |> headerFooterInfos preserved.PartPaths sectionReferenceIds "header" 1
    let footerParts = optionalArray (package?FooterParts) |> headerFooterInfos preserved.PartPaths sectionReferenceIds "footer" 1
    // A header / footer that is what it was when it was read goes back as its text.
    let partSource (part: HeaderFooterPartInfo) = if part.Unchanged then sourceOf part.RelId else None
    // The blocks of the body that are written from the model: their pictures
    // and hyperlinks are the ones the writer makes relationships for.
    let bodyModelBlocks = bodyPlan |> Option.map (fun plan -> plan.Regenerated) |> Option.defaultValue paragraphs
    let bodyImages, nextAfterBody = buildImageScope 1 bodyModelBlocks
    let mutable nextMediaIndex = nextAfterBody
    let partPackage (part: HeaderFooterPartInfo) =
        if (partSource part).IsSome then { Info = part; Images = emptyImageScope; Hyperlinks = Map.empty }
        else
            let images, next = buildImageScope nextMediaIndex part.Blocks
            nextMediaIndex <- next
            { Info = part; Images = images; Hyperlinks = hyperlinkMap part.Blocks }
    let headerPackages = headerParts |> Array.map partPackage
    let footerPackages = footerParts |> Array.map partPackage
    let contentBlocks =
        [| yield! paragraphs
           for part in headerParts do yield! part.Blocks
           for part in footerParts do yield! part.Blocks |]
    // (A link the body had a relationship for keeps it: see hyperlinkMapWith.)
    let hyperlinks = hyperlinkMapWith (if sourceParts.Contains "document" then keptFor "document" else []) bodyModelBlocks
    // The links the relationships part gets an entry of its own for: the others are among the kept ones.
    let newBodyHyperlinks =
        let keptIds = keptFor "document" |> List.map (fun rel -> rel.Id) |> Set.ofList
        hyperlinks |> Map.filter (fun _ id -> not (Set.contains id keptIds))
    // Note and comment bodies come with the package; a reference whose body
    // is not there is written from the text on the reference run.
    let comments = mergeCommentParts (commentPartsFromPackage (package?Comments)) (collectComments contentBlocks)
    let footnotes = mergeNoteParts (notePartsFromPackage (package?Footnotes)) (collectNotes "footnote" contentBlocks)
    let endnotes = mergeNoteParts (notePartsFromPackage (package?Endnotes)) (collectNotes "endnote" contentBlocks)
    // Each of these parts has its own relationships: its images and hyperlinks.
    let bodyPart (blocks: obj array) =
        let images, next = buildImageScope nextMediaIndex blocks
        nextMediaIndex <- next
        images, hyperlinkMap blocks
    let footnoteBlocks = footnotes |> List.toArray |> Array.collect (fun note -> note.Blocks)
    let endnoteBlocks = endnotes |> List.toArray |> Array.collect (fun note -> note.Blocks)
    let commentBlocks = comments |> List.toArray |> Array.collect (fun comment -> comment.Blocks)
    // The notes and the comments as they were read, when nothing of them was changed.
    let footnotesSource = if boolOrDefault false (package?FootnotesUnchanged) then sourceOf "footnotes" else None
    let endnotesSource = if boolOrDefault false (package?EndnotesUnchanged) then sourceOf "endnotes" else None
    let commentsSource = if boolOrDefault false (package?CommentsUnchanged) then sourceOf "comments" else None
    let footnoteImages, footnoteLinks = if footnotesSource.IsSome then emptyImageScope, Map.empty else bodyPart footnoteBlocks
    let endnoteImages, endnoteLinks = if endnotesSource.IsSome then emptyImageScope, Map.empty else bodyPart endnoteBlocks
    let commentImages, commentLinks = if commentsSource.IsSome then emptyImageScope, Map.empty else bodyPart commentBlocks
    let partBody (images: ImageScope) (links: Map<string, string>) (blocks: obj array) =
        blocks |> generateBlocks (generateDocumentItemWithImages links images) (generateTextBoxRun defaultPageSetupInfo links images)
    let allImages =
        [| yield! bodyImages.Assets
           for part in headerPackages do yield! part.Images.Assets
           for part in footerPackages do yield! part.Images.Assets
           yield! footnoteImages.Assets
           yield! endnoteImages.Assets
           yield! commentImages.Assets |]
    let allBlocks =
        [| yield! contentBlocks
           yield! footnoteBlocks
           yield! endnoteBlocks
           yield! commentBlocks |]
    // Rendered first: it picks the default list ids the paragraph writer uses.
    let numberingXml =
        let infos = collectNumberingInfoOf (not (isNullOrUndefined (package?Preserved))) allBlocks
        // Blocks written from the model with a list that has no id of the file.
        let modelBlocks =
            [| yield! bodyModelBlocks
               for part in Array.append headerParts footerParts do
                   if (partSource part).IsNone then yield! part.Blocks
               if footnotesSource.IsNone then yield! footnoteBlocks
               if endnotesSource.IsNone then yield! endnoteBlocks
               if commentsSource.IsNone then yield! commentBlocks |]
        let needsDefaults =
            modelBlocks
            |> Array.toList
            |> List.collect collectItemParagraphs
            |> List.exists (fun para ->
                let listStyle = para?ListStyle |> stringOrDefault "none"
                (listStyle = "bullet" || listStyle = "numbered") && (optionalIntString (para?ListId)).IsNone)
        let merged =
            // (Also with RegenerateAll: the lists are not blocks, and they are as they were.)
            if boolOrDefault false (package?ListsUnchanged) then
                Map.tryFind "numbering" preserved.SourceXml |> Option.bind (fun original -> mergeNumbering original infos needsDefaults)
            else
                // Lists whose format the editor changed: those levels alone are written anew.
                let changed = optionalArray (package?ChangedListLevels) |> Array.map (fun key -> stringOrDefault "" key)
                if changed.Length = 0 then None
                else
                    Map.tryFind "numbering" preserved.SourceXml
                    |> Option.bind (fun original -> mergeNumbering original infos needsDefaults)
                    |> Option.bind (fun merged -> patchNumberingLevels merged infos changed)
        // (Rendered also when it is not written: it picks the default list ids.)
        let generated = lazy (generateNumbering infos)
        match merged with
        | Some text -> Some text
        | None ->
            let text = generated.Force()
            // A document that was opened without a part for its lists, and has no list, is saved without one.
            // (Nor is one made for lists the document refers to and never had a definition of.)
            if not (isNullOrUndefined (package?Preserved)) && not (Map.containsKey "numbering" preserved.SourceXml)
               && (infos |> List.forall (fun info -> info.Undefined)) && not needsDefaults then None
            else Some text
    let hasNumbering = numberingXml.IsSome
    let styles =
        (collectPackageStyles (optionalArray (package?Styles)) @ collectStyles allBlocks)
        |> List.distinctBy (fun style -> style.StyleType, style.StyleId)
        // The document defaults are not a style: a style of that name (inferred by
        // the editor from table cell paragraphs) would take their place when read.
        |> List.filter (fun style -> style.StyleId <> documentDefaultsStyleId || style.StyleType = "defaults")
    let hasComments = not comments.IsEmpty || commentsSource.IsSome
    let hasFootnotes = not footnotes.IsEmpty || footnotesSource.IsSome
    let hasEndnotes = not endnotes.IsEmpty || endnotesSource.IsSome
    // Pictures and other parts that only content written back as it was read refers to.
    let preserved = if sourceParts.Count > 0 then { preserved with Parts = preserved.Parts @ preserved.SourceParts } else preserved
    let partRelsNeeded (key: string) (images: ImageScope) (links: Map<string, string>) =
        images.Assets.Length > 0 || not links.IsEmpty || not (keptFor key).IsEmpty

    let files = [
        yield { Path = "[Content_Types].xml"; Content = generateContentTypesPreserved preserved hasNumbering hasSettings hasComments hasFootnotes hasEndnotes headerParts footerParts allImages }
        yield { Path = "_rels/.rels"; Content = generateRootRelsPreserved preserved }
        yield { Path = "word/document.xml"
                Content =
                    withRelationshipsOf "document" bodyModelBlocks (fun () ->
                        match bodyPlan with
                        | Some plan -> generateDocumentXmlFromPlan plan bodyImages hyperlinks pageSetup headerParts footerParts
                        | None -> generateDocumentXmlPackage paragraphs bodyImages pageSetup headerParts footerParts) }
        yield { Path = "word/_rels/document.xml.rels"; Content = generateDocumentRelsPreserved (keptFor "document") hasNumbering hasSettings newBodyHyperlinks hasComments hasFootnotes hasEndnotes headerParts footerParts bodyImages }
        yield { Path = "word/styles.xml"
                Content =
                    match preserved.StylesXml with
                    | Some original ->
                        let changed =
                            if isNullOrUndefined package?ChangedStyles then None
                            else Some (optionalArray (package?ChangedStyles) |> Array.map (fun key -> unbox<string> key) |> Set.ofArray)
                        mergeStyles original changed styles
                    | None -> generateDefaultStylesWithStyles styles }
        if not (keptFor "styles").IsEmpty then
            yield { Path = "word/_rels/styles.xml.rels"; Content = generatePartRelsWithKept (keptFor "styles") Map.empty emptyImageScope }
        match numberingXml with
        | Some text ->
            yield { Path = "word/numbering.xml"; Content = text }
            // (What the lists refer to: the picture of a picture bullet.)
            if not (keptFor "numbering").IsEmpty then
                yield { Path = "word/_rels/numbering.xml.rels"; Content = generatePartRelsWithKept (keptFor "numbering") Map.empty emptyImageScope }
        | None -> ()
        if hasSettings then
            yield { Path = "word/settings.xml"
                    Content =
                        match preserved.SettingsXml with
                        | Some original -> patchSettings original pageSetup
                        | None -> generateSettings pageSetup }
            if not (keptFor "settings").IsEmpty then
                yield { Path = "word/_rels/settings.xml.rels"; Content = generatePartRelsWithKept (keptFor "settings") Map.empty emptyImageScope }
        for header in headerPackages do
            yield { Path = "word/" + header.Info.Target
                    Content =
                        match partSource header.Info with
                        | Some text -> text
                        | None -> withRelationshipsOf header.Info.RelId header.Info.Blocks (fun () -> generateHeaderFooterXml "w:hdr" header.Hyperlinks header.Images header.Info.Blocks) }
            if partRelsNeeded header.Info.RelId header.Images header.Hyperlinks then
                yield { Path = sprintf "word/_rels/%s.rels" header.Info.Target; Content = generatePartRelsWithKept (keptFor header.Info.RelId) header.Hyperlinks header.Images }
        for footer in footerPackages do
            yield { Path = "word/" + footer.Info.Target
                    Content =
                        match partSource footer.Info with
                        | Some text -> text
                        | None -> withRelationshipsOf footer.Info.RelId footer.Info.Blocks (fun () -> generateHeaderFooterXml "w:ftr" footer.Hyperlinks footer.Images footer.Info.Blocks) }
            if partRelsNeeded footer.Info.RelId footer.Images footer.Hyperlinks then
                yield { Path = sprintf "word/_rels/%s.rels" footer.Info.Target; Content = generatePartRelsWithKept (keptFor footer.Info.RelId) footer.Hyperlinks footer.Images }
        if hasComments then
            yield { Path = "word/comments.xml"
                    Content =
                        match commentsSource with
                        | Some text -> text
                        | None -> withRelationshipsOf "comments" commentBlocks (fun () -> generateCommentsPart (partBody commentImages commentLinks) comments) }
            if partRelsNeeded "comments" commentImages commentLinks then
                yield { Path = "word/_rels/comments.xml.rels"; Content = generatePartRelsWithKept (keptFor "comments") commentLinks commentImages }
        if hasFootnotes then
            yield { Path = "word/footnotes.xml"
                    Content =
                        match footnotesSource with
                        | Some text -> text
                        | None -> withRelationshipsOf "footnotes" footnoteBlocks (fun () -> generateNotesPart "w:footnotes" "w:footnote" "w:footnoteRef" (partBody footnoteImages footnoteLinks) footnotes) }
            if partRelsNeeded "footnotes" footnoteImages footnoteLinks then
                yield { Path = "word/_rels/footnotes.xml.rels"; Content = generatePartRelsWithKept (keptFor "footnotes") footnoteLinks footnoteImages }
        if hasEndnotes then
            yield { Path = "word/endnotes.xml"
                    Content =
                        match endnotesSource with
                        | Some text -> text
                        | None -> withRelationshipsOf "endnotes" endnoteBlocks (fun () -> generateNotesPart "w:endnotes" "w:endnote" "w:endnoteRef" (partBody endnoteImages endnoteLinks) endnotes) }
            if partRelsNeeded "endnotes" endnoteImages endnoteLinks then
                yield { Path = "word/_rels/endnotes.xml.rels"; Content = generatePartRelsWithKept (keptFor "endnotes") endnoteLinks endnoteImages }
    ]

    // Kept parts go back as they were, except where the writer produced a part
    // of the same name.
    let written = files |> List.map (fun file -> file.Path) |> Set.ofList
    let binaryFiles =
        (allImages
         |> Array.choose (fun image -> image.Data |> Option.map (fun data -> sprintf "word/media/%s" image.FileName, data))
         |> Array.toList)
        @ (preserved.Parts |> List.filter (fun (path, _) -> not (Set.contains path written)))

    // Pictures nothing refers to any more do not stay in the package (a block
    // that was deleted, or that is now written from the model with a picture
    // file of its own, leaves the relationship and the file it had): a picture
    // relationship of a part whose XML does not use its id goes, and then every
    // file under word/media that no relationship part names. The same of the
    // relationship of a hyperlink (a paragraph with a link that was deleted, or
    // whose link was taken away): it goes when the part's XML does not use its
    // id any more - one the part did not use when it was read stays, the
    // file's own.
    let files =
        let contentOf = files |> List.map (fun file -> file.Path, file.Content) |> Map.ofList
        let sourceOfPath =
            [ yield "word/document.xml", "document"
              yield "word/footnotes.xml", "footnotes"
              yield "word/endnotes.xml", "endnotes"
              yield "word/comments.xml", "comments"
              for part in Array.append headerParts footerParts do yield "word/" + part.Target, part.RelId ]
            |> List.choose (fun (path, key) ->
                Map.tryFind key preserved.SourceXml
                |> Option.map (fun xml ->
                    let ids (map: Map<string, PreservedRelationshipInfo list>) = map |> Map.tryFind key |> Option.defaultValue [] |> List.map (fun rel -> rel.Id)
                    path, (xml, Set.ofList (ids preserved.SourceRelationships @ ids preserved.PartRelationships))))
            |> Map.ofList
        files
        |> List.map (fun file ->
            let name = file.Path.Substring(file.Path.LastIndexOf '/' + 1)
            if not (file.Path.StartsWith "word/_rels/" && name.EndsWith ".rels") then file
            else
                match Map.tryFind ("word/" + name.Substring(0, name.Length - 5)) contentOf with
                | None -> file
                | Some part ->
                    let partPath = "word/" + name.Substring(0, name.Length - 5)
                    let unused (relationship: string) =
                        let id = System.Text.RegularExpressions.Regex.Match(relationship, "\\sId=\"([^\"]*)\"")
                        let usedIn (text: string) = text.IndexOf("\"" + id.Groups.[1].Value + "\"", System.StringComparison.Ordinal) >= 0
                        let picture = relationship.IndexOf "/relationships/image\"" >= 0 && relationship.IndexOf "TargetMode=\"External\"" < 0
                        // (A link's relationship the part had and did not use when it was read: the file's own.)
                        let inherited =
                            match Map.tryFind partPath sourceOfPath with
                            | Some (source, ids) -> id.Success && Set.contains id.Groups.[1].Value ids && not (usedIn source)
                            | None -> false
                        let link = relationship.IndexOf "/relationships/hyperlink\"" >= 0 && not inherited
                        (picture || link) && id.Success && not (usedIn part)
                    let pieces = System.Text.RegularExpressions.Regex.Matches(file.Content, "<Relationship\\s[^>]*/>")
                    let mutable content = file.Content
                    for i in pieces.Count - 1 .. -1 .. 0 do
                        let piece = pieces.[i]
                        if unused piece.Value then content <- content.Substring(0, piece.Index) + content.Substring(piece.Index + piece.Length)
                    { file with Content = content })
    let binaryFiles =
        let relationshipTexts =
            (files |> List.filter (fun file -> file.Path.EndsWith ".rels") |> List.map (fun file -> file.Content))
            @ (binaryFiles |> List.filter (fun (path, _) -> path.EndsWith ".rels") |> List.map (fun (_, data) -> System.Text.Encoding.UTF8.GetString data))
        binaryFiles
        |> List.filter (fun (path, _) ->
            not (path.StartsWith "word/media/")
            || (let name = path.Substring(path.LastIndexOf '/' + 1)
                relationshipTexts |> List.exists (fun text -> text.IndexOf("media/" + name + "\"", System.StringComparison.Ordinal) >= 0)))

    // The relationship parts and [Content_Types].xml that say what they said when they were read go back as that text.
    let files =
        if regenerateAll || preserved.OriginalText.IsEmpty then files
        else
            let writtenPaths = (files |> List.map (fun file -> file.Path)) @ (binaryFiles |> List.map fst)
            files
            |> List.map (fun file ->
                match Map.tryFind file.Path preserved.OriginalText with
                | None -> file
                | Some original ->
                    if file.Path = "[Content_Types].xml" then
                        if contentTypesStillRight original file.Content writtenPaths then { file with Content = original } else file
                    elif file.Path.EndsWith ".rels" then
                        let directory = relationshipsOwnerDirectory file.Path
                        match relationshipEntries directory original, relationshipEntries directory file.Content with
                        | Some before, Some after when before = after -> { file with Content = original }
                        | _ -> file
                    else file)

    reservedRelationshipIds <- Set.empty
    reservedPartPaths <- Set.empty
    createPackageWithBinary files binaryFiles
