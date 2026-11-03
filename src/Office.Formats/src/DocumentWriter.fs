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
    /// w:vAlign: "center", "both", "bottom" (none or "top": not written).
    VerticalAlign: string option
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
    VerticalAlign = None
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
            VerticalAlign = optionalString (value?VerticalAlign)
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
/// type, pgSz, pgMar, pgNumType, cols, vAlign, titlePg.
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
      match setup |> Option.bind (fun setup -> setup.VerticalAlign) with
      | Some value when value = "center" || value = "both" || value = "bottom" -> yield selfClosing "w:vAlign" [("w:val", value)]
      | _ -> ()
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
        |> Option.map ModelText.wordColor

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
    // (xml:space where the text needs it: white space at an end, or two together.)
    let textElement (value: string) = elementWithText textName (if XmlPatch.needsPreserve value then [("xml:space", "preserve")] else []) value
    if isNull text || text = "" then [ textElement "" ]
    else
        ModelText.textPieces text
        |> List.map (function
            | ModelText.Text value -> textElement value
            | ModelText.Tab -> selfClosing "w:tab" []
            | ModelText.Break -> selfClosing "w:br" [])

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
        let placeholder = ModelText.isFieldPlaceholder fallbackRun.Text
        let result = optionalString (runObj?FieldResult) |> Option.defaultValue (if placeholder then "" else fallbackRun.Text)
        // (A part of the result alone - the paragraph was split in the field's result: what the run shows.)
        let result = if not placeholder && fallbackRun.Text <> "" && fallbackRun.Text.Length < result.Length && result.Contains fallbackRun.Text then fallbackRun.Text else result
        let resultRun = generateRunWithObjectProps runObj { fallbackRun with Text = result }
        element "w:fldSimple" [("w:instr", fieldCode)] [resultRun]
    | parts ->
        fieldResultParts <- []
        element "w:fldSimple" [("w:instr", fieldCode)] (generateRunWithObjectProps runObj fallbackRun :: (parts |> List.map (fun (part, run) -> generateRunWithObjectProps part run)))


/// The bookmarks that start before a run: every one of them, once, in the
/// order of the file. A token "id!" is a bookmark that ends where it starts.
let private bookmarkStartXml (runObj: obj) =
    optionalString (runObj?BookmarkStartId)
    |> Option.orElse (optionalString (runObj?BookmarkId))
    |> Option.map (fun ids ->
        ModelText.bookmarkStarts ids (optionalString (runObj?BookmarkName) |> Option.defaultValue "")
        |> List.map (fun start ->
            selfClosing "w:bookmarkStart" [("w:id", start.Id); ("w:name", start.Name)] + (if start.EndsHere then selfClosing "w:bookmarkEnd" [("w:id", start.Id)] else ""))
        |> String.concat "")

/// The bookmarks that end after a run (and those opened after the last run
/// of a paragraph: "!id=name" ends there too, "!id=name<" ends further on).
let private bookmarkEndXml (runObj: obj) =
    let explicitEnd =
        optionalString (runObj?BookmarkEndId)
        |> Option.orElse (optionalString (runObj?BookmarkId))
    let fieldWrappedEnd =
        optionalString (runObj?FieldCode)
        |> Option.bind (fun _ -> optionalString (runObj?BookmarkStartId))
        |> Option.filter ModelText.isOneOpenBookmark
    explicitEnd
    |> Option.orElse fieldWrappedEnd
    |> Option.map (fun ids ->
        ModelText.bookmarkEnds ids
        |> List.map (function
            | ModelText.Opens (id, name, endsHere) ->
                selfClosing "w:bookmarkStart" [("w:id", id); ("w:name", name)] + (if endsHere then selfClosing "w:bookmarkEnd" [("w:id", id)] else "")
            | ModelText.Ends id -> selfClosing "w:bookmarkEnd" [("w:id", id)])
        |> String.concat "")

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
    /// The same of each header and footer, by its relationship id: the element of the part's root each item
    /// came from, and how many elements the root has.
    PartItemChunks: Map<string, int array * int>
    /// The same of each note and comment, by its part and id ("footnotes:2"): the element of the note's
    /// element each item came from, and how many elements it has.
    NoteItemChunks: Map<string, int array * int>
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
    PartItemChunks = Map.empty
    NoteItemChunks = Map.empty
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
          PartItemChunks =
            optionalArray (value?PartItemChunks)
            |> Array.toList
            |> List.choose (fun entry ->
                optionalString (entry?Part)
                |> Option.map (fun key -> key, (optionalArray (entry?Chunks) |> Array.map (fun chunk -> intOrDefault (-1) chunk), intOrDefault 0 (entry?Count))))
            |> Map.ofList
          NoteItemChunks =
            optionalArray (value?NoteItemChunks)
            |> Array.toList
            |> List.choose (fun entry ->
                match optionalString (entry?Part), optionalString (entry?Id) with
                | Some part, Some id -> Some (part + ":" + id, (optionalArray (entry?Chunks) |> Array.map (fun chunk -> intOrDefault (-1) chunk), intOrDefault 0 (entry?Count)))
                | _ -> None)
            |> Map.ofList
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
    |> List.filter (ModelText.isAnchor >> not)

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
        |> List.filter (fun rel -> rel.External && OpcXml.relationshipKind rel.Type = "hyperlink")
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
    match ModelText.dataUrl src with
    | Some { MediaType = mime; Base64 = Some data } ->
        mimeToExtensionAndContentType mime |> Option.map (fun (extension, contentType) -> extension, contentType, base64ToBytes data)
    | _ -> None

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
                if ModelText.isLinkedSource src then
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
/// flipH="1" / flipV="1" of an a:xfrm, for a shape or picture that is flipped.
let private flipAttributes (flipH: obj) (flipV: obj) : (string * string) list =
    [ if boolOrDefault false flipH then yield ("flipH", "1")
      if boolOrDefault false flipV then yield ("flipV", "1") ]

/// The flip of a VML shape's style ("x", "y", "x y"), for a shape that is flipped.
let private vmlFlip (flipH: obj) (flipV: obj) : string option =
    match boolOrDefault false flipH, boolOrDefault false flipV with
    | true, true -> Some "x y"
    | true, false -> Some "x"
    | false, true -> Some "y"
    | false, false -> None

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
                                element "a:xfrm" (flipAttributes runObj?ImageFlipH runObj?ImageFlipV) [
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
            |> fun setup -> { setup with HeaderReferences = []; FooterReferences = []; PageNumberStart = None; PageNumberFormat = None; VerticalAlign = optionalString (para?SectionVerticalAlign) }
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
        // A drop cap: the frame Word writes for one, at the place of w:framePr.
        match optionalString (para?DropCapKind) with
        | Some kind when kind = "drop" || kind = "margin" ->
            let distance = optionalFloat (para?DropCapDistance) |> Option.defaultValue 0.0
            yield selfClosing "w:framePr" [
                yield ("w:dropCap", kind)
                yield ("w:lines", string (max 1 (optionalInt (para?DropCapLines) |> Option.defaultValue 1)))
                if distance > 0.0 then yield ("w:hSpace", pixelsToTwips distance)
                yield ("w:wrap", "around")
                yield ("w:vAnchor", "text")
                yield ("w:hAnchor", "text") ]
        | _ -> ()
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
        // (The outline level the paragraph states itself: a title's 0, a body text level.)
        match optionalInt (para?OutlineLevel) with
        | Some level when level >= 0 && level <= 9 -> yield selfClosing "w:outlineLvl" [("w:val", string level)]
        | _ -> ()
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
/// The XML of a paragraph that is made of a paragraph that was read and nothing was done to (its
/// SourcePara says Verbatim; its SourceAt is the element), and where that element is in the part.
type private ParagraphSource =
    abstract Of: obj -> string option
    /// Where the paragraph a SourcePara names is in the part (start, end).
    abstract SpanOf: obj -> (int * int) option
let mutable private paragraphSource : ParagraphSource =
    { new ParagraphSource with
        member _.Of _ = None
        member _.SpanOf _ = None }
/// Whether the paragraphs being written may be made of the XML they were read from (see generateTable).
let mutable private paragraphsInContext = true
/// How many paragraphs the last save wrote from the model, and how many of the XML they were read from
/// in blocks that were written from the model.
let mutable private reportGenerated = 0
let mutable private reportDerived = 0
/// Paragraphs written from their source with the model's changes (docs/save-architecture.md): of reportGenerated.
let mutable private reportSpliced = 0
/// Paragraphs changed in the XML they were read from (written from their source: see splicedRuns).
let mutable private reportPatched = 0

// ---- the marks of ranges the model does not hold ---------------------------------
//
// A paragraph that is written from the model is a paragraph that was read,
// and the XML it was read from can hold the marks of ranges the model has
// nothing of: the ranges of comments and of permissions (w:commentRangeStart,
// w:permStart and their ends) wherever they stand, and bookmarks where no run
// is there to have them - every bookmark of a paragraph without runs, and the
// end, before the first run of a paragraph, of a bookmark that began in a
// paragraph before. They are written with the paragraph: at its start what
// stood before its first run (and what opens a range), at its end what stood
// after its last (and what closes one). A range of a comment that is not there
// any more is not. (The ranges of moved text are the model's: the writer makes
// them for the runs that were moved.)

/// The number of the save that is being written (see carried marks: what is set on an item for one save is not another's).
let mutable private saveNumber = 0
/// The ids of the comments the document has, while its body is written.
let mutable private liveCommentIds : Set<string> = Set.empty
/// The ids of the comments the document had when it was read and has no more (their ranges go).
let mutable private deletedCommentIds : Set<string> = Set.empty
/// A mark of a range (w:bookmarkStart, w:commentRangeEnd ...): its kind ("bookmark"), whether it opens the
/// range, and the range's id.
type private RangeMark = { Node: XmlDsl.XmlNode; Kind: string; Opens: bool; Id: string }
    with member mark.Key = mark.Kind + ":" + mark.Id

/// The marks of ranges the writer carries and balances: those it carries between blocks and paragraphs, and
/// the ranges of moved text too (which are the model's: balanced, not carried).
let private carriedRangeKinds = [ "bookmark"; "commentRange"; "perm" ]
let private allRangeKinds = carriedRangeKinds @ [ "moveFromRange"; "moveToRange" ]

/// The element as a mark of a range of one of these kinds (an empty w:...Start or w:...End with a w:id).
let private rangeMarkOf (kinds: string list) (node: XmlDsl.XmlNode) : RangeMark option =
    if XmlPatch.prefix node <> "w" || not (XmlPatch.elements node).IsEmpty then None
    else
        let local = XmlPatch.localName node
        kinds
        |> List.tryPick (fun kind -> if local = kind + "Start" then Some (kind, true) elif local = kind + "End" then Some (kind, false) else None)
        |> Option.bind (fun (kind, opens) -> XmlPatch.attribute node "w:id" |> Option.map (fun id -> { Node = node; Kind = kind; Opens = opens; Id = id }))

/// The marks of ranges of these kinds in a tree, in document order.
let private rangeMarksIn (kinds: string list) (node: XmlDsl.XmlNode) : RangeMark list =
    XmlPatch.descendantsWhere (rangeMarkOf kinds >> Option.isSome) node |> List.choose (rangeMarkOf kinds)

/// The end mark of a range.
let private rangeEndXml (kind: string) (id: string) = selfClosing ("w:" + kind + "End") [ ("w:id", id) ]

/// The marks of a paragraph's XML that the model does not hold: those to write at the start of
/// the paragraph and those to write at its end. (What a shape, a picture or an object anchored in
/// the paragraph holds - the paragraphs of a text box - is not of the paragraph.)
let private carriedMarksOf (paragraphXml: string) : string * string =
    match XmlPatch.read paragraphXml with
    | None -> "", ""
    | Some paragraph ->
        // The marks and the runs (a run, an equation, a simple field) of the paragraph, in order.
        let tokens = ResizeArray<RangeMark option>()
        let rec visit (node: XmlDsl.XmlNode) =
            match XmlPatch.name node with
            | "mc:AlternateContent" | "w:drawing" | "w:pict" | "w:object" -> ()
            | name ->
                match rangeMarkOf carriedRangeKinds node with
                | Some mark -> tokens.Add (Some mark)
                | None ->
                    if name = "w:r" || name = "m:oMath" || name = "m:oMathPara" || name = "w:fldSimple" then tokens.Add None
                    List.iter visit (XmlPatch.elements node)
        List.iter visit (XmlPatch.elements paragraph.Node)
        let tokens = List.ofSeq tokens
        let firstRun = tokens |> List.tryFindIndex Option.isNone |> Option.defaultValue -1
        let lastRun = tokens |> List.tryFindIndexBack Option.isNone |> Option.defaultValue -1
        let atStart = System.Text.StringBuilder()
        let atEnd = System.Text.StringBuilder()
        let opened = System.Collections.Generic.HashSet<string>()
        tokens
        |> List.iteri (fun index token ->
            match token with
            | None -> ()
            | Some mark ->
                let bytes = XmlPatch.bytesOf paragraph.Text mark.Node
                let beforeRuns = firstRun < 0 || index < firstRun
                let afterRuns = lastRun >= 0 && index > lastRun
                if mark.Kind = "bookmark" then
                    // (Those of a run are the model's.)
                    if firstRun < 0 then atStart.Append(bytes) |> ignore
                    elif mark.Opens then opened.Add mark.Id |> ignore
                    elif beforeRuns && not (opened.Contains mark.Id) then atStart.Append(bytes) |> ignore
                elif mark.Kind <> "commentRange" || Set.contains mark.Id liveCommentIds then
                    if beforeRuns then atStart.Append(bytes) |> ignore
                    elif afterRuns then atEnd.Append(bytes) |> ignore
                    elif mark.Opens then atStart.Append(bytes) |> ignore
                    else atEnd.Append(bytes) |> ignore)
        atStart.ToString(), atEnd.ToString()

/// A paragraph the model wrote with marks at its start (after its properties) and at its end.
let private withCarriedMarks (atStart: string) (atEnd: string) (xml: string) : string =
    if atStart = "" && atEnd = "" then xml
    else
        match XmlPatch.read xml with
        | Some paragraph when XmlPatch.name paragraph.Node = "w:p" ->
            let edits =
                [ if atStart <> "" then yield XmlPatch.Prepend atStart
                  if atEnd <> "" then yield XmlPatch.Append atEnd ]
            XmlPatch.apply XmlPatch.Word.order paragraph edits
        | _ -> xml

/// The XML a paragraph that is written from the model was read from, where that is known (see ParagraphSource).
type private ParagraphOrigin =
    abstract Of: obj -> string option
let mutable private paragraphOrigin : ParagraphOrigin = { new ParagraphOrigin with member _.Of _ = None }

/// The marks to write with a paragraph of the model (see above): of the XML it is a paragraph of.
let private carriedMarksOfItem (para: obj) : string * string =
    let source =
        if not (isNullOrUndefined (para?CarrySource)) && intOrDefault -1 (para?CarrySave) = saveNumber then optionalString (para?CarrySource)
        elif not (isNullOrUndefined (para?SourcePara)) || not (boolOrDefault false (para?SourceCopy)) then paragraphOrigin.Of para
        else None
    match source with
    | Some xml -> carriedMarksOf xml
    | None -> "", ""

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
      ImageFlipH = false
      ImageFlipV = false
    }

/// The properties the model writes for a run of the editor's package (none for a picture).
let private modelRunProperties (runObj: obj) : string list =
    // (NoProps: a run of which nothing is taken to be written - see Docx.derivedEdit.)
    if boolOrDefault false (runObj?IsImage) || not (isNullOrUndefined runObj?OpaqueXml) || boolOrDefault false (runObj?NoProps) then []
    else orderedRunProps (runExtrasFromObject runObj) (editorRunOfObject runObj)


// ---------------------------------------------------------------------------
// Splicing a paragraph (docs/save-architecture.md): a paragraph that is written
// from the model and was read from the part is written from its source - what
// of it the model does not hold stays where it stood: the content controls,
// smart tags and custom XML around its runs, the elements between them the
// model has no field for, its properties when they are what they were, and
// every run that is as it was read, as its bytes.
// ---------------------------------------------------------------------------

/// The text of the part being written, as it was read (set by the body writer).
let mutable private spliceSourceText : string = ""

[<Fable.Core.Emit("JSON.stringify($0)")>]
let private jsonText (_value: obj) : string = jsNative
[<Fable.Core.Emit("JSON.stringify(Object.assign({}, $0, { Runs: undefined, SpliceFrom: undefined, SpliceOriginal: undefined, SpliceLeading: undefined, SpliceTrailing: undefined, SpliceJoined: undefined, SpliceBreakKept: undefined, CarrySource: undefined, CarrySave: undefined, SourcePara: undefined, SourceIndex: undefined, SourceUnchanged: undefined, SourceWhy: undefined, SourceTextChanges: undefined }))")>]
let private paragraphPropertiesText (_value: obj) : string = jsNative

/// A paragraph written from its source (see splicedRuns, defined with the merging of properties
/// it uses): the runs, and the properties as they were read where they are what they were.
type private ParagraphSplicer =
    /// The runs, the properties as they were read, and the paragraph's start tag (and what stands before its
    /// properties) as it was read - a copy's without the ids that are the paragraph's alone. `runXml`: the XML
    /// the model writes of a run.
    abstract Splice: para: obj -> pieces: (obj * string * (string -> string)) list -> runXml: (obj -> string) -> wrapOf: (obj -> string -> string) -> (string list * string option * string) option
    /// The properties as they were read (`source`) with what differs between the model's properties then
    /// (`was`) and now (`now`) changed in them; None where that cannot be done in them.
    abstract MergeProperties: source: string -> was: string -> now: string -> string option
let mutable private paragraphSplicer : ParagraphSplicer =
    { new ParagraphSplicer with
        member _.Splice _ _ _ _ = None
        member _.MergeProperties _ _ _ = None }

/// Which of what stands before the first run of its source and after the last a paragraph made of a part of
/// one that was read has (SpliceLeading, SpliceTrailing): a half of a split has what stands on its side of
/// the cut; a paragraph others were joined to has nothing of its own after its runs.
/// `spec`: what says the cut and the joins (the paragraph's SourcePara, or the edit of it).
let private setSpliceCut (spec: obj) (para: obj) =
    if not (isNullOrUndefined spec) then
        let cut = optionalArray (spec?Cut) |> Array.map (fun value -> intOrDefault -1 value)
        if cut.Length = 2 then
            // (The length of the paragraph that was read: its text, or that of the runs it had.)
            let length =
                match optionalString (spec?OldText) with
                | Some text -> text.Length
                | None -> optionalArray (para?SpliceOriginal?Runs) |> Array.sumBy (fun run -> (stringOrDefault "" (run?Text)).Length)
            para?SpliceLeading <- (cut.[0] = 0)
            para?SpliceTrailing <- (cut.[1] >= length)
        if optionalArray (spec?Join) |> Array.isEmpty |> not then
            para?SpliceTrailing <- false
            // (The paragraphs joined to it as they were read: their runs are found in their own XML.)
            para?SpliceJoined <- optionalArray (spec?Join) |> Array.map (fun join -> join?Original) |> Array.filter (isNullOrUndefined >> not)

/// The further runs of a field's result the model holds as several (a part of it typed or formatted), by
/// the field's first run: what the splicer writes of the field (see splicedRuns).
let private spliceFieldParts = System.Collections.Generic.Dictionary<obj, obj list>(HashIdentity.Reference)

let private generateParagraphWithHyperlinksAndImagesWithPageSetup (packageSetup: PageSetupInfo) (hyperlinks: Map<string, string>) (images: ImageScope) (para: obj) : string =
  match (if isNullOrUndefined (para?SourcePara) || not paragraphsInContext then None else paragraphSource.Of para) with
  | Some xml ->
    reportDerived <- reportDerived + 1
    xml
  | None ->
    reportGenerated <- reportGenerated + 1
    // A paragraph made of one that was read (a part of it, or it with others joined) is spliced from that one's
    // XML (see splicedRuns): a part takes the elements of the source that stand among its own runs, and those
    // before or after them where it holds the start or the end of the paragraph.
    if isNullOrUndefined (para?SpliceFrom) && not (isNullOrUndefined (para?SourcePara)) && paragraphsInContext then
        let spec = para?SourcePara
        let original = if isNullOrUndefined (spec?After) then null else spec?After?SpliceOriginal
        match paragraphSource.SpanOf para with
        | Some (start, finish) when not (isNullOrUndefined original) ->
            para?SpliceFrom <- createObj [ "Start" ==> start; "End" ==> finish ]
            para?SpliceOriginal <- original
            setSpliceCut spec para
        | _ -> ()
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
                   // (Runs that make the result as it was read are that field: the next run with the same
                   // code and result is another one - two page numbers side by side. A result that was
                   // typed into - in another format: a run of its own - is not the text that was read
                   // any more; its runs are still one field.)
                   && text <> whole do
                 text <- text + stringOrDefault "" (runsArray.[next]?Text)
                 next <- next + 1
             if next > index + 1 then
                 fieldParts.[index] <- [ for part in index + 1 .. next - 1 -> runsArray.[part], editorRunOfObject runsArray.[part] ]
                 for part in index + 1 .. next - 1 do fieldPartRuns.Add part |> ignore
             else next <- index + 1
         | _ -> ()
         index <- next)
    // (What goes around a run: its tracked change, its link.)
    let wrapOf (runObj: obj) (inner: string) =
        let runXml = revisionWrappedRun runObj inner
        match runHyperlinkTarget runObj with
        | Some target when (ModelText.anchorOf target).IsSome -> element "w:hyperlink" [("w:anchor", (ModelText.anchorOf target).Value); ("w:history", "1")] [runXml]
        | Some target ->
            match Map.tryFind target hyperlinks with
            | Some relId -> element "w:hyperlink" [("r:id", relId); ("w:history", "1")] [runXml]
            | None -> runXml
        | None -> runXml
    let runs =
        runsArray
        |> Array.toList
        |> List.mapi (fun runIndex runObj -> runIndex, runObj)
        |> List.filter (fun (runIndex, _) -> not (fieldPartRuns.Contains runIndex))
        |> List.map (fun (runIndex, runObj) ->
            fieldResultParts <- (match fieldParts.TryGetValue runIndex with | true, parts -> parts | _ -> [])
            if not fieldResultParts.IsEmpty then spliceFieldParts.[runObj] <- fieldResultParts |> List.map fst
            let runXml =
                if not (isNullOrUndefined runObj?OpaqueXml) then
                    // An object kept from the file (chart, SmartArt, OLE object, ...):
                    // its element goes back as it was read. If a relationship it
                    // needs is not in this part (the run was pasted from another
                    // document), only its description is written.
                    let xml = runObj?OpaqueXml |> unbox<string>
                    let referenced = OpcXml.relationshipIdsIn xml
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
            runObj, runXml, wrapOf runObj
        )
    // A paragraph read from the part is written from its source (see splicedRuns).
    // (The XML the model writes of a run of text, for what a run of the source was to the model: see splicedRuns.)
    let textRunXml (runObj: obj) =
        fieldResultParts <- []
        generateReferenceRun runObj (editorRunOfObject runObj)
    // (What wraps a run, to compare its wrapping before and after: its tracked change, its link's target - not
    // the link's relationship, which only the links still in the document have.)
    let wrappingOf (runObj: obj) (inner: string) =
        revisionWrappedRun runObj inner + (match runHyperlinkTarget runObj with Some target -> "\u0003" + target | None -> "")
        // (And the field it is of: text typed beside a simple field is not in it.)
        + (match optionalString (runObj?FieldCode) with Some code -> "\u0004" + code | None -> "")
    let splice = paragraphSplicer.Splice para runs textRunXml wrappingOf
    spliceFieldParts.Clear()
    // (A paragraph written from its source is one changed in the XML it was read from, not one of the model's.)
    if splice.IsSome then
        reportSpliced <- reportSpliced + 1
        reportPatched <- reportPatched + 1
        reportGenerated <- reportGenerated - 1
    let runs =
        match splice with
        | Some (spliced, _, _) -> spliced
        | None -> runs |> List.map (fun (_, inner, wrap) -> wrap inner)
    // (Its properties as they were read, when they are what they were.)
    let pPr =
        match splice with
        | Some (_, Some original, _) when paragraphPropertiesText para = paragraphPropertiesText para?SpliceOriginal -> original
        | Some (_, None, _) when paragraphPropertiesText para = paragraphPropertiesText para?SpliceOriginal -> ""
        // (Changed: what changed, in the properties as they were read - those the model does not hold stay.)
        | Some (_, Some original, _) ->
            let was = generateParaPropsWithPageSetup packageSetup para?SpliceOriginal
            match paragraphSplicer.MergeProperties original was pPr with
            | Some merged -> merged
            | None -> pPr
        | _ -> pPr
    let runs =
        // (A page break after the paragraph is one the paragraph has: a section that ends with it
        // says nothing of a break - the section after it says how it starts.)
        // (Spliced from a source that has the break where it stands: there it stays - see spliceSourceOf.)
        if boolOrDefault false (para?PageBreakAfter) && not (splice.IsSome && boolOrDefault false (para?SpliceBreakKept)) then
            runs @ [element "w:r" [] [selfClosing "w:br" [("w:type", (if boolOrDefault false (para?ColumnBreakAfter) then "column" else "page"))]]]
        else runs

    // (A paragraph that states nothing of its own has no properties element.)
    // (With a start and an end tag also when it is empty: a shape anchored in it goes before its end tag.)
    let written =
        match splice with
        // (Its start tag as it was read.)
        | Some (_, _, head) -> head + (if XmlPatch.isEmpty pPr then "" else pPr) + String.concat "" runs + "</w:p>"
        | None ->
            match (if XmlPatch.isEmpty pPr then runs else pPr :: runs) with
            | [] -> "<w:p></w:p>"
            | children -> element "w:p" [] children
    // (What the XML it was read from holds and the model does not: the marks of ranges - in a paragraph written
    // from its source they stand where they stood.)
    match splice, carriedMarksOfItem para with
    | Some _, _ | None, ("", "") -> written
    | None, (atStart, atEnd) -> withCarriedMarks atStart atEnd written

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

/// Whether a w:tblGrid that was read has the columns the table has now: as many, of the same
/// widths. (A table the model holds no grid of keeps the one it was read with.)
let private sourceGridIs (sourceGrid: string) (table: obj) : bool =
    let now = optionalArray (table?GridColumns) |> Array.choose (fun value -> optionalFloat value) |> Array.map roundToInt
    if now.Length = 0 then true
    else
        let columns = XmlPatch.read sourceGrid |> Option.map (fun grid -> XmlPatch.elements grid.Node |> List.filter (fun node -> XmlPatch.localName node = "gridCol")) |> Option.defaultValue []
        let widths = columns |> List.map (fun column -> XmlPatch.attribute column "w:w" |> Option.bind (fun value -> match System.Double.TryParse value with | true, width -> Some (roundToInt width) | _ -> None))
        let ok = widths |> List.forall Option.isSome
        let read = widths |> List.choose id |> ResizeArray
        ok && read.Count = now.Length && (Array.forall2 (fun (a: int) (b: int) -> abs (a - b) <= 1) (read.ToArray()) now)

/// The properties the model writes for a cell, a row and a table (the children of w:tcPr, w:trPr, w:tblPr).
let private tableCellProperties (cell: obj) : string list =
    // CT_TcPr order: tcW, gridSpan, vMerge, tcBorders, shd, tcMar, vAlign.
    [
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


/// (And the row's w:tblPrEx, "" when it has none.)
let private tableRowProperties (row: obj) : string * string list =
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
    rowExceptions, rowPropChildren

let private tablePropertyElements (table: obj) : string list =
    // CT_TblPr order: tblStyle, tblW, jc, tblCellSpacing, tblBorders, shd, tblLayout, tblCellMar, tblLook.
    [
        // The table keeps its style: the reader resolves the style's formatting
        // into the cells from the reference (and the style itself is kept in
        // styles.xml), so dropping it changed how the saved table read back.
        match optionalString (table?TableStyleId) with
        | Some styleId -> yield selfClosing "w:tblStyle" [("w:val", styleId)]
        | None -> ()
        // A floating table stays one: without its w:tblpPr the table came back in
        // the text flow and the paragraphs that ran beside it below it.
        match optionalString (table?FloatingXml) with
        | Some floatingXml when (XmlPatch.read floatingXml |> Option.map (fun floating -> XmlPatch.name floating.Node)) = Some "w:tblpPr" -> yield floatingXml
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


/// What says how a paragraph of the editor's package looks, but for where it comes from: one made like another
/// (SourceCopy) has, made like it, what that one has.
[<Fable.Core.Emit("JSON.stringify(Object.assign({}, $0, { Runs: undefined, SourceCopy: undefined, SpliceFrom: undefined, SpliceOriginal: undefined, SpliceLeading: undefined, SpliceTrailing: undefined, SpliceJoined: undefined, SpliceBreakKept: undefined, CarrySource: undefined, CarrySave: undefined, SourcePara: undefined, SourceIndex: undefined, SourceUnchanged: undefined, SourceWhy: undefined, SourceTextChanges: undefined }))")>]
let private likePropertiesText (_value: obj) : string = jsNative

/// Whether a paragraph of the editor's package holds nothing: no text, no picture, object, field, note, comment or bookmark.
let private paragraphIsEmpty (para: obj) : bool =
    optionalArray (para?Runs)
    |> Array.forall (fun run ->
        stringOrDefault "" (run?Text) = "" && not (boolOrDefault false (run?IsImage)) && isNullOrUndefined (run?OpaqueXml)
        && [ "FieldCode"; "FieldResult"; "NoteId"; "CommentId"; "BookmarkStartId"; "BookmarkEndId"; "BookmarkId" ] |> List.forall (fun key -> isNullOrUndefined (run?(key))))

/// An empty paragraph made like one that was read (`paragraphXml`: the first of a cell next to which a row or a
/// column was put in): its element without its ids (w14:paraId, w14:textId) and without what it holds but its
/// properties - those of the paragraph and of its mark, as they were read.
let private emptyParagraphLike (paragraphXml: string) : string option =
    match XmlPatch.read paragraphXml with
    | Some paragraph when XmlPatch.name paragraph.Node = "w:p" ->
        let span = XmlPatch.span paragraph.Node
        let startTag = XmlPatch.startTag paragraph.Text paragraph.Node [ XmlPatch.RemoveAttribute "w14:paraId"; XmlPatch.RemoveAttribute "w14:textId" ]
        // (<w:p/>: the start tag is the element.)
        if span.StartTagEnd >= span.End then Some startTag
        else
            let properties = XmlPatch.elements paragraph.Node |> List.tryFind (fun kid -> XmlPatch.localName kid = "pPr") |> Option.map (XmlPatch.bytesOf paragraph.Text) |> Option.defaultValue ""
            Some (startTag + properties + "</" + XmlPatch.name paragraph.Node + ">")
    | _ -> None

/// A table of the body as it was read, for a table that is written from the model (its SourceTable): the
/// text of its properties and grid, and of each of its rows and cells by where its element starts in the
/// part (the origin the model has of it: SourceRow, SourceCell); the layout of each (XmlPatch.layout: its
/// rows, cells, paragraphs among what stands around them - content controls, custom XML, marks, white
/// space). And how the properties that were read are changed (see mergedProperties).
/// CopyStartTag: the start tag for a copy - without the ids that are the element's alone (w14:paraId, w14:textId).
type private SourceCellText = { CellXml: string; CellStartTag: string; CopyStartTag: string; CellProps: string; CellHead: string; CellChildren: XmlPatch.Layout }
/// RowHead / TableHead: the text from the start tag to the end of the properties (and grid), as it was read.
type private SourceRowText = { RowXml: string; RowStartTag: string; CopyStartTag: string; RowExceptions: string; RowProps: string; RowHead: string; RowChildren: XmlPatch.Layout }
type private SourceTableText = {
    /// What the table is in, as it was read (a content control, custom XML; the white space before it): to the table's start, and after the table.
    WrapStart: string
    WrapEnd: string
    TableHead: string
    TableStartTag: string
    TableProps: string
    TableGrid: string
    TableRows: System.Collections.Generic.Dictionary<int, SourceRowText>
    TableCells: System.Collections.Generic.Dictionary<int, SourceCellText>
    TableChildren: XmlPatch.Layout
    /// The table's element in the part's tree (w:tbl).
    TableNode: XmlDsl.XmlNode
}
type private TableSource =
    abstract Of: obj -> SourceTableText option
    /// The table of a w:tbl of the part being written (nothing around it: WrapStart and WrapEnd are "").
    abstract At: XmlDsl.XmlNode -> SourceTableText option
    /// container, the properties that were read, what the model wrote before and writes now -> the properties, or None (see mergedProperties).
    abstract Merge: string * string * string list * string list -> string option
let mutable private tableSource : TableSource =
    { new TableSource with
        member _.Of _ = None
        member _.At _ = None
        member _.Merge (_, _, _, _) = None }
/// The table being written of a table that was read; and whether the paragraphs of what is being written
/// may be made of the XML they were read from: those of a cell that is written with the properties its
/// table, row and cell were read with - what a paragraph's XML leaves to its table's style it takes from there.
let mutable private currentSourceTable : SourceTableText option = None

/// The element a row or cell of a table that was read is of (its SourceRow / SourceCell): where it
/// starts, whether it is a copy of it, whether it is the element still.
let private sourceOfPart (part: obj) : (int * bool * bool) option =
    if isNullOrUndefined part then None
    else
        let at = intOrDefault -1 (part?At)
        if at < 0 then None else Some (at, boolOrDefault false (part?Copy), boolOrDefault false (part?Same))

let rec private generateTableCell (hyperlinks: Map<string, string>) (images: ImageScope) (cell: obj) : string =
    let source =
        match currentSourceTable, sourceOfPart (cell?SourceCell) with
        | Some table, Some (at, copy, same) ->
            (match table.TableCells.TryGetValue at with
             | true, known -> Some (known, copy, same)
             | _ -> None)
        | _ -> None
    match source with
    | Some (known, false, true) -> known.CellXml
    | _ ->
        // Its properties: the ones it was read with (or the one it is a copy of) - changed where the model's changed - or the model's.
        let tcPrChildren = tableCellProperties cell
        let properties =
            match source with
            | Some (known, copy, _) ->
                let before = cell?SourceCell?Before
                if isNullOrUndefined before then Some (if copy then known.CopyStartTag + known.CellProps else known.CellHead)
                else
                    tableSource.Merge ("w:tcPr", known.CellProps, tableCellProperties before, tcPrChildren)
                    |> Option.map (fun props -> (if copy then known.CopyStartTag else known.CellStartTag) + props)
            | None -> None
        let outerContext = paragraphsInContext
        let outerTable = currentSourceTable
        paragraphsInContext <- outerContext && properties.IsSome
        // (Each paragraph with the w:p it is of, when that is one of the cell's and it is the first of it.)
        let paragraphItems = optionalArray (cell?Paragraphs)
        // (A cell made like one that was read - a row or a column put in - that holds one empty paragraph made like
        // that cell's first, as the model has it: that paragraph's element, empty - see emptyParagraphLike.)
        let likeParagraph (para: obj) : string option =
            match source with
            | Some (_, true, _) when paragraphItems.Length = 1 && Array.isEmpty (optionalArray (cell?NestedTables)) && paragraphIsEmpty para ->
                let was = cell?SourceCell?Was
                let first = if isNullOrUndefined was then None else optionalArray (was?Paragraphs) |> Array.tryHead
                match first, optionalInt (para?SourceAt) with
                | Some first, Some at when boolOrDefault false (para?SourceCopy) && optionalInt (first?SourceAt) = Some at && likePropertiesText para = likePropertiesText first ->
                    paragraphOrigin.Of para |> Option.bind emptyParagraphLike
                | _ -> None
            | _ -> None
        let paragraphs =
            paragraphItems
            |> Array.map (fun para ->
                match (if paragraphsInContext then likeParagraph para else None) with
                | Some xml ->
                    reportDerived <- reportDerived + 1
                    xml
                | None -> generateParagraphWithHyperlinksAndImages hyperlinks images para)
        let paragraphAt =
            let used = System.Collections.Generic.HashSet<int>()
            paragraphItems
            |> Array.map (fun item ->
                match source, optionalInt (item?SourceAt) with
                | Some (known, false, _), Some at when not (boolOrDefault false (item?SourceCopy)) && XmlPatch.isItemOf known.CellChildren at && used.Add at -> XmlPatch.AtItem at
                | _ -> XmlPatch.WithPrevious)
        let nestedTables = optionalArray (cell?NestedTables)
        // NestedTableAnchors.[i] = how many of the cell's paragraphs precede nested
        // table i. The writer used to read only Paragraphs, so a nested table and
        // everything in it was deleted on save.
        let anchors = optionalArray (cell?NestedTableAnchors) |> Array.map (fun value -> jsNumber value |> roundToInt)
        let content = ResizeArray<XmlPatch.Placed * string>()
        let mutable lastIsTable = false
        let emitTablesAt index =
            nestedTables
            |> Array.iteri (fun tableIndex nested ->
                let anchor = if tableIndex < anchors.Length then anchors.[tableIndex] else paragraphs.Length
                if (min anchor paragraphs.Length) = index then
                    // (A table in the cell is written from the model: so are its paragraphs.)
                    let inContext = paragraphsInContext
                    paragraphsInContext <- false
                    currentSourceTable <- None
                    content.Add ((XmlPatch.WithPrevious, generateTable hyperlinks images nested))
                    paragraphsInContext <- inContext
                    currentSourceTable <- outerTable
                    lastIsTable <- true)
        for index in 0 .. paragraphs.Length - 1 do
            emitTablesAt index
            content.Add ((paragraphAt.[index], paragraphs.[index]))
            lastIsTable <- false
        emitTablesAt paragraphs.Length
        // A cell must contain, and end with, a paragraph.
        if content.Count = 0 || lastIsTable then content.Add ((XmlPatch.WithPrevious, element "w:p" [] []))
        paragraphsInContext <- outerContext
        match properties, source with
        // (A cell that was read: its paragraphs and tables among what stood around them - content controls, marks.)
        | Some head, Some (known, false, _) -> head + XmlPatch.writeLayout known.CellChildren (content.ToArray()) + "</w:tc>"
        | Some head, _ -> head + (content |> Seq.map snd |> String.concat "") + "</w:tc>"
        | None, _ ->
            let propertiesXml = if tcPrChildren.IsEmpty then "" else element "w:tcPr" [] tcPrChildren
            element "w:tc" [] (propertiesXml :: (content |> Seq.map snd |> List.ofSeq))

and private generateTableRow (hyperlinks: Map<string, string>) (images: ImageScope) (row: obj) : string =
    let source =
        match currentSourceTable, sourceOfPart (row?SourceRow) with
        | Some table, Some (at, copy, same) ->
            (match table.TableRows.TryGetValue at with
             | true, known -> Some (known, copy, same)
             | _ -> None)
        | _ -> None
    match source with
    | Some (known, false, true) -> known.RowXml
    | _ ->
        let rowExceptions, rowPropChildren = tableRowProperties row
        // Its start tag and properties: the ones it was read with (or the one it is a copy of) - changed where the model's changed - or the model's.
        let head =
            match source with
            | Some (known, copy, _) ->
                let startTag = if copy then known.CopyStartTag else known.RowStartTag
                let before = row?SourceRow?Before
                if isNullOrUndefined before then Some (if copy then startTag + known.RowExceptions + known.RowProps else known.RowHead)
                else
                    let (exceptionsBefore, propsBefore) = tableRowProperties before
                    tableSource.Merge ("w:trPr", known.RowProps, propsBefore, rowPropChildren)
                    |> Option.map (fun props -> startTag + (if exceptionsBefore = rowExceptions then known.RowExceptions else rowExceptions) + props)
            | None -> None
        let outerContext = paragraphsInContext
        paragraphsInContext <- outerContext && head.IsSome
        let cells =
            optionalArray (row?Cells)
            |> Array.map (fun cell ->
                let at = sourceOfPart (cell?SourceCell) |> Option.bind (fun (at, copy, _) -> if copy then None else Some at)
                (match at with Some at -> XmlPatch.AtItem at | None -> XmlPatch.WithPrevious), generateTableCell hyperlinks images cell)
        paragraphsInContext <- outerContext
        // (The cells of a row that was read among what stood among them; a copy of a row has its cells alone.)
        let cellsXml =
            match source with
            | Some (known, false, _) -> XmlPatch.writeLayout known.RowChildren cells
            | _ -> cells |> Array.map snd |> String.concat ""
        match head with
        | Some text -> text + cellsXml + "</w:tr>"
        | None ->
            let rowProps = if rowPropChildren.IsEmpty then "" else element "w:trPr" [] rowPropChildren
            element "w:tr" [] [rowExceptions; rowProps; cellsXml]

and private generateTable (hyperlinks: Map<string, string>) (images: ImageScope) (table: obj) : string =
    let tblPrChildren = tablePropertyElements table
    // A table that was read: its properties are the ones it was read with - changed where the
    // model's changed - and its rows and cells may be (see generateTableRow).
    let source = if isNullOrUndefined (table?SourceTable) then None else tableSource.Of (table?SourceTable)
    let properties =
        match source with
        | Some known ->
            let before = table?SourceTable?Before
            // (The grid that was read, unless the table has another by now - a column put in or taken out.)
            let grid = if sourceGridIs known.TableGrid table then known.TableGrid else generateTableGrid table
            if isNullOrUndefined before then Some (if grid = known.TableGrid then known.TableHead else known.TableStartTag + known.TableProps + grid)
            else
                tableSource.Merge ("w:tblPr", known.TableProps, tablePropertyElements before, tblPrChildren)
                |> Option.map (fun props -> known.TableStartTag + props + generateTableGrid table)
        | None -> None
    let outerContext = paragraphsInContext
    let outerTable = currentSourceTable
    paragraphsInContext <- properties.IsSome
    currentSourceTable <- (if properties.IsSome then source else None)
    let rows =
        let written =
            optionalArray (table?Rows)
            |> Array.map (fun row ->
                let at = sourceOfPart (row?SourceRow) |> Option.map (fun (at, _, _) -> at)
                (match at with Some at -> XmlPatch.AtItem at | None -> XmlPatch.WithPrevious), generateTableRow hyperlinks images row)
        match currentSourceTable with
        | Some known -> XmlPatch.writeLayout known.TableChildren written
        | _ -> written |> Array.map snd |> String.concat ""
    paragraphsInContext <- outerContext
    currentSourceTable <- outerTable
    match properties, source with
    | Some text, Some known -> known.WrapStart + text + rows + "</w:tbl>" + known.WrapEnd
    | _ ->
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
    let style = match vmlFlip item?FlipH item?FlipV with Some flip -> style + ";flip:" + flip | None -> style
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
/// The edits of a shape's VML fallback (the mc:Fallback of its alternate content) that give it the place
/// and the size the shape has now, said in VML's way: margin-left / margin-top / width / height of the
/// style of its shape element, in points, from the frame the anchor names
/// (mso-position-horizontal-relative / mso-position-vertical-relative); an alignment in a direction that
/// was moved in (mso-position-horizontal: center) and a size or place in percent go. Everything else of
/// the fallback - its text, fill, line, path - stays the text it was. None: a fallback this does not know
/// how to place (a line, a group, a shape in a line of text, no style): it is left out then, as it says
/// the old place.
let private vmlFallbackEdits (fallback: XmlDsl.XmlNode) (item: obj) (moveH: bool) (moveV: bool) (resize: bool) : (XmlDsl.XmlNode * XmlPatch.ElementEdit list) option =
    let isShape (node: XmlDsl.XmlNode) =
        XmlPatch.prefix node = "v" && List.contains (XmlPatch.localName node) [ "shape"; "rect"; "oval"; "roundrect" ] && (XmlPatch.attribute node "style").IsSome
    let grouped = XmlPatch.descendantsWhere (fun node -> XmlPatch.name node = "v:group") fallback |> List.isEmpty |> not
    match XmlPatch.descendantsWhere isShape fallback with
    | shape :: _ when not grouped ->
        let mutable declarations = XmlPatch.Style.parse (XmlPatch.attribute shape "style" |> Option.defaultValue "")
        let set property value = declarations <- XmlPatch.Style.set property value declarations
        let remove property = declarations <- XmlPatch.Style.remove property declarations
        let points (pixels: float) = string (System.Math.Round(pixels * 75.0) / 100.0) + "pt"
        if (moveH || moveV) && not (declarations |> List.contains ("position", "absolute")) then None
        else
            let mutable ok = true
            if moveH then
                match optionalFloat (item?X) with
                | Some x ->
                    set "margin-left" (points x)
                    remove "mso-position-horizontal"
                    remove "mso-left-percent"
                    match optionalString (item?PositionHRelativeFrom) with
                    | Some "page" -> set "mso-position-horizontal-relative" "page"
                    | Some "margin" -> set "mso-position-horizontal-relative" "margin"
                    | Some "column" -> set "mso-position-horizontal-relative" "text"
                    | Some "character" -> set "mso-position-horizontal-relative" "char"
                    | Some "leftMargin" -> set "mso-position-horizontal-relative" "left-margin-area"
                    | Some "rightMargin" -> set "mso-position-horizontal-relative" "right-margin-area"
                    | Some "insideMargin" -> set "mso-position-horizontal-relative" "inner-margin-area"
                    | Some "outsideMargin" -> set "mso-position-horizontal-relative" "outer-margin-area"
                    | Some _ -> ok <- false
                    | None -> ()
                | None -> ok <- false
            if moveV then
                match optionalFloat (item?Y) with
                | Some y ->
                    set "margin-top" (points y)
                    remove "mso-position-vertical"
                    remove "mso-top-percent"
                    match optionalString (item?PositionVRelativeFrom) with
                    | Some "page" -> set "mso-position-vertical-relative" "page"
                    | Some "margin" -> set "mso-position-vertical-relative" "margin"
                    | Some "paragraph" -> set "mso-position-vertical-relative" "text"
                    | Some "line" -> set "mso-position-vertical-relative" "line"
                    | Some "topMargin" -> set "mso-position-vertical-relative" "top-margin-area"
                    | Some "bottomMargin" -> set "mso-position-vertical-relative" "bottom-margin-area"
                    | Some "insideMargin" -> set "mso-position-vertical-relative" "inner-margin-area"
                    | Some "outsideMargin" -> set "mso-position-vertical-relative" "outer-margin-area"
                    | Some _ -> ok <- false
                    | None -> ()
                | None -> ok <- false
            if resize then
                match optionalFloat (item?Width), optionalFloat (item?Height) with
                | Some width, Some height ->
                    set "width" (points width)
                    set "height" (points height)
                    remove "mso-width-percent"
                    remove "mso-height-percent"
                | _ -> ok <- false
            if ok then Some (shape, [ XmlPatch.SetAttribute ("style", XmlPatch.Style.write declarations) ]) else None
    | _ -> None

/// The element a shape was read from with the shape's place and size as they are now (a shape
/// that was moved or resized and is otherwise what was read: Docx.textBoxMovedFrom): the
/// offsets of wp:positionH / wp:positionV, wp:extent and the extent of the shape's own
/// a:xfrm. Everything else of the element stays the text it was. Its VML fallback says the
/// place and the size in its own way: they are changed in it too (vmlFallbackEdits), and a
/// fallback that cannot be is left out (the DrawingML is what a reader that knows it takes;
/// Word writes the fallback anew). None: not in this element - a shape in a line
/// of text that is moved, a group, a canvas, an anchor without the offsets.
let private patchShapeAnchor (xml: string) (item: obj) : string option =
    let moved = item?SourceMoved
    let moveH = boolOrDefault false (moved?H)
    let moveV = boolOrDefault false (moved?V)
    let resize = boolOrDefault false (moved?Size)
    match XmlPatch.readFragmentRoot xml with
    | None -> None
    | Some fragment ->
        let named (qualified: string) (node: XmlDsl.XmlNode) = XmlPatch.descendantsWhere (fun inner -> XmlPatch.name inner = qualified) node
        // (The DrawingML: the choice of the alternate content, or the drawing itself.)
        let drawing = named "mc:Choice" fragment.Node |> List.tryHead |> Option.defaultValue fragment.Node
        let fallback = named "mc:Fallback" fragment.Node |> List.tryHead
        let unplaceable =
            XmlPatch.descendantsWhere (fun node -> XmlPatch.prefix node = "wpg" || XmlPatch.prefix node = "wpc" || XmlPatch.name node = "w:pict") drawing
        let anchor = named "wp:anchor" drawing |> List.tryHead
        if not unplaceable.IsEmpty then None
        elif (moveH || moveV) && anchor.IsNone then None
        else
            let emu (pixels: float) = string (int64 (System.Math.Round(pixels * 9525.0)))
            let mutable ok = true
            let edits = ResizeArray<XmlDsl.XmlNode * XmlPatch.ElementEdit list>()
            // (The first element of that name in the drawing, with these edits; there being none, the shape is not placed here.)
            let edit (qualified: string) (within: XmlDsl.XmlNode) (made: XmlDsl.XmlNode -> XmlPatch.ElementEdit list) =
                match named qualified within |> List.tryHead with
                | Some node -> edits.Add ((node, made node))
                | None -> ok <- false
            // (A position: its frame, and an offset in the place of an alignment.)
            let position (offset: float) (relativeFrom: string option) (node: XmlDsl.XmlNode) =
                [ match relativeFrom with Some frame -> XmlPatch.SetAttribute ("relativeFrom", frame) | None -> ()
                  XmlPatch.RemoveChild "wp:align"
                  XmlPatch.RemoveChild "wp:posOffset"
                  XmlPatch.AddChild ("wp:posOffset", "<wp:posOffset>" + emu offset + "</wp:posOffset>") ]
            if moveH then
                match optionalFloat (item?X) with
                | Some x -> edit "wp:positionH" drawing (position x (optionalString (item?PositionHRelativeFrom)))
                | None -> ok <- false
            if moveV then
                match optionalFloat (item?Y) with
                | Some y -> edit "wp:positionV" drawing (position y (optionalString (item?PositionVRelativeFrom)))
                | None -> ok <- false
            if resize then
                match optionalFloat (item?Width), optionalFloat (item?Height) with
                | Some width, Some height ->
                    let size = [ XmlPatch.SetAttribute ("cx", emu width); XmlPatch.SetAttribute ("cy", emu height) ]
                    edit "wp:extent" drawing (fun _ -> size)
                    // (The shape's own extent: in its properties.)
                    match XmlPatch.descendantsWhere (fun node -> XmlPatch.localName node = "spPr" && (XmlPatch.prefix node = "wps" || XmlPatch.prefix node = "pic")) drawing with
                    | properties :: _ ->
                        match named "a:ext" properties with
                        | extent :: _ -> edits.Add ((extent, size))
                        | [] -> ()
                    | [] -> ()
                | _ -> ok <- false
            if not ok then None
            else
                // (The fallback at the shape's place too, or left out.)
                match fallback with
                | Some fallback ->
                    match vmlFallbackEdits fallback item moveH moveV resize with
                    | Some placed -> edits.Add placed
                    | None -> edits.Add ((fallback, [ XmlPatch.Removed ]))
                | None -> ()
                Some (XmlPatch.applyFragment (fun _ -> []) fragment (XmlPatch.editsAt fragment.Node (List.ofSeq edits)))

let rec private generateTextBoxRun (packageSetup: PageSetupInfo) (hyperlinks: Map<string, string>) (images: ImageScope) (item: obj) : string =
    // A shape that is what was read goes back as the element it was read from
    // (its DrawingML and its VML fallback, with what the model does not hold
    // of them) - when everything that element refers to is in this part.
    let source =
        (if shapesFromModel then None else optionalString (item?SourceXml))
        |> Option.filter (fun xml ->
            OpcXml.relationshipIdsIn xml
            |> List.forall (fun id -> Set.contains id currentPartRelationshipIds))
    // (Moved or resized, and otherwise what was read: its element with the anchor changed in it.)
    let source =
        match source with
        | Some xml when not (isNullOrUndefined (item?SourceMoved)) -> patchShapeAnchor xml item
        | other -> other
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
          match vmlFlip item?FlipH item?FlipV with
          | Some flip -> yield "flip:" + flip
          | None -> ()
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
                element "a:xfrm" ((if abs rotation >= 0.0001 then [("rot", string (roundToInt (rotation * 60000.0)))] else []) @ flipAttributes item?FlipH item?FlipV) [
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
    let last () = if parts.Count > 0 then XmlPatch.read parts.[parts.Count - 1] else None
    for item in items do
        let isTextBox = (item?Type |> stringOrDefault "Paragraph") = "TextBox"
        // (The paragraph a shape is anchored in: the block before it, or the last paragraph of a table before it.)
        let anchor =
            if not isTextBox then None
            else
                last ()
                |> Option.bind (fun block ->
                    match XmlPatch.name block.Node with
                    | "w:p" -> Some (block, block.Node)
                    // (The paragraph that ends last: not one of a text box in it.)
                    | "w:tbl" ->
                        XmlPatch.descendantsWhere (fun node -> XmlPatch.name node = "w:p") block.Node
                        |> List.sortBy (fun para -> (XmlPatch.span para).End)
                        |> List.tryLast
                        |> Option.map (fun para -> block, para)
                    | _ -> None)
        match anchor with
        | Some (block, para) ->
            parts.[parts.Count - 1] <- XmlPatch.apply XmlPatch.Word.order block (XmlPatch.editsAt block.Node [ para, [ XmlPatch.Append (generateTextBoxRun item) ] ])
        | None ->
            let xml = generateItem item
            // Blocks of one content control (the same w:sdtPr, with an id, one after another) are in one w:sdt.
            let control (source: XmlPatch.SourceElement) =
                if XmlPatch.name source.Node <> "w:sdt" then None
                else
                    match XmlPatch.child "sdtPr" source.Node, XmlPatch.child "sdtContent" source.Node with
                    | Some properties, Some content when (XmlPatch.child "id" properties).IsSome -> Some (XmlPatch.bytesOf source.Text properties, content)
                    | _ -> None
            match last () |> Option.bind (fun previous -> control previous |> Option.map (fun found -> previous, found)), XmlPatch.read xml |> Option.bind (fun now -> control now |> Option.map (fun found -> now, found)) with
            | Some (previous, (properties, content)), Some (now, (sameProperties, nowContent)) when properties = sameProperties ->
                let held = XmlPatch.elements nowContent |> List.map (XmlPatch.bytesOf now.Text) |> String.concat ""
                parts.[parts.Count - 1] <- XmlPatch.apply XmlPatch.Word.order previous (XmlPatch.editsAt previous.Node [ content, [ XmlPatch.Append held ] ])
            | _ -> parts.Add xml
    String.concat "" parts

/// The value of `<w:name w:val="..."/>` in an element's XML.
let private childVal (name: string) (xml: string) =
    XmlPatch.readFragmentRoot xml
    |> Option.bind (fun fragment -> XmlPatch.descendantsWhere (fun node -> XmlPatch.name node = "w:" + name && (XmlPatch.attribute node "w:val").IsSome) fragment.Node |> List.tryHead)
    |> Option.bind (fun node -> XmlPatch.attribute node "w:val")

let private wrapContentControl (item: obj) (content: string) =
    let id, tag, alias = optionalString (item?ContentControlId), optionalString (item?ContentControlTag), optionalString (item?ContentControlAlias)
    match id, tag, alias, optionalString (item?ContentControlXml) with
    | None, None, None, _ -> content
    // The control as the file has it (binding, placeholder, kind) while it is still the one the file had.
    | _, _, _, Some xml when childVal "id" xml = id && childVal "tag" xml = tag && childVal "alias" xml = alias ->
        "<w:sdt>" + xml + "<w:sdtContent>" + content + "</w:sdtContent></w:sdt>"
    | id, tag, alias, _ ->
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
            if PartPath.directory path = "word" && PartPath.extension path = "xml" then Some (PartPath.fileName path) else None)
    let taken = System.Collections.Generic.HashSet<string>()
    for KeyValue (_, path) in originalPaths do
        if PartPath.directory path = "word" then taken.Add ((PartPath.fileName path).ToLowerInvariant()) |> ignore
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

type private ElementRanges = { Starts: int array; Ends: int array; Nodes: XmlDsl.XmlNode array; CloseAt: int }

let rec private storyParagraphCount (item: obj) : int =
    match item?Type |> stringOrDefault "Paragraph" with
    | "TextBox" -> 0
    | "Table" ->
        optionalArray (item?Rows)
        |> Array.sumBy (fun row ->
            optionalArray (row?Cells)
            |> Array.sumBy (fun cell -> (optionalArray (cell?Paragraphs)).Length + (optionalArray (cell?NestedTables) |> Array.sumBy storyParagraphCount)))
    | _ -> 1

/// The paragraph items of an item of the editor's package, in the order their paragraphs have
/// in the XML (as Docx.storyItems: a table in a cell where it is anchored among the cell's paragraphs).
let rec private storyParagraphItems (item: obj) : obj list =
    match item?Type |> stringOrDefault "Paragraph" with
    | "TextBox" -> []
    | "Table" ->
        optionalArray (item?Rows)
        |> Array.toList
        |> List.collect (fun row ->
            optionalArray (row?Cells)
            |> Array.toList
            |> List.collect (fun cell ->
                let paragraphs = optionalArray (cell?Paragraphs) |> Array.toList
                let nested = optionalArray (cell?NestedTables) |> Array.toList
                if nested.IsEmpty then paragraphs
                else
                    let count = paragraphs.Length
                    let anchors = optionalArray (cell?NestedTableAnchors) |> Array.map (fun anchor -> intOrDefault count anchor)
                    let placed = nested |> List.mapi (fun i inner -> inner, (if i < anchors.Length then anchors.[i] else count) |> max 0 |> min count) |> List.sortBy snd
                    let before index = placed |> List.filter (fun (_, anchor) -> anchor = index) |> List.collect (fun (inner, _) -> storyParagraphItems inner)
                    (paragraphs |> List.mapi (fun index para -> before index @ [ para ]) |> List.concat) @ before count))
    | _ -> [ item ]

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

/// The children of a properties element as the model wrote it (`<w:pPr>...</w:pPr>`, `<w:pPr/>`), each its text.
let private propertyElementsOf (element: string) : string list =
    match XmlPatch.read element with
    | Some source -> XmlPatch.elements source.Node |> List.map (XmlPatch.bytesOf source.Text)
    | None -> []

/// The properties element `source` as it was read (w:rPr, w:pPr, w:tcPr ...; "": there is none yet) with the
/// change the model made between writing its children `was` and `now` made in it (XmlPatch): what the model
/// does not write - elements, attributes, a tracked change of the properties - stays as it was read. Where
/// the XML states the complex-script twin of bold or italic (w:bCs, w:iCs) and the model does not, the twin
/// changes with it. "" when nothing is left in it; None when `source` is not one element.
let private mergedProperties (container: string) (source: string) (was: string list) (now: string list) : string option =
    let wrap (children: string list) = "<" + container + ">" + String.concat "" children + "</" + container + ">"
    match XmlPatch.read (if source = "" then "<" + container + "/>" else source), XmlPatch.read (wrap was), XmlPatch.read (wrap now) with
    | Some read, Some before, Some after ->
        let edits = XmlPatch.diff before after
        let named = edits |> List.map XmlPatch.childOf
        // An attribute the model does not write that would overrule one it sets: the same through the
        // document's theme (w:asciiTheme beside w:ascii, w:themeColor beside w:val), or in other units
        // (w:afterLines beside w:after, w:leftChars beside w:left).
        let overrules (element: string) (written: Map<string, string>) (key: string) =
            let lower = key.ToLowerInvariant()
            not (written.ContainsKey key)
            && (if element = "w:rFonts" then lower.Contains "theme" && written |> Map.exists (fun known _ -> lower = known.ToLowerInvariant() + "theme")
                else lower.Contains "theme" || ModelText.isInOtherUnits key)
        let edits =
            edits
            |> List.collect (fun edit ->
                let edit =
                    match edit with
                    // A property element the model changed says what the model writes now in every attribute the
                    // model writes (also where the XML said it another way than the model did before: a spacing
                    // that was automatic); one the model wrote and writes no more goes; one it never writes stays,
                    // but for those that would overrule what it sets.
                    | XmlPatch.EditChild (name, inner, whole) when inner |> List.exists (function XmlPatch.SetAttribute _ | XmlPatch.RemoveAttribute _ -> true | _ -> false) ->
                        let written = XmlPatch.childAttributes after name
                        let gone = XmlPatch.childAttributes read name |> Map.toList |> List.map fst |> List.filter (overrules name written)
                        let removed = inner |> List.filter (function XmlPatch.RemoveAttribute _ -> true | _ -> false)
                        let children = inner |> List.filter (function XmlPatch.SetAttribute _ | XmlPatch.RemoveAttribute _ -> false | _ -> true)
                        let set = written |> Map.toList |> List.map XmlPatch.SetAttribute
                        XmlPatch.EditChild (name, set @ removed @ (gone |> List.map XmlPatch.RemoveAttribute) @ children, whole)
                    | other -> other
                match XmlPatch.childOf edit with
                | "w:b" | "w:i" as name when XmlPatch.hasChild read (name + "Cs") && not (List.contains (name + "Cs") named) -> [ edit; XmlPatch.ofChild (name + "Cs") edit ]
                | _ -> [ edit ])
        if edits.IsEmpty then Some source
        else
            let merged = XmlPatch.apply XmlPatch.Word.order read edits
            Some (if XmlPatch.isEmpty merged then "" else merged)
    | _ -> None

/// A run of a paragraph's XML and the characters of the paragraph's text it stands for.
/// A paragraph as it was read, for one written from its source: the layout of its runs (XmlPatch.layout -
/// in the wrappers the model does not hold, among elements it does not hold), each run's element, its
/// fields, its properties.
type private SpliceSource = {
    Layout: XmlPatch.Layout
    /// Run start -> its element.
    Runs: System.Collections.Generic.Dictionary<int, XmlDsl.XmlNode>
    /// A field's begin run (where it starts) -> where its end run ends: a field the model holds as one run
    /// is written from there, begin mark to end mark (its code, a form field's data), while it is as read.
    Fields: System.Collections.Generic.Dictionary<int, int>
    /// The paragraph's element, for what is in a field's range.
    Paragraph: XmlDsl.XmlNode
    /// Its start tag and what stands after it before its properties (or what it holds), as read; and for a copy.
    Head: string
    CopyHead: string
    /// The w:pPr as it was read.
    Properties: string option
}

/// The elements whose content the model writes itself (the runs in them are the model's) - a simple field;
/// the others it writes (links, tracked changes) are kept from the source while the model's are what they
/// were (see splicedRuns):
/// links, tracked changes, simple fields. The marks of ranges are carried (carriedMarksOf);
/// equations and paragraph-level objects are runs of the model.
let private modelledWrapper (local: string) =
    // (A simple field is a wrapper as it was read, as a link is: a field the model writes otherwise is written outside it.)
    match local with
    | _ -> false
let private modelledLeaf (local: string) =
    match local with
    | "pPr" | "moveFromRangeStart" | "moveFromRangeEnd" | "moveToRangeStart" | "moveToRangeEnd"
    | "oMath" | "oMathPara" | "AlternateContent" -> true
    | _ -> false

/// `bookmarks`: the ids of the bookmarks the model writes in the paragraph (with its runs): those are left out
/// of the source; the others - a bookmark where no run is there to have it - stay where they stood, as do
/// the ranges of comments (of comments the document still has) and of permissions, which the model does
/// not hold.
/// Whether a shape (a text box, a drawing the model holds as an item of its own) is in an element as read.
let rec private holdsShape (node: XmlDsl.XmlNode) : bool =
    match node with
    | XmlDsl.Element (name, _, kids, _) -> name = "mc:AlternateContent" || name = "w:pict" || XmlPatch.prefix node = "wps" || XmlPatch.prefix node = "wpg" || List.exists holdsShape kids
    | _ -> false

/// Whether a run holds a shape the model holds as an item of its own (a text box, a drawing shape or group:
/// DrawingML wps / wpg, or VML with a text box) - not a picture, which is a run of the model's.
let private holdsShapeItem (run: XmlDsl.XmlNode) : bool =
    not (XmlPatch.descendantsWhere (fun node -> XmlPatch.prefix node = "wps" || XmlPatch.prefix node = "wpg" || XmlPatch.name node = "v:textbox") run).IsEmpty

let private spliceSourceOf (text: string) (start: int) (finish: int) (bookmarks: Set<string>) (keepBookmarks: bool) (keepBreaks: string option) (breaksKept: unit -> unit) : SpliceSource option =
    match XmlDsl.tryScanAt text start finish (System.Collections.Generic.HashSet<string>()) with
    | Some (XmlDsl.Element (_, _, _, span) as paragraph) when XmlDsl.Span.isKnown span && span.StartTagEnd < span.End ->
        let runs = System.Collections.Generic.Dictionary<int, XmlDsl.XmlNode>()
        let order = ResizeArray<int>()
        let properties = XmlPatch.child "pPr" paragraph
        let from = match properties with Some (XmlDsl.Element (_, _, _, propsSpan)) -> propsSpan.End | _ -> span.StartTagEnd
        let take (run: XmlDsl.XmlNode) =
            match run with
            | XmlDsl.Element (_, _, _, runSpan) ->
                runs.[runSpan.Start] <- run
                order.Add runSpan.Start
            | _ -> ()
        let skip (node: XmlDsl.XmlNode) =
            match node with
            | XmlDsl.Element (name, _, _, _) ->
                match XmlDsl.Ns.stripNs name with
                // (The model's bookmarks: written by it - or, in the whole of the paragraph, the source's where they stand.
                // One the model does not hold stays: not every bookmark of a file is one the model has.)
                | "bookmarkStart" | "bookmarkEnd" -> not keepBookmarks && (XmlPatch.attribute node "w:id" |> Option.exists bookmarks.Contains)
                | "commentRangeStart" | "commentRangeEnd" -> XmlPatch.attribute node "w:id" |> Option.exists (fun id -> Set.contains id deletedCommentIds)
                | local -> modelledLeaf local
            | _ -> false
        // (The runs that hold the marks of a field that begins or ends in another paragraph - a table of contents
        // - are the paragraph's as they were: the model does not hold that field.)
        let outerFieldRuns =
            let kept = System.Collections.Generic.HashSet<int>()
            let mutable opened : (int list) list = []   // per open field: the runs of its marks, last first
            for run in XmlPatch.descendantsWhere (fun node -> XmlPatch.name node = "w:r") paragraph do
                let start = (XmlPatch.span run).Start
                let marks = XmlPatch.descendants "fldChar" run |> List.choose (fun mark -> XmlPatch.attribute mark "w:fldCharType")
                let isCode = not (XmlPatch.descendants "instrText" run).IsEmpty
                for mark in marks do
                    match mark, opened with
                    | "begin", _ -> opened <- [ start ] :: opened
                    | "end", _ :: rest -> opened <- rest
                    | "end", [] -> kept.Add start |> ignore
                    | _, current :: rest -> opened <- (start :: current) :: rest
                    | _ -> ()
                if isCode && marks.IsEmpty then
                    match opened with
                    | current :: rest -> opened <- (start :: current) :: rest
                    | [] -> ()
            for runs in opened do for start in runs do kept.Add start |> ignore
            kept
        // (A run that holds a shape anchored in the paragraph is the paragraph's as it was too: the model holds the
        // shape as an item of its own, and a block whose shapes changed is not written from its source.)
        // (Nor are the runs in such a shape - a text box's - the paragraph's.)
        let shapeRuns =
            XmlPatch.elements paragraph
            |> List.collect (fun kid -> XmlPatch.descendantsWhere (fun node -> XmlPatch.name node = "w:r" && holdsShapeItem node) kid |> List.truncate 1)
            |> List.map XmlPatch.span
        let inShape (node: XmlDsl.XmlNode) =
            let span = XmlPatch.span node
            shapeRuns |> List.exists (fun outer -> span.Start >= outer.Start && span.End <= outer.End)
        // (Nor is a run that holds nothing but its properties: the model has no run of it.)
        let holdsContent (node: XmlDsl.XmlNode) = XmlPatch.elements node |> List.exists (fun kid -> XmlPatch.localName kid <> "rPr")
        // (A run that holds only a page or column break, where the model holds the paragraph's break - after it,
        // wherever it stands -: the break as it stands.)
        let breakOnly (node: XmlDsl.XmlNode) =
            match keepBreaks, XmlPatch.elements node |> List.filter (fun kid -> XmlPatch.localName kid <> "rPr" && XmlPatch.localName kid <> "lastRenderedPageBreak") with
            | Some kind, [ only ] when XmlPatch.name only = "w:br" && XmlPatch.attribute only "w:type" = Some kind -> breaksKept (); true
            | _ -> false
        let isItem (node: XmlDsl.XmlNode) =
            XmlPatch.localName node = "r" && holdsContent node && not (outerFieldRuns.Contains (XmlPatch.span node).Start) && not (inShape node) && not (breakOnly node)
        XmlPatch.layout isItem modelledWrapper skip take text paragraph from
        |> Option.map (fun laid ->
            // The fields: each begin mark with the end mark that closes it, through the runs in order.
            let fields = System.Collections.Generic.Dictionary<int, int>()
            let mutable opened : int list = []
            for at in order do
                let run = runs.[at]
                for mark in XmlPatch.descendants "fldChar" run do
                    match XmlPatch.attribute mark "w:fldCharType" with
                    | Some "begin" -> opened <- at :: opened
                    | Some "end" ->
                        match opened with
                        | opener :: rest ->
                            fields.[opener] <- (match run with XmlDsl.Element (_, _, _, runSpan) -> runSpan.End | _ -> at)
                            opened <- rest
                        | [] -> ()
                    | _ -> ()
            let headEnd = match properties with Some (XmlDsl.Element (_, _, _, propsSpan)) -> propsSpan.Start | _ -> span.StartTagEnd
            let between = XmlPatch.slice text span.StartTagEnd headEnd
            { Layout = laid; Runs = runs; Fields = fields; Paragraph = paragraph
              Properties = properties |> Option.map (XmlPatch.bytesOf text)
              Head = XmlPatch.slice text span.Start headEnd
              CopyHead = XmlPatch.startTag text paragraph [ XmlPatch.RemoveAttribute "w14:paraId"; XmlPatch.RemoveAttribute "w14:textId" ] + between })
    | _ -> None

/// The runs of a paragraph written from its source (see above): `pieces` are the model's runs
/// with their XML (the run itself, and the link or tracked change around it).
[<Fable.Core.Emit("JSON.stringify(Object.assign({}, $0, { Text: undefined, RenderedPageBreakBefore: undefined }))")>]
let private jsonWithoutTextOrMark (_run: obj) : string = jsNative
[<Fable.Core.Emit("JSON.stringify(Object.assign({}, $0, { Text: undefined }))")>]
let private jsonWithoutText (_run: obj) : string = jsNative

let private splicedRuns (source: SpliceSource) (original: obj) (joined: obj array) (marksBefore: Map<int, string>) (ownBookmarks: Set<string>) (leading: bool) (trailing: bool) (pieces: (obj * string * (string -> string)) list) (runXml: obj -> string) (wrapOf: obj -> string -> string) : string =
    let text = source.Layout.Text
    // The runs read, by source: the run objects the model had for each.
    let originalRuns = System.Collections.Generic.Dictionary<int, ResizeArray<obj>>()
    for run in Array.append (if isNullOrUndefined original then [||] else original?Runs |> unbox<obj array>) (joined |> Array.collect (fun para -> optionalArray (para?Runs))) do
        match optionalInt (run?SourceRun) with
        | Some at -> (match originalRuns.TryGetValue at with | true, list -> list.Add run | _ -> originalRuns.[at] <- ResizeArray [ run ])
        | None -> ()
    let currentCount = System.Collections.Generic.Dictionary<int, int>()
    for runObj, _, _ in pieces do
        match optionalInt (runObj?SourceRun) with
        | Some at -> currentCount.[at] <- (match currentCount.TryGetValue at with | true, n -> n + 1 | _ -> 1)
        | None -> ()
    // (Where several runs are of one source run - text typed next to it takes its style, and so its source -
    // the one that is that run as it was read is it; the others are new.)
    let asRead = System.Collections.Generic.HashSet<obj>(HashIdentity.Reference)
    for runObj, _, _ in pieces do
        match optionalInt (runObj?SourceRun) with
        | Some at when currentCount.[at] > 1 ->
            match originalRuns.TryGetValue at with
            | true, list when list.Count = 1 && jsonText list.[0] = jsonText runObj -> asRead.Add runObj |> ignore
            | _ -> ()
        | _ -> ()
    // What the model writes of a run: the marks before it (a bookmark's start ...), the run, the marks after it.
    let partsOf (inner: string) =
        match XmlPatch.readFragment inner with
        | Some elements ->
            let isRun (element: XmlPatch.SourceElement) = match element.Node with XmlDsl.Element (name, _, _, _) -> name = "w:r" | _ -> false
            match elements |> List.filter isRun with
            | [ run ] ->
                let index = elements |> List.findIndex isRun
                // (Without the bookmarks the source has where they stand: see spliceSourceOf.)
                let theSource (element: XmlPatch.SourceElement) =
                    (XmlPatch.name element.Node = "w:bookmarkStart" || XmlPatch.name element.Node = "w:bookmarkEnd")
                    && (XmlPatch.attribute element.Node "w:id" |> Option.exists ownBookmarks.Contains)
                let bytes (list: XmlPatch.SourceElement list) = list |> List.filter (theSource >> not) |> List.map XmlPatch.bytes |> String.concat ""
                Some (bytes (List.take index elements), XmlPatch.bytes run, bytes (List.skip (index + 1) elements))
            | _ -> None
        | None -> None
    // (Whether what the model writes of a run is marks with it: a simple field alone - the model's form of a field
    // run - is not.)
    let hasMarks (whole: string) =
        match partsOf whole with
        | Some ("", _, "") -> false
        | Some _ -> true
        | None -> (match XmlPatch.read whole with Some field -> XmlPatch.name field.Node <> "w:fldSimple" | None -> true)
    // (A run that was read whose text and formatting are what they were - however the model holds it, as one
    // run or several (a tab, text, a break), merged where they are formatted alike: its bytes, for the first of
    // them, and nothing for the others.)
    let alike (runs: obj seq) =
        runs
        |> Seq.fold (fun (merged: (string * string) list) (run: obj) ->
            let format, text = jsonWithoutText run, stringOrDefault "" (run?Text)
            match merged with
            | (previous, previousText) :: rest when previous = format -> (format, previousText + text) :: rest
            | _ -> (format, text) :: merged) []
    // (Whether a run that was read is in a simple field.)
    let simpleFields = XmlPatch.descendantsWhere (fun node -> XmlPatch.name node = "w:fldSimple") source.Paragraph |> List.map XmlPatch.span
    let inSimpleField (at: int) = simpleFields |> List.exists (fun span -> at > span.Start && at < span.End)
    let wholeAsRead = System.Collections.Generic.HashSet<int>()
    for KeyValue (at, list) in originalRuns do
        let now = pieces |> List.filter (fun (runObj, _, _) -> optionalInt (runObj?SourceRun) = Some at) |> List.map (fun (runObj, _, _) -> runObj)
        // (With no marks of the model's around them: a bookmark on one of them is written with it.)
        let marked = pieces |> List.exists (fun (runObj, whole, _) -> optionalInt (runObj?SourceRun) = Some at && hasMarks whole)
        // (Not the begin run of a field of several runs: that is written from its begin mark to its end mark. A run of
        // a field is as read only as the one run of a simple field, all of the field - none of its result in parts.)
        let fieldAsRead =
            now |> List.forall (fun runObj ->
                isNullOrUndefined (runObj?FieldCode)
                || (inSimpleField at && not (spliceFieldParts.ContainsKey runObj)))
        if not now.IsEmpty && source.Runs.ContainsKey at && not (source.Fields.ContainsKey at) && fieldAsRead && not marked && alike now = alike list then wholeAsRead.Add at |> ignore
    // (A run that was read whose formatting is what it was and whose text changed - however the model holds it, the
    // mark of a page break as it was rendered on its first character, say -: its text changed in the w:t the
    // change is in, the rest of it as it was; written for the first of the model's runs of it.)
    let textOfWhole = System.Collections.Generic.Dictionary<int, XmlPatch.ElementEdit list>()
    let alikeIgnoringMark (runs: obj seq) =
        runs
        |> Seq.fold (fun (merged: (string * string) list) (run: obj) ->
            let format, text = jsonWithoutTextOrMark run, stringOrDefault "" (run?Text)
            match merged with
            | (previous, previousText) :: rest when previous = format -> (format, previousText + text) :: rest
            | _ -> (format, text) :: merged) []
    for KeyValue (at, list) in originalRuns do
        if not (wholeAsRead.Contains at) && source.Runs.ContainsKey at && not (source.Fields.ContainsKey at) then
            let indexes = pieces |> List.indexed |> List.filter (fun (_, (runObj, _, _)) -> optionalInt (runObj?SourceRun) = Some at) |> List.map fst
            let now = indexes |> List.map (fun index -> let runObj, _, _ = pieces.[index] in runObj)
            let marked = indexes |> List.exists (fun index -> let _, whole, _ = pieces.[index] in hasMarks whole)
            let together = not indexes.IsEmpty && List.last indexes - List.head indexes = indexes.Length - 1
            // (Not of a field: the model writes a field as a whole.)
            let ofField = now |> List.exists (fun runObj -> not (isNullOrUndefined (runObj?FieldCode)))
            let plain (value: string) = ModelText.textPieces value |> List.forall (function ModelText.Text _ -> true | _ -> false)
            match alikeIgnoringMark list, alikeIgnoringMark now with
            | [ (formatWas, was) ], [ (formatNow, nowText) ] when formatWas = formatNow && together && not marked && not ofField && was <> nowText && plain was && plain nowText ->
                let run = source.Runs.[at]
                let content = XmlPatch.elements run |> List.filter (fun kid -> XmlPatch.localName kid <> "rPr")
                if not content.IsEmpty && content |> List.forall (fun kid -> XmlPatch.name kid = "w:t" || XmlPatch.name kid = "w:lastRenderedPageBreak")
                   && content |> List.exists (fun kid -> XmlPatch.name kid = "w:t") then
                    let texts = content |> List.filter (fun kid -> XmlPatch.name kid = "w:t") |> List.map XmlPatch.textOf
                    match ModelText.changeInPieces texts was nowText with
                    | Some (index, value) -> textOfWhole.[at] <- [ XmlPatch.EditNth ("w:t", index, [ XmlPatch.SetText value ]) ]
                    | None -> ()
            | _ -> ()
    let firstOfWhole = System.Collections.Generic.HashSet<int>()
    let placedAt = System.Collections.Generic.HashSet<int>()
    let sourceOfRun (runObj: obj) =
        match optionalInt (runObj?SourceRun) |> Option.filter (fun at -> source.Runs.ContainsKey at) with
        | Some at when currentCount.[at] > 1 && asRead.Count > 0 && (pieces |> List.exists (fun (other, _, _) -> asRead.Contains other && optionalInt (other?SourceRun) = Some at)) ->
            if asRead.Contains runObj then Some (at, true) else None
        | Some at -> Some (at, false)
        | None -> None
    // A new run made like a run that was read (text typed next to it, with its style): that run's start tag and
    // properties - with what the model writes differently of them changed - and the model's content.
    let copyOf (runObj: obj) (inner: string) : string option =
        match optionalInt (runObj?SourceRun) |> Option.filter (fun at -> source.Runs.ContainsKey at) with
        | Some at ->
            match originalRuns.TryGetValue at, XmlPatch.read inner with
            // (Of what the model writes as a run: a field it writes - a simple field around its run - is its own.)
            | (true, list), Some now when list.Count = 1 && XmlPatch.name now.Node = "w:r" && not (boolOrDefault false (runObj?IsImage)) && isNullOrUndefined runObj?OpaqueXml ->
                let run = source.Runs.[at]
                let props = XmlPatch.child "rPr" run |> Option.map (XmlPatch.bytesOf text) |> Option.defaultValue ""
                match mergedProperties "w:rPr" props (modelRunProperties list.[0]) (modelRunProperties runObj), run with
                | Some merged, XmlDsl.Element (_, _, _, runSpan) when runSpan.StartTagEnd < runSpan.End ->
                    // (Without the ids of the editing session the run was made in: the text is new.)
                    let rsids = XmlPatch.attributeNames run |> List.filter ModelText.isRevisionId |> List.map XmlPatch.RemoveAttribute
                    Some (XmlPatch.startTag text run rsids + merged + XmlPatch.childrenBytes now (fun name -> name <> "rPr") + "</w:r>")
                | _ -> None
            | _ -> None
        | None -> None
    let covered = ResizeArray<int * int>()
    // What the model writes around a run (its link, its tracked change), as one text: a run whose wrapping is what
    // it was is written in the wrappers of the source (XmlPatch.layout), one whose wrapping changed outside them,
    // in the model's; one that is new in the wrappers of the run before it, when its wrapping is that one's.
    let wrapping (runObj: obj) = wrapOf runObj "\u0002"
    let mutable previousInSource = false
    let mutable previousWrapping = ""
    // (The runs of a field's result: after the run with its separate mark, to its end mark.)
    let fieldResultRunsOf (at: int) (finish: int) =
        let fieldRuns = source.Runs |> Seq.filter (fun (KeyValue (start, _)) -> start >= at && start < finish) |> Seq.sortBy (fun (KeyValue (start, _)) -> start) |> Seq.map (fun (KeyValue (_, node)) -> node) |> List.ofSeq
        let marks (node: XmlDsl.XmlNode) = XmlPatch.descendants "fldChar" node |> List.choose (fun mark -> XmlPatch.attribute mark "w:fldCharType")
        match fieldRuns |> List.tryFindIndex (fun node -> marks node |> List.contains "separate") with
        | Some index -> fieldRuns |> List.skip (index + 1) |> List.filter (fun node -> not (marks node |> List.contains "end"))
        | None -> []
    let marksWritten = System.Collections.Generic.HashSet<int>()
    let written =
        [| for runObj, whole, wrap in pieces do
           // (The starts of ranges that stood before a paragraph joined to this one: where its text begins.)
           match optionalInt (runObj?SourceRun) |> Option.bind (fun at -> Map.tryFind at marksBefore |> Option.map (fun marks -> at, marks)) with
           | Some (at, marks) when marksWritten.Add at -> yield XmlPatch.Outside, marks
           | _ -> ()
           match optionalInt (runObj?SourceRun) with
           | Some at when wholeAsRead.Contains at ->
               if firstOfWhole.Add at then
                   previousInSource <- true
                   previousWrapping <- wrapping runObj
                   yield XmlPatch.AtItem at, XmlPatch.bytesOf text source.Runs.[at]
           | Some at when textOfWhole.ContainsKey at ->
               if firstOfWhole.Add at then
                   placedAt.Add at |> ignore
                   previousInSource <- true
                   previousWrapping <- wrapping runObj
                   yield XmlPatch.AtItem at, XmlPatch.apply XmlPatch.Word.order { XmlPatch.SourceElement.Text = text; XmlPatch.SourceElement.Node = source.Runs.[at] } textOfWhole.[at]
           | _ ->
             let now = wrapping runObj
             // (The marks the model writes around the run go around what is written of it.)
             let before, inner, after = match partsOf whole with Some parts -> parts | None -> "", whole, ""
             let wrap (xml: string) = if xml = whole then wrap xml else wrap (before + xml + after)
             match sourceOfRun runObj with
             | Some (at, alone) ->
                 let run = source.Runs.[at]
                 // The model's run then, when it was one run of the source and is one now (or is the one of several) -
                 // or when the runs it was are of one format and hold the text it holds (a tab, text, a break): its
                 // formatting alone is what can have changed.
                 let sameText =
                     match originalRuns.TryGetValue at with
                     | true, list when currentCount.[at] = 1 ->
                         (match alike list, alike [ runObj ] with
                          | [ (_, before) ], [ (_, after) ] -> before = after
                          | _ -> false)
                     | _ -> false
                 // (A run split - a part of it formatted -: each part is the run with its own change, and text.)
                 let was =
                     match originalRuns.TryGetValue at with
                     | true, list when list.Count = 1 -> Some list.[0]
                     | true, list when sameText -> Some list.[0]
                     // (A run the model held as several of one format - a tab, text, a break - split: each part the
                     // run with the change of format made in it, and the model's content.)
                     | true, list when (alike list).Length = 1 -> Some list.[0]
                     | _ -> None
                 // (Whether the run is one of several the model held it as, and is not all of them now.)
                 let ofSeveral = not sameText && (match originalRuns.TryGetValue at with | true, list -> list.Count > 1 | _ -> false)
                 ignore alone
                 // (A run of a paragraph joined to this one is not an item of its layout: it is written in the model's wrapping.)
                 let wrappedAsBefore =
                     XmlPatch.isItemOf source.Layout at &&
                     match originalRuns.TryGetValue at with
                     | true, list when list.Count > 0 -> list |> Seq.forall (fun model -> wrapping model = now)
                     | _ -> false
                 // (The parts of a run that was split after the first are in its place after it.)
                 let place, wrap =
                     if wrappedAsBefore then (if placedAt.Add at then XmlPatch.AtItem at else XmlPatch.InItem at), (fun xml -> if xml = whole then xml else before + xml + after)
                     else XmlPatch.Outside, wrap
                 previousInSource <- wrappedAsBefore
                 previousWrapping <- now
                 match was with
                 // A field as it was read: its runs, begin mark to end mark, as they are in the source (the bookmarks
                 // in it are the model's, written around it); what stood among them is in them.
                 | Some model when source.Fields.ContainsKey at && jsonText model = jsonText runObj ->
                     let finish = source.Fields.[at]
                     covered.Add ((at, finish))
                     yield place, wrap (XmlPatch.sliceWithout text source.Paragraph at finish (fun name -> name = "bookmarkStart" || name = "bookmarkEnd"))
                 // A field as it was read but for its formatting: its runs, the result's with the change made in their
                 // properties (see mergedProperties); its code and marks as they were.
                 | Some model when source.Fields.ContainsKey at && stringOrDefault "" (model?Text) = stringOrDefault "" (runObj?Text) && optionalString (model?FieldCode) = optionalString (runObj?FieldCode) ->
                     let finish = source.Fields.[at]
                     covered.Add ((at, finish))
                     let was, now = modelRunProperties model, modelRunProperties runObj
                     // (The runs after the one with the field's separate mark, to its end mark.)
                     let fieldRuns = source.Runs |> Seq.filter (fun (KeyValue (start, _)) -> start >= at && start < finish) |> Seq.sortBy (fun (KeyValue (start, _)) -> start) |> Seq.map (fun (KeyValue (_, node)) -> node) |> List.ofSeq
                     let marks (node: XmlDsl.XmlNode) = XmlPatch.descendants "fldChar" node |> List.choose (fun mark -> XmlPatch.attribute mark "w:fldCharType")
                     let separateAt = fieldRuns |> List.tryFindIndex (fun node -> marks node |> List.contains "separate")
                     let resultRuns =
                         match separateAt with
                         | Some index -> fieldRuns |> List.skip (index + 1) |> List.filter (fun node -> not (marks node |> List.contains "end"))
                         | None -> []
                     let replacements =
                         resultRuns |> List.choose (fun node ->
                             let props = XmlPatch.child "rPr" node |> Option.map (XmlPatch.bytesOf text) |> Option.defaultValue ""
                             mergedProperties "w:rPr" props was now
                             |> Option.map (fun merged ->
                                 let span = match node with XmlDsl.Element (_, _, _, sp) -> sp | _ -> XmlDsl.Span.none
                                 let edit = if merged = "" then XmlPatch.RemoveChild "w:rPr" else XmlPatch.SetChild ("w:rPr", merged)
                                 span.Start, span.End, XmlPatch.apply XmlPatch.Word.order { XmlPatch.SourceElement.Text = text; XmlPatch.SourceElement.Node = node } [ edit ]))
                     if replacements.Length = resultRuns.Length && not resultRuns.IsEmpty then
                         let bytes = XmlPatch.sliceWith text at finish replacements
                         // (The bookmarks in it are the model's, written around it.)
                         let bytes = match XmlPatch.readFragment bytes with
                                     | Some elements -> elements |> List.filter (fun element -> match element.Node with XmlDsl.Element (name, _, _, _) -> name <> "w:bookmarkStart" && name <> "w:bookmarkEnd" | _ -> true) |> List.map XmlPatch.bytes |> String.concat ""
                                     | None -> bytes
                         yield place, wrap bytes
                     else
                         covered.RemoveAt (covered.Count - 1)
                         yield place, wrap inner
                 // A field whose result was typed into (its code as it was): its runs as read, its one result run
                 // written as the model's parts of the result - each that run with the part's formatting and text.
                 | Some model when source.Fields.ContainsKey at && optionalString (model?FieldCode) = optionalString (runObj?FieldCode)
                                   && (fieldResultRunsOf at source.Fields.[at]).Length = 1 ->
                     let finish = source.Fields.[at]
                     let resultRun = (fieldResultRunsOf at finish).Head
                     let props = XmlPatch.child "rPr" resultRun |> Option.map (XmlPatch.bytesOf text) |> Option.defaultValue ""
                     let parts = runObj :: (match spliceFieldParts.TryGetValue runObj with | true, parts -> parts | _ -> [])
                     let written =
                         parts |> List.map (fun part ->
                             mergedProperties "w:rPr" props (modelRunProperties model) (modelRunProperties part)
                             |> Option.map (fun merged ->
                                 XmlPatch.startTag text resultRun [] + merged
                                 + String.concat "" (runTextElements (isDeletedRun part) (stringOrDefault "" (part?Text))) + "</w:r>"))
                     if written |> List.forall Option.isSome then
                         covered.Add ((at, finish))
                         let span = XmlPatch.span resultRun
                         let bytes = XmlPatch.sliceWith text at finish [ span.Start, span.End, written |> List.choose id |> String.concat "" ]
                         let bytes = match XmlPatch.readFragment bytes with
                                     | Some elements -> elements |> List.filter (fun element -> match element.Node with XmlDsl.Element (name, _, _, _) -> name <> "w:bookmarkStart" && name <> "w:bookmarkEnd" | _ -> true) |> List.map XmlPatch.bytes |> String.concat ""
                                     | None -> bytes
                         yield place, wrap bytes
                     else yield place, wrap inner
                 // (Written by the model: outside the wrappers it was in - a simple field the model writes is not in the one it was.)
                 | Some _ when not (isNullOrUndefined (runObj?FieldCode)) -> yield XmlPatch.Outside, wrap inner
                 // As it was read in every respect: its bytes.
                 | Some model when jsonText model = jsonText runObj && not ofSeveral -> yield place, wrap (XmlPatch.bytesOf text run)
                 // Changed in its formatting, its text, or both: the run as it was read with what the model changed
                 // made in it (XmlPatch): its properties merged (see mergedProperties), its text set in its w:t.
                 // Where its children are not of the same names in the same order the model's are written in it.
                 | Some model when not (boolOrDefault false (runObj?IsImage)) && isNullOrUndefined runObj?OpaqueXml ->
                     let props = XmlPatch.child "rPr" run
                     let modelRun = match partsOf (runXml model) with Some (_, run, _) -> run | None -> runXml model
                     match mergedProperties "w:rPr" (props |> Option.map (XmlPatch.bytesOf text) |> Option.defaultValue "") (modelRunProperties model) (modelRunProperties runObj), XmlPatch.read modelRun, XmlPatch.read inner with
                     | Some merged, Some before, Some after ->
                         let source = { XmlPatch.SourceElement.Text = text; XmlPatch.SourceElement.Node = run }
                         let propsEdit = if merged = "" then XmlPatch.RemoveChild "w:rPr" else XmlPatch.SetChild ("w:rPr", merged)
                         let content =
                             if sameText then []
                             else XmlPatch.diff before after |> List.filter (fun edit -> let child = XmlPatch.childOf edit in child <> "" && child <> "w:rPr")
                         // (Content of its own - not the run's as it was - where the run's text changed in a way edits do not
                         // say, or the run is one of several the model held it as: its part of the content is the model's.)
                         // (Or where what the run holds is not of the same kinds in the same order: text typed before a tab -
                         // an edit puts a child in at its schema place, and the order of a run's content is its text.)
                         let contentNames (element: XmlPatch.SourceElement) = XmlPatch.elements element.Node |> List.map XmlPatch.name |> List.filter (fun name -> name <> "w:rPr")
                         let reordered = not sameText && contentNames before <> contentNames after
                         // (The run as read holds its text otherwise than the model writes it - in two w:t, say: the edits of
                         // the model's run are not edits of it. Its text written in the w:t the change is in, where the
                         // change is in one; else the model's content.)
                         let sourceShape = contentNames source
                         let otherShape = not sameText && sourceShape <> contentNames before
                         let textOnly (names: string list) = not names.IsEmpty && names |> List.forall (fun name -> name = "w:t")
                         // (The model's text of the run as read: all of the runs it held it as, when they are of one format.)
                         let wasText = match originalRuns.TryGetValue at with | true, list -> (match alike list with [ (_, whole) ] -> Some whole | _ -> None) | _ -> None
                         let textInPlace =
                             let nowText = stringOrDefault "" (runObj?Text)
                             let plain (value: string) = ModelText.textPieces value |> List.forall (function ModelText.Text _ -> true | _ -> false)
                             match wasText with
                             | Some was when not sameText && currentCount.[at] = 1 && textOnly sourceShape && sourceShape.Length > 1 && plain was && plain nowText ->
                                 let pieces = XmlPatch.elements source.Node |> List.filter (fun kid -> XmlPatch.name kid = "w:t") |> List.map XmlPatch.textOf
                                 ModelText.changeInPieces pieces was nowText
                                 |> Option.map (fun (index, value) -> [ XmlPatch.EditNth ("w:t", index, [ XmlPatch.SetText value ]) ])
                             | _ -> None
                         // (Text typed before a run whose content does not begin with its text - a tab, a break -, in its
                         // formatting: a run of its own before it, made like it; the run as it was.)
                         let typedBefore =
                             match wasText with
                             | Some was when currentCount.[at] = 1 && merged = (props |> Option.map (XmlPatch.bytesOf text) |> Option.defaultValue "")
                                             && (match sourceShape with first :: _ -> first <> "w:t" | [] -> false) ->
                                 ModelText.insertedBefore was (stringOrDefault "" (runObj?Text))
                                 |> Option.filter (fun typed -> ModelText.textPieces typed |> List.forall (function ModelText.Text _ -> true | _ -> false))
                             | _ -> None
                         match textInPlace, typedBefore with
                         | _, Some typed ->
                             let rsids = XmlPatch.attributeNames run |> List.filter ModelText.isRevisionId |> List.map XmlPatch.RemoveAttribute
                             let copy = XmlPatch.startTag text run rsids + merged + String.concat "" (runTextElements false typed) + "</w:r>"
                             yield place, wrap (copy + XmlPatch.bytesOf text run)
                         | Some textEdits, None -> yield place, wrap (XmlPatch.apply XmlPatch.Word.order source (propsEdit :: textEdits))
                         | None, None ->
                         if ofSeveral || reordered || otherShape || content |> List.exists (function XmlPatch.SetChildren _ -> true | _ -> false) then
                             let runSpan = match run with XmlDsl.Element (_, _, _, sp) -> sp | _ -> XmlDsl.Span.none
                             let modelContent = XmlPatch.childrenBytes after (fun name -> name <> "rPr")
                             if runSpan.StartTagEnd < runSpan.End then
                                 yield place, wrap (XmlPatch.slice text runSpan.Start runSpan.StartTagEnd + merged + modelContent + XmlPatch.slice text runSpan.EndTagStart runSpan.End)
                             else yield place, wrap inner
                         else yield place, wrap (XmlPatch.apply XmlPatch.Word.order source (propsEdit :: content))
                     | _ -> yield place, wrap inner
                 | _ -> yield place, wrap inner
             | None ->
                 let run = copyOf runObj inner |> Option.defaultValue inner
                 if previousInSource && now = previousWrapping then yield XmlPatch.WithPrevious, before + run + after
                 else
                     previousInSource <- false
                     yield XmlPatch.Outside, wrap run |]
    XmlPatch.writeLayoutWith { KeptBefore = leading; KeptAfter = trailing; Covered = List.ofSeq covered } source.Layout written


paragraphSplicer <-
    { new ParagraphSplicer with
        member _.Splice para pieces runXml wrapOf =
            if isNullOrUndefined (para?SpliceFrom) || isNullOrUndefined (para?SpliceOriginal) || spliceSourceText = "" then None
            else
                // (The bookmarks the model writes: in what it writes of the runs.)
                let bookmarks =
                    [ for _, whole, _ in pieces do
                        match XmlPatch.readFragment whole with
                        | Some elements ->
                            for element in elements do
                                for name in [ "bookmarkStart"; "bookmarkEnd" ] do
                                    for mark in XmlPatch.descendants name element.Node do
                                        yield! Option.toList (XmlPatch.attribute mark "w:id")
                        | None -> () ] |> Set.ofList
                // (Only in the whole of the paragraph: a part of one, or one with others joined to it, has its break where
                // the model says - the source's may stand in another part.)
                let whole = not (para?SpliceLeading = box false) && not (para?SpliceTrailing = box false) && optionalArray (para?SpliceJoined) |> Array.isEmpty
                let keepBreaks =
                    if whole && boolOrDefault false (para?PageBreakAfter) then Some (if boolOrDefault false (para?ColumnBreakAfter) then "column" else "page") else None
                para?SpliceBreakKept <- false
                spliceSourceOf spliceSourceText (intOrDefault -1 (para?SpliceFrom?Start)) (intOrDefault -1 (para?SpliceFrom?End)) bookmarks whole keepBreaks (fun () -> para?SpliceBreakKept <- true)
                |> Option.map (fun source ->
                    let flag (value: obj) = isNullOrUndefined value || unbox<bool> value
                    // The runs of the paragraphs joined to it, in the elements they were read from.
                    let joined = optionalArray (para?SpliceJoined)
                    // (And the starts of the ranges that stood between it and the paragraph before it: before its first run.)
                    let mutable marksBefore = Map.empty
                    let mutable previousEnd = intOrDefault -1 (para?SpliceFrom?End)
                    for other in joined do
                        match paragraphSource.SpanOf other with
                        | Some (start, finish) ->
                            match XmlDsl.tryScanAt spliceSourceText start finish (System.Collections.Generic.HashSet<string>()) with
                            | Some node ->
                                let runs = XmlPatch.descendantsWhere (fun inner -> XmlPatch.name inner = "w:r") node
                                for run in runs do source.Runs.[(XmlPatch.span run).Start] <- run
                                if previousEnd >= 0 && previousEnd <= start && not runs.IsEmpty then
                                    match XmlPatch.readFragmentRoot (XmlPatch.slice spliceSourceText previousEnd start) with
                                    | Some between ->
                                        // (All of them, in their order: a bookmark that began and ended there is where the text begins.)
                                        let starts = rangeMarksIn carriedRangeKinds between.Node |> List.filter (fun mark -> mark.Kind <> "commentRange" || Set.contains mark.Id liveCommentIds) |> List.map (fun mark -> XmlPatch.bytesOf between.Text mark.Node)
                                        if not starts.IsEmpty then marksBefore <- Map.add (XmlPatch.span runs.Head).Start (String.concat "" starts) marksBefore
                                    | None -> ()
                            | None -> ()
                            previousEnd <- finish
                        | None -> previousEnd <- -1
                    // (The bookmarks the model writes that the source has, in the whole of the paragraph: the source's.)
                    let ownBookmarks =
                        if not whole then Set.empty
                        else
                            XmlPatch.descendantsWhere (fun node -> XmlPatch.name node = "w:bookmarkStart" || XmlPatch.name node = "w:bookmarkEnd") source.Paragraph
                            |> List.choose (fun node -> XmlPatch.attribute node "w:id")
                            |> List.filter bookmarks.Contains
                            |> Set.ofList
                    [ splicedRuns source para?SpliceOriginal joined marksBefore ownBookmarks (flag para?SpliceLeading) (flag para?SpliceTrailing) pieces runXml wrapOf ], source.Properties,
                    (if boolOrDefault false (para?SourceCopy) then source.CopyHead else source.Head))
        member _.MergeProperties source was now =
            mergedProperties "w:pPr" source (propertyElementsOf was) (propertyElementsOf now) }

/// What the last save made of the body (see Main's getLastSaveReport).
let mutable private reportInline = 0
let mutable private reportModelBlocks = 0
let mutable private reportModelParagraphs = 0
let mutable private reportSourceBlocks = 0
/// Why blocks and paragraphs were written from the model (the first few).
let private reportReasons = ResizeArray<string>()
let private noteReason (reason: string) = if reportReasons.Count < 400 then reportReasons.Add reason
let lastSaveReport () : obj =
    createObj [
        "Reasons" ==> reportReasons.ToArray()
        "DerivedParagraphs" ==> reportDerived
        "GeneratedParagraphs" ==> reportGenerated
        "SplicedParagraphs" ==> reportSpliced
        "PatchedParagraphs" ==> reportPatched
        "ModelParagraphsInTables" ==> reportInline
        "ModelBlocks" ==> reportModelBlocks
        "ModelBlockParagraphs" ==> reportModelParagraphs
        "SourceBlocks" ==> reportSourceBlocks ]


// ---- a paragraph that was split, and paragraphs that were joined ---------------
//
// A paragraph split in two (Enter) is, in the file, two paragraphs made of
// the one that was read: each has the paragraph's properties and the runs,
// bookmarks, fields and links of its part of the text - the run the split is
// in is split. Paragraphs joined (Backspace at the start of the second) are
// the first with the content of the second at its end. Neither is written
// from the model.

/// The XML of a paragraph of the body as it was read: by the index of its block among the
/// items that were read, and which of the block's paragraphs it is. (Set while a body is
/// written of which the XML is known: see planBody.)
let mutable private sourceBodyXml : string = ""
/// Per element of the body: where its text starts and ends; per item that was read: the element
/// it came from, and how many paragraphs of the body's text it holds.
let mutable private sourceChunkStarts : int array = [||]
/// (Where the element itself starts: what stands before it, after the element before, is between the two.)
let mutable private sourceChunkElementStarts : int array = [||]
let mutable private sourceChunkEnds : int array = [||]
let mutable private sourceItemChunks : int array = [||]
let mutable private sourceStoryCounts : int array = [||]
/// The body part as it was read, as a span tree: its paragraphs, tables, rows and cells by where each
/// starts - the origin (SourceOrigin, SourceAt) the model has of each. Made when first asked for.
let mutable private sourceElements : System.Collections.Generic.Dictionary<int, XmlDsl.XmlNode> option = None
/// The paragraphs, tables, rows and cells of a span tree by where each starts.
let private indexOfElements (root: XmlDsl.XmlNode) =
    let index = System.Collections.Generic.Dictionary<int, XmlDsl.XmlNode>()
    for node in XmlPatch.descendantsWhere (fun node -> List.contains (XmlPatch.localName node) [ "p"; "tbl"; "tr"; "tc" ]) root do
        if XmlDsl.Span.isKnown (XmlPatch.span node) then index.[(XmlPatch.span node).Start] <- node
    index
/// Origins name elements of the body part: they are looked up only while the body is written.
let mutable private writingBody = false
let private sourceElementAt (at: int) : XmlDsl.XmlNode option =
    if not writingBody || sourceBodyXml = "" || at < 0 then None
    else
        let index =
            match sourceElements with
            | Some index -> index
            | None ->
                let index = XmlDsl.tryScan sourceBodyXml (System.Collections.Generic.HashSet<string>()) |> Option.map indexOfElements |> Option.defaultValue (System.Collections.Generic.Dictionary())
                sourceElements <- Some index
                index
        match index.TryGetValue at with
        | true, node -> Some node
        | _ -> None
/// The element a paragraph of the model is of (its SourceAt), as it was read.
let private sourceParagraphOf (para: obj) : XmlDsl.XmlNode option =
    if isNullOrUndefined para then None
    else
        match optionalInt (para?SourceAt) |> Option.bind sourceElementAt with
        | Some (XmlDsl.Element (name, _, _, _) as node) when name = "w:p" -> Some node
        | _ -> None
/// Whether a paragraph as it was read is tied to what is around it, so that it cannot be written on its
/// own: a shape is anchored in it (written with its block), or a field begins in it and ends in another.
let private paragraphIsBound (node: XmlDsl.XmlNode) : bool =
    let marks = XmlPatch.descendants "fldChar" node |> List.choose (fun mark -> XmlPatch.attribute mark "w:fldCharType")
    holdsShape node || (marks |> List.filter ((=) "begin")).Length <> (marks |> List.filter ((=) "end")).Length
/// What is known of the part being written as it was read (see planBody): a header's plan is made with
/// the body's, and each is put back while its part is written.
type private SourceState = {
    Xml: string; Elements: System.Collections.Generic.Dictionary<int, XmlDsl.XmlNode> option
    ChunkStarts: int array; ChunkElementStarts: int array; ChunkEnds: int array; ItemChunks: int array; StoryCounts: int array }
let private sourceState () : SourceState =
    { Xml = sourceBodyXml; Elements = sourceElements; ChunkStarts = sourceChunkStarts; ChunkElementStarts = sourceChunkElementStarts
      ChunkEnds = sourceChunkEnds; ItemChunks = sourceItemChunks; StoryCounts = sourceStoryCounts }
let private useSourceState (state: SourceState) =
    sourceBodyXml <- state.Xml
    sourceElements <- state.Elements
    sourceChunkStarts <- state.ChunkStarts
    sourceChunkElementStarts <- state.ChunkElementStarts
    sourceChunkEnds <- state.ChunkEnds
    sourceItemChunks <- state.ItemChunks
    sourceStoryCounts <- state.StoryCounts
let private forgetSourceBody () =
    sourceElements <- None
    sourceBodyXml <- ""
    sourceChunkStarts <- [||]
    sourceChunkElementStarts <- [||]
    sourceChunkEnds <- [||]
    sourceItemChunks <- [||]
    sourceStoryCounts <- [||]
paragraphSource <-
    { new ParagraphSource with
        // (Nothing was done to it: the bytes it was read as.)
        member _.Of (para: obj) =
            if not (boolOrDefault false (para?SourcePara?Verbatim)) then None
            else sourceParagraphOf para |> Option.filter (paragraphIsBound >> not) |> Option.map (XmlPatch.bytesOf sourceBodyXml)
        member _.SpanOf (para: obj) =
            sourceParagraphOf para
            |> Option.filter (paragraphIsBound >> not)
            |> Option.map (fun node -> match node with XmlDsl.Element (_, _, _, span) -> span.Start, span.End | _ -> -1, -1) }

paragraphOrigin <-
    { new ParagraphOrigin with
        member _.Of (para: obj) = sourceParagraphOf para |> Option.map (XmlPatch.bytesOf sourceBodyXml) }

/// The w:tbl of an element of a body: the table itself, or the one element of a content control or custom XML that is.
let rec private tableElementIn (node: XmlDsl.XmlNode) : XmlDsl.XmlNode option =
    let kidsOf (node: XmlDsl.XmlNode) = XmlPatch.elements node
    match XmlPatch.localName node with
    | "tbl" -> Some node
    | "sdt" ->
        (match kidsOf node |> List.filter (fun kid -> XmlPatch.localName kid = "sdtContent") with
         | [ content ] -> (match kidsOf content with [ one ] -> tableElementIn one | _ -> None)
         | _ -> None)
    | "customXml" -> (match kidsOf node |> List.filter (fun kid -> XmlPatch.localName kid <> "customXmlPr") with [ one ] -> tableElementIn one | _ -> None)
    | _ -> None

/// A w:tbl of `xml` taken apart (see TableSource), with what stands around it in its block (`wrapStart`, `wrapEnd`).
let private sourceTableOf (xml: string) (tableNode: XmlDsl.XmlNode) (wrapStart: string) (wrapEnd: string) : SourceTableText option =
    let local (node: XmlDsl.XmlNode) = match node with XmlDsl.Element (name, _, _, _) -> XmlDsl.Ns.stripNs name | _ -> ""
    let spanOf (node: XmlDsl.XmlNode) = match node with XmlDsl.Element (_, _, _, sp) -> sp | _ -> XmlDsl.Span.none
    let kidsOf (node: XmlDsl.XmlNode) = match node with XmlDsl.Element (_, _, kids, _) -> kids |> List.filter (fun kid -> local kid <> "") | _ -> []
    let slice = XmlPatch.slice xml
    match tableNode with
    | XmlDsl.Element (name, _, children, span) when XmlDsl.Span.isKnown span && name = "w:tbl" ->
        let rows = System.Collections.Generic.Dictionary<int, SourceRowText>()
        let cells = System.Collections.Generic.Dictionary<int, SourceCellText>()
        let mutable sound = true
        let partOf (nodes: XmlDsl.XmlNode list) (name: string) =
            nodes |> List.tryFind (fun kid -> local kid = name) |> Option.map (XmlPatch.bytesOf xml) |> Option.defaultValue ""
        // Where the leading children named `names` of an element end (its start tag's end when it has none).
        let headEndOf (sp: XmlDsl.Span) (nodes: XmlDsl.XmlNode list) (names: string list) =
            kidsOf (XmlDsl.Element ("", Map.empty, nodes, XmlDsl.Span.none))
            |> List.takeWhile (fun kid -> List.contains (local kid) names)
            |> List.fold (fun at kid -> let kidSpan = spanOf kid in if XmlDsl.Span.isKnown kidSpan then kidSpan.End else at) sp.StartTagEnd
        let laid (isItem: string -> bool) (take: XmlDsl.XmlNode -> unit) (node: XmlDsl.XmlNode) (from: int) =
            match XmlPatch.layout (XmlPatch.localName >> isItem) (fun _ -> false) (fun _ -> false) take xml node from with
            | Some found -> found
            | None ->
                sound <- false
                { XmlPatch.Layout.Text = xml; Wrappers = System.Collections.Generic.Dictionary(); Children = []; Leads = System.Collections.Generic.Dictionary(); KeptEnds = System.Collections.Generic.Dictionary(); Tail = (0, 0) }
        let copyTag (node: XmlDsl.XmlNode) = XmlPatch.startTag xml node [ XmlPatch.RemoveAttribute "w14:paraId"; XmlPatch.RemoveAttribute "w14:textId" ]
        let takeCell (node: XmlDsl.XmlNode) =
            let sp = spanOf node
            if sp.StartTagEnd >= sp.End then sound <- false
            else
                let kids = match node with XmlDsl.Element (_, _, kids, _) -> kids | _ -> []
                let headEnd = headEndOf sp kids [ "tcPr" ]
                cells.[sp.Start] <- { CellXml = XmlPatch.bytesOf xml node
                                      CellStartTag = slice sp.Start sp.StartTagEnd
                                      CopyStartTag = copyTag node
                                      CellProps = partOf kids "tcPr"
                                      CellHead = slice sp.Start headEnd
                                      CellChildren = laid (fun name -> name = "p" || name = "tbl") ignore node headEnd }
        let takeRow (node: XmlDsl.XmlNode) =
            let sp = spanOf node
            if sp.StartTagEnd >= sp.End then sound <- false
            else
                let kids = match node with XmlDsl.Element (_, _, kids, _) -> kids | _ -> []
                let headEnd = headEndOf sp kids [ "tblPrEx"; "trPr" ]
                let inner = laid (fun name -> name = "tc") takeCell node headEnd
                rows.[sp.Start] <- { RowXml = XmlPatch.bytesOf xml node
                                     RowStartTag = slice sp.Start sp.StartTagEnd
                                     CopyStartTag = copyTag node
                                     RowExceptions = partOf kids "tblPrEx"
                                     RowProps = partOf kids "trPr"
                                     RowHead = slice sp.Start headEnd
                                     RowChildren = inner }
        let headEnd = headEndOf span children [ "tblPr"; "tblGrid" ]
        let tableChildren = laid (fun name -> name = "tr") takeRow tableNode headEnd
        if not sound then None
        else Some { WrapStart = wrapStart; WrapEnd = wrapEnd
                    TableHead = slice span.Start headEnd; TableStartTag = slice span.Start span.StartTagEnd; TableProps = partOf children "tblPr"; TableGrid = partOf children "tblGrid"
                    TableRows = rows; TableCells = cells; TableChildren = tableChildren; TableNode = tableNode }
    | _ -> None

// The table a block came from, taken apart (see TableSource): a table whose element is the block's alone,
// or the one element in a content control or custom XML that is.
tableSource <-
    { new TableSource with
        member _.Of (spec: obj) =
            let block = intOrDefault -1 (spec?Block)
            let chunk = if block >= 0 && block < sourceItemChunks.Length then sourceItemChunks.[block] else -1
            if chunk < 0 || chunk >= sourceChunkStarts.Length || chunk >= sourceChunkElementStarts.Length || (sourceItemChunks |> Array.filter (fun other -> other = chunk)).Length <> 1 then None
            else
                let xml = sourceBodyXml
                let elementStart = sourceChunkElementStarts.[chunk]
                // (What stands before the block's element - white space alone here: marks there are the plan's.)
                let lead = if elementStart >= sourceChunkStarts.[chunk] then XmlPatch.slice xml sourceChunkStarts.[chunk] elementStart else "x"
                if lead.Trim() <> "" || elementStart >= sourceChunkEnds.[chunk] then None
                else
                    match XmlDsl.tryScanAt xml elementStart sourceChunkEnds.[chunk] (System.Collections.Generic.HashSet<string>()) |> Option.bind tableElementIn with
                    | Some tableNode when XmlDsl.Span.isKnown (XmlPatch.span tableNode) ->
                        let span = XmlPatch.span tableNode
                        sourceTableOf xml tableNode (XmlPatch.slice xml sourceChunkStarts.[chunk] span.Start) (XmlPatch.slice xml span.End sourceChunkEnds.[chunk])
                    | _ -> None
        member _.At (tableNode: XmlDsl.XmlNode) = sourceTableOf sourceBodyXml tableNode "" ""
        member _.Merge (container: string, source: string, was: string list, now: string list) =
            mergedProperties container source was now }

// ---- a table whose rows or columns came or went --------------------------------
//
// A table that was read and changed in more than its paragraphs - a row or a column put in or taken out, a
// cell's properties - is its element as it was read with those changes made in it (XmlPatch): a row or a cell
// that is gone is taken out, the marks of ranges in it (bookmarks, comment and permission ranges) standing
// where it stood; one that is new is put in next to the one before it (made like the one it was made like:
// see generateTableRow, generateTableCell); one that changed is written from its source; the grid has the
// columns put in or taken out; every other row and cell is the bytes it was read as.

/// The marks of ranges (bookmarks, comment and permission ranges) in an element that is taken out: they stay
/// where it stood. (Not the range of a comment the document has no more.)
let private marksOutOf (xml: string) (node: XmlDsl.XmlNode) : string =
    rangeMarksIn carriedRangeKinds node
    |> List.filter (fun mark -> mark.Kind <> "commentRange" || Set.contains mark.Id liveCommentIds)
    |> List.map (fun mark -> XmlPatch.bytesOf xml mark.Node)
    |> String.concat ""

/// The edit that takes an element out: the marks of ranges in it in its place, or nothing.
let private takenOut (xml: string) (node: XmlDsl.XmlNode) : XmlPatch.ElementEdit =
    match marksOutOf xml node with
    | "" -> XmlPatch.Removed
    | marks -> XmlPatch.Replaced marks

/// The items of a layout (where each starts) and the wrappers each is in.
let private layoutItems (laid: XmlPatch.Layout) : (int * int list) list =
    laid.Children |> List.choose (function XmlPatch.LaidItem (at, wrappers) -> Some (at, wrappers) | XmlPatch.LaidKept _ -> None)

/// The rows of a table of the model that are rows of the table as it was read (not copies): where each starts.
let private sourceRowsOf (known: SourceTableText) (table: obj) : int list =
    optionalArray (table?Rows)
    |> Array.toList
    |> List.choose (fun row -> match sourceOfPart (row?SourceRow) with Some (at, false, _) when known.TableRows.ContainsKey at -> Some at | _ -> None)

/// Whether a table of the model can be written as its element as read with its changes made in it: some of its
/// rows are rows that were read, in the order they were read in.
let private tableIsPatchable (known: SourceTableText) (table: obj) : bool =
    let rows = sourceRowsOf known table
    not rows.IsEmpty && (rows |> List.pairwise |> List.forall (fun (a, b) -> a < b))

/// What is taken out of a layout (of rows, of cells) where the items `kept` stay: each item that does not stay, or the
/// outermost of its wrappers that holds none that stays (a content control around rows: a repeating section's item) -
/// where each starts. (`known`: the wrappers that may be taken out - a content control, custom XML.)
let private takenOutOf (laid: XmlPatch.Layout) (known: int -> bool) (kept: int -> bool) : int list =
    let items = layoutItems laid
    items
    |> List.filter (fun (at, _) -> not (kept at))
    |> List.map (fun (at, wrappers) ->
        wrappers
        |> List.tryFind (fun wrapper -> known wrapper && items |> List.forall (fun (other, around) -> not (List.contains wrapper around) || not (kept other)))
        |> Option.defaultValue at)
    |> List.distinct

/// The table written as its element as it was read with the model's changes made in it (see tableIsPatchable).
let private patchedTableXml (hyperlinks: Map<string, string>) (images: ImageScope) (known: SourceTableText) (table: obj) : string =
    let xml = sourceBodyXml
    let tableNode = known.TableNode
    // (The rows and cells of the table and the content controls and custom XML around them, by where each starts.)
    let nodes = System.Collections.Generic.Dictionary<int, XmlDsl.XmlNode>()
    for node in XmlPatch.descendantsWhere (fun node -> List.contains (XmlPatch.localName node) [ "tr"; "tc"; "sdt"; "customXml" ]) tableNode do
        nodes.[(XmlPatch.span node).Start] <- node
    let targets = System.Collections.Generic.Dictionary<int, XmlDsl.XmlNode * ResizeArray<XmlPatch.ElementEdit>>()
    let edit (node: XmlDsl.XmlNode) (change: XmlPatch.ElementEdit) =
        let at = (XmlPatch.span node).Start
        match targets.TryGetValue at with
        | true, (_, edits) -> edits.Add change
        | _ -> targets.[at] <- (node, ResizeArray [ change ])
    // (A child of an element by its local name, with the name it has there.)
    let childNamed (node: XmlDsl.XmlNode) (local: string) = XmlPatch.elements node |> List.tryFind (fun kid -> XmlPatch.localName kid = local)
    let qualified (node: XmlDsl.XmlNode) (local: string) = childNamed node local |> Option.map XmlPatch.name |> Option.defaultValue ("w:" + local)
    // A properties element that is now `now` where it was `was` (as read): set, taken out, or left.
    let setProperties (node: XmlDsl.XmlNode) (local: string) (was: string) (now: string) : XmlPatch.ElementEdit list =
        if now = was then []
        elif now = "" then [ XmlPatch.RemoveChild (qualified node local) ]
        else [ XmlPatch.SetChild (qualified node local, now) ]
    let outerContext = paragraphsInContext
    let outerTable = currentSourceTable
    paragraphsInContext <- true
    currentSourceTable <- Some known
    // The table's properties: those it was read with, changed where the model's changed.
    let before = table?SourceTable?Before
    if not (isNullOrUndefined before) then
        let now = tablePropertyElements table
        match tableSource.Merge ("w:tblPr", known.TableProps, tablePropertyElements before, now) with
        | Some props ->
            for change in setProperties tableNode "tblPr" known.TableProps (if props = "" then "<w:tblPr/>" else props) do edit tableNode change
        | None -> edit tableNode (XmlPatch.SetChild (qualified tableNode "tblPr", element "w:tblPr" [] now))
    // The grid: a column put in like its neighbour, one taken out; otherwise the model's widths.
    if not (sourceGridIs known.TableGrid table) then
        let now = optionalArray (table?GridColumns) |> Array.choose (fun value -> optionalFloat value) |> Array.map roundToInt
        match childNamed tableNode "tblGrid" with
        | None -> edit tableNode (XmlPatch.SetChild ("w:tblGrid", generateTableGrid table))
        | Some grid ->
            let columns = XmlPatch.elements grid |> List.filter (fun node -> XmlPatch.localName node = "gridCol") |> List.toArray
            let widths = columns |> Array.map (fun column -> XmlPatch.attribute column "w:w" |> Option.bind (fun value -> match System.Double.TryParse value with | true, width -> Some (roundToInt width) | _ -> None))
            let near (a: int) (b: int) = abs (a - b) <= 1
            let read = widths |> Array.choose id
            // (The place k at which `longer` has a column `shorter` has not.)
            let extraAt (longer: int array) (shorter: int array) =
                [ 0 .. shorter.Length ]
                |> List.tryFind (fun k -> [ 0 .. shorter.Length - 1 ] |> List.forall (fun i -> near shorter.[i] longer.[if i < k then i else i + 1]))
            let width (column: XmlDsl.XmlNode) (value: int) =
                XmlPatch.apply XmlPatch.Word.order { Text = xml; Node = column } [ XmlPatch.SetAttribute ("w:w", string value) ]
            let whole () = edit grid (XmlPatch.Replaced (generateTableGrid table))
            if read.Length <> widths.Length || columns.Length = 0 || now.Length = 0 then whole ()
            elif now.Length = read.Length + 1 then
                match extraAt now read with
                | Some k when k > 0 -> edit columns.[k - 1] (XmlPatch.InsertAfter (if near read.[k - 1] now.[k] then XmlPatch.bytesOf xml columns.[k - 1] else width columns.[k - 1] now.[k]))
                | Some _ -> edit columns.[0] (XmlPatch.InsertBefore (if near read.[0] now.[0] then XmlPatch.bytesOf xml columns.[0] else width columns.[0] now.[0]))
                | None -> whole ()
            elif now.Length = read.Length - 1 then
                match extraAt read now with
                | Some k -> edit columns.[k] XmlPatch.Removed
                | None -> whole ()
            elif now.Length = read.Length then
                for i in 0 .. read.Length - 1 do
                    if not (near read.[i] now.[i]) then edit columns.[i] (XmlPatch.SetAttribute ("w:w", string now.[i]))
            else whole ()
    // The items of a row or of the table in the model's order - each one that was read (where it starts) or a new
    // one (its XML, and where the one it was made like starts) -: the new ones put in among those that were read.
    // A run of new ones goes before the item that follows it where each was made like that one (a row put in
    // above it), else after the item before it (in the wrappers of that one), else before the first.
    let placeNew (entries: Choice<int, string * int option> list) =
        let rec walk (previous: int option) (entries: Choice<int, string * int option> list) =
            match entries with
            | [] -> ()
            | Choice1Of2 at :: rest -> walk (Some at) rest
            | Choice2Of2 _ :: _ ->
                let run = entries |> List.takeWhile (function Choice2Of2 _ -> true | Choice1Of2 _ -> false) |> List.choose (function Choice2Of2 entry -> Some entry | _ -> None)
                let rest = entries |> List.skip run.Length
                let next = match rest with Choice1Of2 at :: _ -> Some at | _ -> None
                let text = run |> List.map fst |> String.concat ""
                match previous, next with
                | _, Some at when previous.IsNone || run |> List.forall (fun (_, like) -> like = Some at) -> edit nodes.[at] (XmlPatch.InsertBefore text)
                | Some at, _ -> edit nodes.[at] (XmlPatch.InsertAfter text)
                | None, None -> ()
                walk previous rest
        walk None entries
    // The rows.
    let keptRows = Set.ofList (sourceRowsOf known table)
    for at in takenOutOf known.TableChildren nodes.ContainsKey keptRows.Contains do
        match nodes.TryGetValue at with
        | true, node -> edit node (takenOut xml node)
        | _ -> ()
    let likeOf (part: obj) = sourceOfPart part |> Option.map (fun (at, _, _) -> at)
    [ for row in optionalArray (table?Rows) do
        match sourceOfPart (row?SourceRow) with
        | Some (at, false, same) when keptRows.Contains at && nodes.ContainsKey at ->
            let rowNode = nodes.[at]
            if not same then
                let knownRow = known.TableRows.[at]
                // Its properties, changed where the model's changed.
                let head =
                    let rowBefore = row?SourceRow?Before
                    if isNullOrUndefined rowBefore then Some []
                    else
                        let exceptions, props = tableRowProperties row
                        let exceptionsBefore, propsBefore = tableRowProperties rowBefore
                        tableSource.Merge ("w:trPr", knownRow.RowProps, propsBefore, props)
                        |> Option.map (fun merged ->
                            setProperties rowNode "trPr" knownRow.RowProps merged
                            @ (if exceptionsBefore = exceptions then [] else setProperties rowNode "tblPrEx" knownRow.RowExceptions exceptions))
                // Its cells: those that were read in the order they were read in, each where it stood.
                let cells = optionalArray (row?Cells)
                let keptCells =
                    cells |> Array.toList |> List.choose (fun cell ->
                        match sourceOfPart (cell?SourceCell) with
                        | Some (cellAt, false, _) when XmlPatch.isItemOf knownRow.RowChildren cellAt && nodes.ContainsKey cellAt -> Some cellAt
                        | _ -> None)
                let cellsInPlace = not keptCells.IsEmpty && (keptCells |> List.pairwise |> List.forall (fun (a, b) -> a < b))
                match head with
                | Some headEdits when cellsInPlace ->
                    for change in headEdits do edit rowNode change
                    for cellAt in takenOutOf knownRow.RowChildren nodes.ContainsKey (fun cellAt -> List.contains cellAt keptCells) do
                        match nodes.TryGetValue cellAt with
                        | true, node -> edit node (takenOut xml node)
                        | _ -> ()
                    placeNew
                        [ for cell in cells do
                            match sourceOfPart (cell?SourceCell) with
                            | Some (cellAt, false, cellSame) when List.contains cellAt keptCells ->
                                if not cellSame then edit nodes.[cellAt] (XmlPatch.Replaced (generateTableCell hyperlinks images cell))
                                yield Choice1Of2 cellAt
                            | _ -> yield Choice2Of2 (generateTableCell hyperlinks images cell, likeOf (cell?SourceCell)) ]
                | _ -> edit rowNode (XmlPatch.Replaced (generateTableRow hyperlinks images row))
            yield Choice1Of2 at
        | _ ->
            // A row that is new.
            yield Choice2Of2 (generateTableRow hyperlinks images row, likeOf (row?SourceRow)) ]
    |> placeNew
    paragraphsInContext <- outerContext
    currentSourceTable <- outerTable
    let edits = XmlPatch.editsAt tableNode [ for KeyValue (_, (node, changes)) in targets -> node, List.ofSeq changes ]
    XmlPatch.apply XmlPatch.Word.order { Text = xml; Node = tableNode } edits

/// The paragraphs of the items of an element of the body (from `start` to `finish` in the part) that were
/// changed, to be written from their source (see splicedRuns) in the place of their elements (their
/// SourceAt): where each element starts and ends, and the paragraph. None when one of them cannot be: its
/// element is not known, is not in this one, or is bound to what is around it (a shape anchored in it, a
/// field that begins or ends in another), or it cannot be spliced - in a paragraph of the body itself the
/// whole element is then written from the model; in a table, the table.
let private patchSourceChunk (start: int) (finish: int) (items: obj array) : (int * int * obj) list option =
    let inlined = ResizeArray<int * int * obj>()
    let mutable ok = true
    for item in items do
        let changes = item?SourceTextChanges
        if ok && not (isNullOrUndefined changes) then
            for change in optionalArray (changes?Edits) do
                if ok then
                    let after = change?After
                    let index = intOrDefault -1 (change?Paragraph)
                    match (if isNullOrUndefined after then None else sourceParagraphOf after) with
                    | Some (XmlDsl.Element (_, _, _, span) as node) when span.Start >= start && span.End <= finish ->
                        // (A field that begins or ends in another paragraph, a shape anchored in it: not in the way - their
                        // runs stay as they were, see spliceSourceOf. Another kind of shape is.)
                        let otherShape = XmlPatch.descendantsWhere (fun inner -> XmlPatch.name inner = "w:r" && holdsShape inner && not (holdsShapeItem inner)) node
                        if not otherShape.IsEmpty then
                            noteReason (sprintf "paragraph %d of the block: a shape the model does not hold as an item is in it: the block is written from the model" index)
                            ok <- false
                        elif isNullOrUndefined (after?SpliceOriginal) then
                            noteReason (sprintf "paragraph %d of the block: %s" index (optionalString (change?Why) |> Option.defaultValue "changed, and not of a paragraph that was read"))
                            ok <- false
                        else
                            // (The marks of ranges its XML holds go with it: see carriedMarksOf.)
                            after?CarrySource <- XmlPatch.bytesOf sourceBodyXml node
                            after?CarrySave <- saveNumber
                            after?SpliceFrom <- createObj [ "Start" ==> span.Start; "End" ==> span.End ]
                            setSpliceCut change after
                            inlined.Add ((span.Start, span.End, after))
                    | _ ->
                        noteReason (sprintf "paragraph %d of the block: its element as it was read is not known" index)
                        ok <- false
    // (Two paragraphs of one element - the place a vertically merged cell's continuation keeps, made of the
    // paragraph moved out of it -: the first is written in its place, the other leaves the element as it is.
    // Paragraphs of elements one inside the other: not written from their source.)
    let ordered =
        inlined
        |> List.ofSeq
        |> List.distinctBy (fun (start, finish, _) -> start, finish)
        |> List.sortBy (fun (start, _, _) -> start)
    if ordered.Length < inlined.Count then noteReason "two paragraphs of the block are of one element as it was read: the first is written in its place"
    let overlapping = ordered |> List.pairwise |> List.exists (fun ((_, firstEnd, _), (nextStart, _, _)) -> nextStart < firstEnd)
    if overlapping then noteReason "paragraphs of the block are of elements one inside the other: the block is written from the model"
    if ok && not overlapping then Some ordered else None

type private BodySegment =
    /// The part's text from `start` to `finish`, as it was read but for the paragraphs in it written from
    /// their source: in the place of each element (where it starts and ends), its paragraph.
    | SourceText of start: int * finish: int * paragraphs: (int * int * obj) list
    | ModelBlocks of obj array
    /// The part's text from `start` to `finish` (a block that is a table, and what stands before it) as it was read
    /// but for the table, which is its element as it was read with the model's changes made in it (patchedTableXml).
    | PatchedTable of start: int * finish: int * table: SourceTableText * item: obj
    /// The marks of ranges that stood between two blocks, before a block that is written from
    /// the model (or is not there any more): they stay where they stood.
    | CarriedMarks of RangeMark list

type private BodyPlan = {
    /// The part as it was read.
    Original: string
    /// Its root element (w:document).
    Root: XmlDsl.XmlNode
    /// The element that holds the blocks (w:body; the root of a header; a w:footnote, w:endnote, w:comment).
    Container: XmlDsl.XmlNode
    /// Where what w:body holds starts and ends.
    ContentStart: int
    ContentEnd: int
    Segments: BodySegment list
    /// What stands between the last block and the body's section properties (where it starts and ends).
    BeforeSection: int * int
    /// The body's own w:sectPr as it was read (where it is), when the page setup is what it was.
    Section: (int * int) option
    AfterSection: int * int
    /// The items written from the model.
    Regenerated: obj array
}

/// Where the blocks of a part are (see planBody).
type private BlockContainer =
    /// The document's: in w:body, with its section properties.
    | DocumentBody
    /// A header's or a footer's: in the part's root.
    | PartRoot
    /// A note's or a comment's: in its element (w:footnote, w:endnote, w:comment) of this w:id.
    | StoryElement of name: string * id: string

[<Fable.Core.Emit("(typeof $0.SourceIndex === 'number' ? $0.SourceIndex : -1)")>]
let private sourceIndexOf (_item: obj) : int = jsNative

/// Which of the body's elements go back as they were read. None when the
/// text cannot be matched with what the reader made of it (the body is then
/// written from the model, as a new document is), or when nothing is as it was.
/// `container`: where the blocks are - the document's in w:body; a header's or a footer's in its root, a
/// note's or a comment's in its element (neither has section properties of its own).
let private planBody (container: BlockContainer) (xml: string) (itemChunks: int array) (storyCounts: int array) (chunkCount: int) (items: obj array) (pageSetupUnchanged: bool) : BodyPlan option =
    let isBody = (container = DocumentBody)
    // The part's root, its w:body, and the elements the body holds (the blocks, as the reader counted them).
    let bodyOfRoot =
        XmlPatch.readPart xml
        |> Option.bind (fun part ->
            (match container with
             | DocumentBody -> XmlPatch.elements part.Node |> List.tryFind (fun node -> XmlPatch.localName node = "body")
             | PartRoot -> Some part.Node
             | StoryElement (name, id) -> XmlPatch.childrenNamed part name |> List.tryFind (fun node -> XmlPatch.attribute node "w:id" = Some id))
            |> Option.filter (fun body -> (XmlPatch.span body).StartTagEnd < (XmlPatch.span body).End)
            |> Option.map (fun body ->
                let kids = XmlPatch.elements body |> List.toArray
                part.Node, body, { Starts = kids |> Array.map (fun kid -> (XmlPatch.span kid).Start)
                                   Ends = kids |> Array.map (fun kid -> (XmlPatch.span kid).End)
                                   Nodes = kids
                                   CloseAt = (XmlPatch.span body).EndTagStart }))
    match bodyOfRoot with
    | Some (root, bodyNode, body) when body.Starts.Length = chunkCount && chunkCount > 0 ->
        let contentStart = (XmlPatch.span bodyNode).StartTagEnd
        let contentEnd = body.CloseAt
        let isSection index = isBody && XmlPatch.localName body.Nodes.[index] = "sectPr"
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
            // (The paragraphs that were read, for those that are written of them: see sourceParagraphXml.)
            sourceBodyXml <- xml
            sourceElements <- Some (indexOfElements root)
            sourceChunkStarts <- textStart
            sourceChunkElementStarts <- Array.copy body.Starts
            sourceChunkEnds <- Array.init chunkCount (fun chunk -> body.Ends.[chunk])
            sourceItemChunks <- itemChunks
            sourceStoryCounts <- storyCounts
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
            let patchedText = Array.create<(int * int * obj) list option> chunkCount None
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
            writingBody <- true
            for chunk in 0 .. chunkCount - 1 do
                if found.[chunk] <> originalCount.[chunk] then keep.[chunk] <- false
                if keep.[chunk] && patched.[chunk] then
                    patchedText.[chunk] <- patchSourceChunk textStart.[chunk] body.Ends.[chunk] (Array.sub items firstAt.[chunk] originalCount.[chunk])
                    if patchedText.[chunk].IsNone then keep.[chunk] <- false
            writingBody <- false
            // A field that begins in one element and ends in a later one (a table of
            // contents): those elements go back together or not at all.
            let blockChunks = [| for chunk in 0 .. chunkCount - 1 do if originalCount.[chunk] > 0 then yield chunk |]
            let mutable depth = 0
            let group = ResizeArray<int>()
            // An element in which every field that ends also begins holds none of the marks of
            // the field around it (an entry of a table of contents): it can be written from
            // the model while the others stay the text they were.
            // (The marks of the fields of an element's text - begin, separate, end - in order.)
            let fieldMarks =
                Array.init chunkCount (fun chunk ->
                    body.Nodes
                    |> Array.filter (fun node -> (XmlPatch.span node).Start >= textStart.[chunk] && (XmlPatch.span node).End <= body.Ends.[chunk])
                    |> Array.toList
                    |> List.collect (XmlPatch.descendants "fldChar")
                    |> List.choose (fun mark -> XmlPatch.attribute mark "w:fldCharType"))
            let neutral (chunk: int) =
                let depths = fieldMarks.[chunk] |> List.scan (fun depth mark -> match mark with "begin" -> depth + 1 | "end" -> depth - 1 | _ -> depth) 0
                List.forall (fun depth -> depth >= 0) depths && List.last depths = 0
            let inGroup = System.Collections.Generic.HashSet<int>()
            let closeGroup () =
                if group.Count > 1 then
                    for chunk in group do inGroup.Add chunk |> ignore
                    let together =
                        group |> Seq.forall (fun chunk -> keep.[chunk] || (firstAt.[chunk] >= 0 && neutral chunk))
                        && (Seq.pairwise group |> Seq.forall (fun (a, b) -> firstAt.[b] = lastAt.[a] + 1))
                    if not together then
                        let kept = group |> Seq.filter (fun chunk -> keep.[chunk]) |> Seq.length
                        if kept > 0 then noteReason (sprintf "%d blocks: a field begins in one block and ends in another (a table of contents), and one of its %d blocks is written from the model" kept group.Count)
                        for chunk in group do keep.[chunk] <- false
                group.Clear()
            for chunk in blockChunks do
                let begins = fieldMarks.[chunk] |> List.filter ((=) "begin") |> List.length
                let finishes = fieldMarks.[chunk] |> List.filter ((=) "end") |> List.length
                if depth > 0 || begins > finishes then
                    group.Add chunk
                    depth <- max 0 (depth + begins - finishes)
                    if depth = 0 then closeGroup ()
            closeGroup ()
            // A table whose rows or columns came or went (or that changed in more than its paragraphs), that is
            // its element's alone and not one of several held together by a field: its element as it was read with
            // the changes made in it (patchedTableXml), what stands before it as it was read.
            let patchedTables = Array.create<SourceTableText option> chunkCount None
            for chunk in 0 .. chunkCount - 1 do
                if originalCount.[chunk] = 1 && found.[chunk] = 1 && not keep.[chunk] && not (inGroup.Contains chunk) && firstAt.[chunk] >= 0 then
                    let item = items.[firstAt.[chunk]]
                    let next = firstAt.[chunk] + 1
                    if stringOrDefault "" (item?Type) = "Table" && not (isNullOrUndefined (item?SourceTable)) && not (isNullOrUndefined (item?SourceWhy))
                       && not (next < items.Length && isTextBox items.[next]) && body.Starts.[chunk] >= textStart.[chunk] then
                        match tableElementIn body.Nodes.[chunk] |> Option.bind tableSource.At with
                        | Some known when tableIsPatchable known item -> patchedTables.[chunk] <- Some known
                        | _ -> ()
            // In the order the items have now.
            let segments = ResizeArray<BodySegment>()
            let pending = ResizeArray<obj>()
            let regenerated = ResizeArray<obj>()
            let flush () =
                if pending.Count > 0 then
                    segments.Add (ModelBlocks (pending.ToArray()))
                    pending.Clear()
            // The marks of ranges that stand between two blocks (in the text before an element: see
            // textStart) stay where they stood when the element is written from the model, or is
            // not there any more: before the element's first item, or before what comes next.
            // (The elements of paragraphs joined to another: what stood before each goes with its text - see the splice.)
            let joinedChunks =
                let found = System.Collections.Generic.HashSet<int>()
                for item in items do
                    let changes = item?SourceTextChanges
                    if not (isNullOrUndefined changes) then
                        for change in optionalArray (changes?Edits) do
                            for join in optionalArray (change?Join) do
                                let block = intOrDefault -1 (join?Block)
                                if block >= 0 && block < itemChunks.Length && itemChunks.[block] >= 0 then found.Add itemChunks.[block] |> ignore
                found
            let carriedBefore = System.Collections.Generic.Dictionary<int, ResizeArray<RangeMark>>()
            let carriedAtEnd = ResizeArray<RangeMark>()
            for chunk in 0 .. chunkCount - 1 do
                if originalCount.[chunk] > 0 && not keep.[chunk] && patchedTables.[chunk].IsNone && not (joinedChunks.Contains chunk) && body.Starts.[chunk] > textStart.[chunk] then
                    let marks =
                        body.Nodes
                        |> Array.filter (fun node -> (XmlPatch.span node).Start >= textStart.[chunk] && (XmlPatch.span node).End <= body.Starts.[chunk])
                        |> Array.toList
                        |> List.collect (rangeMarksIn carriedRangeKinds)
                    if not marks.IsEmpty then
                            // (Where: the first item of the element, or of the next element that still has one.)
                            let mutable target = -1
                            let mutable next = chunk
                            while target < 0 && next < chunkCount do
                                if firstAt.[next] >= 0 then target <- firstAt.[next]
                                next <- next + 1
                            if target < 0 then carriedAtEnd.AddRange marks
                            else
                                if not (carriedBefore.ContainsKey target) then carriedBefore.[target] <- ResizeArray<RangeMark>()
                                carriedBefore.[target].AddRange marks
            let mutable position = 0
            let mutable anySource = false
            while position < items.Length do
                let chunk = chunkOfItem.[position]
                match carriedBefore.TryGetValue position with
                | true, marks ->
                    flush ()
                    segments.Add (CarriedMarks (List.ofSeq marks))
                | _ -> ()
                if chunk >= 0 && keep.[chunk] then
                    flush ()
                    match (if patched.[chunk] then patchedText.[chunk] else None) with
                    | Some paragraphs ->
                        segments.Add (SourceText (textStart.[chunk], body.Ends.[chunk], paragraphs))
                        for (_, _, para) in paragraphs do regenerated.Add para
                    | None -> segments.Add (SourceText (textStart.[chunk], body.Ends.[chunk], []))
                    anySource <- true
                    reportSourceBlocks <- reportSourceBlocks + originalCount.[chunk]
                    position <- position + originalCount.[chunk]
                elif chunk >= 0 && patchedTables.[chunk].IsSome then
                    flush ()
                    segments.Add (PatchedTable (textStart.[chunk], body.Ends.[chunk], patchedTables.[chunk].Value, items.[position]))
                    // (What of it is written from the model - a row, a cell - takes its pictures and links from it.)
                    regenerated.Add items.[position]
                    anySource <- true
                    reportSourceBlocks <- reportSourceBlocks + 1
                    position <- position + 1
                else
                    pending.Add items.[position]
                    regenerated.Add items.[position]
                    reportModelBlocks <- reportModelBlocks + 1
                    reportModelParagraphs <- reportModelParagraphs + storyParagraphCount items.[position]
                    position <- position + 1
            flush ()
            if carriedAtEnd.Count > 0 then segments.Add (CarriedMarks (List.ofSeq carriedAtEnd))
            // (Nothing as it was read: a new body. Around it the document stays the text it
            // was - the root's start tag, the section properties - when the page setup is.)
            if not anySource && (not pageSetupUnchanged || not isBody) then None
            else
                let sectionStart = if sectIndex >= 0 then body.Starts.[sectIndex] else contentEnd
                let sectionEnd = if sectIndex >= 0 then body.Ends.[sectIndex] else contentEnd
                Some { Original = xml
                       Root = root
                       Container = bodyNode
                       ContentStart = contentStart
                       ContentEnd = contentEnd
                       Segments = List.ofSeq segments
                       BeforeSection = (tailStart, max tailStart sectionStart)
                       Section = if pageSetupUnchanged || not isBody then Some (sectionStart, sectionEnd) else None
                       AfterSection = (sectionEnd, contentEnd)
                       Regenerated = regenerated.ToArray() }
    | _ -> None

/// The namespace declarations the root of a part needs for what was written from the model (`written`):
/// of the prefixes it uses (w, r, m) those the root does not declare. The root element of a document stays
/// as it was read unless something needs more.
let private writerNamespaceEdits (root: XmlDsl.XmlNode) (written: string list) : XmlPatch.ElementEdit list =
    let used =
        written
        |> List.choose XmlPatch.readFragmentRoot
        |> List.collect (fun fragment ->
            let nodes = XmlPatch.descendantsWhere (fun _ -> true) fragment.Node
            (nodes |> List.map XmlPatch.prefix)
            @ (nodes |> List.collect (fun node -> XmlPatch.attributeNames node) |> List.map ModelText.prefixOf))
        |> Set.ofList
    [ "w", Namespaces.word; "r", Namespaces.relationships; "m", Namespaces.math ]
    |> List.filter (fun (prefix, _) -> used.Contains prefix && (XmlPatch.attribute root ("xmlns:" + prefix)).IsNone)
    |> List.map (fun (prefix, uri) -> XmlPatch.SetAttribute ("xmlns:" + prefix, uri))

/// A piece of a part that is put together: a range of the text it was read as, or XML the model wrote.
type private PartPiece =
    | FromSource of start: int * finish: int
    | FromModel of xml: string

/// The ranges of a part that is put together of text as it was read and text
/// written from the model - bookmarks, comment ranges, permission ranges, the
/// ranges of moved text - opened once and closed once each. `parts`: the
/// pieces in order, and whether each was written from the model. Only what
/// was written from the model is changed: a start an earlier piece of the
/// model wrote (an end a later one writes) is not written again; an end of a
/// range nothing opens is left out; a range the model
/// opens and nothing closes - its end stood where the model does not keep
/// one: between two paragraphs, before the first run of another - is closed
/// at the end of its paragraph; one the file's text opens and nothing closes
/// any more is closed at the start of the next piece of the model - or where it
/// is opened, when none follows; an end in the file's text of a range nothing
/// opens any more (the block its start stood in is gone) is left out. What the
/// part had so when it was read (`source`: its text and root) is left as it is.
let private balancedRanges (source: (string * XmlDsl.XmlNode) option) (pieces: PartPiece array) : string array =
    let sourceText, sourceMarks =
        match source with
        | Some (text, root) -> text, rangeMarksIn allRangeKinds root
        | None -> "", []
    let original = source |> Option.map (fun _ -> sourceMarks)
    // (A piece of the source is a range of its text - not always well formed on its own: its marks are those
    // of the tree of the part in that range. A piece of the model is XML of its own.)
    let parts = pieces |> Array.map (function FromSource (start, finish) -> XmlPatch.slice sourceText start finish, false | FromModel text -> text, true)
    let trees = pieces |> Array.map (function FromModel text -> XmlPatch.readFragmentRoot text | FromSource _ -> None)
    let occurrences =
        pieces
        |> Array.mapi (fun index piece ->
            match piece, trees.[index] with
            | FromSource (start, finish), _ ->
                sourceMarks |> List.filter (fun mark -> (XmlPatch.span mark.Node).Start >= start && (XmlPatch.span mark.Node).End <= finish)
            | FromModel _, Some fragment -> rangeMarksIn allRangeKinds fragment.Node
            | FromModel _, None -> [])
    if occurrences |> Array.forall List.isEmpty then parts |> Array.map fst
    else
        let sourceStarts = System.Collections.Generic.Dictionary<string, int>()   // key -> the first piece that has it
        let sourceEnds = System.Collections.Generic.HashSet<string>()
        let modelStarts = System.Collections.Generic.Dictionary<string, int>()    // key -> how many
        let modelEnds = System.Collections.Generic.Dictionary<string, int>()
        let count (table: System.Collections.Generic.Dictionary<string, int>) key = match table.TryGetValue key with | true, n -> n | false, _ -> 0
        let kindOf = System.Collections.Generic.Dictionary<string, string * string>()
        let mutable modelHasAny = false
        for index in 0 .. parts.Length - 1 do
            let isModel = snd parts.[index]
            for mark in occurrences.[index] do
                let key = mark.Key
                kindOf.[key] <- (mark.Kind, mark.Id)
                if isModel then
                    modelHasAny <- true
                    if mark.Opens then modelStarts.[key] <- count modelStarts key + 1 else modelEnds.[key] <- count modelEnds key + 1
                elif mark.Opens then (if not (sourceStarts.ContainsKey key) then sourceStarts.[key] <- index)
                else sourceEnds.Add key |> ignore
        // A range the file's text opens and nothing closes any more: closed at the start of the next piece of the model.
        let closeAt = System.Collections.Generic.Dictionary<int, ResizeArray<string>>()
        // What the file's own text has of a range whose other mark is nowhere any more (the block
        // it stood in was deleted, or the model does not hold it): an end nothing opens, and a
        // start nothing closes that no piece of the model follows. Not what the file had so itself.
        let strayEnds = System.Collections.Generic.HashSet<string>()
        let strayStarts = System.Collections.Generic.HashSet<string>()
        for KeyValue (key, at) in sourceStarts do
            if not (sourceEnds.Contains key) && count modelEnds key = 0 then
                let mutable next = -1
                for index in at + 1 .. parts.Length - 1 do
                    if next < 0 && snd parts.[index] then next <- index
                if next >= 0 then
                    if not (closeAt.ContainsKey next) then closeAt.[next] <- ResizeArray<string>()
                    closeAt.[next].Add key
                else strayStarts.Add key |> ignore
        for key in sourceEnds do
            if not (sourceStarts.ContainsKey key) && count modelStarts key = 0 then strayEnds.Add key |> ignore
        if strayEnds.Count > 0 || strayStarts.Count > 0 || closeAt.Count > 0 then
            // (As the file had it: left as it is.)
            match original with
            | Some marks ->
                let starts = marks |> List.filter (fun mark -> mark.Opens) |> List.map (fun mark -> mark.Key) |> Set.ofList
                let ends = marks |> List.filter (fun mark -> not mark.Opens) |> List.map (fun mark -> mark.Key) |> Set.ofList
                for key in Seq.toArray strayEnds do if not (starts.Contains key) then strayEnds.Remove key |> ignore
                for key in Seq.toArray strayStarts do if not (ends.Contains key) then strayStarts.Remove key |> ignore
                for KeyValue (_, keys) in closeAt do keys.RemoveAll (fun key -> not (ends.Contains key)) |> ignore
            | None ->
                strayEnds.Clear()
                strayStarts.Clear()
        let anyClose = closeAt.Values |> Seq.exists (fun keys -> keys.Count > 0)
        if not modelHasAny && not anyClose && strayEnds.Count = 0 && strayStarts.Count = 0 then parts |> Array.map fst
        else
            let endTag (key: string) = let (kind, id) = kindOf.[key] in rangeEndXml kind id
            let seenStarts = System.Collections.Generic.HashSet<string>()
            let seenEnds = System.Collections.Generic.Dictionary<string, int>()
            // (A part with edits made in it; one that is not well formed stays as it is.)
            let edited (index: int) (edits: (XmlDsl.XmlNode * XmlPatch.ElementEdit list) list) =
                match trees.[index] with
                | Some fragment when not edits.IsEmpty -> XmlPatch.applyFragment XmlPatch.Word.order fragment (XmlPatch.editsAt fragment.Node edits)
                | _ -> fst parts.[index]
            parts
            |> Array.mapi (fun index (text, isModel) ->
                if not isModel then
                    match pieces.[index] with
                    | FromSource (start, finish) when strayEnds.Count > 0 || strayStarts.Count > 0 ->
                        // (A start nothing closes: closed where it is opened. An end nothing opens: left out.)
                        XmlPatch.sliceWith sourceText start finish
                            [ for mark in occurrences.[index] do
                                  let span = XmlPatch.span mark.Node
                                  if mark.Opens && strayStarts.Contains mark.Key then yield span.Start, span.End, XmlPatch.bytesOf sourceText mark.Node + endTag mark.Key
                                  elif not mark.Opens && strayEnds.Contains mark.Key then yield span.Start, span.End, "" ]
                    | _ -> text
                else
                    let opening = match closeAt.TryGetValue index with | true, keys -> keys |> Seq.map endTag |> String.concat "" | false, _ -> ""
                    // (The paragraphs of the part by where they end: the first that ends after a mark is the one it is in, or the next.)
                    let paragraphs =
                        match trees.[index] with
                        | Some fragment -> XmlPatch.descendantsWhere (fun node -> XmlPatch.name node = "w:p") fragment.Node |> List.sortBy (fun para -> (XmlPatch.span para).EndTagStart)
                        | None -> []
                    let atEnd = System.Text.StringBuilder()
                    let edits =
                        [ for mark in occurrences.[index] do
                              let key = mark.Key
                              let keep =
                                  if mark.Opens then
                                      // (The first the model writes. What the file's own text has is not looked at:
                                      // a file can use an id twice, and its text is left as it is.)
                                      seenStarts.Add key
                                  else
                                      let seen = count seenEnds key + 1
                                      seenEnds.[key] <- seen
                                      // (The last one the model writes, of a range that is opened.)
                                      seen = count modelEnds key && (sourceStarts.ContainsKey key || count modelStarts key > 0)
                              if not keep then yield mark.Node, [ XmlPatch.Removed ]
                              elif mark.Opens && not (sourceEnds.Contains key) && count modelEnds key = 0 then
                                  // (A range the model opens and nothing closes: closed at the end of the paragraph it is
                                  // in - or of the first after it -, or at the end of the part.)
                                  match paragraphs |> List.tryFind (fun para -> (XmlPatch.span para).EndTagStart >= (XmlPatch.span mark.Node).End) with
                                  | Some para -> yield para, [ XmlPatch.Append (endTag key) ]
                                  | None -> atEnd.Append(endTag key) |> ignore ]
                    opening + edited index edits + atEnd.ToString())

/// A part that is written from the model whole - the body of a new document, a header or a
/// footer, the notes, the comments - with its ranges opened once and closed once each.
let private balancedPart (xml: string) : string =
    // (The part's root holds what is balanced.)
    match XmlPatch.readPart xml with
    | Some part ->
        let span = XmlPatch.span part.Node
        if span.StartTagEnd >= span.End then xml
        else
            let content = XmlPatch.slice xml span.StartTagEnd span.EndTagStart
            let balanced = (balancedRanges None [| FromModel content |]).[0]
            if balanced = content then xml else XmlPatch.slice xml 0 span.StartTagEnd + balanced + XmlPatch.slice xml span.EndTagStart xml.Length
    | None -> xml

/// What a plan's container holds up to its section properties, written: its blocks (the ranges in them opened
/// and closed once each) and what stands after the last of them. What is written from the model is added to
/// `written`.
let private planContentParts (plan: BodyPlan) (setup: PageSetupInfo) (images: ImageScope) (hyperlinks: Map<string, string>) (written: ResizeArray<string>) : ResizeArray<string> =
    let slice (start: int, finish: int) = XmlPatch.slice plan.Original start finish
    let pieces = ResizeArray<PartPiece>()
    // (Paragraphs written from the model splice the part's text: see splicedRuns. The origins of what is
    // written are of this part: see sourceElementAt.)
    spliceSourceText <- plan.Original
    writingBody <- true
    for segment in plan.Segments do
        match segment with
        | SourceText (start, finish, paragraphs) ->
            // The paragraphs written from their source: each in the place of its element.
            let mutable at = start
            for (paraStart, paraEnd, para) in paragraphs |> List.sortBy (fun (paraStart, _, _) -> paraStart) do
                pieces.Add (FromSource (at, paraStart))
                let splicedBefore = reportSpliced
                let paragraph = generateParagraphWithHyperlinksAndImagesWithPageSetup setup hyperlinks images para
                // (The paragraphs of a block kept as it was read that are not written from their source.)
                if reportSpliced = splicedBefore then reportInline <- reportInline + 1
                written.Add paragraph
                pieces.Add (FromModel paragraph)
                at <- paraEnd
            pieces.Add (FromSource (at, finish))
        | PatchedTable (start, finish, table, item) ->
            let span = XmlPatch.span table.TableNode
            pieces.Add (FromSource (start, span.Start))
            let text = patchedTableXml hyperlinks images table item
            written.Add text
            pieces.Add (FromModel text)
            pieces.Add (FromSource (span.End, finish))
        | ModelBlocks items ->
            let text = items |> generateBlocks (generateDocumentItemWithImagesAndPageSetup setup hyperlinks images) (generateTextBoxRun setup hyperlinks images)
            written.Add text
            pieces.Add (FromModel text)
        | CarriedMarks marks ->
            // (As what the model writes: a mark of a range that is not there any more is left out.)
            // (Nor the range of a comment the document has no more.)
            let live =
                marks
                |> List.filter (fun mark -> mark.Kind <> "commentRange" || Set.contains mark.Id liveCommentIds)
                |> List.map (fun mark -> XmlPatch.bytesOf plan.Original mark.Node)
                |> String.concat ""
            pieces.Add (FromModel live)
    writingBody <- false
    let parts = ResizeArray<string>(balancedRanges (Some (plan.Original, plan.Root)) (pieces.ToArray()))
    parts.Add (slice plan.BeforeSection)
    parts

let private generateDocumentXmlFromPlan (plan: BodyPlan) (images: ImageScope) (hyperlinks: Map<string, string>) (pageSetup: PageSetupInfo option) (headers: HeaderFooterPartInfo array) (footers: HeaderFooterPartInfo array) : string =
    let setup = pageSetup |> Option.defaultValue defaultPageSetupInfo
    let slice (start: int, finish: int) = XmlPatch.slice plan.Original start finish
    let written = ResizeArray<string>()
    let parts = planContentParts plan setup images hyperlinks written
    match plan.Section with
    | Some range -> parts.Add (slice range)
    | None ->
        let globalHeaders = headers |> Array.filter (fun part -> part.IsGlobal)
        let globalFooters = footers |> Array.filter (fun part -> part.IsGlobal)
        let text = generateSectionProperties pageSetup globalHeaders globalFooters
        written.Add text
        parts.Add text
    parts.Add (slice plan.AfterSection)
    // The part: its root's start tag with the namespaces what was written needs, the body's start tag as it
    // was read, what it holds now, and the rest as it was read.
    let rootSpan = XmlPatch.span plan.Root
    let namespaces = if written.Count > 0 then writerNamespaceEdits plan.Root (List.ofSeq written) else []
    XmlPatch.slice plan.Original 0 rootSpan.Start
    + XmlPatch.startTag plan.Original plan.Root namespaces
    + XmlPatch.slice plan.Original rootSpan.StartTagEnd plan.ContentStart
    + String.concat "" parts
    + XmlPatch.slice plan.Original plan.ContentEnd plan.Original.Length

/// The element of a note or a comment written from its plan (see planBody): its start and end tags as they were
/// read, what it holds as `planContentParts` writes it. The run with the note's or comment's own mark
/// (w:footnoteRef, w:endnoteRef, w:annotationRef) - which the model does not hold as a run - is what it was:
/// where what was written has none, the one that was read is put at the start of the first paragraph.
let private generateStoryXmlFromPlan (plan: BodyPlan) (images: ImageScope) (hyperlinks: Map<string, string>) : string =
    let written = ResizeArray<string>()
    let parts = planContentParts plan defaultPageSetupInfo images hyperlinks written
    let span = XmlPatch.span plan.Container
    let text =
        XmlPatch.slice plan.Original span.Start plan.ContentStart
        + String.concat "" parts
        + XmlPatch.slice plan.Original plan.ContentEnd span.End
    let isMark (node: XmlDsl.XmlNode) = List.contains (XmlPatch.localName node) [ "footnoteRef"; "endnoteRef"; "annotationRef" ]
    let markRunOf (node: XmlDsl.XmlNode) =
        XmlPatch.descendantsWhere (fun run -> XmlPatch.localName run = "r" && (XmlPatch.elements run |> List.exists isMark)) node |> List.tryHead
    match markRunOf plan.Container, XmlPatch.readFragmentRoot text with
    | Some mark, Some fragment ->
        match XmlPatch.elements fragment.Node with
        | [ story ] when (markRunOf story).IsNone ->
            match XmlPatch.descendantsWhere (fun node -> XmlPatch.localName node = "p") story |> List.tryHead with
            | Some first -> XmlPatch.applyFragment XmlPatch.Word.order fragment (XmlPatch.editsAt fragment.Node [ first, [ XmlPatch.Prepend (XmlPatch.bytesOf plan.Original mark) ] ])
            | None -> text
        | _ -> text
    | _ -> text


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
    /// The note is what it was when it was read.
    Unchanged: bool
    /// Of one that changed: how many paragraphs each of its blocks had when it was read (its blocks say which
    /// they were then, as the body's do - see planBody).
    SourceStoryCounts: int array
}

type CommentPart = {
    Id: string
    Author: string option
    Date: string option
    Initials: string option
    Blocks: obj array
    PlainText: string option
    Unchanged: bool
    SourceStoryCounts: int array
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
              PlainText = None
              Unchanged = boolOrDefault false (item?SourceUnchanged)
              SourceStoryCounts = optionalArray (item?SourceStoryCounts) |> Array.map (fun count -> intOrDefault 0 count) }))

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
              PlainText = None
              Unchanged = boolOrDefault false (item?SourceUnchanged)
              SourceStoryCounts = optionalArray (item?SourceStoryCounts) |> Array.map (fun count -> intOrDefault 0 count) }))

/// Bodies from the package first; a note that is only known from a reference
/// run (NoteText) is added as plain text.
let private mergeNoteParts (fromPackage: NotePart list) (fromRuns: NoteRef list) : NotePart list =
    let known = fromPackage |> List.map (fun note -> note.Id) |> Set.ofList
    fromPackage
    @ (fromRuns
       |> List.filter (fun note -> not (Set.contains note.Id known))
       |> List.map (fun note -> { Id = note.Id; Kind = None; HasMark = false; MarkStyle = None; Blocks = [||]; PlainText = Some note.Text; Unchanged = false; SourceStoryCounts = [||] }))

let private mergeCommentParts (fromPackage: CommentPart list) (fromRuns: CommentRef list) : CommentPart list =
    let known = fromPackage |> List.map (fun comment -> comment.Id) |> Set.ofList
    fromPackage
    @ (fromRuns
       |> List.filter (fun comment -> not (Set.contains comment.Id known))
       |> List.map (fun comment -> { Id = comment.Id; Author = comment.Author; Date = comment.Date; Initials = None; Blocks = [||]; PlainText = Some comment.Text; Unchanged = false; SourceStoryCounts = [||] }))

let private partRootAttributes =
    [("xmlns:w", Namespaces.word); ("xmlns:r", Namespaces.relationships); ("xmlns:m", Namespaces.math)]

/// footnotes.xml / endnotes.xml with the note bodies as blocks. The reader
/// used to hand over the text of a note only, and this wrote one plain run per
/// note, so a table, a hyperlink or bold text in a footnote was gone after a save.
let private noteElement (noteElementName: string) (markElementName: string) (generateBody: obj array -> string) (note: NotePart) : string =
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
        let isSeparator = note.Kind = Some "separator" || note.Kind = Some "continuationSeparator"
        if not note.HasMark && not isSeparator then body
        else
            match XmlPatch.readFragmentRoot body with
            | Some fragment ->
                match XmlPatch.elements fragment.Node with
                | first :: _ when XmlPatch.name first = "w:p" ->
                    XmlPatch.applyFragment XmlPatch.Word.order fragment (XmlPatch.editsAt fragment.Node [ first, [ XmlPatch.Prepend (markRun note) ] ])
                | _ -> body
            | None -> body
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
    element noteElementName attrs [body]

let private generateNotesPart (rootName: string) (noteElementName: string) (markElementName: string) (generateBody: obj array -> string) (notes: NotePart list) : string =
    document (element rootName partRootAttributes (notes |> List.map (noteElement noteElementName markElementName generateBody)))

let private commentElement (generateBody: obj array -> string) (comment: CommentPart) : string =
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
    element "w:comment" attrs [body]

let private generateCommentsPart (generateBody: obj array -> string) (comments: CommentPart list) : string =
    document (element "w:comments" partRootAttributes (comments |> List.map (commentElement generateBody)))

/// A part of notes or comments (`elementName`: w:footnote, w:endnote, w:comment) as it was read with what
/// changed made in it: the element of each that is what it was (`unchanged`, by id) stays its text, one that
/// changed is what the model writes of it (`write`; `ids`: those the model has, in order), one the model has
/// no more goes, a new one is put in at the end; the root keeps its declarations and gets those what the model
/// wrote needs. None when the part does not read.
let private patchedStoriesPart (xml: string) (elementName: string) (ids: string list) (unchanged: Set<string>) (write: string -> string) : string option =
    XmlPatch.readPart xml
    |> Option.map (fun part ->
        let readIdsFirst = XmlPatch.childrenNamed part elementName |> List.choose (fun node -> XmlPatch.attribute node "w:id") |> Set.ofList
        let written = ids |> List.filter (fun id -> not (Set.contains id unchanged && Set.contains id readIdsFirst)) |> List.map (fun id -> id, write id)
        let byId = Map.ofList written
        let present = Set.ofList ids
        let read = XmlPatch.childrenNamed part elementName |> List.choose (fun node -> XmlPatch.attribute node "w:id" |> Option.map (fun id -> id, node))
        let readIds = read |> List.map fst |> Set.ofList
        let changes =
            [ for id, node in read do
                  let at = (XmlPatch.span node).Start
                  match Map.tryFind id byId with
                  | Some text -> yield XmlPatch.EditAt (at, [ XmlPatch.Replaced text ])
                  | None when not (Set.contains id present) -> yield XmlPatch.EditAt (at, [ XmlPatch.Removed ])
                  | None -> ()
              for id, text in written do
                  if not (Set.contains id readIds) then yield XmlPatch.Append text ]
        let fromModel = written |> List.map snd
        let namespaces = if fromModel.IsEmpty then [] else writerNamespaceEdits part.Node fromModel
        if changes.IsEmpty && namespaces.IsEmpty then xml
        else XmlPatch.applyPart (fun _ -> []) part (namespaces @ changes))

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
    | XmlDsl.Element (rootName, rootAttrs, children, _) ->
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
    match XmlPatch.readPart original with
    | Some part when (match part.Node with XmlDsl.Element (name, _, _, _) -> name = "w:numbering" | _ -> false) ->
        let lists = XmlPatch.childrenNamed part "w:num"
        let known = lists |> List.choose (fun num -> XmlPatch.attribute num "w:numId") |> Set.ofList
        // (The highest number any list or definition has, or refers to.)
        let highest =
            [ for num in lists do yield! Option.toList (XmlPatch.attribute num "w:numId")
              for definition in XmlPatch.childrenNamed part "w:abstractNum" do yield! Option.toList (XmlPatch.attribute definition "w:abstractNumId")
              for reference in XmlPatch.descendants "abstractNumId" part.Node do yield! Option.toList (XmlPatch.attribute reference "w:val")
              for reference in XmlPatch.descendants "numId" part.Node do yield! Option.toList (XmlPatch.attribute reference "w:val") ]
            |> List.choose (fun text -> match System.Int32.TryParse text with | true, value -> Some value | _ -> None)
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
            // (CT_Numbering: the definitions, then the lists - each put in after those of its kind.)
            Some (XmlPatch.applyPart XmlPatch.Word.order part
                    ([ for xml in abstracts -> XmlPatch.AddChild ("w:abstractNum", xml) ] @ [ for xml in nums -> XmlPatch.AddChild ("w:num", xml) ]))
    | _ -> if not needsDefaults && customInfos.IsEmpty then Some original else None

/// The numbering part as it was read with the levels the editor changed
/// ("list id/level") written anew: in the list's own abstract definition (a
/// copy of it, made for the list, when other lists share the definition), the
/// other levels, lists and definitions staying the text they were. None when
/// a list or its definition is not found in the text.
let private patchNumberingLevels (original: string) (customInfos: NumberingInfo list) (changed: string array) : string option =
    let mutable text = Some original
    for key in changed do
        match text |> Option.bind XmlPatch.readPart with
        | None -> text <- None
        | Some part ->
            // (The editor counts levels from one, the file from zero.)
            let numId, level = ModelText.numberingLevel key
            let lists = XmlPatch.childrenNamed part "w:num"
            let list = lists |> List.tryFind (fun num -> XmlPatch.attribute num "w:numId" = Some numId)
            let abstractId = list |> Option.bind (XmlPatch.child "abstractNumId") |> Option.bind (fun reference -> XmlPatch.attribute reference "w:val")
            let definition = abstractId |> Option.bind (fun id -> XmlPatch.childrenNamed part "w:abstractNum" |> List.tryFind (fun node -> XmlPatch.attribute node "w:abstractNumId" = Some id))
            match customInfos |> List.tryFind (fun info -> info.NumId = numId && info.Level = level && not info.Undefined), abstractId, definition with
            | Some info, Some abstractId, Some definition ->
                let shared = (lists |> List.filter (fun num -> (XmlPatch.child "abstractNumId" num |> Option.bind (fun reference -> XmlPatch.attribute reference "w:val")) = Some abstractId)).Length > 1
                let highest =
                    [ for num in lists do yield! Option.toList (XmlPatch.attribute num "w:numId")
                      for node in XmlPatch.childrenNamed part "w:abstractNum" do yield! Option.toList (XmlPatch.attribute node "w:abstractNumId") ]
                    |> List.choose (fun value -> match System.Int32.TryParse value with | true, number -> Some number | _ -> None)
                    |> List.fold max 0
                let ownId = if shared then string (highest + 1) else abstractId
                // The level written anew, in the place of the one that is there.
                let levelXml = numberingLevelWithIndents false info.IndentLeft info.FirstLineIndent info.Level info.Start info.Format info.LevelText
                let levelEdit = XmlPatch.SetChildWhere ("w:lvl", "w:ilvl", string level, levelXml)
                // The list: its definition, and no override of the level but the editor's own start.
                let listEdits =
                    [ if shared then yield XmlPatch.EditChild ("w:abstractNumId", [ XmlPatch.SetAttribute ("w:val", ownId) ], None)
                      yield XmlPatch.RemoveChildWhere ("w:lvlOverride", "w:ilvl", string level)
                      if info.HasOverride then
                          yield XmlPatch.AddChild ("w:lvlOverride", element "w:lvlOverride" [("w:ilvl", string level)] [ selfClosing "w:startOverride" [("w:val", string (info.Start |> Option.defaultValue 1))] ]) ]
                let definitionEdits =
                    if shared then
                        // A definition of its own for this list: with an id and a signature of its own.
                        let own = XmlPatch.apply XmlPatch.Word.order { XmlPatch.SourceElement.Text = part.Text; XmlPatch.SourceElement.Node = definition } [ XmlPatch.SetAttribute ("w:abstractNumId", ownId); XmlPatch.RemoveChild "w:nsid"; XmlPatch.RemoveChild "w:tmpl"; levelEdit ]
                        [ XmlPatch.AddChild ("w:abstractNum", own) ]
                    else [ XmlPatch.EditChildWhere ("w:abstractNum", "w:abstractNumId", abstractId, [ levelEdit ]) ]
                text <- Some (XmlPatch.applyPart XmlPatch.Word.order part (definitionEdits @ [ XmlPatch.EditChildWhere ("w:num", "w:numId", numId, listEdits) ]))
            | _ -> text <- None
    text

let private generateSettings (pageSetup: PageSetupInfo option) =
    let children = [
        match pageSetup with
        | Some setup when setup.EvenAndOddHeaders -> yield selfClosing "w:evenAndOddHeaders" []
        | _ -> ()
    ]
    document (element "w:settings" [("xmlns:w", Namespaces.word)] children)

/// Children of w:settings that come before w:evenAndOddHeaders in CT_Settings; every other comes after it.
let private settingsOrder : XmlPatch.SchemaOrder = fun _ ->
    [ "writeProtection"; "view"; "zoom"; "removePersonalInformation"; "removeDateAndTime"; "doNotDisplayPageBoundaries"
      "displayBackgroundShape"; "printPostScriptOverText"; "printFractionalCharacterWidth"; "printFormsData"; "embedTrueTypeFonts"
      "embedSystemFonts"; "saveSubsetFonts"; "saveFormsData"; "mirrorMargins"; "alignBordersAndEdges"; "bordersDoNotSurroundHeader"
      "bordersDoNotSurroundFooter"; "gutterAtTop"; "hideSpellingErrors"; "hideGrammaticalErrors"; "activeWritingStyle"; "proofState"
      "formsDesign"; "attachedTemplate"; "linkStyles"; "stylePaneFormatFilter"; "stylePaneSortMethod"; "documentType"; "mailMerge"
      "revisionView"; "trackRevisions"; "doNotTrackMoves"; "doNotTrackFormatting"; "documentProtection"; "autoFormatOverride"
      "styleLockTheme"; "styleLockQFSet"; "defaultTabStop"; "autoHyphenation"; "consecutiveHyphenLimit"; "hyphenationZone"
      "doNotHyphenateCaps"; "showEnvelope"; "summaryLength"; "clickAndTypeStyle"; "defaultTableStyle"; "evenAndOddHeaders"; "*" ]

/// The original settings part with the one setting the editor owns, even and
/// odd headers, brought up to date. Everything else (compatibility options,
/// default tab stop, track-revisions flag, footnote properties, rsids, ...)
/// stays: the part used to be dropped, or replaced by one holding this flag only.
let private patchSettings (originalXml: string) (pageSetup: PageSetupInfo option) : string =
    match XmlPatch.readPart originalXml with
    | Some part ->
        let evenAndOdd = pageSetup |> Option.exists (fun setup -> setup.EvenAndOddHeaders)
        let setting = XmlPatch.elements part.Node |> List.tryFind (fun node -> XmlPatch.localName node = "evenAndOddHeaders")
        // The setting is what the part says already: the part as it was read.
        let hadEvenAndOdd =
            setting |> Option.exists (fun node -> match XmlPatch.attribute node "w:val" with Some "0" | Some "false" | Some "off" -> false | _ -> true)
        if hadEvenAndOdd = evenAndOdd then originalXml
        else
            let name = match setting with Some node -> XmlPatch.name node | None -> XmlDsl.Ns.withPrefixOf (XmlPatch.name part.Node) "evenAndOddHeaders"
            XmlPatch.applyPart settingsOrder part
                [ if evenAndOdd then XmlPatch.SetChild (name, selfClosing name []) else XmlPatch.RemoveChild name ]
    | None -> generateSettings pageSetup

// ---- relationship parts and [Content_Types].xml as they were read -------------
// The writer makes these parts anew for every save: its own order, its own
// ids for the parts nothing refers to by id (the styles, the lists, the
// settings, the notes, the comments), absolute targets. A part that was read
// and says what the new one says is written back as the text it was.

/// Relationship types of the parts a document has one of, which nothing refers to by id.
let private isImplicitRelationshipType (relType: string) =
    [ "styles"; "numbering"; "settings"; "comments"; "footnotes"; "endnotes" ] |> List.contains (OpcXml.relationshipKind relType)

/// What a relationships part says: one entry per relationship (its id - but
/// for the parts nothing refers to by id -, type, mode and target, a target
/// in the package as its path), sorted. None when it is not well formed.
let private relationshipEntries (ownerDirectory: string) (xml: string) : string list option =
    OpcXml.relationships xml
    |> Option.map (fun rels ->
        rels
        |> List.map (fun rel ->
            (if isImplicitRelationshipType rel.Type then "" else rel.Id) + "\n" + rel.Type + "\n" + (if rel.External then "external" else "internal") + "\n" + OpcXml.targetPath ownerDirectory rel)
        |> List.sort)

/// Whether the [Content_Types].xml that was read is right for the package
/// that is written: every part has the type the writer gives it, and it names
/// no part that is not written.
let private contentTypesStillRight (original: string) (generated: string) (writtenPaths: string list) =
    let originalType, originalOverrides = OpcXml.contentTypeLookup original
    let generatedType, _ = OpcXml.contentTypeLookup generated
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
    reportGenerated <- 0
    reportDerived <- 0
    reportSpliced <- 0
    saveNumber <- saveNumber + 1
    liveCommentIds <- Set.empty
    forgetSourceBody ()
    shapesFromModel <- regenerateAll
    let sourceOf (key: string) = if regenerateAll then None else Map.tryFind key preserved.SourceXml
    let paragraphs = optionalArray (package?Paragraphs)
    // The body's elements that go back as they were read (see planBody).
    let bodyPlan =
        match sourceOf "document" with
        | Some xml when preserved.BodyItemChunks.Length > 0 ->
            planBody DocumentBody xml preserved.BodyItemChunks (optionalArray (package?SourceStoryCounts) |> Array.map (fun count -> intOrDefault 0 count)) preserved.BodyChunkCount paragraphs (boolOrDefault false (package?PageSetupUnchanged))
        | _ -> None
    let bodyState = sourceState ()
    // A header or footer that was changed: what of it is as it was read goes back as it was, as the body's
    // blocks do (its plan, and what is known of its text while it is written).
    let partPlans =
        Array.append (optionalArray (package?HeaderParts)) (optionalArray (package?FooterParts))
        |> Array.choose (fun part ->
            match optionalString (part?RelationshipId) with
            | Some relId when not (boolOrDefault false (part?SourceUnchanged)) ->
                match sourceOf relId, Map.tryFind relId preserved.PartItemChunks with
                | Some xml, Some (chunks, count) when chunks.Length > 0 ->
                    forgetSourceBody ()
                    let counts = optionalArray (part?SourceStoryCounts) |> Array.map (fun count -> intOrDefault 0 count)
                    planBody PartRoot xml chunks counts count (partBlocks part) false
                    |> Option.map (fun plan -> relId, (plan, sourceState ()))
                | _ -> None
            | _ -> None)
        |> Map.ofArray
    forgetSourceBody ()
    useSourceState bodyState
    /// The parts that are written back as they were read keep all the
    /// relationships they had; a part written from the model keeps those the
    /// writer does not make itself.
    let sourceParts = System.Collections.Generic.HashSet<string>()
    if bodyPlan.IsSome then sourceParts.Add "document" |> ignore
    for KeyValue (relId, _) in partPlans do sourceParts.Add relId |> ignore
    for part in Array.append (optionalArray (package?HeaderParts)) (optionalArray (package?FooterParts)) do
        match optionalString (part?RelationshipId) with
        | Some relId when boolOrDefault false (part?SourceUnchanged) && (sourceOf relId).IsSome -> sourceParts.Add relId |> ignore
        | _ -> ()
    // (A part of notes or comments that was read keeps them whether it is as it was or not: the notes in it
    // that are what they were are written back as they were - see patchedStoriesPart.)
    for key in [ "footnotes"; "endnotes"; "comments" ] do
        if (sourceOf key).IsSome then sourceParts.Add key |> ignore
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
        match partSource part, Map.tryFind part.RelId partPlans with
        | Some _, _ -> { Info = part; Images = emptyImageScope; Hyperlinks = Map.empty }
        | None, Some (plan, _) ->
            // (Its pictures and links are those of what is written from the model; the others keep theirs.)
            let images, next = buildImageScope nextMediaIndex plan.Regenerated
            nextMediaIndex <- next
            { Info = part; Images = images; Hyperlinks = hyperlinkMapWith (keptFor part.RelId) plan.Regenerated }
        | None, None ->
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
    liveCommentIds <- comments |> List.map (fun comment -> comment.Id) |> Set.ofList
    deletedCommentIds <-
        match sourceOf "comments" |> Option.bind XmlPatch.readPart with
        | Some part -> XmlPatch.childrenNamed part "w:comment" |> List.choose (fun comment -> XmlPatch.attribute comment "w:id") |> Set.ofList |> fun read -> Set.difference read liveCommentIds
        | None -> Set.empty
    let footnotes = mergeNoteParts (notePartsFromPackage (package?Footnotes)) (collectNotes "footnote" contentBlocks)
    let endnotes = mergeNoteParts (notePartsFromPackage (package?Endnotes)) (collectNotes "endnote" contentBlocks)
    // A note or comment that was changed: what of it is as it was read goes back as it was, as the body's blocks
    // do (its plan, and what is known of its part while it is written), by its part and id ("footnotes:2").
    let notePlans =
        let planOf (key: string) (elementName: string) (id: string) (unchanged: bool) (blocks: obj array) (counts: int array) =
            if unchanged || blocks.Length = 0 then None
            else
                match sourceOf key, Map.tryFind (key + ":" + id) preserved.NoteItemChunks with
                | Some xml, Some (chunks, count) when chunks.Length > 0 ->
                    forgetSourceBody ()
                    planBody (StoryElement (elementName, id)) xml chunks counts count blocks false
                    |> Option.map (fun plan -> key + ":" + id, (plan, sourceState ()))
                | _ -> None
        [ for note in footnotes do yield planOf "footnotes" "w:footnote" note.Id note.Unchanged note.Blocks note.SourceStoryCounts
          for note in endnotes do yield planOf "endnotes" "w:endnote" note.Id note.Unchanged note.Blocks note.SourceStoryCounts
          for comment in comments do yield planOf "comments" "w:comment" comment.Id comment.Unchanged comment.Blocks comment.SourceStoryCounts ]
        |> List.choose id
        |> Map.ofList
    forgetSourceBody ()
    useSourceState bodyState
    // (Of a note written from its plan, the blocks written from the model.)
    let storyBlocks (key: string) (id: string) (blocks: obj array) =
        match Map.tryFind (key + ":" + id) notePlans with
        | Some (plan, _) -> plan.Regenerated
        | None -> blocks
    // Each of these parts has its own relationships: its images and hyperlinks.
    let bodyPart (key: string) (blocks: obj array) =
        let images, next = buildImageScope nextMediaIndex blocks
        nextMediaIndex <- next
        images, hyperlinkMapWith (keptFor key) blocks
    // (Of a part that was read, those of the notes that changed: the others are written back as they were.)
    let readFor (key: string) = (sourceOf key).IsSome
    let footnoteBlocks = footnotes |> List.filter (fun note -> not (readFor "footnotes" && note.Unchanged)) |> List.toArray |> Array.collect (fun note -> storyBlocks "footnotes" note.Id note.Blocks)
    let endnoteBlocks = endnotes |> List.filter (fun note -> not (readFor "endnotes" && note.Unchanged)) |> List.toArray |> Array.collect (fun note -> storyBlocks "endnotes" note.Id note.Blocks)
    let commentBlocks = comments |> List.filter (fun comment -> not (readFor "comments" && comment.Unchanged)) |> List.toArray |> Array.collect (fun comment -> storyBlocks "comments" comment.Id comment.Blocks)
    // The notes and the comments as they were read, when nothing of them was changed.
    let footnotesSource = if boolOrDefault false (package?FootnotesUnchanged) then sourceOf "footnotes" else None
    let endnotesSource = if boolOrDefault false (package?EndnotesUnchanged) then sourceOf "endnotes" else None
    let commentsSource = if boolOrDefault false (package?CommentsUnchanged) then sourceOf "comments" else None
    let footnoteImages, footnoteLinks = if footnotesSource.IsSome then emptyImageScope, Map.empty else bodyPart "footnotes" footnoteBlocks
    let endnoteImages, endnoteLinks = if endnotesSource.IsSome then emptyImageScope, Map.empty else bodyPart "endnotes" endnoteBlocks
    let commentImages, commentLinks = if commentsSource.IsSome then emptyImageScope, Map.empty else bodyPart "comments" commentBlocks
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
           // (All of them: the lists and styles of those written back as they were are the document's too.)
           for note in footnotes do yield! note.Blocks
           for note in endnotes do yield! note.Blocks
           for comment in comments do yield! comment.Blocks |]
    // Rendered first: it picks the default list ids the paragraph writer uses.
    let numberingXml =
        let infos = collectNumberingInfoOf (not (isNullOrUndefined (package?Preserved))) allBlocks
        // Blocks written from the model with a list that has no id of the file.
        let modelBlocks =
            [| yield! bodyModelBlocks
               for part in Array.append headerParts footerParts do
                   match partSource part, Map.tryFind part.RelId partPlans with
                   | Some _, _ -> ()
                   | None, Some (plan, _) -> yield! plan.Regenerated
                   | None, None -> yield! part.Blocks
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
    // A header or footer: from its plan (what is as it was read stays its text), or from the model.
    let writeHeaderFooter (rootName: string) (part: HeaderFooterPartPackage) =
        match Map.tryFind part.Info.RelId partPlans with
        | Some (plan, state) ->
            useSourceState state
            let text = withRelationshipsOf part.Info.RelId plan.Regenerated (fun () -> generateDocumentXmlFromPlan plan part.Images part.Hyperlinks None [||] [||])
            forgetSourceBody ()
            text
        | None -> withRelationshipsOf part.Info.RelId part.Info.Blocks (fun () -> balancedPart (generateHeaderFooterXml rootName part.Hyperlinks part.Images part.Info.Blocks))
    // (The links of a part that kept the relationships it had: those it had are among the kept ones.)
    let newLinks (key: string) (links: Map<string, string>) =
        let keptIds = keptFor key |> List.map (fun rel -> rel.Id) |> Set.ofList
        links |> Map.filter (fun _ id -> not (Set.contains id keptIds))
    // A part of notes or comments that was read, with what changed made in it (see patchedStoriesPart): each
    // that changed from its plan (what is as it was read stays its text), or the element the model writes with
    // its ranges opened and closed in it.
    let stories (key: string) (elementName: string) (images: ImageScope) (links: Map<string, string>) (items: (string * bool) list) (write: string -> string) =
        sourceOf key
        |> Option.bind (fun xml ->
            let unchanged = items |> List.filter snd |> List.map fst |> Set.ofList
            patchedStoriesPart xml elementName (List.map fst items) unchanged (fun id ->
                match Map.tryFind (key + ":" + id) notePlans with
                | Some (plan, state) ->
                    useSourceState state
                    let text = generateStoryXmlFromPlan plan images links
                    forgetSourceBody ()
                    text
                | None -> (balancedRanges None [| FromModel (write id) |]).[0]))
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
                        | None -> balancedPart (generateDocumentXmlPackage paragraphs bodyImages pageSetup headerParts footerParts)) }
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
                        | None -> writeHeaderFooter "w:hdr" header }
            if partRelsNeeded header.Info.RelId header.Images header.Hyperlinks then
                yield { Path = sprintf "word/_rels/%s.rels" header.Info.Target; Content = generatePartRelsWithKept (keptFor header.Info.RelId) (newLinks header.Info.RelId header.Hyperlinks) header.Images }
        for footer in footerPackages do
            yield { Path = "word/" + footer.Info.Target
                    Content =
                        match partSource footer.Info with
                        | Some text -> text
                        | None -> writeHeaderFooter "w:ftr" footer }
            if partRelsNeeded footer.Info.RelId footer.Images footer.Hyperlinks then
                yield { Path = sprintf "word/_rels/%s.rels" footer.Info.Target; Content = generatePartRelsWithKept (keptFor footer.Info.RelId) (newLinks footer.Info.RelId footer.Hyperlinks) footer.Images }
        if hasComments then
            yield { Path = "word/comments.xml"
                    Content =
                        match commentsSource with
                        | Some text -> text
                        | None ->
                            withRelationshipsOf "comments" commentBlocks (fun () ->
                                let write = commentElement (partBody commentImages commentLinks)
                                stories "comments" "w:comment" commentImages commentLinks (comments |> List.map (fun comment -> comment.Id, comment.Unchanged)) (fun id -> write (comments |> List.find (fun comment -> comment.Id = id)))
                                |> Option.defaultWith (fun () -> balancedPart (generateCommentsPart (partBody commentImages commentLinks) comments))) }
            if partRelsNeeded "comments" commentImages commentLinks then
                yield { Path = "word/_rels/comments.xml.rels"; Content = generatePartRelsWithKept (keptFor "comments") (newLinks "comments" commentLinks) commentImages }
        if hasFootnotes then
            yield { Path = "word/footnotes.xml"
                    Content =
                        match footnotesSource with
                        | Some text -> text
                        | None ->
                            withRelationshipsOf "footnotes" footnoteBlocks (fun () ->
                                let write = noteElement "w:footnote" "w:footnoteRef" (partBody footnoteImages footnoteLinks)
                                stories "footnotes" "w:footnote" footnoteImages footnoteLinks (footnotes |> List.map (fun note -> note.Id, note.Unchanged)) (fun id -> write (footnotes |> List.find (fun note -> note.Id = id)))
                                |> Option.defaultWith (fun () -> balancedPart (generateNotesPart "w:footnotes" "w:footnote" "w:footnoteRef" (partBody footnoteImages footnoteLinks) footnotes))) }
            if partRelsNeeded "footnotes" footnoteImages footnoteLinks then
                yield { Path = "word/_rels/footnotes.xml.rels"; Content = generatePartRelsWithKept (keptFor "footnotes") (newLinks "footnotes" footnoteLinks) footnoteImages }
        if hasEndnotes then
            yield { Path = "word/endnotes.xml"
                    Content =
                        match endnotesSource with
                        | Some text -> text
                        | None ->
                            withRelationshipsOf "endnotes" endnoteBlocks (fun () ->
                                let write = noteElement "w:endnote" "w:endnoteRef" (partBody endnoteImages endnoteLinks)
                                stories "endnotes" "w:endnote" endnoteImages endnoteLinks (endnotes |> List.map (fun note -> note.Id, note.Unchanged)) (fun id -> write (endnotes |> List.find (fun note -> note.Id = id)))
                                |> Option.defaultWith (fun () -> balancedPart (generateNotesPart "w:endnotes" "w:endnote" "w:endnoteRef" (partBody endnoteImages endnoteLinks) endnotes))) }
            if partRelsNeeded "endnotes" endnoteImages endnoteLinks then
                yield { Path = "word/_rels/endnotes.xml.rels"; Content = generatePartRelsWithKept (keptFor "endnotes") (newLinks "endnotes" endnoteLinks) endnoteImages }
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
            match PartPath.ofRelationships file.Path with
            | Some partPath when PartPath.directory file.Path = "word/_rels" ->
                match Map.tryFind partPath contentOf with
                | None -> file
                | Some part ->
                    let usedNow = OpcXml.relationshipIdsOfPart part
                    let source = Map.tryFind partPath sourceOfPath |> Option.map (fun (source, ids) -> OpcXml.relationshipIdsOfPart source, ids)
                    let originalIds =
                        Map.tryFind file.Path preserved.OriginalText |> Option.map (OpcXml.relationshipsOrEmpty >> List.map (fun rel -> rel.Id) >> Set.ofList)
                    let unused (rel: OpcXml.Relationship) =
                        // (A link's relationship the part had and did not use when it was read: the file's own.)
                        let inherited =
                            match source with
                            | Some (usedThen, ids) -> Set.contains rel.Id ids && not (Set.contains rel.Id usedThen)
                            | None -> false
                        // (And a picture's: one the file's relationships named and the part it was read as did not
                        // use - a picture a tool left behind - is the file's own too. Only what a change made
                        // unused goes.)
                        let leftBehind =
                            match source, originalIds with
                            | Some (usedThen, _), Some ids -> not (Set.contains rel.Id usedThen) && Set.contains rel.Id ids
                            | _ -> false
                        let kind = OpcXml.relationshipKind rel.Type
                        let picture = kind = "image" && not rel.External && not leftBehind
                        let link = kind = "hyperlink" && not inherited
                        (picture || link) && not (Set.contains rel.Id usedNow)
                    { file with Content = OpcXml.withoutRelationships unused file.Content }
            | _ -> file)
    let binaryFiles =
        // (Every part a relationship of the package names, as its path.)
        let named =
            (files |> List.map (fun file -> file.Path, file.Content))
            @ (binaryFiles |> List.map (fun (path, data) -> path, System.Text.Encoding.UTF8.GetString data))
            |> List.filter (fun (path, _) -> PartPath.ofRelationships path |> Option.isSome)
            |> List.collect (fun (path, text) ->
                OpcXml.relationshipsOrEmpty text |> List.filter (fun rel -> not rel.External) |> List.map (OpcXml.targetPath (OpcXml.ownerDirectory path)))
            |> Set.ofList
        binaryFiles
        |> List.filter (fun (path, _) -> PartPath.directory path <> "word/media" || Set.contains path named)

    // A document that was read without a styles part gets none while nothing needs one: no
    // style was made or changed, and no part names one. Nor the relationship and the content
    // type of one; and the body's relationships part, which the file did not have, is not
    // written when nothing is left in it. (A shape moved in a document of one part was saved
    // with a styles part of the writer's defaults and a relationships part to name it.)
    let files =
        let stylesChanged = not (isNullOrUndefined package?ChangedStyles) && (optionalArray (package?ChangedStyles)).Length > 0
        let wasRead = Map.containsKey "[Content_Types].xml" preserved.OriginalText
        let namesStyle (file: PackageFile) =
            PartPath.directory file.Path = "word" && PartPath.extension file.Path = "xml" && file.Path <> "word/styles.xml"
            && (match XmlPatch.readPart file.Content with
                | Some part -> [ "pStyle"; "rStyle"; "tblStyle" ] |> List.exists (fun name -> not (XmlPatch.descendants name part.Node).IsEmpty)
                | None -> false)
        if regenerateAll || not wasRead || preserved.StylesXml.IsSome || stylesChanged || not (keptFor "styles").IsEmpty || List.exists namesStyle files then files
        else
            files
            |> List.filter (fun file -> file.Path <> "word/styles.xml")
            |> List.map (fun file ->
                if file.Path = "word/_rels/document.xml.rels" then
                    { file with Content = file.Content |> OpcXml.withoutRelationships (fun rel -> OpcXml.relationshipKind rel.Type = "styles") }
                elif file.Path = "[Content_Types].xml" then
                    let styles = OpcXml.contentTypeEntries file.Content |> Option.defaultValue [] |> List.filter (function OpcXml.Override (name, _) -> name = "/word/styles.xml" | _ -> false)
                    { file with Content = file.Content |> OpcXml.editContentTypes styles [] }
                else file)
            |> List.filter (fun file ->
                not (file.Path = "word/_rels/document.xml.rels"
                     && not (Map.containsKey file.Path preserved.OriginalText)
                     && OpcXml.relationshipsOrEmpty file.Content |> List.isEmpty))

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
                    elif PartPath.ofRelationships file.Path |> Option.isSome then
                        let directory = OpcXml.ownerDirectory file.Path
                        match relationshipEntries directory original, relationshipEntries directory file.Content with
                        | Some before, Some after when before = after -> { file with Content = original }
                        | _ -> file
                    else file)

    reservedRelationshipIds <- Set.empty
    reservedPartPaths <- Set.empty
    forgetSourceBody ()
    createPackageWithBinary files binaryFiles
