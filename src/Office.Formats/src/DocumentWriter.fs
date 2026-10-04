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
let private sectionPropertiesChildren (references: string list) (sectionType: string option) (setup: PageSetupInfo option) (pageNumberStart: int option) =
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
      match pageNumberStart with
      | Some start -> yield selfClosing "w:pgNumType" [("w:start", string (max 0 start))]
      | None -> ()
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
}

let private noRunExtras =
    { StyleId = None; AllCaps = false; SmallCaps = false; DoubleStrike = false; Hidden = false
      UnderlineStyle = None; UnderlineColor = None; Position = 0.0; EmphasisMark = None
      HasExplicit = false; ExplicitBold = None; ExplicitItalic = None; ExplicitUnderline = None
      ExplicitAllCaps = None; ExplicitSmallCaps = None; ExplicitStrike = None; ExplicitDoubleStrike = None
      ExplicitHidden = None; ExplicitFontSize = None }

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
        | None -> ()
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
}

let private isDeletedRun (runObj: obj) =
    optionalString (runObj?RevisionType) = Some "delete"

let private generateRunWithObjectProps (runObj: obj) (run: EditorRun) : string =
    runElement (orderedRunProps (runExtrasFromObject runObj) run) (isDeletedRun runObj) run.Text

let private generateFieldRun (runObj: obj) (fallbackRun: EditorRun) (fieldCode: string) =
    let result = optionalString (runObj?FieldResult) |> Option.defaultValue fallbackRun.Text
    let resultRun = generateRunWithObjectProps runObj { fallbackRun with Text = result }
    element "w:fldSimple" [("w:instr", fieldCode)] [resultRun]

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
                    element "w:r" [] [selfClosing "w:commentReference" [("w:id", commentId)]]
                | None, Some noteId ->
                    let noteType = optionalString (runObj?NoteType) |> Option.defaultValue "footnote"
                    let elementName = if noteType = "endnote" then "w:endnoteReference" else "w:footnoteReference"
                    element "w:r" [] [selfClosing elementName [("w:id", noteId)]]
                | None, None ->
                    if fallbackRun.Text = "" && bookmarkStartXml runObj |> Option.isSome || fallbackRun.Text = "" && bookmarkEndXml runObj |> Option.isSome then ""
                    else generateRunWithObjectProps runObj fallbackRun
    [ bookmarkStartXml runObj; (if inner <> "" then Some inner else None); bookmarkEndXml runObj ]
    |> List.choose id
    |> String.concat ""

let private revisionWrappedRun (runObj: obj) (runXml: string) =
    match optionalString (runObj?RevisionType) with
    | Some revisionType when revisionType = "insert" || revisionType = "delete" ->
        let elementName = if revisionType = "delete" then "w:del" else "w:ins"
        let attrs = [
            yield ("w:id", optionalString (runObj?RevisionId) |> Option.defaultValue "1")
            match optionalString (runObj?RevisionAuthor) with
            | Some author -> yield ("w:author", author)
            | None -> ()
            match optionalString (runObj?RevisionDate) with
            | Some date -> yield ("w:date", date)
            | None -> ()
        ]
        element elementName attrs [runXml]
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
}

let private emptyPreservedPackage = {
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
        { Parts =
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

let private collectParagraphHyperlinks (para: obj) =
    optionalArray (para?Runs)
    |> Array.toList
    |> List.choose (fun run -> optionalString (run?Hyperlink))

let private collectItemHyperlinks (item: obj) =
    collectItemParagraphs item |> List.collect collectParagraphHyperlinks

let private hyperlinkMap (items: obj array) =
    items
    |> Array.toList
    |> List.collect collectItemHyperlinks
    |> List.distinct
    // Not "rId<n>": header/footer parts keep the relationship ids they were read
    // with (rId7, ...), and a hyperlink numbered into the same range pointed at
    // header1.xml after saving.
    |> List.mapi (fun index url -> (url, freshRelationshipId (sprintf "rIdLink%d" (index + 1))))
    |> Map.ofList

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
                if src.StartsWith("http://") || src.StartsWith("https://") then
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
/// Tight and through need a wrap polygon; the model has none, so the bounding box is given.
let private anchorWrapElement (wrap: string option) (wrapText: string option) =
    let side = wrapText |> Option.defaultValue "bothSides"
    let polygon =
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
            let wrap = anchorWrapElement (optionalString (runObj?ImageWrap)) (optionalString (runObj?ImageWrapText))
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
            |> fun setup -> { setup with HeaderReferences = []; FooterReferences = []; PageNumberStart = None }
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
        Some (element "w:sectPr" [] (sectionPropertiesChildren refs sectionBreak (Some setup) pageNumberStart))

let private defaultFontSizePixels = 11.0 * 96.0 / 72.0

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

let private paragraphMarkRunProps (para: obj) =
    // Exact / at-least line spacing is stored relative to the paragraph font
    // size, so pin that size on the paragraph mark: the reader then divides by
    // the same number next time and the value is a fixed point.
    match optionalString (para?LineSpacingRule), optionalFloat (para?LineSpacing) with
    | Some rule, Some _ when rule = "exact" || rule = "atLeast" ->
        let sz = InverseConversion.pixelsToHalfPoints (paragraphFontSize para)
        Some (element "w:rPr" [] [selfClosing "w:sz" [("w:val", sz)]; selfClosing "w:szCs" [("w:val", sz)]])
    | _ -> None

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
        match optionalFloat (para?SpacingBefore) with
        | Some value ->
            yield ("w:before", InverseConversion.pointsToTwips value)
            yield ("w:beforeAutospacing", "0")
        | None -> ()
        match optionalFloat (para?SpacingAfter) with
        | Some value ->
            yield ("w:after", InverseConversion.pointsToTwips value)
            yield ("w:afterAutospacing", "0")
        | None -> ()
        match optionalFloat (para?LineSpacing) with
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
        match optionalFloat (para?IndentLeft) with
        | Some value -> yield ("w:left", InverseConversion.pixelsToTwips value)
        | None -> ()
        match optionalFloat (para?IndentRight) with
        | Some value -> yield ("w:right", InverseConversion.pixelsToTwips value)
        | None -> ()
        match optionalFloat (para?FirstLineIndent) with
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
        | Some styleId when styleId <> "" -> yield selfClosing "w:pStyle" [("w:val", styleId)]
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
        if boolOrDefault false (para?PageBreakBefore) then yield selfClosing "w:pageBreakBefore" []
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
        match optionalString (para?BackgroundColor) with
        | Some color ->
            let wordColor = InverseConversion.cssColorToWord (Some color) |> Option.defaultValue color
            yield selfClosing "w:shd" [("w:val", "clear"); ("w:color", "auto"); ("w:fill", wordColor)]
        | None -> ()
        match tabStopsElement para with
        | Some tabs -> yield tabs
        | None -> ()
        if boolOrDefault false (para?Rtl) then yield selfClosing "w:bidi" []
        if not spacingAttributes.IsEmpty then yield selfClosing "w:spacing" spacingAttributes
        if not indentAttributes.IsEmpty then yield selfClosing "w:ind" indentAttributes
        yield! onOff "w:contextualSpacing" (para?StyleContextualSpacing)
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
let private generateParagraphWithHyperlinksAndImagesWithPageSetup (packageSetup: PageSetupInfo) (hyperlinks: Map<string, string>) (images: ImageScope) (para: obj) : string =
    let pPr = generateParaPropsWithPageSetup packageSetup para
    let runsArray = para?Runs |> unbox<obj array>
    let runs =
        runsArray
        |> Array.toList
        |> List.map (fun runObj ->
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
                    let run: EditorRun = {
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
                    generateReferenceRun runObj run
            let runXml = revisionWrappedRun runObj runXml
            match optionalString (runObj?Hyperlink) |> Option.bind (fun url -> Map.tryFind url hyperlinks) with
            | Some relId -> element "w:hyperlink" [("r:id", relId); ("w:history", "1")] [runXml]
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

    element "w:p" [] (pPr :: runs)

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

    // CT_TblPr order: tblW, jc, tblCellSpacing, tblBorders, shd, tblLayout, tblCellMar.
    let tblPrChildren = [
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
    ]

    element "w:tbl" [] [
        element "w:tblPr" [] tblPrChildren
        generateTableGrid table
        rows
    ]

/// Shape ids must be unique in the document; they are handed out per save.
let mutable private nextShapeId = 1000

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
            element "v:textbox" [] [
                element "w:txbxContent" [] [content]
            ]
        ]
    ]

/// A text box / shape as Word writes it: DrawingML (wp:anchor or wp:inline
/// around a wps:wsp) in an mc:AlternateContent with a VML fallback. Geometry,
/// rotation, solid / gradient / no fill, line colour, width and dash, wrap
/// mode and position are written from the model. (It used to be a bare VML
/// text box rectangle, so an ellipse or a rotated arrow came back as a white box.)
let private generateTextBoxRun (packageSetup: PageSetupInfo) (hyperlinks: Map<string, string>) (images: ImageScope) (item: obj) : string =
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
            yield selfClosing "wps:bodyPr" [("rot", "0"); ("vert", "horz"); ("wrap", "square"); ("lIns", "91440"); ("tIns", "45720"); ("rIns", "91440"); ("bIns", "45720"); ("anchor", "t"); ("anchorCtr", "0")]
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
                anchorWrapElement (optionalString (item?Wrap) |> Option.orElse (Some "none")) None
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

type HeaderFooterPartInfo = { Kind: string; RelId: string; Target: string; PartName: string; Blocks: obj array; IsGlobal: bool }
type HeaderFooterPartPackage = { Info: HeaderFooterPartInfo; Images: ImageScope; Hyperlinks: Map<string, string> }

let private partBlocks (part: obj) =
    if not (isNullOrUndefined part?Blocks) then optionalArray (part?Blocks)
    elif not (isNullOrUndefined part?Paragraphs) then optionalArray (part?Paragraphs)
    else [||]

let private headerFooterInfos scopedRelationshipIds prefix startIndex (parts: obj array) =
    parts
    |> Array.mapi (fun index part ->
        let number = startIndex + index
        let explicitRelId = optionalString (part?RelationshipId)
        let relId = explicitRelId |> Option.defaultValue (freshRelationshipId (sprintf "rId%s%d" prefix number))
        let fileName = sprintf "%s%d.xml" prefix number
        { Kind = part?Kind |> stringOrDefault (part?Type |> stringOrDefault "default")
          RelId = relId
          Target = fileName
          PartName = sprintf "/word/%s" fileName
          Blocks = partBlocks part
          IsGlobal = not (Set.contains relId scopedRelationshipIds) })

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
let private generateContentTypesPreserved (preserved: PreservedPackage) hasSettings hasComments hasFootnotes hasEndnotes (headers: HeaderFooterPartInfo array) (footers: HeaderFooterPartInfo array) (imageAssets: ImageAsset array) =
    let own =
        [ yield "/word/document.xml", preserved.MainContentType |> Option.defaultValue "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
          yield "/word/styles.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
          yield "/word/numbering.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
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
let private documentRelationshipElements (hyperlinks: Map<string, string>) (hasComments: bool) (hasFootnotes: bool) (hasEndnotes: bool) =
    let baseRels = [
        yield selfClosing "Relationship" [("Id", freshRelationshipId "rIdStyles"); ("Type", RelationshipTypes.styles); ("Target", "styles.xml")]
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
let private generateDocumentRelsPreserved (kept: PreservedRelationshipInfo list) hasSettings hyperlinks hasComments hasFootnotes hasEndnotes (headers: HeaderFooterPartInfo array) (footers: HeaderFooterPartInfo array) (images: ImageScope) =
    let headerFooterRels =
        [ if hasSettings then yield selfClosing "Relationship" [("Id", freshRelationshipId "rIdSettings"); ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"); ("Target", "settings.xml")]
          for header in headers -> selfClosing "Relationship" [("Id", header.RelId); ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"); ("Target", header.Target)]
          for footer in footers -> selfClosing "Relationship" [("Id", footer.RelId); ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"); ("Target", footer.Target)] ]
    let rels =
        documentRelationshipElements hyperlinks hasComments hasFootnotes hasEndnotes
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
                    withMark note (if body = "" then element "w:p" [] [selfClosing "w:pPr" []] else body)
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

type NumberingInfo = { NumId: string; Level: int; Format: string; LevelText: string; Start: int option; HasOverride: bool }

/// One entry per (list id, level) used in the document. The list id, number
/// format, level text and start value all come from the paragraph: without
/// them every numbered paragraph was written with numId 2, which merged
/// separate lists into one decimal list and lost letters / roman numerals.
let private collectNumberingInfo (items: obj array) =
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
                { NumId = numId; Level = level; Format = format; LevelText = levelText; Start = start; HasOverride = restart }))
    |> List.distinctBy (fun info -> info.NumId, info.Level)

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
        if styleId = "" then None
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

let private defaultStyleId = "__docDefaults"

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
let private numberingLevel (withIndent: bool) (level: int) (start: int option) (format: string) (levelText: string) =
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
        match (if format = "bullet" then bulletFont levelText else None) with
        | Some font -> yield element "w:rPr" [] [selfClosing "w:rFonts" [("w:ascii", font); ("w:hAnsi", font); ("w:hint", "default")]]
        | None -> ()
    ]

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

    let groups = customInfos |> List.groupBy (fun info -> info.NumId)
    let abstracts =
        [ yield element "w:abstractNum" [("w:abstractNumId", bulletId)] (defaultLevels "bullet" "\u2022")
          yield element "w:abstractNum" [("w:abstractNumId", numberedId)] (defaultLevels "decimal" "")
          for numId, infos in groups do
              let levels =
                  infos
                  |> List.sortBy (fun info -> info.Level)
                  |> List.map (fun info -> numberingLevel false info.Level info.Start info.Format info.LevelText)
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

let writeFromEditorPackage (package: obj) : Fable.Core.JS.Promise<obj> =
    nextShapeId <- 1000
    // What was kept of the opened package: written back around the parts the
    // editor owns. A package without it (a new document) is written from scratch.
    let preserved = preservedPackageFromObject (package?Preserved)
    let keptFor (key: string) = preserved.PartRelationships |> Map.tryFind key |> Option.defaultValue []
    reservedRelationshipIds <-
        preserved.PartRelationships
        |> Map.toList
        |> List.collect (fun (_, rels) -> rels |> List.map (fun rel -> rel.Id))
        |> Set.ofList
    reservedPartPaths <- preserved.Parts |> List.map fst |> Set.ofList
    let withRelationshipsOf (key: string) (generate: unit -> string) =
        currentPartRelationshipIds <- keptFor key |> List.map (fun rel -> rel.Id) |> Set.ofList
        let xml = generate ()
        currentPartRelationshipIds <- Set.empty
        xml
    let paragraphs = optionalArray (package?Paragraphs)
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
    let headerParts = optionalArray (package?HeaderParts) |> headerFooterInfos sectionReferenceIds "header" 1
    let footerParts = optionalArray (package?FooterParts) |> headerFooterInfos sectionReferenceIds "footer" 1
    let bodyImages, nextAfterBody = buildImageScope 1 paragraphs
    let mutable nextMediaIndex = nextAfterBody
    let headerPackages =
        headerParts
        |> Array.map (fun part ->
            let images, next = buildImageScope nextMediaIndex part.Blocks
            nextMediaIndex <- next
            { Info = part; Images = images; Hyperlinks = hyperlinkMap part.Blocks })
    let footerPackages =
        footerParts
        |> Array.map (fun part ->
            let images, next = buildImageScope nextMediaIndex part.Blocks
            nextMediaIndex <- next
            { Info = part; Images = images; Hyperlinks = hyperlinkMap part.Blocks })
    let contentBlocks =
        [| yield! paragraphs
           for part in headerParts do yield! part.Blocks
           for part in footerParts do yield! part.Blocks |]
    let hyperlinks = hyperlinkMap paragraphs
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
    let footnoteImages, footnoteLinks = bodyPart footnoteBlocks
    let endnoteImages, endnoteLinks = bodyPart endnoteBlocks
    let commentImages, commentLinks = bodyPart commentBlocks
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
    let numberingXml = generateNumbering (collectNumberingInfo allBlocks)
    let styles =
        (collectPackageStyles (optionalArray (package?Styles)) @ collectStyles allBlocks)
        |> List.distinctBy (fun style -> style.StyleType, style.StyleId)
    let hasComments = not comments.IsEmpty
    let hasFootnotes = not footnotes.IsEmpty
    let hasEndnotes = not endnotes.IsEmpty
    let partRelsNeeded (key: string) (images: ImageScope) (links: Map<string, string>) =
        images.Assets.Length > 0 || not links.IsEmpty || not (keptFor key).IsEmpty

    let files = [
        yield { Path = "[Content_Types].xml"; Content = generateContentTypesPreserved preserved hasSettings hasComments hasFootnotes hasEndnotes headerParts footerParts allImages }
        yield { Path = "_rels/.rels"; Content = generateRootRelsPreserved preserved }
        yield { Path = "word/document.xml"; Content = withRelationshipsOf "document" (fun () -> generateDocumentXmlPackage paragraphs bodyImages pageSetup headerParts footerParts) }
        yield { Path = "word/_rels/document.xml.rels"; Content = generateDocumentRelsPreserved (keptFor "document") hasSettings hyperlinks hasComments hasFootnotes hasEndnotes headerParts footerParts bodyImages }
        yield { Path = "word/styles.xml"
                Content =
                    match preserved.StylesXml with
                    | Some original ->
                        let changed =
                            if isNullOrUndefined package?ChangedStyles then None
                            else Some (optionalArray (package?ChangedStyles) |> Array.map (fun key -> unbox<string> key) |> Set.ofArray)
                        mergeStyles original changed styles
                    | None -> generateDefaultStylesWithStyles styles }
        yield { Path = "word/numbering.xml"; Content = numberingXml }
        if hasSettings then
            yield { Path = "word/settings.xml"
                    Content =
                        match preserved.SettingsXml with
                        | Some original -> patchSettings original pageSetup
                        | None -> generateSettings pageSetup }
            if not (keptFor "settings").IsEmpty then
                yield { Path = "word/_rels/settings.xml.rels"; Content = generatePartRelsWithKept (keptFor "settings") Map.empty emptyImageScope }
        for header in headerPackages do
            yield { Path = "word/" + header.Info.Target; Content = withRelationshipsOf header.Info.RelId (fun () -> generateHeaderFooterXml "w:hdr" header.Hyperlinks header.Images header.Info.Blocks) }
            if partRelsNeeded header.Info.RelId header.Images header.Hyperlinks then
                yield { Path = sprintf "word/_rels/%s.rels" header.Info.Target; Content = generatePartRelsWithKept (keptFor header.Info.RelId) header.Hyperlinks header.Images }
        for footer in footerPackages do
            yield { Path = "word/" + footer.Info.Target; Content = withRelationshipsOf footer.Info.RelId (fun () -> generateHeaderFooterXml "w:ftr" footer.Hyperlinks footer.Images footer.Info.Blocks) }
            if partRelsNeeded footer.Info.RelId footer.Images footer.Hyperlinks then
                yield { Path = sprintf "word/_rels/%s.rels" footer.Info.Target; Content = generatePartRelsWithKept (keptFor footer.Info.RelId) footer.Hyperlinks footer.Images }
        if hasComments then
            yield { Path = "word/comments.xml"; Content = withRelationshipsOf "comments" (fun () -> generateCommentsPart (partBody commentImages commentLinks) comments) }
            if partRelsNeeded "comments" commentImages commentLinks then
                yield { Path = "word/_rels/comments.xml.rels"; Content = generatePartRelsWithKept (keptFor "comments") commentLinks commentImages }
        if hasFootnotes then
            yield { Path = "word/footnotes.xml"; Content = withRelationshipsOf "footnotes" (fun () -> generateNotesPart "w:footnotes" "w:footnote" "w:footnoteRef" (partBody footnoteImages footnoteLinks) footnotes) }
            if partRelsNeeded "footnotes" footnoteImages footnoteLinks then
                yield { Path = "word/_rels/footnotes.xml.rels"; Content = generatePartRelsWithKept (keptFor "footnotes") footnoteLinks footnoteImages }
        if hasEndnotes then
            yield { Path = "word/endnotes.xml"; Content = withRelationshipsOf "endnotes" (fun () -> generateNotesPart "w:endnotes" "w:endnote" "w:endnoteRef" (partBody endnoteImages endnoteLinks) endnotes) }
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

    reservedRelationshipIds <- Set.empty
    reservedPartPaths <- Set.empty
    createPackageWithBinary files binaryFiles
