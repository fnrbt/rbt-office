/// Word's object model for macros (VBA run by Rbt.Vb), writable: Application, ActiveDocument, Selection,
/// Range (text, character and paragraph formatting, styles, Find and Replace), Paragraphs / Words /
/// Characters / Sentences, Tables (Add, Cell, Rows, Columns, Shading, Borders), Bookmarks and Styles.
///
/// Positions are those of Word's main story: each paragraph's text and then its mark (one character,
/// vbCr), a table cell's last paragraph ending in the end-of-cell mark and each row in an end-of-row
/// mark (one character each, read as vbCr & Chr(7)). A run works on a copy of the document
/// (`MacroDocument`); its host commits the result as one step of the editing session's history.
/// Shared by the browser and the native hosts.
module Rbt.Word.WordMacroHost

open System
open System.Collections.Generic
open Rbt.Vb.Values
open Rbt.Word.Document
open Rbt.Word.WordMacroApi
module TB = Rbt.Word.TextBuffer

// ---- errors ------------------------------------------------------------------------------------------

let private fail number (description: string) : 'a = raiseVbErrorWith number description
let private unsupported (typeName: string) (memberName: string) : 'a =
    fail 438 (sprintf "Object doesn't support this property or method (%s.%s)" typeName memberName)
let private noSuchMember () : 'a = fail 5941 "The requested member of the collection does not exist."
let private notAvailable (what: string) : 'a = fail 4605 (sprintf "This method or property is not available (%s)" what)

let private eq (left: string) (right: string) = String.Equals(left, right, StringComparison.OrdinalIgnoreCase)

// ---- values --------------------------------------------------------------------------------------------

let rec private scalarOf (value: Value) =
    match value with
    | VRef (values, index) when index >= 0 && index < values.Length -> scalarOf values.[index]
    | VFieldRef (instance, field) ->
        match instance.Fields.TryGetValue field with
        | true, inner -> scalarOf inner
        | _ -> VEmpty
    | VHostObject host ->
        match host.GetMember "_Default" with
        | Some inner -> scalarOf inner
        | None -> value
    | other -> other

let private toFloat (value: Value) =
    match scalarOf value with
    | VInteger i -> float i
    | VDouble d -> d
    | VDate serial -> serial
    | VCurrency units -> float units / 10000.0
    | VBoolean b -> if b then -1.0 else 0.0
    | VString s ->
        match Double.TryParse(s, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture) with
        | true, d -> d
        | _ -> fail 13 "Type mismatch"
    | VEmpty | VNull -> 0.0
    | _ -> fail 13 "Type mismatch"

let private toInt (value: Value) = int (Math.Round(toFloat value))

let private toText (value: Value) =
    match scalarOf value with
    | VString s -> s
    | VInteger i -> string i
    | VDouble d -> Rbt.Vb.Api.displayValue (VDouble d)
    | VDate _ | VCurrency _ as v -> Rbt.Vb.Api.displayValue v
    | VBoolean b -> if b then "True" else "False"
    | VEmpty | VNull | VNothing | VUndefined -> ""
    | _ -> fail 13 "Type mismatch"

let private toBool (value: Value) =
    match scalarOf value with
    | VBoolean b -> b
    | VInteger i -> i <> 0
    | VDouble d -> d <> 0.0
    | VString s when eq s "true" -> true
    | VString s when eq s "false" -> false
    | VEmpty -> false
    | _ -> fail 13 "Type mismatch"

let private isMissing (value: Value) = match value with VEmpty | VUndefined -> true | _ -> false
let private arg (args: Value array) (index: int) = if index < args.Length && not (isMissing args.[index]) then Some args.[index] else None
let private intArg args index fallback = arg args index |> Option.map toInt |> Option.defaultValue fallback

let wdUndefined = 9999999
let wdToggle = 9999998
let wdColorAutomatic = -16777216

// ---- colours: VBA's are BGR longs (RGB(r, g, b) = r + g * 256 + b * 65536) -------------------------------

let private hexOfLong (value: int) =
    sprintf "#%02X%02X%02X" (value &&& 0xFF) ((value >>> 8) &&& 0xFF) ((value >>> 16) &&& 0xFF)

let private longOfHex (text: string) : int option =
    let t = text.Trim().TrimStart('#')
    if t.Length = 6 && t |> Seq.forall Uri.IsHexDigit then
        let r = Convert.ToInt32(t.Substring(0, 2), 16)
        let g = Convert.ToInt32(t.Substring(2, 2), 16)
        let b = Convert.ToInt32(t.Substring(4, 2), 16)
        Some (r + g * 256 + b * 65536)
    else None

let private colorOfValue (value: Value) : string option =
    let n = toInt value
    if n = wdColorAutomatic || n < 0 || n > 0xFFFFFF then None else Some (hexOfLong n)

let private colorValueOf (color: string option) =
    match color |> Option.bind longOfHex with
    | Some n -> VInteger n
    | None -> VInteger wdColorAutomatic

/// WdColorIndex: wdBlack 1 ... wdGray25 16, as colours and as w:highlight names.
let private colorIndexColors =
    [| "#000000"; "#0000FF"; "#00FFFF"; "#00FF00"; "#FF00FF"; "#FF0000"; "#FFFF00"; "#FFFFFF"
       "#000080"; "#008080"; "#008000"; "#800080"; "#800000"; "#808000"; "#808080"; "#C0C0C0" |]
let private highlightNames =
    [| "black"; "blue"; "cyan"; "green"; "magenta"; "red"; "yellow"; "white"
       "darkBlue"; "darkCyan"; "darkGreen"; "darkMagenta"; "darkRed"; "darkYellow"; "darkGray"; "lightGray" |]

let private colorOfIndex (index: int) =
    if index >= 1 && index <= colorIndexColors.Length then Some colorIndexColors.[index - 1] else None
let private indexOfColor (color: string option) =
    match color with
    | None -> 0
    | Some c ->
        match colorIndexColors |> Array.tryFindIndex (fun x -> eq x c) with
        | Some i -> i + 1
        | None -> wdUndefined
let private highlightOfIndex (index: int) =
    if index >= 1 && index <= highlightNames.Length then Some highlightNames.[index - 1] else None
let private indexOfHighlight (highlight: string option) =
    match highlight with
    | None -> 0
    | Some name -> match highlightNames |> Array.tryFindIndex (fun x -> eq x name) with Some i -> i + 1 | None -> wdUndefined

// ---- units ----------------------------------------------------------------------------------------------

/// The model's lengths are layout pixels (96 per inch); Word's object model speaks points.
let private pxOfPt (pt: float) = pt / 0.75
let private ptOfPx (px: float) = px * 0.75

// ---- Word's constants ------------------------------------------------------------------------------------

let constants : (string * Value) list =
    [ // units, movement
      "wdCharacter", 1; "wdWord", 2; "wdSentence", 3; "wdParagraph", 4; "wdLine", 5; "wdStory", 6
      "wdScreen", 7; "wdSection", 8; "wdColumn", 9; "wdRow", 10; "wdWindow", 11; "wdCell", 12
      "wdCharacterFormatting", 13; "wdParagraphFormatting", 14; "wdTable", 15; "wdItem", 16
      "wdMove", 0; "wdExtend", 1; "wdCollapseStart", 1; "wdCollapseEnd", 0
      // find
      "wdFindStop", 0; "wdFindContinue", 1; "wdFindAsk", 2
      "wdReplaceNone", 0; "wdReplaceOne", 1; "wdReplaceAll", 2
      // paragraphs
      "wdAlignParagraphLeft", 0; "wdAlignParagraphCenter", 1; "wdAlignParagraphRight", 2
      "wdAlignParagraphJustify", 3; "wdAlignParagraphDistribute", 4
      "wdLineSpaceSingle", 0; "wdLineSpace1pt5", 1; "wdLineSpaceDouble", 2; "wdLineSpaceAtLeast", 3
      "wdLineSpaceExactly", 4; "wdLineSpaceMultiple", 5
      "wdOutlineLevelBodyText", 10
      // breaks
      "wdSectionBreakNextPage", 2; "wdSectionBreakContinuous", 3; "wdSectionBreakEvenPage", 4
      "wdSectionBreakOddPage", 5; "wdLineBreak", 6; "wdPageBreak", 7; "wdColumnBreak", 8
      // styles
      "wdStyleNormal", -1; "wdStyleHeading1", -2; "wdStyleHeading2", -3; "wdStyleHeading3", -4
      "wdStyleHeading4", -5; "wdStyleHeading5", -6; "wdStyleHeading6", -7; "wdStyleHeading7", -8
      "wdStyleHeading8", -9; "wdStyleHeading9", -10; "wdStyleTitle", -63; "wdStyleSubtitle", -75
      "wdStyleQuote", -181; "wdStyleIntenseQuote", -182; "wdStyleListParagraph", -180
      "wdStyleNoSpacing", -158; "wdStyleListBullet", -49; "wdStyleListNumber", -50
      "wdStyleStrong", -88; "wdStyleEmphasis", -89; "wdStyleSubtleEmphasis", -261
      "wdStyleIntenseEmphasis", -262; "wdStyleBodyText", -67; "wdStyleCaption", -35
      "wdStyleTypeParagraph", 1; "wdStyleTypeCharacter", 2; "wdStyleTypeTable", 3
      // character formatting
      "wdUndefined", wdUndefined; "wdToggle", wdToggle
      "wdUnderlineNone", 0; "wdUnderlineSingle", 1; "wdUnderlineWords", 2; "wdUnderlineDouble", 3
      "wdUnderlineDotted", 4; "wdUnderlineThick", 6; "wdUnderlineDash", 7; "wdUnderlineDotDash", 9
      "wdUnderlineDotDotDash", 10; "wdUnderlineWavy", 11
      "wdLowerCase", 0; "wdUpperCase", 1; "wdTitleWord", 2; "wdTitleSentence", 4; "wdToggleCase", 5; "wdNextCase", -1
      // colour indexes
      "wdAuto", 0; "wdNoHighlight", 0; "wdBlack", 1; "wdBlue", 2; "wdTurquoise", 3; "wdBrightGreen", 4
      "wdPink", 5; "wdRed", 6; "wdYellow", 7; "wdWhite", 8; "wdDarkBlue", 9; "wdTeal", 10
      "wdGreen", 11; "wdViolet", 12; "wdDarkRed", 13; "wdDarkYellow", 14; "wdGray50", 15; "wdGray25", 16
      // colours (BGR)
      "wdColorAutomatic", wdColorAutomatic; "wdColorBlack", 0; "wdColorWhite", 16777215
      "wdColorRed", 255; "wdColorBlue", 16711680; "wdColorGreen", 32768; "wdColorBrightGreen", 65280
      "wdColorYellow", 65535; "wdColorOrange", 26367; "wdColorDarkRed", 128; "wdColorDarkBlue", 8388608
      "wdColorTeal", 8421376; "wdColorTurquoise", 16776960; "wdColorViolet", 8388736; "wdColorPink", 16711935
      "wdColorDarkGreen", 13056; "wdColorDarkYellow", 32896; "wdColorGold", 52479; "wdColorIndigo", 10040115
      "wdColorLavender", 16751052; "wdColorLightBlue", 16737843; "wdColorLightGreen", 13434828
      "wdColorLightOrange", 39423; "wdColorLightTurquoise", 16777164; "wdColorLightYellow", 10092543
      "wdColorLime", 52377; "wdColorPlum", 6697881; "wdColorRose", 13408767; "wdColorSeaGreen", 6723891
      "wdColorSkyBlue", 16763904; "wdColorTan", 10079487; "wdColorAqua", 13421619; "wdColorBlueGray", 10053222
      "wdColorBrown", 13209; "wdColorDarkTeal", 6697728; "wdColorOliveGreen", 13107; "wdColorPaleBlue", 16764057
      "wdColorGray05", 15987699; "wdColorGray10", 15132390; "wdColorGray125", 14737632; "wdColorGray15", 14277081
      "wdColorGray20", 13421772; "wdColorGray25", 12632256; "wdColorGray30", 11842740; "wdColorGray35", 10921638
      "wdColorGray40", 10066329; "wdColorGray45", 9211020; "wdColorGray50", 8421504; "wdColorGray55", 7566195
      "wdColorGray60", 6710886; "wdColorGray65", 5855577; "wdColorGray70", 5000268; "wdColorGray75", 4210752
      "wdColorGray80", 3355443; "wdColorGray85", 2500134; "wdColorGray90", 1644825; "wdColorGray95", 789516
      // borders
      "wdBorderTop", -1; "wdBorderLeft", -2; "wdBorderBottom", -3; "wdBorderRight", -4
      "wdBorderHorizontal", -5; "wdBorderVertical", -6; "wdBorderDiagonalDown", -7; "wdBorderDiagonalUp", -8
      "wdLineStyleNone", 0; "wdLineStyleSingle", 1; "wdLineStyleDot", 2; "wdLineStyleDashSmallGap", 3
      "wdLineStyleDashLargeGap", 4; "wdLineStyleDashDot", 5; "wdLineStyleDashDotDot", 6; "wdLineStyleDouble", 7
      "wdLineStyleTriple", 8; "wdLineStyleThinThickSmallGap", 9; "wdLineStyleThickThinSmallGap", 10
      "wdLineStyleSingleWavy", 18; "wdLineStyleDoubleWavy", 19
      "wdLineWidth025pt", 2; "wdLineWidth050pt", 4; "wdLineWidth075pt", 6; "wdLineWidth100pt", 8
      "wdLineWidth150pt", 12; "wdLineWidth225pt", 18; "wdLineWidth300pt", 24; "wdLineWidth450pt", 36; "wdLineWidth600pt", 48
      "wdTextureNone", 0; "wdTextureSolid", 1000
      // tables
      "wdWord8TableBehavior", 0; "wdWord9TableBehavior", 1
      "wdAutoFitFixed", 0; "wdAutoFitContent", 1; "wdAutoFitWindow", 2
      "wdCellAlignVerticalTop", 0; "wdCellAlignVerticalCenter", 1; "wdCellAlignVerticalBottom", 3
      "wdAlignRowLeft", 0; "wdAlignRowCenter", 1; "wdAlignRowRight", 2
      "wdPreferredWidthAuto", 1; "wdPreferredWidthPercent", 2; "wdPreferredWidthPoints", 3
      "wdSortOrderAscending", 0; "wdSortOrderDescending", 1
      // information, statistics
      "wdActiveEndAdjustedPageNumber", 1; "wdActiveEndSectionNumber", 2; "wdActiveEndPageNumber", 3
      "wdNumberOfPagesInDocument", 4; "wdFirstCharacterLineNumber", 10; "wdWithInTable", 12
      "wdStartOfRangeRowNumber", 13; "wdEndOfRangeRowNumber", 14; "wdMaximumNumberOfRows", 15
      "wdStartOfRangeColumnNumber", 16; "wdEndOfRangeColumnNumber", 17; "wdMaximumNumberOfColumns", 18
      "wdAtEndOfRowMarker", 31
      "wdStatisticWords", 0; "wdStatisticLines", 1; "wdStatisticPages", 2; "wdStatisticCharacters", 3
      "wdStatisticParagraphs", 4; "wdStatisticCharactersWithSpaces", 5
      // page setup, views, documents
      "wdOrientPortrait", 0; "wdOrientLandscape", 1
      "wdHeaderFooterPrimary", 1; "wdHeaderFooterFirstPage", 2; "wdHeaderFooterEvenPages", 3
      "wdNormalView", 1; "wdOutlineView", 2; "wdPrintView", 3; "wdWebView", 6; "wdReadingView", 7
      "wdDoNotSaveChanges", 0; "wdSaveChanges", -1; "wdPromptToSaveChanges", -2
      "wdFormatDocument", 0; "wdFormatXMLDocument", 12; "wdFormatXMLDocumentMacroEnabled", 13; "wdFormatPDF", 17
      "wdMainTextStory", 1; "wdAlertsNone", 0; "wdAlertsAll", -1 ]
    |> List.map (fun (name, value) -> name, VInteger value)

// ---- the main story ------------------------------------------------------------------------------------

type Mark = ParagraphMark | CellMark | RowMark

/// A paragraph of the story (Para "" for an end-of-row mark): where its text starts, how long it is, the
/// mark after it, and the table cell (table id, 0-based row and column) it is in ("" for the body).
type Segment = { Para: string; Start: int; Length: int; Mark: Mark; Table: string; Row: int; Col: int }

type Story = {
    Segments: Segment array
    /// The story's characters: a paragraph mark is '\r', an end-of-cell mark '\a', an end-of-row mark '\b'.
    Text: string
    Total: int
}

let private buildStory (doc: Document) : Story =
    let segments = ResizeArray<Segment>()
    let text = Text.StringBuilder()
    let add (para: ParagraphBlock option) mark table row col =
        let start = text.Length
        let length =
            match para with
            | Some p -> let s = TB.toString p.Buffer in text.Append(s) |> ignore; s.Length
            | None -> 0
        text.Append(match mark with ParagraphMark -> '\r' | CellMark -> '\a' | RowMark -> '\b') |> ignore
        segments.Add { Para = (match para with Some p -> p.Id | None -> ""); Start = start; Length = length; Mark = mark; Table = table; Row = row; Col = col }
    for block in doc.Blocks do
        match block with
        | Paragraph p | Heading (_, p) -> add (Some p) ParagraphMark "" -1 -1
        | Table t ->
            t.Rows |> List.iteri (fun r row ->
                row.Cells |> List.iteri (fun c cell ->
                    match cell.Content with
                    | [] -> add None CellMark t.Id r c
                    | paras ->
                        let last = paras.Length - 1
                        paras |> List.iteri (fun i p -> add (Some p) (if i = last then CellMark else ParagraphMark) t.Id r c))
                add None RowMark t.Id r row.Cells.Length)
        | TextBox _ -> ()
    if segments.Count = 0 then add None ParagraphMark "" -1 -1
    { Segments = segments.ToArray(); Text = text.ToString(); Total = text.Length }

/// The segment that holds position `pos` (its text or its mark).
let private locate (story: Story) (pos: int) =
    let segs = story.Segments
    let pos = max 0 (min (story.Total - 1) pos)
    let mutable lo = 0
    let mutable hi = segs.Length - 1
    while lo < hi do
        let mid = (lo + hi + 1) / 2
        if segs.[mid].Start <= pos then lo <- mid else hi <- mid - 1
    lo

let private segmentEnd (seg: Segment) = seg.Start + seg.Length + 1

// ---- words and sentences ----------------------------------------------------------------------------------

let private isWordChar (ch: char) = Char.IsLetterOrDigit ch || ch = '_'

/// Word's words: a run of letters (with the apostrophes inside it) or one other character, with the
/// spaces after it; a mark is a word of its own.
let private wordSpans (text: string) : (int * int) array =
    let spans = ResizeArray<int * int>()
    let n = text.Length
    let mutable i = 0
    while i < n do
        let c = text.[i]
        let mutable j = i + 1
        if c = '\r' || c = '\a' || c = '\b' then ()
        elif isWordChar c then
            while j < n && (isWordChar text.[j] || ((text.[j] = '\'' || text.[j] = '’') && j + 1 < n && isWordChar text.[j + 1])) do j <- j + 1
            while j < n && text.[j] = ' ' do j <- j + 1
        else
            while j < n && text.[j] = ' ' do j <- j + 1
        spans.Add((i, j))
        i <- j
    spans.ToArray()

let private sentenceSpans (text: string) : (int * int) array =
    let spans = ResizeArray<int * int>()
    let n = text.Length
    let mutable start = 0
    let mutable i = 0
    while i < n do
        let c = text.[i]
        if c = '\b' && spans.Count > 0 then
            // (An end-of-row mark belongs to the cell's sentence before it.)
            let a, _ = spans.[spans.Count - 1]
            spans.[spans.Count - 1] <- (a, i + 1)
            start <- i + 1
            i <- i + 1
        elif c = '\r' || c = '\a' then
            spans.Add((start, i + 1))
            start <- i + 1
            i <- i + 1
        elif c = '.' || c = '!' || c = '?' then
            let mutable j = i + 1
            while j < n && (text.[j] = '.' || text.[j] = '!' || text.[j] = '?' || text.[j] = '"' || text.[j] = ')' || text.[j] = '”' || text.[j] = '\'') do j <- j + 1
            if j < n && text.[j] = ' ' then
                while j < n && text.[j] = ' ' do j <- j + 1
                spans.Add((start, j))
                start <- j
            i <- j
        else i <- i + 1
    if start < n then spans.Add((start, n))
    spans.ToArray()

// ---- styles ----------------------------------------------------------------------------------------------

let private plainStyle : NamedStyle =
    { Id = ""; Name = None; StyleType = "paragraph"; IsDefault = false; BasedOn = Some "Normal"; Next = Some "Normal"
      Linked = None; IsQuickStyle = true; Aliases = None; UiPriority = None; SemiHidden = false; UnhideWhenUsed = false
      Text = None; TextBold = None; TextItalic = None; TextUnderline = None; TextAllCaps = None; TextSmallCaps = None
      TextStrike = None; SpacingBefore = None; SpacingAfter = None; ContextualSpacing = None; LineSpacing = None
      LineSpacingRule = None; Alignment = None; IndentLeft = None; IndentRight = None; FirstLineIndent = None
      BackgroundColor = None; OutlineLevel = None; KeepNext = None; KeepLines = None; PageBreakBefore = None; WidowControl = None }

let private textOf (size: float option) (family: string option) (color: string option) =
    let style = { TB.defaultStyle with FontFamily = family; Color = color }
    match size with
    | Some s -> TB.withExplicitFontSize (pxOfPt s) style
    | None -> style

/// Word 2016's built-in styles, as a document made in Word has them, for a macro that applies one the
/// document does not define yet: they are added to the document's styles (and saved with it).
let private builtInStyles : NamedStyle list =
    let heading level size color italic before =
        { plainStyle with
            Id = sprintf "Heading%d" level; Name = Some (sprintf "heading %d" level); Linked = Some (sprintf "Heading%dChar" level)
            Text = Some (textOf (Some size) (Some "Calibri Light") (Some color)); TextItalic = (if italic then Some true else None)
            SpacingBefore = Some before; SpacingAfter = Some 0.0; KeepNext = Some true; KeepLines = Some true
            OutlineLevel = Some (level - 1); UiPriority = Some 9 }
    [ heading 1 16.0 "#2F5496" false 12.0
      heading 2 13.0 "#2F5496" false 2.0
      heading 3 12.0 "#1F3763" false 2.0
      heading 4 11.0 "#2F5496" true 2.0
      heading 5 11.0 "#2F5496" false 2.0
      heading 6 11.0 "#1F3763" false 2.0
      { plainStyle with
          Id = "Title"; Name = Some "Title"; Text = Some ({ textOf (Some 28.0) (Some "Calibri Light") None with CharacterSpacing = pxOfPt -0.5 })
          SpacingAfter = Some 0.0; LineSpacing = Some 1.0; LineSpacingRule = Some "auto"; ContextualSpacing = Some true; UiPriority = Some 10 }
      { plainStyle with
          Id = "Subtitle"; Name = Some "Subtitle"; Text = Some ({ textOf None None (Some "#5A5A5A") with CharacterSpacing = pxOfPt 0.75 })
          SpacingAfter = Some 8.0; UiPriority = Some 11 }
      { plainStyle with
          Id = "Quote"; Name = Some "Quote"; Text = Some (textOf None None (Some "#404040")); TextItalic = Some true
          Alignment = Some Center; SpacingBefore = Some 10.0; SpacingAfter = Some 8.0
          IndentLeft = Some (pxOfPt 43.2); IndentRight = Some (pxOfPt 43.2); UiPriority = Some 29 }
      { plainStyle with
          Id = "IntenseQuote"; Name = Some "Intense Quote"; Text = Some (textOf None None (Some "#4472C4")); TextItalic = Some true
          Alignment = Some Center; SpacingBefore = Some 18.0; SpacingAfter = Some 18.0
          IndentLeft = Some (pxOfPt 43.2); IndentRight = Some (pxOfPt 43.2); UiPriority = Some 30 }
      { plainStyle with
          Id = "ListParagraph"; Name = Some "List Paragraph"; IndentLeft = Some (pxOfPt 36.0); ContextualSpacing = Some true; UiPriority = Some 34 }
      { plainStyle with
          Id = "NoSpacing"; Name = Some "No Spacing"; SpacingAfter = Some 0.0; LineSpacing = Some 1.0; LineSpacingRule = Some "auto"; UiPriority = Some 1 }
      { plainStyle with
          Id = "Strong"; Name = Some "Strong"; StyleType = "character"; BasedOn = Some "DefaultParagraphFont"; Next = None
          Text = Some (TB.withExplicitBold true TB.defaultStyle); TextBold = Some true; UiPriority = Some 22 }
      { plainStyle with
          Id = "Emphasis"; Name = Some "Emphasis"; StyleType = "character"; BasedOn = Some "DefaultParagraphFont"; Next = None
          Text = Some (TB.withExplicitItalic true TB.defaultStyle); TextItalic = Some true; UiPriority = Some 20 }
      { plainStyle with
          Id = "IntenseEmphasis"; Name = Some "Intense Emphasis"; StyleType = "character"; BasedOn = Some "DefaultParagraphFont"; Next = None
          Text = Some (TB.withExplicitItalic true (textOf None None (Some "#4472C4"))); TextItalic = Some true; UiPriority = Some 21 }
      { plainStyle with
          Id = "SubtleEmphasis"; Name = Some "Subtle Emphasis"; StyleType = "character"; BasedOn = Some "DefaultParagraphFont"; Next = None
          Text = Some (TB.withExplicitItalic true (textOf None None (Some "#404040"))); TextItalic = Some true; UiPriority = Some 19 } ]

/// A wdBuiltinStyle constant's style name.
let private builtInStyleName (index: int) =
    match index with
    | -1 -> Some "Normal"
    | n when n <= -2 && n >= -10 -> Some (sprintf "Heading %d" (-n - 1))
    | -63 -> Some "Title" | -75 -> Some "Subtitle" | -181 -> Some "Quote" | -182 -> Some "Intense Quote"
    | -180 -> Some "List Paragraph" | -158 -> Some "No Spacing" | -49 -> Some "List Bullet" | -50 -> Some "List Number"
    | -88 -> Some "Strong" | -89 -> Some "Emphasis" | -261 -> Some "Subtle Emphasis" | -262 -> Some "Intense Emphasis"
    | -67 -> Some "Body Text" | -35 -> Some "Caption"
    | _ -> None

let private compact (name: string) = name.Replace(" ", "")

/// The heading level of a paragraph style (Heading1 .. Heading9, or a style with an outline level).
let private headingLevelOf (style: NamedStyle) =
    let id = style.Id
    if id.StartsWith("Heading", StringComparison.OrdinalIgnoreCase) && id.Length = 8 && Char.IsDigit id.[7] then Some (int id.[7] - int '0')
    else
        match style.OutlineLevel with
        | Some level when level >= 0 && level < 9 && (style.Name |> Option.exists (fun n -> n.StartsWith("heading", StringComparison.OrdinalIgnoreCase))) -> Some (level + 1)
        | _ -> None

let private styleDisplayName (style: NamedStyle) =
    match style.Name with
    | Some name when name.StartsWith("heading ", StringComparison.OrdinalIgnoreCase) -> "Heading " + name.Substring 8
    | Some name -> name
    | None -> style.Id

// ---- the document a run reads and writes ---------------------------------------------------------------

/// A range of the story, kept up to date as the run edits the text before and inside it.
[<AllowNullLiteral>]
type LiveRange(start: int, finish: int) =
    member val Start = start with get, set
    member val End = finish with get, set
    /// Find found the range (a further Execute searches on from its end).
    member val Matched = false with get, set
    /// The Selection's active end is its start (it was extended to the left).
    member val ActiveAtStart = false with get, set
    member val Find : obj = null with get, set

/// The document a run reads and writes: the story's positions over it, the ranges that are alive (kept
/// up to date by edits), the Selection, and the formatting of paragraph marks (what text typed into an
/// empty paragraph gets, as in Word).
type MacroDocument(initial: Document, initialHooks: MacroHooks) =
    let mutable doc = initial
    let mutable hooks = initialHooks
    let mutable story : Story option = None
    let mutable words : (Story * (int * int) array * (int * int) array) option = None
    let mutable resolver : (Document * (ParagraphBlock -> TB.TextStyle -> TB.TextStyle)) option = None
    let ranges = ResizeArray<LiveRange>()
    let markStyles = Dictionary<string, TB.TextStyle -> TB.TextStyle>()
    let settings = Dictionary<string, Value>()
    let mutable selection = LiveRange(0, 0)
    let mutable pending : Dictionary<string, Value> = Dictionary<string, Value>()
    let mutable pendingAt = -1
    let mutable bookmarkCounter = 0

    member this.Story : Story =
        match story with
        | Some s -> s
        | None ->
            let s = buildStory doc
            story <- Some s
            s
    member this.Words =
        let st = this.Story
        match words with
        | Some (s, w, _) when obj.ReferenceEquals(s, st) -> w
        | _ ->
            let w = wordSpans st.Text
            words <- Some (st, w, sentenceSpans st.Text)
            w
    member this.Sentences =
        this.Words |> ignore
        match words with Some (_, _, s) -> s | None -> [||]

    member _.Hooks = hooks
    member _.Document = doc
    member _.Settings = settings
    member _.Pending = pending
    member _.PendingAt with get () = pendingAt and set v = pendingAt <- v
    member _.Selection = selection
    member _.NextBookmarkId() = bookmarkCounter <- bookmarkCounter + 1; bookmarkCounter

    /// A change of the text or of the blocks: the story is built again.
    member _.SetStructure(next: Document) = doc <- next; story <- None
    /// A change of formatting only: the story's positions are what they were.
    member _.SetFormatting(next: Document) = doc <- next

    /// Start over on another document (a program's globals stay bound to this object across runs).
    member this.Reset(next: Document, nextHooks: MacroHooks) =
        doc <- next
        hooks <- nextHooks
        story <- None
        words <- None
        resolver <- None
        ranges.Clear()
        // (The formatting of paragraph marks stays while the project does: the model keeps none for a
        // mark, and the paragraphs keep their ids from one run to the next.)
        settings.Clear()
        pending <- Dictionary<string, Value>()
        pendingAt <- -1
        bookmarkCounter <- 1000
        let start, finish =
            match next.Selection with
            | Some sel ->
                let a = this.PositionOf(sel.Anchor.BlockId, sel.Anchor.Offset)
                let f = this.PositionOf(sel.Focus.BlockId, sel.Focus.Offset)
                min a f, max a f
            | None -> 0, 0
        selection <- LiveRange(start, finish)
        ranges.Add selection

    /// A live range of the story.
    member _.Track(start: int, finish: int) =
        let r = LiveRange(start, finish)
        ranges.Add r
        r

    member this.PositionOf(blockId: string, offset: int) =
        let st = this.Story
        match st.Segments |> Array.tryFind (fun s -> s.Para = blockId) with
        | Some seg -> seg.Start + max 0 (min seg.Length offset)
        | None -> 0

    /// The model's position of a story position (an end-of-row mark is the end of its row's last cell).
    member this.ModelPosition(pos: int) : Position =
        let st = this.Story
        let i = locate st pos
        let seg = st.Segments.[i]
        if seg.Para = "" then
            let rec back j = if j < 0 then { BlockId = ""; Offset = 0 } elif st.Segments.[j].Para <> "" then { BlockId = st.Segments.[j].Para; Offset = st.Segments.[j].Length } else back (j - 1)
            back (i - 1)
        else { BlockId = seg.Para; Offset = max 0 (min seg.Length (pos - seg.Start)) }

    member this.Paragraph(id: string) = findParagraph doc id

    member this.UpdateParagraph(id: string, change: ParagraphBlock -> ParagraphBlock) =
        match findParagraph doc id with
        | Some para ->
            let next = change para
            if not (obj.ReferenceEquals(next, para)) then doc <- replaceParagraph doc id next
        | None -> ()

    /// The text of [start, end) as Word gives it (an end-of-cell or end-of-row mark is vbCr & Chr(7)).
    member this.TextOf(start: int, finish: int) =
        let st = this.Story
        let s = max 0 (min st.Total start)
        let e = max s (min st.Total finish)
        st.Text.Substring(s, e - s).Replace("\a", "\r\a").Replace("\b", "\r\a")

    /// The segments [start, end) touches (the one at `start` for an empty range).
    member this.SegmentsIn(start: int, finish: int) : Segment list =
        let st = this.Story
        let first = locate st start
        let last = if finish > start then locate st (finish - 1) else first
        [ for i in first .. last -> st.Segments.[i] ]

    member this.ParagraphIdsIn(start: int, finish: int) =
        this.SegmentsIn(start, finish) |> List.choose (fun s -> if s.Para = "" then None else Some s.Para)

    /// Positions after an edit that replaced [s, e) with `n` characters.
    member private _.Adjust(s: int, e: int, n: int, except: LiveRange) =
        let moveStart p = if p <= s then p elif p >= e then p - (e - s) + n else s
        let moveEnd p = if p <= s then p elif p >= e then p - (e - s) + n else s + n
        for r in ranges do
            if not (obj.ReferenceEquals(r, except)) then
                r.Start <- moveStart r.Start
                r.End <- max r.Start (moveEnd r.End)

    member this.TableSpans() : (string * int * int) list =
        let st = this.Story
        st.Segments
        |> Array.filter (fun s -> s.Table <> "")
        |> Array.groupBy (fun s -> s.Table)
        |> Array.map (fun (id, segs) -> id, segs.[0].Start, segmentEnd segs.[segs.Length - 1])
        |> Array.toList

    member private this.DeleteSpan(s: int, e: int) =
        let st = this.Story
        if s = 0 && e >= st.Total - 1 then
            // (The whole story: one empty paragraph is left, with the last paragraph's properties.)
            let last = st.Segments |> Array.tryFindBack (fun seg -> seg.Para <> "" && seg.Table = "")
            let metadata = last |> Option.bind (fun seg -> findParagraph doc seg.Para) |> Option.map (fun p -> p.Metadata) |> Option.defaultValue defaultMetadata
            let para = { newEmptyParagraph () with Metadata = { metadata with PageBreakBefore = false } }
            match last with
            | Some seg -> (match markStyles.TryGetValue seg.Para with | true, f -> markStyles.[para.Id] <- f | _ -> ())
            | None -> ()
            this.SetStructure { doc with Blocks = [ Paragraph para ] }
        else
            let covered = this.TableSpans() |> List.filter (fun (_, ts, te) -> s <= ts && te <= e)
            let removed = covered |> List.sumBy (fun (_, ts, te) -> te - ts)
            if not covered.IsEmpty then
                let ids = covered |> List.map (fun (id, _, _) -> id) |> Set.ofList
                this.SetStructure { doc with Blocks = doc.Blocks |> List.filter (function Table t -> not (ids.Contains t.Id) | _ -> true) }
            let e = e - removed
            if e > s then
                let st = this.Story
                let first = locate st s
                let last = locate st (e - 1)
                let segs = st.Segments.[first .. last]
                let place (seg: Segment) = seg.Table, seg.Row, seg.Col
                let crossesCells =
                    segs |> Array.exists (fun seg -> seg.Mark <> ParagraphMark && segmentEnd seg <= e)
                    || segs |> Array.exists (fun seg -> place seg <> place segs.[0])
                if crossesCells then
                    // (Cells keep their marks: what is deleted is the text in them.)
                    for seg in Array.rev segs do
                        if seg.Para <> "" then
                            let a = max s seg.Start - seg.Start
                            let b = min e (seg.Start + seg.Length) - seg.Start
                            if b > a then this.UpdateParagraph(seg.Para, fun p -> { p with Buffer = TB.delete a (b - a) p.Buffer })
                elif first = last then
                    let seg = segs.[0]
                    this.UpdateParagraph(seg.Para, fun p -> { p with Buffer = TB.delete (s - seg.Start) (e - s) p.Buffer })
                else
                    let a = this.ModelPosition s
                    let b = this.ModelPosition e
                    doc <- deleteSelection (setSelection a b doc)
                story <- None

    /// The style that follows a paragraph of `styleId` when a paragraph mark is typed at its end.
    member private this.NextStyle(styleId: string option) : string option option =
        match styleId with
        | None -> None
        | Some id ->
            let named =
                doc.Styles |> List.tryFind (fun s -> s.Id = id && s.StyleType = "paragraph")
                |> Option.orElse (builtInStyles |> List.tryFind (fun s -> s.Id = id))
            match named |> Option.bind (fun s -> s.Next) with
            | Some next when next = id -> None
            | Some next when eq next "Normal" -> Some (if doc.Styles |> List.exists (fun s -> s.Id = "Normal") then Some "Normal" else None)
            | Some next -> Some (Some next)
            | None -> None

    /// Insert `text` at `pos` (vbCr / vbLf / vbCrLf are paragraph marks, Chr(11) a line break, Chr(12) a
    /// page break); the number of positions it takes.
    member private this.InsertAt(pos: int, text: string) : int =
        let text = text.Replace("\r\n", "\r").Replace('\n', '\r').Replace("\a", "").Replace("\b", "")
        if text = "" then 0
        else
            let target = this.ModelPosition(min pos (this.Story.Total - 1))
            if target.BlockId = "" then notAvailable "no paragraph to insert into"
            let mutable para = target.BlockId
            let mutable off = target.Offset
            let flush (piece: string) =
                if piece <> "" then
                    let piece = piece.Replace('\v', '\n')
                    let at = off
                    this.UpdateParagraph(para, fun p ->
                        let wasEmpty = TB.length p.Buffer = 0
                        let buffer = TB.insert at piece p.Buffer
                        let buffer =
                            match markStyles.TryGetValue p.Id with
                            | true, f when wasEmpty -> TB.applyStyle 0 piece.Length f buffer
                            | _ -> buffer
                        { p with Buffer = buffer })
                    off <- off + piece.Length
            let split (pageBreak: bool) =
                let atEnd = match findParagraph doc para with Some p -> off >= TB.length p.Buffer | None -> false
                let next = splitBlock (setCursor para off doc)
                match next.Selection with
                | Some sel when sel.Focus.BlockId <> para ->
                    doc <- next
                    let newId = sel.Focus.BlockId
                    match markStyles.TryGetValue para with
                    | true, f -> markStyles.[newId] <- f
                    | _ -> ()
                    if atEnd then
                        match this.NextStyle((findParagraph doc newId |> Option.bind (fun p -> p.Metadata.Style))) with
                        | Some style -> this.ApplyParagraphStyle(newId, style, None)
                        | None -> ()
                    if pageBreak then this.UpdateParagraph(newId, fun p -> { p with Metadata = { p.Metadata with PageBreakBefore = true } })
                    para <- newId
                    off <- 0
                | _ -> notAvailable "the paragraph cannot be split here"
            let piece = Text.StringBuilder()
            for ch in text do
                if ch = '\r' || ch = '\f' then
                    flush (piece.ToString())
                    piece.Clear() |> ignore
                    split (ch = '\f')
                else piece.Append ch |> ignore
            flush (piece.ToString())
            story <- None
            text.Length

    /// Replace [start, end) with `text`; the range it then takes. The last paragraph mark stays (a range
    /// at the very end inserts before it).
    member this.Replace(start: int, finish: int, text: string, ?owner: LiveRange) : int * int =
        let total = this.Story.Total
        let s = max 0 (min (total - 1) start)
        let e = max s (min (total - 1) finish)
        if e > s then this.DeleteSpan(s, e)
        let n = this.InsertAt(s, text)
        this.Adjust(s, e, n, defaultArg owner null)
        s, s + n

    /// Character formatting of [start, end) (and of the paragraph marks in it).
    member this.Format(start: int, finish: int, change: TB.TextStyle -> TB.TextStyle) =
        for seg in this.SegmentsIn(start, finish) do
            if seg.Para <> "" && finish > start then
                let a = max start seg.Start - seg.Start
                let b = min finish (seg.Start + seg.Length) - seg.Start
                if b > a then this.UpdateParagraph(seg.Para, fun p -> { p with Buffer = TB.applyStyle a (b - a) change p.Buffer })
                let mark = seg.Start + seg.Length
                if mark >= start && mark < finish then
                    markStyles.[seg.Para] <- (match markStyles.TryGetValue seg.Para with | true, f -> f >> change | _ -> change)

    /// The characters of [start, end) mapped (Range.Case): the formatting stays.
    member this.MapText(start: int, finish: int, map: string -> int -> char) =
        for seg in this.SegmentsIn(start, finish) do
            if seg.Para <> "" then
                let a = max start seg.Start - seg.Start
                let b = min finish (seg.Start + seg.Length) - seg.Start
                if b > a then this.UpdateParagraph(seg.Para, fun p -> { p with Buffer = TB.mapCharacters a (b - a) map p.Buffer })
        words <- None
        story <- None

    member this.Resolver =
        match resolver with
        | Some (d, r) when obj.ReferenceEquals(d, doc) -> r
        | _ ->
            let r = WordStyles.effectiveRunStyleResolver doc
            resolver <- Some (doc, r)
            r

    /// The effective character formatting of the text of [start, end) (that of the character at `start`
    /// for an empty range).
    member this.StylesIn(start: int, finish: int) : TB.TextStyle list =
        let resolve = this.Resolver
        let st = this.Story
        let result = ResizeArray<TB.TextStyle>()
        let finish = if finish > start then finish else start + 1
        for seg in this.SegmentsIn(start, finish) do
            // (An end-of-row mark has the formatting of its row's last paragraph.)
            let para =
                if seg.Para <> "" then seg.Para
                else
                    let i = locate st seg.Start
                    if i > 0 then st.Segments.[i - 1].Para else ""
            let seg = if seg.Para = "" then { seg with Para = para; Length = 0 } else seg
            match (if seg.Para = "" then None else findParagraph doc seg.Para) with
            | Some p ->
                let a = max start seg.Start - seg.Start
                let b = min finish (seg.Start + seg.Length) - seg.Start
                if seg.Length = 0 || a >= seg.Length then
                    // (The paragraph mark: the formatting text typed there gets.)
                    let last = p.Buffer.Runs |> List.tryPick (function TB.Text tr -> Some tr.Style | _ -> None) |> Option.defaultValue TB.defaultStyle
                    let mark = match markStyles.TryGetValue p.Id with | true, f -> f last | _ -> last
                    result.Add(resolve p mark)
                else
                    for run in p.Buffer.Runs do
                        match run with
                        | TB.Text tr when tr.Start < b && tr.Start + tr.Length > a -> result.Add(resolve p tr.Style)
                        | TB.Text tr when b = a && tr.Start <= a && a <= tr.Start + tr.Length -> result.Add(resolve p tr.Style)
                        | _ -> ()
            | None -> ()
        ignore st
        if result.Count = 0 then [ TB.defaultStyle ] else List.ofSeq result

    /// Change the metadata of the paragraphs [start, end) touches.
    member this.MapParagraphs(start: int, finish: int, change: BlockMetadata -> BlockMetadata) =
        for id in this.ParagraphIdsIn(start, finish) |> List.distinct do
            this.UpdateParagraph(id, fun p ->
                let m = change p.Metadata
                if m = p.Metadata then p else { p with Metadata = m })

    /// The paragraph style with this name or id (a built-in one the document does not have is added).
    member this.FindStyle(name: string) : NamedStyle option =
        let byName (s: NamedStyle) =
            eq s.Id name || eq (compact s.Id) (compact name)
            || (s.Name |> Option.exists (fun n -> eq n name || eq (compact n) (compact name)))
        match doc.Styles |> List.tryFind byName with
        | Some s -> Some s
        | None ->
            match builtInStyles |> List.tryFind byName with
            | Some s ->
                this.SetFormatting { doc with Styles = doc.Styles @ [ s ] }
                Some s
            | None -> None

    /// A new paragraph's mark has the formatting of another's (text typed into it gets it).
    member _.InheritMark(fromId: string, toId: string) =
        match markStyles.TryGetValue fromId with
        | true, f -> markStyles.[toId] <- f
        | _ -> ()

    member this.ChangeStyle(id: string, change: NamedStyle -> NamedStyle) =
        this.SetFormatting { doc with Styles = doc.Styles |> List.map (fun s -> if s.Id = id then change s else s) }

    /// Give a paragraph a paragraph style (None: the default one); a heading style makes a body paragraph
    /// a heading.
    member this.ApplyParagraphStyle(paraId: string, styleId: string option, level: int option) =
        let level =
            match level, styleId with
            | Some l, _ -> Some l
            | None, Some id -> doc.Styles |> List.tryFind (fun s -> s.Id = id) |> Option.bind headingLevelOf
            | None, None -> None
        // (A paragraph style takes the place of the paragraph's own formatting, as in Word; the run
        // formatting the old style lent the runs goes - WordStyles.setParagraphStyle.)
        doc <- WordStyles.setParagraphStyle (defaultArg styleId "") (setCursor paraId 0 doc)
        let paragraphFlags =
            ParagraphExplicit.SpacingBefore ||| ParagraphExplicit.SpacingAfter ||| ParagraphExplicit.LineSpacing
            ||| ParagraphExplicit.IndentLeft ||| ParagraphExplicit.IndentRight ||| ParagraphExplicit.FirstLineIndent ||| ParagraphExplicit.ContextualSpacing
        let restyle (p: ParagraphBlock) =
            let m = p.Metadata
            { p with
                Metadata =
                    { m with
                        Style = styleId; SpacingBefore = None; SpacingAfter = None; LineSpacing = None; LineSpacingRule = None
                        IndentLeft = None; IndentRight = None; FirstLineIndent = None; ContextualSpacing = None
                        Alignment = Left; ExplicitAlignment = false; KeepNext = None; KeepLines = None
                        Explicit = m.Explicit &&& ~~~paragraphFlags } }
        let mutable found = false
        let blocks =
            doc.Blocks |> List.map (fun block ->
                match block with
                | Paragraph p | Heading (_, p) when p.Id = paraId ->
                    found <- true
                    match level with
                    | Some l -> Heading (max 1 (min 3 l), restyle p)
                    | None -> Paragraph (restyle p)
                | other -> other)
        if found then this.SetFormatting { doc with Blocks = blocks }
        else this.UpdateParagraph(paraId, restyle)

    /// The document as the run leaves it: its selection where the macro's Selection is.
    member this.Result() : Document =
        let a = this.ModelPosition selection.Start
        let f = this.ModelPosition selection.End
        if a.BlockId = "" then doc else { doc with Selection = Some { Anchor = a; Focus = f } }

// ---- the object model --------------------------------------------------------------------------------------

/// The parameters of Word's methods, for calls with named arguments: by (object type, method), lower case.
let private parameterNames : Map<string * string, string list> =
    let rangeMethods =
        [ "collapse", [ "Direction" ]
          "move", [ "Unit"; "Count" ]
          "movestart", [ "Unit"; "Count" ]
          "moveend", [ "Unit"; "Count" ]
          "expand", [ "Unit" ]
          "startof", [ "Unit"; "Extend" ]
          "endof", [ "Unit"; "Extend" ]
          "next", [ "Unit"; "Count" ]
          "previous", [ "Unit"; "Count" ]
          "insertbreak", [ "Type" ]
          "setrange", [ "Start"; "End" ]
          "delete", [ "Unit"; "Count" ]
          "information", [ "Type" ]
          "computestatistics", [ "Statistic"; "IncludeFootnotesAndEndnotes" ]
          "converttotable", [ "Separator"; "NumRows"; "NumColumns"; "InitialColumnWidth"; "Format"; "ApplyBorders" ]
          "insertafter", [ "Text" ]
          "insertbefore", [ "Text" ]
          "typetext", [ "Text" ]
          "homekey", [ "Unit"; "Extend" ]
          "endkey", [ "Unit"; "Extend" ]
          "moveright", [ "Unit"; "Count"; "Extend" ]
          "moveleft", [ "Unit"; "Count"; "Extend" ]
          "moveup", [ "Unit"; "Count"; "Extend" ]
          "movedown", [ "Unit"; "Count"; "Extend" ] ]
    [ for (m, ps) in rangeMethods do
        yield ("Range", m), ps
        yield ("Selection", m), ps
        yield ("Paragraph", m), ps
      yield ("Find", "execute"), [ "FindText"; "MatchCase"; "MatchWholeWord"; "MatchWildcards"; "MatchSoundsLike"; "MatchAllWordForms"
                                   "Forward"; "Wrap"; "Format"; "ReplaceWith"; "Replace"; "MatchKashida"; "MatchDiacritics"; "MatchAlefHamza"; "MatchControl" ]
      yield ("Tables", "add"), [ "Range"; "NumRows"; "NumColumns"; "DefaultTableBehavior"; "AutoFitBehavior" ]
      yield ("Table", "cell"), [ "Row"; "Column" ]
      yield ("Table", "autofitbehavior"), [ "Behavior" ]
      yield ("Table", "sort"), [ "ExcludeHeader"; "FieldNumber"; "SortFieldType"; "SortOrder" ]
      yield ("Rows", "add"), [ "BeforeRow" ]
      yield ("Columns", "add"), [ "BeforeColumn" ]
      yield ("Bookmarks", "add"), [ "Name"; "Range" ]
      yield ("Documents", "add"), [ "Template"; "NewTemplate"; "DocumentType"; "Visible" ]
      yield ("Document", "range"), [ "Start"; "End" ]
      yield ("Document", "saveas"), [ "FileName"; "FileFormat" ]
      yield ("Document", "saveas2"), [ "FileName"; "FileFormat" ]
      yield ("Document", "close"), [ "SaveChanges" ]
      yield ("Paragraphs", "add"), [ "Range" ] ]
    |> List.map (fun ((t, m), ps) -> (t, m), ps |> List.map (fun p -> p.ToLowerInvariant()))
    |> Map.ofList

type private Host(typeName: string, get: string -> Value option, set: string -> Value -> bool,
                  call: string -> Value array -> Value option, enumerate: (unit -> Value seq) option) =
    member _.TypeName = typeName
    interface IHostObject with
        member _.TypeName = typeName
        member _.GetMember name = get name
        member _.SetMember name value = set name value
        member _.CallMethod name args = call name args
        member _.CallNamed name args names =
            let kind = match typeName with "ActiveDocument" | "ThisDocument" -> "Document" | other -> other
            match Map.tryFind (kind, name.ToLowerInvariant()) parameterNames with
            | None -> call name args
            | Some parameters ->
                let slots = ResizeArray<Value>()
                let put index value =
                    while slots.Count <= index do slots.Add VUndefined
                    slots.[index] <- value
                args |> Array.iteri (fun i a ->
                    let argName = if i < names.Length then names.[i] else ""
                    if argName = "" then put i a
                    else
                        match parameters |> List.tryFindIndex ((=) (argName.ToLowerInvariant())) with
                        | Some index -> put index a
                        | None -> raiseVbErrorWith 448 ("Named argument not found: " + argName))
                call name (slots.ToArray())
        member _.Enumerate () = enumerate |> Option.map (fun items -> items ())

/// An object with members by (lower-case) name: `get` answers None for a member it does not have, `set`
/// false; `call` handles methods and falls back to `get` for a property read with arguments.
let private makeObject typeName (get: string -> Value option) (set: string -> Value -> bool) (call: string -> Value array -> Value option) (items: (unit -> Value seq) option) =
    let get' (name: string) = get (name.ToLowerInvariant())
    let set' (name: string) value = if set (name.ToLowerInvariant()) value then true else unsupported typeName name
    let call' (name: string) (args: Value array) =
        let lower = name.ToLowerInvariant()
        match call lower args with
        | Some v -> Some v
        | None ->
            match get lower with
            | Some v when args.Length = 0 -> Some v
            | Some _ -> None
            | None -> unsupported typeName name
    VHostObject (Host(typeName, get', set', call', items) :> IHostObject)

let private noSet (_: string) (_: Value) = false
let private noCall (_: string) (_: Value array) : Value option = None

/// A collection: Count, Item(i) (1-based) / its default member, First, Last, For Each; `more` for its own members.
let private collectionObject typeName (items: unit -> Value array) (moreGet: string -> Value option) (moreSet: string -> Value -> bool) (moreCall: string -> Value array -> Value option) =
    let item (args: Value array) =
        match arg args 0 with
        | Some index ->
            let all = items ()
            let i = toInt index
            if i >= 1 && i <= all.Length then Some all.[i - 1] else noSuchMember ()
        | None -> None
    makeObject typeName
        (fun name ->
            match name with
            | "count" -> Some (VInteger (items ()).Length)
            | "first" -> (match items () with [||] -> noSuchMember () | all -> Some all.[0])
            | "last" -> (match items () with [||] -> noSuchMember () | all -> Some all.[all.Length - 1])
            | _ -> moreGet name)
        moreSet
        (fun name args ->
            match name with
            | "" | "_default" | "item" when args.Length > 0 ->
                match moreCall name args with
                | Some v -> Some v
                | None -> item args
            | _ -> moreCall name args)
        (Some (fun () -> Seq.ofArray (items ())))

// ---- character formatting ------------------------------------------------------------------------------------

let private underlineStyles = [ 1, None; 2, Some "words"; 3, Some "double"; 4, Some "dotted"; 6, Some "thick"; 7, Some "dash"; 9, Some "dotDash"; 10, Some "dotDotDash"; 11, Some "wave" ]

let private toggleOr (value: Value) (current: bool) =
    let n = match scalarOf value with VBoolean b -> (if b then -1 else 0) | other -> toInt other
    if n = wdToggle then not current else n <> 0

/// What setting a Font property does to a run's formatting; None for a property Font does not have.
let private fontChange (name: string) (value: Value) : (TB.TextStyle -> TB.TextStyle) option =
    match name with
    | "bold" -> Some (fun s -> TB.withExplicitBold (toggleOr value s.Bold) s)
    | "italic" -> Some (fun s -> TB.withExplicitItalic (toggleOr value s.Italic) s)
    | "underline" ->
        let n = toInt value
        Some (fun s ->
            if n = 0 then TB.withExplicitUnderline false { s with UnderlineStyle = None }
            else
                let style = underlineStyles |> List.tryPick (fun (k, v) -> if k = n then Some v else None) |> Option.defaultValue None
                TB.withExplicitUnderline true { s with UnderlineStyle = style })
    | "size" -> let size = pxOfPt (toFloat value) in Some (TB.withExplicitFontSize size)
    | "name" | "nameascii" | "nameother" -> let family = toText value in Some (fun s -> { s with FontFamily = Some family })
    | "color" -> let color = colorOfValue value in Some (fun s -> { s with Color = color })
    | "colorindex" -> let color = colorOfIndex (toInt value) in Some (fun s -> { s with Color = color })
    | "strikethrough" -> Some (fun s -> TB.withExplicitStrike (toggleOr value s.Strike) s)
    | "doublestrikethrough" -> Some (fun s -> TB.withExplicitDoubleStrike (toggleOr value s.DoubleStrike) s)
    | "superscript" -> let on = toBool value in Some (fun s -> { s with VertAlign = (if on then Some "superscript" elif s.VertAlign = Some "superscript" then None else s.VertAlign) })
    | "subscript" -> let on = toBool value in Some (fun s -> { s with VertAlign = (if on then Some "subscript" elif s.VertAlign = Some "subscript" then None else s.VertAlign) })
    | "allcaps" -> Some (fun s -> TB.withExplicitAllCaps (toggleOr value s.AllCaps) s)
    | "smallcaps" -> Some (fun s -> TB.withExplicitSmallCaps (toggleOr value s.SmallCaps) s)
    | "hidden" -> Some (fun s -> TB.withExplicitHidden (toggleOr value s.Hidden) s)
    | "spacing" -> let px = pxOfPt (toFloat value) in Some (fun s -> { s with CharacterSpacing = px })
    | "position" -> let pt = toFloat value in Some (fun s -> { s with Position = pt })
    | _ -> None

let private fontRead (name: string) (s: TB.TextStyle) : Value option =
    match name with
    | "bold" -> Some (VInteger (if s.Bold then -1 else 0))
    | "italic" -> Some (VInteger (if s.Italic then -1 else 0))
    | "underline" ->
        if not s.Underline then Some (VInteger 0)
        else Some (VInteger (underlineStyles |> List.tryPick (fun (k, v) -> if v = s.UnderlineStyle then Some k else None) |> Option.defaultValue 1))
    | "size" -> Some (VDouble (Math.Round(ptOfPx s.FontSize * 2.0) / 2.0))
    | "name" | "nameascii" | "nameother" -> Some (VString (defaultArg s.FontFamily "Calibri"))
    | "color" -> Some (colorValueOf s.Color)
    | "colorindex" -> Some (VInteger (indexOfColor s.Color))
    | "strikethrough" -> Some (VInteger (if s.Strike then -1 else 0))
    | "doublestrikethrough" -> Some (VInteger (if s.DoubleStrike then -1 else 0))
    | "superscript" -> Some (VInteger (if s.VertAlign = Some "superscript" then -1 else 0))
    | "subscript" -> Some (VInteger (if s.VertAlign = Some "subscript" then -1 else 0))
    | "allcaps" -> Some (VInteger (if s.AllCaps then -1 else 0))
    | "smallcaps" -> Some (VInteger (if s.SmallCaps then -1 else 0))
    | "hidden" -> Some (VInteger (if s.Hidden then -1 else 0))
    | "spacing" -> Some (VDouble (ptOfPx s.CharacterSpacing))
    | "position" -> Some (VDouble s.Position)
    | _ -> None

/// What a Font is of: a range's text, a set of settings (Find's, Replacement's, the Selection's at a
/// caret), or a style.
type private FontTarget =
    | RangeFont of LiveRange
    | SpecFont of Dictionary<string, Value>
    | StyleFont of string

/// The change a set of Font settings makes.
let private specChange (spec: Dictionary<string, Value>) : TB.TextStyle -> TB.TextStyle =
    let changes = [ for kv in spec -> (fontChange kv.Key kv.Value).Value ]
    fun s -> changes |> List.fold (fun acc f -> f acc) s

let private fontObject (book: MacroDocument) (target: FontTarget) : Value =
    makeObject "Font"
        (fun name ->
            match target with
            | RangeFont r ->
                let styles = book.StylesIn(r.Start, r.End)
                match fontRead name styles.Head with
                | Some first ->
                    let all = styles |> List.map (fun s -> (fontRead name s).Value) |> List.distinct
                    Some (if all.Length = 1 then first else VInteger wdUndefined)
                | None -> None
            | SpecFont spec ->
                match spec.TryGetValue name with
                | true, v -> Some v
                | _ -> fontRead name TB.defaultStyle |> Option.map (fun _ -> VInteger wdUndefined)
            | StyleFont id ->
                let named = book.Document.Styles |> List.tryFind (fun s -> s.Id = id)
                let text = named |> Option.bind (fun s -> s.Text) |> Option.defaultValue TB.defaultStyle
                let text =
                    match named with
                    | Some s -> { text with Bold = defaultArg s.TextBold text.Bold; Italic = defaultArg s.TextItalic text.Italic }
                    | None -> text
                fontRead name text)
        (fun name value ->
            match fontChange name value with
            | None -> false
            | Some change ->
                match target with
                | RangeFont r -> book.Format(r.Start, r.End, change); true
                | SpecFont spec -> spec.[name] <- value; true
                | StyleFont id ->
                    book.ChangeStyle(id, fun s ->
                        let text = change (defaultArg s.Text TB.defaultStyle)
                        { s with
                            Text = Some text
                            TextBold = (if name = "bold" then Some text.Bold else s.TextBold)
                            TextItalic = (if name = "italic" then Some text.Italic else s.TextItalic) })
                    true)
        (fun name _ ->
            match name, target with
            | "reset", RangeFont r -> book.Format(r.Start, r.End, (fun s -> { TB.defaultStyle with Style = s.Style; Hyperlink = s.Hyperlink; BookmarkStarts = s.BookmarkStarts; BookmarkEnds = s.BookmarkEnds; SourceRun = s.SourceRun })); Some VEmpty
            | "reset", SpecFont spec -> spec.Clear(); Some VEmpty
            | "grow", RangeFont r -> book.Format(r.Start, r.End, (fun s -> TB.withExplicitFontSize (s.FontSize + pxOfPt 1.0) s)); Some VEmpty
            | "shrink", RangeFont r -> book.Format(r.Start, r.End, (fun s -> TB.withExplicitFontSize (max 1.0 (s.FontSize - pxOfPt 1.0)) s)); Some VEmpty
            | _ -> None)
        None

// ---- paragraph formatting ------------------------------------------------------------------------------------

let private alignmentOf (n: int) = match n with 1 -> Center | 2 -> Right | 3 | 4 -> Justify | _ -> Left
let private alignmentValue (a: Alignment) = match a with Left -> 0 | Center -> 1 | Right -> 2 | Justify -> 3

let private withFlag flag on (m: BlockMetadata) = { m with Explicit = (if on then m.Explicit ||| flag else m.Explicit &&& ~~~flag) }

/// What setting a ParagraphFormat property does to a paragraph's metadata.
let private paragraphChange (name: string) (value: Value) : (BlockMetadata -> BlockMetadata) option =
    match name with
    | "alignment" -> let a = alignmentOf (toInt value) in Some (fun m -> { m with Alignment = a; ExplicitAlignment = true })
    | "spacebefore" -> let v = toFloat value in Some (fun m -> { m with SpacingBefore = Some v } |> withFlag ParagraphExplicit.SpacingBefore true)
    | "spaceafter" -> let v = toFloat value in Some (fun m -> { m with SpacingAfter = Some v } |> withFlag ParagraphExplicit.SpacingAfter true)
    | "leftindent" -> let v = pxOfPt (toFloat value) in Some (fun m -> { m with IndentLeft = Some v } |> withFlag ParagraphExplicit.IndentLeft true)
    | "rightindent" -> let v = pxOfPt (toFloat value) in Some (fun m -> { m with IndentRight = Some v } |> withFlag ParagraphExplicit.IndentRight true)
    | "firstlineindent" -> let v = pxOfPt (toFloat value) in Some (fun m -> { m with FirstLineIndent = Some v } |> withFlag ParagraphExplicit.FirstLineIndent true)
    | "linespacingrule" ->
        let rule = toInt value
        Some (fun m ->
            let multiple v = { m with LineSpacing = Some v; LineSpacingRule = Some "auto" } |> withFlag ParagraphExplicit.LineSpacing true
            match rule with
            | 0 -> multiple 1.0
            | 1 -> multiple 1.5
            | 2 -> multiple 2.0
            | 3 -> { m with LineSpacingRule = Some "atLeast"; LineSpacing = Some (defaultArg (m.LineSpacing |> Option.filter (fun v -> m.LineSpacingRule = Some "atLeast" || m.LineSpacingRule = Some "exact")) (pxOfPt 12.0)) } |> withFlag ParagraphExplicit.LineSpacing true
            | 4 -> { m with LineSpacingRule = Some "exact"; LineSpacing = Some (defaultArg (m.LineSpacing |> Option.filter (fun v -> m.LineSpacingRule = Some "atLeast" || m.LineSpacingRule = Some "exact")) (pxOfPt 12.0)) } |> withFlag ParagraphExplicit.LineSpacing true
            | _ -> { m with LineSpacingRule = Some "auto" } |> withFlag ParagraphExplicit.LineSpacing true)
    | "linespacing" ->
        let pt = toFloat value
        Some (fun m ->
            match m.LineSpacingRule with
            | Some "exact" | Some "atLeast" -> { m with LineSpacing = Some (pxOfPt pt) } |> withFlag ParagraphExplicit.LineSpacing true
            | _ -> { m with LineSpacing = Some (pt / 12.0); LineSpacingRule = Some "auto" } |> withFlag ParagraphExplicit.LineSpacing true)
    | "keepwithnext" -> let v = toBool value in Some (fun m -> { m with KeepNext = Some v })
    | "keeptogether" -> let v = toBool value in Some (fun m -> { m with KeepLines = Some v })
    | "pagebreakbefore" -> let v = toBool value in Some (fun m -> { m with PageBreakBefore = v })
    | "widowcontrol" -> let v = toBool value in Some (fun m -> { m with WidowControl = Some v })
    | "spacebeforeauto" -> let v = toBool value in Some (fun m -> if v then { m with SpacingBefore = Some 14.0 } |> withFlag ParagraphExplicit.SpacingBefore true else m)
    | "spaceafterauto" -> let v = toBool value in Some (fun m -> if v then { m with SpacingAfter = Some 14.0 } |> withFlag ParagraphExplicit.SpacingAfter true else m)
    | _ -> None

/// What a paragraph's style gives it where it states nothing itself (resolved: based-on styles and the
/// document defaults in), and its heading level.
type private StyleFallback = {
    Fallback: WordStyles.ParagraphStyle option
    Level: int option
}
let private noFallback = { Fallback = None; Level = None }
let private fallbackOfNamed (s: NamedStyle) : StyleFallback =
    { Fallback =
        Some { WordStyles.normal with
                SpacingBefore = s.SpacingBefore; SpacingAfter = s.SpacingAfter; LineSpacing = s.LineSpacing; LineSpacingRule = s.LineSpacingRule
                Alignment = s.Alignment; IndentLeft = s.IndentLeft; IndentRight = s.IndentRight; FirstLineIndent = s.FirstLineIndent
                KeepNext = s.KeepNext; KeepLines = s.KeepLines }
      Level = headingLevelOf s }

let rec private paragraphRead (style: StyleFallback) (name: string) (m: BlockMetadata) : Value option =
    let fromStyle (f: WordStyles.ParagraphStyle -> 'a option) = style.Fallback |> Option.bind f
    match name with
    | "alignment" ->
        let a = if m.ExplicitAlignment then m.Alignment else (fromStyle (fun s -> s.Alignment) |> Option.defaultValue m.Alignment)
        Some (VInteger (alignmentValue a))
    | "spacebefore" -> Some (VDouble (m.SpacingBefore |> Option.orElse (fromStyle (fun s -> s.SpacingBefore)) |> Option.defaultValue 0.0))
    | "spaceafter" -> Some (VDouble (m.SpacingAfter |> Option.orElse (fromStyle (fun s -> s.SpacingAfter)) |> Option.defaultValue 0.0))
    | "leftindent" -> Some (VDouble (ptOfPx (m.IndentLeft |> Option.orElse (fromStyle (fun s -> s.IndentLeft)) |> Option.defaultValue 0.0)))
    | "rightindent" -> Some (VDouble (ptOfPx (m.IndentRight |> Option.orElse (fromStyle (fun s -> s.IndentRight)) |> Option.defaultValue 0.0)))
    | "firstlineindent" -> Some (VDouble (ptOfPx (m.FirstLineIndent |> Option.orElse (fromStyle (fun s -> s.FirstLineIndent)) |> Option.defaultValue 0.0)))
    | "linespacingrule" ->
        match m.LineSpacingRule, m.LineSpacing with
        | _ when m.LineSpacingRule.IsNone && (fromStyle (fun s -> s.LineSpacingRule)).IsSome && m.LineSpacing.IsNone ->
            paragraphRead { style with Fallback = None } name { m with LineSpacingRule = fromStyle (fun s -> s.LineSpacingRule); LineSpacing = fromStyle (fun s -> s.LineSpacing) }
        | Some "exact", _ -> Some (VInteger 4)
        | Some "atLeast", _ -> Some (VInteger 3)
        | _, Some 1.5 -> Some (VInteger 1)
        | _, Some 2.0 -> Some (VInteger 2)
        | _, (Some 1.0 | None) -> Some (VInteger 0)
        | _ -> Some (VInteger 5)
    | "linespacing" when m.LineSpacing.IsNone && (fromStyle (fun s -> s.LineSpacing)).IsSome ->
        paragraphRead { style with Fallback = None } name { m with LineSpacingRule = fromStyle (fun s -> s.LineSpacingRule); LineSpacing = fromStyle (fun s -> s.LineSpacing) }
    | "linespacing" ->
        match m.LineSpacingRule, m.LineSpacing with
        | (Some "exact" | Some "atLeast"), Some px -> Some (VDouble (ptOfPx px))
        | _, Some multiple -> Some (VDouble (multiple * 12.0))
        | _ -> Some (VDouble 12.0)
    | "keepwithnext" -> Some (VBoolean (m.KeepNext |> Option.orElse (fromStyle (fun s -> s.KeepNext)) |> Option.defaultValue false))
    | "keeptogether" -> Some (VBoolean (m.KeepLines |> Option.orElse (fromStyle (fun s -> s.KeepLines)) |> Option.defaultValue false))
    | "pagebreakbefore" -> Some (VBoolean m.PageBreakBefore)
    | "widowcontrol" -> Some (VBoolean (m.WidowControl |> Option.defaultValue true))
    | "outlinelevel" -> Some (VInteger (match style.Level with Some l -> l | None -> 10))
    | "spacebeforeauto" | "spaceafterauto" -> Some (VBoolean false)
    | _ -> None

// ---- shading and borders ---------------------------------------------------------------------------------------

type private CellRef = { TableId: string; Row: int; Col: int }

let private mapTable (book: MacroDocument) (tableId: string) (change: TableBlock -> TableBlock) =
    let doc = book.Document
    book.SetFormatting { doc with Blocks = doc.Blocks |> List.map (function Table t when t.Id = tableId -> Table (change t) | other -> other) }

let private tableOf (book: MacroDocument) (tableId: string) =
    book.Document.Blocks |> List.tryPick (function Table t when t.Id = tableId -> Some t | _ -> None)

let private mapCells (book: MacroDocument) (cells: CellRef list) (change: TableCell -> TableCell) =
    for tableId, group in cells |> List.groupBy (fun c -> c.TableId) do
        let wanted = group |> List.map (fun c -> c.Row, c.Col) |> Set.ofList
        mapTable book tableId (fun t ->
            { t with Rows = t.Rows |> List.mapi (fun r row -> { row with Cells = row.Cells |> List.mapi (fun c cell -> if wanted.Contains (r, c) then change cell else cell) }) })

let private cellOf (book: MacroDocument) (c: CellRef) =
    tableOf book c.TableId |> Option.bind (fun t -> List.tryItem c.Row t.Rows) |> Option.bind (fun row -> List.tryItem c.Col row.Cells)

type private ShadingTarget =
    | ParagraphShading of (unit -> string list)
    | CellShading of (unit -> CellRef list)

let private shadingObject (book: MacroDocument) (target: ShadingTarget) : Value =
    let readColor () =
        match target with
        | ParagraphShading ids -> ids () |> List.tryHead |> Option.bind book.Paragraph |> Option.bind (fun p -> p.Metadata.BackgroundColor)
        | CellShading cells -> cells () |> List.tryHead |> Option.bind (cellOf book) |> Option.bind (fun c -> c.BackgroundColor)
    let setColor (color: string option) =
        match target with
        | ParagraphShading ids ->
            for id in ids () do
                book.UpdateParagraph(id, fun p -> { p with Metadata = { p.Metadata with BackgroundColor = color } |> withFlag ParagraphExplicit.BackgroundColor color.IsSome })
        | CellShading cells -> mapCells book (cells ()) (fun c -> { c with BackgroundColor = color; Shading = None })
    makeObject "Shading"
        (fun name ->
            match name with
            | "backgroundpatterncolor" -> Some (colorValueOf (readColor ()))
            | "backgroundpatterncolorindex" -> Some (VInteger (indexOfColor (readColor ())))
            | "foregroundpatterncolor" -> Some (VInteger wdColorAutomatic)
            | "texture" -> Some (VInteger 0)
            | _ -> None)
        (fun name value ->
            match name with
            | "backgroundpatterncolor" -> setColor (colorOfValue value); true
            | "backgroundpatterncolorindex" -> setColor (colorOfIndex (toInt value)); true
            | "foregroundpatterncolor" | "foregroundpatterncolorindex" -> true
            | "texture" -> true
            | _ -> false)
        noCall None

let private lineStyleNames =
    [ 0, "none"; 1, "single"; 2, "dotted"; 3, "dashSmallGap"; 4, "dashed"; 5, "dotDash"; 6, "dotDotDash"; 7, "double"
      8, "triple"; 9, "thinThickSmallGap"; 10, "thickThinSmallGap"; 18, "wave"; 19, "doubleWave" ]
let private lineStyleName (n: int) = lineStyleNames |> List.tryPick (fun (k, v) -> if k = n then Some v else None) |> Option.defaultValue "single"
let private lineStyleValue (name: string option) =
    match name with
    | None -> 0
    | Some n -> lineStyleNames |> List.tryPick (fun (k, v) -> if eq v n then Some k else None) |> Option.defaultValue 1

/// A border as (style, colour, width in pixels); None is no border.
type private Line = { LineStyle: string; LineColor: string option; LineWidth: float }
let private defaultLine = { LineStyle = "single"; LineColor = Some "#000000"; LineWidth = 4.0 / 6.0 }

type private BorderSide = SideTop | SideLeft | SideBottom | SideRight | SideInsideH | SideInsideV

let private sideOf (n: int) =
    match n with
    | -1 -> Some SideTop | -2 -> Some SideLeft | -3 -> Some SideBottom | -4 -> Some SideRight | -5 -> Some SideInsideH | -6 -> Some SideInsideV
    | _ -> None

type private BordersTarget =
    | ParagraphBorders of (unit -> string list)
    | TableBorders of string
    | CellBorders of (unit -> CellRef list)

let private tableLine (b: TableBorder option) =
    match b with
    | Some tb ->
        match tb.Style with
        | Some s when s <> "none" && s <> "nil" -> Some { LineStyle = s; LineColor = tb.Color; LineWidth = defaultArg tb.Width defaultLine.LineWidth }
        | _ -> None
    | None -> None
let private tableBorderOf (line: Line option) : TableBorder option =
    match line with
    | Some l -> Some ({ Style = Some l.LineStyle; Color = l.LineColor; Width = Some l.LineWidth } : TableBorder)
    | None -> Some ({ Style = Some "nil"; Color = None; Width = None } : TableBorder)
let private emptyTableBorders : TableBorders = { Top = None; Right = None; Bottom = None; Left = None; InsideH = None; InsideV = None }

let private getSide (side: BorderSide) (b: TableBorders) =
    match side with SideTop -> b.Top | SideLeft -> b.Left | SideBottom -> b.Bottom | SideRight -> b.Right | SideInsideH -> b.InsideH | SideInsideV -> b.InsideV
let private setSide (side: BorderSide) (v: TableBorder option) (b: TableBorders) =
    match side with
    | SideTop -> { b with Top = v } | SideLeft -> { b with Left = v } | SideBottom -> { b with Bottom = v }
    | SideRight -> { b with Right = v } | SideInsideH -> { b with InsideH = v } | SideInsideV -> { b with InsideV = v }

let private readLine (book: MacroDocument) (target: BordersTarget) (side: BorderSide) : Line option =
    match target with
    | TableBorders id -> tableOf book id |> Option.bind (fun t -> t.Borders) |> Option.bind (getSide side >> Some) |> Option.bind tableLine
    | CellBorders cells -> cells () |> List.tryHead |> Option.bind (cellOf book) |> Option.bind (fun c -> c.Borders) |> Option.bind (getSide side >> Some) |> Option.bind tableLine
    | ParagraphBorders ids ->
        ids () |> List.tryHead |> Option.bind book.Paragraph |> Option.bind (fun p -> p.Metadata.ParagraphBorders)
        |> Option.bind (fun b -> match side with SideTop -> b.Top | SideLeft -> b.Left | SideBottom -> b.Bottom | SideRight -> b.Right | _ -> None)
        |> Option.bind (fun pb -> match pb.Style with Some s when s <> "none" && s <> "nil" -> Some { LineStyle = s; LineColor = pb.Color; LineWidth = defaultArg pb.Width defaultLine.LineWidth } | _ -> None)

let private writeLine (book: MacroDocument) (target: BordersTarget) (side: BorderSide) (change: Line option -> Line option) =
    match target with
    | TableBorders id ->
        mapTable book id (fun t ->
            let borders = defaultArg t.Borders emptyTableBorders
            { t with Borders = Some (setSide side (tableBorderOf (change (tableLine (getSide side borders)))) borders) })
    | CellBorders cells ->
        mapCells book (cells ()) (fun c ->
            let borders = defaultArg c.Borders emptyTableBorders
            { c with Borders = Some (setSide side (tableBorderOf (change (tableLine (getSide side borders)))) borders) })
    | ParagraphBorders ids ->
        for id in ids () do
            book.UpdateParagraph(id, fun p ->
                let borders = p.Metadata.ParagraphBorders |> Option.defaultValue ({ Top = None; Right = None; Bottom = None; Left = None } : ParagraphBorders)
                let current =
                    match side with SideTop -> borders.Top | SideLeft -> borders.Left | SideBottom -> borders.Bottom | SideRight -> borders.Right | _ -> None
                let line = current |> Option.bind (fun pb -> match pb.Style with Some s when s <> "none" && s <> "nil" -> Some { LineStyle = s; LineColor = pb.Color; LineWidth = defaultArg pb.Width defaultLine.LineWidth } | _ -> None)
                let next = change line |> Option.map (fun l -> ({ Style = Some l.LineStyle; Color = l.LineColor; Width = Some l.LineWidth; Space = Some (pxOfPt 1.0) } : ParagraphBorder))
                let borders : ParagraphBorders =
                    match side with
                    | SideTop -> { borders with Top = next } | SideLeft -> { borders with Left = next }
                    | SideBottom -> { borders with Bottom = next } | SideRight -> { borders with Right = next } | _ -> borders
                { p with Metadata = { p.Metadata with ParagraphBorders = Some borders } })

let private borderObject (book: MacroDocument) (target: BordersTarget) (side: BorderSide) : Value =
    makeObject "Border"
        (fun name ->
            let line = readLine book target side
            match name with
            | "linestyle" -> Some (VInteger (match line with Some l -> lineStyleValue (Some l.LineStyle) | None -> 0))
            | "linewidth" -> Some (VInteger (match line with Some l -> int (Math.Round(l.LineWidth * 6.0)) | None -> 4))
            | "color" -> Some (colorValueOf (line |> Option.bind (fun l -> l.LineColor)))
            | "colorindex" -> Some (VInteger (indexOfColor (line |> Option.bind (fun l -> l.LineColor))))
            | "visible" -> Some (VBoolean line.IsSome)
            | _ -> None)
        (fun name value ->
            let update change = writeLine book target side change; true
            match name with
            | "linestyle" ->
                let n = toInt value
                update (fun line -> if n = 0 then None else Some { (defaultArg line defaultLine) with LineStyle = lineStyleName n })
            | "linewidth" -> let px = float (toInt value) / 6.0 in update (fun line -> Some { (defaultArg line defaultLine) with LineWidth = px })
            | "color" -> let c = colorOfValue value in update (fun line -> Some { (defaultArg line defaultLine) with LineColor = (match c with Some _ -> c | None -> Some "#000000") })
            | "colorindex" -> let c = colorOfIndex (toInt value) in update (fun line -> Some { (defaultArg line defaultLine) with LineColor = (match c with Some _ -> c | None -> Some "#000000") })
            | "visible" -> let on = toBool value in update (fun line -> if on then Some (defaultArg line defaultLine) else None)
            | _ -> false)
        noCall None

let private bordersObject (book: MacroDocument) (target: BordersTarget) : Value =
    let sides =
        match target with
        | ParagraphBorders _ -> [ SideTop; SideLeft; SideBottom; SideRight ]
        | _ -> [ SideTop; SideLeft; SideBottom; SideRight; SideInsideH; SideInsideV ]
    let outside = [ SideTop; SideLeft; SideBottom; SideRight ]
    let inside = sides |> List.filter (fun s -> s = SideInsideH || s = SideInsideV)
    let setAll (which: BorderSide list) change = for side in which do writeLine book target side change
    collectionObject "Borders"
        (fun () -> sides |> List.map (borderObject book target) |> List.toArray)
        (fun name ->
            match name with
            | "enable" -> Some (VInteger (if sides |> List.exists (fun s -> (readLine book target s).IsSome) then -1 else 0))
            | "insidelinestyle" -> Some (VInteger (match inside |> List.tryHead |> Option.bind (readLine book target) with Some l -> lineStyleValue (Some l.LineStyle) | None -> 0))
            | "outsidelinestyle" -> Some (VInteger (match readLine book target SideTop with Some l -> lineStyleValue (Some l.LineStyle) | None -> 0))
            | "shadow" -> Some (VBoolean false)
            | _ -> None)
        (fun name value ->
            match name with
            | "enable" ->
                let on = toBool value
                setAll sides (fun line -> if on then Some (defaultArg line defaultLine) else None); true
            | "insidelinestyle" | "outsidelinestyle" ->
                let n = toInt value
                setAll (if name = "insidelinestyle" then inside else outside) (fun line -> if n = 0 then None else Some { (defaultArg line defaultLine) with LineStyle = lineStyleName n }); true
            | "insidelinewidth" | "outsidelinewidth" ->
                let px = float (toInt value) / 6.0
                setAll (if name = "insidelinewidth" then inside else outside) (fun line -> Some { (defaultArg line defaultLine) with LineWidth = px }); true
            | "insidecolor" | "outsidecolor" | "insidecolorindex" | "outsidecolorindex" ->
                let c = if name.EndsWith "index" then colorOfIndex (toInt value) else colorOfValue value
                let c = match c with Some _ -> c | None -> Some "#000000"
                setAll (if name.StartsWith "inside" then inside else outside) (fun line -> Some { (defaultArg line defaultLine) with LineColor = c }); true
            | "shadow" | "distancefromtop" | "distancefrombottom" | "distancefromleft" | "distancefromright" -> true
            | _ -> false)
        (fun name args ->
            match name with
            | "" | "_default" | "item" ->
                match arg args 0 |> Option.map toInt |> Option.bind sideOf with
                | Some side -> Some (borderObject book target side)
                | None -> Some (borderObject book target SideTop)
            | _ -> None)

// ---- tables -------------------------------------------------------------------------------------------------

/// The story span of a table's cell (its paragraphs and its end-of-cell mark), of a row (to its
/// end-of-row mark), of the whole table.
let private cellSpan (book: MacroDocument) (c: CellRef) =
    let segs = book.Story.Segments |> Array.filter (fun s -> s.Table = c.TableId && s.Row = c.Row && s.Col = c.Col && s.Mark <> RowMark)
    if segs.Length = 0 then noSuchMember () else segs.[0].Start, segmentEnd segs.[segs.Length - 1]
let private rowSpan (book: MacroDocument) (tableId: string) (row: int) =
    let segs = book.Story.Segments |> Array.filter (fun s -> s.Table = tableId && s.Row = row)
    if segs.Length = 0 then noSuchMember () else segs.[0].Start, segmentEnd segs.[segs.Length - 1]
let private tableSpan (book: MacroDocument) (tableId: string) =
    match book.TableSpans() |> List.tryFind (fun (id, _, _) -> id = tableId) with
    | Some (_, s, e) -> s, e
    | None -> fail 5941 "The table is no longer in the document."

/// The Office theme's accent colours, for the built-in table styles a macro applies.
let private accents = [| "#4472C4"; "#ED7D31"; "#A5A5A5"; "#FFC000"; "#5B9BD5"; "#70AD47" |]
let private tint (hex: string) (amount: float) =
    match longOfHex hex with
    | Some n ->
        let mix c = int (Math.Round(float c + (255.0 - float c) * amount))
        sprintf "#%02X%02X%02X" (mix (n &&& 0xFF)) (mix ((n >>> 8) &&& 0xFF)) (mix ((n >>> 16) &&& 0xFF))
    | None -> hex

/// A built-in table style ("Table Grid", "Grid Table 4 - Accent 1", "List Table 3", "Medium Shading 1 -
/// Accent 2", "Light Grid") applied as direct formatting of the cells (the model keeps a table style's
/// formatting resolved into its cells); None for a style this does not know.
let private applyTableStyle (book: MacroDocument) (tableId: string) (name: string) : bool =
    let lower = name.ToLowerInvariant()
    let accent =
        let m = Text.RegularExpressions.Regex.Match(lower, "accent\\s*(\\d)")
        if m.Success then accents.[(int m.Groups.[1].Value - 1) % accents.Length] else "#000000"
    let line color = Some { Style = Some "single"; Color = Some color; Width = Some (4.0 / 6.0) }
    let all color = { Top = line color; Right = line color; Bottom = line color; Left = line color; InsideH = line color; InsideV = line color }
    let none = { Top = None; Right = None; Bottom = None; Left = None; InsideH = None; InsideV = None }
    let header = "header"
    let paint (headerFill: string option) (headerText: string option) (band: string option) (borders: TableBorders) =
        let dark = accent = "#000000"
        mapTable book tableId (fun t ->
            { t with
                Borders = Some borders
                Rows =
                    t.Rows |> List.mapi (fun r row ->
                        let fill =
                            if r = 0 then headerFill
                            elif r % 2 = 1 then band
                            else None
                        { row with IsHeader = (r = 0) || row.IsHeader; Cells = row.Cells |> List.map (fun c -> { c with BackgroundColor = fill; Shading = None }) }) })
        ignore dark
        // The header row's text: bold (white on a dark fill).
        let s, e = rowSpan book tableId 0
        book.Format(s, e, fun st -> { TB.withExplicitBold true st with Color = (match headerText with Some c -> Some c | None -> st.Color) })
        ignore header
    if lower = "table grid" || lower = "tablegrid" then paint None None None (all "#000000"); true
    elif lower.StartsWith "grid table 4" || lower.StartsWith "medium grid" then
        paint (Some accent) (Some "#FFFFFF") (Some (tint accent 0.8)) (all (tint accent 0.4)); true
    elif lower.StartsWith "grid table 5" then
        paint (Some accent) (Some "#FFFFFF") (Some (tint accent 0.6)) (all "#FFFFFF")
        mapTable book tableId (fun t -> { t with Rows = t.Rows |> List.mapi (fun r row -> if r > 0 && r % 2 = 0 then { row with Cells = row.Cells |> List.map (fun c -> { c with BackgroundColor = Some (tint accent 0.8) }) } else row) })
        true
    elif lower.StartsWith "grid table 1" || lower.StartsWith "light grid" then
        paint None None None (all (tint accent 0.6)); true
    elif lower.StartsWith "list table 3" || lower.StartsWith "list table 4" || lower.StartsWith "medium shading" then
        paint (Some accent) (Some "#FFFFFF") (if lower.StartsWith "list table 3" then None else Some (tint accent 0.8))
            { none with Top = line accent; Bottom = line accent; Left = line accent; Right = line accent; InsideH = (if lower.StartsWith "list table 4" then line (tint accent 0.4) else None) }
        true
    elif lower.StartsWith "plain table" || lower.StartsWith "light shading" || lower.StartsWith "list table" then
        paint None None (Some "#F2F2F2") { none with Top = line "#BFBFBF"; Bottom = line "#BFBFBF"; InsideH = line "#BFBFBF" }; true
    else false

// ---- Find -----------------------------------------------------------------------------------------------------

type private FindState() =
    member val Text = "" with get, set
    member val ReplacementText = "" with get, set
    member val Forward = true with get, set
    member val Wrap = 0 with get, set
    member val Format = false with get, set
    member val MatchCase = false with get, set
    member val MatchWholeWord = false with get, set
    member val MatchWildcards = false with get, set
    member val Found = false with get, set
    member val Font = Dictionary<string, Value>() with get
    member val Highlight = -1 with get, set
    member val ReplacementFont = Dictionary<string, Value>() with get
    member val ReplacementHighlight = -1 with get, set
    member val ReplacementStyle : string option = None with get, set
    member val ReplacementParagraph = Dictionary<string, Value>() with get

let private findStateOf (r: LiveRange) : FindState =
    match r.Find with
    | :? FindState as f -> f
    | _ ->
        let f = FindState()
        r.Find <- box f
        f

/// Find's text as a regular expression: its ^ codes (^p ^t ^l ^? ^# ^$ ^w ^^ ^nnn), or Word's wildcards.
let private findPattern (state: FindState) =
    let text = state.Text
    let b = Text.StringBuilder()
    if state.MatchWildcards then
        let mutable i = 0
        while i < text.Length do
            let c = text.[i]
            match c with
            | '?' -> b.Append "[^\\r\\a\\b]" |> ignore
            | '*' -> b.Append "[^\\r\\a\\b]*?" |> ignore
            | '@' -> b.Append "+" |> ignore
            | '<' -> b.Append "\\b(?=\\w)" |> ignore
            | '>' -> b.Append "\\b(?<=\\w)" |> ignore
            | '[' ->
                let close = text.IndexOf(']', i + 1)
                if close < 0 then b.Append "\\[" |> ignore
                else
                    let body = text.Substring(i + 1, close - i - 1)
                    let body = if body.StartsWith "!" then "^" + body.Substring 1 else body
                    b.Append('[').Append(body.Replace("\\", "\\\\")).Append(']') |> ignore
                    i <- close
            | '{' ->
                let close = text.IndexOf('}', i + 1)
                if close < 0 then b.Append "\\{" |> ignore
                else
                    b.Append('{').Append(text.Substring(i + 1, close - i - 1).Replace(";", ",")).Append('}') |> ignore
                    i <- close
            | '(' | ')' -> b.Append c |> ignore
            | '\\' when i + 1 < text.Length ->
                b.Append(Text.RegularExpressions.Regex.Escape(string text.[i + 1])) |> ignore
                i <- i + 1
            | '^' when i + 1 < text.Length && text.[i + 1] = 'p' -> b.Append "\\r" |> ignore; i <- i + 1
            | '^' when i + 1 < text.Length && text.[i + 1] = 't' -> b.Append "\\t" |> ignore; i <- i + 1
            | other -> b.Append(Text.RegularExpressions.Regex.Escape(string other)) |> ignore
            i <- i + 1
    else
        let mutable i = 0
        while i < text.Length do
            let c = text.[i]
            if c = '^' && i + 1 < text.Length then
                let code = text.[i + 1]
                i <- i + 1
                match code with
                | 'p' -> b.Append "\\r" |> ignore
                | 't' -> b.Append "\\t" |> ignore
                | 'l' | '1' when code = 'l' -> b.Append "\\n" |> ignore
                | '?' -> b.Append "[^\\r\\a\\b]" |> ignore
                | '#' -> b.Append "\\d" |> ignore
                | '$' -> b.Append "\\p{L}" |> ignore
                | 'w' -> b.Append "[ \\t\\u00A0]+" |> ignore
                | 's' -> b.Append "\\u00A0" |> ignore
                | '^' -> b.Append "\\^" |> ignore
                | d when Char.IsDigit d ->
                    let mutable j = i
                    while j < text.Length && Char.IsDigit text.[j] do j <- j + 1
                    let n = int (text.Substring(i, j - i))
                    let ch = match n with 13 -> '\r' | 11 -> '\n' | _ -> char n
                    b.Append(Text.RegularExpressions.Regex.Escape(string ch)) |> ignore
                    i <- j - 1
                | other -> b.Append(Text.RegularExpressions.Regex.Escape("^" + string other)) |> ignore
            else
                // (The story's text holds a line break as '\n' and a mark as '\r'.)
                let ch = if c = '\v' then '\n' else c
                b.Append(Text.RegularExpressions.Regex.Escape(string ch)) |> ignore
            i <- i + 1
    let pattern = b.ToString()
    let pattern = if state.MatchWholeWord && not state.MatchWildcards then "(?<![\\w])" + pattern + "(?![\\w])" else pattern
    let options = if state.MatchCase then Text.RegularExpressions.RegexOptions.None else Text.RegularExpressions.RegexOptions.IgnoreCase
    Text.RegularExpressions.Regex(pattern, options)

/// The replacement text for a match: ^& (the match), ^p ^t ^l ^^, and \1 .. \9 with wildcards.
let private replacementFor (state: FindState) (m: Text.RegularExpressions.Match) =
    let text = state.ReplacementText
    let b = Text.StringBuilder()
    let mutable i = 0
    while i < text.Length do
        let c = text.[i]
        if c = '^' && i + 1 < text.Length then
            match text.[i + 1] with
            | '&' -> b.Append(m.Value) |> ignore
            | 'p' -> b.Append '\r' |> ignore
            | 't' -> b.Append '\t' |> ignore
            | 'l' -> b.Append '\v' |> ignore
            | 'm' -> b.Append '\f' |> ignore
            | '^' -> b.Append '^' |> ignore
            | other -> b.Append('^').Append(other) |> ignore
            i <- i + 2
        elif c = '\\' && state.MatchWildcards && i + 1 < text.Length && Char.IsDigit text.[i + 1] then
            b.Append(m.Groups.[int text.[i + 1] - int '0'].Value) |> ignore
            i <- i + 2
        else
            b.Append c |> ignore
            i <- i + 1
    b.ToString()

let private hasFindFormatting (state: FindState) = state.Format && (state.Font.Count > 0 || state.Highlight >= 0)

let private matchesFormatting (book: MacroDocument) (state: FindState) (s: int) (e: int) =
    if not (hasFindFormatting state) then true
    else
        let styles = book.StylesIn(s, e)
        let fontOk =
            state.Font |> Seq.forall (fun kv ->
                let wanted = match fontChange kv.Key kv.Value with Some f -> fontRead kv.Key (f TB.defaultStyle) | None -> None
                styles |> List.forall (fun st -> fontRead kv.Key st = wanted))
        let highlightOk =
            if state.Highlight < 0 then true
            elif state.Highlight = 0 then styles |> List.forall (fun st -> st.Highlight.IsNone)
            else styles |> List.forall (fun st -> st.Highlight.IsSome)
        fontOk && highlightOk

/// The matches of Find in [from, until) of the story, in order.
let private findAll (book: MacroDocument) (state: FindState) (from: int) (until: int) : (int * int * Text.RegularExpressions.Match option) list =
    let st = book.Story
    let from = max 0 from
    let until = min st.Total until
    if until <= from then []
    elif state.Text = "" then
        if not (hasFindFormatting state) then []
        else
            // (Formatting alone: the stretches of text that have it.)
            let result = ResizeArray<int * int * Text.RegularExpressions.Match option>()
            let mutable start = -1
            for p in from .. until - 1 do
                let ok = st.Text.[p] <> '\a' && st.Text.[p] <> '\b' && matchesFormatting book state p (p + 1)
                if ok && start < 0 then start <- p
                if not ok && start >= 0 then
                    result.Add((start, p, None))
                    start <- -1
            if start >= 0 then result.Add((start, until, None))
            List.ofSeq result
    else
        let regex = findPattern state
        let slice = st.Text.Substring(from, until - from)
        [ for m in regex.Matches(slice) |> Seq.cast<Text.RegularExpressions.Match> do
            if m.Length > 0 then
                let s = from + m.Index
                let e = s + m.Length
                if matchesFormatting book state s e then yield (s, e, Some m) ]

let private applyReplacementFormatting (book: MacroDocument) (state: FindState) (s: int) (e: int) =
    if state.ReplacementFont.Count > 0 then book.Format(s, e, specChange state.ReplacementFont)
    if state.ReplacementHighlight >= 0 then
        let color =
            if state.ReplacementHighlight = 0 then None
            else
                let index = match book.Settings.TryGetValue "defaulthighlightcolorindex" with | true, v -> toInt v | _ -> 7
                highlightOfIndex index
        book.Format(s, e, fun st -> { st with Highlight = color })
    match state.ReplacementStyle with
    | Some name ->
        match book.FindStyle name with
        | Some style when style.StyleType = "character" -> book.Format(s, e, fun st -> { st with Style = Some style.Id })
        | Some style -> for id in book.ParagraphIdsIn(s, e) do book.ApplyParagraphStyle(id, Some style.Id, None)
        | None -> ()
    | None -> ()
    if state.ReplacementParagraph.Count > 0 then
        let changes = [ for kv in state.ReplacementParagraph -> (paragraphChange kv.Key kv.Value).Value ]
        book.MapParagraphs(s, e, fun m -> changes |> List.fold (fun acc f -> f acc) m)

let private replacesFormatting (state: FindState) =
    state.ReplacementFont.Count > 0 || state.ReplacementHighlight >= 0 || state.ReplacementStyle.IsSome || state.ReplacementParagraph.Count > 0

/// Replace one match: its text (unless the replacement is only formatting) and its formatting; the
/// range the replacement takes.
let private replaceMatch (book: MacroDocument) (state: FindState) (s: int, e: int, m: Text.RegularExpressions.Match option) =
    let formattingOnly = state.ReplacementText = "" && state.Format && replacesFormatting state
    let text =
        match m with
        | Some m -> replacementFor state m
        | None -> state.ReplacementText
    let s, e =
        if formattingOnly || (match m with Some found -> text = found.Value | None -> false) then s, e
        else book.Replace(s, e, text)
    if state.Format || replacesFormatting state then applyReplacementFormatting book state s e
    s, e

// ---- the objects ---------------------------------------------------------------------------------------------

let private unitSpans (book: MacroDocument) (unit: int) : (int * int) array =
    let st = book.Story
    match unit with
    | 2 -> book.Words
    | 3 -> book.Sentences
    | 4 | 5 | 7 ->
        st.Segments |> Array.filter (fun s -> s.Mark <> RowMark) |> Array.map (fun s -> s.Start, segmentEnd s)
    | 6 | 8 | 11 -> [| 0, st.Total |]
    | 12 ->
        st.Segments |> Array.filter (fun s -> s.Table <> "" && s.Mark <> RowMark) |> Array.groupBy (fun s -> s.Table, s.Row, s.Col)
        |> Array.map (fun (_, segs) -> segs.[0].Start, segmentEnd segs.[segs.Length - 1])
    | 10 ->
        st.Segments |> Array.filter (fun s -> s.Table <> "") |> Array.groupBy (fun s -> s.Table, s.Row)
        |> Array.map (fun (_, segs) -> segs.[0].Start, segmentEnd segs.[segs.Length - 1])
    | 15 -> book.TableSpans() |> List.map (fun (_, s, e) -> s, e) |> List.toArray
    | _ -> [| for p in 0 .. st.Total - 1 -> p, p + 1 |]

let private spanAt (book: MacroDocument) (unit: int) (pos: int) =
    if unit = 1 then pos, pos + 1
    else
        let spans = unitSpans book unit
        match spans |> Array.tryFind (fun (s, e) -> s <= pos && pos < e) with
        | Some span -> span
        | None -> pos, pos

/// A position moved `count` unit boundaries (forward for a positive count).
let private moveBy (book: MacroDocument) (unit: int) (pos: int) (count: int) =
    let total = book.Story.Total
    if unit = 1 then max 0 (min total (pos + count))
    else
        let bounds =
            unitSpans book unit |> Array.collect (fun (s, e) -> [| s; e |]) |> Array.distinct |> Array.sort
        let mutable p = pos
        if count > 0 then
            for _ in 1 .. count do
                match bounds |> Array.tryFind (fun b -> b > p) with
                | Some b -> p <- b
                | None -> p <- total
        else
            for _ in 1 .. -count do
                match bounds |> Array.tryFindBack (fun b -> b < p) with
                | Some b -> p <- b
                | None -> p <- 0
        p

let private statistics (book: MacroDocument) (s: int) (e: int) (kind: int) =
    let text = book.Story.Text.Substring(s, max 0 (e - s))
    let wordCount = text.Split([| ' '; '\r'; '\a'; '\b'; '\t'; '\n'; '\f' |], StringSplitOptions.RemoveEmptyEntries) |> Array.filter (fun w -> w |> Seq.exists Char.IsLetterOrDigit) |> Array.length
    match kind with
    | 0 -> wordCount
    | 1 | 4 -> book.SegmentsIn(s, e) |> List.filter (fun seg -> seg.Mark <> RowMark && (seg.Length > 0 || seg.Mark = ParagraphMark)) |> List.length
    | 2 -> 1 + (book.SegmentsIn(s, e) |> List.filter (fun seg -> seg.Para <> "" && (book.Paragraph seg.Para |> Option.exists (fun p -> p.Metadata.PageBreakBefore))) |> List.length)
    | 3 -> text |> Seq.filter (fun c -> not (Char.IsWhiteSpace c) && c <> '\a' && c <> '\b') |> Seq.length
    | 5 -> text |> Seq.filter (fun c -> c <> '\r' && c <> '\a' && c <> '\b') |> Seq.length
    | _ -> 0

let rec private rangeObject (book: MacroDocument) (r: LiveRange) : Value =
    makeObject "Range" (rangeGet book r) (rangeSet book r) (rangeCall book r) None

and private rangeOf (book: MacroDocument) (s: int) (e: int) = rangeObject book (book.Track(s, e))

and private rangeGet (book: MacroDocument) (r: LiveRange) (name: string) : Value option =
    match name with
    | "" | "_default" | "text" -> Some (VString (book.TextOf(r.Start, r.End)))
    | "formattedtext" | "duplicate" -> Some (rangeOf book r.Start r.End)
    | "start" -> Some (VInteger r.Start)
    | "end" -> Some (VInteger r.End)
    | "storylength" -> Some (VInteger book.Story.Total)
    | "storytype" -> Some (VInteger 1)
    | "font" -> Some (fontObject book (RangeFont r))
    | "paragraphformat" -> Some (paragraphFormatObject book (fun () -> book.ParagraphIdsIn(r.Start, r.End)))
    | "paragraphs" -> Some (paragraphsObject book (fun () -> r.Start, r.End))
    | "words" -> Some (spansObject book "Words" (fun () -> r.Start, r.End) 2)
    | "characters" -> Some (spansObject book "Characters" (fun () -> r.Start, r.End) 1)
    | "sentences" -> Some (spansObject book "Sentences" (fun () -> r.Start, r.End) 3)
    | "tables" -> Some (tablesObject book (fun () -> r.Start, r.End))
    | "cells" -> Some (cellsInRange book r)
    | "rows" -> Some (rowsInRange book r)
    | "style" -> Some (styleOfRange book r)
    | "characterstyle" -> Some (VString "Default Paragraph Font")
    | "paragraphstyle" -> Some (styleOfRange book r)
    | "highlightcolorindex" ->
        let values = book.StylesIn(r.Start, r.End) |> List.map (fun s -> indexOfHighlight s.Highlight) |> List.distinct
        Some (VInteger (match values with [ v ] -> v | _ -> wdUndefined))
    | "bold" | "italic" | "underline" -> (match fontObject book (RangeFont r) with VHostObject h -> h.GetMember name | _ -> None)
    | "case" -> Some (VInteger wdUndefined)
    | "find" -> Some (findObject book r false)
    | "document" -> Some (documentObject book)
    | "shading" -> Some (shadingObject book (ParagraphShading (fun () -> book.ParagraphIdsIn(r.Start, r.End))))
    | "borders" -> Some (bordersObject book (ParagraphBorders (fun () -> book.ParagraphIdsIn(r.Start, r.End))))
    | "listformat" -> Some (listFormatObject book r)
    | "bookmarks" -> Some (bookmarksObject book (Some r))
    | "information" | "computestatistics" -> None
    | "isendofrowmark" -> Some (VBoolean (book.SegmentsIn(r.Start, r.Start) |> List.exists (fun s -> s.Mark = RowMark)))
    | "languageid" -> Some (VInteger 1033)
    | "revisions" | "fields" | "inlineshapes" | "shapes" -> Some (collectionObject "Collection" (fun () -> [||]) (fun _ -> None) noSet noCall)
    | _ -> None

and private rangeSet (book: MacroDocument) (r: LiveRange) (name: string) (value: Value) : bool =
    match name with
    | "" | "_default" | "text" | "formattedtext" ->
        let s, e = book.Replace(r.Start, r.End, toText value, r)
        r.Start <- s; r.End <- e; r.Matched <- false; true
    | "start" ->
        let v = max 0 (min book.Story.Total (toInt value))
        r.Matched <- false
        r.Start <- v
        if r.End < v then r.End <- v
        true
    | "end" ->
        let v = max 0 (min book.Story.Total (toInt value))
        r.Matched <- false
        r.End <- v
        if r.Start > v then r.Start <- v
        true
    | "style" -> setRangeStyle book r value; true
    | "paragraphstyle" -> setRangeStyle book r value; true
    | "characterstyle" -> true
    | "highlightcolorindex" ->
        let color = highlightOfIndex (toInt value)
        book.Format(r.Start, r.End, fun s -> { s with Highlight = color }); true
    | "bold" | "italic" | "underline" -> (match fontObject book (RangeFont r) with VHostObject h -> h.SetMember name value | _ -> false)
    | "case" ->
        let kind = toInt value
        let st = book.Story.Text
        book.MapText(r.Start, r.End, fun text i ->
            let c = text.[i]
            match kind with
            | 0 -> Char.ToLowerInvariant c
            | 1 -> Char.ToUpperInvariant c
            | 2 -> if i = 0 || not (Char.IsLetterOrDigit text.[i - 1] || text.[i - 1] = '\'') then Char.ToUpperInvariant c else Char.ToLowerInvariant c
            | 4 ->
                let rec start j = if j < 0 then true elif Char.IsWhiteSpace text.[j] then start (j - 1) else text.[j] = '.' || text.[j] = '!' || text.[j] = '?'
                if Char.IsLetter c && start (i - 1) then Char.ToUpperInvariant c else Char.ToLowerInvariant c
            | 5 -> if Char.IsUpper c then Char.ToLowerInvariant c else Char.ToUpperInvariant c
            | _ -> c)
        ignore st
        true
    | "languageid" | "noproofing" -> true
    | _ -> false

and private rangeCall (book: MacroDocument) (r: LiveRange) (name: string) (args: Value array) : Value option =
    let collapsedAt p = r.Start <- p; r.End <- p
    match name with
    | "insertafter" ->
        let text = toText (defaultArg (arg args 0) (VString ""))
        let total = book.Story.Total
        let at = min r.End (total - 1)
        let s0 = min r.Start at
        let _, e = book.Replace(at, at, text, r)
        // (A range that took in the last paragraph mark still does.)
        r.Start <- s0; r.End <- (if r.End >= total then book.Story.Total else e); Some VEmpty
    | "insertbefore" ->
        let text = toText (defaultArg (arg args 0) (VString ""))
        let e0 = r.End
        let s, n = book.Replace(r.Start, r.Start, text, r)
        r.Start <- s; r.End <- e0 + (n - s); Some VEmpty
    | "insertparagraphafter" -> rangeCall book r "insertafter" [| VString "\r" |]
    | "insertparagraphbefore" -> rangeCall book r "insertbefore" [| VString "\r" |]
    | "insertparagraph" -> rangeSet book r "text" (VString "\r") |> ignore; Some VEmpty
    | "insertbreak" ->
        let ch = match intArg args 0 7 with 6 -> "\v" | _ -> "\f"
        let s, e = book.Replace(r.Start, r.End, ch, r)
        collapsedAt e; ignore s; Some VEmpty
    | "delete" ->
        if r.End > r.Start then
            let s, _ = book.Replace(r.Start, r.End, "", r)
            collapsedAt s
        else
            let count = intArg args 1 1
            let s, _ = book.Replace(r.Start, moveBy book (intArg args 0 1) r.Start count, "", r)
            collapsedAt s
        Some (VInteger 1)
    | "collapse" ->
        let direction = intArg args 0 1
        // (The end of a range with the last paragraph mark is before that mark: nothing is after it.)
        if direction = 0 then collapsedAt (min r.End (book.Story.Total - 1)) else collapsedAt r.Start
        Some VEmpty
    | "select" ->
        book.Selection.Start <- r.Start; book.Selection.End <- r.End; book.Selection.ActiveAtStart <- false; Some VEmpty
    | "setrange" ->
        let s = intArg args 0 r.Start
        let e = intArg args 1 r.End
        r.Matched <- false
        r.Start <- max 0 (min s e); r.End <- min book.Story.Total (max s e); Some VEmpty
    | "move" ->
        let unit = intArg args 0 1
        let count = intArg args 1 1
        let p =
            if r.End > r.Start then
                let collapsed = if count > 0 then r.End else r.Start
                moveBy book unit collapsed (if count > 0 then count - 1 else count + 1)
            else moveBy book unit r.Start count
        collapsedAt p
        Some (VInteger count)
    | "movestart" ->
        let p = moveBy book (intArg args 0 1) r.Start (intArg args 1 1)
        r.Start <- p
        if r.End < p then r.End <- p
        Some (VInteger 1)
    | "moveend" ->
        let p = moveBy book (intArg args 0 1) r.End (intArg args 1 1)
        r.End <- p
        if r.Start > p then r.Start <- p
        Some (VInteger 1)
    | "expand" ->
        let unit = intArg args 0 2
        let s, _ = spanAt book unit r.Start
        let _, e = spanAt book unit (max r.Start (r.End - 1))
        r.Start <- s; r.End <- max s e; Some (VInteger 1)
    | "startof" ->
        let s, _ = spanAt book (intArg args 0 2) r.Start
        if intArg args 1 0 = 1 then r.Start <- s else collapsedAt s
        Some (VInteger 1)
    | "endof" ->
        let _, e = spanAt book (intArg args 0 2) (max r.Start (r.End - 1))
        if intArg args 1 0 = 1 then r.End <- e else collapsedAt e
        Some (VInteger 1)
    | "next" | "previous" ->
        let unit = intArg args 0 1
        let count = intArg args 1 1
        let p = if name = "next" then moveBy book unit r.End (count - 1) else moveBy book unit r.Start (-count)
        if p >= book.Story.Total then Some VNothing
        else
            let s, e = spanAt book unit p
            Some (rangeOf book s e)
    | "information" ->
        Some (information book r.Start r.End (intArg args 0 0))
    | "computestatistics" -> Some (VInteger (statistics book r.Start r.End (intArg args 0 0)))
    | "inrange" | "isequal" ->
        match arg args 0 |> Option.bind rangeBounds with
        | Some (s, e) -> Some (VBoolean (if name = "inrange" then s <= r.Start && r.End <= e else s = r.Start && r.End = e))
        | None -> fail 13 "Type mismatch"
    | "converttotable" -> Some (convertToTable book r args)
    | "words" | "characters" | "sentences" | "paragraphs" | "tables" | "cells" | "rows" when args.Length > 0 -> None
    | _ -> None

and private information (book: MacroDocument) (s: int) (e: int) (kind: int) : Value =
    let segs = book.SegmentsIn(s, max s e)
    let inTable = segs |> List.exists (fun seg -> seg.Table <> "")
    match kind with
    | 12 -> VBoolean inTable
    | 1 | 3 ->
        let breaks =
            book.Story.Segments
            |> Array.filter (fun seg -> seg.Start <= e && seg.Para <> "" && (book.Paragraph seg.Para |> Option.exists (fun p -> p.Metadata.PageBreakBefore)))
            |> Array.length
        VInteger (1 + breaks)
    | 4 -> VInteger (statistics book 0 book.Story.Total 2)
    | 2 -> VInteger 1
    | 13 | 14 -> (match segs |> List.tryFind (fun seg -> seg.Table <> "") with Some seg -> VInteger (seg.Row + 1) | None -> VInteger -1)
    | 16 | 17 -> (match segs |> List.tryFind (fun seg -> seg.Table <> "") with Some seg -> VInteger (seg.Col + 1) | None -> VInteger -1)
    | 15 | 18 ->
        match segs |> List.tryFind (fun seg -> seg.Table <> "") |> Option.bind (fun seg -> tableOf book seg.Table) with
        | Some t -> VInteger (if kind = 15 then t.Rows.Length else t.Rows |> List.map (fun row -> row.Cells.Length) |> List.max)
        | None -> VInteger -1
    | 31 -> VBoolean (segs |> List.exists (fun seg -> seg.Mark = RowMark))
    | _ -> VEmpty

and private styleOfRange (book: MacroDocument) (r: LiveRange) : Value =
    let ids = book.ParagraphIdsIn(r.Start, r.End) |> List.map (fun id -> book.Paragraph id |> Option.bind (fun p -> p.Metadata.Style)) |> List.distinct
    match ids with
    | [ Some id ] ->
        match book.Document.Styles |> List.tryFind (fun s -> s.Id = id) with
        | Some s -> styleObject book s.Id
        | None -> VString id
    | [ None ] | [] -> VString "Normal"
    | _ -> VNothing

and private setRangeStyle (book: MacroDocument) (r: LiveRange) (value: Value) =
    let name =
        match scalarOf value with
        | VInteger n -> (match builtInStyleName n with Some s -> s | None -> fail 5834 "Item with specified name does not exist.")
        | VDouble d -> (match builtInStyleName (int d) with Some s -> s | None -> fail 5834 "Item with specified name does not exist.")
        | other -> toText other
    if eq name "Normal" && not (book.Document.Styles |> List.exists (fun s -> s.Id = "Normal")) then
        for id in book.ParagraphIdsIn(r.Start, r.End) do book.ApplyParagraphStyle(id, None, None)
    else
        match book.FindStyle name with
        | Some style when style.StyleType = "character" ->
            book.Format(r.Start, r.End, fun s -> { s with Style = Some style.Id })
        | Some style ->
            let level = headingLevelOf style
            for id in book.ParagraphIdsIn(r.Start, r.End) do book.ApplyParagraphStyle(id, Some style.Id, level)
        | None ->
            if eq name "List Bullet" || eq name "List Number" then
                for id in book.ParagraphIdsIn(r.Start, r.End) do
                    book.UpdateParagraph(id, fun p -> { p with Metadata = { p.Metadata with ListStyle = (if eq name "List Bullet" then Bullet else Numbered); ListLevel = max 1 p.Metadata.ListLevel } })
            else fail 5834 (sprintf "Item with specified name does not exist (style \"%s\")." name)

and private paragraphFormatObject (book: MacroDocument) (ids: unit -> string list) : Value =
    let named (id: string) =
        book.Paragraph id |> Option.bind (fun p -> p.Metadata.Style) |> Option.bind (fun s -> book.Document.Styles |> List.tryFind (fun x -> x.Id = s))
    let styleOf (id: string) : StyleFallback =
        let doc = book.Document
        let resolved =
            match doc.Blocks |> List.tryFind (function Paragraph p | Heading (_, p) -> p.Id = id | _ -> false) with
            | Some block -> Some (WordStyles.forBlockInDocument doc block)
            | None -> book.Paragraph id |> Option.map (WordStyles.forTableCellParagraphInDocument doc)
        { Fallback = resolved; Level = named id |> Option.bind headingLevelOf }
    makeObject "ParagraphFormat"
        (fun name ->
            match name with
            | "shading" -> Some (shadingObject book (ParagraphShading ids))
            | "borders" -> Some (bordersObject book (ParagraphBorders ids))
            | "style" ->
                match ids () with
                | id :: _ -> Some (VString (match named id with Some s -> styleDisplayName s | None -> "Normal"))
                | [] -> Some VEmpty
            | _ ->
                let values =
                    ids () |> List.choose (fun id -> book.Paragraph id |> Option.bind (fun p -> paragraphRead (styleOf id) name p.Metadata))
                match values |> List.distinct with
                | [ v ] -> Some v
                | [] -> paragraphRead noFallback name defaultMetadata
                | _ -> Some (VInteger wdUndefined))
        (fun name value ->
            match name with
            | "style" ->
                let all = ids ()
                for id in all do
                    let pos = book.PositionOf(id, 0)
                    setRangeStyle book (LiveRange(pos, pos)) value
                true
            | _ ->
                match paragraphChange name value with
                | Some change ->
                    for id in ids () do
                        book.UpdateParagraph(id, fun p -> let m = change p.Metadata in if m = p.Metadata then p else { p with Metadata = m })
                    true
                | None -> false)
        noCall None

and private listFormatObject (book: MacroDocument) (r: LiveRange) : Value =
    let apply (style: ListStyle option) =
        for id in book.ParagraphIdsIn(r.Start, r.End) do
            book.UpdateParagraph(id, fun p ->
                let m = p.Metadata
                match style with
                | Some s -> { p with Metadata = { m with ListStyle = (if m.ListStyle = s then NoList else s); ListLevel = (if m.ListStyle = s then 0 else max 1 m.ListLevel) } }
                | None -> { p with Metadata = { m with ListStyle = NoList; ListLevel = 0 } })
    makeObject "ListFormat"
        (fun name ->
            match name with
            | "listtype" ->
                match book.ParagraphIdsIn(r.Start, r.End) |> List.tryHead |> Option.bind book.Paragraph with
                | Some p -> Some (VInteger (match p.Metadata.ListStyle with Bullet -> 2 | Numbered -> 3 | NoList -> 0))
                | None -> Some (VInteger 0)
            | "listlevelnumber" ->
                match book.ParagraphIdsIn(r.Start, r.End) |> List.tryHead |> Option.bind book.Paragraph with
                | Some p -> Some (VInteger (max 1 p.Metadata.ListLevel))
                | None -> Some (VInteger 1)
            | _ -> None)
        (fun name value ->
            match name with
            | "listlevelnumber" ->
                let level = max 1 (min 9 (toInt value))
                for id in book.ParagraphIdsIn(r.Start, r.End) do
                    book.UpdateParagraph(id, fun p -> if p.Metadata.ListStyle = NoList then p else { p with Metadata = { p.Metadata with ListLevel = level } })
                true
            | _ -> false)
        (fun name _ ->
            match name with
            | "applybulletdefault" -> apply (Some Bullet); Some VEmpty
            | "applynumberdefault" -> apply (Some Numbered); Some VEmpty
            | "removenumbers" -> apply None; Some VEmpty
            | "listindent" | "listoutdent" ->
                for id in book.ParagraphIdsIn(r.Start, r.End) do
                    book.UpdateParagraph(id, fun p ->
                        if p.Metadata.ListStyle = NoList then p
                        else { p with Metadata = { p.Metadata with ListLevel = max 1 (min 9 (p.Metadata.ListLevel + (if name = "listindent" then 1 else -1))) } })
                Some VEmpty
            | _ -> None)
        None

and private spansObject (book: MacroDocument) (typeName: string) (bounds: unit -> int * int) (unit: int) : Value =
    let items () =
        let s, e = bounds ()
        let e = max e (if e = s then s + 1 else e)
        let spans =
            if unit = 1 then [| for p in s .. min (book.Story.Total - 1) (e - 1) -> p, p + 1 |]
            else unitSpans book unit |> Array.filter (fun (a, b) -> a < e && b > s) |> Array.map (fun (a, b) -> max a s, min b e)
        spans |> Array.map (fun (a, b) -> rangeOf book a b)
    collectionObject typeName items (fun _ -> None) noSet noCall

and private paragraphObject (book: MacroDocument) (r: LiveRange) : Value =
    let ids () =
        match book.ParagraphIdsIn(r.Start, r.End) with
        | [] ->
            // (An end-of-row mark: the format of its row's last paragraph.)
            let st = book.Story
            let i = locate st r.Start
            if i > 0 && st.Segments.[i - 1].Para <> "" then [ st.Segments.[i - 1].Para ] else []
        | found -> found
    let format = paragraphFormatObject book ids
    let formatHost = match format with VHostObject h -> h | _ -> failwith "format"
    makeObject "Paragraph"
        (fun name ->
            match name with
            | "range" -> Some (rangeObject book r)
            | "format" -> Some format
            | "style" -> Some (styleOfRange book r)
            | "next" | "previous" -> None
            | "outlinelevel" | "alignment" | "spacebefore" | "spaceafter" | "leftindent" | "rightindent" | "firstlineindent"
            | "linespacing" | "linespacingrule" | "keepwithnext" | "keeptogether" | "pagebreakbefore" | "widowcontrol"
            | "shading" | "borders" -> formatHost.GetMember name
            | _ -> None)
        (fun name value ->
            match name with
            | "style" -> setRangeStyle book r value; true
            | _ -> (try formatHost.SetMember name value with _ -> false))
        (fun name args ->
            match name with
            | "next" | "previous" ->
                let count = intArg args 0 1
                let st = book.Story
                let paras = st.Segments |> Array.filter (fun s -> s.Mark <> RowMark)
                let here = paras |> Array.tryFindIndex (fun s -> s.Start <= r.Start && r.Start < segmentEnd s) |> Option.defaultValue 0
                let target = if name = "next" then here + count else here - count
                if target < 0 || target >= paras.Length then Some VNothing
                else Some (paragraphObject book (book.Track(paras.[target].Start, segmentEnd paras.[target])))
            | "indent" | "outdent" ->
                let step = if name = "indent" then 36.0 else -36.0
                for id in ids () do
                    book.UpdateParagraph(id, fun p -> { p with Metadata = { p.Metadata with IndentLeft = Some (max 0.0 (defaultArg p.Metadata.IndentLeft 0.0 + pxOfPt step)) } |> withFlag ParagraphExplicit.IndentLeft true })
                Some VEmpty
            | _ -> rangeCall book r name args)
        None

and private paragraphsObject (book: MacroDocument) (bounds: unit -> int * int) : Value =
    let items () =
        let s, e = bounds ()
        // (An end-of-row mark is a paragraph to Word too.)
        book.SegmentsIn(s, e)
        |> List.map (fun seg -> paragraphObject book (book.Track(seg.Start, segmentEnd seg))) |> List.toArray
    let format = paragraphFormatObject book (fun () -> let s, e = bounds () in book.ParagraphIdsIn(s, e))
    let formatHost = match format with VHostObject h -> h | _ -> failwith "format"
    collectionObject "Paragraphs" items
        (fun name -> match name with "format" -> Some format | _ -> None)
        (fun name value -> try formatHost.SetMember name value with _ -> false)
        (fun name args ->
            match name with
            | "add" ->
                let at =
                    match arg args 0 with
                    | Some v -> (match rangeBounds v with Some (s, _) -> s | None -> fail 13 "Type mismatch")
                    | None -> book.Story.Total - 1
                let s, e = book.Replace(at, at, "\r")
                let st = book.Story
                let seg = st.Segments.[locate st e]
                Some (paragraphObject book (book.Track(seg.Start, segmentEnd seg)))
            | "indent" | "outdent" | "space1" | "space15" | "space2" ->
                let s, e = bounds ()
                let change =
                    match name with
                    | "space1" -> (paragraphChange "linespacingrule" (VInteger 0)).Value
                    | "space15" -> (paragraphChange "linespacingrule" (VInteger 1)).Value
                    | "space2" -> (paragraphChange "linespacingrule" (VInteger 2)).Value
                    | step -> (fun m -> { m with IndentLeft = Some (max 0.0 (defaultArg m.IndentLeft 0.0 + pxOfPt (if step = "indent" then 36.0 else -36.0))) } |> withFlag ParagraphExplicit.IndentLeft true)
                book.MapParagraphs(s, e, change); Some VEmpty
            | _ -> None)

and private rangeBounds (value: Value) : (int * int) option =
    match value with
    | VHostObject h ->
        match h.GetMember "Start", h.GetMember "End" with
        | Some s, Some e -> Some (toInt s, toInt e)
        | _ -> None
    | _ -> None

// ---- tables ----------------------------------------------------------------------------------------------------

and private cellObject (book: MacroDocument) (c: CellRef) : Value =
    makeObject "Cell"
        (fun name ->
            match name with
            | "range" -> let s, e = cellSpan book c in Some (rangeOf book s e)
            | "shading" -> Some (shadingObject book (CellShading (fun () -> [ c ])))
            | "borders" -> Some (bordersObject book (CellBorders (fun () -> [ c ])))
            | "rowindex" -> Some (VInteger (c.Row + 1))
            | "columnindex" -> Some (VInteger (c.Col + 1))
            | "verticalalignment" ->
                Some (VInteger (match cellOf book c |> Option.bind (fun x -> x.VerticalAlignment) with Some "center" -> 1 | Some "bottom" -> 3 | _ -> 0))
            | "width" -> Some (VDouble (cellOf book c |> Option.bind (fun x -> x.Width) |> Option.map ptOfPx |> Option.defaultValue 0.0))
            | "next" | "previous" -> None
            | _ -> None)
        (fun name value ->
            match name with
            | "verticalalignment" ->
                let v = match toInt value with 1 -> Some "center" | 3 -> Some "bottom" | _ -> Some "top"
                mapCells book [ c ] (fun x -> { x with VerticalAlignment = v }); true
            | "width" ->
                let px = pxOfPt (toFloat value)
                mapCells book [ c ] (fun x -> { x with Width = Some px; PreferredWidth = Some { Kind = DxaWidth; Value = Some (px * 15.0) } }); true
            | _ -> false)
        (fun name args ->
            match name with
            | "select" -> let s, e = cellSpan book c in book.Selection.Start <- s; book.Selection.End <- e; Some VEmpty
            | "next" | "previous" ->
                match tableOf book c.TableId with
                | Some t ->
                    let cells = [ for r, row in List.indexed t.Rows do for col in 0 .. row.Cells.Length - 1 -> { c with Row = r; Col = col } ]
                    let i = cells |> List.findIndex (fun x -> x.Row = c.Row && x.Col = c.Col)
                    let j = if name = "next" then i + 1 else i - 1
                    if j < 0 || j >= cells.Length then Some VNothing else Some (cellObject book cells.[j])
                | None -> Some VNothing
            | _ -> ignore args; None)
        None

and private tableObject (book: MacroDocument) (tableId: string) : Value =
    let table () = match tableOf book tableId with Some t -> t | None -> fail 5941 "The table is no longer in the document."
    let allCells () = [ for r, row in List.indexed (table ()).Rows do for col in 0 .. row.Cells.Length - 1 -> { TableId = tableId; Row = r; Col = col } ]
    makeObject "Table"
        (fun name ->
            match name with
            | "range" -> let s, e = tableSpan book tableId in Some (rangeOf book s e)
            | "rows" -> Some (rowsObject book tableId)
            | "columns" -> Some (columnsObject book tableId)
            | "borders" -> Some (bordersObject book (TableBorders tableId))
            | "shading" -> Some (shadingObject book (CellShading allCells))
            | "style" -> Some (VString (match book.Settings.TryGetValue ("tablestyle:" + tableId) with | true, v -> toText v | _ -> "Table Grid"))
            | "nestinglevel" -> Some (VInteger 1)
            | "uniform" -> Some (VBoolean true)
            | "title" | "descr" -> Some (VString "")
            | "preferredwidth" -> Some (VDouble 0.0)
            | "preferredwidthtype" -> Some (VInteger 1)
            | "allowautofit" -> Some (VBoolean true)
            | "applystyleheadingrows" | "applystylefirstcolumn" | "applystylelastrow" | "applystylelastcolumn" | "applystylerowbands" -> Some (VBoolean true)
            | _ -> None)
        (fun name value ->
            match name with
            | "style" ->
                let text = match scalarOf value with VInteger n -> string n | v -> toText v
                book.Settings.["tablestyle:" + tableId] <- VString text
                if not (applyTableStyle book tableId text) then
                    // (A style of the document: the reference is kept for the file.)
                    match book.Document.Styles |> List.tryFind (fun s -> s.StyleType = "table" && (eq s.Id text || s.Name |> Option.exists (eq text))) with
                    | Some s -> mapTable book tableId (fun t -> { t with StyleId = Some s.Id })
                    | None -> ()
                true
            | "preferredwidthtype" -> book.Settings.["tablewidthtype:" + tableId] <- value; true
            | "preferredwidth" ->
                let kind = match book.Settings.TryGetValue ("tablewidthtype:" + tableId) with | true, v -> toInt v | _ -> 3
                let width =
                    if kind = 2 then { Kind = PercentWidth; Value = Some (toFloat value * 50.0) }
                    else { Kind = DxaWidth; Value = Some (toFloat value * 20.0) }
                mapTable book tableId (fun t -> { t with PreferredWidth = Some width }); true
            | "allowautofit" | "title" | "descr" | "spacing" | "toppadding" | "bottompadding" | "leftpadding" | "rightpadding" -> true
            | "applystyleheadingrows" | "applystylefirstcolumn" | "applystylelastrow" | "applystylelastcolumn" | "applystylerowbands" | "applystylecolumnbands" -> true
            | _ -> false)
        (fun name args ->
            match name with
            | "cell" ->
                let r = intArg args 0 1 - 1
                let c = intArg args 1 1 - 1
                let t = table ()
                if r < 0 || r >= t.Rows.Length || c < 0 || c >= t.Rows.[r].Cells.Length then fail 5941 "The requested member of the collection does not exist."
                Some (cellObject book { TableId = tableId; Row = r; Col = c })
            | "autofitbehavior" ->
                if intArg args 0 0 = 2 then mapTable book tableId (fun t -> { t with PreferredWidth = Some { Kind = PercentWidth; Value = Some 5000.0 } })
                Some VEmpty
            | "autoformat" -> Some VEmpty
            | "delete" ->
                let s, e = tableSpan book tableId
                book.SetStructure { book.Document with Blocks = book.Document.Blocks |> List.filter (function Table t -> t.Id <> tableId | _ -> true) }
                ignore (s, e)
                Some VEmpty
            | "select" -> let s, e = tableSpan book tableId in book.Selection.Start <- s; book.Selection.End <- e; Some VEmpty
            | "sort" ->
                // (By the first column, or FieldNumber; the header row stays first when ExcludeHeader.)
                let exclude = arg args 0 |> Option.map toBool |> Option.defaultValue false
                let field = intArg args 1 1 - 1
                let descending = intArg args 3 0 = 1
                book.SetStructure (
                    let doc = book.Document
                    { doc with
                        Blocks =
                            doc.Blocks |> List.map (function
                                | Table t when t.Id = tableId ->
                                    let head, body = if exclude then List.truncate 1 t.Rows, List.skip (min 1 t.Rows.Length) t.Rows else [], t.Rows
                                    let key (row: TableRow) =
                                        row.Cells |> List.tryItem field |> Option.map (fun c -> c.Content |> List.map (fun p -> TB.toString p.Buffer) |> String.concat "\r") |> Option.defaultValue ""
                                    let numeric (s: string) = match Double.TryParse(s, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture) with | true, d -> Some d | _ -> None
                                    let sorted =
                                        if body |> List.forall (fun row -> (numeric (key row)).IsSome) then body |> List.sortBy (fun row -> (numeric (key row)).Value)
                                        else body |> List.sortWith (fun a b -> String.Compare(key a, key b, StringComparison.OrdinalIgnoreCase))
                                    Table { t with Rows = head @ (if descending then List.rev sorted else sorted) }
                                | other -> other) })
                Some VEmpty
            | _ -> None)
        None

and private rowObject (book: MacroDocument) (tableId: string) (row: int) : Value =
    let cells () =
        match tableOf book tableId |> Option.bind (fun t -> List.tryItem row t.Rows) with
        | Some r -> [ for c in 0 .. r.Cells.Length - 1 -> { TableId = tableId; Row = row; Col = c } ]
        | None -> []
    let mapRow change = mapTable book tableId (fun t -> { t with Rows = t.Rows |> List.mapi (fun i r -> if i = row then change r else r) })
    makeObject "Row"
        (fun name ->
            match name with
            | "range" -> let s, e = rowSpan book tableId row in Some (rangeOf book s e)
            | "cells" -> Some (collectionObject "Cells" (fun () -> cells () |> List.map (cellObject book) |> List.toArray) (fun _ -> None) noSet noCall)
            | "shading" -> Some (shadingObject book (CellShading cells))
            | "borders" -> Some (bordersObject book (CellBorders cells))
            | "index" -> Some (VInteger (row + 1))
            | "headingformat" -> Some (VBoolean (tableOf book tableId |> Option.bind (fun t -> List.tryItem row t.Rows) |> Option.exists (fun r -> r.IsHeader)))
            | "height" -> Some (VDouble (tableOf book tableId |> Option.bind (fun t -> List.tryItem row t.Rows) |> Option.bind (fun r -> r.Height) |> Option.map ptOfPx |> Option.defaultValue 0.0))
            | "isfirst" -> Some (VBoolean (row = 0))
            | "islast" -> Some (VBoolean (tableOf book tableId |> Option.exists (fun t -> row = t.Rows.Length - 1)))
            | "alignment" -> Some (VInteger 0)
            | _ -> None)
        (fun name value ->
            match name with
            | "headingformat" -> let on = toBool value in mapRow (fun r -> { r with IsHeader = on }); true
            | "height" -> let px = pxOfPt (toFloat value) in mapRow (fun r -> { r with Height = Some px; HeightRule = Some (defaultArg r.HeightRule "atLeast") }); true
            | "heightrule" -> let rule = match toInt value with 2 -> "exact" | 1 -> "atLeast" | _ -> "auto" in mapRow (fun r -> { r with HeightRule = Some rule }); true
            | "allowbreakacrosspages" -> let on = toBool value in mapRow (fun r -> { r with CanSplit = Some on }); true
            | "alignment" -> setTableAlignment book tableId (toInt value); true
            | _ -> false)
        (fun name _ ->
            match name with
            | "delete" ->
                let s, _ = rowSpan book tableId row
                let doc = setCursor (book.ModelPosition s).BlockId 0 book.Document
                book.SetStructure (deleteSelectedTableRow doc); Some VEmpty
            | "select" -> let s, e = rowSpan book tableId row in book.Selection.Start <- s; book.Selection.End <- e; Some VEmpty
            | _ -> None)
        None

and private setTableAlignment (book: MacroDocument) (tableId: string) (n: int) =
    let a = match n with 1 -> Some "center" | 2 -> Some "right" | _ -> Some "left"
    mapTable book tableId (fun t -> { t with Alignment = a })

and private rowsObject (book: MacroDocument) (tableId: string) : Value =
    let count () = match tableOf book tableId with Some t -> t.Rows.Length | None -> 0
    collectionObject "Rows" (fun () -> [| for r in 0 .. count () - 1 -> rowObject book tableId r |])
        (fun name ->
            match name with
            | "alignment" -> Some (VInteger (match tableOf book tableId |> Option.bind (fun t -> t.Alignment) with Some "center" -> 1 | Some "right" -> 2 | _ -> 0))
            | "headingformat" -> Some (VBoolean false)
            | _ -> None)
        (fun name value ->
            match name with
            | "alignment" -> setTableAlignment book tableId (toInt value); true
            | "leftindent" -> let px = pxOfPt (toFloat value) in mapTable book tableId (fun t -> { t with Indent = Some px }); true
            | "allowbreakacrosspages" | "wraparoundtext" -> true
            | "headingformat" -> let on = toBool value in mapTable book tableId (fun t -> { t with Rows = t.Rows |> List.map (fun r -> { r with IsHeader = on }) }); true
            | "height" -> let px = pxOfPt (toFloat value) in mapTable book tableId (fun t -> { t with Rows = t.Rows |> List.map (fun r -> { r with Height = Some px; HeightRule = Some (defaultArg r.HeightRule "atLeast") }) }); true
            | _ -> false)
        (fun name args ->
            match name with
            | "add" ->
                let before = arg args 0 |> Option.bind (fun v -> match v with VHostObject h -> h.GetMember "Index" | _ -> None) |> Option.map toInt
                let n = count ()
                let rowIndex, place = match before with Some i -> i - 1, insertTableRowBeforeSelection | None -> n - 1, insertTableRowAfterSelection
                let s, _ = rowSpan book tableId rowIndex
                let doc = setCursor (book.ModelPosition s).BlockId 0 book.Document
                book.SetStructure (place doc)
                Some (rowObject book tableId (match before with Some i -> i - 1 | None -> n))
            | _ -> None)

and private columnsObject (book: MacroDocument) (tableId: string) : Value =
    let count () = match tableOf book tableId with Some t -> (t.Rows |> List.map (fun r -> r.Cells.Length) |> List.fold max 0) | None -> 0
    let columnObject (col: int) =
        let cells () =
            match tableOf book tableId with
            | Some t -> [ for r, row in List.indexed t.Rows do if col < row.Cells.Length then yield { TableId = tableId; Row = r; Col = col } ]
            | None -> []
        makeObject "Column"
            (fun name ->
                match name with
                | "cells" -> Some (collectionObject "Cells" (fun () -> cells () |> List.map (cellObject book) |> List.toArray) (fun _ -> None) noSet noCall)
                | "shading" -> Some (shadingObject book (CellShading cells))
                | "borders" -> Some (bordersObject book (CellBorders cells))
                | "index" -> Some (VInteger (col + 1))
                | "width" -> Some (VDouble (cells () |> List.tryHead |> Option.bind (cellOf book) |> Option.bind (fun c -> c.Width) |> Option.map ptOfPx |> Option.defaultValue 0.0))
                | _ -> None)
            (fun name value ->
                match name with
                | "width" | "preferredwidth" ->
                    let px = pxOfPt (toFloat value)
                    mapCells book (cells ()) (fun c -> { c with Width = Some px; PreferredWidth = Some { Kind = DxaWidth; Value = Some (px * 15.0) } })
                    mapTable book tableId (fun t ->
                        let n = count ()
                        let grid = if t.GridColumns.Length = n then t.GridColumns else List.replicate n { Kind = AutoWidth; Value = None }
                        { t with GridColumns = grid |> List.mapi (fun i g -> if i = col then { Kind = DxaWidth; Value = Some (px * 15.0) } else g) })
                    true
                | "preferredwidthtype" -> true
                | _ -> false)
            (fun name _ ->
                match name with
                | "select" -> Some VEmpty
                | "delete" ->
                    let s, _ = cellSpan book { TableId = tableId; Row = 0; Col = col }
                    let doc = setCursor (book.ModelPosition s).BlockId 0 book.Document
                    book.SetStructure (deleteSelectedTableColumn doc); Some VEmpty
                | "autofit" -> Some VEmpty
                | _ -> None)
            None
    collectionObject "Columns" (fun () -> [| for c in 0 .. count () - 1 -> columnObject c |])
        (fun _ -> None)
        (fun name _ -> match name with "preferredwidthtype" | "preferredwidth" -> true | _ -> false)
        (fun name args ->
            match name with
            | "add" ->
                let n = count ()
                let s, _ = cellSpan book { TableId = tableId; Row = 0; Col = n - 1 }
                let doc = setCursor (book.ModelPosition s).BlockId 0 book.Document
                book.SetStructure (insertTableColumnAfterSelection doc)
                ignore args
                Some (columnObject n)
            | "autofit" | "distributewidth" -> Some VEmpty
            | _ -> None)

and private tablesObject (book: MacroDocument) (bounds: unit -> int * int) : Value =
    let items () =
        let s, e = bounds ()
        let e = max e (s + 1)
        book.TableSpans() |> List.filter (fun (_, ts, te) -> ts < e && te > s) |> List.map (fun (id, _, _) -> tableObject book id) |> List.toArray
    collectionObject "Tables" items (fun _ -> None) noSet
        (fun name args ->
            match name with
            | "add" ->
                let at =
                    match arg args 0 |> Option.bind rangeBounds with
                    | Some (s, _) -> s
                    | None -> fail 13 "Type mismatch (Tables.Add needs a Range)"
                let rows = max 1 (intArg args 1 1)
                let cols = max 1 (intArg args 2 1)
                let id = addTable book at rows cols
                // (Word's default, wdWord8TableBehavior, makes a table without borders; wdWord9TableBehavior
                // one in the Table Grid style.)
                if intArg args 3 0 <> 1 then mapTable book id (fun t -> { t with Borders = None; StyleId = None })
                Some (tableObject book id)
            | _ -> None)

/// A table of `rows` x `cols` where position `at` is: an empty paragraph there is the one after the
/// table; a paragraph with text is split there.
and private addTable (book: MacroDocument) (at: int) (rows: int) (cols: int) : string =
    let st = book.Story
    let seg = st.Segments.[locate st at]
    if seg.Table <> "" then notAvailable "a table inside a table cell"
    let offset = at - seg.Start
    // The paragraph the table goes before.
    let paraId =
        if offset > 0 && offset <= seg.Length then
            let next = splitBlock (setCursor seg.Para offset book.Document)
            book.SetStructure next
            match next.Selection with Some sel -> sel.Focus.BlockId | None -> seg.Para
        else seg.Para
    let doc = book.Document
    let index = doc.Blocks |> List.findIndex (function Paragraph p | Heading (_, p) -> p.Id = paraId | _ -> false)
    // (insertTable puts the table after the block the caret is in: after the paragraph before, or the
    // paragraph is moved after it.)
    let withTable = insertTable rows cols (setCursor paraId 0 doc)
    let table = withTable.Blocks |> List.item (index + 1)
    let blocks =
        let before = doc.Blocks |> List.take index
        let after = doc.Blocks |> List.skip index
        before @ [ table ] @ after
    // The cells' paragraphs have the style of the paragraph the table was put at.
    let para = doc.Blocks |> List.item index |> getParagraph
    let table =
        match table with
        | Table t ->
            let cellPara (p: ParagraphBlock) = { p with Metadata = { p.Metadata with Style = para.Metadata.Style; SpacingAfter = para.Metadata.SpacingAfter; SpacingBefore = para.Metadata.SpacingBefore; LineSpacing = para.Metadata.LineSpacing; LineSpacingRule = para.Metadata.LineSpacingRule; Explicit = para.Metadata.Explicit &&& (ParagraphExplicit.SpacingAfter ||| ParagraphExplicit.SpacingBefore ||| ParagraphExplicit.LineSpacing) } }
            Table { t with Rows = t.Rows |> List.map (fun r -> { r with Cells = r.Cells |> List.map (fun c -> { c with Content = c.Content |> List.map cellPara }) }) }
        | other -> other
    let blocks = blocks |> List.mapi (fun i b -> if i = index then table else b)
    book.SetStructure { doc with Blocks = blocks }
    match table with
    | Table t -> for row in t.Rows do for cell in row.Cells do for p in cell.Content do book.InheritMark(paraId, p.Id)
    | _ -> ()
    match table with Table t -> t.Id | _ -> ""

and private convertToTable (book: MacroDocument) (r: LiveRange) (args: Value array) : Value =
    let separator =
        match arg args 0 with
        | Some (VString s) when s <> "" -> s
        | Some v -> (match toInt v with 0 -> "\r" | 1 -> "\t" | 2 -> "," | n -> string (char n))
        | None -> "\t"
    let paras = book.SegmentsIn(r.Start, max r.Start (r.End - 1)) |> List.filter (fun s -> s.Para <> "" && s.Table = "")
    let lines = paras |> List.map (fun s -> book.Story.Text.Substring(s.Start, s.Length))
    let cells = lines |> List.map (fun l -> l.Split([| separator |], StringSplitOptions.None))
    let cols = max (intArg args 2 0) (cells |> List.map Array.length |> List.fold max 1)
    let rows = cells.Length
    if rows = 0 then fail 4605 "This method or property is not available because the range has no paragraphs."
    let first = paras.Head.Start
    let last = segmentEnd (List.last paras)
    // The paragraphs go; the table takes their place, cell by cell.
    let s, _ = book.Replace(first, min last (book.Story.Total - 1), "")
    let tableId = addTable book s rows cols
    cells |> List.iteri (fun ri row ->
        row |> Array.iteri (fun ci text ->
            if ci < cols && text <> "" then
                let cs, ce = cellSpan book { TableId = tableId; Row = ri; Col = ci }
                book.Replace(cs, ce - 1, text) |> ignore))
    tableObject book tableId

and private cellsInRange (book: MacroDocument) (r: LiveRange) : Value =
    let items () =
        book.SegmentsIn(r.Start, r.End) |> List.filter (fun s -> s.Table <> "" && s.Mark <> RowMark)
        |> List.map (fun s -> { TableId = s.Table; Row = s.Row; Col = s.Col }) |> List.distinct
    let cellsHost = CellShading items
    collectionObject "Cells" (fun () -> items () |> List.map (cellObject book) |> List.toArray)
        (fun name ->
            match name with
            | "shading" -> Some (shadingObject book cellsHost)
            | "borders" -> Some (bordersObject book (CellBorders items))
            | _ -> None)
        (fun name value ->
            match name with
            | "verticalalignment" ->
                let v = match toInt value with 1 -> Some "center" | 3 -> Some "bottom" | _ -> Some "top"
                mapCells book (items ()) (fun x -> { x with VerticalAlignment = v }); true
            | _ -> false)
        (fun name _ -> match name with "merge" -> notAvailable "Cells.Merge" | _ -> None)

and private rowsInRange (book: MacroDocument) (r: LiveRange) : Value =
    let rows () =
        book.SegmentsIn(r.Start, r.End) |> List.filter (fun s -> s.Table <> "") |> List.map (fun s -> s.Table, s.Row) |> List.distinct
    collectionObject "Rows" (fun () -> rows () |> List.map (fun (t, i) -> rowObject book t i) |> List.toArray) (fun _ -> None) noSet noCall

// ---- Find ---------------------------------------------------------------------------------------------------

and private findObject (book: MacroDocument) (r: LiveRange) (isSelection: bool) : Value =
    let state = findStateOf r
    let replacement =
        makeObject "Replacement"
            (fun name ->
                match name with
                | "text" -> Some (VString state.ReplacementText)
                | "font" -> Some (fontObject book (SpecFont state.ReplacementFont))
                | "highlight" -> Some (VInteger (if state.ReplacementHighlight < 0 then wdUndefined else state.ReplacementHighlight))
                | "style" -> Some (VString (defaultArg state.ReplacementStyle ""))
                | "paragraphformat" -> Some (specParagraphFormat state.ReplacementParagraph)
                | _ -> None)
            (fun name value ->
                match name with
                | "text" -> state.ReplacementText <- toText value; true
                | "highlight" -> state.ReplacementHighlight <- (if toBool value then 1 else 0); true
                | "style" ->
                    state.ReplacementStyle <-
                        (match scalarOf value with
                         | VInteger n -> builtInStyleName n
                         | v -> Some (toText v))
                    true
                | _ -> false)
            (fun name _ ->
                match name with
                | "clearformatting" ->
                    state.ReplacementFont.Clear(); state.ReplacementHighlight <- -1; state.ReplacementStyle <- None; state.ReplacementParagraph.Clear()
                    Some VEmpty
                | _ -> None)
            None
    let execute (args: Value array) : Value =
        arg args 0 |> Option.iter (fun v -> state.Text <- toText v)
        arg args 1 |> Option.iter (fun v -> state.MatchCase <- toBool v)
        arg args 2 |> Option.iter (fun v -> state.MatchWholeWord <- toBool v)
        arg args 3 |> Option.iter (fun v -> state.MatchWildcards <- toBool v)
        arg args 6 |> Option.iter (fun v -> state.Forward <- toBool v)
        arg args 7 |> Option.iter (fun v -> state.Wrap <- toInt v)
        arg args 8 |> Option.iter (fun v -> state.Format <- toBool v)
        arg args 9 |> Option.iter (fun v -> state.ReplacementText <- toText v)
        let replace = intArg args 10 0
        let total = book.Story.Total
        if replace = 2 then
            // Every match: in the range (the whole document for a caret that wraps, or a document's Content).
            let s, e =
                if r.End > r.Start && not r.Matched then r.Start, r.End
                elif state.Wrap = 1 then 0, total
                else r.Start, total
            let matches = findAll book state s e
            for m in List.rev matches do replaceMatch book state m |> ignore
            state.Found <- not matches.IsEmpty
            VBoolean state.Found
        else
            let region =
                if state.Forward then
                    if r.Matched || r.End = r.Start then r.End, total
                    else r.Start, r.End
                else
                    if r.Matched || r.End = r.Start then 0, r.Start else r.Start, r.End
            let candidates = findAll book state (fst region) (snd region)
            let pick (ms: (int * int * Text.RegularExpressions.Match option) list) = if state.Forward then List.tryHead ms else List.tryLast ms
            let found =
                match pick candidates with
                | Some m -> Some m
                | None when state.Wrap = 1 ->
                    let s, e = if state.Forward then 0, fst region else snd region, total
                    pick (findAll book state s e)
                | None -> None
            match found with
            | Some m ->
                let (s, e, _) = m
                let s, e = if replace = 1 then replaceMatch book state m else s, e
                r.Start <- s; r.End <- e; r.Matched <- true
                state.Found <- true
                VBoolean true
            | None ->
                state.Found <- false
                VBoolean false
    makeObject "Find"
        (fun name ->
            match name with
            | "text" -> Some (VString state.Text)
            | "replacement" -> Some replacement
            | "forward" -> Some (VBoolean state.Forward)
            | "wrap" -> Some (VInteger state.Wrap)
            | "format" -> Some (VBoolean state.Format)
            | "matchcase" -> Some (VBoolean state.MatchCase)
            | "matchwholeword" -> Some (VBoolean state.MatchWholeWord)
            | "matchwildcards" -> Some (VBoolean state.MatchWildcards)
            | "matchsoundslike" | "matchallwordforms" | "matchprefix" | "matchsuffix" | "matchdiacritics" | "matchkashida" | "matchalefhamza" | "matchcontrol" | "matchbyte" | "matchfuzzy" | "ignorespace" | "ignorepunct" -> Some (VBoolean false)
            | "found" -> Some (VBoolean state.Found)
            | "font" -> Some (fontObject book (SpecFont state.Font))
            | "highlight" -> Some (VInteger (if state.Highlight < 0 then wdUndefined else state.Highlight))
            | "parent" -> Some (if isSelection then selectionObject book else rangeObject book r)
            | _ -> None)
        (fun name value ->
            match name with
            | "text" -> state.Text <- toText value; true
            | "forward" -> state.Forward <- toBool value; true
            | "wrap" -> state.Wrap <- toInt value; true
            | "format" -> state.Format <- toBool value; true
            | "matchcase" -> state.MatchCase <- toBool value; true
            | "matchwholeword" -> state.MatchWholeWord <- toBool value; true
            | "matchwildcards" -> state.MatchWildcards <- toBool value; true
            | "matchsoundslike" | "matchallwordforms" | "matchprefix" | "matchsuffix" | "matchdiacritics" | "matchkashida" | "matchalefhamza" | "matchcontrol" | "matchbyte" | "matchfuzzy" | "ignorespace" | "ignorepunct" -> true
            | "highlight" -> state.Highlight <- (if toBool value then 1 else 0); true
            | "style" -> true
            | _ -> false)
        (fun name args ->
            match name with
            | "execute" -> Some (execute args)
            | "clearformatting" -> state.Font.Clear(); state.Highlight <- -1; Some VEmpty
            | "clearallfuzzyoptions" | "setallfuzzyoptions" -> Some VEmpty
            | _ -> None)
        None

and private specParagraphFormat (spec: Dictionary<string, Value>) : Value =
    makeObject "ParagraphFormat"
        (fun name -> match spec.TryGetValue name with | true, v -> Some v | _ -> (paragraphRead noFallback name defaultMetadata |> Option.map (fun _ -> VInteger wdUndefined)))
        (fun name value -> match paragraphChange name value with Some _ -> spec.[name] <- value; true | None -> false)
        noCall None

// ---- Selection ---------------------------------------------------------------------------------------------

and private selectionObject (book: MacroDocument) : Value =
    let r () = book.Selection
    let collapsedAt p = let s = r () in s.Start <- p; s.End <- p; s.ActiveAtStart <- false
    /// Move the selection's active end to `p` (extend) or the caret to `p`.
    let moveTo (p: int) (extend: bool) =
        let s = r ()
        let total = book.Story.Total
        let p = max 0 (min (total - 1) p)
        if extend then
            let anchor = if s.ActiveAtStart then s.End else s.Start
            s.Start <- min anchor p; s.End <- max anchor p
            s.ActiveAtStart <- p < anchor
        else collapsedAt p
    let active () = let s = r () in if s.ActiveAtStart then s.Start else s.End
    let typeText (text: string) =
        let s = r ()
        if book.PendingAt <> s.Start || s.End > s.Start then book.Pending.Clear()
        let a, e = book.Replace(s.Start, s.End, text, s)
        if book.Pending.Count > 0 then book.Format(a, e, specChange book.Pending)
        collapsedAt e
        book.PendingAt <- e
    let lineStart p = let st = book.Story in st.Segments.[locate st p].Start
    let lineEnd p = let st = book.Story in let seg = st.Segments.[locate st p] in seg.Start + seg.Length
    makeObject "Selection"
        (fun name ->
            match name with
            | "font" ->
                let s = r ()
                if s.End = s.Start then
                    if book.PendingAt <> s.Start then
                        book.Pending.Clear()
                        book.PendingAt <- s.Start
                    // (At a caret: what is typed next gets it; reading it gives the caret's formatting.)
                    let pending = SpecFont book.Pending
                    let caret = fontObject book (RangeFont s)
                    let caretHost = match caret with VHostObject h -> h | _ -> failwith "font"
                    Some (makeObject "Font"
                            (fun n -> match book.Pending.TryGetValue n with | true, v -> Some v | _ -> caretHost.GetMember n)
                            (fun n v -> match fontObject book pending with VHostObject h -> h.SetMember n v | _ -> false)
                            noCall None)
                else Some (fontObject book (RangeFont s))
            | "range" -> Some (rangeOf book (r ()).Start (r ()).End)
            | "find" -> Some (findObject book (r ()) true)
            | "type" -> Some (VInteger (if (r ()).End > (r ()).Start then 2 else 1))
            | "storytype" -> Some (VInteger 1)
            | "active" -> Some (VBoolean true)
            | "isendofrowmark" -> Some (VBoolean false)
            | _ -> rangeGet book (r ()) name)
        (fun name value ->
            match name with
            | "start" | "end" ->
                book.Pending.Clear()
                rangeSet book (r ()) name value
            | _ -> rangeSet book (r ()) name value)
        (fun name args ->
            let extend index = intArg args index 0 = 1
            match name with
            | "typetext" -> typeText (toText (defaultArg (arg args 0) (VString ""))); Some VEmpty
            | "typeparagraph" -> typeText "\r"; Some VEmpty
            | "typebackspace" ->
                let s = r ()
                if s.End > s.Start then
                    book.Replace(s.Start, s.End, "", s) |> ignore
                    collapsedAt s.Start
                elif s.Start > 0 then
                    let p, _ = book.Replace(s.Start - 1, s.Start, "", s)
                    collapsedAt p
                Some VEmpty
            | "insertparagraph" | "insertparagraphafter" | "insertafter" | "insertbefore" | "insertparagraphbefore" | "insertbreak" ->
                rangeCall book (r ()) name args
            | "homekey" | "endkey" ->
                let unit = intArg args 0 5
                let s = r ()
                let p =
                    match unit, name with
                    | 6, "homekey" -> 0
                    | 6, _ -> book.Story.Total - 1
                    | 10, "homekey" | 15, "homekey" -> fst (spanAt book unit s.Start)
                    | 10, _ | 15, _ -> snd (spanAt book unit s.Start) - 1
                    | _, "homekey" -> lineStart s.Start
                    | _, _ -> lineEnd (active ())
                moveTo p (extend 1)
                Some (VInteger 1)
            | "moveright" | "moveleft" ->
                let unit = intArg args 0 1
                let count = intArg args 1 1
                let count = if name = "moveleft" then -count else count
                let s = r ()
                let ext = extend 2
                if not ext && s.End > s.Start then
                    let collapsed = if count > 0 then s.End else s.Start
                    collapsedAt (moveBy book unit collapsed (if count > 0 then count - 1 else count + 1))
                elif unit = 12 then
                    // (A cell to the right: the next cell, selected as Word selects it.)
                    let spans = unitSpans book 12
                    let i = spans |> Array.tryFindIndex (fun (a, b) -> a <= s.Start && s.Start < b) |> Option.defaultValue -1
                    let j = i + count
                    if j >= 0 && j < spans.Length then
                        let a, b = spans.[j]
                        s.Start <- a; s.End <- b - 1; s.ActiveAtStart <- false
                else moveTo (moveBy book unit (active ()) count) ext
                Some (VInteger (abs count))
            | "moveup" | "movedown" ->
                let unit = intArg args 0 5
                let count = intArg args 1 1
                let count = if name = "moveup" then -count else count
                let unit = if unit = 5 || unit = 7 then 4 else unit
                let from = active ()
                let target =
                    if count > 0 then moveBy book unit from count
                    else moveBy book unit (lineStart from) (count + (if lineStart from = from then 0 else 1))
                moveTo target (extend 2)
                Some (VInteger (abs count))
            | "wholestory" -> let s = r () in s.Start <- 0; s.End <- book.Story.Total; Some VEmpty
            | "collapse" -> book.Pending.Clear(); rangeCall book (r ()) name args
            | "delete" ->
                let s = r ()
                if s.End > s.Start then
                    book.Replace(s.Start, s.End, "", s) |> ignore
                    collapsedAt s.Start
                else
                    let count = intArg args 1 1
                    book.Replace(s.Start, moveBy book (intArg args 0 1) s.Start count, "", s) |> ignore
                    collapsedAt s.Start
                Some (VInteger 1)
            | "escapekey" -> collapsedAt (r ()).Start; Some VEmpty
            | "extend" -> Some VEmpty
            | "selectrow" | "selectcell" | "selectcolumn" ->
                let s = r ()
                let a, b = spanAt book (if name = "selectrow" then 10 else 12) s.Start
                s.Start <- a; s.End <- b; Some VEmpty
            | "goto" -> Some VEmpty
            | "copy" | "cut" | "paste" | "pasteandformat" -> notAvailable ("Selection." + name + " (the clipboard)")
            | _ -> rangeCall book (r ()) name args)
        None

// ---- bookmarks, styles ----------------------------------------------------------------------------------------

and private bookmarksObject (book: MacroDocument) (within: LiveRange option) : Value =
    let infos () =
        bookmarkInfos book.Document
        |> List.map (fun info -> info, book.PositionOf(info.BlockId, info.StartOffset), book.PositionOf(info.BlockId, info.EndOffset))
        |> List.filter (fun (_, s, e) -> match within with Some r -> s >= r.Start && e <= r.End | None -> true)
        |> List.sortBy (fun (_, s, _) -> s)
    let bookmark (info: BookmarkInfo, s: int, e: int) =
        makeObject "Bookmark"
            (fun name ->
                match name with
                | "" | "_default" | "name" -> Some (VString info.Name)
                | "range" -> Some (rangeOf book s e)
                | "start" -> Some (VInteger s)
                | "end" -> Some (VInteger e)
                | "empty" -> Some (VBoolean (s = e))
                | _ -> None)
            noSet
            (fun name _ ->
                match name with
                | "delete" -> book.SetFormatting (removeBookmark info.Id book.Document); Some VEmpty
                | "select" -> book.Selection.Start <- s; book.Selection.End <- e; Some VEmpty
                | _ -> None)
            None
    collectionObject "Bookmarks" (fun () -> infos () |> List.map bookmark |> List.toArray)
        (fun _ -> None) noSet
        (fun name args ->
            match name with
            | "" | "_default" | "item" ->
                match arg args 0 with
                | Some (VString key) ->
                    match infos () |> List.tryFind (fun (info, _, _) -> eq info.Name key) with
                    | Some b -> Some (bookmark b)
                    | None -> fail 5941 (sprintf "The requested member of the collection does not exist (bookmark \"%s\")." key)
                | _ -> None
            | "exists" ->
                let key = toText (defaultArg (arg args 0) (VString ""))
                Some (VBoolean (infos () |> List.exists (fun (info, _, _) -> eq info.Name key)))
            | "add" ->
                let name' = toText (defaultArg (arg args 0) (VString ""))
                if name' = "" || not (Char.IsLetter name'.[0]) then fail 5828 "Bookmark names must begin with a letter."
                let s, e =
                    match arg args 1 |> Option.bind rangeBounds with
                    | Some b -> b
                    | None -> book.Selection.Start, book.Selection.End
                // An old one of that name goes; the new one is on the text of its first paragraph.
                match infos () |> List.tryFind (fun (info, _, _) -> eq info.Name name') with
                | Some (info, _, _) -> book.SetFormatting (removeBookmark info.Id book.Document)
                | None -> ()
                let seg = book.Story.Segments.[locate book.Story s]
                if seg.Para <> "" then
                    let a = s - seg.Start
                    let b = min (seg.Start + seg.Length) e - seg.Start
                    let id = string (book.NextBookmarkId())
                    book.UpdateParagraph(seg.Para, fun p -> { p with Buffer = TB.setBookmark a (max 0 (b - a)) id name' p.Buffer })
                match infos () |> List.tryFind (fun (info, _, _) -> eq info.Name name') with
                | Some b -> Some (bookmark b)
                | None -> Some VNothing
            | _ -> None)

and private styleObject (book: MacroDocument) (id: string) : Value =
    let style () = book.Document.Styles |> List.tryFind (fun s -> s.Id = id)
    makeObject "Style"
        (fun name ->
            match name with
            | "" | "_default" | "namelocal" | "name" -> Some (VString (match style () with Some s -> styleDisplayName s | None -> id))
            | "font" -> Some (fontObject book (StyleFont id))
            | "paragraphformat" -> Some (styleParagraphFormat book id)
            | "type" -> Some (VInteger (match style () |> Option.map (fun s -> s.StyleType) with Some "character" -> 2 | Some "table" -> 3 | _ -> 1))
            | "baseStyle" | "basestyle" -> Some (VString (style () |> Option.bind (fun s -> s.BasedOn) |> Option.defaultValue ""))
            | "builtin" -> Some (VBoolean true)
            | "inuse" -> Some (VBoolean true)
            | "quickstyle" -> Some (VBoolean (style () |> Option.exists (fun s -> s.IsQuickStyle)))
            | _ -> None)
        (fun name value ->
            match name with
            | "quickstyle" | "automaticallyupdate" | "nofollowup" -> true
            | "nextparagraphstyle" -> book.ChangeStyle(id, fun s -> { s with Next = Some (toText value |> compact) }); true
            | "basestyle" -> book.ChangeStyle(id, fun s -> { s with BasedOn = Some (toText value |> compact) }); true
            | _ -> false)
        noCall None

and private styleParagraphFormat (book: MacroDocument) (id: string) : Value =
    makeObject "ParagraphFormat"
        (fun name ->
            let s = book.Document.Styles |> List.tryFind (fun x -> x.Id = id)
            paragraphRead (match s with Some x -> fallbackOfNamed x | None -> noFallback) name defaultMetadata)
        (fun name value ->
            match name with
            | "alignment" -> book.ChangeStyle(id, fun s -> { s with Alignment = Some (alignmentOf (toInt value)) }); true
            | "spacebefore" -> book.ChangeStyle(id, fun s -> { s with SpacingBefore = Some (toFloat value) }); true
            | "spaceafter" -> book.ChangeStyle(id, fun s -> { s with SpacingAfter = Some (toFloat value) }); true
            | "leftindent" -> book.ChangeStyle(id, fun s -> { s with IndentLeft = Some (pxOfPt (toFloat value)) }); true
            | "rightindent" -> book.ChangeStyle(id, fun s -> { s with IndentRight = Some (pxOfPt (toFloat value)) }); true
            | "firstlineindent" -> book.ChangeStyle(id, fun s -> { s with FirstLineIndent = Some (pxOfPt (toFloat value)) }); true
            | "keepwithnext" -> book.ChangeStyle(id, fun s -> { s with KeepNext = Some (toBool value) }); true
            | "keeptogether" -> book.ChangeStyle(id, fun s -> { s with KeepLines = Some (toBool value) }); true
            | "linespacing" -> book.ChangeStyle(id, fun s -> { s with LineSpacing = Some (toFloat value / 12.0); LineSpacingRule = Some "auto" }); true
            | "linespacingrule" -> true
            | _ -> false)
        noCall None

and private stylesObject (book: MacroDocument) : Value =
    let lookup (key: Value) =
        let name =
            match scalarOf key with
            | VInteger n -> (match builtInStyleName n with Some s -> s | None -> fail 5941 "The requested member of the collection does not exist.")
            | VDouble d -> (match builtInStyleName (int d) with Some s -> s | None -> fail 5941 "The requested member of the collection does not exist.")
            | v -> toText v
        if eq name "Normal" && not (book.Document.Styles |> List.exists (fun s -> s.Id = "Normal")) then
            // (A new document's default paragraph style: Normal, as the document defaults make it.)
            let normal = { plainStyle with Id = "Normal"; Name = Some "Normal"; IsDefault = true; BasedOn = None; Next = None; UiPriority = Some 0 }
            book.SetFormatting { book.Document with Styles = book.Document.Styles @ [ normal ] }
        match book.FindStyle name with
        | Some s -> styleObject book s.Id
        | None -> fail 5941 (sprintf "The requested member of the collection does not exist (style \"%s\")." name)
    collectionObject "Styles"
        (fun () -> book.Document.Styles |> List.map (fun s -> styleObject book s.Id) |> List.toArray)
        (fun _ -> None) noSet
        (fun name args ->
            match name with
            | "" | "_default" | "item" -> arg args 0 |> Option.map lookup
            | "add" ->
                let name' = toText (defaultArg (arg args 0) (VString "Style1"))
                let kind = match intArg args 1 1 with 2 -> "character" | 3 -> "table" | _ -> "paragraph"
                let id = compact name'
                if not (book.Document.Styles |> List.exists (fun s -> s.Id = id)) then
                    book.SetFormatting { book.Document with Styles = book.Document.Styles @ [ { plainStyle with Id = id; Name = Some name'; StyleType = kind } ] }
                Some (styleObject book id)
            | _ -> None)

// ---- the document, the application ---------------------------------------------------------------------------

and private documentObject (book: MacroDocument) : Value =
    let content () = rangeOf book 0 book.Story.Total
    let pageSetup () =
        let px (v: float) = VDouble (ptOfPx v)
        makeObject "PageSetup"
            (fun name ->
                let p = book.Document.PageSetup
                match name with
                | "topmargin" -> Some (px p.MarginTop)
                | "bottommargin" -> Some (px p.MarginBottom)
                | "leftmargin" -> Some (px p.MarginLeft)
                | "rightmargin" -> Some (px p.MarginRight)
                | "pagewidth" -> Some (px p.PageWidth)
                | "pageheight" -> Some (px p.PageHeight)
                | "orientation" -> Some (VInteger (if p.PageWidth > p.PageHeight then 1 else 0))
                | "headerdistance" -> Some (px p.HeaderDistance)
                | "footerdistance" -> Some (px p.FooterDistance)
                | _ -> None)
            (fun name value ->
                let v = pxOfPt (toFloat value)
                let set change = book.SetStructure { book.Document with PageSetup = change book.Document.PageSetup }; true
                match name with
                | "topmargin" -> set (fun p -> { p with MarginTop = v })
                | "bottommargin" -> set (fun p -> { p with MarginBottom = v })
                | "leftmargin" -> set (fun p -> { p with MarginLeft = v })
                | "rightmargin" -> set (fun p -> { p with MarginRight = v })
                | "pagewidth" -> set (fun p -> { p with PageWidth = v })
                | "pageheight" -> set (fun p -> { p with PageHeight = v })
                | "headerdistance" -> set (fun p -> { p with HeaderDistance = v })
                | "footerdistance" -> set (fun p -> { p with FooterDistance = v })
                | "orientation" ->
                    let landscape = toInt value = 1
                    set (fun p -> if (p.PageWidth > p.PageHeight) = landscape then p else { p with PageWidth = p.PageHeight; PageHeight = p.PageWidth })
                | _ -> false)
            noCall None
    makeObject "Document"
        (fun name ->
            match name with
            | "content" -> Some (content ())
            | "paragraphs" -> Some (paragraphsObject book (fun () -> 0, book.Story.Total))
            | "words" -> Some (spansObject book "Words" (fun () -> 0, book.Story.Total) 2)
            | "characters" -> Some (spansObject book "Characters" (fun () -> 0, book.Story.Total) 1)
            | "sentences" -> Some (spansObject book "Sentences" (fun () -> 0, book.Story.Total) 3)
            | "tables" -> Some (tablesObject book (fun () -> 0, book.Story.Total))
            | "bookmarks" -> Some (bookmarksObject book None)
            | "styles" -> Some (stylesObject book)
            | "pagesetup" -> Some (pageSetup ())
            | "variables" -> Some (variablesObject book)
            | "name" -> Some (VString book.Hooks.FileName)
            | "fullname" -> Some (VString book.Hooks.FileName)
            | "path" -> Some (VString "")
            | "saved" -> Some (VBoolean false)
            | "readonly" -> Some (VBoolean false)
            | "trackrevisions" -> Some (VBoolean false)
            | "application" -> Some (applicationObject book)
            | "activewindow" -> Some (windowObject book)
            | "selection" -> Some (selectionObject book)
            | "range" | "computestatistics" | "undo" -> None
            | "inlineshapes" | "shapes" | "fields" | "comments" | "revisions" | "footnotes" | "endnotes" | "hyperlinks" | "tablesofcontents" | "lists" ->
                Some (collectionObject "Collection" (fun () -> [||]) (fun _ -> None) noSet noCall)
            | _ -> None)
        (fun name _ ->
            match name with
            | "saved" | "trackrevisions" | "showgrammaticalerrors" | "showspellingerrors" | "spellingchecked" | "grammarchecked" -> true
            | _ -> false)
        (fun name args ->
            match name with
            | "range" ->
                let total = book.Story.Total
                let s = intArg args 0 0
                let e = intArg args 1 (if arg args 0 |> Option.isSome && args.Length = 1 then s else total)
                let e = if args.Length = 0 then total else e
                Some (rangeOf book (max 0 (min total s)) (max 0 (min total e)))
            | "computestatistics" -> Some (VInteger (statistics book 0 book.Story.Total (intArg args 0 0)))
            | "activate" | "select" | "repaginate" | "undoclear" -> Some VEmpty
            | "save" | "saveas" | "saveas2" | "close" | "printout" | "exportasfixedformat" ->
                book.Hooks.Output (sprintf "(Document.%s is not run here: the host saves the document)" name)
                Some VEmpty
            | "undo" | "redo" -> Some (VBoolean false)
            | _ -> None)
        None

/// Document.Variables: values kept by name for the run (Add, Item(name).Value, Count).
and private variablesObject (book: MacroDocument) : Value =
    let names () = [ for kv in book.Settings do if kv.Key.StartsWith "var:" then yield kv.Key.Substring 4 ]
    let variable (key: string) =
        makeObject "Variable"
            (fun name ->
                match name with
                | "" | "_default" | "value" -> (match book.Settings.TryGetValue ("var:" + key) with | true, v -> Some v | _ -> Some VEmpty)
                | "name" -> Some (VString key)
                | _ -> None)
            (fun name value ->
                match name with
                | "" | "_default" | "value" -> book.Settings.["var:" + key] <- VString (toText value); true
                | _ -> false)
            (fun name _ -> match name with "delete" -> book.Settings.Remove("var:" + key) |> ignore; Some VEmpty | _ -> None)
            None
    let find (key: string) = names () |> List.tryFind (eq key)
    collectionObject "Variables" (fun () -> names () |> List.map variable |> List.toArray) (fun _ -> None) noSet
        (fun name args ->
            match name with
            | "" | "_default" | "item" ->
                match arg args 0 with
                | Some (VString key) -> (match find key with Some k -> Some (variable k) | None -> fail 5825 "Object has been deleted.")
                | _ -> None
            | "add" ->
                let key = toText (defaultArg (arg args 0) (VString ""))
                if (find key).IsSome then fail 4172 "The variable name already exists."
                book.Settings.["var:" + key] <- VString (toText (defaultArg (arg args 1) (VString "")))
                Some (variable key)
            | _ -> None)

and private windowObject (book: MacroDocument) : Value =
    let view =
        makeObject "View"
            (fun name -> match name with "type" -> Some (VInteger 3) | "showall" | "showhiddentext" | "showfieldcodes" -> Some (VBoolean false) | "zoom" -> Some (makeObject "Zoom" (fun n -> if n = "percentage" then Some (VInteger 100) else None) (fun _ _ -> true) noCall None) | _ -> None)
            (fun name _ -> match name with "type" | "showall" | "showhiddentext" | "showfieldcodes" | "showbookmarks" | "showparagraphs" | "draft" -> true | _ -> false)
            noCall None
    makeObject "Window"
        (fun name ->
            match name with
            | "view" -> Some view
            | "activepane" -> Some (makeObject "Pane" (fun n -> if n = "view" then Some view else None) noSet (fun n _ -> if n = "smallscroll" || n = "largescroll" then Some VEmpty else None) None)
            | "selection" -> Some (selectionObject book)
            | "caption" -> Some (VString book.Hooks.FileName)
            | "windowstate" -> Some (VInteger 0)
            | "displayrulers" | "displayverticalruler" | "displayhorizontalscrollbar" | "displayverticalscrollbar" -> Some (VBoolean true)
            | _ -> None)
        (fun name _ -> match name with "windowstate" | "caption" | "displayrulers" | "displayverticalruler" | "displayhorizontalscrollbar" | "displayverticalscrollbar" -> true | _ -> false)
        (fun name _ -> match name with "activate" | "scrollintoview" | "smallscroll" | "largescroll" -> Some VEmpty | _ -> None)
        None

and private optionsObject (book: MacroDocument) : Value =
    let known = [ "defaulthighlightcolorindex", VInteger 7; "replaceselection", VBoolean true; "checkspellingasyoutype", VBoolean true
                  "checkgrammarasyoutype", VBoolean true; "autoformatasyoutypereplacequotes", VBoolean true; "smartcutpaste", VBoolean true
                  "pagination", VBoolean true; "measurementunit", VInteger 0; "autoformatasyoutypeapplynumberedlists", VBoolean true ]
    makeObject "Options"
        (fun name ->
            match book.Settings.TryGetValue name with
            | true, v -> Some v
            | _ -> known |> List.tryPick (fun (k, v) -> if k = name then Some v else None))
        (fun name value ->
            if known |> List.exists (fun (k, _) -> k = name) then
                book.Settings.[name] <- value
                true
            else false)
        noCall None

and private applicationObject (book: MacroDocument) : Value =
    let flag name fallback = match book.Settings.TryGetValue ("app:" + name) with | true, v -> v | _ -> fallback
    makeObject "Application"
        (fun name ->
            match name with
            | "activedocument" | "thisdocument" -> Some (documentObject book)
            | "documents" -> Some (documentsObject book)
            | "selection" -> Some (selectionObject book)
            | "options" -> Some (optionsObject book)
            | "activewindow" -> Some (windowObject book)
            | "name" -> Some (VString "Microsoft Word")
            | "version" -> Some (VString "16.0")
            | "build" -> Some (VString "16.0")
            | "username" -> Some (flag "username" (VString "User"))
            | "userinitials" -> Some (VString "U")
            | "pathseparator" -> Some (VString "\\")
            | "screenupdating" -> Some (flag "screenupdating" (VBoolean true))
            | "displayalerts" -> Some (flag "displayalerts" (VInteger -1))
            | "visible" -> Some (VBoolean true)
            | "statusbar" -> Some (flag "statusbar" (VString ""))
            | "caption" -> Some (VString "Microsoft Word")
            | "application" -> Some (applicationObject book)
            | _ -> None)
        (fun name value ->
            match name with
            | "screenupdating" | "displayalerts" | "username" | "visible" | "windowstate" | "displaystatusbar" ->
                book.Settings.["app:" + name] <- value; true
            | "statusbar" -> book.Settings.["app:statusbar"] <- value; true
            | _ -> false)
        (fun name args ->
            let pts factor = Some (VDouble (toFloat (defaultArg (arg args 0) (VInteger 0)) * factor))
            match name with
            | "centimeterstopoints" -> pts (72.0 / 2.54)
            | "millimeterstopoints" -> pts (72.0 / 25.4)
            | "inchestopoints" -> pts 72.0
            | "linestopoints" -> pts 12.0
            | "picastopoints" -> pts 12.0
            | "pointstocentimeters" -> pts (2.54 / 72.0)
            | "pointstoinches" -> pts (1.0 / 72.0)
            | "pointstomillimeters" -> pts (25.4 / 72.0)
            | "screenrefresh" | "activate" | "quit" | "scroll" | "wait" -> Some VEmpty
            | _ -> None)
        None

and private documentsObject (book: MacroDocument) : Value =
    collectionObject "Documents" (fun () -> [| documentObject book |])
        (fun _ -> None) noSet
        (fun name args ->
            match name with
            | "add" ->
                // (One document is open in the editor: a new document is the editor's, made empty.)
                let para = newEmptyParagraph ()
                book.SetStructure { book.Document with Blocks = [ Paragraph para ] }
                book.Selection.Start <- 0; book.Selection.End <- 0
                ignore args
                Some (documentObject book)
            | "open" -> notAvailable "Documents.Open (files)"
            | "" | "_default" | "item" ->
                match arg args 0 with
                | Some (VString _) | Some (VInteger 1) | Some (VDouble 1.0) -> Some (documentObject book)
                | _ -> noSuchMember ()
            | _ -> None)

/// The objects a macro starts from, by name (Application, ActiveDocument, ThisDocument, Selection,
/// Documents, Options, ActiveWindow, the unit conversions), and the wd* constants.
let hostGlobals (book: MacroDocument) : (string * Value) list =
    let app = applicationObject book
    let appHost = match app with VHostObject h -> h | _ -> failwith "application"
    let forward (name: string) =
        // (A global that is what Application has now: its members, its default call and For Each.)
        let current () = match appHost.GetMember name with Some (VHostObject h) -> h | _ -> failwith name
        VHostObject (Host(name,
                          (fun m -> (current ()).GetMember m),
                          (fun m v -> (current ()).SetMember m v),
                          (fun m args -> (current ()).CallMethod m args),
                          Some (fun () -> match (current ()).Enumerate () with Some items -> items | None -> raiseVbError 451)) :> IHostObject)
    let conversion (name: string) =
        VHostObject (Host(name, (fun _ -> None), (fun _ _ -> false),
                          (fun m args -> if m = "" || eq m "_Default" then appHost.CallMethod name args else unsupported name m), None) :> IHostObject)
    [ "Application", app
      "ActiveDocument", forward "ActiveDocument"
      "ThisDocument", forward "ActiveDocument"
      "Selection", forward "Selection"
      "Documents", forward "Documents"
      "Options", forward "Options"
      "ActiveWindow", forward "ActiveWindow"
      "CentimetersToPoints", conversion "CentimetersToPoints"
      "MillimetersToPoints", conversion "MillimetersToPoints"
      "InchesToPoints", conversion "InchesToPoints"
      "LinesToPoints", conversion "LinesToPoints"
      "PointsToCentimeters", conversion "PointsToCentimeters"
      "PointsToInches", conversion "PointsToInches"
      "DoEvents", VHostObject (Host("DoEvents", (fun _ -> Some (VInteger 0)), (fun _ _ -> false), (fun _ _ -> Some (VInteger 0)), None) :> IHostObject) ]
    @ constants

// ---- running macros ---------------------------------------------------------------------------------------------

open Rbt.Vb.Api

type MacroResult =
    | MacroDone of Document
    | MacroFailed of VbFailure

/// How many VBA instructions a run may take before it is stopped as a runaway.
let defaultInstructionLimit = 300_000_000

/// The kind of a module, from its Attribute lines: ThisDocument is a document module, class modules are
/// classes, the rest standard modules.
let moduleKind (source: string) : ModuleKind =
    let header = ModuleSource.parseHeader source
    let attribute name = header.Attributes |> List.tryPick (fun (key, value) -> if eq key name then Some (value.Trim('"')) else None)
    match attribute "VB_Base" with
    | Some _ -> ModuleKind.Document
    | None ->
        match attribute "VB_Creatable", attribute "VB_Exposed", attribute "VB_PredeclaredId" with
        | None, None, None -> Standard
        | _ -> Class

/// The source as the engine compiles it. (Rbt.Vb 0.4.0 does not take `Next c` after `For Each c`, only
/// `Next`: the counter's name after a Next that ends a line is dropped, which changes nothing else - the
/// lines stay where they are.)
let compiledSource (source: string) =
    Text.RegularExpressions.Regex.Replace(
        source, "^([ \\t]*(?:\\d+[ \\t]+)?Next)[ \\t]+[A-Za-z_][A-Za-z0-9_]*[ \\t]*(?=\\r?$|')",
        "$1", Text.RegularExpressions.RegexOptions.Multiline ||| Text.RegularExpressions.RegexOptions.IgnoreCase)

let macroModule (streamName: string) (source: string) : ModuleSource =
    let header = ModuleSource.parseHeader source
    { Name = header.Name |> Option.defaultValue streamName; Kind = moduleKind source; Source = source }

/// A document's VBA project compiled against Word's object model, its module state kept between runs.
type MacroProgram private (project: Project, book: MacroDocument, session: Session) =
    static let hooksOf (book: MacroDocument) : RuntimeHooks =
        { Print = (fun text -> book.Hooks.Output text)
          MsgBox = (fun prompt buttons title -> book.Hooks.MessageBox prompt buttons title)
          InputBox = (fun prompt title defaultText -> book.Hooks.InputBox prompt title defaultText)
          CreateObject = (fun _ -> None)
          Environ = (fun name -> if eq name "USERNAME" then "User" else "") }
    let rec finish (status: ExecutionStatus) =
        match status with
        | Paused _ -> finish (session.RunToEnd defaultInstructionLimit)
        | Finished _ -> MacroDone (book.Result())
        | Failed failure -> MacroFailed failure
        | Running ->
            session.Cancel()
            match session.Step 1 with
            | Failed failure -> MacroFailed { failure with Description = "The macro ran too long and was stopped" }
            | _ -> MacroFailed { Kind = Runtime; Number = 18; Description = "The macro ran too long and was stopped"; Source = ""; Line = 0; Column = 0; Procedure = ""; Module = "" }

    static member Compile(modules: ModuleSource list) : Result<MacroProgram, VbFailure list> =
        let book = MacroDocument(Rbt.Word.Document.empty, WordMacroApi.defaultHooks)
        let options =
            ProjectOptions.vba
            |> ProjectOptions.withHostGlobals (Map.ofList (hostGlobals book))
            |> ProjectOptions.withHooks (hooksOf book)
        match Project.compile options (modules |> List.map (fun m -> { m with Source = compiledSource m.Source })) with
        | Result.Error failures -> Result.Error failures
        | Result.Ok project ->
            match Session.Create(project, defaultInstructionLimit) with
            | Result.Ok session -> Result.Ok (MacroProgram(project, book, session))
            | Result.Error failure -> Result.Error [ failure ]

    member _.Project = project
    member _.Macros : string list =
        Project.macros project |> List.map (fun p -> if p.ModuleKind = Standard then p.Name else p.QualifiedName)
    member _.Has(name: string) = (Project.tryFindProcedure name project).IsSome

    member this.Run(doc: Document, hooks: MacroHooks, name: string, ?args: Value list, ?maxInstructions: int) : MacroResult =
        book.Reset(doc, hooks)
        session.SetHooks(hooksOf book)
        match session.Call(name, defaultArg args []) with
        | Result.Error failure -> MacroFailed failure
        | Result.Ok () -> finish (session.RunToEnd(defaultArg maxInstructions defaultInstructionLimit))

    member this.Immediate(doc: Document, hooks: MacroHooks, line: string, ?moduleName: string) : MacroResult =
        book.Reset(doc, hooks)
        session.SetHooks(hooksOf book)
        match session.Immediate(line, ?moduleName = moduleName) with
        | Result.Error failure -> MacroFailed failure
        | Result.Ok () ->
            let status = session.RunToEnd defaultInstructionLimit
            match status with
            | Finished value when MacroProgram.Prints line -> hooks.Output (displayValue value)
            | _ -> ()
            finish status

    static member Prints(line: string) =
        let text = line.Trim()
        text.StartsWith "?" || text.StartsWith("print ", StringComparison.OrdinalIgnoreCase)

// ---- the macros of an opened document -----------------------------------------------------------------------

/// The modules of the VBA project an opened document keeps (Attribute lines and all).
let macroModulesOf (doc: Document) : ModuleSource list =
    match projectBytes doc with
    | Some bytes ->
        (try VbaProject.readModules bytes with _ -> [])
        |> List.filter (fun (_, source) -> not (String.IsNullOrWhiteSpace source))
        |> List.map (fun (stream, source) -> macroModule stream source)
    | None -> []

/// A module's source as (header, code): the Attribute lines the VBA editor does not show, and the code.
let splitHeader (source: string) : string * string =
    let lines = source.Replace("\r\n", "\n").Split('\n')
    let mutable index = 0
    let mutable inBegin = false
    let isHeader (line: string) =
        let text = line.Trim()
        if inBegin then
            (if text.Equals("END", StringComparison.OrdinalIgnoreCase) then inBegin <- false)
            true
        elif text.StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase) then true
        elif text.StartsWith("VERSION ", StringComparison.OrdinalIgnoreCase) then true
        elif text.Equals("BEGIN", StringComparison.OrdinalIgnoreCase) then inBegin <- true; true
        else false
    while index < lines.Length && isHeader lines.[index] do index <- index + 1
    String.concat "\n" lines.[.. index - 1], String.concat "\n" lines.[index ..]

let joinHeader (header: string) (code: string) =
    if header = "" then code else header + "\n" + code.Replace("\r\n", "\n")

/// A project of the modules, as vbaProject.bin (source only: Word compiles it when it opens the file).
let modulesBytes (modules: ModuleSource list) : byte[] =
    modules
    |> List.map (fun m ->
        let kind =
            match m.Kind with
            | Standard -> VbaProject.StandardModule
            | Class -> VbaProject.ClassModule
            | ModuleKind.Document -> VbaProject.DocumentModule
        m.Name, kind, m.Source)
    |> VbaProject.writeProjectModules
