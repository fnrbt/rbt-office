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
open Rbt.Vb.HostMembers
open Rbt.Word.Document
open Rbt.Word.WordMacroApi
module TB = Rbt.Word.TextBuffer
module HostMembers = Rbt.Vb.HostMembers

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
    | VInteger i | VLong i -> float i
    | VByte b -> float b
    | VSingle f -> float (Rbt.Vb.NumberText.singleAsDouble f)
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
    | VInteger i | VLong i -> string i
    | VByte b -> string b
    | VSingle _ -> Rbt.Vb.Api.displayValue value
    | VDouble d -> Rbt.Vb.Api.displayValue (VDouble d)
    | VDate _ | VCurrency _ as v -> Rbt.Vb.Api.displayValue v
    | VBoolean b -> if b then "True" else "False"
    | VEmpty | VNull | VNothing | VUndefined -> ""
    | _ -> fail 13 "Type mismatch"

let private toBool (value: Value) =
    match scalarOf value with
    | VBoolean b -> b
    | VInteger i | VLong i -> i <> 0
    | VByte b -> b <> 0uy
    | VSingle f -> f <> 0.0f
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
    if t.Length = 6 && t |> Seq.forall (fun ch -> Char.IsDigit ch || (ch >= 'a' && ch <= 'f') || (ch >= 'A' && ch <= 'F')) then
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
    | Some n -> VLong n
    | None -> VLong wdColorAutomatic

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
      "wdTableFormatNone", 0; "wdTableFormatSimple1", 1; "wdTableFormatSimple2", 2; "wdTableFormatSimple3", 3
      "wdTableFormatClassic1", 4; "wdTableFormatClassic2", 5; "wdTableFormatClassic3", 6; "wdTableFormatClassic4", 7
      "wdTableFormatColorful1", 8; "wdTableFormatColorful2", 9; "wdTableFormatColorful3", 10; "wdTableFormatColumns1", 11
      "wdTableFormatGrid1", 16; "wdTableFormatList1", 24; "wdTableFormat3DEffects1", 32; "wdTableFormatContemporary", 35; "wdTableFormatElegant", 36; "wdTableFormatProfessional", 37
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
    |> List.map (fun (name, value) -> name, VLong value)

// ---- the main story ------------------------------------------------------------------------------------

type Mark = ParagraphMark | CellMark | RowMark

/// A paragraph of the story (Para "" for an end-of-row mark): where its text starts, how long it is, the
/// mark after it, and the table cell (table id, 0-based row and column) it is in ("" for the body).
type Segment =
    { Para: string; Start: int; Length: int; Mark: Mark; Table: string; Row: int; Col: int
      /// 1 for a paragraph that starts with a page break (Chr(12) before its text, as Word has it), else 0.
      Lead: int }
    /// Where the paragraph's text starts (after its page break).
    member s.TextStart = s.Start + s.Lead

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
        // (A page break the paragraph starts with is a character of the story: Chr(12).)
        let lead =
            match para with
            | Some p when p.Metadata.PageBreakBefore && ParagraphExplicit.has ParagraphExplicit.PageBreakRunBefore p.Metadata -> 1
            | _ -> 0
        if lead = 1 then text.Append('\f') |> ignore
        let length =
            match para with
            | Some p -> let s = TB.toString p.Buffer in text.Append(s) |> ignore; s.Length
            | None -> 0
        text.Append(match mark with ParagraphMark -> '\r' | CellMark -> '\a' | RowMark -> '\b') |> ignore
        segments.Add { Para = (match para with Some p -> p.Id | None -> ""); Start = start; Length = length; Mark = mark; Table = table; Row = row; Col = col; Lead = lead }
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

let private segmentEnd (seg: Segment) = seg.TextStart + seg.Length + 1

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
        if c = '\r' || c = '\a' || c = '\b' || c = '\t' then ()
        elif isWordChar c then
            while j < n && (isWordChar text.[j] || ((text.[j] = '\'' || text.[j] = '’') && j + 1 < n && isWordChar text.[j + 1])) do j <- j + 1
            while j < n && text.[j] = ' ' do j <- j + 1
        elif c = ' ' then
            while j < n && text.[j] = ' ' do j <- j + 1
        else
            // (A run of punctuation is one word to Word: "++", "%)".)
            while j < n && not (isWordChar text.[j]) && text.[j] <> ' ' && text.[j] <> '\t' && text.[j] <> '\r' && text.[j] <> '\a' && text.[j] <> '\b' && text.[j] <> '\f' do j <- j + 1
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
          Id = "Caption"; Name = Some "caption"; Text = Some (textOf (Some 9.0) None (Some "#44546A")); TextItalic = Some true
          SpacingAfter = Some 10.0; LineSpacing = Some 1.0; LineSpacingRule = Some "auto"; UiPriority = Some 35 }
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
    | Some name when name.Length > 0 && Char.IsLower name.[0] -> string (Char.ToUpperInvariant name.[0]) + name.Substring 1
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
        | Some seg -> seg.TextStart + max 0 (min seg.Length offset)
        | None -> 0

    /// The model's position of a story position (an end-of-row mark is the end of its row's last cell).
    member this.ModelPosition(pos: int) : Position =
        let st = this.Story
        let i = locate st pos
        let seg = st.Segments.[i]
        if seg.Para = "" then
            let rec back j = if j < 0 then { BlockId = ""; Offset = 0 } elif st.Segments.[j].Para <> "" then { BlockId = st.Segments.[j].Para; Offset = st.Segments.[j].Length } else back (j - 1)
            back (i - 1)
        else { BlockId = seg.Para; Offset = max 0 (min seg.Length (pos - seg.TextStart)) }

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
                // (A page break deleted: the paragraph no longer starts with one.)
                for seg in segs do
                    if seg.Lead = 1 && seg.Start >= s && seg.Start < e then
                        this.UpdateParagraph(seg.Para, fun p ->
                            { p with Metadata = { p.Metadata with PageBreakBefore = false; Explicit = p.Metadata.Explicit &&& ~~~ParagraphExplicit.PageBreakRunBefore } })
                let place (seg: Segment) = seg.Table, seg.Row, seg.Col
                let crossesCells =
                    segs |> Array.exists (fun seg -> seg.Mark <> ParagraphMark && segmentEnd seg <= e)
                    || segs |> Array.exists (fun seg -> place seg <> place segs.[0])
                if crossesCells then
                    // (Cells keep their marks: what is deleted is the text in them.)
                    for seg in Array.rev segs do
                        if seg.Para <> "" then
                            let a = max s seg.TextStart - seg.TextStart
                            let b = min e (seg.TextStart + seg.Length) - seg.TextStart
                            if b > a then this.UpdateParagraph(seg.Para, fun p -> { p with Buffer = TB.delete a (b - a) p.Buffer })
                elif first = last then
                    let seg = segs.[0]
                    let a = max 0 (s - seg.TextStart)
                    let b = max a (e - seg.TextStart)
                    this.UpdateParagraph(seg.Para, fun p -> { p with Buffer = TB.delete a (b - a) p.Buffer })
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
            let totalBefore = this.Story.Total
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
                    if pageBreak then
                        // (Word's page break is a paragraph of its own, Chr(12) and a paragraph mark; the
                        // model's is an empty paragraph that starts with the break: the one split off here.)
                        this.UpdateParagraph(para, fun p ->
                            { p with Metadata = { p.Metadata with PageBreakBefore = true; Explicit = p.Metadata.Explicit ||| ParagraphExplicit.PageBreakRunBefore } })
                    para <- newId
                    off <- 0
                | _ -> notAvailable "the paragraph cannot be split here"
            let piece = Text.StringBuilder()
            for ch in text do
                if ch = '\r' then
                    flush (piece.ToString())
                    piece.Clear() |> ignore
                    split false
                elif ch = '\f' then
                    flush (piece.ToString())
                    piece.Clear() |> ignore
                    // (At the start of a paragraph the break goes before it; elsewhere the paragraph is split
                    // first. Right after a table Word keeps an empty paragraph before the break.)
                    let afterTable =
                        let rec look (blocks: Block list) =
                            match blocks with
                            | Table _ :: ((Paragraph p | Heading (_, p)) :: _) when p.Id = para -> true
                            | _ :: rest -> look rest
                            | [] -> false
                        look doc.Blocks
                    if off > 0 || afterTable then split false
                    split true
                else piece.Append ch |> ignore
            flush (piece.ToString())
            story <- None
            this.Story.Total - totalBefore

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
            if seg.Mark = RowMark && seg.Start >= start && seg.Start < finish then
                // (An end-of-row mark has formatting of its own: over what its table was given.)
                let key = sprintf "table:%s:%d" seg.Table seg.Row
                let inherited = match markStyles.TryGetValue ("table:" + seg.Table) with | true, f -> f | _ -> id
                markStyles.[key] <- (match markStyles.TryGetValue key with | true, f -> f >> change | _ -> inherited >> change)
            if seg.Para <> "" && finish > start then
                let a = max start seg.TextStart - seg.TextStart
                let b = min finish (seg.TextStart + seg.Length) - seg.TextStart
                if b > a then this.UpdateParagraph(seg.Para, fun p -> { p with Buffer = TB.applyStyle a (b - a) change p.Buffer })
                let mark = seg.TextStart + seg.Length
                if mark >= start && mark < finish then
                    markStyles.[seg.Para] <- (match markStyles.TryGetValue seg.Para with | true, f -> f >> change | _ -> change)

    /// The characters of [start, end) mapped (Range.Case): the formatting stays.
    member this.MapText(start: int, finish: int, map: string -> int -> char) =
        for seg in this.SegmentsIn(start, finish) do
            if seg.Para <> "" then
                let a = max start seg.TextStart - seg.TextStart
                let b = min finish (seg.TextStart + seg.Length) - seg.TextStart
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
            if seg.Para = "" && seg.Mark = RowMark then
                match (if para = "" then None else findParagraph doc para) with
                | Some p ->
                    let mark =
                        match markStyles.TryGetValue (sprintf "table:%s:%d" seg.Table seg.Row) with
                        | true, f -> f TB.defaultStyle
                        | _ -> match markStyles.TryGetValue ("table:" + seg.Table) with | true, f -> f TB.defaultStyle | _ -> TB.defaultStyle
                    result.Add(resolve { p with Metadata = { p.Metadata with Style = None } } mark)
                | None -> ()
            else
            match (if seg.Para = "" then None else findParagraph doc seg.Para) with
            | Some p ->
                let a = max start seg.TextStart - seg.TextStart
                let b = min finish (seg.TextStart + seg.Length) - seg.TextStart
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
                        ListStyle = NoList; ListLevel = 0
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
//
// Each type of object (Range, Paragraph, Table, Selection, Document...) is a table of its members by name,
// built once (Rbt.Vb.HostMembers), and an object is its type's table with the state of one instance: the
// document and what of it (a live range, a table's id...). A name a type does not have is error 438.

/// Item(i) of a collection: 1-based.
let private itemAt (typeName: string) (items: Value array) (args: Value array) : Value =
    match arg args 0 with
    | Some index ->
        let i = toInt index
        if i >= 1 && i <= items.Length then items.[i - 1] else noSuchMember ()
    | None -> unsupported typeName "Item"

/// A collection's members: Count, First, Last and Item(i) (`item` for an Item of its own: by name, say).
/// A collection has no default member: `obj(i)` is Item (the engine calls Item then), and the collection
/// assigned without Set stays the object.
let private collectionMembers (typeName: string) (items: 'T -> Value array) (item: ('T -> Value array -> Value) option) : Member<'T> list =
    [ get "Count" (fun s -> VLong (items s).Length)
      get "First" (fun s -> match items s with [||] -> noSuchMember () | all -> all.[0])
      get "Last" (fun s -> match items s with [||] -> noSuchMember () | all -> all.[all.Length - 1])
      method "Item" [ "Index" ] (match item with Some f -> f | None -> fun s args -> itemAt typeName (items s) args) ]

/// A type of collection: its items (For Each), its collection members and its own.
let private collectionType (typeName: string) (items: 'T -> Value array) (item: ('T -> Value array -> Value) option) (members: Member<'T> list) =
    HostMembers.Table<'T>(typeName, collectionMembers typeName items item @ members, items = (fun s -> Seq.ofArray (items s)))

/// A member of one type as a member of another, whose state gives the first's (Selection's Range members).
let private lift (view: 'S -> 'T) (m: Member<'T>) : Member<'S> =
    { Name = m.Name
      Parameters = m.Parameters
      Get = m.Get |> Option.map (fun read -> fun s -> read (view s))
      Set = m.Set |> Option.map (fun write -> fun s value -> write (view s) value)
      Call = m.Call |> Option.map (fun run -> fun s args -> run (view s) args) }

/// A type's own members, and those of `inherited` it does not have itself.
let private extending (own: Member<'T> list) (inherited: Member<'T> list) : Member<'T> list =
    let names = own |> List.map (fun m -> m.Name.ToLowerInvariant()) |> Set.ofList
    (inherited |> List.filter (fun m -> not (names.Contains (m.Name.ToLowerInvariant())))) @ own

/// An empty collection (Fields, Shapes, Revisions...: what the model does not have).
let private emptyCollectionType = collectionType "Collection" (fun () -> [||]) None []
let private emptyCollection () = emptyCollectionType.Wrap ()

// ---- character formatting ------------------------------------------------------------------------------------

let private underlineStyles = [ 1, None; 2, Some "words"; 3, Some "double"; 4, Some "dotted"; 6, Some "thick"; 7, Some "dash"; 9, Some "dotDash"; 10, Some "dotDotDash"; 11, Some "wave" ]

let private toggleOr (value: Value) (current: bool) =
    let n = match scalarOf value with VBoolean b -> (if b then -1 else 0) | other -> toInt other
    if n = wdToggle then not current else n <> 0

/// A Font property: what it reads of a run's formatting, and what setting it to a value does to a run's
/// formatting (the value is converted when the change is made, before it is applied).
type private FontProperty = { Name: string; Read: TB.TextStyle -> Value; Change: Value -> TB.TextStyle -> TB.TextStyle }

let private fontProperties : FontProperty list =
    let flag name (read: TB.TextStyle -> bool) (write: bool -> TB.TextStyle -> TB.TextStyle) : FontProperty =
        { Name = name; Read = (fun s -> VLong (if read s then -1 else 0)); Change = (fun value s -> write (toggleOr value (read s)) s) }
    let vertical name (position: string) : FontProperty =
        { Name = name
          Read = (fun s -> VLong (if s.VertAlign = Some position then -1 else 0))
          Change = (fun value ->
            let on = toBool value
            fun s -> { s with VertAlign = (if on then Some position elif s.VertAlign = Some position then None else s.VertAlign) }) }
    let family name : FontProperty =
        { Name = name
          Read = (fun s -> VString (defaultArg s.FontFamily "Calibri"))
          Change = (fun value -> let family = toText value in fun s -> { s with FontFamily = Some family }) }
    [ flag "Bold" (fun s -> s.Bold) TB.withExplicitBold
      flag "Italic" (fun s -> s.Italic) TB.withExplicitItalic
      { Name = "Underline"
        Read = (fun s ->
            if not s.Underline then VLong 0
            else VLong (underlineStyles |> List.tryPick (fun (k, v) -> if v = s.UnderlineStyle then Some k else None) |> Option.defaultValue 1))
        Change = (fun value ->
            let n = toInt value
            fun s ->
                if n = 0 then TB.withExplicitUnderline false { s with UnderlineStyle = None }
                else
                    let style = underlineStyles |> List.tryPick (fun (k, v) -> if k = n then Some v else None) |> Option.defaultValue None
                    TB.withExplicitUnderline true { s with UnderlineStyle = style }) }
      { Name = "Size"
        Read = (fun s -> VDouble (Math.Round(ptOfPx s.FontSize * 2.0) / 2.0))
        Change = (fun value -> let size = pxOfPt (toFloat value) in TB.withExplicitFontSize size) }
      family "Name"
      family "NameAscii"
      family "NameOther"
      { Name = "Color"; Read = (fun s -> colorValueOf s.Color); Change = (fun value -> let color = colorOfValue value in fun s -> { s with Color = color }) }
      { Name = "ColorIndex"; Read = (fun s -> VLong (indexOfColor s.Color)); Change = (fun value -> let color = colorOfIndex (toInt value) in fun s -> { s with Color = color }) }
      flag "StrikeThrough" (fun s -> s.Strike) TB.withExplicitStrike
      flag "DoubleStrikeThrough" (fun s -> s.DoubleStrike) TB.withExplicitDoubleStrike
      vertical "Superscript" "superscript"
      vertical "Subscript" "subscript"
      flag "AllCaps" (fun s -> s.AllCaps) TB.withExplicitAllCaps
      flag "SmallCaps" (fun s -> s.SmallCaps) TB.withExplicitSmallCaps
      flag "Hidden" (fun s -> s.Hidden) TB.withExplicitHidden
      { Name = "Spacing"; Read = (fun s -> VDouble (ptOfPx s.CharacterSpacing)); Change = (fun value -> let px = pxOfPt (toFloat value) in fun s -> { s with CharacterSpacing = px }) }
      { Name = "Position"; Read = (fun s -> VDouble s.Position); Change = (fun value -> let pt = toFloat value in fun s -> { s with Position = pt }) } ]

/// A Font property's key in a set of Font settings (Find's, Replacement's, what is typed at a caret).
let private fontKey (p: FontProperty) = p.Name.ToLowerInvariant()

let private fontPropertyOf : Dictionary<string, FontProperty> =
    let byKey = Dictionary<string, FontProperty>()
    for p in fontProperties do byKey.[fontKey p] <- p
    byKey

/// What a Font is of: a range's text, a set of settings (Find's, Replacement's), a style, or the
/// Selection's caret (what is typed next gets what is set; reading gives the caret's formatting).
type private FontTarget =
    | RangeFont of LiveRange
    | SpecFont of Dictionary<string, Value>
    | StyleFont of string
    | CaretFont of LiveRange

type private FontState = { Book: MacroDocument; Target: FontTarget }

/// The change a set of Font settings makes.
let private specChange (spec: Dictionary<string, Value>) : TB.TextStyle -> TB.TextStyle =
    let changes = [ for kv in spec -> fontPropertyOf.[kv.Key].Change kv.Value ]
    fun s -> changes |> List.fold (fun acc f -> f acc) s

let rec private readFont (font: FontState) (p: FontProperty) : Value =
    let book = font.Book
    match font.Target with
    | RangeFont r ->
        match book.StylesIn(r.Start, r.End) |> List.map p.Read |> List.distinct with
        | [ value ] -> value
        | _ -> VLong wdUndefined
    | SpecFont spec ->
        match spec.TryGetValue (fontKey p) with
        | true, v -> v
        | _ -> VLong wdUndefined
    | CaretFont r ->
        match book.Pending.TryGetValue (fontKey p) with
        | true, v -> v
        | _ -> readFont { font with Target = RangeFont r } p
    | StyleFont id ->
        let named = book.Document.Styles |> List.tryFind (fun s -> s.Id = id)
        let text = named |> Option.bind (fun s -> s.Text) |> Option.defaultValue TB.defaultStyle
        let text =
            match named with
            | Some s -> { text with Bold = defaultArg s.TextBold text.Bold; Italic = defaultArg s.TextItalic text.Italic }
            | None -> text
        p.Read text

let private writeFont (font: FontState) (p: FontProperty) (value: Value) =
    let book = font.Book
    let change = p.Change value
    match font.Target with
    | RangeFont r -> book.Format(r.Start, r.End, change)
    | SpecFont spec -> spec.[fontKey p] <- value
    | CaretFont _ -> book.Pending.[fontKey p] <- value
    | StyleFont id ->
        book.ChangeStyle(id, fun s ->
            let text = change (defaultArg s.Text TB.defaultStyle)
            { s with
                Text = Some text
                TextBold = (if p.Name = "Bold" then Some text.Bold else s.TextBold)
                TextItalic = (if p.Name = "Italic" then Some text.Italic else s.TextItalic) })

let private fontType =
    let rangeOnly (name: string) (change: TB.TextStyle -> TB.TextStyle) =
        method name [] (fun (font: FontState) _ ->
            match font.Target with
            | RangeFont r -> font.Book.Format(r.Start, r.End, change); VEmpty
            | _ -> unsupported "Font" name)
    HostMembers.Table<FontState>("Font",
        (fontProperties |> List.map (fun p -> prop p.Name (fun font -> readFont font p) (fun font value -> writeFont font p value)))
        @ [ method "Reset" [] (fun font _ ->
                match font.Target with
                | RangeFont r ->
                    font.Book.Format(r.Start, r.End, fun s ->
                        { TB.defaultStyle with Style = s.Style; Hyperlink = s.Hyperlink; BookmarkStarts = s.BookmarkStarts; BookmarkEnds = s.BookmarkEnds; SourceRun = s.SourceRun })
                    VEmpty
                | SpecFont spec -> spec.Clear(); VEmpty
                | _ -> unsupported "Font" "Reset")
            rangeOnly "Grow" (fun s -> TB.withExplicitFontSize (s.FontSize + pxOfPt 1.0) s)
            rangeOnly "Shrink" (fun s -> TB.withExplicitFontSize (max 1.0 (s.FontSize - pxOfPt 1.0)) s) ])

let private fontObject (book: MacroDocument) (target: FontTarget) = fontType.Wrap { Book = book; Target = target }

/// A Font property of a Range itself (Range.Bold).
let private rangeFontProperty (name: string) (range: 'T -> MacroDocument * LiveRange) : Member<'T> =
    let p = fontPropertyOf.[name.ToLowerInvariant()]
    let font (s: 'T) : FontState = let book, r = range s in { Book = book; Target = RangeFont r }
    prop p.Name (fun s -> readFont (font s) p) (fun s value -> writeFont (font s) p value)

// ---- paragraph formatting ------------------------------------------------------------------------------------

let private alignmentOf (n: int) = match n with 1 -> Center | 2 -> Right | 3 | 4 -> Justify | _ -> Left
let private alignmentValue (a: Alignment) = match a with Left -> 0 | Center -> 1 | Right -> 2 | Justify -> 3

let private withFlag flag on (m: BlockMetadata) = { m with Explicit = (if on then m.Explicit ||| flag else m.Explicit &&& ~~~flag) }

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

let private fromStyle (style: StyleFallback) (f: WordStyles.ParagraphStyle -> 'a option) = style.Fallback |> Option.bind f

/// The paragraph's line spacing taken from its style (it states none itself).
let private styleLineSpacing (style: StyleFallback) (m: BlockMetadata) =
    { m with LineSpacingRule = fromStyle style (fun s -> s.LineSpacingRule); LineSpacing = fromStyle style (fun s -> s.LineSpacing) }

let private lineSpacingRuleOf (style: StyleFallback) (m: BlockMetadata) =
    let m =
        if m.LineSpacingRule.IsNone && (fromStyle style (fun s -> s.LineSpacingRule)).IsSome && m.LineSpacing.IsNone then styleLineSpacing style m
        else m
    match m.LineSpacingRule, m.LineSpacing with
    | Some "exact", _ -> VLong 4
    | Some "atLeast", _ -> VLong 3
    | _, Some 1.5 -> VLong 1
    | _, Some 2.0 -> VLong 2
    | _, (Some 1.0 | None) -> VLong 0
    | _ -> VLong 5

let private lineSpacingOf (style: StyleFallback) (m: BlockMetadata) =
    let m = if m.LineSpacing.IsNone && (fromStyle style (fun s -> s.LineSpacing)).IsSome then styleLineSpacing style m else m
    match m.LineSpacingRule, m.LineSpacing with
    | (Some "exact" | Some "atLeast"), Some px -> VDouble (ptOfPx px)
    | _, Some multiple -> VDouble (multiple * 12.0)
    | _ -> VDouble 12.0

let private lineSpacingRuleChange (value: Value) : BlockMetadata -> BlockMetadata =
    let rule = toInt value
    fun m ->
        let multiple v = { m with LineSpacing = Some v; LineSpacingRule = Some "auto" } |> withFlag ParagraphExplicit.LineSpacing true
        let fixedAs (kind: string) =
            let current = m.LineSpacing |> Option.filter (fun _ -> m.LineSpacingRule = Some "atLeast" || m.LineSpacingRule = Some "exact")
            { m with LineSpacingRule = Some kind; LineSpacing = Some (defaultArg current (pxOfPt 12.0)) } |> withFlag ParagraphExplicit.LineSpacing true
        match rule with
        | 0 -> multiple 1.0
        | 1 -> multiple 1.5
        | 2 -> multiple 2.0
        | 3 -> fixedAs "atLeast"
        | 4 -> fixedAs "exact"
        | _ -> { m with LineSpacingRule = Some "auto" } |> withFlag ParagraphExplicit.LineSpacing true

/// A ParagraphFormat property: what it reads of a paragraph (with what its style gives it), what setting
/// it does to a paragraph and to a paragraph style (None: it cannot be set there).
type private ParagraphProperty = {
    Name: string
    Read: StyleFallback -> BlockMetadata -> Value
    Change: (Value -> BlockMetadata -> BlockMetadata) option
    StyleChange: (Value -> NamedStyle -> NamedStyle) option
}

let private paragraphProperties : ParagraphProperty list =
    let orStyle (own: float option) (style: StyleFallback) (f: WordStyles.ParagraphStyle -> float option) =
        own |> Option.orElse (fromStyle style f) |> Option.defaultValue 0.0
    let indent name (own: BlockMetadata -> float option) (ofStyle: WordStyles.ParagraphStyle -> float option) flag
               (change: float -> BlockMetadata -> BlockMetadata) (styleChange: float -> NamedStyle -> NamedStyle) : ParagraphProperty =
        { Name = name
          Read = (fun style m -> VDouble (ptOfPx (orStyle (own m) style ofStyle)))
          Change = Some (fun value -> let v = pxOfPt (toFloat value) in fun m -> change v m |> withFlag flag true)
          StyleChange = Some (fun value s -> styleChange (pxOfPt (toFloat value)) s) }
    let keep name (own: BlockMetadata -> bool option) (ofStyle: WordStyles.ParagraphStyle -> bool option)
             (change: bool -> BlockMetadata -> BlockMetadata) (styleChange: bool -> NamedStyle -> NamedStyle) : ParagraphProperty =
        { Name = name
          Read = (fun style m -> VBoolean (own m |> Option.orElse (fromStyle style ofStyle) |> Option.defaultValue false))
          Change = Some (fun value -> let v = toBool value in change v)
          StyleChange = Some (fun value s -> styleChange (toBool value) s) }
    let auto name (change: BlockMetadata -> BlockMetadata) : ParagraphProperty =
        { Name = name
          Read = (fun _ _ -> VBoolean false)
          Change = Some (fun value -> let v = toBool value in fun m -> if v then change m else m)
          StyleChange = None }
    [ { Name = "Alignment"
        Read = (fun style m -> VLong (alignmentValue (if m.ExplicitAlignment then m.Alignment else (fromStyle style (fun s -> s.Alignment) |> Option.defaultValue m.Alignment))))
        Change = Some (fun value -> let a = alignmentOf (toInt value) in fun m -> { m with Alignment = a; ExplicitAlignment = true })
        StyleChange = Some (fun value s -> { s with Alignment = Some (alignmentOf (toInt value)) }) }
      { Name = "SpaceBefore"
        Read = (fun style m -> VDouble (orStyle m.SpacingBefore style (fun s -> s.SpacingBefore)))
        Change = Some (fun value -> let v = toFloat value in fun m -> { m with SpacingBefore = Some v } |> withFlag ParagraphExplicit.SpacingBefore true)
        StyleChange = Some (fun value s -> { s with SpacingBefore = Some (toFloat value) }) }
      { Name = "SpaceAfter"
        Read = (fun style m -> VDouble (orStyle m.SpacingAfter style (fun s -> s.SpacingAfter)))
        Change = Some (fun value -> let v = toFloat value in fun m -> { m with SpacingAfter = Some v } |> withFlag ParagraphExplicit.SpacingAfter true)
        StyleChange = Some (fun value s -> { s with SpacingAfter = Some (toFloat value) }) }
      indent "LeftIndent" (fun m -> m.IndentLeft) (fun s -> s.IndentLeft) ParagraphExplicit.IndentLeft
          (fun v m -> { m with IndentLeft = Some v }) (fun v s -> { s with IndentLeft = Some v })
      indent "RightIndent" (fun m -> m.IndentRight) (fun s -> s.IndentRight) ParagraphExplicit.IndentRight
          (fun v m -> { m with IndentRight = Some v }) (fun v s -> { s with IndentRight = Some v })
      indent "FirstLineIndent" (fun m -> m.FirstLineIndent) (fun s -> s.FirstLineIndent) ParagraphExplicit.FirstLineIndent
          (fun v m -> { m with FirstLineIndent = Some v }) (fun v s -> { s with FirstLineIndent = Some v })
      { Name = "LineSpacingRule"
        Read = lineSpacingRuleOf
        Change = Some lineSpacingRuleChange
        // (A style's rule goes with its LineSpacing.)
        StyleChange = Some (fun _ s -> s) }
      { Name = "LineSpacing"
        Read = lineSpacingOf
        Change = Some (fun value ->
            let pt = toFloat value
            fun m ->
                match m.LineSpacingRule with
                | Some "exact" | Some "atLeast" -> { m with LineSpacing = Some (pxOfPt pt) } |> withFlag ParagraphExplicit.LineSpacing true
                | _ -> { m with LineSpacing = Some (pt / 12.0); LineSpacingRule = Some "auto" } |> withFlag ParagraphExplicit.LineSpacing true)
        StyleChange = Some (fun value s -> { s with LineSpacing = Some (toFloat value / 12.0); LineSpacingRule = Some "auto" }) }
      keep "KeepWithNext" (fun m -> m.KeepNext) (fun s -> s.KeepNext) (fun v m -> { m with KeepNext = Some v }) (fun v s -> { s with KeepNext = Some v })
      keep "KeepTogether" (fun m -> m.KeepLines) (fun s -> s.KeepLines) (fun v m -> { m with KeepLines = Some v }) (fun v s -> { s with KeepLines = Some v })
      { Name = "PageBreakBefore"
        Read = (fun _ m -> VBoolean m.PageBreakBefore)
        Change = Some (fun value -> let v = toBool value in fun m -> { m with PageBreakBefore = v })
        StyleChange = None }
      { Name = "WidowControl"
        Read = (fun _ m -> VBoolean (m.WidowControl |> Option.defaultValue true))
        Change = Some (fun value -> let v = toBool value in fun m -> { m with WidowControl = Some v })
        StyleChange = None }
      { Name = "OutlineLevel"; Read = (fun style _ -> VLong (match style.Level with Some l -> l | None -> 10)); Change = None; StyleChange = None }
      auto "SpaceBeforeAuto" (fun m -> { m with SpacingBefore = Some 14.0 } |> withFlag ParagraphExplicit.SpacingBefore true)
      auto "SpaceAfterAuto" (fun m -> { m with SpacingAfter = Some 14.0 } |> withFlag ParagraphExplicit.SpacingAfter true) ]

/// A ParagraphFormat property's key in a set of settings (Replacement's).
let private paragraphKey (p: ParagraphProperty) = p.Name.ToLowerInvariant()

let private paragraphPropertyOf : Dictionary<string, ParagraphProperty> =
    let byKey = Dictionary<string, ParagraphProperty>()
    for p in paragraphProperties do byKey.[paragraphKey p] <- p
    byKey

/// The change a set of ParagraphFormat settings makes.
let private paragraphSpecChange (spec: Dictionary<string, Value>) : BlockMetadata -> BlockMetadata =
    let changes = [ for kv in spec -> paragraphPropertyOf.[kv.Key].Change.Value kv.Value ]
    fun m -> changes |> List.fold (fun acc f -> f acc) m


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

type private ShadingState = { Book: MacroDocument; Target: ShadingTarget }

let private shadingColor (shading: ShadingState) =
    let book = shading.Book
    match shading.Target with
    | ParagraphShading ids ->
        match ids () |> List.tryHead with
        | Some id ->
            match book.Paragraph id |> Option.bind (fun p -> p.Metadata.BackgroundColor) with
            | Some color -> Some color
            | None ->
                // (A paragraph in a table cell shows the cell's shading, and Word reports it.)
                let st = book.Story
                match st.Segments |> Array.tryFind (fun seg -> seg.Para = id && seg.Table <> "") with
                | Some seg -> cellOf book { TableId = seg.Table; Row = seg.Row; Col = seg.Col } |> Option.bind (fun c -> c.BackgroundColor)
                | None -> None
        | None -> None
    | CellShading cells -> cells () |> List.tryHead |> Option.bind (cellOf book) |> Option.bind (fun c -> c.BackgroundColor)

let private setShadingColor (shading: ShadingState) (color: string option) =
    let book = shading.Book
    match shading.Target with
    | ParagraphShading ids ->
        for id in ids () do
            book.UpdateParagraph(id, fun p -> { p with Metadata = { p.Metadata with BackgroundColor = color } |> withFlag ParagraphExplicit.BackgroundColor color.IsSome })
    | CellShading cells -> mapCells book (cells ()) (fun c -> { c with BackgroundColor = color; Shading = None })

let private shadingType =
    HostMembers.Table<ShadingState>("Shading", [
        prop "BackgroundPatternColor" (fun s -> colorValueOf (shadingColor s)) (fun s value -> setShadingColor s (colorOfValue value))
        prop "BackgroundPatternColorIndex" (fun s -> VLong (indexOfColor (shadingColor s))) (fun s value -> setShadingColor s (colorOfIndex (toInt value)))
        accepted "ForegroundPatternColor" (VLong wdColorAutomatic)
        set "ForegroundPatternColorIndex" (fun _ _ -> ())
        accepted "Texture" (VLong 0) ])

let private shadingObject (book: MacroDocument) (target: ShadingTarget) = shadingType.Wrap { Book = book; Target = target }

let private lineStyleNames =
    [ 0, "none"; 1, "single"; 2, "dotted"; 3, "dashSmallGap"; 4, "dashed"; 5, "dotDash"; 6, "dotDotDash"; 7, "double"
      8, "triple"; 9, "thinThickSmallGap"; 10, "thickThinSmallGap"; 18, "wave"; 19, "doubleWave" ]
let private lineStyleName (n: int) = lineStyleNames |> List.tryPick (fun (k, v) -> if k = n then Some v else None) |> Option.defaultValue "single"
let private lineStyleValue (style: string option) =
    match style with
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

/// A colour set on a border: black for automatic.
let private lineColor (color: string option) = match color with Some _ -> color | None -> Some "#000000"

type private BorderState = { Book: MacroDocument; Target: BordersTarget; Side: BorderSide }

let private borderType =
    let line (b: BorderState) = readLine b.Book b.Target b.Side
    let update (b: BorderState) change = writeLine b.Book b.Target b.Side change
    HostMembers.Table<BorderState>("Border", [
        prop "LineStyle"
            (fun b -> VLong (match line b with Some l -> lineStyleValue (Some l.LineStyle) | None -> 0))
            (fun b value -> let n = toInt value in update b (fun l -> if n = 0 then None else Some { (defaultArg l defaultLine) with LineStyle = lineStyleName n }))
        prop "LineWidth"
            (fun b -> VLong (match line b with Some l -> int (Math.Round(l.LineWidth * 6.0)) | None -> 4))
            (fun b value -> let px = float (toInt value) / 6.0 in update b (fun l -> Some { (defaultArg l defaultLine) with LineWidth = px }))
        prop "Color"
            (fun b -> colorValueOf (line b |> Option.bind (fun l -> l.LineColor)))
            (fun b value -> let c = lineColor (colorOfValue value) in update b (fun l -> Some { (defaultArg l defaultLine) with LineColor = c }))
        prop "ColorIndex"
            (fun b -> VLong (indexOfColor (line b |> Option.bind (fun l -> l.LineColor))))
            (fun b value -> let c = lineColor (colorOfIndex (toInt value)) in update b (fun l -> Some { (defaultArg l defaultLine) with LineColor = c }))
        prop "Visible"
            (fun b -> VBoolean (line b).IsSome)
            (fun b value -> let on = toBool value in update b (fun l -> if on then Some (defaultArg l defaultLine) else None)) ])

let private borderObject (book: MacroDocument) (target: BordersTarget) (side: BorderSide) = borderType.Wrap { Book = book; Target = target; Side = side }

type private BordersState = { Book: MacroDocument; Target: BordersTarget }

let private bordersType =
    let sides (b: BordersState) =
        match b.Target with
        | ParagraphBorders _ -> [ SideTop; SideLeft; SideBottom; SideRight ]
        | _ -> [ SideTop; SideLeft; SideBottom; SideRight; SideInsideH; SideInsideV ]
    let outside (_: BordersState) = [ SideTop; SideLeft; SideBottom; SideRight ]
    let inside (b: BordersState) = sides b |> List.filter (fun s -> s = SideInsideH || s = SideInsideV)
    let setAll (b: BordersState) (which: BorderSide list) change = for side in which do writeLine b.Book b.Target side change
    let lineStyleOf (b: BordersState) (side: BorderSide option) =
        VLong (match side |> Option.bind (readLine b.Book b.Target) with Some l -> lineStyleValue (Some l.LineStyle) | None -> 0)
    let setLineStyle (which: BordersState -> BorderSide list) (b: BordersState) (value: Value) =
        let n = toInt value
        setAll b (which b) (fun line -> if n = 0 then None else Some { (defaultArg line defaultLine) with LineStyle = lineStyleName n })
    let setLineWidth (which: BordersState -> BorderSide list) (b: BordersState) (value: Value) =
        let px = float (toInt value) / 6.0
        setAll b (which b) (fun line -> Some { (defaultArg line defaultLine) with LineWidth = px })
    let setColor (which: BordersState -> BorderSide list) (color: Value -> string option) (b: BordersState) (value: Value) =
        let c = lineColor (color value)
        setAll b (which b) (fun line -> Some { (defaultArg line defaultLine) with LineColor = c })
    let colorIndex (value: Value) = colorOfIndex (toInt value)
    collectionType "Borders"
        (fun b -> sides b |> List.map (borderObject b.Book b.Target) |> List.toArray)
        // (Borders(wdBorderTop) ...; any other index, or none, is the top border.)
        (Some (fun b args -> borderObject b.Book b.Target (arg args 0 |> Option.map toInt |> Option.bind sideOf |> Option.defaultValue SideTop)))
        [ prop "Enable"
            (fun b -> VLong (if sides b |> List.exists (fun s -> (readLine b.Book b.Target s).IsSome) then -1 else 0))
            (fun b value -> let on = toBool value in setAll b (sides b) (fun line -> if on then Some (defaultArg line defaultLine) else None))
          prop "InsideLineStyle" (fun b -> lineStyleOf b (inside b |> List.tryHead)) (setLineStyle inside)
          prop "OutsideLineStyle" (fun b -> lineStyleOf b (Some SideTop)) (setLineStyle outside)
          set "InsideLineWidth" (setLineWidth inside)
          set "OutsideLineWidth" (setLineWidth outside)
          set "InsideColor" (setColor inside colorOfValue)
          set "OutsideColor" (setColor outside colorOfValue)
          set "InsideColorIndex" (setColor inside colorIndex)
          set "OutsideColorIndex" (setColor outside colorIndex)
          accepted "Shadow" (VBoolean false)
          set "DistanceFromTop" (fun _ _ -> ())
          set "DistanceFromBottom" (fun _ _ -> ())
          set "DistanceFromLeft" (fun _ _ -> ())
          set "DistanceFromRight" (fun _ _ -> ()) ]

let private bordersObject (book: MacroDocument) (target: BordersTarget) = bordersType.Wrap { Book = book; Target = target }

// ---- styles of a range -----------------------------------------------------------------------------------------

let private setRangeStyle (book: MacroDocument) (r: LiveRange) (value: Value) =
    let styleName =
        match scalarOf value with
        | VInteger n | VLong n -> (match builtInStyleName n with Some s -> s | None -> fail 5834 "Item with specified name does not exist.")
        | VDouble d -> (match builtInStyleName (int d) with Some s -> s | None -> fail 5834 "Item with specified name does not exist.")
        | other -> toText other
    if eq styleName "Normal" && not (book.Document.Styles |> List.exists (fun s -> s.Id = "Normal")) then
        for id in book.ParagraphIdsIn(r.Start, r.End) do book.ApplyParagraphStyle(id, None, None)
    else
        match book.FindStyle styleName with
        | Some style when style.StyleType = "character" ->
            book.Format(r.Start, r.End, fun s -> { s with Style = Some style.Id })
        | Some style ->
            let level = headingLevelOf style
            for id in book.ParagraphIdsIn(r.Start, r.End) do book.ApplyParagraphStyle(id, Some style.Id, level)
        | None ->
            if eq styleName "List Bullet" || eq styleName "List Number" then
                for id in book.ParagraphIdsIn(r.Start, r.End) do
                    book.UpdateParagraph(id, fun p -> { p with Metadata = { p.Metadata with ListStyle = (if eq styleName "List Bullet" then Bullet else Numbered); ListLevel = max 1 p.Metadata.ListLevel } })
            else fail 5834 (sprintf "Item with specified name does not exist (style \"%s\")." styleName)

// ---- ParagraphFormat ---------------------------------------------------------------------------------------------

/// The bounds of what an object is of: a range, or the whole story (None).
let private boundsOf (book: MacroDocument) (within: LiveRange option) =
    match within with
    | Some r -> r.Start, r.End
    | None -> 0, book.Story.Total

/// What a ParagraphFormat is of: the paragraphs of a range (None: of the document), a set of settings
/// (Replacement's), or a paragraph style.
type private ParagraphTarget =
    | ParagraphsOf of LiveRange option
    | SpecParagraphs of Dictionary<string, Value>
    | StyleParagraphs of string

type private ParagraphFormatState = { Book: MacroDocument; Target: ParagraphTarget }

let private paragraphIds (book: MacroDocument) (within: LiveRange option) =
    let s, e = boundsOf book within
    book.ParagraphIdsIn(s, e)

let private namedStyleOf (book: MacroDocument) (id: string) =
    book.Paragraph id |> Option.bind (fun p -> p.Metadata.Style) |> Option.bind (fun s -> book.Document.Styles |> List.tryFind (fun x -> x.Id = s))

let private paragraphFallback (book: MacroDocument) (id: string) : StyleFallback =
    let doc = book.Document
    let resolved =
        match doc.Blocks |> List.tryFind (function Paragraph p | Heading (_, p) -> p.Id = id | _ -> false) with
        | Some block -> Some (WordStyles.forBlockInDocument doc block)
        | None -> book.Paragraph id |> Option.map (WordStyles.forTableCellParagraphInDocument doc)
    { Fallback = resolved; Level = namedStyleOf book id |> Option.bind headingLevelOf }

/// A style for each of the paragraphs.
let private setParagraphStyles (book: MacroDocument) (within: LiveRange option) (value: Value) =
    for id in paragraphIds book within do
        let pos = book.PositionOf(id, 0)
        setRangeStyle book (LiveRange(pos, pos)) value

let private readParagraphFormat (format: ParagraphFormatState) (p: ParagraphProperty) : Value =
    let book = format.Book
    match format.Target with
    | ParagraphsOf within ->
        let values = paragraphIds book within |> List.choose (fun id -> book.Paragraph id |> Option.map (fun para -> p.Read (paragraphFallback book id) para.Metadata))
        match values |> List.distinct with
        | [ v ] -> v
        | [] ->
            // (An end-of-row mark: the default paragraph style's format.)
            let normal = WordStyles.forBlockInDocument book.Document (Paragraph (newEmptyParagraph ()))
            p.Read { Fallback = Some normal; Level = None } defaultMetadata
        | _ -> VLong wdUndefined
    | SpecParagraphs spec ->
        match spec.TryGetValue (paragraphKey p) with
        | true, v -> v
        | _ -> VLong wdUndefined
    | StyleParagraphs id ->
        let s = book.Document.Styles |> List.tryFind (fun x -> x.Id = id)
        p.Read (match s with Some x -> fallbackOfNamed x | None -> noFallback) defaultMetadata

let private writeParagraphFormat (format: ParagraphFormatState) (p: ParagraphProperty) (value: Value) =
    let book = format.Book
    match format.Target, p.Change, p.StyleChange with
    | ParagraphsOf within, Some change, _ ->
        let change = change value
        for id in paragraphIds book within do
            book.UpdateParagraph(id, fun para -> let m = change para.Metadata in if m = para.Metadata then para else { para with Metadata = m })
    | SpecParagraphs spec, Some change, _ ->
        change value |> ignore
        spec.[paragraphKey p] <- value
    | StyleParagraphs id, _, Some change -> book.ChangeStyle(id, change value)
    | _ -> unsupported "ParagraphFormat" p.Name

let private paragraphFormatType =
    let ofParagraphs (name: string) (f: MacroDocument -> LiveRange option -> Value) (format: ParagraphFormatState) =
        match format.Target with
        | ParagraphsOf within -> f format.Book within
        | _ -> unsupported "ParagraphFormat" name
    HostMembers.Table<ParagraphFormatState>("ParagraphFormat",
        (paragraphProperties |> List.map (fun p -> prop p.Name (fun format -> readParagraphFormat format p) (fun format value -> writeParagraphFormat format p value)))
        @ [ get "Shading" (ofParagraphs "Shading" (fun book within -> shadingObject book (ParagraphShading (fun () -> paragraphIds book within))))
            get "Borders" (ofParagraphs "Borders" (fun book within -> bordersObject book (ParagraphBorders (fun () -> paragraphIds book within))))
            prop "Style"
                (ofParagraphs "Style" (fun book within ->
                    match paragraphIds book within with
                    | id :: _ -> VString (match namedStyleOf book id with Some s -> styleDisplayName s | None -> "Normal")
                    | [] -> VEmpty))
                (fun format value ->
                    match format.Target with
                    | ParagraphsOf within -> setParagraphStyles format.Book within value
                    | _ -> unsupported "ParagraphFormat" "Style") ])

let private paragraphFormatObject (book: MacroDocument) (target: ParagraphTarget) = paragraphFormatType.Wrap { Book = book; Target = target }

/// The ParagraphFormat properties of an object of paragraphs itself (Paragraph.Alignment), read and set
/// through its ParagraphFormat.
let private paragraphFormatProperties (format: 'T -> ParagraphFormatState) : Member<'T> list =
    paragraphProperties |> List.map (fun p -> prop p.Name (fun s -> readParagraphFormat (format s) p) (fun s value -> writeParagraphFormat (format s) p value))

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
    let plainGrid = lower = "table grid" || lower = "tablegrid"
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
        // (Word's table styles set their paragraphs' spacing: none after, single lines.)
        let ts, te = tableSpan book tableId
        book.MapParagraphs(ts, te, fun m ->
            { m with SpacingAfter = Some 0.0; LineSpacing = Some 1.0; LineSpacingRule = Some "auto" }
            |> withFlag ParagraphExplicit.SpacingAfter true |> withFlag ParagraphExplicit.LineSpacing true)
        // The header row's text bold (white on a dark fill), and the first column's (Word's styles have
        // "Header Row" and "First Column" on); Table Grid has neither. (Cell by cell: the end-of-row marks
        // are not of the style.)
        if not plainGrid then
            match tableOf book tableId with
            | Some t ->
                t.Rows |> List.iteri (fun r row ->
                    for c in 0 .. row.Cells.Length - 1 do
                        if r = 0 || c = 0 then
                            let cs, ce = cellSpan book { TableId = tableId; Row = r; Col = c }
                            book.Format(cs, ce, fun st ->
                                if r = 0 then { TB.withExplicitBold true st with Color = (match headerText with Some c -> Some c | None -> st.Color) }
                                else TB.withExplicitBold true st))
            | None -> ()
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
                let p = fontPropertyOf.[kv.Key]
                let wanted = p.Read (p.Change kv.Value TB.defaultStyle)
                styles |> List.forall (fun st -> p.Read st = wanted))
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
        book.MapParagraphs(s, e, paragraphSpecChange state.ReplacementParagraph)

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

// ---- units of text -----------------------------------------------------------------------------------------

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
    let wordCount = text.Split([| ' '; '\r'; '\a'; '\b'; '\t'; '\n'; '\f' |], StringSplitOptions.RemoveEmptyEntries) |> Array.length
    match kind with
    | 0 -> wordCount
    | 1 | 4 -> book.SegmentsIn(s, e) |> List.filter (fun seg -> seg.Mark <> RowMark && (seg.Length > 0 || seg.Mark = ParagraphMark)) |> List.length
    | 2 -> 1 + (book.SegmentsIn(s, e) |> List.filter (fun seg -> seg.Para <> "" && (book.Paragraph seg.Para |> Option.exists (fun p -> p.Metadata.PageBreakBefore))) |> List.length)
    | 3 -> text |> Seq.filter (fun c -> not (Char.IsWhiteSpace c) && c <> '\a' && c <> '\b') |> Seq.length
    | 5 -> text |> Seq.filter (fun c -> c <> '\r' && c <> '\a' && c <> '\b') |> Seq.length
    | _ -> 0

// ---- the objects' state --------------------------------------------------------------------------------------

type private RangeState = { Book: MacroDocument; R: LiveRange }
/// A collection of the document's text (Paragraphs, Words, Tables, Bookmarks): of a range, or of the
/// whole story (None).
type private WithinState = { Book: MacroDocument; Within: LiveRange option }
type private TableState = { Book: MacroDocument; TableId: string }
type private RowState = { Book: MacroDocument; TableId: string; Row: int }
type private ColumnState = { Book: MacroDocument; TableId: string; Col: int }
type private CellState = { Book: MacroDocument; Cell: CellRef }
type private CellSet =
    | CellsInRange of LiveRange
    | CellsOfRow of string * int
    | CellsOfColumn of string * int
type private CellsState = { Book: MacroDocument; Cells: CellSet }
type private FindObjectState = { Book: MacroDocument; R: LiveRange; IsSelection: bool; Find: FindState }
type private ReplacementState = { Book: MacroDocument; Find: FindState }
type private BookmarkState = { Book: MacroDocument; Info: BookmarkInfo; S: int; E: int }
type private StyleState = { Book: MacroDocument; Id: string }
type private VariableState = { Book: MacroDocument; Key: string }
/// The document, or its window (TypeName(ActiveDocument) is "Document", as Word's).
type private NamedState = { Book: MacroDocument }
type private ConversionState = { Conversion: string; Factor: float }

// ---- styles ----------------------------------------------------------------------------------------------------

let private styleType =
    let style (s: StyleState) = s.Book.Document.Styles |> List.tryFind (fun x -> x.Id = s.Id)
    let name (s: StyleState) = VString (match style s with Some x -> styleDisplayName x | None -> s.Id)
    HostMembers.Table<StyleState>("Style",
        [ get "NameLocal" name
          get "Name" name
          get "Font" (fun s -> fontObject s.Book (StyleFont s.Id))
          get "ParagraphFormat" (fun s -> paragraphFormatObject s.Book (StyleParagraphs s.Id))
          get "Type" (fun s -> VLong (match style s |> Option.map (fun x -> x.StyleType) with Some "character" -> 2 | Some "table" -> 3 | _ -> 1))
          prop "BaseStyle"
            (fun s -> VString (style s |> Option.bind (fun x -> x.BasedOn) |> Option.defaultValue ""))
            (fun s value -> s.Book.ChangeStyle(s.Id, fun x -> { x with BasedOn = Some (toText value |> compact) }))
          set "NextParagraphStyle" (fun s value -> s.Book.ChangeStyle(s.Id, fun x -> { x with Next = Some (toText value |> compact) }))
          get "BuiltIn" (fun _ -> VBoolean true)
          get "InUse" (fun _ -> VBoolean true)
          prop "QuickStyle" (fun s -> VBoolean (style s |> Option.exists (fun x -> x.IsQuickStyle))) (fun _ _ -> ())
          set "AutomaticallyUpdate" (fun _ _ -> ())
          set "NoFollowUp" (fun _ _ -> ()) ],
        defaultMember = "NameLocal")

let private styleObject (book: MacroDocument) (id: string) = styleType.Wrap { Book = book; Id = id }

let private stylesType =
    let lookup (book: MacroDocument) (key: Value) =
        let styleName =
            match scalarOf key with
            | VInteger n | VLong n -> (match builtInStyleName n with Some s -> s | None -> fail 5941 "The requested member of the collection does not exist.")
            | VDouble d -> (match builtInStyleName (int d) with Some s -> s | None -> fail 5941 "The requested member of the collection does not exist.")
            | v -> toText v
        if eq styleName "Normal" && not (book.Document.Styles |> List.exists (fun s -> s.Id = "Normal")) then
            // (A new document's default paragraph style: Normal, as the document defaults make it.)
            let normal = { plainStyle with Id = "Normal"; Name = Some "Normal"; IsDefault = true; BasedOn = None; Next = None; UiPriority = Some 0 }
            book.SetFormatting { book.Document with Styles = book.Document.Styles @ [ normal ] }
        match book.FindStyle styleName with
        | Some s -> styleObject book s.Id
        | None -> fail 5941 (sprintf "The requested member of the collection does not exist (style \"%s\")." styleName)
    collectionType "Styles"
        (fun (book: MacroDocument) -> book.Document.Styles |> List.map (fun s -> styleObject book s.Id) |> List.toArray)
        (Some (fun book args -> match arg args 0 with Some key -> lookup book key | None -> unsupported "Styles" "Item"))
        [ method "Add" [ "Name"; "Type" ] (fun book args ->
            let styleName = toText (defaultArg (arg args 0) (VString "Style1"))
            let kind = match intArg args 1 1 with 2 -> "character" | 3 -> "table" | _ -> "paragraph"
            let id = compact styleName
            if not (book.Document.Styles |> List.exists (fun s -> s.Id = id)) then
                book.SetFormatting { book.Document with Styles = book.Document.Styles @ [ { plainStyle with Id = id; Name = Some styleName; StyleType = kind } ] }
            styleObject book id) ]

let private styleOfRange (book: MacroDocument) (r: LiveRange) : Value =
    let ids = book.ParagraphIdsIn(r.Start, r.End) |> List.map (fun id -> book.Paragraph id |> Option.bind (fun p -> p.Metadata.Style)) |> List.distinct
    match ids with
    | [ Some id ] ->
        match book.Document.Styles |> List.tryFind (fun s -> s.Id = id) with
        | Some s -> styleObject book s.Id
        | None -> VString id
    | [ None ] | [] -> VString "Normal"
    | _ -> VNothing

// ---- lists -------------------------------------------------------------------------------------------------------

let private listFormatType =
    let paragraphs (x: RangeState) = x.Book.ParagraphIdsIn(x.R.Start, x.R.End)
    let first (x: RangeState) = paragraphs x |> List.tryHead |> Option.bind x.Book.Paragraph
    /// A list style on the paragraphs (off where they have it already), or none.
    let apply (x: RangeState) (style: ListStyle option) =
        for id in paragraphs x do
            x.Book.UpdateParagraph(id, fun p ->
                let m = p.Metadata
                match style with
                | Some s -> { p with Metadata = { m with ListStyle = (if m.ListStyle = s then NoList else s); ListLevel = (if m.ListStyle = s then 0 else max 1 m.ListLevel) } }
                | None -> { p with Metadata = { m with ListStyle = NoList; ListLevel = 0 } })
    let applyDefault (style: ListStyle) (x: RangeState) (_: Value array) =
        // (Word gives list paragraphs the List Paragraph style.)
        let listParagraph = x.Book.FindStyle "List Paragraph"
        let ids = paragraphs x
        let removing = ids |> List.forall (fun id -> x.Book.Paragraph id |> Option.exists (fun p -> p.Metadata.ListStyle <> NoList))
        if not removing then
            for id in ids do
                match x.Book.Paragraph id |> Option.bind (fun p -> p.Metadata.Style) with
                | None | Some "Normal" -> x.Book.ApplyParagraphStyle(id, listParagraph |> Option.map (fun s -> s.Id), None)
                | Some _ -> ()
        apply x (Some style)
        VEmpty
    let shift (step: int) (x: RangeState) (_: Value array) =
        for id in paragraphs x do
            x.Book.UpdateParagraph(id, fun p ->
                if p.Metadata.ListStyle = NoList then p
                else { p with Metadata = { p.Metadata with ListLevel = max 1 (min 9 (p.Metadata.ListLevel + step)) } })
        VEmpty
    HostMembers.Table<RangeState>("ListFormat", [
        get "ListType" (fun x -> VLong (match first x with Some p -> (match p.Metadata.ListStyle with Bullet -> 2 | Numbered -> 3 | NoList -> 0) | None -> 0))
        prop "ListLevelNumber"
            (fun x -> VLong (match first x with Some p -> max 1 p.Metadata.ListLevel | None -> 1))
            (fun x value ->
                let level = max 1 (min 9 (toInt value))
                for id in paragraphs x do
                    x.Book.UpdateParagraph(id, fun p -> if p.Metadata.ListStyle = NoList then p else { p with Metadata = { p.Metadata with ListLevel = level } }))
        method "ApplyBulletDefault" [] (applyDefault Bullet)
        method "ApplyNumberDefault" [] (applyDefault Numbered)
        method "RemoveNumbers" [] (fun x _ -> apply x None; VEmpty)
        method "ListIndent" [] (shift 1)
        method "ListOutdent" [] (shift -1) ])

// ---- Find's Replacement -------------------------------------------------------------------------------------------

let private replacementType =
    HostMembers.Table<ReplacementState>("Replacement", [
        prop "Text" (fun x -> VString x.Find.ReplacementText) (fun x value -> x.Find.ReplacementText <- toText value)
        get "Font" (fun x -> fontObject x.Book (SpecFont x.Find.ReplacementFont))
        prop "Highlight"
            (fun x -> VLong (if x.Find.ReplacementHighlight < 0 then wdUndefined else x.Find.ReplacementHighlight))
            (fun x value -> x.Find.ReplacementHighlight <- (if toBool value then 1 else 0))
        prop "Style"
            (fun x -> VString (defaultArg x.Find.ReplacementStyle ""))
            (fun x value ->
                x.Find.ReplacementStyle <-
                    (match scalarOf value with
                     | VInteger n | VLong n -> builtInStyleName n
                     | v -> Some (toText v)))
        get "ParagraphFormat" (fun x -> paragraphFormatObject x.Book (SpecParagraphs x.Find.ReplacementParagraph))
        method "ClearFormatting" [] (fun x _ ->
            let f = x.Find
            f.ReplacementFont.Clear(); f.ReplacementHighlight <- -1; f.ReplacementStyle <- None; f.ReplacementParagraph.Clear()
            VEmpty) ])

// ---- the document's page setup, variables; the window, options ---------------------------------------------------

let private pageSetupType =
    let change (book: MacroDocument) (f: PageSetup -> PageSetup) = book.SetStructure { book.Document with PageSetup = f book.Document.PageSetup }
    let length name (read: PageSetup -> float) (write: float -> PageSetup -> PageSetup) =
        prop name (fun (book: MacroDocument) -> VDouble (ptOfPx (read book.Document.PageSetup))) (fun book value -> let v = pxOfPt (toFloat value) in change book (write v))
    HostMembers.Table<MacroDocument>("PageSetup", [
        length "TopMargin" (fun p -> p.MarginTop) (fun v p -> { p with MarginTop = v })
        length "BottomMargin" (fun p -> p.MarginBottom) (fun v p -> { p with MarginBottom = v })
        length "LeftMargin" (fun p -> p.MarginLeft) (fun v p -> { p with MarginLeft = v })
        length "RightMargin" (fun p -> p.MarginRight) (fun v p -> { p with MarginRight = v })
        length "PageWidth" (fun p -> p.PageWidth) (fun v p -> { p with PageWidth = v })
        length "PageHeight" (fun p -> p.PageHeight) (fun v p -> { p with PageHeight = v })
        length "HeaderDistance" (fun p -> p.HeaderDistance) (fun v p -> { p with HeaderDistance = v })
        length "FooterDistance" (fun p -> p.FooterDistance) (fun v p -> { p with FooterDistance = v })
        prop "Orientation"
            (fun book -> let p = book.Document.PageSetup in VLong (if p.PageWidth > p.PageHeight then 1 else 0))
            (fun book value ->
                let landscape = toInt value = 1
                change book (fun p -> if (p.PageWidth > p.PageHeight) = landscape then p else { p with PageWidth = p.PageHeight; PageHeight = p.PageWidth })) ])

/// Document.Variables: values kept by name for the run (Add, Item(name).Value, Count).
let private variableNames (book: MacroDocument) = [ for kv in book.Settings do if kv.Key.StartsWith "var:" then yield kv.Key.Substring 4 ]

let private variableType =
    HostMembers.Table<VariableState>("Variable",
        [ prop "Value"
            (fun v -> match v.Book.Settings.TryGetValue ("var:" + v.Key) with | true, value -> value | _ -> VEmpty)
            (fun v value -> v.Book.Settings.["var:" + v.Key] <- VString (toText value))
          get "Name" (fun v -> VString v.Key)
          method "Delete" [] (fun v _ -> v.Book.Settings.Remove("var:" + v.Key) |> ignore; VEmpty) ],
        defaultMember = "Value")

let private variablesType =
    let variable (book: MacroDocument) (key: string) = variableType.Wrap { Book = book; Key = key }
    let find (book: MacroDocument) (key: string) = variableNames book |> List.tryFind (eq key)
    let items (book: MacroDocument) = variableNames book |> List.map (variable book) |> List.toArray
    collectionType "Variables" items
        (Some (fun book args ->
            match arg args 0 with
            | Some (VString key) -> (match find book key with Some k -> variable book k | None -> fail 5825 "Object has been deleted.")
            | _ -> itemAt "Variables" (items book) args))
        [ method "Add" [ "Name"; "Value" ] (fun book args ->
            let key = toText (defaultArg (arg args 0) (VString ""))
            if (find book key).IsSome then fail 4172 "The variable name already exists."
            book.Settings.["var:" + key] <- VString (toText (defaultArg (arg args 1) (VString "")))
            variable book key) ]

let private zoomType =
    // (Any setting of the zoom is accepted: the host has its own.)
    HostMembers.Table<unit>("Zoom", [ accepted "Percentage" (VLong 100) ],
                            dynamic = ({ Read = (fun _ _ _ -> None); Write = (fun _ _ _ -> true) } : Dynamic<unit>))

let private viewType =
    let noEffect name = set name (fun (_: unit) _ -> ())
    HostMembers.Table<unit>("View", [
        accepted "Type" (VLong 3)
        accepted "ShowAll" (VBoolean false)
        accepted "ShowHiddenText" (VBoolean false)
        accepted "ShowFieldCodes" (VBoolean false)
        noEffect "ShowBookmarks"
        noEffect "ShowParagraphs"
        noEffect "Draft"
        get "Zoom" (fun () -> zoomType.Wrap ()) ])

let private paneType =
    HostMembers.Table<unit>("Pane", [
        get "View" (fun () -> viewType.Wrap ())
        method "SmallScroll" [] (fun _ _ -> VEmpty)
        method "LargeScroll" [] (fun _ _ -> VEmpty) ])

let private optionsType =
    let option name (fallback: Value) =
        let key = (name: string).ToLowerInvariant()
        prop name
            (fun (book: MacroDocument) -> match book.Settings.TryGetValue key with | true, v -> v | _ -> fallback)
            (fun book value -> book.Settings.[key] <- value)
    HostMembers.Table<MacroDocument>("Options", [
        option "DefaultHighlightColorIndex" (VLong 7)
        option "ReplaceSelection" (VBoolean true)
        option "CheckSpellingAsYouType" (VBoolean true)
        option "CheckGrammarAsYouType" (VBoolean true)
        option "AutoFormatAsYouTypeReplaceQuotes" (VBoolean true)
        option "SmartCutPaste" (VBoolean true)
        option "Pagination" (VBoolean true)
        option "MeasurementUnit" (VLong 0)
        option "AutoFormatAsYouTypeApplyNumberedLists" (VBoolean true) ])

/// The unit conversions (Application.InchesToPoints, and the globals of those names).
let private conversions =
    [ "CentimetersToPoints", 72.0 / 2.54; "MillimetersToPoints", 72.0 / 25.4; "InchesToPoints", 72.0; "LinesToPoints", 12.0
      "PicasToPoints", 12.0; "PointsToCentimeters", 2.54 / 72.0; "PointsToInches", 1.0 / 72.0; "PointsToMillimeters", 25.4 / 72.0 ]
let private convert (factor: float) (args: Value array) = VDouble (toFloat (defaultArg (arg args 0) (VLong 0)) * factor)

let private conversionType =
    HostMembers.Table<ConversionState>("Conversion", [ method "Convert" [] (fun c args -> convert c.Factor args) ],
                                       defaultMember = "Convert", typeNameOf = (fun c -> c.Conversion))

let private doEventsType =
    HostMembers.Table<unit>("DoEvents", [ method "DoEvents" [] (fun _ _ -> VLong 0) ], defaultMember = "DoEvents")


// ---- the types of object that refer to one another -----------------------------------------------------------
//
// A Range has Paragraphs and a Paragraph a Range, a Table has Rows and a Row a Range...: these types are
// reached through their references here, each set where its table is built below (once, as the module
// is loaded, before any macro runs).

let mutable private rangeType : HostMembers.Table<RangeState> = null
let mutable private selectionType : HostMembers.Table<MacroDocument> = null
let mutable private paragraphType : HostMembers.Table<RangeState> = null
let mutable private paragraphsType : HostMembers.Table<WithinState> = null
let mutable private wordsType : HostMembers.Table<WithinState> = null
let mutable private charactersType : HostMembers.Table<WithinState> = null
let mutable private sentencesType : HostMembers.Table<WithinState> = null
let mutable private tablesType : HostMembers.Table<WithinState> = null
let mutable private tableType : HostMembers.Table<TableState> = null
let mutable private rowsType : HostMembers.Table<TableState> = null
let mutable private rangeRowsType : HostMembers.Table<RangeState> = null
let mutable private rowType : HostMembers.Table<RowState> = null
let mutable private columnsType : HostMembers.Table<TableState> = null
let mutable private columnType : HostMembers.Table<ColumnState> = null
let mutable private cellsType : HostMembers.Table<CellsState> = null
let mutable private cellType : HostMembers.Table<CellState> = null
let mutable private findType : HostMembers.Table<FindObjectState> = null
let mutable private bookmarksType : HostMembers.Table<WithinState> = null
let mutable private bookmarkType : HostMembers.Table<BookmarkState> = null
let mutable private documentType : HostMembers.Table<NamedState> = null
let mutable private windowType : HostMembers.Table<NamedState> = null
let mutable private documentsType : HostMembers.Table<MacroDocument> = null
let mutable private applicationType : HostMembers.Table<MacroDocument> = null

let private rangeObject (book: MacroDocument) (r: LiveRange) = rangeType.Wrap { Book = book; R = r }
let private rangeOf (book: MacroDocument) (s: int) (e: int) = rangeObject book (book.Track(s, e))
let private selectionObject (book: MacroDocument) = selectionType.Wrap book
let private paragraphObject (book: MacroDocument) (r: LiveRange) = paragraphType.Wrap { Book = book; R = r }
let private paragraphsObject (book: MacroDocument) (within: LiveRange option) = paragraphsType.Wrap { Book = book; Within = within }
let private tablesObject (book: MacroDocument) (within: LiveRange option) = tablesType.Wrap { Book = book; Within = within }
let private tableObject (book: MacroDocument) (tableId: string) = tableType.Wrap { Book = book; TableId = tableId }
let private rowObject (book: MacroDocument) (tableId: string) (row: int) = rowType.Wrap { Book = book; TableId = tableId; Row = row }
let private columnObject (book: MacroDocument) (tableId: string) (col: int) = columnType.Wrap { Book = book; TableId = tableId; Col = col }
let private cellObject (book: MacroDocument) (c: CellRef) = cellType.Wrap { Book = book; Cell = c }
let private cellsObject (book: MacroDocument) (cells: CellSet) = cellsType.Wrap { Book = book; Cells = cells }
let private findObject (book: MacroDocument) (r: LiveRange) (isSelection: bool) = findType.Wrap { Book = book; R = r; IsSelection = isSelection; Find = findStateOf r }
let private bookmarksObject (book: MacroDocument) (within: LiveRange option) = bookmarksType.Wrap { Book = book; Within = within }
let private documentObject (book: MacroDocument) = documentType.Wrap { Book = book }
let private windowObject (book: MacroDocument) = windowType.Wrap { Book = book }
let private documentsObject (book: MacroDocument) = documentsType.Wrap book
let private applicationObject (book: MacroDocument) = applicationType.Wrap book

/// The bounds of a Range, the Selection or a Bookmark.
let private rangeBounds (value: Value) : (int * int) option =
    match rangeType.TryState value with
    | Some x -> Some (x.R.Start, x.R.End)
    | None ->
        match selectionType.TryState value with
        | Some book -> Some (book.Selection.Start, book.Selection.End)
        | None -> bookmarkType.TryState value |> Option.map (fun b -> b.S, b.E)

let private textArg (args: Value array) = toText (defaultArg (arg args 0) (VString ""))

// ---- Range ------------------------------------------------------------------------------------------------------

let private information (book: MacroDocument) (s: int) (e: int) (kind: int) : Value =
    let segs = book.SegmentsIn(s, max s e)
    let inTable = segs |> List.exists (fun seg -> seg.Table <> "")
    match kind with
    | 12 -> VBoolean inTable
    | 1 | 3 ->
        let breaks =
            book.Story.Segments
            |> Array.filter (fun seg -> seg.Start <= e && seg.Para <> "" && (book.Paragraph seg.Para |> Option.exists (fun p -> p.Metadata.PageBreakBefore)))
            |> Array.length
        VLong (1 + breaks)
    | 4 -> VLong (statistics book 0 book.Story.Total 2)
    | 2 -> VLong 1
    | 13 | 14 -> (match segs |> List.tryFind (fun seg -> seg.Table <> "") with Some seg -> VLong (seg.Row + 1) | None -> VLong -1)
    | 16 | 17 -> (match segs |> List.tryFind (fun seg -> seg.Table <> "") with Some seg -> VLong (seg.Col + 1) | None -> VLong -1)
    | 15 | 18 ->
        match segs |> List.tryFind (fun seg -> seg.Table <> "") |> Option.bind (fun seg -> tableOf book seg.Table) with
        | Some t -> VLong (if kind = 15 then t.Rows.Length else t.Rows |> List.map (fun row -> row.Cells.Length) |> List.max)
        | None -> VLong -1
    | 31 -> VBoolean (segs |> List.exists (fun seg -> seg.Mark = RowMark))
    | _ -> VEmpty

/// A table of `rows` x `cols` where position `at` is: an empty paragraph there is the one after the
/// table; a paragraph with text is split there.
let private addTable (book: MacroDocument) (at: int) (rows: int) (cols: int) : string =
    let st = book.Story
    let seg = st.Segments.[locate st at]
    if seg.Table <> "" then notAvailable "a table inside a table cell"
    let offset = at - seg.TextStart
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
            // (Word's Tables.Add: columns of equal width across the text, 0.08" cell margins left and right.)
            let setup = doc.PageSetup
            let column = (setup.PageWidth - setup.MarginLeft - setup.MarginRight) * 15.0 / float cols
            let width : TableWidth = { Kind = DxaWidth; Value = Some (System.Math.Round column) }
            let t =
                { t with
                    GridColumns = List.replicate cols width
                    PaddingLeft = Some 108.0
                    PaddingRight = Some 108.0
                    Rows = t.Rows |> List.map (fun r -> { r with Cells = r.Cells |> List.map (fun c -> { c with PreferredWidth = Some width }) }) }
            let cellPara (p: ParagraphBlock) = { p with Metadata = { p.Metadata with Style = para.Metadata.Style; SpacingAfter = para.Metadata.SpacingAfter; SpacingBefore = para.Metadata.SpacingBefore; LineSpacing = para.Metadata.LineSpacing; LineSpacingRule = para.Metadata.LineSpacingRule; Explicit = para.Metadata.Explicit &&& (ParagraphExplicit.SpacingAfter ||| ParagraphExplicit.SpacingBefore ||| ParagraphExplicit.LineSpacing) } }
            Table { t with Rows = t.Rows |> List.map (fun r -> { r with Cells = r.Cells |> List.map (fun c -> { c with Content = c.Content |> List.map cellPara }) }) }
        | other -> other
    let blocks = blocks |> List.mapi (fun i b -> if i = index then table else b)
    book.SetStructure { doc with Blocks = blocks }
    match table with
    | Table t ->
        for row in t.Rows do for cell in row.Cells do for p in cell.Content do book.InheritMark(paraId, p.Id)
        book.InheritMark(paraId, "table:" + t.Id)
    | _ -> ()
    match table with Table t -> t.Id | _ -> ""

let private convertToTable (book: MacroDocument) (r: LiveRange) (args: Value array) : Value =
    let separator =
        match arg args 0 with
        | Some (VString s) when s <> "" -> s
        | Some v -> (match toInt v with 0 -> "\r" | 1 -> "\t" | 2 -> "," | n -> string (char n))
        | None -> "\t"
    let paras = book.SegmentsIn(r.Start, max r.Start (r.End - 1)) |> List.filter (fun s -> s.Para <> "" && s.Table = "")
    let lines = paras |> List.map (fun s -> book.Story.Text.Substring(s.TextStart, s.Length))
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
                let cs, ce = cellSpan book ({ TableId = tableId; Row = ri; Col = ci } : CellRef)
                book.Replace(cs, ce - 1, text) |> ignore))
    tableObject book tableId

let private collapsedAt (r: LiveRange) (p: int) = r.Start <- p; r.End <- p

/// Range.InsertAfter: the range takes in the text.
let private insertAfter (book: MacroDocument) (r: LiveRange) (text: string) =
    let total = book.Story.Total
    // (A cell's or a row's range: the text goes before its end-of-cell mark.)
    let at =
        let story = book.Story.Text
        let mutable p = min r.End (total - 1)
        if r.End < total then
            while p > r.Start && p > 0 && (story.[p - 1] = '\a' || story.[p - 1] = '\b') do p <- p - 1
        p
    let s0 = min r.Start at
    let _, e = book.Replace(at, at, text, r)
    // (A range that took in the last paragraph mark still does.)
    r.Start <- s0; r.End <- (if r.End >= total then book.Story.Total else e)

let private insertBefore (book: MacroDocument) (r: LiveRange) (text: string) =
    let e0 = r.End
    let s, n = book.Replace(r.Start, r.Start, text, r)
    r.Start <- s; r.End <- e0 + (n - s)

let private setRangeText (x: RangeState) (value: Value) =
    let s, e = x.Book.Replace(x.R.Start, x.R.End, toText value, x.R)
    x.R.Start <- s; x.R.End <- e; x.R.Matched <- false

let private setRangeStart (x: RangeState) (value: Value) =
    let v = max 0 (min x.Book.Story.Total (toInt value))
    x.R.Matched <- false
    x.R.Start <- v
    if x.R.End < v then x.R.End <- v

let private setRangeEnd (x: RangeState) (value: Value) =
    let v = max 0 (min x.Book.Story.Total (toInt value))
    x.R.Matched <- false
    x.R.End <- v
    if x.R.Start > v then x.R.Start <- v

let private collapseRange (x: RangeState) (args: Value array) =
    // (The end of a range with the last paragraph mark is before that mark: nothing is after it.)
    if intArg args 0 1 = 0 then collapsedAt x.R (min x.R.End (x.Book.Story.Total - 1)) else collapsedAt x.R x.R.Start
    VEmpty

/// Range.Case: wdLowerCase, wdUpperCase, wdTitleWord, wdTitleSentence, wdToggleCase.
let private changeCase (x: RangeState) (value: Value) =
    let kind = toInt value
    x.Book.MapText(x.R.Start, x.R.End, fun text i ->
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

let private rangeMembers : Member<RangeState> list =
    let duplicate (x: RangeState) = rangeOf x.Book x.R.Start x.R.End
    let style (x: RangeState) = styleOfRange x.Book x.R
    let setStyle (x: RangeState) (value: Value) = setRangeStyle x.Book x.R value
    let paragraphIdsOf (x: RangeState) () = x.Book.ParagraphIdsIn(x.R.Start, x.R.End)
    let ofRange (x: RangeState) = x.Book, x.R
    let within (x: RangeState) : WithinState = { Book = x.Book; Within = Some x.R }
    let moved (f: RangeState -> Value array -> unit) (x: RangeState) (args: Value array) = f x args; VLong 1
    let nextOrPrevious (forward: bool) (x: RangeState) (args: Value array) =
        let book = x.Book
        let unit = intArg args 0 1
        let count = intArg args 1 1
        let p = if forward then moveBy book unit x.R.End (count - 1) else moveBy book unit x.R.Start (-count)
        if p >= book.Story.Total then VNothing
        else
            let s, e = spanAt book unit p
            rangeOf book s e
    let compared (test: int * int -> RangeState -> bool) (x: RangeState) (args: Value array) =
        match arg args 0 |> Option.bind rangeBounds with
        | Some bounds -> VBoolean (test bounds x)
        | None -> fail 13 "Type mismatch"
    [ prop "Text" (fun x -> VString (x.Book.TextOf(x.R.Start, x.R.End))) setRangeText
      prop "FormattedText" duplicate setRangeText
      get "Duplicate" duplicate
      prop "Start" (fun x -> VLong x.R.Start) setRangeStart
      prop "End" (fun x -> VLong x.R.End) setRangeEnd
      get "StoryLength" (fun x -> VLong x.Book.Story.Total)
      get "StoryType" (fun _ -> VLong 1)
      get "Font" (fun x -> fontObject x.Book (RangeFont x.R))
      rangeFontProperty "Bold" ofRange
      rangeFontProperty "Italic" ofRange
      rangeFontProperty "Underline" ofRange
      prop "HighlightColorIndex"
        (fun x ->
            let values = x.Book.StylesIn(x.R.Start, x.R.End) |> List.map (fun s -> indexOfHighlight s.Highlight) |> List.distinct
            VLong (match values with [ v ] -> v | _ -> wdUndefined))
        (fun x value -> let color = highlightOfIndex (toInt value) in x.Book.Format(x.R.Start, x.R.End, fun s -> { s with Highlight = color }))
      prop "Case" (fun _ -> VLong wdUndefined) changeCase
      get "ParagraphFormat" (fun x -> paragraphFormatObject x.Book (ParagraphsOf (Some x.R)))
      prop "Style" style setStyle
      prop "ParagraphStyle" style setStyle
      prop "CharacterStyle" (fun _ -> VString "Default Paragraph Font") (fun _ _ -> ())
      get "Paragraphs" (fun x -> paragraphsObject x.Book (Some x.R))
      get "Words" (fun x -> wordsType.Wrap (within x))
      get "Characters" (fun x -> charactersType.Wrap (within x))
      get "Sentences" (fun x -> sentencesType.Wrap (within x))
      get "Tables" (fun x -> tablesObject x.Book (Some x.R))
      get "Cells" (fun x -> cellsObject x.Book (CellsInRange x.R))
      get "Rows" (fun x -> rangeRowsType.Wrap x)
      get "Find" (fun x -> findObject x.Book x.R false)
      get "Document" (fun x -> documentObject x.Book)
      get "Shading" (fun x -> shadingObject x.Book (ParagraphShading (paragraphIdsOf x)))
      get "Borders" (fun x -> bordersObject x.Book (ParagraphBorders (paragraphIdsOf x)))
      get "ListFormat" (fun x -> listFormatType.Wrap x)
      get "Bookmarks" (fun x -> bookmarksObject x.Book (Some x.R))
      get "IsEndOfRowMark" (fun x -> VBoolean (x.Book.SegmentsIn(x.R.Start, x.R.Start) |> List.exists (fun s -> s.Mark = RowMark)))
      accepted "LanguageID" (VLong 1033)
      set "NoProofing" (fun _ _ -> ())
      get "Revisions" (fun _ -> emptyCollection ())
      get "Fields" (fun _ -> emptyCollection ())
      get "InlineShapes" (fun _ -> emptyCollection ())
      get "Shapes" (fun _ -> emptyCollection ())

      method "InsertAfter" [ "Text" ] (fun x args -> insertAfter x.Book x.R (textArg args); VEmpty)
      method "InsertBefore" [ "Text" ] (fun x args -> insertBefore x.Book x.R (textArg args); VEmpty)
      method "InsertParagraphAfter" [] (fun x _ -> insertAfter x.Book x.R "\r"; VEmpty)
      method "InsertParagraphBefore" [] (fun x _ -> insertBefore x.Book x.R "\r"; VEmpty)
      method "InsertParagraph" [] (fun x _ -> setRangeText x (VString "\r"); VEmpty)
      method "InsertBreak" [ "Type" ] (fun x args ->
          let ch = match intArg args 0 7 with 6 -> "\v" | _ -> "\f"
          let _, e = x.Book.Replace(x.R.Start, x.R.End, ch, x.R)
          collapsedAt x.R e
          VEmpty)
      method "Delete" [ "Unit"; "Count" ] (fun x args ->
          let book, r = x.Book, x.R
          if r.End > r.Start then
              let s, _ = book.Replace(r.Start, r.End, "", r)
              collapsedAt r s
          else
              let count = intArg args 1 1
              let s, _ = book.Replace(r.Start, moveBy book (intArg args 0 1) r.Start count, "", r)
              collapsedAt r s
          VLong 1)
      method "Collapse" [ "Direction" ] collapseRange
      method "Select" [] (fun x _ ->
          let selection = x.Book.Selection
          selection.Start <- x.R.Start; selection.End <- x.R.End; selection.ActiveAtStart <- false
          VEmpty)
      method "SetRange" [ "Start"; "End" ] (fun x args ->
          let s = intArg args 0 x.R.Start
          let e = intArg args 1 x.R.End
          x.R.Matched <- false
          x.R.Start <- max 0 (min s e); x.R.End <- min x.Book.Story.Total (max s e)
          VEmpty)
      method "Move" [ "Unit"; "Count" ] (fun x args ->
          let unit = intArg args 0 1
          let count = intArg args 1 1
          let p =
              if x.R.End > x.R.Start then
                  let collapsed = if count > 0 then x.R.End else x.R.Start
                  moveBy x.Book unit collapsed (if count > 0 then count - 1 else count + 1)
              else moveBy x.Book unit x.R.Start count
          collapsedAt x.R p
          VLong count)
      method "MoveStart" [ "Unit"; "Count" ] (moved (fun x args ->
          let p = moveBy x.Book (intArg args 0 1) x.R.Start (intArg args 1 1)
          x.R.Start <- p
          if x.R.End < p then x.R.End <- p))
      method "MoveEnd" [ "Unit"; "Count" ] (moved (fun x args ->
          let p = moveBy x.Book (intArg args 0 1) x.R.End (intArg args 1 1)
          x.R.End <- p
          if x.R.Start > p then x.R.Start <- p))
      method "Expand" [ "Unit" ] (moved (fun x args ->
          let unit = intArg args 0 2
          let s, _ = spanAt x.Book unit x.R.Start
          let _, e = spanAt x.Book unit (max x.R.Start (x.R.End - 1))
          x.R.Start <- s; x.R.End <- max s e))
      method "StartOf" [ "Unit"; "Extend" ] (moved (fun x args ->
          let s, _ = spanAt x.Book (intArg args 0 2) x.R.Start
          if intArg args 1 0 = 1 then x.R.Start <- s else collapsedAt x.R s))
      method "EndOf" [ "Unit"; "Extend" ] (moved (fun x args ->
          let _, e = spanAt x.Book (intArg args 0 2) (max x.R.Start (x.R.End - 1))
          if intArg args 1 0 = 1 then x.R.End <- e else collapsedAt x.R e))
      method "Next" [ "Unit"; "Count" ] (nextOrPrevious true)
      method "Previous" [ "Unit"; "Count" ] (nextOrPrevious false)
      method "Information" [ "Type" ] (fun x args -> information x.Book x.R.Start x.R.End (intArg args 0 0))
      method "ComputeStatistics" [ "Statistic"; "IncludeFootnotesAndEndnotes" ] (fun x args -> VLong (statistics x.Book x.R.Start x.R.End (intArg args 0 0)))
      method "InRange" [ "Range" ] (compared (fun (s, e) x -> s <= x.R.Start && x.R.End <= e))
      method "IsEqual" [ "Range" ] (compared (fun (s, e) x -> s = x.R.Start && x.R.End = e))
      method "ConvertToTable" [ "Separator"; "NumRows"; "NumColumns"; "InitialColumnWidth"; "Format"; "ApplyBorders" ] (fun x args -> convertToTable x.Book x.R args) ]

rangeType <- HostMembers.Table<RangeState>("Range", rangeMembers, defaultMember = "Text")

// ---- paragraphs, words, characters, sentences ---------------------------------------------------------------------

let private indentBy (points: float) (m: BlockMetadata) =
    { m with IndentLeft = Some (max 0.0 (defaultArg m.IndentLeft 0.0 + pxOfPt points)) } |> withFlag ParagraphExplicit.IndentLeft true

paragraphType <-
    let format (x: RangeState) : ParagraphFormatState = { Book = x.Book; Target = ParagraphsOf (Some x.R) }
    let ids (x: RangeState) () = x.Book.ParagraphIdsIn(x.R.Start, x.R.End)
    let step (count: int -> int) (x: RangeState) (args: Value array) =
        let paras = x.Book.Story.Segments |> Array.filter (fun s -> s.Mark <> RowMark)
        let here = paras |> Array.tryFindIndex (fun s -> s.Start <= x.R.Start && x.R.Start < segmentEnd s) |> Option.defaultValue 0
        let target = here + count (intArg args 0 1)
        if target < 0 || target >= paras.Length then VNothing
        else paragraphObject x.Book (x.Book.Track(paras.[target].Start, segmentEnd paras.[target]))
    let indent (points: float) (x: RangeState) (_: Value array) =
        for id in ids x () do x.Book.UpdateParagraph(id, fun p -> { p with Metadata = indentBy points p.Metadata })
        VEmpty
    let own : Member<RangeState> list =
        [ get "Range" (fun x -> rangeObject x.Book x.R)
          get "Format" (fun x -> paragraphFormatType.Wrap (format x))
          get "Shading" (fun x -> shadingObject x.Book (ParagraphShading (ids x)))
          get "Borders" (fun x -> bordersObject x.Book (ParagraphBorders (ids x)))
          prop "Style" (fun x -> styleOfRange x.Book x.R) (fun x value -> setRangeStyle x.Book x.R value)
          method "Next" [ "Count" ] (step id)
          method "Previous" [ "Count" ] (step (fun n -> -n))
          method "Indent" [] (indent 36.0)
          method "Outdent" [] (indent -36.0) ]
        @ paragraphFormatProperties format
    // (A paragraph's range's methods are the paragraph's: InsertAfter, Move...)
    HostMembers.Table<RangeState>("Paragraph", extending own (rangeMembers |> List.filter (fun m -> m.Call.IsSome)))

paragraphsType <-
    let items (p: WithinState) =
        let s, e = boundsOf p.Book p.Within
        // (An end-of-row mark is a paragraph to Word too.)
        p.Book.SegmentsIn(s, e) |> List.map (fun seg -> paragraphObject p.Book (p.Book.Track(seg.Start, segmentEnd seg))) |> List.toArray
    let format (p: WithinState) : ParagraphFormatState = { Book = p.Book; Target = ParagraphsOf p.Within }
    let mapAll (change: BlockMetadata -> BlockMetadata) (p: WithinState) (_: Value array) =
        let s, e = boundsOf p.Book p.Within
        p.Book.MapParagraphs(s, e, change)
        VEmpty
    collectionType "Paragraphs" items None
        ([ get "Format" (fun p -> paragraphFormatType.Wrap (format p))
           set "Style" (fun p value -> setParagraphStyles p.Book p.Within value)
           method "Add" [ "Range" ] (fun p args ->
               let book = p.Book
               let at =
                   match arg args 0 with
                   | Some v -> (match rangeBounds v with Some (s, _) -> s | None -> fail 13 "Type mismatch")
                   | None -> book.Story.Total - 1
               let _, e = book.Replace(at, at, "\r")
               let st = book.Story
               let seg = st.Segments.[locate st e]
               paragraphObject book (book.Track(seg.Start, segmentEnd seg)))
           method "Indent" [] (mapAll (indentBy 36.0))
           method "Outdent" [] (mapAll (indentBy -36.0))
           method "Space1" [] (mapAll (lineSpacingRuleChange (VLong 0)))
           method "Space15" [] (mapAll (lineSpacingRuleChange (VLong 1)))
           method "Space2" [] (mapAll (lineSpacingRuleChange (VLong 2))) ]
         // (The paragraphs' format, set through the collection: Paragraphs.SpaceAfter = 6.)
         @ (paragraphProperties |> List.filter (fun p -> p.Change.IsSome) |> List.map (fun p -> set p.Name (fun s value -> writeParagraphFormat (format s) p value))))

/// Words, Characters or Sentences of a range (`unit`: wdWord, wdCharacter, wdSentence).
let private spansType (typeName: string) (unit: int) =
    collectionType typeName
        (fun (w: WithinState) ->
            let book = w.Book
            let s, e = boundsOf book w.Within
            let e = max e (if e = s then s + 1 else e)
            let spans =
                if unit = 1 then [| for p in s .. min (book.Story.Total - 1) (e - 1) -> p, p + 1 |]
                else unitSpans book unit |> Array.filter (fun (a, b) -> a < e && b > s) |> Array.map (fun (a, b) -> max a s, min b e)
            spans |> Array.map (fun (a, b) -> rangeOf book a b))
        None []

wordsType <- spansType "Words" 2
charactersType <- spansType "Characters" 1
sentencesType <- spansType "Sentences" 3

// ---- tables ----------------------------------------------------------------------------------------------------

let private cellsOf (book: MacroDocument) (cells: CellSet) : CellRef list =
    match cells with
    | CellsInRange r ->
        book.SegmentsIn(r.Start, r.End) |> List.filter (fun s -> s.Table <> "" && s.Mark <> RowMark)
        |> List.map (fun s -> ({ TableId = s.Table; Row = s.Row; Col = s.Col } : CellRef)) |> List.distinct
    | CellsOfRow (tableId, row) ->
        match tableOf book tableId |> Option.bind (fun t -> List.tryItem row t.Rows) with
        | Some r -> [ for c in 0 .. r.Cells.Length - 1 -> ({ TableId = tableId; Row = row; Col = c } : CellRef) ]
        | None -> []
    | CellsOfColumn (tableId, col) ->
        match tableOf book tableId with
        | Some t -> [ for r, row in List.indexed t.Rows do if col < row.Cells.Length then yield ({ TableId = tableId; Row = r; Col = col } : CellRef) ]
        | None -> []

let private verticalAlignmentOf (value: Value) = match toInt value with 1 -> Some "center" | 3 -> Some "bottom" | _ -> Some "top"

let private columnCount (book: MacroDocument) (tableId: string) =
    match tableOf book tableId with Some t -> (t.Rows |> List.map (fun r -> r.Cells.Length) |> List.fold max 0) | None -> 0

let private setTableAlignment (book: MacroDocument) (tableId: string) (n: int) =
    let a = match n with 1 -> Some "center" | 2 -> Some "right" | _ -> Some "left"
    mapTable book tableId (fun t -> { t with Alignment = a })

let private selectSpan (book: MacroDocument) (s: int, e: int) = book.Selection.Start <- s; book.Selection.End <- e; VEmpty

/// The structure of a table changed with the caret at position `at` (insert a row before it...).
let private changeTableAt (book: MacroDocument) (at: int) (change: Document -> Document) =
    let doc = setCursor (book.ModelPosition at).BlockId 0 book.Document
    book.SetStructure (change doc)

cellType <-
    let span (c: CellState) = cellSpan c.Book c.Cell
    let width (c: CellState) = cellOf c.Book c.Cell |> Option.bind (fun x -> x.PreferredWidth) |> Option.bind (fun w -> w.Value)
    let neighbour (step: int) (c: CellState) (_: Value array) =
        match tableOf c.Book c.Cell.TableId with
        | Some t ->
            let cells = [ for r, row in List.indexed t.Rows do for col in 0 .. row.Cells.Length - 1 -> { c.Cell with Row = r; Col = col } ]
            let i = cells |> List.findIndex (fun x -> x.Row = c.Cell.Row && x.Col = c.Cell.Col)
            let j = i + step
            if j < 0 || j >= cells.Length then VNothing else cellObject c.Book cells.[j]
        | None -> VNothing
    HostMembers.Table<CellState>("Cell", [
        get "Range" (fun c -> let s, e = span c in rangeOf c.Book s e)
        get "Shading" (fun c -> shadingObject c.Book (CellShading (fun () -> [ c.Cell ])))
        get "Borders" (fun c -> bordersObject c.Book (CellBorders (fun () -> [ c.Cell ])))
        get "RowIndex" (fun c -> VLong (c.Cell.Row + 1))
        get "ColumnIndex" (fun c -> VLong (c.Cell.Col + 1))
        prop "VerticalAlignment"
            (fun c -> VLong (match cellOf c.Book c.Cell |> Option.bind (fun x -> x.VerticalAlignment) with Some "center" -> 1 | Some "bottom" -> 3 | _ -> 0))
            (fun c value -> let v = verticalAlignmentOf value in mapCells c.Book [ c.Cell ] (fun x -> { x with VerticalAlignment = v }))
        prop "Width"
            (fun c -> VDouble (width c |> Option.map (fun twips -> twips / 20.0) |> Option.defaultValue 0.0))
            (fun c value ->
                let px = pxOfPt (toFloat value)
                mapCells c.Book [ c.Cell ] (fun x -> { x with PreferredWidth = Some { Kind = DxaWidth; Value = Some (px * 15.0) } }))
        method "Select" [] (fun c _ -> selectSpan c.Book (span c))
        method "Next" [] (neighbour 1)
        method "Previous" [] (neighbour -1) ])

cellsType <-
    let cells (x: CellsState) () = cellsOf x.Book x.Cells
    collectionType "Cells" (fun x -> cells x () |> List.map (cellObject x.Book) |> List.toArray) None
        [ get "Shading" (fun x -> shadingObject x.Book (CellShading (cells x)))
          get "Borders" (fun x -> bordersObject x.Book (CellBorders (cells x)))
          set "VerticalAlignment" (fun x value -> let v = verticalAlignmentOf value in mapCells x.Book (cells x ()) (fun c -> { c with VerticalAlignment = v }))
          method "Merge" [] (fun _ _ -> notAvailable "Cells.Merge") ]

rowType <-
    let row (x: RowState) = tableOf x.Book x.TableId |> Option.bind (fun t -> List.tryItem x.Row t.Rows)
    let cells (x: RowState) () = cellsOf x.Book (CellsOfRow (x.TableId, x.Row))
    let span (x: RowState) = rowSpan x.Book x.TableId x.Row
    let mapRow (x: RowState) (change: TableRow -> TableRow) =
        mapTable x.Book x.TableId (fun t -> { t with Rows = t.Rows |> List.mapi (fun i r -> if i = x.Row then change r else r) })
    HostMembers.Table<RowState>("Row", [
        get "Range" (fun x -> let s, e = span x in rangeOf x.Book s e)
        get "Cells" (fun x -> cellsObject x.Book (CellsOfRow (x.TableId, x.Row)))
        get "Shading" (fun x -> shadingObject x.Book (CellShading (cells x)))
        get "Borders" (fun x -> bordersObject x.Book (CellBorders (cells x)))
        get "Index" (fun x -> VLong (x.Row + 1))
        prop "HeadingFormat"
            (fun x -> VBoolean (row x |> Option.exists (fun r -> r.IsHeader)))
            (fun x value -> let on = toBool value in mapRow x (fun r -> { r with IsHeader = on }))
        prop "Height"
            (fun x -> VDouble (row x |> Option.bind (fun r -> r.Height) |> Option.map (fun twips -> twips / 20.0) |> Option.defaultValue 0.0))
            (fun x value -> let twips = toFloat value * 20.0 in mapRow x (fun r -> { r with Height = Some twips; HeightRule = Some (defaultArg r.HeightRule "atLeast") }))
        set "HeightRule" (fun x value -> let rule = match toInt value with 2 -> "exact" | 1 -> "atLeast" | _ -> "auto" in mapRow x (fun r -> { r with HeightRule = Some rule }))
        set "AllowBreakAcrossPages" (fun x value -> let on = toBool value in mapRow x (fun r -> { r with CanSplit = Some on }))
        get "IsFirst" (fun x -> VBoolean (x.Row = 0))
        get "IsLast" (fun x -> VBoolean (tableOf x.Book x.TableId |> Option.exists (fun t -> x.Row = t.Rows.Length - 1)))
        prop "Alignment" (fun _ -> VLong 0) (fun x value -> setTableAlignment x.Book x.TableId (toInt value))
        method "Delete" [] (fun x _ -> changeTableAt x.Book (fst (span x)) deleteSelectedTableRow; VEmpty)
        method "Select" [] (fun x _ -> selectSpan x.Book (span x)) ])

rowsType <-
    let count (t: TableState) = match tableOf t.Book t.TableId with Some x -> x.Rows.Length | None -> 0
    let mapRows (t: TableState) (change: TableRow -> TableRow) = mapTable t.Book t.TableId (fun x -> { x with Rows = x.Rows |> List.map change })
    collectionType "Rows" (fun t -> [| for r in 0 .. count t - 1 -> rowObject t.Book t.TableId r |]) None
        [ prop "Alignment"
            (fun t -> VLong (match tableOf t.Book t.TableId |> Option.bind (fun x -> x.Alignment) with Some "center" -> 1 | Some "right" -> 2 | _ -> 0))
            (fun t value -> setTableAlignment t.Book t.TableId (toInt value))
          set "LeftIndent" (fun t value -> let px = pxOfPt (toFloat value) in mapTable t.Book t.TableId (fun x -> { x with Indent = Some px }))
          set "AllowBreakAcrossPages" (fun _ _ -> ())
          set "WrapAroundText" (fun _ _ -> ())
          prop "HeadingFormat" (fun _ -> VBoolean false) (fun t value -> let on = toBool value in mapRows t (fun r -> { r with IsHeader = on }))
          set "Height" (fun t value -> let twips = toFloat value * 20.0 in mapRows t (fun r -> { r with Height = Some twips; HeightRule = Some (defaultArg r.HeightRule "atLeast") }))
          method "Add" [ "BeforeRow" ] (fun t args ->
              let before = arg args 0 |> Option.bind rowType.TryState |> Option.map (fun r -> r.Row + 1)
              let n = count t
              let rowIndex, place = match before with Some i -> i - 1, insertTableRowBeforeSelection | None -> n - 1, insertTableRowAfterSelection
              let s, _ = rowSpan t.Book t.TableId rowIndex
              changeTableAt t.Book s place
              rowObject t.Book t.TableId (match before with Some i -> i - 1 | None -> n)) ]

rangeRowsType <-
    let rows (x: RangeState) =
        x.Book.SegmentsIn(x.R.Start, x.R.End) |> List.filter (fun s -> s.Table <> "") |> List.map (fun s -> s.Table, s.Row) |> List.distinct
    collectionType "Rows" (fun x -> rows x |> List.map (fun (t, i) -> rowObject x.Book t i) |> List.toArray) None []

columnType <-
    let cells (x: ColumnState) () = cellsOf x.Book (CellsOfColumn (x.TableId, x.Col))
    let setWidth (x: ColumnState) (value: Value) =
        let px = pxOfPt (toFloat value)
        mapCells x.Book (cells x ()) (fun c -> { c with PreferredWidth = Some { Kind = DxaWidth; Value = Some (px * 15.0) } })
        mapTable x.Book x.TableId (fun t ->
            let n = columnCount x.Book x.TableId
            let grid = if t.GridColumns.Length = n then t.GridColumns else List.replicate n { Kind = AutoWidth; Value = None }
            { t with GridColumns = grid |> List.mapi (fun i g -> if i = x.Col then { Kind = DxaWidth; Value = Some (px * 15.0) } else g) })
    HostMembers.Table<ColumnState>("Column", [
        get "Cells" (fun x -> cellsObject x.Book (CellsOfColumn (x.TableId, x.Col)))
        get "Shading" (fun x -> shadingObject x.Book (CellShading (cells x)))
        get "Borders" (fun x -> bordersObject x.Book (CellBorders (cells x)))
        get "Index" (fun x -> VLong (x.Col + 1))
        prop "Width"
            (fun x -> VDouble (cells x () |> List.tryHead |> Option.bind (cellOf x.Book) |> Option.bind (fun c -> c.PreferredWidth) |> Option.bind (fun w -> w.Value) |> Option.map (fun twips -> twips / 20.0) |> Option.defaultValue 0.0))
            setWidth
        set "PreferredWidth" setWidth
        set "PreferredWidthType" (fun _ _ -> ())
        method "Select" [] (fun _ _ -> VEmpty)
        method "Delete" [] (fun x _ ->
            let s, _ = cellSpan x.Book ({ TableId = x.TableId; Row = 0; Col = x.Col } : CellRef)
            changeTableAt x.Book s deleteSelectedTableColumn
            VEmpty)
        method "AutoFit" [] (fun _ _ -> VEmpty) ])

columnsType <-
    collectionType "Columns" (fun (t: TableState) -> [| for c in 0 .. columnCount t.Book t.TableId - 1 -> columnObject t.Book t.TableId c |]) None
        [ set "PreferredWidthType" (fun _ _ -> ())
          set "PreferredWidth" (fun _ _ -> ())
          method "Add" [ "BeforeColumn" ] (fun t _ ->
              let n = columnCount t.Book t.TableId
              let s, _ = cellSpan t.Book ({ TableId = t.TableId; Row = 0; Col = n - 1 } : CellRef)
              changeTableAt t.Book s insertTableColumnAfterSelection
              columnObject t.Book t.TableId n)
          method "AutoFit" [] (fun _ _ -> VEmpty)
          method "DistributeWidth" [] (fun _ _ -> VEmpty) ]

tableType <-
    let table (t: TableState) = match tableOf t.Book t.TableId with Some x -> x | None -> fail 5941 "The table is no longer in the document."
    let allCells (t: TableState) () = [ for r, row in List.indexed (table t).Rows do for col in 0 .. row.Cells.Length - 1 -> ({ TableId = t.TableId; Row = r; Col = col } : CellRef) ]
    let setStyle (t: TableState) (value: Value) =
        let book = t.Book
        let text = match scalarOf value with VInteger n | VLong n -> string n | v -> toText v
        book.Settings.["tablestyle:" + t.TableId] <- VString text
        if not (applyTableStyle book t.TableId text) then
            // (A style of the document: the reference is kept for the file.)
            match book.Document.Styles |> List.tryFind (fun s -> s.StyleType = "table" && (eq s.Id text || s.Name |> Option.exists (eq text))) with
            | Some s -> mapTable book t.TableId (fun x -> { x with StyleId = Some s.Id })
            | None -> ()
    let autoFormat (t: TableState) (args: Value array) =
        // (Word's table formats of Word 97 - wdTableFormatColorful2 and the like - as built-in styles.)
        let book, tableId = t.Book, t.TableId
        let format = intArg args 0 0
        let style =
            match format with
            | 0 -> None
            | 9 ->
                // (Word 2019's Colorful 2: a dark red header row in white bold italic, pale yellow
                // rows, the first column bold italic, no borders but the bottom.)
                mapTable book tableId (fun t ->
                    let line = Some ({ Style = Some "single"; Color = Some "#000000"; Width = Some (12.0 / 6.0) } : TableBorder)
                    { t with
                        Borders = Some { emptyTableBorders with Bottom = line }
                        Rows = t.Rows |> List.mapi (fun r row -> { row with Cells = row.Cells |> List.map (fun c -> { c with BackgroundColor = Some (if r = 0 then "#800000" else "#FFFFCC"); Shading = None }) }) })
                let s0, e0 = rowSpan book tableId 0
                book.Format(s0, e0, fun st -> { TB.withExplicitItalic true (TB.withExplicitBold true st) with Color = Some "#FFFFFF" })
                match tableOf book tableId with
                | Some x ->
                    for r in 1 .. x.Rows.Length - 1 do
                        let cs, ce = cellSpan book ({ TableId = tableId; Row = r; Col = 0 } : CellRef)
                        book.Format(cs, ce, fun st -> TB.withExplicitItalic true (TB.withExplicitBold true st))
                | None -> ()
                None
            | 1 | 2 | 3 -> Some "Plain Table 1"
            | 8 | 9 | 10 -> Some (sprintf "Grid Table 5 Dark - Accent %d" (format - 7))
            | 4 | 5 | 6 | 7 -> Some "List Table 3"
            | _ -> Some "Grid Table 4 - Accent 1"
        style |> Option.iter (fun name -> applyTableStyle book tableId name |> ignore)
        VEmpty
    let sort (t: TableState) (args: Value array) =
        // (By the first column, or FieldNumber; the header row stays first when ExcludeHeader.)
        let exclude = arg args 0 |> Option.map toBool |> Option.defaultValue false
        let field = intArg args 1 1 - 1
        let descending = intArg args 3 0 = 1
        let doc = t.Book.Document
        t.Book.SetStructure
            { doc with
                Blocks =
                    doc.Blocks |> List.map (function
                        | Table x when x.Id = t.TableId ->
                            let head, body = if exclude then List.truncate 1 x.Rows, List.skip (min 1 x.Rows.Length) x.Rows else [], x.Rows
                            let key (row: TableRow) =
                                row.Cells |> List.tryItem field |> Option.map (fun c -> c.Content |> List.map (fun p -> TB.toString p.Buffer) |> String.concat "\r") |> Option.defaultValue ""
                            let numeric (s: string) = match Double.TryParse(s, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture) with | true, d -> Some d | _ -> None
                            let sorted =
                                if body |> List.forall (fun row -> (numeric (key row)).IsSome) then body |> List.sortBy (fun row -> (numeric (key row)).Value)
                                else body |> List.sortWith (fun a b -> String.Compare(key a, key b, StringComparison.OrdinalIgnoreCase))
                            Table { x with Rows = head @ (if descending then List.rev sorted else sorted) }
                        | other -> other) }
        VEmpty
    let noEffect name = set name (fun (_: TableState) _ -> ())
    HostMembers.Table<TableState>("Table", [
        get "Range" (fun t -> let s, e = tableSpan t.Book t.TableId in rangeOf t.Book s e)
        get "Rows" (fun t -> rowsType.Wrap t)
        get "Columns" (fun t -> columnsType.Wrap t)
        get "Borders" (fun t -> bordersObject t.Book (TableBorders t.TableId))
        get "Shading" (fun t -> shadingObject t.Book (CellShading (allCells t)))
        prop "Style" (fun t -> VString (match t.Book.Settings.TryGetValue ("tablestyle:" + t.TableId) with | true, v -> toText v | _ -> "Table Grid")) setStyle
        get "NestingLevel" (fun _ -> VLong 1)
        get "Uniform" (fun _ -> VBoolean true)
        accepted "Title" (VString "")
        accepted "Descr" (VString "")
        prop "PreferredWidthType" (fun _ -> VLong 1) (fun t value -> t.Book.Settings.["tablewidthtype:" + t.TableId] <- value)
        prop "PreferredWidth" (fun _ -> VDouble 0.0) (fun t value ->
            let kind = match t.Book.Settings.TryGetValue ("tablewidthtype:" + t.TableId) with | true, v -> toInt v | _ -> 3
            let width =
                if kind = 2 then { Kind = PercentWidth; Value = Some (toFloat value * 50.0) }
                else { Kind = DxaWidth; Value = Some (toFloat value * 20.0) }
            mapTable t.Book t.TableId (fun x -> { x with PreferredWidth = Some width }))
        accepted "AllowAutoFit" (VBoolean true)
        noEffect "Spacing"
        noEffect "TopPadding"
        noEffect "BottomPadding"
        noEffect "LeftPadding"
        noEffect "RightPadding"
        accepted "ApplyStyleHeadingRows" (VBoolean true)
        accepted "ApplyStyleFirstColumn" (VBoolean true)
        accepted "ApplyStyleLastRow" (VBoolean true)
        accepted "ApplyStyleLastColumn" (VBoolean true)
        accepted "ApplyStyleRowBands" (VBoolean true)
        noEffect "ApplyStyleColumnBands"
        method "Cell" [ "Row"; "Column" ] (fun t args ->
            let r = intArg args 0 1 - 1
            let c = intArg args 1 1 - 1
            let x = table t
            if r < 0 || r >= x.Rows.Length || c < 0 || c >= x.Rows.[r].Cells.Length then fail 5941 "The requested member of the collection does not exist."
            cellObject t.Book { TableId = t.TableId; Row = r; Col = c })
        method "AutoFitBehavior" [ "Behavior" ] (fun t args ->
            if intArg args 0 0 = 2 then mapTable t.Book t.TableId (fun x -> { x with PreferredWidth = Some { Kind = PercentWidth; Value = Some 5000.0 } })
            VEmpty)
        method "AutoFormat" [] autoFormat
        method "Delete" [] (fun t _ ->
            let book = t.Book
            book.SetStructure { book.Document with Blocks = book.Document.Blocks |> List.filter (function Table x -> x.Id <> t.TableId | _ -> true) }
            VEmpty)
        method "Select" [] (fun t _ -> selectSpan t.Book (tableSpan t.Book t.TableId))
        method "Sort" [ "ExcludeHeader"; "FieldNumber"; "SortFieldType"; "SortOrder" ] sort ])

tablesType <-
    let items (w: WithinState) =
        let s, e = boundsOf w.Book w.Within
        let e = max e (s + 1)
        w.Book.TableSpans() |> List.filter (fun (_, ts, te) -> ts < e && te > s) |> List.map (fun (id, _, _) -> tableObject w.Book id) |> List.toArray
    collectionType "Tables" items None
        [ method "Add" [ "Range"; "NumRows"; "NumColumns"; "DefaultTableBehavior"; "AutoFitBehavior" ] (fun w args ->
            let book = w.Book
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
            tableObject book id) ]

// ---- Find ---------------------------------------------------------------------------------------------------

let private findExecute (f: FindObjectState) (args: Value array) : Value =
    let book, r, state = f.Book, f.R, f.Find
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

findType <-
    let flag name (read: FindState -> bool) (write: FindState -> bool -> unit) =
        prop name (fun (f: FindObjectState) -> VBoolean (read f.Find)) (fun f value -> write f.Find (toBool value))
    let noEffect name = method name [] (fun (_: FindObjectState) _ -> VEmpty)
    let ignored =
        [ "MatchSoundsLike"; "MatchAllWordForms"; "MatchPrefix"; "MatchSuffix"; "MatchDiacritics"; "MatchKashida"; "MatchAlefHamza"
          "MatchControl"; "MatchByte"; "MatchFuzzy"; "IgnoreSpace"; "IgnorePunct" ]
    HostMembers.Table<FindObjectState>("Find",
        [ prop "Text" (fun f -> VString f.Find.Text) (fun f value -> f.Find.Text <- toText value)
          get "Replacement" (fun f -> replacementType.Wrap { Book = f.Book; Find = f.Find })
          flag "Forward" (fun s -> s.Forward) (fun s v -> s.Forward <- v)
          prop "Wrap" (fun f -> VLong f.Find.Wrap) (fun f value -> f.Find.Wrap <- toInt value)
          flag "Format" (fun s -> s.Format) (fun s v -> s.Format <- v)
          flag "MatchCase" (fun s -> s.MatchCase) (fun s v -> s.MatchCase <- v)
          flag "MatchWholeWord" (fun s -> s.MatchWholeWord) (fun s v -> s.MatchWholeWord <- v)
          flag "MatchWildcards" (fun s -> s.MatchWildcards) (fun s v -> s.MatchWildcards <- v)
          get "Found" (fun f -> VBoolean f.Find.Found)
          get "Font" (fun f -> fontObject f.Book (SpecFont f.Find.Font))
          prop "Highlight"
            (fun f -> VLong (if f.Find.Highlight < 0 then wdUndefined else f.Find.Highlight))
            (fun f value -> f.Find.Highlight <- (if toBool value then 1 else 0))
          get "Parent" (fun f -> if f.IsSelection then selectionObject f.Book else rangeObject f.Book f.R)
          set "Style" (fun _ _ -> ())
          method "Execute" [ "FindText"; "MatchCase"; "MatchWholeWord"; "MatchWildcards"; "MatchSoundsLike"; "MatchAllWordForms"
                             "Forward"; "Wrap"; "Format"; "ReplaceWith"; "Replace"; "MatchKashida"; "MatchDiacritics"; "MatchAlefHamza"; "MatchControl" ]
            findExecute
          method "ClearFormatting" [] (fun f _ -> f.Find.Font.Clear(); f.Find.Highlight <- -1; VEmpty)
          noEffect "ClearAllFuzzyOptions"
          noEffect "SetAllFuzzyOptions" ]
        @ (ignored |> List.map (fun name -> accepted name (VBoolean false))))

// ---- Selection ---------------------------------------------------------------------------------------------

selectionType <-
    let asRange (book: MacroDocument) : RangeState = { Book = book; R = book.Selection }
    let collapsedAt (book: MacroDocument) (p: int) = let s = book.Selection in s.Start <- p; s.End <- p; s.ActiveAtStart <- false
    /// Move the selection's active end to `p` (extend) or the caret to `p`.
    let moveTo (book: MacroDocument) (p: int) (extend: bool) =
        let s = book.Selection
        let p = max 0 (min (book.Story.Total - 1) p)
        if extend then
            let anchor = if s.ActiveAtStart then s.End else s.Start
            s.Start <- min anchor p; s.End <- max anchor p
            s.ActiveAtStart <- p < anchor
        else collapsedAt book p
    let active (book: MacroDocument) = let s = book.Selection in if s.ActiveAtStart then s.Start else s.End
    let typeText (book: MacroDocument) (text: string) =
        let s = book.Selection
        if book.PendingAt <> s.Start || s.End > s.Start then book.Pending.Clear()
        let a, e = book.Replace(s.Start, s.End, text, s)
        if book.Pending.Count > 0 then book.Format(a, e, specChange book.Pending)
        collapsedAt book e
        book.PendingAt <- e
        VEmpty
    let lineStart (book: MacroDocument) p = let st = book.Story in st.Segments.[locate st p].Start
    let lineEnd (book: MacroDocument) p = let st = book.Story in let seg = st.Segments.[locate st p] in seg.TextStart + seg.Length
    let extend (args: Value array) index = intArg args index 0 = 1
    let homeOrEnd (home: bool) (book: MacroDocument) (args: Value array) =
        let unit = intArg args 0 5
        let s = book.Selection
        let p =
            match unit with
            | 6 -> if home then 0 else book.Story.Total - 1
            | 10 | 15 -> if home then fst (spanAt book unit s.Start) else snd (spanAt book unit s.Start) - 1
            | _ -> if home then lineStart book s.Start else lineEnd book (active book)
        moveTo book p (extend args 1)
        VLong 1
    let sideways (sign: int) (book: MacroDocument) (args: Value array) =
        let unit = intArg args 0 1
        let count = sign * intArg args 1 1
        let s = book.Selection
        let ext = extend args 2
        if not ext && s.End > s.Start then
            let collapsed = if count > 0 then s.End else s.Start
            collapsedAt book (moveBy book unit collapsed (if count > 0 then count - 1 else count + 1))
        elif unit = 12 then
            // (A cell to the right: the next cell, selected as Word selects it.)
            let spans = unitSpans book 12
            let i = spans |> Array.tryFindIndex (fun (a, b) -> a <= s.Start && s.Start < b) |> Option.defaultValue -1
            let j = i + count
            if j >= 0 && j < spans.Length then
                let a, b = spans.[j]
                s.Start <- a; s.End <- b - 1; s.ActiveAtStart <- false
        else moveTo book (moveBy book unit (active book) count) ext
        VLong (abs count)
    let upOrDown (sign: int) (book: MacroDocument) (args: Value array) =
        let unit = intArg args 0 5
        let count = sign * intArg args 1 1
        let unit = if unit = 5 || unit = 7 then 4 else unit
        let from = active book
        let target =
            if count > 0 then moveBy book unit from count
            else moveBy book unit (lineStart book from) (count + (if lineStart book from = from then 0 else 1))
        moveTo book target (extend args 2)
        VLong (abs count)
    let selectSpanOf (unit: int) (book: MacroDocument) (_: Value array) =
        let s = book.Selection
        let a, b = spanAt book unit s.Start
        s.Start <- a; s.End <- b
        VEmpty
    let clipboard name = method name [] (fun (_: MacroDocument) _ -> notAvailable ("Selection." + name.ToLowerInvariant() + " (the clipboard)"))
    let own : Member<MacroDocument> list =
        [ get "Font" (fun book ->
            let s = book.Selection
            if s.End = s.Start then
                if book.PendingAt <> s.Start then
                    book.Pending.Clear()
                    book.PendingAt <- s.Start
                fontObject book (CaretFont s)
            else fontObject book (RangeFont s))
          get "Range" (fun book -> rangeOf book book.Selection.Start book.Selection.End)
          get "Find" (fun book -> findObject book book.Selection true)
          get "Type" (fun book -> VLong (if book.Selection.End > book.Selection.Start then 2 else 1))
          get "Active" (fun _ -> VBoolean true)
          get "IsEndOfRowMark" (fun _ -> VBoolean false)
          prop "Start" (fun book -> VLong book.Selection.Start) (fun book value -> book.Pending.Clear(); setRangeStart (asRange book) value)
          prop "End" (fun book -> VLong book.Selection.End) (fun book value -> book.Pending.Clear(); setRangeEnd (asRange book) value)
          method "TypeText" [ "Text" ] (fun book args -> typeText book (textArg args))
          method "TypeParagraph" [] (fun book _ -> typeText book "\r")
          method "TypeBackspace" [] (fun book _ ->
              let s = book.Selection
              if s.End > s.Start then
                  book.Replace(s.Start, s.End, "", s) |> ignore
                  collapsedAt book s.Start
              elif s.Start > 0 then
                  let p, _ = book.Replace(s.Start - 1, s.Start, "", s)
                  collapsedAt book p
              VEmpty)
          method "HomeKey" [ "Unit"; "Extend" ] (homeOrEnd true)
          method "EndKey" [ "Unit"; "Extend" ] (homeOrEnd false)
          method "MoveRight" [ "Unit"; "Count"; "Extend" ] (sideways 1)
          method "MoveLeft" [ "Unit"; "Count"; "Extend" ] (sideways -1)
          method "MoveDown" [ "Unit"; "Count"; "Extend" ] (upOrDown 1)
          method "MoveUp" [ "Unit"; "Count"; "Extend" ] (upOrDown -1)
          method "WholeStory" [] (fun book _ -> let s = book.Selection in s.Start <- 0; s.End <- book.Story.Total; VEmpty)
          method "Collapse" [ "Direction" ] (fun book args -> book.Pending.Clear(); collapseRange (asRange book) args)
          method "Delete" [ "Unit"; "Count" ] (fun book args ->
              let s = book.Selection
              if s.End > s.Start then
                  book.Replace(s.Start, s.End, "", s) |> ignore
                  collapsedAt book s.Start
              else
                  let count = intArg args 1 1
                  book.Replace(s.Start, moveBy book (intArg args 0 1) s.Start count, "", s) |> ignore
                  collapsedAt book s.Start
              VLong 1)
          method "EscapeKey" [] (fun book _ -> collapsedAt book book.Selection.Start; VEmpty)
          method "Extend" [] (fun _ _ -> VEmpty)
          method "SelectRow" [] (selectSpanOf 10)
          method "SelectCell" [] (selectSpanOf 12)
          method "SelectColumn" [] (selectSpanOf 12)
          method "GoTo" [] (fun _ _ -> VEmpty)
          clipboard "Copy"
          clipboard "Cut"
          clipboard "Paste"
          clipboard "PasteAndFormat" ]
    // (The Selection is a Range too: its text, formatting, InsertAfter, Move...)
    HostMembers.Table<MacroDocument>("Selection", extending own (rangeMembers |> List.map (lift asRange)), defaultMember = "Text")

// ---- bookmarks ------------------------------------------------------------------------------------------------

bookmarkType <-
    HostMembers.Table<BookmarkState>("Bookmark",
        [ get "Name" (fun b -> VString b.Info.Name)
          get "Range" (fun b -> rangeOf b.Book b.S b.E)
          get "Start" (fun b -> VLong b.S)
          get "End" (fun b -> VLong b.E)
          get "Empty" (fun b -> VBoolean (b.S = b.E))
          method "Delete" [] (fun b _ -> b.Book.SetFormatting (removeBookmark b.Info.Id b.Book.Document); VEmpty)
          method "Select" [] (fun b _ -> selectSpan b.Book (b.S, b.E)) ],
        defaultMember = "Name")

bookmarksType <-
    let infos (w: WithinState) =
        bookmarkInfos w.Book.Document
        |> List.map (fun info -> info, w.Book.PositionOf(info.BlockId, info.StartOffset), w.Book.PositionOf(info.BlockId, info.EndOffset))
        |> List.filter (fun (_, s, e) -> match w.Within with Some r -> s >= r.Start && e <= r.End | None -> true)
        |> List.sortBy (fun (_, s, _) -> s)
    let named (w: WithinState) (key: string) = infos w |> List.tryFind (fun (info, _, _) -> eq info.Name key)
    let bookmark (book: MacroDocument) (info: BookmarkInfo, s: int, e: int) = bookmarkType.Wrap { Book = book; Info = info; S = s; E = e }
    let items (w: WithinState) = infos w |> List.map (bookmark w.Book) |> List.toArray
    collectionType "Bookmarks" items
        (Some (fun w args ->
            match arg args 0 with
            | Some (VString key) ->
                match named w key with
                | Some b -> bookmark w.Book b
                | None -> fail 5941 (sprintf "The requested member of the collection does not exist (bookmark \"%s\")." key)
            | _ -> itemAt "Bookmarks" (items w) args))
        [ method "Exists" [ "Name" ] (fun w args -> VBoolean (named w (textArg args)).IsSome)
          method "Add" [ "Name"; "Range" ] (fun w args ->
              let book = w.Book
              let name' = textArg args
              if name' = "" || not (Char.IsLetter name'.[0]) then fail 5828 "Bookmark names must begin with a letter."
              let s, e =
                  match arg args 1 |> Option.bind rangeBounds with
                  | Some b -> b
                  | None -> book.Selection.Start, book.Selection.End
              // An old one of that name goes; the new one is on the text of its first paragraph.
              match named w name' with
              | Some (info, _, _) -> book.SetFormatting (removeBookmark info.Id book.Document)
              | None -> ()
              let seg = book.Story.Segments.[locate book.Story s]
              if seg.Para <> "" then
                  let a = s - seg.TextStart
                  let b = min (seg.TextStart + seg.Length) e - seg.TextStart
                  let id = string (book.NextBookmarkId())
                  book.UpdateParagraph(seg.Para, fun p -> { p with Buffer = TB.setBookmark a (max 0 (b - a)) id name' p.Buffer })
              match named w name' with
              | Some b -> bookmark book b
              | None -> VNothing) ]

// ---- the document, its window, the application -----------------------------------------------------------------

documentType <-
    let whole (d: NamedState) : WithinState = { Book = d.Book; Within = None }
    let empty name = get name (fun (_: NamedState) -> emptyCollection ())
    let noEffect name = set name (fun (_: NamedState) _ -> ())
    let nothing name = method name [] (fun (_: NamedState) _ -> VEmpty)
    let notRun name parameters =
        method name parameters (fun (d: NamedState) _ ->
            d.Book.Hooks.Output (sprintf "(Document.%s is not run here: the host saves the document)" ((name: string).ToLowerInvariant()))
            VEmpty)
    HostMembers.Table<NamedState>("Document",
        [ get "Content" (fun d -> rangeOf d.Book 0 d.Book.Story.Total)
          get "Paragraphs" (fun d -> paragraphsObject d.Book None)
          get "Words" (fun d -> wordsType.Wrap (whole d))
          get "Characters" (fun d -> charactersType.Wrap (whole d))
          get "Sentences" (fun d -> sentencesType.Wrap (whole d))
          get "Tables" (fun d -> tablesObject d.Book None)
          get "Bookmarks" (fun d -> bookmarksObject d.Book None)
          get "Styles" (fun d -> stylesType.Wrap d.Book)
          get "PageSetup" (fun d -> pageSetupType.Wrap d.Book)
          get "Variables" (fun d -> variablesType.Wrap d.Book)
          get "Name" (fun d -> VString d.Book.Hooks.FileName)
          get "FullName" (fun d -> VString d.Book.Hooks.FileName)
          get "Path" (fun _ -> VString "")
          accepted "Saved" (VBoolean false)
          get "ReadOnly" (fun _ -> VBoolean false)
          accepted "TrackRevisions" (VBoolean false)
          get "Application" (fun d -> applicationObject d.Book)
          get "ActiveWindow" (fun d -> windowObject d.Book)
          get "Selection" (fun d -> selectionObject d.Book)
          empty "InlineShapes"
          empty "Shapes"
          empty "Fields"
          empty "Comments"
          empty "Revisions"
          empty "Footnotes"
          empty "Endnotes"
          empty "Hyperlinks"
          empty "TablesOfContents"
          empty "Lists"
          noEffect "ShowGrammaticalErrors"
          noEffect "ShowSpellingErrors"
          noEffect "SpellingChecked"
          noEffect "GrammarChecked"
          method "Range" [ "Start"; "End" ] (fun d args ->
              let total = d.Book.Story.Total
              let s = intArg args 0 0
              let e = intArg args 1 (if arg args 0 |> Option.isSome && args.Length = 1 then s else total)
              let e = if args.Length = 0 then total else e
              rangeOf d.Book (max 0 (min total s)) (max 0 (min total e)))
          method "ComputeStatistics" [ "Statistic"; "IncludeFootnotesAndEndnotes" ] (fun d args -> VLong (statistics d.Book 0 d.Book.Story.Total (intArg args 0 0)))
          nothing "Activate"
          nothing "Select"
          nothing "Repaginate"
          nothing "UndoClear"
          notRun "Save" []
          notRun "SaveAs" [ "FileName"; "FileFormat" ]
          notRun "SaveAs2" [ "FileName"; "FileFormat" ]
          notRun "Close" [ "SaveChanges" ]
          notRun "PrintOut" []
          notRun "ExportAsFixedFormat" []
          method "Undo" [] (fun _ _ -> VBoolean false)
          method "Redo" [] (fun _ _ -> VBoolean false) ])

windowType <-
    let display name = accepted name (VBoolean true)
    let nothing name = method name [] (fun (_: NamedState) _ -> VEmpty)
    HostMembers.Table<NamedState>("Window",
        [ get "View" (fun _ -> viewType.Wrap ())
          get "ActivePane" (fun _ -> paneType.Wrap ())
          get "Selection" (fun w -> selectionObject w.Book)
          prop "Caption" (fun w -> VString w.Book.Hooks.FileName) (fun _ _ -> ())
          accepted "WindowState" (VLong 0)
          display "DisplayRulers"
          display "DisplayVerticalRuler"
          display "DisplayHorizontalScrollBar"
          display "DisplayVerticalScrollBar"
          nothing "Activate"
          nothing "ScrollIntoView"
          nothing "SmallScroll"
          nothing "LargeScroll" ])

documentsType <-
    collectionType "Documents" (fun (book: MacroDocument) -> [| documentObject book |])
        (Some (fun book args ->
            match arg args 0 with
            | Some (VString _) | Some (VLong 1) | Some (VDouble 1.0) -> documentObject book
            | _ -> noSuchMember ()))
        [ method "Add" [ "Template"; "NewTemplate"; "DocumentType"; "Visible" ] (fun book _ ->
            // (One document is open in the editor: a new document is the editor's, made empty.)
            let para = newEmptyParagraph ()
            book.SetStructure { book.Document with Blocks = [ Paragraph para ] }
            book.Selection.Start <- 0; book.Selection.End <- 0
            documentObject book)
          method "Open" [] (fun _ _ -> notAvailable "Documents.Open (files)") ]

applicationType <-
    let key (name: string) = "app:" + name.ToLowerInvariant()
    let setting name (fallback: Value) =
        prop name (fun (book: MacroDocument) -> match book.Settings.TryGetValue (key name) with | true, v -> v | _ -> fallback)
                  (fun book value -> book.Settings.[key name] <- value)
    let kept name = set name (fun (book: MacroDocument) value -> book.Settings.[key name] <- value)
    let text name (value: string) = get name (fun (_: MacroDocument) -> VString value)
    let nothing name = method name [] (fun (_: MacroDocument) _ -> VEmpty)
    HostMembers.Table<MacroDocument>("Application",
        [ get "ActiveDocument" documentObject
          get "ThisDocument" documentObject
          get "Documents" documentsObject
          get "Selection" selectionObject
          get "Options" (fun book -> optionsType.Wrap book)
          get "ActiveWindow" windowObject
          text "Name" "Microsoft Word"
          text "Version" "16.0"
          text "Build" "16.0"
          setting "UserName" (VString "User")
          text "UserInitials" "U"
          text "PathSeparator" "\\"
          setting "ScreenUpdating" (VBoolean true)
          setting "DisplayAlerts" (VLong -1)
          prop "Visible" (fun _ -> VBoolean true) (fun book value -> book.Settings.[key "Visible"] <- value)
          kept "WindowState"
          kept "DisplayStatusBar"
          setting "StatusBar" (VString "")
          text "Caption" "Microsoft Word"
          get "Application" applicationObject
          nothing "ScreenRefresh"
          nothing "Activate"
          nothing "Quit"
          nothing "Scroll"
          nothing "Wait" ]
        @ (conversions |> List.map (fun (name, factor) -> method name [] (fun _ args -> convert factor args))))

/// The objects a macro starts from, by name (Application, ActiveDocument, ThisDocument, Selection,
/// Documents, Options, ActiveWindow, the unit conversions), and the wd* constants.
let hostGlobals (book: MacroDocument) : (string * Value) list =
    let globalConversions = [ "CentimetersToPoints"; "MillimetersToPoints"; "InchesToPoints"; "LinesToPoints"; "PointsToCentimeters"; "PointsToInches" ]
    [ "Application", applicationObject book
      "ActiveDocument", documentObject book
      "ThisDocument", documentObject book
      "Selection", selectionObject book
      "Documents", documentsObject book
      "Options", optionsType.Wrap book
      "ActiveWindow", windowObject book ]
    @ [ for name, factor in conversions do
          if List.contains name globalConversions then yield name, conversionType.Wrap { Conversion = name; Factor = factor } ]
    @ [ "DoEvents", doEventsType.Wrap () ]
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
