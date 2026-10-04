module DocxReader

open XmlDsl
open OfficeDsl
open ZipInterop
open Fable.Core.JsInterop
open System

let inline private (==>) name value =
    name, box value

let private objFromFields fields : obj =
#if FABLE_COMPILER
    createObj fields
#else
    let result = System.Collections.Generic.Dictionary<string, obj>()
    for name, value in fields do
        result.[name] <- value
    box result
#endif

// DOCX domain model
type RunProperties = {
    Style: string option
    Bold: bool option
    Italic: bool option
    Underline: bool option
    UnderlineStyle: string option  // w:u@val
    UnderlineColor: string option  // w:u@color
    AllCaps: bool option
    SmallCaps: bool option
    FontSize: string option
    CharacterSpacing: string option
    FontFamily: string option
    ThemeFont: string option
    Color: string option
    ThemeColor: string option
    ThemeTint: string option
    ThemeShade: string option
    Highlight: string option
    Strike: bool option
    DoubleStrike: bool option   // w:dstrike
    Hidden: bool option         // w:vanish
    Position: string option     // w:position@val (half-points, signed)
    EmphasisMark: string option // w:em@val
    VertAlign: string option  // superscript/subscript
}

type Hyperlink = {
    Text: string
    Target: string option
    RelationshipId: string option
    Properties: RunProperties
}

type ImageData = {
    RelationshipId: string
    Linked: bool
    Name: string option
    Title: string option
    Width: float option
    Height: float option
    Description: string option
    Floating: bool
    Wrap: string option
    PositionX: float option
    PositionY: float option
    PositionHRelativeFrom: string option
    PositionVRelativeFrom: string option
    BehindDoc: bool option
    AllowOverlap: bool option
    LayoutInCell: bool option
    DistanceTop: float option
    DistanceBottom: float option
    DistanceLeft: float option
    DistanceRight: float option
    RelativeHeight: string option
    Locked: bool option
    SimplePosition: bool option
    SimplePositionX: float option
    SimplePositionY: float option
    CropLeft: string option
    CropTop: string option
    CropRight: string option
    CropBottom: string option
}

type NoteReference = {
    Id: string
    NoteType: string  // "footnote" or "endnote"
}

type BookmarkReference = {
    Id: string
    Name: string option
}

type FieldType =
    | Date
    | PageNumber
    | NumPages
    | Hyperlink of string
    | Ref of string
    | Unknown of string

type Field = {
    FieldType: FieldType
    Code: string
    Result: string option
}

type RunContent =
    | Text of string
    | Link of Hyperlink
    | Image of ImageData
    | CommentReference of string  // Comment ID
    | NoteRef of NoteReference
    | BookmarkStart of BookmarkReference
    | BookmarkEnd of string
    | FieldRef of Field
    | EquationXml of string
    | RenderedPageBreak
    | PageBreak

type Run = {
    Content: RunContent
    Properties: RunProperties
    BookmarkStart: BookmarkReference option
    BookmarkEnd: string option
    RenderedPageBreakBefore: bool
    RevisionId: string option  // For track changes
    RevisionType: string option
    RevisionAuthor: string option
    RevisionDate: string option
}

type ListInfo = {
    ListId: string
    Level: int
    NumberFormat: string option  // bullet, decimal, etc.
    LevelText: string option
    Start: int option
}

type ParagraphBorder = {
    Style: string option
    Color: string option
    Width: string option
    Space: string option
}

type ParagraphBorders = {
    Top: ParagraphBorder option
    Right: ParagraphBorder option
    Bottom: ParagraphBorder option
    Left: ParagraphBorder option
}

type HeaderFooterReference = {
    Kind: string
    RelationshipId: string
}

type SectionColumns = {
    Count: int
    Gap: float
    Separator: bool
    Widths: (float * float) list
}

type SectionGeometry = {
    PageWidth: float
    PageHeight: float
    MarginTop: float
    MarginBottom: float
    MarginLeft: float
    MarginRight: float
    HeaderDistance: float
    FooterDistance: float
    /// w:titlePg of this section.
    DifferentFirstPage: bool
}

type ParagraphProperties = {
    Style: string option
    ContentControlId: string option
    ContentControlTag: string option
    ContentControlAlias: string option
    Alignment: string option  // left, center, right, justify
    IndentLeft: string option
    IndentRight: string option
    FirstLineIndent: string option
    HangingIndent: string option
    RightTabPosition: string option
    RightTabLeader: string option
    TabStops: (string * string * string option) list  // (pos-twips, alignment, leader)
    Rtl: bool option  // w:bidi
    SpacingBefore: string option
    SpacingBeforeAuto: bool option
    SpacingAfter: string option
    SpacingAfterAuto: bool option
    LineSpacing: string option
    LineRule: string option
    OutlineLevel: string option
    KeepNext: bool option
    KeepLines: bool option
    PageBreakBefore: bool option
    WidowControl: bool option
    ContextualSpacing: bool option
    BackgroundColor: string option
    ParagraphRunProperties: RunProperties option
    ParagraphBorders: ParagraphBorders option
    ListInfo: ListInfo option
    SectionBreak: string option
    SectionColumns: SectionColumns option
    SectionGeometry: SectionGeometry option
    SectionPageNumberStart: int option
    RenderedPageBreakBefore: int option
    HeaderReferences: HeaderFooterReference list
    FooterReferences: HeaderFooterReference list
}

type Paragraph = {
    Runs: Run list
    Properties: ParagraphProperties
}

type BorderStyle = {
    BorderType: string option  // "single", "double", "dashed", etc.
    Size: string option
    Color: string option
    Space: string option
}

type CellBorders = {
    Top: BorderStyle option
    Bottom: BorderStyle option
    Left: BorderStyle option
    Right: BorderStyle option
    InsideH: BorderStyle option
    InsideV: BorderStyle option
}

type Shading = {
    Fill: string option
    Color: string option
    Pattern: string option
}

type TableCellProperties = {
    Borders: CellBorders option
    Shading: Shading option
    Width: string option
    WidthType: string option
    VerticalAlignment: string option
    PaddingTop: string option
    PaddingRight: string option
    PaddingBottom: string option
    PaddingLeft: string option
    ColSpan: int option
    VerticalMerge: string option
}

type TableStyleProperties = {
    Borders: CellBorders option
    Shading: Shading option
    PaddingTop: string option
    PaddingRight: string option
    PaddingBottom: string option
    PaddingLeft: string option
}

type TableRowStyleProperties = {
    IsHeader: bool option
    CanSplit: bool option
}

type Table = {
    Rows: TableRow list
    StyleId: string option
    Look: TableLook option
    Alignment: string option
    Properties: TableStyleProperties option
    Borders: CellBorders option
    Width: string option
    WidthType: string option
    CellSpacing: string option
    LayoutMode: string option
    GridColumns: string list
    Floating: bool
}
and TableLook = {
    FirstRow: bool
    LastRow: bool
    FirstColumn: bool
    LastColumn: bool
    HorizontalBanding: bool
    VerticalBanding: bool
}
and TableRow = {
    Cells: TableCell list
    Height: string option
    HeightRule: string option
    IsHeader: bool
    CanSplit: bool option
    CellSpacing: string option
    PaddingTop: string option
    PaddingRight: string option
    PaddingBottom: string option
    PaddingLeft: string option
    GridBefore: int option
    GridAfter: int option
    WidthBefore: string option
    WidthAfter: string option
    BordersOverride: CellBorders option
}
and TableCell = {
    Paragraphs: Paragraph list
    NestedTables: Table list
    // For each nested table, how many cell paragraphs precede it in document order.
    NestedTableAnchors: int list
    Properties: TableCellProperties option
}

type DocumentContent =
    | Para of Paragraph
    | Tbl of Table

type HeaderFooter = {
    /// Every paragraph of the part, tables flattened (kept for callers that only want text).
    Content: Paragraph list
    /// The part's blocks in order, with tables kept as tables.
    Blocks: DocumentContent list
    Type: string  // "header", "footer", "first", "even"
    RelationshipId: string option
}

type StyleDefinition = {
    Id: string
    Name: string option
    StyleType: string
    IsDefault: bool
    BasedOn: string option
    Next: string option
    Linked: string option
    IsQuickStyle: bool
    Aliases: string option
    UiPriority: string option
    SemiHidden: bool
    UnhideWhenUsed: bool
    RunProperties: RunProperties option
    ParagraphProperties: ParagraphProperties option
    TableProperties: TableStyleProperties option
    FirstRowParagraphProperties: ParagraphProperties option
    FirstRowTableRowProperties: TableRowStyleProperties option
    FirstRowTableProperties: TableStyleProperties option
    LastRowTableProperties: TableStyleProperties option
    FirstColumnTableProperties: TableStyleProperties option
    LastColumnTableProperties: TableStyleProperties option
    Band1HorizontalTableProperties: TableStyleProperties option
    Band2HorizontalTableProperties: TableStyleProperties option
    Band1VerticalTableProperties: TableStyleProperties option
    Band2VerticalTableProperties: TableStyleProperties option
    FirstRowTableCellProperties: TableCellProperties option
    LastRowTableCellProperties: TableCellProperties option
    FirstColumnTableCellProperties: TableCellProperties option
    LastColumnTableCellProperties: TableCellProperties option
    Band1HorizontalTableCellProperties: TableCellProperties option
    Band2HorizontalTableCellProperties: TableCellProperties option
    Band1VerticalTableCellProperties: TableCellProperties option
    Band2VerticalTableCellProperties: TableCellProperties option
}

type Theme = {
    MajorFont: string option
    MinorFont: string option
    Colors: Map<string, string>
}

type DocumentProperties = {
    Title: string option
    Subject: string option
    Creator: string option
    Created: string option
    Modified: string option
    Description: string option
}

// Advanced features
type Comment = {
    Id: string
    Author: string option
    Date: string option
    Text: string
    InitialText: string option
}

type RevisionType =
    | Insert
    | Delete
    | FormatChange
    | MoveFrom
    | MoveTo

type Revision = {
    Id: string
    Author: string option
    Date: string option
    RevisionType: RevisionType
    Text: string
}

type Note = {
    Id: string
    Content: Paragraph list
}

/// One segment of a custom (a:custGeom) path. Op is M/L/C/Q/Z; Pts are the flat
/// coordinate list in the path's own w×h space.
type ShapePathSeg = { Op: string; Pts: float list }

type ShapeGradientStop = { Offset: float; Color: string }
type ShapeGradient = { Stops: ShapeGradientStop list; Angle: float; Radial: bool }

type Shape = {
    Id: string option
    Type: string option  // "textBox", "rectangle", "ellipse", etc.
    Content: Paragraph list
    Width: float option
    Height: float option
    Position: (float * float) option  // (x, y)
    FillColor: string option
    BorderColor: string option
    BorderWidth: float option
    /// a:prstGeom@prst preset name (e.g. "ellipse", "roundRect").
    Preset: string option
    /// a:custGeom path segments + their coordinate-space size (EMU).
    CustomPath: ShapePathSeg list
    CustomW: float option
    CustomH: float option
    Gradient: ShapeGradient option
    /// Shape rotation in degrees (a:xfrm@rot, 60000ths -> degrees).
    Rotation: float
}

type PageSetup = {
    PageWidth: float option
    PageHeight: float option
    MarginLeft: float option
    MarginRight: float option
    MarginTop: float option
    MarginBottom: float option
    HeaderDistance: float option
    FooterDistance: float option
    DifferentFirstPage: bool option
    EvenAndOddHeaders: bool option
    Columns: SectionColumns option
    SectionStart: string option   // body sectPr w:type (e.g. "continuous")
}

type NumberingDefinition = {
    NumId: string
    Level: int
    ParagraphStyle: string option
    NumberFormat: string
    LevelText: string option
    Start: int option
    IndentLeft: string option
    HangingIndent: string option
    FirstLineIndent: string option
    RightTabPosition: string option
}

type WordDocument = {
    Content: DocumentContent list
    Relationships: Relationship list
    Styles: StyleDefinition list
    Theme: Theme option
    Headers: HeaderFooter list
    Footers: HeaderFooter list
    Properties: DocumentProperties option
    Comments: Comment list
    Revisions: Revision list
    Footnotes: Note list
    Endnotes: Note list
    Images: Map<string, string>  // relationshipId -> data URL
    Shapes: Shape list
    PageSetup: PageSetup option
    NumberingDefinitions: NumberingDefinition list
}

// DSL for DOCX-specific XML patterns
module Docx =
    // Namespace-aware element access
    let inline wElem name node = Ns.childNs name node
    let inline wElems name node = Ns.descendantsNs name node

    let private appendEscapedXml (builder: System.Text.StringBuilder) (text: string) =
        if not (String.IsNullOrEmpty text) then
            let mutable segmentStart = 0
            for i = 0 to text.Length - 1 do
                let replacement =
                    match text.[i] with
                    | '&' -> ValueSome "&amp;"
                    | '<' -> ValueSome "&lt;"
                    | '>' -> ValueSome "&gt;"
                    | '"' -> ValueSome "&quot;"
                    | '\'' -> ValueSome "&apos;"
                    | _ -> ValueNone
                match replacement with
                | ValueSome escaped ->
                    if i > segmentStart then
                        builder.Append(text, segmentStart, i - segmentStart) |> ignore
                    builder.Append(escaped) |> ignore
                    segmentStart <- i + 1
                | ValueNone -> ()
            if segmentStart < text.Length then
                builder.Append(text, segmentStart, text.Length - segmentStart) |> ignore

    let rec private appendXmlTo (builder: System.Text.StringBuilder) node =
        match node with
        | XmlDsl.Text text ->
            appendEscapedXml builder text
        | XmlDsl.Empty -> ()
        | XmlDsl.Element(name, attrs, children) ->
            builder.Append('<').Append(name) |> ignore
            for KeyValue(key, value) in attrs do
                builder.Append(' ').Append(key).Append("=\"") |> ignore
                appendEscapedXml builder value
                builder.Append('"') |> ignore
            match children with
            | [] ->
                builder.Append("/>") |> ignore
            | _ ->
                builder.Append('>') |> ignore
                for child in children do
                    appendXmlTo builder child
                builder.Append("</").Append(name).Append('>') |> ignore

    let xmlToString node =
        let builder = System.Text.StringBuilder()
        appendXmlTo builder node
        builder.ToString()

    let attr attrName node =
        Xml.attr attrName node
        |> Option.orElseWith (fun () -> Xml.attr ("w:" + attrName) node)

    let private attrByLocalName localName node =
        match node with
        | XmlDsl.Element(_, attrs, _) ->
            attrs
            |> Map.toList
            |> List.tryPick (fun (name, value) ->
                if Ns.stripNs name = localName then Some value else None)
        | _ -> None

    // Property extraction DSL
    let hasProp propName node =
        wElem propName node |> Option.isSome

    let getPropVal propName attrName node =
        wElem propName node
        |> Option.bind (attr attrName)

    let onOff propName node =
        wElem propName node
        |> Option.map (fun prop ->
            match attr "val" prop |> Option.map (fun value -> value.ToLowerInvariant()) with
            | Some "0"
            | Some "false"
            | Some "off" -> false
            | _ -> true)

    // Parse a sectPr's <w:cols> into a SectionColumns (twips/15 = px).
    let parseColumns (sectPr: XmlNode) : SectionColumns option =
        let twips (s: string option) = s |> Option.bind (fun v -> match System.Double.TryParse v with | true, n -> Some (n / 15.0) | _ -> None)
        Ns.childNs "cols" sectPr
        |> Option.bind (fun cols ->
            let num = attr "num" cols |> Option.bind (fun v -> match System.Int32.TryParse v with | true, n -> Some n | _ -> None) |> Option.defaultValue 1
            let space = twips (attr "space" cols) |> Option.defaultValue 36.0
            // w:sep is an attribute of w:cols (onOff looks for a child element, so
            // the separator line was never read).
            let sep =
                match attr "sep" cols |> Option.map (fun value -> value.ToLowerInvariant()) with
                | Some "1" | Some "true" | Some "on" -> true
                | _ -> false
            let colWidths =
                Ns.descendantsNs "col" cols
                |> List.choose (fun col -> twips (attr "w" col) |> Option.map (fun w -> w, twips (attr "space" col) |> Option.defaultValue 0.0))
            let count = if colWidths.Length >= 2 then colWidths.Length else num
            if count <= 1 then None
            else Some { Count = count; Gap = space; Separator = sep; Widths = (if colWidths.Length >= 2 then colWidths else []) })

    // Parse a sectPr's pgSz/pgMar into a SectionGeometry (twips/15 = px).
    let parseSectionGeometry (sectPr: XmlNode) : SectionGeometry option =
        let twips (s: string option) = s |> Option.bind (fun v -> match System.Double.TryParse v with | true, n -> Some (n / 15.0) | _ -> None)
        match Ns.childNs "pgSz" sectPr, Ns.childNs "pgMar" sectPr with
        | None, None -> None
        | pgSz, pgMar ->
            Some {
                PageWidth = pgSz |> Option.bind (attr "w") |> twips |> Option.defaultValue 816.0
                PageHeight = pgSz |> Option.bind (attr "h") |> twips |> Option.defaultValue 1056.0
                MarginTop = pgMar |> Option.bind (attr "top") |> twips |> Option.defaultValue 96.0
                MarginBottom = pgMar |> Option.bind (attr "bottom") |> twips |> Option.defaultValue 96.0
                MarginLeft = pgMar |> Option.bind (attr "left") |> twips |> Option.defaultValue 96.0
                MarginRight = pgMar |> Option.bind (attr "right") |> twips |> Option.defaultValue 96.0
                HeaderDistance = pgMar |> Option.bind (attr "header") |> twips |> Option.defaultValue 48.0
                FooterDistance = pgMar |> Option.bind (attr "footer") |> twips |> Option.defaultValue 48.0
                DifferentFirstPage = onOff "titlePg" sectPr |> Option.defaultValue false
            }

    // Run properties extraction with defaults
    let runProps node =
        let rPr = wElem "rPr" node
        match rPr with
        | Some props ->
            {
                Style = getPropVal "rStyle" "val" props
                Bold = onOff "b" props
                Italic = onOff "i" props
                // <w:u w:val="none"/> switches an inherited underline off.
                Underline =
                    (match getPropVal "u" "val" props with
                     | Some value when value.ToLowerInvariant() = "none" -> Some false
                     | _ -> onOff "u" props)
                UnderlineStyle = getPropVal "u" "val" props
                UnderlineColor = getPropVal "u" "color" props
                AllCaps = onOff "caps" props
                SmallCaps = onOff "smallCaps" props
                FontSize = getPropVal "sz" "val" props
                CharacterSpacing = getPropVal "spacing" "val" props
                FontFamily = getPropVal "rFonts" "ascii" props
                ThemeFont = getPropVal "rFonts" "asciiTheme" props
                Color = getPropVal "color" "val" props
                ThemeColor = getPropVal "color" "themeColor" props
                ThemeTint = getPropVal "color" "themeTint" props
                ThemeShade = getPropVal "color" "themeShade" props
                Highlight = getPropVal "highlight" "val" props
                Strike = onOff "strike" props
                DoubleStrike = onOff "dstrike" props
                Hidden = onOff "vanish" props
                Position = getPropVal "position" "val" props
                EmphasisMark = getPropVal "em" "val" props
                VertAlign = getPropVal "vertAlign" "val" props
            }
        | None ->
            {
                Style = None
                Bold = None
                Italic = None
                Underline = None
                UnderlineStyle = None
                UnderlineColor = None
                AllCaps = None
                SmallCaps = None
                FontSize = None
                CharacterSpacing = None
                FontFamily = None
                ThemeFont = None
                Color = None
                ThemeColor = None
                ThemeTint = None
                ThemeShade = None
                Highlight = None
                Strike = None
                DoubleStrike = None
                Hidden = None
                Position = None
                EmphasisMark = None
                VertAlign = None
            }

    // Paragraph properties extraction
    let private headerFooterReferences name props =
        wElems name props
        |> List.choose (fun node ->
            Xml.attr "r:id" node
            |> Option.map (fun id -> {
                Kind = attr "type" node |> Option.defaultValue "default"
                RelationshipId = id
            }))

    let private paragraphBorder name props =
        wElem name props
        |> Option.map (fun node -> {
            Style = attr "val" node
            Color = attr "color" node
            Width = attr "sz" node
            Space = attr "space" node
        })

    let paragraphBorders props =
        wElem "pBdr" props
        |> Option.map (fun borders -> {
            Top = paragraphBorder "top" borders
            Right = paragraphBorder "right" borders
            Bottom = paragraphBorder "bottom" borders
            Left = paragraphBorder "left" borders
        })

    let rightTabStop props =
        wElem "tabs" props
        |> Option.bind (fun tabs ->
            Xml.children tabs
            |> List.tryPick (fun tab ->
                if Xml.name tab |> Option.exists (fun name -> Ns.stripNs name = "tab") && attr "val" tab = Some "right" then
                    Some (attr "pos" tab, attr "leader" tab)
                else
                    None))

    /// All real tab stops (pos-twips, alignment, leader). Skips val="clear".
    let allTabStops props : (string * string * string option) list =
        wElem "tabs" props
        |> Option.map (fun tabs ->
            Xml.children tabs
            |> List.choose (fun tab ->
                let isTab = Xml.name tab |> Option.exists (fun name -> Ns.stripNs name = "tab")
                let align = attr "val" tab |> Option.defaultValue "left"
                match (if isTab && align <> "clear" then attr "pos" tab else None) with
                | Some pos -> Some (pos, align, attr "leader" tab)
                | None -> None))
        |> Option.defaultValue []

    let paragraphProps node =
        let pPr = wElem "pPr" node
        match pPr with
        | Some props ->
            let rightTab = rightTabStop props
            let frameAlignment =
                wElem "framePr" props
                |> Option.bind (fun frame -> attr "xAlign" frame)
                |> Option.bind (fun value ->
                    match value.ToLowerInvariant() with
                    | "left"
                    | "center"
                    | "right" -> Some value
                    | _ -> None)
            let listInfo =
                wElem "numPr" props
                |> Option.bind (fun numPr ->
                    getPropVal "numId" "val" numPr
                    |> Option.map (fun numId -> {
                        ListId = numId
                        Level = getPropVal "ilvl" "val" numPr |> Option.map int |> Option.defaultValue 0
                        NumberFormat = None  // Resolved from numbering.xml later
                        LevelText = None
                        Start = None
                    }))

            let sectPr = wElem "sectPr" props
            {
                Style = getPropVal "pStyle" "val" props
                ContentControlId = None
                ContentControlTag = None
                ContentControlAlias = None
                Alignment = getPropVal "jc" "val" props |> Option.orElse frameAlignment
                IndentLeft = getPropVal "ind" "left" props
                IndentRight = getPropVal "ind" "right" props
                FirstLineIndent = getPropVal "ind" "firstLine" props
                HangingIndent = getPropVal "ind" "hanging" props
                RightTabPosition = rightTab |> Option.bind fst
                RightTabLeader = rightTab |> Option.bind snd
                TabStops = allTabStops props
                Rtl = onOff "bidi" props
                SpacingBefore = getPropVal "spacing" "before" props
                SpacingBeforeAuto = getPropVal "spacing" "beforeAutospacing" props |> Option.map (fun value -> value = "1" || value.ToLowerInvariant() = "true")
                SpacingAfter = getPropVal "spacing" "after" props
                SpacingAfterAuto = getPropVal "spacing" "afterAutospacing" props |> Option.map (fun value -> value = "1" || value.ToLowerInvariant() = "true")
                LineSpacing = getPropVal "spacing" "line" props
                LineRule = getPropVal "spacing" "lineRule" props
                OutlineLevel = getPropVal "outlineLvl" "val" props
                KeepNext = onOff "keepNext" props
                KeepLines = onOff "keepLines" props
                PageBreakBefore = onOff "pageBreakBefore" props
                WidowControl = onOff "widowControl" props
                ContextualSpacing = onOff "contextualSpacing" props
                BackgroundColor = getPropVal "shd" "fill" props
                ParagraphRunProperties = wElem "rPr" props |> Option.map (fun _ -> runProps props)
                ParagraphBorders = paragraphBorders props
                ListInfo = listInfo
                SectionBreak = sectPr |> Option.map (fun section -> getPropVal "type" "val" section |> Option.defaultValue "nextPage")
                SectionColumns = sectPr |> Option.bind parseColumns
                SectionGeometry = sectPr |> Option.bind parseSectionGeometry
                SectionPageNumberStart = sectPr |> Option.bind (fun section -> getPropVal "pgNumType" "start" section) |> Option.bind (fun value -> match System.Int32.TryParse(value) with | true, parsed -> Some parsed | false, _ -> None)
                RenderedPageBreakBefore = None
                HeaderReferences = sectPr |> Option.map (headerFooterReferences "headerReference") |> Option.defaultValue []
                FooterReferences = sectPr |> Option.map (headerFooterReferences "footerReference") |> Option.defaultValue []
            }
        | None ->
            {
                Style = None
                ContentControlId = None
                ContentControlTag = None
                ContentControlAlias = None
                Alignment = None
                IndentLeft = None
                IndentRight = None
                FirstLineIndent = None
                HangingIndent = None
                RightTabPosition = None
                RightTabLeader = None
                TabStops = []
                Rtl = None
                SpacingBefore = None
                SpacingBeforeAuto = None
                SpacingAfter = None
                SpacingAfterAuto = None
                LineSpacing = None
                LineRule = None
                OutlineLevel = None
                KeepNext = None
                KeepLines = None
                PageBreakBefore = None
                WidowControl = None
                ContextualSpacing = None
                BackgroundColor = None
                ParagraphRunProperties = None
                ParagraphBorders = None
                ListInfo = None
                SectionBreak = None
                SectionColumns = None
                SectionGeometry = None
                SectionPageNumberStart = None
                RenderedPageBreakBefore = None
                HeaderReferences = []
                FooterReferences = []
            }

    // Image extraction
    let imageData (rNode: XmlNode) : ImageData option =
        // Store DOCX extents as 96-DPI canvas pixels. The Word print path
        // converts these pixels to points when emitting PDF wire.
        let emuToPixels emu = emu / 9525.0
        let ptToPixels pt = pt * 96.0 / 72.0
        let parseFloat (value: string) =
            match System.Double.TryParse(value, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture) with
            | true, parsed -> Some parsed
            | false, _ -> None
        let emptyImage relationshipId linked name title description width height = {
            RelationshipId = relationshipId
            Linked = linked
            Name = name
            Title = title
            Width = width
            Height = height
            Description = description
            Floating = false
            Wrap = None
            PositionX = None
            PositionY = None
            PositionHRelativeFrom = None
            PositionVRelativeFrom = None
            BehindDoc = None
            AllowOverlap = None
            LayoutInCell = None
            DistanceTop = None
            DistanceBottom = None
            DistanceLeft = None
            DistanceRight = None
            RelativeHeight = None
            Locked = None
            SimplePosition = None
            SimplePositionX = None
            SimplePositionY = None
            CropLeft = None
            CropTop = None
            CropRight = None
            CropBottom = None
        }
        let drawingImage =
            wElem "drawing" rNode
            |> Option.bind (fun drawing ->
            let extent = Ns.descendantsNs "extent" drawing |> List.tryHead
            let width = extent |> Option.bind (Xml.attr "cx") |> Option.map (fun cx -> emuToPixels (float cx))
            let height = extent |> Option.bind (Xml.attr "cy") |> Option.map (fun cy -> emuToPixels (float cy))
            let docPr = Ns.descendantsNs "docPr" drawing |> List.tryHead
            let name = docPr |> Option.bind (Xml.attr "name")
            let title = docPr |> Option.bind (Xml.attr "title")
            let desc = docPr |> Option.bind (Xml.attr "descr")
            let blip = Ns.descendantsNs "blip" drawing |> List.tryHead
            match blip with
            | Some b ->
                let embedId = Xml.attr "r:embed" b |> Option.orElseWith (fun () -> attrByLocalName "embed" b)
                let linkId = Xml.attr "r:link" b |> Option.orElseWith (fun () -> attrByLocalName "link" b)
                embedId |> Option.orElse linkId
                |> Option.map (fun rId ->
                    let anchor = Ns.childNs "anchor" drawing
                    let floating = anchor.IsSome
                    let wrap =
                        anchor
                        |> Option.bind (fun a ->
                            Xml.children a
                            |> List.tryPick (fun child ->
                                Xml.name child
                                |> Option.map Ns.stripNs
                                |> Option.bind (function
                                    | "wrapNone" -> Some "none"
                                    | "wrapSquare" -> Some "square"
                                    | "wrapTight" -> Some "tight"
                                    | "wrapThrough" -> Some "through"
                                    | "wrapTopAndBottom" -> Some "topAndBottom"
                                    | _ -> None)))
                    let positionOffset localName =
                        anchor
                        |> Option.bind (Ns.childNs localName)
                        |> Option.bind (Ns.childNs "posOffset")
                        |> Option.bind Xml.text
                        |> Option.bind (fun value ->
                            match System.Double.TryParse(value) with
                            | true, parsed -> Some (emuToPixels parsed)
                            | false, _ -> None)
                    let positionRelativeFrom localName =
                        anchor
                        |> Option.bind (Ns.childNs localName)
                        |> Option.bind (Xml.attr "relativeFrom")
                    let anchorBool attrName =
                        anchor
                        |> Option.bind (Xml.attr attrName)
                        |> Option.map (fun value -> value = "1" || value.ToLowerInvariant() = "true")
                    let anchorDistance attrName =
                        anchor
                        |> Option.bind (Xml.attr attrName)
                        |> Option.bind (fun value ->
                            match System.Double.TryParse(value) with
                            | true, parsed -> Some (emuToPixels parsed)
                            | false, _ -> None)
                    let simplePositionOffset attrName =
                        anchor
                        |> Option.bind (Ns.childNs "simplePos")
                        |> Option.bind (Xml.attr attrName)
                        |> Option.bind (fun value ->
                            match System.Double.TryParse(value) with
                            | true, parsed -> Some (emuToPixels parsed)
                            | false, _ -> None)

                    let positionX = positionOffset "positionH"
                    let positionY = positionOffset "positionV"
                    let positionHRelativeFrom = positionRelativeFrom "positionH"
                    let positionVRelativeFrom = positionRelativeFrom "positionV"
                    let behindDoc = anchorBool "behindDoc"
                    let allowOverlap = anchorBool "allowOverlap"
                    let layoutInCell = anchorBool "layoutInCell"
                    let distanceTop = anchorDistance "distT"
                    let distanceBottom = anchorDistance "distB"
                    let distanceLeft = anchorDistance "distL"
                    let distanceRight = anchorDistance "distR"
                    let relativeHeight = anchor |> Option.bind (Xml.attr "relativeHeight")
                    let locked = anchorBool "locked"
                    let simplePosition = anchorBool "simplePos"
                    let simplePositionX = simplePositionOffset "x"
                    let simplePositionY = simplePositionOffset "y"
                    let srcRect = Ns.descendantsNs "srcRect" drawing |> List.tryHead
                    let cropLeft = srcRect |> Option.bind (Xml.attr "l")
                    let cropTop = srcRect |> Option.bind (Xml.attr "t")
                    let cropRight = srcRect |> Option.bind (Xml.attr "r")
                    let cropBottom = srcRect |> Option.bind (Xml.attr "b")
                    {
                        RelationshipId = rId
                        Linked = embedId.IsNone && linkId.IsSome
                        Name = name
                        Title = title
                        Width = width
                        Height = height
                        Description = desc
                        Floating = floating
                        Wrap = wrap
                        PositionX = positionX
                        PositionY = positionY
                        PositionHRelativeFrom = positionHRelativeFrom
                        PositionVRelativeFrom = positionVRelativeFrom
                        BehindDoc = behindDoc
                        AllowOverlap = allowOverlap
                        LayoutInCell = layoutInCell
                        DistanceTop = distanceTop
                        DistanceBottom = distanceBottom
                        DistanceLeft = distanceLeft
                        DistanceRight = distanceRight
                        RelativeHeight = relativeHeight
                        Locked = locked
                        SimplePosition = simplePosition
                        SimplePositionX = simplePositionX
                        SimplePositionY = simplePositionY
                        CropLeft = cropLeft
                        CropTop = cropTop
                        CropRight = cropRight
                        CropBottom = cropBottom
                    })
            | None ->
                if Ns.descendantsNs "chart" drawing |> List.isEmpty |> not then
                    Some {
                        RelationshipId = "__unsupported-chart"
                        Linked = false
                        Name = name
                        Title = title |> Option.orElse (Some "Unsupported chart")
                        Width = width
                        Height = height
                        Description = desc |> Option.orElse (Some "Chart")
                        Floating = false
                        Wrap = None
                        PositionX = None
                        PositionY = None
                        PositionHRelativeFrom = None
                        PositionVRelativeFrom = None
                        BehindDoc = None
                        AllowOverlap = None
                        LayoutInCell = None
                        DistanceTop = None
                        DistanceBottom = None
                        DistanceLeft = None
                        DistanceRight = None
                        RelativeHeight = None
                        Locked = None
                        SimplePosition = None
                        SimplePositionX = None
                        SimplePositionY = None
                        CropLeft = None
                        CropTop = None
                        CropRight = None
                        CropBottom = None
                    }
                else None)
        let vmlImage =
            wElem "pict" rNode
            |> Option.bind (fun pict ->
                let image = Ns.descendantsNs "imagedata" pict |> List.tryHead
                image
                |> Option.bind (fun image ->
                    let embedId = Xml.attr "r:id" image |> Option.orElseWith (fun () -> attrByLocalName "id" image)
                    embedId
                    |> Option.map (fun rId ->
                        let shape = Ns.descendantsNs "shape" pict |> List.tryHead
                        let style = shape |> Option.bind (Xml.attr "style")
                        let styleValue name =
                            style
                            |> Option.bind (fun value ->
                                value.Split(';')
                                |> Array.tryPick (fun part ->
                                    let pieces = part.Split([| ':' |], 2)
                                    if pieces.Length = 2 && pieces.[0].Trim().Equals(name, System.StringComparison.OrdinalIgnoreCase) then Some (pieces.[1].Trim())
                                    else None))
                        let stylePoints name =
                            styleValue name
                            |> Option.bind (fun value ->
                                Some (if value.EndsWith("pt", System.StringComparison.OrdinalIgnoreCase) then value.Substring(0, value.Length - 2) else value))
                            |> Option.bind parseFloat
                            |> Option.map ptToPixels
                        emptyImage rId false (shape |> Option.bind (fun value -> Xml.attr "id" value)) (Xml.attr "o:title" image) (shape |> Option.bind (fun value -> Xml.attr "alt" value)) (stylePoints "width") (stylePoints "height"))))
        drawingImage |> Option.orElseWith (fun () -> vmlImage)

    // Comment reference extraction
    let commentReference (rNode: XmlNode) : string option =
        wElem "commentReference" rNode
        |> Option.bind (attr "id")

    // Footnote/Endnote reference extraction
    let noteReference (rNode: XmlNode) : NoteReference option =
        let footnoteRef = wElem "footnoteReference" rNode
        let endnoteRef = wElem "endnoteReference" rNode

        match footnoteRef, endnoteRef with
        | Some fn, _ ->
            attr "id" fn
            |> Option.map (fun id -> { Id = id; NoteType = "footnote" })
        | _, Some en ->
            attr "id" en
            |> Option.map (fun id -> { Id = id; NoteType = "endnote" })
        | _ -> None

    // Field extraction
    let fieldTypeFromCode (code: string) =
        let normalized = code.Trim()
        let upper = normalized.ToUpperInvariant()
        if upper.StartsWith("DATE") then Date
        elif upper.StartsWith("PAGE") then PageNumber
        elif upper.StartsWith("NUMPAGES") then NumPages
        elif upper.StartsWith("HYPERLINK") then Hyperlink normalized
        elif upper.StartsWith("REF ") then Ref (normalized.Substring(4).Trim())
        else Unknown normalized

    let simpleField (fldNode: XmlNode) : Run option =
        // OOXML serialises the instruction code on the w: namespaced attribute
        // ("w:instr"). Xml.attr is a plain-string lookup on the attrs map, so
        // we have to ask for the prefixed form for documents that ship the
        // namespace declaration this way. Fall back to the bare local name
        // for parsers/serialisers that strip the prefix.
        (Xml.attr "w:instr" fldNode |> Option.orElseWith (fun () -> Xml.attr "instr" fldNode))
        |> Option.map (fun code ->
            let result =
                wElems "r" fldNode
                |> List.choose (fun run -> wElem "t" run |> Option.bind Xml.text)
                |> String.concat ""
            let props =
                wElems "r" fldNode
                |> List.tryHead
                |> Option.map runProps
                |> Option.defaultValue (runProps fldNode)
            { Content = FieldRef { FieldType = fieldTypeFromCode code; Code = code; Result = if result = "" then None else Some result }
              Properties = props
              BookmarkStart = None
              BookmarkEnd = None
              RenderedPageBreakBefore = false
              RevisionId = None
              RevisionType = None
              RevisionAuthor = None
              RevisionDate = None })

    let fieldData (rNode: XmlNode) : Field option =
        wElem "fldChar" rNode
        |> Option.bind (fun fld ->
            let fldType = attr "fldCharType" fld
            match fldType with
            | Some "begin" ->
                // Look for field code in next runs
                Some {
                    FieldType = Unknown ""
                    Code = ""
                    Result = None
                }
            | _ -> None)

    let textContent (rNode: XmlNode) =
        Xml.children rNode
        |> List.choose (fun child ->
            Xml.name child
            |> Option.bind (fun name ->
                match Ns.stripNs name with
                | "tab" -> Some "\t"
                | "t"
                | "delText" -> Xml.text child
                | _ -> None))
        |> String.concat ""
        |> function
           | "" -> None
           | text -> Some text

    type private RunToken =
        | RunToken of Run
        | FieldBeginToken of RunProperties
        | FieldSeparateToken
        | FieldEndToken
        | FieldInstructionToken of string

    type private ComplexFieldState = {
        Properties: RunProperties
        CodeParts: ResizeArray<string>
        ResultRuns: ResizeArray<Run>
        mutable CollectingResult: bool
    }

    let private textFromRun (run: Run) =
        match run.Content with
        | Text text -> text
        | Link link -> link.Text
        | FieldRef field -> field.Result |> Option.defaultValue field.Code
        | Image img -> img.Description |> Option.map (sprintf "[Image: %s]") |> Option.defaultValue "[Image]"
        | NoteRef note -> sprintf "[%s %s]" note.NoteType note.Id
        | EquationXml _ -> "[Equation]"
        | CommentReference _
        | BookmarkStart _
        | BookmarkEnd _
        | RenderedPageBreak
        | PageBreak -> ""

    let private appendRun (output: Run list) (stack: ComplexFieldState list) (run: Run) =
        match stack with
        | [] -> run :: output, stack
        | current :: rest ->
            current.ResultRuns.Add run
            output, stack

    let private completeFieldRun (field: ComplexFieldState) : Run =
        let codeBuilder = System.Text.StringBuilder()
        for part in field.CodeParts do
            codeBuilder.Append part |> ignore
        let code = codeBuilder.ToString().Trim()
        let resultBuilder = System.Text.StringBuilder()
        for run in field.ResultRuns do
            resultBuilder.Append(textFromRun run) |> ignore
        let result = resultBuilder.ToString()
        let props =
            if field.ResultRuns.Count > 0 then field.ResultRuns.[0].Properties
            else field.Properties
        let firstResult =
            if field.ResultRuns.Count > 0 then Some field.ResultRuns.[0]
            else None
        {
            Content = FieldRef {
                FieldType = fieldTypeFromCode code
                Code = code
                Result = if result = "" then None else Some result
            }
            Properties = props
            BookmarkStart = firstResult |> Option.bind (fun run -> run.BookmarkStart)
            BookmarkEnd = firstResult |> Option.bind (fun run -> run.BookmarkEnd)
            RenderedPageBreakBefore = firstResult |> Option.exists (fun run -> run.RenderedPageBreakBefore)
            RevisionId = firstResult |> Option.bind (fun run -> run.RevisionId)
            RevisionType = firstResult |> Option.bind (fun run -> run.RevisionType)
            RevisionAuthor = firstResult |> Option.bind (fun run -> run.RevisionAuthor)
            RevisionDate = firstResult |> Option.bind (fun run -> run.RevisionDate)
        }

    let private flattenIncompleteField (output: Run list) (stack: ComplexFieldState list) =
        match stack with
        | [] -> output, stack
        | current :: rest ->
            let mutable currentOutput = output
            let mutable currentStack = rest
            for run in current.ResultRuns do
                let nextOutput, nextStack = appendRun currentOutput currentStack run
                currentOutput <- nextOutput
                currentStack <- nextStack
            let output, rest = currentOutput, currentStack
            output, rest

    let private collapseComplexFieldTokens (tokens: RunToken list) =
        let mutable output: Run list = []
        let mutable stack: ComplexFieldState list = []

        for token in tokens do
            match token with
            | RunToken run ->
                let newOutput, newStack = appendRun output stack run
                output <- newOutput
                stack <- newStack
            | FieldBeginToken props ->
                stack <- { Properties = props; CodeParts = ResizeArray<string>(); ResultRuns = ResizeArray<Run>(); CollectingResult = false } :: stack
            | FieldInstructionToken code ->
                match stack with
                | current :: rest when not current.CollectingResult ->
                    current.CodeParts.Add code
                | _ -> ()
            | FieldSeparateToken ->
                match stack with
                | current :: _ -> current.CollectingResult <- true
                | [] -> ()
            | FieldEndToken ->
                match stack with
                | current :: rest ->
                    stack <- rest
                    let fieldRun = completeFieldRun current
                    let newOutput, newStack = appendRun output stack fieldRun
                    output <- newOutput
                    stack <- newStack
                | [] -> ()

        while not stack.IsEmpty do
            let newOutput, newStack = flattenIncompleteField output stack
            output <- newOutput
            stack <- newStack

        List.rev output

    let private attachRunMarkers (runs: Run list) =
        let mutable pendingStart: BookmarkReference option = None
        let mutable pendingRenderedPageBreak = false
        let output = ResizeArray<Run>()

        let attachStart run =
            let withBookmark =
                match pendingStart with
                | None -> run
                | Some first -> { run with BookmarkStart = Some first }
            if pendingRenderedPageBreak then { withBookmark with RenderedPageBreakBefore = true } else withBookmark

        for run in runs do
            match run.Content with
            | BookmarkStart bookmark ->
                if pendingStart.IsNone then
                    pendingStart <- Some bookmark
            | BookmarkEnd id ->
                if output.Count > 0 then
                    let previous = output.[output.Count - 1]
                    output.[output.Count - 1] <- { previous with BookmarkEnd = Some id }
                pendingStart <- None
            | RenderedPageBreak ->
                pendingRenderedPageBreak <- true
            | _ ->
                output.Add(attachStart run)
                pendingStart <- None
                pendingRenderedPageBreak <- false

        output |> Seq.toList

    let private clearTocHyperlinkStyle (run: Run) =
        let props =
            match run.Properties.Style with
            | Some style when style = "Hyperlink" -> { run.Properties with Style = None }
            | _ -> run.Properties
        let content =
            match run.Content with
            | Link link ->
                let linkProps =
                    match link.Properties.Style with
                    | Some style when style = "Hyperlink" -> { link.Properties with Style = None }
                    | _ -> link.Properties
                Link { link with Properties = linkProps }
            | _ -> run.Content
        { run with Content = content; Properties = props }

    let private normalizeTocRuns (props: ParagraphProperties) (runs: Run list) =
        match props.Style with
        | Some style when style.StartsWith("TOC", StringComparison.OrdinalIgnoreCase) ->
            runs |> List.map clearTocHyperlinkStyle
        | _ -> runs

    // Run extraction preserves child order. In WordprocessingML, w:br is a run
    // child at the exact break location, not a paragraph-level flag.
    let private tokensFromRun (rNode: XmlNode) : RunToken list =
        let props = runProps rNode
        let revId = wElem "rPr" rNode
                    |> Option.bind (wElem "rPrChange")
                    |> Option.bind (attr "id")
        let makeRun content =
            { Content = content
              Properties = props
              BookmarkStart = None
              BookmarkEnd = None
              RenderedPageBreakBefore = false
              RevisionId = revId
              RevisionType = None
              RevisionAuthor = None
              RevisionDate = None }

        let childTokens =
            Xml.children rNode
            |> List.choose (fun child ->
                Xml.name child
                |> Option.bind (fun name ->
                    match Ns.stripNs name with
                    | "t"
                    | "delText" -> Xml.text child |> Option.map (Text >> makeRun >> RunToken)
                    | "tab"
                    // w:ptab is an absolute-position tab; a plain tab is the closest the model has.
                    | "ptab" -> Some (makeRun (Text "\t") |> RunToken)
                    // These used to be dropped: "e-mail" written with a non-breaking
                    // hyphen came out as "email".
                    | "noBreakHyphen" -> Some (makeRun (Text "\u2011") |> RunToken)
                    | "softHyphen" -> Some (makeRun (Text "\u00AD") |> RunToken)
                    | "cr" -> Some (makeRun (Text "\n") |> RunToken)
                    | "sym" ->
                        // <w:sym w:font="Wingdings" w:char="F0FC"/>: one character in a named font.
                        attr "char" child
                        |> Option.bind (fun hex ->
                            match System.Int32.TryParse(hex, System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture) with
                            | true, code when code > 0 && code < 0xFFFF && not (code >= 0xD800 && code <= 0xDFFF) -> Some (string (char code))
                            | _ -> None)
                        |> Option.map (fun text ->
                            let run = makeRun (Text text)
                            let run =
                                match attr "font" child with
                                | Some font when font <> "" -> { run with Properties = { run.Properties with FontFamily = Some font; ThemeFont = None } }
                                | _ -> run
                            RunToken run)
                    | "br" ->
                        match attr "type" child with
                        // A column break forces the following content to the next
                        // column; in a multi-column section the layout's page-break
                        // (chunk) boundary is remapped to the next column, and in a
                        // single-column section it behaves as a page break (as Word).
                        | Some "page" | Some "column" -> Some (makeRun PageBreak |> RunToken)
                        | _ -> Some (makeRun (Text "\n") |> RunToken)
                    | "lastRenderedPageBreak" -> Some (makeRun RenderedPageBreak |> RunToken)
                    | "drawing"
                    | "pict" -> imageData rNode |> Option.map (Image >> makeRun >> RunToken)
                    | "commentReference" -> attr "id" child |> Option.map (CommentReference >> makeRun >> RunToken)
                    | "footnoteReference" -> attr "id" child |> Option.map (fun id -> makeRun (NoteRef { Id = id; NoteType = "footnote" }) |> RunToken)
                    | "endnoteReference" -> attr "id" child |> Option.map (fun id -> makeRun (NoteRef { Id = id; NoteType = "endnote" }) |> RunToken)
                    | "fldChar" ->
                        match attr "fldCharType" child with
                        | Some "begin" -> Some (FieldBeginToken props)
                        | Some "separate" -> Some FieldSeparateToken
                        | Some "end" -> Some FieldEndToken
                        | _ -> None
                    | "instrText" -> Xml.text child |> Option.map FieldInstructionToken
                    | _ -> None))

        if not childTokens.IsEmpty then childTokens
        else
            match imageData rNode with
            | Some img -> [makeRun (Image img) |> RunToken]
            | None ->
                match commentReference rNode with
                | Some commentId -> [makeRun (CommentReference commentId) |> RunToken]
                | None ->
                    match noteReference rNode with
                    | Some noteRef -> [makeRun (NoteRef noteRef) |> RunToken]
                    | None ->
                        match fieldData rNode with
                        | Some field -> [makeRun (FieldRef field) |> RunToken]
                        | None -> textContent rNode |> Option.map (Text >> makeRun >> RunToken >> List.singleton) |> Option.defaultValue []

    let runsFromRun (rNode: XmlNode) : Run list =
        tokensFromRun rNode |> collapseComplexFieldTokens |> attachRunMarkers

    let private tokenHasField = function
        | FieldBeginToken _
        | FieldSeparateToken
        | FieldEndToken
        | FieldInstructionToken _ -> true
        | RunToken _ -> false

    let private hyperlinkRunProps (props: RunProperties) =
        match props.Style with
        | Some _ -> props
        | None -> { props with Style = Some "Hyperlink" }

    let private hyperlinkToken = function
        | RunToken run -> RunToken { run with Properties = hyperlinkRunProps run.Properties }
        | FieldBeginToken props -> FieldBeginToken (hyperlinkRunProps props)
        | token -> token

    let private tokensFromHyperlink (hlNode: XmlNode) : RunToken list =
        let rId = Xml.attr "r:id" hlNode
        let tokens = wElems "r" hlNode |> List.collect tokensFromRun

        if tokens |> List.exists tokenHasField then
            tokens |> List.map hyperlinkToken
        else
            let runs = tokens |> collapseComplexFieldTokens |> attachRunMarkers

            // Extract text from runs
            let text =
                runs
                |> List.map (fun r ->
                    match r.Content with
                    | Text t -> t
                    | _ -> "")
                |> String.concat ""

            let props =
                runs
                |> List.tryHead
                |> Option.map (fun r -> hyperlinkRunProps r.Properties)
                |> Option.defaultValue {
                    Style = Some "Hyperlink"
                    Bold = None; Italic = None; Underline = None; UnderlineStyle = None; UnderlineColor = None
                    AllCaps = None; SmallCaps = None
                    FontSize = None; CharacterSpacing = None; FontFamily = None; ThemeFont = None; Color = None
                    ThemeColor = None; ThemeTint = None; ThemeShade = None; Highlight = None; Strike = None
                    DoubleStrike = None; Hidden = None; Position = None; EmphasisMark = None; VertAlign = None
                }

            [RunToken {
                Content = Link {
                    Text = text
                    Target = None  // Resolved from relationships later
                    RelationshipId = rId
                    Properties = props
                }
                Properties = props
                BookmarkStart = None
                BookmarkEnd = None
                RenderedPageBreakBefore = false
                RevisionId = None
                RevisionType = None
                RevisionAuthor = None
                RevisionDate = None
            }]

    // Hyperlink extraction
    let hyperlink (hlNode: XmlNode) : Run list =
        tokensFromHyperlink hlNode |> collapseComplexFieldTokens |> attachRunMarkers

    let private hasRenderedPageBreakBeforeText (pNode: XmlNode) =
        let rec loop seenText nodes =
            match nodes with
            | [] -> false
            | child :: rest ->
                match Xml.name child |> Option.map Ns.stripNs with
                | Some "r" ->
                    let runChildren = Xml.children child
                    let hasBreak = runChildren |> List.exists (fun node -> Xml.name node |> Option.exists (fun name -> Ns.stripNs name = "lastRenderedPageBreak"))
                    let hasText =
                        runChildren
                        |> List.exists (fun node ->
                            Xml.name node
                            |> Option.exists (fun name ->
                                let local = Ns.stripNs name
                                local = "t" || local = "tab" || local = "drawing"))
                    if hasBreak && not seenText then true
                    else loop (seenText || hasText) rest
                | Some "lastRenderedPageBreak" when not seenText -> true
                | _ -> loop seenText rest
        loop false (Xml.children pNode)

    // Paragraph extraction with functional composition
    let private paragraphWithRenderedBreakCounter (renderedBreakCounter: int ref option) (pNode: XmlNode) : Paragraph =
        let hasRenderedPageBreak = Ns.descendantsNs "lastRenderedPageBreak" pNode |> List.isEmpty |> not
        let renderedPageBreakBefore =
            match renderedBreakCounter with
            | Some counter ->
                let beforeText = hasRenderedPageBreakBeforeText pNode
                let breakCount = Ns.descendantsNs "lastRenderedPageBreak" pNode |> List.length
                let ordinal = if beforeText then Some (!counter + 1) else None
                counter := !counter + breakCount
                ordinal
            | None -> None
        let rawProps = paragraphProps pNode
        let props =
            let withRenderedBreak = { rawProps with RenderedPageBreakBefore = renderedPageBreakBefore }
            if hasRenderedPageBreak && rawProps.Style = Some "Heading1" then
                { withRenderedBreak with PageBreakBefore = Some true }
            else
                withRenderedBreak

        let rec tokensFromNode node =
            Xml.children node
            |> List.collect (fun child ->
                match Xml.name child with
                | Some name when Ns.stripNs name = "r" ->
                    tokensFromRun child
                | Some name when Ns.stripNs name = "hyperlink" ->
                    tokensFromHyperlink child
                | Some name when Ns.stripNs name = "fldSimple" ->
                    simpleField child |> Option.map RunToken |> Option.toList
                | Some name when Ns.stripNs name = "oMath" || Ns.stripNs name = "oMathPara" ->
                    [RunToken { Content = EquationXml (xmlToString child); Properties = runProps child; BookmarkStart = None; BookmarkEnd = None; RenderedPageBreakBefore = false; RevisionId = None; RevisionType = None; RevisionAuthor = None; RevisionDate = None }]
                | Some name when Ns.stripNs name = "bookmarkStart" ->
                    attr "id" child
                    |> Option.map (fun id -> RunToken { Content = BookmarkStart { Id = id; Name = attr "name" child }; Properties = runProps child; BookmarkStart = None; BookmarkEnd = None; RenderedPageBreakBefore = false; RevisionId = None; RevisionType = None; RevisionAuthor = None; RevisionDate = None })
                    |> Option.toList
                | Some name when Ns.stripNs name = "bookmarkEnd" ->
                    attr "id" child
                    |> Option.map (fun id -> RunToken { Content = BookmarkEnd id; Properties = runProps child; BookmarkStart = None; BookmarkEnd = None; RenderedPageBreakBefore = false; RevisionId = None; RevisionType = None; RevisionAuthor = None; RevisionDate = None })
                    |> Option.toList
                // Transparent inline wrappers: their runs are ordinary paragraph content.
                | Some name when (match Ns.stripNs name with "smartTag" | "customXml" | "dir" | "bdo" -> true | _ -> false) ->
                    tokensFromNode child
                | Some name when Ns.stripNs name = "AlternateContent" ->
                    // Take the fallback (plain WordprocessingML/VML) when there is one, else the first choice.
                    match Ns.childNs "Fallback" child |> Option.orElseWith (fun () -> Ns.childNs "Choice" child) with
                    | Some branch -> tokensFromNode branch
                    | None -> []
                // A tracked move is an insertion at the destination and a deletion at the
                // source; without this the moved text was missing from both places.
                | Some name when (match Ns.stripNs name with "ins" | "del" | "moveTo" | "moveFrom" -> true | _ -> false) ->
                    let revisionType = if Ns.stripNs name = "del" || Ns.stripNs name = "moveFrom" then "delete" else "insert"
                    let revisionId = Xml.attr "id" child |> Option.orElse (Xml.attr "w:id" child)
                    let author = Xml.attr "author" child |> Option.orElse (Xml.attr "w:author" child)
                    let date = Xml.attr "date" child |> Option.orElse (Xml.attr "w:date" child)
                    tokensFromNode child
                    |> List.map (function
                        | RunToken run -> RunToken { run with RevisionId = revisionId; RevisionType = Some revisionType; RevisionAuthor = author; RevisionDate = date }
                        | token -> token)
                | Some name when Ns.stripNs name = "sdt" ->
                    // A w14:checkbox content control renders as a real checkbox
                    // glyph from its checked state (the stored glyph is often a
                    // symbol-font char that won't render); other controls just
                    // surface their inner content.
                    let checkbox = Ns.childNs "sdtPr" child |> Option.bind (fun pr -> Ns.descendantsNs "checkbox" pr |> List.tryHead)
                    match checkbox with
                    | Some cb ->
                        let isChecked =
                            Ns.descendantsNs "checked" cb |> List.tryHead
                            |> Option.bind (attrByLocalName "val")   // w14:val
                            |> Option.map (fun v -> v = "1" || v = "true")
                            |> Option.defaultValue false
                        let glyph = if isChecked then "☒" else "☐"   // ☒ / ☐
                        [RunToken { Content = Text glyph; Properties = runProps child; BookmarkStart = None; BookmarkEnd = None; RenderedPageBreakBefore = false; RevisionId = None; RevisionType = None; RevisionAuthor = None; RevisionDate = None }]
                    | None ->
                        Ns.childNs "sdtContent" child
                        |> Option.map tokensFromNode
                        |> Option.defaultValue []
                | _ -> [])

        // Extract runs and hyperlinks
        let runs = tokensFromNode pNode |> collapseComplexFieldTokens |> attachRunMarkers |> normalizeTocRuns props

        { Runs = runs; Properties = props }

    let paragraph (pNode: XmlNode) : Paragraph =
        paragraphWithRenderedBreakCounter None pNode

    // Border style extraction
    let borderStyle (borderNode: XmlNode) : BorderStyle option =
        Some {
            BorderType = attr "val" borderNode
            Size = attr "sz" borderNode
            Color = attr "color" borderNode
            Space = attr "space" borderNode
        }

    let bordersFromElement (borders: XmlNode) : CellBorders = {
        Top = wElem "top" borders |> Option.bind borderStyle
        Bottom = wElem "bottom" borders |> Option.bind borderStyle
        Left = wElem "left" borders |> Option.bind borderStyle
        Right = wElem "right" borders |> Option.bind borderStyle
        InsideH = wElem "insideH" borders |> Option.bind borderStyle
        InsideV = wElem "insideV" borders |> Option.bind borderStyle
    }

    let bordersChild childName node =
        wElem childName node
        |> Option.map bordersFromElement

    // Cell borders extraction
    let cellBorders (tcPr: XmlNode) : CellBorders option =
        bordersChild "tcBorders" tcPr

    // Shading extraction
    let shading (node: XmlNode) : Shading option =
        wElem "shd" node
        |> Option.map (fun shd ->
            {
                Fill = attr "fill" shd
                Color = attr "color" shd
                Pattern = attr "val" shd
            })

    // Table cell properties
    let tableCellPropertiesFromTcPr (tcPr: XmlNode) : TableCellProperties =
        let margin name =
            wElem "tcMar" tcPr
            |> Option.bind (wElem name)
            |> Option.bind (attr "w")
        {
            Borders = cellBorders tcPr
            Shading = shading tcPr
            Width = wElem "tcW" tcPr |> Option.bind (attr "w")
            WidthType = wElem "tcW" tcPr |> Option.bind (attr "type")
            VerticalAlignment = wElem "vAlign" tcPr |> Option.bind (attr "val")
            PaddingTop = margin "top"
            PaddingRight = margin "right"
            PaddingBottom = margin "bottom"
            PaddingLeft = margin "left"
            ColSpan = wElem "gridSpan" tcPr |> Option.bind (attr "val") |> Option.bind (fun value -> match System.Int32.TryParse(value) with | true, parsed -> Some parsed | false, _ -> None)
            VerticalMerge = wElem "vMerge" tcPr |> Option.bind (attr "val") |> Option.orElse (wElem "vMerge" tcPr |> Option.map (fun _ -> "continue"))
        }

    let tableCellProps (tcNode: XmlNode) : TableCellProperties option =
        wElem "tcPr" tcNode
        |> Option.map tableCellPropertiesFromTcPr

    let private wChildElems name node =
        Xml.children node
        |> List.filter (fun child -> Xml.name child |> Option.exists (fun childName -> Ns.stripNs childName = name))

    /// Children of a table / row / cell with content-control and customXml
    /// wrappers opened up: <w:sdt><w:sdtContent><w:tr>.. is a row of the table
    /// (a repeating-section or bound row), not something to skip.
    let rec private unwrappedChildren (node: XmlNode) : XmlNode list =
        Xml.children node
        |> List.collect (fun child ->
            match Xml.name child |> Option.map Ns.stripNs with
            | Some "sdt" -> Ns.childNs "sdtContent" child |> Option.map unwrappedChildren |> Option.defaultValue []
            | Some "customXml" -> unwrappedChildren child
            | _ -> [child])

    let private unwrappedChildElems name (node: XmlNode) =
        unwrappedChildren node
        |> List.filter (fun child -> Xml.name child |> Option.exists (fun childName -> Ns.stripNs childName = name))

    // Table extraction with nested document builder
    let rec tableCell (tcNode: XmlNode) : TableCell =
        // Walk p/tbl children in document order so we can record, for each nested
        // table, how many paragraphs precede it (its interleave anchor).
        let mutable paragraphCount = 0
        let mutable paragraphs = []
        let mutable nestedTables = []
        let mutable nestedAnchors = []
        for child in unwrappedChildren tcNode do
            match Xml.name child |> Option.map Ns.stripNs with
            | Some "p" ->
                paragraphs <- paragraph child :: paragraphs
                paragraphCount <- paragraphCount + 1
            | Some "tbl" ->
                nestedTables <- table child :: nestedTables
                nestedAnchors <- paragraphCount :: nestedAnchors
            | _ -> ()
        {
            Paragraphs = List.rev paragraphs
            NestedTables = List.rev nestedTables
            NestedTableAnchors = List.rev nestedAnchors
            Properties = tableCellProps tcNode
        }

    and tableRow (trNode: XmlNode) : TableRow =
        let cells =
            unwrappedChildElems "tc" trNode
            |> List.map tableCell
        let height = wElem "trPr" trNode
                     |> Option.bind (wElem "trHeight")
                     |> Option.bind (attr "val")
        let heightRule = wElem "trPr" trNode
                         |> Option.bind (wElem "trHeight")
                         |> Option.bind (attr "hRule")
        let isHeader =
            wElem "trPr" trNode
            |> Option.bind (onOff "tblHeader")
            |> Option.defaultValue false
        let canSplit =
            wElem "trPr" trNode
            |> Option.bind (onOff "cantSplit")
            |> Option.map not
        let tblPrEx = wElem "tblPrEx" trNode
        let cellSpacing =
            wElem "trPr" trNode
            |> Option.bind (wElem "tblCellSpacing")
            |> Option.bind (attr "w")
            |> Option.orElseWith (fun () ->
                tblPrEx
                |> Option.bind (wElem "tblCellSpacing")
                |> Option.bind (attr "w"))
        let rowMargin name =
            tblPrEx
            |> Option.bind (wElem "tblCellMar")
            |> Option.bind (wElem name)
            |> Option.bind (attr "w")
        let bordersOverride =
            tblPrEx |> Option.bind (bordersChild "tblBorders")
        let trPr = wElem "trPr" trNode
        let intOpt name =
            trPr
            |> Option.bind (wElem name)
            |> Option.bind (attr "val")
            |> Option.bind (fun s -> match System.Int32.TryParse(s) with | true, v -> Some v | _ -> None)
        let widthOpt name =
            trPr
            |> Option.bind (wElem name)
            |> Option.bind (attr "w")
        {
            Cells = cells
            Height = height
            HeightRule = heightRule
            IsHeader = isHeader
            CanSplit = canSplit
            CellSpacing = cellSpacing
            PaddingTop = rowMargin "top"
            PaddingRight = rowMargin "right"
            PaddingBottom = rowMargin "bottom"
            PaddingLeft = rowMargin "left"
            GridBefore = intOpt "gridBefore"
            GridAfter = intOpt "gridAfter"
            WidthBefore = widthOpt "wBefore"
            WidthAfter = widthOpt "wAfter"
            BordersOverride = bordersOverride
        }

    and table (tblNode: XmlNode) : Table =
        let rows =
            unwrappedChildElems "tr" tblNode
            |> List.map tableRow

        let borders =
            wElem "tblPr" tblNode
            |> Option.bind (bordersChild "tblBorders")

        let tblPr = wElem "tblPr" tblNode
        let tableWidth = tblPr |> Option.bind (wElem "tblW")
        let cellSpacing = tblPr |> Option.bind (wElem "tblCellSpacing")
        let layoutMode = tblPr |> Option.bind (wElem "tblLayout") |> Option.bind (attr "type")
        let floating = tblPr |> Option.bind (wElem "tblpPr") |> Option.isSome
        let tableProperties =
            tblPr
            |> Option.map (fun props ->
                let margin name =
                    wElem "tblCellMar" props
                    |> Option.bind (wElem name)
                    |> Option.bind (attr "w")
                { Borders = borders
                  Shading = shading props
                  PaddingTop = margin "top"
                  PaddingRight = margin "right"
                  PaddingBottom = margin "bottom"
                  PaddingLeft = margin "left" })
            |> Option.bind (fun props ->
                if props.Borders.IsNone
                   && props.Shading.IsNone
                   && props.PaddingTop.IsNone
                   && props.PaddingRight.IsNone
                   && props.PaddingBottom.IsNone
                   && props.PaddingLeft.IsNone then None
                else Some props)
        let tableLook =
            tblPr
            |> Option.bind (wElem "tblLook")
            |> Option.map (fun look ->
                let lookValue =
                    attr "val" look
                    |> Option.bind (fun value ->
                        match System.Int32.TryParse(value, System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture) with
                        | true, parsed -> Some parsed
                        | false, _ -> None)
                let bit mask fallback =
                    lookValue
                    |> Option.map (fun value -> (value &&& mask) <> 0)
                    |> Option.defaultValue fallback
                let flag name fallback =
                    attr name look
                    |> Option.map (fun value -> value = "1" || value = "true")
                    |> Option.defaultValue fallback
                { FirstRow = flag "firstRow" (bit 0x0020 false)
                  LastRow = flag "lastRow" (bit 0x0040 false)
                  FirstColumn = flag "firstColumn" (bit 0x0080 false)
                  LastColumn = flag "lastColumn" (bit 0x0100 false)
                  HorizontalBanding = not (flag "noHBand" (bit 0x0200 false))
                  VerticalBanding = not (flag "noVBand" (bit 0x0400 false)) })
        let gridColumns =
            wElem "tblGrid" tblNode
            |> Option.map (fun grid -> wChildElems "gridCol" grid |> List.choose (attr "w"))
            |> Option.defaultValue []

        {
            Rows = rows
            StyleId = tblPr |> Option.bind (wElem "tblStyle") |> Option.bind (attr "val")
            Look = tableLook
            Alignment = tblPr |> Option.bind (wElem "jc") |> Option.bind (attr "val")
            Properties = tableProperties
            Borders = borders
            Width = tableWidth |> Option.bind (attr "w")
            WidthType = tableWidth |> Option.bind (attr "type")
            CellSpacing = cellSpacing |> Option.bind (attr "w")
            LayoutMode = layoutMode
            GridColumns = gridColumns
            Floating = floating
        }

    // Main document content extraction using document builder
    let bodyContentWithFinalSection (finalSectPr: XmlNode option) (bodyNode: XmlNode) : DocumentContent list =
        let renderedBreakCounter = ref 0
        let contentControlProps (sdtNode: XmlNode) =
            let props = Ns.childNs "sdtPr" sdtNode
            let id = props |> Option.bind (Ns.childNs "id") |> Option.bind (attr "val")
            let tag = props |> Option.bind (Ns.childNs "tag") |> Option.bind (attr "val")
            let alias = props |> Option.bind (Ns.childNs "alias") |> Option.bind (attr "val")
            id, tag, alias
        let applyContentControl (id, tag, alias) = function
            | Para para ->
                Para { para with Properties = { para.Properties with ContentControlId = id; ContentControlTag = tag; ContentControlAlias = alias } }
            | item -> item
        let rec contentFromNode node =
            document {
                for child in Xml.children node do
                    match Xml.name child with
                    | Some name when Ns.stripNs name = "p" ->
                        yield Para (paragraphWithRenderedBreakCounter (Some renderedBreakCounter) child)
                    | Some name when Ns.stripNs name = "tbl" ->
                        yield Tbl (table child)
                    | Some name when Ns.stripNs name = "sdt" ->
                        match Ns.childNs "sdtContent" child with
                        | Some content ->
                            let controlProps = contentControlProps child
                            yield! (contentFromNode content |> List.map (applyContentControl controlProps))
                        | None -> ()
                    | Some name when Ns.stripNs name = "customXml" ->
                        yield! contentFromNode child
                    | _ -> ()
            }

        // Iterative: the recursive form was not a tail call and overflowed the JS
        // stack on documents with a few thousand body blocks (the file would not open).
        let normalizeSectionBreaks (items: DocumentContent list) =
            let source = List.toArray items
            source
            |> Array.mapi (fun index item ->
                match item with
                | Para para when index + 1 < source.Length && para.Runs.IsEmpty && para.Properties.SectionBreak = Some "nextPage" ->
                    match source.[index + 1] with
                    | Para nextPara when nextPara.Properties.SectionBreak = Some "continuous" ->
                        Para { para with Properties = { para.Properties with SectionBreak = None } }
                    | _ -> item
                | _ -> item)
            |> Array.toList

        let content = contentFromNode bodyNode |> normalizeSectionBreaks
        match finalSectPr with
        | None -> content
        | Some section ->
            let finalHeaderRefs = headerFooterReferences "headerReference" section
            let finalFooterRefs = headerFooterReferences "footerReference" section
            let finalPageNumberStart =
                getPropVal "pgNumType" "start" section
                |> Option.bind (fun value -> match System.Int32.TryParse(value) with | true, parsed -> Some parsed | false, _ -> None)
            if finalHeaderRefs.IsEmpty && finalFooterRefs.IsEmpty && finalPageNumberStart.IsNone then
                content
            else
                let hasInlineSection =
                    content
                    |> List.exists (function
                        | Para para -> para.Properties.SectionBreak.IsSome || not para.Properties.HeaderReferences.IsEmpty || not para.Properties.FooterReferences.IsEmpty
                        | _ -> false)

                let applyFinalSectionProps (para: Paragraph) : ParagraphProperties =
                    { para.Properties with
                        HeaderReferences = if finalHeaderRefs.IsEmpty then para.Properties.HeaderReferences else finalHeaderRefs
                        FooterReferences = if finalFooterRefs.IsEmpty then para.Properties.FooterReferences else finalFooterRefs
                        SectionPageNumberStart = finalPageNumberStart |> Option.orElse para.Properties.SectionPageNumberStart }

                let rec apply applied pending items =
                    match items with
                    | [] -> []
                    | Para para :: rest ->
                        if pending || (not applied && finalPageNumberStart.IsSome && not hasInlineSection) then
                            Para { para with Properties = applyFinalSectionProps para } :: rest
                        else
                            let startsNextSection = not para.Properties.HeaderReferences.IsEmpty || not para.Properties.FooterReferences.IsEmpty
                            Para para :: apply applied startsNextSection rest
                    | item :: rest -> item :: apply applied pending rest
                apply false false content

    let bodyContent (bodyNode: XmlNode) : DocumentContent list =
        bodyContentWithFinalSection None bodyNode

// Style parsing
module Styles =
    let defaultStyleId = "__docDefaults"

    let parseRunPropsFromStyle (rPr: XmlNode) : RunProperties =
        {
            Style = None
            Bold = Docx.onOff "b" rPr
            Italic = Docx.onOff "i" rPr
            // As in Docx.runProps: <w:u w:val="none"/> in a style means not underlined.
            Underline =
                (match Ns.childNs "u" rPr |> Option.bind (Docx.attr "val") with
                 | Some value when value.ToLowerInvariant() = "none" -> Some false
                 | _ -> Docx.onOff "u" rPr)
            UnderlineStyle = Ns.childNs "u" rPr |> Option.bind (Docx.attr "val")
            UnderlineColor = Ns.childNs "u" rPr |> Option.bind (Docx.attr "color")
            AllCaps = Docx.onOff "caps" rPr
            SmallCaps = Docx.onOff "smallCaps" rPr
            FontSize = Ns.childNs "sz" rPr |> Option.bind (Docx.attr "val")
            CharacterSpacing = Ns.childNs "spacing" rPr |> Option.bind (Docx.attr "val")
            FontFamily = Ns.childNs "rFonts" rPr |> Option.bind (Docx.attr "ascii")
            ThemeFont = Ns.childNs "rFonts" rPr |> Option.bind (Docx.attr "asciiTheme")
            Color = Ns.childNs "color" rPr |> Option.bind (Docx.attr "val")
            ThemeColor = Ns.childNs "color" rPr |> Option.bind (Docx.attr "themeColor")
            ThemeTint = Ns.childNs "color" rPr |> Option.bind (Docx.attr "themeTint")
            ThemeShade = Ns.childNs "color" rPr |> Option.bind (Docx.attr "themeShade")
            Highlight = Ns.childNs "highlight" rPr |> Option.bind (Docx.attr "val")
            Strike = Docx.onOff "strike" rPr
            DoubleStrike = Docx.onOff "dstrike" rPr
            Hidden = Docx.onOff "vanish" rPr
            Position = Ns.childNs "position" rPr |> Option.bind (Docx.attr "val")
            EmphasisMark = Ns.childNs "em" rPr |> Option.bind (Docx.attr "val")
            VertAlign = Ns.childNs "vertAlign" rPr |> Option.bind (Docx.attr "val")
        }

    let parseParagraphPropsFromStyle (pPr: XmlNode) : ParagraphProperties =
        let rightTab = Docx.rightTabStop pPr
        let listInfo =
            Ns.childNs "numPr" pPr
            |> Option.bind (fun numPr ->
                Ns.childNs "numId" numPr
                |> Option.bind (Docx.attr "val")
                |> Option.map (fun numId -> {
                    ListId = numId
                    Level = Ns.childNs "ilvl" numPr |> Option.bind (Docx.attr "val") |> Option.map int |> Option.defaultValue 0
                    NumberFormat = None
                    LevelText = None
                    Start = None
                }))
        {
            Style = Ns.childNs "pStyle" pPr |> Option.bind (Docx.attr "val")
            ContentControlId = None
            ContentControlTag = None
            ContentControlAlias = None
            Alignment = Ns.childNs "jc" pPr |> Option.bind (Docx.attr "val")
            IndentLeft = Ns.childNs "ind" pPr |> Option.bind (Docx.attr "left")
            IndentRight = Ns.childNs "ind" pPr |> Option.bind (Docx.attr "right")
            FirstLineIndent = Ns.childNs "ind" pPr |> Option.bind (Docx.attr "firstLine")
            HangingIndent = Ns.childNs "ind" pPr |> Option.bind (Docx.attr "hanging")
            RightTabPosition = rightTab |> Option.bind fst
            RightTabLeader = rightTab |> Option.bind snd
            TabStops = Docx.allTabStops pPr
            Rtl = Docx.onOff "bidi" pPr
            SpacingBefore = Ns.childNs "spacing" pPr |> Option.bind (Docx.attr "before")
            SpacingBeforeAuto = Ns.childNs "spacing" pPr |> Option.bind (Docx.attr "beforeAutospacing") |> Option.map (fun value -> value = "1" || value.ToLowerInvariant() = "true")
            SpacingAfter = Ns.childNs "spacing" pPr |> Option.bind (Docx.attr "after")
            SpacingAfterAuto = Ns.childNs "spacing" pPr |> Option.bind (Docx.attr "afterAutospacing") |> Option.map (fun value -> value = "1" || value.ToLowerInvariant() = "true")
            LineSpacing = Ns.childNs "spacing" pPr |> Option.bind (Docx.attr "line")
            LineRule = Ns.childNs "spacing" pPr |> Option.bind (Docx.attr "lineRule")
            OutlineLevel = Ns.childNs "outlineLvl" pPr |> Option.bind (Docx.attr "val")
            KeepNext = Docx.onOff "keepNext" pPr
            KeepLines = Docx.onOff "keepLines" pPr
            PageBreakBefore = Docx.onOff "pageBreakBefore" pPr
            WidowControl = Docx.onOff "widowControl" pPr
            ContextualSpacing = Docx.onOff "contextualSpacing" pPr
            BackgroundColor = Ns.childNs "shd" pPr |> Option.bind (Docx.attr "fill")
            ParagraphRunProperties = Ns.childNs "rPr" pPr |> Option.map (fun _ -> Docx.runProps pPr)
            ParagraphBorders = Docx.paragraphBorders pPr
            ListInfo = listInfo
            SectionBreak = None
            SectionGeometry = None
            SectionColumns = None
            SectionPageNumberStart = None
            RenderedPageBreakBefore = None
            HeaderReferences = []
            FooterReferences = []
        }

    let private tableStylePropertiesFrom marginsChild bordersChild (node: XmlNode) : TableStyleProperties option =
        let margin name =
            Ns.childNs marginsChild node
            |> Option.bind (Ns.childNs name)
            |> Option.bind (Docx.attr "w")

        let properties = {
            Borders = Docx.bordersChild bordersChild node
            Shading = Docx.shading node
            PaddingTop = margin "top"
            PaddingRight = margin "right"
            PaddingBottom = margin "bottom"
            PaddingLeft = margin "left"
        }

        if properties.Borders.IsNone
           && properties.Shading.IsNone
           && properties.PaddingTop.IsNone
           && properties.PaddingRight.IsNone
           && properties.PaddingBottom.IsNone
           && properties.PaddingLeft.IsNone then None
        else Some properties

    let private tablePropertiesFromTblPr (node: XmlNode) =
        tableStylePropertiesFrom "tblCellMar" "tblBorders" node

    let private tablePropertiesFromTcPr (node: XmlNode) =
        tableStylePropertiesFrom "tcMar" "tcBorders" node

    let private tableRowPropertiesFromTrPr (node: XmlNode) : TableRowStyleProperties option =
        let properties = {
            IsHeader = Docx.onOff "tblHeader" node
            CanSplit = Docx.onOff "cantSplit" node |> Option.map not
        }

        if properties.IsHeader.IsNone && properties.CanSplit.IsNone then None
        else Some properties

    let parseStyle (styleNode: XmlNode) : StyleDefinition option =
        Docx.attr "styleId" styleNode
        |> Option.map (fun styleId ->
            let name = Ns.childNs "name" styleNode |> Option.bind (Docx.attr "val")
            let basedOn = Ns.childNs "basedOn" styleNode |> Option.bind (Docx.attr "val")
            let next = Ns.childNs "next" styleNode |> Option.bind (Docx.attr "val")
            let linked = Ns.childNs "link" styleNode |> Option.bind (Docx.attr "val")
            let isQuickStyle = Docx.onOff "qFormat" styleNode |> Option.defaultValue false
            let aliases = Ns.childNs "aliases" styleNode |> Option.bind (Docx.attr "val")
            let uiPriority = Ns.childNs "uiPriority" styleNode |> Option.bind (Docx.attr "val")
            let semiHidden = Docx.onOff "semiHidden" styleNode |> Option.defaultValue false
            let unhideWhenUsed = Docx.onOff "unhideWhenUsed" styleNode |> Option.defaultValue false
            let rPr = Ns.childNs "rPr" styleNode |> Option.map parseRunPropsFromStyle
            let pPr = Ns.childNs "pPr" styleNode |> Option.map parseParagraphPropsFromStyle
            let tblPr = Ns.childNs "tblPr" styleNode |> Option.bind tablePropertiesFromTblPr
            let conditionalStyleChild condition =
                Xml.children styleNode
                |> List.tryPick (fun child ->
                    if Xml.name child |> Option.exists (fun name -> Ns.stripNs name = "tblStylePr")
                       && Docx.attr "type" child = Some condition then
                        Some child
                    else None)
            let conditionalChildProperties childName propertyParser condition =
                conditionalStyleChild condition
                |> Option.bind (fun child -> Ns.childNs childName child |> Option.bind propertyParser)
            let conditionalTableProperties condition = conditionalChildProperties "tblPr" tablePropertiesFromTblPr condition
            let conditionalRowProperties condition = conditionalChildProperties "trPr" tableRowPropertiesFromTrPr condition
            let conditionalCellProperties condition = conditionalChildProperties "tcPr" (Docx.tableCellPropertiesFromTcPr >> Some) condition
            let conditionalParagraphProperties condition =
                conditionalStyleChild condition
                |> Option.bind (fun child ->
                    let paragraphProperties =
                        Ns.childNs "pPr" child
                        |> Option.map parseParagraphPropsFromStyle
                    let runProperties =
                        Ns.childNs "rPr" child
                        |> Option.map parseRunPropsFromStyle
                    match paragraphProperties, runProperties with
                    | None, None -> None
                    | Some props, None -> Some props
                    | Some props, Some runProps ->
                        Some { props with ParagraphRunProperties = Some runProps }
                    | None, Some runProps ->
                        Some { Docx.paragraphProps XmlDsl.Empty with ParagraphRunProperties = Some runProps })

            {
                Id = styleId
                Name = name
                StyleType = Docx.attr "type" styleNode |> Option.defaultValue "paragraph"
                IsDefault = Docx.attr "default" styleNode |> Option.map ((=) "1") |> Option.defaultValue false
                BasedOn = basedOn
                Next = next
                Linked = linked
                IsQuickStyle = isQuickStyle
                Aliases = aliases
                UiPriority = uiPriority
                SemiHidden = semiHidden
                UnhideWhenUsed = unhideWhenUsed
                RunProperties = rPr
                ParagraphProperties = pPr
                TableProperties = tblPr
                FirstRowParagraphProperties = conditionalParagraphProperties "firstRow"
                FirstRowTableRowProperties = conditionalRowProperties "firstRow"
                FirstRowTableProperties = conditionalTableProperties "firstRow"
                LastRowTableProperties = conditionalTableProperties "lastRow"
                FirstColumnTableProperties = conditionalTableProperties "firstCol"
                LastColumnTableProperties = conditionalTableProperties "lastCol"
                Band1HorizontalTableProperties = conditionalTableProperties "band1Horz"
                Band2HorizontalTableProperties = conditionalTableProperties "band2Horz"
                Band1VerticalTableProperties = conditionalTableProperties "band1Vert"
                Band2VerticalTableProperties = conditionalTableProperties "band2Vert"
                FirstRowTableCellProperties = conditionalCellProperties "firstRow"
                LastRowTableCellProperties = conditionalCellProperties "lastRow"
                FirstColumnTableCellProperties = conditionalCellProperties "firstCol"
                LastColumnTableCellProperties = conditionalCellProperties "lastCol"
                Band1HorizontalTableCellProperties = conditionalCellProperties "band1Horz"
                Band2HorizontalTableCellProperties = conditionalCellProperties "band2Horz"
                Band1VerticalTableCellProperties = conditionalCellProperties "band1Vert"
                Band2VerticalTableCellProperties = conditionalCellProperties "band2Vert"
            })

    let parseDocDefaults (doc: XmlNode) : StyleDefinition option =
        Ns.descendantsNs "docDefaults" doc
        |> List.tryHead
        |> Option.bind (fun defaults ->
            let rPr =
                Ns.childNs "rPrDefault" defaults
                |> Option.bind (Ns.childNs "rPr")
                |> Option.map parseRunPropsFromStyle
            let pPr =
                Ns.childNs "pPrDefault" defaults
                |> Option.bind (Ns.childNs "pPr")
                |> Option.map parseParagraphPropsFromStyle
            if rPr.IsNone && pPr.IsNone then None
            else
                Some {
                    Id = defaultStyleId
                    Name = Some "Document Defaults"
                    StyleType = "defaults"
                    IsDefault = false
                    BasedOn = None
                    Next = None
                    Linked = None
                    IsQuickStyle = false
                    Aliases = None
                    UiPriority = None
                    SemiHidden = false
                    UnhideWhenUsed = false
                    RunProperties = rPr
                    ParagraphProperties = pPr
                    TableProperties = None
                    FirstRowParagraphProperties = None
                    FirstRowTableRowProperties = None
                    FirstRowTableProperties = None
                    LastRowTableProperties = None
                    FirstColumnTableProperties = None
                    LastColumnTableProperties = None
                    Band1HorizontalTableProperties = None
                    Band2HorizontalTableProperties = None
                    Band1VerticalTableProperties = None
                    Band2VerticalTableProperties = None
                    FirstRowTableCellProperties = None
                    LastRowTableCellProperties = None
                    FirstColumnTableCellProperties = None
                    LastColumnTableCellProperties = None
                    Band1HorizontalTableCellProperties = None
                    Band2HorizontalTableCellProperties = None
                    Band1VerticalTableCellProperties = None
                    Band2VerticalTableCellProperties = None
                })

    let parseStyles (xml: string) : StyleDefinition list =
        let doc = XmlDsl.parse xml
        (parseDocDefaults doc |> Option.toList) @ (Ns.descendantsNs "style" doc |> List.choose parseStyle)

module Theme =
    let private colorValue node =
        Ns.childNs "srgbClr" node
        |> Option.bind (Docx.attr "val")
        |> Option.orElseWith (fun () -> Ns.childNs "sysClr" node |> Option.bind (Docx.attr "lastClr"))
        |> Option.map (fun value -> if value.StartsWith("#") then value.ToUpper() else "#" + value.ToUpper())

    let private latinTypeface node =
        Ns.childNs "latin" node |> Option.bind (Docx.attr "typeface") |> Option.filter ((<>) "")

    let parseTheme (xml: string) : Theme =
        let doc = XmlDsl.parse xml
        let majorFont = Ns.descendantsNs "majorFont" doc |> List.tryHead |> Option.bind latinTypeface
        let minorFont = Ns.descendantsNs "minorFont" doc |> List.tryHead |> Option.bind latinTypeface
        let colorNames = Set.ofList ["dk1"; "lt1"; "dk2"; "lt2"; "accent1"; "accent2"; "accent3"; "accent4"; "accent5"; "accent6"; "hlink"; "folHlink"]
        let colors =
            Ns.descendantsNs "clrScheme" doc
            |> List.tryHead
            |> Option.map (fun scheme ->
                Xml.children scheme
                |> List.choose (fun child ->
                    match Xml.name child with
                    | Some name ->
                        let local = Ns.stripNs name
                        if Set.contains local colorNames then colorValue child |> Option.map (fun color -> local, color) else None
                    | None -> None)
                |> Map.ofList)
            |> Option.defaultValue Map.empty
        { MajorFont = majorFont; MinorFont = minorFont; Colors = colors }

module Numbering =
    type private LevelDefinition = {
        ParagraphStyle: string option
        NumberFormat: string
        LevelText: string option
        Start: int option
        IndentLeft: string option
        HangingIndent: string option
        FirstLineIndent: string option
        RightTabPosition: string option
    }

    let parseNumbering (xml: string) : NumberingDefinition list =
        let doc = XmlDsl.parse xml

        let abstractFormats =
            Ns.descendantsNs "abstractNum" doc
            |> List.choose (fun abstractNum ->
                Docx.attr "abstractNumId" abstractNum
                |> Option.map (fun abstractId ->
                    let levels =
                        Ns.descendantsNs "lvl" abstractNum
                        |> List.choose (fun level ->
                            Docx.attr "ilvl" level
                            |> Option.bind (fun ilvl ->
                                let parseLevel (level: XmlNode) =
                                    let numberFormat =
                                        Ns.childNs "numFmt" level
                                        |> Option.bind (Docx.attr "val")
                                        |> Option.defaultValue "decimal"
                                    let levelText =
                                        Ns.childNs "lvlText" level
                                        |> Option.bind (Docx.attr "val")
                                    let start =
                                        Ns.childNs "start" level
                                        |> Option.bind (Docx.attr "val")
                                        |> Option.bind (fun value ->
                                            match System.Int32.TryParse value with
                                            | true, parsed -> Some parsed
                                            | false, _ -> None)
                                    let pPr = Ns.childNs "pPr" level
                                    let ind = pPr |> Option.bind (Ns.childNs "ind")
                                    let tabPosition =
                                        pPr
                                        |> Option.bind (Ns.childNs "tabs")
                                        |> Option.bind (fun tabs ->
                                            Xml.children tabs
                                            |> List.tryPick (fun tab ->
                                                if Xml.name tab |> Option.exists (fun name -> Ns.stripNs name = "tab") then Docx.attr "pos" tab else None))
                                    {
                                        ParagraphStyle = Ns.childNs "pStyle" level |> Option.bind (Docx.attr "val")
                                        NumberFormat = numberFormat
                                        LevelText = levelText
                                        Start = start
                                        IndentLeft = ind |> Option.bind (Docx.attr "left")
                                        HangingIndent = ind |> Option.bind (Docx.attr "hanging")
                                        FirstLineIndent = ind |> Option.bind (Docx.attr "firstLine")
                                        RightTabPosition = tabPosition
                                    }
                                Some (int ilvl, parseLevel level)))
                    (abstractId, levels)))
            |> Map.ofList

        let abstractFormatMap =
            abstractFormats
            |> Map.map (fun _ levels -> levels |> Map.ofList)

        Ns.descendantsNs "num" doc
        |> List.collect (fun num ->
            match Docx.attr "numId" num, Ns.childNs "abstractNumId" num |> Option.bind (Docx.attr "val") with
            | Some numId, Some abstractId ->
                let startOverrides =
                    Ns.descendantsNs "lvlOverride" num
                    |> List.choose (fun overrideNode ->
                        Docx.attr "ilvl" overrideNode
                        |> Option.bind (fun ilvl ->
                            Ns.childNs "startOverride" overrideNode
                            |> Option.bind (Docx.attr "val")
                            |> Option.bind (fun value ->
                                match System.Int32.TryParse value with
                                | true, parsed -> Some (int ilvl, parsed)
                                | false, _ -> None)))
                    |> Map.ofList
                let levelOverrides =
                    Ns.descendantsNs "lvlOverride" num
                    |> List.choose (fun overrideNode ->
                        Docx.attr "ilvl" overrideNode
                        |> Option.bind (fun ilvl ->
                            Ns.childNs "lvl" overrideNode
                            |> Option.map (fun level ->
                                let numberFormat =
                                    Ns.childNs "numFmt" level
                                    |> Option.bind (Docx.attr "val")
                                    |> Option.defaultValue "decimal"
                                let levelText = Ns.childNs "lvlText" level |> Option.bind (Docx.attr "val")
                                let start = Ns.childNs "start" level |> Option.bind (Docx.attr "val") |> Option.bind (fun value -> match System.Int32.TryParse value with | true, parsed -> Some parsed | false, _ -> None)
                                let pPr = Ns.childNs "pPr" level
                                let ind = pPr |> Option.bind (Ns.childNs "ind")
                                let tabPosition =
                                    pPr
                                    |> Option.bind (Ns.childNs "tabs")
                                    |> Option.bind (fun tabs ->
                                        Xml.children tabs
                                        |> List.tryPick (fun tab ->
                                            if Xml.name tab |> Option.exists (fun name -> Ns.stripNs name = "tab") then Docx.attr "pos" tab else None))
                                int ilvl, {
                                    ParagraphStyle = Ns.childNs "pStyle" level |> Option.bind (Docx.attr "val")
                                    NumberFormat = numberFormat
                                    LevelText = levelText
                                    Start = start
                                    IndentLeft = ind |> Option.bind (Docx.attr "left")
                                    HangingIndent = ind |> Option.bind (Docx.attr "hanging")
                                    FirstLineIndent = ind |> Option.bind (Docx.attr "firstLine")
                                    RightTabPosition = tabPosition
                                })))
                    |> Map.ofList

                abstractFormatMap
                |> Map.tryFind abstractId
                |> Option.defaultValue Map.empty
                |> Map.toList
                |> List.map (fun (level, abstractDef) ->
                    let def = levelOverrides |> Map.tryFind level |> Option.defaultValue abstractDef
                    {
                    NumId = numId
                    Level = level
                    ParagraphStyle = def.ParagraphStyle
                    NumberFormat = def.NumberFormat
                    LevelText = def.LevelText
                    Start = startOverrides |> Map.tryFind level |> Option.orElse def.Start
                    IndentLeft = def.IndentLeft
                    HangingIndent = def.HangingIndent
                    FirstLineIndent = def.FirstLineIndent
                    RightTabPosition = def.RightTabPosition
                })
            | _ -> [])

// Document properties parsing
module Properties =
    let parseProperties (xml: string) : DocumentProperties =
        let doc = XmlDsl.parse xml
        {
            Title = Ns.childNs "title" doc |> Option.bind Xml.text
            Subject = Ns.childNs "subject" doc |> Option.bind Xml.text
            Creator = Ns.childNs "creator" doc |> Option.bind Xml.text
            Created = Ns.childNs "created" doc |> Option.bind Xml.text
            Modified = Ns.childNs "modified" doc |> Option.bind Xml.text
            Description = Ns.childNs "description" doc |> Option.bind Xml.text
        }

/// Key of an image that belongs to a header/footer part in WordDocument.Images.
let partImageKey (partRelationshipId: string) (imageRelationshipId: string) =
    partRelationshipId + "/" + imageRelationshipId

// Header/Footer parsing
module HeaderFooter =
    let parseHeaderFooter (xml: string) (hfType: string) (relationshipId: string option) : HeaderFooter =
        let doc = XmlDsl.parse xml
        // Relationship ids are local to a part: the header's rId1 is not the
        // body's rId1. Image references are qualified with the part's own
        // relationship id so they cannot collide in the document-wide image map.
        let qualify = relationshipId |> Option.map partImageKey |> Option.defaultValue id
        let rec qualifyParagraph (para: Paragraph) =
            { para with
                Runs =
                    para.Runs
                    |> List.map (fun run ->
                        match run.Content with
                        | Image image -> { run with Content = Image { image with RelationshipId = qualify image.RelationshipId } }
                        | _ -> run) }
        and qualifyTable (table: Table) =
            { table with
                Rows =
                    table.Rows
                    |> List.map (fun row ->
                        { row with
                            Cells =
                                row.Cells
                                |> List.map (fun cell ->
                                    { cell with
                                        Paragraphs = cell.Paragraphs |> List.map qualifyParagraph
                                        NestedTables = cell.NestedTables |> List.map qualifyTable }) }) }
        let paragraphs =
            Ns.descendantsNs "p" doc
            |> List.map (Docx.paragraph >> qualifyParagraph)
        let blocks =
            Docx.bodyContent doc
            |> List.map (function
                | Para para -> Para (qualifyParagraph para)
                | Tbl table -> Tbl (qualifyTable table))
        {
            Content = paragraphs
            Blocks = blocks
            Type = hfType
            RelationshipId = relationshipId
        }

// Comments parsing
module Comments =
    let parseComment (commentNode: XmlNode) : Comment option =
        Docx.attr "id" commentNode
        |> Option.map (fun id ->
            let author = Docx.attr "author" commentNode
            let date = Docx.attr "date" commentNode

            let paragraphs =
                Ns.descendantsNs "p" commentNode
                |> List.map Docx.paragraph

            let text =
                paragraphs
                |> List.collect (fun p ->
                    p.Runs |> List.choose (fun r ->
                        match r.Content with
                        | Text t -> Some t
                        | _ -> None))
                |> String.concat " "

            let initialText =
                paragraphs
                |> List.tryHead
                |> Option.bind (fun p ->
                    p.Runs
                    |> List.tryHead
                    |> Option.bind (fun r ->
                        match r.Content with
                        | Text t -> Some t
                        | _ -> None))

            {
                Id = id
                Author = author
                Date = date
                Text = text
                InitialText = initialText
            })

    let parseComments (xml: string) : Comment list =
        let doc = XmlDsl.parse xml
        Ns.descendantsNs "comment" doc
        |> List.choose parseComment

// Track Changes (Revisions) parsing
module Revisions =
    let parseRevisionType (name: string) : RevisionType =
        match name with
        | "ins" -> Insert
        | "del" -> Delete
        | "rPrChange" -> FormatChange
        | "moveFrom" -> MoveFrom
        | "moveTo" -> MoveTo
        | _ -> Insert

    let parseRevision (revNode: XmlNode) : Revision option =
        Xml.name revNode
        |> Option.bind (fun name ->
            Docx.attr "id" revNode
            |> Option.map (fun id ->
                let author = Docx.attr "author" revNode
                let date = Docx.attr "date" revNode
                let revType = parseRevisionType (Ns.stripNs name)

                // Deleted runs carry their text in <w:delText>.
                let text =
                    (Ns.descendantsNs "t" revNode) @ (Ns.descendantsNs "delText" revNode)
                    |> List.choose Xml.text
                    |> String.concat " "

                {
                    Id = id
                    Author = author
                    Date = date
                    RevisionType = revType
                    Text = text
                }))

    let parseRevisions (doc: XmlNode) : Revision list =
        let revisionTypes = ["ins"; "del"; "rPrChange"; "moveFrom"; "moveTo"]
        revisionTypes
        |> List.collect (fun revType ->
            Ns.descendantsNs revType doc
            |> List.choose parseRevision)

// Footnotes and Endnotes parsing
module Notes =
    let parseNote (noteNode: XmlNode) : Note option =
        Docx.attr "id" noteNode
        |> Option.map (fun id ->
            let paragraphs =
                Ns.descendantsNs "p" noteNode
                |> List.map Docx.paragraph
            {
                Id = id
                Content = paragraphs
            })

    let parseNotes (xml: string) : Note list =
        let doc = XmlDsl.parse xml
        Ns.descendantsNs "footnote" doc
        |> List.choose parseNote
        |> fun footnotes ->
            let endnotes =
                Ns.descendantsNs "endnote" doc
                |> List.choose parseNote
            footnotes @ endnotes

// Shapes and Text Boxes parsing
module Shapes =
    let private emuToPixels emu = emu / 9525.0

    let private parseFloat value =
        match System.Double.TryParse(value) with
        | true, parsed -> Some parsed
        | false, _ -> None

    let private shapeNumberFromStyle name (style: string) =
        style.Split(';')
        |> Array.tryPick (fun part ->
            let text = part.Trim()
            if text.StartsWith(name + ":") then
                let raw = text.Substring(name.Length + 1).Trim()
                if raw.EndsWith("pt") then raw.Substring(0, raw.Length - 2) |> parseFloat |> Option.map (fun pt -> pt * 96.0 / 72.0)
                elif raw.EndsWith("px") then raw.Substring(0, raw.Length - 2) |> parseFloat
                else parseFloat raw
            else None)

    let private normalizeColor (value: string option) =
        match value with
        | Some "" | None -> None
        | Some color when color.StartsWith("#") -> Some color
        | Some color -> Some ("#" + color)

    let private strokeWeightFromAttr (value: string option) =
        value
        |> Option.bind (fun raw ->
            let text = raw.Trim()
            if text.EndsWith("pt") then text.Substring(0, text.Length - 2) |> parseFloat |> Option.map (fun pt -> pt * 96.0 / 72.0)
            elif text.EndsWith("px") then text.Substring(0, text.Length - 2) |> parseFloat
            else parseFloat text)

    // --- DrawingML geometry/fill parsing (a:prstGeom / a:custGeom / a:gradFill) ---

    let private srgbOf (node: XmlNode) : string option =
        Ns.descendantsNs "srgbClr" node |> List.tryHead
        |> Option.bind (Xml.attr "val")
        |> Option.map (fun v -> if v.StartsWith("#") then v else "#" + v)

    let private parsePreset (shapeNode: XmlNode) : string option =
        Ns.descendantsNs "prstGeom" shapeNode |> List.tryHead
        |> Option.bind (Xml.attr "prst")

    /// Parse an a:custGeom path into (segments, coordW, coordH).
    let private parseCustom (shapeNode: XmlNode) : (ShapePathSeg list * float * float) option =
        Ns.descendantsNs "custGeom" shapeNode |> List.tryHead
        |> Option.bind (fun cg -> Ns.descendantsNs "path" cg |> List.tryHead)
        |> Option.bind (fun path ->
            let w = Xml.attr "w" path |> Option.bind parseFloat |> Option.defaultValue 0.0
            let h = Xml.attr "h" path |> Option.bind parseFloat |> Option.defaultValue 0.0
            let segs =
                Xml.children path
                |> List.choose (fun child ->
                    let name = Xml.name child |> Option.map Ns.stripNs
                    let pts =
                        Ns.descendantsNs "pt" child
                        |> List.collect (fun p ->
                            [ Xml.attr "x" p |> Option.bind parseFloat |> Option.defaultValue 0.0
                              Xml.attr "y" p |> Option.bind parseFloat |> Option.defaultValue 0.0 ])
                    match name with
                    | Some "moveTo" -> Some { Op = "M"; Pts = pts }
                    | Some "lnTo" -> Some { Op = "L"; Pts = pts }
                    | Some "cubicBezTo" -> Some { Op = "C"; Pts = pts }
                    | Some "quadBezTo" -> Some { Op = "Q"; Pts = pts }
                    | Some "close" -> Some { Op = "Z"; Pts = [] }
                    | _ -> None)   // arcTo unsupported -> dropped (path approximates)
            if segs.IsEmpty || w <= 0.0 || h <= 0.0 then None else Some (segs, w, h))

    let private parseGradient (shapeNode: XmlNode) : ShapeGradient option =
        Ns.descendantsNs "gradFill" shapeNode |> List.tryHead
        |> Option.bind (fun gf ->
            let stops =
                Ns.descendantsNs "gs" gf
                |> List.choose (fun gs ->
                    match srgbOf gs with
                    | Some color ->
                        let pos = Xml.attr "pos" gs |> Option.bind parseFloat |> Option.defaultValue 0.0
                        Some { Offset = pos / 100000.0; Color = color }
                    | None -> None)
            if stops.IsEmpty then None
            else
                let angle =
                    Ns.descendantsNs "lin" gf |> List.tryHead
                    |> Option.bind (Xml.attr "ang")
                    |> Option.bind parseFloat
                    |> Option.map (fun a -> a / 60000.0)
                    |> Option.defaultValue 0.0
                let radial = Ns.descendantsNs "path" gf |> List.isEmpty |> not
                Some { Stops = stops; Angle = angle; Radial = radial })

    /// DrawingML shape fill colour: the spPr's own solidFill (not the line's).
    let private parseDmlFill (shapeNode: XmlNode) : string option =
        Ns.descendantsNs "spPr" shapeNode |> List.tryHead
        |> Option.bind (fun spPr ->
            Xml.children spPr
            |> List.tryPick (fun child ->
                match Xml.name child |> Option.map Ns.stripNs with
                | Some "solidFill" -> srgbOf child
                | _ -> None))

    /// DrawingML line colour from spPr > a:ln > solidFill.
    let private parseDmlLine (shapeNode: XmlNode) : string option =
        Ns.descendantsNs "spPr" shapeNode |> List.tryHead
        |> Option.bind (fun spPr ->
            Xml.children spPr |> List.tryPick (fun c -> if Xml.name c |> Option.map Ns.stripNs = Some "ln" then Some c else None))
        |> Option.bind srgbOf

    let private parseRotation (shapeNode: XmlNode) : float =
        Ns.descendantsNs "xfrm" shapeNode |> List.tryHead
        |> Option.bind (Xml.attr "rot")
        |> Option.bind parseFloat
        |> Option.map (fun r -> r / 60000.0)
        |> Option.defaultValue 0.0

    let parseShape (shapeNode: XmlNode) : Shape option =
        let id = Xml.attr "id" shapeNode
        let shapeType = Xml.attr "type" shapeNode

        let txbxContent = Ns.descendantsNs "txbxContent" shapeNode |> List.tryHead
        let paragraphs =
            match txbxContent with
            | Some content ->
                Ns.descendantsNs "p" content
                |> List.map Docx.paragraph
            | None -> []

        let extent = Ns.descendantsNs "extent" shapeNode |> List.tryHead
        let style = Xml.attr "style" shapeNode
        let width =
            extent |> Option.bind (Xml.attr "cx") |> Option.bind parseFloat |> Option.map emuToPixels
            |> Option.orElseWith (fun () -> style |> Option.bind (shapeNumberFromStyle "width"))
        let height =
            extent |> Option.bind (Xml.attr "cy") |> Option.bind parseFloat |> Option.map emuToPixels
            |> Option.orElseWith (fun () -> style |> Option.bind (shapeNumberFromStyle "height"))

        let posH = Ns.descendantsNs "positionH" shapeNode |> List.tryHead
        let posV = Ns.descendantsNs "positionV" shapeNode |> List.tryHead
        let x =
            posH |> Option.bind (Ns.childNs "posOffset") |> Option.bind Xml.text |> Option.bind parseFloat |> Option.map emuToPixels
            |> Option.orElseWith (fun () -> style |> Option.bind (shapeNumberFromStyle "margin-left"))
        let y =
            posV |> Option.bind (Ns.childNs "posOffset") |> Option.bind Xml.text |> Option.bind parseFloat |> Option.map emuToPixels
            |> Option.orElseWith (fun () -> style |> Option.bind (shapeNumberFromStyle "margin-top"))
        let position = Option.map2 (fun x y -> (x, y)) x y
        // DrawingML xfrm fallback for size/offset (a:ext / a:off) when the wp:
        // anchor extent/position is absent (bare wsp shapes).
        let xfrm = Ns.descendantsNs "xfrm" shapeNode |> List.tryHead
        let extOf node tag = xfrm |> Option.bind (Ns.childNs node) |> Option.bind (Xml.attr tag) |> Option.bind parseFloat |> Option.map emuToPixels
        let width = width |> Option.orElseWith (fun () -> extOf "ext" "cx")
        let height = height |> Option.orElseWith (fun () -> extOf "ext" "cy")
        let fillColor =
            (Xml.attr "fillcolor" shapeNode |> normalizeColor) |> Option.orElseWith (fun () -> parseDmlFill shapeNode)
        let borderColor =
            (Xml.attr "strokecolor" shapeNode |> normalizeColor) |> Option.orElseWith (fun () -> parseDmlLine shapeNode)
        let borderWidth = Xml.attr "strokeweight" shapeNode |> strokeWeightFromAttr
        let custom = parseCustom shapeNode
        // VML shapes use a "#_x0000_t202"-style type ref; only surface a clean
        // DrawingML preset name as the geometry hint.
        let preset = parsePreset shapeNode

        Some {
            Id = id
            Type = shapeType
            Content = paragraphs
            Width = width
            Height = height
            Position = position
            FillColor = fillColor
            BorderColor = borderColor
            BorderWidth = borderWidth
            Preset = preset
            CustomPath = custom |> Option.map (fun (s, _, _) -> s) |> Option.defaultValue []
            CustomW = custom |> Option.map (fun (_, w, _) -> w)
            CustomH = custom |> Option.map (fun (_, _, h) -> h)
            Gradient = parseGradient shapeNode
            Rotation = parseRotation shapeNode
        }

    let parseShapes (doc: XmlNode) : Shape list =
        // Import a node if it carries text (txbxContent) OR a real geometry
        // (prstGeom/custGeom) — so autoshapes with no text still render.
        // A plain picture also has a:prstGeom (inside pic:spPr). It is already
        // read as an image run; without this guard every picture was imported
        // a second time as a rectangle shape and saved as a text box.
        let isPicture node =
            (Ns.descendantsNs "pic" node |> List.isEmpty |> not)
            && (Ns.descendantsNs "wsp" node |> List.isEmpty)
        let hasContent node =
            (Ns.descendantsNs "txbxContent" node |> List.isEmpty |> not)
            || (not (isPicture node)
                && ((Ns.descendantsNs "prstGeom" node |> List.isEmpty |> not)
                    || (Ns.descendantsNs "custGeom" node |> List.isEmpty |> not)))
        // For DrawingML, parse from the wp:anchor / wp:inline wrapper — position
        // (wp:positionH/V) and display size (wp:extent) live there, while the
        // geometry/fill live in the nested wsp. parseShape reads everything via
        // descendant lookups, so passing the whole wrapper picks up both. VML
        // shapes carry their attrs on the v:shape node itself, so pass those.
        (Ns.descendantsNs "anchor" doc) @ (Ns.descendantsNs "inline" doc) @ (Ns.descendantsNs "shape" doc)
        |> List.filter hasContent
        |> List.choose parseShape

// Resolve hyperlink targets from relationships
let resolveHyperlinks (doc: WordDocument) : WordDocument =
    let relMap =
        doc.Relationships
        |> List.choose (fun rel ->
            if rel.Target <> "" then Some (rel.Id, rel.Target) else None)
        |> Map.ofList

    let resolveRun (run: Run) : Run =
        match run.Content with
        | Link link ->
            let target =
                link.RelationshipId
                |> Option.bind (fun rId -> Map.tryFind rId relMap)
            { run with Content = Link { link with Target = target } }
        | _ -> run

    let resolvePara (para: Paragraph) : Paragraph =
        { para with Runs = para.Runs |> List.map resolveRun }

    let rec resolveTable table =
        let resolveCell (cell: TableCell) =
            { cell with
                Paragraphs = cell.Paragraphs |> List.map resolvePara
                NestedTables = cell.NestedTables |> List.map resolveTable }
        let resolveRow (row: TableRow) =
            { row with Cells = row.Cells |> List.map resolveCell }
        { table with Rows = table.Rows |> List.map resolveRow }

    let resolveContent = function
        | Para p -> Para (resolvePara p)
        | Tbl t -> Tbl (resolveTable t)

    { doc with Content = doc.Content |> List.map resolveContent }

// Determine MIME type from file extension
let getMimeType (path: string) : string =
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

let private wordTwipsToPixels value =
    value
    |> Option.bind (fun text ->
        match System.Double.TryParse(text) with
        | true, parsed -> Some (parsed / 15.0)
        | false, _ -> None)

let private parseSettings (xml: string option) =
    xml
    |> Option.map XmlDsl.parse
    |> Option.map (fun settings ->
        // An on/off element: <w:evenAndOddHeaders w:val="0"/> means off.
        match Ns.descendantsNs "evenAndOddHeaders" settings |> List.tryHead with
        | Some node ->
            match Docx.attr "val" node |> Option.map (fun value -> value.ToLowerInvariant()) with
            | Some "0" | Some "false" | Some "off" -> false
            | _ -> true
        | None -> false)
    |> Option.defaultValue false

let private parsePageSetup evenAndOddHeaders (sectPr: XmlNode option) : PageSetup option =
    sectPr
    |> Option.map (fun section ->
        let pgSz = Ns.childNs "pgSz" section
        let pgMar = Ns.childNs "pgMar" section
        let columns = Docx.parseColumns section
        {
            PageWidth = pgSz |> Option.bind (Docx.attr "w") |> wordTwipsToPixels
            PageHeight = pgSz |> Option.bind (Docx.attr "h") |> wordTwipsToPixels
            MarginLeft = pgMar |> Option.bind (Docx.attr "left") |> wordTwipsToPixels
            MarginRight = pgMar |> Option.bind (Docx.attr "right") |> wordTwipsToPixels
            MarginTop = pgMar |> Option.bind (Docx.attr "top") |> wordTwipsToPixels
            MarginBottom = pgMar |> Option.bind (Docx.attr "bottom") |> wordTwipsToPixels
            HeaderDistance = pgMar |> Option.bind (Docx.attr "header") |> wordTwipsToPixels
            FooterDistance = pgMar |> Option.bind (Docx.attr "footer") |> wordTwipsToPixels
            DifferentFirstPage = Some (Docx.onOff "titlePg" section |> Option.defaultValue false)
            EvenAndOddHeaders = Some evenAndOddHeaders
            Columns = columns
            SectionStart = Docx.getPropVal "type" "val" section
        })

// Extract image as data URL from archive
let extractImageDataUrl (archive: ZipArchive) (relationshipId: string) (relationships: Relationship list) : Fable.Core.JS.Promise<string option> =
    promise {
        // Find relationship
        let relOpt = relationships |> List.tryFind (fun r -> r.Id = relationshipId)
        match relOpt with
        | Some rel ->
            if rel.TargetMode = Some "External" then
                return Some rel.Target
            else
                // Normalize path: relationships are relative to word/ folder
                let imagePath = ZipInterop.resolveTarget "word" rel.Target
                let mimeType = getMimeType rel.Target

                // Extract from archive
                return! getFileAsDataUrl archive imagePath mimeType
        | None ->
            return None
    }

let private lastPathSegment (path: string) =
    let parts = (normalizePath path).Split('/')
    parts.[parts.Length - 1]

let private parentPath (path: string) =
    let normalized = normalizePath path
    let index = normalized.LastIndexOf('/')
    if index > 0 then normalized.Substring(0, index) else ""

let private resolveRelationshipTarget baseDir target =
    ZipInterop.resolveTarget baseDir target

let private relationshipTargetBase (rel: Relationship) =
    parentPath (ZipInterop.resolveTarget "word" rel.Target)

let private loadPartRelationships archive (rel: Relationship) =
    promise {
        let relsPath = normalizePath ("word/_rels/" + lastPathSegment rel.Target + ".rels")
        let! xml = getFileText archive relsPath
        return
            xml
            |> Option.map XmlDsl.parse
            |> Option.map parseRelationships
            |> Option.defaultValue []
    }

let private loadImageData archive baseDir (rels: Relationship list) =
    rels
    |> List.filter (fun r -> r.Type.Contains("/image"))
    |> List.map (fun rel ->
        promise {
            if rel.TargetMode = Some "External" then
                return Some (rel.Id, rel.Target)
            else
                let imagePath = resolveRelationshipTarget baseDir rel.Target
                let mimeType = getMimeType rel.Target
                let! dataUrl = getFileAsDataUrl archive imagePath mimeType
                return dataUrl |> Option.map (fun url -> (rel.Id, url))
        })
    |> List.toArray
    |> Promise.all

let private getFileAsDataUrlNative (archive: ZipArchive) (path: string) (mimeType: string) : string option =
    getFileBytesNative archive path
    |> Option.map (fun bytes -> "data:" + mimeType + ";base64," + Convert.ToBase64String(bytes))

let private loadPartRelationshipsNative archive (rel: Relationship) =
    let relsPath = normalizePath ("word/_rels/" + lastPathSegment rel.Target + ".rels")
    getFileTextNative archive relsPath
    |> Option.map XmlDsl.parse
    |> Option.map parseRelationships
    |> Option.defaultValue []

let private loadImageDataNative archive baseDir (rels: Relationship list) =
    rels
    |> List.choose (fun rel ->
        if rel.Type.Contains("/image") then
            if rel.TargetMode = Some "External" then
                Some (rel.Id, rel.Target)
            else
                let imagePath = resolveRelationshipTarget baseDir rel.Target
                let mimeType = getMimeType rel.Target
                getFileAsDataUrlNative archive imagePath mimeType
                |> Option.map (fun url -> (rel.Id, url))
        else
            None)

// Parse DOCX archive
let parseDocument (archive: ZipArchive) : Fable.Core.JS.Promise<WordDocument option> =
    promise {
        // Load main document parts
        let! docXml = getFileText archive "word/document.xml"
        let! relsXml = getFileText archive "word/_rels/document.xml.rels"
        let! stylesXml = getFileText archive "word/styles.xml"
        let! themeXml = getFileText archive "word/theme/theme1.xml"
        let! numberingXml = getFileText archive "word/numbering.xml"
        let! settingsXml = getFileText archive "word/settings.xml"
        let! propsXml = getFileText archive "docProps/core.xml"

        // Load advanced features
        let! commentsXml = getFileText archive "word/comments.xml"
        let! footnotesXml = getFileText archive "word/footnotes.xml"
        let! endnotesXml = getFileText archive "word/endnotes.xml"

        match docXml with
        | Some xml ->
            let doc = XmlDsl.parse xml
            let body = Ns.childNs "body" doc

            let rels =
                relsXml
                |> Option.map XmlDsl.parse
                |> Option.map parseRelationships
                |> Option.defaultValue []

            let styles =
                stylesXml
                |> Option.map Styles.parseStyles
                |> Option.defaultValue []

            let theme = themeXml |> Option.map Theme.parseTheme

            let numberingDefinitions =
                numberingXml
                |> Option.map Numbering.parseNumbering
                |> Option.defaultValue []

            let evenAndOddHeaders = parseSettings settingsXml

            let props =
                propsXml
                |> Option.map Properties.parseProperties

            let headerRefTypes =
                Ns.descendantsNs "headerReference" doc
                |> List.choose (fun refNode ->
                    Xml.attr "r:id" refNode
                    |> Option.map (fun id -> (id, Docx.attr "type" refNode |> Option.defaultValue "default")))
                |> Map.ofList

            let footerRefTypes =
                Ns.descendantsNs "footerReference" doc
                |> List.choose (fun refNode ->
                    Xml.attr "r:id" refNode
                    |> Option.map (fun id -> (id, Docx.attr "type" refNode |> Option.defaultValue "default")))
                |> Map.ofList

            let loadHeaderFooter (rel: Relationship) (hfType: string) =
                promise {
                    let path = ZipInterop.resolveTarget "word" rel.Target
                    let! xml = getFileText archive path
                    return xml |> Option.map (fun value -> HeaderFooter.parseHeaderFooter value hfType (Some rel.Id))
                }

            let! headers =
                rels
                |> List.filter (fun rel -> rel.Type.Contains("/header"))
                |> List.map (fun rel -> loadHeaderFooter rel (headerRefTypes |> Map.tryFind rel.Id |> Option.defaultValue "default"))
                |> List.toArray
                |> Promise.all

            let! footers =
                rels
                |> List.filter (fun rel -> rel.Type.Contains("/footer"))
                |> List.map (fun rel -> loadHeaderFooter rel (footerRefTypes |> Map.tryFind rel.Id |> Option.defaultValue "default"))
                |> List.toArray
                |> Promise.all

            let content =
                body
                |> Option.map (fun bodyNode -> Docx.bodyContentWithFinalSection (Ns.childNs "sectPr" bodyNode) bodyNode)
                |> Option.defaultValue []

            let pageSetup =
                body
                |> Option.bind (Ns.childNs "sectPr")
                |> parsePageSetup evenAndOddHeaders

            // Parse advanced features
            let comments =
                commentsXml
                |> Option.map Comments.parseComments
                |> Option.defaultValue []

            let revisions = Revisions.parseRevisions doc

            let footnotes =
                footnotesXml
                |> Option.map Notes.parseNotes
                |> Option.defaultValue []

            let endnotes =
                endnotesXml
                |> Option.map Notes.parseNotes
                |> Option.defaultValue []

            let shapes = Shapes.parseShapes doc

            // Load image data from zip
            let! images =
                promise {
                    let partRels =
                        rels
                        |> List.filter (fun r -> r.Type.Contains("/header") || r.Type.Contains("/footer"))

                    let! bodyImages = loadImageData archive "word" rels
                    let! runningContentImages =
                        partRels
                        |> List.map (fun partRel ->
                            promise {
                                let! partRelationships = loadPartRelationships archive partRel
                                let! partImages = loadImageData archive (relationshipTargetBase partRel) partRelationships
                                return partImages |> Array.map (Option.map (fun (relId, dataUrl) -> partImageKey partRel.Id relId, dataUrl))
                            })
                        |> List.toArray
                        |> Promise.all

                    return
                        Array.concat (bodyImages :: (runningContentImages |> Array.toList))
                        |> Array.choose id
                        |> Map.ofArray
                }

            let document = {
                Content = content
                Relationships = rels
                Styles = styles
                Theme = theme
                Headers = headers |> Array.choose id |> Array.toList
                Footers = footers |> Array.choose id |> Array.toList
                Properties = props
                Comments = comments
                Revisions = revisions
                Footnotes = footnotes
                Endnotes = endnotes
                Shapes = shapes
                Images = images
                PageSetup = pageSetup
                NumberingDefinitions = numberingDefinitions
            }

            return Some (resolveHyperlinks document)

        | None -> return None
    }

let parseDocumentNative (archive: ZipArchive) : WordDocument option =
    let docXml = getFileTextNative archive "word/document.xml"
    let relsXml = getFileTextNative archive "word/_rels/document.xml.rels"
    let stylesXml = getFileTextNative archive "word/styles.xml"
    let themeXml = getFileTextNative archive "word/theme/theme1.xml"
    let numberingXml = getFileTextNative archive "word/numbering.xml"
    let settingsXml = getFileTextNative archive "word/settings.xml"
    let propsXml = getFileTextNative archive "docProps/core.xml"
    let commentsXml = getFileTextNative archive "word/comments.xml"
    let footnotesXml = getFileTextNative archive "word/footnotes.xml"
    let endnotesXml = getFileTextNative archive "word/endnotes.xml"

    docXml
    |> Option.map (fun xml ->
        let doc = XmlDsl.parse xml
        let body = Ns.childNs "body" doc

        let rels =
            relsXml
            |> Option.map XmlDsl.parse
            |> Option.map parseRelationships
            |> Option.defaultValue []

        let styles =
            stylesXml
            |> Option.map Styles.parseStyles
            |> Option.defaultValue []

        let theme = themeXml |> Option.map Theme.parseTheme

        let numberingDefinitions =
            numberingXml
            |> Option.map Numbering.parseNumbering
            |> Option.defaultValue []

        let evenAndOddHeaders = parseSettings settingsXml
        let props = propsXml |> Option.map Properties.parseProperties

        let headerRefTypes =
            Ns.descendantsNs "headerReference" doc
            |> List.choose (fun refNode ->
                Xml.attr "r:id" refNode
                |> Option.map (fun id -> (id, Docx.attr "type" refNode |> Option.defaultValue "default")))
            |> Map.ofList

        let footerRefTypes =
            Ns.descendantsNs "footerReference" doc
            |> List.choose (fun refNode ->
                Xml.attr "r:id" refNode
                |> Option.map (fun id -> (id, Docx.attr "type" refNode |> Option.defaultValue "default")))
            |> Map.ofList

        let loadHeaderFooterNative (rel: Relationship) (hfType: string) =
            let path = ZipInterop.resolveTarget "word" rel.Target
            getFileTextNative archive path
            |> Option.map (fun value -> HeaderFooter.parseHeaderFooter value hfType (Some rel.Id))

        let headers =
            rels
            |> List.filter (fun rel -> rel.Type.Contains("/header"))
            |> List.choose (fun rel -> loadHeaderFooterNative rel (headerRefTypes |> Map.tryFind rel.Id |> Option.defaultValue "default"))

        let footers =
            rels
            |> List.filter (fun rel -> rel.Type.Contains("/footer"))
            |> List.choose (fun rel -> loadHeaderFooterNative rel (footerRefTypes |> Map.tryFind rel.Id |> Option.defaultValue "default"))

        let content =
            body
            |> Option.map (fun bodyNode -> Docx.bodyContentWithFinalSection (Ns.childNs "sectPr" bodyNode) bodyNode)
            |> Option.defaultValue []

        let pageSetup =
            body
            |> Option.bind (Ns.childNs "sectPr")
            |> parsePageSetup evenAndOddHeaders

        let comments = commentsXml |> Option.map Comments.parseComments |> Option.defaultValue []
        let revisions = Revisions.parseRevisions doc
        let footnotes = footnotesXml |> Option.map Notes.parseNotes |> Option.defaultValue []
        let endnotes = endnotesXml |> Option.map Notes.parseNotes |> Option.defaultValue []
        let shapes = Shapes.parseShapes doc

        let partRels = rels |> List.filter (fun r -> r.Type.Contains("/header") || r.Type.Contains("/footer"))
        let bodyImages = loadImageDataNative archive "word" rels
        let runningContentImages =
            partRels
            |> List.collect (fun partRel ->
                let partRelationships = loadPartRelationshipsNative archive partRel
                loadImageDataNative archive (relationshipTargetBase partRel) partRelationships
                |> List.map (fun (relId, dataUrl) -> partImageKey partRel.Id relId, dataUrl))

        let images =
            (bodyImages @ runningContentImages)
            |> Map.ofList

        {
            Content = content
            Relationships = rels
            Styles = styles
            Theme = theme
            Headers = headers
            Footers = footers
            Properties = props
            Comments = comments
            Revisions = revisions
            Footnotes = footnotes
            Endnotes = endnotes
            Shapes = shapes
            Images = images
            PageSetup = pageSetup
            NumberingDefinitions = numberingDefinitions
        }
        |> resolveHyperlinks)

// Public API for loading DOCX from data
let load (data: obj) : Fable.Core.JS.Promise<WordDocument option> =
    promise {
        let! archive = loadArchive data
        return! parseDocument archive
    }

let loadBytesNative (data: byte[]) : WordDocument option =
    data
    |> loadArchiveNative
    |> parseDocumentNative

// Extract all text with functional pipeline
let extractText (doc: WordDocument) : string =
    let runText (run: Run) =
        match run.Content with
        | Text t -> t
        | Link link -> link.Text
        | Image img -> img.Description |> Option.defaultValue "[Image]"
        | CommentReference id -> $"[Comment:{id}]"
        | NoteRef ref -> $"[{ref.NoteType}:{ref.Id}]"
        | BookmarkStart _ -> ""
        | BookmarkEnd _ -> ""
        | FieldRef field -> field.Result |> Option.defaultValue $"[Field:{field.Code}]"
        | EquationXml _ -> "[Equation]"
        | RenderedPageBreak -> ""
        | PageBreak -> "\n"

    let rec paraText (para: Paragraph) =
        para.Runs |> List.map runText |> String.concat ""

    let rec contentText = function
        | Para p -> paraText p
        | Tbl t ->
            t.Rows
            |> List.collect (fun row ->
                row.Cells
                |> List.collect (fun cell ->
                    cell.Paragraphs
                    |> List.map paraText))
            |> String.concat "\n"

    doc.Content
    |> List.map contentText
    |> String.concat "\n\n"

// Convert DocumentContent to EditorParagraph format for Word editor
let private wordFontSizeToPoints (fontSize: string option) =
    match fontSize with
    | Some value ->
        match System.Double.TryParse(value) with
        | true, halfPoints -> (halfPoints / 2.0) * 96.0 / 72.0
        | false, _ -> 11.0 * 96.0 / 72.0
    | None -> 11.0 * 96.0 / 72.0

/// w:position is in half-points and the editor (Paint) takes points. It used to
/// go through wordFontSizeToPoints, which turned it into px and, worse, turned
/// "no w:position" into the 11pt font-size default instead of 0.
let private wordPositionToPoints (position: string option) =
    match position with
    | Some value ->
        match System.Double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture) with
        | true, halfPoints -> halfPoints / 2.0
        | false, _ -> 0.0
    | None -> 0.0

let private wordCharacterSpacingToCssPixels (spacing: string option) =
    spacing
    |> Option.bind (fun value ->
        match System.Double.TryParse(value) with
        | true, twips -> Some (twips / 15.0)
        | false, _ -> None)

let private resolveThemeFont (theme: Theme option) (themeFont: string option) =
    match theme, themeFont with
    | Some theme, Some value when value.StartsWith("major") -> theme.MajorFont
    | Some theme, Some value when value.StartsWith("minor") -> theme.MinorFont
    | _, Some value when value.StartsWith("major") -> Some "Cambria"
    | _, Some value when value.StartsWith("minor") -> Some "Calibri"
    | _ -> None

let private wordFontFamilyToCss (theme: Theme option) (props: RunProperties) =
    props.FontFamily
    |> Option.orElseWith (fun () -> resolveThemeFont theme props.ThemeFont)
    |> Option.defaultValue "Calibri"

let private wordColorToCss (color: string option) =
    match color with
    | Some value when value = "" || value = "auto" -> None
    | Some value when value.StartsWith("#") -> Some value
    | Some value when value.Length = 8 -> Some ("#" + value.Substring(2).ToUpper())
    | Some value -> Some ("#" + value.ToUpper())
    | None -> None

let private wordShadingColorToEditor (color: string option) =
    match color with
    | Some value when value = "" -> None
    | Some "auto" -> Some "auto"
    | value -> wordColorToCss value

let private wordShadingToEditor (shading: Shading) =
    objFromFields [
        "Pattern" ==> shading.Pattern
        "ForegroundColor" ==> wordShadingColorToEditor shading.Color
        "BackgroundColor" ==> wordShadingColorToEditor shading.Fill
    ]

let private directShadingBackgroundColor (shading: Shading) =
    match shading.Pattern with
    | None
    | Some "clear" -> wordColorToCss shading.Fill
    | _ -> None

let private wordTwipsToPoints (twips: string option) =
    twips
    |> Option.bind (fun value ->
        match System.Double.TryParse(value) with
        | true, parsed -> Some (parsed / 15.0)
        | false, _ -> None)

let private wordSpacingTwipsToPoints (twips: string option) =
    twips
    |> Option.bind (fun value ->
        match System.Double.TryParse(value) with
        | true, parsed -> Some (parsed / 20.0)
        | false, _ -> None)

let private wordAutoSpacing = 14.0

let private wordLineSpacingToMultiple fontSize (line: string option) (lineRule: string option) =
    line
    |> Option.bind (fun value ->
        match System.Double.TryParse(value) with
        | true, parsed ->
            match lineRule with
            | Some rule when rule = "exact" || rule = "atLeast" -> Some ((parsed / 15.0) / max 1.0 fontSize)
            | _ -> Some (parsed / 240.0)
        | false, _ -> None)

let private borderColorToCss (border: BorderStyle option) =
    border
    |> Option.bind (fun border ->
        match border.Color with
        | Some value when value = "" || value = "auto" -> Some "#000000"
        | value -> wordColorToCss value)

let private borderWidthToPoints (border: BorderStyle option) =
    border
    |> Option.bind (fun border -> border.Size)
    |> Option.bind (fun value ->
        match System.Double.TryParse(value) with
        | true, parsed -> Some (parsed / 6.0)
        | false, _ -> None)

let private borderToEditor (border: BorderStyle option) =
    border
    |> Option.map (fun value -> objFromFields [
        "Style" ==> value.BorderType
        "Color" ==> borderColorToCss border
        "Width" ==> borderWidthToPoints border
    ])

let private paragraphBorderColorToCss (border: ParagraphBorder option) =
    border
    |> Option.bind (fun border ->
        match border.Color with
        | Some value when value = "" || value = "auto" -> Some "#000000"
        | value -> wordColorToCss value)

let private paragraphBorderWidthToPoints (border: ParagraphBorder option) =
    border
    |> Option.bind (fun border -> border.Width)
    |> Option.bind (fun value ->
        match System.Double.TryParse(value) with
        | true, parsed -> Some (parsed / 6.0)
        | false, _ -> None)

let private paragraphBorderSpaceToPoints (border: ParagraphBorder option) =
    border
    |> Option.bind (fun border -> border.Space)
    |> Option.bind (fun value ->
        match System.Double.TryParse(value) with
        | true, parsed -> Some parsed
        | false, _ -> None)

let private paragraphBorderAdvanceToLayoutPx (border: ParagraphBorder option) =
    border
    |> Option.bind (fun value ->
        match value.Style |> Option.map (fun style -> style.ToLowerInvariant()) with
        | Some "none"
        | Some "nil" -> None
        | _ ->
            let widthPt =
                value.Width
                |> Option.bind (fun raw ->
                    match System.Double.TryParse(raw) with
                    | true, parsed -> Some (parsed / 8.0)
                    | false, _ -> None)
                |> Option.defaultValue 0.5
            let spacePt =
                value.Space
                |> Option.bind (fun raw ->
                    match System.Double.TryParse(raw) with
                    | true, parsed -> Some parsed
                    | false, _ -> None)
                |> Option.defaultValue 0.0
            Some ((spacePt + widthPt * 2.0) * 96.0 / 72.0))

let private paragraphBorderToEditor (border: ParagraphBorder option) =
    border
    |> Option.map (fun value -> objFromFields [
        "Style" ==> value.Style
        "Color" ==> paragraphBorderColorToCss border
        "Width" ==> paragraphBorderWidthToPoints border
        "Space" ==> paragraphBorderSpaceToPoints border
    ])

let private paragraphBordersToEditor (borders: ParagraphBorders option) =
    borders
    |> Option.map (fun value -> objFromFields [
        "Top" ==> paragraphBorderToEditor value.Top
        "Right" ==> paragraphBorderToEditor value.Right
        "Bottom" ==> paragraphBorderToEditor value.Bottom
        "Left" ==> paragraphBorderToEditor value.Left
    ])

let private bordersToEditor (borders: CellBorders option) =
    borders
    |> Option.map (fun value -> objFromFields [
        "Top" ==> borderToEditor value.Top
        "Right" ==> borderToEditor value.Right
        "Bottom" ==> borderToEditor value.Bottom
        "Left" ==> borderToEditor value.Left
        "InsideH" ==> borderToEditor value.InsideH
        "InsideV" ==> borderToEditor value.InsideV
    ])

let private paragraphLayoutProps fontSize defaultMissingSpacingToZero (borderAdvanceSource: ParagraphBorders option) (props: ParagraphProperties) =
    let missingSpacingFallback = if defaultMissingSpacingToZero then Some 0.0 else None
    let spacingAfter =
        if props.SpacingAfterAuto = Some true then Some wordAutoSpacing
        else wordSpacingTwipsToPoints props.SpacingAfter |> Option.orElse missingSpacingFallback
    let firstLineIndent =
        match wordTwipsToPoints props.FirstLineIndent, wordTwipsToPoints props.HangingIndent with
        | Some firstLine, _ -> Some firstLine
        | None, Some hanging -> Some -hanging
        | None, None -> None
    [
        "ContentControlId" ==> props.ContentControlId
        "ContentControlTag" ==> props.ContentControlTag
        "ContentControlAlias" ==> props.ContentControlAlias
        "IndentLeft" ==> wordTwipsToPoints props.IndentLeft
        "IndentRight" ==> wordTwipsToPoints props.IndentRight
        "FirstLineIndent" ==> firstLineIndent
        "RightTabPosition" ==> wordTwipsToPoints props.RightTabPosition
        "RightTabLeader" ==> props.RightTabLeader
        "TabStops" ==> (props.TabStops
                        |> List.map (fun (pos, align, leader) ->
                            objFromFields [ "Position" ==> wordTwipsToPoints (Some pos); "Alignment" ==> align; "Leader" ==> leader ])
                        |> List.toArray)
        "Rtl" ==> (props.Rtl |> Option.defaultValue false)
        "SpacingBefore" ==> (if props.SpacingBeforeAuto = Some true then Some wordAutoSpacing else wordSpacingTwipsToPoints props.SpacingBefore |> Option.orElse missingSpacingFallback)
        "SpacingAfter" ==> spacingAfter
        "LineSpacing" ==> wordLineSpacingToMultiple fontSize props.LineSpacing props.LineRule
        "LineSpacingRule" ==> props.LineRule
        "BackgroundColor" ==> wordColorToCss props.BackgroundColor
        "ParagraphBorders" ==> paragraphBordersToEditor props.ParagraphBorders
        "ParagraphBorderBefore" ==> (borderAdvanceSource |> Option.bind (fun borders -> paragraphBorderAdvanceToLayoutPx borders.Top))
        "ParagraphBorderAfter" ==> (borderAdvanceSource |> Option.bind (fun borders -> paragraphBorderAdvanceToLayoutPx borders.Bottom))
        "RenderedPageBreakBefore" ==> props.RenderedPageBreakBefore
        // Geometry and columns of the section this paragraph ends (its own w:sectPr).
        // The native importer has always delivered these; without them a saved
        // multi-section document lost every section's page size but the last.
        "SectionPageSetup" ==> (props.SectionGeometry
                                |> Option.map (fun geometry ->
                                    objFromFields [
                                        "PageWidth" ==> geometry.PageWidth
                                        "PageHeight" ==> geometry.PageHeight
                                        "MarginLeft" ==> geometry.MarginLeft
                                        "MarginRight" ==> geometry.MarginRight
                                        "MarginTop" ==> geometry.MarginTop
                                        "MarginBottom" ==> geometry.MarginBottom
                                        "HeaderDistance" ==> geometry.HeaderDistance
                                        "FooterDistance" ==> geometry.FooterDistance
                                        "DifferentFirstPage" ==> geometry.DifferentFirstPage
                                    ]))
        "SectionColumns" ==> (props.SectionColumns
                              |> Option.map (fun columns ->
                                  objFromFields [
                                      "Count" ==> columns.Count
                                      "Gap" ==> columns.Gap
                                      "Separator" ==> columns.Separator
                                      "Widths" ==> (columns.Widths |> List.map (fun (width, space) -> objFromFields [ "W" ==> width; "S" ==> space ]) |> Array.ofList)
                                  ]))
    ]

let private directParagraphFormattingResetsMissingSpacing (props: ParagraphProperties) =
    props.IndentLeft.IsSome
    || props.IndentRight.IsSome
    || props.FirstLineIndent.IsSome
    || props.HangingIndent.IsSome
    || props.RightTabPosition.IsSome
    || props.RightTabLeader.IsSome
    || props.ParagraphRunProperties.IsSome
    || props.BackgroundColor.IsSome
    || props.ParagraphBorders.IsSome
    || props.ContextualSpacing.IsSome

let private tableStyleParagraphResetsMissingSpacing (props: ParagraphProperties) =
    props.LineSpacing.IsSome || props.LineRule.IsSome

let private hasDirectParagraphFormatting (props: ParagraphProperties) =
    props.Style.IsSome
    || props.Alignment.IsSome
    || props.IndentLeft.IsSome
    || props.IndentRight.IsSome
    || props.FirstLineIndent.IsSome
    || props.HangingIndent.IsSome
    || props.RightTabPosition.IsSome
    || props.RightTabLeader.IsSome
    || props.SpacingBefore.IsSome
    || props.SpacingBeforeAuto.IsSome
    || props.SpacingAfter.IsSome
    || props.SpacingAfterAuto.IsSome
    || props.LineSpacing.IsSome
    || props.LineRule.IsSome
    || props.OutlineLevel.IsSome
    || props.KeepNext.IsSome
    || props.KeepLines.IsSome
    || props.PageBreakBefore.IsSome
    || props.WidowControl.IsSome
    || props.ContextualSpacing.IsSome
    || props.BackgroundColor.IsSome
    || props.ParagraphRunProperties.IsSome
    || props.ParagraphBorders.IsSome
    || props.ListInfo.IsSome
    || props.SectionBreak.IsSome
    || props.SectionPageNumberStart.IsSome
    || props.RenderedPageBreakBefore.IsSome
    || not props.HeaderReferences.IsEmpty
    || not props.FooterReferences.IsEmpty

let private isPureEmptyParagraph (para: Paragraph) =
    para.Runs.IsEmpty && not (hasDirectParagraphFormatting para.Properties)

let private referenceObjects (refs: HeaderFooterReference list) =
    refs
    |> List.map (fun ref -> objFromFields [ "Kind" ==> ref.Kind; "RelationshipId" ==> ref.RelationshipId ])
    |> List.toArray

let private emptyRunProperties = {
    Style = None
    Bold = None
    Italic = None
    Underline = None
    UnderlineStyle = None
    UnderlineColor = None
    AllCaps = None
    SmallCaps = None
    FontSize = None
    CharacterSpacing = None
    FontFamily = None
    ThemeFont = None
    Color = None
    ThemeColor = None
    ThemeTint = None
    ThemeShade = None
    Highlight = None
    Strike = None
    DoubleStrike = None
    Hidden = None
    Position = None
    EmphasisMark = None
    VertAlign = None
}

let private mergeRunProperties (baseProps: RunProperties) (props: RunProperties) : RunProperties = {
    Style = props.Style |> Option.orElse baseProps.Style
    Bold = props.Bold |> Option.orElse baseProps.Bold
    Italic = props.Italic |> Option.orElse baseProps.Italic
    Underline = props.Underline |> Option.orElse baseProps.Underline
    UnderlineStyle = props.UnderlineStyle |> Option.orElse baseProps.UnderlineStyle
    UnderlineColor = props.UnderlineColor |> Option.orElse baseProps.UnderlineColor
    AllCaps = props.AllCaps |> Option.orElse baseProps.AllCaps
    SmallCaps = props.SmallCaps |> Option.orElse baseProps.SmallCaps
    Strike = props.Strike |> Option.orElse baseProps.Strike
    DoubleStrike = props.DoubleStrike |> Option.orElse baseProps.DoubleStrike
    Hidden = props.Hidden |> Option.orElse baseProps.Hidden
    Position = props.Position |> Option.orElse baseProps.Position
    EmphasisMark = props.EmphasisMark |> Option.orElse baseProps.EmphasisMark
    FontSize = props.FontSize |> Option.orElse baseProps.FontSize
    CharacterSpacing = props.CharacterSpacing |> Option.orElse baseProps.CharacterSpacing
    FontFamily =
        match props.FontFamily, props.ThemeFont with
        | Some family, _ -> Some family
        | None, Some _ -> None
        | None, None -> baseProps.FontFamily
    ThemeFont =
        match props.FontFamily, props.ThemeFont with
        | Some _, None -> None
        | _, Some themeFont -> Some themeFont
        | None, None -> baseProps.ThemeFont
    Color = props.Color |> Option.orElse baseProps.Color
    ThemeColor = props.ThemeColor |> Option.orElse baseProps.ThemeColor
    ThemeTint = props.ThemeTint |> Option.orElse baseProps.ThemeTint
    ThemeShade = props.ThemeShade |> Option.orElse baseProps.ThemeShade
    Highlight = props.Highlight |> Option.orElse baseProps.Highlight
    VertAlign = props.VertAlign |> Option.orElse baseProps.VertAlign
}

let private effectiveOnOff value = value |> Option.defaultValue false

let private mergeParagraphProperties baseProps props = {
    Style = props.Style |> Option.orElse baseProps.Style
    ContentControlId = props.ContentControlId |> Option.orElse baseProps.ContentControlId
    ContentControlTag = props.ContentControlTag |> Option.orElse baseProps.ContentControlTag
    ContentControlAlias = props.ContentControlAlias |> Option.orElse baseProps.ContentControlAlias
    Alignment = props.Alignment |> Option.orElse baseProps.Alignment
    IndentLeft = props.IndentLeft |> Option.orElse baseProps.IndentLeft
    IndentRight = props.IndentRight |> Option.orElse baseProps.IndentRight
    FirstLineIndent = props.FirstLineIndent |> Option.orElse baseProps.FirstLineIndent
    HangingIndent = props.HangingIndent |> Option.orElse baseProps.HangingIndent
    RightTabPosition = props.RightTabPosition |> Option.orElse baseProps.RightTabPosition
    RightTabLeader = props.RightTabLeader |> Option.orElse baseProps.RightTabLeader
    TabStops = (if List.isEmpty props.TabStops then baseProps.TabStops else props.TabStops)
    Rtl = props.Rtl |> Option.orElse baseProps.Rtl
    SpacingBefore = props.SpacingBefore |> Option.orElse baseProps.SpacingBefore
    SpacingBeforeAuto = props.SpacingBeforeAuto |> Option.orElse baseProps.SpacingBeforeAuto
    SpacingAfter = props.SpacingAfter |> Option.orElse baseProps.SpacingAfter
    SpacingAfterAuto = props.SpacingAfterAuto |> Option.orElse baseProps.SpacingAfterAuto
    LineSpacing = props.LineSpacing |> Option.orElse baseProps.LineSpacing
    LineRule = props.LineRule |> Option.orElse baseProps.LineRule
    OutlineLevel = props.OutlineLevel |> Option.orElse baseProps.OutlineLevel
    KeepNext = props.KeepNext |> Option.orElse baseProps.KeepNext
    KeepLines = props.KeepLines |> Option.orElse baseProps.KeepLines
    PageBreakBefore = props.PageBreakBefore |> Option.orElse baseProps.PageBreakBefore
    WidowControl = props.WidowControl |> Option.orElse baseProps.WidowControl
    ContextualSpacing = props.ContextualSpacing |> Option.orElse baseProps.ContextualSpacing
    BackgroundColor = props.BackgroundColor |> Option.orElse baseProps.BackgroundColor
    ParagraphRunProperties =
        match baseProps.ParagraphRunProperties, props.ParagraphRunProperties with
        | Some baseRun, Some run -> Some (mergeRunProperties baseRun run)
        | None, run -> run
        | run, None -> run
    ParagraphBorders = props.ParagraphBorders |> Option.orElse baseProps.ParagraphBorders
    ListInfo = props.ListInfo |> Option.orElse baseProps.ListInfo
    SectionBreak = props.SectionBreak
    SectionColumns = props.SectionColumns
    SectionGeometry = props.SectionGeometry
    SectionPageNumberStart = props.SectionPageNumberStart
    RenderedPageBreakBefore = props.RenderedPageBreakBefore
    HeaderReferences = props.HeaderReferences
    FooterReferences = props.FooterReferences
}

let private paragraphPropertiesWithRunProperties runProps props =
    match runProps, props with
    | None, value -> value
    | Some runProps, Some props ->
        let paragraphRunProperties =
            match props.ParagraphRunProperties with
            | Some existing -> Some (mergeRunProperties runProps existing)
            | None -> Some runProps
        Some { props with ParagraphRunProperties = paragraphRunProperties }
    | Some runProps, None ->
        Some { Docx.paragraphProps XmlDsl.Empty with ParagraphRunProperties = Some runProps }

let private buildStyleMap (styles: StyleDefinition list) : Map<string, StyleDefinition> =
    styles |> List.map (fun style -> (style.Id, style)) |> Map.ofList

let private mergeCellBorders (baseBorders: CellBorders option) (borders: CellBorders option) =
    match baseBorders, borders with
    | None, value -> value
    | value, None -> value
    | Some baseValue, Some value ->
        Some {
            Top = value.Top |> Option.orElse baseValue.Top
            Bottom = value.Bottom |> Option.orElse baseValue.Bottom
            Left = value.Left |> Option.orElse baseValue.Left
            Right = value.Right |> Option.orElse baseValue.Right
            InsideH = value.InsideH |> Option.orElse baseValue.InsideH
            InsideV = value.InsideV |> Option.orElse baseValue.InsideV
        }

let private mergeTableStyleProperties (baseProps: TableStyleProperties) (props: TableStyleProperties) : TableStyleProperties = {
    Borders = mergeCellBorders baseProps.Borders props.Borders
    Shading = props.Shading |> Option.orElse baseProps.Shading
    PaddingTop = props.PaddingTop |> Option.orElse baseProps.PaddingTop
    PaddingRight = props.PaddingRight |> Option.orElse baseProps.PaddingRight
    PaddingBottom = props.PaddingBottom |> Option.orElse baseProps.PaddingBottom
    PaddingLeft = props.PaddingLeft |> Option.orElse baseProps.PaddingLeft
}

let private mergeCellProperties (baseProps: TableCellProperties option) (props: TableCellProperties option) =
    match baseProps, props with
    | None, value -> value
    | value, None -> value
    | Some baseValue, Some value ->
        Some {
            Borders = mergeCellBorders baseValue.Borders value.Borders
            Shading = value.Shading |> Option.orElse baseValue.Shading
            Width = value.Width |> Option.orElse baseValue.Width
            WidthType = value.WidthType |> Option.orElse baseValue.WidthType
            VerticalAlignment = value.VerticalAlignment |> Option.orElse baseValue.VerticalAlignment
            PaddingTop = value.PaddingTop |> Option.orElse baseValue.PaddingTop
            PaddingRight = value.PaddingRight |> Option.orElse baseValue.PaddingRight
            PaddingBottom = value.PaddingBottom |> Option.orElse baseValue.PaddingBottom
            PaddingLeft = value.PaddingLeft |> Option.orElse baseValue.PaddingLeft
            ColSpan = value.ColSpan |> Option.orElse baseValue.ColSpan
            VerticalMerge = value.VerticalMerge |> Option.orElse baseValue.VerticalMerge
        }

let private mergeTableRowStyleProperties (baseProps: TableRowStyleProperties option) (props: TableRowStyleProperties option) =
    match baseProps, props with
    | None, value -> value
    | value, None -> value
    | Some baseValue, Some value ->
        Some {
            IsHeader = value.IsHeader |> Option.orElse baseValue.IsHeader
            CanSplit = value.CanSplit |> Option.orElse baseValue.CanSplit
        }

let private resolveTableStyle (styleMap: Map<string, StyleDefinition>) styleId : StyleDefinition option =
    let rec resolve visited id =
        if Set.contains id visited then None
        else
            match Map.tryFind id styleMap with
            | None -> None
            | Some style ->
                let baseStyle =
                    match style.BasedOn with
                    | Some basedOn when basedOn <> id -> resolve (Set.add id visited) basedOn
                    | _ -> None
                let mergeConditional baseProps props =
                    match baseProps, props with
                    | Some baseValue, Some value -> Some (mergeTableStyleProperties baseValue value)
                    | None, Some value -> Some value
                    | Some value, None -> Some value
                    | None, None -> None
                let mergeConditionalCell baseProps props = mergeCellProperties baseProps props
                let mergeConditionalParagraph baseProps props =
                    match baseProps, props with
                    | Some baseValue, Some value -> Some (mergeParagraphProperties baseValue value)
                    | None, Some value -> Some value
                    | Some value, None -> Some value
                    | None, None -> None
                Some {
                    style with
                        TableProperties =
                            mergeConditional (baseStyle |> Option.bind (fun style -> style.TableProperties)) style.TableProperties
                        FirstRowParagraphProperties = mergeConditionalParagraph (baseStyle |> Option.bind (fun style -> style.FirstRowParagraphProperties)) style.FirstRowParagraphProperties
                        FirstRowTableRowProperties = mergeTableRowStyleProperties (baseStyle |> Option.bind (fun style -> style.FirstRowTableRowProperties)) style.FirstRowTableRowProperties
                        FirstRowTableProperties = mergeConditional (baseStyle |> Option.bind (fun style -> style.FirstRowTableProperties)) style.FirstRowTableProperties
                        LastRowTableProperties = mergeConditional (baseStyle |> Option.bind (fun style -> style.LastRowTableProperties)) style.LastRowTableProperties
                        FirstColumnTableProperties = mergeConditional (baseStyle |> Option.bind (fun style -> style.FirstColumnTableProperties)) style.FirstColumnTableProperties
                        LastColumnTableProperties = mergeConditional (baseStyle |> Option.bind (fun style -> style.LastColumnTableProperties)) style.LastColumnTableProperties
                        Band1HorizontalTableProperties = mergeConditional (baseStyle |> Option.bind (fun style -> style.Band1HorizontalTableProperties)) style.Band1HorizontalTableProperties
                        Band2HorizontalTableProperties = mergeConditional (baseStyle |> Option.bind (fun style -> style.Band2HorizontalTableProperties)) style.Band2HorizontalTableProperties
                        Band1VerticalTableProperties = mergeConditional (baseStyle |> Option.bind (fun style -> style.Band1VerticalTableProperties)) style.Band1VerticalTableProperties
                        Band2VerticalTableProperties = mergeConditional (baseStyle |> Option.bind (fun style -> style.Band2VerticalTableProperties)) style.Band2VerticalTableProperties
                        FirstRowTableCellProperties = mergeConditionalCell (baseStyle |> Option.bind (fun style -> style.FirstRowTableCellProperties)) style.FirstRowTableCellProperties
                        LastRowTableCellProperties = mergeConditionalCell (baseStyle |> Option.bind (fun style -> style.LastRowTableCellProperties)) style.LastRowTableCellProperties
                        FirstColumnTableCellProperties = mergeConditionalCell (baseStyle |> Option.bind (fun style -> style.FirstColumnTableCellProperties)) style.FirstColumnTableCellProperties
                        LastColumnTableCellProperties = mergeConditionalCell (baseStyle |> Option.bind (fun style -> style.LastColumnTableCellProperties)) style.LastColumnTableCellProperties
                        Band1HorizontalTableCellProperties = mergeConditionalCell (baseStyle |> Option.bind (fun style -> style.Band1HorizontalTableCellProperties)) style.Band1HorizontalTableCellProperties
                        Band2HorizontalTableCellProperties = mergeConditionalCell (baseStyle |> Option.bind (fun style -> style.Band2HorizontalTableCellProperties)) style.Band2HorizontalTableCellProperties
                        Band1VerticalTableCellProperties = mergeConditionalCell (baseStyle |> Option.bind (fun style -> style.Band1VerticalTableCellProperties)) style.Band1VerticalTableCellProperties
                        Band2VerticalTableCellProperties = mergeConditionalCell (baseStyle |> Option.bind (fun style -> style.Band2VerticalTableCellProperties)) style.Band2VerticalTableCellProperties
                }

    styleId |> Option.bind (resolve Set.empty)

let private resolveStyle (styleMap: Map<string, StyleDefinition>) styleId : RunProperties * ParagraphProperties option =
    let rec resolve visited id =
        if Set.contains id visited then (emptyRunProperties, (None: ParagraphProperties option))
        else
            match Map.tryFind id styleMap with
            | None -> (emptyRunProperties, (None: ParagraphProperties option))
            | Some style ->
                let baseRun, basePara =
                    match style.BasedOn with
                    | Some basedOn when basedOn <> id -> resolve (Set.add id visited) basedOn
                    | _ -> (emptyRunProperties, (None: ParagraphProperties option))
                let run = style.RunProperties |> Option.map (mergeRunProperties baseRun) |> Option.defaultValue baseRun
                let para =
                    match basePara, style.ParagraphProperties with
                    | Some baseValue, Some value -> Some (mergeParagraphProperties baseValue value)
                    | None, Some value -> Some value
                    | Some value, None -> Some value
                    | None, None -> None
                (run, para)

    styleId |> Option.map (resolve Set.empty) |> Option.defaultValue (emptyRunProperties, None)

let private defaultParagraphStyleId (styleMap: Map<string, StyleDefinition>) =
    styleMap
    |> Map.toSeq
    |> Seq.tryPick (fun (id, style) ->
        if style.StyleType = "paragraph" && style.IsDefault then Some id else None)

let private resolvedParagraphProperties (styleMap: Map<string, StyleDefinition>) (props: ParagraphProperties) =
    let _, docDefaults = resolveStyle styleMap (Some Styles.defaultStyleId)
    let styleId = props.Style |> Option.orElseWith (fun () -> defaultParagraphStyleId styleMap)
    let _, basePara = resolveStyle styleMap styleId
    let basePara =
        match docDefaults, basePara with
        | Some defaults, Some baseValue -> Some (mergeParagraphProperties defaults baseValue)
        | Some defaults, None -> Some defaults
        | None, Some baseValue -> Some baseValue
        | None, None -> None
    basePara |> Option.map (fun baseValue -> mergeParagraphProperties baseValue props) |> Option.defaultValue props

let private effectiveListInfo (props: ParagraphProperties) =
    props.ListInfo |> Option.filter (fun info -> info.ListId <> "0")

let private resolvedRunProperties (styleMap: Map<string, StyleDefinition>) (paraProps: ParagraphProperties) (runProps: RunProperties) =
    let docDefaultRun, _ = resolveStyle styleMap (Some Styles.defaultStyleId)
    let styleId = paraProps.Style |> Option.orElseWith (fun () -> defaultParagraphStyleId styleMap)
    let baseRun, _ = resolveStyle styleMap styleId
    let characterRun, _ = resolveStyle styleMap runProps.Style
    let paragraphRun = paraProps.ParagraphRunProperties |> Option.defaultValue emptyRunProperties
    let merged = mergeRunProperties (mergeRunProperties (mergeRunProperties (mergeRunProperties docDefaultRun baseRun) paragraphRun) characterRun) runProps
    if merged.FontFamily.IsNone && merged.ThemeFont.IsNone then
        { merged with FontFamily = docDefaultRun.FontFamily; ThemeFont = docDefaultRun.ThemeFont }
    else
        merged

let private paragraphMarkFontSize (styleMap: Map<string, StyleDefinition>) (effectiveParaProps: ParagraphProperties) (para: Paragraph) =
    let fallbackRunProps =
        if isPureEmptyParagraph para then
            let docDefaultRun, _ = resolveStyle styleMap (Some Styles.defaultStyleId)
            docDefaultRun
        else
            resolvedRunProperties styleMap effectiveParaProps emptyRunProperties
    wordFontSizeToPoints fallbackRunProps.FontSize

// Word numbers footnotes with arabic (the sequential id) and endnotes with
// lowercase roman numerals by default.
let private toLowerRoman (n: int) =
    let table = [ 1000, "m"; 900, "cm"; 500, "d"; 400, "cd"; 100, "c"; 90, "xc"; 50, "l"; 40, "xl"; 10, "x"; 9, "ix"; 5, "v"; 4, "iv"; 1, "i" ]
    let rec build rem acc =
        match table |> List.tryFind (fun (v, _) -> rem >= v) with
        | Some (v, s) -> build (rem - v) (acc + s)
        | None -> acc
    if n <= 0 then string n else build n ""

let private noteMark (noteType: string) (id: string) =
    if noteType = "endnote" then
        match System.Int32.TryParse id with
        | true, n -> toLowerRoman n
        | _ -> id
    else id

/// What was set at run level, next to the effective values the run object
/// carries: null = inherited from the styles, true/false (or a size) = set.
/// The editor needs the difference to let a run switch a style's formatting
/// off, and the writer to leave inherited formatting out of the file.
///
/// "Run level" is the run's own w:rPr plus the paragraph's w:pPr/w:rPr, which
/// resolvedRunProperties applies to the paragraph's runs (it is also how a table
/// style's run formatting reaches the cell text): neither comes from a named
/// style the editor could resolve again, so both are direct formatting to it.
let private explicitRunFields (styleMap: Map<string, StyleDefinition>) (paraProps: ParagraphProperties) (run: Run) =
    let props =
        match run.Content with
        | Link link -> link.Properties
        | _ -> run.Properties
    let mark = paraProps.ParagraphRunProperties |> Option.defaultValue emptyRunProperties
    let effective = lazy (resolvedRunProperties styleMap paraProps props)
    let onOff (pick: RunProperties -> bool option) =
        match pick props with
        | Some value -> Some value
        | None when (pick mark).IsSome -> Some (effectiveOnOff (pick effective.Value))
        | None -> None
    [
        "HasExplicit" ==> true
        "ExplicitBold" ==> onOff (fun p -> p.Bold)
        "ExplicitItalic" ==> onOff (fun p -> p.Italic)
        "ExplicitUnderline" ==> onOff (fun p -> p.Underline)
        "ExplicitAllCaps" ==> onOff (fun p -> p.AllCaps)
        "ExplicitSmallCaps" ==> onOff (fun p -> p.SmallCaps)
        "ExplicitStrike" ==> onOff (fun p -> p.Strike)
        "ExplicitDoubleStrike" ==> onOff (fun p -> p.DoubleStrike)
        "ExplicitHidden" ==> onOff (fun p -> p.Hidden)
        "ExplicitFontSize" ==>
            (if props.FontSize.IsSome then Some (wordFontSizeToPoints props.FontSize)
             elif mark.FontSize.IsSome then Some (wordFontSizeToPoints effective.Value.FontSize)
             else None)
    ]

let rec private toEditorParagraphsInternal (theme: Theme option) (styles: StyleDefinition list) (content: DocumentContent list) (images: Map<string, string>) (noteBodies: Map<string * string, string>) : obj array =
    let styleMap = buildStyleMap styles
    let _, documentDefaultParagraphProperties = resolveStyle styleMap (Some Styles.defaultStyleId)
    let documentDefaultMissingSpacingToZero =
        documentDefaultParagraphProperties |> Option.exists tableStyleParagraphResetsMissingSpacing
    content
    |> List.choose (fun item ->
        match item with
        | Para para ->
            let effectiveParaProps = resolvedParagraphProperties styleMap para.Properties
            let paragraphFontSize =
                let fallback = paragraphMarkFontSize styleMap effectiveParaProps para
                let mutable maxSize = fallback
                for run in para.Runs do
                    match run.Content with
                    | Text _
                    | Link _
                    | FieldRef _ ->
                        let runProps = resolvedRunProperties styleMap effectiveParaProps run.Properties
                        maxSize <- max maxSize (wordFontSizeToPoints runProps.FontSize)
                    | _ -> ()
                maxSize
            let pageBreakAfter =
                para.Runs
                |> List.exists (fun run -> match run.Content with | PageBreak -> true | _ -> false)
                || effectiveParaProps.SectionBreak = Some "nextPage"
            let bookmarkRunProps (run: Run) = explicitRunFields styleMap effectiveParaProps run @ [
                "BookmarkStartId" ==> (run.BookmarkStart |> Option.map (fun bookmark -> bookmark.Id))
                "BookmarkName" ==> (run.BookmarkStart |> Option.bind (fun bookmark -> bookmark.Name))
                "BookmarkEndId" ==> run.BookmarkEnd
                "RenderedPageBreakBefore" ==> run.RenderedPageBreakBefore
            ]

            // Convert Run to EditorRun format
            let editorRuns =
                para.Runs
                |> List.choose (fun run ->
                    let effectiveRunProps = resolvedRunProperties styleMap effectiveParaProps run.Properties
                    match run.Content with
                    | Text text ->
                        // Text run
                        Some (objFromFields ([
                            "Text" ==> text
                            "Style" ==> run.Properties.Style
                            "Bold" ==> effectiveOnOff effectiveRunProps.Bold
                            "Italic" ==> effectiveOnOff effectiveRunProps.Italic
                            "Underline" ==> effectiveOnOff effectiveRunProps.Underline
                            "UnderlineStyle" ==> effectiveRunProps.UnderlineStyle
                            "UnderlineColor" ==> effectiveRunProps.UnderlineColor
                            "AllCaps" ==> effectiveOnOff effectiveRunProps.AllCaps
                            "SmallCaps" ==> effectiveOnOff effectiveRunProps.SmallCaps
                            "Strike" ==> effectiveOnOff effectiveRunProps.Strike
                            "DoubleStrike" ==> effectiveOnOff effectiveRunProps.DoubleStrike
                            "Hidden" ==> effectiveOnOff effectiveRunProps.Hidden
                            "Position" ==> wordPositionToPoints effectiveRunProps.Position
                            "EmphasisMark" ==> effectiveRunProps.EmphasisMark
                            "FontSize" ==> wordFontSizeToPoints effectiveRunProps.FontSize
                            "CharacterSpacing" ==> (wordCharacterSpacingToCssPixels effectiveRunProps.CharacterSpacing |> Option.defaultValue 0.0)
                            "FontFamily" ==> wordFontFamilyToCss theme effectiveRunProps
                            "Color" ==> wordColorToCss effectiveRunProps.Color
                            "Highlight" ==> effectiveRunProps.Highlight
                            "VertAlign" ==> effectiveRunProps.VertAlign
                            "IsImage" ==> false
                            "RevisionId" ==> run.RevisionId
                            "RevisionType" ==> run.RevisionType
                            "RevisionAuthor" ==> run.RevisionAuthor
                            "RevisionDate" ==> run.RevisionDate
                        ] @ bookmarkRunProps run))
                    | Image img ->
                        // Image run - resolve relationship ID to data URL
                        let imageSrc =
                            images
                            |> Map.tryFind img.RelationshipId
                            |> Option.defaultValue img.RelationshipId  // Fallback to relId if not found
                        Some (objFromFields [
                            "IsImage" ==> true
                            "ImageSrc" ==> imageSrc
                            "ImageLinked" ==> img.Linked
                            "ImageWidth" ==> img.Width
                            "ImageHeight" ==> img.Height
                            "ImageAlt" ==> img.Description
                            "ImageName" ==> img.Name
                            "ImageTitle" ==> img.Title
                            "ImageFloating" ==> img.Floating
                            "ImageWrap" ==> img.Wrap
                            "ImagePositionX" ==> img.PositionX
                            "ImagePositionY" ==> img.PositionY
                            "ImagePositionHRelativeFrom" ==> img.PositionHRelativeFrom
                            "ImagePositionVRelativeFrom" ==> img.PositionVRelativeFrom
                            "ImageBehindDoc" ==> img.BehindDoc
                            "ImageAllowOverlap" ==> img.AllowOverlap
                            "ImageLayoutInCell" ==> img.LayoutInCell
                            "ImageDistanceTop" ==> img.DistanceTop
                            "ImageDistanceBottom" ==> img.DistanceBottom
                            "ImageDistanceLeft" ==> img.DistanceLeft
                            "ImageDistanceRight" ==> img.DistanceRight
                            "ImageRelativeHeight" ==> img.RelativeHeight
                            "ImageLocked" ==> img.Locked
                            "ImageSimplePosition" ==> img.SimplePosition
                            "ImageSimplePositionX" ==> img.SimplePositionX
                            "ImageSimplePositionY" ==> img.SimplePositionY
                            "ImageCropLeft" ==> img.CropLeft
                            "ImageCropTop" ==> img.CropTop
                            "ImageCropRight" ==> img.CropRight
                            "ImageCropBottom" ==> img.CropBottom
                        ])
                    | Link link ->
                        // Hyperlink as text run with link property
                        Some (objFromFields ([
                            "Text" ==> link.Text
                            "Style" ==> link.Properties.Style
                            "Bold" ==> effectiveOnOff effectiveRunProps.Bold
                            "Italic" ==> effectiveOnOff effectiveRunProps.Italic
                            "Underline" ==> effectiveOnOff effectiveRunProps.Underline
                            "UnderlineStyle" ==> effectiveRunProps.UnderlineStyle
                            "UnderlineColor" ==> effectiveRunProps.UnderlineColor
                            "AllCaps" ==> effectiveOnOff effectiveRunProps.AllCaps
                            "SmallCaps" ==> effectiveOnOff effectiveRunProps.SmallCaps
                            "Strike" ==> effectiveOnOff effectiveRunProps.Strike
                            "DoubleStrike" ==> effectiveOnOff effectiveRunProps.DoubleStrike
                            "Hidden" ==> effectiveOnOff effectiveRunProps.Hidden
                            "Position" ==> wordPositionToPoints effectiveRunProps.Position
                            "EmphasisMark" ==> effectiveRunProps.EmphasisMark
                            "FontSize" ==> wordFontSizeToPoints effectiveRunProps.FontSize
                            "CharacterSpacing" ==> (wordCharacterSpacingToCssPixels effectiveRunProps.CharacterSpacing |> Option.defaultValue 0.0)
                            "FontFamily" ==> wordFontFamilyToCss theme effectiveRunProps
                            "Color" ==> wordColorToCss effectiveRunProps.Color
                            "Highlight" ==> effectiveRunProps.Highlight
                            "VertAlign" ==> effectiveRunProps.VertAlign
                            "IsImage" ==> false
                            "Hyperlink" ==> link.Target
                        ] @ bookmarkRunProps run))
                        | CommentReference id ->
                            Some (objFromFields [
                                "Text" ==> (sprintf "[Comment %s]" id)
                                "Bold" ==> false
                                "Italic" ==> false
                                "Underline" ==> false
                                "AllCaps" ==> false
                                "SmallCaps" ==> false
                                "Strike" ==> false
                                "FontSize" ==> 11.0
                                "FontFamily" ==> "Calibri"
                                "Color" ==> Some "#7c3aed"
                                "Highlight" ==> None
                                "IsImage" ==> false
                                "CommentId" ==> Some id
                            ])
                        | NoteRef note ->
                            // Render the reference as its superscript mark number
                            // (Word writes note ids sequentially) in the normal
                            // text colour, matching the native/PDF path.
                            Some (objFromFields [
                                "Text" ==> noteMark note.NoteType note.Id
                                "Bold" ==> false
                                "Italic" ==> false
                                "Underline" ==> false
                                "AllCaps" ==> false
                                "SmallCaps" ==> false
                                "Strike" ==> false
                                "FontSize" ==> 11.0
                                "FontFamily" ==> "Calibri"
                                "Color" ==> None
                                "Highlight" ==> None
                                "VertAlign" ==> Some "superscript"
                                "IsImage" ==> false
                                "NoteId" ==> Some note.Id
                                "NoteType" ==> Some note.NoteType
                                "NoteText" ==> (noteBodies |> Map.tryFind (note.NoteType, note.Id))
                            ])
                        | BookmarkStart bookmark ->
                            Some (objFromFields [
                                "Text" ==> ""
                                "Bold" ==> false
                                "Italic" ==> false
                                "Underline" ==> false
                                "AllCaps" ==> false
                                "SmallCaps" ==> false
                                "Strike" ==> false
                                "FontSize" ==> 11.0
                                "FontFamily" ==> "Calibri"
                                "IsImage" ==> false
                                "BookmarkStartId" ==> bookmark.Id
                                "BookmarkName" ==> bookmark.Name
                            ])
                        | BookmarkEnd id ->
                            Some (objFromFields [
                                "Text" ==> ""
                                "Bold" ==> false
                                "Italic" ==> false
                                "Underline" ==> false
                                "AllCaps" ==> false
                                "SmallCaps" ==> false
                                "Strike" ==> false
                                "FontSize" ==> 11.0
                                "FontFamily" ==> "Calibri"
                                "IsImage" ==> false
                                "BookmarkEndId" ==> id
                            ])
                        | FieldRef field ->
                            Some (objFromFields ([
                                "Text" ==> (field.Result |> Option.defaultValue (sprintf "[Field:%s]" field.Code))
                                "Style" ==> run.Properties.Style
                                "Bold" ==> effectiveOnOff effectiveRunProps.Bold
                                "Italic" ==> effectiveOnOff effectiveRunProps.Italic
                                "Underline" ==> effectiveOnOff effectiveRunProps.Underline
                                "UnderlineStyle" ==> effectiveRunProps.UnderlineStyle
                                "UnderlineColor" ==> effectiveRunProps.UnderlineColor
                                "AllCaps" ==> effectiveOnOff effectiveRunProps.AllCaps
                                "SmallCaps" ==> effectiveOnOff effectiveRunProps.SmallCaps
                                "Strike" ==> effectiveOnOff effectiveRunProps.Strike
                                "DoubleStrike" ==> effectiveOnOff effectiveRunProps.DoubleStrike
                                "Hidden" ==> effectiveOnOff effectiveRunProps.Hidden
                                "Position" ==> wordPositionToPoints effectiveRunProps.Position
                                "EmphasisMark" ==> effectiveRunProps.EmphasisMark
                                "FontSize" ==> wordFontSizeToPoints effectiveRunProps.FontSize
                                "CharacterSpacing" ==> (wordCharacterSpacingToCssPixels effectiveRunProps.CharacterSpacing |> Option.defaultValue 0.0)
                                "FontFamily" ==> wordFontFamilyToCss theme effectiveRunProps
                                "Color" ==> wordColorToCss effectiveRunProps.Color
                                "Highlight" ==> effectiveRunProps.Highlight
                                "VertAlign" ==> effectiveRunProps.VertAlign
                                "IsImage" ==> false
                                "FieldCode" ==> field.Code
                                "FieldResult" ==> field.Result
                                "RevisionId" ==> run.RevisionId
                                "RevisionType" ==> run.RevisionType
                                "RevisionAuthor" ==> run.RevisionAuthor
                                "RevisionDate" ==> run.RevisionDate
                            ] @ bookmarkRunProps run))
                        | EquationXml xml ->
                            Some (objFromFields [
                                "Text" ==> "[Equation]"
                                "Bold" ==> false
                                "Italic" ==> false
                                "Underline" ==> false
                                "AllCaps" ==> false
                                "SmallCaps" ==> false
                                "Strike" ==> false
                                "FontSize" ==> 11.0
                                "FontFamily" ==> "Cambria Math"
                                "IsImage" ==> false
                                "IsEquation" ==> true
                                "EquationXml" ==> xml
                            ])
                        | RenderedPageBreak -> None
                        | PageBreak -> None
                )
                |> List.toArray

            // Create EditorParagraph
            let listStyle, listLevel =
                match effectiveListInfo effectiveParaProps with
                | Some info ->
                    let style =
                        match info.NumberFormat with
                        | Some "bullet" -> "bullet"
                        | _ -> "numbered"
                    (style, info.Level + 1)
                | None -> ("none", 0)

            let itemType, headingLevel =
                match effectiveParaProps.Style with
                | Some style when style.StartsWith("Heading") ->
                    let suffix = style.Substring("Heading".Length)
                    match System.Int32.TryParse(suffix) with
                    | true, level -> ("Heading", max 1 (min 3 level))
                    | false, _ -> ("Heading", 1)
                | _ -> ("Paragraph", 0)

            let layoutParaProps = { effectiveParaProps with ParagraphBorders = para.Properties.ParagraphBorders }

            let defaultMissingSpacingToZero =
                documentDefaultMissingSpacingToZero
                || directParagraphFormattingResetsMissingSpacing para.Properties

            Some (objFromFields ([
                "Type" ==> itemType
                "HeadingLevel" ==> headingLevel
                "Runs" ==> editorRuns
                "Style" ==> effectiveParaProps.Style
                "StyleKeepNext" ==> effectiveParaProps.KeepNext
                "StyleKeepLines" ==> effectiveParaProps.KeepLines
                "StylePageBreakBefore" ==> effectiveParaProps.PageBreakBefore
                "PageBreakBefore" ==> effectiveParaProps.PageBreakBefore
                "StyleWidowControl" ==> effectiveParaProps.WidowControl
                "StyleContextualSpacing" ==> effectiveParaProps.ContextualSpacing
                "Alignment" ==> (effectiveParaProps.Alignment |> Option.defaultValue "left")
                "HasExplicit" ==> true
                "ExplicitAlignment" ==> para.Properties.Alignment
                "ListStyle" ==> listStyle
                "ListLevel" ==> listLevel
                "ListId" ==> (effectiveParaProps.ListInfo |> Option.map (fun info -> info.ListId))
                "ListNumberFormat" ==> (effectiveListInfo effectiveParaProps |> Option.bind (fun info -> info.NumberFormat))
                "ListLevelText" ==> (effectiveListInfo effectiveParaProps |> Option.bind (fun info -> info.LevelText))
                "ListStart" ==> (effectiveListInfo effectiveParaProps |> Option.bind (fun info -> info.Start))
                "PageBreakAfter" ==> pageBreakAfter
                "SectionBreak" ==> effectiveParaProps.SectionBreak
                "SectionPageNumberStart" ==> effectiveParaProps.SectionPageNumberStart
                "HeaderReferences" ==> referenceObjects effectiveParaProps.HeaderReferences
                "FooterReferences" ==> referenceObjects effectiveParaProps.FooterReferences
                "ParagraphMarkFontSize" ==> paragraphFontSize
            ] @ paragraphLayoutProps paragraphFontSize defaultMissingSpacingToZero effectiveParaProps.ParagraphBorders layoutParaProps))
        | Tbl table ->
            // Convert table to editor format
            let convertCellParagraph defaultMissingSpacingToZero (para: Paragraph) : obj =
                let effectiveParaProps = resolvedParagraphProperties styleMap para.Properties
                let pageBreakAfter =
                    para.Runs
                    |> List.exists (fun run -> match run.Content with | PageBreak -> true | _ -> false)
                    || effectiveParaProps.SectionBreak = Some "nextPage"
                let bookmarkRunProps (run: Run) = explicitRunFields styleMap effectiveParaProps run @ [
                    "BookmarkStartId" ==> (run.BookmarkStart |> Option.map (fun bookmark -> bookmark.Id))
                    "BookmarkName" ==> (run.BookmarkStart |> Option.bind (fun bookmark -> bookmark.Name))
                    "BookmarkEndId" ==> run.BookmarkEnd
                    "RenderedPageBreakBefore" ==> run.RenderedPageBreakBefore
                ]

                let editorRuns =
                    para.Runs
                    |> List.choose (fun run ->
                        let effectiveRunProps = resolvedRunProperties styleMap effectiveParaProps run.Properties
                        match run.Content with
                        | Text text ->
                            Some (objFromFields ([
                                "Text" ==> text
                                "Style" ==> run.Properties.Style
                                "Bold" ==> effectiveOnOff effectiveRunProps.Bold
                                "Italic" ==> effectiveOnOff effectiveRunProps.Italic
                                "Underline" ==> effectiveOnOff effectiveRunProps.Underline
                                "UnderlineStyle" ==> effectiveRunProps.UnderlineStyle
                                "UnderlineColor" ==> effectiveRunProps.UnderlineColor
                                "AllCaps" ==> effectiveOnOff effectiveRunProps.AllCaps
                                "SmallCaps" ==> effectiveOnOff effectiveRunProps.SmallCaps
                                "Strike" ==> effectiveOnOff effectiveRunProps.Strike
                                "DoubleStrike" ==> effectiveOnOff effectiveRunProps.DoubleStrike
                                "Hidden" ==> effectiveOnOff effectiveRunProps.Hidden
                                "Position" ==> wordPositionToPoints effectiveRunProps.Position
                                "EmphasisMark" ==> effectiveRunProps.EmphasisMark
                                "FontSize" ==> wordFontSizeToPoints effectiveRunProps.FontSize
                                "FontFamily" ==> wordFontFamilyToCss theme effectiveRunProps
                                "Color" ==> wordColorToCss effectiveRunProps.Color
                                "Highlight" ==> effectiveRunProps.Highlight
                                "VertAlign" ==> effectiveRunProps.VertAlign
                                "IsImage" ==> false
                                "RevisionId" ==> run.RevisionId
                                "RevisionType" ==> run.RevisionType
                                "RevisionAuthor" ==> run.RevisionAuthor
                                "RevisionDate" ==> run.RevisionDate
                            ] @ bookmarkRunProps run))
                        | Image img ->
                            let imageSrc =
                                images
                                |> Map.tryFind img.RelationshipId
                                |> Option.defaultValue img.RelationshipId
                            Some (objFromFields [
                                "IsImage" ==> true
                                "ImageSrc" ==> imageSrc
                                "ImageLinked" ==> img.Linked
                                "ImageWidth" ==> img.Width
                                "ImageHeight" ==> img.Height
                                "ImageAlt" ==> img.Description
                                "ImageName" ==> img.Name
                                "ImageTitle" ==> img.Title
                                "ImageFloating" ==> img.Floating
                                "ImageWrap" ==> img.Wrap
                                "ImagePositionX" ==> img.PositionX
                                "ImagePositionY" ==> img.PositionY
                                "ImagePositionHRelativeFrom" ==> img.PositionHRelativeFrom
                                "ImagePositionVRelativeFrom" ==> img.PositionVRelativeFrom
                                "ImageBehindDoc" ==> img.BehindDoc
                                "ImageAllowOverlap" ==> img.AllowOverlap
                                "ImageLayoutInCell" ==> img.LayoutInCell
                                "ImageDistanceTop" ==> img.DistanceTop
                                "ImageDistanceBottom" ==> img.DistanceBottom
                                "ImageDistanceLeft" ==> img.DistanceLeft
                                "ImageDistanceRight" ==> img.DistanceRight
                                "ImageRelativeHeight" ==> img.RelativeHeight
                                "ImageLocked" ==> img.Locked
                                "ImageSimplePosition" ==> img.SimplePosition
                                "ImageSimplePositionX" ==> img.SimplePositionX
                                "ImageSimplePositionY" ==> img.SimplePositionY
                                "ImageCropLeft" ==> img.CropLeft
                                "ImageCropTop" ==> img.CropTop
                                "ImageCropRight" ==> img.CropRight
                                "ImageCropBottom" ==> img.CropBottom
                            ])
                        | Link link ->
                            let effectiveLinkProps = resolvedRunProperties styleMap effectiveParaProps link.Properties
                            Some (objFromFields ([
                                "Text" ==> link.Text
                                "Style" ==> link.Properties.Style
                                "Bold" ==> effectiveOnOff effectiveLinkProps.Bold
                                "Italic" ==> effectiveOnOff effectiveLinkProps.Italic
                                "Underline" ==> effectiveOnOff effectiveLinkProps.Underline
                                "UnderlineStyle" ==> effectiveLinkProps.UnderlineStyle
                                "UnderlineColor" ==> effectiveLinkProps.UnderlineColor
                                "AllCaps" ==> effectiveOnOff effectiveLinkProps.AllCaps
                                "SmallCaps" ==> effectiveOnOff effectiveLinkProps.SmallCaps
                                "Strike" ==> effectiveOnOff effectiveLinkProps.Strike
                                "DoubleStrike" ==> effectiveOnOff effectiveLinkProps.DoubleStrike
                                "Hidden" ==> effectiveOnOff effectiveLinkProps.Hidden
                                "Position" ==> wordPositionToPoints effectiveLinkProps.Position
                                "EmphasisMark" ==> effectiveLinkProps.EmphasisMark
                                "FontSize" ==> wordFontSizeToPoints effectiveLinkProps.FontSize
                                "CharacterSpacing" ==> (wordCharacterSpacingToCssPixels effectiveLinkProps.CharacterSpacing |> Option.defaultValue 0.0)
                                "FontFamily" ==> wordFontFamilyToCss theme effectiveLinkProps
                                "Color" ==> wordColorToCss effectiveLinkProps.Color
                                "Highlight" ==> effectiveLinkProps.Highlight
                                "VertAlign" ==> effectiveLinkProps.VertAlign
                                "IsImage" ==> false
                                "Hyperlink" ==> link.Target
                            ] @ bookmarkRunProps run))
                        | CommentReference id ->
                            Some (objFromFields [
                                "Text" ==> (sprintf "[Comment %s]" id)
                                "Bold" ==> false
                                "Italic" ==> false
                                "Underline" ==> false
                                "AllCaps" ==> false
                                "SmallCaps" ==> false
                                "Strike" ==> false
                                "FontSize" ==> 11.0
                                "FontFamily" ==> "Calibri"
                                "Color" ==> Some "#7c3aed"
                                "Highlight" ==> None
                                "IsImage" ==> false
                                "CommentId" ==> Some id
                            ])
                        | NoteRef note ->
                            // Render the reference as its superscript mark number
                            // (Word writes note ids sequentially) in the normal
                            // text colour, matching the native/PDF path.
                            Some (objFromFields [
                                "Text" ==> noteMark note.NoteType note.Id
                                "Bold" ==> false
                                "Italic" ==> false
                                "Underline" ==> false
                                "AllCaps" ==> false
                                "SmallCaps" ==> false
                                "Strike" ==> false
                                "FontSize" ==> 11.0
                                "FontFamily" ==> "Calibri"
                                "Color" ==> None
                                "Highlight" ==> None
                                "VertAlign" ==> Some "superscript"
                                "IsImage" ==> false
                                "NoteId" ==> Some note.Id
                                "NoteType" ==> Some note.NoteType
                                "NoteText" ==> (noteBodies |> Map.tryFind (note.NoteType, note.Id))
                            ])
                        | BookmarkStart bookmark ->
                            Some (objFromFields [
                                "Text" ==> ""
                                "Bold" ==> false
                                "Italic" ==> false
                                "Underline" ==> false
                                "AllCaps" ==> false
                                "SmallCaps" ==> false
                                "Strike" ==> false
                                "FontSize" ==> 11.0
                                "FontFamily" ==> "Calibri"
                                "IsImage" ==> false
                                "BookmarkStartId" ==> bookmark.Id
                                "BookmarkName" ==> bookmark.Name
                            ])
                        | BookmarkEnd id ->
                            Some (objFromFields [
                                "Text" ==> ""
                                "Bold" ==> false
                                "Italic" ==> false
                                "Underline" ==> false
                                "AllCaps" ==> false
                                "SmallCaps" ==> false
                                "Strike" ==> false
                                "FontSize" ==> 11.0
                                "FontFamily" ==> "Calibri"
                                "IsImage" ==> false
                                "BookmarkEndId" ==> id
                            ])
                        | FieldRef field ->
                            Some (objFromFields ([
                                "Text" ==> (field.Result |> Option.defaultValue (sprintf "[Field:%s]" field.Code))
                                "Style" ==> run.Properties.Style
                                "Bold" ==> effectiveOnOff effectiveRunProps.Bold
                                "Italic" ==> effectiveOnOff effectiveRunProps.Italic
                                "Underline" ==> effectiveOnOff effectiveRunProps.Underline
                                "UnderlineStyle" ==> effectiveRunProps.UnderlineStyle
                                "UnderlineColor" ==> effectiveRunProps.UnderlineColor
                                "AllCaps" ==> effectiveOnOff effectiveRunProps.AllCaps
                                "SmallCaps" ==> effectiveOnOff effectiveRunProps.SmallCaps
                                "Strike" ==> effectiveOnOff effectiveRunProps.Strike
                                "DoubleStrike" ==> effectiveOnOff effectiveRunProps.DoubleStrike
                                "Hidden" ==> effectiveOnOff effectiveRunProps.Hidden
                                "Position" ==> wordPositionToPoints effectiveRunProps.Position
                                "EmphasisMark" ==> effectiveRunProps.EmphasisMark
                                "FontSize" ==> wordFontSizeToPoints effectiveRunProps.FontSize
                                "FontFamily" ==> wordFontFamilyToCss theme effectiveRunProps
                                "Color" ==> wordColorToCss effectiveRunProps.Color
                                "Highlight" ==> effectiveRunProps.Highlight
                                "VertAlign" ==> effectiveRunProps.VertAlign
                                "IsImage" ==> false
                                "FieldCode" ==> field.Code
                                "FieldResult" ==> field.Result
                                "RevisionId" ==> run.RevisionId
                                "RevisionType" ==> run.RevisionType
                                "RevisionAuthor" ==> run.RevisionAuthor
                                "RevisionDate" ==> run.RevisionDate
                            ] @ bookmarkRunProps run))
                        | EquationXml xml ->
                            Some (objFromFields [
                                "Text" ==> "[Equation]"
                                "Bold" ==> false
                                "Italic" ==> false
                                "Underline" ==> false
                                "AllCaps" ==> false
                                "SmallCaps" ==> false
                                "Strike" ==> false
                                "FontSize" ==> 11.0
                                "FontFamily" ==> "Cambria Math"
                                "IsImage" ==> false
                                "IsEquation" ==> true
                                "EquationXml" ==> xml
                            ])
                        | RenderedPageBreak -> None
                        | PageBreak -> None
                    )
                    |> List.toArray

                let paragraphFontSize =
                    let fallbackRunProps = resolvedRunProperties styleMap effectiveParaProps emptyRunProperties
                    let fallback = wordFontSizeToPoints fallbackRunProps.FontSize
                    para.Runs
                    |> List.choose (fun run ->
                        match run.Content with
                        | Text _
                        | Link _
                        | FieldRef _ ->
                            let runProps = resolvedRunProperties styleMap effectiveParaProps run.Properties
                            Some (wordFontSizeToPoints runProps.FontSize)
                        | _ -> None)
                    |> function
                        | [] -> fallback
                        | sizes -> List.max sizes

                let layoutParaProps = { effectiveParaProps with ParagraphBorders = para.Properties.ParagraphBorders }

                objFromFields ([
                    "Runs" ==> editorRuns
                    "Style" ==> effectiveParaProps.Style
                    "StyleKeepNext" ==> effectiveParaProps.KeepNext
                    "StyleKeepLines" ==> effectiveParaProps.KeepLines
                    "StylePageBreakBefore" ==> effectiveParaProps.PageBreakBefore
                    "PageBreakBefore" ==> effectiveParaProps.PageBreakBefore
                    "StyleWidowControl" ==> effectiveParaProps.WidowControl
                    "StyleContextualSpacing" ==> effectiveParaProps.ContextualSpacing
                    "Alignment" ==> (effectiveParaProps.Alignment |> Option.defaultValue "left")
                    "HasExplicit" ==> true
                    "ExplicitAlignment" ==> para.Properties.Alignment
                    "ListStyle" ==>
                        (match effectiveListInfo effectiveParaProps with
                         | Some info -> match info.NumberFormat with | Some "bullet" -> "bullet" | _ -> "numbered"
                         | None -> "none")
                    "ListLevel" ==>
                        (match effectiveListInfo effectiveParaProps with
                         | Some info -> info.Level + 1
                         | None -> 0)
                    "ListId" ==> (effectiveParaProps.ListInfo |> Option.map (fun info -> info.ListId))
                    "ListNumberFormat" ==> (effectiveListInfo effectiveParaProps |> Option.bind (fun info -> info.NumberFormat))
                    "ListLevelText" ==> (effectiveListInfo effectiveParaProps |> Option.bind (fun info -> info.LevelText))
                    "ListStart" ==> (effectiveListInfo effectiveParaProps |> Option.bind (fun info -> info.Start))
                    "PageBreakAfter" ==> pageBreakAfter
                    "SectionBreak" ==> effectiveParaProps.SectionBreak
                    "SectionPageNumberStart" ==> effectiveParaProps.SectionPageNumberStart
                    "HeaderReferences" ==> referenceObjects effectiveParaProps.HeaderReferences
                    "FooterReferences" ==> referenceObjects effectiveParaProps.FooterReferences
                    "ParagraphMarkFontSize" ==> paragraphFontSize
                ] @ paragraphLayoutProps paragraphFontSize defaultMissingSpacingToZero effectiveParaProps.ParagraphBorders layoutParaProps)

            let tableStyle = resolveTableStyle styleMap table.StyleId
            let tableStyleProperties = tableStyle |> Option.bind (fun style -> style.TableProperties)
            let tableBorders =
                table.Borders
                |> Option.orElseWith (fun () -> table.Properties |> Option.bind (fun props -> props.Borders))
                |> Option.orElseWith (fun () -> tableStyleProperties |> Option.bind (fun props -> props.Borders))
            let rowCount = table.Rows.Length
            let cellSpan (cell: TableCell) = cell.Properties |> Option.bind (fun props -> props.ColSpan) |> Option.defaultValue 1 |> max 1
            let columnCount = table.Rows |> List.map (fun row -> row.Cells |> List.sumBy cellSpan) |> function | [] -> 0 | counts -> List.max counts

            let conditionalFlags rowIndex columnIndex span =
                let look = table.Look
                let enabled selector = look |> Option.map selector |> Option.defaultValue false
                {| IsFirstRow = rowIndex = 0 && enabled (fun look -> look.FirstRow)
                   IsLastRow = rowIndex = rowCount - 1 && enabled (fun look -> look.LastRow)
                   IsFirstColumn = columnIndex = 0 && enabled (fun look -> look.FirstColumn)
                   IsLastColumn = columnIndex + span >= columnCount && enabled (fun look -> look.LastColumn)
                   HorizontalBanding = enabled (fun look -> look.HorizontalBanding)
                   VerticalBanding = enabled (fun look -> look.VerticalBanding) |}

            let conditionalCellStyleProperties rowIndex columnIndex span =
                let flags = conditionalFlags rowIndex columnIndex span
                let horizontalBand =
                    if flags.HorizontalBanding then
                        if rowIndex % 2 = 0 then tableStyle |> Option.bind (fun style -> style.Band1HorizontalTableCellProperties)
                        else tableStyle |> Option.bind (fun style -> style.Band2HorizontalTableCellProperties)
                    else None
                let verticalBand =
                    if flags.VerticalBanding then
                        if columnIndex % 2 = 0 then tableStyle |> Option.bind (fun style -> style.Band1VerticalTableCellProperties)
                        else tableStyle |> Option.bind (fun style -> style.Band2VerticalTableCellProperties)
                    else None
                None
                |> mergeCellProperties horizontalBand
                |> mergeCellProperties verticalBand
                |> mergeCellProperties (if flags.IsFirstColumn then tableStyle |> Option.bind (fun style -> style.FirstColumnTableCellProperties) else None)
                |> mergeCellProperties (if flags.IsLastColumn then tableStyle |> Option.bind (fun style -> style.LastColumnTableCellProperties) else None)
                |> mergeCellProperties (if flags.IsFirstRow then tableStyle |> Option.bind (fun style -> style.FirstRowTableCellProperties) else None)
                |> mergeCellProperties (if flags.IsLastRow then tableStyle |> Option.bind (fun style -> style.LastRowTableCellProperties) else None)

            let applyTableCellStyle rowIndex columnIndex span (cell: TableCell) =
                { cell with Properties = mergeCellProperties (conditionalCellStyleProperties rowIndex columnIndex span) cell.Properties }

            let conditionalRowStyleProperties rowIndex =
                let flags = conditionalFlags rowIndex 0 columnCount
                if flags.IsFirstRow then tableStyle |> Option.bind (fun style -> style.FirstRowTableRowProperties) else None

            let conditionalParagraphStyleProperties rowIndex =
                let flags = conditionalFlags rowIndex 0 columnCount
                if flags.IsFirstRow then tableStyle |> Option.bind (fun style -> style.FirstRowParagraphProperties) else None

            let tableParagraphStyleProperties rowIndex =
                let baseProperties =
                    tableStyle
                    |> Option.bind (fun style -> style.ParagraphProperties)
                    |> paragraphPropertiesWithRunProperties (tableStyle |> Option.bind (fun style -> style.RunProperties))
                let conditionalProperties = conditionalParagraphStyleProperties rowIndex
                match baseProperties, conditionalProperties with
                | Some baseValue, Some conditionalValue -> Some (mergeParagraphProperties baseValue conditionalValue)
                | Some value, None
                | None, Some value -> Some value
                | None, None -> None

            let intOption value =
                value |> Option.bind (fun value -> match System.Int32.TryParse(value) with | true, parsed -> Some parsed | false, _ -> None)

            let convertCell rowIndex (cell: TableCell) : obj =
                let paragraphStyle = tableParagraphStyleProperties rowIndex
                let defaultMissingSpacingToZero =
                    paragraphStyle |> Option.exists tableStyleParagraphResetsMissingSpacing
                let cellParagraphs =
                    cell.Paragraphs
                    |> List.map (fun para ->
                        match paragraphStyle with
                        | Some styleProps ->
                            let properties = mergeParagraphProperties styleProps para.Properties
                            let properties =
                                if para.Properties.Style.IsNone && styleProps.Style.IsNone then
                                    { properties with Style = Some Styles.defaultStyleId }
                                else
                                    properties
                            { para with Properties = properties }
                        | None -> para)
                    |> List.map (convertCellParagraph defaultMissingSpacingToZero)
                    |> List.toArray
                let nestedTables =
                    cell.NestedTables
                    |> List.choose (fun nested ->
                        toEditorParagraphsInternal theme styles [Tbl nested] images noteBodies
                        |> Array.tryHead)
                    |> List.toArray

                let backgroundColor =
                    cell.Properties
                    |> Option.bind (fun props -> props.Shading)
                    |> Option.bind directShadingBackgroundColor

                let shading =
                    cell.Properties
                    |> Option.bind (fun props -> props.Shading)
                    |> Option.map wordShadingToEditor

                let width =
                    cell.Properties
                    |> Option.bind (fun props -> props.Width)
                    |> Option.bind (fun width -> match System.Int32.TryParse(width) with | true, value -> Some value | false, _ -> None)

                let border =
                    cell.Properties
                    |> Option.bind (fun props -> props.Borders)
                    |> Option.bind (fun borders -> borders.Top |> Option.orElse borders.Left |> Option.orElse borders.Bottom |> Option.orElse borders.Right)

                let verticalAlignment =
                    cell.Properties
                    |> Option.bind (fun props -> props.VerticalAlignment)

                objFromFields [
                    "Paragraphs" ==> cellParagraphs
                    "NestedTables" ==> nestedTables
                    "NestedTableAnchors" ==> (cell.NestedTableAnchors |> List.toArray)
                    "ColSpan" ==> (cell.Properties |> Option.bind (fun props -> props.ColSpan) |> Option.defaultValue 1)
                    "BackgroundColor" ==> backgroundColor
                    "Shading" ==> shading
                    "Width" ==> width
                    "WidthType" ==> (cell.Properties |> Option.bind (fun props -> props.WidthType))
                    "PreferredWidth" ==> width
                    "PreferredWidthType" ==> (cell.Properties |> Option.bind (fun props -> props.WidthType))
                    "BorderColor" ==> borderColorToCss border
                    "BorderWidth" ==> borderWidthToPoints border
                    "Borders" ==> (cell.Properties |> Option.bind (fun props -> bordersToEditor props.Borders))
                    "VerticalAlignment" ==> verticalAlignment
                    "VerticalMerge" ==> (cell.Properties |> Option.bind (fun props -> props.VerticalMerge))
                    "PaddingTop" ==> (cell.Properties |> Option.bind (fun props -> props.PaddingTop) |> Option.bind (fun value -> match System.Int32.TryParse(value) with | true, parsed -> Some parsed | false, _ -> None))
                    "PaddingRight" ==> (cell.Properties |> Option.bind (fun props -> props.PaddingRight) |> Option.bind (fun value -> match System.Int32.TryParse(value) with | true, parsed -> Some parsed | false, _ -> None))
                    "PaddingBottom" ==> (cell.Properties |> Option.bind (fun props -> props.PaddingBottom) |> Option.bind (fun value -> match System.Int32.TryParse(value) with | true, parsed -> Some parsed | false, _ -> None))
                    "PaddingLeft" ==> (cell.Properties |> Option.bind (fun props -> props.PaddingLeft) |> Option.bind (fun value -> match System.Int32.TryParse(value) with | true, parsed -> Some parsed | false, _ -> None))
                ]

            let convertRow rowIndex (row: TableRow) : obj =
                let rowStyle = conditionalRowStyleProperties rowIndex
                let mutable columnIndex = 0
                let cells =
                    row.Cells
                    |> List.map (fun cell ->
                        let span = cellSpan cell
                        let styledCell = applyTableCellStyle rowIndex columnIndex span cell
                        columnIndex <- columnIndex + span
                        convertCell rowIndex styledCell)
                    |> List.toArray

                objFromFields [
                    "Cells" ==> cells
                    "IsHeader" ==> (row.IsHeader || (rowStyle |> Option.bind (fun props -> props.IsHeader) |> Option.defaultValue false))
                    "Height" ==> row.Height
                    "HeightRule" ==> row.HeightRule
                    "CanSplit" ==> (row.CanSplit |> Option.orElseWith (fun () -> rowStyle |> Option.bind (fun props -> props.CanSplit)))
                    "CellSpacing" ==> intOption row.CellSpacing
                    "PaddingTop" ==> intOption row.PaddingTop
                    "PaddingRight" ==> intOption row.PaddingRight
                    "PaddingBottom" ==> intOption row.PaddingBottom
                    "PaddingLeft" ==> intOption row.PaddingLeft
                ]

            let rows =
                table.Rows
                |> List.mapi convertRow
                |> List.toArray

            let tableBackgroundColor =
                table.Properties
                |> Option.bind (fun props -> props.Shading)
                |> Option.orElseWith (fun () -> tableStyleProperties |> Option.bind (fun props -> props.Shading))
                |> Option.bind directShadingBackgroundColor

            let tableShading =
                table.Properties
                |> Option.bind (fun props -> props.Shading)
                |> Option.orElseWith (fun () -> tableStyleProperties |> Option.bind (fun props -> props.Shading))
                |> Option.map wordShadingToEditor

            Some (objFromFields [
                "Type" ==> "Table"
                "Rows" ==> rows
                "PreferredWidth" ==> table.Width
                "PreferredWidthType" ==> table.WidthType
                "TableAlignment" ==> table.Alignment
                "BackgroundColor" ==> tableBackgroundColor
                "Shading" ==> tableShading
                "CellSpacing" ==> intOption table.CellSpacing
                "PaddingTop" ==> intOption (table.Properties |> Option.bind (fun props -> props.PaddingTop) |> Option.orElseWith (fun () -> tableStyle |> Option.bind (fun style -> style.TableProperties) |> Option.bind (fun props -> props.PaddingTop)))
                "PaddingRight" ==> intOption (table.Properties |> Option.bind (fun props -> props.PaddingRight) |> Option.orElseWith (fun () -> tableStyle |> Option.bind (fun style -> style.TableProperties) |> Option.bind (fun props -> props.PaddingRight)))
                "PaddingBottom" ==> intOption (table.Properties |> Option.bind (fun props -> props.PaddingBottom) |> Option.orElseWith (fun () -> tableStyle |> Option.bind (fun style -> style.TableProperties) |> Option.bind (fun props -> props.PaddingBottom)))
                "PaddingLeft" ==> intOption (table.Properties |> Option.bind (fun props -> props.PaddingLeft) |> Option.orElseWith (fun () -> tableStyle |> Option.bind (fun style -> style.TableProperties) |> Option.bind (fun props -> props.PaddingLeft)))
                "LayoutMode" ==> table.LayoutMode
                "GridColumns" ==> (table.GridColumns |> List.toArray)
                "Floating" ==> table.Floating
                "Borders" ==> bordersToEditor tableBorders
            ])
    )
    |> List.toArray

let private applyNumberingDefinitions (numberingDefinitions: NumberingDefinition list) (items: obj array) =
    let isMissing (value: obj) = isNull value || isNull (box value)
    let optionalStringValue (value: obj) =
        if isMissing value then None
        else
            match value with
            | :? (string option) as value -> value |> Option.bind (function null | "" -> None | text -> Some text)
            | :? string as value ->
                match value with
                | null | "" -> None
                | text -> Some text
            | _ ->
                match unbox<string> value with
                | null | "" -> None
                | text -> Some text
    let optionalIntValue (value: obj) =
        if isMissing value then None
        else
            match value with
            | :? (int option) as value -> value
            | :? int as value -> Some value
            | _ -> Some (unbox<int> value)
    let intValue fallback (value: obj) =
        optionalIntValue value |> Option.defaultValue fallback
    let stringValue fallback (value: obj) = optionalStringValue value |> Option.defaultValue fallback
    let getProp (item: obj) name : obj =
#if FABLE_COMPILER
        item?(name)
#else
        match item with
        | :? System.Collections.Generic.IDictionary<string, obj> as fields ->
            match fields.TryGetValue name with
            | true, value -> value
            | false, _ -> null
        | _ -> null
#endif
    let setProp (item: obj) name value =
#if FABLE_COMPILER
        item?(name) <- value
#else
        match item with
        | :? System.Collections.Generic.IDictionary<string, obj> as fields -> fields.[name] <- box value
        | _ -> ()
#endif

    let numberingMap =
        numberingDefinitions
        |> List.map (fun def -> ((def.NumId, def.Level), def))
        |> Map.ofList
    let styleNumberingMap =
        numberingDefinitions
        |> List.choose (fun def -> def.ParagraphStyle |> Option.map (fun styleId -> (styleId, def)))
        |> Map.ofList

    let applyToParagraph (item: obj) =
        let listId = optionalStringValue (getProp item "ListId")
        let listLevel = intValue 0 (getProp item "ListLevel")
        let styleId = optionalStringValue (getProp item "Style")
        let listDefinition =
            match listId with
            | Some id -> Map.tryFind (id, max 0 (listLevel - 1)) numberingMap
            | None -> styleId |> Option.bind (fun id -> Map.tryFind id styleNumberingMap)
        match listDefinition with
        | Some def when def.NumberFormat = "bullet" ->
            setProp item "ListId" def.NumId
            setProp item "ListLevel" (def.Level + 1)
            setProp item "ListStyle" "bullet"
            setProp item "ListNumberFormat" def.NumberFormat
            setProp item "ListLevelText" def.LevelText
            setProp item "ListStart" def.Start
        | Some def ->
            setProp item "ListId" def.NumId
            setProp item "ListLevel" (def.Level + 1)
            setProp item "ListStyle" "numbered"
            setProp item "ListNumberFormat" def.NumberFormat
            setProp item "ListLevelText" def.LevelText
            setProp item "ListStart" def.Start
        | None -> ()
        match listDefinition with
        | Some def ->
            if isMissing (getProp item "IndentLeft") then setProp item "IndentLeft" (wordTwipsToPoints def.IndentLeft)
            if isMissing (getProp item "FirstLineIndent") then
                let firstLineIndent =
                    match wordTwipsToPoints def.FirstLineIndent, wordTwipsToPoints def.HangingIndent with
                    | Some firstLine, _ -> Some firstLine
                    | None, Some hanging -> Some -hanging
                    | None, None -> None
                setProp item "FirstLineIndent" firstLineIndent
            if isMissing (getProp item "RightTabPosition") then setProp item "RightTabPosition" (wordTwipsToPoints def.RightTabPosition)
        | None -> ()
        item

    let rec apply item =
        let itemType = getProp item "Type" |> stringValue "Paragraph"
        if itemType = "Table" then
            let rows = getProp item "Rows" |> unbox<obj array>
            for row in rows do
                let cells = getProp row "Cells" |> unbox<obj array>
                for cell in cells do
                    let paragraphs = getProp cell "Paragraphs" |> unbox<obj array>
                    for para in paragraphs do
                        applyToParagraph para |> ignore
            item
        else
            applyToParagraph item

    items |> Array.map apply

let toEditorParagraphs (content: DocumentContent list) (images: Map<string, string>) : obj array =
    toEditorParagraphsInternal None [] content images Map.empty

let toEditorParagraphsWithStyles (content: DocumentContent list) (images: Map<string, string>) (styles: StyleDefinition list) : obj array =
    toEditorParagraphsInternal None styles content images Map.empty

let toEditorParagraphsWithStylesAndTheme (theme: Theme option) (content: DocumentContent list) (images: Map<string, string>) (styles: StyleDefinition list) : obj array =
    toEditorParagraphsInternal theme styles content images Map.empty

let toEditorParagraphsWithStylesAndNumbering (content: DocumentContent list) (images: Map<string, string>) (styles: StyleDefinition list) (numberingDefinitions: NumberingDefinition list) : obj array =
    toEditorParagraphsInternal None styles content images Map.empty
    |> applyNumberingDefinitions numberingDefinitions

let private noteBodyMap (footnotes: Note list) (endnotes: Note list) : Map<string * string, string> =
    let bodyText (note: Note) =
        note.Content
        |> List.collect (fun para -> para.Runs)
        |> List.map (fun run ->
            match run.Content with
            | Text text -> text
            | Link link -> link.Text
            | _ -> "")
        |> String.concat ""
    (footnotes |> List.map (fun note -> ("footnote", note.Id), bodyText note))
    @ (endnotes |> List.map (fun note -> ("endnote", note.Id), bodyText note))
    |> Map.ofList

let toEditorParagraphsWithStylesThemeAndNumbering (theme: Theme option) (content: DocumentContent list) (images: Map<string, string>) (styles: StyleDefinition list) (numberingDefinitions: NumberingDefinition list) (footnotes: Note list) (endnotes: Note list) : obj array =
    toEditorParagraphsInternal theme styles content images (noteBodyMap footnotes endnotes)
    |> applyNumberingDefinitions numberingDefinitions

let toEditorHeaderFooters (items: HeaderFooter list) (images: Map<string, string>) : obj array =
    items
    |> List.map (fun item ->
        objFromFields [
            "Type" ==> item.Type
            "RelationshipId" ==> item.RelationshipId
            "Paragraphs" ==> toEditorParagraphs item.Blocks images
        ])
    |> List.toArray

let toEditorHeaderFootersWithStyles (items: HeaderFooter list) (images: Map<string, string>) (styles: StyleDefinition list) : obj array =
    items
    |> List.map (fun item ->
        objFromFields [
            "Type" ==> item.Type
            "RelationshipId" ==> item.RelationshipId
            "Paragraphs" ==> toEditorParagraphsWithStyles item.Blocks images styles
        ])
    |> List.toArray

let toEditorHeaderFootersWithStylesAndTheme (theme: Theme option) (items: HeaderFooter list) (images: Map<string, string>) (styles: StyleDefinition list) : obj array =
    items
    |> List.map (fun item ->
        objFromFields [
            "Type" ==> item.Type
            "RelationshipId" ==> item.RelationshipId
            "Paragraphs" ==> toEditorParagraphsWithStylesAndTheme theme item.Blocks images styles
        ])
    |> List.toArray

let toEditorPageSetup (doc: WordDocument) : obj =
    match doc.PageSetup with
    | None -> objFromFields []
    | Some setup ->
        objFromFields [
            "PageWidth" ==> setup.PageWidth
            "PageHeight" ==> setup.PageHeight
            "MarginLeft" ==> setup.MarginLeft
            "MarginRight" ==> setup.MarginRight
            "MarginTop" ==> setup.MarginTop
            "MarginBottom" ==> setup.MarginBottom
            "HeaderDistance" ==> setup.HeaderDistance
            "FooterDistance" ==> setup.FooterDistance
            "DifferentFirstPage" ==> setup.DifferentFirstPage
            "EvenAndOddHeaders" ==> setup.EvenAndOddHeaders
            "ColumnCount" ==> (setup.Columns |> Option.map (fun c -> c.Count) |> Option.defaultValue 1)
            "ColumnSpace" ==> (setup.Columns |> Option.map (fun c -> c.Gap) |> Option.defaultValue 0.0)
            "ColumnSeparator" ==> (setup.Columns |> Option.map (fun c -> c.Separator) |> Option.defaultValue false)
            "ColumnWidths" ==> (setup.Columns |> Option.map (fun c -> c.Widths |> List.map (fun (w, s) -> objFromFields [ "W" ==> w; "S" ==> s ]) |> Array.ofList) |> Option.defaultValue [||])
            "SectionStart" ==> setup.SectionStart
        ]
