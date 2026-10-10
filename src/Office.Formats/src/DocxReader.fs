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
    /// wp:positionH / wp:positionV given as wp:align (left, center, right, top, ...)
    /// instead of an offset.
    PositionHAlign: string option
    PositionVAlign: string option
    /// wrapText of wp:wrapSquare / wrapTight / wrapThrough (bothSides, left, right, largest).
    WrapText: string option
    /// wp:wrapPolygon of wp:wrapTight / wp:wrapThrough: wp:start and the wp:lineTo points (x, y).
    WrapPolygon: (float * float) list
    /// Set for an object the reader does not model (chart, SmartArt, OLE object,
    /// grouped / VML drawing, unknown mc:AlternateContent): the element as it is
    /// in the file, to be written back unchanged. The other fields then describe
    /// its frame; RelationshipId is that of a fallback picture, or "".
    Opaque: OpaqueObject option
}

/// An element kept verbatim. Xml is self-contained: the namespace declarations
/// it relies on are on its root element.
and OpaqueObject = {
    Xml: string
    /// chart | diagram | ole | shape | drawing
    Kind: string
    /// A child of w:p (it holds its own runs) rather than of w:r.
    ParagraphLevel: bool
    /// The body's shape import models this one as a text box / shape.
    ShapeCandidate: bool
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

/// Id, author and date of a tracked change.
type RevisionMark = {
    Id: string option
    Author: string option
    Date: string option
}

type Run = {
    Content: RunContent
    Properties: RunProperties
    BookmarkStart: BookmarkReference option
    BookmarkEnd: string option
    RenderedPageBreakBefore: bool
    RevisionId: string option  // For track changes
    /// insert | delete | moveFrom | moveTo
    RevisionType: string option
    RevisionAuthor: string option
    RevisionDate: string option
    /// A deletion inside a tracked insertion (w:del in w:ins): the insertion.
    /// RevisionType is "delete" and the fields above describe the deletion.
    InsertedRevision: RevisionMark option
    /// Name of the move range (w:moveFromRangeStart / w:moveToRangeStart) a
    /// moveFrom / moveTo run lies in.
    MoveName: string option
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
    /// PageBreakBefore is not what the paragraph's own properties say: the
    /// reader set it for a Heading 1 that started a page when the document was
    /// last laid out. It is not written back as a page break of the paragraph.
    PageBreakBeforeDerived: bool
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
    SectionPageNumberFormat: string option
    RenderedPageBreakBefore: int option
    HeaderReferences: HeaderFooterReference list
    FooterReferences: HeaderFooterReference list
}

type Paragraph = {
    Runs: Run list
    Properties: ParagraphProperties
    /// The paragraph mark is a tracked change (w:ins / w:del in w:pPr/w:rPr):
    /// ("insert" | "delete", id / author / date).
    MarkRevision: (string * RevisionMark) option
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
    /// The w:tblpPr of a floating table as it is in the file (where it floats
    /// and how far text stays away): written back as it was.
    FloatingXml: string option
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
    /// w:initials of the comment author.
    Initials: string option
    /// The comment body in order, with tables kept as tables and run formatting.
    Blocks: DocumentContent list
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
    /// Every paragraph of the note, tables flattened (kept for callers that only want text).
    Content: Paragraph list
    /// The note body in order, with tables kept as tables.
    Blocks: DocumentContent list
    /// w:type: None for an ordinary note, else "separator", "continuationSeparator", "continuationNotice".
    Kind: string option
    /// The body starts with the automatic reference mark (w:footnoteRef / w:endnoteRef).
    HasMark: bool
    /// Character style of the run that holds the mark.
    MarkStyle: string option
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
    /// Index of the body block (paragraph or table) the shape is anchored in.
    AnchorBlock: int option
    /// wp:inline (in the text line) rather than wp:anchor.
    Inline: bool
    /// none | square | tight | through | topAndBottom.
    Wrap: string option
    PositionHRelativeFrom: string option
    PositionVRelativeFrom: string option
    PositionHAlign: string option
    PositionVAlign: string option
    BehindDoc: bool option
    AllowOverlap: bool option
    LayoutInCell: bool option
    RelativeHeight: string option
    DistanceTop: float option
    DistanceBottom: float option
    DistanceLeft: float option
    DistanceRight: float option
    /// a:noFill / a:ln with a:noFill (VML filled="f" / stroked="f").
    NoFill: bool
    NoLine: bool
    /// a:prstDash@val (dash, dot, dashDot, sysDash, ...).
    LineDash: string option
    /// wp:docPr@name.
    Name: string option
    /// The room between the shape's edge and its text, in pixels (wps:bodyPr
    /// lIns / tIns / rIns / bIns; the inset of a v:textbox).
    InsetLeft: float option
    InsetTop: float option
    InsetRight: float option
    InsetBottom: float option
    /// wps:bodyPr@anchor (t | ctr | b); v-text-anchor of a VML shape.
    VerticalAnchor: string option
    /// a:spAutoFit (mso-fit-shape-to-text of a v:textbox): the shape is as high as its text.
    AutoFit: bool
    /// wps:bodyPr@wrap = "none": the text is not broken at the shape's width.
    NoTextWrap: bool
    /// The text is a v:textpath (WordArt, a watermark): drawn as large as the shape.
    FitText: bool
    /// The opacity of the fill (v:fill@opacity), when it is not opaque.
    Opacity: float option
    /// The element the shape was read from as it stood in its run
    /// (mc:AlternateContent, w:drawing, w:pict), with the namespaces it needs.
    SourceXml: string option
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
    /// Header / footer references and page number start of the body-level
    /// sectPr, i.e. of the final section.
    HeaderReferences: HeaderFooterReference list
    FooterReferences: HeaderFooterReference list
    PageNumberStart: int option
    PageNumberFormat: string option
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

/// A relationship kept from the opened package. An internal Target is the
/// absolute part name ("/word/charts/chart1.xml"), so it can be written into
/// the relationships of any part.
type PreservedRelationship = {
    Id: string
    Type: string
    Target: string
    External: bool
}

type PreservedPart = {
    /// Package path without a leading slash.
    Path: string
    Data: byte[]
}

/// What the opened package holds beyond the parts the editor regenerates
/// (document, headers / footers, notes, comments, numbering): kept and written
/// back on save. See DocumentWriter.writeFromEditorPackage.
type PackagePreservation = {
    /// Parts reachable from the kept relationships, with their own .rels parts.
    Parts: PreservedPart list
    /// Package-level relationships other than the main document's.
    RootRelationships: PreservedRelationship list
    /// Id of the package relationship that points at the main document.
    DocumentRelationshipId: string
    /// Kept relationships of each regenerated part. Keys: "document", the
    /// relationship id of a header / footer, "footnotes", "endnotes", "comments".
    PartRelationships: Map<string, PreservedRelationship list>
    /// [Content_Types].xml: Default (extension, type) and Override (part name, type) entries.
    Defaults: (string * string) list
    Overrides: (string * string) list
    /// Content type of the main document part (a template or macro-enabled document has its own).
    MainContentType: string option
    /// The original styles and settings parts: the writer merges into them.
    StylesXml: string option
    SettingsXml: string option
    /// The parts the writer regenerates, as they were read: "document", the
    /// relationship id of a header / footer, "footnotes", "endnotes",
    /// "comments", "numbering". What was not changed in the editor is written
    /// back from here as it was (a body block, a whole header, the notes).
    SourceXml: Map<string, string>
    /// Every relationship of those parts that the writer does not produce
    /// itself (pictures and hyperlinks included): what a part or a body block
    /// written back as it was refers to. Keys as PartRelationships.
    SourceRelationships: Map<string, PreservedRelationship list>
    /// Parts reachable from SourceRelationships only (the pictures the
    /// editor holds as data): written when something is written back as it was.
    SourceParts: PreservedPart list
    /// For each block of the body (WordDocument.Content), the index of the
    /// child element of w:body it was read from.
    BodyContentChunks: int list
    /// How many child elements w:body has.
    BodyChunkCount: int
    /// The package path of each header / footer part, by its relationship id:
    /// a part is written under the name it had.
    PartPaths: Map<string, string>
    /// [Content_Types].xml and the relationship parts of the package and of
    /// the parts the writer produces, by package path, as they were read: one
    /// that says what the written one would say is written back as it was.
    OriginalText: Map<string, string>
}

let emptyPackagePreservation = {
    SourceXml = Map.empty
    SourceRelationships = Map.empty
    SourceParts = []
    BodyContentChunks = []
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
    /// The shapes of the headers and footers, by the part's relationship id
    /// (Shape.AnchorBlock is an index among the part's blocks).
    PartShapes: Map<string, Shape list>
    PageSetup: PageSetup option
    NumberingDefinitions: NumberingDefinition list
    Preserved: PackagePreservation
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
    let headerFooterReferences name props =
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
                PageBreakBeforeDerived = false
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
                SectionPageNumberFormat = sectPr |> Option.bind (fun section -> getPropVal "pgNumType" "fmt" section)
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
                PageBreakBeforeDerived = false
                WidowControl = None
                ContextualSpacing = None
                BackgroundColor = None
                ParagraphRunProperties = None
                ParagraphBorders = None
                ListInfo = None
                SectionBreak = None
                SectionColumns = None
                SectionGeometry = None
                SectionPageNumberStart = None; SectionPageNumberFormat = None
                RenderedPageBreakBefore = None
                HeaderReferences = []
                FooterReferences = []
            }

    let private emuToPixels (emu: float) = emu / 9525.0

    /// The frame of a w:drawing: size, name / description and, for an anchored
    /// one, wrap and position. `rId` is the relationship of the picture shown.
    let drawingImageData (drawing: XmlNode) (rId: string) (linked: bool) : ImageData =
        let extent = Ns.descendantsNs "extent" drawing |> List.tryHead
        let width = extent |> Option.bind (Xml.attr "cx") |> Option.map (fun cx -> emuToPixels (float cx))
        let height = extent |> Option.bind (Xml.attr "cy") |> Option.map (fun cy -> emuToPixels (float cy))
        let docPr = Ns.descendantsNs "docPr" drawing |> List.tryHead
        let name = docPr |> Option.bind (Xml.attr "name")
        let title = docPr |> Option.bind (Xml.attr "title")
        let desc = docPr |> Option.bind (Xml.attr "descr")
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
        let positionAlign localName =
            anchor
            |> Option.bind (Ns.childNs localName)
            |> Option.bind (Ns.childNs "align")
            |> Option.bind Xml.text
        let wrapText =
            anchor
            |> Option.bind (fun a ->
                Xml.children a
                |> List.tryPick (fun child ->
                    match Xml.name child |> Option.map Ns.stripNs with
                    | Some "wrapSquare" | Some "wrapTight" | Some "wrapThrough" -> Xml.attr "wrapText" child
                    | _ -> None))
        let wrapPolygon =
            anchor
            |> Option.bind (fun a -> Ns.childNs "wrapTight" a |> Option.orElseWith (fun () -> Ns.childNs "wrapThrough" a))
            |> Option.bind (Ns.childNs "wrapPolygon")
            |> Option.map (fun polygon ->
                Xml.children polygon
                |> List.choose (fun point ->
                    match Xml.name point |> Option.map Ns.stripNs with
                    | Some "start" | Some "lineTo" ->
                        match Xml.attr "x" point, Xml.attr "y" point with
                        | Some x, Some y ->
                            match System.Double.TryParse(x: string), System.Double.TryParse(y: string) with
                            | (true, px), (true, py) -> Some (px, py)
                            | _ -> None
                        | _ -> None
                    | _ -> None))
            |> Option.defaultValue []
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
            Linked = linked
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
            PositionHAlign = positionAlign "positionH"
            PositionVAlign = positionAlign "positionV"
            WrapText = wrapText
            WrapPolygon = wrapPolygon
            Opaque = None
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
            PositionHAlign = None
            PositionVAlign = None
            WrapText = None
            WrapPolygon = []
            Opaque = None
        }
        let drawingImage =
            wElem "drawing" rNode
            |> Option.bind (fun drawing ->
                match Ns.descendantsNs "blip" drawing |> List.tryHead with
                | Some b ->
                    let embedId = Xml.attr "r:embed" b |> Option.orElseWith (fun () -> attrByLocalName "embed" b)
                    let linkId = Xml.attr "r:link" b |> Option.orElseWith (fun () -> attrByLocalName "link" b)
                    embedId |> Option.orElse linkId
                    |> Option.map (fun rId -> drawingImageData drawing rId (embedId.IsNone && linkId.IsSome))
                // A drawing without a picture (chart, SmartArt, shape) is not an image;
                // it is kept as an opaque object (see opaqueObject).
                | None -> None)
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

    // ---- drawings the reader does not model ---------------------------------

    // A plain picture also has a:prstGeom (inside pic:spPr). It is read as an
    // image run; without this guard every picture was imported a second time
    // as a rectangle shape and saved as a text box.
    let private isPictureDrawing node =
        (Ns.descendantsNs "pic" node |> List.isEmpty |> not)
        && (Ns.descendantsNs "wsp" node |> List.isEmpty)

    // A shape is imported when it carries text (txbxContent) or a real geometry
    // (prstGeom / custGeom), so autoshapes without text still render.
    let private hasShapeContent node =
        (Ns.descendantsNs "txbxContent" node |> List.isEmpty |> not)
        // (WordArt, a watermark: a VML shape whose text is a v:textpath.)
        || (Ns.descendantsNs "textpath" node |> List.exists (fun path -> (Xml.attr "string" path).IsSome))
        || (not (isPictureDrawing node)
            && ((Ns.descendantsNs "prstGeom" node |> List.isEmpty |> not)
                || (Ns.descendantsNs "custGeom" node |> List.isEmpty |> not)))

    /// A drawing the shape model can hold: one shape with text and / or
    /// geometry. Groups, drawing canvases, charts and diagrams are not.
    let isModelledShape (node: XmlNode) =
        hasShapeContent node
        && (Ns.descendantsNs "wgp" node).IsEmpty
        && (Ns.descendantsNs "wpc" node).IsEmpty
        && (Ns.descendantsNs "group" node).IsEmpty
        && (Ns.descendantsNs "chart" node).IsEmpty
        && (Ns.descendantsNs "relIds" node).IsEmpty

    /// The shape roots under `node`, in document order. For DrawingML that is
    /// the wp:anchor / wp:inline wrapper — position (wp:positionH/V) and display
    /// size (wp:extent) live there, the geometry and fill in the nested wsp.
    /// VML shapes carry their attributes on the v:shape / v:rect element itself.
    ///
    /// An mc:AlternateContent holds the same shape twice (DrawingML choice, VML
    /// fallback); only one of them is taken. Scanning the document for anchors
    /// and v:shapes separately imported such a text box two times.
    let rec modelledShapeRoots (node: XmlNode) : XmlNode list =
        match Xml.name node |> Option.map Ns.stripNs with
        | Some "AlternateContent" ->
            let fromChoice = Ns.childNs "Choice" node |> Option.map modelledShapeRoots |> Option.defaultValue []
            if not fromChoice.IsEmpty then fromChoice
            else Ns.childNs "Fallback" node |> Option.map modelledShapeRoots |> Option.defaultValue []
        | Some "anchor" | Some "inline" -> if isModelledShape node then [node] else []
        | Some "shape" | Some "rect" | Some "roundrect" | Some "oval" when (match node with XmlDsl.Element(name, _, _) -> name.StartsWith "v:" | _ -> false) ->
            if isModelledShape node then [node] else []
        | Some _ -> Xml.children node |> List.collect modelledShapeRoots
        | None -> []

    let private hasGraphics (node: XmlNode) =
        not (Ns.descendantsNs "drawing" node).IsEmpty
        || not (Ns.descendantsNs "pict" node).IsEmpty
        || not (Ns.descendantsNs "object" node).IsEmpty

    /// An mc:AlternateContent whose choice is a drawing: kept whole, so the
    /// choice is not lost to its fallback.
    let private isGraphicAlternateContent (node: XmlNode) =
        Ns.childNs "Choice" node |> Option.exists hasGraphics

    let private blankImageData (relationshipId: string) : ImageData =
        { RelationshipId = relationshipId; Linked = false; Name = None; Title = None; Width = None; Height = None; Description = None
          Floating = false; Wrap = None; PositionX = None; PositionY = None; PositionHRelativeFrom = None; PositionVRelativeFrom = None
          BehindDoc = None; AllowOverlap = None; LayoutInCell = None; DistanceTop = None; DistanceBottom = None; DistanceLeft = None
          DistanceRight = None; RelativeHeight = None; Locked = None; SimplePosition = None; SimplePositionX = None; SimplePositionY = None
          CropLeft = None; CropTop = None; CropRight = None; CropBottom = None; PositionHAlign = None; PositionVAlign = None; WrapText = None; WrapPolygon = []
          Opaque = None }

    /// An object the reader does not model, as an image-like run: its frame
    /// (size, floating position), the fallback picture when the document has
    /// one, and the element itself to write back. Charts and SmartArt used to
    /// be replaced by the text "Chart", which destroyed them on save.
    let opaqueObject (paragraphLevel: bool) (node: XmlNode) : ImageData =
        let nodeName = Xml.name node |> Option.map Ns.stripNs |> Option.defaultValue ""
        let fallbackBranch =
            match nodeName with
            | "AlternateContent" -> Ns.childNs "Fallback" node
            | "object" | "pict" -> Some node
            | _ -> None
        let fallbackRelationship =
            fallbackBranch
            |> Option.bind (fun branch ->
                Ns.descendantsNs "imagedata" branch
                |> List.tryHead
                |> Option.bind (fun image -> Xml.attr "r:id" image |> Option.orElseWith (fun () -> attrByLocalName "id" image))
                |> Option.orElseWith (fun () ->
                    Ns.descendantsNs "blip" branch |> List.tryHead |> Option.bind (fun blip -> Xml.attr "r:embed" blip |> Option.orElseWith (fun () -> attrByLocalName "embed" blip))))
            |> Option.defaultValue ""
        let frame =
            match Ns.descendantsNs "drawing" node |> List.tryHead with
            | Some drawing -> { drawingImageData drawing fallbackRelationship false with RelationshipId = fallbackRelationship }
            | None ->
                // VML: the size and position are in the shape's style attribute.
                let style =
                    Ns.descendantsNs "shape" node @ Ns.descendantsNs "rect" node @ Ns.descendantsNs "group" node
                    |> List.tryPick (Xml.attr "style")
                    |> Option.defaultValue ""
                let styleNumber (name: string) =
                    style.Split(';')
                    |> Array.tryPick (fun part ->
                        let pieces = part.Split([| ':' |], 2)
                        if pieces.Length = 2 && pieces.[0].Trim().ToLowerInvariant() = name then
                            let raw = pieces.[1].Trim().ToLowerInvariant()
                            let number (text: string) =
                                match System.Double.TryParse(text, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture) with
                                | true, parsed -> Some parsed
                                | false, _ -> None
                            if raw.EndsWith "pt" then number (raw.Substring(0, raw.Length - 2)) |> Option.map (fun pt -> pt * 96.0 / 72.0)
                            elif raw.EndsWith "in" then number (raw.Substring(0, raw.Length - 2)) |> Option.map (fun inches -> inches * 96.0)
                            elif raw.EndsWith "px" then number (raw.Substring(0, raw.Length - 2))
                            else number raw
                        else None)
                { blankImageData fallbackRelationship with
                    Width = styleNumber "width"
                    Height = styleNumber "height"
                    Floating = style.Replace(" ", "").ToLowerInvariant().Contains "position:absolute"
                    Wrap = Some "none" }
        let shapeCandidate = not (modelledShapeRoots node).IsEmpty
        let kind =
            if nodeName = "object" || not (Ns.descendantsNs "OLEObject" node).IsEmpty then "ole"
            elif not (Ns.descendantsNs "chart" node).IsEmpty then "chart"
            elif not (Ns.descendantsNs "relIds" node).IsEmpty then "diagram"
            elif shapeCandidate || nodeName = "pict" then "shape"
            else "drawing"
        { frame with
            Description = frame.Description |> Option.orElse frame.Name
            Opaque = Some { Xml = xmlToString node; Kind = kind; ParagraphLevel = paragraphLevel; ShapeCandidate = shapeCandidate } }

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
              RevisionDate = None; InsertedRevision = None; MoveName = None })

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
            InsertedRevision = firstResult |> Option.bind (fun run -> run.InsertedRevision)
            MoveName = firstResult |> Option.bind (fun run -> run.MoveName)
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
              RevisionDate = None; InsertedRevision = None; MoveName = None }

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
                    | "pict" ->
                        match imageData rNode with
                        | Some image -> Some (makeRun (Image image) |> RunToken)
                        | None -> Some (makeRun (Image (opaqueObject false child)) |> RunToken)
                    // An embedded (OLE) object, and a drawing offered in two forms.
                    | "object" -> Some (makeRun (Image (opaqueObject false child)) |> RunToken)
                    | "AlternateContent" when isGraphicAlternateContent child -> Some (makeRun (Image (opaqueObject false child)) |> RunToken)
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
        // A link to a bookmark of the document itself (w:anchor): "#name".
        let anchorTarget = attr "anchor" hlNode |> Option.filter (fun name -> name <> "") |> Option.map (fun name -> "#" + name)
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
                    Target = anchorTarget  // The relationship target is resolved later
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
                RevisionDate = None; InsertedRevision = None; MoveName = None
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
        let renderedPageBreaks = Ns.descendantsNs "lastRenderedPageBreak" pNode
        let hasRenderedPageBreak = not renderedPageBreaks.IsEmpty
        let renderedPageBreakBefore =
            match renderedBreakCounter with
            | Some counter ->
                let beforeText = hasRenderedPageBreakBeforeText pNode
                let breakCount = List.length renderedPageBreaks
                let ordinal = if beforeText then Some (!counter + 1) else None
                counter := !counter + breakCount
                ordinal
            | None -> None
        let rawProps = paragraphProps pNode
        let props =
            let withRenderedBreak = { rawProps with RenderedPageBreakBefore = renderedPageBreakBefore }
            if hasRenderedPageBreak && rawProps.Style = Some "Heading1" then
                { withRenderedBreak with PageBreakBefore = Some true; PageBreakBeforeDerived = rawProps.PageBreakBefore <> Some true }
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
                    [RunToken { Content = EquationXml (xmlToString child); Properties = runProps child; BookmarkStart = None; BookmarkEnd = None; RenderedPageBreakBefore = false; RevisionId = None; RevisionType = None; RevisionAuthor = None; RevisionDate = None; InsertedRevision = None; MoveName = None }]
                | Some name when Ns.stripNs name = "bookmarkStart" ->
                    attr "id" child
                    |> Option.map (fun id -> RunToken { Content = BookmarkStart { Id = id; Name = attr "name" child }; Properties = runProps child; BookmarkStart = None; BookmarkEnd = None; RenderedPageBreakBefore = false; RevisionId = None; RevisionType = None; RevisionAuthor = None; RevisionDate = None; InsertedRevision = None; MoveName = None })
                    |> Option.toList
                | Some name when Ns.stripNs name = "bookmarkEnd" ->
                    attr "id" child
                    |> Option.map (fun id -> RunToken { Content = BookmarkEnd id; Properties = runProps child; BookmarkStart = None; BookmarkEnd = None; RenderedPageBreakBefore = false; RevisionId = None; RevisionType = None; RevisionAuthor = None; RevisionDate = None; InsertedRevision = None; MoveName = None })
                    |> Option.toList
                // Transparent inline wrappers: their runs are ordinary paragraph content.
                | Some name when (match Ns.stripNs name with "smartTag" | "customXml" | "dir" | "bdo" -> true | _ -> false) ->
                    tokensFromNode child
                | Some name when Ns.stripNs name = "AlternateContent" ->
                    if isGraphicAlternateContent child then
                        // A drawing in two forms: kept whole rather than reduced to its fallback.
                        [RunToken { Content = Image (opaqueObject true child); Properties = runProps child; BookmarkStart = None; BookmarkEnd = None; RenderedPageBreakBefore = false; RevisionId = None; RevisionType = None; RevisionAuthor = None; RevisionDate = None; InsertedRevision = None; MoveName = None }]
                    else
                        // Take the fallback (plain WordprocessingML/VML) when there is one, else the first choice.
                        match Ns.childNs "Fallback" child |> Option.orElseWith (fun () -> Ns.childNs "Choice" child) with
                        | Some branch -> tokensFromNode branch
                        | None -> []
                // Tracked changes. A move is kept as a move (moveFrom at the source,
                // moveTo at the destination), not turned into a deletion and an insertion.
                | Some name when (match Ns.stripNs name with "ins" | "del" | "moveTo" | "moveFrom" -> true | _ -> false) ->
                    let revisionType =
                        match Ns.stripNs name with
                        | "del" -> "delete"
                        | "moveFrom" -> "moveFrom"
                        | "moveTo" -> "moveTo"
                        | _ -> "insert"
                    let revisionId = Xml.attr "id" child |> Option.orElse (Xml.attr "w:id" child)
                    let author = Xml.attr "author" child |> Option.orElse (Xml.attr "w:author" child)
                    let date = Xml.attr "date" child |> Option.orElse (Xml.attr "w:date" child)
                    tokensFromNode child
                    |> List.map (function
                        // A deletion inside this insertion stays a deletion and remembers
                        // the insertion it is in. (The outer change used to overwrite the
                        // inner one, so deleted text of an insertion read as inserted.)
                        | RunToken run when revisionType = "insert" && run.RevisionType = Some "delete" ->
                            RunToken { run with InsertedRevision = Some { Id = revisionId; Author = author; Date = date } }
                        | RunToken run when revisionType = "delete" && run.RevisionType = Some "insert" ->
                            RunToken { run with
                                        RevisionId = revisionId; RevisionType = Some "delete"; RevisionAuthor = author; RevisionDate = date
                                        InsertedRevision = Some { Id = run.RevisionId; Author = run.RevisionAuthor; Date = run.RevisionDate } }
                        | RunToken run when run.RevisionType.IsSome -> RunToken run
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
                        [RunToken { Content = Text glyph; Properties = runProps child; BookmarkStart = None; BookmarkEnd = None; RenderedPageBreakBefore = false; RevisionId = None; RevisionType = None; RevisionAuthor = None; RevisionDate = None; InsertedRevision = None; MoveName = None }]
                    | None ->
                        Ns.childNs "sdtContent" child
                        |> Option.map tokensFromNode
                        |> Option.defaultValue []
                | _ -> [])

        // Extract runs and hyperlinks
        let runs = tokensFromNode pNode |> collapseComplexFieldTokens |> attachRunMarkers |> normalizeTocRuns props

        // A tracked paragraph mark: w:ins / w:del among the mark's run properties.
        let markRevision =
            wElem "pPr" pNode
            |> Option.bind (wElem "rPr")
            |> Option.bind (fun markProps ->
                [ "ins", "insert"; "del", "delete" ]
                |> List.tryPick (fun (elementName, revisionType) ->
                    wElem elementName markProps
                    |> Option.map (fun node -> revisionType, { Id = attr "id" node; Author = attr "author" node; Date = attr "date" node })))

        { Runs = runs; Properties = props; MarkRevision = markRevision }

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
            FloatingXml = tblPr |> Option.bind (wElem "tblpPr") |> Option.map xmlToString
        }

    // Main document content extraction
    let private contentControlProps (sdtNode: XmlNode) =
        let props = Ns.childNs "sdtPr" sdtNode
        let id = props |> Option.bind (Ns.childNs "id") |> Option.bind (attr "val")
        let tag = props |> Option.bind (Ns.childNs "tag") |> Option.bind (attr "val")
        let alias = props |> Option.bind (Ns.childNs "alias") |> Option.bind (attr "val")
        id, tag, alias
    let private applyContentControl (id, tag, alias) = function
        | Para para ->
            Para { para with Properties = { para.Properties with ContentControlId = id; ContentControlTag = tag; ContentControlAlias = alias } }
        | item -> item

    /// The blocks one child element of the body (of a note, a header, a
    /// content control ...) gives: a paragraph, a table, the blocks of a
    /// block-level content control or of a custom XML element; nothing for
    /// anything else. `renderedBreakCounter` counts the rendered page breaks
    /// of the story's paragraphs, in document order.
    let rec contentOfChild (renderedBreakCounter: int ref) (child: XmlNode) : DocumentContent list =
        match Xml.name child with
        | Some name when Ns.stripNs name = "p" ->
            [ Para (paragraphWithRenderedBreakCounter (Some renderedBreakCounter) child) ]
        | Some name when Ns.stripNs name = "tbl" ->
            [ Tbl (table child) ]
        | Some name when Ns.stripNs name = "sdt" ->
            match Ns.childNs "sdtContent" child with
            | Some content ->
                let controlProps = contentControlProps child
                contentOfChildren renderedBreakCounter content |> List.map (applyContentControl controlProps)
            | None -> []
        | Some name when Ns.stripNs name = "customXml" ->
            contentOfChildren renderedBreakCounter child
        | _ -> []
    and private contentOfChildren (renderedBreakCounter: int ref) (node: XmlNode) : DocumentContent list =
        Xml.children node |> List.collect (contentOfChild renderedBreakCounter)

    /// An empty paragraph that ends a section on a new page, followed by a
    /// paragraph that ends a continuous one, does not break the section: what
    /// `item` is, given the block after it.
    let normalizedSectionBreak (item: DocumentContent) (next: DocumentContent option) : DocumentContent =
        match item, next with
        | Para para, Some (Para nextPara) when para.Runs.IsEmpty && para.Properties.SectionBreak = Some "nextPage" && nextPara.Properties.SectionBreak = Some "continuous" ->
            Para { para with Properties = { para.Properties with SectionBreak = None } }
        | _ -> item

    let bodyContentWithFinalSection (finalSectPr: XmlNode option) (bodyNode: XmlNode) : DocumentContent list =
        let renderedBreakCounter = ref 0

        // Iterative: the recursive form was not a tail call and overflowed the JS
        // stack on documents with a few thousand body blocks (the file would not open).
        let normalizeSectionBreaks (items: DocumentContent list) =
            let source = List.toArray items
            source
            |> Array.mapi (fun index item -> normalizedSectionBreak item (if index + 1 < source.Length then Some source.[index + 1] else None))
            |> Array.toList

        // Every section's properties stay on the section they belong to: a
        // w:sectPr inside a paragraph describes the section that ends with that
        // paragraph, the body-level one the final section (PageSetup). The final
        // section's header / footer references used to be copied onto the
        // paragraph after the first section break, which put the last section's
        // header on the second section of a document with three or more.
        ignore finalSectPr
        contentOfChildren renderedBreakCounter bodyNode |> normalizeSectionBreaks

    let bodyContent (bodyNode: XmlNode) : DocumentContent list =
        bodyContentWithFinalSection None bodyNode

    /// For each block bodyContent gives, the index of the child element of
    /// w:body it comes from (a block-level content control gives several
    /// blocks), and the number of child elements.
    let bodyContentChunks (bodyNode: XmlNode) : int list * int =
        let rec blockCount (node: XmlNode) =
            match Xml.name node |> Option.map Ns.stripNs with
            | Some "p" | Some "tbl" -> 1
            | Some "sdt" ->
                match Ns.childNs "sdtContent" node with
                | Some content -> Xml.children content |> List.sumBy blockCount
                | None -> 0
            | Some "customXml" -> Xml.children node |> List.sumBy blockCount
            | _ -> 0
        let elements = Xml.children bodyNode |> List.filter (fun child -> (Xml.name child).IsSome)
        elements |> List.mapi (fun index element -> List.replicate (blockCount element) index) |> List.concat, elements.Length

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
            PageBreakBeforeDerived = false
            WidowControl = Docx.onOff "widowControl" pPr
            ContextualSpacing = Docx.onOff "contextualSpacing" pPr
            BackgroundColor = Ns.childNs "shd" pPr |> Option.bind (Docx.attr "fill")
            ParagraphRunProperties = Ns.childNs "rPr" pPr |> Option.map (fun _ -> Docx.runProps pPr)
            ParagraphBorders = Docx.paragraphBorders pPr
            ListInfo = listInfo
            SectionBreak = None
            SectionGeometry = None
            SectionColumns = None
            SectionPageNumberStart = None; SectionPageNumberFormat = None
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

/// Image key prefixes of the parts that have no relationship id of their own use.
let footnotesPartKey = "footnotes"
let endnotesPartKey = "endnotes"
let commentsPartKey = "comments"

/// Apply `mapParagraph` to every paragraph in `blocks`, nested tables included.
let rec mapBlockParagraphs (mapParagraph: Paragraph -> Paragraph) (blocks: DocumentContent list) : DocumentContent list =
    let rec mapTable (table: Table) =
        { table with
            Rows =
                table.Rows
                |> List.map (fun row ->
                    { row with
                        Cells =
                            row.Cells
                            |> List.map (fun cell ->
                                { cell with
                                    Paragraphs = cell.Paragraphs |> List.map mapParagraph
                                    NestedTables = cell.NestedTables |> List.map mapTable }) }) }
    blocks
    |> List.map (function
        | Para para -> Para (mapParagraph para)
        | Tbl table -> Tbl (mapTable table))

/// Apply `mapRun` to every run of every paragraph in `blocks`, nested tables included.
let mapBlockRuns (mapRun: Run -> Run) (blocks: DocumentContent list) : DocumentContent list =
    mapBlockParagraphs (fun para -> { para with Runs = para.Runs |> List.map mapRun }) blocks

let private xmlnsDeclarations (attrs: Map<string, string>) : Map<string, string> =
    attrs
    |> Map.toList
    |> List.choose (fun (name, value) -> if name.StartsWith "xmlns:" then Some (name.Substring 6, value) else None)
    |> Map.ofList

let private prefixPattern = System.Text.RegularExpressions.Regex(@"[<\s/]([A-Za-z_][A-Za-z0-9_.\-]*):[A-Za-z_]")
let private startTagAttributePattern = System.Text.RegularExpressions.Regex("\\s([^\\s=]+)=\"([^\"]*)\"")
let private prefixListPattern = System.Text.RegularExpressions.Regex("(?:Requires|Ignorable)=\"([^\"]*)\"")

/// `xml` (one element, serialised) with the namespace declarations of the
/// part's root element it relies on added to its own root, so it stays valid
/// wherever it is written. `rootAttributes` are the attributes of the root
/// element of the part the fragment was taken from.
let selfContainedXml (rootAttributes: Map<string, string>) (xml: string) : string =
    let rootNamespaces = xmlnsDeclarations rootAttributes
    let tagEnd = xml.IndexOf '>'
    if rootNamespaces.IsEmpty || tagEnd < 0 then xml
    else
        let used =
            [ for m in prefixPattern.Matches xml do yield m.Groups.[1].Value
              for m in prefixListPattern.Matches xml do yield! m.Groups.[1].Value.Split([| ' ' |], StringSplitOptions.RemoveEmptyEntries) ]
            |> Set.ofList
        let startTag = xml.Substring(0, tagEnd)
        // Prefixes the part declares ignorable and the fragment uses.
        let ignorable =
            match Map.tryFind "mc:Ignorable" rootAttributes with
            | Some value when not (startTag.Contains "mc:Ignorable=") ->
                value.Split([| ' ' |], StringSplitOptions.RemoveEmptyEntries) |> Array.filter (fun prefix -> Set.contains prefix used) |> Array.toList
            | _ -> []
        let used = if ignorable.IsEmpty then used else Set.add "mc" used
        let additions =
            [ for KeyValue(prefix, uri) in rootNamespaces do
                if Set.contains prefix used && not (startTag.Contains ("xmlns:" + prefix + "=")) then
                    yield "xmlns:" + prefix, uri.Replace("&", "&amp;").Replace("\"", "&quot;").Replace("<", "&lt;")
              if not ignorable.IsEmpty then yield "mc:Ignorable", String.concat " " ignorable ]
        if additions.IsEmpty then xml
        else
            // The root start tag is rebuilt with its attributes in name order, the
            // order the serialiser writes them in: reading the result again and
            // serialising it gives the same text.
            let mutable nameEnd = 1
            while nameEnd < xml.Length && xml.[nameEnd] <> ' ' && xml.[nameEnd] <> '>' && xml.[nameEnd] <> '/' do
                nameEnd <- nameEnd + 1
            let selfClosed = tagEnd > 0 && xml.[tagEnd - 1] = '/'
            let attributeText = xml.Substring(nameEnd, (if selfClosed then tagEnd - 1 else tagEnd) - nameEnd)
            let existing = [ for m in startTagAttributePattern.Matches attributeText -> m.Groups.[1].Value, m.Groups.[2].Value ]
            let attributes =
                (existing @ additions)
                |> Map.ofList
                |> Map.toList
                |> List.map (fun (name, value) -> " " + name + "=\"" + value + "\"")
                |> String.concat ""
            xml.Substring(0, nameEnd) + attributes + xml.Substring(if selfClosed then tagEnd - 1 else tagEnd)

/// Make the XML of every opaque object in `blocks` self-contained (see selfContainedXml).
let completeOpaqueObjects (root: XmlNode) (blocks: DocumentContent list) : DocumentContent list =
    match root with
    | XmlDsl.Element(_, attrs, _) ->
        blocks
        |> mapBlockRuns (fun run ->
            match run.Content with
            | Image ({ Opaque = Some opaque } as image) ->
                { run with Content = Image { image with Opaque = Some { opaque with Xml = selfContainedXml attrs opaque.Xml } } }
            | _ -> run)
    | _ -> blocks

/// The move range each w:moveFrom / w:moveTo element lies in: (kind, revision id)
/// -> range name. Ranges are opened and closed by marker elements that may sit
/// in other paragraphs, so they are followed through the part in document order.
let private moveNamesOfPart (root: XmlNode) : Map<string * string, string> =
    let mutable names = Map.empty
    let mutable fromName : string option = None
    let mutable toName : string option = None
    let rec walk (node: XmlNode) =
        match node with
        | XmlDsl.Element (name, _, children) ->
            match Ns.stripNs name with
            | "moveFromRangeStart" -> fromName <- Docx.attr "name" node
            | "moveFromRangeEnd" -> fromName <- None
            | "moveToRangeStart" -> toName <- Docx.attr "name" node
            | "moveToRangeEnd" -> toName <- None
            | "moveFrom" ->
                match Docx.attr "id" node, fromName with
                | Some id, Some rangeName -> names <- Map.add ("moveFrom", id) rangeName names
                | _ -> ()
            | "moveTo" ->
                match Docx.attr "id" node, toName with
                | Some id, Some rangeName -> names <- Map.add ("moveTo", id) rangeName names
                | _ -> ()
            | _ -> ()
            for child in children do walk child
        | _ -> ()
    walk root
    names

/// Give the moved runs of `blocks` the name of their move range.
let withMoveNames (root: XmlNode) (blocks: DocumentContent list) : DocumentContent list =
    let names = moveNamesOfPart root
    if names.IsEmpty then blocks
    else
        blocks
        |> mapBlockRuns (fun run ->
            match run.RevisionType, run.RevisionId with
            | Some kind, Some id when kind = "moveFrom" || kind = "moveTo" ->
                match Map.tryFind (kind, id) names with
                | Some name -> { run with MoveName = Some name }
                | None -> run
            | _ -> run)

/// The body models these as shapes (text boxes); the opaque copy is dropped there.
let private dropModelledShapeRuns (blocks: DocumentContent list) : DocumentContent list =
    blocks
    |> mapBlockParagraphs (fun para ->
        if para.Runs |> List.exists (fun run -> match run.Content with Image { Opaque = Some { ShapeCandidate = true } } -> true | _ -> false) then
            { para with Runs = para.Runs |> List.filter (fun run -> match run.Content with Image { Opaque = Some { ShapeCandidate = true } } -> false | _ -> true) }
        else para)

/// Relationship ids are local to a part: the header's rId1 is not the body's
/// rId1. Image references of a part other than the body are qualified with a
/// key of that part so they cannot collide in the document-wide image map.
let qualifyImageRun (partKey: string) (run: Run) : Run =
    match run.Content with
    | Image image -> { run with Content = Image { image with RelationshipId = partImageKey partKey image.RelationshipId } }
    | _ -> run

/// Hyperlink targets live in the relationships of the part the link is in.
let resolveHyperlinkRun (relMap: Map<string, string>) (run: Run) : Run =
    match run.Content with
    | Link link ->
        match link.RelationshipId |> Option.bind (fun rId -> Map.tryFind rId relMap) with
        | Some target ->
            // With a w:anchor as well, the anchor is the fragment of the target.
            let fragment = match link.Target with | Some (anchor: string) when anchor.StartsWith "#" && not (target.Contains "#") -> anchor | _ -> ""
            { run with Content = Link { link with Target = Some (target + fragment) } }
        | None -> run
    | _ -> run

let private hyperlinkTargets (rels: Relationship list) =
    rels
    |> List.choose (fun rel -> if rel.Target <> "" then Some (rel.Id, rel.Target) else None)
    |> Map.ofList

/// Paragraphs of `blocks` in document order, table cells included.
let rec blockParagraphs (blocks: DocumentContent list) : Paragraph list =
    let rec tableParagraphs (table: Table) =
        table.Rows
        |> List.collect (fun row ->
            row.Cells
            |> List.collect (fun cell -> cell.Paragraphs @ (cell.NestedTables |> List.collect tableParagraphs)))
    blocks
    |> List.collect (function
        | Para para -> [para]
        | Tbl table -> tableParagraphs table)

// Header/Footer parsing
module HeaderFooter =
    let parseHeaderFooter (xml: string) (hfType: string) (relationshipId: string option) : HeaderFooter =
        let doc = XmlDsl.parse xml
        let qualify blocks =
            match relationshipId with
            | Some partId -> mapBlockRuns (qualifyImageRun partId) blocks
            | None -> blocks
        let blocks = Docx.bodyContent doc |> qualify |> completeOpaqueObjects doc |> withMoveNames doc
        {
            Content =
                Ns.descendantsNs "p" doc
                |> List.map (fun node -> Para (Docx.paragraph node))
                |> qualify
                |> completeOpaqueObjects doc
                |> List.choose (function Para para -> Some para | Tbl _ -> None)
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

            let blocks = Docx.bodyContent commentNode |> mapBlockRuns (qualifyImageRun commentsPartKey)
            let paragraphs = blockParagraphs blocks
            let runText (r: Run) =
                match r.Content with
                | Text t -> Some t
                | Link link -> Some link.Text
                | _ -> None

            // Runs of a paragraph join without a separator; paragraphs with a space.
            let text =
                paragraphs
                |> List.map (fun p -> p.Runs |> List.choose runText |> String.concat "")
                |> List.filter (fun line -> line <> "")
                |> String.concat " "

            let initialText =
                paragraphs
                |> List.tryHead
                |> Option.bind (fun p -> p.Runs |> List.tryPick runText)

            {
                Id = id
                Author = author
                Date = date
                Text = text
                InitialText = initialText
                Initials = Docx.attr "initials" commentNode
                Blocks = blocks
            })

    let parseComments (xml: string) : Comment list =
        let doc = XmlDsl.parse xml
        Ns.descendantsNs "comment" doc
        |> List.choose parseComment
        |> List.map (fun comment -> { comment with Blocks = completeOpaqueObjects doc comment.Blocks })

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

    /// `mayHold name`: false when the part is known to hold no element of that local name.
    let parseRevisionsWhere (mayHold: string -> bool) (doc: XmlNode) : Revision list =
        let revisionTypes = ["ins"; "del"; "rPrChange"; "moveFrom"; "moveTo"]
        revisionTypes
        |> List.collect (fun revType ->
            if not (mayHold revType) then []
            else
                Ns.descendantsNs revType doc
                |> List.choose parseRevision)

    let parseRevisions (doc: XmlNode) : Revision list =
        parseRevisionsWhere (fun _ -> true) doc

// Footnotes and Endnotes parsing
module Notes =
    let private parseNoteOfPart (partKey: string) (noteNode: XmlNode) : Note option =
        Docx.attr "id" noteNode
        |> Option.map (fun id ->
            let blocks = Docx.bodyContent noteNode |> mapBlockRuns (qualifyImageRun partKey)
            // The run holding the automatic reference mark.
            let markRun =
                Ns.descendantsNs "r" noteNode
                |> List.tryFind (fun run ->
                    Xml.children run
                    |> List.exists (fun child ->
                        match Xml.name child |> Option.map Ns.stripNs with
                        | Some "footnoteRef" | Some "endnoteRef" -> true
                        | _ -> false))
            {
                Id = id
                Content = blockParagraphs blocks
                Blocks = blocks
                Kind = Docx.attr "type" noteNode |> Option.filter (fun kind -> kind <> "normal")
                HasMark = markRun.IsSome
                MarkStyle = markRun |> Option.bind (fun run -> (Docx.runProps run).Style)
            })

    let parseNote (noteNode: XmlNode) : Note option =
        let partKey = if Xml.name noteNode |> Option.map Ns.stripNs = Some "endnote" then endnotesPartKey else footnotesPartKey
        parseNoteOfPart partKey noteNode

    let parseNotes (xml: string) : Note list =
        let doc = XmlDsl.parse xml
        Ns.descendantsNs "footnote" doc
        |> List.choose parseNote
        |> fun footnotes ->
            let endnotes =
                Ns.descendantsNs "endnote" doc
                |> List.choose parseNote
            footnotes @ endnotes
        |> List.map (fun note ->
            let blocks = completeOpaqueObjects doc note.Blocks
            { note with Blocks = blocks; Content = blockParagraphs blocks })

    /// The text the note is shown with: every text run of the body, joined.
    let plainText (note: Note) : string =
        note.Content
        |> List.collect (fun para -> para.Runs)
        |> List.map (fun run ->
            match run.Content with
            | Text text -> text
            | Link link -> link.Text
            | _ -> "")
        |> String.concat ""

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
                elif raw.EndsWith("in") then raw.Substring(0, raw.Length - 2) |> parseFloat |> Option.map (fun inches -> inches * 96.0)
                elif raw.EndsWith("mm") then raw.Substring(0, raw.Length - 2) |> parseFloat |> Option.map (fun mm -> mm * 96.0 / 25.4)
                elif raw.EndsWith("cm") then raw.Substring(0, raw.Length - 2) |> parseFloat |> Option.map (fun cm -> cm * 96.0 / 2.54)
                else parseFloat raw
            else None)

    let private namedColors =
        Map.ofList [ "black", "#000000"; "white", "#FFFFFF"; "silver", "#C0C0C0"; "gray", "#808080"; "grey", "#808080"; "red", "#FF0000"; "maroon", "#800000"
                     "yellow", "#FFFF00"; "olive", "#808000"; "lime", "#00FF00"; "green", "#008000"; "aqua", "#00FFFF"; "teal", "#008080"; "blue", "#0000FF"
                     "navy", "#000080"; "fuchsia", "#FF00FF"; "purple", "#800080"; "orange", "#FFA500" ]

    /// A colour as VML writes it: "#rgb", "#rrggbb", a name, each with or without a palette entry after it ("silver [3209]").
    let private normalizeColor (value: string option) =
        match value |> Option.map (fun text -> (text.Trim().Split(' ')).[0]) with
        | Some "" | None -> None
        | Some color when color.StartsWith("#") && color.Length = 4 ->
            Some (sprintf "#%c%c%c%c%c%c" color.[1] color.[1] color.[2] color.[2] color.[3] color.[3])
        | Some color when color.StartsWith("#") -> Some color
        | Some color ->
            match Map.tryFind (color.ToLowerInvariant()) namedColors with
            | Some hex -> Some hex
            | None -> Some ("#" + color)

    /// The value of a property of a VML style ("name:value;name:value").
    let private styleText (name: string) (style: string) : string option =
        style.Split(';')
        |> Array.tryPick (fun part ->
            let text = part.Trim()
            if text.ToLowerInvariant().StartsWith(name + ":") then Some (text.Substring(name.Length + 1).Trim()) else None)

    /// A VML length in pixels: points, inches, millimetres, centimetres, pixels; EMU without a unit.
    let private vmlLength (text: string) : float option =
        let raw = text.Trim().ToLowerInvariant()
        let number (suffix: string) (factor: float) = raw.Substring(0, raw.Length - suffix.Length) |> parseFloat |> Option.map (fun value -> value * factor)
        if raw.EndsWith "pt" then number "pt" (96.0 / 72.0)
        elif raw.EndsWith "in" then number "in" 96.0
        elif raw.EndsWith "mm" then number "mm" (96.0 / 25.4)
        elif raw.EndsWith "cm" then number "cm" (96.0 / 2.54)
        elif raw.EndsWith "px" then number "px" 1.0
        else parseFloat raw |> Option.map (fun emu -> emu / 9525.0)

    let private xmlEscaped (text: string) =
        text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;")

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

    let private childNamed (name: string) (node: XmlNode) =
        Xml.children node |> List.tryFind (fun child -> Xml.name child |> Option.map Ns.stripNs = Some name)

    /// The shape's own properties element (wps:spPr), not a nested picture's.
    let private shapeProperties (shapeNode: XmlNode) =
        Ns.descendantsNs "spPr" shapeNode |> List.tryHead

    let private vmlBool (value: string option) =
        value |> Option.map (fun text -> let t = text.ToLowerInvariant() in t = "t" || t = "true" || t = "1")

    let parseShape (shapeNode: XmlNode) : Shape option =
        let id = Xml.attr "id" shapeNode
        let shapeType = Xml.attr "type" shapeNode

        let txbxContent = Ns.descendantsNs "txbxContent" shapeNode |> List.tryHead
        let vmlStyle = Xml.attr "style" shapeNode |> Option.defaultValue ""
        // WordArt (a watermark): the text of a v:textpath, in its face and the shape's fill colour.
        let textPath = Ns.descendantsNs "textpath" shapeNode |> List.tryFind (fun path -> (Xml.attr "string" path).IsSome)
        let paragraphs =
            match txbxContent, textPath with
            | Some content, _ ->
                Ns.descendantsNs "p" content
                |> List.map Docx.paragraph
            | None, Some path ->
                let pathStyle = Xml.attr "style" path |> Option.defaultValue ""
                let face = styleText "font-family" pathStyle |> Option.map (fun name -> name.Replace("&quot;", "").Replace("\"", "").Trim())
                let color = Xml.attr "fillcolor" shapeNode |> normalizeColor |> Option.map (fun hex -> hex.TrimStart('#'))
                let properties =
                    (match face with Some name when name <> "" -> sprintf "<w:rFonts w:ascii=\"%s\" w:hAnsi=\"%s\"/>" (xmlEscaped name) (xmlEscaped name) | _ -> "")
                    + (if (styleText "font-weight" pathStyle) = Some "bold" then "<w:b/>" else "")
                    + (if (styleText "font-style" pathStyle) = Some "italic" then "<w:i/>" else "")
                    + (match color with Some hex -> sprintf "<w:color w:val=\"%s\"/>" hex | None -> "")
                let text = Xml.attr "string" path |> Option.defaultValue ""
                [ XmlDsl.parse (sprintf "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:pPr><w:jc w:val=\"center\"/></w:pPr><w:r><w:rPr>%s</w:rPr><w:t xml:space=\"preserve\">%s</w:t></w:r></w:p>" properties (xmlEscaped text))
                  |> Docx.paragraph ]
            | None, None -> []
        let bodyPr = Ns.descendantsNs "bodyPr" shapeNode |> List.tryHead
        let vmlTextBox = Ns.descendantsNs "textbox" shapeNode |> List.tryHead
        let vmlInsets =
            vmlTextBox
            |> Option.bind (Xml.attr "inset")
            |> Option.map (fun inset -> inset.Split(',') |> Array.map vmlLength)
            |> Option.defaultValue [||]
        let inset (name: string) (index: int) =
            bodyPr |> Option.bind (Xml.attr name) |> Option.bind parseFloat |> Option.map emuToPixels
            |> Option.orElseWith (fun () -> if index < vmlInsets.Length then vmlInsets.[index] else None)

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
        let spPr = shapeProperties shapeNode
        let line = spPr |> Option.bind (childNamed "ln")
        let fillColor =
            (Xml.attr "fillcolor" shapeNode |> normalizeColor) |> Option.orElseWith (fun () -> parseDmlFill shapeNode)
        let borderColor =
            (Xml.attr "strokecolor" shapeNode |> normalizeColor) |> Option.orElseWith (fun () -> parseDmlLine shapeNode)
        // a:ln@w is EMU (12700 per point); VML strokeweight carries its unit.
        let borderWidth =
            Xml.attr "strokeweight" shapeNode |> strokeWeightFromAttr
            |> Option.orElseWith (fun () -> line |> Option.bind (Xml.attr "w") |> Option.bind parseFloat |> Option.map emuToPixels)
        let custom = parseCustom shapeNode
        // VML shapes use a "#_x0000_t202"-style type ref; only surface a clean
        // DrawingML preset name as the geometry hint.
        let preset =
            parsePreset shapeNode
            |> Option.orElseWith (fun () ->
                match shapeNode with
                | XmlDsl.Element ("v:oval", _, _) -> Some "ellipse"
                | XmlDsl.Element ("v:roundrect", _, _) -> Some "roundRect"
                | _ -> None)
        let isVml = match shapeNode with XmlDsl.Element (name, _, _) -> name.StartsWith "v:" | _ -> false
        // A VML shape says in its style where it stands: aligned or by its margins, relative to what, behind the text.
        let vmlAlign (name: string) = if isVml then styleText name vmlStyle |> Option.filter (fun value -> value <> "absolute" && value <> "") else None
        let vmlRelative (name: string) =
            if not isVml then None
            else
                match styleText name vmlStyle with
                | Some "margin" -> Some "margin"
                | Some "page" -> Some "page"
                | Some "text" -> Some (if name.Contains "horizontal" then "column" else "paragraph")
                | Some "char" -> Some "character"
                | Some "line" -> Some "line"
                | _ -> None
        let vmlWrap =
            if not isVml then None
            else
                Ns.descendantsNs "wrap" shapeNode |> List.tryHead |> Option.bind (Xml.attr "type")
                |> Option.map (fun kind -> match kind with "topAndBottom" | "square" | "tight" | "through" | "none" -> kind | _ -> "none")
        let vmlBehind =
            if not isVml then None
            else styleText "z-index" vmlStyle |> Option.bind parseFloat |> Option.bind (fun z -> if z < 0.0 then Some true else None)
        let vmlRotation =
            if not isVml then None
            else
                styleText "rotation" vmlStyle
                |> Option.bind (fun text -> if text.EndsWith "fd" then text.Substring(0, text.Length - 2) |> parseFloat |> Option.map (fun fd -> fd / 65536.0) else parseFloat text)
        let isAnchor = Xml.name shapeNode |> Option.map Ns.stripNs = Some "anchor"
        let isInline = Xml.name shapeNode |> Option.map Ns.stripNs = Some "inline"
        let anchorBool name = if isAnchor then Xml.attr name shapeNode |> Option.map (fun value -> value = "1" || value.ToLowerInvariant() = "true") else None
        let anchorDistance name = if isAnchor || isInline then Xml.attr name shapeNode |> Option.bind parseFloat |> Option.map emuToPixels else None
        let wrapChild =
            if not isAnchor then None
            else
                Xml.children shapeNode
                |> List.tryPick (fun child ->
                    match Xml.name child |> Option.map Ns.stripNs with
                    | Some "wrapNone" -> Some "none"
                    | Some "wrapSquare" -> Some "square"
                    | Some "wrapTight" -> Some "tight"
                    | Some "wrapThrough" -> Some "through"
                    | Some "wrapTopAndBottom" -> Some "topAndBottom"
                    | _ -> None)

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
            Rotation = vmlRotation |> Option.defaultWith (fun () -> parseRotation shapeNode)
            AnchorBlock = None
            Inline = isInline || (isVml && not (vmlStyle.Replace(" ", "").ToLowerInvariant().Contains "position:absolute"))
            Wrap = wrapChild |> Option.orElse vmlWrap
            PositionHRelativeFrom = posH |> Option.bind (Xml.attr "relativeFrom") |> Option.orElse (vmlRelative "mso-position-horizontal-relative")
            PositionVRelativeFrom = posV |> Option.bind (Xml.attr "relativeFrom") |> Option.orElse (vmlRelative "mso-position-vertical-relative")
            PositionHAlign = posH |> Option.bind (Ns.childNs "align") |> Option.bind Xml.text |> Option.orElse (vmlAlign "mso-position-horizontal")
            PositionVAlign = posV |> Option.bind (Ns.childNs "align") |> Option.bind Xml.text |> Option.orElse (vmlAlign "mso-position-vertical")
            BehindDoc = anchorBool "behindDoc" |> Option.orElse vmlBehind
            AllowOverlap = anchorBool "allowOverlap"
            LayoutInCell = anchorBool "layoutInCell"
            RelativeHeight = if isAnchor then Xml.attr "relativeHeight" shapeNode else None
            DistanceTop = anchorDistance "distT"
            DistanceBottom = anchorDistance "distB"
            DistanceLeft = anchorDistance "distL"
            DistanceRight = anchorDistance "distR"
            NoFill =
                (spPr |> Option.exists (fun props -> (childNamed "noFill" props).IsSome))
                || vmlBool (Xml.attr "filled" shapeNode) = Some false
                // (The fill of WordArt is the colour of its letters.)
                || (txbxContent.IsNone && textPath.IsSome)
            NoLine =
                (line |> Option.exists (fun ln -> (childNamed "noFill" ln).IsSome))
                || vmlBool (Xml.attr "stroked" shapeNode) = Some false
                || (txbxContent.IsNone && textPath.IsSome)
            LineDash = line |> Option.bind (childNamed "prstDash") |> Option.bind (Xml.attr "val")
            // (A VML shape is known by its id: a watermark's is how Word finds it.)
            Name =
                Ns.descendantsNs "docPr" shapeNode |> List.tryHead |> Option.bind (Xml.attr "name")
                |> Option.orElseWith (fun () -> if isVml then Xml.attr "id" shapeNode else None)
            InsetLeft = inset "lIns" 0
            InsetTop = inset "tIns" 1
            InsetRight = inset "rIns" 2
            InsetBottom = inset "bIns" 3
            VerticalAnchor =
                bodyPr |> Option.bind (Xml.attr "anchor")
                |> Option.orElseWith (fun () ->
                    match styleText "v-text-anchor" vmlStyle with
                    | Some "middle" -> Some "ctr"
                    | Some "bottom" -> Some "b"
                    | Some "top" -> Some "t"
                    | _ -> None)
            AutoFit =
                (bodyPr |> Option.exists (fun body -> (childNamed "spAutoFit" body).IsSome))
                || (vmlTextBox |> Option.bind (Xml.attr "style") |> Option.exists (fun style -> style.Replace(" ", "").ToLowerInvariant().Contains "mso-fit-shape-to-text:t"))
            NoTextWrap =
                (bodyPr |> Option.bind (Xml.attr "wrap") = Some "none")
                || (isVml && styleText "mso-wrap-style" vmlStyle = Some "none")
            FitText = txbxContent.IsNone && textPath.IsSome
            Opacity =
                // (Of WordArt: the opacity of its letters.)
                (if txbxContent.IsNone && textPath.IsSome then Ns.descendantsNs "fill" shapeNode else [])
                |> List.tryPick (Xml.attr "opacity")
                |> Option.bind (fun text -> if text.EndsWith "f" then text.Substring(0, text.Length - 1) |> parseFloat |> Option.map (fun f -> f / 65536.0) else parseFloat text)
                |> Option.filter (fun value -> value >= 0.0 && value < 1.0)
            SourceXml = None
        }

    /// Shapes of the document body, each with the index of the body block
    /// (as Docx.bodyContent numbers them) it is anchored in.
    /// The elements under one child element of the body that are blocks (as
    /// Docx.contentOfChild gives one block for each), in document order.
    let rec blockNodesOfChild (child: XmlNode) : XmlNode list =
        match Xml.name child |> Option.map Ns.stripNs with
        | Some "p" | Some "tbl" -> [child]
        | Some "sdt" -> Ns.childNs "sdtContent" child |> Option.map (fun content -> Xml.children content |> List.collect blockNodesOfChild) |> Option.defaultValue []
        | Some "customXml" -> Xml.children child |> List.collect blockNodesOfChild
        | _ -> []

    /// The shapes anchored in the body block of an index.
    let shapesOfBlock (index: int) (block: XmlNode) : Shape list =
        Docx.modelledShapeRoots block
        |> List.choose parseShape
        |> List.map (fun shape -> { shape with AnchorBlock = Some index })

    /// The shapes of a header or footer part (`root`: w:hdr / w:ftr), each with
    /// the index of the part's block it is anchored in.
    let parseShapesOfPart (root: XmlNode) : Shape list =
        Xml.children root
        |> List.collect blockNodesOfChild
        |> List.mapi shapesOfBlock
        |> List.concat

    /// The shapes of a block with the element each was read from: the
    /// block's opaque runs that are shapes (their XML complete with its
    /// namespaces - see completeOpaqueObjects), in the order of the shapes.
    /// A shape whose element is not certain (a w:pict of several shapes) has none.
    let withShapeSources (completed: DocumentContent list) (shapes: Shape list) : Shape list =
        if shapes.IsEmpty then shapes
        else
            let sources =
                blockParagraphs completed
                |> List.collect (fun para ->
                    para.Runs |> List.choose (fun run ->
                        match run.Content with
                        | Image { Opaque = Some opaque } when opaque.ShapeCandidate -> Some opaque.Xml
                        | _ -> None))
            let names (shape: Shape) (xml: string) =
                match shape.Name, shape.Id with
                | Some name, _ -> xml.Contains ("name=\"" + xmlEscaped name + "\"")
                | None, Some id -> xml.Contains ("id=\"" + xmlEscaped id + "\"")
                | None, None -> true
            if sources.Length = shapes.Length && List.forall2 names shapes sources then
                List.map2 (fun (shape: Shape) xml -> { shape with SourceXml = Some xml }) shapes sources
            else shapes

    /// `withShapeSources` for the shapes of several blocks (`blocks`: those
    /// Shape.AnchorBlock counts, before the shapes' runs are dropped).
    let withShapeSourcesOf (blocks: DocumentContent list) (shapes: Shape list) : Shape list =
        if shapes.IsEmpty then shapes
        else
            let blockArray = List.toArray blocks
            shapes
            |> List.groupBy (fun shape -> shape.AnchorBlock)
            |> List.collect (fun (anchor, group) ->
                match anchor with
                | Some index when index >= 0 && index < blockArray.Length -> withShapeSources [ blockArray.[index] ] group
                | _ -> group)

    let parseShapes (doc: XmlNode) : Shape list =
        match Ns.childNs "body" doc with
        | Some body ->
            Xml.children body
            |> List.collect blockNodesOfChild
            |> List.mapi shapesOfBlock
            |> List.concat
        | None -> Docx.modelledShapeRoots doc |> List.choose parseShape

// Resolve hyperlink targets from relationships
let resolveHyperlinks (doc: WordDocument) : WordDocument =
    { doc with Content = doc.Content |> mapBlockRuns (resolveHyperlinkRun (hyperlinkTargets doc.Relationships)) }

/// Hyperlinks of a header / footer resolve against that part's own relationships.
let private resolveHeaderFooterHyperlinks (rels: Relationship list) (part: HeaderFooter) : HeaderFooter =
    if rels.IsEmpty then part
    else
        let resolve = mapBlockRuns (resolveHyperlinkRun (hyperlinkTargets rels))
        { part with
            Blocks = resolve part.Blocks
            Content = part.Content |> List.map Para |> resolve |> List.choose (function Para para -> Some para | Tbl _ -> None) }

let private resolveNoteHyperlinks (rels: Relationship list) (note: Note) : Note =
    if rels.IsEmpty then note
    else
        let blocks = note.Blocks |> mapBlockRuns (resolveHyperlinkRun (hyperlinkTargets rels))
        { note with Blocks = blocks; Content = blockParagraphs blocks }

let private resolveCommentHyperlinks (rels: Relationship list) (comment: Comment) : Comment =
    if rels.IsEmpty then comment
    else { comment with Blocks = comment.Blocks |> mapBlockRuns (resolveHyperlinkRun (hyperlinkTargets rels)) }

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
            HeaderReferences = Docx.headerFooterReferences "headerReference" section
            FooterReferences = Docx.headerFooterReferences "footerReference" section
            PageNumberStart =
                Docx.getPropVal "pgNumType" "start" section
                |> Option.bind (fun value -> match System.Int32.TryParse(value) with | true, parsed -> Some parsed | false, _ -> None)
            PageNumberFormat = Docx.getPropVal "pgNumType" "fmt" section
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

/// The relationships part of the part at `partPath` (word/footnotes.xml -> word/_rels/footnotes.xml.rels).
let private partRelsPath (partPath: string) =
    let dir = parentPath partPath
    (if dir = "" then "" else dir + "/") + "_rels/" + lastPathSegment partPath + ".rels"

/// The package path of the part a relationship of the given type points at, or `fallback`.
let private partPathOfType (baseDir: string) (rels: Relationship list) (typeSuffix: string) (fallback: string) =
    rels
    |> List.tryFind (fun rel -> rel.Type.EndsWith typeSuffix && rel.TargetMode <> Some "External")
    |> Option.map (fun rel -> ZipInterop.resolveTarget baseDir rel.Target)
    |> Option.defaultValue fallback

let private relationshipsOfXml (xml: string option) =
    xml
    |> Option.map XmlDsl.parse
    |> Option.map parseRelationships
    |> Option.defaultValue []

let private relationshipTargetBase (rel: Relationship) =
    parentPath (ZipInterop.resolveTarget "word" rel.Target)

/// The parts of a package, read on demand (.NET) or loaded up front (browser).
type PartSource = {
    Text: string -> string option
    Bytes: string -> byte[] option
}

let private dataUrlOfBytes (mimeType: string) (bytes: byte[]) =
    "data:" + mimeType + ";base64," + Convert.ToBase64String bytes

/// Images of a part: (relationship id, data URL). A linked (external) picture
/// has no data; its target URL is delivered as the source and is marked Linked
/// on the run, so that nothing fetches it unless the host allows it.
let private loadImages (source: PartSource) (baseDir: string) (rels: Relationship list) : (string * string) list =
    rels
    |> List.choose (fun rel ->
        if rel.Type.Contains("/image") then
            if rel.TargetMode = Some "External" then
                Some (rel.Id, rel.Target)
            else
                source.Bytes (resolveRelationshipTarget baseDir rel.Target)
                |> Option.map (fun bytes -> rel.Id, dataUrlOfBytes (getMimeType rel.Target) bytes)
        else
            None)

let private relationshipIdPattern = System.Text.RegularExpressions.Regex("\\s(?:r:[A-Za-z]+|o:relid)=\"([^\"]*)\"")

/// Relationship ids an XML fragment refers to (r:id, r:embed, r:link, r:dm, ..., o:relid).
let relationshipIdsInXml (xml: string) : string list =
    [ for m in relationshipIdPattern.Matches xml -> m.Groups.[1].Value ] |> List.distinct

let private opaqueXmlOfBlocks (blocks: DocumentContent list) : string list =
    blockParagraphs blocks
    |> List.collect (fun para ->
        para.Runs
        |> List.choose (fun run ->
            match run.Content with
            | Image { Opaque = Some opaque } -> Some opaque.Xml
            | _ -> None))

/// Relationship types of parts the writer regenerates (or, for the comment
/// extension parts, that refer into a part it regenerates).
let private isRegeneratedRelationshipType (relType: string) =
    [ "/styles"; "/numbering"; "/settings"; "/header"; "/footer"; "/footnotes"; "/endnotes"; "/comments"
      "/commentsExtended"; "/commentsIds"; "/commentsExtensible" ]
    |> List.exists (fun suffix -> relType.EndsWith suffix)

/// The relationships of a regenerated part that are written back: everything
/// the writer does not produce itself. Pictures and hyperlinks are produced by
/// the writer from the model, so those are kept only when an opaque object of
/// the part refers to them.
let private keptRelationshipsReferenced (baseDir: string) (rels: Relationship list) (referenced: Set<string>) : PreservedRelationship list =
    rels
    |> List.filter (fun rel ->
        if isRegeneratedRelationshipType rel.Type then false
        elif rel.Type.EndsWith "/image" || rel.Type.EndsWith "/hyperlink" then Set.contains rel.Id referenced
        else true)
    |> List.map (fun rel ->
        let isExternal = rel.TargetMode = Some "External"
        { Id = rel.Id
          Type = rel.Type
          Target = if isExternal then rel.Target else "/" + resolveRelationshipTarget baseDir rel.Target
          External = isExternal })

let private keptRelationships (baseDir: string) (rels: Relationship list) (opaqueXml: string list) : PreservedRelationship list =
    keptRelationshipsReferenced baseDir rels (opaqueXml |> List.collect relationshipIdsInXml |> Set.ofList)

/// Every relationship of a part but those of the parts the writer produces
/// itself: what the part, written back as it was read, refers to.
let private sourceRelationships (baseDir: string) (rels: Relationship list) : PreservedRelationship list =
    rels
    |> List.filter (fun rel -> not (isRegeneratedRelationshipType rel.Type))
    |> List.map (fun rel ->
        let isExternal = rel.TargetMode = Some "External"
        { Id = rel.Id
          Type = rel.Type
          Target = if isExternal then rel.Target else "/" + resolveRelationshipTarget baseDir rel.Target
          External = isExternal })

/// Every relationship of a part, as kept relationships (a part that is not
/// written from the model has no relationship the writer makes).
let private allRelationships (baseDir: string) (rels: Relationship list) : PreservedRelationship list =
    rels
    |> List.map (fun rel ->
        let isExternal = rel.TargetMode = Some "External"
        { Id = rel.Id
          Type = rel.Type
          Target = if isExternal then rel.Target else "/" + resolveRelationshipTarget baseDir rel.Target
          External = isExternal })

/// Every part reachable from `roots` through the parts' own relationships,
/// with those relationship parts, in a stable order.
let private reachableParts (source: PartSource) (excluded: Set<string>) (roots: PreservedRelationship list) : PreservedPart list =
    let seen = System.Collections.Generic.HashSet<string>()
    let parts = ResizeArray<PreservedPart>()
    let pending = System.Collections.Generic.Queue<string>()
    for rel in roots do
        if not rel.External then pending.Enqueue (normalizePath rel.Target)
    while pending.Count > 0 do
        let path = pending.Dequeue()
        if not (excluded.Contains path) && seen.Add path then
            match source.Bytes path with
            | Some data ->
                parts.Add { Path = path; Data = data }
                let relsPath = partRelsPath path
                match source.Bytes relsPath, source.Text relsPath with
                | Some relsData, relsText ->
                    if seen.Add relsPath then parts.Add { Path = relsPath; Data = relsData }
                    for rel in relationshipsOfXml relsText do
                        if rel.TargetMode <> Some "External" then
                            pending.Enqueue (resolveRelationshipTarget (parentPath path) rel.Target)
                | _ -> ()
            | None -> ()
    parts |> Seq.sortBy (fun part -> part.Path) |> Seq.toList

let private contentTypeEntries (xml: string option) : (string * string) list * (string * string) list =
    match xml |> Option.map XmlDsl.parse with
    | Some types ->
        let entries (elementName: string) (keyName: string) =
            Xml.children types
            |> List.choose (fun node ->
                if Xml.name node |> Option.map Ns.stripNs = Some elementName then
                    match Xml.attr keyName node, Xml.attr "ContentType" node with
                    | Some key, Some contentType -> Some (key, contentType)
                    | _ -> None
                else None)
        entries "Default" "Extension", entries "Override" "PartName"
    | None -> [], []

/// What is read of a package before the content of its main part: where
/// its parts are, the styles, the theme, the lists, the settings.
type private MainPart = {
    Source: PartSource
    RootRels: Relationship list
    MainRelationship: Relationship option
    MainPath: string
    DocDir: string
    /// The text of the main part.
    Xml: string
    Rels: Relationship list
    StylesPath: string
    SettingsPath: string
    NumberingPath: string
    CommentsPath: string
    FootnotesPath: string
    EndnotesPath: string
    StylesXml: string option
    SettingsXml: string option
    NumberingXml: string option
    CommentsRels: Relationship list
    FootnotesRels: Relationship list
    EndnotesRels: Relationship list
    Styles: StyleDefinition list
    Theme: Theme option
    NumberingDefinitions: NumberingDefinition list
    EvenAndOddHeaders: bool
    Properties: DocumentProperties option
}

/// The parts with text of their own. Headers and footers: (part, its path, its relationships).
type private Stories = {
    HeaderParts: (HeaderFooter * string * Relationship list) list
    FooterParts: (HeaderFooter * string * Relationship list) list
    /// The shapes of those parts, by relationship id.
    PartShapes: Map<string, Shape list>
    Comments: Comment list
    Footnotes: Note list
    Endnotes: Note list
}

/// The main document part is the target of the package's officeDocument
/// relationship (it need not be word/document.xml), and every other part is
/// found through the relationships of the part that refers to it.
let private readMainPart (source: PartSource) : MainPart option =
    let rootRels = source.Text "_rels/.rels" |> relationshipsOfXml
    let mainRelationship = rootRels |> List.tryFind (fun rel -> rel.Type.EndsWith "/officeDocument" && rel.TargetMode <> Some "External")
    let mainPath =
        mainRelationship
        |> Option.map (fun rel -> ZipInterop.resolveTarget "" rel.Target)
        |> Option.filter (fun path -> (source.Text path).IsSome)
        |> Option.defaultValue "word/document.xml"
    let docDir = parentPath mainPath
    source.Text mainPath
    |> Option.map (fun xml ->
        let rels = source.Text (partRelsPath mainPath) |> relationshipsOfXml
        let partPath typeSuffix fallback = partPathOfType docDir rels typeSuffix ((if docDir = "" then "" else docDir + "/") + fallback)

        let stylesPath = partPath "/styles" "styles.xml"
        let settingsPath = partPath "/settings" "settings.xml"
        let stylesXml = source.Text stylesPath
        let settingsXml = source.Text settingsPath
        let themeXml = source.Text (partPath "/theme" "theme/theme1.xml")
        let numberingPath = partPath "/numbering" "numbering.xml"
        let numberingXml = source.Text numberingPath
        let propsXml = source.Text (partPathOfType "" rootRels "/core-properties" "docProps/core.xml")

        // Notes and comments are parts of their own, each with its own relationships.
        let commentsPath = partPath "/comments" "comments.xml"
        let footnotesPath = partPath "/footnotes" "footnotes.xml"
        let endnotesPath = partPath "/endnotes" "endnotes.xml"
        { Source = source
          RootRels = rootRels
          MainRelationship = mainRelationship
          MainPath = mainPath
          DocDir = docDir
          Xml = xml
          Rels = rels
          StylesPath = stylesPath
          SettingsPath = settingsPath
          NumberingPath = numberingPath
          CommentsPath = commentsPath
          FootnotesPath = footnotesPath
          EndnotesPath = endnotesPath
          StylesXml = stylesXml
          SettingsXml = settingsXml
          NumberingXml = numberingXml
          CommentsRels = source.Text (partRelsPath commentsPath) |> relationshipsOfXml
          FootnotesRels = source.Text (partRelsPath footnotesPath) |> relationshipsOfXml
          EndnotesRels = source.Text (partRelsPath endnotesPath) |> relationshipsOfXml
          Styles = stylesXml |> Option.map Styles.parseStyles |> Option.defaultValue []
          Theme = themeXml |> Option.map Theme.parseTheme
          NumberingDefinitions = numberingXml |> Option.map Numbering.parseNumbering |> Option.defaultValue []
          EvenAndOddHeaders = parseSettings settingsXml
          Properties = propsXml |> Option.map Properties.parseProperties })

/// The headers, the footers, the notes and the comments. `headerRefTypes` /
/// `footerRefTypes`: what the sections' references say each part is
/// (relationship id -> "default" | "first" | "even").
let private readStories (main: MainPart) (headerRefTypes: Map<string, string>) (footerRefTypes: Map<string, string>) : Stories =
    let source = main.Source
    // (part, its path, its relationships)
    // The shapes of a header or footer (a text box, a watermark) are modelled as the body's are.
    let partShapes = ResizeArray<string * Shape list>()
    let loadHeaderFooter (refTypes: Map<string, string>) (rel: Relationship) =
        let path = ZipInterop.resolveTarget main.DocDir rel.Target
        let partRels = source.Text (partRelsPath path) |> relationshipsOfXml
        source.Text path
        |> Option.map (fun value ->
            let hfType = refTypes |> Map.tryFind rel.Id |> Option.defaultValue "default"
            let part = HeaderFooter.parseHeaderFooter value hfType (Some rel.Id) |> resolveHeaderFooterHyperlinks partRels
            let holds (name: string) = value.IndexOf(name, System.StringComparison.Ordinal) >= 0
            let part =
                if holds "txbxContent" || holds "prstGeom" || holds "custGeom" || holds "textpath" then
                    let root = XmlDsl.parse value
                    let shapes =
                        Shapes.parseShapesOfPart root
                        |> List.map (fun shape ->
                            { shape with Content = shape.Content |> List.map Para |> completeOpaqueObjects root |> List.choose (function Para para -> Some para | Tbl _ -> None) })
                        |> Shapes.withShapeSourcesOf part.Blocks
                    if not shapes.IsEmpty then partShapes.Add((rel.Id, shapes))
                    { part with
                        Blocks = dropModelledShapeRuns part.Blocks
                        Content = part.Content |> List.map Para |> dropModelledShapeRuns |> List.choose (function Para para -> Some para | Tbl _ -> None) }
                else part
            part, path, partRels)
    let headerParts = main.Rels |> List.filter (fun rel -> rel.Type.EndsWith "/header") |> List.choose (loadHeaderFooter headerRefTypes)
    let footerParts = main.Rels |> List.filter (fun rel -> rel.Type.EndsWith "/footer") |> List.choose (loadHeaderFooter footerRefTypes)
    { HeaderParts = headerParts
      FooterParts = footerParts
      PartShapes = Map.ofSeq partShapes
      Comments =
        source.Text main.CommentsPath |> Option.map Comments.parseComments |> Option.defaultValue [] |> List.map (resolveCommentHyperlinks main.CommentsRels)
      Footnotes =
        source.Text main.FootnotesPath |> Option.map Notes.parseNotes |> Option.defaultValue [] |> List.map (resolveNoteHyperlinks main.FootnotesRels)
      Endnotes =
        source.Text main.EndnotesPath |> Option.map Notes.parseNotes |> Option.defaultValue [] |> List.map (resolveNoteHyperlinks main.EndnotesRels) }

/// Every picture of the package the editor may show: image key (the
/// relationship id; for a part other than the body see partImageKey) -> data URL.
let private allImages (main: MainPart) (stories: Stories) : Map<string, string> =
    let source = main.Source
    let keyedImages key baseDir partRelationships =
        loadImages source baseDir partRelationships
        |> List.map (fun (relId, dataUrl) -> partImageKey key relId, dataUrl)
    (loadImages source main.DocDir main.Rels
     @ ((stories.HeaderParts @ stories.FooterParts)
        |> List.collect (fun (part, path, partRels) ->
            match part.RelationshipId with
            | Some partId -> keyedImages partId (parentPath path) partRels
            | None -> []))
     @ keyedImages footnotesPartKey (parentPath main.FootnotesPath) main.FootnotesRels
     @ keyedImages endnotesPartKey (parentPath main.EndnotesPath) main.EndnotesRels
     @ keyedImages commentsPartKey (parentPath main.CommentsPath) main.CommentsRels)
    |> Map.ofList

/// The relationship ids the opaque objects of blocks refer to.
let private opaqueReferences (blocks: DocumentContent list) : string list =
    opaqueXmlOfBlocks blocks |> List.collect relationshipIdsInXml

/// The document of a package whose parts are read. `content` is the body's
/// blocks (as they are delivered: opaque objects completed, modelled shapes
/// dropped), `documentReferenced` the relationship ids the opaque objects of
/// the body and of its shapes refer to, `bodyChunks` / `bodyChunkCount` what
/// Docx.bodyContentChunks says of the body.
let private assembleDocument (main: MainPart) (stories: Stories) (content: DocumentContent list) (shapes: Shape list) (pageSetup: PageSetup option)
                             (revisions: Revision list) (bodyChunks: int list) (bodyChunkCount: int) (documentReferenced: Set<string>) (images: Map<string, string>) : WordDocument =
    let source = main.Source
    let docDir = main.DocDir
    let rels = main.Rels
    let headerParts = stories.HeaderParts
    let footerParts = stories.FooterParts
    let footnotes = stories.Footnotes
    let endnotes = stories.Endnotes
    let comments = stories.Comments
    // ---- what is kept of the package for the next save ----
    let rootKept =
        keptRelationships "" (main.RootRels |> List.filter (fun rel -> not (rel.Type.EndsWith "/officeDocument"))) []
    let partKept =
        [ yield "document", keptRelationshipsReferenced docDir rels documentReferenced
          for part, path, partRels in headerParts @ footerParts do
              match part.RelationshipId with
              | Some partId -> yield partId, keptRelationships (parentPath path) partRels (opaqueXmlOfBlocks part.Blocks)
              | None -> ()
          yield footnotesPartKey, keptRelationships (parentPath main.FootnotesPath) main.FootnotesRels (footnotes |> List.collect (fun note -> opaqueXmlOfBlocks note.Blocks))
          yield endnotesPartKey, keptRelationships (parentPath main.EndnotesPath) main.EndnotesRels (endnotes |> List.collect (fun note -> opaqueXmlOfBlocks note.Blocks))
          yield commentsPartKey, keptRelationships (parentPath main.CommentsPath) main.CommentsRels (comments |> List.collect (fun comment -> opaqueXmlOfBlocks comment.Blocks))
          // The settings part is patched, not regenerated: all its relationships stay (an attached template).
          yield "settings", keptRelationships (parentPath main.SettingsPath) (source.Text (partRelsPath main.SettingsPath) |> relationshipsOfXml) []
          // The lists and the styles go back as they were read, or with something added: what they
          // refer to (the picture of a picture bullet) stays with them.
          yield "numbering", allRelationships (parentPath main.NumberingPath) (source.Text (partRelsPath main.NumberingPath) |> relationshipsOfXml)
          yield "styles", allRelationships (parentPath main.StylesPath) (source.Text (partRelsPath main.StylesPath) |> relationshipsOfXml) ]
        |> List.filter (fun (_, kept) -> not kept.IsEmpty)
    // Parts the writer produces itself are never carried over as they are.
    let regenerated =
        [ yield main.MainPath
          yield main.StylesPath
          yield main.SettingsPath
          yield main.NumberingPath
          yield main.CommentsPath
          yield main.FootnotesPath
          yield main.EndnotesPath
          for _, path, _ in headerParts @ footerParts do yield path ]
        |> Set.ofList
    let defaults, overrides = contentTypeEntries (source.Text "[Content_Types].xml")
    // ---- what is written back as it was when the editor did not change it ----
    let sourceRels =
        [ yield "document", sourceRelationships docDir rels
          for part, path, partRels in headerParts @ footerParts do
              match part.RelationshipId with
              | Some partId -> yield partId, sourceRelationships (parentPath path) partRels
              | None -> ()
          yield footnotesPartKey, sourceRelationships (parentPath main.FootnotesPath) main.FootnotesRels
          yield endnotesPartKey, sourceRelationships (parentPath main.EndnotesPath) main.EndnotesRels
          yield commentsPartKey, sourceRelationships (parentPath main.CommentsPath) main.CommentsRels ]
        |> List.filter (fun (_, kept) -> not kept.IsEmpty)
    let sourceXml =
        [ yield "document", Some main.Xml
          for part, path, _ in headerParts @ footerParts do
              match part.RelationshipId with
              | Some partId -> yield partId, source.Text path
              | None -> ()
          yield footnotesPartKey, source.Text main.FootnotesPath
          yield endnotesPartKey, source.Text main.EndnotesPath
          yield commentsPartKey, source.Text main.CommentsPath
          yield "numbering", main.NumberingXml ]
        |> List.choose (fun (key, text) -> text |> Option.map (fun value -> key, value))
    let keptParts = reachableParts source regenerated (rootKept @ (partKept |> List.collect snd))
    let keptPaths = keptParts |> List.map (fun part -> part.Path) |> Set.ofList
    let preserved =
        { SourceXml = Map.ofList sourceXml
          SourceRelationships = Map.ofList sourceRels
          SourceParts =
            reachableParts source regenerated (sourceRels |> List.collect snd)
            |> List.filter (fun part -> not (Set.contains part.Path keptPaths))
          BodyContentChunks = bodyChunks
          BodyChunkCount = bodyChunkCount
          PartPaths =
            [ for part, path, _ in headerParts @ footerParts do
                  match part.RelationshipId with
                  | Some partId -> yield partId, path
                  | None -> () ]
            |> Map.ofList
          OriginalText =
            [ yield "[Content_Types].xml"
              yield "_rels/.rels"
              yield partRelsPath main.MainPath
              for _, path, _ in headerParts @ footerParts do yield partRelsPath path
              yield partRelsPath main.FootnotesPath
              yield partRelsPath main.EndnotesPath
              yield partRelsPath main.CommentsPath
              yield partRelsPath main.SettingsPath
              yield partRelsPath main.NumberingPath
              yield partRelsPath main.StylesPath ]
            |> List.distinct
            |> List.choose (fun path -> source.Text path |> Option.map (fun text -> path, text))
            |> Map.ofList
          Parts = keptParts
          RootRelationships = rootKept
          DocumentRelationshipId = main.MainRelationship |> Option.map (fun rel -> rel.Id) |> Option.defaultValue "rId1"
          PartRelationships = Map.ofList partKept
          Defaults = defaults
          Overrides = overrides
          MainContentType = overrides |> List.tryPick (fun (partName, contentType) -> if normalizePath partName = main.MainPath then Some contentType else None)
          StylesXml = main.StylesXml
          SettingsXml = main.SettingsXml }

    {
        Content = content
        Relationships = rels
        Styles = main.Styles
        Theme = main.Theme
        Headers = headerParts |> List.map (fun (part, _, _) -> part)
        Footers = footerParts |> List.map (fun (part, _, _) -> part)
        Properties = main.Properties
        Comments = comments
        Revisions = revisions
        Footnotes = footnotes
        Endnotes = endnotes
        Shapes = shapes
        PartShapes = stories.PartShapes
        Images = images
        PageSetup = pageSetup
        NumberingDefinitions = main.NumberingDefinitions
        Preserved = preserved
    }

/// Parse a DOCX package.
let parseFromSource (source: PartSource) : WordDocument option =
    readMainPart source
    |> Option.map (fun main ->
        // (The main part of a long document is megabytes of XML: it is scanned
        // straight into nodes, and what it holds no element of is not looked for.)
        let doc, docNames = Ns.parseWithLocalNames main.Xml
        let docHolds (name: string) = docNames.Contains name
        let body = Ns.childNs "body" doc

        // (Header and footer references are children of section properties: one walk finds both kinds.)
        let sectionProperties =
            if docHolds "headerReference" || docHolds "footerReference" then Ns.descendantsNs "sectPr" doc else []
        let referenceTypes name =
            sectionProperties
            |> List.collect (fun sectPr -> Xml.children sectPr |> List.filter (fun child -> Xml.name child |> Option.map Ns.stripNs = Some name))
            |> List.choose (fun refNode ->
                Xml.attr "r:id" refNode
                |> Option.map (fun id -> (id, Docx.attr "type" refNode |> Option.defaultValue "default")))
            |> Map.ofList
        let stories = readStories main (referenceTypes "headerReference") (referenceTypes "footerReference")

        // Shapes the body models are text boxes; everything else that is not a
        // picture stays in the text as an opaque object.
        let completedContent =
            body
            |> Option.map (fun bodyNode -> Docx.bodyContentWithFinalSection (Ns.childNs "sectPr" bodyNode) bodyNode)
            |> Option.defaultValue []
            |> completeOpaqueObjects doc
        let content =
            completedContent
            |> dropModelledShapeRuns
            |> (fun blocks -> if docHolds "moveFrom" || docHolds "moveTo" then withMoveNames doc blocks else blocks)

        let pageSetup =
            body
            |> Option.bind (Ns.childNs "sectPr")
            |> parsePageSetup main.EvenAndOddHeaders

        let revisions = Revisions.parseRevisionsWhere docHolds doc
        let shapes =
            // (A modelled shape has text or a geometry: see Docx.isModelledShape.)
            (if docHolds "txbxContent" || docHolds "prstGeom" || docHolds "custGeom" || docHolds "textpath" then Shapes.parseShapes doc else [])
            |> List.map (fun shape ->
                { shape with Content = shape.Content |> List.map Para |> completeOpaqueObjects doc |> List.choose (function Para para -> Some para | Tbl _ -> None) })
            |> Shapes.withShapeSourcesOf completedContent

        let bodyChunks, bodyChunkCount =
            body |> Option.map Docx.bodyContentChunks |> Option.defaultValue ([], 0)
        let documentReferenced =
            opaqueReferences content @ (shapes |> List.collect (fun shape -> opaqueReferences (shape.Content |> List.map Para))) |> Set.ofList
        assembleDocument main stories content shapes pageSetup revisions bodyChunks bodyChunkCount documentReferenced (allImages main stories)
        |> resolveHyperlinks)

// ---- the main part read front first --------------------------------------------
// A long document is megabytes of XML: scanning all of it into nodes and
// making blocks of all of them is most of the time it takes to open one,
// spent before anything is shown. The body is a sequence of child elements
// (a paragraph, a table, a content control ...), each of which gives its
// blocks on its own - apart from two things that span the body, kept across
// elements here (the count of rendered page breaks, and an empty section-break
// paragraph that looks at the block after it). A ProgressiveReader finds the
// body's child elements in the text one after the other and gives the blocks
// of as many as it is asked for; what it has given at the end is what
// `parseFromSource` gives. A main part it cannot read that way (no body-level
// section properties at the end of the body, tracked moves - whose names span
// the part -, a document type declaration) is read at once.

let private textAt (s: string) (at: int) (value: string) : bool =
    if at < 0 || at + value.Length > s.Length then false
    else
        let mutable same = true
        let mutable k = 0
        while same && k < value.Length do
            if s.[at + k] <> value.[k] then same <- false
            k <- k + 1
        same

let private indexOfText (s: string) (value: string) (from: int) : int =
#if FABLE_COMPILER
    s.IndexOf(value, from)
#else
    s.IndexOf(value, from, System.StringComparison.Ordinal)
#endif

let private lastIndexOfText (s: string) (value: string) : int =
#if FABLE_COMPILER
    s.LastIndexOf(value)
#else
    s.LastIndexOf(value, System.StringComparison.Ordinal)
#endif

/// Index of the '>' that ends the tag starting at `lt` (an attribute value may hold one); -1: none.
let private tagEndAt (xml: string) (lt: int) : int =
    let n = xml.Length
    let mutable i = lt + 1
    let mutable inQuote = false
    let mutable quote = '"'
    let mutable found = -1
    while found < 0 && i < n do
        let c = xml.[i]
        if inQuote then
            if c = quote then inQuote <- false
        elif c = '>' then found <- i
        elif c = '"' || c = '\'' then
            inQuote <- true
            quote <- c
        i <- i + 1
    found

let private tagNameAt (xml: string) (lt: int) : string =
    let mutable i = lt + 1
    let mutable fin = false
    while not fin && i < xml.Length do
        let c = xml.[i]
        if c = ' ' || c = '>' || c = '/' || c = '\n' || c = '\t' || c = '\r' then fin <- true else i <- i + 1
    xml.Substring(lt + 1, i - lt - 1)

/// The next tag of an element's content from `pos` on that is not a comment,
/// a CDATA section or a processing instruction: where its '<' is (-1: the
/// text ends, or is not well formed).
let private nextTagAt (xml: string) (pos: int) : int =
    let n = xml.Length
    let mutable at = pos
    let mutable found = -2
    while found = -2 do
        let lt = if at >= 0 && at < n then xml.IndexOf('<', at) else -1
        if lt < 0 || lt + 1 >= n then found <- -1
        elif xml.[lt + 1] = '!' then
            let terminator = if textAt xml lt "<!--" then "-->" elif textAt xml lt "<![CDATA[" then "]]>" else ""
            if terminator = "" then found <- -1
            else
                let fin = indexOfText xml terminator lt
                if fin < 0 then found <- -1 else at <- fin + terminator.Length
        elif xml.[lt + 1] = '?' then
            let fin = indexOfText xml "?>" lt
            if fin < 0 then found <- -1 else at <- fin + 2
        else found <- lt
    found

/// Index just after the element whose start tag begins at `lt`; -1 when the text is not well formed.
let private elementEndAt (xml: string) (lt: int) : int =
    let gt = tagEndAt xml lt
    if gt < 0 then -1
    elif xml.[gt - 1] = '/' then gt + 1
    else
        let mutable depth = 1
        let mutable pos = gt + 1
        while depth > 0 && pos >= 0 do
            let next = nextTagAt xml pos
            if next < 0 then pos <- -1
            elif xml.[next + 1] = '/' then
                let fin = xml.IndexOf('>', next)
                if fin < 0 then pos <- -1
                else
                    depth <- depth - 1
                    pos <- fin + 1
            else
                let fin = tagEndAt xml next
                if fin < 0 then pos <- -1
                else
                    if xml.[fin - 1] <> '/' then depth <- depth + 1
                    pos <- fin + 1
        pos

/// Where the child elements of the body are in the text of the main part,
/// found one at a time.
type private BodyCursor = {
    Xml: string
    /// The root element with its attributes (the namespaces of the part) and no content.
    Root: XmlNode
    /// The root's start tag and its name: what a child element is wrapped in for the platform's parser.
    RootTag: string
    RootName: string
    /// Where the body's end tag starts.
    BodyClose: int
    /// Where the next child element is looked for.
    mutable Pos: int
    /// How many child elements were found.
    mutable Count: int
    mutable AtEnd: bool
}

/// The cursor at the first child element of the body; None for a main part
/// that is read at once (a document type declaration, no body, an empty one).
let private openBodyCursor (xml: string) : BodyCursor option =
    // The root element: after the declaration and comments (a document type declaration is not read here).
    let rootAt = nextTagAt xml 0
    if rootAt < 0 || xml.[rootAt + 1] = '/' then None
    else
        let rootEnd = tagEndAt xml rootAt
        if rootEnd < 0 || xml.[rootEnd - 1] = '/' then None
        else
            let rootName = tagNameAt xml rootAt
            let rootTag = xml.Substring(rootAt, rootEnd + 1 - rootAt)
            match XmlDsl.tryScan (rootTag.Substring(0, rootTag.Length - 1) + "/>") (System.Collections.Generic.HashSet<string>()) with
            | None -> None
            | Some root ->
                // The first child of the root that is a body.
                let mutable pos = rootEnd + 1
                let mutable result = None
                let mutable searching = true
                while searching do
                    let lt = nextTagAt xml pos
                    if lt < 0 || xml.[lt + 1] = '/' then searching <- false
                    else
                        let name = tagNameAt xml lt
                        if Ns.stripNs name = "body" then
                            searching <- false
                            let gt = tagEndAt xml lt
                            let close = lastIndexOfText xml ("</" + name)
                            if gt >= 0 && xml.[gt - 1] <> '/' && close > gt then
                                result <- Some { Xml = xml; Root = root; RootTag = rootTag; RootName = rootName; BodyClose = close; Pos = gt + 1; Count = 0; AtEnd = false }
                        else
                            let fin = elementEndAt xml lt
                            if fin < 0 then searching <- false else pos <- fin
                result

/// The next child element of the body: (where it starts, where it ends, its
/// name); None at the end of the body. (Text that is not well formed ends the body.)
let private nextBodyChild (cursor: BodyCursor) : (int * int * string) option =
    if cursor.AtEnd then None
    else
        let lt = nextTagAt cursor.Xml cursor.Pos
        if lt < 0 || cursor.Xml.[lt + 1] = '/' then
            cursor.AtEnd <- true
            None
        else
            let fin = elementEndAt cursor.Xml lt
            // (An element that does not end, or ends after the body does: the tags do not match.)
            if fin < 0 || fin > cursor.BodyClose then
                failwith "The main part of the document is not well formed"
            else
                cursor.Pos <- fin
                cursor.Count <- cursor.Count + 1
                Some (lt, fin, tagNameAt cursor.Xml lt)

/// The last child element of the body when that is its section properties:
/// (where it starts, where it ends).
let private finalSectionPropertiesAt (cursor: BodyCursor) : (int * int) option =
    let xml = cursor.Xml
    let isSpace (c: char) = c = ' ' || c = '\n' || c = '\t' || c = '\r'
    let mutable fin = cursor.BodyClose
    while fin > 0 && isSpace xml.[fin - 1] do fin <- fin - 1
    if fin < 2 || xml.[fin - 1] <> '>' then None
    else
        // The '<' of the tag that ends there (an attribute value holds none).
        let mutable lt = fin - 1
        while lt >= 0 && xml.[lt] <> '<' do lt <- lt - 1
        if lt < 0 then None
        elif xml.[fin - 2] = '/' then
            if xml.[lt + 1] <> '/' && Ns.stripNs (tagNameAt xml lt) = "sectPr" then Some (lt, fin) else None
        elif xml.[lt + 1] <> '/' then None
        else
            let name = xml.Substring(lt + 2, fin - 1 - (lt + 2)).Trim()
            if Ns.stripNs name <> "sectPr" then None
            else
                // Back to the start tag this end tag closes (section properties may hold
                // earlier ones of the same name, in a tracked change).
                let opening = "<" + name
                let closing = "</" + name
                let mutable depth = 1
                let mutable at = lt - 1
                let mutable start = -1
                while start < 0 && at >= 0 do
                    if xml.[at] = '<' then
                        if textAt xml at closing then depth <- depth + 1
                        elif textAt xml at opening
                             && (let after = at + opening.Length
                                 after < xml.Length && (isSpace xml.[after] || xml.[after] = '>' || xml.[after] = '/')) then
                            let gt = tagEndAt xml at
                            if gt >= 0 && xml.[gt - 1] <> '/' then
                                depth <- depth - 1
                                if depth = 0 then start <- at
                    at <- at - 1
                // (It is a child of the body when the elements before it close up to there:
                // the reader finds out as it goes, see ProgressiveReader.)
                if start < 0 then None else Some (start, fin)

let private referenceTagPattern = System.Text.RegularExpressions.Regex("<(?:[A-Za-z_][A-Za-z0-9_.\\-]*:)?(headerReference|footerReference)[\\s/][^>]*>")

/// What the sections' header / footer references say each part is
/// (relationship id -> type), read off the text of the main part: (headers, footers).
let private referenceTypesInText (xml: string) : Map<string, string> * Map<string, string> =
    let headers = ResizeArray<string * string>()
    let footers = ResizeArray<string * string>()
    if indexOfText xml "Reference" 0 >= 0 then
        let names = System.Collections.Generic.HashSet<string>()
        for m in referenceTagPattern.Matches xml do
            let tag = m.Value
            let closed = if tag.EndsWith "/>" then tag else tag.Substring(0, tag.Length - 1) + "/>"
            match XmlDsl.tryScan closed names with
            | Some node ->
                match Xml.attr "r:id" node with
                | Some id ->
                    let entry = id, (Docx.attr "type" node |> Option.defaultValue "default")
                    if m.Groups.[1].Value = "headerReference" then headers.Add entry else footers.Add entry
                | None -> ()
            | None -> ()
    Map.ofSeq headers, Map.ofSeq footers

/// Blocks a ProgressiveReader has read.
type ProgressiveBatch = {
    /// The index, among the body's blocks, of the first of them.
    First: int
    Blocks: DocumentContent list
    /// The shapes anchored in them (Shape.AnchorBlock is an index among the body's blocks).
    Shapes: Shape list
}

/// A package whose main part is read front first (see above).
type ProgressiveReader = {
    /// The document without its body: the styles, the theme, the lists, the
    /// headers and footers, the notes, the comments, the page setup. (No
    /// Content, Shapes, Images or Preserved: see Finish.)
    Front: WordDocument
    /// Whether the body is read to its end - or as far as it can be read (see Failure).
    IsDone: unit -> bool
    /// Why the reading ended before the end of the body: the text of the main
    /// part is not well formed from some element on, or an element could not
    /// be made blocks of. The blocks before it were given (the last Read gave
    /// those it had); nothing more is read, and Finish fails.
    Failure: unit -> string option
    /// Read on until `count` more blocks are there (fewer at the end of the
    /// body; all the blocks of a main part that is read at once).
    Read: int -> ProgressiveBatch
    /// The page setup of the section the blocks read so far end in: the
    /// section properties that come next in the text - those of a paragraph
    /// that is not read yet, or the body's own (Front.PageSetup) in the last
    /// section. What is read is laid out with it as the page setup of the
    /// document: the pages of a section are then on its paper before the
    /// paragraph that ends the section (and says so) is read.
    OpenSection: unit -> PageSetup option
    /// Whether the main part was read at once (it cannot be read front first).
    AtOnce: bool
    /// How many of the body's blocks are read.
    BlockCount: unit -> int
    /// How far the body is read: 0 .. 1.
    Progress: unit -> float
    /// The pictures blocks show: image key -> data URL, as WordDocument.Images
    /// has them, for the keys the blocks use. (A picture is made a data URL
    /// when a block that shows it is asked for.)
    ImagesOf: DocumentContent list -> Map<string, string>
    /// The document once the body is read (the rest is read first): what
    /// `parseFromSource` gives - Content, Revisions and Images when the reader
    /// was opened to keep them, else empty.
    Finish: unit -> WordDocument
}

/// Image keys of the pictures blocks show.
let private imageKeysOfBlocks (blocks: DocumentContent list) : string list =
    blockParagraphs blocks
    |> List.collect (fun para ->
        para.Runs
        |> List.choose (fun run ->
            match run.Content with
            | Image image -> Some image.RelationshipId
            | _ -> None))

/// A reader of a document that was read at once.
let private wholeReader (doc: WordDocument) : ProgressiveReader =
    let mutable given = false
    { Front = { doc with Content = []; Shapes = []; Images = Map.empty; Revisions = []; Preserved = emptyPackagePreservation }
      IsDone = fun () -> given
      Failure = fun () -> None
      Read =
        fun _ ->
            if given then { First = doc.Content.Length; Blocks = []; Shapes = [] }
            else
                given <- true
                { First = 0; Blocks = doc.Content; Shapes = doc.Shapes }
      OpenSection = fun () -> doc.PageSetup
      AtOnce = true
      BlockCount = fun () -> if given then doc.Content.Length else 0
      Progress = fun () -> if given then 1.0 else 0.0
      ImagesOf = fun _ -> doc.Images
      Finish =
        fun () ->
            given <- true
            doc }

/// Open a package to read its main part front first. `keepContent`: Finish
/// gives the whole document as `parseFromSource` does (the blocks, the
/// revisions and every picture are kept for it); without it the blocks belong
/// to whoever read them. None: not a document.
let openProgressiveFromSource (keepContent: bool) (source: PartSource) : ProgressiveReader option =
    match readMainPart source with
    | None -> None
    | Some main ->
        let xml = main.Xml
        let cursorAndFinal =
            // (The names of tracked moves are found by following their range markers through the part.)
            if indexOfText xml "moveFromRangeStart" 0 >= 0 || indexOfText xml "moveToRangeStart" 0 >= 0 then None
            else
                openBodyCursor xml
                |> Option.bind (fun cursor -> finalSectionPropertiesAt cursor |> Option.map (fun final -> cursor, final))
        let names = System.Collections.Generic.HashSet<string>()
        /// One element of the main part as a node.
        let parseElement (cursor: BodyCursor) (text: string) : XmlNode =
            match XmlDsl.tryScan text names with
            | Some node -> node
            | None ->
                // (What the scanner does not read: by the platform's parser, in the root's namespaces.)
                match XmlDsl.parse (cursor.RootTag + text + "</" + cursor.RootName + ">") with
                | XmlDsl.Element (_, _, children) -> children |> List.tryFind (fun child -> (Xml.name child).IsSome) |> Option.defaultValue XmlDsl.Empty
                | _ -> XmlDsl.Empty
        let finalSetup =
            cursorAndFinal
            |> Option.bind (fun (cursor, (start, fin)) ->
                try Some (cursor, start, parsePageSetup main.EvenAndOddHeaders (Some (parseElement cursor (xml.Substring(start, fin - start)))))
                with _ -> None)
        match finalSetup with
        | None -> parseFromSource source |> Option.map wholeReader
        | Some (cursor, finalStart, finalPageSetup) ->
            let headerRefTypes, footerRefTypes = referenceTypesInText xml
            let stories = readStories main headerRefTypes footerRefTypes
            let hyperlinks = hyperlinkTargets main.Rels
            let renderedBreakCounter = ref 0
            // The last block made, not given yet: what it is depends on the block after it.
            let mutable held : (DocumentContent * Shape list * int * int) option = None
            // Where the element of the last block given ends in the text.
            let mutable givenTo = cursor.Pos
            // The section properties that come next in the text after what is given: where their
            // start tag is (-1: none before the body's own; -2: not looked for), and the page setup they give.
            let mutable openSectionAt = -2
            let mutable openSectionSetup : PageSetup option = None
            let isNameChar (c: char) = System.Char.IsLetterOrDigit c || c = '_' || c = '.' || c = '-'
            /// Where the start tag of the next section properties element at or after `from` begins; -1: none.
            let nextSectionProperties (from: int) : int =
                let mutable at = indexOfText xml "sectPr" from
                let mutable found = -1
                while found < 0 && at >= 0 do
                    let after = at + 6
                    // (A start tag: "<sectPr" or "<prefix:sectPr", the name ending there.)
                    let mutable start = at - 1
                    if start >= 0 && xml.[start] = ':' then
                        start <- start - 1
                        while start >= 0 && isNameChar xml.[start] do start <- start - 1
                    if start >= 0 && xml.[start] = '<' && after < xml.Length && (xml.[after] = ' ' || xml.[after] = '>' || xml.[after] = '/' || xml.[after] = '\n' || xml.[after] = '\t' || xml.[after] = '\r') then
                        found <- start
                    else at <- indexOfText xml "sectPr" after
                found
            let openSection () : PageSetup option =
                // (Looked for again when the paragraph that holds them was given.)
                if openSectionAt = -2 || (openSectionAt >= 0 && givenTo > openSectionAt) then
                    let at = nextSectionProperties givenTo
                    if at < 0 || at >= finalStart then
                        openSectionAt <- -1
                        openSectionSetup <- finalPageSetup
                    else
                        openSectionAt <- at
                        openSectionSetup <-
                            (try
                                let fin = elementEndAt xml at
                                if fin < 0 then finalPageSetup
                                else parsePageSetup main.EvenAndOddHeaders (Some (parseElement cursor (xml.Substring(at, fin - at))))
                             with _ -> finalPageSetup)
                openSectionSetup
            let mutable made = 0
            let mutable finished = false
            // Where the first child of the body that is section properties starts (the page setup is read off it).
            let mutable firstSectionProperties = -1
            let blockChunks = ResizeArray<int>()
            let referenced = System.Collections.Generic.HashSet<string>()
            let allShapes = ResizeArray<Shape>()
            let allContent = ResizeArray<DocumentContent>()
            let revisionTypes = [ "ins"; "del"; "rPrChange"; "moveFrom"; "moveTo" ]
            let revisions = revisionTypes |> List.map (fun _ -> ResizeArray<Revision>())
            let mutable finishedDocument : WordDocument option = None
            // Why the body could not be read to its end (see ProgressiveReader.Failure).
            let mutable failure : string option = None

            // ---- pictures, each made a data URL when it is first shown ----
            let imageCandidates = System.Collections.Generic.Dictionary<string, ResizeArray<string * Relationship>>()
            let addImages (key: string -> string) (baseDir: string) (partRels: Relationship list) =
                for rel in partRels do
                    if rel.Type.Contains "/image" then
                        let imageKey = key rel.Id
                        if not (imageCandidates.ContainsKey imageKey) then imageCandidates.[imageKey] <- ResizeArray()
                        imageCandidates.[imageKey].Add ((baseDir, rel))
            addImages id main.DocDir main.Rels
            for part, path, partRels in stories.HeaderParts @ stories.FooterParts do
                match part.RelationshipId with
                | Some partId -> addImages (partImageKey partId) (parentPath path) partRels
                | None -> ()
            addImages (partImageKey footnotesPartKey) (parentPath main.FootnotesPath) main.FootnotesRels
            addImages (partImageKey endnotesPartKey) (parentPath main.EndnotesPath) main.EndnotesRels
            addImages (partImageKey commentsPartKey) (parentPath main.CommentsPath) main.CommentsRels
            let knownImages = System.Collections.Generic.Dictionary<string, string option>()
            let imageOf (imageKey: string) : string option =
                match knownImages.TryGetValue imageKey with
                | true, known -> known
                | false, _ ->
                    let found =
                        match imageCandidates.TryGetValue imageKey with
                        | true, candidates ->
                            // (The last relationship of an id that has a picture, as a map made of all of them keeps.)
                            let mutable result = None
                            let mutable i = candidates.Count - 1
                            while result.IsNone && i >= 0 do
                                let baseDir, rel = candidates.[i]
                                result <-
                                    if rel.TargetMode = Some "External" then Some rel.Target
                                    else
                                        main.Source.Bytes (resolveRelationshipTarget baseDir rel.Target)
                                        |> Option.map (fun bytes -> dataUrlOfBytes (getMimeType rel.Target) bytes)
                                i <- i - 1
                            result
                        | false, _ -> None
                    knownImages.[imageKey] <- found
                    found
            let imagesOf (blocks: DocumentContent list) : Map<string, string> =
                let mutable images = Map.empty
                for imageKey in imageKeysOfBlocks blocks do
                    if not (Map.containsKey imageKey images) then
                        match imageOf imageKey with
                        | Some source -> images <- Map.add imageKey source images
                        | None -> ()
                images

            /// A block as it is delivered, noted for the end.
            let deliver (released: ResizeArray<DocumentContent>) (shapesOut: ResizeArray<Shape>) (raw: DocumentContent) (shapes: Shape list) (chunk: int) (chunkEnd: int) =
                givenTo <- chunkEnd
                let completed = [ raw ] |> completeOpaqueObjects cursor.Root
                let shapes = Shapes.withShapeSources completed shapes
                let block =
                    completed
                    |> dropModelledShapeRuns
                    |> mapBlockRuns (resolveHyperlinkRun hyperlinks)
                for id in opaqueReferences block do referenced.Add id |> ignore
                for shape in shapes do
                    for id in opaqueReferences (shape.Content |> List.map Para) do referenced.Add id |> ignore
                blockChunks.Add chunk
                allShapes.AddRange shapes
                shapesOut.AddRange shapes
                for item in block do
                    released.Add item
                    if keepContent then allContent.Add item

            let readWhile (more: int -> bool) : ProgressiveBatch =
                let released = ResizeArray<DocumentContent>()
                let shapesOut = ResizeArray<Shape>()
                let first = blockChunks.Count
                while not finished && more released.Count do
                  try
                      match nextBodyChild cursor with
                      | None ->
                          finished <- true
                          match held with
                          | Some (raw, shapes, chunk, chunkEnd) ->
                              held <- None
                              deliver released shapesOut raw shapes chunk chunkEnd
                          | None -> ()
                      | Some (start, fin, name) ->
                          let chunk = cursor.Count - 1
                          let localName = Ns.stripNs name
                          if localName = "sectPr" && firstSectionProperties < 0 then firstSectionProperties <- start
                          let givesBlocks = localName = "p" || localName = "tbl" || localName = "sdt" || localName = "customXml"
                          if givesBlocks || keepContent then
                              let text = xml.Substring(start, fin - start)
                              let node = parseElement cursor text
                              if keepContent then
                                  List.iter2 (fun revisionType (found: ResizeArray<Revision>) ->
                                      if indexOfText text revisionType 0 >= 0 then
                                          found.AddRange (Ns.descendantsNs revisionType node |> List.choose Revisions.parseRevision)) revisionTypes revisions
                              if givesBlocks then
                                  let blocks = Docx.contentOfChild renderedBreakCounter node
                                  // (A modelled shape has text or a geometry: see Docx.isModelledShape.)
                                  let blockNodes =
                                      if indexOfText text "txbxContent" 0 >= 0 || indexOfText text "prstGeom" 0 >= 0 || indexOfText text "custGeom" 0 >= 0 || indexOfText text "textpath" 0 >= 0 then
                                          Shapes.blockNodesOfChild node |> List.toArray
                                      else [||]
                                  blocks
                                  |> List.iteri (fun i raw ->
                                      let shapes =
                                          if i < blockNodes.Length then
                                              Shapes.shapesOfBlock (made + i) blockNodes.[i]
                                              |> List.map (fun shape ->
                                                  { shape with Content = shape.Content |> List.map Para |> completeOpaqueObjects cursor.Root |> List.choose (function Para para -> Some para | Tbl _ -> None) })
                                          else []
                                      match held with
                                      | Some (before, beforeShapes, beforeChunk, beforeEnd) -> deliver released shapesOut (Docx.normalizedSectionBreak before (Some raw)) beforeShapes beforeChunk beforeEnd
                                      | None -> ()
                                      held <- Some (raw, shapes, chunk, fin))
                                  made <- made + blocks.Length
                  with error ->
                    // The element cannot be read: the blocks before it are what there is.
                    failure <- Some error.Message
                    finished <- true
                    match held with
                    | Some (raw, shapes, chunk, chunkEnd) ->
                        held <- None
                        (try deliver released shapesOut raw shapes chunk chunkEnd with _ -> ())
                    | None -> ()
                { First = first; Blocks = List.ofSeq released; Shapes = List.ofSeq shapesOut }
            let read (count: int) : ProgressiveBatch = readWhile (fun given -> given < count)

            let finish () : WordDocument =
                match finishedDocument, failure with
                | Some doc, _ -> doc
                | None, Some message -> failwith message
                | None, None ->
                    while not finished do read System.Int32.MaxValue |> ignore
                    // (The first child of the body that is section properties: the last one, but in a
                    // body that has more than one.)
                    let pageSetup =
                        if firstSectionProperties >= 0 && firstSectionProperties <> finalStart then
                            let fin = elementEndAt xml firstSectionProperties
                            parsePageSetup main.EvenAndOddHeaders (Some (parseElement cursor (xml.Substring(firstSectionProperties, fin - firstSectionProperties))))
                        else finalPageSetup
                    let doc =
                        assembleDocument main stories (List.ofSeq allContent) (List.ofSeq allShapes) pageSetup
                            (revisions |> List.collect List.ofSeq) (List.ofSeq blockChunks) cursor.Count (Set.ofSeq referenced)
                            (if keepContent then allImages main stories else Map.empty)
                    finishedDocument <- Some doc
                    doc

            Some
                { Front =
                    { Content = []
                      Relationships = main.Rels
                      Styles = main.Styles
                      Theme = main.Theme
                      Headers = stories.HeaderParts |> List.map (fun (part, _, _) -> part)
                      Footers = stories.FooterParts |> List.map (fun (part, _, _) -> part)
                      Properties = main.Properties
                      Comments = stories.Comments
                      Revisions = []
                      Footnotes = stories.Footnotes
                      Endnotes = stories.Endnotes
                      Shapes = []
                      PartShapes = stories.PartShapes
                      Images = Map.empty
                      PageSetup = finalPageSetup
                      NumberingDefinitions = main.NumberingDefinitions
                      Preserved = emptyPackagePreservation }
                  IsDone = fun () -> finished
                  Failure = fun () -> failure
                  Read = read
                  OpenSection = (fun () -> if finished then finalPageSetup else openSection ())
                  AtOnce = false
                  BlockCount = fun () -> blockChunks.Count
                  Progress = fun () -> if finished then 1.0 else float (cursor.Pos) / float (max 1 cursor.BodyClose)
                  ImagesOf = imagesOf
                  Finish = finish }

#if FABLE_COMPILER
[<Fable.Core.Emit("new Uint8Array($0)")>]
let private bytesOfArrayBuffer (buffer: obj) : byte[] = Fable.Core.Util.jsNative

[<Fable.Core.Emit("new TextDecoder('utf-8').decode($0)")>]
let private textOfBytes (bytes: byte[]) : string = Fable.Core.Util.jsNative
#else
let private bytesOfArrayBuffer (buffer: obj) : byte[] = unbox buffer
let private textOfBytes (bytes: byte[]) : string = System.Text.Encoding.UTF8.GetString bytes
#endif

let private lookupPart (parts: Map<string, 'T>) (path: string) : 'T option =
    match Map.tryFind path parts with
    | Some value -> Some value
    | None ->
        // OPC part names are case-insensitive.
        let wanted = path.ToLowerInvariant()
        parts |> Map.tryPick (fun name value -> if name.ToLowerInvariant() = wanted then Some value else None)

/// The parts of an archive, every entry read (inflated) once, up front: the
/// parser follows relationships from part to part, which a promise per part
/// would turn inside out.
let sourceOfArchive (archive: ZipArchive) : Fable.Core.JS.Promise<PartSource> =
    promise {
        let! entries =
            archive.Files
            |> Map.toArray
            |> Array.map (fun (name, file) ->
                promise {
                    let! buffer = file.GetArrayBuffer()
                    return name, bytesOfArrayBuffer buffer
                })
            |> Promise.all
        let bytes = Map.ofArray entries
        let texts = System.Collections.Generic.Dictionary<string, string option>()
        let source =
            { Bytes = lookupPart bytes
              Text =
                fun path ->
                    match texts.TryGetValue path with
                    | true, text -> text
                    | false, _ ->
                        let text = lookupPart bytes path |> Option.map textOfBytes
                        texts.[path] <- text
                        text }
        return source
    }

// Parse DOCX archive
let parseDocument (archive: ZipArchive) : Fable.Core.JS.Promise<WordDocument option> =
    promise {
        let! source = sourceOfArchive archive
        return parseFromSource source
    }

let parseDocumentNative (archive: ZipArchive) : WordDocument option =
    parseFromSource { Text = getFileTextNative archive; Bytes = getFileBytesNative archive }

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

/// A package opened to read its main part front first (see ProgressiveReader).
let openProgressiveNative (keepContent: bool) (data: byte[]) : ProgressiveReader option =
    let archive = loadArchiveNative data
    openProgressiveFromSource keepContent { Text = getFileTextNative archive; Bytes = getFileBytesNative archive }

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

/// `props` are the effective paragraph properties, `direct` what the paragraph
/// itself states (its own w:pPr). Next to each effective value the object
/// carries ExplicitX: the value when the paragraph states it, null when it is
/// inherited (from a style, the document defaults or the numbering level) or
/// not set anywhere. The writer writes the explicit ones only, so a save does
/// not turn inherited or unset formatting into direct formatting.
let private paragraphLayoutProps fontSize defaultMissingSpacingToZero (borderAdvanceSource: ParagraphBorders option) (direct: ParagraphProperties) (props: ParagraphProperties) =
    let missingSpacingFallback = if defaultMissingSpacingToZero then Some 0.0 else None
    let spacingBefore =
        if props.SpacingBeforeAuto = Some true then Some wordAutoSpacing else wordSpacingTwipsToPoints props.SpacingBefore |> Option.orElse missingSpacingFallback
    let spacingAfter =
        if props.SpacingAfterAuto = Some true then Some wordAutoSpacing
        else wordSpacingTwipsToPoints props.SpacingAfter |> Option.orElse missingSpacingFallback
    let firstLineIndent =
        // (Stated together, the hanging indent is the one that counts.)
        match wordTwipsToPoints props.FirstLineIndent, wordTwipsToPoints props.HangingIndent with
        | _, Some hanging -> Some -hanging
        | Some firstLine, None -> Some firstLine
        | None, None -> None
    let lineSpacing = wordLineSpacingToMultiple fontSize props.LineSpacing props.LineRule
    let whenDirect (stated: bool) (value: 'a option) = if stated then value else None
    [
        "ExplicitParagraphFormat" ==> true
        "ExplicitSpacingBefore" ==> whenDirect (direct.SpacingBefore.IsSome || direct.SpacingBeforeAuto.IsSome) spacingBefore
        "ExplicitSpacingAfter" ==> whenDirect (direct.SpacingAfter.IsSome || direct.SpacingAfterAuto.IsSome) spacingAfter
        "ExplicitLineSpacing" ==> whenDirect direct.LineSpacing.IsSome lineSpacing
        "ExplicitIndentLeft" ==> whenDirect direct.IndentLeft.IsSome (wordTwipsToPoints props.IndentLeft)
        "ExplicitIndentRight" ==> whenDirect direct.IndentRight.IsSome (wordTwipsToPoints props.IndentRight)
        "ExplicitFirstLineIndent" ==> whenDirect (direct.FirstLineIndent.IsSome || direct.HangingIndent.IsSome) firstLineIndent
        "ExplicitContextualSpacing" ==> direct.ContextualSpacing
        "ExplicitBackgroundColor" ==> whenDirect direct.BackgroundColor.IsSome (wordColorToCss props.BackgroundColor)
        "ExplicitTabs" ==> not direct.TabStops.IsEmpty
        // The paragraph mark has run properties of its own (w:pPr/w:rPr), and the size it states.
        "HasMarkProperties" ==> direct.ParagraphRunProperties.IsSome
        "ExplicitMarkFontSize" ==> (direct.ParagraphRunProperties |> Option.bind (fun mark -> mark.FontSize |> Option.map (fun _ -> wordFontSizeToPoints mark.FontSize)))
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
        "SpacingBefore" ==> spacingBefore
        "SpacingAfter" ==> spacingAfter
        "LineSpacing" ==> lineSpacing
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
    || props.SectionPageNumberFormat.IsSome
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
    // (w:firstLine and w:hanging are the two forms of one property: a paragraph
    // that states either does not take the other from its style. A style's
    // first-line indent used to win over the paragraph's hanging indent.)
    FirstLineIndent = if props.FirstLineIndent.IsSome || props.HangingIndent.IsSome then props.FirstLineIndent else baseProps.FirstLineIndent
    HangingIndent = if props.FirstLineIndent.IsSome || props.HangingIndent.IsSome then props.HangingIndent else baseProps.HangingIndent
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
    PageBreakBeforeDerived = props.PageBreakBeforeDerived
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
    SectionPageNumberFormat = props.SectionPageNumberFormat
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

/// What `resolveStyle` and `defaultParagraphStyleId` found for the style map
/// in hand, by style id: they are asked for every paragraph and every run of
/// a document, and a style's resolution (its chain of based-on styles merged)
/// depends on the map alone.
let mutable private resolvedStylesFor : obj = null
let private resolvedStyles = System.Collections.Generic.Dictionary<string, RunProperties * ParagraphProperties option>()
let mutable private defaultParagraphStyleFor : obj = null
let mutable private defaultParagraphStyleKnown : string option = None

let private resolveStyleUncached (styleMap: Map<string, StyleDefinition>) styleId : RunProperties * ParagraphProperties option =
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

let private resolveStyle (styleMap: Map<string, StyleDefinition>) (styleId: string option) : RunProperties * ParagraphProperties option =
    match styleId with
    | None -> (emptyRunProperties, None)
    | Some id ->
        if not (obj.ReferenceEquals(resolvedStylesFor, styleMap)) then
            resolvedStyles.Clear()
            resolvedStylesFor <- box styleMap
        match resolvedStyles.TryGetValue id with
        | true, known -> known
        | false, _ ->
            let resolved = resolveStyleUncached styleMap styleId
            resolvedStyles.[id] <- resolved
            resolved

let private defaultParagraphStyleId (styleMap: Map<string, StyleDefinition>) =
    if obj.ReferenceEquals(defaultParagraphStyleFor, styleMap) then defaultParagraphStyleKnown
    else
        let found =
            styleMap
            |> Map.toSeq
            |> Seq.tryPick (fun (id, style) ->
                if style.StyleType = "paragraph" && style.IsDefault then Some id else None)
        defaultParagraphStyleFor <- box styleMap
        defaultParagraphStyleKnown <- found
        found

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
        // w:color w:val="auto" on the run: no colour of its own, but it switches a
        // style's colour (a table style's, say) off.
        "ExplicitColorAuto" ==> ((props.Color |> Option.orElse mark.Color) = Some "auto")
    ]

/// The paragraph properties its styles give without the paragraph's direct
/// keepNext / keepLines (only resolved again when it has any).
let private styleKeepProperties (styleMap: Map<string, StyleDefinition>) (para: Paragraph) (effectiveParaProps: ParagraphProperties) =
    if para.Properties.KeepNext.IsNone && para.Properties.KeepLines.IsNone then effectiveParaProps
    else resolvedParagraphProperties styleMap { para.Properties with KeepNext = None; KeepLines = None }

/// Whether a hyperlink target may become an active link: http, https, mailto,
/// tel, or an anchor inside the document ("#bookmark"). A document can carry any
/// string as a target (javascript:alert(1), file:///..., data:...); everything
/// not on this list is delivered as an inactive target: kept so the document
/// saves as it was, never offered as something to follow.
let isActiveHyperlinkTarget (target: string) =
    if String.IsNullOrWhiteSpace target then false
    else
        let lower = target.ToLowerInvariant()
        lower.StartsWith "http://" || lower.StartsWith "https://" || lower.StartsWith "mailto:" || lower.StartsWith "tel:"
        || (lower.StartsWith "#" && lower.Length > 1)

/// A picture source that is an address the document names rather than data
/// from the package: a linked picture, which is not loaded unless the host allows it.
let isExternalImageSource (source: string) =
    not (isNull source) && source.Contains "://" && not (source.StartsWith "data:")

/// Source prefix of an opaque object that has no picture to show: the painter
/// draws a placeholder frame for "opaque:<kind>".
let opaqueSourcePrefix = "opaque:"

/// What an image run shows: the picture behind its relationship; for an opaque
/// object its fallback picture when the browser can draw it, else a placeholder.
let editorImageSource (images: Map<string, string>) (img: ImageData) : string =
    match img.Opaque with
    | Some opaque ->
        match Map.tryFind img.RelationshipId images with
        | Some (source: string) when source.StartsWith "data:image/" && not (source.StartsWith "data:image/x-emf") && not (source.StartsWith "data:image/x-wmf") && not (source.StartsWith "data:image/tiff") -> source
        | _ -> opaqueSourcePrefix + opaque.Kind
    | None ->
        images
        |> Map.tryFind img.RelationshipId
        |> Option.defaultValue img.RelationshipId  // Fallback to relId if not found

/// The colour a comment's mark is shown in when its run has none of its own.
/// (Of the editor's display: the writer does not put it into the file.)
let commentMarkDisplayColor = "#7c3aed"

/// The formatting of the run of a reference mark (a comment's, a note's): what
/// its own w:rPr and its styles give, as for a run of text - the mark made
/// bold, or given a size, is that when the document is read again.
let private referenceMarkFormat (styleMap: Map<string, StyleDefinition>) (theme: Theme option) (paraProps: ParagraphProperties) (run: Run) =
    let props = resolvedRunProperties styleMap paraProps run.Properties
    [ "Style" ==> run.Properties.Style
      "Bold" ==> effectiveOnOff props.Bold
      "Italic" ==> effectiveOnOff props.Italic
      "Underline" ==> effectiveOnOff props.Underline
      "UnderlineStyle" ==> props.UnderlineStyle
      "UnderlineColor" ==> props.UnderlineColor
      "AllCaps" ==> effectiveOnOff props.AllCaps
      "SmallCaps" ==> effectiveOnOff props.SmallCaps
      "Strike" ==> effectiveOnOff props.Strike
      "DoubleStrike" ==> effectiveOnOff props.DoubleStrike
      "Hidden" ==> effectiveOnOff props.Hidden
      "FontSize" ==> wordFontSizeToPoints props.FontSize
      "FontFamily" ==> wordFontFamilyToCss theme props
      "Highlight" ==> props.Highlight
      "IsImage" ==> false ]
    @ explicitRunFields styleMap paraProps run

let rec private toEditorParagraphsOfStyleMap (styleMap: Map<string, StyleDefinition>) (theme: Theme option) (content: DocumentContent list) (images: Map<string, string>) (noteBodies: Map<string * string, string>) : obj array =
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
                            "InsertedRevisionId" ==> (run.InsertedRevision |> Option.bind (fun mark -> mark.Id))
                            "InsertedRevisionAuthor" ==> (run.InsertedRevision |> Option.bind (fun mark -> mark.Author))
                            "InsertedRevisionDate" ==> (run.InsertedRevision |> Option.bind (fun mark -> mark.Date))
                            "RevisionMoveName" ==> run.MoveName
                        ] @ bookmarkRunProps run))
                    | Image img ->
                        // Image run - resolve relationship ID to data URL
                        let imageSrc = editorImageSource images img
                        Some (objFromFields [
                            "IsImage" ==> true
                            "ImageSrc" ==> imageSrc
                            "ImageLinked" ==> (img.Linked || isExternalImageSource imageSrc)
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
                            "ImagePositionHAlign" ==> img.PositionHAlign
                            "ImagePositionVAlign" ==> img.PositionVAlign
                            "ImageWrapText" ==> img.WrapText
                            "ImageWrapPolygon" ==> (if img.WrapPolygon.IsEmpty then null else box (img.WrapPolygon |> List.collect (fun (x, y) -> [ x; y ]) |> List.toArray))
                            "OpaqueXml" ==> (img.Opaque |> Option.map (fun opaque -> opaque.Xml))
                            "OpaqueKind" ==> (img.Opaque |> Option.map (fun opaque -> opaque.Kind))
                            "OpaqueParagraphLevel" ==> (img.Opaque |> Option.map (fun opaque -> opaque.ParagraphLevel))
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
                            "Hyperlink" ==> (link.Target |> Option.filter isActiveHyperlinkTarget)
                            "InactiveHyperlink" ==> (link.Target |> Option.filter (fun target -> not (isActiveHyperlinkTarget target)))
                        ] @ bookmarkRunProps run))
                        | CommentReference id ->
                            Some (objFromFields ([
                                "Text" ==> (sprintf "[Comment %s]" id)
                                "Color" ==> (wordColorToCss effectiveRunProps.Color |> Option.orElse (Some commentMarkDisplayColor))
                                "VertAlign" ==> effectiveRunProps.VertAlign
                                "CommentId" ==> Some id
                            ] @ referenceMarkFormat styleMap theme effectiveParaProps run))
                        | NoteRef note ->
                            // Render the reference as its superscript mark number
                            // (Word writes note ids sequentially) in the normal
                            // text colour, matching the native/PDF path.
                            Some (objFromFields ([
                                "Text" ==> noteMark note.NoteType note.Id
                                "Color" ==> wordColorToCss effectiveRunProps.Color
                                "VertAlign" ==> (effectiveRunProps.VertAlign |> Option.orElse (Some "superscript"))
                                "NoteId" ==> Some note.Id
                                "NoteType" ==> Some note.NoteType
                                "NoteText" ==> (noteBodies |> Map.tryFind (note.NoteType, note.Id))
                            ] @ referenceMarkFormat styleMap theme effectiveParaProps run))
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
                                "InsertedRevisionId" ==> (run.InsertedRevision |> Option.bind (fun mark -> mark.Id))
                                "InsertedRevisionAuthor" ==> (run.InsertedRevision |> Option.bind (fun mark -> mark.Author))
                                "InsertedRevisionDate" ==> (run.InsertedRevision |> Option.bind (fun mark -> mark.Date))
                                "RevisionMoveName" ==> run.MoveName
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
                // Direct keepNext / keepLines are the paragraph's own; StyleKeepX is what
                // its style says without them (the editor infers style definitions from it).
                "KeepNext" ==> para.Properties.KeepNext
                "KeepLines" ==> para.Properties.KeepLines
                "StyleKeepNext" ==> (styleKeepProperties styleMap para effectiveParaProps).KeepNext
                "StyleKeepLines" ==> (styleKeepProperties styleMap para effectiveParaProps).KeepLines
                "StylePageBreakBefore" ==> effectiveParaProps.PageBreakBefore
                "PageBreakBefore" ==> effectiveParaProps.PageBreakBefore
                // The break is the style's, or the reader's own (see PageBreakBeforeDerived): not the paragraph's.
                "PageBreakBeforeDerived" ==> (effectiveParaProps.PageBreakBefore = Some true && (para.Properties.PageBreakBefore <> Some true || para.Properties.PageBreakBeforeDerived))
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
                "SectionPageNumberFormat" ==> effectiveParaProps.SectionPageNumberFormat
                "HeaderReferences" ==> referenceObjects effectiveParaProps.HeaderReferences
                "FooterReferences" ==> referenceObjects effectiveParaProps.FooterReferences
                "ParagraphMarkFontSize" ==> paragraphFontSize
                "MarkRevisionType" ==> (para.MarkRevision |> Option.map fst)
                "MarkRevisionId" ==> (para.MarkRevision |> Option.bind (fun (_, mark) -> mark.Id))
                "MarkRevisionAuthor" ==> (para.MarkRevision |> Option.bind (fun (_, mark) -> mark.Author))
                "MarkRevisionDate" ==> (para.MarkRevision |> Option.bind (fun (_, mark) -> mark.Date))
            ] @ paragraphLayoutProps paragraphFontSize defaultMissingSpacingToZero effectiveParaProps.ParagraphBorders para.Properties layoutParaProps))
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
                                "InsertedRevisionId" ==> (run.InsertedRevision |> Option.bind (fun mark -> mark.Id))
                                "InsertedRevisionAuthor" ==> (run.InsertedRevision |> Option.bind (fun mark -> mark.Author))
                                "InsertedRevisionDate" ==> (run.InsertedRevision |> Option.bind (fun mark -> mark.Date))
                                "RevisionMoveName" ==> run.MoveName
                            ] @ bookmarkRunProps run))
                        | Image img ->
                            let imageSrc = editorImageSource images img
                            Some (objFromFields [
                                "IsImage" ==> true
                                "ImageSrc" ==> imageSrc
                                "ImageLinked" ==> (img.Linked || isExternalImageSource imageSrc)
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
                                "ImagePositionHAlign" ==> img.PositionHAlign
                                "ImagePositionVAlign" ==> img.PositionVAlign
                                "ImageWrapText" ==> img.WrapText
                                "ImageWrapPolygon" ==> (if img.WrapPolygon.IsEmpty then null else box (img.WrapPolygon |> List.collect (fun (x, y) -> [ x; y ]) |> List.toArray))
                                "OpaqueXml" ==> (img.Opaque |> Option.map (fun opaque -> opaque.Xml))
                                "OpaqueKind" ==> (img.Opaque |> Option.map (fun opaque -> opaque.Kind))
                                "OpaqueParagraphLevel" ==> (img.Opaque |> Option.map (fun opaque -> opaque.ParagraphLevel))
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
                                "Hyperlink" ==> (link.Target |> Option.filter isActiveHyperlinkTarget)
                                "InactiveHyperlink" ==> (link.Target |> Option.filter (fun target -> not (isActiveHyperlinkTarget target)))
                            ] @ bookmarkRunProps run))
                        | CommentReference id ->
                            Some (objFromFields ([
                                "Text" ==> (sprintf "[Comment %s]" id)
                                "Color" ==> (wordColorToCss effectiveRunProps.Color |> Option.orElse (Some commentMarkDisplayColor))
                                "VertAlign" ==> effectiveRunProps.VertAlign
                                "CommentId" ==> Some id
                            ] @ referenceMarkFormat styleMap theme effectiveParaProps run))
                        | NoteRef note ->
                            // Render the reference as its superscript mark number
                            // (Word writes note ids sequentially) in the normal
                            // text colour, matching the native/PDF path.
                            Some (objFromFields ([
                                "Text" ==> noteMark note.NoteType note.Id
                                "Color" ==> wordColorToCss effectiveRunProps.Color
                                "VertAlign" ==> (effectiveRunProps.VertAlign |> Option.orElse (Some "superscript"))
                                "NoteId" ==> Some note.Id
                                "NoteType" ==> Some note.NoteType
                                "NoteText" ==> (noteBodies |> Map.tryFind (note.NoteType, note.Id))
                            ] @ referenceMarkFormat styleMap theme effectiveParaProps run))
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
                                "InsertedRevisionId" ==> (run.InsertedRevision |> Option.bind (fun mark -> mark.Id))
                                "InsertedRevisionAuthor" ==> (run.InsertedRevision |> Option.bind (fun mark -> mark.Author))
                                "InsertedRevisionDate" ==> (run.InsertedRevision |> Option.bind (fun mark -> mark.Date))
                                "RevisionMoveName" ==> run.MoveName
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
                    // Direct keepNext / keepLines are the paragraph's own; StyleKeepX is what
                    // its style says without them (the editor infers style definitions from it).
                    "KeepNext" ==> para.Properties.KeepNext
                    "KeepLines" ==> para.Properties.KeepLines
                    "StyleKeepNext" ==> (styleKeepProperties styleMap para effectiveParaProps).KeepNext
                    "StyleKeepLines" ==> (styleKeepProperties styleMap para effectiveParaProps).KeepLines
                    "StylePageBreakBefore" ==> effectiveParaProps.PageBreakBefore
                    "PageBreakBefore" ==> effectiveParaProps.PageBreakBefore
                    "PageBreakBeforeDerived" ==> (effectiveParaProps.PageBreakBefore = Some true && (para.Properties.PageBreakBefore <> Some true || para.Properties.PageBreakBeforeDerived))
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
                    "SectionPageNumberFormat" ==> effectiveParaProps.SectionPageNumberFormat
                    "HeaderReferences" ==> referenceObjects effectiveParaProps.HeaderReferences
                    "FooterReferences" ==> referenceObjects effectiveParaProps.FooterReferences
                    "ParagraphMarkFontSize" ==> paragraphFontSize
                    "MarkRevisionType" ==> (para.MarkRevision |> Option.map fst)
                    "MarkRevisionId" ==> (para.MarkRevision |> Option.bind (fun (_, mark) -> mark.Id))
                    "MarkRevisionAuthor" ==> (para.MarkRevision |> Option.bind (fun (_, mark) -> mark.Author))
                    "MarkRevisionDate" ==> (para.MarkRevision |> Option.bind (fun (_, mark) -> mark.Date))
                ] @ paragraphLayoutProps paragraphFontSize defaultMissingSpacingToZero effectiveParaProps.ParagraphBorders para.Properties layoutParaProps)

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
                        toEditorParagraphsOfStyleMap styleMap theme [Tbl nested] images noteBodies
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
                // The table style and which of its conditional formats apply. Its
                // formatting is resolved into the cells above; the reference is kept so
                // a saved table still has its style.
                "TableStyleId" ==> table.StyleId
                "TableLook" ==> (table.Look
                                 |> Option.map (fun look ->
                                     objFromFields [
                                         "FirstRow" ==> look.FirstRow
                                         "LastRow" ==> look.LastRow
                                         "FirstColumn" ==> look.FirstColumn
                                         "LastColumn" ==> look.LastColumn
                                         "HorizontalBanding" ==> look.HorizontalBanding
                                         "VerticalBanding" ==> look.VerticalBanding ]))
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
                "FloatingXml" ==> table.FloatingXml
                "Borders" ==> bordersToEditor tableBorders
            ])
    )
    |> List.toArray

let private toEditorParagraphsInternal (theme: Theme option) (styles: StyleDefinition list) (content: DocumentContent list) (images: Map<string, string>) (noteBodies: Map<string * string, string>) : obj array =
    toEditorParagraphsOfStyleMap (buildStyleMap styles) theme content images noteBodies

/// What gives the items of a document the lists its numbering part defines.
let private numberingApplier (numberingDefinitions: NumberingDefinition list) : obj array -> obj array =
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
            // The level's own indents: where an indent the paragraph does not state
            // comes from, and what the writer puts back into the numbering definition.
            setProp item "HasListLevelIndents" true
            setProp item "ListLevelIndentLeft" (wordTwipsToPoints def.IndentLeft)
            setProp item "ListLevelFirstLineIndent"
                (match wordTwipsToPoints def.FirstLineIndent, wordTwipsToPoints def.HangingIndent with
                 | Some firstLine, _ -> Some firstLine
                 | None, Some hanging -> Some -hanging
                 | None, None -> None)
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

    fun (items: obj array) -> items |> Array.map apply

let private applyNumberingDefinitions (numberingDefinitions: NumberingDefinition list) (items: obj array) =
    numberingApplier numberingDefinitions items

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
    (footnotes |> List.map (fun note -> ("footnote", note.Id), Notes.plainText note))
    @ (endnotes |> List.map (fun note -> ("endnote", note.Id), Notes.plainText note))
    |> Map.ofList

let toEditorParagraphsWithStylesThemeAndNumbering (theme: Theme option) (content: DocumentContent list) (images: Map<string, string>) (styles: StyleDefinition list) (numberingDefinitions: NumberingDefinition list) (footnotes: Note list) (endnotes: Note list) : obj array =
    toEditorParagraphsInternal theme styles content images (noteBodyMap footnotes endnotes)
    |> applyNumberingDefinitions numberingDefinitions

/// What toEditorParagraphsWithStylesThemeAndNumbering makes of blocks, for a
/// document whose blocks come in steps: the styles, the lists and the notes
/// are looked at once, and each step gives its blocks and their pictures.
let editorParagraphConverter (theme: Theme option) (styles: StyleDefinition list) (numberingDefinitions: NumberingDefinition list) (footnotes: Note list) (endnotes: Note list)
        : DocumentContent list -> Map<string, string> -> obj array =
    let styleMap = buildStyleMap styles
    let noteBodies = noteBodyMap footnotes endnotes
    let applyNumbering = numberingApplier numberingDefinitions
    fun content images -> toEditorParagraphsOfStyleMap styleMap theme content images noteBodies |> applyNumbering

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
            "HeaderReferences" ==> referenceObjects setup.HeaderReferences
            "FooterReferences" ==> referenceObjects setup.FooterReferences
            "PageNumberStart" ==> setup.PageNumberStart
            "PageNumberFormat" ==> setup.PageNumberFormat
        ]

/// Footnote / endnote bodies as editor blocks (paragraphs with runs, tables),
/// one object per note: Id, Kind (w:type of a separator note), HasMark /
/// MarkStyle (the automatic reference mark), Text (the plain text the reference
/// run carries as NoteText) and Blocks.
let toEditorNotes (theme: Theme option) (images: Map<string, string>) (styles: StyleDefinition list) (numberingDefinitions: NumberingDefinition list) (notes: Note list) : obj array =
    notes
    |> List.map (fun note ->
        objFromFields [
            "Id" ==> note.Id
            "Kind" ==> note.Kind
            "HasMark" ==> note.HasMark
            "MarkStyle" ==> note.MarkStyle
            "Text" ==> Notes.plainText note
            "Blocks" ==> (toEditorParagraphsInternal theme styles note.Blocks images Map.empty |> applyNumberingDefinitions numberingDefinitions)
        ])
    |> List.toArray

/// Comment bodies as editor blocks: Id, Author, Date, Initials, Text, Blocks.
let toEditorComments (theme: Theme option) (images: Map<string, string>) (styles: StyleDefinition list) (numberingDefinitions: NumberingDefinition list) (comments: Comment list) : obj array =
    comments
    |> List.map (fun comment ->
        objFromFields [
            "Id" ==> comment.Id
            "Author" ==> comment.Author
            "Date" ==> comment.Date
            "Initials" ==> comment.Initials
            "Text" ==> comment.Text
            "Blocks" ==> (toEditorParagraphsInternal theme styles comment.Blocks images Map.empty |> applyNumberingDefinitions numberingDefinitions)
        ])
    |> List.toArray

/// The kept parts of the package for the editor package object ("Preserved").
/// Part data is a byte array (a Uint8Array in the browser); the object is
/// handed back to DocumentWriter.writeFromEditorPackage unchanged.
let toEditorPreserved (preserved: PackagePreservation) : obj =
    let relationship (rel: PreservedRelationship) =
        objFromFields [ "Id" ==> rel.Id; "Type" ==> rel.Type; "Target" ==> rel.Target; "External" ==> rel.External ]
    let relationships (rels: PreservedRelationship list) = rels |> List.map relationship |> List.toArray
    objFromFields [
        "Parts" ==> (preserved.Parts |> List.map (fun part -> objFromFields [ "Path" ==> part.Path; "Data" ==> part.Data ]) |> List.toArray)
        "RootRelationships" ==> relationships preserved.RootRelationships
        "DocumentRelationshipId" ==> preserved.DocumentRelationshipId
        "PartRelationships" ==> (preserved.PartRelationships
                                 |> Map.toList
                                 |> List.map (fun (key, rels) -> objFromFields [ "Part" ==> key; "Relationships" ==> relationships rels ])
                                 |> List.toArray)
        "Defaults" ==> (preserved.Defaults |> List.map (fun (extension, contentType) -> objFromFields [ "Extension" ==> extension; "ContentType" ==> contentType ]) |> List.toArray)
        "Overrides" ==> (preserved.Overrides |> List.map (fun (partName, contentType) -> objFromFields [ "PartName" ==> partName; "ContentType" ==> contentType ]) |> List.toArray)
        "MainContentType" ==> preserved.MainContentType
        "StylesXml" ==> preserved.StylesXml
        "SourceXml" ==> (preserved.SourceXml |> Map.toList |> List.map (fun (key, text) -> objFromFields [ "Part" ==> key; "Xml" ==> text ]) |> List.toArray)
        "SourceRelationships" ==> (preserved.SourceRelationships
                                   |> Map.toList
                                   |> List.map (fun (key, rels) -> objFromFields [ "Part" ==> key; "Relationships" ==> relationships rels ])
                                   |> List.toArray)
        "SourceParts" ==> (preserved.SourceParts |> List.map (fun part -> objFromFields [ "Path" ==> part.Path; "Data" ==> part.Data ]) |> List.toArray)
        "BodyContentChunks" ==> (preserved.BodyContentChunks |> List.toArray)
        "BodyChunkCount" ==> preserved.BodyChunkCount
        "PartPaths" ==> (preserved.PartPaths |> Map.toList |> List.map (fun (key, path) -> objFromFields [ "Part" ==> key; "Path" ==> path ]) |> List.toArray)
        "OriginalText" ==> (preserved.OriginalText |> Map.toList |> List.map (fun (path, text) -> objFromFields [ "Path" ==> path; "Text" ==> text ]) |> List.toArray)
        "SettingsXml" ==> preserved.SettingsXml
    ]
