module DocxReader

open XmlDsl
open DocxTypes
open OfficeDsl
open ZipInterop
open Fable.Core.JsInterop
open System


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
    /// w:lang@eastAsia: the language of the run's East Asian text (ko-KR, ja-JP, zh-CN ...).
    EastAsiaLanguage: string option
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
    /// The picture is mirrored across (a:xfrm flipH of its pic:spPr; flip:x in
    /// the style of a VML picture) / upside down (flipV; flip:y).
    FlipH: bool
    FlipV: bool
    /// Set for an object the reader does not model (chart, SmartArt, OLE object,
    /// grouped / VML drawing, unknown mc:AlternateContent): the element as it is
    /// in the file, to be written back unchanged. The other fields then describe
    /// its frame; RelationshipId is that of a fallback picture, or "".
    Opaque: OpaqueObject option
}

/// An element kept verbatim (DocxTypes.OpaqueObject).
and OpaqueObject = DocxTypes.OpaqueObject

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
    /// The text of this run where the result of the field is several runs - parts of it are
    /// formatted differently: each run has the field's code and its whole result, and its
    /// own part of the result's text.
    Part: string option
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
    | FieldBoundary of ModelText.FieldBoundary
    | EquationXml of string
    | RenderedPageBreak
    | PageBreak

/// Id, author and date of a tracked change (DocxTypes.RevisionMark).
type RevisionMark = DocxTypes.RevisionMark

type ContentControlLayer = {
    Group: string
    Id: string option
    Tag: string option
    Alias: string option
    Xml: string option
}

type Run = {
    Content: RunContent
    Properties: RunProperties
    /// The enclosing w:hyperlink for field tokens and their result runs.
    /// Ordinary hyperlinks remain Link content; fields need both identities.
    HyperlinkContext: Hyperlink option
    /// The bookmarks that start before the run and those that end (or are opened) after it, in the order
    /// of the file (see attachRunMarkers).
    BookmarkStarts: ModelText.BookmarkStart list
    BookmarkEnds: ModelText.BookmarkEnd list
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
    /// The glyph of a w14:checkbox content control (☐ / ☒): its checked state.
    /// The run is the control's - one character, never text typed beside it.
    Checkbox: bool option
    /// Inline w:sdt wrappers, outermost first; their original properties and
    /// distinct source identities survive a model-only save.
    ContentControls: ContentControlLayer list
    /// The w:r the run was read from, in its part's text (docs/save-architecture.md);
    /// Span.none for a run that is not one (a bookmark, a picture outside a run).
    Source: XmlDsl.Span
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
    /// w:between: the border between the paragraphs of a group with the same borders.
    Between: ParagraphBorder option
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
    /// The w:sdtPr of the block-level content control the paragraph is in,
    /// self-contained (completeOpaqueObjects): its binding, placeholder, kind.
    ContentControlXml: string option
    ContentControlGroup: string option
    ContentControlAncestors: ContentControlLayer list
    /// Controls around this paragraph inside a cell, retained separately from emitted block controls.
    EnclosingCellControls: ContentControlLayer list
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
    /// w:overflowPunct: punctuation may hang past the line's end (East Asian line breaking).
    OverflowPunctuation: bool option
    BackgroundColor: string option
    ShadingPattern: string option
    ShadingForeground: string option
    ParagraphRunProperties: RunProperties option
    ParagraphBorders: ParagraphBorders option
    ListInfo: ListInfo option
    SectionBreak: string option
    SectionColumns: SectionColumns option
    SectionGeometry: SectionGeometry option
    SectionPageNumberStart: int option
    SectionPageNumberFormat: string option
    /// w:vAlign of the section the paragraph ends ("top", "center", "both", "bottom").
    SectionVerticalAlign: string option
    /// A break in the paragraph is a column break (w:br w:type="column"), not a page break.
    ColumnBreak: bool
    RenderedPageBreakBefore: int option
    HeaderReferences: HeaderFooterReference list
    FooterReferences: HeaderFooterReference list
    /// w:framePr w:dropCap: the paragraph is a drop cap - ("drop" | "margin",
    /// w:lines, w:hSpace in twips).
    DropCap: (string * int * string option) option
}

type Paragraph = {
    Runs: Run list
    Properties: ParagraphProperties
    /// The paragraph mark is a tracked change (w:ins / w:del in w:pPr/w:rPr):
    /// ("insert" | "delete", id / author / date).
    MarkRevision: (string * RevisionMark) option
    /// The w:p the paragraph was read from, in its part's text; Span.none for one made.
    Source: XmlDsl.Span
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
    ContentControlId: string option
    ContentControlTag: string option
    ContentControlAlias: string option
    ContentControlXml: string option
    ContentControlGroup: string option
    ContentControlAncestors: ContentControlLayer list
    Look: TableLook option
    Alignment: string option
    Indent: string option
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
    /// The w:tr in its part (offsets into the part's text).
    Source: XmlDsl.Span
}
and TableCell = {
    Paragraphs: Paragraph list
    NestedTables: Table list
    // For each nested table, how many cell paragraphs precede it in document order.
    NestedTableAnchors: int list
    /// Cell-level controls surrounding this cell, outermost first.
    EnclosingControls: ContentControlLayer list
    Properties: TableCellProperties option
    /// The w:tc in its part (offsets into the part's text).
    Source: XmlDsl.Span
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
    /// The child element of the part's root each of Blocks was read from, and how many elements it has
    /// (as Docx.bodyContentChunks says of the body).
    ItemChunks: int list
    ChunkCount: int
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

/// The theme's fonts and colours (DocxTypes.Theme).
type Theme = DocxTypes.Theme

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
    /// The child element of w:comment each of Blocks was read from, and how many elements it has
    /// (as Docx.bodyContentChunks says of the body).
    ItemChunks: int list
    ChunkCount: int
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
    /// The child element of w:footnote / w:endnote each of Blocks was read from, and how many elements it
    /// has (as Docx.bodyContentChunks says of the body).
    ItemChunks: int list
    ChunkCount: int
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
    /// The shape is mirrored across / upside down (a:xfrm flipH / flipV; flip:x / flip:y
    /// in a VML shape's style), before it is turned. Its text is not.
    FlipH: bool
    FlipV: bool
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
    /// The shapes of a group or a drawing canvas (wpg:wgp, wpc:wpc, v:group), in
    /// the order they are drawn: each with its Position and Width / Height in
    /// pixels from the group's top-left corner, as the group's child coordinates
    /// (a:chOff / a:chExt, coordorigin / coordsize) put it in the group's box as
    /// read (ChildSpace); a group in a group has shapes of its own. Empty for
    /// any other shape. A group has no text of its own (Content is empty).
    Children: Shape list
    /// The size (pixels) the places of the children are given in: the group's
    /// own as read - a group drawn larger or smaller scales its shapes with it.
    /// None: they are not scaled (a drawing canvas, whose shapes keep their size
    /// when it is resized).
    ChildSpace: (float * float) option
    /// A picture of a group (pic:pic, a VML shape's v:imagedata): the
    /// relationship id of the picture it shows, drawn in its box.
    ImageRelationshipId: string option
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
    /// w:defaultTabStop of the settings part, in pixels: the distance of the
    /// default tab stops (Word's 720 twips when the part says none).
    DefaultTabStop: float option
    /// The compatibilityMode compatSetting of the settings part (14: Word 2010,
    /// 15: Word 2013 and later); None when it has none (an older document).
    CompatibilityMode: int option
    /// w:useFELayout of the settings part's w:compat (East Asian line breaking).
    UseFarEastLayout: bool option
    /// w:characterSpacingControl of the settings part: doNotCompress, compressPunctuation,
    /// compressPunctuationAndJapaneseKana; None when it has none.
    CharacterSpacingControl: string option
    /// w:autoHyphenation of the settings part, when on: "<zone> <limit> <caps>" - w:hyphenationZone in
    /// twips (-1: none), w:consecutiveHyphenLimit (0: none), w:doNotHyphenateCaps (1 or 0); None when off.
    Hyphenation: string option
    Columns: SectionColumns option
    SectionStart: string option   // body sectPr w:type (e.g. "continuous")
    /// Header / footer references and page number start of the body-level
    /// sectPr, i.e. of the final section.
    HeaderReferences: HeaderFooterReference list
    FooterReferences: HeaderFooterReference list
    PageNumberStart: int option
    PageNumberFormat: string option
    /// w:vAlign of the final section.
    VerticalAlign: string option
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
    /// w:suff: what follows the number ("tab" when not stated, "space", "nothing").
    Suffix: string option
    /// w:lvlJc: how the number stands at its position ("left" when not stated, "center", "right").
    Justification: string option
    /// w:isLgl: the numbers of all levels in this level's text are decimal.
    Legal: bool
    /// w:lvlRestart: the level (1-9) after which this one starts again; 0: never.
    Restart: int option
    /// The run properties of the number (the level's w:rPr): face, bold, italic, colour (w:val), size (half points).
    MarkFont: string option
    MarkBold: bool option
    MarkItalic: bool option
    MarkColor: string option
    MarkSize: string option
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

    /// The XML of `node`, serialised (XmlText.xmlToString).
    let xmlToString node = XmlText.xmlToString node

    let private attrByLocalName localName node =
        match node with
        | XmlDsl.Element(_, attrs, _, _) ->
            attrs
            |> Map.tryPick (fun name value ->
                if Ns.stripNs name = localName then Some value else None)
        | _ -> None

    let attr attrName node =
        Xml.attr attrName node
        |> Option.orElseWith (fun () ->
            Xml.name node
            |> Option.bind (fun name -> Xml.attr (Ns.withPrefixOf name attrName) node))
        |> Option.orElseWith (fun () -> attrByLocalName attrName node)

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
        // (A sectPr that has more than one w:cols - one written by a template, one by a tool after
        // it - is laid out by Word with the last.)
        Xml.children sectPr
        |> List.filter (fun child -> Xml.name child |> Option.map Ns.stripNs = Some "cols")
        |> List.tryLast
        |> Option.bind (fun cols ->
            let num = attr "num" cols |> Option.bind (fun v -> match System.Int32.TryParse v with | true, n -> Some n | _ -> None) |> Option.defaultValue 1
            let space = twips (attr "space" cols) |> Option.defaultValue 36.0
            // w:sep is an attribute of w:cols (onOff looks for a child element, so
            // the separator line was never read).
            let sep =
                match attr "sep" cols |> Option.map (fun value -> value.ToLowerInvariant()) with
                | Some "1" | Some "true" | Some "on" -> true
                | _ -> false
            // (w:equalWidth: the columns are w:num of one width, whatever w:col elements are left there.)
            let equalWidth =
                match attr "equalWidth" cols |> Option.map (fun value -> value.ToLowerInvariant()) with
                | Some "1" | Some "true" | Some "on" -> true
                | _ -> false
            let colWidths =
                (if equalWidth then [] else Ns.descendantsNs "col" cols)
                |> List.choose (fun col -> twips (attr "w" col) |> Option.map (fun w -> w, twips (attr "space" col) |> Option.defaultValue 0.0))
            let count = if colWidths.Length >= 2 then colWidths.Length else num
            if count <= 1 then None
            else Some { Count = count; Gap = space; Separator = sep; Widths = (if colWidths.Length >= 2 then colWidths else []) })

    // Parse a sectPr's pgSz/pgMar into a SectionGeometry (twips/15 = px).
    let parseSectionGeometry (sectPr: XmlNode) : SectionGeometry option =
        let twips (s: string option) = s |> Option.bind (fun v -> match System.Double.TryParse v with | true, n -> Some (n / 15.0) | _ -> None)
        match Ns.childNs "pgSz" sectPr, Ns.childNs "pgMar" sectPr with
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
                // (On one element a theme face beats the named one, as in Word: fixtures rfonts-precedence.)
                FontFamily = (match getPropVal "rFonts" "asciiTheme" props with Some _ -> None | None -> getPropVal "rFonts" "ascii" props)
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
                EastAsiaLanguage = getPropVal "lang" "eastAsia" props
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
                EastAsiaLanguage = None
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
            Between = paragraphBorder "between" borders
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
                ContentControlXml = None
                ContentControlGroup = None
                ContentControlAncestors = []
                EnclosingCellControls = []
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
                OverflowPunctuation = onOff "overflowPunct" props
                BackgroundColor = getPropVal "shd" "fill" props
                ShadingPattern = getPropVal "shd" "val" props
                ShadingForeground = getPropVal "shd" "color" props
                ParagraphRunProperties = wElem "rPr" props |> Option.map (fun _ -> runProps props)
                ParagraphBorders = paragraphBorders props
                ListInfo = listInfo
                SectionBreak = sectPr |> Option.map (fun section -> getPropVal "type" "val" section |> Option.defaultValue "nextPage")
                SectionColumns = sectPr |> Option.bind parseColumns
                SectionGeometry = sectPr |> Option.bind parseSectionGeometry
                SectionPageNumberStart = sectPr |> Option.bind (fun section -> getPropVal "pgNumType" "start" section) |> Option.bind (fun value -> match System.Int32.TryParse(value) with | true, parsed -> Some parsed | false, _ -> None)
                SectionPageNumberFormat = sectPr |> Option.bind (fun section -> getPropVal "pgNumType" "fmt" section)
                SectionVerticalAlign = sectPr |> Option.bind (fun section -> getPropVal "vAlign" "val" section)
                ColumnBreak = false
                RenderedPageBreakBefore = None
                HeaderReferences = sectPr |> Option.map (headerFooterReferences "headerReference") |> Option.defaultValue []
                FooterReferences = sectPr |> Option.map (headerFooterReferences "footerReference") |> Option.defaultValue []
                DropCap =
                    wElem "framePr" props
                    |> Option.bind (fun frame ->
                        match attr "dropCap" frame with
                        | Some ("drop" | "margin" as kind) ->
                            let lines =
                                match attr "lines" frame |> Option.map System.Int32.TryParse with
                                | Some (true, value) when value > 0 -> value
                                | _ -> 1
                            Some (kind, lines, attr "hSpace" frame)
                        | _ -> None)
            }
        | None ->
            {
                Style = None
                ContentControlId = None
                ContentControlTag = None
                ContentControlAlias = None
                ContentControlXml = None
                ContentControlGroup = None
                ContentControlAncestors = []
                EnclosingCellControls = []
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
                OverflowPunctuation = None
                BackgroundColor = None
                ShadingPattern = None
                ShadingForeground = None
                ParagraphRunProperties = None
                ParagraphBorders = None
                ListInfo = None
                SectionBreak = None
                SectionColumns = None
                SectionGeometry = None
                SectionPageNumberStart = None; SectionPageNumberFormat = None
                SectionVerticalAlign = None
                ColumnBreak = false
                RenderedPageBreakBefore = None
                HeaderReferences = []
                FooterReferences = []
                DropCap = None
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
        // (The picture's own transform: pic:pic > pic:spPr > a:xfrm.)
        let pictureFlip attrName =
            Ns.descendantsNs "pic" drawing |> List.tryHead
            |> Option.bind (Ns.childNs "spPr")
            |> Option.bind (Ns.childNs "xfrm")
            |> Option.bind (Xml.attr attrName)
            |> Option.exists (fun value -> value = "1" || value.ToLowerInvariant() = "true")
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
            FlipH = pictureFlip "flipH"
            FlipV = pictureFlip "flipV"
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
            FlipH = false
            FlipV = false
            Opaque = None
        }
        let drawingImage =
            wElem "drawing" rNode
            // (A group or a drawing canvas is not its picture: see Docx.groupElement.)
            |> Option.filter (fun drawing -> (Ns.descendantsNs "wgp" drawing).IsEmpty && (Ns.descendantsNs "wpc" drawing).IsEmpty)
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
            |> Option.filter (fun pict -> (Ns.descendantsNs "group" pict).IsEmpty)
            |> Option.bind (fun pict ->
                let image = Ns.descendantsNs "imagedata" pict |> List.tryHead
                image
                |> Option.bind (fun image ->
                    let embedId = Xml.attr "r:id" image |> Option.orElseWith (fun () -> attrByLocalName "id" image)
                    embedId
                    |> Option.map (fun rId ->
                        let shape = Ns.descendantsNs "shape" pict |> List.tryHead
                        let style = shape |> Option.bind (Xml.attr "style")
                        let styleValue name = style |> Option.bind (ValueText.Css.valueIgnoringCase name)
                        let stylePoints name =
                            styleValue name
                            |> Option.map ValueText.Css.withoutPoints
                            |> Option.bind parseFloat
                            |> Option.map ptToPixels
                        let flipH, flipV = style |> Option.map ValueText.Css.flip |> Option.defaultValue (false, false)
                        { emptyImage rId false (shape |> Option.bind (fun value -> Xml.attr "id" value)) (Xml.attr "o:title" image) (shape |> Option.bind (fun value -> Xml.attr "alt" value)) (stylePoints "width") (stylePoints "height")
                            with FlipH = flipH; FlipV = flipV })))
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

    let private localNameOf (node: XmlNode) = Xml.name node |> Option.map Ns.stripNs |> Option.defaultValue ""

    let private isVmlElement (node: XmlNode) =
        match node with
        | XmlDsl.Element (name, _, _, _) -> Ns.hasPrefix "v" name
        | _ -> false

    /// The group or drawing canvas a drawing is (wpg:wgp, wpc:wpc under the
    /// a:graphicData of a wp:anchor / wp:inline), or a VML group itself.
    let groupElement (node: XmlNode) : XmlNode option =
        match localNameOf node with
        | "group" when isVmlElement node -> Some node
        | "anchor" | "inline" ->
            Ns.childNs "graphic" node
            |> Option.bind (Ns.childNs "graphicData")
            |> Option.bind (fun data -> Xml.children data |> List.tryFind (fun child -> match localNameOf child with "wgp" | "wpc" -> true | _ -> false))
        | _ -> None

    /// The a:ext of an element's own a:xfrm (wps:spPr, pic:spPr, wpg:grpSpPr) - not one of a shape in it.
    let private ownExtent (element: XmlNode) : XmlNode option =
        Xml.children element
        |> List.tryFind (fun child -> match localNameOf child with "spPr" | "grpSpPr" -> true | _ -> false)
        |> Option.bind (Ns.childNs "xfrm")
        |> Option.bind (Ns.childNs "ext")

    /// The relationship id of the picture of a pic:pic (its a:blip's r:embed or r:link).
    let pictureRelationshipOf (picture: XmlNode) : string option =
        Ns.descendantsNs "blip" picture
        |> List.tryHead
        |> Option.bind (fun blip -> Xml.attr "r:embed" blip |> Option.orElse (Xml.attr "r:link" blip))

    /// The relationship id of the picture of a VML shape (its v:imagedata).
    let vmlPictureRelationshipOf (shape: XmlNode) : string option =
        Ns.descendantsNs "imagedata" shape
        |> List.tryHead
        |> Option.bind (fun image -> Xml.attr "r:id" image |> Option.orElse (Xml.attr "o:relid" image))

    /// Whether the shapes of a group are drawn of the element of one (and so whether
    /// the shape model holds the group): a DrawingML shape, picture or group with
    /// its place (a:xfrm); a picture that says which picture it shows; a VML
    /// rectangle, oval, rounded rectangle or line; a VML shape with text, a
    /// picture or an outline the reader knows (a text box's, a rectangle's -
    /// not a path, nor a shape type of its own: an arrow, a connector); a group
    /// of such. Every element of the group is one of those or what the group
    /// says of itself, and there is one at least: a group with a chart, a
    /// diagram, ink or a kind of drawing the reader does not draw is kept as
    /// it is (its fallback picture is what is shown).
    let rec isDrawnGroupChild (node: XmlNode) : bool =
        match node with
        | XmlDsl.Element (name, _, _, _) ->
            let local = Ns.stripNs name
            if Ns.hasPrefix "v" name then
                match local with
                | "rect" | "oval" | "roundrect" | "line" -> true
                | "shape" ->
                    (vmlPictureRelationshipOf node).IsSome
                    || not (Ns.descendantsNs "txbxContent" node).IsEmpty
                    || (match Xml.attr "type" node with
                        | None -> (Xml.attr "path" node).IsNone
                        | Some kind -> kind = "#_x0000_t202" || kind = "#_x0000_t1")
                | "group" -> isGroupOfShapes node
                | _ -> false
            else
                match local with
                | "wsp" -> (ownExtent node).IsSome
                | "pic" -> (ownExtent node).IsSome && (pictureRelationshipOf node).IsSome
                | "grpSp" | "wgp" -> (ownExtent node).IsSome && isGroupOfShapes node
                | _ -> false
        | _ -> false

    and private isGroupOfShapes (group: XmlNode) : bool =
        let vml = isVmlElement group
        let mutable drawn = false
        let mutable known = true
        for child in Xml.children group do
            match child with
            | XmlDsl.Element (name, _, _, _) ->
                let local = Ns.stripNs name
                if vml then
                    if Ns.hasPrefix "v" name then
                        match local with
                        | "shape" | "rect" | "oval" | "roundrect" | "line" | "group" ->
                            if isDrawnGroupChild child then drawn <- true
                            // (A group in it that is not drawn is not: something in it is unknown.)
                            elif local = "group" then known <- false
                        | "shapetype" | "path" | "fill" | "stroke" | "shadow" | "textbox" | "formulas" | "handles" -> ()
                        | _ -> known <- false
                    // (o:lock, w10:wrap: what the group says of itself.)
                else
                    match local with
                    | "wsp" | "pic" | "grpSp" | "wgp" ->
                        if isDrawnGroupChild child then drawn <- true
                        elif local = "grpSp" || local = "wgp" then known <- false
                    | "cNvPr" | "cNvGrpSpPr" | "nvGrpSpPr" | "grpSpPr" | "extLst" | "bg" | "whole" -> ()
                    | _ -> known <- false
            | _ -> ()
        drawn && known

    /// A group or drawing canvas the shape model holds (see isGroupOfShapes).
    let isModelledGroup (node: XmlNode) : bool =
        groupElement node |> Option.exists isGroupOfShapes

    /// A drawing the shape model can hold: one shape with text and / or
    /// geometry. Groups and drawing canvases are held as groups
    /// (isModelledGroup); charts and diagrams are not.
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
        | Some "anchor" | Some "inline" -> if isModelledShape node || isModelledGroup node then [node] else []
        | Some "shape" | Some "rect" | Some "roundrect" | Some "oval" when (match node with XmlDsl.Element(name, _, _, _) -> Ns.hasPrefix "v" name | _ -> false) ->
            if isModelledShape node then [node] else []
        // (A VML group is one drawing: its shapes are not shapes of their own.)
        | Some "group" when isVmlElement node -> if isModelledGroup node then [node] else []
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

    /// The opaque kind of a page or column break inside a paragraph's text (placeInlineBreaks).
    let inlineBreakKind (kind: string) = "break-" + kind
    let isInlineBreakKind (kind: string) = kind = "break-page" || kind = "break-column"

    let private blankImageData (relationshipId: string) : ImageData =
        { RelationshipId = relationshipId; Linked = false; Name = None; Title = None; Width = None; Height = None; Description = None
          Floating = false; Wrap = None; PositionX = None; PositionY = None; PositionHRelativeFrom = None; PositionVRelativeFrom = None
          BehindDoc = None; AllowOverlap = None; LayoutInCell = None; DistanceTop = None; DistanceBottom = None; DistanceLeft = None
          DistanceRight = None; RelativeHeight = None; Locked = None; SimplePosition = None; SimplePositionX = None; SimplePositionY = None
          CropLeft = None; CropTop = None; CropRight = None; CropBottom = None; PositionHAlign = None; PositionVAlign = None; WrapText = None; WrapPolygon = []
          FlipH = false; FlipV = false; Opaque = None }

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
                    ValueText.Css.valueIgnoringCase name style
                    |> Option.bind (fun value ->
                        let raw = value.ToLowerInvariant()
                        let number (text: string) =
                            match System.Double.TryParse(text, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture) with
                            | true, parsed -> Some parsed
                            | false, _ -> None
                        match ValueText.Css.withUnit [ "pt"; "in"; "px" ] raw with
                        | Some (value, "pt") -> number value |> Option.map (fun pt -> pt * 96.0 / 72.0)
                        | Some (value, "in") -> number value |> Option.map (fun inches -> inches * 96.0)
                        | Some (value, _) -> number value
                        | None -> number raw)
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
        if ValueText.FieldCode.startsWith "DATE" upper then Date
        elif ValueText.FieldCode.startsWith "PAGE" upper then PageNumber
        elif ValueText.FieldCode.startsWith "NUMPAGES" upper then NumPages
        elif ValueText.FieldCode.startsWith "HYPERLINK" upper then Hyperlink normalized
        elif ValueText.FieldCode.startsWith "REF " upper then Ref (ValueText.FieldCode.refTarget normalized)
        else Unknown normalized

    /// The runs of a field shown by its own element (w:fldSimple): one run with the field's
    /// code and its result - or, where the runs of its result are formatted differently, as
    /// many runs, each with its own part of the result's text (see completeFieldRuns).
    let simpleFieldRuns (fldNode: XmlNode) : Run list =
        // OOXML serialises the instruction code on the w: namespaced attribute
        // ("w:instr"). Xml.attr is a plain-string lookup on the attrs map, so
        // we have to ask for the prefixed form for documents that ship the
        // namespace declaration this way. Fall back to the bare local name
        // for parsers/serialisers that strip the prefix.
        match (Xml.attr "w:instr" fldNode |> Option.orElseWith (fun () -> Xml.attr "instr" fldNode)) with
        | None -> []
        | Some code ->
            let runs = wElems "r" fldNode
            let texts = runs |> List.map (fun run -> wElem "t" run |> Option.bind Xml.text |> Option.defaultValue "")
            let result = String.concat "" texts
            let fieldType = fieldTypeFromCode code
            // (A run of the result with Word's mark of a page it started: the part from it on has the mark.)
            let marked (run: XmlNode) = (wElem "lastRenderedPageBreak" run).IsSome
            let fieldOf (props: RunProperties) (part: string option) (renderedPageBreak: bool) : Run =
                { Content = FieldRef { FieldType = fieldType; Code = code; Result = (if result = "" then None else Some result); Part = part }
                  Properties = props
                  BookmarkStarts = []
                  BookmarkEnds = []
                  RenderedPageBreakBefore = renderedPageBreak
                  RevisionId = None
                  RevisionType = None
                  RevisionAuthor = None
                  RevisionDate = None; InsertedRevision = None; MoveName = None; Checkbox = None
                  HyperlinkContext = None; ContentControls = []; Source = XmlDsl.Span.none }
            // (A field of one run is of that run: a save writes it from there while it is as read.)
            let whole () =
                let field = fieldOf (runs |> List.tryHead |> Option.map runProps |> Option.defaultValue (runProps fldNode)) None (runs |> List.exists marked)
                [ match runs with
                  | [ XmlDsl.Element (_, _, _, span) ] -> { field with Source = span }
                  | _ -> field ]
            let computed = ValueText.FieldCode.isComputed (code.Trim().ToUpperInvariant())
            if runs.Length < 2 || computed then whole ()
            else
                // The stretches of runs that are formatted alike.
                let groups = ResizeArray<RunProperties * System.Text.StringBuilder * bool>()
                List.iter2 (fun (run: XmlNode) (text: string) ->
                    let props = runProps run
                    let mark = marked run
                    if groups.Count > 0 && not mark && (let (previous, _, _) = groups.[groups.Count - 1] in previous = props) then
                        (let (_, builder, _) = groups.[groups.Count - 1] in builder.Append(text) |> ignore)
                    else groups.Add ((props, System.Text.StringBuilder(text), mark))) runs texts
                let parts = [ for (props, text, mark) in groups do if text.Length > 0 then yield fieldOf props (Some (text.ToString())) mark ]
                if parts.Length < 2 then whole () else parts

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
                    Part = None
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
        /// Source runs and formatting remain attached to structural field tokens.
        | FieldBeginToken of Run
        | FieldSeparateToken of Run
        | FieldEndToken of Run
        | FieldInstructionToken of Run

    type private ComplexFieldState = {
        Properties: RunProperties
        /// The w:r of the field's begin mark: the source of the field's run (see collapseComplexFieldTokens).
        BeginSource: XmlDsl.Span
        BeginHyperlink: Hyperlink option
        CodeParts: ResizeArray<string>
        ResultRuns: ResizeArray<Run>
        mutable CollectingResult: bool
    }

    let private textFromRun (run: Run) =
        match run.Content with
        | Text text -> text
        | Link link -> link.Text
        | FieldRef field -> field.Part |> Option.orElse field.Result |> Option.defaultValue field.Code
        | Image img -> img.Description |> Option.map (sprintf "[Image: %s]") |> Option.defaultValue "[Image]"
        | NoteRef note -> sprintf "[%s %s]" note.NoteType note.Id
        | EquationXml _ -> "[Equation]"
        | CommentReference _
        | BookmarkStart _
        | BookmarkEnd _
        | RenderedPageBreak
        | PageBreak -> ""
        | FieldBoundary _ -> ""

    let private appendRun (output: Run list) (stack: ComplexFieldState list) (run: Run) =
        match stack with
        | [] -> run :: output, stack
        | current :: rest ->
            current.ResultRuns.Add run
            output, stack

    /// The run of a field that ends: its code, and the text of its result. A result of which
    /// parts are formatted differently (text typed into it in another format, a part made
    /// bold) is as many runs, each with the field's code, its whole result and its own part
    /// of the text - as the editor holds a result a part of which was formatted. (Not the
    /// number of a page: the editor shows its own.)
    let private completeFieldRuns (field: ComplexFieldState) : Run list =
        let codeBuilder = System.Text.StringBuilder()
        for part in field.CodeParts do
            codeBuilder.Append part |> ignore
        let code = codeBuilder.ToString().Trim()
        let resultBuilder = System.Text.StringBuilder()
        for run in field.ResultRuns do
            resultBuilder.Append(textFromRun run) |> ignore
        let result = resultBuilder.ToString()
        let fieldType = fieldTypeFromCode code
        let isMark (run: Run) = match run.Content with RenderedPageBreak -> true | _ -> false
        // (Word's mark of a page it started in the result: the field's run has it.)
        let fieldOf (first: Run option) (part: string option) (renderedPageBreak: bool) : Run =
            {
                Content = FieldRef {
                    FieldType = fieldType
                    Code = code
                    Result = if result = "" then None else Some result
                    Part = part
                }
                Properties = (match first with Some run -> run.Properties | None -> field.Properties)
                BookmarkStarts = first |> Option.map (fun run -> run.BookmarkStarts) |> Option.defaultValue []
                BookmarkEnds = first |> Option.map (fun run -> run.BookmarkEnds) |> Option.defaultValue []
                RenderedPageBreakBefore = renderedPageBreak || first |> Option.exists (fun run -> run.RenderedPageBreakBefore)
                RevisionId = first |> Option.bind (fun run -> run.RevisionId)
                RevisionType = first |> Option.bind (fun run -> run.RevisionType)
                RevisionAuthor = first |> Option.bind (fun run -> run.RevisionAuthor)
                RevisionDate = first |> Option.bind (fun run -> run.RevisionDate)
                InsertedRevision = first |> Option.bind (fun run -> run.InsertedRevision)
                MoveName = first |> Option.bind (fun run -> run.MoveName)
                Checkbox = None
                ContentControls = first |> Option.map (fun run -> run.ContentControls) |> Option.defaultValue []
                HyperlinkContext = first |> Option.bind (fun run -> run.HyperlinkContext) |> Option.orElse field.BeginHyperlink
                Source = (match first with Some run -> run.Source | None -> XmlDsl.Span.none)
            }
        let marked = field.ResultRuns |> Seq.exists isMark
        let firstResultRun = field.ResultRuns |> Seq.tryFind (isMark >> not)
        let whole () = [ fieldOf firstResultRun None marked ]
        let computed = ValueText.FieldCode.isComputed (code.ToUpperInvariant())
        // (A mark in the result starts a part of its own there, as a change of format does: the line and the page
        // it marks start where Word started them - docx4j-headers-footers' "§Error! Reference |source not found."
        // in a narrow cell, whose mark at the result's start broke the line after "§". A result that is the mark
        // and one stretch of text is the field's one run, the mark at its start.)
        let plain = field.ResultRuns |> Seq.forall (fun run -> match run.Content with Text _ | RenderedPageBreak -> true | _ -> false)
        if field.ResultRuns.Count < 2 || computed || not plain then whole ()
        else
            // The stretches of runs that are formatted alike, a mark starting one too.
            let groups = ResizeArray<Run * System.Text.StringBuilder * bool>()
            let mutable pendingMark = false
            for run in field.ResultRuns do
                if isMark run then pendingMark <- true
                else
                    let text = textFromRun run
                    let sameAsLast =
                        groups.Count > 0 && not pendingMark
                        && (let (first, _, _) = groups.[groups.Count - 1] in obj.ReferenceEquals(first.Properties, run.Properties) || first.Properties = run.Properties)
                    if sameAsLast then
                        let (_, builder, _) = groups.[groups.Count - 1]
                        builder.Append(text) |> ignore
                    elif text.Length > 0 || groups.Count = 0 then
                        groups.Add ((run, System.Text.StringBuilder(text), pendingMark))
                        pendingMark <- false
            if groups.Count < 2 then whole ()
            else [ for (first, text, mark) in groups do if text.Length > 0 then yield fieldOf (Some first) (Some (text.ToString())) mark ]

    // An unbalanced paragraph belongs to a field spanning paragraphs. Its
    // structural tokens must stay in place; collapsing the inner fields or
    // discarding the unmatched end would change the enclosing field's extent.
    let private spansParagraph (tokens: RunToken list) =
        let mutable depth = 0
        let mutable unmatched = false
        for token in tokens do
            match token with
            | FieldBeginToken _ -> depth <- depth + 1
            | FieldEndToken _ when depth = 0 -> unmatched <- true
            | FieldEndToken _ -> depth <- depth - 1
            | FieldSeparateToken _ | FieldInstructionToken _ when depth = 0 -> unmatched <- true
            | _ -> ()
        unmatched || depth <> 0

    let private boundaryRuns (tokens: RunToken list) =
        tokens |> List.map (function
            | RunToken run -> run
            | FieldBeginToken run -> run
            | FieldInstructionToken run -> run
            | FieldSeparateToken run -> run
            | FieldEndToken run -> run)

    let private collapseComplexFieldTokens (tokens: RunToken list) =
        let mutable output: Run list = []
        let mutable stack: ComplexFieldState list = []

        // A bookmark that starts or ends inside a field (between the field's characters: a form
        // field's name) is not part of the field's result: it starts before the field's run and
        // ends after it. (It used to be lost with the field's characters.)
        let endsAfterField = ResizeArray<Run>()
        for token in tokens do
            match token with
            | RunToken run when not stack.IsEmpty && (match run.Content with BookmarkStart _ -> true | _ -> false) ->
                output <- run :: output
            | RunToken run when not stack.IsEmpty && (match run.Content with BookmarkEnd _ -> true | _ -> false) ->
                endsAfterField.Add run
            | RunToken run ->
                let newOutput, newStack = appendRun output stack run
                output <- newOutput
                stack <- newStack
            | FieldBeginToken run ->
                stack <- { Properties = run.Properties; BeginSource = run.Source; BeginHyperlink = run.HyperlinkContext
                           CodeParts = ResizeArray<string>(); ResultRuns = ResizeArray<Run>(); CollectingResult = false } :: stack
            | FieldInstructionToken run ->
                match stack, run.Content with
                | current :: _, FieldBoundary (ModelText.Instruction code) when not current.CollectingResult ->
                    current.CodeParts.Add code
                | _ -> ()
            | FieldSeparateToken _ ->
                match stack with
                | current :: _ -> current.CollectingResult <- true
                | [] -> ()
            | FieldEndToken _ ->
                match stack with
                | current :: rest ->
                    stack <- rest
                    // (A field that is one run is of the run that begins it: a save writes it - begin mark to
                    // end mark, its form field's data, its code - from there. DocumentWriter.splicedRuns.)
                    let fieldRuns =
                        match completeFieldRuns current with
                        | [ one ] when XmlDsl.Span.isKnown current.BeginSource -> [ { one with Source = current.BeginSource } ]
                        | several -> several
                    for fieldRun in fieldRuns do
                        let newOutput, newStack = appendRun output stack fieldRun
                        output <- newOutput
                        stack <- newStack
                    if stack.IsEmpty then
                        for ending in endsAfterField do output <- ending :: output
                        endsAfterField.Clear()
                | [] -> ()

        for ending in endsAfterField do output <- ending :: output

        List.rev output

    let private runsFromTokens tokens =
        if spansParagraph tokens then boundaryRuns tokens
        else collapseComplexFieldTokens tokens

    // The bookmarks that start before a run and end after it are the run's. A
    // run can have several of each (three w:bookmarkStart before a heading's
    // text): all of them, in the order of the file (only the first start and
    // the last end used to be kept: the others were lost once the paragraph
    // was written from the model). A bookmark that ends where it starts is a
    // start of the run after it that ends here (EndsHere); one that is opened
    // after the last run (and whatever ends there) is written after that run:
    // its Opens, after what ends there.
    let private attachRunMarkers (runs: Run list) =
        let pendingStarts = ResizeArray<ModelText.BookmarkStart>()
        let mutable pendingRenderedPageBreak = false
        let output = ResizeArray<Run>()

        let attachStart run =
            let withBookmark =
                if pendingStarts.Count = 0 then run
                else { run with BookmarkStarts = List.ofSeq pendingStarts }
            if pendingRenderedPageBreak then { withBookmark with RenderedPageBreakBefore = true } else withBookmark
        let addEnd (ending: ModelText.BookmarkEnd) =
            if output.Count > 0 then
                let previous = output.[output.Count - 1]
                output.[output.Count - 1] <- { previous with BookmarkEnds = previous.BookmarkEnds @ [ ending ] }

        for run in runs do
            match run.Content with
            | BookmarkStart bookmark ->
                pendingStarts.Add { Id = bookmark.Id; Name = bookmark.Name |> Option.defaultValue ""; EndsHere = false }
            | BookmarkEnd id ->
                // (A bookmark that ends where it starts: it stands before the next run.)
                let mutable at = -1
                for index in 0 .. pendingStarts.Count - 1 do
                    if at < 0 && pendingStarts.[index].Id = id && not pendingStarts.[index].EndsHere then at <- index
                if at >= 0 then pendingStarts.[at] <- { pendingStarts.[at] with EndsHere = true }
                else addEnd (ModelText.Ends id)
            | RenderedPageBreak ->
                pendingRenderedPageBreak <- true
            | _ ->
                output.Add(attachStart run)
                pendingStarts.Clear()
                pendingRenderedPageBreak <- false
        // (What is opened after the last run.)
        for start in pendingStarts do
            addEnd (ModelText.Opens start)

        output |> Seq.toList

    /// A page or column break (w:br w:type="page" / "column") with text, a picture or a field of the
    /// paragraph before it and after it breaks the page (the column) there: the text after it starts
    /// the next page (column) as a line of the same paragraph (Word 2019: fixtures page-top-space
    /// brkmid, inline-break). It stays an object of no size in the paragraph's text,
    /// written back as it was read; the others are the paragraph's breaks as the model has them
    /// (paragraphPageBreaks: before its text, or after it).
    let private placeInlineBreaks (runs: Run list) =
        let isBreak (run: Run) =
            match run.Content with
            | PageBreak -> true
            | Image image -> image.Opaque |> Option.exists (fun opaque -> isInlineBreakKind opaque.Kind)
            | _ -> false
        let isContent (run: Run) =
            not (isBreak run)
            && (match run.Content with
                | BookmarkStart _ | BookmarkEnd _ | RenderedPageBreak | FieldBoundary _ | Text "" -> false
                | _ -> true)
        if not (runs |> List.exists (fun run -> match run.Content with Image _ -> isBreak run | _ -> false)) then runs
        else
            let items = List.toArray runs
            let mutable contentBefore = false
            [ for index in 0 .. items.Length - 1 do
                let run = items.[index]
                match run.Content with
                | Image _ when isBreak run ->
                    let contentAfter = items |> Array.skip (index + 1) |> Array.exists isContent
                    if contentBefore && contentAfter then yield run else yield { run with Content = PageBreak }
                | _ ->
                    if isContent run then contentBefore <- true
                    yield run ]

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
        | Some style when ValueText.StyleName.isTableOfContents style ->
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
              HyperlinkContext = None
              Properties = props
              BookmarkStarts = []
              BookmarkEnds = []
              RenderedPageBreakBefore = false
              RevisionId = revId
              RevisionType = None
              RevisionAuthor = None
              RevisionDate = None; InsertedRevision = None; MoveName = None; Checkbox = None; ContentControls = []
              Source = (match rNode with XmlDsl.Element (_, _, _, span) -> span | _ -> XmlDsl.Span.none) }

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
                        // (As an object of no size for now: one with text on both sides of it in its
                        // paragraph stays one, the others are made the paragraph's breaks - placeInlineBreaks.)
                        | Some ("page" | "column" as kind) ->
                            Some (makeRun (Image { blankImageData "" with
                                                    Width = Some 0.0; Height = Some 0.0
                                                    Opaque = Some { Xml = xmlToString child; Kind = inlineBreakKind kind; ParagraphLevel = false; ShapeCandidate = false } }) |> RunToken)
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
                        | Some "begin" -> Some (FieldBeginToken (makeRun (FieldBoundary ModelText.Begin)))
                        | Some "separate" -> Some (FieldSeparateToken (makeRun (FieldBoundary ModelText.Separate)))
                        | Some "end" -> Some (FieldEndToken (makeRun (FieldBoundary ModelText.End)))
                        | _ -> None
                    | "instrText" -> Xml.text child |> Option.map (fun code -> FieldInstructionToken (makeRun (FieldBoundary (ModelText.Instruction code))))
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
        tokensFromRun rNode |> runsFromTokens |> attachRunMarkers

    let private tokenHasField = function
        | FieldBeginToken _
        | FieldSeparateToken _
        | FieldEndToken _
        | FieldInstructionToken _ -> true
        | RunToken run -> match run.Content with FieldRef _ -> true | _ -> false

    let private hyperlinkRunProps (props: RunProperties) =
        match props.Style with
        | Some _ -> props
        | None -> { props with Style = Some "Hyperlink" }

    let private hyperlinkToken rId anchorTarget token =
        let withLink (run: Run) =
            let props = hyperlinkRunProps run.Properties
            { run with Properties = props
                       HyperlinkContext = Some { Text = ""; Target = anchorTarget; RelationshipId = rId; Properties = props } }
        match token with
        | RunToken run -> RunToken (withLink run)
        | FieldBeginToken run -> FieldBeginToken (withLink run)
        | FieldSeparateToken run -> FieldSeparateToken (withLink run)
        | FieldEndToken run -> FieldEndToken (withLink run)
        | FieldInstructionToken run -> FieldInstructionToken (withLink run)

    let private tokensFromHyperlink (hlNode: XmlNode) : RunToken list =
        let rId = Xml.attr "r:id" hlNode
        // A link to a bookmark of the document itself (w:anchor): "#name".
        let anchorTarget = attr "anchor" hlNode |> Option.filter (fun name -> name <> "") |> Option.map (fun name -> "#" + name)
        let tokens =
            Xml.children hlNode |> List.collect (fun child ->
                match Xml.name child |> Option.map Ns.stripNs with
                | Some "r" -> tokensFromRun child
                | Some "fldSimple" -> simpleFieldRuns child |> List.map RunToken
                | _ -> [])

        if tokens |> List.exists tokenHasField then
            tokens |> List.map (hyperlinkToken rId anchorTarget)
        else
            let runs = tokens |> runsFromTokens |> attachRunMarkers

            // A link's runs are runs of the link: each with its own text and properties (a part of
            // a link that is bold is bold). A link without text is one run without text.
            let defaultProps = {
                Style = Some "Hyperlink"
                Bold = None; Italic = None; Underline = None; UnderlineStyle = None; UnderlineColor = None
                AllCaps = None; SmallCaps = None
                FontSize = None; CharacterSpacing = None; FontFamily = None; ThemeFont = None; Color = None
                ThemeColor = None; ThemeTint = None; ThemeShade = None; Highlight = None; Strike = None
                DoubleStrike = None; Hidden = None; Position = None; EmphasisMark = None; VertAlign = None
                EastAsiaLanguage = None
            }
            // (The w:r a link run is of, where it is one alone: docs/save-architecture.md.)
            let linkRun (text: string) (props: RunProperties) (source: XmlDsl.Span) (renderedPageBreak: bool) =
                RunToken {
                    Content = Link {
                        Text = text
                        Target = anchorTarget  // The relationship target is resolved later
                        RelationshipId = rId
                        Properties = props
                    }
                    Properties = props
                    BookmarkStarts = []
                    BookmarkEnds = []
                    RenderedPageBreakBefore = renderedPageBreak
                    RevisionId = None
                    RevisionType = None
                    RevisionAuthor = None
                    RevisionDate = None; InsertedRevision = None; MoveName = None; Checkbox = None
                    HyperlinkContext = None; ContentControls = []; Source = source
                }
            // (Runs that are formatted alike are one run, as they were before each was read on its own: of the
            // w:r of the first, where it is of that one alone. A run that Word's mark of a page it started stands
            // before is a run of its own, with the mark: see attachRunMarkers.)
            let texts =
                runs
                |> List.choose (fun r -> match r.Content with | Text t when t <> "" -> Some (t, hyperlinkRunProps r.Properties, r.Source, r.RenderedPageBreakBefore) | _ -> None)
                |> List.fold (fun (acc: (string * RunProperties * XmlDsl.Span * bool) list) (text, props, source, mark) ->
                    match acc with
                    | (previous, previousProps, previousSource, previousMark) :: rest when previousProps = props && not mark ->
                        (previous + text, props, (if previousSource = source then source else XmlDsl.Span.none), previousMark) :: rest
                    | _ -> (text, props, source, mark) :: acc) []
                |> List.rev
            match texts with
            | [] -> [ linkRun "" (runs |> List.tryHead |> Option.map (fun r -> hyperlinkRunProps r.Properties) |> Option.defaultValue defaultProps) XmlDsl.Span.none false ]
            | _ -> texts |> List.map (fun (text, props, source, mark) -> linkRun text props source mark)

    // Hyperlink extraction
    let hyperlink (hlNode: XmlNode) : Run list =
        tokensFromHyperlink hlNode |> runsFromTokens |> attachRunMarkers

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
    let private contentControlProps (sdtNode: XmlNode) : ContentControlLayer =
        let props = Ns.childNs "sdtPr" sdtNode
        let group = match sdtNode with XmlDsl.Element (_, _, _, span) -> string span.Start | _ -> ""
        let endProps = Ns.childNs "sdtEndPr" sdtNode |> Option.map xmlToString |> Option.defaultValue ""
        { Group = group
          Id = props |> Option.bind (Ns.childNs "id") |> Option.bind (attr "val")
          Tag = props |> Option.bind (Ns.childNs "tag") |> Option.bind (attr "val")
          Alias = props |> Option.bind (Ns.childNs "alias") |> Option.bind (attr "val")
          Xml = props |> Option.map (fun props -> xmlToString props + endProps) }

    let private paragraphWithRenderedBreakCounter (renderedBreakCounter: int ref option) (pNode: XmlNode) : Paragraph =
        let renderedPageBreaks = Ns.descendantsNs "lastRenderedPageBreak" pNode
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
            // (A column break goes on in the next column; in one column it is a page break.)
            let columnBreak = Ns.descendantsNs "br" pNode |> List.exists (fun node -> attr "type" node = Some "column")
            // (Not a page break: Word does not break the page where it last did - fixtures heading-mark
            // stale-mark: a Heading 1 that starts with the mark stays in the middle of the page.)
            { rawProps with RenderedPageBreakBefore = renderedPageBreakBefore; ColumnBreak = columnBreak }

        let rec tokensFromNode node =
            Xml.children node
            |> List.collect (fun child ->
                match Xml.name child with
                | Some name when Ns.stripNs name = "r" ->
                    tokensFromRun child
                | Some name when Ns.stripNs name = "hyperlink" ->
                    tokensFromHyperlink child
                | Some name when Ns.stripNs name = "fldSimple" ->
                    simpleFieldRuns child |> List.map RunToken
                | Some name when Ns.stripNs name = "oMath" || Ns.stripNs name = "oMathPara" ->
                    [RunToken { Content = EquationXml (xmlToString child); Properties = runProps child; BookmarkStarts = []; BookmarkEnds = []; RenderedPageBreakBefore = false; RevisionId = None; RevisionType = None; RevisionAuthor = None; RevisionDate = None; InsertedRevision = None; MoveName = None; Checkbox = None; HyperlinkContext = None; ContentControls = []; Source = XmlDsl.Span.none }]
                | Some name when Ns.stripNs name = "bookmarkStart" ->
                    attr "id" child
                    |> Option.map (fun id -> RunToken { Content = BookmarkStart { Id = id; Name = attr "name" child }; Properties = runProps child; BookmarkStarts = []; BookmarkEnds = []; RenderedPageBreakBefore = false; RevisionId = None; RevisionType = None; RevisionAuthor = None; RevisionDate = None; InsertedRevision = None; MoveName = None; Checkbox = None; HyperlinkContext = None; ContentControls = []; Source = XmlDsl.Span.none })
                    |> Option.toList
                | Some name when Ns.stripNs name = "bookmarkEnd" ->
                    attr "id" child
                    |> Option.map (fun id -> RunToken { Content = BookmarkEnd id; Properties = runProps child; BookmarkStarts = []; BookmarkEnds = []; RenderedPageBreakBefore = false; RevisionId = None; RevisionType = None; RevisionAuthor = None; RevisionDate = None; InsertedRevision = None; MoveName = None; Checkbox = None; HyperlinkContext = None; ContentControls = []; Source = XmlDsl.Span.none })
                    |> Option.toList
                // Transparent inline wrappers: their runs are ordinary paragraph content.
                | Some name when (match Ns.stripNs name with "smartTag" | "customXml" | "dir" | "bdo" -> true | _ -> false) ->
                    tokensFromNode child
                | Some name when Ns.stripNs name = "AlternateContent" ->
                    if isGraphicAlternateContent child then
                        // A drawing in two forms: kept whole rather than reduced to its fallback.
                        [RunToken { Content = Image (opaqueObject true child); Properties = runProps child; BookmarkStarts = []; BookmarkEnds = []; RenderedPageBreakBefore = false; RevisionId = None; RevisionType = None; RevisionAuthor = None; RevisionDate = None; InsertedRevision = None; MoveName = None; Checkbox = None; HyperlinkContext = None; ContentControls = []; Source = XmlDsl.Span.none }]
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
                    let revise (run: Run) =
                        // Nested insertion/deletion keeps both identities.
                        if revisionType = "insert" && run.RevisionType = Some "delete" then
                            { run with InsertedRevision = Some { Id = revisionId; Author = author; Date = date } }
                        elif revisionType = "delete" && run.RevisionType = Some "insert" then
                            { run with RevisionId = revisionId; RevisionType = Some "delete"
                                       RevisionAuthor = author; RevisionDate = date
                                       InsertedRevision = Some { Id = run.RevisionId; Author = run.RevisionAuthor; Date = run.RevisionDate } }
                        elif run.RevisionType.IsSome then run
                        else { run with RevisionId = revisionId; RevisionType = Some revisionType; RevisionAuthor = author; RevisionDate = date }
                    tokensFromNode child
                    |> List.map (function
                        | RunToken run -> RunToken (revise run)
                        | FieldBeginToken run -> FieldBeginToken (revise run)
                        | FieldSeparateToken run -> FieldSeparateToken (revise run)
                        | FieldEndToken run -> FieldEndToken (revise run)
                        | FieldInstructionToken run -> FieldInstructionToken (revise run))
                | Some name when Ns.stripNs name = "sdt" ->
                    // A w14:checkbox content control renders as a real checkbox
                    // glyph from its checked state (the stored glyph is often a
                    // symbol-font char that won't render); other controls just
                    // surface their inner content.
                    let layer = contentControlProps child
                    let checkbox = Ns.childNs "sdtPr" child |> Option.bind (fun pr -> Ns.descendantsNs "checkbox" pr |> List.tryHead)
                    let tokens =
                        match checkbox with
                        | Some cb ->
                            let isChecked =
                                Ns.descendantsNs "checked" cb |> List.tryHead
                                |> Option.bind (attrByLocalName "val")
                                |> Option.map (fun v -> v = "1" || v = "true")
                                |> Option.defaultValue false
                            let glyph = if isChecked then "☒" else "☐"
                            let checkboxRunSource =
                                Ns.childNs "sdtContent" child
                                |> Option.bind (fun content -> Ns.descendantsNs "r" content |> List.tryHead)
                                |> Option.map (function XmlDsl.Element (_, _, _, span) -> span | _ -> XmlDsl.Span.none)
                                |> Option.defaultValue XmlDsl.Span.none
                            let checkboxRun = Ns.childNs "sdtContent" child |> Option.bind (fun content -> Ns.descendantsNs "r" content |> List.tryHead)
                            [RunToken { Content = Text glyph; Properties = runProps (checkboxRun |> Option.defaultValue child); BookmarkStarts = []; BookmarkEnds = []; RenderedPageBreakBefore = false; RevisionId = None; RevisionType = None; RevisionAuthor = None; RevisionDate = None; InsertedRevision = None; MoveName = None; Checkbox = Some isChecked; Source = checkboxRunSource; HyperlinkContext = None; ContentControls = [] }]
                        | None ->
                            Ns.childNs "sdtContent" child
                            |> Option.map tokensFromNode
                            |> Option.defaultValue []
                    tokens |> List.map (function
                        | RunToken run -> RunToken { run with ContentControls = layer :: run.ContentControls }
                        | FieldBeginToken run -> FieldBeginToken { run with ContentControls = layer :: run.ContentControls }
                        | FieldSeparateToken run -> FieldSeparateToken { run with ContentControls = layer :: run.ContentControls }
                        | FieldEndToken run -> FieldEndToken { run with ContentControls = layer :: run.ContentControls }
                        | FieldInstructionToken run -> FieldInstructionToken { run with ContentControls = layer :: run.ContentControls })
                | _ -> [])

        // Extract runs and hyperlinks
        let runs = tokensFromNode pNode |> runsFromTokens |> attachRunMarkers |> placeInlineBreaks |> normalizeTocRuns props

        // A tracked paragraph mark: w:ins / w:del among the mark's run properties.
        let markRevision =
            wElem "pPr" pNode
            |> Option.bind (wElem "rPr")
            |> Option.bind (fun markProps ->
                [ "ins", "insert"; "del", "delete" ]
                |> List.tryPick (fun (elementName, revisionType) ->
                    wElem elementName markProps
                    |> Option.map (fun node -> revisionType, { Id = attr "id" node; Author = attr "author" node; Date = attr "date" node })))

        { Runs = runs; Properties = props; MarkRevision = markRevision
          Source = (match pNode with XmlDsl.Element (_, _, _, span) -> span | _ -> XmlDsl.Span.none) }

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
        let rec childrenIn layers node =
            Xml.children node |> List.collect (fun child ->
                match Xml.name child |> Option.map Ns.stripNs with
                | Some "sdt" ->
                    Ns.childNs "sdtContent" child
                    |> Option.map (childrenIn (layers @ [ contentControlProps child ]))
                    |> Option.defaultValue []
                | Some "customXml" -> childrenIn layers child
                | _ -> [ child, layers ])
        for child, controls in childrenIn [] tcNode do
            match Xml.name child |> Option.map Ns.stripNs with
            | Some "p" ->
                let para = paragraph child
                paragraphs <- (if controls.IsEmpty then para else { para with Properties = { para.Properties with EnclosingCellControls = controls } }) :: paragraphs
                paragraphCount <- paragraphCount + 1
            | Some "tbl" ->
                nestedTables <- table child :: nestedTables
                nestedAnchors <- paragraphCount :: nestedAnchors
            | _ -> ()
        {
            Paragraphs = List.rev paragraphs
            NestedTables = List.rev nestedTables
            NestedTableAnchors = List.rev nestedAnchors
            EnclosingControls = []
            Properties = tableCellProps tcNode
            Source = (match tcNode with XmlDsl.Element (_, _, _, span) -> span | _ -> XmlDsl.Span.none)
        }

    and tableRow (trNode: XmlNode) : TableRow =
        let rec cellsIn layers node =
            Xml.children node
            |> List.collect (fun child ->
                match Xml.name child |> Option.map Ns.stripNs with
                | Some "tc" ->
                    let cell = tableCell child
                    [ if List.isEmpty layers then cell else { cell with EnclosingControls = layers } ]
                | Some "sdt" ->
                    Ns.childNs "sdtContent" child
                    |> Option.map (cellsIn (layers @ [ contentControlProps child ]))
                    |> Option.defaultValue []
                | Some "customXml" -> cellsIn layers child
                | _ -> [])
        let cells = cellsIn [] trNode
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
            Source = (match trNode with XmlDsl.Element (_, _, _, span) -> span | _ -> XmlDsl.Span.none)
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
            ContentControlId = None
            ContentControlTag = None
            ContentControlAlias = None
            ContentControlXml = None
            ContentControlGroup = None
            ContentControlAncestors = []
            Look = tableLook
            Alignment = tblPr |> Option.bind (wElem "jc") |> Option.bind (attr "val")
            Indent = tblPr |> Option.bind (wElem "tblInd") |> Option.bind (fun node ->
                match attr "type" node with
                | None | Some "dxa" -> attr "w" node
                | _ -> None)
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
    let private applyContentControl (layer: ContentControlLayer) = function
        | Para para when para.Properties.ContentControlGroup.IsNone ->
            Para {
                para with
                    Properties = {
                        para.Properties with
                            ContentControlId = layer.Id; ContentControlTag = layer.Tag; ContentControlAlias = layer.Alias
                            ContentControlXml = layer.Xml; ContentControlGroup = Some layer.Group
                    }
            }
        | Para para ->
            Para { para with Properties = { para.Properties with ContentControlAncestors = layer :: para.Properties.ContentControlAncestors } }
        | Tbl table when table.ContentControlGroup.IsNone ->
            Tbl { table with ContentControlId = layer.Id; ContentControlTag = layer.Tag; ContentControlAlias = layer.Alias
                             ContentControlXml = layer.Xml; ContentControlGroup = Some layer.Group }
        | Tbl table -> Tbl { table with ContentControlAncestors = layer :: table.ContentControlAncestors }

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
            // (Its rows start pages too: the paragraphs after it are counted from Word's page.)
            renderedBreakCounter.Value <- renderedBreakCounter.Value + List.length (Ns.descendantsNs "lastRenderedPageBreak" child)
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

    /// A paragraph that ends a section is read as it is written: the w:type of
    /// its w:sectPr says how the section it ends starts, and nothing of the
    /// section after it. (An empty paragraph ending a section that starts on a
    /// new page, followed by one ending a continuous section, used to lose its
    /// break here - the first section then took the columns and the header of
    /// the second.)
    let normalizedSectionBreak (item: DocumentContent) (next: DocumentContent option) : DocumentContent =
        ignore next
        item

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
            // (On one element a theme face beats the named one, as in Word: fixtures rfonts-precedence.)
            FontFamily =
                match Ns.childNs "rFonts" rPr |> Option.bind (Docx.attr "asciiTheme") with
                | Some _ -> None
                | None -> Ns.childNs "rFonts" rPr |> Option.bind (Docx.attr "ascii")
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
            EastAsiaLanguage = Ns.childNs "lang" rPr |> Option.bind (Docx.attr "eastAsia")
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
            ContentControlXml = None
            ContentControlGroup = None
            ContentControlAncestors = []
            EnclosingCellControls = []
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
            OverflowPunctuation = Docx.onOff "overflowPunct" pPr
            BackgroundColor = Ns.childNs "shd" pPr |> Option.bind (Docx.attr "fill")
            ShadingPattern = Ns.childNs "shd" pPr |> Option.bind (Docx.attr "val")
            ShadingForeground = Ns.childNs "shd" pPr |> Option.bind (Docx.attr "color")
            ParagraphRunProperties = Ns.childNs "rPr" pPr |> Option.map (fun _ -> Docx.runProps pPr)
            ParagraphBorders = Docx.paragraphBorders pPr
            ListInfo = listInfo
            SectionBreak = None
            SectionGeometry = None
            SectionColumns = None
            SectionPageNumberStart = None; SectionPageNumberFormat = None
            SectionVerticalAlign = None
            ColumnBreak = false
            RenderedPageBreakBefore = None
            HeaderReferences = []
            FooterReferences = []
            DropCap = None
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
            // (A w:pPrDefault, even an empty one, says the paragraph defaults: what it leaves out is
            // Word's own - no space before or after a paragraph. Without one Word adds its own
            // space after; a template's empty w:pPrDefault/ was laid out with it, 8pt too much.)
            let pPr =
                Ns.childNs "pPrDefault" defaults
                |> Option.map (fun pPrDefault ->
                    let props = Ns.childNs "pPr" pPrDefault |> Option.defaultValue XmlDsl.Empty |> parseParagraphPropsFromStyle
                    let unstated = props.SpacingAfter.IsNone && props.SpacingAfterAuto.IsNone
                    { props with
                        SpacingBefore = (if props.SpacingBefore.IsNone && props.SpacingBeforeAuto.IsNone then Some "0" else props.SpacingBefore)
                        SpacingAfter = (if unstated then Some "0" else props.SpacingAfter) })
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
        |> Option.map (fun value -> ValueText.Color.withHash (value.ToUpper()))

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
        Suffix: string option
        Justification: string option
        Legal: bool
        Restart: int option
        MarkFont: string option
        MarkBold: bool option
        MarkItalic: bool option
        MarkColor: string option
        MarkSize: string option
    }

    /// A w:lvl (of an abstract numbering, or of a w:lvlOverride).
    let private parseLevel (level: XmlNode) : LevelDefinition =
        let intOf (name: string) =
            Ns.childNs name level
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
        let rPr = Ns.childNs "rPr" level
        let mark = rPr |> Option.map (fun _ -> Docx.runProps level)
        {
            ParagraphStyle = Ns.childNs "pStyle" level |> Option.bind (Docx.attr "val")
            NumberFormat = Ns.childNs "numFmt" level |> Option.bind (Docx.attr "val") |> Option.defaultValue "decimal"
            LevelText = Ns.childNs "lvlText" level |> Option.bind (Docx.attr "val")
            Start = intOf "start"
            IndentLeft = ind |> Option.bind (fun node -> Docx.attr "left" node |> Option.orElse (Docx.attr "start" node))
            HangingIndent = ind |> Option.bind (Docx.attr "hanging")
            FirstLineIndent = ind |> Option.bind (Docx.attr "firstLine")
            RightTabPosition = tabPosition
            Suffix = Ns.childNs "suff" level |> Option.bind (Docx.attr "val")
            Justification = Ns.childNs "lvlJc" level |> Option.bind (Docx.attr "val")
            Legal = Docx.onOff "isLgl" level |> Option.defaultValue false
            Restart = intOf "lvlRestart"
            MarkFont = mark |> Option.bind (fun props -> props.FontFamily)
            MarkBold = mark |> Option.bind (fun props -> props.Bold)
            MarkItalic = mark |> Option.bind (fun props -> props.Italic)
            MarkColor = mark |> Option.bind (fun props -> props.Color)
            MarkSize = mark |> Option.bind (fun props -> props.FontSize)
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
                            |> Option.map (fun ilvl -> int ilvl, parseLevel level))
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
                            |> Option.map (fun level -> int ilvl, parseLevel level)))
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
                    Suffix = def.Suffix
                    Justification = def.Justification
                    Legal = def.Legal
                    Restart = def.Restart
                    MarkFont = def.MarkFont
                    MarkBold = def.MarkBold
                    MarkItalic = def.MarkItalic
                    MarkColor = def.MarkColor
                    MarkSize = def.MarkSize
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

/// The words of an attribute's value that is a list (mc:Ignorable="w14 wp14", Requires="wps").
let private valueWords (value: string) = ValueText.words value

/// `xml` (one element, serialised) with the namespace declarations of the part's root element it
/// relies on added to its own root (XmlText.selfContainedXml).
let selfContainedXml (rootAttributes: Map<string, string>) (xml: string) : string = XmlText.selfContainedXml rootAttributes xml

/// Make the XML of every opaque object in `blocks` self-contained (see selfContainedXml).
let completeOpaqueObjects (root: XmlNode) (blocks: DocumentContent list) : DocumentContent list =
    match root with
    | XmlDsl.Element(_, attrs, _, _) ->
        // A control can inherit declarations from any ancestor, not only the
        // part root. Its group is the source offset of the enclosing w:sdt.
        let rec controlScopes inherited node =
            match node with
            | XmlDsl.Element(name, own, children, span) ->
                let scope =
                    own
                    |> Map.fold (fun scope key value ->
                        if key = "mc:Ignorable" || (Ns.declaredPrefix key).IsSome then Map.add key value scope
                        else scope) inherited
                let nested = children |> List.collect (controlScopes scope)
                if Ns.stripNs name = "sdt" then (string span.Start, scope) :: nested else nested
            | _ -> []
        let scopes = lazy (controlScopes attrs root |> Map.ofList)
        let complete group xml =
            match xml with
            | None -> None
            | Some value ->
                let scope = group |> Option.bind (fun key -> Map.tryFind key scopes.Value) |> Option.defaultValue attrs
                Some (selfContainedXml scope value)
        let completedLayers = lazy (System.Collections.Generic.Dictionary<ContentControlLayer, ContentControlLayer>())
        let completeLayer (layer: ContentControlLayer) =
            match Portable.Dict.tryFind layer completedLayers.Value with
            | ValueSome completed -> completed
            | _ ->
                let completed = { layer with Xml = complete (Some layer.Group) layer.Xml }
                completedLayers.Value.Add(layer, completed)
                completed
        blocks
        |> mapBlockParagraphs (fun para ->
            let props = para.Properties
            { para with
                Properties = {
                    props with
                        ContentControlXml = complete props.ContentControlGroup props.ContentControlXml
                        ContentControlAncestors = props.ContentControlAncestors |> List.map completeLayer
                } })
        |> List.map (function
            | Tbl table ->
                Tbl {
                    table with
                        ContentControlXml = complete table.ContentControlGroup table.ContentControlXml
                        ContentControlAncestors = table.ContentControlAncestors |> List.map completeLayer
                }
            | block -> block)
        |> mapBlockRuns (fun run ->
            let run =
                if run.ContentControls.IsEmpty then run
                else { run with ContentControls = run.ContentControls |> List.map completeLayer }
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
        | XmlDsl.Element (name, _, children, _) ->
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
    let resolve (link: Hyperlink) =
        match link.RelationshipId |> Option.bind (fun rId -> Map.tryFind rId relMap) with
        | Some target ->
            // With a w:anchor as well, the anchor is the fragment of the target.
            let fragment = match link.Target with | Some (anchor: string) when OpcXml.isAnchorTarget anchor && not (target.Contains "#") -> anchor | _ -> ""
            { link with Target = Some (target + fragment) }
        | None -> link
    let content =
        match run.Content with
        | Link link -> Link (resolve link)
        | other -> other
    { run with Content = content; HyperlinkContext = run.HyperlinkContext |> Option.map resolve }

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
        // (Scanned, as the body is: its paragraphs know where in the part they were read from.)
        let doc, _ = Ns.parseWithLocalNames xml
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
            ItemChunks = fst (Docx.bodyContentChunks doc)
            ChunkCount = snd (Docx.bodyContentChunks doc)
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
                ItemChunks = fst (Docx.bodyContentChunks commentNode)
                ChunkCount = snd (Docx.bodyContentChunks commentNode)
            })

    let parseComments (xml: string) : Comment list =
        // (Scanned, as the body is: its paragraphs know where in the part they were read from.)
        let doc, _ = Ns.parseWithLocalNames xml
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
                ItemChunks = fst (Docx.bodyContentChunks noteNode)
                ChunkCount = snd (Docx.bodyContentChunks noteNode)
            })

    let parseNote (noteNode: XmlNode) : Note option =
        let partKey = if Xml.name noteNode |> Option.map Ns.stripNs = Some "endnote" then endnotesPartKey else footnotesPartKey
        parseNoteOfPart partKey noteNode

    let parseNotes (xml: string) : Note list =
        // (Scanned, as the body is: its paragraphs know where in the part they were read from.)
        let doc, _ = Ns.parseWithLocalNames xml
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

    let private parseFloat (value: string) =
        match System.Double.TryParse(value) with
        | true, parsed -> Some parsed
        | false, _ -> None

    let private shapeNumberFromStyle name (style: string) =
        ValueText.Css.valueOf name style
        |> Option.bind (fun raw ->
            match ValueText.Css.withUnit [ "pt"; "px"; "in"; "mm"; "cm" ] raw with
            | Some (value, "pt") -> parseFloat value |> Option.map (fun pt -> pt * 96.0 / 72.0)
            | Some (value, "px") -> parseFloat value
            | Some (value, "in") -> parseFloat value |> Option.map (fun inches -> inches * 96.0)
            | Some (value, "mm") -> parseFloat value |> Option.map (fun mm -> mm * 96.0 / 25.4)
            | Some (value, _) -> parseFloat value |> Option.map (fun cm -> cm * 96.0 / 2.54)
            | None -> parseFloat raw)

    /// A colour as VML writes it: "#rgb", "#rrggbb", a name, each with or without a palette entry after it ("silver [3209]").
    let private normalizeColor (value: string option) = ValueText.Color.ofVml value

    /// The value of a property of a VML style ("name:value;name:value").
    let private styleText (name: string) (style: string) : string option = ValueText.Css.valueOfLower name style

    /// A VML length in pixels: points, inches, millimetres, centimetres, pixels; EMU without a unit.
    let private vmlLength (text: string) : float option =
        let raw = text.Trim().ToLowerInvariant()
        match ValueText.Css.withUnit [ "pt"; "in"; "mm"; "cm"; "px" ] raw with
        | Some (value, "pt") -> parseFloat value |> Option.map (fun value -> value * (96.0 / 72.0))
        | Some (value, "in") -> parseFloat value |> Option.map (fun value -> value * 96.0)
        | Some (value, "mm") -> parseFloat value |> Option.map (fun value -> value * (96.0 / 25.4))
        | Some (value, "cm") -> parseFloat value |> Option.map (fun value -> value * (96.0 / 2.54))
        | Some (value, _) -> parseFloat value |> Option.map (fun value -> value * 1.0)
        | None -> parseFloat raw |> Option.map (fun emu -> emu / 9525.0)

    let private xmlEscaped (text: string) =
        text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;")

    let private strokeWeightFromAttr (value: string option) =
        value
        |> Option.bind (fun raw ->
            let text = raw.Trim()
            match ValueText.Css.withUnit [ "pt"; "px" ] text with
            | Some (value, "pt") -> parseFloat value |> Option.map (fun pt -> pt * 96.0 / 72.0)
            | Some (value, _) -> parseFloat value
            | None -> parseFloat text)

    // --- DrawingML geometry/fill parsing (a:prstGeom / a:custGeom / a:gradFill) ---

    // --- DrawingML colours (a:srgbClr, a:schemeClr, a:sysClr, a:prstClr) ---
    // A theme colour cannot be known while the shape is read (the theme is a part
    // of its own): it is written "theme:<name>;<modifier>:<value>;..." and
    // resolveColor makes it "#RRGGBB" once the theme is at hand.

    let private colorElementNames = set [ "srgbClr"; "schemeClr"; "sysClr"; "prstClr"; "scrgbClr"; "hslClr" ]

    let private presetColors =
        Map.ofList [ "white", "FFFFFF"; "black", "000000"; "red", "FF0000"; "green", "008000"; "blue", "0000FF"; "yellow", "FFFF00"
                     "gray", "808080"; "grey", "808080"; "orange", "FFA500"; "purple", "800080"; "silver", "C0C0C0"; "navy", "000080" ]

    let private hslOf (r: float, g: float, b: float) =
        let r, g, b = r / 255.0, g / 255.0, b / 255.0
        let mx = max r (max g b)
        let mn = min r (min g b)
        let l = (mx + mn) / 2.0
        if mx = mn then 0.0, 0.0, l
        else
            let d = mx - mn
            let s = if l > 0.5 then d / (2.0 - mx - mn) else d / (mx + mn)
            let h =
                if mx = r then (g - b) / d + (if g < b then 6.0 else 0.0)
                elif mx = g then (b - r) / d + 2.0
                else (r - g) / d + 4.0
            h / 6.0, s, l

    let private rgbOfHsl (h: float, s: float, l: float) =
        if s = 0.0 then l * 255.0, l * 255.0, l * 255.0
        else
            let q = if l < 0.5 then l * (1.0 + s) else l + s - l * s
            let p = 2.0 * l - q
            let channel (t: float) =
                let t = if t < 0.0 then t + 1.0 elif t > 1.0 then t - 1.0 else t
                if t < 1.0 / 6.0 then p + (q - p) * 6.0 * t
                elif t < 0.5 then q
                elif t < 2.0 / 3.0 then p + (q - p) * (2.0 / 3.0 - t) * 6.0
                else p
            channel (h + 1.0 / 3.0) * 255.0, channel h * 255.0, channel (h - 1.0 / 3.0) * 255.0

    /// "#RRGGBB" of "RRGGBB" with DrawingML's colour modifiers applied in order
    /// (lumMod, lumOff, tint, shade; values in 1/1000 of a percent).
    let private withModifiers (hex: string) (modifiers: (string * float) list) : string option =
        let hex = hex.TrimStart('#')
        if hex.Length <> 6 then None
        else
            let channel at = System.Convert.ToInt32(hex.Substring(at, 2), 16) |> float
            let mutable rgb = channel 0, channel 2, channel 4
            for name, value in modifiers do
                let amount = value / 100000.0
                let r, g, b = rgb
                match name with
                | "lumMod" -> (let h, s, l = hslOf rgb in rgb <- rgbOfHsl (h, s, min 1.0 (l * amount)))
                | "lumOff" -> (let h, s, l = hslOf rgb in rgb <- rgbOfHsl (h, s, min 1.0 (max 0.0 (l + amount))))
                | "tint" -> rgb <- r + (255.0 - r) * (1.0 - amount), g + (255.0 - g) * (1.0 - amount), b + (255.0 - b) * (1.0 - amount)
                | "shade" -> rgb <- r * amount, g * amount, b * amount
                | _ -> ()
            let r, g, b = rgb
            let byte (v: float) = int (System.Math.Round(max 0.0 (min 255.0 v)))
            Some (sprintf "#%02X%02X%02X" (byte r) (byte g) (byte b))

    /// The colour of a DrawingML colour element: "#RRGGBB", or "theme:..." for a scheme colour.
    let private colorOfElement (node: XmlNode) : string option =
        let modifiers =
            Xml.children node
            |> List.choose (fun child ->
                match Xml.name child |> Option.map Ns.stripNs, Xml.attr "val" child |> Option.bind parseFloat with
                | Some name, Some value -> Some (name, value)
                | _ -> None)
        match Xml.name node |> Option.map Ns.stripNs with
        | Some "schemeClr" ->
            Xml.attr "val" node
            |> Option.map (fun name ->
                "theme:" + name + (modifiers |> List.map (fun (m, v) -> sprintf ";%s:%g" m v) |> String.concat ""))
        | Some "srgbClr" -> Xml.attr "val" node |> Option.bind (fun hex -> withModifiers hex modifiers)
        | Some "sysClr" ->
            Xml.attr "lastClr" node
            |> Option.orElse (match Xml.attr "val" node with Some "window" -> Some "FFFFFF" | Some _ -> Some "000000" | None -> None)
            |> Option.bind (fun hex -> withModifiers hex modifiers)
        | Some "prstClr" ->
            Xml.attr "val" node |> Option.bind (fun name -> Map.tryFind name presetColors)
            |> Option.bind (fun hex -> withModifiers hex modifiers)
        | _ -> None

    /// The first colour element in a node (a fill, a line, a gradient stop, a style reference).
    let private srgbOf (node: XmlNode) : string option =
        let rec find (nodes: XmlNode list) =
            match nodes with
            | [] -> None
            | head :: rest ->
                match Xml.name head |> Option.map Ns.stripNs with
                | Some name when colorElementNames.Contains name -> colorOfElement head
                | _ ->
                    match find (Xml.children head) with
                    | Some found -> Some found
                    | None -> find rest
        find (Xml.children node)

    /// A shape colour as read (see colorOfElement) as "#RRGGBB": a theme colour
    /// from `colors` (the theme's dk1, lt1, accent1, ...; tx1 / bg1 / tx2 / bg2 are
    /// dk1 / lt1 / dk2 / lt2). One that cannot be resolved is None.
    let resolveColor (colors: Map<string, string>) (value: string) : string option =
        if not (value.StartsWith "theme:") then Some value
        else
            let parts = value.Substring(6).Split(';')
            let name =
                match parts.[0] with
                | "tx1" -> "dk1" | "bg1" -> "lt1" | "tx2" -> "dk2" | "bg2" -> "lt2"
                | "phClr" -> ""
                | other -> other
            let modifiers =
                parts |> Array.skip 1 |> Array.toList
                |> List.choose (fun part ->
                    match part.Split(':') with
                    | [| m; v |] -> parseFloat v |> Option.map (fun v -> m, v)
                    | _ -> None)
            let fallback =
                match name with
                | "dk1" -> Some "000000" | "lt1" -> Some "FFFFFF" | "dk2" -> Some "44546A" | "lt2" -> Some "E7E6E6"
                | "accent1" -> Some "4472C4" | "accent2" -> Some "ED7D31" | "accent3" -> Some "A5A5A5" | "accent4" -> Some "FFC000"
                | "accent5" -> Some "5B9BD5" | "accent6" -> Some "70AD47" | "hlink" -> Some "0563C1" | "folHlink" -> Some "954F72"
                | _ -> None
            Map.tryFind name colors |> Option.orElse fallback
            |> Option.bind (fun hex -> withModifiers hex modifiers)

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

    /// a:xfrm flipH / flipV of the shape's own properties (wps:spPr).
    let private parseFlip (attrName: string) (shapeNode: XmlNode) : bool =
        Ns.descendantsNs "spPr" shapeNode |> List.tryHead
        |> Option.bind (Ns.childNs "xfrm")
        |> Option.bind (Xml.attr attrName)
        |> Option.exists (fun value -> value = "1" || value.ToLowerInvariant() = "true")

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

    let private isVmlNode (node: XmlNode) = match node with XmlDsl.Element (name, _, _, _) -> Ns.hasPrefix "v" name | _ -> false

    let private parseOneShape (shapeNode: XmlNode) : Shape option =
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
            |> Option.map (fun inset -> ValueText.Css.items ',' inset |> Array.map vmlLength)
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
        // (A shape that is aligned in one direction and placed by an offset in the other has the offset
        // it states: the other is 0, and its alignment says where it stands. Both used to be dropped,
        // so a shape centred between the margins an inch down the page stood at the top of the page.)
        let position =
            match x, y with
            | None, None -> None
            | _ -> Some (x |> Option.defaultValue 0.0, y |> Option.defaultValue 0.0)
        // DrawingML xfrm fallback for size/offset (a:ext / a:off) when the wp:
        // anchor extent/position is absent (bare wsp shapes).
        let xfrm = Ns.descendantsNs "xfrm" shapeNode |> List.tryHead
        let extOf node tag = xfrm |> Option.bind (Ns.childNs node) |> Option.bind (Xml.attr tag) |> Option.bind parseFloat |> Option.map emuToPixels
        let width = width |> Option.orElseWith (fun () -> extOf "ext" "cx")
        let height = height |> Option.orElseWith (fun () -> extOf "ext" "cy")
        let spPr = shapeProperties shapeNode
        let line = spPr |> Option.bind (childNamed "ln")
        // (A DrawingML shape that states no fill or line of its own has those of the
        // theme it refers to in wps:style: a:fillRef / a:lnRef with a colour.)
        let styleReference (name: string) =
            Ns.descendantsNs "style" shapeNode |> List.tryHead
            |> Option.bind (Ns.childNs name)
            |> Option.filter (fun reference -> Xml.attr "idx" reference |> Option.bind parseFloat |> Option.exists (fun idx -> idx > 0.0))
            |> Option.bind srgbOf
        let statesOwnFill =
            spPr |> Option.exists (fun props -> [ "solidFill"; "gradFill"; "pattFill"; "blipFill"; "grpFill"; "noFill" ] |> List.exists (fun name -> (childNamed name props).IsSome))
        let statesOwnLine =
            line |> Option.exists (fun ln -> [ "solidFill"; "gradFill"; "pattFill"; "noFill" ] |> List.exists (fun name -> (childNamed name ln).IsSome))
        let fillColor =
            (Xml.attr "fillcolor" shapeNode |> normalizeColor)
            |> Option.orElseWith (fun () -> parseDmlFill shapeNode)
            |> Option.orElseWith (fun () -> if isVmlNode shapeNode || statesOwnFill then None else styleReference "fillRef")
        let borderColor =
            (Xml.attr "strokecolor" shapeNode |> normalizeColor)
            |> Option.orElseWith (fun () -> parseDmlLine shapeNode)
            |> Option.orElseWith (fun () -> if isVmlNode shapeNode || statesOwnLine then None else styleReference "lnRef")
        // A shape filled with a picture (spPr > a:blipFill): the picture in its box.
        let blipFillImage =
            spPr |> Option.bind (childNamed "blipFill")
            |> Option.bind (Ns.childNs "blip")
            |> Option.bind (fun blip -> Xml.attr "r:embed" blip |> Option.orElse (Xml.attr "embed" blip))
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
                | XmlDsl.Element ("v:oval", _, _, _) -> Some "ellipse"
                | XmlDsl.Element ("v:roundrect", _, _, _) -> Some "roundRect"
                | _ -> None)
        let isVml = match shapeNode with XmlDsl.Element (name, _, _, _) -> Ns.hasPrefix "v" name | _ -> false
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
                |> Option.bind (fun text -> match ValueText.Css.withUnit [ "fd" ] text with Some (value, _) -> parseFloat value |> Option.map (fun fd -> fd / 65536.0) | None -> parseFloat text)
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
            FlipH = (if isVml then fst (ValueText.Css.flip vmlStyle) else parseFlip "flipH" shapeNode)
            FlipV = (if isVml then snd (ValueText.Css.flip vmlStyle) else parseFlip "flipV" shapeNode)
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
            // (A DrawingML shape that states no fill and refers to none of the theme's - no fill in its
            // properties, no wps:style - has none; likewise its line. Such a text box used to be read as
            // one that says nothing, which is white with a grey outline: the editor's own.)
            NoFill =
                (spPr |> Option.exists (fun props -> (childNamed "noFill" props).IsSome))
                || (not isVml && spPr.IsSome
                    && (spPr |> Option.exists (fun props -> [ "solidFill"; "gradFill"; "pattFill"; "blipFill"; "grpFill" ] |> List.forall (fun name -> (childNamed name props).IsNone)))
                    && (Ns.descendantsNs "fillRef" shapeNode).IsEmpty)
                || vmlBool (Xml.attr "filled" shapeNode) = Some false
                // (The fill of WordArt is the colour of its letters.)
                || (txbxContent.IsNone && textPath.IsSome)
            NoLine =
                (line |> Option.exists (fun ln -> (childNamed "noFill" ln).IsSome))
                || (not isVml && spPr.IsSome && line.IsNone && (Ns.descendantsNs "lnRef" shapeNode).IsEmpty)
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
                |> Option.bind (fun text -> match ValueText.Css.withUnit [ "f" ] text with Some (value, _) -> parseFloat value |> Option.map (fun f -> f / 65536.0) | None -> parseFloat text)
                |> Option.filter (fun value -> value >= 0.0 && value < 1.0)
            Children = []
            ChildSpace = None
            ImageRelationshipId = blipFillImage
            SourceXml = None
        }

    // --- groups and drawing canvases (wpg:wgp, wpc:wpc, v:group) ------------

    /// Where a group's child coordinates are in pixels from the group's
    /// top-left corner: x -> (x - OriginX) * ScaleX, y -> (y - OriginY) * ScaleY.
    type private ChildMapping = { OriginX: float; OriginY: float; ScaleX: float; ScaleY: float }

    /// EMU from the top-left corner (a drawing canvas's shapes).
    let private emuMapping = { OriginX = 0.0; OriginY = 0.0; ScaleX = 1.0 / 9525.0; ScaleY = 1.0 / 9525.0 }

    let private isTrue (value: string) = value = "1" || value.ToLowerInvariant() = "true"

    /// The a:xfrm of an element's own properties (wps:spPr, pic:spPr,
    /// wpg:grpSpPr) - not one of a shape in it.
    let private ownTransform (element: XmlNode) : XmlNode option =
        Xml.children element
        |> List.tryFind (fun child -> match Xml.name child |> Option.map Ns.stripNs with Some "spPr" | Some "grpSpPr" -> true | _ -> false)
        |> Option.bind (Ns.childNs "xfrm")

    /// A point or size of an a:xfrm (a:off x / y, a:ext cx / cy, a:chOff, a:chExt).
    let private transformPair (childName: string) (xAttr: string) (yAttr: string) (xfrm: XmlNode option) : (float * float) option =
        xfrm
        |> Option.bind (Ns.childNs childName)
        |> Option.bind (fun point ->
            match Xml.attr xAttr point |> Option.bind parseFloat, Xml.attr yAttr point |> Option.bind parseFloat with
            | Some x, Some y -> Some (x, y)
            | _ -> None)

    /// An element's box in its group's child coordinates: a:off and a:ext of its own a:xfrm.
    let private transformBox (xfrm: XmlNode option) : (float * float * float * float) option =
        match transformPair "off" "x" "y" xfrm, transformPair "ext" "cx" "cy" xfrm with
        | Some (x, y), Some (cx, cy) -> Some (x, y, cx, cy)
        | None, Some (cx, cy) -> Some (0.0, 0.0, cx, cy)
        | _ -> None

    let private transformRotation (xfrm: XmlNode option) : float =
        xfrm |> Option.bind (Xml.attr "rot") |> Option.bind parseFloat |> Option.map (fun r -> r / 60000.0) |> Option.defaultValue 0.0

    let private transformFlip (attrName: string) (xfrm: XmlNode option) : bool =
        xfrm |> Option.bind (Xml.attr attrName) |> Option.exists isTrue

    /// A DrawingML group's children in its box of `width` × `height` pixels:
    /// a:chOff / a:chExt of its a:xfrm map onto it (its a:off / a:ext when it
    /// states no child extent; EMU where an extent is 0).
    let private drawingMapping (xfrm: XmlNode option) (width: float) (height: float) : ChildMapping =
        let ox, oy =
            transformPair "chOff" "x" "y" xfrm |> Option.orElse (transformPair "off" "x" "y" xfrm) |> Option.defaultValue (0.0, 0.0)
        let cx, cy = transformPair "chExt" "cx" "cy" xfrm |> Option.orElse (transformPair "ext" "cx" "cy" xfrm) |> Option.defaultValue (0.0, 0.0)
        { OriginX = ox; OriginY = oy
          ScaleX = (if cx > 0.0 then width / cx else 1.0 / 9525.0)
          ScaleY = (if cy > 0.0 then height / cy else 1.0 / 9525.0) }

    /// The two numbers of a VML pair ("0,0", "21600,21600"); one it does not state is `fallback`'s.
    let private vmlPair (fallback: float * float) (value: string option) : float * float =
        match value with
        | Some text ->
            let parts = ValueText.Css.items ',' text |> Array.map (fun part -> parseFloat (part.Trim()))
            (if parts.Length > 0 then parts.[0] else None) |> Option.defaultValue (fst fallback),
            (if parts.Length > 1 then parts.[1] else None) |> Option.defaultValue (snd fallback)
        | None -> fallback

    /// A VML group's children in its box of `width` × `height` pixels: its
    /// coordorigin and coordsize (0,0 and 1000,1000 when it states none).
    let private vmlMapping (group: XmlNode) (width: float) (height: float) : ChildMapping =
        let ox, oy = vmlPair (0.0, 0.0) (Xml.attr "coordorigin" group)
        let cw, ch = vmlPair (1000.0, 1000.0) (Xml.attr "coordsize" group)
        { OriginX = ox; OriginY = oy
          ScaleX = (if cw <> 0.0 then width / cw else 1.0)
          ScaleY = (if ch <> 0.0 then height / ch else 1.0) }

    /// A VML element's box in its group's coordinates: left, top, width and height of its style.
    let private vmlBox (node: XmlNode) : float * float * float * float =
        let style = Xml.attr "style" node |> Option.defaultValue ""
        let number name = ValueText.Css.valueIgnoringCase name style |> Option.bind parseFloat |> Option.defaultValue 0.0
        number "left", number "top", number "width", number "height"

    let private mappedBox (mapping: ChildMapping) (x: float, y: float, w: float, h: float) =
        (x - mapping.OriginX) * mapping.ScaleX, (y - mapping.OriginY) * mapping.ScaleY, w * mapping.ScaleX, h * mapping.ScaleY

    /// A shape of a group at its place in the group (pixels from the group's top-left corner).
    let private placedIn (x: float, y: float, w: float, h: float) (shape: Shape) : Shape =
        { shape with
            Position = Some (x, y); Width = Some (max 0.0 w); Height = Some (max 0.0 h)
            AnchorBlock = None; Inline = false; Wrap = None; PositionHRelativeFrom = None; PositionVRelativeFrom = None
            PositionHAlign = None; PositionVAlign = None; BehindDoc = None; AllowOverlap = None; LayoutInCell = None; RelativeHeight = None
            DistanceTop = None; DistanceBottom = None; DistanceLeft = None; DistanceRight = None }

    /// What a group is of its own: no text, outline or fill (but a canvas's).
    let private asGroup (shape: Shape) : Shape =
        { shape with
            Content = []; Preset = None; CustomPath = []; CustomW = None; CustomH = None; Gradient = None
            FillColor = None; BorderColor = None; BorderWidth = None; NoFill = true; NoLine = true; LineDash = None
            InsetLeft = None; InsetTop = None; InsetRight = None; InsetBottom = None; VerticalAnchor = None
            AutoFit = false; NoTextWrap = false; FitText = false; Opacity = None }

    /// A picture of a group: the picture in its box, no fill or outline.
    let private asPicture (relationshipId: string) (shape: Shape) : Shape =
        { asGroup shape with ImageRelationshipId = Some relationshipId }

    /// The shapes of a group element (wpg:wgp / wpg:grpSp, wpc:wpc, v:group) in
    /// the order they are drawn, `mapping` its child coordinates: those
    /// Docx.isDrawnGroupChild says are (a VML shape the reader does not know
    /// the outline of and that has no text is left out).
    let rec private groupChildren (group: XmlNode) (mapping: ChildMapping) : Shape list =
        Xml.children group
        |> List.filter Docx.isDrawnGroupChild
        |> List.choose (fun node ->
            let vml = match node with XmlDsl.Element (name, _, _, _) -> Ns.hasPrefix "v" name | _ -> false
            match Xml.name node |> Option.map Ns.stripNs with
            | Some ("wsp" | "pic" | "grpSp" | "wgp" as kind) when not vml ->
                let xfrm = ownTransform node
                match transformBox xfrm with
                | None -> None
                | Some box ->
                    let (_, _, w, h) as placed = mappedBox mapping box
                    match kind with
                    | "wsp" -> parseOneShape node |> Option.map (placedIn placed)
                    | "pic" ->
                        Docx.pictureRelationshipOf node
                        |> Option.bind (fun relationshipId ->
                            parseOneShape node
                            |> Option.map (fun shape ->
                                { (placedIn placed shape |> asPicture relationshipId) with
                                    Rotation = transformRotation xfrm; FlipH = transformFlip "flipH" xfrm; FlipV = transformFlip "flipV" xfrm }))
                    | _ ->
                        // (A group in the group: its shapes in its own child coordinates.)
                        parseOneShape node
                        |> Option.map (fun shape ->
                            { (placedIn placed shape |> asGroup) with
                                Type = Some "group"
                                Rotation = transformRotation xfrm; FlipH = transformFlip "flipH" xfrm; FlipV = transformFlip "flipV" xfrm
                                Children = groupChildren node (drawingMapping xfrm w h)
                                ChildSpace = Some (w, h) })
            | Some local when vml ->
                let (_, _, w, h) as placed = mappedBox mapping (vmlBox node)
                match local with
                | "group" ->
                    // (Its turn and flips are those of its style: as parseOneShape reads a VML shape's.)
                    parseOneShape node
                    |> Option.map (fun shape ->
                        { (placedIn placed shape |> asGroup) with
                            Type = Some "group"; Children = groupChildren node (vmlMapping node w h); ChildSpace = Some (w, h) })
                | "line" ->
                    // (From one end to the other: a line across its box, mirrored when it rises.)
                    let x1, y1 = vmlPair (0.0, 0.0) (Xml.attr "from" node)
                    let x2, y2 = vmlPair (0.0, 0.0) (Xml.attr "to" node)
                    parseOneShape node
                    |> Option.map (fun shape ->
                        { (placedIn (mappedBox mapping (min x1 x2, min y1 y2, abs (x2 - x1), abs (y2 - y1))) shape |> asGroup) with
                            Type = Some "line"; Preset = Some "line"
                            BorderColor = shape.BorderColor |> Option.orElse (Some "#000000"); BorderWidth = shape.BorderWidth; NoLine = shape.NoLine
                            FlipH = ((x2 < x1) <> (y2 < y1)); FlipV = false })
                | "shape" | "rect" | "oval" | "roundrect" ->
                    match Docx.vmlPictureRelationshipOf node with
                    | Some relationshipId -> parseOneShape node |> Option.map (fun shape -> { (placedIn placed shape |> asPicture relationshipId) with Rotation = shape.Rotation; FlipH = shape.FlipH; FlipV = shape.FlipV })
                    | None -> parseOneShape node |> Option.map (placedIn placed)
                | _ -> None
            | _ -> None)

    /// A group or drawing canvas where its anchor (wp:anchor / wp:inline around
    /// a wpg:wgp or wpc:wpc) or its style (a VML group) puts it, as large as
    /// its extent, with its shapes in it; a canvas with its background
    /// (wpc:bg) and outline (wpc:whole), a group turned and flipped as its
    /// own a:xfrm says (a VML group: its style).
    let private parseGroupShape (root: XmlNode) (group: XmlNode) : Shape option =
        parseOneShape root
        |> Option.map (fun outer ->
            let vml = match root with XmlDsl.Element (name, _, _, _) -> Ns.hasPrefix "v" name | _ -> false
            let canvas = Xml.name group |> Option.map Ns.stripNs = Some "wpc"
            let xfrm = if vml || canvas then None else ownTransform group
            let extent pick = transformPair "ext" "cx" "cy" xfrm |> Option.map (fun pair -> pick pair / 9525.0)
            let width = outer.Width |> Option.orElse (extent fst) |> Option.defaultValue 0.0
            let height = outer.Height |> Option.orElse (extent snd) |> Option.defaultValue 0.0
            let mapping =
                if vml then vmlMapping group width height
                elif canvas then emuMapping
                else drawingMapping xfrm width height
            let canvasFill = if canvas then Ns.childNs "bg" group |> Option.bind (Ns.childNs "solidFill") |> Option.bind srgbOf else None
            let canvasLine = if canvas then Ns.childNs "whole" group |> Option.bind (Ns.childNs "ln") else None
            let lineColor = canvasLine |> Option.bind (Ns.childNs "solidFill") |> Option.bind srgbOf
            { asGroup outer with
                Type = Some (if canvas then "canvas" else "group")
                Width = Some width
                Height = Some height
                FillColor = canvasFill
                NoFill = canvasFill.IsNone
                BorderColor = lineColor
                NoLine = lineColor.IsNone
                BorderWidth = (if lineColor.IsSome then canvasLine |> Option.bind (Xml.attr "w") |> Option.bind parseFloat |> Option.map emuToPixels else None)
                Rotation = (if vml then outer.Rotation else transformRotation xfrm)
                FlipH = (if vml then outer.FlipH else transformFlip "flipH" xfrm)
                FlipV = (if vml then outer.FlipV else transformFlip "flipV" xfrm)
                // (A VML group is known by its id: see withShapeSources.)
                Name = (if vml then None else outer.Name)
                Children = groupChildren group mapping
                ChildSpace = (if canvas then None else Some (width, height)) })

    /// A shape of a shape root (Docx.modelledShapeRoots): one shape, or a
    /// group or drawing canvas with its shapes.
    let parseShape (shapeNode: XmlNode) : Shape option =
        match Docx.groupElement shapeNode with
        | Some group when Docx.isModelledGroup shapeNode -> parseGroupShape shapeNode group
        | _ -> parseOneShape shapeNode

    /// A shape's text with that of the shapes of it (a group's), changed by `change`.
    let rec mapContent (change: Paragraph list -> Paragraph list) (shape: Shape) : Shape =
        { shape with
            Content = change shape.Content
            Children = shape.Children |> List.map (mapContent change) }

    /// A shape's text and that of the shapes of it (a group's).
    let rec allContent (shape: Shape) : Paragraph list =
        shape.Content @ (shape.Children |> List.collect allContent)

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
let getMimeType (path: string) : string = ValueText.mediaTypeOf path

let private wordTwipsToPixels value =
    value
    |> Option.bind (fun (text: string) ->
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

/// w:defaultTabStop of the settings part, in pixels; None when it says none
/// (or none above zero).
let private parseDefaultTabStop (xml: string option) =
    xml
    |> Option.map XmlDsl.parse
    |> Option.bind (fun settings -> Ns.descendantsNs "defaultTabStop" settings |> List.tryHead)
    |> Option.bind (Docx.attr "val")
    |> wordTwipsToPixels
    |> Option.filter (fun width -> width > 0.0)

/// The compatibilityMode w:compatSetting of the settings part's w:compat.
let private parseCompatibilityMode (xml: string option) =
    xml
    |> Option.map XmlDsl.parse
    |> Option.bind (fun settings ->
        Ns.descendantsNs "compatSetting" settings
        |> List.tryFind (fun node -> Docx.attr "name" node = Some "compatibilityMode"))
    |> Option.bind (Docx.attr "val")
    |> Option.bind (fun value -> match System.Int32.TryParse value with | true, mode -> Some mode | _ -> None)

/// w:useFELayout of the settings part's w:compat.
let private parseUseFarEastLayout (xml: string option) =
    xml
    |> Option.map XmlDsl.parse
    |> Option.bind (fun settings -> Ns.descendantsNs "compat" settings |> List.tryHead)
    |> Option.bind (Docx.onOff "useFELayout")

/// w:characterSpacingControl of the settings part.
let private parseCharacterSpacingControl (xml: string option) =
    xml
    |> Option.map XmlDsl.parse
    |> Option.bind (fun settings -> Ns.descendantsNs "characterSpacingControl" settings |> List.tryHead)
    |> Option.bind (Docx.attr "val")

/// The settings part's hyphenation (see PageSetup.Hyphenation).
let private parseHyphenation (xml: string option) =
    xml
    |> Option.map XmlDsl.parse
    |> Option.bind (fun settings ->
        let on name = Ns.descendantsNs name settings |> List.tryHead |> Option.map (fun node -> Docx.attr "val" node |> Option.forall (fun v -> v <> "0" && v <> "false" && v <> "off"))
        let number name = Ns.descendantsNs name settings |> List.tryHead |> Option.bind (Docx.attr "val") |> Option.bind (fun v -> match System.Int32.TryParse v with | true, n -> Some n | _ -> None)
        if on "autoHyphenation" = Some true then
            Some (sprintf "%d %d %d" (number "hyphenationZone" |> Option.defaultValue -1) (number "consecutiveHyphenLimit" |> Option.defaultValue 0) (if on "doNotHyphenateCaps" = Some true then 1 else 0))
        else None)

let private parsePageSetup evenAndOddHeaders defaultTabStop compatibilityMode useFarEastLayout characterSpacingControl hyphenation (sectPr: XmlNode option) : PageSetup option =
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
            DefaultTabStop = defaultTabStop
            CompatibilityMode = compatibilityMode
            UseFarEastLayout = useFarEastLayout
            CharacterSpacingControl = characterSpacingControl
            Hyphenation = hyphenation
            Columns = columns
            SectionStart = Docx.getPropVal "type" "val" section
            HeaderReferences = Docx.headerFooterReferences "headerReference" section
            FooterReferences = Docx.headerFooterReferences "footerReference" section
            PageNumberStart =
                Docx.getPropVal "pgNumType" "start" section
                |> Option.bind (fun value -> match System.Int32.TryParse(value) with | true, parsed -> Some parsed | false, _ -> None)
            PageNumberFormat = Docx.getPropVal "pgNumType" "fmt" section
            VerticalAlign = Docx.getPropVal "vAlign" "val" section
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

let private lastPathSegment (path: string) = PartPath.fileName (normalizePath path)

let private parentPath (path: string) = PartPath.directory (normalizePath path)

let private resolveRelationshipTarget baseDir target =
    ZipInterop.resolveTarget baseDir target

/// The relationships part of the part at `partPath` (word/footnotes.xml -> word/_rels/footnotes.xml.rels).
let private partRelsPath (partPath: string) =
    let dir = parentPath partPath
    (if dir = "" then "" else dir + "/") + "_rels/" + lastPathSegment partPath + ".rels"

/// The package path of the part a relationship of the given type points at, or `fallback`.
let private partPathOfType (baseDir: string) (rels: Relationship list) (typeSuffix: string) (fallback: string) =
    rels
    |> List.tryFind (fun rel -> ValueText.isRelationshipType typeSuffix rel.Type && rel.TargetMode <> Some "External")
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

/// Relationship ids an XML fragment refers to (r:id, r:embed, r:link, r:dm, ..., o:relid), each once.
let relationshipIdsInXml (xml: string) : string list = OpcXml.relationshipIdsIn xml

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
    |> List.exists (fun suffix -> ValueText.isRelationshipType suffix relType)

/// The relationships of a regenerated part that are written back: everything
/// the writer does not produce itself. Pictures and hyperlinks are produced by
/// the writer from the model, so those are kept only when an opaque object of
/// the part refers to them.
let private keptRelationshipsReferenced (baseDir: string) (rels: Relationship list) (referenced: Set<string>) : PreservedRelationship list =
    rels
    |> List.filter (fun rel ->
        if isRegeneratedRelationshipType rel.Type then false
        elif ValueText.isRelationshipType "/image" rel.Type || ValueText.isRelationshipType "/hyperlink" rel.Type then Set.contains rel.Id referenced
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
    /// See PageSetup.DefaultTabStop.
    DefaultTabStop: float option
    /// See PageSetup.CompatibilityMode.
    CompatibilityMode: int option
    /// See PageSetup.UseFarEastLayout.
    UseFarEastLayout: bool option
    /// See PageSetup.CharacterSpacingControl.
    CharacterSpacingControl: string option
    /// See PageSetup.Hyphenation.
    Hyphenation: string option
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
let private mainRelationshipOf (source: PartSource) =
    let rootRels = source.Text "_rels/.rels" |> relationshipsOfXml
    let mainRelationship = rootRels |> List.tryFind (fun rel -> ValueText.isRelationshipType "/officeDocument" rel.Type && rel.TargetMode <> Some "External")
    let mainPath =
        mainRelationship
        |> Option.map (fun rel -> ZipInterop.resolveTarget "" rel.Target)
        |> Option.filter (fun path -> (source.Text path).IsSome)
        |> Option.defaultValue "word/document.xml"
    rootRels, mainRelationship, mainPath

/// The parts as Word shows them: the content controls bound to custom XML hold their nodes' values
/// (DataBinding), in every part but the custom XML parts themselves.
let withDataBindings (source: PartSource) : PartSource =
    let _, _, mainPath = mainRelationshipOf source
    match DataBinding.storesOf source.Text mainPath with
    | [] -> source
    | stores ->
        let resolved = System.Collections.Generic.Dictionary<string, string option>()
        { source with
            Text =
                fun path ->
                    match Portable.Dict.tryFind path resolved with
                    | ValueSome text -> text
                    | ValueNone ->
                        let text = source.Text path
                        let text =
                            if path.ToLowerInvariant().StartsWith "customxml/" || PartPath.extension path <> "xml" then text
                            else text |> Option.map (DataBinding.resolvePart stores)
                        resolved.[path] <- text
                        text }

let private readMainPart (source: PartSource) : MainPart option =
    let rootRels, mainRelationship, mainPath = mainRelationshipOf source
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
          DefaultTabStop = parseDefaultTabStop settingsXml
          CompatibilityMode = parseCompatibilityMode settingsXml
          UseFarEastLayout = parseUseFarEastLayout settingsXml
          CharacterSpacingControl = parseCharacterSpacingControl settingsXml
          Hyphenation = parseHyphenation settingsXml
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
            let part =
                // (A modelled shape has text or a geometry: see Docx.isModelledShape.)
                if not (XmlDsl.startTagsNamed [ "txbxContent"; "prstGeom"; "custGeom"; "textpath"; "wgp"; "wpc"; "group" ] value).IsEmpty then
                    let root = XmlDsl.parse value
                    let shapes =
                        Shapes.parseShapesOfPart root
                        |> List.map (fun shape ->
                            Shapes.mapContent (fun content -> content |> List.map Para |> completeOpaqueObjects root |> List.choose (function Para para -> Some para | Tbl _ -> None)) shape)
                        |> Shapes.withShapeSourcesOf part.Blocks
                    if not shapes.IsEmpty then partShapes.Add((rel.Id, shapes))
                    { part with
                        Blocks = dropModelledShapeRuns part.Blocks
                        Content = part.Content |> List.map Para |> dropModelledShapeRuns |> List.choose (function Para para -> Some para | Tbl _ -> None) }
                else part
            part, path, partRels)
    let headerParts = main.Rels |> List.filter (fun rel -> ValueText.isRelationshipType "/header" rel.Type) |> List.choose (loadHeaderFooter headerRefTypes)
    let footerParts = main.Rels |> List.filter (fun rel -> ValueText.isRelationshipType "/footer" rel.Type) |> List.choose (loadHeaderFooter footerRefTypes)
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

/// The relationship ids a shape refers to: the opaque objects of its text and
/// of its shapes' - and, of a group, its element's (its pictures: it is
/// written back as it was read).
let private shapeReferences (shape: Shape) : string list =
    opaqueReferences (Shapes.allContent shape |> List.map Para)
    @ (if shape.Children.IsEmpty then [] else shape.SourceXml |> Option.map relationshipIdsInXml |> Option.defaultValue [])

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
        keptRelationships "" (main.RootRels |> List.filter (fun rel -> not (ValueText.isRelationshipType "/officeDocument" rel.Type))) []
    let partKept =
        [ yield "document", keptRelationshipsReferenced docDir rels documentReferenced
          for part, path, partRels in headerParts @ footerParts do
              match part.RelationshipId with
              | Some partId ->
                  // (A group of the part is written back as it was read: its pictures stay with it.)
                  let groupXml =
                      Map.tryFind partId stories.PartShapes |> Option.defaultValue []
                      |> List.choose (fun shape -> if shape.Children.IsEmpty then None else shape.SourceXml)
                  yield partId, keptRelationships (parentPath path) partRels (opaqueXmlOfBlocks part.Blocks @ groupXml)
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
    let source = withDataBindings source
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
            |> parsePageSetup main.EvenAndOddHeaders main.DefaultTabStop main.CompatibilityMode main.UseFarEastLayout main.CharacterSpacingControl main.Hyphenation

        let revisions = Revisions.parseRevisionsWhere docHolds doc
        let shapes =
            // (A modelled shape has text or a geometry: see Docx.isModelledShape.)
            (if docHolds "txbxContent" || docHolds "prstGeom" || docHolds "custGeom" || docHolds "textpath" || docHolds "wgp" || docHolds "wpc" || docHolds "group" then Shapes.parseShapes doc else [])
            |> List.map (fun shape ->
                Shapes.mapContent (fun content -> content |> List.map Para |> completeOpaqueObjects doc |> List.choose (function Para para -> Some para | Tbl _ -> None)) shape)
            |> Shapes.withShapeSourcesOf completedContent

        let bodyChunks, bodyChunkCount =
            body |> Option.map Docx.bodyContentChunks |> Option.defaultValue ([], 0)
        let documentReferenced =
            opaqueReferences content @ (shapes |> List.collect shapeReferences) |> Set.ofList
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

/// Where the child elements of the body are in the text of the main part,
/// read one at a time by the scanner's lexer.
type private BodyCursor = {
    Xml: string
    /// The root element with its attributes (the namespaces of the part) and no content.
    Root: XmlNode
    /// The lexer in what the body holds: it stands on the end of the last child element read.
    Lexer: XmlDsl.XmlLexer
    /// Where the body's end tag starts.
    BodyClose: int
    /// Where what the body holds starts.
    BodyStart: int
    /// The names of the elements of the last child element read.
    Names: System.Collections.Generic.HashSet<string>
    /// How many child elements were found.
    mutable Count: int
    mutable AtEnd: bool
}

/// The cursor at the first child element of the body; None for a main part
/// that is read at once (a document type declaration, no body, an empty one,
/// a body that is not the last element of the root).
let private openBodyCursor (xml: string) : BodyCursor option =
    match XmlDsl.rootStartTag xml with
    | None -> None
    | Some (root, lexer) ->
        // (A root's `xml:space` holds for the text of the whole part: not for elements read one by one.)
        if lexer.SelfClosing || (Xml.attr "xml:space" root).IsSome then None
        else
            // The first child of the root that is a body.
            let mutable result = None
            let mutable searching = true
            while searching do
                match lexer.Next () with
                | XmlDsl.XmlToken.Text | XmlDsl.XmlToken.Comment | XmlDsl.XmlToken.Instruction | XmlDsl.XmlToken.CData -> ()
                | XmlDsl.XmlToken.StartTag when lexer.LocalNameIs "body" ->
                    searching <- false
                    if not lexer.SelfClosing then
                        let bodyName = lexer.Name
                        let bodyStart = lexer.End
                        // Its end tag: the one before the root's.
                        match XmlDsl.endTagBefore xml xml.Length with
                        | Some (rootClose, rootName) when Some rootName = Xml.name root ->
                            match XmlDsl.endTagBefore xml rootClose with
                            | Some (close, name) when name = bodyName && close >= bodyStart ->
                                result <-
                                    Some { Xml = xml; Root = root; Lexer = XmlDsl.XmlLexer(xml, bodyStart, close); BodyClose = close; BodyStart = bodyStart
                                           Names = System.Collections.Generic.HashSet<string>(); Count = 0; AtEnd = false }
                            | _ -> ()
                        | _ -> ()
                | XmlDsl.XmlToken.StartTag -> if lexer.SkipElement () < 0 then searching <- false
                | _ -> searching <- false
            result

/// The next child element of the body: (where it starts, where it ends, its
/// name, the element as the scanner reads it - None where it does not, the
/// platform's parser then reads it); None at the end of the body.
let private nextBodyChild (cursor: BodyCursor) : (int * int * string * XmlNode option) option =
    if cursor.AtEnd then None
    else
        let lexer = cursor.Lexer
        let mutable result = None
        let mutable going = true
        while going do
            match lexer.Next () with
            | XmlDsl.XmlToken.Text | XmlDsl.XmlToken.Comment | XmlDsl.XmlToken.Instruction | XmlDsl.XmlToken.CData -> ()
            | XmlDsl.XmlToken.StartTag ->
                going <- false
                let start = lexer.Start
                let name = lexer.Name
                cursor.Names.Clear ()
                let node = XmlDsl.tryReadElement lexer cursor.Names
                // (An element that does not end, or ends after the body does: the tags do not match.)
                if lexer.Token = XmlDsl.XmlToken.Malformed || lexer.Token = XmlDsl.XmlToken.EndOfText then
                    failwith "The main part of the document is not well formed"
                cursor.Count <- cursor.Count + 1
                result <- Some (start, lexer.End, name, node)
            | XmlDsl.XmlToken.EndOfText ->
                going <- false
                cursor.AtEnd <- true
            | _ -> failwith "The main part of the document is not well formed"
        result

/// The last child element of the body when that is its section properties:
/// (where it starts, the element).
let private finalSectionProperties (cursor: BodyCursor) : (int * XmlNode) option =
    // (It is a child of the body when the elements before it close up to there:
    // the reader finds out as it goes, see ProgressiveReader.)
    match XmlDsl.elementBefore cursor.Xml cursor.BodyClose cursor.BodyStart (System.Collections.Generic.HashSet<string>()) with
    | Some (XmlDsl.Element (name, _, _, span) as node) when Ns.stripNs name = "sectPr" -> Some (span.Start, node)
    | _ -> None

/// What the main part holds that is read before its body: where each section
/// properties element starts, what the sections' header / footer references say
/// each part is (relationship id -> type: headers, footers), and whether it has
/// tracked moves.
type private MainPartOutline = {
    SectionStarts: int[]
    HeaderTypes: Map<string, string>
    FooterTypes: Map<string, string>
    HasMoves: bool
}

let private outlineOfMainPart (xml: string) : MainPartOutline =
    let tags = XmlDsl.startTagsNamed [ "sectPr"; "headerReference"; "footerReference"; "moveFromRangeStart"; "moveToRangeStart" ] xml
    let local (node: XmlNode) = Xml.name node |> Option.map Ns.stripNs |> Option.defaultValue ""
    let referenceTypes (kind: string) =
        tags
        |> List.filter (fun node -> local node = kind)
        |> List.choose (fun node -> Xml.attr "r:id" node |> Option.map (fun id -> id, (Docx.attr "type" node |> Option.defaultValue "default")))
        |> Map.ofList
    { SectionStarts = tags |> List.filter (fun node -> local node = "sectPr") |> List.map (function XmlDsl.Element (_, _, _, span) -> span.Start | _ -> -1) |> Array.ofList
      HeaderTypes = referenceTypes "headerReference"
      FooterTypes = referenceTypes "footerReference"
      HasMoves = tags |> List.exists (fun node -> local node = "moveFromRangeStart" || local node = "moveToRangeStart") }

/// The local names of the elements of a node, from the names the scanner gave (`scanned`), or found in it.
let private localNamesOf (scanned: System.Collections.Generic.HashSet<string> option) (node: XmlNode) : System.Collections.Generic.HashSet<string> =
    let local = System.Collections.Generic.HashSet<string>()
    match scanned with
    | Some names -> for name in names do local.Add (Ns.stripNs name) |> ignore
    | None ->
        let rec visit (current: XmlNode) =
            match current with
            | XmlDsl.Element (name, _, kids, _) ->
                local.Add (Ns.stripNs name) |> ignore
                for kid in kids do visit kid
            | _ -> ()
        visit node
    local

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
    /// The picture of an image key (a relationship id of the body; see
    /// partImageKey) as a data URL: a picture of a group's (Shape.ImageRelationshipId).
    ImageOf: string -> string option
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
      ImageOf = fun key -> Map.tryFind key doc.Images
      Finish =
        fun () ->
            given <- true
            doc }

/// Open a package to read its main part front first. `keepContent`: Finish
/// gives the whole document as `parseFromSource` does (the blocks, the
/// revisions and every picture are kept for it); without it the blocks belong
/// to whoever read them. None: not a document.
let openProgressiveFromSource (keepContent: bool) (source: PartSource) : ProgressiveReader option =
    let source = withDataBindings source
    match readMainPart source with
    | None -> None
    | Some main ->
        let xml = main.Xml
        let outline = outlineOfMainPart xml
        let cursorAndFinal =
            // (The names of tracked moves are found by following their range markers through the part.)
            if outline.HasMoves then None
            else
                openBodyCursor xml
                |> Option.bind (fun cursor -> finalSectionProperties cursor |> Option.map (fun final -> cursor, final))
        /// The element of the main part that starts at `start` as a node (its spans are places in the part).
        let parseElement (cursor: BodyCursor) (start: int) : XmlNode =
            match XmlDsl.tryReadElementAt xml start (System.Collections.Generic.HashSet<string>()) with
            | Some node -> node
            | None ->
                // (What the scanner does not read: by the platform's parser, in the root's namespaces.)
                match XmlDsl.elementEndAt xml start with
                | fin when fin > start -> XmlDsl.parseElementOf xml cursor.Root start fin
                | _ -> failwith "The main part of the document is not well formed"
        let finalSetup =
            cursorAndFinal
            |> Option.bind (fun (cursor, (start, node)) ->
                try Some (cursor, start, parsePageSetup main.EvenAndOddHeaders main.DefaultTabStop main.CompatibilityMode main.UseFarEastLayout main.CharacterSpacingControl main.Hyphenation (Some node))
                with _ -> None)
        match finalSetup with
        | None -> parseFromSource source |> Option.map wholeReader
        | Some (cursor, finalStart, finalPageSetup) ->
            let headerRefTypes, footerRefTypes = outline.HeaderTypes, outline.FooterTypes
            let stories = readStories main headerRefTypes footerRefTypes
            let hyperlinks = hyperlinkTargets main.Rels
            let renderedBreakCounter = ref 0
            // The last block made, not given yet: what it is depends on the block after it.
            let mutable held : (DocumentContent * Shape list * int * int) option = None
            // Where the element of the last block given ends in the text.
            let mutable givenTo = cursor.BodyStart
            // The section properties that come next in the text after what is given: where their
            // start tag is (-1: none before the body's own; -2: not looked for), and the page setup they give.
            let mutable openSectionAt = -2
            let mutable openSectionSetup : PageSetup option = None
            // (The section properties ahead: from the outline of the part, passed by as the blocks are given.)
            let mutable sectionIndex = 0
            /// Where the start tag of the next section properties element at or after `from` begins; -1: none.
            let nextSectionProperties (from: int) : int =
                while sectionIndex < outline.SectionStarts.Length && outline.SectionStarts.[sectionIndex] < from do sectionIndex <- sectionIndex + 1
                if sectionIndex < outline.SectionStarts.Length then outline.SectionStarts.[sectionIndex] else -1
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
                            (try parsePageSetup main.EvenAndOddHeaders main.DefaultTabStop main.CompatibilityMode main.UseFarEastLayout main.CharacterSpacingControl main.Hyphenation (Some (parseElement cursor at))
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
                match Portable.Dict.tryFind imageKey knownImages with
                | ValueSome known -> known
                | ValueNone ->
                    let found =
                        match Portable.Dict.tryFind imageKey imageCandidates with
                        | ValueSome candidates ->
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
                        | ValueNone -> None
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
                    for id in shapeReferences shape do referenced.Add id |> ignore
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
                      | Some (start, fin, name, scanned) ->
                          let chunk = cursor.Count - 1
                          let localName = Ns.stripNs name
                          if localName = "sectPr" && firstSectionProperties < 0 then firstSectionProperties <- start
                          let givesBlocks = localName = "p" || localName = "tbl" || localName = "sdt" || localName = "customXml"
                          if givesBlocks || keepContent then
                              let node =
                                  match scanned with
                                  | Some node -> node
                                  // (What the scanner does not read: by the platform's parser, in the root's namespaces.)
                                  | None -> XmlDsl.parseElementOf xml cursor.Root start fin
                              // The local names of the elements it holds.
                              let holds = localNamesOf (if scanned.IsSome then Some cursor.Names else None) node
                              if keepContent then
                                  List.iter2 (fun revisionType (found: ResizeArray<Revision>) ->
                                      if holds.Contains revisionType then
                                          found.AddRange (Ns.descendantsNs revisionType node |> List.choose Revisions.parseRevision)) revisionTypes revisions
                              if givesBlocks then
                                  let blocks = Docx.contentOfChild renderedBreakCounter node
                                  // (A modelled shape has text or a geometry: see Docx.isModelledShape.)
                                  let blockNodes =
                                      if holds.Contains "txbxContent" || holds.Contains "prstGeom" || holds.Contains "custGeom" || holds.Contains "textpath" || holds.Contains "wgp" || holds.Contains "wpc" || holds.Contains "group" then
                                          Shapes.blockNodesOfChild node |> List.toArray
                                      else [||]
                                  blocks
                                  |> List.iteri (fun i raw ->
                                      let shapes =
                                          if i < blockNodes.Length then
                                              Shapes.shapesOfBlock (made + i) blockNodes.[i]
                                              |> List.map (fun shape ->
                                                  Shapes.mapContent (fun content -> content |> List.map Para |> completeOpaqueObjects cursor.Root |> List.choose (function Para para -> Some para | Tbl _ -> None)) shape)
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
                            parsePageSetup main.EvenAndOddHeaders main.DefaultTabStop main.CompatibilityMode main.UseFarEastLayout main.CharacterSpacingControl main.Hyphenation (Some (parseElement cursor firstSectionProperties))
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
                  Progress = fun () -> if finished then 1.0 else float cursor.Lexer.Position / float (max 1 cursor.BodyClose)
                  ImagesOf = imagesOf
                  ImageOf = imageOf
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
                    match Portable.Dict.tryFind path texts with
                    | ValueSome text -> text
                    | ValueNone ->
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
        | FieldRef field -> field.Part |> Option.orElse field.Result |> Option.defaultValue $"[Field:{field.Code}]"
        | FieldBoundary _ -> ""
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

/// Where a paragraph's page breaks (w:br w:type="page") stand: (at its start, after it). One the paragraph
/// starts with - before any of its text, pictures or fields; bookmarks and Word's marks of where it started a
/// page are none - and that has text after it is a break before the paragraph: its text starts the next page.
/// Any other is a break after its text, as the model has it - a paragraph that is a page break and nothing
/// else too: Word keeps its mark on the page the break ends (its w:lastRenderedPageBreak is in the next
/// paragraph). (A column break stays a break after: see ColumnBreak.)
let paragraphPageBreaks (para: Paragraph) : bool * bool =
    let isBreak (run: Run) = match run.Content with PageBreak -> true | _ -> false
    let isNothing (run: Run) =
        match run.Content with
        | BookmarkStart _ | BookmarkEnd _ | RenderedPageBreak | FieldBoundary _ | Text "" -> true
        | _ -> false
    let rec leading (runs: Run list) =
        match runs with
        | run :: rest when isNothing run -> leading rest
        | run :: rest when isBreak run && not para.Properties.ColumnBreak && rest |> List.exists (fun other -> not (isNothing other) && not (isBreak other)) ->
            true, List.exists isBreak rest
        | _ -> false, List.exists isBreak runs
    leading para.Runs

// Convert DocumentContent to EditorParagraph format for Word editor
/// Text no style, no document default and no run gives a size is 10 pt in Word
/// (w:sz's own default, 20 half-points), not the 11 pt of Normal in Word's template.
let internal wordFontSizeToPoints (fontSize: string option) =
    match fontSize with
    | Some value ->
        match System.Double.TryParse(value) with
        | true, halfPoints -> (halfPoints / 2.0) * 96.0 / 72.0
        | false, _ -> 10.0 * 96.0 / 72.0
    | None -> 10.0 * 96.0 / 72.0

/// w:position is in half-points and the editor (Paint) takes points. It used to
/// go through wordFontSizeToPoints, which turned it into px and, worse, turned
/// "no w:position" into the 11pt font-size default instead of 0.
let internal wordPositionToPoints (position: string option) =
    match position with
    | Some value ->
        match System.Double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture) with
        | true, halfPoints -> halfPoints / 2.0
        | false, _ -> 0.0
    | None -> 0.0

let internal wordCharacterSpacingToCssPixels (spacing: string option) =
    spacing
    |> Option.bind (fun value ->
        match System.Double.TryParse(value) with
        | true, twips -> Some (twips / 15.0)
        | false, _ -> None)

let private resolveThemeFont (theme: Theme option) (themeFont: string option) =
    match theme, themeFont |> Option.bind ValueText.StyleName.isMajorThemeFont with
    | Some theme, Some true -> theme.MajorFont
    | Some theme, Some false -> theme.MinorFont
    | _, Some true -> Some "Cambria"
    | _, Some false -> Some "Calibri"
    | _ -> None

let internal wordFontFamilyToCss (theme: Theme option) (props: RunProperties) =
    props.FontFamily
    |> Option.orElseWith (fun () -> resolveThemeFont theme props.ThemeFont)
    |> Option.defaultValue "Calibri"

let internal wordColorToCss (color: string option) =
    match color with
    | Some value when value = "" || value = "auto" -> None
    | Some value when ValueText.Color.isHash value -> Some value
    | Some value when value.Length = 8 -> Some (ValueText.Color.ofArgb value)
    | Some value -> Some ("#" + value.ToUpper())
    | None -> None

let internal wordShadingColorToEditor (color: string option) =
    match color with
    | Some value when value = "" -> None
    | Some "auto" -> Some "auto"
    | value -> wordColorToCss value

let internal directShadingBackgroundColor (shading: Shading) =
    match shading.Pattern with
    | None
    | Some "clear" -> wordColorToCss shading.Fill
    | _ -> None

let internal wordLineSpacingToMultiple fontSize (line: string option) (lineRule: string option) =
    line
    |> Option.bind (fun value ->
        match System.Double.TryParse(value) with
        | true, parsed ->
            match lineRule with
            | Some rule when rule = "exact" || rule = "atLeast" -> Some ((parsed / 15.0) / max 1.0 fontSize)
            | _ -> Some (parsed / 240.0)
        | false, _ -> None)

let internal directParagraphFormattingResetsMissingSpacing (props: ParagraphProperties) =
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

let internal tableStyleParagraphResetsMissingSpacing (props: ParagraphProperties) =
    props.LineSpacing.IsSome || props.LineRule.IsSome

let internal emptyRunProperties = {
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
    EastAsiaLanguage = None
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
    EastAsiaLanguage = props.EastAsiaLanguage |> Option.orElse baseProps.EastAsiaLanguage
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

let internal mergeParagraphProperties baseProps props = {
    Style = props.Style |> Option.orElse baseProps.Style
    ContentControlId = props.ContentControlId |> Option.orElse baseProps.ContentControlId
    ContentControlTag = props.ContentControlTag |> Option.orElse baseProps.ContentControlTag
    ContentControlAlias = props.ContentControlAlias |> Option.orElse baseProps.ContentControlAlias
    ContentControlXml = props.ContentControlXml |> Option.orElse baseProps.ContentControlXml
    ContentControlGroup = props.ContentControlGroup
    ContentControlAncestors = props.ContentControlAncestors
    EnclosingCellControls = props.EnclosingCellControls
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
    OverflowPunctuation = props.OverflowPunctuation |> Option.orElse baseProps.OverflowPunctuation
    BackgroundColor = if props.ShadingPattern.IsSome || props.ShadingForeground.IsSome || props.BackgroundColor.IsSome then props.BackgroundColor else baseProps.BackgroundColor
    ShadingPattern = if props.ShadingPattern.IsSome || props.ShadingForeground.IsSome || props.BackgroundColor.IsSome then props.ShadingPattern else baseProps.ShadingPattern
    ShadingForeground = if props.ShadingPattern.IsSome || props.ShadingForeground.IsSome || props.BackgroundColor.IsSome then props.ShadingForeground else baseProps.ShadingForeground
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
    SectionVerticalAlign = props.SectionVerticalAlign
    ColumnBreak = props.ColumnBreak
    RenderedPageBreakBefore = props.RenderedPageBreakBefore
    HeaderReferences = props.HeaderReferences
    FooterReferences = props.FooterReferences
    DropCap = props.DropCap
}

let internal paragraphPropertiesWithRunProperties runProps props =
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

let internal mergeCellProperties (baseProps: TableCellProperties option) (props: TableCellProperties option) =
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

let internal resolveTableStyle (styleMap: Map<string, StyleDefinition>) styleId : StyleDefinition option =
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
                        // (The paragraph and run properties of the style it is based on too: a
                        // table style based on Table Grid has its single spacing and no space after.)
                        ParagraphProperties = mergeConditionalParagraph (baseStyle |> Option.bind (fun style -> style.ParagraphProperties)) style.ParagraphProperties
                        RunProperties =
                            match baseStyle |> Option.bind (fun style -> style.RunProperties), style.RunProperties with
                            | Some baseValue, Some value -> Some (mergeRunProperties baseValue value)
                            | baseValue, None -> baseValue
                            | None, value -> value
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

    // (A table that names no style has the default table style - Word's "Normal Table", whose
    // cell margins of 108 twips are those of every table without margins of its own.)
    styleId
    |> Option.orElseWith (fun () ->
        styleMap |> Map.tryPick (fun id style -> if style.StyleType = "table" && style.IsDefault then Some id else None))
    |> Option.bind (resolve Set.empty)

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

let internal resolveStyle (styleMap: Map<string, StyleDefinition>) (styleId: string option) : RunProperties * ParagraphProperties option =
    match styleId with
    | None -> (emptyRunProperties, None)
    | Some id ->
        if not (obj.ReferenceEquals(resolvedStylesFor, styleMap)) then
            resolvedStyles.Clear()
            resolvedStylesFor <- box styleMap
        match Portable.Dict.tryFind id resolvedStyles with
        | ValueSome known -> known
        | ValueNone ->
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

let internal resolvedParagraphProperties (styleMap: Map<string, StyleDefinition>) (props: ParagraphProperties) =
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

let private resolvedRunPropertiesWith includeMark (styleMap: Map<string, StyleDefinition>) (paraProps: ParagraphProperties) (runProps: RunProperties) =
    let docDefaultRun, _ = resolveStyle styleMap (Some Styles.defaultStyleId)
    let styleId = paraProps.Style |> Option.orElseWith (fun () -> defaultParagraphStyleId styleMap)
    let baseRun, _ = resolveStyle styleMap styleId
    let characterRun, _ = resolveStyle styleMap runProps.Style
    let paragraphRun =
        if includeMark then paraProps.ParagraphRunProperties |> Option.defaultValue emptyRunProperties
        else emptyRunProperties
    let merged = mergeRunProperties (mergeRunProperties (mergeRunProperties (mergeRunProperties docDefaultRun baseRun) paragraphRun) characterRun) runProps
    if merged.FontFamily.IsNone && merged.ThemeFont.IsNone then
        { merged with FontFamily = docDefaultRun.FontFamily; ThemeFont = docDefaultRun.ThemeFont }
    else
        merged

/// A run's properties: the document defaults, the paragraph style's, the character
/// style's, its own. Not the paragraph mark's (w:pPr/w:rPr): those are the mark's
/// alone - a run without properties of its own under a mark in Calibri 11 pt is in
/// the style's font and size (Word 2019: Times New Roman 12 pt).
let internal resolvedRunProperties styleMap paraProps runProps = resolvedRunPropertiesWith false styleMap paraProps runProps

/// The paragraph mark's properties: a run's, with the mark's own (w:pPr/w:rPr) over them.
let internal resolvedMarkRunProperties styleMap paraProps runProps = resolvedRunPropertiesWith true styleMap paraProps runProps

/// The paragraph style of a paragraph (its own, or the document's default paragraph style).
let internal paragraphStyleIdOf (styleMap: Map<string, StyleDefinition>) (props: ParagraphProperties) =
    props.Style |> Option.orElseWith (fun () -> defaultParagraphStyleId styleMap)

/// The style of the chain from `styleId` (it, then what it is based on) whose own properties set
/// what `sets` looks for.
let internal styleSetting (styleMap: Map<string, StyleDefinition>) (sets: StyleDefinition -> bool) (styleId: string option) =
    let rec find visited id =
        if Set.contains id visited then None
        else
            match Map.tryFind id styleMap with
            | None -> None
            | Some style when sets style -> Some id
            | Some style -> style.BasedOn |> Option.bind (find (Set.add id visited))
    styleId |> Option.bind (find Set.empty)

/// A table style's paragraph properties as a cell's paragraph takes them: those its paragraph style
/// (`chain`: Normal or the style the paragraph has, with what it is based on, without the document
/// defaults) does not set - the document defaults, the table style, the paragraph style, the
/// paragraph's own (fixtures table-style-props: Normal's first-line indent, spacing, line spacing and
/// justification beat the table style's, its whole table's and its conditional formats' alike).
let internal tableStyleParagraphUnder (chain: ParagraphProperties option) (ts: ParagraphProperties) : ParagraphProperties =
    match chain with
    | None -> ts
    | Some chain ->
        let unless (set: bool) value = if set then None else value
        let firstSet = chain.FirstLineIndent.IsSome || chain.HangingIndent.IsSome
        let beforeSet = chain.SpacingBefore.IsSome || chain.SpacingBeforeAuto.IsSome
        let afterSet = chain.SpacingAfter.IsSome || chain.SpacingAfterAuto.IsSome
        let lineSet = chain.LineSpacing.IsSome || chain.LineRule.IsSome
        let shadingSet = chain.BackgroundColor.IsSome || chain.ShadingPattern.IsSome || chain.ShadingForeground.IsSome
        { ts with
            Alignment = unless chain.Alignment.IsSome ts.Alignment
            IndentLeft = unless chain.IndentLeft.IsSome ts.IndentLeft
            IndentRight = unless chain.IndentRight.IsSome ts.IndentRight
            FirstLineIndent = unless firstSet ts.FirstLineIndent
            HangingIndent = unless firstSet ts.HangingIndent
            RightTabPosition = unless chain.RightTabPosition.IsSome ts.RightTabPosition
            RightTabLeader = unless chain.RightTabPosition.IsSome ts.RightTabLeader
            TabStops = if chain.TabStops.IsEmpty then ts.TabStops else []
            Rtl = unless chain.Rtl.IsSome ts.Rtl
            SpacingBefore = unless beforeSet ts.SpacingBefore
            SpacingBeforeAuto = unless beforeSet ts.SpacingBeforeAuto
            SpacingAfter = unless afterSet ts.SpacingAfter
            SpacingAfterAuto = unless afterSet ts.SpacingAfterAuto
            LineSpacing = unless lineSet ts.LineSpacing
            LineRule = unless lineSet ts.LineRule
            OutlineLevel = unless chain.OutlineLevel.IsSome ts.OutlineLevel
            KeepNext = unless chain.KeepNext.IsSome ts.KeepNext
            KeepLines = unless chain.KeepLines.IsSome ts.KeepLines
            PageBreakBefore = unless chain.PageBreakBefore.IsSome ts.PageBreakBefore
            WidowControl = unless chain.WidowControl.IsSome ts.WidowControl
            ContextualSpacing = unless chain.ContextualSpacing.IsSome ts.ContextualSpacing
            OverflowPunctuation = unless chain.OverflowPunctuation.IsSome ts.OverflowPunctuation
            BackgroundColor = unless shadingSet ts.BackgroundColor
            ShadingPattern = unless shadingSet ts.ShadingPattern
            ShadingForeground = unless shadingSet ts.ShadingForeground
            ParagraphBorders = unless chain.ParagraphBorders.IsSome ts.ParagraphBorders
            ListInfo = unless chain.ListInfo.IsSome ts.ListInfo
            ParagraphRunProperties = None }

/// A table style's run properties as the runs of a cell's paragraph take them: those its paragraph
/// style (`chain`, as for tableStyleParagraphUnder) does not set. An on/off property the paragraph
/// style turns off does not turn the table style's off (fixtures table-style-props nb0-tb: Normal's
/// b 0 under a table style's b is bold).
let internal tableStyleRunUnder (chain: RunProperties) (ts: RunProperties) : RunProperties =
    let unless (set: bool) value = if set then None else value
    let toggle (own: bool option) value = if own = Some true then None else value
    let fontSet = chain.FontFamily.IsSome || chain.ThemeFont.IsSome
    let colorSet = chain.Color.IsSome || chain.ThemeColor.IsSome
    { Style = None
      Bold = toggle chain.Bold ts.Bold; Italic = toggle chain.Italic ts.Italic
      Underline = unless chain.Underline.IsSome ts.Underline
      UnderlineStyle = unless chain.Underline.IsSome ts.UnderlineStyle
      UnderlineColor = unless chain.Underline.IsSome ts.UnderlineColor
      AllCaps = toggle chain.AllCaps ts.AllCaps; SmallCaps = toggle chain.SmallCaps ts.SmallCaps
      FontSize = unless chain.FontSize.IsSome ts.FontSize
      CharacterSpacing = unless chain.CharacterSpacing.IsSome ts.CharacterSpacing
      FontFamily = unless fontSet ts.FontFamily; ThemeFont = unless fontSet ts.ThemeFont
      Color = unless colorSet ts.Color; ThemeColor = unless colorSet ts.ThemeColor
      ThemeTint = unless colorSet ts.ThemeTint; ThemeShade = unless colorSet ts.ThemeShade
      Highlight = unless chain.Highlight.IsSome ts.Highlight
      Strike = toggle chain.Strike ts.Strike; DoubleStrike = toggle chain.DoubleStrike ts.DoubleStrike
      Hidden = toggle chain.Hidden ts.Hidden
      Position = unless chain.Position.IsSome ts.Position
      EmphasisMark = unless chain.EmphasisMark.IsSome ts.EmphasisMark
      VertAlign = unless chain.VertAlign.IsSome ts.VertAlign
      EastAsiaLanguage = unless chain.EastAsiaLanguage.IsSome ts.EastAsiaLanguage }

/// A run's properties in a table's cell: as resolvedRunProperties, with the table style's run
/// properties (`tableRun`: what tableStyleRunUnder left of them) over the paragraph style's.
let internal resolvedCellRunProperties styleMap (tableRun: RunProperties) (paraProps: ParagraphProperties) (runProps: RunProperties) =
    let resolved = resolvedRunProperties styleMap paraProps emptyRunProperties
    let characterRun, _ = resolveStyle styleMap runProps.Style
    mergeRunProperties (mergeRunProperties (mergeRunProperties resolved tableRun) characterRun) runProps

/// The paragraph mark's properties in a table's cell (resolvedMarkRunProperties, with `tableRun`).
let internal resolvedCellMarkRunProperties styleMap (tableRun: RunProperties) (paraProps: ParagraphProperties) (runProps: RunProperties) =
    let mark = paraProps.ParagraphRunProperties |> Option.defaultValue emptyRunProperties
    resolvedCellRunProperties styleMap tableRun paraProps (mergeRunProperties mark runProps)

/// Whether Normal's 12pt and its left justification give way to a table style's size and
/// justification in its cells: in Word 2010's layout and before (compatibilityMode 14 or less, or
/// none), unless the settings' overrideTableStyleFontSizeAndJustification is on (fixtures
/// table-style-props: nsz12-tsz8-c14 / -c12 / -cnone 8pt, nsz12-tsz8-ov-c14 and nsz12-tsz8 12pt).
let tableStyleBeatsNormalSizeAndJustification (settingsXml: string option) =
    let settings = settingsXml |> Option.map XmlDsl.parse
    let setting name =
        settings |> Option.bind (fun settings ->
            Ns.descendantsNs "compatSetting" settings |> List.tryFind (fun node -> Docx.attr "name" node = Some name))
        |> Option.bind (Docx.attr "val")
    let mode = setting "compatibilityMode" |> Option.bind (fun value -> match System.Int32.TryParse value with | true, mode -> Some mode | _ -> None)
    let overridden = setting "overrideTableStyleFontSizeAndJustification" |> Option.exists (fun value -> value = "1" || value = "true" || value = "on")
    mode |> Option.forall (fun mode -> mode <= 14) && not overridden

let internal runPropertiesSetSomething (props: RunProperties) = props <> emptyRunProperties

let internal paragraphPropertiesSetSomething (props: ParagraphProperties) =
    props.ParagraphRunProperties |> Option.exists runPropertiesSetSomething
    || { props with ParagraphRunProperties = None } <> Styles.parseParagraphPropsFromStyle XmlDsl.Empty

/// The paragraph mark's resolved run properties (outside a table's cell): its paragraph style's (Normal's
/// for a bare <w:p/> as well - not the document defaults': fixtures empty-mark, and mark-font `nrm-pure`, `n12-pure`,
/// `-tbl`: Arial 8pt / 12pt lines under Normal over 11pt defaults, after text or a table), its w:rStyle's
/// under its own (`*-cstyle`: an empty paragraph under a mark of a Consolas 8pt character style 9.36pt,
/// as Consolas 8pt stated on the mark).
let internal paragraphMarkRunProperties (styleMap: Map<string, StyleDefinition>) (effectiveParaProps: ParagraphProperties) =
    match effectiveParaProps.ParagraphRunProperties with
    | Some mark when mark.Style.IsSome ->
        let styled = resolvedRunProperties styleMap effectiveParaProps { emptyRunProperties with Style = mark.Style }
        mergeRunProperties styled mark
    | _ -> resolvedMarkRunProperties styleMap effectiveParaProps emptyRunProperties

let internal styleKeepProperties (styleMap: Map<string, StyleDefinition>) (para: Paragraph) (effectiveParaProps: ParagraphProperties) =
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
        ValueText.Link.isActive lower

/// A picture source that is an address the document names rather than data
/// from the package: a linked picture, which is not loaded unless the host allows it.
let isExternalImageSource (source: string) =
    not (isNull source) && source.Contains "://" && not (ValueText.Link.isDataUrl source)

/// Source prefix of an opaque object that has no picture to show: the painter
/// draws a placeholder frame for "opaque:<kind>".
let opaqueSourcePrefix = "opaque:"

/// What an image run shows: the picture behind its relationship; for an opaque
/// object its fallback picture when the browser can draw it, else a placeholder.
let editorImageSource (images: Map<string, string>) (img: ImageData) : string =
    match img.Opaque with
    | Some opaque ->
        match Map.tryFind img.RelationshipId images with
        | Some (source: string) when ValueText.Link.isDrawnPicture source -> source
        | _ -> opaqueSourcePrefix + opaque.Kind
    | None ->
        images
        |> Map.tryFind img.RelationshipId
        |> Option.defaultValue img.RelationshipId  // Fallback to relId if not found

/// The colour a comment's mark is shown in when its run has none of its own.
/// (Of the editor's display: the writer does not put it into the file.)
let commentMarkDisplayColor = "#7c3aed"

