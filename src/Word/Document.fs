module Rbt.Word.Document

open System.Collections.Generic
open System.Text
open Rbt.Word.TextBuffer

/// Block ID for tracking blocks
type BlockId = string

/// Counter for generating unique block IDs
let mutable private blockIdCounter = 0

/// Generate next block ID
let private nextBlockId () =
#if FABLE_COMPILER
    blockIdCounter <- blockIdCounter + 1
    sprintf "block-%d" blockIdCounter
#else
    // Documents may be edited on several threads at once on a server.
    sprintf "block-%d" (System.Threading.Interlocked.Increment &blockIdCounter)
#endif

/// Text style for runs
type TextStyle = {
    Bold: bool
    Italic: bool
    Underline: bool
    FontSize: float option
    Color: string option
}

/// Default text style
let defaultStyle = {
    Bold = false
    Italic = false
    Underline = false
    FontSize = None
    Color = None
}

/// Text run - contiguous text with same formatting
type TextRun = {
    Text: string
    Style: TextStyle
}

/// List style
type ListStyle =
    | NoList
    | Bullet
    | Numbered

/// Text alignment
type Alignment =
    | Left
    | Center
    | Right
    | Justify

type SectionPageSetup = {
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
}

type ParagraphBorder = {
    Style: string option
    Color: string option
    Width: float option
    Space: float option
}

type ParagraphBorders = {
    Top: ParagraphBorder option
    Right: ParagraphBorder option
    Bottom: ParagraphBorder option
    Left: ParagraphBorder option
}

/// Block metadata
/// Newspaper columns for a section (w:cols). Per-column widths in Widths
/// (empty = equal width).
type SectionColumns = {
    Count: int
    Gap: float          // gap between columns, px (equal-width case)
    Separator: bool     // draw a rule between columns (w:sep)
    Widths: (float * float) list  // per-column (width, spaceAfter) px; empty = equal width
}

/// A w:tab stop. Position is px from the left text margin; Alignment is
/// left|center|right|decimal|bar; Leader is the optional fill (dot|hyphen|…).
type TabStop = {
    Position: float
    Alignment: string
    Leader: string option
}

/// A drop cap (w:framePr w:dropCap): the paragraph is the large first letter
/// of the paragraph after it, which goes around it.
type DropCap = {
    /// "margin": the letter stands in the margin, left of the text; else
    /// ("drop") within the text's width, the first lines of text beside it.
    InMargin: bool
    /// The lines of the next paragraph the letter is as high as (w:lines).
    Lines: int
    /// The room between the letter and the text beside it (w:hSpace), in pixels.
    Distance: float
}

/// One level of a list as the file's numbering part defines it (a w:lvl of
/// the list's abstract numbering, with the list's overrides).
type ListLevelDefinition = {
    /// 1-9, as BlockMetadata.ListLevel.
    Level: int
    /// w:numFmt: "decimal", "lowerLetter", "upperRoman", "bullet", "none", ...
    NumberFormat: string
    /// w:lvlText: the number's text, with %1 ... %9 for the numbers of the levels.
    LevelText: string option
    Start: int option
    /// The level's indents (layout pixels; a negative first line is hanging).
    IndentLeft: float option
    FirstLineIndent: float option
    /// w:suff: what follows the number - "tab" (none stated), "space", "nothing".
    Suffix: string option
    /// w:lvlJc: how the number stands at its position - "left" (none stated), "center", "right".
    Justification: string option
    /// w:isLgl: the numbers of all the levels in this level's text are decimal.
    Legal: bool
    /// w:lvlRestart: the level after which this one starts again (0: never); none: after any level above it.
    Restart: int option
    /// The look of the number (the level's run properties), where it states one.
    MarkFont: string option
    MarkBold: bool option
    MarkItalic: bool option
    MarkColor: string option
    MarkSize: float option
}

/// Where a paragraph, row or cell of a document that was read stands in its part: the start of its w:p,
/// w:tr or w:tc in the part's text. The save writes it from that element - what the model does not hold of it is
/// kept - and one that is new from the model alone. ReadAt: the element that was read. LikeAt: one made
/// like it (a row or a column put in next to it, the second half of a paragraph split), written as that element with the model's content.
type SourceOrigin =
    | ReadAt of int
    | LikeAt of int

type ContentControlLayer = TextBuffer.ContentControlLayer

type BlockMetadata = {
    Style: string option
    ContentControlId: string option
    ContentControlTag: string option
    ContentControlAlias: string option
    /// Original properties and end properties of the innermost block control,
    /// including unknown XML and data bindings, retained across edits.
    ContentControlXml: string option
    /// Identity of that control in the source; distinct adjacent controls can
    /// have identical properties. Ancestors are ordered outermost first.
    ContentControlGroup: string option
    /// Source cell-internal controls, for editing constraints rather than wrapper emission.
    EnclosingCellControls: ContentControlLayer list
    ContentControlAncestors: ContentControlLayer list
    /// The w:p the paragraph is of (see SourceOrigin).
    Origin: SourceOrigin option
    ListStyle: ListStyle
    ListLevel: int  // 0 = not in list, 1-9 = nesting level
    ListId: string option
    /// Save-time start overrides (level 1–9, value) for a physical instance.
    ListStartOverrides: (int * int) list
    /// Source numbering instance to clone when this newly authored instance is saved.
    ListSourceId: string option
    /// Start override on the authored instance's current level.
    ListRestart: bool
    /// The marker format/text of this cloned level was explicitly changed.
    ListFormatEdited: bool
    ListNumberFormat: string option
    ListLevelText: string option
    ListStart: int option
    Alignment: Alignment
    /// True when Alignment was set directly on the paragraph. Left is also the
    /// "unset" value, so explicit left under a centred style needs the flag.
    ExplicitAlignment: bool
    IndentLeft: float option
    IndentRight: float option
    FirstLineIndent: float option
    RightTabPosition: float option
    RightTabLeader: string option
    TabStops: TabStop list
    Rtl: bool   // w:bidi — right-to-left paragraph
    SpacingBefore: float option
    SpacingAfter: float option
    ParagraphMarkFontSize: float option
    ContextualSpacing: bool option
    LineSpacing: float option
    LineSpacingRule: string option
    WidowControl: bool option
    /// w:keepNext / w:keepLines set directly on the paragraph (None = as the style says).
    KeepNext: bool option
    KeepLines: bool option
    BackgroundColor: string option
    ParagraphShadingPattern: string option
    ParagraphShadingForeground: string option
    ExplicitParagraphShadingPattern: string option
    ExplicitParagraphShadingForeground: string option
    ParagraphBorders: ParagraphBorders option
    ParagraphBorderBefore: float option
    ParagraphBorderAfter: float option
    PageBreakBefore: bool
    PageBreakAfter: bool
    /// The break after the paragraph (PageBreakAfter) is a column break: the
    /// text goes on in the next column of a section with columns, and on the
    /// next page where there is one column.
    ColumnBreakAfter: bool
    SectionBreak: string option
    SectionColumns: SectionColumns option
    SectionPageNumberStart: int option
    /// w:pgNumType w:fmt of the section the paragraph ends ("decimal", "lowerRoman", ...).
    SectionPageNumberFormat: string option
    /// w:outlineLvl the paragraph states itself (0-9; a title's 0): kept and saved.
    OutlineLevel: int option
    /// w:vAlign of the section the paragraph ends: how the text of a page of
    /// it that is not full stands between the margins ("top", "center",
    /// "both" - spread from top to bottom -, "bottom").
    SectionVerticalAlign: string option
    RenderedPageBreakBefore: int option
    SectionPageSetup: SectionPageSetup option
    HeaderReferences: (string * string) list
    FooterReferences: (string * string) list
    /// Bit set of ParagraphExplicit flags: the properties stated on the paragraph
    /// itself. The values above are the effective ones (the reader resolves
    /// styles and numbering into them); an unflagged one is inherited or not
    /// set anywhere, and a save leaves it out of the paragraph.
    Explicit: int
    /// The paragraph mark has run properties of its own, and the size they state.
    HasMarkProperties: bool
    ExplicitMarkFontSize: float option
    /// The indents (left, first line; a negative first line is hanging) of the
    /// paragraph's list level, when its list came from a file: where an indent
    /// the paragraph does not state comes from.
    ListLevelIndents: (float option * float option) option
    /// All the levels of the paragraph's list, when the list came from a file
    /// (the same list object for every paragraph of the list): what the
    /// paragraph has at another level (see withListLevel), and how the numbers
    /// of the other levels are written in its own.
    ListLevels: ListLevelDefinition list
    /// The paragraph mark itself is a tracked change (w:ins / w:del in the
    /// mark's run properties): inserted = the paragraph was split here, deleted
    /// = it is to be joined with the next one.
    MarkRevision: ParagraphMarkRevision option
    /// The paragraph is the drop cap of the paragraph after it.
    DropCap: DropCap option
}

/// A tracked paragraph mark: RevisionType is "insert" or "delete".
and ParagraphMarkRevision = {
    RevisionType: string
    Mark: RevisionMark
}

/// Flags for BlockMetadata.Explicit.
module ParagraphExplicit =
    [<Literal>]
    let SpacingBefore = 1
    [<Literal>]
    let SpacingAfter = 2
    [<Literal>]
    let LineSpacing = 4
    [<Literal>]
    let IndentLeft = 8
    [<Literal>]
    let IndentRight = 16
    [<Literal>]
    let FirstLineIndent = 32
    [<Literal>]
    let ContextualSpacing = 64
    [<Literal>]
    let BackgroundColor = 128
    [<Literal>]
    let Tabs = 256
    /// Not something the paragraph states, the opposite: its PageBreakBefore
    /// is its style's, or the reader's for a heading that started a page when
    /// the document was last laid out. It is not saved as a break of the paragraph.
    [<Literal>]
    let DerivedPageBreakBefore = 512
    /// The paragraph's PageBreakBefore is a page break its text starts with
    /// (w:br w:type="page" before its first run of text): saved as that break,
    /// not as w:pageBreakBefore. Split, it stays with the first part.
    [<Literal>]
    let PageBreakRunBefore = 1024

    let has (flag: int) (metadata: BlockMetadata) = metadata.Explicit &&& flag <> 0

/// Default block metadata
/// The origin of a paragraph, row or cell made like one with `origin`.
let likeOrigin (origin: SourceOrigin option) : SourceOrigin option =
    origin |> Option.map (function ReadAt at | LikeAt at -> LikeAt at)

let defaultMetadata = {
    Style = None
    ContentControlId = None
    ContentControlTag = None
    ContentControlAlias = None
    ContentControlXml = None
    ContentControlGroup = None
    ContentControlAncestors = []
    EnclosingCellControls = []
    Origin = None
    ListStyle = NoList
    ListLevel = 0
    ListId = None
    ListStartOverrides = []
    ListSourceId = None
    ListRestart = false
    ListFormatEdited = false
    ListNumberFormat = None
    ListLevelText = None
    ListStart = None
    Alignment = Left
    ExplicitAlignment = false
    IndentLeft = None
    IndentRight = None
    FirstLineIndent = None
    RightTabPosition = None
    RightTabLeader = None
    TabStops = []
    Rtl = false
    SpacingBefore = None
    SpacingAfter = None
    ParagraphMarkFontSize = None
    ContextualSpacing = None
    LineSpacing = None
    LineSpacingRule = None
    WidowControl = None
    KeepNext = None
    KeepLines = None
    BackgroundColor = None
    ParagraphShadingPattern = None
    ParagraphShadingForeground = None
    ExplicitParagraphShadingPattern = None
    ExplicitParagraphShadingForeground = None
    ParagraphBorders = None
    ParagraphBorderBefore = None
    ParagraphBorderAfter = None
    PageBreakBefore = false
    PageBreakAfter = false
    ColumnBreakAfter = false
    SectionVerticalAlign = None
    OutlineLevel = None
    SectionBreak = None
    SectionColumns = None
    SectionPageNumberStart = None
    SectionPageNumberFormat = None
    RenderedPageBreakBefore = None
    SectionPageSetup = None
    HeaderReferences = []
    FooterReferences = []
    Explicit = 0
    HasMarkProperties = false
    ExplicitMarkFontSize = None
    ListLevelIndents = None
    ListLevels = []
    MarkRevision = None
    DropCap = None
}

/// Advance the list counters exactly as layout does, without formatting a marker.
let advanceListCounters (metadata: BlockMetadata) (counters: Map<string * int, int>) =
    if metadata.ListStyle <> Numbered then counters
    else
        let level = max 1 metadata.ListLevel
        let id = metadata.ListId |> Option.defaultValue ""
        let key = id, level
        let value = counters |> Map.tryFind key |> Option.map ((+) 1) |> Option.defaultValue (metadata.ListStart |> Option.defaultValue 1)
        let mutable updated = counters |> Map.add key value
        for child in level + 1 .. 9 do
            let restart = metadata.ListLevels |> List.tryFind (fun definition -> definition.Level = child) |> Option.bind (fun definition -> definition.Restart)
            let resets = match restart with Some 0 -> false | Some after -> level <= after | None -> true
            if resets then updated <- updated |> Map.remove (id, child)
        updated

/// One segment of a vector path. Coordinates are in the path's own coordinate
/// space (see ShapeGeometry); the painter scales them to the shape's box.
type PathSeg =
    | PMove of x: float * y: float
    | PLine of x: float * y: float
    | PCubic of c1x: float * c1y: float * c2x: float * c2y: float * x: float * y: float
    | PQuad of cx: float * cy: float * x: float * y: float
    | PClose

/// A gradient colour stop (offset 0..1).
type GradientStop = { Offset: float; Color: string }

/// Linear (angle in degrees) or radial gradient fill.
type GradientSpec = {
    Stops: GradientStop list
    Angle: float
    Radial: bool
}

/// Vector geometry of a drawing shape / text box.
type ShapeGeometry =
    /// Plain rectangle — the legacy text-box shape.
    | GeoRect
    /// An OOXML preset (a:prstGeom@prst, e.g. "ellipse", "roundRect"). Resolved
    /// to a path at paint time by Shapes.presetPath against the shape's w×h.
    | GeoPreset of preset: string
    /// A freeform a:custGeom path in its own coordinate space (coordW × coordH).
    | GeoCustom of segs: PathSeg list * coordW: float * coordH: float

/// Block types
type Block =
    | Paragraph of ParagraphBlock
    | Heading of level: int * ParagraphBlock
    | Table of TableBlock
    | TextBox of TextBoxBlock

and ParagraphBlock = {
    Id: BlockId
    Buffer: TextBuffer
    Metadata: BlockMetadata
}

and TableBlock = {
    Id: BlockId
    Rows: TableRow list
    PreferredWidth: TableWidth option
    Alignment: string option
    /// Table indent in canvas pixels.
    Indent: float option
    LayoutMode: TableLayoutMode option
    GridColumns: TableWidth list
    Borders: TableBorders option
    BackgroundColor: string option
    Shading: Shading option
    CellSpacing: float option
    PaddingTop: float option
    PaddingRight: float option
    PaddingBottom: float option
    PaddingLeft: float option
    Floating: bool
    /// The w:tblpPr of a floating table of the file, kept for the save.
    FloatingXml: string option
    /// The table style (w:tblStyle) and the conditional formats of it that are
    /// switched on (w:tblLook). The style's formatting is already resolved into
    /// the cells; the reference is kept for the file.
    StyleId: string option
    ContentControlId: string option
    ContentControlTag: string option
    ContentControlAlias: string option
    /// The bound control's original properties and end properties.
    ContentControlXml: string option
    ContentControlGroup: string option
    ContentControlAncestors: ContentControlLayer list
    Look: TableLook option
}

and TableLook = {
    FirstRow: bool
    LastRow: bool
    FirstColumn: bool
    LastColumn: bool
    HorizontalBanding: bool
    VerticalBanding: bool
}

and TextBoxBlock = {
    Id: BlockId
    /// The source drawing's identity and kind, separate from the editor block ID.
    SourceShapeId: string option
    SourceShapeKind: string option
    Content: ParagraphBlock list
    X: float option
    Y: float option
    PositionHRelativeFrom: string option
    PositionVRelativeFrom: string option
    Width: float
    Height: float
    FillColor: string option
    BorderColor: string option
    BorderWidth: float option
    RelativeHeight: string option
    Geometry: ShapeGeometry
    Gradient: GradientSpec option
    Rotation: float
    /// The shape is mirrored across / upside down (a:xfrm flipH / flipV) before it
    /// is turned; its text is not.
    FlipH: bool
    FlipV: bool
    /// In the text line (wp:inline) rather than floating (wp:anchor).
    Inline: bool
    /// none | square | tight | through | topAndBottom.
    Wrap: string option
    /// wp:align of the position (left, center, right, top, ...), used instead of X / Y.
    PositionHAlign: string option
    PositionVAlign: string option
    BehindDoc: bool option
    AllowOverlap: bool option
    LayoutInCell: bool option
    DistanceTop: float option
    DistanceBottom: float option
    DistanceLeft: float option
    DistanceRight: float option
    /// Preset line dash (dash, dot, dashDot, sysDash, ...); None = solid.
    BorderDash: string option
    Name: string option
    /// The body block (paragraph or table) the shape is anchored in.
    AnchorBlockId: BlockId option
    /// The room between the shape's edge and its text (wps:bodyPr lIns / tIns /
    /// rIns / bIns), in pixels. None: 0.1 inch left and right, 0.05 inch top and bottom.
    InsetLeft: float option
    InsetTop: float option
    InsetRight: float option
    InsetBottom: float option
    /// Where the text stands in the shape's height (wps:bodyPr anchor): t | ctr | b.
    VerticalAnchor: string option
    /// The shape is as high as its text (a:spAutoFit).
    AutoFit: bool
    /// The text is not broken into lines at the shape's width (wps:bodyPr wrap="none").
    NoTextWrap: bool
    /// The text is drawn as large as the shape (WordArt, a watermark's
    /// v:textpath): one line, scaled to the shape's width and height.
    FitText: bool
    /// The opacity of the fill and of fitted text (1: opaque).
    Opacity: float
    /// The shapes of a group or a drawing canvas (wpg:wgp, wpc:wpc, v:group), in
    /// the order they are drawn: each with its X / Y / Width / Height in pixels
    /// from the group's top-left corner in its ChildSpace (a group in a group
    /// has shapes of its own). Empty for any other shape. A group has no text
    /// of its own; the text of its shapes is drawn, not edited.
    Children: TextBoxBlock list
    /// The size the places of a group's shapes are given in (the group's own as
    /// it was read): a group drawn at another size scales them with it. None:
    /// not scaled (a drawing canvas).
    ChildSpace: (float * float) option
    /// A picture of a group: the picture drawn in the shape's box.
    ImageSource: string option
    /// The element the shape was read from (mc:AlternateContent, w:drawing or
    /// w:pict, as it stood in its run): written back while the shape is what was read.
    SourceXml: string option
}

and TableRow = {
    Cells: TableCell list
    IsHeader: bool
    Height: float option
    HeightRule: string option
    CanSplit: bool option
    CellSpacing: float option
    PaddingTop: float option
    PaddingRight: float option
    PaddingBottom: float option
    PaddingLeft: float option
    GridBefore: int
    GridAfter: int
    WidthBefore: float option
    WidthAfter: float option
    BordersOverride: TableBorders option
    Origin: SourceOrigin option
}

and TableCell = {
    Content: ParagraphBlock list  // Cells can contain multiple paragraphs
    NestedTables: TableBlock list
    // For each entry in NestedTables, the number of cell paragraphs that
    // precede it in document order. Lets layout/paint interleave paragraphs
    // and nested tables instead of hoisting all tables to the cell top.
    NestedTableAnchors: int list
    /// Cell-level source controls, separate from controls emitted inside paragraphs.
    EnclosingControls: ContentControlLayer list
    ColSpan: int
    BackgroundColor: string option
    Shading: Shading option
    Width: float option
    PreferredWidth: TableWidth option
    BorderColor: string option
    BorderWidth: float option
    Borders: TableBorders option
    VerticalAlignment: string option
    VerticalMerge: TableVerticalMerge option
    PaddingTop: float option
    PaddingRight: float option
    PaddingBottom: float option
    PaddingLeft: float option
    Origin: SourceOrigin option
}

and Shading = {
    Pattern: string option
    ForegroundColor: string option
    BackgroundColor: string option
}

and TableWidthKind =
    | AutoWidth
    | DxaWidth
    | PercentWidth
    | PointWidth

and TableWidth = {
    Kind: TableWidthKind
    Value: float option
}

and TableLayoutMode =
    | FixedTableLayout
    | AutoFitTableLayout

and TableBorder = {
    Style: string option
    Color: string option
    Width: float option
}

and TableBorders = {
    Top: TableBorder option
    Right: TableBorder option
    Bottom: TableBorder option
    Left: TableBorder option
    InsideH: TableBorder option
    InsideV: TableBorder option
}

and TableVerticalMerge =
    | RestartMerge
    | ContinueMerge

/// A text box as the editor makes it: a white rectangle with a grey outline.
let defaultTextBox : TextBoxBlock = {
    Id = ""
    SourceShapeId = None
    SourceShapeKind = None
    Content = []
    X = None
    Y = None
    PositionHRelativeFrom = None
    PositionVRelativeFrom = None
    Width = 220.0
    Height = 72.0
    FillColor = Some "#FFFFFF"
    BorderColor = Some "#808080"
    BorderWidth = Some 1.0
    RelativeHeight = None
    Geometry = GeoRect
    Gradient = None
    Rotation = 0.0
    FlipH = false
    FlipV = false
    Inline = false
    Wrap = None
    PositionHAlign = None
    PositionVAlign = None
    BehindDoc = None
    AllowOverlap = None
    LayoutInCell = None
    DistanceTop = None
    DistanceBottom = None
    DistanceLeft = None
    DistanceRight = None
    BorderDash = None
    Name = None
    AnchorBlockId = None
    InsetLeft = None
    InsetTop = None
    InsetRight = None
    InsetBottom = None
    VerticalAnchor = None
    AutoFit = false
    NoTextWrap = false
    FitText = false
    Opacity = 1.0
    Children = []
    ChildSpace = None
    ImageSource = None
    SourceXml = None
}

/// Cursor position
type Position = {
    BlockId: BlockId
    Offset: int  // Character offset in block
}

/// Selection (anchor = start, focus = end/cursor)
type Selection = {
    Anchor: Position
    Focus: Position
}
with
    /// Check if selection is collapsed (cursor only)
    member this.IsCollapsed =
        this.Anchor.BlockId = this.Focus.BlockId &&
        this.Anchor.Offset = this.Focus.Offset

    /// Get normalized range (start, end) accounting for direction - single block only
    member this.Range =
        if this.Anchor.BlockId = this.Focus.BlockId then
            if this.Anchor.Offset <= this.Focus.Offset then
                (this.Anchor, this.Focus)
            else
                (this.Focus, this.Anchor)
        else
            // For now, assume single block
            (this.Anchor, this.Focus)

type BookmarkInfo = {
    Id: string
    Name: string
    BlockId: BlockId
    StartOffset: int
    EndOffset: int
}

type ContentControlInfo = {
    Id: string
    Tag: string
    Alias: string
    BlockId: BlockId
}

type CommentInfo = {
    Id: string
    Text: string
    Author: string option
    Date: string option
    BlockId: BlockId
    StartOffset: int
    EndOffset: int
    Excerpt: string
}

/// Document
type HeaderFooterPart = {
    Kind: string
    RelationshipId: string option
    Blocks: Block list
}


type PageSetup = {
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
    /// The distance of the default tab stops, in pixels (w:defaultTabStop of
    /// the settings part; Word's 720 twips, half an inch, when it says none).
    DefaultTabStop: float
    Columns: SectionColumns option
    SectionStart: string option
    /// Header / footer references (kind, relationship id) of the final section,
    /// the one the body-level w:sectPr describes. The earlier sections carry
    /// theirs on the paragraph that ends them (BlockMetadata.HeaderReferences).
    HeaderReferences: (string * string) list
    FooterReferences: (string * string) list
    /// w:pgNumType w:start of the final section.
    PageNumberStart: int option
    /// w:pgNumType w:fmt of the final section.
    PageNumberFormat: string option
    /// w:vAlign of the final section (see BlockMetadata.SectionVerticalAlign).
    VerticalAlign: string option
}

type NamedStyle = {
    Id: string
    Name: string option
    StyleType: string
    IsDefault: bool
    BasedOn: string option
    Next: string option
    Linked: string option
    IsQuickStyle: bool
    Aliases: string option
    UiPriority: int option
    SemiHidden: bool
    UnhideWhenUsed: bool
    Text: TextBuffer.TextStyle option
    TextBold: bool option
    TextItalic: bool option
    TextUnderline: bool option
    TextAllCaps: bool option
    TextSmallCaps: bool option
    TextStrike: bool option
    SpacingBefore: float option
    SpacingAfter: float option
    ContextualSpacing: bool option
    LineSpacing: float option
    LineSpacingRule: string option
    Alignment: Alignment option
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

/// The body of a footnote or endnote as it is in the file: blocks with their
/// formatting, hyperlinks and tables. The reference run in the text carries
/// the note's plain text (TextStyle.NoteText) for display; `Text` is that
/// plain text as imported, so an edit of the reference's NoteText is noticed
/// (the note is then saved from the edited text instead of these blocks).
type NoteBody = {
    /// "footnote" or "endnote".
    NoteType: string
    Id: string
    /// w:type of a special note (separator, continuationSeparator, ...); None for an ordinary one.
    Kind: string option
    /// The body starts with the automatic reference mark.
    HasMark: bool
    MarkStyle: string option
    Text: string
    Blocks: Block list
}

/// The body of a comment as it is in the file (see NoteBody).
type CommentBody = {
    Id: string
    Author: string option
    Date: string option
    Initials: string option
    Text: string
    Blocks: Block list
}

/// What belongs to the document's package beyond the body and the
/// headers / footers. Carried through editing and written back on save.
type PackageData = {
    Notes: NoteBody list
    Comments: CommentBody list
    /// Theme used to resolve imported styles and retained for package export.
    Theme: DocxReader.Theme option
    /// The parts of the opened package the editor does not model (theme,
    /// settings, document properties, custom XML, chart / diagram / embedded
    /// object parts, the original styles part, ...), exactly as the reader
    /// delivered them. None for a document made from scratch.
    Preserved: WordEditorPackage.PreservedPackage option
    /// The style catalog as imported. A style that still equals its entry here
    /// was not changed in the editor, and its original definition is kept.
    ImportedStyles: NamedStyle list
    /// Mutable only while a progressive import extends its original baseline.
    Imported: ImportedDocument option
}

and Document = {
    Blocks: Block list
    Selection: Selection option
    PageSetup: PageSetup
    Styles: NamedStyle list
    Headers: ParagraphBlock list
    Footers: ParagraphBlock list
    HeaderParts: HeaderFooterPart list
    FooterParts: HeaderFooterPart list
    Package: PackageData
}

and ImportedDocument = {
    mutable Document: Document
}

/// Whether two snapshots contain the same document, independently of the caret.
let sameContent (left: Document) (right: Document) =
    obj.ReferenceEquals(left, right)
    || ((obj.ReferenceEquals(left.Blocks, right.Blocks) || left.Blocks = right.Blocks)
        && left.PageSetup = right.PageSetup
        && (obj.ReferenceEquals(left.Styles, right.Styles) || left.Styles = right.Styles)
        && (obj.ReferenceEquals(left.Headers, right.Headers) || left.Headers = right.Headers)
        && (obj.ReferenceEquals(left.Footers, right.Footers) || left.Footers = right.Footers)
        && (obj.ReferenceEquals(left.HeaderParts, right.HeaderParts) || left.HeaderParts = right.HeaderParts)
        && (obj.ReferenceEquals(left.FooterParts, right.FooterParts) || left.FooterParts = right.FooterParts)
        && (obj.ReferenceEquals(left.Package.Notes, right.Package.Notes) || left.Package.Notes = right.Package.Notes)
        && (obj.ReferenceEquals(left.Package.Comments, right.Package.Comments) || left.Package.Comments = right.Package.Comments)
        && left.Package.Theme = right.Package.Theme
        && obj.ReferenceEquals(left.Package.Preserved, right.Package.Preserved))

let emptyPackageData = {
    Notes = []
    Comments = []
    Theme = None
    Preserved = None
    ImportedStyles = []
    Imported = None
}

let defaultPageSetup = {
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
    DefaultTabStop = 48.0
    Columns = None
    SectionStart = None
    HeaderReferences = []
    FooterReferences = []
    PageNumberStart = None
    PageNumberFormat = None
    VerticalAlign = None
}

/// A paragraph that ends a section: its w:sectPr describes the section that
/// ends with it (break type, page geometry, columns, header / footer
/// references, page number start).
let endsSection (metadata: BlockMetadata) =
    metadata.SectionBreak.IsSome
    || metadata.SectionPageSetup.IsSome
    || not metadata.HeaderReferences.IsEmpty
    || not metadata.FooterReferences.IsEmpty

/// A section's references by kind: its own, and for a kind it does not name
/// the one it inherits from the section before it.
let inheritSectionReferences (previous: (string * string) list) (own: (string * string) list) : (string * string) list =
    own @ (previous |> List.filter (fun (kind, _) -> not (own |> List.exists (fun (ownKind, _) -> ownKind = kind))))


/// The paragraphs of a cell in document order: its own, and those of the
/// tables in it at the places NestedTableAnchors gives (as the layout puts
/// them: a table with no anchor comes after the cell's paragraphs).
let rec cellParagraphsInOrder (cell: TableCell) : ParagraphBlock list =
    if cell.NestedTables.IsEmpty then cell.Content
    else
        let content = cell.Content |> List.toArray
        let anchors = cell.NestedTableAnchors |> List.toArray
        let anchored =
            cell.NestedTables
            |> List.mapi (fun i table -> table, (if i < anchors.Length then anchors.[i] else content.Length) |> max 0 |> min content.Length)
            |> List.sortBy snd
            |> List.toArray
        let result = ResizeArray<ParagraphBlock>()
        let mutable tableIndex = 0
        for paraIndex in 0 .. content.Length do
            while tableIndex < anchored.Length && snd anchored.[tableIndex] = paraIndex do
                result.AddRange(tableParagraphsInOrder (fst anchored.[tableIndex]))
                tableIndex <- tableIndex + 1
            if paraIndex < content.Length then result.Add content.[paraIndex]
        List.ofSeq result

/// The paragraphs of a table, row by row and cell by cell, those of the
/// tables in its cells included.
and tableParagraphsInOrder (table: TableBlock) : ParagraphBlock list =
    table.Rows |> List.collect (fun row -> row.Cells |> List.collect cellParagraphsInOrder)

/// Whether a table, or a table in one of its cells, has the id or holds the
/// paragraph with it.
let rec tableHolds (blockId: string) (table: TableBlock) : bool =
    table.Id = blockId
    || (let mutable found = false
        for row in table.Rows do
            if not found then
                for cell in row.Cells do
                    if not found then
                        for para in cell.Content do
                            if para.Id = blockId then found <- true
                        if not found then
                            for nested in cell.NestedTables do
                                if not found && tableHolds blockId nested then found <- true
        found)

/// The table with an id among tables and the tables in their cells.
let rec tryFindTableIn (tableId: string) (tables: TableBlock list) : TableBlock option =
    tables
    |> List.tryPick (fun table ->
        if table.Id = tableId then Some table
        else table.Rows |> List.tryPick (fun row -> row.Cells |> List.tryPick (fun cell -> if cell.NestedTables.IsEmpty then None else tryFindTableIn tableId cell.NestedTables)))

/// Every paragraph of a block list in document order, including the cells of
/// nested tables and text box content.
let rec paragraphsOfBlocks (blocks: Block list) : ParagraphBlock list =
    let rec tableParagraphs (table: TableBlock) =
        table.Rows
        |> List.collect (fun row ->
            row.Cells
            |> List.collect (fun cell -> cell.Content @ (cell.NestedTables |> List.collect tableParagraphs)))
    blocks
    |> List.collect (function
        | Paragraph para
        | Heading (_, para) -> [para]
        | Table table -> tableParagraphs table
        | TextBox textBox -> textBox.Content)

/// Apply `mapper` to every paragraph of a block list (see paragraphsOfBlocks).
let rec mapParagraphsOfBlocks (mapper: ParagraphBlock -> ParagraphBlock) (blocks: Block list) : Block list =
    let rec mapTable (table: TableBlock) =
        { table with
            Rows =
                table.Rows
                |> List.map (fun row ->
                    { row with
                        Cells =
                            row.Cells
                            |> List.map (fun cell ->
                                { cell with
                                    Content = cell.Content |> List.map mapper
                                    NestedTables = cell.NestedTables |> List.map mapTable }) }) }
    blocks
    |> List.map (function
        | Paragraph para -> Paragraph (mapper para)
        | Heading (level, para) -> Heading (level, mapper para)
        | Table table -> Table (mapTable table)
        | TextBox textBox -> TextBox { textBox with Content = textBox.Content |> List.map mapper })

/// A copy of a block list in which every block has an id of its own:
/// paragraphs (those in table cells, nested tables and text boxes too),
/// tables with the tables in their cells, and text boxes. `newId kind` gives
/// an id no block has ("paragraph" | "table" | "box" | "control").
let copyBlocksWithNewIds (newId: string -> BlockId) (blocks: Block list) : Block list =
    let groups = Dictionary<string, string>()
    let remap group =
        group |> Option.map (fun group ->
            match groups.TryGetValue group with
            | true, found -> found
            | _ ->
                let fresh = newId "control"
                groups.[group] <- fresh
                fresh)
    let ancestors layers =
        layers |> List.map (fun layer ->
            { layer with Group = remap (Some layer.Group) |> Option.get })
    let rec copyTable (table: TableBlock) : TableBlock =
        { table with
            Id = newId "table"
            ContentControlGroup = remap table.ContentControlGroup
            ContentControlAncestors = ancestors table.ContentControlAncestors
            Rows =
                table.Rows
                |> List.map (fun row ->
                    { row with Cells = row.Cells |> List.map (fun cell -> { cell with NestedTables = cell.NestedTables |> List.map copyTable }) }) }
    blocks
    |> mapParagraphsOfBlocks (fun para ->
        { para with
            Id = newId "paragraph"
            Metadata = {
                para.Metadata with
                    ContentControlGroup = remap para.Metadata.ContentControlGroup
                    ContentControlAncestors = ancestors para.Metadata.ContentControlAncestors
            } })
    |> List.map (function
        | Table table -> Table (copyTable table)
        | TextBox textBox -> TextBox { textBox with Id = newId "box" }
        | block -> block)

/// Every paragraph of the document: body, then header and footer parts.
let allParagraphs (doc: Document) : ParagraphBlock list =
    paragraphsOfBlocks doc.Blocks
    @ (doc.HeaderParts |> List.collect (fun part -> paragraphsOfBlocks part.Blocks))
    @ (doc.FooterParts |> List.collect (fun part -> paragraphsOfBlocks part.Blocks))

/// Where page numbering restarts: (id of the block, id of its first paragraph,
/// first page number). A section's w:pgNumType w:start applies from the first
/// page of that section, so the restart point is the first block after the
/// previous section's end, not the paragraph that carries the section
/// properties (the section's last one). A paragraph that carries a start
/// without ending a section restarts at itself.
let pageNumberRestarts (doc: Document) : (BlockId * BlockId option * int) list =
    let result = ResizeArray<BlockId * BlockId option * int>()
    let idOf block =
        match block with
        | Paragraph para
        | Heading (_, para) -> para.Id
        | Table table -> table.Id
        | TextBox textBox -> textBox.Id
    let firstParagraphId block = paragraphsOfBlocks [ block ] |> List.tryHead |> Option.map (fun para -> para.Id)
    let mutable sectionFirst : Block option = None
    for block in doc.Blocks do
        let isFlow = match block with TextBox _ -> false | _ -> true
        if isFlow && sectionFirst.IsNone then sectionFirst <- Some block
        match block with
        | Paragraph para
        | Heading (_, para) ->
            match para.Metadata.SectionPageNumberStart with
            | Some start when endsSection para.Metadata ->
                let first = sectionFirst |> Option.defaultValue block
                result.Add(idOf first, firstParagraphId first, start)
            | Some start -> result.Add(para.Id, Some para.Id, start)
            | None -> ()
            if endsSection para.Metadata then sectionFirst <- None
        | _ -> ()
    match doc.PageSetup.PageNumberStart, sectionFirst with
    | Some start, Some first -> result.Add(idOf first, firstParagraphId first, start)
    | _ -> ()
    List.ofSeq result

let setPageSetup (setup: PageSetup) (doc: Document) =
    if doc.PageSetup = setup then doc else { doc with PageSetup = setup }

type FindMatch = {
    BlockId: BlockId
    Start: int
    End: int
    Text: string
}

type RevisionInfo = {
    BlockId: BlockId
    Start: int
    End: int
    RevisionType: string
    Author: string option
    Date: string option
    Text: string
}

type PageBlockAnchor = {
    PageBlockId: BlockId
    PageNumber: int
}

type PageFieldAnchor = {
    PageFieldBlockId: BlockId
    PageFieldStartOffset: int
    PageFieldCode: string
    PageNumber: int
}

type PageBookmarkAnchor = {
    PageBookmarkName: string
    PageBookmarkBlockId: BlockId
    PageBookmarkStartOffset: int
    PageNumber: int
}

type PageMap = {
    PageCount: int
    BlockAnchors: PageBlockAnchor list
    Fields: PageFieldAnchor list
    Bookmarks: PageBookmarkAnchor list
}

/// Create empty document
let empty = {
    Blocks = [
        Paragraph {
            Id = "block-0"
            Buffer = TextBuffer.empty
            Metadata = defaultMetadata
        }
    ]
    Selection = Some {
        Anchor = { BlockId = "block-0"; Offset = 0 }
        Focus = { BlockId = "block-0"; Offset = 0 }
    }
    PageSetup = defaultPageSetup
    Styles = []
    Headers = []
    Footers = []
    HeaderParts = []
    FooterParts = []
    Package = emptyPackageData
}

let private isPlainTextDocument text doc =
    match doc.Blocks with
    | [ Paragraph para ] ->
        doc.PageSetup = defaultPageSetup
        && doc.Styles.IsEmpty
        && doc.Headers.IsEmpty
        && doc.Footers.IsEmpty
        && doc.HeaderParts.IsEmpty
        && doc.FooterParts.IsEmpty
        && para.Metadata = defaultMetadata
        && TextBuffer.toString para.Buffer = text
    | _ -> false

let setPlainTextContent (text: string) doc =
    if isPlainTextDocument text doc then
        doc
    else
        let buffer = TextBuffer.ofString text
        let paragraph = {
            Id =
                match doc.Blocks with
                | Paragraph para :: _ -> para.Id
                | Heading (_, para) :: _ -> para.Id
                | _ -> nextBlockId()
            Buffer = buffer
            Metadata = defaultMetadata
        }
        let selection =
            Some {
                Anchor = { BlockId = paragraph.Id; Offset = TextBuffer.length buffer }
                Focus = { BlockId = paragraph.Id; Offset = TextBuffer.length buffer }
            }
        { doc with
            Blocks = [ Paragraph paragraph ]
            Selection = selection
            PageSetup = defaultPageSetup
            Styles = []
            Headers = []
            Footers = []
            HeaderParts = []
            FooterParts = []
            Package = emptyPackageData }

/// Replace the document content with one default paragraph per input string
/// (setPlainTextContent puts the whole text into a single paragraph with
/// hard breaks; this creates real separate blocks).
let setParagraphContents (texts: string list) doc =
    let blocks =
        match texts with
        | [] -> [ Paragraph { Id = nextBlockId(); Buffer = TextBuffer.empty; Metadata = defaultMetadata } ]
        | _ ->
            texts
            |> List.map (fun t ->
                Paragraph {
                    Id = nextBlockId()
                    Buffer = TextBuffer.ofString t
                    Metadata = defaultMetadata
                })
    { doc with Blocks = blocks; Selection = None }

// ---- where a block is ---------------------------------------------------------
// Looking a paragraph up by id walked the blocks (and the cells of every
// table before it), comparing ids: the cost of a keystroke grew with the
// caret's place in the document. The top-level index a block id was last
// found at is remembered; a lookup goes there first (and a few places around
// it - blocks inserted or removed before it move it), and only then through
// the document. The hints belong to no document: one that points at another
// block is simply wrong, and corrected by the search.
let private blockPlaces = System.Collections.Generic.Dictionary<string, int>()

let private blockHolds (blockId: string) (block: Block) =
    match block with
    | Paragraph p
    | Heading (_, p) -> p.Id = blockId
    | Table t -> tableHolds blockId t
    | TextBox t -> t.Id = blockId || (t.Content |> List.exists (fun para -> para.Id = blockId))

/// The top-level block that is, or holds, the block with an id, and its
/// index among the document's blocks.
let locateBlock (doc: Document) (blockId: string) : struct (int * Block) voption =
    let hint =
        lock blockPlaces (fun () ->
            match blockPlaces.TryGetValue blockId with
            | true, index -> index
            | false, _ -> -1)
    let mutable result = ValueNone
    if hint >= 0 then
        // The blocks from a few before the hint to a few after it.
        let first = max 0 (hint - 6)
        let mutable rest = doc.Blocks
        let mutable index = 0
        while index < first && not rest.IsEmpty do
            rest <- rest.Tail
            index <- index + 1
        while result.IsNone && index <= hint + 6 && not rest.IsEmpty do
            if blockHolds blockId rest.Head then result <- ValueSome (struct (index, rest.Head))
            rest <- rest.Tail
            index <- index + 1
    if result.IsNone then
        // (The blocks themselves first: the cells of every table before a
        // paragraph of the body are not gone through to find it.)
        let mutable rest = doc.Blocks
        let mutable index = 0
        while result.IsNone && not rest.IsEmpty do
            let isBlock =
                match rest.Head with
                | Paragraph p
                | Heading (_, p) -> p.Id = blockId
                | Table t -> t.Id = blockId
                | TextBox t -> t.Id = blockId
            if isBlock then result <- ValueSome (struct (index, rest.Head))
            rest <- rest.Tail
            index <- index + 1
    if result.IsNone then
        let mutable rest = doc.Blocks
        let mutable index = 0
        while result.IsNone && not rest.IsEmpty do
            (match rest.Head with
             | Table _
             | TextBox _ -> if blockHolds blockId rest.Head then result <- ValueSome (struct (index, rest.Head))
             | _ -> ())
            rest <- rest.Tail
            index <- index + 1
    match result with
    | ValueSome (struct (index, _)) when index <> hint ->
        lock blockPlaces (fun () ->
            if blockPlaces.Count > 200000 then blockPlaces.Clear()
            blockPlaces.[blockId] <- index)
    | _ -> ()
    result

/// Find block by ID
let findBlock doc blockId =
    let idOf (block: Block) =
        match block with
        | Paragraph p
        | Heading (_, p) -> p.Id
        | Table t -> t.Id
        | TextBox t -> t.Id
    match locateBlock doc blockId with
    | ValueSome (struct (_, block)) when idOf block = blockId -> Some block
    | _ -> None

/// Try to get the editable paragraph payload from a block.
let tryGetParagraph = function
    | Paragraph p -> Some p
    | Heading (_, p) -> Some p
    | Table _
    | TextBox _ -> None

/// Get paragraph from block. Callers that can encounter tables should use tryGetParagraph.
let getParagraph = function
    | Paragraph p -> p
    | Heading (_, p) -> p
    | Table _
    | TextBox _ -> invalidOp "This block does not have a paragraph payload"

let private getBlockId = function
    | Paragraph p -> p.Id
    | Heading (_, p) -> p.Id
    | Table t -> t.Id
    | TextBox t -> t.Id

let private withParagraphBuffer block para =
    match block with
    | Paragraph _ -> Some (Paragraph para)
    | Heading (level, _) -> Some (Heading (level, para))
    | Table _
    | TextBox _ -> None

/// Replace block in document
let replaceBlock (doc: Document) blockId newBlock =
    { doc with
        Blocks = doc.Blocks |> List.map (fun b ->
            if getBlockId b = blockId then newBlock else b) }

let private sameReference left right =
    obj.ReferenceEquals(left, right)

let private mapListSharing (mapper: 'T -> 'T) (items: 'T list) =
    let mappedItems = ResizeArray<'T>()
    let mutable remaining = items
    let mutable index = 0
    let mutable firstChanged = false
    let mutable lastChangedCount = 0
    let mutable suffixAfterLastChange = items

    while not remaining.IsEmpty do
        match remaining with
        | item :: tail ->
            let mapped = mapper item
            let same = sameReference mapped item

            if firstChanged then
                mappedItems.Add mapped
            elif not same then
                let mutable prefix = items
                let mutable copied = 0
                while copied < index do
                    match prefix with
                    | prefixItem :: prefixTail ->
                        mappedItems.Add prefixItem
                        prefix <- prefixTail
                        copied <- copied + 1
                    | [] ->
                        copied <- index
                mappedItems.Add mapped
                firstChanged <- true

            if not same then
                lastChangedCount <- mappedItems.Count
                suffixAfterLastChange <- tail

            remaining <- tail
            index <- index + 1
        | [] -> ()

    if not firstChanged then
        items
    else
        let mutable result = suffixAfterLastChange
        for i = lastChangedCount - 1 downto 0 do
            result <- mappedItems.[i] :: result
        result

/// Map the editable paragraphs of the document's blocks until `finished`
/// says the rest can stay as it is (the blocks after that are not looked into).
let private mapEditableParagraphsUntil (finished: unit -> bool) (mapper: ParagraphBlock -> ParagraphBlock) (doc: Document) =
    let rec mapCell (cell: TableCell) =
        let content = mapListSharing mapper cell.Content
        let nested = if cell.NestedTables.IsEmpty then cell.NestedTables else mapListSharing mapTable cell.NestedTables
        if sameReference content cell.Content && sameReference nested cell.NestedTables then cell
        else { cell with Content = content; NestedTables = nested }

    and mapRow (row: TableRow) =
        let cells = mapListSharing mapCell row.Cells
        if sameReference cells row.Cells then row
        else { row with Cells = cells }

    and mapTable (table: TableBlock) =
        if finished () then table else
        let rows = mapListSharing mapRow table.Rows
        if sameReference rows table.Rows then table
        else { table with Rows = rows }

    let mapBlock block =
        if finished () then block else
        match block with
        | Paragraph para ->
            let mapped = mapper para
            if sameReference mapped para then block else Paragraph mapped
        | Heading (level, para) ->
            let mapped = mapper para
            if sameReference mapped para then block else Heading (level, mapped)
        | Table table ->
            let rows = mapListSharing mapRow table.Rows
            if sameReference rows table.Rows then block
            else Table { table with Rows = rows }
        | TextBox textBox ->
            let content = mapListSharing mapper textBox.Content
            if sameReference content textBox.Content then block
            else TextBox { textBox with Content = content }

    let blocks = mapListSharing mapBlock doc.Blocks
    if sameReference blocks doc.Blocks then doc
    else { doc with Blocks = blocks }

let private mapEditableParagraphs (mapper: ParagraphBlock -> ParagraphBlock) (doc: Document) =
    mapEditableParagraphsUntil (fun () -> false) mapper doc

/// Find an editable paragraph anywhere in the document, including table cells.
let findParagraph doc paragraphId =
    (match locateBlock doc paragraphId with
     | ValueSome (struct (_, block)) -> [ block ]
     | ValueNone -> [])
    |> List.tryPick (function
        | Paragraph p
        | Heading (_, p) when p.Id = paragraphId -> Some p
        | Table t ->
            let rec inTable (table: TableBlock) =
                table.Rows
                |> List.tryPick (fun row ->
                    row.Cells
                    |> List.tryPick (fun cell ->
                        match cell.Content |> List.tryFind (fun p -> p.Id = paragraphId) with
                        | Some para -> Some para
                        | None -> if cell.NestedTables.IsEmpty then None else cell.NestedTables |> List.tryPick inTable))
            inTable t
        | TextBox t ->
            t.Content |> List.tryFind (fun p -> p.Id = paragraphId)
        | _ -> None)

/// Replace an editable paragraph anywhere in the document, preserving its container.
let replaceParagraph (doc: Document) paragraphId newPara =
    // (Ids are unique: the block that holds the paragraph is found by its
    // place, the blocks before it are kept as they are and the list after it
    // is the same list.)
    match locateBlock doc paragraphId with
    | ValueNone -> doc
    | ValueSome (struct (index, block)) ->
        let single = { doc with Blocks = [ block ] }
        let mutable replaced = false
        let mapped =
            mapEditableParagraphsUntil
                (fun () -> replaced)
                (fun para ->
                    if not replaced && para.Id = paragraphId then
                        replaced <- true
                        newPara
                    else para)
                single
        match mapped.Blocks with
        | [ changed ] when not (sameReference changed block) ->
            let before = ResizeArray<Block>(index)
            let mutable rest = doc.Blocks
            for _ in 1 .. index do
                before.Add rest.Head
                rest <- rest.Tail
            let mutable blocks = changed :: rest.Tail
            for i = before.Count - 1 downto 0 do
                blocks <- before.[i] :: blocks
            { doc with Blocks = blocks }
        | _ -> doc

let private editableParagraphsInOrder doc =
    doc.Blocks
    |> List.collect (function
        | Paragraph p
        | Heading (_, p) -> [p]
        | Table t -> tableParagraphsInOrder t
        | TextBox t -> t.Content)

/// The paragraphs of the body's own text in reading order (table cells
/// included): what the caret moves through and what "everything" selects. The
/// text of a text box is a story of its own - entered by a click in the box,
/// left by Escape -, wherever the box stands among the blocks.
let private bodyParagraphsInOrder doc =
    doc.Blocks
    |> List.collect (function
        | Paragraph p
        | Heading (_, p) -> [p]
        | Table t -> tableParagraphsInOrder t
        | TextBox _ -> [])

/// The text box of the document a paragraph is in.
let textBoxOfParagraph (doc: Document) (paragraphId: BlockId) : TextBoxBlock option =
    doc.Blocks
    |> List.tryPick (function
        | TextBox box when box.Content |> List.exists (fun para -> para.Id = paragraphId) -> Some box
        | _ -> None)

let private blockParagraphs (block: Block) =
    match block with
    | Paragraph p
    | Heading (_, p) -> [p]
    | Table table -> tableParagraphsInOrder table
    | TextBox textBox -> textBox.Content

let private blockId (block: Block) =
    match block with
    | Paragraph p
    | Heading (_, p) -> p.Id
    | Table table -> table.Id
    | TextBox textBox -> textBox.Id

let private paragraphPageBreakBefore (doc: Document) (para: ParagraphBlock) =
    let styleBreak =
        para.Metadata.Style
        |> Option.bind (fun styleId ->
            doc.Styles
            |> List.tryFind (fun style -> style.StyleType = "paragraph" && style.Id = styleId)
            |> Option.bind (fun style -> style.PageBreakBefore))

    let defaultBreak =
        doc.Styles
        |> List.tryFind (fun style -> style.StyleType = "defaults")
        |> Option.bind (fun style -> style.PageBreakBefore)

    styleBreak |> Option.orElse defaultBreak |> Option.defaultValue false

/// The paragraphs a new page starts after because a section does: those that
/// end a section when the section after it does not go on on the same page.
/// How a section starts is said by its own properties - the w:type of the
/// w:sectPr at its end (on the paragraph that ends it, or the body's for the
/// last section): "continuous" on the page the section before it ends on,
/// anything else (none: nextPage) on a new page.
let sectionPageBreakParagraphs (doc: Document) : System.Collections.Generic.HashSet<BlockId> =
    let result = System.Collections.Generic.HashSet<BlockId>()
    let mutable pending : BlockId option = None
    for block in doc.Blocks do
        match block with
        | Paragraph para
        | Heading (_, para) when para.Metadata.SectionBreak.IsSome ->
            match pending with
            | Some before when para.Metadata.SectionBreak <> Some "continuous" -> result.Add before |> ignore
            | _ -> ()
            pending <- Some para.Id
        | _ -> ()
    match pending with
    | Some before when doc.PageSetup.SectionStart <> Some "continuous" -> result.Add before |> ignore
    | _ -> ()
    result

let buildPageMap (doc: Document) =
    let mutable physicalPageNumber = 1
    let mutable displayPageStart = 1
    let mutable displayPageStartPhysical = 1
    let mutable blockAnchors: PageBlockAnchor list = []
    let mutable fieldAnchors: PageFieldAnchor list = []
    let mutable bookmarkAnchors: PageBookmarkAnchor list = []

    let displayPageNumber () = displayPageStart + physicalPageNumber - displayPageStartPhysical

    let addParagraphAnchors para =
        let pageNumber = displayPageNumber ()
        para.Buffer.Runs
        |> List.iter (function
            | Text run ->
                run.Style.FieldCode
                |> Option.iter (fun code ->
                    fieldAnchors <- fieldAnchors @ [{ PageFieldBlockId = para.Id; PageFieldStartOffset = run.Start; PageFieldCode = code; PageNumber = pageNumber }])
                for (bookmark, opened) in TextBuffer.bookmarksStarted run.Style do
                    if bookmark.Name <> "" then
                        bookmarkAnchors <- bookmarkAnchors @ [{ PageBookmarkName = bookmark.Name; PageBookmarkBlockId = para.Id; PageBookmarkStartOffset = (if opened then run.Start + run.Length else run.Start); PageNumber = pageNumber }]
            | Image _ -> ())

    let restarts = pageNumberRestarts doc
    let sectionPages = sectionPageBreakParagraphs doc
    doc.Blocks
    |> List.iter (fun block ->
        let paragraphs = blockParagraphs block
        if blockAnchors.Length > 0 && paragraphs |> List.exists (paragraphPageBreakBefore doc) then
            physicalPageNumber <- physicalPageNumber + 1
        restarts
        |> List.filter (fun (restartBlockId, _, _) -> restartBlockId = blockId block)
        |> List.map (fun (_, _, start) -> start)
        |> List.tryLast
        |> Option.iter (fun start ->
            displayPageStart <- max 1 start
            displayPageStartPhysical <- physicalPageNumber)
        let pageNumber = displayPageNumber ()
        blockAnchors <- blockAnchors @ [{ PageBlockId = blockId block; PageNumber = pageNumber }]
        paragraphs |> List.iter addParagraphAnchors
        if paragraphs |> List.exists (fun para -> para.Metadata.PageBreakAfter || sectionPages.Contains para.Id) then
            physicalPageNumber <- physicalPageNumber + 1)

    {
        PageCount = max 1 physicalPageNumber
        BlockAnchors = blockAnchors
        Fields = fieldAnchors
        Bookmarks = bookmarkAnchors |> List.distinctBy (fun anchor -> anchor.PageBookmarkName)
    }

let bookmarkInfos (doc: Document) =
    let rawInfos =
        editableParagraphsInOrder doc
        |> List.collect (fun para ->
            para.Buffer.Runs
            |> List.collect (function
                | Text run ->
                    // (A run has all the bookmarks that start before it - and one opened after the last run
                    // of a paragraph, which starts at its end.)
                    TextBuffer.bookmarksStarted run.Style
                    |> List.filter (fun (bookmark, _) -> bookmark.Name <> "")
                    |> List.map (fun (bookmark, opened) ->
                        {
                            Id = bookmark.Id
                            Name = bookmark.Name
                            BlockId = para.Id
                            StartOffset = (if opened then run.Start + run.Length else run.Start)
                            EndOffset = run.Start + run.Length
                        })
                | Image _ -> []))
    rawInfos
    |> List.groupBy (fun info -> info.Name)
    |> List.map (fun (bookmarkName, infos) ->
        let firstInfo = infos |> List.head
        let blockInfos = infos |> List.filter (fun info -> info.BlockId = firstInfo.BlockId)
        {
            Id = firstInfo.Id
            Name = bookmarkName
            BlockId = firstInfo.BlockId
            StartOffset = blockInfos |> List.map (fun info -> info.StartOffset) |> List.min
            EndOffset = blockInfos |> List.map (fun info -> info.EndOffset) |> List.max
        })

type private ContentControlRange = {
    Info: ContentControlInfo
    First: ParagraphBlock
    mutable Last: ParagraphBlock
}

/// Visit physical control membership in document order without constructing
/// panel entries for paint queries or membership sets for panel queries.
let private iterContentControlParagraphs visit (doc: Document) =
    let continuations = Dictionary<string, struct (int * string)>()
    let mutable position = 0
    let add key id tag alias para =
        // Replacing the middle of a control splits its physical XML wrapper,
        // even when the remaining paragraphs retain the original group ID.
        let physicalKey =
            match continuations.TryGetValue key with
            | true, struct (previous, physicalKey) when previous >= position - 1 -> physicalKey
            | _ -> key + "@" + string position
        continuations.[key] <- struct (position, physicalKey)
        visit physicalKey id tag alias para
    let addLayer (layer: ContentControlLayer) para =
        match layer.Id with
        | Some id -> add ("group:" + layer.Group) id layer.Tag layer.Alias para
        | None -> ()
    let addOwn group prefix blockId id tag alias para =
        match id with
        | None -> () // An ID-less binding is not an editable control in the panel.
        | Some controlId ->
            let key =
                match group with
                | Some group -> "group:" + group
                | None -> prefix + blockId
            add key controlId tag alias para
    let addParagraph notifyParent (para: ParagraphBlock) =
        position <- position + 1
        notifyParent para
        let metadata = para.Metadata
        for layer in metadata.ContentControlAncestors do addLayer layer para
        addOwn metadata.ContentControlGroup "paragraph:" para.Id
            metadata.ContentControlId metadata.ContentControlTag metadata.ContentControlAlias para
    let rec addTable (table: TableBlock) (notifyParent: ParagraphBlock -> unit) =
        let addTableLayers para =
            notifyParent para
            for layer in table.ContentControlAncestors do addLayer layer para
            addOwn table.ContentControlGroup "table:" table.Id
                table.ContentControlId table.ContentControlTag table.ContentControlAlias para
        for row in table.Rows do
            for cell in row.Cells do
                if cell.NestedTables.IsEmpty then
                    for para in cell.Content do addParagraph addTableLayers para
                else
                    let content = List.toArray cell.Content
                    let anchors = List.toArray cell.NestedTableAnchors
                    let nested =
                        cell.NestedTables
                        |> List.mapi (fun index inner ->
                            let at = if index < anchors.Length then anchors.[index] else content.Length
                            inner, max 0 (min content.Length at))
                        |> List.sortBy snd
                    let mutable pending = nested
                    for position in 0 .. content.Length do
                        while not pending.IsEmpty && snd pending.Head = position do
                            addTable (fst pending.Head) addTableLayers
                            pending <- pending.Tail
                        if position < content.Length then
                            addParagraph addTableLayers content.[position]
    for block in doc.Blocks do
        match block with
        | Paragraph para | Heading (_, para) -> addParagraph ignore para
        | Table table -> addTable table ignore
        | TextBox box -> box.Content |> List.iter (addParagraph ignore)

let private contentControlRanges (doc: Document) =
    let ranges = ResizeArray<ContentControlRange>()
    let grouped = Dictionary<string, ContentControlRange>()
    iterContentControlParagraphs (fun key id tag alias para ->
        match grouped.TryGetValue key with
        | true, found -> found.Last <- para
        | _ ->
            let resolvedTag = tag |> Option.defaultValue "content-control"
            let resolvedAlias = alias |> Option.defaultValue (tag |> Option.defaultValue id)
            let found = {
                Info = { Id = id; Tag = resolvedTag; Alias = resolvedAlias; BlockId = para.Id }
                First = para; Last = para }
            ranges.Add found
            grouped.[key] <- found) doc
    List.ofSeq ranges

let contentControlInfos (doc: Document) =
    contentControlRanges doc |> List.map (fun range -> range.Info)

/// Paragraph membership of physical controls, including enclosing controls
/// and table controls absent from the cell paragraph's own metadata.
let contentControlMemberships (doc: Document) =
    let memberships = Dictionary<BlockId, HashSet<string>>()
    iterContentControlParagraphs (fun key _ _ _ para ->
        match memberships.TryGetValue para.Id with
        | true, members -> members.Add key |> ignore
        | _ ->
            let members = HashSet<string>()
            members.Add key |> ignore
            memberships.[para.Id] <- members) doc
    memberships

let commentInfos (doc: Document) =
    let infos = ResizeArray<CommentInfo>()
    let seen = HashSet<string>()
    for para in editableParagraphsInOrder doc do
        for runValue in para.Buffer.Runs do
            match runValue with
            | Text run ->
                match run.Style.CommentId with
                | Some id ->
                    let text = run.Style.CommentText |> Option.defaultValue ""
                    let key = id + "\u0000" + text
                    if seen.Add key then
                        infos.Add {
                            Id = id
                            Text = text
                            Author = run.Style.CommentAuthor
                            Date = run.Style.CommentDate
                            BlockId = para.Id
                            StartOffset = run.Start
                            EndOffset = run.Start + run.Length
                            Excerpt = TextBuffer.slice run.Start run.Length para.Buffer
                        }
                | None -> ()
            | Image _ -> ()
    List.ofSeq infos

let selectBookmark bookmarkId doc =
    match bookmarkInfos doc |> List.tryFind (fun info -> info.Id = bookmarkId || info.Name = bookmarkId) with
    | None -> doc
    | Some info ->
        { doc with
            Selection = Some {
                Anchor = { BlockId = info.BlockId; Offset = info.StartOffset }
                Focus = { BlockId = info.BlockId; Offset = info.EndOffset }
            } }

let selectContentControl controlId doc =
    match contentControlRanges doc |> List.tryFind (fun range -> range.Info.Id = controlId || range.Info.Tag = controlId) with
    | None -> doc
    | Some range ->
        { doc with
            Selection = Some {
                Anchor = { BlockId = range.First.Id; Offset = 0 }
                Focus = { BlockId = range.Last.Id; Offset = TextBuffer.length range.Last.Buffer }
            } }

/// The ids of the bookmarks of a document that have this id or this name.
let private bookmarkIdsNamed (bookmarkId: string) (doc: Document) : Set<string> =
    editableParagraphsInOrder doc
    |> List.collect (fun para ->
        para.Buffer.Runs
        |> List.collect (function
            | Text run -> TextBuffer.bookmarksStarted run.Style |> List.choose (fun (bookmark, _) -> if bookmark.Id = bookmarkId || bookmark.Name = bookmarkId then Some bookmark.Id else None)
            | Image _ -> []))
    |> Set.ofList

let removeBookmark bookmarkId doc =
    let ids = bookmarkIdsNamed bookmarkId doc
    let clearBookmarkStyle (style: TextBuffer.TextStyle) =
        // (That bookmark alone - its start and its end, wherever they are: the others of the run stay.)
        TextBuffer.withBookmarks (fun id -> not (ids.Contains id)) (fun bookmark -> bookmark.Name) style

    (doc, editableParagraphsInOrder doc)
    ||> List.fold (fun currentDoc para ->
        let newRuns =
            para.Buffer.Runs
            |> List.map (function
                | Text run -> Text { run with Style = clearBookmarkStyle run.Style }
                | Image run -> Image run)
        replaceParagraph currentDoc para.Id { para with Buffer = { para.Buffer with Runs = newRuns } })

let renameBookmark bookmarkId bookmarkName doc =
    let normalizedName =
        if System.String.IsNullOrWhiteSpace bookmarkName then "Bookmark"
        else bookmarkName.Trim().Replace(" ", "_")

    let ids = bookmarkIdsNamed bookmarkId doc
    let renameBookmarkStyle (style: TextBuffer.TextStyle) =
        TextBuffer.withBookmarks (fun _ -> true) (fun bookmark -> if ids.Contains bookmark.Id then normalizedName else bookmark.Name) style

    (doc, editableParagraphsInOrder doc)
    ||> List.fold (fun currentDoc para ->
        let newRuns =
            para.Buffer.Runs
            |> List.map (function
                | Text run -> Text { run with Style = renameBookmarkStyle run.Style }
                | Image run -> Image run)
        replaceParagraph currentDoc para.Id { para with Buffer = { para.Buffer with Runs = newRuns } })

let private matchesContentControl controlId id tag = id = Some controlId || tag = Some controlId
let private mapContentControlTables controlId update (doc: Document) =
    let matches (table: TableBlock) =
        matchesContentControl controlId table.ContentControlId table.ContentControlTag ||
        (table.ContentControlAncestors |> List.exists (fun layer -> matchesContentControl controlId layer.Id layer.Tag))
    if doc.Blocks |> List.exists (function Table table -> matches table | _ -> false) then
        { doc with
            Blocks = doc.Blocks |> List.map (function
                | Table table when matches table -> Table (update table)
                | block -> block) }
    else doc

let removeContentControl controlId doc =
    let remove (layer: ContentControlLayer) = not (matchesContentControl controlId layer.Id layer.Tag)
    (doc, editableParagraphsInOrder doc)
    ||> List.fold (fun currentDoc para ->
        let m = para.Metadata
        let inner = matchesContentControl controlId m.ContentControlId m.ContentControlTag
        if not inner && List.forall remove m.ContentControlAncestors then currentDoc
        else
            let ancestors = m.ContentControlAncestors |> List.filter remove
            let changed =
                if inner then
                    match List.rev ancestors with
                    | closest :: outer ->
                        { m with ContentControlId = closest.Id; ContentControlTag = closest.Tag; ContentControlAlias = closest.Alias
                                 ContentControlXml = closest.Xml; ContentControlGroup = Some closest.Group
                                 ContentControlAncestors = List.rev outer }
                    | [] ->
                        { m with ContentControlId = None; ContentControlTag = None; ContentControlAlias = None
                                 ContentControlXml = None; ContentControlGroup = None; ContentControlAncestors = [] }
                else { m with ContentControlAncestors = ancestors }
            replaceParagraph currentDoc para.Id { para with Metadata = changed })
    |> mapContentControlTables controlId (fun table ->
        let ancestors = table.ContentControlAncestors |> List.filter remove
        if matchesContentControl controlId table.ContentControlId table.ContentControlTag then
            match List.rev ancestors with
            | closest :: outer ->
                { table with ContentControlId = closest.Id; ContentControlTag = closest.Tag; ContentControlAlias = closest.Alias
                             ContentControlXml = closest.Xml; ContentControlGroup = Some closest.Group
                             ContentControlAncestors = List.rev outer }
            | [] ->
                { table with ContentControlId = None; ContentControlTag = None; ContentControlAlias = None
                             ContentControlXml = None; ContentControlGroup = None; ContentControlAncestors = [] }
        else { table with ContentControlAncestors = ancestors })

let updateContentControl controlId newControlId tag alias doc =
    let normalize fallback (value: string) =
        if System.String.IsNullOrWhiteSpace value then fallback else value.Trim()
    let newControlId = normalize controlId newControlId
    let tag = normalize "content-control" tag
    let alias = normalize tag alias
    let update (layer: ContentControlLayer) =
        if matchesContentControl controlId layer.Id layer.Tag then
            { layer with Id = Some newControlId; Tag = Some tag; Alias = Some alias }
        else layer
    (doc, editableParagraphsInOrder doc)
    ||> List.fold (fun currentDoc para ->
        let m = para.Metadata
        let inner = matchesContentControl controlId m.ContentControlId m.ContentControlTag
        if not inner && not (m.ContentControlAncestors |> List.exists (fun layer -> matchesContentControl controlId layer.Id layer.Tag)) then currentDoc
        else
            let changed =
                { m with ContentControlId = if inner then Some newControlId else m.ContentControlId
                         ContentControlTag = if inner then Some tag else m.ContentControlTag
                         ContentControlAlias = if inner then Some alias else m.ContentControlAlias
                         ContentControlAncestors = m.ContentControlAncestors |> List.map update }
            replaceParagraph currentDoc para.Id { para with Metadata = changed })
    |> mapContentControlTables controlId (fun table ->
        let inner = matchesContentControl controlId table.ContentControlId table.ContentControlTag
        { table with ContentControlId = if inner then Some newControlId else table.ContentControlId
                     ContentControlTag = if inner then Some tag else table.ContentControlTag
                     ContentControlAlias = if inner then Some alias else table.ContentControlAlias
                     ContentControlAncestors = table.ContentControlAncestors |> List.map update })

let private emptyTableCell () =
    {
        Content = [
            {
                Id = nextBlockId ()
                Buffer = TextBuffer.empty
                Metadata = defaultMetadata
            }
        ]
        NestedTables = []
        NestedTableAnchors = []
        EnclosingControls = []
        ColSpan = 1
        BackgroundColor = None
        Shading = None
        Width = None
        PreferredWidth = None
        BorderColor = None
        BorderWidth = None
        Borders = None
        VerticalAlignment = None
        VerticalMerge = None
        PaddingTop = None
        PaddingRight = None
        PaddingBottom = None
        PaddingLeft = None
        Origin = None
    }

let private tryFindTableCellByParagraph doc paragraphId =
    // (The one block that holds the paragraph, found by its place; the table
    // is the innermost one: that of the cell the paragraph is a paragraph of.)
    let rec inTable (table: TableBlock) =
        table.Rows
        |> List.mapi (fun rowIndex row -> rowIndex, row)
        |> List.tryPick (fun (rowIndex, row) ->
            row.Cells
            |> List.mapi (fun colIndex cell -> colIndex, cell)
            |> List.tryPick (fun (colIndex, cell) ->
                if cell.Content |> List.exists (fun para -> para.Id = paragraphId) then
                    Some (table.Id, rowIndex, colIndex)
                elif cell.NestedTables.IsEmpty then None
                else cell.NestedTables |> List.tryPick inTable))
    (match locateBlock doc paragraphId with
     | ValueSome (struct (_, block)) -> [ block ]
     | ValueNone -> [])
    |> List.tryPick (function
        | Table table -> inTable table
        | _ -> None)

/// The table with an id: a block of the document or a table in a cell.
let tryFindTable (doc: Document) (tableId: string) : TableBlock option =
    match locateBlock doc tableId with
    | ValueSome (struct (_, Table table)) -> tryFindTableIn tableId [ table ]
    | _ -> None

let private updateTable (doc: Document) tableId update =
    let rec updateIn (table: TableBlock) =
        if table.Id = tableId then update table
        elif not (tableHolds tableId table) then table
        else
            { table with
                Rows =
                    table.Rows
                    |> List.map (fun row ->
                        { row with
                            Cells =
                                row.Cells
                                |> List.map (fun cell ->
                                    if cell.NestedTables.IsEmpty then cell
                                    else { cell with NestedTables = cell.NestedTables |> List.map updateIn }) }) }
    { doc with
        Blocks =
            doc.Blocks
            |> List.map (function
                | Table table when table.Id = tableId -> Table (update table)
                | Table table when tableHolds tableId table -> Table (updateIn table)
                | block -> block) }

let private currentTableCell doc =
    match doc.Selection with
    | Some sel -> tryFindTableCellByParagraph doc sel.Focus.BlockId
    | None -> None

let private tryGetTableCellParagraph doc tableId rowIndex colIndex =
    tryFindTable doc tableId
    |> Option.bind (fun table -> table.Rows |> List.tryItem rowIndex)
    |> Option.bind (fun row -> row.Cells |> List.tryItem colIndex)
    |> Option.bind (fun cell -> cell.Content |> List.tryHead)

/// Find block index by ID
let findBlockIndex doc blockId =
    match locateBlock doc blockId with
    | ValueSome (struct (index, block)) when getBlockId block = blockId -> Some index
    | _ -> None

/// Get normalized range with block IDs for multi-block selection
let getNormalizedRange (sel: Selection) (doc: Document) =
    // (A range of the body is made of the body's paragraphs: a text box that stands among the blocks
    // it reaches over is not in it. A range in a box that is not entered - set by a host - is the box's.)
    let body = bodyParagraphsInOrder doc
    let indexIn (paragraphs: ParagraphBlock list) (id: BlockId) = paragraphs |> List.tryFindIndex (fun p -> p.Id = id)
    let paragraphs, anchorIdx, focusIdx =
        match indexIn body sel.Anchor.BlockId, indexIn body sel.Focus.BlockId with
        | (Some _ as a), (Some _ as f) -> body, a, f
        | _ ->
            let all = editableParagraphsInOrder doc
            all, indexIn all sel.Anchor.BlockId, indexIn all sel.Focus.BlockId

    match anchorIdx, focusIdx with
    | Some aIdx, Some fIdx ->
        // Determine selection direction
        let (startIdx, startPos, endIdx, endPos) =
            if aIdx < fIdx then
                (aIdx, sel.Anchor, fIdx, sel.Focus)
            elif aIdx > fIdx then
                (fIdx, sel.Focus, aIdx, sel.Anchor)
            else
                // Same block, compare offsets
                if sel.Anchor.Offset <= sel.Focus.Offset then
                    (aIdx, sel.Anchor, fIdx, sel.Focus)
                else
                    (fIdx, sel.Focus, aIdx, sel.Anchor)

        let blockIds =
            paragraphs
            |> List.mapi (fun i p -> (i, p))
            |> List.filter (fun (i, _) -> i >= startIdx && i <= endIdx)
            |> List.map (fun (_, p) -> p.Id)

        (startPos, endPos, blockIds)
    | _ ->
        // Fallback if blocks not found
        (sel.Anchor, sel.Focus, [])

/// The paragraphs a non-collapsed selection touches, in document order, each
/// with the selected range [start, end) inside it. One pass over the document.
/// Call `visit paragraph rangeStart rangeEnd` for each paragraph the selection
/// touches, in document order, until it answers false. False when there is
/// no range selected (no selection, a caret, an end that is not in the document).
let forEachSelectedParagraphRange (doc: Document) (visit: ParagraphBlock -> int -> int -> bool) : bool =
    match doc.Selection with
    | None -> false
    | Some sel when sel.IsCollapsed -> false
    | Some sel ->
        // (One walk over the paragraphs, from the first end of the selection
        // to the other: the selection may be the whole of a long document.)
        let anchorId = sel.Anchor.BlockId
        let focusId = sel.Focus.BlockId
        // Both ends are in the document: the walk below may stop early.
        if (findParagraph doc anchorId).IsNone || (findParagraph doc focusId).IsNone then false
        else
            // 0: before the selection, 1: in it, 2: past it (or told to stop).
            let mutable state = 0
            let mutable endOffset = 0
            let step (para: ParagraphBlock) =
                if state = 0 then
                    let isAnchor = para.Id = anchorId
                    let isFocus = para.Id = focusId
                    if isAnchor && isFocus then
                        visit para (min sel.Anchor.Offset sel.Focus.Offset) (max sel.Anchor.Offset sel.Focus.Offset) |> ignore
                        state <- 2
                    elif isAnchor || isFocus then
                        endOffset <- (if isAnchor then sel.Focus.Offset else sel.Anchor.Offset)
                        state <- (if visit para (if isAnchor then sel.Anchor.Offset else sel.Focus.Offset) (TextBuffer.length para.Buffer) then 1 else 2)
                elif state = 1 then
                    if para.Id = anchorId || para.Id = focusId then
                        visit para 0 endOffset |> ignore
                        state <- 2
                    elif not (visit para 0 (TextBuffer.length para.Buffer)) then state <- 2
            let mutable rest = doc.Blocks
            while state < 2 && not rest.IsEmpty do
                (match rest.Head with
                 | Paragraph p
                 | Heading (_, p) -> step p
                 | Table t -> for p in tableParagraphsInOrder t do (if state < 2 then step p)
                 // (A range that started in the body does not take in a box it reaches over.)
                 | TextBox t -> if state = 0 then (for p in t.Content do (if state < 2 then step p)))
                rest <- rest.Tail
            true

let selectedParagraphRanges (doc: Document) : (ParagraphBlock * int * int) list =
    let found = ResizeArray<ParagraphBlock * int * int>()
    if forEachSelectedParagraphRange doc (fun para rangeStart rangeEnd -> found.Add ((para, rangeStart, rangeEnd)); true) then
        let mutable ranges = []
        for i = found.Count - 1 downto 0 do
            ranges <- found.[i] :: ranges
        ranges
    else []

/// Edit the list of paragraphs a paragraph is in when that is a table cell's
/// (of a table in a cell too) or a text box's. `edit` gets the list and, for
/// a cell, the places of the tables in it (the number of the cell's
/// paragraphs before each), and gives both back; None leaves the document as
/// it is. None when the paragraph is a block of the document.
let private editContainerOf (doc: Document) (paragraphId: string) (edit: ParagraphBlock list -> int list -> (ParagraphBlock list * int list) option) : Document option =
    match locateBlock doc paragraphId with
    | ValueNone -> None
    | ValueSome (struct (index, block)) ->
        let mutable changed = false
        let rec editTable (table: TableBlock) =
            if changed || not (tableHolds paragraphId table) then table
            else { table with Rows = table.Rows |> List.map (fun row -> { row with Cells = row.Cells |> List.map editCell }) }
        and editCell (cell: TableCell) =
            if changed then cell
            elif cell.Content |> List.exists (fun p -> p.Id = paragraphId) then
                match edit cell.Content cell.NestedTableAnchors with
                | Some (content, anchors) ->
                    changed <- true
                    { cell with Content = content; NestedTableAnchors = anchors }
                | None -> cell
            elif cell.NestedTables.IsEmpty then cell
            else { cell with NestedTables = cell.NestedTables |> List.map editTable }
        let edited =
            match block with
            | Table table -> Table (editTable table)
            | TextBox box when box.Content |> List.exists (fun p -> p.Id = paragraphId) ->
                match edit box.Content [] with
                | Some (content, _) ->
                    changed <- true
                    TextBox { box with Content = content }
                | None -> block
            | _ -> block
        if not changed then None
        else Some { doc with Blocks = doc.Blocks |> List.mapi (fun i b -> if i = index then edited else b) }

/// The properties of the first part of a paragraph that is split: what stands
/// at the paragraph's end - the end of its section (Word keeps it in the
/// paragraph mark), a page or column break after its text - goes with the
/// second part, which ends where the paragraph ended. (Both parts had them:
/// Enter in the last paragraph of a section made a section of each part, and
/// Enter before a break a second break.)
let private withoutParagraphEnd (metadata: BlockMetadata) : BlockMetadata =
    if metadata.SectionBreak.IsNone && metadata.SectionPageSetup.IsNone && not metadata.PageBreakAfter && not metadata.ColumnBreakAfter
       && metadata.HeaderReferences.IsEmpty && metadata.FooterReferences.IsEmpty then metadata
    else
        { metadata with
            PageBreakAfter = false; ColumnBreakAfter = false
            SectionBreak = None; SectionColumns = None; SectionPageSetup = None; SectionPageNumberStart = None
            SectionPageNumberFormat = None; SectionVerticalAlign = None
            HeaderReferences = []; FooterReferences = [] }

/// The properties of the second part of a paragraph that is split: not that
/// the paragraph started a page or a column when Word last laid it out
/// (RenderedPageBreakBefore, from its w:lastRenderedPageBreak) - the second
/// part did not; its first part did. (The column flow starts a new column at
/// a paragraph that has it: the second part went to the next column in the
/// editor and stayed in its column in the saved document, which does not say it.)
let private withoutParagraphStart (metadata: BlockMetadata) : BlockMetadata =
    let metadata =
        if metadata.RenderedPageBreakBefore.IsNone then metadata
        else { metadata with RenderedPageBreakBefore = None }
    // (Nor the page break the paragraph started with: it stands before the first part.)
    if ParagraphExplicit.has ParagraphExplicit.PageBreakRunBefore metadata then
        { metadata with PageBreakBefore = false; Explicit = metadata.Explicit &&& ~~~ParagraphExplicit.PageBreakRunBefore }
    else metadata

/// Whether source control properties permit an additional paragraph.
let private controlDisallowsParagraphBreak (xml: string option) =
    xml
    |> Option.bind (fun properties ->
        XmlDsl.tryScan ("<root>" + properties + "</root>") (System.Collections.Generic.HashSet<string>()))
    |> Option.bind (XmlPatch.child "sdtPr")
    |> Option.exists (fun properties ->
        let kind name = XmlPatch.child name properties |> Option.isSome
        if [ "date"; "picture"; "checkbox"; "comboBox"; "dropDownList"; "equation"; "docPartList"; "docPartObj" ] |> List.exists kind then true
        else
            match XmlPatch.child "text" properties with
            | None -> false
            | Some text ->
                XmlPatch.attributes text
                |> Map.toList
                |> List.tryPick (fun (name, value) -> if XmlDsl.Ns.stripNs name = "multiLine" then Some value else None)
                |> Option.exists (fun value -> value = "1" || value = "true" || value = "on")
                |> not)

let private canSplitControlParagraph (doc: Document) (para: ParagraphBlock) offset =
    let paragraphControls =
        para.Metadata.ContentControlXml :: ((para.Metadata.ContentControlAncestors @ para.Metadata.EnclosingCellControls) |> List.map (fun layer -> layer.Xml))
    let cellControls =
        tryFindTableCellByParagraph doc para.Id
        |> Option.bind (fun (tableId, rowIndex, colIndex) ->
            tryFindTable doc tableId
            |> Option.bind (fun table -> table.Rows |> List.tryItem rowIndex)
            |> Option.bind (fun row -> row.Cells |> List.tryItem colIndex))
        |> Option.map (fun cell -> cell.EnclosingControls |> List.map (fun layer -> layer.Xml))
        |> Option.defaultValue []
    let inlineAt side =
        para.Buffer.Runs
        |> List.collect (function
            | TextBuffer.Text run when side run.Start (run.Start + run.Length) -> run.Style.ContentControls |> List.map (fun layer -> layer.Group)
            | _ -> [])
        |> Set.ofList
    let left = inlineAt (fun start _ -> start < offset)
    let right = inlineAt (fun _ finish -> finish > offset)
    not (List.exists controlDisallowsParagraphBreak (paragraphControls @ cellControls))
    && (Set.intersect left right |> Set.isEmpty)

let private controlIsPlainText (xml: string option) =
    xml
    |> Option.bind (fun properties -> XmlDsl.tryScan ("<root>" + properties + "</root>") (System.Collections.Generic.HashSet<string>()))
    |> Option.bind (XmlPatch.child "sdtPr")
    |> Option.bind (XmlPatch.child "text")
    |> Option.isSome

/// Split block at cursor, creating two blocks
let splitBlock doc =
    match doc.Selection with
    | None -> doc
    | Some sel when findParagraph doc sel.Focus.BlockId |> Option.exists (fun para -> not (canSplitControlParagraph doc para sel.Focus.Offset)) -> doc
    | Some sel ->
        match findBlock doc sel.Focus.BlockId with
        | None ->
            // A paragraph of a table cell or a text box: split within the
            // list it is in; the tables of the cell after it stay after it.
            match findParagraph doc sel.Focus.BlockId with
            | None -> doc
            | Some para ->
                let offset = sel.Focus.Offset
                let beforeBuffer = TextBuffer.sliceBuffer 0 offset para.Buffer
                let afterBuffer = TextBuffer.sliceBuffer offset (TextBuffer.length para.Buffer - offset) para.Buffer
                let newBlockId = nextBlockId()
                let newPara = { Id = newBlockId; Buffer = afterBuffer; Metadata = { withoutParagraphStart para.Metadata with Origin = likeOrigin para.Metadata.Origin } }
                let updatedPara = { para with Buffer = beforeBuffer; Metadata = withoutParagraphEnd para.Metadata }
                let split =
                    editContainerOf doc para.Id (fun content anchors ->
                        let index = content |> List.findIndex (fun p -> p.Id = para.Id)
                        Some (content |> List.collect (fun p -> if p.Id = para.Id then [ updatedPara; newPara ] else [ p ]),
                              anchors |> List.map (fun anchor -> if anchor > index then anchor + 1 else anchor)))
                match split with
                | None -> doc
                | Some split ->
                    let caret = { BlockId = newBlockId; Offset = 0 }
                    { split with Selection = Some { Anchor = caret; Focus = caret } }
        | Some (Table _)
        | Some (TextBox _) -> doc
        | Some block ->
            let para = getParagraph block
            let offset = sel.Focus.Offset

            // Split buffer at offset, keeping each side's runs (formatting, images)
            let beforeBuffer = TextBuffer.sliceBuffer 0 offset para.Buffer
            let afterBuffer = TextBuffer.sliceBuffer offset (TextBuffer.length para.Buffer - offset) para.Buffer

            // Create new block for the second part (carry over metadata)
            let newBlockId = nextBlockId()
            let newPara = { Id = newBlockId; Buffer = afterBuffer; Metadata = { withoutParagraphStart para.Metadata with Origin = likeOrigin para.Metadata.Origin } }
            let newBlock =
                match block with
                | Paragraph _ -> Paragraph newPara
                | Heading (level, _) -> Heading (level, newPara)
                | Table t -> Table t
                | TextBox t -> TextBox t

            // Update first block with before text (keep existing metadata)
            let updatedPara = { para with Buffer = beforeBuffer; Metadata = withoutParagraphEnd para.Metadata }
            let updatedBlock =
                match block with
                | Paragraph _ -> Paragraph updatedPara
                | Heading (level, _) -> Heading (level, updatedPara)
                | Table t -> Table t
                | TextBox t -> TextBox t

            // Insert new block after current block
            match findBlockIndex doc para.Id with
            | None -> doc
            | Some idx ->
                let newBlocks =
                    doc.Blocks
                    |> List.mapi (fun i b ->
                        if i = idx then [updatedBlock; newBlock]
                        else [b])
                    |> List.concat

                // Move cursor to start of new block
                let newSel = {
                    Anchor = { BlockId = newBlockId; Offset = 0 }
                    Focus = { BlockId = newBlockId; Offset = 0 }
                }
                { doc with Blocks = newBlocks; Selection = Some newSel }

/// Merge block with previous block (for Backspace at start)
let mergeWithPrevious doc blockId =
    match findBlockIndex doc blockId with
    | None ->
        // A paragraph of a table cell or a text box joins the paragraph
        // before it in the same cell / box - not across a table in the cell.
        match findParagraph doc blockId with
        | None -> doc
        | Some currentPara ->
            let mutable caret = None
            let merged =
                editContainerOf doc blockId (fun content anchors ->
                    let index = content |> List.findIndex (fun p -> p.Id = blockId)
                    if index = 0 || List.contains index anchors then None
                    else
                        let prevPara = content.[index - 1]
                        caret <- Some { BlockId = prevPara.Id; Offset = TextBuffer.length prevPara.Buffer }
                        let mergedPara = { prevPara with Buffer = TextBuffer.append prevPara.Buffer currentPara.Buffer }
                        Some (content |> List.mapi (fun i p -> i, p) |> List.choose (fun (i, p) -> if i = index - 1 then Some mergedPara elif i = index then None else Some p),
                              anchors |> List.map (fun anchor -> if anchor > index then anchor - 1 else anchor)))
            match merged, caret with
            | Some merged, Some caret -> { merged with Selection = Some { Anchor = caret; Focus = caret } }
            | _ -> doc
    | Some idx when idx = 0 -> doc  // First block, can't merge
    | Some idx ->
        match findBlock doc blockId with
        | None -> doc
        | Some currentBlock ->
            let prevBlock = doc.Blocks.[idx - 1]
            // A table or text box above is not something a paragraph merges into.
            match tryGetParagraph currentBlock, tryGetParagraph prevBlock with
            | None, _
            | _, None -> doc
            | Some currentPara, Some prevPara ->

            // Merge buffers: previous + current, keeping both sides' runs
            let prevLength = TextBuffer.length prevPara.Buffer
            let mergedBuffer = TextBuffer.append prevPara.Buffer currentPara.Buffer

            // Update previous block with merged buffer
            let mergedPara = { prevPara with Buffer = mergedBuffer }
            let mergedBlock =
                match prevBlock with
                | Paragraph _ -> Paragraph mergedPara
                | Heading (level, _) -> Heading (level, mergedPara)
                | Table t -> Table t
                | TextBox t -> TextBox t

            // Remove current block
            let newBlocks =
                doc.Blocks
                |> List.mapi (fun i b ->
                    if i = idx - 1 then Some mergedBlock
                    elif i = idx then None
                    else Some b)
                |> List.choose id

            // Move cursor to junction point
            let newSel = {
                Anchor = { BlockId = prevPara.Id; Offset = prevLength }
                Focus = { BlockId = prevPara.Id; Offset = prevLength }
            }
            { doc with Blocks = newBlocks; Selection = Some newSel }

/// Merge block with next block (for Delete at end)
let mergeWithNext doc blockId =
    match findBlockIndex doc blockId with
    | None ->
        // (As mergeWithPrevious: within the cell or the text box.)
        match findParagraph doc blockId with
        | None -> doc
        | Some currentPara ->
            editContainerOf doc blockId (fun content anchors ->
                let index = content |> List.findIndex (fun p -> p.Id = blockId)
                if index >= content.Length - 1 || List.contains (index + 1) anchors then None
                else
                    let nextPara = content.[index + 1]
                    let mergedPara = { currentPara with Buffer = TextBuffer.append currentPara.Buffer nextPara.Buffer }
                    Some (content |> List.mapi (fun i p -> i, p) |> List.choose (fun (i, p) -> if i = index then Some mergedPara elif i = index + 1 then None else Some p),
                          anchors |> List.map (fun anchor -> if anchor > index + 1 then anchor - 1 else anchor)))
            |> Option.defaultValue doc
    | Some idx when idx >= doc.Blocks.Length - 1 -> doc  // Last block, can't merge
    | Some idx ->
        match findBlock doc blockId with
        | None -> doc
        | Some currentBlock ->
            let nextBlock = doc.Blocks.[idx + 1]
            match tryGetParagraph currentBlock, tryGetParagraph nextBlock with
            | None, _
            | _, None -> doc
            | Some currentPara, Some nextPara ->

            // Merge buffers: current + next, keeping both sides' runs
            let mergedBuffer = TextBuffer.append currentPara.Buffer nextPara.Buffer

            // Update current block with merged buffer
            let mergedPara = { currentPara with Buffer = mergedBuffer }
            let mergedBlock =
                match currentBlock with
                | Paragraph _ -> Paragraph mergedPara
                | Heading (level, _) -> Heading (level, mergedPara)
                | Table t -> Table t
                | TextBox t -> TextBox t

            // Remove next block
            let newBlocks =
                doc.Blocks
                |> List.mapi (fun i b ->
                    if i = idx then Some mergedBlock
                    elif i = idx + 1 then None
                    else Some b)
                |> List.choose id

            // Keep cursor at current position
            { doc with Blocks = newBlocks }

/// Delete selected text across blocks
let deleteSelection doc =
    match doc.Selection with
    | None -> doc
    | Some sel when sel.IsCollapsed -> doc  // No selection to delete
    | Some sel ->
        let (startPos, endPos, blockIds) = getNormalizedRange sel doc

        match blockIds with
        | [] -> doc
        | [singleBlock] when singleBlock = startPos.BlockId && singleBlock = endPos.BlockId ->
            // Single block selection
            match findParagraph doc singleBlock with
            | None -> doc
            | Some para ->
                let len = endPos.Offset - startPos.Offset
                let newBuffer = TextBuffer.delete startPos.Offset len para.Buffer
                let newPara = { para with Buffer = newBuffer }
                let newDoc = replaceParagraph doc para.Id newPara

                // Move cursor to start of selection
                let newSel = {
                    Anchor = startPos
                    Focus = startPos
                }
                { newDoc with Selection = Some newSel }
        | _ ->
            // Multi-block selection: keep the text before the selection
            // start and after its end, drop everything in between.
            let firstBlockId = List.head blockIds
            let lastBlockId = List.last blockIds

            match findParagraph doc firstBlockId, findParagraph doc lastBlockId with
            | Some firstPara, Some lastPara ->
                // Index of the top-level block holding a paragraph, and
                // whether the paragraph is that block (as opposed to sitting
                // in a table cell or text box inside it).
                let topLevel paragraphId =
                    let rec cellKey (table: TableBlock) =
                        table.Rows
                        |> List.tryPick (fun row ->
                            row.Cells
                            |> List.tryPick (fun cell ->
                                // (The list of paragraphs it is in, named by the first of them.)
                                if cell.Content |> List.exists (fun p -> p.Id = paragraphId) then Some cell.Content.Head.Id
                                elif cell.NestedTables.IsEmpty then None
                                else cell.NestedTables |> List.tryPick cellKey))
                    match locateBlock doc paragraphId with
                    | ValueNone -> None
                    | ValueSome (struct (index, block)) ->
                        match block with
                        | Paragraph p
                        | Heading (_, p) -> if p.Id = paragraphId then Some (index, true, "") else None
                        | Table table -> cellKey table |> Option.map (fun key -> index, false, key)
                        | TextBox textBox ->
                            if textBox.Content |> List.exists (fun p -> p.Id = paragraphId) then Some (index, false, "") else None

                match topLevel firstBlockId, topLevel lastBlockId with
                | Some (firstIndex, firstIsBlock, firstCell), Some (lastIndex, lastIsBlock, lastCell) ->
                    // Content before the selection start and after its end,
                    // with runs intact
                    let beforeBuffer = TextBuffer.sliceBuffer 0 startPos.Offset firstPara.Buffer
                    let lastBufferLen = TextBuffer.length lastPara.Buffer
                    let afterBuffer = TextBuffer.sliceBuffer endPos.Offset (lastBufferLen - endPos.Offset) lastPara.Buffer

                    // The two ends join into one paragraph only when they
                    // live in the same container: both in the body, or both
                    // in the same cell / text box. Across a table boundary
                    // each end keeps its own paragraph.
                    let joins =
                        (firstIsBlock && lastIsBlock)
                        || (not firstIsBlock && not lastIsBlock && firstIndex = lastIndex && firstCell = lastCell)
                    let newFirstPara =
                        { firstPara with Buffer = if joins then TextBuffer.append beforeBuffer afterBuffer else beforeBuffer }
                    let newLastPara = { lastPara with Buffer = afterBuffer }
                    let removeIds =
                        blockIds
                        |> List.filter (fun id -> id <> firstBlockId && id <> lastBlockId)
                        |> Set.ofList

                    // Paragraphs of a cell or text box. A container never
                    // ends up with no paragraph at all: an emptied one keeps
                    // its first paragraph, blank.
                    let editContent (content: ParagraphBlock list) =
                        let touched =
                            content |> List.exists (fun p -> p.Id = firstBlockId || p.Id = lastBlockId || Set.contains p.Id removeIds)
                        if not touched then content
                        else
                            let kept =
                                content
                                |> List.choose (fun p ->
                                    if p.Id = firstBlockId then Some newFirstPara
                                    elif p.Id = lastBlockId then (if joins then None else Some newLastPara)
                                    elif Set.contains p.Id removeIds then None
                                    else Some p)
                            match kept, content with
                            | [], head :: _ -> [ { head with Buffer = TextBuffer.empty } ]
                            | _ -> kept

                    // A table in a cell goes when the selection runs from
                    // a paragraph of the cell before it to one after it;
                    // the others keep their place among what is left.
                    let rec editTable (table: TableBlock) : TableBlock =
                        { table with Rows = table.Rows |> List.map (fun row -> { row with Cells = row.Cells |> List.map editCell }) }
                    and editCell (cell: TableCell) : TableCell =
                        let content = editContent cell.Content
                        if cell.NestedTables.IsEmpty then { cell with Content = content }
                        else
                            let selected (p: ParagraphBlock) = p.Id = firstBlockId || p.Id = lastBlockId || Set.contains p.Id removeIds
                            let original = cell.Content |> List.toArray
                            let anchors = cell.NestedTableAnchors |> List.toArray
                            let keptIds = content |> List.map (fun p -> p.Id) |> Set.ofList
                            let tables =
                                cell.NestedTables
                                |> List.mapi (fun i nested ->
                                    let anchor = (if i < anchors.Length then anchors.[i] else original.Length) |> max 0 |> min original.Length
                                    let before = original |> Array.take anchor
                                    if Array.exists selected before && (original |> Array.skip anchor |> Array.exists selected) then None
                                    else Some (editTable nested, before |> Array.filter (fun p -> keptIds.Contains p.Id) |> Array.length))
                                |> List.choose id
                            { cell with Content = content; NestedTables = List.map fst tables; NestedTableAnchors = List.map snd tables }

                    // A text box stands among the blocks next to the paragraph it is anchored
                    // in, and goes with the start of that paragraph, where its anchor is: with
                    // a block that is wholly inside the selection, with the first paragraph
                    // when the selection starts at its first character, with the last one
                    // when the selection reaches into it. It stays when its paragraph's start
                    // is not selected - the selection begins further on in it, or ends before
                    // its first character (it is then anchored in the paragraph that one
                    // becomes part of).
                    let goneAnchors = System.Collections.Generic.HashSet<BlockId>()
                    doc.Blocks
                    |> List.iteri (fun index block ->
                        if index > firstIndex && index < lastIndex then
                            match block with
                            | Paragraph p
                            | Heading (_, p) -> goneAnchors.Add p.Id |> ignore
                            | Table table -> goneAnchors.Add table.Id |> ignore
                            | TextBox _ -> ())
                    if firstIsBlock && startPos.Offset = 0 then goneAnchors.Add firstBlockId |> ignore
                    if lastIsBlock && endPos.Offset > 0 then goneAnchors.Add lastBlockId |> ignore
                    let newBlocks =
                        doc.Blocks
                        |> List.mapi (fun index block -> index, block)
                        |> List.choose (fun (index, block) ->
                            match block with
                            // (A box an end of the selection is in: its own text is edited.)
                            | TextBox textBox when index = firstIndex || index = lastIndex ->
                                Some (TextBox { textBox with Content = editContent textBox.Content })
                            | TextBox textBox ->
                                match textBox.AnchorBlockId with
                                | Some anchor when goneAnchors.Contains anchor -> None
                                | Some anchor when anchor = lastBlockId && lastIsBlock && joins -> Some (TextBox { textBox with AnchorBlockId = Some firstBlockId })
                                | Some anchor when (anchor = firstBlockId && firstIsBlock) || (anchor = lastBlockId && lastIsBlock) -> Some block
                                // (Anchored elsewhere, or nowhere: it goes when it stands inside the selection.)
                                | _ when index > firstIndex && index < lastIndex -> None
                                | _ -> Some block
                            | _ ->
                            if index > firstIndex && index < lastIndex then
                                // Wholly inside the selection (a table too).
                                None
                            else
                                match block with
                                | Paragraph p when p.Id = firstBlockId -> Some (Paragraph newFirstPara)
                                | Heading (level, p) when p.Id = firstBlockId -> Some (Heading (level, newFirstPara))
                                | Paragraph p when p.Id = lastBlockId -> if joins then None else Some (Paragraph newLastPara)
                                | Heading (level, p) when p.Id = lastBlockId -> if joins then None else Some (Heading (level, newLastPara))
                                | Table table when index = firstIndex || index = lastIndex ->
                                    Some (Table (editTable table))
                                | block -> Some block)

                    // Move cursor to the start of the deleted range
                    let caret = { BlockId = firstBlockId; Offset = TextBuffer.length beforeBuffer }
                    { doc with Blocks = newBlocks; Selection = Some { Anchor = caret; Focus = caret } }
                | _ -> doc
            | _ -> doc

let private isHighSurrogate (ch: char) =
    let code = int ch
    code >= 0xD800 && code <= 0xDBFF

let private isLowSurrogate (ch: char) =
    let code = int ch
    code >= 0xDC00 && code <= 0xDFFF

/// How many UTF-16 code units the character after `offset` occupies: 2 for a
/// surrogate pair (emoji and other astral characters), otherwise 1. Stepping
/// by this keeps the caret and deletions from splitting a pair.
let private forwardStep buffer offset =
    if offset >= 0 && offset + 2 <= TextBuffer.length buffer then
        let pair = TextBuffer.slice offset 2 buffer
        if pair.Length = 2 && isHighSurrogate pair.[0] && isLowSurrogate pair.[1] then 2 else 1
    else 1

/// How many UTF-16 code units the character before `offset` occupies.
let private backwardStep buffer offset =
    if offset >= 2 && offset <= TextBuffer.length buffer then
        let pair = TextBuffer.slice (offset - 2) 2 buffer
        if pair.Length = 2 && isHighSurrogate pair.[0] && isLowSurrogate pair.[1] then 2 else 1
    else 1

/// Delete forward (Delete key)
// ---- the mark of a comment -----------------------------------------------------
//
// A comment that was read stands in the text as its mark, "[Comment 5]": the
// place of the reference to it (w:commentReference), shown so. The mark is
// one thing: the caret is before it or after it, never in it - an arrow key
// steps over it, a click in it goes to its nearer end, a selection has all of
// it or nothing - text is typed before or after it, and Backspace after it
// (Delete before it) takes all of it. (Text typed into its middle used to
// part it in two marks of the same comment.)

/// Where the marks of comments are in a paragraph's text: (start, end) of each.
let commentMarkRanges (para: ParagraphBlock) : (int * int) list =
    if not (para.Buffer.Runs |> List.exists (function Text run -> run.Style.CommentId.IsSome | Image _ -> false)) then []
    else
        let found = ResizeArray<int * int>()
        // (The runs of one comment that stand together: a part of the mark may be formatted.)
        let mutable current : (string * int * int) option = None
        let close () =
            match current with
            | Some (id, start, fin) ->
                if Rope.slice start (fin - start) para.Buffer.Rope = "[Comment " + id + "]" then found.Add ((start, fin))
                current <- None
            | None -> ()
        for run in para.Buffer.Runs do
            match run with
            | Text text when text.Style.CommentId.IsSome ->
                let id = text.Style.CommentId.Value
                match current with
                | Some (known, start, fin) when known = id && fin = text.Start -> current <- Some (id, start, text.Start + text.Length)
                | _ ->
                    close ()
                    current <- Some (id, text.Start, text.Start + text.Length)
            | _ -> close ()
        close ()
        List.ofSeq found

/// `doc` with no end of its selection inside the mark of a comment. `previous`
/// is the selection before the change that led to `doc`: a caret that came one
/// step into a mark from one of its ends goes on to the other end; any other
/// position in a mark becomes the nearer end.
let snapSelectionToMarks (previous: Selection option) (doc: Document) : Document =
    match doc.Selection with
    | None -> doc
    | Some sel when (match previous with Some was -> obj.ReferenceEquals(was, sel) | None -> false) -> doc
    | Some sel ->
        let snap (position: Position) (was: Position option) : Position =
            match findParagraph doc position.BlockId with
            | None -> position
            | Some para ->
                match commentMarkRanges para |> List.tryFind (fun (start, fin) -> position.Offset > start && position.Offset < fin) with
                | None -> position
                | Some (start, fin) ->
                    let stepped =
                        match was with
                        | Some before when before.BlockId = position.BlockId && before.Offset = start && position.Offset = start + 1 -> Some fin
                        | Some before when before.BlockId = position.BlockId && before.Offset = fin && position.Offset = fin - 1 -> Some start
                        | _ -> None
                    let offset = stepped |> Option.defaultValue (if position.Offset - start < fin - position.Offset then start else fin)
                    { position with Offset = offset }
        let focus = snap sel.Focus (previous |> Option.map (fun was -> was.Focus))
        let anchor = if sel.Anchor = sel.Focus then focus else snap sel.Anchor None
        if focus.Offset = sel.Focus.Offset && anchor.Offset = sel.Anchor.Offset then doc
        else { doc with Selection = Some { Anchor = anchor; Focus = focus } }

/// The mark of a comment that ends (`atEnd`) or starts at a collapsed caret, selected whole.
let private markSelectedAtCaret (atEnd: bool) (doc: Document) : Document option =
    match doc.Selection with
    | Some sel when sel.IsCollapsed ->
        findParagraph doc sel.Focus.BlockId
        |> Option.bind (fun para -> commentMarkRanges para |> List.tryFind (fun (start, fin) -> (if atEnd then fin else start) = sel.Focus.Offset))
        |> Option.map (fun (start, fin) -> { doc with Selection = Some { Anchor = { sel.Focus with Offset = start }; Focus = { sel.Focus with Offset = fin } } })
    | _ -> None

let deleteForward doc =
    match doc.Selection with
    | None -> doc
    | Some sel when not sel.IsCollapsed ->
        // Delete selection if active
        deleteSelection doc
    | Some _ when (markSelectedAtCaret false doc).IsSome -> deleteSelection (markSelectedAtCaret false doc).Value
    | Some sel ->
        match findParagraph doc sel.Focus.BlockId with
        | None -> doc
        | Some para ->
            let maxOffset = TextBuffer.length para.Buffer

            if sel.Focus.Offset >= maxOffset then
                mergeWithNext doc sel.Focus.BlockId
            else
                let newBuffer = TextBuffer.delete sel.Focus.Offset (forwardStep para.Buffer sel.Focus.Offset) para.Buffer
                let newPara = { para with Buffer = newBuffer }
                replaceParagraph doc para.Id newPara

/// Move cursor up to previous block
let moveCursorUp doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        match findBlockIndex doc sel.Focus.BlockId with
        | None -> doc
        | Some idx when idx = 0 -> doc  // Already at first block
        | Some idx ->
            // Nearest paragraph above, stepping over tables and text boxes
            match doc.Blocks |> List.take idx |> List.rev |> List.tryPick tryGetParagraph with
            | None -> doc
            | Some prevPara ->
            let prevLen = TextBuffer.length prevPara.Buffer

            // Move to end of previous block
            let newSel = {
                Anchor = { BlockId = prevPara.Id; Offset = prevLen }
                Focus = { BlockId = prevPara.Id; Offset = prevLen }
            }
            { doc with Selection = Some newSel }

/// Move cursor down to next block
let moveCursorDown doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        match findBlockIndex doc sel.Focus.BlockId with
        | None -> doc
        | Some idx when idx >= doc.Blocks.Length - 1 -> doc  // Already at last block
        | Some idx ->
            // Nearest paragraph below, stepping over tables and text boxes
            match doc.Blocks |> List.skip (idx + 1) |> List.tryPick tryGetParagraph with
            | None -> doc
            | Some nextPara ->

            // Move to start of next block
            let newSel = {
                Anchor = { BlockId = nextPara.Id; Offset = 0 }
                Focus = { BlockId = nextPara.Id; Offset = 0 }
            }
            { doc with Selection = Some newSel }

/// Move cursor to start of current block (Home key)
let moveCursorToLineStart doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        let newSel = {
            Anchor = { sel.Focus with Offset = 0 }
            Focus = { sel.Focus with Offset = 0 }
        }
        { doc with Selection = Some newSel }

/// Move cursor to end of current block (End key)
let moveCursorToLineEnd doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        match findParagraph doc sel.Focus.BlockId with
        | None -> doc
        | Some para ->
            let maxOffset = TextBuffer.length para.Buffer
            let newSel = {
                Anchor = { sel.Focus with Offset = maxOffset }
                Focus = { sel.Focus with Offset = maxOffset }
            }
            { doc with Selection = Some newSel }


/// Typing at a block picture control belongs in a paragraph beside the control;
/// the control itself must continue to contain only its picture.
let private moveTypingOutsidePicture doc =
    match doc.Selection with
    | Some sel when sel.IsCollapsed ->
        match findParagraph doc sel.Focus.BlockId with
        | Some para ->
            let isPicture =
                para.Metadata.ContentControlXml
                |> Option.bind (fun xml -> XmlDsl.tryScan ("<root>" + xml + "</root>") (System.Collections.Generic.HashSet<string>()))
                |> Option.bind (XmlPatch.child "sdtPr")
                |> Option.bind (XmlPatch.child "picture")
                |> Option.isSome
            match isPicture, para.Metadata.ContentControlGroup with
            | true, Some group ->
                let adjacent =
                    { withoutParagraphEnd (withoutParagraphStart para.Metadata) with
                        Origin = None; PageBreakBefore = false; DropCap = None; MarkRevision = None
                        ContentControlId = None; ContentControlTag = None; ContentControlAlias = None
                        ContentControlXml = None; ContentControlGroup = None; ContentControlAncestors = [] }
                let metadata =
                    match List.rev para.Metadata.ContentControlAncestors with
                    | closest :: outer ->
                        { adjacent with ContentControlId = closest.Id; ContentControlTag = closest.Tag
                                        ContentControlAlias = closest.Alias; ContentControlXml = closest.Xml
                                        ContentControlGroup = Some closest.Group; ContentControlAncestors = List.rev outer }
                    | [] -> adjacent
                let newPara = { Id = nextBlockId(); Buffer = TextBuffer.empty; Metadata = metadata }
                let before = sel.Focus.Offset = 0
                let inGroup (p: ParagraphBlock) = p.Metadata.ContentControlGroup = Some group
                let place (content: ParagraphBlock list) =
                    let positions = content |> List.mapi (fun i p -> i, p) |> List.choose (fun (i, p) -> if inGroup p then Some i else None)
                    if positions.IsEmpty then None
                    else
                        let at = if before then List.head positions else List.last positions + 1
                        let paragraphs = content |> List.mapi (fun i p -> if i = at then [newPara; p] else [p]) |> List.concat
                        Some ((if at = content.Length then paragraphs @ [newPara] else paragraphs), at)
                let updated =
                    match findBlockIndex doc para.Id with
                    | Some _ ->
                        let positions =
                            doc.Blocks |> List.mapi (fun i block -> i, block) |> List.choose (fun (i, block) ->
                                match block with Paragraph p | Heading (_, p) when inGroup p -> Some i | _ -> None)
                        if positions.IsEmpty then None
                        else
                            let at = if before then List.head positions else List.last positions + 1
                            let blocks = doc.Blocks |> List.mapi (fun i block -> if i = at then [Paragraph newPara; block] else [block]) |> List.concat
                            Some { doc with Blocks = if at = doc.Blocks.Length then blocks @ [Paragraph newPara] else blocks }
                    | None ->
                        editContainerOf doc para.Id (fun content anchors ->
                            place content |> Option.map (fun (content, at) ->
                                content, anchors |> List.map (fun anchor -> if anchor >= at then anchor + 1 else anchor)))
                match updated with
                | Some changed ->
                    let caret = { BlockId = newPara.Id; Offset = 0 }
                    { changed with Selection = Some { Anchor = caret; Focus = caret } }
                | None -> doc
            | _ -> doc
        | None -> doc
    | _ -> doc

/// Insert text at cursor
let insertText text doc =
    // (Not into the mark of a comment: before it or after it.)
    let doc = snapSelectionToMarks None doc
    match doc.Selection with
    | None -> doc
    | Some sel ->
        // Delete selection if not collapsed
        let doc' = if sel.IsCollapsed then doc else deleteSelection doc
        let doc' = if text = "" then doc' else moveTypingOutsidePicture doc'

        // Insert text at cursor (which is now at the selection start after deletion)
        match doc'.Selection with
        | None -> doc'
        | Some sel' ->
            match findParagraph doc' sel'.Focus.BlockId with
            | None -> doc'
            | Some para ->
                let newBuffer = TextBuffer.insert sel'.Focus.Offset text para.Buffer
                let newPara = { para with Buffer = newBuffer }

                let newDoc = replaceParagraph doc' para.Id newPara

                // Move cursor forward
                let newSel = {
                    Anchor = { sel'.Focus with Offset = sel'.Focus.Offset + text.Length }
                    Focus = { sel'.Focus with Offset = sel'.Focus.Offset + text.Length }
                }
                { newDoc with Selection = Some newSel }

let insertField fieldCode result doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        let doc' = insertText result doc
        match doc'.Selection with
        | None -> doc'
        | Some sel' ->
            let startOffset = max 0 (sel'.Focus.Offset - result.Length)
            match findParagraph doc' sel'.Focus.BlockId with
            | None -> doc'
            | Some para ->
                let newBuffer =
                    TextBuffer.applyStyle startOffset result.Length (fun style ->
                        { style with FieldCode = Some fieldCode; FieldResult = Some result }) para.Buffer
                let newPara = { para with Buffer = newBuffer }
                replaceParagraph doc' para.Id newPara

let private escapeXml (text: string) =
    text
        .Replace("&", "&amp;")
        .Replace("<", "&lt;")
        .Replace(">", "&gt;")
        .Replace("\"", "&quot;")
        .Replace("'", "&apos;")

let insertEquation equationText doc =
    let text = "[Equation]"
    let equationXml = sprintf "<m:oMath><m:r><m:t>%s</m:t></m:r></m:oMath>" (escapeXml equationText)
    match doc.Selection with
    | None -> doc
    | Some sel ->
        let doc' = insertText text doc
        match doc'.Selection with
        | None -> doc'
        | Some sel' ->
            let startOffset = max 0 (sel'.Focus.Offset - text.Length)
            match findParagraph doc' sel'.Focus.BlockId with
            | None -> doc'
            | Some para ->
                let newBuffer =
                    TextBuffer.applyStyle startOffset text.Length (fun style ->
                        { style with EquationXml = Some equationXml; FontFamily = Some "Cambria Math" }) para.Buffer
                let newPara = { para with Buffer = newBuffer }
                replaceParagraph doc' para.Id newPara

/// Insert image at cursor position
let insertImageAtCursor src width height alt doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        // Delete selection if not collapsed
        let doc' = if sel.IsCollapsed then doc else deleteSelection doc

        // Insert image at cursor
        match doc'.Selection with
        | None -> doc'
        | Some sel' ->
            match findParagraph doc' sel'.Focus.BlockId with
            | None -> doc'
            | Some para ->
                let newBuffer = TextBuffer.insertImage sel'.Focus.Offset src width height alt para.Buffer
                let newPara = { para with Buffer = newBuffer }

                let newDoc = replaceParagraph doc' para.Id newPara

                // Move cursor forward by 1 (image occupies 1 character position)
                let newSel = {
                    Anchor = { sel'.Focus with Offset = sel'.Focus.Offset + 1 }
                    Focus = { sel'.Focus with Offset = sel'.Focus.Offset + 1 }
                }
                { newDoc with Selection = Some newSel }

/// Resize image at specified position
let resizeImage blockId offset newWidth newHeight doc =
    match findParagraph doc blockId with
    | None -> doc
    | Some para ->
        let newBuffer = TextBuffer.resizeImage offset newWidth newHeight para.Buffer
        let newPara = { para with Buffer = newBuffer }

        replaceParagraph doc para.Id newPara

let resizeImageNearCursor newWidth newHeight doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        match findParagraph doc sel.Focus.BlockId with
        | None -> doc
        | Some para ->
            let imageOffset =
                para.Buffer.Runs
                |> List.choose (function
                    | TextBuffer.Image img when img.Start <= sel.Focus.Offset -> Some img.Start
                    | _ -> None)
                |> List.sortDescending
                |> List.tryHead
            match imageOffset with
            | None -> doc
            | Some offset -> resizeImage para.Id offset newWidth newHeight doc

/// Move an image run to a new document position.
let moveImage blockId offset targetBlockId targetOffset doc =
    match findParagraph doc blockId with
    | None -> doc
    | Some sourcePara ->
        match sourcePara.Buffer.Runs |> List.tryPick (function TextBuffer.Image img when img.Start = offset -> Some img | _ -> None) with
        | None -> doc
        | Some img ->
            let sourceBuffer = TextBuffer.delete offset 1 sourcePara.Buffer
            let docWithoutImage = replaceParagraph doc sourcePara.Id { sourcePara with Buffer = sourceBuffer }
            match findParagraph docWithoutImage targetBlockId with
            | None -> docWithoutImage
            | Some targetPara ->
                let insertOffset =
                    if blockId = targetBlockId && targetOffset > offset then targetOffset - 1
                    else targetOffset
                let boundedOffset = max 0 (min (TextBuffer.length targetPara.Buffer) insertOffset)
                // The run moves whole: an opaque object keeps its XML, a picture its placement.
                let targetBuffer = TextBuffer.insertImageRun boundedOffset img targetPara.Buffer
                let nextDoc = replaceParagraph docWithoutImage targetPara.Id { targetPara with Buffer = targetBuffer }
                { nextDoc with Selection = Some { Anchor = { BlockId = targetBlockId; Offset = boundedOffset + 1 }; Focus = { BlockId = targetBlockId; Offset = boundedOffset + 1 } } }

/// Delete backward (backspace)
let deleteBackward doc =
    match doc.Selection with
    | None -> doc
    | Some sel when not sel.IsCollapsed ->
        // Delete selection if active
        deleteSelection doc
    | Some sel when sel.Focus.Offset = 0 ->
        match findParagraph doc sel.Focus.BlockId with
        | Some para when para.Metadata.ListStyle <> NoList ->
            // At the start of a list item, as in Word: its number or bullet goes
            // (the paragraph leaves the list, as Enter in an empty item takes it
            // out); the next Backspace joins it to the paragraph before.
            replaceParagraph doc para.Id { para with Metadata = { para.Metadata with ListStyle = NoList; ListLevel = 0 } }
        | _ ->
            // At start of block, merge with previous
            mergeWithPrevious doc sel.Focus.BlockId
    | Some _ when (markSelectedAtCaret true doc).IsSome -> deleteSelection (markSelectedAtCaret true doc).Value
    | Some sel ->
        match findParagraph doc sel.Focus.BlockId with
        | None -> doc
        | Some para ->
            let step = backwardStep para.Buffer sel.Focus.Offset
            let newBuffer = TextBuffer.delete (sel.Focus.Offset - step) step para.Buffer
            let newPara = { para with Buffer = newBuffer }

            let newDoc = replaceParagraph doc para.Id newPara

            // Move cursor back
            let newSel = {
                Anchor = { sel.Focus with Offset = sel.Focus.Offset - step }
                Focus = { sel.Focus with Offset = sel.Focus.Offset - step }
            }
            { newDoc with Selection = Some newSel }

/// Set cursor position
let setCursor blockId offset doc =
    { doc with
        Selection = Some {
            Anchor = { BlockId = blockId; Offset = offset }
            Focus = { BlockId = blockId; Offset = offset }
        } }

/// An empty cell like `cell`: for a row or a column inserted next to it. It
/// has the cell's properties - width, span, borders, shading, vertical
/// alignment, margins - and one empty paragraph with the properties of the
/// cell's first paragraph (style, alignment, indents, spacing and the size
/// of its mark), so that the new row or column looks like the one it is
/// inserted next to, in the editor and in the file. `span`: None keeps the
/// cell's span (a row), Some 1 is a cell of one column.
let private cellLike (span: int option) (cell: TableCell) : TableCell =
    let paragraph =
        match cell.Content with
        | first :: _ ->
            { Id = nextBlockId ()
              Buffer = TextBuffer.empty
              Metadata =
                { first.Metadata with
                    PageBreakBefore = false; PageBreakAfter = false; ColumnBreakAfter = false; RenderedPageBreakBefore = None
                    SectionVerticalAlign = None
                    SectionBreak = None; SectionColumns = None; SectionPageSetup = None; SectionPageNumberStart = None; SectionPageNumberFormat = None
                    HeaderReferences = []; FooterReferences = []
                    ContentControlId = None; ContentControlTag = None; ContentControlAlias = None
                    ContentControlXml = None; ContentControlGroup = None; ContentControlAncestors = []
                    Origin = likeOrigin first.Metadata.Origin } }
        | [] -> { Id = nextBlockId (); Buffer = TextBuffer.empty; Metadata = defaultMetadata }
    { cell with
        Content = [ paragraph ]
        NestedTables = []
        NestedTableAnchors = []
        EnclosingControls = []
        VerticalMerge = None
        ColSpan = (match span with Some columns -> columns | None -> cell.ColSpan)
        Origin = likeOrigin cell.Origin }

/// The grid column a cell of a row starts in.
let private gridColumnOfCell (row: TableRow) (cellIndex: int) : int =
    let mutable column = max 0 row.GridBefore
    row.Cells |> List.iteri (fun i cell -> if i < cellIndex then column <- column + max 1 cell.ColSpan)
    column

/// The table's grid (w:tblGrid) with a column like the one at `like` put in at `at`.
let private gridWithColumn (at: int) (like: int) (grid: TableWidth list) : TableWidth list =
    if grid.IsEmpty || like < 0 || like >= grid.Length then grid
    else
        let at = max 0 (min grid.Length at)
        List.truncate at grid @ [ grid.[like] ] @ List.skip at grid

/// The table's grid without the column at `at`.
let private gridWithoutColumn (at: int) (grid: TableWidth list) : TableWidth list =
    if at < 0 || at >= grid.Length then grid
    else List.truncate at grid @ List.skip (at + 1) grid

/// A row like `row`, empty: the same properties (height, whether it may
/// break, spacing, what stands before and after it in the grid), cell by
/// cell the same cells. It is not a header row.
let private rowLike (row: TableRow) : TableRow =
    { row with Cells = row.Cells |> List.map (cellLike None); IsHeader = false; Origin = likeOrigin row.Origin }

/// Insert a table row below the row containing the selection.
let insertTableRowAfterSelection doc =
    match currentTableCell doc with
    | None -> doc
    | Some (tableId, rowIndex, _) ->
        updateTable doc tableId (fun table ->
            match List.tryItem rowIndex table.Rows with
            | None -> table
            | Some row ->
                let newRow = rowLike row
                { table with
                    Rows =
                        table.Rows
                        |> List.mapi (fun i existing -> if i = rowIndex then [ existing; newRow ] else [ existing ])
                        |> List.concat })

/// A column put into a table next to the cell at `colIndex` of each row
/// (after it, or before it): each row gets a cell like its neighbour, and
/// the table's grid a column as wide as the neighbour's.
let private insertTableColumn (after: bool) doc =
    match currentTableCell doc with
    | None -> doc
    | Some (tableId, rowIndex, colIndex) ->
        updateTable doc tableId (fun table ->
            let grid =
                match List.tryItem rowIndex table.Rows with
                | Some row when colIndex < row.Cells.Length ->
                    let first = gridColumnOfCell row colIndex
                    let last = first + max 1 row.Cells.[colIndex].ColSpan - 1
                    if after then gridWithColumn (last + 1) last table.GridColumns else gridWithColumn first first table.GridColumns
                | _ -> table.GridColumns
            { table with
                GridColumns = grid
                Rows =
                    table.Rows
                    |> List.map (fun row ->
                        { row with
                            Cells =
                                row.Cells
                                |> List.mapi (fun i cell ->
                                    if i <> colIndex then [ cell ]
                                    elif after then [ cell; cellLike (Some 1) cell ]
                                    else [ cellLike (Some 1) cell; cell ])
                                |> List.concat }) })

let insertTableColumnAfterSelection doc = insertTableColumn true doc

/// Insert a table row above the row containing the selection.
let insertTableRowBeforeSelection doc =
    match currentTableCell doc with
    | None -> doc
    | Some (tableId, rowIndex, _) ->
        updateTable doc tableId (fun table ->
            match List.tryItem rowIndex table.Rows with
            | None -> table
            | Some row ->
                // A row like the one it is inserted next to: the same cells, empty.
                let newRow = rowLike row
                { table with
                    Rows =
                        table.Rows
                        |> List.mapi (fun i existing -> if i = rowIndex then [ newRow; existing ] else [ existing ])
                        |> List.concat })

/// Insert a table column to the left of the column containing the selection.
let insertTableColumnBeforeSelection doc = insertTableColumn false doc

/// The table cells the selection spans: (table, first row, first cell, last
/// row, last cell) when both ends of the selection are in cells of one table.
let private selectedTableCells doc =
    match doc.Selection with
    | None -> None
    | Some selection ->
        match tryFindTableCellByParagraph doc selection.Anchor.BlockId, tryFindTableCellByParagraph doc selection.Focus.BlockId with
        | Some (anchorTable, anchorRow, anchorCol), Some (focusTable, focusRow, focusCol) when anchorTable = focusTable ->
            Some (anchorTable, min anchorRow focusRow, min anchorCol focusCol, max anchorRow focusRow, max anchorCol focusCol)
        | _ -> None

/// Tab / Shift+Tab in a table: the caret goes to the next / previous cell
/// (row by row), and Tab in the last cell adds a row and goes to its first
/// cell, as in Word. None when the caret is not in a table (or, for
/// Shift+Tab, is in the first cell).
let moveToAdjacentTableCell (forward: bool) (doc: Document) : Document option =
    match currentTableCell doc with
    | None -> None
    | Some (tableId, rowIndex, colIndex) ->
        let tableOf (current: Document) = tryFindTable current tableId
        let cursorIn (current: Document) (row: int) (col: int) =
            tableOf current
            |> Option.bind (fun table -> List.tryItem row table.Rows)
            |> Option.bind (fun tableRow -> List.tryItem col tableRow.Cells)
            |> Option.bind (fun cell -> List.tryHead cell.Content)
            |> Option.map (fun para -> setCursor para.Id 0 current)
        match tableOf doc with
        | None -> None
        | Some table ->
            let cellsIn (row: int) = match List.tryItem row table.Rows with Some tableRow -> tableRow.Cells.Length | None -> 0
            // (Cells that continue a vertical merge hold nothing to type into: passed over.)
            let continues (row: int) (col: int) =
                match List.tryItem row table.Rows |> Option.bind (fun tableRow -> List.tryItem col tableRow.Cells) with
                | Some cell -> cell.VerticalMerge = Some ContinueMerge
                | None -> false
            let rec step (row: int) (col: int) =
                let nextRow, nextCol =
                    if forward then (if col + 1 < cellsIn row then row, col + 1 else row + 1, 0)
                    else (if col > 0 then row, col - 1 else row - 1, max 0 (cellsIn (row - 1) - 1))
                if nextRow < 0 then None
                elif nextRow >= table.Rows.Length then
                    // Past the last cell: a new row like the last one.
                    let extended = insertTableRowAfterSelection (match cursorIn doc (table.Rows.Length - 1) 0 with Some at -> at | None -> doc)
                    cursorIn extended table.Rows.Length 0
                elif continues nextRow nextCol then step nextRow nextCol
                else cursorIn doc nextRow nextCol
            step rowIndex colIndex

/// Whether the selection spans more than one cell of a table (so that
/// `mergeSelectedTableCells` has something to merge).
let canMergeSelectedTableCells doc =
    match selectedTableCells doc with
    | Some (_, firstRow, firstCol, lastRow, lastCol) -> firstRow <> lastRow || firstCol <> lastCol
    | None -> false

/// Merge the selected cells of a table into one: the cells of each selected
/// row become one cell spanning their columns (holding the content of all of
/// them), and the rows below the first are merged into it vertically.
let mergeSelectedTableCells doc =
    match selectedTableCells doc with
    | Some (tableId, firstRow, firstCol, lastRow, lastCol) when firstRow <> lastRow || firstCol <> lastCol ->
        let isEmptyParagraph (para: ParagraphBlock) = TextBuffer.length para.Buffer = 0
        let merged =
            updateTable doc tableId (fun table ->
                // Everything the merged cells hold ends up in the first one.
                let content =
                    [ for rowIndex in firstRow .. lastRow do
                        match List.tryItem rowIndex table.Rows with
                        | Some row ->
                            for colIndex in firstCol .. min lastCol (row.Cells.Length - 1) do
                                yield! row.Cells.[colIndex].Content
                        | None -> () ]
                let kept =
                    match content |> List.filter (fun para -> not (isEmptyParagraph para)) with
                    | [] -> List.truncate 1 content
                    | nonEmpty -> nonEmpty
                { table with
                    Rows =
                        table.Rows
                        |> List.mapi (fun rowIndex row ->
                            if rowIndex < firstRow || rowIndex > lastRow || firstCol >= row.Cells.Length then row
                            else
                                let last = min lastCol (row.Cells.Length - 1)
                                let span = row.Cells |> List.skip firstCol |> List.truncate (last - firstCol + 1)
                                let first = List.head span
                                let width =
                                    if span |> List.forall (fun cell -> cell.Width.IsSome) then Some (span |> List.sumBy (fun cell -> cell.Width.Value)) else None
                                let mergedCell =
                                    { first with
                                        ColSpan = span |> List.sumBy (fun cell -> max 1 cell.ColSpan)
                                        Width = width
                                        Content = (if rowIndex = firstRow then kept else [ { Id = nextBlockId (); Buffer = TextBuffer.empty; Metadata = defaultMetadata } ])
                                        NestedTables = (if rowIndex = firstRow then span |> List.collect (fun cell -> cell.NestedTables) else [])
                                        NestedTableAnchors = (if rowIndex = firstRow then first.NestedTableAnchors else [])
                                        VerticalMerge =
                                            if firstRow = lastRow then first.VerticalMerge
                                            elif rowIndex = firstRow then Some RestartMerge
                                            else Some ContinueMerge }
                                { row with Cells = List.truncate firstCol row.Cells @ [ mergedCell ] @ List.skip (last + 1) row.Cells }) })
        match tryGetTableCellParagraph merged tableId firstRow firstCol with
        | Some para -> setCursor para.Id 0 merged
        | None -> merged
    | _ -> doc

/// Split the cell the caret is in back into single cells: one per column it
/// spans, and - for a vertically merged cell - one per row.
let splitSelectedTableCell doc =
    match currentTableCell doc with
    | None -> doc
    | Some (tableId, rowIndex, colIndex) ->
        updateTable doc tableId (fun table ->
            match List.tryItem rowIndex table.Rows |> Option.bind (fun row -> List.tryItem colIndex row.Cells) with
            | None -> table
            | Some cell ->
                let span = max 1 cell.ColSpan
                // The rows the cell is merged over: this one and the ones continuing it.
                let lastRow =
                    if cell.VerticalMerge <> Some RestartMerge then rowIndex
                    else
                        let mutable last = rowIndex
                        let mutable continues = true
                        while continues && last + 1 < table.Rows.Length do
                            match List.tryItem colIndex table.Rows.[last + 1].Cells with
                            | Some below when below.VerticalMerge = Some ContinueMerge -> last <- last + 1
                            | _ -> continues <- false
                        last
                if span = 1 && lastRow = rowIndex then table
                else
                    { table with
                        Rows =
                            table.Rows
                            |> List.mapi (fun index row ->
                                if index < rowIndex || index > lastRow then row
                                else
                                    match List.tryItem colIndex row.Cells with
                                    | None -> row
                                    | Some target ->
                                        let width = target.Width |> Option.map (fun total -> total / float (max 1 target.ColSpan))
                                        let first = { target with ColSpan = 1; Width = width; VerticalMerge = None }
                                        let rest = [ for _ in 2 .. max 1 target.ColSpan -> { emptyTableCell () with Width = width; BackgroundColor = target.BackgroundColor; Shading = target.Shading; Borders = target.Borders } ]
                                        { row with Cells = List.truncate colIndex row.Cells @ (first :: rest) @ List.skip (colIndex + 1) row.Cells }) })

let deleteSelectedTableRow doc =
    match currentTableCell doc with
    | None -> doc
    | Some (tableId, rowIndex, colIndex) ->
        let updatedDoc = updateTable doc tableId (fun table ->
            if table.Rows.Length <= 1 then table
            else { table with Rows = table.Rows |> List.mapi (fun i row -> i, row) |> List.choose (fun (i, row) -> if i = rowIndex then None else Some row) })

        match tryFindTable doc tableId |> Option.map (fun table -> table.Rows.Length) with
        | Some oldLength when oldLength > 1 ->
            let targetRow = min rowIndex (oldLength - 2)
            match tryGetTableCellParagraph updatedDoc tableId targetRow colIndex with
            | Some para -> setCursor para.Id 0 updatedDoc
            | None -> updatedDoc
        | _ -> updatedDoc

let deleteSelectedTableColumn doc =
    match currentTableCell doc with
    | None -> doc
    | Some (tableId, _, colIndex) ->
        let updatedDoc = updateTable doc tableId (fun table ->
            let maxColumns = table.Rows |> List.map (fun row -> row.Cells.Length) |> List.fold max 0
            if maxColumns <= 1 then table
            else
                // (The grid loses the column the cell of the selection's row is in, when that cell is one column wide.)
                let grid =
                    match currentTableCell doc |> Option.bind (fun (_, rowIndex, _) -> List.tryItem rowIndex table.Rows) with
                    | Some row when colIndex < row.Cells.Length && row.Cells.[colIndex].ColSpan <= 1 && table.GridColumns.Length > 1 ->
                        gridWithoutColumn (gridColumnOfCell row colIndex) table.GridColumns
                    | _ -> table.GridColumns
                { table with
                    GridColumns = grid
                    Rows =
                        table.Rows
                        |> List.map (fun row ->
                            { row with
                                Cells =
                                    row.Cells
                                    |> List.mapi (fun i cell -> i, cell)
                                    |> List.choose (fun (i, cell) -> if i = colIndex then None else Some cell) }) })

        match currentTableCell updatedDoc with
        | Some _ -> updatedDoc
        | None ->
            match currentTableCell doc with
            | Some (_, rowIndex, _) ->
                let targetCol = max 0 (colIndex - 1)
                match tryGetTableCellParagraph updatedDoc tableId rowIndex targetCol with
                | Some para -> setCursor para.Id 0 updatedDoc
                | None -> updatedDoc
            | None -> updatedDoc

/// A new table of empty cells as insertTable makes one (thin borders, the
/// document's table style for new tables), with ids of its own.
let private newTable rows columns (doc: Document) : TableBlock =
    let rowCount = max 1 rows
    let columnCount = max 1 columns
    // (One layout pixel - three quarters of a point: a thinner line is
    // painted as a faint grey one at 100%.)
    let thin : TableBorder option = Some { Style = Some "single"; Color = Some "#000000"; Width = Some 1.0 }
    // The table refers to the document's table style for new tables: its
    // "Table Grid" when it has one, else its default table style.
    let defaultTableStyleId =
        let tableStyles = doc.Styles |> List.filter (fun style -> style.StyleType = "table")
        tableStyles
        |> List.tryFind (fun style -> style.Id = "TableGrid")
        |> Option.orElse (tableStyles |> List.tryFind (fun style -> style.IsDefault))
        |> Option.map (fun style -> style.Id)
    let table = {
        Id = nextBlockId ()
        Rows = [ for _ in 1 .. rowCount -> { Cells = [ for _ in 1 .. columnCount -> emptyTableCell () ]; IsHeader = false; Height = None; HeightRule = None; CanSplit = None; CellSpacing = None; PaddingTop = None; PaddingRight = None; PaddingBottom = None; PaddingLeft = None; GridBefore = 0; GridAfter = 0; WidthBefore = None; WidthAfter = None; BordersOverride = None; Origin = None } ]
        PreferredWidth = None
        Alignment = None
        Indent = None
        LayoutMode = None
        GridColumns = []
        // A table that was just inserted is seen: single thin borders around
        // every cell (what Word's "Table Grid" gives), stated on the table so
        // that they are there whatever the document's table styles say.
        Borders = Some { Top = thin; Right = thin; Bottom = thin; Left = thin; InsideH = thin; InsideV = thin }
        BackgroundColor = None
        Shading = None
        CellSpacing = None
        PaddingTop = None
        PaddingRight = None
        PaddingBottom = None
        PaddingLeft = None
        Floating = false
        FloatingXml = None
        StyleId = defaultTableStyleId
        ContentControlId = None
        ContentControlTag = None
        ContentControlAlias = None
        ContentControlXml = None
        ContentControlGroup = None
        ContentControlAncestors = []
        Look = None
    }
    table

let insertTable rows columns doc =
    let table = newTable rows columns doc
    let firstCellPara = table.Rows.Head.Cells.Head.Content.Head
    let tableBlock = Table table
    let blocks =
        match doc.Selection with
        | Some sel ->
            let insertAfterIndex =
                doc.Blocks
                |> List.tryFindIndex (function
                    | Paragraph p -> p.Id = sel.Focus.BlockId
                    | Heading (_, p) -> p.Id = sel.Focus.BlockId
                    | Table t -> t.Id = sel.Focus.BlockId
                    | TextBox t -> t.Id = sel.Focus.BlockId)
            match insertAfterIndex with
            | Some index ->
                let before = doc.Blocks |> List.take (index + 1)
                let after = doc.Blocks |> List.skip (index + 1)
                before @ [tableBlock] @ after
            | None -> doc.Blocks @ [tableBlock]
        | None -> doc.Blocks @ [tableBlock]
    { doc with Blocks = blocks } |> setCursor firstCellPara.Id 0

let private insertTopLevelBlockAfterSelection block doc =
    let blocks =
        match doc.Selection with
        | Some sel ->
            let insertAfterIndex =
                doc.Blocks
                |> List.tryFindIndex (function
                    | Paragraph p -> p.Id = sel.Focus.BlockId
                    | Heading (_, p) -> p.Id = sel.Focus.BlockId
                    | Table t -> t.Id = sel.Focus.BlockId
                    | TextBox t -> t.Id = sel.Focus.BlockId)
            match insertAfterIndex with
            | Some index ->
                let before = doc.Blocks |> List.take (index + 1)
                let after = doc.Blocks |> List.skip (index + 1)
                before @ [block] @ after
            | None -> doc.Blocks @ [block]
        | None -> doc.Blocks @ [block]
    { doc with Blocks = blocks }

/// A paragraph without text, with an id of its own.
let newEmptyParagraph () : ParagraphBlock =
    { Id = nextBlockId (); Buffer = TextBuffer.empty; Metadata = defaultMetadata }

/// A text box anchored in the block the caret is in, a little below and
/// right of where that block starts, in front of the text. (It stood on the
/// first page whatever block the caret was in.)
/// A shape (text box) of the body, of a header or of a footer with another place and size:
/// `change` makes it of the shape it is. None: the document has no shape of that id.
let private withTextBox (id: BlockId) (change: TextBoxBlock -> TextBoxBlock) (doc: Document) : Document option =
    let mutable found = false
    let mutable any = false
    // (The list it was when the shape is not in it.)
    let mapBlocks (blocks: Block list) : Block list =
        if not (blocks |> List.exists (function TextBox box -> box.Id = id | _ -> false)) then blocks
        else
            found <- true
            blocks
            |> List.map (fun block ->
                match block with
                | TextBox box when box.Id = id ->
                    let changed = change box
                    if changed = box then block
                    else
                        any <- true
                        TextBox changed
                | _ -> block)
    let mapParts (parts: HeaderFooterPart list) : HeaderFooterPart list =
        parts
        |> List.map (fun part ->
            if found then part
            else
                let blocks = mapBlocks part.Blocks
                if obj.ReferenceEquals(blocks, part.Blocks) then part else { part with Blocks = blocks })
    let blocks = mapBlocks doc.Blocks
    // (Nothing of it changes: the document it was.)
    if found then Some (if any then { doc with Blocks = blocks } else doc)
    else
        let headers = mapParts doc.HeaderParts
        let footers = mapParts doc.FooterParts
        if not found then None
        elif any then Some { doc with HeaderParts = headers; FooterParts = footers }
        else Some doc

/// A shape of the body where `place` says: for each direction the frame its offset is from
/// (None: the one it has) and the offset, in layout pixels. A direction that is given is by
/// its offset - not aligned any more. (Not a shape that stands in a line of text.)
let moveTextBox (id: BlockId) (horizontal: (string option * float) option) (vertical: (string option * float) option) (doc: Document) : Document option =
    doc
    |> withTextBox id (fun box ->
        if box.Inline then box
        else
            let box =
                match horizontal with
                | Some (frame, offset) -> { box with X = Some offset; PositionHAlign = None; PositionHRelativeFrom = frame |> Option.orElse box.PositionHRelativeFrom }
                | None -> box
            match vertical with
            | Some (frame, offset) -> { box with Y = Some offset; PositionVAlign = None; PositionVRelativeFrom = frame |> Option.orElse box.PositionVRelativeFrom }
            | None -> box)

/// A shape of the body as wide and as high as given (layout pixels, at least one). A shape
/// that was as high as its text (a:spAutoFit) is of the given height from then on.
let resizeTextBox (id: BlockId) (width: float) (height: float) (doc: Document) : Document option =
    doc |> withTextBox id (fun box -> { box with Width = max 1.0 width; Height = max 1.0 height; AutoFit = box.AutoFit && abs (height - box.Height) < 0.01 })

/// A shape at another place and - with `size` (width, height) - of another size, in one
/// change: what a drag of a handle of its left or top edge does.
let placeTextBox (id: BlockId) (horizontal: (string option * float) option) (vertical: (string option * float) option) (size: (float * float) option) (doc: Document) : Document option =
    match moveTextBox id horizontal vertical doc with
    | Some moved ->
        match size with
        | Some (width, height) -> resizeTextBox id width height moved
        | None -> Some moved
    | None -> None

/// The rectangle (left, top, width, height) of a shape after one of its handles - "nw", "n",
/// "ne", "e", "se", "s", "sw", "w" - was dragged by (dx, dy): the side (or corner) across from
/// the handle stays where it is, also for a shape that is turned by `rotation` degrees
/// clockwise about its centre (the rectangle is the one before the turn, and the drag is
/// taken in the shape's own directions). At least four pixels wide and high.
let resizedShapeRect (rotation: float) (x: float) (y: float) (width: float) (height: float) (handle: string) (dx: float) (dy: float) : float * float * float * float =
    let angle = rotation * System.Math.PI / 180.0
    let cos, sin = System.Math.Cos angle, System.Math.Sin angle
    let localX, localY = dx * cos + dy * sin, dy * cos - dx * sin
    let least = 4.0
    let west, east = handle.Contains "w", handle.Contains "e"
    let north, south = handle.Contains "n", handle.Contains "s"
    let newWidth = if east then max least (width + localX) elif west then max least (width - localX) else width
    let newHeight = if south then max least (height + localY) elif north then max least (height - localY) else height
    // (How far the centre goes, in the shape's own directions.)
    let shiftX = if east then (newWidth - width) / 2.0 elif west then (width - newWidth) / 2.0 else 0.0
    let shiftY = if south then (newHeight - height) / 2.0 elif north then (height - newHeight) / 2.0 else 0.0
    let centerX = x + width / 2.0 + shiftX * cos - shiftY * sin
    let centerY = y + height / 2.0 + shiftX * sin + shiftY * cos
    centerX - newWidth / 2.0, centerY - newHeight / 2.0, newWidth, newHeight

/// The body without the shape of that id (its text goes with it). The caret, when it was in
/// the shape's text, goes to the block the shape was anchored in.
let deleteTextBox (id: BlockId) (doc: Document) : Document option =
    match doc.Blocks |> List.tryPick (function TextBox box when box.Id = id -> Some box | _ -> None) with
    | None ->
        // A shape of a header or a footer: out of its part.
        let holds (part: HeaderFooterPart) = part.Blocks |> List.exists (function TextBox box -> box.Id = id | _ -> false)
        let without (parts: HeaderFooterPart list) =
            parts |> List.map (fun part -> if holds part then { part with Blocks = part.Blocks |> List.filter (function TextBox box -> box.Id <> id | _ -> true) } else part)
        if List.exists holds doc.HeaderParts then Some { doc with HeaderParts = without doc.HeaderParts }
        elif List.exists holds doc.FooterParts then Some { doc with FooterParts = without doc.FooterParts }
        else None
    | Some box ->
        let blocks = doc.Blocks |> List.filter (function TextBox other -> other.Id <> id | _ -> true)
        let inBox (position: Position) = box.Content |> List.exists (fun para -> para.Id = position.BlockId)
        let selection =
            match doc.Selection with
            | Some sel when inBox sel.Anchor || inBox sel.Focus ->
                let target =
                    box.AnchorBlockId
                    |> Option.bind (fun anchor -> paragraphsOfBlocks blocks |> List.tryFind (fun para -> para.Id = anchor))
                    |> Option.orElse (paragraphsOfBlocks blocks |> List.tryHead)
                target |> Option.map (fun para -> { Anchor = { BlockId = para.Id; Offset = 0 }; Focus = { BlockId = para.Id; Offset = 0 } })
            | other -> other
        Some { doc with Blocks = blocks; Selection = selection }

let insertTextBox text doc =
    let para = {
        Id = nextBlockId ()
        Buffer = TextBuffer.ofString text
        Metadata = defaultMetadata
    }
    let anchor =
        match doc.Selection with
        | Some selection ->
            doc.Blocks
            |> List.tryPick (fun block ->
                match block with
                | TextBox _ -> None
                | Paragraph p
                | Heading (_, p) -> if p.Id = selection.Focus.BlockId then Some p.Id else None
                | Table table -> if blockHolds selection.Focus.BlockId block then Some table.Id else None)
        | None -> None
    let textBox = TextBox {
        defaultTextBox with
            Id = nextBlockId ()
            Content = [para]
            X = Some (if anchor.IsSome then 24.0 else 96.0)
            Y = Some (if anchor.IsSome then 24.0 else 96.0)
            Width = 240.0
            Height = 96.0
            AnchorBlockId = anchor
            PositionHRelativeFrom = (if anchor.IsSome then Some "column" else None)
            PositionVRelativeFrom = (if anchor.IsSome then Some "paragraph" else None)
            Wrap = Some "none"
    }
    insertTopLevelBlockAfterSelection textBox doc

let private textRunsInParagraph para =
    para.Buffer.Runs
    |> List.choose (function
        | TextBuffer.Text run -> Some run
        | TextBuffer.Image _ -> None)

let private textRunsInBlock block =
    match block with
    | Paragraph para
    | Heading (_, para) -> textRunsInParagraph para
    | Table table -> tableParagraphsInOrder table |> List.collect textRunsInParagraph
    | TextBox textBox ->
        textBox.Content |> List.collect textRunsInParagraph

let private nextSequenceResult label doc =
    let prefix = sprintf " SEQ %s " label
    let maxExisting =
        doc.Blocks
        |> List.collect textRunsInBlock
        |> List.choose (fun run ->
            match run.Style.FieldCode, run.Style.FieldResult with
            | Some code, Some result when code.StartsWith(prefix) ->
                match System.Int32.TryParse(result) with
                | true, value -> Some value
                | false, _ -> None
            | _ -> None)
        |> function
            | [] -> 0
            | values -> List.max values
    string (maxExisting + 1)

let private nextBookmarkId doc =
    let maxExisting =
        doc.Blocks
        |> List.collect textRunsInBlock
        // (A run has several: every id among them, of a start, an end or one opened after it.)
        |> List.collect (fun run -> TextBuffer.bookmarkIds run.Style)
        |> List.choose (fun value ->
            match System.Int32.TryParse(value) with
            | true, parsed -> Some parsed
            | false, _ -> None)
        |> function
            | [] -> 0
            | values -> List.max values
    string (maxExisting + 1)

let private nextRevisionId doc =
    let maxExisting =
        doc.Blocks
        |> List.collect textRunsInBlock
        |> List.choose (fun run -> run.Style.RevisionId)
        |> List.choose (fun value ->
            match System.Int32.TryParse(value) with
            | true, parsed -> Some parsed
            | false, _ -> None)
        |> function
            | [] -> 0
            | values -> List.max values
    string (maxExisting + 1)

let insertCaption label captionText doc =
    let result = nextSequenceResult label doc
    let bookmarkId = nextBookmarkId doc
    let bookmarkName = sprintf "_Ref%s" bookmarkId
    let prefix = label + " "
    let suffix = ": " + captionText
    let fullText = prefix + result + suffix
    let buffer =
        TextBuffer.ofString fullText
        |> TextBuffer.applyStyle prefix.Length result.Length (fun style ->
            { style with
                BookmarkStarts = [ { Id = bookmarkId; Name = bookmarkName; EndsHere = false } ]
                BookmarkEnds = [ TextBuffer.Ends bookmarkId ]
                FieldCode = Some (sprintf " SEQ %s \\* ARABIC " label)
                FieldResult = Some result })
    let para = Paragraph {
        Id = nextBlockId ()
        Buffer = buffer
        Metadata = { defaultMetadata with Style = Some "Caption" }
    }
    insertTopLevelBlockAfterSelection para doc

let insertCrossReferenceToLatestCaption label doc =
    let prefix = sprintf " SEQ %s " label
    let latest =
        doc.Blocks
        |> List.collect textRunsInBlock
        |> List.choose (fun run ->
            match run.Style.FieldCode, run.Style.FieldResult, TextBuffer.bookmarkName run.Style with
            | Some code, Some result, Some bookmarkName when code.StartsWith(prefix) ->
                match System.Int32.TryParse(result) with
                | true, number -> Some (number, result, bookmarkName)
                | false, _ -> None
            | _ -> None)
        |> List.sortBy (fun (number, _, _) -> number)
        |> List.tryLast
    match latest with
    | None -> doc
    | Some (_, result, bookmarkName) ->
        insertField (sprintf " REF %s \\h " bookmarkName) result doc

let private fieldCodeParts (code: string) =
    code.Trim().Split([| ' '; '\t'; '\r'; '\n' |], System.StringSplitOptions.RemoveEmptyEntries)
    |> Array.toList

let private dateFieldFormat (code: string) =
    let formatSwitch = code.IndexOf("\\@")
    if formatSwitch < 0 then "yyyy-MM-dd"
    else
        let firstQuote = code.IndexOf('"', formatSwitch)
        if firstQuote < 0 then "yyyy-MM-dd"
        else
            let secondQuote = code.IndexOf('"', firstQuote + 1)
            if secondQuote <= firstQuote then "yyyy-MM-dd"
            else code.Substring(firstQuote + 1, secondQuote - firstQuote - 1)

let private fieldSwitchValue (switchName: string) (code: string) =
    let switchIndex = code.IndexOf(switchName)
    if switchIndex < 0 then None
    else
        let firstQuote = code.IndexOf('"', switchIndex)
        if firstQuote < 0 then None
        else
            let secondQuote = code.IndexOf('"', firstQuote + 1)
            if secondQuote <= firstQuote then None
            else Some (code.Substring(firstQuote + 1, secondQuote - firstQuote - 1))

let private timeFieldFormat (code: string) =
    // TIME shares DATE's "\@ \"fmt\"" switch; default to a clock format when absent.
    if code.IndexOf("\\@") < 0 then "h:mm tt" else dateFieldFormat code

let private normalizeStyleKey (s: string) =
    s.Replace(" ", "").Replace("-", "").ToLowerInvariant()

/// Extract the style name a STYLEREF field points at: the first quoted string,
/// else the first non-switch token after the field name (e.g. STYLEREF Heading1,
/// STYLEREF "Heading 1", STYLEREF 1).
let private styleRefStyleName (code: string) =
    let q1 = code.IndexOf('"')
    if q1 >= 0 then
        let q2 = code.IndexOf('"', q1 + 1)
        if q2 > q1 then Some (code.Substring(q1 + 1, q2 - q1 - 1)) else None
    else
        code.Trim().Split([| ' '; '\t'; '\r'; '\n' |], System.StringSplitOptions.RemoveEmptyEntries)
        |> Array.toList
        |> function
           | _ :: rest -> rest |> List.tryFind (fun t -> not (t.StartsWith("\\")))
           | [] -> None

let private paragraphText para = Rope.toString para.Buffer.Rope

let private tocHeadingRange (code: string) =
    let parseBound fallback (text: string) =
        match System.Int32.TryParse(text: string) with
        | true, value -> max 1 (min 9 value)
        | false, _ -> fallback

    match fieldSwitchValue "\\o" code with
    | None -> 1, 3
    | Some value ->
        match value.Split('-') with
        | parts when parts.Length = 2 ->
            let first = parseBound 1 parts[0]
            let last = parseBound 3 parts[1]
            min first last, max first last
        | _ -> 1, 3

let private tocHeadingLines code pageMap doc =
    let minLevel, maxLevel = tocHeadingRange code
    doc.Blocks
    |> List.choose (function
        | Heading (level, para) when level >= minLevel && level <= maxLevel ->
            let pageNumber =
                pageMap.BlockAnchors
                |> List.tryFind (fun anchor -> anchor.PageBlockId = para.Id)
                |> Option.map (fun anchor -> anchor.PageNumber)
                |> Option.defaultValue 1
            Some (String.replicate (max 0 (level - 1)) "  " + paragraphText para + "\t" + string pageNumber)
        | _ -> None)
    |> String.concat "\n"

let private tocCaptionLines label doc =
    let hasSequenceField para =
        para.Buffer.Runs
        |> List.exists (function
            | Text run ->
                match run.Style.FieldCode with
                | Some code -> (fieldCodeParts code) |> List.exists ((=) label)
                | None -> false
            | Image _ -> false)

    doc.Blocks
    |> List.choose (function
        | Paragraph para
        | Heading (_, para) when hasSequenceField para -> Some (paragraphText para)
        | _ -> None)
    |> String.concat "\n"

/// Carry a caret offset through a set of text replacements in its paragraph.
/// Edits are (original start, original end, new length), sorted and
/// non-overlapping. A caret inside a replaced range lands after the
/// replacement.
let private offsetAfterEdits (edits: ResizeArray<struct (int * int * int)>) offset =
    let mutable delta = 0
    let mutable result = -1
    let mutable index = 0
    while result < 0 && index < edits.Count do
        let struct (editStart, editEnd, newLength) = edits.[index]
        if offset <= editStart then
            result <- offset + delta
        elif offset < editEnd then
            result <- editStart + delta + newLength
        else
            delta <- delta + newLength - (editEnd - editStart)
        index <- index + 1
    if result < 0 then offset + delta else result

/// Re-point the selection after paragraphs were rewritten in place, so it
/// never ends up past the end of a buffer.
let private selectionAfterEdits (edits: Dictionary<BlockId, ResizeArray<struct (int * int * int)>>) (doc: Document) =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        let adjust (position: Position) =
            let moved =
                match edits.TryGetValue position.BlockId with
                | true, paragraphEdits -> offsetAfterEdits paragraphEdits position.Offset
                | false, _ -> position.Offset
            let bounded =
                match findParagraph doc position.BlockId with
                | Some para -> max 0 (min (TextBuffer.length para.Buffer) moved)
                | None -> moved
            if bounded = position.Offset then position else { position with Offset = bounded }
        let anchor = adjust sel.Anchor
        let focus = adjust sel.Focus
        if sameReference anchor sel.Anchor && sameReference focus sel.Focus then doc
        else { doc with Selection = Some { Anchor = anchor; Focus = focus } }

let private collectSequenceRefreshTargets doc =
    let mutable counters: Map<string, int> = Map.empty
    let mutable sequences: Map<string, string list> = Map.empty
    let mutable bookmarks: Map<string, string> = Map.empty

    let addSequence label (bookmarkNames: string list) =
        let next = (counters |> Map.tryFind label |> Option.defaultValue 0) + 1
        counters <- counters |> Map.add label next
        let result = string next
        sequences <- sequences |> Map.add label ((sequences |> Map.tryFind label |> Option.defaultValue []) @ [result])
        for name in bookmarkNames do bookmarks <- bookmarks |> Map.add name result

    editableParagraphsInOrder doc
    |> List.iter (fun para ->
        para.Buffer.Runs
        |> List.iter (function
            | Text run ->
                match run.Style.FieldCode |> Option.map fieldCodeParts with
                | Some ("SEQ" :: label :: _) -> addSequence label (run.Style.BookmarkStarts |> List.map (fun bookmark -> bookmark.Name) |> List.filter (fun name -> name <> ""))
                | _ -> ()
            | Image _ -> ()))
    sequences, bookmarks

let private romanNumeral (number: int) =
    if number <= 0 || number >= 4000 then string number
    else
        let values = [| 1000; 900; 500; 400; 100; 90; 50; 40; 10; 9; 5; 4; 1 |]
        let letters = [| "M"; "CM"; "D"; "CD"; "C"; "XC"; "L"; "XL"; "X"; "IX"; "V"; "IV"; "I" |]
        let text = System.Text.StringBuilder()
        let mutable rest = number
        for i in 0 .. values.Length - 1 do
            while rest >= values.[i] do
                text.Append(letters.[i]) |> ignore
                rest <- rest - values.[i]
        text.ToString()

/// A, B, ... Z, AA, BB, ... as Word numbers pages by letters.
let private letterNumeral (number: int) =
    if number <= 0 then string number
    else
        let letter = string (char (int 'A' + (number - 1) % 26))
        String.replicate ((number - 1) / 26 + 1) letter

/// A page number in a number format of w:pgNumType w:fmt ("decimal",
/// "upperRoman", "lowerRoman", "upperLetter", "lowerLetter", "numberInDash";
/// any other as decimal).
let formatPageNumber (format: string) (number: int) : string =
    match format with
    | "upperRoman" -> romanNumeral number
    | "lowerRoman" -> (romanNumeral number).ToLowerInvariant()
    | "upperLetter" -> letterNumeral number
    | "lowerLetter" -> (letterNumeral number).ToLowerInvariant()
    | "numberInDash" -> "- " + string number + " -"
    | _ -> string number

/// The number format a field states itself (the switch \* roman, \* ROMAN,
/// \* alphabetic, \* ALPHABETIC, \* Arabic), if it does.
let fieldNumberFormat (code: string) : string option =
    let mutable found = None
    let mutable at = code.IndexOf "\\*"
    while found.IsNone && at >= 0 do
        let mutable i = at + 2
        while i < code.Length && System.Char.IsWhiteSpace code.[i] do i <- i + 1
        let start = i
        while i < code.Length && not (System.Char.IsWhiteSpace code.[i]) && code.[i] <> '\\' do i <- i + 1
        let name = code.Substring(start, i - start)
        found <-
            match name with
            | "roman" -> Some "lowerRoman"
            | "ROMAN" | "Roman" -> Some "upperRoman"
            | "alphabetic" -> Some "lowerLetter"
            | "ALPHABETIC" | "Alphabetic" -> Some "upperLetter"
            | "Arabic" | "arabic" | "ARABIC" -> Some "decimal"
            | _ -> None
        at <- code.IndexOf("\\*", i)
    found

/// The page number format of the section each paragraph is in: what the
/// section's properties state (the paragraph that ends the section, the page
/// setup for the last one), "decimal" where they state none.
let sectionNumberFormats (doc: Document) : System.Collections.Generic.Dictionary<BlockId, string> =
    let formats = System.Collections.Generic.Dictionary<BlockId, string>()
    let mutable current = doc.PageSetup.PageNumberFormat |> Option.defaultValue "decimal"
    for block in List.rev doc.Blocks do
        // (The paragraph that carries the break is the last of its section.)
        match block with
        | Paragraph para
        | Heading (_, para) when para.Metadata.SectionBreak.IsSome ->
            current <- para.Metadata.SectionPageNumberFormat |> Option.defaultValue "decimal"
        | _ -> ()
        for para in paragraphsOfBlocks [ block ] do
            formats.[para.Id] <- current
    formats

let refreshFieldsWithPageMap pageMap doc =
    let allSequenceResults, allBookmarkResults = collectSequenceRefreshTargets doc
    let mutable sequenceCounters: Map<string, int> = Map.empty
    let mutable sequenceResults: Map<string, string list> = Map.empty

    let nextSeqResult label =
        let next = (sequenceCounters |> Map.tryFind label |> Option.defaultValue 0) + 1
        sequenceCounters <- sequenceCounters |> Map.add label next
        let result = string next
        let existing = sequenceResults |> Map.tryFind label |> Option.defaultValue []
        sequenceResults <- sequenceResults |> Map.add label (existing @ [result])
        result

    let fieldPage blockId startOffset =
        pageMap.Fields
        |> List.tryFind (fun anchor -> anchor.PageFieldBlockId = blockId && anchor.PageFieldStartOffset = startOffset)
        |> Option.map (fun anchor -> anchor.PageNumber)
        |> Option.defaultValue 1

    let bookmarkPage bookmarkName =
        pageMap.Bookmarks
        |> List.tryFind (fun anchor -> anchor.PageBookmarkName = bookmarkName)
        |> Option.map (fun anchor -> anchor.PageNumber)
        |> Option.defaultValue 1

    let sequencePageResults =
        let mutable counters: Map<string, int> = Map.empty
        let mutable results: Map<string, string list> = Map.empty
        editableParagraphsInOrder doc
        |> List.iter (fun para ->
            para.Buffer.Runs
            |> List.iter (function
                | Text run ->
                    match run.Style.FieldCode |> Option.map fieldCodeParts with
                    | Some ("SEQ" :: label :: _) ->
                        let next = (counters |> Map.tryFind label |> Option.defaultValue 0) + 1
                        counters <- counters |> Map.add label next
                        let pageNumber = TextBuffer.bookmarkName run.Style |> Option.map bookmarkPage |> Option.defaultValue (fieldPage para.Id run.Start)
                        let existing = results |> Map.tryFind label |> Option.defaultValue []
                        results <- results |> Map.add label (existing @ [sprintf "%d\t%d" next pageNumber])
                    | _ -> ()
                | Image _ -> ()))
        results

    // Ordered (blockId, styleKey, text) for STYLEREF: it shows the text of the
    // nearest paragraph of the referenced style — search backward from the field,
    // then forward (matches Word for inline/body STYLEREF use).
    let orderedBodyParas =
        editableParagraphsInOrder doc
        |> List.map (fun p -> p.Id, normalizeStyleKey (p.Metadata.Style |> Option.defaultValue ""), paragraphText p)

    let styleRefResult blockId styleName =
        let wanted =
            match System.Int32.TryParse(styleName: string) with
            | true, n -> "heading" + string n
            | false, _ -> normalizeStyleKey styleName
        let pick = List.tryPick (fun (_, key, t) -> if key = wanted && t <> "" then Some t else None)
        match orderedBodyParas |> List.tryFindIndex (fun (id, _, _) -> id = blockId) with
        | None -> pick orderedBodyParas
        | Some i ->
            let before = orderedBodyParas |> List.take i |> List.rev
            let after = orderedBodyParas |> List.skip i
            pick before |> Option.orElseWith (fun () -> pick after)

    let sectionFormats = lazy (sectionNumberFormats doc)
    let resultForField blockId startOffset style text =
        match style.FieldCode |> Option.map (fun code -> code, fieldCodeParts code) with
        | Some (code, "DATE" :: _) -> System.DateTime.Now.ToString(dateFieldFormat code)
        | Some (code, "TIME" :: _) -> System.DateTime.Now.ToString(timeFieldFormat code)
        | Some (code, "STYLEREF" :: _) ->
            match styleRefStyleName code with
            | Some name -> styleRefResult blockId name |> Option.defaultValue text
            | None -> text
        | Some (code, "TOC" :: _) ->
            match fieldSwitchValue "\\c" code with
            | Some label -> sequencePageResults |> Map.tryFind label |> Option.defaultValue [] |> String.concat "\n"
            | None -> tocHeadingLines code pageMap doc
        | Some (_, "SEQ" :: label :: _) ->
            nextSeqResult label
        | Some (_, "REF" :: bookmarkName :: _) ->
            allBookmarkResults |> Map.tryFind bookmarkName |> Option.defaultValue text
        // The page's number as its section writes page numbers (w:pgNumType w:fmt:
        // roman in a preface, say), unless the field states a format itself. It was
        // always decimal in the text, while the header beside it said "iv".
        | Some (code, "PAGE" :: _) ->
            let format =
                fieldNumberFormat code
                |> Option.defaultWith (fun () ->
                    match sectionFormats.Value.TryGetValue blockId with
                    | true, sectionFormat -> sectionFormat
                    | false, _ -> "decimal")
            formatPageNumber format (fieldPage blockId startOffset)
        | Some (_, "PAGEREF" :: bookmarkName :: _) -> string (bookmarkPage bookmarkName)
        | Some (_, "NUMPAGES" :: _) -> string pageMap.PageCount
        | _ -> text

    let edits = Dictionary<BlockId, ResizeArray<struct (int * int * int)>>()
    (doc, editableParagraphsInOrder doc)
    ||> List.fold (fun currentDoc para ->
        let mutable buffer = para.Buffer
        let mutable shift = 0
        let paragraphEdits = ResizeArray<struct (int * int * int)>()
        edits.[para.Id] <- paragraphEdits
        for run in para.Buffer.Runs do
            match run with
            | Image _ -> ()
            | Text textRun ->
                match textRun.Style.FieldCode with
                | None -> ()
                | Some _ ->
                    let start = textRun.Start + shift
                    let oldText = TextBuffer.slice start textRun.Length buffer
                    let newText = resultForField para.Id textRun.Start textRun.Style oldText
                    if newText <> oldText then
                        let newStyle = { textRun.Style with FieldResult = Some newText }
                        buffer <- buffer |> TextBuffer.delete start textRun.Length |> TextBuffer.insert start newText |> TextBuffer.applyStyle start newText.Length (fun _ -> newStyle)
                        shift <- shift + newText.Length - textRun.Length
                        paragraphEdits.Add(struct (textRun.Start, textRun.Start + textRun.Length, newText.Length))
                    else
                        buffer <- buffer |> TextBuffer.applyStyle start textRun.Length (fun style -> { style with FieldResult = Some oldText })
        replaceParagraph currentDoc para.Id { para with Buffer = buffer })
    |> selectionAfterEdits edits

let refreshFields doc =
    refreshFieldsWithPageMap (buildPageMap doc) doc

let setCurrentTextBoxStyle x y width height fillColor borderColor borderWidth doc =
    let selectedTextBoxId =
        doc.Selection
        |> Option.bind (fun sel ->
            doc.Blocks
            |> List.tryPick (function
                | TextBox textBox when textBox.Id = sel.Focus.BlockId -> Some textBox.Id
                | _ -> None))
    let fallbackTextBoxId =
        doc.Blocks
        |> List.rev
        |> List.tryPick (function
            | TextBox textBox -> Some textBox.Id
            | _ -> None)
    match selectedTextBoxId |> Option.orElse fallbackTextBoxId with
    | None -> doc
    | Some textBoxId ->
        { doc with
            Blocks =
                doc.Blocks
                |> List.map (function
                    | TextBox textBox when textBox.Id = textBoxId ->
                        TextBox { textBox with
                                    X = Some x
                                    Y = Some y
                                    Width = width
                                    Height = height
                                    FillColor = if System.String.IsNullOrWhiteSpace fillColor then None else Some fillColor
                                    BorderColor = if System.String.IsNullOrWhiteSpace borderColor then None else Some borderColor
                                    BorderWidth = if borderWidth > 0.0 then Some borderWidth else None }
                    | block -> block) }

/// Change the table and its current row together in one document update.
let updateSelectedTableProperties (change: TableBlock -> int -> TableBlock) doc =
    match currentTableCell doc with
    | None -> doc
    | Some (tableId, rowIndex, _) -> updateTable doc tableId (fun table -> change table rowIndex)

/// Change only the requested properties of the cell at the caret.
let updateSelectedTableCell (change: TableCell -> TableCell) doc =
    match currentTableCell doc with
    | None -> doc
    | Some (tableId, rowIndex, colIndex) ->
        updateTable doc tableId (fun table ->
            { table with
                Rows =
                    table.Rows |> List.mapi (fun r row ->
                        if r <> rowIndex then row
                        else
                            { row with
                                Cells =
                                    row.Cells |> List.mapi (fun c cell -> if c = colIndex then change cell else cell) }) })

let setSelectedTableCellStyleWithShading backgroundColor shading width borderColor borderWidth verticalAlignment paddingTop paddingRight paddingBottom paddingLeft doc =
    match currentTableCell doc with
    | None -> doc
    | Some (tableId, rowIndex, colIndex) ->
        updateTable doc tableId (fun table ->
            { table with
                Rows =
                    table.Rows
                    |> List.mapi (fun r row ->
                        if r <> rowIndex then row
                        else
                            { row with
                                Cells =
                                    row.Cells
                                    |> List.mapi (fun c cell ->
                                        if c <> colIndex then cell
                                        else
                                            { cell with
                                                BackgroundColor = backgroundColor
                                                Shading = shading
                                                Width = width
                                                BorderColor = borderColor
                                                BorderWidth = borderWidth
                                                VerticalAlignment = verticalAlignment
                                                PaddingTop = paddingTop
                                                PaddingRight = paddingRight
                                                PaddingBottom = paddingBottom
                                                PaddingLeft = paddingLeft }) }) })

/// The shading of the cell at the caret alone: its other properties (width, borders, alignment, margins) stay.
let setSelectedTableCellShading backgroundColor shading doc =
    match currentTableCell doc with
    | None -> doc
    | Some (tableId, rowIndex, colIndex) ->
        updateTable doc tableId (fun table ->
            { table with
                Rows =
                    table.Rows
                    |> List.mapi (fun r row ->
                        if r <> rowIndex then row
                        else { row with Cells = row.Cells |> List.mapi (fun c cell -> if c <> colIndex then cell else { cell with BackgroundColor = backgroundColor; Shading = shading }) }) })

let setSelectedTableCellStyle backgroundColor width borderColor borderWidth verticalAlignment paddingTop paddingRight paddingBottom paddingLeft doc =
    setSelectedTableCellStyleWithShading backgroundColor None width borderColor borderWidth verticalAlignment paddingTop paddingRight paddingBottom paddingLeft doc

let setCurrentTableCellMargins paddingTop paddingRight paddingBottom paddingLeft doc =
    match currentTableCell doc with
    | None -> doc
    | Some (tableId, _, _) ->
        updateTable doc tableId (fun table ->
            { table with
                PaddingTop = paddingTop
                PaddingRight = paddingRight
                PaddingBottom = paddingBottom
                PaddingLeft = paddingLeft })

let setCurrentTableCellMarginsAndSpacing paddingTop paddingRight paddingBottom paddingLeft cellSpacing doc =
    match currentTableCell doc with
    | None -> doc
    | Some (tableId, _, _) ->
        updateTable doc tableId (fun table ->
            { table with
                PaddingTop = paddingTop
                PaddingRight = paddingRight
                PaddingBottom = paddingBottom
                PaddingLeft = paddingLeft
                CellSpacing = cellSpacing })

let private headingFontSize level =
    match level with
    | 1 -> 24.0
    | 2 -> 18.0
    | 3 -> 14.0
    | _ -> 11.0

let private styleHeadingBuffer level buffer =
    let len = TextBuffer.length buffer
    if len = 0 then buffer
    else
        TextBuffer.applyStyle 0 len (fun style ->
            { style with
                Bold = level >= 1 && level <= 3
                FontSize = headingFontSize level }) buffer

/// Paragraphs a paragraph-level command (alignment, lists, indent, heading)
/// applies to: every paragraph the selection touches, wherever it lives
/// (body, table cell or text box). With a caret, just its paragraph.
let private selectedParagraphIds doc =
    match doc.Selection with
    | None -> []
    | Some sel when sel.IsCollapsed -> [ sel.Focus.BlockId ]
    | Some sel ->
        let _, _, blockIds = getNormalizedRange sel doc
        if blockIds.IsEmpty then [ sel.Focus.BlockId ] else blockIds

let private mapSelectedParagraphs (mapper: ParagraphBlock -> ParagraphBlock) doc =
    match selectedParagraphIds doc with
    | [] -> doc
    | ids ->
        let selected = HashSet<BlockId>(ids)
        mapEditableParagraphs (fun para -> if selected.Contains para.Id then mapper para else para) doc

let private mapSelectedMetadata (update: BlockMetadata -> BlockMetadata) doc =
    mapSelectedParagraphs (fun para ->
        let metadata = update para.Metadata
        if metadata = para.Metadata then para else { para with Metadata = metadata }) doc
/// Apply only the requested paragraph metadata as one document edit.
let updateSelectedParagraphMetadata update doc =
    mapSelectedMetadata update doc

/// Paragraphs targeted by paragraph-level formatting.
let paragraphFormattingTargets doc =
    selectedParagraphIds doc |> List.choose (findParagraph doc)

let private selectedParagraphs doc =
    selectedParagraphIds doc |> List.choose (findParagraph doc)

/// Make every selected paragraph a heading. Body paragraphs become Heading
/// blocks; a paragraph inside a table cell or text box cannot be a Heading
/// block in this model, so it only gets the heading's character formatting.
let setHeadingLevel level doc =
    match selectedParagraphIds doc with
    | [] -> doc
    | ids ->
        let selected = HashSet<BlockId>(ids)
        let boundedLevel = max 1 (min 3 level)
        let styled (para: ParagraphBlock) =
            let buffer = styleHeadingBuffer boundedLevel para.Buffer
            if sameReference buffer para.Buffer then para else { para with Buffer = buffer }
        let blocks =
            doc.Blocks
            |> mapListSharing (fun block ->
                match block with
                | Paragraph para when selected.Contains para.Id -> Heading (boundedLevel, styled para)
                | Heading (current, para) when selected.Contains para.Id ->
                    let styledPara = styled para
                    if current = boundedLevel && sameReference styledPara para then block
                    else Heading (boundedLevel, styledPara)
                | _ -> block)
        let withHeadings = if sameReference blocks doc.Blocks then doc else { doc with Blocks = blocks }
        // Body paragraphs are already styled, so this only changes the
        // selected paragraphs inside cells and text boxes.
        mapEditableParagraphs (fun para -> if selected.Contains para.Id then styled para else para) withHeadings

let clearHeading doc =
    match selectedParagraphIds doc with
    | [] -> doc
    | ids ->
        let selected = HashSet<BlockId>(ids)
        let blocks =
            doc.Blocks
            |> mapListSharing (fun block ->
                match block with
                | Heading (_, para) when selected.Contains para.Id -> Paragraph para
                | _ -> block)
        if sameReference blocks doc.Blocks then doc else { doc with Blocks = blocks }

/// The paragraph before or after the given one in the body's reading order
/// (table cells included, text boxes not: see bodyParagraphsInOrder).
let private neighbourParagraph forward doc paragraphId =
    // (From a paragraph of the body to the next one of the body; in a text box that is not entered, within the box.)
    let paragraphs =
        match textBoxOfParagraph doc paragraphId with
        | Some box -> box.Content
        | None -> bodyParagraphsInOrder doc
    match paragraphs |> List.tryFindIndex (fun para -> para.Id = paragraphId) with
    | None -> None
    | Some index ->
        if forward then List.tryItem (index + 1) paragraphs
        elif index > 0 then List.tryItem (index - 1) paragraphs
        else None

/// Position just across the paragraph boundary: the end of the previous
/// paragraph, or the start of the next one.
let private positionAcrossBoundary forward doc paragraphId =
    neighbourParagraph forward doc paragraphId
    |> Option.map (fun para -> { BlockId = para.Id; Offset = if forward then 0 else TextBuffer.length para.Buffer })

let private collapseCharacterSelection forward (selection: Selection) doc =
    let first, last =
        if selection.Anchor.BlockId = selection.Focus.BlockId then
            if selection.Anchor.Offset <= selection.Focus.Offset then selection.Anchor, selection.Focus
            else selection.Focus, selection.Anchor
        else
            let first, last, _ = getNormalizedRange selection doc
            first, last
    let position = if forward then last else first
    { doc with Selection = Some { Anchor = position; Focus = position } }

/// Move cursor left
let moveCursorLeft doc =
    match doc.Selection with
    | None -> doc
    | Some sel when not sel.IsCollapsed -> collapseCharacterSelection false sel doc
    | Some sel when sel.Focus.Offset = 0 ->
        // Start of the paragraph: continue at the end of the previous one
        match positionAcrossBoundary false doc sel.Focus.BlockId with
        | Some position -> { doc with Selection = Some { Anchor = position; Focus = position } }
        | None -> doc
    | Some sel ->
        let step =
            match findParagraph doc sel.Focus.BlockId with
            | Some para -> backwardStep para.Buffer sel.Focus.Offset
            | None -> 1
        let newSel = {
            Anchor = { sel.Focus with Offset = sel.Focus.Offset - step }
            Focus = { sel.Focus with Offset = sel.Focus.Offset - step }
        }
        { doc with Selection = Some newSel }

/// Move cursor right
let moveCursorRight doc =
    match doc.Selection with
    | None -> doc
    | Some sel when not sel.IsCollapsed -> collapseCharacterSelection true sel doc
    | Some sel ->
        match findParagraph doc sel.Focus.BlockId with
        | None -> doc
        | Some para ->
            let maxOffset = TextBuffer.length para.Buffer
            if sel.Focus.Offset >= maxOffset then
                // End of the paragraph: continue at the start of the next one
                match positionAcrossBoundary true doc sel.Focus.BlockId with
                | Some position -> { doc with Selection = Some { Anchor = position; Focus = position } }
                | None -> doc
            else
                let step = forwardStep para.Buffer sel.Focus.Offset
                let newSel = {
                    Anchor = { sel.Focus with Offset = sel.Focus.Offset + step }
                    Focus = { sel.Focus with Offset = sel.Focus.Offset + step }
                }
                { doc with Selection = Some newSel }

/// Extend selection left (Shift+ArrowLeft)
let extendSelectionLeft doc =
    match doc.Selection with
    | None -> doc
    | Some sel when sel.Focus.Offset = 0 ->
        match positionAcrossBoundary false doc sel.Focus.BlockId with
        | Some position -> { doc with Selection = Some { sel with Focus = position } }
        | None -> doc
    | Some sel ->
        // Keep anchor, move focus left
        let step =
            match findParagraph doc sel.Focus.BlockId with
            | Some para -> backwardStep para.Buffer sel.Focus.Offset
            | None -> 1
        let newFocus = { sel.Focus with Offset = sel.Focus.Offset - step }
        { doc with Selection = Some { sel with Focus = newFocus } }

/// Extend selection right (Shift+ArrowRight)
let extendSelectionRight doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        match findParagraph doc sel.Focus.BlockId with
        | None -> doc
        | Some para ->
            let maxOffset = TextBuffer.length para.Buffer
            if sel.Focus.Offset >= maxOffset then
                match positionAcrossBoundary true doc sel.Focus.BlockId with
                | Some position -> { doc with Selection = Some { sel with Focus = position } }
                | None -> doc
            else
                // Keep anchor, move focus right
                let newFocus = { sel.Focus with Offset = sel.Focus.Offset + forwardStep para.Buffer sel.Focus.Offset }
                { doc with Selection = Some { sel with Focus = newFocus } }

/// Get document text (for testing/debugging)
let getText doc =
    let builder = StringBuilder()
    let mutable first = true
    for para in editableParagraphsInOrder doc do
        if first then
            first <- false
        else
            builder.Append('\n') |> ignore
        builder.Append(TextBuffer.toString para.Buffer) |> ignore
    builder.ToString()

let private textForSearch matchCase (text: string) =
    if matchCase then text else text.ToLowerInvariant()

let private findMatchesInText matchCase (query: string) (text: string) =
    if query = "" then
        []
    else
        let source = textForSearch matchCase text
        let needle = textForSearch matchCase query
        let mutable offset = 0
        let mutable matches = []
        let mutable searching = true

        while searching do
            let index = source.IndexOf(needle, offset)
            if index < 0 then
                searching <- false
            else
                matches <- matches @ [{ BlockId = ""; Start = index; End = index + query.Length; Text = text.Substring(index, query.Length) }]
                offset <- index + max 1 query.Length

        matches

let findText query matchCase doc =
    if System.String.IsNullOrEmpty query then
        []
    else
        editableParagraphsInOrder doc
        |> List.collect (fun para ->
            let text = TextBuffer.toString para.Buffer
            findMatchesInText matchCase query text
            |> List.map (fun matchInfo -> { matchInfo with BlockId = para.Id }))

let findTextArray query matchCase doc =
    if System.String.IsNullOrEmpty query then
        [||]
    else
        let matches = ResizeArray<FindMatch>()
        for para in editableParagraphsInOrder doc do
            let text = TextBuffer.toString para.Buffer
            for matchInfo in findMatchesInText matchCase query text do
                matches.Add { matchInfo with BlockId = para.Id }
        matches.ToArray()

let selectFindMatch (matchInfo: FindMatch) doc =
    { doc with
        Selection = Some {
            Anchor = { BlockId = matchInfo.BlockId; Offset = matchInfo.Start }
            Focus = { BlockId = matchInfo.BlockId; Offset = matchInfo.End }
        } }

let replaceAllText query (replacement: string) matchCase doc =
    if query = "" then
        (doc, 0)
    else
        let mutable count = 0
        let edits = Dictionary<BlockId, ResizeArray<struct (int * int * int)>>()

        let replaceParagraphText (para: ParagraphBlock) =
            let ranges = ResizeArray<struct (int * int)>()
            TextBuffer.iterMatchRanges matchCase query para.Buffer (fun start finish ->
                ranges.Add(struct (start, finish)))

            if ranges.Count = 0 then
                para
            else
                count <- count + ranges.Count
                let holdsSelection =
                    match doc.Selection with
                    | Some sel -> sel.Anchor.BlockId = para.Id || sel.Focus.BlockId = para.Id
                    | None -> false
                if holdsSelection then
                    let paragraphEdits = ResizeArray<struct (int * int * int)>(ranges.Count)
                    for struct (start, finish) in ranges do
                        paragraphEdits.Add(struct (start, finish, replacement.Length))
                    edits.[para.Id] <- paragraphEdits
                let originalText = TextBuffer.toString para.Buffer
                let buffer = TextBuffer.replaceRangesWithText originalText ranges replacement para.Buffer
                if sameReference buffer para.Buffer then para
                else { para with Buffer = buffer }

        let updatedDoc = mapEditableParagraphs replaceParagraphText doc |> selectionAfterEdits edits
        (updatedDoc, count)

/// Set selection range
let setSelection anchor focus doc =
    { doc with Selection = Some { Anchor = anchor; Focus = focus } }

/// Select range in current block
let selectRange startOffset endOffset doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        { doc with
            Selection = Some {
                Anchor = { sel.Focus with Offset = startOffset }
                Focus = { sel.Focus with Offset = endOffset }
            } }

/// Select all the text of the body (see bodyParagraphsInOrder; in a text
/// box that is edited the document is the box's text).
let selectAll doc =
    match bodyParagraphsInOrder doc with
    | [] -> doc
    | firstPara :: rest ->
        let lastPara =
            match rest with
            | [] -> firstPara
            | _ -> List.last rest
        let lastLen = TextBuffer.length lastPara.Buffer
        { doc with
            Selection = Some {
                Anchor = { BlockId = firstPara.Id; Offset = 0 }
                Focus = { BlockId = lastPara.Id; Offset = lastLen }
            } }

/// The paragraphs of a selection that is not collapsed, to be mapped: the
/// function given back applies `formatFn paragraph start length buffer` to
/// the selected part of each, until `finished ()` says the rest can stay as
/// it is. (The range is worked out once for all the passes over it.)
let private selectionFormatter (doc: Document) : ((unit -> bool) -> (ParagraphBlock -> int -> int -> TextBuffer.TextBuffer -> TextBuffer.TextBuffer) -> Document) option =
    match doc.Selection with
    | None -> None
    | Some sel when sel.IsCollapsed -> None  // No selection
    | Some sel ->
        let (startPos, endPos, blockIds) = getNormalizedRange sel doc
        let blockCount = List.length blockIds

        if blockCount = 0 then
            None
        else
            let selectedBlockIndexes = Dictionary<BlockId, int>(blockCount)
            blockIds |> List.iteri (fun index blockId -> selectedBlockIndexes.[blockId] <- index)
            // Classify each containing table once, not once per selected cell.
            let mutable plainCellCache: Dictionary<BlockId, HashSet<BlockId>> option = None
            let inPlainTextCell paragraphId =
                match locateBlock doc paragraphId with
                | ValueSome (struct (_, Table table)) ->
                    let cache =
                        match plainCellCache with
                        | Some cache -> cache
                        | None ->
                            let cache = Dictionary<BlockId, HashSet<BlockId>>()
                            plainCellCache <- Some cache
                            cache
                    let paragraphs =
                        match cache.TryGetValue table.Id with
                        | true, paragraphs -> paragraphs
                        | _ ->
                            let paragraphs = HashSet<BlockId>()
                            let rec collect (table: TableBlock) =
                                for row in table.Rows do
                                    for cell in row.Cells do
                                        if cell.EnclosingControls |> List.exists (fun layer -> controlIsPlainText layer.Xml) then
                                            for paragraph in cell.Content do
                                                paragraphs.Add paragraph.Id |> ignore
                                        for nested in cell.NestedTables do collect nested
                            collect table
                            cache.[table.Id] <- paragraphs
                            paragraphs
                    paragraphs.Contains paragraphId
                | _ -> false

            Some (fun finished formatFn ->
                let formatParagraph (para: ParagraphBlock) =
                    match selectedBlockIndexes.TryGetValue para.Id with
                    | false, _ -> para
                    | true, i ->
                        let bufferLen = TextBuffer.length para.Buffer

                        // Determine range to format within this block
                        let (formatStart, formatLen) =
                            if blockCount = 1 then
                                // Single block selection
                                (startPos.Offset, endPos.Offset - startPos.Offset)
                            elif i = 0 then
                                // First block: from startOffset to end
                                (startPos.Offset, bufferLen - startPos.Offset)
                            elif i = blockCount - 1 then
                                // Last block: from 0 to endOffset
                                (0, endPos.Offset)
                            else
                                // Middle block: format entire block
                                (0, bufferLen)

                        // Apply formatting
                        if formatLen > 0 then
                            let selectedGroups =
                                para.Buffer.Runs |> List.collect (function
                                    | TextBuffer.Text run when run.Start < formatStart + formatLen && run.Start + run.Length > formatStart ->
                                        run.Style.ContentControls |> List.choose (fun layer -> if controlIsPlainText layer.Xml then Some layer.Group else None)
                                    | _ -> []) |> Set.ofList
                            let mutable first = formatStart
                            let mutable last = formatStart + formatLen
                            if not selectedGroups.IsEmpty then
                                for run in para.Buffer.Runs do
                                    match run with
                                    | TextBuffer.Text text when text.Style.ContentControls |> List.exists (fun layer -> Set.contains layer.Group selectedGroups) ->
                                        first <- min first text.Start
                                        last <- max last (text.Start + text.Length)
                                    | _ -> ()
                            let isWholePlainControl =
                                (para.Metadata.ContentControlXml :: ((para.Metadata.ContentControlAncestors @ para.Metadata.EnclosingCellControls) |> List.map (fun layer -> layer.Xml))
                                 |> List.exists controlIsPlainText)
                                || inPlainTextCell para.Id
                            let newBuffer =
                                if isWholePlainControl then formatFn para 0 bufferLen para.Buffer
                                else formatFn para first (last - first) para.Buffer
                            if sameReference newBuffer para.Buffer then para
                            else { para with Buffer = newBuffer }
                        else
                            para

                mapEditableParagraphsUntil finished formatParagraph doc)

/// Apply formatting to selection; the formatter also gets the paragraph.
let applyFormattingToSelectionInParagraph (formatFn: ParagraphBlock -> int -> int -> TextBuffer.TextBuffer -> TextBuffer.TextBuffer) doc =
    match selectionFormatter doc with
    | Some format -> format (fun () -> false) formatFn
    | None -> doc

/// Apply formatting to selection
let applyFormattingToSelection formatFn doc =
    applyFormattingToSelectionInParagraph (fun _ start len buffer -> formatFn start len buffer) doc

/// Toggle an on/off character property over the whole selection, deciding
/// once for all of it: clear when every selected run already has it, set
/// otherwise. The decision spans paragraphs, so a selection covering a bold
/// paragraph and a plain one becomes bold throughout.
///
/// `effective` resolves a run style against its paragraph's styles, so the
/// decision is made on what the user sees: text that is bold only through its
/// paragraph style counts as bold, and toggling it yields explicit non-bold.
/// `apply` must record the value as explicit (TextBuffer.withExplicit*).
let private toggleOnSelection (effective: ParagraphBlock -> TextBuffer.TextStyle -> TextBuffer.TextStyle) (isSet: TextBuffer.TextStyle -> bool) (apply: bool -> TextBuffer.TextStyle -> TextBuffer.TextStyle) doc =
    match selectionFormatter doc with
    | None -> doc
    | Some format ->
        // (The first run that has not got it decides: the paragraphs after it are not looked at.)
        let mutable allSet = true
        format (fun () -> not allSet) (fun para start len buffer ->
            if allSet && not (TextBuffer.rangeSatisfies start len (fun style -> isSet (effective para style)) buffer) then
                allSet <- false
            buffer)
        |> ignore
        let turnOn = not allSet
        format (fun () -> false) (fun _ start len buffer -> TextBuffer.applyStyle start len (apply turnOn) buffer)

let private rawStyle (_: ParagraphBlock) (style: TextBuffer.TextStyle) = style

/// Style-aware toggles: `effective` maps a paragraph and a run style to the
/// resolved style (see WordStyles.toggleBold and friends, which supply it).
let toggleBoldWith effective doc = toggleOnSelection effective (fun style -> style.Bold) TextBuffer.withExplicitBold doc
let toggleItalicWith effective doc = toggleOnSelection effective (fun style -> style.Italic) TextBuffer.withExplicitItalic doc
// (A link shows an underline its text need not state: that one is switched off first.)
let toggleUnderlineWith effective doc = toggleOnSelection effective TextBuffer.underlineShown TextBuffer.withExplicitUnderline doc
let toggleAllCapsWith effective doc = toggleOnSelection effective (fun style -> style.AllCaps) TextBuffer.withExplicitAllCaps doc
let toggleSmallCapsWith effective doc = toggleOnSelection effective (fun style -> style.SmallCaps) TextBuffer.withExplicitSmallCaps doc
let toggleStrikeWith effective doc = toggleOnSelection effective (fun style -> style.Strike) TextBuffer.withExplicitStrike doc

/// Toggle bold on selection (decided on the runs' own formatting; hosts with
/// named styles should use WordStyles.toggleBold).
let toggleBold doc = toggleBoldWith rawStyle doc

/// Toggle italic on selection
let toggleItalic doc = toggleItalicWith rawStyle doc

/// Toggle underline on selection
let toggleUnderline doc = toggleUnderlineWith rawStyle doc

let private isCaseWordChar (ch: char) = System.Char.IsLetterOrDigit ch || ch = '\'' || ch = '’'

/// Change Case (Word's Shift+F3) on the selection, or with nothing selected
/// on the word at the caret (the caret stays): text with no capitals becomes
/// Capitalized Each Word, text with some becomes UPPER CASE, text in upper
/// case lower case - one after the other, as Word cycles. The characters
/// keep their formatting.
let changeCase doc =
    let caretWord =
        match doc.Selection with
        | Some sel when sel.IsCollapsed ->
            match findParagraph doc sel.Focus.BlockId with
            | Some para ->
                let text = TextBuffer.toString para.Buffer
                let at = max 0 (min text.Length sel.Focus.Offset)
                let mutable first = at
                while first > 0 && isCaseWordChar text.[first - 1] do first <- first - 1
                let mutable last = at
                while last < text.Length && isCaseWordChar text.[last] do last <- last + 1
                if last > first then Some (setSelection { sel.Focus with Offset = first } { sel.Focus with Offset = last } doc) else None
            | None -> None
        | _ -> None
    let target = caretWord |> Option.defaultValue doc
    match selectionFormatter target with
    | None -> doc
    | Some format ->
        let mutable anyUpper = false
        let mutable anyLower = false
        format (fun () -> false) (fun _ start len buffer ->
            let text = TextBuffer.toString buffer
            for i = start to min text.Length (start + len) - 1 do
                if System.Char.IsUpper text.[i] then anyUpper <- true
                elif System.Char.IsLower text.[i] then anyLower <- true
            buffer)
        |> ignore
        if not anyUpper && not anyLower then doc
        else
            let map: string -> int -> char =
                if not anyUpper then
                    fun text i ->
                        let ch = text.[i]
                        if i = 0 || not (isCaseWordChar text.[i - 1]) then System.Char.ToUpperInvariant ch else ch
                elif not anyLower then fun text i -> System.Char.ToLowerInvariant text.[i]
                else fun text i -> System.Char.ToUpperInvariant text.[i]
            let changed = format (fun () -> false) (fun _ start len buffer -> TextBuffer.mapCharacters start len map buffer)
            { changed with Selection = doc.Selection }

/// Toggle all caps on selection
let toggleAllCaps doc = toggleAllCapsWith rawStyle doc

/// Toggle small caps on selection
let toggleSmallCaps doc = toggleSmallCapsWith rawStyle doc

/// Toggle strikethrough on selection
let toggleStrike doc = toggleStrikeWith rawStyle doc

/// Clear direct character formatting on selection.
let clearFormatting doc =
    applyFormattingToSelection TextBuffer.clearFormatting doc

/// Set vertical alignment on selection
let setVertAlign align doc =
    applyFormattingToSelection (fun start len buf -> TextBuffer.setVertAlign start len align buf) doc

/// Set text color on selection
let setColor color doc =
    applyFormattingToSelection (fun start len buf -> TextBuffer.setColor start len color buf) doc

/// Set highlight color on selection
let setHighlight color doc =
    applyFormattingToSelection (fun start len buf -> TextBuffer.setHighlight start len color buf) doc

/// Set hyperlink on selection
let setHyperlink target doc =
    applyFormattingToSelection (fun start len buf -> TextBuffer.setHyperlink start len target buf) doc

/// Set comment metadata on selection
let setComment commentId commentText author date doc =
    applyFormattingToSelection (fun start len buf -> TextBuffer.setComment start len commentId commentText author date buf) doc

/// Set footnote/endnote metadata on selection
let setNote noteId noteType noteText doc =
    applyFormattingToSelection (fun start len buf -> TextBuffer.setNote start len noteId noteType noteText buf) doc

// ---- the body of a note or comment edited in place --------------------------

/// The blocks with each paragraph (those of tables and text boxes too) mapped;
/// a block nothing of which the mapper changed is the object it was.
let private mapParagraphsKeeping (mapper: ParagraphBlock -> ParagraphBlock) (blocks: Block list) : Block list =
    let mutable changed = false
    let mapList (paras: ParagraphBlock list) =
        let mapped = paras |> List.map mapper
        if List.forall2 (fun (a: ParagraphBlock) b -> obj.ReferenceEquals(a, b)) paras mapped then paras else mapped
    let rec mapTable (table: TableBlock) : TableBlock =
        let rows =
            table.Rows
            |> List.map (fun row ->
                let cells =
                    row.Cells
                    |> List.map (fun cell ->
                        let content = mapList cell.Content
                        let nested = cell.NestedTables |> List.map mapTable
                        if obj.ReferenceEquals(content, cell.Content) && List.forall2 (fun (a: TableBlock) b -> obj.ReferenceEquals(a, b)) cell.NestedTables nested then cell
                        else { cell with Content = content; NestedTables = nested })
                if List.forall2 (fun (a: TableCell) b -> obj.ReferenceEquals(a, b)) row.Cells cells then row else { row with Cells = cells })
        if List.forall2 (fun (a: TableRow) b -> obj.ReferenceEquals(a, b)) table.Rows rows then table else { table with Rows = rows }
    let mapped =
        blocks
        |> List.map (fun block ->
            let next =
                match block with
                | Paragraph para -> (let p = mapper para in if obj.ReferenceEquals(p, para) then block else Paragraph p)
                | Heading (level, para) -> (let p = mapper para in if obj.ReferenceEquals(p, para) then block else Heading (level, p))
                | Table table -> (let t = mapTable table in if obj.ReferenceEquals(t, table) then block else Table t)
                | TextBox textBox -> (let content = mapList textBox.Content in if obj.ReferenceEquals(content, textBox.Content) then block else TextBox { textBox with Content = content })
            if not (obj.ReferenceEquals(next, block)) then changed <- true
            next)
    if changed then mapped else blocks

/// The plain text a note's or comment's reference shows of its body, as the
/// reader gives it: the text of its paragraphs, joined (a comment's, those
/// with text, parted by a space).
let storyPlainText (kind: string) (blocks: Block list) : string =
    let texts = paragraphsOfBlocks blocks |> List.map (fun para -> TextBuffer.toString para.Buffer |> String.filter (fun c -> c <> '\uFFFC'))
    if kind = "comment" then texts |> List.filter (fun text -> text <> "") |> String.concat " "
    else String.concat "" texts

/// The document with `blocks` as the body of the note (kind "footnote" |
/// "endnote") or comment (kind "comment") with this id: the body in the
/// package, its plain text there and on the references that show it (the
/// runs of the text that refer to it). What does not refer to it is the
/// object it was.
let rec withStoryBody (kind: string) (id: string) (blocks: Block list) (doc: Document) : Document =
    withStoryBodyAt kind id blocks None doc

/// The indexes of the blocks of the body that refer to the note (kind
/// "footnote" | "endnote") or comment with this id (see withStoryBodyAt).
and storyReferenceIndexes (kind: string) (id: string) (doc: Document) : int array =
    let refers (style: TextBuffer.TextStyle) =
        if kind = "comment" then style.CommentId = Some id
        else style.NoteId = Some id && (style.NoteType |> Option.defaultValue "footnote") = kind
    doc.Blocks
    |> List.mapi (fun index block -> index, block)
    |> List.filter (fun (_, block) ->
        paragraphsOfBlocks [ block ] |> List.exists (fun para -> para.Buffer.Runs |> List.exists (function TextBuffer.Text run -> refers run.Style | TextBuffer.Image _ -> false)))
    |> List.map fst
    |> List.toArray

/// withStoryBody, the body's blocks other than those at `indexes` (when
/// given: storyReferenceIndexes of `doc`) taken as they are without a look
/// at their runs - what a keystroke in a note of a long document costs.
and withStoryBodyAt (kind: string) (id: string) (blocks: Block list) (indexes: int array option) (doc: Document) : Document =
    let text = storyPlainText kind blocks
    let refers (style: TextBuffer.TextStyle) =
        if kind = "comment" then style.CommentId = Some id && style.CommentText <> Some text
        else style.NoteId = Some id && (style.NoteType |> Option.defaultValue "footnote") = kind && style.NoteText <> Some text
    let mapper (para: ParagraphBlock) =
        let found = para.Buffer.Runs |> List.exists (function TextBuffer.Text run -> refers run.Style | TextBuffer.Image _ -> false)
        if not found then para
        else
            let runs =
                para.Buffer.Runs
                |> List.map (function
                    | TextBuffer.Text run when refers run.Style ->
                        TextBuffer.Text { run with Style = (if kind = "comment" then { run.Style with CommentText = Some text } else { run.Style with NoteText = Some text }) }
                    | other -> other)
            { para with Buffer = { para.Buffer with Runs = runs } }
    let parts (list: HeaderFooterPart list) =
        let mapped = list |> List.map (fun part -> let blocks = mapParagraphsKeeping mapper part.Blocks in if obj.ReferenceEquals(blocks, part.Blocks) then part else { part with Blocks = blocks })
        if List.forall2 (fun (a: HeaderFooterPart) b -> obj.ReferenceEquals(a, b)) list mapped then list else mapped
    let package =
        if kind = "comment" then
            { doc.Package with Comments = doc.Package.Comments |> List.map (fun body -> if body.Id = id then { body with Blocks = blocks; Text = text } else body) }
        else
            { doc.Package with Notes = doc.Package.Notes |> List.map (fun body -> if body.Id = id && body.NoteType = kind && body.Kind.IsNone then { body with Blocks = blocks; Text = text } else body) }
    let body =
        match indexes with
        | None -> mapParagraphsKeeping mapper doc.Blocks
        | Some indexes when indexes.Length = 0 -> doc.Blocks
        | Some indexes ->
            let wanted = System.Collections.Generic.HashSet<int>(indexes)
            doc.Blocks |> List.mapi (fun index block -> if wanted.Contains index then (match mapParagraphsKeeping mapper [ block ] with [ mapped ] -> mapped | _ -> block) else block)
    { doc with
        Blocks = body
        HeaderParts = parts doc.HeaderParts
        FooterParts = parts doc.FooterParts
        Package = package }

/// The body of the note or comment with this id (see withStoryBody), when the document has it.
let storyBody (kind: string) (id: string) (doc: Document) : Block list option =
    if kind = "comment" then doc.Package.Comments |> List.tryFind (fun body -> body.Id = id) |> Option.map (fun body -> body.Blocks)
    else doc.Package.Notes |> List.tryFind (fun body -> body.Id = id && body.NoteType = kind && body.Kind.IsNone) |> Option.map (fun body -> body.Blocks)

/// Set bookmark markers on selection
let setBookmark bookmarkName doc =
    let normalizedName =
        if System.String.IsNullOrWhiteSpace bookmarkName then "Bookmark"
        else bookmarkName.Trim().Replace(" ", "_")
    let bookmarkId = nextBookmarkId doc
    applyFormattingToSelection (fun start len buf -> TextBuffer.setBookmark start len bookmarkId normalizedName buf) doc

/// Set content control metadata on selected paragraphs
let setContentControl controlId tag alias doc =
    let normalize fallback (value: string) =
        if System.String.IsNullOrWhiteSpace value then fallback else value.Trim()
    let controlId = normalize "1" controlId
    let tag = normalize "content-control" tag
    let alias = normalize tag alias
    let blockIds =
        match doc.Selection with
        | None -> []
        | Some sel when sel.IsCollapsed -> [sel.Focus.BlockId]
        | Some sel ->
            let (_, _, ids) = getNormalizedRange sel doc
            ids
    let group = if blockIds.IsEmpty then None else Some (nextBlockId ())
    (doc, blockIds)
    ||> List.fold (fun currentDoc blockId ->
        match findParagraph currentDoc blockId with
        | None -> currentDoc
        | Some para ->
            replaceParagraph currentDoc para.Id {
                para with
                    Metadata = {
                        para.Metadata with
                            ContentControlId = Some controlId
                            ContentControlTag = Some tag
                            ContentControlAlias = Some alias
                            ContentControlGroup = group
                    }
            })

/// Set named paragraph style on selected paragraphs
let setParagraphStyle styleId doc =
    let normalizedStyle =
        if System.String.IsNullOrWhiteSpace styleId then None else Some (styleId.Trim())
    let blockIds =
        match doc.Selection with
        | None -> []
        | Some sel when sel.IsCollapsed -> [sel.Focus.BlockId]
        | Some sel ->
            let (_, _, ids) = getNormalizedRange sel doc
            ids
    (doc, blockIds)
    ||> List.fold (fun currentDoc blockId ->
        match findParagraph currentDoc blockId with
        | None -> currentDoc
        | Some para ->
            replaceParagraph currentDoc para.Id {
                para with Metadata = { para.Metadata with Style = normalizedStyle }
            })

/// Word's Ctrl+Shift+N: the selected paragraphs in the Normal style (the
/// document's default paragraph style), headings no longer.
let setNormalStyle doc =
    let cleared = clearHeading doc
    let paragraphs = selectedParagraphIds cleared |> List.choose (findParagraph cleared)
    if paragraphs |> List.forall (fun para -> para.Metadata.Style.IsNone) then cleared
    else setParagraphStyle "" cleared

/// Set named character style on selection
let setCharacterStyle styleId doc =
    let normalizedStyle =
        if System.String.IsNullOrWhiteSpace styleId then None else Some (styleId.Trim())
    applyFormattingToSelection (fun start len buf -> TextBuffer.setCharacterStyle start len normalizedStyle buf) doc

/// Set revision metadata on selection
let setRevision revisionType revisionId author date doc =
    let normalizedType = if revisionType = "delete" then "delete" else "insert"
    let revisionId = if System.String.IsNullOrWhiteSpace revisionId then "1" else revisionId
    let author = if System.String.IsNullOrWhiteSpace author then None else Some author
    let date = if System.String.IsNullOrWhiteSpace date then None else Some date
    applyFormattingToSelection (fun start len buf -> TextBuffer.setRevision start len normalizedType revisionId author date buf) doc

let private setRevisionWithOptions revisionType revisionId author date doc =
    applyFormattingToSelection (fun start len buf -> TextBuffer.setRevision start len revisionType revisionId author date buf) doc

let private trackDeleteRange revisionId author date doc =
    applyFormattingToSelection (fun start len buf -> TextBuffer.trackDelete start len revisionId author date buf) doc

let private adjacentInsertRevisionId doc (pos: Position) author date =
    findParagraph doc pos.BlockId
    |> Option.bind (fun para ->
        para.Buffer.Runs
        |> List.tryPick (function
            | TextBuffer.Text tr when tr.Start + tr.Length = pos.Offset && tr.Style.RevisionType = Some "insert" && tr.Style.RevisionAuthor = author && tr.Style.RevisionDate = date -> tr.Style.RevisionId
            | _ -> None))

/// Insert text while recording the inserted run as a revision.
let insertTrackedText author date text doc =
    let doc = snapSelectionToMarks None doc
    let author = if System.String.IsNullOrWhiteSpace author then None else Some author
    let date = if System.String.IsNullOrWhiteSpace date then None else Some date
    let deletionRevisionId = nextRevisionId doc
    let docForInsert =
        match doc.Selection with
        | Some sel when not sel.IsCollapsed ->
            let startPos, _, _ = getNormalizedRange sel doc
            let marked = trackDeleteRange deletionRevisionId author date doc
            { marked with Selection = Some { Anchor = startPos; Focus = startPos } }
        | _ -> doc

    let insertionRevisionId =
        match docForInsert.Selection with
        | Some sel -> adjacentInsertRevisionId docForInsert sel.Focus author date |> Option.defaultValue (nextRevisionId docForInsert)
        | None -> nextRevisionId docForInsert
    let inserted = insertText text docForInsert
    match inserted.Selection with
    | None -> inserted
    | Some sel when text.Length > 0 ->
        let startOffset = max 0 (sel.Focus.Offset - text.Length)
        let revisionSelection = {
            Anchor = { sel.Focus with Offset = startOffset }
            Focus = sel.Focus
        }
        let styled = setRevisionWithOptions "insert" insertionRevisionId author date { inserted with Selection = Some revisionSelection }
        { styled with Selection = Some sel }
    | _ -> inserted

/// Mark the selected text as deleted instead of physically removing it.
/// The author's own pending insertions in the selection are removed outright.
let trackDeleteSelection author date doc =
    match doc.Selection with
    | None -> doc
    | Some sel when sel.IsCollapsed -> doc
    | Some sel ->
        let startPos, _, _ = getNormalizedRange sel doc
        let author = if System.String.IsNullOrWhiteSpace author then None else Some author
        let date = if System.String.IsNullOrWhiteSpace date then None else Some date
        let marked = trackDeleteRange (nextRevisionId doc) author date doc
        { marked with Selection = Some { Anchor = startPos; Focus = startPos } }

let private isDeletedRun (run: TextBuffer.TextRun) =
    TextBuffer.isRemovalRevision run.Style

let private topLevelParagraphAt (doc: Document) index =
    if index < 0 || index >= doc.Blocks.Length then None
    else
        match doc.Blocks.[index] with
        | Paragraph para
        | Heading (_, para) -> Some para
        | _ -> None

let private withMarkRevision (revision: ParagraphMarkRevision option) (doc: Document) (paragraphId: BlockId) =
    match findParagraph doc paragraphId with
    | Some para -> replaceParagraph doc para.Id { para with Metadata = { para.Metadata with MarkRevision = revision } }
    | None -> doc

let private revisionAuthor (author: string) = if System.String.IsNullOrWhiteSpace author then None else Some author
let private revisionDate (date: string) = if System.String.IsNullOrWhiteSpace date then None else Some date

/// Enter under track changes: the paragraph is split and the new paragraph
/// mark (the one that now ends the first half) is recorded as inserted, so the
/// split can be accepted or rejected. The second half keeps the original mark.
let trackSplitBlock author date doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        let original = findParagraph doc sel.Focus.BlockId
        let isTopLevel = (findBlockIndex doc sel.Focus.BlockId).IsSome
        let split = splitBlock doc
        match original with
        | Some para when isTopLevel && not (obj.ReferenceEquals(split, doc)) ->
            let inserted =
                { RevisionType = "insert"
                  Mark = { Id = Some (nextRevisionId split); Author = revisionAuthor author; Date = revisionDate date } }
            withMarkRevision (Some inserted) split para.Id
        | _ -> split

/// Joining two paragraphs under track changes (Backspace at the start of the
/// second, Delete at the end of the first): the paragraphs stay apart and the
/// first one's mark is recorded as deleted. A mark the same author inserted
/// (and that is still pending) is simply removed again.
let private trackDeleteMark author date (first: ParagraphBlock) (second: ParagraphBlock) cursor doc =
    let author = revisionAuthor author
    match first.Metadata.MarkRevision with
    | Some revision when revision.RevisionType = "insert" && revision.Mark.Author = author ->
        let offset = TextBuffer.length first.Buffer
        let merged = withMarkRevision second.Metadata.MarkRevision (mergeWithNext doc first.Id) first.Id
        { merged with Selection = Some { Anchor = { BlockId = first.Id; Offset = offset }; Focus = { BlockId = first.Id; Offset = offset } } }
    | Some revision when revision.RevisionType = "delete" ->
        { doc with Selection = Some { Anchor = cursor; Focus = cursor } }
    | _ ->
        let deleted =
            { RevisionType = "delete"
              Mark = { Id = Some (nextRevisionId doc); Author = author; Date = revisionDate date } }
        let marked = withMarkRevision (Some deleted) doc first.Id
        { marked with Selection = Some { Anchor = cursor; Focus = cursor } }

let private nextTrackDeletableOffset para offset =
    para.Buffer.Runs
    |> List.choose (function
        | TextBuffer.Text tr when not (isDeletedRun tr) && tr.Start + tr.Length > offset ->
            Some (max offset tr.Start)
        | _ -> None)
    |> List.tryHead

let private previousTrackDeletableOffset para offset =
    para.Buffer.Runs
    |> List.choose (function
        | TextBuffer.Text tr when not (isDeletedRun tr) && tr.Start < offset ->
            Some ((min offset (tr.Start + tr.Length)) - 1)
        | _ -> None)
    |> List.filter (fun value -> value >= 0)
    |> List.tryLast

let trackDeleteForward author date doc =
    match doc.Selection with
    | None -> doc
    | Some sel when not sel.IsCollapsed -> trackDeleteSelection author date doc
    | Some _ when (markSelectedAtCaret false doc).IsSome -> trackDeleteSelection author date (markSelectedAtCaret false doc).Value
    | Some sel ->
        match findParagraph doc sel.Focus.BlockId with
        | None -> doc
        | Some para ->
            match nextTrackDeletableOffset para sel.Focus.Offset with
            | None ->
                // Nothing left to delete in this paragraph: Delete at its end joins it with the next one.
                match findBlockIndex doc para.Id |> Option.bind (fun index -> topLevelParagraphAt doc (index + 1)) with
                | Some next when sel.Focus.Offset >= TextBuffer.length para.Buffer -> trackDeleteMark author date para next { BlockId = next.Id; Offset = 0 } doc
                | _ -> doc
            | Some offset ->
                let startPos = { sel.Focus with Offset = offset }
                let selection = { Anchor = startPos; Focus = { startPos with Offset = offset + forwardStep para.Buffer offset } }
                trackDeleteSelection author date { doc with Selection = Some selection }

let trackDeleteBackward author date doc =
    match doc.Selection with
    | None -> doc
    | Some sel when not sel.IsCollapsed -> trackDeleteSelection author date doc
    | Some _ when (markSelectedAtCaret true doc).IsSome -> trackDeleteSelection author date (markSelectedAtCaret true doc).Value
    | Some sel ->
        match findParagraph doc sel.Focus.BlockId with
        | None -> doc
        | Some para ->
            match previousTrackDeletableOffset para sel.Focus.Offset with
            | None ->
                // Backspace at the start of a paragraph joins it with the previous one.
                match findBlockIndex doc para.Id |> Option.bind (fun index -> topLevelParagraphAt doc (index - 1)) with
                | Some previous when sel.Focus.Offset = 0 -> trackDeleteMark author date previous para { BlockId = previous.Id; Offset = TextBuffer.length previous.Buffer } doc
                | _ -> doc
            | Some offset ->
                let startPos = { sel.Focus with Offset = offset + 1 - backwardStep para.Buffer (offset + 1) }
                let selection = { Anchor = startPos; Focus = { startPos with Offset = offset + 1 } }
                trackDeleteSelection author date { doc with Selection = Some selection }

type private RevisionEdit = {
    Start: int
    Length: int
    RevisionType: string
}

let private selectionRangeForBlock blockId doc =
    match doc.Selection with
    | None -> None
    | Some sel when sel.IsCollapsed ->
        if sel.Focus.BlockId = blockId then Some (sel.Focus.Offset, sel.Focus.Offset, true) else None
    | Some sel ->
        let startPos, endPos, blockIds = getNormalizedRange sel doc
        if blockIds |> List.contains blockId then
            match findParagraph doc blockId with
            | None -> None
            | Some para ->
                let bufferLen = TextBuffer.length para.Buffer
                let startOffset = if blockId = startPos.BlockId then startPos.Offset else 0
                let endOffset = if blockId = endPos.BlockId then endPos.Offset else bufferLen
                Some (startOffset, endOffset, false)
        else None

let private selectedRevisionEdits para selectionRange =
    let selStart, selEnd, isCollapsed = selectionRange
    let intersects start endPos selStart selEnd isCollapsed =
        if isCollapsed then selStart >= start && selStart <= endPos
        else start < selEnd && endPos > selStart

    para.Buffer.Runs
    |> List.choose (function
        | Text tr ->
            match tr.Style.RevisionType with
            | Some revisionType ->
                let runEnd = tr.Start + tr.Length
                if intersects tr.Start runEnd selStart selEnd isCollapsed then
                    Some { Start = tr.Start; Length = tr.Length; RevisionType = revisionType }
                else None
            | None -> None
        | Image _ -> None)

let private selectedRevisionSpans accept para selectionRange =
    let selStart, selEnd, isCollapsed = selectionRange
    let struct (_, firstStart, changed) =
        TextBuffer.resolveRevisionsInRange accept selStart selEnd isCollapsed para.Buffer
    if changed then
        match firstStart with
        | Some start -> [ { Start = start; Length = 0; RevisionType = "" } ]
        | None -> [ { Start = 0; Length = 0; RevisionType = "" } ]
    else
        []

let revisionInfos doc =
    editableParagraphsInOrder doc
    |> List.collect (fun para ->
        (para.Buffer.Runs
         |> List.choose (function
            | Text tr ->
                tr.Style.RevisionType
                |> Option.map (fun revisionType -> {
                    BlockId = para.Id
                    Start = tr.Start
                    End = tr.Start + tr.Length
                    RevisionType = revisionType
                    Author = tr.Style.RevisionAuthor
                    Date = tr.Style.RevisionDate
                    Text = TextBuffer.slice tr.Start tr.Length para.Buffer
                })
            | Image _ -> None))
        // A tracked paragraph mark, at the end of its paragraph.
        @ (match para.Metadata.MarkRevision with
           | Some revision ->
               let length = TextBuffer.length para.Buffer
               [ { BlockId = para.Id; Start = length; End = length; RevisionType = revision.RevisionType
                   Author = revision.Mark.Author; Date = revision.Mark.Date; Text = "\u00B6" } ]
           | None -> []))

/// Accept or reject the tracked paragraph marks the selection covers: a mark is
/// covered when the selection runs past the end of its paragraph (or, for the
/// last paragraph, when `includeLastMark`; a collapsed cursor at the end of the
/// paragraph covers it too). Accepting an inserted mark or rejecting a deleted
/// one keeps the paragraphs apart; the other way round joins them.
let private applyMarkRevisionDecision accept includeLastMark (selection: Selection option) doc =
    match selection with
    | None -> doc, false
    | Some sel ->
        let startPos, endPos, blockIds =
            if sel.IsCollapsed then sel.Focus, sel.Focus, [ sel.Focus.BlockId ]
            else getNormalizedRange sel doc
        let mutable updated = doc
        let mutable changed = false
        for blockId in List.rev blockIds do
            match findParagraph updated blockId with
            | Some para when para.Metadata.MarkRevision.IsSome ->
                let length = TextBuffer.length para.Buffer
                let covered =
                    if blockId <> endPos.BlockId then true
                    elif sel.IsCollapsed then endPos.Offset >= length
                    else includeLastMark && endPos.Offset >= length
                if covered then
                    let revision = para.Metadata.MarkRevision.Value
                    let join = (accept && revision.RevisionType = "delete") || (not accept && revision.RevisionType = "insert")
                    let next = findBlockIndex updated blockId |> Option.bind (fun index -> topLevelParagraphAt updated (index + 1))
                    match join, next with
                    | true, Some nextPara ->
                        // The joined paragraph ends with the next one's mark.
                        updated <- withMarkRevision nextPara.Metadata.MarkRevision (mergeWithNext updated blockId) blockId
                    | _ -> updated <- withMarkRevision None updated blockId
                    changed <- true
            | _ -> ()
        ignore startPos
        updated, changed

let private applyRevisionDecisionWith includeLastMark accept doc =
    let selectedBlockIds =
        match doc.Selection with
        | None -> []
        | Some sel when sel.IsCollapsed -> [sel.Focus.BlockId]
        | Some sel ->
            let _, _, blockIds = getNormalizedRange sel doc
            blockIds

    let mutable updatedDoc = doc
    let mutable changed = false
    let mutable firstSelection: Position option = None

    for blockId in selectedBlockIds do
        match findParagraph updatedDoc blockId, selectionRangeForBlock blockId updatedDoc with
        | Some para, Some selectionRange ->
            let selStart, selEnd, isCollapsed = selectionRange
            let struct (buffer, firstStart, bufferChanged) =
                TextBuffer.resolveRevisionsInRange accept selStart selEnd isCollapsed para.Buffer
            if bufferChanged then
                updatedDoc <- replaceParagraph updatedDoc para.Id { para with Buffer = buffer }
                changed <- true
                match firstSelection with
                | None -> firstSelection <- Some { BlockId = para.Id; Offset = firstStart |> Option.defaultValue selStart }
                | Some _ -> ()
        | _ -> ()

    // Paragraph marks after the text: joining paragraphs changes the block list.
    let markedDoc, marksChanged = applyMarkRevisionDecision accept includeLastMark doc.Selection updatedDoc
    if changed then
        let selection = firstSelection |> Option.map (fun pos -> { Anchor = pos; Focus = pos })
        { markedDoc with Selection = selection |> Option.orElse markedDoc.Selection }
    elif marksChanged then
        let cursor =
            doc.Selection
            |> Option.map (fun sel -> if sel.IsCollapsed then sel.Focus else (let startPos, _, _ = getNormalizedRange sel doc in startPos))
            |> Option.filter (fun pos -> (findParagraph markedDoc pos.BlockId).IsSome)
        { markedDoc with Selection = cursor |> Option.map (fun pos -> { Anchor = pos; Focus = pos }) |> Option.orElse markedDoc.Selection }
    else updatedDoc

let private applyRevisionDecision accept doc = applyRevisionDecisionWith false accept doc

/// Accept selected revision runs. Insertions become normal text; deletions are removed.
let acceptSelectedRevision doc =
    applyRevisionDecision true doc

/// Reject selected revision runs. Insertions are removed; deletions become normal text.
let rejectSelectedRevision doc =
    applyRevisionDecision false doc

let private allRevisionSelection doc =
    match revisionInfos doc with
    | [] -> None
    | revisions ->
        let first = revisions |> List.head
        let last = revisions |> List.last
        Some {
            Anchor = { BlockId = first.BlockId; Offset = first.Start }
            Focus = { BlockId = last.BlockId; Offset = last.End }
        }

/// Accept all revision runs in document order.
let acceptAllRevisions doc =
    match allRevisionSelection doc with
    | None -> doc
    | Some selection -> applyRevisionDecisionWith true true { doc with Selection = Some selection }

/// Reject all revision runs in document order.
let rejectAllRevisions doc =
    match allRevisionSelection doc with
    | None -> doc
    | Some selection -> applyRevisionDecisionWith true false { doc with Selection = Some selection }

let private revisionSortKey doc revision =
    let paragraphs = editableParagraphsInOrder doc
    let blockIndex = paragraphs |> List.tryFindIndex (fun para -> para.Id = revision.BlockId) |> Option.defaultValue System.Int32.MaxValue
    blockIndex, revision.Start

let private currentRevisionCursor useSelectionStart doc =
    let paragraphs = editableParagraphsInOrder doc
    match doc.Selection with
    | Some sel ->
        let pos =
            if sel.Anchor.BlockId = sel.Focus.BlockId then
                let offset = if useSelectionStart then min sel.Anchor.Offset sel.Focus.Offset else max sel.Anchor.Offset sel.Focus.Offset
                { sel.Focus with Offset = offset }
            elif useSelectionStart then
                let startPos, _, _ = getNormalizedRange sel doc
                startPos
            else
                let _, endPos, _ = getNormalizedRange sel doc
                endPos
        let blockIndex = paragraphs |> List.tryFindIndex (fun para -> para.Id = pos.BlockId) |> Option.defaultValue 0
        blockIndex, pos.Offset
    | None -> 0, 0

let private selectRevision revision doc =
    { doc with
        Selection = Some {
            Anchor = { BlockId = revision.BlockId; Offset = revision.Start }
            Focus = { BlockId = revision.BlockId; Offset = revision.End }
        } }

/// Select next revision, wrapping to the first revision when needed.
let selectNextRevision doc =
    let revisions = revisionInfos doc
    match revisions with
    | [] -> doc
    | _ ->
        let currentBlock, currentOffset = currentRevisionCursor false doc
        let next =
            revisions
            |> List.tryFind (fun revision ->
                let blockIndex, start = revisionSortKey doc revision
                blockIndex > currentBlock || (blockIndex = currentBlock && start > currentOffset))
            |> Option.defaultValue (List.head revisions)
        selectRevision next doc

/// Select previous revision, wrapping to the last revision when needed.
let selectPreviousRevision doc =
    let revisions = revisionInfos doc
    match revisions with
    | [] -> doc
    | _ ->
        let currentBlock, currentOffset = currentRevisionCursor true doc
        let previous =
            revisions
            |> List.filter (fun revision ->
                let blockIndex, start = revisionSortKey doc revision
                blockIndex < currentBlock || (blockIndex = currentBlock && start < currentOffset))
            |> List.tryLast
            |> Option.defaultValue (List.last revisions)
        selectRevision previous doc

/// Set font size on selection
let setFontSize size doc =
    applyFormattingToSelection (fun start len buf -> TextBuffer.setFontSize start len size buf) doc

/// Set font family on selection
let setFontFamily family doc =
    applyFormattingToSelection (fun start len buf -> TextBuffer.setFontFamily start len family buf) doc

/// Word's font sizes in points for Grow Font / Shrink Font (Ctrl+Shift+> / <):
/// those of its size box; above 72 pt steps of 10, below 8 pt steps of 1.
let private fontSizeSteps = [ 8.0; 9.0; 10.0; 10.5; 11.0; 12.0; 14.0; 16.0; 18.0; 20.0; 22.0; 24.0; 26.0; 28.0; 36.0; 48.0; 72.0 ]

/// The size in points a run of `points` gets from Grow Font (grow) or Shrink
/// Font; by one point (Ctrl+] / Ctrl+[) when not `stepwise`. From 1 to 1638 pt.
let nextFontSizePoints (grow: bool) (stepwise: bool) (points: float) : float =
    let points = System.Math.Round(points * 100.0) / 100.0
    let next =
        if not stepwise then (if grow then floor points + 1.0 else ceil points - 1.0)
        elif grow then
            match fontSizeSteps |> List.tryFind (fun size -> size > points) with
            | Some size -> size
            | None -> (floor (points / 10.0) + 1.0) * 10.0
        elif points > 72.0 then max 72.0 ((ceil (points / 10.0) - 1.0) * 10.0)
        elif points > 8.0 then fontSizeSteps |> List.filter (fun size -> size < points) |> List.last
        else ceil points - 1.0
    max 1.0 (min 1638.0 next)

/// Grow Font / Shrink Font on the selection: each run from the size it shows
/// (`effective`: its style resolved against its paragraph's, see
/// WordStyles.growFont) to the next size (nextFontSizePoints).
let resizeFontWith (effective: ParagraphBlock -> TextBuffer.TextStyle -> TextBuffer.TextStyle) (grow: bool) (stepwise: bool) doc =
    applyFormattingToSelectionInParagraph (fun para start len buffer ->
        TextBuffer.applyStyle start len (fun style ->
            let points = (effective para style).FontSize * 72.0 / 96.0
            TextBuffer.withExplicitFontSize (nextFontSizePoints grow stepwise points * 96.0 / 72.0) style) buffer) doc

/// Subscript (Ctrl+=) / superscript (Ctrl+Shift+=) switched on the selection,
/// as bold is: off when all of it is so already, on otherwise.
let toggleVertAlignWith (effective: ParagraphBlock -> TextBuffer.TextStyle -> TextBuffer.TextStyle) (align: string) doc =
    toggleOnSelection effective (fun style -> style.VertAlign = Some align) (fun on style -> { style with VertAlign = (if on then Some align else None) }) doc

/// Get selected text as plain text
let getSelectionText doc =
    match doc.Selection with
    | None -> ""
    | Some sel when sel.Anchor = sel.Focus -> ""  // No selection
    | Some sel ->
        let (startPos, endPos, blockIds) = getNormalizedRange sel doc
        let builder = StringBuilder()
        let mutable i = 0
        for blockId in blockIds do
            if i > 0 then
                builder.Append('\n') |> ignore
            match findParagraph doc blockId with
            | None -> ()
            | Some para ->
                let bufferLen = TextBuffer.length para.Buffer
                let sliceStart = if i = 0 then startPos.Offset else 0
                let sliceEnd = if i = blockIds.Length - 1 then endPos.Offset else bufferLen
                let sliceLen = sliceEnd - sliceStart
                if sliceLen > 0 then
                    builder.Append(TextBuffer.slice sliceStart sliceLen para.Buffer) |> ignore
            i <- i + 1
        builder.ToString()

// ---- copy and paste with formatting ---------------------------------------------
//
// What Copy keeps of a selection to paste it as it was: the selected part of
// each paragraph the selection touches, with its formatting, and the
// paragraph's properties (in Word, those of its paragraph mark - copied with
// every paragraph but the last, whose mark the selection does not reach).

/// A copied part of a paragraph.
type ClipParagraph = {
    /// The selected text, with its formatting and pictures.
    Buffer: TextBuffer.TextBuffer
    /// The paragraph's properties.
    Metadata: BlockMetadata
    /// The heading level of a paragraph of the body that is a heading.
    HeadingLevel: int option
    /// Not a paragraph but a table (pasted HTML): its rows, their cells, the
    /// paragraphs of each cell. Empty for a paragraph.
    TableRows: ClipParagraph list list list
}

/// What Copy and Cut keep of the selection (see copySelection).
type Clip = {
    Paragraphs: ClipParagraph list
    /// The selection's plain text, as getSelectionText gives it: what the
    /// system clipboard was given with the clip.
    Text: string
}

/// The style of copied text: its character formatting and its link; not the
/// comment, note, bookmark, field, revision of the run it was in, nor the
/// run of the file it was read from (pasted text is new text).
let private clipStyle (style: TextBuffer.TextStyle) : TextBuffer.TextStyle =
    let plain = TextBuffer.withoutAnnotations style
    { plain with Hyperlink = style.Hyperlink; InactiveHyperlink = style.InactiveHyperlink; SourceRun = None }

/// The part [start, finish) of a paragraph's text as copied: without the marks
/// of comments and the objects of the file kept as they were read (charts,
/// SmartArt - they are the file's parts, one place each).
let private clipBuffer (para: ParagraphBlock) (start: int) (finish: int) : TextBuffer.TextBuffer =
    if finish <= start then TextBuffer.empty
    else
        let sliced = TextBuffer.sliceBuffer start (finish - start) para.Buffer
        let cuts =
            [ for (markStart, markEnd) in commentMarkRanges para do
                let from = max markStart start
                let until = min markEnd finish
                if from < until then yield (from - start, until - from)
              for run in sliced.Runs do
                match run with
                | TextBuffer.Image image when image.Opaque.IsSome -> yield (image.Start, 1)
                | _ -> () ]
            |> List.sortByDescending fst
        let mutable buffer = sliced
        for (at, len) in cuts do
            buffer <- TextBuffer.delete at len buffer
        let len = TextBuffer.length buffer
        if len = 0 then buffer else TextBuffer.applyStyle 0 len clipStyle buffer

/// The paragraph properties a pasted paragraph mark brings: its formatting,
/// not what ties a paragraph to its place - the section it ends, the content
/// control it is in, a tracked change of its mark, where a page broke before it.
let private withPastedFormatting (pasted: BlockMetadata) (target: BlockMetadata) : BlockMetadata =
    { pasted with
        ContentControlId = target.ContentControlId
        ContentControlTag = target.ContentControlTag
        ContentControlAlias = target.ContentControlAlias
        ContentControlXml = target.ContentControlXml
        ContentControlGroup = target.ContentControlGroup
        ContentControlAncestors = target.ContentControlAncestors
        EnclosingCellControls = target.EnclosingCellControls
        SectionBreak = target.SectionBreak
        SectionColumns = target.SectionColumns
        SectionPageNumberStart = target.SectionPageNumberStart
        SectionPageNumberFormat = target.SectionPageNumberFormat
        SectionVerticalAlign = target.SectionVerticalAlign
        SectionPageSetup = target.SectionPageSetup
        HeaderReferences = target.HeaderReferences
        FooterReferences = target.FooterReferences
        RenderedPageBreakBefore = target.RenderedPageBreakBefore
        MarkRevision = target.MarkRevision }

/// The selection as Copy keeps it: None with nothing selected.
let copySelection (doc: Document) : Clip option =
    match doc.Selection with
    | Some sel when not sel.IsCollapsed ->
        let ranges = selectedParagraphRanges doc
        if ranges.IsEmpty then None
        else
            let headingLevel (id: BlockId) =
                match findBlock doc id with
                | Some (Heading (level, _)) -> Some level
                | _ -> None
            Some {
                Paragraphs =
                    ranges |> List.map (fun (para, start, finish) ->
                        { Buffer = clipBuffer para start finish; Metadata = para.Metadata; HeadingLevel = headingLevel para.Id; TableRows = [] })
                Text = getSelectionText doc }
    | _ -> None

/// A block of the body made a heading of this level, or a paragraph (None).
let private withHeadingLevel (id: BlockId) (level: int option) (doc: Document) : Document =
    let blocks =
        doc.Blocks |> mapListSharing (fun block ->
            match block, level with
            | Paragraph para, Some level when para.Id = id -> Heading (level, para)
            | Heading (current, para), Some level when para.Id = id && current <> level -> Heading (level, para)
            | Heading (_, para), None when para.Id = id -> Paragraph para
            | _ -> block)
    if sameReference blocks doc.Blocks then doc else { doc with Blocks = blocks }

/// A clip with each table (see ClipParagraph.TableRows) as paragraphs, one a
/// row, its cells parted by tabs (a cell's paragraphs by spaces): what is
/// pasted where a table cannot go - a table cell, a text box.
let clipWithoutTables (clip: Clip) : Clip =
    if clip.Paragraphs |> List.forall (fun piece -> piece.TableRows.IsEmpty) then clip
    else
        let separated (separator: string) (buffers: TextBuffer.TextBuffer list) =
            buffers |> List.reduce (fun left right -> TextBuffer.append (TextBuffer.append left (TextBuffer.ofString separator)) right)
        let cellBuffer (paragraphs: ClipParagraph list) =
            if paragraphs.IsEmpty then TextBuffer.empty else paragraphs |> List.map (fun para -> para.Buffer) |> separated " "
        { clip with
            Paragraphs =
                clip.Paragraphs |> List.collect (fun piece ->
                    if piece.TableRows.IsEmpty then [ piece ]
                    else
                        piece.TableRows |> List.map (fun cells ->
                            let buffer = if cells.IsEmpty then TextBuffer.empty else cells |> List.map cellBuffer |> separated "\t"
                            { Buffer = buffer; Metadata = defaultMetadata; HeadingLevel = None; TableRows = [] })) }

/// Paste what Copy kept at the caret, or over the selection (Word's "keep
/// source formatting"): the text with its formatting; with more than one
/// paragraph, each paragraph whose mark was copied with the properties it
/// had, and the last part joined to the text after the caret, in the
/// paragraph that keeps its own properties. The caret ends after the pasted
/// text. Works where Enter does - the body, a table cell, a text box, a
/// header (the paragraphs are split as Enter splits them).
let pasteClip (clip: Clip) (doc: Document) : Document =
    if clip.Paragraphs.IsEmpty then doc else
    let doc = snapSelectionToMarks None doc
    match doc.Selection with
    | None -> doc
    | Some sel ->
        let start = if sel.IsCollapsed then doc else deleteSelection doc
        // The block of the body (or of the story) a paragraph is, by its index.
        let topIndex (current: Document) (id: BlockId) =
            current.Blocks |> List.tryFindIndex (function
                | Paragraph para when para.Id = id -> true
                | Heading (_, para) when para.Id = id -> true
                | _ -> false)
        // A table goes between blocks: where the caret is not in one (a table
        // cell, a text box), its rows are paragraphs, the cells parted by tabs.
        let atTopLevel =
            match start.Selection with
            | Some caret -> (topIndex start caret.Focus.BlockId).IsSome
            | None -> false
        let pieces = if atTopLevel then clip.Paragraphs else (clipWithoutTables clip).Paragraphs
        let last = List.length pieces - 1
        let mutable current = start
        let mutable failed = false
        pieces |> List.iteri (fun index piece ->
            if not failed then
              if not piece.TableRows.IsEmpty then
                // A table: before the paragraph the caret is at the start of
                // (the paragraph split at the caret first), the caret where it was.
                match current.Selection with
                | Some caret when caret.IsCollapsed ->
                    let split = if caret.Focus.Offset = 0 then current else splitBlock current
                    match split.Selection with
                    | Some after when not (obj.ReferenceEquals(split, current) && caret.Focus.Offset > 0) ->
                        match topIndex split after.Focus.BlockId with
                        | Some at ->
                            let columns = piece.TableRows |> List.map List.length |> List.max
                            let table = newTable piece.TableRows.Length columns split
                            let filled =
                                { table with
                                    Rows =
                                        List.zip table.Rows piece.TableRows
                                        |> List.map (fun (row, cells) ->
                                            { row with
                                                Cells =
                                                    row.Cells |> List.mapi (fun column cell ->
                                                        match List.tryItem column cells with
                                                        | Some (first :: rest) ->
                                                            { cell with
                                                                Content =
                                                                    (first :: rest) |> List.map (fun para ->
                                                                        { Id = nextBlockId (); Buffer = para.Buffer; Metadata = withPastedFormatting para.Metadata defaultMetadata }) }
                                                        | _ -> cell) }) }
                            current <- { split with Blocks = List.take at split.Blocks @ (Table filled :: List.skip at split.Blocks) }
                        | None -> failed <- true
                    | _ -> failed <- true
                | _ -> failed <- true
              else
                match current.Selection with
                | Some caret when caret.IsCollapsed ->
                    match findParagraph current caret.Focus.BlockId with
                    | Some para ->
                        let len = TextBuffer.length para.Buffer
                        let at = max 0 (min caret.Focus.Offset len)
                        let buffer =
                            TextBuffer.append
                                (TextBuffer.append (TextBuffer.sliceBuffer 0 at para.Buffer) piece.Buffer)
                                (TextBuffer.sliceBuffer at (len - at) para.Buffer)
                        let past = { caret.Focus with Offset = at + TextBuffer.length piece.Buffer }
                        let inserted = { replaceParagraph current para.Id { para with Buffer = buffer } with Selection = Some { Anchor = past; Focus = past } }
                        if index = last then current <- inserted
                        else
                            // (Split as Enter splits: the part after the caret keeps the
                            // paragraph's own properties, the part before it is given
                            // those of the copied paragraph.)
                            let split = splitBlock inserted
                            if obj.ReferenceEquals(split, inserted) then failed <- true
                            else
                                match findParagraph split para.Id with
                                | Some before ->
                                    let metadata = withPastedFormatting piece.Metadata before.Metadata
                                    let withMetadata = replaceParagraph split para.Id { before with Metadata = metadata }
                                    current <-
                                        match findBlock withMetadata para.Id with
                                        | Some (Paragraph _)
                                        | Some (Heading _) -> withHeadingLevel para.Id piece.HeadingLevel withMetadata
                                        | _ -> withMetadata
                                | None -> failed <- true
                    | None -> failed <- true
                | _ -> failed <- true)
        // (No caret in a paragraph to paste at: the text alone, as insertText puts it.)
        if failed then insertText clip.Text doc
        else current

/// A clip as HTML, for the system clipboard (text/html): what other programs
/// - Word, a browser, a mail editor - paste with its formatting. Paragraphs
/// (headings as h1-h3, list items as li of ul / ol), their alignment, and the
/// character formatting the runs state themselves: bold, italic, underline,
/// strike, a size, a face, a colour, raised / lowered, a link. Pictures that
/// are data are img; others are left out.
let clipToHtml (clip: Clip) : string =
    let html = StringBuilder()
    let escape (text: string) = text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;")
    let points (pixels: float) = System.Math.Round(pixels * 72.0 / 96.0 * 100.0) / 100.0
    let runHtml (text: string) (style: TextBuffer.TextStyle) =
        let css = ResizeArray<string>()
        if style.Bold then css.Add "font-weight:bold"
        if style.Italic then css.Add "font-style:italic"
        let lines = [ if style.Underline then yield "underline"
                      if style.Strike || style.DoubleStrike then yield "line-through" ]
        if not lines.IsEmpty then css.Add ("text-decoration:" + String.concat " " lines)
        if TextBuffer.ExplicitFormat.has TextBuffer.ExplicitFormat.FontSize style then css.Add ("font-size:" + string (points style.FontSize) + "pt")
        style.FontFamily |> Option.iter (fun family -> css.Add (sprintf "font-family:'%s'" (family.Replace("'", ""))))
        style.Color |> Option.iter (fun color -> css.Add ("color:" + color))
        style.Highlight |> Option.iter (fun color -> css.Add ("background-color:" + color))
        match style.VertAlign with
        | Some "superscript" -> css.Add "vertical-align:super"
        | Some "subscript" -> css.Add "vertical-align:sub"
        | _ -> ()
        if style.AllCaps then css.Add "text-transform:uppercase"
        if style.SmallCaps then css.Add "font-variant:small-caps"
        let body = (escape text).Replace("\n", "<br>")
        let spanned = if css.Count = 0 then body else sprintf "<span style=\"%s\">%s</span>" (String.concat ";" css) body
        match style.Hyperlink with
        | Some target -> sprintf "<a href=\"%s\">%s</a>" (escape target) spanned
        | None -> spanned
    let mutable openList : string option = None
    let closeList () =
        match openList with
        | Some tag -> html.Append("</").Append(tag).Append(">") |> ignore; openList <- None
        | None -> ()
    for piece in clip.Paragraphs do
        let metadata = piece.Metadata
        let listTag = match metadata.ListStyle with Bullet -> Some "ul" | Numbered -> Some "ol" | _ -> None
        if listTag <> openList then
            closeList ()
            match listTag with
            | Some tag -> html.Append("<").Append(tag).Append(">") |> ignore; openList <- listTag
            | None -> ()
        let tag =
            match listTag, piece.HeadingLevel with
            | Some _, _ -> "li"
            | None, Some level -> sprintf "h%d" (max 1 (min 6 level))
            | None, None -> "p"
        let css = ResizeArray<string>()
        match metadata.Alignment with
        | Center -> css.Add "text-align:center"
        | Right -> css.Add "text-align:right"
        | Justify -> css.Add "text-align:justify"
        | _ -> ()
        if listTag.IsSome && metadata.ListLevel > 1 then css.Add (sprintf "margin-left:%dem" ((metadata.ListLevel - 1) * 2))
        html.Append("<").Append(tag) |> ignore
        if css.Count > 0 then html.Append(" style=\"").Append(String.concat ";" css).Append("\"") |> ignore
        html.Append(">") |> ignore
        let text = Rope.toString piece.Buffer.Rope
        for run in piece.Buffer.Runs do
            match run with
            | Text run when run.Length > 0 && not run.Style.Hidden ->
                html.Append(runHtml (text.Substring(run.Start, run.Length)) run.Style) |> ignore
            | Image image when image.Src.StartsWith "data:" ->
                html.Append("<img src=\"").Append(escape image.Src).Append("\"") |> ignore
                image.Width |> Option.iter (fun width -> html.Append(" width=\"" + string width + "\"") |> ignore)
                image.Height |> Option.iter (fun height -> html.Append(" height=\"" + string height + "\"") |> ignore)
                image.Alt |> Option.iter (fun alt -> html.Append(" alt=\"").Append(escape alt).Append("\"") |> ignore)
                html.Append(">") |> ignore
            | _ -> ()
        // (An empty paragraph keeps its line.)
        if text.Length = 0 then html.Append("<br>") |> ignore
        html.Append("</").Append(tag).Append(">") |> ignore
    closeList ()
    html.ToString()

/// HTML as a tree (what the browser's parser gives, see Input): text, and
/// elements with their tag (lower case), attributes and children.
type HtmlNode =
    | HtmlText of string
    | HtmlElement of tag: string * attributes: (string * string) list * children: HtmlNode list

let private htmlAttribute (name: string) (attributes: (string * string) list) =
    attributes |> List.tryPick (fun (key, value) -> if key = name then Some value else None)

/// The declarations of an element's style attribute: (property in lower case, value).
let private htmlCss (attributes: (string * string) list) : (string * string) list =
    match htmlAttribute "style" attributes with
    | None -> []
    | Some css ->
        css.Split(';')
        |> Array.choose (fun declaration ->
            let colon = declaration.IndexOf ':'
            if colon <= 0 then None
            else Some (declaration.Substring(0, colon).Trim().ToLowerInvariant(), declaration.Substring(colon + 1).Trim()))
        |> List.ofArray

let private parseNumber (text: string) : float option =
    match System.Double.TryParse(text.Trim()) with
    | true, value -> Some value
    | _ -> None

/// A CSS length in layout pixels (pt and px; a bare number is pixels).
let private htmlLength (value: string) : float option =
    let value = value.Trim().ToLowerInvariant()
    if value.EndsWith "pt" then parseNumber (value.Substring(0, value.Length - 2)) |> Option.map (fun points -> points * 96.0 / 72.0)
    elif value.EndsWith "px" then parseNumber (value.Substring(0, value.Length - 2))
    else parseNumber value

/// A CSS colour as "#RRGGBB" (hex and rgb() forms; named colours are left out).
let private htmlColor (value: string) : string option =
    let value = value.Trim().ToLowerInvariant()
    let digits = "0123456789ABCDEF"
    let hex (n: int) = let n = max 0 (min 255 n) in string digits.[n / 16] + string digits.[n % 16]
    if value.StartsWith "#" && value.Length = 7 then Some (value.ToUpperInvariant())
    elif value.StartsWith "#" && value.Length = 4 then
        Some (("#" + (value.Substring 1 |> Seq.map (fun c -> string c + string c) |> String.concat "")).ToUpperInvariant())
    elif value.StartsWith "rgb" then
        let inner = value.Substring(value.IndexOf '(' + 1).TrimEnd(')')
        let parts = inner.Split(',') |> Array.choose parseNumber
        if parts.Length >= 3 then Some ("#" + hex (int parts.[0]) + hex (int parts.[1]) + hex (int parts.[2])) else None
    else None

/// The character formatting an element gives the text in it.
let private htmlTextStyle (tag: string) (attributes: (string * string) list) (style: TextBuffer.TextStyle) : TextBuffer.TextStyle =
    let mutable style = style
    match tag with
    | "b" | "strong" -> style <- TextBuffer.withExplicitBold true style
    | "i" | "em" | "cite" | "var" | "dfn" -> style <- TextBuffer.withExplicitItalic true style
    | "u" | "ins" -> style <- TextBuffer.withExplicitUnderline true style
    | "s" | "strike" | "del" -> style <- TextBuffer.withExplicitStrike true style
    | "sup" -> style <- { style with VertAlign = Some "superscript" }
    | "sub" -> style <- { style with VertAlign = Some "subscript" }
    | "code" | "kbd" | "tt" | "samp" | "pre" -> style <- { style with FontFamily = Some "Courier New" }
    | "a" ->
        match htmlAttribute "href" attributes with
        | Some target when TextBuffer.isActiveHyperlinkTarget target -> style <- { style with Hyperlink = Some target }
        | _ -> ()
    | _ -> ()
    for (name, value) in htmlCss attributes do
        let value' = value.ToLowerInvariant()
        match name with
        | "font-weight" ->
            match value' with
            | "bold" | "bolder" -> style <- TextBuffer.withExplicitBold true style
            | "normal" | "lighter" -> style <- TextBuffer.withExplicitBold false style
            | _ ->
                match parseNumber value' with
                | Some weight -> style <- TextBuffer.withExplicitBold (weight >= 600.0) style
                | None -> ()
        | "font-style" -> style <- TextBuffer.withExplicitItalic (value' = "italic" || value' = "oblique") style
        | "text-decoration" | "text-decoration-line" ->
            if value'.Contains "none" then
                style <- style |> TextBuffer.withExplicitUnderline false |> TextBuffer.withExplicitStrike false
            else
                if value'.Contains "underline" then style <- TextBuffer.withExplicitUnderline true style
                if value'.Contains "line-through" then style <- TextBuffer.withExplicitStrike true style
        | "font-size" ->
            match htmlLength value with
            | Some size when size > 0.0 -> style <- TextBuffer.withExplicitFontSize size style
            | _ -> ()
        | "font-family" ->
            let first = value.Split(',').[0].Trim().Trim('"', '\'').Trim()
            match first.ToLowerInvariant() with
            | "" | "serif" | "sans-serif" | "monospace" | "cursive" | "fantasy" | "system-ui" | "inherit" | "initial" -> ()
            | _ -> style <- { style with FontFamily = Some first; Explicit = style.Explicit ||| TextBuffer.ExplicitFormat.FontFamily }
        | "color" ->
            match htmlColor value with
            | Some color -> style <- { style with Color = Some color }
            | None -> ()
        | "vertical-align" ->
            match value' with
            | "super" -> style <- { style with VertAlign = Some "superscript" }
            | "sub" -> style <- { style with VertAlign = Some "subscript" }
            | "baseline" -> style <- { style with VertAlign = None }
            | _ -> ()
        | _ -> ()
    style

let private htmlAlignment (attributes: (string * string) list) : Alignment option =
    let value =
        match htmlCss attributes |> List.tryPick (fun (name, value) -> if name = "text-align" then Some value else None) with
        | Some value -> Some value
        | None -> htmlAttribute "align" attributes
    match value |> Option.map (fun v -> v.Trim().ToLowerInvariant()) with
    | Some "center" -> Some Center
    | Some "right" | Some "end" -> Some Right
    | Some "justify" -> Some Justify
    | Some "left" | Some "start" -> Some Left
    | _ -> None

/// Pasted HTML (text/html of the clipboard: from Word, a browser, a mail) as
/// a clip (see pasteClip): a paragraph of each block - p, a heading (h1-h3 a
/// heading of that level, h4-h6 of level 3), a list item (ul: bullets, ol:
/// numbers; nested lists a level down; Word's own list paragraphs, mso-list,
/// too) -, a table (its rows and cells, the paragraphs of each cell; a table
/// in a cell is its rows, the cells parted by tabs), br a line break, the
/// text's character formatting (see htmlTextStyle), pictures that are data.
/// Spaces are collapsed as a browser shows them. None when there is no text.
let rec private clipPiecesOfHtml (allowTables: bool) (baseStyle: TextBuffer.TextStyle) (nodes: HtmlNode list) : ClipParagraph list =
    let paragraphs = ResizeArray<ClipParagraph>()
    let mutable buffer = TextBuffer.empty
    let mutable metadata = defaultMetadata
    let mutable heading : int option = None
    let mutable preformatted = false
    let lastChar () =
        let len = TextBuffer.length buffer
        if len = 0 then None else Some ((TextBuffer.slice (len - 1) 1 buffer).[0])
    let flush (always: bool) =
        // (Spaces at the end of a paragraph are not shown; Word's empty paragraph is a no-break space.)
        let text = TextBuffer.toString buffer
        let trimmed = text.TrimEnd(' ')
        if trimmed.Length < text.Length then buffer <- TextBuffer.delete trimmed.Length (text.Length - trimmed.Length) buffer
        if trimmed = " " then buffer <- TextBuffer.empty
        if always || TextBuffer.length buffer > 0 then
            paragraphs.Add { Buffer = buffer; Metadata = metadata; HeadingLevel = heading; TableRows = [] }
        buffer <- TextBuffer.empty
        heading <- None
    let append (text: string) (style: TextBuffer.TextStyle) =
        if text <> "" then
            let at = TextBuffer.length buffer
            buffer <- TextBuffer.insert at text buffer |> TextBuffer.applyStyle at text.Length (fun _ -> style)
    let appendText (text: string) (style: TextBuffer.TextStyle) =
        if preformatted then append (text.Replace("\r\n", "\n")) style
        else
            let collapsed = System.Text.RegularExpressions.Regex.Replace(text, "[ \t\r\n\f]+", " ")
            // (A space at the start of a line, or after a space, is not shown.)
            let collapsed =
                match lastChar () with
                | None | Some ' ' | Some '\n' -> collapsed.TrimStart(' ')
                | _ -> collapsed
            append collapsed style
    let rec textOf (node: HtmlNode) =
        match node with
        | HtmlText text -> text
        | HtmlElement (_, _, children) -> children |> List.map textOf |> String.concat ""
    let blockTags = set [ "p"; "div"; "h1"; "h2"; "h3"; "h4"; "h5"; "h6"; "li"; "blockquote"; "pre"; "address"; "center"; "dt"; "dd"
                          "tr"; "caption"; "figcaption"; "section"; "article"; "header"; "footer"; "main"; "nav"; "aside"; "table"
                          "tbody"; "thead"; "tfoot"; "ul"; "ol"; "dl"; "form"; "fieldset"; "body"; "html"; "figure" ]
    let paragraphTags = set [ "p"; "h1"; "h2"; "h3"; "h4"; "h5"; "h6"; "li"; "pre"; "dt"; "dd"; "tr"; "caption"; "figcaption"; "address" ]
    let skipped = set [ "head"; "style"; "script"; "title"; "meta"; "link"; "template"; "noscript"; "xml"; "button"; "select"; "textarea"; "input" ]
    // lists: the kind of each list the walk is in, innermost first.
    let rec walk (lists: ListStyle list) (alignment: Alignment option) (style: TextBuffer.TextStyle) (node: HtmlNode) =
        match node with
        | HtmlText text -> appendText text style
        | HtmlElement (tag, attributes, children) ->
            let css = htmlCss attributes
            let msoList = css |> List.tryPick (fun (name, value) -> if name = "mso-list" then Some value else None)
            if skipped.Contains tag || (msoList |> Option.exists (fun value -> value.Trim().ToLowerInvariant() = "ignore")) then ()
            elif tag = "br" then append "\n" style
            elif tag = "img" then
                match htmlAttribute "src" attributes with
                | Some src when src.StartsWith "data:image/" ->
                    let size name = htmlAttribute name attributes |> Option.bind htmlLength
                    buffer <- TextBuffer.insertImage (TextBuffer.length buffer) src (size "width") (size "height") (htmlAttribute "alt" attributes) buffer
                | _ -> ()
            elif tag = "table" && allowTables && not (tableRows children).IsEmpty then
                flush false
                let tableStyle = htmlTextStyle tag attributes style
                let rows =
                    tableRows children |> List.map (fun (rowAttributes, cells) ->
                        let rowStyle = htmlTextStyle "tr" rowAttributes tableStyle
                        cells |> List.map (fun (cellTag, cellAttributes, cellChildren) ->
                            let cellStyle = htmlTextStyle cellTag cellAttributes rowStyle
                            let cellStyle = if cellTag = "th" then TextBuffer.withExplicitBold true cellStyle else cellStyle
                            let content = clipPiecesOfHtml false cellStyle cellChildren
                            match htmlAlignment cellAttributes with
                            | Some cellAlignment ->
                                content |> List.map (fun para ->
                                    if para.Metadata.ExplicitAlignment then para
                                    else { para with Metadata = { para.Metadata with Alignment = cellAlignment; ExplicitAlignment = true } })
                            | None -> content))
                    |> List.filter (fun cells -> not cells.IsEmpty)
                paragraphs.Add { Buffer = TextBuffer.empty; Metadata = defaultMetadata; HeadingLevel = None; TableRows = rows }
            elif tag = "td" || tag = "th" then
                if TextBuffer.length buffer > 0 then append "\t" style
                children |> List.iter (walk lists alignment (htmlTextStyle tag attributes style))
            elif blockTags.Contains tag then
                let alignment = htmlAlignment attributes |> Option.orElse alignment
                let lists =
                    match tag with
                    | "ul" -> Bullet :: lists
                    | "ol" -> Numbered :: lists
                    | _ -> lists
                if paragraphTags.Contains tag || msoList.IsSome then
                    flush false
                    let listed =
                        match tag, msoList with
                        | _, Some value ->
                            // (Word's list paragraph: "l0 level2 lfo1"; its number or bullet is
                            // the text of the span marked mso-list:Ignore.)
                            let level =
                                let found = System.Text.RegularExpressions.Regex.Match(value, "level(\\d)")
                                if found.Success then int found.Groups.[1].Value else 1
                            let rec ignored (node: HtmlNode) =
                                match node with
                                | HtmlElement (_, attrs, kids) ->
                                    if htmlCss attrs |> List.exists (fun (name, v) -> name = "mso-list" && v.Trim().ToLowerInvariant() = "ignore") then Some (textOf node)
                                    else kids |> List.tryPick ignored
                                | HtmlText _ -> None
                            let marker = children |> List.tryPick ignored |> Option.defaultValue ""
                            let numbered = System.Text.RegularExpressions.Regex.IsMatch(marker.Trim(), "^[0-9A-Za-z]+[.)]")
                            Some ((if numbered then Numbered else Bullet), level)
                        | "li", _ -> Some ((match lists with kind :: _ -> kind | [] -> Bullet), max 1 lists.Length)
                        | _ -> None
                    metadata <-
                        { defaultMetadata with
                            Alignment = alignment |> Option.defaultValue Left
                            ExplicitAlignment = alignment.IsSome
                            ListStyle = (match listed with Some (kind, _) -> kind | None -> NoList)
                            ListLevel = (match listed with Some (_, level) -> max 1 (min 9 level) | None -> 0) }
                    heading <-
                        match tag with
                        | "h1" -> Some 1
                        | "h2" -> Some 2
                        | "h3" | "h4" | "h5" | "h6" -> Some 3
                        | _ -> None
                    let wasPreformatted = preformatted
                    preformatted <- preformatted || tag = "pre"
                    let before = paragraphs.Count
                    children |> List.iter (walk lists alignment (htmlTextStyle tag attributes style))
                    preformatted <- wasPreformatted
                    // (An empty one is an empty paragraph - unless a block in it made its paragraphs.)
                    flush (paragraphs.Count = before)
                    metadata <- defaultMetadata
                else
                    flush false
                    children |> List.iter (walk lists alignment (htmlTextStyle tag attributes style))
                    flush false
            else children |> List.iter (walk lists alignment (htmlTextStyle tag attributes style))
    nodes |> List.iter (walk [] None baseStyle)
    flush false
    List.ofSeq paragraphs

/// The rows of a table element's children (in thead, tbody, tfoot or not;
/// not those of a table in a cell): each row's attributes and its cells (tag,
/// attributes, children).
and private tableRows (children: HtmlNode list) : ((string * string) list * (string * (string * string) list * HtmlNode list) list) list =
    children |> List.collect (fun node ->
        match node with
        | HtmlElement ("tr", attributes, cells) ->
            [ attributes, cells |> List.choose (function
                                    | HtmlElement (("td" | "th") as tag, cellAttributes, cellChildren) -> Some (tag, cellAttributes, cellChildren)
                                    | _ -> None) ]
        | HtmlElement (("thead" | "tbody" | "tfoot"), _, rows) -> tableRows rows
        | _ -> [])

/// Pasted HTML as a clip (see clipPiecesOfHtml and pasteClip). None when it has no text.
let clipOfHtml (nodes: HtmlNode list) : Clip option =
    let pieces = clipPiecesOfHtml true TextBuffer.defaultStyle nodes
    if pieces.IsEmpty then None
    else
        let textOf (piece: ClipParagraph) =
            if piece.TableRows.IsEmpty then TextBuffer.toString piece.Buffer
            else
                piece.TableRows
                |> List.map (fun cells -> cells |> List.map (fun paragraphs -> paragraphs |> List.map (fun para -> TextBuffer.toString para.Buffer) |> String.concat " ") |> String.concat "\t")
                |> String.concat "\n"
        Some { Paragraphs = pieces; Text = pieces |> List.map textOf |> String.concat "\n" }

let private toggleListStyle listStyle doc =
    // Word's rule: if every selected paragraph is already in this kind of
    // list, take them all out; otherwise put them all in.
    let paragraphs = selectedParagraphs doc
    let allInList = not paragraphs.IsEmpty && paragraphs |> List.forall (fun para -> para.Metadata.ListStyle = listStyle)
    mapSelectedMetadata (fun metadata ->
        if allInList then { metadata with ListStyle = NoList; ListLevel = 0 }
        elif metadata.ListStyle = listStyle then metadata
        else { metadata with ListStyle = listStyle; ListLevel = max 1 metadata.ListLevel }) doc

/// Toggle bullet list on the selected paragraphs
let toggleBulletList doc =
    toggleListStyle Bullet doc

/// Word's Ctrl+Shift+L (the List Bullet style): the selected paragraphs are
/// bulleted list items; they stay so when they are already.
let applyBulletList doc =
    mapSelectedMetadata (fun metadata ->
        if metadata.ListStyle = Bullet then metadata
        else { metadata with ListStyle = Bullet; ListLevel = max 1 metadata.ListLevel }) doc

/// Toggle numbered list on the selected paragraphs
let toggleNumberedList doc =
    toggleListStyle Numbered doc

/// Apply a numbering marker format to the selected paragraphs and ensure they are a numbered list.
let setNumberFormat format doc =
    let normalizedFormat =
        match format with
        | "lowerLetter" | "upperLetter" | "lowerRoman" | "upperRoman" -> format
        | _ -> "decimal"
    mapSelectedMetadata (fun metadata ->
        { metadata with
            ListStyle = Numbered
            ListLevel = max 1 metadata.ListLevel
            ListNumberFormat = Some normalizedFormat
            ListLevelText = None }) doc

/// A paragraph of a list at another level of it (1-9). A list that came from
/// a file has its levels (BlockMetadata.ListLevels): the paragraph gets the
/// level's number format, text and start, and the level's indents where it
/// states none of its own - what it would have had if the file had said that
/// level. A list made in the editor has the editor's own levels.
let withListLevel (level: int) (metadata: BlockMetadata) : BlockMetadata =
    let level = max 1 (min 9 level)
    match metadata.ListLevels |> List.tryFind (fun definition -> definition.Level = level) with
    | None -> { metadata with ListLevel = level }
    | Some definition ->
        let previous = metadata.ListLevelIndents
        // (An indent the paragraph has from the level it leaves is the new level's; its own stays.)
        let indent (flag: int) (current: float option) (fromOld: float option) (fromNew: float option) =
            if ParagraphExplicit.has flag metadata then current
            else
                match fromNew with
                | Some _ -> fromNew
                | None -> (if current = fromOld then None else current)
        { metadata with
            ListLevel = level
            ListStyle = (if definition.NumberFormat = "bullet" then Bullet else Numbered)
            ListNumberFormat = Some definition.NumberFormat
            ListLevelText = definition.LevelText
            ListStart = definition.Start
            ListLevelIndents = Some (definition.IndentLeft, definition.FirstLineIndent)
            IndentLeft = indent ParagraphExplicit.IndentLeft metadata.IndentLeft (previous |> Option.bind fst) definition.IndentLeft
            FirstLineIndent = indent ParagraphExplicit.FirstLineIndent metadata.FirstLineIndent (previous |> Option.bind snd) definition.FirstLineIndent }

/// Partial list formatting and numbering operation from the shell dialog.
type ListDialogChange = {
    Kind: ListStyle option
    Marker: string option
    Level: int option
    Sequence: string
    Start: int
}

/// Allocate distinct unused numeric ids (including ids in preserved XML).
let private applyListDialogCore (allocateId: unit -> string) (change: ListDialogChange) doc =
    let marker (m: BlockMetadata) =
        match m.ListStyle with
        | Bullet -> m.ListLevelText |> Option.defaultValue "•"
        | Numbered -> m.ListNumberFormat |> Option.defaultValue "decimal"
        | NoList -> ""
    let changesMarker (m: BlockMetadata) =
        let kind = change.Kind |> Option.defaultValue m.ListStyle
        kind <> m.ListStyle || (kind <> NoList && (change.Marker |> Option.exists (fun value -> value <> marker m)))
    // A DOCX numbering instance owns its counter. Split the affected suffix
    // into physical instances with counter snapshots, so the selected marker
    // can change without renumbering any unselected item (including parents
    // referenced by a later nested item).
    let targets = selectedParagraphs doc
    let selectedIds = HashSet<BlockId>(targets |> List.map (fun p -> p.Id))
    let changedIds =
        targets |> List.filter (fun p -> p.Metadata.ListStyle = Numbered && changesMarker p.Metadata)
        |> List.map (fun p -> p.Metadata.ListId) |> Set.ofList
    let segmented =
        if changedIds.IsEmpty then doc
        else
            let story =
                match targets with
                | first :: _ ->
                    match textBoxOfParagraph doc first.Id with
                    | Some box -> box.Content
                    | None -> bodyParagraphsInOrder doc
                | [] -> []
            let mutable counters = Map.empty
            let mutable active = Set.empty
            let replacements = Dictionary<BlockId, BlockMetadata>()
            for paragraph in story do
                let m = paragraph.Metadata
                counters <- advanceListCounters m counters
                if m.ListStyle = Numbered && Set.contains m.ListId changedIds then
                    if selectedIds.Contains paragraph.Id && changesMarker m then active <- Set.add m.ListId active
                    if Set.contains m.ListId active then
                        let key = m.ListId |> Option.defaultValue ""
                        let snapshots =
                            [ for level in 1 .. 9 do
                                match counters |> Map.tryFind (key, level) with
                                | Some value -> yield level, value
                                | None -> () ]
                        let levels =
                            m.ListLevels |> List.map (fun definition ->
                                match snapshots |> List.tryFind (fun (level, _) -> level = definition.Level) with
                                | Some (_, value) -> { definition with Start = Some value }
                                | None -> definition)
                        let start = snapshots |> List.tryFind (fun (level, _) -> level = m.ListLevel) |> Option.map snd
                        replacements.[paragraph.Id] <-
                            { m with ListId = Some (allocateId ()); ListSourceId = m.ListSourceId |> Option.orElse m.ListId
                                     ListRestart = true; ListStart = start; ListStartOverrides = snapshots; ListLevels = levels }
            mapEditableParagraphs (fun paragraph ->
                match replacements.TryGetValue paragraph.Id with
                | true, metadata -> { paragraph with Metadata = metadata }
                | _ -> paragraph) doc
    let ids = System.Collections.Generic.Dictionary<ListStyle * string option, string>()
    let update (m: BlockMetadata) =
        let kind = change.Kind |> Option.defaultValue m.ListStyle
        if kind = NoList then
            if m.ListStyle = NoList then m
            else { m with ListStyle = NoList; ListLevel = 0; ListId = Some "0"; ListSourceId = None; ListRestart = false; ListFormatEdited = false; ListStartOverrides = [] }
        else
            let changed = kind <> m.ListStyle || (change.Marker |> Option.exists (fun value -> value <> marker m))
            let baseMeta =
                if not changed then m
                else
                    let key = m.ListStyle, (if m.ListStyle = NoList then None else m.ListId)
                    let id =
                        match ids.TryGetValue key with
                        | true, id -> id
                        | _ -> let id = allocateId () in ids.[key] <- id; id
                    let source = if m.ListStyle = NoList then None else m.ListSourceId |> Option.orElse m.ListId
                    let chosen = change.Marker |> Option.defaultValue (if kind = Bullet then "•" else "decimal")
                    let numberFormat = if kind = Bullet then "bullet" else chosen
                    let template = if kind = Bullet then Some chosen elif m.ListStyle = Numbered then m.ListLevelText else None
                    { m with ListStyle = kind; ListLevel = max 1 m.ListLevel; ListId = Some id
                             ListSourceId = source; ListRestart = (if m.ListStyle = NoList then false else m.ListRestart); ListFormatEdited = true
                             ListStart = if m.ListStyle = NoList then None else m.ListStart
                             ListStartOverrides = if m.ListStyle = NoList then [] else m.ListStartOverrides
                             ListLevels = if m.ListStyle = NoList then [] else m.ListLevels
                             ListLevelIndents = if m.ListStyle = NoList then None else m.ListLevelIndents
                             ListNumberFormat = Some numberFormat
                             ListLevelText = template }
            let moved =
                match change.Level with
                | Some level when level <> baseMeta.ListLevel ->
                    let result = withListLevel level baseMeta
                    if changed then
                        { result with ListStyle = baseMeta.ListStyle; ListNumberFormat = baseMeta.ListNumberFormat
                                      ListLevelText = if kind = Bullet then Some (change.Marker |> Option.defaultValue "•") elif m.ListStyle = Numbered && result.ListStyle = Numbered then result.ListLevelText else None }
                    else result
                | _ -> baseMeta
            if not changed then moved
            else
                let levels =
                    moved.ListLevels |> List.map (fun definition ->
                        if definition.Level <> moved.ListLevel then definition
                        else
                            { definition with
                                NumberFormat = moved.ListNumberFormat |> Option.defaultValue "decimal"
                                LevelText = moved.ListLevelText
                                MarkFont = if kind = Bullet then None else definition.MarkFont })
                { moved with ListLevels = levels }
    let formatted = mapSelectedMetadata update segmented
    if change.Sequence = "keep" then Ok formatted
    else
        let targets = selectedParagraphs formatted
        match targets with
        | [] -> Error "Select a numbered list first."
        | first :: _ when first.Metadata.ListStyle <> Numbered || (targets |> List.exists (fun p -> p.Metadata.ListStyle = NoList || p.Metadata.ListId <> first.Metadata.ListId)) ->
            Error "Restart or continue requires one numbered list."
        | first :: _ ->
            let story = match textBoxOfParagraph formatted first.Id with Some box -> box.Content | None -> bodyParagraphsInOrder formatted
            match story |> List.tryFindIndex (fun p -> p.Id = first.Id) with
            | None -> Error "The selected list is not in an editable story."
            | Some index ->
                let tail = story |> List.skip index |> List.takeWhile (fun p -> p.Metadata.ListStyle <> NoList && p.Metadata.ListId = first.Metadata.ListId)
                let tailIds = HashSet<BlockId>(tail |> List.map (fun p -> p.Id))
                if targets |> List.exists (fun p -> not (tailIds.Contains p.Id)) then
                    Error "Restart or continue requires a contiguous list selection."
                else
                    let replacement =
                        if change.Sequence = "restart" then
                            let m = first.Metadata
                            Ok { m with ListId = Some (allocateId ()); ListSourceId = m.ListSourceId |> Option.orElse m.ListId; ListRestart = true; ListStart = Some change.Start }
                        else
                            match story |> List.take index |> List.rev |> List.tryFind (fun p -> p.Metadata.ListStyle = Numbered && p.Metadata.ListLevel = first.Metadata.ListLevel) with
                            | None -> Error "There is no preceding numbered list at this level."
                            | Some prior -> Ok prior.Metadata
                    match replacement with
                    | Error message -> Error message
                    | Ok replacement when change.Sequence = "continue" && replacement.ListId = first.Metadata.ListId -> Ok formatted
                    | Ok replacement ->
                        let mapper (p: ParagraphBlock) =
                            if not (tailIds.Contains p.Id) then p
                            else
                                let m = p.Metadata
                                let result =
                                    if change.Sequence = "restart" then
                                        { m with ListId = replacement.ListId; ListSourceId = replacement.ListSourceId
                                                 ListRestart = m.ListLevel = first.Metadata.ListLevel
                                                 ListStart = if m.ListLevel = first.Metadata.ListLevel then Some change.Start else m.ListStart
                                                 ListStartOverrides = if m.ListLevel = first.Metadata.ListLevel then (m.ListLevel, change.Start) :: (m.ListStartOverrides |> List.filter (fun (level, _) -> level <> m.ListLevel)) else m.ListStartOverrides }
                                    else
                                        let template = withListLevel m.ListLevel replacement
                                        { m with ListId = replacement.ListId; ListSourceId = replacement.ListSourceId
                                                 ListRestart = replacement.ListRestart && m.ListLevel = replacement.ListLevel
                                                 ListFormatEdited = replacement.ListFormatEdited && m.ListLevel = replacement.ListLevel
                                                 ListStartOverrides = replacement.ListStartOverrides
                                                 ListStyle = template.ListStyle; ListLevels = replacement.ListLevels; ListNumberFormat = template.ListNumberFormat
                                                 ListLevelText = template.ListLevelText; ListStart = template.ListStart }
                                if result = m then p else { p with Metadata = result }
                        Ok (mapEditableParagraphs mapper formatted)
/// Apply sequence and formatting together as one document transaction.
let applyListDialog (allocateId: unit -> string) (change: ListDialogChange) doc =
    if change.Sequence = "keep" then applyListDialogCore allocateId change doc
    elif change.Kind = Some Bullet || change.Kind = Some NoList then Error "Restart or continue requires a numbered list."
    else
        let prepare =
            match change.Kind with
            | Some Numbered -> applyListDialogCore allocateId { change with Marker = None; Level = None; Sequence = "keep" } doc
            | _ -> Ok doc
        prepare
        |> Result.bind (fun numbered ->
            applyListDialogCore allocateId { change with Kind = None; Marker = None; Level = None } numbered)
        |> Result.bind (fun sequenced ->
            applyListDialogCore allocateId { change with Kind = None; Sequence = "keep" } sequenced)

/// Increase list indentation (Tab) of the selected paragraphs
let increaseIndent doc =
    mapSelectedMetadata (fun metadata ->
        // Only increase if in a list and not at max level
        if metadata.ListStyle <> NoList && metadata.ListLevel < 9 then
            withListLevel (metadata.ListLevel + 1) metadata
        else
            metadata) doc

/// Decrease list indentation (Shift+Tab) of the selected paragraphs
let decreaseIndent doc =
    mapSelectedMetadata (fun metadata ->
        // Only decrease if in a list
        if metadata.ListStyle <> NoList && metadata.ListLevel > 0 then
            let newLevel = metadata.ListLevel - 1
            if newLevel = 0 then
                // Exit list when reaching level 0
                { metadata with ListStyle = NoList; ListLevel = 0 }
            else
                withListLevel newLevel metadata
        else
            metadata) doc

/// Put the selected paragraphs that are in a list at a level of it (1-9).
let setListLevel (level: int) doc =
    mapSelectedMetadata (fun metadata ->
        if metadata.ListStyle <> NoList then withListLevel level metadata else metadata) doc

let private setAlignment alignment doc =
    mapSelectedMetadata (fun metadata -> { metadata with Alignment = alignment; ExplicitAlignment = true }) doc

/// Set alignment to left
let alignLeft doc =
    setAlignment Left doc

/// Set alignment to center
let alignCenter doc =
    setAlignment Center doc

/// Set alignment to right
let alignRight doc =
    setAlignment Right doc

/// Set alignment to justify
let alignJustify doc =
    setAlignment Justify doc

/// Increase Indent (Ctrl+M) / Decrease Indent (Ctrl+Shift+M) of the selected
/// paragraphs, as Word's: a paragraph of a list a level down / up (not out of
/// the list), any other the left indent to the next / previous half inch.
let indentParagraphs (forward: bool) doc =
    let step = 48.0
    mapSelectedMetadata (fun metadata ->
        if metadata.ListStyle <> NoList then
            let level = if forward then min 9 (metadata.ListLevel + 1) else max 1 (metadata.ListLevel - 1)
            if level = metadata.ListLevel then metadata else withListLevel level metadata
        else
            let current = metadata.IndentLeft |> Option.defaultValue 0.0
            let next =
                if forward then (floor (current / step + 0.001) + 1.0) * step
                else max 0.0 ((ceil (current / step - 0.001) - 1.0) * step)
            if Some next = metadata.IndentLeft || (next = 0.0 && metadata.IndentLeft.IsNone) then metadata
            else { metadata with IndentLeft = Some next; Explicit = metadata.Explicit ||| ParagraphExplicit.IndentLeft }) doc

let setCurrentParagraphStyle alignment indentLeft indentRight spacingBefore spacingAfter lineSpacing firstLineIndent backgroundColor doc =
    mapSelectedMetadata (fun metadata ->
        { metadata with
            Alignment = alignment
            ExplicitAlignment = true
            IndentLeft = indentLeft
            IndentRight = indentRight
            SpacingBefore = spacingBefore
            SpacingAfter = spacingAfter
            LineSpacing = lineSpacing
            LineSpacingRule = Some "auto"
            FirstLineIndent = firstLineIndent
            BackgroundColor = backgroundColor
            // What the dialog states is the paragraph's own formatting; a value it
            // leaves empty goes back to being inherited.
            Explicit =
                let flag (isSet: bool) (bit: int) (flags: int) = if isSet then flags ||| bit else flags &&& ~~~bit
                metadata.Explicit
                |> flag indentLeft.IsSome ParagraphExplicit.IndentLeft
                |> flag indentRight.IsSome ParagraphExplicit.IndentRight
                |> flag spacingBefore.IsSome ParagraphExplicit.SpacingBefore
                |> flag spacingAfter.IsSome ParagraphExplicit.SpacingAfter
                |> flag lineSpacing.IsSome ParagraphExplicit.LineSpacing
                |> flag firstLineIndent.IsSome ParagraphExplicit.FirstLineIndent
                |> flag backgroundColor.IsSome ParagraphExplicit.BackgroundColor }) doc

/// Update only the paragraph properties supplied by a dialog. Missing values
/// leave each selected paragraph's own metadata untouched.
let setSelectedParagraphStyleChanges alignment indentLeft indentRight spacingBefore spacingAfter lineSpacing firstLineIndent backgroundColor doc =
    let update metadata =
        let changed = function Some _ -> true | None -> false
        let choose value current = value |> Option.defaultValue current
        let explicit flag value flags =
            if changed value then
                match value with
                | Some (Some _) -> flags ||| flag
                | Some None -> flags &&& ~~~flag
                | None -> flags
            else flags
        { metadata with
            Alignment = choose alignment metadata.Alignment
            ExplicitAlignment = if alignment.IsSome then true else metadata.ExplicitAlignment
            IndentLeft = choose indentLeft metadata.IndentLeft
            IndentRight = choose indentRight metadata.IndentRight
            SpacingBefore = choose spacingBefore metadata.SpacingBefore
            SpacingAfter = choose spacingAfter metadata.SpacingAfter
            LineSpacing = choose lineSpacing metadata.LineSpacing
            LineSpacingRule = if lineSpacing.IsSome then (if lineSpacing.Value.IsSome then Some "auto" else None) else metadata.LineSpacingRule
            FirstLineIndent = choose firstLineIndent metadata.FirstLineIndent
            BackgroundColor = choose backgroundColor metadata.BackgroundColor
            Explicit =
                metadata.Explicit
                |> explicit ParagraphExplicit.IndentLeft indentLeft
                |> explicit ParagraphExplicit.IndentRight indentRight
                |> explicit ParagraphExplicit.SpacingBefore spacingBefore
                |> explicit ParagraphExplicit.SpacingAfter spacingAfter
                |> explicit ParagraphExplicit.LineSpacing lineSpacing
                |> explicit ParagraphExplicit.FirstLineIndent firstLineIndent
                |> explicit ParagraphExplicit.BackgroundColor backgroundColor }
    mapSelectedMetadata update doc

/// Set the line spacing of the selected paragraphs as a multiple of single
/// spacing (1.0, 1.15, 1.5, 2.0 ...); None goes back to the style's.
let setLineSpacing (lineSpacing: float option) doc =
    mapSelectedMetadata (fun metadata ->
        { metadata with
            LineSpacing = lineSpacing
            LineSpacingRule = (if lineSpacing.IsSome then Some "auto" else None)
            Explicit = (if lineSpacing.IsSome then metadata.Explicit ||| ParagraphExplicit.LineSpacing else metadata.Explicit &&& ~~~ParagraphExplicit.LineSpacing) }) doc

/// Set the space before the selected paragraphs (layout pixels); None goes back to the style's.
let setSpacingBefore (spacing: float option) doc =
    mapSelectedMetadata (fun metadata ->
        { metadata with
            SpacingBefore = spacing
            Explicit = (if spacing.IsSome then metadata.Explicit ||| ParagraphExplicit.SpacingBefore else metadata.Explicit &&& ~~~ParagraphExplicit.SpacingBefore) }) doc

/// Set the space after the selected paragraphs (layout pixels); None goes back to the style's.
let setSpacingAfter (spacing: float option) doc =
    mapSelectedMetadata (fun metadata ->
        { metadata with
            SpacingAfter = spacing
            Explicit = (if spacing.IsSome then metadata.Explicit ||| ParagraphExplicit.SpacingAfter else metadata.Explicit &&& ~~~ParagraphExplicit.SpacingAfter) }) doc

/// Insert a page break at the caret: the paragraph is split there and what
/// follows starts a new page.
let insertPageBreak doc =
    let split = splitBlock doc
    match split.Selection with
    | Some selection when not (obj.ReferenceEquals(split, doc)) ->
        match findParagraph split selection.Focus.BlockId with
        | Some para -> replaceParagraph split para.Id { para with Metadata = { para.Metadata with PageBreakBefore = true } }
        | None -> split
    | _ -> split

/// Insert a column break at the caret: the paragraph is split there and what
/// follows starts the next column of its section - in a section of one column
/// the next page, as in Word. (The break stands at the end of the first half:
/// PageBreakAfter with ColumnBreakAfter, saved as <w:br w:type="column"/>.)
/// Only in a paragraph of the body's own (not in a table cell or a text box).
let insertColumnBreak doc =
    match doc.Selection with
    | Some before ->
        match findBlock doc before.Focus.BlockId with
        | Some (Paragraph _)
        | Some (Heading _) ->
            let split = splitBlock doc
            if obj.ReferenceEquals(split, doc) then doc
            else
                match findParagraph split before.Focus.BlockId with
                | Some first -> { replaceParagraph split first.Id { first with Metadata = { first.Metadata with PageBreakAfter = true; ColumnBreakAfter = true } } with Selection = split.Selection }
                | None -> split
        | _ -> doc
    | None -> doc

/// The extent of the hyperlink at an offset of a paragraph: (target, start,
/// end) of the run of text that has the same target.
let hyperlinkAt (para: ParagraphBlock) (offset: int) : (string * int * int) option =
    let length = TextBuffer.length para.Buffer
    let targetAt (index: int) =
        if index < 0 || index >= length then None
        else
            let style = TextBuffer.getStyleAt index para.Buffer
            style.Hyperlink |> Option.orElse style.InactiveHyperlink
    // The character after the caret, or the one before it at the end of a link.
    let probe =
        match targetAt offset with
        | Some _ -> offset
        | None -> offset - 1
    match targetAt probe with
    | None -> None
    | Some target ->
        let mutable first = probe
        while first > 0 && targetAt (first - 1) = Some target do first <- first - 1
        let mutable last = probe + 1
        while last < length && targetAt last = Some target do last <- last + 1
        Some (target, first, last)


/// What the caret is in, for a host's contextual tools. Read off the document
/// alone (no layout).
type CaretContext = {
    InTable: bool
    /// 0-based row and column of the cell the caret is in (-1 outside a table).
    Row: int
    Column: int
    RowCount: int
    ColumnCount: int
    /// The selection spans several cells of one table.
    CanMergeCells: bool
    /// The caret's cell is a merged one.
    CanSplitCell: bool
    InHyperlink: bool
    InList: bool
    /// Always false: headers and footers are not edited in place.
    InHeaderFooter: bool
    /// The selection is one picture.
    ImageSelected: bool
}

let caretContext (doc: Document) : CaretContext =
    let outside =
        { InTable = false; Row = -1; Column = -1; RowCount = 0; ColumnCount = 0; CanMergeCells = false; CanSplitCell = false
          InHyperlink = false; InList = false; InHeaderFooter = false; ImageSelected = false }
    match doc.Selection with
    | None -> outside
    | Some selection ->
        let focus = selection.Focus
        let sameBlock = selection.Anchor.BlockId = focus.BlockId
        let first = if sameBlock then min selection.Anchor.Offset focus.Offset else focus.Offset
        let last = if sameBlock then max selection.Anchor.Offset focus.Offset else focus.Offset
        let paragraph = findParagraph doc focus.BlockId
        let inText =
            match paragraph with
            | None -> outside
            | Some para ->
                { outside with
                    InHyperlink = (hyperlinkAt para first).IsSome
                    InList = para.Metadata.ListStyle <> NoList
                    ImageSelected =
                        sameBlock && last - first = 1
                        && (para.Buffer.Runs |> List.exists (function Image image -> image.Start = first | _ -> false)) }
        match tryFindTableCellByParagraph doc focus.BlockId with
        | None -> inText
        | Some (tableId, rowIndex, colIndex) ->
            let table = tryFindTable doc tableId
            match table with
            | None -> inText
            | Some table ->
                let cell = List.tryItem rowIndex table.Rows |> Option.bind (fun row -> List.tryItem colIndex row.Cells)
                { inText with
                    InTable = true
                    Row = rowIndex
                    Column = colIndex
                    RowCount = table.Rows.Length
                    ColumnCount = table.Rows |> List.fold (fun widest row -> max widest (row.Cells |> List.sumBy (fun cell -> max 1 cell.ColSpan))) 0
                    CanMergeCells =
                        (match selectedTableCells doc with
                         | Some (_, firstRow, firstCol, lastRow, lastCol) -> firstRow <> lastRow || firstCol <> lastCol
                         | None -> false)
                    CanSplitCell =
                        (match cell with
                         | Some cell -> cell.ColSpan > 1 || cell.VerticalMerge = Some RestartMerge
                         | None -> false) }

/// Change the target of the hyperlink the caret is in (the whole link, not
/// just the selected part); None removes the link and keeps its text. With a
/// range selected that is not in a link, the range becomes a link.
let setHyperlinkTarget (target: string option) doc =
    match doc.Selection with
    | None -> doc
    | Some selection ->
        let link =
            if selection.Anchor.BlockId <> selection.Focus.BlockId then None
            else findParagraph doc selection.Focus.BlockId |> Option.bind (fun para -> hyperlinkAt para (min selection.Anchor.Offset selection.Focus.Offset) |> Option.map (fun found -> para, found))
        match link with
        | Some (para, (_, first, last)) ->
            let buffer = TextBuffer.applyStyle first (last - first) (fun style -> { style with Hyperlink = None; InactiveHyperlink = None }) para.Buffer
            let buffer = match target with Some _ -> TextBuffer.setHyperlink first (last - first) target buffer | None -> buffer
            replaceParagraph doc para.Id { para with Buffer = buffer }
        | None -> if selection.IsCollapsed then doc else setHyperlink target doc

/// Insert a hyperlink: with a range selected the range becomes the link;
/// with a caret, `text` (or the target itself) is inserted as a link.
let insertHyperlink (target: string) (text: string) doc =
    match doc.Selection with
    | None -> doc
    | Some selection when not selection.IsCollapsed -> setHyperlink (Some target) doc
    | Some selection ->
        let shown = if System.String.IsNullOrEmpty text then target else text
        let inserted = insertText shown doc
        match inserted.Selection with
        | Some after when after.Focus.BlockId = selection.Focus.BlockId && after.Focus.Offset >= shown.Length ->
            let start = after.Focus.Offset - shown.Length
            let linked =
                { inserted with Selection = Some { Anchor = { after.Focus with Offset = start }; Focus = after.Focus } }
                |> setHyperlink (Some target)
            { linked with Selection = inserted.Selection }
        | _ -> inserted

/// The top-level paragraph the caret is in or, for a caret in a table or a
/// text box, none (a section ends with a paragraph of the body).
let private caretBodyParagraphIndex (doc: Document) : int option =
    match doc.Selection with
    | None -> None
    | Some sel ->
        match locateBlock doc sel.Focus.BlockId with
        | ValueSome (struct (index, (Paragraph _ | Heading _))) -> Some index
        | _ -> None

/// A section break at the caret. `sectionBreak` is how the section that
/// STARTS at the break - the text after the caret's paragraph - starts:
/// "continuous" (on the same page), "nextPage", "evenPage", "oddPage"; None
/// takes the break away.
///
/// In the file a section's properties stand at its end (the w:sectPr of its
/// last paragraph, the body's for the last section) and the w:type among them
/// says how that section starts. So the caret's paragraph becomes the end of a
/// section that starts as the section it is cut from did (and has that
/// section's properties: page, columns, headers), and the type goes to the
/// properties of the section after it - on the next paragraph that ends a
/// section, or the document's own for the last. On a paragraph that ends a
/// section already only the type of the section after it changes. Taking the
/// break away leaves one section with the properties of the one after the
/// break, which starts as the one before the break did.
let setCurrentSectionBreak (sectionBreak: string option) (doc: Document) : Document =
    match caretBodyParagraphIndex doc with
    | None -> doc
    | Some caretIndex ->
        let blocks = List.toArray doc.Blocks
        let paragraphAt index =
            match blocks.[index] with
            | Paragraph para
            | Heading (_, para) -> Some para
            | _ -> None
        let withMetadata index (change: BlockMetadata -> BlockMetadata) =
            match blocks.[index] with
            | Paragraph para -> blocks.[index] <- Paragraph { para with Metadata = change para.Metadata }
            | Heading (level, para) -> blocks.[index] <- Heading (level, { para with Metadata = change para.Metadata })
            | _ -> ()
        // The paragraph that ends the section after the caret's paragraph (none: the last section).
        let mutable following = -1
        let mutable probe = caretIndex + 1
        while following < 0 && probe < blocks.Length do
            match paragraphAt probe with
            | Some para when para.Metadata.SectionBreak.IsSome -> following <- probe
            | _ -> ()
            probe <- probe + 1
        let followingStart =
            if following >= 0 then (paragraphAt following |> Option.bind (fun para -> para.Metadata.SectionBreak))
            else doc.PageSetup.SectionStart
        let setFollowingStart (start: string option) (pageSetup: PageSetup) : PageSetup =
            if following >= 0 then
                // (A paragraph that ends a section says a type: the section's properties are known by it.)
                withMetadata following (fun metadata -> { metadata with SectionBreak = Some (start |> Option.defaultValue "nextPage") })
                pageSetup
            else { pageSetup with SectionStart = start }
        let caret = (paragraphAt caretIndex).Value
        match sectionBreak, caret.Metadata.SectionBreak with
        | None, None -> doc
        | None, Some own ->
            // The break is taken away: the text before it joins the section after it, which starts as its own did.
            withMetadata caretIndex (fun metadata ->
                { metadata with
                    SectionBreak = None; SectionColumns = None; SectionPageSetup = None; SectionPageNumberStart = None
                    SectionPageNumberFormat = None; SectionVerticalAlign = None; HeaderReferences = []; FooterReferences = [] })
            let pageSetup = setFollowingStart (Some own) doc.PageSetup
            { doc with Blocks = List.ofArray blocks; PageSetup = pageSetup }
        | Some kind, Some _ ->
            let pageSetup = setFollowingStart (Some kind) doc.PageSetup
            { doc with Blocks = List.ofArray blocks; PageSetup = pageSetup }
        | Some kind, None ->
            // The section is cut in two: the first part has its properties and starts as it did.
            let columns, geometry, numberStart, numberFormat, vAlign, headers, footers =
                if following >= 0 then
                    let metadata = (paragraphAt following).Value.Metadata
                    metadata.SectionColumns, metadata.SectionPageSetup, metadata.SectionPageNumberStart, metadata.SectionPageNumberFormat, metadata.SectionVerticalAlign, metadata.HeaderReferences, metadata.FooterReferences
                else
                    let setup = doc.PageSetup
                    let geometry : SectionPageSetup =
                        { PageWidth = setup.PageWidth; PageHeight = setup.PageHeight; MarginLeft = setup.MarginLeft; MarginRight = setup.MarginRight
                          MarginTop = setup.MarginTop; MarginBottom = setup.MarginBottom; HeaderDistance = setup.HeaderDistance; FooterDistance = setup.FooterDistance
                          DifferentFirstPage = setup.DifferentFirstPage; EvenAndOddHeaders = setup.EvenAndOddHeaders }
                    setup.Columns, Some geometry, setup.PageNumberStart, setup.PageNumberFormat, setup.VerticalAlign, setup.HeaderReferences, setup.FooterReferences
            withMetadata caretIndex (fun metadata ->
                { metadata with
                    SectionBreak = Some (followingStart |> Option.defaultValue "nextPage")
                    SectionColumns = columns; SectionPageSetup = geometry; SectionPageNumberStart = numberStart
                    SectionPageNumberFormat = numberFormat; SectionVerticalAlign = vAlign; HeaderReferences = headers; FooterReferences = footers })
            // (A page number the section started with is the first part's: the second goes on counting.)
            if following >= 0 then withMetadata following (fun metadata -> { metadata with SectionPageNumberStart = None })
            let pageSetup = setFollowingStart (Some kind) { doc.PageSetup with PageNumberStart = (if following >= 0 then doc.PageSetup.PageNumberStart else None) }
            { doc with Blocks = List.ofArray blocks; PageSetup = pageSetup }

/// The paragraph that ends the section the caret is in (None: the final section, which no paragraph ends).
let private currentSectionEnd (doc: Document) : ParagraphBlock option option =
    match doc.Selection with
    | None -> None
    | Some sel ->
        let containsFocus block =
            match block with
            | Table table -> table.Id = sel.Focus.BlockId || (paragraphsOfBlocks [ block ] |> List.exists (fun para -> para.Id = sel.Focus.BlockId))
            | _ -> paragraphsOfBlocks [ block ] |> List.exists (fun para -> para.Id = sel.Focus.BlockId)
        doc.Blocks
        |> List.tryFindIndex containsFocus
        |> Option.map (fun index ->
            doc.Blocks
            |> List.skip index
            |> List.tryPick (fun block ->
                match block with
                | Paragraph para
                | Heading (_, para) when para.Metadata.SectionBreak.IsSome -> Some para
                | _ -> None))

/// Geometry of the section containing the selection; an unstated geometry
/// uses the same final-section defaults as layout.
let currentSectionPageSetup (doc: Document) : SectionPageSetup option =
    let final = doc.PageSetup
    let defaults: SectionPageSetup =
        { PageWidth = final.PageWidth; PageHeight = final.PageHeight
          MarginLeft = final.MarginLeft; MarginRight = final.MarginRight
          MarginTop = final.MarginTop; MarginBottom = final.MarginBottom
          HeaderDistance = final.HeaderDistance; FooterDistance = final.FooterDistance
          DifferentFirstPage = final.DifferentFirstPage; EvenAndOddHeaders = final.EvenAndOddHeaders }
    match currentSectionEnd doc with
    | None -> None
    | Some (Some para) -> Some (para.Metadata.SectionPageSetup |> Option.defaultValue defaults)
    | Some None -> Some defaults

let setCurrentSectionPageSetup (geometry: SectionPageSetup) (doc: Document) =
    match currentSectionEnd doc, currentSectionPageSetup doc with
    | _, Some old when old = geometry -> doc
    | Some (Some para), Some _ ->
        let updated = { para with Metadata = { para.Metadata with SectionPageSetup = Some geometry } }
        { doc with
            Blocks = doc.Blocks |> List.map (function
                | Paragraph p when p.Id = para.Id -> Paragraph updated
                | Heading (level, p) when p.Id = para.Id -> Heading (level, updated)
                | block -> block) }
    | Some None, Some _ ->
        let old = doc.PageSetup
        { doc with
            PageSetup =
                { old with PageWidth = geometry.PageWidth; PageHeight = geometry.PageHeight
                           MarginLeft = geometry.MarginLeft; MarginRight = geometry.MarginRight
                           MarginTop = geometry.MarginTop; MarginBottom = geometry.MarginBottom
                           HeaderDistance = geometry.HeaderDistance; FooterDistance = geometry.FooterDistance
                           DifferentFirstPage = geometry.DifferentFirstPage; EvenAndOddHeaders = geometry.EvenAndOddHeaders } }
    | _ -> doc

/// How the text of a page that is not full stands between the margins in the section the
/// caret is in (w:vAlign): Some "center", "bottom" or "both" (spread from top to bottom);
/// None (or "top"): from the top. On the paragraph that ends that section, or on the page
/// setup for the final section.
let setCurrentSectionVerticalAlign (align: string option) (doc: Document) : Document =
    let align = match align with Some "top" | Some "" -> None | other -> other
    match currentSectionEnd doc with
    | None -> doc
    | Some (Some para) ->
        if para.Metadata.SectionVerticalAlign = align then doc
        else
            let newPara = { para with Metadata = { para.Metadata with SectionVerticalAlign = align } }
            { doc with
                Blocks =
                    doc.Blocks
                    |> List.map (fun block ->
                        match block with
                        | Paragraph current when current.Id = para.Id -> Paragraph newPara
                        | Heading (level, current) when current.Id = para.Id -> Heading (level, newPara)
                        | other -> other) }
    | Some None ->
        if doc.PageSetup.VerticalAlign = align then doc
        else { doc with PageSetup = { doc.PageSetup with VerticalAlign = align } }

/// The vertical alignment of the section the caret is in ("top" where it states none).
let currentSectionVerticalAlign (doc: Document) : string =
    match currentSectionEnd doc with
    | Some (Some para) -> para.Metadata.SectionVerticalAlign |> Option.defaultValue "top"
    | _ -> doc.PageSetup.VerticalAlign |> Option.defaultValue "top"

/// Set the header / footer references of the section the cursor is in: on the
/// paragraph that ends that section, or on the page setup for the final
/// section (which no paragraph ends).
let setCurrentHeaderFooterReferences headerReferences footerReferences doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        let containsFocus block =
            match block with
            | Table table -> table.Id = sel.Focus.BlockId || (paragraphsOfBlocks [ block ] |> List.exists (fun para -> para.Id = sel.Focus.BlockId))
            | _ -> paragraphsOfBlocks [ block ] |> List.exists (fun para -> para.Id = sel.Focus.BlockId)
        match doc.Blocks |> List.tryFindIndex containsFocus with
        | None -> doc
        | Some index ->
            let sectionEnd =
                doc.Blocks
                |> List.skip index
                |> List.tryPick (fun block ->
                    match block with
                    | Paragraph para
                    | Heading (_, para) when endsSection para.Metadata -> Some para
                    | _ -> None)
            match sectionEnd with
            | Some para ->
                let newPara = { para with Metadata = { para.Metadata with HeaderReferences = headerReferences; FooterReferences = footerReferences } }
                { doc with
                    Blocks =
                        doc.Blocks
                        |> List.map (fun block ->
                            match block with
                            | Paragraph current when current.Id = para.Id -> Paragraph newPara
                            | Heading (level, current) when current.Id = para.Id -> Heading (level, newPara)
                            | other -> other) }
            | None ->
                { doc with PageSetup = { doc.PageSetup with HeaderReferences = headerReferences; FooterReferences = footerReferences } }