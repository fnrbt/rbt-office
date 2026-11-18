module WordEditorPackage

/// A Word part is not an untyped bag of properties: source origins and all
/// content which the editor cannot model remain attached to their own stories.
type SourceOrigin =
    | ReadAt of int
    | LikeAt of int

type SourceFragment = {
    Origin: SourceOrigin option
    Xml: string option
    Copy: bool
    Same: bool
}

type Border = {
    Style: string option
    Color: string option
    Width: float option
    Space: float option
}

type Borders = {
    Top: Border option
    Right: Border option
    Bottom: Border option
    Left: Border option
    InsideH: Border option
    InsideV: Border option
}

type Shading = {
    Pattern: string option
    ForegroundColor: string option
    BackgroundColor: string option
}

type TabStop = { Position: float; Alignment: string; Leader: string option }
type SectionColumns = { Count: int; Gap: float; Separator: bool; Widths: (float * float) list }
type HeaderFooterReference = { Kind: string; RelationshipId: string }

/// The style fields a block carries when a style catalog is absent.
type InferredStyle = {
    Name: string option
    IsDefault: bool option
    BasedOn: string option
    Next: string option
    Linked: string option
    IsQuickStyle: bool option
    Aliases: string option
    UiPriority: int option
    SemiHidden: bool option
    UnhideWhenUsed: bool option
    Bold: bool option
    Italic: bool option
    Underline: bool option
    UnderlineStyle: string option
    UnderlineColor: string option
    AllCaps: bool option
    SmallCaps: bool option
    Strike: bool option
    DoubleStrike: bool option
    Hidden: bool option
    Position: float option
    EmphasisMark: string option
    VertAlign: string option
    CharacterSpacing: float option
    FontSize: float option
    FontFamily: string option
    Color: string option
    Highlight: string option
    SpacingBefore: float option
    SpacingAfter: float option
    ContextualSpacing: bool option
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
    DefaultTabStop: float option
    /// See DocxReader.PageSetup.CompatibilityMode.
    CompatibilityMode: int option
    Columns: SectionColumns option
    SectionStart: string option
    HeaderReferences: HeaderFooterReference list
    FooterReferences: HeaderFooterReference list
    PageNumberStart: int option
    PageNumberFormat: string option
    VerticalAlign: string option
}

type Revision = {
    Id: string option
    Kind: string option
    Author: string option
    Date: string option
    Inserted: DocxReader.RevisionMark option
    MoveName: string option
}

type Image = {
    Source: string option
    Linked: bool
    Width: float option
    Height: float option
    Alt: string option
    Name: string option
    Title: string option
    Floating: bool
    Wrap: string option
    PositionX: float option
    PositionY: float option
    PositionHRelativeFrom: string option
    PositionVRelativeFrom: string option
    PositionHAlign: string option
    PositionVAlign: string option
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
    WrapText: string option
    WrapPolygon: (float * float) list
    FlipH: bool
    FlipV: bool
    Opaque: DocxReader.OpaqueObject option
}

type RunFormatting = {
    Style: string option
    StyleInfo: InferredStyle option
    Bold: bool
    Italic: bool
    Underline: bool
    UnderlineStyle: string option
    UnderlineColor: string option
    AllCaps: bool
    SmallCaps: bool
    Strike: bool
    DoubleStrike: bool
    Hidden: bool
    Position: float
    EmphasisMark: string option
    FontSize: float
    CharacterSpacing: float
    FontFamily: string
    FontFamilyInherited: bool option
    Color: string option
    Highlight: string option
    VertAlign: string option
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
    ExplicitColorAuto: bool
}

type RunContent =
    | Text of string
    | Picture of Image
    | Opaque of DocxReader.OpaqueObject
    | Equation of text: string * xml: string
    | PageBreak
    | FieldBoundary of ModelText.FieldBoundary

type ContentControlLayer = {
    Group: string
    Id: string option
    Tag: string option
    Alias: string option
    Xml: string option
}

type EditorRun = {
    Content: RunContent
    Formatting: RunFormatting
    Hyperlink: string option
    InactiveHyperlink: string option
    CommentId: string option
    CommentText: string option
    CommentAuthor: string option
    CommentDate: string option
    NoteId: string option
    NoteType: string option
    NoteText: string option
    BookmarkStarts: ModelText.BookmarkStart list
    BookmarkEnds: ModelText.BookmarkEnd list
    FieldCode: string option
    FieldResult: string option
    Revision: Revision option
    RenderedPageBreakBefore: bool
    Checkbox: bool option
    /// Inline controls containing this run, outermost first.
    ContentControls: ContentControlLayer list
    /// The start offset of w:r in its source part; None for newly authored runs.
    SourceRun: int option
    NoProps: bool
}


type ParagraphFormatting = {
    Style: string option
    StyleInfo: InferredStyle option
    Alignment: string
    HasExplicit: bool
    ExplicitAlignment: string option
    IndentLeft: float option
    IndentRight: float option
    FirstLineIndent: float option
    SpacingBefore: float option
    SpacingAfter: float option
    LineSpacing: float option
    LineSpacingRule: string option
    RightTabPosition: float option
    RightTabLeader: string option
    TabStops: TabStop list
    Rtl: bool
    ContextualSpacing: bool option
    KeepNext: bool option
    KeepLines: bool option
    WidowControl: bool option
    PageBreakBefore: bool
    PageBreakBeforeDerived: bool
    PageBreakAtStart: bool
    PageBreakAfter: bool
    ColumnBreakAfter: bool
    RenderedPageBreakBefore: int option
    BackgroundColor: string option
    ShadingPattern: string option
    ShadingForeground: string option
    Borders: Borders option
    ParagraphBorderBefore: float option
    ParagraphBorderAfter: float option
    OutlineLevel: int option
    ExplicitParagraphFormat: bool
    ExplicitIndentLeft: float option
    ExplicitIndentRight: float option
    ExplicitFirstLineIndent: float option
    ExplicitSpacingBefore: float option
    ExplicitSpacingAfter: float option
    ExplicitLineSpacing: float option
    ExplicitContextualSpacing: bool option
    ExplicitBackgroundColor: string option
    ExplicitShadingPattern: string option
    ExplicitShadingForeground: string option
    ExplicitTabs: bool
    HasMarkProperties: bool
    ExplicitMarkFontSize: float option
    ParagraphMarkFontSize: float option
    MarkRevision: Revision option
    DropCapKind: string option
    DropCapLines: int option
    DropCapDistance: float option
    ContentControlId: string option
    ContentControlTag: string option
    ContentControlAlias: string option
    ContentControlXml: string option
    ContentControlGroup: string option
    ContentControlAncestors: ContentControlLayer list
    EnclosingCellControls: ContentControlLayer list
}

type ListLevel = {
    Level: int
    NumberFormat: string option
    LevelText: string option
    Start: int option
    IndentLeft: float option
    FirstLineIndent: float option
    Suffix: string option
    Justification: string option
    Legal: bool
    Restart: int option
    MarkFont: string option
    MarkBold: bool option
    MarkItalic: bool option
    MarkColor: string option
    MarkSize: float option
}

type ParagraphList = {
    Style: string
    Level: int
    Id: string option
    SourceId: string option
    Restart: bool
    FormatEdited: bool
    NumberFormat: string option
    LevelText: string option
    Start: int option
    StartOverrides: (int * int) list
    ListLevelIndents: (float option * float option) option
    Levels: ListLevel list
}

type EditorBlock =
    | Paragraph of EditorParagraph
    | Table of EditorTable
    | TextBox of EditorTextBox

and EditorParagraph = {
    Runs: EditorRun array
    HeadingLevel: int option
    Formatting: ParagraphFormatting
    List: ParagraphList option
    SectionBreak: string option
    SectionPageSetup: PageSetup option
    SectionColumns: SectionColumns option
    SectionVerticalAlign: string option
    SectionPageNumberStart: int option
    SectionPageNumberFormat: string option
    HeaderReferences: HeaderFooterReference list
    FooterReferences: HeaderFooterReference list
    Source: SourceFragment
    /// Source paragraph edit, not an XML string: split/join and formatting
    /// changes need both the old and the new model to splice runs safely.
    SourceParagraph: ParagraphEdit option
    SpliceOriginal: EditorParagraph option
    SpliceFrom: (int * int) option
    SpliceLeading: bool
    SpliceTrailing: bool
    SpliceJoined: EditorParagraph array
    SourceIndex: int option
    SourceUnchanged: bool
    SourceWhy: string option
    SourceTextChanges: SourceTextChanges option
}

/// A paragraph edit describes how its source element changes. The original
/// paragraph and joined source paragraphs are retained as typed records.
and ParagraphEdit = {
    Paragraph: int
    Block: int option
    Para: int option
    Text: ParagraphTextEdit option
    OldText: string option
    Before: EditorParagraph option
    After: EditorParagraph option
    NewText: string option
    Regenerate: bool
    Verbatim: bool
    Why: string option
    Cut: (int * int) option
    Join: JoinedParagraph array
}

and JoinedParagraph = {
    Block: int
    Para: int
    Text: string
    Original: EditorParagraph
}

and ParagraphTextEdit = {
    Paragraph: int
    Offset: int
    Deleted: int
    Inserted: string
    JoinBefore: bool
    JoinAfter: bool
    RunsPlain: bool
    Witnesses: int array
    MarkStyle: bool
    NewText: string
}

and SourceTextChanges = {
    Texts: string array
    Edits: ParagraphEdit array
}

and EditorTable = {
    Rows: EditorRow array
    PreferredWidth: float option
    PreferredWidthType: string option
    Alignment: string option
    Indent: float option
    LayoutMode: string option
    GridColumns: float array
    Borders: Borders option
    BackgroundColor: string option
    Shading: Shading option
    CellSpacing: float option
    PaddingTop: float option
    PaddingRight: float option
    PaddingBottom: float option
    PaddingLeft: float option
    Floating: bool
    FloatingXml: string option
    StyleId: string option
    Look: EditorTableLook option
    ContentControlId: string option
    ContentControlTag: string option
    ContentControlAlias: string option
    ContentControlXml: string option
    ContentControlGroup: string option
    ContentControlAncestors: ContentControlLayer list
    Source: SourceFragment
    SourceIndex: int option
    SourceUnchanged: bool
    SourceWhy: string option
    SourceTextChanges: SourceTextChanges option
    SourceTable: TableSource option
}

and EditorTableLook = {
    FirstRow: bool
    LastRow: bool
    FirstColumn: bool
    LastColumn: bool
    HorizontalBanding: bool
    VerticalBanding: bool
}

and EditorRow = {
    Cells: EditorCell array
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
    BordersOverride: Borders option
    Source: SourceFragment
    SourceRow: RowSource option
}

and EditorCell = {
    Paragraphs: EditorBlock array
    NestedTables: EditorTable array
    NestedTableAnchors: int array
    EnclosingControls: ContentControlLayer list
    ColSpan: int
    BackgroundColor: string option
    Shading: Shading option
    Width: float option
    PreferredWidth: float option
    PreferredWidthType: string option
    BorderColor: string option
    BorderWidth: float option
    Borders: Borders option
    VerticalAlignment: string option
    VerticalMerge: string option
    PaddingTop: float option
    PaddingRight: float option
    PaddingBottom: float option
    PaddingLeft: float option
    Source: SourceFragment
    SourceCell: CellSource option
}

and EditorTextBox = {
    Paragraphs: EditorBlock array
    Children: EditorTextBox array
    ShapeId: string option
    ShapeKind: string option
    Name: string option
    Geometry: string
    CustomPath: (string * float array) list
    CustomWidth: float
    CustomHeight: float
    X: float option
    Y: float option
    PositionHRelativeFrom: string option
    PositionVRelativeFrom: string option
    PositionHAlign: string option
    PositionVAlign: string option
    Width: float
    Height: float
    FillColor: string option
    BorderColor: string option
    NoFill: bool
    NoLine: bool
    BorderWidth: float option
    BorderDash: string option
    Gradient: (float * string) list
    GradientAngle: float
    GradientRadial: bool
    Rotation: float
    FlipH: bool
    FlipV: bool
    Inline: bool
    Wrap: string option
    BehindDoc: bool option
    AllowOverlap: bool option
    LayoutInCell: bool option
    DistanceTop: float option
    DistanceBottom: float option
    DistanceLeft: float option
    DistanceRight: float option
    RelativeHeight: string option
    AnchorBlock: int option
    InsetLeft: float option
    InsetTop: float option
    InsetRight: float option
    InsetBottom: float option
    VerticalAnchor: string option
    AutoFit: bool
    NoTextWrap: bool
    FitText: bool
    Opacity: float
    ChildSpace: (float * float) option
    ImageSource: string option
    SourceXml: string option
    SourceMoved: ShapeMovement option
    SourceIndex: int option
    SourceUnchanged: bool
    SourceWhy: string option
    SourceTextChanges: SourceTextChanges option
}

and TableSource = {
    At: int
    Block: int option
    Before: EditorTable option
    Was: EditorTable option
    Copy: bool
    Same: bool
}

and RowSource = {
    At: int
    Before: EditorRow option
    Was: EditorRow option
    Copy: bool
    Same: bool
}

and CellSource = {
    At: int
    Before: EditorCell option
    Was: EditorCell option
    Copy: bool
    Same: bool
}

and ShapeMovement = { H: bool; V: bool; Size: bool }

/// Catalog styles preserve unset vs explicitly false/zero properties; effective
/// text formatting on a run is intentionally a different record.
type StyleRun = {
    Bold: bool option
    Italic: bool option
    Underline: bool option
    UnderlineStyle: string option
    UnderlineColor: string option
    AllCaps: bool option
    SmallCaps: bool option
    Strike: bool option
    DoubleStrike: bool option
    Hidden: bool option
    Position: float option
    EmphasisMark: string option
    VertAlign: string option
    FontSize: float option
    CharacterSpacing: float option
    FontFamily: string option
    Color: string option
    Highlight: string option
}

type StyleParagraph = {
    SpacingBefore: float option
    SpacingAfter: float option
    ContextualSpacing: bool option
    LineSpacing: float option
    LineSpacingRule: string option
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

type Style = {
    Id: string
    Name: string option
    StyleType: string
    IsDefault: bool
    BasedOn: string option
    Next: string option
    Linked: string option
    IsQuickStyle: bool option
    Aliases: string option
    UiPriority: int option
    SemiHidden: bool option
    UnhideWhenUsed: bool option
    Run: StyleRun option
    Paragraph: StyleParagraph option
}

type Story = {
    Id: string
    Kind: string option
    HasMark: bool option
    MarkStyle: string option
    Text: string option
    Blocks: EditorBlock array
    SourceStoryCounts: int array
    Unchanged: bool
}

type HeaderFooterPart = {
    RelationshipId: string option
    Kind: string
    Blocks: EditorBlock array
    SourceStoryCounts: int array
    SourceUnchanged: bool
}

type CommentPart = {
    Id: string
    Author: string option
    Date: string option
    Initials: string option
    Text: string option
    Blocks: EditorBlock array
    SourceStoryCounts: int array
    Unchanged: bool
}

type Relationship = { Id: string; Type: string; Target: string; External: bool }
type PackagePart = { Path: string; Data: byte[] }

type PreservedPackage = {
    Parts: PackagePart list
    RootRelationships: Relationship list
    DocumentRelationshipId: string
    PartRelationships: Map<string, Relationship list>
    Defaults: (string * string) list
    Overrides: (string * string) list
    MainContentType: string option
    StylesXml: string option
    SettingsXml: string option
    SourceXml: Map<string, string>
    SourceRelationships: Map<string, Relationship list>
    SourceParts: PackagePart list
    BodyContentChunks: int array
    BodyItemChunks: int array
    BodyChunkCount: int
    PartItemChunks: Map<string, int array * int>
    NoteItemChunks: Map<string, int array * int>
    PartPaths: Map<string, string>
    OriginalText: Map<string, string>
    /// Original container (not inflated entries) for exact unchanged saves.
    OpenedBytes: byte[] option
}

type EditorPackage = {
    Blocks: EditorBlock array
    Styles: Style array
    HeaderParts: HeaderFooterPart array
    FooterParts: HeaderFooterPart array
    Footnotes: Story array
    Endnotes: Story array
    Comments: CommentPart array
    Theme: DocxReader.Theme option
    PageSetup: PageSetup option
    OriginalPageSetup: PageSetup option
    Preserved: PreservedPackage option
    SourceStoryCounts: int array
    ChangedStyles: string array option
    ChangedListLevels: string array
    PageSetupUnchanged: bool
    ListsUnchanged: bool
    FootnotesUnchanged: bool
    EndnotesUnchanged: bool
    CommentsUnchanged: bool
    Unchanged: bool
    RegenerateAll: bool
}

/// Retain package relationships, unknown entries and source XML from a parsed
/// document without going through a JS object or re-inflating its ZIP entries.
let preservedOfDocument (document: DocxReader.WordDocument) : PreservedPackage =
    let source = document.Preserved
    let relationship (item: DocxReader.PreservedRelationship) : Relationship =
        { Id = item.Id; Type = item.Type; Target = item.Target; External = item.External }
    let relationships = List.map relationship
    let parts (entries: DocxReader.PreservedPart list) =
        entries |> List.map (fun part -> { Path = part.Path; Data = part.Data })
    let shapeChunks (chunks: int array) (shapes: DocxReader.Shape list) =
        shapes
        |> List.map (fun shape ->
            match shape.AnchorBlock with
            | Some index when index >= 0 && index < chunks.Length -> chunks.[index]
            | _ -> -1)
        |> List.toArray
    let bodyChunks = source.BodyContentChunks |> List.toArray
    let partChunks =
        document.Headers @ document.Footers
        |> List.choose (fun part ->
            match part.RelationshipId with
            | Some id when part.ItemChunks.Length = part.Blocks.Length ->
                let chunks = part.ItemChunks |> List.toArray
                let shapes = document.PartShapes |> Map.tryFind id |> Option.defaultValue []
                Some (id, (Array.append chunks (shapeChunks chunks shapes), part.ChunkCount))
            | _ -> None)
        |> Map.ofList
    let noteChunks key (id: string) (blocks: DocxReader.DocumentContent list) (chunks: int list) count =
        if chunks.Length = blocks.Length && count > 0 then Some (key + ":" + id, (List.toArray chunks, count)) else None
    let storyChunks =
        [ yield! document.Footnotes |> List.choose (fun note -> noteChunks "footnotes" note.Id note.Blocks note.ItemChunks note.ChunkCount)
          yield! document.Endnotes |> List.choose (fun note -> noteChunks "endnotes" note.Id note.Blocks note.ItemChunks note.ChunkCount)
          yield! document.Comments |> List.choose (fun comment -> noteChunks "comments" comment.Id comment.Blocks comment.ItemChunks comment.ChunkCount) ]
        |> Map.ofList
    { Parts = parts source.Parts
      RootRelationships = relationships source.RootRelationships
      DocumentRelationshipId = source.DocumentRelationshipId
      PartRelationships = source.PartRelationships |> Map.map (fun _ entries -> relationships entries)
      Defaults = source.Defaults
      Overrides = source.Overrides
      MainContentType = source.MainContentType
      StylesXml = source.StylesXml
      SettingsXml = source.SettingsXml
      SourceXml = source.SourceXml
      SourceRelationships = source.SourceRelationships |> Map.map (fun _ entries -> relationships entries)
      SourceParts = parts source.SourceParts
      BodyContentChunks = bodyChunks
      BodyItemChunks = Array.append bodyChunks (shapeChunks bodyChunks document.Shapes)
      BodyChunkCount = source.BodyChunkCount
      PartItemChunks = partChunks
      NoteItemChunks = storyChunks
      PartPaths = source.PartPaths
      OriginalText = source.OriginalText
      OpenedBytes = None }

#if FABLE_COMPILER
open Fable.Core
open Fable.Core.JsInterop

let inline private (==>) key value = key, box value

[<Emit("$0 == null")>]
let private isMissing (_value: obj) : bool = jsNative

let inline private memberOption<'a> (value: obj) : 'a option =
    if isMissing value then None else Some (unbox<'a> value)

let private members (value: obj) : obj array =
    if isMissing value then [||] else unbox<obj array> value

let private stringMember (value: obj) : string option =
    memberOption<string> value |> Option.filter (fun text -> text <> "")

let private relationshipOfObject (value: obj) : Relationship =
    { Id = unbox<string> value?Id
      Type = unbox<string> value?Type
      Target = unbox<string> value?Target
      External = memberOption<bool> value?External |> Option.defaultValue false }

let private relationshipToObject (value: Relationship) =
    createObj [ "Id" ==> value.Id; "Type" ==> value.Type; "Target" ==> value.Target; "External" ==> value.External ]

let private partOfObject (value: obj) : PackagePart =
    { Path = unbox<string> value?Path; Data = unbox<byte[]> value?Data }

let private partToObject (value: PackagePart) =
    createObj [ "Path" ==> value.Path; "Data" ==> value.Data ]

/// Browser boundary: progressive reader packages retain their original JS
/// shape, while the editor and the portable writer use records.
let preservedOfEditorObject (value: obj) : PreservedPackage option =
    if isMissing value then None
    else
        let relationships value = members value |> Array.map relationshipOfObject |> List.ofArray
        let parts value = members value |> Array.map partOfObject |> List.ofArray
        let keyValues key value name =
            members value |> Array.choose (fun item ->
                stringMember (item?(key)) |> Option.bind (fun label ->
                    stringMember (item?(name)) |> Option.map (fun text -> label, text)))
            |> Map.ofArray
        let relMaps value =
            members value
            |> Array.choose (fun item ->
                stringMember item?Part |> Option.map (fun key -> key, relationships item?Relationships))
            |> Map.ofArray
        let chunksOf (entry: obj) =
            (members entry?Chunks |> Array.map unbox<int>, memberOption<int> entry?Count |> Option.defaultValue 0)
        let chunks value =
            members value
            |> Array.choose (fun entry -> stringMember entry?Part |> Option.map (fun key -> key, chunksOf entry))
            |> Map.ofArray
        let noteChunks =
            members value?NoteItemChunks
            |> Array.choose (fun entry ->
                match stringMember entry?Part, stringMember entry?Id with
                | Some part, Some id -> Some (part + ":" + id, chunksOf entry)
                | _ -> None)
            |> Map.ofArray
        let pairs field1 field2 entries =
            members entries
            |> Array.choose (fun entry ->
                match stringMember (entry?(field1)), stringMember (entry?(field2)) with
                | Some first, Some second -> Some (first, second)
                | _ -> None)
            |> List.ofArray
        Some {
            Parts = parts value?Parts
            RootRelationships = relationships value?RootRelationships
            DocumentRelationshipId = stringMember value?DocumentRelationshipId |> Option.defaultValue "rId1"
            PartRelationships = relMaps value?PartRelationships
            Defaults = pairs "Extension" "ContentType" value?Defaults
            Overrides = pairs "PartName" "ContentType" value?Overrides
            MainContentType = stringMember value?MainContentType
            StylesXml = stringMember value?StylesXml
            SettingsXml = stringMember value?SettingsXml
            SourceXml = keyValues "Part" value?SourceXml "Xml"
            SourceRelationships = relMaps value?SourceRelationships
            SourceParts = parts value?SourceParts
            BodyContentChunks = members value?BodyContentChunks |> Array.map unbox<int>
            BodyItemChunks = members value?BodyItemChunks |> Array.map unbox<int>
            BodyChunkCount = memberOption<int> value?BodyChunkCount |> Option.defaultValue 0
            PartItemChunks = chunks value?PartItemChunks
            NoteItemChunks = noteChunks
            PartPaths = keyValues "Part" value?PartPaths "Path"
            OriginalText = keyValues "Path" value?OriginalText "Text"
            OpenedBytes = memberOption<byte[]> value?OpenedBytes }

let preservedToEditorObject (value: PreservedPackage option) : obj =
    match value with
    | None -> null
    | Some preserved ->
        let mapPairs key field (items: Map<string, string>) =
            items |> Map.toArray |> Array.map (fun (name, text) -> createObj [ key ==> name; field ==> text ])
        let relMap (items: Map<string, Relationship list>) =
            items |> Map.toArray |> Array.map (fun (name, rels) ->
                createObj [ "Part" ==> name; "Relationships" ==> (rels |> List.map relationshipToObject |> List.toArray) ])
        let pairs first second (items: (string * string) list) =
            items |> List.map (fun (name, text) -> createObj [ first ==> name; second ==> text ]) |> List.toArray
        let chunks (items: Map<string, int array * int>) =
            items |> Map.toArray |> Array.map (fun (part, (indices, count)) ->
                createObj [ "Part" ==> part; "Chunks" ==> indices; "Count" ==> count ])
        let noteChunks =
            preserved.NoteItemChunks
            |> Map.toArray
            |> Array.choose (fun (key, (indices, count)) ->
                let colon = key.IndexOf(':')
                if colon < 0 then None else
                    Some (createObj [
                        "Part" ==> key.Substring(0, colon); "Id" ==> key.Substring(colon + 1)
                        "Chunks" ==> indices; "Count" ==> count ]))
        createObj [
            "Parts" ==> (preserved.Parts |> List.map partToObject |> List.toArray)
            "RootRelationships" ==> (preserved.RootRelationships |> List.map relationshipToObject |> List.toArray)
            "DocumentRelationshipId" ==> preserved.DocumentRelationshipId
            "PartRelationships" ==> relMap preserved.PartRelationships
            "Defaults" ==> pairs "Extension" "ContentType" preserved.Defaults
            "Overrides" ==> pairs "PartName" "ContentType" preserved.Overrides
            "MainContentType" ==> preserved.MainContentType
            "StylesXml" ==> preserved.StylesXml
            "SettingsXml" ==> preserved.SettingsXml
            "SourceXml" ==> mapPairs "Part" "Xml" preserved.SourceXml
            "SourceRelationships" ==> relMap preserved.SourceRelationships
            "SourceParts" ==> (preserved.SourceParts |> List.map partToObject |> List.toArray)
            "BodyContentChunks" ==> preserved.BodyContentChunks
            "BodyItemChunks" ==> preserved.BodyItemChunks
            "BodyChunkCount" ==> preserved.BodyChunkCount
            "PartItemChunks" ==> chunks preserved.PartItemChunks
            "NoteItemChunks" ==> noteChunks
            "PartPaths" ==> mapPairs "Part" "Path" preserved.PartPaths
            "OriginalText" ==> mapPairs "Path" "Text" preserved.OriginalText
            "OpenedBytes" ==> preserved.OpenedBytes ]

let private inferredStyleFields (style: InferredStyle option) =
    match style with
    | None -> []
    | Some style -> [
        "StyleName" ==> style.Name; "IsDefaultStyle" ==> style.IsDefault
        "BasedOnStyle" ==> style.BasedOn; "NextStyle" ==> style.Next
        "LinkedStyle" ==> style.Linked; "IsQuickStyle" ==> style.IsQuickStyle
        "StyleAliases" ==> style.Aliases; "StyleUiPriority" ==> style.UiPriority
        "StyleSemiHidden" ==> style.SemiHidden
        "StyleUnhideWhenUsed" ==> style.UnhideWhenUsed
        "StyleBold" ==> style.Bold; "StyleItalic" ==> style.Italic
        "StyleUnderline" ==> style.Underline
        "StyleUnderlineStyle" ==> style.UnderlineStyle
        "StyleUnderlineColor" ==> style.UnderlineColor
        "StyleAllCaps" ==> style.AllCaps; "StyleSmallCaps" ==> style.SmallCaps
        "StyleStrike" ==> style.Strike; "StyleDoubleStrike" ==> style.DoubleStrike
        "StyleHidden" ==> style.Hidden; "StylePosition" ==> style.Position
        "StyleEmphasisMark" ==> style.EmphasisMark; "StyleVertAlign" ==> style.VertAlign
        "StyleFontSize" ==> style.FontSize; "StyleCharacterSpacing" ==> style.CharacterSpacing
        "StyleFontFamily" ==> style.FontFamily; "StyleColor" ==> style.Color
        "StyleHighlight" ==> style.Highlight
        "StyleSpacingBefore" ==> style.SpacingBefore; "StyleSpacingAfter" ==> style.SpacingAfter
        "StyleContextualSpacing" ==> style.ContextualSpacing
        "StyleLineSpacing" ==> style.LineSpacing; "StyleLineRule" ==> style.LineRule
        "StyleAlignment" ==> style.Alignment; "StyleIndentLeft" ==> style.IndentLeft
        "StyleIndentRight" ==> style.IndentRight; "StyleFirstLineIndent" ==> style.FirstLineIndent
        "StyleBackgroundColor" ==> style.BackgroundColor; "StyleOutlineLevel" ==> style.OutlineLevel
        "StyleKeepNext" ==> style.KeepNext; "StyleKeepLines" ==> style.KeepLines
        "StylePageBreakBefore" ==> style.PageBreakBefore; "StyleWidowControl" ==> style.WidowControl ]

let private sourceFields (source: SourceFragment) =
    let at, copy =
        match source.Origin with
        | Some (ReadAt at) -> Some at, false
        | Some (LikeAt at) -> Some at, true
        | None -> None, false
    [ "SourceAt" ==> at; "SourceCopy" ==> (copy || source.Copy); "SourceSame" ==> source.Same ]

let private revisionFields (revision: Revision option) =
    match revision with
    | None -> []
    | Some revision -> [
        "RevisionId" ==> revision.Id; "RevisionType" ==> revision.Kind
        "RevisionAuthor" ==> revision.Author; "RevisionDate" ==> revision.Date
        "InsertedRevisionId" ==> (revision.Inserted |> Option.bind (fun item -> item.Id))
        "InsertedRevisionAuthor" ==> (revision.Inserted |> Option.bind (fun item -> item.Author))
        "InsertedRevisionDate" ==> (revision.Inserted |> Option.bind (fun item -> item.Date))
        "RevisionMoveName" ==> revision.MoveName ]

let private imageFields (image: Image) = [
    "IsImage" ==> true; "ImageSrc" ==> image.Source; "ImageLinked" ==> image.Linked
    "ImageWidth" ==> image.Width; "ImageHeight" ==> image.Height
    "ImageAlt" ==> image.Alt; "ImageName" ==> image.Name; "ImageTitle" ==> image.Title
    "ImageFloating" ==> image.Floating; "ImageWrap" ==> image.Wrap
    "ImagePositionX" ==> image.PositionX; "ImagePositionY" ==> image.PositionY
    "ImagePositionHRelativeFrom" ==> image.PositionHRelativeFrom
    "ImagePositionVRelativeFrom" ==> image.PositionVRelativeFrom
    "ImagePositionHAlign" ==> image.PositionHAlign
    "ImagePositionVAlign" ==> image.PositionVAlign
    "ImageBehindDoc" ==> image.BehindDoc; "ImageAllowOverlap" ==> image.AllowOverlap
    "ImageLayoutInCell" ==> image.LayoutInCell
    "ImageDistanceTop" ==> image.DistanceTop; "ImageDistanceBottom" ==> image.DistanceBottom
    "ImageDistanceLeft" ==> image.DistanceLeft; "ImageDistanceRight" ==> image.DistanceRight
    "ImageRelativeHeight" ==> image.RelativeHeight
    "ImageLocked" ==> image.Locked; "ImageSimplePosition" ==> image.SimplePosition
    "ImageSimplePositionX" ==> image.SimplePositionX; "ImageSimplePositionY" ==> image.SimplePositionY
    "ImageCropLeft" ==> image.CropLeft; "ImageCropTop" ==> image.CropTop
    "ImageCropRight" ==> image.CropRight; "ImageCropBottom" ==> image.CropBottom
    "ImageFlipH" ==> image.FlipH; "ImageFlipV" ==> image.FlipV
    "ImageWrapText" ==> image.WrapText
    "ImageWrapPolygon" ==> (image.WrapPolygon |> List.collect (fun (x, y) -> [ x; y ]) |> List.toArray)
    "OpaqueXml" ==> (image.Opaque |> Option.map (fun item -> item.Xml))
    "OpaqueKind" ==> (image.Opaque |> Option.map (fun item -> item.Kind))
    "OpaqueParagraphLevel" ==> (image.Opaque |> Option.map (fun item -> item.ParagraphLevel)) ]

let private bookmarkStartToObject (start: ModelText.BookmarkStart) =
    createObj [ "Id" ==> start.Id; "Name" ==> start.Name; "EndsHere" ==> start.EndsHere ]

let private bookmarkEndToObject (ending: ModelText.BookmarkEnd) =
    match ending with
    | ModelText.Ends id -> createObj [ "Id" ==> id ]
    | ModelText.Opens start ->
        createObj [ "Id" ==> start.Id; "Name" ==> start.Name; "EndsHere" ==> start.EndsHere; "Opens" ==> true ]

let private controlLayerToObject (layer: ContentControlLayer) =
    createObj [ "Group" ==> layer.Group; "Id" ==> layer.Id; "Tag" ==> layer.Tag
                "Alias" ==> layer.Alias; "Xml" ==> layer.Xml ]

let private runToEditorObject (run: EditorRun) =
    let format = run.Formatting
    let contentFields =
        match run.Content with
        | Text text -> [ "Text" ==> text; "IsImage" ==> false ]
        | Picture picture -> ("Text" ==> "") :: imageFields picture
        | Opaque opaque ->
            [ "Text" ==> ""; "IsImage" ==> false
              "OpaqueXml" ==> opaque.Xml; "OpaqueKind" ==> opaque.Kind
              "OpaqueParagraphLevel" ==> opaque.ParagraphLevel ]
        | Equation (text, xml) -> [ "Text" ==> text; "IsImage" ==> false; "IsEquation" ==> true; "EquationXml" ==> xml ]
        | PageBreak -> [ "Text" ==> "\n"; "IsImage" ==> false ]
        | FieldBoundary boundary ->
            let kind, instruction =
                match boundary with
                | ModelText.Begin -> "begin", None
                | ModelText.Separate -> "separate", None
                | ModelText.End -> "end", None
                | ModelText.Instruction code -> "instruction", Some code
            [ "Text" ==> ""; "IsImage" ==> false
              "FieldBoundary" ==> kind; "FieldInstruction" ==> instruction ]
    // (Assigned one by one, in the order of the fields: no list of pairs built and copied per run.)
    let o = obj ()
    for key, value in contentFields do o?(key) <- value
    o?Style <- format.Style
    o?Bold <- format.Bold
    o?Italic <- format.Italic
    o?Underline <- format.Underline
    o?UnderlineStyle <- format.UnderlineStyle
    o?UnderlineColor <- format.UnderlineColor
    o?AllCaps <- format.AllCaps
    o?SmallCaps <- format.SmallCaps
    o?Strike <- format.Strike
    o?DoubleStrike <- format.DoubleStrike
    o?Hidden <- format.Hidden
    o?Position <- format.Position
    o?EmphasisMark <- format.EmphasisMark
    o?FontSize <- format.FontSize
    o?CharacterSpacing <- format.CharacterSpacing
    o?FontFamily <- format.FontFamily
    o?FontFamilyInherited <- format.FontFamilyInherited
    o?Color <- format.Color
    o?Highlight <- format.Highlight
    o?VertAlign <- format.VertAlign
    o?HasExplicit <- format.HasExplicit
    o?ExplicitBold <- format.ExplicitBold
    o?ExplicitItalic <- format.ExplicitItalic
    o?ExplicitUnderline <- format.ExplicitUnderline
    o?ExplicitAllCaps <- format.ExplicitAllCaps
    o?ExplicitSmallCaps <- format.ExplicitSmallCaps
    o?ExplicitStrike <- format.ExplicitStrike
    o?ExplicitDoubleStrike <- format.ExplicitDoubleStrike
    o?ExplicitHidden <- format.ExplicitHidden
    o?ExplicitFontSize <- format.ExplicitFontSize
    o?ExplicitColorAuto <- format.ExplicitColorAuto
    o?Hyperlink <- run.Hyperlink
    o?InactiveHyperlink <- run.InactiveHyperlink
    o?CommentId <- run.CommentId
    o?CommentText <- run.CommentText
    o?CommentAuthor <- run.CommentAuthor
    o?CommentDate <- run.CommentDate
    o?NoteId <- run.NoteId
    o?NoteType <- run.NoteType
    o?NoteText <- run.NoteText
    o?BookmarkStarts <- (run.BookmarkStarts |> List.map bookmarkStartToObject |> List.toArray)
    o?BookmarkEnds <- (run.BookmarkEnds |> List.map bookmarkEndToObject |> List.toArray)
    o?FieldCode <- run.FieldCode
    o?FieldResult <- run.FieldResult
    o?RenderedPageBreakBefore <- run.RenderedPageBreakBefore
    o?Checkbox <- run.Checkbox
    o?SourceRun <- run.SourceRun
    o?ContentControls <- (run.ContentControls |> List.map controlLayerToObject |> List.toArray)
    o?NoProps <- run.NoProps
    for key, value in revisionFields run.Revision do o?(key) <- value
    for key, value in inferredStyleFields format.StyleInfo do o?(key) <- value
    o

let private borderToObject (border: Border option) : obj =
    match border with
    | None -> null
    | Some border ->
        createObj [ "Style" ==> border.Style; "Color" ==> border.Color
                    "Width" ==> border.Width; "Space" ==> border.Space ]

let private bordersToObject (borders: Borders option) : obj =
    match borders with
    | None -> null
    | Some borders ->
        createObj [
            "Top" ==> borderToObject borders.Top; "Right" ==> borderToObject borders.Right
            "Bottom" ==> borderToObject borders.Bottom; "Left" ==> borderToObject borders.Left
            "InsideH" ==> borderToObject borders.InsideH
            "InsideV" ==> borderToObject borders.InsideV ]

let private shadingToObject (shading: Shading option) : obj =
    match shading with
    | None -> null
    | Some shade ->
        createObj [ "Pattern" ==> shade.Pattern
                    "ForegroundColor" ==> shade.ForegroundColor
                    "BackgroundColor" ==> shade.BackgroundColor ]

let private referencesToObject (refs: HeaderFooterReference list) =
    refs |> List.map (fun value ->
        createObj [ "Kind" ==> value.Kind; "RelationshipId" ==> value.RelationshipId ]) |> List.toArray

let private columnsToObject (columns: SectionColumns option) : obj =
    match columns with
    | None -> null
    | Some columns ->
        createObj [
            "Count" ==> columns.Count; "Gap" ==> columns.Gap
            "Separator" ==> columns.Separator
            "Widths" ==> (columns.Widths |> List.map (fun (w, s) -> createObj [ "W" ==> w; "S" ==> s ]) |> List.toArray) ]

let private pageSetupToObject (setup: PageSetup option) : obj =
    match setup with
    | None -> null
    | Some setup ->
        let columns =
            setup.Columns |> Option.defaultValue { Count = 1; Gap = 0.0; Separator = false; Widths = [] }
        createObj [
            "PageWidth" ==> setup.PageWidth; "PageHeight" ==> setup.PageHeight
            "MarginLeft" ==> setup.MarginLeft; "MarginRight" ==> setup.MarginRight
            "MarginTop" ==> setup.MarginTop; "MarginBottom" ==> setup.MarginBottom
            "HeaderDistance" ==> setup.HeaderDistance; "FooterDistance" ==> setup.FooterDistance
            "DifferentFirstPage" ==> setup.DifferentFirstPage
            "EvenAndOddHeaders" ==> setup.EvenAndOddHeaders
            "DefaultTabStop" ==> setup.DefaultTabStop
            "CompatibilityMode" ==> setup.CompatibilityMode
            "ColumnCount" ==> columns.Count; "ColumnSpace" ==> columns.Gap
            "ColumnSeparator" ==> columns.Separator
            "ColumnWidths" ==> (columns.Widths |> List.map (fun (w, s) -> createObj [ "W" ==> w; "S" ==> s ]) |> List.toArray)
            "SectionStart" ==> setup.SectionStart
            "HeaderReferences" ==> referencesToObject setup.HeaderReferences
            "FooterReferences" ==> referencesToObject setup.FooterReferences
            "PageNumberStart" ==> setup.PageNumberStart
            "PageNumberFormat" ==> setup.PageNumberFormat
            "VerticalAlign" ==> setup.VerticalAlign ]


let private paragraphFormatFields (format: ParagraphFormatting) =
    [
    "Style" ==> format.Style; "Alignment" ==> format.Alignment
    "HasExplicit" ==> format.HasExplicit
    "ExplicitAlignment" ==> format.ExplicitAlignment
    "IndentLeft" ==> format.IndentLeft; "IndentRight" ==> format.IndentRight
    "FirstLineIndent" ==> format.FirstLineIndent
    "RightTabPosition" ==> format.RightTabPosition
    "RightTabLeader" ==> format.RightTabLeader
    "TabStops" ==> (format.TabStops |> List.map (fun stop ->
        createObj [ "Position" ==> stop.Position; "Alignment" ==> stop.Alignment; "Leader" ==> stop.Leader ]) |> List.toArray)
    "Rtl" ==> format.Rtl; "SpacingBefore" ==> format.SpacingBefore
    "SpacingAfter" ==> format.SpacingAfter; "LineSpacing" ==> format.LineSpacing
    "LineSpacingRule" ==> format.LineSpacingRule
    "ContextualSpacing" ==> format.ContextualSpacing
    "KeepNext" ==> format.KeepNext; "KeepLines" ==> format.KeepLines
    "WidowControl" ==> format.WidowControl
    "PageBreakBefore" ==> format.PageBreakBefore
    "PageBreakBeforeDerived" ==> format.PageBreakBeforeDerived
    "PageBreakAtStart" ==> format.PageBreakAtStart
    "PageBreakAfter" ==> format.PageBreakAfter
    "ColumnBreakAfter" ==> format.ColumnBreakAfter
    "RenderedPageBreakBefore" ==> format.RenderedPageBreakBefore
    "BackgroundColor" ==> format.BackgroundColor
    "ParagraphShadingPattern" ==> format.ShadingPattern
    "ParagraphShadingForeground" ==> format.ShadingForeground
    "ParagraphBorders" ==> bordersToObject format.Borders
    "ParagraphBorderBefore" ==> format.ParagraphBorderBefore
    "ParagraphBorderAfter" ==> format.ParagraphBorderAfter
    "OutlineLevel" ==> format.OutlineLevel
    "ExplicitParagraphFormat" ==> format.ExplicitParagraphFormat
    "ExplicitIndentLeft" ==> format.ExplicitIndentLeft
    "ExplicitIndentRight" ==> format.ExplicitIndentRight
    "ExplicitFirstLineIndent" ==> format.ExplicitFirstLineIndent
    "ExplicitSpacingBefore" ==> format.ExplicitSpacingBefore
    "ExplicitSpacingAfter" ==> format.ExplicitSpacingAfter
    "ExplicitLineSpacing" ==> format.ExplicitLineSpacing
    "ExplicitContextualSpacing" ==> format.ExplicitContextualSpacing
    "ExplicitBackgroundColor" ==> format.ExplicitBackgroundColor
    "ExplicitParagraphShadingPattern" ==> format.ExplicitShadingPattern
    "ExplicitParagraphShadingForeground" ==> format.ExplicitShadingForeground
    "ExplicitTabs" ==> format.ExplicitTabs
    "HasMarkProperties" ==> format.HasMarkProperties
    "ExplicitMarkFontSize" ==> format.ExplicitMarkFontSize
    "ParagraphMarkFontSize" ==> format.ParagraphMarkFontSize
    "MarkRevisionType" ==> (format.MarkRevision |> Option.bind (fun mark -> mark.Kind))
    "MarkRevisionId" ==> (format.MarkRevision |> Option.bind (fun mark -> mark.Id))
    "MarkRevisionAuthor" ==> (format.MarkRevision |> Option.bind (fun mark -> mark.Author))
    "MarkRevisionDate" ==> (format.MarkRevision |> Option.bind (fun mark -> mark.Date))
    "DropCapKind" ==> format.DropCapKind
    "DropCapLines" ==> format.DropCapLines
    "DropCapDistance" ==> format.DropCapDistance
    "ContentControlId" ==> format.ContentControlId
    "ContentControlTag" ==> format.ContentControlTag
    "ContentControlAlias" ==> format.ContentControlAlias
    "ContentControlXml" ==> format.ContentControlXml
    "ContentControlGroup" ==> format.ContentControlGroup
    "ContentControlAncestors" ==> (format.ContentControlAncestors |> List.map controlLayerToObject |> List.toArray)
    "EnclosingCellControls" ==> (format.EnclosingCellControls |> List.map controlLayerToObject |> List.toArray)
    ] @ inferredStyleFields format.StyleInfo

let private paragraphListFields (value: ParagraphList option) =
    match value with
    | None -> [ "ListStyle" ==> "none" ]
    | Some item -> [
        "ListStyle" ==> item.Style; "ListLevel" ==> item.Level
        "ListId" ==> item.Id; "ListSourceId" ==> item.SourceId
        "ListRestart" ==> item.Restart; "ListFormatEdited" ==> item.FormatEdited
        "ListNumberFormat" ==> item.NumberFormat; "ListLevelText" ==> item.LevelText
        "ListStart" ==> item.Start
        "ListStartOverrides" ==> (item.StartOverrides |> List.map (fun (level, start) ->
            createObj [ "Level" ==> level; "Start" ==> start ]) |> List.toArray)
        "HasListLevelIndents" ==> item.ListLevelIndents.IsSome
        "ListLevelIndentLeft" ==> (item.ListLevelIndents |> Option.bind fst)
        "ListLevelFirstLineIndent" ==> (item.ListLevelIndents |> Option.bind snd)
        "ListLevels" ==> (item.Levels |> List.map (fun level ->
            createObj [
                "Level" ==> level.Level; "NumberFormat" ==> level.NumberFormat
                "LevelText" ==> level.LevelText; "Start" ==> level.Start
                "IndentLeft" ==> level.IndentLeft; "FirstLineIndent" ==> level.FirstLineIndent
                "Suffix" ==> level.Suffix; "Justification" ==> level.Justification
                "Legal" ==> level.Legal; "Restart" ==> level.Restart
                "MarkFont" ==> level.MarkFont; "MarkBold" ==> level.MarkBold
                "MarkItalic" ==> level.MarkItalic; "MarkColor" ==> level.MarkColor
                "MarkSize" ==> level.MarkSize ]) |> List.toArray) ]

let rec private paragraphToEditorObject (paragraph: EditorParagraph) : obj =
    // (Assigned one by one, in the order of the fields: no list of pairs built and copied per paragraph.)
    let o = obj ()
    o?Type <- (match paragraph.HeadingLevel with Some _ -> "Heading" | None -> "Paragraph")
    o?HeadingLevel <- paragraph.HeadingLevel
    o?Runs <- (paragraph.Runs |> Array.map runToEditorObject)
    o?SectionBreak <- paragraph.SectionBreak
    o?SectionPageSetup <- pageSetupToObject paragraph.SectionPageSetup
    o?SectionColumns <- columnsToObject paragraph.SectionColumns
    o?SectionVerticalAlign <- paragraph.SectionVerticalAlign
    o?SectionPageNumberStart <- paragraph.SectionPageNumberStart
    o?SectionPageNumberFormat <- paragraph.SectionPageNumberFormat
    o?HeaderReferences <- referencesToObject paragraph.HeaderReferences
    o?FooterReferences <- referencesToObject paragraph.FooterReferences
    o?SourcePara <- (paragraph.SourceParagraph |> Option.map paragraphEditToObject)
    o?SpliceOriginal <- (paragraph.SpliceOriginal |> Option.map paragraphToEditorObject)
    o?SpliceFrom <- (paragraph.SpliceFrom |> Option.map (fun (start, finish) ->
              createObj [ "Start" ==> start; "End" ==> finish ]))
    o?SpliceLeading <- paragraph.SpliceLeading
    o?SpliceTrailing <- paragraph.SpliceTrailing
    o?SpliceJoined <- (paragraph.SpliceJoined |> Array.map paragraphToEditorObject)
    o?SourceIndex <- paragraph.SourceIndex
    o?SourceUnchanged <- paragraph.SourceUnchanged
    o?SourceWhy <- paragraph.SourceWhy
    o?SourceTextChanges <- (paragraph.SourceTextChanges |> Option.map sourceChangesToObject)
    for key, value in sourceFields paragraph.Source do o?(key) <- value
    for key, value in paragraphFormatFields paragraph.Formatting do o?(key) <- value
    for key, value in paragraphListFields paragraph.List do o?(key) <- value
    o

and private paragraphEditToObject (edit: ParagraphEdit) : obj =
    createObj [
        "Paragraph" ==> edit.Paragraph; "Block" ==> edit.Block
        "Para" ==> edit.Para; "OldText" ==> edit.OldText
        "Text" ==> (edit.Text |> Option.map (fun change ->
            createObj [
                "Paragraph" ==> change.Paragraph; "Offset" ==> change.Offset
                "Deleted" ==> change.Deleted; "Inserted" ==> change.Inserted
                "JoinBefore" ==> change.JoinBefore; "JoinAfter" ==> change.JoinAfter
                "RunsPlain" ==> change.RunsPlain; "Witnesses" ==> change.Witnesses
                "MarkStyle" ==> change.MarkStyle; "NewText" ==> change.NewText ]))
        "Before" ==> (edit.Before |> Option.map paragraphToEditorObject)
        "After" ==> (edit.After |> Option.map paragraphToEditorObject)
        "NewText" ==> edit.NewText; "Regenerate" ==> edit.Regenerate
        "Verbatim" ==> edit.Verbatim; "Why" ==> edit.Why
        "Cut" ==> (edit.Cut |> Option.map (fun (start, finish) -> [| start; finish |]))
        "Join" ==> (edit.Join |> Array.map (fun joined ->
            createObj [
                "Block" ==> joined.Block; "Para" ==> joined.Para
                "Text" ==> joined.Text; "Original" ==> paragraphToEditorObject joined.Original ])) ]

and private sourceChangesToObject (changes: SourceTextChanges) : obj =
    createObj [
        "Texts" ==> changes.Texts
        "Edits" ==> (changes.Edits |> Array.map paragraphEditToObject) ]

and private blockToEditorObject (block: EditorBlock) : obj =
    match block with
    | Paragraph para -> paragraphToEditorObject para
    | Table table -> tableToEditorObject table
    | TextBox box -> textBoxToEditorObject box

and private sourceTableToObject (source: TableSource option) : obj =
    match source with
    | None -> null
    | Some source ->
        createObj [
            "At" ==> source.At; "Block" ==> source.Block; "Copy" ==> source.Copy; "Same" ==> source.Same
            "Before" ==> (source.Before |> Option.map tableToEditorObject)
            "Was" ==> (source.Was |> Option.map tableToEditorObject) ]

and private sourceRowToObject (source: RowSource option) : obj =
    match source with
    | None -> null
    | Some source ->
        createObj [
            "At" ==> source.At; "Copy" ==> source.Copy; "Same" ==> source.Same
            "Before" ==> (source.Before |> Option.map rowToEditorObject)
            "Was" ==> (source.Was |> Option.map rowToEditorObject) ]

and private sourceCellToObject (source: CellSource option) : obj =
    match source with
    | None -> null
    | Some source ->
        createObj [
            "At" ==> source.At; "Copy" ==> source.Copy; "Same" ==> source.Same
            "Before" ==> (source.Before |> Option.map cellToEditorObject)
            "Was" ==> (source.Was |> Option.map cellToEditorObject) ]

and private tableToEditorObject (table: EditorTable) : obj =
    createObj (
        [ "Type" ==> "Table"
          "Rows" ==> (table.Rows |> Array.map rowToEditorObject)
          "PreferredWidth" ==> table.PreferredWidth
          "PreferredWidthType" ==> table.PreferredWidthType
          "TableAlignment" ==> table.Alignment; "TableIndent" ==> table.Indent
          "LayoutMode" ==> table.LayoutMode; "GridColumns" ==> table.GridColumns
          "Borders" ==> bordersToObject table.Borders
          "BackgroundColor" ==> table.BackgroundColor
          "Shading" ==> shadingToObject table.Shading
          "CellSpacing" ==> table.CellSpacing
          "PaddingTop" ==> table.PaddingTop; "PaddingRight" ==> table.PaddingRight
          "PaddingBottom" ==> table.PaddingBottom; "PaddingLeft" ==> table.PaddingLeft
          "Floating" ==> table.Floating; "FloatingXml" ==> table.FloatingXml
          "TableStyleId" ==> table.StyleId
          "TableLook" ==> (table.Look |> Option.map (fun look -> createObj [
              "FirstRow" ==> look.FirstRow; "LastRow" ==> look.LastRow
              "FirstColumn" ==> look.FirstColumn; "LastColumn" ==> look.LastColumn
              "HorizontalBanding" ==> look.HorizontalBanding
              "VerticalBanding" ==> look.VerticalBanding ]))
          "ContentControlId" ==> table.ContentControlId
          "ContentControlTag" ==> table.ContentControlTag
          "ContentControlAlias" ==> table.ContentControlAlias
          "ContentControlXml" ==> table.ContentControlXml
          "ContentControlGroup" ==> table.ContentControlGroup
          "ContentControlAncestors" ==> (table.ContentControlAncestors |> List.map controlLayerToObject |> List.toArray)
          "SourceTable" ==> sourceTableToObject table.SourceTable
          "SourceIndex" ==> table.SourceIndex
          "SourceUnchanged" ==> table.SourceUnchanged
          "SourceWhy" ==> table.SourceWhy
          "SourceTextChanges" ==> (table.SourceTextChanges |> Option.map sourceChangesToObject) ]
        @ sourceFields table.Source)

and private rowToEditorObject (row: EditorRow) : obj =
    createObj (
        [ "Cells" ==> (row.Cells |> Array.map cellToEditorObject)
          "IsHeader" ==> row.IsHeader; "Height" ==> row.Height
          "HeightRule" ==> row.HeightRule; "CanSplit" ==> row.CanSplit
          "CellSpacing" ==> row.CellSpacing
          "PaddingTop" ==> row.PaddingTop; "PaddingRight" ==> row.PaddingRight
          "PaddingBottom" ==> row.PaddingBottom; "PaddingLeft" ==> row.PaddingLeft
          "GridBefore" ==> row.GridBefore; "GridAfter" ==> row.GridAfter
          "WidthBefore" ==> row.WidthBefore; "WidthAfter" ==> row.WidthAfter
          "BordersOverride" ==> bordersToObject row.BordersOverride
          "SourceRow" ==> sourceRowToObject row.SourceRow ]
        @ sourceFields row.Source)

and private cellToEditorObject (cell: EditorCell) : obj =
    createObj (
        [ "Paragraphs" ==> (cell.Paragraphs |> Array.map blockToEditorObject)
          "NestedTables" ==> (cell.NestedTables |> Array.map tableToEditorObject)
          "NestedTableAnchors" ==> cell.NestedTableAnchors
          "EnclosingControls" ==> (cell.EnclosingControls |> List.map controlLayerToObject |> List.toArray)
          "ColSpan" ==> cell.ColSpan
          "BackgroundColor" ==> cell.BackgroundColor
          "Shading" ==> shadingToObject cell.Shading
          "Width" ==> cell.Width; "PreferredWidth" ==> cell.PreferredWidth
          "PreferredWidthType" ==> cell.PreferredWidthType
          "BorderColor" ==> cell.BorderColor; "BorderWidth" ==> cell.BorderWidth
          "Borders" ==> bordersToObject cell.Borders
          "VerticalAlignment" ==> cell.VerticalAlignment
          "VerticalMerge" ==> cell.VerticalMerge
          "PaddingTop" ==> cell.PaddingTop; "PaddingRight" ==> cell.PaddingRight
          "PaddingBottom" ==> cell.PaddingBottom; "PaddingLeft" ==> cell.PaddingLeft
          "SourceCell" ==> sourceCellToObject cell.SourceCell ]
        @ sourceFields cell.Source)

and private textBoxToEditorObject (box: EditorTextBox) : obj =
    createObj [
        "Type" ==> "TextBox"
        "Paragraphs" ==> (box.Paragraphs |> Array.map blockToEditorObject)
        "Children" ==> (box.Children |> Array.map textBoxToEditorObject)
        "ShapeId" ==> box.ShapeId; "ShapeKind" ==> box.ShapeKind
        "Name" ==> box.Name; "Geometry" ==> box.Geometry
        "CustomPath" ==> (box.CustomPath |> List.map (fun (op, points) ->
            createObj [ "Op" ==> op; "P" ==> points ]) |> List.toArray)
        "CustomW" ==> box.CustomWidth; "CustomH" ==> box.CustomHeight
        "X" ==> box.X; "Y" ==> box.Y
        "PositionHRelativeFrom" ==> box.PositionHRelativeFrom
        "PositionVRelativeFrom" ==> box.PositionVRelativeFrom
        "PositionHAlign" ==> box.PositionHAlign
        "PositionVAlign" ==> box.PositionVAlign
        "Width" ==> box.Width; "Height" ==> box.Height
        "FillColor" ==> box.FillColor; "BorderColor" ==> box.BorderColor
        "NoFill" ==> box.NoFill; "NoLine" ==> box.NoLine
        "BorderWidth" ==> box.BorderWidth; "BorderDash" ==> box.BorderDash
        "Gradient" ==> (if box.Gradient.IsEmpty then null
                         else createObj [
                             "Stops" ==> (box.Gradient |> List.map (fun (offset, color) ->
                                 createObj [ "Offset" ==> offset; "Color" ==> color ]) |> List.toArray)
                             "Angle" ==> box.GradientAngle
                             "Radial" ==> box.GradientRadial ])
        "Rotation" ==> box.Rotation; "FlipH" ==> box.FlipH; "FlipV" ==> box.FlipV
        "Inline" ==> box.Inline; "Wrap" ==> box.Wrap
        "BehindDoc" ==> box.BehindDoc; "AllowOverlap" ==> box.AllowOverlap
        "LayoutInCell" ==> box.LayoutInCell
        "DistanceTop" ==> box.DistanceTop; "DistanceBottom" ==> box.DistanceBottom
        "DistanceLeft" ==> box.DistanceLeft; "DistanceRight" ==> box.DistanceRight
        "RelativeHeight" ==> box.RelativeHeight
        "AnchorBlock" ==> box.AnchorBlock
        "InsetLeft" ==> box.InsetLeft; "InsetTop" ==> box.InsetTop
        "InsetRight" ==> box.InsetRight; "InsetBottom" ==> box.InsetBottom
        "VerticalAnchor" ==> box.VerticalAnchor
        "AutoFit" ==> box.AutoFit; "NoTextWrap" ==> box.NoTextWrap
        "FitText" ==> box.FitText; "Opacity" ==> box.Opacity
        "ChildSpaceWidth" ==> (box.ChildSpace |> Option.map fst)
        "ChildSpaceHeight" ==> (box.ChildSpace |> Option.map snd)
        "ImageSrc" ==> box.ImageSource; "SourceXml" ==> box.SourceXml
        "SourceMoved" ==> (box.SourceMoved |> Option.map (fun movement ->
            createObj [ "H" ==> movement.H; "V" ==> movement.V; "Size" ==> movement.Size ]))
        "SourceIndex" ==> box.SourceIndex
        "SourceUnchanged" ==> box.SourceUnchanged
        "SourceWhy" ==> box.SourceWhy
        "SourceTextChanges" ==> (box.SourceTextChanges |> Option.map sourceChangesToObject) ]

let private styleToEditorObject (style: Style) =
    let run =
        match style.Run with
        | None -> []
        | Some r -> [
            "StyleBold" ==> r.Bold; "StyleItalic" ==> r.Italic
            "StyleUnderline" ==> r.Underline
            "StyleUnderlineStyle" ==> r.UnderlineStyle
            "StyleUnderlineColor" ==> r.UnderlineColor
            "StyleAllCaps" ==> r.AllCaps; "StyleSmallCaps" ==> r.SmallCaps
            "StyleStrike" ==> r.Strike; "StyleDoubleStrike" ==> r.DoubleStrike
            "StyleHidden" ==> r.Hidden; "StylePosition" ==> r.Position
            "StyleEmphasisMark" ==> r.EmphasisMark
            "StyleVertAlign" ==> r.VertAlign; "StyleFontSize" ==> r.FontSize
            "StyleCharacterSpacing" ==> r.CharacterSpacing
            "StyleFontFamily" ==> r.FontFamily
            "StyleColor" ==> r.Color; "StyleHighlight" ==> r.Highlight ]
    let paragraph =
        match style.Paragraph with
        | None -> []
        | Some p -> [
            "StyleSpacingBefore" ==> p.SpacingBefore
            "StyleSpacingAfter" ==> p.SpacingAfter
            "StyleContextualSpacing" ==> p.ContextualSpacing
            "StyleLineSpacing" ==> p.LineSpacing
            "StyleLineRule" ==> p.LineSpacingRule
            "StyleAlignment" ==> p.Alignment
            "StyleIndentLeft" ==> p.IndentLeft
            "StyleIndentRight" ==> p.IndentRight
            "StyleFirstLineIndent" ==> p.FirstLineIndent
            "StyleBackgroundColor" ==> p.BackgroundColor
            "StyleOutlineLevel" ==> p.OutlineLevel
            "StyleKeepNext" ==> p.KeepNext
            "StyleKeepLines" ==> p.KeepLines
            "StylePageBreakBefore" ==> p.PageBreakBefore
            "StyleWidowControl" ==> p.WidowControl ]
    createObj ([
        "StyleId" ==> style.Id; "Name" ==> style.Name
        "StyleType" ==> style.StyleType; "IsDefaultStyle" ==> style.IsDefault
        "BasedOn" ==> style.BasedOn; "NextStyle" ==> style.Next
        "LinkedStyle" ==> style.Linked
        "IsQuickStyle" ==> style.IsQuickStyle
        "StyleAliases" ==> style.Aliases
        "StyleUiPriority" ==> style.UiPriority
        "StyleSemiHidden" ==> style.SemiHidden
        "StyleUnhideWhenUsed" ==> style.UnhideWhenUsed
    ] @ run @ paragraph)

let packageToEditorObject (package: EditorPackage) : obj =
    let parts (items: HeaderFooterPart array) =
        items |> Array.map (fun item ->
            createObj [
                "Kind" ==> item.Kind
                "RelationshipId" ==> item.RelationshipId
                "Paragraphs" ==> (item.Blocks |> Array.map blockToEditorObject)
                "SourceStoryCounts" ==> item.SourceStoryCounts
                "SourceUnchanged" ==> item.SourceUnchanged ])
    let notes (items: Story array) =
        items |> Array.map (fun item ->
            createObj [
                "Id" ==> item.Id; "Kind" ==> item.Kind
                "HasMark" ==> item.HasMark; "MarkStyle" ==> item.MarkStyle
                "Text" ==> item.Text
                "Blocks" ==> (item.Blocks |> Array.map blockToEditorObject)
                "SourceStoryCounts" ==> item.SourceStoryCounts
                "SourceUnchanged" ==> item.Unchanged ])
    let comments (items: CommentPart array) =
        items |> Array.map (fun item ->
            createObj [
                "Id" ==> item.Id; "Author" ==> item.Author; "Date" ==> item.Date
                "Initials" ==> item.Initials; "Text" ==> item.Text
                "Blocks" ==> (item.Blocks |> Array.map blockToEditorObject)
                "SourceStoryCounts" ==> item.SourceStoryCounts
                "SourceUnchanged" ==> item.Unchanged ])
    createObj [
        "Paragraphs" ==> (package.Blocks |> Array.map blockToEditorObject)
        "Styles" ==> (package.Styles |> Array.map styleToEditorObject)
        "HeaderParts" ==> parts package.HeaderParts
        "FooterParts" ==> parts package.FooterParts
        "Footnotes" ==> notes package.Footnotes
        "Endnotes" ==> notes package.Endnotes
        "Comments" ==> comments package.Comments
        "Theme" ==> (package.Theme |> Option.map (fun theme ->
            createObj [
                "MajorFont" ==> theme.MajorFont
                "MinorFont" ==> theme.MinorFont
                "Colors" ==> (theme.Colors |> Map.toList |> List.map (fun (key, color) -> key ==> color) |> createObj) ]))
        "PageSetup" ==> pageSetupToObject package.PageSetup
        "OriginalPageSetup" ==> pageSetupToObject package.OriginalPageSetup
        "Preserved" ==> preservedToEditorObject package.Preserved
        "SourceStoryCounts" ==> package.SourceStoryCounts
        "ChangedStyles" ==> package.ChangedStyles
        "ChangedListLevels" ==> package.ChangedListLevels
        "PageSetupUnchanged" ==> package.PageSetupUnchanged
        "ListsUnchanged" ==> package.ListsUnchanged
        "FootnotesUnchanged" ==> package.FootnotesUnchanged
        "EndnotesUnchanged" ==> package.EndnotesUnchanged
        "CommentsUnchanged" ==> package.CommentsUnchanged
        "Unchanged" ==> package.Unchanged
        "RegenerateAll" ==> package.RegenerateAll ]

// (Inline: each read is a property of a literal name at its call site - a load the engine caches
// per site - not one function reading every property by a variable name.)
let inline private get<'a> (value: obj) (name: string) : 'a option =
    if isMissing value then None
    else memberOption<'a> (value?(name))

let inline private textOf value name = get<string> value name
let inline private text value name = textOf value name |> Option.defaultValue ""
let inline private boolean value name = get<bool> value name |> Option.defaultValue false
[<Emit("Number($0)")>]
let private jsNumber (_value: obj) : float = jsNative
let inline private number value name =
    get<obj> value name
    |> Option.bind (fun raw ->
        let parsed = jsNumber raw
        if System.Double.IsNaN parsed || System.Double.IsInfinity parsed then None else Some parsed)
let inline private amount value name fallback = number value name |> Option.defaultValue fallback
let inline private integer value name = get<int> value name
let inline private intValue value name fallback = integer value name |> Option.defaultValue fallback
let inline private array value name = get<obj array> value name |> Option.defaultValue [||]

let private readSource value =
    let copy = boolean value "SourceCopy"
    { Origin =
        integer value "SourceAt"
        |> Option.map (fun at -> if copy then LikeAt at else ReadAt at)
      Xml = textOf value "SourceXml"
      Copy = copy
      Same = boolean value "SourceSame" }

let private readBorder value =
    if isMissing value then None
    else Some {
        Style = textOf value "Style"; Color = textOf value "Color"
        Width = number value "Width"; Space = number value "Space" }

let private readBorders value =
    if isMissing value then None
    else Some {
        Top = readBorder value?Top; Right = readBorder value?Right
        Bottom = readBorder value?Bottom; Left = readBorder value?Left
        InsideH = readBorder value?InsideH; InsideV = readBorder value?InsideV }

let private readShading value =
    if isMissing value then None
    else Some {
        Pattern = textOf value "Pattern"
        ForegroundColor = textOf value "ForegroundColor"
        BackgroundColor = textOf value "BackgroundColor" }

let private readRefs value name =
    array value name
    |> Array.choose (fun reference ->
        textOf reference "RelationshipId"
        |> Option.map (fun id -> {
            Kind = textOf reference "Kind" |> Option.orElse (textOf reference "Type") |> Option.defaultValue "default"
            RelationshipId = id }))
    |> List.ofArray

let private readColumns value =
    if isMissing value then None
    else Some {
        Count = intValue value "Count" 1
        Gap = amount value "Gap" 0.0
        Separator = boolean value "Separator"
        Widths = array value "Widths" |> Array.map (fun column ->
            amount column "W" 0.0, amount column "S" 0.0) |> List.ofArray }

let private readPageSetup value =
    if isMissing value then None
    else
        let columns =
            if get<obj> value "Columns" |> Option.isSome then readColumns value?Columns
            elif intValue value "ColumnCount" 1 > 1 || (array value "ColumnWidths").Length > 0 then
                Some {
                    Count = intValue value "ColumnCount" 1
                    Gap = amount value "ColumnSpace" 0.0
                    Separator = boolean value "ColumnSeparator"
                    Widths = array value "ColumnWidths" |> Array.map (fun column ->
                        amount column "W" 0.0, amount column "S" 0.0) |> List.ofArray }
            else None
        Some {
            PageWidth = amount value "PageWidth" 816.0
            PageHeight = amount value "PageHeight" 1056.0
            MarginLeft = amount value "MarginLeft" 96.0
            MarginRight = amount value "MarginRight" 96.0
            MarginTop = amount value "MarginTop" 96.0
            MarginBottom = amount value "MarginBottom" 96.0
            HeaderDistance = amount value "HeaderDistance" 48.0
            FooterDistance = amount value "FooterDistance" 48.0
            DifferentFirstPage = boolean value "DifferentFirstPage"
            EvenAndOddHeaders = boolean value "EvenAndOddHeaders"
            DefaultTabStop = number value "DefaultTabStop"
            CompatibilityMode = integer value "CompatibilityMode"
            Columns = columns
            SectionStart = textOf value "SectionStart"
            HeaderReferences = readRefs value "HeaderReferences"
            FooterReferences = readRefs value "FooterReferences"
            PageNumberStart = integer value "PageNumberStart"
            PageNumberFormat = textOf value "PageNumberFormat"
            VerticalAlign = textOf value "VerticalAlign" }

let private styleRunFieldNames = [|
    "StyleBold"; "StyleItalic"; "StyleUnderline"; "StyleUnderlineStyle"; "StyleUnderlineColor"
    "StyleAllCaps"; "StyleSmallCaps"; "StyleStrike"; "StyleDoubleStrike"; "StyleHidden"; "StylePosition"
    "StyleEmphasisMark"; "StyleVertAlign"; "StyleCharacterSpacing"; "StyleFontSize"; "StyleFontFamily"
    "StyleColor"; "StyleHighlight"
|]

let private styleParagraphFieldNames = [|
    "StyleSpacingBefore"; "StyleSpacingAfter"; "StyleContextualSpacing"; "StyleLineSpacing"
    "StyleLineRule"; "StyleLineSpacingRule"; "StyleAlignment"; "StyleIndentLeft"; "StyleIndentRight"
    "StyleFirstLineIndent"; "StyleBackgroundColor"; "StyleOutlineLevel"
    "StyleKeepNext"; "StyleKeepLines"; "StylePageBreakBefore"; "StyleWidowControl"
|]

let private inferredStyleFieldNames =
    Array.concat [|
        [| "StyleName"; "IsDefaultStyle"; "IsDefault"; "BasedOnStyle"; "NextStyle"; "LinkedStyle"
           "IsQuickStyle"; "StyleAliases"; "Aliases"; "StyleUiPriority"; "StyleSemiHidden"; "StyleUnhideWhenUsed" |]
        styleRunFieldNames
        styleParagraphFieldNames
    |]

let private hasAnyStyleField (names: string array) value =
    let mutable found = false
    let mutable index = 0
    while not found && index < names.Length do
        found <- (get<obj> value names.[index]).IsSome
        index <- index + 1
    found

let private readInferredStyle value =
    if not (hasAnyStyleField inferredStyleFieldNames value) then None
    else Some {
        Name = textOf value "StyleName"
        IsDefault = get<bool> value "IsDefaultStyle" |> Option.orElse (get<bool> value "IsDefault")
        BasedOn = textOf value "BasedOnStyle"
        Next = textOf value "NextStyle"
        Linked = textOf value "LinkedStyle"
        IsQuickStyle = get<bool> value "IsQuickStyle"
        Aliases = textOf value "StyleAliases" |> Option.orElse (textOf value "Aliases")
        UiPriority = integer value "StyleUiPriority"
        SemiHidden = get<bool> value "StyleSemiHidden"
        UnhideWhenUsed = get<bool> value "StyleUnhideWhenUsed"
        Bold = get<bool> value "StyleBold"
        Italic = get<bool> value "StyleItalic"
        Underline = get<bool> value "StyleUnderline"
        UnderlineStyle = textOf value "StyleUnderlineStyle"
        UnderlineColor = textOf value "StyleUnderlineColor"
        AllCaps = get<bool> value "StyleAllCaps"
        SmallCaps = get<bool> value "StyleSmallCaps"
        Strike = get<bool> value "StyleStrike"
        DoubleStrike = get<bool> value "StyleDoubleStrike"
        Hidden = get<bool> value "StyleHidden"
        Position = number value "StylePosition"
        EmphasisMark = textOf value "StyleEmphasisMark"
        VertAlign = textOf value "StyleVertAlign"
        CharacterSpacing = number value "StyleCharacterSpacing"
        FontSize = number value "StyleFontSize"
        FontFamily = textOf value "StyleFontFamily"
        Color = textOf value "StyleColor"
        Highlight = textOf value "StyleHighlight"
        SpacingBefore = number value "StyleSpacingBefore"
        SpacingAfter = number value "StyleSpacingAfter"
        ContextualSpacing = get<bool> value "StyleContextualSpacing"
        LineSpacing = number value "StyleLineSpacing"
        LineRule = textOf value "StyleLineRule" |> Option.orElse (textOf value "StyleLineSpacingRule")
        Alignment = textOf value "StyleAlignment"
        IndentLeft = number value "StyleIndentLeft"
        IndentRight = number value "StyleIndentRight"
        FirstLineIndent = number value "StyleFirstLineIndent"
        BackgroundColor = textOf value "StyleBackgroundColor"
        OutlineLevel = integer value "StyleOutlineLevel"
        KeepNext = get<bool> value "StyleKeepNext"
        KeepLines = get<bool> value "StyleKeepLines"
        PageBreakBefore = get<bool> value "StylePageBreakBefore"
        WidowControl = get<bool> value "StyleWidowControl" }

let private readRevision value : Revision option =
    let kind = textOf value "RevisionType"
    let id = textOf value "RevisionId"
    let inserted : DocxReader.RevisionMark option =
        let id = textOf value "InsertedRevisionId"
        let author = textOf value "InsertedRevisionAuthor"
        let date = textOf value "InsertedRevisionDate"
        if id.IsSome || author.IsSome || date.IsSome then
            Some { Id = id; Author = author; Date = date }
        else None
    if kind.IsNone && id.IsNone && inserted.IsNone then None
    else Some {
        Id = id; Kind = kind
        Author = textOf value "RevisionAuthor"
        Date = textOf value "RevisionDate"
        Inserted = inserted
        MoveName = textOf value "RevisionMoveName" }

let private readImage value : Image =
    let opaque : DocxReader.OpaqueObject option =
        textOf value "OpaqueXml"
        |> Option.map (fun xml -> {
            Xml = xml
            Kind = textOf value "OpaqueKind" |> Option.defaultValue "drawing"
            ParagraphLevel = boolean value "OpaqueParagraphLevel"
            ShapeCandidate = false })
    let polygon = array value "ImageWrapPolygon" |> Array.map unbox<float>
    let wrapPolygon =
        [ for i in 0 .. (polygon.Length / 2) - 1 -> polygon.[2 * i], polygon.[2 * i + 1] ]
    { Source = textOf value "ImageSrc"
      Linked = boolean value "ImageLinked"
      Width = number value "ImageWidth"
      Height = number value "ImageHeight"
      Alt = textOf value "ImageAlt"
      Name = textOf value "ImageName"
      Title = textOf value "ImageTitle"
      Floating = boolean value "ImageFloating"
      Wrap = textOf value "ImageWrap"
      PositionX = number value "ImagePositionX"
      PositionY = number value "ImagePositionY"
      PositionHRelativeFrom = textOf value "ImagePositionHRelativeFrom"
      PositionVRelativeFrom = textOf value "ImagePositionVRelativeFrom"
      PositionHAlign = textOf value "ImagePositionHAlign"
      PositionVAlign = textOf value "ImagePositionVAlign"
      BehindDoc = get<bool> value "ImageBehindDoc"
      AllowOverlap = get<bool> value "ImageAllowOverlap"
      LayoutInCell = get<bool> value "ImageLayoutInCell"
      DistanceTop = number value "ImageDistanceTop"
      DistanceBottom = number value "ImageDistanceBottom"
      DistanceLeft = number value "ImageDistanceLeft"
      DistanceRight = number value "ImageDistanceRight"
      RelativeHeight = textOf value "ImageRelativeHeight"
      Locked = get<bool> value "ImageLocked"
      SimplePosition = get<bool> value "ImageSimplePosition"
      SimplePositionX = number value "ImageSimplePositionX"
      SimplePositionY = number value "ImageSimplePositionY"
      CropLeft = textOf value "ImageCropLeft"
      CropTop = textOf value "ImageCropTop"
      CropRight = textOf value "ImageCropRight"
      CropBottom = textOf value "ImageCropBottom"
      WrapText = textOf value "ImageWrapText"
      WrapPolygon = wrapPolygon
      FlipH = boolean value "ImageFlipH"
      FlipV = boolean value "ImageFlipV"
      Opaque = opaque }

let private readRunFormat value = {
    Style = textOf value "Style" |> Option.orElse (textOf value "RunStyle") |> Option.orElse (textOf value "CharacterStyle")
    StyleInfo = readInferredStyle value
    Bold = boolean value "Bold"
    Italic = boolean value "Italic"
    Underline = boolean value "Underline"
    UnderlineStyle = textOf value "UnderlineStyle"
    UnderlineColor = textOf value "UnderlineColor"
    AllCaps = boolean value "AllCaps"
    SmallCaps = boolean value "SmallCaps"
    Strike = boolean value "Strike"
    DoubleStrike = boolean value "DoubleStrike"
    Hidden = boolean value "Hidden"
    Position = amount value "Position" 0.0
    EmphasisMark = textOf value "EmphasisMark"
    FontSize = amount value "FontSize" (10.0 * 96.0 / 72.0)
    CharacterSpacing = amount value "CharacterSpacing" 0.0
    FontFamily = textOf value "FontFamily" |> Option.defaultValue "Calibri"
    FontFamilyInherited = get<bool> value "FontFamilyInherited"
    Color = textOf value "Color"
    Highlight = textOf value "Highlight"
    VertAlign = textOf value "VertAlign"
    HasExplicit = boolean value "HasExplicit"
    ExplicitBold = get<bool> value "ExplicitBold"
    ExplicitItalic = get<bool> value "ExplicitItalic"
    ExplicitUnderline = get<bool> value "ExplicitUnderline"
    ExplicitAllCaps = get<bool> value "ExplicitAllCaps"
    ExplicitSmallCaps = get<bool> value "ExplicitSmallCaps"
    ExplicitStrike = get<bool> value "ExplicitStrike"
    ExplicitDoubleStrike = get<bool> value "ExplicitDoubleStrike"
    ExplicitHidden = get<bool> value "ExplicitHidden"
    ExplicitFontSize = number value "ExplicitFontSize"
    ExplicitColorAuto = boolean value "ExplicitColorAuto" }

let private readBookmarkStart value : ModelText.BookmarkStart =
    { Id = text value "Id"; Name = text value "Name"; EndsHere = boolean value "EndsHere" }

let private readBookmarkEnd value : ModelText.BookmarkEnd =
    if boolean value "Opens" then ModelText.Opens (readBookmarkStart value)
    else ModelText.Ends (text value "Id")

let private readControlLayer value : ContentControlLayer =
    { Group = text value "Group"; Id = textOf value "Id"; Tag = textOf value "Tag"
      Alias = textOf value "Alias"; Xml = textOf value "Xml" }

let private readRun value =
    let content =
        match textOf value "FieldBoundary" with
        | Some "begin" -> FieldBoundary ModelText.Begin
        | Some "separate" -> FieldBoundary ModelText.Separate
        | Some "end" -> FieldBoundary ModelText.End
        | Some "instruction" -> FieldBoundary (ModelText.Instruction (text value "FieldInstruction"))
        | _ ->
            match textOf value "OpaqueXml", boolean value "IsImage", boolean value "IsEquation" with
            | Some _, true, _ -> Picture (readImage value)
            | Some xml, false, _ ->
                Opaque {
                    DocxReader.OpaqueObject.Xml = xml
                    Kind = textOf value "OpaqueKind" |> Option.defaultValue "drawing"
                    ParagraphLevel = boolean value "OpaqueParagraphLevel"
                    ShapeCandidate = false }
            | None, true, _ -> Picture (readImage value)
            | None, false, true -> Equation (text value "Text", text value "EquationXml")
            | _ -> Text (text value "Text")
    { Content = content
      Formatting = readRunFormat value
      Hyperlink = textOf value "Hyperlink"
      InactiveHyperlink = textOf value "InactiveHyperlink"
      CommentId = textOf value "CommentId"
      CommentText = textOf value "CommentText"
      CommentAuthor = textOf value "CommentAuthor"
      CommentDate = textOf value "CommentDate"
      NoteId = textOf value "NoteId"
      NoteType = textOf value "NoteType"
      NoteText = textOf value "NoteText"
      BookmarkStarts = array value "BookmarkStarts" |> Array.map readBookmarkStart |> List.ofArray
      BookmarkEnds = array value "BookmarkEnds" |> Array.map readBookmarkEnd |> List.ofArray
      FieldCode = textOf value "FieldCode"
      FieldResult = textOf value "FieldResult"
      Revision = readRevision value
      RenderedPageBreakBefore = boolean value "RenderedPageBreakBefore"
      Checkbox = get<bool> value "Checkbox"
      ContentControls = array value "ContentControls" |> Array.map readControlLayer |> Array.toList
      SourceRun = integer value "SourceRun"
      NoProps = boolean value "NoProps" }


let private readParagraphFormat value =
    let markRevision =
        let id = textOf value "MarkRevisionId"
        let kind = textOf value "MarkRevisionType"
        if kind.IsNone && id.IsNone then None
        else Some {
            Id = id; Kind = kind
            Author = textOf value "MarkRevisionAuthor"
            Date = textOf value "MarkRevisionDate"
            Inserted = None; MoveName = None }
    { Style = textOf value "Style"
      StyleInfo = readInferredStyle value
      Alignment = textOf value "Alignment" |> Option.defaultValue "left"
      HasExplicit = boolean value "HasExplicit"
      ExplicitAlignment = textOf value "ExplicitAlignment"
      IndentLeft = number value "IndentLeft"
      IndentRight = number value "IndentRight"
      FirstLineIndent = number value "FirstLineIndent"
      SpacingBefore = number value "SpacingBefore"
      SpacingAfter = number value "SpacingAfter"
      LineSpacing = number value "LineSpacing"
      LineSpacingRule = textOf value "LineSpacingRule"
      RightTabPosition = number value "RightTabPosition"
      RightTabLeader = textOf value "RightTabLeader"
      TabStops =
          array value "TabStops"
          |> Array.choose (fun stop ->
              number stop "Position" |> Option.map (fun position -> {
                  Position = position
                  Alignment = textOf stop "Alignment" |> Option.defaultValue "left"
                  Leader = textOf stop "Leader" }))
          |> List.ofArray
      Rtl = boolean value "Rtl"
      ContextualSpacing = get<bool> value "ContextualSpacing"
      KeepNext = get<bool> value "KeepNext"
      KeepLines = get<bool> value "KeepLines"
      WidowControl = get<bool> value "WidowControl"
      PageBreakBefore = boolean value "PageBreakBefore"
      PageBreakBeforeDerived = boolean value "PageBreakBeforeDerived"
      PageBreakAtStart = boolean value "PageBreakAtStart"
      PageBreakAfter = boolean value "PageBreakAfter"
      ColumnBreakAfter = boolean value "ColumnBreakAfter"
      RenderedPageBreakBefore = integer value "RenderedPageBreakBefore"
      BackgroundColor = textOf value "BackgroundColor"
      ShadingPattern = textOf value "ParagraphShadingPattern"
      ShadingForeground = textOf value "ParagraphShadingForeground"
      Borders = readBorders value?ParagraphBorders
      ParagraphBorderBefore = number value "ParagraphBorderBefore"
      ParagraphBorderAfter = number value "ParagraphBorderAfter"
      OutlineLevel = integer value "OutlineLevel"
      ExplicitParagraphFormat = boolean value "ExplicitParagraphFormat"
      ExplicitIndentLeft = number value "ExplicitIndentLeft"
      ExplicitIndentRight = number value "ExplicitIndentRight"
      ExplicitFirstLineIndent = number value "ExplicitFirstLineIndent"
      ExplicitSpacingBefore = number value "ExplicitSpacingBefore"
      ExplicitSpacingAfter = number value "ExplicitSpacingAfter"
      ExplicitLineSpacing = number value "ExplicitLineSpacing"
      ExplicitContextualSpacing = get<bool> value "ExplicitContextualSpacing"
      ExplicitBackgroundColor = textOf value "ExplicitBackgroundColor"
      ExplicitShadingPattern = textOf value "ExplicitParagraphShadingPattern"
      ExplicitShadingForeground = textOf value "ExplicitParagraphShadingForeground"
      ExplicitTabs = boolean value "ExplicitTabs"
      HasMarkProperties = boolean value "HasMarkProperties"
      ExplicitMarkFontSize = number value "ExplicitMarkFontSize"
      ParagraphMarkFontSize = number value "ParagraphMarkFontSize"
      MarkRevision = markRevision
      DropCapKind = textOf value "DropCapKind"
      DropCapLines = integer value "DropCapLines"
      DropCapDistance = number value "DropCapDistance"
      ContentControlId = textOf value "ContentControlId"
      ContentControlTag = textOf value "ContentControlTag"
      ContentControlAlias = textOf value "ContentControlAlias"
      ContentControlXml = textOf value "ContentControlXml"
      ContentControlGroup = textOf value "ContentControlGroup"
      ContentControlAncestors = array value "ContentControlAncestors" |> Array.map readControlLayer |> Array.toList
      EnclosingCellControls = array value "EnclosingCellControls" |> Array.map readControlLayer |> Array.toList }

let private readParagraphList value =
    let style = textOf value "ListStyle" |> Option.defaultValue "none"
    if style = "none" && (textOf value "ListId").IsNone then None
    else
        let levelIndents =
            if boolean value "HasListLevelIndents" then
                Some (number value "ListLevelIndentLeft", number value "ListLevelFirstLineIndent")
            else None
        Some {
            Style = style
            Level = intValue value "ListLevel" 0
            Id = textOf value "ListId"
            SourceId = textOf value "ListSourceId"
            Restart = boolean value "ListRestart"
            FormatEdited = boolean value "ListFormatEdited"
            NumberFormat = textOf value "ListNumberFormat"
            LevelText = textOf value "ListLevelText"
            Start = integer value "ListStart"
            StartOverrides =
                array value "ListStartOverrides"
                |> Array.choose (fun overrideValue ->
                    match integer overrideValue "Level", integer overrideValue "Start" with
                    | Some level, Some start -> Some (level, start)
                    | _ -> None)
                |> List.ofArray
            ListLevelIndents = levelIndents
            Levels =
                array value "ListLevels"
                |> Array.map (fun level -> {
                    Level = intValue level "Level" 1
                    NumberFormat = textOf level "NumberFormat"
                    LevelText = textOf level "LevelText"
                    Start = integer level "Start"
                    IndentLeft = number level "IndentLeft"
                    FirstLineIndent = number level "FirstLineIndent"
                    Suffix = textOf level "Suffix"
                    Justification = textOf level "Justification"
                    Legal = boolean level "Legal"
                    Restart = integer level "Restart"
                    MarkFont = textOf level "MarkFont"
                    MarkBold = get<bool> level "MarkBold"
                    MarkItalic = get<bool> level "MarkItalic"
                    MarkColor = textOf level "MarkColor"
                    MarkSize = number level "MarkSize" })
                |> List.ofArray }

/// Progressive front-first browser import exposes only the newly read slice
/// as JS objects, leaving its typed conversion policy shared with native.
let blocksToEditorObjects (blocks: EditorBlock array) : obj array =
    blocks |> Array.map blockToEditorObject

let pageSetupToEditorObject (setup: PageSetup option) : obj =
    pageSetupToObject setup

let rec private readParagraph value : EditorParagraph =
    let kind = textOf value "Type" |> Option.defaultValue "Paragraph"
    let spliceFrom =
        get<obj> value "SpliceFrom"
        |> Option.bind (fun span ->
            match integer span "Start", integer span "End" with
            | Some start, Some finish -> Some (start, finish)
            | _ -> None)
    { Runs = array value "Runs" |> Array.map readRun
      HeadingLevel = if kind = "Heading" then Some (intValue value "HeadingLevel" 1) else None
      Formatting = readParagraphFormat value
      List = readParagraphList value
      SectionBreak = textOf value "SectionBreak"
      SectionPageSetup = readPageSetup value?SectionPageSetup
      SectionColumns = readColumns value?SectionColumns
      SectionVerticalAlign = textOf value "SectionVerticalAlign"
      SectionPageNumberStart = integer value "SectionPageNumberStart"
      SectionPageNumberFormat = textOf value "SectionPageNumberFormat"
      HeaderReferences = readRefs value "HeaderReferences"
      FooterReferences = readRefs value "FooterReferences"
      Source = readSource value
      SourceParagraph = get<obj> value "SourcePara" |> Option.map readParagraphEdit
      SpliceOriginal = get<obj> value "SpliceOriginal" |> Option.map readParagraph
      SpliceFrom = spliceFrom
      SpliceLeading = get<bool> value "SpliceLeading" |> Option.defaultValue true
      SpliceTrailing = get<bool> value "SpliceTrailing" |> Option.defaultValue true
      SpliceJoined = array value "SpliceJoined" |> Array.map readParagraph
      SourceIndex = integer value "SourceIndex"
      SourceUnchanged = boolean value "SourceUnchanged"
      SourceWhy = textOf value "SourceWhy"
      SourceTextChanges = get<obj> value "SourceTextChanges" |> Option.map readChanges }

and private readParagraphEdit value : ParagraphEdit =
    let change =
        get<obj> value "Text" |> Option.map (fun edit -> {
            Paragraph = intValue edit "Paragraph" (intValue value "Paragraph" -1)
            Offset = intValue edit "Offset" 0
            Deleted = intValue edit "Deleted" 0
            Inserted = text edit "Inserted"
            JoinBefore = boolean edit "JoinBefore"
            JoinAfter = boolean edit "JoinAfter"
            RunsPlain = boolean edit "RunsPlain"
            Witnesses = array edit "Witnesses" |> Array.map unbox<int>
            MarkStyle = boolean edit "MarkStyle"
            NewText = text edit "NewText" })
    let cut =
        match array value "Cut" with
        | [| first; finish |] -> Some (unbox<int> first, unbox<int> finish)
        | _ -> None
    { Paragraph = intValue value "Paragraph" -1
      Block = integer value "Block"
      Para = integer value "Para"
      Text = change
      OldText = textOf value "OldText"
      Before = get<obj> value "Before" |> Option.map readParagraph
      After = get<obj> value "After" |> Option.map readParagraph
      NewText = textOf value "NewText"
      Regenerate = boolean value "Regenerate"
      Verbatim = boolean value "Verbatim"
      Why = textOf value "Why"
      Cut = cut
      Join =
        array value "Join"
        |> Array.choose (fun joined ->
            get<obj> joined "Original" |> Option.map (fun original -> {
                Block = intValue joined "Block" -1
                Para = intValue joined "Para" -1
                Text = text joined "Text"
                Original = readParagraph original })) }

and private readChanges value : SourceTextChanges =
    { Texts = array value "Texts" |> Array.map unbox<string>
      Edits = array value "Edits" |> Array.map readParagraphEdit }

and private readBlock value : EditorBlock =
    match textOf value "Type" |> Option.defaultValue "Paragraph" with
    | "Table" -> Table (readTable value)
    | "TextBox" -> TextBox (readTextBox value)
    | _ -> Paragraph (readParagraph value)

and private readTableSource value : TableSource option =
    if isMissing value then None
    else Some {
        At = intValue value "At" -1
        Block = integer value "Block"
        Before = get<obj> value "Before" |> Option.map readTable
        Was = get<obj> value "Was" |> Option.map readTable
        Copy = boolean value "Copy"
        Same = boolean value "Same" }

and private readRowSource value : RowSource option =
    if isMissing value then None
    else Some {
        At = intValue value "At" -1
        Before = get<obj> value "Before" |> Option.map readRow
        Was = get<obj> value "Was" |> Option.map readRow
        Copy = boolean value "Copy"
        Same = boolean value "Same" }

and private readCellSource value : CellSource option =
    if isMissing value then None
    else Some {
        At = intValue value "At" -1
        Before = get<obj> value "Before" |> Option.map readCell
        Was = get<obj> value "Was" |> Option.map readCell
        Copy = boolean value "Copy"
        Same = boolean value "Same" }

and private readTable value : EditorTable =
    let look = get<obj> value "TableLook" |> Option.map (fun look -> {
        FirstRow = boolean look "FirstRow"
        LastRow = boolean look "LastRow"
        FirstColumn = boolean look "FirstColumn"
        LastColumn = boolean look "LastColumn"
        HorizontalBanding = boolean look "HorizontalBanding"
        VerticalBanding = boolean look "VerticalBanding" })
    { Rows = array value "Rows" |> Array.map readRow
      PreferredWidth = number value "PreferredWidth"
      PreferredWidthType = textOf value "PreferredWidthType"
      Alignment = textOf value "TableAlignment"
      Indent = number value "TableIndent"
      LayoutMode = textOf value "LayoutMode"
      GridColumns = array value "GridColumns" |> Array.map (fun grid -> jsNumber grid)
      Borders = readBorders value?Borders
      BackgroundColor = textOf value "BackgroundColor"
      Shading = readShading value?Shading
      CellSpacing = number value "CellSpacing"
      PaddingTop = number value "PaddingTop"
      PaddingRight = number value "PaddingRight"
      PaddingBottom = number value "PaddingBottom"
      PaddingLeft = number value "PaddingLeft"
      Floating = boolean value "Floating"
      FloatingXml = textOf value "FloatingXml"
      StyleId = textOf value "TableStyleId"
      Look = look
      ContentControlId = textOf value "ContentControlId"
      ContentControlTag = textOf value "ContentControlTag"
      ContentControlAlias = textOf value "ContentControlAlias"
      ContentControlXml = textOf value "ContentControlXml"
      ContentControlGroup = textOf value "ContentControlGroup"
      ContentControlAncestors = array value "ContentControlAncestors" |> Array.map readControlLayer |> Array.toList
      Source = readSource value
      SourceIndex = integer value "SourceIndex"
      SourceUnchanged = boolean value "SourceUnchanged"
      SourceWhy = textOf value "SourceWhy"
      SourceTextChanges = get<obj> value "SourceTextChanges" |> Option.map readChanges
      SourceTable = readTableSource value?SourceTable }

and private readRow value : EditorRow =
    { Cells = array value "Cells" |> Array.map readCell
      IsHeader = boolean value "IsHeader"
      Height = number value "Height"
      HeightRule = textOf value "HeightRule"
      CanSplit = get<bool> value "CanSplit"
      CellSpacing = number value "CellSpacing"
      PaddingTop = number value "PaddingTop"
      PaddingRight = number value "PaddingRight"
      PaddingBottom = number value "PaddingBottom"
      PaddingLeft = number value "PaddingLeft"
      GridBefore = intValue value "GridBefore" 0
      GridAfter = intValue value "GridAfter" 0
      WidthBefore = number value "WidthBefore"
      WidthAfter = number value "WidthAfter"
      BordersOverride = readBorders value?BordersOverride
      Source = readSource value
      SourceRow = readRowSource value?SourceRow }

and private readCell value : EditorCell =
    { Paragraphs = array value "Paragraphs" |> Array.map readBlock
      NestedTables = array value "NestedTables" |> Array.map readTable
      NestedTableAnchors = array value "NestedTableAnchors" |> Array.map unbox<int>
      EnclosingControls = array value "EnclosingControls" |> Array.map readControlLayer |> Array.toList
      ColSpan = intValue value "ColSpan" 1
      BackgroundColor = textOf value "BackgroundColor"
      Shading = readShading value?Shading
      Width = number value "Width"
      PreferredWidth = number value "PreferredWidth"
      PreferredWidthType = textOf value "PreferredWidthType"
      BorderColor = textOf value "BorderColor"
      BorderWidth = number value "BorderWidth"
      Borders = readBorders value?Borders
      VerticalAlignment = textOf value "VerticalAlignment"
      VerticalMerge = textOf value "VerticalMerge"
      PaddingTop = number value "PaddingTop"
      PaddingRight = number value "PaddingRight"
      PaddingBottom = number value "PaddingBottom"
      PaddingLeft = number value "PaddingLeft"
      Source = readSource value
      SourceCell = readCellSource value?SourceCell }

and private readTextBox value : EditorTextBox =
    let gradient = get<obj> value "Gradient"
    let stops = match gradient with Some gradient -> array gradient "Stops" | None -> array value "GradientStops"
    let childSpace =
        match number value "ChildSpaceWidth", number value "ChildSpaceHeight" with
        | Some width, Some height -> Some (width, height)
        | _ -> get<obj> value "ChildSpace" |> Option.bind (fun space ->
            match number space "W", number space "H" with
            | Some width, Some height -> Some (width, height)
            | _ -> None)
    { Paragraphs = array value "Paragraphs" |> Array.map readBlock
      Children = array value "Children" |> Array.map readTextBox
      ShapeId = textOf value "ShapeId"
      ShapeKind = textOf value "ShapeKind"
      Name = textOf value "Name"
      Geometry = textOf value "Geometry" |> Option.defaultValue "rect"
      CustomPath = array value "CustomPath" |> Array.map (fun segment ->
          text segment "Op", array segment "P" |> Array.map jsNumber) |> List.ofArray
      CustomWidth = amount value "CustomW" 0.0
      CustomHeight = amount value "CustomH" 0.0
      X = number value "X"; Y = number value "Y"
      PositionHRelativeFrom = textOf value "PositionHRelativeFrom"
      PositionVRelativeFrom = textOf value "PositionVRelativeFrom"
      PositionHAlign = textOf value "PositionHAlign"
      PositionVAlign = textOf value "PositionVAlign"
      Width = amount value "Width" 200.0
      Height = amount value "Height" 100.0
      FillColor = textOf value "FillColor"
      BorderColor = textOf value "BorderColor"
      NoFill = boolean value "NoFill"
      NoLine = boolean value "NoLine"
      BorderWidth = number value "BorderWidth"
      BorderDash = textOf value "BorderDash"
      Gradient = stops |> Array.map (fun stop ->
          amount stop "Offset" 0.0, text stop "Color") |> List.ofArray
      GradientAngle = gradient |> Option.bind (fun gradient -> number gradient "Angle") |> Option.defaultValue (amount value "GradientAngle" 0.0)
      GradientRadial = gradient |> Option.map (fun gradient -> boolean gradient "Radial") |> Option.defaultValue (boolean value "GradientRadial")
      Rotation = amount value "Rotation" 0.0
      FlipH = boolean value "FlipH"; FlipV = boolean value "FlipV"
      Inline = boolean value "Inline"
      Wrap = textOf value "Wrap"
      BehindDoc = get<bool> value "BehindDoc"
      AllowOverlap = get<bool> value "AllowOverlap"
      LayoutInCell = get<bool> value "LayoutInCell"
      DistanceTop = number value "DistanceTop"
      DistanceBottom = number value "DistanceBottom"
      DistanceLeft = number value "DistanceLeft"
      DistanceRight = number value "DistanceRight"
      RelativeHeight = textOf value "RelativeHeight"
      AnchorBlock = integer value "AnchorBlock"
      InsetLeft = number value "InsetLeft"
      InsetTop = number value "InsetTop"
      InsetRight = number value "InsetRight"
      InsetBottom = number value "InsetBottom"
      VerticalAnchor = textOf value "VerticalAnchor"
      AutoFit = boolean value "AutoFit"
      NoTextWrap = boolean value "NoTextWrap"
      FitText = boolean value "FitText"
      Opacity = amount value "Opacity" 1.0
      ChildSpace = childSpace
      ImageSource = textOf value "ImageSrc"
      SourceXml = textOf value "SourceXml"
      SourceMoved = get<obj> value "SourceMoved" |> Option.map (fun moved -> {
          H = boolean moved "H"; V = boolean moved "V"; Size = boolean moved "Size" })
      SourceIndex = integer value "SourceIndex"
      SourceUnchanged = boolean value "SourceUnchanged"
      SourceWhy = textOf value "SourceWhy"
      SourceTextChanges = get<obj> value "SourceTextChanges" |> Option.map readChanges }

let private readStyle value : Style =
    let run =
        if not (hasAnyStyleField styleRunFieldNames value) then None
        else Some {
            Bold = get<bool> value "StyleBold"
            Italic = get<bool> value "StyleItalic"
            Underline = get<bool> value "StyleUnderline"
            UnderlineStyle = textOf value "StyleUnderlineStyle"
            UnderlineColor = textOf value "StyleUnderlineColor"
            AllCaps = get<bool> value "StyleAllCaps"
            SmallCaps = get<bool> value "StyleSmallCaps"
            Strike = get<bool> value "StyleStrike"
            DoubleStrike = get<bool> value "StyleDoubleStrike"
            Hidden = get<bool> value "StyleHidden"
            Position = number value "StylePosition"
            EmphasisMark = textOf value "StyleEmphasisMark"
            VertAlign = textOf value "StyleVertAlign"
            FontSize = number value "StyleFontSize"
            CharacterSpacing = number value "StyleCharacterSpacing"
            FontFamily = textOf value "StyleFontFamily"
            Color = textOf value "StyleColor"
            Highlight = textOf value "StyleHighlight" }
    let paragraph =
        if not (hasAnyStyleField styleParagraphFieldNames value) then None
        else Some {
            SpacingBefore = number value "StyleSpacingBefore"
            SpacingAfter = number value "StyleSpacingAfter"
            ContextualSpacing = get<bool> value "StyleContextualSpacing"
            LineSpacing = number value "StyleLineSpacing"
            LineSpacingRule = textOf value "StyleLineRule" |> Option.orElse (textOf value "StyleLineSpacingRule")
            Alignment = textOf value "StyleAlignment"
            IndentLeft = number value "StyleIndentLeft"
            IndentRight = number value "StyleIndentRight"
            FirstLineIndent = number value "StyleFirstLineIndent"
            BackgroundColor = textOf value "StyleBackgroundColor"
            OutlineLevel = integer value "StyleOutlineLevel"
            KeepNext = get<bool> value "StyleKeepNext"
            KeepLines = get<bool> value "StyleKeepLines"
            PageBreakBefore = get<bool> value "StylePageBreakBefore"
            WidowControl = get<bool> value "StyleWidowControl" }
    { Id = textOf value "StyleId" |> Option.orElse (textOf value "Id") |> Option.defaultValue ""
      Name = textOf value "Name"
      StyleType = textOf value "StyleType" |> Option.defaultValue "paragraph"
      IsDefault = get<bool> value "IsDefaultStyle" |> Option.orElse (get<bool> value "IsDefault") |> Option.defaultValue false
      BasedOn = textOf value "BasedOn"
      Next = textOf value "NextStyle" |> Option.orElse (textOf value "Next")
      Linked = textOf value "LinkedStyle" |> Option.orElse (textOf value "Linked")
      IsQuickStyle = get<bool> value "IsQuickStyle"
      Aliases = textOf value "StyleAliases" |> Option.orElse (textOf value "Aliases")
      UiPriority = integer value "StyleUiPriority"
      SemiHidden = get<bool> value "StyleSemiHidden"
      UnhideWhenUsed = get<bool> value "StyleUnhideWhenUsed"
      Run = run
      Paragraph = paragraph }

[<Emit("Object.keys($0)")>]
let private keys (_value: obj) : string array = jsNative

let private readTheme value : DocxReader.Theme option =
    if isMissing value then None
    else
        let colors = get<obj> value "Colors"
        Some {
            MajorFont = textOf value "MajorFont"
            MinorFont = textOf value "MinorFont"
            Colors =
                match colors with
                | Some colorMap -> keys colorMap |> Array.map (fun key -> key, text colorMap key) |> Map.ofArray
                | None -> Map.empty }

/// Convert the public browser's package object once at its API boundary.
/// Sparse editor-authored packages use the same defaults as the legacy API.
let packageOfEditorObject (value: obj) : EditorPackage =
    let headerParts name =
        array value name
        |> Array.map (fun part -> {
            RelationshipId = textOf part "RelationshipId"
            Kind = textOf part "Type" |> Option.orElse (textOf part "Kind") |> Option.defaultValue "default"
            Blocks = array part "Paragraphs" |> Array.map readBlock
            SourceStoryCounts = array part "SourceStoryCounts" |> Array.map unbox<int>
            SourceUnchanged = boolean part "SourceUnchanged" })
    let notes name =
        array value name
        |> Array.map (fun note -> {
            Id = text note "Id"
            Kind = textOf note "Kind"
            HasMark = get<bool> note "HasMark"
            MarkStyle = textOf note "MarkStyle"
            Text = textOf note "Text"
            Blocks = array note "Blocks" |> Array.map readBlock
            SourceStoryCounts = array note "SourceStoryCounts" |> Array.map unbox<int>
            Unchanged = boolean note "SourceUnchanged" })
    let comments =
        array value "Comments"
        |> Array.map (fun comment -> {
            Id = text comment "Id"
            Author = textOf comment "Author"
            Date = textOf comment "Date"
            Initials = textOf comment "Initials"
            Text = textOf comment "Text"
            Blocks = array comment "Blocks" |> Array.map readBlock
            SourceStoryCounts = array comment "SourceStoryCounts" |> Array.map unbox<int>
            Unchanged = boolean comment "SourceUnchanged" })
    { Blocks = array value "Paragraphs" |> Array.map readBlock
      Styles = array value "Styles" |> Array.map readStyle
      HeaderParts = headerParts "HeaderParts"
      FooterParts = headerParts "FooterParts"
      Footnotes = notes "Footnotes"
      Endnotes = notes "Endnotes"
      Comments = comments
      Theme = get<obj> value "Theme" |> Option.bind readTheme
      PageSetup = get<obj> value "PageSetup" |> Option.bind readPageSetup
      OriginalPageSetup = get<obj> value "OriginalPageSetup" |> Option.bind readPageSetup
      Preserved = get<obj> value "Preserved" |> Option.bind preservedOfEditorObject
      SourceStoryCounts = array value "SourceStoryCounts" |> Array.map unbox<int>
      ChangedStyles = get<obj array> value "ChangedStyles" |> Option.map (Array.map unbox<string>)
      ChangedListLevels = array value "ChangedListLevels" |> Array.map unbox<string>
      PageSetupUnchanged = boolean value "PageSetupUnchanged"
      ListsUnchanged = boolean value "ListsUnchanged"
      FootnotesUnchanged = boolean value "FootnotesUnchanged"
      EndnotesUnchanged = boolean value "EndnotesUnchanged"
      CommentsUnchanged = boolean value "CommentsUnchanged"
      Unchanged = boolean value "Unchanged"
      RegenerateAll = boolean value "RegenerateAll" }
#endif
