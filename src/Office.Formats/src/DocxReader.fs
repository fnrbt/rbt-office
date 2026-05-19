module DocxReader

open XmlDsl
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
    AllCaps: bool option
    SmallCaps: bool option
    FontSize: string option
    FontFamily: string option
    ThemeFont: string option
    Color: string option
    ThemeColor: string option
    ThemeTint: string option
    ThemeShade: string option
    Highlight: string option
    Strike: bool option
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
    | PageBreak

type Run = {
    Content: RunContent
    Properties: RunProperties
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
    Properties: TableCellProperties option
}

type HeaderFooter = {
    Content: Paragraph list
    Type: string  // "header", "footer", "first", "even"
    RelationshipId: string option
}

type DocumentContent =
    | Para of Paragraph
    | Tbl of Table

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

    let private escapeXml (text: string) =
        text
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;")
            .Replace("'", "&apos;")

    let rec xmlToString node =
        match node with
        | XmlDsl.Text text -> escapeXml text
        | XmlDsl.Empty -> ""
        | XmlDsl.Element(name, attrs, children) ->
            let attrText =
                attrs
                |> Map.toList
                |> List.map (fun (key, value) -> sprintf "%s=\"%s\"" key (escapeXml value))
                |> String.concat " "
            let attrPart = if attrText = "" then "" else " " + attrText
            match children with
            | [] -> sprintf "<%s%s/>" name attrPart
            | _ -> sprintf "<%s%s>%s</%s>" name attrPart (children |> List.map xmlToString |> String.concat "") name

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

    // Run properties extraction with defaults
    let runProps node =
        let rPr = wElem "rPr" node
        match rPr with
        | Some props ->
            {
                Style = getPropVal "rStyle" "val" props
                Bold = onOff "b" props
                Italic = onOff "i" props
                Underline = onOff "u" props
                AllCaps = onOff "caps" props
                SmallCaps = onOff "smallCaps" props
                FontSize = getPropVal "sz" "val" props
                FontFamily = getPropVal "rFonts" "ascii" props
                ThemeFont = getPropVal "rFonts" "asciiTheme" props
                Color = getPropVal "color" "val" props
                ThemeColor = getPropVal "color" "themeColor" props
                ThemeTint = getPropVal "color" "themeTint" props
                ThemeShade = getPropVal "color" "themeShade" props
                Highlight = getPropVal "highlight" "val" props
                Strike = onOff "strike" props
                VertAlign = getPropVal "vertAlign" "val" props
            }
        | None ->
            {
                Style = None
                Bold = None
                Italic = None
                Underline = None
                AllCaps = None
                SmallCaps = None
                FontSize = None
                FontFamily = None
                ThemeFont = None
                Color = None
                ThemeColor = None
                ThemeTint = None
                ThemeShade = None
                Highlight = None
                Strike = None
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

    let paragraphProps node =
        let pPr = wElem "pPr" node
        match pPr with
        | Some props ->
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
                Alignment = getPropVal "jc" "val" props
                IndentLeft = getPropVal "ind" "left" props
                IndentRight = getPropVal "ind" "right" props
                FirstLineIndent = getPropVal "ind" "firstLine" props
                HangingIndent = getPropVal "ind" "hanging" props
                RightTabPosition =
                    wElem "tabs" props
                    |> Option.bind (fun tabs ->
                        Xml.children tabs
                        |> List.tryPick (fun tab ->
                            if Xml.name tab |> Option.exists (fun name -> Ns.stripNs name = "tab") && attr "val" tab = Some "right" then attr "pos" tab else None))
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
              RevisionId = None
              RevisionType = None
              RevisionAuthor = None
              RevisionDate = None })

    let fieldData (rNode: XmlNode) : Field option =
        wElem "fldChar" rNode
        |> Option.bind (fun fld ->
            let fldType = Xml.attr "fldCharType" fld
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

    // Run extraction preserves child order. In WordprocessingML, w:br is a run
    // child at the exact break location, not a paragraph-level flag.
    let runsFromRun (rNode: XmlNode) : Run list =
        let props = runProps rNode
        let revId = wElem "rPr" rNode
                    |> Option.bind (wElem "rPrChange")
                    |> Option.bind (Xml.attr "id")
        let makeRun content =
            { Content = content; Properties = props; RevisionId = revId; RevisionType = None; RevisionAuthor = None; RevisionDate = None }

        let childRuns =
            Xml.children rNode
            |> List.choose (fun child ->
                Xml.name child
                |> Option.bind (fun name ->
                    match Ns.stripNs name with
                    | "t"
                    | "delText" -> Xml.text child |> Option.map (Text >> makeRun)
                    | "tab" -> Some (makeRun (Text "\t"))
                    | "br" ->
                        match attr "type" child with
                        | Some "page" -> Some (makeRun PageBreak)
                        | _ -> Some (makeRun (Text "\n"))
                    | "drawing"
                    | "pict" -> imageData rNode |> Option.map (Image >> makeRun)
                    | "commentReference" -> attr "id" child |> Option.map (CommentReference >> makeRun)
                    | "footnoteReference" -> attr "id" child |> Option.map (fun id -> makeRun (NoteRef { Id = id; NoteType = "footnote" }))
                    | "endnoteReference" -> attr "id" child |> Option.map (fun id -> makeRun (NoteRef { Id = id; NoteType = "endnote" }))
                    | "fldChar" -> fieldData rNode |> Option.map (FieldRef >> makeRun)
                    | _ -> None))

        if not childRuns.IsEmpty then childRuns
        else
            match imageData rNode with
            | Some img -> [makeRun (Image img)]
            | None ->
                match commentReference rNode with
                | Some commentId -> [makeRun (CommentReference commentId)]
                | None ->
                    match noteReference rNode with
                    | Some noteRef -> [makeRun (NoteRef noteRef)]
                    | None ->
                        match fieldData rNode with
                        | Some field -> [makeRun (FieldRef field)]
                        | None -> textContent rNode |> Option.map (Text >> makeRun >> List.singleton) |> Option.defaultValue []

    // Hyperlink extraction
    let hyperlink (hlNode: XmlNode) : Run list =
        let rId = Xml.attr "r:id" hlNode
        let runs = wElems "r" hlNode |> List.collect runsFromRun

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
            |> Option.map (fun r -> r.Properties)
            |> Option.defaultValue {
                Style = None
                Bold = None; Italic = None; Underline = None
                AllCaps = None; SmallCaps = None
                FontSize = None; FontFamily = None; ThemeFont = None; Color = None
                ThemeColor = None; ThemeTint = None; ThemeShade = None; Highlight = None; Strike = None; VertAlign = None
            }

        [{
            Content = Link {
                Text = text
                Target = None  // Resolved from relationships later
                RelationshipId = rId
                Properties = props
            }
            Properties = props
            RevisionId = None
            RevisionType = None
            RevisionAuthor = None
            RevisionDate = None
        }]

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

        let rec runsFromNode node =
            Xml.children node
            |> List.collect (fun child ->
                match Xml.name child with
                | Some name when Ns.stripNs name = "r" ->
                    runsFromRun child
                | Some name when Ns.stripNs name = "hyperlink" ->
                    hyperlink child
                | Some name when Ns.stripNs name = "fldSimple" ->
                    simpleField child |> Option.toList
                | Some name when Ns.stripNs name = "oMath" || Ns.stripNs name = "oMathPara" ->
                    [{ Content = EquationXml (xmlToString child); Properties = runProps child; RevisionId = None; RevisionType = None; RevisionAuthor = None; RevisionDate = None }]
                | Some name when Ns.stripNs name = "bookmarkStart" ->
                    Xml.attr "id" child
                    |> Option.map (fun id -> { Content = BookmarkStart { Id = id; Name = Xml.attr "name" child }; Properties = runProps child; RevisionId = None; RevisionType = None; RevisionAuthor = None; RevisionDate = None })
                    |> Option.toList
                | Some name when Ns.stripNs name = "bookmarkEnd" ->
                    Xml.attr "id" child
                    |> Option.map (fun id -> { Content = BookmarkEnd id; Properties = runProps child; RevisionId = None; RevisionType = None; RevisionAuthor = None; RevisionDate = None })
                    |> Option.toList
                | Some name when Ns.stripNs name = "ins" || Ns.stripNs name = "del" ->
                    let revisionType = if Ns.stripNs name = "del" then "delete" else "insert"
                    let revisionId = Xml.attr "id" child |> Option.orElse (Xml.attr "w:id" child)
                    let author = Xml.attr "author" child |> Option.orElse (Xml.attr "w:author" child)
                    let date = Xml.attr "date" child |> Option.orElse (Xml.attr "w:date" child)
                    runsFromNode child
                    |> List.map (fun run -> { run with RevisionId = revisionId; RevisionType = Some revisionType; RevisionAuthor = author; RevisionDate = date })
                | Some name when Ns.stripNs name = "sdt" ->
                    Ns.childNs "sdtContent" child
                    |> Option.map runsFromNode
                    |> Option.defaultValue []
                | _ -> [])

        // Extract runs and hyperlinks
        let runs = runsFromNode pNode

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

    // Table extraction with nested document builder
    let rec tableCell (tcNode: XmlNode) : TableCell =
        let paragraphs =
            wChildElems "p" tcNode
            |> List.map paragraph
        {
            Paragraphs = paragraphs
            Properties = tableCellProps tcNode
        }

    let tableRow (trNode: XmlNode) : TableRow =
        let cells =
            wChildElems "tc" trNode
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

    let table (tblNode: XmlNode) : Table =
        let rows =
            wChildElems "tr" tblNode
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
                let flag name fallback =
                    attr name look
                    |> Option.map (fun value -> value = "1" || value = "true")
                    |> Option.defaultValue fallback
                { FirstRow = flag "firstRow" false
                  LastRow = flag "lastRow" false
                  FirstColumn = flag "firstColumn" false
                  LastColumn = flag "lastColumn" false
                  HorizontalBanding = not (flag "noHBand" false)
                  VerticalBanding = not (flag "noVBand" false) })
        let gridColumns =
            wElem "tblGrid" tblNode
            |> Option.map (fun grid -> wChildElems "gridCol" grid |> List.choose (attr "w"))
            |> Option.defaultValue []

        {
            Rows = rows
            StyleId = tblPr |> Option.bind (wElem "tblStyle") |> Option.bind (attr "val")
            Look = tableLook
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
            let id = props |> Option.bind (Ns.childNs "id") |> Option.bind (Xml.attr "val")
            let tag = props |> Option.bind (Ns.childNs "tag") |> Option.bind (Xml.attr "val")
            let alias = props |> Option.bind (Ns.childNs "alias") |> Option.bind (Xml.attr "val")
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
                    | _ -> ()
            }

        let rec normalizeSectionBreaks = function
            | Para para :: (Para nextPara :: _ as rest)
                when para.Runs.IsEmpty && para.Properties.SectionBreak = Some "nextPage" && nextPara.Properties.SectionBreak = Some "continuous" ->
                Para { para with Properties = { para.Properties with SectionBreak = None } } :: normalizeSectionBreaks rest
            | item :: rest -> item :: normalizeSectionBreaks rest
            | [] -> []

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
            Underline = Docx.onOff "u" rPr
            AllCaps = Docx.onOff "caps" rPr
            SmallCaps = Docx.onOff "smallCaps" rPr
            FontSize = Ns.childNs "sz" rPr |> Option.bind (Docx.attr "val")
            FontFamily = Ns.childNs "rFonts" rPr |> Option.bind (Docx.attr "ascii")
            ThemeFont = Ns.childNs "rFonts" rPr |> Option.bind (Docx.attr "asciiTheme")
            Color = Ns.childNs "color" rPr |> Option.bind (Docx.attr "val")
            ThemeColor = Ns.childNs "color" rPr |> Option.bind (Docx.attr "themeColor")
            ThemeTint = Ns.childNs "color" rPr |> Option.bind (Docx.attr "themeTint")
            ThemeShade = Ns.childNs "color" rPr |> Option.bind (Docx.attr "themeShade")
            Highlight = Ns.childNs "highlight" rPr |> Option.bind (Docx.attr "val")
            Strike = Docx.onOff "strike" rPr
            VertAlign = Ns.childNs "vertAlign" rPr |> Option.bind (Docx.attr "val")
        }

    let parseParagraphPropsFromStyle (pPr: XmlNode) : ParagraphProperties =
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
            RightTabPosition =
                Ns.childNs "tabs" pPr
                |> Option.bind (fun tabs ->
                    Xml.children tabs
                    |> List.tryPick (fun tab ->
                        if Xml.name tab |> Option.exists (fun name -> Ns.stripNs name = "tab") && Docx.attr "val" tab = Some "right" then Docx.attr "pos" tab else None))
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
            let conditionalChildProperties childName propertyParser condition =
                Xml.children styleNode
                |> List.tryPick (fun child ->
                    if Xml.name child |> Option.exists (fun name -> Ns.stripNs name = "tblStylePr")
                       && Docx.attr "type" child = Some condition then
                        Ns.childNs childName child |> Option.bind propertyParser
                    else None)
            let conditionalTableProperties condition = conditionalChildProperties "tblPr" tablePropertiesFromTblPr condition
            let conditionalRowProperties condition = conditionalChildProperties "trPr" tableRowPropertiesFromTrPr condition
            let conditionalCellProperties condition = conditionalChildProperties "tcPr" (Docx.tableCellPropertiesFromTcPr >> Some) condition
            let conditionalParagraphProperties condition = conditionalChildProperties "pPr" (parseParagraphPropsFromStyle >> Some) condition

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

// Header/Footer parsing
module HeaderFooter =
    let parseHeaderFooter (xml: string) (hfType: string) (relationshipId: string option) : HeaderFooter =
        let doc = XmlDsl.parse xml
        let paragraphs =
            Ns.descendantsNs "p" doc
            |> List.map Docx.paragraph
        {
            Content = paragraphs
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
            Xml.attr "id" revNode
            |> Option.map (fun id ->
                let author = Xml.attr "author" revNode
                let date = Xml.attr "date" revNode
                let revType = parseRevisionType (Ns.stripNs name)

                let text =
                    Ns.descendantsNs "t" revNode
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
        let fillColor = Xml.attr "fillcolor" shapeNode |> normalizeColor
        let borderColor = Xml.attr "strokecolor" shapeNode |> normalizeColor
        let borderWidth = Xml.attr "strokeweight" shapeNode |> strokeWeightFromAttr

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
        }

    let parseShapes (doc: XmlNode) : Shape list =
        (Ns.descendantsNs "wsp" doc) @ (Ns.descendantsNs "shape" doc)
        |> List.filter (fun node -> Ns.descendantsNs "txbxContent" node |> List.isEmpty |> not)
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

    let resolveContent = function
        | Para p -> Para (resolvePara p)
        | Tbl t ->
            let resolveCell (cell: TableCell) =
                { cell with Paragraphs = cell.Paragraphs |> List.map resolvePara }
            let resolveRow (row: TableRow) =
                { row with Cells = row.Cells |> List.map resolveCell }
            Tbl { t with Rows = t.Rows |> List.map resolveRow }

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
    |> Option.map (fun settings -> Ns.descendantsNs "evenAndOddHeaders" settings |> List.isEmpty |> not)
    |> Option.defaultValue false

let private parsePageSetup evenAndOddHeaders (sectPr: XmlNode option) : PageSetup option =
    sectPr
    |> Option.map (fun section ->
        let pgSz = Ns.childNs "pgSz" section
        let pgMar = Ns.childNs "pgMar" section
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
                let imagePath = normalizePath ("word/" + rel.Target)
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
    let rec resolve dir remaining =
        let normalized = normalizePath remaining
        if normalized.StartsWith("../") then
            resolve (parentPath dir) (normalized.Substring(3))
        elif normalized.StartsWith("/") then
            normalizePath normalized
        elif dir = "" then
            normalizePath normalized
        else
            normalizePath (dir + "/" + normalized)
    resolve baseDir target

let private relationshipTargetBase (rel: Relationship) =
    parentPath (normalizePath ("word/" + rel.Target))

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
                    let path = normalizePath ("word/" + rel.Target)
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
                                return! loadImageData archive (relationshipTargetBase partRel) partRelationships
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
            let path = normalizePath ("word/" + rel.Target)
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
                loadImageDataNative archive (relationshipTargetBase partRel) partRelationships)

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
    createObj [
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
    |> Option.map (fun value -> createObj [
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

let private paragraphBorderToEditor (border: ParagraphBorder option) =
    border
    |> Option.map (fun value -> createObj [
        "Style" ==> value.Style
        "Color" ==> paragraphBorderColorToCss border
        "Width" ==> paragraphBorderWidthToPoints border
        "Space" ==> paragraphBorderSpaceToPoints border
    ])

let private paragraphBordersToEditor (borders: ParagraphBorders option) =
    borders
    |> Option.map (fun value -> createObj [
        "Top" ==> paragraphBorderToEditor value.Top
        "Right" ==> paragraphBorderToEditor value.Right
        "Bottom" ==> paragraphBorderToEditor value.Bottom
        "Left" ==> paragraphBorderToEditor value.Left
    ])

let private bordersToEditor (borders: CellBorders option) =
    borders
    |> Option.map (fun value -> createObj [
        "Top" ==> borderToEditor value.Top
        "Right" ==> borderToEditor value.Right
        "Bottom" ==> borderToEditor value.Bottom
        "Left" ==> borderToEditor value.Left
        "InsideH" ==> borderToEditor value.InsideH
        "InsideV" ==> borderToEditor value.InsideV
    ])

let private paragraphLayoutProps fontSize (props: ParagraphProperties) =
    let spacingAfter =
        if props.SpacingAfterAuto = Some true then Some 0.0
        else wordTwipsToPoints props.SpacingAfter
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
        "SpacingBefore" ==> (if props.SpacingBeforeAuto = Some true then Some 0.0 else wordTwipsToPoints props.SpacingBefore)
        "SpacingAfter" ==> spacingAfter
        "LineSpacing" ==> wordLineSpacingToMultiple fontSize props.LineSpacing props.LineRule
        "LineSpacingRule" ==> props.LineRule
        "BackgroundColor" ==> wordColorToCss props.BackgroundColor
        "ParagraphBorders" ==> paragraphBordersToEditor props.ParagraphBorders
        "RenderedPageBreakBefore" ==> props.RenderedPageBreakBefore
    ]

let private referenceObjects (refs: HeaderFooterReference list) =
    refs
    |> List.map (fun ref -> createObj [ "Kind" ==> ref.Kind; "RelationshipId" ==> ref.RelationshipId ])
    |> List.toArray

let private emptyRunProperties = {
    Style = None
    Bold = None
    Italic = None
    Underline = None
    AllCaps = None
    SmallCaps = None
    FontSize = None
    FontFamily = None
    ThemeFont = None
    Color = None
    ThemeColor = None
    ThemeTint = None
    ThemeShade = None
    Highlight = None
    Strike = None
    VertAlign = None
}

let private mergeRunProperties (baseProps: RunProperties) (props: RunProperties) : RunProperties = {
    Style = props.Style |> Option.orElse baseProps.Style
    Bold = props.Bold |> Option.orElse baseProps.Bold
    Italic = props.Italic |> Option.orElse baseProps.Italic
    Underline = props.Underline |> Option.orElse baseProps.Underline
    AllCaps = props.AllCaps |> Option.orElse baseProps.AllCaps
    SmallCaps = props.SmallCaps |> Option.orElse baseProps.SmallCaps
    Strike = props.Strike |> Option.orElse baseProps.Strike
    FontSize = props.FontSize |> Option.orElse baseProps.FontSize
    FontFamily = props.FontFamily |> Option.orElse baseProps.FontFamily
    ThemeFont = props.ThemeFont |> Option.orElse baseProps.ThemeFont
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
    SectionPageNumberStart = props.SectionPageNumberStart
    RenderedPageBreakBefore = props.RenderedPageBreakBefore
    HeaderReferences = props.HeaderReferences
    FooterReferences = props.FooterReferences
}

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
    mergeRunProperties (mergeRunProperties (mergeRunProperties (mergeRunProperties docDefaultRun baseRun) paragraphRun) characterRun) runProps

let private toEditorParagraphsInternal (theme: Theme option) (styles: StyleDefinition list) (content: DocumentContent list) (images: Map<string, string>) : obj array =
    let styleMap = buildStyleMap styles
    content
    |> List.choose (fun item ->
        match item with
        | Para para ->
            let effectiveParaProps = resolvedParagraphProperties styleMap para.Properties
            let paragraphFontSize =
                let fallbackRunProps = resolvedRunProperties styleMap effectiveParaProps emptyRunProperties
                let fallback = wordFontSizeToPoints fallbackRunProps.FontSize
                let runSizes =
                    para.Runs
                    |> List.choose (fun run ->
                        match run.Content with
                        | Text _
                        | Link _
                        | FieldRef _ ->
                            let runProps = resolvedRunProperties styleMap effectiveParaProps run.Properties
                            Some (wordFontSizeToPoints runProps.FontSize)
                        | _ -> None)
                if List.isEmpty runSizes then fallback else List.reduce max runSizes
            let pageBreakAfter =
                para.Runs
                |> List.exists (fun run -> match run.Content with | PageBreak -> true | _ -> false)
                || effectiveParaProps.SectionBreak = Some "nextPage"

            // Convert Run to EditorRun format
            let editorRuns =
                para.Runs
                |> List.choose (fun run ->
                    let effectiveRunProps = resolvedRunProperties styleMap effectiveParaProps run.Properties
                    match run.Content with
                    | Text text ->
                        // Text run
                        Some (createObj [
                            "Text" ==> text
                            "Style" ==> run.Properties.Style
                            "Bold" ==> effectiveOnOff effectiveRunProps.Bold
                            "Italic" ==> effectiveOnOff effectiveRunProps.Italic
                            "Underline" ==> effectiveOnOff effectiveRunProps.Underline
                            "AllCaps" ==> effectiveOnOff effectiveRunProps.AllCaps
                            "SmallCaps" ==> effectiveOnOff effectiveRunProps.SmallCaps
                            "Strike" ==> effectiveOnOff effectiveRunProps.Strike
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
                        ])
                    | Image img ->
                        // Image run - resolve relationship ID to data URL
                        let imageSrc =
                            images
                            |> Map.tryFind img.RelationshipId
                            |> Option.defaultValue img.RelationshipId  // Fallback to relId if not found
                        Some (createObj [
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
                        Some (createObj [
                            "Text" ==> link.Text
                            "Style" ==> link.Properties.Style
                            "Bold" ==> effectiveOnOff effectiveRunProps.Bold
                            "Italic" ==> effectiveOnOff effectiveRunProps.Italic
                            "Underline" ==> effectiveOnOff effectiveRunProps.Underline
                            "AllCaps" ==> effectiveOnOff effectiveRunProps.AllCaps
                            "SmallCaps" ==> effectiveOnOff effectiveRunProps.SmallCaps
                            "Strike" ==> effectiveOnOff effectiveRunProps.Strike
                            "FontSize" ==> wordFontSizeToPoints effectiveRunProps.FontSize
                            "FontFamily" ==> wordFontFamilyToCss theme effectiveRunProps
                            "Color" ==> wordColorToCss effectiveRunProps.Color
                            "Highlight" ==> effectiveRunProps.Highlight
                            "VertAlign" ==> effectiveRunProps.VertAlign
                            "IsImage" ==> false
                            "Hyperlink" ==> link.Target
                        ])
                        | CommentReference id ->
                            Some (createObj [
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
                            Some (createObj [
                                "Text" ==> (sprintf "[%s %s]" note.NoteType note.Id)
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
                                "NoteId" ==> Some note.Id
                                "NoteType" ==> Some note.NoteType
                            ])
                        | BookmarkStart bookmark ->
                            Some (createObj [
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
                            Some (createObj [
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
                            Some (createObj [
                                "Text" ==> (field.Result |> Option.defaultValue (sprintf "[Field:%s]" field.Code))
                                "Style" ==> run.Properties.Style
                                "Bold" ==> effectiveOnOff effectiveRunProps.Bold
                                "Italic" ==> effectiveOnOff effectiveRunProps.Italic
                                "Underline" ==> effectiveOnOff effectiveRunProps.Underline
                                "AllCaps" ==> effectiveOnOff effectiveRunProps.AllCaps
                                "SmallCaps" ==> effectiveOnOff effectiveRunProps.SmallCaps
                                "Strike" ==> effectiveOnOff effectiveRunProps.Strike
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
                            ])
                        | EquationXml xml ->
                            Some (createObj [
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

            Some (createObj ([
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
            ] @ paragraphLayoutProps paragraphFontSize effectiveParaProps))
        | Tbl table ->
            // Convert table to editor format
            let convertCellParagraph (para: Paragraph) : obj =
                let effectiveParaProps = resolvedParagraphProperties styleMap para.Properties
                let pageBreakAfter =
                    para.Runs
                    |> List.exists (fun run -> match run.Content with | PageBreak -> true | _ -> false)
                    || effectiveParaProps.SectionBreak = Some "nextPage"

                let editorRuns =
                    para.Runs
                    |> List.choose (fun run ->
                        let effectiveRunProps = resolvedRunProperties styleMap effectiveParaProps run.Properties
                        match run.Content with
                        | Text text ->
                            Some (createObj [
                                "Text" ==> text
                                "Style" ==> run.Properties.Style
                                "Bold" ==> effectiveOnOff effectiveRunProps.Bold
                                "Italic" ==> effectiveOnOff effectiveRunProps.Italic
                                "Underline" ==> effectiveOnOff effectiveRunProps.Underline
                                "AllCaps" ==> effectiveOnOff effectiveRunProps.AllCaps
                                "SmallCaps" ==> effectiveOnOff effectiveRunProps.SmallCaps
                                "Strike" ==> effectiveOnOff effectiveRunProps.Strike
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
                            ])
                        | Image img ->
                            let imageSrc =
                                images
                                |> Map.tryFind img.RelationshipId
                                |> Option.defaultValue img.RelationshipId
                            Some (createObj [
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
                            Some (createObj [
                                "Text" ==> link.Text
                                "Style" ==> link.Properties.Style
                                "Bold" ==> effectiveOnOff effectiveLinkProps.Bold
                                "Italic" ==> effectiveOnOff effectiveLinkProps.Italic
                                "Underline" ==> effectiveOnOff effectiveLinkProps.Underline
                                "AllCaps" ==> effectiveOnOff effectiveLinkProps.AllCaps
                                "SmallCaps" ==> effectiveOnOff effectiveLinkProps.SmallCaps
                                "Strike" ==> effectiveOnOff effectiveLinkProps.Strike
                                "FontSize" ==> wordFontSizeToPoints effectiveLinkProps.FontSize
                                "FontFamily" ==> wordFontFamilyToCss theme effectiveLinkProps
                                "Color" ==> wordColorToCss effectiveLinkProps.Color
                                "Highlight" ==> effectiveLinkProps.Highlight
                                "VertAlign" ==> effectiveLinkProps.VertAlign
                                "IsImage" ==> false
                                "Hyperlink" ==> link.Target
                            ])
                        | CommentReference id ->
                            Some (createObj [
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
                            Some (createObj [
                                "Text" ==> (sprintf "[%s %s]" note.NoteType note.Id)
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
                                "NoteId" ==> Some note.Id
                                "NoteType" ==> Some note.NoteType
                            ])
                        | BookmarkStart bookmark ->
                            Some (createObj [
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
                            Some (createObj [
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
                            Some (createObj [
                                "Text" ==> (field.Result |> Option.defaultValue (sprintf "[Field:%s]" field.Code))
                                "Style" ==> run.Properties.Style
                                "Bold" ==> effectiveOnOff effectiveRunProps.Bold
                                "Italic" ==> effectiveOnOff effectiveRunProps.Italic
                                "Underline" ==> effectiveOnOff effectiveRunProps.Underline
                                "AllCaps" ==> effectiveOnOff effectiveRunProps.AllCaps
                                "SmallCaps" ==> effectiveOnOff effectiveRunProps.SmallCaps
                                "Strike" ==> effectiveOnOff effectiveRunProps.Strike
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
                            ])
                        | EquationXml xml ->
                            Some (createObj [
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

                createObj ([
                    "Runs" ==> editorRuns
                    "Style" ==> effectiveParaProps.Style
                    "StyleKeepNext" ==> effectiveParaProps.KeepNext
                    "StyleKeepLines" ==> effectiveParaProps.KeepLines
                    "StylePageBreakBefore" ==> effectiveParaProps.PageBreakBefore
                    "PageBreakBefore" ==> effectiveParaProps.PageBreakBefore
                    "StyleWidowControl" ==> effectiveParaProps.WidowControl
                    "StyleContextualSpacing" ==> effectiveParaProps.ContextualSpacing
                    "Alignment" ==> (effectiveParaProps.Alignment |> Option.defaultValue "left")
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
                ] @ paragraphLayoutProps paragraphFontSize effectiveParaProps)

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
                let baseProperties = tableStyle |> Option.bind (fun style -> style.ParagraphProperties)
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
                    |> List.map convertCellParagraph
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

                createObj [
                    "Paragraphs" ==> cellParagraphs
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

                createObj [
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

            Some (createObj [
                "Type" ==> "Table"
                "Rows" ==> rows
                "PreferredWidth" ==> table.Width
                "PreferredWidthType" ==> table.WidthType
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
            match unbox<string> value with
            | null | "" -> None
            | text -> Some text
    let intValue fallback (value: obj) =
        if isMissing value then fallback else unbox<int> value
    let stringValue fallback (value: obj) = optionalStringValue value |> Option.defaultValue fallback

    let numberingMap =
        numberingDefinitions
        |> List.map (fun def -> ((def.NumId, def.Level), def))
        |> Map.ofList
    let styleNumberingMap =
        numberingDefinitions
        |> List.choose (fun def -> def.ParagraphStyle |> Option.map (fun styleId -> (styleId, def)))
        |> Map.ofList

    let applyToParagraph (item: obj) =
        let listId = optionalStringValue (item?ListId)
        let listLevel = intValue 0 (item?ListLevel)
        let styleId = optionalStringValue (item?Style)
        let listDefinition =
            match listId with
            | Some id -> Map.tryFind (id, max 0 (listLevel - 1)) numberingMap
            | None -> styleId |> Option.bind (fun id -> Map.tryFind id styleNumberingMap)
        match listDefinition with
        | Some def when def.NumberFormat = "bullet" ->
            item?ListId <- def.NumId
            item?ListLevel <- def.Level + 1
            item?ListStyle <- "bullet"
            item?ListNumberFormat <- def.NumberFormat
            item?ListLevelText <- def.LevelText
            item?ListStart <- def.Start
        | Some def ->
            item?ListId <- def.NumId
            item?ListLevel <- def.Level + 1
            item?ListStyle <- "numbered"
            item?ListNumberFormat <- def.NumberFormat
            item?ListLevelText <- def.LevelText
            item?ListStart <- def.Start
        | None -> ()
        match listDefinition with
        | Some def ->
            if isMissing item?IndentLeft then item?IndentLeft <- wordTwipsToPoints def.IndentLeft
            if isMissing item?FirstLineIndent then
                let firstLineIndent =
                    match wordTwipsToPoints def.FirstLineIndent, wordTwipsToPoints def.HangingIndent with
                    | Some firstLine, _ -> Some firstLine
                    | None, Some hanging -> Some -hanging
                    | None, None -> None
                item?FirstLineIndent <- firstLineIndent
            if isMissing item?RightTabPosition then item?RightTabPosition <- wordTwipsToPoints def.RightTabPosition
        | None -> ()
        item

    let rec apply item =
        let itemType = item?Type |> stringValue "Paragraph"
        if itemType = "Table" then
            let rows = item?Rows |> unbox<obj array>
            for row in rows do
                let cells = row?Cells |> unbox<obj array>
                for cell in cells do
                    let paragraphs = cell?Paragraphs |> unbox<obj array>
                    for para in paragraphs do
                        applyToParagraph para |> ignore
            item
        else
            applyToParagraph item

    items |> Array.map apply

let toEditorParagraphs (content: DocumentContent list) (images: Map<string, string>) : obj array =
    toEditorParagraphsInternal None [] content images

let toEditorParagraphsWithStyles (content: DocumentContent list) (images: Map<string, string>) (styles: StyleDefinition list) : obj array =
    toEditorParagraphsInternal None styles content images

let toEditorParagraphsWithStylesAndTheme (theme: Theme option) (content: DocumentContent list) (images: Map<string, string>) (styles: StyleDefinition list) : obj array =
    toEditorParagraphsInternal theme styles content images

let toEditorParagraphsWithStylesAndNumbering (content: DocumentContent list) (images: Map<string, string>) (styles: StyleDefinition list) (numberingDefinitions: NumberingDefinition list) : obj array =
    toEditorParagraphsInternal None styles content images
    |> applyNumberingDefinitions numberingDefinitions

let toEditorParagraphsWithStylesThemeAndNumbering (theme: Theme option) (content: DocumentContent list) (images: Map<string, string>) (styles: StyleDefinition list) (numberingDefinitions: NumberingDefinition list) : obj array =
    toEditorParagraphsInternal theme styles content images
    |> applyNumberingDefinitions numberingDefinitions

let toEditorHeaderFooters (items: HeaderFooter list) (images: Map<string, string>) : obj array =
    items
    |> List.map (fun item ->
        createObj [
            "Type" ==> item.Type
            "RelationshipId" ==> item.RelationshipId
            "Paragraphs" ==> toEditorParagraphs (item.Content |> List.map Para) images
        ])
    |> List.toArray

let toEditorHeaderFootersWithStyles (items: HeaderFooter list) (images: Map<string, string>) (styles: StyleDefinition list) : obj array =
    items
    |> List.map (fun item ->
        createObj [
            "Type" ==> item.Type
            "RelationshipId" ==> item.RelationshipId
            "Paragraphs" ==> toEditorParagraphsWithStyles (item.Content |> List.map Para) images styles
        ])
    |> List.toArray

let toEditorHeaderFootersWithStylesAndTheme (theme: Theme option) (items: HeaderFooter list) (images: Map<string, string>) (styles: StyleDefinition list) : obj array =
    items
    |> List.map (fun item ->
        createObj [
            "Type" ==> item.Type
            "RelationshipId" ==> item.RelationshipId
            "Paragraphs" ==> toEditorParagraphsWithStylesAndTheme theme (item.Content |> List.map Para) images styles
        ])
    |> List.toArray

let toEditorPageSetup (doc: WordDocument) : obj =
    match doc.PageSetup with
    | None -> createObj []
    | Some setup ->
        createObj [
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
        ]
