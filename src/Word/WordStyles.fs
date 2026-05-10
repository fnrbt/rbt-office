module RedBat.Word.WordStyles

open RedBat.Word.Document
open RedBat.Word.TextBuffer

type ParagraphStyle = {
    Text: TextStyle
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
    PageBreakBefore: bool option
    KeepNext: bool option
    KeepLines: bool option
    WidowControl: bool option
}

let private styleMap (doc: Document) =
    doc.Styles |> List.map (fun style -> (style.StyleType, style.Id), style) |> Map.ofList

let private documentDefaults (doc: Document) =
    doc.Styles |> List.tryFind (fun style -> style.StyleType = "defaults")

let private mergeTextStyle (baseStyle: TextStyle) (overrideStyle: TextStyle) =
    {
        overrideStyle with
            Bold = baseStyle.Bold || overrideStyle.Bold
            Italic = baseStyle.Italic || overrideStyle.Italic
            Underline = baseStyle.Underline || overrideStyle.Underline
            AllCaps = baseStyle.AllCaps || overrideStyle.AllCaps
            SmallCaps = baseStyle.SmallCaps || overrideStyle.SmallCaps
            Strike = baseStyle.Strike || overrideStyle.Strike
            FontSize = if overrideStyle.FontSize = TextBuffer.defaultStyle.FontSize then baseStyle.FontSize else overrideStyle.FontSize
            FontFamily = overrideStyle.FontFamily |> Option.orElse baseStyle.FontFamily
            Color = overrideStyle.Color |> Option.orElse baseStyle.Color
            Highlight = overrideStyle.Highlight |> Option.orElse baseStyle.Highlight
            EquationXml = overrideStyle.EquationXml |> Option.orElse baseStyle.EquationXml
            RevisionId = overrideStyle.RevisionId |> Option.orElse baseStyle.RevisionId
            RevisionType = overrideStyle.RevisionType |> Option.orElse baseStyle.RevisionType
            RevisionAuthor = overrideStyle.RevisionAuthor |> Option.orElse baseStyle.RevisionAuthor
            RevisionDate = overrideStyle.RevisionDate |> Option.orElse baseStyle.RevisionDate
            VertAlign = overrideStyle.VertAlign |> Option.orElse baseStyle.VertAlign
    }

let private mergeNamedStyle (baseStyle: NamedStyle) (overrideStyle: NamedStyle) =
    {
        overrideStyle with
            Text =
                match baseStyle.Text, overrideStyle.Text with
                | Some baseText, Some overrideText -> Some (mergeTextStyle baseText overrideText)
                | Some baseText, None -> Some baseText
                | None, Some overrideText -> Some overrideText
                | None, None -> None
            SpacingBefore = overrideStyle.SpacingBefore |> Option.orElse baseStyle.SpacingBefore
            SpacingAfter = overrideStyle.SpacingAfter |> Option.orElse baseStyle.SpacingAfter
            ContextualSpacing = overrideStyle.ContextualSpacing |> Option.orElse baseStyle.ContextualSpacing
            LineSpacing = overrideStyle.LineSpacing |> Option.orElse baseStyle.LineSpacing
            LineSpacingRule = overrideStyle.LineSpacingRule |> Option.orElse baseStyle.LineSpacingRule
            Alignment = overrideStyle.Alignment |> Option.orElse baseStyle.Alignment
            IndentLeft = overrideStyle.IndentLeft |> Option.orElse baseStyle.IndentLeft
            IndentRight = overrideStyle.IndentRight |> Option.orElse baseStyle.IndentRight
            FirstLineIndent = overrideStyle.FirstLineIndent |> Option.orElse baseStyle.FirstLineIndent
            BackgroundColor = overrideStyle.BackgroundColor |> Option.orElse baseStyle.BackgroundColor
            OutlineLevel = overrideStyle.OutlineLevel |> Option.orElse baseStyle.OutlineLevel
            KeepNext = overrideStyle.KeepNext |> Option.orElse baseStyle.KeepNext
            KeepLines = overrideStyle.KeepLines |> Option.orElse baseStyle.KeepLines
            PageBreakBefore = overrideStyle.PageBreakBefore |> Option.orElse baseStyle.PageBreakBefore
            WidowControl = overrideStyle.WidowControl |> Option.orElse baseStyle.WidowControl
    }

let private resolveNamedStyle (styles: Map<(string * string), NamedStyle>) styleType styleId =
    let rec resolve visited id =
        if Set.contains id visited then None
        else
            match Map.tryFind (styleType, id) styles with
            | None -> None
            | Some style ->
                match style.BasedOn with
                | Some basedOn when basedOn <> id ->
                    match resolve (Set.add id visited) basedOn with
                    | Some baseStyle -> Some (mergeNamedStyle baseStyle style)
                    | None -> Some style
                | _ -> Some style
    resolve Set.empty styleId

let private applyNamedParagraphStyle (styles: Map<(string * string), NamedStyle>) (metadata: BlockMetadata) (paragraphStyle: ParagraphStyle) =
    match metadata.Style |> Option.bind (resolveNamedStyle styles "paragraph") with
    | None -> paragraphStyle
    | Some named ->
        { paragraphStyle with
            Text = named.Text |> Option.map (fun text -> mergeTextStyle paragraphStyle.Text text) |> Option.defaultValue paragraphStyle.Text
            SpacingBefore = named.SpacingBefore |> Option.orElse paragraphStyle.SpacingBefore
            SpacingAfter = named.SpacingAfter |> Option.orElse paragraphStyle.SpacingAfter
            ContextualSpacing = named.ContextualSpacing |> Option.orElse paragraphStyle.ContextualSpacing
            LineSpacing = named.LineSpacing |> Option.orElse paragraphStyle.LineSpacing
            LineSpacingRule = named.LineSpacingRule |> Option.orElse paragraphStyle.LineSpacingRule
            Alignment = named.Alignment |> Option.orElse paragraphStyle.Alignment
            IndentLeft = named.IndentLeft |> Option.orElse paragraphStyle.IndentLeft
            IndentRight = named.IndentRight |> Option.orElse paragraphStyle.IndentRight
            FirstLineIndent = named.FirstLineIndent |> Option.orElse paragraphStyle.FirstLineIndent
            BackgroundColor = named.BackgroundColor |> Option.orElse paragraphStyle.BackgroundColor
            OutlineLevel = named.OutlineLevel |> Option.orElse paragraphStyle.OutlineLevel
            PageBreakBefore = named.PageBreakBefore |> Option.orElse paragraphStyle.PageBreakBefore
            KeepNext = named.KeepNext |> Option.orElse paragraphStyle.KeepNext
            KeepLines = named.KeepLines |> Option.orElse paragraphStyle.KeepLines
            WidowControl = named.WidowControl |> Option.orElse paragraphStyle.WidowControl }

let normal = {
    Text = { TextBuffer.defaultStyle with FontFamily = Some "Calibri" }
    SpacingBefore = None
    SpacingAfter = Some 10.0
    ContextualSpacing = None
    LineSpacing = None
    LineSpacingRule = None
    Alignment = None
    IndentLeft = None
    IndentRight = None
    FirstLineIndent = None
    BackgroundColor = None
    OutlineLevel = None
    PageBreakBefore = None
    KeepNext = None
    KeepLines = None
    WidowControl = None
}

let private applyDocumentDefaults doc (paragraphStyle: ParagraphStyle) : ParagraphStyle =
    match documentDefaults doc with
    | None -> paragraphStyle
    | Some defaults ->
        { paragraphStyle with
            Text = defaults.Text |> Option.map (fun text -> mergeTextStyle paragraphStyle.Text text) |> Option.defaultValue paragraphStyle.Text
            SpacingBefore = defaults.SpacingBefore |> Option.orElse paragraphStyle.SpacingBefore
            SpacingAfter = defaults.SpacingAfter |> Option.orElse paragraphStyle.SpacingAfter
            ContextualSpacing = defaults.ContextualSpacing |> Option.orElse paragraphStyle.ContextualSpacing
            LineSpacing = defaults.LineSpacing |> Option.orElse paragraphStyle.LineSpacing
            LineSpacingRule = defaults.LineSpacingRule |> Option.orElse paragraphStyle.LineSpacingRule
            Alignment = defaults.Alignment |> Option.orElse paragraphStyle.Alignment
            IndentLeft = defaults.IndentLeft |> Option.orElse paragraphStyle.IndentLeft
            IndentRight = defaults.IndentRight |> Option.orElse paragraphStyle.IndentRight
            FirstLineIndent = defaults.FirstLineIndent |> Option.orElse paragraphStyle.FirstLineIndent
            BackgroundColor = defaults.BackgroundColor |> Option.orElse paragraphStyle.BackgroundColor
            OutlineLevel = defaults.OutlineLevel |> Option.orElse paragraphStyle.OutlineLevel
            PageBreakBefore = defaults.PageBreakBefore |> Option.orElse paragraphStyle.PageBreakBefore
            KeepNext = defaults.KeepNext |> Option.orElse paragraphStyle.KeepNext
            KeepLines = defaults.KeepLines |> Option.orElse paragraphStyle.KeepLines
            WidowControl = defaults.WidowControl |> Option.orElse paragraphStyle.WidowControl }

let private headingText size = {
    TextBuffer.defaultStyle with
        Bold = true
        FontSize = size
        FontFamily = Some "Calibri"
}

let heading level =
    match level with
    | 1 -> { normal with Text = headingText 22.0; SpacingBefore = Some 0.0; SpacingAfter = Some 12.0; LineSpacing = Some 1.2; LineSpacingRule = Some "auto"; OutlineLevel = Some 1 }
    | 2 -> { normal with Text = headingText 18.0; SpacingBefore = Some 12.0; SpacingAfter = Some 8.0; LineSpacing = Some 1.2; LineSpacingRule = Some "auto"; OutlineLevel = Some 2 }
    | 3 -> { normal with Text = headingText 16.0; SpacingBefore = Some 10.0; SpacingAfter = Some 6.0; LineSpacing = Some 1.2; LineSpacingRule = Some "auto"; OutlineLevel = Some 3 }
    | _ -> normal

let forBlock = function
    | Heading (level, _) -> heading level
    | Paragraph _
    | Table _
    | TextBox _ -> normal

let forBlockInDocument (doc: Document) block =
    let baseStyle = forBlock block |> applyDocumentDefaults doc
    match block with
    | Paragraph para
    | Heading (_, para) -> applyNamedParagraphStyle (styleMap doc) para.Metadata baseStyle
    | Table _
    | TextBox _ -> baseStyle

let forTableCellParagraphInDocument (doc: Document) (para: ParagraphBlock) =
    let style = forBlockInDocument doc (Paragraph para)
    { style with
        SpacingBefore = Some 0.0
        SpacingAfter = Some 0.0
        LineSpacing = Some 1.0
        LineSpacingRule = Some "auto" }

let tableCellParagraphForLayout (para: ParagraphBlock) =
    { para with
        Metadata =
            { para.Metadata with
                SpacingBefore = Some 0.0
                SpacingAfter = Some 0.0
                LineSpacing = Some 1.0
                LineSpacingRule = Some "auto" } }

let private inheritFontFamily baseStyle runStyle =
    match runStyle.FontFamily with
    | Some _ -> runStyle.FontFamily
    | None -> baseStyle.FontFamily

let effectiveTextStyle (paragraphStyle: ParagraphStyle) (runStyle: TextStyle) =
    let defaultStyle = TextBuffer.defaultStyle
    {
        runStyle with
            Bold = paragraphStyle.Text.Bold || runStyle.Bold
            Italic = paragraphStyle.Text.Italic || runStyle.Italic
            FontSize = if runStyle.FontSize = defaultStyle.FontSize then paragraphStyle.Text.FontSize else runStyle.FontSize
            FontFamily = inheritFontFamily paragraphStyle.Text runStyle
            Color = runStyle.Color |> Option.orElse paragraphStyle.Text.Color
            Highlight = runStyle.Highlight |> Option.orElse paragraphStyle.Text.Highlight
    }

let effectiveTextStyleInDocument (doc: Document) (paragraphStyle: ParagraphStyle) (runStyle: TextStyle) =
    let styles = styleMap doc
    let styledRun =
        match runStyle.Style |> Option.bind (resolveNamedStyle styles "character") |> Option.bind (fun style -> style.Text) with
        | Some namedText -> mergeTextStyle namedText runStyle
        | None -> runStyle
    effectiveTextStyle paragraphStyle styledRun

let effectiveBuffer paragraphStyle (buffer: TextBuffer) =
    { buffer with
        Runs =
            buffer.Runs
            |> List.map (function
                | Text run -> Text { run with Style = effectiveTextStyle paragraphStyle run.Style }
                | Image image -> Image image) }

let effectiveBufferInDocument doc paragraphStyle (buffer: TextBuffer) =
    { buffer with
        Runs =
            buffer.Runs
            |> List.map (function
                | Text run -> Text { run with Style = effectiveTextStyleInDocument doc paragraphStyle run.Style }
                | Image image -> Image image) }

let spacingBefore (paragraphStyle: ParagraphStyle) (metadata: BlockMetadata) = metadata.SpacingBefore |> Option.orElse paragraphStyle.SpacingBefore |> Option.defaultValue 0.0
let spacingAfter (paragraphStyle: ParagraphStyle) (metadata: BlockMetadata) = metadata.SpacingAfter |> Option.orElse paragraphStyle.SpacingAfter |> Option.defaultValue 10.0
let lineSpacing (paragraphStyle: ParagraphStyle) (metadata: BlockMetadata) = metadata.LineSpacing |> Option.orElse paragraphStyle.LineSpacing |> Option.defaultValue 1.0
let lineSpacingRule (paragraphStyle: ParagraphStyle) (metadata: BlockMetadata) = metadata.LineSpacingRule |> Option.orElse paragraphStyle.LineSpacingRule |> Option.defaultValue "auto"
let widowControl (paragraphStyle: ParagraphStyle) (metadata: BlockMetadata) = metadata.WidowControl |> Option.orElse paragraphStyle.WidowControl |> Option.defaultValue true

let contextualSpacingBetween (paragraphStyle: ParagraphStyle) (current: ParagraphBlock) (next: ParagraphBlock option) =
    let enabled = current.Metadata.ContextualSpacing |> Option.orElse paragraphStyle.ContextualSpacing |> Option.defaultValue false
    enabled && next |> Option.exists (fun value -> current.Metadata.Style.IsSome && current.Metadata.Style = value.Metadata.Style)

let effectiveMetadata (paragraphStyle: ParagraphStyle) (metadata: BlockMetadata) =
    { metadata with
        Alignment = if metadata.Alignment = Left then paragraphStyle.Alignment |> Option.defaultValue metadata.Alignment else metadata.Alignment
        IndentLeft = metadata.IndentLeft |> Option.orElse paragraphStyle.IndentLeft
        IndentRight = metadata.IndentRight |> Option.orElse paragraphStyle.IndentRight
        FirstLineIndent = metadata.FirstLineIndent |> Option.orElse paragraphStyle.FirstLineIndent
        ContextualSpacing = metadata.ContextualSpacing |> Option.orElse paragraphStyle.ContextualSpacing
        LineSpacingRule = metadata.LineSpacingRule |> Option.orElse paragraphStyle.LineSpacingRule
        WidowControl = metadata.WidowControl |> Option.orElse paragraphStyle.WidowControl
        BackgroundColor = metadata.BackgroundColor |> Option.orElse paragraphStyle.BackgroundColor }
