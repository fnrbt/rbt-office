module Rbt.Word.WordStyles

open Rbt.Word.Document
open Rbt.Word.TextBuffer

let private resizeArrayToList (items: ResizeArray<'T>) =
    let mutable result = []
    for i = items.Count - 1 downto 0 do
        result <- items.[i] :: result
    result

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

type DocumentStyleContext = {
    Styles: Map<(string * string), NamedStyle>
    ResolvedParagraphStyles: Map<string, NamedStyle>
    ResolvedCharacterStyles: Map<string, NamedStyle>
    NormalStyle: ParagraphStyle
    DefaultParagraphStyle: ParagraphStyle
    ParagraphStyles: Map<string, ParagraphStyle>
    Defaults: NamedStyle option
    DefaultParagraphStyleId: string option
}

/// An on/off property of `style` (the overriding run or style) against the
/// value it would inherit. Set directly (TextStyle.Explicit) it wins either
/// way, so a run can be non-bold under a bold style; otherwise "off" is the
/// unset value and the inherited value shows through.
let private inheritToggle (flag: int) (inherited: bool) (own: bool) (style: TextStyle) =
    if ExplicitFormat.has flag style then own else inherited || own

/// Font size likewise: the default size doubles as "unset" unless the run
/// carries the explicit flag.
let private inheritFontSize (inherited: float) (style: TextStyle) =
    if ExplicitFormat.has ExplicitFormat.FontSize style || style.FontSize <> TextBuffer.defaultStyle.FontSize then style.FontSize
    else inherited

let private mergeTextStyle (baseStyle: TextStyle) (overrideStyle: TextStyle) =
    {
        overrideStyle with
            Bold = inheritToggle ExplicitFormat.Bold baseStyle.Bold overrideStyle.Bold overrideStyle
            Italic = inheritToggle ExplicitFormat.Italic baseStyle.Italic overrideStyle.Italic overrideStyle
            Underline = inheritToggle ExplicitFormat.Underline baseStyle.Underline overrideStyle.Underline overrideStyle
            UnderlineStyle = overrideStyle.UnderlineStyle |> Option.orElse baseStyle.UnderlineStyle
            UnderlineColor = overrideStyle.UnderlineColor |> Option.orElse baseStyle.UnderlineColor
            AllCaps = inheritToggle ExplicitFormat.AllCaps baseStyle.AllCaps overrideStyle.AllCaps overrideStyle
            SmallCaps = inheritToggle ExplicitFormat.SmallCaps baseStyle.SmallCaps overrideStyle.SmallCaps overrideStyle
            Strike = inheritToggle ExplicitFormat.Strike baseStyle.Strike overrideStyle.Strike overrideStyle
            DoubleStrike = inheritToggle ExplicitFormat.DoubleStrike baseStyle.DoubleStrike overrideStyle.DoubleStrike overrideStyle
            Hidden = inheritToggle ExplicitFormat.Hidden baseStyle.Hidden overrideStyle.Hidden overrideStyle
            Position = if abs overrideStyle.Position < 0.001 then baseStyle.Position else overrideStyle.Position
            EmphasisMark = overrideStyle.EmphasisMark |> Option.orElse baseStyle.EmphasisMark
            FontSize = inheritFontSize baseStyle.FontSize overrideStyle
            CharacterSpacing = if abs overrideStyle.CharacterSpacing < 0.001 then baseStyle.CharacterSpacing else overrideStyle.CharacterSpacing
            FontFamily = overrideStyle.FontFamily |> Option.orElse baseStyle.FontFamily
            Color = overrideStyle.Color |> Option.orElse baseStyle.Color
            Highlight = overrideStyle.Highlight |> Option.orElse baseStyle.Highlight
            EquationXml = overrideStyle.EquationXml |> Option.orElse baseStyle.EquationXml
            RevisionId = overrideStyle.RevisionId |> Option.orElse baseStyle.RevisionId
            RevisionType = overrideStyle.RevisionType |> Option.orElse baseStyle.RevisionType
            RevisionAuthor = overrideStyle.RevisionAuthor |> Option.orElse baseStyle.RevisionAuthor
            RevisionDate = overrideStyle.RevisionDate |> Option.orElse baseStyle.RevisionDate
            DeletedInsertion = overrideStyle.DeletedInsertion |> Option.orElse baseStyle.DeletedInsertion
            MoveName = overrideStyle.MoveName |> Option.orElse baseStyle.MoveName
            VertAlign = overrideStyle.VertAlign |> Option.orElse baseStyle.VertAlign
            EastAsiaLanguage = overrideStyle.EastAsiaLanguage |> Option.orElse baseStyle.EastAsiaLanguage
    }

let private mergeNamedTextStyle (baseStyle: TextStyle) (named: NamedStyle) =
    match named.Text with
    | None -> baseStyle
    | Some overrideStyle ->
        let merged = mergeTextStyle baseStyle overrideStyle
        {
            merged with
                Bold = named.TextBold |> Option.defaultValue merged.Bold
                Italic = named.TextItalic |> Option.defaultValue merged.Italic
                Underline = named.TextUnderline |> Option.defaultValue merged.Underline
                AllCaps = named.TextAllCaps |> Option.defaultValue merged.AllCaps
                SmallCaps = named.TextSmallCaps |> Option.defaultValue merged.SmallCaps
                Strike = named.TextStrike |> Option.defaultValue merged.Strike
        }

let private mergeNamedStyle (baseStyle: NamedStyle) (overrideStyle: NamedStyle) =
    {
        overrideStyle with
            Text =
                match baseStyle.Text, overrideStyle.Text with
                | Some baseText, Some _ -> Some (mergeNamedTextStyle baseText overrideStyle)
                | Some baseText, None -> Some baseText
                | None, Some overrideText -> Some overrideText
                | None, None -> None
            TextBold = overrideStyle.TextBold |> Option.orElse baseStyle.TextBold
            TextItalic = overrideStyle.TextItalic |> Option.orElse baseStyle.TextItalic
            TextUnderline = overrideStyle.TextUnderline |> Option.orElse baseStyle.TextUnderline
            TextAllCaps = overrideStyle.TextAllCaps |> Option.orElse baseStyle.TextAllCaps
            TextSmallCaps = overrideStyle.TextSmallCaps |> Option.orElse baseStyle.TextSmallCaps
            TextStrike = overrideStyle.TextStrike |> Option.orElse baseStyle.TextStrike
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

let private defaultParagraphStyleId (styles: Map<(string * string), NamedStyle>) =
    // Word treats the styles.xml paragraph style marked w:default="1" (usually
    // "Normal") as the implicit style for any paragraph that doesn't set its
    // own pStyle.  Without this fallback, body paragraphs miss the Normal
    // style's line-spacing / indent / spacing-after values and accumulate
    // pixel drift across the page.
    styles
    |> Map.toSeq
    |> Seq.tryPick (fun ((styleType, styleId), named) ->
        if styleType = "paragraph" && named.IsDefault then Some styleId else None)

let private resolvedStyleMap (styles: Map<(string * string), NamedStyle>) styleType =
    styles
    |> Map.toSeq
    |> Seq.choose (fun ((candidateType, styleId), _) ->
        if candidateType = styleType then
            resolveNamedStyle styles styleType styleId
            |> Option.map (fun resolved -> styleId, resolved)
        else
            None)
    |> Map.ofSeq

let private applyNamedParagraphStyleWithDefault (styles: Map<(string * string), NamedStyle>) defaultStyleId (metadata: BlockMetadata) (paragraphStyle: ParagraphStyle) =
    let effectiveStyleId =
        match metadata.Style with
        | Some _ as explicit -> explicit
        | None -> defaultStyleId
    match effectiveStyleId |> Option.bind (resolveNamedStyle styles "paragraph") with
    | None -> paragraphStyle
    | Some named ->
        { paragraphStyle with
            Text = mergeNamedTextStyle paragraphStyle.Text named
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

let private applyResolvedParagraphStyleWithDefault (resolvedStyles: Map<string, NamedStyle>) defaultStyleId (metadata: BlockMetadata) (paragraphStyle: ParagraphStyle) =
    let effectiveStyleId =
        match metadata.Style with
        | Some _ as explicit -> explicit
        | None -> defaultStyleId
    match effectiveStyleId |> Option.bind (fun styleId -> Map.tryFind styleId resolvedStyles) with
    | None -> paragraphStyle
    | Some named ->
        { paragraphStyle with
            Text = mergeNamedTextStyle paragraphStyle.Text named
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

let private applyNamedParagraphStyle (styles: Map<(string * string), NamedStyle>) metadata paragraphStyle =
    applyNamedParagraphStyleWithDefault styles (defaultParagraphStyleId styles) metadata paragraphStyle

let normal = {
    Text = { TextBuffer.defaultStyle with FontFamily = Some "Calibri" }
    SpacingBefore = None
    SpacingAfter = Some 8.0
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
            Text = mergeNamedTextStyle paragraphStyle.Text defaults
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

let private applyDocumentDefaultsFrom defaults (paragraphStyle: ParagraphStyle) : ParagraphStyle =
    match defaults with
    | None -> paragraphStyle
    | Some defaults ->
        { paragraphStyle with
            Text = mergeNamedTextStyle paragraphStyle.Text defaults
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

let buildStyleContext (doc: Document) =
    let styles = styleMap doc
    let resolvedParagraphStyles = resolvedStyleMap styles "paragraph"
    let defaultParagraphStyleId = defaultParagraphStyleId styles
    let defaults = documentDefaults doc
    let normalStyle = normal |> applyDocumentDefaultsFrom defaults
    let styleMetadata styleId =
        { defaultMetadata with Style = Some styleId }
    let paragraphStyles =
        resolvedParagraphStyles
        |> Map.toSeq
        |> Seq.map (fun (styleId, _) ->
            styleId,
            applyResolvedParagraphStyleWithDefault resolvedParagraphStyles defaultParagraphStyleId (styleMetadata styleId) normalStyle)
        |> Map.ofSeq
    {
        Styles = styles
        ResolvedParagraphStyles = resolvedParagraphStyles
        ResolvedCharacterStyles = resolvedStyleMap styles "character"
        NormalStyle = normalStyle
        DefaultParagraphStyle = applyResolvedParagraphStyleWithDefault resolvedParagraphStyles defaultParagraphStyleId defaultMetadata normalStyle
        ParagraphStyles = paragraphStyles
        Defaults = defaults
        DefaultParagraphStyleId = defaultParagraphStyleId
    }

/// The style contexts built last, with the style lists they were built from.
/// A context depends on the document's style list alone, and that list is
/// the same object from one edit to the next: resolving every style again for
/// each layout and paint pass (and once per table) was 30-50 ms of a
/// keystroke.
let mutable private knownStyleContexts : (NamedStyle list * DocumentStyleContext) list = []

let styleContext (doc: Document) =
    match knownStyleContexts |> List.tryFind (fun (styles, _) -> obj.ReferenceEquals(styles, doc.Styles)) with
    | Some (_, context) -> context
    | None ->
        let context = buildStyleContext doc
        knownStyleContexts <- (doc.Styles, context) :: List.truncate 3 knownStyleContexts
        context

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

let forBlockInContext (context: DocumentStyleContext) block =
    match block with
    | Paragraph para ->
        match para.Metadata.Style with
        // (A style the document does not define - "Header" on a header made here - is the
        // default paragraph style to Word, not the bare document defaults.)
        | Some styleId -> context.ParagraphStyles |> Map.tryFind styleId |> Option.defaultValue context.DefaultParagraphStyle
        | None -> context.DefaultParagraphStyle
    | Heading (level, para) ->
        match para.Metadata.Style with
        | Some styleId ->
            match context.ParagraphStyles |> Map.tryFind styleId with
            | Some style -> style
            | None ->
                heading level
                |> applyDocumentDefaultsFrom context.Defaults
                |> applyResolvedParagraphStyleWithDefault context.ResolvedParagraphStyles context.DefaultParagraphStyleId para.Metadata
        | None ->
            forBlock block
            |> applyDocumentDefaultsFrom context.Defaults
            |> applyResolvedParagraphStyleWithDefault context.ResolvedParagraphStyles context.DefaultParagraphStyleId para.Metadata
    | Table _
    | TextBox _ -> context.NormalStyle

let forBlockInDocument (doc: Document) block =
    forBlockInContext (styleContext doc) block

let forTableCellParagraphInDocument (doc: Document) (para: ParagraphBlock) =
    let style = forBlockInDocument doc (Paragraph para)
    { style with
        SpacingBefore = Some 0.0 }

let forTableCellParagraphInContext (context: DocumentStyleContext) (para: ParagraphBlock) =
    let style = forBlockInContext context (Paragraph para)
    { style with
        SpacingBefore = Some 0.0 }

let tableCellParagraphForLayout (para: ParagraphBlock) =
    { para with
        Metadata =
            { para.Metadata with
                SpacingBefore = para.Metadata.SpacingBefore |> Option.orElse (Some 0.0) } }

let private inheritFontFamily baseStyle runStyle =
    match runStyle.FontFamily with
    | Some _ -> runStyle.FontFamily
    | None -> baseStyle.FontFamily

let effectiveTextStyle (paragraphStyle: ParagraphStyle) (runStyle: TextStyle) =
    let inherited = paragraphStyle.Text
    {
        runStyle with
            Bold = inheritToggle ExplicitFormat.Bold inherited.Bold runStyle.Bold runStyle
            Italic = inheritToggle ExplicitFormat.Italic inherited.Italic runStyle.Italic runStyle
            Underline = inheritToggle ExplicitFormat.Underline inherited.Underline runStyle.Underline runStyle
            AllCaps = inheritToggle ExplicitFormat.AllCaps inherited.AllCaps runStyle.AllCaps runStyle
            SmallCaps = inheritToggle ExplicitFormat.SmallCaps inherited.SmallCaps runStyle.SmallCaps runStyle
            Strike = inheritToggle ExplicitFormat.Strike inherited.Strike runStyle.Strike runStyle
            DoubleStrike = inheritToggle ExplicitFormat.DoubleStrike inherited.DoubleStrike runStyle.DoubleStrike runStyle
            Hidden = inheritToggle ExplicitFormat.Hidden inherited.Hidden runStyle.Hidden runStyle
            FontSize = inheritFontSize inherited.FontSize runStyle
            CharacterSpacing = if abs runStyle.CharacterSpacing < 0.001 then paragraphStyle.Text.CharacterSpacing else runStyle.CharacterSpacing
            FontFamily = inheritFontFamily paragraphStyle.Text runStyle
            // "Automatic" colour stated on the run: the style's colour does not apply.
            Color = if ExplicitFormat.has ExplicitFormat.ColorAuto runStyle then runStyle.Color else runStyle.Color |> Option.orElse paragraphStyle.Text.Color
            Highlight = runStyle.Highlight |> Option.orElse paragraphStyle.Text.Highlight
    }

let private effectiveTextStyleWithStyles (styles: Map<(string * string), NamedStyle>) (paragraphStyle: ParagraphStyle) (runStyle: TextStyle) =
    let styledRun =
        match runStyle.Style |> Option.bind (resolveNamedStyle styles "character") |> Option.bind (fun style -> style.Text) with
        | Some namedText -> mergeTextStyle namedText runStyle
        | None -> runStyle
    effectiveTextStyle paragraphStyle styledRun

let private effectiveTextStyleWithResolved (resolvedStyles: Map<string, NamedStyle>) (paragraphStyle: ParagraphStyle) (runStyle: TextStyle) =
    let styledRun =
        match runStyle.Style |> Option.bind (fun styleId -> Map.tryFind styleId resolvedStyles) |> Option.bind (fun style -> style.Text) with
        | Some namedText -> mergeTextStyle namedText runStyle
        | None -> runStyle
    effectiveTextStyle paragraphStyle styledRun

let effectiveTextStyleInDocument (doc: Document) (paragraphStyle: ParagraphStyle) (runStyle: TextStyle) =
    let context = styleContext doc
    effectiveTextStyleWithResolved context.ResolvedCharacterStyles paragraphStyle runStyle

let effectiveTextStyleInContext (context: DocumentStyleContext) (paragraphStyle: ParagraphStyle) (runStyle: TextStyle) =
    effectiveTextStyleWithResolved context.ResolvedCharacterStyles paragraphStyle runStyle

/// Resolves a run style of a paragraph of `doc` to what is displayed:
/// paragraph style, character style and the run's own formatting.
let effectiveRunStyleResolver (doc: Document) : ParagraphBlock -> TextStyle -> TextStyle =
    let context = styleContext doc
    let headingLevels =
        doc.Blocks
        |> List.choose (function
            | Heading (level, para) -> Some (para.Id, level)
            | _ -> None)
        |> Map.ofList
    fun para runStyle ->
        let block =
            match Map.tryFind para.Id headingLevels with
            | Some level -> Heading (level, para)
            | None -> Paragraph para
        effectiveTextStyleWithResolved context.ResolvedCharacterStyles (forBlockInContext context block) runStyle
/// Imported runs may cache the effective formatting supplied by their old named
/// styles. Drop only matching, non-direct values before changing the style - the
/// fonts, the colour and the on/off properties (bold, italic ...): under the old
/// style the text looks the same, and the new style's values can apply. Unrelated
/// values (including table formatting) and explicit overrides stay. (A colour has
/// no explicit flag: one equal to the old style's is taken as the style's.)
let private releaseInheritedFonts (inherited: TextStyle) (style: TextStyle) =
    let family =
        if not (ExplicitFormat.has ExplicitFormat.FontFamily style) && style.FontFamily = inherited.FontFamily then None
        else style.FontFamily
    let size =
        if not (ExplicitFormat.has ExplicitFormat.FontSize style) && abs (style.FontSize - inherited.FontSize) < 0.01 then TextBuffer.defaultStyle.FontSize
        else style.FontSize
    let color =
        if not (ExplicitFormat.has ExplicitFormat.ColorAuto style) && style.Color.IsSome && style.Color = inherited.Color then None
        else style.Color
    // (An on/off property on because the old style had it on, and not set directly.)
    let release flag (value: bool) (inheritedValue: bool) = value && inheritedValue && not (ExplicitFormat.has flag style)
    let bold = release ExplicitFormat.Bold style.Bold inherited.Bold
    let italic = release ExplicitFormat.Italic style.Italic inherited.Italic
    let underline = release ExplicitFormat.Underline style.Underline inherited.Underline
    let allCaps = release ExplicitFormat.AllCaps style.AllCaps inherited.AllCaps
    let smallCaps = release ExplicitFormat.SmallCaps style.SmallCaps inherited.SmallCaps
    if family = style.FontFamily && size = style.FontSize && color = style.Color && not (bold || italic || underline || allCaps || smallCaps) then style
    else
        { style with
            FontFamily = family; FontSize = size; Color = color
            Bold = style.Bold && not bold; Italic = style.Italic && not italic; Underline = style.Underline && not underline
            AllCaps = style.AllCaps && not allCaps; SmallCaps = style.SmallCaps && not smallCaps }

let private inheritedFontsFor (context: DocumentStyleContext) (doc: Document) (para: ParagraphBlock) =
    let block = findBlock doc para.Id |> Option.defaultValue (Paragraph para)
    let paragraph = forBlockInContext context block
    fun (style: TextStyle) ->
        effectiveTextStyleInContext context paragraph { TextBuffer.defaultStyle with Style = style.Style }

/// Apply a named paragraph style, retaining direct run formatting.
let setParagraphStyle styleId (doc: Document) =
    let normalized = if System.String.IsNullOrWhiteSpace styleId then None else Some (styleId.Trim())
    let context = styleContext doc
    let updated =
        (doc, Document.paragraphFormattingTargets doc)
        ||> List.fold (fun current para ->
            if para.Metadata.Style = normalized then current
            else
                let inherited = inheritedFontsFor context doc para
                let buffer = TextBuffer.applyStyle 0 (TextBuffer.length para.Buffer) (fun style -> releaseInheritedFonts (inherited style) style) para.Buffer
                if obj.ReferenceEquals(buffer, para.Buffer) then current
                else Document.replaceParagraph current para.Id { para with Buffer = buffer })
    updated |> Document.clearHeading |> Document.setParagraphStyle styleId

/// `para` with another paragraph style (None: the default one), as setParagraphStyle gives a
/// paragraph: the fonts its runs had from its old style (`context`, the styles before the change)
/// are let go, its direct formatting stays. Its block (a Heading or not) is the caller's.
let restyleParagraph (context: DocumentStyleContext) (doc: Document) (styleId: string option) (para: ParagraphBlock) : ParagraphBlock =
    if para.Metadata.Style = styleId then para
    else
        let inherited = inheritedFontsFor context doc para
        let buffer = TextBuffer.applyStyle 0 (TextBuffer.length para.Buffer) (fun style -> releaseInheritedFonts (inherited style) style) para.Buffer
        { para with Buffer = buffer; Metadata = { para.Metadata with Style = styleId } }

/// Apply a named character style to a selected range (also used by the
/// pending-typing probe for a caret).
let setCharacterStyle styleId (doc: Document) =
    let normalized = if System.String.IsNullOrWhiteSpace styleId then None else Some (styleId.Trim())
    let context = styleContext doc
    Document.applyFormattingToSelectionInParagraph (fun para start length buffer ->
        let inherited = inheritedFontsFor context doc para
        TextBuffer.applyStyle start length (fun style ->
            if style.Style = normalized then style
            else
                let released = releaseInheritedFonts (inherited style) style
                { released with Style = normalized }) buffer) doc


/// Toggle commands decided on the effective formatting of the selection, so
/// toggling text that is bold through its style produces explicit non-bold.
let toggleBold (doc: Document) = Document.toggleBoldWith (effectiveRunStyleResolver doc) doc
let toggleItalic (doc: Document) = Document.toggleItalicWith (effectiveRunStyleResolver doc) doc
let toggleUnderline (doc: Document) = Document.toggleUnderlineWith (effectiveRunStyleResolver doc) doc
let toggleAllCaps (doc: Document) = Document.toggleAllCapsWith (effectiveRunStyleResolver doc) doc
let toggleSmallCaps (doc: Document) = Document.toggleSmallCapsWith (effectiveRunStyleResolver doc) doc
let toggleStrike (doc: Document) = Document.toggleStrikeWith (effectiveRunStyleResolver doc) doc
/// Grow Font / Shrink Font (by Word's size steps, or by a point when not `stepwise`) from the size the text shows.
let growFont (grow: bool) (stepwise: bool) (doc: Document) = Document.resizeFontWith (effectiveRunStyleResolver doc) grow stepwise doc
let toggleSubscript (doc: Document) = Document.toggleVertAlignWith (effectiveRunStyleResolver doc) "subscript" doc
let toggleSuperscript (doc: Document) = Document.toggleVertAlignWith (effectiveRunStyleResolver doc) "superscript" doc

let effectiveBufferWith transformStyle (buffer: TextBuffer) =
    match buffer.Runs with
    | [] -> buffer
    | [ Text textRun ] ->
        let style = transformStyle textRun.Style
        if style = textRun.Style then
            buffer
        else
            { buffer with Runs = [ Text { textRun with Style = style } ] }
    | [ Image _ ] -> buffer
    | _ ->
        let mutable runs = Unchecked.defaultof<ResizeArray<Run>>
        let mutable changed = false
        let copyOriginalPrefix count =
            runs <- ResizeArray<Run>()
            let mutable index = 0
            let mutable remaining = buffer.Runs
            while index < count do
                match remaining with
                | original :: tail ->
                    runs.Add original
                    remaining <- tail
                    index <- index + 1
                | [] -> index <- count
            changed <- true
        let ensureChanged prefixCount =
            if not changed then
                copyOriginalPrefix prefixCount
        let mutable hasPendingText = false
        let mutable pendingIndex = 0
        let mutable pendingStart = 0
        let mutable pendingLength = 0
        let mutable pendingStyle = TextBuffer.defaultStyle
        let flushPendingText () =
            if changed && hasPendingText then
                runs.Add(Text { Start = pendingStart; Length = pendingLength; Style = pendingStyle })
            hasPendingText <- false
        let startPendingText index (run: TextRun) style =
            pendingIndex <- index
            pendingStart <- run.Start
            pendingLength <- run.Length
            pendingStyle <- style
            hasPendingText <- true
        let addText index (run: TextRun) style =
            if hasPendingText && pendingStart + pendingLength = run.Start && pendingStyle.Checkbox.IsNone && (obj.ReferenceEquals(pendingStyle, style) || pendingStyle = style) then
                ensureChanged pendingIndex
                pendingLength <- pendingLength + run.Length
                changed <- true
            else
                if changed then
                    flushPendingText ()
                    startPendingText index run style
                elif style <> run.Style then
                    ensureChanged index
                    hasPendingText <- false
                    startPendingText index run style
                else
                    startPendingText index run style
        let mutable index = 0
        for run in buffer.Runs do
            match run with
            | Text textRun ->
                let style = transformStyle textRun.Style
                addText index textRun style
            | Image image ->
                if changed then
                    flushPendingText ()
                    runs.Add(Image image)
                else
                    hasPendingText <- false
            index <- index + 1
        if changed then
            flushPendingText ()
            { buffer with Runs = resizeArrayToList runs }
        else
            buffer

let effectiveBuffer (paragraphStyle: ParagraphStyle) (buffer: TextBuffer) =
    effectiveBufferWith (effectiveTextStyle paragraphStyle) buffer

let effectiveBufferInDocument (doc: Document) (paragraphStyle: ParagraphStyle) (buffer: TextBuffer) =
    let context = styleContext doc
    effectiveBufferWith (effectiveTextStyleWithResolved context.ResolvedCharacterStyles paragraphStyle) buffer

let effectiveBufferInContext (context: DocumentStyleContext) (paragraphStyle: ParagraphStyle) (buffer: TextBuffer) =
    effectiveBufferWith (effectiveTextStyleWithResolved context.ResolvedCharacterStyles paragraphStyle) buffer

// DocumentProcessors.twipsToPoints stores SpacingBefore/SpacingAfter as
// points (twips / 20). Layout runs in CSS-pixels at 96 DPI (margins came
// out of DocxReader via /15 i.e. twips → px directly), so the spacing
// values get applied at 75% of their intended size unless we convert
// pt → px here at the layout boundary.
let private ptToLayoutPx (pt: float) = pt * 96.0 / 72.0

let spacingBefore (paragraphStyle: ParagraphStyle) (metadata: BlockMetadata) =
    metadata.SpacingBefore
    |> Option.orElse paragraphStyle.SpacingBefore
    |> Option.map ptToLayoutPx
    |> Option.defaultValue 0.0

let spacingAfter (paragraphStyle: ParagraphStyle) (metadata: BlockMetadata) =
    metadata.SpacingAfter
    |> Option.orElse paragraphStyle.SpacingAfter
    |> Option.map ptToLayoutPx
    |> Option.defaultValue 10.7
let lineSpacing (paragraphStyle: ParagraphStyle) (metadata: BlockMetadata) = metadata.LineSpacing |> Option.orElse paragraphStyle.LineSpacing |> Option.defaultValue 1.0
let lineSpacingRule (paragraphStyle: ParagraphStyle) (metadata: BlockMetadata) = metadata.LineSpacingRule |> Option.orElse paragraphStyle.LineSpacingRule |> Option.defaultValue "auto"
/// Keep with next / keep lines together: the paragraph's own setting, else its style's.
let keepNext (paragraphStyle: ParagraphStyle) (metadata: BlockMetadata) = metadata.KeepNext |> Option.orElse paragraphStyle.KeepNext |> Option.defaultValue false
let keepLines (paragraphStyle: ParagraphStyle) (metadata: BlockMetadata) = metadata.KeepLines |> Option.orElse paragraphStyle.KeepLines |> Option.defaultValue false
let widowControl (paragraphStyle: ParagraphStyle) (metadata: BlockMetadata) = metadata.WidowControl |> Option.orElse paragraphStyle.WidowControl |> Option.defaultValue true

let contextualSpacingBetween (paragraphStyle: ParagraphStyle) (current: ParagraphBlock) (next: ParagraphBlock option) =
    let enabled = current.Metadata.ContextualSpacing |> Option.orElse paragraphStyle.ContextualSpacing |> Option.defaultValue false
    enabled && next |> Option.exists (fun value -> current.Metadata.Style.IsSome && current.Metadata.Style = value.Metadata.Style)

let effectiveMetadata (paragraphStyle: ParagraphStyle) (metadata: BlockMetadata) =
    // Left is also the unset value; ExplicitAlignment tells direct left (which
    // must beat a centred style) from "inherit".
    let alignment =
        if metadata.ExplicitAlignment || metadata.Alignment <> Left then metadata.Alignment
        else paragraphStyle.Alignment |> Option.defaultValue metadata.Alignment
    // Nothing to take from the style: the metadata is its own effective form
    // (layout asks for this per paragraph, so no copy is made then).
    let inline unchanged (own: 'a option) (styled: 'a option) = own.IsSome || styled.IsNone
    if alignment = metadata.Alignment
       && unchanged metadata.IndentLeft paragraphStyle.IndentLeft
       && unchanged metadata.IndentRight paragraphStyle.IndentRight
       && unchanged metadata.FirstLineIndent paragraphStyle.FirstLineIndent
       && unchanged metadata.ContextualSpacing paragraphStyle.ContextualSpacing
       && unchanged metadata.LineSpacingRule paragraphStyle.LineSpacingRule
       && unchanged metadata.WidowControl paragraphStyle.WidowControl
       && unchanged metadata.BackgroundColor paragraphStyle.BackgroundColor then
        metadata
    else
        { metadata with
            Alignment = alignment
            IndentLeft = metadata.IndentLeft |> Option.orElse paragraphStyle.IndentLeft
            IndentRight = metadata.IndentRight |> Option.orElse paragraphStyle.IndentRight
            FirstLineIndent = metadata.FirstLineIndent |> Option.orElse paragraphStyle.FirstLineIndent
            ContextualSpacing = metadata.ContextualSpacing |> Option.orElse paragraphStyle.ContextualSpacing
            LineSpacingRule = metadata.LineSpacingRule |> Option.orElse paragraphStyle.LineSpacingRule
            WidowControl = metadata.WidowControl |> Option.orElse paragraphStyle.WidowControl
            BackgroundColor = metadata.BackgroundColor |> Option.orElse paragraphStyle.BackgroundColor }

/// A property over a selection: nothing to report (no text selected), the
/// same value everywhere, or different values.
type FormatValue<'T> =
    | NoValue
    | Uniform of 'T
    | Mixed

module FormatValue =
    let inline add (value: 'T) (current: FormatValue<'T>) =
        match current with
        | NoValue -> Uniform value
        | Uniform existing when existing = value -> current
        | _ -> Mixed

/// Whether a style id names an internal entry of the style list rather than
/// a style of the document ("__docDefaults": the document defaults).
let isInternalStyleId (styleId: string) =
    not (isNull styleId) && styleId.StartsWith "__"

/// The formatting a toolbar shows for the selection: what is displayed
/// (styles resolved), not just what the runs state themselves.
type SelectionFormat = {
    Bold: FormatValue<bool>
    Italic: FormatValue<bool>
    Underline: FormatValue<bool>
    UnderlineMode: FormatValue<string>
    Strike: FormatValue<bool>
    Hidden: FormatValue<bool>
    AllCaps: FormatValue<bool>
    SmallCaps: FormatValue<bool>
    FontFamily: FormatValue<string option>
    FontSize: FormatValue<float>
    Color: FormatValue<string option>
    Highlight: FormatValue<string option>
    VertAlign: FormatValue<string option>
    CharacterStyleId: FormatValue<string option>
    StyleId: FormatValue<string option>
    HeadingLevel: FormatValue<int>
    Alignment: FormatValue<Alignment>
    ListKind: FormatValue<ListStyle>
    ListLevel: FormatValue<int>
    IndentLeft: FormatValue<float>
    IndentRight: FormatValue<float>
    FirstLineIndent: FormatValue<float>
    BackgroundColor: FormatValue<string option>
    NumberFormat: FormatValue<string option>
    ParagraphLineSpacing: FormatValue<float option>
    ParagraphLineSpacingRule: FormatValue<string option>
    ParagraphSpacingBefore: FormatValue<float>
    ParagraphSpacingAfter: FormatValue<float>
}

/// Effective character and paragraph formatting of the selection. With a
/// caret, the formatting typing would continue with (the character before
/// the caret, or the paragraph's own when it is empty) and the caret's
/// paragraph; with a range, every selected run and every paragraph touched.
///
/// (A selection may be the whole of a long document - Ctrl+A - and a host
/// asks for its format with every change of the selection: each property is
/// followed on its own until it is "mixed", nothing is resolved for a
/// paragraph or a run once every property it could change is mixed, a run in
/// the style of the run before it adds nothing, and the walk ends when
/// everything is mixed.)
let selectionFormat (doc: Document) : SelectionFormat =
    let mutable bold = NoValue
    let mutable italic = NoValue
    let mutable underline = NoValue
    let mutable underlineMode = NoValue
    let mutable strike = NoValue
    let mutable hidden = NoValue
    let mutable allCaps = NoValue
    let mutable smallCaps = NoValue
    let mutable fontFamily = NoValue
    let mutable fontSize = NoValue
    let mutable color = NoValue
    let mutable highlight = NoValue
    let mutable vertAlign = NoValue
    let mutable characterStyleId = NoValue
    let mutable styleId = NoValue
    let mutable headingLevel = NoValue
    let mutable alignment = NoValue
    let mutable listKind = NoValue
    let mutable listLevel = NoValue
    let mutable indentLeft = NoValue
    let mutable indentRight = NoValue
    let mutable firstLineIndent = NoValue
    let mutable backgroundColor = NoValue
    let mutable numberFormat = NoValue
    let mutable lineSpacing = NoValue
    let mutable lineSpacingRule = NoValue
    let mutable spacingBefore = NoValue
    let mutable spacingAfter = NoValue
    let inline isMixed (value: FormatValue<'T>) = (match value with Mixed -> true | _ -> false)
    let charactersMixed () =
        isMixed bold && isMixed italic && isMixed underline && isMixed underlineMode && isMixed strike
        && isMixed allCaps && isMixed smallCaps && isMixed hidden
        && isMixed fontFamily && isMixed fontSize && isMixed color && isMixed highlight && isMixed vertAlign && isMixed characterStyleId
    let paragraphsMixed () =
        isMixed styleId && isMixed headingLevel && isMixed alignment && isMixed listKind && isMixed listLevel
        && isMixed indentLeft && isMixed indentRight && isMixed firstLineIndent && isMixed backgroundColor && isMixed numberFormat
        && isMixed lineSpacing && isMixed lineSpacingRule && isMixed spacingBefore && isMixed spacingAfter
    match doc.Selection with
    | None -> ()
    | Some selection ->
        let context = styleContext doc
        // (Asked of the paragraphs of the selection, each found by its place -
        // or, once a selection turns out to be of many paragraphs, of the
        // document's headings found in one walk: not a search of the document
        // per paragraph.)
        let mutable headings : System.Collections.Generic.Dictionary<string, int> = null
        let mutable visited = 0
        let headingLevelOf (id: string) =
            if visited > 16 then
                if isNull headings then
                    headings <- System.Collections.Generic.Dictionary<string, int>()
                    for block in doc.Blocks do
                        match block with
                        | Heading (level, para) -> headings.[para.Id] <- level
                        | _ -> ()
                match Portable.Dict.tryFind id headings with
                | ValueSome level -> Some level
                | ValueNone -> None
            else
                match findBlock doc id with
                | Some (Heading (level, _)) -> Some level
                | _ -> None
        // (A style's paragraph format and whether it is one a user sees, looked up once per style.)
        let styleFormats = System.Collections.Generic.Dictionary<string, ParagraphStyle>()
        let visibleStyles = System.Collections.Generic.Dictionary<string, bool>()
        let paragraphStyleFor (level: int option) (para: ParagraphBlock) =
            match level with
            | Some level -> forBlockInContext context (Heading (level, para))
            | None ->
                match para.Metadata.Style with
                | None -> context.DefaultParagraphStyle
                | Some styleId ->
                    match Portable.Dict.tryFind styleId styleFormats with
                    | ValueSome known -> known
                    | ValueNone ->
                        let found = forBlockInContext context (Paragraph para)
                        styleFormats.[styleId] <- found
                        found
        let isVisibleStyle (styleId: string) =
            match Portable.Dict.tryFind styleId visibleStyles with
            | ValueSome known -> known
            | ValueNone ->
                let visible =
                    not (isInternalStyleId styleId)
                    && (match Map.tryFind ("paragraph", styleId) context.Styles with
                        | Some style -> style.StyleType <> "defaults"
                        | None -> true)
                visibleStyles.[styleId] <- visible
                visible
        let paragraphStyleOf (para: ParagraphBlock) = paragraphStyleFor (headingLevelOf para.Id) para
        // (The same object, or - in JavaScript, where an optional string is the string - the same text, at once.)
        let inline sameText (left: string option) (right: string option) = obj.ReferenceEquals(left, right) || left = right
        let mutable lastParagraphStyle = Unchecked.defaultof<ParagraphStyle>
        let mutable lastRunStyle = Unchecked.defaultof<TextStyle>
        let addCharacter (paragraphStyle: ParagraphStyle) (runStyle: TextStyle) =
            // (What is reported comes from these properties of the run alone.)
            let asLast =
                obj.ReferenceEquals(paragraphStyle, lastParagraphStyle)
                && not (isNull (box lastRunStyle))
                && (obj.ReferenceEquals(runStyle, lastRunStyle)
                    || (runStyle.Explicit = lastRunStyle.Explicit
                        && runStyle.Bold = lastRunStyle.Bold && runStyle.Italic = lastRunStyle.Italic
                        && runStyle.FontSize = lastRunStyle.FontSize
                        && runStyle.Underline = lastRunStyle.Underline && runStyle.Strike = lastRunStyle.Strike
                        && runStyle.DoubleStrike = lastRunStyle.DoubleStrike && runStyle.Hidden = lastRunStyle.Hidden
                        && sameText runStyle.UnderlineStyle lastRunStyle.UnderlineStyle
                        && runStyle.AllCaps = lastRunStyle.AllCaps && runStyle.SmallCaps = lastRunStyle.SmallCaps
                        && sameText runStyle.FontFamily lastRunStyle.FontFamily && sameText runStyle.Color lastRunStyle.Color
                        && sameText runStyle.Highlight lastRunStyle.Highlight && sameText runStyle.Style lastRunStyle.Style
                        && sameText runStyle.VertAlign lastRunStyle.VertAlign))
            if not asLast then
                if not (isMixed characterStyleId) then characterStyleId <- FormatValue.add (runStyle.Style |> Option.filter (isInternalStyleId >> not)) characterStyleId
                lastRunStyle <- runStyle
                lastParagraphStyle <- paragraphStyle
                // (The run's own style when a character style has a part in it: what
                // effectiveTextStyleWithResolved resolves. Each property is then worked out as
                // effectiveTextStyle does, and only while it is not known to be mixed - no
                // style record is made per run.)
                let own =
                    match runStyle.Style with
                    | None -> runStyle
                    | Some styleId ->
                        match Map.tryFind styleId context.ResolvedCharacterStyles |> Option.bind (fun style -> style.Text) with
                        | Some namedText -> mergeTextStyle namedText runStyle
                        | None -> runStyle
                let inherited = paragraphStyle.Text
                if not (isMixed fontFamily) then fontFamily <- FormatValue.add (own.FontFamily |> Option.orElse inherited.FontFamily) fontFamily
                if not (isMixed bold) then bold <- FormatValue.add (inheritToggle ExplicitFormat.Bold inherited.Bold own.Bold own) bold
                if not (isMixed italic) then italic <- FormatValue.add (inheritToggle ExplicitFormat.Italic inherited.Italic own.Italic own) italic
                if not (isMixed underline) then underline <- FormatValue.add (inheritToggle ExplicitFormat.Underline inherited.Underline own.Underline own || TextBuffer.underlineShown own) underline
                if not (isMixed underlineMode) then
                    let shown = inheritToggle ExplicitFormat.Underline inherited.Underline own.Underline own || TextBuffer.underlineShown own
                    underlineMode <- FormatValue.add (if shown then own.UnderlineStyle |> Option.orElse inherited.UnderlineStyle |> Option.defaultValue "single" else "none") underlineMode
                if not (isMixed strike) then
                    strike <-
                        FormatValue.add
                            (inheritToggle ExplicitFormat.Strike inherited.Strike own.Strike own
                             || inheritToggle ExplicitFormat.DoubleStrike inherited.DoubleStrike own.DoubleStrike own)
                            strike
                if not (isMixed hidden) then hidden <- FormatValue.add (inheritToggle ExplicitFormat.Hidden inherited.Hidden own.Hidden own) hidden
                if not (isMixed allCaps) then allCaps <- FormatValue.add (inheritToggle ExplicitFormat.AllCaps inherited.AllCaps own.AllCaps own) allCaps
                if not (isMixed smallCaps) then smallCaps <- FormatValue.add (inheritToggle ExplicitFormat.SmallCaps inherited.SmallCaps own.SmallCaps own) smallCaps
                if not (isMixed fontSize) then fontSize <- FormatValue.add (inheritFontSize inherited.FontSize own) fontSize
                if not (isMixed color) then
                    color <- FormatValue.add (if ExplicitFormat.has ExplicitFormat.ColorAuto own then own.Color else own.Color |> Option.orElse inherited.Color) color
                if not (isMixed highlight) then highlight <- FormatValue.add (own.Highlight |> Option.orElse inherited.Highlight) highlight
                if not (isMixed vertAlign) then vertAlign <- FormatValue.add (own.VertAlign |> Option.orElse inherited.VertAlign) vertAlign
        let addParagraph (level: int option) (para: ParagraphBlock) (paragraphStyle: ParagraphStyle) =
            let own = para.Metadata
            // (What effectiveMetadata gives for the properties reported here.)
            if not (isMixed alignment) then
                let effective =
                    if own.ExplicitAlignment || own.Alignment <> Left then own.Alignment
                    else paragraphStyle.Alignment |> Option.defaultValue own.Alignment
                alignment <- FormatValue.add effective alignment
            if not (isMixed headingLevel) then
                let value =
                    match level with
                    | Some level -> level
                    | None ->
                        match paragraphStyle.OutlineLevel with
                        | Some level when level >= 1 && level <= 9 -> level
                        | _ -> 0
                headingLevel <- FormatValue.add value headingLevel
            if not (isMixed styleId) then
                // The style a user sees: an internal entry (the document defaults,
                // "__docDefaults", which table-cell paragraphs are imported with)
                // is no style - the paragraph has the default one.
                let visibleStyle = own.Style |> Option.filter isVisibleStyle
                styleId <- FormatValue.add visibleStyle styleId
            if not (isMixed lineSpacing) then lineSpacing <- FormatValue.add (own.LineSpacing |> Option.orElse paragraphStyle.LineSpacing) lineSpacing
            if not (isMixed lineSpacingRule) then lineSpacingRule <- FormatValue.add (own.LineSpacingRule |> Option.orElse paragraphStyle.LineSpacingRule) lineSpacingRule
            if not (isMixed spacingBefore) then spacingBefore <- FormatValue.add (own.SpacingBefore |> Option.orElse paragraphStyle.SpacingBefore |> Option.defaultValue 0.0) spacingBefore
            if not (isMixed spacingAfter) then spacingAfter <- FormatValue.add (own.SpacingAfter |> Option.orElse paragraphStyle.SpacingAfter |> Option.defaultValue 0.0) spacingAfter
            if not (isMixed listKind) then listKind <- FormatValue.add own.ListStyle listKind
            if not (isMixed listLevel) then listLevel <- FormatValue.add (if own.ListStyle = NoList then 0 else own.ListLevel) listLevel
            if not (isMixed indentLeft) then indentLeft <- FormatValue.add (own.IndentLeft |> Option.orElse paragraphStyle.IndentLeft |> Option.defaultValue 0.0) indentLeft
            if not (isMixed indentRight) then indentRight <- FormatValue.add (own.IndentRight |> Option.orElse paragraphStyle.IndentRight |> Option.defaultValue 0.0) indentRight
            if not (isMixed firstLineIndent) then firstLineIndent <- FormatValue.add (own.FirstLineIndent |> Option.orElse paragraphStyle.FirstLineIndent |> Option.defaultValue 0.0) firstLineIndent
            if not (isMixed backgroundColor) then backgroundColor <- FormatValue.add (own.BackgroundColor |> Option.orElse paragraphStyle.BackgroundColor) backgroundColor
            if not (isMixed numberFormat) then numberFormat <- FormatValue.add (if own.ListStyle = Numbered then own.ListNumberFormat else None) numberFormat
        let caretStyle (para: ParagraphBlock) offset =
            // What typing here continues with: the character before the caret.
            TextBuffer.getStyleAt (if offset > 0 then offset - 1 else 0) para.Buffer
        let addCaret (blockId: BlockId) offset =
            match findParagraph doc blockId with
            | None -> ()
            | Some para ->
                let level = headingLevelOf para.Id
                let paragraphStyle = paragraphStyleFor level para
                addParagraph level para paragraphStyle
                addCharacter paragraphStyle (caretStyle para offset)
        if selection.IsCollapsed then
            addCaret selection.Focus.BlockId selection.Focus.Offset
        else
            let mutable charactersDone = false
            let mutable paragraphsDone = false
            let mutable first : (ParagraphBlock * int) option = None
            forEachSelectedParagraphRange doc (fun para rangeStart rangeEnd ->
                visited <- visited + 1
                if first.IsNone then first <- Some (para, rangeStart)
                let level = headingLevelOf para.Id
                let paragraphStyle = paragraphStyleFor level para
                if not paragraphsDone then
                    addParagraph level para paragraphStyle
                    paragraphsDone <- paragraphsMixed ()
                if not charactersDone then
                    for run in para.Buffer.Runs do
                        match run with
                        | Text textRun when textRun.Length > 0 && textRun.Start < rangeEnd && textRun.Start + textRun.Length > rangeStart ->
                            addCharacter paragraphStyle textRun.Style
                        | _ -> ()
                    charactersDone <- charactersMixed ()
                not (charactersDone && paragraphsDone))
            |> ignore
            // A range with no text in it (only a picture, or only paragraph marks).
            if bold = NoValue then
                match first with
                | Some (para, rangeStart) -> addCharacter (paragraphStyleOf para) (caretStyle para rangeStart)
                | None -> ()
            if styleId = NoValue then addCaret selection.Focus.BlockId selection.Focus.Offset
    { SelectionFormat.Bold = bold; Italic = italic; Underline = underline; UnderlineMode = underlineMode; Strike = strike
      Hidden = hidden; AllCaps = allCaps; SmallCaps = smallCaps
      FontFamily = fontFamily; FontSize = fontSize; Color = color; Highlight = highlight; VertAlign = vertAlign
      CharacterStyleId = characterStyleId; StyleId = styleId; HeadingLevel = headingLevel; Alignment = alignment
      ListKind = listKind; ListLevel = listLevel; IndentLeft = indentLeft; IndentRight = indentRight; FirstLineIndent = firstLineIndent
      BackgroundColor = backgroundColor; NumberFormat = numberFormat
      ParagraphLineSpacing = lineSpacing; ParagraphLineSpacingRule = lineSpacingRule; ParagraphSpacingBefore = spacingBefore; ParagraphSpacingAfter = spacingAfter }
