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

let styleContext (doc: Document) =
    buildStyleContext doc

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
        | Some styleId -> context.ParagraphStyles |> Map.tryFind styleId |> Option.defaultValue context.NormalStyle
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

/// Toggle commands decided on the effective formatting of the selection, so
/// toggling text that is bold through its style produces explicit non-bold.
let toggleBold (doc: Document) = Document.toggleBoldWith (effectiveRunStyleResolver doc) doc
let toggleItalic (doc: Document) = Document.toggleItalicWith (effectiveRunStyleResolver doc) doc
let toggleUnderline (doc: Document) = Document.toggleUnderlineWith (effectiveRunStyleResolver doc) doc
let toggleAllCaps (doc: Document) = Document.toggleAllCapsWith (effectiveRunStyleResolver doc) doc
let toggleSmallCaps (doc: Document) = Document.toggleSmallCapsWith (effectiveRunStyleResolver doc) doc
let toggleStrike (doc: Document) = Document.toggleStrikeWith (effectiveRunStyleResolver doc) doc

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
            if hasPendingText && pendingStart + pendingLength = run.Start && (obj.ReferenceEquals(pendingStyle, style) || pendingStyle = style) then
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
