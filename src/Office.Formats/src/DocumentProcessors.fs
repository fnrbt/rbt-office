module DocumentProcessors

open DocxReader

// ============================================================================
// TypeConversion Module
// Converts OOXML string representations to typed values
// ============================================================================

module TypeConversion =

    /// Convert Word font size (half-points) to CSS-pixels-at-96-DPI, which is
    /// the unit the rest of the Word renderer pipeline operates in. Margins,
    /// indents, spacing-after, and image positions all reach Layout as
    /// CSS-pixels; font sizes were the last value still flowing through as
    /// points, which caused 75%-of-correct glyph sizes everywhere.
    /// Input: "24" (Word half-points) -> Output: 16.0 (12pt as 16 CSS-px).
    /// Default (no size stated anywhere): Word's 10pt, rendered as 13.333... CSS-px.
    let fontSizeToPoints (size: string option) : float =
        let pt =
            match size with
            | Some s ->
                match System.Double.TryParse(s) with
                | (true, v) -> v / 2.0
                | _ -> 10.0
            | None -> 10.0
        pt * 96.0 / 72.0

    /// Normalize color from Word format to CSS format
    /// Input: "FF0000" or "auto" -> Output: Some "#FF0000" or None
    /// Handles 6-char (RRGGBB) and 8-char (AARRGGBB) formats
    let normalizeColor (color: string option) : string option =
        match color with
        | Some c when c = "auto" || c = "" -> None
        | Some c when c.Length = 6 ->
            Some (sprintf "#%s" (c.ToUpper()))
        | Some c when c.Length = 8 ->
            // AARRGGBB format, strip alpha channel
            Some (sprintf "#%s" (c.Substring(2).ToUpper()))
        | Some c ->
            // Try to parse as is
            if c.StartsWith("#") then Some c else Some (sprintf "#%s" c)
        | None -> None

    /// Normalize alignment from Word values to web standards
    /// Maps "both" and "distribute" to "justify"
    /// Output: "left" | "center" | "right" | "justify"
    let normalizeAlignment (alignment: string option) : string =
        match alignment with
        | Some "center" -> "center"
        | Some "right" -> "right"
        | Some "justify" -> "justify"
        | Some "both" -> "justify"
        | Some "distribute" -> "justify"
        | _ -> "left"

    /// Normalize font family, handling theme fonts
    /// Theme fonts like "+majorFont" map to default fonts
    /// Default: "Calibri"
    let normalizeFontFamily (family: string option) : string =
        match family with
        | Some f when f.StartsWith("+") -> "Calibri"  // Theme fonts
        | Some f when f <> "" -> f
        | _ -> "Calibri"

    let resolveThemeFont (theme: Theme option) (themeFont: string option) =
        match theme, themeFont with
        | Some theme, Some value when value.StartsWith("major") -> theme.MajorFont
        | Some theme, Some value when value.StartsWith("minor") -> theme.MinorFont
        | _ -> None

    let normalizeFontFamilyWithTheme (theme: Theme option) (family: string option) (themeFont: string option) : string =
        match family with
        | Some f when f <> "" && not (f.StartsWith("+")) -> f
        | _ -> resolveThemeFont theme themeFont |> Option.defaultValue (normalizeFontFamily family)

    let private hexByte (text: string) =
        let nibble ch =
            if ch >= '0' && ch <= '9' then Some (int ch - int '0')
            elif ch >= 'A' && ch <= 'F' then Some (10 + int ch - int 'A')
            elif ch >= 'a' && ch <= 'f' then Some (10 + int ch - int 'a')
            else None
        if text.Length <> 2 then None
        else
            match nibble text.[0], nibble text.[1] with
            | Some high, Some low -> Some (high * 16 + low)
            | _ -> None

    let private clampByte value = max 0 (min 255 value)

    let private applyThemeShade shade channel =
        int (float channel * (float shade / 255.0)) |> clampByte

    let private applyThemeTint tint channel =
        int (float channel + (float (255 - channel) * (float tint / 255.0))) |> clampByte

    let private transformThemeColor (tint: string option) (shade: string option) (color: string) =
        if color.Length <> 7 || not (color.StartsWith("#")) then color
        else
            let transformed =
                [ color.Substring(1, 2); color.Substring(3, 2); color.Substring(5, 2) ]
                |> List.map (fun part ->
                    let channel = hexByte part |> Option.defaultValue 0
                    let shaded = shade |> Option.bind hexByte |> Option.map (fun value -> applyThemeShade value channel) |> Option.defaultValue channel
                    tint |> Option.bind hexByte |> Option.map (fun value -> applyThemeTint value shaded) |> Option.defaultValue shaded)
            sprintf "#%02X%02X%02X" transformed.[0] transformed.[1] transformed.[2]

    let normalizeColorWithTheme (theme: Theme option) (color: string option) (themeColor: string option) (themeTint: string option) (themeShade: string option) : string option =
        match normalizeColor color with
        | Some resolved -> Some resolved
        | None ->
            match theme, themeColor with
            | Some theme, Some key -> Map.tryFind key theme.Colors |> Option.map (transformThemeColor themeTint themeShade)
            | _ -> None

    let twipsToPoints (twips: string option) : float option =
        twips
        |> Option.bind (fun value ->
            match System.Double.TryParse(value) with
            | true, parsed -> Some (parsed / 20.0)
            | false, _ -> None)

    let twipsToCanvasPixels (twips: string option) : float option =
        twips
        |> Option.bind (fun value ->
            match System.Double.TryParse(value) with
            | true, parsed -> Some (parsed / 15.0)
            | false, _ -> None)

    let wordAutoSpacingPoints = 14.0

    let lineSpacingToMultiple (line: string option) : float option =
        line
        |> Option.bind (fun value ->
            match System.Double.TryParse(value) with
            | true, parsed -> Some (parsed / 240.0)
            | false, _ -> None)

// ============================================================================
// StyleResolution Module
// Implements CSS-like cascade model for style inheritance
// ============================================================================

module StyleResolution =

    let documentDefaultsStyleId = "__docDefaults"

    /// Fully resolved run properties (no optionals)
    type ResolvedRunProperties = {
        Bold: bool
        Italic: bool
        Underline: bool
        Strike: bool
        FontSize: float
        CharacterSpacing: float
        FontFamily: string
        Color: string option
        VertAlign: string option
    }

    /// Fully resolved paragraph properties (no optionals for key properties)
    type ResolvedParagraphProperties = {
        Alignment: string
        IndentLeft: float option
        IndentRight: float option
        FirstLineIndent: float option
        RightTabPosition: float option
        RightTabLeader: string option
        SpacingBefore: float option
        SpacingAfter: float option
        ContextualSpacing: bool option
        LineSpacing: float option
        ParagraphBorders: ParagraphBorders option
    }

    /// Default run properties (Word document defaults)
    let defaultRunProps : ResolvedRunProperties = {
        Bold = false
        Italic = false
        Underline = false
        Strike = false
        // CSS px like every parsed size (fontSizeToPoints): Word's 10pt. A bare 11.0 here
        // was 11px = 8.25pt for any document that never states w:sz.
        FontSize = TypeConversion.fontSizeToPoints None
        CharacterSpacing = 0.0
        FontFamily = "Calibri"
        Color = None
        VertAlign = None
    }

    /// Default paragraph properties
    let defaultParaProps : ResolvedParagraphProperties = {
        Alignment = "left"
        IndentLeft = None
        IndentRight = None
        FirstLineIndent = None
        RightTabPosition = None
        RightTabLeader = None
        SpacingBefore = None
        SpacingAfter = None
        ContextualSpacing = None
        LineSpacing = None
        ParagraphBorders = None
    }

    /// Merge run properties: base + override
    let mergeRunPropsWithTheme (theme: Theme option) (baseProps: ResolvedRunProperties) (overrideProps: RunProperties option) : ResolvedRunProperties =
        match overrideProps with
        | None -> baseProps
        | Some props ->
            {
                Bold = props.Bold |> Option.defaultValue baseProps.Bold
                Italic = props.Italic |> Option.defaultValue baseProps.Italic
                Underline = props.Underline |> Option.defaultValue baseProps.Underline
                Strike = props.Strike |> Option.defaultValue baseProps.Strike
                FontSize =
                    match props.FontSize with
                    | Some _ -> TypeConversion.fontSizeToPoints props.FontSize
                    | None -> baseProps.FontSize
                CharacterSpacing =
                    match props.CharacterSpacing with
                    | Some _ -> TypeConversion.twipsToCanvasPixels props.CharacterSpacing |> Option.defaultValue baseProps.CharacterSpacing
                    | None -> baseProps.CharacterSpacing
                FontFamily =
                    match props.FontFamily, props.ThemeFont with
                    | Some _, _
                    | _, Some _ -> TypeConversion.normalizeFontFamilyWithTheme theme props.FontFamily props.ThemeFont
                    | None, None -> baseProps.FontFamily
                Color =
                    match props.Color, props.ThemeColor with
                    | Some _, _
                    | _, Some _ -> TypeConversion.normalizeColorWithTheme theme props.Color props.ThemeColor props.ThemeTint props.ThemeShade
                    | None, None -> baseProps.Color
                VertAlign =
                    match props.VertAlign with
                    | Some v -> Some v
                    | None -> baseProps.VertAlign
            }

    let mergeRunProps baseProps overrideProps = mergeRunPropsWithTheme None baseProps overrideProps

    let overlayCharacterStyle
        (styleMap: Map<string, ResolvedRunProperties * ResolvedParagraphProperties>)
        (baseProps: ResolvedRunProperties)
        (styleId: string option)
        : ResolvedRunProperties =

        match styleId |> Option.bind (fun id -> styleMap.TryFind(id) |> Option.map fst) with
        | None -> baseProps
        | Some styleProps ->
            let defaults =
                styleMap.TryFind(documentDefaultsStyleId)
                |> Option.map fst
                |> Option.defaultValue defaultRunProps
            {
                Bold = if styleProps.Bold <> defaults.Bold then styleProps.Bold else baseProps.Bold
                Italic = if styleProps.Italic <> defaults.Italic then styleProps.Italic else baseProps.Italic
                Underline = if styleProps.Underline <> defaults.Underline then styleProps.Underline else baseProps.Underline
                Strike = if styleProps.Strike <> defaults.Strike then styleProps.Strike else baseProps.Strike
                FontSize = if styleProps.FontSize <> defaults.FontSize then styleProps.FontSize else baseProps.FontSize
                CharacterSpacing = if abs (styleProps.CharacterSpacing - defaults.CharacterSpacing) >= 0.001 then styleProps.CharacterSpacing else baseProps.CharacterSpacing
                FontFamily = if styleProps.FontFamily <> defaults.FontFamily then styleProps.FontFamily else baseProps.FontFamily
                Color = if styleProps.Color <> defaults.Color then styleProps.Color else baseProps.Color
                VertAlign = if styleProps.VertAlign <> defaults.VertAlign then styleProps.VertAlign else baseProps.VertAlign
            }

    /// Merge paragraph properties: base + override
    let mergeParaProps (baseProps: ResolvedParagraphProperties) (overrideProps: ParagraphProperties option) : ResolvedParagraphProperties =
        match overrideProps with
        | None -> baseProps
        | Some props ->
            {
                Alignment =
                    match props.Alignment with
                    | Some _ -> TypeConversion.normalizeAlignment props.Alignment
                    | None -> baseProps.Alignment
                IndentLeft =
                    match props.IndentLeft with
                    | Some _ -> TypeConversion.twipsToPoints props.IndentLeft
                    | None -> baseProps.IndentLeft
                IndentRight =
                    match props.IndentRight with
                    | Some _ -> TypeConversion.twipsToPoints props.IndentRight
                    | None -> baseProps.IndentRight
                FirstLineIndent =
                    match TypeConversion.twipsToPoints props.FirstLineIndent, TypeConversion.twipsToPoints props.HangingIndent with
                    | _, Some hanging -> Some -hanging
                    | Some firstLine, None -> Some firstLine
                    | None, None -> baseProps.FirstLineIndent
                RightTabPosition =
                    match props.RightTabPosition with
                    | Some _ -> TypeConversion.twipsToCanvasPixels props.RightTabPosition
                    | None -> baseProps.RightTabPosition
                RightTabLeader = props.RightTabLeader |> Option.orElse baseProps.RightTabLeader
                SpacingBefore =
                    match props.SpacingBeforeAuto, props.SpacingBefore with
                    | Some true, _ -> Some TypeConversion.wordAutoSpacingPoints
                    | _, Some _ -> TypeConversion.twipsToPoints props.SpacingBefore
                    | _ -> baseProps.SpacingBefore
                SpacingAfter =
                    match props.SpacingAfterAuto, props.SpacingAfter with
                    | Some true, _ -> Some TypeConversion.wordAutoSpacingPoints
                    | _, Some _ -> TypeConversion.twipsToPoints props.SpacingAfter
                    | _ -> baseProps.SpacingAfter
                ContextualSpacing = props.ContextualSpacing |> Option.orElse baseProps.ContextualSpacing
                LineSpacing =
                    match props.LineSpacing with
                    | Some _ -> TypeConversion.lineSpacingToMultiple props.LineSpacing
                    | None -> baseProps.LineSpacing
                ParagraphBorders = props.ParagraphBorders |> Option.orElse baseProps.ParagraphBorders
            }

    /// Build a map of style ID -> resolved properties
    /// Handles BasedOn inheritance chains with memoization
    /// Detects circular references and breaks them with defaults
    let buildStyleMapWithTheme (theme: Theme option) (styles: StyleDefinition list) : Map<string, ResolvedRunProperties * ResolvedParagraphProperties> =
        let docDefaultStyle = styles |> List.tryFind (fun style -> style.StyleType = "defaults")
        let documentRunProps =
            match docDefaultStyle with
            | Some style -> mergeRunPropsWithTheme theme defaultRunProps style.RunProperties
            | None -> defaultRunProps
        let documentParaProps =
            match docDefaultStyle with
            | Some style -> mergeParaProps defaultParaProps style.ParagraphProperties
            | None -> defaultParaProps

        let rec resolveStyle (styleId: string) (visited: Set<string>) (styleDict: Map<string, StyleDefinition>) : ResolvedRunProperties * ResolvedParagraphProperties =
            // Check for circular reference
            if visited.Contains(styleId) then
                (documentRunProps, documentParaProps)
            else
                match styleDict.TryFind(styleId) with
                | None -> (documentRunProps, documentParaProps)
                | Some style ->
                    // Get base style properties if BasedOn is specified
                    let (baseRunProps, baseParaProps) =
                        match style.BasedOn with
                        | Some basedOnId when basedOnId <> styleId ->
                            resolveStyle basedOnId (visited.Add(styleId)) styleDict
                        | _ -> (documentRunProps, documentParaProps)

                    // Apply this style's properties on top of base
                    let runProps = mergeRunPropsWithTheme theme baseRunProps style.RunProperties
                    let paraProps = mergeParaProps baseParaProps style.ParagraphProperties
                    (runProps, paraProps)

        let styleDict = styles |> List.map (fun s -> (s.Id, s)) |> Map.ofList

        // Pre-compute all styles
        styles
        |> List.map (fun style ->
            let (runProps, paraProps) = resolveStyle style.Id Set.empty styleDict
            (style.Id, (runProps, paraProps)))
        |> Map.ofList

    let buildStyleMap (styles: StyleDefinition list) : Map<string, ResolvedRunProperties * ResolvedParagraphProperties> =
        buildStyleMapWithTheme None styles

    /// Resolve run properties for a specific run
    /// Applies cascade: document defaults -> paragraph style -> direct formatting
    let resolveRunProps
        (styleMap: Map<string, ResolvedRunProperties * ResolvedParagraphProperties>)
        (paraProps: ParagraphProperties)
        (runProps: RunProperties)
        : ResolvedRunProperties =

        // Start with paragraph style (if specified)
        let baseProps =
            match paraProps.Style with
            | Some styleId ->
                match styleMap.TryFind(styleId) with
                | Some (runProps, _) -> runProps
                | None ->
                    styleMap.TryFind(documentDefaultsStyleId)
                    |> Option.map fst
                    |> Option.defaultValue defaultRunProps
            | None ->
                styleMap.TryFind(documentDefaultsStyleId)
                |> Option.map fst
                |> Option.defaultValue defaultRunProps

        // Apply direct run formatting
        mergeRunProps baseProps (Some runProps)

    /// Resolve paragraph properties
    let resolveParagraphProps
        (styleMap: Map<string, ResolvedRunProperties * ResolvedParagraphProperties>)
        (paraProps: ParagraphProperties)
        : ResolvedParagraphProperties =

        // Get base from style
        let baseProps =
            match paraProps.Style with
            | Some styleId ->
                match styleMap.TryFind(styleId) with
                | Some (_, paraProps) -> paraProps
                | None ->
                    styleMap.TryFind(documentDefaultsStyleId)
                    |> Option.map snd
                    |> Option.defaultValue defaultParaProps
            | None ->
                styleMap.TryFind(documentDefaultsStyleId)
                |> Option.map snd
                |> Option.defaultValue defaultParaProps

        // Apply direct paragraph formatting
        mergeParaProps baseProps (Some paraProps)

// ============================================================================
// EditorProcessor Module
// Transforms parsed documents to editor-friendly format
// ============================================================================

module EditorProcessor =

    /// Editor-friendly run format (fully resolved properties)
    type EditorRun = {
        Text: string
        Bold: bool
        Italic: bool
        Underline: bool
        Strike: bool
        FontSize: float
        FontFamily: string
        Color: string option
        Highlight: string option
        VertAlign: string option
        CharacterSpacing: float
        CommentId: string option
        CommentText: string option
        CommentAuthor: string option
        CommentDate: string option
        NoteId: string option
        NoteType: string option
        NoteText: string option
        Hyperlink: string option
        /// The bookmarks that start before the run, and what ends (or is opened) after it.
        BookmarkStarts: ModelText.BookmarkStart list
        BookmarkEnds: ModelText.BookmarkEnd list
        FieldCode: string option
        FieldResult: string option
        EquationXml: string option
        RenderedPageBreakBefore: bool
        // Image support
        IsImage: bool
        ImageSrc: string option      // Data URL for image
        ImageLinked: bool
        ImageWidth: float option     // Width in pixels
        ImageHeight: float option    // Height in pixels
        ImageAlt: string option      // Alt text
        ImageName: string option
        ImageTitle: string option
        ImageFloating: bool
        ImageWrap: string option
        ImagePositionX: float option
        ImagePositionY: float option
        ImagePositionHRelativeFrom: string option
        ImagePositionVRelativeFrom: string option
        ImageBehindDoc: bool option
        ImageAllowOverlap: bool option
        ImageLayoutInCell: bool option
        ImageDistanceTop: float option
        ImageDistanceBottom: float option
        ImageDistanceLeft: float option
        ImageDistanceRight: float option
        ImageRelativeHeight: string option
        ImageLocked: bool option
        ImageSimplePosition: bool option
        ImageSimplePositionX: float option
        ImageSimplePositionY: float option
        ImageCropLeft: string option
        ImageCropTop: string option
        ImageCropRight: string option
        ImageCropBottom: string option
        /// The picture is mirrored across / upside down (a:xfrm flipH / flipV).
        ImageFlipH: bool
        ImageFlipV: bool
    }

    /// Editor-friendly paragraph format
    type EditorParagraph = {
        Runs: EditorRun list
        Alignment: string
        IndentLeft: float option
        IndentRight: float option
        FirstLineIndent: float option
        RightTabPosition: float option
        RightTabLeader: string option
        SpacingBefore: float option
        SpacingAfter: float option
        LineSpacing: float option
    }

    /// Extract text content from RunContent
    let extractRunText (comments: Map<string, Comment>) (content: RunContent) : string =
        match content with
        | Text text -> text
        | Link link -> link.Text
        | Image img ->
            match img.Description with
            | Some desc -> sprintf "[Image: %s]" desc
            | None -> "[Image]"
        | CommentReference _ ->
            // Word treats a w:commentReference as a sidebar/margin annotation,
            // not as inline text.  When the renderer's "show comments inline"
            // mode (eg. the editor's comment-pop) wants to display the comment
            // body it can look it up via the CommentId on the editor run; the
            // base text extracted here is empty so a print/PDF render doesn't
            // splatter the comment body into the document flow.
            ""
        | NoteRef note -> sprintf "[%s %s]" note.NoteType note.Id
        | BookmarkStart _ -> ""
        | BookmarkEnd _ -> ""
        | FieldRef field ->
            match field.Part |> Option.orElse field.Result with
            | Some result -> result
            | None -> field.Code
        | FieldBoundary _ -> ""
        | EquationXml _ -> "[Equation]"
        | RenderedPageBreak -> ""
        | PageBreak -> ""

    let private noteBodyText (note: Note) =
        note.Content
        |> List.collect (fun para -> para.Runs)
        |> List.map (fun run -> extractRunText Map.empty run.Content)
        |> String.concat ""

    let private noteText (notes: Map<string * string, string>) (note: NoteReference) =
        notes |> Map.tryFind (note.NoteType, note.Id)

    let private toLowerRoman (n: int) =
        let table = [ 1000, "m"; 900, "cm"; 500, "d"; 400, "cd"; 100, "c"; 90, "xc"; 50, "l"; 40, "xl"; 10, "x"; 9, "ix"; 5, "v"; 4, "iv"; 1, "i" ]
        let rec build rem acc =
            match table |> List.tryFind (fun (v, _) -> rem >= v) with
            | Some (v, s) -> build (rem - v) (acc + s)
            | None -> acc
        if n <= 0 then string n else build n ""

    /// The displayed mark for a note reference. Word defaults footnotes to arabic
    /// (the sequential id) and endnotes to lowercase roman numerals.
    let noteMark (noteType: string) (id: string) =
        if noteType = "endnote" then
            match System.Int32.TryParse id with
            | true, n -> toLowerRoman n
            | _ -> id
        else id

    /// Convert a run to EditorRun format (without image data)
    let toEditorRun
        (theme: Theme option)
        (styleMap: Map<string, StyleResolution.ResolvedRunProperties * StyleResolution.ResolvedParagraphProperties>)
        (resolved: StyleResolution.ResolvedRunProperties)
        (comments: Map<string, Comment>)
        (notes: Map<string * string, string>)
        (run: Run)
        : EditorRun option =

        match run.Content with
        | Image img ->
            // Image run - ImageSrc will be populated later if archive is provided
            Some {
                Text = img.RelationshipId  // Store RelationshipId temporarily in Text field
                Bold = false
                Italic = false
                Underline = false
                Strike = false
                FontSize = 11.0
                CharacterSpacing = 0.0
                FontFamily = "Calibri"
                Color = None
                Highlight = None
                VertAlign = None
                CommentId = None
                CommentText = None
                CommentAuthor = None
                CommentDate = None
                NoteId = None
                NoteType = None
                NoteText = None
                Hyperlink = None
                BookmarkStarts = []
                BookmarkEnds = []
                FieldCode = None
                FieldResult = None
                EquationXml = None
                RenderedPageBreakBefore = run.RenderedPageBreakBefore
                IsImage = true
                ImageSrc = None  // Will be populated when we have archive access
                ImageLinked = img.Linked
                ImageWidth = img.Width
                ImageHeight = img.Height
                ImageAlt = img.Description
                ImageName = img.Name
                ImageTitle = img.Title
                ImageFloating = img.Floating
                ImageWrap = img.Wrap
                ImagePositionX = img.PositionX
                ImagePositionY = img.PositionY
                ImagePositionHRelativeFrom = img.PositionHRelativeFrom
                ImagePositionVRelativeFrom = img.PositionVRelativeFrom
                ImageBehindDoc = img.BehindDoc
                ImageAllowOverlap = img.AllowOverlap
                ImageLayoutInCell = img.LayoutInCell
                ImageDistanceTop = img.DistanceTop
                ImageDistanceBottom = img.DistanceBottom
                ImageDistanceLeft = img.DistanceLeft
                ImageDistanceRight = img.DistanceRight
                ImageRelativeHeight = img.RelativeHeight
                ImageLocked = img.Locked
                ImageSimplePosition = img.SimplePosition
                ImageSimplePositionX = img.SimplePositionX
                ImageSimplePositionY = img.SimplePositionY
                ImageCropLeft = img.CropLeft
                ImageCropTop = img.CropTop
                ImageCropRight = img.CropRight
                ImageCropBottom = img.CropBottom
                ImageFlipH = img.FlipH
                ImageFlipV = img.FlipV
            }
        | _ ->
            let text = extractRunText comments run.Content

            // Filter out empty runs
            if text = "" then
                None
            else
                // Merge resolved props with run-specific overrides
                let characterProps = StyleResolution.overlayCharacterStyle styleMap resolved run.Properties.Style
                let finalProps = StyleResolution.mergeRunPropsWithTheme theme characterProps (Some run.Properties)

                let comment =
                    match run.Content with
                    | CommentReference id -> comments |> Map.tryFind id
                    | _ -> None

                let note =
                    match run.Content with
                    | NoteRef note -> Some note
                    | _ -> None

                let field =
                    match run.Content with
                    | FieldRef field -> Some field
                    | _ -> None

                let hyperlink =
                    match run.Content with
                    | Link link -> link.Target
                    | _ -> None

                // A footnote/endnote reference renders as its (superscript) mark
                // number, not the "[footnote N]" placeholder that extractRunText
                // produces. Word writes note ids sequentially, so the id doubles
                // as the display number, and references are always superscript.
                let noteRefText = note |> Option.map (fun value -> noteMark value.NoteType value.Id)
                let displayText = noteRefText |> Option.defaultValue text
                let displayVertAlign =
                    match note with
                    | Some _ -> Some "superscript"
                    | None -> finalProps.VertAlign

                Some {
                    Text = displayText
                    Bold = finalProps.Bold
                    Italic = finalProps.Italic
                    Underline = finalProps.Underline
                    Strike = finalProps.Strike
                    FontSize = finalProps.FontSize
                    CharacterSpacing = finalProps.CharacterSpacing
                    FontFamily = finalProps.FontFamily
                    Color = finalProps.Color
                    Highlight = run.Properties.Highlight
                    VertAlign = displayVertAlign
                    CommentId = match run.Content with CommentReference id -> Some id | _ -> None
                    CommentText = comment |> Option.map (fun value -> value.Text)
                    CommentAuthor = comment |> Option.bind (fun value -> value.Author)
                    CommentDate = comment |> Option.bind (fun value -> value.Date)
                    NoteId = note |> Option.map (fun value -> value.Id)
                    NoteType = note |> Option.map (fun value -> value.NoteType)
                    NoteText = note |> Option.bind (noteText notes)
                    Hyperlink = hyperlink
                    BookmarkStarts = run.BookmarkStarts
                    BookmarkEnds = run.BookmarkEnds
                    FieldCode = field |> Option.map (fun value -> value.Code)
                    FieldResult = field |> Option.bind (fun value -> value.Result)
                    EquationXml = (match run.Content with EquationXml xml -> Some xml | _ -> None)
                    RenderedPageBreakBefore = run.RenderedPageBreakBefore
                    IsImage = false
                    ImageSrc = None
                    ImageLinked = false
                    ImageWidth = None
                    ImageHeight = None
                    ImageAlt = None
                    ImageName = None
                    ImageTitle = None
                    ImageFloating = false
                    ImageWrap = None
                    ImagePositionX = None
                    ImagePositionY = None
                    ImagePositionHRelativeFrom = None
                    ImagePositionVRelativeFrom = None
                    ImageBehindDoc = None
                    ImageAllowOverlap = None
                    ImageLayoutInCell = None
                    ImageDistanceTop = None
                    ImageDistanceBottom = None
                    ImageDistanceLeft = None
                    ImageDistanceRight = None
                    ImageRelativeHeight = None
                    ImageLocked = None
                    ImageSimplePosition = None
                    ImageSimplePositionX = None
                    ImageSimplePositionY = None
                    ImageCropLeft = None
                    ImageCropTop = None
                    ImageCropRight = None
                    ImageCropBottom = None
                    ImageFlipH = false
                    ImageFlipV = false
                }

    /// Convert a paragraph to EditorParagraph format
    let toEditorParagraph
        (theme: Theme option)
        (styleMap: Map<string, StyleResolution.ResolvedRunProperties * StyleResolution.ResolvedParagraphProperties>)
        (comments: Map<string, Comment>)
        (notes: Map<string * string, string>)
        (para: Paragraph)
        : EditorParagraph =

        // Resolve paragraph properties
        let paraProps = StyleResolution.resolveParagraphProps styleMap para.Properties

        // Get base run properties from paragraph style
        let baseRunProps =
            match para.Properties.Style with
            | Some styleId ->
                match styleMap.TryFind(styleId) with
                | Some (runProps, _) -> runProps
                | None -> StyleResolution.defaultRunProps
            | None ->
                styleMap.TryFind(StyleResolution.documentDefaultsStyleId)
                |> Option.map fst
                |> Option.defaultValue StyleResolution.defaultRunProps
        let baseRunProps =
            StyleResolution.mergeRunPropsWithTheme theme baseRunProps para.Properties.ParagraphRunProperties

        // Convert all runs
        let runs =
            para.Runs
            |> List.choose (toEditorRun theme styleMap baseRunProps comments notes)

        {
            Runs = runs
            Alignment = paraProps.Alignment
            IndentLeft = paraProps.IndentLeft
            IndentRight = paraProps.IndentRight
            FirstLineIndent = paraProps.FirstLineIndent
            RightTabPosition = paraProps.RightTabPosition
            RightTabLeader = paraProps.RightTabLeader
            SpacingBefore = paraProps.SpacingBefore
            SpacingAfter = paraProps.SpacingAfter
            LineSpacing = paraProps.LineSpacing
        }
