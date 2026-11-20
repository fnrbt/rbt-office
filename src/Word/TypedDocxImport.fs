module Rbt.Word.TypedDocxImport

open Rbt.Word.Document
open Rbt.Word.TextBuffer

let private origin = Option.map (function WordEditorPackage.ReadAt at -> ReadAt at | WordEditorPackage.LikeAt at -> LikeAt at)
let private alignment (value: string) =
    match value.ToLowerInvariant() with
    | "center" -> Center
    | "right" | "end" -> Right
    | "justify" | "both" | "distribute" | "thaidistribute" -> Justify
    | _ -> Left

let private columns (value: WordEditorPackage.SectionColumns) : SectionColumns =
    { Count = value.Count; Gap = value.Gap; Separator = value.Separator; Widths = value.Widths }

let private references (items: WordEditorPackage.HeaderFooterReference list) =
    items |> List.map (fun item -> item.Kind, item.RelationshipId)

let private sectionSetup (value: WordEditorPackage.PageSetup) : SectionPageSetup =
    { PageWidth = value.PageWidth; PageHeight = value.PageHeight
      MarginLeft = value.MarginLeft; MarginRight = value.MarginRight
      MarginTop = value.MarginTop; MarginBottom = value.MarginBottom
      HeaderDistance = value.HeaderDistance; FooterDistance = value.FooterDistance
      DifferentFirstPage = value.DifferentFirstPage; EvenAndOddHeaders = value.EvenAndOddHeaders }

let private pageSetup (value: WordEditorPackage.PageSetup) : PageSetup =
    { PageWidth = value.PageWidth; PageHeight = value.PageHeight
      MarginLeft = value.MarginLeft; MarginRight = value.MarginRight
      MarginTop = value.MarginTop; MarginBottom = value.MarginBottom
      HeaderDistance = value.HeaderDistance; FooterDistance = value.FooterDistance
      DifferentFirstPage = value.DifferentFirstPage; EvenAndOddHeaders = value.EvenAndOddHeaders
      DefaultTabStop = value.DefaultTabStop |> Option.filter (fun x -> x > 0.0) |> Option.defaultValue defaultPageSetup.DefaultTabStop
      // (A document without the setting - or without a settings part - is older than Word 2010:
      // Word opens it in compatibility mode.)
      CompatibilityMode = value.CompatibilityMode |> Option.defaultValue 12
      UseFarEastLayout = value.UseFarEastLayout |> Option.defaultValue false
      CharacterSpacingControl = value.CharacterSpacingControl
      Hyphenation = value.Hyphenation
      Columns = value.Columns |> Option.filter (fun x -> x.Count >= 2) |> Option.map columns
      SectionStart = value.SectionStart; HeaderReferences = references value.HeaderReferences
      FooterReferences = references value.FooterReferences; PageNumberStart = value.PageNumberStart
      PageNumberFormat = value.PageNumberFormat; VerticalAlign = value.VerticalAlign }

let private paragraphBorder (value: WordEditorPackage.Border) : ParagraphBorder =
    { Style = value.Style; Color = value.Color; Width = value.Width; Space = value.Space }
let private paragraphBorders (value: WordEditorPackage.Borders) : ParagraphBorders =
    { Top = Option.map paragraphBorder value.Top; Right = Option.map paragraphBorder value.Right
      Bottom = Option.map paragraphBorder value.Bottom; Left = Option.map paragraphBorder value.Left
      Between = Option.map paragraphBorder value.InsideH }
let private tableBorder (value: WordEditorPackage.Border) : TableBorder =
    { Style = value.Style; Color = value.Color; Width = value.Width }
let private tableBorders (value: WordEditorPackage.Borders) : TableBorders =
    { Top = Option.map tableBorder value.Top; Right = Option.map tableBorder value.Right
      Bottom = Option.map tableBorder value.Bottom; Left = Option.map tableBorder value.Left
      InsideH = Option.map tableBorder value.InsideH; InsideV = Option.map tableBorder value.InsideV }
let private shading (value: WordEditorPackage.Shading) : Shading =
    { Pattern = value.Pattern; ForegroundColor = value.ForegroundColor; BackgroundColor = value.BackgroundColor }
let private width (amount: float option) kind =
    let kind = match kind with Some "dxa" -> DxaWidth | Some "pct" -> PercentWidth | Some "auto" | Some "nil" -> AutoWidth | _ -> PointWidth
    if amount.IsSome || kind = AutoWidth then Some { Kind = kind; Value = amount } else None
let private pixels value = value |> Option.map (fun x -> x / WordLayout.canvasScale.DxaPerUnit)

let private listLevel (value: WordEditorPackage.ListLevel) : ListLevelDefinition =
    { Level = value.Level; NumberFormat = value.NumberFormat |> Option.defaultValue "decimal"
      LevelText = value.LevelText; Start = value.Start; IndentLeft = value.IndentLeft
      FirstLineIndent = value.FirstLineIndent; Suffix = value.Suffix
      Justification = value.Justification; Legal = value.Legal; Restart = value.Restart
      MarkFont = value.MarkFont; MarkBold = value.MarkBold; MarkItalic = value.MarkItalic
      MarkColor = value.MarkColor; MarkSize = value.MarkSize }

type LevelCache = System.Collections.Generic.Dictionary<WordEditorPackage.ListLevel list, ListLevelDefinition list>

let newLevelCache () = LevelCache(HashIdentity.Structural)

let private convertedLevels (cache: LevelCache) levels =
    if List.isEmpty levels then []
    else
        match Portable.Dict.tryFind levels cache with
        | ValueSome converted -> converted
        | ValueNone ->
            let converted = List.map listLevel levels
            cache.Add(levels, converted)
            converted

let private metadata cache (item: WordEditorPackage.EditorParagraph) =
    let f = item.Formatting
    let listStyle =
        match item.List |> Option.map (fun x -> x.Style.ToLowerInvariant()) with
        | Some "bullet" -> Bullet
        | Some "numbered" -> Numbered
        | _ -> NoList
    let flag condition bit = if condition then bit else 0
    let flags =
        if f.ExplicitParagraphFormat then
            flag f.ExplicitSpacingBefore.IsSome ParagraphExplicit.SpacingBefore |||
            flag f.ExplicitSpacingAfter.IsSome ParagraphExplicit.SpacingAfter |||
            flag f.ExplicitLineSpacing.IsSome ParagraphExplicit.LineSpacing |||
            flag f.ExplicitIndentLeft.IsSome ParagraphExplicit.IndentLeft |||
            flag f.ExplicitIndentRight.IsSome ParagraphExplicit.IndentRight |||
            flag f.ExplicitFirstLineIndent.IsSome ParagraphExplicit.FirstLineIndent |||
            flag f.ExplicitContextualSpacing.IsSome ParagraphExplicit.ContextualSpacing |||
            flag f.ExplicitBackgroundColor.IsSome ParagraphExplicit.BackgroundColor |||
            flag f.ExplicitTabs ParagraphExplicit.Tabs
        else
            flag f.SpacingBefore.IsSome ParagraphExplicit.SpacingBefore |||
            flag f.SpacingAfter.IsSome ParagraphExplicit.SpacingAfter |||
            flag f.LineSpacing.IsSome ParagraphExplicit.LineSpacing |||
            flag f.IndentLeft.IsSome ParagraphExplicit.IndentLeft |||
            flag f.IndentRight.IsSome ParagraphExplicit.IndentRight |||
            flag f.FirstLineIndent.IsSome ParagraphExplicit.FirstLineIndent |||
            flag f.ContextualSpacing.IsSome ParagraphExplicit.ContextualSpacing |||
            flag f.BackgroundColor.IsSome ParagraphExplicit.BackgroundColor |||
            flag (not f.TabStops.IsEmpty || f.RightTabPosition.IsSome) ParagraphExplicit.Tabs
    let flags = flags ||| flag (f.PageBreakBefore && f.PageBreakBeforeDerived) ParagraphExplicit.DerivedPageBreakBefore ||| flag (f.PageBreakBefore && f.PageBreakAtStart) ParagraphExplicit.PageBreakRunBefore
    let listValue pick = item.List |> Option.bind pick
    { defaultMetadata with
        Style = f.Style; ContentControlId = f.ContentControlId; ContentControlTag = f.ContentControlTag
        ContentControlAlias = f.ContentControlAlias; ContentControlXml = f.ContentControlXml
        ContentControlGroup = f.ContentControlGroup
        ContentControlAncestors = f.ContentControlAncestors |> List.map (fun layer ->
            ({ Group = layer.Group; Id = layer.Id; Tag = layer.Tag; Alias = layer.Alias; Xml = layer.Xml } : ContentControlLayer))
        EnclosingCellControls = f.EnclosingCellControls |> List.map (fun layer ->
            ({ Group = layer.Group; Id = layer.Id; Tag = layer.Tag; Alias = layer.Alias; Xml = layer.Xml } : ContentControlLayer))
        Origin = origin item.Source.Origin; Alignment = alignment f.Alignment
        ExplicitAlignment = f.ExplicitAlignment.IsSome; ListStyle = listStyle
        ListLevel = if listStyle = NoList then 0 else item.List |> Option.map (fun x -> max 1 x.Level) |> Option.defaultValue 1
        ListId = listValue (fun x -> x.Id); ListSourceId = listValue (fun x -> x.SourceId)
        ListRestart = item.List |> Option.exists (fun x -> x.Restart)
        ListFormatEdited = item.List |> Option.exists (fun x -> x.FormatEdited)
        ListNumberFormat = listValue (fun x -> x.NumberFormat); ListLevelText = listValue (fun x -> x.LevelText)
        ListStart = listValue (fun x -> x.Start)
        ListStartOverrides = item.List |> Option.map (fun x -> x.StartOverrides) |> Option.defaultValue []
        ListLevels = if listStyle = NoList then [] else item.List |> Option.map (fun x -> convertedLevels cache x.Levels) |> Option.defaultValue []
        IndentLeft = f.IndentLeft; IndentRight = f.IndentRight; FirstLineIndent = f.FirstLineIndent
        RightTabPosition = f.RightTabPosition; RightTabLeader = f.RightTabLeader
        TabStops = f.TabStops |> List.map (fun x -> { Position = x.Position; Alignment = x.Alignment; Leader = x.Leader })
        Rtl = f.Rtl; SpacingBefore = f.SpacingBefore; SpacingAfter = f.SpacingAfter
        ParagraphMarkFontSize = f.ParagraphMarkFontSize; ParagraphMarkFontFamily = f.ParagraphMarkFontFamily; ContextualSpacing = f.ContextualSpacing
        LineSpacing = f.LineSpacing; LineSpacingRule = f.LineSpacingRule; WidowControl = f.WidowControl
        KeepNext = f.KeepNext; KeepLines = f.KeepLines; BackgroundColor = f.BackgroundColor
        ParagraphShadingPattern = f.ShadingPattern; ParagraphShadingForeground = f.ShadingForeground
        ExplicitParagraphShadingPattern = f.ExplicitShadingPattern; ExplicitParagraphShadingForeground = f.ExplicitShadingForeground
        ParagraphBorders = Option.map paragraphBorders f.Borders
        PageBreakBefore = f.PageBreakBefore; PageBreakAfter = f.PageBreakAfter
        ColumnBreakAfter = f.PageBreakAfter && f.ColumnBreakAfter
        SectionBreak = item.SectionBreak; SectionPageNumberStart = item.SectionPageNumberStart
        SectionPageNumberFormat = item.SectionPageNumberFormat; SectionVerticalAlign = item.SectionVerticalAlign
        OutlineLevel = f.OutlineLevel; RenderedPageBreakBefore = f.RenderedPageBreakBefore
        SectionPageSetup = Option.map sectionSetup item.SectionPageSetup
        EffectiveParagraphBorders = Option.map paragraphBorders f.EffectiveBorders
        OverflowPunctuation = f.OverflowPunctuation
        ListLevelIndents = item.List |> Option.bind (fun x -> x.ListLevelIndents)
        DropCap = f.DropCapKind |> Option.map (fun kind -> { InMargin = kind = "margin"; Lines = max 1 (Option.defaultValue 1 f.DropCapLines); Distance = Option.defaultValue 0.0 f.DropCapDistance })
        SectionColumns = item.SectionColumns |> Option.filter (fun x -> x.Count >= 2) |> Option.map columns
        HeaderReferences = references item.HeaderReferences; FooterReferences = references item.FooterReferences
        Explicit = flags; HasMarkProperties = f.HasMarkProperties; ExplicitMarkFontSize = f.ExplicitMarkFontSize
        MarkRevision = f.MarkRevision |> Option.bind (fun r -> r.Kind |> Option.map (fun kind -> { RevisionType = kind; Mark = { Id = r.Id; Author = r.Author; Date = r.Date } })) }

let private bookmark (value: ModelText.BookmarkStart) : BookmarkStart =
    { Id = value.Id; Name = value.Name; EndsHere = value.EndsHere }

let private runStyle (run: WordEditorPackage.EditorRun) =
    let f = run.Formatting
    let flag value bit = if value then bit else 0
    let flags =
        flag f.ExplicitBold.IsSome ExplicitFormat.Bold ||| flag f.ExplicitItalic.IsSome ExplicitFormat.Italic |||
        flag f.ExplicitUnderline.IsSome ExplicitFormat.Underline ||| flag f.ExplicitAllCaps.IsSome ExplicitFormat.AllCaps |||
        flag f.ExplicitSmallCaps.IsSome ExplicitFormat.SmallCaps ||| flag f.ExplicitStrike.IsSome ExplicitFormat.Strike |||
        flag f.ExplicitDoubleStrike.IsSome ExplicitFormat.DoubleStrike ||| flag f.ExplicitHidden.IsSome ExplicitFormat.Hidden |||
        flag f.ExplicitFontSize.IsSome ExplicitFormat.FontSize ||| flag (f.FontFamilyInherited = Some false) ExplicitFormat.FontFamily |||
        flag (f.ExplicitColorAuto && f.Color.IsNone) ExplicitFormat.ColorAuto
    let link = run.Hyperlink |> Option.filter isActiveHyperlinkTarget
    let inactive = run.InactiveHyperlink |> Option.orElse (run.Hyperlink |> Option.filter (isActiveHyperlinkTarget >> not))
    let revision = run.Revision
    { TextBuffer.defaultStyle with
        Style = f.Style; Bold = Option.defaultValue f.Bold f.ExplicitBold
        Italic = Option.defaultValue f.Italic f.ExplicitItalic
        Underline = Option.defaultValue f.Underline f.ExplicitUnderline
        UnderlineStyle = f.UnderlineStyle |> Option.filter ((<>) "single")
        UnderlineColor = f.UnderlineColor; AllCaps = Option.defaultValue f.AllCaps f.ExplicitAllCaps
        SmallCaps = Option.defaultValue f.SmallCaps f.ExplicitSmallCaps; Strike = Option.defaultValue f.Strike f.ExplicitStrike
        DoubleStrike = Option.defaultValue f.DoubleStrike f.ExplicitDoubleStrike; Hidden = Option.defaultValue f.Hidden f.ExplicitHidden
        Position = f.Position; EmphasisMark = f.EmphasisMark; FontSize = Option.defaultValue f.FontSize f.ExplicitFontSize
        CharacterSpacing = f.CharacterSpacing; FontFamily = if f.FontFamily = "" then None else Some f.FontFamily
        Color = f.Color; Highlight = f.Highlight; VertAlign = f.VertAlign; EastAsiaLanguage = f.EastAsiaLanguage; Hyperlink = link; InactiveHyperlink = inactive
        CommentId = run.CommentId; CommentText = run.CommentText; CommentAuthor = run.CommentAuthor; CommentDate = run.CommentDate
        NoteId = run.NoteId; NoteType = run.NoteType; NoteText = run.NoteText
        BookmarkStarts = List.map bookmark run.BookmarkStarts
        BookmarkEnds = run.BookmarkEnds |> List.map (function ModelText.Ends id -> Ends id | ModelText.Opens start -> Opens (bookmark start))
        FieldCode = run.FieldCode; FieldResult = run.FieldResult
        FieldBoundary =
            match run.Content with
            | WordEditorPackage.FieldBoundary boundary ->
                match boundary with
                | ModelText.Begin -> Some TextBuffer.FieldBoundary.Begin
                | ModelText.Instruction text -> Some (TextBuffer.FieldBoundary.Instruction text)
                | ModelText.Separate -> Some TextBuffer.FieldBoundary.Separate
                | ModelText.End -> Some TextBuffer.FieldBoundary.End
            | _ -> None
        RenderedPageBreakBefore = run.RenderedPageBreakBefore
        EquationXml = match run.Content with WordEditorPackage.Equation (_, xml) -> Some xml | _ -> None
        RevisionId = revision |> Option.bind (fun r -> r.Id); RevisionType = revision |> Option.bind (fun r -> r.Kind)
        RevisionAuthor = revision |> Option.bind (fun r -> r.Author); RevisionDate = revision |> Option.bind (fun r -> r.Date)
        DeletedInsertion = revision |> Option.bind (fun r -> r.Inserted) |> Option.map (fun r -> { Id = r.Id; Author = r.Author; Date = r.Date })
        MoveName = revision |> Option.bind (fun r -> r.MoveName); SourceRun = run.SourceRun; Checkbox = run.Checkbox
        ContentControls = run.ContentControls |> List.map (fun layer ->
            ({ Group = layer.Group; Id = layer.Id; Tag = layer.Tag; Alias = layer.Alias; Xml = layer.Xml } : TextBuffer.ContentControlLayer))
        Explicit = flags }

let private buffer (runs: WordEditorPackage.EditorRun array) =
    let text = System.Text.StringBuilder()
    let output = ResizeArray<Run>()
    let addText (run: WordEditorPackage.EditorRun) value =
        if not (System.String.IsNullOrEmpty value) then
            let start = text.Length
            text.Append(value) |> ignore
            let style = runStyle run
            if style.RenderedPageBreakBefore && value.Length > 1 then
                output.Add(Text { Start = start; Length = 1; Style = style })
                output.Add(Text { Start = start + 1; Length = value.Length - 1; Style = { style with RenderedPageBreakBefore = false } })
            else output.Add(Text { Start = start; Length = value.Length; Style = style })
    for run in runs do
        match run.Content with
        | WordEditorPackage.Text value -> addText run value
        | WordEditorPackage.Equation (value, _) -> addText run value
        | WordEditorPackage.FieldBoundary _ ->
            output.Add(Text { Start = text.Length; Length = 0; Style = runStyle run })
        | WordEditorPackage.PageBreak -> ()
        | WordEditorPackage.Picture image ->
            let start = text.Length
            text.Append(objectReplacementChar) |> ignore
            output.Add(Image {
                Start = start; Src = Option.defaultValue "" image.Source; Linked = image.Linked
                Width = image.Width; Height = image.Height; Alt = image.Alt; Name = image.Name; Title = image.Title
                Floating = image.Floating; Wrap = image.Wrap; PositionX = image.PositionX; PositionY = image.PositionY
                PositionHRelativeFrom = image.PositionHRelativeFrom; PositionVRelativeFrom = image.PositionVRelativeFrom
                BehindDoc = image.BehindDoc; AllowOverlap = image.AllowOverlap; LayoutInCell = image.LayoutInCell
                DistanceTop = image.DistanceTop; DistanceBottom = image.DistanceBottom; DistanceLeft = image.DistanceLeft
                DistanceRight = image.DistanceRight; RelativeHeight = image.RelativeHeight; Locked = image.Locked
                SimplePosition = image.SimplePosition; SimplePositionX = image.SimplePositionX; SimplePositionY = image.SimplePositionY
                CropLeft = image.CropLeft; CropTop = image.CropTop; CropRight = image.CropRight; CropBottom = image.CropBottom
                PositionHAlign = image.PositionHAlign; PositionVAlign = image.PositionVAlign; WrapText = image.WrapText
                WrapPolygon = image.WrapPolygon; FlipH = image.FlipH; FlipV = image.FlipV
                Opaque = image.Opaque |> Option.map (fun o -> { Xml = o.Xml; Kind = o.Kind; ParagraphLevel = o.ParagraphLevel }); SourceRun = run.SourceRun })
        | WordEditorPackage.Opaque opaque ->
            let start = text.Length
            text.Append(objectReplacementChar) |> ignore
            output.Add(Image {
                Start = start; Src = opaqueSourcePrefix + opaque.Kind; Linked = false
                Width = None; Height = None; Alt = None; Name = None; Title = None; Floating = false; Wrap = None
                PositionX = None; PositionY = None; PositionHRelativeFrom = None; PositionVRelativeFrom = None
                BehindDoc = None; AllowOverlap = None; LayoutInCell = None; DistanceTop = None; DistanceBottom = None
                DistanceLeft = None; DistanceRight = None; RelativeHeight = None; Locked = None; SimplePosition = None
                SimplePositionX = None; SimplePositionY = None; CropLeft = None; CropTop = None; CropRight = None; CropBottom = None
                PositionHAlign = None; PositionVAlign = None; WrapText = None; WrapPolygon = []; FlipH = false; FlipV = false
                Opaque = Some { Xml = opaque.Xml; Kind = opaque.Kind; ParagraphLevel = opaque.ParagraphLevel }; SourceRun = run.SourceRun })
    { Rope = Rbt.Word.Rope.ofString (text.ToString()); Runs = List.ofSeq output }

let mutable private importedBlockCounter = 1000
let private nextId () =
#if FABLE_COMPILER
    importedBlockCounter <- importedBlockCounter + 1
    "imported-block-" + string importedBlockCounter
#else
    "imported-block-" + string (System.Threading.Interlocked.Increment &importedBlockCounter)
#endif

let private paragraph cache (item: WordEditorPackage.EditorParagraph) : ParagraphBlock =
    { Id = nextId (); Buffer = buffer item.Runs; Metadata = metadata cache item }

let private geometry (shape: WordEditorPackage.EditorTextBox) =
    let segment (op, p: float array) =
        match op, p with
        | "M", [| a; b |] -> Some (PMove (a, b))
        | "L", [| a; b |] -> Some (PLine (a, b))
        | "C", [| a; b; c; d; e; f |] -> Some (PCubic (a, b, c, d, e, f))
        | "Q", [| a; b; c; d |] -> Some (PQuad (a, b, c, d))
        | "Z", _ -> Some PClose
        | _ -> None
    match shape.Geometry with
    | "custom" ->
        let path = List.choose segment shape.CustomPath
        if path.IsEmpty || shape.CustomWidth <= 0.0 || shape.CustomHeight <= 0.0 then GeoRect
        else GeoCustom (path, shape.CustomWidth, shape.CustomHeight)
    | "" | "rect" -> GeoRect
    | preset -> GeoPreset preset

let rec private shape cache (item: WordEditorPackage.EditorTextBox) : TextBoxBlock =
    { defaultTextBox with
        Id = nextId (); SourceShapeId = item.ShapeId; SourceShapeKind = item.ShapeKind
        Content = paragraphs cache item.Paragraphs; X = item.X; Y = item.Y
        PositionHRelativeFrom = item.PositionHRelativeFrom; PositionVRelativeFrom = item.PositionVRelativeFrom
        Width = item.Width; Height = item.Height
        FillColor = if item.NoFill then None else Option.orElse (Some "#FFFFFF") item.FillColor
        BorderColor = if item.NoLine then None else Option.orElse (Some "#808080") item.BorderColor
        BorderWidth = Option.orElse (Some 1.0) item.BorderWidth; RelativeHeight = item.RelativeHeight; Geometry = geometry item
        Gradient = if item.Gradient.IsEmpty then None else Some { Stops = item.Gradient |> List.map (fun (offset, color) -> { Offset = offset; Color = color }); Angle = item.GradientAngle; Radial = item.GradientRadial }
        Rotation = item.Rotation; FlipH = item.FlipH; FlipV = item.FlipV; Inline = item.Inline; Wrap = item.Wrap
        PositionHAlign = item.PositionHAlign; PositionVAlign = item.PositionVAlign; BehindDoc = item.BehindDoc
        AllowOverlap = item.AllowOverlap; LayoutInCell = item.LayoutInCell; DistanceTop = item.DistanceTop
        DistanceBottom = item.DistanceBottom; DistanceLeft = item.DistanceLeft; DistanceRight = item.DistanceRight
        BorderDash = item.BorderDash; Name = item.Name; InsetLeft = item.InsetLeft; InsetTop = item.InsetTop
        InsetRight = item.InsetRight; InsetBottom = item.InsetBottom; VerticalAnchor = item.VerticalAnchor
        AutoFit = item.AutoFit; NoTextWrap = item.NoTextWrap; FitText = item.FitText; Opacity = item.Opacity
        Children = item.Children |> Array.map (shape cache) |> Array.toList; ChildSpace = item.ChildSpace
        ImageSource = item.ImageSource; SourceXml = item.SourceXml }
and private paragraphs cache items =
    items |> Array.choose (function WordEditorPackage.Paragraph p -> Some (paragraph cache p) | _ -> None) |> Array.toList
and private table cache (item: WordEditorPackage.EditorTable) : TableBlock =
    let cell (c: WordEditorPackage.EditorCell) : TableCell =
        { Content = paragraphs cache c.Paragraphs
          NestedTables = c.NestedTables |> Array.map (fun t -> let result = table cache t in { result with Rows = WordLayout.Table.distributeVerticalMergeContent result.Rows }) |> Array.toList
          NestedTableAnchors = Array.toList c.NestedTableAnchors; ColSpan = max 1 c.ColSpan
          EnclosingControls = c.EnclosingControls |> List.map (fun layer ->
              ({ Group = layer.Group; Id = layer.Id; Tag = layer.Tag; Alias = layer.Alias; Xml = layer.Xml } : ContentControlLayer))
          BackgroundColor = c.BackgroundColor; Shading = Option.map shading c.Shading; Width = c.Width
          PreferredWidth = width c.PreferredWidth c.PreferredWidthType; BorderColor = c.BorderColor; BorderWidth = c.BorderWidth
          Borders = Option.map tableBorders c.Borders; VerticalAlignment = c.VerticalAlignment
          VerticalMerge = match c.VerticalMerge with Some "restart" -> Some RestartMerge | Some "continue" -> Some ContinueMerge | _ -> None
          PaddingTop = pixels c.PaddingTop; PaddingRight = pixels c.PaddingRight; PaddingBottom = pixels c.PaddingBottom
          PaddingLeft = pixels c.PaddingLeft; Origin = origin c.Source.Origin }
    let row (r: WordEditorPackage.EditorRow) : TableRow =
        { Cells = r.Cells |> Array.map cell |> Array.toList; IsHeader = r.IsHeader; Height = r.Height
          HeightRule = r.HeightRule; CanSplit = r.CanSplit; CellSpacing = pixels r.CellSpacing
          PaddingTop = pixels r.PaddingTop; PaddingRight = pixels r.PaddingRight; PaddingBottom = pixels r.PaddingBottom
          PaddingLeft = pixels r.PaddingLeft; GridBefore = max 0 r.GridBefore; GridAfter = max 0 r.GridAfter
          WidthBefore = pixels r.WidthBefore; WidthAfter = pixels r.WidthAfter
          BordersOverride = Option.map tableBorders r.BordersOverride; Origin = origin r.Source.Origin }
    { Id = nextId (); Rows = item.Rows |> Array.map row |> Array.toList
      PreferredWidth = width item.PreferredWidth item.PreferredWidthType; Alignment = item.Alignment
      Indent = pixels item.Indent
      LayoutMode = match item.LayoutMode with Some "fixed" -> Some FixedTableLayout | Some "autofit" | Some "autoFit" -> Some AutoFitTableLayout | _ -> None
      GridColumns = item.GridColumns |> Array.map (fun n -> { Kind = DxaWidth; Value = Some n }) |> Array.toList
      Borders = Option.map tableBorders item.Borders; BackgroundColor = item.BackgroundColor; Shading = Option.map shading item.Shading
      CellSpacing = pixels item.CellSpacing; PaddingTop = pixels item.PaddingTop; PaddingRight = pixels item.PaddingRight
      PaddingBottom = pixels item.PaddingBottom; PaddingLeft = pixels item.PaddingLeft; Floating = item.Floating
      FloatingXml = item.FloatingXml; StyleId = item.StyleId
      ContentControlId = item.ContentControlId; ContentControlTag = item.ContentControlTag
      ContentControlAlias = item.ContentControlAlias; ContentControlXml = item.ContentControlXml
      ContentControlGroup = item.ContentControlGroup
      ContentControlAncestors = item.ContentControlAncestors |> List.map (fun layer ->
          ({ Group = layer.Group; Id = layer.Id; Tag = layer.Tag; Alias = layer.Alias; Xml = layer.Xml } : ContentControlLayer))
      Look = item.Look |> Option.map (fun x -> { FirstRow = x.FirstRow; LastRow = x.LastRow; FirstColumn = x.FirstColumn; LastColumn = x.LastColumn; HorizontalBanding = x.HorizontalBanding; VerticalBanding = x.VerticalBanding }) }

let importBlock cache item =
    match item with
    | WordEditorPackage.Paragraph p ->
        let result = paragraph cache p
        match p.HeadingLevel with Some level -> Heading (min 3 (max 1 level), result) | None -> Paragraph result
    | WordEditorPackage.Table t -> Table (table cache t)
    | WordEditorPackage.TextBox s -> TextBox (shape cache s)

let private blocks cache (items: WordEditorPackage.EditorBlock array) =
    let mapped = Array.map (importBlock cache) items
    let idAt i =
        if i < 0 || i >= mapped.Length then None else
        match mapped.[i] with Paragraph p | Heading (_, p) -> Some p.Id | Table t -> Some t.Id | TextBox _ -> None
    mapped |> Array.mapi (fun i block ->
        match items.[i], block with
        | WordEditorPackage.TextBox source, TextBox box ->
            TextBox { box with AnchorBlockId = source.AnchorBlock |> Option.bind idAt }
        | _ -> block) |> Array.toList

let private namedStyle (s: WordEditorPackage.Style) : NamedStyle =
    let text = s.Run
    let para = s.Paragraph
    let get pick = text |> Option.bind pick
    let p pick = para |> Option.bind pick
    { Id = s.Id; Name = s.Name; StyleType = s.StyleType; IsDefault = s.IsDefault
      BasedOn = s.BasedOn; Next = s.Next; Linked = s.Linked
      IsQuickStyle = Option.defaultValue false s.IsQuickStyle; Aliases = s.Aliases; UiPriority = s.UiPriority
      SemiHidden = Option.defaultValue false s.SemiHidden; UnhideWhenUsed = Option.defaultValue false s.UnhideWhenUsed
      Text = text |> Option.filter (fun r ->
        r.Bold.IsSome || r.Italic.IsSome || r.Underline.IsSome || r.AllCaps.IsSome ||
        r.SmallCaps.IsSome || r.Strike.IsSome || r.VertAlign.IsSome || r.FontSize.IsSome ||
        r.CharacterSpacing.IsSome || r.FontFamily.IsSome || r.Color.IsSome || r.Highlight.IsSome)
        |> Option.map (fun r ->
            { TextBuffer.defaultStyle with
                Bold = Option.defaultValue false r.Bold; Italic = Option.defaultValue false r.Italic
                Underline = Option.defaultValue false r.Underline; UnderlineStyle = r.UnderlineStyle |> Option.filter ((<>) "single")
                UnderlineColor = r.UnderlineColor; AllCaps = Option.defaultValue false r.AllCaps; SmallCaps = Option.defaultValue false r.SmallCaps
                Strike = Option.defaultValue false r.Strike; DoubleStrike = Option.defaultValue false r.DoubleStrike
                Hidden = Option.defaultValue false r.Hidden; Position = Option.defaultValue 0.0 r.Position; EmphasisMark = r.EmphasisMark
                VertAlign = r.VertAlign; FontSize = Option.defaultValue TextBuffer.defaultStyle.FontSize r.FontSize
                Explicit = if r.FontSize.IsSome then ExplicitFormat.FontSize else 0
                CharacterSpacing = Option.defaultValue 0.0 r.CharacterSpacing; FontFamily = r.FontFamily; Color = r.Color; Highlight = r.Highlight })
      TextBold = get (fun r -> r.Bold); TextItalic = get (fun r -> r.Italic); TextUnderline = get (fun r -> r.Underline)
      TextAllCaps = get (fun r -> r.AllCaps); TextSmallCaps = get (fun r -> r.SmallCaps); TextStrike = get (fun r -> r.Strike)
      SpacingBefore = p (fun x -> x.SpacingBefore); SpacingAfter = p (fun x -> x.SpacingAfter)
      ContextualSpacing = p (fun x -> x.ContextualSpacing); LineSpacing = p (fun x -> x.LineSpacing)
      LineSpacingRule = p (fun x -> x.LineSpacingRule); Alignment = p (fun x -> x.Alignment) |> Option.map alignment
      IndentLeft = p (fun x -> x.IndentLeft); IndentRight = p (fun x -> x.IndentRight); FirstLineIndent = p (fun x -> x.FirstLineIndent)
      BackgroundColor = p (fun x -> x.BackgroundColor); OutlineLevel = p (fun x -> x.OutlineLevel)
      KeepNext = p (fun x -> x.KeepNext); KeepLines = p (fun x -> x.KeepLines)
      PageBreakBefore = p (fun x -> x.PageBreakBefore); WidowControl = p (fun x -> x.WidowControl) }

let private inferredStyle kind id (info: WordEditorPackage.InferredStyle option) =
    let get pick = info |> Option.bind pick
    let hasText = (get (fun x -> x.Bold)).IsSome || (get (fun x -> x.Italic)).IsSome || (get (fun x -> x.Underline)).IsSome ||
                  (get (fun x -> x.AllCaps)).IsSome || (get (fun x -> x.SmallCaps)).IsSome || (get (fun x -> x.Strike)).IsSome ||
                  (get (fun x -> x.VertAlign)).IsSome || (get (fun x -> x.FontSize)).IsSome ||
                  (get (fun x -> x.CharacterSpacing)).IsSome || (get (fun x -> x.FontFamily)).IsSome ||
                  (get (fun x -> x.Color)).IsSome || (get (fun x -> x.Highlight)).IsSome
    let text : WordEditorPackage.StyleRun option =
        if not hasText then None else
        Some { Bold = get (fun x -> x.Bold); Italic = get (fun x -> x.Italic); Underline = get (fun x -> x.Underline)
               UnderlineStyle = get (fun x -> x.UnderlineStyle); UnderlineColor = get (fun x -> x.UnderlineColor)
               AllCaps = get (fun x -> x.AllCaps); SmallCaps = get (fun x -> x.SmallCaps)
               Strike = get (fun x -> x.Strike); DoubleStrike = get (fun x -> x.DoubleStrike); Hidden = get (fun x -> x.Hidden)
               Position = get (fun x -> x.Position); EmphasisMark = get (fun x -> x.EmphasisMark)
               VertAlign = get (fun x -> x.VertAlign); FontSize = get (fun x -> x.FontSize); CharacterSpacing = get (fun x -> x.CharacterSpacing)
               FontFamily = get (fun x -> x.FontFamily); Color = get (fun x -> x.Color); Highlight = get (fun x -> x.Highlight) }
    let para : WordEditorPackage.StyleParagraph option =
        if kind <> "paragraph" then None else
        Some { SpacingBefore = get (fun x -> x.SpacingBefore); SpacingAfter = get (fun x -> x.SpacingAfter)
               ContextualSpacing = get (fun x -> x.ContextualSpacing); LineSpacing = get (fun x -> x.LineSpacing)
               LineSpacingRule = get (fun x -> x.LineRule); Alignment = get (fun x -> x.Alignment)
               IndentLeft = get (fun x -> x.IndentLeft); IndentRight = get (fun x -> x.IndentRight); FirstLineIndent = get (fun x -> x.FirstLineIndent)
               BackgroundColor = get (fun x -> x.BackgroundColor); OutlineLevel = get (fun x -> x.OutlineLevel)
               KeepNext = get (fun x -> x.KeepNext); KeepLines = get (fun x -> x.KeepLines)
               PageBreakBefore = get (fun x -> x.PageBreakBefore); WidowControl = get (fun x -> x.WidowControl) }
    namedStyle { Id = id; Name = get (fun x -> x.Name); StyleType = kind; IsDefault = get (fun x -> x.IsDefault) |> Option.defaultValue false
                 BasedOn = get (fun x -> x.BasedOn); Next = get (fun x -> x.Next); Linked = get (fun x -> x.Linked)
                 IsQuickStyle = get (fun x -> x.IsQuickStyle); Aliases = get (fun x -> x.Aliases); UiPriority = get (fun x -> x.UiPriority)
                 SemiHidden = get (fun x -> x.SemiHidden); UnhideWhenUsed = get (fun x -> x.UnhideWhenUsed); Run = text; Paragraph = para }

let rec inferredStylesOfBlock item =
    match item with
    | WordEditorPackage.Paragraph p ->
        [ match p.Formatting.Style with Some id -> yield inferredStyle "paragraph" id p.Formatting.StyleInfo | None -> ()
          for run in p.Runs do
              match run.Formatting.Style with Some id -> yield inferredStyle "character" id run.Formatting.StyleInfo | None -> () ]
    | WordEditorPackage.Table t ->
        [ for row in t.Rows do
            for cell in row.Cells do
                for child in cell.Paragraphs do yield! inferredStylesOfBlock child
                for nested in cell.NestedTables do yield! inferredStylesOfBlock (WordEditorPackage.Table nested) ]
    | WordEditorPackage.TextBox s -> [ for child in s.Paragraphs do yield! inferredStylesOfBlock child ]

let private paragraphBlocks blocks =
    blocks |> List.choose (function Paragraph p | Heading (_, p) -> Some p | _ -> None)

let importPackageWithCache cache (package: WordEditorPackage.EditorPackage) : Document =
    let body = blocks cache package.Blocks
    let parts (items: WordEditorPackage.HeaderFooterPart array) =
        items |> Array.map (fun p -> { Kind = p.Kind; RelationshipId = p.RelationshipId; Blocks = blocks cache p.Blocks } : HeaderFooterPart) |> Array.toList
    let headers = parts package.HeaderParts
    let footers = parts package.FooterParts
    let stories kind (items: WordEditorPackage.Story array) =
        items |> Array.map (fun s ->
            { NoteType = kind; Id = s.Id; Kind = s.Kind; HasMark = Option.defaultValue true s.HasMark
              MarkStyle = s.MarkStyle; Text = Option.defaultValue "" s.Text; Blocks = blocks cache s.Blocks } : NoteBody) |> Array.toList
    let comments =
        package.Comments |> Array.map (fun c ->
            { Id = c.Id; Author = c.Author; Date = c.Date; Initials = c.Initials
              Text = Option.defaultValue "" c.Text; Blocks = blocks cache c.Blocks } : CommentBody) |> Array.toList
    let styles =
        let seen = System.Collections.Generic.HashSet<string * string>()
        seq {
            for style in package.Styles do yield namedStyle style
            for block in package.Blocks do yield! inferredStylesOfBlock block
        } |> Seq.filter (fun s -> seen.Add (s.StyleType, s.Id)) |> Seq.toList
    let byId = comments |> List.map (fun c -> c.Id, c) |> Map.ofList
    let fill (para: ParagraphBlock) =
        let runs = para.Buffer.Runs |> List.map (function
            | Text run when run.Style.CommentId.IsSome && run.Style.CommentText.IsNone ->
                match run.Style.CommentId |> Option.bind (fun id -> Map.tryFind id byId) with
                | Some c ->
                    let style =
                        { run.Style with
                            CommentText = Some c.Text
                            CommentAuthor = Option.orElse c.Author run.Style.CommentAuthor
                            CommentDate = Option.orElse c.Date run.Style.CommentDate }
                    Text { run with Style = style }
                | None -> Text run
            | other -> other)
        { para with Buffer = { para.Buffer with Runs = runs } }
    let fillPart (part: HeaderFooterPart) = { part with Blocks = mapParagraphsOfBlocks fill part.Blocks }
    let body = if comments.IsEmpty then body else mapParagraphsOfBlocks fill body
    let headers = if comments.IsEmpty then headers else List.map fillPart headers
    let footers = if comments.IsEmpty then footers else List.map fillPart footers
    let selection =
        body |> List.tryPick (function
            | Paragraph p -> Some { Anchor = { BlockId = p.Id; Offset = 0 }; Focus = { BlockId = p.Id; Offset = 0 } }
            | _ -> None)
    let doc =
        { Blocks = body; Selection = selection; PageSetup = package.PageSetup |> Option.map pageSetup |> Option.defaultValue defaultPageSetup
          Styles = styles; HeaderParts = headers; FooterParts = footers
          Headers = headers |> List.tryHead |> Option.map (fun x -> paragraphBlocks x.Blocks) |> Option.defaultValue []
          Footers = footers |> List.tryHead |> Option.map (fun x -> paragraphBlocks x.Blocks) |> Option.defaultValue []
          Package = { emptyPackageData with Notes = stories "footnote" package.Footnotes @ stories "endnote" package.Endnotes
                                            Comments = comments; Theme = package.Theme; Preserved = package.Preserved; ImportedStyles = styles } }
    if package.Preserved.IsNone then doc
    else { doc with Package = { doc.Package with Imported = Some { Document = doc } } }

let importPackage package = importPackageWithCache (newLevelCache ()) package
