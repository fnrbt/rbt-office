module WordEditorPackageReader

open System
open WordEditorPackage
module D = DocxReader
module TC = DocumentProcessors.TypeConversion

let private number value =
    value |> Option.bind (fun (text: string) ->
        match Double.TryParse(text, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture) with
        | true, value -> Some value
        | _ -> None)
let private integer value =
    value |> Option.bind (fun (text: string) -> match Int32.TryParse text with true, value -> Some value | _ -> None)
let private pixels value = number value |> Option.map (fun value -> value / 15.0)
let private points value = number value |> Option.map (fun value -> value / 20.0)
let private on value = value |> Option.defaultValue false
let private firstLine first hanging =
    match pixels first, pixels hanging with
    | _, Some value -> Some -value
    | value, None -> value
let private source (span: XmlDsl.Span) : SourceFragment =
    { Origin = if XmlDsl.Span.isKnown span then Some (ReadAt span.Start) else None
      Xml = None; Copy = false; Same = false }
let private noSource = source XmlDsl.Span.none
let private references (items: D.HeaderFooterReference list) : HeaderFooterReference list =
    items |> List.map (fun item -> { Kind = item.Kind; RelationshipId = item.RelationshipId })
let private columns (value: D.SectionColumns) : SectionColumns =
    { Count = value.Count; Gap = value.Gap; Separator = value.Separator; Widths = value.Widths }
let private controlLayer (layer: D.ContentControlLayer) : ContentControlLayer =
    { Group = layer.Group; Id = layer.Id; Tag = layer.Tag; Alias = layer.Alias; Xml = layer.Xml }
let private borderColor value =
    match value with Some "" | Some "auto" -> Some "#000000" | value -> D.wordColorToCss value
let private cellBorder (value: D.BorderStyle) : Border =
    { Style = value.BorderType; Color = borderColor value.Color
      Width = number value.Size |> Option.map (fun width -> width / 6.0); Space = None }
let private cellBorders (value: D.CellBorders) : Borders =
    { Top = Option.map cellBorder value.Top; Right = Option.map cellBorder value.Right
      Bottom = Option.map cellBorder value.Bottom; Left = Option.map cellBorder value.Left
      InsideH = Option.map cellBorder value.InsideH; InsideV = Option.map cellBorder value.InsideV }
let private paraBorder (value: D.ParagraphBorder) : Border =
    { Style = value.Style; Color = borderColor value.Color
      Width = number value.Width |> Option.map (fun width -> width / 6.0); Space = number value.Space }
let private paraBorders (value: D.ParagraphBorders) : Borders =
    { Top = Option.map paraBorder value.Top; Right = Option.map paraBorder value.Right
      Bottom = Option.map paraBorder value.Bottom; Left = Option.map paraBorder value.Left
      InsideH = None; InsideV = None }
let private shading (value: D.Shading) : Shading =
    { Pattern = value.Pattern; ForegroundColor = D.wordShadingColorToEditor value.Color
      BackgroundColor = D.wordShadingColorToEditor value.Fill }
let private revision kind (value: D.RevisionMark) : Revision =
    { Id = value.Id; Kind = Some kind; Author = value.Author; Date = value.Date; Inserted = None; MoveName = None }
let private defaultPage : PageSetup =
    { PageWidth = 816.0; PageHeight = 1056.0; MarginLeft = 96.0; MarginRight = 96.0
      MarginTop = 96.0; MarginBottom = 96.0; HeaderDistance = 48.0; FooterDistance = 48.0
      DifferentFirstPage = false; EvenAndOddHeaders = false; DefaultTabStop = None
      Columns = None; SectionStart = None; HeaderReferences = []; FooterReferences = []
      PageNumberStart = None; PageNumberFormat = None; VerticalAlign = None }
let pageSetup (value: D.PageSetup) : PageSetup =
    { PageWidth = defaultArg value.PageWidth defaultPage.PageWidth; PageHeight = defaultArg value.PageHeight defaultPage.PageHeight
      MarginLeft = defaultArg value.MarginLeft defaultPage.MarginLeft; MarginRight = defaultArg value.MarginRight defaultPage.MarginRight
      MarginTop = defaultArg value.MarginTop defaultPage.MarginTop; MarginBottom = defaultArg value.MarginBottom defaultPage.MarginBottom
      HeaderDistance = defaultArg value.HeaderDistance defaultPage.HeaderDistance; FooterDistance = defaultArg value.FooterDistance defaultPage.FooterDistance
      DifferentFirstPage = on value.DifferentFirstPage; EvenAndOddHeaders = on value.EvenAndOddHeaders
      DefaultTabStop = value.DefaultTabStop; Columns = Option.map columns value.Columns; SectionStart = value.SectionStart
      HeaderReferences = references value.HeaderReferences; FooterReferences = references value.FooterReferences
      PageNumberStart = value.PageNumberStart; PageNumberFormat = value.PageNumberFormat; VerticalAlign = value.VerticalAlign }
let private sectionSetup (value: D.SectionGeometry) : PageSetup =
    { defaultPage with PageWidth = value.PageWidth; PageHeight = value.PageHeight
                       MarginLeft = value.MarginLeft; MarginRight = value.MarginRight
                       MarginTop = value.MarginTop; MarginBottom = value.MarginBottom
                       HeaderDistance = value.HeaderDistance; FooterDistance = value.FooterDistance
                       DifferentFirstPage = value.DifferentFirstPage }

// The source reader's style cache is synchronous. Native calls share the same
// critical section; the browser does not need or compile a thread primitive.
#if !FABLE_COMPILER
let private conversionGate = obj()
#endif
let private serialized action =
#if FABLE_COMPILER
    action ()
#else
    lock conversionGate action
#endif

let private runFormatting styleMap theme (para: D.ParagraphProperties) (run: D.Run) : RunFormatting =
    let direct = match run.Content with D.Link link -> link.Properties | _ -> run.Properties
    let effective = D.resolvedRunProperties styleMap para direct
    let mark = para.ParagraphRunProperties |> Option.defaultValue D.emptyRunProperties
    let explicit pick =
        match pick direct with
        | Some value -> Some value
        | None when (pick mark).IsSome -> Some (on (pick effective))
        | None -> None
    { Style = direct.Style; StyleInfo = None
      Bold = on effective.Bold; Italic = on effective.Italic; Underline = on effective.Underline
      UnderlineStyle = effective.UnderlineStyle; UnderlineColor = effective.UnderlineColor
      AllCaps = on effective.AllCaps; SmallCaps = on effective.SmallCaps; Strike = on effective.Strike
      DoubleStrike = on effective.DoubleStrike; Hidden = on effective.Hidden
      Position = D.wordPositionToPoints effective.Position; EmphasisMark = effective.EmphasisMark
      FontSize = D.wordFontSizeToPoints effective.FontSize
      CharacterSpacing = defaultArg (D.wordCharacterSpacingToCssPixels effective.CharacterSpacing) 0.0
      FontFamily = D.wordFontFamilyToCss theme effective
      FontFamilyInherited = Some (direct.FontFamily.IsNone && direct.ThemeFont.IsNone && mark.FontFamily.IsNone && mark.ThemeFont.IsNone)
      Color = D.wordColorToCss effective.Color; Highlight = effective.Highlight; VertAlign = effective.VertAlign
      HasExplicit = true; ExplicitBold = explicit (fun p -> p.Bold); ExplicitItalic = explicit (fun p -> p.Italic)
      ExplicitUnderline = explicit (fun p -> p.Underline); ExplicitAllCaps = explicit (fun p -> p.AllCaps)
      ExplicitSmallCaps = explicit (fun p -> p.SmallCaps); ExplicitStrike = explicit (fun p -> p.Strike)
      ExplicitDoubleStrike = explicit (fun p -> p.DoubleStrike); ExplicitHidden = explicit (fun p -> p.Hidden)
      ExplicitFontSize = if direct.FontSize.IsSome then Some (D.wordFontSizeToPoints direct.FontSize) elif mark.FontSize.IsSome then Some (D.wordFontSizeToPoints effective.FontSize) else None
      ExplicitColorAuto = (direct.Color |> Option.orElse mark.Color) = Some "auto" }

let private image images (value: D.ImageData) : Image =
    let src = D.editorImageSource images value
    { Source = Some src; Linked = value.Linked || D.isExternalImageSource src
      Width = value.Width; Height = value.Height; Alt = value.Description; Name = value.Name; Title = value.Title
      Floating = value.Floating; Wrap = value.Wrap; PositionX = value.PositionX; PositionY = value.PositionY
      PositionHRelativeFrom = value.PositionHRelativeFrom; PositionVRelativeFrom = value.PositionVRelativeFrom
      PositionHAlign = value.PositionHAlign; PositionVAlign = value.PositionVAlign
      BehindDoc = value.BehindDoc; AllowOverlap = value.AllowOverlap; LayoutInCell = value.LayoutInCell
      DistanceTop = value.DistanceTop; DistanceBottom = value.DistanceBottom; DistanceLeft = value.DistanceLeft; DistanceRight = value.DistanceRight
      RelativeHeight = value.RelativeHeight; Locked = value.Locked; SimplePosition = value.SimplePosition
      SimplePositionX = value.SimplePositionX; SimplePositionY = value.SimplePositionY
      CropLeft = value.CropLeft; CropTop = value.CropTop; CropRight = value.CropRight; CropBottom = value.CropBottom
      WrapText = value.WrapText; WrapPolygon = value.WrapPolygon; FlipH = value.FlipH; FlipV = value.FlipV; Opaque = value.Opaque }

let private convertRun styleMap theme images noteBodies para (value: D.Run) : EditorRun option =
    match value.Content with
    | D.RenderedPageBreak | D.PageBreak -> None
    | _ ->
        let formatting = runFormatting styleMap theme para value
        let initial =
            { Content = Text ""; Formatting = formatting
              Hyperlink = value.HyperlinkContext |> Option.bind (fun link -> link.Target) |> Option.filter D.isActiveHyperlinkTarget
              InactiveHyperlink = value.HyperlinkContext |> Option.bind (fun link -> link.Target) |> Option.filter (D.isActiveHyperlinkTarget >> not)
              CommentId = None; CommentText = None; CommentAuthor = None; CommentDate = None
              NoteId = None; NoteType = None; NoteText = None
              BookmarkStarts = value.BookmarkStarts; BookmarkEnds = value.BookmarkEnds
              FieldCode = None; FieldResult = None
              Revision =
                if value.RevisionType.IsSome || value.RevisionId.IsSome then
                    Some { Id = value.RevisionId; Kind = value.RevisionType; Author = value.RevisionAuthor
                           Date = value.RevisionDate; Inserted = value.InsertedRevision; MoveName = value.MoveName }
                else None
              RenderedPageBreakBefore = value.RenderedPageBreakBefore; Checkbox = value.Checkbox
              ContentControls = value.ContentControls |> List.map controlLayer
              SourceRun = if XmlDsl.Span.isKnown value.Source then Some value.Source.Start else None
              NoProps = false }
        let converted =
            match value.Content with
            | D.Text text -> { initial with Content = Text text }
            | D.Link link ->
                { initial with Content = Text link.Text
                               Hyperlink = link.Target |> Option.filter D.isActiveHyperlinkTarget
                               InactiveHyperlink = link.Target |> Option.filter (D.isActiveHyperlinkTarget >> not) }
            | D.Image picture -> { initial with Content = Picture (image images picture) }
            | D.EquationXml xml -> { initial with Content = Equation ("[Equation]", xml) }
            | D.CommentReference id ->
                { initial with Content = Text (sprintf "[Comment %s]" id); CommentId = Some id
                               Formatting = { formatting with Color = formatting.Color |> Option.orElse (Some D.commentMarkDisplayColor) } }
            | D.NoteRef note ->
                { initial with Content = Text (DocumentProcessors.EditorProcessor.noteMark note.NoteType note.Id)
                               NoteId = Some note.Id; NoteType = Some note.NoteType
                               NoteText = Map.tryFind (note.NoteType, note.Id) noteBodies
                               Formatting = { formatting with VertAlign = formatting.VertAlign |> Option.orElse (Some "superscript") } }
            | D.FieldRef field ->
                { initial with Content = Text (field.Part |> Option.orElse field.Result |> Option.defaultValue (sprintf "[Field:%s]" field.Code))
                               FieldCode = Some field.Code; FieldResult = field.Result }
            | D.FieldBoundary boundary ->
                { initial with Content = FieldBoundary boundary }
            | D.BookmarkStart bookmark ->
                { initial with BookmarkStarts = [ { Id = bookmark.Id; Name = defaultArg bookmark.Name ""; EndsHere = false } ] }
            | D.BookmarkEnd id -> { initial with BookmarkEnds = [ ModelText.Ends id ] }
            | D.RenderedPageBreak | D.PageBreak -> initial
        Some converted

let private emptyInferred : InferredStyle =
    { Name = None; IsDefault = None; BasedOn = None; Next = None; Linked = None; IsQuickStyle = None
      Aliases = None; UiPriority = None; SemiHidden = None; UnhideWhenUsed = None
      Bold = None; Italic = None; Underline = None; UnderlineStyle = None; UnderlineColor = None
      AllCaps = None; SmallCaps = None; Strike = None; DoubleStrike = None; Hidden = None
      Position = None; EmphasisMark = None; VertAlign = None; CharacterSpacing = None; FontSize = None; FontFamily = None
      Color = None; Highlight = None; SpacingBefore = None; SpacingAfter = None; ContextualSpacing = None
      LineSpacing = None; LineRule = None; Alignment = None; IndentLeft = None; IndentRight = None
      FirstLineIndent = None; BackgroundColor = None; OutlineLevel = None; KeepNext = None; KeepLines = None
      PageBreakBefore = None; WidowControl = None }

let private levelOf (value: D.NumberingDefinition) : ListLevel =
    { Level = value.Level + 1; NumberFormat = Some value.NumberFormat; LevelText = value.LevelText; Start = value.Start
      IndentLeft = pixels value.IndentLeft
      FirstLineIndent = pixels value.FirstLineIndent |> Option.orElse (pixels value.HangingIndent |> Option.map (~-))
      Suffix = value.Suffix; Justification = value.Justification; Legal = value.Legal; Restart = value.Restart
      MarkFont = value.MarkFont; MarkBold = value.MarkBold; MarkItalic = value.MarkItalic
      MarkColor = D.wordColorToCss value.MarkColor; MarkSize = value.MarkSize |> Option.map (Some >> D.wordFontSizeToPoints) }

let private paragraph styleMap theme images notes numbering styleNumbering levels defaultSpacing inCell (value: D.Paragraph) : EditorParagraph =
    let direct = value.Properties
    let effective = D.resolvedParagraphProperties styleMap direct
    let fontSizes =
        value.Runs |> List.choose (fun run ->
            match run.Content with
            | D.Text _ | D.Link _ | D.FieldRef _ -> Some (D.wordFontSizeToPoints (D.resolvedRunProperties styleMap effective run.Properties).FontSize)
            | _ -> None)
    let markSize =
        if inCell then
            match fontSizes with
            | [] -> D.wordFontSizeToPoints (D.resolvedRunProperties styleMap effective D.emptyRunProperties).FontSize
            | sizes -> List.max sizes
        else List.fold max (D.paragraphMarkFontSize styleMap effective value) fontSizes
    let breakAtStart, breakAfter = D.paragraphPageBreaks value
    let breakAtStart = not inCell && breakAtStart
    let breakAfter = if inCell then value.Runs |> List.exists (fun run -> run.Content = D.PageBreak) else breakAfter
    let spacing number auto =
        if auto = Some true then Some 14.0
        else points number |> Option.orElse (if defaultSpacing then Some 0.0 else None)
    let before = spacing effective.SpacingBefore effective.SpacingBeforeAuto
    let after = spacing effective.SpacingAfter effective.SpacingAfterAuto
    let line = D.wordLineSpacingToMultiple markSize effective.LineSpacing effective.LineRule
    let first = firstLine effective.FirstLineIndent effective.HangingIndent
    let onlyWhen condition value = if condition then value else None
    let keeps = D.styleKeepProperties styleMap value effective
    let format : ParagraphFormatting =
        { Style = effective.Style
          StyleInfo =
              if keeps.KeepNext.IsNone && keeps.KeepLines.IsNone && effective.PageBreakBefore.IsNone &&
                 effective.WidowControl.IsNone && effective.ContextualSpacing.IsNone then None
              else
                  Some { emptyInferred with
                            KeepNext = keeps.KeepNext
                            KeepLines = keeps.KeepLines
                            PageBreakBefore = effective.PageBreakBefore
                            WidowControl = effective.WidowControl
                            ContextualSpacing = effective.ContextualSpacing }
          Alignment = defaultArg effective.Alignment "left"; HasExplicit = true; ExplicitAlignment = direct.Alignment
          IndentLeft = pixels effective.IndentLeft; IndentRight = pixels effective.IndentRight; FirstLineIndent = first
          SpacingBefore = before; SpacingAfter = after; LineSpacing = line; LineSpacingRule = effective.LineRule
          RightTabPosition = pixels effective.RightTabPosition; RightTabLeader = effective.RightTabLeader
          TabStops = effective.TabStops |> List.choose (fun (position, alignment, leader) -> pixels (Some position) |> Option.map (fun position -> { Position = position; Alignment = alignment; Leader = leader }))
          Rtl = on effective.Rtl; ContextualSpacing = effective.ContextualSpacing
          KeepNext = direct.KeepNext; KeepLines = direct.KeepLines; WidowControl = effective.WidowControl
          PageBreakBefore = breakAtStart || on effective.PageBreakBefore
          PageBreakBeforeDerived = effective.PageBreakBefore = Some true && (direct.PageBreakBefore <> Some true || direct.PageBreakBeforeDerived)
          PageBreakAtStart = breakAtStart; PageBreakAfter = breakAfter; ColumnBreakAfter = effective.ColumnBreak
          RenderedPageBreakBefore = effective.RenderedPageBreakBefore; BackgroundColor = D.wordColorToCss effective.BackgroundColor
          ShadingPattern = effective.ShadingPattern; ShadingForeground = D.wordColorToCss effective.ShadingForeground
          Borders = direct.ParagraphBorders |> Option.map paraBorders
          ParagraphBorderBefore = effective.ParagraphBorders |> Option.bind (fun borders -> D.paragraphBorderAdvanceToLayoutPx borders.Top)
          ParagraphBorderAfter = effective.ParagraphBorders |> Option.bind (fun borders -> D.paragraphBorderAdvanceToLayoutPx borders.Bottom)
          OutlineLevel = integer direct.OutlineLevel; ExplicitParagraphFormat = true
          ExplicitIndentLeft = onlyWhen direct.IndentLeft.IsSome (pixels effective.IndentLeft)
          ExplicitIndentRight = onlyWhen direct.IndentRight.IsSome (pixels effective.IndentRight)
          ExplicitFirstLineIndent = onlyWhen (direct.FirstLineIndent.IsSome || direct.HangingIndent.IsSome) first
          ExplicitSpacingBefore = onlyWhen (direct.SpacingBefore.IsSome || direct.SpacingBeforeAuto.IsSome) before
          ExplicitSpacingAfter = onlyWhen (direct.SpacingAfter.IsSome || direct.SpacingAfterAuto.IsSome) after
          ExplicitLineSpacing = onlyWhen direct.LineSpacing.IsSome line
          ExplicitContextualSpacing = direct.ContextualSpacing
          ExplicitBackgroundColor = onlyWhen direct.BackgroundColor.IsSome (D.wordColorToCss effective.BackgroundColor)
          ExplicitShadingPattern = direct.ShadingPattern; ExplicitShadingForeground = D.wordColorToCss direct.ShadingForeground
          ExplicitTabs = not direct.TabStops.IsEmpty; HasMarkProperties = direct.ParagraphRunProperties.IsSome
          ExplicitMarkFontSize = direct.ParagraphRunProperties |> Option.bind (fun mark -> mark.FontSize |> Option.map (Some >> D.wordFontSizeToPoints))
          ParagraphMarkFontSize = Some markSize
          MarkRevision = value.MarkRevision |> Option.map (fun (kind, mark) -> revision kind mark)
          DropCapKind = effective.DropCap |> Option.map (fun (kind, _, _) -> kind)
          DropCapLines = effective.DropCap |> Option.map (fun (_, lines, _) -> lines)
          DropCapDistance = effective.DropCap |> Option.bind (fun (_, _, space) -> pixels space)
          ContentControlId = effective.ContentControlId; ContentControlTag = effective.ContentControlTag
          ContentControlAlias = effective.ContentControlAlias; ContentControlXml = effective.ContentControlXml
          ContentControlGroup = effective.ContentControlGroup
          ContentControlAncestors = effective.ContentControlAncestors |> List.map controlLayer
          EnclosingCellControls = effective.EnclosingCellControls |> List.map controlLayer }
    let definition =
        match effective.ListInfo with
        | Some info -> Map.tryFind (info.ListId, max 0 info.Level) numbering
        | None -> effective.Style |> Option.bind (fun id -> Map.tryFind id styleNumbering)
    let listInfo =
        match direct.ListInfo with
        | Some info when info.ListId = "0" ->
            // A direct numId=0 cancels numbering inherited from the paragraph
            // style. It is not the same as an unnumbered paragraph without numPr.
            Some { Style = "none"; Level = 0; Id = Some "0"; SourceId = None
                   Restart = false; FormatEdited = false; NumberFormat = None
                   LevelText = None; Start = None; StartOverrides = []
                   ListLevelIndents = None; Levels = [] }
        | _ ->
            match definition, effective.ListInfo with
            | Some (def: D.NumberingDefinition), _ ->
                let level = levelOf def
                Some { Style = if def.NumberFormat = "bullet" then "bullet" else "numbered"
                       Level = def.Level + 1; Id = Some def.NumId; SourceId = Some def.NumId
                       Restart = false; FormatEdited = false; NumberFormat = Some def.NumberFormat
                       LevelText = def.LevelText; Start = def.Start; StartOverrides = []
                       ListLevelIndents = Some (level.IndentLeft, level.FirstLineIndent)
                       Levels = Map.tryFind def.NumId levels |> Option.defaultValue [] }
            | None, Some info when info.ListId <> "0" ->
                Some { Style = if info.NumberFormat = Some "bullet" then "bullet" else "numbered"
                       Level = info.Level + 1; Id = Some info.ListId; SourceId = Some info.ListId
                       Restart = false; FormatEdited = false; NumberFormat = info.NumberFormat
                       LevelText = info.LevelText; Start = info.Start; StartOverrides = []; ListLevelIndents = None; Levels = [] }
            | _ -> None
    let format =
        match definition with
        | Some def ->
            let level = levelOf def
            let ownNumbering = direct.ListInfo.IsSome
            { format with
                IndentLeft = if format.IndentLeft.IsNone || (ownNumbering && format.ExplicitIndentLeft.IsNone && level.IndentLeft.IsSome) then level.IndentLeft else format.IndentLeft
                FirstLineIndent = if format.FirstLineIndent.IsNone || (ownNumbering && format.ExplicitFirstLineIndent.IsNone && level.FirstLineIndent.IsSome) then level.FirstLineIndent else format.FirstLineIndent
                RightTabPosition = format.RightTabPosition |> Option.orElse (pixels def.RightTabPosition) }
        | None -> format
    { Runs = value.Runs |> List.choose (convertRun styleMap theme images notes effective) |> List.toArray
      HeadingLevel = if inCell then None else effective.Style |> Option.bind ValueText.StyleName.headingSuffix |> Option.map (fun suffix -> integer (Some suffix) |> Option.defaultValue 1 |> max 1 |> min 3)
      Formatting = format; List = listInfo; SectionBreak = effective.SectionBreak
      SectionPageSetup = effective.SectionGeometry |> Option.map sectionSetup; SectionColumns = effective.SectionColumns |> Option.map columns
      SectionVerticalAlign = effective.SectionVerticalAlign; SectionPageNumberStart = effective.SectionPageNumberStart; SectionPageNumberFormat = effective.SectionPageNumberFormat
      HeaderReferences = references effective.HeaderReferences; FooterReferences = references effective.FooterReferences
      Source = source value.Source; SourceParagraph = None; SpliceOriginal = None; SpliceFrom = None
      SpliceLeading = false; SpliceTrailing = false; SpliceJoined = [||]
      SourceIndex = None; SourceUnchanged = false; SourceWhy = None; SourceTextChanges = None }

/// Reuse style and numbering resolution across slices of a progressive open.
/// The callback consumes only the new source blocks and their available media.
let converter (document: D.WordDocument) : D.DocumentContent list -> Map<string, string> -> EditorBlock array =
    let styleMap = document.Styles |> List.map (fun style -> style.Id, style) |> Map.ofList
    let numbering = document.NumberingDefinitions |> List.map (fun item -> (item.NumId, item.Level), item) |> Map.ofList
    let styleNumbering = document.NumberingDefinitions |> List.choose (fun item -> item.ParagraphStyle |> Option.map (fun style -> style, item)) |> Map.ofList
    let levels = document.NumberingDefinitions |> List.groupBy (fun item -> item.NumId) |> List.map (fun (id, items) -> id, items |> List.sortBy (fun item -> item.Level) |> List.map levelOf) |> Map.ofList
    let noteBodies =
        [ yield! document.Footnotes |> List.map (fun note -> ("footnote", note.Id), D.Notes.plainText note)
          yield! document.Endnotes |> List.map (fun note -> ("endnote", note.Id), D.Notes.plainText note) ] |> Map.ofList
    let convert content images =
        let _, defaults = D.resolveStyle styleMap (Some D.Styles.defaultStyleId)
        let defaultSpacing = defaults |> Option.exists D.tableStyleParagraphResetsMissingSpacing
        let para spacing inCell value =
            paragraph styleMap document.Theme images noteBodies numbering styleNumbering levels spacing inCell value
        let rec table (value: D.Table) : EditorTable =
            let style = D.resolveTableStyle styleMap value.StyleId
            let tableStyle = style |> Option.bind (fun style -> style.TableProperties)
            let borders =
                value.Borders
                |> Option.orElseWith (fun () -> value.Properties |> Option.bind (fun properties -> properties.Borders))
                |> Option.orElseWith (fun () -> tableStyle |> Option.bind (fun properties -> properties.Borders))
            let cellSpan (cell: D.TableCell) = cell.Properties |> Option.bind (fun p -> p.ColSpan) |> Option.defaultValue 1 |> max 1
            let columnCount = value.Rows |> List.fold (fun count row -> max count (row.Cells |> List.sumBy cellSpan)) 0
            let rowCount = value.Rows.Length
            let enabled pick = value.Look |> Option.exists pick
            let firstRow index = index = 0 && enabled (fun look -> look.FirstRow)
            let lastRow index = index = rowCount - 1 && enabled (fun look -> look.LastRow)
            let conditionalCell row column span =
                let get pick = style |> Option.bind pick
                None
                |> D.mergeCellProperties (if enabled (fun look -> look.HorizontalBanding) then get (fun style -> if row % 2 = 0 then style.Band1HorizontalTableCellProperties else style.Band2HorizontalTableCellProperties) else None)
                |> D.mergeCellProperties (if enabled (fun look -> look.VerticalBanding) then get (fun style -> if column % 2 = 0 then style.Band1VerticalTableCellProperties else style.Band2VerticalTableCellProperties) else None)
                |> D.mergeCellProperties (if column = 0 && enabled (fun look -> look.FirstColumn) then get (fun style -> style.FirstColumnTableCellProperties) else None)
                |> D.mergeCellProperties (if column + span >= columnCount && enabled (fun look -> look.LastColumn) then get (fun style -> style.LastColumnTableCellProperties) else None)
                |> D.mergeCellProperties (if firstRow row then get (fun style -> style.FirstRowTableCellProperties) else None)
                |> D.mergeCellProperties (if lastRow row then get (fun style -> style.LastRowTableCellProperties) else None)
            let paragraphStyle index =
                let common =
                    style |> Option.bind (fun style -> style.ParagraphProperties)
                    |> D.paragraphPropertiesWithRunProperties (style |> Option.bind (fun style -> style.RunProperties))
                let conditional = if firstRow index then style |> Option.bind (fun style -> style.FirstRowParagraphProperties) else None
                match common, conditional with
                | Some baseValue, Some current -> Some (D.mergeParagraphProperties baseValue current)
                | Some current, None | None, Some current -> Some current
                | None, None -> None
            let preferred amount kind = amount |> Option.map (fun width -> if kind = Some "pct" then width / 50.0 else width)
            let cell index column (value: D.TableCell) : EditorCell =
                let properties = D.mergeCellProperties (conditionalCell index column (cellSpan value)) value.Properties
                let get pick = properties |> Option.bind pick
                let style = paragraphStyle index
                let spacing = style |> Option.exists D.tableStyleParagraphResetsMissingSpacing
                let paragraphs =
                    value.Paragraphs |> List.map (fun original ->
                        let styled =
                            match style with
                            | Some style ->
                                let properties = D.mergeParagraphProperties style original.Properties
                                let properties =
                                    if original.Properties.Style.IsNone && style.Style.IsNone then
                                        { properties with Style = Some D.Styles.defaultStyleId }
                                    else properties
                                { original with Properties = properties }
                            | None -> original
                        Paragraph (para spacing true styled)) |> List.toArray
                let cellBorder =
                    get (fun p -> p.Borders)
                    |> Option.bind (fun borders -> borders.Top |> Option.orElse borders.Left |> Option.orElse borders.Bottom |> Option.orElse borders.Right)
                let width = get (fun p -> p.Width) |> number
                let widthType = get (fun p -> p.WidthType)
                { Paragraphs = paragraphs; NestedTables = value.NestedTables |> List.map table |> List.toArray
                  NestedTableAnchors = List.toArray value.NestedTableAnchors
                  EnclosingControls = value.EnclosingControls |> List.map controlLayer
                  ColSpan = get (fun p -> p.ColSpan) |> Option.defaultValue 1
                  BackgroundColor = get (fun p -> p.Shading) |> Option.bind D.directShadingBackgroundColor
                  Shading = get (fun p -> p.Shading) |> Option.map shading
                  Width = width; PreferredWidth = preferred width widthType; PreferredWidthType = widthType
                  BorderColor = cellBorder |> Option.bind (fun border -> borderColor border.Color)
                  BorderWidth = cellBorder |> Option.bind (fun border -> number border.Size) |> Option.map (fun width -> width / 6.0)
                  Borders = get (fun p -> p.Borders) |> Option.map cellBorders
                  VerticalAlignment = get (fun p -> p.VerticalAlignment); VerticalMerge = get (fun p -> p.VerticalMerge)
                  PaddingTop = get (fun p -> p.PaddingTop) |> number; PaddingRight = get (fun p -> p.PaddingRight) |> number
                  PaddingBottom = get (fun p -> p.PaddingBottom) |> number; PaddingLeft = get (fun p -> p.PaddingLeft) |> number
                  Source = source value.Source; SourceCell = None }
            let rows =
                value.Rows |> List.mapi (fun index row ->
                    let rowStyle = if firstRow index then style |> Option.bind (fun style -> style.FirstRowTableRowProperties) else None
                    let mutable column = 0
                    let cells =
                        row.Cells |> List.map (fun item ->
                            let converted = cell index column item
                            column <- column + cellSpan item
                            converted) |> List.toArray
                    { Cells = cells; IsHeader = row.IsHeader || (rowStyle |> Option.bind (fun style -> style.IsHeader) |> on)
                      Height = number row.Height; HeightRule = row.HeightRule
                      CanSplit = row.CanSplit |> Option.orElseWith (fun () -> rowStyle |> Option.bind (fun style -> style.CanSplit))
                      CellSpacing = number row.CellSpacing; PaddingTop = number row.PaddingTop; PaddingRight = number row.PaddingRight
                      PaddingBottom = number row.PaddingBottom; PaddingLeft = number row.PaddingLeft
                      GridBefore = defaultArg row.GridBefore 0; GridAfter = defaultArg row.GridAfter 0
                      WidthBefore = number row.WidthBefore; WidthAfter = number row.WidthAfter
                      BordersOverride = row.BordersOverride |> Option.map cellBorders
                      Source = source row.Source; SourceRow = None }) |> List.toArray
            let getTable pick = value.Properties |> Option.bind pick |> Option.orElseWith (fun () -> tableStyle |> Option.bind pick)
            { Rows = rows; PreferredWidth = preferred (number value.Width) value.WidthType; PreferredWidthType = value.WidthType
              Alignment = value.Alignment; Indent = number value.Indent; LayoutMode = value.LayoutMode
              GridColumns = value.GridColumns |> List.choose (Some >> number) |> List.toArray
              Borders = borders |> Option.map cellBorders
              BackgroundColor = getTable (fun p -> p.Shading) |> Option.bind D.directShadingBackgroundColor
              Shading = getTable (fun p -> p.Shading) |> Option.map shading; CellSpacing = number value.CellSpacing
              PaddingTop = getTable (fun p -> p.PaddingTop) |> number; PaddingRight = getTable (fun p -> p.PaddingRight) |> number
              PaddingBottom = getTable (fun p -> p.PaddingBottom) |> number; PaddingLeft = getTable (fun p -> p.PaddingLeft) |> number
              Floating = value.Floating; FloatingXml = value.FloatingXml; StyleId = value.StyleId
              ContentControlId = value.ContentControlId; ContentControlTag = value.ContentControlTag
              ContentControlAlias = value.ContentControlAlias; ContentControlXml = value.ContentControlXml
              ContentControlGroup = value.ContentControlGroup
              ContentControlAncestors = value.ContentControlAncestors |> List.map controlLayer
              Look = value.Look |> Option.map (fun look ->
                  { FirstRow = look.FirstRow; LastRow = look.LastRow; FirstColumn = look.FirstColumn
                    LastColumn = look.LastColumn; HorizontalBanding = look.HorizontalBanding; VerticalBanding = look.VerticalBanding })
              Source = noSource; SourceIndex = None; SourceUnchanged = false; SourceWhy = None
              SourceTextChanges = None; SourceTable = None }
        content |> List.map (function
            | D.Para value -> Paragraph (para (defaultSpacing || D.directParagraphFormattingResetsMissingSpacing value.Properties) false value)
            | D.Tbl value -> Table (table value)) |> List.toArray
    fun content images -> serialized (fun () -> convert content images)

let private styleRun theme (value: D.RunProperties) : StyleRun =
    { Bold = value.Bold; Italic = value.Italic; Underline = value.Underline
      UnderlineStyle = value.UnderlineStyle; UnderlineColor = value.UnderlineColor
      AllCaps = value.AllCaps; SmallCaps = value.SmallCaps; Strike = value.Strike; DoubleStrike = value.DoubleStrike
      Hidden = value.Hidden; Position = value.Position |> Option.map (Some >> D.wordPositionToPoints)
      EmphasisMark = value.EmphasisMark; VertAlign = value.VertAlign
      FontSize = value.FontSize |> Option.map (Some >> D.wordFontSizeToPoints); CharacterSpacing = pixels value.CharacterSpacing
      FontFamily = value.FontFamily |> Option.orElseWith (fun () -> TC.resolveThemeFont theme value.ThemeFont)
      Color = TC.normalizeColorWithTheme theme value.Color value.ThemeColor value.ThemeTint value.ThemeShade
      Highlight = value.Highlight }

let private styleParagraph fontSize (value: D.ParagraphProperties) : StyleParagraph =
    let spacing amount auto = if auto = Some true then Some 14.0 else points amount
    { SpacingBefore = spacing value.SpacingBefore value.SpacingBeforeAuto; SpacingAfter = spacing value.SpacingAfter value.SpacingAfterAuto
      ContextualSpacing = value.ContextualSpacing
      LineSpacing = D.wordLineSpacingToMultiple (defaultArg fontSize (10.0 * 96.0 / 72.0)) value.LineSpacing value.LineRule
      LineSpacingRule = value.LineRule; Alignment = value.Alignment
      IndentLeft = pixels value.IndentLeft; IndentRight = pixels value.IndentRight
      FirstLineIndent = pixels value.FirstLineIndent |> Option.orElse (pixels value.HangingIndent |> Option.map (~-))
      BackgroundColor = D.wordColorToCss value.BackgroundColor; OutlineLevel = integer value.OutlineLevel
      KeepNext = value.KeepNext; KeepLines = value.KeepLines; PageBreakBefore = value.PageBreakBefore; WidowControl = value.WidowControl }

let private style theme (value: D.StyleDefinition) : Style =
    let run = value.RunProperties |> Option.map (styleRun theme)
    { Id = value.Id; Name = value.Name; StyleType = value.StyleType; IsDefault = value.IsDefault
      BasedOn = value.BasedOn; Next = value.Next; Linked = value.Linked; IsQuickStyle = Some value.IsQuickStyle
      Aliases = value.Aliases; UiPriority = integer value.UiPriority; SemiHidden = Some value.SemiHidden; UnhideWhenUsed = Some value.UnhideWhenUsed
      Run = run; Paragraph = value.ParagraphProperties |> Option.map (styleParagraph (run |> Option.bind (fun run -> run.FontSize))) }

let private visibleShape (shape: D.Shape) =
    not shape.Content.IsEmpty || shape.Preset.IsSome || not shape.CustomPath.IsEmpty || not shape.Children.IsEmpty

/// The shapes of a part. `theme`: the document's, for the theme colours of their fills and lines.
let shapes (theme: D.Theme option) (convert: D.DocumentContent list -> Map<string, string> -> EditorBlock array) images imageOf values =
    let colors = theme |> Option.map (fun theme -> theme.Colors) |> Option.defaultValue Map.empty
    let color (value: string option) = value |> Option.bind (D.Shapes.resolveColor colors)
    let rec shape (value: D.Shape) : EditorTextBox =
        { Paragraphs = convert (value.Content |> List.map D.Para) images
          Children = value.Children |> List.map shape |> List.toArray
          ShapeId = value.Id; ShapeKind = value.Type
          Name = value.Name; Geometry = if value.CustomPath.IsEmpty then defaultArg value.Preset "rect" else "custom"
          CustomPath = value.CustomPath |> List.map (fun part -> part.Op, List.toArray part.Pts)
          CustomWidth = defaultArg value.CustomW 0.0; CustomHeight = defaultArg value.CustomH 0.0
          X = value.Position |> Option.map fst; Y = value.Position |> Option.map snd
          PositionHRelativeFrom = value.PositionHRelativeFrom; PositionVRelativeFrom = value.PositionVRelativeFrom
          PositionHAlign = value.PositionHAlign; PositionVAlign = value.PositionVAlign
          Width = defaultArg value.Width 200.0; Height = defaultArg value.Height 100.0
          FillColor = color value.FillColor; BorderColor = color value.BorderColor; NoFill = value.NoFill; NoLine = value.NoLine
          BorderWidth = value.BorderWidth; BorderDash = value.LineDash
          Gradient = value.Gradient |> Option.map (fun gradient -> gradient.Stops |> List.choose (fun stop -> color (Some stop.Color) |> Option.map (fun c -> stop.Offset, c))) |> Option.defaultValue []
          GradientAngle = value.Gradient |> Option.map (fun gradient -> gradient.Angle) |> Option.defaultValue 0.0
          GradientRadial = value.Gradient |> Option.exists (fun gradient -> gradient.Radial)
          Rotation = value.Rotation; FlipH = value.FlipH; FlipV = value.FlipV; Inline = value.Inline; Wrap = value.Wrap
          BehindDoc = value.BehindDoc; AllowOverlap = value.AllowOverlap; LayoutInCell = value.LayoutInCell
          DistanceTop = value.DistanceTop; DistanceBottom = value.DistanceBottom; DistanceLeft = value.DistanceLeft; DistanceRight = value.DistanceRight
          RelativeHeight = value.RelativeHeight; AnchorBlock = value.AnchorBlock
          InsetLeft = value.InsetLeft; InsetTop = value.InsetTop; InsetRight = value.InsetRight; InsetBottom = value.InsetBottom
          VerticalAnchor = value.VerticalAnchor; AutoFit = value.AutoFit; NoTextWrap = value.NoTextWrap; FitText = value.FitText
          Opacity = defaultArg value.Opacity 1.0; ChildSpace = value.ChildSpace; ImageSource = value.ImageRelationshipId |> Option.bind imageOf
          SourceXml = value.SourceXml; SourceMoved = None; SourceIndex = None; SourceUnchanged = false; SourceWhy = None; SourceTextChanges = None }
    values |> List.filter visibleShape |> List.map (shape >> TextBox) |> List.toArray

/// Import one parsed package with its source provenance; no JS object
/// construction or property lookup participates in this conversion.
let ofDocument (document: D.WordDocument) : EditorPackage =
    let convert = converter document
    let shapeItems = shapes document.Theme convert document.Images (fun id -> Map.tryFind id document.Images) document.Shapes
    let headers (parts: D.HeaderFooter list) =
        parts |> List.map (fun part ->
            let partShapes = part.RelationshipId |> Option.bind (fun id -> Map.tryFind id document.PartShapes) |> Option.defaultValue []
            let imageOf id = Map.tryFind (D.partImageKey (defaultArg part.RelationshipId "") id) document.Images
            { RelationshipId = part.RelationshipId; Kind = part.Type
              Blocks = Array.append (convert part.Blocks document.Images) (shapes document.Theme convert document.Images imageOf partShapes)
              SourceStoryCounts = [||]; SourceUnchanged = false }) |> List.toArray
    let notes (items: D.Note list) =
        items |> List.map (fun note ->
            { Id = note.Id; Kind = note.Kind; HasMark = Some note.HasMark; MarkStyle = note.MarkStyle
              Text = Some (D.Notes.plainText note); Blocks = convert note.Blocks document.Images
              SourceStoryCounts = [||]; Unchanged = false }) |> List.toArray
    // The source indexes must cover exactly the shapes that are emitted.
    let provenance =
        { document with Shapes = document.Shapes |> List.filter visibleShape
                        PartShapes = document.PartShapes |> Map.map (fun _ items -> List.filter visibleShape items) }
        |> preservedOfDocument
    { Blocks = Array.append (convert document.Content document.Images) shapeItems
      Styles = document.Styles |> List.map (style document.Theme) |> List.toArray
      HeaderParts = headers document.Headers; FooterParts = headers document.Footers
      Footnotes = notes document.Footnotes; Endnotes = notes document.Endnotes
      Comments = document.Comments |> List.map (fun comment ->
          { Id = comment.Id; Author = comment.Author; Date = comment.Date; Initials = comment.Initials; Text = Some comment.Text
            Blocks = convert comment.Blocks document.Images; SourceStoryCounts = [||]; Unchanged = false }) |> List.toArray
      PageSetup = document.PageSetup |> Option.map pageSetup; OriginalPageSetup = None
      Theme = document.Theme
      Preserved = Some provenance; SourceStoryCounts = [||]; ChangedStyles = None; ChangedListLevels = [||]
      PageSetupUnchanged = false; ListsUnchanged = false; FootnotesUnchanged = false; EndnotesUnchanged = false
      CommentsUnchanged = false; Unchanged = false; RegenerateAll = false }

#if !FABLE_COMPILER
let loadBytesNative (bytes: byte[]) : EditorPackage option =
    D.loadBytesNative bytes |> Option.map (fun document ->
        let package = ofDocument document
        { package with Preserved = package.Preserved |> Option.map (fun preserved -> { preserved with OpenedBytes = Some (Array.copy bytes) }) })
#endif
