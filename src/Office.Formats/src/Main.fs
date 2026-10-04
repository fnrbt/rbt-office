module Main

open Fable.Core
open Fable.Core.JsInterop
open ZipInterop
open System.Collections.Generic

// Public API for DOCX reading
module Word =
    /// Load a DOCX file from array buffer or blob
    let load = DocxReader.load

    /// Extract all text from a Word document
    let extractText = DocxReader.extractText

    /// Get document content (paragraphs and tables)
    let getContent (doc: DocxReader.WordDocument) = doc.Content |> List.toArray

    /// Get all paragraphs from document
    let paragraphs (doc: DocxReader.WordDocument) =
        doc.Content
        |> List.choose (function
            | DocxReader.Para p -> Some p
            | _ -> None)

    /// Get all tables from document
    let tables (doc: DocxReader.WordDocument) =
        doc.Content
        |> List.choose (function
            | DocxReader.Tbl t -> Some t
            | _ -> None)

    /// Get document styles
    let styles (doc: DocxReader.WordDocument) = doc.Styles

    /// Get document headers
    let headers (doc: DocxReader.WordDocument) = doc.Headers

    /// Get document footers
    let footers (doc: DocxReader.WordDocument) = doc.Footers

    /// Get document properties (title, author, etc.)
    let properties (doc: DocxReader.WordDocument) = doc.Properties

    /// Get all hyperlinks from document
    let hyperlinks (doc: DocxReader.WordDocument) =
        let extractFromPara (para: DocxReader.Paragraph) =
            para.Runs
            |> List.choose (fun run ->
                match run.Content with
                | DocxReader.Link link -> Some link
                | _ -> None)

        doc.Content
        |> List.collect (function
            | DocxReader.Para p -> extractFromPara p
            | DocxReader.Tbl t ->
                t.Rows
                |> List.collect (fun row ->
                    row.Cells
                    |> List.collect (fun cell ->
                        cell.Paragraphs
                        |> List.collect extractFromPara))
        )

    /// Get all images from document
    let images (doc: DocxReader.WordDocument) =
        let extractFromPara (para: DocxReader.Paragraph) =
            para.Runs
            |> List.choose (fun run ->
                match run.Content with
                | DocxReader.Image img -> Some img
                | _ -> None)

        doc.Content
        |> List.collect (function
            | DocxReader.Para p -> extractFromPara p
            | DocxReader.Tbl t ->
                t.Rows
                |> List.collect (fun row ->
                    row.Cells
                    |> List.collect (fun cell ->
                        cell.Paragraphs
                        |> List.collect extractFromPara))
        )

    /// Get all comments from document
    let comments (doc: DocxReader.WordDocument) = doc.Comments

    /// Get all revisions/track changes from document
    let revisions (doc: DocxReader.WordDocument) = doc.Revisions

    /// Get footnotes from document
    let footnotes (doc: DocxReader.WordDocument) = doc.Footnotes

    /// Get endnotes from document
    let endnotes (doc: DocxReader.WordDocument) = doc.Endnotes

    /// Get shapes and text boxes from document
    let shapes (doc: DocxReader.WordDocument) = doc.Shapes

// Public API for Document Processors
module Processors =
    let private wordColorToCss (color: string option) =
        color
        |> Option.bind (fun value ->
            if value = "auto" || value = "" then None
            elif value.StartsWith("#") then Some value
            elif value.Length = 8 then Some ("#" + value.Substring(2).ToUpper())
            else Some ("#" + value.ToUpper()))

    let private wordFontSizeToCssPixels (size: string option) =
        size
        |> Option.bind (fun value ->
            match System.Double.TryParse(value) with
            | true, parsed -> Some ((parsed / 2.0) * 96.0 / 72.0)
            | false, _ -> None)

    let private wordTwipsToCssPixels (value: string option) =
        value
        |> Option.bind (fun text ->
            match System.Double.TryParse(text) with
            | true, parsed -> Some (parsed / 15.0)
            | false, _ -> None)

    let private wordCharacterSpacingToCssPixels = wordTwipsToCssPixels

    let private wordTwipsToPoints (value: string option) =
        value
        |> Option.bind (fun text ->
            match System.Double.TryParse(text) with
            | true, parsed -> Some (parsed / 20.0)
            | false, _ -> None)

    let private wordAutoSpacingPoints = 14.0

    let private wordStyleSpacing value auto =
        if auto = Some true then Some wordAutoSpacingPoints else wordTwipsToPoints value

    let private wordLineSpacingValue (fontSize: float option) (value: string option) (rule: string option) =
        value
        |> Option.bind (fun text ->
            match System.Double.TryParse(text) with
            | true, parsed ->
                match rule with
                | Some "exact"
                | Some "atLeast" -> Some (parsed / 15.0)
                | _ -> Some (parsed / 240.0)
            | false, _ -> None)

    let private wordInt (value: string option) =
        value
        |> Option.bind (fun text ->
            match System.Int32.TryParse(text) with
            | true, parsed -> Some parsed
            | false, _ -> None)

    let private resolveThemeFont (theme: DocxReader.Theme option) (themeFont: string option) =
        match theme, themeFont with
        | Some theme, Some value when value.StartsWith("major") -> theme.MajorFont
        | Some theme, Some value when value.StartsWith("minor") -> theme.MinorFont
        | _ -> None

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

    let private transformThemeColor (themeTint: string option) (themeShade: string option) (color: string) =
        if color.Length <> 7 || not (color.StartsWith("#")) then color
        else
            let transformed =
                [ color.Substring(1, 2); color.Substring(3, 2); color.Substring(5, 2) ]
                |> List.map (fun part ->
                    let channel = hexByte part |> Option.defaultValue 0
                    let shaded = themeShade |> Option.bind hexByte |> Option.map (fun shade -> int (float channel * (float shade / 255.0)) |> clampByte) |> Option.defaultValue channel
                    themeTint |> Option.bind hexByte |> Option.map (fun tint -> int (float shaded + (float (255 - shaded) * (float tint / 255.0))) |> clampByte) |> Option.defaultValue shaded)
            sprintf "#%02X%02X%02X" transformed.[0] transformed.[1] transformed.[2]

    let private resolveThemeColor (theme: DocxReader.Theme option) (themeColor: string option) (themeTint: string option) (themeShade: string option) =
        match theme, themeColor with
        | Some theme, Some key -> Map.tryFind key theme.Colors |> Option.map (transformThemeColor themeTint themeShade)
        | _ -> None

    let private styleToObj (theme: DocxReader.Theme option) (style: DocxReader.StyleDefinition) =
        createObj [
            "StyleId" ==> style.Id
            "Name" ==> style.Name
            "StyleType" ==> style.StyleType
            "IsDefaultStyle" ==> style.IsDefault
            "BasedOn" ==> style.BasedOn
            "NextStyle" ==> style.Next
            "LinkedStyle" ==> style.Linked
            "IsQuickStyle" ==> style.IsQuickStyle
            "StyleAliases" ==> style.Aliases
            "StyleUiPriority" ==> (style.UiPriority |> Option.bind (fun value -> match System.Int32.TryParse(value) with | true, parsed -> Some parsed | false, _ -> None))
            "StyleSemiHidden" ==> style.SemiHidden
            "StyleUnhideWhenUsed" ==> style.UnhideWhenUsed
            "StyleBold" ==> (style.RunProperties |> Option.bind (fun props -> props.Bold))
            "StyleItalic" ==> (style.RunProperties |> Option.bind (fun props -> props.Italic))
            "StyleUnderline" ==> (style.RunProperties |> Option.bind (fun props -> props.Underline))
            "StyleAllCaps" ==> (style.RunProperties |> Option.bind (fun props -> props.AllCaps))
            "StyleSmallCaps" ==> (style.RunProperties |> Option.bind (fun props -> props.SmallCaps))
            "StyleStrike" ==> (style.RunProperties |> Option.bind (fun props -> props.Strike))
            "StyleVertAlign" ==> (style.RunProperties |> Option.bind (fun props -> props.VertAlign))
            "StyleFontSize" ==> (style.RunProperties |> Option.bind (fun props -> wordFontSizeToCssPixels props.FontSize))
            "StyleCharacterSpacing" ==> (style.RunProperties |> Option.bind (fun props -> wordCharacterSpacingToCssPixels props.CharacterSpacing))
            "StyleFontFamily" ==> (style.RunProperties |> Option.bind (fun props -> props.FontFamily |> Option.orElse (resolveThemeFont theme props.ThemeFont)))
            "StyleColor" ==> (style.RunProperties |> Option.bind (fun props -> wordColorToCss props.Color |> Option.orElse (resolveThemeColor theme props.ThemeColor props.ThemeTint props.ThemeShade)))
            "StyleHighlight" ==> (style.RunProperties |> Option.bind (fun props -> props.Highlight))
            "StyleSpacingBefore" ==> (style.ParagraphProperties |> Option.bind (fun props -> wordStyleSpacing props.SpacingBefore props.SpacingBeforeAuto))
            "StyleSpacingAfter" ==> (style.ParagraphProperties |> Option.bind (fun props -> wordStyleSpacing props.SpacingAfter props.SpacingAfterAuto))
            "StyleLineSpacing" ==> (style.ParagraphProperties |> Option.bind (fun props -> wordLineSpacingValue (style.RunProperties |> Option.bind (fun run -> wordFontSizeToCssPixels run.FontSize)) props.LineSpacing props.LineRule))
            "StyleLineRule" ==> (style.ParagraphProperties |> Option.bind (fun props -> props.LineRule))
            "StyleAlignment" ==> (style.ParagraphProperties |> Option.bind (fun props -> props.Alignment))
            "StyleIndentLeft" ==> (style.ParagraphProperties |> Option.bind (fun props -> wordTwipsToCssPixels props.IndentLeft))
            "StyleIndentRight" ==> (style.ParagraphProperties |> Option.bind (fun props -> wordTwipsToCssPixels props.IndentRight))
            // A hanging indent is a negative first-line indent, as on a paragraph.
            "StyleFirstLineIndent" ==> (style.ParagraphProperties
                                        |> Option.bind (fun props ->
                                            wordTwipsToCssPixels props.FirstLineIndent
                                            |> Option.orElse (wordTwipsToCssPixels props.HangingIndent |> Option.map (fun hanging -> -hanging))))
            "StyleBackgroundColor" ==> (style.ParagraphProperties |> Option.bind (fun props -> wordColorToCss props.BackgroundColor))
            "StyleOutlineLevel" ==> (style.ParagraphProperties |> Option.bind (fun props -> wordInt props.OutlineLevel))
            "StyleKeepNext" ==> (style.ParagraphProperties |> Option.bind (fun props -> props.KeepNext))
            "StyleKeepLines" ==> (style.ParagraphProperties |> Option.bind (fun props -> props.KeepLines))
            "StylePageBreakBefore" ==> (style.ParagraphProperties |> Option.bind (fun props -> props.PageBreakBefore))
            "StyleWidowControl" ==> (style.ParagraphProperties |> Option.bind (fun props -> props.WidowControl))
        ]

    let private editorParagraphToObj (para: DocumentProcessors.EditorProcessor.EditorParagraph) =
        createObj [
            "Runs" ==> (para.Runs |> List.toArray)
            "Alignment" ==> para.Alignment
            "IndentLeft" ==> para.IndentLeft
            "IndentRight" ==> para.IndentRight
            "FirstLineIndent" ==> para.FirstLineIndent
            "RightTabPosition" ==> para.RightTabPosition
            "RightTabLeader" ==> para.RightTabLeader
            "SpacingBefore" ==> para.SpacingBefore
            "SpacingAfter" ==> para.SpacingAfter
            "LineSpacing" ==> para.LineSpacing
        ]

    /// Build editor "TextBox" objects (incl. preset/custom geometry + gradient)
    /// for every drawing shape in the document.
    let private shapeObjectsFor (doc: DocxReader.WordDocument) : obj array =
        let styleMap = DocumentProcessors.StyleResolution.buildStyleMapWithTheme doc.Theme doc.Styles
        let comments = doc.Comments |> List.map (fun comment -> comment.Id, comment) |> Map.ofList
        let noteText (note: DocxReader.Note) =
            note.Content
            |> List.collect (fun para -> para.Runs)
            |> List.map (fun run -> DocumentProcessors.EditorProcessor.extractRunText Map.empty run.Content)
            |> String.concat ""
        let notes =
            (doc.Footnotes |> List.map (fun note -> ("footnote", note.Id), noteText note)) @
            (doc.Endnotes |> List.map (fun note -> ("endnote", note.Id), noteText note))
            |> Map.ofList
        doc.Shapes
        |> List.filter (fun shape -> not shape.Content.IsEmpty || shape.Preset.IsSome || not shape.CustomPath.IsEmpty)
        |> List.map (fun shape ->
            let paragraphs =
                shape.Content
                |> List.map (DocumentProcessors.EditorProcessor.toEditorParagraph doc.Theme styleMap comments notes >> editorParagraphToObj)
                |> List.toArray
            let geometry =
                if not shape.CustomPath.IsEmpty then "custom"
                else shape.Preset |> Option.defaultValue "rect"
            let customPath =
                shape.CustomPath
                |> List.map (fun seg -> createObj [ "Op" ==> seg.Op; "P" ==> (seg.Pts |> List.toArray) ])
                |> List.toArray
            let gradient =
                shape.Gradient
                |> Option.map (fun g ->
                    createObj [
                        "Stops" ==> (g.Stops |> List.map (fun s -> createObj [ "Offset" ==> s.Offset; "Color" ==> s.Color ]) |> List.toArray)
                        "Angle" ==> g.Angle
                        "Radial" ==> g.Radial ])
                |> Option.toObj
            createObj [
                "Type" ==> "TextBox"
                "ShapeId" ==> shape.Id
                "ShapeKind" ==> shape.Type
                "Width" ==> shape.Width
                "Height" ==> shape.Height
                "X" ==> (shape.Position |> Option.map fst)
                "Y" ==> (shape.Position |> Option.map snd)
                "FillColor" ==> shape.FillColor
                "BorderColor" ==> shape.BorderColor
                "BorderWidth" ==> shape.BorderWidth
                "Geometry" ==> geometry
                "CustomPath" ==> customPath
                "CustomW" ==> (shape.CustomW |> Option.defaultValue 0.0)
                "CustomH" ==> (shape.CustomH |> Option.defaultValue 0.0)
                "Gradient" ==> gradient
                "Rotation" ==> shape.Rotation
                "AnchorBlock" ==> shape.AnchorBlock
                "Inline" ==> shape.Inline
                "Wrap" ==> shape.Wrap
                "PositionHRelativeFrom" ==> shape.PositionHRelativeFrom
                "PositionVRelativeFrom" ==> shape.PositionVRelativeFrom
                "PositionHAlign" ==> shape.PositionHAlign
                "PositionVAlign" ==> shape.PositionVAlign
                "BehindDoc" ==> shape.BehindDoc
                "AllowOverlap" ==> shape.AllowOverlap
                "LayoutInCell" ==> shape.LayoutInCell
                "RelativeHeight" ==> shape.RelativeHeight
                "DistanceTop" ==> shape.DistanceTop
                "DistanceBottom" ==> shape.DistanceBottom
                "DistanceLeft" ==> shape.DistanceLeft
                "DistanceRight" ==> shape.DistanceRight
                "NoFill" ==> shape.NoFill
                "NoLine" ==> shape.NoLine
                "BorderDash" ==> shape.LineDash
                "Name" ==> shape.Name
                "Paragraphs" ==> paragraphs
            ])
        |> List.toArray

    let forEditorPackage (fileData: obj) : Fable.Core.JS.Promise<obj> =
        promise {
            let! archive = loadArchive fileData
            let! docOpt = DocxReader.parseDocument archive

            match docOpt with
            | Some doc ->
                let paragraphs =
                    DocxReader.toEditorParagraphsWithStylesThemeAndNumbering doc.Theme doc.Content doc.Images doc.Styles doc.NumberingDefinitions doc.Footnotes doc.Endnotes
                let paragraphs = Array.append paragraphs (shapeObjectsFor doc)
                return createObj [
                    "Paragraphs" ==> paragraphs
                    "Styles" ==> (doc.Styles |> List.map (styleToObj doc.Theme) |> List.toArray)
                    "HeaderParts" ==> DocxReader.toEditorHeaderFootersWithStylesAndTheme doc.Theme doc.Headers doc.Images doc.Styles
                    "FooterParts" ==> DocxReader.toEditorHeaderFootersWithStylesAndTheme doc.Theme doc.Footers doc.Images doc.Styles
                    "PageSetup" ==> DocxReader.toEditorPageSetup doc
                    "Footnotes" ==> DocxReader.toEditorNotes doc.Theme doc.Images doc.Styles doc.NumberingDefinitions doc.Footnotes
                    "Endnotes" ==> DocxReader.toEditorNotes doc.Theme doc.Images doc.Styles doc.NumberingDefinitions doc.Endnotes
                    "Comments" ==> DocxReader.toEditorComments doc.Theme doc.Images doc.Styles doc.NumberingDefinitions doc.Comments
                    "Preserved" ==> DocxReader.toEditorPreserved doc.Preserved
                    "Theme" ==> (doc.Theme |> Option.map (fun theme -> createObj [ "MajorFont" ==> theme.MajorFont; "MinorFont" ==> theme.MinorFont; "Colors" ==> createObj (theme.Colors |> Map.toList |> List.map (fun (key, value) -> key ==> value)) ]))
                ]
            | None ->
                return createObj [
                    "Paragraphs" ==> [||]
                    "Styles" ==> [||]
                ]
        }

// Public API for Word Writing
module WordWriter =
    /// Write paragraphs to DOCX (Phase 1: plain text only)
    let write (paragraphs: string list) = DocxWriter.writeMinimalDocx paragraphs

    /// Write from EditorFormat with formatting (Phase 2)
    let writeFromEditor (paragraphs: obj array) =
        DocumentWriter.writeFromEditor paragraphs

    let writeFromEditorPackage (package: obj) =
        DocumentWriter.writeFromEditorPackage package

    /// Create document builder (Phase 3)
    let createDocument () = DocumentBuilder.create ()

    /// Add paragraph to builder
    let addParagraph = DocumentBuilder.addParagraph

    /// Add paragraph with runs to builder
    let addParagraphWithRuns = DocumentBuilder.addParagraphWithRuns

    /// Save document
    let saveDocument = DocumentBuilder.save

    /// Helper to create a run
    let createRun = DocumentBuilder.createRun

    /// Create a simple table from 2D array (Phase 4)
    let createSimpleTable = TableWriter.createSimpleTable

    /// Create a table with borders (Phase 4)
    let createTableWithBorders = TableWriter.createTableWithBorders

    /// Generate table XML (Phase 4)
    let generateTableXml = TableWriter.generateTable

// Public API for Excel Writing
module ExcelWriter =
    /// Write 2D array to XLSX (Phase 1)
    let write (sheets: XlsxWriter.SheetData list) = XlsxWriter.writeMinimalXlsx sheets

    /// Write formatted XLSX (Phase 2)
    let writeFormatted (sheets: XlsxWriter.FormattedSheet list) = XlsxWriter.writeFormattedXlsx sheets

    /// Create workbook builder (Phase 3)
    let createWorkbook () = WorkbookBuilder.create ()

    /// Add sheet to workbook
    let addSheet = WorkbookBuilder.addSheet

    /// Add simple sheet to workbook
    let addSimpleSheet = WorkbookBuilder.addSimpleSheet

    /// Save workbook
    let saveWorkbook = WorkbookBuilder.saveWorkbook

    /// Add row to sheet builder
    let addRow = WorkbookBuilder.addRow

    /// Add simple row to sheet builder
    let addSimpleRow = WorkbookBuilder.addSimpleRow

    /// Helper to create a formatted cell
    let createCell = WorkbookBuilder.createCell

    /// Helper to create a header cell
    let createHeaderCell = WorkbookBuilder.createHeaderCell

    /// Helper to create a numeric cell with format
    let createNumericCell = WorkbookBuilder.createNumericCell

    /// Helper to create a formula cell (Phase 4)
    let createFormulaCell = WorkbookBuilder.createFormulaCell

    /// Add merged range to sheet builder (Phase 4)
    let addMergedRange = WorkbookBuilder.addMergedRange

    /// Create merged cell range (Phase 4)
    let createMergedRange = XlsxWriter.createMergedRange

// Public API for XLSX reading
module Excel =
    type private CellRangeBounds = {
        RowMin: int
        RowMax: int
        ColMin: int
        ColMax: int
    }

    let private columnIndex (letters: string) =
        let mutable acc = 0
        for ch in letters do
            if System.Char.IsLetter ch then
                acc <- acc * 26 + (int (System.Char.ToUpperInvariant ch) - int 'A' + 1)
        acc - 1

    let private refToCoord (ref: string) =
        let mutable index = 0
        let mutable col = 0
        let mutable row = 0
        let mutable hasCol = false
        let mutable hasRow = false
        let mutable inRow = false
        let mutable valid = not (System.String.IsNullOrWhiteSpace ref)
        while valid && index < ref.Length do
            let ch = ref.[index]
            if ch = '$' then
                ()
            elif System.Char.IsLetter ch && not inRow then
                hasCol <- true
                let digit = int (System.Char.ToUpperInvariant ch) - int 'A' + 1
                if digit < 1 || digit > 26 || col > (System.Int32.MaxValue - digit) / 26 then
                    valid <- false
                else
                    col <- col * 26 + digit
            elif System.Char.IsDigit ch then
                inRow <- true
                hasRow <- true
                let digit = int ch - int '0'
                if row > (System.Int32.MaxValue - digit) / 10 then
                    valid <- false
                else
                    row <- row * 10 + digit
            else
                valid <- false
            index <- index + 1
        if valid && hasCol && hasRow && row > 0 && col > 0 then Some (row - 1, col - 1)
        else None

    let private tryParseRangeBounds (range: string) =
        let parts = range.Split(':')
        match parts.Length with
        | 1 ->
            refToCoord parts.[0]
            |> Option.map (fun (r, c) -> { RowMin = r; RowMax = r; ColMin = c; ColMax = c })
        | 2 ->
            match refToCoord parts.[0], refToCoord parts.[1] with
            | Some (r1, c1), Some (r2, c2) ->
                Some {
                    RowMin = min r1 r2
                    RowMax = max r1 r2
                    ColMin = min c1 c2
                    ColMax = max c1 c2
                }
            | _ -> None
        | _ -> None

    let private rangeBoundsContains row col range =
        row >= range.RowMin
        && row <= range.RowMax
        && col >= range.ColMin
        && col <= range.ColMax

    let private rangeContains (row: int) (col: int) (range: string) =
        tryParseRangeBounds range
        |> Option.exists (rangeBoundsContains row col)

    let private rangeListContains row col (ranges: string) =
        ranges.Split([| ' ' |], System.StringSplitOptions.RemoveEmptyEntries)
        |> Array.exists (rangeContains row col)

    let private listDropdownRangeBounds (sheet: XlsxReader.Worksheet) =
        let ranges = ResizeArray<CellRangeBounds>()
        for validation in sheet.DataValidations do
            if validation.ValidationType = "list" then
                for range in validation.Range.Split([| ' ' |], System.StringSplitOptions.RemoveEmptyEntries) do
                    match tryParseRangeBounds range with
                    | Some bounds -> ranges.Add bounds
                    | None -> ()
        ranges.ToArray()

    let private rangeBoundsArrayContains row col (ranges: CellRangeBounds array) =
        let mutable found = false
        let mutable index = 0
        while not found && index < ranges.Length do
            if rangeBoundsContains row col ranges.[index] then
                found <- true
            index <- index + 1
        found

    let private optString (value: string option) =
        value |> Option.defaultValue null

    let private objField name (value: obj) =
#if FABLE_COMPILER
        name ==> value
#else
        name, value
#endif

    let private objFromFields fields : obj =
#if FABLE_COMPILER
        createObj fields
#else
        let result = Dictionary<string, obj>()
        for name, value in fields do
            result.[name] <- value
        box result
#endif

    let private addField (fields: ResizeArray<_>) name value =
        fields.Add(objField name value)

    let private addFieldWhen (fields: ResizeArray<_>) condition name value =
        if condition then addField fields name value

    let private addFieldOption (fields: ResizeArray<_>) name value =
        match value with
        | Some v -> addField fields name (box v)
        | None -> ()

    let private addFieldObjOption (fields: ResizeArray<_>) name value =
        match value with
        | Some v -> addField fields name v
        | None -> ()

    let private cellRefText (cellRef: XlsxReader.CellReference) =
        sprintf "%s%d" cellRef.Column cellRef.Row

    let private optFloat (value: float option) : obj =
        match value with
        | Some v -> box v
        | None -> null

    let private borderObj (side: XlsxReader.CellBorderSide option) : obj =
        match side with
        | Some s ->
            objFromFields [
                objField "style" s.Style
                objField "color" (optString s.Color)
            ]
        | None -> null

    let private cssColor (value: string option) : obj =
        match value with
        | Some raw ->
            let trimmed = raw.TrimStart('#')
            let cleaned = if trimmed.Length = 8 then trimmed.Substring(2) else trimmed
            if cleaned.Length = 6 then box ("#" + cleaned) else null
        | None -> null

    let private richRunObj (text: string, props: XlsxReader.RunProps) =
        let fields = ResizeArray()
        addField fields "text" text
        addFieldWhen fields props.Bold "bold" true
        addFieldWhen fields props.Italic "italic" true
        addFieldWhen fields props.Underline "underline" true
        addFieldWhen fields props.UnderlineDouble "underlineDouble" true
        addFieldWhen fields props.Strikethrough "strikethrough" true
        let color = cssColor props.Color
        if not (isNull color) then
            addField fields "color" color
        addFieldOption fields "fontName" props.FontName
        addFieldObjOption fields "fontSize" (props.FontSize |> Option.map box)
        match props.VertAlign with
        | XlsxReader.VertAlignSuperscript -> addField fields "vertAlign" "superscript"
        | XlsxReader.VertAlignSubscript -> addField fields "vertAlign" "subscript"
        | XlsxReader.VertAlignBaseline -> ()
        objFromFields fields

    let private richRunsForCell (workbook: XlsxReader.Workbook) (cell: XlsxReader.Cell) =
        let lift runs =
            if List.length runs <= 1 then [||]
            else runs |> List.map richRunObj |> List.toArray
        if not (List.isEmpty cell.InlineRuns) then
            lift cell.InlineRuns
        else
            match cell.Value with
            | XlsxReader.SharedString idx
                when idx >= 0
                     && idx < workbook.SharedStringRuns.Length
                     && not (List.isEmpty workbook.SharedStringRuns.[idx]) ->
                lift workbook.SharedStringRuns.[idx]
            | _ -> [||]

    let private richRunsObjForCell (workbook: XlsxReader.Workbook) (cell: XlsxReader.Cell) =
        let lift runs =
            match runs with
            | _ :: _ :: _ -> box (runs |> List.map richRunObj |> List.toArray)
            | _ -> null
        if not (List.isEmpty cell.InlineRuns) then
            lift cell.InlineRuns
        else
            match cell.Value with
            | XlsxReader.SharedString idx
                when idx >= 0
                     && idx < workbook.SharedStringRuns.Length
                     && not (List.isEmpty workbook.SharedStringRuns.[idx]) ->
                lift workbook.SharedStringRuns.[idx]
            | _ -> null

    /// Formula text (no leading "="), or null for a plain value cell.
    let private cellFormula (cell: XlsxReader.Cell) =
        match cell.Formula with
        | Some f -> f
        | None -> null

    /// The cell's value — for a formula cell, the result cached in the file.
    /// null for empty cells, error cells (see `cellErrorObj`) and formulas
    /// the file carries no cached result for.
    let private cellValueObj (workbook: XlsxReader.Workbook) (cell: XlsxReader.Cell) : obj =
        match XlsxReader.resolveValue workbook cell.Value with
        | XlsxReader.String s -> box s
        | XlsxReader.Number n -> box n
        | XlsxReader.Boolean b -> box b
        | XlsxReader.Date d -> box d
        | _ -> null

    /// Error code without the leading "#" ("DIV/0!", "N/A") for an error
    /// cell or a formula whose cached result is an error; null otherwise.
    let private cellErrorObj (cell: XlsxReader.Cell) : obj =
        match cell.Value with
        | XlsxReader.Error e -> box e
        | _ -> null

    let private listDropdownAt row col (sheet: XlsxReader.Worksheet) =
        sheet.DataValidations
        |> List.exists (fun dv ->
            dv.ValidationType = "list"
            && rangeListContains row col dv.Range)

    let private commentAtCoord row col (sheet: XlsxReader.Worksheet) =
        sheet.Comments
                |> List.tryFind (fun c ->
            c.CellRef.Row - 1 = row
            && columnIndex c.CellRef.Column = col)

    let private addStyleFields fields (style: XlsxReader.CellStyle option) =
        match style with
        | Some s ->
            addFieldWhen fields s.Bold "bold" true
            addFieldWhen fields s.Italic "italic" true
            addFieldWhen fields s.Underline "underline" true
            addFieldWhen fields s.UnderlineDouble "underlineDouble" true
            addFieldWhen fields s.Strikethrough "strikethrough" true
            addFieldOption fields "fontName" s.FontName
            addFieldOption fields "fontColor" s.FontColor
            addFieldOption fields "fillColor" s.FillColor
            addFieldObjOption fields "fontSize" (s.FontSize |> Option.map box)
            addFieldOption fields "numberFormat" s.NumberFormat
            addFieldOption fields "horizontalAlign" s.HorizontalAlign
            addFieldOption fields "verticalAlign" s.VerticalAlign
            addFieldWhen fields s.WrapText "wrapText" true
            addFieldObjOption fields "borderTop" (s.BorderTop |> Option.map (fun side -> borderObj (Some side)))
            addFieldObjOption fields "borderBottom" (s.BorderBottom |> Option.map (fun side -> borderObj (Some side)))
            addFieldObjOption fields "borderLeft" (s.BorderLeft |> Option.map (fun side -> borderObj (Some side)))
            addFieldObjOption fields "borderRight" (s.BorderRight |> Option.map (fun side -> borderObj (Some side)))
            addFieldObjOption fields "borderDiagonal" (s.BorderDiagonal |> Option.map (fun side -> borderObj (Some side)))
            addFieldWhen fields s.DiagonalUp "diagonalUp" true
            addFieldWhen fields s.DiagonalDown "diagonalDown" true
        | None -> ()

    let private styleObj (style: XlsxReader.CellStyle option) =
        let fields = ResizeArray()
        addStyleFields fields style
        objFromFields fields

    let private effectiveStyleIndex (workbook: XlsxReader.Workbook) (cell: XlsxReader.Cell) =
        match cell.StyleIndex with
        | Some idx when idx >= 0 && idx < workbook.Styles.Length -> idx
        | None when workbook.Styles.Length > 0 && cell.Style.IsSome -> 0
        | _ -> -1

    let private commentCoordSet (sheet: XlsxReader.Worksheet) =
        let comments = HashSet<struct (int * int)>()
        for comment in sheet.Comments do
            let row = comment.CellRef.Row - 1
            let col = columnIndex comment.CellRef.Column
            if row >= 0 && col >= 0 then
                comments.Add(struct (row, col)) |> ignore
        comments

    let private cellWithStyleObjWithFlags (workbook: XlsxReader.Workbook option) (c: XlsxReader.Cell) hasComment hasListDropdown : obj =
        let style = c.Style
        let fields = ResizeArray()
        addField fields "ref" (cellRefText c.Reference)
        match workbook with
        | Some wb ->
            let value = cellValueObj wb c
            if not (isNull value) then
                addField fields "value" value
        | None -> ()
        let formula = cellFormula c
        if not (isNull formula) then
            addField fields "formula" formula
        let error = cellErrorObj c
        if not (isNull error) then
            addField fields "error" error
        match style with
        | Some _ -> addStyleFields fields style
        | None -> ()
        addFieldWhen fields hasComment "hasComment" true
        addFieldWhen fields hasListDropdown "hasListDropdown" true
        match workbook with
        | Some wb ->
            let richRuns = richRunsForCell wb c
            if richRuns.Length > 0 then
                addField fields "richRuns" richRuns
        | None -> ()
        objFromFields fields

    let private cellWithStyleObj (workbook: XlsxReader.Workbook option) (sheet: XlsxReader.Worksheet) (c: XlsxReader.Cell) : obj =
        let row = c.Reference.Row - 1
        let col = columnIndex c.Reference.Column
        cellWithStyleObjWithFlags workbook c (commentAtCoord row col sheet |> Option.isSome) (listDropdownAt row col sheet)

    /// Load an XLSX file from array buffer or blob
    let load = XlsxReader.load

    /// Load only one worksheet from an XLSX file. None selects the first sheet.
    let loadSelected = XlsxReader.loadSelected

    /// Get all worksheet names, hidden ones included (formulas may refer to
    /// them). Pair with `sheetStates`, or use `visibleSheetNames` for tabs.
    let sheetNames (wb: XlsxReader.Workbook) =
        wb.Sheets |> List.map (fun s -> s.Name) |> List.toArray

    let private sheetStateText (state: XlsxReader.SheetState) =
        match state with
        | XlsxReader.SheetVisible -> "visible"
        | XlsxReader.SheetHidden -> "hidden"
        | XlsxReader.SheetVeryHidden -> "veryHidden"

    /// "visible" | "hidden" | "veryHidden" for each sheet, parallel to `sheetNames`.
    let sheetStates (wb: XlsxReader.Workbook) =
        wb.Sheets |> List.map (fun s -> sheetStateText s.State) |> List.toArray

    /// Names of the sheets Excel shows a tab for.
    let visibleSheetNames (wb: XlsxReader.Workbook) =
        wb.Sheets
        |> List.filter (fun s -> s.State = XlsxReader.SheetVisible)
        |> List.map (fun s -> s.Name)
        |> List.toArray

    /// Get a worksheet by name
    let sheet name (wb: XlsxReader.Workbook) =
        wb.Sheets |> List.tryFind (fun s -> s.Name = name)

    /// Convert worksheet to 2D array
    let toArray = XlsxReader.sheetToArray

    /// Used range dimensions as one-based row/column counts.
    let dimensions (sheet: XlsxReader.Worksheet) =
        let mutable rowCount = 0
        let mutable colCount = 0
        for row in sheet.Rows do
            if row.Index > rowCount then
                rowCount <- row.Index
            for cell in row.Cells do
                let col = columnIndex cell.Reference.Column + 1
                if col > colCount then
                    colCount <- col
        objFromFields [
            objField "rowCount" (box rowCount)
            objField "colCount" (box colCount)
        ]

    let private findCellByRef (cellRef: XlsxReader.CellReference) (sheet: XlsxReader.Worksheet) =
        sheet.Rows
        |> List.tryFind (fun row -> row.Index = cellRef.Row)
        |> Option.bind (fun row ->
            row.Cells
            |> List.tryFind (fun cell -> cell.Reference = cellRef))

    /// Get cell value at reference (e.g., "A1")
    let cellAt (ref: string) (sheet: XlsxReader.Worksheet) =
        let cellRef = XlsxReader.Xlsx.parseCellRef ref
        sheet
        |> findCellByRef cellRef
        |> Option.map (fun c -> c.Value)

    /// Get cell with style at reference
    let cellWithStyleAt (ref: string) (sheet: XlsxReader.Worksheet) : obj =
        let cellRef = XlsxReader.Xlsx.parseCellRef ref
        sheet
        |> findCellByRef cellRef
        |> Option.map (cellWithStyleObj None sheet)
        |> Option.defaultValue null

    /// Stream actual worksheet cells with values/styles in document order.
    /// This avoids the browser importer's old coordinate-by-coordinate scan,
    /// which is pathological for wide style-only sheets.
    let cellsWithStyle (workbook: XlsxReader.Workbook) (sheet: XlsxReader.Worksheet) =
        let comments = commentCoordSet sheet
        let dropdownRanges = listDropdownRangeBounds sheet
        let cells = ResizeArray<obj>()
        for row in sheet.Rows do
            for cell in row.Cells do
                let rowIndex = cell.Reference.Row - 1
                let colIndex = columnIndex cell.Reference.Column
                let hasComment =
                    rowIndex >= 0
                    && colIndex >= 0
                    && comments.Contains(struct (rowIndex, colIndex))
                let hasListDropdown =
                    rowIndex >= 0
                    && colIndex >= 0
                    && rangeBoundsArrayContains rowIndex colIndex dropdownRanges
                cells.Add(cellWithStyleObjWithFlags (Some workbook) cell hasComment hasListDropdown)
        cells.ToArray()

    /// Compact cell stream for importers. Slots are:
    /// row0, col0, value, formula, styleIndex, hasComment, hasListDropdown, richRuns.
    /// (Error results are only reported by `cellsColumnar` / `cellsWithStyle`.)
    let cellsCompact (workbook: XlsxReader.Workbook) (sheet: XlsxReader.Worksheet) =
        let comments = commentCoordSet sheet
        let dropdownRanges = listDropdownRangeBounds sheet
        let cells = ResizeArray<obj>()
        for row in sheet.Rows do
            for cell in row.Cells do
                let rowIndex = cell.Reference.Row - 1
                let colIndex = columnIndex cell.Reference.Column
                let hasComment =
                    rowIndex >= 0
                    && colIndex >= 0
                    && comments.Contains(struct (rowIndex, colIndex))
                let hasListDropdown =
                    rowIndex >= 0
                    && colIndex >= 0
                    && rangeBoundsArrayContains rowIndex colIndex dropdownRanges
                cells.Add(box [|
                    box rowIndex
                    box colIndex
                    cellValueObj workbook cell
                    box (cellFormula cell)
                    box (effectiveStyleIndex workbook cell)
                    box hasComment
                    box hasListDropdown
                    richRunsObjForCell workbook cell
                |])
        cells.ToArray()

    /// Columnar compact cell stream for importers. This avoids one tiny
    /// array/object per cell when importing very large sheets.
    let cellsColumnar (workbook: XlsxReader.Workbook) (sheet: XlsxReader.Worksheet) =
        let comments = commentCoordSet sheet
        let dropdownRanges = listDropdownRangeBounds sheet
        let rows = ResizeArray<int>()
        let cols = ResizeArray<int>()
        let values = ResizeArray<obj>()
        let formulas = ResizeArray<obj>()
        let styleIndexes = ResizeArray<int>()
        let hasComments = ResizeArray<bool>()
        let hasListDropdowns = ResizeArray<bool>()
        let richRuns = ResizeArray<obj>()
        let errors = ResizeArray<obj>()
        for row in sheet.Rows do
            for cell in row.Cells do
                let rowIndex = cell.Reference.Row - 1
                let colIndex = columnIndex cell.Reference.Column
                let hasComment =
                    rowIndex >= 0
                    && colIndex >= 0
                    && comments.Contains(struct (rowIndex, colIndex))
                let hasListDropdown =
                    rowIndex >= 0
                    && colIndex >= 0
                    && rangeBoundsArrayContains rowIndex colIndex dropdownRanges
                rows.Add rowIndex
                cols.Add colIndex
                values.Add(cellValueObj workbook cell)
                formulas.Add(box (cellFormula cell))
                styleIndexes.Add(effectiveStyleIndex workbook cell)
                hasComments.Add hasComment
                hasListDropdowns.Add hasListDropdown
                richRuns.Add(richRunsObjForCell workbook cell)
                errors.Add(cellErrorObj cell)
        objFromFields [
            objField "rows" (box (rows.ToArray()))
            objField "cols" (box (cols.ToArray()))
            objField "values" (box (values.ToArray()))
            objField "formulas" (box (formulas.ToArray()))
            objField "styleIndexes" (box (styleIndexes.ToArray()))
            objField "hasComments" (box (hasComments.ToArray()))
            objField "hasListDropdowns" (box (hasListDropdowns.ToArray()))
            objField "richRuns" (box (richRuns.ToArray()))
            objField "errors" (box (errors.ToArray()))
        ]

    let styleAt (index: int) (workbook: XlsxReader.Workbook) =
        if index >= 0 && index < workbook.Styles.Length then
            styleObj (Some workbook.Styles.[index])
        else
            null

    /// Get per-run rich text for a cell. Returns an empty JS array for plain cells.
    let richRunsAt (ref: string) (workbook: XlsxReader.Workbook) (sheet: XlsxReader.Worksheet) =
        let cellRef = XlsxReader.Xlsx.parseCellRef ref
        sheet
        |> findCellByRef cellRef
        |> Option.map (richRunsForCell workbook)
        |> Option.defaultValue [||]

    /// Get a cell comment/note at reference.
    let commentAt (ref: string) (sheet: XlsxReader.Worksheet) : obj =
        match refToCoord ref with
        | Some (row, col) ->
            match commentAtCoord row col sheet with
            | Some c ->
                createObj [
                    "text" ==> c.Text
                    "author" ==> optString c.Author
                ]
            | None -> null
        | None -> null

    /// Effective column width in Excel character units.
    let private defaultColumnWidth (sheet: XlsxReader.Worksheet) =
        match sheet.Format.DefaultColWidth with
        | Some width -> width
        | None ->
            match sheet.Format.BaseColWidth with
            | Some width -> width + 1.0
            | None -> 8.43

    let columnWidth (col: int) (sheet: XlsxReader.Worksheet) =
        let col1 = col + 1
        match sheet.Columns |> List.tryFind (fun c -> col1 >= c.Index && col1 <= c.MaxIndex) with
        | Some c when c.Hidden -> 0.0
        | Some c -> c.Width |> Option.defaultValue 0.0
        | None -> defaultColumnWidth sheet

    /// Effective column widths in Excel character units for zero-based columns [0, count).
    let columnWidths (count: int) (sheet: XlsxReader.Worksheet) =
        let count = max 0 count
        let widths = Array.create count (defaultColumnWidth sheet)
        for column in sheet.Columns do
            let first = max 0 (column.Index - 1)
            let last = min (count - 1) (column.MaxIndex - 1)
            if first <= last then
                let width = if column.Hidden then 0.0 else column.Width |> Option.defaultValue 0.0
                for col = first to last do
                    widths.[col] <- width
        widths

    /// Effective row height in points.
    let private defaultRowHeight (sheet: XlsxReader.Worksheet) =
        sheet.Format.DefaultRowHeight |> Option.defaultValue 15.0

    let private effectiveRowHeight (sheet: XlsxReader.Worksheet) (row: XlsxReader.Row) =
        row.Height |> Option.orElse sheet.Format.DefaultRowHeight |> Option.defaultValue 15.0

    let rowHeight (row: int) (sheet: XlsxReader.Worksheet) =
        let row1 = row + 1
        match sheet.Rows |> List.tryFind (fun r -> r.Index = row1) with
        | Some r -> effectiveRowHeight sheet r
        | None -> defaultRowHeight sheet

    /// Effective row heights in points for zero-based rows [0, count).
    let rowHeights (count: int) (sheet: XlsxReader.Worksheet) =
        let count = max 0 count
        let heights = Array.create count (defaultRowHeight sheet)
        for row in sheet.Rows do
            let index = row.Index - 1
            if index >= 0 && index < count then
                heights.[index] <- effectiveRowHeight sheet row
        heights

    /// Get merged cells from worksheet
    let mergedCells (sheet: XlsxReader.Worksheet) =
        sheet.MergedCells
        |> List.map (fun mc -> sprintf "%s:%s" (cellRefText mc.StartRef) (cellRefText mc.EndRef))
        |> List.toArray

    /// Get column information
    let columns (sheet: XlsxReader.Worksheet) = sheet.Columns |> List.toArray

    /// Get workbook styles
    let styles (wb: XlsxReader.Workbook) = wb.Styles

    /// Resolve shared strings and cell values
    let resolveCell (wb: XlsxReader.Workbook) (cell: XlsxReader.Cell) =
        let value = XlsxReader.resolveValue wb cell.Value
        { cell with Value = value }

    /// Get cell comments from a worksheet
    let comments (sheet: XlsxReader.Worksheet) = sheet.Comments

    /// Get conditional formatting rules from a worksheet
    let conditionalFormats (sheet: XlsxReader.Worksheet) = sheet.ConditionalFormats

    /// Get data validations from a worksheet
    let dataValidations (sheet: XlsxReader.Worksheet) = sheet.DataValidations

    /// Get charts from a worksheet
    let charts (sheet: XlsxReader.Worksheet) = sheet.Charts

    /// Get pivot tables from a worksheet
    let pivotTables (sheet: XlsxReader.Worksheet) = sheet.PivotTables

    /// Get named ranges from workbook
    let namedRanges (wb: XlsxReader.Workbook) = wb.NamedRanges

    let private optObj (value: string option) : obj =
        match value with
        | Some v -> box v
        | None -> null

    let private optIntObj (value: int option) : obj =
        match value with
        | Some v -> box v
        | None -> null

    /// Header / band fills for a built-in table style, from the workbook
    /// theme: Excel's `TableStyle{Light,Medium,Dark}N` cycle through black
    /// and the six accents (N mod 7).
    let private tableFills (workbook: XlsxReader.Workbook) (styleName: string) : string * string =
        let digits = styleName |> Seq.filter System.Char.IsDigit |> Seq.toArray |> System.String
        let n = match System.Int32.TryParse(digits, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture) with | true, v -> v | _ -> 2
        let slot = (max 1 n - 1) % 7
        let accent =
            if slot = 0 || slot - 1 >= workbook.ThemeAccents.Length then "000000"
            else workbook.ThemeAccents.[slot - 1]
        "#" + accent, "#" + XlsxReader.Xlsx.applyColorTint accent 0.8

    /// Workbook-level facts the cell streams do not carry: the date system,
    /// every sheet with its visibility, and the defined names. `localSheetId`
    /// is the index of the owning sheet in the workbook (null = global).
    let workbookInfo (wb: XlsxReader.Workbook) : obj =
        objFromFields [
            objField "date1904" (box wb.Date1904)
            objField "sheets"
                (box [| for s in wb.Sheets ->
                          objFromFields [ objField "name" (box s.Name); objField "state" (box (sheetStateText s.State)) ] |])
            objField "definedNames"
                (box [| for n in wb.NamedRanges ->
                          objFromFields [
                              objField "name" (box n.Name)
                              objField "refersTo" (box n.RefersTo)
                              objField "localSheetId"
                                  (n.Scope
                                   |> Option.bind (fun s ->
                                       match System.Int32.TryParse(s, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture) with
                                       | true, v -> Some v
                                       | _ -> None)
                                   |> optIntObj)
                          ] |])
            objField "themeAccents" (box wb.ThemeAccents)
        ]

    /// Everything about a sheet besides its cells, as plain objects/arrays an
    /// importer can read without knowing the reader's record types: tab
    /// state, freeze panes, hidden rows/columns (0-based), tables,
    /// data validations, conditional formats (with the dxf each rule
    /// applies already resolved), hyperlinks and print setup.
    let sheetMeta (workbook: XlsxReader.Workbook) (sheet: XlsxReader.Worksheet) : obj =
        let tables =
            [| for t in sheet.Tables ->
                 objFromFields [
                     objField "name" (box t.Name)
                     objField "startRow" (box t.StartRow)
                     objField "startCol" (box t.StartCol)
                     objField "endRow" (box t.EndRow)
                     objField "endCol" (box t.EndCol)
                     objField "headerRowCount" (box t.HeaderRowCount)
                     objField "totalsRowCount" (box t.TotalsRowCount)
                     objField "columns" (box (List.toArray t.Columns))
                     objField "autoFilter"
                         (match t.AutoFilter with
                          | Some (c1, c2) -> box [| c1; c2 |]
                          | None -> null)
                     objField "style"
                         (match t.Style with
                          | Some s ->
                              let headerFill, stripeFill = tableFills workbook s.Name
                              objFromFields [
                                  objField "name" (box s.Name)
                                  objField "showFirstColumn" (box s.ShowFirstColumn)
                                  objField "showLastColumn" (box s.ShowLastColumn)
                                  objField "showRowStripes" (box s.ShowRowStripes)
                                  objField "showColumnStripes" (box s.ShowColumnStripes)
                                  objField "headerFill" (box headerFill)
                                  objField "stripeFill" (box stripeFill)
                              ]
                          | None -> null)
                 ] |]
        let validations =
            [| for dv in sheet.DataValidations ->
                 objFromFields [
                     objField "range" (box dv.Range)
                     objField "type" (box dv.ValidationType)
                     objField "operator" (optObj dv.Operator)
                     objField "formula1" (optObj dv.Formula1)
                     objField "formula2" (optObj dv.Formula2)
                     objField "allowBlank" (box dv.AllowBlank)
                     objField "showDropDown" (box dv.ShowDropDown)
                     objField "error" (optObj dv.ErrorMessage)
                     objField "errorTitle" (optObj dv.ErrorTitle)
                     objField "errorStyle" (optObj dv.ErrorStyle)
                     objField "showErrorMessage" (box dv.ShowErrorMessage)
                     objField "prompt" (optObj dv.Prompt)
                     objField "promptTitle" (optObj dv.PromptTitle)
                     objField "showInputMessage" (box dv.ShowInputMessage)
                 ] |]
        let conditionalFormats =
            [| for cf in sheet.ConditionalFormats ->
                 objFromFields [
                     objField "range" (box cf.Range)
                     objField "rules"
                         (box [| for rule in cf.Rules ->
                                   let dxf =
                                       rule.DxfId
                                       |> Option.bind (fun id -> if id >= 0 && id < workbook.Dxfs.Length then Some workbook.Dxfs.[id] else None)
                                   objFromFields [
                                       objField "type" (box rule.Type)
                                       objField "operator" (optObj rule.Operator)
                                       objField "formulas" (box (List.toArray rule.Formulas))
                                       objField "text" (optObj rule.Text)
                                       objField "priority" (optIntObj rule.Priority)
                                       objField "stopIfTrue" (box rule.StopIfTrue)
                                       objField "rank" (optIntObj rule.Rank)
                                       objField "percent" (box rule.Percent)
                                       objField "bottom" (box rule.Bottom)
                                       objField "aboveAverage" (box rule.AboveAverage)
                                       objField "equalAverage" (box rule.EqualAverage)
                                       objField "stdDev" (optIntObj rule.StdDev)
                                       objField "colorScale"
                                           (box [| for stop in rule.ColorScale ->
                                                     objFromFields [
                                                         objField "type" (box stop.CfvoType)
                                                         objField "value" (optObj stop.CfvoValue)
                                                         objField "color" (box stop.Color)
                                                     ] |])
                                       objField "dataBarColor" (optObj rule.DataBarColor)
                                       objField "dataBarCfvo"
                                           (box [| for (kind, value) in rule.DataBarCfvo ->
                                                     objFromFields [ objField "type" (box kind); objField "value" (optObj value) ] |])
                                       objField "iconSet" (optObj rule.IconSet)
                                       objField "iconSetCfvo"
                                           (box [| for (kind, value) in rule.IconSetCfvo ->
                                                     objFromFields [ objField "type" (box kind); objField "value" (optObj value) ] |])
                                       objField "iconSetReverse" (box rule.IconSetReverse)
                                       objField "style"
                                           (match dxf with
                                            | Some d ->
                                                objFromFields [
                                                    objField "fontColor" (cssColor d.FontColor)
                                                    objField "fillColor" (cssColor d.FillColor)
                                                    objField "bold" (box (d.Bold = Some true))
                                                    objField "italic" (box (d.Italic = Some true))
                                                ]
                                            | None -> null)
                                   ] |])
                 ] |]
        let hyperlinks =
            [| for h in sheet.Hyperlinks ->
                 objFromFields [
                     objField "range" (box h.Range)
                     objField "target" (box h.Target)
                     objField "isInternal" (box h.IsInternal)
                 ] |]
        let hiddenRows = [| for r in sheet.Rows do if r.Hidden then yield r.Index - 1 |]
        let hiddenCols =
            [| for c in sheet.Columns do
                 // A hidden range can run to the last column (16384); the
                 // importer only needs the ones near the data.
                 if c.Hidden then
                     for idx in c.Index .. min c.MaxIndex (c.Index + 1023) do
                         yield idx - 1 |]
        let pageSetup =
            match sheet.PageSetup with
            | None -> null
            | Some p ->
                objFromFields [
                    objField "paperSize" (optIntObj p.PaperSize)
                    objField "orientation"
                        (match p.Orientation with
                         | XlsxReader.OrientationLandscape -> box "landscape"
                         | XlsxReader.OrientationPortrait -> box "portrait"
                         | XlsxReader.OrientationDefault -> null)
                    objField "fitToPage" (box p.FitToPage)
                    objField "fitToWidth" (box p.FitToWidth)
                    objField "fitToHeight" (box p.FitToHeight)
                    objField "scale" (box p.Scale)
                    objField "firstPageNumber" (box p.FirstPageNumber)
                    objField "useFirstPageNumber" (box p.UseFirstPageNumber)
                    objField "pageOrder" (box (match p.PageOrder with XlsxReader.OverThenDown -> "overThenDown" | XlsxReader.DownThenOver -> "downThenOver"))
                    objField "blackAndWhite" (box p.BlackAndWhite)
                    objField "cellComments"
                        (box (match p.CellComments with
                              | XlsxReader.CellCommentsAsDisplayed -> "asDisplayed"
                              | XlsxReader.CellCommentsAtEnd -> "atEnd"
                              | XlsxReader.CellCommentsNone -> "none"))
                ]
        let pageMargins =
            match sheet.PageMargins with
            | None -> null
            | Some m ->
                objFromFields [
                    objField "top" (box m.Top)
                    objField "bottom" (box m.Bottom)
                    objField "left" (box m.Left)
                    objField "right" (box m.Right)
                    objField "header" (box m.Header)
                    objField "footer" (box m.Footer)
                ]
        let hf = sheet.HeaderFooter
        objFromFields [
            objField "name" (box sheet.Name)
            objField "state" (box (sheetStateText sheet.State))
            objField "frozenRows" (box (sheet.FrozenPane |> Option.map (fun p -> p.YSplit) |> Option.defaultValue 0))
            objField "frozenCols" (box (sheet.FrozenPane |> Option.map (fun p -> p.XSplit) |> Option.defaultValue 0))
            objField "showGridLines" (box sheet.ShowGridLines)
            objField "showZeros" (box sheet.ShowZeros)
            objField "hiddenRows" (box hiddenRows)
            objField "hiddenCols" (box hiddenCols)
            objField "tables" (box tables)
            objField "validations" (box validations)
            objField "conditionalFormats" (box conditionalFormats)
            objField "hyperlinks" (box hyperlinks)
            objField "pageSetup" pageSetup
            objField "pageMargins" pageMargins
            objField "printOptions"
                (objFromFields [
                    objField "gridLines" (box sheet.PrintOptions.GridLines)
                    objField "headings" (box sheet.PrintOptions.Headings)
                    objField "horizontalCentered" (box sheet.PrintOptions.HorizontalCentered)
                    objField "verticalCentered" (box sheet.PrintOptions.VerticalCentered)
                ])
            objField "headerFooter"
                (objFromFields [
                    objField "oddHeader" (optObj hf.OddHeader)
                    objField "oddFooter" (optObj hf.OddFooter)
                    objField "evenHeader" (optObj hf.EvenHeader)
                    objField "evenFooter" (optObj hf.EvenFooter)
                    objField "firstHeader" (optObj hf.FirstHeader)
                    objField "firstFooter" (optObj hf.FirstFooter)
                    objField "differentOddEven" (box hf.DifferentOddEven)
                    objField "differentFirst" (box hf.DifferentFirst)
                    objField "scaleWithDoc" (box hf.ScaleWithDoc)
                ])
            objField "rowBreaks" (box (List.toArray sheet.RowBreaks))
            objField "colBreaks" (box (List.toArray sheet.ColBreaks))
            objField "printArea" (optObj sheet.PrintArea)
            objField "printTitleRows" (optObj sheet.PrintTitleRows)
            objField "printTitleCols" (optObj sheet.PrintTitleCols)
        ]

    /// Evaluate a cell's formula (stub)
    let evaluateFormula (wb: XlsxReader.Workbook) (sheet: XlsxReader.Worksheet) (cell: XlsxReader.Cell) =
        XlsxReader.FormulaEngine.evaluateCell wb sheet cell

// Example usage demonstrating the DSL
[<ExportDefault>]
let api () = {|
    word = {|
        load = Word.load
        extractText = Word.extractText
        getContent = Word.getContent
        paragraphs = Word.paragraphs
        tables = Word.tables
        styles = Word.styles
        headers = Word.headers
        footers = Word.footers
        properties = Word.properties
        hyperlinks = Word.hyperlinks
        images = Word.images
        comments = Word.comments
        revisions = Word.revisions
        footnotes = Word.footnotes
        endnotes = Word.endnotes
        shapes = Word.shapes
        write = WordWriter.write
        writeFromEditor = WordWriter.writeFromEditor
        writeFromEditorPackage = WordWriter.writeFromEditorPackage
        createDocument = WordWriter.createDocument
        addParagraph = WordWriter.addParagraph
        addParagraphWithRuns = WordWriter.addParagraphWithRuns
        saveDocument = WordWriter.saveDocument
        createRun = WordWriter.createRun
        createSimpleTable = WordWriter.createSimpleTable
        createTableWithBorders = WordWriter.createTableWithBorders
        generateTableXml = WordWriter.generateTableXml
    |}
    processors = {|
        forEditorPackage = Processors.forEditorPackage
    |}
    excel = {|
        load = Excel.load
        loadSelected = Excel.loadSelected
        sheetNames = Excel.sheetNames
        sheetStates = Excel.sheetStates
        visibleSheetNames = Excel.visibleSheetNames
        workbookInfo = Excel.workbookInfo
        sheetMeta = Excel.sheetMeta
        sheet = Excel.sheet
        toArray = Excel.toArray
        dimensions = Excel.dimensions
        cellAt = Excel.cellAt
        cellWithStyleAt = Excel.cellWithStyleAt
        cellsWithStyle = Excel.cellsWithStyle
        cellsCompact = Excel.cellsCompact
        cellsColumnar = Excel.cellsColumnar
        styleAt = Excel.styleAt
        richRunsAt = Excel.richRunsAt
        commentAt = Excel.commentAt
        columnWidth = Excel.columnWidth
        columnWidths = Excel.columnWidths
        rowHeight = Excel.rowHeight
        rowHeights = Excel.rowHeights
        mergedCells = Excel.mergedCells
        columns = Excel.columns
        styles = Excel.styles
        resolveCell = Excel.resolveCell
        comments = Excel.comments
        conditionalFormats = Excel.conditionalFormats
        dataValidations = Excel.dataValidations
        charts = Excel.charts
        pivotTables = Excel.pivotTables
        namedRanges = Excel.namedRanges
        evaluateFormula = Excel.evaluateFormula
        write = ExcelWriter.write
        writeFormatted = ExcelWriter.writeFormatted
        createWorkbook = ExcelWriter.createWorkbook
        addSheet = ExcelWriter.addSheet
        addSimpleSheet = ExcelWriter.addSimpleSheet
        saveWorkbook = ExcelWriter.saveWorkbook
        addRow = ExcelWriter.addRow
        addSimpleRow = ExcelWriter.addSimpleRow
        createCell = ExcelWriter.createCell
        createHeaderCell = ExcelWriter.createHeaderCell
        createNumericCell = ExcelWriter.createNumericCell
        createFormulaCell = ExcelWriter.createFormulaCell
        addMergedRange = ExcelWriter.addMergedRange
        createMergedRange = ExcelWriter.createMergedRange
    |}
|}

// Browser-friendly usage example
[<Emit("console.log($0)")>]
let log (msg: obj) : unit = jsNative

// Example: Read DOCX from file input
let readWordDoc (file: obj) : Fable.Core.JS.Promise<unit> =
    promise {
        let! doc = Word.load file
        match doc with
        | Some d ->
            let text = Word.extractText d
            log {| docText = text; paragraphCount = (Word.paragraphs d).Length |}
        | None ->
            log "Failed to load document"
    }

// Example: Read XLSX from file input
let readExcelDoc (file: obj) : Fable.Core.JS.Promise<unit> =
    promise {
        let! wb = Excel.load file
        match wb with
        | Some w ->
            let sheets = Excel.sheetNames w
            log {| workbook = w; sheetNames = sheets |}

            // Get first sheet as array
            match w.Sheets with
            | sheet :: _ ->
                let arr = Excel.toArray w sheet
                log {| sheetData = arr |}
            | [] -> ()
        | None ->
            log "Failed to load workbook"
    }
