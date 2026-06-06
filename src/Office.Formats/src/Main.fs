module Main

open Fable.Core
open Fable.Core.JsInterop
open ZipInterop

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
            "StyleFirstLineIndent" ==> (style.ParagraphProperties |> Option.bind (fun props -> wordTwipsToCssPixels props.FirstLineIndent))
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
                "Paragraphs" ==> paragraphs
            ])
        |> List.toArray

    /// Transform Word file data to editor-friendly format with embedded resources (images) resolved
    /// Takes raw file data and returns paragraphs with fully resolved styles and embedded images
    let forEditor (fileData: obj) : Fable.Core.JS.Promise<obj array> =
        promise {
            let! archive = loadArchive fileData
            let! docOpt = DocxReader.parseDocument archive

            match docOpt with
            | Some doc ->
                let paragraphs = DocumentProcessors.EditorProcessor.extractForEditor doc

                // Resolve images for each run
                let! resolvedParagraphs =
                    paragraphs
                    |> List.map (fun para ->
                        promise {
                            let! resolvedRuns =
                                para.Runs
                                |> List.map (fun run ->
                                    promise {
                                        if run.IsImage then
                                            // RelationshipId is temporarily stored in Text field
                                            let relationshipId = run.Text
                                            let! imageSrcOpt = DocxReader.extractImageDataUrl archive relationshipId doc.Relationships

                                            return {
                                                run with
                                                    Text = run.ImageAlt |> Option.defaultValue ""
                                                    ImageSrc = imageSrcOpt
                                            }
                                        else
                                            return run
                                    })
                                |> Array.ofList
                                |> ZipInterop.Promise.all

                            return
                                editorParagraphToObj {
                                    para with Runs = resolvedRuns |> Array.toList
                                }
                        })
                    |> Array.ofList
                    |> ZipInterop.Promise.all

                let shapeObjects = shapeObjectsFor doc

                return Array.append resolvedParagraphs shapeObjects
            | None ->
                return [||]
        }

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
    let private columnIndex (letters: string) =
        let mutable acc = 0
        for ch in letters do
            if System.Char.IsLetter ch then
                acc <- acc * 26 + (int (System.Char.ToUpperInvariant ch) - int 'A' + 1)
        acc - 1

    let private refToCoord (ref: string) =
        let clean = ref.Replace("$", "")
        let letters = clean |> Seq.takeWhile System.Char.IsLetter |> Seq.toArray |> System.String
        let digits = clean |> Seq.skipWhile System.Char.IsLetter |> Seq.toArray |> System.String
        match System.Int32.TryParse digits with
        | true, row when letters.Length > 0 -> Some (row - 1, columnIndex letters)
        | _ -> None

    let private rangeContains (row: int) (col: int) (range: string) =
        let parts = range.Split(':')
        match parts.Length with
        | 1 ->
            refToCoord parts.[0]
            |> Option.exists (fun (r, c) -> r = row && c = col)
        | 2 ->
            match refToCoord parts.[0], refToCoord parts.[1] with
            | Some (r1, c1), Some (r2, c2) ->
                row >= min r1 r2 && row <= max r1 r2
                && col >= min c1 c2 && col <= max c1 c2
            | _ -> false
        | _ -> false

    let private rangeListContains row col (ranges: string) =
        ranges.Split([| ' ' |], System.StringSplitOptions.RemoveEmptyEntries)
        |> Array.exists (rangeContains row col)

    let private optString (value: string option) =
        value |> Option.defaultValue null

    let private cellRefText (cellRef: XlsxReader.CellReference) =
        sprintf "%s%d" cellRef.Column cellRef.Row

    let private optFloat (value: float option) : obj =
        match value with
        | Some v -> box v
        | None -> null

    let private borderObj (side: XlsxReader.CellBorderSide option) : obj =
        match side with
        | Some s ->
            createObj [
                "style" ==> s.Style
                "color" ==> optString s.Color
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
        createObj [
            "text" ==> text
            if props.Bold then "bold" ==> true
            if props.Italic then "italic" ==> true
            if props.Underline then "underline" ==> true
            if props.UnderlineDouble then "underlineDouble" ==> true
            if props.Strikethrough then "strikethrough" ==> true
            "color" ==> cssColor props.Color
            "fontName" ==> optString props.FontName
            "fontSize" ==> optFloat props.FontSize
            match props.VertAlign with
            | XlsxReader.VertAlignSuperscript -> "vertAlign" ==> "superscript"
            | XlsxReader.VertAlignSubscript -> "vertAlign" ==> "subscript"
            | XlsxReader.VertAlignBaseline -> ()
        ]

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

    let private cellFormula (cell: XlsxReader.Cell) =
        match cell.Value with
        | XlsxReader.Formula f -> f
        | _ -> null

    let private cellValueObj (workbook: XlsxReader.Workbook) (cell: XlsxReader.Cell) : obj =
        match XlsxReader.resolveValue workbook cell.Value with
        | XlsxReader.String s -> box s
        | XlsxReader.Number n -> box n
        | XlsxReader.Boolean b -> box b
        | XlsxReader.Date d -> box d
        | XlsxReader.Formula f -> box f
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

    let private cellWithStyleObj (workbook: XlsxReader.Workbook option) (sheet: XlsxReader.Worksheet) (c: XlsxReader.Cell) : obj =
        let row = c.Reference.Row - 1
        let col = columnIndex c.Reference.Column
        let style = c.Style
        createObj [
            "ref" ==> cellRefText c.Reference
            "value" ==>
                (match workbook with
                 | Some wb -> cellValueObj wb c
                 | None -> null)
            "formula" ==> cellFormula c
            "bold" ==> (style |> Option.map (fun s -> s.Bold) |> Option.defaultValue false)
            "italic" ==> (style |> Option.map (fun s -> s.Italic) |> Option.defaultValue false)
            "underline" ==> (style |> Option.map (fun s -> s.Underline) |> Option.defaultValue false)
            "underlineDouble" ==> (style |> Option.map (fun s -> s.UnderlineDouble) |> Option.defaultValue false)
            "strikethrough" ==> (style |> Option.map (fun s -> s.Strikethrough) |> Option.defaultValue false)
            "fontName" ==> (style |> Option.bind (fun s -> s.FontName) |> optString)
            "fontColor" ==> (style |> Option.bind (fun s -> s.FontColor) |> optString)
            "fillColor" ==> (style |> Option.bind (fun s -> s.FillColor) |> optString)
            "fontSize" ==> (style |> Option.bind (fun s -> s.FontSize) |> optFloat)
            "numberFormat" ==> (style |> Option.bind (fun s -> s.NumberFormat) |> optString)
            "horizontalAlign" ==> (style |> Option.bind (fun s -> s.HorizontalAlign) |> optString)
            "verticalAlign" ==> (style |> Option.bind (fun s -> s.VerticalAlign) |> optString)
            "wrapText" ==> (style |> Option.map (fun s -> s.WrapText) |> Option.defaultValue false)
            "borderTop" ==> (style |> Option.bind (fun s -> s.BorderTop) |> borderObj)
            "borderBottom" ==> (style |> Option.bind (fun s -> s.BorderBottom) |> borderObj)
            "borderLeft" ==> (style |> Option.bind (fun s -> s.BorderLeft) |> borderObj)
            "borderRight" ==> (style |> Option.bind (fun s -> s.BorderRight) |> borderObj)
            "borderDiagonal" ==> (style |> Option.bind (fun s -> s.BorderDiagonal) |> borderObj)
            "diagonalUp" ==> (style |> Option.map (fun s -> s.DiagonalUp) |> Option.defaultValue false)
            "diagonalDown" ==> (style |> Option.map (fun s -> s.DiagonalDown) |> Option.defaultValue false)
            "hasComment" ==> (commentAtCoord row col sheet |> Option.isSome)
            "hasListDropdown" ==> listDropdownAt row col sheet
            "richRuns" ==>
                (match workbook with
                 | Some wb -> richRunsForCell wb c
                 | None -> [||])
        ]

    /// Load an XLSX file from array buffer or blob
    let load = XlsxReader.load

    /// Get all worksheet names
    let sheetNames (wb: XlsxReader.Workbook) =
        wb.Sheets |> List.map (fun s -> s.Name) |> List.toArray

    /// Get a worksheet by name
    let sheet name (wb: XlsxReader.Workbook) =
        wb.Sheets |> List.tryFind (fun s -> s.Name = name)

    /// Convert worksheet to 2D array
    let toArray = XlsxReader.sheetToArray

    /// Get cell value at reference (e.g., "A1")
    let cellAt (ref: string) (sheet: XlsxReader.Worksheet) =
        let cellRef = XlsxReader.Xlsx.parseCellRef ref
        sheet.Rows
        |> List.collect (fun r -> r.Cells)
        |> List.tryFind (fun c -> c.Reference = cellRef)
        |> Option.map (fun c -> c.Value)

    /// Get cell with style at reference
    let cellWithStyleAt (ref: string) (sheet: XlsxReader.Worksheet) : obj =
        let cellRef = XlsxReader.Xlsx.parseCellRef ref
        sheet.Rows
        |> List.collect (fun r -> r.Cells)
        |> List.tryFind (fun c -> c.Reference = cellRef)
        |> Option.map (cellWithStyleObj None sheet)
        |> Option.defaultValue null

    /// Stream actual worksheet cells with values/styles in document order.
    /// This avoids the browser importer's old coordinate-by-coordinate scan,
    /// which is pathological for wide style-only sheets.
    let cellsWithStyle (workbook: XlsxReader.Workbook) (sheet: XlsxReader.Worksheet) =
        sheet.Rows
        |> List.collect (fun r -> r.Cells)
        |> List.map (cellWithStyleObj (Some workbook) sheet)
        |> List.toArray

    /// Get per-run rich text for a cell. Returns an empty JS array for plain cells.
    let richRunsAt (ref: string) (workbook: XlsxReader.Workbook) (sheet: XlsxReader.Worksheet) =
        let cellRef = XlsxReader.Xlsx.parseCellRef ref
        sheet.Rows
        |> List.collect (fun r -> r.Cells)
        |> List.tryFind (fun c -> c.Reference = cellRef)
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
    let columnWidth (col: int) (sheet: XlsxReader.Worksheet) =
        let col1 = col + 1
        match sheet.Columns |> List.tryFind (fun c -> col1 >= c.Index && col1 <= c.MaxIndex) with
        | Some c when c.Hidden -> 0.0
        | Some c -> c.Width |> Option.defaultValue 0.0
        | None ->
            match sheet.Format.DefaultColWidth with
            | Some width -> width
            | None ->
                match sheet.Format.BaseColWidth with
                | Some width -> width + 1.0
                | None -> 8.43

    /// Effective row height in points.
    let rowHeight (row: int) (sheet: XlsxReader.Worksheet) =
        let row1 = row + 1
        match sheet.Rows |> List.tryFind (fun r -> r.Index = row1) with
        | Some r -> r.Height |> Option.orElse sheet.Format.DefaultRowHeight |> Option.defaultValue 15.0
        | None -> sheet.Format.DefaultRowHeight |> Option.defaultValue 15.0

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
        forEditor = Processors.forEditor
        forEditorPackage = Processors.forEditorPackage
    |}
    excel = {|
        load = Excel.load
        sheetNames = Excel.sheetNames
        sheet = Excel.sheet
        toArray = Excel.toArray
        cellAt = Excel.cellAt
        cellWithStyleAt = Excel.cellWithStyleAt
        cellsWithStyle = Excel.cellsWithStyle
        richRunsAt = Excel.richRunsAt
        commentAt = Excel.commentAt
        columnWidth = Excel.columnWidth
        rowHeight = Excel.rowHeight
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
