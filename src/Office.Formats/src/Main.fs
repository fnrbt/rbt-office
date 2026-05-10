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

    let private wordFontSizeToPoints (size: string option) =
        size
        |> Option.bind (fun value ->
            match System.Double.TryParse(value) with
            | true, parsed -> Some (parsed / 2.0)
            | false, _ -> None)

    let private wordTwipsToPoints (value: string option) =
        value
        |> Option.bind (fun text ->
            match System.Double.TryParse(text) with
            | true, parsed -> Some (parsed / 20.0)
            | false, _ -> None)

    let private wordLineSpacingValue (fontSize: float option) (value: string option) (rule: string option) =
        value
        |> Option.bind (fun text ->
            match System.Double.TryParse(text) with
            | true, parsed ->
                match rule with
                | Some "exact"
                | Some "atLeast" -> Some (parsed / 20.0)
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
            "StyleFontSize" ==> (style.RunProperties |> Option.bind (fun props -> wordFontSizeToPoints props.FontSize))
            "StyleFontFamily" ==> (style.RunProperties |> Option.bind (fun props -> props.FontFamily |> Option.orElse (resolveThemeFont theme props.ThemeFont)))
            "StyleColor" ==> (style.RunProperties |> Option.bind (fun props -> wordColorToCss props.Color |> Option.orElse (resolveThemeColor theme props.ThemeColor props.ThemeTint props.ThemeShade)))
            "StyleHighlight" ==> (style.RunProperties |> Option.bind (fun props -> props.Highlight))
            "StyleSpacingBefore" ==> (style.ParagraphProperties |> Option.bind (fun props -> wordTwipsToPoints props.SpacingBefore))
            "StyleSpacingAfter" ==> (style.ParagraphProperties |> Option.bind (fun props -> wordTwipsToPoints props.SpacingAfter))
            "StyleLineSpacing" ==> (style.ParagraphProperties |> Option.bind (fun props -> wordLineSpacingValue (style.RunProperties |> Option.bind (fun run -> wordFontSizeToPoints run.FontSize)) props.LineSpacing props.LineRule))
            "StyleLineRule" ==> (style.ParagraphProperties |> Option.bind (fun props -> props.LineRule))
            "StyleAlignment" ==> (style.ParagraphProperties |> Option.bind (fun props -> props.Alignment))
            "StyleIndentLeft" ==> (style.ParagraphProperties |> Option.bind (fun props -> wordTwipsToPoints props.IndentLeft))
            "StyleIndentRight" ==> (style.ParagraphProperties |> Option.bind (fun props -> wordTwipsToPoints props.IndentRight))
            "StyleFirstLineIndent" ==> (style.ParagraphProperties |> Option.bind (fun props -> wordTwipsToPoints props.FirstLineIndent))
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
            "SpacingBefore" ==> para.SpacingBefore
            "SpacingAfter" ==> para.SpacingAfter
            "LineSpacing" ==> para.LineSpacing
        ]

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
                let shapeObjects =
                    doc.Shapes
                    |> List.filter (fun shape -> not shape.Content.IsEmpty)
                    |> List.map (fun shape ->
                        let paragraphs =
                            shape.Content
                            |> List.map (DocumentProcessors.EditorProcessor.toEditorParagraph doc.Theme styleMap comments notes >> editorParagraphToObj)
                            |> List.toArray
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
                            "Paragraphs" ==> paragraphs
                        ])
                    |> List.toArray

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
                let! paragraphs = forEditor fileData
                return createObj [
                    "Paragraphs" ==> paragraphs
                    "Styles" ==> (doc.Styles |> List.map (styleToObj doc.Theme) |> List.toArray)
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
    /// Load an XLSX file from array buffer or blob
    let load = XlsxReader.load

    /// Get all worksheet names
    let sheetNames (wb: XlsxReader.Workbook) =
        wb.Sheets |> List.map (fun s -> s.Name)

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
    let cellWithStyleAt (ref: string) (sheet: XlsxReader.Worksheet) =
        let cellRef = XlsxReader.Xlsx.parseCellRef ref
        sheet.Rows
        |> List.collect (fun r -> r.Cells)
        |> List.tryFind (fun c -> c.Reference = cellRef)

    /// Get merged cells from worksheet
    let mergedCells (sheet: XlsxReader.Worksheet) = sheet.MergedCells

    /// Get column information
    let columns (sheet: XlsxReader.Worksheet) = sheet.Columns

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
