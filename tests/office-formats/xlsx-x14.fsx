#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

// The x14 extension of a sheet (Excel 2010+): data bars with their x14 twin,
// icon sets with custom icons, rules and validations whose formulas read
// another sheet. Read from a package laid out the way Excel writes it (and
// from a real workbook), modelled, written the same way, read again; what is
// not understood is carried as it is.

open System
open System.IO
open System.IO.Compression
open Rbt.Spreadsheet.Model
module Bridge = Rbt.Spreadsheet.XlsxBridge
module CF = Rbt.Spreadsheet.ConditionalFormatting

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-34s  %s" label detail)
let eq label (got: 'a) (expected: 'a) =
    ok label (got = expected) (sprintf "expected %A  got %A" expected got)

let root = Path.Combine(__SOURCE_DIRECTORY__, "..")
let loadBytes (bytes: byte[]) = XlsxReader.loadBytesNative bytes |> Option.get
let partsOf (bytes: byte[]) : Map<string, byte[]> =
    use archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read)
    [ for entry in archive.Entries ->
        use stream = entry.Open()
        use buffer = new MemoryStream()
        stream.CopyTo buffer
        entry.FullName, buffer.ToArray() ]
    |> Map.ofList
let textOf (parts: Map<string, byte[]>) (path: string) =
    match Map.tryFind path parts with
    | Some data -> Text.Encoding.UTF8.GetString data
    | None -> ""
let zipOf (parts: (string * string) list) : byte[] =
    use buffer = new MemoryStream()
    (use archive = new ZipArchive(buffer, ZipArchiveMode.Create, true)
     for path, text in parts do
        let entry = archive.CreateEntry path
        use stream = entry.Open()
        let data = Text.Encoding.UTF8.GetBytes (text: string)
        stream.Write(data, 0, data.Length))
    buffer.ToArray()
let count (needle: string) (text: string) =
    let mutable n = 0
    let mutable at = text.IndexOf needle
    while at >= 0 do
        n <- n + 1
        at <- text.IndexOf (needle, at + needle.Length)
    n
let addr row col : CellAddress = { Row = row; Col = col }
let range sr sc er ec : CellRange = { From = addr sr sc; To = addr er ec }
let export (imported: Bridge.ImportedWorkbook) (sheets: SpreadsheetModel array) =
    XlsxWriter.writeWorkbookPackageBytes
        (Bridge.Export.toWorkbookOut
            { Sheets = sheets
              Names = imported.Sheets |> Array.map (fun sheet -> sheet.Name)
              Visible = Array.create sheets.Length true
              ActiveSheetIndex = 0
              NameInfo = Map.empty
              DefinedNames = imported.Context.Names
              Source = Some imported.Source
              Macros = None; MacroProject = None })
let modelsOf (imported: Bridge.ImportedWorkbook) = imported.Sheets |> Array.map (fun sheet -> sheet.Model)

// ---- a package as Excel writes it ----------------------------------------------
let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
let x14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
let xm = "http://schemas.microsoft.com/office/excel/2006/main"
let decl = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
let rows (values: float list list) =
    values
    |> List.mapi (fun r cells ->
        sprintf "<row r=\"%d\">%s</row>" (r + 1)
            (cells |> List.mapi (fun c v -> sprintf "<c r=\"%s%d\"><v>%g</v></c>" (string (char (int 'A' + c))) (r + 1) v) |> String.concat ""))
    |> String.concat ""
let barId = "{C0DA50BC-9888-411D-B3C3-DCE2CC2ECCB4}"
let barId2 = "{11111111-2222-4333-8444-555555555555}"
let sparklines =
    "<ext uri=\"{05C60535-1F16-4fd2-B633-F4F36F0B64E0}\" xmlns:x14=\"" + x14 + "\"><x14:sparklineGroups xmlns:xm=\"" + xm + "\"><x14:sparklineGroup displayEmptyCellsAs=\"gap\"><x14:colorSeries rgb=\"FF376092\"/><x14:sparklines><x14:sparkline><xm:f>Data!A1:C1</xm:f><xm:sqref>F1</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext>"
/// `extraRule`: another x14 rule in the list (for the "not understood" case).
let dataSheet (extraRule: string) =
    decl + "<worksheet xmlns=\"" + ns + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">"
    + "<sheetData>" + rows [ [ -50.0; 1.0; 5.0 ]; [ -25.0; 2.0; 4.0 ]; [ 0.0; 3.0; 3.0 ]; [ 50.0; 4.0; 2.0 ]; [ 100.0; 5.0; 1.0 ] ] + "</sheetData>"
    // The 2007 data bars, each naming its x14 twin; an ordinary rule between them.
    + "<conditionalFormatting sqref=\"A1:A5\"><cfRule type=\"dataBar\" priority=\"1\"><dataBar><cfvo type=\"min\"/><cfvo type=\"max\"/><color rgb=\"FF638EC6\"/></dataBar><extLst><ext uri=\"{B025F937-C7B1-47D3-B67F-A62EFF666E3E}\" xmlns:x14=\"" + x14 + "\"><x14:id>" + barId + "</x14:id></ext></extLst></cfRule></conditionalFormatting>"
    + "<conditionalFormatting sqref=\"B1:B5\"><cfRule type=\"cellIs\" dxfId=\"0\" priority=\"3\" operator=\"greaterThan\"><formula>3</formula></cfRule><cfRule type=\"dataBar\" priority=\"2\"><dataBar><cfvo type=\"num\" val=\"0\"/><cfvo type=\"max\"/><color rgb=\"FF63C384\"/></dataBar><extLst><ext uri=\"{B025F937-C7B1-47D3-B67F-A62EFF666E3E}\" xmlns:x14=\"" + x14 + "\"><x14:id>" + barId2 + "</x14:id></ext></extLst></cfRule></conditionalFormatting>"
    + "<dataValidations count=\"1\"><dataValidation type=\"whole\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"D1:D5\"><formula1>1</formula1><formula2>9</formula2></dataValidation></dataValidations>"
    + "<extLst>"
    + "<ext uri=\"{78C0D931-6437-407d-A8EE-F0AAD7539E65}\" xmlns:x14=\"" + x14 + "\"><x14:conditionalFormattings>"
    + "<x14:conditionalFormatting xmlns:xm=\"" + xm + "\"><x14:cfRule type=\"dataBar\" id=\"" + barId + "\"><x14:dataBar minLength=\"0\" maxLength=\"100\" gradient=\"0\"><x14:cfvo type=\"autoMin\"/><x14:cfvo type=\"autoMax\"/><x14:negativeFillColor rgb=\"FFFF0000\"/><x14:axisColor rgb=\"FF000000\"/></x14:dataBar></x14:cfRule><xm:sqref>A1:A5</xm:sqref></x14:conditionalFormatting>"
    + "<x14:conditionalFormatting xmlns:xm=\"" + xm + "\"><x14:cfRule type=\"dataBar\" id=\"" + barId2 + "\"><x14:dataBar minLength=\"5\" maxLength=\"95\" showValue=\"0\" border=\"1\" direction=\"rightToLeft\" negativeBarColorSameAsPositive=\"1\" negativeBarBorderColorSameAsPositive=\"0\" axisPosition=\"middle\"><x14:cfvo type=\"num\"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type=\"formula\"><xm:f>MAX($B$1:$B$5)*2</xm:f></x14:cfvo><x14:borderColor rgb=\"FF2E7D32\"/><x14:negativeBorderColor rgb=\"FF990000\"/><x14:axisColor rgb=\"FF7F7F7F\"/></x14:dataBar></x14:cfRule><xm:sqref>B1:B5</xm:sqref></x14:conditionalFormatting>"
    + "<x14:conditionalFormatting xmlns:xm=\"" + xm + "\"><x14:cfRule type=\"iconSet\" priority=\"4\" id=\"{AAAAAAAA-0000-4000-8000-000000000001}\"><x14:iconSet iconSet=\"3Triangles\" showValue=\"0\" custom=\"1\"><x14:cfvo type=\"percent\"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type=\"num\"><xm:f>2</xm:f></x14:cfvo><x14:cfvo type=\"formula\" gte=\"0\"><xm:f>Lists!$B$1</xm:f></x14:cfvo><x14:cfIcon iconSet=\"3Arrows\" iconId=\"0\"/><x14:cfIcon iconSet=\"NoIcons\" iconId=\"0\"/><x14:cfIcon iconSet=\"3Stars\" iconId=\"2\"/></x14:iconSet></x14:cfRule><xm:sqref>C1:C5</xm:sqref></x14:conditionalFormatting>"
    + "<x14:conditionalFormatting xmlns:xm=\"" + xm + "\"><x14:cfRule type=\"expression\" priority=\"5\" stopIfTrue=\"1\" id=\"{AAAAAAAA-0000-4000-8000-000000000002}\"><xm:f>Lists!$B$1&gt;$A1</xm:f><x14:dxf><font><b/><i/><color rgb=\"FF9C0006\"/></font><fill><patternFill><bgColor rgb=\"FFFFC7CE\"/></patternFill></fill></x14:dxf></x14:cfRule>"
    + "<x14:cfRule type=\"cellIs\" priority=\"6\" operator=\"between\" id=\"{AAAAAAAA-0000-4000-8000-000000000003}\"><xm:f>Lists!$B$1</xm:f><xm:f>Lists!$B$2</xm:f><x14:dxf><fill><patternFill><bgColor rgb=\"FFC6EFCE\"/></patternFill></fill></x14:dxf></x14:cfRule><xm:sqref>E1:E5 G1:G3</xm:sqref></x14:conditionalFormatting>"
    + extraRule
    + "</x14:conditionalFormattings></ext>"
    + "<ext uri=\"{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}\" xmlns:x14=\"" + x14 + "\"><x14:dataValidations count=\"2\" xmlns:xm=\"" + xm + "\">"
    + "<x14:dataValidation type=\"list\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" errorTitle=\"Pick one\" error=\"Not in the list\"><x14:formula1><xm:f>Lists!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>H1:H5</xm:sqref></x14:dataValidation>"
    + "<x14:dataValidation type=\"decimal\" operator=\"lessThanOrEqual\" showErrorMessage=\"1\"><x14:formula1><xm:f>Lists!$B$2</xm:f></x14:formula1><xm:sqref>I1:I5 K2</xm:sqref></x14:dataValidation>"
    + "</x14:dataValidations></ext>"
    + sparklines
    + "</extLst></worksheet>"
let listsSheet =
    decl + "<worksheet xmlns=\"" + ns + "\"><sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>red</t></is></c><c r=\"B1\"><v>3</v></c></row><row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>green</t></is></c><c r=\"B2\"><v>40</v></c></row><row r=\"3\"><c r=\"A3\" t=\"inlineStr\"><is><t>blue</t></is></c></row></sheetData></worksheet>"
let packageOf (extraRule: string) =
    zipOf
        [ "[Content_Types].xml",
          decl + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/>"
          + "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
          + "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
          + "<Override PartName=\"/xl/worksheets/sheet2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
          + "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/></Types>"
          "_rels/.rels", decl + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>"
          "xl/workbook.xml", decl + "<workbook xmlns=\"" + ns + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"Data\" sheetId=\"1\" r:id=\"rId1\"/><sheet name=\"Lists\" sheetId=\"2\" r:id=\"rId2\"/></sheets></workbook>"
          "xl/_rels/workbook.xml.rels",
          decl + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet2.xml\"/><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/></Relationships>"
          "xl/styles.xml",
          decl + "<styleSheet xmlns=\"" + ns + "\"><fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts><fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills><borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs><cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs><dxfs count=\"1\"><dxf><fill><patternFill><bgColor rgb=\"FFFFEB9C\"/></patternFill></fill></dxf></dxfs></styleSheet>"
          "xl/worksheets/sheet1.xml", dataSheet extraRule
          "xl/worksheets/sheet2.xml", listsSheet ]

// ---- reading -------------------------------------------------------------------
let cfvo kind value : CfThreshold = { Kind = kind; Value = value }
let expectedRules : (string * ConditionalFormatRule * CellFormat * bool) list =
    [ "A1:A5", DataBarEx (cfvo CfvoAutoMin None, cfvo CfvoAutoMax None, "#638EC6",
                          { Gradient = false; Border = None; NegativeFill = Some "#FF0000"; NegativeBorder = None; Axis = AxisAutomatic; AxisColor = Some "#000000"
                            Direction = BarContext; ShowValue = true; MinLength = 0; MaxLength = 100 }), defaultFormat, false
      "B1:B5", DataBarEx (cfvo CfvoNumber (Some "0"), cfvo CfvoFormula (Some "MAX($B$1:$B$5)*2"), "#63C384",
                          { Gradient = true; Border = Some "#2E7D32"; NegativeFill = None; NegativeBorder = Some "#990000"; Axis = AxisMiddle; AxisColor = Some "#7F7F7F"
                            Direction = BarRightToLeft; ShowValue = false; MinLength = 5; MaxLength = 95 }), defaultFormat, false
      "B1:B5", CellIs (OpGreaterThan, "3", None), { defaultFormat with BgColor = Some "#FFEB9C" }, false
      "C1:C5", IconSetEx ("3Triangles", [ cfvo CfvoPercent (Some "0"); cfvo CfvoNumber (Some "2"); cfvo CfvoFormula (Some "Lists!$B$1") ], false,
                          { ShowValue = false; Custom = [ "3Arrows", 0; "NoIcons", 0; "3Stars", 2 ] }), defaultFormat, false
      "E1:E5,G1:G3", Expression "Lists!$B$1>$A1", { defaultFormat with Bold = true; Italic = true; TextColor = Some "#9C0006"; BgColor = Some "#FFC7CE" }, true
      "E1:E5,G1:G3", CellIs (OpBetween, "Lists!$B$1", Some "Lists!$B$2"), { defaultFormat with BgColor = Some "#C6EFCE" }, false ]
let rangesText (ranges: CellRange list) =
    ranges |> List.map (fun r -> formatCellRef r.From.Row r.From.Col + ":" + formatCellRef r.To.Row r.To.Col) |> String.concat ","
let rulesOf (model: SpreadsheetModel) =
    model.ConditionalFormats |> List.sortBy (fun entry -> entry.Priority) |> List.map (fun entry -> rangesText entry.Ranges, entry.Format.Rule, entry.Format.Style, entry.StopIfTrue)
let validationsOf (model: SpreadsheetModel) =
    model.Validations |> List.map (fun v -> rangesText v.Ranges, v.Validation.Rule, v.Validation.ErrorTitle, v.Validation.ErrorMessage)
let expectedValidations =
    [ "D1:D5", Comparison (WholeKind, OpBetween, "1", Some "9"), None, None
      "H1:H5", ListSource "Lists!$A$1:$A$3", Some "Pick one", Some "Not in the list"
      "I1:I5,K2:K2", Comparison (DecimalKind, OpLessOrEqual, "Lists!$B$2", None), None, None ]

let file = packageOf ""
let imported = Bridge.Import.importWorkbook (loadBytes file)
let data = imported.Sheets.[0].Model
eq "read: the rules, in order of precedence" (rulesOf data) expectedRules
eq "read: the validations" (validationsOf data |> List.map (fun (r, rule, t, m) -> r, (match rule with WholeNumber (a, b) -> Comparison (WholeKind, OpBetween, string (int a.Value), b |> Option.map (fun x -> string (int x))) | other -> other), t, m)) expectedValidations
eq "read: the x14 rules are not also kept as they were" (imported.Source.Workbook.Sheets.[0].Extras.KeptCellExtensions |> List.map (fun ext -> ext.Contains "sparklineGroups", ext.Contains "conditionalFormattings")) [ true, false ]

// ---- writing ---------------------------------------------------------------------
let checkWritten (label: string) (bytes: byte[]) =
    let sheet = textOf (partsOf bytes) "xl/worksheets/sheet1.xml"
    let main = sheet.Substring(0, sheet.IndexOf "<extLst><ext uri=\"{78C0D931")
    let ext = sheet.Substring(main.Length)
    // The main part: two data bars that name their twins, and the ordinary rule.
    eq (label + ": main rules") (count "<cfRule type=\"dataBar\"" main, count "<cfRule type=\"cellIs\"" main, count "<cfRule " main) (2, 1, 3)
    let ids = [ for m in Text.RegularExpressions.Regex.Matches(main, "<x14:id>([^<]*)</x14:id>") -> m.Groups.[1].Value ]
    eq (label + ": each data bar names an x14 rule") (ids.Length, ids |> List.forall (fun id -> ext.Contains ("<x14:cfRule type=\"dataBar\" id=\"" + id + "\">")), (List.distinct ids).Length) (2, true, 2)
    ok (label + ": the 2007 bar has min / max ends") (main.Contains "<dataBar><cfvo type=\"min\"/><cfvo type=\"max\"/><color rgb=\"FF638EC6\"/></dataBar>") main
    ok (label + ": the first x14 bar") (ext.Contains "<x14:dataBar minLength=\"0\" maxLength=\"100\" gradient=\"0\"><x14:cfvo type=\"autoMin\"/><x14:cfvo type=\"autoMax\"/><x14:fillColor rgb=\"FF638EC6\"/><x14:negativeFillColor rgb=\"FFFF0000\"/><x14:axisColor rgb=\"FF000000\"/></x14:dataBar>") ext
    ok (label + ": the second x14 bar") (ext.Contains "<x14:dataBar minLength=\"5\" maxLength=\"95\" showValue=\"0\" border=\"1\" direction=\"rightToLeft\" negativeBarColorSameAsPositive=\"1\" negativeBarBorderColorSameAsPositive=\"0\" axisPosition=\"middle\"><x14:cfvo type=\"num\"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type=\"formula\"><xm:f>MAX($B$1:$B$5)*2</xm:f></x14:cfvo><x14:fillColor rgb=\"FF63C384\"/><x14:borderColor rgb=\"FF2E7D32\"/><x14:negativeBorderColor rgb=\"FF990000\"/><x14:axisColor rgb=\"FF7F7F7F\"/></x14:dataBar>") ext
    ok (label + ": the icon set") (ext.Contains "<x14:iconSet iconSet=\"3Triangles\" showValue=\"0\" custom=\"1\"><x14:cfvo type=\"percent\"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type=\"num\"><xm:f>2</xm:f></x14:cfvo><x14:cfvo type=\"formula\"><xm:f>Lists!$B$1</xm:f></x14:cfvo><x14:cfIcon iconSet=\"3Arrows\" iconId=\"0\"/><x14:cfIcon iconSet=\"NoIcons\" iconId=\"0\"/><x14:cfIcon iconSet=\"3Stars\" iconId=\"2\"/></x14:iconSet>") ext
    ok (label + ": the rule that reads another sheet") (ext.Contains "<xm:f>Lists!$B$1&gt;$A1</xm:f><x14:dxf><font><b/><i/><color rgb=\"FF9C0006\"/></font><fill><patternFill><bgColor rgb=\"FFFFC7CE\"/></patternFill></fill></x14:dxf></x14:cfRule>" && ext.Contains "stopIfTrue=\"1\"") ext
    ok (label + ": its range") (ext.Contains "<xm:sqref>E1:E5 G1:G3</xm:sqref>") ext
    eq (label + ": no formula that reads another sheet in the main part") (main.Contains "Lists!") false
    // Validations: the ordinary one in the main part, the two others in the extension.
    eq (label + ": validations") (count "<dataValidation " main, count "<x14:dataValidation " ext, ext.Contains "<x14:dataValidations count=\"2\"") (1, 2, true)
    ok (label + ": the list source") (ext.Contains "<x14:formula1><xm:f>Lists!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>H1:H5</xm:sqref>") ext
    ok (label + ": two ranges of a validation") (ext.Contains "<xm:sqref>I1:I5 K2</xm:sqref>") ext
    eq (label + ": one extension of each kind") (count "{78C0D931-6437-407d-A8EE-F0AAD7539E65}" sheet, count "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" sheet, count "<x14:sparklineGroups" sheet, count "<extLst>" ext) (1, 1, 1, 1)
    ok (label + ": well-formed") (try XmlDsl.parse sheet |> ignore; true with e -> (printfn "%s" e.Message; false)) ""
    let again = Bridge.Import.importWorkbook (loadBytes bytes)
    eq (label + ": read again, the rules") (rulesOf again.Sheets.[0].Model) expectedRules
    eq (label + ": read again, the validations") (validationsOf again.Sheets.[0].Model |> List.map (fun (r, _, t, m) -> r, t, m)) (expectedValidations |> List.map (fun (r, _, t, m) -> r, t, m))
    again

// Saved with nothing of the rules touched. (A table added: a sheet nothing at
// all was changed in goes back as the part it was read from, one with another
// zoom, a row height or a link as that part with the change in it, and this is
// about what the writer makes of the model.)
let zoomed (models: SpreadsheetModel array) =
    models
    |> Array.mapi (fun index model ->
        let top = 300 + 4 * model.Tables.Length
        createTable (sprintf "Added%d_%d" index model.Tables.Length) top 0 (top + 1) 1 false model)
let untouched = export imported (zoomed (modelsOf imported))
let second = checkWritten "save" untouched
checkWritten "second save" (export second (zoomed (modelsOf second))) |> ignore
// ... and with nothing touched at all: the sheet is the file's part.
ok "save: an untouched sheet is the part that was read" ((partsOf (export imported (modelsOf imported))).["xl/worksheets/sheet1.xml"] = (partsOf file).["xl/worksheets/sheet1.xml"]) ""
// The reader's workbook straight to the writer.
checkWritten "reader to writer" (XlsxWriter.writeWorkbookPackageBytes (XlsxWriter.packageFromReaderWorkbook (loadBytes file))) |> ignore

// The rules of the sheet edited (one added): the x14 rules are still there.
do
    let edited = addConditionalFormatRange [ range 0 9 4 9 ] { Rule = Expression "$J1>0"; Style = { defaultFormat with BgColor = Some "#DDEBF7" } } data
    let models = modelsOf imported
    models.[0] <- edited
    let bytes = export imported models
    let sheet = textOf (partsOf bytes) "xl/worksheets/sheet1.xml"
    eq "rules edited: x14 rules written" (count "<x14:cfRule type=\"dataBar\"" sheet, count "<x14:cfRule type=\"iconSet\"" sheet, count "<x14:cfRule type=\"expression\"" sheet, count "<x14:cfRule type=\"cellIs\"" sheet) (2, 1, 1, 1)
    let again = Bridge.Import.importWorkbook (loadBytes bytes)
    eq "rules edited: every rule read again" (rulesOf again.Sheets.[0].Model |> List.map (fun (r, rule, _, _) -> r, rule) |> List.sort) (("J1:J5", Expression "$J1>0") :: (expectedRules |> List.map (fun (r, rule, _, _) -> r, rule)) |> List.sort)
    // A validation edited: the cross-sheet ones are still in the extension.
    let fewer = { data with Validations = data.Validations |> List.filter (fun v -> rangesText v.Ranges <> "D1:D5") }
    models.[0] <- fewer
    let sheet = textOf (partsOf (export imported models)) "xl/worksheets/sheet1.xml"
    eq "validations edited" (count "<dataValidation " sheet, count "<x14:dataValidation " sheet) (0, 2)
    // A row inserted above: the x14 rules follow (they are rules of the model now).
    models.[0] <- applyStructureEdit (InsertRows (0, 2)) data
    let moved = Bridge.Import.importWorkbook (loadBytes (export imported models))
    eq "row inserted: the ranges moved" (rulesOf moved.Sheets.[0].Model |> List.map (fun (r, _, _, _) -> r)) [ "A3:A7"; "B3:B7"; "B3:B7"; "C3:C7"; "E3:E7,G3:G5"; "E3:E7,G3:G5" ]
    eq "row inserted: the cross-sheet validations moved" (validationsOf moved.Sheets.[0].Model |> List.map (fun (r, _, _, _) -> r)) [ "D3:D7"; "H3:H7"; "I3:I7,K4:K4" ]

// ---- what is drawn -------------------------------------------------------------
do
    let ctx = CF.createContext Rbt.Spreadsheet.Model.evaluateFormulaAt data
    let bar row col = CF.effectAt ctx (addr row col) |> Option.bind (fun e -> e.DataBar)
    eq "drawn: A1 (-50) left of the axis in red, solid" (bar 0 0 |> Option.map (fun b -> Math.Round(b.From, 3), Math.Round(b.To, 3), b.Color, b.Gradient, b.Axis |> Option.map (fun (x, c) -> Math.Round(x, 3), c))) (Some (0.0, 0.333, "#FF0000", false, Some (0.333, "#000000")))
    eq "drawn: A5 (100) right of the axis" (bar 4 0 |> Option.map (fun b -> Math.Round(b.From, 3), Math.Round(b.To, 3), b.Color)) (Some (0.333, 1.0, "#638EC6"))
    eq "drawn: B5 right to left from the middle, bordered, no value" (bar 4 1 |> Option.map (fun b -> Math.Round(b.From, 3), Math.Round(b.To, 3), b.Border, b.HideValue, b.Gradient)) (Some (0.25, 0.5, Some "#2E7D32", true, true))

// ---- opaque rules survive edits alongside supported rules -------------------------
do
    let pivotRule = "<x14:conditionalFormatting pivot=\"1\" xmlns:z=\"urn:rbt:future\" z:group=\"keep\" xmlns:xm=\"" + xm + "\"><x14:cfRule type=\"futureKind\" priority=\"9\" id=\"{BBBBBBBB-0000-4000-8000-000000000001}\"><x14:somethingNew level=\"2\"/></x14:cfRule><xm:sqref>M1:M5</xm:sqref></x14:conditionalFormatting>"
    let source = packageOf pivotRule
    let futureGroup (sheet: string) =
        let doc = System.Xml.Linq.XDocument.Parse sheet
        let future =
            doc.Descendants(System.Xml.Linq.XName.Get("cfRule", x14))
            |> Seq.filter (fun rule ->
                let kind = rule.Attribute(System.Xml.Linq.XName.Get "type")
                not (isNull kind) && kind.Value = "futureKind")
            |> Seq.exactlyOne
        let group = System.Xml.Linq.XElement(future.Parent)
        // Compare expanded names and data, not prefix declarations or attribute order.
        for node in group.DescendantsAndSelf() do
            for attribute in node.Attributes() |> Seq.filter _.IsNamespaceDeclaration |> Seq.toArray do
                attribute.Remove()
        // Inserting a higher-precedence rule legitimately renumbers priorities.
        for rule in group.Descendants(System.Xml.Linq.XName.Get("cfRule", x14)) do
            let priority = rule.Attribute(System.Xml.Linq.XName.Get "priority")
            if not (isNull priority) then priority.Remove()
        group, int (future.Attribute(System.Xml.Linq.XName.Get "priority").Value)
    let originalGroup, _ = futureGroup (textOf (partsOf source) "xl/worksheets/sheet1.xml")
    let odd = Bridge.Import.importWorkbook (loadBytes source)
    let unchanged = textOf (partsOf (export odd (modelsOf odd))) "xl/worksheets/sheet1.xml"
    eq "opaque rule and pivot range survive an unchanged save"
        (System.Xml.Linq.XNode.DeepEquals(originalGroup, futureGroup unchanged |> fst)) true
    let models = modelsOf odd
    let added = { Rule = Expression "Lists!$B$1>0"; Style = { defaultFormat with BgColor = Some "#DDEBF7" } }
    models.[0] <- addConditionalFormatRange [ range 0 9 4 9 ] added models.[0]
    let saved = export odd models
    let edited = textOf (partsOf saved) "xl/worksheets/sheet1.xml"
    let editedGroup, opaquePriority = futureGroup edited
    ok "opaque rule identity, payload and pivot range survive a supported rule edit"
        (System.Xml.Linq.XNode.DeepEquals(originalGroup, editedGroup))
        (sprintf "%O -> %O" originalGroup editedGroup)
    let reopened = Bridge.Import.importWorkbook (loadBytes saved)
    eq "supported rules and the new cross-sheet rule survive beside an opaque rule"
        (rulesOf reopened.Sheets.[0].Model)
        (("J1:J5", added.Rule, added.Style, false) :: expectedRules)
    let knownPriority = reopened.Sheets.[0].Model.ConditionalFormats |> List.map _.Priority |> List.max
    ok "opaque rule remains below the supported rules in precedence"
        (opaquePriority > knownPriority) (sprintf "opaque=%d known=%d" opaquePriority knownPriority)

// ---- a real workbook ---------------------------------------------------------------
do
    let original = File.ReadAllBytes(Path.Combine(root, "real-world/dashboard-sales-analysis.xlsx"))
    let real = Bridge.Import.importWorkbook (loadBytes original)
    let first = real.Sheets.[0].Model
    let bars = first.ConditionalFormats |> List.choose (fun entry -> match entry.Format.Rule with DataBarEx (low, high, _, options) -> Some (rangesText entry.Ranges, low.Kind, high.Kind, options.Gradient, options.NegativeFill, options.Axis) | _ -> None)
    eq "dashboard: its data bar with the x14 options" bars [ "E15:E22", CfvoAutoMin, CfvoAutoMax, false, Some "#FF0000", AxisAutomatic ]
    // An edit elsewhere, a save: the bar is written with its twin; the sparklines stay.
    let models = modelsOf real
    models.[0] <- setCell (addr 40 0) { Value = CellValue.String "edited"; Format = defaultFormat; Comment = None } first
    let sheet = textOf (partsOf (export real models)) "xl/worksheets/sheet1.xml"
    eq "dashboard: saved" (count "<x14:cfRule type=\"dataBar\"" sheet, count "<x14:id>" sheet, count "<x14:sparklineGroup " sheet, sheet.Contains "<x14:dataBar minLength=\"0\" maxLength=\"100\" gradient=\"0\"><x14:cfvo type=\"autoMin\"/><x14:cfvo type=\"autoMax\"/>") (1, 1, 3, true)
    let back = Bridge.Import.importWorkbook (loadBytes (export real models))
    eq "dashboard: read again" (back.Sheets.[0].Model.ConditionalFormats |> List.map (fun entry -> entry.Ranges, entry.Format)) (first.ConditionalFormats |> List.map (fun entry -> entry.Ranges, entry.Format))

printfn ""
printfn "XLSX x14 tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
