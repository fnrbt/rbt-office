#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

// Dynamic array formulas in XLSX: `<f t="array" ref>` on a cell whose `cm`
// names the XLDAPR record of xl/metadata.xml, as Excel writes them; legacy
// (Ctrl+Shift+Enter) array formulas, which have no such metadata, stay what
// they are; a metadata part that holds other things is added to, not replaced.

open System
open System.IO
open System.IO.Compression
open Rbt.Spreadsheet.Model
module Bridge = Rbt.Spreadsheet.XlsxBridge

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-34s  %s" label detail)
let eq label (got: 'a) (expected: 'a) =
    ok label (got = expected) (sprintf "expected %A  got %A" expected got)

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
let export (source: Bridge.ImportedWorkbook option) (sheets: SpreadsheetModel array) (names: string array) =
    XlsxWriter.writeWorkbookPackageBytes
        (Bridge.Export.toWorkbookOut
            { Sheets = sheets
              Names = names
              Visible = Array.create sheets.Length true
              ActiveSheetIndex = 0
              NameInfo = Map.empty
              DefinedNames = (match source with Some imported -> imported.Context.Names | None -> Map.empty)
              Source = source |> Option.map (fun imported -> imported.Source)
              Macros = None })
/// The start tag of a cell and of its formula: `<c r="C1" …><f …>`.
let cellTags (sheetXml: string) (cell: string) =
    let m = Text.RegularExpressions.Regex.Match(sheetXml, sprintf "<c r=\"%s\"[^>]*>(<f[^>]*>)?" cell)
    if m.Success then m.Value else ""

// ---- the package as Excel writes it ------------------------------------------
let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
let decl = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
let excelMetadata =
    decl + "<metadata xmlns=\"" + ns + "\" xmlns:xda=\"http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray\">"
    + "<metadataTypes count=\"1\"><metadataType name=\"XLDAPR\" minSupportedVersion=\"120000\" copy=\"1\" pasteAll=\"1\" pasteValues=\"1\" merge=\"1\" splitFirst=\"1\" rowColShift=\"1\" clearFormats=\"1\" clearComments=\"1\" assign=\"1\" coerce=\"1\" cellMeta=\"1\"/></metadataTypes>"
    + "<futureMetadata name=\"XLDAPR\" count=\"1\"><bk><extLst><ext uri=\"{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}\"><xda:dynamicArrayProperties fDynamic=\"1\" fCollapsed=\"0\"/></ext></extLst></bk></futureMetadata>"
    + "<cellMetadata count=\"1\"><bk><rc t=\"1\" v=\"0\"/></bk></cellMetadata></metadata>"
/// Rich values only (a picture in a cell): no XLDAPR type, no cell metadata.
let richMetadata =
    decl + "<metadata xmlns=\"" + ns + "\" xmlns:xlrd=\"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata\">"
    + "<metadataTypes count=\"1\"><metadataType name=\"XLRICHVALUE\" minSupportedVersion=\"120000\" copy=\"1\" pasteAll=\"1\" pasteValues=\"1\" merge=\"1\" splitFirst=\"1\" rowColShift=\"1\" clearFormats=\"1\" clearComments=\"1\" assign=\"1\" coerce=\"1\"/></metadataTypes>"
    + "<futureMetadata name=\"XLRICHVALUE\" count=\"1\"><bk><extLst><ext uri=\"{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}\"><xlrd:rvb i=\"0\"/></ext></extLst></bk></futureMetadata>"
    + "<valueMetadata count=\"1\"><bk><rc t=\"1\" v=\"0\"/></bk></valueMetadata></metadata>"
let richValueXml =
    decl + "<rvData xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata\" count=\"1\"><rv s=\"0\"><v>0</v><v>5</v></rv></rvData>"

let packageOf (sheetData: string) (metadata: string option) (rich: bool) =
    let types =
        decl + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
        + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/>"
        + "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
        + "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
        + (if metadata.IsSome then "<Override PartName=\"/xl/metadata.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml\"/>" else "")
        + (if rich then "<Override PartName=\"/xl/richData/rdrichvalue.xml\" ContentType=\"application/vnd.ms-excel.rdrichvalue+xml\"/>" else "")
        + "</Types>"
    let workbookRels =
        decl + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
        + (if metadata.IsSome then "<Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata\" Target=\"metadata.xml\"/>" else "")
        + (if rich then "<Relationship Id=\"rId5\" Type=\"http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue\" Target=\"richData/rdrichvalue.xml\"/>" else "")
        + "</Relationships>"
    zipOf
        [ yield "[Content_Types].xml", types
          yield "_rels/.rels", decl + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>"
          yield "xl/workbook.xml", decl + "<workbook xmlns=\"" + ns + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>"
          yield "xl/_rels/workbook.xml.rels", workbookRels
          yield "xl/worksheets/sheet1.xml", decl + "<worksheet xmlns=\"" + ns + "\"><sheetData>" + sheetData + "</sheetData></worksheet>"
          match metadata with Some xml -> yield "xl/metadata.xml", xml | None -> ()
          if rich then yield "xl/richData/rdrichvalue.xml", richValueXml ]

/// A1:A3 = 3, 1, 2; C1 a dynamic SORT spilled over C1:C3; E1 a legacy array over
/// E1:E3; G1 a legacy single-cell array; H1 a dynamic array that fills one cell.
let excelRows (dynamicCm: string) =
    "<row r=\"1\"><c r=\"A1\"><v>3</v></c>"
    + "<c r=\"C1\"" + dynamicCm + "><f t=\"array\" ref=\"C1:C3\">_xlfn._xlws.SORT(A1:A3)</f><v>1</v></c>"
    + "<c r=\"E1\"><f t=\"array\" ref=\"E1:E3\">A1:A3*2</f><v>6</v></c>"
    + "<c r=\"G1\"><f t=\"array\" ref=\"G1\">SUM(A1:A3*A1:A3)</f><v>14</v></c>"
    + "<c r=\"H1\"" + dynamicCm + "><f t=\"array\" ref=\"H1\">_xlfn.SEQUENCE(1)</f><v>1</v></c></row>"
    + "<row r=\"2\"><c r=\"A2\"><v>1</v></c><c r=\"C2\"><v>2</v></c><c r=\"E2\"><v>2</v></c></row>"
    + "<row r=\"3\"><c r=\"A3\"><v>2</v></c><c r=\"C3\"><v>3</v></c><c r=\"E3\"><v>4</v></c></row>"

// ---- the metadata part is understood ----------------------------------------
eq "metadata: the dynamic cell-metadata record" (XlsxReader.dynamicArrayCellMetadata excelMetadata) (Set.ofList [ 1 ])
eq "metadata: rich values are not dynamic arrays" (XlsxReader.dynamicArrayCellMetadata richMetadata) Set.empty

// ---- reading what Excel wrote ------------------------------------------------
let excelFile = packageOf (excelRows " cm=\"1\"") (Some excelMetadata) false
let checkExcelRead (label: string) (bytes: byte[]) =
    let wb = loadBytes bytes
    let cells = wb.Sheets.[0].Rows |> List.collect (fun row -> row.Cells) |> List.map (fun cell -> cell.Reference.Column + string cell.Reference.Row, cell) |> Map.ofList
    let arrayOf (cell: string) = cells.[cell].ArrayRef, cells.[cell].ArrayDynamic
    eq (label + ": C1 is a dynamic array") (arrayOf "C1") (Some "C1:C3", true)
    eq (label + ": E1 is a legacy array") (arrayOf "E1") (Some "E1:E3", false)
    eq (label + ": G1 is a legacy array in one cell") (arrayOf "G1") (Some "G1", false)
    eq (label + ": H1 is a dynamic array in one cell") (arrayOf "H1") (Some "H1", true)
    wb
let excelWorkbook = checkExcelRead "read" excelFile
let imported = Bridge.Import.importWorkbook excelWorkbook
let model = imported.Sheets.[0].Model
eq "import: the marks" (model.ArrayMarks |> Map.toList |> List.map (fun (cell, mark) -> cell.Col, mark.Legacy, mark.Rows, mark.Cols)) [ 2, false, 3, 1; 4, true, 3, 1; 6, true, 1, 1; 7, false, 1, 1 ]
ok "import: the spilled cells belong to their formula" (Map.tryFind (addr 2 2) model.Spill = Some (addr 0 2) && Map.tryFind (addr 1 4) model.Spill = Some (addr 0 4)) (sprintf "%A" model.Spill)

// ---- writing it again --------------------------------------------------------
let checkExcelWritten (label: string) (bytes: byte[]) =
    let parts = partsOf bytes
    let sheet = textOf parts "xl/worksheets/sheet1.xml"
    ok (label + ": C1 written dynamic") ((cellTags sheet "C1").Contains "cm=\"1\"" && (cellTags sheet "C1").Contains "<f t=\"array\" ref=\"C1:C3\">") (cellTags sheet "C1")
    ok (label + ": E1 written legacy") (not ((cellTags sheet "E1").Contains "cm=") && (cellTags sheet "E1").Contains "<f t=\"array\" ref=\"E1:E3\">") (cellTags sheet "E1")
    ok (label + ": G1 written legacy") (not ((cellTags sheet "G1").Contains "cm=") && (cellTags sheet "G1").Contains "<f t=\"array\" ref=\"G1\">") (cellTags sheet "G1")
    ok (label + ": H1 written dynamic") ((cellTags sheet "H1").Contains "cm=\"1\"" && (cellTags sheet "H1").Contains "<f t=\"array\" ref=\"H1\">") (cellTags sheet "H1")
    eq (label + ": two cells name the metadata") (count " cm=\"" sheet) 2
    eq (label + ": one metadata relationship") (count "/sheetMetadata\"" (textOf parts "xl/_rels/workbook.xml.rels")) 1
    eq (label + ": one metadata content type") (count "PartName=\"/xl/metadata.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml\"" (textOf parts "[Content_Types].xml")) 1
    parts
let written = export (Some imported) [| model |] [| "Sheet1" |]
let writtenParts = checkExcelWritten "save" written
eq "save: Excel's metadata part is kept as it was" (textOf writtenParts "xl/metadata.xml") excelMetadata
checkExcelRead "read again" written |> ignore
// A second save of the saved file is as good.
let importedAgain = Bridge.Import.importWorkbook (loadBytes written)
checkExcelWritten "second save" (export (Some importedAgain) [| importedAgain.Sheets.[0].Model |] [| "Sheet1" |]) |> ignore

// The reader -> writer path that does not go through the model.
let direct = XlsxWriter.writeWorkbookPackageBytes (XlsxWriter.packageFromReaderWorkbook excelWorkbook)
checkExcelWritten "reader to writer" direct |> ignore

// ---- edits --------------------------------------------------------------------
// A row inserted above: the marks follow their formulas.
let shifted = applyStructureEdit (InsertRows (0, 1)) model
eq "insert row: the marks moved" (shifted.ArrayMarks |> Map.toList |> List.map (fun (cell, mark) -> cell.Row, cell.Col, mark.Legacy)) [ 1, 2, false; 1, 4, true; 1, 6, true; 1, 7, false ]
let shiftedSheet = textOf (partsOf (export (Some imported) [| shifted |] [| "Sheet1" |])) "xl/worksheets/sheet1.xml"
ok "insert row: the legacy array is still one" ((cellTags shiftedSheet "E2").Contains "<f t=\"array\" ref=\"E2:E4\">" && not ((cellTags shiftedSheet "E2").Contains "cm=")) (cellTags shiftedSheet "E2")
ok "insert row: the dynamic array is still one" ((cellTags shiftedSheet "C2").Contains "cm=\"1\"") (cellTags shiftedSheet "C2")
// Another formula typed over the legacy array: it is not a legacy array.
let retyped = setCell (addr 0 6) { Value = CellValue.Formula ("=SUM(A1:A3)", CellValue.Number 6.0); Format = defaultFormat; Comment = None } model
let retypedSheet = textOf (partsOf (export (Some imported) [| retyped |] [| "Sheet1" |])) "xl/worksheets/sheet1.xml"
ok "retyped: an ordinary formula" ((cellTags retypedSheet "G1").EndsWith "<f>") (cellTags retypedSheet "G1")

// ---- a workbook made here -----------------------------------------------------
let fresh =
    createEmpty 20 10
    |> setCell (addr 0 0) { Value = CellValue.Formula ("=SEQUENCE(3)", CellValue.Empty); Format = defaultFormat; Comment = None }
    |> setCell (addr 0 2) { Value = CellValue.Formula ("=1+1", CellValue.Empty); Format = defaultFormat; Comment = None }
    |> Rbt.Spreadsheet.Formula.recalcAndSpill None (Some "Sheet1")
eq "new: the formula spilled" (fresh.Spill |> Map.toList |> List.map (fun (cell, _) -> cell.Row)) [ 1; 2 ]
let freshBytes = export None [| fresh |] [| "Sheet1" |]
let freshParts = partsOf freshBytes
let freshSheet = textOf freshParts "xl/worksheets/sheet1.xml"
ok "new: written as a dynamic array" ((cellTags freshSheet "A1").Contains "cm=\"1\"" && (cellTags freshSheet "A1").Contains "<f t=\"array\" ref=\"A1:A3\">") (cellTags freshSheet "A1")
ok "new: an ordinary formula is ordinary" ((cellTags freshSheet "C1").EndsWith "<f>") (cellTags freshSheet "C1")
let freshMetadata = textOf freshParts "xl/metadata.xml"
ok "new: the metadata part" (freshMetadata.Contains "<futureMetadata name=\"XLDAPR\" count=\"1\">" && freshMetadata.Contains "xda:dynamicArrayProperties fDynamic=\"1\" fCollapsed=\"0\"" && freshMetadata.Contains "<cellMetadata count=\"1\"><bk><rc t=\"1\" v=\"0\"/></bk></cellMetadata>") freshMetadata
eq "new: the metadata part is what Excel writes" (freshMetadata.Replace("\r\n", "\n")) (excelMetadata.Replace("\r\n", "\n"))
eq "new: its relationship" (count "relationships/sheetMetadata\" Target=\"metadata.xml\"" (textOf freshParts "xl/_rels/workbook.xml.rels")) 1
eq "new: its content type" (count "PartName=\"/xl/metadata.xml\"" (textOf freshParts "[Content_Types].xml")) 1
let freshBack = Bridge.Import.importWorkbook (loadBytes freshBytes)
eq "new: read back as a dynamic array" (freshBack.Sheets.[0].Model.ArrayMarks |> Map.toList |> List.map (fun (cell, mark) -> cell.Row, cell.Col, mark.Legacy, mark.Rows)) [ 0, 0, false, 3 ]
// A workbook without dynamic arrays has no metadata part.
let plainParts = partsOf (export None [| createEmpty 5 5 |> setCell (addr 0 0) { Value = CellValue.Number 1.0; Format = defaultFormat; Comment = None } |] [| "Sheet1" |])
ok "plain: no metadata part" (not (Map.containsKey "xl/metadata.xml" plainParts) && not ((textOf plainParts "xl/_rels/workbook.xml.rels").Contains "sheetMetadata")) "metadata written"

// ---- legacy arrays in a file without metadata ---------------------------------
do
    let legacyFile = packageOf (excelRows "") None false
    let legacy = Bridge.Import.importWorkbook (loadBytes legacyFile)
    eq "legacy file: every array is legacy" (legacy.Sheets.[0].Model.ArrayMarks |> Map.toList |> List.map (fun (_, mark) -> mark.Legacy)) [ true; true; true; true ]
    let parts = partsOf (export (Some legacy) [| legacy.Sheets.[0].Model |] [| "Sheet1" |])
    let sheet = textOf parts "xl/worksheets/sheet1.xml"
    eq "legacy file: no cell names metadata" (count " cm=\"" sheet) 0
    eq "legacy file: four array formulas" (count "<f t=\"array\"" sheet) 4
    ok "legacy file: no metadata part" (not (Map.containsKey "xl/metadata.xml" parts)) "metadata written"

// ---- a metadata part with rich values gets the dynamic array record ------------
do
    let extended, cm = XlsxWriter.metadataWithDynamicArrays richMetadata
    eq "extend: the record's number" cm 1
    eq "extend: the reader finds it" (XlsxReader.dynamicArrayCellMetadata extended) (Set.ofList [ 1 ])
    ok "extend: the rich value type is kept, XLDAPR is second" (extended.Contains "<metadataTypes count=\"2\"><metadataType name=\"XLRICHVALUE\"" && extended.Contains "<rc t=\"2\" v=\"0\"/>") extended
    ok "extend: rich value metadata kept" (extended.Contains "<futureMetadata name=\"XLRICHVALUE\" count=\"1\"><bk><extLst><ext uri=\"{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}\"><xlrd:rvb i=\"0\"/></ext></extLst></bk></futureMetadata>" && extended.Contains "<valueMetadata count=\"1\"><bk><rc t=\"1\" v=\"0\"/></bk></valueMetadata>") extended
    // CT_Metadata's order: types, future metadata, cellMetadata, valueMetadata.
    let at (token: string) = extended.IndexOf token
    ok "extend: schema order" (at "<metadataTypes" < at "<futureMetadata name=\"XLRICHVALUE\"" && at "<futureMetadata name=\"XLRICHVALUE\"" < at "<futureMetadata name=\"XLDAPR\"" && at "<futureMetadata name=\"XLDAPR\"" < at "<cellMetadata" && at "<cellMetadata" < at "<valueMetadata") extended
    ok "extend: well-formed" (try XmlDsl.parse extended |> ignore; true with _ -> false) extended
    // Extending what is already extended changes nothing.
    eq "extend: once" (XlsxWriter.metadataWithDynamicArrays extended) (extended, 1)
    // Cell metadata that is something else: the record is added after it.
    let other = richMetadata.Replace("<valueMetadata", "<cellMetadata count=\"1\"><bk><rc t=\"1\" v=\"0\"/></bk></cellMetadata><valueMetadata")
    let otherExtended, otherCm = XlsxWriter.metadataWithDynamicArrays other
    eq "extend: after another cell-metadata record" (otherCm, otherExtended.Contains "<cellMetadata count=\"2\"><bk><rc t=\"1\" v=\"0\"/></bk><bk><rc t=\"2\" v=\"0\"/></bk></cellMetadata>") (2, true)
    eq "extend: the reader finds the second record" (XlsxReader.dynamicArrayCellMetadata otherExtended) (Set.ofList [ 2 ])

    // In a package: the rich value parts and the metadata are carried, a new
    // spilling formula names the added record.
    let richFile = packageOf "<row r=\"1\"><c r=\"A1\" t=\"e\" vm=\"1\"><v>#VALUE!</v></c><c r=\"B1\"><v>2</v></c></row>" (Some richMetadata) true
    let rich = Bridge.Import.importWorkbook (loadBytes richFile)
    let untouched = partsOf (export (Some rich) [| rich.Sheets.[0].Model |] [| "Sheet1" |])
    eq "rich: metadata kept byte for byte without dynamic arrays" (textOf untouched "xl/metadata.xml") richMetadata
    eq "rich: the rich value part is kept" (textOf untouched "xl/richData/rdrichvalue.xml") richValueXml
    let spilled =
        rich.Sheets.[0].Model
        |> setCell (addr 2 0) { Value = CellValue.Formula ("=SEQUENCE(2)", CellValue.Empty); Format = defaultFormat; Comment = None }
        |> Rbt.Spreadsheet.Formula.recalcAndSpill None (Some "Sheet1")
    let parts = partsOf (export (Some rich) [| spilled |] [| "Sheet1" |])
    let metadata = textOf parts "xl/metadata.xml"
    eq "rich + spill: the metadata was extended" metadata extended
    ok "rich + spill: the formula names the record" ((cellTags (textOf parts "xl/worksheets/sheet1.xml") "A3").Contains "cm=\"1\"") (cellTags (textOf parts "xl/worksheets/sheet1.xml") "A3")
    eq "rich + spill: one metadata relationship, one content type" (count "/sheetMetadata\"" (textOf parts "xl/_rels/workbook.xml.rels"), count "PartName=\"/xl/metadata.xml\"" (textOf parts "[Content_Types].xml")) (1, 1)
    eq "rich + spill: the rich value part and its relationship" (textOf parts "xl/richData/rdrichvalue.xml", count "/rdRichValue\"" (textOf parts "xl/_rels/workbook.xml.rels"), count "PartName=\"/xl/richData/rdrichvalue.xml\"" (textOf parts "[Content_Types].xml")) (richValueXml, 1, 1)
    let back = Bridge.Import.importWorkbook (loadBytes (export (Some rich) [| spilled |] [| "Sheet1" |]))
    eq "rich + spill: read back dynamic" (back.Sheets.[0].Model.ArrayMarks |> Map.toList |> List.map (fun (cell, mark) -> cell.Row, mark.Legacy, mark.Rows)) [ 2, false, 2 ]

printfn "xlsx-dynamic-arrays: %d passed, %d failed" pass fail
if fail > 0 then exit 1
