#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Vb.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

// A macro-enabled workbook's VBA project survives a save: byte for byte, with
// its signature parts, when its macros were not edited (the browser, which
// holds no macro source, and the native host, whose source is the project's);
// written as a new project, without the stale signature, when they were.

open System.IO
open System.IO.Compression
module Bridge = Rbt.Spreadsheet.XlsxBridge

let mutable pass = 0
let mutable fail = 0
let check label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else fail <- fail + 1; printfn "FAIL %-55s expected %A got %A" label expected got

let parts (bytes: byte[]) =
    use archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read)
    [ for e in archive.Entries -> e.FullName, (use s = e.Open() in use m = new MemoryStream() in s.CopyTo m; m.ToArray()) ] |> Map.ofList
let text (bytes: byte[]) = System.Text.Encoding.UTF8.GetString bytes

let withParts (bytes: byte[]) (edit: Map<string, byte[]> -> Map<string, byte[]>) =
    use output = new MemoryStream()
    (use archive = new ZipArchive(output, ZipArchiveMode.Create, true)
     for KeyValue(name, data) in edit (parts bytes) do
        use s = archive.CreateEntry(name).Open()
        s.Write(data, 0, data.Length))
    output.ToArray()

// An .xlsm with two modules' worth of source in one project, and a signature as Excel writes one.
let source = "Sub Main()\n  Range(\"A1\").Value = 1\nEnd Sub\n\nFunction Twice(x)\n  Twice = x * 2\nEnd Function"
let signature = Array.init 64 byte
let xlsm =
    let injected = VbaProject.injectIntoXlsx (File.ReadAllBytes (Path.Combine(__SOURCE_DIRECTORY__, "../real-world/dashboard-sales-analysis.xlsx"))) "Module1" source
    withParts injected (fun p ->
        p
        |> Map.add "xl/vbaProjectSignature.bin" signature
        |> Map.add "xl/_rels/vbaProject.bin.rels"
            (System.Text.Encoding.UTF8.GetBytes "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature\" Target=\"vbaProjectSignature.bin\"/></Relationships>")
        // The code names that tie ThisWorkbook and the sheet modules to their parts.
        |> Map.change "xl/workbook.xml" (Option.map (fun data ->
            (text data).Replace("<workbookPr hidePivotFieldList=\"1\"/>", "<workbookPr codeName=\"ThisWorkbook\" hidePivotFieldList=\"1\"/>") |> System.Text.Encoding.UTF8.GetBytes))
        |> Map.change "xl/worksheets/sheet1.xml" (Option.map (fun data ->
            let xml = text data
            let at = xml.IndexOf("<dimension")
            (xml.Substring(0, at) + "<sheetPr codeName=\"Sheet1\"/>" + xml.Substring at) |> System.Text.Encoding.UTF8.GetBytes))
        |> Map.change "[Content_Types].xml" (Option.map (fun data ->
            (text data).Replace("</Types>", "<Override PartName=\"/xl/vbaProjectSignature.bin\" ContentType=\"application/vnd.ms-office.vbaProjectSignature\"/></Types>")
            |> System.Text.Encoding.UTF8.GetBytes)))
let original = parts xlsm

let workbook = XlsxReader.loadBytesNative xlsm |> Option.get
let imported = Bridge.Import.importWorkbookWith Bridge.editorImportOptions workbook
let save macros =
    let input : Bridge.ExportInput =
        { Sheets = imported.Sheets |> Array.map (fun s -> s.Model); Names = imported.Sheets |> Array.map (fun s -> s.Name)
          Visible = Array.create imported.Sheets.Length true; ActiveSheetIndex = 0; NameInfo = Map.empty
          DefinedNames = imported.Context.Names; Source = Some imported.Source; Macros = macros }
    parts (XlsxWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut input))

let kept label (saved: Map<string, byte[]>) =
    let contentTypes = text saved.["[Content_Types].xml"]
    let rels = text saved.["xl/_rels/workbook.xml.rels"]
    check (label + ": the project, byte for byte") (saved.TryFind "xl/vbaProject.bin") (original.TryFind "xl/vbaProject.bin")
    check (label + ": its signature") (saved.TryFind "xl/vbaProjectSignature.bin") (Some signature)
    check (label + ": the project's relationships") (saved.TryFind "xl/_rels/vbaProject.bin.rels") (original.TryFind "xl/_rels/vbaProject.bin.rels")
    check (label + ": one relationship to the project") ((rels.Split("vbaProject\"").Length - 1)) 1
    check (label + ": macro-enabled workbook") (contentTypes.Contains "application/vnd.ms-excel.sheet.macroEnabled.main+xml") true
    check (label + ": the project typed as one") (contentTypes.Contains "<Override PartName=\"/xl/vbaProject.bin\" ContentType=\"application/vnd.ms-office.vbaProject\"/>") true
    check (label + ": typed once") (contentTypes.Split("/xl/vbaProject.bin\"").Length - 1) 1
    check (label + ": the workbook's code name") ((text saved.["xl/workbook.xml"]).Contains "codeName=\"ThisWorkbook\"") true
    check (label + ": the sheet's code name") ((text saved.["xl/worksheets/sheet1.xml"]).Contains "codeName=\"Sheet1\"") true
    check (label + ": signature typed") (contentTypes.Contains "/xl/vbaProjectSignature.bin") true

// The browser holds no macro source; the native host holds the project's.
kept "no source (browser)" (save None)
kept "the project's own source (native)" (save (Some (VbaProject.readMacroSource original.["xl/vbaProject.bin"])))
kept "the same source with CRLF and spacing" (save (Some ("\r\n" + (VbaProject.readMacroSource original.["xl/vbaProject.bin"]).Replace("\n", "\r\n") + "\r\n")))

// Edited macros: a new project of the edited source; the signature, which signed the old one, is gone.
let edited = save (Some "Sub Main()\n  Range(\"A1\").Value = 2\nEnd Sub")
let editedRels = text edited.["xl/_rels/workbook.xml.rels"]
check "edited: a new project" (edited.TryFind "xl/vbaProject.bin" <> original.TryFind "xl/vbaProject.bin" && edited.ContainsKey "xl/vbaProject.bin") true
check "edited: holding the edited source" (VbaProject.readMacroSource edited.["xl/vbaProject.bin"]) "Sub Main()\n  Range(\"A1\").Value = 2\nEnd Sub"
check "edited: no stale signature" (edited.ContainsKey "xl/vbaProjectSignature.bin", edited.ContainsKey "xl/_rels/vbaProject.bin.rels") (false, false)
check "edited: one relationship to the project" ((editedRels.Split("vbaProject\"").Length - 1)) 1
check "edited: macro-enabled workbook" ((text edited.["[Content_Types].xml"]).Contains "macroEnabled.main+xml") true
check "edited: code names kept" ((text edited.["xl/workbook.xml"]).Contains "codeName=\"ThisWorkbook\"", (text edited.["xl/worksheets/sheet1.xml"]).Contains "codeName=\"Sheet1\"") (true, true)
check "edited: no signature content type" ((text edited.["[Content_Types].xml"]).Contains "vbaProjectSignature") false
check "edited: the new project typed once" ((text edited.["[Content_Types].xml"]).Split("/xl/vbaProject.bin\"").Length - 1) 1
check "edited: printer settings keep the .bin default" ((text edited.["[Content_Types].xml"]).Contains "Extension=\"bin\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings\"") true

// A sheet edited and written from the model keeps its code name.
let editedSheet =
    let model = imported.Sheets.[0].Model
    let cell : Rbt.Spreadsheet.Model.CellData = { Value = Rbt.Spreadsheet.Model.CellValue.Number 42.0; Format = Rbt.Spreadsheet.Model.defaultFormat; Comment = None }
    let input : Bridge.ExportInput =
        { Sheets = imported.Sheets |> Array.mapi (fun i s -> if i = 0 then Rbt.Spreadsheet.Model.setCell { Row = 40; Col = 0 } cell model else s.Model)
          Names = imported.Sheets |> Array.map (fun s -> s.Name)
          Visible = Array.create imported.Sheets.Length true; ActiveSheetIndex = 0; NameInfo = Map.empty
          DefinedNames = imported.Context.Names; Source = Some imported.Source; Macros = None }
    parts (XlsxWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut input))
check "edited sheet: the cell written" ((text editedSheet.["xl/worksheets/sheet1.xml"]).Contains "r=\"A41\"") true
check "edited sheet: its code name kept" ((text editedSheet.["xl/worksheets/sheet1.xml"]).Contains "codeName=\"Sheet1\"") true
check "edited sheet: the project kept" (editedSheet.TryFind "xl/vbaProject.bin") (original.TryFind "xl/vbaProject.bin")

// A plain workbook stays plain.
let plain =
    let input : Bridge.ExportInput =
        let wb = XlsxReader.loadBytesNative (File.ReadAllBytes (Path.Combine(__SOURCE_DIRECTORY__, "../real-world/dashboard-sales-analysis.xlsx"))) |> Option.get
        let imp = Bridge.Import.importWorkbookWith Bridge.editorImportOptions wb
        { Sheets = imp.Sheets |> Array.map (fun s -> s.Model); Names = imp.Sheets |> Array.map (fun s -> s.Name)
          Visible = Array.create imp.Sheets.Length true; ActiveSheetIndex = 0; NameInfo = Map.empty
          DefinedNames = imp.Context.Names; Source = Some imp.Source; Macros = None }
    parts (XlsxWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut input))
check "plain workbook: no project" (plain.ContainsKey "xl/vbaProject.bin") false
check "plain workbook: not macro-enabled" ((text plain.["[Content_Types].xml"]).Contains "macroEnabled") false

printfn "xlsm-vba-project: %d passed, %d failed" pass fail
if fail > 0 then exit 1
