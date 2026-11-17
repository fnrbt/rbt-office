#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

// A workbook in the 1904 date system stays in it: the serials in the cells
// are the file's, dates are made of them in that system, and a save says
// `date1904` and writes untouched sheets as the parts they were read from.
//
// Usage: dotnet fsi tests/office-formats/xlsx-date1904.fsx

open System
open System.IO
open System.IO.Compression
module Bridge = Rbt.Spreadsheet.XlsxBridge
module M = Rbt.Spreadsheet.Model

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-34s  %s" label detail)

let root = Path.Combine(__SOURCE_DIRECTORY__, "..")
let partsOf (bytes: byte[]) : Map<string, byte[]> =
    use archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read)
    [ for entry in archive.Entries do
        if not (entry.FullName.EndsWith "/") then
            use stream = entry.Open()
            use buffer = new MemoryStream()
            stream.CopyTo buffer
            yield entry.FullName, buffer.ToArray() ]
    |> Map.ofList
let text (data: byte[]) = Text.Encoding.UTF8.GetString data
let says1904 (parts: Map<string, byte[]>) = Text.RegularExpressions.Regex.IsMatch(text parts.["xl/workbook.xml"], "<workbookPr\\b[^>]*\\sdate1904=\"(1|true)\"")

for name in [ "pandas-times_1904.xlsx"; "poi-47504.xlsx"; "poi-47804.xlsx"; "poi-47813.xlsx" ] do
    let bytes = File.ReadAllBytes(Path.Combine(root, "test-documents-excel", name))
    let before = partsOf bytes
    let book = XlsxReader.loadBytesNative bytes |> Option.get
    let opened = Bridge.Import.importWorkbook book
    ok (name + ": the file is in the 1904 system, and the import says so") (says1904 before && book.Date1904 && opened.Date1904 && Rbt.Spreadsheet.Workbook.date1904 opened.Context && opened.Sheets |> Array.forall (fun sheet -> sheet.Model.Date1904 && (M.resolveSheet sheet.Model).Date1904)) (sprintf "%b %b %A" book.Date1904 opened.Date1904 (opened.Sheets |> Array.map (fun sheet -> sheet.Model.Date1904)))
    // The serials are the file's: every number of the first sheet is the <v> of its cell.
    let sheetPath = book.Sheets.[0].Extras.SourcePart |> Option.get |> fst
    let sheetXml = text before.[sheetPath]
    let model = M.resolveSheet opened.Sheets.[0].Model
    let wrong =
        [ for found in Text.RegularExpressions.Regex.Matches(sheetXml, "<c r=\"([A-Z]+)([0-9]+)\"(?: s=\"[0-9]+\")?(?: t=\"n\")?><v>([^<]+)</v></c>") do
            let col = (found.Groups.[1].Value |> Seq.fold (fun acc ch -> acc * 26 + int ch - int 'A' + 1) 0) - 1
            let at : M.CellAddress = { Row = int found.Groups.[2].Value - 1; Col = col }
            match M.getCell at model |> Option.map (fun cell -> cell.Value) with
            | Some (M.CellValue.Number n) when abs (n - Double.Parse(found.Groups.[3].Value, Globalization.CultureInfo.InvariantCulture)) < 1e-9 -> ()
            | other -> yield sprintf "%s%s: %A, the file has %s" found.Groups.[1].Value found.Groups.[2].Value other found.Groups.[3].Value ]
    ok (name + ": the numbers in the cells are the file's serials") wrong.IsEmpty (String.concat "; " (List.truncate 3 wrong))
    let export (edit: M.SpreadsheetModel -> M.SpreadsheetModel) =
        let sheets = opened.Sheets |> Array.map (fun sheet -> sheet.Model)
        sheets.[0] <- edit sheets.[0]
        let input : Bridge.ExportInput =
            { Sheets = sheets; Names = opened.Sheets |> Array.map (fun sheet -> sheet.Name); Visible = opened.Sheets |> Array.map (fun sheet -> sheet.Visible)
              ActiveSheetIndex = book.Preserved.ActiveTab; NameInfo = opened.Context.NameInfo; DefinedNames = opened.Context.Names
              Source = Some opened.Source; Macros = None; MacroProject = None }
        partsOf (XlsxWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut input))
    let same = export id
    ok (name + ": saved untouched: the workbook says date1904") (says1904 same) (text same.["xl/workbook.xml"] |> fun xml -> xml.Substring(0, min 400 xml.Length))
    let sheetParts = before |> Map.toList |> List.map fst |> List.filter (fun path -> path.StartsWith "xl/worksheets/" && not (path.Contains "_rels"))
    ok (name + ": saved untouched: every sheet is the part it was read from")
        (sheetParts |> List.forall (fun path -> Map.tryFind path same = Some before.[path])) (sheetParts |> List.filter (fun path -> Map.tryFind path same <> Some before.[path]) |> String.concat ", ")
    // A date typed into a cell: the sheet's part with that cell in it, the serial in the workbook's system.
    let at : M.CellAddress = { Row = 500; Col = 3 }
    let typed = Rbt.Spreadsheet.Input.parseInputFor model.Date1904 2024 "2018-12-28"
    let edited = export (fun model -> M.setCell at { Value = typed.Value; Format = { M.defaultFormat with NumberFormat = typed.NumberFormat }; Comment = None } (M.resolveSheet model))
    let now = text edited.[sheetPath]
    ok (name + ": a typed date is the 1904 serial") (typed.Value = M.CellValue.Number 42000.0 && Text.RegularExpressions.Regex.IsMatch(now, "<c r=\"D501\"[^>]*><v>42000</v></c>")) (sprintf "%A" typed.Value)
    ok (name + ": edited: the workbook says date1904") (says1904 edited) ""
    let strip (xml: string) = Text.RegularExpressions.Regex.Replace(xml, "<c r=\"D501\"[^>]*>.*?</c>| spans=\"[^\"]*\"|<dimension[^>]*>", "").Replace("<row r=\"501\"></row>", "").Replace("<sheetData></sheetData>", "<sheetData/>")
    ok (name + ": edited: the rest of the sheet's part is the file's") (strip now = strip sheetXml) (sprintf "%d chars, the file's %d" now.Length sheetXml.Length)
    // Read again: still 1904, the same serials, and the typed date reads as that date.
    let again = XlsxReader.loadBytesNative (XlsxWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut { Sheets = [||]; Names = [||]; Visible = [||]; ActiveSheetIndex = 0; NameInfo = Map.empty; DefinedNames = Map.empty; Source = None; Macros = None; MacroProject = None })) 
    ok (name + ": a workbook that is not a file's is in the 1900 system") (again |> Option.forall (fun book -> not book.Date1904)) ""
    // A workbook in the 1900 system opened in between leaves this one as it is: the system is the workbook's own.
    let between = Bridge.Import.importWorkbook (XlsxReader.loadBytesNative (File.ReadAllBytes(Path.Combine(root, "real-world/milk-street-office.xlsx"))) |> Option.get)
    let shown (m: M.SpreadsheetModel) = Rbt.Spreadsheet.CellText.cellDisplayTextIn m.Date1904 None 1.0e6 { M.defaultFormat with NumberFormat = Some "yyyy-mm-dd" } (M.CellValue.Number 42000.0)
    ok (name + ": the typed serial shown as a date, a 1900 workbook opened in between") (shown model = "2018-12-28" && shown (M.resolveSheet between.Sheets.[0].Model) = "2014-12-27") (shown model + " " + shown (M.resolveSheet between.Sheets.[0].Model))
    // The lazily opened workbook says the same of every sheet, converted or not.
    let lazily = Bridge.Import.importWorkbookLazy (fun index -> index = 0) (XlsxReader.loadBytesNative bytes |> Option.get)
    ok (name + ": opened lazily: every sheet says 1904") (lazily.Sheets |> Array.forall (fun sheet -> sheet.Model.Date1904 && (M.resolveSheet sheet.Model).Date1904)) ""

// A workbook in the 1900 system after one in the 1904 system.
do
    let opened = Bridge.Import.importWorkbook (XlsxReader.loadBytesNative (File.ReadAllBytes(Path.Combine(root, "real-world/milk-street-office.xlsx"))) |> Option.get)
    ok "a 1900 workbook opened afterwards is in the 1900 system" (not opened.Date1904 && opened.Sheets |> Array.forall (fun sheet -> not sheet.Model.Date1904)) ""

// A workbook made without a file, its sheets in the 1904 system: the save says date1904, and it reads back so.
do
    let sheet = { M.createEmpty 10 5 with Date1904 = true } |> M.setCell { Row = 0; Col = 0 } { Value = M.CellValue.Number 42000.0; Format = { M.defaultFormat with NumberFormat = Some "yyyy-mm-dd" }; Comment = None }
    let bytes = XlsxWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut { Sheets = [| sheet |]; Names = [| "Sheet1" |]; Visible = [| true |]; ActiveSheetIndex = 0; NameInfo = Map.empty; DefinedNames = Map.empty; Source = None; Macros = None; MacroProject = None })
    let back = Bridge.Import.importWorkbook (XlsxReader.loadBytesNative bytes |> Option.get)
    ok "a new workbook in the 1904 system: saved and read again" (says1904 (partsOf bytes) && back.Date1904 && back.Sheets.[0].Model.Date1904 && M.getCellValue { Row = 0; Col = 0 } (M.resolveSheet back.Sheets.[0].Model) = M.CellValue.Number 42000.0) ""

printfn "\nXLSX 1904 date system tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
