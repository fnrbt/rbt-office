#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

// What a save carries over from the opened file and how it follows edits:
// pivot tables, pivot caches, slicers, slicer caches and external links (kept
// parts with their relationships and content types), what-if data tables,
// hyperlinks, column default styles, and the identity of a sheet through
// renames, reordering and added or removed sheets.

open System
open System.IO
open System.IO.Compression
open Rbt.Spreadsheet.Model
module Bridge = Rbt.Spreadsheet.XlsxBridge
module Wb = Rbt.Spreadsheet.Workbook

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-34s  %s" label detail)
let eq label (got: 'a) (expected: 'a) =
    ok label (got = expected) (sprintf "expected %A  got %A" expected got)


let root = Path.Combine(__SOURCE_DIRECTORY__, "..")
let bytesOf (relative: string) = File.ReadAllBytes(Path.Combine(root, relative))
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
let inputOf (imported: Bridge.ImportedWorkbook) (sheets: SpreadsheetModel array) (names: string array) : Bridge.ExportInput =
    { Sheets = sheets
      Names = names
      Visible = Array.create sheets.Length true
      ActiveSheetIndex = 0
      NameInfo = Map.empty; DefinedNames = imported.Context.Names
      Source = Some imported.Source
      Macros = None }
let modelsOf (imported: Bridge.ImportedWorkbook) = imported.Sheets |> Array.map (fun sheet -> sheet.Model)
let namesOf (imported: Bridge.ImportedWorkbook) = imported.Sheets |> Array.map (fun sheet -> sheet.Name)
let export (imported: Bridge.ImportedWorkbook) (sheets: SpreadsheetModel array) (names: string array) =
    XlsxWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut (inputOf imported sheets names))
let count (needle: string) (text: string) =
    let mutable n = 0
    let mutable at = text.IndexOf needle
    while at >= 0 do
        n <- n + 1
        at <- text.IndexOf (needle, at + needle.Length)
    n
let addr row col : CellAddress = { Row = row; Col = col }
let range sr sc er ec : CellRange = { From = addr sr sc; To = addr er ec }

// ---- pivot tables, pivot caches, slicers, slicer caches, external links -----
let isKeptPart (path: string) =
    [ "xl/pivotTables/"; "xl/pivotCache/"; "xl/slicers/"; "xl/slicerCaches/"; "xl/externalLinks/" ] |> List.exists path.StartsWith

do
    let original = bytesOf "real-world/dashboard-sales-analysis.xlsx"
    let before = partsOf original
    let imported = Bridge.Import.importWorkbook (loadBytes original)
    let models = modelsOf imported
    let names = namesOf imported
    let kept = before |> Map.toList |> List.map fst |> List.filter isKeptPart
    eq "dashboard: the file has the parts" (kept |> List.filter (fun p -> p.StartsWith "xl/pivotTables/pivotTable") |> List.length) 8

    // An edit that has nothing to do with them: every part goes back byte for byte.
    let edited = Array.copy models
    edited.[0] <- setCell (addr 40 0) { Value = CellValue.String "edited"; Format = defaultFormat; Comment = None } edited.[0]
    let written = export imported edited names
    let after = partsOf written
    let differing = kept |> List.filter (fun path -> Map.tryFind path after <> Some before.[path])
    eq "dashboard: pivot / slicer / external-link parts byte-identical" differing []
    let types = textOf after "[Content_Types].xml"
    let untyped =
        kept
        |> List.filter (fun path -> path.EndsWith ".xml")
        |> List.filter (fun path -> not (types.Contains (sprintf "PartName=\"/%s\"" path)))
    eq "dashboard: every kept part has its content type" untyped []
    let workbookXml = textOf after "xl/workbook.xml"
    let workbookRels = textOf after "xl/_rels/workbook.xml.rels"
    eq "dashboard: pivot cache listed" (count "<pivotCache " workbookXml, workbookXml.Contains "cacheId=\"10\"") (1, true)
    eq "dashboard: external references listed" (count "<externalReference " workbookXml) 2
    eq "dashboard: slicer caches listed" (count "slicerCache r:id" workbookXml) 3
    eq "dashboard: workbook relationships" (count "/pivotCacheDefinition\"" workbookRels, count "/slicerCache\"" workbookRels, count "/externalLink\"" workbookRels) (1, 3, 2)
    // The ids workbook.xml names them by are ids of workbook.xml.rels.
    let relIds = DocxReader.relationshipIdsInXml workbookXml
    eq "dashboard: every r:id of workbook.xml resolves" (relIds |> List.filter (fun id -> not (workbookRels.Contains (sprintf "Id=\"%s\"" id)))) []
    eq "dashboard: no relationship id used twice" (relIds |> List.countBy id |> List.filter (fun (_, n) -> n > 1)) []
    // Sheets keep their sheetId: a slicer cache names pivot tables by tabId.
    let sheetIdOf (name: string) (xml: string) =
        let m = Text.RegularExpressions.Regex.Match(xml, sprintf "<sheet name=\"%s\" sheetId=\"(\\d+)\"" name)
        if m.Success then int m.Groups.[1].Value else -1
    eq "dashboard: sheetIds kept" (names |> Array.map (fun name -> sheetIdOf name workbookXml)) (names |> Array.map (fun name -> sheetIdOf name (textOf before "xl/workbook.xml")))
    // Each sheet still points at its pivot tables and slicers.
    let sheetTargets (parts: Map<string, byte[]>) (sheetNumber: int) =
        let rels = textOf parts (sprintf "xl/worksheets/_rels/sheet%d.xml.rels" sheetNumber)
        [ for m in Text.RegularExpressions.Regex.Matches(rels, "Type=\"[^\"]*/(pivotTable|slicer)\" Target=\"([^\"]*)\"") -> m.Groups.[1].Value, (m.Groups.[2].Value.Split('/') |> Array.last) ]
        |> List.sort
    eq "dashboard: sheet relationships" [ for n in 1 .. 8 -> sheetTargets after n ] [ for n in 1 .. 8 -> sheetTargets before n ]
    let sheet1 = textOf after "xl/worksheets/sheet1.xml"
    ok "dashboard: the slicer list is in the sheet" (sheet1.Contains "<x14:slicerList" && sheet1.Contains "x14:slicer r:id=") "no slicerList"
    let slicerRel = Text.RegularExpressions.Regex.Match(sheet1, "x14:slicer r:id=\"([^\"]*)\"").Groups.[1].Value
    ok "dashboard: and its id is a relationship of the sheet" ((textOf after "xl/worksheets/_rels/sheet1.xml.rels").Contains (sprintf "Id=\"%s\"" slicerRel)) slicerRel
    // The written file opens again with the same kept content.
    let again = loadBytes written
    eq "dashboard: pivot locations after a save" (again.Sheets |> List.map (fun s -> s.Extras.PivotLocations |> List.map snd |> List.sort)) (imported.Source.Workbook.Sheets |> List.map (fun s -> s.Extras.PivotLocations |> List.map snd |> List.sort))
    eq "dashboard: kept workbook relationships after a save" (again.Preserved.WorkbookRelationships |> List.map (fun (_, t, target, _) -> t, target) |> List.sort) (imported.Source.Workbook.Preserved.WorkbookRelationships |> List.map (fun (_, t, target, _) -> t, target) |> List.sort)
    // ... and a second save is as good as the first.
    let secondImport = Bridge.Import.importWorkbook again
    let twice = partsOf (export secondImport (modelsOf secondImport) (namesOf secondImport))
    eq "dashboard: byte-identical after two saves" (kept |> List.filter (fun path -> Map.tryFind path twice <> Some before.[path])) []

    // Rows inserted above a pivot table: its location follows; the rest of the part is as it was.
    let linePivot = names |> Array.findIndex (fun name -> name = "Line Pivot")
    let pivotPath = models.[linePivot].KeptRanges.Head.Key
    eq "pivot: where the file put it" (XlsxReader.pivotLocationOf (textOf before pivotPath)) (Some "A3:D27")
    let moved = Array.copy models
    moved.[linePivot] <- models.[linePivot] |> applyStructureEdit (InsertRows (0, 2)) |> applyStructureEdit (InsertCols (0, 1))
    eq "pivot: the model's range follows the edit" moved.[linePivot].KeptRanges.Head.Range (range 4 1 28 4)
    let movedParts = partsOf (export imported moved names)
    let movedXml = textOf movedParts pivotPath
    eq "pivot: location re-anchored" (XlsxReader.pivotLocationOf movedXml) (Some "B5:E29")
    eq "pivot: nothing else of the part changed" (movedXml.Replace("ref=\"B5:E29\"", "ref=\"A3:D27\"")) (textOf before pivotPath)
    eq "pivot: other parts untouched" (kept |> List.filter (fun path -> path <> pivotPath && Map.tryFind path movedParts <> Some before.[path])) []

    // A sheet keeps what the file gave it when it is renamed, moved and joined by a new sheet.
    let shuffled = Array.concat [ [| createEmpty 20 5 |]; Array.rev models ]
    let shuffledNames = Array.concat [ [| "New" |]; names |> Array.rev |> Array.map (fun name -> name + " (renamed)") ]
    let shuffledBytes = export imported shuffled shuffledNames
    let shuffledParts = partsOf shuffledBytes
    eq "identity: kept parts byte-identical after rename / reorder / add" (kept |> List.filter (fun path -> Map.tryFind path shuffledParts <> Some before.[path])) []
    let shuffledWorkbook = loadBytes shuffledBytes
    let pivotCounts (wb: XlsxReader.Workbook) = wb.Sheets |> List.map (fun s -> s.Name.Replace(" (renamed)", ""), (s.Extras.PivotLocations.Length, s.Extras.KeptRelationships.Length, s.Extras.Objects.Length, s.Extras.SheetId)) |> List.filter (fun (name, _) -> name <> "New") |> List.sort
    eq "identity: each sheet has its own pivots, slicers, charts and sheetId" (pivotCounts shuffledWorkbook) (pivotCounts imported.Source.Workbook)
    let newSheetId = shuffledWorkbook.Sheets |> List.find (fun s -> s.Name = "New") |> fun s -> s.Extras.SheetId
    ok "identity: the new sheet got a free sheetId" (newSheetId.IsSome && not (imported.Source.Workbook.Sheets |> List.exists (fun s -> s.Extras.SheetId = newSheetId))) (sprintf "%A" newSheetId)
    // A copy of a sheet is a new sheet: the original keeps the carry-over.
    let copied = Array.append models [| models.[linePivot] |]
    let copiedWorkbook = loadBytes (export imported copied (Array.append names [| "Copy" |]))
    eq "identity: a copied sheet does not duplicate the pivot" (copiedWorkbook.Sheets |> List.map (fun s -> s.Extras.PivotLocations.Length) |> List.sum) 8
    // A removed sheet: the others keep theirs.
    let fewer = models |> Array.indexed |> Array.filter (fun (i, _) -> i <> 2) |> Array.map snd
    let fewerNames = names |> Array.indexed |> Array.filter (fun (i, _) -> i <> 2) |> Array.map snd
    let fewerWorkbook = loadBytes (export imported fewer fewerNames)
    eq "identity: after removing a sheet" (fewerWorkbook.Sheets |> List.map (fun s -> s.Name, s.Extras.PivotLocations.Length)) (imported.Source.Workbook.Sheets |> List.filter (fun s -> s.Name <> names.[2]) |> List.map (fun s -> s.Name, s.Extras.PivotLocations.Length))
    // Models of another import are not this file's sheets.
    let other = Bridge.Import.importWorkbook (loadBytes (bytesOf "real-world/chart-zoo.xlsx"))
    let foreign = loadBytes (export imported (modelsOf other) (namesOf other))
    eq "identity: another file's sheets take nothing from this one" (foreign.Sheets |> List.map (fun s -> s.Extras.PivotLocations.Length + s.Extras.Objects.Length) |> List.sum) 0

// ---- what-if data tables ---------------------------------------------------
for file, sheetName, expected in [ "real-world/milk-street-office.xlsx", "Annual", 10; "real-world/saopaulo-condos.xlsx", "Assumptions", 8 ] do
    let original = bytesOf file
    let wb = loadBytes original
    let imported = Bridge.Import.importWorkbook wb
    let index = imported.Sheets |> Array.findIndex (fun sheet -> sheet.Name = sheetName)
    let model = imported.Sheets.[index].Model
    let label = sprintf "data tables %s" sheetName
    eq (label + ": read") model.DataTables.Length expected
    let sheetPath (parts: Map<string, byte[]>) =
        parts |> Map.toList |> List.map fst |> List.filter (fun p -> p.StartsWith "xl/worksheets/sheet") |> List.filter (fun p -> (textOf parts p).Contains "t=\"dataTable\"")
    let before = partsOf original
    let formulasOf (parts: Map<string, byte[]>) =
        sheetPath parts
        |> List.collect (fun path ->
            [ for m in Text.RegularExpressions.Regex.Matches(textOf parts path, "<c r=\"([A-Z]+[0-9]+)\"[^>]*><f ([^>]*t=\"dataTable\"[^>]*)/><v>([^<]*)</v>") ->
                let attrs = [ for a in Text.RegularExpressions.Regex.Matches(m.Groups.[2].Value, "([A-Za-z0-9]+)=\"([^\"]*)\"") -> a.Groups.[1].Value, a.Groups.[2].Value ] |> List.filter (fun (name, _) -> name <> "aca") |> List.sort
                let value =
                    match Double.TryParse(m.Groups.[3].Value, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture) with
                    | true, number -> number.ToString("R", Globalization.CultureInfo.InvariantCulture)
                    | _ -> m.Groups.[3].Value
                m.Groups.[1].Value, attrs, value ])
        |> List.sort
    let written = partsOf (export imported (modelsOf imported) (namesOf imported))
    eq (label + ": written back as they were") (formulasOf written) (formulasOf before)
    eq (label + ": all of them") (formulasOf written).Length expected
    // A recalculation leaves the results the file had.
    let valuesIn (m: SpreadsheetModel) =
        [ for table in m.DataTables do
            for row in table.Range.From.Row .. table.Range.To.Row do
                for col in table.Range.From.Col .. table.Range.To.Col -> getCell (addr row col) m |> Option.map (fun data -> data.Value) ]
    let recalculated = Rbt.Spreadsheet.Formula.recalculateWorkbook imported.Context
    let recalculatedModel = recalculated.Sheets.[Bridge.sheetIdOfIndex index].Model
    eq (label + ": results survive a recalculation") (valuesIn recalculatedModel) (valuesIn model)
    ok (label + ": the range holds values, not formulas") (valuesIn model |> List.forall (function Some (CellValue.Formula _) -> false | _ -> true)) "a formula in a data table"
    // Rows and columns inserted before them: ranges and input cells follow.
    let sheetId = Bridge.sheetIdOfIndex index
    let edited, _ = Wb.applyStructureEdit sheetId (InsertRows (0, 3)) imported.Context
    let edited, _ = Wb.applyStructureEdit sheetId (InsertCols (0, 1)) edited
    let movedModel = edited.Sheets.[sheetId].Model
    let shift (text: string) =
        // "B5" -> "C8", "$B$5" -> "$C$8" (one column right, three rows down)
        Text.RegularExpressions.Regex.Replace(text, "(\\$?)([A-Z]+)(\\$?)([0-9]+)", fun m ->
            m.Groups.[1].Value + columnToLetter (letterToColumn m.Groups.[2].Value + 1) + m.Groups.[3].Value + string (int m.Groups.[4].Value + 3))
    let describe (m: SpreadsheetModel) (map: string -> string) =
        m.DataTables |> List.map (fun t -> map (formatCellRef t.Range.From.Row t.Range.From.Col), map (formatCellRef t.Range.To.Row t.Range.To.Col), map t.Input1, t.Input2 |> Option.map map)
    eq (label + ": ranges and inputs follow a structure edit") (describe movedModel id) (describe model shift)
    eq (label + ": the results moved with them") (valuesIn movedModel |> List.map (Option.map unwrapFormulaValue)) (valuesIn model |> List.map (Option.map unwrapFormulaValue))
    let models = modelsOf imported |> Array.mapi (fun i m -> edited.Sheets.[Bridge.sheetIdOfIndex i].Model)
    let movedParts = partsOf (export imported models (namesOf imported))
    eq (label + ": written at their new place")
        (formulasOf movedParts |> List.map (fun (cell, attrs, value) -> cell, attrs, value))
        (formulasOf before |> List.map (fun (cell, attrs, value) -> shift cell, attrs |> List.map (fun (name, text) -> name, (if name = "ref" || name = "r1" || name = "r2" then shift text else text)), value) |> List.sort)
    // ... and read again.
    let again = loadBytes (XlsxWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut (inputOf imported models (namesOf imported))))
    eq (label + ": read back after the edit") (again.Sheets.[index].Extras.DataTables |> List.map (fun t -> t.Ref) |> List.sort) (movedModel.DataTables |> List.map (fun t -> formatCellRef t.Range.From.Row t.Range.From.Col + ":" + formatCellRef t.Range.To.Row t.Range.To.Col) |> List.sort)

// ---- hyperlinks follow their cells -----------------------------------------
do
    // A workbook with hyperlinks: one from the corpus, with links added through the writer.
    let wb = loadBytes (bytesOf "real-world/dcf-stock-analysis.xlsx")
    let package = XlsxWriter.packageFromReaderWorkbook wb
    let links : XlsxWriter.XlsxHyperlink list =
        [ { Range = "B4"; Target = "https://example.com/a"; IsInternal = false; Tooltip = None; Display = None }
          { Range = "C6:D7"; Target = "https://example.com/b?x=1&y=2"; IsInternal = false; Tooltip = None; Display = None }
          { Range = "A10"; Target = "Sheet1!A1"; IsInternal = true; Tooltip = None; Display = None } ]
    let withLinks = { package with Sheets = package.Sheets |> List.mapi (fun i sheet -> if i = 0 then { sheet with Hyperlinks = links } else sheet) }
    let bytes = XlsxWriter.writeWorkbookPackageBytes withLinks
    let imported = Bridge.Import.importWorkbook (loadBytes bytes)
    let model = imported.Sheets.[0].Model
    eq "hyperlinks: in the model" (model.Hyperlinks |> List.map (fun l -> l.Range, l.Target, l.IsInternal)) [ range 3 1 3 1, "https://example.com/a", false; range 5 2 6 3, "https://example.com/b?x=1&y=2", false; range 9 0 9 0, "Sheet1!A1", true ]
    let hyperlinksAfter (models: SpreadsheetModel array) =
        (loadBytes (export imported models (namesOf imported))).Sheets.[0].Hyperlinks |> List.map (fun l -> l.Range, l.Target, l.IsInternal)
    eq "hyperlinks: unedited" (hyperlinksAfter (modelsOf imported)) (links |> List.map (fun l -> l.Range, l.Target, l.IsInternal))
    let edit (edits: StructureEdit list) =
        let models = modelsOf imported
        models.[0] <- edits |> List.fold (fun m e -> applyStructureEdit e m) model
        hyperlinksAfter models
    eq "hyperlinks: a row inserted above" (edit [ InsertRows (0, 1) ]) [ "B5", "https://example.com/a", false; "C7:D8", "https://example.com/b?x=1&y=2", false; "A11", "Sheet1!A1", true ]
    eq "hyperlinks: a column inserted to the left" (edit [ InsertCols (1, 2) ]) [ "D4", "https://example.com/a", false; "E6:F7", "https://example.com/b?x=1&y=2", false; "A10", "Sheet1!A1", true ]
    eq "hyperlinks: a row inserted inside a linked range grows it" (edit [ InsertRows (6, 1) ]) [ "B4", "https://example.com/a", false; "C6:D8", "https://example.com/b?x=1&y=2", false; "A11", "Sheet1!A1", true ]
    eq "hyperlinks: the linked row deleted" (edit [ DeleteRows (3, 1) ]) [ "C5:D6", "https://example.com/b?x=1&y=2", false; "A9", "Sheet1!A1", true ]
    eq "hyperlinks: part of a linked range deleted" (edit [ DeleteCols (3, 1) ]) [ "B4", "https://example.com/a", false; "C6:C7", "https://example.com/b?x=1&y=2", false; "A10", "Sheet1!A1", true ]

// ---- comments on a sheet whose VML drawing holds other shapes ---------------
do
    let source = bytesOf "real-world/dashboard-ecommerce-sales.xlsx"
    let original = partsOf source
    let vmlPath = "xl/drawings/vmlDrawing1.vml"
    let wb = loadBytes source
    let imported = Bridge.Import.importWorkbook wb
    // The sheet that names the VML drawing.
    let index = wb.Sheets |> List.findIndex (fun sheet -> sheet.Extras.KeptRelationships |> List.exists (fun (_, _, target, _) -> target.EndsWith "vmlDrawing1.vml"))
    let pictures = count "ObjectType=\"Pict\"" (textOf original vmlPath)
    ok "vml: the file's drawing has camera pictures and no notes" (pictures > 0 && count "ObjectType=\"Note\"" (textOf original vmlPath) = 0) (string pictures)
    let withComment (cells: (int * int) list) (model: SpreadsheetModel) =
        cells |> List.fold (fun m (row, col) ->
            let cell = Map.tryFind (addr row col) m.Cells |> Option.defaultValue { Value = CellValue.Empty; Format = defaultFormat; Comment = None }
            { m with Cells = Map.add (addr row col) { cell with Comment = Some { Text = sprintf "note %d %d" row col; Author = Some "Tester" } } m.Cells }) model
    let models = modelsOf imported
    models.[index] <- withComment [ (3, 2); (40, 5) ] models.[index]
    let saved = export imported models (namesOf imported)
    let written = partsOf saved
    let vml = textOf written vmlPath
    eq "vml: the pictures stay" (count "ObjectType=\"Pict\"" vml) pictures
    eq "vml: each new comment has a box" (count "ObjectType=\"Note\"" vml) 2
    ok "vml: the boxes are at the comments' cells" (vml.Contains "<x:Row>3</x:Row><x:Column>2</x:Column>" && vml.Contains "<x:Row>40</x:Row><x:Column>5</x:Column>") ""
    ok "vml: the note shape type is declared once" (count "id=\"_x0000_t202\"" vml = 1) ""
    let ids = [ for m in Text.RegularExpressions.Regex.Matches(vml, "_x0000_s(\\d+)\"") -> m.Groups.[1].Value ]
    ok "vml: shape ids are distinct" (ids.Length = (List.distinct ids).Length) (String.concat "," ids)
    ok "vml: well formed" (try Xml.Linq.XDocument.Parse vml |> ignore; true with _ -> false) ""
    ok "vml: one VML drawing for the sheet" (written |> Map.toList |> List.filter (fun (path, _) -> path.EndsWith ".vml") |> List.length = (original |> Map.toList |> List.filter (fun (path, _) -> path.EndsWith ".vml") |> List.length)) ""
    // Opened again: the comments are there; one removed, its box goes and the other stays as it is.
    let again = loadBytes saved
    let reimported = Bridge.Import.importWorkbook again
    let commentsOf (model: SpreadsheetModel) = model.Cells |> Map.toList |> List.choose (fun (a, d) -> d.Comment |> Option.map (fun _ -> a.Row, a.Col))
    eq "vml: the comments are read back" (commentsOf reimported.Sheets.[index].Model) [ (3, 2); (40, 5) ]
    let models2 = modelsOf reimported
    models2.[index] <- { models2.[index] with Cells = models2.[index].Cells |> Map.add (addr 3 2) { models2.[index].Cells.[addr 3 2] with Comment = None } }
    let vml2 = textOf (partsOf (export reimported models2 (namesOf reimported))) vmlPath
    eq "vml: a removed comment's box is gone" (count "ObjectType=\"Note\"" vml2, vml2.Contains "<x:Row>40</x:Row>", count "ObjectType=\"Pict\"" vml2) (1, true, pictures)
    // Nothing changed: the drawing is the file's.
    ok "vml: untouched comments leave the drawing alone" ((partsOf (export reimported (modelsOf reimported) (namesOf reimported))).[vmlPath] = (partsOf saved).[vmlPath]) ""

// ---- a changed table keeps what the model does not hold ---------------------
do
    let source = bytesOf "test-documents-excel/pandas-test_decimal.xlsx"
    let original = partsOf source
    let wb = loadBytes source
    let imported = Bridge.Import.importWorkbook wb
    let index = imported.Sheets |> Array.findIndex (fun sheet -> not sheet.Model.Tables.IsEmpty)
    let model = imported.Sheets.[index].Model
    let table = model.Tables.Head
    // One more row, a header renamed.
    let header = addr (table.StartRow - 1) (table.StartCol)
    let edited =
        { model with
            Tables = [ { table with EndRow = table.EndRow + 1 } ]
            Cells = model.Cells |> Map.add header { model.Cells.[header] with Value = CellValue.String "Amount" } }
    let models = modelsOf imported
    models.[index] <- edited
    let written = partsOf (export imported models (namesOf imported))
    let xml = textOf written "xl/tables/table1.xml"
    ok "table: the changed table is written" (xml.Contains "ref=\"A1:F6\"" && xml.Contains "<autoFilter ref=\"A1:F6\"") xml
    ok "table: its revision ids stay" (xml.Contains "xr:uid=\"{0F235FC8-2683-4264-B306-3D553F7838E8}\"" && xml.Contains "xr3:uid=\"{1BB94865-8D62-4199-B271-3EA12BCA4449}\"") xml
    ok "table: the columns' formats stay" (count "dataDxfId=" xml = 6 && xml.Contains "name=\"Amount\" queryTableFieldId=\"2\" dataDxfId=\"4\"") xml
    ok "table: the query table binding stays" (xml.Contains "tableType=\"queryTable\"" && count "queryTableFieldId=" xml = 6 && count "uniqueName=" xml = 6) xml
    ok "table: the query table part and its relationship stay"
        (written.ContainsKey "xl/queryTables/queryTable1.xml" && written.["xl/tables/_rels/table1.xml.rels"] = original.["xl/tables/_rels/table1.xml.rels"]) ""
    ok "table: well formed" (try Xml.Linq.XDocument.Parse xml |> ignore; true with _ -> false) ""
    eq "table: one table part" (written |> Map.toList |> List.filter (fun (path, _) -> path.StartsWith "xl/tables/table") |> List.length) 1
    let back = (loadBytes (export imported models (namesOf imported))).Sheets.[index].Tables.Head
    eq "table: read back" (back.Name, back.EndRow, back.Columns.[1]) ("test_decimal", table.EndRow + 1, "Amount")
    // Renamed and one column narrower: still the file's part.
    let models = modelsOf imported
    models.[index] <- { model with Tables = [ { table with Name = "Decimals"; EndCol = table.EndCol - 1 } ] }
    let xml = textOf (partsOf (export imported models (namesOf imported))) "xl/tables/table1.xml"
    ok "table: renamed and narrowed" (xml.Contains "displayName=\"Decimals\"" && xml.Contains "ref=\"A1:E5\"" && xml.Contains "<tableColumns count=\"5\">" && count "<tableColumn " xml = 5 && not (xml.Contains "name=\"Column6\"") && xml.Contains "xr:uid=") xml

// The same on a part with calculated columns, filter criteria, totals and a sort order.
do
    let part =
        "<?xml version=\"1.0\"?><table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:xr=\"urn:xr\" id=\"7\" xr:uid=\"{T}\" name=\"Sales\" displayName=\"Sales\" ref=\"B2:E10\" totalsRowCount=\"1\" headerRowDxfId=\"9\" dataDxfId=\"8\">"
        + "<autoFilter ref=\"B2:E9\" xr:uid=\"{F}\"><filterColumn colId=\"1\"><filters><filter val=\"x\"/></filters></filterColumn></autoFilter>"
        + "<sortState ref=\"B3:E9\"><sortCondition ref=\"C3:C9\"/></sortState>"
        + "<tableColumns count=\"4\"><tableColumn id=\"1\" name=\"Item\" totalsRowLabel=\"Total\"/><tableColumn id=\"2\" name=\"Qty\" totalsRowFunction=\"sum\" dataDxfId=\"3\"/>"
        + "<tableColumn id=\"5\" name=\"Price\"/><tableColumn id=\"9\" name=\"Amount\" totalsRowFunction=\"custom\" dataDxfId=\"2\"><calculatedColumnFormula>Sales[[#This Row],[Qty]]*Sales[[#This Row],[Price]]</calculatedColumnFormula><totalsRowFormula>SUM(E3:E9)/2</totalsRowFormula></tableColumn></tableColumns>"
        + "<tableStyleInfo name=\"TableStyleLight1\" showRowStripes=\"1\"/><extLst><ext uri=\"u\"/></extLst></table>"
    let style : XlsxWriter.XlsxTableStyle = { Name = "TableStyleMedium2"; ShowFirstColumn = true; ShowLastColumn = false; ShowRowStripes = true; ShowColumnStripes = false }
    let table : XlsxWriter.XlsxTableOut =
        { Name = "Sales"; StartRow = 2; StartCol = 2; EndRow = 12; EndCol = 6; Columns = [ "Item"; "Qty"; "Tax"; "Price"; "Amount" ]; HeaderRowCount = 1; TotalsRowCount = 1
          Totals = [ "label", "All"; "function", "average"; "", ""; "", ""; "custom", "MAX(F3:F11)" ]; AutoFilter = Some (2, 6); Style = Some style; SourcePart = None }
    let xml = XlsxWriter.tableXmlFromSource part table |> Option.defaultValue ""
    ok "table part: well formed" (try Xml.Linq.XDocument.Parse xml |> ignore; true with _ -> false) xml
    ok "table part: root" (xml.Contains "id=\"7\" xr:uid=\"{T}\"" && xml.Contains "ref=\"B2:F12\"" && xml.Contains "headerRowDxfId=\"9\" dataDxfId=\"8\"" && xml.Contains "totalsRowCount=\"1\"") xml
    ok "table part: the filter keeps its criteria on the new range" (xml.Contains "<autoFilter ref=\"B2:F11\" xr:uid=\"{F}\"><filterColumn colId=\"1\">") xml
    ok "table part: a sort order of another range is dropped" (not (xml.Contains "sortState")) xml
    ok "table part: columns keep their ids and formats; a new one gets a free id"
        (xml.Contains "<tableColumns count=\"5\"><tableColumn id=\"1\" name=\"Item\" totalsRowLabel=\"All\"/><tableColumn id=\"2\" name=\"Qty\" totalsRowFunction=\"average\" dataDxfId=\"3\"/><tableColumn id=\"10\" name=\"Tax\"/><tableColumn id=\"5\" name=\"Price\"/><tableColumn id=\"9\" name=\"Amount\" totalsRowFunction=\"custom\" dataDxfId=\"2\">") xml
    ok "table part: the calculated column formula stays, the total's formula is the new one"
        (xml.Contains "<calculatedColumnFormula>Sales[[#This Row],[Qty]]*Sales[[#This Row],[Price]]</calculatedColumnFormula><totalsRowFormula>MAX(F3:F11)</totalsRowFormula></tableColumn>") xml
    ok "table part: style and what follows" (xml.Contains "<tableStyleInfo name=\"TableStyleMedium2\" showFirstColumn=\"1\" showLastColumn=\"0\" showRowStripes=\"1\" showColumnStripes=\"0\"/><extLst><ext uri=\"u\"/></extLst></table>") xml
    // No totals row any more, no filter, no header.
    let bare = XlsxWriter.tableXmlFromSource part { table with StartRow = 2; EndRow = 9; EndCol = 5; Columns = [ "Item"; "Qty"; "Price"; "Amount" ]; TotalsRowCount = 0; Totals = []; AutoFilter = None; HeaderRowCount = 0 } |> Option.defaultValue ""
    ok "table part: without totals, filter and header"
        (bare.Contains "totalsRowShown=\"0\"" && not (bare.Contains "totalsRowCount") && bare.Contains "headerRowCount=\"0\"" && not (bare.Contains "<autoFilter")
         && bare.Contains "totalsRowFunction=\"sum\"" && (try Xml.Linq.XDocument.Parse bare |> ignore; true with _ -> false)) bare

    // A new column of a query table names itself (uniqueName), as the file's columns do.
    let queryPart = part.Replace("<table xmlns", "<table tableType=\"queryTable\" xmlns").Replace("<tableColumn id=\"1\" name", "<tableColumn id=\"1\" uniqueName=\"1\" name")
    let queried = XlsxWriter.tableXmlFromSource queryPart table |> Option.defaultValue ""
    ok "table part: the test's part is a query table" (queryPart.Contains "tableType=\"queryTable\"" && queryPart <> part) queryPart
    ok "table part: a new column of a query table gets a uniqueName" (queried.Contains "<tableColumn id=\"10\" uniqueName=\"10\" name=\"Tax\"/>") queried
    ok "table part: and no other table's does" (not (xml.Contains "uniqueName")) xml

// ---- a format changed by the user is made from the file's -----------------------
// The cell's new cellXfs entry is the file's entry with what the change touched
// replaced: protection, the quote prefix, the named style and the parts of the
// format the change left alone (their fonts, fills and borders as the file has them) stay.
do
    let bytes = bytesOf "real-world/milk-street-office.xlsx"
    let original = partsOf bytes
    let imported = Bridge.Import.importWorkbook (loadBytes bytes)
    let models = modelsOf imported
    let xfsOf (styles: string) =
        let body = styles.Substring(styles.IndexOf "<cellXfs", styles.IndexOf "</cellXfs>" - styles.IndexOf "<cellXfs")
        [| for m in Text.RegularExpressions.Regex.Matches(body, "<xf(?=[\\s>/])(?:[^>]*/>|[^>]*>[\\s\\S]*?</xf>)") -> m.Value |]
    let attrOf (name: string) (xml: string) =
        let m = Text.RegularExpressions.Regex.Match(xml, "\\s" + name + "=\"([^\"]*)\"")
        if m.Success then m.Groups.[1].Value else ""
    let sourceXfs = xfsOf (textOf original "xl/styles.xml")
    let sheetXml = textOf original "xl/worksheets/sheet1.xml"
    // A cell of the first sheet whose format is unlocked, and one with a quote prefix.
    let cellWith (wanted: string -> bool) =
        [ for m in Text.RegularExpressions.Regex.Matches(sheetXml, "<c r=\"([A-Z]+)(\\d+)\" s=\"(\\d+)\"") do
            let xf = int m.Groups.[3].Value
            if xf < sourceXfs.Length && wanted sourceXfs.[xf] then
                yield { Row = int m.Groups.[2].Value - 1; Col = Bridge.columnIndexOfLetters m.Groups.[1].Value }, xf ]
        |> List.tryFind (fun (cell, _) -> Map.containsKey cell models.[0].Cells)
    let unlocked = cellWith (fun xf -> xf.Contains "<protection locked=\"0\"" && attrOf "xfId" xf <> "0" && attrOf "xfId" xf <> "")
    let unlocked = match unlocked with Some found -> Some found | None -> cellWith (fun xf -> xf.Contains "<protection locked=\"0\"")
    ok "format change: the workbook has an unlocked cell" unlocked.IsSome ""
    match unlocked with
    | Some (cell, xf) ->
        let data = models.[0].Cells.[cell]
        let edited = Array.copy models
        edited.[0] <- { models.[0] with Cells = Map.add cell { data with Format = { data.Format with Italic = not data.Format.Italic } } models.[0].Cells }
        let written = partsOf (export imported edited (namesOf imported))
        let nowXfs = xfsOf (textOf written "xl/styles.xml")
        let nowSheet = textOf written "xl/worksheets/sheet1.xml"
        let reference = columnToLetter cell.Col + string (cell.Row + 1)
        let m = Text.RegularExpressions.Regex.Match(nowSheet, "<c r=\"" + reference + "\" s=\"(\\d+)\"")
        let now = if m.Success then int m.Groups.[1].Value else -1
        ok "format change: the cell has a new format behind the file's" (now >= sourceXfs.Length && now < nowXfs.Length) (sprintf "%d of %d (the file has %d)" now nowXfs.Length sourceXfs.Length)
        if now >= sourceXfs.Length && now < nowXfs.Length then
            let was, made = sourceXfs.[xf], nowXfs.[now]
            ok "format change: the file's formats keep their places" (Array.truncate sourceXfs.Length nowXfs = sourceXfs) ""
            ok "format change: protection stays" (made.Contains "<protection locked=\"0\"") (was + "  ->  " + made)
            eq "format change: the named style, number format, fill and border stay"
                (attrOf "xfId" made, attrOf "numFmtId" made, attrOf "fillId" made, attrOf "borderId" made) (attrOf "xfId" was, attrOf "numFmtId" was, attrOf "fillId" was, attrOf "borderId" was)
            ok "format change: the font is a new one" (attrOf "fontId" made <> attrOf "fontId" was && attrOf "applyFont" made = "1") made
            eq "format change: the alignment stays" (Text.RegularExpressions.Regex.Match(made, "<alignment[^>]*>").Value) (Text.RegularExpressions.Regex.Match(was, "<alignment[^>]*>").Value)
            // Read back: the cell is italic now, and still what it was otherwise.
            let again = Bridge.Import.importWorkbook (loadBytes (export imported edited (namesOf imported)))
            eq "format change: read back" again.Sheets.[0].Model.Cells.[cell].Format { data.Format with Italic = not data.Format.Italic }
    | None -> ()
    // A quote prefix stays through a change of the alignment (the entry gets an alignment of its own).
    match cellWith (fun xf -> xf.Contains "quotePrefix=\"1\"") with
    | Some (cell, xf) ->
        let data = models.[0].Cells.[cell]
        let edited = Array.copy models
        edited.[0] <- { models.[0] with Cells = Map.add cell { data with Format = { data.Format with HorizontalAlign = Some "right"; WrapText = true } } models.[0].Cells }
        let written = partsOf (export imported edited (namesOf imported))
        let nowXfs = xfsOf (textOf written "xl/styles.xml")
        let reference = columnToLetter cell.Col + string (cell.Row + 1)
        let m = Text.RegularExpressions.Regex.Match(textOf written "xl/worksheets/sheet1.xml", "<c r=\"" + reference + "\" s=\"(\\d+)\"")
        let now = if m.Success then int m.Groups.[1].Value else -1
        ok "format change (alignment): a new format" (now >= sourceXfs.Length && now < nowXfs.Length) (string now)
        if now >= sourceXfs.Length && now < nowXfs.Length then
            let was, made = sourceXfs.[xf], nowXfs.[now]
            ok "format change (alignment): the quote prefix stays" (made.Contains "quotePrefix=\"1\"") made
            ok "format change (alignment): the alignment is the new one" (made.Contains "horizontal=\"right\"" && made.Contains "wrapText=\"1\"" && attrOf "applyAlignment" made = "1") made
            eq "format change (alignment): the font stays" (attrOf "fontId" made) (attrOf "fontId" was)
            ok "format change (alignment): well formed" (try Xml.Linq.XElement.Parse made |> ignore; true with _ -> false) made
    | None -> ok "format change: the workbook has a cell with a quote prefix on its first sheet" false ""

// ---- sheets renamed or deleted: the sheets that do not name them go back as they were read ----
for file in [ "real-world/lyric-multifamily.xlsx"; "real-world/dashboard-ecommerce-sales.xlsx" ] do
    let label = Path.GetFileNameWithoutExtension file
    let bytes = bytesOf file
    let original = partsOf bytes
    let reader = loadBytes bytes
    let imported = Bridge.Import.importWorkbook reader
    let names = namesOf imported
    let partOf (index: int) = (List.item index reader.Sheets).Extras.SourcePart |> Option.map fst
    let isWorksheet (index: int) = match (List.item index reader.Sheets).Extras.SourcePart with Some (_, relType) -> relType.EndsWith "/worksheet" | None -> false
    let mentions (index: int) (name: string) =
        match partOf index with
        | Some path -> (textOf original path).ToLowerInvariant().Contains (name.ToLowerInvariant())
        | None -> true
    // The sheet the fewest others name.
    let target =
        [ 0 .. names.Length - 1 ]
        |> List.minBy (fun k -> [ 0 .. names.Length - 1 ] |> List.filter (fun i -> i <> k && mentions i names.[k]) |> List.length)
    let others = [ 0 .. names.Length - 1 ] |> List.filter (fun i -> i <> target && isWorksheet i)
    let silent = others |> List.filter (fun i -> not (mentions i names.[target]))
    ok (label + ": there are sheets that do not name the renamed one") (not silent.IsEmpty) names.[target]
    // Renamed.
    let renamed = names |> Array.mapi (fun i name -> if i = target then "Renamed Tab" else name)
    let written = partsOf (export imported (modelsOf imported) renamed)
    // (But for the sheet the file had on display: its tabSelected goes to the first sheet.)
    let unselected (text: string) = text.Replace(" tabSelected=\"1\"", "")
    let same (parts: Map<string, byte[]>) (index: int) = match partOf index with Some path -> unselected (textOf parts path) = unselected (textOf original path) | None -> false
    let differing = silent |> List.filter (fun i -> not (same written i))
    ok (label + ": renamed: the sheets that do not name it are the file's parts") differing.IsEmpty (sprintf "%A of %A" (differing |> List.map (fun i -> names.[i])) (silent |> List.map (fun i -> names.[i])))
    ok (label + ": renamed: the workbook lists the new name") ((textOf written "xl/workbook.xml").Contains "name=\"Renamed Tab\"" && not ((textOf written "xl/workbook.xml").Contains ("<sheet name=\"" + names.[target] + "\""))) ""
    for i in others |> List.filter (fun i -> mentions i names.[target]) do
        match partOf i with
        | Some path ->
            let now = textOf written path
            let quoted = "'" + names.[target] + "'!"
            ok (sprintf "%s: renamed: %s follows" label names.[i]) (not (now.Contains (names.[target] + "!")) && not (now.Contains quoted) || not ((textOf original path).Contains (names.[target] + "!") || (textOf original path).Contains quoted)) path
        | None -> ()
    let again = Bridge.Import.importWorkbook (loadBytes (export imported (modelsOf imported) renamed))
    eq (label + ": renamed: read back") (namesOf again) renamed
    eq (label + ": renamed: the cells of a sheet that went back as it was") (silent |> List.map (fun i -> again.Sheets.[i].Model.Cells = imported.Sheets.[i].Model.Cells)) (silent |> List.map (fun _ -> true))
    // Deleted.
    let keep = [| for i in 0 .. names.Length - 1 do if i <> target then yield i |]
    let withoutIt = partsOf (export imported (keep |> Array.map (fun i -> (modelsOf imported).[i])) (keep |> Array.map (fun i -> names.[i])))
    let differing = silent |> List.filter (fun i -> not (same withoutIt i))
    ok (label + ": deleted: the sheets that do not name it are the file's parts") differing.IsEmpty (sprintf "%A" (differing |> List.map (fun i -> names.[i])))
    let again = Bridge.Import.importWorkbook (loadBytes (export imported (keep |> Array.map (fun i -> (modelsOf imported).[i])) (keep |> Array.map (fun i -> names.[i]))))
    eq (label + ": deleted: read back") (namesOf again) (keep |> Array.map (fun i -> names.[i]))

// ---- a print area of several areas --------------------------------------------
do
    let source = bytesOf "real-world/milk-street-office.xlsx"
    let imported = Bridge.Import.importWorkbook (loadBytes source)
    let names = namesOf imported
    let models = modelsOf imported
    let areas = [ range 0 0 9 2; range 0 4 9 6; range 20 0 29 2 ]
    models.[0] <- { models.[0] with PageSetup = { models.[0].PageSetup with PrintArea = Some areas.Head; ExtraPrintAreas = areas.Tail } }
    let saved = export imported models names
    let workbookXml = (textOf (partsOf saved) "xl/workbook.xml").Replace("&apos;", "'")
    let quoted = if names.[0] |> Seq.forall Char.IsLetterOrDigit then names.[0] else "'" + names.[0] + "'"
    ok "print areas: written as one name, each area with its sheet"
        (workbookXml.Contains (sprintf ">%s!$A$1:$C$10,%s!$E$1:$G$10,%s!$A$21:$C$30</definedName>" quoted quoted quoted) || workbookXml.Contains (sprintf ">'%s'!$A$1:$C$10,'%s'!$E$1:$G$10,'%s'!$A$21:$C$30</definedName>" names.[0] names.[0] names.[0]))
        (let at = workbookXml.IndexOf "_xlnm.Print_Area" in if at < 0 then "no print area" else workbookXml.Substring(at, min 200 (workbookXml.Length - at)))
    let back = loadBytes saved
    eq "print areas: read as the areas, without the sheet" back.Sheets.[0].PrintArea (Some "$A$1:$C$10,$E$1:$G$10,$A$21:$C$30")
    let again = Bridge.Import.importWorkbook back
    eq "print areas: in the model" (printAreasOf again.Sheets.[0].Model.PageSetup) areas
    // Untouched, they go back as the file states them; a row inserted above moves them all.
    let savedAgain = export again (modelsOf again) (namesOf again)
    eq "print areas: an unchanged save" ((loadBytes savedAgain).Sheets.[0].PrintArea) (Some "$A$1:$C$10,$E$1:$G$10,$A$21:$C$30")
    let moved = modelsOf again
    moved.[0] <- applyStructureEdit (InsertRows (0, 2)) moved.[0]
    eq "print areas: follow inserted rows" ((loadBytes (export again moved (namesOf again))).Sheets.[0].PrintArea) (Some "$A$3:$C$12,$E$3:$G$12,$A$23:$C$32")
    let cut = modelsOf again
    cut.[0] <- applyStructureEdit (DeleteCols (0, 3)) cut.[0]
    eq "print areas: an area whose columns are deleted goes" ((loadBytes (export again cut (namesOf again))).Sheets.[0].PrintArea) (Some "$B$1:$D$10")

// ---- column default styles follow their columns --------------------------------
do
    let wb = loadBytes (bytesOf "real-world/saopaulo-condos.xlsx")
    let imported = Bridge.Import.importWorkbook wb
    let index = imported.Sheets |> Array.findIndex (fun sheet -> not sheet.Model.ColumnStyles.IsEmpty)
    let model = imported.Sheets.[index].Model
    let stylesOf (sheet: XlsxReader.Worksheet) (book: XlsxReader.Workbook) =
        sheet.Columns
        |> List.choose (fun column ->
            match column.Style with
            | Some style when style > 0 && style < book.Styles.Length -> Some (column.Index - 1, column.MaxIndex - 1, book.Styles.[style])
            | _ -> None)
    let original = stylesOf wb.Sheets.[index] wb
    ok "column styles: the file has some" (not original.IsEmpty) "none"
    eq "column styles: in the model" (model.ColumnStyles |> List.map (fun s -> s.First, s.Last)) (original |> List.map (fun (first, last, _) -> first, last))
    let after (edits: StructureEdit list) =
        let models = modelsOf imported
        models.[index] <- edits |> List.fold (fun m e -> applyStructureEdit e m) model
        let again = loadBytes (export imported models (namesOf imported))
        stylesOf again.Sheets.[index] again
    // A column a style is written for, expanded to single columns (the writer merges neighbours).
    let perColumn (styles: (int * int * XlsxReader.CellStyle) list) =
        styles |> List.collect (fun (first, last, style) -> [ for col in first .. min last 400 -> col, style ]) |> List.sortBy fst
    eq "column styles: unedited" (perColumn (after [])) (perColumn original)
    let firstStyled = original |> List.map (fun (first, _, _) -> first) |> List.min
    eq "column styles: a column inserted before them"
        (perColumn (after [ InsertCols (firstStyled, 1) ]))
        (perColumn (original |> List.map (fun (first, last, style) -> first + 1, (if last >= 16383 then last else last + 1), style)))
    eq "column styles: the first styled column deleted"
        (perColumn (after [ DeleteCols (firstStyled, 1) ]) |> List.map fst)
        (perColumn (original |> List.choose (fun (first, last, style) ->
            if first = firstStyled && last = firstStyled then None
            elif first = firstStyled then Some (first, last - 1, style)
            else Some (first - 1, (if last >= 16383 then last else last - 1), style))) |> List.map fst)
    eq "column styles: row edits leave them" (perColumn (after [ InsertRows (0, 2); DeleteRows (5, 1) ])) (perColumn original)

printfn ""
printfn "XLSX carry-over tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
