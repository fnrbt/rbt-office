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
      DefinedNames = imported.Context.Names
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
        [ { Range = "B4"; Target = "https://example.com/a"; IsInternal = false }
          { Range = "C6:D7"; Target = "https://example.com/b?x=1&y=2"; IsInternal = false }
          { Range = "A10"; Target = "Sheet1!A1"; IsInternal = true } ]
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
