#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

// What a save keeps of a sheet's part when more than cells changed.
//
// Without a row or column inserted or deleted the rows stay the file's: a
// changed row height, a hidden or grouped row, a filter, another page setup
// are put into the part as it is; a link or a note makes the rest of the
// part be written from the sheet, its rows still the file's (their styles,
// the shared formulas). With rows or columns inserted or deleted the sheet
// is written from its cells, and still says what its part said beyond what
// a sheet holds: the code name, ignoredErrors (where their cells are now),
// the rows' own styles, the selection, and formulas written once for the
// cells that share them.
//
// Usage: dotnet fsi tests/office-formats/xlsx-regenerated.fsx

open System
open System.IO
open System.IO.Compression
open System.Text.RegularExpressions
module Bridge = Rbt.Spreadsheet.XlsxBridge
module M = Rbt.Spreadsheet.Model
module Ops = Rbt.Spreadsheet.SheetOps

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
let element (name: string) (xml: string) =
    let found = Regex.Match(xml, sprintf "<%s(?=[\\s>/])[^>]*/>|<%s(?=[\\s>])[^>]*>[\\s\\S]*?</%s>" name name name)
    if found.Success then found.Value else ""
let count (pattern: string) (xml: string) = Regex.Matches(xml, pattern).Count
let short (value: string) = value.Substring(0, min 300 value.Length)
let addr row col : M.CellAddress = { Row = row; Col = col }
let rowTag (number: int) (xml: string) = (let found = Regex.Match(xml, sprintf "<row r=\"%d\"[^>]*>" number) in if found.Success then found.Value else "")
let rowTags (xml: string) = [ for found in Regex.Matches(xml, "<row\\b[^>]*>") -> Regex.Replace(found.Value, " spans=\"[^\"]*\"", "") ]
let withoutRow (number: int) (xml: string) = Regex.Replace(xml, sprintf "<row r=\"%d\"[^>]*>" number, "")
let attribute (name: string) (tag: string) = (let found = Regex.Match(tag, sprintf "\\s%s=\"([^\"]*)\"" name) in if found.Success then Some found.Groups.[1].Value else None)

let milk = File.ReadAllBytes(Path.Combine(root, "real-world/milk-street-office.xlsx"))
let workbook = XlsxReader.loadBytesNative milk |> Option.get
let imported = Bridge.Import.importWorkbook workbook
let index = imported.Sheets |> Array.findIndex (fun sheet -> sheet.Name = "Annual")
let sheetPath = (List.item index workbook.Sheets).Extras.SourcePart |> Option.get |> fst
let before = partsOf milk
let was = text before.[sheetPath]
let model0 = imported.Sheets.[index].Model
let exportWith (edit: M.SpreadsheetModel -> M.SpreadsheetModel) =
    let sheets = imported.Sheets |> Array.map (fun sheet -> sheet.Model)
    sheets.[index] <- edit sheets.[index]
    let input : Bridge.ExportInput =
        { Sheets = sheets; Names = imported.Sheets |> Array.map (fun sheet -> sheet.Name); Visible = imported.Sheets |> Array.map (fun sheet -> sheet.Visible)
          ActiveSheetIndex = workbook.Preserved.ActiveTab; NameInfo = imported.Context.NameInfo; DefinedNames = imported.Context.Names
          Source = Some imported.Source; Macros = None }
    let saved = XlsxWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut input)
    saved, sheets.[index]
let sheetOf (saved: byte[]) = text (partsOf saved).[sheetPath]
let reread (saved: byte[]) = (Bridge.Import.importWorkbook (XlsxReader.loadBytesNative saved |> Option.get)).Sheets.[index].Model
/// The formulas of a sheet, by cell.
let formulas (model: M.SpreadsheetModel) =
    model.Cells |> Map.toList |> List.choose (fun (at, cell) -> match cell.Value with M.CellValue.Formula (formula, _) -> Some (at, formula) | _ -> None)
let sameFormulas (label: string) (saved: byte[]) (edited: M.SpreadsheetModel) =
    let now, wanted = formulas (reread saved), formulas edited
    let wrong = List.zip now wanted |> List.filter (fun (a, b) -> a <> b)
    ok label (now.Length = wanted.Length && wrong.IsEmpty)
        (sprintf "%d formulas, %d wanted; first that differs: %A" now.Length wanted.Length (List.tryHead (if now.Length = wanted.Length then wrong else [])))
/// A part every element of which is where the schema has it.
let childOrder =
    [ "sheetPr"; "dimension"; "sheetViews"; "sheetFormatPr"; "cols"; "sheetData"; "sheetCalcPr"; "sheetProtection"; "protectedRanges"; "scenarios"
      "autoFilter"; "sortState"; "dataConsolidate"; "customSheetViews"; "mergeCells"; "phoneticPr"; "conditionalFormatting"; "dataValidations"
      "hyperlinks"; "printOptions"; "pageMargins"; "pageSetup"; "headerFooter"; "rowBreaks"; "colBreaks"; "customProperties"; "cellWatches"
      "ignoredErrors"; "smartTags"; "drawing"; "legacyDrawing"; "legacyDrawingHF"; "drawingHF"; "picture"; "oleObjects"; "controls"
      "webPublishItems"; "tableParts"; "extLst" ]
let inOrder (label: string) (xml: string) =
    let outside = Regex.Replace(xml, "<sheetData>[\\s\\S]*</sheetData>", "<sheetData/>")
    let doc = Xml.Linq.XDocument.Parse outside
    let names = [ for child in doc.Root.Elements() -> child.Name.LocalName ]
    let places = names |> List.map (fun name -> List.tryFindIndex ((=) name) childOrder |> Option.defaultValue -1)
    ok label (places |> List.forall (fun place -> place >= 0) && places = List.sort places) (String.concat " " names)
    // (... and the whole part is well-formed.)
    ok (label + " (well-formed)") (try (Xml.Linq.XDocument.Parse xml |> ignore; true) with _ -> false) "not XML"

ok "the file has what the checks are about"
    ((element "sheetPr" was).Contains "codeName" && element "ignoredErrors" was <> "" && count "t=\"shared\"" was > 100 && count "<row [^>]*\\bs=\"" was > 10
     && (element "sheetViews" was).Contains "<selection")
    (short (element "sheetViews" was))
let untouchedElements = [ "sheetPr"; "sheetViews"; "sheetFormatPr"; "cols"; "ignoredErrors"; "pageMargins"; "pageSetup"; "rowBreaks"; "mergeCells"; "dataValidations"; "extLst"; "legacyDrawing" ]

// ---- a row's height -----------------------------------------------------------
// Row 6 gets a height: its <row> says so, and nothing else of the part differs.
do
    let saved, _ = exportWith (M.setRowHeight 5 53.32)
    let now = sheetOf saved
    ok "row height: only that row's tag differs from the file's part" (withoutRow 6 now = withoutRow 6 was) (sprintf "%d chars, the file's %d" now.Length was.Length)
    ok "row height: the row states it" (attribute "ht" (rowTag 6 now) = Some "40" && attribute "customHeight" (rowTag 6 now) = Some "1") (rowTag 6 now)
    ok "row height: the row keeps its other attributes"
        (attribute "s" (rowTag 6 now) = attribute "s" (rowTag 6 was) && attribute "x14ac:dyDescent" (rowTag 6 now) = attribute "x14ac:dyDescent" (rowTag 6 was)) (rowTag 6 now + " <- " + rowTag 6 was)
    ok "row height: read back" (Map.tryFind 5 (reread saved).RowHeights |> Option.exists (fun size -> abs (size - 53.32) < 0.1)) (sprintf "%A" (Map.tryFind 5 (reread saved).RowHeights))

// ---- a hidden row, a group ----------------------------------------------------
do
    let saved, _ = exportWith (Ops.hideAxis true 9 9)
    let now = sheetOf saved
    ok "hidden row: only that row's tag differs from the file's part" (withoutRow 10 now = withoutRow 10 was) (sprintf "%d chars, the file's %d" now.Length was.Length)
    ok "hidden row: the row says so, and keeps what it said" (rowTag 10 now = (rowTag 10 was).Replace(">", " hidden=\"1\">")) (rowTag 10 now + " <- " + rowTag 10 was)
    ok "hidden row: read back" (Map.tryFind 9 (reread saved).RowHeights = Some 0.0) (sprintf "%A" (Map.tryFind 9 (reread saved).RowHeights))
    let grouped, _ = exportWith (fun model -> { model with Axes = { model.Axes with RowLevels = model.Axes.RowLevels |> Map.add 11 3 |> Map.add 12 3 } })
    let now = sheetOf grouped
    ok "grouped rows: the rows state their level" (attribute "outlineLevel" (rowTag 12 now) = Some "3" && attribute "outlineLevel" (rowTag 13 now) = Some "3") (rowTag 12 now)
    ok "grouped rows: sheetFormatPr states the deepest level, and what it said" ((element "sheetFormatPr" was).Contains "outlineLevelRow=\"2\"" && (element "sheetFormatPr" now) = (element "sheetFormatPr" was).Replace("outlineLevelRow=\"2\"", "outlineLevelRow=\"3\"")) (element "sheetFormatPr" now)
    let strip (xml: string) = xml |> withoutRow 12 |> withoutRow 13 |> fun rest -> rest.Replace(element "sheetFormatPr" rest, "")
    ok "grouped rows: the file has the rows at another level" (attribute "outlineLevel" (rowTag 12 was) = Some "1") (rowTag 12 was)
    ok "grouped rows: nothing else differs from the file's part" (strip now = strip was) (sprintf "%d chars, the file's %d" now.Length was.Length)
    ok "grouped rows: read back" ((reread grouped).Axes.RowLevels |> Map.tryFind 12 = Some 3) (sprintf "%A" (reread grouped).Axes.RowLevels)

// ---- the page setup -----------------------------------------------------------
do
    let saved, _ = exportWith (fun model -> { model with PageSetup = { model.PageSetup with Orientation = M.LandscapeOrientation } })
    let now = sheetOf saved
    ok "page setup: the part says landscape, and what it said of the printer"
        ((element "pageSetup" now).Contains "orientation=\"landscape\"" && (element "pageSetup" now).Contains "horizontalDpi=\"1200\"" && (element "pageSetup" now).Contains "r:id=\"rId1\"") (element "pageSetup" now)
    let strip (xml: string) = xml.Replace(element "pageSetup" xml, "")
    ok "page setup: nothing else differs from the file's part" (strip now = strip was) (sprintf "%d chars, the file's %d" now.Length was.Length)
    ok "page setup: read back" ((reread saved).PageSetup.Orientation = M.LandscapeOrientation) "portrait"

// ---- an AutoFilter ------------------------------------------------------------
do
    let range : M.CellRange = { From = addr 3 1; To = addr 20 4 }
    let saved, edited = exportWith (M.setAutoFilter (Some range))
    let now = sheetOf saved
    ok "auto filter: the part has it" ((element "autoFilter" now).Contains "ref=\"B4:E21\"") (element "autoFilter" now)
    let strip (xml: string) = match element "autoFilter" xml with "" -> xml | filter -> xml.Replace(filter, "")
    ok "auto filter: nothing else differs from the file's part" (strip now = strip was) (sprintf "%d chars, the file's %d" now.Length was.Length)
    inOrder "auto filter: the part's elements are in the schema's order" now
    ok "auto filter: read back" ((reread saved).AutoFilter |> Option.map (fun filter -> filter.Range) = (edited.AutoFilter |> Option.map (fun filter -> filter.Range))) (sprintf "%A" (reread saved).AutoFilter)

// ---- a cell whose formula other cells share -----------------------------------
// The cell that holds a shared formula gets a number: the next cell of the
// group holds the formula from now on, the others still take it from there.
do
    let master = Regex.Match(was, "<c r=\"([A-Z]+)([0-9]+)\"[^>]*><f t=\"shared\" ref=\"([A-Z]+[0-9]+):([A-Z]+[0-9]+)\" si=\"([0-9]+)\">")
    let column = master.Groups.[1].Value |> Seq.fold (fun acc ch -> acc * 26 + int ch - int 'A' + 1) 0
    let at = addr (int master.Groups.[2].Value - 1) (column - 1)
    let saved, edited = exportWith (M.setCell at { Value = M.CellValue.Number 77.0; Format = (M.getCell at model0 |> Option.get).Format; Comment = None })
    let now = sheetOf saved
    ok "shared formula's cell edited: the other cells still share the formula" (count "t=\"shared\"" now = count "t=\"shared\"" was - 1)
        (sprintf "%d shared, the file has %d (%s%s)" (count "t=\"shared\"" now) (count "t=\"shared\"" was) master.Groups.[1].Value master.Groups.[2].Value)
    ok "shared formula's cell edited: the rows keep their attributes" (rowTags now = rowTags was) "row tags differ"
    for name in untouchedElements do
        ok (sprintf "shared formula's cell edited: <%s> is the file's" name) (element name now = element name was) (short (element name now))
    sameFormulas "shared formula's cell edited: every formula reads as it did" saved edited
    ok "shared formula's cell edited: the cell holds the number" ((M.getCell at (reread saved) |> Option.map (fun cell -> cell.Value)) = Some (M.CellValue.Number 77.0)) "not 77"

// ---- a link, a note -----------------------------------------------------------
// They are in other parts, or name one: the rest of the part is made from the
// sheet, the rows are the file's, and what the part said beyond a sheet stays.
let keptOutsideRows (label: string) (now: string) =
    ok (label + ": sheetPr is the file's") (element "sheetPr" now = element "sheetPr" was) (element "sheetPr" now)
    ok (label + ": the views are the file's (selection, where it is scrolled to)") (element "sheetViews" now = element "sheetViews" was) (short (element "sheetViews" now))
    ok (label + ": sheetFormatPr is the file's") (element "sheetFormatPr" now = element "sheetFormatPr" was) (element "sheetFormatPr" now)
    ok (label + ": ignoredErrors are the file's") (element "ignoredErrors" now = element "ignoredErrors" was) (short (element "ignoredErrors" now))
    ok (label + ": the worksheet's namespaces are the file's") (Regex.Match(now, "<worksheet[^>]*>").Value = Regex.Match(was, "<worksheet[^>]*>").Value) (Regex.Match(now, "<worksheet[^>]*>").Value)
    inOrder (label + ": the part's elements are in the schema's order") now
do
    let range : M.CellRange = { From = addr 400 1; To = addr 400 1 }
    let saved, edited = exportWith (Ops.setHyperlink [ range ] "https://example.com/a?b=1&c=2" false (Some "a link") None)
    let now = sheetOf saved
    ok "link added: the part names it" ((element "hyperlinks" now).Contains "ref=\"B401\"") (element "hyperlinks" now)
    ok "link added: the rows keep their attributes" (rowTags now |> List.filter (fun tag -> not (tag.Contains "r=\"401\"")) = rowTags was) "row tags differ"
    ok "link added: the shared formulas stay shared" (count "t=\"shared\"" now = count "t=\"shared\"" was) (sprintf "%d of %d" (count "t=\"shared\"" now) (count "t=\"shared\"" was))
    keptOutsideRows "link added" now
    sameFormulas "link added: every formula reads as it did" saved edited
    ok "link added: read back" ((reread saved).Hyperlinks |> List.exists (fun link -> link.Target = "https://example.com/a?b=1&c=2" && link.Range = range)) (sprintf "%A" (reread saved).Hyperlinks)
    let noted, edited = exportWith (Ops.setComment (addr 401 2) "a note" (Some "me"))
    let now = sheetOf noted
    ok "note added: the rows keep their attributes" (rowTags now |> List.filter (fun tag -> not (tag.Contains "r=\"402\"")) = rowTags was) "row tags differ"
    ok "note added: the shared formulas stay shared" (count "t=\"shared\"" now = count "t=\"shared\"" was) (sprintf "%d of %d" (count "t=\"shared\"" now) (count "t=\"shared\"" was))
    keptOutsideRows "note added" now
    sameFormulas "note added: every formula reads as it did" noted edited
    ok "note added: read back" ((M.getCell (addr 401 2) (reread noted) |> Option.bind (fun cell -> cell.Comment)).IsSome) "no note"

// ---- the comments part: a note that is what it was is the file's element -----------
do
    let commentsOf (saved: byte[]) = text (partsOf saved).["xl/comments1.xml"]
    let commentsWas = text before.["xl/comments1.xml"]
    let noteOf (reference: string) (xml: string) = Regex.Match(xml, sprintf "<comment ref=\"%s\"[^>]*>[\\s\\S]*?</comment>" reference).Value
    let fileNote = noteOf "E113" commentsWas
    ok "the file has a note in runs, with what a sheet does not hold of it" (fileNote.Contains "<color indexed=\"81\"/>" && fileNote.Contains "<family val=\"2\"/>" && fileNote.Contains "shapeId=\"0\"" && fileNote.Contains "xr:uid=") (short fileNote)
    let noted, _ = exportWith (Ops.setComment (addr 401 2) "a note" (Some "me"))
    let now = commentsOf noted
    ok "note added: the file's note is the element the file wrote" (noteOf "E113" now = fileNote) (short (noteOf "E113" now))
    ok "note added: the part's root and its authors are the file's, the new author behind them"
        (Regex.Match(now, "<comments[^>]*>").Value = Regex.Match(commentsWas, "<comments[^>]*>").Value && now.Contains "<authors><author>BIWS</author><author>me</author></authors>") (short now)
    ok "note added: the new note names its author" ((noteOf "C402" now).Contains "authorId=\"1\"" && (noteOf "C402" now).Contains "a note") (noteOf "C402" now)
    ok "note added: the part is the file's but for the author and the note" (now.Replace("<author>me</author>", "").Replace(noteOf "C402" now, "") = commentsWas) (sprintf "%d chars, the file's %d" now.Length commentsWas.Length)
    ok "note added: read back, both" (let again = reread noted in (M.getCell (addr 401 2) again |> Option.bind (fun cell -> cell.Comment)).IsSome && (M.getCell (addr 112 4) again |> Option.bind (fun cell -> cell.Comment)).IsSome) "a note is missing"
    // The VML drawing: the file's shapes stay, the new note has a box.
    let vmlPath = before |> Map.toList |> List.map fst |> List.find (fun path -> path.Contains "vmlDrawing")
    let vmlWas, vmlNow = text before.[vmlPath], text (partsOf noted).[vmlPath]
    let shapes (xml: string) = [ for found in Regex.Matches(xml, "<v:shape\\b[\\s\\S]*?</v:shape>") -> found.Value ]
    ok "note added: the file's comment boxes are the shapes the file wrote, and there is one more" (shapes vmlWas |> List.forall (fun shape -> List.contains shape (shapes vmlNow)) && (shapes vmlNow).Length = (shapes vmlWas).Length + 1) (sprintf "%d shapes, the file has %d" (shapes vmlNow).Length (shapes vmlWas).Length)
    // A row inserted above the note: the element is the file's, in the cell the note is in now.
    let moved, _ = exportWith (M.insertRow 100)
    ok "row inserted: the note is the file's element, one row down" (noteOf "E114" (commentsOf moved) = fileNote.Replace("ref=\"E113\"", "ref=\"E114\"") && commentsOf moved = commentsWas.Replace("ref=\"E113\"", "ref=\"E114\"")) (short (commentsOf moved))
    // The note's text changed: that note is written from the sheet, a second note of the file would stay.
    let changed, _ = exportWith (Ops.setComment (addr 112 4) "another text" (Some "BIWS") >> Ops.setComment (addr 401 2) "a note" None)
    let now = commentsOf changed
    ok "note changed: it is written from the sheet, by the file's author" ((noteOf "E113" now).Contains "another text" && (noteOf "E113" now).Contains "authorId=\"0\"" && not ((noteOf "E113" now).Contains "Based on actual")) (short (noteOf "E113" now))
    ok "note changed: the part is still XML" (try (Xml.Linq.XDocument.Parse now |> ignore; true) with _ -> false) (short now)

// ---- rows and columns inserted and deleted --------------------------------------
// The sheet is written from its cells, and says what its part said of rows
// and cells where they are now.
let movedRow (insertAt: int) (countRows: int) (number: int) = if number - 1 >= insertAt then number + countRows else number
do
    let saved, edited = exportWith (M.insertRow 100)
    let now = sheetOf saved
    ok "row inserted: sheetPr is the file's (the code name)" (element "sheetPr" now = element "sheetPr" was) (element "sheetPr" now)
    ok "row inserted: sheetFormatPr is the file's" (element "sheetFormatPr" now = element "sheetFormatPr" was) (element "sheetFormatPr" now)
    ok "row inserted: the worksheet's namespaces are the file's" (Regex.Match(now, "<worksheet[^>]*>").Value = Regex.Match(was, "<worksheet[^>]*>").Value) (Regex.Match(now, "<worksheet[^>]*>").Value)
    // ignoredErrors: E96:E109 is E96:E100 and below; F137:K157 moved down by one.
    let ignored = element "ignoredErrors" now
    ok "row inserted: ignoredErrors follow their cells"
        (ignored.Contains "E96:E110" && ignored.Contains "F138:K158" && ignored.Contains "J13 E22:E23" && ignored.Contains "<ignoredError sqref=\"H242:J242\" formula=\"1\"/>"
         && count "<ignoredError " ignored = count "<ignoredError " (element "ignoredErrors" was))
        (short ignored)
    // Every row of the file that had a style of its own has it where the row is now.
    let styled (xml: string) = [ for found in Regex.Matches(xml, "<row r=\"([0-9]+)\"[^>]*>") do
                                    match attribute "s" found.Value with
                                    | Some style -> yield int found.Groups.[1].Value, style, attribute "customFormat" found.Value
                                    | None -> () ]
    let wanted = styled was |> List.map (fun (number, style, custom) -> movedRow 100 1 number, style, custom)
    ok "row inserted: the rows keep their own styles" (styled now = wanted && not wanted.IsEmpty) (sprintf "%d styled rows, the file has %d; %A" (styled now).Length wanted.Length (List.tryHead (styled now)))
    // The view is scrolled to row 308 with B308 selected: one further down now.
    ok "row inserted: the selection and where the view is scrolled to follow"
        ((element "sheetViews" now).Contains "topLeftCell=\"A309\"" && (element "sheetViews" now).Contains "activeCell=\"B309\"" && (element "sheetViews" now).Contains "sqref=\"B309\"") (short (element "sheetViews" now))
    ok "row inserted: formulas that read alike are written once" (count "t=\"shared\"" now * 10 >= count "t=\"shared\"" was * 8)
        (sprintf "%d shared, the file has %d" (count "t=\"shared\"" now) (count "t=\"shared\"" was))
    sameFormulas "row inserted: every formula reads as the sheet has it" saved edited
    inOrder "row inserted: the part's elements are in the schema's order" now
    ok "row inserted: the rows after it keep the heights the file gave them"
        (let heights (xml: string) = [ for found in Regex.Matches(xml, "<row r=\"([0-9]+)\"[^>]*>") do
                                        match attribute "ht" found.Value with
                                        | Some height -> yield int found.Groups.[1].Value, float height, attribute "customHeight" found.Value, attribute "hidden" found.Value
                                        | None -> () ]
         let wanted = heights was |> List.map (fun (number, height, custom, hidden) -> movedRow 100 1 number, height, custom, hidden)
         let close (a: (int * float * string option * string option) list) (b: (int * float * string option * string option) list) =
            a.Length = b.Length && List.forall2 (fun (n1, h1, c1, x1) (n2, h2, c2, x2) -> n1 = n2 && abs (h1 - h2) < 0.01 && c1 = c2 && x1 = x2) a b
         close (heights now) wanted && wanted.Length > 5)
        (String.concat " " [ for found in Regex.Matches(now, "<row r=\"([0-9]+)\"[^>]*>") do if (attribute "ht" found.Value).IsSome then yield found.Value ] |> short)
// ---- a row inserted: the styles, the shared strings, the calculation chain ---------
do
    let saved, _ = exportWith (M.insertRow 100)
    let parts = partsOf saved
    let styles, stylesWas = text parts.["xl/styles.xml"], text before.["xl/styles.xml"]
    let xfs (xml: string) = count "<xf[\s>/]" (element "cellXfs" xml)
    ok "row inserted: the cells that moved have the file's formats (no cellXfs entry is added)" (xfs styles = xfs stylesWas) (sprintf "%d entries, the file has %d" (xfs styles) (xfs stylesWas))
    ok "row inserted: the styles part is the file's" (styles = stylesWas) (sprintf "%d chars, the file's %d" styles.Length stylesWas.Length)
    // The shared strings' count is the number of cells that name one.
    let uses (all: Map<string, byte[]>) = all |> Map.toList |> List.sumBy (fun (path, data) -> if path.StartsWith "xl/worksheets/sheet" then count " t=\"s\"" (text data) else 0)
    let said (all: Map<string, byte[]>) = attribute "count" (Regex.Match(text all.["xl/sharedStrings.xml"], "<sst[^>]*>").Value) |> Option.map int
    ok "row inserted: the shared strings' count is the cells that use them" (said parts = Some (uses parts)) (sprintf "count %A, %d uses (the file: %A, %d)" (said parts) (uses parts) (said before) (uses before))
    // The calculation chain lists cells that moved: the part goes, with its content type and its relationship.
    ok "the file has a calculation chain" (before.ContainsKey "xl/calcChain.xml") ""
    ok "row inserted: no calculation chain, no content type and no relationship of one"
        (not (parts.ContainsKey "xl/calcChain.xml") && not ((text parts.["[Content_Types].xml"]).Contains "calcChain") && not ((text parts.["xl/_rels/workbook.xml.rels"]).Contains "calcChain"))
        (sprintf "part %b, type %b, relationship %b" (parts.ContainsKey "xl/calcChain.xml") ((text parts.["[Content_Types].xml"]).Contains "calcChain") ((text parts.["xl/_rels/workbook.xml.rels"]).Contains "calcChain"))
    // Text typed into two cells: two more uses, one more string.
    let typed, _ = exportWith (M.setCell (addr 600 1) { Value = M.CellValue.String "a text no cell has"; Format = M.defaultFormat; Comment = None } >> M.setCell (addr 601 1) { Value = M.CellValue.String "a text no cell has"; Format = M.defaultFormat; Comment = None })
    let parts = partsOf typed
    ok "text typed twice: the shared strings' count is the cells that use them" (said parts = Some (uses parts) && uses parts = uses before + 2) (sprintf "count %A, %d uses (the file: %A, %d)" (said parts) (uses parts) (said before) (uses before))

// ---- manual page breaks -------------------------------------------------------------
do
    let breaksWas = [ for found in Regex.Matches(element "rowBreaks" was, "<brk\\b[^>]*/>") -> found.Value ]
    ok "the file has manual row breaks" (not breaksWas.IsEmpty) (element "rowBreaks" was)
    let saved, _ = exportWith (fun model -> { model with PageSetup = { model.PageSetup with ManualRowBreaks = List.sort (450 :: model.PageSetup.ManualRowBreaks); ManualColBreaks = [ 3 ] } })
    let now = sheetOf saved
    let breaks = [ for found in Regex.Matches(element "rowBreaks" now, "<brk\\b[^>]*/>") -> found.Value ]
    ok "break added: the file's breaks are as the file wrote them" (breaksWas |> List.forall (fun tag -> List.contains tag breaks) && breaks.Length = breaksWas.Length + 1) (element "rowBreaks" now)
    ok "break added: a row break reaches as far as the file's do, a column break over the rows"
        ((breaks |> List.exists (fun tag -> attribute "id" tag = Some "450" && attribute "max" tag = Some "11" && attribute "man" tag = Some "1")) && (element "colBreaks" now).Contains "<brk id=\"3\" max=\"1048575\" man=\"1\"/>")
        (element "rowBreaks" now + element "colBreaks" now)

do
    let saved, edited = exportWith (M.deleteRow 100 >> M.deleteRow 100)
    let now = sheetOf saved
    let ignored = element "ignoredErrors" now
    ok "rows deleted: ignoredErrors follow their cells (a range shrinks)" (ignored.Contains "E96:E107" && ignored.Contains "F135:K155" && ignored.Contains "J13 E22:E23") (short ignored)
    ok "rows deleted: sheetPr is the file's" (element "sheetPr" now = element "sheetPr" was) (element "sheetPr" now)
    sameFormulas "rows deleted: every formula reads as the sheet has it" saved edited
    inOrder "rows deleted: the part's elements are in the schema's order" now
do
    let saved, edited = exportWith (M.insertColumn 4)
    let now = sheetOf saved
    let ignored = element "ignoredErrors" now
    ok "column inserted: ignoredErrors follow their cells" (ignored.Contains "K13 F22:F23" && ignored.Contains "G87:L91" && ignored.Contains "D348:L364") (short ignored)
    ok "column inserted: the rows keep their own styles"
        (count "<row [^>]*\\bs=\"" now = count "<row [^>]*\\bs=\"" was && attribute "s" (rowTag 6 now) = attribute "s" (rowTag 6 was)) (rowTag 6 now)
    ok "column inserted: formulas that read alike are written once" (count "t=\"shared\"" now * 10 >= count "t=\"shared\"" was * 8)
        (sprintf "%d shared, the file has %d" (count "t=\"shared\"" now) (count "t=\"shared\"" was))
    sameFormulas "column inserted: every formula reads as the sheet has it" saved edited
    inOrder "column inserted: the part's elements are in the schema's order" now

// ---- a note without an author ------------------------------------------------------
// It is no one's: the authors list has an empty entry for it (not the first author of the list).
do
    let bytes = File.ReadAllBytes(Path.Combine(root, "test-documents-excel/synthetic-sparklines.xlsx"))
    let opened = Bridge.Import.importWorkbook (XlsxReader.loadBytesNative bytes |> Option.get)
    let edited = M.resolveSheet opened.Sheets.[0].Model |> Ops.setComment (addr 5 0) "signed" (Some "Ann") |> Ops.setComment (addr 6 0) "unsigned" None
    let input : Bridge.ExportInput =
        { Sheets = [| edited |]; Names = [| opened.Sheets.[0].Name |]; Visible = [| true |]
          ActiveSheetIndex = 0; NameInfo = opened.Context.NameInfo; DefinedNames = opened.Context.Names; Source = Some opened.Source; Macros = None }
    let saved = XlsxWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut input)
    let comments = partsOf saved |> Map.toList |> List.pick (fun (path, data) -> if path.StartsWith "xl/comments" then Some (text data) else None)
    let authors = [ for found in Regex.Matches(comments, "<author(?:/>|>([^<]*)</author>)") -> found.Groups.[1].Value ]
    let authorIdOf (reference: string) = int (Regex.Match(comments, sprintf "<comment ref=\"%s\" authorId=\"([0-9]+)\"" reference).Groups.[1].Value)
    ok "note without an author: the authors list has an empty entry, and the note names it"
        (List.contains "" authors && List.contains "Ann" authors && List.item (authorIdOf "A7") authors = "" && List.item (authorIdOf "A6") authors = "Ann") (sprintf "%A; A6 %d, A7 %d" authors (authorIdOf "A6") (authorIdOf "A7"))
    let again = (Bridge.Import.importWorkbook (XlsxReader.loadBytesNative saved |> Option.get)).Sheets.[0].Model |> M.resolveSheet
    let authorAt at = M.getCell at again |> Option.bind (fun cell -> cell.Comment) |> Option.map (fun note -> note.Author |> Option.defaultValue "")
    ok "note without an author: read back, it is not the other note's author's" (authorAt (addr 6 0) = Some "" && authorAt (addr 5 0) = Some "Ann") (sprintf "%A %A" (authorAt (addr 5 0)) (authorAt (addr 6 0)))

// ---- sparklines follow their cells ----------------------------------------------
do
    let bytes = File.ReadAllBytes(Path.Combine(root, "test-documents-excel/synthetic-sparklines.xlsx"))
    let book = XlsxReader.loadBytesNative bytes |> Option.get
    let opened = Bridge.Import.importWorkbook book
    let save (edit: M.SpreadsheetModel -> M.SpreadsheetModel) =
        let input : Bridge.ExportInput =
            { Sheets = [| edit (M.resolveSheet opened.Sheets.[0].Model) |]; Names = [| opened.Sheets.[0].Name |]; Visible = [| true |]
              ActiveSheetIndex = 0; NameInfo = opened.Context.NameInfo; DefinedNames = opened.Context.Names; Source = Some opened.Source; Macros = None }
        text (partsOf (XlsxWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut input))).["xl/worksheets/sheet1.xml"]
    let sparklines (xml: string) = [ for found in Regex.Matches(xml, "<x14:sparkline><xm:f>([^<]*)</xm:f><xm:sqref>([^<]*)</xm:sqref></x14:sparkline>") -> found.Groups.[1].Value, found.Groups.[2].Value ]
    ok "sparklines: the file has three" (sparklines (text (partsOf bytes).["xl/worksheets/sheet1.xml"]) = [ "Sheet1!B2:H2", "A2"; "Sheet1!B3:H3", "A3"; "Sheet1!B4:H4", "A4" ]) ""
    let inserted = save (M.insertRow 0)
    ok "sparklines: a row inserted above: they and their data are one row down" (sparklines inserted = [ "Sheet1!B3:H3", "A3"; "Sheet1!B4:H4", "A4"; "Sheet1!B5:H5", "A5" ]) (sprintf "%A" (sparklines inserted))
    inOrder "sparklines: a row inserted: the part's elements are in the schema's order" inserted
    let wider = save (M.insertColumn 3)
    ok "sparklines: a column inserted in their data: the data is one column wider" (sparklines wider = [ "Sheet1!B2:I2", "A2"; "Sheet1!B3:I3", "A3"; "Sheet1!B4:I4", "A4" ]) (sprintf "%A" (sparklines wider))
    let fewer = save (M.deleteRow 2)
    ok "sparklines: a row deleted: its sparkline and its group are gone, the one below moved up"
        (sparklines fewer = [ "Sheet1!B2:H2", "A2"; "Sheet1!B3:H3", "A3" ] && count "<x14:sparklineGroup[ >]" fewer = 2 && fewer.Contains "type=\"stacked\"" && not (fewer.Contains "type=\"column\""))
        (sprintf "%A, %d groups" (sparklines fewer) (count "<x14:sparklineGroup[ >]" fewer))
    let none = save (M.deleteRow 1 >> M.deleteRow 1 >> M.deleteRow 1)
    ok "sparklines: every row of them deleted: no sparklines, and no empty list of them" (not (none.Contains "sparkline") && (try (Xml.Linq.XDocument.Parse none |> ignore; true) with _ -> false)) (short (none.Substring(none.IndexOf "</sheetData>")))

// ---- every workbook -------------------------------------------------------------
// A link added (the rows are the file's, the rest of the part is made from
// the sheet), a row hidden and another page orientation (the part with that
// changed in it), a row and a column inserted (written from the cells): the
// parts are XML with their elements in the schema's order, the file opens,
// and every formula of the sheet reads as the sheet has it.
let mutable books = 0
for dir in [ "real-world"; "test-documents-excel" ] do
    for file in Directory.GetFiles(Path.Combine(root, dir)) |> Array.sort do
        if file.EndsWith ".xlsx" || file.EndsWith ".xlsm" then
            let name = Path.GetFileName file
            match (try XlsxReader.loadBytesNative (File.ReadAllBytes file) with _ -> None) with
            | None -> ()
            | Some book ->
                match (try Some (Bridge.Import.importWorkbook book) with _ -> None) with
                | None -> ()
                | Some opened ->
                    let target =
                        book.Sheets |> List.tryFindIndex (fun sheet ->
                            match sheet.Extras.SourcePart with
                            | Some (_, relType) -> relType.EndsWith "/worksheet" && not sheet.Extras.Verbatim
                            | None -> false)
                    match target with
                    | Some target when target < opened.Sheets.Length ->
                        books <- books + 1
                        let path = (List.item target book.Sheets).Extras.SourcePart |> Option.get |> fst
                        let start = M.resolveSheet opened.Sheets.[target].Model
                        let placesOf (xml: string) =
                            let names = [ for child in (Xml.Linq.XDocument.Parse xml).Root.Elements() -> child.Name.LocalName ]
                            names, names |> List.map (fun name -> List.tryFindIndex ((=) name) childOrder |> Option.defaultValue -1)
                        // (A file whose own part is not in that order: its order is kept.)
                        let fileInOrder = (let _, places = placesOf (text (partsOf (File.ReadAllBytes file)).[path]) in places |> List.forall (fun place -> place >= 0) && places = List.sort places)
                        let edits : (string * (M.SpreadsheetModel -> M.SpreadsheetModel)) list =
                            [ "a link", Ops.setHyperlink [ { From = addr 2000 30; To = addr 2000 30 } ] "https://example.com/" false (Some "link") None
                              "a hidden row", (Ops.hideAxis true 1 1 >> fun model -> { model with PageSetup = { model.PageSetup with Orientation = M.LandscapeOrientation } })
                              "a row inserted", M.insertRow 1
                              "a column deleted and one inserted", (M.deleteColumn 40 >> M.insertColumn 1) ]
                        for label, edit in edits do
                            let label = name + ": " + label
                            try
                                let sheets = opened.Sheets |> Array.map (fun sheet -> sheet.Model)
                                let edited = edit start
                                sheets.[target] <- edited
                                let input : Bridge.ExportInput =
                                    { Sheets = sheets; Names = opened.Sheets |> Array.map (fun sheet -> sheet.Name); Visible = opened.Sheets |> Array.map (fun sheet -> sheet.Visible)
                                      ActiveSheetIndex = book.Preserved.ActiveTab; NameInfo = opened.Context.NameInfo; DefinedNames = opened.Context.Names
                                      Source = Some opened.Source; Macros = None }
                                let saved = XlsxWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut input)
                                let xml = text (partsOf saved).[path]
                                let names, places = placesOf xml
                                // The sheet's protection is the file's whatever was changed.
                                let fileXml = text (partsOf (File.ReadAllBytes file)).[path]
                                if element "sheetProtection" fileXml <> "" then
                                    ok (label + ": the sheet's protection is the file's") (element "sheetProtection" xml = element "sheetProtection" fileXml) (element "sheetProtection" xml)
                                ok (label + ": the part's elements are in the schema's order") (not fileInOrder || (places |> List.forall (fun place -> place >= 0) && places = List.sort places)) (String.concat " " names)
                                let again = M.resolveSheet (Bridge.Import.importWorkbook (XlsxReader.loadBytesNative saved |> Option.get)).Sheets.[target].Model
                                let now, wanted = formulas again, formulas edited
                                let wrong = if now.Length = wanted.Length then List.zip now wanted |> List.filter (fun (a, b) -> a <> b) else []
                                ok (label + ": every formula reads as the sheet has it") (now.Length = wanted.Length && wrong.IsEmpty)
                                    (sprintf "%d formulas, %d wanted; first that differs: %A" now.Length wanted.Length (List.tryHead wrong))
                            with error -> ok (label + ": saved") false (error.Message.Split('\n').[0])
                    | _ -> ()

printfn "\nXLSX regenerated-sheet tests: %d passed, %d failed (%d workbooks)" pass fail books
if fail > 0 then exit 1
