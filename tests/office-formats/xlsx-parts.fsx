#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

// Nothing is lost on a save: after an unedited export every part of the
// opened package is in the written one (under its own name), with a content
// type and reachable through relationships, for every workbook under
// tests/real-world and tests/test-documents-excel.
//
// Usage: dotnet fsi tests/office-formats/xlsx-parts.fsx [--list] [file.xlsx ...]
// (--list prints what is missing per workbook instead of the first few.)

open System
open System.IO
open System.IO.Compression
module Bridge = Rbt.Spreadsheet.XlsxBridge

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-34s  %s" label detail)

let root = Path.Combine(__SOURCE_DIRECTORY__, "..")
let args = fsi.CommandLineArgs |> Array.skip 1
let listAll = args |> Array.contains "--list"
let named = args |> Array.filter (fun a -> a <> "--list")
let files =
    if named.Length > 0 then named |> Array.map Path.GetFullPath
    else
        [| for dir in [ "real-world"; "test-documents-excel" ] do
            for file in Directory.GetFiles(Path.Combine(root, dir)) |> Array.sort do
                if file.EndsWith ".xlsx" || file.EndsWith ".xlsm" then yield file |]

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

/// The targets of a .rels part, resolved to part names (internal ones).
let relTargets (relsPath: string) (xml: string) : (string * string) list =
    let baseDir =
        let dir = Path.GetDirectoryName(relsPath).Replace('\\', '/')            // ".../_rels"
        let parent = if dir.EndsWith "_rels" then dir.Substring(0, dir.Length - 5).TrimEnd('/') else dir
        parent
    let resolve (target: string) =
        if target.StartsWith "/" then target.TrimStart('/')
        else
            let stack = Collections.Generic.List<string>(baseDir.Split([| '/' |], StringSplitOptions.RemoveEmptyEntries))
            for piece in target.Split('/') do
                if piece = ".." then (if stack.Count > 0 then stack.RemoveAt(stack.Count - 1))
                elif piece <> "." && piece <> "" then stack.Add piece
            String.Join("/", stack)
    [ for m in Text.RegularExpressions.Regex.Matches(xml, "<Relationship\\b[^>]*>") do
        let attr (name: string) =
            let found = Text.RegularExpressions.Regex.Match(m.Value, "\\b" + name + "=\"([^\"]*)\"")
            if found.Success then found.Groups.[1].Value else ""
        if attr "TargetMode" <> "External" then yield attr "Type", resolve (Uri.UnescapeDataString (attr "Target")) ]

/// Every part reachable from the package's root relationships.
let reachable (parts: Map<string, byte[]>) : Set<string> =
    let seen = Collections.Generic.HashSet<string>()
    let rec visit (relsPath: string) =
        match Map.tryFind relsPath parts with
        | Some data ->
            for relType, target in relTargets relsPath (text data) do
                ignore relType
                if seen.Add target then
                    let dir = Path.GetDirectoryName(target).Replace('\\', '/')
                    visit ((if dir = "" then "" else dir + "/") + "_rels/" + Path.GetFileName target + ".rels")
        | None -> ()
    visit "_rels/.rels"
    Set.ofSeq seen

/// Every part reachable from one part.
let reachableFrom (parts: Map<string, byte[]>) (start: string) : Set<string> =
    let seen = Collections.Generic.HashSet<string>()
    let rec visit (part: string) =
        let dir = Path.GetDirectoryName(part).Replace('\\', '/')
        let relsPath = (if dir = "" then "" else dir + "/") + "_rels/" + Path.GetFileName part + ".rels"
        match Map.tryFind relsPath parts with
        | Some data -> for _, target in relTargets relsPath (text data) do if seen.Add target then visit target
        | None -> ()
    visit start
    Set.ofSeq seen

let hasContentType (types: string) (part: string) =
    types.Contains ("PartName=\"/" + part + "\"")
    || (let ext = Path.GetExtension(part).TrimStart('.') in ext <> "" && Text.RegularExpressions.Regex.IsMatch(types, "<Default\\b[^>]*Extension=\"" + Text.RegularExpressions.Regex.Escape ext + "\"", Text.RegularExpressions.RegexOptions.IgnoreCase))

// Parts the writer makes itself (their content is the writer's; calcChain is
// derived data Excel rebuilds).
let regenerated (part: string) = part = "xl/calcChain.xml"

let mutable books = 0
let mutable skipped = 0
for file in files do
    let name = Path.GetFileName file
    let original = try Some (File.ReadAllBytes file |> partsOf) with _ -> None
    match original with
    | None -> skipped <- skipped + 1          // not a zip package (an encrypted or deliberately broken test file)
    | Some original ->
        match (try XlsxReader.loadBytesNative (File.ReadAllBytes file) with _ -> None) with
        | None -> skipped <- skipped + 1
        | Some workbook ->
            books <- books + 1
            try
                let imported = Bridge.Import.importWorkbook workbook
                // (The sheet on display stays the one the file shows.)
                let activeTab =
                    let found = Text.RegularExpressions.Regex.Match(text original.["xl/workbook.xml"], "activeTab=\"(\\d+)\"")
                    if found.Success then min (int found.Groups.[1].Value) (imported.Sheets.Length - 1) else 0
                let input : Bridge.ExportInput =
                    { Sheets = imported.Sheets |> Array.map (fun sheet -> sheet.Model)
                      Names = imported.Sheets |> Array.map (fun sheet -> sheet.Name)
                      Visible = imported.Sheets |> Array.map (fun sheet -> sheet.Visible)
                      ActiveSheetIndex = activeTab
                      NameInfo = Map.empty; DefinedNames = imported.Context.Names
                      Source = Some imported.Source
                      Macros = None; MacroProject = None }
                let written = XlsxPackageWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut input) |> partsOf
                // Only what the original package itself reaches counts (a stray entry nothing refers to is not a part).
                let wanted = reachable original |> Set.filter (fun part -> Map.containsKey part original && not (regenerated part))
                let missing = wanted |> Set.filter (fun part -> not (Map.containsKey part written)) |> Set.toList
                let show (items: string list) = (if listAll then items else List.truncate 6 items) |> String.concat ", " |> fun s -> sprintf "%d: %s" items.Length s
                ok (name + ": parts") missing.IsEmpty (show missing)
                // What the writer does not make from the model goes back byte for byte: drawings, tables,
                // charts, media, pivot and slicer parts, the theme, document properties, embedded data ...
                let rewritten (part: string) =
                    [ "[Content_Types].xml"; "_rels/.rels"; "xl/workbook.xml"; "xl/_rels/workbook.xml.rels"; "xl/styles.xml"; "xl/sharedStrings.xml"; "xl/metadata.xml" ] |> List.contains part
                    || [ "xl/worksheets/"; "xl/comments"; "xl/threadedComments/"; "xl/persons/" ] |> List.exists (fun prefix -> part.StartsWith prefix)
                let changed =
                    wanted |> Set.toList
                    |> List.filter (fun part -> not (rewritten part) && (match Map.tryFind part written with Some data -> data <> original.[part] | None -> false))
                ok (name + ": untouched parts identical") changed.IsEmpty (show changed)
                // The sheets, chart sheets among them, are listed in the file's order, each kind of sheet as what it was.
                let sheetList (parts: Map<string, byte[]>) =
                    let rels = relTargets "xl/_rels/workbook.xml.rels" (text parts.["xl/_rels/workbook.xml.rels"])
                    let ids =
                        [ for m in Text.RegularExpressions.Regex.Matches(text parts.["xl/_rels/workbook.xml.rels"], "<Relationship\\b[^>]*>") do
                            let attr (name: string) = Text.RegularExpressions.Regex.Match(m.Value, "\\b" + name + "=\"([^\"]*)\"").Groups.[1].Value
                            yield attr "Id", attr "Type" ] |> Map.ofList
                    ignore rels
                    [ for m in Text.RegularExpressions.Regex.Matches(text parts.["xl/workbook.xml"], "<sheet\\b[^>]*>") do
                        let attr (name: string) = Text.RegularExpressions.Regex.Match(m.Value, "\\b" + name + "=\"([^\"]*)\"").Groups.[1].Value
                        let relType = Map.tryFind (attr "r:id") ids |> Option.defaultValue "?"
                        yield attr "name", relType.Substring(relType.LastIndexOf '/' + 1) ]
                if original.ContainsKey "xl/workbook.xml" && original.ContainsKey "xl/_rels/workbook.xml.rels" then
                    ok (name + ": sheets") (sheetList original = sheetList written) (sprintf "%A / %A" (sheetList original) (sheetList written))
                let reached = reachable written
                let unreached = wanted |> Set.filter (fun part -> Map.containsKey part written && not (Set.contains part reached)) |> Set.toList
                ok (name + ": related") unreached.IsEmpty (show unreached)
                let types = match Map.tryFind "[Content_Types].xml" written with Some data -> text data | None -> ""
                let untyped = written |> Map.toList |> List.map fst |> List.filter (fun part -> part <> "[Content_Types].xml" && not (hasContentType types part))
                ok (name + ": content types") untyped.IsEmpty (show untyped)
                // One cell of the first worksheet edited: every other worksheet goes back as the part it was
                // read from, byte for byte, and still means what it meant - the styles and shared strings
                // its cells index are the file's, added to. (Not for a workbook in the 1904 date system: its
                // serials are converted on import, so every sheet is written from the model.)
                let isWorksheet (sheet: XlsxReader.Worksheet) = match sheet.Extras.SourcePart with Some (_, relType) -> relType.EndsWith "/worksheet" | None -> false
                let sheetsRead = List.toArray workbook.Sheets
                match sheetsRead |> Array.tryFindIndex isWorksheet with
                | Some target when sheetsRead.Length > 1 ->
                    let models = imported.Sheets |> Array.map (fun sheet -> sheet.Model)
                    let at : Rbt.Spreadsheet.Model.CellAddress = { Row = 2; Col = 1 }
                    let cell : Rbt.Spreadsheet.Model.CellData =
                        { Value = Rbt.Spreadsheet.Model.CellValue.String "edited here"; Comment = None
                          Format = { Rbt.Spreadsheet.Model.defaultFormat with Bold = true; BgColor = Some "#ABCDEF"; NumberFormat = Some "0.0000\"u\"" } }
                    models.[target] <- { models.[target] with Cells = Map.add at cell models.[target].Cells }
                    let editedBytes = XlsxPackageWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut { input with Sheets = models })
                    let edited = partsOf editedBytes
                    let others = sheetsRead |> Array.mapi (fun i sheet -> i, sheet) |> Array.filter (fun (i, sheet) -> i <> target && isWorksheet sheet)
                    if not workbook.Date1904 then
                        let differing =
                            others |> Array.toList
                            |> List.choose (fun (_, sheet) ->
                                let path = fst sheet.Extras.SourcePart.Value
                                match Map.tryFind path edited with
                                | Some data when data = original.[path] -> None
                                | _ -> Some path)
                        ok (name + ": other sheets identical after an edit") differing.IsEmpty (show differing)
                        // ... and everything those sheets reach is still there, as it was.
                        let lost =
                            wanted |> Set.toList
                            |> List.filter (fun part ->
                                not (part.StartsWith "xl/worksheets/sheet") && not (rewritten part && not (part.StartsWith "xl/comments" || part.StartsWith "xl/threadedComments/"))
                                && (others |> Array.exists (fun (_, sheet) -> reachableFrom original (fst sheet.Extras.SourcePart.Value) |> Set.contains part))
                                && (match Map.tryFind part edited with Some data -> data <> original.[part] | None -> true))
                        ok (name + ": what the other sheets reach is identical") lost.IsEmpty (show lost)
                    // The styles are the file's, added to: what the editor has no model of stays (table styles, the
                    // palette, extensions, protection, quote prefix, pivot buttons), and every format keeps its place.
                    (match Map.tryFind "xl/styles.xml" original, Map.tryFind "xl/styles.xml" edited with
                     | Some was, Some now ->
                        let was, now = text was, text now
                        let element (name: string) (xml: string) =
                            let found = Text.RegularExpressions.Regex.Match(xml, "<" + name + "(?=[\\s>/])[^>]*/>|<" + name + "(?=[\\s>])[^>]*>[\\s\\S]*?</" + name + ">")
                            if found.Success then found.Value else ""
                        let inner (name: string) (xml: string) = let e = element name xml in if e = "" || e.EndsWith "/>" then "" else e.Substring(e.IndexOf '>' + 1, e.LastIndexOf "</" - e.IndexOf '>' - 1)
                        let lost = [ "tableStyles"; "colors"; "extLst"; "cellStyles"; "cellStyleXfs" ] |> List.filter (fun name -> element name was <> element name now)
                        ok (name + ": styles keep what is not modelled") lost.IsEmpty (show lost)
                        let grown = [ "cellXfs"; "fonts"; "fills"; "borders"; "dxfs"; "numFmts" ] |> List.filter (fun name -> not ((inner name now).StartsWith (inner name was)))
                        ok (name + ": styles are added to, nothing renumbered") grown.IsEmpty (show grown)
                     | _ -> ())
                    // The workbook part keeps what is not the sheet list, the names or the calculation flags.
                    (let was, now = text original.["xl/workbook.xml"], text edited.["xl/workbook.xml"]
                     let tags (xml: string) = [ for m in Text.RegularExpressions.Regex.Matches(xml, "<(fileVersion|workbookPr|workbookProtection|fileSharing|workbookView)(?=[\\s>/])[^>]*>") -> Text.RegularExpressions.Regex.Replace(m.Value, "\\s(activeTab|firstSheet|date1904)=\"[^\"]*\"", "") ]
                     ok (name + ": workbook properties, version and window stay") (tags was = tags now) (sprintf "%A / %A" (tags was) (tags now))
                     let extLst (xml: string) = let at = xml.LastIndexOf "<extLst" in if at < 0 then "" else xml.Substring at
                     ok (name + ": workbook extensions stay") (extLst was = extLst now) "")
                    // The comments of the edited sheet were not touched: its comments part is the file's.
                    (let commentParts = reachableFrom original (fst sheetsRead.[target].Extras.SourcePart.Value) |> Set.filter (fun part -> part.StartsWith "xl/comments" && Map.containsKey part original)
                     let changedComments = commentParts |> Set.filter (fun part -> Map.tryFind part edited <> Some original.[part]) |> Set.toList
                     ok (name + ": the edited sheet's comments part is the file's") changedComments.IsEmpty (show changedComments))
                    // Read again: every cell of every sheet has the value and the format it had (but the edited one).
                    match (try XlsxReader.loadBytesNative editedBytes with _ -> None) with
                    | None -> ok (name + ": edited export opens") false ""
                    | Some again ->
                        let describe (book: XlsxReader.Workbook) (sheet: XlsxReader.Worksheet) =
                            [ for row in sheet.Rows do
                                for c in row.Cells do
                                    let value =
                                        match c.Value with
                                        | XlsxReader.SharedString index when index >= 0 && index < book.SharedStrings.Length -> XlsxReader.String book.SharedStrings.[index]
                                        | other -> other
                                    let format = XlsxWriter.formatOfReaderStyle c.Style
                                    if value <> XlsxReader.Empty || c.Formula.IsSome || format <> XlsxWriter.defaultXlsxCellFormat then
                                        yield (c.Reference.Row, c.Reference.Column), (value, c.Formula, format, (if workbook.Date1904 || c.StyleIndex = Some 0 then None else c.StyleIndex)) ]
                            |> Map.ofList
                        let wrong =
                            [ for i in 0 .. sheetsRead.Length - 1 do
                                if i < again.Sheets.Length then
                                    let was, now = describe workbook sheetsRead.[i], describe again again.Sheets.[i]
                                    let was, now = (if i = target then Map.remove (3, "B") was, Map.remove (3, "B") now else was, now)
                                    if was <> now then
                                        let first = was |> Map.toSeq |> Seq.tryFind (fun (key, cell) -> Map.tryFind key now <> Some cell) |> Option.map (fun ((row, col), _) -> sprintf "%s%d" col row) |> Option.defaultValue "+"
                                        yield sprintf "%s!%s" sheetsRead.[i].Name first ]
                        ok (name + ": cells mean what they meant after an edit") wrong.IsEmpty (show wrong)
                        let editedCell = again.Sheets.[target].Rows |> List.collect (fun row -> row.Cells) |> List.tryFind (fun c -> c.Reference.Row = 3 && c.Reference.Column = "B")
                        ok (name + ": the edited cell")
                            (match editedCell with
                             | Some c ->
                                (match c.Value with XlsxReader.SharedString index -> again.SharedStrings.[index] = "edited here" | XlsxReader.String text -> text = "edited here" | _ -> false)
                                && (match c.Style with Some style -> style.Bold && style.NumberFormat = Some "0.0000\"u\"" && style.FillColor |> Option.exists (fun color -> color.ToUpperInvariant().EndsWith "ABCDEF") | None -> false)
                             | None -> false) (sprintf "%A" (editedCell |> Option.map (fun c -> c.Value, c.Style)))
                | _ -> ()
            with error -> ok (name + ": export") false (error.Message.Split('\n').[0])

// ---- an edited sheet is the file's part with the edit in it ------------------
// A value typed into a cell, a fill on another, formulas with other results:
// the sheet's part is the file's with those <c> elements replaced, nothing
// else of it written anew - sheetPr, the views, ignoredErrors, the rows and
// their attributes, the shared formulas and the page setup are the file's.
if named.Length = 0 then
    let M = Rbt.Spreadsheet.Model.setCell
    let milk = File.ReadAllBytes(Path.Combine(root, "real-world/milk-street-office.xlsx"))
    let workbook = XlsxReader.loadBytesNative milk |> Option.get
    let imported = Bridge.Import.importWorkbook workbook
    let index = imported.Sheets |> Array.findIndex (fun sheet -> sheet.Name = "Annual")
    let sheetPath = (List.item index workbook.Sheets).Extras.SourcePart |> Option.get |> fst
    let before = partsOf milk
    let was = text before.[sheetPath]
    let addr row col : Rbt.Spreadsheet.Model.CellAddress = { Row = row; Col = col }
    let cellOfRef (reference: string) =
        let letters = reference |> Seq.takeWhile Char.IsLetter |> Seq.toArray |> String
        addr (int (reference.Substring letters.Length) - 1) ((letters |> Seq.fold (fun acc ch -> acc * 26 + int ch - int 'A' + 1) 0) - 1)
    // A cell that takes its formula from another (`<f t="shared" si/>`), holding a number.
    let sharedRef = Text.RegularExpressions.Regex.Match(was, "<c r=\"([A-Z]+[0-9]+)\"(?: s=\"[0-9]+\")?><f t=\"shared\" si=\"[0-9]+\"/><v>").Groups.[1].Value
    let exportWith (edit: Rbt.Spreadsheet.Model.SpreadsheetModel -> Rbt.Spreadsheet.Model.SpreadsheetModel) =
        let sheets = imported.Sheets |> Array.map (fun sheet -> sheet.Model)
        sheets.[index] <- edit sheets.[index]
        let input : Bridge.ExportInput =
            { Sheets = sheets; Names = imported.Sheets |> Array.map (fun sheet -> sheet.Name); Visible = imported.Sheets |> Array.map (fun sheet -> sheet.Visible)
              ActiveSheetIndex = workbook.Preserved.ActiveTab; NameInfo = imported.Context.NameInfo; DefinedNames = imported.Context.Names
              Source = Some imported.Source; Macros = None; MacroProject = None }
        XlsxPackageWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut input)
    let typed (model: Rbt.Spreadsheet.Model.SpreadsheetModel) =
        let cell at = Rbt.Spreadsheet.Model.getCell at model
        let b3 = cell (addr 2 1) |> Option.defaultValue { Value = Rbt.Spreadsheet.Model.CellValue.Empty; Format = Rbt.Spreadsheet.Model.defaultFormat; Comment = None }
        let shared = cell (cellOfRef sharedRef) |> Option.get
        model
        |> M (addr 397 13) { Value = Rbt.Spreadsheet.Model.CellValue.Number 5.0; Format = Rbt.Spreadsheet.Model.defaultFormat; Comment = None }
        |> M (addr 2 1) { b3 with Format = { b3.Format with BgColor = Some "#FFEE00" } }
        |> M (cellOfRef sharedRef) { shared with Value = (match shared.Value with Rbt.Spreadsheet.Model.CellValue.Formula (formula, _) -> Rbt.Spreadsheet.Model.CellValue.Formula (formula, Rbt.Spreadsheet.Model.CellValue.Number 123.25) | other -> other) }
    let saved = exportWith typed
    let after = partsOf saved
    let now = text after.[sheetPath]
    let element (name: string) (xml: string) =
        let found = Text.RegularExpressions.Regex.Match(xml, sprintf "<%s(?=[\\s>/])[^>]*/>|<%s(?=[\\s>])[^>]*>[\\s\\S]*?</%s>" name name name)
        if found.Success then found.Value else ""
    let count (pattern: string) (xml: string) = Text.RegularExpressions.Regex.Matches(xml, pattern).Count
    ok "edited sheet: the file has what the check is about"
        (sharedRef <> "" && (element "sheetPr" was).Contains "codeName" && element "ignoredErrors" was <> "" && count "t=\"shared\"" was > 100 && element "headerFooter" was = "")
        (sprintf "shared %s, sheetPr %s" sharedRef (element "sheetPr" was))
    for name in [ "sheetPr"; "sheetViews"; "sheetFormatPr"; "cols"; "ignoredErrors"; "pageMargins"; "pageSetup"; "headerFooter"; "mergeCells"; "extLst"; "drawing"; "legacyDrawing" ] do
        ok (sprintf "edited sheet: <%s> is the file's" name) (element name now = element name was) (sprintf "%s  <-  %s" (element name now |> fun t -> t.Substring(0, min 200 t.Length)) (element name was |> fun t -> t.Substring(0, min 200 t.Length)))
    let rowTags (xml: string) = [ for found in Text.RegularExpressions.Regex.Matches(xml, "<row\\b[^>]*>") -> Text.RegularExpressions.Regex.Replace(found.Value, " spans=\"[^\"]*\"", "") ]
    ok "edited sheet: the rows keep their attributes (style, height)" (rowTags now = rowTags was) (sprintf "%d rows, %d in the file" (rowTags now).Length (rowTags was).Length)
    ok "edited sheet: the shared formulas stay shared" (count "t=\"shared\"" now = count "t=\"shared\"" was) (sprintf "%d of %d" (count "t=\"shared\"" now) (count "t=\"shared\"" was))
    // But for the three cells the part is the file's, character for character.
    let without (xml: string) =
        [ "N398"; "B3"; sharedRef ]
        |> List.fold (fun (acc: string) reference -> Text.RegularExpressions.Regex.Replace(acc, sprintf "<c r=\"%s\"(?:[^>]*/>|[^>]*>[\\s\\S]*?</c>)" reference, "")) xml
        // (A row with a new cell loses its `spans`; the dimension holds the new cell.)
        |> fun stripped -> Text.RegularExpressions.Regex.Replace(stripped, " spans=\"[^\"]*\"|<dimension[^>]*>", "")
    ok "edited sheet: nothing but the edited cells differs from the file's part" (without now = without was) (sprintf "%d chars, the file's %d" (without now).Length (without was).Length)
    ok "edited sheet: the dimension holds the new cell" ((element "dimension" was).Contains "B1:L398" && (element "dimension" now).Contains "B1:N398") (element "dimension" now)
    ok "edited sheet: the typed value is in its cell" (Text.RegularExpressions.Regex.IsMatch(now, "<c r=\"N398\"[^>]*><v>5</v></c>")) "N398"
    ok "edited sheet: a formula with another result keeps its shared formula"
        (Text.RegularExpressions.Regex.IsMatch(now, sprintf "<c r=\"%s\"[^>]*><f t=\"shared\" si=\"[0-9]+\"/><v>123.25</v></c>" sharedRef)) sharedRef
    ok "edited sheet: the shared strings are the file's" (after.["xl/sharedStrings.xml"] = before.["xl/sharedStrings.xml"]) "sharedStrings.xml differs"
    let reread = Bridge.Import.importWorkbook (XlsxReader.loadBytesNative saved |> Option.get)
    let cellNow at = Rbt.Spreadsheet.Model.getCell at reread.Sheets.[index].Model
    ok "edited sheet: read back: the value, the fill and the formula's result"
        ((cellNow (addr 397 13) |> Option.map (fun cell -> cell.Value)) = Some (Rbt.Spreadsheet.Model.CellValue.Number 5.0)
         && (cellNow (addr 2 1) |> Option.bind (fun cell -> cell.Format.BgColor) |> Option.map (fun color -> color.ToUpperInvariant())) = Some "#FFEE00"
         && (match cellNow (cellOfRef sharedRef) |> Option.map (fun cell -> cell.Value) with Some (Rbt.Spreadsheet.Model.CellValue.Formula (_, Rbt.Spreadsheet.Model.CellValue.Number 123.25)) -> true | _ -> false))
        (sprintf "%A" (cellNow (addr 2 1) |> Option.map (fun cell -> cell.Format.BgColor)))
    // What holds the package together is the file's: the content types, the relationships, the workbook part.
    let plumbing = [ "[Content_Types].xml"; "_rels/.rels"; "xl/_rels/workbook.xml.rels"; "xl/workbook.xml" ]
    ok "edited sheet: content types, relationships and the workbook part are the file's" (plumbing |> List.forall (fun path -> after.[path] = before.[path]))
        (plumbing |> List.filter (fun path -> after.[path] <> before.[path]) |> String.concat ", ")
    // Only a value typed: every part but the sheet's is the file's, byte for byte (no format was added to the styles).
    let onlyTyped = partsOf (exportWith (M (addr 397 13) { Value = Rbt.Spreadsheet.Model.CellValue.Number 5.0; Format = Rbt.Spreadsheet.Model.defaultFormat; Comment = None }))
    let differing = before |> Map.toList |> List.filter (fun (path, data) -> Map.tryFind path onlyTyped <> Some data) |> List.map fst
    ok "a value typed: only the sheet's part is not the file's" (differing = [ sheetPath ] && onlyTyped.Count = before.Count) (String.concat ", " differing)
    // Another zoom and tab colour: the view keeps its selection and where it is scrolled to.
    let zoomed = text (partsOf (exportWith (fun model -> { model with Zoom = 0.75; TabColor = Some "#00B050" }))).[sheetPath]
    let viewNow, viewWas = element "sheetViews" zoomed, element "sheetViews" was
    ok "zoomed sheet: the view is the file's with the new zoom"
        (viewNow = viewWas.Replace(" tabSelected=\"1\"", " tabSelected=\"1\" zoomScale=\"75\"").Replace(" zoomScaleNormal=\"100\" workbookViewId", " zoomScaleNormal=\"100\" workbookViewId") || (viewNow.Contains "zoomScale=\"75\"" && viewNow.Contains "topLeftCell=\"A308\"" && viewNow.Contains "<selection activeCell=\"B308\"" && viewNow.Length = viewWas.Length + " zoomScale=\"75\"".Length))
        viewNow
    ok "zoomed sheet: the tab colour is in the file's sheetPr" (element "sheetPr" zoomed = (element "sheetPr" was).Replace("<pageSetUpPr", "<tabColor rgb=\"FF00B050\"/><pageSetUpPr")) (element "sheetPr" zoomed)
    let strip (xml: string) = xml.Replace(element "sheetViews" xml, "").Replace(element "sheetPr" xml, "")
    ok "zoomed sheet: the rest of the part is the file's" (strip zoomed = strip was) "differs"
    // A row inserted: the sheet is written from its model. The results of
    // formulas that are texts stay in their cells - none becomes a shared string.
    let restructured = partsOf (exportWith (Rbt.Spreadsheet.Model.insertRow 500))
    let strings (parts: Map<string, byte[]>) = count "<si[\\s>]" (text parts.["xl/sharedStrings.xml"])
    ok "row inserted: no result of a formula is added to the shared strings" (strings restructured = strings before) (sprintf "%d strings, the file has %d" (strings restructured) (strings before))
    let restructuredSheet = text restructured.[sheetPath]
    ok "row inserted: nothing the file did not say is added (an empty header and footer, fit-to-page counts)"
        (not (restructuredSheet.Contains "<headerFooter") && not (restructuredSheet.Contains "fitToWidth") && not (restructuredSheet.Contains "fitToHeight") && not (was.Contains "fitToWidth"))
        (element "pageSetup" restructuredSheet + element "headerFooter" restructuredSheet)
    ok "row inserted: text results are written in their cells" (count "t=\"str\"[^>]*><f" (text restructured.[sheetPath]) = count "t=\"str\"[^>]*><f" was) (sprintf "%d of %d" (count "t=\"str\"[^>]*><f" (text restructured.[sheetPath])) (count "t=\"str\"[^>]*><f" was))

printfn "\nXLSX part carry-over tests: %d passed, %d failed (%d workbooks, %d files skipped)" pass fail books skipped
if fail > 0 then exit 1
