#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

// How a sheet's part is written after one edit, for every workbook under
// tests/real-world and tests/test-documents-excel and per kind of edit:
//
//   patched     the file's part with what the edit is about changed in it: every other element
//               of the worksheet is the file's text (cols with bestFit, pageSetup with pageOrder,
//               the views, the extension list ...), every row keeps its start tag (spans, styles)
//   rows kept   the rows are the file's, the rest of the part is made from the sheet
//   from model  the part is written from the sheet's cells
//
// The table is printed; what is asserted is that the edits a part can take in its text - a cell,
// a link added or removed, a note added, edited or removed (also the last one) - are patched in
// every workbook whose part is laid out as the writer expects, and that every saved file opens
// with the sheet as it was edited.
//
// Usage: dotnet fsi tests/office-formats/xlsx-patched-corpus.fsx [--list] [file.xlsx ...]
// (--list names the workbooks behind every count that is not "patched".)

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
let element (name: string) (xml: string) =
    let found = Regex.Match(xml, sprintf "<%s(?=[\\s>/])[^>]*/>|<%s(?=[\\s>])[^>]*>[\\s\\S]*?</%s>" name name name)
    if found.Success then found.Value else ""
let addr row col : M.CellAddress = { Row = row; Col = col }
let childOrder =
    [ "sheetPr"; "dimension"; "sheetViews"; "sheetFormatPr"; "cols"; "sheetCalcPr"; "sheetProtection"; "protectedRanges"; "scenarios"
      "autoFilter"; "sortState"; "dataConsolidate"; "customSheetViews"; "mergeCells"; "phoneticPr"; "conditionalFormatting"; "dataValidations"
      "hyperlinks"; "printOptions"; "pageMargins"; "pageSetup"; "headerFooter"; "rowBreaks"; "colBreaks"; "customProperties"; "cellWatches"
      "ignoredErrors"; "smartTags"; "drawing"; "legacyDrawing"; "legacyDrawingHF"; "drawingHF"; "picture"; "oleObjects"; "controls"
      "webPublishItems"; "tableParts"; "extLst" ]
let rowTags (xml: string) = [ for found in Regex.Matches(xml, "<row\\b[^>]*>") -> found.Value ]
let rowNumber (tag: string) = Regex.Match(tag, "\\sr=\"([0-9]+)\"").Groups.[1].Value

/// How the part was written: see the head of the file. `own`: the elements the edit is about;
/// `rows`: the rows (as the part numbers them) the edit put a cell into.
let written (was: string) (now: string) (own: string list) (rows: string list) : string =
    let rootOf (xml: string) = Regex.Match(xml, "<worksheet\\b[^>]*>").Value
    let outsideRows (xml: string) = Regex.Replace(xml, "<sheetData\\b[\\s\\S]*</sheetData>|<sheetData\\b[^>]*/>", "")
    let wasOut, nowOut = outsideRows was, outsideRows now
    let others = childOrder |> List.filter (fun name -> name <> "dimension" && not (List.contains name own))
    let same = rootOf was = rootOf now && others |> List.forall (fun name -> element name nowOut = element name wasOut)
    let tagsOf (xml: string) = rowTags xml |> List.filter (fun tag -> not (List.contains (rowNumber tag) rows))
    let rowsSame = tagsOf now = tagsOf was
    if same && rowsSame then "patched" elif rowsSame then "rows kept" else "from model"

// The edits: a name, the elements of the part each is about, and the edit (None: the sheet has nothing to make it on).
let far = addr 2000 30
let edits : (string * string list * string list * (M.SpreadsheetModel -> M.SpreadsheetModel option)) list =
    [ "a cell typed", [], [ "2001" ], (fun model -> Some (M.setCell far { Value = M.CellValue.Number 12.5; Format = M.defaultFormat; Comment = None } model))
      "a link added", [ "hyperlinks" ], [ "2001" ], (fun model -> Some (Ops.setHyperlink [ { From = far; To = far } ] "https://example.com/a?b=1&c=2" false (Some "link") None model))
      "a link to a cell added", [ "hyperlinks" ], [ "2001" ], (fun model -> Some (Ops.setHyperlink [ { From = far; To = far } ] "A1" true None None model))
      "one of several links removed", [ "hyperlinks" ], [],
        (fun model -> if model.Hyperlinks.Length < 2 then None else Some (fst (Ops.removeHyperlinks [ model.Hyperlinks.Head.Range ] model)))
      "every link removed", [ "hyperlinks" ], [],
        (fun model -> if model.Hyperlinks.IsEmpty then None else Some (fst (Ops.removeHyperlinks (model.Hyperlinks |> List.map (fun link -> link.Range)) model)))
      "a note added", [ "legacyDrawing" ], [], (fun model -> Some (Ops.setComment far "a note" (Some "me") model))
      "a note edited", [ "legacyDrawing" ], [],
        (fun model -> match Ops.allComments model |> List.filter (fun note -> note.Thread.IsNone) with first :: _ -> Some (Ops.setComment first.Cell "said again" (Some "me") model) | [] -> None)
      "one of several notes removed", [ "legacyDrawing" ], [],
        (fun model -> match Ops.allComments model |> List.filter (fun note -> note.Thread.IsNone) with first :: _ :: _ -> Some (Ops.deleteComment first.Cell model) | _ -> None)
      "the last note removed", [ "legacyDrawing" ], [],
        (fun model ->
            match Ops.allComments model with
            | [] -> None
            | notes when notes |> List.exists (fun note -> note.Thread.IsSome) -> None
            | notes -> Some (notes |> List.fold (fun sheet note -> Ops.deleteComment note.Cell sheet) model))
      "a table added", [ "tableParts" ], [ "2001"; "2002"; "2003" ], (fun model -> Some (M.createTable "PatchedCorpusTable" 2000 30 2002 31 false model))
      "a row inserted", [], [], (fun model -> Some (M.insertRow 1 model)) ]
/// The edits a part takes in its text: asserted to be patched wherever the part is laid out as the writer expects.
let patchable = set [ "a cell typed"; "a link added"; "a link to a cell added"; "one of several links removed"; "every link removed"; "a note added"; "a note edited"; "one of several notes removed"; "the last note removed" ]

let counts = Collections.Generic.Dictionary<string * string, ResizeArray<string>>()
let note (edit: string) (how: string) (book: string) =
    match counts.TryGetValue ((edit, how)) with
    | true, books -> books.Add book
    | false, _ -> counts.[(edit, how)] <- ResizeArray [ book ]
let mutable books = 0
for file in files do
    let name = Path.GetFileName file
    let bytes = File.ReadAllBytes file
    match (try XlsxReader.loadBytesNative bytes with _ -> None) with
    | None -> ()
    | Some book ->
        match (try Some (Bridge.Import.importWorkbook book) with _ -> None) with
        | None -> ()
        | Some opened ->
            let target =
                book.Sheets |> List.tryFindIndex (fun sheet ->
                    match sheet.Extras.SourcePart with
                    | Some (_, relType) -> relType.EndsWith "/worksheet"
                    | None -> false)
            match target with
            | Some target when target < opened.Sheets.Length ->
                books <- books + 1
                let path = (List.item target book.Sheets).Extras.SourcePart |> Option.get |> fst
                let was = text (partsOf bytes).[path]
                let start = M.resolveSheet opened.Sheets.[target].Model
                let danglingBefore =
                    let file0 = partsOf bytes
                    set [ for KeyValue(part, data) in file0 do
                            if part.EndsWith ".rels" then
                                let folder = part.Substring(0, part.IndexOf "_rels/").TrimEnd('/')
                                for found in Regex.Matches(text data, "<Relationship\\b[^>]*>") do
                                    let target = Regex.Match(found.Value, "\\sTarget=\"([^\"]*)\"").Groups.[1].Value
                                    if not (found.Value.Contains "TargetMode=\"External\"") && target <> "" && not (target.StartsWith "#") then
                                        let stack = Collections.Generic.List<string>((if target.StartsWith "/" then "" else folder).Split([| '/' |], StringSplitOptions.RemoveEmptyEntries))
                                        for piece in (Uri.UnescapeDataString target).Split('/') do
                                            if piece = ".." then (if stack.Count > 0 then stack.RemoveAt(stack.Count - 1))
                                            elif piece <> "." && piece <> "" then stack.Add piece
                                        let resolved = String.Join("/", stack)
                                        if not (file0.ContainsKey resolved) then yield part + " -> " + resolved ]
                // (A part the writer does not patch at all - a prefix on its elements, rows without `r` -: a cell typed says so.)
                let mutable laidOut = true
                for label, own, rows, edit in edits do
                    match (try edit start with _ -> None) with
                    | None -> ()
                    | Some edited ->
                        try
                            let sheets = opened.Sheets |> Array.map (fun sheet -> sheet.Model)
                            sheets.[target] <- edited
                            let input : Bridge.ExportInput =
                                { Sheets = sheets; Names = opened.Sheets |> Array.map (fun sheet -> sheet.Name); Visible = opened.Sheets |> Array.map (fun sheet -> sheet.Visible)
                                  ActiveSheetIndex = book.Preserved.ActiveTab; NameInfo = opened.Context.NameInfo; DefinedNames = opened.Context.Names
                                  Source = Some opened.Source; Macros = None }
                            let saved = XlsxWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut input)
                            let now = text (partsOf saved).[path]
                            let how = written was now own rows
                            if label = "a cell typed" && how <> "patched" then laidOut <- false
                            note label how name
                            if patchable.Contains label && laidOut then
                                ok (name + ": " + label + ": the part is the file's with that changed in it") (how = "patched") how
                            // The file opens, and the sheet says what was edited.
                            let again = M.resolveSheet (Bridge.Import.importWorkbook (XlsxReader.loadBytesNative saved |> Option.get)).Sheets.[target].Model
                            let links (model: M.SpreadsheetModel) = model.Hyperlinks |> List.map (fun link -> link.Range, link.Target, link.IsInternal, link.Display, link.Tooltip) |> List.sort
                            let notes (model: M.SpreadsheetModel) = Ops.allComments model |> List.map (fun n -> n.Cell.Row, n.Cell.Col, n.Text) |> List.sort
                            ok (name + ": " + label + ": the links read back") (links again = links edited) (sprintf "%A, wanted %A" (links again |> List.truncate 3) (links edited |> List.truncate 3))
                            ok (name + ": " + label + ": the notes read back") (notes again = notes edited) (sprintf "%A, wanted %A" (notes again |> List.truncate 3) (notes edited |> List.truncate 3))
                            // Every relationship of the sheet names a part that is there, once.
                            let now = partsOf saved
                            // The package says what it holds: every part the content types name is there, every
                            // part has a type, and what every relationships part leads to is in the package.
                            let types = text now.["[Content_Types].xml"]
                            let overrides = [ for found in Regex.Matches(types, "<Override\\b[^>]*PartName=\"/([^\"]*)\"") -> Uri.UnescapeDataString found.Groups.[1].Value ]
                            let defaults = set [ for found in Regex.Matches(types, "<Default\\b[^>]*Extension=\"([^\"]*)\"") -> found.Groups.[1].Value.ToLowerInvariant() ]
                            ok (name + ": " + label + ": every part the content types name is in the package") (overrides |> List.forall now.ContainsKey) (overrides |> List.filter (now.ContainsKey >> not) |> String.concat " ")
                            let untyped = now |> Map.toList |> List.map fst |> List.filter (fun part -> part <> "[Content_Types].xml" && not (List.contains part overrides) && not (defaults.Contains (Path.GetExtension(part).TrimStart('.').ToLowerInvariant())))
                            ok (name + ": " + label + ": every part has a content type") untyped.IsEmpty (String.concat " " untyped)
                            let dangling =
                                [ for KeyValue(part, data) in now do
                                    if part.EndsWith ".rels" then
                                        let folder = part.Substring(0, part.IndexOf "_rels/").TrimEnd('/')
                                        for found in Regex.Matches(text data, "<Relationship\\b[^>]*>") do
                                            let target = Regex.Match(found.Value, "\\sTarget=\"([^\"]*)\"").Groups.[1].Value
                                            if not (found.Value.Contains "TargetMode=\"External\"") && target <> "" && not (target.StartsWith "#") then
                                                let stack = Collections.Generic.List<string>((if target.StartsWith "/" then "" else folder).Split([| '/' |], StringSplitOptions.RemoveEmptyEntries))
                                                for piece in (Uri.UnescapeDataString target).Split('/') do
                                                    if piece = ".." then (if stack.Count > 0 then stack.RemoveAt(stack.Count - 1))
                                                    elif piece <> "." && piece <> "" then stack.Add piece
                                                let resolved = String.Join("/", stack)
                                                if not (now.ContainsKey resolved) then yield part + " -> " + resolved ]
                            // (What the file itself left dangling is not the save's.)
                            ok (name + ": " + label + ": no relationship leads to a part that is not there") (dangling |> List.forall (fun entry -> danglingBefore.Contains entry)) (String.concat "; " (dangling |> List.truncate 4))
                            let relsPath = path.Substring(0, path.LastIndexOf '/') + "/_rels/" + path.Substring(path.LastIndexOf '/' + 1) + ".rels"
                            match now.TryFind relsPath with
                            | Some rels ->
                                let entries = [ for found in Regex.Matches(text rels, "<Relationship\\b[^>]*/>") -> found.Value ]
                                let ids = entries |> List.map (fun entry -> Regex.Match(entry, "\\sId=\"([^\"]*)\"").Groups.[1].Value)
                                ok (name + ": " + label + ": the sheet's relationship ids are unique") (List.distinct ids = ids) (String.concat " " ids)
                                let used = [ for found in Regex.Matches(text now.[path], "\\br:id=\"([^\"]*)\"") -> found.Groups.[1].Value ]
                                ok (name + ": " + label + ": every relationship the part names is there") (used |> List.forall (fun id -> List.contains id ids)) (sprintf "names %A, has %A" used ids)
                            | None ->
                                ok (name + ": " + label + ": a part without relationships names none") (not (Regex.IsMatch(text now.[path], "\\br:id=\""))) "r:id without a relationships part"
                        with error -> ok (name + ": " + label + ": saved") false (error.Message.Split('\n').[0])
            | _ -> ()

printfn "\nHow the edited sheet's part was written, by kind of edit (%d workbooks, the first worksheet of each):" books
printfn "  %-32s %8s %10s %11s" "" "patched" "rows kept" "from model"
for label, _, _, _ in edits do
    let countOf (how: string) = match counts.TryGetValue ((label, how)) with | true, found -> found.Count | false, _ -> 0
    printfn "  %-32s %8d %10d %11d" label (countOf "patched") (countOf "rows kept") (countOf "from model")
    if listAll then
        for how in [ "rows kept"; "from model" ] do
            match counts.TryGetValue ((label, how)) with
            | true, found -> printfn "      %s: %s" how (String.concat ", " found)
            | false, _ -> ()

printfn "\nXLSX patched-corpus tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
