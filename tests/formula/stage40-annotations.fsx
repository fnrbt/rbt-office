#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
module Ops = Rbt.Spreadsheet.SheetOps
module Formula = Rbt.Spreadsheet.Formula

// What is attached to cells besides their content: comments (notes and
// threads with replies), hyperlinks and pictures - the operations behind
// the host API, and how they follow rows and columns that move.

let mutable pass = 0
let mutable fail = 0
let eq label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-44s expected %A  got %A" label expected got)
let throws label (run: unit -> unit) (part: string) =
    try
        run ()
        fail <- fail + 1
        printfn "FAIL %-44s did not throw" label
    with error ->
        if error.Message.Contains part then pass <- pass + 1
        else (fail <- fail + 1; printfn "FAIL %-44s threw %s" label error.Message)
let isError label (result: Result<'a, string>) (part: string) =
    match result with
    | Result.Error message when message.Contains part -> pass <- pass + 1
    | Result.Error message -> fail <- fail + 1; printfn "FAIL %-44s error was %s" label message
    | Ok _ -> fail <- fail + 1; printfn "FAIL %-44s no error" label
let okOf (result: Result<'a, string>) = match result with Ok value -> value | Result.Error message -> failwith message

Formula.formulaHooksInstalled |> ignore

let at row col : CellAddress = { Row = row; Col = col }
let cell (value: CellValue) = { Value = value; Format = defaultFormat; Comment = None }
let sheetOf cells = cells |> List.fold (fun m (addr, data) -> setCell addr data m) (createEmpty 60 10)
let range a b c d : CellRange = { From = at a b; To = at c d }

// ---- comments -------------------------------------------------------------------
let empty = sheetOf [ at 0 0, cell (CellValue.Number 1.0) ]
let noted = Ops.setComment (at 1 1) "check this" (Some "Ann") empty
eq "note: on a cell that did not exist" (Ops.commentAt noted (at 1 1) |> Option.map (fun c -> c.Text, c.Author, c.Thread.IsSome)) (Some ("check this", Some "Ann", false))
eq "note: the cell shows the marker" (getCell (at 1 1) noted |> Option.map (fun data -> data.Format.HasComment)) (Some true)
eq "note: no comment elsewhere" (Ops.commentAt noted (at 0 0)) None
let renoted = Ops.setComment (at 1 1) "checked" None noted
// (No author given: the note keeps the one it has.)
eq "note: replaced" (Ops.commentAt renoted (at 1 1) |> Option.map (fun c -> c.Text, c.Author)) (Some ("checked", Some "Ann"))
let onValue = Ops.setComment (at 0 0) "one" None empty
eq "note: the value stays" (getCell (at 0 0) onValue |> Option.map (fun data -> data.Value, data.Comment |> Option.map (fun c -> c.Text))) (Some (CellValue.Number 1.0, Some "one"))
let removed = Ops.deleteComment (at 1 1) noted
eq "delete: the note goes, and the cell that was only there for it" (Ops.commentAt removed (at 1 1), getCell (at 1 1) removed) (None, None)
eq "delete: a cell with a value keeps it, without the marker" (Ops.deleteComment (at 0 0) onValue |> getCell (at 0 0) |> Option.map (fun data -> data.Value, data.Comment, data.Format.HasComment)) (Some (CellValue.Number 1.0, None, false))
eq "delete: nothing to delete gives the same model" (obj.ReferenceEquals(Ops.deleteComment (at 5 5) noted, noted)) true
isError "reply: no comment" (Ops.addCommentReply (at 4 4) "hi" None None noted) "has no comment"
// A reply turns a note into a thread.
let replied, replyId = Ops.addCommentReply (at 1 1) "done" (Some "Bob") (Some "2026-01-02T03:04:05.00") noted |> okOf
let thread = (Ops.commentAt replied (at 1 1)).Value.Thread.Value
eq "reply: the note is the first comment" (thread.Text, thread.Author, thread.Replies |> List.map (fun r -> r.Id, r.Text, r.Author, r.Date)) ("check this", Some "Ann", [ replyId, "done", Some "Bob", Some "2026-01-02T03:04:05.00" ])
eq "reply: the note itself is gone, the marker stays" (getCell (at 1 1) replied |> Option.map (fun data -> data.Comment, data.Format.HasComment)) (Some (None, true))
eq "reply: ids are braces GUIDs and differ" (thread.Id.StartsWith "{" && thread.Id.Length = 38 && thread.Id <> replyId) true
let replied2, replyId2 = Ops.addCommentReply (at 1 1) "thanks" None None replied |> okOf
eq "reply: a second one follows" ((Ops.commentAt replied2 (at 1 1)).Value.Thread.Value.Replies |> List.map (fun r -> r.Text), replyId2 <> replyId) ([ "done"; "thanks" ], true)
// SetComment on a thread edits its first comment.
let edited = Ops.setComment (at 1 1) "check that" None replied2
eq "set on a thread: first comment edited, replies kept" ((Ops.commentAt edited (at 1 1)).Value |> fun c -> c.Text, c.Author, c.Thread.Value.Replies.Length) ("check that", Some "Ann", 2)
eq "set on a thread: a new author drops the person" ((Ops.setComment (at 1 1) "x" (Some "Cy") { replied2 with Threads = replied2.Threads |> Map.map (fun _ t -> { t with PersonId = Some "p1" }) }).Threads.[at 1 1] |> fun t -> t.Author, t.PersonId) (Some "Cy", None)
let unreplied = Ops.deleteCommentReply (at 1 1) replyId replied2 |> okOf
eq "delete reply: by id" (unreplied.Threads.[at 1 1].Replies |> List.map (fun r -> r.Text)) [ "thanks" ]
isError "delete reply: unknown id" (Ops.deleteCommentReply (at 1 1) "{nope}" replied2) "no reply with the id"
isError "delete reply: no thread" (Ops.deleteCommentReply (at 0 0) replyId replied2) "no comment thread"
let resolved = Ops.setCommentResolved (at 1 1) true replied2 |> okOf
eq "resolve: marked" (resolved.Threads.[at 1 1].Done, (Ops.setCommentResolved (at 1 1) false resolved |> okOf).Threads.[at 1 1].Done) (true, false)
isError "resolve: a note cannot be" (Ops.setCommentResolved (at 1 1) true noted) "cannot be resolved"
eq "delete: a thread goes with its replies" (Ops.deleteComment (at 1 1) replied2 |> fun m -> m.Threads.IsEmpty, getCell (at 1 1) m) (true, None)
eq "all comments, row by row" (Ops.allComments (Ops.setComment (at 0 3) "first" None replied2) |> List.map (fun c -> c.Cell, c.Text, c.Thread.IsSome)) [ at 0 3, "first", false; at 1 1, "check this", true ]
eq "placeholder note of a thread" (Ops.threadPlaceholderNote thread |> fun note -> note.Author = Some ("tc=" + thread.Id), note.Text.Contains "Comment:\n    check this", note.Text.Contains "Reply:\n    done") (true, true, true)
eq "clear all takes notes and threads, clear contents neither" (Ops.clearCells Ops.ClearEverything (0, 0, 9, 9) (Ops.setComment (at 0 3) "first" None replied2) |> Ops.allComments, Ops.clearCells Ops.ClearContents (0, 0, 9, 9) (Ops.setComment (at 0 3) "first" None replied2) |> Ops.allComments |> List.length) ([], 2)
// Comments follow rows and columns.
let shifted = applyStructureEdit (InsertRows (0, 2)) (Ops.setComment (at 0 3) "first" None replied2)
eq "insert rows: notes and threads move" (Ops.allComments shifted |> List.map (fun c -> c.Cell)) [ at 2 3; at 3 1 ]
eq "delete the row: the comments go" (applyStructureEdit (DeleteRows (1, 1)) replied2 |> Ops.allComments) []

// ---- hyperlinks -----------------------------------------------------------------
let linkOf target isInternal : SheetHyperlink = { Range = range 0 0 0 0; Target = target; IsInternal = isInternal; Tooltip = None; Display = None }
eq "scheme" ([ "https://example.com"; "MAILTO:a@b.c"; "tel:+1"; "javascript:alert(1)"; "Sheet2!A1"; "no scheme"; "C:\\file.txt"; "" ] |> List.map Ops.hyperlinkScheme)
    [ "https"; "mailto"; "tel"; "javascript"; ""; ""; "c"; "" ]
eq "safe: http, https, mailto, tel and nothing else" ([ "http://a"; "HTTPS://a"; "mailto:a@b.c"; "tel:1"; "javascript:alert(1)"; "data:text/html,x"; "file:///etc/passwd"; "ftp://a"; "vbscript:x"; "example.com" ] |> List.map (fun t -> Ops.hyperlinkIsSafe (linkOf t false)))
    [ true; true; true; true; false; false; false; false; false; false ]
eq "safe: a place in the workbook is not an address" (Ops.hyperlinkIsSafe (linkOf "Sheet2!A1" true)) false
let classify = Ops.classifyHyperlinkTarget (fun name -> name = "Totals")
eq "classify" ([ "https://example.com/x#y"; "#Sheet2!A1"; "Sheet2!A1"; "'My sheet'!B2:C3"; "A1"; "Totals"; "#Totals"; "mailto:a@b.c"; "example.com/page"; "javascript:alert(1)"; " tel:5 " ] |> List.map classify)
    [ "https://example.com/x#y", false; "Sheet2!A1", true; "Sheet2!A1", true; "'My sheet'!B2:C3", true; "A1", true; "Totals", true; "Totals", true; "mailto:a@b.c", false; "example.com/page", false; "javascript:alert(1)", false; "tel:5", false ]
let linkBase = sheetOf [ at 0 0, cell (CellValue.String "site"); at 2 0, cell (CellValue.Number 5.0); at 3 0, { cell (CellValue.String "red") with Format = { defaultFormat with TextColor = Some "#FF0000" } } ]
let linkedOnce = Ops.setHyperlink [ range 0 0 0 0 ] "https://example.com" false None (Some "the site") linkBase
eq "set: the link" (linkedOnce.Hyperlinks |> List.map (fun l -> l.Range, l.Target, l.IsInternal, l.Tooltip, l.Display)) [ range 0 0 0 0, "https://example.com", false, Some "the site", None ]
eq "set: the text stays, the cell looks like a link" (getCell (at 0 0) linkedOnce |> Option.map (fun d -> d.Value, d.Format.TextColor, d.Format.Underline)) (Some (CellValue.String "site", Some "#0563C1", true))
let linkedEmpty = Ops.setHyperlink [ range 1 1 1 1 ] "https://example.com/a" false None None linkBase
eq "set: an empty cell shows the target" (getCell (at 1 1) linkedEmpty |> Option.map (fun d -> d.Value)) (Some (CellValue.String "https://example.com/a"))
let linkedShown = Ops.setHyperlink [ range 2 0 2 0 ] "Sheet2!A1" true (Some "go") None linkBase
eq "set: display replaces the text" (getCell (at 2 0) linkedShown |> Option.map (fun d -> d.Value), linkedShown.Hyperlinks.[0].Display, linkedShown.Hyperlinks.[0].IsInternal) (Some (CellValue.String "go"), Some "go", true)
eq "set: a cell with a colour of its own keeps it" (Ops.setHyperlink [ range 3 0 3 0 ] "https://x.y" false None None linkBase |> getCell (at 3 0) |> Option.map (fun d -> d.Format.TextColor, d.Format.Underline)) (Some (Some "#FF0000", false))
let linkedBlock = Ops.setHyperlink [ range 0 0 3 1 ] "https://example.com/block" false None None linkedOnce
eq "set over a link: the new one replaces it" (linkedBlock.Hyperlinks |> List.map (fun l -> l.Range, l.Target)) [ range 0 0 3 1, "https://example.com/block" ]
eq "set on a block: only cells that hold something are styled, none is made but the first" (linkedBlock.Cells |> Map.toList |> List.map (fun (a, d) -> a, d.Format.Underline)) [ at 0 0, true; at 2 0, true; at 3 0, false ]
eq "lookup: the block's cells have the link, others none" (Ops.hyperlinkAt linkedBlock (at 2 1) |> Option.map (fun l -> l.Target), Ops.hyperlinkAt linkedBlock (at 4 0)) (Some "https://example.com/block", None)
let overlaid = Ops.setHyperlink [ range 1 0 1 0 ] "Sheet2!B2" true None None linkedBlock
eq "set inside a block: the block is cut around the cell" (overlaid.Hyperlinks |> List.map (fun l -> l.Range, l.Target) |> List.sort)
    ([ range 0 0 0 1, "https://example.com/block"; range 1 1 1 1, "https://example.com/block"; range 2 0 3 1, "https://example.com/block"; range 1 0 1 0, "Sheet2!B2" ] |> List.sort)
eq "lookup: the cell's own link" (Ops.hyperlinkAt overlaid (at 1 0) |> Option.map (fun l -> l.Target), Ops.hyperlinkAt overlaid (at 1 1) |> Option.map (fun l -> l.Target)) (Some "Sheet2!B2", Some "https://example.com/block")
let unlinked, removedCount = Ops.removeHyperlinks [ range 0 0 0 0 ] linkedOnce
eq "remove: the link and its look go, the text stays" (removedCount, unlinked.Hyperlinks, getCell (at 0 0) unlinked |> Option.map (fun d -> d.Value, d.Format.TextColor, d.Format.Underline)) (1, [], Some (CellValue.String "site", None, false))
let partly, partlyCount = Ops.removeHyperlinks [ range 0 0 1 1 ] linkedBlock
eq "remove part of a block: the rest keeps the link" (partlyCount, partly.Hyperlinks |> List.map (fun l -> l.Range), getCell (at 2 0) partly |> Option.map (fun d -> d.Format.Underline), getCell (at 0 0) partly |> Option.map (fun d -> d.Format.Underline)) (1, [ range 2 0 3 1 ], Some true, Some false)
eq "remove where there is none: the same model" (Ops.removeHyperlinks [ range 8 8 9 9 ] linkedOnce |> fun (m, n) -> obj.ReferenceEquals(m, linkedOnce), n) (true, 0)
eq "remove: the target shown in an empty cell stays as text" (Ops.removeHyperlinks [ range 1 1 1 1 ] linkedEmpty |> fst |> getCell (at 1 1) |> Option.map (fun d -> d.Value, d.Format.Underline)) (Some (CellValue.String "https://example.com/a", false))
eq "a whole column: no cell is made" (Ops.setHyperlink [ { From = at 0 5; To = at Ops.maxSheetRow 5 } ] "https://example.com" false None None (createEmpty 60 10) |> fun m -> m.Cells.Count, m.Hyperlinks.Length) (1, 1)
eq "links follow rows" ((applyStructureEdit (InsertRows (0, 3)) overlaid).Hyperlinks |> List.map (fun l -> l.Range.From.Row) |> List.sort) [ 3; 4; 4; 5 ]
// (Many single-cell links are looked up by cell.)
let manyLinks = { createEmpty 3000 10 with Hyperlinks = [ for row in 0 .. 1999 -> { Range = range row 1 row 1; Target = sprintf "https://example.com/%d" row; IsInternal = false; Tooltip = None; Display = None } ] }
eq "lookup among many" ([ 0; 1234; 1999; 2000 ] |> List.map (fun row -> Ops.hyperlinkAt manyLinks (at row 1) |> Option.map (fun l -> l.Target))) [ Some "https://example.com/0"; Some "https://example.com/1234"; Some "https://example.com/1999"; None ]

// ---- pictures -------------------------------------------------------------------
let png (width: int) (height: int) : byte[] =
    Array.concat [ [| 0x89uy; 0x50uy; 0x4Euy; 0x47uy; 0x0Duy; 0x0Auy; 0x1Auy; 0x0Auy; 0uy; 0uy; 0uy; 13uy; 0x49uy; 0x48uy; 0x44uy; 0x52uy |]
                   [| byte (width >>> 24); byte (width >>> 16); byte (width >>> 8); byte width; byte (height >>> 24); byte (height >>> 16); byte (height >>> 8); byte height |]
                   [| 8uy; 6uy; 0uy; 0uy; 0uy; 0uy; 0uy; 0uy; 0uy |] ]
let gif (width: int) (height: int) : byte[] =
    Array.concat [ "GIF89a"B; [| byte width; byte (width >>> 8); byte height; byte (height >>> 8); 0x80uy; 0uy; 0uy |] ]
// A JPEG with an APP0 segment, a comment and a quantisation table before its (progressive) frame.
let jpeg (width: int) (height: int) : byte[] =
    Array.concat [ [| 0xFFuy; 0xD8uy; 0xFFuy; 0xE0uy; 0uy; 16uy |]; Array.zeroCreate 14
                   [| 0xFFuy; 0xFEuy; 0uy; 5uy; 1uy; 2uy; 3uy |]
                   [| 0xFFuy; 0xDBuy; 0uy; 4uy; 0uy; 0uy |]
                   [| 0xFFuy; 0xC2uy; 0uy; 17uy; 8uy; byte (height >>> 8); byte height; byte (width >>> 8); byte width; 3uy |]; Array.zeroCreate 9 ]
let infoOf (data: byte[]) = Ops.imageInfo data |> Option.map (fun info -> info.ContentType, info.Extension, info.Width, info.Height)
eq "image: PNG" (infoOf (png 640 480)) (Some ("image/png", "png", 640, 480))
eq "image: a large PNG" (infoOf (png 70000 3)) (Some ("image/png", "png", 70000, 3))
eq "image: GIF" (infoOf (gif 300 2)) (Some ("image/gif", "gif", 300, 2))
eq "image: JPEG, the size is in the frame segment" (infoOf (jpeg 1024 768)) (Some ("image/jpeg", "jpeg", 1024, 768))
eq "image: not an image" ([ "<svg xmlns='http://www.w3.org/2000/svg'/>"B; "BM......................"B; [||]; [| 0xFFuy; 0xD8uy |]; (png 5 5).[0..15]; Array.append (jpeg 4 4).[0..21] [| 0xFFuy |] ] |> List.map infoOf) [ None; None; None; None; None; None ]
eq "image: a PNG without a size" (infoOf (png 0 9)) None
let noPictures = sheetOf [ at 0 0, cell (CellValue.Number 1.0) ]
let pngInfo = (Ops.imageInfo (png 40 30)).Value
let onePicture = Ops.insertPicture (Ops.nextDrawingId noPictures) pngInfo (png 40 30) (at 2 1) 40.0 30.0 None noPictures
let describe (m: SpreadsheetModel) = m.Drawings |> List.map (fun d -> d.Id, d.Name, d.Anchor.Kind, (d.Anchor.FromRow, d.Anchor.FromCol), (d.Anchor.FromOffsetX, d.Anchor.FromOffsetY), (d.Anchor.Width, d.Anchor.Height))
eq "insert: anchored to its cell, with its size and a name" (describe onePicture) [ 1, Some "Picture 1", "oneCell", (2, 1), (0.0, 0.0), (Some 40.0, Some 30.0) ]
let chart : SheetDrawing = { Id = 7; Anchor = Ops.pictureAnchor (at 0 0) 0.0 0.0 100.0 80.0; Name = Some "Chart 1"; Description = None; Content = KeptObject ("chart", "<xdr:graphicFrame/>") }
let withChart = { onePicture with Drawings = onePicture.Drawings @ [ chart ] }
let twoPictures = Ops.insertPicture (Ops.nextDrawingId withChart) pngInfo (png 40 30) (at 5 5) 80.0 60.0 (Some "Logo") withChart
eq "insert: the next id is above every one, the new picture in front" (twoPictures.Drawings |> List.map (fun d -> d.Id, d.Name)) [ 1, Some "Picture 1"; 7, Some "Chart 1"; 8, Some "Logo" ]
let placed = Ops.placePicture 1 (at 4 3) 5.0 7.0 20.0 10.0 twoPictures |> okOf
eq "place: another cell, an offset and a size; the others as they were" (describe placed |> List.head, placed.Drawings.Tail = twoPictures.Drawings.Tail) ((1, Some "Picture 1", "oneCell", (4, 3), (5.0, 7.0), (Some 20.0, Some 10.0)), true)
isError "place: no such picture" (Ops.placePicture 99 (at 0 0) 0.0 0.0 1.0 1.0 twoPictures) "no picture with the id 99"
isError "place: a chart is not a picture" (Ops.placePicture 7 (at 0 0) 0.0 0.0 1.0 1.0 twoPictures) "is not a picture"
isError "delete: a chart is not a picture" (Ops.deletePicture 7 twoPictures) "is not a picture"
eq "delete: by id" ((Ops.deletePicture 1 twoPictures |> okOf).Drawings |> List.map (fun d -> d.Id)) [ 7; 8 ]
eq "delete: an id is not handed out again while a higher one is there" (Ops.nextDrawingId (Ops.deletePicture 1 twoPictures |> okOf)) 9
eq "describe: alternative text and name" ((Ops.describePicture 8 None (Some "our logo") twoPictures |> okOf).Drawings.[2] |> fun d -> d.Name, d.Description) (Some "Logo", Some "our logo")
eq "describe: an empty text removes it" ((Ops.describePicture 8 (Some "L") (Some "") (Ops.describePicture 8 None (Some "x") twoPictures |> okOf) |> okOf).Drawings.[2] |> fun d -> d.Name, d.Description) (Some "L", None)
eq "a two-cell anchor becomes a cell and a size when placed" ((Ops.placePicture 1 (at 1 1) 0.0 0.0 30.0 30.0 { twoPictures with Drawings = twoPictures.Drawings |> List.map (fun d -> if d.Id = 1 then { d with Anchor = { d.Anchor with Kind = "twoCell"; ToRow = Some 9; ToCol = Some 9; EditAs = Some "oneCell" } } else d) } |> okOf).Drawings.[0].Anchor |> fun a -> a.Kind, a.ToRow, a.EditAs) ("oneCell", None, None)
eq "pictures follow rows and columns, keeping their ids" ((applyStructureEdit (InsertRows (0, 2)) (applyStructureEdit (InsertCols (0, 1)) twoPictures)).Drawings |> List.map (fun d -> d.Id, d.Anchor.FromRow, d.Anchor.FromCol)) [ 1, 4, 2; 7, 2, 1; 8, 7, 6 ]
let fixedAnchor =
    { onePicture.Drawings.Head.Anchor with
        Kind = "twoCell"; FromRow = 4; FromCol = 3
        FromOffsetY = 5.0; FromOffsetX = 90.0
        ToRow = Some 7; ToCol = Some 5; ToOffsetY = 35.0; ToOffsetX = 10.0
        EditAs = Some "absolute" }
let fixedSheet =
    { onePicture with
        DefaultRowHeight = 20.0; DefaultColumnWidth = 80.0
        RowHeights = Map.ofList [ 0, 40.0; 2, 0.0 ]
        ColumnWidths = Map.ofList [ 0, 0.0; 2, 140.0 ]
        Drawings = [ { onePicture.Drawings.Head with Anchor = fixedAnchor } ] }
let physicalCorners (sheet: SpreadsheetModel) =
    let a = sheet.Drawings.Head.Anchor
    let pos sizes defaultSize index offset =
        ([ 0 .. index - 1 ] |> List.sumBy (fun i -> Map.tryFind i sizes |> Option.defaultValue defaultSize)) + offset
    pos sheet.RowHeights sheet.DefaultRowHeight a.FromRow a.FromOffsetY,
    pos sheet.ColumnWidths sheet.DefaultColumnWidth a.FromCol a.FromOffsetX,
    pos sheet.RowHeights sheet.DefaultRowHeight a.ToRow.Value a.ToOffsetY,
    pos sheet.ColumnWidths sheet.DefaultColumnWidth a.ToCol.Value a.ToOffsetX
let expectedCorners = (85.0, 310.0, 175.0, 390.0)
eq "absolute anchor: original coordinates include hidden and nonuniform cells" (physicalCorners fixedSheet) expectedCorners
let fixedInserted = fixedSheet |> applyStructureEdit (InsertRows (0, 1)) |> applyStructureEdit (InsertCols (1, 1))
eq "absolute anchor: inserted cells preserve physical corners" (physicalCorners fixedInserted) expectedCorners
let fixedDeleted = fixedSheet |> applyStructureEdit (DeleteRows (0, 1)) |> applyStructureEdit (DeleteCols (2, 1))
eq "absolute anchor: deleting a sized row and column preserves physical corners" (physicalCorners fixedDeleted) expectedCorners

// A note edited without naming an author keeps the author it has.
do
    let at : CellAddress = { Row = 3; Col = 3 }
    let authorOf (model: SpreadsheetModel) = getCell at model |> Option.bind (fun cell -> cell.Comment) |> Option.map (fun note -> note.Text, note.Author)
    let noted = Ops.setComment at "first" (Some "Ann") (createEmpty 10 10)
    eq "a note edited without an author keeps its author" (authorOf (Ops.setComment at "second" None noted)) (Some ("second", Some "Ann"))
    eq "... and takes the author it is given" (authorOf (Ops.setComment at "second" (Some "Bob") noted)) (Some ("second", Some "Bob"))
    eq "runs set without an author keep the note's author" (authorOf (Ops.setCommentRuns at [ "bold", { emptyRunOverride with Bold = Some true }; " text", emptyRunOverride ] None noted)) (Some ("bold text", Some "Ann"))
    eq "a new note without an author has none" (authorOf (Ops.setComment at "new" None (createEmpty 10 10))) (Some ("new", None))

printfn "\nAnnotation tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
