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
eq "note: replaced" (Ops.commentAt renoted (at 1 1) |> Option.map (fun c -> c.Text, c.Author)) (Some ("checked", None))
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

printfn "\nAnnotation tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
