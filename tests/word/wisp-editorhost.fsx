#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../office-wisp/bin/Debug/net10.0/Rbt.Wisp.dll"
#r "../../office-wisp/bin/Debug/net10.0/Rbt.Renderer.dll"
#r "../../office-wisp/bin/Debug/net10.0/Platform.dll"
#r "../../office-wisp/word-render/bin/Debug/net10.0/Rbt.OfficeNative.dll"
#r "../../office-wisp/word-render/bin/Debug/net10.0/Rbt.WordNative.dll"
#load "../../word-wisp/EditorHost.fs"

open Rbt.Word.Document
open Rbt.WordWisp.EditorHost

let mutable checks = 0
let check name expected actual =
    checks <- checks + 1
    if actual <> expected then failwithf "%s: expected %A, got %A" name expected actual

// Fit is a view transform: the same document point selects the same offset
// and the document extent scales without changing line wrapping.
let coordinateDoc = empty |> setCursor "block-0" 0 |> insertText "A coordinate boundary test"
let coordinateSettings = Rbt.WordNative.Render.layoutSettings coordinateDoc
let coordinateLayouts = Rbt.Word.Layout.layoutDocument coordinateSettings Rbt.WordNative.Render.measureText coordinateDoc
let coordinateLine = coordinateLayouts.Head.Lines.Head
let coordinateX = coordinateSettings.PaddingLeft + Rbt.Word.Layout.xForOffset Rbt.WordNative.Render.measureText coordinateLine 5
let coordinateY = coordinateLine.Y + coordinateLine.Height / 2.0
let selectedAt width =
    let host = create coordinateDoc
    let scale = Rbt.WordNative.Render.scale width coordinateDoc
    handlePointer host width Wisp.Layout.Down (coordinateX * scale) (coordinateY * scale)
    host.Doc.Value.Selection.Value.Focus.Offset, documentHeight host width coordinateDoc
let offset1, height1 = selectedAt coordinateDoc.PageSetup.PageWidth
let offset2, height2 = selectedAt (coordinateDoc.PageSetup.PageWidth / 2.0)
check "fit pointer selects same character" offset1 offset2
check "fit scales document height" true (abs (height2 * 2.0 - height1) < 0.001)

let doc = empty |> setCursor "block-0" 0 |> insertText "original"
let st = create doc
applyNoHistory st (setCursor "block-0" 0) |> ignore
check "selection changes do not dirty the document" false st.Session.IsModified
check "selection changes retain original-byte eligibility" true (st.Session.CaptureSave()).Unchanged
edit st (insertText "first ")
let captured = st.Session.CaptureSave()
check "edited document cannot reuse original bytes" false captured.Unchanged
edit st (insertText "later ")
st.Session.MarkSaved captured
check "saving an earlier snapshot leaves later typing dirty" true st.Session.IsModified
undo st |> ignore
check "undo returns to saved snapshot" false st.Session.IsModified
check "saved edits still cannot reuse opened bytes" false (st.Session.CaptureSave()).Unchanged
redo st |> ignore
check "redo past saved snapshot is dirty" true st.Session.IsModified
undo st |> ignore
undo st |> ignore
check "undo to opened document restores original-byte eligibility" true (st.Session.CaptureSave()).Unchanged
check "opened state differs from the saved edited state" true st.Session.IsModified

let replacement = empty |> setCursor "block-0" 0 |> insertText "replacement"
replaceDocument st replacement
edit st (insertText " edited")
st.Session.MarkSaved captured
check "save from replaced document cannot clean replacement edits" true st.Session.IsModified
undo st |> ignore
check "replacement starts clean with independent history" false st.Session.IsModified
undo st |> ignore
check "replacement undo cannot reach previous document" replacement st.Doc.Value
let other = create doc
other.Session.MarkSaved(st.Session.CaptureSave())
check "snapshot from another host cannot change saved state" false other.Session.IsModified
let typed = create (empty |> setCursor "block-0" 0)
let press key =
    handleKey typed 760.0 { Key = key; Shift = false; Ctrl = false; Alt = false } |> ignore
let typedText () =
    (findParagraph typed.Doc.Value "block-0").Value.Buffer |> Rbt.Word.TextBuffer.toString
for key in [ "a"; "b"; "c" ] do press key
check "typing changes the document" "abc" (typedText ())
undo typed |> ignore
check "one undo reverts consecutive typing" "" (typedText ())
redo typed |> ignore
check "redo restores consecutive typing" "abc" (typedText ())
let firstSave = typed.Session.CaptureSave()
press "d"
typed.Session.MarkSaved firstSave
check "captured save does not clean a later edit" true typed.Session.IsModified
undo typed |> ignore
check "undo returns to captured save" false typed.Session.IsModified
press "ArrowLeft"
press "X"
undo typed |> ignore
check "navigation ends the typing run without adding history" "abc" (typedText ())


let styledText (st: EditorState) id =
    let buffer = (findParagraph st.Doc.Value id).Value.Buffer
    Rbt.Word.TextBuffer.toString buffer, Rbt.Word.TextBuffer.getStyleAt 0 buffer
let shortcut st key =
    handleKey st 760.0 { Key = key; Shift = false; Ctrl = true; Alt = false } |> ignore
let key st key =
    handleKey st 760.0 { Key = key; Shift = false; Ctrl = false; Alt = false } |> ignore
let deletion = create (empty |> setCursor "block-0" 0 |> insertText "hello")
key deletion "Backspace"
key deletion "Backspace"
check "consecutive backspaces delete text" "hel" (styledText deletion "block-0" |> fst)
shortcut deletion "z"
check "consecutive backspaces undo together" "hello" (styledText deletion "block-0" |> fst)
shortcut deletion "Home"
key deletion "Delete"
key deletion "Delete"
check "consecutive forward deletes remove text" "llo" (styledText deletion "block-0" |> fst)
shortcut deletion "z"
check "consecutive forward deletes undo together" "hello" (styledText deletion "block-0" |> fst)

let formatted = create (empty |> setCursor "block-0" 0)
shortcut formatted "b"
shortcut formatted "i"
check "formatting at caret leaves document clean" false formatted.Session.IsModified
check "formatting at caret adds no undo" 0 formatted.Session.History.Undos.Length
check "formatting at caret does not alter captured bytes" true (formatted.Session.CaptureSave()).Unchanged
key formatted "a"
key formatted "b"
let runText, runStyle = styledText formatted "block-0"
check "consecutive typed text has pending bold and italic" ("ab", true, true) (runText, runStyle.Bold, runStyle.Italic)
undo formatted |> ignore
check "one undo reverts formatted typing together" "" (styledText formatted "block-0" |> fst)
redo formatted |> ignore
check "redo restores formatted typing" "ab" (styledText formatted "block-0" |> fst)
shortcut formatted "b"
key formatted "c"
check "format change applies to subsequent typing" false
    (Rbt.Word.TextBuffer.getStyleAt 2 (findParagraph formatted.Doc.Value "block-0").Value.Buffer).Bold
undo formatted |> ignore
check "format change separates typing undo" "ab" (styledText formatted "block-0" |> fst)

applyNoHistory formatted (setCursor "block-0" 0) |> ignore
check "caret movement drops pending formatting" true formatted.Session.PendingFormat.IsNone
key formatted "Z"
undo formatted |> ignore
check "movement breaks the formatted typing run" "ab" (styledText formatted "block-0" |> fst)
replaceDocument formatted (empty |> setCursor "block-0" 0)
key formatted "R"
check "replacement drops pending formatting" false (styledText formatted "block-0" |> snd).Bold
check "replacement has independent typing history" 1 formatted.Session.History.Undos.Length

let pendingUndo = create (empty |> setCursor "block-0" 0)
shortcut pendingUndo "u"
key pendingUndo "x"
undo pendingUndo |> ignore
key pendingUndo "y"
check "undo drops pending underline" false (styledText pendingUndo "block-0" |> snd).Underline
let movedPending = create (empty |> setCursor "block-0" 0 |> insertText "x ")
shortcut movedPending "b"
key movedPending "ArrowLeft"
key movedPending "Z"
check "movement does not style text at its destination" false
    (Rbt.Word.TextBuffer.getStyleAt 1 (findParagraph movedPending.Doc.Value "block-0").Value.Buffer).Bold

let selected = create (empty |> setCursor "block-0" 0 |> insertText "one two"
                       |> setSelection { BlockId = "block-0"; Offset = 0 } { BlockId = "block-0"; Offset = 3 })
shortcut selected "b"
let selectedBuffer = (findParagraph selected.Doc.Value "block-0").Value.Buffer
check "selected text formats without replacing text" ("one two", true, false)
    (Rbt.Word.TextBuffer.toString selectedBuffer,
     (Rbt.Word.TextBuffer.getStyleAt 0 selectedBuffer).Bold,
     (Rbt.Word.TextBuffer.getStyleAt 4 selectedBuffer).Bold)
undo selected |> ignore
check "selected formatting is one undo step" false (styledText selected "block-0" |> snd).Bold
let insideWord = create (empty |> setCursor "block-0" 0 |> insertText "hello" |> setCursor "block-0" 2)
shortcut insideWord "b"
check "caret inside word formats the word, not pending typing" (true, true, true)
    ((styledText insideWord "block-0" |> snd).Bold,
     (Rbt.Word.TextBuffer.getStyleAt 4 (findParagraph insideWord.Doc.Value "block-0").Value.Buffer).Bold,
     insideWord.Session.PendingFormat.IsNone)

let character = Rbt.Word.Session.WordCommand.Character
let importedStyle = create (empty |> setCursor "block-0" 0 |> insertText "hello" |> setCursor "block-0" 2)
let styleResult = executeCommand importedStyle (character (Rbt.Word.Session.CharacterCommand.NamedStyle "ImportedCustom")) None
check "named character style inside a word is pending, not an edit"
    (Ok Rbt.Word.Session.NoChange, false, 0, true)
    (styleResult, importedStyle.Session.IsModified, importedStyle.Session.History.Undos.Length,
     importedStyle.Session.PendingFormat.IsSome)
key importedStyle "X"
check "typing retains a character style absent from the local catalog" (Some "ImportedCustom")
    (Rbt.Word.TextBuffer.getStyleAt 2 (findParagraph importedStyle.Doc.Value "block-0").Value.Buffer).Style

let unchangedRange = create (empty |> setCursor "block-0" 0 |> insertText "one two"
                             |> setSelection { BlockId = "block-0"; Offset = 0 } { BlockId = "block-0"; Offset = 3 })
check "range color is one session edit" (Ok Rbt.Word.Session.ContentChanged, 1)
    (executeCommand unchangedRange (character (Rbt.Word.Session.CharacterCommand.Color (Some "#123456"))) None,
     unchangedRange.Session.History.Undos.Length)
let savedColor = unchangedRange.Session.CaptureSave()
unchangedRange.Session.MarkSaved savedColor
undo unchangedRange |> ignore
let undoRevision = unchangedRange.Session.Revision
let redoCount = unchangedRange.Session.History.Redos.Length
check "unnecessary color clear leaves redo and revision untouched"
    (Ok Rbt.Word.Session.NoChange, undoRevision, redoCount)
    (executeCommand unchangedRange (character (Rbt.Word.Session.CharacterCommand.Color None)) None,
     unchangedRange.Session.Revision, unchangedRange.Session.History.Redos.Length)
redo unchangedRange |> ignore
check "redo after no-op restores saved range color" (Some "#123456", false)
    ((Rbt.Word.TextBuffer.getStyleAt 0 (findParagraph unchangedRange.Doc.Value "block-0").Value.Buffer).Color,
     unchangedRange.Session.IsModified)

let pasteHost = create (empty |> setCursor "block-0" 0)
shortcut pasteHost "b"
Clip.text <- "pasted"
shortcut pasteHost "v"
check "pasted text uses pending bold" ("pasted", true)
    ((styledText pasteHost "block-0" |> fst), (styledText pasteHost "block-0" |> snd).Bold)
undo pasteHost |> ignore
check "pending-style paste is one undo step" "" (styledText pasteHost "block-0" |> fst)
let multilinePaste = create (empty |> setCursor "block-0" 0)
shortcut multilinePaste "b"
Clip.text <- "first\nsecond"
shortcut multilinePaste "v"
let multilineParagraphs = multilinePaste.Doc.Value.Blocks |> List.choose (function Paragraph p -> Some p | _ -> None)
check "multiline paste applies pending formatting to each paragraph"
    ([ "first"; "second" ], [ true; true ])
    (multilineParagraphs |> List.map (fun p -> Rbt.Word.TextBuffer.toString p.Buffer),
     multilineParagraphs |> List.map (fun p -> (Rbt.Word.TextBuffer.getStyleAt 0 p.Buffer).Bold))
undo multilinePaste |> ignore
check "multiline paste undoes in one step" "" (styledText multilinePaste "block-0" |> fst)


let emptyHistory = create (empty |> setCursor "block-0" 0)
shortcut emptyHistory "b"
undo emptyHistory |> ignore
redo emptyHistory |> ignore
key emptyHistory "x"
check "empty undo and redo preserve pending style" true (styledText emptyHistory "block-0" |> snd).Bold

let paragraph id text : ParagraphBlock =
    { Id = id; Buffer = Rbt.Word.TextBuffer.ofString text; Metadata = defaultMetadata }
let caret id offset : Selection =
    let pos : Position = { BlockId = id; Offset = offset }
    { Anchor = pos; Focus = pos }
let story main blocks isHeader partIndex box note : Rbt.Word.Session.StorySession =
    { IsHeader = isHeader; PageIndex = 0; Kind = "default"; Section = 0; PartIndex = partIndex
      Label = "story"; Main = main; Edited = main; PartBlocks = blocks; EntryBlocks = blocks
      Folded = ResizeArray(); Box = box; BoxFolded = ResizeArray()
      Note = note; NoteFolded = ResizeArray() }
let body = empty |> setCursor "block-0" 0 |> insertText "Body"
let bodySelection = body.Selection
let header = Paragraph (paragraph "header-text" "Head")
let footer = Paragraph (paragraph "footer-text" "Foot")
let parts =
    { body with
        HeaderParts = [ { Kind = "default"; RelationshipId = Some "header-1"; Blocks = [ header ] } ]
        FooterParts = [ { Kind = "default"; RelationshipId = Some "footer-1"; Blocks = [ footer ] } ]
        PageSetup =
            { body.PageSetup with HeaderReferences = [ "default", "header-1" ]; FooterReferences = [ "default", "footer-1" ] } }
let headerFormat = create parts
enterStory headerFormat (story parts [ header ] true 0 None None)
    (Some { Anchor = { BlockId = "header-text"; Offset = 0 }; Focus = { BlockId = "header-text"; Offset = 4 } })
check "header character formatting edits its story in one undo step"
    (Ok Rbt.Word.Session.ContentChanged, 1, true, false)
    (executeCommand headerFormat (character Rbt.Word.Session.CharacterCommand.Bold) None,
     headerFormat.Session.History.Undos.Length,
     (headerFormat.Session.WholeDocument.HeaderParts.Head.Blocks.Head |> function
      | Paragraph p -> (Rbt.Word.TextBuffer.getStyleAt 0 p.Buffer).Bold | _ -> false),
     (Rbt.Word.TextBuffer.getStyleAt 0 (findParagraph headerFormat.Session.WholeDocument "block-0").Value.Buffer).Bold)
exitStory headerFormat |> ignore
undo headerFormat |> ignore
check "header character formatting undoes after leaving its story" false
    (headerFormat.Session.WholeDocument.HeaderParts.Head.Blocks.Head |> function
     | Paragraph p -> (Rbt.Word.TextBuffer.getStyleAt 0 p.Buffer).Bold | _ -> true)

let storyHost = create parts
enterStory storyHost (story parts [ header ] true 0 None None) (Some (caret "header-text" 4))
edit storyHost (insertText "!")
check "header edit folds into whole document" "Head!"
    ((storyHost.Session.CaptureSave().Document.HeaderParts.Head.Blocks.Head |> function Paragraph p -> Rbt.Word.TextBuffer.toString p.Buffer | _ -> ""))
check "header edit leaves body and footer intact" (parts.Blocks, parts.FooterParts)
    (storyHost.Session.WholeDocument.Blocks, storyHost.Session.WholeDocument.FooterParts)
check "header edit marks whole document dirty" true storyHost.Session.IsModified
let headerSave = storyHost.Session.CaptureSave()
storyHost.Session.MarkSaved headerSave
check "saving folded header cleans the document" false storyHost.Session.IsModified
check "exiting header restores body selection" true (exitStory storyHost)
check "body caret restored after header" bodySelection storyHost.Doc.Value.Selection
undo storyHost |> ignore
check "undo outside header re-enters header story" true storyHost.Session.Story.IsSome
check "undo removes header edit from saved document" "Head"
    ((storyHost.Session.CaptureSave().Document.HeaderParts.Head.Blocks.Head |> function Paragraph p -> Rbt.Word.TextBuffer.toString p.Buffer | _ -> ""))
check "undoing saved header dirties the document" true storyHost.Session.IsModified
redo storyHost |> ignore
check "redo restores header edit" "Head!"
    ((storyHost.Session.CaptureSave().Document.HeaderParts.Head.Blocks.Head |> function Paragraph p -> Rbt.Word.TextBuffer.toString p.Buffer | _ -> ""))
check "redo of saved header returns to clean revision" false storyHost.Session.IsModified
let replacedStory = create parts
enterStory replacedStory (story parts [ header ] true 0 None None) (Some (caret "header-text" 4))
replaceDocument replacedStory body
check "replacing an active story resets context and save identity" (false, true)
    (replacedStory.Session.Story.IsSome, (replacedStory.Session.CaptureSave()).Unchanged)

exitStory storyHost |> ignore
let current = storyHost.Session.WholeDocument
enterStory storyHost (story current [ footer ] false 0 None None) (Some (caret "footer-text" 4))
edit storyHost (insertText "?")
check "footer edit folds into saved document" "Foot?"
    ((storyHost.Session.CaptureSave().Document.FooterParts.Head.Blocks.Head |> function Paragraph p -> Rbt.Word.TextBuffer.toString p.Buffer | _ -> ""))
check "footer edit after saved header is dirty" true storyHost.Session.IsModified
exitStory storyHost |> ignore
check "body caret restored after footer" bodySelection storyHost.Doc.Value.Selection
undo storyHost |> ignore
check "undo enters footer instead of previous header" false storyHost.Session.Story.Value.IsHeader
check "undo of footer returns to saved whole document" false storyHost.Session.IsModified
exitStory storyHost |> ignore
check "header remains after undo of footer" "Head!"
    ((storyHost.Session.WholeDocument.HeaderParts.Head.Blocks.Head |> function Paragraph p -> Rbt.Word.TextBuffer.toString p.Buffer | _ -> ""))

let noteBlocks = [ Paragraph (paragraph "note-text" "Note") ]
let noteBody : NoteBody =
    { NoteType = "footnote"; Id = "1"; Kind = None; HasMark = true; MarkStyle = None
      Text = "Note"; Blocks = noteBlocks }
let noteMain =
    { body with Package = { body.Package with Notes = [ noteBody ] } }
let noteHost = create noteMain
enterStory noteHost (story noteMain noteBlocks false -1 None (Some ("footnote", "1"))) (Some (caret "note-text" 4))
edit noteHost (insertText "!")
let noteSave = noteHost.Session.CaptureSave()
check "note edit appears in saved whole document" "Note!"
    (noteSave.Document.Package.Notes.Head.Blocks.Head |> function Paragraph p -> Rbt.Word.TextBuffer.toString p.Buffer | _ -> "")
check "note edit preserves body" noteMain.Blocks noteSave.Document.Blocks
exitStory noteHost |> ignore
check "body caret restored after note" bodySelection noteHost.Doc.Value.Selection
undo noteHost |> ignore
check "note undo restores original note" "Note" (noteHost.Session.WholeDocument.Package.Notes.Head.Text)

let boxPara = paragraph "box-text" "Box"
let textBox = { defaultTextBox with Id = "story-box"; Content = [ boxPara ]; AnchorBlockId = Some "block-0" }
let boxMain = { body with Blocks = body.Blocks @ [ TextBox textBox ] }
let boxHost = create boxMain
let boxBlocks = [ Paragraph boxPara ]
enterStory boxHost (story boxMain boxBlocks false -1 (Some "story-box") None) (Some (caret "box-text" 3))
edit boxHost (insertText "!")
let boxSave = boxHost.Session.CaptureSave()
check "textbox edit appears in saved whole document" "Box!"
    (boxSave.Document.Blocks |> List.pick (function TextBox b -> Some (Rbt.Word.TextBuffer.toString b.Content.Head.Buffer) | _ -> None))
exitStory boxHost |> ignore
check "body caret restored after textbox" bodySelection boxHost.Doc.Value.Selection
undo boxHost |> ignore
check "textbox undo restores original" "Box"
    (boxHost.Session.WholeDocument.Blocks |> List.pick (function TextBox b -> Some (Rbt.Word.TextBuffer.toString b.Content.Head.Buffer) | _ -> None))

let lazyHost = create body
let lazyHeader =
    Rbt.Word.Session.headerFooterStory body 0 true "first" 0 None "first page header" id
    |> Option.get
enterStory lazyHost lazyHeader None
check "entering absent header does not create document part" true (lazyHost.Session.CaptureSave()).Unchanged
let lazyId = lazyHost.Doc.Value.Selection.Value.Focus.BlockId
key lazyHost "H"
let lazySave = lazyHost.Session.CaptureSave()
let lazyText = lazySave.Document.HeaderParts.Head.Blocks.Head |> function Paragraph p -> Rbt.Word.TextBuffer.toString p.Buffer | _ -> ""
let lazyReferenced = lazySave.Document.PageSetup.HeaderReferences |> List.exists (fun (kind, _) -> kind = "first")
check "first header edit creates a referenced part" ("H", true) (lazyText, lazyReferenced)
check "new header remains a story while typing" lazyId lazyHost.Doc.Value.Selection.Value.Focus.BlockId
undo lazyHost |> ignore
check "undo of new header removes part and reference" ([], []) (lazyHost.Session.WholeDocument.HeaderParts, lazyHost.Session.WholeDocument.PageSetup.HeaderReferences)
check "undo to original whole document is clean" false lazyHost.Session.IsModified
exitStory lazyHost |> ignore
check "new header exit restores body caret" bodySelection lazyHost.Doc.Value.Selection

// Wisp structural commands and keyboard edits use the shared session.
let structural = create (empty |> setCursor "block-0" 0 |> insertText "one")
let run (st: EditorState) cmd = executeCommand st cmd None
check "Wisp alignment changes the paragraph atomically" (Ok Rbt.Word.Session.ContentChanged, Center, 1)
    (run structural Rbt.Word.Session.AlignCenter,
     (findParagraph structural.Doc.Value "block-0").Value.Metadata.Alignment,
     structural.Session.History.Undos.Length)
let savedAlignment = structural.Session.CaptureSave()
structural.Session.MarkSaved savedAlignment
let unchangedAligned = structural.Doc.Value
check "Wisp repeat alignment preserves the saved object" (Ok Rbt.Word.Session.NoChange, true, false)
    (run structural Rbt.Word.Session.AlignCenter,
     obj.ReferenceEquals(unchangedAligned, structural.Doc.Value), structural.Session.IsModified)
undo structural |> ignore
let beforeNoOp = structural.Doc.Value
let redoDepth = structural.Session.History.Redos.Length
check "Wisp out-of-table navigation preserves redo" (Ok Rbt.Word.Session.NoChange, true, redoDepth)
    (run structural (Rbt.Word.Session.MoveCell false),
     obj.ReferenceEquals(beforeNoOp, structural.Doc.Value), structural.Session.History.Redos.Length)
redo structural |> ignore
check "Wisp redo restores saved paragraph alignment" (Center, false)
    ((findParagraph structural.Doc.Value "block-0").Value.Metadata.Alignment, structural.Session.IsModified)
check "Wisp list command is an atomic change" (Ok Rbt.Word.Session.ContentChanged, Bullet)
    (run structural Rbt.Word.Session.ToggleBulletList,
     (findParagraph structural.Doc.Value "block-0").Value.Metadata.ListStyle)
undo structural |> ignore
check "Wisp list undo returns to saved paragraph" (NoList, false)
    ((findParagraph structural.Doc.Value "block-0").Value.Metadata.ListStyle, structural.Session.IsModified)

let tableState = create (empty |> setCursor "block-0" 0)
let tableOf (st: EditorState) : TableBlock =
    st.Doc.Value.Blocks |> List.pick (function Table table -> Some table | _ -> None)
check "Wisp inserts table in one edit" (Ok Rbt.Word.Session.ContentChanged, 2, 1)
    (run tableState (Rbt.Word.Session.InsertTable(2, 2)), (tableOf tableState).Rows.Length,
     tableState.Session.History.Undos.Length)
let firstCellId = tableState.Doc.Value.Selection.Value.Focus.BlockId
for _ in 1 .. 3 do key tableState "Tab"
check "Wisp selection-only cell navigation does not add undo" (false, 1)
    (firstCellId = tableState.Doc.Value.Selection.Value.Focus.BlockId, tableState.Session.History.Undos.Length)
let lastCellId = tableState.Doc.Value.Selection.Value.Focus.BlockId
key tableState "Tab"
check "Wisp last-cell Tab inserts exactly one row transaction" (3, 2)
    ((tableOf tableState).Rows.Length, tableState.Session.History.Undos.Length)
undo tableState |> ignore
check "Wisp row append undo restores original table caret" (2, lastCellId)
    ((tableOf tableState).Rows.Length, tableState.Doc.Value.Selection.Value.Focus.BlockId)
redo tableState |> ignore
check "Wisp row append redo restores row" 3 ((tableOf tableState).Rows.Length)
let tableSave = tableState.Session.CaptureSave()
tableState.Session.MarkSaved tableSave
let firstTablePara = (tableOf tableState).Rows.Head.Cells.Head.Content.Head.Id
applyNoHistory tableState (setCursor firstTablePara 0) |> ignore
let firstCellCount () = (tableOf tableState).Rows.Head.Cells.Head.Content.Length
key tableState "Enter"
check "Wisp Enter splits inside the table cell" 2 (firstCellCount ())
undo tableState |> ignore
check "Wisp table Enter undo returns to saved table" (1, false) (firstCellCount (), tableState.Session.IsModified)
let beforeShiftTab = tableState.Doc.Value
handleKey tableState 760.0 { Key = "Tab"; Shift = true; Ctrl = false; Alt = false } |> ignore
check "Wisp ShiftTab in first cell leaves saved table untouched" (true, false)
    (obj.ReferenceEquals(beforeShiftTab, tableState.Doc.Value), tableState.Session.IsModified)
handleKey tableState 760.0 { Key = "Tab"; Shift = false; Ctrl = true; Alt = false } |> ignore
check "Wisp CtrlTab inserts a tab within the cell" ("\t", firstTablePara)
    (Rbt.Word.TextBuffer.toString ((tableOf tableState).Rows.Head.Cells.Head.Content.Head.Buffer),
     tableState.Doc.Value.Selection.Value.Focus.BlockId)
undo tableState |> ignore
check "Wisp CtrlTab undo restores saved table" false tableState.Session.IsModified


let listState = create (empty |> setCursor "block-0" 0 |> insertText "list")
run listState Rbt.Word.Session.ToggleNumberedList |> ignore
key listState "Enter"
let listParagraphs () = listState.Doc.Value.Blocks |> List.choose (function Paragraph para -> Some para | _ -> None)
check "Wisp Enter continues the numbered list" (2, Numbered)
    ((listParagraphs ()).Length, (listParagraphs ()).[1].Metadata.ListStyle)
key listState "Tab"
check "Wisp Tab indents the list item" 2 ((listParagraphs ()).[1].Metadata.ListLevel)
undo listState |> ignore
check "Wisp list indentation has its own undo" 1 ((listParagraphs ()).[1].Metadata.ListLevel)
redo listState |> ignore
handleKey listState 760.0 { Key = "Tab"; Shift = true; Ctrl = false; Alt = false } |> ignore
check "Wisp ShiftTab restores original list level" 1 ((listParagraphs ()).[1].Metadata.ListLevel)
key listState "Enter"
check "Wisp empty list Enter leaves list without adding a paragraph" (2, NoList)
    ((listParagraphs ()).Length, (listParagraphs ()).[1].Metadata.ListStyle)

let trackedState = create (empty |> setCursor "block-0" 0 |> insertText "tracked")
let tracked = executeCommand trackedState Rbt.Word.Session.Enter (Some ("Wisp author", "2026-11-12T12:00:00Z"))
let trackedMark = (findParagraph trackedState.Doc.Value "block-0").Value.Metadata.MarkRevision
check "Wisp tracked Enter keeps author and date on paragraph mark"
    (Ok Rbt.Word.Session.ContentChanged, Some "Wisp author", Some "2026-11-12T12:00:00Z")
    (tracked, trackedMark |> Option.bind (fun revision -> revision.Mark.Author),
     trackedMark |> Option.bind (fun revision -> revision.Mark.Date))
let rejectedBreak = create parts
enterStory rejectedBreak (story parts [ header ] true 0 None None) (Some (caret "header-text" 4))
let beforeBreak = rejectedBreak.Doc.Value
check "Wisp story rejects page break without touching saved revision" (Ok Rbt.Word.Session.NoChange, true, true)
    (run rejectedBreak Rbt.Word.Session.PageBreak,
     obj.ReferenceEquals(beforeBreak, rejectedBreak.Doc.Value), (rejectedBreak.Session.CaptureSave()).Unchanged)
handleKey rejectedBreak 760.0 { Key = "Enter"; Shift = false; Ctrl = true; Alt = false } |> ignore
check "Wisp CtrlEnter also rejects a header break" (true, 0)
    (obj.ReferenceEquals(beforeBreak, rejectedBreak.Doc.Value), rejectedBreak.Session.History.Undos.Length)


// Clipboard captures preserve source formatting across hosts; cut commits only
// after transfer and still belongs to the source story's undo history.
let richSource =
    create (empty |> setCursor "block-0" 0 |> insertText "Bold" |> setSelection
        { BlockId = "block-0"; Offset = 0 } { BlockId = "block-0"; Offset = 4 })
executeCommand richSource (character Rbt.Word.Session.CharacterCommand.Bold) None |> ignore
shortcut richSource "c"
let richTarget = create (empty |> setCursor "block-0" 0)
shortcut richTarget "v"
let richBuffer = (findParagraph richTarget.Doc.Value "block-0").Value.Buffer
check "Wisp cross-host copy keeps rich character formatting and atomic paste" ("Bold", true, 1)
    (Rbt.Word.TextBuffer.toString richBuffer, (Rbt.Word.TextBuffer.getStyleAt 0 richBuffer).Bold,
     richTarget.Session.History.Undos.Length)
undo richTarget |> ignore
check "rich paste undo leaves target empty" "" (styledText richTarget "block-0" |> fst)
shortcut richSource "x"
check "Wisp cut removes selected source in one undo step" ("", 2)
    ((findParagraph richSource.Doc.Value "block-0").Value.Buffer |> Rbt.Word.TextBuffer.toString,
     richSource.Session.History.Undos.Length)
undo richSource |> ignore
check "cut undo restores selected formatted text" ("Bold", true)
    (styledText richSource "block-0" |> fun (text, style) -> text, style.Bold)

let trackedPaste = create (empty |> setCursor "block-0" 0)
let richTransfer = Clip.capture.Value
let trackedOutcome =
    trackedPaste.Session.PasteClipboard(richTransfer.Text, richTransfer.Html, (fun _ -> Some richTransfer.Clip),
                                        Some ("Paste author", "2026-11-12T12:00:00Z"))
let trackedBuffer = (findParagraph trackedPaste.Session.Document "block-0").Value.Buffer
let trackedStyle = Rbt.Word.TextBuffer.getStyleAt 0 trackedBuffer
check "tracked paste inserts a revision instead of transferring source formatting"
    (Rbt.Word.Session.ContentChanged, "Bold", false, Some "insert", 1)
    (trackedOutcome, Rbt.Word.TextBuffer.toString trackedBuffer, trackedStyle.Bold,
     trackedStyle.RevisionType, trackedPaste.Session.History.Undos.Length)

let markedBuffer =
    Rbt.Word.TextBuffer.ofString "Link"
    |> Rbt.Word.TextBuffer.setHyperlink 0 4 (Some "https://example.com")
    |> Rbt.Word.TextBuffer.setBookmark 0 4 "bookmark-1" "source-mark"
let markedSource =
    create
        { empty with
            Blocks = [ Paragraph { (paragraph "block-0" "") with Buffer = markedBuffer } ]
            Selection = Some { Anchor = { BlockId = "block-0"; Offset = 0 }
                               Focus = { BlockId = "block-0"; Offset = 4 } } }
let markedCapture = markedSource.Session.CaptureClipboard(fun _ -> "marked-html") |> Option.get
let markedTarget = create (empty |> setCursor "block-0" 0)
let markedOutcome =
    markedTarget.Session.PasteClipboard(markedCapture.Text, markedCapture.Html,
                                        (fun _ -> Some markedCapture.Clip), None)
let markedStyle = (findParagraph markedTarget.Session.Document "block-0").Value.Buffer |> Rbt.Word.TextBuffer.getStyleAt 0
check "shared clipboard keeps hyperlink but never clones bookmark identity"
    (Rbt.Word.Session.ContentChanged, Some "https://example.com", [], [], 1)
    (markedOutcome, markedStyle.Hyperlink, markedStyle.BookmarkStarts, markedStyle.BookmarkEnds,
     markedTarget.Session.History.Undos.Length)

let stale = create (empty |> setCursor "block-0" 0 |> insertText "Original" |> setSelection
    { BlockId = "block-0"; Offset = 0 } { BlockId = "block-0"; Offset = 8 })
let staleCapture = stale.Session.CaptureClipboard(fun _ -> "internal-marker") |> Option.get
key stale "ArrowLeft"
check "moving the selection cancels deferred cut without changing revision" (None, false, 0)
    (stale.Session.CommitCut(staleCapture, None), stale.Session.IsModified, stale.Session.History.Undos.Length)
applyNoHistory stale (setSelection { BlockId = "block-0"; Offset = 0 } { BlockId = "block-0"; Offset = 8 }) |> ignore
check "restoring the same selection does not revive an asynchronous cut" (None, "Original")
    (stale.Session.CommitCut(staleCapture, None),
     (findParagraph stale.Doc.Value "block-0").Value.Buffer |> Rbt.Word.TextBuffer.toString)
replaceDocument stale (empty |> setCursor "block-0" 0 |> insertText "Replacement")
check "replacement invalidates deferred cut even with same paragraph id" (None, "Replacement")
    (stale.Session.CommitCut(staleCapture, None),
     (findParagraph stale.Doc.Value "block-0").Value.Buffer |> Rbt.Word.TextBuffer.toString)
let noopDoc = stale.Session.Document
let noopHistory = stale.Session.History
check "empty clipboard paste preserves snapshot and redo" (Rbt.Word.Session.NoChange, true, true)
    (stale.Session.PasteClipboard("", "", Rbt.Word.Session.noExternalClip, None),
     obj.ReferenceEquals(noopDoc, stale.Session.Document),
     obj.ReferenceEquals(noopHistory, stale.Session.History))

let headerClipboard = create parts
enterStory headerClipboard (story parts [ header ] true 0 None None)
    (Some { Anchor = { BlockId = "header-text"; Offset = 0 }; Focus = { BlockId = "header-text"; Offset = 4 } })
shortcut headerClipboard "x"
check "cut in header leaves body untouched and records whole-story edit" ("", parts.Blocks, 1)
    ((headerClipboard.Session.WholeDocument.HeaderParts.Head.Blocks.Head |> function Paragraph p -> Rbt.Word.TextBuffer.toString p.Buffer | _ -> "?"),
     headerClipboard.Session.WholeDocument.Blocks, headerClipboard.Session.History.Undos.Length)
undo headerClipboard |> ignore
check "undo of header cut restores header and body" ("Head", parts.Blocks)
    ((headerClipboard.Session.WholeDocument.HeaderParts.Head.Blocks.Head |> function Paragraph p -> Rbt.Word.TextBuffer.toString p.Buffer | _ -> "?"),
     headerClipboard.Session.WholeDocument.Blocks)

let clipboardCell text : ClipParagraph =
    { Buffer = Rbt.Word.TextBuffer.ofString text; Metadata = defaultMetadata
      HeadingLevel = None; TableRows = [] }
let tableClip : Clip =
    { Text = "A\tB"
      Paragraphs =
        [ { Buffer = Rbt.Word.TextBuffer.empty; Metadata = defaultMetadata; HeadingLevel = None
            TableRows = [ [ [ clipboardCell "A" ]; [ clipboardCell "B" ] ] ] } ] }
let boxPaste = create boxMain
enterStory boxPaste (story boxMain boxBlocks false -1 (Some "story-box") None) (Some (caret "box-text" 3))
let boxPasteOutcome = boxPaste.Session.PasteClipboard(tableClip.Text, "external-html", (fun _ -> Some tableClip), None)
let boxPasteBody = boxPaste.Session.WholeDocument.Blocks |> List.filter (function TextBox _ -> false | _ -> true)
let boxOriginalBody = boxMain.Blocks |> List.filter (function TextBox _ -> false | _ -> true)
check "paste table into textbox flattens cells and keeps body in one undo step"
    (Rbt.Word.Session.ContentChanged, "BoxA\tB", true, 1)
    (boxPasteOutcome,
     (findParagraph boxPaste.Session.Document "box-text").Value.Buffer |> Rbt.Word.TextBuffer.toString,
     boxPasteBody = boxOriginalBody, boxPaste.Session.History.Undos.Length)

undo boxPaste |> ignore
check "textbox paste undo restores only box content" ("Box", true)
    ((findParagraph boxPaste.Session.Document "box-text").Value.Buffer |> Rbt.Word.TextBuffer.toString,
     boxPaste.Session.WholeDocument.Blocks = boxMain.Blocks)

let wordNavigation = create (empty |> setCursor "block-0" 0 |> insertText "one two" |> setCursor "block-0" 0)
shortcut wordNavigation "ArrowRight"
handleKey wordNavigation 760.0 { Key = "ArrowRight"; Shift = true; Ctrl = true; Alt = false } |> ignore
let wordSelection = wordNavigation.Session.Document.Selection.Value
check "Ctrl-arrow moves by word and Ctrl-Shift-arrow selects the next word without editing"
    (4, 7, 0, false)
    (wordSelection.Anchor.Offset, wordSelection.Focus.Offset,
     wordNavigation.Session.History.Undos.Length, wordNavigation.Session.IsModified)
shortcut wordNavigation "ArrowLeft"
check "Ctrl-left collapses a word selection without history" (4, true, 0)
    (wordNavigation.Doc.Value.Selection.Value.Focus.Offset,
     wordNavigation.Doc.Value.Selection.Value.IsCollapsed, wordNavigation.Session.History.Undos.Length)
handleKey wordNavigation 760.0 { Key = "Home"; Shift = true; Ctrl = true; Alt = false } |> ignore
check "Ctrl-Shift-Home extends selection to document start" (4, 0)
    (wordNavigation.Doc.Value.Selection.Value.Anchor.Offset, wordNavigation.Doc.Value.Selection.Value.Focus.Offset)

printfn "Word Wisp file-session checks: %d passed" checks
