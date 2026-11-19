#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"

// The outline of a document (Outline.fs): levels from a paragraph's own outline level, its style's
// and the built-in heading styles; the headings and the subtext of a heading; promote, demote,
// promote to Heading 1, demote to body text, a level - as Word 2019 does them (checked with VBA on
// the Windows VM: Paragraph.OutlinePromote / OutlineDemote / OutlineDemoteToBody / OutlineLevel);
// a heading moved with its subtext past its sibling, blocks moved, a heading deleted with its
// subtext; each command one undo step of the editing session, and a save read again keeps it.

open Rbt.Word.Document
module Doc = Rbt.Word.Document
module TB = Rbt.Word.TextBuffer
module Outline = Rbt.Word.Outline
module Session = Rbt.Word.Session

let mutable pass = 0
let mutable fail = 0
let check label got expected =
    if got = expected then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-50s expected %A got %A" label expected got

let normal : NamedStyle =
    { Outline.builtInHeadingStyle 1 with
        Id = "Normal"; Name = Some "Normal"; IsDefault = true; BasedOn = None; Next = None; Linked = None
        Text = None; TextItalic = None; SpacingBefore = None; SpacingAfter = Some 8.0; OutlineLevel = None
        KeepNext = None; KeepLines = None; UiPriority = Some 0; SemiHidden = false; UnhideWhenUsed = false }
let custom : NamedStyle = { normal with Id = "MyHead"; Name = Some "MyHead"; IsDefault = false; BasedOn = Some "Normal"; OutlineLevel = Some 2 }
let styles = [ normal; Outline.builtInHeadingStyle 1; Outline.builtInHeadingStyle 2; Outline.builtInHeadingStyle 9; custom ]

let mutable counter = 0
let para (style: string option) (text: string) : ParagraphBlock =
    counter <- counter + 1
    { Id = sprintf "p%d" counter; Buffer = TB.ofString text; Metadata = { defaultMetadata with Style = style } }
let h level text = Heading (min 3 level, para (Some (sprintf "Heading%d" level)) text)
let n text = Paragraph (para (Some "Normal") text)
let direct level text =
    let p = para (Some "Normal") text
    Paragraph { p with Metadata = { p.Metadata with OutlineLevel = Some (level - 1) } }
let mine text = Paragraph (para (Some "MyHead") text)
let table text =
    let host = para None "x"
    let made = Doc.insertTable 1 1 (Doc.setCursor host.Id 0 { Doc.empty with Blocks = [ Paragraph host ] })
    let t = made.Blocks |> List.pick (function Table t -> Some t | _ -> None)
    let cell = t.Rows.Head.Cells.Head
    let content = { cell.Content.Head with Buffer = TB.ofString text; Metadata = { cell.Content.Head.Metadata with Style = Some "Normal" } }
    Table { t with Rows = [ { t.Rows.Head with Cells = [ { cell with Content = [ content ] } ] } ] }
let docOf blocks = { Doc.empty with Blocks = blocks; Styles = styles; Selection = None }

let textOf (block: Block) =
    match block with
    | Paragraph p | Heading (_, p) -> TB.toString p.Buffer
    | Table t -> "[" + (t.Rows.[0].Cells.[0].Content |> List.map (fun p -> TB.toString p.Buffer) |> String.concat "") + "]"
    | TextBox _ -> "[box]"
let texts (doc: Document) = doc.Blocks |> List.map textOf
let levels (doc: Document) = (Outline.ofDocument doc).Levels |> Array.toList
let styleOf (doc: Document) (index: int) =
    match List.item index doc.Blocks with
    | Paragraph p | Heading (_, p) -> p.Metadata.Style |> Option.defaultValue ""
    | _ -> ""
let ownLevelOf (doc: Document) (index: int) =
    match List.item index doc.Blocks with
    | Paragraph p | Heading (_, p) -> p.Metadata.OutlineLevel
    | _ -> None
let idAt (doc: Document) (index: int) = (Outline.ofDocument doc).BlockIds.[index]
let caretIn (doc: Document) (index: int) = Doc.setCursor (idAt doc index) 0 doc
let selectFrom (doc: Document) (first: int) (last: int) =
    Doc.setSelection { BlockId = idAt doc first; Offset = 0 } { BlockId = idAt doc last; Offset = 0 } doc

// The document of the VM's checks: Heading 1 "A", body, Heading 2 "B", body, a paragraph of level 3
// by its own outline level, one of level 3 by its style, a table, Heading 1 "C".
let sample () = docOf [ h 1 "A"; n "a body"; h 2 "B"; n "b body"; direct 3 "direct3"; mine "custom3"; table "cell"; h 1 "C" ]
let d0 = sample ()

// ---- levels, headings, subtext --------------------------------------------------------------------------
check "levels: own, style, built-in, body, table" (levels d0) [ 1; 10; 2; 10; 3; 3; 10; 1 ]
check "headings: their places" ((Outline.ofDocument d0).Headings |> Array.toList) [ 0; 2; 4; 5; 7 ]
let o0 = Outline.ofDocument d0
check "the outline is computed once for a state of the document" (obj.ReferenceEquals(Outline.ofDocument d0, o0)) true
check "subtext: a heading's runs to the next heading of its level or higher" (Outline.subtextEnd o0 0) 7
check "subtext: of Heading 2 too (deeper headings are in it)" (Outline.subtextEnd o0 2) 7
check "subtext: a level-3 paragraph ends at the next of level 3" (Outline.subtextEnd o0 4) 5
check "subtext: the last heading's runs to the end" (Outline.subtextEnd o0 7) 8
check "subtext: body text has none" (Outline.subtextEnd o0 1) 2
check "heading at: a table is under the heading before it" (Outline.headingAt o0 6) 5
check "heading at: before the first heading" (Outline.headingAt (Outline.ofDocument (docOf [ n "x"; h 1 "A" ])) 0) -1
check "a heading style in a table cell is no heading of the outline" (levels (docOf [ table "cell" ])) [ 10 ]
check "a document without styles: a heading made here is its level" (levels { Doc.empty with Blocks = [ Heading (2, para None "made here"); Paragraph (para None "body") ] }) [ 2; 10 ]
let undefinedHeading = { Doc.empty with Blocks = [ Paragraph (para (Some "Heading4") "four") ] }
check "a Heading N the document does not define is Word's" (levels undefinedHeading) [ 4 ]
let named = { normal with Id = "berschrift1"; Name = Some "heading 1"; IsDefault = false; OutlineLevel = None }
check "a built-in heading by its name, whatever its id" (levels { Doc.empty with Styles = [ normal; named ]; Blocks = [ Paragraph (para (Some "berschrift1") "x") ] }) [ 1 ]
let ownOnHeading =
    let p = para (Some "Heading1") "x"
    docOf [ Heading (1, { p with Metadata = { p.Metadata with OutlineLevel = Some 4 } }) ]
check "a built-in heading's level is its style's, not the paragraph's own" (levels ownOnHeading) [ 1 ]

// ---- promote / demote (Word's results on the VM) ---------------------------------------------------------
let promote index doc = Outline.changeSelectedLevels Outline.Promote (caretIn doc index)
let demote index doc = Outline.changeSelectedLevels Outline.Demote (caretIn doc index)
let p1 = promote 1 d0
check "body text promoted is the level of the heading before it (Heading 1)" (styleOf p1 1, levels p1 |> List.item 1) ("Heading1", 1)
check "promoted body text is a Heading block, as the reader makes one" (match List.item 1 p1.Blocks with Heading (1, _) -> true | _ -> false) true
let p2 = demote 1 d0
check "body text demoted is one level below the heading before it" (styleOf p2 1, levels p2 |> List.item 1) ("Heading2", 2)
let p3 = promote 3 d0
check "body text after Heading 2 promoted is Heading 2" (styleOf p3 3) "Heading2"
let p4 = promote 2 d0
check "Heading 2 promoted is Heading 1" (styleOf p4 2, levels p4 |> List.item 2) ("Heading1", 1)
check "Heading 1 promoted stays (the same document)" (obj.ReferenceEquals(promote 0 d0 |> fun d -> d.Blocks, d0.Blocks)) true
let p5 = demote 0 d0
check "Heading 1 demoted is Heading 2" (styleOf p5 0) "Heading2"
let p6 = promote 4 d0
check "a paragraph's own level promoted: its own level changes, its style stays" (styleOf p6 4, ownLevelOf p6 4, levels p6 |> List.item 4) ("Normal", Some 1, 2)
let p7 = demote 4 p6
check "and demoted again" (ownLevelOf p7 4) (Some 2)
let p8 = promote 5 d0
check "a style's own level promoted: the paragraph is given a level of its own, the style stays" (styleOf p8 5, ownLevelOf p8 5, levels p8 |> List.item 5) ("MyHead", Some 1, 2)
let nine = docOf [ h 1 "A"; h 9 "nine"; h 2 "two" ]
check "Heading 9 demoted stays" (obj.ReferenceEquals((demote 1 nine).Blocks, nine.Blocks)) true
let p9 = demote 2 nine
check "Heading 3 that the document does not define is added (Word's built-in)" (styleOf p9 2, (p9.Styles |> List.exists (fun s -> s.Id = "Heading3")), levels p9 |> List.item 2) ("Heading3", true, 3)
check "an added heading style is Word's (outline level, size)" (p9.Styles |> List.find (fun s -> s.Id = "Heading3") |> fun s -> s.OutlineLevel, s.BasedOn) (Some 2, Some "Normal")
let toBody = Outline.changeSelectedLevels Outline.DemoteToBody (caretIn d0 5)
check "demote to body text: Normal, no level of its own" (styleOf toBody 5, ownLevelOf toBody 5, levels toBody |> List.item 5) ("Normal", None, 10)
let toBody2 = Outline.changeSelectedLevels Outline.DemoteToBody (caretIn d0 4)
check "demote to body text: a paragraph's own level goes" (ownLevelOf toBody2 4, levels toBody2 |> List.item 4) (None, 10)
let spaced =
    let p = para (Some "Heading1") "spaced"
    docOf [ Heading (1, { p with Metadata = { p.Metadata with SpacingAfter = Some 30.0; Explicit = ParagraphExplicit.SpacingAfter } }) ]
let spacedDemoted = demote 0 spaced
check "demote keeps the paragraph's own formatting (as Word)" (match spacedDemoted.Blocks.Head with Heading (_, p) -> p.Metadata.SpacingAfter | _ -> None) (Some 30.0)
let spacedBody = Outline.changeSelectedLevels Outline.DemoteToBody (caretIn spaced 0)
check "demote to body text takes the style's spacing (as Word)" (match spacedBody.Blocks.Head with Paragraph p -> p.Metadata.SpacingAfter | _ -> Some -1.0) None
let toOne = Outline.changeSelectedLevels Outline.PromoteToHeading1 (selectFrom d0 1 2)
check "promote to Heading 1: every selected paragraph" (styleOf toOne 1, styleOf toOne 2) ("Heading1", "Heading1")
let multi = docOf [ h 1 "A"; h 2 "B"; h 3 "C"; n "body" ]
let multiPromoted = Outline.changeSelectedLevels Outline.Promote (selectFrom multi 1 3)
check "a selection with headings: headings promoted, body text left (as Word)" (levels multiPromoted) [ 1; 1; 2; 10 ]
let bodies = docOf [ h 2 "A"; n "x"; n "y"; n "z" ]
let bodiesPromoted = Outline.changeSelectedLevels Outline.Promote (selectFrom bodies 1 2)
check "body text alone promoted: each the level of the heading before" (levels bodiesPromoted) [ 2; 2; 2; 10 ]
let firstBody = docOf [ n "first"; n "second"; h 1 "A" ]
check "body text before any heading promoted is Heading 1" (levels (promote 0 firstBody)) [ 1; 10; 1 ]
check "body text before any heading demoted is Heading 2" (levels (demote 0 firstBody)) [ 2; 10; 1 ]
let afterOwn = docOf [ h 1 "A"; direct 3 "d"; n "after" ]
check "body text after a paragraph of level 3 promoted is Heading 3" (styleOf (promote 2 afterOwn) 2) "Heading3"
let toLevel level index doc = Outline.changeSelectedLevels (Outline.ToLevel level) (caretIn doc index)
check "a level chosen for body text: that heading style" (styleOf (toLevel 4 1 d0) 1) "Heading4"
check "a level chosen for a paragraph with its own level: its own" (ownLevelOf (toLevel 6 4 d0) 4) (Some 5)
check "body text chosen for a heading: Normal" (styleOf (toLevel 10 2 d0) 2, levels (toLevel 10 2 d0) |> List.item 2) ("Normal", 10)
let own level index doc = Outline.changeLevels (Outline.OwnLevel level) [ idAt doc index ] doc
check "OutlineLevel set on a built-in heading: no change (as Word)" (obj.ReferenceEquals((own 2 0 d0).Blocks, d0.Blocks)) true
check "OutlineLevel set on Normal: its own level, style Normal" (own 2 1 d0 |> fun d -> styleOf d 1, ownLevelOf d 1) ("Normal", Some 1)
check "OutlineLevel body text set on a style with a level: its own body level" (own 10 5 d0 |> fun d -> ownLevelOf d 5, levels d |> List.item 5) (Some 9, 10)
check "OutlineLevel body text set where only its own gave the level: none" (own 10 4 d0 |> fun d -> ownLevelOf d 4) None
let cellDoc = docOf [ h 1 "A"; table "in cell" ]
let cellId = match List.item 1 cellDoc.Blocks with Table t -> t.Rows.[0].Cells.[0].Content.Head.Id | _ -> ""
let cellPromoted = Outline.changeLevels Outline.Demote [ cellId ] cellDoc
check "a paragraph of a table cell (VBA) gets the style, its table stays a table" ((match List.item 1 cellPromoted.Blocks with Table t -> t.Rows.[0].Cells.[0].Content.Head.Metadata.Style | _ -> None), levels cellPromoted) (Some "Heading2", [ 1; 10 ])
check "a selection in a table cell has no body paragraph" (Outline.selectedBodyParagraphs (Doc.setCursor cellId 0 cellDoc)) []

// ---- a heading and its subheadings (the navigation pane) ---------------------------------------------------
let branch = Outline.changeBranch true (idAt d0 2) d0
check "a branch promoted: the heading and its subheadings" (levels branch) [ 1; 10; 1; 10; 2; 2; 10; 1 ]
check "a branch with a Heading 1 is not promoted" (obj.ReferenceEquals((Outline.changeBranch true (idAt d0 0) d0).Blocks, d0.Blocks)) true
check "a branch demoted" (levels (Outline.changeBranch false (idAt d0 2) d0)) [ 1; 10; 3; 10; 4; 4; 10; 1 ]

// ---- moves ------------------------------------------------------------------------------------------------
let three = docOf [ h 1 "A"; n "a body"; h 2 "A sub"; n "a sub body"; h 1 "B"; n "b body"; h 1 "C" ]
check "a heading moved down with its subtext past its next sibling" (texts (Outline.moveHeading false (idAt three 0) three)) [ "B"; "b body"; "A"; "a body"; "A sub"; "a sub body"; "C" ]
check "and up past its previous one" (texts (Outline.moveHeading true (idAt three 4) three)) [ "B"; "b body"; "A"; "a body"; "A sub"; "a sub body"; "C" ]
check "the last sibling moves up past the one before" (texts (Outline.moveHeading true (idAt three 6) three)) [ "A"; "a body"; "A sub"; "a sub body"; "C"; "B"; "b body" ]
check "a first child has no previous sibling: no move" (obj.ReferenceEquals((Outline.moveHeading true (idAt three 2) three).Blocks, three.Blocks)) true
check "an only child has no next sibling: no move" (obj.ReferenceEquals((Outline.moveHeading false (idAt three 2) three).Blocks, three.Blocks)) true
check "body text moves by one block" (texts (Outline.moveHeading false (idAt three 1) three)) [ "A"; "A sub"; "a body"; "a sub body"; "B"; "b body"; "C" ]
check "the caret stays where it was in the moved heading" ((Outline.moveHeading false (idAt three 0) (caretIn three 0)).Selection |> Option.map (fun s -> s.Focus.BlockId)) (Some (idAt three 0))
check "selected blocks move by one (Range.Relocate, Alt+Shift+Down: as Word)" (texts (Outline.moveSelectedBlocks false (caretIn three 0))) [ "a body"; "A"; "A sub"; "a sub body"; "B"; "b body"; "C" ]
check "a selection of two blocks moved up" (texts (Outline.moveSelectedBlocks true (selectFrom three 4 5))) [ "A"; "a body"; "A sub"; "B"; "b body"; "a sub body"; "C" ]
check "the first block does not move up" (obj.ReferenceEquals((Outline.moveSelectedBlocks true (caretIn three 0)).Blocks, three.Blocks)) true
check "blocks moved before another (a drop)" (texts (Outline.moveBlocks 4 6 0 three)) [ "B"; "b body"; "A"; "a body"; "A sub"; "a sub body"; "C" ]
check "blocks moved to the end" (texts (Outline.moveBlocks 0 2 7 three)) [ "A sub"; "a sub body"; "B"; "b body"; "C"; "A"; "a body" ]
check "blocks moved into their own place: no move" (obj.ReferenceEquals((Outline.moveBlocks 2 4 3 three).Blocks, three.Blocks)) true
let deleted = Outline.deleteHeading (idAt three 0) three
check "a heading deleted with its subtext" (texts deleted) [ "B"; "b body"; "C" ]
check "the caret where it stood" (deleted.Selection |> Option.map (fun s -> s.Focus.BlockId)) (Some (idAt three 4))
let onlyOne = docOf [ h 1 "only"; n "text" ]
let emptied = Outline.deleteHeading (idAt onlyOne 0) onlyOne
check "deleting everything leaves an empty paragraph" (texts emptied, emptied.Selection.IsSome) ([ "" ], true)
check "go to a heading: the caret at its start" ((Outline.goTo (idAt three 4) three).Selection |> Option.map (fun s -> s.Focus.BlockId, s.Focus.Offset)) (Some (idAt three 4, 0))

// ---- one undo step each, in the editing session ---------------------------------------------------------------
let session = Session.EditingSession(caretIn d0 1)
let run change = session.Execute(Session.ParagraphFormat change, None)
check "promote through the session changes the document" (run (Outline.changeSelectedLevels Outline.Promote)) (Ok Session.ContentChanged)
check "it is one undo step" (List.length session.History.Undos) 1
session.Undo() |> ignore
check "undone, the paragraph is body text again" (levels session.Document) (levels d0)
session.Document <- caretIn session.Document 0
check "a promote that changes nothing is no step" (run (Outline.changeSelectedLevels Outline.Promote)) (Ok Session.NoChange)
check "no step recorded" (List.length session.History.Undos) 0
check "a heading moved through the session" (run (fun d -> Outline.moveHeading true (idAt d 7) d)) (Ok Session.ContentChanged)
check "moved: Heading 1 C before A" (texts session.Document |> List.head) "C"
check "the move is one step" (List.length session.History.Undos) 1
session.Undo() |> ignore
check "the move undone in one step" (texts session.Document) (texts d0)

// ---- the save read again keeps it ------------------------------------------------------------------------------
let changed =
    caretIn d0 1
    |> Outline.changeSelectedLevels Outline.Demote
    |> fun d -> Outline.moveHeading true (idAt d 7) d
    |> fun d -> Outline.changeSelectedLevels (Outline.ToLevel 3) (caretIn d 2)
let roundTrip (doc: Document) =
    let package = Rbt.Word.Docx.exportTypedDocumentPackage doc
    let bytes = DocumentWriter.writeTypedPackageBytes { package with Unchanged = false }
    Rbt.Word.Docx.loadBytesNative bytes |> Option.defaultWith (fun () -> failwith "the saved DOCX did not load")
let reread = roundTrip changed
check "saved and read again: the blocks in their order" (texts reread) (texts changed)
check "saved and read again: the levels" (levels reread) (levels changed)
check "saved and read again: the heading styles" ([ for i in 0 .. 3 -> styleOf reread i ]) ([ for i in 0 .. 3 -> styleOf changed i ])
check "saved and read again: a paragraph's own level" (ownLevelOf reread 5) (ownLevelOf changed 5)

printfn "outline: %d passed, %d failed" pass fail
if fail > 0 then exit 1
