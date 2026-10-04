#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"

open Rbt.Word.TextBuffer
open Rbt.Word.Layout
open Rbt.Word.Paint
module Doc = Rbt.Word.Document
module TB = Rbt.Word.TextBuffer

// Model-level commands and paint behaviour the editor's toolbar and preview
// rely on: the placeholder label of an object the editor does not draw,
// content-control frames, paragraph spacing, page breaks, table rows /
// columns / merge / split, and hyperlinks.

let mutable pass = 0
let mutable fail = 0
let check label got expected =
    if got = expected then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-30s expected %A got %A" label expected got

let settings : LayoutSettings =
    { Width = 600.0; PageHeight = 1000.0; FontSize = 16.0; LineHeight = 24.0; PaddingLeft = 0.0; PaddingTop = 0.0
      MarginRight = 0.0; MarginBottom = 0.0; InTableCell = false; TableContinuationHeaderHeight = 0.0; LongTrailingSpaceFitSafety = 0.25 }
// Every character is 7 px wide.
let measureText (text: string) (_style: TextStyle) = float text.Length * 7.0
let measureRange = rangeMeasurerFromTextMeasurer measureText
let paragraph id (text: string) : Doc.ParagraphBlock = { Id = id; Buffer = TB.empty |> TB.insert 0 text; Metadata = Doc.defaultMetadata }
let textOf (para: Doc.ParagraphBlock) = TB.toString para.Buffer
let docOf (paragraphs: Doc.ParagraphBlock list) = { Doc.empty with Blocks = paragraphs |> List.map Doc.Paragraph; Selection = None }
let caret id offset (doc: Doc.Document) = Doc.setCursor id offset doc
let select id first last (doc: Doc.Document) = Doc.setSelection { BlockId = id; Offset = first } { BlockId = id; Offset = last } doc

// ---- the placeholder label: its own width, in the middle of the frame ------
let opaqueDoc (width: float) (height: float) =
    let buffer = TB.empty |> TB.insertImage 0 (TB.opaqueSourcePrefix + "chart") (Some width) (Some height) None
    docOf [ { Id = "opaque"; Buffer = buffer; Metadata = Doc.defaultMetadata } ]
let labelOf (width: float) (height: float) =
    let _, commands, _ = collectDocumentCommands settings "Calibri" measureText measureRange (opaqueDoc width height)
    let frame = commands |> List.pick (function StrokeRect (_, _, x, y, w, h) -> Some (x, y, w, h) | _ -> None)
    let label = commands |> List.tryPick (function DrawText (text, _, x, baseline, w) -> Some (text, x, baseline, w) | _ -> None)
    frame, label
let (frameX, frameY, frameW, frameH), wide = labelOf 400.0 200.0
check "label: the text" (wide |> Option.map (fun (text, _, _, _) -> text)) (Some "Chart")
check "label: drawn at its own width (5 characters), not the frame's" (wide |> Option.map (fun (_, _, _, width) -> width)) (Some 35.0)
check "label: centred horizontally" (wide |> Option.map (fun (_, x, _, width) -> x + width / 2.0)) (Some (frameX + frameW / 2.0))
check "label: its baseline is around the middle of the frame" (wide |> Option.map (fun (_, _, baseline, _) -> baseline > frameY + frameH / 2.0 - 2.0 && baseline < frameY + frameH / 2.0 + 12.0)) (Some true)
let _, narrow = labelOf 40.0 60.0
check "label: a narrow frame shows what fits, with an ellipsis" (narrow |> Option.map (fun (text, _, _, width) -> text.EndsWith "…" && text.Length < 6 && width <= 28.0)) (Some true)
let _, tiny = labelOf 12.0 60.0
check "label: nothing in a frame too narrow for any of it" tiny None
let _, flat = labelOf 400.0 14.0
check "label: nothing in a frame too low for it" flat None

// ---- content-control frames ------------------------------------------------------
let controlDoc =
    docOf [ paragraph "before" "Before."; paragraph "in1" "Inside one."; paragraph "in2" "Inside two."; paragraph "after" "After." ]
    |> Doc.setSelection { BlockId = "in1"; Offset = 0 } { BlockId = "in2"; Offset = 3 }
    |> Doc.setContentControl "cc" "tag" "alias"
let controlLayouts (doc: Doc.Document) = layoutDocument settings measureText doc
let frames (doc: Doc.Document) =
    markupMarkerCommands settings doc (controlLayouts doc) |> List.choose (function StrokeRect (_, _, _, y, _, h) -> Some (y, h) | _ -> None)
let bodyHasControlBlue (doc: Doc.Document) =
    let _, commands, _ = collectDocumentCommands settings "Calibri" measureText measureRange doc
    commands |> List.exists (function StrokeRect ("#2563EB", _, _, _, _, _) | FillRect ("#2563EB", _, _, _, _) -> true | _ -> false)
check "control frame: not part of the page" (bodyHasControlBlue controlDoc) false
check "control frame: none without a selection" (frames { controlDoc with Selection = None }) []
check "control frame: none with the caret outside the control" (frames (caret "before" 2 controlDoc), frames (caret "after" 1 controlDoc)) ([], [])
check "control frame: around both paragraphs of the control the caret is in" (frames (caret "in1" 3 controlDoc) |> List.length, frames (caret "in2" 0 controlDoc) |> List.length) (2, 2)
check "control frame: a selection ending in the control shows it" (frames (Doc.setSelection { BlockId = "before"; Offset = 1 } { BlockId = "in1"; Offset = 2 } controlDoc) |> List.length) 2
let twoControls =
    controlDoc |> caret "after" 0 |> Doc.setContentControl "other" "tag2" "alias2"
check "control frame: only the caret's control" (frames (caret "after" 2 twoControls) |> List.length, frames (caret "in1" 2 twoControls) |> List.length) (1, 2)

// ---- paragraph spacing -----------------------------------------------------------
let spaced =
    docOf [ paragraph "p1" "One"; paragraph "p2" "Two"; paragraph "p3" "Three" ]
    |> Doc.setSelection { BlockId = "p1"; Offset = 1 } { BlockId = "p2"; Offset = 1 }
let metadataOf id (doc: Doc.Document) = (Doc.findParagraph doc id).Value.Metadata
let afterSpacing = spaced |> Doc.setLineSpacing (Some 1.5) |> Doc.setSpacingBefore (Some 12.0) |> Doc.setSpacingAfter (Some 6.0)
check "spacing: on every selected paragraph" ([ "p1"; "p2" ] |> List.map (fun id -> let m = metadataOf id afterSpacing in m.LineSpacing, m.LineSpacingRule, m.SpacingBefore, m.SpacingAfter)) [ (Some 1.5, Some "auto", Some 12.0, Some 6.0); (Some 1.5, Some "auto", Some 12.0, Some 6.0) ]
check "spacing: not on the others" (metadataOf "p3" afterSpacing = Doc.defaultMetadata) true
check "spacing: stated on the paragraph (saved with it)" (Doc.ParagraphExplicit.has Doc.ParagraphExplicit.LineSpacing (metadataOf "p1" afterSpacing), Doc.ParagraphExplicit.has Doc.ParagraphExplicit.SpacingBefore (metadataOf "p1" afterSpacing), Doc.ParagraphExplicit.has Doc.ParagraphExplicit.SpacingAfter (metadataOf "p1" afterSpacing)) (true, true, true)
let cleared = afterSpacing |> Doc.setLineSpacing None |> Doc.setSpacingBefore None |> Doc.setSpacingAfter None
check "spacing: None goes back to the style's" (metadataOf "p1" cleared = Doc.defaultMetadata) true

// ---- page break ------------------------------------------------------------------
let broken = docOf [ paragraph "p1" "Hello world"; paragraph "p2" "Next" ] |> caret "p1" 6 |> Doc.insertPageBreak
let brokenParagraphs = Doc.allParagraphs broken
check "page break: the paragraph is split at the caret" (brokenParagraphs |> List.map textOf) [ "Hello "; "world"; "Next" ]
check "page break: what follows starts a new page" (brokenParagraphs |> List.map (fun para -> para.Metadata.PageBreakBefore)) [ false; true; false ]
check "page break: the caret is at the start of the new page" (broken.Selection |> Option.map (fun s -> s.Focus.BlockId = brokenParagraphs.[1].Id, s.Focus.Offset)) (Some (true, 0))
let brokenLayouts, brokenPages = layoutDocumentPages settings measureText measureRange broken
check "page break: the layout puts it on the second page" (brokenPages.Length, (brokenLayouts |> List.find (fun layout -> layout.BlockId = brokenParagraphs.[1].Id)).Y >= 1000.0) (2, true)

// ---- tables ----------------------------------------------------------------------
let withTable = docOf [ paragraph "p1" "Text" ] |> caret "p1" 4 |> Doc.insertTable 2 3
let tableOf (doc: Doc.Document) = doc.Blocks |> List.pick (function Doc.Table table -> Some table | _ -> None)
let shape (doc: Doc.Document) = (tableOf doc).Rows |> List.map (fun row -> row.Cells |> List.map (fun cell -> cell.ColSpan))
let cellId row col (doc: Doc.Document) = (tableOf doc).Rows.[row].Cells.[col].Content.Head.Id
let cellText row col (doc: Doc.Document) = (tableOf doc).Rows.[row].Cells.[col].Content |> List.map textOf |> String.concat "|"
check "table: 2 x 3" (shape withTable) [ [ 1; 1; 1 ]; [ 1; 1; 1 ] ]
let filled =
    [ 0, 0, "a"; 0, 1, "b"; 0, 2, "c"; 1, 0, "d"; 1, 1, "e"; 1, 2, "f" ]
    |> List.fold (fun doc (row, col, text) -> doc |> caret (cellId row col doc) 0 |> Doc.insertText text) withTable
let inB = filled |> caret (cellId 0 1 filled) 0
let rowBefore = Doc.insertTableRowBeforeSelection inB
check "table: a row inserted before" (shape rowBefore, cellText 0 1 rowBefore, cellText 1 1 rowBefore) ([ [ 1; 1; 1 ]; [ 1; 1; 1 ]; [ 1; 1; 1 ] ], "", "b")
let rowAfter = Doc.insertTableRowAfterSelection inB
check "table: a row inserted after" (cellText 0 1 rowAfter, cellText 1 1 rowAfter, cellText 2 1 rowAfter) ("b", "", "e")
let colBefore = Doc.insertTableColumnBeforeSelection inB
check "table: a column inserted before" ((tableOf colBefore).Rows |> List.map (fun row -> row.Cells |> List.map (fun cell -> cell.Content |> List.map textOf |> String.concat ""))) [ [ "a"; ""; "b"; "c" ]; [ "d"; ""; "e"; "f" ] ]
let colAfter = Doc.insertTableColumnAfterSelection inB
check "table: a column inserted after" ((tableOf colAfter).Rows.[0].Cells |> List.map (fun cell -> cell.Content |> List.map textOf |> String.concat "")) [ "a"; "b"; ""; "c" ]
// Merge a..b of the first row.
check "table: a caret in one cell has nothing to merge" (Doc.canMergeSelectedTableCells inB) false
let rowSelection = filled |> Doc.setSelection { BlockId = cellId 0 0 filled; Offset = 0 } { BlockId = cellId 0 1 filled; Offset = 1 }
check "table: a selection over two cells can be merged" (Doc.canMergeSelectedTableCells rowSelection) true
let mergedRow = Doc.mergeSelectedTableCells rowSelection
check "table: merged across columns" (shape mergedRow, cellText 0 0 mergedRow, cellText 0 1 mergedRow) ([ [ 2; 1 ]; [ 1; 1; 1 ] ], "a|b", "c")
check "table: the caret is in the merged cell" (mergedRow.Selection |> Option.map (fun s -> s.Focus.BlockId)) (Some (cellId 0 0 mergedRow))
let splitRow = mergedRow |> Doc.splitSelectedTableCell
check "table: split back into single cells" (shape splitRow, cellText 0 0 splitRow, cellText 0 1 splitRow, cellText 0 2 splitRow) ([ [ 1; 1; 1 ]; [ 1; 1; 1 ] ], "a|b", "", "c")
// Merge b and e (one column, two rows).
let columnSelection = filled |> Doc.setSelection { BlockId = cellId 0 1 filled; Offset = 0 } { BlockId = cellId 1 1 filled; Offset = 0 }
let mergedColumn = Doc.mergeSelectedTableCells columnSelection
check "table: merged across rows" ((tableOf mergedColumn).Rows |> List.map (fun row -> row.Cells.[1].VerticalMerge), cellText 0 1 mergedColumn, cellText 1 1 mergedColumn) ([ Some Doc.RestartMerge; Some Doc.ContinueMerge ], "b|e", "")
let splitColumn = mergedColumn |> caret (cellId 0 1 mergedColumn) 0 |> Doc.splitSelectedTableCell
check "table: a vertical merge split" ((tableOf splitColumn).Rows |> List.map (fun row -> row.Cells.[1].VerticalMerge)) [ None; None ]
// A block of four cells.
let blockSelection = filled |> Doc.setSelection { BlockId = cellId 0 1 filled; Offset = 0 } { BlockId = cellId 1 2 filled; Offset = 0 }
let mergedBlock = Doc.mergeSelectedTableCells blockSelection
check "table: a block of cells merged" (shape mergedBlock, cellText 0 1 mergedBlock, (tableOf mergedBlock).Rows |> List.map (fun row -> row.Cells.[1].VerticalMerge)) ([ [ 1; 2 ]; [ 1; 2 ] ], "b|c|e|f", [ Some Doc.RestartMerge; Some Doc.ContinueMerge ])
check "table: text outside a table is not merged" (Doc.canMergeSelectedTableCells (docOf [ paragraph "x" "one"; paragraph "y" "two" ] |> Doc.setSelection { BlockId = "x"; Offset = 0 } { BlockId = "y"; Offset = 1 })) false

// ---- hyperlinks ------------------------------------------------------------------
let plain = docOf [ paragraph "p" "Visit the site today." ]
let linked = plain |> select "p" 10 14 |> Doc.insertHyperlink "https://example.com/a" ""
let paraOf (doc: Doc.Document) = (Doc.findParagraph doc "p").Value
check "link: the selection becomes the link" (Doc.hyperlinkAt (paraOf linked) 12) (Some ("https://example.com/a", 10, 14))
check "link: found from its first and its last character" (Doc.hyperlinkAt (paraOf linked) 10, Doc.hyperlinkAt (paraOf linked) 14) (Some ("https://example.com/a", 10, 14), Some ("https://example.com/a", 10, 14))
check "link: none elsewhere" (Doc.hyperlinkAt (paraOf linked) 3, Doc.hyperlinkAt (paraOf linked) 16) (None, None)
let retargeted = linked |> caret "p" 12 |> Doc.setHyperlinkTarget (Some "https://example.com/b")
check "link: a new target for the whole link" (Doc.hyperlinkAt (paraOf retargeted) 11, textOf (paraOf retargeted)) (Some ("https://example.com/b", 10, 14), "Visit the site today.")
let unlinked = retargeted |> caret "p" 12 |> Doc.setHyperlinkTarget None
check "link: removed, the text stays" (Doc.hyperlinkAt (paraOf unlinked) 12, textOf (paraOf unlinked)) (None, "Visit the site today.")
let inserted = plain |> caret "p" 6 |> Doc.insertHyperlink "https://example.com/c" "our page "
check "link: inserted at the caret with its text" (textOf (paraOf inserted), Doc.hyperlinkAt (paraOf inserted) 8) ("Visit our page the site today.", Some ("https://example.com/c", 6, 15))
check "link: the caret is after it" (inserted.Selection |> Option.map (fun s -> s.Focus.Offset, s.IsCollapsed)) (Some (15, true))
let bare = plain |> caret "p" 0 |> Doc.insertHyperlink "https://example.com/d" ""
check "link: without a text the target is the text" ((textOf (paraOf bare)).StartsWith "https://example.com/dVisit", Doc.hyperlinkAt (paraOf bare) 3 |> Option.map (fun (target, first, last) -> target, first, last)) (true, Some ("https://example.com/d", 0, 21))

// ---- a table that was just inserted; what the caret is in --------------------------
module Styles = Rbt.Word.WordStyles
let newWithTable = docOf [ paragraph "before" "Before" ] |> Doc.setCursor "before" 3 |> Doc.insertTable 3 3
let newInsertedTable = newWithTable.Blocks |> List.pick (function Doc.Table table -> Some table | _ -> None)
let newCellParagraph row col = newInsertedTable.Rows.[row].Cells.[col].Content.Head
let newSide (border: Doc.TableBorder option) = border |> Option.map (fun b -> b.Style, b.Color, b.Width)
check "new table: single borders on every edge"
    (newInsertedTable.Borders |> Option.map (fun b -> [ newSide b.Top; newSide b.Right; newSide b.Bottom; newSide b.Left; newSide b.InsideH; newSide b.InsideV ] |> List.distinct))
    (Some [ Some (Some "single", Some "#000000", Some 1.0) ])
check "new table: no table style in a document without one" newInsertedTable.StyleId None
check "new table: the caret is in its first cell" (newWithTable.Selection |> Option.map (fun sel -> sel.Focus.BlockId, sel.Focus.Offset, sel.IsCollapsed)) (Some ((newCellParagraph 0 0).Id, 0, true))
let newContextOf (doc: Doc.Document) = let c = Doc.caretContext doc in c.InTable, c.Row, c.Column, c.RowCount, c.ColumnCount
check "caret context: the first cell" (newContextOf newWithTable) (true, 0, 0, 3, 3)
check "caret context: outside a table" (newContextOf (newWithTable |> Doc.setCursor "before" 1)) (false, -1, -1, 0, 0)
let tab (doc: Doc.Document) = Doc.moveToAdjacentTableCell true doc |> Option.get
let backTab (doc: Doc.Document) = Doc.moveToAdjacentTableCell false doc
check "Tab: the next cell" (newContextOf (tab newWithTable)) (true, 0, 1, 3, 3)
check "Tab at the end of a row: the next row" (newContextOf (newWithTable |> tab |> tab |> tab)) (true, 1, 0, 3, 3)
check "Shift+Tab: the cell before" (newWithTable |> tab |> tab |> tab |> backTab |> Option.map newContextOf) (Some (true, 0, 2, 3, 3))
check "Shift+Tab in the first cell: nowhere" (backTab newWithTable |> Option.isSome) false
check "Tab outside a table: not a cell move" (Doc.moveToAdjacentTableCell true (newWithTable |> Doc.setCursor "before" 1) |> Option.isSome) false
let newAtLast = [ 1 .. 8 ] |> List.fold (fun doc _ -> tab doc) newWithTable
check "eight Tabs: the last cell" (newContextOf newAtLast) (true, 2, 2, 3, 3)
check "Tab in the last cell: a new row, its first cell" (newContextOf (tab newAtLast)) (true, 3, 0, 4, 3)
// The layout has a place for the empty paragraph of every cell (the caret is drawn there).
let newTableLayouts = layoutDocumentWithRange settings measureText measureRange newWithTable
check "layout: every empty cell paragraph has a place"
    ([ for row in 0 .. 2 do for col in 0 .. 2 -> newTableLayouts |> List.exists (fun layout -> layout.BlockId = (newCellParagraph row col).Id) ] |> List.forall id) true
let newCellLayout row col = newTableLayouts |> List.find (fun layout -> layout.BlockId = (newCellParagraph row col).Id)
check "layout: cells of a row newSide by newSide, rows below each other"
    ((newCellLayout 0 1).Indent > (newCellLayout 0 0).Indent + 50.0, (newCellLayout 0 2).Indent > (newCellLayout 0 1).Indent + 50.0, (newCellLayout 1 0).Y > (newCellLayout 0 0).Y + 10.0, (newCellLayout 0 1).Y = (newCellLayout 0 0).Y)
    (true, true, true, true)
// Merge / split and what the context says of them.
let newOverTwo = newWithTable |> Doc.setSelection { BlockId = (newCellParagraph 0 0).Id; Offset = 0 } { BlockId = (newCellParagraph 0 1).Id; Offset = 0 }
check "caret context: a selection over two cells can be merged" (let c = Doc.caretContext newOverTwo in c.CanMergeCells, c.CanSplitCell) (true, false)
let newMergedCells = Doc.mergeSelectedTableCells newOverTwo
let newMergedTable = newMergedCells.Blocks |> List.pick (function Doc.Table table -> Some table | _ -> None)
let newInMerged = newMergedCells |> Doc.setCursor newMergedTable.Rows.Head.Cells.Head.Content.Head.Id 0
check "caret context: a merged cell can be split" (let c = Doc.caretContext newInMerged in c.InTable, c.CanSplitCell, c.ColumnCount) (true, true, 3)
// Lists, links.
let newListed = docOf [ { paragraph "li" "item" with Metadata = { Doc.defaultMetadata with ListStyle = Doc.Bullet } } ] |> Doc.setCursor "li" 2
check "caret context: in a list" (let c = Doc.caretContext newListed in c.InList, c.InTable, c.InHyperlink, c.InHeaderFooter, c.ImageSelected) (true, false, false, false, false)
check "caret context: in a hyperlink" ((Doc.caretContext (bare |> Doc.setCursor "p" 3)).InHyperlink) true
// The selection format: an internal style entry is not a style; spacing is reported.
let internalStyled = docOf [ { paragraph "cellp" "text" with Metadata = { Doc.defaultMetadata with Style = Some "__docDefaults" } } ] |> Doc.setCursor "cellp" 1
check "selection format: an internal style id is no style" (Styles.selectionFormat internalStyled).StyleId (Styles.Uniform None)
check "internal style ids" (Styles.isInternalStyleId "__docDefaults", Styles.isInternalStyleId "Heading1", Styles.isInternalStyleId null) (true, false, false)
let newSpaced = docOf [ paragraph "sp" "text" ] |> Doc.setCursor "sp" 1 |> Doc.setLineSpacing (Some 1.5) |> Doc.setSpacingBefore (Some 8.0) |> Doc.setSpacingAfter (Some 16.0)
let spacedFormat = Styles.selectionFormat newSpaced
check "selection format: line spacing and the space before / after"
    (spacedFormat.ParagraphLineSpacing, spacedFormat.ParagraphSpacingBefore, spacedFormat.ParagraphSpacingAfter)
    (Styles.Uniform (Some 1.5), Styles.Uniform 8.0, Styles.Uniform 16.0)

printfn ""
printfn "Word command tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
