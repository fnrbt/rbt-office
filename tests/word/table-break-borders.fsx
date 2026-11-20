// The border between two rows of a table belongs to one of them: the upper
// row's bottom or the lower row's top, whichever won where the two meet. A
// page break between the rows left one side without its line: the first row
// on the next page had no top border (or the last row before the break no
// bottom border), and a header row repeated on the next page no line below it.
// Each side of a break has the line now; a row that follows its neighbour on
// the same page still has one line between them, not two.
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"

open Rbt.Word.Document
open Rbt.Word.Layout
module TB = Rbt.Word.TextBuffer
module Paint = Rbt.Word.Paint

let mutable pass = 0
let mutable fail = 0
let check label got expected =
    if got = expected then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-50s expected %A got %A" label expected got

let settings : LayoutSettings =
    { Width = 816.0; PageHeight = 1056.0; FontSize = 16.0; LineHeight = 24.0; PaddingLeft = 96.0; PaddingTop = 96.0
      MarginRight = 96.0; MarginBottom = 96.0; InTableCell = false; TableContinuationHeaderHeight = 0.0; LongTrailingSpaceFitSafety = 3.5; DefaultTabStop = 48.0; CompatibilityMode = 12; HangingPunctuation = Rbt.Word.Layout.NoHanging; SpaceShrink = 0.0 }
let measureText (text: string) (style: TB.TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = rangeMeasurerFromTextMeasurer measureText
let para id (text: string) : ParagraphBlock = { Id = id; Buffer = TB.ofString text; Metadata = defaultMetadata }
let pitch = settings.PageHeight + pageGap
let pageOf (y: float) = int (floor (y / pitch))

/// `rows` rows of two cells after `fillers` paragraphs; the first `headers` rows are header rows.
let build (fillers: int) (rows: int) (headers: int) =
    let position : Position = { BlockId = "p0"; Offset = 0 }
    let made = insertTable rows 2 { empty with Blocks = [ Paragraph (para "p0" "before") ]; Selection = Some { Anchor = position; Focus = position } }
    let table = made.Blocks |> List.pick (function Table table -> Some table | _ -> None)
    let table =
        { table with
            Rows =
                table.Rows |> List.mapi (fun r row ->
                    { row with
                        IsHeader = r < headers
                        CanSplit = Some false
                        Cells = row.Cells |> List.mapi (fun c cell -> { cell with Content = [ para (sprintf "r%d-%d" r c) (if r < headers then "HEAD" else sprintf "row %d" r) ] }) }) }
    { made with Blocks = [ for i in 1 .. fillers -> Paragraph (para (sprintf "f%d" i) "filler") ] @ [ Table table; Paragraph (para "after" "after") ]; Selection = None }

/// The horizontal border lines of a document, and the y of each row's first line by page.
let bordersOf (doc: Document) =
    let layouts, pages = layoutDocumentPages settings measureText measureRange doc
    let _, commands, _ = Paint.collectDocumentCommandsWithLayoutsAndPages settings "Calibri" measureText measureRange doc layouts pages
    let horizontals =
        commands
        |> List.choose (function Paint.StrokeLine (_, _, _, _, x1, y1, x2, y2) when abs (y1 - y2) < 0.1 && abs (x2 - x1) > 50.0 -> Some y1 | _ -> None)
        |> List.distinct
        |> List.sort
    let rowLines (r: int) = layouts |> List.filter (fun layout -> layout.BlockId = sprintf "r%d-0" r) |> List.collect (fun layout -> layout.Lines)
    layouts, horizontals, rowLines

/// A horizontal line within `reach` above the row's first line / below its last line.
let lineAbove (horizontals: float list) (line: MeasuredLine) = horizontals |> List.exists (fun y -> y <= line.Y + 0.5 && y >= line.Y - 14.0)
let lineBelow (horizontals: float list) (line: MeasuredLine) = horizontals |> List.exists (fun y -> y >= line.Y + line.Height - 0.5 && y <= line.Y + line.Height + 14.0)
let linesBetween (horizontals: float list) (upper: MeasuredLine) (lower: MeasuredLine) =
    horizontals |> List.filter (fun y -> y >= upper.Y + upper.Height - 0.5 && y <= lower.Y + 0.5) |> List.length

// ---- a table without header rows ----
do
    let rows = 40
    let _, horizontals, rowLines = bordersOf (build 20 rows 0)
    let firstOnSecond = [ 0 .. rows - 1 ] |> List.find (fun r -> pageOf (List.head (rowLines r)).Y = 1)
    check "the table goes on to a second page" (firstOnSecond > 1 && firstOnSecond < rows - 1) true
    let last = List.head (rowLines (firstOnSecond - 1))
    let first = List.head (rowLines firstOnSecond)
    check "the last row before the break has its bottom border" (lineBelow (horizontals |> List.filter (fun y -> pageOf y = 0)) last) true
    check "the first row after the break has its top border" (lineAbove (horizontals |> List.filter (fun y -> pageOf y = 1)) first) true
    check "nothing is drawn in the margin below the last row" (horizontals |> List.exists (fun y -> pageOf y = 0 && y > last.Y + last.Height + 14.0)) false
    // Rows that follow each other on a page share one line.
    let a, b = List.head (rowLines 1), List.head (rowLines 2)
    check "one line between two rows on the same page" (linesBetween horizontals a b) 1
    let c, d = List.head (rowLines firstOnSecond), List.head (rowLines (firstOnSecond + 1))
    check "also after the break" (linesBetween horizontals c d) 1
    check "the first row of the table has its top border, the last its bottom" (lineAbove horizontals (List.head (rowLines 0)), lineBelow horizontals (List.head (rowLines (rows - 1)))) (true, true)

// ---- header rows repeated on the next page ----
do
    let rows = 40
    let layouts, horizontals, rowLines = bordersOf (build 20 rows 2)
    let headLines (r: int) = layouts |> List.filter (fun layout -> layout.BlockId = sprintf "r%d-0" r) |> List.collect (fun layout -> layout.Lines)
    let firstOnSecond = [ 2 .. rows - 1 ] |> List.find (fun r -> pageOf (List.head (rowLines r)).Y = 1)
    let onSecond = horizontals |> List.filter (fun y -> pageOf y = 1)
    let first = List.head (rowLines firstOnSecond)
    // The header rows are painted again above the first row of the page: their
    // borders are lines above that row, two rows' worth.
    let aboveFirst = onSecond |> List.filter (fun y -> y <= first.Y + 0.5)
    check "the repeated header rows have their borders: top, between them, and below" (aboveFirst.Length >= 3) true
    check "the row below the repeated header rows has a line above it" (lineAbove onSecond first) true
    check "the last row before the break has its bottom border" (lineBelow (horizontals |> List.filter (fun y -> pageOf y = 0)) (List.head (rowLines (firstOnSecond - 1)))) true
    check "the header rows at the top of the table are bordered as before" (lineAbove horizontals (List.head (headLines 0)), linesBetween horizontals (List.head (headLines 0)) (List.head (headLines 1)), linesBetween horizontals (List.head (headLines 1)) (List.head (rowLines 2))) (true, 1, 1)

printfn "word table-break-borders: %d passed, %d failed" pass fail
if fail > 0 then exit 1
