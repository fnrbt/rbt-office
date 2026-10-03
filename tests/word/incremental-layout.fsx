#load "../../src/Word/Rope.fs"
#load "../../src/Word/TextBuffer.fs"
#load "../../src/Word/Document.fs"
#load "../../src/Word/Shapes.fs"
#load "../../src/Word/WordStyles.fs"
#load "../../src/Word/BorderGrid.fs"
#load "../../src/Word/WordLayout.fs"
#load "../../src/Word/Math.fs"
#load "../../src/Word/Layout.fs"

open Rbt.Word.TextBuffer
module Doc = Rbt.Word.Document
module Layout = Rbt.Word.Layout

let mutable pass = 0
let mutable fail = 0

let check label got expected =
    if got = expected then
        pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-42s expected %A got %A" label expected got

let checkTrue label value =
    check label value true

let settings : Layout.LayoutSettings =
    { Width = 420.0
      PageHeight = 100000.0
      FontSize = 16.0
      LineHeight = 24.0
      PaddingLeft = 0.0
      PaddingTop = 0.0
      MarginRight = 0.0
      MarginBottom = 0.0
      InTableCell = false
      TableContinuationHeaderHeight = 0.0
      LongTrailingSpaceFitSafety = 0.25 }

let pageSetup =
    { Doc.defaultPageSetup with
        PageWidth = settings.Width
        PageHeight = settings.PageHeight
        MarginLeft = settings.PaddingLeft
        MarginRight = settings.MarginRight
        MarginTop = settings.PaddingTop
        MarginBottom = settings.MarginBottom }

let measureText (text: string) (_style: TextStyle) =
    float text.Length * 7.0

let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText

let zeroSpacing =
    { Doc.defaultMetadata with SpacingBefore = Some 0.0; SpacingAfter = Some 0.0 }

let para id text : Doc.ParagraphBlock =
    { Id = id; Buffer = ofString text; Metadata = zeroSpacing }

let document blocks =
    { Doc.empty with Blocks = blocks; PageSetup = pageSetup; Styles = []; Selection = None }

let geometry (layouts: Layout.BlockLayout list) =
    layouts |> List.map (fun layout -> layout.BlockId, System.Math.Round(layout.Y, 6), System.Math.Round(layout.Height, 6), layout.Lines.Length)

let assertEquivalent label doc edit =
    let state = Layout.createIncrementalLayoutState()
    let cold0 = Layout.layoutDocumentWithRange settings measureText measureRange doc
    let inc0 = Layout.layoutDocumentIncrementalWithRange state settings measureText measureRange doc
    check (label + " initial geometry") (geometry inc0) (geometry cold0)
    let edited = edit doc
    Layout.resetIncrementalLayoutMeasurements state
    let cold1 = Layout.layoutDocumentWithRange settings measureText measureRange edited
    let inc1 = Layout.layoutDocumentIncrementalWithRange state settings measureText measureRange edited
    check (label + " edited geometry") (geometry inc1) (geometry cold1)
    check (label + " page count") (Layout.uniformPageInfos settings inc1).Length (Layout.uniformPageInfos settings cold1).Length
    Layout.incrementalLayoutMeasuredBlockCount state

let wrappedDoc =
    [ for i in 1 .. 200 ->
        let text = String.concat " " [ for j in 1 .. 18 -> sprintf "word%d_%d" i j ]
        let buffer =
            ofString text
            |> applyStyle 0 (min 12 text.Length) (fun style -> { style with Bold = i % 2 = 0; Italic = i % 3 = 0 })
        Doc.Paragraph { Id = sprintf "p%d" i; Buffer = buffer; Metadata = zeroSpacing } ]
    |> document

let editP100 doc =
    doc |> Doc.setCursor "p100" 4 |> Doc.insertText "X"

let measuredWrapped = assertEquivalent "wrapped mixed paragraphs" wrappedDoc editP100
checkTrue "wrapped edit remeasures <= 2 blocks" (measuredWrapped <= 2)

let tableCell text id : Doc.TableCell =
    { Content = [ para id text ]
      NestedTables = []
      NestedTableAnchors = []
      ColSpan = 1
      BackgroundColor = None
      Shading = None
      Width = None
      PreferredWidth = None
      BorderColor = None
      BorderWidth = None
      Borders = None
      VerticalAlignment = None
      VerticalMerge = None
      PaddingTop = None
      PaddingRight = None
      PaddingBottom = None
      PaddingLeft = None }

let tableRow a b : Doc.TableRow =
    { Cells = [ tableCell a (a + "-id"); tableCell b (b + "-id") ]
      IsHeader = false
      Height = None
      HeightRule = None
      CanSplit = None
      CellSpacing = None
      PaddingTop = None
      PaddingRight = None
      PaddingBottom = None
      PaddingLeft = None
      GridBefore = 0
      GridAfter = 0
      WidthBefore = None
      WidthAfter = None
      BordersOverride = None }

let inlineImageBuffer =
    ofString "image "
    |> insertImage 6 "memory://one" (Some 32.0) (Some 24.0) (Some "inline")
    |> insert 7 " after image text wraps"

let tableDoc =
    document
        [ Doc.Paragraph { (para "before" "before table paragraph with wrapped text") with Buffer = inlineImageBuffer }
          Doc.Table
            { Id = "table1"
              Rows = [ tableRow "r1c1" "r1c2"; tableRow "r2c1" "r2c2" ]
              PreferredWidth = None
              Alignment = None
              LayoutMode = None
              GridColumns = []
              Borders = None
              BackgroundColor = None
              Shading = None
              CellSpacing = None
              PaddingTop = None
              PaddingRight = None
              PaddingBottom = None
              PaddingLeft = None
              Floating = false }
          Doc.Paragraph (para "after" "after table paragraph") ]

let tableMeasured = assertEquivalent "table plus inline image" tableDoc (fun doc -> doc |> Doc.setCursor "after" 5 |> Doc.insertText "X")
checkTrue "table edit remeasures <= 2 blocks" (tableMeasured <= 2)

let footnoteBuffer =
    ofString "body text with note"
    |> setNote 15 1 (Some "1") (Some "footnote") (Some "footnote body text")

let footnoteDoc =
    document
        [ Doc.Paragraph { Id = "fn1"; Buffer = footnoteBuffer; Metadata = zeroSpacing }
          Doc.Paragraph (para "fn2" "paragraph after footnote") ]

let footnoteMeasured = assertEquivalent "footnote document" footnoteDoc (fun doc -> doc |> Doc.setCursor "fn2" 3 |> Doc.insertText "X")
checkTrue "footnote edit remeasures <= 2 blocks" (footnoteMeasured <= 2)

let columns : Doc.SectionColumns =
    { Count = 2; Gap = 24.0; Separator = true; Widths = [] }

let multiColumnDoc =
    { document [ for i in 1 .. 12 -> Doc.Paragraph (para (sprintf "mc%d" i) (String.concat " " [ for _ in 1 .. 14 -> "column text" ])) ] with
        PageSetup = { pageSetup with Columns = Some columns } }

let multiMeasured = assertEquivalent "multi-column document" multiColumnDoc (fun doc -> doc |> Doc.setCursor "mc6" 2 |> Doc.insertText "X")
check "multi-column cold fallback measurements" multiMeasured 0

printfn "incremental-layout: %d passed, %d failed" pass fail
if fail > 0 then failwithf "%d incremental layout checks failed" fail
