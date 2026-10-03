#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Vb.dll"

open Rbt.Word.TextBuffer
open System.Reflection
module Doc = Rbt.Word.Document
module Cmd = Rbt.Word.Commands
module Vba = Rbt.Word.Vba

let mutable pass = 0
let mutable fail = 0

let check label got expected =
    if got = expected then
        pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-28s expected %A got %A" label expected got

let repeatedAppend =
    let mutable buffer = empty
    for _ in 1 .. 1000 do
        buffer <- insert (length buffer) "x" buffer
    buffer

check "append text length" (length repeatedAppend) 1000
check "append text content" (toString repeatedAppend) (System.String('x', 1000))
check "append merges runs" repeatedAppend.Runs.Length 1
check "single run style lookup" (getStyleAt 999 repeatedAppend).Bold false
check "style before buffer default" (getStyleAt -1 repeatedAppend).Bold false
check "style after buffer default" (getStyleAt 1000 repeatedAppend).Bold false

let styleLookupA =
    { Rope = Rbt.Word.Rope.ofString "ab"
      Runs =
        [ Text { Start = 0; Length = 1; Style = { defaultStyle with Bold = true } }
          Text { Start = 1; Length = 1; Style = defaultStyle } ] }
let styleLookupB =
    { Rope = Rbt.Word.Rope.ofString "ab"
      Runs =
        [ Text { Start = 0; Length = 1; Style = defaultStyle }
          Text { Start = 1; Length = 1; Style = { defaultStyle with Italic = true } } ] }
check "style lookup fragmented first" (getStyleAt 0 styleLookupA).Bold true
check "style lookup switches" (getStyleAt 1 styleLookupB).Italic true
check "style lookup returns" (getStyleAt 0 styleLookupA).Bold true
check "style lookup keeps switched buffer" (getStyleAt 1 styleLookupB).Italic true
// removed: tested deleted cache/perf machinery

let middleInsert =
    let mutable buffer = ofString "abcd"
    buffer <- insert 2 "XY" buffer
    check "middle insert text" (toString buffer) "abXYcd"
    check "middle insert merges runs" buffer.Runs.Length 1

let styleBoundary =
    let styled =
        ofString "abcd"
        |> applyStyle 0 2 (fun style -> { style with Bold = true })
    let inserted = insert 2 "X" styled
    check "style boundary text" (toString inserted) "abXcd"
    check "style boundary run count" inserted.Runs.Length 2

let fullStyle =
    ofString "abcd"
    |> applyStyle 0 4 (fun style -> { style with Bold = true })
check "style full text" (toString fullStyle) "abcd"
check "style full run count" fullStyle.Runs.Length 1
check "style full bold" (getStyleAt 0 fullStyle).Bold true

let fragmented =
    { Rope = Rbt.Word.Rope.ofString "abcd"
      Runs =
        [ for i in 0 .. 3 ->
            Text { Start = i; Length = 1; Style = defaultStyle } ] }
let fragmentedEndInsert =
    let boldLast =
        { fragmented with
            Runs =
                [ Text { Start = 0; Length = 3; Style = defaultStyle }
                  Text { Start = 3; Length = 1; Style = { defaultStyle with Bold = true } } ] }
    insert 4 "z" boldLast
check "insert fragmented end text" (toString fragmentedEndInsert) "abcdz"
check "insert fragmented end style" (getStyleAt 4 fragmentedEndInsert).Bold true
let deletedFragment = delete 0 1 fragmented
check "delete fragmented text" (toString deletedFragment) "bcd"
check "delete fragmented merges" deletedFragment.Runs.Length 1
let deleteAfterFragment = delete 99 1 fragmented
check "delete fragmented after text" (toString deleteAfterFragment) "abcd"
let noopStyledFragment = applyStyle 1 2 id fragmented
check "style fragmented noop text" (toString noopStyledFragment) "abcd"

let bridgeDelete =
    let bold = { defaultStyle with Bold = true }
    { Rope = Rbt.Word.Rope.ofString "abXcd"
      Runs =
        [ Text { Start = 0; Length = 2; Style = bold }
          Text { Start = 2; Length = 1; Style = defaultStyle }
          Text { Start = 3; Length = 2; Style = bold } ] }
    |> delete 2 1
check "delete bridge text" (toString bridgeDelete) "abcd"
check "delete bridge merges" bridgeDelete.Runs.Length 1

let largeText = String.init 5000 (fun i -> string (char (int 'a' + (i % 26))))
let largeRope = Rbt.Word.Rope.ofString largeText
check "rope slice all" (Rbt.Word.Rope.slice 0 largeText.Length largeRope) largeText
check "rope slice first half" (Rbt.Word.Rope.slice 0 2500 largeRope) (largeText.Substring(0, 2500))
check "rope slice second half" (Rbt.Word.Rope.slice 2500 2500 largeRope) (largeText.Substring(2500, 2500))
check "rope slice cross leaf" (Rbt.Word.Rope.slice 1010 80 largeRope) (largeText.Substring(1010, 80))
check "rope char first" (Rbt.Word.Rope.charAt 0 largeRope) (Some largeText.[0])
check "rope char middle" (Rbt.Word.Rope.charAt 2517 largeRope) (Some largeText.[2517])
check "rope char last" (Rbt.Word.Rope.charAt (largeText.Length - 1) largeRope) (Some largeText.[largeText.Length - 1])
check "rope char before" (Rbt.Word.Rope.charAt -1 largeRope) None
check "rope char after" (Rbt.Word.Rope.charAt largeText.Length largeRope) None
check "rope linecol same leaf" (Rbt.Word.Rope.lineColToPosition 0 10 largeRope) 10
check "rope linecol later leaf" (Rbt.Word.Rope.lineColToPosition 0 2500 largeRope) 2500
check "rope pos linecol later leaf" (Rbt.Word.Rope.positionToLineCol 2500 largeRope) (0, 2500)
let leafSource = System.String([| 'l'; 'e'; 'a'; 'f' |])

let multiLineText =
    String.init 5000 (fun i ->
        if i % 100 = 99 then "\n"
        else string (char (int 'a' + (i % 26))))
let multiLineRope = Rbt.Word.Rope.ofString multiLineText
check "rope linecol multiline" (Rbt.Word.Rope.lineColToPosition 20 17 multiLineRope) 2017
check "rope pos linecol multiline" (Rbt.Word.Rope.positionToLineCol 2017 multiLineRope) (20, 17)
check "rope linecol multiline tail" (Rbt.Word.Rope.lineColToPosition 49 99 multiLineRope) 4999
check "rope pos linecol multiline tail" (Rbt.Word.Rope.positionToLineCol 4999 multiLineRope) (49, 99)
check "rope linecol after final newline" (Rbt.Word.Rope.lineColToPosition 50 0 multiLineRope) 5000
let longParagraphText = System.String('x', 4096)
let longTextDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph { Id = "long"; Buffer = ofString longParagraphText; Metadata = Doc.defaultMetadata }
              Doc.Paragraph { Id = "tail"; Buffer = ofString "done"; Metadata = Doc.defaultMetadata } ] }
check "get text long paragraph" (Doc.getText longTextDoc) (longParagraphText + "\ndone")
let changedLongTextDoc =
    { longTextDoc with Blocks = [ Doc.Paragraph { Id = "changed"; Buffer = ofString "changed"; Metadata = Doc.defaultMetadata } ] }
check "get text changed blocks" (Doc.getText changedLongTextDoc) "changed"
// removed: tested deleted cache/perf machinery
let plainContentDoc = Doc.setPlainTextContent "plain" Doc.empty
check "plain content text" (Doc.getText plainContentDoc) "plain"
let styledPlainContentDoc =
    { plainContentDoc with
        Blocks = [ Doc.Paragraph { Id = "block-0"; Buffer = ofString "plain" |> setColor 0 5 (Some "#ff0000"); Metadata = Doc.defaultMetadata } ] }
let resetPlainContentDoc = Doc.setPlainTextContent "plain" styledPlainContentDoc
match resetPlainContentDoc.Blocks with
| [ Doc.Paragraph para ] -> check "plain content preserves style" (getStyleAt 0 para.Buffer).Color (Some "#ff0000")
| _ ->
    fail <- fail + 1
    printfn "FAIL %-28s expected plain paragraph" "plain content resets"
check "nonbreaking blank empty" (isNonbreakingBlank empty) false
check "nonbreaking blank spaces" (isNonbreakingBlank (ofString " \t\n")) true
check "nonbreaking blank nbsp" (isNonbreakingBlank (ofString "\u00A0\u00A0")) true
check "nonbreaking blank text" (isNonbreakingBlank (ofString "\u00A0x")) false

let red = { defaultStyle with Color = Some "#ff0000" }
let styledBuffer =
    { Rope = Rbt.Word.Rope.ofString "abcd"
      Runs =
        [ Text { Start = 0; Length = 2; Style = { defaultStyle with Bold = true } }
          Text { Start = 2; Length = 2; Style = red } ] }
let slicedBuffer = sliceBuffer 1 2 styledBuffer
check "sliceBuffer text" (toString slicedBuffer) "bc"
check "sliceBuffer first style" (getStyleAt 0 slicedBuffer).Bold true
check "sliceBuffer second style" (getStyleAt 1 slicedBuffer).Color (Some "#ff0000")
let fullSlicedBuffer = sliceBuffer 0 (length styledBuffer) styledBuffer
let clampedFullSlicedBuffer = sliceBuffer -5 (length styledBuffer + 10) styledBuffer
let appendedBuffer = append (sliceBuffer 0 2 styledBuffer) (sliceBuffer 2 2 styledBuffer)
check "append buffer text" (toString appendedBuffer) "abcd"
check "append preserves left style" (getStyleAt 1 appendedBuffer).Bold true
check "append preserves right style" (getStyleAt 2 appendedBuffer).Color (Some "#ff0000")
let sameStyleAppend = append (ofString "ab") (ofString "cd")
check "append same style text" (toString sameStyleAppend) "abcd"
check "append same style merges" sameStyleAppend.Runs.Length 1

let para id buffer : Doc.ParagraphBlock = { Id = id; Buffer = buffer; Metadata = Doc.defaultMetadata }
let replaceHelperPara = para "replace-helper" (ofString "same")
let replaceHelperDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph replaceHelperPara
              Doc.Paragraph (para "replace-helper-tail" (ofString "tail")) ] }
let tableCell id text : Doc.TableCell =
    { Content = [para id (ofString text)]
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
let singleCellTable tableId cell : Doc.TableBlock =
    { Id = tableId
      Rows =
        [{ Cells = [cell]
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
           BordersOverride = None }]
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
let singleColumnTableRow cell : Doc.TableRow =
    { Cells = [ cell ]
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
let singleColumnTableDoc prefix docIndex rowCount text selection =
    let tableId = sprintf "%s-table-%02d" prefix docIndex
    let rows =
        [ for rowIndex in 0 .. rowCount - 1 ->
            singleColumnTableRow (tableCell (sprintf "%s-%02d-%05d" prefix docIndex rowIndex) text) ]
    let table = { (singleCellTable tableId (tableCell (sprintf "%s-template" prefix) "")) with Rows = rows }
    { Doc.empty with Blocks = [ Doc.Table table ]; Selection = selection }
let duplicateParagraphDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph (para "dup" (ofString "first"))
              Doc.Paragraph (para "dup" (ofString "second")) ] }
check "find block first duplicate" (Doc.findBlock duplicateParagraphDoc "dup" |> Option.map (fun b -> Doc.getParagraph b |> fun p -> toString p.Buffer)) (Some "first")
check "find block index duplicate" (Doc.findBlockIndex duplicateParagraphDoc "dup") (Some 0)
check "find paragraph first duplicate" (Doc.findParagraph duplicateParagraphDoc "dup" |> Option.map (fun p -> toString p.Buffer)) (Some "first")
let replacedDuplicateParagraphDoc =
    Doc.replaceBlock duplicateParagraphDoc "dup" (Doc.Paragraph (para "dup" (ofString "changed")))
check "replace block all duplicates"
    (replacedDuplicateParagraphDoc.Blocks
     |> List.choose (function
         | Doc.Paragraph p -> Some (toString p.Buffer)
         | _ -> None))
    [ "changed"; "changed" ]
let invalidatedParagraphDoc =
    { duplicateParagraphDoc with
        Blocks = [ Doc.Paragraph (para "other" (ofString "other")) ] }
check "find block changed blocks" (Doc.findBlock invalidatedParagraphDoc "dup") None
check "find block index invalidates" (Doc.findBlockIndex invalidatedParagraphDoc "dup") None
check "find paragraph changed blocks" (Doc.findParagraph invalidatedParagraphDoc "dup") None

// removed: tested deleted cache/perf machinery
let smallBulkReplaceTableRowCount = 20
let smallBulkReplaceTableEnd = sprintf "bulk-replace-table-small-00-%05d" (smallBulkReplaceTableRowCount - 1)
let smallBulkReplaceTableDoc =
    { singleColumnTableDoc "bulk-replace-table-small" 0 smallBulkReplaceTableRowCount "cell" None with
        Selection =
            Some {
                Anchor = { BlockId = "bulk-replace-table-small-00-00000"; Offset = 0 }
                Focus = { BlockId = smallBulkReplaceTableEnd; Offset = 4 }
            } }
match Doc.setColor (Some "#123456") smallBulkReplaceTableDoc |> Doc.findParagraph <| smallBulkReplaceTableEnd with
| Some paragraph -> check "bulk replace table changed style" (getStyleAt 0 paragraph.Buffer).Color (Some "#123456")
| None ->
    fail <- fail + 1
    printfn "FAIL %-28s expected paragraph" "bulk replace table changed style"

let nestedLookupTable =
    singleCellTable "nested-lookup-table" (tableCell "nested-lookup" "nested")
let nestedLookupCell =
    { tableCell "outer-before" "before" with
        Content =
            [ para "outer-before" (ofString "before")
              para "outer-after" (ofString "after") ]
        NestedTables = [ nestedLookupTable ]
        NestedTableAnchors = [ 1 ] }
let nestedLookupDoc =
    { Doc.empty with
        Blocks = [ Doc.Table (singleCellTable "outer-lookup-table" nestedLookupCell) ] }
check "find paragraph table before nested" (Doc.findParagraph nestedLookupDoc "outer-before" |> Option.map (fun p -> toString p.Buffer)) (Some "before")
check "find paragraph table after nested" (Doc.findParagraph nestedLookupDoc "outer-after" |> Option.map (fun p -> toString p.Buffer)) (Some "after")
let deepNestedLookupDepth = 50
let mutable deepNestedLookupTable =
    singleCellTable "deep-nested-leaf-table" (tableCell "deep-nested-leaf" "leaf")
for depth = deepNestedLookupDepth - 1 downto 0 do
    let cell =
        { tableCell (sprintf "deep-nested-holder-%05d" depth) "" with
            Content = []
            NestedTables = [ deepNestedLookupTable ]
            NestedTableAnchors = [ 0 ] }
    deepNestedLookupTable <- singleCellTable (sprintf "deep-nested-table-%05d" depth) cell
let deepNestedLookupDoc = { Doc.empty with Blocks = [ Doc.Table deepNestedLookupTable ] }
let deepNestedLookupResult = Doc.findParagraph deepNestedLookupDoc "deep-nested-leaf"

let deepNestedVbaStats =
    Vba.runDocument deepNestedLookupDoc
        """
Dim x
x = ActiveDocument.Characters.Count
x = ActiveDocument.ComputeStatistics(wdStatisticCharacters)
"""
check "vba deep nested table source counts"
    (match deepNestedVbaStats with | Ok _ -> true | Error _ -> false)
    true

let vbaTextStatsDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph (para "vba-stats-1" (ofString "alpha beta"))
              Doc.Paragraph (para "vba-stats-2" (ofString "gamma")) ] }
let vbaTextStats =
    Vba.runDocument vbaTextStatsDoc
        """
Dim x
x = ActiveDocument.Words.Count
x = ActiveDocument.Characters.Count
x = ActiveDocument.ComputeStatistics(wdStatisticWords)
x = ActiveDocument.ComputeStatistics(wdStatisticCharacters)
"""
check "vba text statistics ok"
    (match vbaTextStats with | Ok _ -> true | Error _ -> false)
    true

// removed: tested deleted cache/perf machinery

let contentControlPara id text controlId tag alias : Doc.ParagraphBlock =
    { para id (ofString text) with
        Metadata = {
            Doc.defaultMetadata with
                ContentControlId = Some controlId
                ContentControlTag = tag
                ContentControlAlias = alias
        } }
let pos blockId offset : Doc.Position = { BlockId = blockId; Offset = offset }

let resizeImagesBuffer =
    empty
    |> insertImage 0 "first" (Some 10.0) (Some 10.0) None
    |> insertImage 1 "second" (Some 20.0) (Some 20.0) None
let resizeImageDoc =
    { Doc.empty with
        Blocks = [ Doc.Paragraph (para "resize-images" resizeImagesBuffer) ]
        Selection = Some { Anchor = pos "resize-images" 1; Focus = pos "resize-images" 1 } }
let resizedImageDoc = Doc.resizeImageNearCursor 40.0 50.0 resizeImageDoc
match Doc.findParagraph resizedImageDoc "resize-images" with
| Some p ->
    match p.Buffer.Runs with
    | [ Image first; Image second ] ->
        check "resize near keeps first width" first.Width (Some 10.0)
        check "resize near updates second width" second.Width (Some 40.0)
        check "resize near updates second height" second.Height (Some 50.0)
    | runs ->
        fail <- fail + 1
        printfn "FAIL %-28s expected two image runs got %A" "resize near image runs" runs.Length
| None ->
    fail <- fail + 1
    printfn "FAIL %-28s expected paragraph" "resize near image paragraph"
let resizeMissBuffer = resizeImage 99 40.0 50.0 resizeImagesBuffer
let resizeSameBuffer = resizeImage 0 10.0 10.0 resizeImagesBuffer

let moveImagesBuffer =
    empty
    |> insertImage 0 "first" (Some 10.0) (Some 10.0) None
    |> insertImage 1 "second" (Some 20.0) (Some 20.0) None
    |> insertImage 2 "third" (Some 30.0) (Some 30.0) None
let moveImageDoc =
    { Doc.empty with Blocks = [ Doc.Paragraph (para "move-images" moveImagesBuffer) ] }
let imageSources buffer =
    buffer.Runs
    |> List.choose (function Image img -> Some img.Src | _ -> None)
match moveImagesBuffer.Runs with
| Image first :: _ ->
    match tryMoveImageWithinBuffer 0 0 first moveImagesBuffer with
    | Some sameMoveBuffer -> ()
    | None ->
        fail <- fail + 1
        printfn "FAIL %-28s expected same buffer" "move image same buffer"
| _ ->
    fail <- fail + 1
    printfn "FAIL %-28s expected image run" "move image same source"
let movedForwardDoc = Doc.moveImage "move-images" 0 "move-images" 3 moveImageDoc
match Doc.findParagraph movedForwardDoc "move-images", movedForwardDoc.Selection with
| Some p, Some sel ->
    check "move image forward order" (imageSources p.Buffer) ["second"; "third"; "first"]
    check "move image forward cursor" sel.Focus.Offset 3
| _ ->
    fail <- fail + 1
    printfn "FAIL %-28s expected paragraph selection" "move image forward"
let movedBackwardDoc = Doc.moveImage "move-images" 2 "move-images" 0 moveImageDoc
match Doc.findParagraph movedBackwardDoc "move-images", movedBackwardDoc.Selection with
| Some p, Some sel ->
    check "move image backward order" (imageSources p.Buffer) ["third"; "first"; "second"]
    check "move image backward cursor" sel.Focus.Offset 1
| _ ->
    fail <- fail + 1
    printfn "FAIL %-28s expected paragraph selection" "move image backward"

let splitDoc =
    { Doc.empty with
        Blocks = [ Doc.Paragraph (para "p1" styledBuffer) ]
        Selection = Some { Anchor = pos "p1" 2; Focus = pos "p1" 2 } }
match Doc.splitBlock splitDoc with
| { Blocks = [ Doc.Paragraph before; Doc.Paragraph after ] } ->
    check "split preserves before text" (toString before.Buffer) "ab"
    check "split preserves after text" (toString after.Buffer) "cd"
    check "split keeps before style" (getStyleAt 0 before.Buffer).Bold true
    check "split keeps after style" (getStyleAt 0 after.Buffer).Color (Some "#ff0000")
| doc ->
    fail <- fail + 1
    printfn "FAIL %-28s expected two paragraphs got %A" "split block shape" doc.Blocks.Length

let findDoc =
    { Doc.empty with
        Blocks = [ Doc.Paragraph (para "find1" (ofString "Target target")) ] }
match Doc.findText "Target" true findDoc with
| [ matchInfo ] ->
    check "find case text" matchInfo.Text "Target"
| matches ->
    fail <- fail + 1
    printfn "FAIL %-28s expected one match got %A" "find case count" matches.Length
match Doc.findText "target" false findDoc with
| first :: _ ->
    check "find nocase preserves text" first.Text "Target"
| matches ->
    fail <- fail + 1
    printfn "FAIL %-28s expected matches got %A" "find nocase count" matches.Length
let findListA = Doc.findText "target" false findDoc
let findListB = Doc.findText "target" false findDoc
let findArray = Doc.findTextArray "target" false findDoc
check "find array count" findArray.Length 2
check "find array second text" findArray.[1].Text "target"
check "find empty list" (Doc.findText "" false findDoc) []
check "find empty array" (Doc.findTextArray "" false findDoc |> Array.length) 0
let changedFindDoc =
    { findDoc with Blocks = [ Doc.Paragraph (para "find1" (ofString "missing")) ] }
check "find changed blocks" (Doc.findTextArray "missing" false changedFindDoc |> Array.length) 1
// removed: tested deleted cache/perf machinery
let missingReplaceDoc, missingReplaceCount = Doc.replaceAllText "missing" "x" false findDoc
check "replace missing count" missingReplaceCount 0
let changedMissingReplaceDoc, changedMissingReplaceCount = Doc.replaceAllText "missing" "x" false changedFindDoc
check "replace changed missing count" changedMissingReplaceCount 1
check "replace changed missing text" (Doc.getText changedMissingReplaceDoc) "x"
let directMissingDoc =
    { Doc.empty with Blocks = [ Doc.Paragraph (para "direct-missing" (ofString "alpha beta")) ] }
let directMissingReplaceA, directMissingCountA = Doc.replaceAllText "absent-direct" "x" false directMissingDoc
let directMissingReplaceB, directMissingCountB = Doc.replaceAllText "absent-direct" "x" false directMissingDoc
check "replace direct missing first count" directMissingCountA 0
check "replace direct missing second count" directMissingCountB 0
let directMissingChanged =
    { directMissingDoc with Blocks = [ Doc.Paragraph (para "direct-missing" (ofString "absent-direct")) ] }
let directMissingChangedReplace, directMissingChangedCount = Doc.replaceAllText "absent-direct" "x" false directMissingChanged
check "replace direct missing changed count" directMissingChangedCount 1
check "replace direct missing changed text" (Doc.getText directMissingChangedReplace) "x"
let replaceMissingParagraphCount = 200
let replaceMissingParagraphDoc =
    { Doc.empty with
        Blocks =
            [ for index in 0 .. replaceMissingParagraphCount - 1 ->
                Doc.Paragraph (para (sprintf "replace-missing-%06d" index) (ofString "alpha beta")) ] }
let replaceMissingParagraphResult, replaceMissingParagraphResultCount =
    Doc.replaceAllText "replace-large-absent" "x" false replaceMissingParagraphDoc
check "replace missing paragraphs count" replaceMissingParagraphResultCount 0
let replaceMissingTableDoc =
    singleColumnTableDoc "replace-missing-table-small" 0 200 "alpha beta" None
let findMissingTableResults = Doc.findTextArray "find-large-table-absent" false replaceMissingTableDoc
check "find missing table count" findMissingTableResults.Length 0
let replaceMissingTableResult, replaceMissingTableCount =
    Doc.replaceAllText "replace-large-table-absent" "x" false replaceMissingTableDoc
check "replace missing table count" replaceMissingTableCount 0
let selectedFindDoc = Doc.selectFindMatch findArray.[0] findDoc
let crossLeafText = System.String('a', 1023) + "NeEdLe" + System.String('b', 8)
let crossLeafDoc =
    { Doc.empty with
        Blocks = [ Doc.Paragraph (para "find-cross-leaf" (ofString crossLeafText)) ] }
let crossLeafMatches = Doc.findTextArray "needle" false crossLeafDoc
check "find cross leaf count" crossLeafMatches.Length 1
check "find cross leaf start" crossLeafMatches.[0].Start 1023
check "find cross leaf text" crossLeafMatches.[0].Text "NeEdLe"
let crossLeafReplaced, crossLeafReplaceCount = Doc.replaceAllText "needle" "x" false crossLeafDoc
check "replace cross leaf count" crossLeafReplaceCount 1
check "replace cross leaf text" (Doc.getText crossLeafReplaced) (System.String('a', 1023) + "x" + System.String('b', 8))
let sameInsensitiveReplaceDoc =
    { Doc.empty with
        Blocks = [ Doc.Paragraph (para "same-insensitive-replace" (ofString "target target")) ] }
let sameInsensitiveReplaced, sameInsensitiveReplaceCount = Doc.replaceAllText "TARGET" "target" false sameInsensitiveReplaceDoc
check "replace same insensitive count" sameInsensitiveReplaceCount 2

let bookmarkStyle =
    { defaultStyle with
        BookmarkStartId = Some "bookmark-id-1"
        BookmarkName = Some "Bookmark_1"
        BookmarkEndId = Some "bookmark-id-1" }
let bookmarkDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph
                { para "bm1" (ofString "")
                    with Buffer =
                            { Rope = Rbt.Word.Rope.ofString "abcdef"
                              Runs =
                                [ Text { Start = 0; Length = 1; Style = defaultStyle }
                                  Text { Start = 1; Length = 2; Style = bookmarkStyle }
                                  Text { Start = 3; Length = 2; Style = bookmarkStyle }
                                  Text { Start = 5; Length = 1; Style = defaultStyle } ] } } ] }
let selectedBookmarkByName = Doc.selectBookmark "Bookmark_1" bookmarkDoc
match selectedBookmarkByName.Selection with
| Some selection ->
    check "bookmark select name start" selection.Anchor.Offset 1
    check "bookmark select name end" selection.Focus.Offset 5
| None ->
    fail <- fail + 1
    printfn "FAIL %-28s expected selection" "bookmark select name"
let selectedBookmarkById = Doc.selectBookmark "bookmark-id-1" bookmarkDoc
match selectedBookmarkById.Selection with
| Some selection ->
    check "bookmark select id block" selection.Anchor.BlockId "bm1"
    check "bookmark select id end" selection.Focus.Offset 5
| None ->
    fail <- fail + 1
    printfn "FAIL %-28s expected selection" "bookmark select id"
// removed: tested deleted cache/perf machinery
let renamedBookmarkDoc = Doc.renameBookmark "bookmark-id-1" "Renamed Bookmark" bookmarkDoc
match Doc.bookmarkInfos renamedBookmarkDoc with
| [ info ] -> check "bookmark rename changed name" info.Name "Renamed_Bookmark"
| infos ->
    fail <- fail + 1
    printfn "FAIL %-28s expected one bookmark got %A" "bookmark rename changed" infos.Length
let bookmarkIdParagraphCount = 200
let bookmarkIdBlockId = "bookmark-id-small-00000"
let mutable bookmarkIdDoc =
    { Doc.empty with
        Blocks =
            [ for index in 0 .. bookmarkIdParagraphCount - 1 ->
                Doc.Paragraph (para (sprintf "bookmark-id-small-%05d" index) (ofString "plain text")) ] }
for index in 1 .. 20 do
    bookmarkIdDoc <-
        Doc.setBookmark (sprintf "Bookmark %d" index)
            { bookmarkIdDoc with
                Selection = Some { Anchor = pos bookmarkIdBlockId 0; Focus = pos bookmarkIdBlockId 1 } }
let bookmarkIdLatest =
    Doc.findParagraph bookmarkIdDoc bookmarkIdBlockId
    |> Option.map (fun p -> (getStyleAt 0 p.Buffer).BookmarkStartId)
    |> Option.defaultValue None
check "bookmark id numbering" bookmarkIdLatest (Some "20")
// removed: tested deleted cache/perf machinery

let contentControlDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph (contentControlPara "cc1" "Alpha" "control-1" (Some "invoice") (Some "Invoice"))
              Doc.Paragraph (contentControlPara "cc2" "Beta" "control-2" None None) ] }
// removed: tested deleted cache/perf machinery
let commentStyle =
    { defaultStyle with
        CommentId = Some "comment-1"
        CommentText = Some "Needs review"
        CommentAuthor = Some "author"
        CommentDate = Some "date" }
let commentDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph
                { para "comment-p" (ofString "")
                    with Buffer =
                            { Rope = Rbt.Word.Rope.ofString "abcdef"
                              Runs =
                                [ Text { Start = 0; Length = 1; Style = defaultStyle }
                                  Text { Start = 1; Length = 3; Style = commentStyle }
                                  Text { Start = 4; Length = 2; Style = defaultStyle } ] } } ] }
match Doc.commentInfos commentDoc with
| [ info ] ->
    check "comment infos id" info.Id "comment-1"
    check "comment infos block" info.BlockId "comment-p"
    check "comment infos start" info.StartOffset 1
    check "comment infos end" info.EndOffset 4
    check "comment infos excerpt" info.Excerpt "bcd"
| infos ->
    fail <- fail + 1
    printfn "FAIL %-28s expected one comment got %A" "comment infos count" infos.Length
// removed: tested deleted cache/perf machinery
// removed: tested deleted cache/perf machinery
let selectedByTag = Doc.selectContentControl "invoice" contentControlDoc
match selectedByTag.Selection with
| Some selection ->
    check "content select tag block" selection.Anchor.BlockId "cc1"
    check "content select tag focus" selection.Focus.Offset 5
| None ->
    fail <- fail + 1
    printfn "FAIL %-28s expected selection" "content select tag"
let selectedByDefaultTag = Doc.selectContentControl "content-control" contentControlDoc
match selectedByDefaultTag.Selection with
| Some selection ->
    check "content select default" selection.Anchor.BlockId "cc2"
| None ->
    fail <- fail + 1
    printfn "FAIL %-28s expected selection" "content select default"
let removedByTag = Doc.removeContentControl "invoice" contentControlDoc
match Doc.findParagraph removedByTag "cc1" with
| Some p ->
    check "content remove tag id" p.Metadata.ContentControlId None
    check "content remove tag tag" p.Metadata.ContentControlTag None
| None ->
    fail <- fail + 1
    printfn "FAIL %-28s expected paragraph" "content remove tag"
let updatedByTag = Doc.updateContentControl "invoice" "control-3" "paid" "Paid" contentControlDoc
match Doc.findParagraph updatedByTag "cc1" with
| Some p ->
    check "content update tag id" p.Metadata.ContentControlId (Some "control-3")
    check "content update tag tag" p.Metadata.ContentControlTag (Some "paid")
    check "content update tag alias" p.Metadata.ContentControlAlias (Some "Paid")
| None ->
    fail <- fail + 1
    printfn "FAIL %-28s expected paragraph" "content update tag"

let noFieldRefreshDoc =
    { Doc.empty with
        Blocks =
            [ for i in 0 .. 9 ->
                Doc.Paragraph (para (sprintf "no-field-%i" i) (ofString "plain text")) ] }
// removed: tested deleted cache/perf machinery

let pageFieldStyle =
    { defaultStyle with FieldCode = Some " PAGE "; FieldResult = Some "0" }
let pageRefFieldStyle =
    { defaultStyle with FieldCode = Some " PAGEREF Bookmark_1 "; FieldResult = Some "0" }
let fieldRefreshDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph
                { Id = "field-ref"
                  Buffer =
                    { Rope = Rbt.Word.Rope.ofString "mark"
                      Runs = [ Text { Start = 0; Length = 4; Style = bookmarkStyle } ] }
                  Metadata = { Doc.defaultMetadata with PageBreakAfter = true } }
              Doc.Paragraph
                { Id = "field-page"
                  Buffer =
                    { Rope = Rbt.Word.Rope.ofString "0 0"
                      Runs =
                        [ Text { Start = 0; Length = 1; Style = pageFieldStyle }
                          Text { Start = 1; Length = 1; Style = defaultStyle }
                          Text { Start = 2; Length = 1; Style = pageRefFieldStyle } ] }
                  Metadata = Doc.defaultMetadata } ] }
let refreshedFields = Doc.refreshFields fieldRefreshDoc
match Doc.findParagraph refreshedFields "field-page" with
| Some p ->
    check "field refresh page" (toString p.Buffer) "2 1"
| None ->
    fail <- fail + 1
    printfn "FAIL %-28s expected paragraph" "field refresh page"

let pageBreakNamedStyle styleType id pageBreakBefore : Doc.NamedStyle =
    {
        Id = id
        Name = Some id
        StyleType = styleType
        IsDefault = false
        BasedOn = None
        Next = None
        Linked = None
        IsQuickStyle = false
        Aliases = None
        UiPriority = None
        SemiHidden = false
        UnhideWhenUsed = false
        Text = None
        TextBold = None
        TextItalic = None
        TextUnderline = None
        TextAllCaps = None
        TextSmallCaps = None
        TextStrike = None
        SpacingBefore = None
        SpacingAfter = None
        ContextualSpacing = None
        LineSpacing = None
        LineSpacingRule = None
        Alignment = None
        IndentLeft = None
        IndentRight = None
        FirstLineIndent = None
        BackgroundColor = None
        OutlineLevel = None
        KeepNext = None
        KeepLines = None
        PageBreakBefore = pageBreakBefore
        WidowControl = None
    }
let pageMapStyleDoc =
    { Doc.empty with
        Styles =
            [ pageBreakNamedStyle "defaults" "defaults" (Some true)
              pageBreakNamedStyle "paragraph" "NoBreak" (Some false) ]
        Blocks =
            [ Doc.Paragraph (para "pb1" (ofString "one"))
              Doc.Paragraph { para "pb2" (ofString "two") with Metadata = { Doc.defaultMetadata with Style = Some "NoBreak" } }
              Doc.Paragraph { para "pb3" (ofString "three") with Metadata = { Doc.defaultMetadata with Style = Some "MissingStyle" } } ] }
let pageMapStyle = Doc.buildPageMap pageMapStyleDoc
let pageFor blockId =
    pageMapStyle.BlockAnchors
    |> List.find (fun anchor -> anchor.PageBlockId = blockId)
    |> fun anchor -> anchor.PageNumber
check "page map default break" (pageFor "pb1", pageFor "pb2", pageFor "pb3") (1, 1, 2)
let pageMapStyleChanged =
    Doc.buildPageMap
        { pageMapStyleDoc with
            Styles =
                [ pageBreakNamedStyle "defaults" "defaults" (Some false)
                  pageBreakNamedStyle "paragraph" "NoBreak" (Some true) ] }
let changedPageFor blockId =
    pageMapStyleChanged.BlockAnchors
    |> List.find (fun anchor -> anchor.PageBlockId = blockId)
    |> fun anchor -> anchor.PageNumber
check "page map style changed" (changedPageFor "pb1", changedPageFor "pb2", changedPageFor "pb3") (1, 2, 2)
let pageMapStyleAgain = Doc.buildPageMap pageMapStyleDoc
let pageAgainFor blockId =
    pageMapStyleAgain.BlockAnchors
    |> List.find (fun anchor -> anchor.PageBlockId = blockId)
    |> fun anchor -> anchor.PageNumber
check "page map style original" (pageAgainFor "pb1", pageAgainFor "pb2", pageAgainFor "pb3") (1, 1, 2)
// removed: tested deleted cache/perf machinery

let tocFieldStyle =
    { defaultStyle with FieldCode = Some " TOC \\o \"1-2\" "; FieldResult = Some "" }
let tocRefreshDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph
                { Id = "toc-field"
                  Buffer =
                    { Rope = Rbt.Word.Rope.ofString "toc"
                      Runs = [ Text { Start = 0; Length = 3; Style = tocFieldStyle } ] }
                  Metadata = Doc.defaultMetadata }
              Doc.Heading (1, para "toc-h1" (ofString "Intro"))
              Doc.Heading (2, para "toc-h2" (ofString "Details"))
              Doc.Heading (3, para "toc-h3" (ofString "Hidden")) ] }
let refreshedToc = Doc.refreshFields tocRefreshDoc
match Doc.findParagraph refreshedToc "toc-field" with
| Some p ->
    check "field refresh toc headings" (toString p.Buffer) "Intro\t1\n  Details\t1"
| None ->
    fail <- fail + 1
    printfn "FAIL %-28s expected paragraph" "field refresh toc"

let fieldParagraph id text fieldCode bookmarkName =
    let style = { defaultStyle with FieldCode = Some fieldCode; FieldResult = Some text; BookmarkName = bookmarkName }
    Doc.Paragraph
        { Id = id
          Buffer =
            { Rope = Rbt.Word.Rope.ofString text
              Runs = [ Text { Start = 0; Length = text.Length; Style = style } ] }
          Metadata = Doc.defaultMetadata }
let sequenceFieldDoc =
    { Doc.empty with
        Blocks =
            [ fieldParagraph "seq-1" "0" " SEQ Figure \\* ARABIC " (Some "Fig1")
              fieldParagraph "seq-2" "0" " SEQ Figure \\* ARABIC " (Some "Fig2")
              fieldParagraph "ref-2" "0" " REF Fig2 \\h " None
              fieldParagraph "toc-figures" "toc" " TOC \\c \"Figure\" " None ] }
let refreshedSequenceFields = Doc.refreshFields sequenceFieldDoc
let refreshedText blockId =
    Doc.findParagraph refreshedSequenceFields blockId |> Option.map (fun p -> toString p.Buffer) |> Option.defaultValue ""
check "field refresh seq" (refreshedText "seq-1", refreshedText "seq-2") ("1", "2")
check "field refresh ref" (refreshedText "ref-2") "2"
check "field refresh toc captions" (refreshedText "toc-figures") "1\t1\n2\t1"
let refreshedSequenceFieldsAgain = Doc.refreshFields refreshedSequenceFields

let captionParagraphCount = 20
let captionLastBlockId = sprintf "caption-small-%05d" (captionParagraphCount - 1)
let captionDoc =
    { Doc.empty with
        Blocks =
            [ for index in 0 .. captionParagraphCount - 1 ->
                Doc.Paragraph (para (sprintf "caption-small-%05d" index) (ofString "plain text")) ]
        Selection = Some { Anchor = pos captionLastBlockId 10; Focus = pos captionLastBlockId 10 } }
let mutable captionCurrent = Doc.insertCaption "Figure" "caption 1" captionDoc
for index in 2 .. 5 do
    captionCurrent <- Doc.insertCaption "Figure" (sprintf "caption %d" index) captionCurrent
let captionTexts =
    captionCurrent.Blocks
    |> List.choose (function
        | Doc.Paragraph p when p.Metadata.Style = Some "Caption" -> Some (toString p.Buffer)
        | _ -> None)
check "insert caption count" captionTexts.Length 5
check "insert caption numbering" (captionTexts |> List.exists ((=) "Figure 5: caption 5")) true
let captionCrossRef = Doc.insertCrossReferenceToLatestCaption "Figure" captionCurrent
check "insert caption cross ref"
    (Doc.findParagraph captionCrossRef captionLastBlockId |> Option.map (fun p -> toString p.Buffer) |> Option.defaultValue "")
    "plain text5"
// removed: tested deleted cache/perf machinery

let styleRefFieldStyle =
    { defaultStyle with FieldCode = Some " STYLEREF Heading1 "; FieldResult = Some "old" }
let styleRefDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph { para "style-source" (ofString "Chapter One") with Metadata = { Doc.defaultMetadata with Style = Some "Heading1" } }
              Doc.Paragraph
                { Id = "style-ref"
                  Buffer =
                    { Rope = Rbt.Word.Rope.ofString "old"
                      Runs = [ Text { Start = 0; Length = 3; Style = styleRefFieldStyle } ] }
                  Metadata = Doc.defaultMetadata } ] }
let refreshedStyleRef = Doc.refreshFields styleRefDoc
check "field refresh styleref"
    (Doc.findParagraph refreshedStyleRef "style-ref" |> Option.map (fun p -> toString p.Buffer) |> Option.defaultValue "")
    "Chapter One"

let mergeDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph (para "p1" (sliceBuffer 0 2 styledBuffer))
              Doc.Paragraph (para "p2" (sliceBuffer 2 2 styledBuffer)) ]
        Selection = Some { Anchor = pos "p2" 0; Focus = pos "p2" 0 } }
match Doc.mergeWithPrevious mergeDoc "p2" with
| { Blocks = [ Doc.Paragraph merged ] } ->
    check "merge preserves text" (toString merged.Buffer) "abcd"
    check "merge keeps first style" (getStyleAt 0 merged.Buffer).Bold true
    check "merge keeps second style" (getStyleAt 3 merged.Buffer).Color (Some "#ff0000")
| doc ->
    fail <- fail + 1
    printfn "FAIL %-28s expected one paragraph got %A" "merge block shape" doc.Blocks.Length
match Doc.mergeWithNext mergeDoc "p1" with
| { Blocks = [ Doc.Paragraph merged ] } ->
    check "merge next preserves text" (toString merged.Buffer) "abcd"
    check "merge next keeps second style" (getStyleAt 3 merged.Buffer).Color (Some "#ff0000")
| doc ->
    fail <- fail + 1
    printfn "FAIL %-28s expected one paragraph got %A" "merge next block shape" doc.Blocks.Length
let mergeTailBlocks prefix =
    [ for index in 0 .. 4 ->
        Doc.Paragraph (para (sprintf "%s-tail-%05d" prefix index) (ofString "tail")) ]
let bulkFormatSmallDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph (para "bulk-format-small-a" (ofString "abcd"))
              Doc.Paragraph (para "bulk-format-small-b" (ofString "efgh")) ]
        Selection = Some { Anchor = pos "bulk-format-small-a" 1; Focus = pos "bulk-format-small-b" 2 } }
match Doc.setColor (Some "#654321") bulkFormatSmallDoc with
| { Blocks = [ Doc.Paragraph first; Doc.Paragraph second ] } ->
    check "bulk format first style" (getStyleAt 1 first.Buffer).Color (Some "#654321")
    check "bulk format second style" (getStyleAt 1 second.Buffer).Color (Some "#654321")
| doc ->
    fail <- fail + 1
    printfn "FAIL %-28s expected formatted paragraphs got %A" "bulk format style" doc.Blocks.Length
// removed: tested deleted cache/perf machinery

let selectionDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph (para "p1" styledBuffer)
              Doc.Paragraph (para "p2" styledBuffer) ]
        Selection = Some { Anchor = pos "p1" 1; Focus = pos "p2" 3 } }
let changedSelectionDoc = { selectionDoc with Selection = Some { Anchor = pos "p1" 1; Focus = pos "p1" 3 } }
check "selection text changed selection" (Doc.getSelectionText changedSelectionDoc) "bc"
// removed: tested deleted cache/perf machinery
let sameBlockSelectionDoc = { selectionDoc with Selection = Some { Anchor = pos "p1" 1; Focus = pos "p1" 3 } }
let cursorDoc = { selectionDoc with Selection = Some { Anchor = pos "p1" 1; Focus = pos "p1" 1 } }
let lineStartDoc = { selectionDoc with Selection = Some { Anchor = pos "p1" 0; Focus = pos "p1" 0 } }
let lineEndDoc = { selectionDoc with Selection = Some { Anchor = pos "p1" (length styledBuffer); Focus = pos "p1" (length styledBuffer) } }
let inputDocumentsEqualForChange =
    typeof<Doc.Document>.Assembly.GetType("Rbt.Word.Input").GetMethod(
        "documentsEqualForChange",
        System.Reflection.BindingFlags.NonPublic ||| System.Reflection.BindingFlags.Static)
if isNull inputDocumentsEqualForChange then
    fail <- fail + 1
    printfn "FAIL %-28s expected private documentsEqualForChange method" "input change equality reflection"
else
    let blocks =
        [ for index in 1 .. 200 ->
            Doc.Paragraph (para (sprintf "input-change-equality-%05d" index) empty) ]
    let left =
        { Doc.empty with
            Blocks = blocks
            Selection = Some { Anchor = pos "input-change-equality-00001" 0; Focus = pos "input-change-equality-00001" 0 } }
    let right =
        { left with
            Selection = Some { Anchor = pos "input-change-equality-00001" 1; Focus = pos "input-change-equality-00001" 1 } }
    let equal = inputDocumentsEqualForChange.Invoke(null, [| box left; box right |]) :?> bool
    check "input change equality detects selection" equal false
let selectAllNestedDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Table (singleCellTable "select-all-table" (tableCell "select-all-cell" "cell"))
              Doc.Paragraph (para "select-all-last" (ofString "tail")) ] }
match Doc.selectAll selectAllNestedDoc with
| { Selection = Some sel } ->
    check "select all nested anchor" sel.Anchor.BlockId "select-all-cell"
    check "select all nested focus" sel.Focus.BlockId "select-all-last"
    check "select all nested end" sel.Focus.Offset 4
| _ ->
    fail <- fail + 1
    printfn "FAIL %-28s expected selection" "select all nested"
let selectedAllNestedDoc = Doc.selectAll selectAllNestedDoc
let changedSelectAllNestedDoc =
    { selectedAllNestedDoc with Blocks = [ Doc.Paragraph (para "select-all-changed" (ofString "changed")) ] }
match Doc.selectAll changedSelectAllNestedDoc with
| { Selection = Some sel } ->
    check "select all changed block" sel.Anchor.BlockId "select-all-changed"
    check "select all changed end" sel.Focus.Offset 7
| _ ->
    fail <- fail + 1
    printfn "FAIL %-28s expected selection" "select all invalidates"
// removed: tested deleted cache/perf machinery
match Doc.deleteSelection selectionDoc with
| { Blocks = [ Doc.Paragraph merged ] } ->
    check "delete selection merged text" (toString merged.Buffer) "ad"
| doc ->
    fail <- fail + 1
    printfn "FAIL %-28s expected one paragraph got %A" "delete selection shape" doc.Blocks.Length

let reverseSelectionDoc =
    { selectionDoc with
        Selection = Some { Anchor = pos "p2" 3; Focus = pos "p1" 1 } }
match Doc.deleteSelection reverseSelectionDoc with
| { Blocks = [ Doc.Paragraph merged ] } ->
    check "delete reversed selection text" (toString merged.Buffer) "ad"
| doc ->
    fail <- fail + 1
    printfn "FAIL %-28s expected one paragraph got %A" "delete reversed shape" doc.Blocks.Length

// removed: tested deleted cache/perf machinery

let reverseFormatDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph (para "fmt1" (ofString "abcd"))
              Doc.Paragraph (para "fmt2" (ofString "abcd")) ]
        Selection = Some { Anchor = pos "fmt2" 2; Focus = pos "fmt1" 1 } }
match Doc.toggleBold reverseFormatDoc with
| { Blocks = [ Doc.Paragraph p1; Doc.Paragraph p2 ] } ->
    check "format reversed first before" (getStyleAt 0 p1.Buffer).Bold false
    check "format reversed first selected" (getStyleAt 1 p1.Buffer).Bold true
    check "format reversed second selected" (getStyleAt 1 p2.Buffer).Bold true
    check "format reversed second after" (getStyleAt 2 p2.Buffer).Bold false
| doc ->
    fail <- fail + 1
    printfn "FAIL %-28s expected two paragraphs got %A" "format reversed shape" doc.Blocks.Length

let sameBlockDeleteDoc =
    { Doc.empty with
        Blocks = [ Doc.Paragraph (para "same-del" (ofString "abcd")) ]
        Selection = Some { Anchor = pos "same-del" 3; Focus = pos "same-del" 1 } }
match Doc.deleteSelection sameBlockDeleteDoc with
| { Blocks = [ Doc.Paragraph p ]; Selection = Some sel } ->
    check "delete same block reversed" (toString p.Buffer) "ad"
    check "delete same block cursor" sel.Focus.Offset 1
| doc ->
    fail <- fail + 1
    printfn "FAIL %-28s expected one paragraph got %A" "delete same block shape" doc.Blocks.Length

let sameBlockFormatDoc =
    { Doc.empty with
        Blocks = [ Doc.Paragraph (para "same-fmt" (ofString "abcd")) ]
        Selection = Some { Anchor = pos "same-fmt" 3; Focus = pos "same-fmt" 1 } }
match Doc.toggleBold sameBlockFormatDoc with
| { Blocks = [ Doc.Paragraph p ] } ->
    check "format same block text" (toString p.Buffer) "abcd"
    check "format same block before" (getStyleAt 0 p.Buffer).Bold false
    check "format same block first" (getStyleAt 1 p.Buffer).Bold true
    check "format same block last" (getStyleAt 2 p.Buffer).Bold true
    check "format same block after" (getStyleAt 3 p.Buffer).Bold false
| doc ->
    fail <- fail + 1
    printfn "FAIL %-28s expected one paragraph got %A" "format same block shape" doc.Blocks.Length

let collapsedFormatDoc =
    { sameBlockFormatDoc with Selection = Some { Anchor = pos "same-fmt" 2; Focus = pos "same-fmt" 2 } }
let sameColorBuffer = ofString "abcd" |> setColor 1 2 (Some "#c00000")
let sameColorDoc =
    { Doc.empty with
        Blocks = [ Doc.Paragraph (para "same-color" sameColorBuffer) ]
        Selection = Some { Anchor = pos "same-color" 3; Focus = pos "same-color" 1 } }
let multiColorDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph (para "same-color-a" (ofString "ab" |> setColor 0 2 (Some "#c00000")))
              Doc.Paragraph (para "same-color-b" (ofString "cd" |> setColor 0 2 (Some "#c00000"))) ]
        Selection = Some { Anchor = pos "same-color-a" 0; Focus = pos "same-color-b" 2 } }
let sameHighlightBuffer = ofString "abcd" |> setHighlight 1 2 (Some "yellow")
let sameHighlightDoc =
    { Doc.empty with
        Blocks = [ Doc.Paragraph (para "same-highlight" sameHighlightBuffer) ]
        Selection = Some { Anchor = pos "same-highlight" 3; Focus = pos "same-highlight" 1 } }
let multiHighlightDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph (para "same-highlight-a" (ofString "ab" |> setHighlight 0 2 (Some "yellow")))
              Doc.Paragraph (para "same-highlight-b" (ofString "cd" |> setHighlight 0 2 (Some "yellow"))) ]
        Selection = Some { Anchor = pos "same-highlight-a" 0; Focus = pos "same-highlight-b" 2 } }
let sameHyperlinkBuffer = ofString "abcd" |> setHyperlink 1 2 (Some "https://example.com")
let sameHyperlinkDoc =
    { Doc.empty with
        Blocks = [ Doc.Paragraph (para "same-hyperlink" sameHyperlinkBuffer) ]
        Selection = Some { Anchor = pos "same-hyperlink" 3; Focus = pos "same-hyperlink" 1 } }
let multiHyperlinkDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph (para "same-hyperlink-a" (ofString "ab" |> setHyperlink 0 2 (Some "https://example.com")))
              Doc.Paragraph (para "same-hyperlink-b" (ofString "cd" |> setHyperlink 0 2 (Some "https://example.com"))) ]
        Selection = Some { Anchor = pos "same-hyperlink-a" 0; Focus = pos "same-hyperlink-b" 2 } }
let sameVertBuffer = ofString "abcd" |> setVertAlign 1 2 (Some "superscript")
let sameVertDoc =
    { Doc.empty with
        Blocks = [ Doc.Paragraph (para "same-vert" sameVertBuffer) ]
        Selection = Some { Anchor = pos "same-vert" 3; Focus = pos "same-vert" 1 } }
let multiVertDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph (para "same-vert-a" (ofString "ab" |> setVertAlign 0 2 (Some "superscript")))
              Doc.Paragraph (para "same-vert-b" (ofString "cd" |> setVertAlign 0 2 (Some "superscript"))) ]
        Selection = Some { Anchor = pos "same-vert-a" 0; Focus = pos "same-vert-b" 2 } }
let sameFontSizeBuffer = ofString "abcd" |> setFontSize 1 2 14.0
let sameFontSizeDoc =
    { Doc.empty with
        Blocks = [ Doc.Paragraph (para "same-font-size" sameFontSizeBuffer) ]
        Selection = Some { Anchor = pos "same-font-size" 3; Focus = pos "same-font-size" 1 } }
let sameFontFamilyBuffer = ofString "abcd" |> setFontFamily 1 2 "Aptos"
let sameFontFamilyDoc =
    { Doc.empty with
        Blocks = [ Doc.Paragraph (para "same-font-family" sameFontFamilyBuffer) ]
        Selection = Some { Anchor = pos "same-font-family" 3; Focus = pos "same-font-family" 1 } }
let sameCharacterStyleBuffer = ofString "abcd" |> setCharacterStyle 1 2 (Some "Emphasis")
let sameCharacterStyleDoc =
    { Doc.empty with
        Blocks = [ Doc.Paragraph (para "same-character-style" sameCharacterStyleBuffer) ]
        Selection = Some { Anchor = pos "same-character-style" 3; Focus = pos "same-character-style" 1 } }
let multiCharacterStyleDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph (para "same-character-style-a" (ofString "ab" |> setCharacterStyle 0 2 (Some "Emphasis")))
              Doc.Paragraph (para "same-character-style-b" (ofString "cd" |> setCharacterStyle 0 2 (Some "Emphasis"))) ]
        Selection = Some { Anchor = pos "same-character-style-a" 0; Focus = pos "same-character-style-b" 2 } }
let sameCommentBuffer = ofString "abcd" |> setComment 1 2 (Some "comment-1") (Some "Needs review") (Some "author") (Some "date")
let sameCommentDoc =
    { Doc.empty with
        Blocks = [ Doc.Paragraph (para "same-comment" sameCommentBuffer) ]
        Selection = Some { Anchor = pos "same-comment" 3; Focus = pos "same-comment" 1 } }
let multiCommentDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph (para "same-comment-a" (ofString "ab" |> setComment 0 2 (Some "comment-1") (Some "Needs review") (Some "author") (Some "date")))
              Doc.Paragraph (para "same-comment-b" (ofString "cd" |> setComment 0 2 (Some "comment-1") (Some "Needs review") (Some "author") (Some "date"))) ]
        Selection = Some { Anchor = pos "same-comment-a" 0; Focus = pos "same-comment-b" 2 } }
let sameNoteBuffer = ofString "abcd" |> setNote 1 2 (Some "note-1") (Some "footnote") (Some "note text")
let sameNoteDoc =
    { Doc.empty with
        Blocks = [ Doc.Paragraph (para "same-note" sameNoteBuffer) ]
        Selection = Some { Anchor = pos "same-note" 3; Focus = pos "same-note" 1 } }
let multiNoteDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph (para "same-note-a" (ofString "ab" |> setNote 0 2 (Some "note-1") (Some "footnote") (Some "note text")))
              Doc.Paragraph (para "same-note-b" (ofString "cd" |> setNote 0 2 (Some "note-1") (Some "footnote") (Some "note text"))) ]
        Selection = Some { Anchor = pos "same-note-a" 0; Focus = pos "same-note-b" 2 } }
let sameRevisionBuffer = ofString "abcd" |> setRevision 1 2 "insert" "r1" (Some "author") (Some "date")
let sameRevisionDoc =
    { Doc.empty with
        Blocks = [ Doc.Paragraph (para "same-revision" sameRevisionBuffer) ]
        Selection = Some { Anchor = pos "same-revision" 3; Focus = pos "same-revision" 1 } }
let multiRevisionDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph (para "same-revision-a" (ofString "ab" |> setRevision 0 2 "insert" "r1" (Some "author") (Some "date")))
              Doc.Paragraph (para "same-revision-b" (ofString "cd" |> setRevision 0 2 "insert" "r1" (Some "author") (Some "date"))) ]
        Selection = Some { Anchor = pos "same-revision-a" 0; Focus = pos "same-revision-b" 2 } }
let cleanRevisionBuffer = ofString "abcd"
let headingDoc = Doc.setHeadingLevel 2 collapsedFormatDoc
let selectedCellDoc =
    { Doc.empty with
        Blocks = [ Doc.Table (singleCellTable "same-cell-style-table" (tableCell "same-cell-style-cell" "cell")) ]
        Selection = Some { Anchor = pos "same-cell-style-cell" 0; Focus = pos "same-cell-style-cell" 0 } }
let styledCellDoc =
    Doc.setSelectedTableCellStyle (Some "#fff000") (Some 80.0) (Some "#222222") (Some 1.0) (Some "center") (Some 2.0) (Some 3.0) (Some 4.0) (Some 5.0) selectedCellDoc
let tableMarginsDoc = Doc.setCurrentTableCellMargins (Some 2.0) (Some 3.0) (Some 4.0) (Some 5.0) styledCellDoc
let tableSpacingDoc = Doc.setCurrentTableCellMarginsAndSpacing (Some 2.0) (Some 3.0) (Some 4.0) (Some 5.0) (Some 6.0) tableMarginsDoc
let customPageSetup =
    { Doc.defaultPageSetup with PageWidth = 700.0; MarginLeft = 72.0 }
let pageSetupDoc = Doc.setPageSetup customPageSetup Doc.empty
let textBoxDoc =
    { Doc.empty with
        Blocks =
            [ Doc.TextBox
                { Id = "tb1"
                  Content = [ para "tbp1" (ofString "box") ]
                  X = Some 10.0
                  Y = Some 20.0
                  PositionHRelativeFrom = None
                  PositionVRelativeFrom = None
                  Width = 120.0
                  Height = 60.0
                  FillColor = Some "#ffffff"
                  BorderColor = Some "#111111"
                  BorderWidth = Some 1.0
                  RelativeHeight = None
                  Geometry = Doc.GeoRect
                  Gradient = None
                  Rotation = 0.0 } ]
        Selection = Some { Anchor = pos "tb1" 0; Focus = pos "tb1" 0 } }
// removed: tested deleted cache/perf machinery

let selectionTextDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph (para "sel1" (ofString "abcd"))
              Doc.Paragraph (para "sel2" (ofString "efgh")) ]
        Selection = Some { Anchor = pos "sel1" 1; Focus = pos "sel2" 3 } }
check "selection text multi" (Doc.getSelectionText selectionTextDoc) "bcd\nefg"
check "selection text reversed" (Doc.getSelectionText { selectionTextDoc with Selection = Some { Anchor = pos "sel2" 3; Focus = pos "sel1" 1 } }) "bcd\nefg"
check "selection text same" (Doc.getSelectionText { selectionTextDoc with Selection = Some { Anchor = pos "sel1" 1; Focus = pos "sel1" 3 } }) "bc"
check "selection text same reversed" (Doc.getSelectionText { selectionTextDoc with Selection = Some { Anchor = pos "sel1" 3; Focus = pos "sel1" 1 } }) "bc"
let farSelectionDoc =
    { Doc.empty with
        Blocks =
            [ for i in 0 .. 1499 ->
                Doc.Paragraph (para (sprintf "far-sel-%d" i) (ofString "abcd")) ] }
let farSelection : Doc.Selection =
    { Anchor = pos "far-sel-1200" 1; Focus = pos "far-sel-1201" 3 }
let farStart, farEnd, farIds = Doc.getNormalizedRange farSelection farSelectionDoc
check "normalized range far start" farStart.BlockId "far-sel-1200"
check "normalized range far end" farEnd.BlockId "far-sel-1201"
check "normalized range far ids" farIds [ "far-sel-1200"; "far-sel-1201" ]
check "selection text far indexed"
    (Doc.getSelectionText { farSelectionDoc with Selection = Some { Anchor = pos "far-sel-1201" 3; Focus = pos "far-sel-1200" 1 } })
    "bcd\nabc"
// removed: tested deleted cache/perf machinery
// removed: tested deleted cache/perf machinery

let replaceDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph (para "p1" (ofString "one target two target"))
              Doc.Paragraph (para "p2" (styledBuffer |> insert (length styledBuffer) " target")) ] }
let replacedDoc, replacedCount = Doc.replaceAllText "target" "x" false replaceDoc
check "replace all count" replacedCount 3
check "replace all text" (Doc.getText replacedDoc) "one x two x\nabcd x"
let replaceSharingBefore = para "replace-share-before" (ofString "before")
let replaceSharingTarget = para "replace-share-target" (ofString "needle")
let replaceSharingAfter = para "replace-share-after" (ofString "after")
let replaceSharingDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph replaceSharingBefore
              Doc.Paragraph replaceSharingTarget
              Doc.Paragraph replaceSharingAfter ] }
let replaceSharingResult, replaceSharingCount = Doc.replaceAllText "needle" "x" true replaceSharingDoc
check "replace sharing count" replaceSharingCount 1
match replaceSharingDoc.Blocks, replaceSharingResult.Blocks with
| Doc.Paragraph originalBefore :: Doc.Paragraph originalChanged :: originalTail,
  Doc.Paragraph newBefore :: Doc.Paragraph newChanged :: newTail ->
    check "replace shares before para" (obj.ReferenceEquals(originalBefore, newBefore)) true
    check "replace changes target para" (obj.ReferenceEquals(originalChanged, newChanged)) false
    check "replace shares block tail" (obj.ReferenceEquals(originalTail, newTail)) true
    match originalTail, newTail with
    | [ Doc.Paragraph originalAfter ], [ Doc.Paragraph newAfter ] ->
        check "replace shares after para" (obj.ReferenceEquals(originalAfter, newAfter)) true
    | _ ->
        fail <- fail + 1
        printfn "FAIL %-28s expected one shared tail paragraph" "replace sharing tail shape"
| _ ->
    fail <- fail + 1
    printfn "FAIL %-28s expected three paragraphs" "replace sharing shape"
let noReplaceDoc, noReplaceCount = Doc.replaceAllText "missing" "x" false replaceDoc
check "replace no match count" noReplaceCount 0
let sameReplaceDoc, sameReplaceCount = Doc.replaceAllText "target" "target" true replaceDoc
check "replace same count" sameReplaceCount 3
let singleRunRanges = ResizeArray<struct (int * int)>()
singleRunRanges.Add(struct (4, 10))
singleRunRanges.Add(struct (15, 21))
let singleRunReplaced = replaceRanges singleRunRanges "x" (ofString "one target two target")
check "replace single-run text" (toString singleRunReplaced) "one x two x"
check "replace single-run keeps one run" singleRunReplaced.Runs.Length 1
check "replace single-run style" (getStyleAt 4 singleRunReplaced).Bold false
let styledSingleRunReplaceDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph (
                para "styled-single-replace"
                    { Rope = Rbt.Word.Rope.ofString "target target"
                      Runs = [ Text { Start = 0; Length = 13; Style = { defaultStyle with Bold = true } } ] }) ] }
let styledSingleRunReplacedDoc, styledSingleRunReplaceCount = Doc.replaceAllText "target" "x" false styledSingleRunReplaceDoc
check "replace styled single-run count" styledSingleRunReplaceCount 2
match styledSingleRunReplacedDoc.Blocks with
| [ Doc.Paragraph p ] ->
    check "replace styled single-run text" (toString p.Buffer) "x x"
    check "replace styled single-run keeps one run" p.Buffer.Runs.Length 1
    check "replace styled single-run style" (getStyleAt 0 p.Buffer).Bold true
| blocks ->
    fail <- fail + 1
    printfn "FAIL %-28s expected one paragraph got %A" "replace styled single-run shape" blocks.Length
match replacedDoc.Blocks with
| [ Doc.Paragraph _; Doc.Paragraph replacedStyled ] ->
    check "replace all preserves style before match" (getStyleAt 1 replacedStyled.Buffer).Bold true
    check "replace all preserves style before appended match" (getStyleAt 2 replacedStyled.Buffer).Color (Some "#ff0000")
| blocks ->
    fail <- fail + 1
    printfn "FAIL %-28s expected two paragraphs got %A" "replace all shape" blocks.Length

let boundaryReplaceBuffer =
    { Rope = Rbt.Word.Rope.ofString "abtargetcd"
      Runs =
        [ Text { Start = 0; Length = 2; Style = { defaultStyle with Bold = true } }
          Text { Start = 2; Length = 8; Style = red } ] }
let boundaryReplaceDoc =
    { Doc.empty with
        Blocks = [ Doc.Paragraph (para "rb1" boundaryReplaceBuffer) ] }
let boundaryReplacedDoc, boundaryReplacedCount = Doc.replaceAllText "target" "x" false boundaryReplaceDoc
check "replace boundary count" boundaryReplacedCount 1
match boundaryReplacedDoc.Blocks with
| [ Doc.Paragraph p ] ->
    check "replace boundary text" (toString p.Buffer) "abxcd"
    check "replace boundary uses left style" (getStyleAt 2 p.Buffer).Bold true
    check "replace boundary keeps right style" (getStyleAt 3 p.Buffer).Color (Some "#ff0000")
| blocks ->
    fail <- fail + 1
    printfn "FAIL %-28s expected one paragraph got %A" "replace boundary shape" blocks.Length

let fragmentedReplaceText = "target target"
let fragmentedReplaceBuffer =
    { Rope = Rbt.Word.Rope.ofString fragmentedReplaceText
      Runs =
        [ for i in 0 .. fragmentedReplaceText.Length - 1 ->
            let style = if i % 2 = 0 then { defaultStyle with Bold = true } else red
            Text { Start = i; Length = 1; Style = style } ] }
let fragmentedReplaceDoc =
    { Doc.empty with
        Blocks = [ Doc.Paragraph (para "rf1" fragmentedReplaceBuffer) ] }
let fragmentedReplacedDoc, fragmentedReplacedCount = Doc.replaceAllText "target" "x" false fragmentedReplaceDoc
check "replace fragmented count" fragmentedReplacedCount 2
match fragmentedReplacedDoc.Blocks with
| [ Doc.Paragraph p ] ->
    check "replace fragmented text" (toString p.Buffer) "x x"
    check "replace fragmented first style" (getStyleAt 0 p.Buffer).Bold true
    check "replace fragmented second style" (getStyleAt 2 p.Buffer).Bold true
| blocks ->
    fail <- fail + 1
    printfn "FAIL %-28s expected one paragraph got %A" "replace fragmented shape" blocks.Length

let revisionBuffer revisionType =
    ofString "aREVb"
    |> setRevision 1 3 revisionType "r1" (Some "author") (Some "date")

let fragmentedRevisionBuffer =
    let style =
        { defaultStyle with
            RevisionType = Some "insert"
            RevisionId = Some "r1"
            RevisionAuthor = Some "author"
            RevisionDate = Some "date" }
    { Rope = Rbt.Word.Rope.ofString "abcd"
      Runs =
        [ for i in 0 .. 3 ->
            Text { Start = i; Length = 1; Style = style } ] }
let fragmentedRevisionDoc =
    { Doc.empty with Blocks = [ Doc.Paragraph (para "fragrev" fragmentedRevisionBuffer) ] }
check "revision infos fragmented runs" (Doc.revisionInfos fragmentedRevisionDoc |> List.length) 4
match Doc.selectNextRevision { fragmentedRevisionDoc with Selection = Some { Anchor = pos "fragrev" 0; Focus = pos "fragrev" 0 } } with
| { Selection = Some sel } ->
    check "select next fragmented revision start" sel.Anchor.Offset 1
    check "select next fragmented revision end" sel.Focus.Offset 2
| _ ->
    fail <- fail + 1
    printfn "FAIL %-28s expected selected revision" "select next fragmented revision"
match Doc.selectPreviousRevision { fragmentedRevisionDoc with Selection = Some { Anchor = pos "fragrev" 4; Focus = pos "fragrev" 4 } } with
| { Selection = Some sel } ->
    check "select previous fragmented revision start" sel.Anchor.Offset 3
    check "select previous fragmented revision end" sel.Focus.Offset 4
| _ ->
    fail <- fail + 1
    printfn "FAIL %-28s expected selected revision" "select previous fragmented revision"
let selectedFragmentedRevisionDoc = Doc.selectNextRevision { fragmentedRevisionDoc with Selection = Some { Anchor = pos "fragrev" 0; Focus = pos "fragrev" 0 } }
// removed: tested deleted cache/perf machinery
let splitRevisionBuffer =
    let style revisionId =
        { defaultStyle with
            RevisionType = Some "insert"
            RevisionId = Some revisionId }
    { Rope = Rbt.Word.Rope.ofString "abcd"
      Runs =
        [ Text { Start = 0; Length = 2; Style = style "r1" }
          Text { Start = 2; Length = 2; Style = style "r2" } ] }
let splitRevisionDoc =
    { Doc.empty with Blocks = [ Doc.Paragraph (para "splitrev" splitRevisionBuffer) ] }
check "revision infos split revision ids" (Doc.revisionInfos splitRevisionDoc |> List.length) 2
match Doc.selectNextRevision { splitRevisionDoc with Selection = Some { Anchor = pos "splitrev" 1; Focus = pos "splitrev" 1 } } with
| { Selection = Some sel } ->
    check "select next split revision start" sel.Anchor.Offset 2
    check "select next split revision end" sel.Focus.Offset 4
| _ ->
    fail <- fail + 1
    printfn "FAIL %-28s expected selected revision" "select next split revision"
match Doc.selectNextRevision { splitRevisionDoc with Selection = Some { Anchor = pos "splitrev" 4; Focus = pos "splitrev" 4 } } with
| { Selection = Some sel } ->
    check "select next split revision wrap start" sel.Anchor.Offset 0
    check "select next split revision wrap end" sel.Focus.Offset 2
| _ ->
    fail <- fail + 1
    printfn "FAIL %-28s expected selected revision" "select next split revision wrap"
match Doc.selectPreviousRevision { splitRevisionDoc with Selection = Some { Anchor = pos "splitrev" 2; Focus = pos "splitrev" 2 } } with
| { Selection = Some sel } ->
    check "select previous split revision start" sel.Anchor.Offset 0
    check "select previous split revision end" sel.Focus.Offset 2
| _ ->
    fail <- fail + 1
    printfn "FAIL %-28s expected selected revision" "select previous split revision"
match Doc.selectPreviousRevision { splitRevisionDoc with Selection = Some { Anchor = pos "splitrev" 0; Focus = pos "splitrev" 0 } } with
| { Selection = Some sel } ->
    check "select previous split wrap start" sel.Anchor.Offset 2
    check "select previous split wrap end" sel.Focus.Offset 4
| _ ->
    fail <- fail + 1
    printfn "FAIL %-28s expected selected revision" "select previous split wrap"
// removed: tested deleted cache/perf machinery

let trackedInsertBaseStyle =
    { defaultStyle with
        RevisionType = Some "insert"
        RevisionId = Some "7"
        RevisionAuthor = Some "author"
        RevisionDate = Some "date" }
let trackedInsertBaseDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph (para "tracked" {
                Rope = Rbt.Word.Rope.ofString "abc"
                Runs = [ Text { Start = 0; Length = 3; Style = trackedInsertBaseStyle } ]
            }) ] }
let trackedAdjacent =
    Doc.insertTrackedText "author" "date" "x"
        { trackedInsertBaseDoc with Selection = Some { Anchor = pos "tracked" 3; Focus = pos "tracked" 3 } }
match Doc.findParagraph trackedAdjacent "tracked" with
| Some p ->
    check "tracked insert adjacent text" (toString p.Buffer) "abcx"
    check "tracked insert adjacent id" (getStyleAt 3 p.Buffer).RevisionId (Some "7")
| None ->
    fail <- fail + 1
    printfn "FAIL %-28s expected paragraph" "tracked insert adjacent"
let trackedNonAdjacent =
    Doc.insertTrackedText "author" "date" "x"
        { trackedInsertBaseDoc with Selection = Some { Anchor = pos "tracked" 0; Focus = pos "tracked" 0 } }
match Doc.findParagraph trackedNonAdjacent "tracked" with
| Some p ->
    check "tracked insert new id text" (toString p.Buffer) "xabc"
    check "tracked insert new id" (getStyleAt 0 p.Buffer).RevisionId (Some "8")
| None ->
    fail <- fail + 1
    printfn "FAIL %-28s expected paragraph" "tracked insert new id"
let trackedDeleteBaseDoc =
    { Doc.empty with
        Blocks = [ Doc.Paragraph (para "trackdel" (ofString "abc")) ] }
let trackedDeleteForward =
    Doc.trackDeleteForward "author" "date"
        { trackedDeleteBaseDoc with Selection = Some { Anchor = pos "trackdel" 1; Focus = pos "trackdel" 1 } }
match Doc.findParagraph trackedDeleteForward "trackdel", trackedDeleteForward.Selection with
| Some p, Some sel ->
    check "tracked delete forward type" (getStyleAt 1 p.Buffer).RevisionType (Some "delete")
    check "tracked delete forward cursor" sel.Focus.Offset 1
| _ ->
    fail <- fail + 1
    printfn "FAIL %-28s expected paragraph selection" "tracked delete forward"
let trackedDeleteBackward =
    Doc.trackDeleteBackward "author" "date"
        { trackedDeleteBaseDoc with Selection = Some { Anchor = pos "trackdel" 2; Focus = pos "trackdel" 2 } }
match Doc.findParagraph trackedDeleteBackward "trackdel", trackedDeleteBackward.Selection with
| Some p, Some sel ->
    check "tracked delete backward type" (getStyleAt 1 p.Buffer).RevisionType (Some "delete")
    check "tracked delete backward cursor" sel.Focus.Offset 1
| _ ->
    fail <- fail + 1
    printfn "FAIL %-28s expected paragraph selection" "tracked delete backward"
let trackedInsertParagraphCount = 200
let trackedInsertBlockId = "tracked-small-00000"
let mutable trackedInsertDoc =
    { Doc.empty with
        Blocks =
            [ for index in 0 .. trackedInsertParagraphCount - 1 ->
                Doc.Paragraph (para (sprintf "tracked-small-%05d" index) (ofString "plain text")) ] }
for _ in 1 .. 20 do
    trackedInsertDoc <-
        Doc.insertTrackedText "author" "date" "x"
            { trackedInsertDoc with
                Selection = Some { Anchor = pos trackedInsertBlockId 0; Focus = pos trackedInsertBlockId 0 } }
let trackedInsertRevisionId =
    Doc.findParagraph trackedInsertDoc trackedInsertBlockId
    |> Option.map (fun p -> (getStyleAt 0 p.Buffer).RevisionId)
    |> Option.defaultValue None
check "tracked insert revision numbering" trackedInsertRevisionId (Some "20")
// removed: tested deleted cache/perf machinery

// removed: tested deleted cache/perf machinery

let tableRow cells =
    { (singleCellTable "template-table" (tableCell "template-cell" "")).Rows.Head with Cells = cells }
let tableInsertRowDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Table
                { (singleCellTable "insert-row-table" (tableCell "ir0c0" "a")) with
                    Rows =
                        [ tableRow [ tableCell "ir0c0" "a"; tableCell "ir0c1" "b" ]
                          tableRow [ tableCell "ir1c0" "c"; tableCell "ir1c1" "d" ] ] } ]
        Selection = Some { Anchor = pos "ir1c1" 0; Focus = pos "ir1c1" 0 } }
let insertedTableRowDoc = Doc.insertTableRowAfterSelection tableInsertRowDoc
match insertedTableRowDoc.Blocks with
| [ Doc.Table table ] ->
    check "table insert row count" table.Rows.Length 3
    check "table insert row width" table.Rows.[2].Cells.Length 2
    check "table insert row empty text" (toString table.Rows.[2].Cells.[1].Content.Head.Buffer) ""
| blocks ->
    fail <- fail + 1
    printfn "FAIL %-28s expected one table got %A" "table insert row shape" blocks.Length
let multiParagraphCell =
    { tableCell "mp-first" "a" with
        Content = [ para "mp-first" (ofString "a"); para "mp-second" (ofString "b") ] }
let tableInsertRowSecondParagraphDoc =
    { Doc.empty with
        Blocks = [ Doc.Table (singleCellTable "insert-row-second-para-table" multiParagraphCell) ]
        Selection = Some { Anchor = pos "mp-second" 0; Focus = pos "mp-second" 0 } }
match Doc.insertTableRowAfterSelection tableInsertRowSecondParagraphDoc with
| { Blocks = [ Doc.Table table ] } ->
    check "table lookup second para row" table.Rows.Length 2
| result ->
    fail <- fail + 1
    printfn "FAIL %-28s expected table got %A" "table lookup second para row" result.Blocks.Length
// removed: tested deleted cache/perf machinery
let tableDeleteRowDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Table
                { (singleCellTable "delete-row-table" (tableCell "dr0c0" "a")) with
                    Rows =
                        [ tableRow [ tableCell "dr0c0" "a"; tableCell "dr0c1" "b" ]
                          tableRow [ tableCell "dr1c0" "c"; tableCell "dr1c1" "d" ]
                          tableRow [ tableCell "dr2c0" "e"; tableCell "dr2c1" "f" ] ] } ]
        Selection = Some { Anchor = pos "dr2c1" 0; Focus = pos "dr2c1" 0 } }
let deletedTableRowDoc = Doc.deleteSelectedTableRow tableDeleteRowDoc
match deletedTableRowDoc.Blocks, deletedTableRowDoc.Selection with
| [ Doc.Table table ], Some sel ->
    check "table delete row count" table.Rows.Length 2
    check "table delete row removed" (table.Rows |> List.exists (fun row -> row.Cells |> List.exists (fun cell -> cell.Content.Head.Id = "dr2c1"))) false
    check "table delete row cursor" sel.Focus.BlockId "dr1c1"
| blocks, selection ->
    fail <- fail + 1
    printfn "FAIL %-28s expected table selection got blocks=%A selection=%A" "table delete row shape" blocks.Length selection
let tableDeleteColumnDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Table
                { (singleCellTable "delete-column-table" (tableCell "dc0c0" "a")) with
                    Rows =
                        [ tableRow [ tableCell "dc0c0" "a"; tableCell "dc0c1" "b"; tableCell "dc0c2" "c" ]
                          tableRow [ tableCell "dc1c0" "d"; tableCell "dc1c1" "e"; tableCell "dc1c2" "f" ] ] } ]
        Selection = Some { Anchor = pos "dc1c2" 0; Focus = pos "dc1c2" 0 } }
let deletedTableColumnDoc = Doc.deleteSelectedTableColumn tableDeleteColumnDoc
match deletedTableColumnDoc.Blocks, deletedTableColumnDoc.Selection with
| [ Doc.Table table ], Some sel ->
    check "table delete column width" table.Rows.[0].Cells.Length 2
    check "table delete column removed" (table.Rows.[1].Cells |> List.exists (fun cell -> cell.Content.Head.Id = "dc1c2")) false
    check "table delete column cursor" sel.Focus.BlockId "dc1c1"
| blocks, selection ->
    fail <- fail + 1
    printfn "FAIL %-28s expected table selection got blocks=%A selection=%A" "table delete column shape" blocks.Length selection

let tableEditSize = 20
let tableEditWideFocus = sprintf "table-edit-wide-0-%05d" (tableEditSize - 1)
let tableEditWideDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Table
                { (singleCellTable "table-edit-wide" (tableCell "table-edit-wide-template" "")) with
                    Rows =
                        [ tableRow
                            [ for col in 0 .. tableEditSize - 1 ->
                                tableCell (sprintf "table-edit-wide-0-%05d" col) "cell" ] ] } ]
        Selection = Some { Anchor = pos tableEditWideFocus 0; Focus = pos tableEditWideFocus 0 } }
let tableEditTallFocus = sprintf "table-edit-tall-%05d-0" (tableEditSize - 1)
let tableEditTallDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Table
                { (singleCellTable "table-edit-tall" (tableCell "table-edit-tall-template" "")) with
                    Rows =
                        [ for row in 0 .. tableEditSize - 1 ->
                            tableRow [ tableCell (sprintf "table-edit-tall-%05d-0" row) "cell" ] ] } ]
        Selection = Some { Anchor = pos tableEditTallFocus 0; Focus = pos tableEditTallFocus 0 } }
let tableEditInsertedColumn = Doc.insertTableColumnAfterSelection tableEditWideDoc
let tableEditDeletedColumn = Doc.deleteSelectedTableColumn tableEditWideDoc
let tableEditInsertedRow = Doc.insertTableRowAfterSelection tableEditTallDoc
let tableEditDeletedRow = Doc.deleteSelectedTableRow tableEditTallDoc
let tableEditValid =
    match tableEditInsertedColumn.Blocks, tableEditDeletedColumn.Blocks, tableEditInsertedRow.Blocks, tableEditDeletedRow.Blocks with
    | [ Doc.Table insertedColumn ], [ Doc.Table deletedColumn ], [ Doc.Table insertedRow ], [ Doc.Table deletedRow ] ->
        insertedColumn.Rows.Head.Cells.Length = tableEditSize + 1
        && deletedColumn.Rows.Head.Cells.Length = tableEditSize - 1
        && insertedRow.Rows.Length = tableEditSize + 1
        && deletedRow.Rows.Length = tableEditSize - 1
    | _ -> false
check "table edit wide/tall result" tableEditValid true
let tableEditTallColumnFocus = sprintf "table-edit-tall-column-%05d-1" (tableEditSize - 1)
let tableEditTallColumnDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Table
                { (singleCellTable "table-edit-tall-column" (tableCell "table-edit-tall-column-template" "")) with
                    Rows =
                        [ for row in 0 .. tableEditSize - 1 ->
                            tableRow
                                [ tableCell (sprintf "table-edit-tall-column-%05d-0" row) "cell"
                                  tableCell (sprintf "table-edit-tall-column-%05d-1" row) "cell" ] ] } ]
        Selection = Some { Anchor = pos tableEditTallColumnFocus 0; Focus = pos tableEditTallColumnFocus 0 } }
match Doc.deleteSelectedTableColumn tableEditTallColumnDoc with
| { Blocks = [ Doc.Table table ]; Selection = Some sel } ->
    check "table edit tall delete column width" table.Rows.[0].Cells.Length 1
    check "table edit tall delete column cursor" sel.Focus.BlockId (sprintf "table-edit-tall-column-%05d-0" (tableEditSize - 1))
| doc ->
    fail <- fail + 1
    printfn "FAIL %-28s expected table selection got blocks=%A" "table edit tall delete column" doc.Blocks.Length
// removed: tested deleted cache/perf machinery

let tableTailPrefixCount = 20
let tableTailPrefixBlocks =
    [ for index in 0 .. tableTailPrefixCount - 1 ->
        Doc.Paragraph (para (sprintf "table-tail-prefix-%05d" index) (ofString "p")) ]
let tableTailDoc =
    { Doc.empty with
        Blocks =
            [ yield! tableTailPrefixBlocks
              yield Doc.Table { (singleCellTable "table-tail-table" (tableCell "table-tail-cell" "cell")) with Rows = [ tableRow [ tableCell "table-tail-cell" "cell" ] ] } ]
        Selection = Some { Anchor = pos "table-tail-cell" 0; Focus = pos "table-tail-cell" 0 } }
match Doc.insertTableRowAfterSelection tableTailDoc with
| { Blocks = blocks } ->
    match blocks |> List.rev |> List.tryHead with
    | Some (Doc.Table table) -> check "table tail insert rows" table.Rows.Length 2
    | _ ->
        fail <- fail + 1
        printfn "FAIL %-28s expected table at tail" "table tail insert rows"
// removed: tested deleted cache/perf machinery

// removed: tested deleted cache/perf machinery

let insertRevisionDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph (para "r1" (revisionBuffer "insert"))
              Doc.Paragraph (para "r2" (revisionBuffer "insert")) ] }
let acceptedInsertDoc = Doc.acceptAllRevisions insertRevisionDoc
check "accept insert revision text" (Doc.getText acceptedInsertDoc) "aREVb\naREVb"
check "accept insert revision cleared" (Doc.revisionInfos acceptedInsertDoc |> List.length) 0
let rejectedInsertDoc = Doc.rejectAllRevisions insertRevisionDoc
check "reject insert revision text" (Doc.getText rejectedInsertDoc) "ab\nab"
check "reject insert revision cleared" (Doc.revisionInfos rejectedInsertDoc |> List.length) 0
let selectedInsertRevisionDoc =
    { insertRevisionDoc with Selection = Some { Anchor = pos "r1" 1; Focus = pos "r2" 4 } }
let acceptedSelectedInsertDoc = Doc.acceptSelectedRevision selectedInsertRevisionDoc
check "accept selected insert text" (Doc.getText acceptedSelectedInsertDoc) "aREVb\naREVb"
check "accept selected insert cleared" (Doc.revisionInfos acceptedSelectedInsertDoc |> List.length) 0
let rejectedSelectedInsertDoc = Doc.rejectSelectedRevision selectedInsertRevisionDoc
check "reject selected insert text" (Doc.getText rejectedSelectedInsertDoc) "ab\nab"
check "reject selected insert cleared" (Doc.revisionInfos rejectedSelectedInsertDoc |> List.length) 0
let fullInsertRevisionDoc =
    { Doc.empty with
        Blocks = [ Doc.Paragraph (para "fr1" (ofString "REV" |> setRevision 0 3 "insert" "r2" None None)) ] }
let rejectedFullInsertDoc = Doc.rejectAllRevisions fullInsertRevisionDoc
check "reject full insert revision text" (Doc.getText rejectedFullInsertDoc) ""
check "reject full insert revision cleared" (Doc.revisionInfos rejectedFullInsertDoc |> List.length) 0

let deleteRevisionDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph (para "d1" (revisionBuffer "delete"))
              Doc.Paragraph (para "d2" (revisionBuffer "delete")) ] }
let acceptedDeleteDoc = Doc.acceptAllRevisions deleteRevisionDoc
check "accept delete revision text" (Doc.getText acceptedDeleteDoc) "ab\nab"
check "accept delete revision cleared" (Doc.revisionInfos acceptedDeleteDoc |> List.length) 0
let rejectedDeleteDoc = Doc.rejectAllRevisions deleteRevisionDoc
check "reject delete revision text" (Doc.getText rejectedDeleteDoc) "aREVb\naREVb"
check "reject delete revision cleared" (Doc.revisionInfos rejectedDeleteDoc |> List.length) 0

// Splitting keeps an inline image as an image, not a bare placeholder character.
let splitImageDoc =
    { Doc.empty with
        Blocks = [ Doc.Paragraph (para "img" (ofString "ab" |> insertImage 2 "pic" (Some 10.0) (Some 10.0) None)) ]
        Selection = Some { Anchor = pos "img" 1; Focus = pos "img" 1 } }
match Doc.splitBlock splitImageDoc with
| { Blocks = [ _; Doc.Paragraph after ] } -> check "split keeps image run" (imageSources after.Buffer) [ "pic" ]
| doc ->
    fail <- fail + 1
    printfn "FAIL %-28s expected two paragraphs got %d" "split image shape" doc.Blocks.Length

// Text typed after an inline image lands after it in the runs as well as the rope.
let afterImageBuffer =
    ofString "ab"
    |> insertImage 2 "pic" (Some 10.0) (Some 10.0) None
    |> insert 3 "cd"
check "insert after image text" (toString afterImageBuffer) "ab\uFFFCcd"
check "insert after image runs"
    (afterImageBuffer.Runs |> List.map (function Image img -> "I", img.Start, 1 | Text tr -> "T", tr.Start, tr.Length))
    [ "T", 0, 2; "I", 2, 1; "T", 3, 2 ]

// Deleting across paragraphs keeps the surviving runs on both sides.
let crossDeleteDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph (para "d1" styledBuffer)
              Doc.Paragraph (para "d2" styledBuffer) ]
        Selection = Some { Anchor = pos "d1" 1; Focus = pos "d2" 3 } }
match Doc.deleteSelection crossDeleteDoc with
| { Blocks = [ Doc.Paragraph merged ] } ->
    check "cross delete text" (toString merged.Buffer) "ad"
    check "cross delete keeps first style" (getStyleAt 0 merged.Buffer).Bold true
    check "cross delete keeps last style" (getStyleAt 1 merged.Buffer).Color (Some "#ff0000")
| doc ->
    fail <- fail + 1
    printfn "FAIL %-28s expected one paragraph got %d" "cross delete shape" doc.Blocks.Length

// A table neighbour is stepped over or left alone rather than throwing.
let besideTableDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph (para "above" (ofString "above"))
              Doc.Table (singleCellTable "tbl" (tableCell "cell" "x"))
              Doc.Paragraph (para "below" (ofString "below")) ]
        Selection = Some { Anchor = pos "below" 0; Focus = pos "below" 0 } }
check "backspace below table is a no-op" (Doc.mergeWithPrevious besideTableDoc "below").Blocks.Length 3
check "delete above table is a no-op" (Doc.mergeWithNext besideTableDoc "above").Blocks.Length 3
check "arrow up steps over table"
    ((Doc.moveCursorUp besideTableDoc).Selection |> Option.map (fun sel -> sel.Focus.BlockId, sel.Focus.Offset))
    (Some ("above", 5))
check "arrow down steps over table"
    ((Doc.moveCursorDown { besideTableDoc with Selection = Some { Anchor = pos "above" 0; Focus = pos "above" 0 } }).Selection
     |> Option.map (fun sel -> sel.Focus.BlockId))
    (Some "below")

// Text typed at the edge of an annotated run takes its formatting, not its identity.
let caretDoc id buffer offset =
    { Doc.empty with
        Blocks = [ Doc.Paragraph (para id buffer) ]
        Selection = Some { Anchor = pos id offset; Focus = pos id offset } }
let paraText id doc =
    Doc.findParagraph doc id |> Option.map (fun p -> toString p.Buffer) |> Option.defaultValue "<missing>"
let afterFieldDoc =
    caretDoc "fld" (ofString "Page ") 5
    |> Doc.insertField "PAGE" "1"
    |> Doc.insertText "Z"
check "typing after field text" (paraText "fld" afterFieldDoc) "Page 1Z"
check "typing after field is not field"
    (Doc.findParagraph afterFieldDoc "fld" |> Option.map (fun p -> (getStyleAt 5 p.Buffer).FieldCode, (getStyleAt 6 p.Buffer).FieldCode))
    (Some (Some "PAGE", None))
check "typing after field survives refresh" (paraText "fld" (Doc.refreshFields afterFieldDoc)) "Page 1Z"
let annotated =
    { defaultStyle with
        Bold = true
        Hyperlink = Some "http://x"
        CommentId = Some "c1"
        CommentText = Some "note"
        NoteId = Some "n1"
        BookmarkStartId = Some "b1"
        BookmarkName = Some "bm"
        BookmarkEndId = Some "b1"
        RevisionType = Some "delete"
        RevisionId = Some "r1" }
let annotatedBuffer = { Rope = Rbt.Word.Rope.ofString "link"; Runs = [ Text { Start = 0; Length = 4; Style = annotated } ] }
let afterAnnotated = getStyleAt 4 (insert 4 "x" annotatedBuffer)
check "typing after annotated run keeps formatting" afterAnnotated.Bold true
check "typing after annotated run drops identity"
    (afterAnnotated.Hyperlink, afterAnnotated.CommentId, afterAnnotated.CommentText, afterAnnotated.NoteId, afterAnnotated.BookmarkName, afterAnnotated.BookmarkEndId, afterAnnotated.RevisionType)
    (None, None, None, None, None, None, None)
let beforeAnnotated = getStyleAt 0 (insert 0 "x" annotatedBuffer)
check "typing before annotated run drops identity" (beforeAnnotated.Bold, beforeAnnotated.Hyperlink, beforeAnnotated.NoteId) (true, None, None)
let insideAnnotated = insert 2 "x" annotatedBuffer
check "typing inside hyperlink stays in it" (getStyleAt 2 insideAnnotated).Hyperlink (Some "http://x")
check "typing inside revision is not a revision" (getStyleAt 2 insideAnnotated).RevisionType None
let twoRunLink =
    { Rope = Rbt.Word.Rope.ofString "abcd"
      Runs =
        [ Text { Start = 0; Length = 2; Style = { defaultStyle with Bold = true; Hyperlink = Some "http://x" } }
          Text { Start = 2; Length = 2; Style = { defaultStyle with Hyperlink = Some "http://x" } } ] }
check "typing between runs of one hyperlink stays in it" (getStyleAt 2 (insert 2 "x" twoRunLink)).Hyperlink (Some "http://x")
check "typing after last hyperlink run leaves it" (getStyleAt 4 (insert 4 "x" twoRunLink)).Hyperlink None
let replaceAfterField =
    { Rope = Rbt.Word.Rope.ofString "1abc"
      Runs =
        [ Text { Start = 0; Length = 1; Style = { defaultStyle with Italic = true; FieldCode = Some "PAGE" } }
          Text { Start = 1; Length = 3; Style = defaultStyle } ] }
let replacedAfterField =
    let ranges = ResizeArray<struct (int * int)>()
    ranges.Add(struct (1, 4))
    replaceRanges ranges "Z" replaceAfterField
check "replace after field is not field" ((getStyleAt 1 replacedAfterField).FieldCode, (getStyleAt 1 replacedAfterField).Italic) (None, true)

// removed: command-history retention/budget tests targeted the deleted cache machinery (RetainedSourceTextCharCount, trimHistory)
printfn ""
printfn "Word textbuffer tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
