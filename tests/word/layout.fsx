#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"

open Rbt.Word.TextBuffer
open Rbt.Word.Layout
open Rbt.Word.Paint
module Doc = Rbt.Word.Document
module Styles = Rbt.Word.WordStyles

let mutable pass = 0
let mutable fail = 0

let check label got expected =
    if got = expected then
        pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-28s expected %A got %A" label expected got

let settings : LayoutSettings =
    { Width = 240.0
      PageHeight = 1000.0
      FontSize = 16.0
      LineHeight = 24.0
      PaddingLeft = 0.0
      PaddingTop = 0.0
      MarginRight = 0.0
      MarginBottom = 0.0
      InTableCell = false
      TableContinuationHeaderHeight = 0.0
      LongTrailingSpaceFitSafety = 0.25 }

let measureText (text: string) (_style: TextStyle) =
    float text.Length * 7.0

let textOfSegment = function
    | TextSegment (text, _, _, _, _) -> text
    | ImageSegment _ -> "\uFFFC"
    | MathSegment _ -> ""

let lines =
    layoutLinesStyled settings measureText (ofString "alpha beta gamma delta") 90.0

check "wrapped line count" lines.Length 2
check "wrapped line text 1" (lines.[0].Segments |> List.map textOfSegment |> String.concat "") "alpha beta "
check "wrapped line text 2" (lines.[1].Segments |> List.map textOfSegment |> String.concat "") "gamma delta"
check "wrapped coalesced line 1" lines.[0].Segments.Length 1
check "wrapped coalesced line 2" lines.[1].Segments.Length 1

let wideCoalesced =
    layoutLinesStyled settings measureText (ofString (String.concat " " [ for _ in 1 .. 20 -> "word" ])) 2000.0
check "wide line coalesced count" wideCoalesced.[0].Segments.Length 1
check "wide line coalesced text"
    (wideCoalesced.[0].Segments |> List.map textOfSegment |> String.concat "")
    (String.concat " " [ for _ in 1 .. 20 -> "word" ])

let fragmented =
    { Rope = Rbt.Word.Rope.ofString "alpha beta gamma"
      Runs =
        [ Text { Start = 0; Length = 6; Style = defaultStyle }
          Text { Start = 6; Length = 4; Style = { defaultStyle with Bold = true } }
          Text { Start = 10; Length = 6; Style = defaultStyle } ] }
let fragmentedLines = layoutLinesStyled settings measureText fragmented 90.0
check "fragmented line count" fragmentedLines.Length 2
check "fragmented line text 1" (fragmentedLines.[0].Segments |> List.map textOfSegment |> String.concat "") "alpha beta "
check "fragmented line text 2" (fragmentedLines.[1].Segments |> List.map textOfSegment |> String.concat "") "gamma"
let reusePara : Doc.ParagraphBlock =
    { Id = "reuse"; Buffer = ofString "plain text"; Metadata = Doc.defaultMetadata }
let reuseLayout, _ = layoutParagraphAtWithRange settings measureText (rangeMeasurerFromTextMeasurer measureText) 500.0 0.0 0.0 reusePara Map.empty
check "layout zero-offset keeps segment" (match reuseLayout.Lines.[0].Segments with TextSegment (_, _, x, _, _) :: _ -> x | _ -> -1.0) 0.0
let indentedReusePara = { reusePara with Metadata = { reusePara.Metadata with IndentLeft = Some 12.0 } }
let indentedReuseLayout, _ = layoutParagraphAtWithRange settings measureText (rangeMeasurerFromTextMeasurer measureText) 500.0 0.0 0.0 indentedReusePara Map.empty
check "layout indent shifts segment" (match indentedReuseLayout.Lines.[0].Segments with TextSegment (_, _, x, _, _) :: _ -> x | _ -> -1.0) 12.0

let effectiveMerged =
    Styles.effectiveBuffer Styles.normal
        { Rope = Rbt.Word.Rope.ofString "abcd"
          Runs =
            [ Text { Start = 0; Length = 2; Style = defaultStyle }
              Text { Start = 2; Length = 2; Style = defaultStyle } ] }
check "effective buffer merges" effectiveMerged.Runs.Length 1
let effectiveSingle = ofString "single run"
let noOpParagraphStyle = { Styles.normal with Text = defaultStyle }
check "effective buffer single reuse" (obj.ReferenceEquals(Styles.effectiveBuffer noOpParagraphStyle effectiveSingle, effectiveSingle)) true
let boldParagraphStyle = { Styles.normal with Text = { Styles.normal.Text with Bold = true } }
let effectiveSingleBold = Styles.effectiveBuffer boldParagraphStyle effectiveSingle
check "effective buffer single changed" (obj.ReferenceEquals(effectiveSingleBold, effectiveSingle)) false
check "effective buffer single bold" (getStyleAt 0 effectiveSingleBold).Bold true

let namedStyle id basedOn textStyle : Doc.NamedStyle =
    { Id = id
      Name = Some id
      StyleType = "character"
      IsDefault = false
      BasedOn = basedOn
      Next = None
      Linked = None
      IsQuickStyle = false
      Aliases = None
      UiPriority = None
      SemiHidden = false
      UnhideWhenUsed = false
      Text = textStyle
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
      PageBreakBefore = None
      WidowControl = None }

let styleDoc =
    { Doc.empty with
        Styles =
            [ namedStyle "base-char" None (Some { defaultStyle with Bold = true })
              namedStyle "child-char" (Some "base-char") (Some { defaultStyle with Color = Some "#111111" }) ] }
let inheritedStyleBuffer =
    Styles.effectiveBufferInDocument styleDoc Styles.normal
        { Rope = Rbt.Word.Rope.ofString "ab"
          Runs =
            [ Text { Start = 0; Length = 2; Style = { defaultStyle with Style = Some "child-char" } } ] }
let inheritedStyle =
    match inheritedStyleBuffer.Runs with
    | [ Text run ] -> run.Style
    | _ -> defaultStyle
check "effective document char basedOn bold" inheritedStyle.Bold true
check "effective document char basedOn color" inheritedStyle.Color (Some "#111111")
// removed: tested deleted cache machinery

let blockLayout id y line : BlockLayout =
    { BlockId = id
      Lines = [ { line with Y = y } ]
      Y = y
      Height = line.Height
      AvailableWidth = 120.0
      Metadata = Rbt.Word.Document.defaultMetadata
      Prefix = ""
      PrefixStyle = None
      PrefixWidth = 0.0
      Indent = 0.0
      VAlignOffset = 0.0 }

let hitLine =
    { Segments = [ TextSegment ("abcdef", defaultStyle, 0.0, 42.0, 0) ]
      Y = 0.0
      Width = 42.0
      Height = 24.0
      StartOffset = 0
      EndOffset = 6 }
let hitLayouts =
    [ blockLayout "h1" 0.0 hitLine
      blockLayout "h2" 24.0 hitLine
      blockLayout "h3" 48.0 hitLine ]
let hitArray = hitLayouts |> List.toArray
check "hittest array middle" (hitTestLayoutArrayWithRange (rangeMeasurerFromTextMeasurer measureText) 0.0 hitArray 14.0 26.0) (hitTestLayoutsWithRange (rangeMeasurerFromTextMeasurer measureText) 0.0 hitLayouts 14.0 26.0)
check "hittest array last" (hitTestLayoutArrayWithRange (rangeMeasurerFromTextMeasurer measureText) 0.0 hitArray 14.0 50.0) (hitTestLayoutsWithRange (rangeMeasurerFromTextMeasurer measureText) 0.0 hitLayouts 14.0 50.0)
let pageInfoHighLine = { hitLine with Y = 1300.0 }
let nonMonotonicPageLayouts =
    [ { blockLayout "page-info-high" 1300.0 pageInfoHighLine with Height = pageInfoHighLine.Height }
      blockLayout "page-info-low" 0.0 hitLine ]
check "uniform page infos use max line bottom"
    ((uniformPageInfos settings nonMonotonicPageLayouts).Length > 1)
    true
let countedPageInfos, countedLayoutCount = uniformPageInfosWithLayoutCount settings nonMonotonicPageLayouts
check "uniform page infos counted count" countedLayoutCount nonMonotonicPageLayouts.Length
check "uniform page infos counted pages" countedPageInfos (uniformPageInfos settings nonMonotonicPageLayouts)

let lineLookupLines =
    [ for i in 0 .. 999 ->
        { hitLine with
            Y = float i * 24.0
            StartOffset = i * 10
            EndOffset = (i + 1) * 10 } ]
let lineLookupLayout =
    { (blockLayout "line-lookup" 0.0 hitLine) with
        Lines = lineLookupLines
        Height = 24000.0 }
check "line lookup first" (lineForOffset lineLookupLayout 0).StartOffset 0
check "line lookup boundary next" (lineForOffset lineLookupLayout 10).StartOffset 10
check "line lookup last" (lineForOffset lineLookupLayout 9999).StartOffset 9990
// removed: tested deleted cache machinery

let imageLine =
    { hitLine with
        Segments = [ ImageSegment ("img", 7.0, 0.0, 32.0, 24.0, 0, false, None, None, None, None, None, None) ]
        Width = 32.0
        EndOffset = 1 }
let imageLayouts =
    [ blockLayout "i1" 0.0 hitLine
      blockLayout "i2" 24.0 imageLine ]
let imageArray = imageLayouts |> List.toArray
check "image hittest array hit" (hitTestForImageLayoutArray 0.0 imageArray 14.0 26.0) (hitTestForImageLayouts 0.0 imageLayouts 14.0 26.0)
check "image hittest array miss" (hitTestForImageLayoutArray 0.0 imageArray 14.0 2.0) (hitTestForImageLayouts 0.0 imageLayouts 14.0 2.0)

let selectionDoc =
    { Doc.empty with
        Selection =
            Some {
                Anchor = { BlockId = "h3"; Offset = 2 }
                Focus = { BlockId = "h3"; Offset = 2 }
            } }
let layoutLookup blockId =
    hitArray |> Array.tryFind (fun layout -> layout.BlockId = blockId)
check "selection lookup command"
    (selectionCommandsWithRangeLookup settings (rangeMeasurerFromTextMeasurer measureText) selectionDoc layoutLookup)
    (selectionCommandsWithRange settings (rangeMeasurerFromTextMeasurer measureText) selectionDoc hitLayouts)
let imageLookup blockId =
    imageArray |> Array.tryFind (fun layout -> layout.BlockId = blockId)
check "hover lookup command"
    (hoverCommandsLookup settings imageLookup (Some ("i2", 0)))
    (hoverCommands settings imageLayouts (Some ("i2", 0)))
let hoverLookupPerfLayouts =
    [ for i in 0 .. 19999 ->
        let baseLine =
            if i = 19999 then
                { imageLine with Y = float i * 24.0 }
            else
                { hitLine with Y = float i * 24.0 }
        blockLayout (sprintf "hover-lookup-%d" i) (float i * 24.0) baseLine ]
check "hover layout lookup command count"
    (hoverCommands settings hoverLookupPerfLayouts (Some ("hover-lookup-19999", 0)) |> List.length)
    5
// removed: tested deleted cache machinery

// removed: tested deleted cache machinery

let headerPageLayout =
    { blockLayout "page-body" 0.0 hitLine with
        Lines = [ { hitLine with Y = pageHeight settings + pageGap + 12.0 } ]
        Height = pageHeight settings + pageGap + 36.0 }

let paragraph id text : Doc.ParagraphBlock =
    { Id = id
      Buffer = ofString text
      Metadata = Doc.defaultMetadata }

let tableCell content merge : Doc.TableCell =
    { Content = content
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
      VerticalMerge = merge
      PaddingTop = None
      PaddingRight = None
      PaddingBottom = None
      PaddingLeft = None }

let tableRow cells : Doc.TableRow =
    { Cells = cells
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

let verticalMergeRows =
    Rbt.Word.WordLayout.Table.distributeVerticalMergeContent
        [ tableRow [ tableCell [ paragraph "vm-1" "one"; paragraph "vm-2" "two"; paragraph "vm-3" "three" ] (Some Doc.RestartMerge) ]
          tableRow [ tableCell [ paragraph "vm-empty-1" "" ] (Some Doc.ContinueMerge) ]
          tableRow [ tableCell [ paragraph "vm-empty-2" "" ] (Some Doc.ContinueMerge) ] ]
let verticalMergeText rowIndex =
    verticalMergeRows.[rowIndex].Cells.[0].Content
    |> List.map (fun para -> toString para.Buffer)
check "vertical merge restart content" (verticalMergeText 0) [ "one" ]
check "vertical merge continue content 1" (verticalMergeText 1) [ "two" ]
check "vertical merge continue content 2" (verticalMergeText 2) [ "three" ]
let verticalMergeDistributionColumnCount = 16000
let verticalMergeDistributionRows =
    [ tableRow
        [ for col in 0 .. verticalMergeDistributionColumnCount - 1 ->
            tableCell
                [ paragraph (sprintf "vm-dist-%05d-a" col) "a"
                  paragraph (sprintf "vm-dist-%05d-b" col) "b" ]
                (Some Doc.RestartMerge) ]
      tableRow
        [ for col in 0 .. verticalMergeDistributionColumnCount - 1 ->
            tableCell [ paragraph (sprintf "vm-dist-%05d-empty" col) "" ] (Some Doc.ContinueMerge) ] ]
let verticalMergeDistributionSw = System.Diagnostics.Stopwatch.StartNew()
let verticalMergeDistributionResolved =
    Rbt.Word.WordLayout.Table.distributeVerticalMergeContent verticalMergeDistributionRows
verticalMergeDistributionSw.Stop()
let verticalMergeDistributionFirst =
    verticalMergeDistributionResolved.[1].Cells.[0].Content |> List.map (fun para -> toString para.Buffer)
let verticalMergeDistributionLast =
    verticalMergeDistributionResolved.[1].Cells.[verticalMergeDistributionColumnCount - 1].Content |> List.map (fun para -> toString para.Buffer)
check "vertical merge wide distribution"
    (verticalMergeDistributionFirst, verticalMergeDistributionLast)
    ([ "b" ], [ "b" ])
check "vertical merge distribution perf"
    (verticalMergeDistributionSw.Elapsed.TotalMilliseconds < 200.0)
    true

let wideFastRow =
    [ for i in 0 .. 31 -> tableCell [ paragraph $"wide-{i}" "x" ] None ]
    |> tableRow
let wideFastTable : Doc.TableBlock =
    { Id = "wide-fast"
      Rows = [ wideFastRow ]
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
let nestedAnchorTableCount = 2500
let nestedAnchorTables =
    [ for index in 0 .. nestedAnchorTableCount - 1 ->
        { wideFastTable with
            Id = sprintf "nested-anchor-%04d" index
            Rows = [ tableRow [ tableCell [ paragraph (sprintf "nested-anchor-p-%04d" index) "" ] None ] ] } ]
let nestedAnchorCell =
    { tableCell [ paragraph "nested-anchor-root" "" ] None with
        NestedTables = nestedAnchorTables
        NestedTableAnchors = [ for index in 0 .. nestedAnchorTableCount - 1 -> if index % 2 = 0 then 0 else 1 ] }
let nestedAnchorTable =
    { wideFastTable with
        Id = "nested-anchor-parent"
        Rows = [ tableRow [ nestedAnchorCell ] ] }
let nestedAnchorSettings = { settings with Width = 100000.0; PageHeight = 100000.0 }
let nestedAnchorSw = System.Diagnostics.Stopwatch.StartNew()
let nestedAnchorPlan =
    buildTableLayoutPlanForDocumentAt nestedAnchorSettings measureText (rangeMeasurerFromTextMeasurer measureText) Doc.empty nestedAnchorTable 0.0
nestedAnchorSw.Stop()
check "nested table anchors perf"
    (nestedAnchorSw.Elapsed.TotalMilliseconds < 300.0 && (tableCellFragmentPlans nestedAnchorPlan |> List.length) > 0)
    true
let wideFastWidths = tableColumnWidths 3200.0 wideFastTable
let wideFastHeight =
    tryMeasureTableRowHeightWithStyle settings measureText wideFastTable 3200.0 wideFastWidths wideFastRow (fun _ -> Styles.normal) (fun style textStyle -> Styles.effectiveTextStyle style textStyle)
check "wide table fast height" (wideFastHeight |> Option.map (fun height -> height > 0.0)) (Some true)
let wideResolved = Rbt.Word.WordLayout.Table.resolveTable Rbt.Word.WordLayout.canvasScale (Some 3200.0) wideFastTable
check "wide table resolve cells" wideResolved.BodyRows.[0].Cells.Length 32
let wideResolvedHasBorders =
    wideResolved.BodyRows
    |> List.exists (fun row ->
        row.Cells
        |> List.exists (fun cell ->
            cell.Borders.Top.IsSome
            || cell.Borders.Right.IsSome
            || cell.Borders.Bottom.IsSome
            || cell.Borders.Left.IsSome))
check "wide table no border inputs" wideResolvedHasBorders false
let vMergeResolveTable =
    { wideFastTable with
        Id = "resolve-vmerge"
        Rows =
            [ tableRow [ tableCell [ paragraph "resolve-vm-a" "a" ] (Some Doc.RestartMerge); tableCell [ paragraph "resolve-vm-b" "b" ] (Some Doc.RestartMerge) ]
              tableRow [ tableCell [ paragraph "resolve-vm-a2" "" ] (Some Doc.ContinueMerge); tableCell [ paragraph "resolve-vm-b2" "" ] (Some Doc.ContinueMerge) ]
              tableRow [ tableCell [ paragraph "resolve-vm-a3" "" ] (Some Doc.ContinueMerge); tableCell [ paragraph "resolve-vm-b3" "" ] (Some Doc.ContinueMerge) ] ] }
let vMergeResolved = Rbt.Word.WordLayout.Table.resolveTable Rbt.Word.WordLayout.canvasScale (Some 3200.0) vMergeResolveTable
check "vmerge resolve row spans" (vMergeResolved.BodyRows.[0].Cells |> List.map _.RowSpan) [ 3; 3 ]
check "vmerge resolve continuations" (vMergeResolved.BodyRows.[1].Cells.Length, vMergeResolved.BodyRows.[2].Cells.Length) (0, 0)
let vMergeLookupColumnCount = 16000
let vMergeLookupSettings = { settings with Width = 100000.0; PageHeight = 100000.0 }
let vMergeLookupCell col merge align =
    { tableCell [ paragraph (sprintf "vmerge-lookup-%d" col) "x" ] merge with VerticalAlignment = align }
let vMergeLookupTable =
    { wideFastTable with
        Id = "vmerge-lookup-perf"
        Rows =
            [ tableRow [ for col in 0 .. vMergeLookupColumnCount - 1 -> vMergeLookupCell col (Some Doc.RestartMerge) (Some "center") ]
              tableRow [ for col in 0 .. vMergeLookupColumnCount - 1 -> vMergeLookupCell col (Some Doc.ContinueMerge) None ] ] }
let vMergeLookupPlan =
    buildTableLayoutPlanForDocumentAt vMergeLookupSettings measureText (rangeMeasurerFromTextMeasurer measureText) Doc.empty vMergeLookupTable 0.0
let vMergeLookupSw = System.Diagnostics.Stopwatch.StartNew()
let vMergeLookupResolver = tableCellVerticalAlignmentHeightResolverForPlan vMergeLookupPlan
let mutable vMergeLookupTotal = 0.0
for col in 0 .. vMergeLookupColumnCount - 1 do
    vMergeLookupTotal <- vMergeLookupTotal + vMergeLookupResolver 0 col 24.0
vMergeLookupSw.Stop()
check "vmerge row-span lookup perf"
    (vMergeLookupSw.Elapsed.TotalMilliseconds < 200.0 && vMergeLookupTotal > 0.0)
    true
let vMergePaintRows = 256
let vMergePaintColumns = 256
let vMergePaintSettings = { settings with Width = 100000.0; PageHeight = 100000.0 }
let vMergePaintCell row col merge =
    tableCell [ paragraph (sprintf "vmerge-paint-%d-%d" row col) (if row = 0 then "x" else "") ] merge
let vMergePaintTable =
    { wideFastTable with
        Id = "vmerge-paint-perf"
        Rows =
            [ for row in 0 .. vMergePaintRows - 1 ->
                let merge = if row = 0 then Some Doc.RestartMerge else Some Doc.ContinueMerge
                tableRow [ for col in 0 .. vMergePaintColumns - 1 -> vMergePaintCell row col merge ] ] }
let vMergePaintSw = System.Diagnostics.Stopwatch.StartNew()
let _, vMergePaintNextY = tableBorderCommands vMergePaintSettings measureText vMergePaintTable 0.0
vMergePaintSw.Stop()
check "vmerge paint row-span perf"
    (vMergePaintSw.Elapsed.TotalMilliseconds < 800.0 && vMergePaintNextY > 0.0)
    true
let testBorder : Doc.TableBorder =
    { Style = Some "single"; Color = Some "#000000"; Width = Some 1.0 }
let testBorders : Doc.TableBorders =
    { Top = Some testBorder
      Right = Some testBorder
      Bottom = Some testBorder
      Left = Some testBorder
      InsideH = Some testBorder
      InsideV = Some testBorder }
let borderedResolveTable = { wideFastTable with Borders = Some testBorders; Rows = [ tableRow [ tableCell [ paragraph "border-left" "x" ] None; tableCell [ paragraph "border-right" "x" ] None ] ] }
let borderedResolve = Rbt.Word.WordLayout.Table.resolveTable Rbt.Word.WordLayout.canvasScale (Some 3200.0) borderedResolveTable
let borderedLeft = borderedResolve.BodyRows.[0].Cells.[0].Borders
let borderedRight = borderedResolve.BodyRows.[0].Cells.[1].Borders
check "bordered table shared border" (borderedLeft.Right.IsSome, borderedRight.Left.IsNone) (true, true)
let _, wideFastNextY = tableBorderCommands settings measureText wideFastTable 0.0
check "wide table border advance" (wideFastNextY > 0.0) true
let borderGridWideCellCount = 5000
let borderGridStroke : RedBat.BorderGrid.Stroke =
    { Color = "#000000"; Width = 1.0; Style = Some "single"; Priority = 1 }
let borderGridEdges : RedBat.BorderGrid.Edges =
    { Top = Some borderGridStroke
      Right = Some borderGridStroke
      Bottom = Some borderGridStroke
      Left = Some borderGridStroke }
let borderGridWideCells = ResizeArray<RedBat.BorderGrid.Cell>(borderGridWideCellCount)
for col in 0 .. borderGridWideCellCount - 1 do
    borderGridWideCells.Add
        { Row = 0
          Column = col
          RowSpan = 1
          ColumnSpan = 1
          X = float col * 10.0
          Y = 0.0
          Width = 10.0
          Height = 18.0
          Edges = borderGridEdges }
let borderGridWideSw = System.Diagnostics.Stopwatch.StartNew()
let borderGridWideSegments = RedBat.BorderGrid.resolveResizeArray borderGridWideCells
borderGridWideSw.Stop()
check "border grid wide row perf"
    (borderGridWideSw.Elapsed.TotalMilliseconds < 250.0
     && borderGridWideSegments.Length > borderGridWideCellCount)
    true
let wideFastPlan =
    buildTableLayoutPlanForDocumentAt settings measureText (rangeMeasurerFromTextMeasurer measureText) Doc.empty wideFastTable 0.0
check "wide table paint fragments" (tableCellPaintFragmentPlans wideFastPlan |> List.length) 32
let tallFastTableRowCount = 12000
let tallFastTable =
    { wideFastTable with
        Id = "tall-fast"
        Rows =
            [ for row in 0 .. tallFastTableRowCount - 1 ->
                tableRow [ tableCell [] None ] ] }
let tallFastSettings = { settings with Width = 1000.0; PageHeight = 10000000.0 }
let tallFastPlanSw = System.Diagnostics.Stopwatch.StartNew()
let tallFastPlan =
    buildTableLayoutPlanForDocumentAt tallFastSettings measureText (rangeMeasurerFromTextMeasurer measureText) Doc.empty tallFastTable 0.0
tallFastPlanSw.Stop()
let tallFastFragmentsSw = System.Diagnostics.Stopwatch.StartNew()
let tallFastFragments = tableCellPaintFragmentPlans tallFastPlan |> List.length
tallFastFragmentsSw.Stop()
let tallFastOk =
    tallFastPlan.Rows.Length = tallFastTableRowCount
    && tallFastFragments = tallFastTableRowCount
    && tallFastPlanSw.Elapsed.TotalMilliseconds + tallFastFragmentsSw.Elapsed.TotalMilliseconds < 1500.0
if not tallFastOk then
    printfn
        "tall table plan details: rows=%d fragments=%d plan=%.1fms fragments=%.1fms"
        tallFastPlan.Rows.Length
        tallFastFragments
        tallFastPlanSw.Elapsed.TotalMilliseconds
        tallFastFragmentsSw.Elapsed.TotalMilliseconds
check "tall table plan perf"
    tallFastOk
    true
let tallWideVirtualGridRows = 8000
let tallWideVirtualGridColumns = 512
let tallWideVirtualGridRow =
    { tableRow [ tableCell [] None ] with
        GridAfter = tallWideVirtualGridColumns - 1 }
let tallWideVirtualGridTable =
    { wideFastTable with
        Id = "tall-wide-virtual-grid"
        Rows = [ for _ in 1 .. tallWideVirtualGridRows -> tallWideVirtualGridRow ] }
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let tallWideVirtualGridBefore = System.GC.GetAllocatedBytesForCurrentThread()
let tallWideVirtualGridSw = System.Diagnostics.Stopwatch.StartNew()
let tallWideVirtualGridPlan =
    buildTableLayoutPlanForDocumentAt tallFastSettings measureText (rangeMeasurerFromTextMeasurer measureText) Doc.empty tallWideVirtualGridTable 0.0
let tallWideVirtualGridFragments =
    tableCellPaintFragmentPlans tallWideVirtualGridPlan |> List.length
tallWideVirtualGridSw.Stop()
let tallWideVirtualGridBytes =
    System.GC.GetAllocatedBytesForCurrentThread() - tallWideVirtualGridBefore
check "tall wide grid table plan"
    (tallWideVirtualGridPlan.Rows.Length = tallWideVirtualGridRows
     && tallWideVirtualGridFragments = tallWideVirtualGridRows
     && tallWideVirtualGridSw.Elapsed.TotalMilliseconds < 350.0)
    true
check "tall wide grid table allocation"
    (tallWideVirtualGridBytes < 40000000L)
    true
let tallWideGridPaintRows = 4000
let tallWideGridPaintRow =
    { tableRow [ tableCell [ paragraph "tall-wide-grid-paint-cell" "x" ] None ] with
        GridAfter = tallWideVirtualGridColumns - 1 }
let tallWideGridPaintTable =
    { wideFastTable with
        Id = "tall-wide-grid-paint"
        Rows = [ for _ in 1 .. tallWideGridPaintRows -> tallWideGridPaintRow ] }
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let tallWideGridPaintBefore = System.GC.GetAllocatedBytesForCurrentThread()
let tallWideGridPaintSw = System.Diagnostics.Stopwatch.StartNew()
let _, tallWideGridPaintNextY =
    tableBorderCommands tallFastSettings measureText tallWideGridPaintTable 0.0
tallWideGridPaintSw.Stop()
let tallWideGridPaintBytes =
    System.GC.GetAllocatedBytesForCurrentThread() - tallWideGridPaintBefore
check "tall wide grid paint perf"
    (tallWideGridPaintNextY > 0.0
     && tallWideGridPaintSw.Elapsed.TotalMilliseconds < 350.0)
    true
check "tall wide grid paint allocation"
    (tallWideGridPaintBytes < 35000000L)
    true
let repeatedHeaderFastHeaderCount = 5000
let repeatedHeaderFastBodyCount = 5000
let repeatedHeaderFastTable =
    { wideFastTable with
        Id = "repeated-header-fast"
        Rows =
            [ for row in 0 .. repeatedHeaderFastHeaderCount - 1 ->
                { tableRow [ tableCell [] None ] with IsHeader = true }
              for row in 0 .. repeatedHeaderFastBodyCount - 1 ->
                tableRow [ tableCell [] None ] ] }
let repeatedHeaderFastSw = System.Diagnostics.Stopwatch.StartNew()
let repeatedHeaderFastPlan =
    buildTableLayoutPlanForDocumentAt tallFastSettings measureText (rangeMeasurerFromTextMeasurer measureText) Doc.empty repeatedHeaderFastTable 0.0
repeatedHeaderFastSw.Stop()
let repeatedHeaderFastOk =
    repeatedHeaderFastPlan.Rows.Length = repeatedHeaderFastHeaderCount + repeatedHeaderFastBodyCount
    && repeatedHeaderFastPlan.HeaderRowCount = repeatedHeaderFastHeaderCount
    && repeatedHeaderFastSw.Elapsed.TotalMilliseconds < 1500.0
if not repeatedHeaderFastOk then
    printfn
        "repeated header table plan details: headers=%d rows=%d plan=%.1fms"
        repeatedHeaderFastPlan.HeaderRowCount
        repeatedHeaderFastPlan.Rows.Length
        repeatedHeaderFastSw.Elapsed.TotalMilliseconds
check "repeated header table plan perf"
    repeatedHeaderFastOk
    true
let percentShadedCell index =
    { tableCell [ paragraph (sprintf "percent-shaded-%d" index) "x" ] None with
        Shading = Some { Pattern = Some "pct50"; ForegroundColor = Some "#000000"; BackgroundColor = Some "#ffffff" } }
let percentShadedTable =
    { wideFastTable with
        Id = "percent-shaded"
        Rows = [ tableRow [ for index in 0 .. 2 -> percentShadedCell index ] ] }
let percentShadedCommands, _ = tableBorderCommands settings measureText percentShadedTable 0.0
let percentShadedFillColors =
    percentShadedCommands
    |> List.choose (function
        | FillRect (color, _, _, _, _) -> Some color
        | _ -> None)
check "percent shaded table blended fills" percentShadedFillColors [ "#808080"; "#808080"; "#808080" ]

let deepNestedPaintDepth = 20000
let deepNestedPaintSettings = { settings with Width = 50000.0; PageHeight = 1000000.0 }
let tableWithSingleCell id cell =
    { wideFastTable with
        Id = id
        Rows = [ tableRow [ cell ] ] }
let mutable deepNestedPaintTable =
    tableWithSingleCell
        "deep-nested-paint-leaf-table"
        (tableCell [ paragraph "deep-nested-paint-leaf" "leaf" ] None)
for depth = deepNestedPaintDepth - 1 downto 0 do
    let nestedCell =
        { tableCell [] None with
            NestedTables = [ deepNestedPaintTable ]
            NestedTableAnchors = [ 0 ] }
    deepNestedPaintTable <-
        tableWithSingleCell (sprintf "deep-nested-paint-%05d" depth) nestedCell
let deepNestedPaintDoc =
    { Doc.empty with
        Blocks = [ Doc.Table deepNestedPaintTable ]
        Selection = None }
let deepNestedPaintSw = System.Diagnostics.Stopwatch.StartNew()
let deepNestedPaintLayouts, deepNestedPaintCommands, deepNestedPaintPages =
    collectDocumentCommandsWithLayouts
        deepNestedPaintSettings
        "Calibri"
        measureText
        (rangeMeasurerFromTextMeasurer measureText)
        deepNestedPaintDoc
        (Some [])
deepNestedPaintSw.Stop()
check "deep nested table paint stack"
    (deepNestedPaintSw.Elapsed.TotalMilliseconds < 3000.0
     && deepNestedPaintLayouts.Length = 0
     && deepNestedPaintCommands.Length = 0
     && deepNestedPaintPages.Length = 1)
    true

let headerTexts doc =
    let _, commands, _ =
        collectDocumentCommandsWithLayouts settings "Calibri" measureText (rangeMeasurerFromTextMeasurer measureText) doc (Some [ headerPageLayout ])
    commands
    |> List.choose (function
        | DrawText (text, _, _, baseline, _) -> Some (text, baseline)
        | _ -> None)

let staticHeaderDoc =
    { Doc.empty with
        Blocks = []
        Selection = None
        Headers = [ paragraph "static-header" "Static header" ] }
let staticHeaderTexts = headerTexts staticHeaderDoc
check "static header per page" (staticHeaderTexts |> List.map fst) [ "Static header"; "Static header" ]
check "static header y shift"
    (match staticHeaderTexts with | [ (_, first); (_, second) ] -> int (round (second - first)) | _ -> -1)
    (int (round (pageHeight settings + pageGap)))

// removed: tested deleted cache machinery

let referencedHeaderLayout =
    { headerPageLayout with Metadata = { Doc.defaultMetadata with HeaderReferences = [ "default", "target-header" ] } }
let referencedHeaderDoc =
    { Doc.empty with
        HeaderParts =
            [ { Kind = "default"; RelationshipId = Some "other-header"; Blocks = [ Doc.Paragraph (paragraph "wrong-header" "Wrong header") ] }
              { Kind = "default"; RelationshipId = Some "target-header"; Blocks = [ Doc.Paragraph (paragraph "right-header" "Right header") ] } ] }
let referencedHeaderTexts =
    let _, commands, _ =
        collectDocumentCommandsWithLayouts settings "Calibri" measureText (rangeMeasurerFromTextMeasurer measureText) referencedHeaderDoc (Some [ referencedHeaderLayout ])
    commands
    |> List.choose (function | DrawText (text, _, _, _, _) -> Some text | _ -> None)
check "referenced header part" referencedHeaderTexts [ "Right header"; "Right header" ]
let headerDepthDoc =
    { Doc.empty with
        HeaderParts =
            [ { Kind = "default"
                RelationshipId = None
                Blocks =
                    [ Doc.Paragraph (paragraph "depth-header-1" "one")
                      Doc.Paragraph (paragraph "depth-header-2" "two")
                      Doc.Paragraph (paragraph "depth-header-3" "three") ] } ] }
let headerDepthBodyTop = Rbt.Word.WordLayout.Page.bodyTopWithRunningHeader headerDepthDoc
let headerDepthChanged =
    Rbt.Word.WordLayout.Page.bodyTopWithRunningHeader
        { headerDepthDoc with HeaderParts = [ { Kind = "other"; RelationshipId = None; Blocks = [] } ] }
check "header part lookup depth updates" (headerDepthBodyTop > Doc.defaultPageSetup.MarginTop, headerDepthChanged) (true, Doc.defaultPageSetup.MarginTop)
let oversizedHeaderDepthParagraph = paragraph "oversized-header-depth-para" ""
let oversizedHeaderDepthCell =
    tableCell
        [ for _ in 0 .. 200000 -> oversizedHeaderDepthParagraph ]
        None
let oversizedHeaderDepthTable : Doc.TableBlock =
    { Id = "oversized-header-depth-table"
      Rows = [ tableRow [ oversizedHeaderDepthCell ] ]
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
let oversizedHeaderDepthDoc =
    { Doc.empty with
        HeaderParts =
            [ { Kind = "default"
                RelationshipId = None
                Blocks = [ Doc.Table oversizedHeaderDepthTable ] } ] }
let oversizedHeaderDepthTop =
    Rbt.Word.WordLayout.Page.bodyTopWithRunningHeader oversizedHeaderDepthDoc
check "header depth paragraph count cap"
    oversizedHeaderDepthTop
    (Doc.defaultPageSetup.HeaderDistance + 100001.0 * 30.0)
let allocatedBytesFor iterations f =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetAllocatedBytesForCurrentThread()
    for _ in 1 .. iterations do
        f()
    System.GC.GetAllocatedBytesForCurrentThread() - before
let modernHeaderSkipsLegacyFallbackDoc =
    { Doc.empty with
        Headers = [ for i in 0 .. 4999 -> paragraph (sprintf "unused-legacy-header-%04d" i) "" ]
        HeaderParts =
            [ { Kind = "default"
                RelationshipId = None
                Blocks = [ Doc.Paragraph (paragraph "modern-header-paint" "Modern header") ] } ] }
let mutable modernHeaderPaintIndex = 0
let modernHeaderSkipsLegacyFallbackBytes =
    allocatedBytesFor 20 (fun () ->
        modernHeaderPaintIndex <- modernHeaderPaintIndex + 1
        collectDocumentCommandsWithLayouts
            settings
            (sprintf "Calibri-%d" modernHeaderPaintIndex)
            measureText
            (rangeMeasurerFromTextMeasurer measureText)
            modernHeaderSkipsLegacyFallbackDoc
            (Some [ headerPageLayout ])
        |> ignore)
check "modern header skips legacy fallback allocation" (modernHeaderSkipsLegacyFallbackBytes < 4000000L) true
let modernBlocksSkipLegacyFallbackDoc =
    { Doc.empty with
        Headers = [ for i in 0 .. 4999 -> paragraph (sprintf "unused-default-header-%04d" i) "" ]
        Footers = [ for i in 0 .. 4999 -> paragraph (sprintf "unused-default-footer-%04d" i) "" ]
        HeaderParts =
            [ { Kind = "default"
                RelationshipId = None
                Blocks = [ Doc.Paragraph (paragraph "modern-default-header" "Modern header blocks") ] } ]
        FooterParts =
            [ { Kind = "default"
                RelationshipId = None
                Blocks = [ Doc.Paragraph (paragraph "modern-default-footer" "Modern footer blocks") ] } ] }
let defaultHeaderBlocks = Rbt.Word.WordLayout.Page.defaultHeaderBlocks
let defaultFooterBlocks = Rbt.Word.WordLayout.Page.defaultFooterBlocks
let blockText = function
    | Doc.Paragraph paragraph -> toString paragraph.Buffer
    | _ -> ""
check "modern header blocks prefer part" (defaultHeaderBlocks modernBlocksSkipLegacyFallbackDoc |> List.map blockText) [ "Modern header blocks" ]
check "modern footer blocks prefer part" (defaultFooterBlocks modernBlocksSkipLegacyFallbackDoc |> List.map blockText) [ "Modern footer blocks" ]
let modernBlocksSkipLegacyFallbackBytes =
    allocatedBytesFor 100 (fun () ->
        defaultHeaderBlocks modernBlocksSkipLegacyFallbackDoc |> ignore
        defaultFooterBlocks modernBlocksSkipLegacyFallbackDoc |> ignore)
check "modern header blocks skip legacy allocation" (modernBlocksSkipLegacyFallbackBytes < 1000000L) true
let selectionLookupMeasureRange = rangeMeasurerFromTextMeasurer measureText
let selectionLookupPerfLayouts =
    [ for i in 0 .. 19999 ->
        blockLayout (sprintf "selection-lookup-%d" i) (float i * 24.0) hitLine ]
let selectionLookupPerfDoc =
    { Doc.empty with
        Selection =
            Some {
                Anchor = { BlockId = "selection-lookup-19999"; Offset = 3 }
                Focus = { BlockId = "selection-lookup-19999"; Offset = 3 }
            } }
check "selection layout lookup command count"
    (selectionCommandsWithRange settings selectionLookupMeasureRange selectionLookupPerfDoc selectionLookupPerfLayouts |> List.length)
    1
let selectionLookupMultiBlockDoc =
    { Doc.empty with
        Blocks =
            [ for i in 0 .. 19999 ->
                Doc.Paragraph (paragraph (sprintf "selection-lookup-%d" i) "abcdef") ]
        Selection =
            Some {
                Anchor = { BlockId = "selection-lookup-19990"; Offset = 1 }
                Focus = { BlockId = "selection-lookup-19999"; Offset = 4 }
            } }
check "selection layout lookup multi-block count"
    (selectionCommandsWithRange settings selectionLookupMeasureRange selectionLookupMultiBlockDoc selectionLookupPerfLayouts |> List.length)
    11
// removed: tested deleted cache machinery
let bodyTopWithRunningHeader = Rbt.Word.WordLayout.Page.bodyTopWithRunningHeader
let legacyHeaderBodyTopDoc =
    { Doc.empty with
        Headers = [ for i in 0 .. 4999 -> paragraph (sprintf "legacy-header-%04d" i) "" ] }
check "legacy header body top depth" (bodyTopWithRunningHeader legacyHeaderBodyTopDoc > Doc.defaultPageSetup.MarginTop) true
let legacyHeaderBodyTopBytes =
    allocatedBytesFor 100 (fun () ->
        bodyTopWithRunningHeader legacyHeaderBodyTopDoc |> ignore)
check "legacy header body top allocation" (legacyHeaderBodyTopBytes < 1000000L) true
// removed: tested deleted cache machinery

let roundRectShapePath = Rbt.Word.Shapes.presetPath "roundRect" 320.0 180.0
let ellipseShapePath = Rbt.Word.Shapes.presetPath "ellipse" 320.0 180.0
let star32ShapePath = Rbt.Word.Shapes.presetPath "star32" 320.0 180.0
check "shape roundrect segments" roundRectShapePath.Length 10
check "shape ellipse segments" ellipseShapePath.Length 6
check "shape star32 segments" star32ShapePath.Length 65
check "shape star32 endpoints"
    ((match star32ShapePath with | Doc.PMove _ :: _ -> true | _ -> false),
     (match List.last star32ShapePath with | Doc.PClose -> true | _ -> false))
    (true, true)
// removed: tested deleted cache machinery

let pageFieldBuffer =
    ofString "0"
    |> applyStyle 0 1 (fun style -> { style with FieldCode = Some "PAGE" })
let pageFieldHeaderDoc =
    { staticHeaderDoc with
        Headers = [ { paragraph "page-field-header" "0" with Buffer = pageFieldBuffer } ] }
check "page field header per page" (headerTexts pageFieldHeaderDoc |> List.map fst) [ "1"; "2" ]

let mutable deepNestedFieldTable =
    tableWithSingleCell
        "deep-nested-field-leaf-table"
        (tableCell [ { paragraph "deep-nested-field-leaf" "0" with Buffer = pageFieldBuffer } ] None)
for depth = deepNestedPaintDepth - 1 downto 0 do
    let nestedCell =
        { tableCell [] None with
            NestedTables = [ deepNestedFieldTable ]
            NestedTableAnchors = [ 0 ] }
    deepNestedFieldTable <-
        tableWithSingleCell (sprintf "deep-nested-field-%05d" depth) nestedCell
let deepNestedFieldHeaderDoc =
    { staticHeaderDoc with
        Headers = []
        HeaderParts =
            [ { Kind = "default"
                RelationshipId = None
                Blocks = [ Doc.Table deepNestedFieldTable ] } ] }
let deepNestedFieldHeaderSw = System.Diagnostics.Stopwatch.StartNew()
let deepNestedFieldHeaderLayouts, _, deepNestedFieldHeaderPages =
    collectDocumentCommandsWithLayouts
        settings
        "Calibri"
        measureText
        (rangeMeasurerFromTextMeasurer measureText)
        deepNestedFieldHeaderDoc
        (Some [ headerPageLayout ])
deepNestedFieldHeaderSw.Stop()
check "deep nested header field stack"
    (deepNestedFieldHeaderSw.Elapsed.TotalMilliseconds < 3000.0
     && deepNestedFieldHeaderLayouts.Length = 1
     && deepNestedFieldHeaderPages.Length = 2)
    true

// ---------------------------------------------------------------------------
// Line spacing rules. The importer hands paragraph-level "exact"/"atLeast"
// spacing as (line height px / paragraph font size px) with the font size in
// ParagraphMarkFontSize, and style-level spacing as absolute px.
// ---------------------------------------------------------------------------
let lineRuleSettings = { settings with Width = 240.0; PageHeight = 100000.0 }
let lineRuleFontPx = 11.0 * 96.0 / 72.0
let lineRulePara id metadata =
    { paragraph id "alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu" with Metadata = metadata }
let lineRuleLayout (doc: Doc.Document) id =
    layoutDocument lineRuleSettings measureText doc |> List.find (fun layout -> layout.BlockId = id)
let lineRulePitch (layout: BlockLayout) =
    match layout.Lines with
    | first :: second :: _ -> System.Math.Round(second.Y - first.Y, 3)
    | _ -> -1.0
let lineRuleMetadata rule px =
    { Doc.defaultMetadata with
        SpacingBefore = Some 0.0
        SpacingAfter = Some 0.0
        LineSpacing = Some (px / lineRuleFontPx)
        LineSpacingRule = Some rule
        ParagraphMarkFontSize = Some lineRuleFontPx }
let lineRuleDoc blocks styles = { Doc.empty with Blocks = blocks; Styles = styles; Selection = None }
let exactLayout = lineRuleLayout (lineRuleDoc [ Doc.Paragraph (lineRulePara "exact" (lineRuleMetadata "exact" 20.0)) ] []) "exact"
check "exact spacing wraps" (exactLayout.Lines.Length >= 3) true
check "exact spacing line pitch" (lineRulePitch exactLayout) 20.0
check "exact spacing line height" (System.Math.Round(exactLayout.Lines.Head.Height, 3)) 20.0
let autoLayout =
    lineRuleLayout (lineRuleDoc [ Doc.Paragraph (lineRulePara "auto" { Doc.defaultMetadata with SpacingBefore = Some 0.0; SpacingAfter = Some 0.0 }) ] []) "auto"
let atLeastSmall = lineRuleLayout (lineRuleDoc [ Doc.Paragraph (lineRulePara "atleast-small" (lineRuleMetadata "atLeast" 5.0)) ] []) "atleast-small"
check "atLeast below natural keeps single spacing" (lineRulePitch atLeastSmall > 12.0 && lineRulePitch atLeastSmall < 24.0) true
let atLeastLarge = lineRuleLayout (lineRuleDoc [ Doc.Paragraph (lineRulePara "atleast-large" (lineRuleMetadata "atLeast" 40.0)) ] []) "atleast-large"
check "atLeast above natural uses the minimum" (lineRulePitch atLeastLarge) 40.0
let exactParagraphStyle =
    { namedStyle "ExactStyle" None None with StyleType = "paragraph"; LineSpacing = Some 20.0; LineSpacingRule = Some "exact" }
let styledExactMetadata = { Doc.defaultMetadata with Style = Some "ExactStyle"; SpacingBefore = Some 0.0; SpacingAfter = Some 0.0 }
let styleExactLayout =
    lineRuleLayout (lineRuleDoc [ Doc.Paragraph (lineRulePara "style-exact" styledExactMetadata) ] [ exactParagraphStyle ]) "style-exact"
check "style exact spacing is absolute px" (lineRulePitch styleExactLayout) 20.0
// w:line without w:lineRule on the paragraph is "auto" even under an exact style.
let overrideLayout =
    lineRuleLayout
        (lineRuleDoc [ Doc.Paragraph (lineRulePara "style-override" { styledExactMetadata with LineSpacing = Some 2.0 }) ] [ exactParagraphStyle ])
        "style-override"
check "paragraph multiple overrides exact style" (lineRulePitch overrideLayout > lineRulePitch autoLayout * 1.5) true

printfn ""
printfn "Word layout tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
