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
      LongTrailingSpaceFitSafety = 0.25; DefaultTabStop = 48.0 }

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

// First-line indents and tab stops, as Word sets them. Every line is as wide
// as the paragraph between its indents; only the first starts further in (or
// out, hanging). Tab stops - custom and default - are measured from the
// margin, not from the paragraph's indent; a hanging indent is a tab stop of
// the first line.
let indentLines (metadata: Doc.BlockMetadata) text =
    let para : Doc.ParagraphBlock = { Id = "indent"; Buffer = ofString text; Metadata = metadata }
    let layout, _ = layoutParagraphAtWithRange settings measureText (rangeMeasurerFromTextMeasurer measureText) 210.0 0.0 0.0 para Map.empty
    layout.Lines |> List.map (fun line -> line.Segments |> List.map (function TextSegment (t, _, x, _, _) -> x, t | _ -> -1.0, ""))
let lineText (segments: (float * string) list) = segments |> List.map snd |> String.concat ""
let lineStart (segments: (float * string) list) = fst segments.Head
// 11 words of 4 letters, 7 px a letter: five words and their spaces (35 px a word) fill 175 px.
let elevenWords = "aaaa bbbb cccc dddd eeee ffff gggg hhhh iiii jjjj kkkk"
let firstLineIndented = indentLines { Doc.defaultMetadata with FirstLineIndent = Some 28.0 } elevenWords
check "first-line indent: the first line starts at it" (lineStart firstLineIndented.[0]) 28.0
check "first-line indent: the lines after it are as wide as the paragraph (6 words, not 5)"
    (firstLineIndented |> List.map lineText) [ "aaaa bbbb cccc dddd eeee "; "ffff gggg hhhh iiii jjjj kkkk" ]
let hangingLines = indentLines { Doc.defaultMetadata with IndentLeft = Some 48.0; FirstLineIndent = Some -48.0 } elevenWords
check "hanging indent: the first line starts at the margin and has its width (6 words)"
    (hangingLines |> List.map (fun line -> lineStart line, lineText line)) [ 0.0, "aaaa bbbb cccc dddd eeee ffff "; 48.0, "gggg hhhh iiii jjjj "; 48.0, "kkkk" ]
let hangingTab = indentLines { Doc.defaultMetadata with IndentLeft = Some 48.0; FirstLineIndent = Some -48.0 } "1.\tterm"
check "hanging indent: a tab on the first line goes to the indent" (hangingTab.[0] |> List.last) (48.0, "term")
let customStop = indentLines { Doc.defaultMetadata with IndentLeft = Some 48.0; TabStops = [ { Position = 96.0; Alignment = "left"; Leader = None } ] } "a\tb"
check "a custom tab stop is measured from the margin, not the indent" (customStop.[0] |> List.last) (96.0, "b")
// Justified: every line but the last reaches the right indent, the first-line-indented and hanging ones too.
let lineReach (segments: (float * string) list) = segments |> List.map (fun (x, t) -> x + 7.0 * float (t.TrimEnd().Length)) |> List.max
let justifiedWords = elevenWords + " llll mmmm nnnn oooo pppp qqqq"
for label, metadata in [ "first-line", { Doc.defaultMetadata with FirstLineIndent = Some 28.0 }; "hanging", { Doc.defaultMetadata with IndentLeft = Some 48.0; FirstLineIndent = Some -48.0 } ] do
    let justified = indentLines { metadata with Alignment = Doc.Justify; ExplicitAlignment = true } justifiedWords
    check (sprintf "justified, %s indent: the lines but the last reach the right edge" label)
        (justified |> List.take (justified.Length - 1) |> List.map (fun line -> System.Math.Round(lineReach line, 1))) (List.replicate (justified.Length - 1) 210.0)
let defaultStop = indentLines { Doc.defaultMetadata with IndentLeft = Some 20.0 } "a\tb"
check "default tab stops are every half inch from the margin" (defaultStop.[0] |> List.last) (48.0, "b")

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
        Segments = [ ImageSegment ("img", 7.0, 0.0, 32.0, 24.0, 0, false, None, None, None, None, None, None, (false, false)) ]
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
      EnclosingControls = []
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
      PaddingLeft = None; Origin = None }

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
      BordersOverride = None; Origin = None }

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
      Indent = None
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
      Floating = false
      FloatingXml = None
      StyleId = None
      ContentControlId = None; ContentControlTag = None; ContentControlAlias = None; ContentControlXml = None
      ContentControlGroup = None; ContentControlAncestors = []
      Look = None }
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
let tallWideGridPaintSw = System.Diagnostics.Stopwatch.StartNew()
let _, tallWideGridPaintNextY =
    tableBorderCommands tallFastSettings measureText tallWideGridPaintTable 0.0
tallWideGridPaintSw.Stop()
check "tall wide grid paint perf"
    (tallWideGridPaintNextY > 0.0
     && tallWideGridPaintSw.Elapsed.TotalMilliseconds < 350.0)
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
// Three sections: the first two carry their header on the paragraph that ends
// them, the final one in the page setup. The final section's header used to be
// attached to the paragraph after the first break, i.e. inside section two.
let sectionPagePitch = pageHeight settings + pageGap
let sectionEndLayout id pageIndex (refs: (string * string) list) sectionBreak =
    { blockLayout id (float pageIndex * sectionPagePitch + 40.0) hitLine with
        Height = 20.0
        Metadata = { Doc.defaultMetadata with HeaderReferences = refs; SectionBreak = sectionBreak } }
let threeSectionDoc =
    { Doc.empty with
        PageSetup = { Doc.defaultPageSetup with HeaderReferences = [ "default", "h3" ] }
        HeaderParts =
            [ { Kind = "default"; RelationshipId = Some "h1"; Blocks = [ Doc.Paragraph (paragraph "sec-h1" "Header one") ] }
              { Kind = "default"; RelationshipId = Some "h2"; Blocks = [ Doc.Paragraph (paragraph "sec-h2" "Header two") ] }
              { Kind = "default"; RelationshipId = Some "h3"; Blocks = [ Doc.Paragraph (paragraph "sec-h3" "Header three") ] } ] }
let sectionHeaderTexts doc layouts =
    let _, commands, _ =
        collectDocumentCommandsWithLayouts settings "Calibri" measureText (rangeMeasurerFromTextMeasurer measureText) doc (Some layouts)
    commands |> List.choose (function | DrawText (text, _, _, _, _) when text.StartsWith "Header" -> Some text | _ -> None)
let threeSectionLayouts =
    [ sectionEndLayout "sec-end-1" 0 [ "default", "h1" ] (Some "nextPage")
      { blockLayout "sec-2-body" (1.0 * sectionPagePitch + 40.0) hitLine with Height = 20.0 }
      sectionEndLayout "sec-end-2" 2 [ "default", "h2" ] (Some "nextPage")
      { blockLayout "sec-3-body" (4.0 * sectionPagePitch + 40.0) hitLine with Height = 20.0 } ]
check "three sections: each page gets its section's header"
    (sectionHeaderTexts threeSectionDoc threeSectionLayouts)
    [ "Header one"; "Header two"; "Header two"; "Header three"; "Header three" ]
// A section that names no header keeps the previous section's ("link to previous").
let inheritingLayouts =
    [ sectionEndLayout "inh-end-1" 0 [ "default", "h1" ] (Some "nextPage")
      sectionEndLayout "inh-end-2" 1 [] (Some "nextPage")
      { blockLayout "inh-3-body" (2.0 * sectionPagePitch + 40.0) hitLine with Height = 20.0 } ]
check "section without a header reference inherits the previous one"
    (sectionHeaderTexts { threeSectionDoc with PageSetup = Doc.defaultPageSetup } inheritingLayouts)
    [ "Header one"; "Header one"; "Header one" ]

// An object kept from the file without a picture is painted as a labelled frame of its size.
let opaqueRunDoc kind =
    let buffer = Rbt.Word.TextBuffer.empty |> Rbt.Word.TextBuffer.insertImage 0 (Rbt.Word.TextBuffer.opaqueSourcePrefix + kind) (Some 200.0) (Some 100.0) None
    { Doc.empty with
        Blocks = [ Doc.Paragraph { Id = "opaque-para"; Buffer = buffer; Metadata = Doc.defaultMetadata } ]
        Selection = None }
let opaqueCommands kind =
    let _, commands, _ = collectDocumentCommands settings "Calibri" measureText (rangeMeasurerFromTextMeasurer measureText) (opaqueRunDoc kind)
    commands
check "opaque chart paints a labelled placeholder frame"
    (opaqueCommands "chart"
     |> List.choose (function
         | StrokeRect (_, _, _, _, width, height) -> Some (sprintf "frame %gx%g" width height)
         | DrawText (text, _, _, _, _) -> Some ("text " + text)
         | DrawImage _ -> Some "image"
         | _ -> None))
    [ "frame 200x100"; "text Chart" ]
check "opaque drawing without a picture paints nothing"
    (opaqueCommands "drawing" |> List.exists (function StrokeRect _ | DrawImage _ | DrawText _ -> true | _ -> false))
    false

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
      Indent = None
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
      Floating = false
      FloatingXml = None
      StyleId = None
      ContentControlId = None; ContentControlTag = None; ContentControlAlias = None; ContentControlXml = None
      ContentControlGroup = None; ContentControlAncestors = []
      Look = None }
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

// ---------------------------------------------------------------------------
// lastRenderedPageBreak hints must not be honoured on the 1,000,000px page
// used to measure a table row's natural height.
// ---------------------------------------------------------------------------
let hintText = "first line of the paragraph second line after the hint"
let hintBuffer =
    ofString hintText
    |> applyStyle 28 (hintText.Length - 28) (fun style -> { style with RenderedPageBreakBefore = true })
let hintCellContent =
    [ for i in 1 .. 8 -> paragraph (sprintf "hint-cell-lead-%d" i) "x" ]
    @ [ { paragraph "hint-cell" hintText with Buffer = hintBuffer } ]
let hintTable =
    { wideFastTable with
        Id = "hint-table"
        Rows = [ tableRow [ tableCell hintCellContent None ] ] }
let hintTableSettings = { settings with Width = 600.0; PageHeight = 1000.0 }
let hintRowMeasurement =
    measureTableRowForDocumentRowLocal hintTableSettings measureText (rangeMeasurerFromTextMeasurer measureText) Doc.empty hintTable 600.0 [ 600.0 ] 0.0 0.0 hintTable.Rows.Head
check "rendered break hint row-local height" (hintRowMeasurement.Height < 1000.0) true
let hintDocLayouts =
    layoutDocument hintTableSettings measureText
        { Doc.empty with
            Blocks = [ Doc.Table hintTable; Doc.Paragraph (paragraph "hint-after" "after") ]
            Selection = None }
check "rendered break hint keeps document short"
    ((uniformPageInfos hintTableSettings hintDocLayouts).Length <= 2)
    true

// ---------------------------------------------------------------------------
// Consecutive tabs each advance to the next tab stop.
// ---------------------------------------------------------------------------
let segmentXOf (needle: string) (line: MeasuredLine) =
    line.Segments
    |> List.pick (function
        | TextSegment (text, _, x, _, _) when text.StartsWith needle -> Some x
        | _ -> None)
let tabLineFor text = (layoutLinesStyled settings measureText (ofString text) 2000.0).Head
// "a" is 7px wide; default tab stops are every 48px.
check "one tab reaches the first stop" (segmentXOf "b" (tabLineFor "a\tb")) 48.0
check "two tabs reach the second stop" (segmentXOf "b" (tabLineFor "a\t\tb")) 96.0
check "three tabs reach the third stop" (segmentXOf "b" (tabLineFor "a\t\t\tb")) 144.0
check "space then tab keeps the tab" (segmentXOf "b" (tabLineFor "a \tb")) 48.0
check "tabs keep their offsets" (tabLineFor "a\t\tb").EndOffset 4

// ---------------------------------------------------------------------------
// A click puts the caret on the nearer side of the clicked glyph.
// ---------------------------------------------------------------------------
let clickLayouts =
    layoutDocument { settings with Width = 2000.0 } measureText
        { Doc.empty with Blocks = [ Doc.Paragraph { paragraph "click" "abcdefgh" with Metadata = { Doc.defaultMetadata with SpacingBefore = Some 0.0 } } ]; Selection = None }
let clickY = clickLayouts.Head.Lines.Head.Y + 2.0
let clickAt x = hitTestLayouts measureText 0.0 clickLayouts x clickY |> Option.map snd
// Glyphs are 7px wide: "c" spans 14..21.
check "click in left half of a glyph" (clickAt 15.0) (Some 2)
check "click in right half of a glyph" (clickAt 20.0) (Some 3)
check "click in left half of first glyph" (clickAt 1.0) (Some 0)
check "click in right half of last glyph" (clickAt 55.0) (Some 8)
check "click past the line end" (clickAt 500.0) (Some 8)

// ---------------------------------------------------------------------------
// The same floating picture used twice is two wrap obstacles.
// ---------------------------------------------------------------------------
let floatingParagraph id =
    let buffer =
        ofString (String.concat " " [ for _ in 1 .. 30 -> "wrap" ])
        |> insertImageWithMetadata 0 "memory://same" false (Some 100.0) (Some 60.0) None None None true (Some "square") (Some 0.0) (Some 0.0) None None None None None None None None None None None None None None None None None None
    { paragraph id "" with Buffer = buffer; Metadata = { Doc.defaultMetadata with SpacingBefore = Some 0.0 } }
let floatSettings = { settings with Width = 400.0; PageHeight = 100000.0 }
let floatLayouts =
    layoutDocument floatSettings measureText
        { Doc.empty with
            Blocks =
                [ Doc.Paragraph (floatingParagraph "float-a")
                  Doc.Paragraph (paragraph "float-gap-1" "gap")
                  Doc.Paragraph (paragraph "float-gap-2" "gap")
                  Doc.Paragraph (paragraph "float-gap-3" "gap")
                  Doc.Paragraph (paragraph "float-gap-4" "gap")
                  Doc.Paragraph (floatingParagraph "float-b") ]
            Selection = None }
let firstTextX id =
    floatLayouts
    |> List.find (fun layout -> layout.BlockId = id)
    |> fun layout ->
        layout.Lines.Head.Segments
        |> List.pick (function TextSegment (_, _, x, _, _) -> Some x | _ -> None)
check "first float pushes text aside" (firstTextX "float-a" >= 100.0) true
check "second use of the picture pushes text aside" (firstTextX "float-b" >= 100.0) true

// ---------------------------------------------------------------------------
// A blank line made by soft breaks is a line (the caret needs one).
// ---------------------------------------------------------------------------
let softBreakLines = layoutLinesStyled settings measureText (ofString "a\n\nb") 2000.0
check "soft break blank line count" softBreakLines.Length 3
check "soft break blank line offsets" (softBreakLines.[1].StartOffset, softBreakLines.[1].EndOffset) (2, 2)
check "soft break blank line has height" (softBreakLines.[1].Height > 0.0) true
check "soft break blank line is between" (softBreakLines.[0].Y < softBreakLines.[1].Y && softBreakLines.[1].Y < softBreakLines.[2].Y) true
let softBreakPara = { paragraph "soft" "a\n\nb" with Metadata = { Doc.defaultMetadata with SpacingBefore = Some 0.0 } }
let softBreakLayout, _ = layoutParagraphAtWithRange settings measureText (rangeMeasurerFromTextMeasurer measureText) 500.0 0.0 0.0 softBreakPara Map.empty
check "caret on blank line finds that line" (lineForOffset softBreakLayout 2).Y softBreakLayout.Lines.[1].Y
// A soft break at the very end of the paragraph opens a last, empty line for the caret.
let trailingBreakLines = layoutLinesStyled settings measureText (ofString "a\n") 2000.0
check "trailing soft break adds a line" trailingBreakLines.Length 2
check "trailing soft break line offsets" (trailingBreakLines.[1].StartOffset, trailingBreakLines.[1].EndOffset, trailingBreakLines.[1].Segments.Length) (2, 2, 0)
check "trailing soft break line is below" (trailingBreakLines.[1].Y > trailingBreakLines.[0].Y && trailingBreakLines.[1].Height > 0.0) true
let trailingBreakPara text = { paragraph "trailing-soft" text with Metadata = { Doc.defaultMetadata with SpacingBefore = Some 0.0 } }
let trailingBreakLayout text = layoutParagraphAtWithRange settings measureText (rangeMeasurerFromTextMeasurer measureText) 500.0 0.0 0.0 (trailingBreakPara text) Map.empty |> fst
check "caret after a trailing soft break is on the new line"
    ((lineForOffset (trailingBreakLayout "a\n") 2).Y, (lineForOffset (trailingBreakLayout "a\n") 1).Y)
    ((trailingBreakLayout "a\n").Lines.[1].Y, (trailingBreakLayout "a\n").Lines.[0].Y)
check "paragraph grows by the trailing line"
    ((trailingBreakLayout "a\n").Height > (trailingBreakLayout "a").Height, (trailingBreakLayout "a\n").Lines.Length, (trailingBreakLayout "a\nb").Lines.Length)
    (true, 2, 2)
check "two trailing soft breaks make two empty lines"
    ((layoutLinesStyled settings measureText (ofString "a\n\n") 2000.0) |> List.map (fun line -> line.StartOffset, line.EndOffset))
    [ 0, 1; 2, 2; 3, 3 ]
check "a lone soft break is two lines"
    ((layoutLinesStyled settings measureText (ofString "\n") 2000.0) |> List.map (fun line -> line.StartOffset, line.EndOffset))
    [ 0, 0; 1, 1 ]

// ---------------------------------------------------------------------------
// Numbered lists inside table cells count up.
// ---------------------------------------------------------------------------
let numberedCellPara id listId =
    { paragraph id "item" with Metadata = { Doc.defaultMetadata with ListStyle = Doc.Numbered; ListLevel = 1; ListId = Some listId } }
let listTable =
    { wideFastTable with
        Id = "list-table"
        Rows =
            [ tableRow [ tableCell [ numberedCellPara "lt-1" "L"; numberedCellPara "lt-2" "L" ] None
                         tableCell [ numberedCellPara "lt-3" "L" ] None ]
              tableRow [ tableCell [ numberedCellPara "lt-4" "L" ] None
                         tableCell [ numberedCellPara "lt-other" "M" ] None ] ] }
let listTableLayouts =
    layoutDocument { settings with Width = 600.0 } measureText { Doc.empty with Blocks = [ Doc.Table listTable ]; Selection = None }
let listPrefix id = (listTableLayouts |> List.find (fun layout -> layout.BlockId = id)).Prefix.Trim()
check "cell list item 1" (listPrefix "lt-1") "1."
check "cell list item 2" (listPrefix "lt-2") "2."
check "cell list continues in next cell" (listPrefix "lt-3") "3."
check "cell list continues in next row" (listPrefix "lt-4") "4."
check "second list in table starts at 1" (listPrefix "lt-other") "1."

// A list that starts in the body continues inside a table and on after it.
let flowSettingsForLists = { settings with Width = 1200.0; PageHeight = 2000.0 }
let bodyListPara id : Doc.ParagraphBlock =
    { paragraph id "item" with Metadata = { Doc.defaultMetadata with ListStyle = Doc.Numbered; ListLevel = 1; ListId = Some "L" } }
let continuedListTable =
    { wideFastTable with
        Id = "continued-list-table"
        Rows =
            [ tableRow [ tableCell [ numberedCellPara "ct-1" "L" ] None; tableCell [ numberedCellPara "ct-2" "L"; numberedCellPara "ct-other" "M" ] None ]
              tableRow [ tableCell [ numberedCellPara "ct-3" "L" ] None; tableCell [ paragraph "ct-plain" "plain" ] None ] ] }
let continuedListDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph (bodyListPara "cb-1"); Doc.Paragraph (bodyListPara "cb-2"); Doc.Paragraph (paragraph "cb-text" "between")
              Doc.Table continuedListTable
              Doc.Paragraph (bodyListPara "cb-after") ]
        Selection = None }
let continuedLayouts = layoutDocument { settings with Width = 600.0 } measureText continuedListDoc
let continuedPrefix id = (continuedLayouts |> List.find (fun layout -> layout.BlockId = id)).Prefix.Trim()
check "body list numbers" (continuedPrefix "cb-1", continuedPrefix "cb-2") ("1.", "2.")
check "list continues from the body into the table" (continuedPrefix "ct-1", continuedPrefix "ct-2", continuedPrefix "ct-3") ("3.", "4.", "5.")
check "another list in the table starts at 1" (continuedPrefix "ct-other") "1."
check "list continues after the table" (continuedPrefix "cb-after") "6."
let continuedPlan =
    buildTableLayoutPlanForDocumentAt { settings with Width = 600.0 } measureText (rangeMeasurerFromTextMeasurer measureText) continuedListDoc continuedListTable 0.0
let planPrefix id =
    continuedPlan.Rows
    |> List.collect (fun row -> row.Measurement.CellGroups |> List.collect (fun group -> group.Layouts))
    |> List.find (fun layout -> layout.BlockId = id)
    |> fun layout -> layout.Prefix.Trim()
check "table plan numbers cells like the document pass" (planPrefix "ct-1", planPrefix "ct-3") ("3.", "5.")
// ... and in a multi-column (column-flow) document too.
let continuedColumnLayouts =
    layoutDocument flowSettingsForLists measureText { continuedListDoc with PageSetup = { continuedListDoc.PageSetup with Columns = Some { Count = 2; Gap = 20.0; Separator = false; Widths = [] } } }
let continuedColumnPrefix id = (continuedColumnLayouts |> List.find (fun layout -> layout.BlockId = id)).Prefix.Trim()
check "list continues into a table in a column flow" (continuedColumnPrefix "ct-1", continuedColumnPrefix "ct-3", continuedColumnPrefix "cb-after") ("3.", "5.", "6.")

// ---------------------------------------------------------------------------
// A table in a header paints its text, not only its borders.
// ---------------------------------------------------------------------------
let headerTextTable =
    { wideFastTable with
        Id = "header-text-table"
        Rows = [ tableRow [ tableCell [ paragraph "header-cell-a" "Left cell" ] None; tableCell [ paragraph "header-cell-b" "Right cell" ] None ] ] }
let headerTableDoc =
    { staticHeaderDoc with
        Headers = []
        HeaderParts = [ { Kind = "default"; RelationshipId = None; Blocks = [ Doc.Table headerTextTable ] } ] }
let headerTableTexts = headerTexts headerTableDoc
let headerTableHas (needle: string) = headerTableTexts |> List.filter (fun (text, _) -> text.Contains needle)
check "header table text painted on each page" ((headerTableHas "Left").Length, (headerTableHas "Right").Length) (2, 2)
let headerTableBaselines = headerTableHas "Left" |> List.map snd
check "header table text sits at the page top"
    (headerTableBaselines |> List.forall (fun baseline ->
        let pageTop = floor (baseline / (pageHeight settings + pageGap)) * (pageHeight settings + pageGap)
        baseline - pageTop > 0.0 && baseline - pageTop < 120.0))
    true

// ---------------------------------------------------------------------------
// Columns and sections: a paragraph continued in the next column keeps the
// paragraph's own offsets, and footnotes are laid out.
// ---------------------------------------------------------------------------
let flowSettings = { settings with Width = 500.0; PageHeight = 400.0; PaddingTop = 40.0; MarginBottom = 40.0 }
let flowPageSetup =
    { Doc.defaultPageSetup with
        PageWidth = 500.0; PageHeight = 400.0; MarginLeft = 0.0; MarginRight = 0.0; MarginTop = 40.0; MarginBottom = 40.0 }
let twoColumns : Doc.SectionColumns = { Count = 2; Gap = 20.0; Separator = false; Widths = [] }
let longFlowText = String.concat " " [ for i in 1 .. 220 -> sprintf "w%03d" i ]
let columnFlowDoc =
    { Doc.empty with
        Blocks = [ Doc.Paragraph (paragraph "flow-lead" "lead"); Doc.Paragraph (paragraph "flow-long" longFlowText) ]
        PageSetup = { flowPageSetup with Columns = Some twoColumns }
        Selection = None }
let columnFlowLayouts = layoutDocument flowSettings measureText columnFlowDoc
let longFragments = columnFlowLayouts |> List.filter (fun layout -> layout.BlockId = "flow-long")
check "long paragraph continues in further columns" (longFragments.Length >= 2) true
let fragmentRanges =
    longFragments |> List.map (fun layout -> layout.Lines.Head.StartOffset, (layout.Lines |> List.last).EndOffset)
check "first fragment starts at 0" (fst fragmentRanges.Head) 0
check "fragments continue where the previous one ended"
    (fragmentRanges |> List.pairwise |> List.forall (fun ((_, previousEnd), (nextStart, _)) -> nextStart = previousEnd))
    true
check "fragments cover the whole paragraph" (snd (List.last fragmentRanges)) longFlowText.Length
let fragmentText =
    longFragments
    |> List.collect (fun layout -> layout.Lines)
    |> List.collect (fun line -> line.Segments)
    |> List.forall (function
        | TextSegment (text, _, _, _, startOffset) -> longFlowText.Substring(startOffset, text.Length) = text
        | _ -> true)
check "fragment segments map to the paragraph's own characters" fragmentText true

let footnotedPara id =
    let text = "body text with a note x"
    { paragraph id text with Buffer = ofString text |> setNote (text.Length - 1) 1 (Some "1") (Some "footnote") (Some "the footnote text") }
let isFootnoteLayout (layout: BlockLayout) = layout.BlockId.StartsWith footnoteBlockIdPrefix
let columnFootnoteLayouts =
    layoutDocument flowSettings measureText
        { columnFlowDoc with Blocks = [ Doc.Paragraph (footnotedPara "flow-note"); Doc.Paragraph (paragraph "flow-after" "after") ] }
check "two-column document lays its footnote out" (columnFootnoteLayouts |> List.filter isFootnoteLayout |> List.length) 1
let sectionBreakPara = { paragraph "sec-break" "end of first section" with Metadata = { Doc.defaultMetadata with SectionBreak = Some "nextPage" } }
let sectionFootnoteLayouts, sectionFootnotePages =
    layoutDocumentPages flowSettings measureText (rangeMeasurerFromTextMeasurer measureText)
        { Doc.empty with
            Blocks = [ Doc.Paragraph sectionBreakPara; Doc.Paragraph (footnotedPara "sec-note"); Doc.Paragraph (paragraph "sec-after" "after") ]
            PageSetup = flowPageSetup
            Selection = None }
let sectionFootnotes = sectionFootnoteLayouts |> List.filter isFootnoteLayout
check "sectioned document lays its footnote out" sectionFootnotes.Length 1
let sectionNoteBody = sectionFootnoteLayouts |> List.find (fun layout -> layout.BlockId = "sec-note")
let secondPage = sectionFootnotePages.[1]
check "section footnote sits at the bottom of the referencing page"
    (sectionFootnotes |> List.forall (fun layout ->
        layout.Y > sectionNoteBody.Y + sectionNoteBody.Height - 0.5
        && layout.Y >= secondPage.Top
        && layout.Y + layout.Height <= secondPage.Top + secondPage.Height - secondPage.MarginBottom + 0.5))
    true

// ---------------------------------------------------------------------------
// The column / section flow behaves like the single-column pass: endnotes are
// produced, adjoining paragraph spacing collapses, keep-with-next is honoured.
// ---------------------------------------------------------------------------
let endnotedPara id =
    let text = "body text with an endnote x"
    { paragraph id text with Buffer = ofString text |> setNote (text.Length - 1) 1 (Some "1") (Some "endnote") (Some "the endnote text") }
let isEndnoteLayout (layout: BlockLayout) = layout.BlockId.StartsWith endnoteBlockIdPrefix
let bodyBottomOf (layouts: BlockLayout list) =
    layouts |> List.filter (fun layout -> not (isEndnoteLayout layout) && not (isFootnoteLayout layout)) |> List.map (fun layout -> layout.Y + layout.Height) |> List.max
let columnEndnoteLayouts =
    layoutDocument flowSettings measureText
        { columnFlowDoc with Blocks = [ Doc.Paragraph (endnotedPara "flow-endnote"); Doc.Paragraph (paragraph "flow-after-endnote" "after") ] }
let columnEndnotes = columnEndnoteLayouts |> List.filter isEndnoteLayout
check "two-column document lays its endnote out" columnEndnotes.Length 1
check "column endnote sits below the body" (columnEndnotes |> List.forall (fun layout -> layout.Y >= bodyBottomOf columnEndnoteLayouts)) true
check "column endnote holds the note text"
    (columnEndnotes |> List.collect (fun layout -> layout.Lines) |> List.collect (fun line -> line.Segments) |> List.map textOfSegment |> String.concat "" |> fun text -> text.Contains "the endnote text")
    true
let sectionEndnoteLayouts, sectionEndnotePages =
    layoutDocumentPages flowSettings measureText (rangeMeasurerFromTextMeasurer measureText)
        { Doc.empty with
            Blocks = [ Doc.Paragraph sectionBreakPara; Doc.Paragraph (endnotedPara "sec-endnote"); Doc.Paragraph (paragraph "sec-after-endnote" "after") ]
            PageSetup = flowPageSetup
            Selection = None }
let sectionEndnotes = sectionEndnoteLayouts |> List.filter isEndnoteLayout
check "sectioned document lays its endnote out" sectionEndnotes.Length 1
check "section endnote sits below the body, on a page"
    (sectionEndnotes |> List.forall (fun layout ->
        layout.Y >= bodyBottomOf sectionEndnoteLayouts
        && sectionEndnotePages |> List.exists (fun page -> layout.Y >= page.Top + page.MarginTop - 0.5 && layout.Y + layout.Height <= page.Top + page.Height - page.MarginBottom + 0.5)))
    true
// Many endnotes spill onto further pages, which the page list covers.
let manyEndnotesPara index =
    let text = sprintf "note %d x" index
    { paragraph (sprintf "many-endnote-%d" index) text with Buffer = ofString text |> setNote (text.Length - 1) 1 (Some (string index)) (Some "endnote") (Some (String.concat " " [ for _ in 1 .. 40 -> "endnote" ])) }
let manyEndnoteLayouts, manyEndnotePages =
    layoutDocumentPages flowSettings measureText (rangeMeasurerFromTextMeasurer measureText)
        { Doc.empty with
            Blocks = Doc.Paragraph sectionBreakPara :: [ for index in 1 .. 12 -> Doc.Paragraph (manyEndnotesPara index) ]
            PageSetup = flowPageSetup
            Selection = None }
let lastManyPage = List.last manyEndnotePages
check "section endnotes all laid out" (manyEndnoteLayouts |> List.filter isEndnoteLayout |> List.length) 12
check "pages cover endnotes that spill over"
    (manyEndnoteLayouts |> List.filter isEndnoteLayout |> List.forall (fun layout -> layout.Y + layout.Height <= lastManyPage.Top + lastManyPage.Height))
    true
// Spacing after / before collapses in a column exactly as in the single-column pass.
let spacedPara id after before =
    { paragraph id "spaced" with Metadata = { Doc.defaultMetadata with SpacingAfter = Some after; SpacingBefore = Some before } }
let spacingBlocks = [ Doc.Paragraph (spacedPara "sp-1" 30.0 0.0); Doc.Paragraph (spacedPara "sp-2" 0.0 12.0); Doc.Paragraph (spacedPara "sp-3" 0.0 40.0) ]
let spacingGaps (layouts: BlockLayout list) =
    let y id = (layouts |> List.find (fun layout -> layout.BlockId = id)).Lines.Head.Y
    System.Math.Round(y "sp-2" - y "sp-1", 3), System.Math.Round(y "sp-3" - y "sp-2", 3)
let singleColumnSpacing = layoutDocument flowSettings measureText { columnFlowDoc with Blocks = spacingBlocks; PageSetup = flowPageSetup }
let twoColumnSpacing = layoutDocument flowSettings measureText { columnFlowDoc with Blocks = spacingBlocks }
check "column flow collapses paragraph spacing like the single-column pass" (spacingGaps twoColumnSpacing) (spacingGaps singleColumnSpacing)
let uncollapsedGap =
    let line = (singleColumnSpacing |> List.find (fun layout -> layout.BlockId = "sp-1")).Lines.Head
    System.Math.Round(line.Height + (30.0 + 12.0) * 96.0 / 72.0, 3)
check "the collapsed gap is smaller than the sum of the spacings" (fst (spacingGaps twoColumnSpacing) < uncollapsedGap) true
// Keep-with-next: a heading at the foot of a column moves with the text it introduces.
let keepNextStyle : Doc.NamedStyle = { namedStyle "KeepHead" None None with StyleType = "paragraph"; KeepNext = Some true }
let fillerPara index = { paragraph (sprintf "kn-fill-%d" index) "filler" with Metadata = { Doc.defaultMetadata with SpacingAfter = Some 0.0 } }
let keepDoc fillers keep =
    { columnFlowDoc with
        Styles = [ keepNextStyle ]
        Blocks =
            [ for index in 1 .. fillers -> Doc.Paragraph (fillerPara index) ]
            @ [ Doc.Paragraph { paragraph "kn-head" "heading" with Metadata = { Doc.defaultMetadata with Style = (if keep then Some "KeepHead" else None); SpacingAfter = Some 0.0 } }
                Doc.Paragraph { paragraph "kn-body" "body" with Metadata = { Doc.defaultMetadata with SpacingAfter = Some 0.0 } } ] }
let columnOfBlock (layouts: BlockLayout list) id =
    let layout = layouts |> List.find (fun layout -> layout.BlockId = id)
    match layout.Lines.Head.Segments with
    | TextSegment (_, _, x, _, _) :: _ -> if x >= 250.0 then 1 else 0
    | _ -> -1
let keepColumns fillers keep =
    let layouts = layoutDocument flowSettings measureText (keepDoc fillers keep)
    columnOfBlock layouts "kn-head", columnOfBlock layouts "kn-body"
// The filler count that leaves the heading alone at the foot of the first column.
let strandingFillers = [ 1 .. 40 ] |> List.tryFind (fun fillers -> keepColumns fillers false = (0, 1))
check "a heading can be stranded at the foot of a column" strandingFillers.IsSome true
check "keep-with-next takes the heading to the next column"
    (strandingFillers |> Option.map (fun fillers -> keepColumns fillers true))
    (Some (1, 1))

// keepNext / keepLines set directly on the paragraph work like the style's.
let directKeepDoc fillers keep =
    { columnFlowDoc with
        Blocks =
            [ for index in 1 .. fillers -> Doc.Paragraph (fillerPara index) ]
            @ [ Doc.Paragraph { paragraph "kn-head" "heading" with Metadata = { Doc.defaultMetadata with KeepNext = (if keep then Some true else None); SpacingAfter = Some 0.0 } }
                Doc.Paragraph { paragraph "kn-body" "body" with Metadata = { Doc.defaultMetadata with SpacingAfter = Some 0.0 } } ] }
let directKeepColumns fillers keep =
    let layouts = layoutDocument flowSettings measureText (directKeepDoc fillers keep)
    columnOfBlock layouts "kn-head", columnOfBlock layouts "kn-body"
check "direct keep-with-next moves the heading to the next column"
    (strandingFillers |> Option.map (fun fillers -> directKeepColumns fillers false, directKeepColumns fillers true))
    (Some ((0, 1), (1, 1)))
let pageOfBlock (layouts: BlockLayout list) id =
    let layout = layouts |> List.find (fun layout -> layout.BlockId = id)
    int (floor (layout.Lines.Head.Y / (pageHeight flowSettings + pageGap)))
let directKeepPages fillers keep =
    let layouts = layoutDocument flowSettings measureText { directKeepDoc fillers keep with PageSetup = flowPageSetup }
    pageOfBlock layouts "kn-head", pageOfBlock layouts "kn-body"
let strandingPageFillers = [ 1 .. 40 ] |> List.tryFind (fun fillers -> directKeepPages fillers false = (0, 1))
check "a heading can be stranded at the foot of a page" strandingPageFillers.IsSome true
check "direct keep-with-next takes the heading to the next page"
    (strandingPageFillers |> Option.map (fun fillers -> directKeepPages fillers true))
    (Some (1, 1))
check "a paragraph can switch its style's keep-with-next off"
    (Styles.keepNext { Styles.normal with KeepNext = Some true } { Doc.defaultMetadata with KeepNext = Some false },
     Styles.keepNext { Styles.normal with KeepNext = Some true } Doc.defaultMetadata,
     Styles.keepLines Styles.normal { Doc.defaultMetadata with KeepLines = Some true })
    (false, true, true)
// keepLines on the paragraph: it moves to the next page whole instead of splitting.
let keepLinesText = String.concat " " [ for index in 1 .. 40 -> sprintf "word%02d" index ]
let keepLinesDoc fillers keep =
    { columnFlowDoc with
        PageSetup = flowPageSetup
        Blocks =
            [ for index in 1 .. fillers -> Doc.Paragraph (fillerPara index) ]
            @ [ Doc.Paragraph { paragraph "kl-para" keepLinesText with Metadata = { Doc.defaultMetadata with KeepLines = (if keep then Some true else None); SpacingAfter = Some 0.0; WidowControl = Some false } } ] }
let keepLinesPages fillers keep =
    let layout = layoutDocument { flowSettings with Width = 300.0 } measureText (keepLinesDoc fillers keep) |> List.find (fun layout -> layout.BlockId = "kl-para")
    layout.Lines |> List.map (fun line -> int (floor (line.Y / (pageHeight flowSettings + pageGap)))) |> List.distinct
let splittingFillers = [ 1 .. 40 ] |> List.tryFind (fun fillers -> keepLinesPages fillers false = [ 0; 1 ])
check "a paragraph can split across pages" splittingFillers.IsSome true
check "direct keep-lines keeps the paragraph on one page"
    (splittingFillers |> Option.map (fun fillers -> keepLinesPages fillers true))
    (Some [ 1 ])

// ---------------------------------------------------------------------------
// Sections with their own page size: painting from an existing layout pass
// keeps the per-section page geometry instead of assuming uniform pages.
// ---------------------------------------------------------------------------
let landscapeSetup : Doc.SectionPageSetup =
    { PageWidth = 700.0; PageHeight = 300.0; MarginLeft = 0.0; MarginRight = 0.0; MarginTop = 40.0; MarginBottom = 40.0
      HeaderDistance = 10.0; FooterDistance = 10.0; DifferentFirstPage = false; EvenAndOddHeaders = false }
let mixedSectionDoc =
    { Doc.empty with
        Blocks =
            [ Doc.Paragraph { paragraph "mixed-1" "wide section" with Metadata = { Doc.defaultMetadata with SectionBreak = Some "nextPage"; SectionPageSetup = Some landscapeSetup } }
              Doc.Paragraph (paragraph "mixed-2" "body section") ]
        PageSetup = flowPageSetup
        Selection = None }
let mixedLayouts, mixedPages = layoutDocumentPages flowSettings measureText (rangeMeasurerFromTextMeasurer measureText) mixedSectionDoc
let pageSizes (pages: PageInfo list) = pages |> List.map (fun page -> page.Top, page.Width, page.Height)
check "sections lay out to pages of their own size" (pageSizes mixedPages) [ 0.0, 700.0, 300.0; 300.0 + pageGap, 500.0, 400.0 ]
let _, _, cachedOnlyPages =
    collectDocumentCommandsWithLayouts flowSettings "Calibri" measureText (rangeMeasurerFromTextMeasurer measureText) mixedSectionDoc (Some mixedLayouts)
check "painting from cached layouts keeps section pages" (pageSizes cachedOnlyPages) (pageSizes mixedPages)
let _, _, cachedPagesPages =
    collectDocumentCommandsWithLayoutsAndPages flowSettings "Calibri" measureText (rangeMeasurerFromTextMeasurer measureText) mixedSectionDoc mixedLayouts mixedPages
check "painting with the pass's pages keeps them" (pageSizes cachedPagesPages) (pageSizes mixedPages)
check "documentHasSections" (documentHasSections mixedSectionDoc, documentHasSections columnFlowDoc) (true, false)
check "page index follows non-uniform pages"
    ([ 10.0; 299.0; 300.0 + pageGap + 1.0; 5000.0 ] |> List.map (pageIndexForYInPages mixedPages))
    [ 0; 0; 1; 1 ]
check "pages extend over a note area below the last page"
    (pageInfosCoveringBottom (300.0 + pageGap + 400.0 + pageGap + 50.0) mixedPages |> pageSizes |> List.last)
    (300.0 + pageGap + 400.0 + pageGap, 500.0, 400.0)
check "pages are left alone when they cover the content" (pageInfosCoveringBottom 500.0 mixedPages |> List.length) 2

// ---------------------------------------------------------------------------
// A tall paragraph in a vertically merged cell spreads over the rows it spans
// instead of inflating the restart row; only what does not fit grows the last
// row. The plan, the block flow and the border painter agree on the heights.
// ---------------------------------------------------------------------------
let mergeFlowSettings = { settings with Width = 600.0; PageHeight = 100000.0 }
let mergeFlowTable id (longText: string) merged : Doc.TableBlock =
    let continuation index = if merged then tableCell [ paragraph (sprintf "%s-a%d" id index) "" ] (Some Doc.ContinueMerge) else tableCell [ paragraph (sprintf "%s-a%d" id index) "" ] None
    { wideFastTable with
        Id = id
        Rows =
            [ tableRow [ tableCell [ paragraph (id + "-a1") longText ] (if merged then Some Doc.RestartMerge else None); tableCell [ paragraph (id + "-b1") "b1" ] None ]
              tableRow [ continuation 2; tableCell [ paragraph (id + "-b2") "b2" ] None ]
              tableRow [ continuation 3; tableCell [ paragraph (id + "-b3") "b3" ] None ] ] }
let mergeFlowHeights table =
    (buildTableLayoutPlanForDocumentAt mergeFlowSettings measureText (rangeMeasurerFromTextMeasurer measureText) Doc.empty table 0.0).Rows
    |> List.map (fun row -> System.Math.Round(row.Height, 3))
let tallText = String.concat " " [ for index in 1 .. 60 -> sprintf "word%d" index ]
let shortText = String.concat " " [ for index in 1 .. 9 -> sprintf "word%d" index ]
let plainHeights = mergeFlowHeights (mergeFlowTable "mf-plain" "a1" false)
let unmergedTallHeights = mergeFlowHeights (mergeFlowTable "mf-unmerged" tallText false)
let mergedTallHeights = mergeFlowHeights (mergeFlowTable "mf-tall" tallText true)
let mergedShortHeights = mergeFlowHeights (mergeFlowTable "mf-short" shortText true)
check "unmerged tall cell inflates its row" (unmergedTallHeights.[0] > 3.0 * plainHeights.[0]) true
check "merged tall cell leaves restart row alone" (mergedTallHeights.[0], mergedTallHeights.[1]) (plainHeights.[0], plainHeights.[1])
check "merged tall cell grows the last spanned row" (mergedTallHeights.[2] > plainHeights.[2]) true
check "merged rows together hold the content"
    (abs (List.sum mergedTallHeights - unmergedTallHeights.[0]) < 0.01)
    true
check "merged content that fits changes no row" mergedShortHeights plainHeights
let mergeFlowDoc table = { Doc.empty with Blocks = [ Doc.Table table ]; Selection = None }
let mergeFlowLayouts = layoutDocument mergeFlowSettings measureText (mergeFlowDoc (mergeFlowTable "mf-tall" tallText true))
let mergeFlowY id = (mergeFlowLayouts |> List.find (fun layout -> layout.BlockId = id)).Y
check "block flow places rows by the merged heights"
    (System.Math.Round(mergeFlowY "mf-tall-b2" - mergeFlowY "mf-tall-b1", 3), System.Math.Round(mergeFlowY "mf-tall-b3" - mergeFlowY "mf-tall-b2", 3))
    (mergedTallHeights.[0], mergedTallHeights.[1])
let unmergedFlowLayouts = layoutDocument mergeFlowSettings measureText (mergeFlowDoc (mergeFlowTable "mf-unmerged" tallText false))
let linesOf (layouts: BlockLayout list) id = (layouts |> List.find (fun layout -> layout.BlockId = id)).Lines |> List.map (fun line -> System.Math.Round(line.Y, 3))
check "merged paragraph keeps all its lines, flowing over the rows below"
    (linesOf mergeFlowLayouts "mf-tall-a1")
    (linesOf unmergedFlowLayouts "mf-unmerged-a1")
check "merged paragraph has many lines" ((linesOf mergeFlowLayouts "mf-tall-a1").Length > 6) true
let borderAdvance table = snd (tableBorderCommands mergeFlowSettings measureText table 0.0)
// (reference rows hold an empty paragraph: the painter's single-line shortcut
// does not take those, so both tables are measured the same way)
let plainEmptyTable = mergeFlowTable "mf-plain-e" "" false
check "border painter agrees on merged heights"
    (abs ((borderAdvance (mergeFlowTable "mf-tall" tallText true) - List.sum mergedTallHeights)
          - (borderAdvance plainEmptyTable - List.sum (mergeFlowHeights plainEmptyTable))) < 0.01)
    true
// Content spread over the merge by the importer keeps sizing its own rows.
let distributedRows =
    Rbt.Word.WordLayout.Table.distributeVerticalMergeContent
        [ tableRow [ tableCell [ paragraph "mfd-1" tallText; paragraph "mfd-2" "two" ] (Some Doc.RestartMerge); tableCell [ paragraph "mfd-b1" "b1" ] None ]
          tableRow [ tableCell [ paragraph "mfd-e" "" ] (Some Doc.ContinueMerge); tableCell [ paragraph "mfd-b2" "b2" ] None ] ]
check "distributed merge content still sizes its row"
    ((mergeFlowHeights { wideFastTable with Id = "mf-dist"; Rows = distributedRows }).[0] > 3.0 * plainHeights.[0])
    true
// A nested table gets the same treatment.
let nestedMergeHost inner =
    { wideFastTable with Id = "mf-host"; Rows = [ tableRow [ { tableCell [ paragraph "mf-host-p" "" ] None with NestedTables = [ inner ] } ] ] }
let nestedHeight inner = (mergeFlowHeights (nestedMergeHost inner)).[0]
check "nested merged table spreads content too"
    (abs (nestedHeight (mergeFlowTable "mf-n-tall" tallText true) - nestedHeight (mergeFlowTable "mf-n-unmerged" tallText false)
          + 2.0 * plainHeights.[1]) < 0.01)
    true

// ---------------------------------------------------------------------------
// A w:lastRenderedPageBreak at a paragraph's start (Word started page N there
// when it last laid the document out) starts a page unless this layout has
// started that page already: lines that went on to a page of their own before
// the mark (the text came out taller than Word's) are on Word's page N, and a
// page break at the mark left them alone on a page (docx4j-getting-started:
// five pages of one line each).
// ---------------------------------------------------------------------------
let hintLine index = { paragraph (sprintf "hint-line-%d" index) (sprintf "line %d" index) with Metadata = { Doc.defaultMetadata with SpacingAfter = Some 0.0 } }
// (Word's page `ordinal`, from 0; the mark is the first run's too, as the reader gives it.)
let hintedPara id ordinal derivedBreak =
    { paragraph id "hinted paragraph" with
        Buffer = ofString "hinted paragraph" |> applyStyle 0 16 (fun style -> { style with RenderedPageBreakBefore = true })
        Metadata =
            { Doc.defaultMetadata with
                SpacingAfter = Some 0.0
                RenderedPageBreakBefore = Some ordinal
                // (The reader makes a Heading 1's mark a page break of the paragraph's: see DocxReader.)
                PageBreakBefore = derivedBreak
                Explicit = if derivedBreak then Doc.ParagraphExplicit.DerivedPageBreakBefore else 0 } }
let hintPages blocks =
    let layouts, pages = layoutDocumentPages flowSettings measureText (rangeMeasurerFromTextMeasurer measureText) { Doc.empty with Blocks = blocks; PageSetup = flowPageSetup; Selection = None }
    fun id -> layouts |> List.find (fun layout -> layout.BlockId = id) |> fun layout -> pageIndexForYInPages pages layout.Lines.Head.Y
// As many lines as fit on a page and two more, which go on to the next.
let linesOnFirstPage =
    let pageOf = hintPages [ for index in 1 .. 100 -> Doc.Paragraph (hintLine index) ]
    [ 1 .. 100 ] |> List.filter (fun index -> pageOf (sprintf "hint-line-%d" index) = 0) |> List.length
let overflowing = [ for index in 1 .. linesOnFirstPage + 2 -> Doc.Paragraph (hintLine index) ]
let lastOverflowing = sprintf "hint-line-%d" (linesOnFirstPage + 2)
check "the text before Word's page 1 runs on to page 1" ((hintPages overflowing) lastOverflowing) 1
check "a mark of Word's page 1 on our page 1 starts no page"
    ((hintPages (overflowing @ [ Doc.Paragraph (hintedPara "hinted" 1 false) ])) "hinted") 1
check "nor does a Heading 1's page break made of the mark"
    ((hintPages (overflowing @ [ Doc.Paragraph (hintedPara "hinted" 1 true) ])) "hinted") 1
check "a mark of Word's page 1 on our page 0 starts page 1"
    ((hintPages [ Doc.Paragraph (hintLine 1); Doc.Paragraph (hintedPara "hinted" 1 false) ]) "hinted") 1
// From the mark on the pages are counted as Word's: the next mark, with nothing run on to a
// page since, starts its page.
check "the next mark starts the page after"
    ((hintPages (overflowing @ [ Doc.Paragraph (hintedPara "hinted" 1 false); Doc.Paragraph (hintLine 0); Doc.Paragraph (hintedPara "hinted-2" 2 false) ])) "hinted-2") 2
// In a section after the first, whose pages are numbered from the section's first: before its
// first mark, a page the section started by itself is the mark's.
let firstSection = { hintLine 0 with Metadata = { (hintLine 0).Metadata with SectionBreak = Some "nextPage" } }
let sectionPages blocks = hintPages (Doc.Paragraph firstSection :: blocks)
check "in a later section the text runs on to its second page" ((sectionPages overflowing) lastOverflowing) 2
check "a mark in a later section on a page it started starts no page"
    ((sectionPages (overflowing @ [ Doc.Paragraph (hintedPara "hinted" 2 false) ])) "hinted") 2
check "a mark in a later section on its first page starts a page"
    ((sectionPages [ Doc.Paragraph (hintLine 1); Doc.Paragraph (hintedPara "hinted" 2 false) ]) "hinted") 2

// ---------------------------------------------------------------------------
// Each line is as tall as its own text: a paragraph whose first word is set
// large has one tall line, and the lines after it are those of the small text
// (phpword-sample-11: an 18 pt run made all four lines 23.78 pt; Word and
// LibreOffice make the others 12.4 pt). An exact line height stays the same on
// every line; an at-least one is the least of each.
// ---------------------------------------------------------------------------
let mixedSizeText = "BIG " + String.concat " " [ for i in 1 .. 60 -> sprintf "w%02d" i ]
let mixedSizePara (metadata: Doc.BlockMetadata) =
    { paragraph "mixed-size" mixedSizeText with
        Buffer = ofString mixedSizeText |> applyStyle 0 3 (fun style -> { style with FontSize = 32.0 })
        Metadata = { metadata with SpacingAfter = Some 0.0; SpacingBefore = Some 0.0 } }
let mixedLineHeights metadata =
    let layout, _ = layoutParagraphAtWithRange settings measureText (rangeMeasurerFromTextMeasurer measureText) 120.0 0.0 0.0 (mixedSizePara metadata) Map.empty
    layout.Lines |> List.map (fun line -> System.Math.Round(line.Height, 2))
let smallLineHeight =
    let layout, _ = layoutParagraphAtWithRange settings measureText (rangeMeasurerFromTextMeasurer measureText) 120.0 0.0 0.0 { paragraph "small" "w01 w02" with Metadata = { Doc.defaultMetadata with SpacingAfter = Some 0.0; SpacingBefore = Some 0.0 } } Map.empty
    System.Math.Round(layout.Lines.Head.Height, 2)
let autoHeights = mixedLineHeights Doc.defaultMetadata
check "the large word's line is taller than the small text's" (autoHeights.Head > smallLineHeight) true
check "the lines after it are the small text's" (autoHeights.Tail |> List.forall (fun height -> height = smallLineHeight)) true
let exactHeights = mixedLineHeights { Doc.defaultMetadata with LineSpacing = Some (30.0 / 32.0); LineSpacingRule = Some "exact"; ParagraphMarkFontSize = Some 32.0 }
check "an exact line height is every line's" (exactHeights |> List.distinct) [ 30.0 ]
let atLeastHeights = mixedLineHeights { Doc.defaultMetadata with LineSpacing = Some (20.0 / 32.0); LineSpacingRule = Some "atLeast"; ParagraphMarkFontSize = Some 32.0 }
check "an at-least line height is each line's least" (atLeastHeights.Tail |> List.distinct) [ max 20.0 smallLineHeight ]
check "... and the large word's line is its own" (atLeastHeights.Head) autoHeights.Head

// ---------------------------------------------------------------------------
// A line box is of its font's metrics: Word sets a line at the font's GDI
// height (usWinAscent + usWinDescent, and the hhea line gap beyond them).
// Arial's is 1.1499 of the size and Calibri's 1.2207, so Arial lines are 6%
// lower than Calibri's at the same spacing (they were all Calibri's). Text in
// two fonts on one line: the line is as tall as the taller font's.
// ---------------------------------------------------------------------------
let fontLineHeights family (metadata: Doc.BlockMetadata) text =
    let para =
        { paragraph "font-line" text with
            Buffer = ofString text |> applyStyle 0 (String.length text) (fun style -> { style with FontFamily = Some family; FontSize = 20.0 })
            Metadata = { metadata with SpacingAfter = Some 0.0; SpacingBefore = Some 0.0 } }
    let layout, _ = layoutParagraphAtWithRange settings measureText (rangeMeasurerFromTextMeasurer measureText) 120.0 0.0 0.0 para Map.empty
    layout.Lines |> List.map (fun line -> line.Height)
let multiple = { Doc.defaultMetadata with LineSpacing = Some (259.0 / 240.0); LineSpacingRule = Some "auto" }
let calibriLine = (fontLineHeights "Calibri" multiple "w01").Head
let arialLine = (fontLineHeights "Arial" multiple "w01").Head
check "Arial lines are Arial's height, not Calibri's" (System.Math.Round(arialLine / calibriLine, 3)) (System.Math.Round(1.1499 / 1.2207, 3))
let singleCalibri = (fontLineHeights "Calibri" Doc.defaultMetadata "w01").Head
let singleArial = (fontLineHeights "Arial" Doc.defaultMetadata "w01").Head
check "... at single spacing too" (System.Math.Round(singleArial / singleCalibri, 3)) (System.Math.Round(1.1499 / 1.2207, 3))
let mixedFontText = "aaa " + String.concat " " [ for i in 1 .. 30 -> sprintf "w%02d" i ]
let mixedFontLines =
    let para =
        { paragraph "font-mix" mixedFontText with
            Buffer =
                ofString mixedFontText
                |> applyStyle 0 mixedFontText.Length (fun style -> { style with FontFamily = Some "Arial"; FontSize = 20.0 })
                |> applyStyle 0 3 (fun style -> { style with FontFamily = Some "Calibri" })
            Metadata = { multiple with SpacingAfter = Some 0.0; SpacingBefore = Some 0.0 } }
    let layout, _ = layoutParagraphAtWithRange settings measureText (rangeMeasurerFromTextMeasurer measureText) 120.0 0.0 0.0 para Map.empty
    layout.Lines |> List.map (fun line -> System.Math.Round(line.Height, 2))
check "a line with Calibri text in an Arial paragraph is Calibri's" mixedFontLines.Head (System.Math.Round(calibriLine, 2))
check "... and the Arial lines after it Arial's" (mixedFontLines.Tail |> List.distinct) [ System.Math.Round(arialLine, 2) ]

// ---------------------------------------------------------------------------
// A page break a paragraph's text starts with (w:br at its start, read as
// PageBreakBefore with PageBreakRunBefore) puts the paragraph on the next
// page - the first paragraph of the document too: Word's first page is then
// blank. (w:pageBreakBefore on the first paragraph starts no page.)
// ---------------------------------------------------------------------------
let breakRunPara id =
    { hintLine 0 with
        Id = id
        Metadata = { (hintLine 0).Metadata with PageBreakBefore = true; Explicit = Doc.ParagraphExplicit.PageBreakRunBefore } }
let propertyBreakPara id = { hintLine 0 with Id = id; Metadata = { (hintLine 0).Metadata with PageBreakBefore = true } }
check "a page break the first paragraph starts with puts it on page 2" ((hintPages [ Doc.Paragraph (breakRunPara "broken") ]) "broken") 1
check "w:pageBreakBefore on the first paragraph starts no page" ((hintPages [ Doc.Paragraph (propertyBreakPara "broken") ]) "broken") 0
check "a page break a later paragraph starts with puts it on the next page" ((hintPages [ Doc.Paragraph (hintLine 1); Doc.Paragraph (breakRunPara "broken") ]) "broken") 1
// (Split, the break stays before the first part.)
let splitBreakDoc =
    let para = breakRunPara "split-break"
    let doc = { Doc.empty with Blocks = [ Doc.Paragraph para ]; Selection = Some { Anchor = { BlockId = para.Id; Offset = 4 }; Focus = { BlockId = para.Id; Offset = 4 } } }
    Doc.splitBlock doc
let splitBreaks =
    splitBreakDoc.Blocks |> List.choose (function Doc.Paragraph para -> Some (para.Metadata.PageBreakBefore, Doc.ParagraphExplicit.has Doc.ParagraphExplicit.PageBreakRunBefore para.Metadata) | _ -> None)
check "a split paragraph's break stays before its first part" splitBreaks [ (true, true); (false, false) ]

// ---------------------------------------------------------------------------
// A paragraph's first line starts where the paragraph does, whatever its size:
// text larger than half the top margin went down by twice its size less the
// margin (sample-2: a 55 pt title on a 39 pt margin started 71 pt down, every
// one-page section ran over on to a second page - 120 pages where Word has 60).
// ---------------------------------------------------------------------------
let largeTextTop =
    let narrowTop = { settings with PaddingTop = 52.0 }
    let para =
        { paragraph "large-title" "Lorem Ipsum" with
            Buffer = ofString "Lorem Ipsum" |> applyStyle 0 11 (fun style -> { style with FontSize = 73.33 })
            Metadata = { Doc.defaultMetadata with SpacingAfter = Some 0.0; SpacingBefore = Some 0.0 } }
    let layout, _ = layoutParagraphAtWithRange narrowTop measureText (rangeMeasurerFromTextMeasurer measureText) 500.0 52.0 0.0 para Map.empty
    layout.Lines.Head.Y
check "a large paragraph's first line is at the paragraph's top" largeTextTop 52.0

printfn ""
printfn "Word layout tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
