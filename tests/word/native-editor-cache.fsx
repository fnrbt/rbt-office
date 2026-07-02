#r "/home/dave/Dev/browser/renderer/bin/Debug/net10.0/Renderer.dll"
#r "/home/dave/Dev/browser/widgets/bin/Debug/net10.0/Widgets.dll"
#r "/home/dave/Dev/browser/office/word-native/bin/Debug/net10.0/WordWidgets.dll"
#r "/home/dave/Dev/browser/office/src/Word/bin/Debug/netstandard2.0/RedBat.Word.dll"
#r "/home/dave/Dev/browser/fsvb/bin/Debug/netstandard2.0/VBScriptParser.dll"

open System.Reflection
open System.Collections.Generic
open Renderer

module Doc = RedBat.Word.Document
module Layout = RedBat.Word.Layout
module WordLayout = RedBat.Word.WordLayout
module Paint = RedBat.Word.Paint
module Shapes = RedBat.Word.Shapes
module Styles = RedBat.Word.WordStyles
module TB = RedBat.Word.TextBuffer

let mutable pass = 0
let mutable fail = 0

let check label got expected =
    if got = expected then
        pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-34s expected %A got %A" label expected got

let pos blockId offset : Doc.Position = { BlockId = blockId; Offset = offset }
let para id text : Doc.ParagraphBlock = { Id = id; Buffer = TB.ofString text; Metadata = Doc.defaultMetadata }
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
let cloneString (text: string) = System.String(text.ToCharArray())

let paragraphCount = 10000
let selectionProbeCount = 80
let blocks =
    [ for index in 0 .. paragraphCount - 1 ->
        Doc.Paragraph (para (sprintf "native-cache-%05d" index) "word cache text") ]

let docWithSelection offset =
    { Doc.empty with
        Blocks = blocks
        Selection = Some { Anchor = pos "native-cache-00000" offset; Focus = pos "native-cache-00000" offset } }

let initialDoc = docWithSelection 0
let state = Composit.Ui.WordEditor.createState initialDoc
let wordEditorType = typeof<Composit.Ui.WordEditor.WordEditorState>.Assembly.GetType("Composit.Ui.WordEditor")
let staticBindingFlags = BindingFlags.NonPublic ||| BindingFlags.Public ||| BindingFlags.Static
let documentModuleType = typeof<Doc.Document>.Assembly.GetType("RedBat.Word.Document")
let cachedLayouts =
    wordEditorType.GetMethod("cachedLayouts", BindingFlags.NonPublic ||| BindingFlags.Static)
let documentHeightCached =
    wordEditorType.GetMethod("documentHeightCached", staticBindingFlags)
let commandsInViewport =
    wordEditorType.GetMethod("commandsInViewport", BindingFlags.NonPublic ||| BindingFlags.Static)
let imageMoveTarget =
    wordEditorType.GetMethod("imageMoveTarget", BindingFlags.NonPublic ||| BindingFlags.Static)
let measureText =
    wordEditorType.GetMethod("measureText", staticBindingFlags)
let glyphLayoutForPaintText =
    wordEditorType.GetMethod("glyphLayoutForPaintText", staticBindingFlags)
let trimHistory =
    wordEditorType.GetMethod("trimHistory", BindingFlags.NonPublic ||| BindingFlags.Static)
let documentsEqualForChange =
    wordEditorType.GetMethod("documentsEqualForChange", BindingFlags.NonPublic ||| BindingFlags.Static)
let splitBlockPreservingStyles =
    wordEditorType.GetMethod("splitBlockPreservingStyles", BindingFlags.NonPublic ||| BindingFlags.Static)
let tryLoadWordFont =
    wordEditorType.GetMethod("tryLoadWordFont", staticBindingFlags)
let wordFontCacheProperty =
    wordEditorType.GetProperty("wordFontCache", staticBindingFlags)
let documentCaptionIdMetrics =
    documentModuleType.GetMethod("captionIdMetrics", BindingFlags.NonPublic ||| BindingFlags.Static)
let documentBookmarkIdMetrics =
    documentModuleType.GetMethod("bookmarkIdMetrics", BindingFlags.NonPublic ||| BindingFlags.Static)
let documentRevisionIdMetrics =
    documentModuleType.GetMethod("revisionIdMetrics", BindingFlags.NonPublic ||| BindingFlags.Static)
let vbaDocumentText =
    typeof<Doc.Document>.Assembly.GetType("RedBat.Word.Vba").GetMethod("documentText", BindingFlags.NonPublic ||| BindingFlags.Static)
let vbaLayoutCacheFor =
    typeof<Doc.Document>.Assembly.GetType("RedBat.Word.Vba").GetMethod("layoutCacheFor", BindingFlags.NonPublic ||| BindingFlags.Static)
let vbaCompileSourceCached =
    typeof<Doc.Document>.Assembly.GetType("RedBat.Word.Vba").GetMethod("compileSourceCached", BindingFlags.NonPublic ||| BindingFlags.Static)
let vbaStatisticsByConstant =
    typeof<Doc.Document>.Assembly.GetType("RedBat.Word.Vba").GetMethod("statisticsByConstant", BindingFlags.NonPublic ||| BindingFlags.Static)

if isNull cachedLayouts then
    fail <- fail + 1
    printfn "FAIL %-34s expected private cachedLayouts method" "native layout cache reflection"
else
    cachedLayouts.Invoke(null, [| box 900.0; box state; box initialDoc |]) |> ignore

    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()

    let sw = System.Diagnostics.Stopwatch.StartNew()
    let mutable cachedProbeCount = 0
    for index in 1 .. selectionProbeCount do
        let doc = docWithSelection (index % 2)
        let result = cachedLayouts.Invoke(null, [| box 900.0; box state; box doc |])
        if not (isNull result) then
            cachedProbeCount <- cachedProbeCount + 1
    sw.Stop()

    check "native layout cache selection probes" cachedProbeCount selectionProbeCount
    check "native layout cache selection perf"
        (sw.Elapsed.TotalMilliseconds < 80.0)
        true

if isNull documentHeightCached then
    fail <- fail + 1
    printfn "FAIL %-34s expected private documentHeightCached method" "native height cache reflection"
else
    let heightParagraphCount = 30000
    let heightDoc =
        { Doc.empty with
            Blocks =
                [ for index in 0 .. heightParagraphCount - 1 ->
                    Doc.Paragraph (para (sprintf "native-height-%05d" index) "height text") ] }
    let heightState = Composit.Ui.WordEditor.createState heightDoc
    let firstHeight = documentHeightCached.Invoke(null, [| box 900.0; box heightState; box heightDoc |]) :?> float

    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()

    let sw = System.Diagnostics.Stopwatch.StartNew()
    let mutable totalHeight = 0.0
    for _ in 1 .. 100 do
        totalHeight <- totalHeight + (documentHeightCached.Invoke(null, [| box 900.0; box heightState; box heightDoc |]) :?> float)
    sw.Stop()

    check "native doc height cache stable"
        (firstHeight > 0.0 && abs (totalHeight - firstHeight * 100.0) < 0.01)
        true
    check "native doc height cache perf"
        (sw.Elapsed.TotalMilliseconds < 15.0)
        true

if isNull measureText then
    fail <- fail + 1
    printfn "FAIL %-34s expected private measureText method" "native measure reflection"
else
    let style = TB.defaultStyle
    let invokeMeasure text =
        measureText.Invoke(null, [| box text; box style |]) :?> float
    let firstWidth = invokeMeasure "Repeated word text 12345"

    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()

    let measureSw = System.Diagnostics.Stopwatch.StartNew()
    let mutable measuredTotal = 0.0
    for _ in 1 .. 20000 do
        measuredTotal <- measuredTotal + invokeMeasure "Repeated word text 12345"
    measureSw.Stop()
    check "native measure cache stable"
        (firstWidth > 0.0 && abs (measuredTotal - firstWidth * 20000.0) < 0.01)
        true
    check "native measure cache perf"
        (measureSw.Elapsed.TotalMilliseconds < 50.0)
        true

if isNull glyphLayoutForPaintText then
    fail <- fail + 1
    printfn "FAIL %-34s expected private glyphLayoutForPaintText method" "native glyph cache reflection"
else
    let paintStyle: Paint.TextPaintStyle = {
        Bold = false
        Italic = false
        Underline = false
        AllCaps = false
        SmallCaps = false
        Strike = false
        FontSize = 11.0
        CharacterSpacing = 0.0
        FontFamily = None
        Color = "#000000"
        Highlight = None
        VertAlign = None
    }
    let firstLayout = glyphLayoutForPaintText.Invoke(null, [| box "Repeated word render text"; box paintStyle |])
    let secondLayout = glyphLayoutForPaintText.Invoke(null, [| box "Repeated word render text"; box paintStyle |])
    check "native glyph layout cache reuse"
        (obj.ReferenceEquals(firstLayout, secondLayout))
        true

    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()

    let glyphSw = System.Diagnostics.Stopwatch.StartNew()
    let mutable glyphProbeCount = 0
    for _ in 1 .. 20000 do
        if not (isNull (glyphLayoutForPaintText.Invoke(null, [| box "Repeated word render text"; box paintStyle |]))) then
            glyphProbeCount <- glyphProbeCount + 1
    glyphSw.Stop()
    check "native glyph layout cache perf"
        (glyphSw.Elapsed.TotalMilliseconds < 50.0 && glyphProbeCount = 20000)
        true

let documentTextCycleDocs =
    [| for docIndex in 0 .. 64 ->
        { Doc.empty with
            Blocks =
                [ for index in 0 .. 499 ->
                    Doc.Paragraph (
                        para
                            (sprintf "native-document-text-source-cycle-%02d-%04d" docIndex index)
                            "alpha beta gamma delta epsilon zeta eta theta iota kappa lambda") ] } |]

for doc in documentTextCycleDocs do
    Doc.getText doc |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let documentTextCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let documentTextCycleSw = System.Diagnostics.Stopwatch.StartNew()
let mutable documentTextCycleTotal = 0L
for _pass in 1 .. 100 do
    for doc in documentTextCycleDocs do
        documentTextCycleTotal <- documentTextCycleTotal + int64 (Doc.getText doc).Length
documentTextCycleSw.Stop()
let documentTextCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - documentTextCycleAllocatedBefore
check "native document text source cycle"
    (documentTextCycleSw.Elapsed.TotalMilliseconds < 80.0
     && documentTextCycleAllocated < 512L * 1024L
     && documentTextCycleTotal > 0L)
    true

let editableEndpointCycleDocs =
    [| for docIndex in 0 .. 64 ->
        { Doc.empty with
            Blocks =
                [ for index in 0 .. 499 ->
                    Doc.Paragraph (
                        para
                            (sprintf "native-select-all-source-cycle-%02d-%04d" docIndex index)
                            "abcdefghijklmnopqrstuvwxyz") ] } |]

for doc in editableEndpointCycleDocs do
    Doc.selectAll doc |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let editableEndpointCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let editableEndpointCycleSw = System.Diagnostics.Stopwatch.StartNew()
let mutable editableEndpointCycleTotal = 0
for _pass in 1 .. 100 do
    for doc in editableEndpointCycleDocs do
        match Doc.selectAll doc with
        | { Selection = Some selection } -> editableEndpointCycleTotal <- editableEndpointCycleTotal + selection.Focus.Offset
        | _ -> ()
editableEndpointCycleSw.Stop()
let editableEndpointCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - editableEndpointCycleAllocatedBefore
check "native select all endpoint source cycle"
    (editableEndpointCycleSw.Elapsed.TotalMilliseconds < 80.0
     && editableEndpointCycleAllocated < 16L * 1024L
     && editableEndpointCycleTotal > 0)
    true

let orderedParagraphCycleDocs =
    [| for docIndex in 0 .. 64 ->
        let ids =
            [| for index in 0 .. 499 ->
                sprintf "native-ordered-source-cycle-%02d-%04d" docIndex index |]
        { Doc.empty with
            Blocks =
                [ for paragraphId in ids ->
                    Doc.Paragraph (para paragraphId "abcdefghijklmnopqrstuvwxyz") ] },
        ids.[0] |]

for doc, blockId in orderedParagraphCycleDocs do
    let selection: Doc.Selection = { Anchor = pos blockId 0; Focus = pos blockId 1 }
    Doc.getNormalizedRange selection doc |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let orderedParagraphCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let orderedParagraphCycleSw = System.Diagnostics.Stopwatch.StartNew()
let mutable orderedParagraphCycleTotal = 0
for pass in 1 .. 100 do
    let offset = pass % 20
    for doc, blockId in orderedParagraphCycleDocs do
        let selection: Doc.Selection =
            { Anchor = pos blockId offset
              Focus = pos blockId (offset + 1) }
        let startPos, endPos, blockIds = Doc.getNormalizedRange selection doc
        orderedParagraphCycleTotal <- orderedParagraphCycleTotal + startPos.Offset + endPos.Offset + blockIds.Length
orderedParagraphCycleSw.Stop()
let orderedParagraphCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - orderedParagraphCycleAllocatedBefore
check "native ordered paragraph source cycle"
    (orderedParagraphCycleSw.Elapsed.TotalMilliseconds < 80.0
     && orderedParagraphCycleAllocated < 1024L * 1024L
     && orderedParagraphCycleTotal > 0)
    true

let contentInfoCycleDocs =
    [| for docIndex in 0 .. 64 ->
        { Doc.empty with
            Blocks =
                [ for index in 0 .. 499 ->
                    Doc.Paragraph {
                        Id = sprintf "native-content-info-source-cycle-%02d-%04d" docIndex index
                        Buffer = TB.ofString "plain content control text"
                        Metadata =
                            { Doc.defaultMetadata with
                                ContentControlId = Some (sprintf "control-%02d-%04d" docIndex index)
                                ContentControlTag = Some "tag"
                                ContentControlAlias = Some "alias" }
                    } ] } |]

for doc in contentInfoCycleDocs do
    Doc.contentControlInfos doc |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let contentInfoCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let contentInfoCycleSw = System.Diagnostics.Stopwatch.StartNew()
let mutable contentInfoCycleTotal = 0
for _pass in 1 .. 100 do
    for doc in contentInfoCycleDocs do
        contentInfoCycleTotal <- contentInfoCycleTotal + (Doc.contentControlInfos doc).Length
contentInfoCycleSw.Stop()
let contentInfoCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - contentInfoCycleAllocatedBefore
check "native content info source cycle"
    (contentInfoCycleSw.Elapsed.TotalMilliseconds < 80.0
     && contentInfoCycleAllocated < 512L * 1024L
     && contentInfoCycleTotal = contentInfoCycleDocs.Length * 100 * 500)
    true

let revisionCacheCycleDocs =
    [| for docIndex in 0 .. 64 ->
        let firstBlockId = sprintf "native-revision-source-cycle-%02d-0000" docIndex
        { Doc.empty with
            Blocks =
                [ for paragraphIndex in 0 .. 499 ->
                    let buffer =
                        TB.ofString "abcdefghijklmnopqrstuvwxyz"
                        |> TB.applyStyle 3 12 (fun style ->
                            { style with
                                RevisionType = Some "insert"
                                RevisionId = Some (sprintf "rev-%02d-%04d" docIndex paragraphIndex)
                                RevisionAuthor = Some "author"
                                RevisionDate = Some "date" })
                    Doc.Paragraph {
                        Id = sprintf "native-revision-source-cycle-%02d-%04d" docIndex paragraphIndex
                        Buffer = buffer
                        Metadata = Doc.defaultMetadata
                    } ]
            Selection = Some { Anchor = pos firstBlockId 0; Focus = pos firstBlockId 0 } } |]

for doc in revisionCacheCycleDocs do
    Doc.revisionInfos doc |> ignore
    Doc.selectNextRevision doc |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let revisionInfoCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let revisionInfoCycleSw = System.Diagnostics.Stopwatch.StartNew()
let mutable revisionInfoCycleTotal = 0
for _pass in 1 .. 100 do
    for doc in revisionCacheCycleDocs do
        revisionInfoCycleTotal <- revisionInfoCycleTotal + (Doc.revisionInfos doc).Length
revisionInfoCycleSw.Stop()
let revisionInfoCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - revisionInfoCycleAllocatedBefore
check "native revision info source cycle"
    (revisionInfoCycleSw.Elapsed.TotalMilliseconds < 80.0
     && revisionInfoCycleAllocated < 512L * 1024L
     && revisionInfoCycleTotal = revisionCacheCycleDocs.Length * 100 * 500)
    true

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let revisionNavigationCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let revisionNavigationCycleSw = System.Diagnostics.Stopwatch.StartNew()
let mutable revisionNavigationCycleTotal = 0
for _pass in 1 .. 100 do
    for doc in revisionCacheCycleDocs do
        match Doc.selectNextRevision doc with
        | { Selection = Some selection } ->
            revisionNavigationCycleTotal <-
                revisionNavigationCycleTotal + selection.Anchor.Offset + selection.Focus.Offset
        | _ -> ()
revisionNavigationCycleSw.Stop()
let revisionNavigationCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - revisionNavigationCycleAllocatedBefore
check "native revision navigation source cycle"
    (revisionNavigationCycleSw.Elapsed.TotalMilliseconds < 80.0
     && revisionNavigationCycleAllocated < 16L * 1024L
     && revisionNavigationCycleTotal = revisionCacheCycleDocs.Length * 100 * (3 + 15))
    true

let pageBreakLookupCycleDocCount = 65
let pageBreakLookupCycleStyleCount = 2000
let pageBreakLookupCycleParagraphCount = 500
let pageBreakLookupCycleDocs =
    [| for docIndex in 0 .. pageBreakLookupCycleDocCount - 1 ->
        let styleIds =
            [| for styleIndex in 0 .. pageBreakLookupCycleStyleCount - 1 ->
                sprintf "pb%02d%04d" docIndex styleIndex |]
        let styles =
            [ for styleIndex in 0 .. pageBreakLookupCycleStyleCount - 1 ->
                { namedStyle styleIds.[styleIndex] None None with
                    StyleType = "paragraph"
                    PageBreakBefore = Some (styleIndex % 8 = 0) } ]
        { Doc.empty with
            Styles = styles
            Blocks =
                [ for paragraphIndex in 0 .. pageBreakLookupCycleParagraphCount - 1 ->
                    Doc.Paragraph {
                        Id = sprintf "native-page-break-source-cycle-%02d-%04d" docIndex paragraphIndex
                        Buffer = TB.ofString "paragraph text"
                        Metadata =
                            { Doc.defaultMetadata with
                                Style = Some styleIds.[paragraphIndex % styleIds.Length] }
                    } ] } |]

for doc in pageBreakLookupCycleDocs do
    Doc.buildPageMap doc |> ignore

let pageBreakLookupCycleExpectedPerDoc =
    let pageMap = Doc.buildPageMap pageBreakLookupCycleDocs.[0]
    pageMap.PageCount + pageMap.BlockAnchors.Length

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let pageBreakLookupCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let pageBreakLookupCycleSw = System.Diagnostics.Stopwatch.StartNew()
let mutable pageBreakLookupCycleTotal = 0
for _pass in 1 .. 20 do
    for doc in pageBreakLookupCycleDocs do
        let pageMap = Doc.buildPageMap doc
        pageBreakLookupCycleTotal <-
            pageBreakLookupCycleTotal + pageMap.PageCount + pageMap.BlockAnchors.Length
pageBreakLookupCycleSw.Stop()
let pageBreakLookupCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - pageBreakLookupCycleAllocatedBefore
check "native page break lookup source cycle"
    (pageBreakLookupCycleSw.Elapsed.TotalMilliseconds < 80.0
     && pageBreakLookupCycleAllocated < 1024L * 1024L
     && pageBreakLookupCycleTotal = pageBreakLookupCycleDocs.Length * 20 * pageBreakLookupCycleExpectedPerDoc)
    true

let styleContextCycleDocCount = 80
let styleContextCycleStyleCount = 160
let styleContextCycleDocs =
    [| for docIndex in 0 .. styleContextCycleDocCount - 1 ->
        { Doc.empty with
            Styles =
                [ for styleIndex in 0 .. styleContextCycleStyleCount - 1 ->
                    namedStyle
                        (sprintf "native-cycle-style-%02d-%03d" docIndex styleIndex)
                        None
                        (Some { TB.defaultStyle with FontSize = 10.0 + float (styleIndex % 9); Bold = styleIndex % 3 = 0 }) ] } |]
for doc in styleContextCycleDocs do
    Styles.styleContext doc |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let styleContextCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let styleContextCycleSw = System.Diagnostics.Stopwatch.StartNew()
let mutable styleContextCycleTotal = 0
for index in 1 .. 800 do
    let context = Styles.styleContext styleContextCycleDocs.[index % styleContextCycleDocs.Length]
    styleContextCycleTotal <-
        styleContextCycleTotal
        + context.Styles.Count
        + context.ResolvedCharacterStyles.Count
styleContextCycleSw.Stop()
let styleContextCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - styleContextCycleAllocatedBefore
check "native style context cache document cycle"
    (styleContextCycleSw.Elapsed.TotalMilliseconds < 80.0
     && styleContextCycleAllocated < 512L * 1024L
     && styleContextCycleTotal = 256000)
    true

let effectiveBufferContext = Styles.styleContext Doc.empty
let effectiveBufferStyle = Styles.normal
let effectiveBufferSource = TB.ofString "abcdef"
let effectiveBufferA = Styles.effectiveBufferInContext effectiveBufferContext effectiveBufferStyle effectiveBufferSource
let effectiveBufferB = Styles.effectiveBufferInContext effectiveBufferContext effectiveBufferStyle effectiveBufferSource
check "native effective buffer cache reuse"
    (obj.ReferenceEquals(effectiveBufferA, effectiveBufferB))
    true

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let effectiveBufferAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let effectiveBufferSw = System.Diagnostics.Stopwatch.StartNew()
let mutable effectiveBufferTotal = 0
for _ in 1 .. 100000 do
    effectiveBufferTotal <-
        effectiveBufferTotal
        + TB.length (Styles.effectiveBufferInContext effectiveBufferContext effectiveBufferStyle effectiveBufferSource)
effectiveBufferSw.Stop()
let effectiveBufferAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - effectiveBufferAllocatedBefore
check "native effective buffer cache hot hit"
    (effectiveBufferSw.Elapsed.TotalMilliseconds < 20.0
     && effectiveBufferAllocated < 8L * 1024L * 1024L
     && effectiveBufferTotal = TB.length effectiveBufferA * 100000)
    true

let effectiveBufferCycleSources =
    [| TB.ofString "abcdef"
       TB.ofString "ghijkl"
       TB.ofString "mnopqr" |]
for source in effectiveBufferCycleSources do
    Styles.effectiveBufferInContext effectiveBufferContext effectiveBufferStyle source |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let effectiveBufferCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let effectiveBufferCycleSw = System.Diagnostics.Stopwatch.StartNew()
let mutable effectiveBufferCycleTotal = 0
for index in 1 .. 300000 do
    effectiveBufferCycleTotal <-
        effectiveBufferCycleTotal
        + TB.length (Styles.effectiveBufferInContext effectiveBufferContext effectiveBufferStyle effectiveBufferCycleSources.[index % effectiveBufferCycleSources.Length])
effectiveBufferCycleSw.Stop()
let effectiveBufferCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - effectiveBufferCycleAllocatedBefore
check "native effective buffer cache three-position cycle"
    (effectiveBufferCycleSw.Elapsed.TotalMilliseconds < 80.0
     && effectiveBufferCycleAllocated < 512L * 1024L
     && effectiveBufferCycleTotal = 1800000)
    true

let effectiveBufferSourceCycleCount = 80
let effectiveBufferSourceCycleRunCount = 1000
let makeEffectiveBufferCycleSource sourceIndex =
    let text = System.String('x', effectiveBufferSourceCycleRunCount)
    let runs =
        [ for index in 0 .. effectiveBufferSourceCycleRunCount - 1 ->
            TB.Text
                { Start = index
                  Length = 1
                  Style =
                    { TB.defaultStyle with
                        Bold = (index + sourceIndex) % 2 = 0
                        FontSize = 10.0 + float ((index + sourceIndex) % 7) } } ]
    { Rope = RedBat.Word.Rope.ofString text; Runs = runs } : TB.TextBuffer

let effectiveBufferSourceCycleSources =
    [| for sourceIndex in 0 .. effectiveBufferSourceCycleCount - 1 ->
        makeEffectiveBufferCycleSource sourceIndex |]

for source in effectiveBufferSourceCycleSources do
    Styles.effectiveBufferInContext effectiveBufferContext effectiveBufferStyle source |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let effectiveBufferSourceCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let effectiveBufferSourceCycleSw = System.Diagnostics.Stopwatch.StartNew()
let mutable effectiveBufferSourceCycleTotal = 0
for _pass in 1 .. 20 do
    for source in effectiveBufferSourceCycleSources do
        effectiveBufferSourceCycleTotal <-
            effectiveBufferSourceCycleTotal
            + TB.length (Styles.effectiveBufferInContext effectiveBufferContext effectiveBufferStyle source)
effectiveBufferSourceCycleSw.Stop()
let effectiveBufferSourceCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - effectiveBufferSourceCycleAllocatedBefore
check "native effective buffer source cycle"
    (effectiveBufferSourceCycleSw.Elapsed.TotalMilliseconds < 80.0
     && effectiveBufferSourceCycleAllocated < 512L * 1024L
     && effectiveBufferSourceCycleTotal = 1600000)
    true

let runArraySourceCycleCount = 80
let runArraySourceCycleRunCount = 1000
let makeRunArrayCycleSource sourceIndex =
    let text = System.String('x', runArraySourceCycleRunCount)
    let runs =
        [ for index in 0 .. runArraySourceCycleRunCount - 1 ->
            TB.Text
                { Start = index
                  Length = 1
                  Style =
                    { TB.defaultStyle with
                        Bold = (index + sourceIndex) % 2 = 0
                        FontSize = 10.0 + float ((index + sourceIndex) % 7) } } ]
    { Rope = RedBat.Word.Rope.ofString text; Runs = runs } : TB.TextBuffer

let runArraySourceCycleSources =
    [| for sourceIndex in 0 .. runArraySourceCycleCount - 1 ->
        makeRunArrayCycleSource sourceIndex |]

for source in runArraySourceCycleSources do
    TB.getStyleAt (runArraySourceCycleRunCount / 2) source |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let runArraySourceCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let runArraySourceCycleSw = System.Diagnostics.Stopwatch.StartNew()
let mutable runArraySourceCycleTotal = 0.0
for _pass in 1 .. 20 do
    for source in runArraySourceCycleSources do
        let style = TB.getStyleAt (runArraySourceCycleRunCount / 2) source
        runArraySourceCycleTotal <-
            runArraySourceCycleTotal
            + style.FontSize
            + (if style.Bold then 1.0 else 0.0)
runArraySourceCycleSw.Stop()
let runArraySourceCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - runArraySourceCycleAllocatedBefore
check "native text buffer run array source cycle"
    (runArraySourceCycleSw.Elapsed.TotalMilliseconds < 80.0
     && runArraySourceCycleAllocated < 512L * 1024L
     && abs (runArraySourceCycleTotal - 21660.0) < 0.01)
    true

let rendererCommandParagraphCount = 5000
let rendererCommandDoc =
    { Doc.empty with
        Blocks =
            [ for index in 0 .. rendererCommandParagraphCount - 1 ->
                Doc.Paragraph (para (sprintf "native-render-%05d" index) "word render text") ]
        Selection = None }
let rendererCommandState = Composit.Ui.WordEditor.createState rendererCommandDoc
let noImage : string -> Framebuffer option = fun _ -> None
let rendererCommandsA =
    Composit.Ui.WordEditor.commands 900.0 600.0 noImage None rendererCommandState rendererCommandDoc
let rendererCommandsB =
    Composit.Ui.WordEditor.commands 900.0 600.0 noImage None rendererCommandState rendererCommandDoc
check "native renderer command cache reuse"
    (obj.ReferenceEquals(rendererCommandsA, rendererCommandsB))
    true

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let rendererCommandsSw = System.Diagnostics.Stopwatch.StartNew()
let mutable rendererCommandTotal = 0
for _ in 1 .. 10 do
    rendererCommandTotal <-
        rendererCommandTotal
        + (Composit.Ui.WordEditor.commands 900.0 600.0 noImage None rendererCommandState rendererCommandDoc).Length
rendererCommandsSw.Stop()
check "native renderer command cache perf"
    (rendererCommandsSw.Elapsed.TotalMilliseconds < 20.0
     && rendererCommandTotal = rendererCommandsA.Length * 10)
    true

let documentPaintSettings : Layout.LayoutSettings =
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
let documentPaintMeasureText (text: string) _style =
    float text.Length * 7.0
let documentPaintLine : Layout.MeasuredLine =
    { Segments = [ Layout.TextSegment ("abcdef", TB.defaultStyle, 0.0, 42.0, 0) ]
      Y = 0.0
      Width = 42.0
      Height = 24.0
      StartOffset = 0
      EndOffset = 6 }
let alternatingLineLookupLines =
    [ for index in 0 .. 999 ->
        { documentPaintLine with
            Y = float index * 24.0
            StartOffset = index * 10
            EndOffset = (index + 1) * 10 } ]
let alternatingLineLookupLayout : Layout.BlockLayout =
    { BlockId = "native-line-lookup-cache"
      Lines = alternatingLineLookupLines
      Y = 0.0
      Height = 24000.0
      AvailableWidth = 120.0
      Metadata = Doc.defaultMetadata
      Prefix = ""
      PrefixStyle = None
      PrefixWidth = 0.0
      Indent = 0.0
      VAlignOffset = 0.0 }
Layout.lineForOffset alternatingLineLookupLayout 9999 |> ignore
Layout.lineForOffset alternatingLineLookupLayout 10 |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let alternatingLineLookupSw = System.Diagnostics.Stopwatch.StartNew()
let mutable alternatingLineLookupChecksum = 0L
for index in 1 .. 500000 do
    let offset = if index % 2 = 0 then 9999 else 10
    alternatingLineLookupChecksum <-
        alternatingLineLookupChecksum
        + int64 (Layout.lineForOffset alternatingLineLookupLayout offset).EndOffset
alternatingLineLookupSw.Stop()
check "native alternating line lookup cache"
    (alternatingLineLookupSw.Elapsed.TotalMilliseconds < 30.0
     && alternatingLineLookupChecksum = 2505000000L)
    true
let lineLookupCycleOffsets = [| 9999; 10; 5010 |]
for offset in lineLookupCycleOffsets do
    Layout.lineForOffset alternatingLineLookupLayout offset |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let lineLookupCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let lineLookupCycleSw = System.Diagnostics.Stopwatch.StartNew()
let mutable lineLookupCycleChecksum = 0L
for index in 1 .. 500000 do
    lineLookupCycleChecksum <-
        lineLookupCycleChecksum
        + int64 (Layout.lineForOffset alternatingLineLookupLayout lineLookupCycleOffsets.[index % lineLookupCycleOffsets.Length]).EndOffset
lineLookupCycleSw.Stop()
let lineLookupCycleAllocated = System.GC.GetAllocatedBytesForCurrentThread() - lineLookupCycleAllocatedBefore
check "native line lookup cache three-position cycle"
    (lineLookupCycleSw.Elapsed.TotalMilliseconds < 30.0
     && lineLookupCycleAllocated < 512L * 1024L
     && lineLookupCycleChecksum = 2506661680L)
    true

let lineArraySourceCycleCount = 80
let lineArraySourceCycleLineCount = 1000
let makeLineArraySourceCycleLayout layoutIndex =
    let lines =
        [ for lineIndex in 0 .. lineArraySourceCycleLineCount - 1 ->
            { documentPaintLine with
                Y = float lineIndex * 24.0
                StartOffset = lineIndex * 10
                EndOffset = (lineIndex + 1) * 10 } ]
    { alternatingLineLookupLayout with
        BlockId = sprintf "native-line-array-source-cycle-%02d" layoutIndex
        Lines = lines
        Height = float lineArraySourceCycleLineCount * 24.0 }
let lineArraySourceCycleLayouts =
    [| for layoutIndex in 0 .. lineArraySourceCycleCount - 1 ->
        makeLineArraySourceCycleLayout layoutIndex |]
for layout in lineArraySourceCycleLayouts do
    Layout.lineForOffset layout 9999 |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let lineArraySourceCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let lineArraySourceCycleSw = System.Diagnostics.Stopwatch.StartNew()
let mutable lineArraySourceCycleChecksum = 0L
for _pass in 1 .. 20 do
    for layout in lineArraySourceCycleLayouts do
        lineArraySourceCycleChecksum <-
            lineArraySourceCycleChecksum + int64 (Layout.lineForOffset layout 9999).EndOffset
lineArraySourceCycleSw.Stop()
let lineArraySourceCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - lineArraySourceCycleAllocatedBefore
check "native line array cache source cycle"
    (lineArraySourceCycleSw.Elapsed.TotalMilliseconds < 80.0
     && lineArraySourceCycleAllocated < 512L * 1024L
     && lineArraySourceCycleChecksum = 16000000L)
    true
let hitTestMeasureRange (_text: string) (_style: TB.TextStyle) (_start: int) (length: int) =
    float length
let hitTestLongLineCount = 5000
let hitTestLongLines =
    [ for index in 0 .. hitTestLongLineCount - 1 ->
        let startOffset = index * 10
        { documentPaintLine with
            Segments = [ Layout.TextSegment ("abcdef", TB.defaultStyle, 0.0, 6.0, startOffset) ]
            Y = float index * 24.0
            Width = 6.0
            StartOffset = startOffset
            EndOffset = startOffset + 6 } ]
let hitTestLongLayout : Layout.BlockLayout =
    { alternatingLineLookupLayout with
        BlockId = "native-hit-test-long-lines"
        Lines = hitTestLongLines
        Height = float hitTestLongLineCount * 24.0 }
let hitTestLongLayouts = [ hitTestLongLayout ]
let hitTestLongY = float (hitTestLongLineCount - 1) * 24.0 + 2.0
Layout.hitTestLayoutsWithRange hitTestMeasureRange 0.0 hitTestLongLayouts 3.0 hitTestLongY |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let hitTestLongSw = System.Diagnostics.Stopwatch.StartNew()
let mutable hitTestLongTotal = 0
for _ in 1 .. 20000 do
    match Layout.hitTestLayoutsWithRange hitTestMeasureRange 0.0 hitTestLongLayouts 3.0 hitTestLongY with
    | Some (_, offset) -> hitTestLongTotal <- hitTestLongTotal + offset
    | None -> ()
hitTestLongSw.Stop()
check "native long line hit-test cache"
    (hitTestLongSw.Elapsed.TotalMilliseconds < 80.0
     && hitTestLongTotal = 999860000)
    true
let hitTestNonMonotonicLines =
    let high =
        { documentPaintLine with
            Segments = [ Layout.TextSegment ("high", TB.defaultStyle, 0.0, 4.0, 10) ]
            Y = 1000.0
            Width = 4.0
            StartOffset = 10
            EndOffset = 14 }
    let low =
        { documentPaintLine with
            Segments = [ Layout.TextSegment ("low", TB.defaultStyle, 0.0, 3.0, 20) ]
            Y = 0.0
            Width = 3.0
            StartOffset = 20
            EndOffset = 23 }
    let rest =
        [ for index in 2 .. 40 ->
            { documentPaintLine with
                Y = 1000.0 + float index * 24.0
                StartOffset = index * 10
                EndOffset = index * 10 + 6 } ]
    high :: low :: rest
let hitTestNonMonotonicLayout : Layout.BlockLayout =
    { alternatingLineLookupLayout with
        BlockId = "native-hit-test-nonmonotonic"
        Lines = hitTestNonMonotonicLines
        Height = 2400.0 }
check "native hit-test nonmonotonic fallback"
    (Layout.hitTestLayoutsWithRange hitTestMeasureRange 0.0 [hitTestNonMonotonicLayout] 3.0 2.0)
    (Some ("native-hit-test-nonmonotonic", 23))
let xOffsetSegmentCount = 256
let xOffsetSegmentText = "abcde"
let xOffsetLine : Layout.MeasuredLine =
    { Segments =
        [ for index in 0 .. xOffsetSegmentCount - 1 ->
            Layout.TextSegment (
                xOffsetSegmentText,
                TB.defaultStyle,
                float (index * xOffsetSegmentText.Length),
                float xOffsetSegmentText.Length,
                index * xOffsetSegmentText.Length) ]
      Y = 0.0
      Width = float (xOffsetSegmentCount * xOffsetSegmentText.Length)
      Height = 24.0
      StartOffset = 0
      EndOffset = xOffsetSegmentCount * xOffsetSegmentText.Length }
let xOffsetMeasureRange (_text: string) (_style: TB.TextStyle) (_start: int) (length: int) =
    float length
let xOffsetHotOffset = xOffsetLine.EndOffset - 1
let xOffsetHotX = Layout.xForOffsetWithRange xOffsetMeasureRange xOffsetLine xOffsetHotOffset

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let xOffsetAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let xOffsetSw = System.Diagnostics.Stopwatch.StartNew()
let mutable xOffsetTotal = 0.0
for _ in 1 .. 200000 do
    xOffsetTotal <- xOffsetTotal + Layout.xForOffsetWithRange xOffsetMeasureRange xOffsetLine xOffsetHotOffset
xOffsetSw.Stop()
let xOffsetAllocated = System.GC.GetAllocatedBytesForCurrentThread() - xOffsetAllocatedBefore
check "native x offset cache hot hit"
    (xOffsetSw.Elapsed.TotalMilliseconds < 80.0
     && xOffsetAllocated < 1024L * 1024L
     && xOffsetHotX = 1279.0
     && xOffsetTotal = xOffsetHotX * 200000.0)
    true

let xOffsetCycleOffsets = [| 1279; 640; 7 |]
for offset in xOffsetCycleOffsets do
    Layout.xForOffsetWithRange xOffsetMeasureRange xOffsetLine offset |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let xOffsetCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let xOffsetCycleSw = System.Diagnostics.Stopwatch.StartNew()
let mutable xOffsetCycleTotal = 0.0
for index in 1 .. 200000 do
    xOffsetCycleTotal <-
        xOffsetCycleTotal
        + Layout.xForOffsetWithRange xOffsetMeasureRange xOffsetLine xOffsetCycleOffsets.[index % xOffsetCycleOffsets.Length]
xOffsetCycleSw.Stop()
let xOffsetCycleAllocated = System.GC.GetAllocatedBytesForCurrentThread() - xOffsetCycleAllocatedBefore
check "native x offset cache three-position cycle"
    (xOffsetCycleSw.Elapsed.TotalMilliseconds < 40.0
     && xOffsetCycleAllocated < 512L * 1024L
     && xOffsetCycleTotal = 128399363.0)
    true
let xOffsetOversizedStyleWeakReference () =
    let payload = System.String('X', 6 * 1024 * 1024)
    let line : Layout.MeasuredLine =
        { xOffsetLine with
            Segments =
                [ Layout.TextSegment (
                    "x",
                    { TB.defaultStyle with CommentText = Some payload },
                    0.0,
                    1.0,
                    0) ]
            Width = 1.0
            EndOffset = 1 }
    Layout.xForOffsetWithRange xOffsetMeasureRange line 1 |> ignore
    System.WeakReference(payload)
let xOffsetOversizedStyleWeak = xOffsetOversizedStyleWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
check "native x offset cache skips oversized style"
    xOffsetOversizedStyleWeak.IsAlive
    false

let shapePresetPathCycleSizes =
    [| for index in 0 .. 128 ->
        100.0 + float index, 80.0 + float (index % 17) |]

for w, h in shapePresetPathCycleSizes do
    Shapes.presetPath "star32" w h |> ignore

let shapePresetPathCycleExpectedLength =
    (Shapes.presetPath "star32" 100.0 80.0).Length

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let shapePresetPathCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let shapePresetPathCycleSw = System.Diagnostics.Stopwatch.StartNew()
let mutable shapePresetPathCycleTotal = 0
for _pass in 1 .. 1000 do
    for w, h in shapePresetPathCycleSizes do
        shapePresetPathCycleTotal <- shapePresetPathCycleTotal + (Shapes.presetPath "star32" w h).Length
shapePresetPathCycleSw.Stop()
let shapePresetPathCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - shapePresetPathCycleAllocatedBefore
check "native shape preset path source cycle"
    (shapePresetPathCycleSw.Elapsed.TotalMilliseconds < 100.0
     && shapePresetPathCycleAllocated < 512L * 1024L
     && shapePresetPathCycleTotal = shapePresetPathCycleSizes.Length * 1000 * shapePresetPathCycleExpectedLength)
    true

let documentPaintLayout : Layout.BlockLayout =
    { BlockId = "native-document-paint-cache"
      Lines = [ documentPaintLine ]
      Y = 0.0
      Height = documentPaintLine.Height
      AvailableWidth = 120.0
      Metadata = Doc.defaultMetadata
      Prefix = ""
      PrefixStyle = None
      PrefixWidth = 0.0
      Indent = 0.0
      VAlignOffset = 0.0 }
let documentPaintLayouts = [ documentPaintLayout ]
let documentPaintCachedLayouts = Some documentPaintLayouts
let documentPaintDoc =
    { Doc.empty with
        Blocks = [ Doc.Paragraph (para "native-document-paint-cache" "abcdef") ]
        Selection = None }
let documentPaintRangeMeasure = Layout.rangeMeasurerFromTextMeasurer documentPaintMeasureText
let _, documentPaintCommandsA, _ =
    Paint.collectDocumentCommandsWithLayouts
        documentPaintSettings
        "Calibri"
        documentPaintMeasureText
        documentPaintRangeMeasure
        documentPaintDoc
        documentPaintCachedLayouts
let _, documentPaintCommandsB, _ =
    Paint.collectDocumentCommandsWithLayouts
        documentPaintSettings
        "Calibri"
        documentPaintMeasureText
        documentPaintRangeMeasure
        documentPaintDoc
        documentPaintCachedLayouts
check "native document paint cache reuse"
    (obj.ReferenceEquals(documentPaintCommandsA, documentPaintCommandsB))
    true

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let documentPaintProbeCount = 100000
let documentPaintAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let documentPaintSw = System.Diagnostics.Stopwatch.StartNew()
let mutable documentPaintTotal = 0
for _ in 1 .. documentPaintProbeCount do
    let _, commands, _ =
        Paint.collectDocumentCommandsWithLayouts
            documentPaintSettings
            "Calibri"
            documentPaintMeasureText
            documentPaintRangeMeasure
            documentPaintDoc
            documentPaintCachedLayouts
    documentPaintTotal <- documentPaintTotal + commands.Length
documentPaintSw.Stop()
let documentPaintAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - documentPaintAllocatedBefore
check "native document paint cache hot hit"
    (documentPaintSw.Elapsed.TotalMilliseconds < 40.0
     && documentPaintAllocated < 4L * 1024L * 1024L
     && documentPaintTotal = documentPaintCommandsA.Length * documentPaintProbeCount)
    true

let documentPaintCycleSources =
    [| for index in 0 .. 79 ->
        let blockId = sprintf "native-document-paint-cycle-%02d" index
        let line = documentPaintLine
        let layout : Layout.BlockLayout =
            { BlockId = blockId
              Lines = [ line ]
              Y = 0.0
              Height = line.Height
              AvailableWidth = 120.0
              Metadata = Doc.defaultMetadata
              Prefix = ""
              PrefixStyle = None
              PrefixWidth = 0.0
              Indent = 0.0
              VAlignOffset = 0.0 }
        { Doc.empty with
            Blocks = [ Doc.Paragraph (para blockId "abcdef") ]
            Selection = None },
        [ layout ] |]

for doc, layouts in documentPaintCycleSources do
    Paint.collectDocumentCommandsWithLayouts
        documentPaintSettings
        "Calibri"
        documentPaintMeasureText
        documentPaintRangeMeasure
        doc
        (Some layouts)
    |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let documentPaintCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let documentPaintCycleSw = System.Diagnostics.Stopwatch.StartNew()
let mutable documentPaintCycleTotal = 0
for _pass in 1 .. 500 do
    for doc, layouts in documentPaintCycleSources do
        let _, commands, _ =
            Paint.collectDocumentCommandsWithLayouts
                documentPaintSettings
                "Calibri"
                documentPaintMeasureText
                documentPaintRangeMeasure
                doc
                (Some layouts)
        documentPaintCycleTotal <- documentPaintCycleTotal + commands.Length
documentPaintCycleSw.Stop()
let documentPaintCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - documentPaintCycleAllocatedBefore
check "native document paint cache document cycle"
    (documentPaintCycleSw.Elapsed.TotalMilliseconds < 80.0
     && documentPaintCycleAllocated < 4L * 1024L * 1024L
     && documentPaintCycleTotal = 40000)
    true

let selectionLayoutLookupCycleSources =
    [| for sourceIndex in 0 .. 64 ->
        let blocks, layouts =
            [ for layoutIndex in 0 .. 499 ->
                let blockId = sprintf "native-layout-lookup-cycle-%02d-%04d" sourceIndex layoutIndex
                let line =
                    { documentPaintLine with
                        Y = float layoutIndex * documentPaintLine.Height }
                Doc.Paragraph (para blockId "abcdef"),
                ({ BlockId = blockId
                   Lines = [ line ]
                   Y = line.Y
                   Height = line.Height
                   AvailableWidth = 120.0
                   Metadata = Doc.defaultMetadata
                   Prefix = ""
                   PrefixStyle = None
                   PrefixWidth = 0.0
                   Indent = 0.0
                   VAlignOffset = 0.0 } : Layout.BlockLayout) ]
            |> List.unzip
        let targetBlockId = sprintf "native-layout-lookup-cycle-%02d-%04d" sourceIndex 499
        { Doc.empty with
            Blocks = blocks
            Selection = Some { Anchor = pos targetBlockId 3; Focus = pos targetBlockId 3 } },
        layouts |]

for doc, layouts in selectionLayoutLookupCycleSources do
    Paint.selectionCommandsWithRange documentPaintSettings documentPaintRangeMeasure doc layouts |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let selectionLayoutLookupCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let selectionLayoutLookupCycleSw = System.Diagnostics.Stopwatch.StartNew()
let mutable selectionLayoutLookupCycleTotal = 0
for _pass in 1 .. 1000 do
    for doc, layouts in selectionLayoutLookupCycleSources do
        selectionLayoutLookupCycleTotal <-
            selectionLayoutLookupCycleTotal
            + (Paint.selectionCommandsWithRange documentPaintSettings documentPaintRangeMeasure doc layouts).Length
selectionLayoutLookupCycleSw.Stop()
let selectionLayoutLookupCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - selectionLayoutLookupCycleAllocatedBefore
check "native selection layout lookup source cycle"
    (selectionLayoutLookupCycleSw.Elapsed.TotalMilliseconds < 180.0
     && selectionLayoutLookupCycleAllocated < 512L * 1024L
     && selectionLayoutLookupCycleTotal = selectionLayoutLookupCycleSources.Length * 1000)
    true

let markupMarkerCycleSources: Layout.BlockLayout list array =
    [| for index in 0 .. 79 ->
        [ for blockIndex in 0 .. 9 ->
            let line =
                { documentPaintLine with
                    Y = float blockIndex * 24.0 }
            { BlockId = sprintf "native-marker-cycle-%02d-%02d" index blockIndex
              Lines = [ line ]
              Y = float blockIndex * 24.0
              Height = line.Height
              AvailableWidth = 120.0
              Metadata =
                { Doc.defaultMetadata with
                    ContentControlId = Some (sprintf "cc-%02d-%02d" index blockIndex) }
              Prefix = ""
              PrefixStyle = None
              PrefixWidth = 0.0
              Indent = 0.0
              VAlignOffset = 0.0 } ] |]

for layouts in markupMarkerCycleSources do
    Paint.markupMarkerCommands documentPaintSettings layouts |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let markupMarkerCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let markupMarkerCycleSw = System.Diagnostics.Stopwatch.StartNew()
let mutable markupMarkerCycleTotal = 0
for _pass in 1 .. 500 do
    for layouts in markupMarkerCycleSources do
        markupMarkerCycleTotal <-
            markupMarkerCycleTotal
            + (Paint.markupMarkerCommands documentPaintSettings layouts).Length
markupMarkerCycleSw.Stop()
let markupMarkerCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - markupMarkerCycleAllocatedBefore
check "native markup marker cache entry cycle"
    (markupMarkerCycleSw.Elapsed.TotalMilliseconds < 80.0
     && markupMarkerCycleAllocated < 1024L * 1024L
     && markupMarkerCycleTotal = 800000)
    true

if isNull commandsInViewport then
    fail <- fail + 1
    printfn "FAIL %-34s expected private commandsInViewport method" "native viewport commands reflection"
else
    let viewportCommandDoc =
        { Doc.empty with
            Blocks =
                [ for index in 0 .. 10000 - 1 ->
                    Doc.Paragraph (para (sprintf "native-viewport-%05d" index) "word render text") ]
            Selection = None }
    let viewportCommandState = Composit.Ui.WordEditor.createState viewportCommandDoc
    documentHeightCached.Invoke(null, [| box 900.0; box viewportCommandState; box viewportCommandDoc |]) |> ignore
    let invokeViewportCommands scrollY =
        commandsInViewport.Invoke(null, [| box 900.0; box 600.0; box scrollY; box noImage; box None; box viewportCommandState; box viewportCommandDoc |]) :?> Renderer.Cmd list
    let viewportCommandsA = invokeViewportCommands 200000.0
    let viewportCommandsB = invokeViewportCommands 200000.0
    check "native viewport command size"
        (obj.ReferenceEquals(viewportCommandsA, viewportCommandsB)
         && viewportCommandsA.Length < 100)
        true

    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()

    let viewportAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
    let viewportSw = System.Diagnostics.Stopwatch.StartNew()
    let mutable viewportTotal = 0
    for index in 0 .. 19 do
        viewportTotal <- viewportTotal + (invokeViewportCommands (200000.0 + float (index % 2) * 600.0)).Length
    viewportSw.Stop()
    let viewportAllocated = System.GC.GetAllocatedBytesForCurrentThread() - viewportAllocatedBefore
    check "native viewport command alloc"
        (viewportSw.Elapsed.TotalMilliseconds < 40.0
         && viewportAllocated < 256L * 1024L
         && viewportTotal < 2000)
        true

if isNull imageMoveTarget then
    fail <- fail + 1
    printfn "FAIL %-34s expected private imageMoveTarget method" "native image target reflection"
else
    let imageTargetParagraphCount = 30000
    let imageTargetDoc =
        { Doc.empty with
            Blocks =
                [ for index in 0 .. imageTargetParagraphCount - 1 ->
                    Doc.Paragraph (para (sprintf "native-image-target-%05d" index) "image target text") ]
            Selection = None }
    let imageTargetState = Composit.Ui.WordEditor.createState imageTargetDoc
    documentHeightCached.Invoke(null, [| box 900.0; box imageTargetState; box imageTargetDoc |]) |> ignore
    let invokeImageMoveTarget () =
        imageMoveTarget.Invoke(
            null,
            [|
                box 900.0
                box 160.0
                box 600000.0
                box "native-image-target-00000"
                box 0
                box imageTargetState
                box imageTargetDoc
            |]) :?> Doc.Position
    let firstImageTarget = invokeImageMoveTarget ()
    check "native image target resolves"
        (firstImageTarget.BlockId.StartsWith("native-image-target-"))
        true

    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()

    let imageTargetAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
    let imageTargetSw = System.Diagnostics.Stopwatch.StartNew()
    let mutable imageTargetChecksum = 0
    for _ in 1 .. 20 do
        let target = invokeImageMoveTarget ()
        imageTargetChecksum <- imageTargetChecksum + target.Offset + target.BlockId.Length
    imageTargetSw.Stop()
    let imageTargetAllocated = System.GC.GetAllocatedBytesForCurrentThread() - imageTargetAllocatedBefore
    check "native image target cached hit-test"
        (imageTargetSw.Elapsed.TotalMilliseconds < 80.0
         && imageTargetAllocated < 8L * 1024L * 1024L
         && imageTargetChecksum > 0)
        true

let rendererSelectionDoc =
    { rendererCommandDoc with
        Selection =
            Some {
                Anchor = pos "native-render-00000" 0
                Focus = pos "native-render-00000" 0
            } }
let rendererSelectionState = Composit.Ui.WordEditor.createState rendererSelectionDoc
let rendererSelectionCommandsA =
    Composit.Ui.WordEditor.commands 900.0 600.0 noImage None rendererSelectionState rendererSelectionDoc
let rendererSelectionCommandsB =
    Composit.Ui.WordEditor.commands 900.0 600.0 noImage None rendererSelectionState rendererSelectionDoc
check "native selection command cache reuse"
    (obj.ReferenceEquals(rendererSelectionCommandsA, rendererSelectionCommandsB))
    true

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let rendererSelectionAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let rendererSelectionSw = System.Diagnostics.Stopwatch.StartNew()
let mutable rendererSelectionTotal = 0
for _ in 1 .. 50 do
    rendererSelectionTotal <-
        rendererSelectionTotal
        + (Composit.Ui.WordEditor.commands 900.0 600.0 noImage None rendererSelectionState rendererSelectionDoc).Length
rendererSelectionSw.Stop()
let rendererSelectionAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - rendererSelectionAllocatedBefore
check "native selection command cache alloc"
    (rendererSelectionSw.Elapsed.TotalMilliseconds < 20.0
     && rendererSelectionAllocated < 1024L * 1024L
     && rendererSelectionTotal = rendererSelectionCommandsA.Length * 50)
    true

let rendererMovingSelectionState = Composit.Ui.WordEditor.createState rendererSelectionDoc
let rendererMovingSelectionCommandLength =
    (Composit.Ui.WordEditor.commands 900.0 600.0 noImage None rendererMovingSelectionState rendererSelectionDoc).Length

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let rendererMovingSelectionAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let rendererMovingSelectionSw = System.Diagnostics.Stopwatch.StartNew()
let mutable rendererMovingSelectionTotal = 0
for index in 0 .. 49 do
    let offset = index % 2
    let doc =
        { rendererCommandDoc with
            Selection =
                Some {
                    Anchor = pos "native-render-00000" offset
                    Focus = pos "native-render-00000" offset
                } }
    rendererMovingSelectionTotal <-
        rendererMovingSelectionTotal
        + (Composit.Ui.WordEditor.commands 900.0 600.0 noImage None rendererMovingSelectionState doc).Length
rendererMovingSelectionSw.Stop()
let rendererMovingSelectionAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - rendererMovingSelectionAllocatedBefore
check "native moving selection command alloc"
    (rendererMovingSelectionSw.Elapsed.TotalMilliseconds < 20.0
     && rendererMovingSelectionAllocated < 128L * 1024L
     && rendererMovingSelectionTotal = rendererMovingSelectionCommandLength * 50)
    true

let rendererFullSelectionDoc =
    { rendererCommandDoc with
        Selection =
            Some {
                Anchor = pos "native-render-00000" 0
                Focus = pos (sprintf "native-render-%05d" (rendererCommandParagraphCount - 1)) 16
            } }
let rendererFullSelectionState = Composit.Ui.WordEditor.createState rendererFullSelectionDoc
let rendererFullSelectionCommandsA =
    Composit.Ui.WordEditor.commands 900.0 600.0 noImage None rendererFullSelectionState rendererFullSelectionDoc
let rendererFullSelectionCommandsB =
    Composit.Ui.WordEditor.commands 900.0 600.0 noImage None rendererFullSelectionState rendererFullSelectionDoc
check "native full selection command cache reuse"
    (obj.ReferenceEquals(rendererFullSelectionCommandsA, rendererFullSelectionCommandsB))
    true

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let rendererFullSelectionAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let rendererFullSelectionSw = System.Diagnostics.Stopwatch.StartNew()
let mutable rendererFullSelectionTotal = 0
for _ in 1 .. 20 do
    rendererFullSelectionTotal <-
        rendererFullSelectionTotal
        + (Composit.Ui.WordEditor.commands 900.0 600.0 noImage None rendererFullSelectionState rendererFullSelectionDoc).Length
rendererFullSelectionSw.Stop()
let rendererFullSelectionAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - rendererFullSelectionAllocatedBefore
check "native full selection command cache alloc"
    (rendererFullSelectionSw.Elapsed.TotalMilliseconds < 20.0
     && rendererFullSelectionAllocated < 1024L * 1024L
     && rendererFullSelectionTotal = rendererFullSelectionCommandsA.Length * 20)
    true

if isNull trimHistory then
    fail <- fail + 1
    printfn "FAIL %-34s expected private trimHistory method" "native history trim reflection"
else
    let largeHistoryDocCount = 12
    let largeHistoryTextChars = 2 * 1024 * 1024
    let largeHistoryDocs =
        [ for index in 0 .. largeHistoryDocCount - 1 ->
            let blockId = sprintf "native-history-%02d" index
            { Doc.empty with
                Blocks =
                    [ Doc.Paragraph
                        (para
                            blockId
                            (System.String(char (int 'A' + (index % 26)), largeHistoryTextChars))) ] } ]
    let sw = System.Diagnostics.Stopwatch.StartNew()
    let trimmed = trimHistory.Invoke(null, [| box largeHistoryDocs |]) :?> Doc.Document list
    sw.Stop()
    let retainedTextChars =
        trimmed
        |> List.sumBy (fun doc ->
            match doc.Blocks with
            | [ Doc.Paragraph para ] -> TB.length para.Buffer
            | _ -> 0)
    check "native history trims large docs"
        (trimmed.Length < largeHistoryDocs.Length && retainedTextChars <= 20 * 1024 * 1024)
        true
    check "native history trim perf"
        (sw.Elapsed.TotalMilliseconds < 80.0)
        true

    let largeMetadataDocCount = 6
    let largeMetadataPayloadChars = 256 * 1024
    let metadataPayload index label =
        let suffix = sprintf "-%02d-%s" index label
        System.String(char (int 'a' + (index % 26)), largeMetadataPayloadChars - suffix.Length) + suffix
    let metadataBorder index label : Doc.TableBorder = {
        Style = Some (metadataPayload index (label + "-style"))
        Color = Some (metadataPayload index (label + "-color"))
        Width = Some 1.0
    }
    let metadataBorders index label : Doc.TableBorders = {
        Top = Some (metadataBorder index (label + "-top"))
        Right = Some (metadataBorder index (label + "-right"))
        Bottom = Some (metadataBorder index (label + "-bottom"))
        Left = Some (metadataBorder index (label + "-left"))
        InsideH = Some (metadataBorder index (label + "-inside-h"))
        InsideV = Some (metadataBorder index (label + "-inside-v"))
    }
    let metadataShading index label : Doc.Shading = {
        Pattern = Some (metadataPayload index (label + "-pattern"))
        ForegroundColor = Some (metadataPayload index (label + "-foreground"))
        BackgroundColor = Some (metadataPayload index (label + "-background"))
    }
    let metadataTable index : Doc.TableBlock = {
        Id = sprintf "native-metadata-table-%02d" index
        Rows = [
            {
                Cells = [
                    {
                        Content = [ para (sprintf "native-metadata-cell-%02d" index) "x" ]
                        NestedTables = []
                        NestedTableAnchors = []
                        ColSpan = 1
                        BackgroundColor = None
                        Shading = Some (metadataShading index "cell-shading")
                        Width = None
                        PreferredWidth = None
                        BorderColor = None
                        BorderWidth = Some 1.0
                        Borders = Some (metadataBorders index "cell-borders")
                        VerticalAlignment = None
                        VerticalMerge = None
                        PaddingTop = None
                        PaddingRight = None
                        PaddingBottom = None
                        PaddingLeft = None
                    }
                ]
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
                BordersOverride = Some (metadataBorders index "row-borders")
            }
        ]
        PreferredWidth = None
        Alignment = None
        LayoutMode = None
        GridColumns = []
        Borders = Some (metadataBorders index "table-borders")
        BackgroundColor = None
        Shading = Some (metadataShading index "table-shading")
        CellSpacing = None
        PaddingTop = None
        PaddingRight = None
        PaddingBottom = None
        PaddingLeft = None
        Floating = false
    }
    let largeMetadataDocs =
        [ for index in 0 .. largeMetadataDocCount - 1 ->
            { Doc.empty with Blocks = [ Doc.Table (metadataTable index) ] } ]
    let metadataTrimmed = trimHistory.Invoke(null, [| box largeMetadataDocs |]) :?> Doc.Document list
    check "native history trims table metadata"
        (metadataTrimmed.Length < largeMetadataDocs.Length)
        true
    check "native history metadata bound"
        (metadataTrimmed.Length <= 1)
        true

    let oversizedColumnsWidths = [ for index in 0 .. 200000 -> float index, 12.0 ]
    let oversizedColumns: Doc.SectionColumns =
        { Count = 200001
          Gap = 12.0
          Separator = false
          Widths = oversizedColumnsWidths }
    let oversizedColumnsDoc =
        { Doc.empty with
            PageSetup = { Doc.defaultPageSetup with Columns = Some oversizedColumns } }
    let columnsTrimmed = trimHistory.Invoke(null, [| box [ oversizedColumnsDoc ] |]) :?> Doc.Document list
    check "native history skips section columns"
        columnsTrimmed.Length
        0

    let oversizedShapeSegments =
        [ for index in 0 .. 200000 ->
            Doc.PLine (float index, 0.0) ]
    let oversizedShapeTextBox: Doc.TextBoxBlock =
        { Id = "native-history-shape-geometry"
          Content = []
          X = None
          Y = None
          PositionHRelativeFrom = None
          PositionVRelativeFrom = None
          Width = 1.0
          Height = 1.0
          FillColor = None
          BorderColor = None
          BorderWidth = None
          RelativeHeight = None
          Geometry = Doc.GeoCustom (oversizedShapeSegments, 1.0, 1.0)
          Gradient = None
          Rotation = 0.0 }
    let oversizedShapeDoc =
        { Doc.empty with
            Blocks = [ Doc.TextBox oversizedShapeTextBox ] }
    let shapeTrimmed = trimHistory.Invoke(null, [| box [ oversizedShapeDoc ] |]) :?> Doc.Document list
    check "native history skips shape geometry"
        shapeTrimmed.Length
        0

if isNull documentsEqualForChange then
    fail <- fail + 1
    printfn "FAIL %-34s expected private documentsEqualForChange method" "native change equality reflection"
else
    let largeEqualityParagraphCount = 100000
    let equalityBlocks =
        [ for index in 0 .. largeEqualityParagraphCount - 1 ->
            Doc.Paragraph (para (sprintf "native-eq-%05d" index) "word equality text") ]
    let equalityDoc offset =
        { Doc.empty with
            Blocks = equalityBlocks
            Selection = Some { Anchor = pos "native-eq-00000" offset; Focus = pos "native-eq-00000" offset } }
    let left = equalityDoc 0
    let right = equalityDoc 1
    let sw = System.Diagnostics.Stopwatch.StartNew()
    let mutable equal = true
    for _ in 1 .. 100 do
        equal <- documentsEqualForChange.Invoke(null, [| box left; box right |]) :?> bool
    sw.Stop()
    check "native change equality selection"
        equal
        false
    check "native change equality perf"
        (sw.Elapsed.TotalMilliseconds < 80.0)
        true

if isNull splitBlockPreservingStyles then
    fail <- fail + 1
    printfn "FAIL %-34s expected private splitBlockPreservingStyles method" "native split reflection"
else
    let splitParagraphCount = 5000
    let splitTarget = sprintf "native-split-%05d" (splitParagraphCount - 1)
    let splitStyledBuffer =
        TB.ofString "abcdefgh"
        |> TB.applyStyle 0 4 (fun style -> { style with Bold = true })
    let splitDoc =
        { Doc.empty with
            Blocks =
                [ for index in 0 .. splitParagraphCount - 1 ->
                    let blockId = sprintf "native-split-%05d" index
                    if blockId = splitTarget then
                        Doc.Paragraph { (para blockId "") with Buffer = splitStyledBuffer }
                    else
                        Doc.Paragraph (para blockId "abcdefgh") ]
            Selection = Some { Anchor = pos splitTarget 4; Focus = pos splitTarget 4 } }
    let sw = System.Diagnostics.Stopwatch.StartNew()
    let splitResult = splitBlockPreservingStyles.Invoke(null, [| box splitDoc |]) :?> Doc.Document
    sw.Stop()
    check "native split large doc count" splitResult.Blocks.Length (splitParagraphCount + 1)
    match splitResult.Blocks |> List.rev |> List.truncate 2 |> List.rev with
    | [ Doc.Paragraph beforePara; Doc.Paragraph afterPara ] ->
        check "native split before text" (TB.toString beforePara.Buffer) "abcd"
        check "native split after text" (TB.toString afterPara.Buffer) "efgh"
        check "native split preserves style" (TB.getStyleAt 1 beforePara.Buffer).Bold true
    | tail ->
        fail <- fail + 1
        printfn "FAIL %-34s expected split paragraph tail got %A" "native split shape" tail.Length
    check "native split large doc perf"
        (sw.Elapsed.TotalMilliseconds < 30.0)
        true

    let splitFirstParagraphCount = 50000
    let splitFirstTarget = "native-split-first-00000"
    let splitFirstDoc =
        { Doc.empty with
            Blocks =
                [ for index in 0 .. splitFirstParagraphCount - 1 ->
                    Doc.Paragraph (para (sprintf "native-split-first-%05d" index) "abcdefgh") ]
            Selection = Some { Anchor = pos splitFirstTarget 4; Focus = pos splitFirstTarget 4 } }

    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()

    let splitFirstAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
    let splitFirstSw = System.Diagnostics.Stopwatch.StartNew()
    let splitFirstResult = splitBlockPreservingStyles.Invoke(null, [| box splitFirstDoc |]) :?> Doc.Document
    splitFirstSw.Stop()
    let splitFirstAllocated = System.GC.GetAllocatedBytesForCurrentThread() - splitFirstAllocatedBefore
    let splitFirstTailShared =
        match splitFirstDoc.Blocks, splitFirstResult.Blocks with
        | _ :: originalTail, _ :: _ :: resultTail -> obj.ReferenceEquals(originalTail, resultTail)
        | _ -> false
    check "native split first tail shares"
        splitFirstTailShared
        true
    check "native split first alloc"
        (splitFirstResult.Blocks.Length = splitFirstParagraphCount + 1
         && splitFirstSw.Elapsed.TotalMilliseconds < 30.0
         && splitFirstAllocated < 1024L * 1024L)
        true

let resizeParagraphCount = 50000
let resizeTarget = sprintf "native-resize-%05d" (resizeParagraphCount - 1)
let resizeImageBuffer =
    TB.ofString "ab"
    |> TB.insertImage 1 "data:image/png;base64,AA==" (Some 100.0) (Some 80.0) (Some "img")
let resizeDoc =
    { Doc.empty with
        Blocks =
            [ for index in 0 .. resizeParagraphCount - 1 ->
                let blockId = sprintf "native-resize-%05d" index
                if blockId = resizeTarget then
                    Doc.Paragraph { (para blockId "") with Buffer = resizeImageBuffer }
                else
                    Doc.Paragraph (para blockId "resize text") ] }
let mutable resizedDoc = resizeDoc
for step in 1 .. 3 do
    resizedDoc <- Doc.resizeImage resizeTarget 1 (110.0 + float step) 80.0 resizedDoc

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let resizeAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let resizeSw = System.Diagnostics.Stopwatch.StartNew()
for step in 1 .. 20 do
    resizedDoc <- Doc.resizeImage resizeTarget 1 (200.0 + float step) 80.0 resizedDoc
resizeSw.Stop()
let resizeAllocated = System.GC.GetAllocatedBytesForCurrentThread() - resizeAllocatedBefore
let resizeWidth =
    match Doc.findParagraph resizedDoc resizeTarget with
    | Some para ->
        para.Buffer.Runs
        |> List.tryPick (function TB.Image image when image.Start = 1 -> image.Width | _ -> None)
    | None -> None
check "native image resize applied"
    resizeWidth
    (Some 220.0)
check "native image resize alloc"
    (resizeSw.Elapsed.TotalMilliseconds < 140.0
     && resizeAllocated < 36L * 1024L * 1024L)
    true

let replaceParagraphCount = 50000
let replaceTarget = sprintf "native-replace-%05d" (replaceParagraphCount - 1)
let replaceDoc =
    { Doc.empty with
        Blocks =
            [ for index in 0 .. replaceParagraphCount - 1 ->
                Doc.Paragraph (para (sprintf "native-replace-%05d" index) "replace text") ] }
let mutable replacedDoc = replaceDoc
for step in 1 .. 3 do
    replacedDoc <- Doc.replaceParagraph replacedDoc replaceTarget (para replaceTarget (sprintf "warm replace %d" step))

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let replaceAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let replaceSw = System.Diagnostics.Stopwatch.StartNew()
for step in 1 .. 20 do
    replacedDoc <- Doc.replaceParagraph replacedDoc replaceTarget (para replaceTarget (sprintf "replace %d" step))
replaceSw.Stop()
let replaceAllocated = System.GC.GetAllocatedBytesForCurrentThread() - replaceAllocatedBefore
check "native replace paragraph value"
    (Doc.findParagraph replacedDoc replaceTarget |> Option.map (fun para -> TB.toString para.Buffer))
    (Some "replace 20")
check "native replace paragraph alloc"
    (replaceSw.Elapsed.TotalMilliseconds < 170.0
     && replaceAllocated < 36L * 1024L * 1024L)
    true

let replaceAllParagraphCount = 50000
let replaceAllDoc =
    { Doc.empty with
        Blocks =
            [ for index in 0 .. replaceAllParagraphCount - 1 ->
                let text = if index = 0 then "needle only here" else "ordinary paragraph"
                Doc.Paragraph (para (sprintf "native-replace-all-%05d" index) text) ] }

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let replaceAllAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let replaceAllSw = System.Diagnostics.Stopwatch.StartNew()
let replaceAllResult, replaceAllCount = Doc.replaceAllText "needle" "thread" false replaceAllDoc
replaceAllSw.Stop()
let replaceAllAllocated = System.GC.GetAllocatedBytesForCurrentThread() - replaceAllAllocatedBefore
let replaceAllTailShared =
    match replaceAllDoc.Blocks, replaceAllResult.Blocks with
    | _ :: originalTail, _ :: resultTail -> obj.ReferenceEquals(originalTail, resultTail)
    | _ -> false
check "native replace all count"
    replaceAllCount
    1
check "native replace all tail shares"
    replaceAllTailShared
    true
check "native replace all alloc"
    (replaceAllSw.Elapsed.TotalMilliseconds < 40.0
     && replaceAllAllocated < 4L * 1024L * 1024L)
    true

let findTextHotParagraphCount = 1000
let findTextHotDoc =
    { Doc.empty with
        Blocks =
            [ for index in 0 .. findTextHotParagraphCount - 1 ->
                let text = if index % 100 = 0 then "alpha target omega" else "alpha omega"
                Doc.Paragraph (para (sprintf "native-find-text-hot-%05d" index) text) ] }
let findTextHotFirst = Doc.findTextArray "target" false findTextHotDoc

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let findTextHotProbeCount = 200000
let findTextHotAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let findTextHotSw = System.Diagnostics.Stopwatch.StartNew()
let mutable findTextHotTotal = 0
let mutable findTextHotReused = true
for _ in 1 .. findTextHotProbeCount do
    let matches = Doc.findTextArray "target" false findTextHotDoc
    findTextHotTotal <- findTextHotTotal + matches.Length
    if not (obj.ReferenceEquals(matches, findTextHotFirst)) then
        findTextHotReused <- false
findTextHotSw.Stop()
let findTextHotAllocated = System.GC.GetAllocatedBytesForCurrentThread() - findTextHotAllocatedBefore
check "native find text cache hot hit"
    (findTextHotSw.Elapsed.TotalMilliseconds < 120.0
     && findTextHotAllocated < 1024L * 1024L
     && findTextHotFirst.Length = findTextHotParagraphCount / 100
     && findTextHotTotal = findTextHotFirst.Length * findTextHotProbeCount
     && findTextHotReused)
    true

let findTextEntryCycleQueryCount = 80
let findTextEntryCycleQueries =
    [| for index in 0 .. findTextEntryCycleQueryCount - 1 ->
        sprintf "missing-entry-%02d" index |]
let findTextEntryCycleDoc =
    { Doc.empty with
        Blocks =
            [ for index in 0 .. findTextHotParagraphCount - 1 ->
                Doc.Paragraph (para (sprintf "native-find-entry-%05d" index) "alpha omega text") ] }
for query in findTextEntryCycleQueries do
    Doc.findTextArray query false findTextEntryCycleDoc |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let findTextEntryCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let findTextEntryCycleSw = System.Diagnostics.Stopwatch.StartNew()
let mutable findTextEntryCycleTotal = 0
for _pass in 1 .. 20 do
    for query in findTextEntryCycleQueries do
        findTextEntryCycleTotal <- findTextEntryCycleTotal + (Doc.findTextArray query false findTextEntryCycleDoc).Length
findTextEntryCycleSw.Stop()
let findTextEntryCycleAllocated = System.GC.GetAllocatedBytesForCurrentThread() - findTextEntryCycleAllocatedBefore
check "native find text entry cycle"
    (findTextEntryCycleSw.Elapsed.TotalMilliseconds < 80.0
     && findTextEntryCycleAllocated < 512L * 1024L
     && findTextEntryCycleTotal = 0)
    true

let findTextCycleSourceCount = 80
let findTextCycleParagraphCount = 1000
let findTextCycleQuery = "missing-needle"
let findTextCycleDocs =
    [| for sourceIndex in 0 .. findTextCycleSourceCount - 1 ->
        { Doc.empty with
            Blocks =
                [ for index in 0 .. findTextCycleParagraphCount - 1 ->
                    Doc.Paragraph (para (sprintf "native-find-cycle-%02d-%05d" sourceIndex index) "alpha omega text") ] } |]
for doc in findTextCycleDocs do
    Doc.findTextArray findTextCycleQuery false doc |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let findTextCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let findTextCycleSw = System.Diagnostics.Stopwatch.StartNew()
let mutable findTextCycleTotal = 0
for _pass in 1 .. 20 do
    for doc in findTextCycleDocs do
        findTextCycleTotal <- findTextCycleTotal + (Doc.findTextArray findTextCycleQuery false doc).Length
findTextCycleSw.Stop()
let findTextCycleAllocated = System.GC.GetAllocatedBytesForCurrentThread() - findTextCycleAllocatedBefore
check "native find text source cycle"
    (findTextCycleSw.Elapsed.TotalMilliseconds < 80.0
     && findTextCycleAllocated < 512L * 1024L
     && findTextCycleTotal = 0)
    true

let bookmarkParagraphCount = 50000
let bookmarkInterval = 20
let bookmarkId = "native-bookmark-perf"
let bookmarkPlainBuffer = TB.ofString "ordinary paragraph text"
let bookmarkMarkedBuffer =
    TB.ofString "bookmarked paragraph text"
    |> TB.applyStyle 0 10 (fun style -> { style with BookmarkStartId = Some bookmarkId; BookmarkName = Some bookmarkId })
let bookmarkDoc =
    { Doc.empty with
        Blocks =
            [ for index in 0 .. bookmarkParagraphCount - 1 ->
                let buffer = if index % bookmarkInterval = 0 then bookmarkMarkedBuffer else bookmarkPlainBuffer
                Doc.Paragraph { (para (sprintf "native-bookmark-%05d" index) "") with Buffer = buffer } ] }

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let bookmarkAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let bookmarkSw = System.Diagnostics.Stopwatch.StartNew()
let bookmarkResult = Doc.removeBookmark bookmarkId bookmarkDoc
bookmarkSw.Stop()
let bookmarkAllocated = System.GC.GetAllocatedBytesForCurrentThread() - bookmarkAllocatedBefore
let rec bookmarkTailAt index items =
    if index <= 0 then items
    else
        match items with
        | _ :: tail -> bookmarkTailAt (index - 1) tail
        | [] -> []
let bookmarkLastChangedIndex = ((bookmarkParagraphCount - 1) / bookmarkInterval) * bookmarkInterval
let bookmarkTailShared =
    obj.ReferenceEquals(
        bookmarkTailAt (bookmarkLastChangedIndex + 1) bookmarkDoc.Blocks,
        bookmarkTailAt (bookmarkLastChangedIndex + 1) bookmarkResult.Blocks)
let bookmarkFirstCleared =
    match bookmarkResult.Blocks with
    | Doc.Paragraph para :: _ -> (TB.getStyleAt 0 para.Buffer).BookmarkStartId
    | _ -> Some "missing"
check "native remove bookmark clears"
    bookmarkFirstCleared
    None
check "native remove bookmark tail shares"
    bookmarkTailShared
    true
check "native remove bookmark alloc"
    (bookmarkSw.Elapsed.TotalMilliseconds < 40.0
     && bookmarkAllocated < 13L * 1024L * 1024L)
    true

let blockLookupIdPrefix = System.String('b', 5000)
let blockLookupIds =
    [| for index in 0 .. 511 -> blockLookupIdPrefix + sprintf "%04d" index |]
let blockLookupDoc =
    { Doc.empty with
        Blocks =
            [ for blockId in blockLookupIds ->
                Doc.Paragraph (para blockId "lookup text") ] }
Doc.findBlock blockLookupDoc blockLookupIds.[0] |> ignore
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let blockLookupAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let blockLookupSw = System.Diagnostics.Stopwatch.StartNew()
let mutable blockLookupHits = 0
for _pass in 1 .. 2000 do
    for blockId in blockLookupIds do
        match Doc.findBlock blockLookupDoc blockId with
        | Some _ -> blockLookupHits <- blockLookupHits + 1
        | None -> ()
blockLookupSw.Stop()
let blockLookupAllocated = System.GC.GetAllocatedBytesForCurrentThread() - blockLookupAllocatedBefore
check "native block lookup long id perf"
    (blockLookupSw.Elapsed.TotalMilliseconds < 400.0
     && blockLookupAllocated < 2L * 1024L * 1024L
     && blockLookupHits = 1024000)
    true
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let blockLookupIndexAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let blockLookupIndexSw = System.Diagnostics.Stopwatch.StartNew()
let mutable blockLookupIndexHits = 0
for _pass in 1 .. 2000 do
    for blockId in blockLookupIds do
        match Doc.findBlockIndex blockLookupDoc blockId with
        | Some _ -> blockLookupIndexHits <- blockLookupIndexHits + 1
        | None -> ()
blockLookupIndexSw.Stop()
let blockLookupIndexAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - blockLookupIndexAllocatedBefore
check "native block index lookup long id perf"
    (blockLookupIndexSw.Elapsed.TotalMilliseconds < 400.0
     && blockLookupIndexAllocated < 2L * 1024L * 1024L
     && blockLookupIndexHits = 1024000)
    true
let clonedBlockLookupIds = blockLookupIds |> Array.map cloneString
Doc.findBlock blockLookupDoc clonedBlockLookupIds.[0] |> ignore
Doc.findBlockIndex blockLookupDoc clonedBlockLookupIds.[0] |> ignore
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let clonedBlockLookupAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let clonedBlockLookupSw = System.Diagnostics.Stopwatch.StartNew()
let mutable clonedBlockLookupHits = 0
for _pass in 1 .. 750 do
    for blockId in clonedBlockLookupIds do
        match Doc.findBlock blockLookupDoc blockId with
        | Some _ -> clonedBlockLookupHits <- clonedBlockLookupHits + 1
        | None -> ()
        match Doc.findBlockIndex blockLookupDoc blockId with
        | Some _ -> clonedBlockLookupHits <- clonedBlockLookupHits + 1
        | None -> ()
clonedBlockLookupSw.Stop()
let clonedBlockLookupAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - clonedBlockLookupAllocatedBefore
check "native block lookup cloned long id cache"
    (clonedBlockLookupSw.Elapsed.TotalMilliseconds < 380.0
     && clonedBlockLookupAllocated < 512L * 1024L
     && clonedBlockLookupHits = 768000)
    true

let blockLookupSourceCycleDocs =
    [| for docIndex in 0 .. 64 ->
        let ids = [| for index in 0 .. 499 -> sprintf "native-block-source-cycle-%02d-%04d" docIndex index |]
        { Doc.empty with
            Blocks = [ for blockId in ids -> Doc.Paragraph (para blockId "alpha beta gamma delta") ] },
        ids.[ids.Length - 1] |]

for doc, target in blockLookupSourceCycleDocs do
    Doc.findBlock doc target |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let blockLookupSourceCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let blockLookupSourceCycleSw = System.Diagnostics.Stopwatch.StartNew()
let mutable blockLookupSourceCycleHits = 0
for _pass in 1 .. 100 do
    for doc, target in blockLookupSourceCycleDocs do
        match Doc.findBlock doc target with
        | Some _ -> blockLookupSourceCycleHits <- blockLookupSourceCycleHits + 1
        | None -> ()
blockLookupSourceCycleSw.Stop()
let blockLookupSourceCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - blockLookupSourceCycleAllocatedBefore
check "native block lookup source cycle"
    (blockLookupSourceCycleSw.Elapsed.TotalMilliseconds < 80.0
     && blockLookupSourceCycleAllocated < 512L * 1024L
     && blockLookupSourceCycleHits = blockLookupSourceCycleDocs.Length * 100)
    true
let cursorBoundaryDownDoc =
    { blockLookupDoc with
        Selection =
            Some {
                Anchor = pos blockLookupIds.[blockLookupIds.Length - 1] 0
                Focus = pos blockLookupIds.[blockLookupIds.Length - 1] 0
            } }
Doc.findBlock cursorBoundaryDownDoc blockLookupIds.[blockLookupIds.Length - 1] |> ignore
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let cursorBoundaryDownAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let cursorBoundaryDownSw = System.Diagnostics.Stopwatch.StartNew()
let mutable cursorBoundaryDownSame = 0
for _pass in 1 .. 1024000 do
    if obj.ReferenceEquals(Doc.moveCursorDown cursorBoundaryDownDoc, cursorBoundaryDownDoc) then
        cursorBoundaryDownSame <- cursorBoundaryDownSame + 1
cursorBoundaryDownSw.Stop()
let cursorBoundaryDownAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - cursorBoundaryDownAllocatedBefore
check "native cursor down boundary cache hot hit"
    (cursorBoundaryDownSw.Elapsed.TotalMilliseconds < 500.0
     && cursorBoundaryDownAllocated < 2L * 1024L * 1024L
     && cursorBoundaryDownSame = 1024000)
    true
let cursorBoundaryUpDoc =
    { blockLookupDoc with
        Selection =
            Some {
                Anchor = pos blockLookupIds.[0] 0
                Focus = pos blockLookupIds.[0] 0
            } }
Doc.findBlock cursorBoundaryUpDoc blockLookupIds.[0] |> ignore
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let cursorBoundaryUpAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let cursorBoundaryUpSw = System.Diagnostics.Stopwatch.StartNew()
let mutable cursorBoundaryUpSame = 0
for _pass in 1 .. 1024000 do
    if obj.ReferenceEquals(Doc.moveCursorUp cursorBoundaryUpDoc, cursorBoundaryUpDoc) then
        cursorBoundaryUpSame <- cursorBoundaryUpSame + 1
cursorBoundaryUpSw.Stop()
let cursorBoundaryUpAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - cursorBoundaryUpAllocatedBefore
check "native cursor up boundary cache hot hit"
    (cursorBoundaryUpSw.Elapsed.TotalMilliseconds < 500.0
     && cursorBoundaryUpAllocated < 2L * 1024L * 1024L
     && cursorBoundaryUpSame = 1024000)
    true
let headingNoopPara =
    { para "native-heading-noop" "Title" with
        Buffer =
            TB.applyStyle
                0
                5
                (fun style -> { style with Bold = true; FontSize = 24.0 })
                (TB.ofString "Title") }
let headingNoopDoc =
    { Doc.empty with
        Blocks = [ Doc.Heading (1, headingNoopPara) ]
        Selection = Some { Anchor = pos "native-heading-noop" 0; Focus = pos "native-heading-noop" 0 } }
Doc.findBlock headingNoopDoc "native-heading-noop" |> ignore
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let headingNoopAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let headingNoopSw = System.Diagnostics.Stopwatch.StartNew()
let mutable headingNoopSame = 0
for _pass in 1 .. 200000 do
    if obj.ReferenceEquals(Doc.setHeadingLevel 1 headingNoopDoc, headingNoopDoc) then
        headingNoopSame <- headingNoopSame + 1
headingNoopSw.Stop()
let headingNoopAllocated = System.GC.GetAllocatedBytesForCurrentThread() - headingNoopAllocatedBefore
check "native heading noop cache hot hit"
    (headingNoopSw.Elapsed.TotalMilliseconds < 250.0
     && headingNoopAllocated < 128L * 1024L
     && headingNoopSame = 200000)
    true
let tableCellMarginCell : Doc.TableCell =
    { Content = [ para "native-table-margin-cell" "x" ]
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
let tableCellMarginRow : Doc.TableRow =
    { Cells = [ tableCellMarginCell ]
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
let tableCellMarginTable : Doc.TableBlock =
    { Id = "native-table-margin"
      Rows = [ tableCellMarginRow ]
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
let tableCellMarginDoc =
    { Doc.empty with
        Blocks = [ Doc.Table tableCellMarginTable ]
        Selection =
            Some {
                Anchor = pos "native-table-margin-cell" 0
                Focus = pos "native-table-margin-cell" 0
            } }
Doc.setCurrentTableCellMargins None None None None tableCellMarginDoc |> ignore
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let tableCellMarginAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let tableCellMarginSw = System.Diagnostics.Stopwatch.StartNew()
let mutable tableCellMarginSame = 0
for _pass in 1 .. 200000 do
    if obj.ReferenceEquals(Doc.setCurrentTableCellMargins None None None None tableCellMarginDoc, tableCellMarginDoc) then
        tableCellMarginSame <- tableCellMarginSame + 1
tableCellMarginSw.Stop()
let tableCellMarginAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - tableCellMarginAllocatedBefore
check "native table margin noop cache hot hit"
    (tableCellMarginSw.Elapsed.TotalMilliseconds < 250.0
     && tableCellMarginAllocated < 512L * 1024L
     && tableCellMarginSame = 200000)
    true

let clonedTableCellMarginIdPrefix = System.String('c', 5000)
let clonedTableCellMarginIds =
    [| for index in 0 .. 511 -> clonedTableCellMarginIdPrefix + sprintf "%04d" index |]
let clonedTableCellMarginTable =
    { tableCellMarginTable with
        Id = "native-table-margin-cloned"
        Rows =
            [ for paragraphId in clonedTableCellMarginIds ->
                { tableCellMarginRow with
                    Cells =
                        [ { tableCellMarginCell with
                                Content = [ para paragraphId "x" ] } ] } ] }
let clonedTableCellMarginBlocks = [ Doc.Table clonedTableCellMarginTable ]
let clonedTableCellMarginDocs =
    [| for paragraphId in clonedTableCellMarginIds ->
        { Doc.empty with
            Blocks = clonedTableCellMarginBlocks
            Selection =
                Some {
                    Anchor = pos (cloneString paragraphId) 0
                    Focus = pos (cloneString paragraphId) 0
                } } |]
Doc.setCurrentTableCellMargins
    None
    None
    None
    None
    { clonedTableCellMarginDocs.[0] with
        Selection =
            Some {
                Anchor = pos clonedTableCellMarginIds.[0] 0
                Focus = pos clonedTableCellMarginIds.[0] 0
            } }
|> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let clonedTableCellMarginAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let clonedTableCellMarginSw = System.Diagnostics.Stopwatch.StartNew()
let mutable clonedTableCellMarginSame = 0
for _pass in 1 .. 500 do
    for doc in clonedTableCellMarginDocs do
        if obj.ReferenceEquals(Doc.setCurrentTableCellMargins None None None None doc, doc) then
            clonedTableCellMarginSame <- clonedTableCellMarginSame + 1
clonedTableCellMarginSw.Stop()
let clonedTableCellMarginAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - clonedTableCellMarginAllocatedBefore
check "native table margin cloned id cache"
    (clonedTableCellMarginSw.Elapsed.TotalMilliseconds < 220.0
     && clonedTableCellMarginAllocated < 512L * 1024L
     && clonedTableCellMarginSame = 256000)
    true

let tableCellMarginSourceCycleCount = 80
let tableCellMarginSourceCycleRows = 1000
let tableCellMarginSourceCycleDocs =
    [| for sourceIndex in 0 .. tableCellMarginSourceCycleCount - 1 ->
        let selectedId = sprintf "native-table-margin-source-%02d-%05d" sourceIndex (tableCellMarginSourceCycleRows - 1)
        let table =
            { tableCellMarginTable with
                Id = sprintf "native-table-margin-source-%02d" sourceIndex
                Rows =
                    [ for rowIndex in 0 .. tableCellMarginSourceCycleRows - 1 ->
                        { tableCellMarginRow with
                            Cells =
                                [ { tableCellMarginCell with
                                        Content = [ para (sprintf "native-table-margin-source-%02d-%05d" sourceIndex rowIndex) "x" ] } ] } ] }
        { Doc.empty with
            Blocks = [ Doc.Table table ]
            Selection =
                Some {
                    Anchor = pos selectedId 0
                    Focus = pos selectedId 0
                } } |]

for doc in tableCellMarginSourceCycleDocs do
    Doc.setCurrentTableCellMargins None None None None doc |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let tableCellMarginSourceCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let tableCellMarginSourceCycleSw = System.Diagnostics.Stopwatch.StartNew()
let mutable tableCellMarginSourceCycleSame = 0
for _pass in 1 .. 20 do
    for doc in tableCellMarginSourceCycleDocs do
        if obj.ReferenceEquals(Doc.setCurrentTableCellMargins None None None None doc, doc) then
            tableCellMarginSourceCycleSame <- tableCellMarginSourceCycleSame + 1
tableCellMarginSourceCycleSw.Stop()
let tableCellMarginSourceCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - tableCellMarginSourceCycleAllocatedBefore
check "native table margin source cycle"
    (tableCellMarginSourceCycleSw.Elapsed.TotalMilliseconds < 80.0
     && tableCellMarginSourceCycleAllocated < 512L * 1024L
     && tableCellMarginSourceCycleSame = 1600)
    true

let paragraphLookupIdPrefix = System.String('p', 5000)
let paragraphLookupIds =
    [| for index in 0 .. 511 -> paragraphLookupIdPrefix + sprintf "%04d" index |]
let paragraphLookupDoc =
    { Doc.empty with
        Blocks =
            [ for paragraphId in paragraphLookupIds ->
                Doc.Paragraph (para paragraphId "lookup text") ] }
Doc.findParagraph paragraphLookupDoc paragraphLookupIds.[0] |> ignore
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let paragraphLookupAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let paragraphLookupSw = System.Diagnostics.Stopwatch.StartNew()
let mutable paragraphLookupHits = 0
for _pass in 1 .. 2000 do
    for paragraphId in paragraphLookupIds do
        match Doc.findParagraph paragraphLookupDoc paragraphId with
        | Some _ -> paragraphLookupHits <- paragraphLookupHits + 1
        | None -> ()
paragraphLookupSw.Stop()
let paragraphLookupAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - paragraphLookupAllocatedBefore
check "native paragraph lookup long id perf"
    (paragraphLookupSw.Elapsed.TotalMilliseconds < 500.0
     && paragraphLookupAllocated < 2L * 1024L * 1024L
     && paragraphLookupHits = 1024000)
    true
let clonedParagraphLookupIds = paragraphLookupIds |> Array.map cloneString
Doc.findParagraph paragraphLookupDoc clonedParagraphLookupIds.[0] |> ignore
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let clonedParagraphLookupAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let clonedParagraphLookupSw = System.Diagnostics.Stopwatch.StartNew()
let mutable clonedParagraphLookupHits = 0
for _pass in 1 .. 750 do
    for paragraphId in clonedParagraphLookupIds do
        match Doc.findParagraph paragraphLookupDoc paragraphId with
        | Some _ -> clonedParagraphLookupHits <- clonedParagraphLookupHits + 1
        | None -> ()
clonedParagraphLookupSw.Stop()
let clonedParagraphLookupAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - clonedParagraphLookupAllocatedBefore
check "native paragraph lookup cloned long id cache"
    (clonedParagraphLookupSw.Elapsed.TotalMilliseconds < 340.0
     && clonedParagraphLookupAllocated < 512L * 1024L
     && clonedParagraphLookupHits = 384000)
    true

let paragraphLookupSourceCycleDocs =
    [| for docIndex in 0 .. 64 ->
        let ids = [| for index in 0 .. 499 -> sprintf "native-paragraph-source-cycle-%02d-%04d" docIndex index |]
        { Doc.empty with
            Blocks = [ for paragraphId in ids -> Doc.Paragraph (para paragraphId "alpha beta gamma delta") ] },
        ids.[ids.Length - 1] |]

for doc, target in paragraphLookupSourceCycleDocs do
    Doc.findParagraph doc target |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let paragraphLookupSourceCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let paragraphLookupSourceCycleSw = System.Diagnostics.Stopwatch.StartNew()
let mutable paragraphLookupSourceCycleHits = 0
for _pass in 1 .. 100 do
    for doc, target in paragraphLookupSourceCycleDocs do
        match Doc.findParagraph doc target with
        | Some _ -> paragraphLookupSourceCycleHits <- paragraphLookupSourceCycleHits + 1
        | None -> ()
paragraphLookupSourceCycleSw.Stop()
let paragraphLookupSourceCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - paragraphLookupSourceCycleAllocatedBefore
check "native paragraph lookup source cycle"
    (paragraphLookupSourceCycleSw.Elapsed.TotalMilliseconds < 80.0
     && paragraphLookupSourceCycleAllocated < 512L * 1024L
     && paragraphLookupSourceCycleHits = paragraphLookupSourceCycleDocs.Length * 100)
    true

let normalizedRangeLookupIdPrefix = System.String('r', 5000)
let normalizedRangeLookupIds =
    [| for index in 0 .. 63 -> normalizedRangeLookupIdPrefix + sprintf "%04d" index |]
let normalizedRangeLookupDoc =
    { Doc.empty with
        Blocks =
            [ for blockId in normalizedRangeLookupIds ->
                Doc.Paragraph (para blockId "lookup text") ] }
let normalizedRangeLookupSelections : Doc.Selection array =
    [| for index in 0 .. 63 ->
        { Anchor = pos normalizedRangeLookupIds.[index] 0
          Focus = pos normalizedRangeLookupIds.[index] 3 } |]
for selection in normalizedRangeLookupSelections do
    Doc.getNormalizedRange selection normalizedRangeLookupDoc |> ignore
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let normalizedRangeLookupAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let normalizedRangeLookupSw = System.Diagnostics.Stopwatch.StartNew()
let mutable normalizedRangeLookupTotal = 0
for _pass in 1 .. 10000 do
    for selection in normalizedRangeLookupSelections do
        let startPos, endPos, blockIds = Doc.getNormalizedRange selection normalizedRangeLookupDoc
        normalizedRangeLookupTotal <- normalizedRangeLookupTotal + startPos.Offset + endPos.Offset + blockIds.Length
normalizedRangeLookupSw.Stop()
let normalizedRangeLookupAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - normalizedRangeLookupAllocatedBefore
check "native normalized range long id perf"
    (normalizedRangeLookupSw.Elapsed.TotalMilliseconds < 2600.0
     && normalizedRangeLookupAllocated < 30L * 1024L * 1024L
     && normalizedRangeLookupTotal = 2560000)
    true

let selectionTextLookupIdPrefix = System.String('u', 5000)
let selectionTextLookupIds =
    [| for index in 0 .. 63 -> selectionTextLookupIdPrefix + sprintf "%04d" index |]
let selectionTextLookupBaseDoc =
    { Doc.empty with
        Blocks =
            [ for blockId in selectionTextLookupIds ->
                Doc.Paragraph (para blockId "abcdef") ] }
let selectionTextLookupDocs : Doc.Document array =
    [| for index in 0 .. 63 ->
        { selectionTextLookupBaseDoc with
            Selection =
                Some {
                    Anchor = pos selectionTextLookupIds.[index] 1
                    Focus = pos selectionTextLookupIds.[index] 4
                } } |]
for doc in selectionTextLookupDocs do
    Doc.getSelectionText doc |> ignore
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let selectionTextLookupAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let selectionTextLookupSw = System.Diagnostics.Stopwatch.StartNew()
let mutable selectionTextLookupTotal = 0
for _pass in 1 .. 10000 do
    for doc in selectionTextLookupDocs do
        selectionTextLookupTotal <- selectionTextLookupTotal + (Doc.getSelectionText doc).Length
selectionTextLookupSw.Stop()
let selectionTextLookupAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - selectionTextLookupAllocatedBefore
check "native selection text long id perf"
    (selectionTextLookupSw.Elapsed.TotalMilliseconds < 500.0
     && selectionTextLookupAllocated < 128L * 1024L
     && selectionTextLookupTotal = 1920000)
    true

let selectionTextEntryCycleIdPrefix = System.String('v', 5000)
let selectionTextEntryCycleCount = 80
let selectionTextEntryCycleIds =
    [| for index in 0 .. selectionTextEntryCycleCount - 1 ->
        selectionTextEntryCycleIdPrefix + sprintf "%04d" index |]
let selectionTextEntryCycleBaseDoc =
    { Doc.empty with
        Blocks =
            [ for blockId in selectionTextEntryCycleIds ->
                Doc.Paragraph (para blockId "abcdef") ] }
let selectionTextEntryCycleDocs =
    [| for index in 0 .. selectionTextEntryCycleCount - 1 ->
        { selectionTextEntryCycleBaseDoc with
            Selection =
                Some {
                    Anchor = pos selectionTextEntryCycleIds.[index] 1
                    Focus = pos selectionTextEntryCycleIds.[index] 4
                } } |]
for doc in selectionTextEntryCycleDocs do
    Doc.getSelectionText doc |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let selectionTextEntryCyclePasses = 1000
let selectionTextEntryCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let selectionTextEntryCycleSw = System.Diagnostics.Stopwatch.StartNew()
let mutable selectionTextEntryCycleTotal = 0
for _pass in 1 .. selectionTextEntryCyclePasses do
    for doc in selectionTextEntryCycleDocs do
        selectionTextEntryCycleTotal <- selectionTextEntryCycleTotal + (Doc.getSelectionText doc).Length
selectionTextEntryCycleSw.Stop()
let selectionTextEntryCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - selectionTextEntryCycleAllocatedBefore
check "native selection text entry cycle"
    (selectionTextEntryCycleSw.Elapsed.TotalMilliseconds < 120.0
     && selectionTextEntryCycleAllocated < 4L * 1024L * 1024L
     && selectionTextEntryCycleTotal = 240000)
    true

let selectionTextHotDoc =
    { Doc.empty with
        Blocks = [ Doc.Paragraph (para "native-selection-text-hot" "abcdef") ]
        Selection =
            Some {
                Anchor = pos "native-selection-text-hot" 1
                Focus = pos "native-selection-text-hot" 4
            } }
let selectionTextHotFirst = Doc.getSelectionText selectionTextHotDoc

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let selectionTextHotProbeCount = 200000
let selectionTextHotAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let selectionTextHotSw = System.Diagnostics.Stopwatch.StartNew()
let mutable selectionTextHotTotal = 0
for _ in 1 .. selectionTextHotProbeCount do
    selectionTextHotTotal <- selectionTextHotTotal + (Doc.getSelectionText selectionTextHotDoc).Length
selectionTextHotSw.Stop()
let selectionTextHotAllocated = System.GC.GetAllocatedBytesForCurrentThread() - selectionTextHotAllocatedBefore
check "native selection text cache hot hit"
    (selectionTextHotSw.Elapsed.TotalMilliseconds < 80.0
     && selectionTextHotAllocated < 1024L * 1024L
     && selectionTextHotFirst = "bcd"
     && selectionTextHotTotal = selectionTextHotFirst.Length * selectionTextHotProbeCount)
    true

let clonedSelectionTextBlockId = "native-selection-text-cloned"
let clonedSelectionTextBlocks = [ Doc.Paragraph (para clonedSelectionTextBlockId "abcdef") ]
let clonedSelectionTextDocs =
    [| for _ in 1 .. 200 ->
        { Doc.empty with
            Blocks = clonedSelectionTextBlocks
            Selection =
                Some {
                    Anchor = pos (cloneString clonedSelectionTextBlockId) 1
                    Focus = pos (cloneString clonedSelectionTextBlockId) 4
                } } |]
for doc in clonedSelectionTextDocs do
    Doc.getSelectionText doc |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let clonedSelectionTextAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let clonedSelectionTextSw = System.Diagnostics.Stopwatch.StartNew()
let mutable clonedSelectionTextTotal = 0
for _pass in 1 .. 500 do
    for doc in clonedSelectionTextDocs do
        clonedSelectionTextTotal <- clonedSelectionTextTotal + (Doc.getSelectionText doc).Length
clonedSelectionTextSw.Stop()
let clonedSelectionTextAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - clonedSelectionTextAllocatedBefore
check "native selection text cloned id cache"
    (clonedSelectionTextSw.Elapsed.TotalMilliseconds < 80.0
     && clonedSelectionTextAllocated < 512L * 1024L
     && clonedSelectionTextTotal = 300000)
    true

let selectionTextSourceCycleDocCount = 80
let selectionTextSourceCycleParagraphCount = 1000
let selectionTextSourceCycleDocs =
    [| for docIndex in 0 .. selectionTextSourceCycleDocCount - 1 ->
        let blockId index = sprintf "native-selection-source-cycle-%02d-%04d" docIndex index
        { Doc.empty with
            Blocks =
                [ for index in 0 .. selectionTextSourceCycleParagraphCount - 1 ->
                    Doc.Paragraph (para (blockId index) "abcdef") ]
            Selection =
                Some {
                    Anchor = pos (blockId 1) 1
                    Focus = pos (blockId (selectionTextSourceCycleParagraphCount - 2)) 4
                } } |]
for doc in selectionTextSourceCycleDocs do
    Doc.getSelectionText doc |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let selectionTextSourceCyclePasses = 20
let selectionTextSourceCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let selectionTextSourceCycleSw = System.Diagnostics.Stopwatch.StartNew()
let mutable selectionTextSourceCycleTotal = 0
for _pass in 1 .. selectionTextSourceCyclePasses do
    for doc in selectionTextSourceCycleDocs do
        selectionTextSourceCycleTotal <- selectionTextSourceCycleTotal + (Doc.getSelectionText doc).Length
selectionTextSourceCycleSw.Stop()
let selectionTextSourceCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - selectionTextSourceCycleAllocatedBefore
check "native selection text source cycle"
    (selectionTextSourceCycleSw.Elapsed.TotalMilliseconds < 80.0
     && selectionTextSourceCycleAllocated < 512L * 1024L
     && selectionTextSourceCycleTotal = 11171200)
    true

let sameBlockSelectionTextParagraphCount = 50000
let sameBlockSelectionTextDoc =
    { Doc.empty with
        Blocks =
            [ for index in 0 .. sameBlockSelectionTextParagraphCount - 1 ->
                Doc.Paragraph (para (sprintf "native-selection-text-same-%05d" index) "abcd") ]
        Selection =
            Some {
                Anchor = pos "native-selection-text-same-00000" 1
                Focus = pos "native-selection-text-same-00000" 3
            } }

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let sameBlockSelectionTextAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let sameBlockSelectionTextSw = System.Diagnostics.Stopwatch.StartNew()
let sameBlockSelectionText = Doc.getSelectionText sameBlockSelectionTextDoc
sameBlockSelectionTextSw.Stop()
let sameBlockSelectionTextAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - sameBlockSelectionTextAllocatedBefore
check "native selection text same-block cold"
    (sameBlockSelectionText = "bc"
     && sameBlockSelectionTextSw.Elapsed.TotalMilliseconds < 30.0
     && sameBlockSelectionTextAllocated < 200000L)
    true

let clonedCursorFormatIdPrefix = System.String('f', 5000)
let clonedCursorFormatIds =
    [| for index in 0 .. 511 -> clonedCursorFormatIdPrefix + sprintf "%04d" index |]
let clonedCursorFormatTable =
    { tableCellMarginTable with
        Id = "native-cursor-format-cloned"
        Rows =
            [ for paragraphId in clonedCursorFormatIds ->
                { tableCellMarginRow with
                    Cells =
                        [ { tableCellMarginCell with
                                Content = [ para paragraphId "abcdef" ] } ] } ] }
let clonedCursorFormatBlocks = [ Doc.Table clonedCursorFormatTable ]
let clonedCursorFormatDocs =
    [| for paragraphId in clonedCursorFormatIds ->
        let blockId = cloneString paragraphId
        { Doc.empty with
            Blocks = clonedCursorFormatBlocks
            Selection =
                Some {
                    Anchor = pos blockId 1
                    Focus = pos blockId 4
                } } |]
Doc.setColor
    None
    { clonedCursorFormatDocs.[0] with
        Selection =
            Some {
                Anchor = pos clonedCursorFormatIds.[0] 1
                Focus = pos clonedCursorFormatIds.[0] 4
            } }
|> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let clonedCursorFormatAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let clonedCursorFormatSw = System.Diagnostics.Stopwatch.StartNew()
let mutable clonedCursorFormatSame = 0
for _pass in 1 .. 250 do
    for doc in clonedCursorFormatDocs do
        if obj.ReferenceEquals(Doc.setColor None doc, doc) then
            clonedCursorFormatSame <- clonedCursorFormatSame + 1
clonedCursorFormatSw.Stop()
let clonedCursorFormatAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - clonedCursorFormatAllocatedBefore
check "native cursor paragraph cloned id cache"
    (clonedCursorFormatSw.Elapsed.TotalMilliseconds < 450.0
     && clonedCursorFormatAllocated < 128L * 1024L
     && clonedCursorFormatSame = 128000)
    true

let normalizedRangeClonedId = "native-normalized-range-cloned"
let normalizedRangeClonedDoc =
    { Doc.empty with
        Blocks = [ Doc.Paragraph (para normalizedRangeClonedId "abcdef") ] }
let normalizedRangeClonedSelections : Doc.Selection array =
    [| for _ in 1 .. 2000 ->
        { Anchor = pos (cloneString normalizedRangeClonedId) 1
          Focus = pos (cloneString normalizedRangeClonedId) 4 } |]
for selection in normalizedRangeClonedSelections do
    Doc.getNormalizedRange selection normalizedRangeClonedDoc |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let normalizedRangeClonedAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let normalizedRangeClonedSw = System.Diagnostics.Stopwatch.StartNew()
let mutable normalizedRangeClonedTotal = 0
for _pass in 1 .. 100 do
    for selection in normalizedRangeClonedSelections do
        let startPos, endPos, blockIds = Doc.getNormalizedRange selection normalizedRangeClonedDoc
        normalizedRangeClonedTotal <- normalizedRangeClonedTotal + startPos.Offset + endPos.Offset + blockIds.Length
normalizedRangeClonedSw.Stop()
let normalizedRangeClonedAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - normalizedRangeClonedAllocatedBefore
check "native normalized range cloned id cache"
    (normalizedRangeClonedSw.Elapsed.TotalMilliseconds < 180.0
     && normalizedRangeClonedAllocated < 512L * 1024L
     && normalizedRangeClonedTotal = 1200000)
    true

let normalizedRangeCapacityIds =
    [| for index in 0 .. 511 -> normalizedRangeLookupIdPrefix + "capacity" + sprintf "%04d" index |]
let normalizedRangeCapacityDoc =
    { Doc.empty with
        Blocks =
            [ for blockId in normalizedRangeCapacityIds ->
                Doc.Paragraph (para blockId "lookup text") ] }
let normalizedRangeCapacitySelections : Doc.Selection array =
    [| for index in 0 .. 511 ->
        { Anchor = pos normalizedRangeCapacityIds.[index] 0
          Focus = pos normalizedRangeCapacityIds.[index] 3 } |]
for selection in normalizedRangeCapacitySelections do
    Doc.getNormalizedRange selection normalizedRangeCapacityDoc |> ignore
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let normalizedRangeCapacitySw = System.Diagnostics.Stopwatch.StartNew()
let mutable normalizedRangeCapacityTotal = 0
for _pass in 1 .. 1000 do
    for selection in normalizedRangeCapacitySelections do
        let startPos, endPos, blockIds = Doc.getNormalizedRange selection normalizedRangeCapacityDoc
        normalizedRangeCapacityTotal <- normalizedRangeCapacityTotal + startPos.Offset + endPos.Offset + blockIds.Length
normalizedRangeCapacitySw.Stop()
check "native normalized range capacity perf"
    (normalizedRangeCapacitySw.Elapsed.TotalMilliseconds < 7500.0
     && normalizedRangeCapacityTotal = 2048000)
    true

let normalizedRangeSourceCycleCount = 80
let normalizedRangeSourceCycleParagraphCount = 1000
let normalizedRangeSourceCycleDocs =
    [| for sourceIndex in 0 .. normalizedRangeSourceCycleCount - 1 ->
        let firstId = sprintf "native-range-source-%02d-%05d" sourceIndex 0
        let lastId = sprintf "native-range-source-%02d-%05d" sourceIndex (normalizedRangeSourceCycleParagraphCount - 1)
        let doc =
            { Doc.empty with
                Blocks =
                    [ for index in 0 .. normalizedRangeSourceCycleParagraphCount - 1 ->
                        Doc.Paragraph (para (sprintf "native-range-source-%02d-%05d" sourceIndex index) "range text") ] }
        let selection : Doc.Selection =
            { Anchor = pos firstId 0
              Focus = pos lastId 5 }
        doc, selection |]

for doc, selection in normalizedRangeSourceCycleDocs do
    Doc.getNormalizedRange selection doc |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let normalizedRangeSourceCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let normalizedRangeSourceCycleSw = System.Diagnostics.Stopwatch.StartNew()
let mutable normalizedRangeSourceCycleTotal = 0
for _pass in 1 .. 20 do
    for doc, selection in normalizedRangeSourceCycleDocs do
        let startPos, endPos, blockIds = Doc.getNormalizedRange selection doc
        normalizedRangeSourceCycleTotal <-
            normalizedRangeSourceCycleTotal + startPos.Offset + endPos.Offset + blockIds.Length
normalizedRangeSourceCycleSw.Stop()
let normalizedRangeSourceCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - normalizedRangeSourceCycleAllocatedBefore
check "native normalized range source cycle"
    (normalizedRangeSourceCycleSw.Elapsed.TotalMilliseconds < 80.0
     && normalizedRangeSourceCycleAllocated < 512L * 1024L
     && normalizedRangeSourceCycleTotal = 1608000)
    true

let headerFooterPartSourceCycleCount = 80
let headerFooterPartSourceCyclePartCount = 1000
let headerFooterPartSourceCycleDocs =
    [| for docIndex in 0 .. headerFooterPartSourceCycleCount - 1 ->
        let parts: Doc.HeaderFooterPart list =
            [ for partIndex in 0 .. headerFooterPartSourceCyclePartCount - 1 ->
                { Kind = if partIndex = 0 then "default" else sprintf "kind-%04d" partIndex
                  RelationshipId = Some (sprintf "r-%02d-%04d" docIndex partIndex)
                  Blocks =
                    [ Doc.Paragraph (
                        para
                            (sprintf "native-header-footer-part-%02d-%04d" docIndex partIndex)
                            "header") ] } ]
        { Doc.empty with HeaderParts = parts } |]
for doc in headerFooterPartSourceCycleDocs do
    WordLayout.Page.defaultHeaderParagraphs doc |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let headerFooterPartSourceCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let headerFooterPartSourceCycleSw = System.Diagnostics.Stopwatch.StartNew()
let mutable headerFooterPartSourceCycleTotal = 0
for _pass in 1 .. 20 do
    for doc in headerFooterPartSourceCycleDocs do
        headerFooterPartSourceCycleTotal <-
            headerFooterPartSourceCycleTotal + (WordLayout.Page.defaultHeaderParagraphs doc).Length
headerFooterPartSourceCycleSw.Stop()
let headerFooterPartSourceCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - headerFooterPartSourceCycleAllocatedBefore
check "native header/footer part source cycle"
    (headerFooterPartSourceCycleSw.Elapsed.TotalMilliseconds < 80.0
     && headerFooterPartSourceCycleAllocated < 512L * 1024L
     && headerFooterPartSourceCycleTotal = 1600)
    true

if isNull documentCaptionIdMetrics || isNull documentBookmarkIdMetrics || isNull documentRevisionIdMetrics then
    fail <- fail + 1
    printfn "FAIL %-34s expected private document id metrics helpers" "native id metrics reflection"
else
    let idMetricDoc docIndex =
        { Doc.empty with
            Blocks =
                [ for index in 0 .. 24 ->
                    Doc.Paragraph (para (sprintf "native-id-metric-%03d-%03d" docIndex index) "plain text") ] }
    let idMetricDocs =
        [| for index in 0 .. 63 -> idMetricDoc index |]
    let captionMetricDocs =
        [| for index in 0 .. 15 -> idMetricDoc (100 + index) |]

    for doc in idMetricDocs do
        documentBookmarkIdMetrics.Invoke(null, [| box doc |]) |> ignore
        documentRevisionIdMetrics.Invoke(null, [| box doc |]) |> ignore
    for doc in captionMetricDocs do
        documentCaptionIdMetrics.Invoke(null, [| box "Figure"; box doc |]) |> ignore

    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()

    let idMetricSw = System.Diagnostics.Stopwatch.StartNew()
    let mutable idMetricHits = 0
    for _pass in 1 .. 2000 do
        for doc in idMetricDocs do
            if not (isNull (documentBookmarkIdMetrics.Invoke(null, [| box doc |]))) then
                idMetricHits <- idMetricHits + 1
            if not (isNull (documentRevisionIdMetrics.Invoke(null, [| box doc |]))) then
                idMetricHits <- idMetricHits + 1
    for _pass in 1 .. 1000 do
        for doc in captionMetricDocs do
            if not (isNull (documentCaptionIdMetrics.Invoke(null, [| box "Figure"; box doc |]))) then
                idMetricHits <- idMetricHits + 1
    idMetricSw.Stop()

    check "native id metrics cache many docs"
        (idMetricSw.Elapsed.TotalMilliseconds < 120.0
         && idMetricHits = 272000)
        true

    let idMetricCycleDocs =
        [| for index in 0 .. 79 -> idMetricDoc (200 + index) |]
    for doc in idMetricCycleDocs do
        documentBookmarkIdMetrics.Invoke(null, [| box doc |]) |> ignore
        documentRevisionIdMetrics.Invoke(null, [| box doc |]) |> ignore

    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()

    let idMetricCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
    let idMetricCycleSw = System.Diagnostics.Stopwatch.StartNew()
    let mutable idMetricCycleHits = 0
    for _pass in 1 .. 500 do
        for doc in idMetricCycleDocs do
            if not (isNull (documentBookmarkIdMetrics.Invoke(null, [| box doc |]))) then
                idMetricCycleHits <- idMetricCycleHits + 1
            if not (isNull (documentRevisionIdMetrics.Invoke(null, [| box doc |]))) then
                idMetricCycleHits <- idMetricCycleHits + 1
    idMetricCycleSw.Stop()
    let idMetricCycleAllocated =
        System.GC.GetAllocatedBytesForCurrentThread() - idMetricCycleAllocatedBefore

    check "native id metrics cache document cycle"
        (idMetricCycleSw.Elapsed.TotalMilliseconds < 120.0
         && idMetricCycleAllocated < 12L * 1024L * 1024L
         && idMetricCycleHits = 80000)
        true

    let captionMetricCycleDocs =
        [| for index in 0 .. 79 -> idMetricDoc (300 + index) |]
    for doc in captionMetricCycleDocs do
        documentCaptionIdMetrics.Invoke(null, [| box "Figure"; box doc |]) |> ignore

    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()

    let captionMetricCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
    let captionMetricCycleSw = System.Diagnostics.Stopwatch.StartNew()
    let mutable captionMetricCycleHits = 0
    for _pass in 1 .. 500 do
        for doc in captionMetricCycleDocs do
            if not (isNull (documentCaptionIdMetrics.Invoke(null, [| box "Figure"; box doc |]))) then
                captionMetricCycleHits <- captionMetricCycleHits + 1
    captionMetricCycleSw.Stop()
    let captionMetricCycleAllocated =
        System.GC.GetAllocatedBytesForCurrentThread() - captionMetricCycleAllocatedBefore

    check "native caption metrics cache document cycle"
        (captionMetricCycleSw.Elapsed.TotalMilliseconds < 80.0
         && captionMetricCycleAllocated < 4L * 1024L * 1024L
         && captionMetricCycleHits = 40000)
        true

let dateFieldParagraphCount = 50000
let dateFieldBuffer =
    TB.ofString "old"
    |> TB.applyStyle 0 3 (fun style -> { style with FieldCode = Some " DATE "; FieldResult = Some "old" })
let dateFieldDoc =
    { Doc.empty with
        Blocks =
            [ for index in 0 .. dateFieldParagraphCount - 1 ->
                let buffer = if index = 0 then dateFieldBuffer else TB.ofString "ordinary paragraph text"
                Doc.Paragraph { (para (sprintf "native-date-field-%05d" index) "") with Buffer = buffer } ] }

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let dateFieldAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let dateFieldSw = System.Diagnostics.Stopwatch.StartNew()
let dateFieldResult = Doc.refreshFields dateFieldDoc
dateFieldSw.Stop()
let dateFieldAllocated = System.GC.GetAllocatedBytesForCurrentThread() - dateFieldAllocatedBefore
let dateFieldText =
    match dateFieldResult.Blocks with
    | Doc.Paragraph para :: _ -> TB.toString para.Buffer
    | _ -> "missing"
check "native date field refreshed"
    (dateFieldText <> "old")
    true
check "native date field no page map alloc"
    (dateFieldSw.Elapsed.TotalMilliseconds < 60.0
     && dateFieldAllocated < 6L * 1024L * 1024L)
    true

let fieldPresenceCycleDocs =
    [| for docIndex in 0 .. 64 ->
        { Doc.empty with
            Blocks =
                [ for index in 0 .. 499 ->
                    Doc.Paragraph (
                        para
                            (sprintf "native-field-presence-source-cycle-%02d-%04d" docIndex index)
                            "plain text without fields") ] } |]

for doc in fieldPresenceCycleDocs do
    Doc.refreshFields doc |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let fieldPresenceCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let fieldPresenceCycleSw = System.Diagnostics.Stopwatch.StartNew()
let mutable fieldPresenceCycleIdentityHits = 0
for _pass in 1 .. 100 do
    for doc in fieldPresenceCycleDocs do
        if obj.ReferenceEquals(Doc.refreshFields doc, doc) then
            fieldPresenceCycleIdentityHits <- fieldPresenceCycleIdentityHits + 1
fieldPresenceCycleSw.Stop()
let fieldPresenceCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - fieldPresenceCycleAllocatedBefore
check "native field presence source cycle"
    (fieldPresenceCycleSw.Elapsed.TotalMilliseconds < 80.0
     && fieldPresenceCycleAllocated < 512L * 1024L
     && fieldPresenceCycleIdentityHits = fieldPresenceCycleDocs.Length * 100)
    true

if isNull vbaDocumentText then
    fail <- fail + 1
    printfn "FAIL %-34s expected private documentText method" "native vba text cache reflection"
else
    let vbaTextCacheDocs =
        [| for docIndex in 0 .. 63 ->
            { Doc.empty with
                Blocks =
                    [ for index in 0 .. 999 ->
                        Doc.Paragraph (para (sprintf "native-vba-text-%02d-%04d" docIndex index) "alpha beta gamma delta") ] } |]

    for doc in vbaTextCacheDocs do
        vbaDocumentText.Invoke(null, [| box doc |]) |> ignore

    let vbaTextCacheSw = System.Diagnostics.Stopwatch.StartNew()
    let mutable vbaTextCacheTotal = 0L
    for _pass in 1 .. 2000 do
        for doc in vbaTextCacheDocs do
            vbaTextCacheTotal <-
                vbaTextCacheTotal
                + int64 ((vbaDocumentText.Invoke(null, [| box doc |]) :?> string).Length)
    vbaTextCacheSw.Stop()

    check "native vba text cache many docs"
        (vbaTextCacheSw.Elapsed.TotalMilliseconds < 80.0
         && vbaTextCacheTotal > 0L)
        true

    let vbaTextCycleDocs =
        [| for docIndex in 0 .. 79 ->
            { Doc.empty with
                Blocks =
                    [ for index in 0 .. 999 ->
                        Doc.Paragraph (para (sprintf "native-vba-text-cycle-%02d-%04d" docIndex index) "alpha beta gamma delta") ] } |]

    for doc in vbaTextCycleDocs do
        vbaDocumentText.Invoke(null, [| box doc |]) |> ignore

    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()

    let vbaTextCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
    let vbaTextCycleSw = System.Diagnostics.Stopwatch.StartNew()
    let mutable vbaTextCycleTotal = 0L
    for _pass in 1 .. 200 do
        for doc in vbaTextCycleDocs do
            vbaTextCycleTotal <-
                vbaTextCycleTotal
                + int64 ((vbaDocumentText.Invoke(null, [| box doc |]) :?> string).Length)
    vbaTextCycleSw.Stop()
    let vbaTextCycleAllocated =
        System.GC.GetAllocatedBytesForCurrentThread() - vbaTextCycleAllocatedBefore

    check "native vba text cache document cycle"
        (vbaTextCycleSw.Elapsed.TotalMilliseconds < 80.0
         && vbaTextCycleAllocated < 8L * 1024L * 1024L
         && vbaTextCycleTotal > 0L)
        true

if isNull vbaStatisticsByConstant || isNull vbaDocumentText then
    fail <- fail + 1
    printfn "FAIL %-34s expected private statistics cache helpers" "native vba statistics cache reflection"
else
    let vbaStatisticsCycleDocs =
        [| for docIndex in 0 .. 79 ->
            { Doc.empty with
                Blocks =
                    [ for index in 0 .. 999 ->
                        Doc.Paragraph (para (sprintf "native-vba-statistics-cycle-%02d-%04d" docIndex index) "alpha beta gamma delta epsilon zeta eta theta") ] } |]
    let vbaStatisticsCycleLengths =
        vbaStatisticsCycleDocs
        |> Array.map (fun doc -> (vbaDocumentText.Invoke(null, [| box doc |]) :?> string).Length)

    for index in 0 .. vbaStatisticsCycleDocs.Length - 1 do
        vbaStatisticsByConstant.Invoke(
            null,
            [| box vbaStatisticsCycleDocs.[index]; box 0; box vbaStatisticsCycleLengths.[index]; box 0 |])
        |> ignore

    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()

    let vbaStatisticsCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
    let vbaStatisticsCycleSw = System.Diagnostics.Stopwatch.StartNew()
    let mutable vbaStatisticsCycleHits = 0
    for _pass in 1 .. 200 do
        for index in 0 .. vbaStatisticsCycleDocs.Length - 1 do
            let result =
                vbaStatisticsByConstant.Invoke(
                    null,
                    [| box vbaStatisticsCycleDocs.[index]; box 0; box vbaStatisticsCycleLengths.[index]; box 0 |])
            if not (isNull result) then
                vbaStatisticsCycleHits <- vbaStatisticsCycleHits + 1
    vbaStatisticsCycleSw.Stop()
    let vbaStatisticsCycleAllocated =
        System.GC.GetAllocatedBytesForCurrentThread() - vbaStatisticsCycleAllocatedBefore

    check "native vba statistics document cycle"
        (vbaStatisticsCycleSw.Elapsed.TotalMilliseconds < 80.0
         && vbaStatisticsCycleAllocated < 8L * 1024L * 1024L
         && vbaStatisticsCycleHits = 16000)
        true

    let vbaStatisticsEntryDoc =
        { Doc.empty with
            Blocks =
                [ for index in 0 .. 1499 ->
                    Doc.Paragraph (
                        para
                            (sprintf "native-vba-statistics-entry-%04d" index)
                            "alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi rho sigma tau") ] }
    let vbaStatisticsEntryTextLength = (Doc.getText vbaStatisticsEntryDoc).Length
    let vbaStatisticsEntryRangeCount = 160
    let vbaStatisticsEntryRangeLength = 4096
    let vbaStatisticsEntryMaxStart =
        max 1 (vbaStatisticsEntryTextLength - vbaStatisticsEntryRangeLength - 1)
    let vbaStatisticsEntryRanges =
        [| for index in 0 .. vbaStatisticsEntryRangeCount - 1 ->
            let startOffset = (index * 997) % vbaStatisticsEntryMaxStart
            startOffset, startOffset + vbaStatisticsEntryRangeLength |]

    for startOffset, endOffset in vbaStatisticsEntryRanges do
        vbaStatisticsByConstant.Invoke(
            null,
            [| box vbaStatisticsEntryDoc; box startOffset; box endOffset; box 0 |])
        |> ignore

    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()

    let vbaStatisticsEntryAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
    let vbaStatisticsEntrySw = System.Diagnostics.Stopwatch.StartNew()
    let mutable vbaStatisticsEntryHits = 0
    for _pass in 1 .. 120 do
        for startOffset, endOffset in vbaStatisticsEntryRanges do
            let result =
                vbaStatisticsByConstant.Invoke(
                    null,
                    [| box vbaStatisticsEntryDoc; box startOffset; box endOffset; box 0 |])
            if not (isNull result) then
                vbaStatisticsEntryHits <- vbaStatisticsEntryHits + 1
    vbaStatisticsEntrySw.Stop()
    let vbaStatisticsEntryAllocated =
        System.GC.GetAllocatedBytesForCurrentThread() - vbaStatisticsEntryAllocatedBefore

    check "native vba statistics entry cycle"
        (vbaStatisticsEntrySw.Elapsed.TotalMilliseconds < 80.0
         && vbaStatisticsEntryAllocated < 8L * 1024L * 1024L
         && vbaStatisticsEntryHits = vbaStatisticsEntryRanges.Length * 120)
        true

if isNull vbaLayoutCacheFor then
    fail <- fail + 1
    printfn "FAIL %-34s expected private layoutCacheFor method" "native vba layout cache reflection"
else
    let vbaLayoutCacheDocs =
        [| for docIndex in 0 .. 23 ->
            { Doc.empty with
                Blocks =
                    [ for index in 0 .. 199 ->
                        Doc.Paragraph (para (sprintf "native-vba-layout-%02d-%04d" docIndex index) "alpha beta gamma delta epsilon zeta eta theta") ] } |]

    for doc in vbaLayoutCacheDocs do
        vbaLayoutCacheFor.Invoke(null, [| box doc |]) |> ignore

    let vbaLayoutCacheSw = System.Diagnostics.Stopwatch.StartNew()
    let mutable vbaLayoutCacheHits = 0
    for _pass in 1 .. 5000 do
        for doc in vbaLayoutCacheDocs do
            if not (isNull (vbaLayoutCacheFor.Invoke(null, [| box doc |]))) then
                vbaLayoutCacheHits <- vbaLayoutCacheHits + 1
    vbaLayoutCacheSw.Stop()

    check "native vba layout cache many docs"
        (vbaLayoutCacheSw.Elapsed.TotalMilliseconds < 40.0
         && vbaLayoutCacheHits = 120000)
        true

    let vbaLayoutCycleDocs =
        [| for docIndex in 0 .. 79 ->
            { Doc.empty with
                Blocks =
                    [ for index in 0 .. 199 ->
                        Doc.Paragraph (para (sprintf "native-vba-layout-cycle-%02d-%04d" docIndex index) "alpha beta gamma delta epsilon zeta eta theta") ] } |]

    for doc in vbaLayoutCycleDocs do
        vbaLayoutCacheFor.Invoke(null, [| box doc |]) |> ignore

    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()

    let vbaLayoutCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
    let vbaLayoutCycleSw = System.Diagnostics.Stopwatch.StartNew()
    let mutable vbaLayoutCycleHits = 0
    for _pass in 1 .. 100 do
        for doc in vbaLayoutCycleDocs do
            if not (isNull (vbaLayoutCacheFor.Invoke(null, [| box doc |]))) then
                vbaLayoutCycleHits <- vbaLayoutCycleHits + 1
    vbaLayoutCycleSw.Stop()
    let vbaLayoutCycleAllocated =
        System.GC.GetAllocatedBytesForCurrentThread() - vbaLayoutCycleAllocatedBefore

    check "native vba layout cache document cycle"
        (vbaLayoutCycleSw.Elapsed.TotalMilliseconds < 120.0
         && vbaLayoutCycleAllocated < 8L * 1024L * 1024L
         && vbaLayoutCycleHits = 8000)
        true

if isNull vbaCompileSourceCached then
    fail <- fail + 1
    printfn "FAIL %-34s expected private compileSourceCached method" "native vba compile cache reflection"
else
    let vbaCompileOptions = FsVb.ParseOptions.vba
    let vbaCompileSources =
        [| for sourceIndex in 0 .. 63 ->
            System.String.Join(
                "\n",
                [|
                    yield "Dim x"
                    yield sprintf "' native compile source %02d" sourceIndex
                    for line in 1 .. 20 do
                        yield sprintf "x = %d" (sourceIndex + line)
                |]) |]

    for source in vbaCompileSources do
        vbaCompileSourceCached.Invoke(null, [| box vbaCompileOptions; box source |]) |> ignore

    let vbaCompileCacheSw = System.Diagnostics.Stopwatch.StartNew()
    let mutable vbaCompileCacheHits = 0
    for _pass in 1 .. 5000 do
        for source in vbaCompileSources do
            let result = vbaCompileSourceCached.Invoke(null, [| box vbaCompileOptions; box source |])
            if not (isNull result) then
                vbaCompileCacheHits <- vbaCompileCacheHits + 1
    vbaCompileCacheSw.Stop()

    check "native vba compile cache many sources"
        (vbaCompileCacheSw.Elapsed.TotalMilliseconds < 300.0
         && vbaCompileCacheHits = 320000)
        true

    let vbaCompileCycleSources =
        [| for sourceIndex in 0 .. 79 ->
            System.String.Join(
                "\n",
                [|
                    yield "Dim x"
                    yield sprintf "' native compile source cycle %02d" sourceIndex
                    for line in 1 .. 20 do
                        yield sprintf "x = %d" (sourceIndex + line)
                |]) |]

    for source in vbaCompileCycleSources do
        vbaCompileSourceCached.Invoke(null, [| box vbaCompileOptions; box source |]) |> ignore

    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()

    let vbaCompileCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
    let vbaCompileCycleSw = System.Diagnostics.Stopwatch.StartNew()
    let mutable vbaCompileCycleHits = 0
    for _pass in 1 .. 20 do
        for source in vbaCompileCycleSources do
            let result = vbaCompileSourceCached.Invoke(null, [| box vbaCompileOptions; box source |])
            if not (isNull result) then
                vbaCompileCycleHits <- vbaCompileCycleHits + 1
    vbaCompileCycleSw.Stop()
    let vbaCompileCycleAllocated =
        System.GC.GetAllocatedBytesForCurrentThread() - vbaCompileCycleAllocatedBefore

    check "native vba compile cache source cycle"
        (vbaCompileCycleSw.Elapsed.TotalMilliseconds < 80.0
         && vbaCompileCycleAllocated < 8L * 1024L * 1024L
         && vbaCompileCycleHits = 1600)
        true

if isNull tryLoadWordFont || isNull wordFontCacheProperty then
    fail <- fail + 1
    printfn "FAIL %-34s expected private font cache helpers" "native font cache reflection"
else
    let cache = wordFontCacheProperty.GetValue(null) :?> Dictionary<string, Renderer.Text.FontData option>
    cache.Clear()
    for index in 0 .. 199 do
        let name = sprintf "UnsupportedFont%03d%s" index (System.String('x', 1024))
        tryLoadWordFont.Invoke(null, [| box (Some name); box false; box false |]) |> ignore
    check "native unsupported font cache"
        cache.Count
        0

    for name in [ "Calibri"; " calibri "; "ARIAL"; "Times New Roman"; "Cambria" ] do
        tryLoadWordFont.Invoke(null, [| box (Some name); box false; box false |]) |> ignore
    let supportedKeys = cache.Keys |> Seq.toArray |> Array.sort
    check "native supported font cache count"
        supportedKeys.Length
        4
    check "native font cache normalizes"
        (supportedKeys |> Array.contains "calibri|false|false")
        true

printfn ""
printfn "Word native editor cache tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
