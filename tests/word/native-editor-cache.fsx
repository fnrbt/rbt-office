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
module Paint = RedBat.Word.Paint
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
    check "native viewport command size"
        (viewportCommandsA.Length < 100)
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
         && viewportAllocated < 4L * 1024L * 1024L
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
     && rendererMovingSelectionAllocated < 1024L * 1024L
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
     && resizeAllocated < 70L * 1024L * 1024L)
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
     && replaceAllocated < 70L * 1024L * 1024L)
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
     && bookmarkAllocated < 20L * 1024L * 1024L)
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
let blockLookupSw = System.Diagnostics.Stopwatch.StartNew()
let mutable blockLookupHits = 0
for _pass in 1 .. 2000 do
    for blockId in blockLookupIds do
        match Doc.findBlock blockLookupDoc blockId with
        | Some _ -> blockLookupHits <- blockLookupHits + 1
        | None -> ()
blockLookupSw.Stop()
check "native block lookup long id perf"
    (blockLookupSw.Elapsed.TotalMilliseconds < 400.0
     && blockLookupHits = 1024000)
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
let paragraphLookupSw = System.Diagnostics.Stopwatch.StartNew()
let mutable paragraphLookupHits = 0
for _pass in 1 .. 2000 do
    for paragraphId in paragraphLookupIds do
        match Doc.findParagraph paragraphLookupDoc paragraphId with
        | Some _ -> paragraphLookupHits <- paragraphLookupHits + 1
        | None -> ()
paragraphLookupSw.Stop()
check "native paragraph lookup long id perf"
    (paragraphLookupSw.Elapsed.TotalMilliseconds < 500.0
     && paragraphLookupHits = 1024000)
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
let selectionTextLookupSw = System.Diagnostics.Stopwatch.StartNew()
let mutable selectionTextLookupTotal = 0
for _pass in 1 .. 10000 do
    for doc in selectionTextLookupDocs do
        selectionTextLookupTotal <- selectionTextLookupTotal + (Doc.getSelectionText doc).Length
selectionTextLookupSw.Stop()
check "native selection text long id perf"
    (selectionTextLookupSw.Elapsed.TotalMilliseconds < 1000.0
     && selectionTextLookupTotal = 1920000)
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
