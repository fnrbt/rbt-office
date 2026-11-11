#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.Pdf.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.Unicode.Core.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.Bidi.Core.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.OpenType.Core.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.Shaper.Core.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.Text.Engine.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/ShaperText/bin/Debug/net10.0/Rbt.Word.ShaperText.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.Word.Pdf.dll"

open System
module Doc = Rbt.Word.Document
module Layout = Rbt.Word.Layout
module Paint = Rbt.Word.Paint
module Pdf = Rbt.Word.Pdf

let mutable pass = 0
let mutable fail = 0

let check label condition detail =
    if condition then
        pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-32s %s" label detail

let allocatedBytesFor (f: unit -> 'a) =
    GC.Collect()
    GC.WaitForPendingFinalizers()
    GC.Collect()
    let before = GC.GetAllocatedBytesForCurrentThread()
    let result = f()
    GC.KeepAlive result
    GC.GetAllocatedBytesForCurrentThread() - before

let retainedBytesFor (f: unit -> 'a) =
    GC.Collect()
    GC.WaitForPendingFinalizers()
    GC.Collect()
    let before = GC.GetTotalMemory(true)
    let result = f()
    GC.KeepAlive result
    GC.Collect()
    GC.WaitForPendingFinalizers()
    GC.Collect()
    GC.GetTotalMemory(true) - before

let paragraph index =
    Doc.Paragraph {
        Id = sprintf "pdf-p-%04d" index
        Buffer = Rbt.Word.TextBuffer.ofString "The quick brown fox jumps over the lazy dog."
        Metadata = Doc.defaultMetadata
    }

let repeatedTextDoc =
    { Doc.empty with
        Blocks = [ for i in 1 .. 4000 -> paragraph i ]
        Selection = None
        PageSetup = { Doc.defaultPageSetup with PageHeight = 300000.0 } }

let fontBytes = IO.File.ReadAllBytes("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf")
let options: Pdf.RenderOptions =
    { DefaultFontFamily = "DejaVu Sans"
      Fonts = [
        { FaceName = "DejaVu Sans"
          Family = "DejaVu Sans"
          Bold = false
          Italic = false
          Data = fontBytes }
      ]
      MeasureText = None
      ShapeText = None
      Images = [] }

let warmBytes = Pdf.toBytes options repeatedTextDoc
check "word pdf emits bytes" (warmBytes.Length > 1000) (sprintf "bytes=%d" warmBytes.Length)

// Word.Pdf no longer owns the deleted font-context cache. Retained bytes here include
// process-level caches in sibling text/engine and text/shaper, so keep this as a
// coarse same-document growth sanity check rather than a cache eviction budget.
let repeatedExportCount = 8
let repeatedExportRetainedBound = int64 repeatedExportCount * 2L * 1024L * 1024L
let repeatedExportRetained =
    retainedBytesFor (fun () ->
        let mutable total = 0
        for _ in 1 .. repeatedExportCount do
            total <- total + (Pdf.toBytes options repeatedTextDoc).Length
        total)
check "word pdf repeated export retained sanity"
    (repeatedExportRetained < repeatedExportRetainedBound)
    (sprintf
        "retained %.1f MB bound %.1f MB"
        (float repeatedExportRetained / 1024.0 / 1024.0)
        (float repeatedExportRetainedBound / 1024.0 / 1024.0))

let blankPlan: Pdf.PaintPlan =
    { Settings =
        { Width = 240.0
          PageHeight = 320.0
          FontSize = 14.0
          LineHeight = 21.0
          PaddingLeft = 0.0
          PaddingTop = 0.0
          MarginRight = 0.0
          MarginBottom = 0.0
          InTableCell = false
          TableContinuationHeaderHeight = 0.0
          LongTrailingSpaceFitSafety = 0.25; DefaultTabStop = 48.0 }
      Commands = []
      Pages =
        [ { Top = 0.0
            Width = 240.0
            Height = 320.0
            MarginTop = 0.0
            MarginBottom = 0.0
            MarginLeft = 0.0
            MarginRight = 0.0 } ]
      ImageSources = []
      ImageUses = []
      PageUnit = Pdf.CssPixels }
// The chosen PDF font must be the face that produced the glyph IDs and
// advances, not merely a face with the requested weight in another family.
let serifBoldBytes = IO.File.ReadAllBytes("/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf")
let regularA: Pdf.FontFace =
    { FaceName = "FamilyA"; Family = "FamilyA"; Bold = false; Italic = false; Data = fontBytes }
let styledB: Pdf.FontFace =
    { FaceName = "FamilyB-Bold"; Family = "FamilyB"; Bold = true; Italic = false; Data = serifBoldBytes }
let italicB: Pdf.FontFace =
    { styledB with FaceName = "FamilyB-Italic"; Bold = false; Italic = true }
let paintStyle: Paint.TextPaintStyle =
    { Bold = false; Italic = false; Underline = false; AllCaps = false
      SmallCaps = false; Strike = false; FontSize = 18.0; CharacterSpacing = 0.0
      FontFamily = Some "FamilyA"; Color = "#000000"; Highlight = None; VertAlign = None }
let facePlan style =
    { blankPlan with Commands = [ Paint.DrawText ("WAV fi", style, 12.0, 48.0, 65.0) ] }
let renderFaces fonts shape style =
    Pdf.toBytesWithPaintPlan
        { options with DefaultFontFamily = "FamilyB"; Fonts = fonts; ShapeText = shape }
        (facePlan style)
// Font resource names may differ even when the subset and glyph stream are
// identical. Ignore only those names and the xref offsets they change.
let comparablePdf (bytes: byte[]) =
    let pdf = Text.Encoding.GetEncoding("ISO-8859-1").GetString bytes
    let xref = pdf.LastIndexOf("\nxref", StringComparison.Ordinal)
    let body = if xref >= 0 then pdf.Substring(0, xref) else pdf
    Text.RegularExpressions.Regex.Replace(body, @"(/(?:BaseFont|FontName)\s*)/[^\s<>\[\]()]+", "$1/NormalizedFont")


for label, bold, italic in [ "bold", true, false; "italic", false, true; "bold italic", true, true ] do
    let style = { paintStyle with Bold = bold; Italic = italic }
    let actual = renderFaces [ regularA; styledB; italicB ] None style
    let expected = renderFaces [ regularA ] None style
    let wrongFace = renderFaces [ styledB; italicB ] None style
    check ("shaping face fallback " + label) (actual = expected && actual <> wrongFace)
        "PDF must contain the same subset, glyphs and widths as the requested family's regular face"

// The external shaper may select another registered face, or even one that
// was not registered in Fonts. In both cases its returned font owns the IDs.
let externalFont = Rbt.Text.Engine.loadFont serifBoldBytes
let externalShape text (style: Rbt.Word.TextBuffer.TextStyle) =
    Some (externalFont, Rbt.Text.Engine.shapeText externalFont style.FontSize text)
let externalActual = renderFaces [ regularA; styledB ] (Some externalShape) paintStyle
let externalExpected = renderFaces [ styledB ] (Some externalShape) paintStyle
check "external shaper registered face" (externalActual = externalExpected)
    "external shaping face, not the requested family, must determine PDF glyphs and widths"

let externalUnregistered = renderFaces [ regularA ] (Some externalShape) paintStyle
let externalRegistered = renderFaces [ regularA; styledB ] (Some externalShape) paintStyle
let externalWrong = renderFaces [ regularA ] None paintStyle
check "external shaper unregistered face"
    (comparablePdf externalUnregistered = comparablePdf externalRegistered
     && comparablePdf externalUnregistered <> comparablePdf externalWrong)
    "the external face's own bytes must be embedded when absent from Fonts"

// Duplicate names must not pair a shaped face with the other face's bytes.
let collisionA = { regularA with FaceName = "shared-name" }
let collisionB = { styledB with FaceName = "shared-name" }
let collisionActual = renderFaces [ collisionA; collisionB ] None { paintStyle with Bold = true }
let collisionExpected = renderFaces [ regularA; collisionB ] None { paintStyle with Bold = true }
let collisionWrong = renderFaces [ collisionB ] None { paintStyle with Bold = true }
check "duplicate face name keeps identity"
    (comparablePdf collisionActual = comparablePdf collisionExpected
     && comparablePdf collisionActual <> comparablePdf collisionWrong)
    "a name collision must not substitute the other font's bytes"


let imageResource source dataSize : Pdf.ImageResource =
    { Source = source
      Width = 64
      Height = max 1 (dataSize / (64 * 3))
      ColorSpace = "DeviceRGB"
      BitsPerComponent = 8
      Data = Array.create dataSize 127uy
      Filter = None }

let unusedImageOptions =
    { options with
        Images = [ for index in 1 .. 8 -> imageResource (sprintf "unused-%02d" index) (512 * 1024) ] }

let unusedImageAllocated =
    allocatedBytesFor (fun () -> Pdf.toBytesWithPaintPlan unusedImageOptions blankPlan)
let unusedImageBytes = Pdf.toBytesWithPaintPlan unusedImageOptions blankPlan
check "word pdf skips unused image bytes"
    (unusedImageBytes.Length < 20000)
    (sprintf "bytes=%d" unusedImageBytes.Length)
check "word pdf skips unused image alloc"
    (unusedImageAllocated < 2L * 1024L * 1024L)
    (sprintf "allocated %.1f MB" (float unusedImageAllocated / 1024.0 / 1024.0))

let usedImage = imageResource "used-image" (64 * 64 * 3)
let usedImagePlan =
    { blankPlan with
        Commands = [ Paint.DrawImage ("used-image", 12.0, 12.0, 64.0, 64.0, None, None, None, None) ]
        ImageSources = [ "used-image" ]
        ImageUses = [ ({ Source = "used-image"; Width = 64.0; Height = 64.0 }: Pdf.ImageUse) ] }
let usedImageBytes = Pdf.toBytesWithPaintPlan { options with Images = [ usedImage ] } usedImagePlan
check "word pdf keeps used image"
    (usedImageBytes.Length > unusedImageBytes.Length + usedImage.Data.Length)
    (sprintf "used=%d blank=%d data=%d" usedImageBytes.Length unusedImageBytes.Length usedImage.Data.Length)

let duplicateImagePlan =
    { usedImagePlan with
        Commands = [ Paint.DrawImage ("duplicate-image", 12.0, 12.0, 64.0, 64.0, None, None, None, None) ]
        ImageSources = [ "duplicate-image" ]
        ImageUses = [ ({ Source = "duplicate-image"; Width = 64.0; Height = 64.0 }: Pdf.ImageUse) ] }
let duplicateImageBytes =
    Pdf.toBytesWithPaintPlan
        { options with
            Images = [
                imageResource "duplicate-image" (512 * 1024)
                imageResource "duplicate-image" (64 * 64 * 3)
            ] }
        duplicateImagePlan
check "word pdf dedupes image resources"
    (duplicateImageBytes.Length < unusedImageBytes.Length + 64 * 64 * 3 + 20000)
    (sprintf "bytes=%d blank=%d" duplicateImageBytes.Length unusedImageBytes.Length)

let measuredPlanDoc =
    { Doc.empty with
        Blocks = [ paragraph 90000 ]
        Selection = None }

let probeOptions measureText index =
    let padding = Array.create (8 * 1024 * 1024) (byte (index % 251))
    let data = Array.append fontBytes padding
    let family = sprintf "ProbeFont%d" index
    { options with
        DefaultFontFamily = family
        Fonts = [
            { FaceName = family
              Family = family
              Bold = false
              Italic = false
              Data = data }
        ]
        MeasureText = measureText }

let measuredPlanRetained =
    retainedBytesFor (fun () ->
        let mutable total = 0
        for index in 0 .. 11 do
            let plan =
                Pdf.collectPaintPlan
                    (probeOptions (Some (fun text _ -> float text.Length * 7.0)) index)
                    measuredPlanDoc
            total <- total + plan.Commands.Length + plan.Pages.Length
        total)
check "word pdf measured plan font retention"
    (measuredPlanRetained < 32L * 1024L * 1024L)
    (sprintf "retained %.1f MB" (float measuredPlanRetained / 1024.0 / 1024.0))

let manyPageDoc =
    { Doc.empty with
        Blocks = [ for i in 1 .. 1600 -> paragraph (10000 + i) ]
        Selection = None }

let manyPagePlan = Pdf.collectPaintPlan options manyPageDoc
check "word pdf many page plan" (manyPagePlan.Pages.Length > 20)
    (sprintf "pages=%d" manyPagePlan.Pages.Length)

Pdf.toBytesWithPaintPlan options manyPagePlan |> ignore
let manyPageSw = Diagnostics.Stopwatch.StartNew()
let manyPageBytes = Pdf.toBytesWithPaintPlan options manyPagePlan
manyPageSw.Stop()
check "word pdf many page emits" (manyPageBytes.Length > 1000)
    (sprintf "bytes=%d" manyPageBytes.Length)
check "word pdf many page perf"
    (manyPageSw.Elapsed.TotalMilliseconds < 1500.0)
    (sprintf "many-page export took %.1f ms" manyPageSw.Elapsed.TotalMilliseconds)

// ---------------------------------------------------------------------------
// Page size: the unit comes from the plan (documents are CSS px), not from a
// guess based on the page width.
// ---------------------------------------------------------------------------
let mediaBoxes (bytes: byte[]) =
    let text = Text.Encoding.GetEncoding("ISO-8859-1").GetString bytes
    [ for m in Text.RegularExpressions.Regex.Matches(text, @"/MediaBox\s*\[\s*([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s*\]") ->
        let value (index: int) = Double.Parse(m.Groups.[index].Value, Globalization.CultureInfo.InvariantCulture)
        Math.Round(value 3 - value 1, 1), Math.Round(value 4 - value 2, 1) ]
    |> List.distinct

let pageSizedDoc width height =
    { Doc.empty with
        Blocks = [ paragraph 20000 ]
        PageSetup = { Doc.defaultPageSetup with PageWidth = width; PageHeight = height; MarginLeft = 48.0; MarginRight = 48.0; MarginTop = 48.0; MarginBottom = 48.0 }
        Selection = None }

// A5 is 148 x 210 mm = 559.4 x 793.7 px = 419.5 x 595.3 pt.
let a5Boxes = mediaBoxes (Pdf.toBytes options (pageSizedDoc 559.37 793.7))
// (The PDF writer emits whole points.)
check "word pdf A5 page is A5 in points"
    (match a5Boxes with [ width, height ] -> abs (width - 419.5) <= 1.0 && abs (height - 595.3) <= 1.0 | _ -> false)
    (sprintf "%A" a5Boxes)
let letterBoxes = mediaBoxes (Pdf.toBytes options (pageSizedDoc 816.0 1056.0))
check "word pdf Letter page is Letter in points" (letterBoxes = [ 612.0, 792.0 ]) (sprintf "%A" letterBoxes)
let smallBoxes = mediaBoxes (Pdf.toBytes options (pageSizedDoc 400.0 600.0))
check "word pdf small page is scaled like any other" (smallBoxes = [ 300.0, 450.0 ]) (sprintf "%A" smallBoxes)
// A plan built in points keeps its size, even when wide (landscape Letter).
let pointsPlan =
    { blankPlan with
        Settings = { blankPlan.Settings with Width = 792.0; PageHeight = 612.0 }
        Pages = [ { blankPlan.Pages.Head with Width = 792.0; Height = 612.0 } ]
        PageUnit = Pdf.Points }
let pointsBoxes = mediaBoxes (Pdf.toBytesWithPaintPlan options pointsPlan)
check "word pdf points plan is not shrunk" (pointsBoxes = [ 792.0, 612.0 ]) (sprintf "%A" pointsBoxes)

printfn ""
printfn "Word PDF tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
