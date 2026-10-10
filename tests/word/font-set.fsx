// ShaperText's font sets: a font file is read once for as long as its bytes
// are around (per array of bytes - not once per set, nor per name a family
// goes by), its outlines when a glyph of it is first drawn, and a set is
// extended with faces (addFaces) without reading the files it has again.
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.Unicode.Core.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.Bidi.Core.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.OpenType.Core.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.Shaper.Core.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.Text.Engine.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/ShaperText/bin/Debug/net10.0/Rbt.Word.ShaperText.dll"

open System
module Shaper = Rbt.Word.ShaperText
module TB = Rbt.Word.TextBuffer

let mutable pass = 0
let mutable fail = 0
let check label got expected =
    if got = expected then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-60s expected %A got %A" label expected got)

let sans = IO.File.ReadAllBytes "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
let bold = IO.File.ReadAllBytes "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
let serif = IO.File.ReadAllBytes "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf"
let face family isBold data : Shaper.FontFace = { Family = family; Bold = isBold; Italic = false; Data = data }
let style family isBold = { TB.defaultStyle with FontFamily = Some family; Bold = isBold; FontSize = 12.0 }
let text = "The quick brown fox jumps over the lazy dog"
let stats () = let s = Shaper.fontSetStats () in s.ShapingParses, s.OutlineParses
let since (shaping, outlines) = let s, o = stats () in s - shaping, o - outlines

// A family under two names, two faces: two files.
let start = stats ()
let faces = [ face "Sans" false sans; face "Sans" true bold; face "Alias" false sans; face "Alias" true bold ]
let set = Shaper.createFontSet "Sans" faces
check "a file is read once, whatever the names it goes by; no outlines yet" (since start) (2, 0)
let again = Shaper.createFontSet "Sans" faces
check "a set of the same files reads nothing" (since start) (2, 0)
let width = Shaper.measureText set text (style "Sans" false)
check "text is measured" (width > 100.0, Shaper.measureText set text (style "Alias" false), Shaper.measureText again text (style "Sans" false)) (true, width, width)
check "bold is another face" (Shaper.measureText set text (style "Sans" true) > width) true

// A face arrives: the set is extended, the files it has are not read again.
let serifFallback = Shaper.measureText set text (style "Serif" false)
check "a family the set has no face of is set in the default family" serifFallback width
let extended = Shaper.addFaces set [ face "Serif" false serif ]
check "adding a face reads its file, and that one only" (since start) (3, 0)
let serifWidth = Shaper.measureText extended text (style "Serif" false)
check "the new set has the face" (serifWidth <> width && serifWidth > 100.0) true
check "the faces it had measure as before" (Shaper.measureText extended text (style "Sans" false), Shaper.measureText extended text (style "Alias" true)) (width, Shaper.measureText set text (style "Sans" true))
check "the set it was made of is not changed" (Shaper.measureText set text (style "Serif" false)) serifFallback
check "adding the faces a set has reads nothing" (Shaper.addFaces extended faces |> ignore; since start) (3, 0)
check "the same set made at once measures the same" (Shaper.measureText (Shaper.createFontSet "Sans" (faces @ [ face "Serif" false serif ])) text (style "Serif" false), since start) (serifWidth, (3, 0))

// The outlines: when a glyph is first drawn, once per file.
let glyph (fontSet: Shaper.FontSet) family isBold =
    match Shaper.shapeText fontSet "A" (style family isBold) with
    | Some (_, shaped) -> (Shaper.glyphPath fontSet (style family isBold) shaped.Runs.Head.Glyphs.[0].GlyphId).Length
    | None -> -1
let path = glyph set "Sans" false
check "a glyph's outline is read from the file" (path > 2, since start) (true, (3, 1))
check "again, under another name, from another set: the same outlines" (glyph set "Sans" false, glyph set "Alias" false, glyph extended "Sans" false, since start) (path, path, path, (3, 1))
check "the units of a face need no outlines" (Shaper.unitsPerEm set (style "Sans" true), since start) (2048, (3, 1))
// Ahead of the first glyph, a file at a time.
check "prepareOutline reads the outlines of a file that has none, and says whether more are to come" (Shaper.prepareOutline extended, since start) (true, (3, 2))
check "the last one" (Shaper.prepareOutline extended, since start) (false, (3, 3))
check "nothing left" (Shaper.prepareOutline extended, Shaper.prepareOutline set, glyph extended "Serif" false > 2, since start) (false, false, true, (3, 3))

// Bytes are known by the array that holds them: a copy is another file.
let copy = Array.copy sans
Shaper.createFontSet "Sans" [ face "Sans" false copy ] |> ignore
check "a copy of the bytes is read as a file of its own" (since start) (4, 3)

printfn "font set: %d passed, %d failed" pass fail
if fail > 0 then exit 1
