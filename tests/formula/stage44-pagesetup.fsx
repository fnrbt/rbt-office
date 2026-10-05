#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
module Ops = Rbt.Spreadsheet.SheetOps

// Page setup as hosts change it: the sections of a header / footer, paper
// names, the checks on each value, manual page breaks, and what rows and
// columns that move do to the print area, the print titles and the breaks.

let mutable pass = 0
let mutable fail = 0
let eq label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-44s expected %A  got %A" label expected got)
let throws label (run: unit -> unit) (part: string) =
    try
        run ()
        fail <- fail + 1
        printfn "FAIL %-44s did not throw" label
    with error ->
        if error.Message.Contains part then pass <- pass + 1
        else (fail <- fail + 1; printfn "FAIL %-44s threw %s" label error.Message)

let sections (raw: string) = let s = Ops.splitHeaderFooter raw in s.Left, s.Center, s.Right
let join left center right = Ops.joinHeaderFooter { Left = left; Center = center; Right = right }

// ---- header / footer sections ----
eq "three sections" (sections "&LLeft&CCentre&RRight") ("Left", "Centre", "Right")
eq "text before a marker is the centre" (sections "Title&RPage &P") ("", "Title", "Page &P")
eq "codes stay in their section" (sections "&L&D &T&C&A&R&P of &N") ("&D &T", "&A", "&P of &N")
eq "a doubled ampersand is no marker" (sections "&LR&&D&CL&&R") ("R&&D", "L&&R", "")
eq "a font name is no marker" (sections "&L&\"Calibri,Regular\"Left&C&\"Courier,Bold\"&12Centre") ("&\"Calibri,Regular\"Left", "&\"Courier,Bold\"&12Centre", "")
eq "markers in lower case" (sections "&lA&cB&rC") ("A", "B", "C")
eq "a section named twice is joined" (sections "&LA&CB&LC") ("AC", "B", "")
eq "an empty string" (sections "") ("", "", "")
eq "an ampersand at the end" (sections "&LA&") ("A&", "", "")
eq "colour codes with letters of markers" (sections "&C&KCC0000Red&RX") ("", "&KCC0000Red", "X")
eq "joined in order" (join "L" "C" "R") (Some "&LL&CC&RR")
eq "empty sections are left out" (join "" "Page &P" "") (Some "&CPage &P")
eq "nothing gives none" (join "" "" "") None
for raw in [ "&LLeft&CCentre&RRight"; "&CPage &P of &N"; "&L&D&R&\"Arial,Bold\"&14Report" ] do
    eq ("split then joined: " + raw) (Ops.joinHeaderFooter (Ops.splitHeaderFooter raw)) (Some raw)

// ---- paper ----
eq "paper names" ([ Some 1; Some 9; Some 8; Some 5; Some 11; Some 13; None; Some 41 ] |> List.map Ops.paperSizeName) [ Some "letter"; Some "a4"; Some "a3"; Some "legal"; Some "a5"; Some "b5"; Some "letter"; None ]
eq "paper codes" ([ "A4"; " letter "; "legal"; "a0" ] |> List.map Ops.paperSizeCode) [ Some 9; Some 1; Some 5; None ]

// ---- changes ----
let apply changes = Ops.applyPageSetupChanges "SetPageSetup" changes defaultPageSetup
let a4 = apply [ Ops.SetPaper 9; Ops.SetOrientation LandscapeOrientation ]
eq "paper: the code and its size in points" (a4.PaperSize, a4.PageWidth, a4.PageHeight, a4.Orientation) (Some 9, 595.27, 841.89, LandscapeOrientation)
let odd = apply [ Ops.SetPaper 41 ]
eq "an unknown paper keeps its code (Letter's size stands in)" (odd.PaperSize, odd.PageWidth) (Some 41, 612.0)
let margins = apply [ Ops.SetMargin ("top", 1.0); Ops.SetMargin ("header", 0.5) ]
eq "margins: the sides named" margins.Margins { defaultPageMargins with Top = 1.0; Header = 0.5 }
let fitted = apply [ Ops.SetFitTo (Some 1, Some 0) ]
eq "fit to pages" (fitted.FitToPage, fitted.FitToWidth, fitted.FitToHeight) (true, 1, 0)
let halfFitted = Ops.applyPageSetupChanges "SetPageSetup" [ Ops.SetFitTo (None, Some 3) ] fitted
eq "one side of the fit keeps the other" (halfFitted.FitToWidth, halfFitted.FitToHeight) (1, 3)
let scaled = Ops.applyPageSetupChanges "SetPageSetup" [ Ops.SetScale 75 ] fitted
eq "a scale ends the fit" (scaled.FitToPage, scaled.Scale) (false, 75)
let numbered = apply [ Ops.SetFirstPageNumber (Some 5) ]
eq "first page number" (numbered.FirstPageNumber, numbered.UseFirstPageNumber) (5, true)
eq "... automatic again" (let p = Ops.applyPageSetupChanges "SetPageSetup" [ Ops.SetFirstPageNumber None ] numbered in p.FirstPageNumber, p.UseFirstPageNumber) (1, false)
let headed = apply [ Ops.SetHeaderFooterText ("header", Some "&CTitle"); Ops.SetHeaderFooterText ("evenFooter", Some "&RPage &P"); Ops.SetDifferentOddEven true ]
eq "headers and footers" (headed.HeaderFooter.OddHeader, headed.HeaderFooter.EvenFooter, headed.HeaderFooter.DifferentOddEven) (Some "&CTitle", Some "&RPage &P", true)
eq "an empty header is none" (Ops.applyPageSetupChanges "SetPageSetup" [ Ops.SetHeaderFooterText ("header", Some "") ] headed).HeaderFooter.OddHeader None
let flags = apply [ Ops.SetPrintGridLines true; Ops.SetPrintHeadings true; Ops.SetCenterHorizontally true; Ops.SetBlackAndWhite true; Ops.SetPageOrder OverThenDown; Ops.SetCommentsPrint CellCommentsAtEnd; Ops.SetScaleWithDocument false ]
eq "print options" (flags.Options, flags.BlackAndWhite, flags.PageOrder, flags.CellComments, flags.HeaderFooter.ScaleWithDoc)
    ({ GridLines = true; Headings = true; HorizontalCentered = true; VerticalCentered = false }, true, OverThenDown, CellCommentsAtEnd, false)
let range r0 c0 r1 c1 : CellRange = { From = { Row = r0; Col = c0 }; To = { Row = r1; Col = c1 } }
let area = apply [ Ops.SetPrintArea (Some (range 0 0 39 5)); Ops.SetRepeatRows (Some (0, 1)); Ops.SetRepeatCols (Some (0, 0)) ]
eq "print area and titles" (area.PrintArea |> Option.map Ops.rangeText, area.RepeatRows, area.RepeatCols) (Some "A1:F40", Some (0, 1), Some (0, 0))
eq "nothing to change gives the same setup" (obj.ReferenceEquals(Ops.applyPageSetupChanges "SetPageSetup" [] area, area)) true

throws "scale below 10" (fun () -> apply [ Ops.SetScale 5 ] |> ignore) "SetPageSetup: scale must be between 10 and 400"
throws "scale above 400" (fun () -> apply [ Ops.SetScale 401 ] |> ignore) "scale must be between"
throws "a negative fit" (fun () -> apply [ Ops.SetFitTo (Some -1, None) ] |> ignore) "fitToWidth must be between 0 and 32767"
throws "a negative margin" (fun () -> apply [ Ops.SetMargin ("left", -0.1) ] |> ignore) "margins.left must be between 0 and 49 inches"
throws "a margin that is no side" (fun () -> apply [ Ops.SetMargin ("inner", 1.0) ] |> ignore) "margins has no side \"inner\""
throws "a paper code of 0" (fun () -> apply [ Ops.SetPaper 0 ] |> ignore) "is not a paper size"
throws "a header that is too long" (fun () -> apply [ Ops.SetHeaderFooterText ("header", Some (String.replicate 256 "x")) ] |> ignore) "header is longer than 255 characters"
throws "a first page number out of range" (fun () -> apply [ Ops.SetFirstPageNumber (Some 40000) ] |> ignore) "firstPageNumber must be between"
throws "titles the wrong way round" (fun () -> apply [ Ops.SetRepeatRows (Some (3, 1)) ] |> ignore) "printTitles.rows is not a run of rows"

// ---- page breaks ----
let broken = Ops.setPageBreaks "SetPageBreaks" [ 40; 20; 20 ] [ 5 ] defaultPageSetup
eq "breaks: sorted, once each" (Ops.pageBreaksOf broken) ([ 20; 40 ], [ 5 ])
eq "the model keeps the row a break comes after, 1-based: the same number" (broken.ManualRowBreaks, broken.ManualColBreaks) ([ 20; 40 ], [ 5 ])
eq "the same breaks give the same setup" (obj.ReferenceEquals(Ops.setPageBreaks "SetPageBreaks" [ 20; 40 ] [ 5 ] broken, broken)) true
throws "no break before the first row" (fun () -> Ops.setPageBreaks "SetPageBreaks" [ 0 ] [] defaultPageSetup |> ignore) "SetPageBreaks: there can be no page break before row 0"
throws "no break outside the sheet" (fun () -> Ops.setPageBreaks "SetPageBreaks" [] [ 16384 ] defaultPageSetup |> ignore) "no page break before column 16384"

// ---- rows and columns that move ----
let sheet = { createEmpty 200 20 with PageSetup = { area with ManualRowBreaks = [ 20; 40 ]; ManualColBreaks = [ 5 ] } }
let inserted = applyStructureEdit (InsertRows (10, 2)) sheet
eq "inserted rows: the area grows, breaks below move, titles stay"
    (inserted.PageSetup.PrintArea |> Option.map Ops.rangeText, inserted.PageSetup.ManualRowBreaks, inserted.PageSetup.RepeatRows) (Some "A1:F42", [ 22; 42 ], Some (0, 1))
let deleted = applyStructureEdit (DeleteCols (0, 1)) sheet
eq "a deleted column: the area shrinks, the break moves, the title column is gone"
    (deleted.PageSetup.PrintArea |> Option.map Ops.rangeText, deleted.PageSetup.ManualColBreaks, deleted.PageSetup.RepeatCols) (Some "A1:E40", [ 4 ], None)

printfn "Page setup tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
