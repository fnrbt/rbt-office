#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Workbook
open Rbt.Spreadsheet.Formula
open Rbt.Spreadsheet.ConditionalFormatting

// Conditional formatting and data validation in the model: every rule kind
// evaluated per cell, rules stored by range, priority and stop-if-true,
// structure edits, the OOXML shape round trip, and validation with formulas,
// messages and alert styles.

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-30s  %s" label detail)
let clip (text: string) = if text.Length > 300 then text.Substring(0, 300) + "…" else text
let eq label (got: 'a) (expected: 'a) =
    ok label (got = expected) (clip (sprintf "expected %A  got %A" expected got))

let cell v = { Value = v; Format = defaultFormat; Comment = None }
let num x = cell (CellValue.Number x)
let str x = cell (CellValue.String x)
let fml f = cell (CellValue.Formula (f, CellValue.Empty))
let at r c = { Row = r; Col = c }
let modelOf cells = { createEmpty 100 26 with Cells = Map.ofList cells } |> recalcAndSpill None None
let range r0 c0 r1 c1 : CellRange = { From = at r0 c0; To = at r1 c1 }
let wholeColumn c : CellRange = { From = at 0 c; To = at 1048575 c }
let fill color = { defaultFormat with BgColor = Some color }
let rule ranges r style priority : ConditionalFormatRange =
    { Id = priority; Ranges = ranges; Format = { Rule = r; Style = style }; Priority = priority; StopIfTrue = false }
let withRules rules (m: SpreadsheetModel) = { m with ConditionalFormats = rules }
/// The conditional fill of a cell, if any.
let fillAt r c (m: SpreadsheetModel) =
    effectAt (createContext evaluateFormulaAt m) (at r c) |> Option.bind (fun e -> e.BgColor)
let effectOf r c (m: SpreadsheetModel) = effectAt (createContext evaluateFormulaAt m) (at r c)
let fills c rows m = [ for r in rows -> fillAt r c m |> Option.isSome ]

// A1:A6 = 10, 20, 30, 40, 50, 20   B1:B6 = text   C1:C6 = thresholds
let data =
    modelOf [ at 0 0, num 10.0; at 1 0, num 20.0; at 2 0, num 30.0; at 3 0, num 40.0; at 4 0, num 50.0; at 5 0, num 20.0
              at 0 1, str "Apple"; at 1 1, str "banana"; at 2 1, str "apple"; at 3 1, str "Cherry pie"; at 4 1, str ""; at 5 1, fml "=1/0"
              at 0 2, num 25.0; at 1 2, num 5.0; at 2 2, num 35.0; at 3 2, num 35.0; at 4 2, num 100.0; at 5 2, num 20.0 ]
let colA = [ range 0 0 5 0 ]
let cellIs op a b = data |> withRules [ rule colA (CellIs (op, a, b)) (fill "#FF0000") 1 ]

// --- cellIs with every operator --------------------------------------------
eq "cellis-gt"          (fills 0 [ 0 .. 5 ] (cellIs OpGreaterThan "30" None))        [ false; false; false; true; true; false ]
eq "cellis-lt"          (fills 0 [ 0 .. 5 ] (cellIs OpLessThan "20" None))           [ true; false; false; false; false; false ]
eq "cellis-ge"          (fills 0 [ 0 .. 5 ] (cellIs OpGreaterOrEqual "30" None))     [ false; false; true; true; true; false ]
eq "cellis-le"          (fills 0 [ 0 .. 5 ] (cellIs OpLessOrEqual "20" None))        [ true; true; false; false; false; true ]
eq "cellis-eq"          (fills 0 [ 0 .. 5 ] (cellIs OpEqual "20" None))              [ false; true; false; false; false; true ]
eq "cellis-ne"          (fills 0 [ 0 .. 5 ] (cellIs OpNotEqual "20" None))           [ true; false; true; true; true; false ]
eq "cellis-between"     (fills 0 [ 0 .. 5 ] (cellIs OpBetween "20" (Some "40")))     [ false; true; true; true; false; true ]
eq "cellis-not-between" (fills 0 [ 0 .. 5 ] (cellIs OpNotBetween "40" (Some "20")))  [ true; false; false; false; true; false ]
// Operands are formulas, relative to the range's top-left cell.
eq "cellis-relative"    (fills 0 [ 0 .. 5 ] (cellIs OpGreaterThan "C1" None))        [ false; true; false; true; false; false ]
eq "cellis-absolute"    (fills 0 [ 0 .. 5 ] (cellIs OpGreaterThan "$C$1" None))      [ false; false; true; true; true; false ]
eq "cellis-formula"     (fills 0 [ 0 .. 5 ] (cellIs OpGreaterThan "=AVERAGE($A$1:$A$6)" None)) [ false; false; true; true; true; false ]
let textRule = data |> withRules [ rule [ range 0 1 5 1 ] (CellIs (OpEqual, "\"APPLE\"", None)) (fill "#00FF00") 1 ]
eq "cellis-text"        (fills 1 [ 0 .. 3 ] textRule)                                [ true; false; true; false ]
// A blank cell counts as 0 for a "cell value" rule, as in Excel.
eq "cellis-blank-is-zero" (fillAt 9 0 (data |> withRules [ rule [ range 0 0 9 0 ] (CellIs (OpLessThan, "5", None)) (fill "#FF0000") 1 ])) (Some "#FF0000")

// --- expression rules -------------------------------------------------------
let expr f ranges = data |> withRules [ rule ranges (Expression f) (fill "#0000FF") 1 ]
eq "expr-relative"   (fills 0 [ 0 .. 5 ] (expr "=A1>C1" colA))             [ false; true; false; true; false; false ]
eq "expr-row-band"   (fills 0 [ 0 .. 5 ] (expr "=MOD(ROW(),2)=0" colA))    [ false; true; false; true; false; true ]
eq "expr-mixed-lock" (fills 1 [ 0 .. 5 ] (expr "=$A1>=30" [ range 0 0 5 1 ])) [ false; false; true; true; true; false ]
eq "expr-whole-row"  (fillAt 3 2 (expr "=$A1>=30" [ range 0 0 5 2 ]))      (Some "#0000FF")
eq "expr-countif"    (fills 0 [ 0 .. 5 ] (expr "=COUNTIF($A$1:$A$6,A1)>1" colA)) [ false; true; false; false; false; true ]
eq "expr-error-no-match" (fillAt 0 0 (expr "=1/0" colA)) None
eq "expr-second-range-anchors-first" (fillAt 2 0 (expr "=A1>C1" [ range 0 0 0 0; range 2 0 3 0 ])) None

// --- text rules -------------------------------------------------------------
let text kind value = data |> withRules [ rule [ range 0 1 5 1 ] (TextRule (kind, value)) (fill "#AAAAAA") 1 ]
eq "text-contains"     (fills 1 [ 0 .. 4 ] (text TextContains "APP"))    [ true; false; true; false; false ]
eq "text-not-contains" (fills 1 [ 0 .. 3 ] (text TextNotContains "an"))  [ true; false; true; true ]
eq "text-begins"       (fills 1 [ 0 .. 3 ] (text TextBeginsWith "ch"))   [ false; false; false; true ]
eq "text-ends"         (fills 1 [ 0 .. 3 ] (text TextEndsWith "PIE"))    [ false; false; false; true ]
eq "blanks"            (fills 1 [ 3; 4 ] (data |> withRules [ rule [ range 0 1 5 1 ] (Blanks true) (fill "#111111") 1 ])) [ false; true ]
eq "errors"            (fills 1 [ 4; 5 ] (data |> withRules [ rule [ range 0 1 5 1 ] (Errors true) (fill "#111111") 1 ])) [ false; true ]
eq "legacy-contains"   (fillAt 0 1 (data |> withRules [ rule [ range 0 1 5 1 ] (ContainsText "pp") (fill "#222222") 1 ])) (Some "#222222")
eq "legacy-on-formula" (fillAt 0 3 (modelOf [ at 0 3, fml "=40+2" ] |> withRules [ rule [ range 0 3 0 3 ] (GreaterThan 41.0) (fill "#222222") 1 ])) (Some "#222222")

// --- top / bottom, averages, duplicates ------------------------------------
let ranked r = data |> withRules [ rule colA r (fill "#FFFF00") 1 ]
eq "top-2"          (fills 0 [ 0 .. 5 ] (ranked (TopBottom (2, false, false))))  [ false; false; false; true; true; false ]
eq "bottom-2"       (fills 0 [ 0 .. 5 ] (ranked (TopBottom (2, false, true))))   [ true; true; false; false; false; true ]
eq "top-50-percent" (fills 0 [ 0 .. 5 ] (ranked (TopBottom (50, true, false))))  [ false; false; true; true; true; false ]
eq "bottom-20-pct"  (fills 0 [ 0 .. 5 ] (ranked (TopBottom (20, true, true))))   [ true; false; false; false; false; false ]
// mean = 28.33
eq "above-average"  (fills 0 [ 0 .. 5 ] (ranked (AboveAverage (true, false, 0))))  [ false; false; true; true; true; false ]
eq "below-average"  (fills 0 [ 0 .. 5 ] (ranked (AboveAverage (false, false, 0)))) [ true; true; false; false; false; true ]
eq "above-1-stddev" (fills 0 [ 0 .. 5 ] (ranked (AboveAverage (true, false, 1))))  [ false; false; false; false; true; false ]
let flat = modelOf [ at 0 0, num 5.0; at 1 0, num 5.0 ]
eq "equal-average"  (fillAt 0 0 (flat |> withRules [ rule colA (AboveAverage (true, true, 0)) (fill "#FFFF00") 1 ]) |> Option.isSome,
                     fillAt 0 0 (flat |> withRules [ rule colA (AboveAverage (true, false, 0)) (fill "#FFFF00") 1 ]) |> Option.isSome) (true, false)
eq "duplicates"     (fills 0 [ 0 .. 5 ] (ranked DuplicateValues))                 [ false; true; false; false; false; true ]
eq "uniques"        (fills 0 [ 0 .. 5 ] (ranked UniqueValues))                    [ true; false; true; true; true; false ]
eq "duplicate-text-ignores-case" (fills 1 [ 0 .. 3 ] (data |> withRules [ rule [ range 0 1 5 1 ] DuplicateValues (fill "#FFFF00") 1 ])) [ true; false; true; false ]

// --- colour scales, data bars, icon sets -----------------------------------
let cfvo kind value : CfThreshold = { Kind = kind; Value = value }
let lo, hi = cfvo CfvoMin None, cfvo CfvoMax None
let scale2 = data |> withRules [ rule colA (ColorScale [ lo, "#FF0000"; hi, "#00FF00" ]) defaultFormat 1 ]
eq "scale2-min"  (fillAt 0 0 scale2) (Some "#FF0000")
eq "scale2-max"  (fillAt 4 0 scale2) (Some "#00FF00")
eq "scale2-mid"  (fillAt 2 0 scale2) (Some "#808000")
let scale3 = data |> withRules [ rule colA (ColorScale [ lo, "#FF0000"; cfvo CfvoPercentile (Some "50"), "#FFFFFF"; hi, "#0000FF" ]) defaultFormat 1 ]
eq "scale3-low"       (fillAt 0 0 scale3) (Some "#FF0000")
eq "scale3-median"    (fillAt 5 0 scale3 = fillAt 1 0 scale3, fillAt 4 0 scale3) (true, Some "#0000FF")
eq "scale3-at-median" (fillAt 0 3 (modelOf [ at 0 3, num 0.0; at 1 3, num 5.0; at 2 3, num 10.0 ]
                                   |> withRules [ rule [ range 0 3 2 3 ] (ColorScale [ lo, "#FF0000"; cfvo CfvoPercent (Some "50"), "#FFFFFF"; hi, "#0000FF" ]) defaultFormat 1 ]
                                   |> fun m -> { m with ConditionalFormats = m.ConditionalFormats }) ) (Some "#FF0000")
let scaleNum = data |> withRules [ rule colA (ColorScale [ cfvo CfvoNumber (Some "20"), "#000000"; cfvo CfvoNumber (Some "40"), "#FFFFFF" ]) defaultFormat 1 ]
eq "scale-number-stops" (fillAt 0 0 scaleNum, fillAt 2 0 scaleNum, fillAt 4 0 scaleNum) (Some "#000000", Some "#808080", Some "#FFFFFF")
eq "scale-skips-text"   (fillAt 0 1 (data |> withRules [ rule [ range 0 0 5 1 ] (ColorScale [ lo, "#FF0000"; hi, "#00FF00" ]) defaultFormat 1 ])) None

let bars = data |> withRules [ rule colA (DataBar (lo, hi, "#638EC6")) defaultFormat 1 ]
eq "databar-max"  (effectOf 4 0 bars |> Option.bind (fun e -> e.DataBar)) (Some (plainDataBar 1.0 "#638EC6"))
eq "databar-part" (effectOf 0 0 bars |> Option.bind (fun e -> e.DataBar)) (Some (plainDataBar 0.2 "#638EC6"))
let barsNum = data |> withRules [ rule colA (DataBar (cfvo CfvoNumber (Some "10"), cfvo CfvoNumber (Some "30"), "638EC6")) defaultFormat 1 ]
eq "databar-bounds" ([ 0; 1; 4 ] |> List.map (fun r -> effectOf r 0 barsNum |> Option.bind (fun e -> e.DataBar) |> Option.map (fun bar -> bar.To))) [ Some 0.0; Some 0.5; Some 1.0 ]
let icons = data |> withRules [ rule colA (IconSet ("3TrafficLights1", [ cfvo CfvoPercent (Some "0"); cfvo CfvoPercent (Some "33"); cfvo CfvoPercent (Some "67") ], false)) defaultFormat 1 ]
eq "icons"          ([ 0; 2; 4 ] |> List.map (fun r -> effectOf r 0 icons |> Option.bind (fun e -> e.IconSet) |> Option.map (fun icon -> icon.Index, icon.Count))) [ Some (0, 3); Some (1, 3); Some (2, 3) ]
let iconsRev = data |> withRules [ rule colA (IconSet ("5Arrows", [ for p in [ "0"; "20"; "40"; "60"; "80" ] -> cfvo CfvoPercent (Some p) ], true)) defaultFormat 1 ]
eq "icons-5-reverse" ([ 0; 4 ] |> List.map (fun r -> effectOf r 0 iconsRev |> Option.bind (fun e -> e.IconSet) |> Option.map (fun icon -> icon.Index, icon.Count))) [ Some (4, 5); Some (0, 5) ]
// The painter's format carries the effect.
eq "format-databar" ((formatAt bars (at 4 0) defaultFormat).DataBar) (Some (plainDataBar 1.0 "#638EC6"))
eq "format-fill"    ((formatAt scale2 (at 0 0) defaultFormat).BgColor) (Some "#FF0000")
eq "format-untouched" (formatAt scale2 (at 0 1) defaultFormat) defaultFormat

// The portable painter draws them: fill, bar and icon rectangles.
let viewport : Rbt.Spreadsheet.Layout.Viewport =
    { ScrollX = 0.0; ScrollY = 0.0; Width = 600.0; Height = 300.0; HeaderHeight = 20.0; RowHeaderWidth = 30.0; Measurer = None }
let fillColors (m: SpreadsheetModel) =
    Rbt.Spreadsheet.Paint.buildPaintItems m viewport
    |> List.choose (fun item -> match item with Rbt.Spreadsheet.Paint.FillRect (rect, color) -> Some (color.ToUpperInvariant(), rect.Width) | _ -> None)
eq "paint-plain-has-none"  (fillColors data |> List.exists (fun (c, _) -> c = "#FF0000" || c = "#638EC6")) false
eq "paint-color-scale"     (fillColors scale2 |> List.filter (fun (c, _) -> c = "#FF0000" || c = "#00FF00") |> List.length) 2
eq "paint-cellis-fill"     (fillColors (cellIs OpGreaterThan "30" None) |> List.filter (fun (c, _) -> c = "#FF0000") |> List.length) 2
let barWidths = fillColors bars |> List.filter (fun (c, _) -> c = "#638EC6") |> List.map snd |> List.sort
eq "paint-data-bars"       (barWidths.Length, barWidths.Head < barWidths.[barWidths.Length - 1]) (6, true)
eq "paint-icons"           (fillColors icons |> List.filter (fun (_, w) -> w = 9.0) |> List.length) 6

// --- priority, merging, stop-if-true ---------------------------------------
let red = rule colA (CellIs (OpGreaterThan, "15", None)) { defaultFormat with BgColor = Some "#FF0000"; Bold = true } 2
let blue = rule colA (CellIs (OpGreaterThan, "25", None)) { defaultFormat with BgColor = Some "#0000FF"; TextColor = Some "#FFFFFF" } 1
let layered = data |> withRules [ red; blue ]
eq "priority-wins"     (fillAt 2 0 layered) (Some "#0000FF")
eq "priority-merges"   (effectOf 2 0 layered |> Option.map (fun e -> e.Bold, e.TextColor)) (Some (Some true, Some "#FFFFFF"))
eq "priority-fallback" (fillAt 1 0 layered) (Some "#FF0000")
let stopped = data |> withRules [ red; { blue with StopIfTrue = true } ]
eq "stop-if-true"      (effectOf 2 0 stopped |> Option.map (fun e -> e.BgColor, e.Bold)) (Some (Some "#0000FF", None))
eq "stop-only-on-match" (fillAt 1 0 stopped) (Some "#FF0000")
eq "add-newest-first"  (data |> addConditionalFormat (at 0 0) { Rule = GreaterThan 0.0; Style = fill "#111111" }
                             |> addConditionalFormat (at 0 0) { Rule = GreaterThan 5.0; Style = fill "#222222" } |> fillAt 0 0) (Some "#222222")
eq "legacy-getter"     (getConditionalFormat (at 0 0) (CellValue.Number 10.0) (data |> addConditionalFormat (at 0 0) { Rule = GreaterThan 0.0; Style = fill "#111111" })
                        |> Option.bind (fun f -> f.BgColor)) (Some "#111111")
let removed = data |> withRules [ rule [ range 0 0 5 1 ] (Blanks false) (fill "#333333") 1 ] |> removeConditionalFormats (at 2 0)
eq "remove-one-cell"   (fillAt 2 0 removed, fillAt 1 0 removed, fillAt 2 1 removed, fillAt 3 0 removed) (None, Some "#333333", Some "#333333", Some "#333333")

// --- rules are kept by range: whole columns are not clipped -----------------
let wide = data |> withRules [ rule [ wholeColumn 0 ] (CellIs (OpGreaterThan, "30", None)) (fill "#FF0000") 1 ]
let grown = wide |> setCell (at 5000 0) (num 99.0)
eq "whole-column-far-cell" (fillAt 5000 0 grown) (Some "#FF0000")
eq "whole-column-one-rule" (List.length grown.ConditionalFormats, grown.ConditionalFormats.Head.Ranges) (1, [ wholeColumn 0 ])
eq "whole-column-stats"    (fillAt 5000 0 (grown |> withRules [ rule [ wholeColumn 0 ] (TopBottom (1, false, false)) (fill "#FF0000") 1 ])) (Some "#FF0000")

// --- structure edits move the ranges and the formulas ----------------------
let editable = data |> withRules [ rule [ range 1 0 4 0 ] (Expression "=A2>$C$1") (fill "#FF0000") 1
                                   rule [ wholeColumn 0 ] (CellIs (OpEqual, "C3", None)) (fill "#00FF00") 2 ]
let inserted = insertRow 2 editable
eq "edit-range-grows"   inserted.ConditionalFormats.[0].Ranges [ range 1 0 5 0 ]
eq "edit-formula-moves" inserted.ConditionalFormats.[0].Format.Rule (Expression "=A2>$C$1")
eq "edit-whole-col-kept" inserted.ConditionalFormats.[1].Ranges [ wholeColumn 0 ]
eq "edit-operand-moves" inserted.ConditionalFormats.[1].Format.Rule (CellIs (OpEqual, "C4", None))
let above = insertRow 0 editable
eq "edit-range-shifts"  (above.ConditionalFormats.[0].Ranges, above.ConditionalFormats.[0].Format.Rule) ([ range 2 0 5 0 ], Expression "=A3>$C$2")
eq "edit-still-matches" (fillAt 2 0 above, fillAt 3 0 above) (None, Some "#FF0000")
let gone = Rbt.Spreadsheet.Model.applyStructureEdit (DeleteRows (1, 4)) editable
eq "edit-rule-deleted"  (gone.ConditionalFormats |> List.map (fun r -> r.Ranges)) [ [ wholeColumn 0 ] ]
let colMoved = insertColumn 0 editable
eq "edit-col-shift"     (colMoved.ConditionalFormats.[0].Ranges, colMoved.ConditionalFormats.[1].Ranges) ([ range 1 1 4 1 ], [ wholeColumn 1 ])

// --- OOXML shape -----------------------------------------------------------
eq "sqref-parse"   (parseSqref "B2:C3 $A:$A 3:4 D5") [ range 1 1 2 2; wholeColumn 0; { From = at 2 0; To = at 3 16383 }; range 4 3 4 3 ]
eq "sqref-print"   (sqrefText [ range 1 1 2 2; wholeColumn 0; { From = at 2 0; To = at 3 16383 }; range 4 3 4 3 ]) "B2:C3 A:A 3:4 D5"
let everyRule =
    [ rule colA (CellIs (OpNotBetween, "1", Some "C1*2")) { defaultFormat with BgColor = Some "#FF0000"; TextColor = Some "#FFFFFF"; Bold = true } 1
      rule [ wholeColumn 1; range 0 3 4 4 ] (Expression "$A1>3") { defaultFormat with Italic = true } 2
      rule colA (TopBottom (3, true, true)) (fill "#00FF00") 3
      rule colA (AboveAverage (false, true, 2)) (fill "#00FF00") 4
      rule colA DuplicateValues (fill "#00FF00") 5
      rule colA UniqueValues (fill "#00FF00") 6
      rule colA (TextRule (TextBeginsWith, "ab")) (fill "#00FF00") 7
      rule colA (Blanks true) (fill "#00FF00") 8
      rule colA (ColorScale [ lo, "#F8696B"; cfvo CfvoPercentile (Some "50"), "#FFEB84"; hi, "#63BE7B" ]) defaultFormat 9
      rule colA (DataBar (lo, cfvo CfvoNumber (Some "100"), "#638EC6")) defaultFormat 10
      { rule colA (IconSet ("4Arrows", [ for p in [ "0"; "25"; "50"; "75" ] -> cfvo CfvoPercent (Some p) ], true)) defaultFormat 11 with StopIfTrue = true } ]
let viaOoxml = everyRule |> toOoxml |> ofOoxml |> List.sortBy (fun r -> r.Priority)
for (i, (a, b)) in List.zip viaOoxml everyRule |> List.indexed do
    eq (sprintf "ooxml-round-trip-%d" (i + 1)) (a.Ranges, a.Format.Rule, a.Priority, a.StopIfTrue, a.Format.Style = b.Format.Style)
                                               (b.Ranges, b.Format.Rule, b.Priority, b.StopIfTrue, true)
eq "ooxml-round-trip-count" viaOoxml.Length everyRule.Length
let ooxml = toOoxml everyRule
eq "ooxml-blocks"     (ooxml |> List.map (fun f -> f.Sqref, f.Rules.Length)) [ "A1:A6", 10; "B:B D1:E5", 1 ]
eq "ooxml-types"      (ooxml.Head.Rules |> List.map (fun r -> r.Type))
                      [ "cellIs"; "top10"; "aboveAverage"; "duplicateValues"; "uniqueValues"; "beginsWith"; "containsBlanks"; "colorScale"; "dataBar"; "iconSet" ]
eq "ooxml-cellis"     (ooxml.Head.Rules.Head.Operator, ooxml.Head.Rules.Head.Formulas) (Some "notBetween", [ "1"; "C1*2" ])
eq "ooxml-visual-no-style" (ooxml.Head.Rules |> List.rev |> List.head |> fun r -> r.Style, r.IconSetReverse, r.StopIfTrue) (None, true, true)
eq "ooxml-unknown-skipped" (ofOoxml [ { Sqref = "A1"; Rules = [ emptyOoxmlRule "timePeriod" ] } ]) []
eq "ooxml-legacy-export"   (toOoxml [ rule colA (Between (1.0, 2.5)) (fill "#FF0000") 1 ] |> List.head |> fun f -> f.Rules.Head.Operator, f.Rules.Head.Formulas)
                           (Some "between", [ "1"; "2.5" ])

// --- validation -------------------------------------------------------------
let outcome addr value m = checkEntry addr value m
let invalid o = match o with EntryInvalid _ -> true | EntryValid -> false
let V = CellValue.Number
let vmodel =
    modelOf [ at 0 0, num 5.0; at 1 0, num 5.0; at 0 2, num 10.0; at 1 2, num 20.0 ]
    |> setValidation (at 0 1) (validationOf (Comparison (WholeKind, OpNotBetween, "3", Some "7")))
    |> setValidation (at 1 1) (validationOf (Comparison (DecimalKind, OpNotEqual, "0", None)))
    |> setValidation (at 2 1) (validationOf (Comparison (TextLengthKind, OpLessOrEqual, "3", None)))
    |> setValidation (at 3 1) (validationOf (Comparison (TimeKind, OpBetween, "0.25", Some "0.75")))
    |> setValidation (at 4 1) (validationOf (Comparison (DecimalKind, OpGreaterThan, "=C1", None)))
    |> setValidation (at 5 1) { validationOf (CustomFormula "=ISNUMBER(B5)") with Anchor = Some (at 4 1) }
    |> setValidation (at 6 1) { validationOf (CustomFormula "=COUNTIF($A$1:$A$9,B7)=0") with ErrorTitle = Some "Taken"; ErrorMessage = Some "Already used" }
    |> setValidation (at 7 1) { validationOf (WholeNumber (Some 1.0, Some 10.0)) with Style = WarningStyle; ErrorMessage = Some "Unusual" }
    |> setValidation (at 8 1) { validationOf (WholeNumber (Some 1.0, Some 10.0)) with Style = InformationStyle }
    |> setValidation (at 9 1) { validationOf (WholeNumber (Some 1.0, Some 10.0)) with ShowError = false }
    |> setValidation (at 10 1) { validationOf (Comparison (DateKind, OpGreaterOrEqual, "=DATE(2024,1,1)", None)) with InputTitle = Some "When"; InputMessage = Some "A date in 2024 or later" }
eq "v-not-between"     ([ 2.0; 3.0; 5.0; 7.0; 8.0 ] |> List.map (fun x -> invalid (outcome (at 0 1) (V x) vmodel))) [ false; true; true; true; false ]
eq "v-whole-decimal"   (invalid (outcome (at 0 1) (V 8.5) vmodel)) true
eq "v-not-equal"       (invalid (outcome (at 1 1) (V 0.0) vmodel), invalid (outcome (at 1 1) (V 0.1) vmodel)) (true, false)
eq "v-text-length"     (invalid (outcome (at 2 1) (CellValue.String "abcd") vmodel), invalid (outcome (at 2 1) (CellValue.String "abc") vmodel)) (true, false)
eq "v-time"            (invalid (outcome (at 3 1) (V 0.5) vmodel), invalid (outcome (at 3 1) (V 0.9) vmodel)) (false, true)
eq "v-formula-operand" (invalid (outcome (at 4 1) (V 10.0) vmodel), invalid (outcome (at 4 1) (V 10.5) vmodel)) (true, false)
eq "v-custom-anchor"   (invalid (outcome (at 5 1) (V 1.0) vmodel), invalid (outcome (at 5 1) (CellValue.String "x") vmodel)) (false, true)
eq "v-custom-sees-entry" (invalid (outcome (at 6 1) (V 5.0) vmodel), invalid (outcome (at 6 1) (V 6.0) vmodel)) (true, false)
eq "v-custom-message"  (outcome (at 6 1) (V 5.0) vmodel) (EntryInvalid (StopStyle, Some "Taken", "Already used"))
eq "v-formula-entry"   (invalid (outcome (at 0 1) (CellValue.Formula ("=2+3", CellValue.Empty)) vmodel), invalid (outcome (at 0 1) (CellValue.Formula ("=4*2", CellValue.Empty)) vmodel)) (true, false)
eq "v-blank-allowed"   (outcome (at 0 1) CellValue.Empty vmodel) EntryValid
eq "v-warning-outcome" (outcome (at 7 1) (V 50.0) vmodel) (EntryInvalid (WarningStyle, None, "Unusual"))
eq "v-info-outcome"    (match outcome (at 8 1) (V 50.0) vmodel with EntryInvalid (InformationStyle, _, _) -> true | _ -> false) true
// validateEdit refuses only what Excel's "stop" refuses.
let refused addr x = match validateEdit addr (V x) vmodel with Result.Error _ -> true | Result.Ok () -> false
eq "v-stop-refuses"    (refused (at 0 1) 5.0, refused (at 7 1) 50.0, refused (at 8 1) 50.0, refused (at 9 1) 50.0) (true, false, false, false)
eq "v-no-alert-valid"  (outcome (at 9 1) (V 50.0) vmodel) EntryValid
eq "v-date-formula"    (invalid (outcome (at 10 1) (V 45000.0) vmodel), invalid (outcome (at 10 1) (V 45292.0) vmodel)) (true, false)
eq "v-input-message"   ((getValidation (at 10 1) vmodel).Value.InputTitle, (getValidation (at 10 1) vmodel).Value.InputMessage) (Some "When", Some "A date in 2024 or later")
eq "v-literal-only"    (validateValue (CustomFormula "=FALSE") (V 1.0), (match validateValue (Comparison (WholeKind, OpNotEqual, "4", None)) (V 4.0) with Result.Error _ -> true | _ -> false))
                       (Result.Ok (), true)

// Row/column edits carry the rule's formulas and anchor along.
let vEdited = insertRow 0 vmodel
eq "v-edit-moves-cell"    ((getValidation (at 6 1) vEdited).Value.Rule, (getValidation (at 6 1) vEdited).Value.Anchor)
                          (CustomFormula "=ISNUMBER(B6)", Some (at 5 1))
eq "v-edit-operand"       (getValidation (at 5 1) vEdited).Value.Rule (Comparison (DecimalKind, OpGreaterThan, "=C2", None))
eq "v-edit-absolute"      (getValidation (at 7 1) vEdited).Value.Rule (CustomFormula "=COUNTIF($A$2:$A$10,B8)=0")
eq "v-edit-still-works"   (invalid (checkEntry (at 5 1) (V 10.0) vEdited), invalid (checkEntry (at 5 1) (V 10.5) vEdited)) (true, false)

// OOXML shape of validations.
let ov kind op f1 f2 : OoxmlValidation =
    { Type = kind; Operator = op; Formula1 = f1; Formula2 = f2; AllowBlank = true; ShowErrorMessage = true; ShowInputMessage = true; ShowDropDown = true
      ErrorStyle = Some "warning"; ErrorTitle = Some "T"; Error = Some "E"; PromptTitle = Some "PT"; Prompt = Some "P" }
eq "vo-not-between" (validationRuleOfOoxml (ov "whole" (Some "notBetween") (Some "1") (Some "5"))) (Some (Comparison (WholeKind, OpNotBetween, "1", Some "5")))
eq "vo-between"     (validationRuleOfOoxml (ov "decimal" None (Some "1") (Some "5"))) (Some (DecimalRange (Some 1.0, Some 5.0)))
eq "vo-custom"      (validationRuleOfOoxml (ov "custom" None (Some "A1>0") None)) (Some (CustomFormula "A1>0"))
eq "vo-time"        (validationRuleOfOoxml (ov "time" (Some "greaterThan") (Some "0.5") None)) (Some (Comparison (TimeKind, OpGreaterThan, "0.5", None)))
eq "vo-list"        (validationRuleOfOoxml (ov "list" None (Some "\"Yes, No\"") None)) (Some (ListValues [ "Yes"; "No" ]))
eq "vo-list-range"  (validationRuleOfOoxml (ov "list" None (Some "$A$1:$A$3") None)) (Some (ListSource "$A$1:$A$3"))
eq "vo-any"         (validationRuleOfOoxml (ov "none" None None None)) (Some AnyValue)
let dvFull = validationWithOoxml (at 2 3) (ov "textLength" (Some "notEqual") (Some "4") None) (Comparison (TextLengthKind, OpNotEqual, "4", None))
eq "vo-alert-fields" (dvFull.Style, dvFull.ErrorTitle, dvFull.ErrorMessage, dvFull.InputTitle, dvFull.InputMessage, dvFull.Anchor)
                     (WarningStyle, Some "T", Some "E", Some "PT", Some "P", Some (at 2 3))
eq "vo-round-trip"   (ooxmlOfValidation dvFull) (ov "textLength" (Some "notEqual") (Some "4") None)
let grouped =
    [ { Id = 1; Ranges = [ { From = at 1 1; To = at 3 2 }; { From = at 0 0; To = at 1048575 0 } ]; Validation = dvFull }
      { Id = 2; Ranges = [ { From = at 9 9; To = at 9 9 } ]; Validation = validationOf (CustomFormula "=TRUE") } ]
    |> validationsToOoxml |> List.map fst |> List.sort
eq "vo-ranges-sqref" grouped [ "B2:C4 A1:A1048576"; "J10" ]

printfn ""
printfn "Stage 26 rule tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
