#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
open Rbt.Spreadsheet.ConditionalFormatting
module Ops = Rbt.Spreadsheet.SheetOps

// The conditional formats of Excel 2010+ (the x14 extension) in the model:
// data bars with negative values, an axis, a border, a gradient or solid fill
// and a direction; icon sets with custom icons, hidden values and the icon
// sets 3Stars / 3Triangles / 5Boxes; what a cell draws for them; the rules as
// plain data and back; the file shape.

let mutable pass = 0
let mutable fail = 0
let eq label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-40s expected %A  got %A" label expected got)
let throws label (run: unit -> unit) (start: string) =
    let message = try run (); "(nothing thrown)" with e -> e.Message
    if message.StartsWith start then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-40s expected an error starting \"%s\"  got \"%s\"" label start message)

let at row col : CellAddress = { Row = row; Col = col }
let block r0 c0 r1 c1 : CellRange = { From = at r0 c0; To = at r1 c1 }
let cfvo kind value : CfThreshold = { Kind = kind; Value = value }
let num x = { Value = CellValue.Number x; Format = defaultFormat; Comment = None }
let rule ranges r priority : ConditionalFormatRange =
    { Id = priority; Ranges = ranges; Format = { Rule = r; Style = defaultFormat }; Priority = priority; StopIfTrue = false }
/// A1:A5 = -50, -25, 0, 50, 100.
let data = { createEmpty 50 10 with Cells = Map.ofList [ for i, v in List.indexed [ -50.0; -25.0; 0.0; 50.0; 100.0 ] -> at i 0, num v ] }
let barOf (r: ConditionalFormatRule) row =
    effectAt (createContext evaluateFormulaAt { data with ConditionalFormats = [ rule [ block 0 0 4 0 ] r 1 ] }) (at row 0) |> Option.bind (fun e -> e.DataBar)
let iconOf (m: SpreadsheetModel) (r: ConditionalFormatRule) row =
    effectAt (createContext evaluateFormulaAt { m with ConditionalFormats = [ rule [ block 0 0 4 0 ] r 1 ] }) (at row 0) |> Option.bind (fun e -> e.IconSet)
let round3 (x: float) = System.Math.Round(x, 3)
let span (bar: DataBarDraw option) = bar |> Option.map (fun b -> round3 b.From, round3 b.To)
let auto = cfvo CfvoAutoMin None, cfvo CfvoAutoMax None

// ---- the default data bar of Excel 2010+: automatic ends, automatic axis -------
let standard = DataBarEx (fst auto, snd auto, "#638EC6", defaultDataBarOptions)
// The axis is where zero is between -50 and 100: a third of the way across.
eq "axis: the highest value fills right of the axis" (span (barOf standard 4)) (Some (0.333, 1.0))
eq "axis: half of the highest" (span (barOf standard 3)) (Some (0.333, 0.667))
eq "axis: the lowest value fills left of the axis" (span (barOf standard 0)) (Some (0.0, 0.333))
eq "axis: half of the lowest" (span (barOf standard 1)) (Some (0.167, 0.333))
eq "axis: zero has the axis and no bar" (barOf standard 2 |> Option.map (fun b -> round3 b.From, round3 b.To, b.Axis |> Option.map (fun (x, c) -> round3 x, c))) (Some (0.0, 0.0, Some (0.333, "#000000")))
eq "negative values have their colour" ([ 0; 4 ] |> List.map (fun r -> barOf standard r |> Option.map (fun b -> b.Color))) [ Some "#FF0000"; Some "#638EC6" ]
eq "a gradient fades away from the axis" ([ 0; 4 ] |> List.map (fun r -> barOf standard r |> Option.map (fun b -> b.Gradient, b.FadeRight))) [ Some (true, false); Some (true, true) ]
eq "every cell has the axis" ([ 0 .. 4 ] |> List.map (fun r -> barOf standard r |> Option.bind (fun b -> b.Axis) |> Option.isSome)) [ true; true; true; true; true ]

// ---- options -------------------------------------------------------------------
let withOptions options = DataBarEx (fst auto, snd auto, "#638EC6", options)
let middle = withOptions { defaultDataBarOptions with Axis = AxisMiddle }
eq "axis in the middle: the largest magnitude fills a half" ([ 0; 1; 3; 4 ] |> List.map (fun r -> span (barOf middle r))) [ Some (0.25, 0.5); Some (0.375, 0.5); Some (0.5, 0.75); Some (0.5, 1.0) ]
let none = withOptions { defaultDataBarOptions with Axis = AxisNone }
eq "no axis: every bar from the left edge" ([ 0; 2; 4 ] |> List.map (fun r -> barOf none r |> Option.map (fun b -> round3 b.From, round3 b.To, b.Axis))) [ None; Some (0.0, 0.333, None); Some (0.0, 1.0, None) ]
eq "no axis: a negative value keeps its colour" (barOf none 1 |> Option.map (fun b -> b.Color)) (Some "#FF0000")
let lengths = withOptions { defaultDataBarOptions with Axis = AxisNone; MinLength = 10; MaxLength = 90 }
eq "shortest and longest bar" ([ 0; 4 ] |> List.map (fun r -> span (barOf lengths r))) [ Some (0.0, 0.1); Some (0.0, 0.9) ]
let rightToLeft = withOptions { defaultDataBarOptions with Direction = BarRightToLeft }
eq "right to left: mirrored" ([ 0; 4 ] |> List.map (fun r -> barOf rightToLeft r |> Option.map (fun b -> round3 b.From, round3 b.To, b.FadeRight, b.Axis |> Option.map (fst >> round3)))) [ Some (0.667, 1.0, true, Some 0.667); Some (0.0, 0.667, false, Some 0.667) ]
let bordered = withOptions { defaultDataBarOptions with Gradient = false; Border = Some "#1F4E79"; NegativeBorder = Some "#990000"; NegativeFill = None; ShowValue = false }
eq "solid, bordered, the same fill for negative values, no value" ([ 0; 4 ] |> List.map (fun r -> barOf bordered r |> Option.map (fun b -> b.Gradient, b.Color, b.Border, b.HideValue))) [ Some (false, "#638EC6", Some "#990000", true); Some (false, "#638EC6", Some "#1F4E79", true) ]
// All values positive: no axis, bars from the left, automatic minimum is zero.
let positive = { data with Cells = Map.ofList [ for i, v in List.indexed [ 10.0; 20.0; 30.0; 40.0; 50.0 ] -> at i 0, num v ] }
let positiveBar row = effectAt (createContext evaluateFormulaAt { positive with ConditionalFormats = [ rule [ block 0 0 4 0 ] standard 1 ] }) (at row 0) |> Option.bind (fun e -> e.DataBar)
eq "all positive: from the edge, measured from zero" ([ 0; 4 ] |> List.map (fun r -> positiveBar r |> Option.map (fun b -> round3 b.From, round3 b.To, b.Axis))) [ Some (0.0, 0.2, None); Some (0.0, 1.0, None) ]
// The plain data bar is what it was.
eq "a plain data bar" (barOf (DataBar (cfvo CfvoMin None, cfvo CfvoMax None, "#638EC6")) 4) (Some { plainDataBar 1.0 "#638EC6" with Gradient = true })

// ---- icon sets -------------------------------------------------------------------
let steps = [ cfvo CfvoPercent (Some "0"); cfvo CfvoPercent (Some "33"); cfvo CfvoPercent (Some "67") ]
let stars = IconSetEx ("3Stars", steps, false, defaultIconSetOptions)
eq "3Stars: lowest and highest" ([ 0; 4 ] |> List.map (fun r -> iconOf data stars r)) [ Some { Family = "3Stars"; Index = 0; Count = 3; HideValue = false }; Some { Family = "3Stars"; Index = 2; Count = 3; HideValue = false } ]
let custom = IconSetEx ("3TrafficLights1", steps, false, { ShowValue = false; Custom = [ "3Arrows", 2; "NoIcons", 0; "5Rating", 4 ] })
eq "custom icons: by step, none for NoIcons, values hidden" ([ 0; 2; 4 ] |> List.map (fun r -> iconOf data custom r)) [ Some { Family = "3Arrows"; Index = 2; Count = 3; HideValue = true }; None; Some { Family = "5Rating"; Index = 4; Count = 5; HideValue = true } ]
eq "icon glyphs: arrows point down, sideways, up" ([ 0; 1; 2 ] |> List.map (fun i -> fst (iconGlyph "3Arrows" i))) [ "↓"; "→"; "↑" ]
eq "icon glyphs: the lowest is red, the highest green" (snd (iconGlyph "3TrafficLights1" 0), snd (iconGlyph "3TrafficLights1" 2)) ("#E53935", "#43A047")

// ---- as plain data and back ------------------------------------------------------
for r in [ standard; middle; bordered; rightToLeft; stars; custom
           DataBarEx (cfvo CfvoNumber (Some "-10"), cfvo CfvoFormula (Some "$C$1*2"), "#63C384", { defaultDataBarOptions with AxisColor = None; Axis = AxisNone }) ] do
    let format = { Rule = r; Style = defaultFormat }
    eq (sprintf "there and back: %A" r) (Ops.ruleOfDescription "T" defaultFormat (Ops.describeRule format)) format
let described = Ops.describeRule { Rule = standard; Style = defaultFormat }
eq "a described data bar: its ends are automatic" (described.Low |> Option.map (fun t -> t.Kind), described.High |> Option.map (fun t -> t.Kind), described.Bar) (Some "autoMin", Some "autoMax", Some defaultDataBarOptions)
eq "a plain data bar has no options" (Ops.describeRule { Rule = DataBar (cfvo CfvoMin None, cfvo CfvoMax None, "#638EC6"); Style = defaultFormat }).Bar None
let d kind = Ops.emptyRuleDescription kind
throws "automatic ends are for the data bar of Excel 2010+" (fun () -> Ops.ruleOfDescription "Add" defaultFormat { d "dataBar" with Low = Some { Kind = "autoMin"; Value = None } } |> ignore) "Add: rule.min.type must be one of"
throws "a bad negative colour" (fun () -> Ops.ruleOfDescription "Add" defaultFormat { d "dataBar" with Bar = Some { defaultDataBarOptions with NegativeFill = Some "red" } } |> ignore) "Add: rule.negativeColor must be a colour"
throws "bar lengths" (fun () -> Ops.ruleOfDescription "Add" defaultFormat { d "dataBar" with Bar = Some { defaultDataBarOptions with MinLength = 80; MaxLength = 20 } } |> ignore) "Add: rule.minLength and rule.maxLength"
throws "as many custom icons as steps" (fun () -> Ops.ruleOfDescription "Add" defaultFormat { d "iconSet" with Icons = Some { ShowValue = true; Custom = [ "3Arrows", 0 ] } } |> ignore) "Add: rule.icons needs 3 icons"
throws "a custom icon of its set" (fun () -> Ops.ruleOfDescription "Add" defaultFormat { d "iconSet" with Icons = Some { ShowValue = true; Custom = [ "3Arrows", 0; "3Arrows", 3; "NoIcons", 0 ] } } |> ignore) "Add: rule.icons[1].index must be 0..2"
eq "an icon set of Excel 2010+ by name" (Ops.ruleOfDescription "Add" defaultFormat { d "iconSet" with IconSet = Some "5Boxes" }).Rule
    (IconSetEx ("5Boxes", [ for p in [ "0"; "20"; "40"; "60"; "80" ] -> cfvo CfvoPercent (Some p) ], false, defaultIconSetOptions))

// ---- the file shape ---------------------------------------------------------------
let entries = [ rule [ block 0 0 4 0 ] standard 1; rule [ block 0 0 4 0 ] bordered 2; rule [ block 0 1 4 1 ] stars 3; rule [ block 0 1 4 1 ] custom 4
                { rule [ block 0 2 4 2 ] (Expression "Other!$A$1>A1") 5 with Format = { Rule = Expression "Other!$A$1>A1"; Style = { defaultFormat with BgColor = Some "#FFC7CE" } } }
                rule [ block 0 2 4 2 ] (Expression "$A1>0") 6 ]
let ooxml = toOoxml entries
let flat = ooxml |> List.collect (fun cf -> cf.Rules)
eq "file shape: which rules are in the x14 list only" (flat |> List.map (fun r -> r.Type, r.Ext |> Option.map (fun e -> e.InExtList))) [ "dataBar", Some false; "dataBar", Some false; "iconSet", Some true; "iconSet", Some true; "expression", Some true; "expression", None ]
eq "file shape: the 2007 twin of a data bar has min / max ends" (flat.[0].DataBarCfvo, flat.[0].Ext |> Option.bind (fun e -> e.DataBar) |> Option.map (fun b -> b.Cfvo)) ([ "min", None; "max", None ], Some [ "autoMin", None; "autoMax", None ])
eq "file shape: and back" (ofOoxml ooxml |> List.map (fun e -> e.Ranges, e.Format, e.Priority)) (entries |> List.map (fun e -> e.Ranges, e.Format, e.Priority))

// ---- the portable painter -----------------------------------------------------------
module Paint = Rbt.Spreadsheet.Paint
module Layout = Rbt.Spreadsheet.Layout
let painted (r: ConditionalFormatRule) =
    Paint.buildPaintItems { data with ConditionalFormats = [ rule [ block 0 0 4 0 ] r 1 ] } (Layout.defaultViewport 800.0 500.0)
let fills color items = items |> List.choose (function Paint.FillRect (rect, c) when c = color -> Some rect | _ -> None)
let solid = withOptions { defaultDataBarOptions with Gradient = false; Border = Some "#1F4E79" }
let solidItems = painted solid
// Two negative values and two positive ones have a bar; zero has none.
eq "painter: solid bars in the two colours" ((fills "#FF0000" solidItems).Length, (fills "#638EC6" solidItems).Length) (2, 2)
let paintedNegative = (fills "#FF0000" solidItems) |> List.minBy (fun rect -> rect.Y)
let paintedPositive = (fills "#638EC6" solidItems) |> List.maxBy (fun rect -> rect.Y)
eq "painter: the negative bar ends where the positive one starts (the axis)" (round3 (paintedNegative.X + paintedNegative.Width)) (round3 paintedPositive.X)
eq "painter: the highest value's bar is twice the lowest value's" (round3 (paintedPositive.Width / paintedNegative.Width)) 2.0
eq "painter: a border around each bar" (solidItems |> List.choose (function Paint.StrokeRect (_, c, _) when c = "#1F4E79" -> Some () | _ -> None) |> List.length) 4
let axisPieces = fills "#000000" solidItems |> List.filter (fun rect -> rect.Width = 1.0 && rect.Height <= 2.0)
eq "painter: the axis is a dashed line down every cell, at one x" (axisPieces.Length > 20, axisPieces |> List.map (fun rect -> round3 rect.X) |> List.distinct |> List.length) (true, 1)
eq "painter: the axis is where the bars meet" (round3 axisPieces.Head.X) (round3 paintedPositive.X)
// A gradient is strips of the colour getting lighter.
let gradientItems = painted standard
eq "painter: a gradient bar is drawn in strips" ((fills "#638EC6" gradientItems).Length, gradientItems |> List.choose (function Paint.FillRect (_, c) when c <> "#638EC6" && c.StartsWith "#" && c.Length = 7 -> Some c | _ -> None) |> List.distinct |> List.length > 8) (2, true)
// Values hidden: no text in the cells of the rule.
let textsOf items = items |> List.choose (function Paint.DrawText text -> Some text.Text | _ -> None)
eq "painter: values shown" (textsOf solidItems |> List.filter (fun t -> t = "100" || t = "-50") |> List.sort) [ "-50"; "100" ]
eq "painter: values hidden" (textsOf (painted bordered) |> List.filter (fun t -> t = "100" || t = "-50")) []
// Icons: arrows are drawn as shapes, pointing down for the lowest and up for the highest.
let arrowItems = painted (IconSetEx ("3TrafficLights1", steps, false, { ShowValue = true; Custom = [ "3Arrows", 0; "NoIcons", 0; "3Arrows", 2 ] }))
let arrows = arrowItems |> List.choose (function Paint.VecPath (segments, Some color, None) -> Some (color, segments) | _ -> None)
eq "painter: an arrow per cell that has an icon, red down and green up" (arrows |> List.map fst |> List.countBy id |> List.sort) [ "#43A047", 1; "#E53935", 2 ]

printfn ""
printfn "Stage 47 x14 rule tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
