#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
open Rbt.Spreadsheet.ConditionalFormatting
module Ops = Rbt.Spreadsheet.SheetOps

// What the conditional-format part of the host API stands on: a rule as plain
// data and back (every kind the model has), what a description is refused
// for, rules addressed by id (add, update, remove, reorder, clear by range),
// ids through structure edits and the file shape, and "a date occurring".

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
let fill color = { defaultFormat with BgColor = Some color }
let empty = createEmpty 100 26

// ---- several areas, and a range without some cells --------------------------------
eq "areas by spaces" (Ops.tryParseRanges "A1:A5 C1:C5") (Some [ block 0 0 4 0; block 0 2 4 2 ])
eq "areas by commas" (Ops.tryParseRanges "A:A, C3") (Some [ block 0 0 Ops.maxSheetRow 0; block 2 2 2 2 ])
eq "one bad area spoils it" (Ops.tryParseRanges "A1 nope") None
eq "no area" (Ops.tryParseRanges "  ") None
eq "areas as text" (Ops.rangesText [ block 0 0 4 0; block 0 2 Ops.maxSheetRow 2 ]) "A1:A5 C:C"
throws "parseRanges names the member" (fun () -> Ops.parseRanges "AddConditionalFormat" "A1;;x" |> ignore) "AddConditionalFormat: \"A1;;x\" is not a range"
eq "a hole in the middle" (Ops.subtractRange (block 2 2 3 3) (block 0 0 5 5)) [ block 0 0 1 5; block 2 0 3 1; block 2 4 3 5; block 4 0 5 5 ]
eq "a hole over a corner" (Ops.subtractRange (block 0 0 1 1) (block 1 1 3 3)) [ block 1 2 1 3; block 2 1 3 3 ]
eq "a hole elsewhere" (Ops.subtractRange (block 10 10 11 11) (block 0 0 5 5)) [ block 0 0 5 5 ]
eq "a hole over everything" (Ops.subtractRange (block 0 0 9 9) (block 1 1 3 3)) []
eq "rows out of a whole column" (Ops.subtractRange (block 2 0 4 3) (block 0 1 Ops.maxSheetRow 1)) [ block 0 1 1 1; block 5 1 Ops.maxSheetRow 1 ]

// ---- every kind of rule, as plain data and back -----------------------------------
let lo, hi = cfvo CfvoMin None, cfvo CfvoMax None
let everyRule : ConditionalFormatRule list =
    [ for op in [ OpGreaterThan; OpLessThan; OpGreaterOrEqual; OpLessOrEqual; OpEqual; OpNotEqual ] -> CellIs (op, "10", None)
      yield CellIs (OpBetween, "$C$1", Some "\"m\"")
      yield CellIs (OpNotBetween, "1", Some "9")
      yield Expression "AND($A1>0,MOD(ROW(),2)=0)"
      for kind in [ TextContains; TextNotContains; TextBeginsWith; TextEndsWith ] -> TextRule (kind, "a\"b")
      yield TopBottom (3, false, false)
      yield TopBottom (25, true, true)
      yield AboveAverage (true, false, 0)
      yield AboveAverage (false, true, 2)
      yield DuplicateValues
      yield UniqueValues
      yield Blanks true
      yield Blanks false
      yield Errors true
      yield Errors false
      for period in timePeriods -> TimePeriod period
      yield ColorScale [ lo, "#F8696B"; hi, "#63BE7B" ]
      yield ColorScale [ lo, "#F8696B"; cfvo CfvoPercentile (Some "50"), "#FFEB84"; cfvo CfvoNumber (Some "90"), "#63BE7B" ]
      yield DataBar (cfvo CfvoPercent (Some "10"), cfvo CfvoFormula (Some "$C$1*2"), "#638EC6")
      yield IconSet ("3TrafficLights1", [ cfvo CfvoPercent (Some "0"); cfvo CfvoPercent (Some "33"); cfvo CfvoPercent (Some "67") ], false)
      yield IconSet ("5Arrows", [ for p in [ "0"; "20"; "40"; "60"; "80" ] -> cfvo CfvoPercentile (Some p) ], true) ]
let style = { defaultFormat with BgColor = Some "#FFC7CE"; TextColor = Some "#9C0006"; Bold = true; Underline = true; Strikethrough = true }
for rule in everyRule do
    let format = { Rule = rule; Style = (if Ops.isVisualRule rule then defaultFormat else style) }
    eq (sprintf "there and back: %A" rule) (Ops.ruleOfDescription "T" defaultFormat (Ops.describeRule format)) format
eq "every rule type is described" (everyRule |> List.map (fun rule -> (Ops.describeRule { Rule = rule; Style = style }).Type) |> List.distinct |> List.sort) (List.sort Ops.ruleTypes)
// The literal rules of the old dialog are the general ones they are equal to.
let described rule = let d = Ops.describeRule { Rule = rule; Style = style } in d.Type, d.Operator, d.Formula1, d.Formula2, d.Text
eq "literal greater than" (described (GreaterThan 5.0)) ("cellIs", Some "greaterThan", Some "5", None, None)
eq "literal less than" (described (LessThan 2.5)) ("cellIs", Some "lessThan", Some "2.5", None, None)
eq "literal equal" (described (Equal -3.0)) ("cellIs", Some "equal", Some "-3", None, None)
eq "literal between" (described (Between (1.0, 9.0))) ("cellIs", Some "between", Some "1", Some "9", None)
eq "literal contains" (described (ContainsText "north")) ("containsText", None, None, None, Some "north")
eq "formulas are given without =" (Ops.describeRule { Rule = Expression "=A1>5"; Style = style }).Formula1 (Some "A1>5")
eq "a rule's format is the six properties" (Ops.describeRule { Rule = DuplicateValues; Style = { style with FontSize = Some 20; WrapText = true } }).Style (Some style)
eq "a visual rule has no format" (Ops.describeRule { Rule = DataBar (lo, hi, "#638EC6"); Style = style }).Style None

// ---- what a description is refused for ----------------------------------------------
let d kind = Ops.emptyRuleDescription kind
let build description = Ops.ruleOfDescription "Add" defaultFormat description
eq "a leading = is dropped" (build { d "expression" with Formula1 = Some " =A1>5 " }).Rule (Expression "A1>5")
eq "the fallback style is used" (Ops.ruleOfDescription "Add" style (d "blanks")).Style style
eq "a given style wins, cut to what a rule applies" (Ops.ruleOfDescription "Add" style { d "blanks" with Style = Some { fill "#00FF00" with Indent = 3 } }).Style (fill "#00FF00")
eq "data bar defaults" (build (d "dataBar")).Rule (DataBar (lo, hi, "#638EC6"))
eq "icon set defaults" (build { d "iconSet" with IconSet = Some "4Rating" }).Rule
   (IconSet ("4Rating", [ for p in [ "0"; "25"; "50"; "75" ] -> cfvo CfvoPercent (Some p) ], false))
throws "unknown type" (fun () -> build (d "sparkle") |> ignore) "Add: unknown rule type \"sparkle\""
throws "cellIs without operator" (fun () -> build { d "cellIs" with Formula1 = Some "1" } |> ignore) "Add: rule.operator must be one of"
throws "cellIs with a bad operator" (fun () -> build { d "cellIs" with Operator = Some ">"; Formula1 = Some "1" } |> ignore) "Add: rule.operator must be one of"
throws "cellIs without operand" (fun () -> build { d "cellIs" with Operator = Some "equal" } |> ignore) "Add: a \"cellIs\" rule needs rule.formula1"
throws "between without the other bound" (fun () -> build { d "cellIs" with Operator = Some "between"; Formula1 = Some "1" } |> ignore) "Add: a \"cellIs\" rule needs rule.formula2"
throws "a second operand that is not used" (fun () -> build { d "cellIs" with Operator = Some "equal"; Formula1 = Some "1"; Formula2 = Some "2" } |> ignore) "Add: rule.formula2 is only for"
throws "expression without formula" (fun () -> build { d "expression" with Formula1 = Some "=" } |> ignore) "Add: a \"expression\" rule needs rule.formula"
throws "text rule without text" (fun () -> build (d "beginsWith") |> ignore) "Add: a \"beginsWith\" rule needs rule.text"
throws "rank 0" (fun () -> build { d "top" with Rank = 0 } |> ignore) "Add: rule.rank must be between 1 and 1000"
throws "percent above 100" (fun () -> build { d "top" with Rank = 101; Percent = true } |> ignore) "Add: rule.rank must be between 1 and 100,"
throws "stdDev 4" (fun () -> build { d "aboveAverage" with StdDev = 4 } |> ignore) "Add: rule.stdDev must be between 0 and 3"
throws "an unknown period" (fun () -> build { d "dateOccurring" with Period = Some "someday" } |> ignore) "Add: rule.period must be one of \"today\""
throws "one stop" (fun () -> build { d "colorScale" with Stops = [ { Kind = "min"; Value = None }, "#FF0000" ] } |> ignore) "Add: a colour scale has 2 or 3 stops, got 1"
let stop kind value = ({ Kind = kind; Value = value } : Ops.ThresholdDescription)
throws "a stop of an unknown type" (fun () -> build { d "colorScale" with Stops = [ stop "lowest" None, "#FF0000"; stop "max" None, "#00FF00" ] } |> ignore) "Add: rule.stops[0].type must be one of"
throws "a stop without its number" (fun () -> build { d "colorScale" with Stops = [ stop "min" None, "#FF0000"; stop "percent" None, "#00FF00" ] } |> ignore) "Add: rule.stops[1].value is missing"
throws "a stop that is no number" (fun () -> build { d "colorScale" with Stops = [ stop "min" None, "#FF0000"; stop "number" (Some "lots"), "#00FF00" ] } |> ignore) "Add: rule.stops[1].value must be a number"
throws "a percentage above 100" (fun () -> build { d "colorScale" with Stops = [ stop "min" None, "#FF0000"; stop "percentile" (Some "120"), "#00FF00" ] } |> ignore) "Add: rule.stops[1].value must be between 0 and 100"
throws "a stop of a bad colour" (fun () -> build { d "colorScale" with Stops = [ stop "min" None, "red"; stop "max" None, "#00FF00" ] } |> ignore) "Add: rule.stops[0].color must be a colour"
throws "a bar of a bad colour" (fun () -> build { d "dataBar" with Color = Some "#12" } |> ignore) "Add: rule.color must be a colour"
throws "an unknown icon set" (fun () -> build { d "iconSet" with IconSet = Some "3Stars" } |> ignore) "Add: rule.iconSet must be one of 3Arrows"
throws "too few thresholds" (fun () -> build { d "iconSet" with IconSet = Some "5Rating"; Thresholds = Ops.defaultIconThresholds 3 } |> ignore) "Add: icon set \"5Rating\" needs 5 thresholds"
throws "min as an icon threshold" (fun () -> build { d "iconSet" with IconSet = Some "3Flags"; Thresholds = [ stop "min" None; stop "percent" (Some "33"); stop "percent" (Some "67") ] } |> ignore) "Add: rule.thresholds[0].type must be one of \"number\""
throws "a format on a visual rule" (fun () -> build { d "dataBar" with Style = Some style } |> ignore) "Add: a \"dataBar\" rule draws by itself"

// ---- rules by id ------------------------------------------------------------------------
let add id ranges rule = Ops.addConditionalFormatWithId id ranges false { Rule = rule; Style = fill "#FF0000" }
let three = empty |> add 1 [ block 0 0 9 0 ] DuplicateValues |> add 2 [ block 0 1 9 1 ] UniqueValues |> add 3 [ block 0 2 9 2 ] (Blanks true)
let order (m: SpreadsheetModel) = Ops.conditionalFormatsInOrder m |> List.map (fun entry -> entry.Id)
eq "the newest rule comes first" (order three) [ 3; 2; 1 ]
eq "highest id" (Ops.highestConditionalFormatId three, Ops.highestConditionalFormatId empty) (3, 0)
eq "the model's own add numbers on" ((three |> addConditionalFormat (at 5 5) { Rule = Errors true; Style = defaultFormat }).ConditionalFormats.Head.Id) 4
eq "find by id" (Ops.tryFindConditionalFormat 2 three |> Option.map (fun entry -> entry.Format.Rule), Ops.tryFindConditionalFormat 9 three) (Some UniqueValues, None)
let reordered = Ops.reorderConditionalFormats "Reorder" [ 1; 3; 2 ] three
eq "reorder" (order reordered, Ops.conditionalFormatsInOrder reordered |> List.map (fun entry -> entry.Priority)) ([ 1; 3; 2 ], [ 1; 2; 3 ])
eq "reorder to the same order changes nothing" (obj.ReferenceEquals(Ops.reorderConditionalFormats "Reorder" [ 1; 3; 2 ] reordered, reordered)) true
throws "reorder with an unknown id" (fun () -> Ops.reorderConditionalFormats "Reorder" [ 1; 2; 7 ] three |> ignore) "Reorder: the sheet has no conditional format with id 7"
throws "reorder with one left out" (fun () -> Ops.reorderConditionalFormats "Reorder" [ 1; 2 ] three |> ignore) "Reorder: every conditional format of the sheet must be listed once (2 listed, the sheet has 3: 3, 2, 1)"
throws "reorder with one twice" (fun () -> Ops.reorderConditionalFormats "Reorder" [ 1; 2; 2 ] three |> ignore) "Reorder: an id is listed twice"
eq "move to the front" (order (Ops.moveConditionalFormat "Move" 1 1 three)) [ 1; 3; 2 ]
eq "move to the end" (order (Ops.moveConditionalFormat "Move" 3 3 three)) [ 2; 1; 3 ]
eq "move to the middle" (order (Ops.moveConditionalFormat "Move" 3 2 three)) [ 2; 3; 1 ]
throws "move past the end" (fun () -> Ops.moveConditionalFormat "Move" 3 4 three |> ignore) "Move: priority must be between 1 and 3, got 4"
let updated = Ops.updateConditionalFormat "Update" 2 (fun entry -> { entry with Id = 99; Ranges = [ block 0 5 1 5 ]; StopIfTrue = true }) three
eq "update keeps the id and the place" (order updated, Ops.tryFindConditionalFormat 2 updated |> Option.map (fun entry -> entry.Ranges, entry.StopIfTrue)) ([ 3; 2; 1 ], Some ([ block 0 5 1 5 ], true))
eq "an update that changes nothing gives the same model" (obj.ReferenceEquals(Ops.updateConditionalFormat "Update" 2 id three, three)) true
throws "update of an unknown id" (fun () -> Ops.updateConditionalFormat "Update" 8 id three |> ignore) "Update: the sheet has no conditional format with id 8"
eq "remove" (order (Ops.removeConditionalFormat "Remove" 2 three)) [ 3; 1 ]
throws "remove of an unknown id" (fun () -> Ops.removeConditionalFormat "Remove" 8 three |> ignore) "Remove: the sheet has no conditional format with id 8"
eq "clear everything" (Ops.clearConditionalFormats None three).ConditionalFormats []
eq "clear nothing on an empty sheet" (obj.ReferenceEquals(Ops.clearConditionalFormats None empty, empty)) true
let cut = Ops.clearConditionalFormats (Some [ block 0 0 9 0; block 2 1 3 1 ]) three
eq "clear by range: one gone, one cut, one kept" (Ops.conditionalFormatsInOrder cut |> List.map (fun entry -> entry.Id, entry.Ranges))
   [ 3, [ block 0 2 9 2 ]; 2, [ block 0 1 1 1; block 4 1 9 1 ] ]
eq "clear where no rule is" (obj.ReferenceEquals(Ops.clearConditionalFormats (Some [ block 50 0 60 9 ]) three, three)) true

// Ids stay with their rules when rows move, and a file's rules are numbered by precedence.
let moved = insertRow 0 three
eq "ids after inserting a row" (Ops.conditionalFormatsInOrder moved |> List.map (fun entry -> entry.Id, entry.Ranges.Head.From.Row)) [ 3, 1; 2, 1; 1, 1 ]
let viaFile = Ops.conditionalFormatsInOrder reordered |> toOoxml |> ofOoxml
eq "a file's rules: 1, 2, 3 by precedence" (viaFile |> List.map (fun entry -> entry.Id, entry.Format.Rule)) [ 1, DuplicateValues; 2, Blanks true; 3, UniqueValues ]
let styled = empty |> Ops.addConditionalFormatWithId 1 [ block 0 0 9 0 ] true { Rule = Errors false; Style = style }
eq "the whole format and stop-if-true through the file shape" (styled.ConditionalFormats |> toOoxml |> ofOoxml |> List.map (fun entry -> entry.Format.Style, entry.StopIfTrue)) [ style, true ]

// ---- a date occurring ---------------------------------------------------------------------
let today = match evaluateFormulaAt "=TODAY()" 0 0 empty (at 0 0) |> unwrapFormulaValue with CellValue.Number n -> n | _ -> nan
let weekday = float ((int today - 1) % 7 + 1)     // 1 = Sunday
let date offset = { Value = CellValue.Number (today + offset + 0.5); Format = defaultFormat; Comment = None }
// A1:A9 = today, yesterday, tomorrow, 6 days ago, 7 days ago, last Saturday, next Sunday, a text, a blank
let lastSaturday, nextSunday = -weekday, 8.0 - weekday
let dates =
    { empty with
        Cells = Map.ofList [ at 0 0, date 0.0; at 1 0, date -1.0; at 2 0, date 1.0; at 3 0, date -6.0; at 4 0, date -7.0
                             at 5 0, date lastSaturday; at 6 0, date nextSunday
                             at 7 0, { Value = CellValue.String "soon"; Format = defaultFormat; Comment = None } ] }
let occurring period =
    let m = dates |> add 1 [ block 0 0 8 0 ] (TimePeriod period)
    [ for row in 0 .. 8 -> (effectAt (createContext evaluateFormulaAt m) (at row 0)).IsSome ]
eq "today is known" (today > 40000.0) true
eq "today" (occurring "today") [ true; false; false; false; false; false; false; false; false ]
eq "yesterday" (occurring "yesterday") [ false; true; false; false; false; lastSaturday = -1.0; false; false; false ]
eq "tomorrow" (occurring "tomorrow") [ false; false; true; false; false; false; nextSunday = 1.0; false; false ]
eq "last 7 days" (occurring "last7Days") [ true; true; false; true; false; lastSaturday >= -6.0; false; false; false ]
eq "this week" ((occurring "thisWeek").[0], (occurring "thisWeek").[5], (occurring "thisWeek").[6], (occurring "thisWeek").[7]) (true, false, false, false)
eq "last week" ((occurring "lastWeek").[0], (occurring "lastWeek").[5], (occurring "lastWeek").[6]) (false, true, false)
eq "next week" ((occurring "nextWeek").[0], (occurring "nextWeek").[5], (occurring "nextWeek").[6]) (false, false, true)
eq "this month" ((occurring "thisMonth").[0], (occurring "thisMonth").[7], (occurring "thisMonth").[8]) (true, false, false)
let monthAway offset = { dates with Cells = Map.ofList [ at 0 0, { Value = evaluateFormulaAt (sprintf "=EDATE(TODAY(),%d)" offset) 0 0 empty (at 0 0) |> unwrapFormulaValue; Format = defaultFormat; Comment = None } ] }
let matches period (m: SpreadsheetModel) = (effectAt (createContext evaluateFormulaAt (m |> add 1 [ block 0 0 8 0 ] (TimePeriod period))) (at 0 0)).IsSome
eq "last month" (matches "lastMonth" (monthAway -1), matches "lastMonth" dates, matches "nextMonth" (monthAway -1)) (true, false, false)
eq "next month" (matches "nextMonth" (monthAway 1), matches "nextMonth" dates, matches "thisMonth" (monthAway 1)) (true, false, false)
// The rule follows its cells: written for the first cell of its range, wherever that is.
let shifted = { empty with Cells = Map.ofList [ at 4 3, date 0.0; at 5 3, date -1.0 ] } |> add 1 [ block 4 3 5 3 ] (TimePeriod "today")
eq "a range away from A1" [ for row in 4 .. 5 -> (effectAt (createContext evaluateFormulaAt shifted) (at row 3)).IsSome ] [ true; false ]
eq "the file shape of a period" (toOoxml (dates |> add 1 [ block 0 0 8 0 ] (TimePeriod "today")).ConditionalFormats |> List.collect (fun cf -> cf.Rules |> List.map (fun r -> r.Type, r.TimePeriod, r.Formulas)))
   [ "timePeriod", Some "today", [ "FLOOR(A1,1)=TODAY()" ] ]

printfn "stage34-cfapi: %d passed, %d failed" pass fail
if fail > 0 then exit 1
