module Rbt.Spreadsheet.Managers

#if FABLE_COMPILER
open System
open Fable.Core
open Fable.Core.JsInterop
open Rbt.FableUI.OfficeUiDialog
open Rbt.Office.Ui
open Rbt.Spreadsheet.Dom

/// Editor-owned workflow boundary: the grid supplies its existing validated,
/// undoable host operations without depending on dialog/chrome modules.
type SpreadsheetPort = {
    GetWorkbookEpoch: unit -> int
    GetChangeSerial: unit -> int
    GetActiveSheetIndex: unit -> int
    IsReadOnly: unit -> bool
    IsEditing: unit -> bool
    GetConditionalFormats: obj -> obj array
    GetSelectionRef: unit -> string
    UpdateConditionalFormat: obj * obj -> unit
    AddConditionalFormat: obj * obj * obj * obj * obj -> int
    RemoveConditionalFormat: obj -> unit
    GetDefinedNames: unit -> obj array
    GetSheetNames: unit -> string array
    SetDefinedName: obj * obj * obj * obj -> unit
    DeleteDefinedName: obj * obj -> unit
}

[<Emit("$0 == null")>]
let private missing (item: obj) : bool = jsNative
[<Emit("Object.assign({}, $0)")>]
let private copy (item: obj) : obj = jsNative
[<Emit("String($0)")>]
let private text (item: obj) : string = jsNative
[<Emit("String(($0 && $0.message) || $0)")>]
let private errorText (item: obj) : string = jsNative

let private prop item name = if missing item then null else getProp item name
let private value item name = let v = prop item name in if missing v then "" else text v
let private flag item name = value item name = "true"
let private choices entries = entries |> List.map (fun (key, label) -> choice key label)
let private select id title entries initial = dialogSelect id title (choices entries) initial
let private status = dialogStatus "error" "Validation" ""
let private fail (dialog: DialogState) (error: exn) =
    match dialog.Open.Peek() with
    | Some opened ->
        let message = errorText (box error)
        opened.Values.["error"].Value <- message
        if opened.Values.ContainsKey "error-format" then opened.Values.["error-format"].Value <- message
    | None -> ()
    false

// A dialog observes one workbook version. A successful action opens a fresh
// manager, with a new version; changes elsewhere never retarget an open form.
type private Context = { Epoch: int; Serial: int; Sheet: int; Selection: string }
let private capture (grid: SpreadsheetPort) =
    { Epoch = grid.GetWorkbookEpoch(); Serial = grid.GetChangeSerial()
      Sheet = grid.GetActiveSheetIndex(); Selection = grid.GetSelectionRef() }
let private ready (grid: SpreadsheetPort) (context: Context) =
    if grid.IsReadOnly() then invalidOp "The workbook is read-only."
    if grid.GetWorkbookEpoch() <> context.Epoch || grid.GetChangeSerial() <> context.Serial
       || grid.GetActiveSheetIndex() <> context.Sheet then
        invalidOp "The workbook or sheet changed while the dialog was open. Reopen the manager."
    if grid.IsEditing() then invalidOp "Correct or cancel the pending cell entry before applying."

let private changed (before: Map<string, string>) (after: Map<string, string>) key = before.[key] <> after.[key]
let private integer label (input: string) =
    match Int32.TryParse(input.Trim()) with
    | true, number -> number
    | _ -> invalidArg label ("Enter a whole number for " + label + ".")
let private color label (input: string) =
    let input = input.Trim()
    if input <> "" && not (System.Text.RegularExpressions.Regex.IsMatch(input, "^#[0-9a-fA-F]{6}$")) then
        invalidArg label ("Enter #RRGGBB or leave " + label + " empty.")
    input

let private ruleKinds = [
    "cellIs", "Compare cells"; "expression", "Formula is true"
    "containsText", "Text contains"; "notContainsText", "Text does not contain"
    "beginsWith", "Text begins with"; "endsWith", "Text ends with"
    "top", "Top or bottom values"; "aboveAverage", "Above or below average"
    "duplicateValues", "Duplicate values"; "uniqueValues", "Unique values"
    "blanks", "Blanks"; "noBlanks", "Nonblank cells"; "errors", "Errors"; "noErrors", "No errors"
    "dateOccurring", "Dates in period"; "colorScale", "Color scale"; "dataBar", "Data bar"; "iconSet", "Icon set" ]
let private comparisons = [
    "greaterThan", "Greater than"; "lessThan", "Less than"
    "greaterThanOrEqual", "Greater than or equal"; "lessThanOrEqual", "Less than or equal"
    "equal", "Equal"; "notEqual", "Not equal"; "between", "Between"; "notBetween", "Not between" ]
let private periods = [
    "today", "Today"; "yesterday", "Yesterday"; "tomorrow", "Tomorrow"
    "last7Days", "Last seven days"; "thisWeek", "This week"; "lastWeek", "Last week"
    "nextWeek", "Next week"; "thisMonth", "This month"; "lastMonth", "Last month"; "nextMonth", "Next month" ]
let private thresholdKinds = [
    "min", "Minimum"; "max", "Maximum"; "autoMin", "Automatic minimum"; "autoMax", "Automatic maximum"
    "number", "Number"; "percent", "Percent"; "percentile", "Percentile"; "formula", "Formula" ]
let private iconFamilies =
    Rbt.Spreadsheet.SheetOps.iconSetNames @ Rbt.Spreadsheet.ConditionalFormatting.extendedIconSets
    |> List.map (fun family -> family, family)
let private thresholdFields key title (source: obj) defaultKind defaultValue allowed = [
    select (key + "Type") (title + " threshold") allowed
        (if value source "type" = "" then defaultKind else value source "type")
    dialogText (key + "Value") (title + " value (number or formula)") "0"
        (if missing source then defaultValue else value source "value") ]
let private thresholdFrom (fields: Map<string, string>) key =
    let kind = fields.[key + "Type"]
    let result = createObj [ "type" ==> kind ]
    if not (List.contains kind [ "min"; "max"; "autoMin"; "autoMax" ]) then
        setProp result "value" (box (fields.[key + "Value"].Trim()))
    result

let private ruleFields kind rule =
    let field name = value rule name
    match kind with
    | "cellIs" -> [
        select "operator" "Comparison" comparisons (if field "operator" = "" then "greaterThan" else field "operator")
        dialogText "formula1" "First operand" "10, $A$1 or formula" (field "formula1")
        dialogText "formula2" "Second operand (between only)" "20" (field "formula2") ]
    | "expression" -> [ dialogText "formula" "Formula (for the first cell in the range)" "A1>10" (field "formula") ]
    | "containsText" | "notContainsText" | "beginsWith" | "endsWith" ->
        [ dialogText "text" "Match text" "Text to match" (field "text") ]
    | "top" -> [ dialogNumber "rank" "Count" (Some "1") (Some "1000") (if field "rank" = "" then "10" else field "rank")
                 dialogCheckbox "percent" "Count is a percentage" (flag rule "percent")
                 dialogCheckbox "bottom" "Bottom rather than top" (flag rule "bottom") ]
    | "aboveAverage" ->
        [ dialogCheckbox "above" "Above rather than below average" (missing rule || flag rule "above")
          dialogCheckbox "orEqual" "Include values equal to the average" (flag rule "orEqual")
          dialogNumber "stdDev" "Standard deviations (0–3)" (Some "0") (Some "3") (if field "stdDev" = "" then "0" else field "stdDev") ]
    | "dateOccurring" -> [ select "period" "Date period" periods (if field "period" = "" then "today" else field "period") ]
    | "colorScale" ->
        let stops = if missing (prop rule "stops") then [||] else unbox<obj array> (prop rule "stops")
        [ select "stopCount" "Stops" [ "2", "Two stops"; "3", "Three stops" ] (if stops.Length = 3 then "3" else "2") ]
        @ [ for index in 0..2 do
                let source = if index < stops.Length then stops.[index] else null
                let last = if stops.Length = 3 then 2 else 1
                yield! thresholdFields ("stop" + string index) (sprintf "Stop %d" (index + 1)) source
                    (if index = 0 then "min" elif index = last then "max" else "percentile") "50"
                    (thresholdKinds |> List.filter (fun (name, _) -> name <> "autoMin" && name <> "autoMax"))
                yield dialogText ("stop" + string index + "Color") (sprintf "Stop %d color" (index + 1)) "#RRGGBB"
                    (if missing source then (if index = 0 then "#F8696B" elif index = last then "#63BE7B" else "#FFEB84")
                     else value source "color") ]
    | "dataBar" ->
        thresholdFields "min" "Minimum" (prop rule "min") "min" ""
            (thresholdKinds |> List.filter (fun (name, _) -> name <> "autoMax"))
        @ thresholdFields "max" "Maximum" (prop rule "max") "max" ""
            (thresholdKinds |> List.filter (fun (name, _) -> name <> "autoMin"))
        @ [ dialogText "barColor" "Bar color" "#RRGGBB" (if field "color" = "" then "#638EC6" else field "color") ]
    | "iconSet" ->
        let thresholds = if missing (prop rule "thresholds") then [||] else unbox<obj array> (prop rule "thresholds")
        let family = if field "iconSet" = "" then "3TrafficLights1" else field "iconSet"
        let defaults = Rbt.Spreadsheet.SheetOps.defaultIconThresholds (int family.[0] - int '0') |> List.toArray
        [ select "family" "Icon family" iconFamilies family
          dialogCheckbox "reverse" "Reverse icon order" (flag rule "reverse") ]
        @ [ for index in 0..4 do
                let source = if index < thresholds.Length then thresholds.[index] else null
                let defaultValue = if index < defaults.Length then defaults.[index].Value |> Option.defaultValue ""
                                   else string (index * 20)
                yield! thresholdFields ("icon" + string index) (sprintf "Icon %d floor" (index + 1)) source
                    "percent" defaultValue
                    (thresholdKinds |> List.filter (fun (name, _) -> List.contains name [ "number"; "percent"; "percentile"; "formula" ])) ]
    | _ -> []

let private formatFields (rule: obj) =
    let format = prop rule "format"
    [ dialogText "fill" "Fill (#RRGGBB; empty: no fill)" "#RRGGBB"
          (if missing rule then "#FFCCCC" else value format "fill")
      dialogText "textColor" "Text color (#RRGGBB; empty: automatic)" "#RRGGBB" (value format "textColor")
      dialogCheckbox "bold" "Bold" (flag format "bold")
      dialogCheckbox "italic" "Italic" (flag format "italic")
      dialogCheckbox "underline" "Underline" (flag format "underline")
      dialogCheckbox "strikethrough" "Strikethrough" (flag format "strikethrough") ]

let private updateRule (kind: string) (original: obj) (before: Map<string, string>) (fields: Map<string, string>) =
    let existing = not (missing original) && value original "type" = kind
    let rule = if existing then copy original else createObj [ "type" ==> kind ]
    let assign key field = if not existing || changed before fields field then setProp rule key (box fields.[field])
    let boolean key = if not existing || changed before fields key then setProp rule key (box (fields.[key] = "true"))
    let number key = if not existing || changed before fields key then setProp rule key (box (integer key fields.[key]))
    match kind with
    | "cellIs" ->
        assign "operator" "operator"; assign "formula1" "formula1"
        if List.contains fields.["operator"] [ "between"; "notBetween" ] then assign "formula2" "formula2"
        elif fields.["formula2"].Trim() <> "" && changed before fields "formula2" then
            invalidArg "formula2" "A second operand is only used for between and not between."
        elif existing && changed before fields "operator" then setProp rule "formula2" null
    | "expression" -> assign "formula" "formula"
    | "containsText" | "notContainsText" | "beginsWith" | "endsWith" -> assign "text" "text"
    | "top" -> number "rank"; boolean "percent"; boolean "bottom"
    | "aboveAverage" -> boolean "above"; boolean "orEqual"; number "stdDev"
    | "dateOccurring" -> assign "period" "period"
    | "colorScale" ->
        let count = integer "stopCount" fields.["stopCount"]
        let oldStops = if existing then unbox<obj array> (prop original "stops") else [||]
        let stops = [| for i in 0 .. count - 1 do
                         let key = "stop" + string i
                         if existing && i < oldStops.Length &&
                            ([ key + "Type"; key + "Value"; key + "Color" ] |> List.forall (fun name -> not (changed before fields name))) then
                             yield oldStops.[i]
                         else
                             let stop = thresholdFrom fields key
                             setProp stop "color" (box (color (key + "Color") fields.[key + "Color"]))
                             yield stop |]
        setProp rule "stops" (box stops)
    | "dataBar" ->
        for key in [ "min"; "max" ] do
            if not existing || changed before fields (key + "Type") || changed before fields (key + "Value") then
                setProp rule key (thresholdFrom fields key)
        if not existing || changed before fields "barColor" then setProp rule "color" (box (color "bar color" fields.["barColor"]))
    | "iconSet" ->
        assign "iconSet" "family"; boolean "reverse"
        let family = fields.["family"]
        let count = int family.[0] - int '0'
        let oldSteps = if existing then unbox<obj array> (prop original "thresholds") else [||]
        if not existing || changed before fields "family" ||
           ([ 0 .. count - 1 ] |> List.exists (fun i -> changed before fields ("icon" + string i + "Type") || changed before fields ("icon" + string i + "Value"))) then
            let familyChanged = existing && changed before fields "family"
            let defaults = Rbt.Spreadsheet.SheetOps.defaultIconThresholds count |> List.toArray
            setProp rule "thresholds" (box [| for i in 0 .. count - 1 do
                                                  let key = "icon" + string i
                                                  let edited = changed before fields (key + "Type") || changed before fields (key + "Value")
                                                  if existing && not familyChanged && i < oldSteps.Length && not edited then
                                                      yield oldSteps.[i]
                                                  elif familyChanged && not edited then
                                                      yield createObj [ "type" ==> "percent"; "value" ==> (defaults.[i].Value |> Option.defaultValue "0") ]
                                                  else yield thresholdFrom fields key |])
    | _ -> ()
    if not (List.contains kind [ "colorScale"; "dataBar"; "iconSet" ]) then
        if not existing then
            let format = createObj [ "fill" ==> "#FFCCCC" ]
            setProp rule "format" format
        let format =
            if existing then
                if missing (prop original "format") then createObj [] else copy (prop original "format")
            else prop rule "format"
        for key in [ "fill"; "textColor"; "bold"; "italic"; "underline"; "strikethrough" ] do
            if changed before fields key then
                let item = if key = "fill" || key = "textColor" then
                               let input = color key fields.[key]
                               if input = "" then null else box input
                           else box (fields.[key] = "true")
                setProp format key item
        setProp rule "format" format
    rule

let openConditionalManager (grid: SpreadsheetPort) (dialog: DialogState) =
    let rec listing () =
        let context = capture grid
        let rules = grid.GetConditionalFormats(null)
        let options = [ "", "Choose a rule" ] @
                      [ for rule in rules ->
                            let id = value rule "id"
                            let kind = value (prop rule "rule") "type"
                            id, sprintf "%s · #%s · %s · %s" (value rule "priority") id kind (value rule "range") ]
        let selected (fields: Map<string, string>) =
            let id = fields.["selected"]
            match rules |> Array.tryFind (fun rule -> value rule "id" = id) with
            | Some rule -> rule
            | None -> invalidOp "Choose a rule from the list."
        let action callback fields =
            try ready grid context; callback fields
            with error -> fail dialog error
        openDialog dialog {
            TestId = "xl-dlg-conditional-manager"; Title = "Conditional formatting"
            Subtitle = Some "Rules are ordered by priority. Stop-if-true prevents lower-priority rules from applying."
            Fields = [ select "selected" "Sheet rules (priority · id · kind · range)" options (if rules.Length = 0 then "" else value rules.[0] "id")
                       dialogComputed "details" "Rule details" (fun fields ->
                           let rule = rules |> Array.tryFind (fun r -> value r "id" = fields.["selected"])
                           match rule with
                           | None -> "No rule selected."
                           | Some entry -> sprintf "Range: %s · Priority: %s · Stop if true: %s"
                                                (value entry "range") (value entry "priority") (value entry "stopIfTrue"))
                       status ]
            Tabs = []
            Buttons = [ dialogCancel
                        dialogButton "xl-dlg-conditional-add" "New rule" false (action (fun _ -> picking context; false))
                        dialogButton "xl-dlg-conditional-edit" "Edit rule" true (action (fun fields -> editing context (Some (selected fields)) None; false))
                        dialogButton "xl-dlg-conditional-up" "Move up" false (action (fun fields ->
                            let rule = selected fields
                            let priority = integer "priority" (value rule "priority")
                            if priority > 1 then grid.UpdateConditionalFormat(box (integer "id" (value rule "id")), createObj [ "priority" ==> (priority - 1) ])
                            listing(); false))
                        dialogButton "xl-dlg-conditional-down" "Move down" false (action (fun fields ->
                            let rule = selected fields
                            let priority = integer "priority" (value rule "priority")
                            if priority < rules.Length then grid.UpdateConditionalFormat(box (integer "id" (value rule "id")), createObj [ "priority" ==> (priority + 1) ])
                            listing(); false))
                        dialogButton "xl-dlg-conditional-delete" "Delete rule" false (action (fun fields ->
                            let rule = selected fields
                            grid.RemoveConditionalFormat(box (integer "id" (value rule "id")))
                            listing(); false)) ] }
    and picking context =
        openDialog dialog {
            TestId = "xl-dlg-conditional-kind"; Title = "New conditional rule"
            Subtitle = Some "Choose a rule type; its editor opens next."
            Fields = [ select "kind" "Rule type" ruleKinds "cellIs"; status ]; Tabs = []
            Buttons = [ dialogCancel
                        dialogButton "xl-dlg-conditional-next" "Next" true (fun fields ->
                            try ready grid context; editing context None (Some fields.["kind"]); false
                            with error -> fail dialog error) ] }
    and editing context current newKind =
        let original = current |> Option.map (fun item -> prop item "rule") |> Option.defaultValue null
        let kind = newKind |> Option.defaultValue (value original "type")
        let known = ruleKinds |> List.exists (fun (key, _) -> key = kind)
        let visual = List.contains kind [ "colorScale"; "dataBar"; "iconSet" ]
        let range = current |> Option.map (fun item -> value item "range") |> Option.defaultValue context.Selection
        let count = grid.GetConditionalFormats(null).Length
        let fields =
            [ dialogText "range" "Applies to (one or more areas)" "A1:A20 C1:C20" range
              if current.IsSome then
                  dialogNumber "priority" "Priority (1 = highest)" (Some "1") (Some (string (max 1 count)))
                      (value (Option.get current) "priority")
              dialogCheckbox "stopIfTrue" "Stop if true" (current |> Option.exists (fun item -> flag item "stopIfTrue")) ]
        let ruleFields = if known then ruleFields kind original else [ dialogStatus "unsupported" "Imported rule" "This rule type is read-only. Range, priority and stop-if-true remain editable." ]
        let all = fields @ ruleFields @ (if visual || not known then [] else formatFields original) @ [ status ]
        let initial = all |> List.map (fun item -> item.Id, item.Initial) |> Map.ofList
        let submit (values: Map<string, string>) =
            try
                ready grid context
                if current.IsNone && grid.GetSelectionRef() <> context.Selection then
                    invalidOp "The selected cells changed while the rule was being created. Reopen the manager."
                let modifications = createObj []
                if changed initial values "range" then setProp modifications "range" (box (values.["range"].Trim()))
                if changed initial values "stopIfTrue" then setProp modifications "stopIfTrue" (box (values.["stopIfTrue"] = "true"))
                if current.IsSome && changed initial values "priority" then
                    setProp modifications "priority" (box (integer "priority" values.["priority"]))
                let ruleChanged =
                    known && (current.IsNone || ((ruleFields @ (if visual then [] else formatFields original))
                                                |> List.exists (fun item -> changed initial values item.Id)))
                if ruleChanged then setProp modifications "rule" (updateRule kind original initial values)
                if current.IsNone then
                    if not known then invalidOp "Imported rule types cannot be created here."
                    grid.AddConditionalFormat(box (values.["range"].Trim()), prop modifications "rule",
                                              createObj [ "stopIfTrue" ==> (values.["stopIfTrue"] = "true") ], null, null) |> ignore
                elif not (missing (prop modifications "rule") && missing (prop modifications "range")
                           && missing (prop modifications "stopIfTrue") && missing (prop modifications "priority")) then
                    grid.UpdateConditionalFormat(prop (Option.get current) "id", modifications)
                listing()
                false
            with error -> fail dialog error
        openDialog dialog {
            TestId = "xl-dlg-conditional-rule"; Title = (if current.IsSome then "Edit conditional rule" else "New conditional rule")
            Subtitle = Some (if known then sprintf "Rule type: %s. Only changed fields are rewritten; imported advanced settings stay intact." kind
                             else "Unsupported imported rule: only its placement and stop-if-true can be edited.")
            Fields = []
            Tabs = [ "Rule", fields @ ruleFields @ [ status ] ] @
                   (if visual || not known then [] else [ "Format", formatFields original @ [ dialogStatus "error-format" "Validation" "" ] ])
            Buttons = [ dialogCancel; dialogButton "xl-dlg-conditional-apply" "Apply" true submit ] }
    listing()

let openNameManager (grid: SpreadsheetPort) (dialog: DialogState) =
    let rec listing () =
        let context = capture grid
        let names = grid.GetDefinedNames()
        let key entry =
            let scope = prop entry "scope"
            (if missing scope then "workbook" else text scope) + "|" + value entry "name"
        let selected (fields: Map<string, string>) =
            match names |> Array.tryFind (fun entry -> key entry = fields.["selected"]) with
            | Some entry -> entry
            | None -> invalidOp "Choose a name from the list."
        let guarded callback fields =
            try ready grid context; callback fields
            with error -> fail dialog error
        let choices = [ "", "Choose a name" ] @
                      [ for name in names ->
                            key name, sprintf "%s · %s%s%s" (value name "name")
                                (if missing (prop name "scope") then "Workbook" else value name "scopeName")
                                (if flag name "isBuiltIn" then " · built-in (read-only)" else "")
                                (if flag name "hidden" then " · hidden in workbook" else "") ]
        openDialog dialog {
            TestId = "xl-dlg-name-manager"; Title = "Name manager"
            Subtitle = Some "Names with identical spelling in different scopes are separate. Built-in names cannot be changed."
            Fields = [ select "selected" "Defined names (name · scope)" choices (if names.Length = 0 then "" else key names.[0])
                       dialogComputed "details" "Reference and comment" (fun fields ->
                           match names |> Array.tryFind (fun item -> key item = fields.["selected"]) with
                           | None -> "No name selected."
                           | Some name -> sprintf "%s%s" (value name "refersTo")
                                              (if value name "comment" = "" then "" else " · " + value name "comment"))
                       status ]; Tabs = []
            Buttons = [ dialogCancel
                        dialogButton "xl-dlg-name-add" "New name" false (guarded (fun _ -> editing context None; false))
                        dialogButton "xl-dlg-name-edit" "Edit name" true (guarded (fun fields -> editing context (Some (selected fields)); false))
                        dialogButton "xl-dlg-name-delete" "Delete name" false (guarded (fun fields ->
                            let entry = selected fields
                            if flag entry "isBuiltIn" then invalidOp "Built-in names cannot be deleted."
                            grid.DeleteDefinedName(box (value entry "name"), prop entry "scope")
                            listing(); false)) ] }
    and editing context current =
        let entry = current |> Option.defaultValue null
        let sheets = grid.GetSheetNames()
        let scopes = [ "", "Workbook" ] @ [ for i in 0 .. sheets.Length - 1 -> string i, sheets.[i] ]
        let scope = if missing (prop entry "scope") then "" else value entry "scope"
        let locked = flag entry "isBuiltIn"
        let fields = [
            if current.IsSome then
                yield dialogStatus "identity" "Name (immutable)" (value entry "name")
                yield dialogStatus "scopeLabel" "Scope (immutable)"
                    (if scope = "" then "Workbook" else sheets.[integer "scope" scope])
            else
                yield dialogText "name" "Name" "SalesTotal" ""
                yield select "scope" "Scope" scopes ""
            if locked then
                yield dialogStatus "refersTo" "Refers to (read-only)" (value entry "refersTo")
                yield dialogStatus "comment" "Comment (read-only)" (value entry "comment")
            else
                yield dialogText "refersTo" "Refers to" "Sheet1!$A$1:$A$20 or a formula" (value entry "refersTo")
                yield dialogTextArea "comment" "Comment" "Optional description" 2 (value entry "comment")
            yield status ]
        let initial = fields |> List.map (fun field -> field.Id, field.Initial) |> Map.ofList
        openDialog dialog {
            TestId = "xl-dlg-name-editor"; Title = (if current.IsSome then "Edit defined name" else "New defined name")
            Subtitle = Some (if locked then "This built-in name is read-only."
                             elif current.IsSome then "Name and scope are immutable. Edit its reference or comment; use New for another identity."
                             else "A name may be workbook-wide or local to one sheet.")
            Fields = fields; Tabs = []
            Buttons =
                if locked then [ dialogCancel ]
                else [ dialogCancel
                       dialogButton "xl-dlg-name-apply" "Apply" true (fun values ->
                           try
                               ready grid context
                               let name = if current.IsSome then value entry "name" else values.["name"].Trim()
                               let scope = if current.IsSome then prop entry "scope"
                                           elif values.["scope"] = "" then null
                                           else box (integer "scope" values.["scope"])
                               if current.IsNone &&
                                  (grid.GetDefinedNames() |> Array.exists (fun existing ->
                                      String.Equals(value existing "name", name, StringComparison.OrdinalIgnoreCase)
                                      && ((missing scope && missing (prop existing "scope"))
                                          || (not (missing scope) && value existing "scope" = text scope)))) then
                                   invalidArg "name" "A name with this scope already exists. Edit the existing name."
                               if current.IsNone || changed initial values "refersTo" || changed initial values "comment" then
                                   grid.SetDefinedName(box name, box (values.["refersTo"].Trim()), scope,
                                                       box values.["comment"])
                               listing()
                               false
                           with error -> fail dialog error) ] }
    listing()
#endif
