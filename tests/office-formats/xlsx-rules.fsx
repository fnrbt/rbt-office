#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
module CF = Rbt.Spreadsheet.ConditionalFormatting

// Conditional formats, data validations and the three sheet states through
// the real XLSX writer and reader: model → OOXML shape → package bytes →
// reader → OOXML shape → model.

let mutable pass = 0
let mutable fail = 0
let clip (text: string) = if text.Length > 400 then text.Substring(0, 400) + "…" else text
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-30s  %s" label (clip detail))
let eq label (got: 'a) (expected: 'a) = ok label (got = expected) (sprintf "expected %A  got %A" expected got)

let at r c = { Row = r; Col = c }
let range r0 c0 r1 c1 : CellRange = { From = at r0 c0; To = at r1 c1 }
let wholeColumn c : CellRange = { From = at 0 c; To = at 1048575 c }
let cfvo kind value : CfThreshold = { Kind = kind; Value = value }
let rule ranges r style priority stop : ConditionalFormatRange =
    { Ranges = ranges; Format = { Rule = r; Style = style }; Priority = priority; StopIfTrue = stop }

// ---- the two small mappings a host needs (see spreadsheet-wisp) -------------
let toWriter (dxfs: ResizeArray<XlsxReader.Dxf>) (formats: ConditionalFormatRange list) : XlsxReader.ConditionalFormat list =
    let dxfIndex (dxf: XlsxReader.Dxf) =
        match dxfs.IndexOf dxf with
        | -1 -> dxfs.Add dxf; dxfs.Count - 1
        | index -> index
    CF.toOoxml formats
    |> List.map (fun cf ->
        ({ Range = cf.Sqref
           Rules =
            cf.Rules |> List.map (fun r ->
                ({ Type = r.Type; Operator = r.Operator; Formulas = r.Formulas
                   DxfId = r.Style |> Option.map (fun s ->
                       dxfIndex { FontColor = s.FontColor |> Option.map (fun c -> c.TrimStart '#'); FillColor = s.FillColor |> Option.map (fun c -> c.TrimStart '#')
                                  Bold = (if s.Bold then Some true else None); Italic = (if s.Italic then Some true else None) })
                   ColorScale = r.ColorScale |> List.map (fun (k, v, c) -> ({ CfvoType = k; CfvoValue = v; Color = c } : XlsxReader.CfColorStop))
                   DataBarColor = r.DataBarColor; DataBarCfvo = r.DataBarCfvo; Text = r.Text
                   IconSet = r.IconSet; IconSetCfvo = r.IconSetCfvo; IconSetReverse = r.IconSetReverse
                   Rank = r.Rank; Percent = r.Percent; Bottom = r.Bottom
                   AboveAverage = r.AboveAverage; EqualAverage = r.EqualAverage; StdDev = r.StdDev
                   Priority = r.Priority; StopIfTrue = r.StopIfTrue } : XlsxReader.ConditionalFormatRule)) } : XlsxReader.ConditionalFormat))

let fromReader (wb: XlsxReader.Workbook) (sheet: XlsxReader.Worksheet) : ConditionalFormatRange list =
    sheet.ConditionalFormats
    |> List.map (fun cf ->
        ({ Sqref = cf.Range
           Rules =
            cf.Rules |> List.map (fun r ->
                let dxf = r.DxfId |> Option.bind (fun id -> if id >= 0 && id < wb.Dxfs.Length then Some wb.Dxfs.[id] else None)
                ({ Type = r.Type; Operator = r.Operator; Formulas = r.Formulas; Text = r.Text
                   Rank = r.Rank; Percent = r.Percent; Bottom = r.Bottom
                   AboveAverage = r.AboveAverage; EqualAverage = r.EqualAverage; StdDev = r.StdDev
                   ColorScale = r.ColorScale |> List.map (fun s -> s.CfvoType, s.CfvoValue, s.Color)
                   DataBarColor = r.DataBarColor; DataBarCfvo = r.DataBarCfvo
                   IconSet = r.IconSet; IconSetCfvo = r.IconSetCfvo; IconSetReverse = r.IconSetReverse
                   Priority = r.Priority; StopIfTrue = r.StopIfTrue
                   Style = dxf |> Option.map (fun d -> ({ FontColor = d.FontColor; FillColor = d.FillColor; Bold = (d.Bold = Some true); Italic = (d.Italic = Some true) } : CF.OoxmlStyle)) } : CF.OoxmlRule)) } : CF.OoxmlConditionalFormat))
    |> CF.ofOoxml

let validationsOut (validations: Map<CellAddress, DataValidation>) : XlsxWriter.XlsxDataValidation list =
    CF.validationsToOoxml validations
    |> List.map (fun (sqref, dv) ->
        ({ Sqref = sqref; ValidationType = dv.Type; Operator = dv.Operator; Formula1 = dv.Formula1; Formula2 = dv.Formula2
           AllowBlank = dv.AllowBlank; ShowDropDown = true; ErrorMessage = dv.Error; InputMessage = dv.Prompt
           ErrorTitle = dv.ErrorTitle; PromptTitle = dv.PromptTitle; ErrorStyle = dv.ErrorStyle
           ShowErrorMessage = dv.ShowErrorMessage; ShowInputMessage = dv.ShowInputMessage } : XlsxWriter.XlsxDataValidation))

let validationsIn (sheet: XlsxReader.Worksheet) : Map<CellAddress, DataValidation> =
    let mutable result = Map.empty
    for dv in sheet.DataValidations do
        let ooxml : CF.OoxmlValidation =
            { Type = dv.ValidationType; Operator = dv.Operator; Formula1 = dv.Formula1; Formula2 = dv.Formula2
              AllowBlank = dv.AllowBlank; ShowErrorMessage = dv.ShowErrorMessage; ShowInputMessage = dv.ShowInputMessage
              ErrorStyle = dv.ErrorStyle; ErrorTitle = dv.ErrorTitle; Error = dv.ErrorMessage
              PromptTitle = dv.PromptTitle; Prompt = dv.Prompt }
        for area in CF.parseSqref dv.Range do
            match CF.validationRuleOfOoxml ooxml with
            | Some rule ->
                let validation = CF.validationWithOoxml area.From ooxml rule
                for r in area.From.Row .. area.To.Row do
                    for c in area.From.Col .. area.To.Col do
                        result <- Map.add (at r c) validation result
            | None -> ()
    result

let blankSheet name state cfs dvs : XlsxWriter.XlsxSheetOut =
    { Name = name; State = state
      Cells = [ { Row = 0; Col = 0; Value = XlsxWriter.XlsxNumber 1.0; Format = XlsxWriter.defaultXlsxCellFormat; Comment = None } ]
      MergedCells = []; ColumnWidths = []; RowHeights = []; FrozenRows = 0; FrozenCols = 0; ShowGridLines = true; Tables = []
      DataValidations = dvs; ConditionalFormats = cfs; Hyperlinks = []
      PageMargins = None; PageSetup = None; PrintOptions = None; HeaderFooter = None
      RowBreaks = []; ColBreaks = []; PrintArea = None; PrintTitleRows = None; PrintTitleCols = None }

// ---- the rules ---------------------------------------------------------------
let fill c = { defaultFormat with BgColor = Some c }
let colA = [ range 0 0 9 0 ]
let lo, hi = cfvo CfvoMin None, cfvo CfvoMax None
let rules =
    [ rule colA (CellIs (OpGreaterOrEqual, "10", None)) { defaultFormat with BgColor = Some "#FFC7CE"; TextColor = Some "#9C0006"; Bold = true } 1 false
      rule colA (CellIs (OpLessOrEqual, "B1", None)) (fill "#C6EFCE") 2 false
      rule colA (CellIs (OpNotEqual, "5", None)) (fill "#FFEB9C") 3 true
      rule colA (CellIs (OpNotBetween, "1", Some "9")) { defaultFormat with Italic = true; TextColor = Some "#0000FF" } 4 false
      rule [ wholeColumn 2; range 0 4 3 5 ] (Expression "AND($A1>0,MOD(ROW(),2)=0)") (fill "#DDEBF7") 5 false
      rule colA (TopBottom (3, false, false)) (fill "#FFC7CE") 6 false
      rule colA (TopBottom (25, true, true)) (fill "#FFC7CE") 7 false
      rule colA (AboveAverage (true, false, 0)) (fill "#C6EFCE") 8 false
      rule colA (AboveAverage (false, true, 1)) (fill "#C6EFCE") 9 false
      rule colA DuplicateValues (fill "#FFC7CE") 10 false
      rule colA UniqueValues (fill "#FFC7CE") 11 false
      rule colA (TextRule (TextContains, "a\"b")) (fill "#FFC7CE") 12 false
      rule colA (TextRule (TextEndsWith, "z")) (fill "#FFC7CE") 13 false
      rule colA (Blanks false) (fill "#FFC7CE") 14 false
      rule colA (Errors true) (fill "#FFC7CE") 15 false
      rule colA (ColorScale [ lo, "#F8696B"; hi, "#63BE7B" ]) defaultFormat 16 false
      rule colA (ColorScale [ lo, "#F8696B"; cfvo CfvoPercentile (Some "50"), "#FFEB84"; cfvo CfvoNumber (Some "90"), "#63BE7B" ]) defaultFormat 17 false
      rule colA (DataBar (lo, hi, "#638EC6")) defaultFormat 18 false
      rule colA (IconSet ("3TrafficLights1", [ cfvo CfvoPercent (Some "0"); cfvo CfvoPercent (Some "33"); cfvo CfvoPercent (Some "67") ], false)) defaultFormat 19 false
      rule colA (IconSet ("5Arrows", [ for p in [ "0"; "20"; "40"; "60"; "80" ] -> cfvo CfvoPercent (Some p) ], true)) defaultFormat 20 false ]

let dv r = validationOf r
let validations =
    Map.ofList
        [ for r in 0 .. 4 -> at r 1, { dv (Comparison (WholeKind, OpNotBetween, "1", Some "5")) with Anchor = Some (at 0 1); ErrorTitle = Some "Range"; ErrorMessage = Some "Not 1 to 5"; Style = WarningStyle }
          yield at 0 2, { dv (Comparison (DecimalKind, OpNotEqual, "0", None)) with Anchor = Some (at 0 2); Style = InformationStyle }
          yield at 0 3, { dv (CustomFormula "ISNUMBER(D1)") with Anchor = Some (at 0 3); InputTitle = Some "Hint"; InputMessage = Some "A number" }
          yield at 0 4, { dv (Comparison (TimeKind, OpGreaterThan, "0.5", None)) with Anchor = Some (at 0 4) }
          yield at 0 5, { dv (Comparison (TextLengthKind, OpLessThan, "8", None)) with Anchor = Some (at 0 5); ShowError = false; AllowBlank = false }
          yield at 0 6, { dv (Comparison (DateKind, OpGreaterOrEqual, "TODAY()", None)) with Anchor = Some (at 0 6) }
          yield at 0 7, { dv (ListValues [ "Yes"; "No" ]) with Anchor = Some (at 0 7) }
          yield at 0 8, { dv (WholeNumber (Some 1.0, Some 10.0)) with Anchor = Some (at 0 8) } ]

let dxfs = ResizeArray<XlsxReader.Dxf>()
let sheets =
    [ blankSheet "Rules" XlsxReader.SheetVisible (toWriter dxfs rules) (validationsOut validations)
      blankSheet "Tucked" XlsxReader.SheetHidden [] []
      blankSheet "Deep" XlsxReader.SheetVeryHidden [] [] ]
let out : XlsxWriter.XlsxWorkbookOut =
    { Sheets = sheets; ActiveSheetIndex = 0; DefinedNames = []; VbaProjectSource = None; Dxfs = List.ofSeq dxfs }
let bytes = XlsxWriter.writeWorkbookPackageBytes out

match XlsxReader.loadBytesNative bytes with
| None -> ok "package-reloads" false "reader returned None"
| Some wb ->
    // 11. All three sheet states survive.
    eq "sheet-states" (wb.Sheets |> List.map (fun s -> s.Name, s.State))
                      [ "Rules", XlsxReader.SheetVisible; "Tucked", XlsxReader.SheetHidden; "Deep", XlsxReader.SheetVeryHidden ]
    let sheet = wb.Sheets.Head
    // 9. Every rule kind comes back as the same model rule, range and style.
    let back = fromReader wb sheet |> List.sortBy (fun r -> r.Priority)
    eq "cf-count" back.Length rules.Length
    for (a, b) in List.zip back rules do
        eq (sprintf "cf-%02d-%s" b.Priority (match b.Format.Rule with CellIs (op, _, _) -> sprintf "%A" op | r -> (sprintf "%A" r).Split(' ', '(').[0]))
           (a.Ranges, a.Format.Rule, a.Priority, a.StopIfTrue, a.Format.Style)
           (b.Ranges, b.Format.Rule, b.Priority, b.StopIfTrue, b.Format.Style)
    eq "cf-whole-column-kept" (sheet.ConditionalFormats |> List.exists (fun cf -> cf.Range = "C:C E1:F4")) true
    eq "cf-dxfs-shared" (wb.Dxfs.Length < rules.Length && wb.Dxfs.Length >= 5) true
    // The same rules evaluate the same after the trip.
    let numbers = [ for r in 0 .. 9 -> at r 0, { Value = CellValue.Number (float (r + 1) * 2.0); Format = defaultFormat; Comment = None } ]
    let before = { createEmpty 50 10 with Cells = Map.ofList numbers; ConditionalFormats = rules }
    let after = { before with ConditionalFormats = back }
    let effects (m: SpreadsheetModel) = [ for r in 0 .. 9 -> CF.effectAt (CF.createContext evaluateFormulaAt m) (at r 0) ]
    eq "cf-same-effects" (effects after) (effects before)
    // 10. Validations: kinds, operators, formulas, alert texts and styles.
    let dvBack = validationsIn sheet
    eq "dv-count" (Map.count dvBack) (Map.count validations)
    for KeyValue(addr, expected) in validations do
        eq (sprintf "dv-%s" (formatCellRef addr.Row addr.Col)) (Map.tryFind addr dvBack) (Some expected)
    eq "dv-grouped" (sheet.DataValidations |> List.exists (fun d -> d.Range = "B1:B5" && d.Operator = Some "notBetween" && d.ErrorStyle = Some "warning")) true
    eq "dv-custom"  (sheet.DataValidations |> List.exists (fun d -> d.ValidationType = "custom" && d.Formula1 = Some "ISNUMBER(D1)" && d.Prompt = Some "A number" && d.PromptTitle = Some "Hint")) true
    eq "dv-time"    (sheet.DataValidations |> List.exists (fun d -> d.ValidationType = "time" && d.Operator = Some "greaterThan")) true

printfn ""
printfn "XLSX rules round trip: %d passed, %d failed" pass fail
if fail > 0 then exit 1
