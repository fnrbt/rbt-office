module Rbt.Spreadsheet.ChartAuthoring

open System
open Rbt.Spreadsheet.Model

/// The four classic chart layouts the grid can author. The source is always a
/// single rectangular block; columns (not rows) become series.
type StandardChartKind = Column | Bar | Line | Pie

let private escaped (text: string) =
    text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&apos;")

let private pointCache (isNumber: bool) (values: (string * float option) array) =
    let kind = if isNumber then "num" else "str"
    let result = Text.StringBuilder()
    result.Append("<c:").Append(kind).Append("Cache>") |> ignore
    if isNumber then result.Append("<c:formatCode>General</c:formatCode>") |> ignore
    result.Append("<c:ptCount val=\"").Append(values.Length).Append("\"/>") |> ignore
    for index = 0 to values.Length - 1 do
        let text, number = values.[index]
        if isNumber then
            match number with
            | Some value when not (Double.IsNaN value || Double.IsInfinity value) ->
                result.Append("<c:pt idx=\"").Append(index).Append("\"><c:v>")
                      .Append(value.ToString("R", Globalization.CultureInfo.InvariantCulture))
                      .Append("</c:v></c:pt>") |> ignore
            | _ -> ()
        else
            result.Append("<c:pt idx=\"").Append(index).Append("\"><c:v>")
                  .Append(escaped text).Append("</c:v></c:pt>") |> ignore
    result.Append("</c:").Append(kind).Append("Cache>").ToString()

let private rangeRef (sheet: string) r0 c0 r1 c1 =
    let cell r c = "$" + columnToLetter c + "$" + string (r + 1)
    let prefix = "'" + sheet.Replace("'", "''") + "'!"
    prefix + cell r0 c0 + (if r0 = r1 && c0 = c1 then "" else ":" + cell r1 c1)

let private refXml number formula values =
    let kind = if number then "num" else "str"
    sprintf "<c:%sRef><c:f>%s</c:f>%s</c:%sRef>" kind (escaped formula) (pointCache number values) kind

let private richTitle (text: string) =
    "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang=\"en-US\"/>"
    + "<a:t>" + escaped text + "</a:t></a:r></a:p></c:rich></c:tx><c:layout/></c:title>"

/// Create a chart part with cached data for interoperability and cell formulas
/// for the shared Charts renderer. `read` gives displayed text and numeric value
/// at a 0-based cell; blanks remain holes in numeric caches.
let standardChartXml (kind: StandardChartKind) (title: string) (sheetName: string)
    (r0: int) (c0: int) (r1: int) (c1: int) (headers: bool) (categories: bool)
    (read: int -> int -> string * float option) : string =
    let firstRow = r0 + (if headers then 1 else 0)
    let firstCol = c0 + (if categories then 1 else 0)
    let count = r1 - firstRow + 1
    let series = c1 - firstCol + 1
    let labels =
        Array.init count (fun index ->
            if categories then read (firstRow + index) c0
            else string (index + 1), None)
    let category =
        if categories then
            "<c:cat>" + refXml false (rangeRef sheetName firstRow c0 r1 c0) labels + "</c:cat>"
        else "<c:cat>" + (pointCache false labels).Replace("strCache", "strLit") + "</c:cat>"
    let seriesXml =
        [ for index in 0 .. series - 1 do
            let column = firstCol + index
            let values = Array.init count (fun at -> read (firstRow + at) column)
            let name =
                if headers then
                    let text = fst (read r0 column)
                    "<c:tx>" + refXml false (rangeRef sheetName r0 column r0 column) [| text, None |] + "</c:tx>"
                else "<c:tx><c:v>" + escaped (sprintf "Series %d" (index + 1)) + "</c:v></c:tx>"
            yield sprintf "<c:ser><c:idx val=\"%d\"/><c:order val=\"%d\"/>%s%s<c:val>%s</c:val></c:ser>"
                index index name category (refXml true (rangeRef sheetName firstRow column r1 column) values) ]
        |> String.concat ""
    let chartKind, opening =
        match kind with
        | Column -> "barChart", "<c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>"
        | Bar -> "barChart", "<c:barDir val=\"bar\"/><c:grouping val=\"clustered\"/>"
        | Line -> "lineChart", "<c:grouping val=\"standard\"/>"
        | Pie -> "pieChart", "<c:varyColors val=\"1\"/>"
    // CT_LineChart places marker after its series, before its axis ids.
    let afterSeries = if kind = Line then "<c:marker val=\"1\"/>" else ""
    let axes = if kind = Pie then "" else "<c:axId val=\"1048576\"/><c:axId val=\"1048577\"/>"
    let categoryAxis, valueAxis =
        if kind = Pie then "", ""
        else
            let catPosition, valPosition = if kind = Bar then "l", "b" else "b", "l"
            sprintf "<c:catAx><c:axId val=\"1048576\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:axPos val=\"%s\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"1048577\"/><c:crosses val=\"autoZero\"/></c:catAx>" catPosition,
            sprintf "<c:valAx><c:axId val=\"1048577\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:axPos val=\"%s\"/><c:majorGridlines/><c:numFmt formatCode=\"General\" sourceLinked=\"1\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"1048576\"/><c:crosses val=\"autoZero\"/></c:valAx>" valPosition
    let legend = if series < 2 || kind = Pie then "" else "<c:legend><c:legendPos val=\"r\"/><c:layout/></c:legend>"
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
    + "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"
    + "<c:chart>" + (if title.Trim() = "" then "<c:autoTitleDeleted val=\"1\"/>" else richTitle title)
    + "<c:plotArea><c:layout/><c:" + chartKind + ">" + opening + seriesXml + afterSeries + axes + "</c:" + chartKind + ">"
    + categoryAxis + valueAxis + "</c:plotArea>" + legend + "<c:plotVisOnly val=\"1\"/></c:chart></c:chartSpace>"

/// An ordinary one-cell anchor: moves with its top-left cell, keeps its size.
let standardChartAnchorXml (id: int) (relationshipId: string) (name: string) (row: int) (col: int) (width: float) (height: float) =
    let cx = int64 (Math.Round(width * 9525.0))
    let cy = int64 (Math.Round(height * 9525.0))
    sprintf "<xdr:oneCellAnchor xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><xdr:from><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx=\"%d\" cy=\"%d\"/><xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"%d\" name=\"%s\"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"%d\" cy=\"%d\"/></xdr:xfrm><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart r:id=\"%s\"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>"
        col row cx cy id (escaped name) cx cy (escaped relationshipId)

/// Update the saved caches from the cells the chart currently refers to.
/// The grid always reads live cells; caches are for applications showing an
/// unopened workbook, and must not describe the pre-edit values.
let refreshCaches (valuesOf: string -> (string * float option) array option) (xml: string) =
    let container = Text.RegularExpressions.Regex("<c:(numRef|strRef)>[\\s\\S]*?</c:\\1>")
    let formula = Text.RegularExpressions.Regex("<c:f>([^<]*)</c:f>")
    let cache = Text.RegularExpressions.Regex("<c:(?:numCache|strCache)>[\\s\\S]*?</c:(?:numCache|strCache)>")
    container.Replace(xml, Text.RegularExpressions.MatchEvaluator(fun found ->
        let reference = found.Value
        let matched = formula.Match reference
        if not matched.Success then reference
        else
            let raw = matched.Groups.[1].Value
            let value = raw.Replace("&lt;", "<").Replace("&gt;", ">").Replace("&quot;", "\"").Replace("&apos;", "'").Replace("&amp;", "&")
            let current = pointCache (found.Groups.[1].Value = "numRef") (valuesOf value |> Option.defaultValue [||])
            if cache.IsMatch reference then cache.Replace(reference, current, 1)
            else reference.Replace(matched.Value, matched.Value + current)))

/// The data of a series of a chart a macro made: from cells (a formula such as 'Sheet1'!$B$2:$B$9 with
/// their values) or literal (an array assigned to Values / XValues).
type MacroSeriesData = { Formula: string option; Points: (string * float option) array }

/// A series of a macro's chart: its name, its x values or categories, its values.
type MacroSeries = { SeriesName: string option; XData: MacroSeriesData option; YData: MacroSeriesData }

let macroSeries (name: string option) (x: MacroSeriesData option) (y: MacroSeriesData) = { SeriesName = name; XData = x; YData = y }

/// A chart a macro made (Shapes.AddChart, ChartObjects.Add): Excel's chart type (XlChartType), a title,
/// a legend, its series.
type MacroChart = { ChartType: int; Title: string option; HasLegend: bool; Series: MacroSeries list }

/// The chart part of a macro's chart, for the shared renderer (and a save): scatter, line, column, bar,
/// area and pie charts; another chart type is drawn as a clustered column chart.
let macroChartXml (chart: MacroChart) : string =
    let data (number: bool) (d: MacroSeriesData) =
        match d.Formula with
        | Some formula -> refXml number formula d.Points
        | None ->
            let cached = pointCache number d.Points
            if number then cached.Replace("numCache>", "numLit>") else cached.Replace("strCache>", "strLit>")
    let numbers (d: MacroSeriesData) = d.Points |> Array.forall (fun (_, n) -> n.IsSome)
    let scatter = chart.ChartType = -4169 || (chart.ChartType >= 72 && chart.ChartType <= 75)
    let kind, opening, markers =
        match chart.ChartType with
        | t when scatter ->
            let style = if t = 72 || t = 73 then "smoothMarker" elif t = -4169 then "lineMarker" else "lineMarker"
            "scatterChart", sprintf "<c:scatterStyle val=\"%s\"/><c:varyColors val=\"0\"/>" style, ""
        | 4 | 65 | 63 | 64 | 66 | 67 -> "lineChart", "<c:grouping val=\"standard\"/><c:varyColors val=\"0\"/>", (if chart.ChartType = 4 then "<c:marker val=\"1\"/>" else "<c:marker val=\"1\"/>")
        | 1 | 76 | 77 -> "areaChart", "<c:grouping val=\"standard\"/><c:varyColors val=\"0\"/>", ""
        | 5 | -4102 | 69 | 70 | -4120 -> "pieChart", "<c:varyColors val=\"1\"/>", ""
        | 57 | 58 | 59 -> "barChart", "<c:barDir val=\"bar\"/><c:grouping val=\"clustered\"/><c:varyColors val=\"0\"/>", ""
        | _ -> "barChart", "<c:barDir val=\"col\"/><c:grouping val=\"clustered\"/><c:varyColors val=\"0\"/>", ""
    let series =
        chart.Series |> List.mapi (fun index s ->
            let name =
                match s.SeriesName with
                | Some text -> "<c:tx><c:v>" + escaped text + "</c:v></c:tx>"
                | None -> "<c:tx><c:v>" + escaped (sprintf "Series%d" (index + 1)) + "</c:v></c:tx>"
            let look =
                if not scatter then ""
                elif chart.ChartType = -4169 then "<c:spPr><a:ln w=\"25400\"><a:noFill/></a:ln></c:spPr>"
                elif chart.ChartType = 75 || chart.ChartType = 73 then "<c:marker><c:symbol val=\"none\"/></c:marker>"
                else ""
            let x, y =
                if scatter then
                    (match s.XData with Some d -> "<c:xVal>" + data (numbers d) d + "</c:xVal>" | None -> ""),
                    "<c:yVal>" + data true s.YData + "</c:yVal>"
                else
                    (match s.XData with Some d -> "<c:cat>" + data (numbers d) d + "</c:cat>" | None -> ""),
                    "<c:val>" + data true s.YData + "</c:val>"
            let smooth = if scatter then sprintf "<c:smooth val=\"%d\"/>" (if chart.ChartType = 72 || chart.ChartType = 73 then 1 else 0) else ""
            sprintf "<c:ser><c:idx val=\"%d\"/><c:order val=\"%d\"/>%s%s%s%s%s</c:ser>" index index name look x y smooth)
        |> String.concat ""
    let pie = kind = "pieChart"
    let axes = if pie then "" else "<c:axId val=\"1048576\"/><c:axId val=\"1048577\"/>"
    let axisXml =
        if pie then ""
        else
            let firstAxis =
                if scatter then "<c:valAx><c:axId val=\"1048576\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:axPos val=\"b\"/><c:numFmt formatCode=\"General\" sourceLinked=\"1\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"1048577\"/><c:crosses val=\"autoZero\"/></c:valAx>"
                else sprintf "<c:catAx><c:axId val=\"1048576\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:axPos val=\"%s\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"1048577\"/><c:crosses val=\"autoZero\"/></c:catAx>" (if kind = "barChart" && opening.Contains "\"bar\"" then "l" else "b")
            firstAxis + sprintf "<c:valAx><c:axId val=\"1048577\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:axPos val=\"%s\"/><c:majorGridlines/><c:numFmt formatCode=\"General\" sourceLinked=\"1\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"1048576\"/><c:crosses val=\"autoZero\"/></c:valAx>" (if kind = "barChart" && opening.Contains "\"bar\"" then "b" else "l")
    let title =
        match chart.Title with
        | Some text when text.Trim() <> "" -> richTitle text
        | _ -> "<c:autoTitleDeleted val=\"1\"/>"
    let legend = if chart.HasLegend then "<c:legend><c:legendPos val=\"r\"/><c:layout/></c:legend>" else ""
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
    + "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"
    + "<c:chart>" + title
    + "<c:plotArea><c:layout/><c:" + kind + ">" + opening + series + markers + axes + "</c:" + kind + ">"
    + axisXml + "</c:plotArea>" + legend + "<c:plotVisOnly val=\"1\"/></c:chart></c:chartSpace>"
