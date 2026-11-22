#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeUi.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Vb.dll"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../office-wisp/spreadsheet-native/bin/Debug/net10.0/Rbt.OpenType.Core.dll"
#r "../../office-wisp/spreadsheet-native/bin/Debug/net10.0/Rbt.Renderer.dll"
#r "../../office-wisp/spreadsheet-native/bin/Debug/net10.0/StbImageNext.dll"
#r "../../office-wisp/spreadsheet-native/bin/Debug/net10.0/Rbt.OfficeNative.dll"
#r "../../office-wisp/spreadsheet-native/bin/Debug/net10.0/Rbt.SpreadsheetNative.dll"
#load "../../tools/excel-native-runner/Import.fs"

open System.IO
open Rbt.Renderer
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Paint

let path = Path.Combine(__SOURCE_DIRECTORY__, "../test-documents-excel/synthetic-drawing-extras.xlsx")
let workbook = XlsxReader.loadBytesNative (File.ReadAllBytes path) |> Option.get
let context = ExcelNativeRunner.buildWorkbookContext workbook
let model = context.Sheets.["sheet1"].Model
if getCellValue { Row = 0; Col = 0 } model <> Number 42.0 then failwith "native import lost A1"
let viewport = Rbt.SpreadsheetNative.GridRender.viewportOf model 1100.0 720.0
let items = buildViewportPaintItems model viewport |> List.map (fun item -> item.Item)
let pictureRect =
    items |> List.pick (function DrawImage (rect, _) -> Some rect | _ -> None)
let frame =
    Rbt.SpreadsheetNative.GridRender.renderSceneCommands 1100.0 720.0 items
    |> Rbt.Renderer.Scene.render (Rbt.Renderer.Raster.create()) 1100 720
let sampleX = int (pictureRect.X + pictureRect.Width * 0.75)
let sampleY = int (pictureRect.Y + pictureRect.Height * 0.25)
let pixels = frame.Px
let pixel = pixels.Span.[sampleY * frame.W + sampleX]
if (Pixel.r pixel, Pixel.g pixel, Pixel.b pixel) <> (32uy, 96uy, 200uy) then
    failwith "native import did not render the embedded picture's blue content"

// Exercise the native adapter with a deterministic in-memory reader workbook:
// cached and uncached formulas, a second sheet, a name, a table, and axes.
let sourceSheet = workbook.Sheets.[0]
let sourceRow = sourceSheet.Rows.[0]
let sourceCell = sourceRow.Cells.[0]
let formula column expression value =
    { sourceCell with Reference = { sourceCell.Reference with Column = column }
                      Formula = Some expression; Value = value }
let firstRow =
    { sourceRow with Height = Some 19.4; RawHeight = Some 19.4; CustomHeight = Some true
                     Cells = [ sourceCell
                               formula "B" "A1*2" (XlsxReader.Number 777.0)
                               formula "C" "Data!A1+Rate" (XlsxReader.Formula "Data!A1+Rate")
                               formula "D" "SUM(Sales[Amount])" (XlsxReader.Formula "SUM(Sales[Amount])") ] }
let hiddenRow = { sourceRow with Index = 2; Cells = []; Hidden = true }
let lastRow =
    { sourceRow with Index = 1048576; Height = None; RawHeight = None; CustomHeight = None
                     Cells = [ { sourceCell with Reference = { sourceCell.Reference with Row = 1048576 } } ] }
let visibleColumn : XlsxReader.ColumnInfo =
    { Index = 1; MaxIndex = 1; Width = Some 10.0; BestFit = false; Hidden = false; OutlineLevel = 0; Collapsed = false; Style = None }
let hiddenColumn = { visibleColumn with Index = 4; MaxIndex = 4; Hidden = true }
let table : XlsxReader.TableDefinition =
    { Name = "Sales"; Columns = [ "Amount" ]; StartRow = 1; StartCol = 1; EndRow = 2; EndCol = 1
      HeaderRowCount = 1; TotalsRowCount = 0; Totals = []; Style = None; AutoFilter = None }
let first =
    { sourceSheet with Name = "First"; RowSource = XlsxReader.SheetRows.Of [ firstRow; hiddenRow; lastRow ]
                       Columns = [ visibleColumn; hiddenColumn ]; Tables = [ table ] }
let second = { sourceSheet with Name = "Data"; Tables = [] }
let name : XlsxReader.NamedRange =
    { Name = "Rate"; RefersTo = "Data!$A$1"; Scope = None; Comment = None; Hidden = false }
let synthetic = { workbook with Sheets = [ first; second ]; NamedRanges = [ name ] }
let native = ExcelNativeRunner.buildWorkbookContext synthetic
let nativeFirst = native.Sheets.["sheet1"].Model
let value col = getCellValue { Row = 0; Col = col } nativeFirst
match value 1, value 2, value 3 with
| Formula (_, Number 777.0), Formula (_, Number 84.0), Formula (_, Number 0.0) -> ()
| values -> failwithf "native import changed cached, cross-sheet/name or structured-reference values: %A" values
if nativeFirst.RowHeights.[0] <> 25.0 || nativeFirst.RowHeights.[1] <> 0.0
   || nativeFirst.ColumnWidths.[3] <> 0.0
   || nativeFirst.ColumnWidths.[0] <> System.Math.Round nativeFirst.ColumnWidths.[0] then
    failwith "native import lost rounded or hidden axis geometry"
if nativeFirst.RowHeights.Count > 4 then
    failwithf "sparse last row expanded to %d row heights" nativeFirst.RowHeights.Count
printfn "native XLSX drawing regression passed"
