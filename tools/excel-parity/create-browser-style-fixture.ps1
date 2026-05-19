param(
  [Parameter(Mandatory = $true)][string]$InputPath,
  [Parameter(Mandatory = $true)][string]$OutputPath
)

$ErrorActionPreference = "Stop"

$excel = New-Object -ComObject Excel.Application
try {
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $wb = $excel.Workbooks.Add()
  try {
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "BrowserStyles"
    $ws.Columns.Item(1).ColumnWidth = 24
    $ws.Columns.Item(2).ColumnWidth = 22

    $labels = @(
      "solid orange fill",
      "light gray fill",
      "medium gray fill",
      "dark gray fill",
      "navy fill",
      "green fill",
      "pink fill",
      "larger heading font",
      "double underline",
      "diagonal up",
      "diagonal down red",
      "diagonal both blue",
      "plain text"
    )

    for ($i = 0; $i -lt $labels.Count; $i++) {
      $row = $i + 1
      $ws.Cells.Item($row, 1).Value2 = $labels[$i]
      $ws.Rows.Item($row).RowHeight = if ($row -ge 10 -and $row -le 12) { 28 } else { 22 }
    }

    $colors = @(49407, 14277081, 10921638, 5592405, 8388608, 5287936, 13408767)
    for ($row = 1; $row -le 7; $row++) {
      $ws.Cells.Item($row, 2).Interior.Color = $colors[$row - 1]
    }

    $ws.Cells.Item(8, 2).Value2 = "HEADING"
    $ws.Cells.Item(8, 2).Font.Size = 14
    $ws.Cells.Item(9, 2).Value2 = "important"
    $ws.Cells.Item(9, 2).Font.Underline = -4119 # xlUnderlineStyleDouble

    $ws.Cells.Item(10, 2).Value2 = "N/A"
    $ws.Cells.Item(10, 2).Borders.Item(6).LineStyle = 1 # xlDiagonalUp
    $ws.Cells.Item(10, 2).Borders.Item(6).Color = 0
    $ws.Cells.Item(10, 2).Borders.Item(6).Weight = 2

    $ws.Cells.Item(11, 2).Value2 = "cancel"
    $ws.Cells.Item(11, 2).Borders.Item(5).LineStyle = 1 # xlDiagonalDown
    $ws.Cells.Item(11, 2).Borders.Item(5).Color = 192
    $ws.Cells.Item(11, 2).Borders.Item(5).Weight = 2

    $ws.Cells.Item(12, 2).Borders.Item(5).LineStyle = 1
    $ws.Cells.Item(12, 2).Borders.Item(5).Color = 13369344
    $ws.Cells.Item(12, 2).Borders.Item(5).Weight = -4138 # xlMedium
    $ws.Cells.Item(12, 2).Borders.Item(6).LineStyle = 1
    $ws.Cells.Item(12, 2).Borders.Item(6).Color = 13369344
    $ws.Cells.Item(12, 2).Borders.Item(6).Weight = -4138

    $ws.Cells.Item(13, 2).Value2 = "plain text"
    $ws.Range("XFD1048576").Select() | Out-Null

    $xlsxFormat = 51
    $full = [System.IO.Path]::GetFullPath($OutputPath)
    $parent = Split-Path -Parent $full
    if ($parent -and -not (Test-Path -LiteralPath $parent)) {
      New-Item -ItemType Directory -Force -Path $parent | Out-Null
    }
    $wb.SaveAs($full, $xlsxFormat)
  }
  finally {
    $wb.Close($false)
  }
}
finally {
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}

Write-Output ([System.IO.Path]::GetFullPath($OutputPath))
