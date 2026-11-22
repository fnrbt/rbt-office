# chartsheet-probe.ps1 OUT.tsv PNGDIR [WORKBOOK ...]: Excel 2019's chart sheets - those of the workbooks given and of
# synthetic-chartsheets.xlsx, which it makes first (PNGDIR\synthetic-chartsheets.xlsx: a column chart on chart sheets
# of Letter landscape (Excel's default), portrait, A4, wide margins, a chart sized with the window (zoomToFit), and a
# pie) - each activated in a maximised window: its chart's size and plot area through COM, the window's zoom and the
# pane's size, the page set up; a picture of the pane (PrintWindow, PNGDIR\<book>-<sheet>-view.png) and the chart's
# own (Chart.Export, PNGDIR\<book>-<sheet>.png). Run in the desktop session:
#   PNGDIR=dir tools/excel-parity/probe-on-vm.sh tools/excel-parity/chartsheet-probe.ps1 out.tsv [book.xlsx ...]
# (probe-on-vm.sh passes OUT, the workbooks, then PNGDIR.)
$ErrorActionPreference = "Stop"
$out = $args[0]
$pngDir = $args[$args.Count - 1]
$books = @()
if ($args.Count -gt 2) { $books = $args[1..($args.Count - 2)] }
Add-Type -AssemblyName System.Drawing
Add-Type -Namespace RbChart -Name Win32 -MemberDefinition @'
[StructLayout(LayoutKind.Sequential)]
public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
public delegate bool EnumChildProc(IntPtr hwnd, IntPtr lParam);
[DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);
[DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr hwndParent, EnumChildProc lpEnumFunc, IntPtr lParam);
[DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hwnd, IntPtr hdcBlt, uint nFlags);
'@
function Find-Pane([IntPtr]$root) {
  $found = New-Object System.Collections.Generic.List[object]
  $cb = [RbChart.Win32+EnumChildProc] {
    param([IntPtr]$h, [IntPtr]$lp)
    $sb = New-Object System.Text.StringBuilder 128
    [RbChart.Win32]::GetClassName($h, $sb, $sb.Capacity) | Out-Null
    if ($sb.ToString() -eq "EXCEL7") {
      $r = New-Object RbChart.Win32+RECT
      if ([RbChart.Win32]::GetWindowRect($h, [ref]$r)) { $found.Add([pscustomobject]@{ Hwnd = $h; Rect = $r; Area = ($r.Right - $r.Left) * ($r.Bottom - $r.Top) }) }
    }
    return $true
  }
  [RbChart.Win32]::EnumChildWindows($root, $cb, [IntPtr]::Zero) | Out-Null
  $found | Sort-Object -Property Area -Descending | Select-Object -First 1
}
$lines = New-Object System.Collections.Generic.List[string]
$excel = New-Object -ComObject Excel.Application
try {
  $excel.Visible = $true
  $excel.DisplayAlerts = $false
  $excel.AskToUpdateLinks = $false
  $excel.EnableEvents = $false
  # ---- the fixture
  $fixture = Join-Path $pngDir "synthetic-chartsheets.xlsx"
  $wb = $excel.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $ws.Name = "Data"
  $ws.Range("A1").Value2 = "Quarter"; $ws.Range("B1").Value2 = "Sales"; $ws.Range("C1").Value2 = "Costs"
  $q = @("Q1", "Q2", "Q3", "Q4"); $s = @(120, 135, 160, 150); $c = @(80, 95, 90, 110)
  for ($i = 0; $i -lt 4; $i++) { $ws.Cells.Item($i + 2, 1).Value2 = $q[$i]; $ws.Cells.Item($i + 2, 2).Value2 = [double]$s[$i]; $ws.Cells.Item($i + 2, 3).Value2 = [double]$c[$i] }
  $specs = @(
    @{ Name = "Landscape"; Type = 51; Orient = 0; Paper = 0; Margin = 0; Fit = $false },
    @{ Name = "Portrait"; Type = 51; Orient = 1; Paper = 0; Margin = 0; Fit = $false },
    @{ Name = "A4"; Type = 51; Orient = 2; Paper = 9; Margin = 0; Fit = $false },
    @{ Name = "Margins"; Type = 51; Orient = 2; Paper = 0; Margin = 2; Fit = $false },
    @{ Name = "Fit"; Type = 51; Orient = 2; Paper = 0; Margin = 0; Fit = $true },
    @{ Name = "Pie"; Type = 5; Orient = 0; Paper = 0; Margin = 0; Fit = $false }
  )
  foreach ($spec in $specs) {
    $ch = $wb.Charts.Add([Type]::Missing, $wb.Sheets.Item($wb.Sheets.Count))
    $ch.Name = $spec.Name
    $ch.ChartType = $spec.Type
    if ($spec.Type -eq 5) { $ch.SetSourceData($ws.Range("A1:B5")) } else { $ch.SetSourceData($ws.Range("A1:C5")) }
    $ch.HasTitle = $true
    $ch.ChartTitle.Text = $spec.Name
    if ($spec.Orient -eq 1) { $ch.PageSetup.Orientation = 1 }
    if ($spec.Orient -eq 2) { $ch.PageSetup.Orientation = 2 }
    if ($spec.Paper -ne 0) { $ch.PageSetup.PaperSize = $spec.Paper }
    if ($spec.Margin -ne 0) { $ch.PageSetup.LeftMargin = 144; $ch.PageSetup.RightMargin = 144; $ch.PageSetup.TopMargin = 108; $ch.PageSetup.BottomMargin = 108 }
    if ($spec.Fit) { $ch.Activate(); $excel.ActiveWindow.Zoom = $true; try { $ch.SizeWithWindow = $true } catch { } }
  }
  $wb.Worksheets.Item(1).Activate()
  $wb.SaveAs($fixture, 51)
  $wb.Close($false)
  # ---- each chart sheet
  foreach ($book in @($fixture) + $books) {
    $wb = $excel.Workbooks.Open($book, $false, $true)
    try { $excel.WindowState = -4137 } catch { }
    $base = [System.IO.Path]::GetFileNameWithoutExtension($book)
    foreach ($ch in $wb.Charts) {
      try {
        $ch.Activate()
        Start-Sleep -Milliseconds 1500
        $hwnd = [IntPtr]$excel.Hwnd
        [RbChart.Win32]::ShowWindow($hwnd, 3) | Out-Null
        Start-Sleep -Milliseconds 750
        $pane = Find-Pane $hwnd
        $w = $pane.Rect.Right - $pane.Rect.Left; $h = $pane.Rect.Bottom - $pane.Rect.Top
        $win = $excel.ActiveWindow
        $ps = $ch.PageSetup
        $pa = $ch.PlotArea
        $lines.Add(($base, $ch.Name, "chart", $ch.ChartArea.Width, $ch.ChartArea.Height, $ch.ChartArea.Left, $ch.ChartArea.Top) -join "`t")
        $lines.Add(($base, $ch.Name, "window", $w, $h, $win.Zoom, $win.UsableWidth, $win.UsableHeight, $win.VisibleRange) -join "`t")
        $sw = ""; try { $sw = $ch.SizeWithWindow } catch { }
        $lines.Add(($base, $ch.Name, "page", $ps.Orientation, $ps.PaperSize, $ps.LeftMargin, $ps.RightMargin, $ps.TopMargin, $ps.BottomMargin, $ps.HeaderMargin, $ps.FooterMargin, $ps.ChartSize, "sizeWithWindow $sw") -join "`t")
        $lines.Add(($base, $ch.Name, "plot", $pa.Left, $pa.Top, $pa.Width, $pa.Height, $pa.InsideLeft, $pa.InsideTop, $pa.InsideWidth, $pa.InsideHeight) -join "`t")
        if ($ch.HasTitle) { $lines.Add(($base, $ch.Name, "title", $ch.ChartTitle.Left, $ch.ChartTitle.Top, $ch.ChartTitle.Width, $ch.ChartTitle.Height, $ch.ChartTitle.Font.Size) -join "`t") }
        if ($ch.HasLegend) { $lines.Add(($base, $ch.Name, "legend", $ch.Legend.Left, $ch.Legend.Top, $ch.Legend.Width, $ch.Legend.Height, $ch.Legend.Font.Size) -join "`t") }
        try { $ax = $ch.Axes(2); $lines.Add(($base, $ch.Name, "axis", $ax.MinimumScale, $ax.MaximumScale, $ax.MajorUnit, $ax.TickLabels.Font.Size) -join "`t") } catch { }
        $bmp = New-Object System.Drawing.Bitmap $w, $h
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $hdc = $g.GetHdc()
        [RbChart.Win32]::PrintWindow($pane.Hwnd, $hdc, 2) | Out-Null
        $g.ReleaseHdc($hdc)
        $bmp.Save((Join-Path $pngDir ("$base-" + $ch.Name.Replace(" ", "_") + "-view.png")), [System.Drawing.Imaging.ImageFormat]::Png)
        $g.Dispose(); $bmp.Dispose()
        $ch.Export((Join-Path $pngDir ("$base-" + $ch.Name.Replace(" ", "_") + ".png"))) | Out-Null
      }
      catch { $lines.Add(($base, $ch.Name, "ERROR", $_.Exception.Message) -join "`t") }
    }
    $wb.Close($false)
  }
}
finally {
  [System.IO.File]::WriteAllLines($out, $lines)
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
