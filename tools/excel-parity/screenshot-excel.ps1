param(
  [Parameter(Mandatory = $true)][string]$InputPath,
  [Parameter(Mandatory = $true)][string]$OutputPath
)

# Captures a PNG of the live Excel main window so the row/column headers
# and gridlines are visible — what `print-excel-pdf.vbs` cannot produce.
# Designed to run under an interactive-user scheduled task because Excel COM
# automation fails in non-interactive sessions.

$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type -Namespace RbExcelShot -Name Win32 -MemberDefinition @'
[StructLayout(LayoutKind.Sequential)]
public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
public delegate bool EnumChildProc(IntPtr hwnd, IntPtr lParam);
[DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
[DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr hWnd);
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
[DllImport("user32.dll")] public static extern bool IsWindow(IntPtr hWnd);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);
[DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr hwndParent, EnumChildProc lpEnumFunc, IntPtr lParam);
[DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hwnd, IntPtr hdcBlt, uint nFlags);
'@

# Walk Excel's child window tree and return the screen rect of every visible
# pane whose class name is "EXCEL7" (the worksheet pane class — there's one
# per visible split/scroll region, but the one we want is the largest by
# area). The largest such pane is the main worksheet area whose top-left
# corner is the corner gray box where col/row headers begin.
function Find-WorksheetPane([IntPtr]$rootHwnd) {
  $candidates = New-Object System.Collections.Generic.List[object]
  $callback = [RbExcelShot.Win32+EnumChildProc] {
    param([IntPtr]$h, [IntPtr]$lp)
    $sb = New-Object System.Text.StringBuilder 128
    [RbExcelShot.Win32]::GetClassName($h, $sb, $sb.Capacity) | Out-Null
    if ($sb.ToString() -eq "EXCEL7") {
      $r = New-Object RbExcelShot.Win32+RECT
      if ([RbExcelShot.Win32]::GetWindowRect($h, [ref]$r)) {
        $candidates.Add([pscustomobject]@{
          Hwnd = $h
          Rect = $r
          Area = [Math]::Max(0, $r.Right - $r.Left) * [Math]::Max(0, $r.Bottom - $r.Top)
        })
      }
    }
    return $true
  }
  [RbExcelShot.Win32]::EnumChildWindows($rootHwnd, $callback, [IntPtr]::Zero) | Out-Null
  if ($candidates.Count -eq 0) { return $null }
  $candidates | Sort-Object -Property Area -Descending | Select-Object -First 1
}

$outputFull = [System.IO.Path]::GetFullPath($OutputPath)
$outputParent = Split-Path -Parent $outputFull
if ($outputParent -and -not (Test-Path -LiteralPath $outputParent)) {
  New-Item -ItemType Directory -Force -Path $outputParent | Out-Null
}

$excel = New-Object -ComObject Excel.Application
try {
  $excel.Visible = $true
  $excel.DisplayAlerts = $false
  $excel.AskToUpdateLinks = $false
  $excel.EnableEvents = $false

  $workbook = $excel.Workbooks.Open($InputPath, $false, $true)
  try {
    # Calculation can only be set once a workbook is open. Manual mode avoids
    # surprise recompute cycles that delay the screenshot.
    try { $excel.Calculation = -4135 } catch { }
    try { $excel.WindowState = -4137 } catch { }   # xlMaximized
    # Activate the first visible sheet, freeze the view at A1 so the
    # screenshot composition is reproducible across runs.
    foreach ($sheet in $workbook.Worksheets) {
      if ($sheet.Visible -eq $true -or $sheet.Visible -eq -1) {
        $sheet.Activate()
        $sheet.Range("A1").Select()
        $excel.ActiveWindow.ScrollRow = 1
        $excel.ActiveWindow.ScrollColumn = 1
        break
      }
    }

    # Give Excel a moment to finish painting the worksheet, then promote the
    # main window to the foreground so the screenshot capture sees pixels.
    Start-Sleep -Milliseconds 1500
    $hwnd = [IntPtr]$excel.Hwnd
    if (-not [RbExcelShot.Win32]::IsWindow($hwnd)) {
      throw "Excel main window handle is invalid"
    }
    [RbExcelShot.Win32]::ShowWindow($hwnd, 3) | Out-Null   # SW_MAXIMIZE
    [RbExcelShot.Win32]::BringWindowToTop($hwnd) | Out-Null
    [RbExcelShot.Win32]::SetForegroundWindow($hwnd) | Out-Null
    Start-Sleep -Milliseconds 750

    # Prefer the worksheet pane: it begins exactly at the corner gray box
    # (the top-left of the row/column headers), so the capture aligns with
    # our renderer's (0, 0). Fall back to the full window otherwise.
    $pane = Find-WorksheetPane $hwnd
    $targetHwnd = $hwnd
    $rect = $null
    if ($null -ne $pane) {
      $targetHwnd = $pane.Hwnd
      $rect = $pane.Rect
    } else {
      Write-Warning "EXCEL7 worksheet pane not found; falling back to full Excel window"
      $rect = New-Object RbExcelShot.Win32+RECT
      if (-not [RbExcelShot.Win32]::GetWindowRect($hwnd, [ref]$rect)) {
        throw "GetWindowRect failed and no EXCEL7 pane was found"
      }
    }

    $width = [Math]::Max(1, $rect.Right - $rect.Left)
    $height = [Math]::Max(1, $rect.Bottom - $rect.Top)

    $bitmap = New-Object System.Drawing.Bitmap $width, $height
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    try {
      # CopyFromScreen needs an active desktop session; the wordvm sessions
      # are disconnected, so we use PrintWindow which asks the target window
      # to render itself into our DC and works without a visible desktop.
      # PW_RENDERFULLCONTENT (0x00000002) is required for DirectComposition
      # surfaces — Office 2016+ uses DC for its content rendering.
      $hdc = $graphics.GetHdc()
      try {
        $ok = [RbExcelShot.Win32]::PrintWindow($targetHwnd, $hdc, 2)
        if (-not $ok) {
          throw "PrintWindow returned false (no display session and Excel refused to render)"
        }
      }
      finally {
        $graphics.ReleaseHdc($hdc)
      }
      $bitmap.Save($outputFull, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
      $graphics.Dispose()
      $bitmap.Dispose()
    }
  }
  finally {
    $workbook.Close($false)
  }
}
finally {
  try { $excel.Calculation = -4105 } catch { }   # xlCalculationAutomatic
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}

if (-not (Test-Path -LiteralPath $outputFull)) {
  throw "Screenshot output was not created: $outputFull"
}

Write-Output $outputFull
