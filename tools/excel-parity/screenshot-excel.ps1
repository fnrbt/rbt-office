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
[DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
[DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr hWnd);
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
[DllImport("user32.dll")] public static extern bool IsWindow(IntPtr hWnd);
'@

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

    $rect = New-Object RbExcelShot.Win32+RECT
    if (-not [RbExcelShot.Win32]::GetWindowRect($hwnd, [ref]$rect)) {
      throw "GetWindowRect failed"
    }

    $width = [Math]::Max(1, $rect.Right - $rect.Left)
    $height = [Math]::Max(1, $rect.Bottom - $rect.Top)

    $bitmap = New-Object System.Drawing.Bitmap $width, $height
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    try {
      $graphics.CopyFromScreen($rect.Left, $rect.Top, 0, 0, [System.Drawing.Size]::new($width, $height))
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
