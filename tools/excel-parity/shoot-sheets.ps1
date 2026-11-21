param([Parameter(Mandatory = $true)][string]$ListPath, [switch]$KeepZoom)
# Screenshots of Excel's worksheet pane (PrintWindow of the EXCEL7 window, as screenshot-excel.ps1), one per
# line of LIST: workbook path, sheet name, the cell to scroll to the top left, output PNG (tab separated).
# Zoom 100% (-KeepZoom: the zoom the sheet is saved with, which is what a user sees), the selection parked out of sight. Run it in the desktop session (run-vbs-in-session.ps1).
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing
Add-Type -Namespace RbShoot -Name Win32 -MemberDefinition @'
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
  $cb = [RbShoot.Win32+EnumChildProc] {
    param([IntPtr]$h, [IntPtr]$lp)
    $sb = New-Object System.Text.StringBuilder 128
    [RbShoot.Win32]::GetClassName($h, $sb, $sb.Capacity) | Out-Null
    if ($sb.ToString() -eq "EXCEL7") {
      $r = New-Object RbShoot.Win32+RECT
      if ([RbShoot.Win32]::GetWindowRect($h, [ref]$r)) { $found.Add([pscustomobject]@{ Hwnd = $h; Rect = $r; Area = ($r.Right - $r.Left) * ($r.Bottom - $r.Top) }) }
    }
    return $true
  }
  [RbShoot.Win32]::EnumChildWindows($root, $cb, [IntPtr]::Zero) | Out-Null
  $found | Sort-Object -Property Area -Descending | Select-Object -First 1
}

$excel = New-Object -ComObject Excel.Application
try {
  $excel.Visible = $true
  $excel.DisplayAlerts = $false
  $excel.AskToUpdateLinks = $false
  $excel.EnableEvents = $false
  foreach ($line in Get-Content -LiteralPath $ListPath) {
    if ($line.Trim() -eq "") { continue }
    $book, $sheetName, $cell, $out = $line.Split("`t")
    try {
      $wb = $excel.Workbooks.Open($book, $false, $true)
      try {
        try { $excel.WindowState = -4137 } catch { }
        $ws = $wb.Worksheets.Item($sheetName)
        $ws.Activate()
        if (-not $KeepZoom) { $excel.ActiveWindow.Zoom = 100 }
        try { $ws.Range("XFD1048576").Select() } catch { }
        $excel.Goto($ws.Range($cell), $true)
        try { $ws.Range("XFD1048576").Select() } catch { }
        $excel.ActiveWindow.ScrollRow = $ws.Range($cell).Row
        $excel.ActiveWindow.ScrollColumn = $ws.Range($cell).Column
        Start-Sleep -Milliseconds 2000
        $hwnd = [IntPtr]$excel.Hwnd
        [RbShoot.Win32]::ShowWindow($hwnd, 3) | Out-Null
        Start-Sleep -Milliseconds 750
        $pane = Find-Pane $hwnd
        $w = $pane.Rect.Right - $pane.Rect.Left; $h = $pane.Rect.Bottom - $pane.Rect.Top
        $bmp = New-Object System.Drawing.Bitmap $w, $h
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $hdc = $g.GetHdc()
        [RbShoot.Win32]::PrintWindow($pane.Hwnd, $hdc, 2) | Out-Null
        $g.ReleaseHdc($hdc)
        $bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
        $g.Dispose(); $bmp.Dispose()
        Write-Output "ok`t$out`t$w x $h"
      }
      finally { $wb.Close($false) }
    }
    catch { Write-Output "ERROR`t$out`t$($_.Exception.Message)" }
  }
}
finally {
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
Write-Output "DONE"
