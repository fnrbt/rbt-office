$ErrorActionPreference = 'Stop'
Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;
public class GdiProbe2 {
  [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
  public struct LOGFONT {
    public int lfHeight, lfWidth, lfEscapement, lfOrientation, lfWeight;
    public byte lfItalic, lfUnderline, lfStrikeOut, lfCharSet, lfOutPrecision, lfClipPrecision, lfQuality, lfPitchAndFamily;
    [MarshalAs(UnmanagedType.ByValTStr, SizeConst=32)] public string lfFaceName;
  }
  public struct SIZE { public int cx, cy; }
  [DllImport("gdi32.dll")] public static extern IntPtr CreateCompatibleDC(IntPtr hdc);
  [DllImport("gdi32.dll", CharSet=CharSet.Unicode, EntryPoint="CreateFontIndirectW")] public static extern IntPtr CreateFontIndirect(ref LOGFONT lf);
  [DllImport("gdi32.dll")] public static extern IntPtr SelectObject(IntPtr hdc, IntPtr h);
  [DllImport("gdi32.dll")] public static extern bool DeleteObject(IntPtr h);
  [DllImport("gdi32.dll")] public static extern bool DeleteDC(IntPtr hdc);
  [DllImport("gdi32.dll", CharSet=CharSet.Unicode)] public static extern int GetTextFaceW(IntPtr hdc, int n, StringBuilder sb);
  [DllImport("gdi32.dll", CharSet=CharSet.Unicode)] public static extern bool GetTextExtentPoint32W(IntPtr hdc, string s, int n, out SIZE sz);
}
"@
function Probe([string]$face, [int]$ppem, [int]$italic) {
  $lf = New-Object GdiProbe2+LOGFONT
  $lf.lfHeight = -$ppem; $lf.lfWeight = 400
  $lf.lfItalic = if ($italic -ne 0) { 1 } else { 0 }
  $lf.lfCharSet = 1; $lf.lfQuality = 0; $lf.lfPitchAndFamily = 0; $lf.lfFaceName = $face
  $font = [GdiProbe2]::CreateFontIndirect([ref]$lf)
  $dc = [GdiProbe2]::CreateCompatibleDC([IntPtr]::Zero)
  $old = [GdiProbe2]::SelectObject($dc, $font)
  $sb = New-Object System.Text.StringBuilder 64
  [void][GdiProbe2]::GetTextFaceW($dc, 64, $sb)
  $sel = $sb.ToString()
  $size = New-Object GdiProbe2+SIZE
  $texts = @()
  foreach ($n in 1..12) { $texts += ('1' * $n); $texts += ('8' * $n) }
  $texts += '1234567.89%','1111111111.00%','1.00%','11.00%','111111.00%','1111111.00%','11111111.00%','111111.00%','.','%','#','$','(',')','-',','
  foreach ($t in $texts) {
    [void][GdiProbe2]::GetTextExtentPoint32W($dc, $t, $t.Length, [ref]$size)
    "$face`t$ppem`t$italic`t$sel`textent`t$t`t$($size.cx)"
  }
  [void][GdiProbe2]::SelectObject($dc, $old)
  [void][GdiProbe2]::DeleteObject($font)
  [void][GdiProbe2]::DeleteDC($dc)
}
"face`tppem`titalic`tselected`tkind`ttext`tcx"
Probe 'Verdana' 81 0
Probe 'Verdana' 108 0
Probe 'Verdana' 109 0
Probe 'Verdana' 120 0
Probe 'Tahoma' 87 0
Probe 'Arial' 80 0
Probe 'Georgia' 113 0
Probe 'Georgia' 113 1
