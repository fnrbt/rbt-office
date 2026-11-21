# gdi-metrics.ps1 -InputPath x -OutputPath out.tsv : GDI's TEXTMETRIC of fonts at pixel heights (lfHeight = -px).
param([string]$InputPath = "", [string]$OutputPath = "")
$ErrorActionPreference = "Stop"
if ($OutputPath -eq "") { $OutputPath = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "gdi.tsv" }
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class Gdi {
  [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
  public struct TEXTMETRIC { public int tmHeight, tmAscent, tmDescent, tmInternalLeading, tmExternalLeading, tmAveCharWidth, tmMaxCharWidth, tmWeight, tmOverhang, tmDigitizedAspectX, tmDigitizedAspectY; public char tmFirstChar, tmLastChar, tmDefaultChar, tmBreakChar; public byte tmItalic, tmUnderlined, tmStruckOut, tmPitchAndFamily, tmCharSet; }
  [DllImport("gdi32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr CreateFontW(int h, int w, int esc, int ori, int weight, uint italic, uint underline, uint strike, uint charset, uint outPrec, uint clipPrec, uint quality, uint pitch, string face);
  [DllImport("gdi32.dll")] public static extern IntPtr SelectObject(IntPtr dc, IntPtr obj);
  [DllImport("gdi32.dll")] public static extern bool DeleteObject(IntPtr obj);
  [DllImport("gdi32.dll", CharSet = CharSet.Unicode)] public static extern bool GetTextMetricsW(IntPtr dc, out TEXTMETRIC tm);
  [DllImport("gdi32.dll", CharSet = CharSet.Unicode)] public static extern int GetTextFaceW(IntPtr dc, int n, System.Text.StringBuilder face);
  [DllImport("user32.dll")] public static extern IntPtr GetDC(IntPtr wnd);
  public static string Metrics(string face, int px, bool bold, bool italic) {
    IntPtr dc = GetDC(IntPtr.Zero);
    IntPtr f = CreateFontW(-px, 0, 0, 0, bold ? 700 : 400, italic ? 1u : 0u, 0, 0, 1, 0, 0, 0, 0, face);
    IntPtr old = SelectObject(dc, f);
    TEXTMETRIC tm; GetTextMetricsW(dc, out tm);
    var sb = new System.Text.StringBuilder(64); GetTextFaceW(dc, 64, sb);
    SelectObject(dc, old); DeleteObject(f);
    return tm.tmHeight + "\t" + tm.tmAscent + "\t" + tm.tmDescent + "\t" + tm.tmInternalLeading + "\t" + tm.tmExternalLeading + "\t" + tm.tmAveCharWidth + "\t" + sb.ToString();
  }
}
"@
$faces = @("Calibri", "Arial", "Times New Roman", "Cambria", "Verdana", "Tahoma", "Segoe UI", "Courier New", "Georgia", "Calibri Light", "Garamond", "Century Gothic", "Aptos", "Trebuchet MS", "Aptos Narrow", "Arial Narrow", "Consolas", "Candara", "Corbel", "Book Antiqua", "Lucida Sans", "Franklin Gothic Book", "Gill Sans MT", "Palatino Linotype", "Constantia", "Roboto", "Aptos Display", "Myriad Pro", "Segoe UI Black", "MS Sans Serif", "Arial MT", "Small Fonts")
$lines = New-Object System.Collections.Generic.List[string]
foreach ($face in $faces) {
  foreach ($px in 6..60) {
    foreach ($k in 0..3) {
      $lines.Add($face + "`t" + $px + "`t" + $k + "`t" + [Gdi]::Metrics($face, $px, ($k -eq 1 -or $k -eq 3), ($k -ge 2)))
    }
  }
}
[System.IO.File]::WriteAllLines($OutputPath, $lines)
