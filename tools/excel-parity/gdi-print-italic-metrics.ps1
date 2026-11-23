# gdi-print-italic-metrics.ps1 - the GDI printer measurements for the printed italic reserve
# (docs/excel-parity-tracker.md, "Print italic reserve"). Run in Excel's desktop session with run-vbs-in-session.ps1
# (PowerShell 5): CreateDC on the default printer, then for each face/size/style record GetTextMetricsW (ascent,
# descent, internal leading, tmOverhang), GetTextExtentPoint32W vs the sum of GetCharABCWidthsW A+B+C, GetCharWidth32W,
# GetGlyphOutlineW, and raw OS/2/head/hhea/post font data. It loads the Aptos Narrow files copied beside it privately
# and measures the same faces with synthesized italic where a real italic face is missing.
# Output is UTF-8 TSV; each run row also carries extentEx and the first A / last C side bearings.
param([string]$Out = "")
$ErrorActionPreference = "Stop"
if ([string]::IsNullOrWhiteSpace($Out)) { $Out = Join-Path $PSScriptRoot "gdi-metrics.tsv" }
$src = @'
using System;
using System.Text;
using System.Runtime.InteropServices;

public static class GdiProbe {
  [DllImport("winspool.drv", CharSet=CharSet.Unicode, SetLastError=true)]
  public static extern bool GetDefaultPrinter(StringBuilder pszBuffer, ref int pcchBuffer);
  [DllImport("gdi32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
  public static extern IntPtr CreateDCW(string pwszDriver, string pwszDevice, string pszPort, IntPtr pdm);
  [DllImport("gdi32.dll", SetLastError=true)]
  public static extern bool DeleteDC(IntPtr hdc);
  [DllImport("user32.dll", SetLastError=true)]
  public static extern IntPtr GetDC(IntPtr hwnd);
  [DllImport("user32.dll", SetLastError=true)]
  public static extern int ReleaseDC(IntPtr hwnd, IntPtr hdc);
  [DllImport("gdi32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
  public static extern IntPtr CreateFontW(int cHeight,int cWidth,int cEscapement,int cOrientation,int cWeight,
    uint bItalic,uint bUnderline,uint bStrikeOut,uint iCharSet,uint iOutPrecision,uint iClipPrecision,
    uint iQuality,uint iPitchAndFamily,string pszFaceName);
  [DllImport("gdi32.dll", SetLastError=true)] public static extern IntPtr SelectObject(IntPtr hdc, IntPtr obj);
  [DllImport("gdi32.dll", SetLastError=true)] public static extern bool DeleteObject(IntPtr obj);
  [DllImport("gdi32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
  public static extern int GetTextFaceW(IntPtr hdc, int nCount, StringBuilder lpFaceName);
  [DllImport("gdi32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
  public static extern bool GetTextMetricsW(IntPtr hdc, ref TEXTMETRICW tm);
  [DllImport("gdi32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
  public static extern bool GetTextExtentPoint32W(IntPtr hdc, string s, int len, ref SIZE size);
  [DllImport("gdi32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
  public static extern bool GetTextExtentExPointW(IntPtr hdc, string s, int len, int maxExtent,
    IntPtr lpnFit, IntPtr alpDx, ref SIZE size);
  [DllImport("gdi32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
  public static extern bool GetCharABCWidthsW(IntPtr hdc, uint first, uint last, [Out] ABC[] abc);
  [DllImport("gdi32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
  public static extern bool GetCharABCWidthsFloatW(IntPtr hdc, uint first, uint last, [Out] ABCFLOAT[] abc);
  [DllImport("gdi32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
  public static extern bool GetCharWidth32W(IntPtr hdc, uint first, uint last, [Out] int[] widths);
  [DllImport("gdi32.dll", SetLastError=true)]
  public static extern uint GetGlyphOutlineW(IntPtr hdc, uint ch, uint format, ref GLYPHMETRICS gm,
    uint cbBuffer, IntPtr buffer, ref MAT2 mat2);
  [DllImport("gdi32.dll", SetLastError=true)]
  public static extern uint GetOutlineTextMetricsW(IntPtr hdc, uint cbData, byte[] data);
  [DllImport("gdi32.dll", SetLastError=true)]
  public static extern uint GetFontData(IntPtr hdc, uint table, uint offset, byte[] buffer, uint cbData);
  [DllImport("gdi32.dll", SetLastError=true)]
  public static extern int GetDeviceCaps(IntPtr hdc, int index);
  [DllImport("gdi32.dll", SetLastError=true)]
  public static extern uint GetGlyphIndicesW(IntPtr hdc, string s, int count, [Out] ushort[] indices, uint flags);
  [DllImport("gdi32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
  public static extern int AddFontResourceExW(string name, uint fl, IntPtr pdv);
  [DllImport("gdi32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
  public static extern bool RemoveFontResourceExW(string name, uint fl, IntPtr pdv);
  [DllImport("gdi32.dll", SetLastError=true)]
  public static extern bool GetCharABCWidthsI(IntPtr hdc, uint giFirst, uint cgi, [Out] ushort[] pgi, [Out] ABC[] abc);

  [StructLayout(LayoutKind.Sequential)] public struct SIZE { public int cx; public int cy; }
  [StructLayout(LayoutKind.Sequential)] public struct ABC { public int A; public int B; public int C; }
  [StructLayout(LayoutKind.Sequential)] public struct ABCFLOAT { public float A; public float B; public float C; }
  [StructLayout(LayoutKind.Sequential)] public struct POINT { public int x; public int y; }
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int left; public int top; public int right; public int bottom; }
  [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
  public struct TEXTMETRICW {
    public int tmHeight, tmAscent, tmDescent, tmInternalLeading, tmExternalLeading;
    public int tmAveCharWidth, tmMaxCharWidth, tmWeight, tmOverhang, tmDigitizedAspectX, tmDigitizedAspectY;
    public char tmFirstChar, tmLastChar, tmDefaultChar, tmBreakChar;
    public byte tmItalic, tmUnderlined, tmStruckOut, tmPitchAndFamily, tmCharSet;
  }
  [StructLayout(LayoutKind.Sequential)] public struct GLYPHMETRICS {
    public uint gmBlackBoxX, gmBlackBoxY;
    public POINT gmptGlyphOrigin;
    public short gmCellIncX, gmCellIncY;
  }
  [StructLayout(LayoutKind.Sequential)] public struct MAT2 { public FIXED eM11, eM12, eM21, eM22; }
  [StructLayout(LayoutKind.Sequential)] public struct FIXED { public ushort fract; public short value; }

  public static string DefaultPrinter() {
    int len = 1024; var sb = new StringBuilder(len);
    if (!GetDefaultPrinter(sb, ref len)) return "?";
    return sb.ToString();
  }
  public static string Face(IntPtr hdc) {
    var sb = new StringBuilder(256);
    int n = GetTextFaceW(hdc, 256, sb);
    return n > 0 ? sb.ToString() : "?";
  }
  public static string Hex(byte[] b, int n) {
    var sb = new StringBuilder(n * 2);
    for (int i = 0; i < n; i++) sb.Append(b[i].ToString("x2"));
    return sb.ToString();
  }
}
'@
Add-Type -TypeDefinition $src -Language CSharp

function Q([string]$s) { $s.Replace("`t"," ").Replace("`r"," ").Replace("`n"," ") }
$enc = New-Object System.Text.UTF8Encoding($false)
$w = New-Object System.IO.StreamWriter($Out, $false, $enc)
function W([string]$line) { $w.WriteLine($line) }

$printer = [GdiProbe]::DefaultPrinter()
$hdc = [GdiProbe]::CreateDCW("WINSPOOL", $printer, $null, [IntPtr]::Zero)
if ($hdc -eq [IntPtr]::Zero) { throw "CreateDC failed for $printer : $([Runtime.InteropServices.Marshal]::GetLastWin32Error())" }
$screen = [GdiProbe]::GetDC([IntPtr]::Zero)

W ("device`tdefault_printer`t" + (Q $printer))
$names = @{ 0="driverVersion";2="technology";4="horzSizeMm";6="vertSizeMm";8="horzRes";10="vertRes";12="bitsPixel";14="planes";16="numBrushes";18="numPens";20="numMarkers";22="numFonts";24="numColors";26="pdeviceSize";28="curveCaps";36="clipCaps";38="rasterCaps";40="aspectX";42="aspectY";44="aspectXY";88="logPixelsX";90="logPixelsY";104="sizePalette";106="numReserved";108="colorRes";110="physicalWidth";111="physicalHeight";112="physicalOffsetX";113="physicalOffsetY" }
foreach ($idx in ($names.Keys | Sort-Object)) {
  $v = [GdiProbe]::GetDeviceCaps($hdc, $idx)
  W ("devcap`t600`t$($names[$idx])`t$v")
}
foreach ($idx in ($names.Keys | Sort-Object)) {
  $v = [GdiProbe]::GetDeviceCaps($screen, $idx)
  W ("devcap`t96`t$($names[$idx])`t$v")
}

$strings = @("-98.41%","0.00%","#######","0.000%","%","#","0","-0.98","(0.98)","1234567.00","1,234,567.00","0.50")

function Emit-Font([string]$family,[bool]$italic,[int]$pt,[int]$scale,[int]$quality,[int]$outp,[string]$tag) {
  $dpi = 600
  $h = -[int][math]::Floor($pt * $scale / 100.0 * $dpi / 72.0 + 0.5)
  $lf = [GdiProbe]::CreateFontW($h, 0, 0, 0, 400, [uint32]$italic, 0, 0, 1, [uint32]$outp, 0, [uint32]$quality, 0, $family)
  if ($lf -eq [IntPtr]::Zero) { W "font`t$tag`terror`tCreateFont failed"; return }
  [void][GdiProbe]::SelectObject($hdc, $lf)
  $tm = New-Object GdiProbe+TEXTMETRICW
  [void][GdiProbe]::GetTextMetricsW($hdc, [ref]$tm)
  $face = Q ([GdiProbe]::Face($hdc))
  W ("font`t$tag`t$face`tpt=$pt`tscale=$scale`theightPx=$(-$h)`tquality=$quality`toutprecis=$outp`titalic=$italic`tweight=400`ttmHeight=$($tm.tmHeight)`ttmAscent=$($tm.tmAscent)`ttmDescent=$($tm.tmDescent)`ttmInternalLeading=$($tm.tmInternalLeading)`ttmExternalLeading=$($tm.tmExternalLeading)`ttmAveCharWidth=$($tm.tmAveCharWidth)`ttmMaxCharWidth=$($tm.tmMaxCharWidth)`ttmOverhang=$($tm.tmOverhang)`ttmItalic=$($tm.tmItalic)`ttmPitch=0x$('{0:x2}' -f $tm.tmPitchAndFamily)")
  $abc = New-Object 'GdiProbe+ABC[]' 256
  $ok = [GdiProbe]::GetCharABCWidthsW($hdc, 0, 255, $abc)
  if ($ok) { for ($c = 32; $c -le 126; $c++) { W ("abc`t$tag`t$c`t$($abc[$c].A)`t$($abc[$c].B)`t$($abc[$c].C)") } }
  $cw = New-Object 'int[]' 256
  if ([GdiProbe]::GetCharWidth32W($hdc,0,255,$cw)) { for ($c=32;$c -le 126;$c++){ W ("charwidth32`t$tag`t$c`t$($cw[$c])") } }
  foreach ($cp in @(35,37,48,57,45,40,41)) {
    $gm = New-Object GdiProbe+GLYPHMETRICS
    $mat = New-Object GdiProbe+MAT2; $mat.eM11.value=1; $mat.eM22.value=1
    $r = [GdiProbe]::GetGlyphOutlineW($hdc,[uint32]$cp,0,[ref]$gm,0,[IntPtr]::Zero,[ref]$mat)
    W ("glyph`t$tag`t$cp`tbbx=$($gm.gmBlackBoxX)`toriginX=$($gm.gmptGlyphOrigin.x)`tincX=$($gm.gmCellIncX)`tret=$r")
  }
  foreach ($s in $strings) {
    $sz = New-Object GdiProbe+SIZE
    [void][GdiProbe]::GetTextExtentPoint32W($hdc, $s, $s.Length, [ref]$sz)
    $sumABC=0; $sumB=0; $lastC=0
    for ($i=0;$i -lt $s.Length;$i++) { $cc=[int]$s[$i]; $sumABC += $abc[$cc].A+$abc[$cc].B+$abc[$cc].C; $sumB += $abc[$cc].B; if ($i -eq $s.Length-1){$lastC=$abc[$cc].C} }
    W ("run`t$tag`t$s`textent=$($sz.cx)`tsumABC=$sumABC`tsumB=$sumB`tlastC=$lastC")
  }
  [void][GdiProbe]::SelectObject($hdc,[IntPtr]::Zero); [void][GdiProbe]::DeleteObject($lf)
}

# synthesized italics: families with no italic face installed
$synth = @(
  @("Arial Black",14),@("Arial Black",8),@("Lucida Console",14),@("Lucida Console",8),
  @("MS Sans Serif",14),@("MS Sans Serif",8),@("Small Fonts",14),@("Small Fonts",8),
  @("Courier New",14),@("Tahoma",14)
)
$i=0
foreach ($c in $synth) {
  $i++
  Emit-Font $c[0] $false $c[1] 100 0 0 ("s{0:D2}u" -f $i)
  Emit-Font $c[0] $true  $c[1] 100 0 0 ("s{0:D2}i" -f $i)
}
# Aptos Narrow private font
$added=@()
foreach ($f in @("aptosnarrow-27970306811.ttf","aptosnarrow-28903934312.ttf","aptosnarrow-25065980124.ttf")) {
  $path = Join-Path $PSScriptRoot $f
  if (Test-Path $path) { $n=[GdiProbe]::AddFontResourceExW($path,0x10,[IntPtr]::Zero); W "fontadd`t$f`t$n"; $added += $f }
}
$i=0
foreach ($pt in @(8,9,11,12,14,20)) {
  foreach ($sc in @(100,75)) {
    $i++
    Emit-Font "Aptos Narrow" $false $pt $sc 0 0 ("a{0:D2}u" -f $i)
    Emit-Font "Aptos Narrow" $true  $pt $sc 0 0 ("a{0:D2}i" -f $i)
  }
}
foreach ($f in $added) { [void][GdiProbe]::RemoveFontResourceExW((Join-Path $PSScriptRoot $f),0x10,[IntPtr]::Zero) }

$w.Close(); [void][GdiProbe]::DeleteDC($hdc); [void][GdiProbe]::ReleaseDC([IntPtr]::Zero,$screen)
Write-Output "wrote $Out"
