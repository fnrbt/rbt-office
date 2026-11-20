# Side by side: Word (left) and rbt-office in Edge (right) on one document; Word scrolls, the page follows.
param([string]$Doc = 'magazine-feature.docx', [int]$Steps = 400, [int]$StepMs = 40, [int]$Lines = 1,
      [int]$HoldMs = 2500, [string]$Out = '', [int]$ProbeY = 420, [switch]$Test, [switch]$Keep, [int]$Wheel = 0, [int]$Seconds = 20, [int]$DragFrom = 0, [int]$DragTo = 0, [int]$DragX = 948, [switch]$Take, [int]$TakeMs = 9000, [switch]$Probe, [int]$LagMs = 0)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -Name W -Namespace S -MemberDefinition '[DllImport("user32.dll")] public static extern IntPtr GetDC(IntPtr h); [DllImport("gdi32.dll")] public static extern uint GetPixel(IntPtr dc, int x, int y); [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y); [DllImport("user32.dll")] public static extern void mouse_event(int f, int x, int y, int d, int e); [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow(); [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int c);'
[void][S.W]::ShowWindow([S.W]::GetConsoleWindow(), 0)
Get-Process OneDrive -ErrorAction SilentlyContinue | Stop-Process -Force
Get-Process msedge -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -like '*rbt-office*' } | Stop-Process -Force
$web = New-Object System.Net.WebClient
$post = { param($s) try { [void]$web.UploadString('http://localhost:16990/pos', $s) } catch {} }
& $post ''

$w = New-Object -ComObject Word.Application
$w.Visible = $true
$d = $w.Documents.Open("C:\sbs\$Doc", $false, $true)
$w.WindowState = 0
$w.Left = 0; $w.Top = 0; $w.Width = 720; $w.Height = 780      # points: 960 x 1040 px
$win = $w.ActiveWindow
$win.View.Type = 3; $win.View.Zoom.Percentage = 100
$win.DocumentMap = $false
$d.ActiveWindow.VerticalPercentScrolled = 0
$edge = 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
Start-Process $edge "--app=http://localhost:16990/#$Doc --window-position=960,0 --window-size=960,1040 --user-data-dir=C:\sbs\edge --no-first-run --remote-debugging-port=9333 --disable-background-timer-throttling --disable-renderer-backgrounding --disable-backgrounding-occluded-windows"
Start-Sleep -Seconds 12
$w.Activate()

$px = 96 / 72 * $win.View.Zoom.Percentage / 100
function Position {
  # The point of Word's page under a probe line: from the text range there (an object hit: its range,
  # else another probe height).
  $get = { param($o, $n, $a) [System.__ComObject].InvokeMember($n, [Reflection.BindingFlags]::GetProperty, $null, $o, $a) }
  foreach ($py in @($ProbeY, 300, 900, 200, 700, 1000)) {
    try {
      $r = $win.RangeFromPoint(170, $py); if ($null -eq $r) { continue }
      try { $page = & $get $r 'Information' @(3) } catch { $r = & $get $r 'Range' $null; $page = & $get $r 'Information' @(3) }
      $y = & $get $r 'Information' @(6)
      $l = 0; $t = 0; $wd = 0; $h = 0; $win.GetPoint([ref]$l, [ref]$t, [ref]$wd, [ref]$h, $r)
      return "{0} {1:0.00} {2}" -f $page, ($y + ($py - $t) / $px), $py
    } catch { }
  }
  return $null
}
if ($Probe) {
  foreach ($y in 300, 420, 600, 900) { foreach ($x in 170, 300, 600) {
    try { $r = $win.RangeFromPoint($x, $y); $kind = if ($null -eq $r) { 'null' } else { [Microsoft.VisualBasic.Information]::TypeName($r) }
          $l = 0; $t = 0; $wd = 0; $h = 0; $win.GetPoint([ref]$l, [ref]$t, [ref]$wd, [ref]$h, $r)
          $pg = [System.__ComObject].InvokeMember('Information', [Reflection.BindingFlags]::GetProperty, $null, $r, @(3))
          "$x,$y $kind top=$t page=$pg" } catch { "$x,$y ERR $($_.Exception.Message)" } } }
  "position: $(Position)"
  $d.Close($false); $w.Quit(); exit
}
function StartRec {
  $psi = New-Object Diagnostics.ProcessStartInfo 'C:\sbs\ffmpeg.exe', "-v error -y -f gdigrab -framerate 30 -draw_mouse 0 -i desktop -c:v libx264 -preset ultrafast -crf 14 -pix_fmt yuv420p C:\sbs\$Out"
  $psi.UseShellExecute = $false; $psi.RedirectStandardInput = $true; $psi.CreateNoWindow = $true
  return [Diagnostics.Process]::Start($psi)
}
function StopRec($p) { if ($p) { $p.StandardInput.Write('q'); $p.StandardInput.Close(); [void]$p.WaitForExit(30000) } }
$ff = $null
if ($Out -and -not $Take) {
  $ff = StartRec
}
& $post (Position)
Start-Sleep -Milliseconds $HoldMs
$log = @()
function Thumb {
  # Hover shows the scroll bar; the thumb is the white run (track 243) between the arrows.
  [void][S.W]::SetCursorPos($DragX, 600); Start-Sleep -Milliseconds 700
  $dc = [S.W]::GetDC([IntPtr]::Zero); $top = -1; $bottom = -1
  for ($y = 176; $y -le 1000; $y++) { $v = [S.W]::GetPixel($dc, $DragX, $y) -band 0xFF; if ($v -ge 250) { if ($top -lt 0) { $top = $y }; $bottom = $y } elseif ($top -ge 0) { break } }
  return @($top, $bottom)
}
function Ease([double]$u) { return $u * $u * (3 - 2 * $u) }
function Drag([int]$from, [int]$to, [int]$ms, [string]$plan = '') {
  [void][S.W]::SetCursorPos($DragX, $from); Start-Sleep -Milliseconds 200
  [S.W]::mouse_event(0x0002, 0, 0, 0, 0)
  $sw = [Diagnostics.Stopwatch]::StartNew()
  if ($plan) { & $post $plan }
  while ($sw.ElapsedMilliseconds -lt $ms) { [void][S.W]::SetCursorPos($DragX, [int]($from + ($to - $from) * (Ease ($sw.ElapsedMilliseconds / $ms)))); Start-Sleep -Milliseconds 4 }
  [void][S.W]::SetCursorPos($DragX, $to); Start-Sleep -Milliseconds 50
  [S.W]::mouse_event(0x0004, 0, 0, 0, 0)
  [void][S.W]::SetCursorPos(1900, 1070)
}
if ($Take) {
  $t = Thumb; $from = [int](($t[0] + $t[1]) / 2); $to = $from + (1000 - $t[1])
  "thumb $($t[0])-$($t[1]); drag $from -> $to"
  $start = Position
  Drag $from $to $TakeMs
  Start-Sleep -Milliseconds 300; $end = Position; "rehearsal: $start -> $end"
  $win.VerticalPercentScrolled = 0; Start-Sleep -Milliseconds 300
  $start = Position; & $post $start
  StopRec $ff; $ff = $null
  if ($Out) { $ff = StartRec }
  Start-Sleep -Milliseconds $HoldMs
  Drag $from $to $TakeMs ("plan $start $end $TakeMs $LagMs")
  Start-Sleep -Milliseconds 300; $real = Position; "take: $start -> $real (planned $end)"
  Start-Sleep -Milliseconds $HoldMs
  $Steps = 0
}
if ($DragFrom -gt 0) {
  # Word's scroll bar thumb dragged down, a pixel at a time: Word scrolls continuously.
  [void][S.W]::SetCursorPos($DragX, $DragFrom); Start-Sleep -Milliseconds 300
  [S.W]::mouse_event(0x0002, 0, 0, 0, 0)
  $sw = [Diagnostics.Stopwatch]::StartNew(); $prev = ''
  for ($y = $DragFrom; $y -le $DragTo; $y++) {
    [void][S.W]::SetCursorPos($DragX, $y)
    $p = $null; try { $p = Position } catch {}
    if ($p -and $p -ne $prev) { & $post $p; $prev = $p }
    if ($Test) { $log += ('{0} {1} {2}' -f $sw.ElapsedMilliseconds, $y, $p) }
    while ($sw.ElapsedMilliseconds -lt ($y - $DragFrom + 1) * $StepMs) { Start-Sleep -Milliseconds 2 }
  }
  [S.W]::mouse_event(0x0004, 0, 0, 0, 0)
  $Steps = 0
}
if ($Wheel -ne 0) {
  [void][S.W]::SetCursorPos(480, 600)
  $sw = [Diagnostics.Stopwatch]::StartNew(); $next = 0; $prev = ''
  while ($sw.ElapsedMilliseconds -lt $Seconds * 1000) {
    if ($sw.ElapsedMilliseconds -ge $next) { [S.W]::mouse_event(0x0800, 0, 0, -$Wheel, 0); $next += $StepMs }
    $p = Position
    if ($p -and $p -ne $prev) { & $post $p; $prev = $p; if ($Test) { $log += ('{0} {1}' -f $sw.ElapsedMilliseconds, $p) } }
    Start-Sleep -Milliseconds 10
  }
  $Steps = 0
}
for ($i = 0; $i -lt $Steps; $i++) {
  $win.SmallScroll($Lines)
  $p = Position
  if ($p) { & $post $p }
  if ($Test) { $log += $p }
  Start-Sleep -Milliseconds $StepMs
}
Start-Sleep -Milliseconds $HoldMs
StopRec $ff; $ff = $null
if ($Test) { $log }
if (-not $Keep) { $d.Close($false); $w.Quit() }
