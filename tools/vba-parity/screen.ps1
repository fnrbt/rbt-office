param([string]$InputPath, [Parameter(Mandatory = $true)][string]$OutputPath)
# A PNG of the whole interactive desktop of the VM (what a stuck Office dialog shows), and the titles of
# the visible top-level windows beside it (OutputPath with .txt).
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
$bitmap = New-Object System.Drawing.Bitmap $bounds.Width, $bounds.Height
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size)
$bitmap.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
Get-Process | Where-Object { $_.MainWindowTitle } | ForEach-Object { "$($_.ProcessName)`t$($_.MainWindowTitle)" } | Set-Content -LiteralPath ([System.IO.Path]::ChangeExtension($OutputPath, ".txt"))
# Every visible top-level window with its title and the texts of its child controls (a dialog's message).
Add-Type -Namespace RbtScreen -Name Win32 -MemberDefinition @'
public delegate bool EnumProc(System.IntPtr hwnd, System.IntPtr lParam);
[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc f, System.IntPtr p);
[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern bool EnumChildWindows(System.IntPtr parent, EnumProc f, System.IntPtr p);
[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern bool IsWindowVisible(System.IntPtr hwnd);
[System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)] public static extern int GetWindowText(System.IntPtr hwnd, System.Text.StringBuilder s, int n);
[System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)] public static extern int GetClassName(System.IntPtr hwnd, System.Text.StringBuilder s, int n);
'@
$lines = New-Object System.Collections.Generic.List[string]
$text = { param($h) $b = New-Object System.Text.StringBuilder 512; [void][RbtScreen.Win32]::GetWindowText($h, $b, 512); $b.ToString() }
$class = { param($h) $b = New-Object System.Text.StringBuilder 256; [void][RbtScreen.Win32]::GetClassName($h, $b, 256); $b.ToString() }
[RbtScreen.Win32]::EnumWindows({ param($h, $p)
  if ([RbtScreen.Win32]::IsWindowVisible($h)) {
    $lines.Add("WINDOW [" + (& $class $h) + "] " + (& $text $h))
    [RbtScreen.Win32]::EnumChildWindows($h, { param($c, $q) $t = & $text $c; if ($t) { $lines.Add("    [" + (& $class $c) + "] " + $t) }; $true }, [System.IntPtr]::Zero) | Out-Null
  }
  $true }, [System.IntPtr]::Zero) | Out-Null
$lines | Add-Content -LiteralPath ([System.IO.Path]::ChangeExtension($OutputPath, ".txt"))
