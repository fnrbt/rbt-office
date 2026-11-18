param([Parameter(Mandatory = $true)][string]$InputPath, [Parameter(Mandatory = $true)][string]$OutputPath)
# Compile errors as Excel's VBA gives them: each case (InputPath/*.bas, a standard module; a case's
# <name>.<Class>.cls files are class modules beside it) goes into a blank workbook, and VBE's
# "Compile VBAProject" (control 578) compiles the project. A compile error is a dialog: a watcher process
# reads its text and closes it. OutputPath: case <TAB> "ok" | the dialog's text (lines joined by " | ").
# Run in the interactive session (tools/excel-parity/run-excel-vbs-task.ps1 -Engine powershell).
$ErrorActionPreference = 'Stop'
$watcher = @'
param([string]$Out, [int]$Seconds)
Add-Type -Namespace RbtCompile -Name Win32 -MemberDefinition @"
public delegate bool EnumProc(System.IntPtr hwnd, System.IntPtr lParam);
[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc f, System.IntPtr p);
[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern bool EnumChildWindows(System.IntPtr w, EnumProc f, System.IntPtr p);
[System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)] public static extern int GetWindowText(System.IntPtr w, System.Text.StringBuilder s, int n);
[System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)] public static extern int GetClassName(System.IntPtr w, System.Text.StringBuilder s, int n);
[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern bool IsWindowVisible(System.IntPtr w);
[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern System.IntPtr SendMessage(System.IntPtr w, int msg, System.IntPtr a, System.IntPtr b);
"@
function Get-WindowText($h) { $sb = New-Object System.Text.StringBuilder 1024; [void][RbtCompile.Win32]::GetWindowText($h, $sb, 1024); $sb.ToString() }
function Get-WindowClass($h) { $sb = New-Object System.Text.StringBuilder 256; [void][RbtCompile.Win32]::GetClassName($h, $sb, 256); $sb.ToString() }
$deadline = (Get-Date).AddSeconds($Seconds)
while ((Get-Date) -lt $deadline) {
    $found = $null
    [RbtCompile.Win32]::EnumWindows({ param($h, $l)
        if ([RbtCompile.Win32]::IsWindowVisible($h) -and (Get-WindowClass $h) -eq '#32770' -and (Get-WindowText $h) -like 'Microsoft Visual Basic*') { $script:found = $h; return $false }
        return $true }, [System.IntPtr]::Zero) | Out-Null
    if ($script:found) {
        $texts = New-Object System.Collections.ArrayList
        $ok = $null
        [RbtCompile.Win32]::EnumChildWindows($script:found, { param($h, $l)
            $c = Get-WindowClass $h; $t = Get-WindowText $h
            if ($c -eq 'Static' -and $t) { [void]$texts.Add($t) }
            if ($c -eq 'Button' -and $t -eq 'OK') { $script:ok = $h }
            return $true }, [System.IntPtr]::Zero) | Out-Null
        [System.IO.File]::WriteAllText($Out, (($texts -join ' | ') -replace "`r?`n", ' '))
        if ($script:ok) { [void][RbtCompile.Win32]::SendMessage($script:ok, 0x00F5, [System.IntPtr]::Zero, [System.IntPtr]::Zero) }
        else { [void][RbtCompile.Win32]::SendMessage($script:found, 0x0010, [System.IntPtr]::Zero, [System.IntPtr]::Zero) }
        exit
    }
    Start-Sleep -Milliseconds 200
}
# (Nothing found: what was on the screen, for a look.)
$seen = New-Object System.Collections.ArrayList
[RbtCompile.Win32]::EnumWindows({ param($h, $l)
    if ([RbtCompile.Win32]::IsWindowVisible($h)) { [void]$seen.Add((Get-WindowClass $h) + "`t" + (Get-WindowText $h)) }
    return $true }, [System.IntPtr]::Zero) | Out-Null
[System.IO.File]::WriteAllText($Out + '.windows', ($seen -join "`n"))
'@
$work = Split-Path -Parent ([System.IO.Path]::GetFullPath($OutputPath))
$watcherPath = Join-Path $work 'compile-watcher.ps1'
Set-Content -LiteralPath $watcherPath -Value $watcher -Encoding UTF8
$lines = New-Object System.Collections.ArrayList
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false
try {
    foreach ($case in Get-ChildItem -LiteralPath $InputPath -Filter *.bas | Sort-Object Name) {
        $name = $case.BaseName
        $book = $excel.Workbooks.Add()
        $project = $book.VBProject
        $module = $project.VBComponents.Add(1)
        $module.CodeModule.AddFromString([System.IO.File]::ReadAllText($case.FullName))
        foreach ($cls in Get-ChildItem -LiteralPath $InputPath -Filter "$name.*.cls") {
            $component = $project.VBComponents.Add(2)
            $component.Name = $cls.BaseName.Substring($name.Length + 1)
            $component.CodeModule.AddFromString([System.IO.File]::ReadAllText($cls.FullName))
        }
        $dialog = Join-Path $work "dialog-$name.txt"
        Remove-Item -LiteralPath $dialog -ErrorAction SilentlyContinue
        $watch = Start-Process powershell.exe -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $watcherPath, '-Out', $dialog, '-Seconds', '8') -PassThru -WindowStyle Hidden
        # (Compile acts on the VBE's active project; the VBE's window is shown for its commands to work.)
        $excel.VBE.MainWindow.Visible = $true
        $excel.VBE.ActiveVBProject = $project
        $control = $excel.VBE.CommandBars.FindControl([Type]::Missing, 578)
        $state = if ($control -eq $null) { 'no control' } elseif (-not $control.Enabled) { 'disabled' } else { $control.Execute(); '' }
        $watch.WaitForExit()
        $result = if (Test-Path -LiteralPath $dialog) { Get-Content -LiteralPath $dialog -Raw } elseif ($state) { "?? $state" } else { 'ok' }
        [void]$lines.Add("$name`t$result")
        $book.Close($false)
    }
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
[System.IO.File]::WriteAllText($OutputPath, (($lines -join "`n") + "`n"), (New-Object System.Text.UTF8Encoding($false)))
