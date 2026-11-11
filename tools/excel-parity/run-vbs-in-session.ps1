# Runs a VBS (or a .ps1) in the desktop session (a scheduled task with /IT: Excel cannot be started from an SSH session)
# with the arguments given, and waits for it. Unlike run-excel-vbs-task.ps1 it kills nothing by name: on a
# timeout only the Excel processes that started while the script ran are stopped, by their ids, so an
# Excel or Word that someone else runs on the VM is left alone.
#   powershell -File run-vbs-in-session.ps1 -ScriptPath C:\...\open-all.vbs -Arguments 'C:\...\dir' [-TimeoutSeconds 3600]
# Writes the script's output to <script>.log beside it and exits with the script's exit code.
param(
  [Parameter(Mandatory = $true)][string]$ScriptPath,
  [string]$Arguments = "",
  [int]$TimeoutSeconds = 3600
)

$ErrorActionPreference = "Stop"
$scriptFull = (Resolve-Path -LiteralPath $ScriptPath).Path
$workDir = Split-Path -Parent $scriptFull
$taskName = "RedBatExcelParity-" + ([guid]::NewGuid().ToString("N"))
$logPath = [System.IO.Path]::ChangeExtension($scriptFull, ".log")
$donePath = [System.IO.Path]::ChangeExtension($scriptFull, ".done")
$cmdPath = Join-Path $workDir ($taskName + ".cmd")
$runner = if ($scriptFull.EndsWith(".ps1")) { "powershell -NoProfile -ExecutionPolicy Bypass -File" } else { "cscript //nologo" }
$before = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })

Set-Content -LiteralPath $cmdPath -Encoding ASCII -Value @"
@echo off
cd /d "$workDir"
del /q "$donePath" 2>nul
$runner "$scriptFull" $Arguments > "$logPath" 2>&1
echo %ERRORLEVEL% > "$donePath"
"@

try {
  $start = (Get-Date).AddMinutes(10)
  & schtasks.exe /Create /TN $taskName /TR ('cmd.exe /c "' + $cmdPath + '"') /SC ONCE /SD $start.ToString("MM/dd/yyyy") /ST $start.ToString("HH:mm") /F /IT | Out-Null
  if ($LASTEXITCODE -ne 0) { throw "schtasks /Create failed: $LASTEXITCODE" }
  & schtasks.exe /Run /TN $taskName | Out-Null
  if ($LASTEXITCODE -ne 0) { throw "schtasks /Run failed: $LASTEXITCODE" }
  $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
  while ((Get-Date) -lt $deadline) {
    if (Test-Path -LiteralPath $donePath) {
      $code = 1
      [void][int]::TryParse((Get-Content -LiteralPath $donePath -Raw).Trim(), [ref]$code)
      if (Test-Path -LiteralPath $logPath) { Get-Content -LiteralPath $logPath }
      exit $code
    }
    Start-Sleep -Seconds 2
  }
  & schtasks.exe /End /TN $taskName 2>$null | Out-Null
  Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object { Stop-Process -Id $_.Id -Force }
  if (Test-Path -LiteralPath $logPath) { Get-Content -LiteralPath $logPath }
  throw "timed out after $TimeoutSeconds s"
}
finally {
  & schtasks.exe /Delete /TN $taskName /F 2>$null | Out-Null
  Remove-Item -LiteralPath $cmdPath -Force -ErrorAction SilentlyContinue
}
