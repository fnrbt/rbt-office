param(
  [Parameter(Mandatory = $true)][string]$ScriptPath,
  [Parameter(Mandatory = $true)][string]$InputPath,
  [Parameter(Mandatory = $true)][string]$OutputPath,
  [int]$TimeoutSeconds = 600
)

$ErrorActionPreference = "Stop"

function Quote-Arg([string]$Value) {
  '"' + $Value.Replace('"', '\"') + '"'
}

function New-TaskStartArgs([int]$MinutesFromNow) {
  $start = (Get-Date).AddMinutes($MinutesFromNow)
  @("/SC", "ONCE", "/SD", $start.ToString("MM/dd/yyyy"), "/ST", $start.ToString("HH:mm"))
}

$scriptFull = (Resolve-Path -LiteralPath $ScriptPath).Path
$inputFull = (Resolve-Path -LiteralPath $InputPath).Path
$outputFull = [System.IO.Path]::GetFullPath($OutputPath)
$workDir = Split-Path -Parent $scriptFull
$taskName = "RedBatWordParity-" + ([guid]::NewGuid().ToString("N"))
$logPath = [System.IO.Path]::ChangeExtension($outputFull, ".task.log")
$donePath = [System.IO.Path]::ChangeExtension($outputFull, ".task.done")
$cmdPath = Join-Path $workDir ($taskName + ".cmd")

$outputParent = Split-Path -Parent $outputFull
if ($outputParent -and -not (Test-Path -LiteralPath $outputParent)) {
  New-Item -ItemType Directory -Force -Path $outputParent | Out-Null
}

function Stop-InteractiveWord {
  $cleanupCmd = Join-Path $workDir ("kill-word-" + ([guid]::NewGuid().ToString("N")) + ".cmd")
  Set-Content -LiteralPath $cleanupCmd -Encoding ASCII -Value @(
    "@echo off",
    "for /L %%i in (1,1,5) do (",
    "  taskkill /F /T /IM WINWORD.EXE >nul 2>nul",
    "  ping 127.0.0.1 -n 2 >nul",
    ")"
  )

  try {
    $cleanupTaskName = "RedBatWordParity-KillWord-" + ([guid]::NewGuid().ToString("N"))
    try {
      $cleanupRun = 'cmd.exe /c "' + $cleanupCmd + '"'
      $cleanupStartArgs = New-TaskStartArgs 10
      & schtasks.exe /Create /TN $cleanupTaskName /TR $cleanupRun @cleanupStartArgs /F /IT 2>$null | Out-Null
      if ($LASTEXITCODE -eq 0) {
        & schtasks.exe /Run /TN $cleanupTaskName 2>$null | Out-Null
        Start-Sleep -Seconds 7
      }
    }
    finally {
      & schtasks.exe /Delete /TN $cleanupTaskName /F 2>$null | Out-Null
    }
  }
  finally {
    Remove-Item -LiteralPath $cleanupCmd -Force -ErrorAction SilentlyContinue
  }
}

$cmd = @"
@echo off
cd /d "$workDir"
del /q "$donePath" 2>nul
cscript //nologo "$scriptFull" "$inputFull" "$outputFull" > "$logPath" 2>&1
set RB_EXIT=%ERRORLEVEL%
echo %RB_EXIT% > "$donePath"
exit /b %RB_EXIT%
"@
Set-Content -LiteralPath $cmdPath -Value $cmd -Encoding ASCII

try {
  $startArgs = New-TaskStartArgs 10
  $taskRun = 'cmd.exe /c "' + $cmdPath + '"'
  & schtasks.exe /Create /TN $taskName /TR $taskRun @startArgs /F /IT | Out-Null
  if ($LASTEXITCODE -ne 0) {
    throw "schtasks /Create failed with exit code $LASTEXITCODE"
  }
  & schtasks.exe /Run /TN $taskName | Out-Null
  if ($LASTEXITCODE -ne 0) {
    throw "schtasks /Run failed with exit code $LASTEXITCODE"
  }

  $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
  while ((Get-Date) -lt $deadline) {
    if (Test-Path -LiteralPath $donePath) {
      $exitCodeText = (Get-Content -LiteralPath $donePath -Raw).Trim()
      $exitCode = 1
      [void][int]::TryParse($exitCodeText, [ref]$exitCode)
      if ($exitCode -ne 0) {
        if (Test-Path -LiteralPath $logPath) {
          Get-Content -LiteralPath $logPath | Write-Error
        }
        throw "VBS task failed with exit code $exitCode"
      }
      if (-not (Test-Path -LiteralPath $outputFull)) {
        throw "VBS task completed but output was not created: $outputFull"
      }
      Write-Output $outputFull
      exit 0
    }
    Start-Sleep -Seconds 1
  }

  if (Test-Path -LiteralPath $logPath) {
    Get-Content -LiteralPath $logPath | Write-Error
  }
  & schtasks.exe /End /TN $taskName 2>$null | Out-Null
  Stop-InteractiveWord
  throw "Timed out waiting for VBS task after $TimeoutSeconds seconds"
}
finally {
  & schtasks.exe /Delete /TN $taskName /F 2>$null | Out-Null
  Remove-Item -LiteralPath $cmdPath -Force -ErrorAction SilentlyContinue
  Stop-InteractiveWord
}
