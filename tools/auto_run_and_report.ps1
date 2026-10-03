# Auto-run a SolidWorks scheduled task and report feedback.
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File auto_run_and_report.ps1
#     [-TaskName CodexSW_Gripper_Auto] [-StatusFile run_status.txt]
#     [-ConsoleFile run_console.log] [-BuildLog build_log.txt]
#     [-DesktopPattern Desktop\*.SLDPRT] [-TimeoutSec 900]
param(
    [string]$TaskName = "CodexSW_Gripper_Auto",
    [string]$OutDir = (Join-Path $env:USERPROFILE "Documents\Codex\2026-08-15\he-2\outputs"),
    [string]$StatusFile = "run_status.txt",
    [string]$ConsoleFile = "run_console.log",
    [string]$BuildLog = "build_log.txt",
    [string]$DesktopPattern = "Desktop\*.SLDPRT",
    [int]$TimeoutSec = 900
)
$ErrorActionPreference = "Stop"

$taskName    = $TaskName
$statusFile  = Join-Path $OutDir $StatusFile
$consoleLog  = Join-Path $OutDir $ConsoleFile
$buildLog    = Join-Path $OutDir $BuildLog
$deadlineSec = $TimeoutSec

# 1) task must exist
$task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if (-not $task) { throw "Scheduled task '$taskName' not found. Register it first." }

# 2) clear previous status so we only wait for THIS run
Remove-Item $statusFile -ErrorAction SilentlyContinue

# 3) trigger the task (runs in the interactive user session)
schtasks /run /tn $taskName | Out-Null
if ($LASTEXITCODE -ne 0) { throw "schtasks /run failed with code $LASTEXITCODE" }

# 4) wait for completion (status file gets an 'exit=' line)
$deadline = (Get-Date).AddSeconds($deadlineSec)
$done = $false
while ((Get-Date) -lt $deadline) {
    if (Test-Path $statusFile) {
        $content = Get-Content $statusFile -Raw -ErrorAction SilentlyContinue
        if ($content -match "exit=") { $done = $true; break }
    }
    Start-Sleep -Seconds 3
}
if (-not $done) { throw "Timed out waiting for scheduled task to finish" }

# 5) feedback report
Write-Host "===== run_status ====="
Get-Content $statusFile
Write-Host "`n===== build_log.txt ====="
if (Test-Path $buildLog) { Get-Content $buildLog -Encoding UTF8 }
Write-Host "`n===== run_console.log (last 45 lines) ====="
if (Test-Path $consoleLog) { Get-Content $consoleLog -Encoding UTF8 | Select-Object -Last 45 }
Write-Host "`n===== newest .SLDPRT on Desktop ====="
$part = Get-ChildItem (Join-Path $env:USERPROFILE $DesktopPattern) -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($part) {
    $part | Select-Object FullName, Length, LastWriteTime | Format-List
} else {
    Write-Host "NOT FOUND: no .SLDPRT on Desktop"
}
