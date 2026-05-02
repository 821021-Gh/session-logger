# SessionReport Scheduler Setup
# Usage:
#   setup_report_task.ps1            → interactive (asks for time)
#   setup_report_task.ps1 21:00      → set send time to 21:00 daily
#   setup_report_task.ps1 disable     → remove the scheduled task

param(
    [string]$Time   = $null,
    [string]$Action  = "setup"
)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$reportPy  = Join-Path $scriptDir "session_report.py"
$taskName  = "SessionReport_Daily"

if ($Action -eq "disable") {
    schtasks /Delete /TN $taskName /F *>$null 2>$null
    Write-Host "[OK] Report scheduler removed."
    exit 0
}

if (-not $Time) {
    Write-Host "Session Report Daily Email Scheduler"
    Write-Host "===================================="
    Write-Host "Enter the time to send daily report (24h format, e.g. 20:00): "
    $Time = Read-Host
}

if ($Time -notmatch "^\d{1,2}:\d{2}$") {
    Write-Host "[ERROR] Invalid format. Use HH:MM (e.g. 20:00)"
    exit 1
}

$h,[int]$m = $Time -split ':'
if ($h -lt 0 -or $h -gt 23 -or $m -lt 0 -or $m -gt 59) {
    Write-Host "[ERROR] Time out of range."
    exit 1
}

schtasks /Delete /TN $taskName /F *>$null 2>$null
schtasks /Create /TN $taskName /SC DAILY /ST $Time /TR "python `"$reportPy`" --email-only" /F 2>$null | Out-Null

if ($LASTEXITCODE -eq 0) {
    Write-Host "[OK] Report scheduled daily at $Time"
    Write-Host "    Task: $taskName"
    Write-Host "    Command: python session_report.py --email-only"
} else {
    Write-Host "[ERROR] Failed to create scheduled task."
}