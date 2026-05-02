$taskName = "SessionReport_Daily"
$scriptPath = "C:\Users\spunk\.qclaw-oversea\workspace\session_report.py"

$action = New-ScheduledTaskAction -Execute "python" -Argument "`"$scriptPath`" --email-only" -WorkingDirectory "C:\Users\spunk\.qclaw-oversea\workspace"
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited
$trigger = New-ScheduledTaskTrigger -Daily -At "20:00"
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 1)

Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description "Daily session report: generate Excel chart and email to spunk.chang@gmail.com"
Write-Host "[OK] Task created: $taskName" -ForegroundColor Green
Write-Host "Time: Daily at 20:00 | Action: Excel report + email to spunk.chang@gmail.com" -ForegroundColor Cyan
schtasks /query /tn "$taskName"