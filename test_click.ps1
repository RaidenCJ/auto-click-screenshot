# Test Auto Click - Run after 10 seconds
# Triggers the auto_click.ps1 script via Task Scheduler

Write-Host ""
Write-Host "=========================================="
Write-Host "  Auto Click TEST Run"
Write-Host "=========================================="
Write-Host ""
Write-Host "Test will run in 10 seconds..."
Write-Host "You can minimize this window or close remote connection."
Write-Host ""

Start-Sleep -Seconds 10

# Trigger via task scheduler
$taskName = "AutoClick_Immediate"
$scriptPath = Split-Path $PSScriptRoot -Parent | Join-Path -ChildPath "auto_click.ps1"

# Fallback: use known path
if (-not (Test-Path $scriptPath)) {
    $scriptPath = "C:\Users\Raiden\.qclaw\workspace\auto_click.ps1"
}

Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue

$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -File `"$scriptPath`""
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddSeconds(2)
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -User "Raiden" -Force | Out-Null

Write-Host "Test task TRIGGERED at $(Get-Date -Format 'HH:mm:ss')"
Write-Host "Check your Feishu for screenshots!"
Write-Host ""
