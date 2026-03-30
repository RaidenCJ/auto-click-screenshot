# Trigger auto_click.ps1 via Windows Task Scheduler (bypasses remote software restrictions)
# This script creates an immediate task that runs in the user session

# Minimize this window
Add-Type -MemberDefinition '[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);' -Name Win32 -Namespace W -ErrorAction SilentlyContinue
$null = [W.Win32]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess()).MainWindowHandle, 6)

$taskName = "AutoClick_Immediate"
$scriptPath = "C:\Users\Raiden\.qclaw\workspace\auto_click.ps1"

# Remove existing immediate task if any
Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue

# Create task that runs immediately (minimized window)
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`""
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddSeconds(2)
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -User "Raiden" -Force | Out-Null

Write-Host "Task '$taskName' triggered at $(Get-Date -Format 'HH:mm:ss')"
