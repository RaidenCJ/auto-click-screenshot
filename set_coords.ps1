# Set Click Coordinates - Simple Version
# Move mouse to each position, wait 4 seconds to capture

param(
    [switch]$Test  # If set, run test after capturing
)

Add-Type -AssemblyName System.Windows.Forms

# Ensure DPI awareness
try {
    $null = Add-Type -MemberDefinition '[DllImport("user32.dll")] public static extern bool SetProcessDPIAware();' -Name U32 -Namespace W -ErrorAction SilentlyContinue
    [W.U32]::SetProcessDPIAware()
} catch {}

$configPath = "$PSScriptRoot\auto_click_config.json"

Write-Host ""
Write-Host "=========================================="
Write-Host "  Auto Click Coordinate Picker"
Write-Host "=========================================="
Write-Host ""
Write-Host "INSTRUCTIONS:"
Write-Host "  1. Move your mouse to the target position"
Write-Host "  2. Keep mouse still for 3 seconds"
Write-Host "  3. Position will be captured automatically"
Write-Host ""
Write-Host "  (Run with -Test to auto-run after setting)"
Write-Host ""

$coords = @()
$positions = @(
    "Position 1 (main button - click only, NO screenshot)",
    "Position 2 (click + screenshot)",
    "Position 3 (click + screenshot)", 
    "Position 4 (click + screenshot)",
    "Position 5 (click only, NO screenshot)"
)

for ($i = 0; $i -lt 5; $i++) {
    Write-Host "[$($i+1)/5] $($positions[$i])"
    Write-Host "         Move mouse now... (capturing in 4 seconds)"
    
    Start-Sleep -Seconds 4
    
    $pos = [System.Windows.Forms.Cursor]::Position
    $coords += @{x = $pos.X; y = $pos.Y}
    
    Write-Host "         CAPTURED: ($($pos.X), $($pos.Y))"
    Write-Host ""
}

# Save to config
if (Test-Path $configPath) {
    $config = Get-Content $configPath | ConvertFrom-Json
    $config.coordinates = $coords
    $config.last_updated = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $config | ConvertTo-Json -Depth 10 | Set-Content $configPath -Encoding UTF8
    Write-Host "=========================================="
    Write-Host "  Config saved!"
    Write-Host "=========================================="
} else {
    Write-Host "ERROR: Config file not found at $configPath"
    exit 1
}

Write-Host ""
Write-Host "Coordinates saved:"
for ($i = 0; $i -lt 5; $i++) {
    Write-Host "  Position $($i+1): ($($coords[$i].x), $($coords[$i].y))"
}

# Test run if requested
if ($Test) {
    Write-Host ""
    Write-Host "=========================================="
    Write-Host "  Test will run in 10 seconds..."
    Write-Host "  (You can minimize this window)"
    Write-Host "=========================================="
    
    Start-Sleep -Seconds 10
    
    # Trigger via task scheduler
    $taskName = "AutoClick_Immediate"
    $scriptPath = "$PSScriptRoot\auto_click.ps1"
    
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
    
    $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -File `"$scriptPath`""
    $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddSeconds(2)
    $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
    
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -User "Raiden" -Force | Out-Null
    
    Write-Host ""
    Write-Host "Test task TRIGGERED at $(Get-Date -Format 'HH:mm:ss')"
    Write-Host "Check your Feishu for screenshots!"
}

Write-Host ""
