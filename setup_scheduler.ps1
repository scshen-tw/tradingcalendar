# CB Calendar - Windows Task Scheduler Setup
# Run as Administrator

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$BatFile   = Join-Path $ScriptDir "auto_update.bat"
$TaskName  = "CB-Calendar-AutoUpdate"

Write-Host "=== CB Calendar Scheduler Setup ===" -ForegroundColor Cyan
Write-Host "Bat file: $BatFile"

if (-not (Test-Path $BatFile)) {
    Write-Host "ERROR: auto_update.bat not found at $BatFile" -ForegroundColor Red
    exit 1
}

# Remove old task if exists
$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existing) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "Removed old task."
}

# Action
$Action = New-ScheduledTaskAction `
    -Execute "cmd.exe" `
    -Argument "/c `"$BatFile`"" `
    -WorkingDirectory $ScriptDir

# Triggers: 08:00 / 13:30 / 19:00 daily
$Trigger1 = New-ScheduledTaskTrigger -Daily -At "08:00"
$Trigger2 = New-ScheduledTaskTrigger -Daily -At "13:30"
$Trigger3 = New-ScheduledTaskTrigger -Daily -At "19:00"

# Settings
$Settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 5) `
    -MultipleInstances IgnoreNew `
    -StartWhenAvailable

$Principal = New-ScheduledTaskPrincipal `
    -UserId ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) `
    -LogonType Interactive `
    -RunLevel Highest

# Register
Register-ScheduledTask `
    -TaskName   $TaskName `
    -Action     $Action `
    -Trigger    $Trigger1, $Trigger2, $Trigger3 `
    -Settings   $Settings `
    -Principal  $Principal `
    -Description "CB Calendar auto-update at 08:00 / 13:30 / 19:00" `
    -Force | Out-Null

Write-Host ""
Write-Host "OK: Task created -> $TaskName" -ForegroundColor Green
Write-Host "    Schedule: 08:00 / 13:30 / 19:00 daily"
Write-Host "    Bat: $BatFile"
Write-Host ""
Write-Host "To verify: open Task Scheduler -> Task Scheduler Library -> $TaskName"
