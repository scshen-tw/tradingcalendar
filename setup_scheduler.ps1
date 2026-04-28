# Trading Calendar - Windows Task Scheduler setup
#
# Run this from a normal logged-in desktop session. Outlook automation requires
# the same interactive user session as Outlook, so these tasks intentionally do
# not use RunLevel Highest.

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

function Register-CalendarTask {
    param(
        [Parameter(Mandatory = $true)][string]$TaskName,
        [Parameter(Mandatory = $true)][string]$BatName,
        [Parameter(Mandatory = $true)][string[]]$Times,
        [Parameter(Mandatory = $true)][string]$Description
    )

    $BatFile = Join-Path $ScriptDir $BatName
    Write-Host "Task: $TaskName" -ForegroundColor Cyan
    Write-Host "Bat:  $BatFile"

    if (-not (Test-Path $BatFile)) {
        Write-Host "ERROR: $BatName not found." -ForegroundColor Red
        exit 1
    }

    $existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($existing) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Host "Removed old task."
    }

    $action = New-ScheduledTaskAction `
        -Execute "cmd.exe" `
        -Argument "/c `"$BatFile`"" `
        -WorkingDirectory $ScriptDir

    $triggers = foreach ($time in $Times) {
        New-ScheduledTaskTrigger -Daily -At $time
    }

    $settings = New-ScheduledTaskSettingsSet `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 10) `
        -MultipleInstances IgnoreNew `
        -StartWhenAvailable

    $principal = New-ScheduledTaskPrincipal `
        -UserId ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) `
        -LogonType Interactive `
        -RunLevel Limited

    Register-ScheduledTask `
        -TaskName $TaskName `
        -Action $action `
        -Trigger $triggers `
        -Settings $settings `
        -Principal $principal `
        -Description $Description `
        -Force | Out-Null

    Write-Host "OK: $TaskName -> $($Times -join ', ')" -ForegroundColor Green
    Write-Host ""
}

Write-Host "=== Trading Calendar Scheduler Setup ===" -ForegroundColor Cyan
Write-Host "Directory: $ScriptDir"
Write-Host ""

Register-CalendarTask `
    -TaskName "TradingCalendar-CB-OutlookUpdate" `
    -BatName "auto_update.bat" `
    -Times @("08:00", "13:30", "19:00") `
    -Description "Update CB events from Outlook cbas mail and stock auction events."

Register-CalendarTask `
    -TaskName "TradingCalendar-StockUpdate" `
    -BatName "update_stocks.bat" `
    -Times @("09:10") `
    -Description "Update stock auction events without Outlook."

Write-Host "Done. Keep Outlook open for the CB task." -ForegroundColor Green
Write-Host "Verify in Task Scheduler Library:"
Write-Host "  TradingCalendar-CB-OutlookUpdate"
Write-Host "  TradingCalendar-StockUpdate"
