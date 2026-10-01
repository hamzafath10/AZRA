<#
.SYNOPSIS
    Planifie MacroBrief tous les jours (par défaut à 07:30, du lundi au vendredi).
.EXAMPLE
    .\Install-MacroBriefTask.ps1 -Time 08:00
.EXAMPLE
    .\Install-MacroBriefTask.ps1 -Remove
#>
param(
    [string]$Time = '07:30',
    [switch]$Weekend,
    [switch]$Remove
)
$TaskName = 'MacroBrief'

if ($Remove) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Tâche '$TaskName' supprimée." -ForegroundColor Yellow
    return
}

$script = Join-Path $PSScriptRoot 'MacroBrief.ps1'
$action = New-ScheduledTaskAction -Execute 'powershell.exe' `
    -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$script`" -Quiet"
$days = 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'
if ($Weekend) { $days += 'Saturday', 'Sunday' }
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $days -At $Time
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings `
    -Description 'Briefing macro / marchés quotidien (rapport HTML ouvert dans le navigateur)' -Force | Out-Null

Write-Host "Tâche '$TaskName' planifiée à $Time ($($days -join ', '))." -ForegroundColor Green
Write-Host "Le rapport s'ouvrira automatiquement dans ton navigateur." -ForegroundColor DarkGray
