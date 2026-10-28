<#
.SYNOPSIS
  One-time setup on a Windows laptop: Python venv, first sign-in, daily scheduled task.
.EXAMPLE
  .\setup-windows.ps1 -UsageDir "$env:USERPROFILE\Contoso\Team - Documents\Usage" -CrSite "contoso.sharepoint.com:/sites/Team"
#>
param(
  [Parameter(Mandatory)] [string] $UsageDir,
  [string] $CrSite = '',
  [string] $At = '09:00'
)
$ErrorActionPreference = 'Stop'
if (-not (Test-Path $UsageDir)) { throw "Usage folder not found: $UsageDir (is the Teams folder synced?)" }
$home_ = Join-Path $env:LOCALAPPDATA 'copilot-credits'
New-Item -ItemType Directory -Force -Path $home_ | Out-Null

$py = Get-Command py, python -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $py) {
  winget install -e --id Python.Python.3.12 --scope user --accept-source-agreements --accept-package-agreements
  $env:Path = [Environment]::GetEnvironmentVariable('Path', 'User') + ';' + $env:Path
  $py = Get-Command py, python | Select-Object -First 1
}
& $py.Source -m venv (Join-Path $home_ 'venv')
& (Join-Path $home_ 'venv\Scripts\python.exe') -m pip install -q -r "$PSScriptRoot\requirements.txt"

Write-Host 'First run: sign in with the device code below (admin account).'
& "$PSScriptRoot\run.ps1" -UsageDir $UsageDir -CrSite $CrSite -Interactive

$arg = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PSScriptRoot\run.ps1`" -UsageDir `"$UsageDir`" -CrSite `"$CrSite`""
$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $arg
$trigger = New-ScheduledTaskTrigger -Daily -At $At
# StartWhenAvailable: if the laptop was off/asleep at $At, run as soon as it is back.
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
  -ExecutionTimeLimit (New-TimeSpan -Hours 3)  # audit searches can take an hour each
Register-ScheduledTask -TaskName 'Copilot credits report' -Action $action -Trigger $trigger -Settings $settings -Force | Out-Null
Write-Host "Scheduled daily at $At (task 'Copilot credits report'). Log: $home_\run.log"
