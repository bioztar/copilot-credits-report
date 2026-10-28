<#
.SYNOPSIS
  Pull today's Copilot credits/usage from Graph, then rebuild the workbook + dashboard.
  Called by the scheduled task; use -Interactive once to sign in (prints a device code).
#>
param(
  [Parameter(Mandatory)] [string] $UsageDir,
  [string] $CrSite = '',
  [switch] $Interactive
)
$ErrorActionPreference = 'Continue'  # PS 5.1 turns native stderr into errors when redirected; rely on exit codes
$home_ = Join-Path $env:LOCALAPPDATA 'copilot-credits'
$py = Join-Path $home_ 'venv\Scripts\python.exe'
$log = Join-Path $home_ 'run.log'
$env:CR_SITE = $CrSite

if ($Interactive) {
  & $py "$PSScriptRoot\pull_credits.py" $UsageDir
} else {
  "`n=== $(Get-Date -Format s)" | Add-Content $log
  & $py "$PSScriptRoot\pull_credits.py" $UsageDir --unattended *>> $log
}
if ($LASTEXITCODE -ne 0) { throw "pull_credits failed (exit $LASTEXITCODE), see $log" }
& $py "$PSScriptRoot\build_report.py" $UsageDir *>> $log
if ($LASTEXITCODE -ne 0) { throw "build_report failed (exit $LASTEXITCODE), see $log" }

# Purview audit log (non-fatal: a search can take up to an hour, credits must not wait for it).
#  1. Cowork detail: last 8 days, merged into cowork_audit/cowork_events.csv, then its HTML report.
#  2. All Copilot / AI activity for yesterday -> ai_activity/daily_<day>_<day+1>.csv, then the cost-over-time report.
# The interactive run asks for the audit permission once (second device code).
if ($Interactive) { & $py "$PSScriptRoot\cowork_audit.py" $UsageDir --days 1 }  # shows the device code on screen
$today = (Get-Date).ToString('yyyy-MM-dd'); $yday = (Get-Date).AddDays(-1).ToString('yyyy-MM-dd')
& $py "$PSScriptRoot\cowork_audit.py" $UsageDir --days 8 --unattended *>> $log
if ($LASTEXITCODE -eq 0) { & $py "$PSScriptRoot\cowork_audit_report.py" $UsageDir *>> $log }
else { "cowork_audit failed (exit $LASTEXITCODE), credits report unaffected" | Add-Content $log }
if (-not (Test-Path (Join-Path $UsageDir "ai_activity\daily_$($yday -replace '-')_$($today -replace '-').csv"))) {
  & $py "$PSScriptRoot\ai_activity.py" $UsageDir --from $yday --to $today --unattended *>> $log
  if ($LASTEXITCODE -ne 0) { "ai_activity failed (exit $LASTEXITCODE)" | Add-Content $log }
}
& $py "$PSScriptRoot\ai_cost_report.py" $UsageDir *>> $log
