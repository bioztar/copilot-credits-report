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

# Cowork activity from the Purview audit log (last 8 days, merged into cowork_events.csv), then its HTML report.
# Non-fatal: the audit search can take up to an hour, and credits must not wait for it.
if (-not $Interactive) {
  & $py "$PSScriptRoot\cowork_audit.py" $UsageDir --days 8 --unattended *>> $log
  if ($LASTEXITCODE -eq 0) { & $py "$PSScriptRoot\cowork_audit_report.py" $UsageDir *>> $log }
  else { "cowork_audit failed (exit $LASTEXITCODE), credits report unaffected" | Add-Content $log }
}
