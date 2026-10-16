#Requires -Version 5.1
<#
.SYNOPSIS
  Register monthly Task Scheduler job (day 15, 09:00) for expense-review mail.
.EXAMPLE
  .\scripts\register_expense_review_mail_task.ps1
  .\scripts\register_expense_review_mail_task.ps1 -Prod
#>
param(
    [string]$Root = (Split-Path -Parent $PSScriptRoot),
    [switch]$Prod,
    [string]$TaskName = "Finance-ExpenseReviewMail"
)

$ErrorActionPreference = "Stop"
$Root = (Resolve-Path $Root).Path
$python = Join-Path $Root ".venv\Scripts\python.exe"
$script = Join-Path $Root "scripts\send_expense_review_mail.py"

if (-not (Test-Path $python)) { throw "Missing venv python: $python" }
if (-not (Test-Path $script)) { throw "Missing script: $script" }

$modeFlag = if ($Prod) { "--prod" } else { "--test" }
# schtasks needs the whole command quoted as one TR argument
$tr = "`"$python`" `"$script`" $modeFlag"

# Monthly on day 15 at 09:00 (interactive logon so network/R: drive is available)
$createArgs = @(
    "/Create", "/F",
    "/TN", $TaskName,
    "/TR", $tr,
    "/SC", "MONTHLY",
    "/D", "15",
    "/ST", "09:00",
    "/RL", "LIMITED",
    "/IT"
)

& schtasks.exe @createArgs
if ($LASTEXITCODE -ne 0) {
    throw "schtasks failed with exit code $LASTEXITCODE"
}

Write-Host "Registered task: $TaskName"
Write-Host "Mode: $(if ($Prod) { 'PROD' } else { 'TEST (mis@ only)' })"
Write-Host "Schedule: monthly day 15, 09:00"
Write-Host "Requires secrets\mail_smtp.env and R: drive access."
Write-Host "Run once now:  schtasks /Run /TN `"$TaskName`""
