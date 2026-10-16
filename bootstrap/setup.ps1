#Requires -Version 5.1
param(
    [Alias('R')]
    [string]$Root = (Split-Path -Parent $PSScriptRoot)
)

$ErrorActionPreference = 'Stop'

$Root = (Resolve-Path $Root).Path
$versionFile = Join-Path $Root '.python-version'
if (-not (Test-Path $versionFile)) {
    throw "Missing .python-version in $Root"
}
$PyVersion = (Get-Content $versionFile -Raw).Trim()

function Invoke-Python {
    param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Args)
    if (Get-Command py -ErrorAction SilentlyContinue) {
        & py "-$PyVersion" @Args
    } else {
        & python @Args
    }
    if ($LASTEXITCODE -ne 0) { throw "python failed: $Args" }
}

$dirs = @(
    (Join-Path $Root 'data'),
    (Join-Path $Root 'output'),
    (Join-Path $Root 'logs'),
    (Join-Path $Root 'scripts'),
    (Join-Path $Root 'src\finance_tools')
)
foreach ($d in $dirs) {
    New-Item -ItemType Directory -Force -Path $d | Out-Null
}

$requirements = @'
pandas>=2.2
openpyxl>=3.1
'@
$pathsPy = @'
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DATA = ROOT / "data"
OUTPUT = ROOT / "output"
LOGS = ROOT / "logs"
SCRIPTS = ROOT / "scripts"

# Shared network folders for monthly expense-review mail
EMAIL_ZONE = Path(r"R:\費用檢討E-MAIL專區")
RECIPIENT_DIR = EMAIL_ZONE / "E-MIAL"
ATTACHMENT_DIR = EMAIL_ZONE / "夾檔"
CONTENT_FILE = EMAIL_ZONE / "MAIL內文.xlsx"


def ensure_dirs() -> None:
    for p in (DATA, OUTPUT, LOGS):
        p.mkdir(parents=True, exist_ok=True)
'@

$reqPath = Join-Path $Root 'requirements.txt'
if (-not (Test-Path $reqPath)) {
    Set-Content -Path $reqPath -Value $requirements -Encoding UTF8
}
Set-Content -Path (Join-Path $Root 'src\finance_tools\paths.py') -Value $pathsPy -Encoding UTF8
$initPath = Join-Path $Root 'src\finance_tools\__init__.py'
if (-not (Test-Path $initPath)) {
    New-Item -ItemType File -Force -Path $initPath | Out-Null
}

$venv = Join-Path $Root '.venv'
if (-not (Test-Path $venv)) {
    Invoke-Python -Args @('-m', 'venv', $venv)
}

$venvPython = Join-Path $venv 'Scripts\python.exe'
& $venvPython -m pip install -U pip
& $venvPython -m pip install -r $reqPath

$env:PYTHONPATH = Join-Path $Root 'src'
& $venvPython -c "from finance_tools.paths import ensure_dirs; ensure_dirs()"

Write-Host "Ready: $Root"
Write-Host "venv:  $venv\Scripts\Activate.ps1"
