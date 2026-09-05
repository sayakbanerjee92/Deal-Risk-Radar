<# Runs the Deal Contract Risk Radar locally with Python only. #>
[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

if (-not (Get-Command py -ErrorAction SilentlyContinue)) {
    throw "Python 3.11+ was not found. Install it from https://www.python.org/downloads/ and reopen VS Code."
}

$projectRoot = Split-Path -Parent $PSCommandPath
Set-Location -LiteralPath $projectRoot

if (-not (Test-Path -LiteralPath ".venv\Scripts\python.exe")) {
    Write-Host "Creating Python virtual environment…" -ForegroundColor Cyan
    & py -3.11 -m venv .venv
}

$python = Join-Path $projectRoot ".venv\Scripts\python.exe"
Write-Host "Installing local Python dependencies…" -ForegroundColor Cyan
& $python -m pip install --upgrade pip
& $python -m pip install --upgrade -r requirements.txt

Write-Host "Starting Deal Contract Risk Radar…" -ForegroundColor Green
& $python -m streamlit run app.py

