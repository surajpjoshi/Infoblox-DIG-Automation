$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectDir
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "       INFOBLOX DIG AUTOMATION - SETUP" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "[1/7] Checking Python..." -ForegroundColor Yellow
if (-not (Get-Command python -ErrorAction SilentlyContinue)) { Write-Host "Python not found. Install Python 3.11 or 3.12 and enable Add Python to PATH." -ForegroundColor Red; exit 1 }
python --version
Write-Host "[2/7] Checking project files..." -ForegroundColor Yellow
foreach ($file in @("app.py","commands.txt","requirements.txt")) { if (-not (Test-Path $file)) { Write-Host "Missing file: $file" -ForegroundColor Red; exit 1 } }
Write-Host "Project files OK." -ForegroundColor Green
Write-Host "[3/7] Creating virtual environment..." -ForegroundColor Yellow
if (-not (Test-Path ".venv")) { python -m venv .venv }
$PythonExe = Join-Path $ProjectDir ".venv\Scripts\python.exe"
if (-not (Test-Path $PythonExe)) { Write-Host "Virtual environment creation failed." -ForegroundColor Red; exit 1 }
Write-Host "[4/7] Upgrading pip..." -ForegroundColor Yellow
& $PythonExe -m pip install --upgrade pip
Write-Host "[5/7] Installing required packages..." -ForegroundColor Yellow
& $PythonExe -m pip install -r requirements.txt
Write-Host "[6/7] Checking PuTTYgen..." -ForegroundColor Yellow
$puttygen = $null
foreach ($path in @("$env:ProgramFiles\PuTTY\puttygen.exe","${env:ProgramFiles(x86)}\PuTTY\puttygen.exe","$env:LOCALAPPDATA\Programs\PuTTY\puttygen.exe")) { if ($path -and (Test-Path $path)) { $puttygen=$path; break } }
if (-not $puttygen) { $cmd=Get-Command puttygen.exe -ErrorAction SilentlyContinue; if ($cmd) { $puttygen=$cmd.Source } }
if ($puttygen) { Write-Host "PuTTYgen found: $puttygen" -ForegroundColor Green } else { Write-Host "PuTTYgen not found. PPK authentication requires PuTTYgen." -ForegroundColor Yellow }
Write-Host "[7/7] Testing Python modules..." -ForegroundColor Yellow
& $PythonExe -c "import customtkinter, paramiko, openpyxl; print('Python modules: OK')"
if ($LASTEXITCODE -ne 0) { Write-Host "Module verification failed." -ForegroundColor Red; exit 1 }
Write-Host "SETUP COMPLETE" -ForegroundColor Green
$runNow=Read-Host "Launch Infoblox DIG Automation now? (Y/N)"
if ($runNow -match '^[Yy]$') { & $PythonExe ".\app.py" } else { Write-Host "Launch later with: .\.venv\Scripts\python.exe .\app.py" -ForegroundColor Cyan }
