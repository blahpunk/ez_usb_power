param(
    [string]$Version = "0.1.0"
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot

Write-Host "[1/5] Installing dependencies..."
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyinstaller

Write-Host "[2/5] Cleaning old build artifacts..."
if (Test-Path ".\build") { Remove-Item ".\build" -Recurse -Force }
if (Test-Path ".\dist") { Remove-Item ".\dist" -Recurse -Force }
if (Test-Path ".\release") { Remove-Item ".\release" -Recurse -Force }
if (Test-Path ".\EZ_USB_Power.spec") { Remove-Item ".\EZ_USB_Power.spec" -Force }

Write-Host "[3/5] Building standalone EXE..."
python -m PyInstaller `
    --noconfirm `
    --clean `
    --windowed `
    --onefile `
    --name "EZ_USB_Power" `
    .\usb_power_gui.py

Write-Host "[4/5] Packaging release..."
New-Item -ItemType Directory -Path ".\release" | Out-Null
$ExePath = ".\dist\EZ_USB_Power.exe"
if (-not (Test-Path $ExePath)) {
    throw "Expected build output not found: $ExePath"
}

$ZipName = "EZ_USB_Power_v$Version`_win64.zip"
$ZipPath = Join-Path ".\release" $ZipName
Compress-Archive -Path $ExePath -DestinationPath $ZipPath -Force

Write-Host "[5/5] Done"
Write-Host "EXE: $ExePath"
Write-Host "ZIP: $ZipPath"
