# Master Automated Build Script for Gemini Antigravity Office Plugin
# Builds assets, compiles standalone daemon binary, and packages Inno Setup .exe installer

$ErrorActionPreference = "Stop"

Write-Host "=========================================================="
Write-Host "  Gemini Antigravity Office Plugin - Master Build Pipeline"
Write-Host "=========================================================="

$projectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $projectRoot

# Step 1: Generate Branding & Icons
Write-Host "`n[1/5] Generating vector & bitmap assets..."
python scripts/generate_assets.py

# Step 2: Generate Localhost SSL Certificate
Write-Host "`n[2/5] Generating localhost TLS certificate..."
python daemon/cert_gen.py

# Step 3: Compile Standalone Daemon Binary via PyInstaller
Write-Host "`n[3/5] Compiling standalone daemon via PyInstaller..."
python -m PyInstaller --noconfirm gemini_office_daemon.spec

# Step 4: Verify Daemon Executable
$daemonExe = Join-Path $projectRoot "dist\gemini_office_daemon\gemini_office_daemon.exe"
if (-not (Test-Path $daemonExe)) {
    throw "PyInstaller build failed: $daemonExe not found!"
}
Write-Host "Daemon compiled successfully: $daemonExe"

# Step 5: Compile Inno Setup Installer
Write-Host "`n[4/5] Compiling Inno Setup standalone installer..."
$isccPath = "C:\Users\user\AppData\Local\Programs\Inno Setup 6\ISCC.exe"
if (-not (Test-Path $isccPath)) {
    throw "Inno Setup compiler not found at $isccPath"
}

$issFile = Join-Path $projectRoot "installer\setup.iss"
& $isccPath /Qp $issFile

# Step 6: Verify Final Installer Output
Write-Host "`n[5/5] Verifying output installer..."
$setupExe = Join-Path $projectRoot "dist\GeminiAntigravityOffice_Setup.exe"
if (Test-Path $setupExe) {
    $item = Get-Item $setupExe
    $mbSize = [math]::Round($item.Length / 1MB, 2)
    Write-Host "=========================================================="
    Write-Host "  BUILD SUCCESSFUL!"
    Write-Host "  Installer: $setupExe"
    Write-Host "  Size: $mbSize MB"
    Write-Host "=========================================================="
} else {
    throw "Build failed: $setupExe was not created!"
}
