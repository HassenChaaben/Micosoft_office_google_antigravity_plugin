# Silent Office Add-in Unregistration Script
param (
    [string]$InstallDir = "$env:LOCALAPPDATA\GeminiAntigravityOffice"
)

$ErrorActionPreference = "SilentlyContinue"

$guid = "{7F89B2E3-8E5A-4B21-996A-12E4A9F200A1}"

# Remove WEF Developer Key
Remove-Item -Path "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer\$guid" -Recurse -Force

# Remove WEF Trusted Catalog
Remove-Item -Path "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$guid" -Recurse -Force

# Remove Startup Run Key
Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "GeminiAntigravityOffice" -Force

# Clean WEF ribbon cache and manifests
Get-ChildItem -Path "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef" -Recurse -Filter "*7f89b2e3*" -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
Get-ChildItem -Path "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef" -Recurse -Filter "*wa200099999*" -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue

# Kill running daemon process if alive
Get-Process -Name "gemini_office_daemon" | Stop-Process -Force

Write-Host "Unregistration completed successfully."
