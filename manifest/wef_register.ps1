# Silent Office Add-in Registration and SSL Certificate Trust Script
param (
    [string]$InstallDir = "$env:LOCALAPPDATA\GeminiAntigravityOffice"
)

$ErrorActionPreference = "Continue"

Write-Host "=========================================================="
Write-Host "  Registering Gemini Antigravity Office Add-in"
Write-Host "  Install Directory: $InstallDir"
Write-Host "=========================================================="

$guid = "{7F89B2E3-8E5A-4B21-996A-12E4A9F200A1}"
$manifestPath = Join-Path $InstallDir "manifest.xml"
$certPath = Join-Path $InstallDir "certs\localhost.crt"

# 0. Sync fresh manifest and assets if running from repository root
$scriptDir = $PSScriptRoot
if (-not $scriptDir) {
    if ($MyInvocation.MyCommand.Definition) {
        $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
    } else {
        $scriptDir = (Get-Location).Path
    }
}
$repoRoot = Split-Path -Parent $scriptDir
if ((Test-Path (Join-Path $repoRoot "manifest\manifest.xml")) -and ($InstallDir -ne $repoRoot)) {
    Write-Host "[Sync] Syncing latest manifest and assets from repository to $InstallDir..."
    if (-not (Test-Path $InstallDir)) {
        New-Item -Path $InstallDir -ItemType Directory -Force | Out-Null
    }
    Copy-Item -Path (Join-Path $repoRoot "manifest\manifest.xml") -Destination (Join-Path $InstallDir "manifest.xml") -Force -ErrorAction SilentlyContinue
    if (Test-Path (Join-Path $repoRoot "assets")) {
        if (-not (Test-Path (Join-Path $InstallDir "assets"))) {
            New-Item -Path (Join-Path $InstallDir "assets") -ItemType Directory -Force | Out-Null
        }
        Copy-Item -Path (Join-Path $repoRoot "assets\*") -Destination (Join-Path $InstallDir "assets\") -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# 1. Clean stale Office WEF ribbon and manifest cache
Write-Host "[Cache] Cleaning stale Office WEF and AppCommands cache..."
Get-ChildItem -Path "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef" -Recurse -Filter "*7f89b2e3*" -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
Get-ChildItem -Path "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef" -Recurse -Filter "*wa200099999*" -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue

# 1. Install & Trust Localhost Certificate
if (Test-Path $certPath) {
    Write-Host "[Cert] Trusting localhost SSL certificate in CurrentUser\TrustedPeople..."
    & certutil -addstore -f -user TrustedPeople "$certPath" | Out-Null
    
    # Also attempt Root store if running elevated
    try {
        & certutil -addstore -f Root "$certPath" 2>$null | Out-Null
    } catch {}
}

# 2. Register Office WEF Developer Key (Word & PowerPoint instant sideload)
Write-Host "[Registry] Writing WEF Developer Add-in Key..."
$devKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer\$guid"
if (-not (Test-Path $devKey)) {
    New-Item -Path $devKey -Force | Out-Null
}
Set-ItemProperty -Path $devKey -Name "ManifestPath" -Value $manifestPath -Force
Set-ItemProperty -Path $devKey -Name "UseResourceTokens" -Type DWord -Value 0 -Force

# 3. Register Trusted Catalog
Write-Host "[Registry] Registering WEF Trusted Catalog..."
$catKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$guid"
if (-not (Test-Path $catKey)) {
    New-Item -Path $catKey -Force | Out-Null
}
Set-ItemProperty -Path $catKey -Name "Url" -Value $InstallDir -Force
Set-ItemProperty -Path $catKey -Name "Flags" -Type DWord -Value 1 -Force

# 4. Register Windows Startup Run key for Background Daemon
Write-Host "[Registry] Registering Windows Startup Daemon..."
$runKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
$daemonExe = Join-Path $InstallDir "gemini_office_daemon.exe"
if (Test-Path $daemonExe) {
    Set-ItemProperty -Path $runKey -Name "GeminiAntigravityOffice" -Value "`"$daemonExe`" --background" -Force
}

# 5. Direct WEF AppCommands Ribbon Cache Injection (Word, Excel & PowerPoint across all locales)
Write-Host "[Ribbon] Injecting into Office Ribbon Cache across all installed locales..."
$baseAppCommands = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef\AppCommands"
if (Test-Path $baseAppCommands) {
    $locales = @("en-US", "fr-FR", "it-IT", "de-DE", "es-ES")
    $sysCulture = (Get-UICulture).Name
    if ($sysCulture -and -not ($locales -contains $sysCulture)) {
        $locales += $sysCulture
    }

    $cacheFiles = Get-ChildItem -Path $baseAppCommands -Filter "*.RibbonCache.*" -Recurse -File -ErrorAction SilentlyContinue
    foreach ($cf in $cacheFiles) {
        if ($cf.Name -match "\.RibbonCache\.(.+)$") {
            $foundLoc = $Matches[1]
            if (-not ($locales -contains $foundLoc)) {
                $locales += $foundLoc
            }
        }
    }

    $targetDirs = Get-ChildItem -Path $baseAppCommands -Directory -Recurse -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
    if (-not $targetDirs -or $targetDirs.Count -eq 0) {
        $targetDirs = @("$baseAppCommands\18.0")
        New-Item -Path "$baseAppCommands\18.0" -ItemType Directory -Force | Out-Null
    }

    $versions = @("16.0", "18.0")
    foreach ($dir in $targetDirs) {
        foreach ($loc in $locales) {
            foreach ($ver in $versions) {
                $wordCmd = "$ver 0 $loc WA200099999 WA200099999 wa200099999 1:0:0:1$loc 0 https://localhost:18890/index.html Gemini Gemini Antigravity in Microsoft Office 1:PrimaryCommandSurface:2:2 PrimaryCommandSurface 1:4 TabHome 1:3 GeminiCommandsGroup.Word 1:1 GeminiTaskpaneButton.Word 0:Gemini  Gemini Gemini 3:16:https://localhost:18890/assets/icon-gemini-16.png 32:https://localhost:18890/assets/icon-gemini-32.png 80:https://localhost:18890/assets/icon-gemini-80.png {""enabled"":true,""apioverride"":false,""visible"":true}   2:4:2:https://localhost:18890/index.html 3:Office.AutoShowTaskpaneWithDocument 4: 8:Taskpane.Url Gemini   5:16:https://localhost:18890/assets/icon-gemini-16.png 32:https://localhost:18890/assets/icon-gemini-32.png 32:https://localhost:18890/assets/icon-gemini-32.png 64:https://localhost:18890/assets/icon-gemini-64.png 80:https://localhost:18890/assets/icon-gemini-80.png {""enabled"":true,""apioverride"":false,""visible"":true}      0:   0:5:https://localhost:18890/assets/icon-gemini-32.png https://localhost:18890/assets/icon-gemini-64.png https://localhost:18890/assets/icon-gemini-16.png https://localhost:18890/assets/icon-gemini-32.png https://localhost:18890/assets/icon-gemini-80.png 0:`n"
                $pptCmd = "$ver 0 $loc WA200099999 WA200099999 wa200099999 1:0:0:1$loc 0 https://localhost:18890/index.html Gemini Gemini Antigravity in Microsoft Office 1:PrimaryCommandSurface:2:2 PrimaryCommandSurface 1:4 TabHome 1:3 GeminiCommandsGroup.PowerPoint 1:1 GeminiTaskpaneButton.PowerPoint 0:Gemini  Gemini Gemini 3:16:https://localhost:18890/assets/icon-gemini-16.png 32:https://localhost:18890/assets/icon-gemini-32.png 80:https://localhost:18890/assets/icon-gemini-80.png {""enabled"":true,""apioverride"":false,""visible"":true}   2:4:2:https://localhost:18890/index.html 3:Office.AutoShowTaskpaneWithDocument 4: 8:Taskpane.Url Gemini   5:16:https://localhost:18890/assets/icon-gemini-16.png 32:https://localhost:18890/assets/icon-gemini-32.png 32:https://localhost:18890/assets/icon-gemini-32.png 64:https://localhost:18890/assets/icon-gemini-64.png 80:https://localhost:18890/assets/icon-gemini-80.png {""enabled"":true,""apioverride"":false,""visible"":true}      0:   0:5:https://localhost:18890/assets/icon-gemini-32.png https://localhost:18890/assets/icon-gemini-64.png https://localhost:18890/assets/icon-gemini-16.png https://localhost:18890/assets/icon-gemini-32.png https://localhost:18890/assets/icon-gemini-80.png 0:`n"
                $excelCmd = "$ver 0 $loc WA200099999 WA200099999 wa200099999 1:0:0:1$loc 0 https://localhost:18890/index.html Gemini Gemini Antigravity in Microsoft Office 1:PrimaryCommandSurface:2:2 PrimaryCommandSurface 1:4 TabHome 1:3 GeminiCommandsGroup.Excel 1:1 GeminiTaskpaneButton.Excel 0:Gemini  Gemini Gemini 3:16:https://localhost:18890/assets/icon-gemini-16.png 32:https://localhost:18890/assets/icon-gemini-32.png 80:https://localhost:18890/assets/icon-gemini-80.png {""enabled"":true,""apioverride"":false,""visible"":true}   2:4:2:https://localhost:18890/index.html 3:Office.AutoShowTaskpaneWithDocument 4: 8:Taskpane.Url Gemini   5:16:https://localhost:18890/assets/icon-gemini-16.png 32:https://localhost:18890/assets/icon-gemini-32.png 32:https://localhost:18890/assets/icon-gemini-32.png 64:https://localhost:18890/assets/icon-gemini-64.png 80:https://localhost:18890/assets/icon-gemini-80.png {""enabled"":true,""apioverride"":false,""visible"":true}      0:   0:5:https://localhost:18890/assets/icon-gemini-32.png https://localhost:18890/assets/icon-gemini-64.png https://localhost:18890/assets/icon-gemini-16.png https://localhost:18890/assets/icon-gemini-32.png https://localhost:18890/assets/icon-gemini-80.png 0:`n"

                Set-Content -Path (Join-Path $dir "wa200099999${ver}Word${loc}") -Value $wordCmd -Encoding UTF8 -ErrorAction SilentlyContinue
                Set-Content -Path (Join-Path $dir "wa200099999${ver}PowerPoint${loc}") -Value $pptCmd -Encoding UTF8 -ErrorAction SilentlyContinue
                Set-Content -Path (Join-Path $dir "wa200099999${ver}Excel${loc}") -Value $excelCmd -Encoding UTF8 -ErrorAction SilentlyContinue
            }

            $appCacheList = @(
                @{ File = "Word.RibbonCache.$loc" },
                @{ File = "PowerPoint.RibbonCache.$loc" },
                @{ File = "Excel.RibbonCache.$loc" }
            )
            foreach ($cacheItem in $appCacheList) {
                $cPath = Join-Path $dir $cacheItem.File
                if (Test-Path $cPath) {
                    $content = Get-Content $cPath -Raw -ErrorAction SilentlyContinue
                    if ($content -notmatch "WA200099999") {
                        Add-Content -Path $cPath -Value " 1:WA200099999+Omex 1:0 Omex 1:Omex WA200099999 WA200099999 1:0:0:1$loc " -NoNewline -ErrorAction SilentlyContinue
                    }
                }
            }
        }
    }
}

Write-Host "Registration completed successfully. Word, Excel, and PowerPoint will now load Gemini Antigravity automatically!"
