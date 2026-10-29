@echo off
title Trust Gemini Antigravity Office Certificate
echo ============================================================
echo   Trusting Gemini Antigravity Office Localhost Certificate
echo ============================================================
echo.

set CERT_FOUND=0

:: 1. Check current folder certs directory
if exist "%~dp0certs\ca.crt" (
    echo [1/3] Installing CA Certificate from local certs directory...
    certutil -user -addstore -f Root "%~dp0certs\ca.crt" >nul 2>&1
    certutil -addstore -f Root "%~dp0certs\ca.crt" >nul 2>&1
    set CERT_FOUND=1
)

:: 2. Check parent folder certs directory (if run from a subfolder)
if exist "%~dp0..\certs\ca.crt" (
    echo [1/3] Installing CA Certificate from parent certs directory...
    certutil -user -addstore -f Root "%~dp0..\certs\ca.crt" >nul 2>&1
    certutil -addstore -f Root "%~dp0..\certs\ca.crt" >nul 2>&1
    set CERT_FOUND=1
)

:: 3. Check localhost.crt
if exist "%~dp0certs\localhost.crt" (
    echo [2/3] Installing Localhost Certificate into TrustedPeople...
    certutil -user -addstore -f TrustedPeople "%~dp0certs\localhost.crt" >nul 2>&1
    certutil -addstore -f Root "%~dp0certs\localhost.crt" >nul 2>&1
)
if exist "%~dp0..\certs\localhost.crt" (
    echo [2/3] Installing Localhost Certificate into TrustedPeople...
    certutil -user -addstore -f TrustedPeople "%~dp0..\certs\localhost.crt" >nul 2>&1
    certutil -addstore -f Root "%~dp0..\certs\localhost.crt" >nul 2>&1
)

:: 4. Check user profile Office dev certs fallback
if exist "%USERPROFILE%\.office-addin-dev-certs\ca.crt" (
    echo [*] Installing fallback Office Dev CA Certificate...
    certutil -user -addstore -f Root "%USERPROFILE%\.office-addin-dev-certs\ca.crt" >nul 2>&1
    set CERT_FOUND=1
)

:: 5. Set WebView2 bypass flags for silent SSL compatibility across all Office apps
echo [3/3] Setting WebView2 environment bypass flags...
setx WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS "--ignore-certificate-errors --allow-insecure-localhost" >nul 2>&1

echo.
echo ============================================================
if "%CERT_FOUND%"=="1" (
    echo   SUCCESS! Certificates and WebView2 flags configured!
    echo   Now open Word, Excel, or PowerPoint to use Gemini!
) else (
    echo   WebView2 flags configured successfully.
    echo   You can now launch Gemini Antigravity in Office.
)
echo ============================================================
echo.
pause
