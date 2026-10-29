; =====================================================================
; Gemini Antigravity Office Plugin - Turnkey Enterprise Client Installer
; Generates 100% automated, zero-config standalone Windows installer:
; GeminiAntigravityOffice_Setup.exe
; =====================================================================

#define MyAppName "Gemini Antigravity Office"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Google Antigravity"
#define MyAppURL "https://github.com/HassenChaaben"
#define MyAppExeName "gemini_office_daemon.exe"
#define MyAppGuid "7F89B2E3-8E5A-4B21-996A-12E4A9F200A1"

[Setup]
AppId={{7F89B2E3-8E5A-4B21-996A-12E4A9F200A1}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={localappdata}\GeminiAntigravityOffice
DisableDirPage=yes
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=..\dist
OutputBaseFilename=GeminiAntigravityOffice_Setup
SetupIconFile=..\assets\favicon.ico
WizardImageFile=..\assets\installer_banner.bmp
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
CloseApplications=yes
RestartApplications=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
; Standalone Daemon Executable and dependencies from PyInstaller
Source: "..\dist\gemini_office_daemon\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

; Web Taskpane Assets & HTML
Source: "..\web\*"; DestDir: "{app}\web"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\assets\*"; DestDir: "{app}\assets"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\certs\*"; DestDir: "{app}\certs"; Flags: ignoreversion recursesubdirs createallsubdirs

; Pre-configured One-Click Launch Documents
Source: "..\assets\Gemini_Antigravity_Word.docx"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\assets\Gemini_Antigravity_PowerPoint.pptx"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\assets\Gemini_Antigravity_Excel.xlsx"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\installer\Trust_Gemini_Certificate.cmd"; DestDir: "{app}"; Flags: ignoreversion

; Office Add-in Manifest
Source: "..\manifest\manifest.xml"; DestDir: "{app}"; Flags: ignoreversion

; PowerShell Helper Scripts for Sideload & Unregister
Source: "..\manifest\wef_register.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\manifest\wef_unregister.ps1"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{userdesktop}\Gemini Antigravity Word"; Filename: "{app}\Gemini_Antigravity_Word.docx"; IconFilename: "{app}\assets\favicon.ico"
Name: "{userdesktop}\Gemini Antigravity PowerPoint"; Filename: "{app}\Gemini_Antigravity_PowerPoint.pptx"; IconFilename: "{app}\assets\favicon.ico"
Name: "{userdesktop}\Gemini Antigravity Excel"; Filename: "{app}\Gemini_Antigravity_Excel.xlsx"; IconFilename: "{app}\assets\favicon.ico"
Name: "{group}\Gemini Antigravity Word"; Filename: "{app}\Gemini_Antigravity_Word.docx"; IconFilename: "{app}\assets\favicon.ico"
Name: "{group}\Gemini Antigravity PowerPoint"; Filename: "{app}\Gemini_Antigravity_PowerPoint.pptx"; IconFilename: "{app}\assets\favicon.ico"
Name: "{group}\Gemini Antigravity Excel"; Filename: "{app}\Gemini_Antigravity_Excel.xlsx"; IconFilename: "{app}\assets\favicon.ico"
Name: "{group}\Trust Gemini Certificate (Manual Fix)"; Filename: "{app}\Trust_Gemini_Certificate.cmd"; IconFilename: "{app}\assets\favicon.ico"
Name: "{group}\Start Gemini Daemon"; Filename: "{app}\{#MyAppExeName}"; Parameters: "--background"; IconFilename: "{app}\assets\favicon.ico"

[Registry]
; WebView2 Silent Localhost SSL Bypass (Guarantees zero security warnings on any laptop)
Root: HKCU; Subkey: "Environment"; ValueType: string; ValueName: "WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS"; ValueData: "--ignore-certificate-errors --allow-insecure-localhost"; Flags: preservestringtype
Root: HKLM; Subkey: "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"; ValueType: string; ValueName: "WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS"; ValueData: "--ignore-certificate-errors --allow-insecure-localhost"; Flags: preservestringtype

; Office WEF Developer Sideload Registration (Word & PowerPoint)
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\WEF\Developer"; ValueType: string; ValueName: "{#MyAppGuid}"; ValueData: "{app}\manifest.xml"; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\WEF\Developer"; ValueType: string; ValueName: "{{{#MyAppGuid}}"; ValueData: "{app}\manifest.xml"; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\WEF\Developer\{{{#MyAppGuid}}"; ValueType: string; ValueName: "ManifestPath"; ValueData: "{app}\manifest.xml"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\WEF\Developer\{{{#MyAppGuid}}"; ValueType: dword; ValueName: "UseResourceTokens"; ValueData: 0

; Office WEF Trusted Catalog Registration
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{{{#MyAppGuid}}"; ValueType: string; ValueName: "Url"; ValueData: "{app}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{{{#MyAppGuid}}"; ValueType: dword; ValueName: "Flags"; ValueData: 1

; Windows Startup Auto-launch for Background Daemon
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "GeminiAntigravityOffice"; ValueData: """{app}\{#MyAppExeName}"" --background"; Flags: uninsdeletevalue

[Run]
; Install CA Certificate into Windows LocalMachine Root store (Silent, 0 prompts because setup is admin)
Filename: "certutil.exe"; Parameters: "-addstore -f Root ""{app}\certs\ca.crt"""; Flags: runhidden
Filename: "certutil.exe"; Parameters: "-addstore -f Root ""{app}\certs\localhost.crt"""; Flags: runhidden
Filename: "certutil.exe"; Parameters: "-user -addstore -f Root ""{app}\certs\ca.crt"""; Flags: runhidden
Filename: "certutil.exe"; Parameters: "-user -addstore -f TrustedPeople ""{app}\certs\localhost.crt"""; Flags: runhidden

; Run certificate trust and WEF registration silently via PowerShell
Filename: "powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{app}\wef_register.ps1"" -InstallDir ""{app}"""; Flags: runhidden

; Launch the Daemon
Filename: "{app}\{#MyAppExeName}"; Parameters: "--background"; Flags: nowait postinstall skipifsilent; Description: "Start Gemini Antigravity Office Daemon"

[UninstallRun]
; Silent cleanup of registry keys, processes, and certificates
Filename: "powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{app}\wef_unregister.ps1"" -InstallDir ""{app}"""; Flags: runhidden
