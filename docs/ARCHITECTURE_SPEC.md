# Reverse Engineering Architecture (REA) Specification
## Gemini Antigravity Office Plugin (Word & PowerPoint)

### 1. Architectural Overview & Deconstruction of Claude Office
The official Claude Office Add-in implements a hybrid client-daemon architecture that bypasses typical web sandbox and enterprise Office limitations:
1. **The Office WebView2 Sandbox Container**:
   - Word and PowerPoint host modern Web Add-ins inside a dedicated Microsoft Edge WebView2 process (`winsta.exe` or `mso.exe`).
   - The UI runs in a secure appcontainer with standard browser sandbox constraints (isolated storage, restricted localhost access, and blocked direct COM/Win32 APIs).
2. **The Localhost Companion Daemon (`https://localhost:18890`)**:
   - Claude Office spawns a background Windows process running locally.
   - This daemon serves the Taskpane web application over a trusted local HTTPS certificate.
   - It functions as an unrestricted agent runtime:
     * Accesses the local filesystem to load MCP configuration, plugins, and skills (e.g., `~/.gemini/config/`).
     * Directs bi-directional COM automation with Word and PowerPoint via `win32com.client` / Win32 IDispatch interfaces for high-privilege operations (Track Changes toggling, background slide manipulation, headless batch operations).
     * Connects directly to Google Gemini via browser session OAuth2 / Antigravity authentication without requiring raw API keys.
3. **Silent Windows Registration (Zero-Sideload Deployment)**:
   - Rather than forcing users through the Office 365 Store or manual XML upload dialogs, the installer writes directly to the Office Web Extension Framework (WEF) developer registry branch:
     `HKCU\Software\Microsoft\Office\16.0\WEF\Developer\{7F89B2E3-8E5A-4B21-996A-12E4A9F200A1}`
     * `ManifestPath` = Path to local `manifest.xml`
     * `UseResourceTokens` = 0
   - Configures a Trusted Catalog:
     `HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{7F89B2E3-8E5A-4B21-996A-12E4A9F200A1}`
     * `Url` = Manifest folder path
     * `Flags` = 1 (Trusted)
   - On the next launch of Word or PowerPoint, Office reads the WEF developer tree and automatically mounts the ribbon button in the Home (`TabHome`) tab.

---

### 2. Detailed Component Architecture

```
+-----------------------------------------------------------------------------------+
|                            Microsoft Office Host Process                          |
|                       (WINWORD.EXE / POWERPNT.EXE - Office 16)                    |
|                                                                                   |
|  +-----------------------------------------------------------------------------+  |
|  | Ribbon (Home Tab): [Gemini Antigravity] Button (Id: GeminiAntigravityBtn)   |  |
|  +-----------------------------------------------------------------------------+  |
|                                         |                                         |
|                                    Launches                                       |
|                                         v                                         |
|  +-----------------------------------------------------------------------------+  |
|  | WebView2 Taskpane Container: https://localhost:18890/index.html              |  |
|  |                                                                             |  |
|  |  [Header: Title, History Clock, New Chat +, Menu ⋮]                         |  |
|  |  [Skills Onboarding: /check-doc, /copy-edit, /summarize-contract, etc.]     |  |
|  |  [Chat Stream & Diff Preview]                                               |  |
|  |  [Dock: (+) Flyout | Context Globe | Model Dropdown | Input | Mic]          |  |
|  |                                                                             |  |
|  |  +-----------------------------------------------------------------------+  |  |
|  |  | Office.js Runtime (Word.run / PowerPoint.run)                        |  |  |
|  |  | In-sandbox document reading, inline comments, native styling         |  |  |
|  |  +-----------------------------------------------------------------------+  |  |
|  +-----------------------------------------------------------------------------+  |
+---------------------------------------|-------------------------------------------+
                                        | HTTPS / WSS / REST
                                        v
+-----------------------------------------------------------------------------------+
|                   Gemini Antigravity Localhost Daemon Process                     |
|                       (gemini_office_daemon.exe - Port 18890)                     |
|                                                                                   |
|  +---------------------------+  +----------------------------------------------+  |
|  | Fast API / Uvicorn HTTPS  |  | Auth Bridge: Google Gemini Subscription      |  |
|  | TLS 1.3 Localhost Cert    |  | OAuth2 Loopback / Antigravity Session        |  |
|  +---------------------------+  +----------------------------------------------+  |
|  | MCP Connectors Engine     |  | Document & Presentation Engine               |  |
|  | Notion / Drive / Git / FS |  | win32com.client Word.Application             |  |
|  | Syncs ~/.gemini/config    |  | win32com.client PowerPoint.Application       |  |
|  +---------------------------+  +----------------------------------------------+  |
|  | Skills Execution Engine   |  | Gemini Streaming Client                      |  |
|  | /check-doc, /storylining  |  | 3.8 Flash High / Pro / Deep Think            |  |
|  +---------------------------+  +----------------------------------------------+  |
+-----------------------------------------------------------------------------------+
```

---

### 3. Registry & Certificate Specification
1. **SSL Certificate**:
   - Common Name: `localhost`
   - Subject Alternative Name: `DNS:localhost, IP:127.0.0.1`
   - Key Usage: Digital Signature, Key Encipherment, Server Authentication (`1.3.6.1.5.5.7.3.1`)
   - Installed in: `Cert:\CurrentUser\Root` (Windows Trusted Root Certification Authorities).
2. **Registry Keys**:
   - `HKCU\Software\Microsoft\Office\16.0\WEF\Developer\{7F89B2E3-8E5A-4B21-996A-12E4A9F200A1}`:
     - `ManifestPath`: string (`%LOCALAPPDATA%\GeminiAntigravityOffice\manifest.xml`)
     - `UseResourceTokens`: dword (`0`)
   - `HKCU\Software\Microsoft\Windows\CurrentVersion\Run`:
     - `GeminiAntigravityOffice`: string (`"%LOCALAPPDATA%\GeminiAntigravityOffice\gemini_office_daemon.exe" --background`)
3. **Office Manifest**:
   - Strict Office 1.1 XML Schema compliance.
   - Dual-host definition: `<Host Name="Document" />` and `<Host Name="Presentation" />`.
   - Ribbon group located on Home Tab (`idQ="mso:TabHome"`).
