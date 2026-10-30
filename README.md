<p align="center">
  <img src="assets/logo.svg" width="96" height="96" alt="Gemini Antigravity Logo">
</p>

<h1 align="center">Gemini Antigravity Microsoft Office Plugin</h1>

<p align="center">
  <b>Autonomous, Publication-Grade AI Copilot for Microsoft Word, PowerPoint, and Excel</b><br>
  <i>Powered by Google Gemini and the Antigravity Agentic Cognitive Runtime</i>
</p>

<p align="center">
  <a href="LICENSE"><img src="https://img.shields.io/badge/License-MIT-blue.svg" alt="License: MIT"></a>
  <a href="https://github.com/HassenChaaben/Micosoft_office_google_antigravity_plugin"><img src="https://img.shields.io/badge/Platform-Windows%2011%20%7C%2010-0078D4.svg?logo=windows" alt="Platform: Windows 11 / 10"></a>
  <a href="https://python.org"><img src="https://img.shields.io/badge/Python-3.12-3776AB.svg?logo=python" alt="Python: 3.12"></a>
  <a href="https://fastapi.tiangolo.com"><img src="https://img.shields.io/badge/FastAPI-0.115+-009688.svg?logo=fastapi" alt="FastAPI"></a>
  <a href="https://github.com/HassenChaaben/Micosoft_office_google_antigravity_plugin/releases"><img src="https://img.shields.io/badge/Release-v1.0.1--Standalone-success.svg" alt="GitHub Releases"></a>
</p>

> 🔗 **Repository**: [https://github.com/HassenChaaben/Micosoft_office_google_antigravity_plugin.git](https://github.com/HassenChaaben/Micosoft_office_google_antigravity_plugin.git)  
> 📦 **Download Standalone Setup**: [Releases / GeminiAntigravityOffice_Setup.exe](https://github.com/HassenChaaben/Micosoft_office_google_antigravity_plugin/releases)

---

## 📸 Visual Walkthrough & Real-World Use Cases

The following real-world screenshots demonstrate Gemini Antigravity operating live inside Microsoft Word on an executive technical dossier (*"Progetto di Automazione e Monitoraggio Esecutivo"*):

### 1. Executive Document Drafting & Publication-Grade Table Insertion
The user requests a comprehensive project delivery schedule. Gemini Antigravity perceives the document outline and inserts **Tabella 1** (*"Stato di avanzamento dei deliverable e della documentazione"*), populated with 6 records across 3 columns (`Elemento / Documento`, `Descrizione / Specifiche`, `Stato / Valore`), shaded header styling, automated column widths, and contextual narrative text without paragraph duplication:

![Executive Document Drafting & Table Insertion](docs/images/01_document_drafting_table.png)

### 2. Multi-Table Business & Financial Modeling
The user asks to expand the dossier with budget allocation and financial tracking. The agent builds **Tabella 2** (*"Ripartizione Budget ed Economica delle Fasi Operative"*), comparing planned estimates against actual costs with formatted currency numbers and harmonized grid formatting matching the document's design tokens:

![Multi-Table Financial Modeling](docs/images/02_business_financial_modeling.png)

### 3. Autonomous Process Flowcharts & Executive Visual Cards
Transforming raw tabular metrics into visual roadmap components. The agent generates **Figura 1** (*"Diagramma di flusso e avanzamento per fasi"*), creating a multi-stage visual process flow with milestone cards, KPI badges, and status colors directly rendered onto the Word canvas:

![Autonomous Process Flowcharts & Executive Visual Cards](docs/images/03_flowchart_diagram_generation.png)

### 4. Multi-Turn Conversational Memory & Persistent Session Drawer
Iterative refinement across complex documents. The user opens the **"Your conversations"** session drawer in the taskpane to view conversation history (`"cretae a table of 3 colu..."`), switch sessions, remove phantom blank pages, re-index headings, and customize diagram palettes:

![Multi-Turn Conversational Memory & Persistent Session Drawer](docs/images/04_iterative_refinement_styling.png)

---

## ⚡ What We Have Built

Gemini Antigravity is an autonomous Office agent that operates directly on document objects rather than just returning markdown text. It bridges the modern **Google Gemini** model family with native desktop Office applications via a local background daemon.

### 1. Autonomous Microsoft Word Agent (`daemon/office/`)
- **Perceive-First Loop**: Automatically inspects document structure, paragraph count, tables, headings, and active cursor selection *before* sending prompts to the model.
- **Native Table Insertion (`word_insert_table`)**: Creates publication-grade Word tables with specified headers, row counts, and auto-formatting without creating phantom duplicates.
- **In-Place Paragraph Editing (`word_set_paragraph_text`, `word_replace_text`)**: Edits existing paragraphs directly in place, preventing unwanted duplicate blocks when refining documents.
- **Form Field Filling (`word_fill_form_fields`)**: Identifies label-value pairs and template cells, filling adjacent cells automatically.
- **Real-Time Stop & Cancel**: Immediate task cancellation via the UI `⏹ Stop` button or the `Esc` keyboard shortcut.

### 2. AI Text-to-Image Generation (Gemini Imagen 3)
- **Zero-Subscription Default**: The plugin defaults to your active Antigravity subscription runtime.
- **Imagen 3 Integration (`word_generate_image`)**: Users can input their personal Gemini API key in Settings to unlock direct AI image generation. The agent creates the illustration and inserts it directly into Word with automatic captions and alignment.

### 3. Multimodal Prompt Picture Attachments
- **Clipboard Screenshots (`Ctrl+V`)**: Take a screenshot with `Win + Shift + S` and paste it directly into the prompt box.
- **Photo Picker (`#btnAttachPhoto`)**: Attach `.png`, `.jpg`, `.webp`, `.gif`, or `.bmp` files.
- **Drag & Drop**: Drop images straight onto the input dock card.
- **Local Persistence & Injection**: Attachments are saved to `%LOCALAPPDATA%\GeminiAntigravityOffice\attachments\` and fed directly into the model's prompt for inspection, analysis, or insertion via `word_insert_image`.

### 4. Turnkey Client Architecture
- **Local HTTPS Daemon (`gemini_office_daemon.exe`)**: Self-contained background service compiled with PyInstaller. Runs on port `18890` with auto-generated and pre-trusted Windows root SSL certificates.
- **Office WEF Sideloading (`wef_register.ps1`)**: Automatically registers the manifest across Word, PowerPoint, and Excel across all Windows UI locales (`en-US`, `fr-FR`, `it-IT`, `de-DE`, `es-ES`).
- **Windows Startup**: Automatically launches silently on Windows user logon via `HKCU\...\Run`.

![Local Daemon Terminal](docs/images/daemon_terminal.png)

---

## 🔍 Exact Capabilities & What is Missing (Zero-Fluff Matrix)

To provide total clarity on what this tool currently does and what remains on the roadmap:

| Capability | Status | Implementation Details & Current Reality |
|---|:---:|---|
| **Microsoft Word Win32 COM Automation** | ✅ Production | Paragraph rewrites, native table insertion, direct formatting, styles, outline inspection. |
| **Multimodal Image Prompt Attachments** | ✅ Production | Clipboard paste (`Ctrl+V`), photo picker (`#btnAttachPhoto`), drag & drop, local caching in `%LOCALAPPDATA%`. |
| **Gemini Imagen 3 AI Image Generation** | ✅ Production | Activated when user enters Gemini API key in Settings; inserts illustrations with captions. |
| **Live Streaming & Task Interruption** | ✅ Production | Real-time SSE token delivery with `⏹ Stop` button and `Esc` keyboard abort. |
| **Word Ribbon Home Tab Integration** | ✅ Production | Permanent ribbon button across all installed Office language packs (`en-US`, `fr-FR`, `it-IT`, `de-DE`, `es-ES`). |
| **Offline Test Mock Mode** | ✅ Production | Mock responses for headless CI test suites (`GEMINI_OFFICE_BACKEND=offline`). |
| **One-Click Turnkey Windows Setup** | ✅ Production | Single 51 MB installer (`GeminiAntigravityOffice_Setup.exe`) with auto-cert and auto-registry injection. |
| **`/skills` Dynamic Skill Execution** | ❌ Not Working Yet | The 6 pills (`/check-doc`, `/copy-edit`, etc.) are **UI preview pills** only; no isolated server-side skill sandboxes exist yet. |
| **`/plugins` & Custom Connectors Menu** | ❌ Not Working Yet | The flyout menu exists in the UI, but external third-party API connectors (GitHub, Jira, Notion) are **unconnected stubs**. |
| **`search_web` Internet Research Tool** | ❌ Not Working Yet | The agent is **strictly offline/document-bound**; there is no live web search API integration connected. |
| **Excel 3-Statement Financial Modeling** | ⚠️ Partial | Basic table output and cell formatting work; dynamic multi-sheet formula linkage is in progress. |
| **PowerPoint Executive Storylining** | ⚠️ Partial | Slide creation and text boxes work; custom vector shape coordinate layout is in progress. |
| **macOS / Office Online (Web) Support** | ❌ Not Supported | Architecture relies strictly on native Windows Win32 COM and local Windows daemon IPC. |
| **Cross-Window Multi-Document Sync** | ❌ Not Supported | Agent targets active foreground document; multi-doc switching across separate windows is in progress. |

---

### Detailed Breakdown of Incomplete & Stubbed Features

To ensure complete engineering transparency for contributors and users, here is exactly what is NOT yet functional:

#### 1. The `/skills` Pills Are UI Blueprints Only (No Dynamic Server Execution)
In the onboarding view of the taskpane, there are 6 clickable skill pill buttons:
- `/check-doc` (Review structure & formatting)
- `/copy-edit` (Enhance style & clarity)
- `/summarize-contract` (Legal clause overview)
- `/flag-issues` (Risk & discrepancy identification)
- `/industry-overview` (Market analysis draft)
- `/storylining` (Executive narrative structuring)

> ⚠️ **Current Reality**: Clicking these buttons currently only pastes suggested starter text into `#promptInput`. There is **no dynamic server-side Python execution engine** that loads isolated skills from `daemon/skills/`. All requests are routed through the single monolithic `word_agent` / `gemini_client` prompt pipeline. PRs to introduce a modular skill runner are warmly welcomed!

#### 2. The `/plugins` & Connectors Menu Is a Non-Functional Prototype
In the input dock card, clicking the `+` flyout button displays a "Connectors" option (`#menuConnectors`) intended for third-party integrations (GitHub, Jira, Google Workspace, Notion).

> ⚠️ **Current Reality**: This menu is currently a **visual mockup**. No backend API connectors exist to fetch remote files, synchronize issues, or query external tools. The agent currently only operates on the locally active Office document.

#### 3. The `search_web` Tool Is Not Connected
When the agent drafts scientific or business sections (e.g. the Quantum Computing overview on Page 5 of the sample document), users may expect live internet browsing or citations.

> ⚠️ **Current Reality**: The agent currently has **no live search provider connected** (e.g., Google Custom Search API, Serper, or Tavily). The agent generates document text purely based on the LLM's parametric knowledge and the text extracted from the user's active document. Live web browsing and retrieval-augmented verification are planned for v1.1.0.

## 🚀 Quick Setup & Installation

### Option A: Install via Standalone Windows Setup *(Recommended)*

No Python, Node.js, or command line required:

1. **Download the Installer**:
   Download [`GeminiAntigravityOffice_Setup.exe`](https://github.com/HassenChaaben/Micosoft_office_google_antigravity_plugin/releases) from the **[Releases](https://github.com/HassenChaaben/Micosoft_office_google_antigravity_plugin/releases)** page.
2. **Run the Installer**:
   Double-click the setup wizard. It automatically:
   - Deploys the standalone daemon and web taskpane to `%LOCALAPPDATA%\GeminiAntigravityOffice`.
   - Generates and trusts the localhost TLS certificate.
   - Injects the add-in into Word, PowerPoint, and Excel ribbons.
   - Configures the background service to start on Windows logon.
3. **Launch Microsoft Word**:
   - Open Word (or double-click the **Gemini Antigravity Word** desktop shortcut).
   - In the top ribbon on the **Accueil** (or **Home**) tab, click the black **Gemini** icon button.
   - The taskpane will open and connect automatically.

---

### Option B: Build & Run from Source (Developers)

#### Prerequisites
- Windows 11 or Windows 10
- Python 3.12+
- Microsoft Office (Word 2016, 2019, 2021, or Microsoft 365)
- Inno Setup 6 (optional, to compile the `.exe` installer)

#### 1. Clone the Repository
```bash
git clone https://github.com/HassenChaaben/Micosoft_office_google_antigravity_plugin.git
cd Micosoft_office_google_antigravity_plugin
```

#### 2. Install Dependencies
```powershell
pip install -r requirements.txt
```

#### 3. Run Automated Tests
```powershell
# Run JavaScript Taskpane & Agent tests
node --test tests/js/taskpane-utils.test.mjs tests/js/word-agent.test.mjs

# Run Python Backend & Attachment tests
pytest tests/test_imagen_client.py tests/test_image_attachments.py -v
```

#### 4. Build Daemon & Installer
```powershell
powershell -ExecutionPolicy Bypass -File scripts/build.ps1
```
The output installer will be generated at:
`dist\GeminiAntigravityOffice_Setup.exe`

---

## 🤝 Open for Contribution

This project is **100% open-source** and actively welcomes community contributions!

We welcome pull requests for:
- 💡 **New Office Skills**: Adding specialized `/commands` for legal contracts, medical papers, or academic theses.
- 📊 **Excel Autonomous Modeling**: Enhancing formula graphs and financial statements.
- 🎨 **PowerPoint Slide Automation**: Richer slide layouts and visual themes.
- 🍎 **macOS & Web Office.js Compatibility**: Expanding pure Office.js fallback tools for cross-platform support.

### How to Contribute
1. Fork the repository: [https://github.com/HassenChaaben/Micosoft_office_google_antigravity_plugin](https://github.com/HassenChaaben/Micosoft_office_google_antigravity_plugin).
2. Create a feature branch: `git checkout -b feat/my-new-feature`.
3. Commit your changes: `git commit -m "feat: add support for custom chart themes"`.
4. Push to your fork and submit a **Pull Request**.

---

## 📄 License

Distributed under the **MIT License**. See `LICENSE` for more information.

---

<p align="center">
  <b>Gemini Antigravity Microsoft Office Plugin</b> • Created & Maintained by <a href="https://github.com/HassenChaaben">Hassen Ben Chaaben</a>
</p>
