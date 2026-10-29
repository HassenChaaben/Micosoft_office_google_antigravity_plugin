# Unified Microsoft Office Suite Controller Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Expand Gemini Antigravity Office into a unified, all-in-one controller across all Microsoft Office applications (Word, Excel, PowerPoint, and Outlook) with bidirectional Office.js + Win32 COM automation, host-aware adaptive UI, cross-application workflows, and turnkey Windows installer.

**Architecture:** 
- Unified Office Manifest supporting Word (`Document`), Excel (`Workbook`), and PowerPoint (`Presentation`), plus Mailbox COM hooks for Outlook.
- Four specialized autonomous engines in `daemon/`: `WordEngine`, `PowerPointEngine`, `ExcelEngine`, and `OutlookEngine`, coordinated by a master `SuiteOrchestrator`.
- Host-adaptive WebView2 taskpane in `web/` that automatically detects the active application (`Word`, `Excel`, `PowerPoint`, `Outlook`) and renders tailored skill pills, tools, and actions.
- Cross-app pipelines: Word document summarization to PowerPoint slides, Excel financial tables to Word reports, and executive summaries to Outlook email drafts.

**Tech Stack:** Office.js, Win32 COM (`win32com.client`), FastAPI / Python 3.12, Edge WebView2, OpenXML (`python-docx`, `openpyxl`, `python-pptx`), Inno Setup 6, PyInstaller.

**Spec:** `docs/ARCHITECTURE_SPEC.md`

## Global Constraints
- Turnkey zero-config installation on any Windows 10/11 system.
- Zero manual SSL certificate warnings (trusted root + WebView2 ignore policies).
- Identical visual fidelity and typography matching modern Microsoft Office Ribbon design.
- Zero AI attribution in commits or production logs.
- English code identifiers, Italian documentation where Cloud4Job applies.

## Review Focus
1. **Excel Range & Formula Integrity:** Ensure Excel formulas (`SUM`, `XLOOKUP`, `AVERAGE`) use localized or standard formula notation and preserve formatting.
2. **Outlook Security & Permissions:** Ensure email reading and drafting operate via secure local COM/MAPI without triggering security popups.
3. **Office Manifest Multi-Host Schema Compliance:** Validate manifest with `npx office-addin-manifest validate` for Document, Workbook, and Presentation.
4. **Cross-App COM Concurrency:** Ensure `pythoncom.CoInitialize()` is properly handled per thread and releases locks cleanly.
5. **Turnkey Sideloading Across All Hosts:** Ensure registry keys cover Word (`16.0\WEF\Developer`), Excel, PowerPoint, and Trusted Catalogs.

---

### Task 1: Microsoft Excel Autonomous Engine (`daemon/excel_engine.py`)
**Files:**
- Create: `daemon/excel_engine.py`
- Modify: `daemon/main.py`
- Test: `tests/unit/test_excel_engine.py`

- [x] Write unit test for `ExcelEngine` covering `read_sheet()`, `write_cell_range()`, `apply_financial_model()`, `insert_formula()`, and `create_chart()`.
- [x] Implement `daemon/excel_engine.py` with Win32 COM (`Excel.Application`).
- [x] Implement `read_sheet()`: returns active sheet name, used range, column headers, cell matrix, and selected range.
- [x] Implement `write_cell_range(start_cell, data_matrix)`: writes 2D matrix of values and formulas into active sheet.
- [x] Implement `apply_financial_model(model_type, data)`: creates formatted 3-statement or revenue projections with bold headers, currency styling, zebra shading, and sum formulas.
- [x] Implement `create_chart(chart_type, data_range, title)`: inserts native Excel chart (Clustered Column, Line, or Pie).
- [x] Add REST endpoints in `daemon/main.py`:
  - `POST /api/document/excel/read`
  - `POST /api/document/excel/write`
  - `POST /api/document/excel/financial_model`
  - `POST /api/document/excel/chart`
- [x] Run `python -m pytest tests/unit/test_excel_engine.py` and verify all tests pass.

---

### Task 2: Microsoft Outlook Autonomous Engine (`daemon/outlook_engine.py`)
**Files:**
- Create: `daemon/outlook_engine.py`
- Modify: `daemon/main.py`
- Test: `tests/unit/test_outlook_engine.py`

- [x] Write unit test for `OutlookEngine` covering `read_current_email()`, `draft_reply()`, `create_email()`, and `schedule_meeting()`.
- [x] Implement `daemon/outlook_engine.py` using Win32 COM (`Outlook.Application`).
- [x] Implement `read_current_email()`: reads active inspector/explorer item (Subject, Sender, Body, ReceivedTime, Recipients).
- [x] Implement `draft_reply(reply_body, reply_all)`: creates in-memory draft reply with formatted HTML.
- [x] Implement `create_email(to, subject, body_html)`: creates new draft email.
- [x] Implement `schedule_meeting(subject, start_time, duration_min, attendees)`: creates Outlook appointment item.
- [x] Add REST endpoints in `daemon/main.py`:
  - `POST /api/document/outlook/read`
  - `POST /api/document/outlook/reply`
  - `POST /api/document/outlook/compose`
  - `POST /api/document/outlook/meeting`
- [x] Run `python -m pytest tests/unit/test_outlook_engine.py` and verify all tests pass.

---

### Task 3: Unified Suite Orchestrator & Cross-App Pipelines (`daemon/suite_orchestrator.py`)
**Files:**
- Create: `daemon/suite_orchestrator.py`
- Modify: `daemon/main.py`
- Modify: `daemon/gemini_client.py`
- Test: `tests/unit/test_suite_orchestrator.py`

- [x] Write unit test for cross-application pipelines.
- [x] Implement `SuiteOrchestrator`:
  - `doc_to_slides()`: reads active Word document, extracts core thesis and sections, builds multi-slide PowerPoint presentation.
  - `excel_to_word_report()`: reads active Excel financial/data table, formats it into an executive Word report with commentary.
  - `email_summary()`: reads active Word doc or Excel sheet, drafts executive summary in Outlook.
- [x] Update `daemon/gemini_client.py` with specialized skills:
  - `/analyze-sheet`: statistical summary, anomalies, and trend analysis for Excel.
  - `/financial-model`: builds balance sheet, DCF, or revenue projection model.
  - `/reply-email`: professional executive email reply with action points.
  - `/doc-to-slides`: autonomous PowerPoint deck generation from Word.
- [x] Add REST endpoints in `daemon/main.py`:
  - `POST /api/suite/doc_to_slides`
  - `POST /api/suite/excel_to_report`
  - `POST /api/suite/email_summary`
- [x] Run tests and verify green.

---

### Task 4: Office Multi-Host Manifest Expansion (`manifest/manifest.xml`)
**Files:**
- Modify: `manifest/manifest.xml`
- Test: `npx office-addin-manifest validate manifest/manifest.xml`

- [x] Add `<Host Name="Workbook" />` to `<Hosts>`.
- [x] Add `<Host xsi:type="Workbook">` to `<VersionOverrides><Hosts>`:
  - Add primary ribbon group `GeminiExcelGroup` on `TabHome`.
  - Add `GeminiExcelBtn` button triggering `GeminiAntigravityTaskpane`.
- [x] Validate manifest XML using `npx office-addin-manifest validate manifest/manifest.xml`. Confirm 100% valid schema without warnings.

---

### Task 5: Host-Adaptive Taskpane Web UI (`web/`)
**Files:**
- Modify: `web/office-bridge.js`
- Modify: `web/taskpane.js`
- Modify: `web/taskpane.css`
- Modify: `web/index.html`

- [x] Extend `web/office-bridge.js` to detect and control `Excel` via Office.js:
  - `readExcelContent()`: loads active range values, formulas, and sheet names.
  - `writeExcelRange(rangeAddress, valuesMatrix)`: inserts formatted values and formulas into Excel.
  - `insertExcelChart(type, rangeAddress)`: creates native Excel chart.
- [x] Extend `web/taskpane.js` with Host-Aware Adaptation:
  - Detects if host is `Word`, `PowerPoint`, `Excel`, `Outlook`, or `Standalone`.
  - Updates host badge with application color (Word: `#2B579A`, Excel: `#217346`, PowerPoint: `#D24726`, Outlook: `#0078D4`).
  - Swaps the onboarding skills pills dynamically based on active host:
    - **Word**: `/check-doc`, `/copy-edit`, `/summarize-contract`, `/flag-issues`, `/doc-to-slides`
    - **Excel**: `/analyze-sheet`, `/financial-model`, `/formula-helper`, `/generate-chart`, `/excel-to-report`
    - **PowerPoint**: `/storylining`, `/slide-designer`, `/executive-summary`
    - **Outlook**: `/reply-email`, `/summarize-thread`, `/schedule-meeting`, `/draft-memo`
  - Updates slash (`/`) autocomplete catalog to include all Excel, PowerPoint, Word, Outlook skills and connectors.
  - Updates bottom floating dock action buttons:
    - In Excel: `📊 Insert Financial Model into Excel`, `📈 Generate Chart`, `📋 Copy Data`.
    - In Word: `✨ Insert Formatted Report into Word`, `💬 Review Comment`.
    - In PowerPoint: `📊 Create Formatted Slide`.

---

### Task 6: Turnkey Registration, Test Launch Documents & Packaging
**Files:**
- Create: `assets/Gemini_Antigravity_Excel.xlsx` (pre-configured OpenXML webextension)
- Modify: `manifest/wef_register.ps1`
- Modify: `installer/setup.iss`
- Modify: `scripts/verify.py`

- [x] Create pre-configured test file `assets/Gemini_Antigravity_Excel.xlsx` with embedded webextension sideloading `{7F89B2E3-8E5A-4B21-996A-12E4A9F200A1}`.
- [x] Update `manifest/wef_register.ps1` to write WEF Developer keys for Excel and Outlook in addition to Word and PowerPoint.
- [x] Update `installer/setup.iss`:
  - Add Excel test document `Gemini_Antigravity_Excel.xlsx` to desktop shortcuts.
  - Write WEF Developer and Trusted Catalog registry keys for Excel.
- [x] Update `scripts/verify.py` to test all 4 application engines, manifest validation, API routes, and installer output.
- [x] Recompile PyInstaller executable `dist/gemini_office_daemon/`.
- [x] Recompile standalone installer `dist/GeminiAntigravityOffice_Setup.exe`.
- [x] Execute `python scripts/verify.py` to confirm 100% green verification across the entire suite.
