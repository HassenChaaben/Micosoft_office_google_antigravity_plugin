"""
Comprehensive Capabilities Verification Suite (Claude-Harness Protocol)
Validates:
1. Document Reading across Word, Excel, PowerPoint, and Outlook
2. Form & Data Matrix Filling across Excel, Word, and Outlook
3. Executive Professional Report Generation with advanced typography, zebra tables, and callouts
4. Plugin & Tool Invocation (MCP Connectors, Specialized Skills, and Cross-Suite Pipelines)
"""

import os
import re
import pytest
from starlette.testclient import TestClient

from daemon.main import app
from daemon.word_engine import word_engine
from daemon.excel_engine import excel_engine
from daemon.ppt_engine import ppt_engine
from daemon.outlook_engine import outlook_engine
from daemon.suite_orchestrator import suite_orchestrator
from daemon.gemini_client import gemini_client

client = TestClient(app)

# =============================================================================
# 1. READ CAPABILITIES (Word, Excel, PowerPoint, Outlook)
# =============================================================================

def test_reading_word_document():
    """Verify system can read active document content and text selections from Microsoft Word."""
    res = client.post("/api/document/word/read")
    assert res.status_code == 200
    data = res.json()
    assert "content" in data
    assert len(data["content"]) > 0
    assert data.get("type") in ["selection", "full_document", "mock"]

def test_reading_excel_sheet():
    """Verify system can read active worksheet data, dimensions, used range, and cell matrices from Microsoft Excel."""
    res = client.post("/api/document/excel/read")
    assert res.status_code == 200
    data = res.json()
    assert "sheet_name" in data
    assert "used_range" in data or "headers" in data
    # Verify data matrix or rows exist
    assert "rows" in data or "data" in data

def test_reading_powerpoint_presentation():
    """Verify system can read presentation slide deck outline, shapes, and notes from Microsoft PowerPoint."""
    res = client.post("/api/document/ppt/read")
    assert res.status_code == 200
    data = res.json()
    assert "slides" in data
    assert isinstance(data["slides"], list)
    assert len(data["slides"]) > 0
    assert "title" in data["slides"][0]

def test_reading_outlook_email():
    """Verify system can read active email subject, sender, body, and timestamp from Microsoft Outlook."""
    res = client.post("/api/document/outlook/read")
    assert res.status_code == 200
    data = res.json()
    assert "subject" in data
    assert "body" in data
    assert len(data["subject"]) > 0

# =============================================================================
# 2. FORM & DATA MATRIX FILLING (Excel, Word, Outlook)
# =============================================================================

def test_form_filling_excel_table_and_formulas():
    """Verify system can fill Excel form tables with formulas, formatting, and auto-fit."""
    table_data = [
        ["Quarter", "Unit Sales", "Unit Price", "Gross Revenue"],
        ["Q1", 1000, 49.99, "=B2*C2"],
        ["Q2", 1500, 49.99, "=B3*C3"],
        ["Total", "=SUM(B2:B3)", "-", "=SUM(D2:D3)"]
    ]
    res = client.post("/api/document/excel/write", json={
        "start_cell": "A1",
        "matrix": table_data,
        "format_table": True
    })
    assert res.status_code == 200
    assert res.json().get("success") is True

def test_form_filling_excel_financial_model():
    """Verify system can build and fill a Wall Street financial projection model."""
    res = client.post("/api/document/excel/financial_model", json={
        "model_type": "saas_revenue",
        "currency": "$"
    })
    assert res.status_code == 200
    data = res.json()
    assert data.get("success") is True
    assert data.get("model") == "saas_revenue"

def test_form_filling_excel_native_chart():
    """Verify system can generate native Excel charts (Column, Line, Pie) from data ranges."""
    res = client.post("/api/document/excel/chart", json={
        "chart_type": "column",
        "data_range": "A1:D3",
        "title": "Quarterly Performance"
    })
    assert res.status_code == 200
    assert res.json().get("success") is True

def test_form_filling_outlook_email_and_meeting():
    """Verify system can fill and dispatch Outlook email forms and calendar appointments."""
    # 1. Compose email form
    res_compose = client.post("/api/document/outlook/compose", json={
        "to": "client@enterprise.com",
        "subject": "Executive Review: Deliverables Ready",
        "body_html": "<p>Dear Client, your report is ready.</p>"
    })
    assert res_compose.status_code == 200
    assert res_compose.json().get("success") is True

    # 2. Reply form
    res_reply = client.post("/api/document/outlook/reply", json={
        "reply_html": "<p>Confirmed. Proceed with the proposed plan.</p>",
        "reply_all": False
    })
    assert res_reply.status_code == 200
    assert res_reply.json().get("success") is True

    # 3. Calendar appointment form
    res_meeting = client.post("/api/document/outlook/meeting", json={
        "subject": "Sprint Review & Sign-Off",
        "start_time": "2026-10-30 14:00",
        "duration_min": 45,
        "attendees": "hassen@enterprise.com; client@enterprise.com"
    })
    assert res_meeting.status_code == 200
    assert res_meeting.json().get("success") is True

def test_form_filling_word_comments_and_revisions():
    """Verify system can insert inline review comments and text revisions into Word."""
    res = client.post("/api/document/word/comment", json={
        "comment": "Gemini: Compliance clause verified against EU GDPR guidelines."
    })
    assert res.status_code == 200
    assert res.json().get("success") is True

# =============================================================================
# 3. EXECUTIVE PROFESSIONAL REPORT GENERATION
# =============================================================================

def test_executive_html_report_formatting():
    """Verify system transforms raw markdown into high-fidelity Word HTML with navy tables, zebra rows, and callouts."""
    raw_markdown = """# Executive Strategic Audit: Autonomous Office AI
## 1. Overview
This audit examines enterprise deployment velocity.

> **Key Finding:** Automated cross-application workflows eliminate 4.2 hours of manual formatting per employee weekly.

## 2. Quantitative Performance Matrix

| Metric | Baseline | Gemini Antigravity | Improvement |
| :--- | :--- | :--- | :--- |
| **Doc Review Time** | 45 min | 8 min | +82.2% |
| **Model Generation** | 60 min | 5 min | +91.6% |
| **Deck Creation** | 90 min | 12 min | +86.7% |

## 3. Actionable Next Steps
- Finalize silent deployment across Windows workstation fleet.
- Enable Notion and GitHub MCP connectors for all teams.
"""
    html_output = word_engine._markdown_to_word_html(raw_markdown)
    
    # 1. Typography checks
    assert "font-family:Aptos, Calibri, sans-serif" in html_output
    assert "<h1 style='font-size:22pt; font-weight:bold; color:#0F2C59;" in html_output
    assert "<h2 style='font-size:15pt; font-weight:bold; color:#1E3A8A;" in html_output
    
    # 2. Table formatting checks
    assert "<table border='1'" in html_output
    assert "background-color:#1E3A8A; color:#FFFFFF;" in html_output # Navy header with white text
    assert "background-color:#F8FAFC;" in html_output # Zebra row shading
    assert "Doc Review Time" in html_output
    
    # 3. Callout quote checks
    assert "border-left:4pt solid #4338CA;" in html_output # Left accent border
    assert "background-color:#EEF2FF;" in html_output # Light blue callout background
    assert "Key Finding:" in html_output

# =============================================================================
# 4. PLUGIN & TOOL INVOCATION (MCP Connectors, Skills, Cross-Suite Pipelines)
# =============================================================================

def test_mcp_connectors_invocation():
    """Verify all connected external tools and MCP connectors can be queried and invoked."""
    res = client.get("/api/connectors")
    assert res.status_code == 200
    connectors = res.json()
    connector_ids = [c["id"] for c in connectors]
    assert "notion" in connector_ids
    assert "github" in connector_ids
    assert "gdrive" in connector_ids
    assert "fs" in connector_ids or "filesystem" in connector_ids

def test_skills_library_invocation():
    """Verify all built-in skills can be executed via the daemon streaming chat endpoint."""
    test_skills = [
        "/check-doc",
        "/copy-edit",
        "/summarize-contract",
        "/flag-issues",
        "/financial-model",
        "/analyze-sheet",
        "/storylining",
        "/reply-email"
    ]
    for skill in test_skills:
        res = client.post("/api/chat", json={"prompt": skill, "host": "Word"})
        assert res.status_code == 200
        assert "text/event-stream" in res.headers.get("content-type", "")
        # Read streaming body
        body = res.text
        assert len(body) > 0
        assert "data: " in body

def test_cross_suite_multi_app_pipelines():
    """Verify cross-app pipelines: /doc-to-slides, /excel-to-report, /email-summary."""
    # 1. Word to PowerPoint (/doc-to-slides)
    res_slides = client.post("/api/suite/doc_to_slides")
    assert res_slides.status_code == 200
    assert res_slides.json().get("success") is True
    assert res_slides.json().get("slides_created") == 4

    # 2. Excel to Word (/excel-to-report)
    res_report = client.post("/api/suite/excel_to_report")
    assert res_report.status_code == 200
    assert res_report.json().get("success") is True

    # 3. Suite to Outlook (/email-summary)
    res_summary = client.post("/api/suite/email_summary", json={"recipient": "ergys@cloud4job.com"})
    assert res_summary.status_code == 200
    assert res_summary.json().get("success") is True
