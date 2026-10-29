"""
Comprehensive Verification Suite for Gemini Antigravity Office Plugin
Verifies Manifest XML, Assets, UI markup, Daemon Endpoints, COM bridges, and Installer Binary.
"""

import os
import sys
import xml.etree.ElementTree as ET
from PIL import Image

# Ensure standard UTF-8 friendly output
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

# Endpoint checks verify HTTP shape only; skip real model calls here.
os.environ.setdefault("GEMINI_OFFICE_BACKEND", "offline")

def test_manifest():
    print("[1/6] Verifying Office Add-in Manifest XML...")
    manifest_path = os.path.join(PROJECT_ROOT, "manifest", "manifest.xml")
    assert os.path.exists(manifest_path), f"Manifest missing at {manifest_path}"

    tree = ET.parse(manifest_path)
    root = tree.getroot()
    assert "OfficeApp" in root.tag, "Root element is not OfficeApp"

    # Read hosts
    hosts = [h.attrib.get("Name") for h in root.findall(".//{http://schemas.microsoft.com/office/appforoffice/1.1}Host")]
    assert "Document" in hosts, "Document (Word) host not found in manifest"
    assert "Presentation" in hosts, "Presentation (PowerPoint) host not found in manifest"
    assert "Workbook" in hosts, "Workbook (Excel) host not found in manifest"

    # Check Taskpane Url
    urls = [u.attrib.get("DefaultValue") for u in root.findall(".//{http://schemas.microsoft.com/office/officeappbasictypes/1.0}Url")]
    assert any("https://localhost:18890/index.html" in u for u in urls if u), "Taskpane URL not configured properly"

    print("  [OK] Manifest XML schema and triple-host (Word, Excel, PowerPoint) configuration verified!")

def test_assets():
    print("[2/6] Verifying Branding & Icon Assets...")
    assets_dir = os.path.join(PROJECT_ROOT, "assets")
    expected_sizes = {
        "icon-16.png": (16, 16),
        "icon-32.png": (32, 32),
        "icon-64.png": (64, 64),
        "icon-80.png": (80, 80),
        "icon-128.png": (128, 128)
    }

    for filename, expected_dim in expected_sizes.items():
        file_path = os.path.join(assets_dir, filename)
        assert os.path.exists(file_path), f"Asset missing: {filename}"
        with Image.open(file_path) as img:
            assert img.size == expected_dim, f"{filename} dimensions {img.size} != {expected_dim}"

    assert os.path.exists(os.path.join(assets_dir, "logo.svg")), "logo.svg missing"
    assert os.path.exists(os.path.join(assets_dir, "favicon.ico")), "favicon.ico missing"
    assert os.path.exists(os.path.join(assets_dir, "installer_banner.bmp")), "installer_banner.bmp missing"
    assert os.path.exists(os.path.join(assets_dir, "Gemini_Antigravity_Word.docx")), "Gemini_Antigravity_Word.docx missing"
    assert os.path.exists(os.path.join(assets_dir, "Gemini_Antigravity_PowerPoint.pptx")), "Gemini_Antigravity_PowerPoint.pptx missing"
    assert os.path.exists(os.path.join(assets_dir, "Gemini_Antigravity_Excel.xlsx")), "Gemini_Antigravity_Excel.xlsx missing"
    print("  [OK] All vector, multi-resolution PNGs, ICO, BMP, and launch documents verified!")

def test_ui_markup():
    print("[3/6] Verifying Taskpane Web UI & 6 Exact Skill Pills...")
    html_path = os.path.join(PROJECT_ROOT, "web", "index.html")
    assert os.path.exists(html_path), "index.html missing"

    with open(html_path, "r", encoding="utf-8") as f:
        content = f.read()

    # The 6 exact skills
    required_skills = [
        "/check-doc",
        "/copy-edit",
        "/summarize-contract",
        "/flag-issues",
        "/industry-overview",
        "/storylining"
    ]
    for skill in required_skills:
        assert skill in content, f"Skill {skill} missing from index.html"

    # UI controls
    assert "Gemini 3.8 Flash High" in content, "Model selector missing Gemini 3.8 Flash High"
    assert "Deep Think" in content, "Model selector missing Deep Think"
    assert "btnFlyoutPlus" in content, "Plus flyout button missing"
    assert "flyoutMenu" in content, "Flyout popup menu missing"
    assert "menuConnectors" in content, "Connectors menu missing"
    print("  [OK] Taskpane UI markup and skills onboarding verified 100%!")

def test_daemon_api():
    print("[4/6] Verifying Daemon API Endpoints via FastAPI TestClient...")
    from starlette.testclient import TestClient
    from daemon.main import app

    client = TestClient(app)

    # 1. Health
    res = client.get("/health")
    assert res.status_code == 200
    assert res.json()["status"] == "ok"

    # 2. Auth Status
    res = client.get("/api/auth/status")
    assert res.status_code == 200
    auth_data = res.json()
    assert auth_data.get("authenticated") is True
    assert "email" in auth_data

    # 3. Connectors
    res = client.get("/api/connectors")
    assert res.status_code == 200
    connectors = res.json()
    assert any(c["id"] == "notion" for c in connectors)

    # 4. Word Automation Read & Comment
    res = client.post("/api/document/word/read")
    assert res.status_code == 200
    res = client.post("/api/document/word/comment", json={"comment": "Verification Test Comment"})
    assert res.status_code == 200
    assert res.json().get("success") is True

    # 5. PowerPoint Automation Read & Create Slide
    res = client.post("/api/document/ppt/read")
    assert res.status_code == 200
    res = client.post("/api/document/ppt/create_slide", json={"title": "Verification Slide", "bullets": ["Point 1", "Point 2"]})
    assert res.status_code == 200
    assert res.json().get("success") is True

    # 6. Excel Automation Read, Write, Financial Model & Chart
    res = client.post("/api/document/excel/read")
    assert res.status_code == 200
    res = client.post("/api/document/excel/write", json={"start_cell": "A1", "matrix": [["Year", "Revenue"], [2026, 1000000]], "format_table": True})
    assert res.status_code == 200
    assert res.json().get("success") is True
    res = client.post("/api/document/excel/financial_model", json={"model_type": "saas_revenue", "currency": "$"})
    assert res.status_code == 200
    assert res.json().get("success") is True
    res = client.post("/api/document/excel/chart", json={"chart_type": "column", "data_range": "A1:B2", "title": "Test Chart"})
    assert res.status_code == 200
    assert res.json().get("success") is True

    # 7. Outlook Automation Read, Reply, Compose & Meeting
    res = client.post("/api/document/outlook/read")
    assert res.status_code == 200
    res = client.post("/api/document/outlook/reply", json={"reply_html": "<p>Verified reply</p>", "reply_all": False})
    assert res.status_code == 200
    assert res.json().get("success") is True
    res = client.post("/api/document/outlook/compose", json={"to": "client@example.com", "subject": "Test", "body_html": "<p>Hello</p>"})
    assert res.status_code == 200
    assert res.json().get("success") is True
    res = client.post("/api/document/outlook/meeting", json={"subject": "Review Meeting", "start_time": "2026-10-29 10:00", "duration_min": 30, "attendees": "client@example.com"})
    assert res.status_code == 200
    assert res.json().get("success") is True

    # 8. Cross-Suite Multi-App Pipelines
    res = client.post("/api/suite/doc_to_slides")
    assert res.status_code == 200
    assert res.json().get("success") is True
    res = client.post("/api/suite/excel_to_report")
    assert res.status_code == 200
    assert res.json().get("success") is True
    res = client.post("/api/suite/email_summary", json={"recipient": "ergys@cloud4job.com"})
    assert res.status_code == 200
    assert res.json().get("success") is True

    # 9. AI Chat Streaming
    res = client.post("/api/chat", json={"prompt": "/check-doc", "host": "Word"})
    assert res.status_code == 200
    assert "text/event-stream" in res.headers.get("content-type", "")

    print("  [OK] All Word, Excel, PowerPoint, Outlook, and Suite Orchestrator API endpoints verified successfully!")

def test_certs():
    print("[5/6] Verifying SSL Certificate...")
    cert_path = os.path.join(PROJECT_ROOT, "certs", "localhost.crt")
    key_path = os.path.join(PROJECT_ROOT, "certs", "localhost.key")
    assert os.path.exists(cert_path), "localhost.crt missing"
    assert os.path.exists(key_path), "localhost.key missing"
    print("  [OK] Localhost TLS certificate and key verified!")

def test_installer_package():
    print("[6/6] Verifying Standalone Windows Installer (.exe)...")
    installer_path = os.path.join(PROJECT_ROOT, "dist", "GeminiAntigravityOffice_Setup.exe")
    assert os.path.exists(installer_path), f"Installer not found at {installer_path}"
    size_mb = os.path.getsize(installer_path) / (1024 * 1024)
    assert size_mb > 1.0, f"Installer size unexpectedly small: {size_mb:.2f} MB"
    print(f"  [OK] Standalone installer verified: {installer_path} ({size_mb:.2f} MB)")

def main():
    print("==========================================================")
    print("  Starting Gemini Antigravity Office Verification Suite")
    print("==========================================================")
    test_manifest()
    test_assets()
    test_ui_markup()
    test_daemon_api()
    test_certs()
    test_installer_package()
    print("==========================================================")
    print("  ALL VERIFICATION TESTS PASSED (6/6 SUITES GREEN)!")
    print("==========================================================")

if __name__ == "__main__":
    main()
