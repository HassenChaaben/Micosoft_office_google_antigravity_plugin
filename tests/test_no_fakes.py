import asyncio
import json
import os
import subprocess
from unittest.mock import MagicMock, patch

import pytest

from daemon.auth import GeminiAuthBridge, auth_bridge
from daemon.gemini_client import GeminiClient
from daemon.mcp_bridge import MCPBridge, mcp_bridge
from daemon.word_engine import WordEngine, word_engine


# =============================================================================
# 1. WordEngine: No NameError and No Disguising Errors as Success
# =============================================================================

def test_word_engine_replace_selection_assigns_doc_on_error():
    engine = WordEngine()

    mock_app = MagicMock()
    mock_doc = MagicMock()
    mock_app.ActiveDocument = mock_doc
    # Simulate Selection.InsertFile failing, then DocContent.InsertFile failing
    mock_app.Selection.InsertFile.side_effect = Exception("COM InsertFile failed")
    mock_doc.Content.InsertFile.side_effect = Exception("DocContent InsertFile failed")
    # Simulate Content.InsertAfter raising an error
    mock_app.Selection = None
    mock_doc.Content.InsertAfter.side_effect = Exception("COM InsertAfter failed")

    with patch.object(engine, "_get_word_app", return_value=mock_app):
        result = engine.replace_selection("# Markdown Heading\n| col1 | col2 |")
        # Must report failure, not success!
        assert result.get("success") is False
        assert "error" in result
        assert "InsertAfter failed" in result["error"]
        # Ensure no NameError occurred for 'doc'
        assert "name 'doc' is not defined" not in result["error"]


def test_word_engine_replace_selection_plain_text_error_reported():
    engine = WordEngine()

    mock_app = MagicMock()
    mock_doc = MagicMock()
    mock_app.ActiveDocument = mock_doc
    mock_app.Selection = None
    mock_doc.Content.InsertAfter.side_effect = RuntimeError("Disk full")

    with patch.object(engine, "_get_word_app", return_value=mock_app):
        result = engine.replace_selection("plain text replacement")
        assert result.get("success") is False
        assert result.get("error") == "Disk full"


def test_word_engine_insert_comment_error_not_disguised():
    engine = WordEngine()

    mock_app = MagicMock()
    mock_doc = MagicMock()
    mock_app.ActiveDocument = mock_doc
    mock_doc.Comments.Add.side_effect = Exception("Permission denied")

    with patch.object(engine, "_get_word_app", return_value=mock_app):
        result = engine.insert_comment("Test comment")
        assert result.get("success") is False
        assert "Permission denied" in result.get("error", "")


# =============================================================================
# 2. GeminiClient: No Canned Fake Responses
# =============================================================================

def _collect_stream(agen):
    async def run():
        chunks = []
        async for chunk in agen:
            chunks.append(chunk)
        return "".join(chunks)
    return asyncio.run(run())


def test_gemini_client_offline_mode_returns_honest_mock():
    client = GeminiClient()
    response = client._generate_mock_response("/check-doc", "Word")
    assert response == "[Mock offline response for Word: /check-doc]"

    # Verify none of the removed canned texts exist
    assert "Superconducting Transmons" not in response
    assert "Daily Tasks Tracking Matrix" not in response
    assert "TSK-101" not in response
    assert "Executive Contract Summary" not in response


def test_gemini_client_no_canned_answers_for_arbitrary_prompts():
    client = GeminiClient()
    quantum_resp = client._generate_mock_response("quantum computing", "Word")
    assert "Superconducting Transmons" not in quantum_resp
    assert quantum_resp == "[Mock offline response for Word: quantum computing]"

    table_resp = client._generate_mock_response("create a table of daily tasks", "Word")
    assert "Daily Tasks Tracking Matrix" not in table_resp
    assert table_resp == "[Mock offline response for Word: create a table of daily tasks]"


def test_gemini_client_offline_streaming_uses_mock_response(monkeypatch):
    monkeypatch.setenv("GEMINI_OFFICE_BACKEND", "offline")
    client = GeminiClient()
    body = _collect_stream(client.stream_chat({"prompt": "test prompt", "host": "Word"}))
    deltas = "".join(json.loads(line[6:])["delta"] for line in body.splitlines() if line.startswith("data: {") and "delta" in line)
    assert deltas == "[Mock offline response for Word: test prompt]"
    assert "data: [DONE]" in body


# =============================================================================
# 3. MCPBridge: Real CLI Queries (agy mcp list & agy plugin list)
# =============================================================================

def test_mcp_bridge_query_agy_mcp_list_parser():
    bridge = MCPBridge()
    sample_stdout = (
        "NAME         TYPE  STATUS   COMMAND/URL\n"
        "word-office  http  enabled  https://127.0.0.1:18890/mcp\n"
        "custom-mcp   stdio enabled  C:\\tools\\custom.exe\n"
    )

    mock_res = MagicMock()
    mock_res.returncode = 0
    mock_res.stdout = sample_stdout

    with patch("subprocess.run", return_value=mock_res):
        conns = bridge._query_agy_mcp_list("fake_agy.exe")
        assert len(conns) == 2
        assert conns[0]["id"] == "word-office"
        assert conns[0]["type"] == "mcp"
        assert conns[0]["status"] == "enabled"
        assert conns[0]["target"] == "https://127.0.0.1:18890/mcp"
        assert conns[1]["id"] == "custom-mcp"


def test_mcp_bridge_query_agy_plugin_list_parser():
    bridge = MCPBridge()
    sample_stdout = (
        "Installed plugins:\n"
        "- notion (enabled, 1.0.0)\n"
        "    Components: 1 skill\n"
        "- github (enabled, 1.2.0)\n"
        "- gdrive (disabled, 0.0.8)\n"
    )

    mock_res = MagicMock()
    mock_res.returncode = 0
    mock_res.stdout = sample_stdout

    with patch("subprocess.run", return_value=mock_res):
        plugins = bridge._query_agy_plugin_list("fake_agy.exe")
        assert len(plugins) == 3
        assert plugins[0]["id"] == "notion"
        assert plugins[0]["status"] == "enabled"
        assert plugins[0]["version"] == "1.0.0"
        assert plugins[1]["id"] == "github"
        assert plugins[2]["id"] == "gdrive"
        assert plugins[2]["status"] == "disabled"


def test_mcp_bridge_query_connector_no_hardcoded_items():
    bridge = MCPBridge()
    # Query an unconfigured connector
    res = bridge.query_connector("unknown_connector", "search term")
    assert res.get("status") == "not_found"
    assert "items" not in res

    # Query an active connector
    bridge.connectors = [{"id": "notion", "name": "Notion", "status": "active", "type": "mcp"}]
    res_notion = bridge.query_connector("notion", "search term")
    assert res_notion.get("source") == "notion"
    assert "Sprint Goals 2026" not in str(res_notion)
    assert "Project Overview" not in str(res_notion)


def test_mcp_bridge_live_or_filesystem_connectors():
    conns = mcp_bridge.get_connector_list()
    ids = [c["id"] for c in conns]
    assert "fs" in ids or "filesystem" in ids


# =============================================================================
# 4. GeminiAuthBridge: Real Auth Checks via Windows Credential Manager & agy models
# =============================================================================

def test_auth_bridge_unauthenticated_when_no_credentials():
    bridge = GeminiAuthBridge.__new__(GeminiAuthBridge)
    bridge.session_data = {}

    with patch.object(bridge, "_read_win32_credential", return_value=None), \
         patch.object(bridge, "_check_agy_models", return_value={"ok": False, "error": "Not logged in"}), \
         patch.object(bridge, "get_api_key", return_value=None), \
         patch.dict(os.environ, {}, clear=True), \
         patch("builtins.open", side_effect=FileNotFoundError):
        session = bridge._load_session()
        assert session["authenticated"] is False
        assert session["tier"] == "unauthenticated"
        assert session["token"] == ""
        # Must not contain dummy session tokens
        assert session["token"] != "sess_gemini_active_bridge"


def test_auth_bridge_reads_real_windows_credential():
    bridge = GeminiAuthBridge.__new__(GeminiAuthBridge)
    bridge.session_data = {}

    mock_blob = {
        "id_token_email": "real.user@example.com",
        "token": {"access_token": "ya29.real_oauth_token"},
        "name": "Real User",
    }

    with patch.object(bridge, "_read_win32_credential", return_value=mock_blob), \
         patch.object(bridge, "_save_session"):
        session = bridge._load_session()
        assert session["authenticated"] is True
        assert session["email"] == "real.user@example.com"
        assert session["token"] == "ya29.real_oauth_token"
        assert session["source"] == "windows_credential_manager"


def test_auth_bridge_agy_models_check():
    bridge = GeminiAuthBridge.__new__(GeminiAuthBridge)

    mock_res = MagicMock()
    mock_res.returncode = 0
    mock_res.stdout = "Fetching available models...\ngemini-3.8-flash-high\ngemini-pro\n"

    with patch("daemon.auth.find_agy", return_value="agy.exe"), \
         patch("subprocess.run", return_value=mock_res):
        res = bridge._check_agy_models()
        assert res["ok"] is True
        assert res["models_count"] == 2
