"""
Tests for COM Tools wiring into DAEMON_TOOLS and MCP server execution.
Verifies all 16 native Word COM tools from daemon.office.word_com_tools.TOOLS,
execution via ComWorker, document targeting by doc_url, visual insertion ga:<id> anchors,
perception cache invalidation, and error handling.
"""

import asyncio
import json
from unittest.mock import MagicMock, AsyncMock, patch

import pytest
from fastapi.testclient import TestClient

from daemon.main import app
from daemon.agent_session import agent_session
from daemon.core.store import Store
from daemon.mcp_server import DAEMON_TOOLS, DAEMON_TOOL_NAMES, COM_TOOL_NAMES, VISUAL_TOOLS
from daemon.office.word_com_tools import TOOLS as COM_TOOLS

client = TestClient(app)

EXPECTED_16_TOOLS = {
    "word_insert_chart",
    "word_insert_smartart",
    "word_insert_diagram",
    "word_insert_toc",
    "word_insert_caption",
    "word_insert_cross_reference",
    "word_update_fields",
    "word_insert_section_break",
    "word_set_header_footer",
    "word_track_changes",
    "word_add_comment",
    "word_modify_style",
    "word_apply_style",
    "word_link_numbering_to_headings",
    "word_set_checkbox",
    "word_table_clone_row",
}


class MockContentControl:
    def __init__(self, cc_type, rng):
        self.Type = cc_type
        self.Range = rng
        self.Tag = ""
        self.Title = ""


class MockBookmarks:
    def __init__(self):
        self._bms = {}

    def Add(self, name, rng):
        self._bms[name] = rng
        return rng

    def Exists(self, name):
        return name in self._bms


class MockContentControls:
    def __init__(self):
        self._ccs = []

    def Add(self, cc_type, rng):
        cc = MockContentControl(cc_type, rng)
        self._ccs.append(cc)
        return cc

    def SelectByTag(self, tag):
        matching = [c for c in self._ccs if c.Tag == tag]
        res = MagicMock()
        res.Count = len(matching)
        res.side_effect = lambda i: matching[i - 1]
        return res


class MockDocument:
    def __init__(self, full_name="C:\\docs\\report.docx", name="report.docx"):
        self.FullName = full_name
        self.Name = name
        self.TrackRevisions = False
        self._bms = MockBookmarks()
        self.Bookmarks = self._bms
        self._ccs = MockContentControls()
        self.ContentControls = self._ccs
        self.SelectContentControlsByTag = self._ccs.SelectByTag

        # Range and content
        self.Content = MagicMock()
        self.Content.WordOpenXML = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>'

        # Shapes & InlineShapes collections
        self._inline_shapes = []
        self._shapes = {}

        mock_ish_coll = MagicMock()
        mock_ish_coll.Count = 0
        mock_ish_coll.side_effect = lambda i: self._inline_shapes[i - 1]
        self.InlineShapes = mock_ish_coll

        mock_sh_coll = MagicMock()
        mock_sh_coll.Count = 0
        mock_sh_coll.side_effect = lambda key: self._shapes[key] if isinstance(key, str) else list(self._shapes.values())[key - 1]
        self.Shapes = mock_sh_coll

        # UndoRecord
        app_mock = MagicMock()
        undo_mock = MagicMock()
        undo_mock.IsRecordingCustomRecord = False
        app_mock.UndoRecord = undo_mock
        self.Application = app_mock

    def add_mock_inline_shape(self):
        ish = MagicMock()
        ish.Range = MagicMock()
        self._inline_shapes.append(ish)
        self.InlineShapes.Count = len(self._inline_shapes)
        return ish

    def add_mock_shape(self, name):
        sh = MagicMock()
        sh.Name = name
        sh.Anchor = MagicMock()
        self._shapes[name] = sh
        self.Shapes.Count = len(self._shapes)
        return sh


class MockWorker:
    def __init__(self, doc=None):
        self.doc = doc or MockDocument()
        self.calls = []
        self.requested_urls = []

    def get_document(self, doc_url=None):
        self.requested_urls.append(doc_url)
        return self.doc

    async def run(self, fn, *args, **kwargs):
        self.calls.append((fn, args, kwargs))
        if asyncio.iscoroutinefunction(fn):
            return await fn(*args, **kwargs)
        return fn(*args, **kwargs)


@pytest.fixture(autouse=True)
def setup_test_env(tmp_path):
    store = Store(tmp_path / "com_wiring.db")
    doc_url = "C:\\docs\\report.docx"
    chat = store.create_chat(host="Word", doc_url=doc_url)
    agent_session.configure(store=store)
    agent_session._current_chat = chat["id"]
    agent_session._current_doc_url = doc_url
    agent_session.turn_snapshot = {"stats": {"words": 10}, "doc_type": "report"}
    agent_session.invalidate_perception_cache()
    mock_worker = MockWorker()
    agent_session.com_worker = mock_worker
    yield mock_worker
    agent_session._current_doc_url = None
    agent_session._com_worker = None


def test_all_16_com_tools_registered():
    """Requirement 1: Verify all 16 COM tools are present in DAEMON_TOOLS."""
    assert len(EXPECTED_16_TOOLS) == 16
    assert EXPECTED_16_TOOLS == set(COM_TOOLS.keys())
    daemon_names = {t["name"] for t in DAEMON_TOOLS}
    for tool_name in EXPECTED_16_TOOLS:
        assert tool_name in daemon_names, f"{tool_name} missing from DAEMON_TOOLS"
        assert tool_name in DAEMON_TOOL_NAMES, f"{tool_name} missing from DAEMON_TOOL_NAMES"


def test_mcp_tools_list_contains_com_tools_without_duplicates():
    """Verify tools/list endpoint returns all COM tools and no duplicate names."""
    res = client.post("/mcp", json={"jsonrpc": "2.0", "id": 1, "method": "tools/list"}).json()
    assert "result" in res
    tools = res["result"]["tools"]
    names = [t["name"] for t in tools]
    # Check no duplicate tool names
    assert len(names) == len(set(names)), f"Duplicate tool names found: {[n for n in names if names.count(n) > 1]}"
    # Verify all 16 COM tools appear in tools/list
    for expected in EXPECTED_16_TOOLS:
        assert expected in names


def test_com_tool_execution_via_worker_targets_doc_url(setup_test_env):
    """Requirement 2: Execute tool through ComWorker, targeting agent_session._current_doc_url."""
    mock_worker = setup_test_env
    agent_session._current_doc_url = "C:\\docs\\report.docx"

    payload = {
        "jsonrpc": "2.0",
        "id": 10,
        "method": "tools/call",
        "params": {
            "name": "word_track_changes",
            "arguments": {"on": True}
        }
    }
    res = client.post("/mcp", json=payload)
    assert res.status_code == 200
    data = res.json()
    content = json.loads(data["result"]["content"][0]["text"])

    assert content["ok"] is True
    assert content["track_changes"] is True
    assert mock_worker.doc.TrackRevisions is True
    assert "C:\\docs\\report.docx" in mock_worker.requested_urls


def test_visual_tool_chart_anchor_wrapping_and_ga_return(setup_test_env):
    """Requirement 3: word_insert_chart wraps with ContentControl/Bookmark and returns ga:<id>."""
    mock_worker = setup_test_env
    # Prepare mock inline shape for chart
    ish = mock_worker.doc.add_mock_inline_shape()

    # Mock word_insert_chart in COM_TOOLS to return success with inline_shape_index
    with patch.dict(COM_TOOLS, {
        "word_insert_chart": {
            "fn": lambda doc, **kwargs: {"ok": True, "chart_type": "column", "inline_shape_index": 0, "anchor": "i0"},
            "description": "Mock chart",
            "inputSchema": {"type": "object"}
        }
    }):
        payload = {
            "jsonrpc": "2.0",
            "id": 20,
            "method": "tools/call",
            "params": {
                "name": "word_insert_chart",
                "arguments": {"type": "column", "categories": ["Q1"], "series": [{"name": "Rev", "values": [10]}], "id": "my_chart_1"}
            }
        }
        res = client.post("/mcp", json=payload)
        assert res.status_code == 200
        content = json.loads(res.json()["result"]["content"][0]["text"])

        assert content["ok"] is True
        assert content["anchor"] == "ga:my_chart_1"
        assert content["anchor_id"] == "my_chart_1"
        assert content["bookmark"] == "ga_my_chart_1"

        # Check that ContentControl was added with tag my_chart_1
        assert len(mock_worker.doc._ccs._ccs) == 1
        cc = mock_worker.doc._ccs._ccs[0]
        assert cc.Tag == "my_chart_1"
        assert cc.Title == "ga:my_chart_1"

        # Check that Bookmark was added
        assert mock_worker.doc.Bookmarks.Exists("ga_my_chart_1")


def test_visual_tool_smartart_anchor_wrapping_and_ga_return(setup_test_env):
    """Requirement 3: word_insert_smartart wraps and returns ga:<id>."""
    mock_worker = setup_test_env
    ish = mock_worker.doc.add_mock_inline_shape()

    with patch.dict(COM_TOOLS, {
        "word_insert_smartart": {
            "fn": lambda doc, **kwargs: {"ok": True, "layout": "process", "anchor": "i0"},
            "description": "Mock SmartArt",
            "inputSchema": {"type": "object"}
        }
    }):
        payload = {
            "jsonrpc": "2.0",
            "id": 21,
            "method": "tools/call",
            "params": {
                "name": "word_insert_smartart",
                "arguments": {"layout": "process", "nodes": [{"text": "Step 1"}], "anchor_id": "sa_process"}
            }
        }
        res = client.post("/mcp", json=payload)
        assert res.status_code == 200
        content = json.loads(res.json()["result"]["content"][0]["text"])

        assert content["ok"] is True
        assert content["anchor"] == "ga:sa_process"
        assert content["bookmark"] == "ga_sa_process"
        assert mock_worker.doc._ccs._ccs[-1].Tag == "sa_process"
        assert mock_worker.doc.Bookmarks.Exists("ga_sa_process")


def test_visual_tool_diagram_anchor_wrapping_and_ga_return(setup_test_env):
    """Requirement 3: word_insert_diagram wraps canvas shape and returns ga:<id>."""
    mock_worker = setup_test_env
    sh = mock_worker.doc.add_mock_shape("canvas_diag_1")

    with patch.dict(COM_TOOLS, {
        "word_insert_diagram": {
            "fn": lambda doc, **kwargs: {"ok": True, "kind": "flowchart", "name": "canvas_diag_1"},
            "description": "Mock Diagram",
            "inputSchema": {"type": "object"}
        }
    }):
        payload = {
            "jsonrpc": "2.0",
            "id": 22,
            "method": "tools/call",
            "params": {
                "name": "word_insert_diagram",
                "arguments": {"kind": "flowchart", "nodes": [{"id": "a", "text": "Start"}]}
            }
        }
        res = client.post("/mcp", json=payload)
        assert res.status_code == 200
        content = json.loads(res.json()["result"]["content"][0]["text"])

        assert content["ok"] is True
        assert content["anchor"].startswith("ga:")
        anchor_id = content["anchor_id"]
        assert mock_worker.doc._ccs._ccs[-1].Tag == anchor_id
        assert mock_worker.doc.Bookmarks.Exists(content["bookmark"])


def test_perception_cache_invalidated_after_com_tool(setup_test_env):
    """Requirement 4: Invalidate perception snapshot cache after COM tool execution."""
    mock_worker = setup_test_env

    # Populate perception cache with a dummy entry
    perception = agent_session.perception
    perception._cache["cached_key"] = {"snapshot": {"fake": True}, "map": "map", "summary": {}}
    assert "cached_key" in perception._cache

    # Execute a COM tool
    payload = {
        "jsonrpc": "2.0",
        "id": 30,
        "method": "tools/call",
        "params": {
            "name": "word_track_changes",
            "arguments": {"on": False}
        }
    }
    res = client.post("/mcp", json=payload)
    assert res.status_code == 200

    # Verify stale cache was invalidated (cached_key removed)
    assert "cached_key" not in perception._cache


def test_com_tool_error_handling(setup_test_env):
    """Error handling: Word failure or document error is cleanly returned as ok: False."""
    mock_worker = setup_test_env

    # Failure inside COM tool function
    with patch.dict(COM_TOOLS, {
        "word_update_fields": {
            "fn": lambda doc, **kwargs: {"ok": False, "error": "Word is busy in modal dialog"},
            "description": "Mock update",
            "inputSchema": {"type": "object"}
        }
    }):
        payload = {
            "jsonrpc": "2.0",
            "id": 40,
            "method": "tools/call",
            "params": {
                "name": "word_update_fields",
                "arguments": {}
            }
        }
        res = client.post("/mcp", json=payload)
        assert res.status_code == 200
        content = json.loads(res.json()["result"]["content"][0]["text"])
        assert content["ok"] is False
        assert "Word is busy" in content["error"]


def test_com_tool_document_not_open(setup_test_env):
    """Error handling: When worker.get_document raises error, tool returns ok: False."""
    mock_worker = setup_test_env
    mock_worker.get_document = MagicMock(side_effect=RuntimeError("No Word document is open"))

    payload = {
        "jsonrpc": "2.0",
        "id": 41,
        "method": "tools/call",
        "params": {
            "name": "word_track_changes",
            "arguments": {"on": True}
        }
    }
    res = client.post("/mcp", json=payload)
    assert res.status_code == 200
    content = json.loads(res.json()["result"]["content"][0]["text"])
    assert content["ok"] is False
    assert "No Word document is open" in content["error"]


def test_legacy_arg_adaptation(setup_test_env):
    """Verify legacy paragraph_index and comment parameters are adapted."""
    mock_worker = setup_test_env
    captured_args = {}

    def mock_apply(doc, **kwargs):
        captured_args.update(kwargs)
        return {"ok": True}

    with patch.dict(COM_TOOLS, {
        "word_apply_style": {
            "fn": mock_apply,
            "description": "Apply style",
            "inputSchema": {"type": "object"}
        }
    }):
        payload = {
            "jsonrpc": "2.0",
            "id": 50,
            "method": "tools/call",
            "params": {
                "name": "word_apply_style",
                "arguments": {"paragraph_index": 3, "style": "Heading 2"}
            }
        }
        res = client.post("/mcp", json=payload)
        assert res.status_code == 200
        assert captured_args.get("anchor") == "p3"
        assert captured_args.get("style") == "Heading 2"


def test_targeting_caption_and_cross_reference_after_visual_insert(setup_test_env):
    """Verify visual insertion anchor can be targeted by caption and cross-reference."""
    mock_worker = setup_test_env
    ish = mock_worker.doc.add_mock_inline_shape()

    # 1. Insert chart with specific ID
    with patch.dict(COM_TOOLS, {
        "word_insert_chart": {
            "fn": lambda doc, **kwargs: {"ok": True, "chart_type": "column", "inline_shape_index": 0, "anchor": "i0"},
            "description": "Mock chart",
            "inputSchema": {"type": "object"}
        }
    }):
        payload = {
            "jsonrpc": "2.0",
            "id": 60,
            "method": "tools/call",
            "params": {
                "name": "word_insert_chart",
                "arguments": {"type": "column", "categories": ["A"], "series": [{"name": "S", "values": [1]}], "id": "fig_rev_2026"}
            }
        }
        res = client.post("/mcp", json=payload)
        chart_result = json.loads(res.json()["result"]["content"][0]["text"])
        assert chart_result["ok"] is True
        assert chart_result["anchor"] == "ga:fig_rev_2026"
        assert chart_result["bookmark"] == "ga_fig_rev_2026"

    # 2. Add caption targeting the returned anchor ga:fig_rev_2026
    caption_called_with = {}
    def mock_caption(doc, **kwargs):
        caption_called_with.update(kwargs)
        # Check that doc.SelectContentControlsByTag finds the control
        tag = kwargs.get("anchor", "")[3:]
        ccs = doc.SelectContentControlsByTag(tag)
        assert ccs.Count >= 1
        return {"ok": True, "label": "Figure", "text": "Figure 1: Revenue"}

    with patch.dict(COM_TOOLS, {
        "word_insert_caption": {
            "fn": mock_caption,
            "description": "Insert caption",
            "inputSchema": {"type": "object"}
        }
    }):
        payload = {
            "jsonrpc": "2.0",
            "id": 61,
            "method": "tools/call",
            "params": {
                "name": "word_insert_caption",
                "arguments": {"anchor": chart_result["anchor"], "label": "Figure", "text": "Revenue"}
            }
        }
        res = client.post("/mcp", json=payload)
        cap_result = json.loads(res.json()["result"]["content"][0]["text"])
        assert cap_result["ok"] is True
        assert caption_called_with["anchor"] == "ga:fig_rev_2026"

    # 3. Insert cross reference targeting ga:fig_rev_2026 normalized to bookmark
    xref_called_with = {}
    def mock_xref(doc, **kwargs):
        xref_called_with.update(kwargs)
        # Verify the bookmark exists in doc
        bm = kwargs.get("target_anchor")
        assert doc.Bookmarks.Exists(bm)
        return {"ok": True, "text": "Figure 1"}

    with patch.dict(COM_TOOLS, {
        "word_insert_cross_reference": {
            "fn": mock_xref,
            "description": "Insert xref",
            "inputSchema": {"type": "object"}
        }
    }):
        payload = {
            "jsonrpc": "2.0",
            "id": 62,
            "method": "tools/call",
            "params": {
                "name": "word_insert_cross_reference",
                "arguments": {"kind": "bookmark", "target_anchor": chart_result["anchor"]}
            }
        }
        res = client.post("/mcp", json=payload)
        xref_result = json.loads(res.json()["result"]["content"][0]["text"])
        assert xref_result["ok"] is True
        assert xref_called_with["target_anchor"] == "ga_fig_rev_2026"


def test_all_16_tools_callable_via_mcp(setup_test_env):
    """Verify every single one of the 16 COM tools can be called through the MCP endpoint."""
    mock_worker = setup_test_env

    for idx, tool_name in enumerate(sorted(EXPECTED_16_TOOLS)):
        called = {"status": False}

        def make_mock_fn(name):
            def _fn(doc, **kwargs):
                called["status"] = True
                return {"ok": True, "tool": name}
            return _fn

        with patch.dict(COM_TOOLS, {
            tool_name: {
                "fn": make_mock_fn(tool_name),
                "description": f"Mock {tool_name}",
                "inputSchema": {"type": "object"}
            }
        }):
            payload = {
                "jsonrpc": "2.0",
                "id": 100 + idx,
                "method": "tools/call",
                "params": {
                    "name": tool_name,
                    "arguments": {}
                }
            }
            res = client.post("/mcp", json=payload)
            assert res.status_code == 200
            data = res.json()
            assert "result" in data, f"Tool {tool_name} failed: {data}"
            content = json.loads(data["result"]["content"][0]["text"])
            assert content.get("ok") is True, f"Tool {tool_name} returned ok=False: {content}"
            assert called["status"] is True, f"Tool function for {tool_name} was not called"


def test_verify_document_with_com_worker_fallback(setup_test_env):
    """Verify office_verify clears cache and uses com_worker fallback when taskpane is offline."""
    mock_worker = setup_test_env
    agent_session.is_running = False  # Taskpane offline
    agent_session.turn_snapshot = {
        "stats": {"words": 10, "fields_total": 0, "fields_empty": 0, "tables": 0},
        "doc_type": "document",
        "fields": [],
        "tables": []
    }

    # Pre-populate perception cache
    agent_session.perception._cache["stale_key"] = {"snapshot": {"old": True}, "map": "stale", "summary": {}}
    assert "stale_key" in agent_session.perception._cache

    payload = {
        "jsonrpc": "2.0",
        "id": 200,
        "method": "tools/call",
        "params": {
            "name": "office_verify",
            "arguments": {"expect_all_fields_filled": False}
        }
    }
    res = client.post("/mcp", json=payload)
    assert res.status_code == 200
    data = res.json()
    content = json.loads(data["result"]["content"][0]["text"])

    assert content["ok"] is True
    # Perception cache was cleared before verification
    assert "stale_key" not in agent_session.perception._cache


def test_refresh_perception_method(setup_test_env):
    """Verify agent_session.refresh_perception() reads OOXML and parses snapshot."""
    mock_worker = setup_test_env
    # Clear perception
    agent_session.perception._cache.clear()

    snapshot = asyncio.run(agent_session.refresh_perception("C:\\docs\\report.docx"))
    assert snapshot is not None
    assert "stats" in snapshot
    # Cache should now have the entry
    assert len(agent_session.perception._cache) > 0


def test_agent_session_com_worker_lifecycle():
    """Verify com_worker property lazy instantiation, setter and shutdown."""
    agent_session._com_worker = None
    # Lazy property creates ComWorker
    worker = agent_session.com_worker
    assert worker is not None

    # Setter allows injection
    custom_worker = MockWorker()
    agent_session.com_worker = custom_worker
    assert agent_session.com_worker is custom_worker

    # Shutdown stops worker
    custom_worker.stop = MagicMock()
    asyncio.run(agent_session.shutdown())
    assert custom_worker.stop.called
    assert agent_session._com_worker is None
