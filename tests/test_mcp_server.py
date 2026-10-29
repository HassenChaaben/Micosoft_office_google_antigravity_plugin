import json
import pytest
import asyncio
from fastapi.testclient import TestClient
from daemon.main import app
from daemon.agent_session import agent_session

client = TestClient(app)


def test_mcp_health_probe():
    res = client.get("/mcp")
    assert res.status_code == 200
    assert "Word Office MCP" in res.text


def test_mcp_initialize():
    payload = {
        "jsonrpc": "2.0",
        "id": 1,
        "method": "initialize",
        "params": {
            "protocolVersion": "2025-03-26",
            "capabilities": {}
        }
    }
    res = client.post("/mcp", json=payload)
    assert res.status_code == 200
    data = res.json()
    assert data["id"] == 1
    assert data["result"]["serverInfo"]["name"] == "word-office"
    assert "tools" in data["result"]["capabilities"]


def test_mcp_tools_list():
    payload = {
        "jsonrpc": "2.0",
        "id": 2,
        "method": "tools/list",
        "params": {}
    }
    res = client.post("/mcp", json=payload)
    assert res.status_code == 200
    data = res.json()
    tools = data["result"]["tools"]
    names = [t["name"] for t in tools]
    assert "word_get_document_outline" in names
    assert "word_insert_table" in names
    assert "word_format_table" in names


def test_mcp_tools_call_without_active_session():
    # When no agent is active in Word
    agent_session.is_running = False
    payload = {
        "jsonrpc": "2.0",
        "id": 3,
        "method": "tools/call",
        "params": {
            "name": "word_get_document_outline",
            "arguments": {}
        }
    }
    res = client.post("/mcp", json=payload)
    assert res.status_code == 200
    data = res.json()
    content = json.loads(data["result"]["content"][0]["text"])
    assert content["ok"] is False
    assert "Word is not currently connected" in content["error"]


def test_mcp_unknown_tool():
    payload = {
        "jsonrpc": "2.0",
        "id": 4,
        "method": "tools/call",
        "params": {
            "name": "non_existent_tool",
            "arguments": {}
        }
    }
    res = client.post("/mcp", json=payload)
    assert res.status_code == 200
    data = res.json()
    assert "error" in data
    assert "Unknown tool" in data["error"]["message"]
