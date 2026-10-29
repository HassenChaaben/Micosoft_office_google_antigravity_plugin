"""
Internal Model Context Protocol (MCP) Server for Microsoft Word
Implements JSON-RPC 2.0 endpoints for Streamable HTTP MCP communication.
Exposes word-office tools to the local Antigravity AI CLI (`agy`).
"""

import json
from fastapi import APIRouter, Request, Response
from fastapi.responses import JSONResponse

from .word_tools import get_tool_definitions, get_tool
from .agent_session import agent_session

router = APIRouter()


@router.get("/mcp")
async def mcp_get():
    """Streamable HTTP MCP probe / health check."""
    return Response(content="Word Office MCP Server Active", media_type="text/plain")


@router.post("/mcp")
async def mcp_post(request: Request):
    """
    Handles JSON-RPC 2.0 messages from the Antigravity CLI MCP client.
    Methods: initialize, notifications/initialized, tools/list, tools/call
    """
    try:
        body = await request.json()
    except Exception:
        return JSONResponse(
            {"jsonrpc": "2.0", "id": None, "error": {"code": -32700, "message": "Parse error"}},
            status_code=400
        )

    method = body.get("method")
    msg_id = body.get("id")

    # Notifications do not have an ID and do not return a JSON-RPC body
    if msg_id is None:
        return Response(status_code=202)

    # 1. Initialize
    if method == "initialize":
        client_version = body.get("params", {}).get("protocolVersion", "2025-03-26")
        return JSONResponse({
            "jsonrpc": "2.0",
            "id": msg_id,
            "result": {
                "protocolVersion": client_version,
                "capabilities": {
                    "tools": {}
                },
                "serverInfo": {
                    "name": "word-office",
                    "version": "1.0.0"
                }
            }
        })

    # 2. Tools List
    if method == "tools/list":
        return JSONResponse({
            "jsonrpc": "2.0",
            "id": msg_id,
            "result": {
                "tools": get_tool_definitions()
            }
        })

    # 3. Tools Call
    if method == "tools/call":
        params = body.get("params", {})
        tool_name = params.get("name")
        arguments = params.get("arguments", {})

        tool_def = get_tool(tool_name)
        if not tool_def:
            return JSONResponse({
                "jsonrpc": "2.0",
                "id": msg_id,
                "error": {
                    "code": -32601,
                    "message": f"Unknown tool: '{tool_name}'"
                }
            })

        # Delegate execution to the active Word taskpane session
        tool_result = await agent_session.call_tool(tool_name, arguments)

        return JSONResponse({
            "jsonrpc": "2.0",
            "id": msg_id,
            "result": {
                "content": [
                    {
                        "type": "text",
                        "text": json.dumps(tool_result, ensure_ascii=False)
                    }
                ]
            }
        })

    # Method not found
    return JSONResponse({
        "jsonrpc": "2.0",
        "id": msg_id,
        "error": {
            "code": -32601,
            "message": f"Method not found: {method}"
        }
    })
