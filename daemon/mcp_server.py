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

# Tools executed by the daemon itself (not forwarded as-is to the taskpane).
DAEMON_TOOLS = [
    {
        "name": "office_verify",
        "description": (
            "Verify your work like an Office expert before you report back. Re-reads the whole document "
            "and compares it with the state at the start of this turn: fields filled, fields still empty, "
            "fields accidentally cleared, tables/rows removed, leftover placeholders, heading level skips, "
            "stale table of contents. Call it after your edits; fix every issue it reports, then call it again."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "expect_all_fields_filled": {
                    "type": "boolean",
                    "description": "True (default) when the user asked to fill/complete the document.",
                }
            },
        },
    }
]
BLOCKS_DOC = (
    "List of blocks. Types: heading{level 1-6,text} | title{text} | subtitle{text} | "
    "paragraph{text | runs[{text,bold,italic,underline,color '#RRGGBB',size,highlight,font}], align left|center|right|justify, style} | "
    "bullets{items} | numbered{items} (item = string or {text, children:[items]}, max 3 levels) | "
    "table{columns[{width_pct}], header_rows=1, repeat_header=true, style, borders, header_shading, "
    "rows[[cell]]} where cell = string or {text|runs, colspan, rowspan, shading, bold, align, valign}; do NOT write the cells "
    "covered by a rowspan/colspan | page_break | caption{label,text} | quote{text} | toc{levels:'1-3'}."
)
DAEMON_TOOLS.append({
    "name": "word_insert_content",
    "description": (
        "FASTEST and preferred way to write any multi-part content (report sections, headings + text, lists, "
        "tables with merged cells and shading, captions, TOC). Everything is inserted in ONE atomic native Word "
        "operation with built-in styles (Heading 1-6, Title, Subtitle, Caption...), which work in any Word language. "
        "Returns block_id; the content is wrapped in content control 'ga:<block_id>' so you can target it later "
        "(location 'after:ga:<id>' or 'replace:ga:<id>')."
    ),
    "inputSchema": {
        "type": "object",
        "properties": {
            "blocks": {"type": "array", "items": {"type": "object"}, "description": BLOCKS_DOC},
            "location": {
                "type": "string",
                "description": "cursor (default) | start | end | replace_selection | after:p<index> | after:<heading text> "
                               "| after:ga:<block_id> | replace:ga:<block_id>",
            },
        },
        "required": ["blocks"],
    },
})
DAEMON_TOOL_NAMES = {t["name"] for t in DAEMON_TOOLS}


async def run_daemon_tool(name: str, args: dict) -> dict:
    if name == "office_verify":
        return await agent_session.verify_document(args)
    if name == "word_insert_content":
        import uuid
        from .office.ooxml_builder import build_ooxml, validate_blocks

        blocks = args.get("blocks") or []
        errors = validate_blocks(blocks)
        if errors:
            return {"ok": False, "error": "Invalid blocks, fix and retry:\n- " + "\n- ".join(errors)}
        block_id = args.get("block_id") or uuid.uuid4().hex[:8]
        ooxml = build_ooxml(blocks, block_id=block_id)
        res = await agent_session.call_tool(
            "word_insert_ooxml", {"ooxml": ooxml, "location": args.get("location") or "cursor", "block_id": block_id})
        return {**res, "block_id": block_id, "blocks": len(blocks)} if res.get("ok") else res
    return {"ok": False, "error": f"Unknown daemon tool {name}"}


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
                "tools": get_tool_definitions() + DAEMON_TOOLS
            }
        })

    # 3. Tools Call
    if method == "tools/call":
        params = body.get("params", {})
        tool_name = params.get("name")
        arguments = params.get("arguments", {})

        if tool_name in DAEMON_TOOL_NAMES:
            tool_result = await run_daemon_tool(tool_name, arguments)
            return JSONResponse({
                "jsonrpc": "2.0",
                "id": msg_id,
                "result": {
                    "content": [{"type": "text", "text": json.dumps(tool_result, ensure_ascii=False)}],
                    "isError": tool_result.get("ok") is False and "error" in tool_result,
                },
            })

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
