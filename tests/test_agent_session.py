import pytest
import asyncio
from daemon.agent_session import AgentSession


def test_agent_session_tool_dispatch_and_resolution():
    async def _run():
        session = AgentSession()
        session.is_running = True

        async def worker():
            return await session.call_tool("word_insert_table", {"headers": ["A", "B"], "rows": [["1", "2"]]}, timeout_sec=5.0)

        task = asyncio.create_task(worker())

        # Wait for event to arrive in queue
        event = await asyncio.wait_for(session.event_queue.get(), timeout=2.0)
        assert event["event"] == "tool_call"
        assert event["tool"] == "word_insert_table"
        call_id = event["call_id"]
        assert call_id in session.pending_calls

        # Resolve tool call
        resolved = session.resolve_tool(call_id, {"ok": True, "table_index": 0})
        assert resolved is True

        result = await asyncio.wait_for(task, timeout=2.0)
        assert result["ok"] is True
        assert result["table_index"] == 0

    asyncio.run(_run())


def test_agent_session_timeout():
    async def _run():
        session = AgentSession()
        session.is_running = True

        result = await session.call_tool("word_search", {"search_text": "nonexistent"}, timeout_sec=0.1)
        assert result["ok"] is False
        assert "Timed out" in result["error"]

    asyncio.run(_run())


def test_agent_session_cancel():
    async def _run():
        session = AgentSession()
        session.is_running = True

        async def worker():
            return await session.call_tool("word_search", {"search_text": "foo"}, timeout_sec=5.0)

        task = asyncio.create_task(worker())
        await asyncio.sleep(0.05)
        await session.cancel()

        result = await task
        assert result["ok"] is False
        assert "cancelled" in result["error"]

    asyncio.run(_run())
