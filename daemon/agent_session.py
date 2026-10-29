"""
Autonomous Office Agent Session Orchestrator
Coordinates communication between the local Antigravity AI CLI (`agy`),
the internal Model Context Protocol (MCP) server, and the Microsoft Word Taskpane.
"""

import os
import json
import uuid
import asyncio
import subprocess
from typing import Dict, Any, Optional, AsyncGenerator

from . import antigravity_backend
from .word_tools import is_destructive

AGENT_SYSTEM_RULES = (
    "You are Gemini Antigravity, an expert autonomous AI agent operating directly inside Microsoft Word.\n"
    "You manipulate Word like a human expert by executing tools from the 'word-office' MCP server.\n\n"
    "CORE OPERATIONAL RULES:\n"
    "1. OPERATE VIA TOOLS ONLY: To create tables, format text, add headings, modify paragraphs, or insert pictures, always call the appropriate `word_*` tool. Never dump raw markdown, HTML, or code in your final reply.\n"
    "2. YOU ALREADY SEE THE DOCUMENT: a DOCUMENT MAP (type, styles, tables with merged cells, every field to fill with its path and T/r/c location, outline, selection, headers/footers) is given below, analysed just before this turn. Start by stating in one line what kind of document it is and what must be preserved, then act. Do NOT call `word_get_document_outline` just to look around; use read tools only for details the map does not show. When the map says DIRECT FORMATTING is used, never change fonts, bold, sizes or shading of existing cells.\n"
    "3. IN-PLACE MODIFICATIONS (NEVER DUPLICATE TEXT):\n"
    "   - When the user asks to modify, rewrite, change, fix, expand, or correct an existing paragraph or text: DO NOT CALL `word_insert_paragraph`! `word_insert_paragraph` creates a duplicate paragraph!\n"
    "   - Instead, find the matching paragraph index from the outline or `selected_text`, and call `word_set_paragraph_text` (to rewrite it) or `word_replace_text` (to replace specific phrases) or `word_format_text` (for styling).\n"
    "   - When the user asks to enhance, format, resize, or style an existing table: inspect the outline to find `table_index` and call `word_format_table`, `word_add_table_row`, or `word_set_table_cell`. Never create a second table!\n"
    "4. CREATING TABLES: When asked for a new table, call `word_insert_table` ONCE with exact requested columns in `headers` and requested row count in `rows`.\n"
    "5. INSERTING PICTURES & VISUALS: When asked to add a picture, logo, diagram, or chart, call `word_insert_image` with the URL, keyword ('logo', 'chart', 'diagram', 'banner'), or base64 data, and an optional caption.\n"
    "6. TABLE & DOCUMENT MANAGEMENT: Use `word_read_table(table_index)` to read table cells and `word_delete_table(table_index)` to remove a table if needed.\n"
    "7. BUILT-IN STYLES: Use standard built-in style names: 'Heading 1', 'Heading 2', 'Title', 'Normal'. Localization (French Titre 1, Italian Titolo 1) is handled automatically.\n"
    "8. CONCISE SUMMARY: When finished with all tool actions, provide a single friendly sentence summarizing the changes made in Word.\n"
    "9. NO LOCAL OS/FILE EDITS: Never call file, shell, terminal, or browser tools. Only use `word-office` MCP tools.\n"
    "10. TEMPLATES, FORMS & PM DOCUMENTS:\n"
    "   - When working with templates (Project Charter, Business Case, Risk Register, etc.), forms, or tables with fields to fill:\n"
    "   - Call `word_fill_form_fields(fields={label: value})` to populate fields automatically by matching labels (e.g. {'Codice': '...', 'Titolo': '...', 'Budget complessivo': '...'}). This is completely immune to merged cells and row shifts!\n"
    "   - Or call `word_batch_update_table_cells(table_index, updates=[{row, col, text}])` to populate or update multiple specific cells in a single atomic batch.\n"
    "   - When the user asks to modify a field in an existing form or table (e.g. 'Change the budget to €75.000' or 'Update sponsor to Cloud4Job'), use `word_fill_form_fields` or `word_batch_update_table_cells` to modify it directly in-place without touching other cells.\n"
    "11. FORMATTING & ALIGNMENT (CENTER, SUBTITLES, FONTS):\n"
    "   - When the user asks to center text (e.g. 'located in center', 'align center', 'centrato'): call `word_format_text(paragraph_index, alignment='center')`!\n"
    "   - When the user asks for a subtitle (e.g. 'make it subtitle', 'sottotitolo'): call `word_apply_style(paragraph_index, 'Subtitle')` or `word_format_text(paragraph_index, style='Subtitle', alignment='center', size_pt=14)`. Never use Title or Heading 1 for a subtitle!\n"
    "   - PRESERVE PRIOR WORK: Use the conversation history and document outline to understand what was created in previous turns. Never delete or distort previously built content unless specifically asked to edit it.\n"
)


class AgentSession:
    def __init__(self):
        self.active_id: Optional[str] = None
        self.is_running: bool = False
        self.event_queue: asyncio.Queue = asyncio.Queue()
        self.pending_calls: Dict[str, asyncio.Future] = {}
        self.current_process: Optional[subprocess.Popen] = None
        self._loop: Optional[asyncio.AbstractEventLoop] = None
        self.conversation_history: list = []
        self._current_turn_tools: list = []

    def reset_history(self):
        """Clears multi-turn conversation memory."""
        self.conversation_history.clear()
        self._current_turn_tools.clear()

    def is_active(self) -> bool:
        return self.is_running

    async def call_tool(self, name: str, args: Dict[str, Any], timeout_sec: float = 60.0) -> Dict[str, Any]:
        """
        Called by the MCP server when the AI invokes a word-office tool.
        Dispatches the tool call to the taskpane via SSE and awaits the client's result.
        """
        if not self.is_running:
            return {
                "ok": False,
                "error": "Word is not currently connected to an active agent session. Open the Gemini Antigravity taskpane in Word."
            }

        loop = asyncio.get_running_loop()
        call_id = str(uuid.uuid4())
        future = loop.create_future()
        self.pending_calls[call_id] = future
        self._current_turn_tools.append(f"{name}({json.dumps(args)})")

        # Emit tool_call event to taskpane
        await self.event_queue.put({
            "event": "tool_call",
            "call_id": call_id,
            "tool": name,
            "arguments": args,
            "destructive": is_destructive(name)
        })

        try:
            result = await asyncio.wait_for(future, timeout=timeout_sec)
            return result
        except asyncio.TimeoutError:
            return {
                "ok": False,
                "error": f"Timed out waiting for Microsoft Word to complete action '{name}' ({timeout_sec}s)."
            }
        except asyncio.CancelledError:
            return {"ok": False, "error": "Agent run was cancelled by user."}
        finally:
            self.pending_calls.pop(call_id, None)

    def resolve_tool(self, call_id: str, result: Dict[str, Any]) -> bool:
        """
        Called by POST /api/agent/tool_result when the taskpane finishes executing an action in Word.
        """
        fut = self.pending_calls.get(call_id)
        if fut and not fut.done():
            fut.set_result(result)
            return True
        return False

    async def cancel(self):
        """Cancels any running agent process and aborts pending tool calls."""
        self.is_running = False
        if self.current_process and self.current_process.poll() is None:
            try:
                self.current_process.kill()
            except Exception:
                pass
            self.current_process = None

        for fut in list(self.pending_calls.values()):
            if not fut.done():
                fut.cancel()
        self.pending_calls.clear()

        await self.event_queue.put({"event": "cancelled"})

    async def run(self, prompt: str, host: str = "Word", model: str = "gemini-3.8-flash-high",
                  document_map: str = "", analysis: Optional[Dict[str, Any]] = None) -> AsyncGenerator[str, None]:
        """
        Runs the autonomous agent session. Streams SSE lines to the client.
        """
        if self.is_running:
            yield f"data: {json.dumps({'event': 'error', 'error': 'Another agent task is already running. Please wait or cancel it.'})}\n\n"
            yield "data: [DONE]\n\n"
            return

        if analysis:
            yield f"data: {json.dumps({'event': 'analysis', **analysis})}\n\n"

        self.is_running = True
        self.active_id = str(uuid.uuid4())
        self._loop = asyncio.get_running_loop()
        self.event_queue = asyncio.Queue()
        self.pending_calls.clear()
        self._current_turn_tools.clear()

        agy_path = antigravity_backend.find_agy()
        if not agy_path:
            self.is_running = False
            err_msg = "Antigravity CLI (`agy`) not found. Please verify Antigravity is installed."
            yield f"data: {json.dumps({'event': 'error', 'error': err_msg})}\n\n"
            yield "data: [DONE]\n\n"
            return

        # Ensure MCP server word-office is registered with agy
        antigravity_backend.ensure_mcp_registered()

        history_context = ""
        if self.conversation_history:
            history_context = "PREVIOUS DIALOGUE HISTORY & EXECUTED ACTIONS IN THIS WORD DOCUMENT SESSION:\n"
            for turn in self.conversation_history[-6:]:
                history_context += f"[User]: {turn.get('content', '')}\n"
                if turn.get("tools"):
                    history_context += f"[Actions Executed]: {', '.join(turn['tools'])}\n"
                if turn.get("summary"):
                    history_context += f"[Assistant Summary]: {turn['summary']}\n"
            history_context += "\n"

        full_prompt = (
            f"{AGENT_SYSTEM_RULES}\n\n"
            f"{history_context}"
            f"{document_map}\n\n"
            f"Current User Instruction:\n{prompt}"
        )

        cmd = [
            agy_path,
            "--output-format", "stream-json",
            "--model", antigravity_backend.MODEL_MAP.get(model, antigravity_backend.DEFAULT_MODEL),
            "--disable-slash-commands",
            "--sandbox",
            "--dangerously-skip-permissions",
            "--print-timeout", "240s",
            "-p", full_prompt
        ]

        # Start child process in worker thread
        os.makedirs(antigravity_backend.WORKSPACE_DIR, exist_ok=True)
        try:
            self.current_process = subprocess.Popen(
                cmd,
                cwd=antigravity_backend.WORKSPACE_DIR,
                stdin=subprocess.DEVNULL,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                encoding="utf-8",
                errors="replace",
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)
            )
        except OSError as e:
            self.is_running = False
            yield f"data: {json.dumps({'event': 'error', 'error': f'Failed to launch Antigravity agent process: {e}'})}\n\n"
            yield "data: [DONE]\n\n"
            return

        # Background reader task reading stdout lines from process and queuing text deltas/status
        transcript_task: Optional[asyncio.Task] = None

        async def transcript_watcher(conv_id: str):
            brain_dir = os.path.expandvars(r"%USERPROFILE%\.gemini\antigravity-cli\brain")
            transcript_path = os.path.join(brain_dir, conv_id, ".system_generated", "logs", "transcript.jsonl")
            processed_lines = 0
            while self.is_running:
                await asyncio.sleep(0.25)
                if not os.path.isfile(transcript_path):
                    continue
                try:
                    with open(transcript_path, "r", encoding="utf-8", errors="replace") as f:
                        lines = f.readlines()
                    if len(lines) > processed_lines:
                        for line in lines[processed_lines:]:
                            line = line.strip()
                            if not line:
                                continue
                            try:
                                obj = json.loads(line)
                                thought = obj.get("thinking")
                                if thought and isinstance(thought, str) and thought.strip():
                                    await self.event_queue.put({
                                        "event": "thinking",
                                        "thought": thought.strip()
                                    })
                            except Exception:
                                pass
                        processed_lines = len(lines)
                except Exception:
                    pass

        async def stdout_reader():
            nonlocal transcript_task
            loop = asyncio.get_running_loop()
            proc = self.current_process
            if not proc or not proc.stdout:
                return

            def read_line():
                return proc.stdout.readline()

            while self.is_running and proc.poll() is None:
                line = await loop.run_in_executor(None, read_line)
                if not line:
                    break
                line_str = line.strip()
                if not line_str.startswith("{"):
                    continue
                try:
                    data = json.loads(line_str)
                except Exception:
                    continue

                event_type = data.get("event")
                if event_type == "init":
                    conv_id = data.get("conversation_id")
                    if conv_id and not transcript_task:
                        transcript_task = asyncio.create_task(transcript_watcher(conv_id))

                elif event_type == "step_update":
                    step = data.get("step_update", {})
                    # Detect tool invocation in progress
                    if step.get("step_type") == "tool" and step.get("state") == "ACTIVE":
                        tool_info = step.get("tool_info", {})
                        tool_name = tool_info.get("parameters", {}).get("ToolName") or step.get("tool_name", "")
                        await self.event_queue.put({
                            "event": "step_status",
                            "status": "active",
                            "tool": tool_name
                        })
                    # Stream text deltas from the final agent explanation
                    elif step.get("step_type") == "agent_response":
                        delta = step.get("text_delta")
                        if delta:
                            await self.event_queue.put({
                                "event": "text",
                                "delta": delta
                            })
                elif event_type == "result":
                    res = data.get("result", {})
                    if res.get("status") == "SUCCESS":
                        await self.event_queue.put({"event": "agent_finished"})
                    else:
                        err = res.get("error") or res.get("message") or "Agent execution failed"
                        await self.event_queue.put({"event": "error", "error": err})

            if self.is_running:
                await self.event_queue.put({"event": "agent_finished"})

        reader_task = asyncio.create_task(stdout_reader())

        response_deltas = []
        try:
            # Yield events from event_queue until agent_finished or error
            while self.is_running:
                item = await self.event_queue.get()
                ev = item.get("event")
                if ev == "text" and item.get("delta"):
                    response_deltas.append(item["delta"])
                yield f"data: {json.dumps(item)}\n\n"
                if ev in ("agent_finished", "error", "cancelled"):
                    break
        finally:
            turn_record = {
                "role": "user",
                "content": prompt,
                "tools": list(self._current_turn_tools),
                "summary": "".join(response_deltas).strip() if response_deltas else ""
            }
            self.conversation_history.append(turn_record)
            self.is_running = False
            reader_task.cancel()
            if transcript_task:
                transcript_task.cancel()
            if self.current_process and self.current_process.poll() is None:
                try:
                    self.current_process.kill()
                except Exception:
                    pass
                self.current_process = None
            yield "data: [DONE]\n\n"


agent_session = AgentSession()
