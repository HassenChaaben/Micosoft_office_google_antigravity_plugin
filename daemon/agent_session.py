"""
Autonomous Office Agent Session Orchestrator
Coordinates communication between the local Antigravity AI CLI (`agy`),
the internal Model Context Protocol (MCP) server, and the Microsoft Word Taskpane.
"""

import os
import json
import uuid
import time
import asyncio
from typing import Dict, Any, Optional, AsyncGenerator

from . import antigravity_backend
from .word_tools import is_destructive
from .core.agy_process import AgyProcess, Init, TextDelta, Thinking, ToolStatus, TurnDone, TurnError, find_agy
from .core.skills import load_skills, parse_forced, render_for_prompt, route
from .core.store import Store

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
    "4. WRITING CONTENT: for any new content longer than one paragraph (report sections, headings with text, lists, tables with merged cells/shading, captions, TOC) call `word_insert_content` ONCE with all the blocks: it is one atomic native insert with built-in styles and it is much faster than many small calls. Use exact requested columns and row counts. Keep the returned block_id to edit that content later (location 'replace:ga:<id>').\n"
    "5. INSERTING PICTURES & VISUALS: When asked to add a picture, logo, diagram, or chart, call `word_insert_image` with the URL, keyword ('logo', 'chart', 'diagram', 'banner'), or base64 data, and an optional caption.\n"
    "6. TABLE & DOCUMENT MANAGEMENT: Use `word_read_table(table_index)` to read table cells and `word_delete_table(table_index)` to remove a table if needed.\n"
    "7. BUILT-IN STYLES: Use standard built-in style names: 'Heading 1', 'Heading 2', 'Title', 'Normal'. Localization (French Titre 1, Italian Titolo 1) is handled automatically.\n"
    "7b. PLAN AND REMEMBER: for any multi-step task call `agent_plan` first (short steps). When the user states a preference or a project fact that should last (format rules, names, codes, budget, dates), call `memory_save`; facts already remembered are listed in your prompt. Use `skill_load` to read an expert skill that is listed but not shown in full.\n"
    "8. VERIFY, THEN SUMMARIZE: after your edits ALWAYS call `office_verify` (pass expect_all_fields_filled=true when the user asked to fill/complete). Fix every issue it reports and verify again. Then give a short summary of what changed and anything left '[da definire]'.\n"
    "9. NO LOCAL OS/FILE EDITS: Never call file, shell, terminal, or browser tools. Only use `word-office` MCP tools.\n"
    "10. TEMPLATES, FORMS & PM DOCUMENTS: fill fields with ONE `word_fill_fields` call per table (or for the whole document), using the field paths printed in the DOCUMENT MAP (e.g. 'Cliente › Azienda'), grid keys '<grid> › #<record> › <column>' (e.g. 'Minacce › #2 › Causa') or 'T1r2c3'. It writes in place and keeps the template formatting. To change one value later, call it again with that key. If a grid needs more rows than it has, clone the formatted row first (word_table_clone_row). Unknown values: write '[da definire]', never invent.\n"
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
        self._loop: Optional[asyncio.AbstractEventLoop] = None
        self.turn_snapshot: Optional[Dict[str, Any]] = None
        self._current_turn_tools: list = []
        self._current_chat: Optional[str] = None
        self._current_doc_url: Optional[str] = None
        self._store: Optional[Store] = None
        self._factory = None
        self._skills = None
        self._procs: Dict[str, Any] = {}
        self._sent_skills: Dict[str, set] = {}
        self._com_worker: Optional[Any] = None
        self._perception: Optional[Any] = None

    @property
    def com_worker(self):
        if self._com_worker is None:
            from .office.com_worker import ComWorker
            self._com_worker = ComWorker()
        return self._com_worker

    @com_worker.setter
    def com_worker(self, worker):
        self._com_worker = worker

    @property
    def perception(self):
        if self._perception is None:
            from .office.perception import Perception
            self._perception = Perception()
        return self._perception

    @perception.setter
    def perception(self, p):
        self._perception = p

    def invalidate_perception_cache(self) -> None:
        """Clear perception cache so verification and future turns see fresh document state."""
        if self._perception is not None:
            self._perception._cache.clear()
        import sys
        main_mod = sys.modules.get("daemon.main")
        if main_mod and hasattr(main_mod, "perception"):
            try:
                main_mod.perception._cache.clear()
            except Exception:
                pass

    async def refresh_perception(self, doc_url: Optional[str] = None) -> Optional[Dict[str, Any]]:
        """Invalidates cache and attempts to refresh the perception snapshot via COM worker."""
        self.invalidate_perception_cache()
        url = doc_url or getattr(self, "_current_doc_url", None)
        if self._com_worker is not None:
            try:
                def _get_ooxml():
                    doc = self.com_worker.get_document(url)
                    return doc.Content.WordOpenXML, doc.Name, doc.FullName

                ooxml, name, full_name = await self.com_worker.run(_get_ooxml)
                if ooxml:
                    payload = {
                        "document_ooxml": ooxml,
                        "doc_title": name,
                        "doc_url": full_name,
                    }
                    res = await self.perception.analyze(payload)
                    return res.snapshot
            except Exception:
                pass
        return None

    async def verify_document(self, args: Dict[str, Any]) -> Dict[str, Any]:
        """office_verify: capture the document again through the taskpane and compare with the turn start."""
        from .office.verify import verify

        self.invalidate_perception_cache()
        captured = await self.call_tool("__capture_document", {})
        if not captured.get("ok") or not captured.get("document_ooxml"):
            if self._com_worker is not None:
                try:
                    def _get_doc_ooxml():
                        doc = self.com_worker.get_document(getattr(self, "_current_doc_url", None))
                        return doc.Content.WordOpenXML, doc.Name, doc.FullName
                    ooxml, name, full_name = await self.com_worker.run(_get_doc_ooxml)
                    if ooxml:
                        captured = {"ok": True, "document_ooxml": ooxml, "doc_title": name, "doc_url": full_name}
                except Exception:
                    pass
        if not captured.get("ok") or not captured.get("document_ooxml"):
            return {"ok": False, "error": captured.get("error") or "The document could not be read for verification."}
        result = await self.perception.analyze(captured)
        return verify(self.turn_snapshot, result.snapshot,
                      expect_all_fields_filled=args.get("expect_all_fields_filled", True))

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

        started = time.monotonic()
        try:
            result = await asyncio.wait_for(future, timeout=timeout_sec)
            self._record_tool(name, args, result, started)
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

    def _record_tool(self, name: str, args: Dict[str, Any], result: Dict[str, Any], started: float):
        if not self._current_chat or name.startswith("__"):
            return  # document captures are huge and not useful in history
        try:
            slim_args = {k: (v if len(json.dumps(v, ensure_ascii=False)) < 4000 else "[large]") for k, v in args.items()}
            self.store.add_tool_call(self._current_chat, None, name, slim_args, result,
                                     ok=bool(result.get("ok")), duration_ms=int((time.monotonic() - started) * 1000))
        except Exception as e:
            print(f"[AgentSession] could not record tool call {name}: {e}")

    def resolve_tool(self, call_id: str, result: Dict[str, Any]) -> bool:
        """
        Called by POST /api/agent/tool_result when the taskpane finishes executing an action in Word.
        """
        fut = self.pending_calls.get(call_id)
        if fut and not fut.done():
            fut.set_result(result)
            return True
        return False

    # --- configuration ---------------------------------------------------------------

    def configure(self, store=None, process_factory=None, skills=None, com_worker=None, perception=None):
        """Inject the history store, the agy process factory and the skill list (used by tests)."""
        if store is not None:
            self._store = store
        if process_factory is not None:
            self._factory = process_factory
        if skills is not None:
            self._skills = skills
        if com_worker is not None:
            self._com_worker = com_worker
        if perception is not None:
            self._perception = perception
        self._procs = {}
        self._sent_skills = {}

    @property
    def store(self) -> Store:
        if self._store is None:
            self._store = Store()
        return self._store

    def skills(self):
        if self._skills is None:
            self._skills = load_skills()
        return self._skills

    def _new_process(self, model: str, effort: Optional[str], conversation_id: Optional[str]) -> AgyProcess:
        if self._factory is not None:
            return self._factory(model, effort=effort, conversation_id=conversation_id)
        return AgyProcess(find_agy(), model, effort=effort, conversation_id=conversation_id,
                          extra_args=("--sandbox", "--dangerously-skip-permissions", "--disable-slash-commands"))

    async def _process_for(self, chat: Dict[str, Any], model: str, effort: Optional[str]) -> AgyProcess:
        """One warm agy process per chat; restarted (resuming the conversation) when the model changes."""
        entry = self._procs.get(chat["id"])
        if entry and (entry[1], entry[2]) == (model, effort):
            return entry[0]
        if entry:
            await entry[0].close()
        proc = self._new_process(model, effort, chat.get("agy_conversation_id"))
        self._procs[chat["id"]] = (proc, model, effort)
        return proc

    async def shutdown(self):
        for proc, _, _ in list(self._procs.values()):
            await proc.close()
        self._procs.clear()
        if self._com_worker is not None:
            try:
                self._com_worker.stop()
            except Exception:
                pass
            self._com_worker = None

    # --- control -------------------------------------------------------------------------

    async def cancel(self):
        """Cancels the running turn (the agy conversation is resumed on the next message)."""
        self.is_running = False
        entry = self._procs.get(self._current_chat or "")
        if entry:
            await entry[0].cancel()
        for fut in list(self.pending_calls.values()):
            if not fut.done():
                fut.cancel()
        self.pending_calls.clear()
        await self.event_queue.put({"event": "cancelled"})

    def reset_history(self):
        """Kept for the legacy /api/agent/reset endpoint: chats now live in the history database."""
        self._current_turn_tools.clear()

    # --- one turn ------------------------------------------------------------------------

    def _compose_message(self, chat: Dict[str, Any], first_turn: bool, prompt: str, host: str,
                         doc_type: Optional[str], document_map: str, doc_url: Optional[str]):
        skills = self.skills()
        forced, cleaned = parse_forced(prompt, known={s.name for s in skills})
        routed = route(skills, host.lower(), doc_type, cleaned, forced=forced)
        sent = self._sent_skills.setdefault(chat["id"], set())
        new_full = [s for s in routed["full"] if s.name not in sent]
        parts = []
        if first_turn:
            parts.append(AGENT_SYSTEM_RULES)
        if new_full or first_turn:
            parts.append(render_for_prompt({"full": new_full, "listed": routed["listed"]}))
        scopes = ["global"] + ([f"doc:{doc_url}"] if doc_url else [])
        memories = self.store.search_memories(cleaned, scopes, limit=8)
        if memories:
            parts.append("REMEMBERED USER PREFERENCES AND FACTS (from memory):\n"
                         + "\n".join(f"- {m['text']}" for m in memories))
        if document_map:
            parts.append(document_map)
        parts.append(f"User request:\n{cleaned}")
        return "\n\n".join(parts), routed, new_full

    async def run(self, prompt: str, host: str = "Word", model: str = "gemini-3.8-flash-high",
                  document_map: str = "", analysis: Optional[Dict[str, Any]] = None,
                  snapshot: Optional[Dict[str, Any]] = None, chat_id: Optional[str] = None,
                  doc_url: Optional[str] = None, doc_title: Optional[str] = None,
                  effort: Optional[str] = None) -> AsyncGenerator[str, None]:
        """Runs one agent turn in the chat's warm agy process. Streams SSE lines to the taskpane."""
        def sse(obj):
            return f"data: {json.dumps(obj, ensure_ascii=False)}\n\n"

        if self.is_running:
            yield sse({"event": "error", "error": "Another agent task is already running. Please wait or cancel it."})
            yield "data: [DONE]\n\n"
            return

        model = antigravity_backend.MODEL_MAP.get(model, model)
        chat = self.store.get_chat(chat_id) if chat_id else None
        if chat is None:
            chat = self.store.create_chat(host=host, doc_url=doc_url, doc_title=doc_title, model=model, effort=effort)
        else:
            self.store.update_chat(chat["id"], model=model, effort=effort)
        first_turn = not chat.get("agy_conversation_id") and chat["id"] not in self._sent_skills
        self._current_chat = chat["id"]
        self._current_doc_url = doc_url or chat.get("doc_url")
        yield sse({"event": "chat", "chat_id": chat["id"], "title": chat.get("title")})

        self.turn_snapshot = snapshot
        if analysis:
            yield sse({"event": "analysis", **analysis})

        message, routed, new_full = self._compose_message(
            chat, first_turn, prompt, host, (analysis or {}).get("doc_type"), document_map, doc_url)
        yield sse({"event": "skills", "names": [s.name for s in routed["full"]]})
        self.store.add_message(chat["id"], "user", prompt)

        self.is_running = True
        self.active_id = str(uuid.uuid4())
        self.event_queue = asyncio.Queue()
        self.pending_calls.clear()
        self._current_turn_tools.clear()

        try:
            proc = await self._process_for(chat, model, effort)
        except Exception as e:  # agy not installed, bad path...
            self.is_running = False
            yield sse({"event": "error", "error": f"Could not start the Antigravity agent: {e}", "code": "other"})
            yield "data: [DONE]\n\n"
            return

        answer: list = []
        outcome = {"ok": False}

        async def pump():
            try:
                async for ev in proc.send(message):
                    if isinstance(ev, Init):
                        if ev.conversation_id and ev.conversation_id != chat.get("agy_conversation_id"):
                            self.store.set_conversation_id(chat["id"], ev.conversation_id)
                            chat["agy_conversation_id"] = ev.conversation_id
                    elif isinstance(ev, TextDelta):
                        answer.append(ev.text)
                        await self.event_queue.put({"event": "text", "delta": ev.text})
                    elif isinstance(ev, Thinking):
                        await self.event_queue.put({"event": "thinking", "thought": ev.text})
                    elif isinstance(ev, ToolStatus):
                        await self.event_queue.put({"event": "step_status", "status": ev.state, "tool": ev.tool})
                    elif isinstance(ev, TurnDone):
                        outcome["ok"] = True
                        await self.event_queue.put({"event": "agent_finished"})
                    elif isinstance(ev, TurnError):
                        if ev.code == "cancelled":
                            await self.event_queue.put({"event": "cancelled"})
                        else:
                            await self.event_queue.put({"event": "error", "error": ev.message, "code": ev.code,
                                                        "retry_after": ev.retry_after})
            except Exception as e:
                await self.event_queue.put({"event": "error", "error": f"Agent runtime error: {e}", "code": "other"})

        pump_task = asyncio.create_task(pump())
        try:
            while True:
                item = await self.event_queue.get()
                yield sse(item)
                if item.get("event") in ("agent_finished", "error", "cancelled"):
                    break
        finally:
            self.is_running = False
            if not pump_task.done():
                pump_task.cancel()
            text = "".join(answer).strip()
            if text:
                self.store.add_message(chat["id"], "assistant", text)
            if outcome["ok"]:
                self._sent_skills.setdefault(chat["id"], set()).update(s.name for s in new_full)
            yield "data: [DONE]\n\n"


agent_session = AgentSession()
