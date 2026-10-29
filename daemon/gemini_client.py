"""
Gemini Streaming AI Engine
Streams reasoning, analysis, and document transformations using Gemini 3.8 Flash High, Pro, and Deep Think.
Supports live Google Gemini API streaming (via API Key or Subscription Bridge) and intelligent local domain synthesis.
"""

import os
import re
import time
import json
import asyncio
from typing import AsyncGenerator, Dict, Any
from . import antigravity_backend
from .auth import auth_bridge
from .skills_library import get_skill_definition

class GeminiClient:
    def __init__(self):
        pass

    async def stream_chat(self, payload: Dict[str, Any]) -> AsyncGenerator[str, None]:
        prompt = payload.get("prompt", "")
        model = payload.get("model", "gemini-3.8-flash-high")
        host = payload.get("host", "Word")
        web_search = payload.get("web_search", True)
        doc_context = payload.get("document_context", "")

        backend = os.environ.get("GEMINI_OFFICE_BACKEND", "auto").lower()
        errors = []

        # 0. Antigravity subscription via the signed-in `agy` CLI
        agy_path = antigravity_backend.find_agy() if backend != "offline" else None
        if agy_path:
            events = antigravity_backend.stream_events(agy_path, prompt, host, doc_context, model)
            has_streamed = False
            while True:
                event = await asyncio.to_thread(next, events, None)
                if event is None:
                    break
                kind, text = event
                if kind == "delta":
                    has_streamed = True
                    yield f"data: {json.dumps({'delta': text})}\n\n"
                elif kind == "error":
                    errors.append(text)
                    if has_streamed:
                        yield f"data: {json.dumps({'delta': chr(10) * 2 + '> ⚠️ ' + text})}\n\n"
                    break
                else:
                    break
            if has_streamed:
                yield "data: [DONE]\n\n"
                return
        elif backend != "offline":
            errors.append("Antigravity CLI (`agy`) was not found on this PC. Install/sign in to Antigravity CLI.")

        api_key = auth_bridge.get_api_key() if backend != "offline" else None

        # 1. Live Google Gemini API Stream (if API Key configured)
        if api_key:
            try:
                import google.generativeai as genai
                genai.configure(api_key=api_key)

                # Map UI model selector to Google Gemini model name
                model_map = {
                    "gemini-3.8-flash-high": "gemini-flash-latest",
                    "gemini-pro": "gemini-pro-latest",
                    "gemini-deep-think": "gemini-pro-latest"
                }
                model_id = model_map.get(model, "gemini-flash-latest")

                system_instruction = (
                    f"You are Gemini Antigravity, an autonomous executive AI operating inside Microsoft {host}. "
                    "Generate comprehensive, structured, publication-grade deliverables with headings, tables, bullet points, and actionable details. "
                    "Do not include conversational filler like 'Sure, here is...' or 'I hope this helps'."
                )

                full_prompt = prompt
                if doc_context:
                    full_prompt = (
                        f"Active Microsoft {host} Document Context:\n"
                        f"\"\"\"\n{doc_context}\n\"\"\"\n\n"
                        f"User Request:\n{prompt}"
                    )

                gen_model = genai.GenerativeModel(model_id, system_instruction=system_instruction)
                response = gen_model.generate_content(full_prompt, stream=True)

                has_streamed = False
                for chunk in response:
                    if hasattr(chunk, "text") and chunk.text:
                        has_streamed = True
                        sse_data = json.dumps({"delta": chunk.text})
                        yield f"data: {sse_data}\n\n"

                if has_streamed:
                    yield "data: [DONE]\n\n"
                    return
            except Exception as e:
                print(f"[GeminiClient] Live Google API call failed: {e}")
                errors.append(f"Gemini API key call failed: {e}")

        if backend != "offline":
            message = (
                "### ⚠️ No AI answer\n\n"
                "The request could not reach a real AI model, so nothing was generated.\n\n"
                + "\n".join(f"- {err}" for err in errors)
                + "\n\nCheck that Antigravity CLI is signed in (run `agy models` in a terminal), then try again."
            )
            yield f"data: {json.dumps({'delta': message})}\n\n"
            yield "data: [DONE]\n\n"
            return

        # 2. Offline mock mode (strictly for automated mock tests when GEMINI_OFFICE_BACKEND=offline)
        response_text = self._generate_mock_response(prompt, host, doc_context, web_search, model)

        # Stream words as SSE deltas to simulate realistic real-time token streaming
        words = response_text.split(" ")
        is_test = bool(os.environ.get("PYTEST_CURRENT_TEST"))
        for i, word in enumerate(words):
            chunk = word + (" " if i < len(words) - 1 else "")
            sse_data = json.dumps({"delta": chunk})
            yield f"data: {sse_data}\n\n"
            if not is_test:
                await asyncio.sleep(0.012)

        yield "data: [DONE]\n\n"

    def _generate_mock_response(self, prompt: str, host: str, doc_context: str = "", web_search: bool = False, model: str = "") -> str:
        """
        Mock response strictly for automated test suites when GEMINI_OFFICE_BACKEND=offline.
        Never used in real runtime.
        """
        clean_p = prompt.strip()
        return f"[Mock offline response for {host}: {clean_p}]"

    def _generate_intelligent_response(self, prompt: str, host: str, doc_context: str = "", web_search: bool = False, model: str = "") -> str:
        return self._generate_mock_response(prompt, host, doc_context, web_search, model)

gemini_client = GeminiClient()
