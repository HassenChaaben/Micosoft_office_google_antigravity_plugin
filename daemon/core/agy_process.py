"""Warm, multi-turn Antigravity CLI (`agy`) process speaking stream-json over stdin/stdout.

One process per chat:
    agy --input-format stream-json --output-format stream-json --model <id> [--effort e]
        [--conversation <id>] [extra args] -p ""
Each user turn is one stdin line: {"event":"user","message":{"content":"<text>"}}.
Output is NDJSON (init, step_update, result, error). Errors may also arrive on stderr as
`AGY_ERROR: {json}`. A turn that ends without a `result` is ALWAYS an error, never success.
"""

from __future__ import annotations

import asyncio
import json
import os
import re
import shutil
import subprocess
import sys
import threading
import time
from collections import deque
from dataclasses import dataclass, field
from typing import AsyncIterator, Dict, List, Mapping, Optional, Sequence, Union

# ------------------------------------------------------------------ events


@dataclass(frozen=True)
class Init:
    conversation_id: str
    model: str


@dataclass(frozen=True)
class TextDelta:
    text: str


@dataclass(frozen=True)
class Thinking:
    text: str


@dataclass(frozen=True)
class ToolStatus:
    tool: str
    state: str


@dataclass(frozen=True)
class TurnDone:
    response: str
    usage: Dict = field(default_factory=dict)


@dataclass(frozen=True)
class TurnError:
    message: str
    # "quota" | "auth" | "model" | "other" ; AgyProcess also emits "cancelled"
    code: str = "other"
    retry_after: Optional[str] = None


AgyEvent = Union[Init, TextDelta, Thinking, ToolStatus, TurnDone, TurnError]

_RETRY_RE = re.compile(r"resets? in\s+([0-9]+[0-9hmsd]*)", re.IGNORECASE)
_THINKING_STEP_TYPES = {"thinking", "agent_thinking", "planner_thinking", "reasoning"}


def classify_error(text: str) -> tuple:
    """Return (code, retry_after) for an agy error text."""
    low = (text or "").lower()
    match = _RETRY_RE.search(text or "")
    retry_after = match.group(1) if match else None
    if "resource_exhausted" in low or "quota" in low or "code 429" in low or "rate limit" in low:
        return "quota", retry_after
    if ("not logged" in low or "unauthenticated" in low or "log in" in low or "sign in" in low
            or "/login" in low or "code 401" in low):
        return "auth", retry_after
    if "model" in low and any(k in low for k in ("not found", "unknown", "invalid", "unsupported",
                                                   "not supported", "not available")):
        return "model", retry_after
    return "other", retry_after


def _error(message: str, hint: str = "") -> TurnError:
    code, retry = classify_error(f"{message} {hint}")
    return TurnError(message=message, code=code, retry_after=retry)


def _error_text(value) -> str:
    if isinstance(value, dict):
        return str(value.get("message") or value.get("short_error") or value.get("error") or json.dumps(value))
    return str(value) if value is not None else ""


def parse_event(line: Union[str, bytes]) -> Optional[AgyEvent]:
    """Parse one stdout NDJSON line or one stderr `AGY_ERROR:` line. Returns None for anything else."""
    if isinstance(line, bytes):
        line = line.decode("utf-8", errors="replace")
    s = (line or "").strip()
    if not s:
        return None
    if s.startswith("AGY_ERROR:"):
        payload = s[len("AGY_ERROR:"):].strip()
        try:
            data = json.loads(payload)
        except json.JSONDecodeError:
            return _error(payload)
        if not isinstance(data, dict):
            return _error(payload)
        msg = _error_text(data.get("short_error") or data.get("message") or data.get("error") or payload)
        return _error(msg, f"{data.get('status', '')} code {data.get('error_code', '')}")
    if not s.startswith("{"):
        return None
    try:
        obj = json.loads(s)
    except json.JSONDecodeError:
        return None
    if not isinstance(obj, dict):
        return None

    kind = obj.get("event")
    if kind == "init":
        init = obj.get("init") or {}
        return Init(conversation_id=obj.get("conversation_id") or init.get("conversation_id") or "",
                    model=init.get("model") or "")
    if kind == "step_update":
        st = obj.get("step_update") or {}
        stype = st.get("step_type") or ""
        delta = st.get("text_delta")
        if st.get("thinking_delta"):
            return Thinking(text=st["thinking_delta"])
        if stype in _THINKING_STEP_TYPES:
            return Thinking(text=delta) if delta else None
        if stype == "agent_response":
            return TextDelta(text=delta) if delta else None
        tool_info = st.get("tool_info") or {}
        if stype == "tool" or tool_info or st.get("tool_name"):
            params = tool_info.get("parameters") or {}
            name = (params.get("ToolName") or st.get("tool_name") or tool_info.get("name")
                    or tool_info.get("tool_name") or "tool")
            return ToolStatus(tool=str(name), state=str(st.get("state") or ""))
        return None
    if kind == "result":
        res = obj.get("result") or {}
        if res.get("status") == "SUCCESS":
            return TurnDone(response=res.get("response") or "", usage=res.get("usage") or {})
        msg = _error_text(res.get("error")) or f"agy returned status {res.get('status') or 'unknown'}"
        return _error(msg)
    if kind == "error":
        msg = _error_text(obj.get("error") or obj.get("message")) or s
        return _error(msg)
    return None


# ------------------------------------------------------------------ discovery & models


def find_agy() -> Optional[str]:
    """agy on PATH, else %LOCALAPPDATA%\\agy\\bin\\agy.exe, else None."""
    found = shutil.which("agy")
    if found:
        return found
    local = os.environ.get("LOCALAPPDATA")
    if local:
        candidate = os.path.join(local, "agy", "bin", "agy.exe")
        if os.path.isfile(candidate):
            return candidate
    return None


_EFFORTS = ("low", "medium", "high", "xhigh", "max")


def parse_models(text: str) -> List[Dict[str, Optional[str]]]:
    """Parse `agy models` output: lines '<id>\\t<label>'."""
    models = []
    for raw in (text or "").splitlines():
        if "\t" not in raw:
            continue
        mid, label = raw.split("\t", 1)
        mid, label = mid.strip(), label.strip()
        if not mid or " " in mid:
            continue
        family, effort = mid, None
        head, _, tail = mid.rpartition("-")
        if head and tail in _EFFORTS:
            family, effort = head, tail
        models.append({"id": mid, "label": label or mid, "family": family, "effort": effort})
    return models


class AgyModelsError(RuntimeError):
    def __init__(self, message: str, code: str = "other"):
        super().__init__(message)
        self.code = code


def _as_command(agy: Union[str, Sequence[str]]) -> List[str]:
    return [agy] if isinstance(agy, str) else list(agy)


_CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def build_env(extra: Optional[Mapping[str, str]] = None) -> Dict[str, str]:
    """Inherited environment + extra, guaranteeing the profile variables agy needs for its login."""
    env = dict(os.environ)
    if extra:
        env.update({k: str(v) for k, v in extra.items()})
    home = env.get("USERPROFILE") or env.get("HOME") or os.path.expanduser("~")
    env.setdefault("USERPROFILE", home)
    if not env.get("HOME"):
        env["HOME"] = home
    if not env.get("APPDATA"):
        env["APPDATA"] = os.path.join(home, "AppData", "Roaming")
    if not env.get("LOCALAPPDATA"):
        env["LOCALAPPDATA"] = os.path.join(home, "AppData", "Local")
    return env


def list_models(agy: Union[str, Sequence[str]], timeout: float = 30.0,
                env: Optional[Mapping[str, str]] = None) -> List[Dict[str, Optional[str]]]:
    """Run `agy models` (agy = path or command prefix list). Raises AgyModelsError on failure."""
    try:
        proc = subprocess.run(_as_command(agy) + ["models"], stdin=subprocess.DEVNULL, capture_output=True,
                              encoding="utf-8", errors="replace", timeout=timeout,
                              creationflags=_CREATE_NO_WINDOW, env=build_env(env))
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise AgyModelsError(f"Could not run agy models: {exc}") from exc
    models = parse_models(proc.stdout)
    if proc.returncode != 0 or not models:
        detail = (proc.stderr or proc.stdout or "").strip()[-500:] or f"exit code {proc.returncode}"
        code, _ = classify_error(detail)
        raise AgyModelsError(f"agy models failed: {detail}", code)
    return models


class ModelCache:
    """Thread-safe TTL cache around list_models."""

    def __init__(self, ttl: float = 300.0):
        self.ttl = ttl
        self._lock = threading.Lock()
        self._data: Dict[tuple, tuple] = {}

    def get(self, agy: Union[str, Sequence[str]], force: bool = False) -> List[Dict[str, Optional[str]]]:
        key = tuple(_as_command(agy))
        with self._lock:
            hit = self._data.get(key)
            if hit and not force and time.monotonic() - hit[0] < self.ttl:
                return hit[1]
        models = list_models(agy)
        with self._lock:
            self._data[key] = (time.monotonic(), models)
        return models

    def invalidate(self) -> None:
        with self._lock:
            self._data.clear()


# ------------------------------------------------------------------ warm process


def default_workspace() -> str:
    base = os.environ.get("LOCALAPPDATA") or os.path.join(os.path.expanduser("~"), "AppData", "Local")
    return os.path.join(base, "GeminiAntigravityOffice", "agy_workspace")


_STREAM_LIMIT = 32 * 1024 * 1024  # a whole long answer can arrive in one `result` line


class AgyProcess:
    """A warm agy process for one chat. One turn at a time; auto-restarts (resuming) when it died."""

    def __init__(self, agy_path: Optional[str], model: str, effort: Optional[str] = None,
                 conversation_id: Optional[str] = None, cwd: Optional[str] = None,
                 extra_args: Sequence[str] = (), env: Optional[Mapping[str, str]] = None,
                 command_prefix: Optional[Sequence[str]] = None):
        if command_prefix is None and not agy_path:
            agy_path = find_agy() or "agy"
        self.command_prefix: List[str] = list(command_prefix) if command_prefix else [agy_path]
        self.model = model
        self.effort = effort
        self.cwd = cwd or default_workspace()
        self.extra_args = list(extra_args)
        self.env = env  # kept by reference: read again on every (re)start
        self._conversation_id = conversation_id
        self._proc: Optional[asyncio.subprocess.Process] = None
        self._stderr_task: Optional[asyncio.Task] = None
        self._stderr_lines: deque = deque(maxlen=200)
        self._stdout_noise: deque = deque(maxlen=20)
        self._lock = asyncio.Lock()
        self._cancelled = False

    # -- properties / helpers

    @property
    def conversation_id(self) -> Optional[str]:
        return self._conversation_id

    @property
    def running(self) -> bool:
        return self._proc is not None and self._proc.returncode is None

    def build_args(self) -> List[str]:
        args = list(self.command_prefix) + ["--input-format", "stream-json", "--output-format", "stream-json",
                                            "--model", self.model]
        if self.effort:
            args += ["--effort", self.effort]
        if self._conversation_id:
            args += ["--conversation", self._conversation_id]
        args += self.extra_args
        args += ["-p", ""]
        return args

    def build_env(self) -> Dict[str, str]:
        return build_env(self.env)

    # -- lifecycle

    async def start(self) -> None:
        async with self._lock:
            if not self.running:
                await self._spawn()

    async def _spawn(self) -> None:
        await self._reap()
        os.makedirs(self.cwd, exist_ok=True)
        self._stderr_lines.clear()
        self._stdout_noise.clear()
        self._cancelled = False
        self._proc = await asyncio.create_subprocess_exec(
            *self.build_args(), stdin=asyncio.subprocess.PIPE, stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE, cwd=self.cwd, env=self.build_env(),
            creationflags=_CREATE_NO_WINDOW, limit=_STREAM_LIMIT)
        self._stderr_task = asyncio.ensure_future(self._read_stderr(self._proc))

    async def _read_stderr(self, proc) -> None:
        try:
            while True:
                raw = await proc.stderr.readline()
                if not raw:
                    return
                self._stderr_lines.append(raw.decode("utf-8", errors="replace").rstrip("\r\n"))
        except (asyncio.CancelledError, Exception):
            return

    async def _kill(self) -> None:
        proc = self._proc
        if proc is None or proc.returncode is not None:
            return
        if sys.platform == "win32":
            # agy may spawn helpers that inherit our pipes; kill the whole tree.
            try:
                tk = await asyncio.create_subprocess_exec(
                    "taskkill", "/PID", str(proc.pid), "/T", "/F", stdout=asyncio.subprocess.DEVNULL,
                    stderr=asyncio.subprocess.DEVNULL, creationflags=_CREATE_NO_WINDOW)
                await asyncio.wait_for(tk.wait(), 10)
            except Exception:
                pass
        try:
            proc.kill()
        except ProcessLookupError:
            pass
        try:
            await asyncio.wait_for(proc.wait(), 10)
        except asyncio.TimeoutError:
            pass

    async def _reap(self) -> None:
        """Forget a previous (dead or abandoned) process."""
        if self._proc is not None:
            await self._kill()
        if self._stderr_task is not None:
            try:
                await asyncio.wait_for(self._stderr_task, 1)
            except (asyncio.TimeoutError, asyncio.CancelledError, Exception):
                self._stderr_task.cancel()
        self._proc = None
        self._stderr_task = None

    async def cancel(self) -> None:
        """Stop the running turn now. The next send() restarts, resuming the known conversation."""
        self._cancelled = True
        await self._kill()

    async def close(self) -> None:
        async with self._lock:
            proc = self._proc
            if proc is not None and proc.returncode is None:
                try:
                    proc.stdin.close()
                    await asyncio.wait_for(proc.wait(), 3)
                except Exception:
                    pass
            await self._reap()

    # -- turns

    async def _stderr_error(self, wait: float) -> Optional[TurnError]:
        """Most relevant AGY_ERROR from stderr; optionally wait for the stderr reader to finish."""
        if wait and self._stderr_task is not None:
            try:
                await asyncio.wait_for(asyncio.shield(self._stderr_task), wait)
            except (asyncio.TimeoutError, asyncio.CancelledError, Exception):
                pass
        for line in reversed(self._stderr_lines):
            ev = parse_event(line)
            if isinstance(ev, TurnError):
                return ev
        return None

    async def _eof_error(self) -> TurnError:
        proc = self._proc
        code = None
        if proc is not None:
            try:
                code = await asyncio.wait_for(proc.wait(), 5)
            except asyncio.TimeoutError:
                await self._kill()
                code = proc.returncode
        err = await self._stderr_error(wait=2)
        if err is not None:
            return err
        tail = " | ".join([l for l in self._stderr_lines if l.strip()][-5:] + list(self._stdout_noise)[-3:])
        msg = f"agy exited (code {code}) without a result" + (f": {tail}" if tail else "")
        return _error(msg[:2000])

    async def send(self, text: str) -> AsyncIterator[AgyEvent]:
        """Send one user turn; yield events until TurnDone or TurnError (always exactly one of them last)."""
        async with self._lock:
            finished = False
            try:
                retried = False
                while True:
                    fresh = False
                    if not self.running:
                        try:
                            await self._spawn()
                        except OSError as exc:
                            finished = True
                            yield TurnError(message=f"Could not start agy ({self.command_prefix[0]}): {exc}",
                                            code="other")
                            return
                        fresh = True
                    self._cancelled = False
                    self._stderr_lines.clear()
                    proc = self._proc
                    line = json.dumps({"event": "user", "message": {"content": text}}, ensure_ascii=False) + "\n"
                    write_failed = False
                    try:
                        proc.stdin.write(line.encode("utf-8"))
                        await proc.stdin.drain()
                    except (BrokenPipeError, ConnectionResetError, OSError, RuntimeError):
                        write_failed = True

                    got_any = False
                    if not write_failed:
                        while True:
                            raw = await proc.stdout.readline()
                            if not raw:
                                break
                            ev = parse_event(raw)
                            if ev is None:
                                noise = raw.decode("utf-8", errors="replace").strip()
                                if noise:
                                    self._stdout_noise.append(noise)
                                continue
                            got_any = True
                            if isinstance(ev, Init):
                                if ev.conversation_id:
                                    self._conversation_id = ev.conversation_id
                                yield ev
                                continue
                            if isinstance(ev, TurnError) and ev.code == "other":
                                better = await self._stderr_error(wait=0.3)
                                if better is not None and better.code != "other":
                                    ev = better
                            if isinstance(ev, (TurnDone, TurnError)):
                                finished = True
                                yield ev
                                return
                            yield ev

                    # EOF (or broken stdin) before a result.
                    if self._cancelled:
                        finished = True
                        await self._reap()
                        yield TurnError(message="Cancelled", code="cancelled")
                        return
                    if not fresh and not got_any and not retried:
                        # The warm process died between turns without us noticing: restart and resume once.
                        retried = True
                        await self._reap()
                        continue
                    err = await self._eof_error()
                    await self._reap()
                    finished = True
                    yield err
                    return
            finally:
                if not finished:
                    # Consumer abandoned the turn mid-stream: the process still owes output, so drop it.
                    await self._reap()
