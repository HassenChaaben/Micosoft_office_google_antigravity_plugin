"""Tests for daemon.core.agy_process (warm multi-turn agy stream-json process)."""

import asyncio
import json
import os
import sys
from pathlib import Path

import pytest

from daemon.core.agy_process import (
    AgyProcess,
    Init,
    ModelCache,
    TextDelta,
    Thinking,
    ToolStatus,
    TurnDone,
    TurnError,
    find_agy,
    list_models,
    parse_event,
    parse_models,
)

FIXTURES = Path(__file__).parent / "fixtures" / "agy_stream"
FAKE_AGY = str(Path(__file__).parent / "fake_agy.py")
PREFIX = [sys.executable, "-u", FAKE_AGY]


def _parse_file(name):
    events = []
    for line in (FIXTURES / name).read_text(encoding="utf-8").splitlines():
        ev = parse_event(line)
        if ev is not None:
            events.append(ev)
    return events


# ---------------------------------------------------------------- parse_event


def test_parse_success_fixture():
    events = _parse_file("single_turn_success.ndjson")
    assert events[0] == Init(conversation_id="f048210c-942d-4c3d-91d8-4ee1d7188df8", model="gemini-3.8-flash-low")
    deltas = [e.text for e in events if isinstance(e, TextDelta)]
    assert deltas == ["BANANA", "\n"]
    done = events[-1]
    assert isinstance(done, TurnDone)
    assert done.response == "BANANA\n"
    assert done.usage["input_tokens"] == 28486


def test_parse_quota_from_result():
    events = _parse_file("quota_exhausted.ndjson")
    err = events[-1]
    assert isinstance(err, TurnError)
    assert err.code == "quota"
    assert err.retry_after == "96h24m30s"
    assert "quota" in err.message.lower()


def test_parse_quota_from_stderr_agy_error():
    lines = (FIXTURES / "quota_exhausted.stderr.txt").read_text(encoding="utf-8").splitlines()
    events = [e for e in (parse_event(l) for l in lines) if e is not None]
    assert len(events) == 1
    err = events[0]
    assert isinstance(err, TurnError)
    assert err.code == "quota"
    assert err.retry_after == "96h24m30s"
    assert "RESOURCE_EXHAUSTED" in err.message


def test_parse_error_no_content_fixture_ignores_plain_text_line():
    events = _parse_file("input_error_no_content.ndjson")
    assert isinstance(events[0], Init)
    assert isinstance(events[-1], TurnError)
    assert events[-1].code == "other"
    assert len(events) == 2  # the bare "error: ..." line is not JSON -> None


def test_parse_auth_error():
    line = json.dumps({"event": "result", "result": {"status": "ERROR", "error": "You are not logged into Antigravity. Run agy and /login."}})
    ev = parse_event(line)
    assert isinstance(ev, TurnError) and ev.code == "auth"


def test_parse_model_error():
    line = json.dumps({"event": "error", "error": "model 'foo-bar' not found"})
    ev = parse_event(line)
    assert isinstance(ev, TurnError) and ev.code == "model"


def test_parse_tool_and_thinking():
    tool = {"event": "step_update", "step_update": {"step_type": "tool", "state": "ACTIVE",
                                                     "tool_info": {"parameters": {"ToolName": "word_insert_table"}}}}
    ev = parse_event(json.dumps(tool))
    assert ev == ToolStatus(tool="word_insert_table", state="ACTIVE")
    tool2 = {"event": "step_update", "step_update": {"step_type": "tool", "state": "DONE", "tool_name": "x"}}
    assert parse_event(json.dumps(tool2)) == ToolStatus(tool="x", state="DONE")
    th = {"event": "step_update", "step_update": {"step_type": "thinking", "state": "ACTIVE", "text_delta": "hmm"}}
    assert parse_event(json.dumps(th)) == Thinking(text="hmm")
    th2 = {"event": "step_update", "step_update": {"step_type": "agent_response", "thinking_delta": "deep"}}
    assert parse_event(json.dumps(th2)) == Thinking(text="deep")


def test_parse_garbage_returns_none():
    assert parse_event("") is None
    assert parse_event("hello") is None
    assert parse_event("{not json") is None
    assert parse_event(json.dumps({"event": "step_update", "step_update": {"step_type": "user_input"}})) is None


# ---------------------------------------------------------------- models


SAMPLE_MODELS = (
    "Fetching available models...\n"
    "gemini-4-argon-medium\tGemini 4 Argon (Medium)\n"
    "gemini-3.8-flash-low\tGemini 3.8 Flash (Low)\n"
    "claude-opus-5-5-high\tClaude Opus 5.5 (High)\n"
    "gpt-oss-120b-medium\tGPT-OSS 120B (Medium)\n"
    "\n"
)


def test_parse_models():
    models = parse_models(SAMPLE_MODELS)
    assert [m["id"] for m in models] == [
        "gemini-4-argon-medium", "gemini-3.8-flash-low", "claude-opus-5-5-high", "gpt-oss-120b-medium"]
    assert models[0] == {"id": "gemini-4-argon-medium", "label": "Gemini 4 Argon (Medium)",
                         "family": "gemini-4-argon", "effort": "medium"}
    assert models[2]["family"] == "claude-opus-5-5"
    assert models[3]["family"] == "gpt-oss-120b"


def test_list_models_with_fake_and_cache(tmp_path, monkeypatch):
    monkeypatch.setenv("FAKE_AGY_ARGV_LOG", str(tmp_path / "argv.log"))
    models = list_models(PREFIX)
    assert any(m["id"] == "gemini-3.8-flash-low" for m in models)
    cache = ModelCache(ttl=60)
    first = cache.get(PREFIX)
    second = cache.get(PREFIX)
    assert first == second
    calls = (tmp_path / "argv.log").read_text(encoding="utf-8").splitlines()
    assert len(calls) == 2  # list_models + one cached fetch


def test_find_agy_localappdata(tmp_path, monkeypatch):
    monkeypatch.setenv("PATH", str(tmp_path / "empty"))
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    assert find_agy() is None
    exe = tmp_path / "agy" / "bin" / "agy.exe"
    exe.parent.mkdir(parents=True)
    exe.write_bytes(b"")
    assert find_agy() == str(exe)


# ---------------------------------------------------------------- AgyProcess


async def _collect(proc, text):
    return [ev async for ev in proc.send(text)]


def _env(tmp_path, mode="success", **extra):
    env = {"FAKE_AGY_MODE": mode, "FAKE_AGY_ARGV_LOG": str(tmp_path / "argv.log"),
           "FAKE_AGY_STATE_DIR": str(tmp_path / "state")}
    env.update(extra)
    return env


def _text(events):
    return "".join(e.text for e in events if isinstance(e, TextDelta))


def test_single_turn_success(tmp_path):
    async def main():
        proc = AgyProcess(None, "gemini-3.8-flash-low", command_prefix=PREFIX, cwd=str(tmp_path), env=_env(tmp_path))
        await proc.start()
        events = await _collect(proc, "hello world")
        await proc.close()
        return proc, events

    proc, events = asyncio.run(main())
    assert isinstance(events[0], Init)
    assert proc.conversation_id == events[0].conversation_id
    assert _text(events) == "ECHO: hello world"
    assert isinstance(events[-1], TurnDone)
    assert events[-1].response == "ECHO: hello world"
    argv = json.loads((tmp_path / "argv.log").read_text(encoding="utf-8").splitlines()[0])
    assert argv[argv.index("--model") + 1] == "gemini-3.8-flash-low"
    assert argv[argv.index("--input-format") + 1] == "stream-json"
    assert argv[argv.index("--output-format") + 1] == "stream-json"
    assert argv[-2:] == ["-p", ""]


def test_args_include_effort_conversation_and_extra(tmp_path):
    async def main():
        proc = AgyProcess(None, "m", effort="high", conversation_id="conv-1", command_prefix=PREFIX,
                          cwd=str(tmp_path), extra_args=("--sandbox",), env=_env(tmp_path))
        events = await _collect(proc, "x")  # send() auto-starts
        await proc.close()
        return events

    events = asyncio.run(main())
    assert isinstance(events[-1], TurnDone)
    argv = json.loads((tmp_path / "argv.log").read_text(encoding="utf-8").splitlines()[0])
    assert argv[argv.index("--effort") + 1] == "high"
    assert argv[argv.index("--conversation") + 1] == "conv-1"
    assert "--sandbox" in argv


def test_multi_turn_memory_same_process(tmp_path):
    async def main():
        proc = AgyProcess(None, "m", command_prefix=PREFIX, cwd=str(tmp_path), env=_env(tmp_path, "remember"))
        t1 = await _collect(proc, "the code word is PISTACHIO")
        t2 = await _collect(proc, "what was the code word?")
        await proc.close()
        return t1, t2

    t1, t2 = asyncio.run(main())
    assert _text(t1) == "PREV: NONE"
    assert _text(t2) == "PREV: the code word is PISTACHIO"
    # still one process
    assert len((tmp_path / "argv.log").read_text(encoding="utf-8").splitlines()) == 1


def test_unicode_and_newlines_roundtrip(tmp_path):
    msg = "Caffè «test» €75.000\nseconda riga"

    async def main():
        proc = AgyProcess(None, "m", command_prefix=PREFIX, cwd=str(tmp_path), env=_env(tmp_path))
        events = await _collect(proc, msg)
        await proc.close()
        return events

    assert _text(asyncio.run(main())) == "ECHO: " + msg


def test_crash_is_turn_error_not_success(tmp_path):
    async def main():
        proc = AgyProcess(None, "m", command_prefix=PREFIX, cwd=str(tmp_path), env=_env(tmp_path, "crash"))
        events = await _collect(proc, "hi")
        await proc.close()
        return events

    events = asyncio.run(main())
    assert not any(isinstance(e, TurnDone) for e in events)
    err = events[-1]
    assert isinstance(err, TurnError)
    assert "boom" in err.message


def test_quota_turn_error_with_retry_after(tmp_path):
    async def main():
        proc = AgyProcess(None, "m", command_prefix=PREFIX, cwd=str(tmp_path), env=_env(tmp_path, "quota"))
        events = await _collect(proc, "hi")
        await proc.close()
        return events

    err = asyncio.run(main())[-1]
    assert isinstance(err, TurnError)
    assert err.code == "quota"
    assert err.retry_after == "96h24m30s"


def test_auth_error(tmp_path):
    async def main():
        proc = AgyProcess(None, "m", command_prefix=PREFIX, cwd=str(tmp_path), env=_env(tmp_path, "auth"))
        events = await _collect(proc, "hi")
        await proc.close()
        return events

    err = asyncio.run(main())[-1]
    assert isinstance(err, TurnError) and err.code == "auth"


def test_restart_resumes_after_process_died_between_turns(tmp_path):
    async def main():
        env = _env(tmp_path, "remember", FAKE_AGY_EXIT_AFTER_TURN="1")
        proc = AgyProcess(None, "m", command_prefix=PREFIX, cwd=str(tmp_path), env=env)
        t1 = await _collect(proc, "remember OLIVE")
        await asyncio.sleep(0.3)  # fake exits after the turn
        t2 = await _collect(proc, "what?")
        await proc.close()
        return proc, t1, t2

    proc, t1, t2 = asyncio.run(main())
    assert isinstance(t1[-1], TurnDone)
    assert isinstance(t2[-1], TurnDone)
    assert _text(t2) == "PREV: remember OLIVE"
    lines = (tmp_path / "argv.log").read_text(encoding="utf-8").splitlines()
    assert len(lines) == 2
    argv2 = json.loads(lines[1])
    assert argv2[argv2.index("--conversation") + 1] == proc.conversation_id


def test_cancel_then_next_send_resumes_with_conversation(tmp_path):
    async def main():
        env = _env(tmp_path, "slow")
        proc = AgyProcess(None, "m", command_prefix=PREFIX, cwd=str(tmp_path), env=env)
        events = []

        async def run_turn():
            async for ev in proc.send("long task"):
                events.append(ev)
                if isinstance(ev, TextDelta):
                    await proc.cancel()

        await asyncio.wait_for(run_turn(), timeout=20)
        conv = proc.conversation_id
        env["FAKE_AGY_MODE"] = "success"
        t2 = await asyncio.wait_for(_collect(proc, "again"), timeout=20)
        await proc.close()
        return conv, events, t2

    conv, events, t2 = asyncio.run(main())
    assert conv
    assert isinstance(events[-1], TurnError)
    assert events[-1].code == "cancelled"
    assert isinstance(t2[-1], TurnDone)
    lines = (tmp_path / "argv.log").read_text(encoding="utf-8").splitlines()
    assert len(lines) == 2
    argv2 = json.loads(lines[1])
    assert argv2[argv2.index("--conversation") + 1] == conv


def test_abandoned_turn_does_not_leak_output_into_next_turn(tmp_path):
    async def main():
        env = _env(tmp_path, "slow")
        proc = AgyProcess(None, "m", command_prefix=PREFIX, cwd=str(tmp_path), env=env)
        gen = proc.send("first")
        async for ev in gen:
            if isinstance(ev, TextDelta):
                break
        await gen.aclose()
        env["FAKE_AGY_MODE"] = "success"
        t2 = await asyncio.wait_for(_collect(proc, "second"), timeout=20)
        await proc.close()
        return t2

    t2 = asyncio.run(main())
    assert _text(t2) == "ECHO: second"


def test_missing_executable_is_turn_error(tmp_path):
    async def main():
        proc = AgyProcess(str(tmp_path / "nope.exe"), "m", cwd=str(tmp_path))
        events = await _collect(proc, "hi")
        await proc.close()
        return events

    events = asyncio.run(main())
    assert len(events) == 1 and isinstance(events[0], TurnError)


def test_env_has_profile_vars(tmp_path):
    proc = AgyProcess(None, "m", command_prefix=PREFIX, cwd=str(tmp_path), env={"FOO": "1"})
    env = proc.build_env()
    assert env["FOO"] == "1"
    for key in ("USERPROFILE", "HOME", "APPDATA", "LOCALAPPDATA"):
        if sys.platform == "win32":
            assert env.get(key), key
