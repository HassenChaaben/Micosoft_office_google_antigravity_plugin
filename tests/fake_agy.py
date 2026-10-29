"""Fake `agy` CLI for tests. Mimics `agy --input-format stream-json --output-format stream-json -p ""`.

Env:
  FAKE_AGY_MODE            success | remember | quota | auth | crash | slow   (default success)
  FAKE_AGY_ARGV_LOG        append one JSON line with argv[1:] per launch
  FAKE_AGY_STATE_DIR       persist "previous user message" per conversation id (proves --conversation resume)
  FAKE_AGY_EXIT_AFTER_TURN "1" -> exit after the first turn's result
Also supports `fake_agy.py models`.
"""

import json
import os
import sys
import time
import uuid

QUOTA = "Individual quota reached. Please upgrade your subscription to increase your limits. Resets in 96h24m30s."
MODELS = [
    ("gemini-4-argon-medium", "Gemini 4 Argon (Medium)"),
    ("gemini-3.8-flash-low", "Gemini 3.8 Flash (Low)"),
    ("claude-sonnet-5-5-medium", "Claude Sonnet 5.5 (Medium)"),
]


def emit(obj):
    sys.stdout.write(json.dumps(obj, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def step(conv, text, state="ACTIVE", step_type="agent_response"):
    emit({"event": "step_update", "step_update": {"conversation_id": conv, "step_index": 1, "state": state,
                                                    "step_type": step_type, "text_delta": text}})


def result(conv, status, response="", error=None):
    res = {"conversation_id": conv, "status": status, "response": response, "num_turns": 1,
           "usage": {"input_tokens": 10, "output_tokens": 2, "total_tokens": 12}}
    if error:
        res["error"] = error
    emit({"event": "result", "result": res})


def main():
    for stream in (sys.stdin, sys.stdout, sys.stderr):
        stream.reconfigure(encoding="utf-8")
    argv = sys.argv[1:]
    log = os.environ.get("FAKE_AGY_ARGV_LOG")
    if log:
        with open(log, "a", encoding="utf-8") as f:
            f.write(json.dumps(argv) + "\n")

    if argv[:1] == ["models"]:
        print("Fetching available models...")
        for mid, label in MODELS:
            print(f"{mid}\t{label}")
        return 0

    mode = os.environ.get("FAKE_AGY_MODE", "success")
    conv = argv[argv.index("--conversation") + 1] if "--conversation" in argv else str(uuid.uuid4())
    model = argv[argv.index("--model") + 1] if "--model" in argv else "fake"
    state_dir = os.environ.get("FAKE_AGY_STATE_DIR")
    state_file = os.path.join(state_dir, conv + ".json") if state_dir else None
    prev = None
    if state_file and os.path.exists(state_file):
        with open(state_file, encoding="utf-8") as f:
            prev = json.load(f).get("prev")

    emit({"event": "init", "conversation_id": conv, "init": {"model": model, "cwd": os.getcwd(), "tools": []}})

    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            content = json.loads(line)["message"]["content"]
        except Exception:
            print('error: stream input "user" message has no content')
            result(conv, "ERROR", error='stream input "user" message has no content')
            continue

        if mode == "quota":
            result(conv, "ERROR", error=QUOTA)
            sys.stderr.write('AGY_ERROR: {"short_error":"RESOURCE_EXHAUSTED (code 429): ' + QUOTA +
                             '","status":"RESOURCE_EXHAUSTED","error_code":429}\n')
            sys.stderr.flush()
            return 3
        if mode == "auth":
            result(conv, "ERROR", error="You are not logged into Antigravity.")
            continue
        if mode == "crash":
            step(conv, "partial")
            sys.stderr.write("boom: fatal internal error\n")
            sys.stderr.flush()
            return 1
        if mode == "slow":
            step(conv, "working...")
            time.sleep(60)
            result(conv, "SUCCESS", response="late")
            continue

        if mode == "remember":
            reply = "PREV: " + (prev if prev is not None else "NONE")
            prev = content
            if state_file:
                os.makedirs(state_dir, exist_ok=True)
                with open(state_file, "w", encoding="utf-8") as f:
                    json.dump({"prev": prev}, f)
        else:
            reply = "ECHO: " + content
        step(conv, "thinking about it", step_type="thinking")
        half = len(reply) // 2
        step(conv, reply[:half])
        step(conv, reply[half:], state="DONE")
        result(conv, "SUCCESS", response=reply)
        if os.environ.get("FAKE_AGY_EXIT_AFTER_TURN") == "1":
            return 0
    return 0


if __name__ == "__main__":
    sys.exit(main())
