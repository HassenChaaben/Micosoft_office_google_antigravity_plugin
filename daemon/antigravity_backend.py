"""
Antigravity subscription backend.
Sends the user's request to the locally installed Antigravity CLI (`agy -p`), which is
already signed in with the user's Antigravity subscription, and streams the answer back.
"""

import json
import os
import shutil
import subprocess
from typing import Iterator, List, Optional, Tuple

DEFAULT_MODEL = "gemini-3.8-flash-high"
MODEL_MAP = {
    "gemini-3.8-flash-high": "gemini-3.8-flash-high",
    "gemini-3.8-flash-medium": "gemini-3.8-flash-medium",
    "gemini-pro": "gemini-4-argon-medium",
    "gemini-deep-think": "gemini-4-argon-high",
}
# Windows caps a command line at 32,767 chars and the prompt travels as an argument.
MAX_PROMPT_CHARS = 24_000
TIMEOUT_SECONDS = 180

WORKSPACE_DIR = os.path.expandvars(r"%LOCALAPPDATA%\GeminiAntigravityOffice\agy_workspace")

SYSTEM_RULES = (
    "You are Gemini Antigravity, a writing assistant running inside Microsoft {host}.\n"
    "Rules:\n"
    "- Reply ONLY with the content to show the user, in GitHub-flavored markdown.\n"
    "- Do NOT call tools, run commands, browse, or create/edit files.\n"
    "- Follow the user's request exactly (e.g. if they name columns, use exactly those columns).\n"
    "- Use markdown tables for tabular data and headings for structure.\n"
    "- No filler like 'Sure, here is'.\n"
)


def find_agy() -> Optional[str]:
    found = shutil.which("agy")
    if found:
        return found
    default = os.path.expandvars(r"%LOCALAPPDATA%\agy\bin\agy.exe")
    return default if os.path.isfile(default) else None


def build_prompt(prompt: str, host: str, doc_context: str) -> str:
    parts = [SYSTEM_RULES.format(host=host)]
    if doc_context:
        budget = max(0, MAX_PROMPT_CHARS - len(parts[0]) - len(prompt) - 200)
        ctx = doc_context if len(doc_context) <= budget else doc_context[:budget] + "\n[...document truncated...]"
        parts.append(f'Active {host} document content:\n"""\n{ctx}\n"""\n')
    parts.append(f"User request:\n{prompt}")
    return "\n".join(parts)


def build_command(agy_path: str, full_prompt: str, model: str) -> List[str]:
    return [
        agy_path,
        "--output-format", "stream-json",
        "--model", MODEL_MAP.get(model, DEFAULT_MODEL),
        "--disable-slash-commands",
        "--sandbox",
        "--print-timeout", f"{TIMEOUT_SECONDS}s",
        "-p", full_prompt,
    ]


def parse_stream_line(line: str) -> Optional[Tuple[str, str]]:
    line = line.strip()
    if not line.startswith("{"):
        return None
    try:
        event = json.loads(line)
    except json.JSONDecodeError:
        return None

    kind = event.get("event")
    if kind == "step_update":
        step = event.get("step_update") or {}
        text = step.get("text_delta")
        if step.get("step_type") == "agent_response" and text:
            return ("delta", text)
        return None
    if kind == "result":
        result = event.get("result") or {}
        if result.get("status") == "SUCCESS":
            return ("done", "")
        detail = result.get("error") or result.get("message") or result.get("status") or "unknown error"
        return ("error", f"Antigravity returned an error: {detail}")
    if kind == "error":
        return ("error", f"Antigravity error: {event.get('error') or event.get('message') or line}")
    return None


def stream_events(agy_path: str, prompt: str, host: str, doc_context: str, model: str) -> Iterator[Tuple[str, str]]:
    """Yields ("delta", text), then ("done", "") or ("error", message)."""
    os.makedirs(WORKSPACE_DIR, exist_ok=True)
    cmd = build_command(agy_path, build_prompt(prompt, host, doc_context), model)
    try:
        proc = subprocess.Popen(
            cmd,
            cwd=WORKSPACE_DIR,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            encoding="utf-8",
            errors="replace",
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except OSError as e:
        yield ("error", f"Could not start the Antigravity CLI ({agy_path}): {e}")
        return

    finished = False
    try:
        for line in proc.stdout:
            parsed = parse_stream_line(line)
            if parsed is None:
                continue
            yield parsed
            if parsed[0] in ("done", "error"):
                finished = True
                return
        proc.wait(timeout=10)
        if not finished:
            stderr = (proc.stderr.read() or "").strip()[-500:]
            yield ("error", f"Antigravity CLI exited with code {proc.returncode}. {stderr}".strip())
    finally:
        # The generator can be abandoned mid-stream (client closed the taskpane), so never leak the child.
        if proc.poll() is None:
            proc.kill()


WORD_OFFICE_INSTRUCTIONS = """# word-office MCP Server Instructions

## Native Microsoft Word Manipulation Best Practices
You are Gemini Antigravity, an expert AI operating directly inside Microsoft Word. You manipulate Word like a human expert using native Word objects.

### 1. Document Exploration
- Always call `word_get_document_outline` first to inspect:
  - Paragraph count and text preview
  - Existing tables (indices, row/column counts)
  - Active cursor/selection status
- Always check the outline before deciding where to insert new content or which element to modify.

### 2. Creating Tables (Strict Precision & Anti-Duplication)
- When the user asks for a table:
  - Check `word_get_document_outline` first.
  - If the user specifies columns (e.g. "product_name, quantity, price"), your `headers` array MUST contain EXACTLY those columns: `["product_name", "quantity", "price"]`.
  - If the user specifies a row count (e.g. '5 rows'), provide exactly that number of data rows in `rows`. If unspecified, provide 3 to 4 realistic data rows.
  - Call `word_insert_table` ONCE.
  - When `word_insert_table` returns `ok: true`, the table is 100% created and formatted in Word. DO NOT call `word_insert_table` again!

### 3. Modifying Existing Text & Paragraphs (In-Place, Never Duplicate)
- When the user asks to modify, rewrite, change, fix, expand, or correct existing text or paragraphs:
  - DO NOT call `word_insert_paragraph`! That creates an unwanted duplicate paragraph!
  - Check `selected_text` or inspect the outline to find the target paragraph's `index`.
  - Call `word_set_paragraph_text(paragraph_index, new_text)` to rewrite the paragraph in-place.
  - Call `word_replace_text(search_text, replacement_text)` to swap words or phrases.
  - Call `word_format_text` to modify bold, size, color, or highlight.

### 4. Enhancing & Modifying Tables In-Place
- When the user asks to "enhance font size, style", "format table", "change colors", or "add rows":
  - DO NOT create a second table!
  - Inspect the outline to find the index of the table (e.g. `table_index: 0`).
  - Use `word_format_table` to apply built-in styling, bold headers, or auto-fit.
  - Use `word_add_table_row` to append data rows to an existing table.
  - Use `word_set_table_cell` to update specific cell values or bold headers.
  - When the user asks to remove, delete, or drop a column (e.g. "remove third column", "elimina la seconda colonna"):
    - NEVER delete and recreate the table! Call `word_delete_column(table_index, col_index)` to delete the column directly in-place.
  - When the user asks to remove or delete a row:
    - Call `word_delete_row(table_index, row_index)` to delete the row directly in-place.

### 5. Inserting Pictures & Visuals
- Use `word_insert_image(source, caption, location, width_px, height_px)` to add pictures.
- `source` can be a URL, base64 data, or a keyword like `'logo'`, `'chart'`, `'diagram'`, `'banner'`.

### 6. Reading and Deleting Tables
- Use `word_read_table(table_index)` to read all cell values of a table.
- Use `word_delete_table(table_index)` to delete an entire table if requested or if an unwanted duplicate exists.

### 7. Styles & Formatting
- Standard styles: 'Heading 1', 'Heading 2', 'Heading 3', 'Normal', 'Title'.
- Note: Word may run in English, French (Titre 1), or Italian (Titolo 1). Always pass standard English names ('Heading 1', 'Normal'); the executor translates them automatically.

### 8. Forms, Project Management Templates & Batch Cell Updates
- When filling out structured tables or forms (e.g. Project Charter, Business Case, Risk Register, Cost Baseline):
  - Use `word_fill_form_fields(fields={"Label": "Value"})` to fill fields by label name automatically. This eliminates merged-cell coordinate mistakes and handles in-place updates seamlessly!
  - Use `word_batch_update_table_cells(table_index, updates=[{"row": r, "col": c, "text": "..."}])` to update multiple cells across a table at once in a single fast, atomic call.
  - When the user asks for modifications (e.g. "Change budget to €75.000" or "Set status to Completed"): use `word_fill_form_fields` or `word_batch_update_table_cells` to modify in-place.
"""


def ensure_mcp_registered(server_name: str = "word-office", server_url: str = "https://127.0.0.1:18890/mcp"):
    """
    Registers the word-office MCP server in the user's Antigravity CLI configuration if missing.
    Also ensures instructions.md is deployed to guide the model with native Word best practices.
    """
    # Write instructions.md into both cli and app directories
    for sub in ["antigravity-cli", "antigravity"]:
        target_dir = os.path.expandvars(rf"%USERPROFILE%\.gemini\{sub}\mcp\{server_name}")
        try:
            os.makedirs(target_dir, exist_ok=True)
            with open(os.path.join(target_dir, "instructions.md"), "w", encoding="utf-8") as f:
                f.write(WORD_OFFICE_INSTRUCTIONS)
        except Exception:
            pass

    agy_path = find_agy()
    if not agy_path:
        return False
    try:
        # Check if already present
        res = subprocess.run(
            [agy_path, "mcp", "list"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            encoding="utf-8",
            errors="replace",
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            timeout=10
        )
        if server_name in (res.stdout or ""):
            return True

        # Register server
        add_res = subprocess.run(
            [agy_path, "mcp", "add", server_name, server_url],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            encoding="utf-8",
            errors="replace",
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            timeout=15
        )
        return add_res.returncode == 0
    except Exception as e:
        print(f"[AntigravityBackend] Failed to ensure MCP server '{server_name}' registered: {e}")
        return False
