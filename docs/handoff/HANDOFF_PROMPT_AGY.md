# Handoff prompt for agy (Antigravity) — Office agent, branch `office_agent_001`

Paste everything below the line into agy, after running `/wake_up`.

---

/wake_up

You are continuing work that Claude Code started on the Gemini Antigravity Office add-in.

- **Project:** `C:\Users\user\Downloads\Personal_projects\Google_antigravity_microsoft_office_plugin`
- **Git branch:** `office_agent_001`. Keep the git identity `hassen <haythemnuira@gmail.com>`; this is a personal project.
- **Python:** `C:\Python312\python.exe`. Do NOT use `python -I`, because pywin32 is in the user site.

## Read first, in this order
1. `C:\Users\user\.claude\plans\this-is-exe-version-federated-flask.md`: the approved plan (Milestones 0–6).
2. `docs/spikes/2026-10-29-m0-spike.md`: verified facts about agy, the models, auth and snapshot speed.
3. `git log --oneline`: every milestone is one commit with an Italian message.
4. `daemon/agent_session.py` (the agent loop and its rules), `daemon/mcp_server.py` (the daemon tools) and `daemon/main.py` (the endpoints).

## Goal (what the user wants)
A Microsoft Word agent that works like a **senior Office expert**:
- It **sees and analyses the document before every prompt**, so it is never blind.
- It plans and acts with **native Office features** (styles, numbering, sections, fields/TOC, content controls, SmartArt, native charts, shapes). It uses the web only for facts, never for pictures.
- It verifies its own work.
- It remembers chats.
- It is fast.
- It is installable with **one click by non-technical clients**, and Antigravity auth is automatic.

After Word, Excel and PowerPoint will reuse the same core.

## Already DONE and committed (do not redo)
- **Perception:**
  - `daemon/office/word_snapshot.py` turns OOXML into a snapshot plus the Document Map (field paths, record grids, merges, styles, direct formatting).
  - `daemon/office/perception.py` handles caching.
  - The taskpane `captureDocumentContext()` (`web/word-agent.js`) sends `body.getOoxml()`, the selection and the headers/footers on every run.
- **Verify:** `daemon/office/verify.py` plus the `office_verify` tool, which re-captures the document through `__capture_document`.
- **Content:** `daemon/office/ooxml_builder.py` plus the `word_insert_content` tool. The taskpane runs `word_insert_ooxml` as one `insertOoxml` call. Locations: `cursor | start | end | replace_selection | after:p<N> | after:<heading> | after:ga:<id> | replace:ga:<id>`.
- **Forms:** `daemon/office/fill.py` plus the `word_fill_fields` tool. Keys are a path, `#record`, `T{t}r{r}c{c}` or a content-control tag. The taskpane runs `word_fill_cells` as one batched write that keeps the formatting.
- **Runtime:**
  - `daemon/core/agy_process.py`: a warm `agy --input-format stream-json` process per chat. It resumes with `--conversation`, and quota and auth errors are typed.
  - `daemon/core/store.py`: SQLite history in `%LOCALAPPDATA%\GeminiAntigravityOffice\history.db`.
  - `AgentSession.run()` wires it all together. Its SSE events are `chat`, `analysis`, `skills`, `plan`, `text`, `tool_call`, `agent_finished`, `error{code,retry_after}`.
- **Skills:**
  - 13 expert skills in `daemon/skills/*/SKILL.md`.
  - The router `daemon/core/skills.py` injects them into the first message of a chat.
  - Daemon tools: `skill_load`, `memory_save`, `memory_search`, `agent_plan`.
- **Endpoints:** `/api/models`, `/api/chats` (GET, PATCH, DELETE), `/api/chats/{id}`.
- **Tests:** 194 Python unit tests pass, plus 20 JS tests.

## Work that may be UNCOMMITTED in the working tree (review → test → commit)
Two subagents were still running when Claude stopped. Check with `git status`.

1. **Native COM tools**: `daemon/office/com_worker.py`, `daemon/office/word_com_tools.py`, `tests/test_word_com_tools.py`.
   - Run `python -m pytest -q tests/test_word_com_tools.py`. These tests use a HIDDEN `DispatchEx` Word instance; never touch the user's open Word.
   - If they pass, commit.
   - If the files are incomplete, finish them. The tool contract is: `word_insert_chart`, `word_insert_smartart`, `word_insert_diagram` (flowchart, wbs, network, gantt), `word_insert_toc`, `word_insert_caption`, `word_insert_cross_reference`, `word_update_fields`, `word_insert_section_break`, `word_set_header_footer`, `word_track_changes`, `word_add_comment`, `word_modify_style`, `word_apply_style`, `word_link_numbering_to_headings`, `word_set_checkbox`, `word_table_clone_row`.
2. **Taskpane UI**: this one is now COMMITTED. It added the history drawer, model picker, quota card with retry, plan, skills and analysis cards, and safe markdown. JS tests: `node --test tests/js/insert-ooxml.test.mjs tests/js/word-agent.test.mjs tests/js/fill-cells.test.mjs tests/js/taskpane-utils.test.mjs` (27 pass). Gaps it found, to do:
   - add a **Stop button** that calls `POST /api/agent/cancel`;
   - give the "already running" error a `code`;
   - make `/api/chats/{id}` return the plans and tool calls per turn, so a reopened chat shows its cards;
   - add an endpoint to start sign-in (see M6);
   - send display names for model families from `/api/models`;
   - unsaved documents (empty `doc_url`) cannot resume a chat.

## REMAINING WORK (in order)
1. **Wire the COM tools into the agent.**
   - Expose them through `daemon/mcp_server.py` as daemon tools with `executor=com`. They run in the daemon through `ComWorker`, which targets the document by matching `doc_url` against `Documents(i).FullName`. `AgentSession._current_doc_url` holds the URL.
   - Each visual must return a `ga:<id>` anchor (a bookmark or content control) so captions and cross-references can target it.
   - Re-run perception after COM edits.
2. **Close the contract gaps reported by the skills writer:**
   - add a `before:<anchor>` location;
   - define the `word_insert_cross_reference` kinds (`label_and_number`, `heading_number`, `page`);
   - define the `word_insert_diagram` fields (`critical`, `start`/`finish` gg/mm/aaaa, `milestone`, edge `type` FS/SS/FF/SF);
   - add a table of figures, image insertion with alt text, and the page-number format and restart;
   - let `word_set_checkbox` take a label anchor.
3. **Snapshot edge cases:**
   - nested tables are not indexed;
   - Office.js `getCell` on vertically merged rows may fail (COM `Rows(n)` does). Add a COM fallback in `fill`.
4. **Remove the remaining fakes:**
   - the canned offline answers in `daemon/gemini_client.py` (keep the offline mode for tests only);
   - the stub connectors in `daemon/mcp_bridge.py`: replace them with the real `agy mcp list` and `agy plugin list`;
   - the fake `daemon/auth.py`.
   
   Do not touch the legacy `/api/document/*` COM routes unless needed. `daemon/word_engine.py` has a `doc` NameError at lines 147 and 155.
5. **Live test inside real Word.**
   - Rebuild with `scripts/build.ps1`. Add the new packages and the `daemon/skills/**` data files to `gemini_office_daemon.spec`.
   - Test with the PM templates in `practical_tests/.../Template_Progetto_PM`:
     - "compila questo documento per il progetto Cloud4Job" on `2 Project charter.docx` and `23 Registro dei rischi.docx`;
     - a report with TOC, chart and SmartArt in a blank document;
     - a WBS diagram in `8 WBS…`;
     - a network with the critical path in `19 Reticolo…`;
     - memory: make "Progetto di Automazione Cloud4Job" a centred subtitle, restart, reopen the chat, then ask "make it blue".
6. **M6, one-click setup** (the clients are non-technical):
   - `installer/setup.iss`: **REMOVE** the `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--ignore-certificate-errors` entries in HKCU and HKLM. They are a security hole. The uninstaller must remove them too.
   - Generate a per-machine CA and localhost certificate at install time (`daemon/cert_gen.py`, exposed as `--setup-cert`). Stop shipping `certs/*.key`.
   - Check prerequisites: Office, WebView2, and the agy CLI (install it silently if missing).
   - Start the daemon, self-test (`/health` plus `agy models`), and offer "Open Word now".
7. **M6, auth:** the login is already shared in Windows Credential Manager (`gemini:antigravity`).
   - `/api/auth/status` checks `agy models`.
   - If the user is not signed in, show a "Sign in with Google" button.
   - Settings → Account → "Switch account" opens a visible interactive `agy` console where the user runs `/logout` and `/login`. The daemon then polls `agy models`.
8. **Manifest:** add `<Requirements>` (WordApi 1.3 minimum) and feature checks. Hide Outlook in the UI.

## Rules (from the user's CLAUDE.md — follow strictly)
- Strict TDD: write the failing test first and watch it fail, then implement. Check pytest's REAL exit code; never pipe it through `tail` alone.
- Run the unit suite: `python -m pytest -q tests --ignore=tests/test_final_e2e_live.py --ignore=tests/test_pm_templates_36.py --deselect tests/test_claude_harness_capabilities.py::test_reading_word_document`. That last test attaches to the real Word and is flaky.
- Make one commit per finished step, with an Italian message (`feat(area): …`).
- **Never add AI attribution** to commits, comments or PRs.
- Explain things to the user in simple English (B1–B2).
- Word tests must use a hidden `DispatchEx` instance and never touch the user's open documents.
- Never fake success: return real errors to the model and the user.
