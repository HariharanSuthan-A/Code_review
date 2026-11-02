# Code‑Review Report – `voice‑excel‑agent`

**Project overview**  
The codebase implements a “text‑only” Excel assistant.  The backend is a
FastAPI server that exposes a RESTful API for workbook operations and a
WebSocket stream for real‑time status updates.  The agent layer is built on
LangGraph’s ReAct pattern and operates only over a tightly‑controlled
tool‑set that directly manipulates an `openpyxl` workbook.  All Excel
formula strings are evaluated by a custom, *whitelisted* parser – no
`eval` or `exec` is used anywhere.  The application uses a small SQLite
store (`history.sqlite3`) to persist chat history.

| Layer | Primary Language | Main framework / library | Purpose |
|-------|-----------------|--------------------------|---------|
| Backend API | **Python 3.11** | FastAPI, uvicorn, `langchain_core` (LLM & tools), `openpyxl` | Handles HTTP/WS I/O, calls the tool set. |
| Agent logic | Python | LangGraph (ReAct) + LangChain models | Chat, tool‑calling, confirmation gating. |
| Workbook engine | Python | `openpyxl` | In‑memory workbook representation + snapshot/undo. |
| Formula evaluator | Python | Custom parser in `formula_engine.py` | Safe evaluation of typical Excel formulas. |
| Persistence | SQLite (Python `sqlite3`) | lightweight chat history store | Persist human–assistant dialogue. |
| CLI / scripts | Python | `dotenv` | Load environment variables. |

> **Key design decision** – *No code execution*: All user‑supplied strings are
> parsed and validated before being passed to any tool.  The only "dynamic"
> behaviour is the LLM’s choice of which pre‑written tool to invoke.

---

## 1. Strengths

| Area | Benefit / Evidence |
|------|--------------------|
| **Explicit API surface** | `excel_tools.py` exposes a *finite*, type‑annotated list of functions (see `ALL_TOOLS` list).  This makes the attack surface tiny – the LLM can only request the operations it knows. |
| **Safe formula evaluation** | `formula_engine.py` implements a hand‑rolled parser that does *not* use Python `eval`.  The evaluator is limited to the whitelist in `_FUNCTIONS`.  Any unsupported token raises `FormulaEvalError`, which is co‑erced to `None`. |
| **File‑system isolation** |  
> `WorkbookManager._safe_filename` and `_contained_path` ensure uploads cannot leave the `workbooks/` dir.  Even a malicious file name like `../../tmp/pwned.xlsx` is normalised to `pwned.xlsx` (see lines ~115‑139). |
| **Undo capacity** | Snapshots are kept in memory as `bytes` and a 20‑step stack is enforced (`MAX_UNDO_STACK`), which protects against accidental data loss (chunk ~190‑208). |
| **Clear confirmation gating** | `DESTRUCTIVE_TOOLS` set in `excel_tools.py` (<200) plus `LARGE_RANGE_ROW_THRESHOLD` guarantee the agent will *screen* destructive operations before actually executing them. |
| **Observability** | WebSocket status framings (“Thinking…”, “Editing Excel…”, etc.) give the UI a simple way to show live progress. |
| **No arbitrary code execution** | The `llm_client.py` module *only* creates a `ChatOllama` or `ChatOpenAI` client; no Python code is run on behalf of the user.  This meets the core requirement behind the "Security Note" in `excel_tools.py`. |
| **Persisted chat history** | SQLite table `chat_history` (see `db.py`, lines ~12‑34) is sanitized via parameterised queries, preventing SQL injection. |

---

## 2. Potential Technical Issues & Security Concerns

| Category | Issue | Supporting code | Why it matters | Suggested fix |
|----------|-------|-----------------|----------------|----------------|
| **Concurrency / session handling** | Single global `WorkbookSession` – collisions between users or concurrent WebSocket connections. | `WorkbookManager` holds a single `session`; WebSocket handler does not scope per‑user (`session_id`). | A malicious user could read/write another user’s workbook. | Replace with a session dictionary keyed by `session_id` (e.g., `self.sessions[session_id]`). Use per‑session locks. |
| **File upload payload size** | No limit on uploaded file size. | `upload_workbook` in `routes.py` (lines ~36‑48). | DOS via huge files. | Enforce a max size (`max_upload_size = 10_000_000`) and return 413 if exceeded. |
| **Database path injection** | `HISTORY_DB_PATH` constructed from env variable with a relative path. | `db.py` line ~4. | Someone could craft `HISTORY_DB_PATH="../outside.db"` to store history elsewhere. | Normalize via `os.path.abspath` and reject paths outside app root. |
| **Error surface** | `_guard` returns raw `str(exc)` for any exception, which can leak implementation details (e.g., stack traces, workbook names). | `excel_tools._guard` (lines ~50‑58). | Might expose internal file names / exception types. | Mask unexpected exceptions: return a generic `"Tool error"` message. |
| **Large range clearing** | `delete_row` & `clear_range` use no check for range size. | `delete_row` (lines ~274‑283), `clear_range` (lines ~551‑564). | A malicious user could clear millions of cells. | Enforce `DELETE_ROW` threshold or prompt confirmation for > `LARGE_RANGE_ROW_THRESHOLD`. |
| **No authentication** | All endpoints are public. | `routes.py` – no security decorators. | Anyone can manipulate any workbook. | Add JWT or token‐based auth; restrict `session_id` scope. |
| **OpenAI / Ollama request leak** | LLM calls are not rate‑limited; environment variable may expose key. | `llm_client.py` – loads `LLM_API_KEY` but never sanitises. | API abuse or credential leakage. | Use IAM roles or separate secrets manager; enforce rate limits. |
| **SQLite concurrency** | SQLite allows multiple readers but only one writer; simultaneous writes may block. | `db.py` – uses context manager per operation. | Under high concurrency the DB may become a bottleneck. | Use a connection pool or an async wrapper (e.g., `aiosqlite`). |
| **Potential memory leak** | Undo snapshots stored as large bytes may grow if many edits and if users open multiple workbooks. | `WorkbookManager._snapshot` (lines ~106‑116). | With a global session, memory can balloon. | Move stack into per‑session dict and drop older snapshots after a threshold not just size. |
| **Locale / timezone** | `WorkbookManager` imports `datetime, timezone` but never uses them. | Import at top of file. | Unnecessary imports; could cause confusion. | Remove unused imports. |
| **Unvalidated tool arguments** | `excel_tools` functions use `_coerce_value` for all string inputs.  If user sends an empty string for numeric field, it maps to `None`. But some workbook methods expect integers (e.g. row numbers). | `add_row`, `write_cell`, etc., call `_guard(wm.add_row, at_row=at_row if at_row > 0 else None)` (lines ~266). | Passing `0` returns `None`, leading to no `"at_row"` argument → correct, but other tools may receive `None` incorrectly. | Add explicit validation in static type guards or raise `ValueError` for `None` where not allowed. |
| **Missing `apply_formula` implementation** | Only the first ~5000 chars of `workbook_manager.apply_formula` are seen; the tail is truncated. | `workbook_manager.apply_formula` top part (lines ~232‑225). | If the implementation is incomplete it could silently do nothing or raise errors at runtime. | Verify the method is fully implemented or added via code generation. |
| **Path traversal in `.download_workbook`** | The response returns the absolute file path from `wm.session.path`. | `routes.py` line ~70. | If the session path could be manipulated, a user might download arbitrary files. | Verify the file path is under the allowed `workbooks/` directory before streaming. |

---

## 3. Code‑Specific Observations

### `excel_tools.py` – Tool wrappers
* **Guarding policy** – All tools are wrapped with `_guard` (lines ~47‑55) that turns every exception into a dict with `"success": False`.  This is good for robustness but the exception string is leaked.  
* **Optional arguments** – Heavy use of sentinel `""` or `0` to indicate absence.  This can lead to confusion for the LLM which must explicitly send those sentinels, otherwise the underlying workbook methods are called with `None` where they expect a value.  Example: `add_row(at_row: int = 0)` expects `0` to mean “append”; if the LLM sends omitted argument it may get `None`.  
* **Use of `list` and `dict` only** – For `OLLAMA` compatibility as documented.  Great.

### `workbook_manager.py`
* **Header styling** – `_style_header_row` is idempotent (calls once per sheet).  
* **Cell coordinate lookup** – `write_cell` and `write_range` perform regex validation and handling of row 1 header styling.  
* **Large writes** – `write_range` accepts a 2‑D list; there is no check for size.  The post‑safety net lies in the UI threshold for confirmation.  
* **Undo** – Works by serialising the entire workbook to `bytes`.  This approach is memory‑intensive but simple.  
* **Thread safety** – `WorkbookSession.lock` is an `RLock`, but not used in every method that mutates state; many mutating methods call `_snapshot()` without acquiring the lock.  The optional `with self.session.lock:` wrapper is only applied to `new_workbook` / `open_workbook` / `undo`.  A concurrent request that triggers two simultaneous writes could race and corrupt the workbook.  Recommendation: wrap all mutating operations in `with self.session.lock`.  

### `formula_engine.py`
* **Parser** – Handles typical formulas: `SUM`, `AVERAGE`, `IF`, `LOOKUP`, `TEXTJOIN`, etc.  
* **Safety** – `evaluate_on_workbook` recursively resolves cell references (context).  Circular references raise `FormulaEvalError`.  
* **Return type** – All functions return plain Python scalars; when the workbook stores a formula, `WorkbookManager._display_value` calls `try_evaluate_on_workbook` to expose the answer.  

---

## 4. Suggested Improvements

| Area | Recommendation | How to implement |
|------|----------------|------------------|
| **Per‑user session isolation** | Replace single‐global `WorkbookSession` with `WorkbookSessionManager` that keeps a dict keyed by `session_id`. | `WorkbookManager.__init__` → `self.sessions: Dict[str, WorkbookSession] = {}`.  Methods now `session = self.sessions.get(sid, DEFAULT_SESSION)`. |
| **Input validation** | Add Pydantic models for tool arguments instead of ad‑hoc coercion. | Define `BaseModel` for each tool; use `@tool` signature `(arg: Foo)`.  This removes the need for manual `_coerce_*` functions. |
| **Error sanitisation** | Replace `_err` with a generic error‑message. | In `excel_tools._err` do `return {"success": False, "error": "Tool execution failed"}`. Keep detailed logging in server logs. |
| **Upload size limit** | Add middleware or explicit check in `upload_workbook`. | `if file.size > MAX_UPLOAD_SIZE: raise HTTPException(413, "File too large")`. |
| **Auth** | Add FastAPI dependency that extracts `X-Session-ID` and validates against session store. | `def get_current_session(session_id: str = Header(...))`. |
| **Rate limiting** | Wrap API endpoints with `fastapi-limiter` or manual token bucket. | Decorate `upload_workbook`, `agent_message`, etc. |
| **Undo fallback** | Persist undo stack to SQLite or disk to survive restarts. | Add `undo_history.sqlite3` table; on startup recover stack. |
| **Memory usage** | Switch snapshot from raw bytes → `io.BytesIO()` → `numpy` array to reduce overhead?  Or use `openpyxl.writer.write_only` for huge workbooks. | Only if workbook size tends to > 50 MB. |
| **Logging** | Use structured JSON logging.  Add `LOG_LEVEL` from env. | Replace `print` with `logger`. |
| **Unit tests** | Expand test suite to cover edge cases: deleting last sheet, circular formulas, large range clearing. | Add `tests/test_workbook_manager.py` & `pytest`. |
| **File download security** | Validate download path before streaming. | `if not wm.session.path.startswith(os.path.abspath("workbooks/")): raise HTTPException(403)`. |

---

## 5. Questions for the Development Team

1. **Session handling** – Is the application ever intended to serve multiple users concurrently, or will it remain single‑user MVP?  
2. **Undo persistence** – Do you anticipate a need to recover the undo history after the process restarts?  
3. **Max upload limit** – Is there an accepted maximum workbook size that must be enforced?  
4. **Security review** – Have you run a static analysis or OWASP tool (e.g., Bandit) on the code base?  
5. **Testing** – What is the current coverage of `workbook_manager`’s mutation methods?  

---

## 6. Summary

The codebase demonstrates a solid, Security‑first design: explicit tool set, safe formula evaluation, and careful file‑system boundaries.  The main technical risk arises from **multi‑user concurrency**: a single global workbook session is shared across all API/WS calls, which could lead to data corruption or leakage.  Additional mitigations (upload size limits, auth, error masking and session isolation) are recommended before deploying in production.

*Open issues*: The `apply_formula` implementation appears cut off in the provided snapshot – verify that the method fully writes the value and handles errors.  

**Next steps**  
1. Add per‑user session support.  
2. Introduce request validation (Pydantic).  
3. Mask exception messages.  
4. Implement auth + rate limiting.  
5. Run automated security scan (Bandit, Trivy).  

Once these changes are merged, the system will be well positioned for safe, scalable operation.