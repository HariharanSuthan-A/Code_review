# Excela AI

Control and edit an Excel workbook with natural language (text chat for now) —
"talking to Excel through an AI assistant," not just dictation.

```
User Text → AI Excel Agent → Python Excel Tools → Excel Workbook → Verify → Response
```

## What's inside

- **FastAPI backend** (`backend/app`) with a LangGraph ReAct agent that can
  only act through a fixed set of ~46 typed Excel tools (`app/tools/excel_tools.py`)
  wrapping an OpenPyXL manager (`app/excel/workbook_manager.py`). **No `eval`/`exec`,
  no arbitrary code execution — ever.**
- **Multi-provider LLM** — point it at Ollama (fully local/free), OpenAI, or any
  other OpenAI-compatible endpoint (OpenRouter, LM Studio, vLLM…) via `.env`.
- **SQLite** for the chat transcript; an in-memory undo stack (workbook snapshots)
  for one-click undo.
- **Vanilla HTML/CSS/JS frontend** — spreadsheet preview, text chat composer, live
  status ("Analyzing… / Updating Excel… / ✓ Completed") streamed over the
  `/ws/agent` websocket, chat log, confirmation prompts for destructive actions.

## Quick start

### 1. Backend

```bash
cd backend
python -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt

cp ../.env.example .env
# edit .env — at minimum pick LLM_PROVIDER / LLM_MODEL

uvicorn app.main:app --reload --port 8000
```

Open **http://localhost:8000** — the FastAPI app serves the frontend directly,
so there's no separate frontend server to run.

### 2. Get an LLM running

**Easiest (fully local, free):**
```bash
ollama pull llama3.1        # or qwen2.5:14b, mistral-nemo, etc. — pick one good at tool calling
ollama serve
```
`.env` defaults already point at `http://localhost:11434/v1`.

**Or use OpenAI:**
```env
LLM_PROVIDER=openai
LLM_BASE_URL=https://api.openai.com/v1
LLM_API_KEY=sk-...
LLM_MODEL=gpt-4o-mini
```

### 3. Try it

1. Click **New** (or **Open .xlsx** to upload an existing workbook).
2. Type: *"Add Arun, age 24, department AI, salary 35000."*
3. Try a follow-up: *"Set his salary to 40000."*
4. Try a multi-step command: *"Create a Summary sheet, calculate the total
   salary from Sheet1, and put the result in A1."*
5. Try something destructive: *"Delete that row."* — the agent will ask you
   to confirm before it touches anything.
6. **Save** / **Download** when you're done.

## Project structure

```
.
├── backend/
│   ├── app/
│   │   ├── agent/        # LLM factory + LangGraph ReAct agent + system prompt
│   │   ├── excel/        # WorkbookManager: openpyxl + undo snapshots
│   │   ├── tools/        # The only surface the LLM can act through
│   │   ├── api/          # FastAPI routes + websocket status stream
│   │   └── main.py       # App entrypoint, serves frontend + mounts routes
│   └── requirements.txt
├── frontend/
│   ├── index.html
│   ├── style.css
│   └── app.js
├── workbooks/             # Uploaded/created .xlsx files + history.sqlite3 live here
├── tests/
│   └── test_workbook_manager.py   # 26 tests, no LLM required — run with pytest
├── .env.example
└── README.md
```

## API reference

| Method | Path | Purpose |
|---|---|---|
| POST | `/api/workbook/upload` | Upload an existing `.xlsx`/`.xlsm` |
| POST | `/api/workbook/new` | Create a blank workbook |
| GET | `/api/workbook` | Current workbook summary |
| GET | `/api/workbook/sheets` | List sheet names |
| GET | `/api/workbook/preview` | Grid preview of a sheet |
| POST | `/api/workbook/save` | Save to disk |
| GET | `/api/workbook/download` | Download the saved `.xlsx` |
| POST | `/api/agent/message` | Send a text command to the agent |
| POST | `/api/undo` | Revert the last mutating operation |
| GET | `/api/history` | Chat transcript + workbook action log |
| WS | `/ws/agent` | Live status stream + async chat |

## Tool list

Inspection (read-only — the agent calls these liberally): `get_workbook_summary`,
`preview_sheet`, `read_range`, `find_value`, `find_row_by_column_value`,
`get_headers`, `list_sheets`, `analyze_sheet`.

Writes & structure: `write_cell`, `write_range`, `add_row`, `update_row`,
`delete_row`, `add_column`, `rename_column`, `delete_column`,
`insert_column_before`, `insert_column_after`, `insert_row_before`,
`insert_row_after`, `move_column`, `move_row`, `copy_range`, `clear_range`,
`merge_cells`, `unmerge_cells`, `freeze_panes`, `set_column_width`,
`set_row_height`, `autofit_columns`, `duplicate_sheet`, `set_cell_comment`.

Sheets: `create_sheet`, `rename_sheet`, `delete_sheet`, `set_active_sheet`.

Sort/filter: `sort_range`, `filter_rows`. Formatting: `format_cells`,
`conditional_format` (`greater_than`, `less_than`, `between`, `color_scale`,
`data_bars`, `duplicate_values`), `data_validation` (`list` dropdown,
`whole_number`, `decimal`, `date`). Charts: `add_chart` (`bar`/`line`/`pie`).

## Formula-writing behavior

`apply_formula(cell, formula, sheet)` accepts **any valid Excel formula string**
(e.g. `=VLOOKUP(A2,Sheet2!A:C,3,FALSE)`, `=IF(B2>30000,"High","Low")`). It is
validated only for basic syntax sanity (starts with `=`, balanced parentheses),
stored as a **live formula** for Excel to recalculate, and additionally run through
a whitelisted display parser (`backend/app/excel/formula_engine.py` — no
`eval`/`exec`) so previews/reads and the tool result show the **computed answer**
(e.g. `Excellent: 2, Good: 1`), never the raw `=...` text. `apply_simple_formula` remains as a one-shot convenience
wrapper for the common aggregate case (`SUM`/`AVERAGE`/`COUNT`/`MIN`/`MAX`), so
"sum column C" stays a single tool call.

For plain-English calculations ("total sales where region is West and month is
Jan"), the agent uses `suggest_formula(description)`, which returns the sheet's
headers plus a data sample; the agent then composes the correct live formula
itself (`SUMIFS`/`VLOOKUP`/`INDEX-MATCH`/`IF`/`COUNTIFS`/…) with its own
reasoning — no separate ML model — and writes it via `apply_formula`. The system
prompt instructs the agent to prefer such live, recalculating formulas over
precomputing static values whenever the request implies a formula.

## How the "no arbitrary code" guarantee works

The agent's LLM never sees a Python interpreter. Its **entire** action space is
the `@tool`-decorated functions in `app/tools/excel_tools.py`, each a thin,
typed wrapper around one `WorkbookManager` method (`write_cell`, `add_row`,
`sort_range`, `apply_formula`, …). There is no tool that accepts a code string
or shell command. `apply_formula` accepts an Excel formula **string only** —
never executed by Python, only stored as a cell formula for Excel to evaluate.
If a request can't be expressed through these tools, the system
prompt instructs the agent to say so rather than improvise.

Destructive/structural tools (`delete_row`, `delete_column`, `delete_sheet`,
`insert_column_before/after`, `insert_row_before/after`, `move_column/row`,
`merge_cells`, `clear_range` — see `DESTRUCTIVE_TOOLS` in `excel_tools.py`, plus
large-range overwrites beyond `LARGE_RANGE_ROW_THRESHOLD`) require explicit user
confirmation (system-prompt rule 4 + the frontend Yes/No prompt). Every
write/structural tool snapshots before mutating so undo keeps working.

## Roadmap

- **Phase 2 — voice (STT/TTS):** re-attach speech-to-text and text-to-speech
  around the existing `/ws/agent` status-stream contract (`Listening...` /
  `Analyzing...` / `Updating Excel...` / `✓ Completed`). The agent and tool
  layers are voice-ready by design — only the transport (API routes + frontend)
  changes, no rearchitecting needed.
- **More charts / pivot-style summaries** → `workbook_manager.py` + a matching
  tool in `excel_tools.py`.
- **Multi-user sessions** → replace the single global `WorkbookSession` in
  `workbook_manager.py` with a dict keyed by session id.
- **Persistent long-term memory** (beyond one conversation thread) → swap
  LangGraph's `MemorySaver` for `SqliteSaver`/`PostgresSaver`.
- **Real-time push during long multi-step commands** → the `/ws/agent`
  websocket and `broadcast_status()` helper in `routes.py` are already wired
  for this; call `broadcast_status(...)` between tool calls in `agent.py` if
  you want per-step UI updates instead of one final status.
