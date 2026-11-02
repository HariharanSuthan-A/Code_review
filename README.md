# CodeReview Crew

A multi-agent [crewAI](https://crewai.com) project for reviewing a codebase end
to end.

## One goal

> Given a folder of source code, understand its technology stack, verify
> practices against authoritative documentation, identify evidence-based
> security weaknesses, and produce an actionable remediation prompt.

That single goal is split across **4 agents** and **4 sequential tasks**. Each
agent has a narrow responsibility, and each task feeds its output to the next.

## Crew structure

| Step | Agent | Task | Consumes | Produces |
|------|-------|------|----------|----------|
| 1 | `code_review` | `code_review_task` | Target folder contents | `output/review.md` |
| 2 | `code_structure_info` | `code_structure_info_task` | Task 1 output | `structure.json` |
| 3 | `code_vulnerability_check` | `code_vulnerability_check_task` | Tasks 1–2 | `vulnerabilities.json` |
| 4 | `write_prompt` | `write_prompt_task` | Tasks 1–3 | `prompt.md` |

```text
code_review_task
      │
      ▼
code_structure_info_task
      │
      ▼
code_vulnerability_check_task
      │
      ▼
write_prompt_task
```

### What each agent does

1. **`code_review`** — reviews the supplied code line by line. Identifies
   languages, frameworks, libraries, dependencies, architecture, coding
   patterns, quality issues, and suspicious areas that warrant security
   investigation. Evidence-based; it does not assume a vulnerability exists
   without support from the code.

2. **`code_structure_info`** — takes the stack identified in step 1 and
   researches official documentation. Summarizes recommended project structure,
   language and framework conventions, secure implementation patterns, common
   pitfalls, and relevant security guidance, with sources.

3. **`code_vulnerability_check`** — performs a systematic security assessment
   using the code review and the researched guidance. Every finding includes an
   identifier, title, category, affected location, evidence, root cause, impact,
   severity, confidence, and remediation. Confirmed vulnerabilities are
   separated from risks that need verification.

4. **`write_prompt`** — converts the findings into a prioritized,
   implementation-ready remediation prompt for a developer or coding agent,
   including fixes, secure code examples, tests, and acceptance criteria.

The areas each agent covers are defined in
`src/code_review/config/agents.yaml`, and the tasks, expected outputs, and
context dependencies are in `src/code_review/config/tasks.yaml`.

## Requirements

- Python >= 3.10, < 3.14
- [uv](https://docs.astral.sh/uv/) for dependency management
- An Ollama instance with the configured model (see
  [Model configuration](#model-configuration))
- A [Tavily](https://tavily.com/) API key for documentation search

## Setup

Install dependencies from the project root:

```bash
uv sync
```

Add your Tavily key to `.env`:

```dotenv
TAVILY_API_KEY=tvly-your-key-here
```

Either `TAVILY_API_KEY` or `tavily_api_key` is accepted. The key is read from
the environment at runtime; nothing is hardcoded in the source.

## Running the crew

Set the folder to review in `src/code_review/main.py`:

```python
folder = "./voice-excel-agent"
user = f"""
Inspect this {folder} and find the improvements and risks and give the prompt to overcome it
"""
```

Then run from the project root:

```bash
crewai run
```

or:

```bash
uv run code_review
```

The crew runs the four tasks in order and writes the output files listed in the
table above.

## Model configuration

The model is configured in `src/code_review/crew.py`:

```python
llm = LLM(
    model="ollama/gpt-oss:20b-cloud",
    base_url="http://localhost:11434",
)
```

Change `model` and `base_url` there to use a different Ollama model or another
provider supported by `crewai.LLM` (for example `"openai/gpt-4o"` with
`OPENAI_API_KEY` set).

## Tools

All four agents share the tools exported from
`src/code_review/tools/custom_tool.py`. Each tool returns a string and reports
failures as readable messages the agent can react to.

| Tool | Purpose | Notes |
|------|---------|-------|
| `tavily_search(query)` | Search official documentation sites | Advanced search, top 10 results, domain allowlist, returns an answer plus excerpts |
| `list_files_folders(dir=".")` | List a directory's immediate contents | Marks `[DIR]`/`[FILE]` and includes file sizes |
| `open_folders(folder_name=".")` | Show a folder as a tree | Three levels deep; skips `.git`, `.venv`, `node_modules`, `__pycache__`, and similar; capped at 500 entries |
| `read_files(file_name, max_chars=50000)` | Read a text file | Detects binaries, falls back from UTF-8 to latin-1, truncates large files; `max_chars` capped at 200,000 |

Paths resolve from the current working directory at call time. The Tavily
client is created lazily, so the module imports cleanly even without a key.

### Documentation allowlist

`tavily_search` only queries official and authoritative sources (language and
framework documentation, MDN, OWASP, vendor docs, and similar). To add or
remove domains, edit `OFFICIAL_DOMAINS` in `custom_tool.py`.

## Project structure

```text
code_review/
├── src/code_review/
│   ├── config/
│   │   ├── agents.yaml        # 4 agents: roles, goals, backstories
│   │   └── tasks.yaml         # 4 tasks: descriptions, outputs, context
│   ├── tools/
│   │   └── custom_tool.py     # Filesystem and Tavily tools
│   ├── crew.py                # Crew wiring: agents, tasks, LLM, outputs
│   └── main.py                # Entry point and target folder
├── knowledge/
├── output/                    # Review report
├── .env                       # API keys (not committed)
└── pyproject.toml
```

## Customizing

- `config/agents.yaml` — change what each agent focuses on
- `config/tasks.yaml` — change task instructions, expected outputs, or context
  dependencies
- `crew.py` — change the model, tool assignment, output paths, or process
- `main.py` — change the folder or prompt passed to the crew
- `tools/custom_tool.py` — change tool behavior, limits, or the documentation
  allowlist

## Troubleshooting

| Symptom | Cause and fix |
|---------|---------------|
| `Tavily API key not found` | Add `TAVILY_API_KEY` to `.env` or export it in the shell |
| `Tavily search failed: ...` | Network issue, invalid key, or rate limit; the message includes the underlying error |
| `Error: directory does not exist` | Paths resolve from the current working directory; pass a valid relative or absolute path |
| `appears to be a binary file` | `read_files` only reads text; use appropriate tooling for archives, spreadsheets, or images |
| Ollama connection error | Start Ollama and confirm the model named in `crew.py` is available |

## Security note

The Tavily key is read from the environment. If a key was ever committed or
shared, rotate it in the Tavily dashboard, and keep `.env` out of version
control.

The filesystem tools honor the paths they are given, including absolute paths.
Run the crew only against code you intend to review, and review its outputs
before acting on them.

## Support

- [crewAI documentation](https://docs.crewai.com)
- [crewAI GitHub repository](https://github.com/joaomdmoura/crewai)
- [crewAI Discord](https://discord.com/invite/X4JWnZnxPb)
- [Chat with the docs](https://chatg.pt/DWjSBZn)
