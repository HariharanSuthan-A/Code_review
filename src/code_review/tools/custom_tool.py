"""Custom tools used by the code review crew."""

from __future__ import annotations

import os
from pathlib import Path

from crewai.tools import tool
from dotenv import load_dotenv
from tavily import TavilyClient

# Load variables from the nearest .env so the API key is not hardcoded.
load_dotenv()

MAX_RESULTS = 10
MAX_FILE_CHARS = 50_000
MAX_LIST_ENTRIES = 500
MAX_TREE_DEPTH = 3

# Directories that are noise for a code review and can contain thousands of files.
IGNORED_DIRS = {
    ".git",
    ".hg",
    ".svn",
    ".venv",
    "venv",
    "env",
    "node_modules",
    "__pycache__",
    ".pytest_cache",
    ".mypy_cache",
    ".ruff_cache",
    ".idea",
    ".vscode",
    "dist",
    "build",
    ".next",
    ".nuxt",
    ".turbo",
    "coverage",
    "htmlcov",
}

OFFICIAL_DOMAINS = [
    "docs.python.org",
    "nodejs.org",
    "developer.mozilla.org",
    "typescriptlang.org",
    "react.dev",
    "angular.dev",
    "vuejs.org",
    "svelte.dev",
    "nextjs.org",
    "nuxt.com",
    "expressjs.com",
    "nestjs.com",
    "fastapi.tiangolo.com",
    "djangoproject.com",
    "flask.palletsprojects.com",
    "spring.io",
    "learn.microsoft.com",
    "go.dev",
    "php.net",
    "laravel.com",
    "rubyonrails.org",
    "rust-lang.org",
    "kotlinlang.org",
    "oracle.com",
    "dev.java",
    "learn-c.org",
    "isocpp.org",
    "ruby-lang.org",
    "perl.org",
    "swift.org",
    "developer.android.com",
    "developer.apple.com",
    "flutter.dev",
    "reactnative.dev",
    "getbootstrap.com",
    "tailwindcss.com",
    "mui.com",
    "ant.design",
    "webpack.js.org",
    "vite.dev",
    "babeljs.io",
    "eslint.org",
    "prettier.io",
    "git-scm.com",
    "github.com",
    "docker.com",
    "kubernetes.io",
    "jenkins.io",
    "postgresql.org",
    "dev.mysql.com",
    "mongodb.com",
    "redis.io",
    "firebase.google.com",
    "supabase.com",
    "aws.amazon.com",
    "azure.microsoft.com",
    "cloud.google.com",
    "graphql.org",
    "swagger.io",
    "spec.openapis.org",
    "jestjs.io",
    "vitest.dev",
    "testing-library.com",
    "playwright.dev",
    "selenium.dev",
    "postman.com",
    "owasp.org",
    "12factor.net",
    "roadmap.sh",
]

_tavily_client: TavilyClient | None = None


def _get_api_key() -> str:
    """Return the Tavily key from the environment, accepting either casing."""
    for variable in ("TAVILY_API_KEY", "tavily_api_key"):
        value = (os.getenv(variable) or "").strip().strip("\"'")
        if value:
            return value
    return ""


def _get_tavily_client() -> TavilyClient:
    """Create the Tavily client lazily so importing this module never fails."""
    global _tavily_client
    if _tavily_client is None:
        api_key = _get_api_key()
        if not api_key:
            raise RuntimeError(
                "Tavily API key not found. Set TAVILY_API_KEY in the environment "
                "or in the project's .env file."
            )
        _tavily_client = TavilyClient(api_key=api_key)
    return _tavily_client


def _resolve_path(raw_path: str | None) -> Path:
    """Resolve a user-supplied path against the current working directory."""
    text = str(raw_path).strip() if raw_path is not None else ""
    if text in ("", "."):
        return Path.cwd().resolve()

    path = Path(text).expanduser()
    if not path.is_absolute():
        path = Path.cwd() / path
    return path.resolve()


def _human_size(size: int) -> str:
    if size < 1024:
        return f"{size} B"
    for unit in ("KB", "MB", "GB"):
        size /= 1024.0
        if size < 1024.0:
            return f"{size:.1f} {unit}"
    return f"{size:.1f} TB"


def _format_entry(path: Path) -> str:
    try:
        if path.is_symlink() and not path.exists():
            return f"[BROKEN LINK] {path.name}"
        if path.is_dir():
            return f"[DIR]  {path.name}/"
        return f"[FILE] {path.name} ({_human_size(path.stat().st_size)})"
    except OSError:
        return f"[?]    {path.name}"


def _sorted_entries(path: Path) -> list[Path]:
    return sorted(path.iterdir(), key=lambda item: (not item.is_dir(), item.name.lower()))


def _add_tree(
    directory: Path,
    prefix: str,
    depth: int,
    lines: list[str],
    counter: list[int],
) -> None:
    if depth > MAX_TREE_DEPTH or counter[0] >= MAX_LIST_ENTRIES:
        return

    try:
        entries = _sorted_entries(directory)
    except OSError as exc:
        lines.append(f"{prefix}[cannot list: {exc}]")
        return

    for index, entry in enumerate(entries):
        if counter[0] >= MAX_LIST_ENTRIES:
            lines.append(f"{prefix}... (truncated at {MAX_LIST_ENTRIES} entries)")
            return

        counter[0] += 1
        is_last = index == len(entries) - 1
        connector = "└── " if is_last else "├── "
        suffix = "/" if entry.is_dir() else ""
        lines.append(f"{prefix}{connector}{entry.name}{suffix}")

        if (
            entry.is_dir()
            and not entry.is_symlink()
            and entry.name not in IGNORED_DIRS
        ):
            extension = "    " if is_last else "│   "
            _add_tree(entry, prefix + extension, depth + 1, lines, counter)


@tool("tavily_search")
def tavily_search(query: str) -> str:
    """Search official documentation sites for programming languages and frameworks.

    Use this tool to gather authoritative technical and security documentation
    from vendor and project websites (for example Python, Node.js, MDN,
    framework docs, and OWASP).

    Args:
        query: The search query, for example "Django SQL injection prevention".

    Returns:
        A formatted list of matching pages with titles, URLs, and excerpts.
    """
    query = (query or "").strip()
    if not query:
        return "Error: 'query' must be a non-empty string."

    try:
        response = _get_tavily_client().search(
            query=query,
            max_results=MAX_RESULTS,
            search_depth="advanced",
            include_answer=True,
            include_domains=OFFICIAL_DOMAINS,
        )
    except Exception as exc:  # noqa: BLE001 - surface the failure to the agent
        return f"Tavily search failed: {type(exc).__name__}: {exc}"

    lines: list[str] = []
    answer = (response.get("answer") or "").strip()
    if answer:
        lines.append(f"Answer: {answer}")

    results = response.get("results") or []
    if not results:
        lines.append(f"No results found for: {query}")
        return "\n\n".join(lines)

    for index, item in enumerate(results, start=1):
        title = (item.get("title") or "Untitled").strip()
        url = (item.get("url") or "").strip()
        content = (item.get("content") or "").strip()
        if len(content) > 800:
            content = content[:800].rsplit(" ", 1)[0] + "..."
        lines.append(f"{index}. {title}\n   URL: {url}\n   {content}")

    return "\n\n".join(lines)


@tool("list_files_folders")
def list_files_folders(dir: str = ".") -> str:
    """List the files and folders directly inside a directory.

    Use this tool first to discover the top-level structure of a project.

    Args:
        dir: Directory to inspect. Relative paths are resolved from the current
            working directory. Defaults to "." (the current directory).

    Returns:
        One line per entry, marked as [DIR] or [FILE] with its size.
    """
    path = _resolve_path(dir)
    if not path.exists():
        return f"Error: directory does not exist: {path}"
    if not path.is_dir():
        return f"Error: {path} is a file, not a directory. Use read_files to read it."

    try:
        entries = _sorted_entries(path)
    except OSError as exc:
        return f"Error: cannot list {path}: {exc}"

    if not entries:
        return f"{path} is empty."

    lines = [f"Contents of {path} ({len(entries)} entries):"]
    lines.extend(_format_entry(entry) for entry in entries[:MAX_LIST_ENTRIES])
    if len(entries) > MAX_LIST_ENTRIES:
        lines.append(f"... (truncated at {MAX_LIST_ENTRIES} entries)")
    return "\n".join(lines)


@tool("open_folders")
def open_folders(folder_name: str = ".") -> str:
    """Show the contents of a folder as a tree, up to three levels deep.

    Use this tool to explore a project folder after list_files_folders has
    shown which folders exist. VCS, virtualenv, dependency, and cache folders
    are skipped automatically.

    Args:
        folder_name: Folder to open, for example "voice-excel-agent" or
            "src/code_review". Defaults to "." (the current directory).

    Returns:
        A tree view of the folder's files and subfolders.
    """
    path = _resolve_path(folder_name)
    if not path.exists():
        return f"Error: folder does not exist: {path}"
    if not path.is_dir():
        return f"Error: {path} is a file, not a folder. Use read_files to read it."

    lines = [f"Tree of {path} (max depth {MAX_TREE_DEPTH}, ignored: {', '.join(sorted(IGNORED_DIRS))}):"]
    counter = [0]
    _add_tree(path, "", 1, lines, counter)
    if len(lines) == 1:
        lines.append("(empty)")
    return "\n".join(lines)


@tool("read_files")
def read_files(file_name: str, max_chars: int = MAX_FILE_CHARS) -> str:
    """Read the text contents of a file.

    Use this tool to inspect source code, configuration, and documentation
    files discovered with list_files_folders or open_folders.

    Args:
        file_name: Path to the file to read. Relative paths are resolved from
            the current working directory.
        max_chars: Maximum number of characters to return. Large files are
            truncated. Defaults to 50000 and is capped at 200000.

    Returns:
        The file contents, or a clear error message if it cannot be read.
    """
    path = _resolve_path(file_name)
    if not path.exists():
        return f"Error: file does not exist: {path}"
    if path.is_dir():
        return f"Error: {path} is a directory. Use open_folders to list it."

    try:
        data = path.read_bytes()
    except OSError as exc:
        return f"Error: cannot read {path}: {exc}"

    if b"\x00" in data[:2048]:
        return f"Error: {path} appears to be a binary file ({len(data)} bytes); it cannot be read as text."

    try:
        text = data.decode("utf-8")
    except UnicodeDecodeError:
        text = data.decode("latin-1")

    try:
        limit = max(1, min(int(max_chars), 200_000))
    except (TypeError, ValueError):
        limit = MAX_FILE_CHARS

    original_length = len(text)
    truncated = original_length > limit
    if truncated:
        text = text[:limit]

    result = f"Contents of {path}:\n\n{text}"
    if truncated:
        result += f"\n\n... [truncated after {limit} of {original_length} characters]"
    return result


available_tools = [tavily_search, list_files_folders, open_folders, read_files]
