"""
Lightweight SQLite store for chat history.

This is intentionally simple: one table, append-only, used to (a) render
the chat panel on reload and (b) give /api/history something durable to
return. The agent's own short-term conversational memory (for resolving
"his"/"that row") is handled separately by LangGraph's MemorySaver
checkpointer in app.agent.agent — this DB is for the human-facing
transcript, not the agent's reasoning state.
"""
from __future__ import annotations

import os
import sqlite3
import time
from contextlib import contextmanager

DB_PATH = os.getenv("HISTORY_DB_PATH", os.path.join("workbooks", "history.sqlite3"))


def _ensure_dir() -> None:
    d = os.path.dirname(DB_PATH)
    if d:
        os.makedirs(d, exist_ok=True)


@contextmanager
def _conn():
    _ensure_dir()
    conn = sqlite3.connect(DB_PATH)
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db() -> None:
    with _conn() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS chat_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                session_id TEXT NOT NULL,
                role TEXT NOT NULL,           -- 'user' | 'assistant'
                content TEXT NOT NULL,
                created_at REAL NOT NULL
            )
            """
        )


def add_message(session_id: str, role: str, content: str) -> None:
    with _conn() as conn:
        conn.execute(
            "INSERT INTO chat_history (session_id, role, content, created_at) VALUES (?, ?, ?, ?)",
            (session_id, role, content, time.time()),
        )


def get_history(session_id: str = "default", limit: int = 100) -> list[dict]:
    with _conn() as conn:
        rows = conn.execute(
            "SELECT role, content, created_at FROM chat_history WHERE session_id = ? ORDER BY id DESC LIMIT ?",
            (session_id, limit),
        ).fetchall()
    return [{"role": r[0], "content": r[1], "created_at": r[2]} for r in reversed(rows)]
