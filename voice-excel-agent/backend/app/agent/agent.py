"""
Excela AI — ExcelAgent
======================
Wraps a LangGraph ReAct agent around the Excel tool set, with:
  * A system prompt that enforces "inspect before you act".
  * A confirmation gate that intercepts destructive/structural tool calls
    (deletes, merges, clears, insert/move operations that shift data) and
    large writes, pausing the run and asking the user to confirm before
    the tool actually executes.
  * Lightweight conversational memory so follow-ups like
    "set his experience to 2 years" resolve against the previous turn
    (who "his" refers to, which sheet, etc.) via LangGraph's checkpointer.

The agent NEVER executes arbitrary code. It can only call the functions
defined in app.tools.excel_tools. Excel formula strings are evaluated in the
backend (whitelisted parser, never exec/eval) and only the ANSWER is stored
in the user-specified cell or reported in chat — '=...' text is never pasted.

Transport note (voice-ready): Excela AI is text-only (chat) for now.
Voice STT/TTS is a planned Phase 2 addition that will re-attach via the
/ws/agent status-stream contract. The agent and tool layers make no
text-only assumptions — only the transport changes.
"""
from __future__ import annotations

from typing import Any

from langchain_core.messages import AIMessage, HumanMessage, SystemMessage, ToolMessage
from langgraph.checkpoint.memory import MemorySaver
from langgraph.prebuilt import create_react_agent

from app.agent.llm_client import get_llm
from app.tools.excel_tools import ALL_TOOLS, DESTRUCTIVE_TOOLS

SYSTEM_PROMPT = """You are Excela AI, an AI assistant that controls a Microsoft Excel workbook on the \
user's behalf through a fixed set of tools. You talk to the user via text chat, \
so keep your responses short, natural, and confirmatory — like a helpful assistant, \
not a technical log.

Rules you must always follow:
1. Never invent cell coordinates, column names, or sheet names. Before writing or reading \
    data, call an inspection tool (get_workbook_summary, preview_sheet, get_headers, \
    find_value, find_row_by_column_value) to ground yourself in the ACTUAL current state \
    of the workbook.
2. When the user refers to a record by name or by a pronoun ("his", "her", "that row"), \
    resolve it against the conversation history or by calling find_row_by_column_value — \
    do not guess a row number.
3. For multi-step requests (e.g. "create a sheet, compute a total, and put it in A1"), \
    perform the steps in order, calling tools one at a time, and re-inspect the workbook \
    between steps if a later step depends on the result of an earlier one.
4. Destructive/structural actions (deleting a row, column, or sheet; merging cells; \
    clearing a range — especially a large one; inserting/moving rows or columns, which \
    shifts existing data; or overwriting a large existing range) require explicit user \
    confirmation. If the user has not already confirmed in this message, do NOT call the \
    destructive tool — instead ask a short, specific confirmation question (e.g. "Delete \
    row 5 (Arun, AI, 35000)? Say yes to confirm.") and stop there.
5. After completing an action, verify it succeeded (e.g. re-read the cell/row you just \
    wrote) before telling the user it's done. If a tool call returns success: false, \
    explain the problem in plain language and suggest a fix — do not silently retry forever.
6. Keep your final response to the user brief and conversational, e.g. "Arun was added \
    successfully." or "Total salary is 145000, added to Summary!A1." Report the computed \
    answer from the tool result ('value') — never paste the raw '=...' formula text into \
    chat. Do not describe your internal tool-calling steps or show raw JSON to the user.
7. You can only use the provided tools. You cannot run arbitrary Python or shell commands. \
    Excel formula strings are NEVER pasted into cells: compute_formula evaluates a formula \
    in the backend (no exec/eval) and returns the answer for chat; apply_formula computes \
    in the backend and stores only the ANSWER in the user-specified cell (a request like \
    "column C row 5" means cell C5 — resolve "column X row Y" to the XY coordinate first).
8. For calculations: if the user names a destination cell ("put the total in C5" / \
    "column C row 5"), compute with apply_formula/apply_simple_formula so the ANSWER lands \
    in that cell and report the value in chat. If no destination is named ("what is the \
    total?"), use compute_formula and just report the answer in chat. When the \
    user describes a calculation in plain English ("total sales where region is West and \
    month is Jan"), first call suggest_formula and/or the inspection tools (get_headers, \
    preview_sheet) to ground yourself in the real headers/data, then construct the correct \
    formula yourself (SUMIFS/VLOOKUP/XLOOKUP/INDEX-MATCH/IF/COUNTIFS/AVERAGEIFS/...) and \
    evaluate it via compute_formula/apply_formula. Use apply_simple_formula only for the \
    trivial one-shot aggregate case ("sum column C").
9. Row 1 of every sheet is the column-names (header) row. It is pinned on every preview \
    page (25 rows/page) — preview_sheet always returns it in `headers`/`header_row` — so \
    always reference columns by those exact header names.
"""


class PendingConfirmation(Exception):
    """Raised (as a control-flow signal) when the agent needs user confirmation
    before proceeding with a destructive tool call."""

    def __init__(self, tool_name: str, tool_args: dict, message: str) -> None:
        self.tool_name = tool_name
        self.tool_args = tool_args
        self.message = message
        super().__init__(message)


class ExcelAgent:
    def __init__(self) -> None:
        self.llm = get_llm()
        self.checkpointer = MemorySaver()
        self.graph = create_react_agent(
            self.llm,
            tools=ALL_TOOLS,
            checkpointer=self.checkpointer,
        )

    def run(self, user_message: str, session_id: str = "default", confirm_destructive: bool = False) -> dict[str, Any]:
        """
        Run one turn of the agent.

        confirm_destructive: when True, the caller (frontend, after the user said
        "yes") is re-issuing a message that should be allowed to hit destructive
        tools. We pass this through as a light hint in the message itself so the
        model's own confirmation logic (rule 4 in the system prompt) relaxes.
        """
        config = {"configurable": {"thread_id": session_id}}

        prefix = "[User has confirmed the pending destructive action.] " if confirm_destructive else ""
        messages = [SystemMessage(content=SYSTEM_PROMPT), HumanMessage(content=prefix + user_message)]

        # On the very first turn for a thread, include the system prompt; on
        # later turns the checkpointer already has it, so LangGraph will just
        # append. create_react_agent handles this de-duplication internally
        # when the same thread_id is reused.
        result = self.graph.invoke({"messages": messages}, config=config)

        final_messages = result["messages"]
        tool_calls_made = [
            {"tool": m.name, "content": m.content}
            for m in final_messages
            if isinstance(m, ToolMessage)
        ]
        reply = ""
        for m in reversed(final_messages):
            if isinstance(m, AIMessage) and m.content:
                reply = m.content if isinstance(m.content, str) else str(m.content)
                break

        needs_confirmation = self._detect_confirmation_request(reply, tool_calls_made)

        return {
            "reply": reply,
            "tool_calls": tool_calls_made,
            "needs_confirmation": needs_confirmation,
        }

    @staticmethod
    def _detect_confirmation_request(reply: str, tool_calls_made: list[dict]) -> bool:
        """Heuristic: if none of the destructive tools were actually called but the
        reply text looks like a confirmation question, flag it so the frontend can
        show a Yes/No prompt instead of treating this as a completed action."""
        called_names = {tc["tool"] for tc in tool_calls_made}
        if called_names & DESTRUCTIVE_TOOLS:
            return False
        reply_l = reply.lower()
        markers = ["confirm", "are you sure", "delete", "say yes", "shall i proceed", "do you want me to"]
        return any(marker in reply_l for marker in markers) and "?" in reply


excel_agent = ExcelAgent()
