from __future__ import annotations

import asyncio
import os
import shutil

from fastapi import APIRouter, File, HTTPException, UploadFile, WebSocket, WebSocketDisconnect
from fastapi.responses import FileResponse
from pydantic import BaseModel

from app import db
from app.agent.agent import excel_agent
from app.excel.workbook_manager import WorkbookError, workbook_manager as wm

router = APIRouter()

UPLOAD_DIR = "workbooks"
os.makedirs(UPLOAD_DIR, exist_ok=True)


# ------------------------------------------------------------------ #
# Workbook endpoints
# ------------------------------------------------------------------ #
@router.post("/workbook/upload")
async def upload_workbook(file: UploadFile = File(...)):
    filename = os.path.basename(file.filename or "")
    if not filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(400, "Only .xlsx/.xlsm files are supported")
    if not filename or filename in (".", ".."):
        raise HTTPException(400, "Invalid filename")
    dest = os.path.join(UPLOAD_DIR, filename)
    with open(dest, "wb") as f:
        shutil.copyfileobj(file.file, f)
    try:
        summary = wm.open_workbook(dest)
    except WorkbookError as exc:
        if os.path.exists(dest):
            os.remove(dest)  # don't keep files that aren't valid workbooks
        raise HTTPException(400, str(exc)) from exc
    return summary


@router.post("/workbook/new")
async def new_workbook(filename: str = "untitled.xlsx"):
    return wm.new_workbook(filename)


@router.get("/workbook")
async def get_workbook():
    if not wm.is_loaded():
        return {"loaded": False}
    return {"loaded": True, **wm.get_summary()}


@router.get("/workbook/sheets")
async def get_sheets():
    try:
        return {"sheets": wm.list_sheets()}
    except WorkbookError as exc:
        raise HTTPException(400, str(exc)) from exc


@router.get("/workbook/preview")
async def preview(sheet: str | None = None, max_rows: int = 25, max_cols: int = 12, start_row: int = 1):
    try:
        return wm.sheet_preview(sheet=sheet, max_rows=max_rows, max_cols=max_cols, start_row=start_row)
    except WorkbookError as exc:
        raise HTTPException(400, str(exc)) from exc


@router.post("/workbook/save")
async def save_workbook(save_as: str | None = None):
    try:
        return wm.save_workbook(save_as=save_as)
    except WorkbookError as exc:
        raise HTTPException(400, str(exc)) from exc


@router.get("/workbook/download")
async def download_workbook():
    if not wm.session.path or not os.path.exists(wm.session.path):
        raise HTTPException(404, "No saved workbook file found. Save first.")
    return FileResponse(wm.session.path, filename=os.path.basename(wm.session.path))


@router.post("/undo")
async def undo():
    try:
        return wm.undo()
    except WorkbookError as exc:
        raise HTTPException(400, str(exc)) from exc


@router.get("/history")
async def history(session_id: str = "default", limit: int = 100):
    return {"chat": db.get_history(session_id, limit), "workbook_actions": wm.get_history(limit)}


# ------------------------------------------------------------------ #
# Agent endpoints
# ------------------------------------------------------------------ #
class AgentMessageRequest(BaseModel):
    message: str
    session_id: str = "default"
    confirm_destructive: bool = False


class AgentMessageResponse(BaseModel):
    reply: str
    needs_confirmation: bool
    tool_calls: list[dict]


@router.post("/agent/message", response_model=AgentMessageResponse)
async def agent_message(req: AgentMessageRequest):
    if not wm.is_loaded():
        raise HTTPException(400, "No workbook is open. Upload or create one first.")

    db.add_message(req.session_id, "user", req.message)
    result = excel_agent.run(req.message, session_id=req.session_id, confirm_destructive=req.confirm_destructive)
    db.add_message(req.session_id, "assistant", result["reply"])

    return AgentMessageResponse(
        reply=result["reply"],
        needs_confirmation=result["needs_confirmation"],
        tool_calls=result["tool_calls"],
    )


# ------------------------------------------------------------------ #
# WebSocket — live status streaming for text chat (voice-ready contract)
# ------------------------------------------------------------------ #
class ConnectionManager:
    def __init__(self) -> None:
        self.active: list[WebSocket] = []

    async def connect(self, ws: WebSocket) -> None:
        await ws.accept()
        self.active.append(ws)

    def disconnect(self, ws: WebSocket) -> None:
        if ws in self.active:
            self.active.remove(ws)

    async def broadcast(self, message: dict) -> None:
        for ws in list(self.active):
            try:
                await ws.send_json(message)
            except Exception:
                self.disconnect(ws)


manager = ConnectionManager()

# Tools that never mutate the workbook — anything else counts as an edit,
# so the WS stream can emit a truthful "Updating Excel..." phase.
READ_ONLY_TOOLS = frozenset({
    "get_workbook_summary",
    "preview_sheet",
    "read_range",
    "find_value",
    "find_row_by_column_value",
    "get_headers",
    "list_sheets",
    "analyze_sheet",
    "suggest_formula",
})


@router.websocket("/ws/agent")
async def ws_agent(websocket: WebSocket):
    """
    Live status stream + async chat over a single socket.

    Status contract (shared with the planned Phase 2 voice transport):
    "Thinking..." / "Figuring out the edits..." / "Reading workbook..." /
    "Editing Excel..." / "Verifying..." / "Updating Excel..." / "✓ Completed".
    Text chat reuses this contract today: the frontend sends
    {"message", "session_id", "confirm_destructive"} and renders each
    {"status", ...} frame as live UI status + typing indicator text.
    A future voice transport can attach STT/TTS around the same frames
    without rearchitecting the agent or tool layer.
    """
    await manager.connect(websocket)
    try:
        while True:
            data = await websocket.receive_json()
            message = data.get("message", "")
            session_id = data.get("session_id", "default")
            confirm = data.get("confirm_destructive", False)

            if not message:
                continue
            if not wm.is_loaded():
                await websocket.send_json({"status": "error", "reply": "No workbook is open."})
                continue

            await websocket.send_json({"status": "Thinking…"})
            db.add_message(session_id, "user", message)

            # Run the blocking LangGraph ReAct loop in a worker thread so
            # the event loop stays free to stream heartbeat statuses.
            # The agent can take 10-60s (multiple LLM + tool round-trips),
            # so we cycle through human-readable phases while it works.
            phases = [
                "Thinking…",
                "Figuring out the edits…",
                "Reading workbook…",
                "Editing Excel…",
                "Verifying…",
            ]
            stop_heartbeat = asyncio.Event()

            async def _heartbeat() -> None:
                i = 1  # index 0 ("Thinking…") was just sent above
                try:
                    while not stop_heartbeat.is_set():
                        await asyncio.sleep(2.2)
                        if stop_heartbeat.is_set():
                            break
                        await websocket.send_json({"status": phases[i % len(phases)]})
                        i += 1
                except Exception:
                    pass

            heartbeat_task = asyncio.create_task(_heartbeat())
            try:
                result = await asyncio.to_thread(
                    excel_agent.run, message,
                    session_id=session_id, confirm_destructive=confirm,
                )
            except Exception as exc:
                stop_heartbeat.set()
                await heartbeat_task
                await websocket.send_json({"status": "error", "reply": f"Agent error: {exc}"})
                continue
            finally:
                stop_heartbeat.set()
            await heartbeat_task

            db.add_message(session_id, "assistant", result["reply"])
            tool_names = {tc.get("tool", "") for tc in result.get("tool_calls", [])}
            did_write = bool(tool_names - READ_ONLY_TOOLS)
            if did_write:
                await websocket.send_json({"status": "Updating Excel..."})
                await asyncio.sleep(0.35)  # let the UI paint the edit phase
            await websocket.send_json(
                {
                    "status": "✓ Completed",
                    "reply": result["reply"],
                    "tool_calls": result["tool_calls"],
                    "needs_confirmation": result["needs_confirmation"],
                }
            )
    except WebSocketDisconnect:
        manager.disconnect(websocket)


async def broadcast_status(status: str) -> None:
    await manager.broadcast({"status": status})
