from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv

# Load .env before any app.* imports that read os.getenv at import time
# (app.db, app.agent.llm_client, app.speech.*).
load_dotenv(Path(__file__).resolve().parent.parent / ".env", override=False)
load_dotenv(Path(__file__).resolve().parent.parent.parent / ".env", override=False)

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse

from app import db
from app.api.routes import router as api_router

app = FastAPI(title="Excela AI", version="0.2.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # MVP: relax for local dev. Tighten before any real deployment.
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(api_router, prefix="/api")


@app.on_event("startup")
async def on_startup() -> None:
    db.init_db()


FRONTEND_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "frontend")

if os.path.isdir(FRONTEND_DIR):
    app.mount("/static", StaticFiles(directory=FRONTEND_DIR), name="static")

    @app.get("/")
    async def index():
        return FileResponse(os.path.join(FRONTEND_DIR, "index.html"))


@app.get("/api/health")
async def health():
    return {"status": "ok"}


@app.get("/api/llm/status")
async def llm_status():
    """Echo resolved LLM config (no secret) + live reachability check."""
    from app.agent.llm_client import get_llm_config

    cfg = get_llm_config()
    safe = {k: v for k, v in cfg.items() if k != "api_key"}
    safe["api_key_configured"] = bool(cfg.get("api_key"))
    # Live ping of the OpenAI-compatible endpoint (short timeout, no exception leak).
    import urllib.request
    import json as _json

    reachable, models = False, []
    try:
        req = urllib.request.Request(
            cfg["base_url"].rstrip("/") + "/models",
            headers={"Authorization": f"Bearer {cfg['api_key']}"},
        )
        with urllib.request.urlopen(req, timeout=5) as resp:
            payload = _json.loads(resp.read().decode("utf-8", "replace"))
            reachable = True
            models = [m.get("id") for m in payload.get("data", [])][:20]
    except Exception as exc:  # endpoint down / wrong URL / auth — report, don't 500
        safe["reachability_error"] = f"{type(exc).__name__}: {exc}"
    safe["reachable"] = reachable
    if models:
        safe["available_models"] = models
        # Ollama Cloud lists "gpt-oss:20b" while the local stub / user config
        # may say "gpt-oss:20b-cloud" — treat those as the same model.
        base = cfg["model"].removesuffix("-cloud")
        safe["configured_model_available"] = cfg["model"] in models or base in models
    return safe
