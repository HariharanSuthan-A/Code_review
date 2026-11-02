"""
Multi-provider LLM factory.

- provider "ollama" -> native ChatOllama client against Ollama's /api/chat
  endpoint (far more reliable for tool calling with local GGUF models than
  the OpenAI-compat /v1 shim, which 500s on union-type tool schemas).
- provider "openai" | "openai_compatible" -> ChatOpenAI against any
  OpenAI-compatible chat-completions endpoint (OpenAI, OpenRouter, vLLM,
  LM Studio, ...).

Env vars (see .env.example):
    LLM_PROVIDER        "ollama" | "openai" | "openai_compatible"  (default: ollama)
    LLM_BASE_URL         base URL of the endpoint
    LLM_API_KEY           API key (Ollama ignores this; pass any non-empty string)
    LLM_MODEL             model name, e.g. "llama3.1", "qwen2.5:14b", "gpt-4o-mini"
    LLM_TEMPERATURE        default 0.1 (keep it low — this agent calls tools, not prose)
"""
from __future__ import annotations

import logging
import os
from pathlib import Path

from dotenv import load_dotenv

# Load backend/.env (and fall back to repo-root .env / environment).
# This module is imported before FastAPI startup, so we cannot rely on
# main.py to have called load_dotenv() first — do it here explicitly.
for _candidate in (
    Path(__file__).resolve().parents[2] / ".env",  # backend/.env
    Path(__file__).resolve().parents[3] / ".env",  # repo-root .env
):
    if _candidate.is_file():
        load_dotenv(_candidate, override=False)
        break
else:
    load_dotenv(override=False)

logger = logging.getLogger(__name__)

DEFAULT_BASE_URLS = {
    "ollama": "http://localhost:11434/v1",
    "openai": "https://api.openai.com/v1",
}


def get_llm_config() -> dict:
    """Resolve LLM settings from the environment (.env). Single source of truth."""
    provider = os.getenv("LLM_PROVIDER", "ollama").strip().lower() or "ollama"
    base_url = os.getenv("LLM_BASE_URL") or DEFAULT_BASE_URLS.get(provider, DEFAULT_BASE_URLS["ollama"])
    api_key = os.getenv("LLM_API_KEY", "ollama-local-key") or "ollama-local-key"
    model = os.getenv("LLM_MODEL", "llama3.1").strip() or "llama3.1"
    try:
        temperature = float(os.getenv("LLM_TEMPERATURE", "0.1"))
    except ValueError:
        temperature = 0.1
    return {
        "provider": provider,
        "base_url": base_url.rstrip("/"),
        "api_key": api_key,
        "model": model,
        "temperature": temperature,
    }


def get_llm():
    cfg = get_llm_config()
    logger.info(
        "LLM config: provider=%s model=%s base_url=%s temperature=%s",
        cfg["provider"],
        cfg["model"],
        cfg["base_url"],
        cfg["temperature"],
    )
    if cfg["provider"] == "ollama":
        # Native Ollama API: strip the /v1 compat suffix, ChatOllama
        # expects the server root (http://host:11434).
        from langchain_ollama import ChatOllama

        native_base = cfg["base_url"]
        if native_base.endswith("/v1"):
            native_base = native_base[: -len("/v1")]
        return ChatOllama(
            model=cfg["model"],
            base_url=native_base,
            temperature=cfg["temperature"],
        )
    from langchain_openai import ChatOpenAI

    return ChatOpenAI(
        model=cfg["model"],
        base_url=cfg["base_url"],
        api_key=cfg["api_key"],
        temperature=cfg["temperature"],
        timeout=120,
    )
