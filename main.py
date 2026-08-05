"""FastAPI service for the Indian Logistics Data Agent."""

from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import Any

from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field

from agents.master_agent import MasterAgent
from agents.pipeline_agent import DEFAULT_MODEL
from template_schema import TEMPLATE_COLUMNS

load_dotenv()

logging.basicConfig(
    level=os.getenv("LOG_LEVEL", "INFO").upper(),
    format="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
)
logger = logging.getLogger("logistics_data_agent")

app = FastAPI(
    title="India Logistics Data Agent",
    description=(
        "Collects Indian multi-modal logistics hub (MMLH / MMLP) data from public "
        "websites and fills a 30-column master CSV template via Cursor SDK structured extraction."
    ),
    version="1.0.0",
)

master = MasterAgent()
pipeline = master.pipeline_agent


class RunRequest(BaseModel):
    urls: list[str] = Field(..., min_length=1, description="Public page URLs to process")


class SearchRequest(BaseModel):
    query: str = Field(..., min_length=2, description="Search topic for DuckDuckGo")
    max_urls: int = Field(default=8, ge=1, le=25)


@app.get("/health")
def health() -> dict[str, Any]:
    model = os.getenv("CURSOR_MODEL") or DEFAULT_MODEL
    return {
        "status": "ok",
        "service": "india-logistics-data-agent",
        "model": model,
        "cursor_api_key_configured": bool(os.getenv("CURSOR_API_KEY", "").strip()),
        "output_csv": str(Path(pipeline.output_path)),
    }


@app.get("/template")
def template() -> dict[str, Any]:
    return {
        "columns": TEMPLATE_COLUMNS,
        "count": len(TEMPLATE_COLUMNS),
    }


@app.post("/run")
def run(request: RunRequest) -> dict[str, Any]:
    logger.info("POST /run with %s URL(s)", len(request.urls))
    return master.run_urls(request.urls)


@app.post("/search")
def search(request: SearchRequest) -> dict[str, Any]:
    logger.info("POST /search query=%r max_urls=%s", request.query, request.max_urls)
    return master.run_search(request.query, max_urls=request.max_urls)


@app.get("/results")
def results() -> dict[str, Any]:
    rows = pipeline.read_all_rows()
    return {
        "count": len(rows),
        "columns": TEMPLATE_COLUMNS,
        "rows": rows,
    }


@app.get("/download")
def download() -> FileResponse:
    path = Path(pipeline.output_path)
    if not path.exists() or path.stat().st_size == 0:
        raise HTTPException(status_code=404, detail="Master CSV not found yet. Run /run or /search first.")
    return FileResponse(
        path=path,
        media_type="text/csv",
        filename="India_MMLH_Master_Collected_Data.csv",
    )


if __name__ == "__main__":
    import uvicorn

    host = os.getenv("HOST", "0.0.0.0")
    port = int(os.getenv("PORT", "8000"))
    uvicorn.run("main:app", host=host, port=port, reload=False)
