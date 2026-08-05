"""Master agent: orchestrate search → fetch → extract for batches of URLs."""

from __future__ import annotations

import logging
from typing import Any

from agents.fetcher_agent import FetcherAgent
from agents.pipeline_agent import PipelineAgent
from agents.search_agent import SearchAgent

logger = logging.getLogger(__name__)


class MasterAgent:
    """Coordinate the four-stage logistics data collection workflow."""

    def __init__(
        self,
        search_agent: SearchAgent | None = None,
        fetcher_agent: FetcherAgent | None = None,
        pipeline_agent: PipelineAgent | None = None,
    ) -> None:
        self.search_agent = search_agent or SearchAgent()
        self.fetcher_agent = fetcher_agent or FetcherAgent()
        self.pipeline_agent = pipeline_agent or PipelineAgent()

    def run_urls(self, urls: list[str]) -> dict[str, Any]:
        """Fetch + extract each URL; never abort the batch on a single failure."""
        cleaned_urls = [u.strip() for u in urls if isinstance(u, str) and u.strip()]
        logger.info("MasterAgent starting batch of %s URL(s)", len(cleaned_urls))

        reports: list[dict[str, Any]] = []
        total_added = 0

        for url in cleaned_urls:
            report = self._process_one(url)
            reports.append(report)
            total_added += int(report.get("rows_added") or 0)

        summary = {
            "urls_processed": len(cleaned_urls),
            "rows_added_total": total_added,
            "ok": sum(1 for r in reports if r["status"] == "ok"),
            "fetch_failed": sum(1 for r in reports if r["status"] == "fetch_failed"),
            "extract_failed": sum(1 for r in reports if r["status"] == "extract_failed"),
            "no_data": sum(1 for r in reports if r["status"] == "no_data"),
        }
        logger.info("MasterAgent batch complete: %s", summary)
        return {"summary": summary, "reports": reports}

    def run_search(self, query: str, max_urls: int = 8) -> dict[str, Any]:
        """Search for candidate URLs, then run the fetch/extract pipeline."""
        logger.info("MasterAgent search workflow: query=%r max_urls=%s", query, max_urls)
        try:
            urls = self.search_agent.search(query, max_urls=max_urls)
        except Exception as exc:  # noqa: BLE001
            logger.exception("Search stage failed: %s", exc)
            return {
                "query": query,
                "urls": [],
                "summary": {
                    "urls_processed": 0,
                    "rows_added_total": 0,
                    "ok": 0,
                    "fetch_failed": 0,
                    "extract_failed": 0,
                    "no_data": 0,
                },
                "reports": [],
                "error": f"search_failed: {exc}",
            }

        result = self.run_urls(urls)
        result["query"] = query
        result["urls"] = urls
        return result

    def _process_one(self, url: str) -> dict[str, Any]:
        logger.info("Processing URL: %s", url)
        try:
            fetched = self.fetcher_agent.fetch(url)
        except Exception as exc:  # noqa: BLE001
            logger.exception("Unexpected fetch error for %s: %s", url, exc)
            return {
                "url": url,
                "status": "fetch_failed",
                "rows_added": 0,
                "error": str(exc),
                "title": "",
            }

        if not fetched.ok:
            return {
                "url": url,
                "status": "fetch_failed",
                "rows_added": 0,
                "error": fetched.error,
                "title": fetched.title,
            }

        try:
            pipeline = self.pipeline_agent.extract_and_save(fetched.text, fetched.url or url)
        except Exception as exc:  # noqa: BLE001
            logger.exception("Unexpected pipeline error for %s: %s", url, exc)
            return {
                "url": url,
                "status": "extract_failed",
                "rows_added": 0,
                "error": str(exc),
                "title": fetched.title,
            }

        return {
            "url": url,
            "final_url": fetched.url,
            "title": fetched.title,
            "status": pipeline.get("status", "extract_failed"),
            "rows_added": pipeline.get("rows_added", 0),
            "error": pipeline.get("error"),
            "hubs": [
                {"Hub_Name": h.get("Hub_Name"), "Location": h.get("Location")}
                for h in pipeline.get("hubs") or []
            ],
        }
