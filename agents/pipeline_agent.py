"""Pipeline agent: Cursor SDK structured extraction into the master CSV template."""

from __future__ import annotations

import csv
import json
import logging
import os
import re
import tempfile
from pathlib import Path
from typing import Any

from pydantic import ValidationError

from template_schema import (
    TEMPLATE_COLUMNS,
    ExtractionResult,
    LogisticsHubRecord,
)

logger = logging.getLogger(__name__)

DEFAULT_OUTPUT = Path("output") / "India_MMLH_Master_Collected_Data.csv"
DEFAULT_MODEL = "composer-2.5"

_FENCE_RE = re.compile(
    r"```(?:json)?\s*(\{.*?\}|\[.*?\])\s*```",
    re.DOTALL | re.IGNORECASE,
)


class PipelineAgent:
    """Extract hub rows via Cursor Agent structured JSON and append to the CSV."""

    def __init__(
        self,
        api_key: str | None = None,
        model: str | None = None,
        output_path: str | Path | None = None,
    ) -> None:
        self.api_key = (api_key or os.getenv("CURSOR_API_KEY", "").strip()).strip()
        self.model = (model or os.getenv("CURSOR_MODEL") or DEFAULT_MODEL).strip()
        self.output_path = Path(output_path or DEFAULT_OUTPUT)

    def extract_and_save(self, cleaned_text: str, source_url: str) -> dict[str, Any]:
        """
        Extract hub records from cleaned page text and append new rows.

        Returns a small report dict with status and rows_added.
        """
        logger.info("Extracting structured data from %s via Cursor model %s", source_url, self.model)
        try:
            result = self._call_cursor(cleaned_text, source_url)
        except Exception as exc:  # noqa: BLE001 — isolate LLM failures
            logger.exception("Cursor extraction failed for %s: %s", source_url, exc)
            return {
                "status": "extract_failed",
                "rows_added": 0,
                "error": str(exc),
                "hubs": [],
            }

        hubs = result.hubs or []
        if not hubs:
            logger.info("No hub data extracted from %s", source_url)
            return {
                "status": "no_data",
                "rows_added": 0,
                "error": None,
                "hubs": [],
            }

        prepared: list[dict[str, str]] = []
        for hub in hubs:
            row = self._record_to_row(hub, source_url)
            if not (row.get("Hub_Name") or "").strip():
                continue
            prepared.append(row)

        if not prepared:
            logger.info("Extraction produced hubs without Hub_Name for %s", source_url)
            return {
                "status": "no_data",
                "rows_added": 0,
                "error": None,
                "hubs": [],
            }

        added = self._append_unique_rows(prepared)
        # Extraction succeeded (hubs found). Duplicate-only saves still count as ok
        # with rows_added == 0; no_data is reserved for empty extractions.
        logger.info(
            "Pipeline finished for %s: status=ok rows_added=%s",
            source_url,
            added,
        )
        return {
            "status": "ok",
            "rows_added": added,
            "error": None,
            "hubs": prepared,
        }

    def _call_cursor(self, cleaned_text: str, source_url: str) -> ExtractionResult:
        if not self.api_key:
            raise RuntimeError(
                "CURSOR_API_KEY is not set. Add it as a Cursor secret "
                "(or copy .env.example to .env) with a key from "
                "https://cursor.com/dashboard/api"
            )

        from cursor_sdk import (
            Agent,
            AgentOptions,
            LocalAgentOptions,
            SandboxOptions,
        )

        schema_json = json.dumps(ExtractionResult.model_json_schema(), indent=2)
        prompt = (
            "You are a data-extraction worker. Do NOT use tools, do NOT edit files, "
            "do NOT run shell commands. Reply with a single JSON object only — no "
            "markdown, no commentary.\n\n"
            "Extract structured facts about Indian Multi-Modal Logistics Hubs "
            "(MMLH) and Multi-Modal Logistics Parks (MMLP) from the page text.\n\n"
            "Rules:\n"
            "- Only extract information explicitly supported by the text.\n"
            "- If a field is unknown, use null.\n"
            "- Prefer concise factual values over marketing copy.\n"
            "- One object per distinct hub / park mentioned with usable details.\n"
            "- Always set Data_Source_Verification_Link to the source URL below.\n"
            "- Ignore unrelated navigation, ads, or generic India logistics boilerplate.\n"
            "- Return {\"hubs\": []} if the page has no usable hub data.\n\n"
            f"Source URL: {source_url}\n\n"
            "JSON Schema (response must validate against this):\n"
            f"{schema_json}\n\n"
            "PAGE TEXT:\n"
            f"{cleaned_text}"
        )

        # Isolated empty cwd + sandbox so the agent cannot touch the project tree.
        with tempfile.TemporaryDirectory(prefix="mmlh_extract_") as tmp:
            logger.info("Calling Cursor Agent.prompt model=%s cwd=%s", self.model, tmp)
            run = Agent.prompt(
                prompt,
                AgentOptions(
                    api_key=self.api_key,
                    model=self.model,
                    local=LocalAgentOptions(
                        cwd=tmp,
                        sandbox_options=SandboxOptions(enabled=True),
                    ),
                ),
            )

        raw = (run.result or "").strip()
        if not raw:
            raise RuntimeError("Cursor agent returned an empty response")

        logger.info(
            "Cursor run id=%s status=%s duration_ms=%s",
            getattr(run, "id", ""),
            getattr(run, "status", ""),
            getattr(run, "duration_ms", 0),
        )
        return self._parse_extraction(raw)

    def _parse_extraction(self, raw: str) -> ExtractionResult:
        candidates = [raw]
        fence = _FENCE_RE.search(raw)
        if fence:
            candidates.insert(0, fence.group(1).strip())

        # Also try the outermost {...} slice if the model added prose.
        start = raw.find("{")
        end = raw.rfind("}")
        if start != -1 and end != -1 and end > start:
            candidates.append(raw[start : end + 1])

        errors: list[str] = []
        for candidate in candidates:
            try:
                return ExtractionResult.model_validate_json(candidate)
            except ValidationError as exc:
                errors.append(str(exc))
            except Exception as exc:  # noqa: BLE001
                errors.append(str(exc))

        raise RuntimeError(
            "Cursor JSON failed schema validation: "
            + (errors[0] if errors else "unparseable response")
        )

    def _record_to_row(self, hub: LogisticsHubRecord, source_url: str) -> dict[str, str]:
        data = hub.model_dump()
        if not data.get("Data_Source_Verification_Link"):
            data["Data_Source_Verification_Link"] = source_url
        row: dict[str, str] = {}
        for column in TEMPLATE_COLUMNS:
            value = data.get(column)
            if value is None:
                row[column] = ""
            else:
                row[column] = str(value).strip()
        return row

    def _append_unique_rows(self, rows: list[dict[str, str]]) -> int:
        self.output_path.parent.mkdir(parents=True, exist_ok=True)
        existing = self._load_existing_keys()
        to_write: list[dict[str, str]] = []

        for row in rows:
            key = self._dedupe_key(row)
            if not key[0]:
                continue
            if key in existing:
                logger.info(
                    "Skipping duplicate hub: %s @ %s",
                    row.get("Hub_Name"),
                    row.get("Location"),
                )
                continue
            existing.add(key)
            to_write.append(row)

        if not to_write:
            return 0

        write_header = not self.output_path.exists() or self.output_path.stat().st_size == 0
        with self.output_path.open("a", encoding="utf-8", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=TEMPLATE_COLUMNS, extrasaction="ignore")
            if write_header:
                writer.writeheader()
            writer.writerows(to_write)

        logger.info("Appended %s new row(s) to %s", len(to_write), self.output_path)
        return len(to_write)

    def _load_existing_keys(self) -> set[tuple[str, str]]:
        keys: set[tuple[str, str]] = set()
        if not self.output_path.exists():
            return keys
        try:
            with self.output_path.open("r", encoding="utf-8", newline="") as handle:
                reader = csv.DictReader(handle)
                for row in reader:
                    keys.add(self._dedupe_key(row))
        except Exception as exc:  # noqa: BLE001
            logger.warning("Could not read existing CSV for dedupe: %s", exc)
        return keys

    @staticmethod
    def _dedupe_key(row: dict[str, Any]) -> tuple[str, str]:
        name = str(row.get("Hub_Name") or "").strip().lower()
        location = str(row.get("Location") or "").strip().lower()
        return name, location

    def read_all_rows(self) -> list[dict[str, str]]:
        """Return every saved CSV row as a list of dicts."""
        if not self.output_path.exists():
            return []
        with self.output_path.open("r", encoding="utf-8", newline="") as handle:
            return list(csv.DictReader(handle))
