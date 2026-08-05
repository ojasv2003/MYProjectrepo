"""Pipeline agent: Gemini structured extraction into the master CSV template."""

from __future__ import annotations

import csv
import logging
import os
from pathlib import Path
from typing import Any

from google import genai
from google.genai import types
from pydantic import ValidationError

from template_schema import (
    TEMPLATE_COLUMNS,
    ExtractionResult,
    LogisticsHubRecord,
)

logger = logging.getLogger(__name__)

DEFAULT_OUTPUT = Path("output") / "India_MMLH_Master_Collected_Data.csv"
DEFAULT_MODEL = "gemini-2.5-flash"


class PipelineAgent:
    """Extract hub rows via Gemini structured output and append to the CSV."""

    def __init__(
        self,
        api_key: str | None = None,
        model: str | None = None,
        output_path: str | Path | None = None,
    ) -> None:
        self.api_key = (
            api_key
            or os.getenv("GEMINI_API_KEY", "").strip()
            or os.getenv("GOOGLE_API_KEY", "").strip()
        )
        self.model = (model or os.getenv("GEMINI_MODEL") or DEFAULT_MODEL).strip()
        self.output_path = Path(output_path or DEFAULT_OUTPUT)
        self._client: genai.Client | None = None

    @property
    def client(self) -> genai.Client:
        if self._client is None:
            if not self.api_key:
                raise RuntimeError(
                    "GEMINI_API_KEY is not set. Add it as a Cursor secret "
                    "(or copy .env.example to .env) with your Google AI Studio key."
                )
            self._client = genai.Client(api_key=self.api_key)
        return self._client

    def extract_and_save(self, cleaned_text: str, source_url: str) -> dict[str, Any]:
        """
        Extract hub records from cleaned page text and append new rows.

        Returns a small report dict with status and rows_added.
        """
        logger.info("Extracting structured data from %s via %s", source_url, self.model)
        try:
            result = self._call_gemini(cleaned_text, source_url)
        except Exception as exc:  # noqa: BLE001 — isolate Gemini failures
            logger.exception("Gemini extraction failed for %s: %s", source_url, exc)
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

    def _call_gemini(self, cleaned_text: str, source_url: str) -> ExtractionResult:
        prompt = (
            "You are extracting structured facts about Indian Multi-Modal Logistics "
            "Hubs (MMLH) and Multi-Modal Logistics Parks (MMLP) from a web page.\n\n"
            "Rules:\n"
            "- Only extract information explicitly supported by the text.\n"
            "- If a field is unknown, leave it null.\n"
            "- Prefer concise factual values over marketing copy.\n"
            "- One object per distinct hub / park mentioned with usable details.\n"
            "- Always set Data_Source_Verification_Link to the source URL below.\n"
            "- Ignore unrelated navigation, ads, or generic India logistics boilerplate.\n\n"
            f"Source URL: {source_url}\n\n"
            "PAGE TEXT:\n"
            f"{cleaned_text}"
        )

        response = self.client.models.generate_content(
            model=self.model,
            contents=prompt,
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                response_json_schema=ExtractionResult.model_json_schema(),
                temperature=0.1,
            ),
        )

        raw = (response.text or "").strip()
        if not raw:
            raise RuntimeError("Gemini returned an empty response")

        try:
            return ExtractionResult.model_validate_json(raw)
        except ValidationError as exc:
            raise RuntimeError(f"Gemini JSON failed schema validation: {exc}") from exc

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
