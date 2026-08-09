"""Fetcher agent: download a URL and return clean, table-preserving text."""

from __future__ import annotations

import logging
from dataclasses import dataclass

import requests
from bs4 import BeautifulSoup, Comment

logger = logging.getLogger(__name__)

DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}

STRIP_TAGS = (
    "script",
    "style",
    "noscript",
    "iframe",
    "svg",
    "canvas",
    "form",
    "button",
    "input",
    "select",
    "textarea",
    "nav",
    "footer",
    "header",
    "aside",
)

AD_HINTS = (
    "ad-",
    "ads-",
    "advert",
    "sponsor",
    "cookie",
    "newsletter",
    "popup",
    "modal",
    "social-share",
    "share-buttons",
)


@dataclass
class FetchResult:
    url: str
    ok: bool
    text: str = ""
    title: str = ""
    error: str | None = None
    content_type: str = ""


class FetcherAgent:
    """Download public pages and strip chrome while keeping tables."""

    def __init__(self, timeout: float = 25.0, max_chars: int = 80_000) -> None:
        self.timeout = timeout
        self.max_chars = max_chars
        self.session = requests.Session()
        self.session.headers.update(DEFAULT_HEADERS)

    def fetch(self, url: str) -> FetchResult:
        """Fetch ``url`` and return cleaned text, or a failed result."""
        logger.info("Fetching URL: %s", url)
        try:
            response = self.session.get(url, timeout=self.timeout, allow_redirects=True)
        except requests.Timeout:
            logger.warning("Fetch timeout: %s", url)
            return FetchResult(url=url, ok=False, error="timeout")
        except requests.RequestException as exc:
            logger.warning("Fetch failed for %s: %s", url, exc)
            return FetchResult(url=url, ok=False, error=f"request_error: {exc}")

        content_type = (response.headers.get("Content-Type") or "").lower()
        if response.status_code >= 400:
            logger.warning("HTTP %s for %s", response.status_code, url)
            return FetchResult(
                url=url,
                ok=False,
                error=f"http_{response.status_code}",
                content_type=content_type,
            )

        if "html" not in content_type and not self._looks_like_html(response.text):
            logger.warning("Non-HTML content for %s (%s)", url, content_type or "unknown")
            return FetchResult(
                url=url,
                ok=False,
                error="non_html",
                content_type=content_type,
            )

        try:
            title, text = self._clean_html(response.text)
        except Exception as exc:  # noqa: BLE001
            logger.warning("HTML clean failed for %s: %s", url, exc)
            return FetchResult(
                url=url,
                ok=False,
                error=f"parse_error: {exc}",
                content_type=content_type,
            )

        if not text.strip():
            logger.warning("Empty cleaned text for %s", url)
            return FetchResult(
                url=url,
                ok=False,
                error="empty_content",
                title=title,
                content_type=content_type,
            )

        if len(text) > self.max_chars:
            logger.info(
                "Truncating cleaned text for %s from %s to %s chars",
                url,
                len(text),
                self.max_chars,
            )
            text = text[: self.max_chars]

        logger.info("Fetched OK (%s chars): %s", len(text), url)
        return FetchResult(
            url=str(response.url or url),
            ok=True,
            text=text,
            title=title,
            content_type=content_type,
        )

    @staticmethod
    def _looks_like_html(body: str) -> bool:
        sample = (body or "")[:2000].lower()
        return "<html" in sample or "<body" in sample or "<div" in sample

    def _clean_html(self, html: str) -> tuple[str, str]:
        soup = BeautifulSoup(html, "html.parser")

        for comment in soup.find_all(string=lambda value: isinstance(value, Comment)):
            comment.extract()

        for tag_name in STRIP_TAGS:
            for node in soup.find_all(tag_name):
                node.decompose()

        for node in soup.find_all(True):
            attrs = node.attrs or {}
            identity = " ".join(
                [
                    str(attrs.get("id", "")),
                    " ".join(attrs.get("class", []) if isinstance(attrs.get("class"), list) else [str(attrs.get("class", ""))]),
                    str(attrs.get("role", "")),
                ]
            ).lower()
            if any(hint in identity for hint in AD_HINTS):
                node.decompose()

        title = ""
        if soup.title and soup.title.string:
            title = soup.title.string.strip()

        root = soup.body or soup
        # Preserve tables as readable row/column text before get_text flattening.
        for table in root.find_all("table"):
            table.replace_with(soup.new_string("\n" + self._table_to_text(table) + "\n"))

        for br in root.find_all("br"):
            br.replace_with("\n")

        for block in root.find_all(["p", "li", "tr", "h1", "h2", "h3", "h4", "section", "article"]):
            block.append("\n")

        text = root.get_text(separator=" ", strip=True)
        # Collapse noisy whitespace while keeping paragraph breaks.
        lines = [" ".join(line.split()) for line in text.splitlines()]
        compact = "\n".join(line for line in lines if line)
        return title, compact

    @staticmethod
    def _table_to_text(table) -> str:
        rows: list[str] = []
        for tr in table.find_all("tr"):
            cells = [
                " ".join(cell.get_text(" ", strip=True).split())
                for cell in tr.find_all(["th", "td"])
            ]
            cells = [c for c in cells if c]
            if cells:
                rows.append(" | ".join(cells))
        return "\n".join(rows)
