"""Search agent: find candidate source URLs via DuckDuckGo (no Gemini, no API key)."""

from __future__ import annotations

import logging
import re
from urllib.parse import parse_qs, unquote, urlparse

import requests
from bs4 import BeautifulSoup

logger = logging.getLogger(__name__)

DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}

# Skip obvious non-content destinations.
BLOCKED_HOST_FRAGMENTS = (
    "duckduckgo.com",
    "google.com",
    "bing.com",
    "youtube.com",
    "facebook.com",
    "twitter.com",
    "x.com",
    "instagram.com",
    "linkedin.com",
    "wikipedia.org",
)


class SearchAgent:
    """Locate public web pages about Indian multi-modal logistics hubs."""

    def __init__(self, timeout: float = 20.0) -> None:
        self.timeout = timeout
        self.session = requests.Session()
        self.session.headers.update(DEFAULT_HEADERS)

    def search(self, query: str, max_urls: int = 8) -> list[str]:
        """Return up to ``max_urls`` unique candidate URLs for ``query``."""
        topic = (query or "").strip()
        if not topic:
            logger.warning("Search skipped: empty query")
            return []

        max_urls = max(1, min(int(max_urls), 25))
        logger.info("Searching DuckDuckGo for up to %s URLs: %s", max_urls, topic)

        urls: list[str] = []
        seen: set[str] = set()

        # Prefer the HTML endpoint; fall back to lite if needed.
        for fetcher in (self._search_html, self._search_lite):
            try:
                candidates = fetcher(topic)
            except Exception as exc:  # noqa: BLE001 — never fail the batch
                logger.warning("DuckDuckGo %s failed: %s", fetcher.__name__, exc)
                continue

            for url in candidates:
                normalized = self._normalize_url(url)
                if not normalized or normalized in seen:
                    continue
                if self._is_blocked(normalized):
                    continue
                seen.add(normalized)
                urls.append(normalized)
                if len(urls) >= max_urls:
                    logger.info("Search complete: %s URLs", len(urls))
                    return urls

        logger.info("Search complete: %s URLs", len(urls))
        return urls

    def _search_html(self, query: str) -> list[str]:
        response = self.session.get(
            "https://html.duckduckgo.com/html/",
            params={"q": query},
            timeout=self.timeout,
        )
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")
        urls: list[str] = []

        for anchor in soup.select("a.result__a"):
            href = anchor.get("href")
            if not href:
                continue
            urls.append(self._unwrap_ddg_redirect(href))

        if not urls:
            # Alternate markup occasionally used by DDG HTML.
            for anchor in soup.select("a.result-link, a[href*='uddg=']"):
                href = anchor.get("href")
                if href:
                    urls.append(self._unwrap_ddg_redirect(href))

        logger.debug("html.duckduckgo.com returned %s raw links", len(urls))
        return urls

    def _search_lite(self, query: str) -> list[str]:
        response = self.session.post(
            "https://lite.duckduckgo.com/lite/",
            data={"q": query},
            timeout=self.timeout,
        )
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")
        urls: list[str] = []

        for anchor in soup.select("a[href]"):
            href = anchor.get("href", "")
            if not href.startswith("http"):
                continue
            if "duckduckgo.com" in href:
                continue
            urls.append(href)

        logger.debug("lite.duckduckgo.com returned %s raw links", len(urls))
        return urls

    @staticmethod
    def _unwrap_ddg_redirect(href: str) -> str:
        """Extract the real target from DuckDuckGo redirect wrappers."""
        if "uddg=" in href:
            parsed = urlparse(href)
            qs = parse_qs(parsed.query)
            if "uddg" in qs and qs["uddg"]:
                return unquote(qs["uddg"][0])
        if href.startswith("//"):
            return "https:" + href
        return href

    @staticmethod
    def _normalize_url(url: str) -> str | None:
        url = (url or "").strip()
        if not url.startswith(("http://", "https://")):
            return None
        # Drop fragment noise.
        parsed = urlparse(url)
        cleaned = parsed._replace(fragment="").geturl()
        return cleaned.rstrip("/")

    @staticmethod
    def _is_blocked(url: str) -> bool:
        host = urlparse(url).netloc.lower()
        if any(fragment in host for fragment in BLOCKED_HOST_FRAGMENTS):
            return True
        # Skip binary / document-only links that the fetcher cannot usefully parse.
        if re.search(r"\.(pdf|docx?|xlsx?|zip|rar)(\?|$)", url, re.I):
            return True
        return False
