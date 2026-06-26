"""Firecrawl-based source reader for product characteristic extraction."""

from __future__ import annotations

import re
from typing import Any

import requests

from config import Settings
from parsers import ParsedSource
from search_engine import SearchResult, is_blacklisted


# Firecrawl docs currently use the v2 scrape endpoint.
FIRECRAWL_ENDPOINT = "https://api.firecrawl.dev/v2/scrape"


def _clean_text(value: str) -> str:
    value = re.sub(r"[ \t]+", " ", value or "")
    value = re.sub(r"\n{3,}", "\n\n", value)
    return value.strip()


def _pick_text(data: dict[str, Any]) -> tuple[str, str]:
    """Extract title and the best available text payload from Firecrawl response."""
    payload = data.get("data") if isinstance(data.get("data"), dict) else data
    metadata = payload.get("metadata") or {}
    title = str(metadata.get("title") or payload.get("title") or "").strip()
    description = str(metadata.get("description") or "").strip()
    parts = [
        str(payload.get("markdown") or "").strip(),
        str(payload.get("text") or "").strip(),
        str(payload.get("html") or "").strip(),
        description,
    ]
    text = _clean_text("\n\n".join(part for part in parts if part))
    return title, text


def _response_error(response: requests.Response) -> str:
    body = (response.text or "").strip().replace("\n", " ")
    if len(body) > 240:
        body = body[:240] + "..."
    return f"Firecrawl HTTP {response.status_code}: {body or response.reason}"


class FirecrawlClient:
    """Small HTTP client for Firecrawl scrape API.

    It intentionally does not try to bypass CAPTCHA or anti-bot protections.
    If Firecrawl cannot read a URL, the caller receives an error source and
    continues with the next candidate URL.
    """

    def __init__(self, settings: Settings):
        self.settings = settings
        self.session = requests.Session()

    def scrape(self, result: SearchResult) -> ParsedSource:
        if is_blacklisted(result.url):
            return ParsedSource(result.url, "blacklist", provider=result.provider)
        if not self.settings.firecrawl_api_key:
            return ParsedSource(
                result.url,
                "firecrawl недоступен",
                provider=result.provider,
                error="FIRECRAWL_API_KEY не задан в Streamlit Secrets",
            )

        try:
            response = self.session.post(
                FIRECRAWL_ENDPOINT,
                headers={
                    "Authorization": f"Bearer {self.settings.firecrawl_api_key}",
                    "Content-Type": "application/json",
                },
                json={
                    "url": result.url,
                    "formats": ["markdown"],
                    "onlyMainContent": True,
                },
                timeout=self.settings.firecrawl_timeout + 10,
            )
            if response.status_code in {401, 403}:
                return ParsedSource(
                    result.url,
                    "firecrawl ошибка",
                    provider=result.provider,
                    error=_response_error(response),
                )
            if response.status_code in {400, 408, 409, 425, 429, 500, 502, 503, 504}:
                return ParsedSource(
                    result.url,
                    "firecrawl ошибка",
                    provider=result.provider,
                    error=_response_error(response),
                )
            response.raise_for_status()
            payload = response.json()
            title, text = _pick_text(payload)
            text = text[: self.settings.max_source_chars]
            if len(text) < self.settings.min_page_text:
                return ParsedSource(
                    result.url,
                    "пусто firecrawl",
                    title=title,
                    text=text,
                    provider=result.provider,
                    error="Firecrawl вернул слишком мало текста",
                )
            return ParsedSource(
                result.url,
                "открыт",
                title=title or result.title,
                text=text,
                provider=f"{result.provider} + Firecrawl",
            )
        except Exception as exc:
            return ParsedSource(
                result.url,
                "firecrawl ошибка",
                provider=result.provider,
                error=str(exc)[:300],
            )
