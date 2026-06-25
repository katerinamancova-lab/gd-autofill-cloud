"""Search providers with mandatory blacklist exclusions."""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass
from urllib.parse import unquote, urlparse

import requests

from config import BLACKLIST, Settings

logger = logging.getLogger(__name__)


@dataclass
class SearchResult:
    title: str
    url: str
    snippet: str = ""
    provider: str = ""


def hostname(url: str) -> str:
    return (urlparse(url).hostname or "").lower().removeprefix("www.")


def is_blacklisted(url: str) -> bool:
    host = hostname(url)
    return any(host == domain or host.endswith("." + domain) for domain in BLACKLIST)


def exclusion_suffix() -> str:
    return " ".join(f"-site:{domain}" for domain in sorted(BLACKLIST))


def build_queries(product_name: str, category: str) -> list[str]:
    quoted = f'"{product_name.strip()}"'
    suffix = exclusion_suffix()
    queries = [
        f"{quoted} характеристики {suffix}",
        f"{quoted} технические характеристики {suffix}",
        f"{quoted} specs specification {suffix}",
        f"{quoted} паспорт инструкция manual pdf {suffix}",
    ]
    if category and category != "Не определена":
        queries.append(f"{product_name} {category} характеристики {suffix}")
    return queries


GENERIC_MODEL_WORDS = {
    "pro",
    "airdeck",
    "лодка",
    "лодки",
    "пвх",
    "мотор",
    "характеристики",
}


def _model_tokens(product_name: str) -> list[str]:
    return [
        token
        for token in re.findall(r"[a-zа-яё0-9]+", product_name.lower())
        if len(token) >= 2 and token not in GENERIC_MODEL_WORDS
    ]


def relevance_score(product_name: str, result: SearchResult) -> float:
    """Prefer exact model pages and reject generic/category-only results."""
    tokens = _model_tokens(product_name)
    if not tokens:
        return 1.0
    haystack = unquote(f"{result.title} {result.snippet} {result.url}").lower()
    matched = sum(token in haystack for token in tokens)
    numeric_tokens = [token for token in tokens if any(char.isdigit() for char in token)]
    if numeric_tokens and not any(token in haystack for token in numeric_tokens):
        return 0.0
    return matched / len(tokens)


class SearchEngine:
    def __init__(self, settings: Settings):
        self.settings = settings
        self.session = requests.Session()
        self.session.headers.update({"User-Agent": settings.user_agent})

    def _serper(self, query: str) -> list[SearchResult]:
        if not self.settings.serper_api_key:
            return []
        response = self.session.post(
            "https://google.serper.dev/search",
            headers={"X-API-KEY": self.settings.serper_api_key},
            json={"q": query, "num": 10},
            timeout=self.settings.search_timeout,
        )
        response.raise_for_status()
        items = response.json().get("organic", [])
        return [
            SearchResult(
                title=item.get("title", ""),
                url=item.get("link", ""),
                snippet=item.get("snippet", ""),
                provider="Serper",
            )
            for item in items
        ]

    def _bing(self, query: str) -> list[SearchResult]:
        if not self.settings.bing_api_key:
            return []
        response = self.session.get(
            "https://api.bing.microsoft.com/v7.0/search",
            headers={"Ocp-Apim-Subscription-Key": self.settings.bing_api_key},
            params={"q": query, "count": 10, "textDecorations": False},
            timeout=self.settings.search_timeout,
        )
        response.raise_for_status()
        items = response.json().get("webPages", {}).get("value", [])
        return [
            SearchResult(
                title=item.get("name", ""),
                url=item.get("url", ""),
                snippet=item.get("snippet", ""),
                provider="Bing",
            )
            for item in items
        ]

    def _duckduckgo(self, query: str) -> list[SearchResult]:
        try:
            from ddgs import DDGS
        except ImportError:
            return []
        items = DDGS(timeout=self.settings.search_timeout).text(
            query, max_results=10, safesearch="moderate"
        )
        return [
            SearchResult(
                title=item.get("title", ""),
                url=item.get("href", ""),
                snippet=item.get("body", ""),
                provider="DuckDuckGo",
            )
            for item in items
        ]

    def search_product(self, product_name: str, category: str) -> list[SearchResult]:
        """Search using available providers, de-duplicate and enforce blacklist."""
        results: list[SearchResult] = []
        seen: set[str] = set()
        providers = (self._serper, self._bing, self._duckduckgo)

        for query in build_queries(product_name, category):
            for provider in providers:
                try:
                    found = provider(query)
                except Exception as exc:
                    logger.warning("Search provider failed: %s", exc)
                    continue
                for item in found:
                    clean_url = item.url.split("#", 1)[0]
                    if not clean_url or clean_url in seen or is_blacklisted(clean_url):
                        continue
                    if relevance_score(product_name, item) < 0.5:
                        continue
                    seen.add(clean_url)
                    item.url = clean_url
                    results.append(item)
                    if len(results) >= self.settings.max_results_per_product:
                        return results
                # A configured API is preferred; DDG remains a fallback.
                if found:
                    break
        return results
