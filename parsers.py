"""Web/PDF parsing and evidence-bound characteristic extraction."""

from __future__ import annotations

import io
import json
import random
import re
import time
from dataclasses import dataclass, field
from html import unescape
from typing import Any
from urllib.parse import urlparse

import requests
from bs4 import BeautifulSoup

from config import Settings
from search_engine import SearchResult, is_blacklisted


@dataclass
class ParsedSource:
    url: str
    status: str
    title: str = ""
    text: str = ""
    provider: str = ""
    error: str = ""
    used: bool = False
    facts: dict[str, str] = field(default_factory=dict)


def _looks_blocked(status_code: int, text: str) -> bool:
    sample = text[:8_000].lower()
    markers = (
        "captcha",
        "cloudflare",
        "verify you are human",
        "проверка браузера",
        "доступ ограничен",
        "антибот",
    )
    return status_code in {401, 403, 429, 503} or any(marker in sample for marker in markers)


def _pdf_text(content: bytes, max_pages: int = 40) -> str:
    try:
        from pypdf import PdfReader
    except ImportError:
        try:
            from PyPDF2 import PdfReader
        except ImportError as exc:
            raise RuntimeError(
                "Чтение PDF недоступно: установите пакет pypdf из requirements.txt"
            ) from exc

    reader = PdfReader(io.BytesIO(content))
    chunks = []
    for page in reader.pages[:max_pages]:
        chunks.append(page.extract_text() or "")
    return "\n".join(chunks)


def _html_text(content: bytes) -> tuple[str, str]:
    soup = BeautifulSoup(content, "html.parser")
    title = soup.title.get_text(" ", strip=True) if soup.title else ""
    description = soup.find("meta", attrs={"name": re.compile("description", re.I)})
    description_text = description.get("content", "") if description else ""

    for tag in soup(["script", "style", "noscript", "svg", "form", "nav", "footer"]):
        tag.decompose()

    chunks = [title, description_text]
    for table in soup.find_all("table"):
        rows = []
        for row in table.find_all("tr"):
            cells = [cell.get_text(" ", strip=True) for cell in row.find_all(["th", "td"])]
            if cells:
                rows.append(" | ".join(cells))
        chunks.append("\n".join(rows))
    chunks.append(soup.get_text("\n", strip=True))
    text = unescape("\n".join(filter(None, chunks)))
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return title, text


class SourceFetcher:
    """Polite per-run HTTP client with normal cookies and domain throttling."""

    def __init__(self, settings: Settings):
        self.settings = settings
        self.session = requests.Session()
        self.session.headers.update(
            {
                "User-Agent": settings.user_agent,
                "Accept": "text/html,application/xhtml+xml,application/pdf;q=0.9,*/*;q=0.8",
                "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.7",
            }
        )
        self.last_request: dict[str, float] = {}
        self.blocked_until: dict[str, float] = {}

    @staticmethod
    def _domain(url: str) -> str:
        return (urlparse(url).hostname or "").lower().removeprefix("www.")

    def _wait_for_domain(self, domain: str) -> None:
        now = time.monotonic()
        blocked_for = self.blocked_until.get(domain, 0.0) - now
        if blocked_for > 0:
            raise RuntimeError(
                f"Домен на охлаждении после CAPTCHA/ограничения: ещё {blocked_for:.0f} сек."
            )

        random_pause = random.uniform(
            self.settings.request_delay_min, self.settings.request_delay_max
        )
        earliest = self.last_request.get(domain, 0.0) + self.settings.domain_cooldown
        delay = max(random_pause, earliest - now)
        if delay > 0:
            time.sleep(delay)

    def fetch(self, result: SearchResult) -> ParsedSource:
        if is_blacklisted(result.url):
            return ParsedSource(result.url, "blacklist", provider=result.provider)

        domain = self._domain(result.url)
        try:
            self._wait_for_domain(domain)
        except RuntimeError as exc:
            return ParsedSource(
                result.url, "охлаждение", provider=result.provider, error=str(exc)
            )

        try:
            response = self.session.get(
                result.url,
                timeout=self.settings.page_timeout,
                allow_redirects=True,
            )
            self.last_request[domain] = time.monotonic()
            if is_blacklisted(response.url):
                return ParsedSource(result.url, "blacklist", provider=result.provider)

            content_type = response.headers.get("content-type", "").lower()
            preview = (
                response.text if "text" in content_type or "html" in content_type else ""
            )
            if _looks_blocked(response.status_code, preview):
                retry_after = response.headers.get("Retry-After", "")
                try:
                    cooldown = max(
                        self.settings.blocked_domain_cooldown, float(retry_after)
                    )
                except (TypeError, ValueError):
                    cooldown = self.settings.blocked_domain_cooldown
                self.blocked_until[domain] = time.monotonic() + cooldown
                return ParsedSource(
                    response.url,
                    "капча — ручная проверка",
                    provider=result.provider,
                    error=(
                        "Автоматические запросы к домену остановлены. "
                        f"Охлаждение {cooldown:.0f} сек.; URL можно открыть вручную."
                    ),
                )

            response.raise_for_status()
            if "pdf" in content_type or response.url.lower().endswith(".pdf"):
                title, text = result.title, _pdf_text(response.content)
            else:
                title, text = _html_text(response.content)
            text = text[: self.settings.max_source_chars]
            status = (
                "открыт"
                if len(text) >= self.settings.min_page_text
                else "мало текста"
            )
            return ParsedSource(response.url, status, title, text, result.provider)
        except Exception as exc:
            self.last_request[domain] = time.monotonic()
            return ParsedSource(
                result.url, "ошибка", provider=result.provider, error=str(exc)[:300]
            )


def fetch_source(
    result: SearchResult, settings: Settings, fetcher: SourceFetcher | None = None
) -> ParsedSource:
    """Backward-compatible entry point; prefer one SourceFetcher per processing run."""
    return (fetcher or SourceFetcher(settings)).fetch(result)


def _normalize_key(value: str) -> str:
    value = re.sub(r"\[[^\]]+]", "", value.lower())
    return re.sub(r"[^a-zа-я0-9]+", "", value)


def _canonical_url(url: str) -> str:
    parsed = urlparse(url.strip())
    host = (parsed.hostname or "").lower().removeprefix("www.")
    path = re.sub(r"/+$", "", parsed.path or "")
    return f"{host}{path}".lower()


def _normalized_evidence(value: str) -> str:
    return re.sub(r"\s+", " ", value.lower()).strip()


def source_matches_product(product_name: str, source: ParsedSource) -> bool:
    """Reject pages for a different model before sending them to Gemini."""
    generic = {
        "pro",
        "airdeck",
        "efi",
        "лодка",
        "лодки",
        "пвх",
        "мотоцикл",
        "мотор",
        "с",
        "псм",
        "seats",
    }
    tokens = [
        token
        for token in re.findall(r"[a-zа-яё0-9]+", product_name.lower())
        if len(token) >= 2 and token not in generic
    ]
    if not tokens:
        return True
    haystack = f"{source.title} {source.text[:12_000]}".lower()
    numeric = [token for token in tokens if any(char.isdigit() for char in token)]
    if numeric and not any(token in haystack for token in numeric):
        return False
    return sum(token in haystack for token in tokens) >= max(1, len(tokens) // 2)


def extract_pairs(text: str) -> dict[str, list[str]]:
    """Conservative local extraction from specification-like lines."""
    pairs: dict[str, list[str]] = {}
    for raw_line in text.splitlines():
        line = raw_line.strip(" |•\t")
        if not 3 <= len(line) <= 220:
            continue
        match = re.match(r"^([^:|]{2,100})\s*(?::|\|)\s*(.{1,100})$", line)
        if not match:
            continue
        key, value = (part.strip() for part in match.groups())
        if len(value) > 100 or not re.search(r"\d|да|нет|yes|no", value, re.I):
            continue
        pairs.setdefault(_normalize_key(key), []).append(value)
    return pairs


def _extract_json(text: str) -> dict[str, Any]:
    text = text.strip().removeprefix("```json").removeprefix("```").removesuffix("```")
    start, end = text.find("{"), text.rfind("}")
    if start < 0 or end < start:
        return {}
    payload = json.loads(text[start : end + 1])
    return payload if isinstance(payload, dict) else {}


def extract_with_gemini(
    product_name: str,
    category: str,
    columns: list[str],
    sources: list[ParsedSource],
    settings: Settings,
) -> dict[str, Any]:
    if not settings.gemini_api_key or not sources:
        return {}
    evidence = "\n\n".join(
        f"SOURCE: {source.url}\n{source.text[:12_000]}" for source in sources if source.text
    )
    prompt = f"""
Ты извлекаешь характеристики товара только из предоставленных источников.
Не используй знания из памяти и не делай предположений.
Товар: {product_name}
Категория: {category}
Разрешённые ключи JSON (точно как в Excel):
{json.dumps(columns, ensure_ascii=False)}

Верни только JSON-объект. Значение каждого ключа должно иметь вид:
{{"value": "значение", "evidence": "короткая цитата/фрагмент", "source": "URL"}}
Если подтверждения нет — не добавляй ключ. Не добавляй другие ключи.

ИСТОЧНИКИ:
{evidence[:48_000]}
""".strip()
    models = [settings.gemini_model, "gemini-2.5-flash-lite"]
    if settings.gemini_model == "gemini-2.0-flash":
        models.insert(0, "gemini-2.5-flash")
    last_error: Exception | None = None
    for model in dict.fromkeys(models):
        url = (
            "https://generativelanguage.googleapis.com/v1beta/models/"
            f"{model}:generateContent"
        )
        for attempt in range(3):
            try:
                response = requests.post(
                    url,
                    headers={"x-goog-api-key": settings.gemini_api_key},
                    json={
                        "contents": [{"parts": [{"text": prompt}]}],
                        "generationConfig": {
                            "temperature": 0,
                            "responseMimeType": "application/json",
                        },
                    },
                    timeout=75,
                )
                if response.status_code in {429, 500, 503, 504}:
                    last_error = RuntimeError(
                        f"Gemini {model}: временная ошибка HTTP {response.status_code}"
                    )
                    if attempt < 2:
                        time.sleep(2 ** attempt * 2)
                        continue
                    break
                response.raise_for_status()
                parts = response.json()["candidates"][0]["content"]["parts"]
                payload = _extract_json(
                    "".join(part.get("text", "") for part in parts)
                )
                if payload:
                    return payload
                last_error = RuntimeError(f"Gemini {model}: получен пустой JSON")
                break
            except requests.RequestException as exc:
                status_code = getattr(getattr(exc, "response", None), "status_code", None)
                last_error = RuntimeError(
                    f"Gemini {model}: ошибка HTTP {status_code or 'сети'}"
                )
                if status_code in {429, 500, 503, 504} and attempt < 2:
                    time.sleep(2 ** attempt * 2)
                    continue
                break
            except Exception as exc:
                last_error = RuntimeError(
                    f"Gemini {model}: {type(exc).__name__}"
                )
                break
    if last_error:
        raise last_error
    return {}


def extract_locally(columns: list[str], sources: list[ParsedSource]) -> dict[str, Any]:
    per_source = [(source, extract_pairs(source.text)) for source in sources]
    output: dict[str, Any] = {}
    for column in columns:
        normalized = _normalize_key(column)
        candidates: list[tuple[str, str]] = []
        for source, pairs in per_source:
            for key, values in pairs.items():
                if normalized and (normalized in key or key in normalized):
                    candidates.extend((value, source.url) for value in values)
        if candidates:
            value, url = candidates[0]
            output[column] = {"value": value, "evidence": value, "source": url}
    return output


def validate_extraction(
    extracted: dict[str, Any], columns: list[str], sources: list[ParsedSource]
) -> dict[str, dict[str, str]]:
    """Reject unknown keys, unsupported values and blacklisted source URLs."""
    allowed = set(columns)
    source_text = {
        _canonical_url(source.url): _normalized_evidence(source.text) for source in sources
    }
    valid: dict[str, dict[str, str]] = {}
    for key, raw in extracted.items():
        if key not in allowed:
            continue
        item = raw if isinstance(raw, dict) else {"value": str(raw)}
        value = str(item.get("value", "")).strip()
        source_url = str(item.get("source", "")).strip()
        evidence = str(item.get("evidence", "")).strip()
        if not value or not source_url or is_blacklisted(source_url):
            continue
        text = source_text.get(_canonical_url(source_url), "")
        if not text:
            continue
        # Require either exact value or evidence to exist in fetched content.
        normalized_value = _normalized_evidence(value)
        normalized_quote = _normalized_evidence(evidence)
        if normalized_value not in text and (
            not normalized_quote or normalized_quote not in text
        ):
            continue
        valid[key] = {"value": value, "source": source_url, "evidence": evidence}
    return valid
