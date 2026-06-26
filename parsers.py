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

        last_request = self.last_request.get(domain)
        if last_request is None:
            return
        random_pause = random.uniform(
            self.settings.request_delay_min, self.settings.request_delay_max
        )
        earliest = last_request + self.settings.domain_cooldown
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
    return re.sub(r"[^a-zа-яё0-9]+", "", value)


def _canonical_url(url: str) -> str:
    parsed = urlparse(url.strip())
    host = (parsed.hostname or "").lower().removeprefix("www.")
    path = re.sub(r"/+$", "", parsed.path or "")
    return f"{host}{path}".lower()


def _normalized_evidence(value: str) -> str:
    return re.sub(r"\s+", " ", value.lower()).strip()


def _column_key_map(columns: list[str]) -> dict[str, str]:
    """Map exact Excel headers and their visible labels back to exact headers."""
    mapping: dict[str, str] = {}
    for column in columns:
        mapping[column] = column
        mapping[_display_label(column)] = column
        mapping[_normalize_key(column)] = column
        mapping[_normalize_key(_display_label(column))] = column
    return {key: value for key, value in mapping.items() if key}


def source_matches_product(product_name: str, source: ParsedSource) -> bool:
    """Reject pages for a different model before sending them to Gemini."""
    generic = {
        "pro",
        "airdeck",
        "фб",
        "fb",
        "нднд",
        "ндвд",
        "efi",
        "лодка",
        "лодки",
        "пвх",
        "мотоцикл",
        "мотоциклы",
        "мотор",
        "с",
        "псм",
        "abs",
        "tour",
        "rally",
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
    brand_or_model_hits = sum(token in haystack for token in tokens)
    return brand_or_model_hits >= 1


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


def _column_aliases(column: str) -> list[str]:
    label = _display_label(column).lower()
    code_match = re.search(r"\[([^\]]+)]", column)
    code = (code_match.group(1).lower() if code_match else "")
    aliases = [label]
    if "бренд" in label or code == "brand":
        aliases += ["марка", "бренд", "brand", "manufacturer"]
    if "объём двигателя" in label or "объем двигателя" in label or "engine_capacity" in code:
        aliases += [
            "объем двигателя",
            "объём двигателя",
            "рабочий объем",
            "рабочий объём",
            "кубатура",
            "engine capacity",
            "engine displacement",
            "displacement",
        ]
    if "мощность" in label or "power" in code:
        aliases += ["мощность", "максимальная мощность", "power", "max power"]
    if "охлаждение" in label or "cooling" in code:
        aliases += ["охлаждение", "система охлаждения", "cooling"]
    if "двигатель" in label and "объ" not in label:
        aliases += ["двигатель", "модель двигателя", "engine", "engine type"]
    if "колёсная база" in label or "колесная база" in label or "wheelbase" in code:
        aliases += ["колесная база", "колёсная база", "wheelbase"]
    if "вес" in label or "weight" in code:
        aliases += ["вес", "масса", "сухая масса", "снаряженная масса", "weight"]
    if "клиренс" in label or "clearance" in code:
        aliases += ["клиренс", "дорожный просвет", "clearance", "ground clearance"]
    if "топлив" in label and "тип" not in label:
        aliases += ["система питания", "подача топлива", "топливная система", "fuel system", "fuel supply"]
    if "трансмис" in label:
        aliases += ["трансмиссия", "коробка передач", "кпп", "transmission"]
    if "запуск" in label or "starting" in code:
        aliases += ["система запуска", "запуск", "стартер", "starting system"]
    if "страна бренда" in label:
        aliases += ["страна бренда", "страна марки", "brand country"]
    if "страна производства" in label or "manufacturer" in code:
        aliases += ["страна производства", "производство", "made in", "manufacturer country"]
    if "материал рамы" in label:
        aliases += ["рама", "материал рамы", "frame", "frame material"]
    if "тип топлива" in label:
        aliases += ["тип топлива", "топливо", "fuel type"]
    if "количество тактов" in label or "stroke" in code:
        aliases += ["количество тактов", "тактов", "тактность", "stroke"]
    if "тип двигателя" in label:
        aliases += ["тип двигателя", "двигатель", "engine type"]
    if "объём бака" in label or "объем бака" in label or "fuel_capacity" in code:
        aliases += ["объем бака", "объём бака", "топливный бак", "бак", "fuel tank", "fuel capacity"]
    if "цилиндр" in label:
        aliases += ["количество цилиндров", "цилиндров", "cylinders"]
    if "подвеска передняя" in label:
        aliases += ["передняя подвеска", "подвеска передняя", "front suspension"]
    if "подвеска задняя" in label:
        aliases += ["задняя подвеска", "подвеска задняя", "rear suspension"]
    if "тормоза передние" in label:
        aliases += ["передние тормоза", "тормоз передний", "front brake", "front brakes"]
    if "тормоза задние" in label:
        aliases += ["задние тормоза", "тормоз задний", "rear brake", "rear brakes"]
    if label.startswith("длина"):
        aliases += ["длина", "length"]
    if label.startswith("ширина"):
        aliases += ["ширина", "width"]
    if label.startswith("высота") and "седл" not in label:
        aliases += ["высота", "height"]
    if "высота по седлу" in label or "seat_height" in code:
        aliases += ["высота по седлу", "высота сиденья", "seat height"]
    if "скорость" in label:
        aliases += ["максимальная скорость", "макс. скорость", "max speed", "maximum speed"]
    if "колеса передние" in label or "колёса передние" in label:
        aliases += ["переднее колесо", "передние колеса", "передние колёса", "front wheel", "front tire"]
    if "колеса задние" in label or "колёса задние" in label:
        aliases += ["заднее колесо", "задние колеса", "задние колёса", "rear wheel", "rear tire"]
    return list(dict.fromkeys(_normalize_key(alias) for alias in aliases if alias))


def _line_regex_value(source: ParsedSource, aliases: list[str]) -> str:
    for raw_line in source.text.splitlines():
        line = re.sub(r"\s+", " ", raw_line.strip(" |•\t"))
        if not 4 <= len(line) <= 240:
            continue
        normalized_line = _normalize_key(line)
        if not any(alias and alias in normalized_line for alias in aliases):
            continue
        match = re.search(r"(?:[:|–—-]\s*|\s{2,})([^:|]{1,120})$", line)
        if match:
            return match.group(1).strip()
    return ""


def _regex_value(column: str, source: ParsedSource) -> str:
    text = re.sub(r"\s+", " ", source.text)
    label = _display_label(column).lower()
    patterns: list[str] = []
    if "объём двигателя" in label or "объем двигателя" in label:
        patterns = [r"(?:об[ъь]е[мё]\s+двигателя|рабочий\s+об[ъь]е[мё]|displacement)[^\d]{0,40}(\d{2,4}(?:[,.]\d+)?)"]
    elif "мощность" in label:
        patterns = [r"(?:мощность|max(?:imum)?\s+power|power)[^\d]{0,40}(\d{1,3}(?:[,.]\d+)?)"]
    elif "колёсная база" in label or "колесная база" in label:
        patterns = [r"(?:кол[её]сная\s+база|wheelbase)[^\d]{0,40}(\d{3,4})"]
    elif "вес" in label:
        patterns = [r"(?:сухая\s+масса|снаряженная\s+масса|масса|вес|weight)[^\d]{0,40}(\d{2,4}(?:[,.]\d+)?)"]
    elif "клиренс" in label:
        patterns = [r"(?:клиренс|дорожный\s+просвет|clearance)[^\d]{0,40}(\d{2,4})"]
    elif "объём бака" in label or "объем бака" in label:
        patterns = [r"(?:об[ъь]е[мё]\s+бака|топливный\s+бак|fuel\s+(?:tank|capacity))[^\d]{0,40}(\d{1,2}(?:[,.]\d+)?)"]
    elif "количество цилиндров" in label:
        patterns = [r"(?:цилиндров|cylinders)[^\d]{0,40}(\d{1,2})"]
    elif "количество тактов" in label:
        patterns = [r"(?:тактов|тактный|stroke)[^\d]{0,40}(\d)"]
    elif "высота по седлу" in label:
        patterns = [r"(?:высота\s+(?:по\s+)?седл[ау]|seat\s+height)[^\d]{0,40}(\d{3,4})"]
    elif "скорость" in label:
        patterns = [r"(?:макс(?:имальная)?\.?\s+скорость|max(?:imum)?\s+speed)[^\d]{0,40}(\d{2,3})"]
    for pattern in patterns:
        match = re.search(pattern, text, re.I)
        if match:
            return match.group(1)
    return ""


def _extract_json(text: str) -> dict[str, Any]:
    text = text.strip().removeprefix("```json").removeprefix("```").removesuffix("```")
    start, end = text.find("{"), text.rfind("}")
    if start < 0 or end < start:
        return {}
    payload = json.loads(text[start : end + 1])
    return payload if isinstance(payload, dict) else {}


def _display_label(column: str) -> str:
    return re.sub(r"\s*\[[^\]]+]\s*$", "", column).strip()


def _column_kind(column: str) -> str:
    label = _display_label(column).lower()
    if "объём двигателя" in label and "диапазон" not in label:
        return "engine_cc_number"
    if "мощность" in label and "диапазон" not in label:
        return "power_hp_number"
    if any(part in label for part in ("колёсная база", "колесная база", "клиренс", "высота по седлу")):
        return "mm_number"
    if "подвеска" in label:
        return "suspension"
    if label.startswith(("длина", "ширина", "высота")) and "седл" not in label:
        return "cm_number"
    if "вес" in label:
        return "kg_number"
    if "объём бака" in label or "объем бака" in label:
        return "liter_number"
    if "количество цилиндров" in label or "количество тактов" in label:
        return "small_integer"
    if "макс" in label and "скор" in label:
        return "speed_number"
    if "тип топлива" in label:
        return "fuel_type"
    if "тип мотоцикла" in label:
        return "motorcycle_type"
    if "перед" in label and ("тормоз" in label or "колес" in label or "колёс" in label):
        return "front_only"
    if "зад" in label and ("тормоз" in label or "колес" in label or "колёс" in label):
        return "rear_only"
    return "text"


def _column_guide(columns: list[str]) -> list[dict[str, str]]:
    guide = []
    for column in columns:
        label = _display_label(column)
        guide.append(
            {
                "key": column,
                "label": label,
                "kind": _column_kind(column),
                "rule": (
                    "Ориентируйся на label. Текст в квадратных скобках — только внутренний код; "
                    "если код противоречит label, используй label."
                ),
            }
        )
    return guide


def extract_with_gemini(
    product_name: str,
    category: str,
    columns: list[str],
    sources: list[ParsedSource],
    settings: Settings,
) -> dict[str, Any]:
    if not settings.gemini_api_key or not sources:
        return {}
    column_guide = json.dumps(_column_guide(columns), ensure_ascii=False)

    def build_prompt(source_limit: int, total_limit: int) -> str:
        evidence = "\n\n".join(
            f"SOURCE: {source.url}\n{source.text[:source_limit]}"
            for source in sources
            if source.text
        )
        return f"""
Ты извлекаешь характеристики товара только из предоставленных источников.
Не используй знания из памяти и не делай предположений.
Если значение найдено хотя бы в одном нормальном источнике и относится к этому товару — верни его.
Не требуй подтверждения в нескольких источниках.
Не заполняй поле, если источник говорит о другом товаре, другой модификации или значение относится к другой характеристике.
Не бери телефоны, email, адреса, цены, SEO-текст, крошки меню, условия доставки и рекламные фразы как характеристики.
Если в названии столбца есть текст в квадратных скобках, это внутренний код. Главный смысл столбца — русский label до скобок.
Например: "Тормоза передние [REAR_BRAKE]" означает именно передние тормоза, несмотря на код REAR_BRAKE.
Для числовых столбцов возвращай только число без лишнего текста и единиц, если это возможно.
Для длины/ширины/высоты в сантиметрах переводи миллиметры в сантиметры.
Ключи JSON должны совпадать со строками из поля "key" один в один.
Товар: {product_name}
Категория: {category}
Столбцы для заполнения:
{column_guide}

Верни только JSON-объект. Значение каждого ключа должно иметь вид:
{{"value": "значение", "evidence": "короткая цитата/фрагмент", "source": "URL"}}
Если подтверждения нет — не добавляй ключ. Не добавляй другие ключи.

ИСТОЧНИКИ:
{evidence[:total_limit]}
""".strip()

    prompt_variants = [
        build_prompt(settings.max_gemini_source_chars, settings.max_gemini_total_chars),
        build_prompt(12_000, 35_000),
    ]
    models = [settings.gemini_model, "gemini-2.5-flash-lite"]
    if settings.gemini_model == "gemini-2.0-flash":
        models.insert(0, "gemini-2.5-flash")
    last_error: Exception | None = None
    for model in dict.fromkeys(models):
        url = (
            "https://generativelanguage.googleapis.com/v1beta/models/"
            f"{model}:generateContent"
        )
        attempts = 2 if model == settings.gemini_model else 1
        for prompt_index, prompt in enumerate(prompt_variants):
            if prompt_index > 0:
                time.sleep(1)
            for attempt in range(attempts):
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
                        timeout=45,
                    )
                    if response.status_code in {429, 500, 503, 504}:
                        last_error = RuntimeError(
                            f"Gemini {model}: временная ошибка HTTP {response.status_code}"
                        )
                        if attempt < attempts - 1:
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
                    if status_code in {429, 500, 503, 504} and attempt < attempts - 1:
                        time.sleep(2 ** attempt * 2)
                        continue
                    break
                except Exception as exc:
                    last_error = RuntimeError(
                        f"Gemini {model}: {type(exc).__name__}"
                    )
                    break
            if not last_error or "HTTP 429" not in str(last_error):
                break
    if last_error:
        raise last_error
    return {}


def extract_locally(
    columns: list[str], sources: list[ParsedSource], product_name: str = ""
) -> dict[str, Any]:
    per_source = [(source, extract_pairs(source.text)) for source in sources]
    output: dict[str, Any] = {}
    for column in columns:
        normalized = _normalize_key(column)
        aliases = _column_aliases(column)
        label = _display_label(column).lower()
        candidates: list[tuple[str, str]] = []

        if ("бренд" in label or "[BRAND]" in column.upper()) and product_name:
            brand = re.split(r"\s+", product_name.strip())[0]
            if brand and any(brand.lower() in f"{source.title} {source.text[:4000]}".lower() for source in sources):
                output[column] = {
                    "value": brand,
                    "evidence": product_name,
                    "source": sources[0].url if sources else "manual://product-name",
                }
                continue

        for source, pairs in per_source:
            for key, values in pairs.items():
                if normalized and (normalized in key or key in normalized):
                    candidates.extend((value, source.url) for value in values)
                    continue
                if any(alias and (alias in key or key in alias) for alias in aliases):
                    candidates.extend((value, source.url) for value in values)
            if not candidates:
                value = _line_regex_value(source, aliases) or _regex_value(column, source)
                if value:
                    candidates.append((value, source.url))
        if candidates:
            value, url = candidates[0]
            output[column] = {"value": value, "evidence": value, "source": url}
    return output


CONTACT_RE = re.compile(
    r"(@|email|e-mail|тел\.?|телефон|whatsapp|viber|\+?\d[\d\s().-]{8,})",
    re.I,
)


def _first_number(value: str) -> str:
    match = re.search(r"\d+(?:[.,]\d+)?", value)
    return match.group(0).replace(",", ".") if match else ""


def _format_number(value: float) -> str:
    if abs(value - round(value)) < 0.05:
        return str(int(round(value)))
    return f"{value:.1f}".rstrip("0").rstrip(".")


def _has_front(value: str) -> bool:
    return bool(re.search(r"передн|front", value, re.I))


def _has_rear(value: str) -> bool:
    return bool(re.search(r"задн|rear", value, re.I))


def _sanitize_value_for_column(column: str, value: str, evidence: str) -> str:
    """Return a normalized safe value for a column, or empty string to reject it."""
    raw = re.sub(r"\s+", " ", str(value or "")).strip()
    if not raw:
        return ""
    check_text = f"{raw} {evidence}"
    if CONTACT_RE.search(raw):
        return ""
    if len(raw) > 180:
        return ""

    kind = _column_kind(column)
    lowered = raw.lower()
    combined = check_text.lower()

    if kind in {
        "engine_cc_number",
        "power_hp_number",
        "mm_number",
        "cm_number",
        "kg_number",
        "liter_number",
        "small_integer",
        "speed_number",
    }:
        number = _first_number(raw)
        if not number:
            return ""
        numeric = float(number)
        if kind == "engine_cc_number" and not (40 <= numeric <= 3000):
            return ""
        if kind == "power_hp_number" and not (1 <= numeric <= 250):
            return ""
        if kind == "mm_number" and not (50 <= numeric <= 2500):
            return ""
        if "высота по седлу" in _display_label(column).lower() and not (400 <= numeric <= 1000):
            return ""
        if kind == "cm_number":
            # The Excel header asks for centimeters. Many sources publish overall
            # dimensions in millimeters, so convert obvious mm values.
            if numeric > 1000 or "мм" in lowered or " mm" in lowered:
                numeric = numeric / 10
            if not (20 <= numeric <= 500):
                return ""
        if kind == "kg_number" and not (10 <= numeric <= 800):
            return ""
        if kind == "liter_number" and not (1 <= numeric <= 80):
            return ""
        if kind == "small_integer" and not (1 <= numeric <= 12):
            return ""
        if kind == "speed_number" and not (5 <= numeric <= 350):
            return ""
        return _format_number(numeric)

    if kind == "fuel_type":
        if any(word in lowered for word in ("abs", "тормоз", "подвес", "диск", "электронно управляем")):
            return ""
        if not any(word in lowered for word in ("бензин", "аи-", "diesel", "gasoline", "petrol", "электро")):
            return ""
        return raw[:80]

    if kind == "motorcycle_type":
        if any(word in lowered for word in ("abs", "тормоз", "диск", "электронно управляем", "pgm-fi")):
            return ""
        return raw[:80]

    if kind == "front_only":
        if _has_front(lowered) and _has_rear(lowered):
            return ""
        if _has_rear(lowered) and not _has_front(lowered):
            return ""
        return raw[:140]

    if kind == "rear_only":
        if _has_front(lowered) and _has_rear(lowered):
            return ""
        if _has_front(lowered) and not _has_rear(lowered):
            return ""
        return raw[:140]

    if kind == "suspension":
        if len(raw) < 8 or re.fullmatch(r"\d+(?:[.,]\d+)?", raw):
            return ""
        if not any(word in lowered for word in ("вилка", "аморт", "подвес", "маятник", "телескоп", "моно")):
            return ""
        return raw[:160]

    if any(bad in lowered for bad in ("email", "e-mail", "телефон", "whatsapp", "корзина", "купить")):
        return ""
    return raw[:180]


def validate_extraction(
    extracted: dict[str, Any], columns: list[str], sources: list[ParsedSource]
) -> dict[str, dict[str, str]]:
    """Reject unknown keys, unsupported values and blacklisted source URLs."""
    key_map = _column_key_map(columns)
    source_text = {
        _canonical_url(source.url): _normalized_evidence(source.text) for source in sources
    }
    known_urls = {_canonical_url(source.url): source.url for source in sources}
    valid: dict[str, dict[str, str]] = {}
    for key, raw in extracted.items():
        column = key_map.get(str(key).strip()) or key_map.get(_normalize_key(str(key)))
        if not column:
            continue
        item = raw if isinstance(raw, dict) else {"value": str(raw)}
        value = str(item.get("value", "")).strip()
        source_url = str(item.get("source", "")).strip()
        evidence = str(item.get("evidence", "")).strip()
        if not value:
            continue
        if not source_url and sources:
            source_url = sources[0].url
        if not source_url or is_blacklisted(source_url):
            continue
        value = _sanitize_value_for_column(column, value, evidence)
        if not value:
            continue
        canonical = _canonical_url(source_url)
        text = source_text.get(canonical, "")
        normalized_value = _normalized_evidence(value)
        normalized_quote = _normalized_evidence(evidence)

        if not text:
            # Gemini sometimes returns a canonical/redirect URL not byte-identical
            # to the search result. Find the fetched source that contains the
            # quoted evidence or the normalized value and attribute the fill to it.
            for source in sources:
                candidate_text = source_text.get(_canonical_url(source.url), "")
                if (
                    normalized_quote
                    and normalized_quote in candidate_text
                    or normalized_value in candidate_text
                ):
                    text = candidate_text
                    source_url = source.url
                    break
        if not text or (normalized_quote not in text and normalized_value not in text):
            continue
        valid[column] = {"value": value, "source": source_url, "evidence": evidence}
    return valid
