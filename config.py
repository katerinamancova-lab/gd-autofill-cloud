"""Configuration for GD AutoFill."""

from __future__ import annotations

import os
from dataclasses import dataclass, field


BLACKLIST = {
    "globaldrive.ru",
    "more-motorov-spb.ru",
    "spb.menstechnic.ru",
    "nordkit.ru",
    "mot-motor.ru",
    "moskva.x-tehnika.ru",
    "murmansk.activattor.ru",
    "lodka-motor.com",
}

PROTECTED_COLUMN_MARKERS = {
    "uid",
    "уид",
    "активность",
    "розничная цена",
    "retail price",
}

REPORT_SHEETS = {
    "Отчёт",
    "Проверить",
    "Источники",
    # Backward compatibility with older builds that had mojibake sheet names.
    "РћС‚С‡С‘С‚",
    "РџСЂРѕРІРµСЂРёС‚СЊ",
    "РСЃС‚РѕС‡РЅРёРєРё",
}


@dataclass(frozen=True)
class Settings:
    serper_api_key: str = field(default_factory=lambda: os.getenv("SERPER_API_KEY", ""))
    bing_api_key: str = field(default_factory=lambda: os.getenv("BING_API_KEY", ""))
    gemini_api_key: str = field(default_factory=lambda: os.getenv("GEMINI_API_KEY", ""))
    firecrawl_api_key: str = field(default_factory=lambda: os.getenv("FIRECRAWL_API_KEY", ""))
    gemini_model: str = field(
        default_factory=lambda: os.getenv("GEMINI_MODEL", "gemini-2.5-flash")
    )
    search_timeout: int = 30
    page_timeout: int = 20
    max_results_per_product: int = 40
    max_pages_per_product: int = 8
    max_firecrawl_urls_per_product: int = 8
    firecrawl_timeout: int = 45
    max_source_chars: int = 100_000
    max_gemini_source_chars: int = 100_000
    max_gemini_total_chars: int = 220_000
    min_page_text: int = 180
    request_delay_min: float = 3.0
    request_delay_max: float = 8.0
    domain_cooldown: float = 20.0
    blocked_domain_cooldown: float = 900.0
    product_time_budget: int = 180
    user_agent: str = (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36"
    )


def load_settings() -> Settings:
    """Load settings from environment variables and Streamlit secrets."""
    values: dict[str, str] = {}
    try:
        import streamlit as st

        for key in (
            "SERPER_API_KEY",
            "BING_API_KEY",
            "GEMINI_API_KEY",
            "GEMINI_MODEL",
            "FIRECRAWL_API_KEY",
        ):
            if key in st.secrets:
                values[key] = str(st.secrets[key])
    except Exception:
        pass

    for key, value in values.items():
        os.environ[key] = value
    return Settings()
