"""GD AutoFill — Streamlit application."""

from __future__ import annotations

import time
import re
from collections import Counter
from pathlib import Path

import streamlit as st

from category_rules import detect_category
from config import load_settings
from excel_writer import (
    add_report_sheets,
    inspect_workbook,
    workbook_bytes,
    writable_columns,
    write_values,
)
from firecrawl_client import FirecrawlClient
from parsers import (
    ParsedSource,
    SourceFetcher,
    extract_locally,
    extract_with_gemini,
    fetch_source,
    source_matches_product,
    validate_extraction,
)
from search_engine import SearchEngine

st.set_page_config(page_title="GD AutoFill", page_icon="⚙️", layout="wide")

st.markdown(
    """
<style>
    .block-container {max-width: 1120px; padding-top: 2rem;}
    [data-testid="stFileUploader"] {border: 1px dashed #94a3b8; border-radius: 14px; padding: 12px;}
    .gd-card {background:#f8fafc;border:1px solid #e2e8f0;border-radius:14px;padding:18px;}
    .gd-muted {color:#64748b;font-size:.92rem;}
</style>
""",
    unsafe_allow_html=True,
)


def _product_tokens(product_name: str) -> list[str]:
    return [
        token
        for token in re.findall(r"[a-zа-яё0-9]+", product_name.lower())
        if len(token) >= 2
    ]


def _snippet_sources(product_name: str, results) -> list[ParsedSource]:
    """Use search snippets as weak but useful fallback evidence."""
    output = []
    for result in results:
        text = " ".join(part for part in (result.title, result.snippet) if part).strip()
        if len(text) < 20:
            continue
        tokens = _product_tokens(product_name)
        if tokens and not any(token in text.lower() for token in tokens):
            continue
        output.append(
            ParsedSource(
                result.url,
                "открыт",
                title=result.title,
                text=f"{result.title}\n{result.snippet}\nURL: {result.url}",
                provider=f"{result.provider} snippet",
            )
        )
    return output


def _manual_source_for_product(product_name: str, manual_text: str) -> ParsedSource | None:
    """Pick a user-pasted source block for the current product."""
    manual_text = (manual_text or "").strip()
    if len(manual_text) < 20:
        return None
    lower = manual_text.lower()
    tokens = _product_tokens(product_name)
    if tokens and not any(token in lower for token in tokens):
        return None

    chunks = re.split(r"\n\s*(?:={3,}|-{3,}|#{2,})\s*\n", manual_text)
    chosen = manual_text
    for chunk in chunks:
        chunk_lower = chunk.lower()
        if tokens and all(token in chunk_lower for token in tokens[:2]):
            chosen = chunk
            break

    return ParsedSource(
        f"manual://{product_name}",
        "открыт",
        title=f"Ручной источник: {product_name}",
        text=chosen[:80_000],
        provider="Ручной текст",
    )


def process_file(
    uploaded_file,
    test_mode: bool,
    progress,
    status,
    *,
    existing_bytes: bytes | None = None,
    start_position: int = 0,
    batch_size: int | None = None,
    saved_report_rows: list | None = None,
    saved_review_rows: list | None = None,
    saved_source_rows: list | None = None,
    manual_sources_text: str = "",
):
    settings = load_settings()
    workbook, layout = inspect_workbook(existing_bytes or uploaded_file.getvalue())
    all_rows = layout.data_rows[:3] if test_mode else layout.data_rows
    end_position = (
        len(all_rows)
        if batch_size is None
        else min(len(all_rows), start_position + batch_size)
    )
    rows = all_rows[start_position:end_position]
    engine = SearchEngine(settings)
    fetcher = SourceFetcher(settings)
    firecrawl = FirecrawlClient(settings)
    report_rows = list(saved_report_rows or [])
    review_rows = list(saved_review_rows or [])
    source_rows = list(saved_source_rows or [])

    for position, row in enumerate(rows, start=start_position + 1):
        started = time.monotonic()
        product_name = str(layout.sheet.cell(row, layout.name_column).value).strip()
        columns = writable_columns(layout, row)
        category, confidence = detect_category(
            product_name,
            layout.headers,
            layout.sheet.title,
            uploaded_file.name,
        )
        status.write(f"**{position}/{len(all_rows)}** · {product_name} · {category}")

        results = engine.search_product(product_name, category)
        parsed = _snippet_sources(product_name, results)
        manual_source = _manual_source_for_product(product_name, manual_sources_text)
        if manual_source:
            parsed.insert(0, manual_source)
        if settings.firecrawl_api_key:
            firecrawl_candidates = results[: settings.max_firecrawl_urls_per_product]
            for result in firecrawl_candidates:
                if time.monotonic() - started > settings.product_time_budget:
                    break
                parsed.append(firecrawl.scrape(result))
        else:
            for result in results[: settings.max_pages_per_product]:
                if time.monotonic() - started > settings.product_time_budget:
                    break
                parsed.append(fetch_source(result, settings, fetcher))

        usable = [
            source
            for source in parsed
            if source.status == "открыт" and source_matches_product(product_name, source)
        ]
        extracted = {}
        gemini_error = ""
        if columns and usable:
            try:
                extracted = extract_with_gemini(
                    product_name, category, columns, usable, settings
                )
            except Exception as exc:
                gemini_error = str(exc)[:500]
                extracted = {}
            if not extracted:
                extracted = extract_locally(columns, usable)

        values = validate_extraction(extracted, columns, usable)
        changed = write_values(layout, row, values)
        used_urls = {item["source"] for item in values.values()}
        for source in parsed:
            source.used = source.url in used_urls
            source_rows.append(
                [
                    row,
                    product_name,
                    source.url,
                    source.provider,
                    source.status,
                    "да" if source.used else "нет",
                    source.error,
                ]
            )

        unresolved = [column for column in columns if column not in values]
        for column in unresolved:
            review_rows.append(
                [
                    row,
                    product_name,
                    column,
                    "Нет подтверждённого значения в доступных источниках",
                    "",
                ]
            )
        if not results:
            review_rows.append(
                [row, product_name, "Все поля", "Поиск не вернул разрешённых источников", ""]
            )
        if gemini_error:
            review_rows.append(
                [
                    row,
                    product_name,
                    "Gemini API",
                    "ИИ-извлечение не выполнилось; использован ограниченный локальный разбор",
                    gemini_error,
                ]
            )
        captcha_urls = [
            source.url for source in parsed if source.status == "капча — ручная проверка"
        ]
        if captcha_urls:
            review_rows.append(
                [
                    row,
                    product_name,
                    "Источник",
                    "CAPTCHA: откройте URL вручную; автоматические запросы к домену остановлены",
                    " | ".join(captcha_urls),
                ]
            )

        skipped = sum(source.status != "открыт" for source in parsed)
        report_rows.append(
            [
                row,
                product_name,
                category,
                f"{confidence:.0%}",
                changed,
                len(results),
                len(usable),
                skipped,
                len(unresolved),
            ]
        )
        progress.progress(
            position / len(all_rows),
            text=f"Обработано {position} из {len(all_rows)}",
        )

    add_report_sheets(workbook, report_rows, review_rows, source_rows)
    return (
        workbook_bytes(workbook),
        report_rows,
        review_rows,
        source_rows,
        end_position,
        len(all_rows),
    )


st.title("GD AutoFill")
st.caption("Заполнение характеристик товаров по подтверждённым внешним источникам")

left, right = st.columns([1.7, 1])
with left:
    uploaded = st.file_uploader(
        "Загрузите Excel-шаблон GlobalDrive",
        type=["xlsx"],
        help=(
            "Поддерживаются массовые выгрузки и формы нового товара. "
            "В форме нового товара сначала замените пример в B2 точным названием."
        ),
    )
with right:
    st.markdown(
        """
<div class="gd-card">
<b>Что будет в результате</b><br>
<span class="gd-muted">Исходный Excel с заполненными пустыми полями и листами
«Отчёт», «Проверить», «Источники». Новые значения выделяются жёлтым.</span>
</div>
""",
        unsafe_allow_html=True,
    )

mode = st.radio(
    "Режим обработки",
    ["Полный файл — все товары", "Тест — первые 3 товара"],
    index=0,
    horizontal=True,
)
st.caption(
    "Если это пустая форма нового товара, замените в B2 текст-пример "
    "«БРЕНД Модель» на точное название товара. Служебный столбец A не изменяется."
)

with st.expander("Подключения и безопасность"):
    settings = load_settings()
    cols = st.columns(4)
    cols[0].metric("Serper", "подключён" if settings.serper_api_key else "не задан")
    cols[1].metric("Bing", "подключён" if settings.bing_api_key else "не задан")
    cols[2].metric("Firecrawl", "подключён" if settings.firecrawl_api_key else "не задан")
    cols[3].metric("Gemini", "подключён" if settings.gemini_api_key else "не задан")
    st.caption(
        "Без Serper/Bing используется DuckDuckGo. Без Gemini включается более "
        "ограниченный локальный разбор. Домены из blacklist никогда не используются."
    )
    st.caption(
        "Запросы выполняются с паузами и охлаждением каждого домена. Обычные cookies "
        "сохраняются только внутри текущей обработки. CAPTCHA не обходится: домен "
        "ставится на паузу, а ссылка попадает в лист «Проверить»."
    )

manual_sources_text = st.text_area(
    "Ручные источники, если сайт открылся только после CAPTCHA",
    height=180,
    placeholder=(
        "Не обязательно. Если сайт просит CAPTCHA, откройте его вручную, скопируйте блок "
        "характеристик и вставьте сюда. Для нескольких товаров разделяйте блоки строкой ---.\n\n"
        "Пример:\n"
        "VOGE DS800 Rally\n"
        "Объём двигателя: 798 см3\n"
        "Мощность: 95 л.с.\n"
        "Передний тормоз: два диска 310 мм\n"
        "---\n"
        "Honda CB400F\n"
        "Объём двигателя: 399 см3"
    ),
)

if "result" not in st.session_state:
    st.session_state.result = None
    st.session_state.summary = None
    st.session_state.filename = None
if "job" not in st.session_state:
    st.session_state.job = None



BATCH_SIZE_FULL = 5
BATCH_SIZE_TEST = 3
AUTO_CONTINUE_DELAY_SECONDS = 1.5

start_clicked = st.button(
    "Заполнить Excel",
    type="primary",
    disabled=uploaded is None or bool(st.session_state.job),
    use_container_width=True,
)

should_process = False

if start_clicked:
    batch_size = BATCH_SIZE_TEST if mode.startswith("Тест") else BATCH_SIZE_FULL
    st.session_state.job = {
        "original": uploaded.getvalue(),
        "current": None,
        "name": uploaded.name,
        "test_mode": mode.startswith("Тест"),
        "next": 0,
        "total": None,
        "batch_size": batch_size,
        "auto_continue": True,
        "manual_sources_text": manual_sources_text,
        "report_rows": [],
        "review_rows": [],
        "source_rows": [],
    }
    st.session_state.result = None
    st.session_state.summary = None
    st.session_state.filename = None
    should_process = True

if st.session_state.job:
    job = st.session_state.job
    processed = job["next"]
    total = job.get("total")
    batch_size = job.get("batch_size", BATCH_SIZE_FULL)
    auto_continue = job.get("auto_continue", True)

    if total:
        st.info(
            f"Готово {processed} из {total}. "
            f"Программа сама идёт пачками по {batch_size} товаров."
        )
    else:
        st.info(
            f"Начинаю заполнение. "
            f"Для устойчивости файл сохраняется после каждых {batch_size} товаров."
        )

    action_left, action_right = st.columns([3, 1])
    with action_left:
        continue_clicked = st.button(
            f"Продолжить вручную ? ещё до {batch_size} товаров",
            type="secondary",
            use_container_width=True,
        )
    with action_right:
        stop_clicked = st.button("Остановить", use_container_width=True)

    if stop_clicked:
        job["auto_continue"] = False
        st.warning("Автопродолжение остановлено. Можно скачать уже готовую часть или нажать ?Продолжить?." )

    should_process = should_process or continue_clicked

if st.session_state.job and should_process:
    job = st.session_state.job
    batch_size = job.get("batch_size", BATCH_SIZE_FULL)
    progress = st.progress(0, text=f"Ищу и заполняю до {batch_size} товаров?")
    status = st.empty()

    class StoredUpload:
        name = job["name"]

        def getvalue(self):
            return job["original"]

    try:
        (
            result,
            report_rows,
            review_rows,
            source_rows,
            next_position,
            total,
        ) = process_file(
            StoredUpload(),
            job["test_mode"],
            progress,
            status,
            existing_bytes=job["current"],
            start_position=job["next"],
            batch_size=batch_size,
            saved_report_rows=job["report_rows"],
            saved_review_rows=job["review_rows"],
            saved_source_rows=job["source_rows"],
            manual_sources_text=job.get("manual_sources_text", ""),
        )
        job.update(
            current=result,
            next=next_position,
            total=total,
            report_rows=report_rows,
            review_rows=review_rows,
            source_rows=source_rows,
        )
        st.session_state.result = result
        st.session_state.summary = report_rows
        st.session_state.filename = f"{Path(job['name']).stem}_filled.xlsx"
        if next_position >= total:
            st.session_state.job = None
            status.success("Готово. Все товары обработаны. Скачайте Excel ниже.")
        else:
            status.success(
                f"Сохранено {next_position} из {total}. "
                "Сейчас автоматически перейду к следующей пачке."
            )
            if job.get("auto_continue", True):
                time.sleep(AUTO_CONTINUE_DELAY_SECONDS)
                st.rerun()
    except Exception as exc:
        job["auto_continue"] = False
        status.error(
            "Обработка остановилась, но уже готовая часть сохранена. "
            f"Причина: {exc}. Можно скачать готовую часть или нажать ?Продолжить?."
        )

if st.session_state.result:
    summary = st.session_state.summary or []
    total_filled = sum(row[4] for row in summary)
    total_review = sum(row[8] for row in summary)
    categories = Counter(row[2] for row in summary)
    c1, c2, c3 = st.columns(3)
    c1.metric("Товаров обработано", len(summary))
    c2.metric("Полей заполнено", total_filled)
    c3.metric("Полей на проверку", total_review)
    if categories:
        st.caption(
            "Категории: "
            + ", ".join(f"{name} — {count}" for name, count in categories.items())
        )
    st.download_button(
        "Скачать готовый Excel",
        data=st.session_state.result,
        file_name=st.session_state.filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
