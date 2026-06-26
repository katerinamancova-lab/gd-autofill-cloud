"""GD AutoFill — Streamlit application."""

from __future__ import annotations

import time
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
from parsers import (
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
        parsed = []
        for result in results[: settings.max_pages_per_product]:
            if time.monotonic() - started > 115:
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
    cols = st.columns(3)
    cols[0].metric("Serper", "подключён" if settings.serper_api_key else "не задан")
    cols[1].metric("Bing", "подключён" if settings.bing_api_key else "не задан")
    cols[2].metric("Gemini", "подключён" if settings.gemini_api_key else "не задан")
    st.caption(
        "Без Serper/Bing используется DuckDuckGo. Без Gemini включается более "
        "ограниченный локальный разбор. Домены из blacklist никогда не используются."
    )
    st.caption(
        "Запросы выполняются с паузами и охлаждением каждого домена. Обычные cookies "
        "сохраняются только внутри текущей обработки. CAPTCHA не обходится: домен "
        "ставится на паузу, а ссылка попадает в лист «Проверить»."
    )

if "result" not in st.session_state:
    st.session_state.result = None
    st.session_state.summary = None
    st.session_state.filename = None
if "job" not in st.session_state:
    st.session_state.job = None

start_clicked = st.button(
    "????????? Excel",
    type="primary",
    disabled=uploaded is None or bool(st.session_state.job),
    use_container_width=True,
)

should_process = False

if start_clicked:
    st.session_state.job = {
        "original": uploaded.getvalue(),
        "current": None,
        "name": uploaded.name,
        "test_mode": mode.startswith("????"),
        "next": 0,
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
    status = st.empty()
    processed = job["next"]
    status.info(
        f"????????? ???????: {processed}. "
        "??????? ?????? ????, ????? ?????????? ????????? ?????."
    )

    action_left, action_right = st.columns([3, 1])
    with action_left:
        continue_clicked = st.button(
            "?????????? ? ?????????? ????????? ?????",
            type="primary",
            use_container_width=True,
        )
    with action_right:
        reset_clicked = st.button("?????? ??????", use_container_width=True)

    if reset_clicked:
        st.session_state.job = None
        st.session_state.result = None
        st.session_state.summary = None
        st.session_state.filename = None
        st.rerun()

    should_process = should_process or continue_clicked

if st.session_state.job and should_process:
    job = st.session_state.job
    progress = st.progress(0, text="??????????? ???? ??????")
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
            batch_size=1,
            saved_report_rows=job["report_rows"],
            saved_review_rows=job["review_rows"],
            saved_source_rows=job["source_rows"],
        )
        job.update(
            current=result,
            next=next_position,
            report_rows=report_rows,
            review_rows=review_rows,
            source_rows=source_rows,
        )
        st.session_state.result = result
        st.session_state.summary = report_rows
        st.session_state.filename = f"{Path(job['name']).stem}_filled.xlsx"
        if next_position >= total:
            st.session_state.job = None
            status.success("??????. ??? ?????? ??????????.")
        else:
            status.success(
                f"????? {next_position} ?? {total} ????????. "
                "????? ??????? ????????????? Excel ??? ?????????? ?????????."
            )
    except Exception as exc:
        status.error(
            "????????? ????????????, ?? ??? ??????? ????? ?????????. "
            f"???????: {exc}. ????? ????????? ????????? ????? ??? ?????? ??????."
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
