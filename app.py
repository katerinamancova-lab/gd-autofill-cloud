import os
import time
import streamlit as st
from dotenv import load_dotenv

from search_engine import find_sources, fetch_text
from parsers import parse_basic_specs, gemini_extract
from excel_writer import prepare_workbook, header_map, set_if_col, export_workbook
from category_rules import detect_category, base_rules

load_dotenv()


def get_secret(name: str) -> str:
    try:
        value = st.secrets.get(name, "")
    except Exception:
        value = os.getenv(name, "")
    return str(value or "").strip()


st.set_page_config(page_title="GD AutoFill Search Engine v1", layout="centered")
st.title("GD AutoFill Search Engine v1")
st.write("Ищет характеристики в интернете, исключая ваши сайты, анализирует источники и заполняет Excel.")

serper_key = get_secret("SERPER_API_KEY")
gemini_key = get_secret("GEMINI_API_KEY")

st.info(f"Serper API: {'✅ найден' if serper_key else '❌ не найден'}")
st.info(f"Gemini API: {'✅ найден' if gemini_key else '❌ не найден'}")

mode = st.radio("Режим", ["Быстрый тест 3 товара", "Полный файл"], index=0)
use_gemini = st.checkbox("Использовать Gemini для анализа найденных источников", value=True)

uploaded = st.file_uploader("Загрузите Excel", type=["xlsx"])

if uploaded:
    st.success(f"Файл загружен: {uploaded.name}")

    if st.button("Заполнить и скачать"):
        wb, ws, report, check, sources = prepare_workbook(uploaded)
        hmap = header_map(ws)
        headers = list(hmap.keys())

        rows = []
        for r in range(2, ws.max_row + 1):
            name = str(ws.cell(r, 1).value or "").strip()
            if name:
                rows.append((r, name))

        if mode == "Быстрый тест 3 товара":
            rows = rows[:3]

        category = detect_category(headers, [n for _, n in rows[:5]])
        changed_total = 0
        source_ok = 0
        ai_ok = 0
        ai_fail = 0

        progress = st.progress(0)

        for idx, (row_num, product_name) in enumerate(rows, start=1):
            row_category = detect_category(headers, [product_name])
            spec = base_rules(product_name, row_category)

            urls, snippets, logs = find_sources(product_name, row_category, serper_key)
            source_text = "\\n".join(snippets)

            for log in logs[:12]:
                check.append([row_num, product_name, "Поиск", "", log])

            for url in urls:
                text, status = fetch_text(url)
                sources.append([row_num, product_name, url, status])
                if text:
                    source_ok += 1
                    source_text += "\\n\\n" + text

            parsed = parse_basic_specs(source_text, row_category)
            spec.update(parsed)

            if use_gemini and source_text.strip():
                ai_spec, status = gemini_extract(product_name, row_category, headers, source_text, gemini_key)
                if ai_spec:
                    spec.update(ai_spec)
                    ai_ok += 1
                else:
                    ai_fail += 1
                    check.append([row_num, product_name, "Gemini", "", status])

            row_changed = 0
            for h, v in spec.items():
                row_changed += set_if_col(ws, hmap, row_num, h, v)

            changed_total += row_changed
            if row_changed == 0:
                check.append([row_num, product_name, "Заполнение", "", "Не изменилось: нет точных данных или нет колонок"])

            progress.progress(idx / max(len(rows), 1))
            time.sleep(0.5)

        for row in [
            ["Категория файла", category],
            ["Обработано товаров", len(rows)],
            ["Источники открылись", source_ok],
            ["Gemini успешно", ai_ok],
            ["Gemini ошибки", ai_fail],
            ["Изменено ячеек", changed_total],
            ["Исключены сайты", "globaldrive.ru, more-motorov-spb.ru, spb.menstechnic.ru, nordkit.ru, mot-motor.ru, moskva.x-tehnika.ru, murmansk.activattor.ru, lodka-motor.com"],
        ]:
            report.append(row)

        result = export_workbook(wb)
        st.success("Готово")
        st.download_button(
            "Скачать заполненный Excel",
            data=result,
            file_name=uploaded.name.replace(".xlsx", "_SEARCH_ENGINE_v1.xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
