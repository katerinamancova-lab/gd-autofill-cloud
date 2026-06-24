import io
import os
import re
import json
import time
from urllib.parse import urlparse

import requests
import streamlit as st
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from dotenv import load_dotenv

load_dotenv()

try:
    import google.generativeai as genai
except Exception:
    genai = None


YELLOW = PatternFill(fill_type="solid", fgColor="FFFF00")

BLACKLIST_DOMAINS = [
    "globaldrive.ru",
    "more-motorov-spb.ru",
    "spb.menstechnic.ru",
    "nordkit.ru",
    "mot-motor.ru",
    "moskva.x-tehnika.ru",
    "murmansk.activattor.ru",
    "lodka-motor.com",
]

BAD_DOMAINS = [
    "avito", "ozon", "wildberries", "youtube", "vk.com", "dzen",
    "instagram", "pinterest", "cart", "login", "compare", "forum",
    "drive2", "market.yandex", "maps.yandex", "2gis", "wikipedia"
]

HEADERS = {
    "User-Agent": "Mozilla/5.0 Chrome/124 Safari/537.36",
    "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.8",
}


def get_secret(name: str) -> str:
    try:
        value = st.secrets.get(name, "")
    except Exception:
        value = os.getenv(name, "")
    return str(value or "").strip()


def is_blocked_domain(url: str) -> bool:
    u = (url or "").lower()
    return any(d in u for d in BLACKLIST_DOMAINS)


def is_bad_url(url: str) -> bool:
    u = (url or "").lower()
    return is_blocked_domain(u) or any(d in u for d in BAD_DOMAINS)


def clean_name(name: str) -> str:
    s = str(name or "").strip()
    s = re.sub(r"\([^)]*(цвет|red|black|blue|green|white|желт|красн|черн|син|бел|202[0-9])[^)]*\)", " ", s, flags=re.I)
    s = re.sub(r"\b(новый|new|202[0-9]|год|цвет|красный|черный|чёрный|синий|белый|желтый|зелёный|зеленый)\b", " ", s, flags=re.I)
    return re.sub(r"\s+", " ", s).strip() or str(name or "").strip()


def header_map(ws):
    return {str(ws.cell(1, c).value or "").strip(): c for c in range(1, ws.max_column + 1)}


def set_cell(ws, hmap, row, header, value):
    if value in (None, ""):
        return 0
    if header not in hmap:
        return 0
    if any(x in header.lower() for x in ["uid", "уид", "активность", "розничная цена"]):
        return 0

    cell = ws.cell(row, hmap[header])
    if str(cell.value or "").strip() == str(value or "").strip():
        return 0

    cell.value = value
    cell.fill = YELLOW
    return 1


def detect_category(headers, names):
    joined = " ".join(str(h or "").lower() for h in headers)
    n = " ".join(names).lower()

    if any(x in joined for x in ["дейдвуд", "вращение винта", "тип насадки"]):
        return "лодочный мотор"
    if any(x in joined for x in ["тип днища", "плотность материала", "диаметр борта"]):
        return "лодка пвх"
    if any(x in joined for x in ["наличие псм", "лебедка", "тип привода"]):
        return "квадроцикл"
    if any(x in joined for x in ["тип мотоцикла", "наличие птс"]):
        return "мотоцикл"
    if any(x in joined for x in ["пассажировместимость", "мощность (вт)", "запас хода"]):
        return "гольфкар"

    if any(x in n for x in ["bf", "mfs", "tohatsu", "mercury", "hidea", "parsun", "suzuki df"]):
        return "лодочный мотор"
    if any(x in n for x in ["пвх", "лодка", "нднд", "airdeck"]):
        return "лодка пвх"
    if any(x in n for x in ["квадроцикл", "atv", "outlander", "sprmotors", "blade"]):
        return "квадроцикл"
    if any(x in n for x in ["гольфкар", "greencamel"]):
        return "гольфкар"
    return "мотоцикл"


def blacklist_query_part() -> str:
    return " ".join(f"-site:{d}" for d in BLACKLIST_DOMAINS)


def build_queries(product, category):
    name = clean_name(product)
    minus = blacklist_query_part()
    base = [
        f'"{name}" характеристики {minus}',
        f'"{name}" технические характеристики {minus}',
        f'"{name}" specs {minus}',
        f'{name} {category} характеристики {minus}',
        f'{name} паспорт характеристики {minus}',
        f'{name} инструкция характеристики {minus}',
    ]
    return base


def search_serper(query, api_key):
    if not api_key:
        return [], [], "SERPER_API_KEY не найден"
    try:
        r = requests.post(
            "https://google.serper.dev/search",
            json={"q": query, "gl": "ru", "hl": "ru", "num": 10},
            headers={"X-API-KEY": api_key, "Content-Type": "application/json"},
            timeout=20,
        )
        r.raise_for_status()
        data = r.json()
    except Exception as e:
        return [], [], f"Serper ошибка: {e}"

    urls, snippets = [], []
    for item in data.get("organic", []):
        url = item.get("link") or ""
        if not url.startswith("http") or is_bad_url(url):
            continue
        title = item.get("title", "")
        snippet = item.get("snippet", "")
        urls.append(url)
        snippets.append(f"{title}. {snippet}")
    return urls, snippets, "ok"


def fetch_page_text(url):
    if is_bad_url(url):
        return "", "blacklist/bad"
    try:
        r = requests.get(url, headers=HEADERS, timeout=8)
        r.raise_for_status()
    except Exception as e:
        return "", f"не открылось: {e}"

    try:
        soup = BeautifulSoup(r.text or "", "lxml")

        meta = []
        for m in soup.find_all("meta"):
            c = m.get("content")
            if c:
                meta.append(c)

        tables = []
        for tr in soup.select("tr"):
            cells = [c.get_text(" ", strip=True) for c in tr.find_all(["td", "th"])]
            cells = [c for c in cells if c]
            if len(cells) >= 2:
                tables.append(" : ".join(cells))

        for tag in soup(["script", "style", "noscript", "svg", "footer", "nav"]):
            tag.decompose()

        visible = soup.get_text(" ", strip=True)
        text = re.sub(r"\s+", " ", " ".join(meta + tables + [visible])).strip()

        low = text.lower()
        if len(text) < 250 and any(x in low for x in ["captcha", "access denied", "докажите", "робот"]):
            return "", "капча/антибот"

        return text[:10000], "ok" if len(text) > 100 else "мало текста"
    except Exception as e:
        return "", f"ошибка чтения: {e}"


def collect_sources(product, category, serper_key, max_pages=5):
    all_urls, all_snippets, logs = [], [], []

    for q in build_queries(product, category):
        urls, snippets, status = search_serper(q, serper_key)
        logs.append(f"Serper: {q} | ссылок {len(urls)} | {status}")
        all_urls.extend(urls)
        all_snippets.extend(snippets)

    unique = []
    seen = set()
    for u in all_urls:
        if u not in seen and not is_bad_url(u):
            seen.add(u)
            unique.append(u)

    text = "\n".join(all_snippets[:30])
    opened = []

    for url in unique[:max_pages]:
        page_text, status = fetch_page_text(url)
        opened.append((url, status))
        if page_text:
            text += "\n\n" + page_text

    return text, opened, logs


def extract_json(text):
    if not text:
        return {}
    text = text.strip().replace("```json", "").replace("```", "").strip()
    try:
        obj = json.loads(text)
        return obj if isinstance(obj, dict) else {}
    except Exception:
        pass

    m = re.search(r"\{.*\}", text, re.S)
    if m:
        try:
            obj = json.loads(m.group(0))
            return obj if isinstance(obj, dict) else {}
        except Exception:
            return {}
    return {}


def gemini_extract(product, category, headers, source_text, gemini_key):
    if not gemini_key or genai is None or not source_text.strip():
        return {}, "Gemini недоступен или нет источников"

    safe_headers = [
        h for h in headers
        if h and not any(x in h.lower() for x in ["uid", "уид", "активность", "розничная цена"])
    ]

    genai.configure(api_key=gemini_key)

    prompt = f"""
Ты контент-менеджер каталога техники.
Нужно заполнить Excel строго по найденным источникам.

Товар: {product}
Категория: {category}

Колонки Excel:
{json.dumps(safe_headers, ensure_ascii=False)}

Источники:
{source_text[:22000]}

Верни только JSON.
Ключи JSON должны ТОЧНО совпадать с колонками Excel.
Заполняй только точные характеристики из источников.
Если характеристики нет — пропусти.
Не заполняй UID, УИД, Активность, Розничная цена.
"""

    for model in ["gemini-2.0-flash", "gemini-flash-latest"]:
        try:
            response = genai.GenerativeModel(model).generate_content(
                prompt,
                generation_config={"temperature": 0, "response_mime_type": "application/json"},
            )
            data = extract_json(response.text)
            if data:
                return data, f"ok {model}"
        except Exception as e:
            last = str(e)
    return {}, f"Gemini ошибка: {last if 'last' in locals() else ''}"


def basic_parse(source_text, category):
    spec = {}
    text = source_text or ""
    low = text.lower()

    patterns = {
        "Длина, см [LENGTH_CM]": [r"длина[^0-9]{0,40}(\d{3})"],
        "Ширина, см [WIDTH]": [r"ширина[^0-9]{0,40}(\d{2,3})"],
        "Вес, кг [WEIGHT]": [r"(?:вес|масса)[^0-9]{0,40}(\d{1,3}[,.]?\d*)"],
        "Сухой вес, кг [DRY_WEIGHT]": [r"(?:сухой вес|вес лодки|масса)[^0-9]{0,40}(\d{1,3}[,.]?\d*)"],
        "Грузоподъемность, кг [LOAD_CAPACITY]": [r"грузопод[ъье]мность[^0-9]{0,40}(\d{2,4})"],
        "Макс. мощность мотора, л.с. [MAX_POWER]": [r"(?:макс\.?\s*мощность|мощность мотора)[^0-9]{0,40}(\d{1,3})"],
        "Диаметр борта, см [BOAT_SIDE_DIAMETER]": [r"(?:диаметр баллона|диаметр борта)[^0-9]{0,40}(\d{2})"],
        "Объём двигателя, куб [ENGINE_CAPACITY]": [r"(?:объем|объём) двигателя[^0-9]{0,40}(\d{2,4})"],
        "Мощность, л.с. [POWER_HP]": [r"мощность[^0-9]{0,40}(\d{1,3})\s*(?:л\.с|лс)"],
    }

    for header, pats in patterns.items():
        for pat in pats:
            m = re.search(pat, low)
            if m:
                spec[header] = m.group(1).replace(",", ".")
                break

    if category == "лодка пвх":
        spec.setdefault("Тип лодки [TYPE_BOAT]", "Под мотор")
        if "нднд" in low or "надувное дно низкого давления" in low:
            spec["Тип днища [TYPE_BOTTOM]"] = "Надувное, низкого давления"
        elif "airdeck" in low:
            spec["Тип днища [TYPE_BOTTOM]"] = "Надувное, высокого давления"

    return spec


def process_excel(uploaded, mode):
    wb = load_workbook(uploaded)
    ws = wb.active
    hmap = header_map(ws)
    headers = list(hmap.keys())

    for s in ["Отчет", "Проверить", "Источники"]:
        if s in wb.sheetnames:
            del wb[s]

    report = wb.create_sheet("Отчет")
    report.append(["Показатель", "Значение"])
    check = wb.create_sheet("Проверить")
    check.append(["Строка", "Товар", "Поле", "Значение", "Комментарий"])
    sources_ws = wb.create_sheet("Источники")
    sources_ws.append(["Строка", "Товар", "URL", "Статус"])

    rows = []
    for r in range(2, ws.max_row + 1):
        name = str(ws.cell(r, 1).value or "").strip()
        if name:
            rows.append((r, name))

    if mode == "Тест: первые 3 товара":
        rows = rows[:3]

    serper_key = get_secret("SERPER_API_KEY")
    gemini_key = get_secret("GEMINI_API_KEY")
    category = detect_category(headers, [n for _, n in rows[:5]])

    progress = st.progress(0)
    changed = 0
    ai_ok = 0
    opened_ok = 0

    for idx, (row_num, product) in enumerate(rows, start=1):
        row_category = detect_category(headers, [product])

        source_text, opened, logs = collect_sources(product, row_category, serper_key)

        for log in logs[:10]:
            check.append([row_num, product, "Поиск", "", log])

        for url, status in opened:
            sources_ws.append([row_num, product, url, status])
            if status == "ok":
                opened_ok += 1

        spec = basic_parse(source_text, row_category)

        ai_spec, status = gemini_extract(product, row_category, headers, source_text, gemini_key)
        if ai_spec:
            spec.update(ai_spec)
            ai_ok += 1
        else:
            check.append([row_num, product, "Gemini", "", status])

        row_changed = 0
        for h, v in spec.items():
            row_changed += set_cell(ws, hmap, row_num, h, v)

        changed += row_changed

        if row_changed == 0:
            check.append([row_num, product, "Заполнение", "", "Не удалось найти точные характеристики по разрешенным сайтам"])

        progress.progress(idx / max(len(rows), 1))
        time.sleep(0.2)

    for row in [
        ["Обработано товаров", len(rows)],
        ["Категория файла", category],
        ["Открыто источников", opened_ok],
        ["Gemini успешно", ai_ok],
        ["Изменено ячеек", changed],
        ["Исключенные сайты", ", ".join(BLACKLIST_DOMAINS)],
    ]:
        report.append(row)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


st.set_page_config(page_title="GD AutoFill Cloud Content v3", layout="centered")
st.title("GD AutoFill Cloud Content v3")
st.write("Контентщик загружает Excel → программа ищет характеристики вне ваших сайтов → отдаёт заполненный файл.")

st.info(f"Serper API: {'✅ найден' if get_secret('SERPER_API_KEY') else '❌ не найден'}")
st.info(f"Gemini API: {'✅ найден' if get_secret('GEMINI_API_KEY') else '❌ не найден'}")

mode = st.radio("Режим", ["Тест: первые 3 товара", "Полный файл"], index=0)
uploaded = st.file_uploader("Загрузите Excel", type=["xlsx"])

if uploaded:
    st.success(f"Файл загружен: {uploaded.name}")

    if st.button("Заполнить и скачать"):
        with st.spinner("Ищу характеристики и заполняю файл..."):
            try:
                result = process_excel(uploaded, mode)
                st.success("Готово")
                st.download_button(
                    "Скачать заполненный Excel",
                    data=result,
                    file_name=uploaded.name.replace(".xlsx", "_CLOUD_CONTENT_v3.xlsx"),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error(f"Ошибка: {e}")
