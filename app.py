import io
import os
import re
import json
import time

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


def get_secret(name):
    try:
        return str(st.secrets.get(name, "") or "").strip()
    except Exception:
        return str(os.getenv(name, "") or "").strip()


def is_bad_url(url):
    u = (url or "").lower()
    return any(d in u for d in BLACKLIST_DOMAINS) or any(d in u for d in BAD_DOMAINS)


def header_map(ws):
    return {str(ws.cell(1, c).value or "").strip(): c for c in range(1, ws.max_column + 1)}


def set_cell(ws, hmap, row, header, value):
    if value in ("", None):
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


def clean_name(name):
    s = str(name or "").strip()
    s = re.sub(r"\([^)]*(цвет|red|black|blue|green|white|желт|красн|черн|син|бел|202[0-9])[^)]*\)", " ", s, flags=re.I)
    s = re.sub(r"\b(новый|new|202[0-9]|год|цвет|красный|черный|чёрный|синий|белый|желтый|зелёный|зеленый)\b", " ", s, flags=re.I)
    return re.sub(r"\s+", " ", s).strip() or str(name or "").strip()


def detect_category(headers, names):
    joined = " ".join(str(h or "").lower() for h in headers)
    names_text = " ".join(names).lower()

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

    if any(x in names_text for x in ["bf", "mfs", "tohatsu", "mercury", "hidea", "parsun", "suzuki df"]):
        return "лодочный мотор"
    if any(x in names_text for x in ["пвх", "лодка", "нднд", "airdeck"]):
        return "лодка пвх"
    if any(x in names_text for x in ["квадроцикл", "atv", "outlander", "sprmotors", "blade"]):
        return "квадроцикл"
    if any(x in names_text for x in ["гольфкар", "greencamel"]):
        return "гольфкар"
    return "мотоцикл"


def minus_sites():
    return " ".join(f"-site:{d}" for d in BLACKLIST_DOMAINS)


def build_queries(product, category):
    name = clean_name(product)
    minus = minus_sites()

    return [
        f'"{name}" характеристики {minus}',
        f'"{name}" технические характеристики {minus}',
        f'"{name}" specs {minus}',
        f'"{name}" specification {minus}',
        f'"{name}" паспорт {minus}',
        f'"{name}" инструкция {minus}',
        f'"{name}" manual pdf {minus}',
        f'"{name}" каталог характеристики {minus}',
        f'{name} {category} характеристики {minus}',
        f'{name} купить характеристики {minus}',
    ]


def serper_search(query, key):
    if not key:
        return [], "нет SERPER_API_KEY"
    try:
        r = requests.post(
            "https://google.serper.dev/search",
            json={"q": query, "gl": "ru", "hl": "ru", "num": 10},
            headers={"X-API-KEY": key, "Content-Type": "application/json"},
            timeout=18,
        )
        r.raise_for_status()
        data = r.json()
    except Exception as e:
        return [], f"ошибка Serper: {e}"

    items = []
    for item in data.get("organic", []):
        url = item.get("link") or ""
        if not url.startswith("http") or is_bad_url(url):
            continue
        items.append({
            "url": url,
            "title": item.get("title", ""),
            "snippet": item.get("snippet", ""),
        })
    return items, "ok"


def fetch_text_fast(url):
    if is_bad_url(url):
        return "", "blacklist"
    try:
        r = requests.get(url, headers=HEADERS, timeout=6)
        r.raise_for_status()
    except Exception as e:
        return "", f"не открылось: {e}"

    try:
        soup = BeautifulSoup(r.text or "", "lxml")
        meta = [m.get("content") for m in soup.find_all("meta") if m.get("content")]

        tables = []
        for tr in soup.select("tr"):
            cells = [c.get_text(" ", strip=True) for c in tr.find_all(["td", "th"])]
            cells = [c for c in cells if c]
            if len(cells) >= 2:
                tables.append(" : ".join(cells))

        for tag in soup(["script", "style", "noscript", "svg", "footer", "nav"]):
            tag.decompose()

        text = " ".join(meta + tables + [soup.get_text(" ", strip=True)])
        text = re.sub(r"\s+", " ", text).strip()

        low = text.lower()
        if len(text) < 300 and any(x in low for x in ["captcha", "access denied", "робот", "докажите"]):
            return "", "капча/антибот"

        if len(text) < 150:
            return "", "мало текста"

        return text[:7000], "ok"
    except Exception as e:
        return "", f"ошибка чтения: {e}"


def collect_sources(product, category, serper_key):
    source_parts = []
    source_rows = []
    logs = []
    seen_urls = set()

    for q in build_queries(product, category):
        items, status = serper_search(q, serper_key)
        logs.append(f"{q} | результатов: {len(items)} | {status}")

        for item in items:
            url = item["url"]
            if url in seen_urls:
                continue
            seen_urls.add(url)

            # Главное: даже если сайт не откроется, title/snippet из Google уже используем.
            title = item.get("title", "")
            snippet = item.get("snippet", "")
            source_parts.append(f"Источник выдачи: {title}. {snippet}. URL: {url}")

            source_rows.append([url, "snippet"])

    opened_count = 0
    # Открываем только первые 6 сайтов, чтобы приложение не зависало.
    for url in list(seen_urls)[:6]:
        txt, status = fetch_text_fast(url)
        source_rows.append([url, status])
        if txt:
            opened_count += 1
            source_parts.append(f"Текст страницы {url}: {txt}")

    return "\n\n".join(source_parts)[:28000], source_rows, logs, opened_count


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


def regex_extract(source_text, category):
    spec = {}
    low = (source_text or "").lower()

    patterns = {
        "Длина, см [LENGTH_CM]": r"длина[^0-9]{0,60}(\d{3})",
        "Ширина, см [WIDTH]": r"ширина[^0-9]{0,60}(\d{2,3})",
        "Высота, см [HEIGHT]": r"высота[^0-9]{0,60}(\d{2,3})",
        "Вес, кг [WEIGHT]": r"(?:вес|масса)[^0-9]{0,60}(\d{1,3}[,.]?\d*)",
        "Сухой вес, кг [DRY_WEIGHT]": r"(?:сухой вес|вес лодки|масса)[^0-9]{0,60}(\d{1,3}[,.]?\d*)",
        "Грузоподъемность, кг [LOAD_CAPACITY]": r"грузопод[ъье]мность[^0-9]{0,60}(\d{2,4})",
        "Макс. мощность мотора, л.с. [MAX_POWER]": r"(?:макс\.?\s*мощность|мощность мотора)[^0-9]{0,60}(\d{1,3})",
        "Диаметр борта, см [BOAT_SIDE_DIAMETER]": r"(?:диаметр баллона|диаметр борта)[^0-9]{0,60}(\d{2})",
        "Пассажировместимость, чел [PASSENGER]": r"(?:пассажировместимость|пассажиров|мест)[^0-9]{0,60}(\d{1,2})",
        "Объём двигателя, куб [ENGINE_CAPACITY]": r"(?:объем|объём) двигателя[^0-9]{0,60}(\d{2,4})",
        "Мощность, л.с. [POWER_HP]": r"мощность[^0-9]{0,60}(\d{1,3})\s*(?:л\.с|лс)",
    }

    for h, pat in patterns.items():
        m = re.search(pat, low)
        if m:
            spec[h] = m.group(1).replace(",", ".")

    if category == "лодка пвх":
        spec.setdefault("Тип лодки [TYPE_BOAT]", "Под мотор")
        if "нднд" in low or "надувное дно низкого давления" in low:
            spec["Тип днища [TYPE_BOTTOM]"] = "Надувное, низкого давления"
        elif "airdeck" in low:
            spec["Тип днища [TYPE_BOTTOM]"] = "Надувное, высокого давления"

    return spec


def gemini_schema_extract(product, category, headers, source_text, key):
    if not key or genai is None:
        return {}, "нет Gemini"
    if not source_text.strip():
        return {}, "нет источников"

    safe_headers = [
        h for h in headers
        if h and not any(x in h.lower() for x in ["uid", "уид", "активность", "розничная цена"])
    ]

    prompt = f"""
Ты заполняешь карточку товара в Excel.
Заполняй только по найденным источникам из Google/сайтов. Не выдумывай.

Товар: {product}
Категория: {category}

Колонки Excel:
{json.dumps(safe_headers, ensure_ascii=False)}

Найденные источники:
{source_text[:26000]}

Верни JSON.
Ключи — ТОЧНО такие же названия колонок Excel.
Заполняй максимум точных характеристик.
Если характеристики нет в источниках — не добавляй ключ.
Не заполняй UID, УИД, Активность, Розничная цена.
"""

    genai.configure(api_key=key)

    for model_name in ["gemini-2.0-flash", "gemini-flash-latest"]:
        try:
            model = genai.GenerativeModel(model_name)
            resp = model.generate_content(
                prompt,
                generation_config={"temperature": 0, "response_mime_type": "application/json"},
            )
            data = extract_json(resp.text)
            if data:
                return data, f"ok {model_name}"
        except Exception as e:
            last = str(e)

    return {}, f"Gemini ошибка: {last if 'last' in locals() else ''}"


def process(uploaded, mode):
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

    src_ws = wb.create_sheet("Источники")
    src_ws.append(["Строка", "Товар", "URL", "Статус"])

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
    source_snippets = 0

    for i, (row_num, product) in enumerate(rows, start=1):
        row_category = detect_category(headers, [product])
        source_text, source_rows, logs, opened_count = collect_sources(product, row_category, serper_key)

        opened_ok += opened_count

        for log in logs[:10]:
            check.append([row_num, product, "Поиск", "", log])

        for url, status in source_rows:
            src_ws.append([row_num, product, url, status])
            if status == "snippet":
                source_snippets += 1

        spec = regex_extract(source_text, row_category)

        ai_spec, status = gemini_schema_extract(product, row_category, headers, source_text, gemini_key)
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
            check.append([row_num, product, "Заполнение", "", "Нет точных данных по разрешенным источникам"])

        progress.progress(i / max(len(rows), 1))
        time.sleep(0.15)

    for row in [
        ["Категория файла", category],
        ["Обработано товаров", len(rows)],
        ["Сниппетов из Google", source_snippets],
        ["Открыто сайтов", opened_ok],
        ["Gemini успешно", ai_ok],
        ["Изменено ячеек", changed],
        ["Принцип v5", "не ждёт капчу: использует выдачу Google + открытые сайты + Gemini"],
        ["Исключены сайты", ", ".join(BLACKLIST_DOMAINS)],
    ]:
        report.append(row)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


st.set_page_config(page_title="GD AutoFill Cloud v5 Final", layout="centered")
st.title("GD AutoFill Cloud v5 Final")
st.write("Версия для работы по ссылке: ищет вне ваших сайтов, не зависает на капче, использует выдачу Google + открытые сайты + Gemini.")

st.info(f"Serper API: {'✅ найден' if get_secret('SERPER_API_KEY') else '❌ не найден'}")
st.info(f"Gemini API: {'✅ найден' if get_secret('GEMINI_API_KEY') else '❌ не найден'}")

mode = st.radio("Режим", ["Тест: первые 3 товара", "Полный файл"], index=0)
uploaded = st.file_uploader("Загрузите Excel", type=["xlsx"])

if uploaded:
    st.success(f"Файл загружен: {uploaded.name}")
    if st.button("Заполнить и скачать"):
        with st.spinner("Ищу характеристики и заполняю файл..."):
            try:
                result = process(uploaded, mode)
                st.success("Готово")
                st.download_button(
                    "Скачать заполненный Excel",
                    data=result,
                    file_name=uploaded.name.replace(".xlsx", "_CLOUD_v5_FINAL.xlsx"),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error(f"Ошибка: {e}")
