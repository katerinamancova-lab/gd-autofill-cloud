
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


BAD_DOMAINS = [
    "avito", "ozon", "wildberries", "youtube", "vk.com", "dzen",
    "instagram", "pinterest", "images", "cart", "login", "compare",
    "forum", "drive2", "market.yandex"
]

TRUSTED_DOMAINS = [
    "globaldrive.ru", "more-motorov-spb.ru", "rollingmoto.ru",
    "motomarine.ru", "vodnik", "honda", "tohatsu", "mercury",
    "suzuki", "yamaha", "hidea", "parsun", "sea-pro", "linhai",
    "greencamel", "brp", "can-am", "segway", "bajaj", "benda",
    "voge", "qjmotor", "royalenfield", "ktm"
]

HEADERS = {
    "User-Agent": "Mozilla/5.0 Chrome/124 Safari/537.36",
    "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.8",
}

YELLOW = PatternFill(fill_type="solid", fgColor="FFFF00")

CATEGORY_KEYWORDS = {
    "лодочный мотор": ["tohatsu", "honda bf", "mercury", "suzuki df", "hidea", "hdx", "parsun", "sea-pro t", "yamaha f", "mfs", "jet"],
    "лодка пвх": ["лодка", "пвх", "нднд", "hunter", "хантер", "байкал", "сапфир", "urex", "access sp", "smarine"],
    "квадроцикл": ["квадроцикл", "atv", "linhai", "segway", "outlander", "aodes", "росомаха", "brp"],
    "мотоцикл": ["bajaj", "boxer", "benda", "darkflag", "groza", "brz", "yx140", "voge", "qjmotor", "ktm", "honda cb", "gold wing", "мотоцикл"],
    "гольфкар": ["greencamel", "сонора", "u10k", "гольфкар", "2+2", "4x4"],
    "снегоход": ["снегоход", "snowmobile"],
    "мотобуксировщик": ["мотобуксировщик", "мотособака", "буксировщик"],
}


def get_secret(name: str) -> str:
    try:
        return st.secrets.get(name, "")
    except Exception:
        return os.getenv(name, "")


def is_bad_url(url: str) -> bool:
    u = (url or "").lower()
    return any(x in u for x in BAD_DOMAINS)


def score_url(url: str, product_name: str) -> int:
    u = (url or "").lower()
    p = (product_name or "").lower()
    if is_bad_url(u):
        return -1000

    score = 0
    if any(d in u for d in TRUSTED_DOMAINS):
        score += 100

    for token in re.findall(r"[a-zA-Zа-яА-Я0-9]+", p):
        if len(token) > 2 and token.lower() in u:
            score += 7

    if any(x in u for x in ["product", "catalog", "character", "spec", "harakter", "kharakteristiki"]):
        score += 15

    return score


def detect_category_by_name(name: str) -> str:
    s = str(name or "").lower()
    scores = {}
    for category, keys in CATEGORY_KEYWORDS.items():
        scores[category] = sum(1 for k in keys if k in s)
    best = max(scores, key=scores.get)
    return best if scores[best] > 0 else "не определено"


def detect_category_by_headers(headers) -> str:
    joined = " ".join(str(h or "").lower() for h in headers)
    if "дейдвуд" in joined or "вращение винта" in joined:
        return "лодочный мотор"
    if "тип днища" in joined or "плотность материала" in joined:
        return "лодка пвх"
    if "наличие псм" in joined or "лебедка" in joined:
        return "квадроцикл"
    if "пассажировместимость" in joined and "мощность (вт)" in joined:
        return "гольфкар"
    if "тип мотоцикла" in joined or "наличие птс" in joined:
        return "мотоцикл"
    return "не определено"


def search_serper(query: str, max_results: int = 10):
    api_key = get_secret("SERPER_API_KEY")
    if not api_key:
        return [], [], "SERPER_API_KEY не найден"

    try:
        r = requests.post(
            "https://google.serper.dev/search",
            json={"q": query, "gl": "ru", "hl": "ru", "num": max_results},
            headers={"X-API-KEY": api_key, "Content-Type": "application/json"},
            timeout=25,
        )
        r.raise_for_status()
        data = r.json()
    except Exception as e:
        return [], [], f"Ошибка Serper: {e}"

    links = []
    snippets = []

    for item in data.get("organic", []):
        link = item.get("link")
        if link and not is_bad_url(link):
            links.append(link)
            snippets.append((item.get("title", "") + " " + item.get("snippet", "")).strip())

    return links, snippets, "ok"


def find_pages(product_name: str, category: str, max_pages: int = 6):
    queries = [
        f'"{product_name}" характеристики',
        f'"{product_name}" технические характеристики',
        f'{product_name} {category} характеристики',
        f'{product_name} specs',
        f'{product_name} site:globaldrive.ru',
        f'{product_name} site:more-motorov-spb.ru',
        f'{product_name} site:rollingmoto.ru',
        f'{product_name} site:motomarine.ru',
    ]

    all_links = []
    all_snippets = []
    log = []

    for q in queries:
        links, snippets, status = search_serper(q)
        all_links.extend(links)
        all_snippets.extend(snippets)
        log.append(f"Запрос: {q} | Найдено: {len(links)} | {status}")

    unique = []
    seen = set()

    for url in all_links:
        if url not in seen:
            seen.add(url)
            unique.append(url)

    ranked = sorted(unique, key=lambda u: score_url(u, product_name), reverse=True)
    return ranked[:max_pages], all_snippets[:25], log


def fetch_text(url: str):
    try:
        r = requests.get(url, headers=HEADERS, timeout=25)
        r.raise_for_status()
    except Exception as e:
        return "", f"не удалось открыть: {e}"

    soup = BeautifulSoup(r.text, "lxml")
    for tag in soup(["script", "style", "noscript", "svg", "footer", "nav"]):
        tag.decompose()

    text = re.sub(r"\s+", " ", soup.get_text(" ", strip=True))
    return text[:8000], "ok"


def extract_json(text: str):
    if not text:
        return {}

    text = text.strip().replace("```json", "").replace("```", "").strip()

    try:
        data = json.loads(text)
        return data if isinstance(data, dict) else {}
    except Exception:
        pass

    match = re.search(r"\{.*\}", text, re.S)
    if match:
        try:
            data = json.loads(match.group(0))
            return data if isinstance(data, dict) else {}
        except Exception:
            return {}

    return {}


def call_gemini(product_name: str, category: str, headers: list[str], source_text: str):
    api_key = get_secret("GEMINI_API_KEY")
    if not api_key:
        return {}, "GEMINI_API_KEY не найден"

    if genai is None:
        return {}, "google-generativeai не установлен"

    genai.configure(api_key=api_key)

    safe_headers = [
        h for h in headers
        if h and not any(x in str(h).lower() for x in ["uid", "уид", "активность", "цена", "количество", "сортировка"])
    ]

    prompt = f"""
Ты профессиональный контент-менеджер интернет-магазина техники.
Нужно заполнить Excel-карточку товара.

Товар:
{product_name}

Категория:
{category}

Колонки Excel, которые можно заполнять:
{json.dumps(safe_headers, ensure_ascii=False)}

Текст источников из интернета:
{source_text[:20000]}

Правила:
1. Верни ТОЛЬКО JSON-объект.
2. Ключи JSON должны точно совпадать с колонками Excel.
3. Не заполняй УИД, UID, Активность, цену, количество, сортировку.
4. Числа пиши с точкой, не с запятой.
5. Если не уверен в точной цифре — не заполняй это поле.
6. Справочные поля выбирай коротко: Инжектор, Карбюратор, Водяное, Жидкостное, Электростартер, Бензиновый и т.д.
7. Заполни максимум характеристик по источникам.
"""

    try:
        model = genai.GenerativeModel("gemini-1.5-flash")
        response = model.generate_content(
            prompt,
            generation_config={
                "temperature": 0,
                "response_mime_type": "application/json",
            },
        )
        data = extract_json(response.text)
        if data:
            return data, "ok"
        return {}, "Gemini не вернул JSON"
    except Exception as e:
        return {}, f"Ошибка Gemini: {e}"


def prop_code(header: str):
    match = re.search(r"\[([A-Za-z0-9_]+)\]", str(header or ""))
    return match.group(1).lower() if match else ""


def norm_header(header: str):
    s = str(header or "").lower()
    s = re.sub(r"\[[^\]]+\]", " ", s)
    s = re.sub(r"[^a-zа-я0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def make_header_maps(ws):
    exact = {}
    codes = {}
    normalized = {}

    for c in range(1, ws.max_column + 1):
        h = str(ws.cell(1, c).value or "").strip()
        exact[h] = c

        code = prop_code(h)
        if code:
            codes[code] = c

        normalized[norm_header(h)] = c

    return exact, codes, normalized


def find_column(header: str, exact, codes, normalized):
    h = str(header or "").strip()

    if h in exact:
        return exact[h]

    code = prop_code(h)
    if code and code in codes:
        return codes[code]

    nh = norm_header(h)
    if nh in normalized:
        return normalized[nh]

    return None


def is_protected_header(header: str):
    h = str(header or "").lower()
    return any(x in h for x in ["uid", "уид", "активность"])


def is_manual_value(value):
    return str(value or "").strip().lower() == "заполните вручную в админке"


def process_excel(uploaded_file, category_mode: str, max_products: int):
    wb = load_workbook(uploaded_file)
    ws = wb.active

    # remove old reports
    for name in ["Отчет", "Источники", "Проверить", "Лог поиска"]:
        if name in wb.sheetnames:
            del wb[name]

    report = wb.create_sheet("Отчет")
    report.append(["Показатель", "Значение"])

    sources_ws = wb.create_sheet("Источники")
    sources_ws.append(["Строка", "Товар", "Категория", "Источник"])

    check_ws = wb.create_sheet("Проверить")
    check_ws.append(["Строка", "Товар", "Поле", "Значение", "Причина"])

    log_ws = wb.create_sheet("Лог поиска")
    log_ws.append(["Строка", "Товар", "Сообщение"])

    exact, codes, normalized = make_header_maps(ws)
    headers = [str(ws.cell(1, c).value or "").strip() for c in range(1, ws.max_column + 1)]

    file_category = detect_category_by_headers(headers)

    rows = []
    for r in range(2, ws.max_row + 1):
        name = str(ws.cell(r, 1).value or "").strip()
        if name:
            rows.append((r, name))

    if max_products:
        rows = rows[:max_products]

    changed = 0
    found_links = 0
    ai_ok = 0
    ai_fail = 0

    progress = st.progress(0)

    for idx, (row_num, product_name) in enumerate(rows, start=1):
        if category_mode == "Авто по строкам":
            category = detect_category_by_name(product_name)
            if category == "не определено":
                category = file_category
        elif category_mode == "Авто":
            category = file_category
        else:
            category = category_mode.lower()

        urls, snippets, logs = find_pages(product_name, category)

        for log in logs:
            log_ws.append([row_num, product_name, log])

        texts = []

        for url in urls:
            sources_ws.append([row_num, product_name, category, url])
            txt, status = fetch_text(url)
            log_ws.append([row_num, product_name, f"Открытие: {url} | {status} | символов: {len(txt)}"])
            if txt:
                texts.append(txt)

        found_links += len(urls)
        source_text = "\n".join(snippets) + "\n\n" + "\n\n".join(texts)

        spec, status = call_gemini(product_name, category, headers, source_text)

        if status == "ok":
            ai_ok += 1
        else:
            ai_fail += 1
            check_ws.append([row_num, product_name, "Gemini", "", status])

        if not spec:
            check_ws.append([row_num, product_name, "Товар", "", "Не удалось извлечь характеристики"])
            progress.progress(idx / len(rows))
            continue

        row_changed = 0

        for header, value in spec.items():
            col = find_column(header, exact, codes, normalized)

            if not col:
                check_ws.append([row_num, product_name, header, value, "Нет такой колонки в шаблоне"])
                continue

            real_header = str(ws.cell(1, col).value or "")

            if is_protected_header(real_header):
                continue

            cell = ws.cell(row_num, col)

            if is_manual_value(cell.value):
                cell.value = None
                continue

            if str(cell.value or "").strip() == str(value or "").strip():
                continue

            cell.value = value
            cell.fill = YELLOW
            changed += 1
            row_changed += 1

        if row_changed == 0:
            check_ws.append([row_num, product_name, "Заполнение", "", "Характеристики найдены, но ячейки не изменились"])

        progress.progress(idx / len(rows))
        time.sleep(0.2)

    report_rows = [
        ("Категория файла", file_category),
        ("Режим", category_mode),
        ("Обработано товаров", len(rows)),
        ("Найдено ссылок", found_links),
        ("Gemini успешно", ai_ok),
        ("Gemini ошибки", ai_fail),
        ("Изменено ячеек", changed),
    ]

    for row in report_rows:
        report.append(row)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


st.set_page_config(page_title="GD AutoFill AI Cloud", layout="centered")
st.title("GD AutoFill AI Cloud")
st.write("Загрузите Excel — программа найдёт характеристики в интернете и заполнит файл через Gemini.")

gemini_ok = bool(get_secret("GEMINI_API_KEY"))
serper_ok = bool(get_secret("SERPER_API_KEY"))

st.info(f"Gemini API: {'✅ найден' if gemini_ok else '❌ не найден'}")
st.info(f"Serper API: {'✅ найден' if serper_ok else '❌ не найден'}")

category_mode = st.selectbox(
    "Категория",
    ["Авто по строкам", "Авто", "лодочный мотор", "лодка пвх", "квадроцикл", "мотоцикл", "гольфкар", "снегоход", "мотобуксировщик"],
)

max_products = st.number_input(
    "Сколько товаров обработать за раз",
    min_value=1,
    max_value=100,
    value=5,
    help="Для бесплатной версии лучше начинать с 3–5 товаров.",
)

with st.expander("Проверить поиск"):
    test_product = st.text_input("Название товара", value="Honda BF 100 XRTU")
    test_category = st.selectbox("Категория теста", ["лодочный мотор", "лодка пвх", "квадроцикл", "мотоцикл", "гольфкар"])
    if st.button("Проверить поиск"):
        urls, snippets, logs = find_pages(test_product, test_category)
        st.write("Лог:")
        for item in logs:
            st.write(item)
        st.write("Ссылки:")
        for url in urls:
            st.write(url)
        st.write("Сниппеты:")
        for snippet in snippets[:5]:
            st.write(snippet)

uploaded = st.file_uploader("Загрузите Excel", type=["xlsx"])

if uploaded:
    st.success(f"Файл загружен: {uploaded.name}")

    if st.button("Заполнить характеристики"):
        if not gemini_ok or not serper_ok:
            st.error("Не найдены GEMINI_API_KEY или SERPER_API_KEY. Добавь их в secrets.")
        else:
            with st.spinner("Ищу характеристики и заполняю Excel..."):
                try:
                    result = process_excel(uploaded, category_mode, max_products)
                    st.success("Готово")
                    st.download_button(
                        "Скачать заполненный Excel",
                        data=result,
                        file_name=uploaded.name.replace(".xlsx", "_GD_AI_CLOUD.xlsx"),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                except Exception as e:
                    st.error(f"Ошибка обработки: {e}")
