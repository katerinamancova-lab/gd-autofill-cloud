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
    "voge", "qjmotor", "royalenfield", "ktm", "hondaset", "mymotors",
    "lodki-piter", "vodomotorika"
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
    "мотоцикл": ["bajaj", "boxer", "benda", "darkflag", "groza", "brz", "yx140", "voge", "qjmotor", "ktm", "honda cb", "gold wing", "мотоцикл", "royal enfield", "benelli", "stels"],
    "гольфкар": ["greencamel", "сонора", "u10k", "гольфкар", "2+2", "4x4"],
    "снегоход": ["снегоход", "snowmobile"],
    "мотобуксировщик": ["мотобуксировщик", "мотособака", "буксировщик"],
}

SERVICE_DEFAULTS = {
    "Комплектация": '<ul><li style="font-size: 13pt; font-family: Acrom">Мотоцикл</li><li style="font-size: 13pt; font-family: Acrom">Сервисная книжка</li></ul>',
    "Гарантия на товар": '<span style="font-family: Acrom; font-size: 13pt;">Гарантия на товар составляет 1 год</span>',
    "Скидка": 11,
    "Доступное количество": 1000,
    "Нет в продаже": None,
    "Сортировка": 500,
    "Привязка к аксессуарам (новая)": "Шлем кроссовый Sharmax SH536 Red/Black;Шлем кроссовый Sharmax SH336 Blue/Black;Мотозащита Sharmax черепаха RT 8;Очки кроссовые Sharmax Premium Black;Очки кроссовые Sharmax Gray/Black;Наколенники Sharmax пластик KP 48 Красные;Наколенники Sharmax SH-32K;Мотоперчатки Sharmax GL-SH 47 White ;Мотоперчатки Sharmax GL-SH 48 Yellow;Мотоперчатки Sharmax GL-SH 49 Red\n",
}

MOTO_KNOWN = {
    "voge ds800": {"cc": 798, "hp": 94, "brand": "VOGE", "brand_country": "Китай", "made": "Китай", "type": "Тур-эндуро", "cooling": "Жидкостное", "fuel_supply": "Инжектор"},
    "honda cb400": {"cc": 399, "hp": 46, "brand": "Honda", "brand_country": "Япония", "made": "Япония", "type": "Дорожный", "cooling": "Жидкостное", "fuel_supply": "Карбюратор"},
    "gold wing": {"cc": 1833, "hp": 126, "brand": "Honda", "brand_country": "Япония", "made": "Япония", "type": "Туристический", "cooling": "Жидкостное", "fuel_supply": "Инжектор"},
    "srk 921": {"cc": 921, "hp": 129, "brand": "QJMotor", "brand_country": "Китай", "made": "Китай", "type": "Нэйкед", "cooling": "Жидкостное", "fuel_supply": "Инжектор"},
    "srk 800": {"cc": 799, "hp": 95, "brand": "QJMotor", "brand_country": "Китай", "made": "Китай", "type": "Нэйкед", "cooling": "Жидкостное", "fuel_supply": "Инжектор"},
    "srk 600": {"cc": 554, "hp": 61, "brand": "QJMotor", "brand_country": "Китай", "made": "Китай", "type": "Нэйкед", "cooling": "Жидкостное", "fuel_supply": "Инжектор"},
    "srk 550": {"cc": 554, "hp": 56, "brand": "QJMotor", "brand_country": "Китай", "made": "Китай", "type": "Нэйкед", "cooling": "Жидкостное", "fuel_supply": "Инжектор"},
    "srv 400": {"cc": 385, "hp": 41, "brand": "QJMotor", "brand_country": "Китай", "made": "Китай", "type": "Классический", "cooling": "Жидкостное", "fuel_supply": "Инжектор"},
    "srv 550": {"cc": 554, "hp": 61, "brand": "QJMotor", "brand_country": "Китай", "made": "Китай", "type": "Классический", "cooling": "Жидкостное", "fuel_supply": "Инжектор"},
    "srk 125": {"cc": 125, "hp": 15, "brand": "QJMotor", "brand_country": "Китай", "made": "Китай", "type": "Нэйкед", "cooling": "Жидкостное", "fuel_supply": "Инжектор"},
    "pulsar ns400": {"cc": 373, "hp": 40, "brand": "Bajaj", "brand_country": "Индия", "made": "Индия", "type": "Нэйкед", "cooling": "Жидкостное", "fuel_supply": "Инжектор"},
    "ss400": {"cc": 373, "hp": 40, "brand": "Bajaj", "brand_country": "Индия", "made": "Индия", "type": "Спортивный", "cooling": "Жидкостное", "fuel_supply": "Инжектор"},
    "c4 300": {"cc": 300, "hp": 26, "brand": "M1NSK", "brand_country": "Беларусь", "made": "Беларусь", "type": "Дорожный", "cooling": "Воздушное", "fuel_supply": "Карбюратор"},
    "d4 125": {"cc": 125, "hp": 11, "brand": "M1NSK", "brand_country": "Беларусь", "made": "Беларусь", "type": "Дорожный", "cooling": "Воздушное", "fuel_supply": "Карбюратор"},
    "390 adventure": {"cc": 373, "hp": 43, "brand": "KTM", "brand_country": "Австрия", "made": "Индия", "type": "Тур-эндуро", "cooling": "Жидкостное", "fuel_supply": "Инжектор"},
    "390 duke": {"cc": 373, "hp": 43, "brand": "KTM", "brand_country": "Австрия", "made": "Индия", "type": "Нэйкед", "cooling": "Жидкостное", "fuel_supply": "Инжектор"},
    "himalayan 450": {"cc": 452, "hp": 40, "brand": "Royal Enfield", "brand_country": "Индия", "made": "Индия", "type": "Тур-эндуро", "cooling": "Жидкостное", "fuel_supply": "Инжектор"},
    "guerrilla 450": {"cc": 452, "hp": 40, "brand": "Royal Enfield", "brand_country": "Индия", "made": "Индия", "type": "Нэйкед", "cooling": "Жидкостное", "fuel_supply": "Инжектор"},
    "752 s": {"cc": 754, "hp": 76, "brand": "Benelli", "brand_country": "Италия", "made": "Китай", "type": "Нэйкед", "cooling": "Жидкостное", "fuel_supply": "Инжектор"},
    "hunter 350": {"cc": 349, "hp": 20, "brand": "Royal Enfield", "brand_country": "Индия", "made": "Индия", "type": "Классический", "cooling": "Воздушное", "fuel_supply": "Инжектор"},
    "interceptor": {"cc": 648, "hp": 47, "brand": "Royal Enfield", "brand_country": "Индия", "made": "Индия", "type": "Классический", "cooling": "Воздушно-масляное", "fuel_supply": "Инжектор"},
    "continental gt": {"cc": 648, "hp": 47, "brand": "Royal Enfield", "brand_country": "Индия", "made": "Индия", "type": "Классический", "cooling": "Воздушно-масляное", "fuel_supply": "Инжектор"},
    "rk125": {"cc": 125, "hp": 12, "brand": "Stels", "brand_country": "Россия", "made": "Китай", "type": "Дорожный", "cooling": "Воздушное", "fuel_supply": "Карбюратор"},
    "m502n": {"cc": 500, "hp": 47, "brand": "Stels", "brand_country": "Россия", "made": "Китай", "type": "Нэйкед", "cooling": "Жидкостное", "fuel_supply": "Инжектор"},
    "monster plus 125": {"cc": 125, "hp": 11, "brand": "Vento", "brand_country": "Китай", "made": "Китай", "type": "Дорожный", "cooling": "Воздушное", "fuel_supply": "Карбюратор"},
    "pulsar as 150": {"cc": 150, "hp": 17, "brand": "Bajaj", "brand_country": "Индия", "made": "Индия", "type": "Дорожный", "cooling": "Воздушное", "fuel_supply": "Карбюратор"},
    "pulsar as 200": {"cc": 199, "hp": 24, "brand": "Bajaj", "brand_country": "Индия", "made": "Индия", "type": "Дорожный", "cooling": "Жидкостное", "fuel_supply": "Карбюратор"},
    "boxer bm125": {"cc": 125, "hp": 10, "brand": "Bajaj", "brand_country": "Индия", "made": "Индия", "type": "Дорожный", "cooling": "Воздушное", "fuel_supply": "Карбюратор"},
    "v 150": {"cc": 150, "hp": 12, "brand": "Bajaj", "brand_country": "Индия", "made": "Индия", "type": "Классический", "cooling": "Воздушное", "fuel_supply": "Карбюратор"},
}

BRAND_COUNTRIES = {
    "Honda": ("Япония", "Япония"),
    "Tohatsu": ("Япония", "Япония"),
    "Suzuki": ("Япония", "Япония"),
    "Yamaha": ("Япония", "Япония"),
    "Mercury": ("США", "Китай"),
    "Hidea": ("Китай", "Китай"),
    "HDX": ("Китай", "Китай"),
    "Parsun": ("Китай", "Китай"),
    "Sea-Pro": ("Китай", "Китай"),
    "Seanovo": ("Китай", "Китай"),
    "Reef Rider": ("Китай", "Китай"),
}


def get_secret(name: str) -> str:
    try:
        value = st.secrets.get(name, "")
    except Exception:
        value = os.getenv(name, "")
    return str(value or "").strip()


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
        if h and not any(x in str(h).lower() for x in ["uid", "уид", "активность", "цена"])
    ]

    prompt = f"""
Ты профессиональный контент-менеджер интернет-магазина техники.
Нужно заполнить Excel-карточку товара.

Товар: {product_name}
Категория: {category}

Колонки Excel, которые можно заполнять:
{json.dumps(safe_headers, ensure_ascii=False)}

Текст источников из интернета:
{source_text[:18000]}

Верни только JSON. Без markdown и без пояснений.

Правила:
1. Ключи JSON должны точно совпадать с колонками Excel.
2. Не заполняй УИД, UID, Активность, Розничная цена.
3. Числа пиши с точкой, не с запятой.
4. Если есть точное значение в источнике — заполни.
5. Если не уверен в точной цифре — лучше пропусти.
6. Для служебных полей можно использовать:
   - Комплектация: Мотоцикл + Сервисная книжка
   - Гарантия на товар: Гарантия на товар составляет 1 год
   - Скидка: 11
   - Доступное количество: 1000
   - Сортировка: 500
"""

    try:
        model = genai.GenerativeModel("gemini-flash-latest")
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
    return any(x in h for x in ["uid", "уид", "активность", "розничная цена"])


def is_manual_value(value):
    return str(value or "").strip().lower() == "заполните вручную в админке"


def range_cc_moto(cc):
    try:
        cc = float(cc)
    except Exception:
        return ""
    if cc <= 199:
        return "до 199"
    if cc <= 300:
        return "200 - 300"
    if cc <= 600:
        return "301 - 600"
    if cc <= 800:
        return "601 - 800"
    return "от 801"


def range_hp_moto(hp):
    try:
        hp = float(hp)
    except Exception:
        return ""
    if hp <= 8:
        return "2 - 8"
    if hp <= 15:
        return "9 - 15"
    if hp <= 25:
        return "16 - 25"
    if hp <= 40:
        return "26 - 40"
    if hp <= 60:
        return "41 - 60"
    if hp <= 100:
        return "61 - 100"
    return "Более 100"


def range_hp_motor(hp):
    try:
        hp = float(hp)
    except Exception:
        return ""
    if hp <= 3.9:
        return "до 3.9"
    if hp <= 6.9:
        return "4 - 6.9"
    if hp <= 9.8:
        return "7 - 9.8"
    if hp <= 20:
        return "9.9 - 20"
    if hp <= 39:
        return "21 - 39"
    if hp <= 59:
        return "40 - 59"
    if hp <= 79:
        return "60 - 79"
    if hp <= 130:
        return "80 - 130"
    if hp <= 150:
        return "131 - 150"
    return "Более 151"


def brand_from_name(name):
    p = name.lower()
    for b in ["Honda", "Mercury", "Tohatsu", "Suzuki", "Yamaha", "Hidea", "HDX", "Parsun", "Sea-Pro", "Seanovo", "Reef Rider"]:
        if b.lower() in p:
            return b
    return ""


def hp_motor_from_name(name):
    p = name.lower()
    patterns = [
        r"bf\s*([0-9]{2,3})",
        r"f\s*([0-9]{2,3})",
        r"mfs\s*([0-9]{2,3})",
        r"df\s*([0-9]{2,3})",
        r"hdef\s*([0-9]{2,3})",
        r"hd\s*([0-9]{2,3})",
        r"t\s*([0-9]{2,3})",
        r"([0-9]{2,3})\s*л\.?с",
    ]
    for pat in patterns:
        m = re.search(pat, p)
        if m:
            val = int(m.group(1))
            if 2 <= val <= 350:
                return val
    return None


def rule_spec(product_name, category, headers):
    spec = {}
    p = product_name.lower()

    # Служебные поля заполняем всегда, если такие колонки есть
    for k, v in SERVICE_DEFAULTS.items():
        if k in headers:
            spec[k] = v

    if category == "мотоцикл":
        data = None
        for key, val in MOTO_KNOWN.items():
            if key in p:
                data = val
                break

        if data:
            spec["Бренд [BRAND]"] = data.get("brand")
            spec["Объём двигателя, куб [ENGINE_CAPACITY]"] = str(data.get("cc"))
            spec["Объём двигателя (по диапазонам) [ENGINE_CAPACITY1]"] = range_cc_moto(data.get("cc"))
            spec["Мощность, л.с. [POWER_HP]"] = str(data.get("hp"))
            spec["Мощность (по диапазонам) [POWER_HP1]"] = range_hp_moto(data.get("hp"))
            spec["Тип мотоцикла [Motorcycle_Type]"] = data.get("type")
            spec["Охлаждение [COOLING]"] = data.get("cooling")
            spec["Система подачи топлива [Fuel_supply_system]"] = data.get("fuel_supply")
            spec["Страна бренда [BRAND_COUNTRY]"] = data.get("brand_country")
            spec["Страна производства [MANUFACTURER]"] = data.get("made")

        spec.setdefault("Гарантия [WARRANTY]", "1 год")
        spec.setdefault("Наличие ПТС [NALICHIE_PTS]", "Да")
        spec.setdefault("Система запуска [STARTING_SYSTEM]", "Электростартер")
        spec.setdefault("Материал рамы [FRAME_MATERIAL]", "Сталь")
        spec.setdefault("Тип топлива [TYPE_FUEL]", "АИ92-95")
        spec.setdefault("Количество тактов [STROKE]", "4")
        spec.setdefault("Тип двигателя [TYPE_ENGINE]", "Бензиновый")
        spec.setdefault("Трансмиссия [TRANSMISSION]", "Механическая")
        spec.setdefault("Двигатель [ENGINE]", "-")
        spec.setdefault("Количество цилиндров [CYLINDERS]", "1")

    elif category == "лодочный мотор":
        brand = brand_from_name(product_name)
        if brand:
            spec["Бренд [BRAND]"] = brand
            bc, made = BRAND_COUNTRIES.get(brand, ("", ""))
            if bc:
                spec["Страна бренда [BRAND_COUNTRY]"] = bc
            if made:
                spec["Страна производства [MANUFACTURER]"] = made

        hp = hp_motor_from_name(product_name)
        if hp:
            spec["Мощность, л.с. [POWER_HP1]"] = str(hp)
            spec["Мощность (л.с.) [POWER_HP]"] = range_hp_motor(hp)
            spec["Мощность (кВт) [POWER_KW]"] = str(round(hp * 0.7355, 1))

        spec.setdefault("Управление [OPERATION]", "Румпельное" if any(x in p for x in ["fh", "румп", "hes"]) else "Дистанционное")
        spec.setdefault("Система запуска [STARTING_SYSTEM]", "Электростартер" if any(x in p for x in ["e", "etl", "elpt", "xrtu", "efi"]) else "Ручной стартер/электростартер")
        spec.setdefault("Тип насадки [NOZZLETYPE]", "Водомёт" if "jet" in p or "водом" in p else "Винт")
        spec.setdefault("Система подачи топлива [Fuel_supply_system]", "Инжектор" if "efi" in p else "Карбюратор")
        spec.setdefault("Система подъёма [LIFTING_SYSTEM]", "Гидравлическая" if any(x in p for x in ["pt", "trim", "xrtu", "elpt", "btx"]) else "Ручная")
        spec.setdefault("Количество тактов [STROKE]", "2" if "2 такт" in p or "2-такт" in p or "t 40" in p else "4")
        spec.setdefault("Охлаждение [COOLING]", "Водяное")
        spec.setdefault("Тип двигателя [TYPE_ENGINE]", "Бензиновый")
        spec.setdefault("Передачи [GEAR]", "F-N-R")
        spec.setdefault("Тип топлива [TYPE_FUEL]", "АИ92-95")
        spec.setdefault("Вращение винта [ROTATION_SCREW]", "Водомётная насадка" if spec.get("Тип насадки [NOZZLETYPE]") == "Водомёт" else "Правое")
        spec.setdefault("Гарантия [WARRANTY]", "5 лет" if brand in ["Honda", "Tohatsu"] else "3 года")

    return {k: v for k, v in spec.items() if v is not None}


def process_excel(uploaded_file, category_mode: str, max_products: int):
    wb = load_workbook(uploaded_file)
    ws = wb.active

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
    rule_ok = 0

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

        spec = rule_spec(product_name, category, headers)
        if spec:
            rule_ok += 1
            log_ws.append([row_num, product_name, f"Правила: подготовлено полей {len(spec)}"])

        ai_spec, status = call_gemini(product_name, category, headers, source_text)

        if status == "ok":
            ai_ok += 1
            spec.update(ai_spec)
            log_ws.append([row_num, product_name, f"Gemini: получено полей {len(ai_spec)}"])
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
        ("Правила сработали", rule_ok),
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
