import io
import os
import re
import json
import time
import requests
from bs4 import BeautifulSoup
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from dotenv import load_dotenv

load_dotenv()

try:
    import google.generativeai as genai
except Exception:
    genai = None

YELLOW = PatternFill(fill_type="solid", fgColor="FFFF00")

MOTO_ACCESSORIES = (
    "Шлем кроссовый Sharmax SH536 Red/Black;"
    "Шлем кроссовый Sharmax SH336 Blue/Black;"
    "Мотозащита Sharmax черепаха RT 8;"
    "Очки кроссовые Sharmax Premium Black;"
    "Очки кроссовые Sharmax Gray/Black;"
    "Наколенники Sharmax пластик KP 48 Красные;"
    "Наколенники Sharmax SH-32K;"
    "Мотоперчатки Sharmax GL-SH 47 White ;"
    "Мотоперчатки Sharmax GL-SH 48 Yellow;"
    "Мотоперчатки Sharmax GL-SH 49 Red\n"
)

SERVICE_DEFAULTS = {
    "мотоцикл": {
        "Комплектация": '<ul><li style="font-size: 13pt; font-family: Acrom">Мотоцикл</li><li style="font-size: 13pt; font-family: Acrom">Сервисная книжка</li></ul>',
        "Гарантия на товар": '<span style="font-family: Acrom; font-size: 13pt;">Гарантия на товар составляет 1 год</span>',
        "Скидка": 11,
        "Доступное количество": 1000,
        "Нет в продаже": None,
        "Сортировка": 500,
        "Привязка к аксессуарам (новая)": MOTO_ACCESSORIES,
    },
    "лодочный мотор": {
        "Гарантия на товар": '<span style="font-family: Acrom; font-size: 13pt;">Гарантия на товар составляет 1 год</span>',
        "Скидка": 11,
        "Доступное количество": 1000,
        "Нет в продаже": None,
        "Сортировка": 500,
    },
    "лодка пвх": {
        "Гарантия на товар": '<span style="font-family: Acrom; font-size: 13pt;">Гарантия на товар составляет 1 год</span>',
        "Скидка": 11,
        "Доступное количество": 1000,
        "Нет в продаже": None,
        "Сортировка": 500,
    },
    "квадроцикл": {
        "Гарантия на товар": '<span style="font-family: Acrom; font-size: 13pt;">Гарантия на товар составляет 1 год</span>',
        "Скидка": 11,
        "Доступное количество": 1000,
        "Нет в продаже": None,
        "Сортировка": 500,
    },
    "гольфкар": {
        "Гарантия на товар": '<span style="font-family: Acrom; font-size: 13pt;">Гарантия на товар составляет 1 год</span>',
        "Скидка": 11,
        "Доступное количество": 1000,
        "Нет в продаже": None,
        "Сортировка": 500,
    },
}

MOTO_KNOWN = {
    "voge ds800": ("VOGE", 798, "501 - 800", 94, "80 - 99", "Жидкостное", "4-тактный 2-цилиндровый", "Тур-эндуро", "Китай", "Китай", "Инжектор", 2),
    "honda cb400": ("Honda", 399, "301 - 600", 46, "41 - 60", "Жидкостное", "4-тактный 4-цилиндровый", "Дорожный", "Япония", "Япония", "Карбюратор", 4),
    "gold wing": ("Honda", 1833, "от 801", 126, "Более 100", "Жидкостное", "4-тактный 6-цилиндровый", "Туристический", "Япония", "Япония", "Инжектор", 6),
    "srk 921": ("QJMotor", 921, "от 801", 129, "Более 100", "Жидкостное", "4-тактный 4-цилиндровый", "Нэйкед", "Китай", "Китай", "Инжектор", 4),
    "srk 800": ("QJMotor", 799, "501 - 800", 95, "80 - 99", "Жидкостное", "4-тактный 4-цилиндровый", "Спортивный", "Китай", "Китай", "Инжектор", 4),
    "srk 600": ("QJMotor", 554, "301 - 600", 61, "61 - 100", "Жидкостное", "4-тактный 4-цилиндровый", "Нэйкед", "Китай", "Китай", "Инжектор", 4),
    "srk 550": ("QJMotor", 554, "301 - 600", 56, "41 - 60", "Жидкостное", "4-тактный 2-цилиндровый", "Нэйкед", "Китай", "Китай", "Инжектор", 2),
    "srv 400": ("QJMotor", 385, "301 - 600", 41, "41 - 60", "Жидкостное", "4-тактный 2-цилиндровый", "Классический", "Китай", "Китай", "Инжектор", 2),
    "srv 550": ("QJMotor", 554, "301 - 600", 61, "61 - 100", "Жидкостное", "4-тактный 2-цилиндровый", "Классический", "Китай", "Китай", "Инжектор", 2),
    "pulsar ns400": ("Bajaj", 373, "301 - 600", 40, "26 - 40", "Жидкостное", "4-тактный 1-цилиндровый", "Нэйкед", "Индия", "Индия", "Инжектор", 1),
    "ss400": ("Bajaj", 373, "301 - 600", 40, "26 - 40", "Жидкостное", "4-тактный 1-цилиндровый", "Спортивный", "Индия", "Индия", "Инжектор", 1),
    "c4 300": ("M1NSK", 300, "200 - 300", 26, "26 - 40", "Воздушное", "4-тактный 1-цилиндровый", "Дорожный", "Беларусь", "Беларусь", "Карбюратор", 1),
    "d4 125": ("M1NSK", 125, "до 199", 11, "9 - 15", "Воздушное", "4-тактный 1-цилиндровый", "Дорожный", "Беларусь", "Беларусь", "Карбюратор", 1),
    "390 adventure": ("KTM", 373, "301 - 600", 43, "41 - 60", "Жидкостное", "4-тактный 1-цилиндровый", "Тур-эндуро", "Австрия", "Индия", "Инжектор", 1),
    "390 duke": ("KTM", 373, "301 - 600", 43, "41 - 60", "Жидкостное", "4-тактный 1-цилиндровый", "Нэйкед", "Австрия", "Индия", "Инжектор", 1),
    "himalayan 450": ("Royal Enfield", 452, "301 - 600", 40, "26 - 40", "Жидкостное", "4-тактный 1-цилиндровый", "Тур-эндуро", "Индия", "Индия", "Инжектор", 1),
    "guerrilla 450": ("Royal Enfield", 452, "301 - 600", 40, "26 - 40", "Жидкостное", "4-тактный 1-цилиндровый", "Нэйкед", "Индия", "Индия", "Инжектор", 1),
    "benelli 752": ("Benelli", 754, "501 - 800", 76, "61 - 100", "Жидкостное", "4-тактный 2-цилиндровый", "Нэйкед", "Италия", "Китай", "Инжектор", 2),
    "hunter 350": ("Royal Enfield", 349, "301 - 600", 20, "16 - 25", "Воздушное", "4-тактный 1-цилиндровый", "Классический", "Индия", "Индия", "Инжектор", 1),
    "interceptor": ("Royal Enfield", 648, "601 - 800", 47, "41 - 60", "Воздушно-масляное", "4-тактный 2-цилиндровый", "Классический", "Индия", "Индия", "Инжектор", 2),
    "continental gt": ("Royal Enfield", 648, "601 - 800", 47, "41 - 60", "Воздушно-масляное", "4-тактный 2-цилиндровый", "Классический", "Индия", "Индия", "Инжектор", 2),
    "rk125": ("Stels", 125, "до 199", 12, "9 - 15", "Воздушное", "4-тактный 1-цилиндровый", "Дорожный", "Россия", "Китай", "Карбюратор", 1),
    "m502n": ("Stels", 500, "301 - 600", 47, "41 - 60", "Жидкостное", "4-тактный 2-цилиндровый", "Нэйкед", "Россия", "Китай", "Инжектор", 2),
    "monster plus": ("Vento", 125, "до 199", 11, "9 - 15", "Воздушное", "4-тактный 1-цилиндровый", "Дорожный", "Китай", "Китай", "Карбюратор", 1),
    "pulsar as 150": ("Bajaj", 150, "до 199", 17, "16 - 25", "Воздушное", "4-тактный 1-цилиндровый", "Дорожный", "Индия", "Индия", "Карбюратор", 1),
    "pulsar as 200": ("Bajaj", 199, "до 199", 24, "16 - 25", "Жидкостное", "4-тактный 1-цилиндровый", "Дорожный", "Индия", "Индия", "Карбюратор", 1),
    "boxer bm125": ("Bajaj", 125, "до 199", 10, "9 - 15", "Воздушное", "4-тактный 1-цилиндровый", "Дорожный", "Индия", "Индия", "Карбюратор", 1),
    "bajaj v 150": ("Bajaj", 150, "до 199", 12, "9 - 15", "Воздушное", "4-тактный 1-цилиндровый", "Классический", "Индия", "Индия", "Карбюратор", 1),
}

BRAND_DEFAULTS = {
    "honda": ("Honda", "Япония", "Япония"),
    "tohatsu": ("Tohatsu", "Япония", "Япония"),
    "suzuki": ("Suzuki", "Япония", "Япония"),
    "yamaha": ("Yamaha", "Япония", "Япония"),
    "mercury": ("Mercury", "США", "Китай"),
    "hidea": ("Hidea", "Китай", "Китай"),
    "hdx": ("HDX", "Китай", "Китай"),
    "parsun": ("Parsun", "Китай", "Китай"),
    "sea-pro": ("Sea-Pro", "Китай", "Китай"),
    "seanovo": ("Seanovo", "Китай", "Китай"),
    "reef rider": ("Reef Rider", "Китай", "Китай"),
    "bajaj": ("Bajaj", "Индия", "Индия"),
    "voge": ("VOGE", "Китай", "Китай"),
    "qjmotor": ("QJMotor", "Китай", "Китай"),
    "royal enfield": ("Royal Enfield", "Индия", "Индия"),
    "ktm": ("KTM", "Австрия", "Индия"),
    "benelli": ("Benelli", "Италия", "Китай"),
    "stels": ("Stels", "Россия", "Китай"),
    "benda": ("Benda", "Китай", "Китай"),
    "linhai": ("Linhai Yamaha", "Китай", "Китай"),
    "greencamel": ("GreenCamel", "Россия", "Китай"),
}

CATEGORY_KEYS = {
    "лодочный мотор": ["дейдвуд", "вращение винта", "тип насадки", "передачи"],
    "лодка пвх": ["тип днища", "плотность материала", "диаметр борта", "внутренняя длина"],
    "квадроцикл": ["наличие псм", "лебедка", "тип привода", "защита рук", "фаркоп"],
    "гольфкар": ["пассажировместимость", "мощность (вт)", "запас хода", "педаль акселератора"],
    "мотоцикл": ["тип мотоцикла", "наличие птс", "колеса передние", "колеса задние"],
}


def get_secret(name: str) -> str:
    try:
        value = st.secrets.get(name, "")
    except Exception:
        value = os.getenv(name, "")
    return str(value or "").strip()


def header_map(ws):
    return {str(ws.cell(1, c).value or "").strip(): c for c in range(1, ws.max_column + 1)}


def detect_category(headers, first_names):
    joined = " ".join(str(h or "").lower() for h in headers)
    for cat, keys in CATEGORY_KEYS.items():
        if any(k in joined for k in keys):
            return cat
    names = " ".join(first_names).lower()
    if any(x in names for x in ["bf", "mfs", "mercury", "tohatsu", "hidea", "parsun", "sea-pro", "suzuki df"]):
        return "лодочный мотор"
    if any(x in names for x in ["пвх", "лодка", "нднд", "hunter", "байкал", "сапфир"]):
        return "лодка пвх"
    if any(x in names for x in ["atv", "outlander", "segway", "linhai", "квадроцикл"]):
        return "квадроцикл"
    if any(x in names for x in ["greencamel", "гольфкар", "сонора"]):
        return "гольфкар"
    return "мотоцикл"


def set_if_col(ws, hmap, row, header, value):
    if value is None:
        return 0
    col = hmap.get(header)
    if not col:
        return 0
    if any(x in header.lower() for x in ["uid", "уид", "активность", "розничная цена"]):
        return 0
    cell = ws.cell(row, col)
    old = cell.value
    if str(old or "").strip() == str(value or "").strip():
        return 0
    cell.value = value
    cell.fill = YELLOW
    return 1


def extract_json(text):
    if not text:
        return {}
    text = text.strip().replace("```json", "").replace("```", "").strip()
    try:
        data = json.loads(text)
        return data if isinstance(data, dict) else {}
    except Exception:
        pass
    m = re.search(r"\{.*\}", text, re.S)
    if m:
        try:
            data = json.loads(m.group(0))
            return data if isinstance(data, dict) else {}
        except Exception:
            return {}
    return {}



BAD_DOMAINS = [
    "avito", "ozon", "wildberries", "youtube", "vk.com", "dzen",
    "instagram", "pinterest", "images", "cart", "login", "compare",
    "forum", "drive2", "market.yandex", "maps.yandex"
]

TRUSTED_DOMAINS = [
    "globaldrive.ru", "more-motorov-spb.ru", "rollingmoto.ru",
    "motomarine.ru", "honda", "tohatsu", "mercury", "suzuki",
    "yamaha", "hidea", "parsun", "sea-pro", "linhai", "greencamel",
    "brp", "can-am", "segway", "bajaj", "benda", "voge", "qjmotor",
    "royalenfield", "ktm", "benelli", "stels", "mymotors", "hondaset"
]

HTTP_HEADERS = {
    "User-Agent": "Mozilla/5.0 Chrome/124 Safari/537.36",
    "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.8",
}


def is_bad_url(url):
    u = (url or "").lower()
    return any(x in u for x in BAD_DOMAINS)


def score_url(url, product_name):
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


def search_serper(query, max_results=10):
    api_key = get_secret("SERPER_API_KEY")
    if not api_key:
        return [], [], "SERPER_API_KEY не найден"

    try:
        r = requests.post(
            "https://google.serper.dev/search",
            json={"q": query, "gl": "ru", "hl": "ru", "num": max_results},
            headers={"X-API-KEY": api_key, "Content-Type": "application/json"},
            timeout=20,
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


def find_pages(product_name, category, max_pages=5):
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
    logs = []

    for q in queries:
        links, snippets, status = search_serper(q)
        all_links.extend(links)
        all_snippets.extend(snippets)
        logs.append(f"Запрос: {q} | найдено: {len(links)} | {status}")

    unique = []
    seen = set()
    for url in all_links:
        if url not in seen:
            seen.add(url)
            unique.append(url)

    ranked = sorted(unique, key=lambda u: score_url(u, product_name), reverse=True)
    return ranked[:max_pages], all_snippets[:20], logs


def fetch_text(url):
    try:
        r = requests.get(url, headers=HTTP_HEADERS, timeout=15)
        r.raise_for_status()
    except Exception as e:
        return "", f"не открылось: {e}"

    # Если сайт показывает капчу/защиту, честно пишем в лог и идём дальше.
    html = r.text or ""
    low = html.lower()
    if any(x in low for x in ["captcha", "recaptcha", "cloudflare", "access denied", "докажите", "робот"]):
        return "", "сайт просит капчу/антибот, пропущено"

    try:
        soup = BeautifulSoup(html, "lxml")
        for tag in soup(["script", "style", "noscript", "svg", "footer", "nav"]):
            tag.decompose()
        text = re.sub(r"\s+", " ", soup.get_text(" ", strip=True))
        return text[:8000], "ok"
    except Exception as e:
        return "", f"ошибка чтения страницы: {e}"



def gemini_by_name(product_name, headers, category, source_text=""):
    api_key = get_secret("GEMINI_API_KEY")
    if not api_key or genai is None:
        return {}, "Gemini недоступен"

    genai.configure(api_key=api_key)
    safe_headers = [
        h for h in headers
        if h and not any(x in h.lower() for x in ["uid", "уид", "активность", "розничная цена"])
    ]

    source_part = ""
    if source_text and source_text.strip():
        source_part = f"""
Найденные источники из интернета:
{source_text[:18000]}

Используй источники как главный источник данных.
"""
    else:
        source_part = """
Источники не открылись или сайт заблокировал доступ.
Заполни по названию товара, категории и технической логике.
Если точную цифру не знаешь — пропусти.
"""

    prompt = f"""
Верни только JSON для заполнения Excel.
Товар: {product_name}
Категория: {category}

Колонки Excel:
{json.dumps(safe_headers, ensure_ascii=False)}

{source_part}

Правила:
- ключи JSON должны точно совпадать с колонками;
- не заполняй УИД, UID, Активность, Розничная цена;
- если точную цифру не знаешь — пропусти;
- бренды/страны/тип двигателя/топливо/запуск можно определить логически;
- заполни максимум характеристик.
"""

    for model_name in ["gemini-2.0-flash", "gemini-flash-latest", "gemini-2.5-flash"]:
        try:
            model = genai.GenerativeModel(model_name)
            response = model.generate_content(
                prompt,
                generation_config={"temperature": 0, "response_mime_type": "application/json"},
            )
            data = extract_json(response.text)
            if data:
                return data, f"ok {model_name}"
        except Exception as e:
            last = str(e)
            continue
    return {}, f"Gemini не сработал, файл всё равно будет создан. Ошибка: {last if 'last' in locals() else ''}"


def range_hp_motor(hp):
    try: hp = float(hp)
    except Exception: return ""
    if hp <= 3.9: return "до 3.9"
    if hp <= 6.9: return "4 - 6.9"
    if hp <= 9.8: return "7 - 9.8"
    if hp <= 20: return "9.9 - 20"
    if hp <= 39: return "21 - 39"
    if hp <= 59: return "40 - 59"
    if hp <= 79: return "60 - 79"
    if hp <= 130: return "80 - 130"
    if hp <= 150: return "131 - 150"
    return "Более 151"


def hp_from_name(name):
    p = name.lower()
    for pat in [r"bf\s*([0-9]{2,3})", r"f\s*([0-9]{2,3})", r"mfs\s*([0-9]{2,3})", r"df\s*([0-9]{2,3})", r"hdef\s*([0-9]{2,3})", r"hd\s*([0-9]{2,3})", r"t\s*([0-9]{2,3})", r"([0-9]{2,3})\s*л\.?с"]:
        m = re.search(pat, p)
        if m:
            v = int(m.group(1))
            if 2 <= v <= 350:
                return v
    return None


def apply_brand(spec, name):
    p = name.lower()
    for key, (brand, country, made) in BRAND_DEFAULTS.items():
        if key in p:
            spec.setdefault("Бренд [BRAND]", brand)
            spec.setdefault("Страна бренда [BRAND_COUNTRY]", country)
            spec.setdefault("Страна производства [MANUFACTURER]", made)
            return spec
    return spec


def rules_motorcycle(name):
    p = name.lower()
    spec = {}
    for key, row in MOTO_KNOWN.items():
        if key in p:
            brand, cc, cc_range, hp, hp_range, cooling, engine, mtype, bc, made, fuel, cyl = row
            spec.update({
                "Бренд [BRAND]": brand,
                "Объём двигателя, куб [ENGINE_CAPACITY]": cc,
                "Объём двигателя (по диапазонам) [ENGINE_CAPACITY1]": cc_range,
                "Мощность, л.с. [POWER_HP]": hp,
                "Мощность (по диапазонам) [POWER_HP1]": hp_range,
                "Охлаждение [COOLING]": cooling,
                "Двигатель [ENGINE]": engine,
                "Тип мотоцикла [Motorcycle_Type]": mtype,
                "Страна бренда [BRAND_COUNTRY]": bc,
                "Страна производства [MANUFACTURER]": made,
                "Система подачи топлива [Fuel_supply_system]": fuel,
                "Количество цилиндров [CYLINDERS]": cyl,
            })
            break
    apply_brand(spec, name)
    spec.setdefault("Гарантия [WARRANTY]", "1 год")
    spec.setdefault("Наличие ПТС [NALICHIE_PTS]", "Да")
    spec.setdefault("Система запуска [STARTING_SYSTEM]", "Электростартер")
    spec.setdefault("Материал рамы [FRAME_MATERIAL]", "Сталь")
    spec.setdefault("Тип топлива [TYPE_FUEL]", "АИ92-95")
    spec.setdefault("Количество тактов [STROKE]", 4)
    spec.setdefault("Тип двигателя [TYPE_ENGINE]", "Бензиновый")
    spec.setdefault("Трансмиссия [TRANSMISSION]", "Механическая")
    spec.setdefault("Карбюратор [Carburetor]", "Нет")
    return spec


def rules_boat_motor(name):
    p = name.lower()
    spec = {}
    apply_brand(spec, name)
    hp = hp_from_name(name)
    if hp:
        spec["Мощность, л.с. [POWER_HP1]"] = hp
        spec["Мощность (л.с.) [POWER_HP]"] = range_hp_motor(hp)
        spec["Мощность, л.с. [POWER_HP]"] = hp
        spec["Мощность (кВт) [POWER_KW]"] = round(hp * 0.7355, 1)
    spec.setdefault("Управление [OPERATION]", "Румпельное" if any(x in p for x in ["fh", "румп", "hes"]) else "Дистанционное")
    spec.setdefault("Система запуска [STARTING_SYSTEM]", "Электростартер" if any(x in p for x in ["e", "etl", "elpt", "xrtu", "efi"]) else "Ручной стартер/электростартер")
    spec.setdefault("Дейдвуд [DEADWOOD]", "635 (XL)" if any(x in p for x in ["xrtu", "fex", "xl"]) else ("508 (L)" if any(x in p for x in ["etl", "elpt", "fel", "lrt", "fvel"]) else "381 (S)"))
    spec.setdefault("Тип насадки [NOZZLETYPE]", "Водомёт" if "jet" in p or "водом" in p else "Винт")
    spec.setdefault("Система подачи топлива [Fuel_supply_system]", "Инжектор" if "efi" in p else "Карбюратор")
    spec.setdefault("Система подъёма [LIFTING_SYSTEM]", "Гидравлическая" if any(x in p for x in ["pt", "trim", "xrtu", "elpt", "btx"]) else "Ручная")
    spec.setdefault("Количество тактов [STROKE]", "2" if "2 такт" in p or "2-такт" in p or "t 40" in p else "4")
    spec.setdefault("Охлаждение [COOLING]", "Водяное")
    spec.setdefault("Тип двигателя [TYPE_ENGINE]", "Бензиновый")
    spec.setdefault("Передачи [GEAR]", "F-N-R")
    spec.setdefault("Тип топлива [TYPE_FUEL]", "АИ92-95")
    spec.setdefault("Вращение винта [ROTATION_SCREW]", "Водомётная насадка" if spec.get("Тип насадки [NOZZLETYPE]") == "Водомёт" else "Правое")
    spec.setdefault("Гарантия [WARRANTY]", "5 лет" if any(x in p for x in ["honda", "tohatsu"]) else "3 года")
    return spec


def rules_pvc_boat(name):
    p = name.lower()
    spec = {}
    apply_brand(spec, name)
    spec.setdefault("Тип лодки [TYPE_BOAT]", "Под мотор")
    spec.setdefault("Гарантия [WARRANTY]", "1 год")
    if "нднд" in p or "air" in p or "aero" in p:
        spec["Тип днища [TYPE_BOTTOM]"] = "Надувное, низкого давления"
    elif any(x in p for x in ["ал", "al"]):
        spec["Тип днища [TYPE_BOTTOM]"] = "Алюминиевые пайолы"
    else:
        spec["Тип днища [TYPE_BOTTOM]"] = "Надувное, низкого давления"
    spec.setdefault("Сливной клапан [DRAIN_VALVE]", "Есть")
    spec.setdefault("Надувной киль [INFLATABLE_KEEL]", "Есть")
    return spec


def rules_quad(name):
    p = name.lower()
    spec = {}
    apply_brand(spec, name)
    spec.setdefault("Охлаждение [COOLING]", "Жидкостное" if any(x in p for x in ["700", "800", "1000", "efi"]) else "Воздушное")
    spec.setdefault("Тип привода [Tipprivoda]", "Полный")
    spec.setdefault("Система привода [DRIVE_SYSTEM]", "Карданный")
    spec.setdefault("Трансмиссия [TRANSMISSION]", "Вариатор")
    spec.setdefault("Наличие ПСМ [NALICHIE_PSM]", "Есть" if "псм" in p else "Нет")
    spec.setdefault("Тип двигателя [TYPE_ENGINE]", "Бензиновый")
    spec.setdefault("Система подачи топлива [Fuel_supply_system]", "Инжектор" if "efi" in p else "Карбюратор")
    spec.setdefault("Лебедка [WINCH]", "Есть")
    spec.setdefault("Классификация [CLASSIFICATION]", "Утилитарный")
    spec.setdefault("Система запуска [STARTING_SYSTEM]", "Электростартер")
    spec.setdefault("Материал рамы [FRAME_MATERIAL]", "Сталь")
    spec.setdefault("Гарантия [WARRANTY]", "1 год")
    return spec


def rules_golfcar(name):
    p = name.lower()
    spec = {}
    apply_brand(spec, name)
    spec.setdefault("Тип двигателя [TYPE_ENGINE]", "Электрический")
    spec.setdefault("Система запуска [STARTING_SYSTEM]", "Электростартер")
    spec.setdefault("Система привода", "Полный привод" if "4x4" in p else "Задний привод")
    spec.setdefault("Пассажировместимость", "4" if "2+2" in p else "2")
    spec.setdefault("Страна бренда [BRAND_COUNTRY]", "Россия")
    spec.setdefault("Страна производства [MANUFACTURER]", "Китай")
    spec.setdefault("Гарантия [WARRANTY]", "1 год")
    return spec


def make_rules(name, category):
    if category == "мотоцикл" or category in ["дорожный мотоцикл", "внедорожный мотоцикл"]:
        spec = rules_motorcycle(name)
        spec.update(SERVICE_DEFAULTS.get("мотоцикл", {}))
        return spec
    if category == "лодочный мотор":
        spec = rules_boat_motor(name)
        spec.update(SERVICE_DEFAULTS.get("лодочный мотор", {}))
        return spec
    if category == "лодка пвх":
        spec = rules_pvc_boat(name)
        spec.update(SERVICE_DEFAULTS.get("лодка пвх", {}))
        return spec
    if category == "квадроцикл":
        spec = rules_quad(name)
        spec.update(SERVICE_DEFAULTS.get("квадроцикл", {}))
        return spec
    if category == "гольфкар":
        spec = rules_golfcar(name)
        spec.update(SERVICE_DEFAULTS.get("гольфкар", {}))
        return spec
    return {}


def process_excel(uploaded_file, category_mode, max_products, use_ai, use_search):
    wb = load_workbook(uploaded_file)
    ws = wb.active
    hmap = header_map(ws)
    headers = list(hmap.keys())

    for s in ["Отчет", "Проверить", "Лог поиска"]:
        if s in wb.sheetnames:
            del wb[s]

    report = wb.create_sheet("Отчет")
    report.append(["Показатель", "Значение"])
    check = wb.create_sheet("Проверить")
    check.append(["Строка", "Товар", "Поле", "Значение", "Комментарий"])

    rows = []
    for r in range(2, ws.max_row + 1):
        name = str(ws.cell(r, 1).value or "").strip()
        if name:
            rows.append((r, name))
    rows = rows[:max_products]

    if category_mode == "Авто по шаблону":
        category = detect_category(headers, [n for _, n in rows[:5]])
    else:
        category = category_mode

    changed = 0
    ai_ok = 0
    ai_fail = 0
    rules_ok = 0
    progress = st.progress(0)

    for idx, (r, name) in enumerate(rows, start=1):
        row_category = detect_category(headers, [name]) if category_mode == "Авто по строкам" else category
        spec = make_rules(name, row_category)
        if spec:
            rules_ok += 1

        source_text = ""
        if use_search:
            urls, snippets, search_logs = find_pages(name, row_category)
            source_text = "\n".join(snippets)
            for url in urls:
                txt, fetch_status = fetch_text(url)
                check.append([r, name, "Источник", url, fetch_status])
                if txt:
                    source_text += "\n\n" + txt
            if not urls:
                check.append([r, name, "Поиск", "", "Ссылки не найдены, будет fallback через AI по названию"])

        if use_ai:
            ai_spec, status = gemini_by_name(name, headers, row_category, source_text)
            if ai_spec:
                spec.update(ai_spec)
                ai_ok += 1
            else:
                ai_fail += 1
                check.append([r, name, "Gemini", "", status])
            time.sleep(2)

        if not spec:
            check.append([r, name, "Товар", "", "Нет правил и AI не вернул данные"])
            progress.progress(idx / len(rows))
            continue

        row_changed = 0
        for h, v in spec.items():
            row_changed += set_if_col(ws, hmap, r, h, v)

        changed += row_changed
        if row_changed == 0:
            check.append([r, name, "Заполнение", "", "Нечего изменить или нет колонок"])

        progress.progress(idx / len(rows))

    for row in [
        ["Категория", category],
        ["Режим", category_mode],
        ["Обработано товаров", len(rows)],
        ["Правила сработали", rules_ok],
        ["Gemini успешно", ai_ok],
        ["Gemini ошибки", ai_fail],
        ["Изменено ячеек", changed],
        ["Режим стабильности", "файл возвращается всегда"],
    ]:
        report.append(row)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


st.set_page_config(page_title="GD AutoFill Stable v17", layout="centered")
st.title("GD AutoFill Stable v17")
st.write("Стабильная версия со всеми основными категориями. Ищет характеристики в интернете, пропускает капчу и дозаполняет через Gemini.")

gemini_ok = bool(get_secret("GEMINI_API_KEY"))
st.info(f"Gemini API: {'✅ найден' if gemini_ok else '❌ не найден'}")

category_mode = st.selectbox(
    "Категория",
    [
        "Авто по шаблону",
        "Авто по строкам",
        "лодочный мотор",
        "лодка пвх",
        "квадроцикл",
        "дорожный мотоцикл",
        "внедорожный мотоцикл",
        "мотоцикл",
        "гольфкар",
    ],
)

max_products = st.number_input("Сколько товаров обработать за раз", min_value=1, max_value=500, value=500, help="500 означает: обработать все строки, которые есть в файле.")
use_search = st.checkbox("Искать характеристики в интернете через Serper/Google", value=True)
use_ai = st.checkbox("Дозаполнять новые товары через Gemini AI", value=False)

if use_ai:
    st.warning("ИИ может работать медленно на бесплатном тарифе. Для теста лучше 1–3 товара.")

uploaded = st.file_uploader("Загрузите Excel", type=["xlsx"])

if uploaded:
    st.success(f"Файл загружен: {uploaded.name}")
    try:
        _wb_preview = load_workbook(uploaded, read_only=True)
        _ws_preview = _wb_preview.active
        _names_preview = []
        for _r in range(2, _ws_preview.max_row + 1):
            _name = str(_ws_preview.cell(_r, 1).value or "").strip()
            if _name:
                _names_preview.append(_name)
        st.info(f"В файле найдено товаров: {len(_names_preview)}")
        uploaded.seek(0)
    except Exception as _e:
        st.warning(f"Не смогла посчитать строки файла: {_e}")
        uploaded.seek(0)

    if st.button("Заполнить и скачать"):
        with st.spinner("Заполняю файл..."):
            try:
                result = process_excel(uploaded, category_mode, int(max_products), use_ai, use_search)
                st.success("Готово")
                st.download_button(
                    "Скачать заполненный Excel",
                    data=result,
                    file_name=uploaded.name.replace(".xlsx", "_STABLE_v17.xlsx"),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error(f"Ошибка: {e}")
