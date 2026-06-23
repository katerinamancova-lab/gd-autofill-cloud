import io
import os
import re
import json
import time
import requests
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

SERVICE_DEFAULTS = {
    "Комплектация": '<ul><li style="font-size: 13pt; font-family: Acrom">Мотоцикл</li><li style="font-size: 13pt; font-family: Acrom">Сервисная книжка</li></ul>',
    "Гарантия на товар": '<span style="font-family: Acrom; font-size: 13pt;">Гарантия на товар составляет 1 год</span>',
    "Скидка": 11,
    "Доступное количество": 1000,
    "Нет в продаже": None,
    "Сортировка": 500,
    "Привязка к аксессуарам (новая)": "Шлем кроссовый Sharmax SH536 Red/Black;Шлем кроссовый Sharmax SH336 Blue/Black;Мотозащита Sharmax черепаха RT 8;Очки кроссовые Sharmax Premium Black;Очки кроссовые Sharmax Gray/Black;Наколенники Sharmax пластик KP 48 Красные;Наколенники Sharmax SH-32K;Мотоперчатки Sharmax GL-SH 47 White ;Мотоперчатки Sharmax GL-SH 48 Yellow;Мотоперчатки Sharmax GL-SH 49 Red\n",
}

# База правил не заменяет поиск, а страхует программу, чтобы файл точно заполнялся и отдавался.
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
    "bajaj": ("Bajaj", "Индия", "Индия"),
    "voge": ("VOGE", "Китай", "Китай"),
    "qjmotor": ("QJMotor", "Китай", "Китай"),
    "royal enfield": ("Royal Enfield", "Индия", "Индия"),
    "ktm": ("KTM", "Австрия", "Индия"),
    "benelli": ("Benelli", "Италия", "Китай"),
    "stels": ("Stels", "Россия", "Китай"),
    "benda": ("Benda", "Китай", "Китай"),
}


def get_secret(name: str) -> str:
    try:
        value = st.secrets.get(name, "")
    except Exception:
        value = os.getenv(name, "")
    return str(value or "").strip()


def header_map(ws):
    return {str(ws.cell(1, c).value or "").strip(): c for c in range(1, ws.max_column + 1)}


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


def gemini_by_name(product_name, headers, category):
    api_key = get_secret("GEMINI_API_KEY")
    if not api_key or genai is None:
        return {}, "Gemini недоступен"

    genai.configure(api_key=api_key)
    safe_headers = [
        h for h in headers
        if h and not any(x in h.lower() for x in ["uid", "уид", "активность", "розничная цена"])
    ]

    prompt = f"""
Верни только JSON для заполнения Excel.
Товар: {product_name}
Категория: {category}

Колонки Excel:
{json.dumps(safe_headers, ensure_ascii=False)}

Правила:
- ключи JSON должны точно совпадать с колонками;
- не заполняй УИД, UID, Активность, Розничная цена;
- если точную цифру не знаешь — пропусти;
- бренды/страны/тип двигателя/топливо/запуск можно определить логически;
- заполни максимум характеристик.
"""

    for model_name in ["gemini-flash-latest", "gemini-2.0-flash", "gemini-1.5-flash-latest"]:
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
    return {}, f"Gemini не сработал: {last if 'last' in locals() else ''}"


def rules_motorcycle(product_name):
    p = product_name.lower()
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

    # Если модель новая, хотя бы определяем бренд/страны
    if "Бренд [BRAND]" not in spec:
        for key, (brand, bc, made) in BRAND_DEFAULTS.items():
            if key in p:
                spec["Бренд [BRAND]"] = brand
                spec["Страна бренда [BRAND_COUNTRY]"] = bc
                spec["Страна производства [MANUFACTURER]"] = made
                break

    spec.setdefault("Гарантия [WARRANTY]", "1 год")
    spec.setdefault("Наличие ПТС [NALICHIE_PTS]", "Да")
    spec.setdefault("Система запуска [STARTING_SYSTEM]", "Электростартер")
    spec.setdefault("Материал рамы [FRAME_MATERIAL]", "Сталь")
    spec.setdefault("Тип топлива [TYPE_FUEL]", "АИ92-95")
    spec.setdefault("Количество тактов [STROKE]", 4)
    spec.setdefault("Тип двигателя [TYPE_ENGINE]", "Бензиновый")
    spec.setdefault("Трансмиссия [TRANSMISSION]", "Механическая")
    spec.setdefault("Карбюратор [Carburetor]", "Нет")

    for h, v in SERVICE_DEFAULTS.items():
        spec[h] = v

    return spec


def process_excel(uploaded_file, category_mode, max_products, use_ai):
    wb = load_workbook(uploaded_file)
    ws = wb.active
    hmap = header_map(ws)
    headers = list(hmap.keys())

    # Удаляем старые отчёты
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

    changed = 0
    ai_ok = 0
    ai_fail = 0
    rules_ok = 0

    progress = st.progress(0)

    for idx, (r, name) in enumerate(rows, start=1):
        spec = {}
        if category_mode == "мотоцикл":
            spec.update(rules_motorcycle(name))
            rules_ok += 1

        if use_ai:
            ai_spec, status = gemini_by_name(name, headers, category_mode)
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

    report_rows = [
        ["Категория", category_mode],
        ["Обработано товаров", len(rows)],
        ["Правила сработали", rules_ok],
        ["Gemini успешно", ai_ok],
        ["Gemini ошибки", ai_fail],
        ["Изменено ячеек", changed],
        ["Режим", "стабильный: файл возвращается всегда"],
    ]
    for row in report_rows:
        report.append(row)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


st.set_page_config(page_title="GD AutoFill Stable v14", layout="centered")
st.title("GD AutoFill Stable v14")
st.write("Стабильная версия: файл всегда отдаётся обратно. ИИ можно включать только для новых товаров.")

gemini_ok = bool(get_secret("GEMINI_API_KEY"))
st.info(f"Gemini API: {'✅ найден' if gemini_ok else '❌ не найден'}")

category_mode = st.selectbox("Категория", ["мотоцикл"])
max_products = st.number_input("Сколько товаров обработать за раз", min_value=1, max_value=100, value=30)
use_ai = st.checkbox("Дозаполнять новые товары через Gemini AI", value=False)

if use_ai:
    st.warning("ИИ может работать медленно на бесплатном тарифе. Для теста лучше 1–3 товара.")

uploaded = st.file_uploader("Загрузите Excel", type=["xlsx"])

if uploaded:
    st.success(f"Файл загружен: {uploaded.name}")
    if st.button("Заполнить и скачать"):
        with st.spinner("Заполняю файл..."):
            try:
                result = process_excel(uploaded, category_mode, int(max_products), use_ai)
                st.success("Готово")
                st.download_button(
                    "Скачать заполненный Excel",
                    data=result,
                    file_name=uploaded.name.replace(".xlsx", "_STABLE_v14.xlsx"),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error(f"Ошибка: {e}")
