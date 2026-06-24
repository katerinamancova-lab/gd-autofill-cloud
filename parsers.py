import re
import json

try:
    import google.generativeai as genai
except Exception:
    genai = None


def extract_json(text: str):
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


def num_after(text, labels, min_val=None, max_val=None):
    t = text.lower()
    for label in labels:
        lab = label.lower()
        start = 0
        while True:
            pos = t.find(lab, start)
            if pos == -1:
                break
            frag = text[pos:pos + 260]
            m = re.search(r"(\d+[,.]?\d*)", frag)
            if m:
                val = m.group(1).replace(",", ".")
                try:
                    f = float(val)
                    if (min_val is None or f >= min_val) and (max_val is None or f <= max_val):
                        return val
                except Exception:
                    pass
            start = pos + len(lab)
    return ""


def parse_basic_specs(source_text: str, category: str):
    spec = {}
    text = source_text or ""
    if not text.strip():
        return spec

    common_fields = {
        "Длина, см [LENGTH_CM]": (["длина лодки", "габаритная длина", "длина"], 100, 600),
        "Ширина, см [WIDTH]": (["ширина лодки", "габаритная ширина", "ширина"], 50, 300),
        "Вес, кг [WEIGHT]": (["вес комплекта", "вес лодки", "сухой вес", "масса", "вес"], 5, 500),
        "Сухой вес, кг [DRY_WEIGHT]": (["сухой вес", "вес лодки", "масса"], 5, 500),
        "Грузоподъемность, кг [LOAD_CAPACITY]": (["грузоподъемность", "грузоподъёмность"], 50, 2000),
        "Макс. мощность мотора, л.с. [MAX_POWER]": (["максимальная мощность мотора", "мощность мотора", "макс. мощность"], 1, 200),
        "Диаметр борта, см [BOAT_SIDE_DIAMETER]": (["диаметр баллона", "диаметр борта"], 20, 90),
        "Пассажировместимость, чел [PASSENGER]": (["пассажировместимость", "количество пассажиров", "вместимость"], 1, 20),
        "Объём двигателя, куб [ENGINE_CAPACITY]": (["объем двигателя", "объём двигателя", "куб.см", "см3"], 20, 3000),
        "Мощность, л.с. [POWER_HP]": (["мощность", "л.с."], 1, 350),
    }

    for header, (labels, mn, mx) in common_fields.items():
        val = num_after(text, labels, mn, mx)
        if val:
            spec[header] = val

    if category == "лодка пвх":
        spec.setdefault("Тип лодки [TYPE_BOAT]", "Под мотор")
        if "нднд" in text.lower() or "надувное дно" in text.lower():
            spec["Тип днища [TYPE_BOTTOM]"] = "Надувное, низкого давления"

    return spec


def gemini_extract(product_name: str, category: str, headers: list[str], source_text: str, api_key: str):
    if not api_key or genai is None:
        return {}, "Gemini недоступен"

    genai.configure(api_key=api_key)
    safe_headers = [
        h for h in headers
        if h and not any(x in h.lower() for x in ["uid", "уид", "активность", "розничная цена"])
    ]

    prompt = f"""
Ты контент-менеджер интернет-магазина техники.
Нужно заполнить Excel карточку товара строго по найденным источникам.

Товар: {product_name}
Категория: {category}

Колонки Excel:
{json.dumps(safe_headers, ensure_ascii=False)}

Текст найденных источников:
{source_text[:22000]}

Верни только JSON.
Ключи JSON должны точно совпадать с колонками Excel.
Не выдумывай. Если точной характеристики нет — пропусти.
Не заполняй УИД, UID, Активность, Розничная цена.
"""

    for model_name in ["gemini-2.0-flash", "gemini-flash-latest"]:
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

    return {}, f"Gemini ошибка: {last if 'last' in locals() else ''}"
